import * as XLSX from "xlsx";

export type Row = Record<string, unknown>;

export interface ParsedSheet {
  sheetNames: string[];
  rowsBySheet: Record<string, Row[]>;
}

export async function parseExcel(file: File): Promise<ParsedSheet> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const rowsBySheet: Record<string, Row[]> = {};
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    rowsBySheet[name] = XLSX.utils.sheet_to_json<Row>(ws, { defval: null, raw: true });
  }
  return { sheetNames: wb.SheetNames, rowsBySheet };
}

function normalizeKey(k: string) {
  return k
    .toString()
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, "_");
}

export function normalizeRows(rows: Row[]): Row[] {
  return rows.map((r) => {
    const out: Row = {};
    for (const k of Object.keys(r)) out[normalizeKey(k)] = r[k];
    return out;
  });
}

// Excel serial date -> JS Date
function excelSerialToDate(n: number): Date {
  // Excel epoch: 1899-12-30. Treat serial as a "naive" wall-clock value
  // (matches how pandas/openpyxl read it) so 18:00 stays 18:00 in local time
  // regardless of the user's timezone.
  const whole = Math.floor(n);
  const frac = n - whole;
  const base = new Date(1899, 11, 30);
  base.setDate(base.getDate() + whole);
  const totalSec = Math.round(frac * 86400);
  const hh = Math.floor(totalSec / 3600);
  const mm = Math.floor((totalSec % 3600) / 60);
  const ss = totalSec % 60;
  base.setHours(hh, mm, ss, 0);
  return base;
}

const MONTHS_PT: Record<string, number> = {
  jan: 0, fev: 1, mar: 2, abr: 3, mai: 4, jun: 5,
  jul: 6, ago: 7, set: 8, out: 9, nov: 10, dez: 11,
};

export function parseDate(value: unknown): Date | null {
  if (value == null || value === "") return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (typeof value === "number") {
    const d = excelSerialToDate(value);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(value).trim();
  if (!s) return null;

  // dd/mm/yyyy [hh:mm[:ss]]
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const [, dd, mm, yy, hh, mi, ss] = m;
    const year = yy.length === 2 ? 2000 + parseInt(yy, 10) : parseInt(yy, 10);
    return new Date(year, +mm - 1, +dd, +(hh ?? 0), +(mi ?? 0), +(ss ?? 0));
  }
  // yyyy-mm-dd[ hh:mm[:ss]]
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    const [, yy, mm, dd, hh, mi, ss] = m;
    return new Date(+yy, +mm - 1, +dd, +(hh ?? 0), +(mi ?? 0), +(ss ?? 0));
  }
  // dd-mmm-yyyy ex: 12-jan-2024
  m = s.match(/^(\d{1,2})[- ]([a-zç]{3,})[- ](\d{2,4})(?:[ T](\d{1,2}):(\d{2}))?$/i);
  if (m) {
    const [, dd, mon, yy, hh, mi] = m;
    const mi2 = MONTHS_PT[mon.toLowerCase().slice(0, 3)];
    if (mi2 != null) {
      const year = yy.length === 2 ? 2000 + parseInt(yy, 10) : parseInt(yy, 10);
      return new Date(year, mi2, +dd, +(hh ?? 0), +(mi ?? 0));
    }
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

export function toInt(v: unknown): number {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return Math.trunc(v);
  const s = String(v).replace(/\./g, "").replace(",", ".").trim();
  const n = parseFloat(s);
  return isFinite(n) ? Math.trunc(n) : 0;
}

export interface Prepared {
  rows: Array<{
    qt: number;
    dt: Date;
    separador: string;
    endereco: string | null;
    zona: string | null;
    turno: string | null;
    horaInicio: Date;
    diaInicio: Date;
  }>;
  enderecoField: string | null;
  missing: string[];
}

const ENDERECO_CANDIDATES = ["endereco", "endereco_picking", "end_picking", "rua", "posicao", "local", "endereco_separacao", "cd_endereco"];

const QT_CANDIDATES = ["qt_separada", "quantidade_separada", "qtd_separada", "qt_separado"];
const DT_CANDIDATES = ["dt_separacao", "data_separacao", "data_hora_separacao"];
const NOME_CANDIDATES = ["separador", "nome_separador", "rotulo_separador", "nm_funcionario"];
const COD_CANDIDATES = ["codigo_separador", "cod_separador", "cd_separador", "matricula", "cd_funcionario"];
const ZONA_CANDIDATES = ["cd_classe", "classe", "zona"];
const TURNO_CANDIDATES = ["turno_rota", "turno", "cd_turno"];

// Some exports append the SQL type to the header (e.g. "QT_SEPARADO NUMBER").
const TYPE_SUFFIX = /_(varchar2|varchar|number|char|date|float|integer|int|numeric|timestamp)$/;

function withAliases(rows: Row[]): Row[] {
  return rows.map((r) => {
    const out: Row = { ...r };
    for (const k of Object.keys(r)) {
      const alias = k.replace(TYPE_SUFFIX, "");
      if (alias !== k && !(alias in out)) out[alias] = r[k];
    }
    return out;
  });
}

function pick(obj: Row, cands: string[]): string | null {
  for (const c of cands) if (c in obj) return c;
  return null;
}

export function prepare(rawRows: Row[]): Prepared {
  const rows = withAliases(normalizeRows(rawRows));
  const missing: string[] = [];
  const first = rows[0] ?? {};
  const qtKey = pick(first, QT_CANDIDATES);
  const dtKey = pick(first, DT_CANDIDATES);
  const nomeKey = pick(first, NOME_CANDIDATES);
  const codKey = pick(first, COD_CANDIDATES);
  const zonaKey = pick(first, ZONA_CANDIDATES);
  const turnoKey = pick(first, TURNO_CANDIDATES);
  if (!qtKey) missing.push("Qt. Separada");
  if (!dtKey) missing.push("Dt. Separação");
  if (!nomeKey && !codKey) missing.push("Separador");

  let enderecoField: string | null = null;
  for (const c of ENDERECO_CANDIDATES) {
    if (first && c in first) { enderecoField = c; break; }
  }

  const prepared: Prepared["rows"] = [];
  if (missing.length) return { rows: prepared, enderecoField, missing };
  for (const r of rows) {
    const dt = parseDate(r[dtKey!]);
    if (!dt) continue;
    const qt = toInt(r[qtKey!]);
    const cod = codKey ? String(r[codKey] ?? "").trim() : "";
    const nome = nomeKey ? String(r[nomeKey] ?? "").trim() : "";
    const sep = (cod && nome) ? `${cod} - ${nome}` : (cod || nome || "—");
    const horaInicio = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours());
    const diaInicio = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    const endereco = enderecoField ? (r[enderecoField] == null ? null : String(r[enderecoField])) : null;
    const zonaRaw = zonaKey ? (r[zonaKey] == null ? "" : String(r[zonaKey]).trim()) : "";
    const zona = zonaRaw || (endereco && endereco.trim().length >= 3 ? endereco.trim().slice(0, 3).toUpperCase() : null);
    const turnoRaw = turnoKey ? (r[turnoKey] == null ? "" : String(r[turnoKey]).trim()) : "";
    prepared.push({
      qt,
      dt,
      separador: sep,
      endereco,
      zona: zona || null,
      turno: turnoRaw || null,
      horaInicio,
      diaInicio,
    });
  }

  return { rows: prepared, enderecoField, missing };
}


export interface RankItem {
  separador: string;
  total: number;
  horasAtivas: number;
  enderecos: number;
  prodHora: number;
}

function uniqueHours(dates: Date[]) {
  const s = new Set(dates.map((d) => d.getTime()));
  return s.size;
}

export function rankByInterval(
  rows: Prepared["rows"],
  start: Date,
  end: Date,
): { ranking: RankItem[]; media: number; totalGeral: number; separadoresAtivos: number } {
  const filtered = rows.filter((r) => r.dt >= start && r.dt < end);
  const bySep = new Map<string, typeof filtered>();
  for (const r of filtered) {
    const arr = bySep.get(r.separador) ?? [];
    arr.push(r);
    bySep.set(r.separador, arr);
  }
  const ranking: RankItem[] = [];
  for (const [sep, items] of bySep) {
    const total = items.reduce((a, b) => a + b.qt, 0);
    const horas = uniqueHours(items.map((i) => i.horaInicio));
    const enderecos = new Set(items.map((i) => i.endereco).filter((x) => x != null) as string[]).size;
    const prodHora = horas > 0 ? total / horas : 0;
    ranking.push({ separador: sep, total, horasAtivas: horas, enderecos, prodHora });
  }
  ranking.sort((a, b) => b.prodHora - a.prodHora);
  const media = ranking.length > 0 ? ranking.reduce((a, b) => a + b.prodHora, 0) / ranking.length : 0;
  const totalGeral = ranking.reduce((a, b) => a + b.total, 0);
  return { ranking, media, totalGeral, separadoresAtivos: ranking.length };
}

export function summaryByHour(rows: Prepared["rows"], start: Date, end: Date) {
  const filtered = rows.filter((r) => r.dt >= start && r.dt < end);
  const map = new Map<number, { hora: Date; total: number; separadores: Set<string> }>();
  for (const r of filtered) {
    const k = r.horaInicio.getTime();
    const cur = map.get(k) ?? { hora: r.horaInicio, total: 0, separadores: new Set<string>() };
    cur.total += r.qt;
    cur.separadores.add(r.separador);
    map.set(k, cur);
  }
  return Array.from(map.values())
    .sort((a, b) => a.hora.getTime() - b.hora.getTime())
    .map((v) => ({
      hora: v.hora,
      total: v.total,
      separadoresAtivos: v.separadores.size,
      media: v.separadores.size > 0 ? v.total / v.separadores.size : 0,
    }));
}

export function summaryByDay(rows: Prepared["rows"], start: Date, end: Date) {
  const filtered = rows.filter((r) => r.dt >= start && r.dt < end);
  const map = new Map<number, { dia: Date; total: number; separadores: Set<string>; horas: Set<number> }>();
  for (const r of filtered) {
    const k = r.diaInicio.getTime();
    const cur = map.get(k) ?? { dia: r.diaInicio, total: 0, separadores: new Set<string>(), horas: new Set<number>() };
    cur.total += r.qt;
    cur.separadores.add(r.separador);
    cur.horas.add(r.horaInicio.getTime());
    map.set(k, cur);
  }
  return Array.from(map.values())
    .sort((a, b) => a.dia.getTime() - b.dia.getTime())
    .map((v) => ({
      dia: v.dia,
      total: v.total,
      separadoresAtivos: v.separadores.size,
      horas: v.horas.size,
      prodHora: v.horas.size > 0 ? v.total / v.horas.size : 0,
    }));
}

export function combineDateAndTime(dateStr: string, timeStr: string): Date {
  const [y, m, d] = dateStr.split("-").map(Number);
  const [hh, mm] = timeStr.split(":").map(Number);
  return new Date(y, m - 1, d, hh, mm, 0);
}