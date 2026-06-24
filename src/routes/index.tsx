import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useRef, useState } from "react";
import { toPng } from "html-to-image";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, ReferenceLine, Tooltip, LabelList, CartesianGrid,
} from "recharts";
import {
  parseExcel, prepare, rankByInterval, summaryByHour, summaryByDay, combineDateAndTime,
  type Prepared,
} from "@/lib/produtividade";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Produtividade por Hora — Separação" },
      { name: "description", content: "Dashboard de produtividade de separação por hora e por dia a partir de Excel." },
      { property: "og:title", content: "Produtividade por Hora — Separação" },
      { property: "og:description", content: "Dashboard de produtividade de separação por hora e por dia." },
    ],
  }),
  component: Index,
});

function Index() {
  const [fileName, setFileName] = useState<string>("");
  const [sheets, setSheets] = useState<string[]>([]);
  const [rowsBySheet, setRowsBySheet] = useState<Record<string, Record<string, unknown>[]>>({});
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [prepared, setPrepared] = useState<Prepared | null>(null);

  const [date, setDate] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [h1Start, setH1Start] = useState("14:00");
  const [h1End, setH1End] = useState("15:00");
  const [d2Start, setD2Start] = useState("06:00");
  const [d2End, setD2End] = useState("17:00");
  const [meta, setMeta] = useState<number>(35);

  const [error, setError] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [calculated, setCalculated] = useState(false);

  async function handleFile(f: File) {
    setError("");
    setLoading(true);
    try {
      const parsed = await parseExcel(f);
      setFileName(f.name);
      setSheets(parsed.sheetNames);
      setRowsBySheet(parsed.rowsBySheet);
      const first = parsed.sheetNames[0];
      setSelectedSheet(first);
      const prep = prepare(parsed.rowsBySheet[first] ?? []);
      setPrepared(prep);
      if (prep.missing.length) {
        setError(`Colunas obrigatórias ausentes: ${prep.missing.join(", ")}`);
      }
      setCalculated(false);
    } catch (e) {
      setError(`Falha ao ler o arquivo: ${(e as Error).message}`);
    } finally {
      setLoading(false);
    }
  }

  function onSheetChange(name: string) {
    setSelectedSheet(name);
    const prep = prepare(rowsBySheet[name] ?? []);
    setPrepared(prep);
    setError(prep.missing.length ? `Colunas obrigatórias ausentes: ${prep.missing.join(", ")}` : "");
    setCalculated(false);
  }

  const result = useMemo(() => {
    if (!prepared || !calculated) return null;
    if (h1Start >= h1End) return { err: "Intervalo 1: hora inicial deve ser menor que a final" };
    if (d2Start >= d2End) return { err: "Intervalo 2: hora inicial deve ser menor que a final" };
    const i1Start = combineDateAndTime(date, h1Start);
    const i1End = combineDateAndTime(date, h1End);
    const i2Start = combineDateAndTime(date, d2Start);
    const i2End = combineDateAndTime(date, d2End);
    return {
      i1: rankByInterval(prepared.rows, i1Start, i1End),
      i2: rankByInterval(prepared.rows, i2Start, i2End),
      sumHour: summaryByHour(prepared.rows, i2Start, i2End),
      sumDay: summaryByDay(prepared.rows, i2Start, i2End),
    };
  }, [prepared, calculated, date, h1Start, h1End, d2Start, d2End]);

  function calcular() {
    setError("");
    if (!prepared) { setError("Faça o upload de um arquivo Excel primeiro."); return; }
    if (prepared.missing.length) { setError(`Colunas obrigatórias ausentes: ${prepared.missing.join(", ")}`); return; }
    if (!date) { setError("Selecione uma data."); return; }
    if (h1Start >= h1End) { setError("Intervalo 1: hora inicial deve ser menor que a final."); return; }
    if (d2Start >= d2End) { setError("Intervalo 2: hora inicial deve ser menor que a final."); return; }
    setCalculated(true);
  }

  function limpar() {
    setCalculated(false);
    setError("");
  }

  return (
    <div className="min-h-screen text-foreground">
      <header>
        <div className="mx-auto flex max-w-7xl items-center gap-3 px-6 pt-10 pb-6">
          <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-primary text-primary-foreground shadow-sm">
            <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 3v18h18"/><path d="M7 14l4-4 3 3 5-6"/></svg>
          </div>
          <div>
            <h1 className="text-2xl font-bold tracking-tight">Produtividade por Hora — Separação</h1>
            <p className="mt-0.5 text-sm text-muted-foreground">
              Faça upload da planilha, defina os filtros e analise o desempenho dos separadores.
            </p>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 px-6 py-6">
        {/* Filtros */}
        <section className="rounded-xl border bg-card p-5 shadow-sm">
          <div className="grid gap-4 md:grid-cols-12">
            <div className="md:col-span-4">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Arquivo Excel</label>
              <label className="flex h-10 cursor-pointer items-center gap-2 rounded-md border bg-background px-3 text-sm hover:bg-accent">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
                />
                <span className="font-medium">{loading ? "Lendo…" : "Upload Excel"}</span>
                <span className="truncate text-muted-foreground">{fileName || "Nenhum arquivo"}</span>
              </label>
            </div>

            <div className="md:col-span-3">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Aba</label>
              <select
                value={selectedSheet}
                onChange={(e) => onSheetChange(e.target.value)}
                disabled={!sheets.length}
                className="h-10 w-full rounded-md border bg-background px-3 text-sm disabled:opacity-50"
              >
                {sheets.length === 0 && <option>—</option>}
                {sheets.map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>

            <div className="md:col-span-3">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Data</label>
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="h-10 w-full rounded-md border bg-background px-3 text-sm"
              />
            </div>

            <div className="md:col-span-2">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Meta (prod/h)</label>
              <input
                type="number"
                step="0.1"
                value={meta}
                onChange={(e) => setMeta(parseFloat(e.target.value) || 0)}
                className="h-10 w-full rounded-md border bg-background px-3 text-sm"
              />
            </div>

            <div className="md:col-span-3">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Intervalo 1 (hora)</label>
              <div className="flex items-center gap-2">
                <input type="time" value={h1Start} onChange={(e) => setH1Start(e.target.value)} className="h-10 w-full rounded-md border bg-background px-3 text-sm" />
                <span className="text-muted-foreground">—</span>
                <input type="time" value={h1End} onChange={(e) => setH1End(e.target.value)} className="h-10 w-full rounded-md border bg-background px-3 text-sm" />
              </div>
            </div>

            <div className="md:col-span-3">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Intervalo 2 (dia)</label>
              <div className="flex items-center gap-2">
                <input type="time" value={d2Start} onChange={(e) => setD2Start(e.target.value)} className="h-10 w-full rounded-md border bg-background px-3 text-sm" />
                <span className="text-muted-foreground">—</span>
                <input type="time" value={d2End} onChange={(e) => setD2End(e.target.value)} className="h-10 w-full rounded-md border bg-background px-3 text-sm" />
              </div>
            </div>

            <div className="flex items-end gap-2 md:col-span-6">
              <button
                onClick={calcular}
                className="h-10 flex-1 rounded-md bg-primary px-4 text-sm font-medium text-primary-foreground transition hover:opacity-90"
              >
                Calcular
              </button>
              <button
                onClick={limpar}
                className="h-10 rounded-md border bg-background px-4 text-sm font-medium hover:bg-accent"
              >
                Limpar
              </button>
              <button
                onClick={() => window.print()}
                className="h-10 rounded-md border bg-background px-4 text-sm font-medium hover:bg-accent"
              >
                Imprimir
              </button>
            </div>
          </div>

          {error && (
            <div className="mt-4 rounded-md border border-destructive/30 bg-destructive/10 px-3 py-2 text-sm text-destructive">
              {error}
            </div>
          )}
        </section>

        {result && "err" in result && (
          <div className="rounded-md border border-destructive/30 bg-destructive/10 px-3 py-2 text-sm text-destructive">
            {result.err}
          </div>
        )}

        {result && !("err" in result) && (
          <>
            <KpiRow
              totalSeparado={result.i2.totalGeral}
              separadoresAtivos={result.i2.separadoresAtivos}
              media={result.i2.media}
              meta={meta}
            />

            <div className="grid gap-6 lg:grid-cols-2">
              <ChartCard
                title="Desempenho por separador (prod/h)"
                subtitle={`Intervalo 1 (Hora) — ${h1Start} às ${h1End}`}
                data={result.i1.ranking}
                media={result.i1.media}
                meta={meta}
              />
              <ChartCard
                title="Desempenho por separador (prod/h)"
                subtitle={`Intervalo 2 (Dia) — ${d2Start} às ${d2End}`}
                data={result.i2.ranking}
                media={result.i2.media}
                meta={meta}
              />
            </div>

            <div className="grid gap-6 lg:grid-cols-2">
              <SummaryTable
                title="Resumo por hora"
                head={["Hora", "Total", "Separadores", "Média/sep"]}
                rows={result.sumHour.map((r) => [
                  fmtHour(r.hora),
                  fmtInt(r.total),
                  String(r.separadoresAtivos),
                  fmt1(r.media),
                ])}
              />
              <SummaryTable
                title="Resumo por dia"
                head={["Dia", "Total", "Horas", "Separadores", "Prod/h"]}
                rows={result.sumDay.map((r) => [
                  fmtDay(r.dia),
                  fmtInt(r.total),
                  String(r.horas),
                  String(r.separadoresAtivos),
                  fmt1(r.prodHora),
                ])}
              />
            </div>
          </>
        )}

        {!result && (
          <div className="rounded-xl border border-dashed bg-card p-10 text-center text-sm text-muted-foreground">
            Faça upload de uma planilha e clique em <span className="font-medium text-foreground">Calcular</span> para ver os indicadores.
          </div>
        )}
      </main>
    </div>
  );
}

function KpiRow({ totalSeparado, separadoresAtivos, media, meta }: {
  totalSeparado: number; separadoresAtivos: number; media: number; meta: number;
}) {
  const diff = media - meta;
  return (
    <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-5">
      <Kpi label="Total separado" value={fmtInt(totalSeparado)} />
      <Kpi label="Separadores ativos" value={String(separadoresAtivos)} />
      <Kpi label="Média por separador (prod/h)" value={fmt1(media)} />
      <Kpi label="Meta (prod/h)" value={fmt1(meta)} />
      <Kpi
        label="Diferença para meta"
        value={(diff >= 0 ? "+" : "") + fmt1(diff)}
        tone={diff >= 0 ? "positive" : "negative"}
      />
    </div>
  );
}

function Kpi({ label, value, tone }: { label: string; value: string; tone?: "positive" | "negative" }) {
  const toneCls = tone === "positive" ? "text-emerald-600" : tone === "negative" ? "text-destructive" : "text-foreground";
  return (
    <div className="rounded-xl border bg-card p-4 shadow-sm">
      <div className="text-xs font-medium uppercase tracking-wide text-muted-foreground">{label}</div>
      <div className={`mt-2 text-2xl font-semibold ${toneCls}`}>{value}</div>
    </div>
  );
}

function ChartCard({
  title, subtitle, data, media, meta,
}: {
  title: string;
  subtitle: string;
  data: { separador: string; prodHora: number; total: number; horasAtivas: number }[];
  media: number;
  meta: number;
}) {
  const height = Math.max(220, data.length * 36 + 80);
  const ref = useRef<HTMLDivElement>(null);
  const [copyState, setCopyState] = useState<"idle" | "copying" | "done" | "err">("idle");

  async function copyChart() {
    if (!ref.current) return;
    setCopyState("copying");
    try {
      const dataUrl = await toPng(ref.current, {
        pixelRatio: 3,
        backgroundColor: "#ffffff",
        cacheBust: true,
      });
      const blob = await (await fetch(dataUrl)).blob();
      if (navigator.clipboard && "write" in navigator.clipboard && typeof ClipboardItem !== "undefined") {
        await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        setCopyState("done");
      } else {
        const a = document.createElement("a");
        a.href = dataUrl;
        a.download = `${title}.png`;
        a.click();
        setCopyState("done");
      }
      setTimeout(() => setCopyState("idle"), 1800);
    } catch {
      setCopyState("err");
      setTimeout(() => setCopyState("idle"), 2200);
    }
  }

  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm">
      <div className="mb-1 flex items-baseline justify-between gap-3">
        <h3 className="text-base font-semibold">{title}</h3>
        <button
          onClick={copyChart}
          disabled={data.length === 0 || copyState === "copying"}
          className="rounded-md border bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent disabled:opacity-50"
          title="Copiar gráfico para a área de transferência"
        >
          {copyState === "copying" ? "Copiando…"
            : copyState === "done" ? "Copiado ✓"
            : copyState === "err" ? "Erro ao copiar"
            : "Copiar gráfico"}
        </button>
      </div>
      <p className="mb-4 text-xs text-muted-foreground">{subtitle}</p>
      {data.length === 0 ? (
        <div className="py-12 text-center text-sm text-muted-foreground">Sem dados no intervalo.</div>
      ) : (
        <div ref={ref} className="bg-card p-3" style={{ width: "100%", height: height + 60 }}>
          <div className="mb-2 text-center text-sm font-semibold">{subtitle}</div>
          <div style={{ width: "100%", height }}>
          <ResponsiveContainer>
            <BarChart data={data} layout="vertical" margin={{ left: 24, right: 36, top: 8, bottom: 24 }}>
              <CartesianGrid horizontal={false} stroke="var(--border)" />
              <XAxis type="number" tick={{ fontSize: 11, fill: "var(--muted-foreground)" }} />
              <YAxis
                type="category"
                dataKey="separador"
                width={180}
                tick={{ fontSize: 11, fill: "var(--foreground)" }}
              />
              <Tooltip
                cursor={{ fill: "var(--accent)" }}
                contentStyle={{ background: "var(--card)", border: "1px solid var(--border)", borderRadius: 8, fontSize: 12 }}
                formatter={(v: number) => [v.toFixed(1), "prod/h"]}
              />
              <ReferenceLine x={meta} stroke="var(--chart-4)" strokeDasharray="4 4" label={{ value: `Meta ${meta.toFixed(1)}`, position: "top", fill: "var(--chart-4)", fontSize: 11 }} />
              <ReferenceLine x={media} stroke="var(--chart-2)" strokeDasharray="2 4" label={{ value: `Média ${media.toFixed(1)}`, position: "insideTopRight", fill: "var(--chart-2)", fontSize: 11 }} />
              <Bar dataKey="prodHora" fill="var(--primary)" radius={[0, 6, 6, 0]} barSize={20}>
                <LabelList dataKey="prodHora" position="right" formatter={(v: number) => v.toFixed(1)} style={{ fontSize: 11, fill: "var(--foreground)" }} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          </div>
        </div>
      )}
    </div>
  );
}

function SummaryTable({ title, head, rows }: { title: string; head: string[]; rows: string[][] }) {
  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm">
      <h3 className="mb-3 text-base font-semibold">{title}</h3>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b text-left text-xs uppercase tracking-wide text-muted-foreground">
              {head.map((h) => <th key={h} className="py-2 pr-3 font-medium">{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={head.length} className="py-6 text-center text-muted-foreground">Sem dados.</td></tr>
            ) : rows.map((r, i) => (
              <tr key={i} className="border-b last:border-0">
                {r.map((c, j) => <td key={j} className="py-2 pr-3 tabular-nums">{c}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function fmtInt(n: number) { return new Intl.NumberFormat("pt-BR").format(Math.round(n)); }
function fmt1(n: number) { return new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(n); }
function pad(n: number) { return n.toString().padStart(2, "0"); }
function fmtHour(d: Date) { return `${pad(d.getHours())}:00`; }
function fmtDay(d: Date) { return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; }
