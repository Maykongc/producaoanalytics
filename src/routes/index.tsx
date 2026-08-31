import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { toPng } from "html-to-image";
import {
  BarChart, Bar, XAxis, YAxis, ResponsiveContainer, ReferenceLine, LabelList, CartesianGrid, Cell,
} from "recharts";
import {
  parseExcel, prepare, rankByInterval, combineDateAndTime,
  type Prepared, type RankItem,
} from "@/lib/produtividade";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Produtividade - Separação" },
      { name: "description", content: "Dashboard de produtividade de separação por hora e por dia a partir de Excel." },
      { property: "og:title", content: "Produtividade - Separação" },
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

  const [date, setDate] = useState<string>(() => {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dd}`;
  });
  const [h1Start, setH1Start] = useState("");
  const [h1End, setH1End] = useState("");
  const [d2Start, setD2Start] = useState("");
  const [d2End, setD2End] = useState("");
  const [meta, setMeta] = useState<number>(35);
  const [zonaSel, setZonaSel] = useState<string[]>([]);
  
  const [funcionarioSel, setFuncionarioSel] = useState<string[]>([]);


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

  const zonas = useMemo(() => {
    if (!prepared) return [] as string[];
    const s = new Set<string>();
    for (const r of prepared.rows) if (r.zona) s.add(r.zona);
    return Array.from(s).sort();
  }, [prepared]);

  const funcionarios = useMemo(() => {
    if (!prepared) return [] as string[];
    const s = new Set<string>();
    for (const r of prepared.rows) if (r.separador) s.add(r.separador);
    return Array.from(s).sort();
  }, [prepared]);

  useEffect(() => {
    setFuncionarioSel(funcionarios);
  }, [funcionarios]);



  const result = useMemo(() => {
    if (!prepared || !calculated) return null;
    const has1 = Boolean(h1Start && h1End);
    const has2 = Boolean(d2Start && d2End);
    if (!has1 && !has2) return { err: "Preencha os horários de pelo menos um intervalo." };
    if (has1 && h1Start >= h1End) return { err: "Intervalo 1: hora inicial deve ser menor que a final" };
    if (has2 && d2Start >= d2End) return { err: "Intervalo 2: hora inicial deve ser menor que a final" };
    const rows = prepared.rows.filter((r) =>
      (zonaSel.length === 0 || zonaSel.includes(r.zona ?? "")) &&
      (funcionarioSel.length === 0 || funcionarioSel.includes(r.separador ?? "")),
    );
    const i1Start = has1 ? combineDateAndTime(date, h1Start) : null;
    const i1End = has1 ? combineDateAndTime(date, h1End) : null;
    const i2Start = has2 ? combineDateAndTime(date, d2Start) : null;
    const i2End = has2 ? combineDateAndTime(date, d2End) : null;
    return {
      i1: i1Start && i1End ? rankByInterval(rows, i1Start, i1End) : null,
      i2: i2Start && i2End ? rankByInterval(rows, i2Start, i2End) : null,
      i1Start, i1End, i2Start, i2End,
    };
  }, [prepared, calculated, date, h1Start, h1End, d2Start, d2End, zonaSel, funcionarioSel]);


  function calcular() {
    setError("");
    if (!prepared) { setError("Faça o upload de um arquivo Excel primeiro."); return; }
    if (prepared.missing.length) { setError(`Colunas obrigatórias ausentes: ${prepared.missing.join(", ")}`); return; }
    if (!date) { setError("Selecione uma data."); return; }
    const has1 = Boolean(h1Start && h1End);
    const has2 = Boolean(d2Start && d2End);
    if (!has1 && !has2) { setError("Preencha os horários de pelo menos um intervalo (hora ou dia)."); return; }
    if ((h1Start || h1End) && !has1) { setError("Intervalo 1: preencha as horas inicial e final."); return; }
    if ((d2Start || d2End) && !has2) { setError("Intervalo 2: preencha as horas inicial e final."); return; }
    if (has1 && h1Start >= h1End) { setError("Intervalo 1: hora inicial deve ser menor que a final."); return; }
    if (has2 && d2Start >= d2End) { setError("Intervalo 2: hora inicial deve ser menor que a final."); return; }
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
            <h1 className="text-2xl font-bold tracking-tight">Produtividade - Separação</h1>
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
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Zona</label>
              <MultiSelect
                options={zonas}
                selected={zonaSel}
                onChange={setZonaSel}
                allLabel="Todas as zonas"
                placeholder=""
              />
            </div>

            <div className="md:col-span-5">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Funcionário</label>
              <MultiSelect
                options={funcionarios}
                selected={funcionarioSel}
                onChange={setFuncionarioSel}
                allLabel="Todos os funcionários"
                searchable
                searchPlaceholder="Buscar por matrícula ou nome…"
              />
            </div>



            <div className="md:col-span-6">
              <label className="mb-1 block text-xs font-medium text-muted-foreground">Intervalo 1 (hora)</label>
              <div className="flex items-center gap-2">
                <input type="time" value={h1Start} onChange={(e) => setH1Start(e.target.value)} className="h-10 w-full rounded-md border bg-background px-3 text-sm" />
                <span className="text-muted-foreground">—</span>
                <input type="time" value={h1End} onChange={(e) => setH1End(e.target.value)} className="h-10 w-full rounded-md border bg-background px-3 text-sm" />
              </div>
            </div>

            <div className="md:col-span-6">
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

        {result && !("err" in result) && (() => {
          const zonaLabel = zonaSel.length === 0 ? "Todas" : zonaSel.join(", ");
          return (
          <>
            <ChartReport
              kind="hora"
              headerTitle={`Produtividade por Hora - Separação - Zona (${zonaLabel})`}
              chartTitle={`Desempenho por separador (prod/h) - Intervalo 1 (Hora)`}
              intervaloLabel={`${fmtDay(result.i1Start)} ${h1Start} às ${h1End}`}
              ranking={result.i1.ranking}
              media={result.i1.media}
              total={result.i1.totalGeral}
              separadoresAtivos={result.i1.separadoresAtivos}
              meta={meta}
            />
            <ChartReport
              kind="dia"
              headerTitle={`Produtividade por Dia - Separação - Zona (${zonaLabel})`}
              chartTitle={`Desempenho por separador (prod/h) - Intervalo 2 (Dia)`}
              intervaloLabel={`${fmtDay(result.i2Start)}`}
              ranking={result.i2.ranking}
              media={result.i2.media}
              total={result.i2.totalGeral}
              separadoresAtivos={result.i2.separadoresAtivos}
              meta={meta}
            />
          </>
          );
        })()}

        {!result && (
          <div className="rounded-xl border border-dashed bg-card p-10 text-center text-sm text-muted-foreground">
            Faça upload de uma planilha e clique em <span className="font-medium text-foreground">Calcular</span> para ver os indicadores.
          </div>
        )}
      </main>
    </div>
  );
}

function ChartReport({
  kind, headerTitle, chartTitle, intervaloLabel, ranking, media, total, separadoresAtivos, meta,
}: {
  kind: "hora" | "dia";
  headerTitle: string;
  chartTitle: string;
  intervaloLabel: string;
  ranking: RankItem[];
  media: number;
  total: number;
  separadoresAtivos: number;
  meta: number;
}) {
  const mediaPorSep = separadoresAtivos > 0 ? total / separadoresAtivos : 0;
  const data = [...ranking].sort((a, b) => b.total - a.total);
  const height = Math.max(260, data.length * 28 + 120);
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
        a.download = `${headerTitle}.png`;
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
    <section className="rounded-xl border bg-card p-5 shadow-sm">
      <div className="mb-3 flex items-center justify-end">
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
      {data.length === 0 ? (
        <div className="py-12 text-center text-sm text-muted-foreground">Sem dados no intervalo.</div>
      ) : (
        <div ref={ref} style={{ background: "#ffffff", padding: 24 }}>
          <h2 style={{ textAlign: "center", fontSize: 20, fontWeight: 700, color: "#1f2937", margin: 0 }}>
            {headerTitle}
          </h2>
          <h3 style={{ fontSize: 15, fontWeight: 700, color: "#111827", marginTop: 18, marginBottom: 4 }}>
            Gráfico - Desempenho por separador (prod/h)
          </h3>
          <p style={{ textAlign: "center", fontSize: 12, color: "#374151", marginTop: 4, marginBottom: 8 }}>
            {chartTitle}
          </p>
          <div style={{ width: "100%", height }}>
            <ResponsiveContainer>
              <BarChart data={data} layout="vertical" margin={{ left: 8, right: 60, top: 8, bottom: 40 }}>
                <CartesianGrid horizontal={false} stroke="#e5e7eb" />
                <XAxis
                  type="number"
                  tick={{ fontSize: 11, fill: "#374151" }}
                  label={{ value: "Produtos por hora (prod/h)", position: "insideBottom", offset: -10, fill: "#374151", fontSize: 12 }}
                />
                <YAxis
                  type="category"
                  dataKey="separador"
                  width={210}
                  tick={{ fontSize: 10, fill: "#111827" }}
                  interval={0}
                />
                <ReferenceLine
                  x={meta}
                  stroke="#ea8c2f"
                  strokeDasharray="6 4"
                  strokeWidth={1.5}
                />
                <ReferenceLine
                  x={media}
                  stroke="#2f6fea"
                  strokeDasharray="2 4"
                  strokeWidth={1.5}
                />
                <Bar dataKey="total" radius={[0, 2, 2, 0]} barSize={16}>
                  {data.map((d, i) => (
                    <Cell key={i} fill={d.prodHora >= meta ? "#3aa84a" : "#d94a4a"} />
                  ))}
                  <LabelList
                    dataKey="total"
                    position="right"
                    content={(props: any) => {
                      const { x, y, width, height, index } = props;
                      const d = data[index];
                      if (!d) return null;
                      return (
                        <text
                          x={Number(x) + Number(width) + 6}
                          y={Number(y) + Number(height) / 2}
                          fill="#111827"
                          fontSize={10}
                          dominantBaseline="middle"
                        >
                          {`${fmtInt(d.total)} prod • ${d.enderecos} end.`}
                        </text>
                      );
                    }}
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 16, fontSize: 11, color: "#374151", marginTop: -8 }}>
            <span><span style={{ display: "inline-block", width: 18, borderTop: "2px dashed #ea8c2f", verticalAlign: "middle", marginRight: 6 }} />Meta: {fmt1(meta)} prod/h</span>
            <span><span style={{ display: "inline-block", width: 18, borderTop: "2px dotted #2f6fea", verticalAlign: "middle", marginRight: 6 }} />Média geral: {fmt1(media)} prod/h</span>
          </div>
          <h3 style={{ fontSize: 15, fontWeight: 700, color: "#111827", marginTop: 18, marginBottom: 8 }}>
            {kind === "hora" ? "Resumo por hora" : "Resumo por dia"}
          </h3>
          <table style={{ borderCollapse: "collapse", fontSize: 11, color: "#111827" }}>
            <thead>
              <tr>
                {["intervalo", "qt_separada_total", "qt_separadores_ativos", "qt_separada_media_por_sep"].map((h) => (
                  <th key={h} style={{ border: "1px solid #cbd5e1", padding: "6px 10px", background: "#f1f5f9", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style={{ border: "1px solid #cbd5e1", padding: "6px 10px", textAlign: "center" }}>{intervaloLabel}</td>
                <td style={{ border: "1px solid #cbd5e1", padding: "6px 10px", textAlign: "center" }}>{fmtInt(total)}</td>
                <td style={{ border: "1px solid #cbd5e1", padding: "6px 10px", textAlign: "center" }}>{separadoresAtivos}</td>
                <td style={{ border: "1px solid #cbd5e1", padding: "6px 10px", textAlign: "center" }}>{fmt1(mediaPorSep)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

function fmtInt(n: number) { return new Intl.NumberFormat("pt-BR").format(Math.round(n)); }
function fmt1(n: number) { return new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(n); }
function pad(n: number) { return n.toString().padStart(2, "0"); }
function fmtDay(d: Date) { return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; }

function MultiSelect({
  options, selected, onChange, allLabel, placeholder, searchable, searchPlaceholder,
}: {
  options: string[];
  selected: string[];
  onChange: (v: string[]) => void;
  allLabel: string;
  placeholder?: string;
  searchable?: boolean;
  searchPlaceholder?: string;
}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const box = useRef<HTMLDivElement>(null);
  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return options;
    return options.filter((o) => o.toLowerCase().includes(q));
  }, [options, query]);

  useEffect(() => {
    function onDoc(e: MouseEvent) {
      if (box.current && !box.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  function toggle(v: string) {
    onChange(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v]);
  }

  const label = selected.length === 0
    ? (placeholder ?? allLabel)
    : (options.length > 0 && selected.length === options.length ? allLabel : selected.join(", "));

  return (
    <div ref={box} className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex h-10 w-full items-center justify-between gap-2 rounded-md border bg-background px-3 text-left text-sm hover:bg-accent"
      >
        <span className="truncate">{label}</span>
        <span className="text-muted-foreground">▾</span>
      </button>
      {open && (
        <div className="absolute z-20 mt-1 max-h-60 w-full overflow-auto rounded-md border bg-popover p-1 shadow-md">
          <div className="sticky top-0 z-10 border-b bg-popover pb-1 mb-1">
            {searchable && (
              <input
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder={searchPlaceholder ?? "Buscar…"}
                className="mb-1 h-8 w-full rounded border bg-background px-2 text-sm"
              />
            )}
            <div className="flex gap-1">
              <button
                type="button"
                onClick={() => onChange(options)}
                className="flex-1 rounded px-2 py-1.5 text-left text-xs font-medium text-foreground hover:bg-accent"
              >
                {allLabel}
              </button>
              <button
                type="button"
                onClick={() => onChange([])}
                className="rounded px-2 py-1.5 text-left text-xs text-muted-foreground hover:bg-accent"
              >
                Limpar seleção
              </button>
            </div>
          </div>

          {visible.length === 0 && (
            <div className="px-2 py-3 text-center text-xs text-muted-foreground">Nenhum resultado</div>
          )}
          {visible.map((o) => (
            <label key={o} className="flex cursor-pointer items-center gap-2 rounded px-2 py-1.5 text-sm hover:bg-accent">
              <input
                type="checkbox"
                checked={selected.includes(o)}
                onChange={() => toggle(o)}
                className="h-4 w-4"
              />
              <span className="truncate">{o}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
}
