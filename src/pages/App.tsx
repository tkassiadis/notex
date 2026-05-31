// ============================================================
// src/pages/App.tsx
// Componente raiz. Preserva 100% do visual e UX original.
//
// O que mudou em relação ao original:
//   1. useState(rawItems) + localStorage  →  useAtividades (Supabase)
//   2. handleSave/handleDelete diretos    →  addAtividade/updateAtividade/deleteAtividade
//   3. onImport do ImportPanel            →  importAtividades
//   4. exportToExcel                      →  importado de lib/xlsx.ts
//   5. getDisciplineStats                 →  importado de lib/calculos.ts (recebe meta)
//   6. Adicionado: useAuth (sessão) + tela de Login
//   7. Adicionado: botão de logout no header (discreto)
//   8. Adicionado: indicador de erro de rede (banner não-intrusivo)
//
// O que NÃO mudou:
//   - Todos os componentes visuais
//   - Todas as tabs e modais
//   - Toda a lógica de cálculo
//   - Todo o visual e layout
// ============================================================

import { useState, useCallback, useMemo } from "react";
import { PieChart, Pie, Cell, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

import { useAuth } from "../hooks/useAuth";
import { useAtividades } from "../hooks/useAtividades";
import { useProfile } from "../hooks/useProfile";
import { enrichItem, getDisciplineStats, calcNota, calcDaysRemaining } from "../lib/calculos";
import { exportToExcel, parseXlsxFile } from "../lib/xlsx";
import { Login } from "./Login";
import type { Atividade, AtividadeEnriquecida, DisciplinaStats } from "../types";

// ─────────────────────────────────────────────────────────────
// Constantes — PRESERVADAS DO ORIGINAL
// ─────────────────────────────────────────────────────────────
const STATUS_OPTIONS = ["Não iniciado", "Estudo inicial", "Estudo médio", "Estudo avançado", "Finalizado"];
const AVALIACAO_OPTIONS = ["AP1", "AP2", "AS", "AF", "Trabalho"];
const STATUS_COLORS: Record<string, string> = { "Não iniciado": "#64748b", "Estudo inicial": "#f59e0b", "Estudo médio": "#f97316", "Estudo avançado": "#06b6d4", "Finalizado": "#10b981" };
const DISC_COLORS = ["#6366f1", "#8b5cf6", "#ec4899", "#f43f5e", "#f97316", "#eab308", "#10b981", "#06b6d4"];

// ─────────────────────────────────────────────────────────────
// Componentes UI — PRESERVADOS INTEGRALMENTE DO ORIGINAL
// ─────────────────────────────────────────────────────────────
function Badge({ children, color = "slate" }: { children: React.ReactNode; color?: string }) {
  const map: Record<string, string> = { green: "bg-emerald-500/20 text-emerald-400 border border-emerald-500/30", red: "bg-red-500/20 text-red-400 border border-red-500/30", yellow: "bg-amber-500/20 text-amber-400 border border-amber-500/30", cyan: "bg-cyan-500/20 text-cyan-400 border border-cyan-500/30", slate: "bg-slate-500/20 text-slate-400 border border-slate-500/30", indigo: "bg-indigo-500/20 text-indigo-400 border border-indigo-500/30", orange: "bg-orange-500/20 text-orange-400 border border-orange-500/30" };
  return <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium ${map[color] || map.slate}`}>{children}</span>;
}

function StatusBadge({ status }: { status: string }) {
  const map: Record<string, string> = { "Não iniciado": "slate", "Estudo inicial": "yellow", "Estudo médio": "orange", "Estudo avançado": "cyan", "Finalizado": "green" };
  return <Badge color={map[status] || "slate"}>{status}</Badge>;
}

function ProgressBar({ value, max = 100, color = "#6366f1", height = 4 }: { value: number; max?: number; color?: string; height?: number }) {
  const pct = max > 0 ? Math.min(100, (value / max) * 100) : 0;
  return (
    <div className="w-full rounded-full overflow-hidden" style={{ height, background: "rgba(255,255,255,0.07)" }}>
      <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 999, transition: "width 0.5s ease" }} />
    </div>
  );
}

function MetricCard({ label, value, sub, color = "#6366f1", icon }: { label: string; value: any; sub?: string; color?: string; icon?: string }) {
  return (
    <div className="rounded-2xl p-5 border border-white/5 flex flex-col gap-1" style={{ background: "rgba(255,255,255,0.04)" }}>
      <div className="flex items-center justify-between mb-1">
        <span className="text-xs text-slate-500 font-medium uppercase tracking-widest">{label}</span>
        {icon && <span className="text-lg">{icon}</span>}
      </div>
      <span className="text-3xl font-bold tracking-tight" style={{ color }}>{value ?? "—"}</span>
      {sub && <span className="text-xs text-slate-500 mt-1">{sub}</span>}
    </div>
  );
}

function Modal({ open, onClose, title, children }: { open: boolean; onClose: () => void; title: string; children: React.ReactNode }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4">
      <div className="absolute inset-0 bg-black/70 backdrop-blur-sm" onClick={onClose} />
      <div className="relative w-full sm:max-w-lg rounded-t-3xl sm:rounded-2xl border border-white/10 overflow-hidden" style={{ background: "#111827", maxHeight: "92vh", overflowY: "auto" }}>
        <div className="flex items-center justify-between px-6 py-4 border-b border-white/10 sticky top-0 z-10" style={{ background: "#111827" }}>
          <h3 className="text-base font-bold text-white">{title}</h3>
          <button onClick={onClose} className="w-8 h-8 rounded-full flex items-center justify-center text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.07)" }}>✕</button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

function FormField({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1.5">
      <label className="text-xs text-slate-400 font-medium uppercase tracking-wider">{label}</label>
      {children}
    </div>
  );
}

function parseDecimal(str: string): number | null {
  if (str === "" || str == null) return null;
  const n = parseFloat(String(str).replace(",", "."));
  return isNaN(n) ? null : n;
}

function ItemForm({ item, onSave, onClose, disciplines }: {
  item: Atividade | null;
  onSave: (form: Omit<Atividade, "id">) => void;
  onClose: () => void;
  disciplines: string[];
}) {
  const toStr = (v: any) => (v == null ? "" : String(v).replace(".", ","));
  const [form, setForm] = useState<any>(() => {
    const base: any = item || { avaliacao: "AP1", instrumento: "", disciplina: disciplines[0] || "", subdivisao: "", status: "Não iniciado", data: "", pesoAvaliacao: 0.2, pesoInstrumento: 1.0, pontuacaoMaxima: null, pontuacao: null, observacoes: "" };
    return { ...base, _pa: toStr(base.pesoAvaliacao), _pi: toStr(base.pesoInstrumento), _pm: toStr(base.pontuacaoMaxima), _p: toStr(base.pontuacao) };
  });
  const set = (k: string, v: any) => setForm((f: any) => ({ ...f, [k]: v }));
  const handleSave = () => onSave({ ...form, pesoAvaliacao: parseDecimal(form._pa) ?? 0, pesoInstrumento: parseDecimal(form._pi) ?? 0, pontuacaoMaxima: parseDecimal(form._pm), pontuacao: parseDecimal(form._p) });
  const cls = "w-full rounded-xl px-3 py-2.5 text-sm text-white border border-white/10 outline-none focus:border-indigo-500 transition";
  const sty = { background: "rgba(255,255,255,0.06)" };
  return (
    <div className="flex flex-col gap-4">
      <div className="grid grid-cols-2 gap-3">
        <FormField label="Avaliação"><select value={form.avaliacao} onChange={e => set("avaliacao", e.target.value)} className={cls} style={sty}>{AVALIACAO_OPTIONS.map(o => <option key={o}>{o}</option>)}</select></FormField>
        <FormField label="Status"><select value={form.status} onChange={e => set("status", e.target.value)} className={cls} style={sty}>{STATUS_OPTIONS.map(o => <option key={o}>{o}</option>)}</select></FormField>
      </div>
      <FormField label="Instrumento"><input value={form.instrumento} onChange={e => set("instrumento", e.target.value)} className={cls} style={sty} placeholder="Nome do instrumento" /></FormField>
      <FormField label="Disciplina">
        <input value={form.disciplina} onChange={e => set("disciplina", e.target.value)} list="disc-list" className={cls} style={sty} placeholder="Nome da disciplina" />
        <datalist id="disc-list">{disciplines.map(d => <option key={d} value={d} />)}</datalist>
      </FormField>
      <FormField label="Subdivisão"><input value={form.subdivisao} onChange={e => set("subdivisao", e.target.value)} className={cls} style={sty} placeholder="Ex: Pneumo e Cardio" /></FormField>
      <FormField label="Data"><input type="date" value={form.data} onChange={e => set("data", e.target.value)} className={cls} style={sty} /></FormField>
      <div className="grid grid-cols-2 gap-3">
        <FormField label="Peso Avaliação (0–1)"><input type="text" inputMode="decimal" value={form._pa} onChange={e => set("_pa", e.target.value)} className={cls} style={sty} placeholder="Ex: 0,3" /></FormField>
        <FormField label="Peso Instrumento (0–1)"><input type="text" inputMode="decimal" value={form._pi} onChange={e => set("_pi", e.target.value)} className={cls} style={sty} placeholder="Ex: 0,5" /></FormField>
      </div>
      <div className="grid grid-cols-2 gap-3">
        <FormField label="Pontuação Máxima"><input type="text" inputMode="decimal" value={form._pm} onChange={e => set("_pm", e.target.value)} className={cls} style={sty} placeholder="Ex: 10" /></FormField>
        <FormField label="Pontuação Obtida"><input type="text" inputMode="decimal" value={form._p} onChange={e => set("_p", e.target.value)} className={cls} style={sty} placeholder="Ex: 7,5" /></FormField>
      </div>
      <FormField label="Observações"><textarea value={form.observacoes} onChange={e => set("observacoes", e.target.value)} rows={2} className={cls} style={sty} placeholder="Anotações opcionais..." /></FormField>
      <div className="flex gap-2 mt-2">
        <button onClick={onClose} className="flex-1 py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10">Cancelar</button>
        <button onClick={handleSave} className="flex-1 py-3 rounded-xl text-sm font-semibold text-white" style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)" }}>Salvar</button>
      </div>
    </div>
  );
}

function SimulationPanel({ stats, onClose }: { stats: DisciplinaStats[]; onClose: () => void }) {
  const [simScores, setSimScores] = useState<Record<string, number>>({});
  const [openDisc, setOpenDisc] = useState<string | null>(stats[0]?.disciplina ?? null);
  const simStats = useMemo(() => stats.map(s => {
    let ws = 0, tw = 0;
    s.items.forEach(r => { const w = r.pesoAvaliacao * r.pesoInstrumento; tw += w; let n = calcNota(r.pontuacao, r.pontuacaoMaxima); if (n == null && simScores[r.id] != null) n = simScores[r.id]; if (n != null) ws += n * w; });
    const m = tw > 0 ? ws / tw : null;
    return { ...s, mediaSimulada: m != null ? Math.round(m * 100) / 100 : null };
  }), [stats, simScores]);
  return (
    <div className="flex flex-col gap-4">
      <p className="text-sm text-slate-400">Simule notas hipotéticas por disciplina. <span className="text-amber-400">Dados reais não serão alterados.</span></p>
      {simStats.map(s => {
        const pending = s.items.filter(it => it.pontuacao == null);
        const isOpen = openDisc === s.disciplina;
        const mc = s.mediaSimulada == null ? "#64748b" : s.mediaSimulada >= 7 ? "#10b981" : s.mediaSimulada >= 5 ? "#f59e0b" : "#ef4444";
        return (
          <div key={s.disciplina} className="rounded-2xl border border-white/8 overflow-hidden" style={{ background: "rgba(255,255,255,0.04)" }}>
            <button onClick={() => setOpenDisc(p => p === s.disciplina ? null : s.disciplina)} className="w-full flex items-center justify-between px-4 py-3 text-left">
              <div className="flex items-center gap-3 min-w-0">
                <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: s.color }} />
                <span className="text-sm font-bold text-white truncate">{s.disciplina}</span>
                {pending.length > 0 && <span className="text-xs text-slate-500 shrink-0">{pending.length} pend.</span>}
              </div>
              <div className="flex items-center gap-2 shrink-0 ml-2">
                <div className="text-right">
                  {s.mediaAtual != null && <p className="text-xs text-slate-500">Atual: {s.mediaAtual.toFixed(2)}</p>}
                  <p className="text-sm font-bold" style={{ color: mc }}>{s.mediaSimulada != null ? `→ ${s.mediaSimulada.toFixed(2)}` : "—"}</p>
                </div>
                <Badge color={s.mediaSimulada == null ? "slate" : s.mediaSimulada >= 7 ? "green" : s.mediaSimulada >= 5 ? "yellow" : "red"}>{s.mediaSimulada == null ? "—" : s.mediaSimulada >= 7 ? "Aprovado" : "Reprovado"}</Badge>
                <span className="text-slate-500 text-xs">{isOpen ? "▲" : "▼"}</span>
              </div>
            </button>
            {isOpen && (
              <div className="border-t border-white/8 px-4 pb-4 pt-3 flex flex-col gap-3">
                {pending.length === 0 ? <p className="text-xs text-slate-500 italic">Todos os instrumentos já possuem nota lançada.</p> : pending.map(it => (
                  <div key={it.id} className="rounded-xl p-3 border border-white/8" style={{ background: "rgba(255,255,255,0.04)" }}>
                    <div className="flex items-start justify-between gap-2 mb-2">
                      <div><p className="text-sm font-semibold text-white">{it.instrumento}</p><p className="text-xs text-slate-500">{it.avaliacao} · peso {(it.pesoAvaliacao * it.pesoInstrumento * 100).toFixed(0)}%</p></div>
                      <span className="text-sm font-bold text-indigo-400 shrink-0">{simScores[it.id] != null ? simScores[it.id] : "—"}</span>
                    </div>
                    <input type="range" min="0" max="10" step="0.5" value={simScores[it.id] ?? 5} onChange={e => setSimScores(sc => ({ ...sc, [it.id]: parseFloat(e.target.value) }))} className="w-full" />
                    <div className="flex justify-between text-xs text-slate-600 mt-0.5"><span>0</span><span>5</span><span>10</span></div>
                  </div>
                ))}
                <div className="rounded-xl p-3 border border-white/8 flex items-center justify-between" style={{ background: "rgba(99,102,241,0.07)" }}>
                  <span className="text-xs text-slate-400">Média simulada</span>
                  <div className="flex items-center gap-2">
                    {s.mediaAtual != null && <span className="text-xs text-slate-500">{s.mediaAtual.toFixed(2)} atual</span>}
                    <span className="text-sm font-bold" style={{ color: mc }}>{s.mediaSimulada != null ? s.mediaSimulada.toFixed(2) : "—"}</span>
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })}
      <button onClick={onClose} className="w-full py-3 rounded-xl text-sm font-semibold text-white" style={{ background: "rgba(255,255,255,0.08)" }}>Fechar</button>
    </div>
  );
}

function ImportPanel({ onImport, onClose }: { onImport: (items: Omit<Atividade, "id">[]) => void; onClose: () => void }) {
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const handleFile = useCallback(async (file: File) => {
    if (!file) return;
    setStatus("Lendo arquivo...");
    try {
      const parsed = await parseXlsxFile(file);
      setStatus(`✅ ${parsed.length} atividades importadas!`);
      setTimeout(() => { onImport(parsed); onClose(); }, 1200);
    } catch (e: any) { setStatus("❌ Erro ao ler o arquivo. Verifique se é um .xlsx válido."); }
  }, [onImport, onClose]);
  return (
    <div className="flex flex-col gap-4">
      <p className="text-sm text-slate-400">Importe uma planilha Excel (.xlsx). Os dados existentes serão substituídos.</p>
      <div className={`rounded-2xl border-2 border-dashed p-8 flex flex-col items-center gap-3 transition cursor-pointer ${dragging ? "border-indigo-500" : "border-white/15"}`} style={{ background: dragging ? "rgba(99,102,241,0.08)" : "rgba(255,255,255,0.03)" }} onDragOver={e => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={e => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]); }} onClick={() => document.getElementById("xl-input")?.click()}>
        <span className="text-3xl">📂</span><p className="text-sm text-white font-medium">Arraste o arquivo ou clique para selecionar</p><p className="text-xs text-slate-500">.xlsx · Aba Sheet1</p>
        <input id="xl-input" type="file" accept=".xlsx,.xls" className="hidden" onChange={e => e.target.files?.[0] && handleFile(e.target.files[0])} />
      </div>
      {status && <div className="rounded-xl p-3 border border-white/10 text-sm text-center" style={{ background: "rgba(255,255,255,0.05)" }}>{status}</div>}
      <button onClick={onClose} className="w-full py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10">Cancelar</button>
    </div>
  );
}

function ChartHeader({ label, chartKey, expandedChart, setExpandedChart }: { label: string; chartKey: string; expandedChart: string | null; setExpandedChart: (k: string | null) => void }) {
  return (
    <div className="flex items-center justify-between mb-4">
      <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">{label}</p>
      <button onClick={() => setExpandedChart(expandedChart === chartKey ? null : chartKey)} className="px-2 py-1 rounded-lg text-xs text-slate-400 hover:text-white transition border border-white/10" style={{ background: "rgba(255,255,255,0.04)" }}>
        {expandedChart === chartKey ? "▲ Recolher" : "▼ Expandir"}
      </button>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// DashboardTab — PRESERVADO DO ORIGINAL
// ─────────────────────────────────────────────────────────────
function DashboardTab({ items, stats }: { items: AtividadeEnriquecida[]; stats: DisciplinaStats[] }) {
  const [expandedChart, setExpandedChart] = useState<string | null>(null);
  const upcomingAll = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 30).sort((a, b) => (a.daysRemaining ?? 0) - (b.daysRemaining ?? 0)), [items]);
  const aguardandoCorrecaoAll = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining < 0 && it.pontuacao == null), [items]);
  const totalFinished = items.filter(it => it.status === "Finalizado").length;
  const totalInProgress = items.filter(it => ["Estudo inicial", "Estudo médio", "Estudo avançado"].includes(it.status)).length;
  const totalPending = items.filter(it => it.status === "Não iniciado").length;
  const avgGrade = useMemo(() => { const graded = stats.filter(s => s.mediaAtual != null); return graded.length ? (graded.reduce((a, s) => a + (s.mediaAtual ?? 0), 0) / graded.length).toFixed(2) : null; }, [stats]);
  const statusPieData = [{ name: "Finalizado", value: totalFinished, color: "#10b981" }, { name: "Em andamento", value: totalInProgress, color: "#06b6d4" }, { name: "Pendente", value: totalPending, color: "#64748b" }].filter(d => d.value > 0);
  const statusPieTotal = statusPieData.reduce((a, d) => a + d.value, 0);
  const pendingByStatus = useMemo(() => { const pending = items.filter(it => it.pontuacao == null); const groups: Record<string, AtividadeEnriquecida[]> = {}; pending.forEach(it => { if (!groups[it.status]) groups[it.status] = []; groups[it.status].push(it); }); return STATUS_OPTIONS.map(s => ({ status: s, items: groups[s] || [], color: STATUS_COLORS[s] })).filter(g => g.items.length > 0); }, [items]);
  const pontosChartData = useMemo(() => stats.map((s, idx) => {
    let obtidos = 0, futuros = 0;
    s.items.forEach(r => { const contrib = r.pesoAvaliacao * r.pesoInstrumento * 10; if (r.pontuacao != null && r.pontuacaoMaxima) { const nota = (r.pontuacao / r.pontuacaoMaxima) * 10; obtidos += (nota / 10) * contrib; } else { futuros += contrib; } });
    return { name: s.disciplina.length > 14 ? s.disciplina.split(" ").slice(0, 2).join(" ") : s.disciplina, fullName: s.disciplina, obtidos: Math.round(obtidos * 100) / 100, futuros: Math.round(futuros * 100) / 100, fill: s.color };
  }), [stats]);

  return (
    <div className="flex flex-col gap-6">
      <div className="grid grid-cols-2 gap-3">
        <MetricCard label="Média Geral" value={avgGrade ?? "—"} sub="Ponderada" color="#6366f1" icon="📊" />
        <MetricCard label="Próximas (30d)" value={upcomingAll.length} sub="atividades" color="#f59e0b" icon="⏰" />
      </div>
      {stats.filter(s => s.emRisco).length > 0 && (
        <div className="rounded-2xl p-4 border border-red-500/30" style={{ background: "rgba(239,68,68,0.08)" }}>
          <div className="flex items-center gap-2 mb-3"><span>⚠️</span><p className="text-sm font-bold text-red-400">Disciplinas em Risco</p></div>
          {stats.filter(s => s.emRisco).map(s => (<div key={s.disciplina} className="flex items-center justify-between"><span className="text-sm text-white">{s.disciplina}</span><span className="text-sm font-bold text-red-400">{s.mediaAtual?.toFixed(2)}</span></div>))}
        </div>
      )}
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <ChartHeader label="Pontos na Média Final (de 10)" chartKey="grades" expandedChart={expandedChart} setExpandedChart={setExpandedChart} />
        <p className="text-xs text-slate-600 mb-4 -mt-2">Pontos já conquistados + pontos ainda disponíveis em avaliações futuras</p>
        <div className="flex flex-col gap-3">
          {pontosChartData.map((d, i) => {
            const s = stats[i]; const obtPct = (d.obtidos / 10) * 100; const futPct = (d.futuros / 10) * 100;
            return (
              <div key={d.fullName}>
                <div className="flex items-center justify-between mb-1.5">
                  <span className="text-xs text-slate-300 font-semibold truncate pr-2" style={{ maxWidth: "55%" }}>{d.fullName}</span>
                  <div className="flex items-center gap-2 shrink-0"><span className="text-xs font-bold" style={{ color: s.color }}>{d.obtidos.toFixed(2)}</span><span className="text-xs text-slate-600">+</span><span className="text-xs text-slate-400">{d.futuros.toFixed(2)}</span></div>
                </div>
                <div className="relative w-full rounded-lg overflow-hidden flex" style={{ height: 22, background: "rgba(255,255,255,0.05)" }}>
                  {obtPct > 0 && (<div className="flex items-center justify-end pr-1.5 relative" style={{ width: `${obtPct}%`, background: s.color, minWidth: d.obtidos > 0 ? 2 : 0, transition: "width 0.6s ease" }}>{obtPct > 12 && (<span className="text-xs font-bold text-white drop-shadow" style={{ fontSize: 10 }}>{d.obtidos.toFixed(1)}</span>)}</div>)}
                  {obtPct > 0 && futPct > 0 && (<div style={{ width: 2, background: "#0a0f1a", flexShrink: 0 }} />)}
                  {futPct > 0 && (<div className="flex items-center pl-1.5" style={{ width: `${futPct}%`, background: "rgba(255,255,255,0.12)", minWidth: d.futuros > 0 ? 2 : 0 }}>{futPct > 12 && (<span className="text-xs text-slate-400" style={{ fontSize: 10 }}>{d.futuros.toFixed(1)}</span>)}</div>)}
                </div>
                {i === pontosChartData.length - 1 && (<div className="flex justify-between mt-1.5 px-0">{[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(n => (<span key={n} className="text-xs" style={{ color: n === 7 ? "#f59e0b" : "#475569", fontSize: 9, fontWeight: n === 0 || n === 10 || n === 7 ? 700 : 400 }}>{n}</span>))}</div>)}
              </div>
            );
          })}
        </div>
        <div className="flex items-center gap-4 mt-4 pt-3 border-t border-white/5">
          <div className="flex items-center gap-1.5"><div className="w-10 h-3 rounded-sm" style={{ background: "linear-gradient(90deg,#6366f1,#8b5cf6)" }} /><span className="text-xs text-slate-400">Conquistados</span></div>
          <div className="flex items-center gap-1.5"><div className="w-10 h-3 rounded-sm" style={{ background: "rgba(255,255,255,0.14)", border: "1px solid rgba(255,255,255,0.1)" }} /><span className="text-xs text-slate-400">Disponíveis</span></div>
          <div className="flex items-center gap-1.5 ml-auto"><div className="w-1 h-3" style={{ background: "#f59e0b" }} /><span className="text-xs text-amber-400">7 = meta</span></div>
        </div>
        {expandedChart === "grades" && (<div className="mt-4 flex flex-col gap-3 border-t border-white/8 pt-4">{pontosChartData.map((d, i) => { const s = stats[i]; return (<div key={d.fullName}><div className="flex items-center justify-between mb-1"><span className="text-xs text-slate-300 font-semibold">{d.fullName}</span><div className="flex items-center gap-3"><span className="text-xs text-slate-500">disponível: +{d.futuros.toFixed(2)}</span><span className="text-sm font-bold" style={{ color: d.obtidos >= 7 ? "#10b981" : d.obtidos >= 5 ? "#f59e0b" : "#ef4444" }}>{d.obtidos.toFixed(2)}/10</span></div></div>{s.notaNecessaria != null && s.notaNecessaria > 0 && (<p className="text-xs text-slate-500">Precisa de <span className="font-semibold" style={{ color: s.notaNecessaria > 10 ? "#ef4444" : "#f59e0b" }}>{s.notaNecessaria.toFixed(2)}</span> nas próximas para fechar 7</p>)}</div>); })}</div>)}
      </div>
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <ChartHeader label="Status dos Estudos" chartKey="status" expandedChart={expandedChart} setExpandedChart={setExpandedChart} />
        <div className="flex items-center gap-4">
          <ResponsiveContainer width={120} height={120}><PieChart><Pie data={statusPieData} dataKey="value" innerRadius={35} outerRadius={55} paddingAngle={3}>{statusPieData.map((e, i) => <Cell key={i} fill={e.color} />)}</Pie></PieChart></ResponsiveContainer>
          <div className="flex flex-col gap-2.5 flex-1">{statusPieData.map(d => (<div key={d.name} className="flex items-center justify-between"><div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: d.color }} /><span className="text-xs text-slate-400">{d.name}</span></div><div className="flex items-center gap-2"><span className="text-xs text-slate-500">{statusPieTotal > 0 ? Math.round(d.value / statusPieTotal * 100) : 0}%</span><span className="text-sm font-bold text-white w-5 text-right">{d.value}</span></div></div>))}</div>
        </div>
        {expandedChart === "status" && (<div className="mt-4 flex flex-col gap-4 border-t border-white/8 pt-4"><p className="text-xs text-slate-500">Instrumentos pendentes por status de estudo:</p>{pendingByStatus.map(group => (<div key={group.status}><div className="flex items-center gap-2 mb-2"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: group.color }} /><span className="text-xs font-bold text-white">{group.status}</span><span className="text-xs text-slate-500">({group.items.length})</span></div><div className="flex flex-col gap-1 pl-4">{group.items.map(it => (<div key={it.id} className="flex items-center justify-between py-1 border-b border-white/5"><div className="min-w-0 flex-1"><p className="text-xs text-slate-300 truncate">{it.instrumento}</p><p className="text-xs text-slate-600">{it.disciplina} · {it.avaliacao}</p></div><div className="shrink-0 ml-2">{it.daysRemaining != null && <span className="text-xs" style={{ color: it.daysRemaining < 0 ? "#f59e0b" : it.daysRemaining <= 7 ? "#f59e0b" : "#64748b" }}>{it.daysRemaining < 0 ? `${Math.abs(it.daysRemaining)}d atrás` : `em ${it.daysRemaining}d`}</span>}</div></div>))}</div></div>))}</div>)}
      </div>
      {upcomingAll.length > 0 && (<div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold mb-4">Próximas Atividades</p><div className="flex flex-col gap-2">{upcomingAll.slice(0, 6).map(it => (<div key={it.id} className="flex items-center gap-3 py-2 border-b border-white/5"><div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0" style={{ background: (it.daysRemaining ?? 99) <= 3 ? "rgba(239,68,68,0.15)" : (it.daysRemaining ?? 99) <= 7 ? "rgba(245,158,11,0.15)" : "rgba(99,102,241,0.15)" }}><span className={`text-xs font-bold ${(it.daysRemaining ?? 99) <= 3 ? "text-red-400" : (it.daysRemaining ?? 99) <= 7 ? "text-amber-400" : "text-indigo-400"}`}>{it.daysRemaining}d</span></div><div className="flex-1 min-w-0"><p className="text-sm font-medium text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {it.avaliacao}</p></div><StatusBadge status={it.status} /></div>))}</div></div>)}
      {aguardandoCorrecaoAll.length > 0 && (<div className="rounded-2xl border border-amber-500/20 overflow-hidden" style={{ background: "rgba(245,158,11,0.06)" }}><button onClick={() => setExpandedChart(p => p === "aguardando" ? null : "aguardando")} className="w-full flex items-center justify-between px-4 py-3 text-left"><div className="flex items-center gap-2"><span>⏳</span><span className="text-xs text-amber-400 uppercase tracking-widest font-semibold">Aguardando Correção</span><Badge color="yellow">{aguardandoCorrecaoAll.length}</Badge></div><span className="text-xs text-slate-500">{expandedChart === "aguardando" ? "▲" : "▼"}</span></button>{expandedChart === "aguardando" && (<div className="border-t border-amber-500/15 px-4 pb-4 pt-3 flex flex-col gap-1"><p className="text-xs text-slate-500 mb-2">Realizadas, mas ainda sem nota lançada pelo professor.</p>{aguardandoCorrecaoAll.map(it => (<div key={it.id} className="flex items-center justify-between py-1"><div className="min-w-0"><span className="text-sm text-white block truncate">{it.instrumento}</span><span className="text-xs text-slate-500">{it.disciplina} · {it.avaliacao}</span></div><Badge color="yellow">{Math.abs(it.daysRemaining ?? 0)}d atrás</Badge></div>))}</div>)}</div>)}
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <ChartHeader label="Ranking de Disciplinas" chartKey="ranking" expandedChart={expandedChart} setExpandedChart={setExpandedChart} />
        <div className="flex flex-col gap-3">{[...stats].sort((a, b) => (b.mediaAtual ?? -1) - (a.mediaAtual ?? -1)).map((s, i) => (<div key={s.disciplina}><div className="flex items-center justify-between mb-1"><div className="flex items-center gap-2"><span className="text-xs text-slate-600 w-4">#{i + 1}</span><span className="text-sm text-white">{s.disciplina}</span></div><div className="flex items-center gap-2"><span className="text-sm font-bold" style={{ color: s.color }}>{s.mediaAtual?.toFixed(2) ?? "—"}</span>{s.emRisco && <Badge color="red">Risco</Badge>}</div></div><ProgressBar value={s.mediaAtual ?? 0} max={10} color={s.color} height={3} />{expandedChart === "ranking" && (<div className="mt-2 flex flex-col gap-1 pl-5 border-l-2" style={{ borderColor: s.color + "44" }}>{s.items.map(it => { const n = calcNota(it.pontuacao, it.pontuacaoMaxima); return (<div key={it.id} className="flex items-center justify-between py-0.5"><div className="flex items-center gap-2 min-w-0"><div className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: STATUS_COLORS[it.status] }} /><span className="text-xs text-slate-400 truncate">{it.instrumento}</span><span className="text-xs text-slate-600 shrink-0">{it.avaliacao}</span></div><span className="text-xs font-semibold shrink-0 ml-2" style={{ color: n == null ? "#64748b" : n >= 7 ? "#10b981" : n >= 5 ? "#f59e0b" : "#ef4444" }}>{n != null ? n.toFixed(1) : "—"}</span></div>); })}</div>)}</div>))}</div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// DisciplineTab — PRESERVADO DO ORIGINAL
// ─────────────────────────────────────────────────────────────
function DisciplineTab({ stats, onEditItem, onDeleteItem }: { stats: DisciplinaStats[]; onEditItem: (item: Atividade) => void; onDeleteItem: (id: string) => void }) {
  const [selected, setSelected] = useState<string | null>(null);
  const disc = selected ? stats.find(s => s.disciplina === selected) : null;
  if (disc) return (
    <div className="flex flex-col gap-4">
      <button onClick={() => setSelected(null)} className="flex items-center gap-2 text-sm text-slate-400 hover:text-white transition"><span>←</span> Todas as disciplinas</button>
      <div className="rounded-2xl p-5 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <div className="flex items-center gap-3 mb-4"><div className="w-10 h-10 rounded-xl flex items-center justify-center text-lg" style={{ background: `${disc.color}22` }}>📖</div><div><h2 className="text-lg font-bold text-white">{disc.disciplina}</h2><p className="text-xs text-slate-500">{disc.items.length} instrumentos</p></div></div>
        <div className="grid grid-cols-2 gap-3 mb-4">
          <div className="rounded-xl p-3 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><p className="text-xs text-slate-500 mb-1">Média Atual</p><p className="text-2xl font-bold" style={{ color: disc.color }}>{disc.mediaAtual?.toFixed(2) ?? "—"}</p></div>
          <div className="rounded-xl p-3 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><p className="text-xs text-slate-500 mb-1">Nota Necessária</p><p className="text-2xl font-bold" style={{ color: (disc.notaNecessaria ?? 0) > 10 ? "#ef4444" : (disc.notaNecessaria ?? 0) > 7 ? "#f59e0b" : "#10b981" }}>{disc.notaNecessaria != null ? disc.notaNecessaria.toFixed(2) : "—"}</p></div>
        </div>
        <div className="mb-2"><div className="flex items-center justify-between mb-1"><span className="text-xs text-slate-500">Progresso concluído</span><span className="text-xs text-slate-400">{disc.pesoConcluido.toFixed(1)}%</span></div><ProgressBar value={disc.pesoConcluido} max={100} color={disc.color} height={6} /></div>
        {disc.notaNecessaria != null && (<div className="mt-3 p-3 rounded-xl border" style={{ background: "rgba(99,102,241,0.08)", borderColor: "rgba(99,102,241,0.2)" }}><p className="text-xs text-indigo-300">{disc.notaNecessaria <= 0 ? "✅ Média 7 já garantida!" : disc.notaNecessaria > 10 ? "❌ Média 7 não é mais matematicamente possível." : `📌 Você precisa de média ${disc.notaNecessaria.toFixed(2)} nas próximas avaliações para fechar com 7.`}</p></div>)}
      </div>
      <div className="flex flex-col gap-2">{disc.items.map(it => { const nota = calcNota(it.pontuacao, it.pontuacaoMaxima); return (<div key={it.id} className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><div className="flex items-start justify-between gap-2 mb-2"><div className="flex-1 min-w-0"><p className="text-sm font-semibold text-white leading-tight">{it.instrumento}</p><div className="flex items-center gap-2 mt-1 flex-wrap"><Badge color="indigo">{it.avaliacao}</Badge>{it.subdivisao && <Badge color="slate">{it.subdivisao}</Badge>}<StatusBadge status={it.status} /></div></div><div className="flex gap-1 shrink-0"><button onClick={() => onEditItem(it)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }}>✏️</button><button onClick={() => onDeleteItem(it.id)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs text-red-400 hover:text-red-300 transition" style={{ background: "rgba(239,68,68,0.1)" }}>🗑</button></div></div><div className="grid grid-cols-3 gap-2 text-center mt-2"><div><p className="text-xs text-slate-600">Nota/10</p><p className="text-sm font-bold" style={{ color: nota != null && nota >= 7 ? "#10b981" : nota != null && nota >= 5 ? "#f59e0b" : nota != null ? "#ef4444" : "#64748b" }}>{nota != null ? nota.toFixed(2) : "—"}</p></div><div><p className="text-xs text-slate-600">Peso Aval.</p><p className="text-sm font-bold text-slate-300">{(it.pesoAvaliacao * 100).toFixed(0)}%</p></div><div><p className="text-xs text-slate-600">Data</p><p className="text-xs font-semibold" style={{ color: it.daysRemaining != null && it.daysRemaining < 0 ? "#ef4444" : it.daysRemaining != null && it.daysRemaining <= 7 ? "#f59e0b" : "#94a3b8" }}>{it.daysRemaining != null && it.daysRemaining < 0 ? `${Math.abs(it.daysRemaining)}d atrás` : it.daysRemaining != null ? `${it.daysRemaining}d` : it.data || "—"}</p></div></div>{it.pontuacao != null && it.pontuacaoMaxima != null && <div className="mt-2"><ProgressBar value={it.pontuacao} max={it.pontuacaoMaxima} color={nota != null && nota >= 7 ? "#10b981" : nota != null && nota >= 5 ? "#f59e0b" : "#ef4444"} height={4} /><p className="text-xs text-slate-600 mt-1 text-right">{it.pontuacao}/{it.pontuacaoMaxima} pts</p></div>}{it.observacoes && <p className="text-xs text-slate-500 mt-2 italic">{it.observacoes}</p>}</div>); })}</div>
    </div>
  );
  return (
    <div className="flex flex-col gap-4">
      <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">Selecione uma disciplina</p>
      {stats.map(s => (<button key={s.disciplina} onClick={() => setSelected(s.disciplina)} className="rounded-2xl p-5 border border-white/5 text-left w-full hover:border-white/15 transition-all" style={{ background: "rgba(255,255,255,0.04)" }}><div className="flex items-center justify-between mb-3"><div className="flex items-center gap-3"><div className="w-3 h-3 rounded-full shrink-0" style={{ background: s.color }} /><p className="text-sm font-bold text-white">{s.disciplina}</p></div><div className="flex items-center gap-2">{s.emRisco && <Badge color="red">Risco</Badge>}<span className="text-lg font-bold" style={{ color: s.color }}>{s.mediaAtual?.toFixed(2) ?? "—"}</span></div></div><ProgressBar value={s.pesoConcluido} max={100} color={s.color} height={5} /><div className="flex items-center justify-between mt-2"><span className="text-xs text-slate-500">{s.pesoConcluido.toFixed(1)}% concluído</span><div className="flex gap-2"><span className="text-xs text-emerald-400">{s.statusCounts["Finalizado"]} ok</span><span className="text-xs text-cyan-400">{s.statusCounts["Em andamento"]} em estudo</span><span className="text-xs text-slate-500">{s.statusCounts["Não iniciado"]} pendentes</span></div></div></button>))}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// AllItemsTab — PRESERVADO DO ORIGINAL
// ─────────────────────────────────────────────────────────────
function AllItemsTab({ items, stats, onEditItem, onDeleteItem }: { items: AtividadeEnriquecida[]; stats: DisciplinaStats[]; onEditItem: (item: Atividade) => void; onDeleteItem: (id: string) => void }) {
  const [filter, setFilter] = useState(""); const [statusFilter, setStatusFilter] = useState("Todos"); const [sort, setSort] = useState("data-asc");
  const notaNecMap = useMemo(() => { const m: Record<string, number> = {}; stats.forEach(s => { m[s.disciplina] = s.notaNecessaria ?? 0; }); return m; }, [stats]);
  const SORT_OPTIONS = [{ value: "data-asc", label: "Data ↑" }, { value: "data-desc", label: "Data ↓" }, { value: "nota-desc", label: "Nota nec. ↑" }, { value: "nota-asc", label: "Nota nec. ↓" }];
  const filtered = useMemo(() => { const list = items.filter(it => { const q = filter.toLowerCase(); return (!q || it.disciplina.toLowerCase().includes(q) || it.instrumento.toLowerCase().includes(q) || it.avaliacao.toLowerCase().includes(q)) && (statusFilter === "Todos" || it.status === statusFilter); }); return [...list].sort((a, b) => { if (sort === "data-asc") return (a.daysRemaining ?? 9999) - (b.daysRemaining ?? 9999); if (sort === "data-desc") return (b.daysRemaining ?? -9999) - (a.daysRemaining ?? -9999); if (sort === "nota-desc") return (notaNecMap[b.disciplina] ?? 0) - (notaNecMap[a.disciplina] ?? 0); if (sort === "nota-asc") return (notaNecMap[a.disciplina] ?? 0) - (notaNecMap[b.disciplina] ?? 0); return 0; }); }, [items, filter, statusFilter, sort, notaNecMap]);
  return (
    <div className="flex flex-col gap-4">
      <input value={filter} onChange={e => setFilter(e.target.value)} placeholder="Buscar disciplina, instrumento..." className="w-full rounded-xl px-4 py-3 text-sm text-white border border-white/10 outline-none focus:border-indigo-500 transition" style={{ background: "rgba(255,255,255,0.05)" }} />
      <div className="flex gap-2 overflow-x-auto pb-1">{["Todos", ...STATUS_OPTIONS].map(s => <button key={s} onClick={() => setStatusFilter(s)} className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${statusFilter === s ? "text-white" : "text-slate-400 border border-white/10"}`} style={statusFilter === s ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>{s}</button>)}</div>
      <div className="flex items-center gap-2"><span className="text-xs text-slate-500 shrink-0">Ordenar:</span><div className="flex gap-1.5 overflow-x-auto pb-0.5">{SORT_OPTIONS.map(o => <button key={o.value} onClick={() => setSort(o.value)} className={`shrink-0 px-2.5 py-1 rounded-lg text-xs font-medium transition ${sort === o.value ? "text-white border" : "text-slate-400 border border-white/10"}`} style={sort === o.value ? { background: "rgba(99,102,241,0.4)", borderColor: "rgba(99,102,241,0.5)" } : {}}>{o.label}</button>)}</div></div>
      <p className="text-xs text-slate-500">{filtered.length} atividades</p>
      <div className="flex flex-col gap-2">{filtered.map(it => { const nota = calcNota(it.pontuacao, it.pontuacaoMaxima); return (<div key={it.id} className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><div className="flex items-start justify-between gap-2"><div className="flex-1 min-w-0"><p className="text-sm font-semibold text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500 mt-0.5">{it.disciplina} · {it.avaliacao}</p><div className="flex gap-1 mt-1.5 flex-wrap"><StatusBadge status={it.status} />{it.daysRemaining != null && <Badge color={it.daysRemaining < 0 && it.pontuacao == null ? "yellow" : it.daysRemaining < 0 ? "slate" : it.daysRemaining <= 7 ? "yellow" : "slate"}>{it.daysRemaining < 0 ? (it.pontuacao == null ? `Aguard. correção (${Math.abs(it.daysRemaining)}d)` : `${Math.abs(it.daysRemaining)}d atrás`) : `${it.daysRemaining}d`}</Badge>}</div></div><div className="flex flex-col items-end gap-2 shrink-0"><span className="text-base font-bold" style={{ color: nota == null ? "#64748b" : nota >= 7 ? "#10b981" : nota >= 5 ? "#f59e0b" : "#ef4444" }}>{nota != null ? nota.toFixed(1) : "—"}/10</span><div className="flex gap-1"><button onClick={() => onEditItem(it)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs transition" style={{ background: "rgba(255,255,255,0.06)" }}>✏️</button><button onClick={() => onDeleteItem(it.id)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs transition" style={{ background: "rgba(239,68,68,0.1)" }}>🗑</button></div></div></div></div>); })}</div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// AlertsTab — PRESERVADO DO ORIGINAL
// ─────────────────────────────────────────────────────────────
function AlertsTab({ items, stats }: { items: AtividadeEnriquecida[]; stats: DisciplinaStats[] }) {
  const urgentes = items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 3 && it.pontuacao == null);
  const proximas = items.filter(it => it.daysRemaining != null && it.daysRemaining > 3 && it.daysRemaining <= 14 && it.pontuacao == null);
  const aguardando = items.filter(it => it.daysRemaining != null && it.daysRemaining < 0 && it.pontuacao == null);
  const emRisco = stats.filter(s => s.emRisco); const notaAlta = stats.filter(s => s.notaNecessaria != null && s.notaNecessaria > 8); const notaImpossivel = stats.filter(s => s.notaNecessaria != null && s.notaNecessaria > 10);
  const Section = ({ title, color, icon, children, count }: any) => (<div className="rounded-2xl overflow-hidden border" style={{ borderColor: color + "33", background: color + "0d" }}><div className="flex items-center gap-2 px-4 py-3 border-b" style={{ borderColor: color + "22" }}><span className="text-base">{icon}</span><span className="text-sm font-bold" style={{ color }}>{title}</span><span className="ml-auto text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: color + "22", color }}>{count}</span></div><div className="flex flex-col gap-0">{children}</div></div>);
  const AlertRow = ({ it, color }: any) => (<div className="flex items-center justify-between px-4 py-2.5 border-b border-white/4 last:border-0"><div className="min-w-0"><p className="text-sm text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {it.avaliacao}</p></div><span className="text-xs font-bold shrink-0 ml-3" style={{ color }}>{it.daysRemaining < 0 ? `${Math.abs(it.daysRemaining)}d atrás` : it.daysRemaining === 0 ? "Hoje" : `${it.daysRemaining}d`}</span></div>);
  const hasAlerts = urgentes.length || proximas.length || aguardando.length || emRisco.length || notaAlta.length;
  return (
    <div className="flex flex-col gap-4">
      <div className="flex items-center justify-between"><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">Central de Alertas</p>{hasAlerts ? <Badge color="red">{urgentes.length + proximas.length + emRisco.length + notaAlta.length} alertas</Badge> : <Badge color="green">Tudo ok</Badge>}</div>
      {!hasAlerts && (<div className="rounded-2xl p-8 border border-white/5 flex flex-col items-center gap-3" style={{ background: "rgba(255,255,255,0.03)" }}><span className="text-4xl">✅</span><p className="text-sm font-semibold text-white">Nenhum alerta ativo</p><p className="text-xs text-slate-500 text-center">Todas as disciplinas estão em dia e sem atividades urgentes.</p></div>)}
      {notaImpossivel.length > 0 && (<Section title="Aprovação Matematicamente Impossível" color="#ef4444" icon="❌" count={notaImpossivel.length}>{notaImpossivel.map((s: DisciplinaStats) => (<div key={s.disciplina} className="flex items-center justify-between px-4 py-2.5 border-b border-white/4 last:border-0"><div><p className="text-sm text-white">{s.disciplina}</p><p className="text-xs text-slate-500">Média atual: {s.mediaAtual?.toFixed(2) ?? "—"}</p></div><Badge color="red">Impossível</Badge></div>))}</Section>)}
      {urgentes.length > 0 && (<Section title="Urgente — Próximas 72 horas" color="#ef4444" icon="🚨" count={urgentes.length}>{urgentes.map((it: AtividadeEnriquecida) => <AlertRow key={it.id} it={it} color="#ef4444" />)}</Section>)}
      {emRisco.length > 0 && (<Section title="Disciplinas em Risco (média < 6)" color="#f43f5e" icon="⚠️" count={emRisco.length}>{emRisco.map((s: DisciplinaStats) => (<div key={s.disciplina} className="flex items-center justify-between px-4 py-2.5 border-b border-white/4 last:border-0"><div><p className="text-sm text-white">{s.disciplina}</p><p className="text-xs text-slate-500">Nota necessária: {s.notaNecessaria?.toFixed(2) ?? "—"}</p></div><span className="text-sm font-bold text-red-400">{s.mediaAtual?.toFixed(2) ?? "—"}</span></div>))}</Section>)}
      {notaAlta.filter((s: DisciplinaStats) => (s.notaNecessaria ?? 0) <= 10).length > 0 && (<Section title="Nota Necessária Alta (> 8)" color="#f97316" icon="📈" count={notaAlta.filter((s: DisciplinaStats) => (s.notaNecessaria ?? 0) <= 10).length}>{notaAlta.filter((s: DisciplinaStats) => (s.notaNecessaria ?? 0) <= 10).map((s: DisciplinaStats) => (<div key={s.disciplina} className="flex items-center justify-between px-4 py-2.5 border-b border-white/4 last:border-0"><div><p className="text-sm text-white">{s.disciplina}</p><p className="text-xs text-slate-500">Média atual: {s.mediaAtual?.toFixed(2) ?? "—"}</p></div><span className="text-sm font-bold text-orange-400">{s.notaNecessaria!.toFixed(2)} nec.</span></div>))}</Section>)}
      {proximas.length > 0 && (<Section title="Próximas — Em até 14 dias" color="#f59e0b" icon="⏰" count={proximas.length}>{proximas.map((it: AtividadeEnriquecida) => <AlertRow key={it.id} it={it} color="#f59e0b" />)}</Section>)}
      {aguardando.length > 0 && (<Section title="Aguardando Correção" color="#64748b" icon="⏳" count={aguardando.length}>{aguardando.map((it: AtividadeEnriquecida) => <AlertRow key={it.id} it={it} color="#94a3b8" />)}</Section>)}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// CalendarTab — PRESERVADO DO ORIGINAL
// ─────────────────────────────────────────────────────────────
function CalendarTab({ items }: { items: AtividadeEnriquecida[] }) {
  const today = new Date(); const [viewDate, setViewDate] = useState(new Date(today.getFullYear(), today.getMonth(), 1));
  const year = viewDate.getFullYear(); const month = viewDate.getMonth();
  const monthNames = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  const dayNames = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  const firstDay = new Date(year, month, 1).getDay(); const daysInMonth = new Date(year, month + 1, 0).getDate();
  const byDate = useMemo(() => { const m: Record<string, AtividadeEnriquecida[]> = {}; items.forEach(it => { if (!it.data) return; if (!m[it.data]) m[it.data] = []; m[it.data].push(it); }); return m; }, [items]);
  const [selected, setSelected] = useState<number | null>(null);
  const cells: (number | null)[] = []; for (let i = 0; i < firstDay; i++) cells.push(null); for (let d = 1; d <= daysInMonth; d++) cells.push(d);
  const getDateStr = (d: number) => `${year}-${String(month + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  const todayStr = today.toISOString().split("T")[0]; const selectedStr = selected ? getDateStr(selected) : null; const selectedItems = selectedStr ? (byDate[selectedStr] || []) : [];
  const getDotColor = (d: number) => { const dateStr = getDateStr(d); const its = byDate[dateStr]; if (!its || its.length === 0) return null; const days = Math.round((new Date(dateStr).getTime() - today.getTime()) / 86400000); if (days < 0) return its.some(i => i.pontuacao == null) ? "#f59e0b" : "#10b981"; if (days <= 3) return "#ef4444"; if (days <= 7) return "#f97316"; return "#6366f1"; };
  return (
    <div className="flex flex-col gap-4">
      <div className="flex items-center justify-between"><button onClick={() => setViewDate(new Date(year, month - 1, 1))} className="w-9 h-9 rounded-xl flex items-center justify-center text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }}>‹</button><p className="text-sm font-bold text-white">{monthNames[month]} {year}</p><button onClick={() => setViewDate(new Date(year, month + 1, 1))} className="w-9 h-9 rounded-xl flex items-center justify-center text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }}>›</button></div>
      <div className="rounded-2xl overflow-hidden border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <div className="grid grid-cols-7 border-b border-white/5">{dayNames.map(d => <div key={d} className="py-2 text-center text-xs font-semibold text-slate-500">{d}</div>)}</div>
        <div className="grid grid-cols-7">{cells.map((d, i) => { if (!d) return <div key={`e${i}`} className="h-12" />; const dateStr = getDateStr(d); const isToday = dateStr === todayStr; const isSelected = d === selected; const dotColor = getDotColor(d); const count = byDate[dateStr]?.length || 0; return (<button key={d} onClick={() => setSelected(d === selected ? null : d)} className="h-12 flex flex-col items-center justify-center gap-0.5 relative transition-all" style={{ background: isSelected ? "rgba(99,102,241,0.25)" : isToday ? "rgba(99,102,241,0.1)" : "transparent" }}><span className={`text-xs font-semibold ${isToday ? "text-indigo-400" : isSelected ? "text-white" : "text-slate-300"}`} style={isToday ? { fontWeight: 800 } : {}}>{d}</span>{dotColor && (<div className="flex gap-0.5">{count > 3 ? <div className="w-4 h-1 rounded-full" style={{ background: dotColor }} /> : Array.from({ length: Math.min(count, 3) }).map((_, j) => <div key={j} className="w-1 h-1 rounded-full" style={{ background: dotColor }} />)}</div>)}</button>); })}</div>
      </div>
      <div className="flex flex-wrap gap-3">{[["#ef4444", "≤ 3 dias"], ["#f97316", "≤ 7 dias"], ["#6366f1", "Futuro"], ["#f59e0b", "Aguard. correção"], ["#10b981", "Concluído"]].map(([c, l]) => (<div key={l} className="flex items-center gap-1.5"><div className="w-2 h-2 rounded-full" style={{ background: c }} /><span className="text-xs text-slate-500">{l}</span></div>))}</div>
      {selected && (<div className="rounded-2xl border border-white/5 overflow-hidden" style={{ background: "rgba(255,255,255,0.04)" }}><div className="px-4 py-3 border-b border-white/5"><p className="text-sm font-bold text-white">{selected} de {monthNames[month]}</p><p className="text-xs text-slate-500">{selectedItems.length === 0 ? "Sem atividades" : `${selectedItems.length} atividade(s)`}</p></div>{selectedItems.length === 0 ? (<div className="px-4 py-6 text-center"><p className="text-xs text-slate-500">Nenhuma atividade neste dia</p></div>) : (<div className="flex flex-col">{selectedItems.map(it => { const nota = calcNota(it.pontuacao, it.pontuacaoMaxima); const days = it.daysRemaining; const dotC = days == null ? "#64748b" : days < 0 && it.pontuacao == null ? "#f59e0b" : days < 0 ? "#10b981" : days <= 3 ? "#ef4444" : days <= 7 ? "#f97316" : "#6366f1"; return (<div key={it.id} className="flex items-center gap-3 px-4 py-3 border-b border-white/4 last:border-0"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: dotC }} /><div className="flex-1 min-w-0"><p className="text-sm text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {it.avaliacao}</p></div><div className="text-right shrink-0">{nota != null ? <p className="text-sm font-bold" style={{ color: nota >= 7 ? "#10b981" : nota >= 5 ? "#f59e0b" : "#ef4444" }}>{nota.toFixed(1)}/10</p> : <StatusBadge status={it.status} />}</div></div>); })}</div>)}</div>)}
      <div><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold mb-3">Próximas Atividades</p><div className="flex flex-col gap-2">{items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0).sort((a, b) => (a.daysRemaining ?? 0) - (b.daysRemaining ?? 0)).slice(0, 8).map(it => (<div key={it.id} className="rounded-xl px-4 py-3 border border-white/5 flex items-center gap-3" style={{ background: "rgba(255,255,255,0.03)" }}><div className="w-10 h-10 rounded-xl flex flex-col items-center justify-center shrink-0 text-center" style={{ background: (it.daysRemaining ?? 99) <= 3 ? "rgba(239,68,68,0.15)" : (it.daysRemaining ?? 99) <= 7 ? "rgba(249,115,22,0.15)" : "rgba(99,102,241,0.15)" }}><span className="text-xs font-bold leading-tight" style={{ color: (it.daysRemaining ?? 99) <= 3 ? "#ef4444" : (it.daysRemaining ?? 99) <= 7 ? "#f97316" : "#818cf8" }}>{it.daysRemaining}d</span></div><div className="flex-1 min-w-0"><p className="text-sm font-medium text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {it.data}</p></div><StatusBadge status={it.status} /></div>))}</div></div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────
// App — componente raiz
// ─────────────────────────────────────────────────────────────
export default function App() {
  const { user, loading: authLoading, signIn, signUp, signOut } = useAuth();
  const { atividades, loading: dataLoading, error, addAtividade, updateAtividade, deleteAtividade, importAtividades } = useAtividades(user);
  const { profile } = useProfile(user);

  const [tab, setTab] = useState("dashboard");
  const [modal, setModal] = useState<string | null>(null);
  const [editingItem, setEditingItem] = useState<Atividade | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);

  const meta = profile?.metaAprovacao ?? 7;
  const items = useMemo(() => atividades.map(enrichItem), [atividades]);
  const stats = useMemo(() => getDisciplineStats(items, meta), [items, meta]);
  const disciplines = useMemo(() => [...new Set(items.map(it => it.disciplina))], [items]);

  const handleSave = useCallback(async (form: Omit<Atividade, "id">) => {
    if (editingItem) {
      await updateAtividade({ ...form, id: editingItem.id });
    } else {
      await addAtividade(form);
    }
    setModal(null);
    setEditingItem(null);
  }, [editingItem, addAtividade, updateAtividade]);

  const handleEdit = useCallback((item: Atividade) => { setEditingItem(item); setModal("edit"); }, []);
  const handleDelete = useCallback((id: string) => { setDeleteId(id); setModal("delete"); }, []);
  const confirmDelete = useCallback(async () => {
    if (deleteId) await deleteAtividade(deleteId);
    setModal(null);
    setDeleteId(null);
  }, [deleteId, deleteAtividade]);

  const TABS = [
    { id: "dashboard", label: "Dashboard", icon: "⚡" },
    { id: "disciplines", label: "Disciplinas", icon: "📖" },
    { id: "all", label: "Atividades", icon: "📋" },
    { id: "alerts", label: "Alertas", icon: "🔔" },
    { id: "calendar", label: "Calendário", icon: "📅" },
  ];

  // Telas de loading
  if (authLoading) return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: "#0a0f1a" }}>
      <p className="text-slate-500 text-sm">Carregando...</p>
    </div>
  );

  // Tela de login
  if (!user) return <Login onSignIn={signIn} onSignUp={signUp} />;

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#0a0f1a", fontFamily: "'DM Sans',system-ui,sans-serif" }}>

      {/* Header — PRESERVADO. Adicionado botão de logout discreto */}
      <header className="sticky top-0 z-40 border-b border-white/5" style={{ background: "rgba(10,15,26,0.95)", backdropFilter: "blur(20px)" }}>
        <div className="max-w-2xl mx-auto px-4 py-3 flex items-center justify-between">
          <div>
            <h1 className="text-base font-bold text-white tracking-tight">Controle do Semestre</h1>
            <p className="text-xs text-slate-500">{items.length} atividades · {disciplines.length} disciplinas</p>
          </div>
          <div className="flex gap-2">
            <button onClick={() => setModal("import")} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Importar">📂</button>
            <button onClick={() => exportToExcel(atividades)} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Exportar">📥</button>
            <button onClick={() => setModal("sim")} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Simulações">🔮</button>
            <button onClick={signOut} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Sair">🚪</button>
            <button onClick={() => { setEditingItem(null); setModal("add"); }} className="px-3 h-8 rounded-xl text-xs font-semibold text-white flex items-center gap-1" style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)" }}>+ Novo</button>
          </div>
        </div>
      </header>

      {/* Banner de erro de rede — não-intrusivo */}
      {error && (
        <div className="border-b border-red-500/20" style={{ background: "rgba(239,68,68,0.08)" }}>
          <div className="max-w-2xl mx-auto px-4 py-2">
            <p className="text-xs text-red-400">⚠️ {error}</p>
          </div>
        </div>
      )}

      {/* Banner de alerta — PRESERVADO DO ORIGINAL */}
      {(stats.some(s => s.emRisco) || items.some(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 3 && it.pontuacao == null)) && (
        <div className="border-b border-amber-500/20" style={{ background: "rgba(245,158,11,0.08)" }}>
          <div className="max-w-2xl mx-auto px-4 py-2">
            <p className="text-xs text-amber-400 flex items-center gap-2">⚠️
              {stats.filter(s => s.emRisco).length > 0 && <span>{stats.filter(s => s.emRisco).length} disciplina(s) em risco</span>}
              {items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 3 && it.pontuacao == null).length > 0 && <span>· {items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 3 && it.pontuacao == null).length} atividade(s) em até 3 dias</span>}
            </p>
          </div>
        </div>
      )}

      {/* Tabs — PRESERVADAS DO ORIGINAL */}
      <div className="sticky top-[57px] z-30 border-b border-white/5" style={{ background: "rgba(10,15,26,0.95)", backdropFilter: "blur(20px)" }}>
        <div className="max-w-2xl mx-auto px-4">
          <div className="flex overflow-x-auto">
            {TABS.map(t => <button key={t.id} onClick={() => setTab(t.id)} className={`flex-1 min-w-[64px] py-3 text-xs font-semibold flex flex-col items-center gap-0.5 transition border-b-2 ${tab === t.id ? "border-indigo-500 text-indigo-400" : "border-transparent text-slate-500 hover:text-slate-300"}`}><span>{t.icon}</span><span className="whitespace-nowrap">{t.label}</span></button>)}
          </div>
        </div>
      </div>

      {/* Loading de dados */}
      {dataLoading ? (
        <div className="flex-1 flex items-center justify-center">
          <p className="text-slate-500 text-sm">Carregando suas atividades...</p>
        </div>
      ) : (
        <main className="flex-1 max-w-2xl mx-auto w-full px-4 py-6 pb-24">
          {tab === "dashboard" && <DashboardTab items={items} stats={stats} />}
          {tab === "disciplines" && <DisciplineTab stats={stats} onEditItem={handleEdit} onDeleteItem={handleDelete} />}
          {tab === "all" && <AllItemsTab items={items} stats={stats} onEditItem={handleEdit} onDeleteItem={handleDelete} />}
          {tab === "alerts" && <AlertsTab items={items} stats={stats} />}
          {tab === "calendar" && <CalendarTab items={items} />}
        </main>
      )}

      {/* FAB — PRESERVADO DO ORIGINAL */}
      <button onClick={() => { setEditingItem(null); setModal("add"); }} className="fixed bottom-6 right-4 w-14 h-14 rounded-2xl text-2xl text-white shadow-2xl z-30 flex items-center justify-center sm:hidden" style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)", boxShadow: "0 8px 32px rgba(99,102,241,0.4)" }}>+</button>

      {/* Modais — PRESERVADOS DO ORIGINAL */}
      <Modal open={modal === "add" || modal === "edit"} onClose={() => { setModal(null); setEditingItem(null); }} title={modal === "edit" ? "Editar Atividade" : "Nova Atividade"}>
        <ItemForm item={editingItem} onSave={handleSave} onClose={() => { setModal(null); setEditingItem(null); }} disciplines={disciplines} />
      </Modal>
      <Modal open={modal === "import"} onClose={() => setModal(null)} title="Importar Planilha">
        <ImportPanel onImport={importAtividades} onClose={() => setModal(null)} />
      </Modal>
      <Modal open={modal === "sim"} onClose={() => setModal(null)} title="🔮 Simulação Acadêmica">
        <SimulationPanel stats={stats} onClose={() => setModal(null)} />
      </Modal>
      <Modal open={modal === "delete"} onClose={() => setModal(null)} title="Excluir Atividade">
        <div className="flex flex-col gap-4">
          <p className="text-sm text-slate-400">Tem certeza que deseja excluir esta atividade? Esta ação não pode ser desfeita.</p>
          <div className="flex gap-2">
            <button onClick={() => setModal(null)} className="flex-1 py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10">Cancelar</button>
            <button onClick={confirmDelete} className="flex-1 py-3 rounded-xl text-sm font-semibold text-white" style={{ background: "linear-gradient(135deg,#ef4444,#dc2626)" }}>Excluir</button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
