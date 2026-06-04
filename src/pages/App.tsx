// ============================================================
// src/pages/App.tsx
// Versão com Supabase + Tab Planejamento + Seletor de Disciplina
// ============================================================

import { useState, useCallback, useMemo, useEffect, useRef } from "react";
import { PieChart, Pie, Cell, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

import { useAuth } from "../hooks/useAuth";
import { useAtividades } from "../hooks/useAtividades";
import { useProfile } from "../hooks/useProfile";
import { useDisciplinas } from "../hooks/useDisciplinas";
import { enrichItem, getDisciplineStats, calcNota, calcDaysRemaining } from "../lib/calculos";
import { exportToExcel, parseXlsxFile } from "../lib/xlsx";
import { Login } from "./Login";
import type { Atividade, AtividadeEnriquecida, DisciplinaStats, TipoDisciplina, ParteDisciplina, Disciplina } from "../types";

// ─── CONSTANTES ──────────────────────────────────────────────
const STATUS_OPTIONS = ["Não iniciado", "Estudo inicial", "Estudo médio", "Estudo avançado", "Finalizado"];
const AVALIACAO_OPTIONS = ["AP1", "AP2", "AS", "AF"];
const STATUS_COLORS: Record<string, string> = {
  "Não iniciado": "#64748b", "Estudo inicial": "#f59e0b",
  "Estudo médio": "#f97316", "Estudo avançado": "#06b6d4", "Finalizado": "#10b981",
};
const DISC_COLORS = ["#6366f1","#8b5cf6","#ec4899","#f43f5e","#f97316","#eab308","#10b981","#06b6d4"];

function buildExportHref(items: Atividade[]) {
  try {
    const headers = ["#","Avaliação","Instrumento","Disciplina","Subdivisão","Status","Data","Restam (Dias)","Peso Avaliação","Peso Instrumento","Pontuação Máxima","Pontuação","Nota (de 10)","Observações"];
    const rows = items.map((it, i) => {
      const nota = calcNota(it.pontuacao, it.pontuacaoMaxima);
      const days = calcDaysRemaining(it.data);
      return [i+1,it.avaliacao,it.instrumento,it.disciplina,it.subdivisao||"",it.status,it.data||"",days!=null?(days<0?"Já passou":days):"",it.pesoAvaliacao,it.pesoInstrumento,it.pontuacaoMaxima??"",it.pontuacao??"",nota!=null?Math.round(nota*100)/100:"",it.observacoes||""];
    });
    const ws = XLSX.utils.aoa_to_sheet([headers,...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
    return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + b64;
  } catch { return null; }
}

// ─── UI PRIMITIVOS ───────────────────────────────────────────
const BADGE_STYLES: Record<string, string> = {
  green: "bg-emerald-500/20 text-emerald-400 border border-emerald-500/30",
  red: "bg-red-500/20 text-red-400 border border-red-500/30",
  yellow: "bg-amber-500/20 text-amber-400 border border-amber-500/30",
  cyan: "bg-cyan-500/20 text-cyan-400 border border-cyan-500/30",
  slate: "bg-slate-500/20 text-slate-400 border border-slate-500/30",
  indigo: "bg-indigo-500/20 text-indigo-400 border border-indigo-500/30",
  orange: "bg-orange-500/20 text-orange-400 border border-orange-500/30",
  purple: "bg-purple-500/20 text-purple-400 border border-purple-500/30",
};

function Badge({ children, color = "slate" }: { children: React.ReactNode; color?: string }) {
  return <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium ${BADGE_STYLES[color] || BADGE_STYLES.slate}`}>{children}</span>;
}

const STATUS_BADGE_MAP: Record<string, string> = {
  "Não iniciado": "slate", "Estudo inicial": "yellow", "Estudo médio": "orange", "Estudo avançado": "cyan", "Finalizado": "green",
};
function StatusBadge({ status }: { status: string }) {
  return <Badge color={STATUS_BADGE_MAP[status] || "slate"}>{status}</Badge>;
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
      <div className="relative w-full sm:max-w-lg rounded-t-3xl sm:rounded-2xl border border-white/10 overflow-hidden" style={{ background: "#111827", maxHeight: "92dvh", overflowY: "auto" }}>
        <div className="flex items-center justify-between px-6 py-4 border-b border-white/10 sticky top-0 z-10" style={{ background: "#111827" }}>
          <h3 className="text-base font-bold text-white">{title}</h3>
          <button onClick={onClose} className="w-8 h-8 rounded-full flex items-center justify-center text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.07)" }}>✕</button>
        </div>
        <div className="p-6 modal-safe">{children}</div>
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

function parseDecimal(str: string): number | null {
  if (str === "" || str == null) return null;
  const n = parseFloat(String(str).replace(",", "."));
  return isNaN(n) ? null : n;
}

function notaColor(nota: number | null, meta: number) {
  if (nota == null) return "#64748b";
  if (nota >= meta) return "#10b981";
  if (nota >= meta - 2) return "#f59e0b";
  return "#ef4444";
}

// ─── CLASSIFICAÇÃO DE ITEM (P10) ──────────────────────────────
// Categorias automáticas, válidas para Avaliações e Eventos.
type CatId = "ate3" | "ate7" | "ate14" | "ate30" | "aguardando" | "concluido" | "futuro_distante";
const CATEGORIAS: { id: CatId; label: string; color: string }[] = [
  { id: "ate3",       label: "Até 3 dias",          color: "#ef4444" },
  { id: "ate7",       label: "Até 7 dias",          color: "#f97316" },
  { id: "ate14",      label: "Até 14 dias",         color: "#eab308" },
  { id: "ate30",      label: "Até 30 dias",         color: "#6366f1" },
  { id: "aguardando", label: "Aguardando Correção", color: "#f59e0b" },
  { id: "concluido",  label: "Concluído",           color: "#10b981" },
];
function classificarItem(it: AtividadeEnriquecida): { id: CatId; color: string; label: string } {
  const isEvento = it.tipo === "evento";
  const d = it.daysRemaining;
  // Concluído: avaliação com nota lançada OU evento com data passada
  const concluido = isEvento ? (d != null && d < 0) : (it.pontuacao != null);
  if (concluido) return CATEGORIAS.find(c => c.id === "concluido")!;
  // Aguardando correção: avaliação cuja data passou e ainda sem nota
  if (!isEvento && d != null && d < 0 && it.pontuacao == null) return CATEGORIAS.find(c => c.id === "aguardando")!;
  if (d != null) {
    if (d <= 3)  return CATEGORIAS.find(c => c.id === "ate3")!;
    if (d <= 7)  return CATEGORIAS.find(c => c.id === "ate7")!;
    if (d <= 14) return CATEGORIAS.find(c => c.id === "ate14")!;
    if (d <= 30) return CATEGORIAS.find(c => c.id === "ate30")!;
  }
  return { id: "futuro_distante", color: "#64748b", label: "Futuro" };
}

// ─── ITEM FORM (com seletor de disciplina existente/nova) ─────
const INPUT_CLS = "w-full rounded-xl px-3 py-2.5 text-sm text-white border border-white/10 outline-none focus:border-indigo-500 transition";
const INPUT_STY = { background: "rgba(255,255,255,0.06)" };

function ItemForm({ item, onSave, onClose, disciplines, disciplinasFull, tipoInicial }: {
  item: Atividade | null; onSave: (form: Omit<Atividade, "id">) => void | Promise<void>;
  onClose: () => void; disciplines: string[]; disciplinasFull: Disciplina[]; tipoInicial?: "avaliacao" | "evento";
}) {
  const toStr = (v: any) => (v == null ? "" : String(v).replace(".", ","));
  const isEditing = !!item;
  const itemDiscExists = item && disciplines.includes(item.disciplina);
  const [discMode, setDiscMode] = useState(
    isEditing ? (itemDiscExists ? "existing" : "new") : (disciplines.length > 0 ? "existing" : "new")
  );
  const [newDiscName, setNewDiscName] = useState(isEditing && !itemDiscExists ? (item?.disciplina || "") : "");
  const [tipo, setTipo] = useState<"avaliacao" | "evento">(item?.tipo === "evento" ? "evento" : (tipoInicial || "avaliacao"));
  const [form, setForm] = useState<any>(() => {
    const base: any = item || { avaliacao: "AP1", instrumento: "", disciplina: disciplines[0] || "", disciplinaId: null, parte: "unica", subdivisao: "", status: "Não iniciado", data: "", pesoAvaliacao: 0.2, pesoInstrumento: 1.0, pontuacaoMaxima: null, pontuacao: null, observacoes: "" };
    return { ...base, _pa: toStr(base.pesoAvaliacao), _pi: toStr(base.pesoInstrumento), _pm: toStr(base.pontuacaoMaxima), _p: toStr(base.pontuacao) };
  });
  const [errors, setErrors] = useState<any>({});
  const [saving, setSaving] = useState(false);          // FIX BUG2: trava contra duplo-clique
  const set = (k: string, v: any) => setForm((f: any) => ({ ...f, [k]: v }));
  const isEvento = tipo === "evento";
  const disciplinaFinal = discMode === "existing" ? form.disciplina : newDiscName.trim();
  // Disciplina selecionada como entidade (para saber tipo, subdivisões, id)
  const discObj = disciplinasFull.find(d => d.nome === disciplinaFinal) || null;
  const isMista = discObj?.tipo === "Mista";
  const subdivisoesDisc = discObj?.subdivisoes ?? [];
  const handleSave = async () => {
    if (saving) return;                                  // FIX BUG2: ignora cliques repetidos
    const errs: any = {};
    if (!form.instrumento.trim()) errs.instrumento = true;
    // Disciplina é obrigatória para avaliação, mas OPCIONAL para evento
    if (!isEvento && !disciplinaFinal) errs.disciplina = true;
    if (Object.keys(errs).length) { setErrors(errs); return; }
    setErrors({});
    setSaving(true);
    try {
      const parteFinal = isMista ? (form.parte === "pratica" ? "pratica" : "teorica") : "unica";
      if (isEvento) {
        // Eventos: sem nota, sem peso — não afetam cálculos. Disciplina opcional.
        await onSave({ ...form, tipo: "evento", avaliacao: "Evento", disciplina: disciplinaFinal || "", disciplinaId: discObj?.id ?? null, parte: "unica", pesoAvaliacao: 0, pesoInstrumento: 0, pontuacaoMaxima: null, pontuacao: null });
      } else {
        await onSave({ ...form, tipo: "avaliacao", disciplina: disciplinaFinal, disciplinaId: discObj?.id ?? null, parte: parteFinal, pesoAvaliacao: parseDecimal(form._pa) ?? 0, pesoInstrumento: parseDecimal(form._pi) ?? 0, pontuacaoMaxima: parseDecimal(form._pm), pontuacao: parseDecimal(form._p) });
      }
    } finally {
      setSaving(false);
    }
  };
  const errCls = "border-red-500/60";
  return (
    <div className="flex flex-col gap-4">
      {/* Seletor de tipo: Avaliação ou Evento */}
      <div className="flex rounded-xl overflow-hidden border border-white/10" style={{ background: "rgba(255,255,255,0.03)" }}>
        <button type="button" onClick={() => setTipo("avaliacao")}
          className={`flex-1 py-2.5 text-xs font-semibold transition flex items-center justify-center gap-1.5 ${tipo === "avaliacao" ? "text-white" : "text-slate-500 hover:text-slate-300"}`}
          style={tipo === "avaliacao" ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>
          📊 Avaliação
        </button>
        <button type="button" onClick={() => setTipo("evento")}
          className={`flex-1 py-2.5 text-xs font-semibold transition flex items-center justify-center gap-1.5 ${tipo === "evento" ? "text-white" : "text-slate-500 hover:text-slate-300"}`}
          style={tipo === "evento" ? { background: "linear-gradient(135deg,#06b6d4,#0891b2)" } : {}}>
          📅 Evento
        </button>
      </div>
      {isEvento && <p className="text-xs text-slate-500 -mt-1">Eventos (palestras, entregas, visitas) não têm nota nem peso e não afetam suas médias.</p>}
      <div className="grid grid-cols-2 gap-3">
        {isEvento ? (
          <FormField label="Tipo"><div className={INPUT_CLS} style={{ ...INPUT_STY, opacity: 0.7 }}>📅 Evento</div></FormField>
        ) : (
          <FormField label="Avaliação"><select value={form.avaliacao} onChange={e => set("avaliacao", e.target.value)} className={INPUT_CLS} style={INPUT_STY}>{AVALIACAO_OPTIONS.map(o => <option key={o}>{o}</option>)}</select></FormField>
        )}
        <FormField label="Status"><select value={form.status} onChange={e => set("status", e.target.value)} className={INPUT_CLS} style={INPUT_STY}>{STATUS_OPTIONS.map(o => <option key={o}>{o}</option>)}</select></FormField>
      </div>
      <FormField label={`${isEvento ? "Evento" : "Instrumento"}${errors.instrumento ? " — obrigatório" : ""}`}>
        <input value={form.instrumento} onChange={e => set("instrumento", e.target.value)} className={`${INPUT_CLS} ${errors.instrumento ? errCls : ""}`} style={INPUT_STY} placeholder={isEvento ? "Nome do evento" : "Nome do instrumento"} />
      </FormField>
      <FormField label={`Disciplina${isEvento ? " (opcional)" : ""}${errors.disciplina ? " — obrigatório" : ""}`}>
        <div className="flex rounded-xl overflow-hidden border border-white/10 mb-2" style={{ background: "rgba(255,255,255,0.03)" }}>
          <button type="button" onClick={() => { setDiscMode("existing"); if (!form.disciplina && disciplines.length > 0) set("disciplina", disciplines[0]); }}
            className={`flex-1 py-2 text-xs font-semibold transition ${discMode === "existing" ? "text-white" : "text-slate-500 hover:text-slate-300"}`}
            style={discMode === "existing" ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>
            Disciplina existente
          </button>
          <button type="button" onClick={() => setDiscMode("new")}
            className={`flex-1 py-2 text-xs font-semibold transition ${discMode === "new" ? "text-white" : "text-slate-500 hover:text-slate-300"}`}
            style={discMode === "new" ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>
            + Nova disciplina
          </button>
        </div>
        {isEvento && <p className="text-xs text-slate-600 mb-2 px-1">Eventos podem ficar sem disciplina (aparecem em "Geral").</p>}
        {discMode === "existing" ? (
          disciplines.length === 0 ? (
            <p className="text-xs text-slate-500 italic px-1">Nenhuma disciplina ainda. <button type="button" className="text-indigo-400 underline" onClick={() => setDiscMode("new")}>Criar nova</button></p>
          ) : (
            <select value={form.disciplina} onChange={e => set("disciplina", e.target.value)} className={`${INPUT_CLS} ${errors.disciplina && !form.disciplina ? errCls : ""}`} style={INPUT_STY}>
              {isEvento && <option value="">Sem disciplina (Geral)</option>}
              {disciplines.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
          )
        ) : (
          <input value={newDiscName} onChange={e => setNewDiscName(e.target.value)} className={`${INPUT_CLS} ${errors.disciplina && !newDiscName.trim() ? errCls : ""}`} style={INPUT_STY} placeholder="Nome da nova disciplina" autoFocus />
        )}
        {discMode === "existing" && form.disciplina && <p className="text-xs text-slate-600 mt-1 px-1">Disciplina: <span className="text-slate-400">{form.disciplina}</span>{discObj ? ` · ${discObj.tipo}` : ""}</p>}
        {discMode === "new" && newDiscName.trim() && <p className="text-xs text-emerald-600 mt-1 px-1">✓ Será criada: <span className="text-emerald-400">{newDiscName.trim()}</span></p>}
      </FormField>
      {/* Parte da disciplina — só para avaliação em disciplina MISTA */}
      {!isEvento && isMista && (
        <FormField label="Parte da disciplina">
          <div className="flex rounded-xl overflow-hidden border border-white/10" style={{ background: "rgba(255,255,255,0.03)" }}>
            <button type="button" onClick={() => set("parte", "teorica")} className={`flex-1 py-2 text-xs font-semibold transition ${form.parte !== "pratica" ? "text-white" : "text-slate-500"}`} style={form.parte !== "pratica" ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>Teórica ({discObj?.pesoTeorica.toFixed(0)}%)</button>
            <button type="button" onClick={() => set("parte", "pratica")} className={`flex-1 py-2 text-xs font-semibold transition ${form.parte === "pratica" ? "text-white" : "text-slate-500"}`} style={form.parte === "pratica" ? { background: "linear-gradient(135deg,#06b6d4,#0891b2)" } : {}}>Prática ({discObj?.pesoPratica.toFixed(0)}%)</button>
          </div>
        </FormField>
      )}
      {/* Subdivisão: dropdown se a disciplina tem subdivisões; senão campo livre */}
      {subdivisoesDisc.length > 0 ? (
        <FormField label="Subdivisão"><select value={form.subdivisao} onChange={e => set("subdivisao", e.target.value)} className={INPUT_CLS} style={INPUT_STY}><option value="">Sem subdivisão</option>{subdivisoesDisc.map(s => <option key={s} value={s}>{s}</option>)}</select></FormField>
      ) : (
        <FormField label="Subdivisão (opcional)"><input value={form.subdivisao} onChange={e => set("subdivisao", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: Pneumo e Cardio" /></FormField>
      )}
      <FormField label="Data"><input type="date" value={form.data} onChange={e => set("data", e.target.value)} className={INPUT_CLS} style={INPUT_STY} /></FormField>
      {!isEvento && (
        <>
          <div className="grid grid-cols-2 gap-3">
            <FormField label="Peso Avaliação (0–1)"><input type="text" inputMode="decimal" value={form._pa} onChange={e => set("_pa", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: 0,3" /></FormField>
            <FormField label="Peso Instrumento (0–1)"><input type="text" inputMode="decimal" value={form._pi} onChange={e => set("_pi", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: 0,5" /></FormField>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <FormField label="Pontuação Máxima"><input type="text" inputMode="decimal" value={form._pm} onChange={e => set("_pm", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: 10" /></FormField>
            <FormField label="Pontuação Obtida"><input type="text" inputMode="decimal" value={form._p} onChange={e => set("_p", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: 7,5" /></FormField>
          </div>
        </>
      )}
      <FormField label="Observações"><textarea value={form.observacoes} onChange={e => set("observacoes", e.target.value)} rows={2} className={INPUT_CLS} style={INPUT_STY} placeholder="Anotações opcionais..." /></FormField>
      <div className="flex gap-2 mt-2">
        <button onClick={onClose} disabled={saving} className="flex-1 py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10 disabled:opacity-40">Cancelar</button>
        <button onClick={handleSave} disabled={saving} className="flex-1 py-3 rounded-xl text-sm font-semibold text-white disabled:opacity-60 flex items-center justify-center gap-2" style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)" }}>
          {saving ? "Salvando..." : "Salvar"}
        </button>
      </div>
    </div>
  );
}

// ─── IMPORT PANEL ─────────────────────────────────────────────
function ImportPanel({ onImport, onClose }: { onImport: (items: Omit<Atividade, "id">[]) => void; onClose: () => void }) {
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const handleFile = useCallback(async (file: File) => {
    if (!file) return;
    setStatus("Lendo arquivo...");
    try {
      const parsed = await parseXlsxFile(file);
      setStatus(`✅ ${parsed.length} atividades importadas!`);
      setTimeout(() => { onImport(parsed); onClose(); }, 1200);
    } catch { setStatus("❌ Erro ao ler o arquivo. Verifique se é um .xlsx válido."); }
  }, [onImport, onClose]);
  return (
    <div className="flex flex-col gap-4">
      <p className="text-sm text-slate-400">Importe uma planilha Excel (.xlsx). Os dados existentes serão substituídos.</p>
      <div className={`rounded-2xl border-2 border-dashed p-8 flex flex-col items-center gap-3 transition cursor-pointer ${dragging ? "border-indigo-500" : "border-white/15"}`} style={{ background: dragging ? "rgba(99,102,241,0.08)" : "rgba(255,255,255,0.03)" }} onDragOver={e => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={e => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }} onClick={() => inputRef.current?.click()}>
        <span className="text-3xl">📂</span>
        <p className="text-sm text-white font-medium">Arraste o arquivo ou clique para selecionar</p>
        <p className="text-xs text-slate-500">.xlsx · Aba Sheet1</p>
      </div>
      <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f); e.target.value = ""; }} />
      <button onClick={() => inputRef.current?.click()} className="w-full py-3 rounded-xl text-sm font-semibold text-white border border-indigo-500/40 transition hover:border-indigo-500" style={{ background: "rgba(99,102,241,0.15)" }}>Selecionar arquivo .xlsx</button>
      {status && <div className="rounded-xl p-3 border border-white/10 text-sm text-center" style={{ background: "rgba(255,255,255,0.05)" }}>{status}</div>}
      <button onClick={onClose} className="w-full py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10">Cancelar</button>
    </div>
  );
}

// ─── EXPORT PANEL ─────────────────────────────────────────────
function ExportPanel({ items, onClose }: { items: Atividade[]; onClose: () => void }) {
  const [href, setHref] = useState<string | null>(null);
  const [erro, setErro] = useState<string | null>(null);
  useEffect(() => {
    const h = buildExportHref(items);
    if (h) setHref(h);
    else setErro("Não foi possível gerar o arquivo.");
  }, [items]);
  return (
    <div className="flex flex-col gap-4">
      <p className="text-sm text-slate-400">Toque no botão abaixo para baixar o Excel com todas as atividades atuais.</p>
      {erro && <div className="rounded-xl p-3 border border-red-500/30 text-sm text-red-400" style={{ background: "rgba(239,68,68,0.08)" }}>{erro}</div>}
      {href && <a href={href} download="controle-semestre.xlsx" className="w-full py-4 rounded-xl text-sm font-semibold text-white text-center block" style={{ background: "linear-gradient(135deg,#10b981,#059669)" }}>⬇️ Baixar controle-semestre.xlsx<span className="block text-xs font-normal mt-1 opacity-75">{items.length} atividades</span></a>}
      {!href && !erro && <div className="rounded-xl p-4 text-center text-sm text-slate-500">Gerando arquivo...</div>}
      <button onClick={onClose} className="w-full py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10">Fechar</button>
    </div>
  );
}

// ─── DASHBOARD TAB ────────────────────────────────────────────
function DashboardTab({ items, stats, meta }: { items: AtividadeEnriquecida[]; stats: DisciplinaStats[]; meta: number }) {
  const [expandedChart, setExpandedChart] = useState<string | null>(null);
  const upcomingAll = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 30).sort((a, b) => (a.daysRemaining ?? 0) - (b.daysRemaining ?? 0)), [items]);
  const { totalFinished, totalInProgress, totalPending } = useMemo(() => ({ totalFinished: items.filter(it => it.status === "Finalizado").length, totalInProgress: items.filter(it => ["Estudo inicial","Estudo médio","Estudo avançado"].includes(it.status)).length, totalPending: items.filter(it => it.status === "Não iniciado").length }), [items]);
  const avgGrade = useMemo(() => { const graded = stats.filter(s => s.mediaAtual != null); return graded.length ? (graded.reduce((a, s) => a + (s.mediaAtual ?? 0), 0) / graded.length).toFixed(2) : null; }, [stats]);
  const statusPieData = useMemo(() => [{ name: "Finalizado", value: totalFinished, color: "#10b981" }, { name: "Em andamento", value: totalInProgress, color: "#06b6d4" }, { name: "Pendente", value: totalPending, color: "#64748b" }].filter(d => d.value > 0), [totalFinished, totalInProgress, totalPending]);
  const statusPieTotal = useMemo(() => statusPieData.reduce((a, d) => a + d.value, 0), [statusPieData]);
  const pontosChartData = useMemo(() => stats.map(s => {
    let obtidos = 0, futuros = 0;
    s.items.forEach(r => { const contrib = r.pesoAvaliacao * r.pesoInstrumento * 10; if (r.pontuacao != null && r.pontuacaoMaxima) { obtidos += ((r.pontuacao / r.pontuacaoMaxima)) * contrib; } else { futuros += contrib; } });
    return { fullName: s.disciplina, obtidos: Math.round(obtidos * 100) / 100, futuros: Math.round(futuros * 100) / 100, color: s.color };
  }), [stats]);
  const pendingByStatus = useMemo(() => {
    const groups = [
      { status: "Não iniciado", color: "#64748b" },
      { status: "Estudo inicial", color: "#f59e0b" },
      { status: "Estudo médio", color: "#f97316" },
      { status: "Estudo avançado", color: "#06b6d4" },
    ];
    return groups.map(g => ({ ...g, items: items.filter(it => it.status === g.status && it.pontuacao == null) })).filter(g => g.items.length > 0);
  }, [items]);
  return (
    <div className="flex flex-col gap-6">
      <div className="grid grid-cols-2 gap-3">
        <MetricCard label="Média Geral" value={avgGrade ?? "—"} sub="Ponderada" color="#6366f1" icon="📊" />
        <MetricCard label="Próximas (30d)" value={upcomingAll.length} sub="atividades" color="#f59e0b" icon="⏰" />
      </div>
      {stats.some(s => s.emRisco) && (
        <div className="rounded-2xl p-4 border border-red-500/30" style={{ background: "rgba(239,68,68,0.08)" }}>
          <div className="flex items-center gap-2 mb-3"><span>⚠️</span><p className="text-sm font-bold text-red-400">Disciplinas em Risco</p></div>
          {stats.filter(s => s.emRisco).map(s => <div key={s.disciplina} className="flex items-center justify-between"><span className="text-sm text-white">{s.disciplina}</span><span className="text-sm font-bold text-red-400">{s.mediaAtual?.toFixed(2)}</span></div>)}
        </div>
      )}
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <ChartHeader label="Pontos na Média Final (de 10)" chartKey="grades" expandedChart={expandedChart} setExpandedChart={setExpandedChart} />
        <p className="text-xs text-slate-600 mb-4 -mt-2">Pontos já conquistados + pontos ainda disponíveis</p>
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
                  {obtPct > 0 && <div className="flex items-center justify-end pr-1.5" style={{ width: `${obtPct}%`, background: s.color, minWidth: 2, transition: "width 0.6s ease" }}>{obtPct > 12 && <span style={{ fontSize: 10, color: "#fff", fontWeight: 700 }}>{d.obtidos.toFixed(1)}</span>}</div>}
                  {obtPct > 0 && futPct > 0 && <div style={{ width: 2, background: "#0a0f1a", flexShrink: 0 }} />}
                  {futPct > 0 && <div className="flex items-center pl-1.5" style={{ width: `${futPct}%`, background: "rgba(255,255,255,0.12)", minWidth: 2 }}>{futPct > 12 && <span style={{ fontSize: 10, color: "#94a3b8" }}>{d.futuros.toFixed(1)}</span>}</div>}
                </div>
                {i === pontosChartData.length - 1 && <div className="flex justify-between mt-1.5">{[0,1,2,3,4,5,6,7,8,9,10].map(n => <span key={n} style={{ color: n === meta ? "#f59e0b" : "#475569", fontSize: 9, fontWeight: n === meta ? 700 : 400 }}>{n}</span>)}</div>}
              </div>
            );
          })}
        </div>
        <div className="flex items-center gap-4 mt-4 pt-3 border-t border-white/5">
          <div className="flex items-center gap-1.5"><div className="w-10 h-3 rounded-sm" style={{ background: "linear-gradient(90deg,#6366f1,#8b5cf6)" }} /><span className="text-xs text-slate-400">Conquistados</span></div>
          <div className="flex items-center gap-1.5"><div className="w-10 h-3 rounded-sm" style={{ background: "rgba(255,255,255,0.14)", border: "1px solid rgba(255,255,255,0.1)" }} /><span className="text-xs text-slate-400">Disponíveis</span></div>
          <div className="flex items-center gap-1.5 ml-auto"><div className="w-1 h-3" style={{ background: "#f59e0b" }} /><span className="text-xs text-amber-400">{meta} = meta</span></div>
        </div>
        {expandedChart === "grades" && (
          <div className="mt-4 flex flex-col gap-3 border-t border-white/8 pt-4">
            {pontosChartData.map((d, i) => { const s = stats[i]; return (
              <div key={d.fullName}>
                <div className="flex items-center justify-between mb-1">
                  <span className="text-xs text-slate-300 font-semibold truncate pr-2">{d.fullName}</span>
                  <div className="flex items-center gap-3 shrink-0"><span className="text-xs text-slate-500">disponível: +{d.futuros.toFixed(2)}</span><span className="text-sm font-bold" style={{ color: d.obtidos >= meta ? "#10b981" : d.obtidos >= meta - 2 ? "#f59e0b" : "#ef4444" }}>{d.obtidos.toFixed(2)}/10</span></div>
                </div>
                {s.notaNecessaria != null && s.notaNecessaria > 0 && !s.aprovacaoGarantida && (<p className="text-xs text-slate-500">Precisa de <span className="font-semibold" style={{ color: s.aprovacaoImpossivel ? "#ef4444" : "#f59e0b" }}>{s.aprovacaoImpossivel ? "nota impossível" : s.notaNecessaria.toFixed(2)}</span> nas próximas para fechar {meta}</p>)}
                {s.aprovacaoGarantida && <p className="text-xs text-emerald-500">✓ Meta {meta} já garantida</p>}
              </div>
            ); })}
          </div>
        )}
      </div>
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <ChartHeader label="Status dos Estudos" chartKey="status" expandedChart={expandedChart} setExpandedChart={setExpandedChart} />
        <div className="flex items-center gap-4">
          <ResponsiveContainer width={120} height={120}><PieChart><Pie data={statusPieData} dataKey="value" innerRadius={35} outerRadius={55} paddingAngle={3}>{statusPieData.map((e, i) => <Cell key={i} fill={e.color} />)}</Pie></PieChart></ResponsiveContainer>
          <div className="flex flex-col gap-2.5 flex-1">{statusPieData.map(d => (<div key={d.name} className="flex items-center justify-between"><div className="flex items-center gap-2"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: d.color }} /><span className="text-xs text-slate-400">{d.name}</span></div><div className="flex items-center gap-2"><span className="text-xs text-slate-500">{statusPieTotal > 0 ? Math.round(d.value / statusPieTotal * 100) : 0}%</span><span className="text-sm font-bold text-white w-5 text-right">{d.value}</span></div></div>))}</div>
        </div>
        {expandedChart === "status" && (
          <div className="mt-4 flex flex-col gap-4 border-t border-white/8 pt-4">
            <p className="text-xs text-slate-500">Instrumentos pendentes por status de estudo:</p>
            {pendingByStatus.length === 0 ? <p className="text-xs text-slate-600 italic">Nenhum instrumento pendente.</p> : pendingByStatus.map(group => (
              <div key={group.status}>
                <div className="flex items-center gap-2 mb-2"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: group.color }} /><span className="text-xs font-bold text-white">{group.status}</span><span className="text-xs text-slate-500">({group.items.length})</span></div>
                <div className="flex flex-col gap-1 pl-4">{group.items.map(it => (<div key={it.id} className="flex items-center justify-between py-1 border-b border-white/5"><div className="min-w-0 flex-1"><p className="text-xs text-slate-300 truncate">{it.instrumento}</p><p className="text-xs text-slate-600">{it.disciplina} · {it.avaliacao}</p></div><div className="shrink-0 ml-2">{it.daysRemaining != null && <span className="text-xs" style={{ color: it.daysRemaining < 0 ? "#f59e0b" : it.daysRemaining <= 7 ? "#f59e0b" : "#64748b" }}>{it.daysRemaining < 0 ? `${Math.abs(it.daysRemaining)}d atrás` : `em ${it.daysRemaining}d`}</span>}</div></div>))}</div>
              </div>
            ))}
          </div>
        )}
      </div>
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <ChartHeader label="Ranking de Disciplinas" chartKey="ranking" expandedChart={expandedChart} setExpandedChart={setExpandedChart} />
        <div className="flex flex-col gap-3">{[...stats].sort((a, b) => (b.mediaAtual ?? -1) - (a.mediaAtual ?? -1)).map((s, i) => (<div key={s.disciplina}><div className="flex items-center justify-between mb-1"><div className="flex items-center gap-2"><span className="text-xs text-slate-600 w-4">#{i+1}</span><span className="text-sm text-white">{s.disciplina}</span></div><div className="flex items-center gap-2"><span className="text-sm font-bold" style={{ color: s.color }}>{s.mediaAtual?.toFixed(2) ?? "—"}</span>{s.emRisco && <Badge color="red">Risco</Badge>}</div></div><ProgressBar value={s.mediaAtual ?? 0} max={10} color={s.color} height={3} />{expandedChart === "ranking" && (<div className="mt-2 flex flex-col gap-1 pl-5 border-l-2" style={{ borderColor: s.color + "44" }}>{s.items.map(it => { const n = calcNota(it.pontuacao, it.pontuacaoMaxima); return (<div key={it.id} className="flex items-center justify-between py-0.5"><div className="flex items-center gap-2 min-w-0"><div className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: STATUS_COLORS[it.status] }} /><span className="text-xs text-slate-400 truncate">{it.instrumento}</span><span className="text-xs text-slate-600 shrink-0">{it.avaliacao}</span></div><span className="text-xs font-semibold shrink-0 ml-2" style={{ color: n == null ? "#64748b" : n >= meta ? "#10b981" : n >= meta - 2 ? "#f59e0b" : "#ef4444" }}>{n != null ? n.toFixed(1) : "—"}</span></div>); })}</div>)}</div>))}</div>
      </div>
    </div>
  );
}

// ─── DISCIPLINE TAB ───────────────────────────────────────────
function DisciplineTab({ items, stats, meta, onEditItem, onDeleteItem, onQuickGrade }: { items: AtividadeEnriquecida[]; stats: DisciplinaStats[]; meta: number; onEditItem: (item: Atividade) => void; onDeleteItem: (id: string) => void; onQuickGrade?: (item: Atividade) => void }) {
  const [selected, setSelected] = useState<string | null>(null);
  const disc = selected ? stats.find(s => s.disciplina === selected) : null;
  if (disc) return (
    <div className="flex flex-col gap-4">
      <button onClick={() => setSelected(null)} className="flex items-center gap-2 text-sm text-slate-400 hover:text-white transition"><span>←</span> Todas as disciplinas</button>
      <div className="rounded-2xl p-5 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <div className="flex items-center gap-3 mb-4"><div className="w-10 h-10 rounded-xl flex items-center justify-center text-lg" style={{ background: `${disc.color}22` }}>📖</div><div><h2 className="text-lg font-bold text-white">{disc.disciplina}</h2><p className="text-xs text-slate-500">{disc.items.length} instrumentos</p></div></div>
        <div className="grid grid-cols-2 gap-3 mb-4">
          <div className="rounded-xl p-3 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><p className="text-xs text-slate-500 mb-1">Média Atual</p><p className="text-2xl font-bold" style={{ color: disc.color }}>{disc.mediaAtual?.toFixed(2) ?? "—"}</p></div>
          <div className="rounded-xl p-3 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><p className="text-xs text-slate-500 mb-1">Nota Necessária</p><p className="text-2xl font-bold" style={{ color: disc.aprovacaoImpossivel ? "#ef4444" : disc.notaNecessaria == null ? "#64748b" : (disc.notaNecessaria > meta ? "#f59e0b" : "#10b981") }}>{disc.aprovacaoImpossivel ? "Impossível" : disc.notaNecessaria != null ? disc.notaNecessaria.toFixed(2) : "—"}</p></div>
        </div>
        <div className="mb-2"><div className="flex items-center justify-between mb-1"><span className="text-xs text-slate-500">Progresso concluído</span><span className="text-xs text-slate-400">{disc.pesoConcluido.toFixed(1)}%</span></div><ProgressBar value={disc.pesoConcluido} max={100} color={disc.color} height={6} /></div>
        {disc.notaNecessaria != null && <div className="mt-3 p-3 rounded-xl border" style={{ background: "rgba(99,102,241,0.08)", borderColor: "rgba(99,102,241,0.2)" }}><p className="text-xs text-indigo-300">{disc.aprovacaoGarantida ? `✅ Média ${meta} já garantida!` : disc.aprovacaoImpossivel ? `❌ Média ${meta} não é mais matematicamente possível.` : `📌 Você precisa de média ${disc.notaNecessaria.toFixed(2)} nas próximas avaliações para fechar com ${meta}.`}</p></div>}
      </div>
      <ItemsList items={items.filter(it => it.disciplina === disc.disciplina)} stats={stats} meta={meta} onEditItem={onEditItem} onDeleteItem={onDeleteItem} onQuickGrade={onQuickGrade} title={`Itens de ${disc.disciplina}`} />
    </div>
  );
  return (
    <div className="flex flex-col gap-4">
      <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">Selecione uma disciplina</p>
      {stats.length === 0 && <div className="rounded-2xl p-8 border border-white/5 flex flex-col items-center gap-2 text-center" style={{ background: "rgba(255,255,255,0.03)" }}><span className="text-3xl">📚</span><p className="text-sm font-semibold text-white">Nenhuma disciplina ainda</p><p className="text-xs text-slate-500">Toque em "+ Novo" para criar sua primeira disciplina ou atividade.</p></div>}
      {stats.map(s => (<button key={s.disciplina} onClick={() => setSelected(s.disciplina)} className="rounded-2xl p-5 border border-white/5 text-left w-full hover:border-white/15 transition-all" style={{ background: "rgba(255,255,255,0.04)" }}><div className="flex items-center justify-between mb-3"><div className="flex items-center gap-3"><div className="w-3 h-3 rounded-full shrink-0" style={{ background: s.color }} /><p className="text-sm font-bold text-white">{s.disciplina}</p></div><div className="flex items-center gap-2">{s.emRisco && <Badge color="red">Risco</Badge>}<span className="text-lg font-bold" style={{ color: s.color }}>{s.mediaAtual?.toFixed(2) ?? "—"}</span></div></div><ProgressBar value={s.pesoConcluido} max={100} color={s.color} height={5} /><div className="flex items-center justify-between mt-2"><span className="text-xs text-slate-500">{s.pesoConcluido.toFixed(1)}% concluído</span><div className="flex gap-2"><span className="text-xs text-emerald-400">{s.statusCounts["Finalizado"]} ok</span><span className="text-xs text-cyan-400">{s.statusCounts["Em andamento"]} em estudo</span><span className="text-xs text-slate-500">{s.statusCounts["Não iniciado"]} pendentes</span></div></div></button>))}
      {items.length > 0 && <div className="mt-2 pt-2 border-t border-white/5"><ItemsList items={items} stats={stats} meta={meta} onEditItem={onEditItem} onDeleteItem={onDeleteItem} onQuickGrade={onQuickGrade} title="Todas as atividades" /></div>}
    </div>
  );
}

// ─── ITEMS LIST (reutilizável: dentro de Disciplinas) ─────────
// Filtros: Todos / Não iniciado / Em Estudo / Finalizado
// Ordenação: Data (asc/desc), Nota Necessária (asc/desc)
const STATUS_FILTERS = [
  { id: "Todos", label: "Todos", match: (_it: AtividadeEnriquecida) => true },
  { id: "Não iniciado", label: "Não iniciado", match: (it: AtividadeEnriquecida) => it.status === "Não iniciado" },
  { id: "Em Estudo", label: "Em Estudo", match: (it: AtividadeEnriquecida) => ["Estudo inicial","Estudo médio","Estudo avançado"].includes(it.status) },
  { id: "Finalizado", label: "Finalizado", match: (it: AtividadeEnriquecida) => it.status === "Finalizado" },
];

function ItemsList({ items, stats, meta, onEditItem, onDeleteItem, onQuickGrade, title }: {
  items: AtividadeEnriquecida[]; stats: DisciplinaStats[]; meta: number;
  onEditItem: (item: Atividade) => void; onDeleteItem: (id: string) => void;
  onQuickGrade?: (item: Atividade) => void; title?: string;
}) {
  const [filter, setFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("Todos");
  const [sortField, setSortField] = useState<"data" | "nota">("data");
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");
  const notaNecMap = useMemo(() => { const m: Record<string, number> = {}; stats.forEach(s => { m[s.disciplina] = s.notaNecessaria ?? 0; }); return m; }, [stats]);
  const toggleSort = (field: "data" | "nota") => {
    if (sortField === field) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortField(field); setSortDir("asc"); }
  };
  const filtered = useMemo(() => {
    const q = filter.toLowerCase();
    const sf = STATUS_FILTERS.find(f => f.id === statusFilter) || STATUS_FILTERS[0];
    const list = items.filter(it =>
      (!q || it.disciplina.toLowerCase().includes(q) || it.instrumento.toLowerCase().includes(q) || it.avaliacao.toLowerCase().includes(q)) &&
      sf.match(it)
    );
    const dir = sortDir === "asc" ? 1 : -1;
    return [...list].sort((a, b) => {
      if (sortField === "data") return ((a.daysRemaining ?? 99999) - (b.daysRemaining ?? 99999)) * dir;
      return ((notaNecMap[a.disciplina] ?? 0) - (notaNecMap[b.disciplina] ?? 0)) * dir;
    });
  }, [items, filter, statusFilter, sortField, sortDir, notaNecMap]);
  const arrow = sortDir === "asc" ? "↑" : "↓";
  return (
    <div className="flex flex-col gap-4">
      {title && <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">{title}</p>}
      <input value={filter} onChange={e => setFilter(e.target.value)} placeholder="Buscar disciplina, instrumento, avaliação..." className="w-full rounded-xl px-4 py-3 text-sm text-white border border-white/10 outline-none focus:border-indigo-500 transition" style={{ background: "rgba(255,255,255,0.05)" }} />
      <div className="flex gap-2 overflow-x-auto pb-1">{STATUS_FILTERS.map(f => <button key={f.id} onClick={() => setStatusFilter(f.id)} className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${statusFilter === f.id ? "text-white" : "text-slate-400 border border-white/10"}`} style={statusFilter === f.id ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>{f.label}</button>)}</div>
      <div className="flex items-center gap-2"><span className="text-xs text-slate-500 shrink-0">Ordenar:</span><div className="flex gap-1.5">
        <button onClick={() => toggleSort("data")} className={`px-3 py-1 rounded-lg text-xs font-medium transition border ${sortField === "data" ? "text-white" : "text-slate-400 border-white/10"}`} style={sortField === "data" ? { background: "rgba(99,102,241,0.4)", borderColor: "rgba(99,102,241,0.5)" } : {}}>Data {sortField === "data" ? arrow : ""}</button>
        <button onClick={() => toggleSort("nota")} className={`px-3 py-1 rounded-lg text-xs font-medium transition border ${sortField === "nota" ? "text-white" : "text-slate-400 border-white/10"}`} style={sortField === "nota" ? { background: "rgba(99,102,241,0.4)", borderColor: "rgba(99,102,241,0.5)" } : {}}>Nota Necessária {sortField === "nota" ? arrow : ""}</button>
      </div></div>
      <p className="text-xs text-slate-500">{filtered.length} {filtered.length === 1 ? "item" : "itens"}</p>
      <div className="flex flex-col gap-2">{filtered.length === 0 ? <div className="rounded-2xl p-8 border border-white/5 text-center" style={{ background: "rgba(255,255,255,0.03)" }}><p className="text-sm text-slate-500">Nenhum item encontrado.</p></div> : filtered.map(it => { const nota = calcNota(it.pontuacao, it.pontuacaoMaxima); const isEvento = it.tipo === "evento"; return (<div key={it.id} className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><div className="flex items-start justify-between gap-2"><div className="flex-1 min-w-0"><p className="text-sm font-semibold text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500 mt-0.5">{it.disciplina} · {isEvento ? "📅 Evento" : it.avaliacao}</p><div className="flex gap-1 mt-1.5 flex-wrap"><StatusBadge status={it.status} />{it.daysRemaining != null && <Badge color={it.daysRemaining < 0 && it.pontuacao == null && !isEvento ? "yellow" : it.daysRemaining < 0 ? "slate" : it.daysRemaining <= 7 ? "yellow" : "slate"}>{it.daysRemaining < 0 ? (it.pontuacao == null && !isEvento ? `Aguard. correção (${Math.abs(it.daysRemaining)}d)` : `${Math.abs(it.daysRemaining)}d atrás`) : `${it.daysRemaining}d`}</Badge>}</div></div><div className="flex flex-col items-end gap-2 shrink-0">{isEvento ? <span className="text-xs text-cyan-400 font-semibold">evento</span> : <span className="text-base font-bold" style={{ color: notaColor(nota, meta) }}>{nota != null ? nota.toFixed(1) : "—"}/10</span>}<div className="flex gap-1">{!isEvento && onQuickGrade && <button onClick={() => onQuickGrade(it)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs transition" style={{ background: "rgba(16,185,129,0.15)" }} title="Lançar nota">🎯</button>}<button onClick={() => onEditItem(it)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Editar">✏️</button><button onClick={() => onDeleteItem(it.id)} className="w-7 h-7 rounded-lg flex items-center justify-center text-xs transition" style={{ background: "rgba(239,68,68,0.1)" }} title="Excluir">🗑</button></div></div></div></div>); })}</div>
    </div>
  );
}

// ─── ALERTS TAB ───────────────────────────────────────────────
function AlertSection({ title, color, icon, count, children }: any) {
  return <div className="rounded-2xl overflow-hidden border" style={{ borderColor: color + "33", background: color + "0d" }}><div className="flex items-center gap-2 px-4 py-3 border-b" style={{ borderColor: color + "22" }}><span className="text-base">{icon}</span><span className="text-sm font-bold" style={{ color }}>{title}</span><span className="ml-auto text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: color + "22", color }}>{count}</span></div><div className="flex flex-col">{children}</div></div>;
}
function AlertRow({ it, color }: any) {
  return <div className="flex items-center justify-between px-4 py-2.5 border-b border-white/5 last:border-0"><div className="min-w-0"><p className="text-sm text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {it.avaliacao}</p></div><span className="text-xs font-bold shrink-0 ml-3" style={{ color }}>{it.daysRemaining < 0 ? `${Math.abs(it.daysRemaining)}d atrás` : it.daysRemaining === 0 ? "Hoje" : `${it.daysRemaining}d`}</span></div>;
}
function AlertsTab({ items, stats, meta }: { items: AtividadeEnriquecida[]; stats: DisciplinaStats[]; meta: number }) {
  const urgentes = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 3 && it.pontuacao == null), [items]);
  const proximas = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining > 3 && it.daysRemaining <= 14 && it.pontuacao == null), [items]);
  const aguardando = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining < 0 && it.pontuacao == null), [items]);
  const emRisco = useMemo(() => stats.filter(s => s.emRisco), [stats]);
  const notaAlta = useMemo(() => stats.filter(s => s.notaNecessaria != null && s.notaNecessaria > 8 && !s.aprovacaoImpossivel), [stats]);
  const notaImpossivel = useMemo(() => stats.filter(s => s.aprovacaoImpossivel), [stats]);
  const hasAlerts = urgentes.length || proximas.length || aguardando.length || emRisco.length || notaAlta.length;
  return (
    <div className="flex flex-col gap-4">
      <div className="flex items-center justify-between"><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">Central de Alertas</p>{hasAlerts ? <Badge color="red">{urgentes.length + proximas.length + emRisco.length + notaAlta.length} alertas</Badge> : <Badge color="green">Tudo ok</Badge>}</div>
      {!hasAlerts && <div className="rounded-2xl p-8 border border-white/5 flex flex-col items-center gap-3" style={{ background: "rgba(255,255,255,0.03)" }}><span className="text-4xl">✅</span><p className="text-sm font-semibold text-white">Nenhum alerta ativo</p><p className="text-xs text-slate-500 text-center">Todas as disciplinas estão em dia.</p></div>}
      {notaImpossivel.length > 0 && <AlertSection title={`Aprovação Impossível (meta ${meta})`} color="#ef4444" icon="❌" count={notaImpossivel.length}>{notaImpossivel.map((s: DisciplinaStats) => <div key={s.disciplina} className="flex items-center justify-between px-4 py-2.5 border-b border-white/5 last:border-0"><div><p className="text-sm text-white">{s.disciplina}</p><p className="text-xs text-slate-500">Máximo: {s.notaMaxima?.toFixed(2) ?? "—"}</p></div><Badge color="red">Impossível</Badge></div>)}</AlertSection>}
      {urgentes.length > 0 && <AlertSection title="Urgente — Próximas 72 horas" color="#ef4444" icon="🚨" count={urgentes.length}>{urgentes.map((it: AtividadeEnriquecida) => <AlertRow key={it.id} it={it} color="#ef4444" />)}</AlertSection>}
      {emRisco.length > 0 && <AlertSection title="Disciplinas em Risco" color="#f43f5e" icon="⚠️" count={emRisco.length}>{emRisco.map((s: DisciplinaStats) => <div key={s.disciplina} className="flex items-center justify-between px-4 py-2.5 border-b border-white/5 last:border-0"><div><p className="text-sm text-white">{s.disciplina}</p><p className="text-xs text-slate-500">Nota nec.: {s.notaNecessaria?.toFixed(2) ?? "—"}</p></div><span className="text-sm font-bold text-red-400">{s.mediaAtual?.toFixed(2) ?? "—"}</span></div>)}</AlertSection>}
      {notaAlta.length > 0 && <AlertSection title="Nota Necessária Alta (> 8)" color="#f97316" icon="📈" count={notaAlta.length}>{notaAlta.map((s: DisciplinaStats) => <div key={s.disciplina} className="flex items-center justify-between px-4 py-2.5 border-b border-white/5 last:border-0"><div><p className="text-sm text-white">{s.disciplina}</p><p className="text-xs text-slate-500">Média: {s.mediaAtual?.toFixed(2) ?? "—"}</p></div><span className="text-sm font-bold text-orange-400">{s.notaNecessaria!.toFixed(2)} nec.</span></div>)}</AlertSection>}
      {proximas.length > 0 && <AlertSection title="Próximas — Em até 14 dias" color="#f59e0b" icon="⏰" count={proximas.length}>{proximas.map((it: AtividadeEnriquecida) => <AlertRow key={it.id} it={it} color="#f59e0b" />)}</AlertSection>}
      {aguardando.length > 0 && <AlertSection title="Aguardando Correção" color="#64748b" icon="⏳" count={aguardando.length}>{aguardando.map((it: AtividadeEnriquecida) => <AlertRow key={it.id} it={it} color="#94a3b8" />)}</AlertSection>}
    </div>
  );
}

// ─── CALENDAR TAB ─────────────────────────────────────────────
function CalendarTab({ items }: { items: AtividadeEnriquecida[] }) {
  const today = useMemo(() => { const d = new Date(); d.setHours(0,0,0,0); return d; }, []);
  const todayStr = useMemo(() => today.toISOString().split("T")[0], [today]);
  const [viewDate, setViewDate] = useState(() => new Date(today.getFullYear(), today.getMonth(), 1));
  const [selected, setSelected] = useState<number | null>(null);
  const year = viewDate.getFullYear(); const month = viewDate.getMonth();
  const MONTH_NAMES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
  const DAY_NAMES = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  const byDate = useMemo(() => { const m: Record<string, AtividadeEnriquecida[]> = {}; items.forEach(it => { if (!it.data) return; if (!m[it.data]) m[it.data] = []; m[it.data].push(it); }); return m; }, [items]);
  const cells = useMemo(() => { const firstDay = new Date(year, month, 1).getDay(); const daysInMonth = new Date(year, month + 1, 0).getDate(); const arr: (number | null)[] = []; for (let i = 0; i < firstDay; i++) arr.push(null); for (let d = 1; d <= daysInMonth; d++) arr.push(d); return arr; }, [year, month]);
  const getDateStr = useCallback((d: number) => `${year}-${String(month+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`, [year, month]);
  const getDotColor = useCallback((d: number) => {
    const dateStr = getDateStr(d);
    const its = byDate[dateStr];
    if (!its || its.length === 0) return null;
    // Prioridade visual: usa a categoria de maior urgência entre os itens do dia
    const ordem: CatId[] = ["ate3","ate7","aguardando","ate14","ate30","concluido"];
    let best: { color: string; rank: number } | null = null;
    its.forEach(it => {
      const cat = classificarItem(it);
      const rank = ordem.indexOf(cat.id);
      const r = rank === -1 ? 99 : rank;
      if (best === null || r < best.rank) best = { color: cat.color, rank: r };
    });
    return best !== null ? (best as { color: string; rank: number }).color : "#64748b";
  }, [byDate, getDateStr]);
  const selectedStr = selected ? getDateStr(selected) : null;
  const selectedItems = selectedStr ? (byDate[selectedStr] || []) : [];
  const upcomingList = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0).sort((a, b) => (a.daysRemaining ?? 0) - (b.daysRemaining ?? 0)).slice(0, 8), [items]);
  return (
    <div className="flex flex-col gap-4">
      <div className="flex items-center justify-between"><button onClick={() => setViewDate(new Date(year, month-1, 1))} className="w-9 h-9 rounded-xl flex items-center justify-center text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }}>‹</button><p className="text-sm font-bold text-white">{MONTH_NAMES[month]} {year}</p><button onClick={() => setViewDate(new Date(year, month+1, 1))} className="w-9 h-9 rounded-xl flex items-center justify-center text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }}>›</button></div>
      <div className="rounded-2xl overflow-hidden border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}><div className="grid grid-cols-7 border-b border-white/5">{DAY_NAMES.map(d => <div key={d} className="py-2 text-center text-xs font-semibold text-slate-500">{d}</div>)}</div><div className="grid grid-cols-7">{cells.map((d, i) => { if (!d) return <div key={`e${i}`} className="h-12" />; const dateStr = getDateStr(d); const isToday = dateStr === todayStr; const isSelected = d === selected; const dotColor = getDotColor(d); const count = byDate[dateStr]?.length || 0; return (<button key={d} onClick={() => setSelected(d === selected ? null : d)} className="h-12 flex flex-col items-center justify-center gap-0.5 transition-all" style={{ background: isSelected ? "rgba(99,102,241,0.25)" : isToday ? "rgba(99,102,241,0.1)" : "transparent" }}><span className={`text-xs font-semibold ${isToday ? "text-indigo-400" : isSelected ? "text-white" : "text-slate-300"}`} style={isToday ? { fontWeight: 800 } : {}}>{d}</span>{dotColor && <div className="flex gap-0.5">{count > 3 ? <div className="w-4 h-1 rounded-full" style={{ background: dotColor }} /> : Array.from({ length: Math.min(count, 3) }).map((_, j) => <div key={j} className="w-1 h-1 rounded-full" style={{ background: dotColor }} />)}</div>}</button>); })}</div></div>
      <div className="flex flex-wrap gap-3">{CATEGORIAS.map(c => (<div key={c.id} className="flex items-center gap-1.5"><div className="w-2 h-2 rounded-full" style={{ background: c.color }} /><span className="text-xs text-slate-500">{c.label}</span></div>))}</div>
      {selected && <div className="rounded-2xl border border-white/5 overflow-hidden" style={{ background: "rgba(255,255,255,0.04)" }}><div className="px-4 py-3 border-b border-white/5"><p className="text-sm font-bold text-white">{selected} de {MONTH_NAMES[month]}</p><p className="text-xs text-slate-500">{selectedItems.length === 0 ? "Sem atividades" : `${selectedItems.length} atividade(s)`}</p></div>{selectedItems.length === 0 ? <div className="px-4 py-6 text-center"><p className="text-xs text-slate-500">Nenhuma atividade neste dia</p></div> : <div className="flex flex-col">{selectedItems.map(it => { const nota = calcNota(it.pontuacao, it.pontuacaoMaxima); const cat = classificarItem(it); const isEvento = it.tipo === "evento"; return (<div key={it.id} className="flex items-center gap-3 px-4 py-3 border-b border-white/5 last:border-0"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: cat.color }} /><div className="flex-1 min-w-0"><p className="text-sm text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {isEvento ? "📅 Evento" : it.avaliacao} · <span style={{ color: cat.color }}>{cat.label}</span></p></div><div className="text-right shrink-0">{isEvento ? <span className="text-xs text-cyan-400">evento</span> : nota != null ? <p className="text-sm font-bold" style={{ color: notaColor(nota, 7) }}>{nota.toFixed(1)}/10</p> : <StatusBadge status={it.status} />}</div></div>); })}</div>}</div>}
      <div><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold mb-3">Próximas Atividades</p><div className="flex flex-col gap-2">{upcomingList.map(it => (<div key={it.id} className="rounded-xl px-4 py-3 border border-white/5 flex items-center gap-3" style={{ background: "rgba(255,255,255,0.03)" }}><div className="w-10 h-10 rounded-xl flex flex-col items-center justify-center shrink-0" style={{ background: (it.daysRemaining ?? 99) <= 3 ? "rgba(239,68,68,0.15)" : (it.daysRemaining ?? 99) <= 7 ? "rgba(249,115,22,0.15)" : "rgba(99,102,241,0.15)" }}><span className="text-xs font-bold" style={{ color: (it.daysRemaining ?? 99) <= 3 ? "#ef4444" : (it.daysRemaining ?? 99) <= 7 ? "#f97316" : "#818cf8" }}>{it.daysRemaining}d</span></div><div className="flex-1 min-w-0"><p className="text-sm font-medium text-white truncate">{it.instrumento}</p><p className="text-xs text-slate-500">{it.disciplina} · {it.data}</p></div><StatusBadge status={it.status} /></div>))}</div></div>
    </div>
  );
}

// ─── PLANEJAMENTO TAB ─────────────────────────────────────────
function projetarMedia(s: DisciplinaStats, mult: number | null, customScores: Record<string, number> = {}): number | null {
  let ws = 0, tw = 0;
  s.items.forEach((r: any) => {
    const w = r.pesoAvaliacao * r.pesoInstrumento; tw += w;
    let n = calcNota(r.pontuacao, r.pontuacaoMaxima);
    if (n == null) n = customScores[r.id] != null ? customScores[r.id] : (mult != null ? 10 * mult : 7.5);
    if (n != null) ws += n * w;
  });
  return tw > 0 ? Math.round((ws / tw) * 100) / 100 : null;
}

const CENARIO_OPTS = [
  { id: "pessimista", label: "Pessimista", desc: "5,0 em tudo", icon: "📉", color: "#ef4444" },
  { id: "realista", label: "Realista", desc: "7,5 em tudo", icon: "📊", color: "#f59e0b" },
  { id: "otimista", label: "Otimista", desc: "10,0 em tudo", icon: "📈", color: "#10b981" },
  { id: "custom", label: "Personalizado", desc: "Defina abaixo", icon: "🎛️", color: "#8b5cf6" },
];
const META_OPTS = [{ label: "Meta 6", value: 6 }, { label: "Meta 7", value: 7 }, { label: "Meta 8", value: 8 }, { label: "Meta 9", value: 9 }];

function scColor(s: any) {
  if (s.aprovacaoGarantida) return "#10b981";
  if (s.aprovacaoImpossivel) return "#ef4444";
  if (s.notaNecessaria == null) return "#64748b";
  if (s.notaNecessaria > 9) return "#ef4444";
  if (s.notaNecessaria > 7) return "#f97316";
  return "#f59e0b";
}
function prColor(v: number | null, meta: number) {
  if (v == null) return "#64748b";
  return v >= meta ? "#10b981" : v >= meta - 1.5 ? "#f59e0b" : "#ef4444";
}

function PlanejamentoTab({ stats, meta, onChangeMeta }: { stats: DisciplinaStats[]; meta: number; onChangeMeta: (m: number) => void }) {
  const [simScores, setSimScores] = useState<Record<string, number>>({});
  const [cenario, setCenario] = useState("realista");
  const discStats = useMemo(() => stats.map(s => ({ ...s, proj: { pessimista: projetarMedia(s, 0.5), realista: projetarMedia(s, 0.75), otimista: projetarMedia(s, 1.0), custom: projetarMedia(s, null, simScores) } })), [stats, simScores]);
  const mediaGeral = useMemo(() => { const g = discStats.filter(s => s.mediaAtual != null); return g.length ? g.reduce((a, s) => a + (s.mediaAtual ?? 0), 0) / g.length : null; }, [discStats]);
  const mediaProjetada = useMemo(() => { const vals = discStats.map(s => (s as any).proj[cenario]).filter((v: any) => v != null); return vals.length ? vals.reduce((a: number, v: number) => a + v, 0) / vals.length : null; }, [discStats, cenario]);
  const seguras = useMemo(() => discStats.filter(s => s.aprovacaoGarantida), [discStats]);
  const emRisco = useMemo(() => discStats.filter(s => !s.aprovacaoGarantida && !s.aprovacaoImpossivel && s.notaNecessaria != null && s.notaNecessaria > 0), [discStats]);
  const impossiveis = useMemo(() => discStats.filter(s => s.aprovacaoImpossivel), [discStats]);

  function MediaBar({ value, color }: { value: number | null; color: string }) {
    const pct = value != null ? Math.min(100, (value / 10) * 100) : 0;
    const metaPct = (meta / 10) * 100;
    return (
      <div className="relative w-full rounded-lg overflow-hidden" style={{ height: 8, background: "rgba(255,255,255,0.07)" }}>
        <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 999, transition: "width 0.5s ease" }} />
        <div style={{ position: "absolute", left: `${metaPct}%`, top: 0, bottom: 0, width: 2, background: "#f59e0b", transform: "translateX(-1px)" }} />
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-5">
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.04)" }}>
        <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold mb-3">Meta de Aprovação</p>
        <div className="flex gap-2 flex-wrap">{META_OPTS.map(o => <button key={o.value} onClick={() => onChangeMeta(o.value)} className={`flex-1 min-w-[60px] py-2.5 rounded-xl text-sm font-bold transition border ${meta === o.value ? "text-white border-indigo-500" : "text-slate-400 border-white/10 hover:border-white/20"}`} style={meta === o.value ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : { background: "rgba(255,255,255,0.04)" }}>{o.label}</button>)}</div>
        <div className="mt-3 flex items-center gap-3"><span className="text-xs text-slate-500 shrink-0">Personalizada:</span><input type="number" min="1" max="10" step="0.5" value={meta} onChange={e => { const v = parseFloat(e.target.value); if (!isNaN(v) && v >= 1 && v <= 10) onChangeMeta(v); }} className="w-24 rounded-xl px-3 py-1.5 text-sm text-white border border-white/10 outline-none focus:border-indigo-500" style={{ background: "rgba(255,255,255,0.06)" }} /><span className="text-xs text-slate-600">de 10</span></div>
      </div>
      <div className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl p-4 border border-white/5 flex flex-col gap-1" style={{ background: "rgba(255,255,255,0.04)" }}><span className="text-xs text-slate-500 uppercase tracking-widest">Média atual</span><span className="text-3xl font-bold" style={{ color: mediaGeral != null ? prColor(mediaGeral, meta) : "#64748b" }}>{mediaGeral != null ? mediaGeral.toFixed(2) : "—"}</span><span className="text-xs text-slate-500">das disciplinas avaliadas</span></div>
        <div className="rounded-2xl p-4 border border-white/5 flex flex-col gap-1" style={{ background: "rgba(255,255,255,0.04)" }}><span className="text-xs text-slate-500 uppercase tracking-widest">Meta</span><span className="text-3xl font-bold text-amber-400">{meta.toFixed(1)}</span><span className="text-xs text-slate-500">{seguras.length}/{discStats.length} garantidas</span></div>
      </div>
      <div className="rounded-2xl border border-white/5 overflow-hidden" style={{ background: "rgba(255,255,255,0.04)" }}>
        <div className="px-4 py-3 border-b border-white/5"><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">Situação por Disciplina</p></div>
        <div className="flex flex-col">{discStats.map((s, i) => { const sc = scColor(s); return (<div key={s.disciplina} className={`px-4 py-3 ${i < discStats.length - 1 ? "border-b border-white/5" : ""}`}><div className="flex items-center justify-between mb-2"><div className="flex items-center gap-2 min-w-0"><div className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} /><span className="text-sm font-bold text-white truncate">{s.disciplina}</span></div><div className="flex items-center gap-2 shrink-0 ml-2">{s.aprovacaoGarantida && <Badge color="green">✓ Garantida</Badge>}{s.aprovacaoImpossivel && <Badge color="red">✗ Impossível</Badge>}{!s.aprovacaoGarantida && !s.aprovacaoImpossivel && s.notaNecessaria != null && <span className="text-xs font-bold" style={{ color: sc }}>Nec. {s.notaNecessaria.toFixed(2)}</span>}{s.mediaAtual == null && <Badge color="slate">Sem dados</Badge>}</div></div><div className="flex items-center gap-2"><span className="text-xs text-slate-500 w-10 shrink-0">{s.mediaAtual != null ? s.mediaAtual.toFixed(1) : "—"}</span><div className="flex-1"><MediaBar value={s.mediaAtual} color={s.color} /></div><span className="text-xs text-slate-600 w-12 shrink-0 text-right">máx {s.notaMaxima != null ? s.notaMaxima.toFixed(1) : "10.0"}</span></div>{!s.aprovacaoGarantida && !s.aprovacaoImpossivel && s.notaNecessaria != null && s.notaNecessaria > 0 && <p className="text-xs mt-1.5" style={{ color: sc }}>{s.notaNecessaria > 10 ? "Aprovação matematicamente impossível." : `Precisa de ${s.notaNecessaria.toFixed(2)} nas ${s.items.filter((r: any) => r.pontuacao == null).length} avaliações restantes.`}</p>}</div>); })}</div>
      </div>
      {(seguras.length > 0 || emRisco.length > 0 || impossiveis.length > 0) && (
        <div className="flex flex-col gap-3">
          {seguras.length > 0 && <div className="rounded-2xl p-4 border border-emerald-500/20" style={{ background: "rgba(16,185,129,0.06)" }}><div className="flex items-center gap-2 mb-3"><span>✅</span><p className="text-sm font-bold text-emerald-400">Aprovação Garantida ({seguras.length})</p></div>{seguras.map(s => <div key={s.disciplina} className="flex items-center justify-between"><div className="flex items-center gap-2"><div className="w-1.5 h-1.5 rounded-full" style={{ background: s.color }} /><span className="text-sm text-white">{s.disciplina}</span></div><span className="text-sm font-bold text-emerald-400">{s.mediaAtual?.toFixed(2)}</span></div>)}</div>}
          {emRisco.length > 0 && <div className="rounded-2xl p-4 border border-amber-500/20" style={{ background: "rgba(245,158,11,0.06)" }}><div className="flex items-center gap-2 mb-3"><span>⚠️</span><p className="text-sm font-bold text-amber-400">Atenção Necessária ({emRisco.length})</p></div>{emRisco.map(s => <div key={s.disciplina} className="flex items-center justify-between"><div className="flex items-center gap-2 min-w-0"><div className="w-1.5 h-1.5 rounded-full" style={{ background: s.color }} /><span className="text-sm text-white truncate">{s.disciplina}</span></div><div className="flex items-center gap-2 shrink-0"><span className="text-xs text-slate-500">{s.mediaAtual != null ? `${s.mediaAtual.toFixed(1)} →` : ""}</span><span className="text-sm font-bold" style={{ color: scColor(s) }}>nec. {s.notaNecessaria?.toFixed(2)}</span></div></div>)}</div>}
          {impossiveis.length > 0 && <div className="rounded-2xl p-4 border border-red-500/30" style={{ background: "rgba(239,68,68,0.08)" }}><div className="flex items-center gap-2 mb-3"><span>❌</span><p className="text-sm font-bold text-red-400">Aprovação Impossível ({impossiveis.length})</p></div>{impossiveis.map(s => <div key={s.disciplina} className="flex items-center justify-between"><div className="flex items-center gap-2 min-w-0"><div className="w-1.5 h-1.5 rounded-full" style={{ background: s.color }} /><span className="text-sm text-white truncate">{s.disciplina}</span></div><span className="text-xs text-red-400">máx {s.notaMaxima?.toFixed(2)}</span></div>)}</div>}
        </div>
      )}
      <div className="rounded-2xl border border-white/5 overflow-hidden" style={{ background: "rgba(255,255,255,0.04)" }}>
        <div className="px-4 py-3 border-b border-white/5"><p className="text-xs text-slate-500 uppercase tracking-widest font-semibold">Simulação de Cenários</p><p className="text-xs text-slate-600 mt-0.5">Projeção da média final nas avaliações restantes</p></div>
        <div className="grid grid-cols-2 gap-2 p-4 pb-0">{CENARIO_OPTS.map(c => <button key={c.id} onClick={() => setCenario(c.id)} className="rounded-xl p-3 text-left border transition" style={{ background: cenario === c.id ? `${c.color}15` : "rgba(255,255,255,0.03)", borderColor: cenario === c.id ? c.color + "44" : "rgba(255,255,255,0.08)" }}><div className="flex items-center gap-1.5 mb-1"><span className="text-sm">{c.icon}</span><span className="text-xs font-bold text-white">{c.label}</span></div><p className="text-xs" style={{ color: c.color }}>{c.desc}</p></button>)}</div>
        {cenario === "custom" && (
          <div className="px-4 pt-4 pb-2 flex flex-col gap-3">
            <p className="text-xs text-slate-500">Arraste para definir a nota esperada em cada avaliação pendente:</p>
            {discStats.map(s => { const pending = s.items.filter((r: any) => r.pontuacao == null); if (pending.length === 0) return null; return (<div key={s.disciplina}><div className="flex items-center gap-2 mb-2"><div className="w-2 h-2 rounded-full" style={{ background: s.color }} /><span className="text-xs font-bold text-white">{s.disciplina}</span></div>{pending.map((it: any) => (<div key={it.id} className="mb-3 pl-4"><div className="flex items-center justify-between mb-1"><span className="text-xs text-slate-400 truncate pr-2">{it.instrumento}</span><span className="text-xs font-bold text-indigo-400 shrink-0">{(simScores[it.id] ?? 7.5).toFixed(1)}</span></div><input type="range" min="0" max="10" step="0.5" value={simScores[it.id] ?? 7.5} onChange={e => setSimScores(sc => ({ ...sc, [it.id]: parseFloat(e.target.value) }))} className="w-full" /><div className="flex justify-between text-xs text-slate-700 mt-0.5"><span>0</span><span>5</span><span>10</span></div></div>))}</div>); })}
          </div>
        )}
        <div className="px-4 py-4 flex flex-col gap-2">
          {discStats.map(s => { const projVal = (s as any).proj[cenario]; const hasPending = s.items.some((r: any) => r.pontuacao == null); if (!hasPending && s.mediaAtual != null) return (<div key={s.disciplina} className="flex items-center justify-between py-1.5 border-b border-white/5 last:border-0"><div className="flex items-center gap-2 min-w-0"><div className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: s.color }} /><span className="text-xs text-slate-400 truncate">{s.disciplina}</span></div><div className="flex items-center gap-2 shrink-0"><span className="text-xs text-slate-600">já concluída</span><span className="text-sm font-bold" style={{ color: s.mediaAtual >= meta ? "#10b981" : "#ef4444" }}>{s.mediaAtual.toFixed(2)}</span></div></div>); return (<div key={s.disciplina} className="py-1.5 border-b border-white/5 last:border-0"><div className="flex items-center justify-between mb-1"><div className="flex items-center gap-2 min-w-0"><div className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: s.color }} /><span className="text-xs text-white font-medium truncate">{s.disciplina}</span></div><div className="flex items-center gap-2 shrink-0">{s.mediaAtual != null && <span className="text-xs text-slate-600">{s.mediaAtual.toFixed(1)} →</span>}<span className="text-sm font-bold" style={{ color: prColor(projVal, meta) }}>{projVal != null ? projVal.toFixed(2) : "—"}</span><Badge color={projVal == null ? "slate" : projVal >= meta ? "green" : "red"}>{projVal == null ? "—" : projVal >= meta ? "✓" : "✗"}</Badge></div></div><ProgressBar value={projVal ?? 0} max={10} color={prColor(projVal, meta)} height={3} /></div>); })}
          <div className="mt-3 rounded-xl p-3 border border-white/10 flex items-center justify-between" style={{ background: "rgba(99,102,241,0.08)" }}><div><p className="text-xs text-slate-400">Média projetada geral</p><p className="text-xs text-slate-600 mt-0.5">Cenário: {CENARIO_OPTS.find(c => c.id === cenario)?.label}</p></div><div className="text-right">{mediaGeral != null && <p className="text-xs text-slate-500">Atual: {mediaGeral.toFixed(2)}</p>}<p className="text-xl font-bold" style={{ color: prColor(mediaProjetada, meta) }}>{mediaProjetada != null ? mediaProjetada.toFixed(2) : "—"}</p></div></div>
          {mediaProjetada != null && <div className="flex items-center gap-3 mt-1"><div className="flex-1 rounded-lg p-2 text-center border border-emerald-500/20" style={{ background: "rgba(16,185,129,0.08)" }}><p className="text-lg font-bold text-emerald-400">{discStats.filter(s => ((s as any).proj[cenario] ?? 0) >= meta).length}</p><p className="text-xs text-slate-500">aprovadas</p></div><div className="flex-1 rounded-lg p-2 text-center border border-red-500/20" style={{ background: "rgba(239,68,68,0.08)" }}><p className="text-lg font-bold text-red-400">{discStats.filter(s => ((s as any).proj[cenario] ?? 0) < meta).length}</p><p className="text-xs text-slate-500">reprovadas</p></div></div>}
        </div>
      </div>
    </div>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────
// ─── QUICK GRADE PANEL (lançamento rápido de nota) ────────────
function QuickGradePanel({ item, stats, meta, onSave, onClose }: {
  item: Atividade; stats: DisciplinaStats[]; meta: number;
  onSave: (form: Omit<Atividade, "id">) => void | Promise<void>; onClose: () => void;
}) {
  const max = item.pontuacaoMaxima && item.pontuacaoMaxima > 0 ? item.pontuacaoMaxima : 10;
  const toStr = (v: number) => (v % 1 === 0 ? String(v) : String(v).replace(".", ","));
  const inicial = item.pontuacao ?? Math.round(max * 0.7 * 100) / 100;
  const [pontuacao, setPontuacao] = useState<number>(inicial);
  const [pontTexto, setPontTexto] = useState<string>(toStr(inicial));   // digitação livre
  const [saving, setSaving] = useState(false);
  const disc = stats.find(s => s.disciplina === item.disciplina);
  const nota10 = (pontuacao / max) * 10;
  const pct = (pontuacao / max) * 100;

  // Sincroniza o texto digitado → número (aceita vírgula e qualquer casa decimal)
  const onChangeTexto = (txt: string) => {
    setPontTexto(txt);
    const n = parseDecimal(txt);
    if (n != null) {
      const clamped = Math.max(0, Math.min(max, n));
      setPontuacao(clamped);
    }
  };
  // Slider → número e texto
  const onChangeSlider = (v: number) => {
    setPontuacao(v);
    setPontTexto(toStr(Math.round(v * 100) / 100));
  };

  // Projeção: recalcula a média da disciplina substituindo esta nota
  const mediaProjetada = useMemo(() => {
    if (!disc) return null;
    let ws = 0, tw = 0;
    disc.items.forEach((r) => {
      if (r.tipo === "evento") return;
      const w = r.pesoAvaliacao * r.pesoInstrumento;
      let n = calcNota(r.pontuacao, r.pontuacaoMaxima);
      if (r.id === item.id) n = nota10;     // usa a nota simulada para este item
      if (n != null) { ws += n * w; tw += w; }
    });
    return tw > 0 ? Math.round((ws / tw) * 100) / 100 : null;
  }, [disc, item.id, nota10]);

  const color = nota10 >= meta ? "#10b981" : nota10 >= meta - 2 ? "#f59e0b" : "#ef4444";
  const handleSave = async () => {
    if (saving) return;
    setSaving(true);
    try {
      await onSave({ ...item, pontuacao, pontuacaoMaxima: max, status: "Finalizado" });
    } finally { setSaving(false); }
  };

  return (
    <div className="flex flex-col gap-5">
      <div>
        <p className="text-sm font-bold text-white">{item.instrumento}</p>
        <p className="text-xs text-slate-500">{item.disciplina} · {item.avaliacao}</p>
      </div>

      {/* Campo de digitação livre da pontuação */}
      <FormField label={`Pontuação obtida (0 a ${toStr(max)})`}>
        <input type="text" inputMode="decimal" value={pontTexto} onChange={e => onChangeTexto(e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder={`Ex: 8,75`} />
      </FormField>

      {/* Nota grande + percentual */}
      <div className="rounded-2xl p-5 border border-white/5 flex items-center justify-around" style={{ background: "rgba(255,255,255,0.04)" }}>
        <div className="text-center"><p className="text-xs text-slate-500 mb-1">Pontuação</p><p className="text-2xl font-bold text-white">{pontuacao % 1 === 0 ? pontuacao : pontuacao.toFixed(2)}<span className="text-sm text-slate-500">/{toStr(max)}</span></p></div>
        <div className="text-center"><p className="text-xs text-slate-500 mb-1">Nota /10</p><p className="text-2xl font-bold" style={{ color }}>{nota10.toFixed(2)}</p></div>
        <div className="text-center"><p className="text-xs text-slate-500 mb-1">Percentual</p><p className="text-2xl font-bold" style={{ color }}>{pct.toFixed(0)}%</p></div>
      </div>

      {/* Slider (passo fino, ajuste rápido) */}
      <div>
        <input type="range" min="0" max={max} step={max <= 10 ? 0.05 : max <= 100 ? 0.5 : 1} value={pontuacao} onChange={e => onChangeSlider(parseFloat(e.target.value))} className="w-full" style={{ accentColor: color }} />
        <div className="flex justify-between text-xs text-slate-600 mt-1"><span>0</span><span>{(max/2) % 1 === 0 ? max/2 : (max/2).toFixed(1)}</span><span>{toStr(max)}</span></div>
      </div>

      {/* Impacto na média */}
      {disc && (
        <div className="rounded-xl p-3 border border-white/10 flex items-center justify-between" style={{ background: "rgba(99,102,241,0.08)" }}>
          <div><p className="text-xs text-slate-400">Média da disciplina</p><p className="text-xs text-slate-600 mt-0.5">{disc.mediaAtual != null ? `Atual: ${disc.mediaAtual.toFixed(2)}` : "Sem notas ainda"}</p></div>
          <div className="text-right"><p className="text-xs text-slate-500">com esta nota</p><p className="text-xl font-bold" style={{ color: mediaProjetada == null ? "#64748b" : mediaProjetada >= meta ? "#10b981" : "#ef4444" }}>{mediaProjetada != null ? mediaProjetada.toFixed(2) : "—"}</p></div>
        </div>
      )}

      <div className="flex gap-2">
        <button onClick={onClose} disabled={saving} className="flex-1 py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10 disabled:opacity-40">Cancelar</button>
        <button onClick={handleSave} disabled={saving} className="flex-1 py-3 rounded-xl text-sm font-semibold text-white disabled:opacity-60" style={{ background: "linear-gradient(135deg,#10b981,#059669)" }}>{saving ? "Salvando..." : "Lançar Nota"}</button>
      </div>
    </div>
  );
}

// ─── NOVA DISCIPLINA (cadastro em lote) ───────────────────────
const TIPO_DISCIPLINA = ["Teórica", "Prática", "Mista"];
const TIPO_AVALIACAO_LOTE = ["AP1", "AP2", "AS", "AF", "Seminário", "Projeto", "Avaliação Prática", "Avaliação Teórica"];

// Nova Disciplina (Fase D): cria a entidade Disciplina de verdade.
// Avaliações são OPCIONAIS — pode salvar disciplina vazia e lançar depois.
function NovaDisciplinaForm({ onSaveDisciplina, onClose }: {
  onSaveDisciplina: (
    disc: { nome: string; tipo: TipoDisciplina; observacoes: string; pesoTeorica: number; pesoPratica: number; subdivisoes: string[] },
    avaliacoes: { avaliacao: string; instrumento: string; parte: ParteDisciplina; subdivisao: string; data: string; pesoAvaliacao: number; pesoInstrumento: number; pontuacaoMaxima: number | null; observacoes: string }[]
  ) => void | Promise<void>;
  onClose: () => void;
}) {
  const [nome, setNome] = useState("");
  const [tipoDisc, setTipoDisc] = useState<TipoDisciplina>("Teórica");
  const [obsDisc, setObsDisc] = useState("");
  const [pesoTeorica, setPesoTeorica] = useState("60");   // só usado em Mista
  const [subdivisoes, setSubdivisoes] = useState<string[]>([]);
  const [novaSub, setNovaSub] = useState("");
  const [avaliacoes, setAvaliacoes] = useState<any[]>([]);
  const [saving, setSaving] = useState(false);
  const [erro, setErro] = useState("");

  const isMista = tipoDisc === "Mista";
  const pTeorica = parseDecimal(pesoTeorica) ?? 0;
  const pPratica = 100 - pTeorica;

  const addAvaliacao = (parte: ParteDisciplina) => setAvaliacoes(a => [...a, { _id: Date.now() + Math.random(), nome: "", tipo: "AP1", parte, subdivisao: "", data: "", peso: "", pesoInstrumento: "100", pontuacaoMaxima: "10", observacoes: "" }]);
  const removeAvaliacao = (id: number) => setAvaliacoes(a => a.filter(x => x._id !== id));
  const setAv = (id: number, k: string, v: any) => setAvaliacoes(a => a.map(x => x._id === id ? { ...x, [k]: v } : x));
  const addSub = () => { const s = novaSub.trim(); if (s && !subdivisoes.includes(s)) { setSubdivisoes(p => [...p, s]); setNovaSub(""); } };
  const removeSub = (s: string) => { setSubdivisoes(p => p.filter(x => x !== s)); setAvaliacoes(a => a.map(x => x.subdivisao === s ? { ...x, subdivisao: "" } : x)); };

  // Soma de pesos por parte (cada parte deve fechar 100%)
  const somaParte = (parte: ParteDisciplina) => avaliacoes.filter(a => a.parte === parte).reduce((acc, a) => acc + (parseDecimal(a.peso) ?? 0), 0);
  const somaUnica = somaParte("unica");
  const somaT = somaParte("teorica");
  const somaP = somaParte("pratica");

  const barColor = (soma: number, tem: boolean) => !tem ? "#64748b" : Math.abs(soma - 100) < 0.01 ? "#10b981" : soma > 100 ? "#ef4444" : "#f59e0b";

  const handleSave = async () => {
    if (saving) return;
    if (!nome.trim()) { setErro("Dê um nome à disciplina."); return; }
    if (isMista && (pTeorica < 0 || pTeorica > 100)) { setErro("Peso da parte teórica deve ficar entre 0 e 100."); return; }
    setErro("");
    setSaving(true);
    try {
      const disc = {
        nome: nome.trim(),
        tipo: tipoDisc,
        observacoes: obsDisc.trim(),
        pesoTeorica: isMista ? pTeorica : (tipoDisc === "Prática" ? 0 : 100),
        pesoPratica: isMista ? pPratica : (tipoDisc === "Prática" ? 100 : 0),
        subdivisoes,
      };
      const validas = avaliacoes.filter(a => a.nome.trim());
      const avs = validas.map(a => ({
        avaliacao: a.tipo,
        instrumento: a.nome.trim(),
        parte: (isMista ? a.parte : "unica") as ParteDisciplina,
        subdivisao: a.subdivisao || "",
        data: a.data || "",
        pesoAvaliacao: (parseDecimal(a.peso) ?? 0) / 100,
        pesoInstrumento: (parseDecimal(a.pesoInstrumento) ?? 100) / 100,
        pontuacaoMaxima: parseDecimal(a.pontuacaoMaxima),
        observacoes: a.observacoes.trim(),
      }));
      await onSaveDisciplina(disc, avs);
    } finally {
      setSaving(false);
    }
  };

  // Renderiza um grupo de avaliações de uma parte
  const renderAvaliacoes = (parte: ParteDisciplina, label: string) => {
    const lista = avaliacoes.filter(a => a.parte === parte);
    const soma = somaParte(parte);
    const cor = barColor(soma, lista.length > 0);
    return (
      <div className="rounded-2xl p-3 border border-white/10" style={{ background: "rgba(255,255,255,0.03)" }}>
        <div className="flex items-center justify-between mb-2">
          <span className="text-xs font-bold text-white">{label}</span>
          {lista.length > 0 && <span className="text-xs font-bold" style={{ color: cor }}>{soma.toFixed(0)}% · resta {(100 - soma).toFixed(0)}%</span>}
        </div>
        {lista.length > 0 && <div className="w-full rounded-full overflow-hidden mb-3" style={{ height: 5, background: "rgba(255,255,255,0.07)" }}><div style={{ width: `${Math.min(100, soma)}%`, height: "100%", background: cor, transition: "width 0.3s" }} /></div>}
        <div className="flex flex-col gap-3">
          {lista.map((a, idx) => (
            <div key={a._id} className="rounded-xl p-3 border border-white/10" style={{ background: "rgba(255,255,255,0.04)" }}>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-400">{label} · {idx + 1}</span>
                <button onClick={() => removeAvaliacao(a._id)} className="text-xs text-red-400">remover</button>
              </div>
              <div className="flex flex-col gap-2">
                <input value={a.nome} onChange={e => setAv(a._id, "nome", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Nome (ex: Prova 1)" />
                <div className="grid grid-cols-2 gap-2">
                  <select value={a.tipo} onChange={e => setAv(a._id, "tipo", e.target.value)} className={INPUT_CLS} style={INPUT_STY}>{TIPO_AVALIACAO_LOTE.map(t => <option key={t}>{t}</option>)}</select>
                  {subdivisoes.length > 0
                    ? <select value={a.subdivisao} onChange={e => setAv(a._id, "subdivisao", e.target.value)} className={INPUT_CLS} style={INPUT_STY}><option value="">Sem subdivisão</option>{subdivisoes.map(s => <option key={s}>{s}</option>)}</select>
                    : <input type="date" value={a.data} onChange={e => setAv(a._id, "data", e.target.value)} className={INPUT_CLS} style={INPUT_STY} />}
                </div>
                {subdivisoes.length > 0 && <input type="date" value={a.data} onChange={e => setAv(a._id, "data", e.target.value)} className={INPUT_CLS} style={INPUT_STY} />}
                <div className="grid grid-cols-2 gap-2">
                  <input type="text" inputMode="decimal" value={a.peso} onChange={e => setAv(a._id, "peso", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Peso na parte % (ex: 40)" />
                  <input type="text" inputMode="decimal" value={a.pontuacaoMaxima} onChange={e => setAv(a._id, "pontuacaoMaxima", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Pontuação máx (ex: 10)" />
                </div>
                <input value={a.observacoes} onChange={e => setAv(a._id, "observacoes", e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Observações (opcional)" />
              </div>
            </div>
          ))}
        </div>
        <button onClick={() => addAvaliacao(parte)} className="w-full mt-3 py-2.5 rounded-xl text-sm font-semibold text-indigo-300 border border-indigo-500/30" style={{ background: "rgba(99,102,241,0.1)" }}>+ Adicionar avaliação</button>
      </div>
    );
  };

  return (
    <div className="flex flex-col gap-5">
      {/* Dados da disciplina */}
      <div className="flex flex-col gap-3">
        <FormField label="Nome da disciplina"><input value={nome} onChange={e => setNome(e.target.value)} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: Clínica Médica" autoFocus /></FormField>
        <FormField label="Tipo da disciplina">
          <div className="flex rounded-xl overflow-hidden border border-white/10" style={{ background: "rgba(255,255,255,0.03)" }}>
            {TIPO_DISCIPLINA.map(t => (
              <button key={t} type="button" onClick={() => setTipoDisc(t as TipoDisciplina)} className={`flex-1 py-2 text-xs font-semibold transition ${tipoDisc === t ? "text-white" : "text-slate-500 hover:text-slate-300"}`} style={tipoDisc === t ? { background: "linear-gradient(135deg,#6366f1,#8b5cf6)" } : {}}>{t}</button>
            ))}
          </div>
        </FormField>
        {isMista && (
          <div className="rounded-xl p-3 border border-indigo-500/20" style={{ background: "rgba(99,102,241,0.06)" }}>
            <p className="text-xs text-slate-400 mb-2">Quanto cada parte vale da nota final:</p>
            <div className="flex items-center gap-3">
              <div className="flex-1"><label className="text-xs text-slate-500">Teórica %</label><input type="text" inputMode="decimal" value={pesoTeorica} onChange={e => setPesoTeorica(e.target.value)} className={INPUT_CLS} style={INPUT_STY} /></div>
              <span className="text-slate-600 mt-4">+</span>
              <div className="flex-1"><label className="text-xs text-slate-500">Prática %</label><div className={INPUT_CLS} style={{ ...INPUT_STY, opacity: 0.7 }}>{pPratica.toFixed(0)}</div></div>
            </div>
            <p className="text-xs mt-2" style={{ color: pTeorica >= 0 && pTeorica <= 100 ? "#10b981" : "#ef4444" }}>Teórica {pTeorica.toFixed(0)}% + Prática {pPratica.toFixed(0)}% = 100%</p>
          </div>
        )}
        <FormField label="Observações (opcional)"><textarea value={obsDisc} onChange={e => setObsDisc(e.target.value)} rows={2} className={INPUT_CLS} style={INPUT_STY} placeholder="Anotações sobre a disciplina (ficam salvas)" /></FormField>
      </div>

      {/* Subdivisões */}
      <div className="rounded-2xl p-4 border border-white/5" style={{ background: "rgba(255,255,255,0.03)" }}>
        <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold mb-1">Subdivisões (opcional)</p>
        <p className="text-xs text-slate-600 mb-2">Crie aqui; escolha em cada avaliação.</p>
        <div className="flex gap-2 mb-2">
          <input value={novaSub} onChange={e => setNovaSub(e.target.value)} onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); addSub(); } }} className={INPUT_CLS} style={INPUT_STY} placeholder="Ex: Cardio, Módulo A..." />
          <button onClick={addSub} className="px-3 rounded-xl text-sm font-semibold text-white shrink-0" style={{ background: "rgba(99,102,241,0.4)" }}>+</button>
        </div>
        {subdivisoes.length > 0 && <div className="flex flex-wrap gap-2">{subdivisoes.map(s => <span key={s} className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs text-white" style={{ background: "rgba(99,102,241,0.25)" }}>{s}<button onClick={() => removeSub(s)} className="text-slate-400 hover:text-white">✕</button></span>)}</div>}
      </div>

      {/* Avaliações — opcionais */}
      <div>
        <p className="text-xs text-slate-500 uppercase tracking-widest font-semibold mb-1">Avaliações (opcional)</p>
        <p className="text-xs text-slate-600 mb-3">Você pode salvar a disciplina vazia e lançar avaliações depois.</p>
        {isMista ? (
          <div className="flex flex-col gap-3">
            {renderAvaliacoes("teorica", `Parte Teórica (${pTeorica.toFixed(0)}%)`)}
            {renderAvaliacoes("pratica", `Parte Prática (${pPratica.toFixed(0)}%)`)}
          </div>
        ) : renderAvaliacoes("unica", "Avaliações")}
      </div>

      {erro && <p className="text-xs text-red-400">{erro}</p>}

      <div className="flex gap-2">
        <button onClick={onClose} disabled={saving} className="flex-1 py-3 rounded-xl text-sm font-semibold text-slate-400 border border-white/10 disabled:opacity-40">Cancelar</button>
        <button onClick={handleSave} disabled={saving} className="flex-1 py-3 rounded-xl text-sm font-semibold text-white disabled:opacity-60" style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)" }}>{saving ? "Salvando..." : "Salvar Disciplina"}</button>
      </div>
    </div>
  );
}

const TABS = [
  { id: "dashboard", label: "Dashboard", icon: "⚡" },
  { id: "disciplines", label: "Disciplinas", icon: "📖" },
  { id: "plano", label: "Planejamento", icon: "🎯" },
  { id: "alerts", label: "Alertas", icon: "🔔" },
  { id: "calendar", label: "Calendário", icon: "📅" },
];

export default function App() {
  const { user, loading: authLoading, signIn, signUp, signOut } = useAuth();
  const { atividades, loading: dataLoading, error, addAtividade, updateAtividade, deleteAtividade, importAtividades } = useAtividades(user);
  const { profile, updateMeta } = useProfile(user);
  const { disciplinas, addDisciplina } = useDisciplinas(user);

  const [tab, setTab] = useState("dashboard");
  const [modal, setModal] = useState<string | null>(null);
  const [editingItem, setEditingItem] = useState<Atividade | null>(null);
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [localMeta, setLocalMeta] = useState(7);
  const [showAddMenu, setShowAddMenu] = useState(false);          // menu do botão "+"
  const [novoTipo, setNovoTipo] = useState<"avaliacao" | "evento">("avaliacao"); // pré-seleção do ItemForm

  const meta = profile?.metaAprovacao ?? localMeta;

  const handleChangeMeta = useCallback((m: number) => {
    setLocalMeta(m);
    updateMeta(m);
  }, [updateMeta]);

  const items = useMemo(() => atividades.map(enrichItem), [atividades]);
  const stats = useMemo(() => getDisciplineStats(items, meta, disciplinas), [items, meta, disciplinas]);
  const disciplines = useMemo<string[]>(() => {
    const nomes = new Set<string>(items.map(it => it.disciplina));
    disciplinas.forEach(d => nomes.add(d.nome));
    return Array.from(nomes).filter(Boolean);
  }, [items, disciplinas]);

  const handleSave = useCallback(async (form: Omit<Atividade, "id">) => {
    // Erros sobem para o ItemForm (que libera o botão); modal só fecha no sucesso
    if (editingItem) await updateAtividade({ ...form, id: editingItem.id });
    else await addAtividade(form);
    setModal(null); setEditingItem(null);
  }, [editingItem, addAtividade, updateAtividade]);

  // Salvamento em lote (fluxo Nova Disciplina) — cria todas as avaliações de uma vez
  const handleSaveDisciplina = useCallback(async (
    disc: { nome: string; tipo: TipoDisciplina; observacoes: string; pesoTeorica: number; pesoPratica: number; subdivisoes: string[] },
    avs: { avaliacao: string; instrumento: string; parte: ParteDisciplina; subdivisao: string; data: string; pesoAvaliacao: number; pesoInstrumento: number; pontuacaoMaxima: number | null; observacoes: string }[]
  ) => {
    // 1. Cria a disciplina na tabela
    const novaDisc = await addDisciplina(disc);
    // 2. Cria as avaliações ligadas a ela (se houver)
    if (novaDisc) {
      for (const a of avs) {
        await addAtividade({
          tipo: "avaliacao",
          avaliacao: a.avaliacao,
          instrumento: a.instrumento,
          disciplina: disc.nome,
          disciplinaId: novaDisc.id,
          parte: a.parte,
          subdivisao: a.subdivisao,
          status: "Não iniciado",
          data: a.data,
          pesoAvaliacao: a.pesoAvaliacao,
          pesoInstrumento: a.pesoInstrumento,
          pontuacaoMaxima: a.pontuacaoMaxima,
          pontuacao: null,
          observacoes: a.observacoes,
        });
      }
    }
    setModal(null);
  }, [addDisciplina, addAtividade]);

  // Abre o ItemForm já com o tipo certo (avaliação ou evento)
  const openNovaAvaliacao = useCallback(() => { setEditingItem(null); setNovoTipo("avaliacao"); setShowAddMenu(false); setModal("add"); }, []);
  const openNovoEvento = useCallback(() => { setEditingItem(null); setNovoTipo("evento"); setShowAddMenu(false); setModal("add"); }, []);
  const openNovaDisciplina = useCallback(() => { setShowAddMenu(false); setModal("novadisc"); }, []);

  const handleEdit = useCallback((item: Atividade) => { setEditingItem(item); setModal("edit"); }, []);
  const handleQuickGrade = useCallback((item: Atividade) => { setEditingItem(item); setModal("quickgrade"); }, []);
  const handleDelete = useCallback((id: string) => { setDeleteId(id); setModal("delete"); }, []);
  const confirmDelete = useCallback(async () => {
    if (deleteId) await deleteAtividade(deleteId);
    setModal(null); setDeleteId(null);
  }, [deleteId, deleteAtividade]);

  const urgentCount = useMemo(() => items.filter(it => it.daysRemaining != null && it.daysRemaining >= 0 && it.daysRemaining <= 3 && it.pontuacao == null).length, [items]);

  if (authLoading) return <div className="min-h-screen flex items-center justify-center" style={{ background: "#0a0f1a" }}><p className="text-slate-500 text-sm">Carregando...</p></div>;
  if (!user) return <Login onSignIn={signIn} onSignUp={signUp} />;

  return (
    <div className="min-h-screen flex flex-col" style={{ background: "#0a0f1a", fontFamily: "'DM Sans',system-ui,sans-serif" }}>
      <header className="sticky top-0 z-40 border-b border-white/5 header-safe" style={{ background: "rgba(10,15,26,0.95)", backdropFilter: "blur(20px)" }}>
        <div className="max-w-2xl mx-auto px-4 py-3 flex items-center justify-between">
          <div className="min-w-0"><h1 className="text-base font-bold text-white tracking-tight truncate">Controle do Semestre</h1><p className="text-xs text-slate-500 truncate">{items.length} atividades · {disciplines.length} disciplinas · meta {meta}</p></div>
          <div className="flex gap-2 shrink-0">
            <button onClick={() => setModal("import")} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Importar">📂</button>
            <button onClick={() => setModal("export")} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Exportar">📥</button>
            <button onClick={signOut} className="w-8 h-8 rounded-xl flex items-center justify-center text-sm text-slate-400 hover:text-white transition" style={{ background: "rgba(255,255,255,0.06)" }} title="Sair">🚪</button>
            <div className="relative">
              <button onClick={() => setShowAddMenu(v => !v)} className="px-3 h-8 rounded-xl text-xs font-semibold text-white flex items-center gap-1" style={{ background: "linear-gradient(135deg,#6366f1,#8b5cf6)" }}>+ Novo</button>
              {showAddMenu && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setShowAddMenu(false)} />
                  <div className="absolute right-0 mt-2 w-48 rounded-2xl border border-white/10 overflow-hidden z-50 shadow-2xl" style={{ background: "#111827" }}>
                    <button onClick={openNovaDisciplina} className="w-full flex items-center gap-3 px-4 py-3 text-left text-sm font-medium text-white hover:bg-white/5 transition"><span className="text-base">📚</span>Nova Disciplina</button>
                    <button onClick={openNovaAvaliacao} className="w-full flex items-center gap-3 px-4 py-3 text-left text-sm font-medium text-white hover:bg-white/5 transition"><span className="text-base">📊</span>Nova Avaliação</button>
                    <button onClick={openNovoEvento} className="w-full flex items-center gap-3 px-4 py-3 text-left text-sm font-medium text-white hover:bg-white/5 transition"><span className="text-base">📅</span>Novo Evento</button>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      </header>

      {error && <div className="border-b border-red-500/20" style={{ background: "rgba(239,68,68,0.08)" }}><div className="max-w-2xl mx-auto px-4 py-2"><p className="text-xs text-red-400">⚠️ {error}</p></div></div>}

      {(stats.some(s => s.emRisco) || urgentCount > 0) && (
        <div className="border-b border-amber-500/20" style={{ background: "rgba(245,158,11,0.08)" }}>
          <div className="max-w-2xl mx-auto px-4 py-2">
            <p className="text-xs text-amber-400 flex items-center gap-2">⚠️
              {stats.filter(s => s.emRisco).length > 0 && <span>{stats.filter(s => s.emRisco).length} disciplina(s) em risco</span>}
              {urgentCount > 0 && <span>· {urgentCount} atividade(s) em até 3 dias</span>}
            </p>
          </div>
        </div>
      )}

      {dataLoading ? (
        <div className="flex-1 flex items-center justify-center"><p className="text-slate-500 text-sm">Carregando suas atividades...</p></div>
      ) : (
        <main className="flex-1 max-w-2xl mx-auto w-full px-4 py-6 px-safe" style={{ paddingBottom: "calc(80px + var(--sab, 0px))" }}>
          {tab === "dashboard"   && <DashboardTab   items={items} stats={stats} meta={meta} />}
          {tab === "disciplines" && <DisciplineTab  items={items} stats={stats} meta={meta} onEditItem={handleEdit} onDeleteItem={handleDelete} onQuickGrade={handleQuickGrade} />}
          {tab === "plano"       && <PlanejamentoTab stats={stats} meta={meta} onChangeMeta={handleChangeMeta} />}
          {tab === "alerts"      && <AlertsTab      items={items} stats={stats} meta={meta} />}
          {tab === "calendar"    && <CalendarTab    items={items} />}
        </main>
      )}

      {/* Barra de navegação fixa estilo app — ícones (mobile) + rótulo (desktop) */}
      <nav className="fixed bottom-0 left-0 right-0 z-40 border-t border-white/10 nav-safe" style={{ background: "rgba(10,15,26,0.97)", backdropFilter: "blur(20px)" }} aria-label="Navegação principal">
        <div className="max-w-2xl mx-auto px-2 flex">
          {TABS.map(t => {
            const active = tab === t.id;
            return (
              <button key={t.id} onClick={() => setTab(t.id)} aria-label={t.label} aria-current={active ? "page" : undefined}
                className={`flex-1 py-2.5 flex flex-col sm:flex-row items-center justify-center gap-0.5 sm:gap-2 transition relative ${active ? "text-indigo-400" : "text-slate-500 hover:text-slate-300"}`}>
                <span className="text-xl sm:text-base leading-none" style={active ? { transform: "scale(1.1)" } : {}}>{t.icon}</span>
                <span className={`text-[10px] sm:text-xs font-semibold ${active ? "" : "sm:opacity-80"} hidden sm:inline`}>{t.label}</span>
                <span className="text-[10px] font-medium sm:hidden">{active ? t.label : ""}</span>
                {active && <span className="absolute top-0 left-1/2 -translate-x-1/2 w-8 h-0.5 rounded-full" style={{ background: "linear-gradient(90deg,#6366f1,#8b5cf6)" }} />}
              </button>
            );
          })}
        </div>
      </nav>

      <Modal open={modal === "novadisc"} onClose={() => setModal(null)} title="📚 Nova Disciplina">
        <NovaDisciplinaForm onSaveDisciplina={handleSaveDisciplina} onClose={() => setModal(null)} />
      </Modal>
      <Modal open={modal === "add" || modal === "edit"} onClose={() => { setModal(null); setEditingItem(null); }} title={modal === "edit" ? "Editar Atividade" : (novoTipo === "evento" ? "Novo Evento" : "Nova Avaliação")}>
        <ItemForm item={editingItem} onSave={handleSave} onClose={() => { setModal(null); setEditingItem(null); }} disciplines={disciplines} disciplinasFull={disciplinas} tipoInicial={novoTipo} />
      </Modal>
      <Modal open={modal === "quickgrade" && !!editingItem} onClose={() => { setModal(null); setEditingItem(null); }} title="🎯 Lançar Nota">
        {editingItem && <QuickGradePanel item={editingItem} stats={stats} meta={meta} onSave={handleSave} onClose={() => { setModal(null); setEditingItem(null); }} />}
      </Modal>
      <Modal open={modal === "import"} onClose={() => setModal(null)} title="Importar Planilha">
        <ImportPanel onImport={importAtividades} onClose={() => setModal(null)} />
      </Modal>
      <Modal open={modal === "export"} onClose={() => setModal(null)} title="📥 Exportar Planilha">
        <ExportPanel items={atividades} onClose={() => setModal(null)} />
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
