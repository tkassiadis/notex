// ============================================================
// src/lib/calculos.ts
// Funções matemáticas do sistema.
// CÓDIGO PRESERVADO INTEGRALMENTE do arquivo original.
// Única adição: tipagem TypeScript e parâmetro `meta` em getDisciplineStats.
// ============================================================

import type { Atividade, AtividadeEnriquecida, DisciplinaStats } from "../types";

const DISC_COLORS = [
  "#6366f1", "#8b5cf6", "#ec4899", "#f43f5e",
  "#f97316", "#eab308", "#10b981", "#06b6d4",
];

// ------------------------------------------------------------
// calcDaysRemaining — PRESERVADO DO ORIGINAL
// Diferença em dias entre hoje e a data da atividade
// ------------------------------------------------------------
export function calcDaysRemaining(dateStr: string | null | undefined): number | null {
  if (!dateStr) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const d = new Date(dateStr);
  d.setHours(0, 0, 0, 0);
  return Math.round((d.getTime() - today.getTime()) / 86400000);
}

// ------------------------------------------------------------
// calcNota — PRESERVADO DO ORIGINAL
// Converte pontuação para escala 0–10
// ------------------------------------------------------------
export function calcNota(
  pontuacao: number | null | undefined,
  pontuacaoMaxima: number | null | undefined
): number | null {
  if (pontuacao == null || !pontuacaoMaxima) return null;
  return (pontuacao / pontuacaoMaxima) * 10;
}

// ------------------------------------------------------------
// enrichItem — PRESERVADO DO ORIGINAL
// Adiciona daysRemaining calculado ao item
// ------------------------------------------------------------
export function enrichItem(item: Atividade): AtividadeEnriquecida {
  return { ...item, daysRemaining: calcDaysRemaining(item.data) };
}

// ------------------------------------------------------------
// getDisciplineStats — PRESERVADO DO ORIGINAL
// Única alteração: recebe `meta` como parâmetro (antes hardcoded como 7)
// Isso permite que o usuário configure sua própria meta de aprovação
// ------------------------------------------------------------
export function getDisciplineStats(
  items: AtividadeEnriquecida[],
  meta: number = 7
): DisciplinaStats[] {
  const byDisc: Record<string, AtividadeEnriquecida[]> = {};
  items.forEach((it) => {
    if (!byDisc[it.disciplina]) byDisc[it.disciplina] = [];
    byDisc[it.disciplina].push(it);
  });

  return Object.entries(byDisc).map(([disc, rows], idx) => {
    let weightedSum = 0;
    let weightUsed = 0;

    rows.forEach((r) => {
      const nota = calcNota(r.pontuacao, r.pontuacaoMaxima);
      if (nota != null) {
        weightedSum += nota * r.pesoAvaliacao * r.pesoInstrumento;
        weightUsed += r.pesoAvaliacao * r.pesoInstrumento;
      }
    });

    const mediaAtual = weightUsed > 0 ? weightedSum / weightUsed : null;
    const pesoConcluido = rows.reduce(
      (a, r) => (r.pontuacao != null ? a + r.pesoAvaliacao * r.pesoInstrumento : a),
      0
    );
    const pesoRestante = rows.reduce(
      (a, r) => (r.pontuacao == null ? a + r.pesoAvaliacao * r.pesoInstrumento : a),
      0
    );
    const pesoTotal = pesoConcluido + pesoRestante;

    // `meta` substituiu o hardcoded 7
    const notaNecessaria =
      pesoRestante > 0.001
        ? (meta * pesoTotal - weightedSum) / pesoRestante
        : null;

    return {
      disciplina: disc,
      items: rows,
      mediaAtual: mediaAtual != null ? Math.round(mediaAtual * 100) / 100 : null,
      pesoConcluido: Math.round(pesoConcluido * 1000) / 10,
      pesoRestante: Math.round(pesoRestante * 1000) / 10,
      notaNecessaria:
        notaNecessaria != null
          ? Math.max(0, Math.round(notaNecessaria * 100) / 100)
          : null,
      statusCounts: {
        "Não iniciado": rows.filter((r) => r.status === "Não iniciado").length,
        "Em andamento": rows.filter((r) =>
          ["Estudo inicial", "Estudo médio", "Estudo avançado"].includes(r.status)
        ).length,
        Finalizado: rows.filter((r) => r.status === "Finalizado").length,
      },
      proximas: rows.filter(
        (r) => r.daysRemaining != null && r.daysRemaining >= 0 && r.daysRemaining <= 14
      ),
      aguardandoCorrecao: rows.filter(
        (r) => r.daysRemaining != null && r.daysRemaining < 0 && r.pontuacao == null
      ),
      color: DISC_COLORS[idx % DISC_COLORS.length],
      emRisco: mediaAtual != null && mediaAtual < meta - 1, // risco = 1 ponto abaixo da meta
    };
  });
}
