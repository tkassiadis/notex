// ============================================================
// src/lib/calculos.ts
// ============================================================

import type { Atividade, AtividadeEnriquecida, DisciplinaStats, Disciplina } from "../types";

const DISC_COLORS = [
  "#6366f1", "#8b5cf6", "#ec4899", "#f43f5e",
  "#f97316", "#eab308", "#10b981", "#06b6d4",
];

export function calcDaysRemaining(dateStr: string | null | undefined): number | null {
  if (!dateStr) return null;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const d = new Date(dateStr + "T00:00:00"); d.setHours(0, 0, 0, 0);
  return Math.round((d.getTime() - today.getTime()) / 86400000);
}

export function calcNota(
  pontuacao: number | null | undefined,
  pontuacaoMaxima: number | null | undefined
): number | null {
  if (pontuacao == null || !pontuacaoMaxima) return null;
  return (pontuacao / pontuacaoMaxima) * 10;
}

export function enrichItem(item: Atividade): AtividadeEnriquecida {
  return { ...item, daysRemaining: calcDaysRemaining(item.data) };
}

export function getDisciplineStats(
  items: AtividadeEnriquecida[],
  meta: number = 7,
  disciplinas: Disciplina[] = []
): DisciplinaStats[] {
  // Mapa nome → metadados da disciplina (tipo, observações, id)
  const metaByName: Record<string, Disciplina> = {};
  disciplinas.forEach((d) => { metaByName[d.nome] = d; });

  // Agrupa por nome de disciplina. Itens sem disciplina (ex: eventos gerais) → "Geral".
  const byDisc: Record<string, AtividadeEnriquecida[]> = {};
  items.forEach((it) => {
    const chave = it.disciplina && it.disciplina.trim() ? it.disciplina : "Geral";
    if (!byDisc[chave]) byDisc[chave] = [];
    byDisc[chave].push(it);
  });
  // Garante que disciplinas cadastradas mas ainda sem atividades apareçam
  disciplinas.forEach((d) => {
    if (!byDisc[d.nome]) byDisc[d.nome] = [];
  });

  return Object.entries(byDisc).map(([disc, rows], idx) => {
    // EVENTOS não entram em nenhum cálculo acadêmico — apenas avaliações.
    const avaliacoes = rows.filter((r) => r.tipo !== "evento");

    let weightedSum = 0, weightUsed = 0;

    avaliacoes.forEach((r) => {
      const nota = calcNota(r.pontuacao, r.pontuacaoMaxima);
      if (nota != null) {
        weightedSum += nota * r.pesoAvaliacao * r.pesoInstrumento;
        weightUsed  += r.pesoAvaliacao * r.pesoInstrumento;
      }
    });

    const mediaAtual    = weightUsed > 0 ? weightedSum / weightUsed : null;
    const pesoConcluido = avaliacoes.reduce((a, r) => r.pontuacao != null ? a + r.pesoAvaliacao * r.pesoInstrumento : a, 0);
    const pesoRestante  = avaliacoes.reduce((a, r) => r.pontuacao == null ? a + r.pesoAvaliacao * r.pesoInstrumento : a, 0);
    const pesoTotal     = pesoConcluido + pesoRestante;

    const notaNecessariaRaw = pesoRestante > 0.001
      ? (meta * pesoTotal - weightedSum) / pesoRestante
      : null;

    const notaMaxima = pesoTotal > 0.001
      ? (weightedSum + 10 * pesoRestante) / pesoTotal
      : mediaAtual;

    const aprovacaoGarantida  = notaNecessariaRaw !== null && notaNecessariaRaw <= 0;
    const aprovacaoImpossivel = notaMaxima !== null && notaMaxima < meta;

    // notaNecessaria exposta sem Math.max para preservar sinalizacao de impossibilidade
    const notaNecessaria = notaNecessariaRaw != null
      ? Math.round(notaNecessariaRaw * 100) / 100
      : null;

    return {
      disciplina: disc,
      disciplinaId: metaByName[disc]?.id ?? null,
      tipoDisciplina: metaByName[disc]?.tipo ?? "Teórica",
      observacoes: metaByName[disc]?.observacoes ?? "",
      items: rows,        // lista completa (avaliações + eventos) para UI/calendário/alertas
      mediaAtual:    mediaAtual != null ? Math.round(mediaAtual * 100) / 100 : null,
      pesoConcluido: Math.round(pesoConcluido * 1000) / 10,
      pesoRestante:  Math.round(pesoRestante  * 1000) / 10,
      notaNecessaria,
      notaMaxima:    notaMaxima != null ? Math.round(notaMaxima * 100) / 100 : null,
      aprovacaoGarantida,
      aprovacaoImpossivel,
      statusCounts: {
        "Não iniciado": rows.filter((r) => r.status === "Não iniciado").length,
        "Em andamento": rows.filter((r) => ["Estudo inicial","Estudo médio","Estudo avançado"].includes(r.status)).length,
        "Finalizado":   rows.filter((r) => r.status === "Finalizado").length,
      },
      proximas:           rows.filter((r) => r.daysRemaining != null && r.daysRemaining >= 0 && r.daysRemaining <= 14),
      aguardandoCorrecao: avaliacoes.filter((r) => r.daysRemaining != null && r.daysRemaining < 0 && r.pontuacao == null),
      color:   DISC_COLORS[idx % DISC_COLORS.length],
      emRisco: mediaAtual != null && !aprovacaoGarantida &&
               (aprovacaoImpossivel || (notaNecessaria != null && notaNecessaria > 8)),
    };
  });
}
