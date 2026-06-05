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

  // Agrupa por nome de disciplina. Itens sem disciplina (ex: eventos gerais) → "Eventos Gerais".
  const byDisc: Record<string, AtividadeEnriquecida[]> = {};
  items.forEach((it) => {
    const chave = it.disciplina && it.disciplina.trim() ? it.disciplina : "Eventos Gerais";
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
    const dMeta = metaByName[disc];
    const ehMista = dMeta?.tipo === "Mista";

    // ── Métricas de um conjunto de avaliações (uma parte ou a disciplina toda) ──
    // Modelo de DOIS NÍVEIS:
    //  • Instrumentos são agrupados pela "avaliação" (AP1, AP2...).
    //  • Dentro de cada avaliação, cada instrumento contribui com nota × pesoInstrumento.
    //    Instrumentos faltando contam como ZERO (não re-normaliza).
    //  • Cada avaliação contribui para a disciplina com seu pesoAvaliacao.
    // Retorna: ws (soma ponderada das notas já obtidas), wUsed (peso de avaliação com
    // alguma nota), wConcl/wRest em termos de peso de avaliação × fração concluída.
    function metricasDe(avs: AtividadeEnriquecida[]) {
      // Agrupa por nome de avaliação
      const grupos: Record<string, AtividadeEnriquecida[]> = {};
      avs.forEach((r) => {
        const chave = r.avaliacao || "—";
        if (!grupos[chave]) grupos[chave] = [];
        grupos[chave].push(r);
      });

      let ws = 0, wUsed = 0, wConcl = 0, wRest = 0;
      Object.values(grupos).forEach((insts) => {
        // Peso da avaliação no total da disciplina (todos os instrumentos do grupo
        // compartilham o mesmo peso de avaliação; usamos o primeiro informado).
        const pesoAval = insts[0].pesoAvaliacao;
        if (pesoAval <= 0) return;

        // Nota da avaliação = Σ (nota_instrumento × pesoInstrumento).
        // Instrumento sem nota contribui 0. Não re-normaliza (faltando = zero).
        let notaAval = 0;       // 0..10 (assumindo instrumentos somando 100%)
        let fracComNota = 0;    // fração da avaliação já avaliada (Σ pesoInstrumento com nota)
        insts.forEach((r) => {
          const nota = calcNota(r.pontuacao, r.pontuacaoMaxima);
          if (nota != null) { notaAval += nota * r.pesoInstrumento; fracComNota += r.pesoInstrumento; }
        });

        // A parte concluída desta avaliação (em peso de avaliação) é proporcional
        // à fração de instrumentos já avaliados.
        const concl = pesoAval * fracComNota;
        const rest  = pesoAval * (1 - fracComNota);
        wConcl += concl;
        wRest  += Math.max(0, rest);

        if (fracComNota > 0) {
          // contribui com a nota já obtida (sobre a fração avaliada)
          ws += notaAval * pesoAval;       // notaAval já está ponderado pelos pesoInstrumento
          wUsed += pesoAval * fracComNota;
        }
      });

      const media = wUsed > 0 ? ws / wUsed : null;
      return { ws, wUsed, wConcl, wRest, media };
    }

    let mediaAtual: number | null;
    let mediaTeorica: number | null = null;
    let mediaPratica: number | null = null;
    let pesoConcluido: number;   // 0..1 (fração do total da disciplina concluída)
    let pesoRestante: number;
    let notaMaxima: number | null;
    let notaNecessariaRaw: number | null;
    const pesoParteTeorica = dMeta?.pesoTeorica ?? 100;
    const pesoPartePratica = dMeta?.pesoPratica ?? 0;

    if (ehMista) {
      // Cada parte é normalizada para fechar 100% dentro de si.
      const fT = pesoParteTeorica / 100;
      const fP = pesoPartePratica / 100;
      const teo = metricasDe(avaliacoes.filter((r) => r.parte !== "pratica")); // teorica + unica
      const pra = metricasDe(avaliacoes.filter((r) => r.parte === "pratica"));

      mediaTeorica = teo.media != null ? Math.round(teo.media * 100) / 100 : null;
      mediaPratica = pra.media != null ? Math.round(pra.media * 100) / 100 : null;

      // Média geral combinada: só considera as partes que já têm nota,
      // re-normalizando os pesos das partes presentes (para não penalizar parte ainda sem nota).
      let somaComb = 0, pesoComb = 0;
      if (teo.media != null) { somaComb += teo.media * fT; pesoComb += fT; }
      if (pra.media != null) { somaComb += pra.media * fP; pesoComb += fP; }
      mediaAtual = pesoComb > 0 ? somaComb / pesoComb : null;

      // Progresso e projeções ponderados pelo peso de cada parte (corrige o bug dos 200%).
      const progT = (teo.wConcl + teo.wRest) > 0 ? teo.wConcl / (teo.wConcl + teo.wRest) : 0;
      const progP = (pra.wConcl + pra.wRest) > 0 ? pra.wConcl / (pra.wConcl + pra.wRest) : 0;
      pesoConcluido = progT * fT + progP * fP;       // 0..1
      pesoRestante  = (1 - progT) * fT + (1 - progP) * fP;

      // Nota máxima: parte sem nota pode chegar a 10.
      const maxT = teo.media != null
        ? (teo.wUsed + teo.wRest > 0 ? (teo.ws + 10 * teo.wRest) / (teo.wUsed + teo.wRest) : teo.media)
        : 10;
      const maxP = pra.media != null
        ? (pra.wUsed + pra.wRest > 0 ? (pra.ws + 10 * pra.wRest) / (pra.wUsed + pra.wRest) : pra.media)
        : 10;
      notaMaxima = maxT * fT + maxP * fP;

      // Nota necessária nas avaliações restantes para atingir a meta no total.
      const restoFrac = pesoRestante;
      const jaGarantido = (mediaAtual != null ? mediaAtual : 0) * pesoConcluido;
      notaNecessariaRaw = restoFrac > 0.001 ? (meta - jaGarantido) / restoFrac : null;
    } else {
      // Disciplina simples (Teórica ou Prática): comportamento original.
      const m = metricasDe(avaliacoes);
      mediaAtual = m.media != null ? Math.round(m.media * 100) / 100 : null;
      const pesoTotal = m.wConcl + m.wRest;
      pesoConcluido = pesoTotal > 0 ? m.wConcl / pesoTotal : 0;
      pesoRestante  = pesoTotal > 0 ? m.wRest  / pesoTotal : 0;
      notaNecessariaRaw = m.wRest > 0.001 ? (meta * pesoTotal - m.ws) / m.wRest : null;
      notaMaxima = pesoTotal > 0.001 ? (m.ws + 10 * m.wRest) / pesoTotal : mediaAtual;
    }

    const aprovacaoGarantida  = notaNecessariaRaw !== null && notaNecessariaRaw <= 0;
    const aprovacaoImpossivel = notaMaxima !== null && notaMaxima < meta;

    // ── Pontos na média final (escala 0–10) ──
    // Conquistados: o que já está garantido na média final.
    // A conquistar: o máximo ainda obtível nas avaliações restantes (tirar 10).
    // Perdidos: o que já foi perdido nas avaliações feitas + o que ficou impossível.
    const pontosConquistados = (mediaAtual != null ? mediaAtual : 0) * pesoConcluido;
    const pontosAConquistar = pesoRestante * 10;
    const pontosPerdidos = Math.max(0, 10 - pontosConquistados - pontosAConquistar);
    const notaNecessaria = notaNecessariaRaw != null
      ? Math.round(notaNecessariaRaw * 100) / 100
      : null;

    return {
      disciplina: disc,
      disciplinaId: dMeta?.id ?? null,
      tipoDisciplina: dMeta?.tipo ?? "Teórica",
      observacoes: dMeta?.observacoes ?? "",
      items: rows,        // lista completa (avaliações + eventos) para UI/calendário/alertas
      mediaAtual:    mediaAtual != null ? Math.round(mediaAtual * 100) / 100 : null,
      pesoConcluido: Math.round(pesoConcluido * 1000) / 10,   // → porcentagem 0..100
      pesoRestante:  Math.round(pesoRestante  * 1000) / 10,
      notaNecessaria,
      notaMaxima:    notaMaxima != null ? Math.round(notaMaxima * 100) / 100 : null,
      aprovacaoGarantida,
      aprovacaoImpossivel,
      mediaTeorica,
      mediaPratica,
      pesoParteTeorica,
      pesoPartePratica,
      pontosConquistados: Math.round(pontosConquistados * 100) / 100,
      pontosAConquistar: Math.round(pontosAConquistar * 100) / 100,
      pontosPerdidos: Math.round(pontosPerdidos * 100) / 100,
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
