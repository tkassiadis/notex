// ============================================================
// src/types/index.ts
// ============================================================

export interface Atividade {
  id: string;
  tipo: TipoItem;                // "avaliacao" (afeta médias) | "evento" (só registro)
  avaliacao: string;
  instrumento: string;
  disciplina: string;
  disciplinaId: string | null;   // FK para a tabela disciplinas
  parte: ParteDisciplina;        // "unica" | "teorica" | "pratica" (para mistas)
  subdivisao: string;
  status: StatusOption;
  data: string;
  pesoAvaliacao: number;
  pesoInstrumento: number;
  pontuacaoMaxima: number | null;
  pontuacao: number | null;
  observacoes: string;
}

export type ParteDisciplina = "unica" | "teorica" | "pratica";

export type TipoDisciplina = "Teórica" | "Prática" | "Mista";

export interface Disciplina {
  id: string;
  nome: string;
  tipo: TipoDisciplina;
  observacoes: string;
  pesoTeorica: number;           // % da parte teórica no total (mistas)
  pesoPratica: number;           // % da parte prática no total (mistas)
  subdivisoes: string[];         // lista de nomes de subdivisões
}

export interface DisciplinaRow {
  id: string;
  user_id: string;
  nome: string;
  tipo: string;
  observacoes: string;
  peso_teorica: number;
  peso_pratica: number;
  subdivisoes: string[] | null;
  created_at: string;
  updated_at: string;
}

export interface AtividadeEnriquecida extends Atividade {
  daysRemaining: number | null;
}

export type TipoItem = "avaliacao" | "evento";

export type StatusOption =
  | "Não iniciado"
  | "Estudo inicial"
  | "Estudo médio"
  | "Estudo avançado"
  | "Finalizado";

export type AvaliacaoOption = "AP1" | "AP2" | "AS" | "AF" | "Evento";

export interface DisciplinaStats {
  disciplina: string;
  disciplinaId: string | null;         // referência à entidade Disciplina
  tipoDisciplina: TipoDisciplina;      // Teórica | Prática | Mista
  observacoes: string;                 // observações da disciplina (persistentes)
  items: AtividadeEnriquecida[];
  mediaAtual: number | null;
  pesoConcluido: number;
  pesoRestante: number;
  notaNecessaria: number | null;
  notaMaxima: number | null;           // nota máxima alcançável
  aprovacaoGarantida: boolean;         // meta já garantida matematicamente
  aprovacaoImpossivel: boolean;        // impossível atingir a meta
  // Para disciplinas MISTAS: média de cada parte (null se não houver notas)
  mediaTeorica: number | null;
  mediaPratica: number | null;
  pesoParteTeorica: number;            // % da parte teórica no total
  pesoPartePratica: number;            // % da parte prática no total
  // Pontos na média final (escala 0–10), para o gráfico do Dashboard
  pontosConquistados: number;          // já garantidos na média
  pontosAConquistar: number;           // máximo ainda possível (recuperável)
  pontosPerdidos: number;              // perdido em avaliações feitas + impossível
  statusCounts: {
    "Não iniciado": number;
    "Em andamento": number;
    "Finalizado": number;
  };
  proximas: AtividadeEnriquecida[];
  aguardandoCorrecao: AtividadeEnriquecida[];
  color: string;
  emRisco: boolean;
}

export interface Profile {
  id: string;
  email: string;
  nome: string | null;
  metaAprovacao: number;
  createdAt: string;
  updatedAt: string;
}

export interface AtividadeRow {
  id: string;
  user_id: string;
  tipo: string;
  avaliacao: string;
  instrumento: string;
  disciplina: string;
  disciplina_id: string | null;
  parte: string;
  subdivisao: string;
  status: string;
  data: string | null;
  peso_avaliacao: number;
  peso_instrumento: number;
  pontuacao_maxima: number | null;
  pontuacao: number | null;
  observacoes: string;
  created_at: string;
  updated_at: string;
}

export interface ProfileRow {
  id: string;
  email: string;
  nome: string | null;
  meta_aprovacao: number;
  created_at: string;
  updated_at: string;
}
