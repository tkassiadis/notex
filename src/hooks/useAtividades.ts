// ============================================================
// src/hooks/useAtividades.ts
// SUBSTITUI o useState(rawItems) + useEffect do App original.
// Expõe a MESMA interface que o estado anterior:
//   rawItems, setRawItems equivalentes via addAtividade/updateAtividade/deleteAtividade
//
// Internamente: sincroniza com Supabase + escuta realtime.
// ============================================================

import { useState, useEffect, useCallback } from "react";
import type { User } from "@supabase/supabase-js";
import { supabase } from "../lib/supabase";
import type { Atividade, AtividadeRow } from "../types";

// ------------------------------------------------------------
// Conversão snake_case (banco) ↔ camelCase (frontend)
// Preserva todos os nomes de campo do original
// ------------------------------------------------------------
function rowToAtividade(row: AtividadeRow): Atividade {
  return {
    id: row.id,
    avaliacao: row.avaliacao,
    instrumento: row.instrumento,
    disciplina: row.disciplina,
    subdivisao: row.subdivisao,
    status: row.status as Atividade["status"],
    data: row.data ?? "",
    pesoAvaliacao: row.peso_avaliacao,
    pesoInstrumento: row.peso_instrumento,
    pontuacaoMaxima: row.pontuacao_maxima,
    pontuacao: row.pontuacao,
    observacoes: row.observacoes,
  };
}

function atividadeToInsert(
  item: Omit<Atividade, "id">,
  userId: string
): Omit<AtividadeRow, "id" | "created_at" | "updated_at"> {
  return {
    user_id: userId,
    avaliacao: item.avaliacao,
    instrumento: item.instrumento,
    disciplina: item.disciplina,
    subdivisao: item.subdivisao,
    status: item.status,
    data: item.data || null,
    peso_avaliacao: item.pesoAvaliacao,
    peso_instrumento: item.pesoInstrumento,
    pontuacao_maxima: item.pontuacaoMaxima,
    pontuacao: item.pontuacao,
    observacoes: item.observacoes,
  };
}

function atividadeToUpdate(
  item: Atividade
): Partial<Omit<AtividadeRow, "id" | "user_id" | "created_at" | "updated_at">> {
  return {
    avaliacao: item.avaliacao,
    instrumento: item.instrumento,
    disciplina: item.disciplina,
    subdivisao: item.subdivisao,
    status: item.status,
    data: item.data || null,
    peso_avaliacao: item.pesoAvaliacao,
    peso_instrumento: item.pesoInstrumento,
    pontuacao_maxima: item.pontuacaoMaxima,
    pontuacao: item.pontuacao,
    observacoes: item.observacoes,
  };
}

// ------------------------------------------------------------
// Hook principal
// ------------------------------------------------------------
interface UseAtividadesReturn {
  atividades: Atividade[];
  loading: boolean;
  error: string | null;
  addAtividade: (item: Omit<Atividade, "id">) => Promise<void>;
  updateAtividade: (item: Atividade) => Promise<void>;
  deleteAtividade: (id: string) => Promise<void>;
  importAtividades: (items: Omit<Atividade, "id">[]) => Promise<void>;
}

export function useAtividades(user: User | null): UseAtividadesReturn {
  const [atividades, setAtividades] = useState<Atividade[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // ----------------------------------------------------------
  // Carregamento inicial: busca todas as atividades do usuário
  // ----------------------------------------------------------
  useEffect(() => {
    if (!user) {
      setAtividades([]);
      setLoading(false);
      return;
    }

    setLoading(true);
    setError(null);

    supabase
      .from("atividades")
      .select("*")
      .eq("user_id", user.id)
      .order("created_at", { ascending: true })
      .then(({ data, error }) => {
        if (error) {
          setError("Erro ao carregar atividades: " + error.message);
        } else {
          setAtividades((data as AtividadeRow[]).map(rowToAtividade));
        }
        setLoading(false);
      });
  }, [user?.id]);

  // ----------------------------------------------------------
  // Realtime: escuta INSERT, UPDATE e DELETE no banco
  // Atualiza o estado local automaticamente (sincronização entre dispositivos)
  // ----------------------------------------------------------
  useEffect(() => {
    if (!user) return;

    const channel = supabase
      .channel(`atividades:${user.id}`)
      .on(
        "postgres_changes",
        {
          event: "INSERT",
          schema: "public",
          table: "atividades",
          filter: `user_id=eq.${user.id}`,
        },
        (payload) => {
          const nova = rowToAtividade(payload.new as AtividadeRow);
          setAtividades((prev) => {
            // Evita duplicatas caso o insert tenha vindo do próprio dispositivo
            if (prev.find((a) => a.id === nova.id)) return prev;
            return [...prev, nova];
          });
        }
      )
      .on(
        "postgres_changes",
        {
          event: "UPDATE",
          schema: "public",
          table: "atividades",
          filter: `user_id=eq.${user.id}`,
        },
        (payload) => {
          const atualizada = rowToAtividade(payload.new as AtividadeRow);
          setAtividades((prev) =>
            prev.map((a) => (a.id === atualizada.id ? atualizada : a))
          );
        }
      )
      .on(
        "postgres_changes",
        {
          event: "DELETE",
          schema: "public",
          table: "atividades",
          filter: `user_id=eq.${user.id}`,
        },
        (payload) => {
          setAtividades((prev) =>
            prev.filter((a) => a.id !== (payload.old as AtividadeRow).id)
          );
        }
      )
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [user?.id]);

  // ----------------------------------------------------------
  // addAtividade — equivalente a setRawItems(prev => [...prev, novoItem])
  // ----------------------------------------------------------
  const addAtividade = useCallback(
    async (item: Omit<Atividade, "id">): Promise<void> => {
      if (!user) return;
      const { data, error } = await supabase
        .from("atividades")
        .insert(atividadeToInsert(item, user.id))
        .select()
        .single();
      if (error) {
        setError("Erro ao salvar atividade: " + error.message);
        throw error;
      }
      // O realtime já vai capturar o INSERT, mas atualizamos localmente
      // para resposta imediata na UI (sem esperar o evento de rede)
      const nova = rowToAtividade(data as AtividadeRow);
      setAtividades((prev) =>
        prev.find((a) => a.id === nova.id) ? prev : [...prev, nova]
      );
    },
    [user]
  );

  // ----------------------------------------------------------
  // updateAtividade — equivalente a setRawItems(prev => prev.map(...))
  // ----------------------------------------------------------
  const updateAtividade = useCallback(
    async (item: Atividade): Promise<void> => {
      if (!user) return;
      const { error } = await supabase
        .from("atividades")
        .update(atividadeToUpdate(item))
        .eq("id", item.id)
        .eq("user_id", user.id);
      if (error) {
        setError("Erro ao atualizar atividade: " + error.message);
        throw error;
      }
      // Atualização otimista local
      setAtividades((prev) =>
        prev.map((a) => (a.id === item.id ? item : a))
      );
    },
    [user]
  );

  // ----------------------------------------------------------
  // deleteAtividade — equivalente a setRawItems(prev => prev.filter(...))
  // ----------------------------------------------------------
  const deleteAtividade = useCallback(
    async (id: string): Promise<void> => {
      if (!user) return;
      const { error } = await supabase
        .from("atividades")
        .delete()
        .eq("id", id)
        .eq("user_id", user.id);
      if (error) {
        setError("Erro ao excluir atividade: " + error.message);
        return;
      }
      // Remoção otimista local
      setAtividades((prev) => prev.filter((a) => a.id !== id));
    },
    [user]
  );

  // ----------------------------------------------------------
  // importAtividades — equivalente ao onImport do ImportPanel
  // Deleta tudo do usuário e insere o novo lote
  // ----------------------------------------------------------
  const importAtividades = useCallback(
    async (items: Omit<Atividade, "id">[]): Promise<void> => {
      if (!user) return;

      // Deleta todas as atividades existentes do usuário
      const { error: deleteError } = await supabase
        .from("atividades")
        .delete()
        .eq("user_id", user.id);
      if (deleteError) {
        setError("Erro ao limpar dados para importação: " + deleteError.message);
        return;
      }

      // Insere todas as novas atividades em lote
      const inserts = items.map((item) => atividadeToInsert(item, user.id));
      const { data, error: insertError } = await supabase
        .from("atividades")
        .insert(inserts)
        .select();
      if (insertError) {
        setError("Erro ao importar atividades: " + insertError.message);
        return;
      }

      setAtividades((data as AtividadeRow[]).map(rowToAtividade));
    },
    [user]
  );

  return {
    atividades,
    loading,
    error,
    addAtividade,
    updateAtividade,
    deleteAtividade,
    importAtividades,
  };
}
