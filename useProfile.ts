// ============================================================
// src/hooks/useProfile.ts
// Gerencia o perfil do usuário, incluindo meta_aprovacao.
// ============================================================

import { useState, useEffect, useCallback } from "react";
import type { User } from "@supabase/supabase-js";
import { supabase } from "../lib/supabase";
import type { Profile, ProfileRow } from "../types";

function rowToProfile(row: ProfileRow): Profile {
  return {
    id: row.id,
    email: row.email,
    nome: row.nome,
    metaAprovacao: row.meta_aprovacao,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

interface UseProfileReturn {
  profile: Profile | null;
  loading: boolean;
  updateMeta: (meta: number) => Promise<void>;
}

export function useProfile(user: User | null): UseProfileReturn {
  const [profile, setProfile] = useState<Profile | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!user) {
      setProfile(null);
      setLoading(false);
      return;
    }

    supabase
      .from("profiles")
      .select("*")
      .eq("id", user.id)
      .single()
      .then(({ data, error }) => {
        if (!error && data) {
          setProfile(rowToProfile(data as ProfileRow));
        }
        setLoading(false);
      });
  }, [user?.id]);

  const updateMeta = useCallback(
    async (meta: number): Promise<void> => {
      if (!user) return;
      const { error } = await supabase
        .from("profiles")
        .update({ meta_aprovacao: meta })
        .eq("id", user.id);
      if (!error) {
        setProfile((prev) => (prev ? { ...prev, metaAprovacao: meta } : prev));
      }
    },
    [user]
  );

  return { profile, loading, updateMeta };
}
