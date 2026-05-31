// ============================================================
// src/lib/supabase.ts
// Client Supabase singleton.
// As variáveis de ambiente são lidas do .env (Vite).
// ============================================================

import { createClient } from "@supabase/supabase-js";
import type { AtividadeRow, ProfileRow } from "../types";

// Tipos do banco para o client tipado do Supabase
export type Database = {
  public: {
    Tables: {
      atividades: {
        Row: AtividadeRow;
        Insert: Omit<AtividadeRow, "id" | "created_at" | "updated_at">;
        Update: Partial<Omit<AtividadeRow, "id" | "user_id" | "created_at" | "updated_at">>;
      };
      profiles: {
        Row: ProfileRow;
        Insert: Omit<ProfileRow, "created_at" | "updated_at">;
        Update: Partial<Pick<ProfileRow, "nome" | "meta_aprovacao">>;
      };
    };
  };
};

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL as string;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string;

if (!supabaseUrl || !supabaseAnonKey) {
  throw new Error(
    "Variáveis de ambiente VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY são obrigatórias."
  );
}

export const supabase = createClient<Database>(supabaseUrl, supabaseAnonKey, {
  auth: {
    // Persiste a sessão no localStorage do navegador
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
  realtime: {
    params: {
      eventsPerSecond: 10,
    },
  },
});
