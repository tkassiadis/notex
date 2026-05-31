// ============================================================
// src/lib/xlsx.ts
// Funções de importação e exportação de planilhas.
// CÓDIGO PRESERVADO INTEGRALMENTE do arquivo original.
// Única adição: tipagem TypeScript.
// ============================================================

import * as XLSX from "xlsx";
import type { Atividade } from "../types";
import { calcNota, calcDaysRemaining } from "./calculos";

// ------------------------------------------------------------
// exportToExcel — PRESERVADO DO ORIGINAL
// ------------------------------------------------------------
export async function exportToExcel(items: Atividade[]): Promise<void> {
  try {
    const headers = [
      "#", "Avaliação", "Instrumento", "Disciplina", "Subdivisão Disciplinar",
      "Status", "Data", "Restam (Dias)", "Peso Avaliação", "Peso Instrumento",
      "Pontuação Máxima", "Pontuação", "Nota (de 10)", "Observações",
    ];

    const rows = items.map((it, i) => {
      const nota = calcNota(it.pontuacao, it.pontuacaoMaxima);
      const days = calcDaysRemaining(it.data);
      return [
        i + 1,
        it.avaliacao,
        it.instrumento,
        it.disciplina,
        it.subdivisao || "",
        it.status,
        it.data || "",
        days != null ? (days < 0 ? "Já passou" : days) : "",
        it.pesoAvaliacao,
        it.pesoInstrumento,
        it.pontuacaoMaxima ?? "",
        it.pontuacao ?? "",
        nota != null ? Math.round(nota * 100) / 100 : "",
        it.observacoes || "",
      ];
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "controle-semestre-atualizado.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (e: any) {
    alert("Erro ao exportar: " + e.message);
  }
}

// ------------------------------------------------------------
// parseXlsxFile — PRESERVADO DO ORIGINAL (era inline no ImportPanel)
// Extrai e normaliza as atividades de um arquivo .xlsx
// ------------------------------------------------------------
export async function parseXlsxFile(file: File): Promise<Atividade[]> {
  const statusMap: Record<string, string> = {
    "nao iniciado": "Não iniciado",
    "não iniciado": "Não iniciado",
    "estudo inicial": "Estudo inicial",
    "estudo medio": "Estudo médio",
    "estudo médio": "Estudo médio",
    "estudo avancado": "Estudo avançado",
    "estudo avançado": "Estudo avançado",
    finalizado: "Finalizado",
  };

  const wb = XLSX.read(await file.arrayBuffer());
  const ws = wb.Sheets["Sheet1"] || wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

  const parsed = raw
    .slice(1)
    .filter((r) => r[0] != null && r[0] !== "")
    .map((r, i) => {
      let dateStr = "";
      const dv = r[6];
      if (dv instanceof Date) dateStr = dv.toISOString().split("T")[0];
      else if (typeof dv === "number")
        dateStr = new Date((dv - 25569) * 86400000).toISOString().split("T")[0];
      else if (typeof dv === "string" && dv.length > 0)
        dateStr = dv.split("T")[0];

      return {
        id: `import-${Date.now()}-${i}`, // temporário, substituído pelo UUID do banco
        avaliacao: String(r[1] || "").trim(),
        instrumento: String(r[2] || "").trim(),
        disciplina: String(r[3] || "").trim(),
        subdivisao: String(r[4] || "").trim(),
        status:
          statusMap[String(r[5] || "").toLowerCase().trim()] || "Não iniciado",
        data: dateStr,
        pesoAvaliacao: parseFloat(r[8]) || 0,
        pesoInstrumento: parseFloat(r[9]) || 0,
        pontuacaoMaxima:
          r[10] != null && r[10] !== "" ? parseFloat(r[10]) : null,
        pontuacao:
          r[11] != null && r[11] !== "" ? parseFloat(r[11]) : null,
        observacoes: String(r[13] || "").trim(),
      } as Atividade;
    })
    .filter((r) => r.disciplina && r.instrumento);

  return parsed;
}
