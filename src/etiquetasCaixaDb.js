/** Persistência de etiquetas de caixa (contador sequencial) no Supabase. */

import { supabase } from "./supabase";
import {
  formatContador,
  parseContadorInicial,
  readStoredContador,
  writeStoredContador,
} from "./etiquetaZpl";

export function normalizeContadorScan(raw) {
  const digits = String(raw ?? "").replace(/\D/g, "");
  if (!digits) return "";
  // Mantém o valor bipado (com zeros à esquerda se o leitor enviar).
  return digits;
}

export function contadorKeyVariants(raw) {
  const digits = normalizeContadorScan(raw);
  if (!digits) return [];
  const n = Number(digits) || 0;
  const variants = new Set([digits, formatContador(n, Math.max(digits.length, 7))]);
  // Também tenta sem zeros à esquerda e com 7 dígitos.
  variants.add(String(n));
  variants.add(formatContador(n, 7));
  return [...variants];
}

/** Próximo contador_n: max(DB, localStorage) + 1. */
export async function peekNextContadorN() {
  const local = parseContadorInicial(readStoredContador());
  let dbN = 0;
  const { data, error } = await supabase
    .from("etiquetas_caixa")
    .select("contador_n")
    .order("contador_n", { ascending: false })
    .limit(1)
    .maybeSingle();
  if (error) {
    // Tabela ainda não criada no Supabase — cai no localStorage.
    if (isMissingTableError(error)) {
      return { n: local.n, width: local.width, fromDb: false, error };
    }
    return { n: local.n, width: local.width, fromDb: false, error };
  }
  dbN = Number(data?.contador_n) || 0;
  const n = Math.max(dbN + 1, local.n);
  const width = Math.max(local.width, 7);
  return { n, width, fromDb: true, error: null };
}

export function isMissingTableError(error) {
  const msg = String(error?.message || error?.details || "").toLowerCase();
  const code = String(error?.code || "");
  return (
    code === "42P01" ||
    code === "PGRST205" ||
    (msg.includes("etiquetas_caixa") &&
      (msg.includes("does not exist") || msg.includes("schema cache")))
  );
}

/**
 * Registra etiquetas impressas/baixadas. Idempotente por contador (upsert).
 * @param {Array<{contador:string, sku:string, ref:string, cor:string, tamanho:number|string, sigla?:string, prog?:string}>} labels
 */
export async function registrarEtiquetasImpressas(labels, origem = "print") {
  const rows = (labels || [])
    .map((l) => {
      const contador = String(l.contador || "").trim();
      const { n, width } = parseContadorInicial(contador);
      if (!contador || !l.sku) return null;
      return {
        contador: formatContador(n, width),
        contador_n: n,
        sku: String(l.sku).trim().toUpperCase(),
        ref: String(l.ref || "").trim(),
        cor: String(l.cor || "").trim(),
        tamanho: Number(l.tamanho) || 0,
        sigla: String(l.sigla || "").trim() || null,
        programacao: String(l.prog || l.programacao || "").trim() || null,
        origem: String(origem || "print"),
      };
    })
    .filter((r) => r && r.ref && r.cor && r.tamanho > 0);

  if (!rows.length) {
    return { ok: false, error: new Error("Nenhuma etiqueta válida para registrar."), inserted: 0 };
  }

  const { data, error } = await supabase
    .from("etiquetas_caixa")
    .upsert(rows, { onConflict: "contador", ignoreDuplicates: true })
    .select("contador");

  if (error) {
    return { ok: false, error, inserted: 0 };
  }
  return { ok: true, error: null, inserted: (data || []).length };
}

/** Reserva bip (anti-duplicidade). Marca bipado_em se ainda livre. */
export async function reservarBipPorContador(rawContador) {
  const variants = contadorKeyVariants(rawContador);
  if (!variants.length) {
    return { ok: false, code: "invalid", error: "Código sequencial inválido." };
  }

  let row = null;
  for (const key of variants) {
    const { data, error } = await supabase
      .from("etiquetas_caixa")
      .select("*")
      .eq("contador", key)
      .maybeSingle();
    if (error) {
      if (isMissingTableError(error)) {
        return {
          ok: false,
          code: "missing_table",
          error:
            "Tabela etiquetas_caixa não existe no Supabase. Rode scripts/etiquetas-caixa-schema.sql.",
        };
      }
      return { ok: false, code: "db", error: error.message || "Erro ao buscar etiqueta." };
    }
    if (data) {
      row = data;
      break;
    }
  }

  // Fallback: busca por contador_n numérico.
  if (!row) {
    const n = Number(normalizeContadorScan(rawContador)) || 0;
    if (n > 0) {
      const { data, error } = await supabase
        .from("etiquetas_caixa")
        .select("*")
        .eq("contador_n", n)
        .maybeSingle();
      if (error) {
        return { ok: false, code: "db", error: error.message || "Erro ao buscar etiqueta." };
      }
      row = data;
    }
  }

  if (!row) {
    return {
      ok: false,
      code: "not_found",
      error: "Etiqueta não encontrada. Imprima/baixe o ZPL antes de bipar na esteira.",
    };
  }

  if (row.confirmado_em || row.bipado_em) {
    return {
      ok: false,
      code: "already",
      error: `Caixa ${row.contador} já foi bipada — não pode ler de novo.`,
      row,
    };
  }

  const now = new Date().toISOString();
  const { data: updated, error: updErr } = await supabase
    .from("etiquetas_caixa")
    .update({ bipado_em: now })
    .eq("contador", row.contador)
    .is("bipado_em", null)
    .select("*")
    .maybeSingle();

  if (updErr) {
    return { ok: false, code: "db", error: updErr.message || "Falha ao reservar bip." };
  }
  if (!updated) {
    return {
      ok: false,
      code: "already",
      error: `Caixa ${row.contador} já foi bipada — não pode ler de novo.`,
      row,
    };
  }

  return { ok: true, code: "ok", row: updated, error: null };
}

/** Desfaz reserva do último bip (só se ainda não confirmado). */
export async function liberarBipPorContador(contador) {
  const key = String(contador || "").trim();
  if (!key) return { ok: false, error: "Contador vazio." };
  const { data, error } = await supabase
    .from("etiquetas_caixa")
    .update({ bipado_em: null })
    .eq("contador", key)
    .is("confirmado_em", null)
    .select("contador")
    .maybeSingle();
  if (error) return { ok: false, error: error.message };
  if (!data) return { ok: false, error: "Não foi possível desfazer (já confirmada?)." };
  return { ok: true, error: null };
}

/** Marca lote como confirmado (após M→PA + Bling). */
export async function confirmarBipsNoBanco(contadores) {
  const keys = (contadores || []).map((c) => String(c).trim()).filter(Boolean);
  if (!keys.length) return { ok: true, error: null };
  const now = new Date().toISOString();
  const { error } = await supabase
    .from("etiquetas_caixa")
    .update({ confirmado_em: now, bipado_em: now })
    .in("contador", keys);
  if (error) return { ok: false, error };
  return { ok: true, error: null };
}

export function syncLocalContadorFromN(n, width = 7) {
  writeStoredContador(formatContador(n, width));
}
