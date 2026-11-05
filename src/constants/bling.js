/** Sigla limpa: só letras/números, maiúscula. */
export function normalizeBlingSigla(sigla) {
  return String(sigla || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

/** SKU filho Bling: REF + SIGLA + tamanho (ex.: TNCV010RS34). */
export function buildBlingSku(ref, sigla, size) {
  const r = String(ref || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
  const s = normalizeBlingSigla(sigla);
  const n = String(size ?? "").trim();
  if (!r || !s || !n) return "";
  return `${r}${s}${n}`;
}

/**
 * Linha de Nova referência: "ROSE|RS" ou só "ROSE".
 * @returns {{ cor: string, sigla: string }}
 */
export function parseCorSiglaLine(linha) {
  const raw = String(linha || "").trim();
  if (!raw) return { cor: "", sigla: "" };
  const pipe = raw.indexOf("|");
  if (pipe === -1) {
    return { cor: raw.toUpperCase(), sigla: "" };
  }
  const cor = raw.slice(0, pipe).trim().toUpperCase();
  const sigla = normalizeBlingSigla(raw.slice(pipe + 1));
  return { cor, sigla };
}

export const BLING_DEPOSITO_PA_LABEL = "PRODUTO ACABADO";
export const BLING_DEPOSITO_EST_LABEL = "OVERLOQUE";
