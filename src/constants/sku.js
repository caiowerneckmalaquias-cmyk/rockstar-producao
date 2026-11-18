/** Sigla limpa: só letras/números, maiúscula. */
export function normalizeSigla(sigla) {
  return String(sigla || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

/** SKU da etiqueta: REF + SIGLA + tamanho (ex.: TNCV010RS34). */
export function buildSku(ref, sigla, size) {
  const r = String(ref || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
  const s = normalizeSigla(sigla);
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
  const sigla = normalizeSigla(raw.slice(pipe + 1));
  return { cor, sigla };
}
