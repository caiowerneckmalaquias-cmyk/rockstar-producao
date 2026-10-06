/** Offset Shopee/Bling: estoque exibido = real + OFFSET */
export const BLING_STOCK_OFFSET = 1000;

/**
 * Bling/Shopee → estoque real do Rock Star.
 * Ex.: 1010 → 10; 1000 ou menos → 0.
 */
export function blingToReal(blingQty, offset = BLING_STOCK_OFFSET) {
  const n = Number(blingQty);
  if (!Number.isFinite(n)) return 0;
  const off = Number(offset) || 0;
  return Math.max(0, n - off);
}

/**
 * Estoque real → valor a gravar no Bling (para a Shopee).
 * Ex.: 10 → 1010. Se real é 0, sobe 0 (zerado).
 */
export function realToBling(realQty, offset = BLING_STOCK_OFFSET) {
  const n = Number(realQty);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return n + (Number(offset) || 0);
}

/** Estoque no Bling ≤ offset ⇒ real zerou ⇒ forçar Bling = 0. */
export function shouldZeroBling(blingQty, offset = BLING_STOCK_OFFSET) {
  const n = Number(blingQty);
  if (!Number.isFinite(n)) return true;
  return n <= (Number(offset) || 0);
}

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
