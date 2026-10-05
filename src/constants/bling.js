/** Offset Shopee/Bling: estoque exibido = real + OFFSET */
export const BLING_STOCK_OFFSET = 1000;

/**
 * Bling/Shopee → estoque real do Rock Star.
 * Ex.: 1010 → 10; 1000 ou menos → 0.
 */
export function blingToReal(blingQty, offset = BLING_STOCK_OFFSET) {
  const n = Number(blingQty);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, n - Number(offset) || 0);
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

export const BLING_DEPOSITO_PA_LABEL = "PRODUTO ACABADO";
export const BLING_DEPOSITO_EST_LABEL = "OVERLOQUE";
