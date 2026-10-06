/**
 * Smoke test da regra +1000 e SKU (sem rede).
 * Uso: node scripts/verify-bling-offset.mjs
 */
import {
  BLING_STOCK_OFFSET,
  blingToReal,
  realToBling,
  shouldZeroBling,
  buildBlingSku,
  parseCorSiglaLine,
  normalizeBlingSigla,
} from "../src/constants/bling.js";

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

assert(BLING_STOCK_OFFSET === 1000, "offset padrão 1000");
assert(realToBling(10) === 1010, "real 10 → 1010");
assert(realToBling(0) === 0, "real 0 → 0 no Bling");
assert(blingToReal(1010) === 10, "1010 → real 10");
assert(blingToReal(1000) === 0, "1000 → real 0");
assert(blingToReal(999) === 0, "999 → real 0");
assert(shouldZeroBling(1000) === true, "≤1000 zera");
assert(shouldZeroBling(1001) === false, "1001 não zera");
assert(shouldZeroBling(0) === true, "0 zera");
assert(buildBlingSku("TNCV010", "RS", 34) === "TNCV010RS34", "SKU");
assert(buildBlingSku("tncv010", "rs", "35") === "TNCV010RS35", "SKU case");
assert(normalizeBlingSigla("r-s") === "RS", "sigla limpa");
const parsed = parseCorSiglaLine("ROSE|RS");
assert(parsed.cor === "ROSE" && parsed.sigla === "RS", "parse COR|SIGLA");
assert(parseCorSiglaLine("PRETO").sigla === "", "cor sem sigla");

console.log("PASS: regra Bling +1000 + SKU/sigla");
