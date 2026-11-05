/**
 * Smoke test de SKU/sigla do Bling (sem rede).
 * Uso: node scripts/verify-bling-offset.mjs
 */
import {
  buildBlingSku,
  parseCorSiglaLine,
  normalizeBlingSigla,
} from "../src/constants/bling.js";

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

assert(buildBlingSku("TNCV010", "RS", 34) === "TNCV010RS34", "SKU");
assert(buildBlingSku("tncv010", "rs", "35") === "TNCV010RS35", "SKU case");
assert(normalizeBlingSigla("r-s") === "RS", "sigla limpa");
const parsed = parseCorSiglaLine("ROSE|RS");
assert(parsed.cor === "ROSE" && parsed.sigla === "RS", "parse COR|SIGLA");
assert(parseCorSiglaLine("PRETO").sigla === "", "cor sem sigla");

console.log("PASS: SKU/sigla Bling");
