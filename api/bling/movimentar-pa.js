import {
  STOCK_OFFSET,
  calcularAjusteComOffset,
  fetchSaldosProduto,
  findProductByCodigo,
  getValidAccessToken,
  json,
  lancarMovimentoEstoque,
  saldoNoDeposito,
  sleep,
} from "../_lib/bling.js";

/**
 * Aplica +N/−N pares no depósito PA do Bling por SKU (Esteira).
 * Body: { movimentos: [{ sku, delta }] }. Responde o resultado de cada SKU para o app tirar da fila só o que deu certo.
 */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const body = req.body && typeof req.body === "object" ? req.body : {};
    const porSku = new Map();
    (Array.isArray(body.movimentos) ? body.movimentos : []).forEach((m) => {
      const sku = String(m?.sku || "").trim().toUpperCase();
      const delta = Math.trunc(Number(m?.delta) || 0);
      if (!sku || !delta) return;
      porSku.set(sku, (porSku.get(sku) || 0) + delta);
    });

    if (!porSku.size) {
      return json(res, 200, { ok: true, resultados: [] });
    }

    const { accessToken, conexao } = await getValidAccessToken();
    const depositoPaId = conexao.deposito_pa_id;
    if (!depositoPaId) {
      return json(res, 400, {
        error: "Depósito PA não vinculado. Reconecte o Bling (PRODUTO ACABADO).",
      });
    }

    const resultados = [];
    for (const [sku, delta] of porSku) {
      if (!delta) {
        resultados.push({ sku, delta, ok: true, semMudanca: true });
        continue;
      }
      try {
        await sleep(350);
        const produto = await findProductByCodigo(accessToken, sku);
        if (!produto?.id) {
          resultados.push({ sku, delta, ok: false, permanente: true, error: "SKU não encontrado no Bling" });
          continue;
        }

        await sleep(350);
        const saldos = await fetchSaldosProduto(accessToken, produto.id);
        const saldoAntes = saldoNoDeposito(saldos, depositoPaId);
        const ajuste = calcularAjusteComOffset(saldoAntes, delta, STOCK_OFFSET);

        if (ajuste.diff !== 0) {
          await sleep(350);
          await lancarMovimentoEstoque({
            accessToken,
            produtoId: produto.id,
            depositoId: depositoPaId,
            operacao: ajuste.diff > 0 ? "E" : "S",
            quantidade: Math.abs(ajuste.diff),
            observacoes: `Rock Star Esteira: ${delta > 0 ? "+" : ""}${delta} par(es)`,
          });
        }

        resultados.push({
          sku,
          delta,
          ok: true,
          saldoBlingAntes: saldoAntes,
          saldoBlingDepois: ajuste.alvoBling,
          realDepois: ajuste.realDepois,
        });
      } catch (e) {
        resultados.push({ sku, delta, ok: false, error: e.message || "Falha no Bling" });
      }
    }

    const falhas = resultados.filter((r) => !r.ok).length;
    return json(res, 200, { ok: falhas === 0, falhas, resultados, stockOffset: STOCK_OFFSET });
  } catch (e) {
    console.error("movimentar-pa", e);
    return json(res, 500, { error: e.message || "Falha ao movimentar PA no Bling" });
  }
}
