import {
  createServiceClient,
  fetchSaldosProduto,
  findProductByCodigo,
  getValidAccessToken,
  json,
  listarSkusComSigla,
  saldoNoDeposito,
  setEstoqueSaldo,
  sleep,
} from "../_lib/bling.js";

export const config = { maxDuration: 300 };

const LOTE_PADRAO = 25;
const SENHA = String(process.env.SENHA_ZERAR_PA || process.env.VITE_SENHA_ZERAR_EST || "132701").trim();

/**
 * Zera o depósito PRODUTO ACABADO (PA) no Bling, em lotes. Só lança balanço nos SKUs com saldo ≠ 0,
 * então rodar de novo depois de uma interrupção continua de onde parou.
 * Body: { senha, confirmacao: "ZERAR", inicio?, lote? }. Responde `proximo` (null quando terminou).
 */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const body = req.body && typeof req.body === "object" ? req.body : {};
    if (String(body.senha || "").trim() !== SENHA) {
      return json(res, 403, { error: "Senha incorreta." });
    }
    if (String(body.confirmacao || "").trim().toUpperCase() !== "ZERAR") {
      return json(res, 400, { error: "Confirmação ZERAR ausente." });
    }

    const inicio = Math.max(0, Math.trunc(Number(body.inicio) || 0));
    const lote = Math.min(50, Math.max(1, Math.trunc(Number(body.lote) || LOTE_PADRAO)));

    const { accessToken, conexao } = await getValidAccessToken();
    const depositoPaId = conexao.deposito_pa_id;
    if (!depositoPaId) {
      return json(res, 400, { error: "Depósito PA não vinculado. Reconecte o Bling (PRODUTO ACABADO)." });
    }

    const { skus } = await listarSkusComSigla(createServiceClient());
    const fatia = skus.slice(inicio, inicio + lote);
    let zerados = 0;
    let paresRemovidos = 0;
    let naoAchados = 0;
    let erros = 0;
    const samples = [];

    for (const sku of fatia) {
      try {
        await sleep(350);
        const produto = await findProductByCodigo(accessToken, sku);
        if (!produto?.id) {
          naoAchados += 1;
          continue;
        }

        await sleep(350);
        const saldos = await fetchSaldosProduto(accessToken, produto.id);
        const saldo = saldoNoDeposito(saldos, depositoPaId);
        if (saldo === 0) continue;

        await sleep(350);
        await setEstoqueSaldo({
          accessToken,
          produtoId: produto.id,
          depositoId: depositoPaId,
          quantidade: 0,
          observacoes: "Rock Star: zerar produto acabado (operação protegida)",
        });
        zerados += 1;
        paresRemovidos += Math.max(0, saldo);
      } catch (e) {
        erros += 1;
        if (samples.length < 8) samples.push({ sku, error: e.message });
      }
    }

    const fim = inicio + fatia.length;
    return json(res, 200, {
      ok: true,
      total: skus.length,
      inicio,
      fim,
      proximo: fim < skus.length ? fim : null,
      zerados,
      paresRemovidos,
      naoAchados,
      erros,
      samples,
    });
  } catch (e) {
    console.error("zerar-pa", e);
    return json(res, 500, { error: e.message || "Falha ao zerar produto acabado" });
  }
}
