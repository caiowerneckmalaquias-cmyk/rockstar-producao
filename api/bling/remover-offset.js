import {
  OFFSET_LEGADO,
  buildBlingSku,
  createServiceClient,
  fetchSaldosProduto,
  findProductByCodigo,
  getValidAccessToken,
  json,
  lancarMovimentoEstoque,
  saldoNoDeposito,
  sleep,
} from "../_lib/bling.js";

export const config = { maxDuration: 300 };

const SIZES = [34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44];
const LOTE_PADRAO = 25;

/**
 * Conversão única: tira os +1000 antigos de PA (PRODUTO ACABADO) e EST (OVERLOQUE).
 * Só mexe em saldo >= 1000 e usa saída relativa, então rodar de novo não desconta duas vezes
 * e vendas lançadas durante a conversão são respeitadas.
 * Body: { inicio?: number, lote?: number }. Responde `proximo` (null quando terminou).
 */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const body = req.body && typeof req.body === "object" ? req.body : {};
    const inicio = Math.max(0, Math.trunc(Number(body.inicio) || 0));
    const lote = Math.min(50, Math.max(1, Math.trunc(Number(body.lote) || LOTE_PADRAO)));

    const { accessToken, conexao } = await getValidAccessToken();
    const depositos = [
      { id: conexao.deposito_pa_id, nome: "PA" },
      { id: conexao.deposito_est_id, nome: "EST" },
    ].filter((d) => d.id);
    if (!depositos.length) {
      return json(res, 400, {
        error: "Depósitos PA/EST não vinculados. Reconecte o Bling (PRODUTO ACABADO + OVERLOQUE).",
      });
    }

    const supabase = createServiceClient();
    const { data: produtos, error: erroProd } = await supabase
      .from("produtos")
      .select("ref,cor,bling_sigla");
    if (erroProd) throw new Error(erroProd.message);

    const skus = Array.from(
      new Set(
        (produtos || [])
          .filter((p) => String(p.bling_sigla || "").trim())
          .flatMap((p) => SIZES.map((size) => buildBlingSku(p.ref, p.bling_sigla, size)))
          .filter(Boolean)
      )
    ).sort();
    const semSigla = (produtos || []).filter((p) => !String(p.bling_sigla || "").trim()).length;

    const fatia = skus.slice(inicio, inicio + lote);
    let convertidos = 0;
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

        for (const dep of depositos) {
          const saldo = saldoNoDeposito(saldos, dep.id);
          if (saldo < OFFSET_LEGADO) continue;
          await sleep(350);
          await lancarMovimentoEstoque({
            accessToken,
            produtoId: produto.id,
            depositoId: dep.id,
            operacao: "S",
            quantidade: OFFSET_LEGADO,
            observacoes: "Rock Star: remover regra +1000",
          });
          convertidos += 1;
          if (samples.length < 8) {
            samples.push({ sku, deposito: dep.nome, antes: saldo, depois: saldo - OFFSET_LEGADO });
          }
        }
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
      convertidos,
      naoAchados,
      erros,
      semSigla,
      samples,
    });
  } catch (e) {
    console.error("remover-offset", e);
    return json(res, 500, { error: e.message || "Falha ao remover a regra dos 1000" });
  }
}
