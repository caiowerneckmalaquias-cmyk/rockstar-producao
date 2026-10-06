import {
  STOCK_OFFSET,
  buildBlingSku,
  createServiceClient,
  findProductByCodigo,
  getValidAccessToken,
  json,
  realToBling,
  setEstoqueSaldo,
  sleep,
} from "../_lib/bling.js";

const SIZES = [34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44];

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const { accessToken, conexao } = await getValidAccessToken();
    const depositoPaId = conexao.deposito_pa_id;
    const depositoEstId = conexao.deposito_est_id;
    if (!depositoPaId || !depositoEstId) {
      return json(res, 400, {
        error:
          "Depósitos PA/EST não vinculados. Reconecte o Bling (PRODUTO ACABADO + OVERLOQUE).",
      });
    }

    const supabase = createServiceClient();
    const { data: produtos, error: erroProd } = await supabase
      .from("produtos")
      .select("ref,cor,bling_sigla");
    if (erroProd) throw new Error(erroProd.message);

    const comSigla = (produtos || []).filter((p) =>
      String(p.bling_sigla || "").trim()
    );
    const semSigla = (produtos || []).length - comSigla.length;

    const { data: estoque, error: erroEst } = await supabase
      .from("estoque")
      .select("ref,cor,numero,pa,est");
    if (erroEst) throw new Error(erroEst.message);

    const estoqueMap = new Map();
    (estoque || []).forEach((e) => {
      estoqueMap.set(`${e.ref}__${e.cor}__${e.numero}`, e);
    });

    const skuCache = new Map();
    let pushed = 0;
    let notFound = 0;
    const samples = [];

    for (const prod of comSigla) {
      const ref = String(prod.ref || "").trim();
      const cor = String(prod.cor || "").trim();
      const sigla = String(prod.bling_sigla || "").trim();

      for (const size of SIZES) {
        const sku = buildBlingSku(ref, sigla, size);
        if (!sku) continue;

        let produtoBling = skuCache.get(sku);
        if (produtoBling === undefined) {
          await sleep(350);
          try {
            produtoBling = await findProductByCodigo(accessToken, sku);
          } catch (e) {
            produtoBling = null;
            if (samples.length < 8) samples.push({ sku, error: e.message });
          }
          skuCache.set(sku, produtoBling || null);
        }

        if (!produtoBling?.id) {
          notFound += 1;
          continue;
        }

        const cell = estoqueMap.get(`${ref}__${cor}__${size}`) || {};
        const paBling = realToBling(cell.pa, STOCK_OFFSET);
        const estBling = realToBling(cell.est, STOCK_OFFSET);

        await sleep(350);
        try {
          await setEstoqueSaldo({
            accessToken,
            produtoId: produtoBling.id,
            depositoId: depositoPaId,
            quantidade: paBling,
            observacoes: "Rock Star: push PA (real+offset)",
          });
          pushed += 1;
        } catch (e) {
          if (samples.length < 8) {
            samples.push({ sku, deposito: "PA", error: e.message });
          }
        }

        await sleep(350);
        try {
          await setEstoqueSaldo({
            accessToken,
            produtoId: produtoBling.id,
            depositoId: depositoEstId,
            quantidade: estBling,
            observacoes: "Rock Star: push EST (real+offset)",
          });
          pushed += 1;
        } catch (e) {
          if (samples.length < 8) {
            samples.push({ sku, deposito: "EST", error: e.message });
          }
        }
      }
    }

    return json(res, 200, {
      ok: true,
      pushed,
      semSigla,
      notFound,
      stockOffset: STOCK_OFFSET,
      samples,
    });
  } catch (e) {
    console.error("push-estoque", e);
    return json(res, 500, { error: e.message || "Falha ao subir estoque" });
  }
}
