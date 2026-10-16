import {
  STOCK_OFFSET,
  blingToReal,
  buildBlingSku,
  createServiceClient,
  fetchSaldosProduto,
  findProductByCodigo,
  getValidAccessToken,
  json,
  listContasSecundariasProntas,
  saldoNoDeposito,
  setEstoqueSaldo,
  shouldZeroBling,
  sleep,
} from "../_lib/bling.js";

/** Vercel Pro: até 300s — o pull é sequencial e demora com muitos SKUs. */
export const config = { maxDuration: 300 };

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

    let secundarias = [];
    try {
      secundarias = await listContasSecundariasProntas();
    } catch (e) {
      console.warn("listContasSecundariasProntas:", e.message);
      secundarias = [];
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

    const { data: estoqueAtual, error: erroEst } = await supabase
      .from("estoque")
      .select("ref,cor,numero,pa,est,m,p,disp");
    if (erroEst) throw new Error(erroEst.message);

    const estoqueMap = new Map();
    (estoqueAtual || []).forEach((e) => {
      estoqueMap.set(`${e.ref}__${e.cor}__${e.numero}`, e);
    });

    const skuCachePrincipal = new Map();
    /** cache por slot → Map(sku → produto|null) */
    const skuCacheSec = new Map();
    secundarias.forEach(({ conta }) => {
      skuCacheSec.set(conta.slot, new Map());
    });

    const toUpsert = [];
    let ok = 0;
    let notFound = 0;
    let zeroedOnBling = 0;
    const samples = [];

    for (const prod of comSigla) {
      const ref = String(prod.ref || "").trim();
      const cor = String(prod.cor || "").trim();
      const sigla = String(prod.bling_sigla || "").trim();

      for (const size of SIZES) {
        const sku = buildBlingSku(ref, sigla, size);
        if (!sku) continue;

        let produtoBling = skuCachePrincipal.get(sku);
        if (produtoBling === undefined) {
          await sleep(350);
          try {
            produtoBling = await findProductByCodigo(accessToken, sku);
          } catch (e) {
            produtoBling = null;
            if (samples.length < 8) {
              samples.push({ sku, error: e.message });
            }
          }
          skuCachePrincipal.set(sku, produtoBling || null);
        }

        if (!produtoBling?.id) {
          notFound += 1;
          continue;
        }

        await sleep(350);
        let saldos;
        try {
          saldos = await fetchSaldosProduto(accessToken, produtoBling.id);
        } catch (e) {
          if (samples.length < 8) {
            samples.push({ sku, error: `saldos: ${e.message}` });
          }
          continue;
        }

        let paBling = saldoNoDeposito(saldos, depositoPaId);
        let estBling = saldoNoDeposito(saldos, depositoEstId);

        if (shouldZeroBling(paBling, STOCK_OFFSET)) {
          if (paBling !== 0) {
            await sleep(350);
            try {
              await setEstoqueSaldo({
                accessToken,
                produtoId: produtoBling.id,
                depositoId: depositoPaId,
                quantidade: 0,
                observacoes: "Rock Star: ≤ offset → zerar PA",
              });
              zeroedOnBling += 1;
            } catch (_) {
              /* segue com real 0 */
            }
          }
          paBling = 0;
        }
        if (shouldZeroBling(estBling, STOCK_OFFSET)) {
          if (estBling !== 0) {
            await sleep(350);
            try {
              await setEstoqueSaldo({
                accessToken,
                produtoId: produtoBling.id,
                depositoId: depositoEstId,
                quantidade: 0,
                observacoes: "Rock Star: ≤ offset → zerar EST",
              });
              zeroedOnBling += 1;
            } catch (_) {
              /* segue */
            }
          }
          estBling = 0;
        }

        const paReal = blingToReal(paBling, STOCK_OFFSET);
        const estReal = blingToReal(estBling, STOCK_OFFSET);

        // DISP = PA principal + soma(Geral das secundárias). Geral sem offset.
        let somaGeral = 0;
        for (const { accessToken: tokenSec, conta } of secundarias) {
          const cache = skuCacheSec.get(conta.slot);
          let prodSec = cache.get(sku);
          if (prodSec === undefined) {
            await sleep(350);
            try {
              prodSec = await findProductByCodigo(tokenSec, sku);
            } catch (_) {
              prodSec = null;
            }
            cache.set(sku, prodSec || null);
          }
          if (!prodSec?.id) continue;

          await sleep(350);
          try {
            const saldosSec = await fetchSaldosProduto(tokenSec, prodSec.id);
            const geralRaw = saldoNoDeposito(saldosSec, conta.deposito_geral_id);
            somaGeral += Number(geralRaw) || 0;
          } catch (e) {
            if (samples.length < 8) {
              samples.push({
                sku,
                error: `sec${conta.slot} saldos: ${e.message}`,
              });
            }
          }
        }

        const dispReal = Math.round(paReal + somaGeral);
        const prev = estoqueMap.get(`${ref}__${cor}__${size}`) || {};

        toUpsert.push({
          ref,
          cor,
          numero: size,
          pa: paReal,
          est: estReal,
          disp: dispReal,
          m: Number(prev.m) || 0,
          p: Number(prev.p) || 0,
        });
        ok += 1;
      }
    }

    if (toUpsert.length) {
      const chunk = 200;
      for (let i = 0; i < toUpsert.length; i += chunk) {
        const slice = toUpsert.slice(i, i + chunk);
        const { error } = await supabase
          .from("estoque")
          .upsert(slice, { onConflict: "ref,cor,numero" });
        if (error) throw new Error(error.message);
      }
    }

    const secundariasUsadas = secundarias.length;
    return json(res, 200, {
      ok: true,
      updated: toUpsert.length,
      skusOk: ok,
      semSigla,
      notFound,
      zeroedOnBling,
      secundariasUsadas,
      secundariasSlots: secundarias.map(({ conta }) => ({
        slot: conta.slot,
        depositoGeralId: conta.deposito_geral_id,
        depositoGeralNome: conta.deposito_geral_nome,
      })),
      stockOffset: STOCK_OFFSET,
      samples,
      warning:
        secundariasUsadas > 0
          ? null
          : "Nenhuma conta secundária pronta — DISP = PA. Conecte as 4 contas (Geral) e puxe de novo.",
    });
  } catch (e) {
    console.error("pull-estoque", e);
    return json(res, 500, { error: e.message || "Falha ao puxar estoque" });
  }
}
