import {
  STOCK_OFFSET,
  blingToReal,
  buildBlingSku,
  createServiceClient,
  discoverMultiempresaFromSaldos,
  fetchSaldosProduto,
  findProductByCodigo,
  getValidAccessToken,
  json,
  saldoNoDeposito,
  setEstoqueSaldo,
  shouldZeroBling,
  sleep,
  upsertConexao,
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
    let depositoMultiId = conexao.deposito_multi_id ?? null;
    let depositoMultiNome = conexao.deposito_multi_nome ?? null;
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

    const { data: estoqueAtual, error: erroEst } = await supabase
      .from("estoque")
      .select("ref,cor,numero,pa,est,m,p,disp");
    if (erroEst) throw new Error(erroEst.message);

    const estoqueMap = new Map();
    (estoqueAtual || []).forEach((e) => {
      estoqueMap.set(`${e.ref}__${e.cor}__${e.numero}`, e);
    });

    const skuCache = new Map();
    const toUpsert = [];
    let ok = 0;
    let notFound = 0;
    let zeroedOnBling = 0;
    const samples = [];
    let multiDescobertoNoPull = false;

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
            if (samples.length < 8) {
              samples.push({ sku, error: e.message });
            }
          }
          skuCache.set(sku, produtoBling || null);
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

        // Fallback: virtual Multiempresa às vezes não vem em GET /depositos.
        if (depositoMultiId == null) {
          const found = discoverMultiempresaFromSaldos(saldos);
          if (found?.id != null) {
            depositoMultiId = found.id;
            depositoMultiNome = found.nome || "Multiempresa";
            multiDescobertoNoPull = true;
            try {
              await upsertConexao({
                access_token: conexao.access_token,
                refresh_token: conexao.refresh_token,
                token_type: conexao.token_type || "Bearer",
                expires_at: conexao.expires_at,
                scopes: conexao.scopes,
                deposito_pa_id: depositoPaId,
                deposito_est_id: depositoEstId,
                deposito_multi_id: depositoMultiId,
                deposito_pa_nome: conexao.deposito_pa_nome,
                deposito_est_nome: conexao.deposito_est_nome,
                deposito_multi_nome: depositoMultiNome,
                connected_at: conexao.connected_at,
              });
            } catch (persistErr) {
              console.warn("persist multi from saldos:", persistErr.message);
            }
          }
        }

        const multiVinculado = depositoMultiId != null;

        let paBling = saldoNoDeposito(saldos, depositoPaId);
        let estBling = saldoNoDeposito(saldos, depositoEstId);
        const multiRaw = multiVinculado
          ? saldoNoDeposito(saldos, depositoMultiId)
          : null;

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
        // Multiempresa = disponível consolidado (unidades reais, sem offset Shopee).
        const dispReal = multiVinculado
          ? Math.max(0, Math.round(Number(multiRaw) || 0))
          : paReal;
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

    const multiVinculado = depositoMultiId != null;

    return json(res, 200, {
      ok: true,
      updated: toUpsert.length,
      skusOk: ok,
      semSigla,
      notFound,
      zeroedOnBling,
      multiVinculado,
      multiDescobertoNoPull,
      depositoMultiId: depositoMultiId ?? null,
      depositoMultiNome: depositoMultiNome ?? null,
      stockOffset: STOCK_OFFSET,
      samples,
      warning: multiVinculado
        ? null
        : "Depósito Multiempresa não vinculado — DISP = PA (fallback). Use Vincular depósitos ou reconecte o Bling.",
    });
  } catch (e) {
    console.error("pull-estoque", e);
    return json(res, 500, { error: e.message || "Falha ao puxar estoque" });
  }
}
