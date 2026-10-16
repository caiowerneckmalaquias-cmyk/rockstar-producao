import { getValidAccessToken, json, upsertConexao } from "../_lib/bling.js";

/**
 * Grava manualmente o id do depósito Multiempresa (DISP).
 * Body: { depositoMultiId: number|string, depositoMultiNome?: string }
 */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const { conexao } = await getValidAccessToken();
    const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {};
    const rawId = body.depositoMultiId ?? body.deposito_multi_id;
    const idNum = Number(rawId);
    if (!Number.isFinite(idNum) || idNum <= 0) {
      return json(res, 400, {
        error: "Informe um depósito Multiempresa id numérico válido.",
      });
    }

    const nome =
      String(body.depositoMultiNome || body.deposito_multi_nome || "Multiempresa").trim() ||
      "Multiempresa";

    const updated = await upsertConexao({
      access_token: conexao.access_token,
      refresh_token: conexao.refresh_token,
      token_type: conexao.token_type || "Bearer",
      expires_at: conexao.expires_at,
      scopes: conexao.scopes,
      deposito_pa_id: conexao.deposito_pa_id,
      deposito_est_id: conexao.deposito_est_id,
      deposito_multi_id: Math.trunc(idNum),
      deposito_pa_nome: conexao.deposito_pa_nome,
      deposito_est_nome: conexao.deposito_est_nome,
      deposito_multi_nome: nome,
      connected_at: conexao.connected_at,
    });

    return json(res, 200, {
      ok: true,
      depositoMultiId: updated.deposito_multi_id ?? Math.trunc(idNum),
      depositoMultiNome: updated.deposito_multi_nome ?? nome,
    });
  } catch (e) {
    console.error("bling set-deposito-multi", e);
    return json(res, 500, { error: e.message || "Falha ao salvar Multiempresa" });
  }
}
