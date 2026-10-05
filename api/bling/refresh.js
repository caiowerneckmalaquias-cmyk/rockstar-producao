import {
  exchangeToken,
  expiresAtFromToken,
  getConexao,
  json,
  upsertConexao,
} from "../_lib/bling.js";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const row = await getConexao();
    if (!row?.refresh_token) {
      return json(res, 400, { error: "Nenhuma conexão Bling para renovar" });
    }

    const tokenJson = await exchangeToken({
      grantType: "refresh_token",
      refreshToken: row.refresh_token,
    });

    await upsertConexao({
      access_token: tokenJson.access_token,
      refresh_token: tokenJson.refresh_token || row.refresh_token,
      token_type: tokenJson.token_type || "Bearer",
      expires_at: expiresAtFromToken(tokenJson),
      scopes: tokenJson.scope || row.scopes,
      deposito_pa_id: row.deposito_pa_id,
      deposito_est_id: row.deposito_est_id,
      deposito_pa_nome: row.deposito_pa_nome,
      deposito_est_nome: row.deposito_est_nome,
      connected_at: row.connected_at,
    });

    return json(res, 200, {
      ok: true,
      expiresAt: expiresAtFromToken(tokenJson),
    });
  } catch (e) {
    console.error("bling refresh", e);
    return json(res, 500, { error: e.message || "Falha no refresh" });
  }
}
