import {
  exchangeToken,
  expiresAtFromToken,
  getEnv,
  listDepositos,
  matchDepositos,
  redirect,
  upsertConexao,
} from "../_lib/bling.js";

function readCookie(req, name) {
  const raw = req.headers.cookie || "";
  const parts = raw.split(";").map((p) => p.trim());
  for (const part of parts) {
    const eq = part.indexOf("=");
    if (eq === -1) continue;
    if (part.slice(0, eq) === name) return decodeURIComponent(part.slice(eq + 1));
  }
  return "";
}

export default async function handler(req, res) {
  const { appUrl } = getEnv();
  const fail = (msg) =>
    redirect(res, `${appUrl}/?bling=error&msg=${encodeURIComponent(msg)}`);

  try {
    if (req.method !== "GET") return fail("Método inválido");

    const url = new URL(req.url, `https://${req.headers.host}`);
    const code = url.searchParams.get("code");
    const state = url.searchParams.get("state") || "";
    const err = url.searchParams.get("error");
    const errDesc = url.searchParams.get("error_description");

    if (err) return fail(errDesc || err);
    if (!code) return fail("Código de autorização ausente");

    const expected = readCookie(req, "bling_oauth_state");
    if (expected && state && expected !== state) {
      return fail("State OAuth inválido");
    }

    const tokenJson = await exchangeToken({
      grantType: "authorization_code",
      code,
    });

    const accessToken = tokenJson.access_token;
    const refreshToken = tokenJson.refresh_token;
    if (!accessToken || !refreshToken) {
      return fail("Tokens não retornados pelo Bling");
    }

    let depositos = [];
    try {
      depositos = await listDepositos(accessToken);
    } catch (e) {
      // Conecta mesmo se o escopo de depósitos falhar; IDs ficam null
      console.warn("listDepositos:", e.message);
    }

    const matched = matchDepositos(depositos);

    await upsertConexao({
      access_token: accessToken,
      refresh_token: refreshToken,
      token_type: tokenJson.token_type || "Bearer",
      expires_at: expiresAtFromToken(tokenJson),
      scopes: tokenJson.scope || null,
      deposito_pa_id: matched.depositoPaId,
      deposito_est_id: matched.depositoEstId,
      deposito_pa_nome: matched.depositoPaNome,
      deposito_est_nome: matched.depositoEstNome,
      connected_at: new Date().toISOString(),
    });

    res.setHeader(
      "Set-Cookie",
      "bling_oauth_state=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0"
    );
    return redirect(res, `${appUrl}/?bling=connected`);
  } catch (e) {
    console.error("bling callback", e);
    return fail(e.message || "Erro no callback Bling");
  }
}
