import {
  BLING_AUTHORIZE_URL,
  createState,
  getEnv,
  json,
  redirect,
} from "../_lib/bling.js";

export default function handler(req, res) {
  if (req.method !== "GET") {
    return json(res, 405, { error: "Method not allowed" });
  }

  const { clientId, redirectUri } = getEnv();
  if (!clientId) {
    return json(res, 500, {
      error: "BLING_CLIENT_ID não configurado nas variáveis de ambiente da Vercel",
    });
  }

  const state = createState();
  // Cookie curto só para validar o round-trip OAuth
  res.setHeader(
    "Set-Cookie",
    `bling_oauth_state=${state}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=600`
  );

  const url = new URL(BLING_AUTHORIZE_URL);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("state", state);
  url.searchParams.set("redirect_uri", redirectUri);

  return redirect(res, url.toString());
}
