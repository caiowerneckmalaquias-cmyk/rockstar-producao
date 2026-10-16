import {
  BLING_AUTHORIZE_URL,
  createState,
  getEnv,
  json,
  normalizeSecundariaSlot,
  redirect,
} from "../_lib/bling.js";

function buildSetCookie(name, value, maxAge = 600) {
  return `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${maxAge}`;
}

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

  const urlReq = new URL(req.url, `https://${req.headers.host}`);
  const tipo = String(urlReq.searchParams.get("tipo") || "principal").toLowerCase();
  const slotRaw = urlReq.searchParams.get("slot");
  let destino = "principal";

  if (tipo === "secundaria") {
    const slot = normalizeSecundariaSlot(slotRaw);
    if (slot == null) {
      return json(res, 400, { error: "Slot secundário inválido (use 1–4)" });
    }
    destino = `secundaria:${slot}`;
  }

  const state = createState();
  res.setHeader("Set-Cookie", [
    buildSetCookie("bling_oauth_state", state),
    buildSetCookie("bling_oauth_destino", destino),
  ]);

  const url = new URL(BLING_AUTHORIZE_URL);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("state", state);
  url.searchParams.set("redirect_uri", redirectUri);

  return redirect(res, url.toString());
}
