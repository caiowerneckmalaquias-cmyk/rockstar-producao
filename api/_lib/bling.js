import { createClient } from "@supabase/supabase-js";

export const BLING_AUTHORIZE_URL = "https://www.bling.com.br/Api/v3/oauth/authorize";
export const BLING_TOKEN_URL = "https://www.bling.com.br/Api/v3/oauth/token";
export const BLING_API_BASE = "https://api.bling.com.br/Api/v3";

export const STOCK_OFFSET = Number(process.env.BLING_STOCK_OFFSET) || 1000;

const DEPOSITO_PA_NAMES = ["PRODUTO ACABADO", "PRODUTOACABADO"];
const DEPOSITO_EST_NAMES = ["OVERLOQUE"];

export function getEnv() {
  const clientId = process.env.BLING_CLIENT_ID || "";
  const clientSecret = process.env.BLING_CLIENT_SECRET || "";
  const redirectUri =
    process.env.BLING_REDIRECT_URI ||
    "https://rockstar-producao.vercel.app/bling/callback";
  const appUrl =
    process.env.APP_URL || "https://rockstar-producao.vercel.app";
  const supabaseUrl =
    process.env.SUPABASE_URL ||
    process.env.VITE_SUPABASE_URL ||
    "https://hnjawylcwfhxjhkvtyyc.supabase.co";
  const supabaseServiceKey =
    process.env.SUPABASE_SERVICE_ROLE_KEY ||
    process.env.SUPABASE_SERVICE_KEY ||
    "";
  return {
    clientId,
    clientSecret,
    redirectUri,
    appUrl,
    supabaseUrl,
    supabaseServiceKey,
  };
}

export function basicAuthHeader(clientId, clientSecret) {
  const raw = Buffer.from(`${clientId}:${clientSecret}`).toString("base64");
  return `Basic ${raw}`;
}

export function normalizeDepositName(name) {
  return String(name || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

export function matchDepositos(depositos) {
  const list = Array.isArray(depositos) ? depositos : [];
  let depositoPaId = null;
  let depositoEstId = null;
  let depositoPaNome = null;
  let depositoEstNome = null;

  for (const dep of list) {
    const nome = normalizeDepositName(dep?.descricao || dep?.nome || dep?.name);
    const id = dep?.id ?? null;
    if (id == null) continue;
    const compact = nome.replace(/\s+/g, "");
    if (
      !depositoPaId &&
      (DEPOSITO_PA_NAMES.includes(nome) || DEPOSITO_PA_NAMES.includes(compact) || nome.includes("PRODUTO ACABADO"))
    ) {
      depositoPaId = id;
      depositoPaNome = dep?.descricao || dep?.nome || nome;
    }
    if (
      !depositoEstId &&
      (DEPOSITO_EST_NAMES.includes(nome) || nome.includes("OVERLOQUE"))
    ) {
      depositoEstId = id;
      depositoEstNome = dep?.descricao || dep?.nome || nome;
    }
  }

  return { depositoPaId, depositoEstId, depositoPaNome, depositoEstNome };
}

export function createServiceClient() {
  const { supabaseUrl, supabaseServiceKey } = getEnv();
  if (!supabaseUrl || !supabaseServiceKey) {
    throw new Error("SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY não configurados");
  }
  return createClient(supabaseUrl, supabaseServiceKey, {
    auth: { persistSession: false, autoRefreshToken: false },
  });
}

export async function exchangeToken({ grantType, code, refreshToken }) {
  const { clientId, clientSecret, redirectUri } = getEnv();
  if (!clientId || !clientSecret) {
    throw new Error("BLING_CLIENT_ID / BLING_CLIENT_SECRET não configurados");
  }

  const body = new URLSearchParams();
  body.set("grant_type", grantType);
  if (grantType === "authorization_code") {
    body.set("code", code);
    body.set("redirect_uri", redirectUri);
  } else if (grantType === "refresh_token") {
    body.set("refresh_token", refreshToken);
  } else {
    throw new Error(`grant_type inválido: ${grantType}`);
  }

  const res = await fetch(BLING_TOKEN_URL, {
    method: "POST",
    headers: {
      Authorization: basicAuthHeader(clientId, clientSecret),
      "Content-Type": "application/x-www-form-urlencoded",
      Accept: "application/json",
      "enable-jwt": "1",
    },
    body,
  });

  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg =
      json?.error?.description ||
      json?.error?.message ||
      json?.message ||
      JSON.stringify(json);
    throw new Error(`Falha ao obter token Bling (${res.status}): ${msg}`);
  }
  return json;
}

export async function blingFetch(path, accessToken, options = {}) {
  const url = path.startsWith("http") ? path : `${BLING_API_BASE}${path}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      "enable-jwt": "1",
      ...(options.headers || {}),
    },
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg =
      json?.error?.description ||
      json?.error?.message ||
      json?.message ||
      JSON.stringify(json);
    const err = new Error(`Bling API ${res.status}: ${msg}`);
    err.status = res.status;
    throw err;
  }
  return json;
}

export async function listDepositos(accessToken) {
  const json = await blingFetch("/depositos?pagina=1&limite=100", accessToken);
  return json?.data || json || [];
}

export async function upsertConexao(row) {
  const supabase = createServiceClient();
  const payload = {
    id: 1,
    ...row,
    updated_at: new Date().toISOString(),
  };
  const { data, error } = await supabase
    .from("bling_conexao")
    .upsert(payload, { onConflict: "id" })
    .select("*")
    .single();
  if (error) throw new Error(`Supabase bling_conexao: ${error.message}`);
  return data;
}

export async function getConexao() {
  const supabase = createServiceClient();
  const { data, error } = await supabase
    .from("bling_conexao")
    .select("*")
    .eq("id", 1)
    .maybeSingle();
  if (error) throw new Error(`Supabase bling_conexao: ${error.message}`);
  return data;
}

export function expiresAtFromToken(tokenJson) {
  const seconds = Number(tokenJson?.expires_in) || 3600;
  return new Date(Date.now() + seconds * 1000).toISOString();
}

export function redirect(res, url) {
  res.statusCode = 302;
  res.setHeader("Location", url);
  res.end();
}

export function json(res, status, body) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(body));
}

export function createState() {
  const nonce = Buffer.from(
    `${Date.now()}-${Math.random().toString(36).slice(2)}`
  ).toString("base64url");
  return nonce;
}
