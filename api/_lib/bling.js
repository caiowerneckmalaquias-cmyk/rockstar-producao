import { createClient } from "@supabase/supabase-js";

export const BLING_AUTHORIZE_URL = "https://www.bling.com.br/Api/v3/oauth/authorize";
export const BLING_TOKEN_URL = "https://www.bling.com.br/Api/v3/oauth/token";
export const BLING_API_BASE = "https://api.bling.com.br/Api/v3";

export const STOCK_OFFSET = Number(process.env.BLING_STOCK_OFFSET) || 1000;

const DEPOSITO_PA_NAMES = ["PRODUTO ACABADO", "PRODUTOACABADO"];
const DEPOSITO_EST_NAMES = ["OVERLOQUE"];
const DEPOSITO_GERAL_NAMES = ["GERAL"];
export const BLING_SECUNDARIA_SLOTS = [1, 2, 3, 4];

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

/** Nome bruto do depósito em vários formatos da API Bling. */
export function depositRawName(dep) {
  if (!dep || typeof dep !== "object") return "";
  return (
    dep.descricao ||
    dep.nome ||
    dep.name ||
    dep.descricaoDeposito ||
    dep.deposito?.descricao ||
    dep.deposito?.nome ||
    dep.deposito?.name ||
    ""
  );
}

export function depositId(dep) {
  if (!dep || typeof dep !== "object") return null;
  const id = dep.id ?? dep.deposito?.id ?? null;
  return id == null ? null : id;
}

/** Multiempresa / Multiempresas / Multi Empresa (com ou sem espaços). */
export function isMultiempresaName(rawNome) {
  const nome = normalizeDepositName(rawNome);
  if (!nome) return false;
  const compact = nome.replace(/\s+/g, "");
  return (
    compact.includes("MULTIEMPRESA") ||
    nome.includes("MULTI EMPRESA") ||
    /\bMULTI\s*EMPRESAS?\b/.test(nome)
  );
}

export function matchDepositos(depositos) {
  const list = Array.isArray(depositos) ? depositos : [];
  let depositoPaId = null;
  let depositoEstId = null;
  let depositoMultiId = null;
  let depositoPaNome = null;
  let depositoEstNome = null;
  let depositoMultiNome = null;

  for (const dep of list) {
    const rawNome = depositRawName(dep);
    const nome = normalizeDepositName(rawNome);
    const id = depositId(dep);
    if (id == null) continue;
    const compact = nome.replace(/\s+/g, "");
    if (!depositoMultiId && isMultiempresaName(rawNome)) {
      depositoMultiId = id;
      depositoMultiNome = rawNome || nome;
    }
    if (
      !depositoPaId &&
      (DEPOSITO_PA_NAMES.includes(nome) ||
        DEPOSITO_PA_NAMES.includes(compact) ||
        nome.includes("PRODUTO ACABADO"))
    ) {
      depositoPaId = id;
      depositoPaNome = rawNome || nome;
    }
    if (
      !depositoEstId &&
      (DEPOSITO_EST_NAMES.includes(nome) || nome.includes("OVERLOQUE"))
    ) {
      depositoEstId = id;
      depositoEstNome = rawNome || nome;
    }
  }

  return {
    depositoPaId,
    depositoEstId,
    depositoMultiId,
    depositoPaNome,
    depositoEstNome,
    depositoMultiNome,
  };
}

/** Casa depósito Geral (vendas das contas secundárias). */
export function matchDepositoGeral(depositos) {
  const list = Array.isArray(depositos) ? depositos : [];
  let exact = null;
  let fuzzy = null;
  for (const dep of list) {
    const rawNome = depositRawName(dep);
    const nome = normalizeDepositName(rawNome);
    const id = depositId(dep);
    if (id == null || !nome) continue;
    if (isMultiempresaName(rawNome)) continue;
    const compact = nome.replace(/\s+/g, "");
    if (DEPOSITO_GERAL_NAMES.includes(nome) || DEPOSITO_GERAL_NAMES.includes(compact)) {
      exact = { id, nome: rawNome || nome };
      break;
    }
    if (!fuzzy && (nome === "GERAL" || compact === "GERAL" || nome.startsWith("GERAL "))) {
      fuzzy = { id, nome: rawNome || nome };
    }
  }
  const hit = exact || fuzzy;
  return {
    depositoGeralId: hit?.id ?? null,
    depositoGeralNome: hit?.nome ?? null,
  };
}

/** Lista resumida para debug na UI. */
export function summarizeDepositos(depositos) {
  const list = Array.isArray(depositos) ? depositos : [];
  return list
    .map((dep) => {
      const id = depositId(dep);
      const nome = depositRawName(dep) || "(sem nome)";
      if (id == null) return null;
      return { id, nome: String(nome) };
    })
    .filter(Boolean);
}

/**
 * Descobre Multiempresa a partir do payload de saldos do produto
 * (quando GET /depositos não lista o virtual).
 */
export function discoverMultiempresaFromSaldos(saldosPayload) {
  const items = Array.isArray(saldosPayload?.data)
    ? saldosPayload.data
    : Array.isArray(saldosPayload)
      ? saldosPayload
      : [];

  for (const item of items) {
    const deps = item?.depositos || item?.saldosDepositos || [];
    if (Array.isArray(deps)) {
      for (const d of deps) {
        const rawNome = depositRawName(d);
        const id = depositId(d);
        if (id != null && isMultiempresaName(rawNome)) {
          return { id, nome: rawNome || "Multiempresa" };
        }
      }
    }
    const nestedNome = depositRawName(item?.deposito || item);
    const nestedId = depositId(item?.deposito || item);
    if (nestedId != null && isMultiempresaName(nestedNome)) {
      return { id: nestedId, nome: nestedNome || "Multiempresa" };
    }
  }
  return null;
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
  const raw = json?.data ?? json;
  if (Array.isArray(raw)) return raw;
  if (Array.isArray(raw?.data)) return raw.data;
  return [];
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

export function normalizeSecundariaSlot(slot) {
  const n = Number(slot);
  if (!BLING_SECUNDARIA_SLOTS.includes(n)) return null;
  return n;
}

export async function listContasSecundarias() {
  const supabase = createServiceClient();
  const { data, error } = await supabase
    .from("bling_conta_secundaria")
    .select("*")
    .order("slot", { ascending: true });
  if (error) throw new Error(`Supabase bling_conta_secundaria: ${error.message}`);
  return data || [];
}

export async function getContaSecundaria(slot) {
  const s = normalizeSecundariaSlot(slot);
  if (s == null) return null;
  const supabase = createServiceClient();
  const { data, error } = await supabase
    .from("bling_conta_secundaria")
    .select("*")
    .eq("slot", s)
    .maybeSingle();
  if (error) throw new Error(`Supabase bling_conta_secundaria: ${error.message}`);
  return data;
}

export async function upsertContaSecundaria(row) {
  const s = normalizeSecundariaSlot(row?.slot);
  if (s == null) throw new Error("Slot secundário inválido (use 1–4)");
  const supabase = createServiceClient();
  const payload = {
    ...row,
    slot: s,
    label: row.label || `Conta secundária ${s}`,
    updated_at: new Date().toISOString(),
  };
  const { data, error } = await supabase
    .from("bling_conta_secundaria")
    .upsert(payload, { onConflict: "slot" })
    .select("*")
    .single();
  if (error) throw new Error(`Supabase bling_conta_secundaria: ${error.message}`);
  return data;
}

/** Token válido de uma conta secundária (renova se precisar). */
export async function getValidAccessTokenSecundaria(slot) {
  const row = await getContaSecundaria(slot);
  if (!row?.access_token) {
    throw new Error(`Conta secundária ${slot} não conectada.`);
  }

  const expiresAt = row.expires_at ? new Date(row.expires_at).getTime() : 0;
  const skewMs = 60_000;
  if (expiresAt && expiresAt - skewMs > Date.now()) {
    return { accessToken: row.access_token, conta: row };
  }

  if (!row.refresh_token) {
    throw new Error(`Token da conta secundária ${slot} expirado. Reconecte.`);
  }

  const tokenJson = await exchangeToken({
    grantType: "refresh_token",
    refreshToken: row.refresh_token,
  });

  const updated = await upsertContaSecundaria({
    slot: row.slot,
    label: row.label,
    access_token: tokenJson.access_token,
    refresh_token: tokenJson.refresh_token || row.refresh_token,
    token_type: tokenJson.token_type || "Bearer",
    expires_at: expiresAtFromToken(tokenJson),
    scopes: tokenJson.scope || row.scopes,
    deposito_geral_id: row.deposito_geral_id,
    deposito_geral_nome: row.deposito_geral_nome,
    connected_at: row.connected_at,
  });

  return { accessToken: updated.access_token, conta: updated };
}

/** Contas secundárias prontas para somar no DISP (token + Geral). */
export async function listContasSecundariasProntas() {
  const rows = await listContasSecundarias();
  const prontas = [];
  for (const row of rows) {
    if (!row?.access_token || row.deposito_geral_id == null) continue;
    try {
      const { accessToken, conta } = await getValidAccessTokenSecundaria(row.slot);
      if (conta.deposito_geral_id != null) {
        prontas.push({ accessToken, conta });
      }
    } catch (e) {
      console.warn(`secundaria slot ${row.slot}:`, e.message);
    }
  }
  return prontas;
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

export function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function blingToReal(blingQty, offset = STOCK_OFFSET) {
  const n = Number(blingQty);
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, n - (Number(offset) || 0));
}

export function realToBling(realQty, offset = STOCK_OFFSET) {
  const n = Number(realQty);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return n + (Number(offset) || 0);
}

export function shouldZeroBling(blingQty, offset = STOCK_OFFSET) {
  const n = Number(blingQty);
  if (!Number.isFinite(n)) return true;
  return n <= (Number(offset) || 0);
}

export function buildBlingSku(ref, sigla, size) {
  const r = String(ref || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
  const s = String(sigla || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
  const n = String(size ?? "").trim();
  if (!r || !s || !n) return "";
  return `${r}${s}${n}`;
}

/** Access token válido; renova com refresh se perto de expirar. */
export async function getValidAccessToken() {
  const row = await getConexao();
  if (!row?.access_token) {
    throw new Error("Bling não conectado. Use Conectar Bling primeiro.");
  }

  const expiresAt = row.expires_at ? new Date(row.expires_at).getTime() : 0;
  const skewMs = 60_000;
  if (expiresAt && expiresAt - skewMs > Date.now()) {
    return { accessToken: row.access_token, conexao: row };
  }

  if (!row.refresh_token) {
    throw new Error("Token Bling expirado. Reconecte o Bling.");
  }

  const tokenJson = await exchangeToken({
    grantType: "refresh_token",
    refreshToken: row.refresh_token,
  });

  const updated = await upsertConexao({
    access_token: tokenJson.access_token,
    refresh_token: tokenJson.refresh_token || row.refresh_token,
    token_type: tokenJson.token_type || "Bearer",
    expires_at: expiresAtFromToken(tokenJson),
    scopes: tokenJson.scope || row.scopes,
    deposito_pa_id: row.deposito_pa_id,
    deposito_est_id: row.deposito_est_id,
    deposito_multi_id: row.deposito_multi_id,
    deposito_pa_nome: row.deposito_pa_nome,
    deposito_est_nome: row.deposito_est_nome,
    deposito_multi_nome: row.deposito_multi_nome,
    connected_at: row.connected_at,
  });

  return { accessToken: updated.access_token, conexao: updated };
}

export async function findProductByCodigo(accessToken, codigo) {
  const code = encodeURIComponent(String(codigo || "").trim());
  if (!code) return null;
  const json = await blingFetch(
    `/produtos?codigo=${code}&pagina=1&limite=10`,
    accessToken
  );
  const list = json?.data || [];
  if (!Array.isArray(list) || !list.length) return null;
  const exact = list.find(
    (p) =>
      String(p.codigo || "").toUpperCase() === String(codigo).toUpperCase()
  );
  return exact || list[0];
}

/** Saldo do produto em um depósito (0 se não achar). */
export function saldoNoDeposito(saldosPayload, depositoId) {
  const depId = Number(depositoId);
  const items = Array.isArray(saldosPayload?.data)
    ? saldosPayload.data
    : Array.isArray(saldosPayload)
      ? saldosPayload
      : [];

  for (const item of items) {
    const deps = item?.depositos || item?.saldosDepositos || [];
    if (Array.isArray(deps)) {
      for (const d of deps) {
        if (Number(d?.id || d?.deposito?.id) === depId) {
          const saldo =
            d?.saldo ??
            d?.saldoVirtual ??
            d?.saldoFisico ??
            d?.quantidade ??
            0;
          return Number(saldo) || 0;
        }
      }
    }
    if (Number(item?.deposito?.id) === depId) {
      return Number(item?.saldo ?? item?.quantidade ?? 0) || 0;
    }
  }

  // Fallback: saldo geral do primeiro item
  if (items[0] && depositoId == null) {
    return Number(items[0].saldo ?? items[0].saldoVirtual ?? 0) || 0;
  }
  return 0;
}

export async function fetchSaldosProduto(accessToken, produtoId) {
  const json = await blingFetch(
    `/estoques/saldos?idsProdutos[]=${encodeURIComponent(produtoId)}`,
    accessToken
  );
  return json;
}

export async function setEstoqueSaldo({
  accessToken,
  produtoId,
  depositoId,
  quantidade,
  observacoes,
}) {
  return blingFetch("/estoques", accessToken, {
    method: "POST",
    body: JSON.stringify({
      produto: { id: Number(produtoId) },
      deposito: { id: Number(depositoId) },
      operacao: "B",
      quantidade: Number(quantidade) || 0,
      observacoes: observacoes || "Rock Star Produção",
    }),
  });
}
