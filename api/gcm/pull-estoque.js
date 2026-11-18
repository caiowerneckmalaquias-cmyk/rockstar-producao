const MPLUS_BASE = "https://nolanapi.mplus.net.br";
const SIZES = [34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44];
const FILIAL_PA = 100;
const FILIAL_EST = 200;
const DATA_INICIO = "2010-01-01";
const PAGE_LIMIT = 100;
const LOTE = 8;
const MAX_PAGINAS = 800;

export const maxDuration = 60;

function normalizeKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "")
    .trim()
    .toUpperCase();
}

function normalizeSigla(sigla) {
  return String(sigla || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function emptyGrid() {
  return Object.fromEntries(SIZES.map((size) => [size, 0]));
}

function piso(value) {
  const n = Math.trunc(Number(value) || 0);
  return n > 0 ? n : 0;
}

function hojeIso() {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "America/Sao_Paulo" }).format(new Date());
}

function catalogoIndex(catalogo) {
  const porSigla = new Map();
  const porNome = new Map();
  (Array.isArray(catalogo) ? catalogo : []).forEach((item) => {
    const ref = normalizeKey(item?.ref);
    const cor = String(item?.cor || "").trim();
    if (!ref || !cor) return;
    const sigla = normalizeSigla(item?.sigla);
    if (sigla) porSigla.set(`${ref}__${sigla}`, cor);
    porNome.set(`${ref}__${normalizeKey(cor)}`, cor);
  });
  return { porSigla, porNome };
}

function corDoCadastro(index, ref, codigoCor) {
  const refKey = normalizeKey(ref);
  const sigla = normalizeSigla(codigoCor);
  if (sigla) {
    const pelaSigla = index.porSigla.get(`${refKey}__${sigla}`);
    if (pelaSigla) return pelaSigla;
  }
  const peloNome = index.porNome.get(`${refKey}__${normalizeKey(codigoCor)}`);
  if (peloNome) return peloNome;
  return sigla || String(codigoCor || "").trim().toUpperCase();
}

function cursorValido(cursor) {
  const codigo = Number(cursor?.codigo);
  if (!Number.isFinite(codigo) || codigo <= 0) return null;
  const data = String(cursor?.data || "").slice(0, 10);
  return { codigo, data: /^\d{4}-\d{2}-\d{2}$/.test(data) ? data : DATA_INICIO };
}

function gradeDe(origem) {
  const grid = emptyGrid();
  SIZES.forEach((size) => {
    grid[size] = piso(origem?.[size]);
  });
  return grid;
}

function baseIndex(base) {
  const map = new Map();
  (Array.isArray(base) ? base : []).forEach((item) => {
    const ref = normalizeKey(item?.ref);
    const cor = normalizeKey(item?.cor);
    if (!ref || !cor) return;
    map.set(`${ref}__${cor}`, item);
  });
  return map;
}

function maiorCursor(atual, mov) {
  const codigo = Number(mov?.codigo) || 0;
  if (!codigo) return atual;
  const data = String(mov?.data || "").slice(0, 10) || atual?.data || DATA_INICIO;
  if (!atual || codigo > Number(atual.codigo || 0)) return { codigo, data };
  return atual;
}

export function aplicarMovimentos({ movimentos, catalogo, base, modo, cursor }) {
  const alvo = modo === "pa" || modo === "est" ? modo : "ambos";
  const querPa = alvo !== "est";
  const querEst = alvo !== "pa";
  const cursorPa = querPa ? cursorValido(cursor?.pa) : null;
  const cursorEst = querEst ? cursorValido(cursor?.est) : null;
  const absolutoPa = querPa && !cursorPa;
  const absolutoEst = querEst && !cursorEst;
  const index = catalogoIndex(catalogo);
  const bases = baseIndex(base);
  const porCor = new Map();
  let cursorPaNovo = cursorPa;
  let cursorEstNovo = cursorEst;

  (movimentos || []).forEach((mov) => {
    const filial = Number(mov?.filial);
    const size = Number(mov?.numeracao);
    if (!SIZES.includes(size)) return;
    const es = String(mov?.entrada_saida || "").trim().toUpperCase();
    if (es !== "E" && es !== "S") return;
    const qtd = piso(mov?.quantidade);
    if (!qtd) return;

    const afetaPa = querPa && filial !== FILIAL_EST && (filial === FILIAL_PA || es === "S");
    const afetaEst = querEst && filial === FILIAL_EST;
    if (afetaPa) cursorPaNovo = maiorCursor(cursorPaNovo, mov);
    if (afetaEst) cursorEstNovo = maiorCursor(cursorEstNovo, mov);

    const novoPa = afetaPa && (absolutoPa || Number(mov?.codigo) > cursorPa.codigo);
    const novoEst = afetaEst && (absolutoEst || Number(mov?.codigo) > cursorEst.codigo);
    if (!novoPa && !novoEst) return;

    const ref = String(mov?.produto || "").trim().toUpperCase();
    if (!ref) return;
    const cor = corDoCadastro(index, ref, mov?.padronizacao);
    const key = `${normalizeKey(ref)}__${normalizeKey(cor)}`;
    if (!porCor.has(key)) {
      const salvo = bases.get(key);
      porCor.set(key, {
        ref,
        cor,
        data: absolutoPa ? emptyGrid() : gradeDe(salvo?.pa),
        dataEst: absolutoEst ? emptyGrid() : gradeDe(salvo?.est),
        tocouPa: false,
        tocouEst: false,
      });
    }
    const item = porCor.get(key);
    const sinal = es === "E" ? 1 : -1;
    if (novoPa) {
      if (filial === FILIAL_PA) item.data[size] += sinal * qtd;
      else item.data[size] -= qtd;
      item.tocouPa = true;
    }
    if (novoEst) {
      item.dataEst[size] += sinal * qtd;
      item.tocouEst = true;
    }
  });

  const itens = [...porCor.values()]
    .filter((item) => item.tocouPa || item.tocouEst)
    .map((item) => {
      if (item.tocouPa) SIZES.forEach((size) => {
        item.data[size] = piso(item.data[size]);
      });
      if (item.tocouEst) SIZES.forEach((size) => {
        item.dataEst[size] = piso(item.dataEst[size]);
      });
      const total = item.tocouPa
        ? SIZES.reduce((acc, size) => acc + item.data[size], 0)
        : 0;
      const totalEst = item.tocouEst
        ? SIZES.reduce((acc, size) => acc + item.dataEst[size], 0)
        : 0;
      return {
        ref: item.ref,
        cor: item.cor,
        data: item.data,
        dataEst: item.dataEst,
        atualizaPa: item.tocouPa,
        temOverloque: item.tocouEst,
        total,
        totalEst,
      };
    });
  itens.sort((a, b) => a.ref.localeCompare(b.ref) || a.cor.localeCompare(b.cor));
  const totalPares = itens.reduce((acc, item) => acc + item.total, 0);
  const totalEst = itens.reduce((acc, item) => acc + item.totalEst, 0);
  return {
    itens,
    totalPares,
    totalEst,
    modo: alvo,
    historico: { pa: absolutoPa, est: absolutoEst },
    cursor: {
      pa: querPa ? cursorPaNovo : null,
      est: querEst ? cursorEstNovo : null,
    },
  };
}

async function mplusJson(path, { method = "GET", token, credential } = {}) {
  const headers = { Accept: "application/json" };
  if (credential) headers.Authorization = credential;
  if (token) headers["x-access-token"] = token;
  const res = await fetch(`${MPLUS_BASE}${path}`, { method, headers });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const error = new Error(data.error || data.message || `MPlus HTTP ${res.status}`);
    error.statusCode = res.status;
    throw error;
  }
  return data;
}

async function listarMovimentos(token, { filial, dataInicial }) {
  const params = new URLSearchParams({
    page: "1",
    limit: String(PAGE_LIMIT),
    data_inicial: dataInicial || DATA_INICIO,
    data_final: hojeIso(),
  });
  if (filial) params.set("filial", String(filial));
  const first = await mplusJson(
    `/api/integracao/estoque/produto_acabado_movimento?${params}`,
    { token }
  );
  const totalPages = Math.min(MAX_PAGINAS, Math.max(1, Number(first.totalPages) || 1));
  const movimentos = [...(first.data || [])];
  const paginas = [];
  for (let page = 2; page <= totalPages; page += 1) paginas.push(page);
  for (let i = 0; i < paginas.length; i += LOTE) {
    const lote = paginas.slice(i, i + LOTE);
    const bodies = await Promise.all(
      lote.map((page) => {
        const query = new URLSearchParams(params);
        query.set("page", String(page));
        return mplusJson(
          `/api/integracao/estoque/produto_acabado_movimento?${query}`,
          { token }
        );
      })
    );
    bodies.forEach((body) => movimentos.push(...(body.data || [])));
  }
  return movimentos;
}

function dataInicial(modo, cursor) {
  const datas = [];
  if (modo !== "est") {
    const pa = cursorValido(cursor?.pa);
    if (!pa) return DATA_INICIO;
    datas.push(pa.data);
  }
  if (modo !== "pa") {
    const est = cursorValido(cursor?.est);
    if (!est) return DATA_INICIO;
    datas.push(est.data);
  }
  return datas.sort()[0] || DATA_INICIO;
}

export async function pullEstoqueGcm(entrada) {
  const body = entrada && !Array.isArray(entrada) && typeof entrada === "object" && ("modo" in entrada || "cursor" in entrada || "base" in entrada)
    ? entrada
    : { catalogo: entrada };
  const modo = body.modo === "pa" || body.modo === "est" ? body.modo : "ambos";
  const credential = String(process.env.MPLUS_INTEGRATION_KEY || "").trim();
  if (!credential) {
    const error = new Error("MPLUS_INTEGRATION_KEY não configurada.");
    error.statusCode = 500;
    throw error;
  }

  const { token } = await mplusJson("/api/integracao/auth/token", {
    method: "POST",
    credential,
  });
  if (!token) {
    const error = new Error("A MPlus não devolveu o token de acesso.");
    error.statusCode = 502;
    throw error;
  }

  const movimentos = await listarMovimentos(token, {
    filial: modo === "est" ? FILIAL_EST : "",
    dataInicial: dataInicial(modo, body.cursor),
  });
  return aplicarMovimentos({
    movimentos,
    catalogo: body.catalogo,
    base: body.base,
    modo,
    cursor: body.cursor,
  });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }
  try {
    const body = req.body && typeof req.body === "object" ? req.body : {};
    const data = await pullEstoqueGcm(body);
    res.status(200).json(data);
  } catch (error) {
    console.error("gcm pull-estoque", error.message);
    res.status(error.statusCode || 500).json({
      error: error.message || "Falha ao puxar estoque do GCM",
    });
  }
}
