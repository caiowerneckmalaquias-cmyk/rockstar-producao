import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import { supabase } from "./supabase";

import {
  sizes,
  navItems,
  initialRows,
  initialMinimos,
  initialVendas,
  initialTempoProducao,
  LIMITE_PROGRAMACAO_DIA
} from "./constants/production";

const makeEmptyGrid = () => Object.fromEntries(sizes.map((s) => [s, 0]));

function deepClonePlain(obj) {
  try {
    if (typeof structuredClone === "function") return structuredClone(obj);
  } catch (_) {
    /* ignore */
  }
  return JSON.parse(JSON.stringify(obj));
}

const PROG_FICHAS_LANCADAS_STORAGE_KEY = "rockstar-prog-fichas-lancadas-v1";

function readProgFichasLancadasFromStorage() {
  try {
    const raw = localStorage.getItem(PROG_FICHAS_LANCADAS_STORAGE_KEY);
    if (!raw) return [];
    const p = JSON.parse(raw);
    return Array.isArray(p) ? p.filter((x) => typeof x === "string") : [];
  } catch {
    return [];
  }
}

const PROG_RESERVA_TOP_PCT_KEY = "rockstar-prog-reserva-top-pct-v1";
const PROG_TOP_N_KEY = "rockstar-prog-top-n-v1";
const PROG_TOP_MODE_KEY = "rockstar-prog-top-mode-v1";
const PROG_TOP_MANUAL_KEYS_KEY = "rockstar-prog-top-manual-keys-v1";
const PROG_VALORES_PAGAMENTO_KEY = "rockstar-prog-valores-pagamento-v1";

function readProgReservaTopPctFromStorage() {
  try {
    const raw = localStorage.getItem(PROG_RESERVA_TOP_PCT_KEY);
    if (raw == null || raw === "") return 0;
    const n = Number(raw);
    if (Number.isNaN(n)) return 0;
    return Math.min(100, Math.max(0, n));
  } catch {
    return 0;
  }
}

function readProgTopNFromStorage() {
  try {
    const raw = localStorage.getItem(PROG_TOP_N_KEY);
    if (raw == null || raw === "") return 10;
    const n = Number(raw);
    if (Number.isNaN(n)) return 10;
    return Math.min(200, Math.max(1, Math.round(n)));
  } catch {
    return 10;
  }
}

function readProgTopModeFromStorage() {
  try {
    const raw = String(localStorage.getItem(PROG_TOP_MODE_KEY) || "").trim().toLowerCase();
    return raw === "manual" ? "manual" : "auto";
  } catch {
    return "auto";
  }
}

function readProgTopManualKeysFromStorage() {
  try {
    const raw = localStorage.getItem(PROG_TOP_MANUAL_KEYS_KEY);
    if (!raw) return [];
    const p = JSON.parse(raw);
    return Array.isArray(p) ? p.filter((x) => typeof x === "string" && x.includes("__")) : [];
  } catch {
    return [];
  }
}

function readProgValoresPagamentoFromStorage() {
  try {
    const raw = localStorage.getItem(PROG_VALORES_PAGAMENTO_KEY);
    if (!raw) return { weverton: "", romulo: "" };
    const p = JSON.parse(raw);
    return {
      weverton: String(p?.weverton ?? ""),
      romulo: String(p?.romulo ?? ""),
    };
  } catch {
    return { weverton: "", romulo: "" };
  }
}

function parseDecimalInput(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return Number.NaN;
  const normalized = raw.replace(/\s+/g, "").replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function parseTopManualKeysFromDb(raw) {
  try {
    const base = typeof raw === "string" ? JSON.parse(raw) : raw;
    return Array.isArray(base) ? base.filter((x) => typeof x === "string" && x.includes("__")) : [];
  } catch {
    return [];
  }
}

let pdfCaptureProbeCtx = null;
function getPdfCaptureProbeCtx() {
  if (!pdfCaptureProbeCtx) {
    const c = document.createElement("canvas");
    c.width = 1;
    c.height = 1;
    pdfCaptureProbeCtx = c.getContext("2d");
  }
  return pdfCaptureProbeCtx;
}

/** Converte um token de cor (oklch/lab/etc.) para hex/rgb via Canvas — formato que html2canvas entende. */
function coerceColorTokenForHtml2Canvas(token) {
  const t = String(token || "").trim();
  if (!t) return "#808080";
  const ctx = getPdfCaptureProbeCtx();
  try {
    ctx.fillStyle = "#000000";
    ctx.fillStyle = t;
    const out = String(ctx.fillStyle);
    if (out && !/oklch|lab\(|lch\(|color\(/i.test(out)) return out;
  } catch {
    /* ignore */
  }
  return "#808080";
}

/**
 * Browsers podem devolver getComputedStyle já em oklch(); html2canvas não parseia.
 * Normaliza tokens modernos e, em último caso, remove oklch/color-mix restantes.
 */
function sanitizeCssValueForHtml2Canvas(value) {
  if (typeof value !== "string" || !value.trim()) return value;
  if (!/oklch|lab\(|lch\(|color\(|color-mix\(/i.test(value)) return value;

  const ctx = getPdfCaptureProbeCtx();
  try {
    ctx.fillStyle = "#000000";
    ctx.fillStyle = value.trim();
    const asFill = String(ctx.fillStyle);
    if (asFill && !/oklch|lab\(|lch\(|color\(/i.test(asFill)) return asFill;
  } catch {
    /* continuar com substituições */
  }

  let out = value
    .replace(/\boklch\([^)]+\)/gi, (m) => coerceColorTokenForHtml2Canvas(m))
    .replace(/\blab\([^)]+\)/gi, (m) => coerceColorTokenForHtml2Canvas(m))
    .replace(/\blch\([^)]+\)/gi, (m) => coerceColorTokenForHtml2Canvas(m))
    .replace(/\bcolor\([^)]+\)/gi, (m) => coerceColorTokenForHtml2Canvas(m));

  if (/oklch|color-mix\(/i.test(out)) {
    out = out
      .replace(/\boklch\([^)]*\)/gi, "#808080")
      .replace(/\bcolor-mix\([^)]*\)/gi, "#808080");
  }
  return out;
}

/**
 * html2canvas não interpreta cores oklch() (Tailwind v4). Copiamos o estilo *calculado*
 * para inline no clone (sanitizado), removemos classes e depois removemos folhas no onclone.
 */
function syncInlineComputedStylesForPdfCapture(origRoot, cloneRoot) {
  if (
    origRoot?.nodeType !== Node.ELEMENT_NODE ||
    cloneRoot?.nodeType !== Node.ELEMENT_NODE
  ) {
    return;
  }
  const win = origRoot.ownerDocument?.defaultView;
  if (!win) return;

  try {
    const cs = win.getComputedStyle(origRoot);
    for (let i = 0; i < cs.length; i++) {
      const name = cs[i];
      try {
        let value = cs.getPropertyValue(name);
        const priority = cs.getPropertyPriority(name);
        if (value) {
          if (/oklch|lab\(|lch\(|color\(|color-mix\(/i.test(value)) {
            value = sanitizeCssValueForHtml2Canvas(value);
          }
          cloneRoot.style.setProperty(name, value, priority);
        }
      } catch {
        /* propriedade pode não aplicar-se ao nó clonado */
      }
    }
  } catch {
    /* ignore */
  }
  cloneRoot.removeAttribute("class");

  const oCh = origRoot.children;
  const cCh = cloneRoot.children;
  const len = Math.min(oCh.length, cCh.length);
  for (let i = 0; i < len; i++) {
    syncInlineComputedStylesForPdfCapture(oCh[i], cCh[i]);
  }
}

/** PDF a partir do mesmo DOM da folha de impressão (html2canvas + A4 com margens iguais ao @page). */
async function buildProgramacaoPrintSheetPdfBlobFromElement(element) {
  const canvas = await html2canvas(element, {
    scale: 2,
    useCORS: true,
    logging: false,
    backgroundColor: "#ffffff",
    scrollX: 0,
    scrollY: 0,
    width: element.scrollWidth,
    height: element.scrollHeight,
    onclone(clonedDoc) {
      const clonedRoot = clonedDoc.getElementById("programacao-print-sheet-root");
      if (clonedRoot) syncInlineComputedStylesForPdfCapture(element, clonedRoot);
      clonedDoc.querySelectorAll('link[rel="stylesheet"]').forEach((n) => n.remove());
      clonedDoc.querySelectorAll("style").forEach((n) => n.remove());
    },
  });
  const imgData = canvas.toDataURL("image/jpeg", 0.93);
  const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait", compress: true });
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = 10;
  const contentW = pageW - 2 * margin;
  const contentH = pageH - 2 * margin;
  const imgW = contentW;
  const imgH = (canvas.height * imgW) / canvas.width;
  let page = 0;
  while (page * contentH < imgH - 0.05) {
    if (page > 0) pdf.addPage();
    pdf.addImage(imgData, "JPEG", margin, margin - page * contentH, imgW, imgH);
    page += 1;
  }
  return pdf.output("blob");
}

/** Garante chaves numéricas por numeração (evita mismatch "34" vs 34 vindo do Supabase/JSON). */
const normalizeProductData = (rawData) =>
  Object.fromEntries(
    sizes.map((s) => {
      const cell = rawData?.[s] ?? rawData?.[String(s)];
      return [
        s,
        {
          pa: Number(cell?.pa) || 0,
          est: Number(cell?.est) || 0,
          m: Number(cell?.m) || 0,
          p: Number(cell?.p) || 0,
        },
      ];
    })
  );

const calcTotal = (item) => item.pa + item.est + item.m + item.p;
const round12 = (n) => (n <= 0 ? 0 : Math.ceil(n / 12) * 12);
const normalizeKey = (value) =>
  String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, " ")
    .trim()
    .toUpperCase();

function statusFor(item, minimo) {
  const prod = item.est + item.m + item.p;
  if (item.pa < minimo.pa && prod < minimo.prod) return "CRÍTICO";
  if (item.pa < minimo.pa) return "ATENÇÃO PA";
  if (prod < minimo.prod) return "ATENÇÃO PROD";
  return "OK";
}

function tone(status) {
  if (status === "CRÍTICO") return "bg-red-50";
  if (status === "ATENÇÃO PA") return "bg-amber-50";
  if (status === "ATENÇÃO PROD") return "bg-sky-50";
  return "bg-white";
}

function badge(status) {
  if (status === "CRÍTICO") return "bg-red-100 text-red-700 border-red-200";
  if (status === "ATENÇÃO PA") return "bg-amber-100 text-amber-700 border-amber-200";
  if (status === "ATENÇÃO PROD") return "bg-sky-100 text-[#8B1E2D] border-sky-200";
  return "bg-emerald-100 text-emerald-700 border-emerald-200";
}

function vendaDiaFromMes(vendaMes) {
  return (Number(vendaMes) || 0) / 30;
}

function coberturaDias(pa, vendaMes) {
  const vendaDia = vendaDiaFromMes(vendaMes);
  if (vendaDia <= 0) return null;
  return pa / vendaDia;
}

function coberturaBadgeClass(cobertura, tempoTotal) {
  if (cobertura == null) return "bg-slate-100 text-slate-700 border-slate-200";
  if (cobertura < tempoTotal) return "bg-red-100 text-red-700 border-red-200";
  if (cobertura <= tempoTotal + 2) return "bg-amber-100 text-amber-700 border-amber-200";
  return "bg-emerald-100 text-emerald-700 border-emerald-200";
}

function coberturaLabel(cobertura, tempoTotal) {
  if (cobertura == null) return "Sem giro";
  if (cobertura < tempoTotal) return "Cobertura crítica";
  if (cobertura <= tempoTotal + 2) return "Cobertura curta";
  return "Cobertura ok";
}

function parseDateBrToDate(value) {
  if (!value) return null;
  const parts = String(value).trim().split("/");
  if (parts.length !== 3) return null;
  const [dd, mm, yyyy] = parts;
  const date = new Date(`${yyyy}-${mm}-${dd}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDateToBr(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function parseFeriadosText(text) {
  return String(text || "")
    .split(/[,;]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map(parseDateBrToDate)
    .filter(Boolean)
    .map((date) => formatDateToBr(date));
}

function contarDiasUteisDoMesAteHoje(currentDate, feriadosList = []) {
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();
  const feriadosSet = new Set(feriadosList);
  let total = 0;

  for (let day = 1; day <= currentDate.getDate(); day += 1) {
    const date = new Date(year, month, day);
    const diaSemana = date.getDay();
    const isFimDeSemana = diaSemana === 0 || diaSemana === 6;
    const dataBr = formatDateToBr(date);
    if (!isFimDeSemana && !feriadosSet.has(dataBr)) {
      total += 1;
    }
  }

  return Math.max(1, total);
}

function parseGcmRawText(rawText) {
  const lines = String(rawText || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const resultado = [];
  let atual = null;
  let tamanhos = [];

  const extrairNumeros = (texto) =>
    (texto.match(/\d+/g) || []).map(Number);

  const extrairCor = (texto) => {
  const partes = texto.split("-");

  // pega a parte principal (onde tem ADULTO / COURINO / INFANTIL)
  const descricao = partes[1] || "";

  const palavras = descricao.trim().split(" ");

  // encontra onde começa a cor
  const idx = palavras.findIndex((p) =>
    ["ADULTO", "COURINO", "INFANTIL"].includes(p)
  );

  if (idx === -1) return "";

  // tudo depois disso é cor
  return palavras.slice(idx + 1).join(" ").trim();
};

  const finalizar = () => {
    if (!atual) return;
    atual.total = sizes.reduce((acc, s) => acc + (atual.data[s] || 0), 0);
    resultado.push(atual);
  };

  lines.forEach((line) => {
    const texto = line.toUpperCase();

    // ignora overloque
    if (texto.includes("OVERLOQUE")) return;

    // cabeçalho produto
    if (
      texto.includes("-") &&
      (texto.includes("ADULTO") || texto.includes("COURINO") || texto.includes("INFANTIL"))
    ) {
      finalizar();

      atual = {
        ref: texto.split("-")[0].trim(),
        cor: extrairCor(texto),
        data: Object.fromEntries(sizes.map((s) => [s, 0])),
        total: 0,
      };

      return;
    }

    if (!atual) return;

    // linha de tamanhos
    if (texto.includes("QTDE")) {
      tamanhos = extrairNumeros(texto).filter((n) => sizes.includes(n));
      return;
    }

    // linha de estoque
    if (texto.includes("ESTOQUE")) {
      const numeros = extrairNumeros(texto);

      tamanhos.forEach((size, idx) => {
        atual.data[size] = numeros[idx] || 0;
      });

      return;
    }
  });

  finalizar();

  return resultado;
}

function buildSuggestions(rows, minimos, vendas, tempoProducao) {
  const montagem = [];
  const pesponto = [];

  const diasPesponto = Number(tempoProducao?.pesponto) || 0;
  const diasMontagem = Number(tempoProducao?.montagem) || 0;
  const diasTotal = diasPesponto + diasMontagem;
  
    const LIMITE_POR_NUMERO = 84;

  const roundUp12 = (value) => {
    const numero = Number(value) || 0;
    if (numero <= 0) return 0;
    return Math.ceil(numero / 12) * 12;
  };

  /**
   * Mínimo múltiplo de 12 → sugere em múltiplos de 12.
   * Outro mínimo > 0 (ex.: 4) → arredonda para cima em múltiplos desse mínimo.
   * Sem mínimo de referência (0) → mantém a necessidade calculada.
   */
  const sugerirQtdPorMinimo = (necessidade, minReferencia) => {
    const n = Math.max(0, Number(necessidade) || 0);
    const minRef = Number(minReferencia) || 0;
    if (n <= 0) return 0;
    if (minRef > 0 && minRef % 12 === 0) {
      return roundUp12(n);
    }
    if (minRef > 0) {
      return Math.ceil(n / minRef) * minRef;
    }
    return n;
  };

  /** 12, 24, 84 → lote 12; caso contrário o próprio mínimo (ex.: 4). */
  const passoLotePorMinimo = (minRef) => {
    const m = Number(minRef) || 0;
    if (m <= 0) return 0;
    if (m % 12 === 0) return 12;
    return m;
  };

  const getVendaDia = (vendaMes) => (Number(vendaMes) || 0) / 30;

  const getCoberturaDias = (estoque, vendaDia) => {
    if (vendaDia <= 0) return estoque > 0 ? 999 : 0;
    return estoque / vendaDia;
  };

  rows.forEach((row) => {
    const mins = minimos?.[row.ref]?.[row.cor] || {};
    const sales = vendas?.[row.ref]?.[row.cor] || {};

    const montSizes = makeEmptyGrid();
    const pespSizes = makeEmptyGrid();

    let montTotal = 0;
    let pespTotal = 0;

    let prioridadeMontagem = 0;
    let prioridadePesponto = 0;

    let itemUrgenteMontagem = false;
    let itemUrgentePesponto = false;

    let itemRecuperacaoMontagem = false;
    let itemRecuperacaoPesponto = false;

    let tamanhosCriticosMontagem = 0;
    let tamanhosPendentesMontagem = 0;

    let tamanhosCriticosPesponto = 0;
    let tamanhosPendentesPesponto = 0;

    let vendaTotalRefCor = 0;

    sizes.forEach((size) => {
      const item = row?.data?.[size] || { pa: 0, est: 0, m: 0, p: 0 };
      const minimo = mins?.[size] || { pa: 0, prod: 0 };
      const vendaMes = Number(sales?.[size]) || 0;
      const vendaDia = getVendaDia(vendaMes);

      vendaTotalRefCor += vendaMes;

      const pa = Number(item?.pa) || 0;
      const est = Number(item?.est) || 0;
      const m = Number(item?.m) || 0;
      const p = Number(item?.p) || 0;

      const minPA = Number(minimo?.pa) || 0;
      const minProd = Number(minimo?.prod) || 0;

      const prodAtual = est + m + p;

      const itemRelevante =
        vendaMes > 0 ||
        minPA > 0 ||
        minProd > 0;

      if (!itemRelevante) {
        montSizes[size] = 0;
        pespSizes[size] = 0;
        return;
      }

      const temMinimo = minPA > 0 || minProd > 0;

      const consumoDuranteMontagem = Math.ceil(vendaDia * diasMontagem);
      const consumoDuranteCicloTotal = Math.ceil(vendaDia * diasTotal);

      const needPA = Math.max(0, minPA + consumoDuranteMontagem - pa);

      const needProd = Math.max(0, minProd + consumoDuranteCicloTotal - prodAtual);

      let mont = 0;
      let pesp = 0;

      if (temMinimo) {
        const refMont = minPA > 0 ? minPA : minProd > 0 ? minProd : 0;
        const refPesp = minProd > 0 ? minProd : minPA > 0 ? minPA : 0;

        const montDesejado = Math.min(sugerirQtdPorMinimo(needPA, refMont), LIMITE_POR_NUMERO);

        const passoMont = passoLotePorMinimo(refMont);
        mont =
          passoMont > 0
            ? Math.floor(Math.min(montDesejado, est) / passoMont) * passoMont
            : Math.min(montDesejado, est);

        const faltaBrutaMontagem = Math.max(0, montDesejado - mont);
        // Abate falta com pipeline acima do mínimo de prod; sem min prod, considera P (WIP pesponto).
        const creditoPipeline =
          minProd > 0 ? Math.max(0, prodAtual - minProd) : p;
        const faltaParaMontagem = Math.max(0, faltaBrutaMontagem - creditoPipeline);

        pesp = Math.min(
          sugerirQtdPorMinimo(needProd + faltaParaMontagem, refPesp),
          LIMITE_POR_NUMERO
        );

        montSizes[size] = mont;
        pespSizes[size] = pesp;

        montTotal += mont;
        pespTotal += pesp;
      } else {
        montSizes[size] = 0;
        pespSizes[size] = 0;
      }

      const coberturaPA = getCoberturaDias(pa, vendaDia);
      const coberturaFutura = getCoberturaDias(pa + prodAtual, vendaDia);

      const paZero = pa === 0;
      const paAbaixoMinimo = pa < minPA;
      const prodAbaixoMinimo = prodAtual < minProd;
      const semReposicao = prodAtual <= 0;

      const coberturaCriticaPA =
        vendaDia > 0 && coberturaPA <= Math.max(2, diasMontagem || 1);

      const coberturaRuimPA =
        vendaDia > 0 && coberturaPA <= Math.max(5, (diasMontagem || 1) + 2);

      const coberturaFuturaCritica =
        vendaDia > 0 && coberturaFutura <= Math.max(4, diasTotal || 2);

      const coberturaFuturaRuim =
        vendaDia > 0 && coberturaFutura <= Math.max(8, (diasTotal || 2) + 3);

      const produtoTopPorTamanho = vendaMes >= 40;

      const urgenteMontagem =
        (paZero && itemRelevante) ||
        coberturaCriticaPA ||
        (paAbaixoMinimo && vendaDia > 0 && semReposicao);

      const urgentePesponto =
        (paZero && itemRelevante) ||
        coberturaCriticaPA ||
        (coberturaFuturaCritica && vendaDia > 0);

      const recuperacaoMontagem =
        !urgenteMontagem &&
        vendaDia > 0 &&
        (
          (paAbaixoMinimo && est > 0) ||
          (coberturaRuimPA && est > 0) ||
          (produtoTopPorTamanho && coberturaRuimPA)
        );

      const recuperacaoPesponto =
        !urgentePesponto &&
        vendaDia > 0 &&
        (
          (prodAbaixoMinimo && prodAtual > 0) ||
          (coberturaFuturaRuim && prodAtual > 0) ||
          (produtoTopPorTamanho && coberturaFuturaRuim)
        );

      if (urgenteMontagem) itemUrgenteMontagem = true;
      if (urgentePesponto) itemUrgentePesponto = true;

      if (recuperacaoMontagem) itemRecuperacaoMontagem = true;
      if (recuperacaoPesponto) itemRecuperacaoPesponto = true;

      // recuperação por grade
      const tamanhoCriticoMontagem =
        itemRelevante &&
        (
          paZero ||
          paAbaixoMinimo ||
          coberturaCriticaPA ||
          coberturaRuimPA
        );

      const tamanhoPendenteMontagem =
        tamanhoCriticoMontagem &&
        (
          mont <= 0 ||
          (pa + est) < minPA
        );

      const tamanhoCriticoPesponto =
        itemRelevante &&
        (
          paZero ||
          prodAbaixoMinimo ||
          coberturaFuturaCritica ||
          coberturaFuturaRuim
        );

      const tamanhoPendentePesponto =
        tamanhoCriticoPesponto &&
        (
          pesp <= 0 ||
          (pa + prodAtual) < minPA
        );

      if (tamanhoCriticoMontagem) tamanhosCriticosMontagem += 1;
      if (tamanhoPendenteMontagem) tamanhosPendentesMontagem += 1;

      if (tamanhoCriticoPesponto) tamanhosCriticosPesponto += 1;
      if (tamanhoPendentePesponto) tamanhosPendentesPesponto += 1;

      let scoreMontagem = 0;
      let scorePesponto = 0;

      // 1. Estado crítico
      if (paZero) {
        scoreMontagem += 100000;
        scorePesponto += 100000;
      }

      if (coberturaCriticaPA) {
        scoreMontagem += 45000;
        scorePesponto += 38000;
      }

      if (paAbaixoMinimo) {
        scoreMontagem += 18000 + Math.max(0, minPA - pa) * 350;
      }

      if (coberturaFuturaCritica) {
        scorePesponto += 22000;
      }

      // 2. Estado de recuperação
      if (recuperacaoMontagem) {
        scoreMontagem += 35000;
      }

      if (recuperacaoPesponto) {
        scorePesponto += 32000;
      }

      // 3. Peso de vendas
      scoreMontagem += vendaMes * 45;
      scorePesponto += vendaMes * 45;

      // 4. Necessidade calculada
      scoreMontagem += needPA * 180;
      scorePesponto += needProd * 180;

      // 5. Mínimos como ajuste fino
      scoreMontagem += Math.max(0, minPA - pa) * 100;
      scorePesponto += Math.max(0, minProd - prodAtual) * 100;

      // 6. Reposição em andamento: reduz, mas pouco
      scoreMontagem -= est * 12;
      scorePesponto -= prodAtual * 10;

      if (semReposicao) {
        scoreMontagem += 7000;
        scorePesponto += 7000;
      }

      prioridadeMontagem += scoreMontagem;
      prioridadePesponto += scorePesponto;
    });

    const produtoTopRefCor = vendaTotalRefCor >= 120;

    const gradeRecuperacaoMontagem =
      produtoTopRefCor &&
      tamanhosCriticosMontagem >= 2 &&
      tamanhosPendentesMontagem >= 1;

    const gradeRecuperacaoPesponto =
      produtoTopRefCor &&
      tamanhosCriticosPesponto >= 2 &&
      tamanhosPendentesPesponto >= 1;

    if (gradeRecuperacaoMontagem) {
      prioridadeMontagem += 90000;
      itemRecuperacaoMontagem = true;
    }

    if (gradeRecuperacaoPesponto) {
      prioridadePesponto += 85000;
      itemRecuperacaoPesponto = true;
    }

    if (montTotal > 0) {
      montagem.push({
        tipo: "Montagem",
        ref: row.ref,
        cor: row.cor,
        sizes: montSizes,
        total: montTotal,
        prioridade: prioridadeMontagem,
        urgente: itemUrgenteMontagem,
        recuperacao: itemRecuperacaoMontagem,
        gradeRecuperacao: gradeRecuperacaoMontagem,
        tamanhosCriticos: tamanhosCriticosMontagem,
        tamanhosPendentes: tamanhosPendentesMontagem,
        vendaTotal: vendaTotalRefCor,
      });
    }

    if (pespTotal > 0) {
      pesponto.push({
        tipo: "Pesponto",
        ref: row.ref,
        cor: row.cor,
        sizes: pespSizes,
        total: pespTotal,
        prioridade: prioridadePesponto,
        urgente: itemUrgentePesponto,
        recuperacao: itemRecuperacaoPesponto,
        gradeRecuperacao: gradeRecuperacaoPesponto,
        tamanhosCriticos: tamanhosCriticosPesponto,
        tamanhosPendentes: tamanhosPendentesPesponto,
        vendaTotal: vendaTotalRefCor,
      });
    }
  });

  montagem.sort((a, b) => {
    if (Number(b.urgente) !== Number(a.urgente)) return Number(b.urgente) - Number(a.urgente);
    if (Number(b.gradeRecuperacao) !== Number(a.gradeRecuperacao)) return Number(b.gradeRecuperacao) - Number(a.gradeRecuperacao);
    if (Number(b.recuperacao) !== Number(a.recuperacao)) return Number(b.recuperacao) - Number(a.recuperacao);
    if (b.prioridade !== a.prioridade) return b.prioridade - a.prioridade;
    return b.total - a.total;
  });

  pesponto.sort((a, b) => {
    if (Number(b.urgente) !== Number(a.urgente)) return Number(b.urgente) - Number(a.urgente);
    if (Number(b.gradeRecuperacao) !== Number(a.gradeRecuperacao)) return Number(b.gradeRecuperacao) - Number(a.gradeRecuperacao);
    if (Number(b.recuperacao) !== Number(a.recuperacao)) return Number(b.recuperacao) - Number(a.recuperacao);
    if (b.prioridade !== a.prioridade) return b.prioridade - a.prioridade;
    return b.total - a.total;
  });

  return { montagem, pesponto };
}

function splitIntoFichas(sizesObj, maxPorFicha = 396) {
  const limite = Math.max(1, Number(maxPorFicha) || 396);
  const fichas = [];

  const tamanhos = Object.entries(sizesObj)
    .filter(([_, qtd]) => qtd > 0)
    .map(([size, qtd]) => ({
      size: Number(size),
      qtd: Number(qtd),
    }));

  if (!tamanhos.length) return fichas;

  let restante = tamanhos.map(t => ({ ...t }));

  let contador = 1;

  while (restante.some(t => t.qtd > 0)) {
    let ficha = {};
    let total = 0;

    for (let i = 0; i < restante.length; i++) {
      const item = restante[i];
      if (item.qtd <= 0) continue;

      const podeAdicionar = Math.min(item.qtd, limite - total);

      if (podeAdicionar <= 0) break;

      ficha[item.size] = podeAdicionar;
      item.qtd -= podeAdicionar;
      total += podeAdicionar;

      if (total >= limite) break;
    }

    if (total > 0) {
      fichas.push({
        nome: `Ficha ${contador}`,
        sizes: ficha,
        total,
      });
      contador++;
    } else {
      break;
    }
  }

  return fichas;
}

function buildProgramacaoPeriodo(
  fichasBase,
  suggestionsBase,
  capacidadeDia = 396,
  dias = 1,
  tipo = "",
  options = {}
) {
  const pctTop = Math.min(100, Math.max(0, Number(options.pctCapacidadeTop) || 0));
  const topN = Math.max(1, Math.min(200, Number(options.topN) || 10));
  const topMode = options.topMode === "manual" ? "manual" : "auto";
  const topManualSet = new Set(
    Array.isArray(options.topManualKeys) ? options.topManualKeys.filter((x) => typeof x === "string") : []
  );

  const prioridadeMap = new Map();
  const urgenciaMap = new Map();
  const recuperacaoMap = new Map();
  const gradeRecuperacaoMap = new Map();
  const vendaMap = new Map();

  (suggestionsBase || []).forEach((item) => {
    const key = `${item.ref}__${item.cor}`;
    prioridadeMap.set(key, item.prioridade || 0);
    urgenciaMap.set(key, !!item.urgente);
    recuperacaoMap.set(key, !!item.recuperacao);
    gradeRecuperacaoMap.set(key, !!item.gradeRecuperacao);
    vendaMap.set(key, Number(item.vendaTotal) || 0);
  });

  const grupos = {};
  (fichasBase || []).forEach((ficha) => {
    const key = `${ficha.ref}__${ficha.cor}`;

    if (!grupos[key]) {
      const prioridade = prioridadeMap.get(key) || 0;
      const urgente = urgenciaMap.get(key) || false;
      const recuperacao = recuperacaoMap.get(key) || false;
      const gradeRecuperacao = gradeRecuperacaoMap.get(key) || false;

      grupos[key] = {
        key,
        tipo,
        ref: ficha.ref,
        cor: ficha.cor,
        prioridade,
        urgente,
        recuperacao,
        gradeRecuperacao,
        peso: Math.max(1, prioridade),
        current: 0,
        fichas: [],
        topTier: false,
      };
    }

    grupos[key].fichas.push({ ...ficha, tipo });
  });

  const keysGrupo = Object.keys(grupos);
  const ranked = keysGrupo
    .map((key) => ({ key, venda: vendaMap.get(key) || 0 }))
    .sort((a, b) => b.venda - a.venda || String(a.key).localeCompare(String(b.key), "pt-BR"));
  const topSet =
    topMode === "manual"
      ? new Set(
          ranked
            .filter((x) => topManualSet.has(x.key))
            .map((x) => x.key)
        )
      : new Set(ranked.slice(0, Math.min(topN, ranked.length)).map((x) => x.key));

  keysGrupo.forEach((key) => {
    grupos[key].topTier = topSet.has(key);
  });

  Object.values(grupos).forEach((grupo) => {
    // Mantem a ordem de execucao da maior para a menor ficha em cada grupo.
    grupo.fichas.sort((a, b) => b.total - a.total || a.nome.localeCompare(b.nome, "pt-BR"));
  });

  const ativos = Object.values(grupos).filter((grupo) => grupo.fichas.length > 0);
  const totalPeso = ativos.reduce((acc, grupo) => acc + grupo.peso, 0) || 1;

  const usarReservaTop = pctTop > 0 && ativos.some((g) => g.topTier);
  const ativosTop = ativos.filter((g) => g.topTier);

  const diasProgramados = [];
  let gruposDiaAnterior = new Set();

  const escolherGrupo = (gruposDisponiveis, restante, ultimoGrupoDia, gruposBloqueadosDiaAnterior) => {
    const candidatos = gruposDisponiveis.filter(
      (grupo) => grupo.fichas.length > 0 && Number(grupo.fichas[0]?.total || 0) <= restante
    );

    if (!candidatos.length) return null;

    candidatos.forEach((grupo) => {
      grupo.current += grupo.peso;
    });

    const ordenados = [...candidatos].sort((a, b) => {
      const aFoiOntem = gruposBloqueadosDiaAnterior.has(a.key);
      const bFoiOntem = gruposBloqueadosDiaAnterior.has(b.key);

      const penalRepeticaoA =
        a.urgente ? 0 : a.gradeRecuperacao ? 0 : a.recuperacao ? 1 : 3;

      const penalRepeticaoB =
        b.urgente ? 0 : b.gradeRecuperacao ? 0 : b.recuperacao ? 1 : 3;

      const penalA =
        (a.key === ultimoGrupoDia ? 1 : 0) +
        (aFoiOntem ? penalRepeticaoA : 0);

      const penalB =
        (b.key === ultimoGrupoDia ? 1 : 0) +
        (bFoiOntem ? penalRepeticaoB : 0);

      if (Number(b.urgente) !== Number(a.urgente)) return Number(b.urgente) - Number(a.urgente);
      if (Number(b.gradeRecuperacao) !== Number(a.gradeRecuperacao)) return Number(b.gradeRecuperacao) - Number(a.gradeRecuperacao);
      if (Number(b.recuperacao) !== Number(a.recuperacao)) return Number(b.recuperacao) - Number(a.recuperacao);
      if (penalA !== penalB) return penalA - penalB;
      if (b.current !== a.current) return b.current - a.current;
      if (b.prioridade !== a.prioridade) return b.prioridade - a.prioridade;
      return b.fichas.length - a.fichas.length;
    });

    const escolhido = ordenados[0] || null;
    if (!escolhido) return null;

    escolhido.current -= totalPeso;
    return escolhido;
  };

  for (let dia = 1; dia <= dias; dia += 1) {
    let ultimoGrupoDia = "";
    const selecionadas = [];
    const gruposSelecionadosNoDia = new Set();

    const runPhase = (restanteInicial, pool, marcarReservaTop) => {
      let restante = restanteInicial;
      while (restante > 0) {
        const grupoEscolhido = escolherGrupo(pool, restante, ultimoGrupoDia, gruposDiaAnterior);

        if (!grupoEscolhido) break;

        const fichaEscolhida = grupoEscolhido.fichas[0];

        if (!fichaEscolhida || fichaEscolhida.total > restante) {
          grupoEscolhido.current -= totalPeso;
          break;
        }

        grupoEscolhido.fichas = grupoEscolhido.fichas.filter((f) => f.nome !== fichaEscolhida.nome);

        selecionadas.push({
          ...fichaEscolhida,
          prioridade: grupoEscolhido.prioridade,
          urgente: grupoEscolhido.urgente,
          recuperacao: grupoEscolhido.recuperacao,
          gradeRecuperacao: grupoEscolhido.gradeRecuperacao,
          grupoKey: grupoEscolhido.key,
          ...(marcarReservaTop ? { programacaoReservaTop: true } : {}),
        });

        gruposSelecionadosNoDia.add(grupoEscolhido.key);
        restante -= fichaEscolhida.total;
        ultimoGrupoDia = grupoEscolhido.key;
      }
      return restante;
    };

    let restanteFinal;
    if (usarReservaTop && ativosTop.length > 0) {
      const reservaTopPares = Math.round((capacidadeDia * pctTop) / 100);
      let restanteGeral = capacidadeDia - reservaTopPares;
      const sobraTop = runPhase(reservaTopPares, ativosTop, true);
      restanteGeral += sobraTop;
      restanteFinal = runPhase(restanteGeral, ativos, false);
    } else {
      restanteFinal = runPhase(capacidadeDia, ativos, false);
    }

    diasProgramados.push({
      dia,
      capacidadeDia,
      totalProgramado: capacidadeDia - restanteFinal,
      restante: restanteFinal,
      fichas: selecionadas,
    });

    gruposDiaAnterior = new Set(gruposSelecionadosNoDia);
  }

  const todasFichas = diasProgramados.flatMap((item) => item.fichas);

  return {
    tipo,
    dias,
    capacidadeDia,
    totalProgramado: todasFichas.reduce((acc, item) => acc + item.total, 0),
    totalRestante: diasProgramados.reduce((acc, item) => acc + item.restante, 0),
    diasProgramados,
    totalFichas: todasFichas.length,
    reservaTopVendasPct: usarReservaTop ? pctTop : 0,
    reservaTopN: topN,
    reservaTopAtiva: usarReservaTop,
    reservaTopModo: topMode,
    reservaTopManualSelecionados: topSet.size,
  };
}

function SummaryCard({ title, value, subtitle, action }) {
  return (
    <div className="relative overflow-hidden rounded-[28px] border border-[#E5E7EB] bg-white p-5 shadow-[0_12px_30px_rgba(15,23,42,0.08)] backdrop-blur">
      <div className="absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-[#8B1E2D] via-[#6F1421] to-[#0F172A]" />
      <div className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-400">{title}</div>
      <div className="mt-3 text-3xl font-black tracking-tight text-[#0F172A]">{value}</div>
      <div className="mt-2 text-sm text-slate-500">{subtitle}</div>
      {action ? <div className="mt-3">{action}</div> : null}
    </div>
  );
}

function RelatorioProducaoPrintDocument({ data }) {
  const titulo = data.titulo || "RELATÓRIO DE PRODUÇÃO";
  const obs = data.observacoes?.trim();

  return (
    <div className="relatorio-print-doc text-slate-900 [print-color-adjust:exact]">
      <div className="flex flex-col sm:flex-row sm:items-stretch sm:justify-between gap-4 border-b-4 border-[#8B1E2D] pb-4 print-section">
        <div className="flex flex-col justify-center min-w-[120px]">
          <div className="text-lg font-black tracking-[0.12em] text-[#8B1E2D]">ROCK STAR</div>
          <div className="text-[10px] uppercase tracking-[0.2em] text-slate-500 mt-1">Módulo Produção</div>
        </div>
        <div className="text-center flex-1 px-2 min-w-0">
          <div className="text-xl sm:text-2xl font-black text-[#0F172A] leading-tight">{titulo}</div>
          <div className="text-xs text-slate-500 mt-1">Documento gerado pelo sistema</div>
        </div>
        <div className="text-xs text-left sm:text-right text-slate-600 sm:min-w-[170px] sm:self-end">
          <div className="font-semibold text-slate-800">Gerado em</div>
          <div className="mt-1">{data.dataGeracao}</div>
        </div>
      </div>

      {obs ? (
        <div className="mt-4 rounded-xl border border-amber-300/80 bg-amber-50 p-3 text-sm print-section">
          <div className="text-[11px] font-bold uppercase tracking-wide text-amber-900">Observações</div>
          <div className="mt-2 whitespace-pre-wrap text-amber-950 leading-relaxed">{obs}</div>
        </div>
      ) : null}

      <div className="mt-4 rounded-xl border border-slate-200 bg-gradient-to-br from-slate-50 to-white p-3 text-xs print-section">
        <div className="font-bold text-slate-800">Filtros aplicados</div>
        <div className="mt-2 grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-1 text-slate-600">
          <div>
            <span className="text-slate-400">Período:</span> {data.filtros.periodo}
          </div>
          <div>
            <span className="text-slate-400">REF:</span> {data.filtros.ref}
          </div>
          <div>
            <span className="text-slate-400">Cor:</span> {data.filtros.cor}
          </div>
          <div>
            <span className="text-slate-400">Setor:</span> {data.filtros.setor}
          </div>
          <div className="sm:col-span-2">
            <span className="text-slate-400">Status:</span> {data.filtros.status}
          </div>
        </div>
      </div>

      <div className="mt-4 grid grid-cols-2 md:grid-cols-4 gap-3 print-section">
        <div className="rounded-xl border border-slate-200 border-l-4 border-l-[#8B1E2D] bg-white p-3 text-center shadow-sm">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Programações</div>
          <div className="text-2xl font-black text-[#0F172A] mt-1">{data.resumo.programacoes}</div>
        </div>
        <div className="rounded-xl border border-slate-200 border-l-4 border-l-[#0F172A] bg-white p-3 text-center shadow-sm">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Total pares</div>
          <div className="text-2xl font-black text-[#0F172A] mt-1">{data.resumo.totalPares}</div>
        </div>
        <div className="rounded-xl border border-slate-200 border-l-4 border-l-amber-500 bg-white p-3 text-center shadow-sm">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Em aberto</div>
          <div className="text-2xl font-black text-amber-800 mt-1">{data.resumo.emAberto}</div>
        </div>
        <div className="rounded-xl border border-slate-200 border-l-4 border-l-emerald-600 bg-white p-3 text-center shadow-sm">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Finalizados</div>
          <div className="text-2xl font-black text-emerald-800 mt-1">{data.resumo.finalizados}</div>
        </div>
      </div>

      {data.gruposPorSetor.map((grupo) => (
        <div key={`doc-${grupo.setor}`} className="mt-6 print-section">
          <div
            className={`text-sm font-black uppercase tracking-[0.15em] pb-2 border-b-2 ${
              grupo.setor === "Pesponto" ? "text-amber-800 border-amber-300" : "text-sky-900 border-sky-300"
            }`}
          >
            {grupo.setor}
          </div>
          {grupo.linhas.length === 0 ? (
            <div className="mt-3 rounded-xl border border-dashed border-slate-300 bg-slate-50 p-4 text-sm text-slate-500 text-center">
              Sem registros neste setor.
            </div>
          ) : (
            <div className="mt-3 overflow-x-auto rounded-xl border border-slate-200">
              <table className="w-full border-collapse text-[11px] min-w-[640px]">
                <thead>
                  <tr className="bg-[#0F172A] text-white">
                    <th className="px-2 py-2.5 text-left font-semibold border border-[#1e293b]">Data</th>
                    <th className="px-2 py-2.5 text-left font-semibold border border-[#1e293b]">Programação</th>
                    <th className="px-2 py-2.5 text-left font-semibold border border-[#1e293b]">REF</th>
                    <th className="px-2 py-2.5 text-left font-semibold border border-[#1e293b]">Cor</th>
                    <th className="px-2 py-2.5 text-center font-semibold border border-[#1e293b]">Pares</th>
                    <th className="px-2 py-2.5 text-center font-semibold border border-[#1e293b]">Status</th>
                    <th className="px-2 py-2.5 text-left font-semibold border border-[#1e293b]">Detalhe</th>
                  </tr>
                </thead>
                <tbody>
                  {grupo.linhas.map((item, idx) => {
                    const dataRef = item.status === "Finalizado" && item.dataFinalizacao ? item.dataFinalizacao : item.dataLancamento;
                    const finalizado = item.status === "Finalizado";
                    return (
                      <tr key={`doc-${grupo.setor}-${item.id}`} className={idx % 2 === 0 ? "bg-white" : "bg-slate-50/90"}>
                        <td className="border border-slate-200 px-2 py-2 align-top text-slate-800">{dataRef || "—"}</td>
                        <td className="border border-slate-200 px-2 py-2 align-top font-medium text-slate-900">{item.programacao || "—"}</td>
                        <td className="border border-slate-200 px-2 py-2 align-top">{item.ref || "—"}</td>
                        <td className="border border-slate-200 px-2 py-2 align-top">{item.cor || "—"}</td>
                        <td className="border border-slate-200 px-2 py-2 text-center font-bold text-slate-900">{item.total ?? 0}</td>
                        <td className="border border-slate-200 px-2 py-2 text-center">
                          <span
                            className={`inline-block px-2 py-0.5 rounded-md text-[10px] font-bold ${
                              finalizado ? "bg-emerald-100 text-emerald-800 border border-emerald-200" : "bg-amber-100 text-amber-900 border border-amber-200"
                            }`}
                          >
                            {item.status || "—"}
                          </span>
                        </td>
                        <td className="border border-slate-200 px-2 py-2 text-slate-700">
                          {(item.items || []).map((entry) => `${entry.size}: ${entry.qtd}`).join(" · ") || "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      ))}

      <div className="mt-10 grid grid-cols-2 gap-8 print-section">
        <div className="text-center pt-6 border-t-2 border-slate-300">
          <div className="h-8" />
          <div className="text-xs font-semibold text-slate-600">Responsável Produção</div>
        </div>
        <div className="text-center pt-6 border-t-2 border-slate-300">
          <div className="h-8" />
          <div className="text-xs font-semibold text-slate-600">Conferência</div>
        </div>
      </div>
    </div>
  );
}

function ProgramacaoDiaFolhaImpressao({
  titulo,
  logoSrc = "/logo-rockstar-bandeira.png",
  setor,
  modoLabel,
  diasCount,
  dataImpressao,
  observacoes,
  itens,
  sizesList,
  folhaSubtitle = "Fichas selecionadas — otimizado para A4",
  copiasPorPagina = 1,
  etiquetaFichaCustom = "",
  cabecalhoFolha = "completo",
  valoresParTerceiros = {},
  tipoFolhaImpressao = "folha1",
}) {
  const obs = observacoes?.trim();
  const etiquetaBase = String(etiquetaFichaCustom || "").trim();
  const modoCabecalho = cabecalhoFolha === "minimo" || cabecalhoFolha === "oculto" ? cabecalhoFolha : "completo";
  const tamanhosFolha = (Array.isArray(sizesList) ? sizesList : [])
    .map((s) => Number(s))
    .filter((s) => Number.isFinite(s) && s >= 34 && s <= 44);
  const tamanhosGrade = tamanhosFolha.length > 0 ? tamanhosFolha : (Array.isArray(sizesList) ? sizesList : []);
  const destinatariosFolha1 = [
    { nome: "WEVERTON", chaveValor: "weverton", regra: "valor fixo" },
    { nome: "ROMULO", chaveValor: "romulo", regra: "valor fixo" },
    { nome: "GU", chaveValor: "guPorReferencia", regra: "por referencia" },
  ];
  const getValorGuPorReferencia = (ref) => {
    const codigo = String(ref || "").trim().toUpperCase();
    if (codigo === "BTCV010" || codigo === "TNCV010") return 0.4;
    if (codigo === "CRVTNCV") return 0.3;
    return Number.NaN;
  };
  const getNomeReferencia = (ref) => {
    const codigo = String(ref || "").trim().toUpperCase();
    if (codigo === "TNCV010") return "CANO BAIXO";
    if (codigo === "BTCV010") return "CANO ALTO";
    if (codigo === "CRVTNCV") return "COURINO";
    return "";
  };
  const formatarMoedaBr = (valor) => {
    const num = Number(valor);
    if (!Number.isFinite(num)) return "—";
    return num.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
  };
  const extractFichaToken = (nome = "") => {
    const m = String(nome).match(/ficha\s*(\d+)/i);
    return m ? `Ficha ${String(m[1]).padStart(2, "0")}` : "";
  };

  const extractNomeSemFicha = (nome = "", ref = "", cor = "") => {
    let base = String(nome || "");
    if (ref) base = base.replace(String(ref), "");
    if (cor) base = base.replace(String(cor), "");
    base = base.replace(/[-–]\s*/g, " ").replace(/\bficha\s*\d+\b/i, "").replace(/\s+/g, " ").trim();
    return base;
  };

  const gruposFichaMap = new Map();
  (itens || []).forEach((row, idx) => {
    const { ficha, dia, ordem } = row;
    const copia = Math.max(1, Number(row?.copia) || 1);
    const programacaoNomeGrupo = String(row?.programacaoNome || "").trim();
    const agrupador = programacaoNomeGrupo ? `prog__${programacaoNomeGrupo}` : `dia__${dia}`;
    const key = `${copia}__${agrupador}`;
    if (!gruposFichaMap.has(key)) {
      gruposFichaMap.set(key, {
        key,
        copia,
        dia,
        programacaoNome: programacaoNomeGrupo,
        ordem: ordem || idx + 1,
        fichaToken: "",
        itens: [],
      });
    }
    gruposFichaMap.get(key).itens.push(row);
  });

  const gruposFicha = Array.from(gruposFichaMap.values()).sort((a, b) => a.copia - b.copia || a.dia - b.dia || a.ordem - b.ordem);
  const totalCopias = Math.max(1, ...gruposFicha.map((g) => Number(g.copia) || 1), 1);
  const copiasPorPaginaEfetivo = Math.max(1, Math.min(4, Number(copiasPorPagina) || 1));
  const isEconomico = copiasPorPaginaEfetivo > 1;
  const blocosPorPaginaEconomico = copiasPorPaginaEfetivo;
  const repeticoesPorFichaBase = gruposFicha.reduce((acc, grupo) => {
    const baseKey = `${grupo.dia}__${grupo.fichaToken}`;
    acc[baseKey] = (acc[baseKey] || 0) + 1;
    return acc;
  }, {});

  return (
    <div
      className={`programacao-print-doc text-slate-900 ${modoCabecalho === "oculto" ? "programacao-print-doc--cabecalho-oculto" : ""} ${
        isEconomico ? `programacao-print-economico programacao-print-economico-${blocosPorPaginaEconomico}` : ""
      }`}
    >
      {modoCabecalho === "completo" ? (
        <>
          <div className="flex flex-wrap items-center justify-between gap-3 border-b-[3px] border-[#8B1E2D] pb-3 mb-3">
            <img src={logoSrc} alt="Logo" className="h-11 w-auto max-w-[min(100%,280px)] object-contain object-left shrink-0" />
            <div className="flex-1 text-center px-2 min-w-[120px]">
              <h1 className="text-[15pt] font-black text-[#0F172A] leading-tight">{titulo || "Programação do Dia"}</h1>
              <p className="text-[8pt] text-slate-500 mt-1">{folhaSubtitle}</p>
            </div>
            <div className="text-[8pt] text-slate-600 text-right shrink-0">
              <div className="font-semibold text-slate-800">Impresso em</div>
              <div>{dataImpressao}</div>
            </div>
          </div>

          <div className="flex flex-wrap gap-1.5 mb-3 text-[8pt]">
            <span className="px-2 py-0.5 rounded-md border border-slate-300 bg-slate-100 font-bold text-[#0F172A]">Setor: {setor}</span>
            <span className="px-2 py-0.5 rounded-md border border-slate-200 bg-white">{modoLabel}</span>
            <span className="px-2 py-0.5 rounded-md border border-slate-200 bg-white">Período: {diasCount} dia(s)</span>
            <span className="px-2 py-0.5 rounded-md border border-slate-200 bg-white">{itens.length} ficha(s)</span>
          </div>
        </>
      ) : modoCabecalho === "minimo" ? (
        <div className="programacao-print-cabecalho-minimo mb-2 pb-1.5 border-b border-slate-300 text-[7pt] text-slate-700 flex flex-wrap items-baseline justify-between gap-x-3 gap-y-1">
          <span className="font-black text-slate-900">{titulo || "Programação do Dia"}</span>
          <span className="text-slate-600">
            {setor} · {modoLabel} · {diasCount} dia(s) · {itens.length} bloco(s)
          </span>
          <span className="text-slate-500 tabular-nums">{dataImpressao}</span>
        </div>
      ) : null}

      <div className={isEconomico ? "programacao-print-grid-economico" : "space-y-3"}>
        {gruposFicha.map((grupo, i) => {
          const totalProg = grupo.itens.reduce((acc, row) => acc + (Number(row.ficha?.total) || 0), 0);
          const refsMap = new Map();
          grupo.itens.forEach((row) => {
            const ficha = row.ficha || {};
            const ref = String(ficha.ref || "—");
            const cor = String(ficha.cor || "—");
            if (!refsMap.has(ref)) {
              refsMap.set(ref, {
                ref,
                nome: extractNomeSemFicha(ficha.nome || "", ficha.ref || "", ficha.cor || ""),
                rows: [],
              });
            }
            refsMap.get(ref).rows.push({
              cor,
              sizes: Object.fromEntries(tamanhosGrade.map((s) => [s, Number(ficha.sizes?.[s] || 0)])),
              total: Number(ficha.total) || 0,
            });
          });
          const refs = Array.from(refsMap.values());
          const linhasCor = refs.reduce((acc, ref) => acc + ref.rows.length, 0);
          const blocoGrande = refs.length > 2 || linhasCor > 4;
          const baseKey = `${grupo.dia}__${grupo.fichaToken}`;
          const fichaDuplicada = Number(repeticoesPorFichaBase[baseKey] || 0) > 1;
          const quebraEconomica =
            isEconomico &&
            (((!fichaDuplicada && blocoGrande) || (((i + 1) % blocosPorPaginaEconomico === 0) && i !== gruposFicha.length - 1)));
          const linhaEtiquetaFicha = etiquetaBase
            ? etiquetaBase.replace(/\{dia\}/gi, String(grupo.dia).padStart(2, "0"))
            : (grupo.programacaoNome
              ? `PROGRAMAÇÃO: ${grupo.programacaoNome}`
              : `PROGRAMAÇÃO: Dia ${String(grupo.dia).padStart(2, "0")}`);
          const isFolha1 = tipoFolhaImpressao === "folha1";
          const destinatario = isFolha1 ? (destinatariosFolha1[grupo.copia - 1] || null) : null;
          const etiquetaDestinatario = destinatario?.nome || "";
          const valorParBaseNumero = parseDecimalInput(destinatario ? valoresParTerceiros?.[destinatario.chaveValor] : "");
          const valorParTexto =
            destinatario?.chaveValor === "guPorReferencia"
              ? "BTCV010/TNCV010: R$ 0,40 · CRVTNCV: R$ 0,30"
              : formatarMoedaBr(valorParBaseNumero);
          const totalFichaValor =
            destinatario?.chaveValor === "guPorReferencia"
              ? refs.reduce((accRefs, refBlock) => {
                  const valorRef = getValorGuPorReferencia(refBlock.ref);
                  const valorAplicado = Number.isFinite(valorRef) ? valorRef : Number.NaN;
                  const totalRef = refBlock.rows.reduce((accRows, rowCor) => accRows + (Number(rowCor.total) || 0), 0);
                  if (!Number.isFinite(valorAplicado)) return accRefs;
                  return accRefs + totalRef * valorAplicado;
                }, 0)
              : (Number.isFinite(valorParBaseNumero) ? totalProg * valorParBaseNumero : NaN);
          const totalFichaTexto = formatarMoedaBr(totalFichaValor);
          const etiquetaFolha2 = `PESPONTO (${grupo.copia}/${totalCopias})`;
          return (
            <div
              key={`pf-group-${grupo.key}-${i}`}
              className={`programacao-print-ficha rounded-lg border border-slate-400 bg-white p-2.5 break-inside-avoid page-break-inside-avoid shadow-sm ${
                isEconomico ? "programacao-print-ficha-economica" : ""
              } ${quebraEconomica ? "programacao-print-item-break" : ""}`}
            >
              <div className="border-b border-slate-300 pb-1.5 mb-2">
                <div className="text-[9pt] font-black text-slate-900 flex items-center justify-between gap-2">
                  <span className="min-w-0 whitespace-normal break-words leading-tight">{linhaEtiquetaFicha}</span>
                  {totalCopias > 1 ? (
                    <span className="text-[7pt] font-bold text-slate-600">
                      {isFolha1
                        ? (etiquetaDestinatario
                          ? `${etiquetaDestinatario} (${grupo.copia}/${totalCopias})`
                          : `Cópia ${grupo.copia}/${totalCopias}`)
                        : etiquetaFolha2}
                    </span>
                  ) : null}
                </div>
                <div className="text-[7pt] text-slate-500">
                  {setor} · {modoLabel}
                </div>
                {isFolha1 && destinatario ? (
                  <div className="mt-0.5 text-[7pt] text-slate-700">
                    <span className="font-semibold">Valor/par: </span>
                    <span className="font-bold">{valorParTexto}</span>
                    <span className="text-slate-500"> · {destinatario.regra}</span>
                  </div>
                ) : null}
              </div>

              {refs.map((refBlock, refIdx) => (
                <div
                  key={`refblock-${grupo.key}-${refBlock.ref}-${refIdx}`}
                  className={`programacao-print-ref-block ${refIdx > 0 ? "mt-2.5" : ""}`}
                >
                  <div className="text-[8pt] font-bold text-slate-800 mb-1">
                    {`${refBlock.ref}${getNomeReferencia(refBlock.ref) ? ` - ${getNomeReferencia(refBlock.ref)}` : ""}`}
                  </div>
                  <table className="programacao-print-grade-table w-full border-collapse text-[8pt] leading-normal table-fixed box-border">
                    <colgroup>
                      <col className="w-[28mm]" />
                      {tamanhosGrade.map((s) => (
                        <col key={`${grupo.key}-${refBlock.ref}-col-${s}`} className="w-[8mm]" />
                      ))}
                      <col className="w-[10mm]" />
                    </colgroup>
                    <thead>
                      <tr className="bg-slate-600 text-white">
                        <th className="border border-slate-500 bg-slate-600 px-1 py-1 align-middle text-left font-bold text-white">COR</th>
                        {tamanhosGrade.map((s) => (
                          <th
                            key={`${grupo.key}-${refBlock.ref}-h-${s}`}
                            className="border border-slate-500 bg-slate-600 px-0.5 py-1 align-middle text-center font-bold text-white"
                          >
                            {s}
                          </th>
                        ))}
                        <th className="border border-slate-500 bg-slate-600 px-1 py-1 align-middle text-right font-black text-white">TOT</th>
                      </tr>
                    </thead>
                    <tbody>
                      {refBlock.rows.map((corRow, rowIdx) => (
                        <tr key={`${grupo.key}-${refBlock.ref}-${corRow.cor}-${rowIdx}`}>
                          <td className="border border-slate-300 px-1 py-1 align-middle text-left font-semibold text-slate-800 break-words [overflow-wrap:anywhere]">
                            {corRow.cor}
                          </td>
                          {tamanhosGrade.map((s) => (
                            <td
                              key={`${grupo.key}-${refBlock.ref}-${corRow.cor}-${rowIdx}-${s}`}
                              className="border border-slate-300 px-0.5 py-1 align-middle text-center tabular-nums text-slate-700"
                            >
                              {corRow.sizes[s] > 0 ? corRow.sizes[s] : ""}
                            </td>
                          ))}
                          <td className="border border-slate-300 px-1 py-1 align-middle text-right font-bold tabular-nums text-slate-900">
                            {corRow.total}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}

              {isFolha1 ? (
                <div className="mt-1 rounded-md border border-slate-300 bg-slate-50 px-2 py-1.5 print-section">
                  <div className="text-[6.5pt] font-semibold uppercase tracking-wide text-slate-500">Resumo da ficha</div>
                  <div className="mt-0.5 flex items-baseline justify-between gap-2">
                    <span className="text-[9pt] font-bold text-slate-800">TOTAL PROG</span>
                    <span className="text-[11pt] font-black tabular-nums text-slate-900">{totalProg}</span>
                  </div>
                  {destinatario ? (
                    <div className="mt-0.5 flex items-baseline justify-between gap-2 border-t border-slate-200 pt-0.5">
                      <span className="text-[9pt] font-black text-slate-900">TOTAL FICHA</span>
                      <span className="text-[12pt] font-black tabular-nums text-slate-900">{totalFichaTexto}</span>
                    </div>
                  ) : null}
                </div>
              ) : (
                <div className="mt-0.5 text-[7pt] text-slate-700 print-section">
                  <span className="font-bold">TOTAL PROG:</span>{" "}
                  <span className="font-black tabular-nums">{totalProg}</span>
                </div>
              )}

              {obs ? (
                <div className={`mt-1 text-[7pt] text-slate-800 print-section ${isEconomico ? "text-[6pt] leading-snug" : ""}`}>
                  <span className="font-bold">Obs.: </span>
                  <span className="whitespace-pre-wrap break-words">{obs}</span>
                </div>
              ) : null}

              <div className="mt-3 grid grid-cols-2 gap-3 border-t border-slate-300 pt-2 text-[7pt] print-section">
                <div className="text-center">
                  <div className="h-5 border-b border-slate-400" />
                  <div className="mt-1 font-semibold text-slate-700">Responsável</div>
                </div>
                <div className="text-center">
                  <div className="h-5 border-b border-slate-400" />
                  <div className="mt-1 font-semibold text-slate-700">Conferência</div>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PageShell({ children, title, subtitle, action }) {
  return (
    <div className="space-y-6">
      <header className="rounded-[32px] border border-[#E5E7EB] bg-white p-6 shadow-[0_14px_35px_rgba(15,23,42,0.08)] backdrop-blur">
        <div className="flex flex-col gap-5 lg:flex-row lg:items-start lg:justify-between">
          <div>
            <div className="inline-flex items-center rounded-full border border-[#E7C7CC] bg-[#FFF7F8] px-3 py-1 text-[11px] font-semibold uppercase tracking-[0.2em] text-[#8B1E2D]">
              Rock Star • Produção
            </div>
            <h1 className="mt-4 text-3xl font-black tracking-tight text-[#0F172A] lg:text-4xl">{title}</h1>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-slate-500 lg:text-[15px]">{subtitle}</p>
          </div>
          {action ? <div className="shrink-0">{action}</div> : null}
        </div>
      </header>
      {children}
    </div>
  );
}

export default function ModuloProducaoPreviewRecuperado() {
  const [active, setActive] = useState("Dashboard");
const [rows, setRows] = useState([]);
const [minimos, setMinimos] = useState({});
const [vendas, setVendas] = useState({});
const [importText, setImportText] = useState("");
const [importFileName, setImportFileName] = useState("");
const [importFeedback, setImportFeedback] = useState("");
const [importPreview, setImportPreview] = useState([]);
const [ultimaImportacaoGcm, setUltimaImportacaoGcm] = useState(null);
const [salesImportFileName, setSalesImportFileName] = useState("");
const [salesImportFeedback, setSalesImportFeedback] = useState("");
const [salesImportPreview, setSalesImportPreview] = useState([]);
const [vendasDraft, setVendasDraft] = useState({});
const [vendasDirty, setVendasDirty] = useState(false);
const [historicoVendasManuais, setHistoricoVendasManuais] = useState([]);
const [pespontoForm, setPespontoForm] = useState({
  ref: "",
  cor: "",
  grid: makeEmptyGrid(),
  programacao: "Programação A",
});
const [montagemForm, setMontagemForm] = useState({
  ref: "",
  cor: "",
  grid: makeEmptyGrid(),
  programacao: "Programação A",
});
const [pespontoLancamentos, setPespontoLancamentos] = useState([]);
const [montagemLancamentos, setMontagemLancamentos] = useState([]);
const [previewFicha, setPreviewFicha] = useState(null);
const [confirmImport, setConfirmImport] = useState(false);
const [importMode, setImportMode] = useState("replace");
const [movError, setMovError] = useState({ Pesponto: "", Montagem: "" });
const [confirmMov, setConfirmMov] = useState(null);
const [editingMov, setEditingMov] = useState(null);
const [controleFiltroRef, setControleFiltroRef] = useState("TODAS");
const [controleFiltroCor, setControleFiltroCor] = useState("TODAS");
const [controleFiltroNumero, setControleFiltroNumero] = useState("TODAS");
const [ajusteEstForm, setAjusteEstForm] = useState({
  ref: "",
  cor: "",
  tipo: "entrada",
  size: 34,
  qtd: 0,
  motivo: "",
});
const [ajustesEst, setAjustesEst] = useState([]);
const [ajusteEstErro, setAjusteEstErro] = useState("");
const [draftMinimos, setDraftMinimos] = useState({});
const [dirtyMinimos, setDirtyMinimos] = useState(false);
const [capacidadePespontoDia, setCapacidadePespontoDia] = useState(396);
const [capacidadeMontagemDia, setCapacidadeMontagemDia] = useState(396);
const [programacaoReservaTopPct, setProgramacaoReservaTopPct] = useState(readProgReservaTopPctFromStorage);
const [programacaoTopN, setProgramacaoTopN] = useState(readProgTopNFromStorage);
const [programacaoTopModo, setProgramacaoTopModo] = useState(readProgTopModeFromStorage);
const [programacaoTopManualKeys, setProgramacaoTopManualKeys] = useState(readProgTopManualKeysFromStorage);
const [tempoProducao, setTempoProducao] = useState(initialTempoProducao);
const [tempoProducaoDraft, setTempoProducaoDraft] = useState(initialTempoProducao);
const [fichasAbertas, setFichasAbertas] = useState({});
const [programacaoDias, setProgramacaoDias] = useState(7);
const [confirmAction, setConfirmAction] = useState(null);
const [relatorioDataInicial, setRelatorioDataInicial] = useState("");
const [relatorioDataFinal, setRelatorioDataFinal] = useState("");
const [relatorioSetor, setRelatorioSetor] = useState("TODOS");
const [relatorioStatus, setRelatorioStatus] = useState("TODOS");
const [relatorioRef, setRelatorioRef] = useState("TODAS");
const [relatorioCor, setRelatorioCor] = useState("TODAS");
const [printRelatorioData, setPrintRelatorioData] = useState(null);
const [relatorioPrintModalOpen, setRelatorioPrintModalOpen] = useState(false);
const [relatorioPrintDraft, setRelatorioPrintDraft] = useState(null);
const [relatorioPrintTitulo, setRelatorioPrintTitulo] = useState("RELATÓRIO DE PRODUÇÃO");
const [relatorioPrintObs, setRelatorioPrintObs] = useState("");
const [programacaoSubAba, setProgramacaoSubAba] = useState("Pesponto");
const [programacaoModoVisual, setProgramacaoModoVisual] = useState("normal");
const [programacaoObsImpressao, setProgramacaoObsImpressao] = useState("");
const [programacaoLogoImpressao, setProgramacaoLogoImpressao] = useState("/logo-rockstar-bandeira.png");
const [programacaoCopiasPorPagina, setProgramacaoCopiasPorPagina] = useState(1);
const [programacaoEtiquetaFicha, setProgramacaoEtiquetaFicha] = useState("");
const [programacaoNomeLoteImpressao, setProgramacaoNomeLoteImpressao] = useState("");
const [programacaoCabecalhoFolha, setProgramacaoCabecalhoFolha] = useState("completo");
const [programacaoValoresTerceiros, setProgramacaoValoresTerceiros] = useState(readProgValoresPagamentoFromStorage);
const [programacaoTipoFolha, setProgramacaoTipoFolha] = useState("folha1");
const [movImpressaoSelecao, setMovImpressaoSelecao] = useState({ Pesponto: {}, Montagem: {} });
const [programacaoFichaSelecao, setProgramacaoFichaSelecao] = useState({});
const [programacaoPdfBusy, setProgramacaoPdfBusy] = useState(false);
const programacaoPrintSheetRef = useRef(null);
const [lancarFichaDaProgramacao, setLancarFichaDaProgramacao] = useState(null);
/** Chaves de fichas já lançadas (persistido em localStorage). */
const [fichasProgramacaoLancadas, setFichasProgramacaoLancadas] = useState(readProgFichasLancadasFromStorage);
/** Plano congelado na Programação do Dia (não muda com estoque até Recalcular ou mudança período/capacidade). */
const [programacaoPlanoCongelado, setProgramacaoPlanoCongelado] = useState(null);
const [movListPage, setMovListPage] = useState({ Pesponto: 1, Montagem: 1 });
const [feriadosTexto, setFeriadosTexto] = useState("");
const [dashboardFeriadosAberto, setDashboardFeriadosAberto] = useState(false);
const [dashboardMaisKpisAberto, setDashboardMaisKpisAberto] = useState(false);
const [dashboardMobileTab, setDashboardMobileTab] = useState("visao");

const addLaunchedProgFichaKey = useCallback((key) => {
  if (!key) return;
  setFichasProgramacaoLancadas((prev) => {
    if (prev.includes(key)) return prev;
    const next = [...prev, key];
    try {
      localStorage.setItem(PROG_FICHAS_LANCADAS_STORAGE_KEY, JSON.stringify(next));
    } catch (_) {
      /* ignore */
    }
    return next;
  });
}, []);

const limparMarcasFichasLancadasProgramacao = useCallback(() => {
  setFichasProgramacaoLancadas([]);
  try {
    localStorage.removeItem(PROG_FICHAS_LANCADAS_STORAGE_KEY);
  } catch (_) {
    /* ignore */
  }
}, []);

const refs = useMemo(() => rows.map((r) => `${r.ref}__${r.cor}`), [rows]);
const firstRef = refs[0] ? refs[0].split("__")[0] : "";
const firstCor = refs[0] ? refs[0].split("__")[1] : "";

  const rowsNormalized = useMemo(
    () => rows.map((row) => ({ ...row, data: normalizeProductData(row.data) })),
    [rows]
  );

const previewBySelection = (form) => {
  const row = rowsNormalized.find((item) => item.ref === form.ref && item.cor === form.cor);
  if (!row) return null;

  const totalPA = sizes.reduce((acc, size) => acc + (row.data[size]?.pa || 0), 0);
  const totalPesponto = sizes.reduce((acc, size) => acc + (row.data[size]?.p || 0), 0);
  const totalMontagem = sizes.reduce((acc, size) => acc + (row.data[size]?.m || 0), 0);
  const totalEst = sizes.reduce((acc, size) => acc + (row.data[size]?.est || 0), 0);

  return {
    row,
    totalPA,
    totalPesponto,
    totalMontagem,
    totalEst,
  };
};

  const vendaGridForRow = (row) => {
    const grid = vendasDraft?.[row.ref]?.[row.cor] || vendas?.[row.ref]?.[row.cor];
    if (grid) return grid;
    return Object.fromEntries(sizes.map((size) => [size, 0]));
  };

  const controleRefs = useMemo(() => ["TODAS", ...Array.from(new Set(rows.map((row) => row.ref)))], [rows]);
  const controleCores = useMemo(() => {
    const base = rows.filter((row) => controleFiltroRef === "TODAS" || row.ref === controleFiltroRef);
    return ["TODAS", ...Array.from(new Set(base.map((row) => row.cor)))];
  }, [rows, controleFiltroRef]);

  const controleRows = useMemo(() => {
    return rowsNormalized.filter((row) => {
      if (controleFiltroRef !== "TODAS" && row.ref !== controleFiltroRef) return false;
      if (controleFiltroCor !== "TODAS" && row.cor !== controleFiltroCor) return false;
      return true;
    });
  }, [rowsNormalized, controleFiltroRef, controleFiltroCor]);

  const visibleSizesControle = useMemo(() => {
    if (controleFiltroNumero === "TODAS") return sizes;
    return [Number(controleFiltroNumero)];
  }, [controleFiltroNumero]);

  const sortedRowsByRefCor = useMemo(() => {
    return [...rowsNormalized].sort((a, b) => {
      const refCompare = String(a.ref).localeCompare(String(b.ref), "pt-BR", { numeric: true, sensitivity: "base" });
      if (refCompare !== 0) return refCompare;
      return String(a.cor).localeCompare(String(b.cor), "pt-BR", { sensitivity: "base" });
    });
  }, [rowsNormalized]);

  const pespontoFinalizadosHistorico = useMemo(() => {
    const ts = (lanc) => {
      const s = lanc?.dataFinalizacao || lanc?.dataLancamento || "";
      const p = String(s).split("/");
      if (p.length !== 3) return 0;
      const [dd, mm, yyyy] = p;
      const t = new Date(`${yyyy}-${mm}-${dd}T12:00:00`).getTime();
      return Number.isNaN(t) ? 0 : t;
    };
    return [...pespontoLancamentos]
      .filter((l) => String(l.status || "").trim() === "Finalizado")
      .sort((a, b) => ts(b) - ts(a));
  }, [pespontoLancamentos]);

  const metrics = useMemo(() => {
    let criticos = 0;
    let atencaoPA = 0;
    let atencaoProd = 0;
    let ok = 0;
    let costura = 0;
    rowsNormalized.forEach((row) => {
      sizes.forEach((size) => {
        const item = row.data[size] || { pa: 0, est: 0, m: 0, p: 0 };
        costura += item.est;
        const minimo = minimos?.[row.ref]?.[row.cor]?.[size] || { pa: 0, prod: 0 };
        const st = statusFor(item, minimo);
        if (st === "CRÍTICO") criticos += 1;
        else if (st === "ATENÇÃO PA") atencaoPA += 1;
        else if (st === "ATENÇÃO PROD") atencaoProd += 1;
        else ok += 1;
      });
    });
    return { criticos, atencaoPA, atencaoProd, ok, costura };
  }, [rowsNormalized, minimos]);

  const suggestions = useMemo(
  () => buildSuggestions(rowsNormalized, minimos, vendas, tempoProducao),
  [rowsNormalized, minimos, vendas, tempoProducao]
);

const topVendasCadastroOpcoes = useMemo(
  () =>
    sortedRowsByRefCor
      .map((row) => {
        const vendaGrid = vendasDraft?.[row.ref]?.[row.cor] || vendas?.[row.ref]?.[row.cor] || {};
        const vendaTotal = sizes.reduce((acc, size) => acc + (Number(vendaGrid[size]) || 0), 0);
        return {
          key: `${row.ref}__${row.cor}`,
          ref: row.ref,
          cor: row.cor,
          vendaTotal,
        };
      })
      .sort((a, b) => b.vendaTotal - a.vendaTotal || a.ref.localeCompare(b.ref, "pt-BR") || a.cor.localeCompare(b.cor, "pt-BR")),
  [sortedRowsByRefCor, vendasDraft, vendas]
);

const limiteFichaPesponto = Math.max(1, Number(capacidadePespontoDia) || 396);
const limiteFichaMontagem = Math.max(1, Number(capacidadeMontagemDia) || 396);

const fichasMontagem = useMemo(
  () =>
    (suggestions.montagem || []).flatMap((item) =>
      splitIntoFichas(item.sizes, limiteFichaMontagem).map((ficha, index) => ({
        ...ficha,
        nome: `${item.ref} - ${item.cor} - Ficha ${index + 1}`,
        ref: item.ref,
        cor: item.cor,
        tipo: item.tipo,
        prioridade: item.prioridade,
        urgente: item.urgente,
        recuperacao: item.recuperacao,
      }))
    ),
  [suggestions, limiteFichaMontagem]
);

const fichasPesponto = useMemo(
  () =>
    (suggestions.pesponto || []).flatMap((item) =>
      splitIntoFichas(item.sizes, limiteFichaPesponto).map((ficha, index) => ({
        ...ficha,
        nome: `${item.ref} - ${item.cor} - Ficha ${index + 1}`,
        ref: item.ref,
        cor: item.cor,
        tipo: item.tipo,
        prioridade: item.prioridade,
        urgente: item.urgente,
        recuperacao: item.recuperacao,
      }))
    ),
  [suggestions, limiteFichaPesponto]
);

const programacaoPesponto = useMemo(
  () =>
    buildProgramacaoPeriodo(
      fichasPesponto,
      suggestions.pesponto,
      Number(capacidadePespontoDia) || 396,
      programacaoDias,
      "Pesponto",
      { pctCapacidadeTop: programacaoReservaTopPct, topN: programacaoTopN, topMode: programacaoTopModo, topManualKeys: programacaoTopManualKeys }
    ),
  [fichasPesponto, suggestions, programacaoDias, capacidadePespontoDia, programacaoReservaTopPct, programacaoTopN, programacaoTopModo, programacaoTopManualKeys]
);

const programacaoMontagem = useMemo(
  () =>
    buildProgramacaoPeriodo(
      fichasMontagem,
      suggestions.montagem,
      Number(capacidadeMontagemDia) || 396,
      programacaoDias,
      "Montagem",
      { pctCapacidadeTop: programacaoReservaTopPct, topN: programacaoTopN, topMode: programacaoTopModo, topManualKeys: programacaoTopManualKeys }
    ),
  [fichasMontagem, suggestions, programacaoDias, capacidadeMontagemDia, programacaoReservaTopPct, programacaoTopN, programacaoTopModo, programacaoTopManualKeys]
);

  useEffect(() => {
    if (active !== "Programação do Dia") return;
    const capP = Number(capacidadePespontoDia) || 396;
    const capM = Number(capacidadeMontagemDia) || 396;
    const reservaPct = Math.min(100, Math.max(0, Number(programacaoReservaTopPct) || 0));
    const topN = Math.max(1, Math.min(200, Number(programacaoTopN) || 10));
    const topMode = programacaoTopModo === "manual" ? "manual" : "auto";
    const topManualSignature = [...new Set(programacaoTopManualKeys)].sort((a, b) => String(a).localeCompare(String(b), "pt-BR")).join("|");
    setProgramacaoPlanoCongelado((prev) => {
      if (prev != null) {
        const configOk =
          prev.dias === programacaoDias &&
          prev.capP === capP &&
          prev.capM === capM &&
          prev.reservaTopPct === reservaPct &&
          prev.topN === topN &&
          prev.topMode === topMode &&
          prev.topManualSignature === topManualSignature;
        if (configOk) return prev;
      }
      return {
        pesponto: deepClonePlain(programacaoPesponto),
        montagem: deepClonePlain(programacaoMontagem),
        dias: programacaoDias,
        capP,
        capM,
        reservaTopPct: reservaPct,
        topN,
        topMode,
        topManualSignature,
      };
    });
  }, [
    active,
    programacaoDias,
    capacidadePespontoDia,
    capacidadeMontagemDia,
    programacaoReservaTopPct,
    programacaoTopN,
    programacaoTopModo,
    programacaoTopManualKeys,
    programacaoPesponto,
    programacaoMontagem,
  ]);

  const recalcularProgramacaoCongelada = useCallback(() => {
    const capP = Number(capacidadePespontoDia) || 396;
    const capM = Number(capacidadeMontagemDia) || 396;
    const reservaPct = Math.min(100, Math.max(0, Number(programacaoReservaTopPct) || 0));
    const topN = Math.max(1, Math.min(200, Number(programacaoTopN) || 10));
    const topMode = programacaoTopModo === "manual" ? "manual" : "auto";
    const topManualSignature = [...new Set(programacaoTopManualKeys)].sort((a, b) => String(a).localeCompare(String(b), "pt-BR")).join("|");
    setProgramacaoPlanoCongelado({
      pesponto: deepClonePlain(programacaoPesponto),
      montagem: deepClonePlain(programacaoMontagem),
      dias: programacaoDias,
      capP,
      capM,
      reservaTopPct: reservaPct,
      topN,
      topMode,
      topManualSignature,
    });
    setProgramacaoFichaSelecao({});
  }, [
    programacaoPesponto,
    programacaoMontagem,
    programacaoDias,
    capacidadePespontoDia,
    capacidadeMontagemDia,
    programacaoReservaTopPct,
    programacaoTopN,
    programacaoTopModo,
    programacaoTopManualKeys,
  ]);

  const dashboardData = useMemo(() => {
    const tempoTotal = (Number(tempoProducao?.pesponto) || 0) + (Number(tempoProducao?.montagem) || 0);

    const totalPA = rowsNormalized.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.pa || 0), 0), 0);
    const totalEst = rowsNormalized.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.est || 0), 0), 0);
    const totalMontagemAtual = rowsNormalized.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.m || 0), 0), 0);
    const totalPespontoAtual = rowsNormalized.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.p || 0), 0), 0);
    const vendaMensalTotal = rowsNormalized.reduce((acc, row) => {
      const vendasRow = vendas?.[row.ref]?.[row.cor] || {};
      return acc + sizes.reduce((sum, size) => sum + (Number(vendasRow[size]) || 0), 0);
    }, 0);
    const vendaDiariaTotal = vendaMensalTotal / 30;

    const hoje = new Date();
    const feriadosLista = parseFeriadosText(feriadosTexto).filter((dataBr) => {
      const parsed = parseDateBrToDate(dataBr);
      return parsed && parsed.getMonth() === hoje.getMonth() && parsed.getFullYear() === hoje.getFullYear();
    });
    const diasUteisNoMes = contarDiasUteisDoMesAteHoje(hoje, feriadosLista);
    const diasUteisSafe = Math.max(1, diasUteisNoMes);
    const mesAtual = hoje.getMonth();
    const anoAtual = hoje.getFullYear();
    const hojeTexto = hoje.toLocaleDateString("pt-BR");

    const finalizadosMesPesponto = pespontoLancamentos.filter((item) => {
      if (item.status !== "Finalizado" || !item.dataFinalizacao) return false;
      const data = parseDateBrToDate(item.dataFinalizacao);
      return data && data.getMonth() === mesAtual && data.getFullYear() === anoAtual;
    });

    const finalizadosMesMontagem = montagemLancamentos.filter((item) => {
      if (item.status !== "Finalizado" || !item.dataFinalizacao) return false;
      const data = parseDateBrToDate(item.dataFinalizacao);
      return data && data.getMonth() === mesAtual && data.getFullYear() === anoAtual;
    });

    const totalFinalizadoMesPesponto = finalizadosMesPesponto.reduce((acc, item) => acc + (item.total || 0), 0);
    const totalFinalizadoMesMontagem = finalizadosMesMontagem.reduce((acc, item) => acc + (item.total || 0), 0);
    const mediaFinalizadaPesponto = totalFinalizadoMesPesponto / diasUteisSafe;
    const mediaFinalizadaMontagem = totalFinalizadoMesMontagem / diasUteisSafe;

    const finalizadoHojePesponto = finalizadosMesPesponto
      .filter((item) => item.dataFinalizacao === hojeTexto)
      .reduce((acc, item) => acc + (item.total || 0), 0);
    const finalizadoHojeMontagem = finalizadosMesMontagem
      .filter((item) => item.dataFinalizacao === hojeTexto)
      .reduce((acc, item) => acc + (item.total || 0), 0);

    const menorCobertura = rowsNormalized
      .flatMap((row) =>
        sizes.map((size) => {
          const vendaMes = Number(vendas?.[row.ref]?.[row.cor]?.[size]) || 0;
          const cobertura = coberturaDias(row.data[size]?.pa || 0, vendaMes);
          return { ref: row.ref, cor: row.cor, size, cobertura, pa: row.data[size]?.pa || 0, vendaMes };
        })
      )
      .filter((item) => item.cobertura !== null)
      .sort((a, b) => a.cobertura - b.cobertura)
      .slice(0, 5);

    const topModelos = rowsNormalized
      .map((row) => {
        const vendasRow = vendas?.[row.ref]?.[row.cor] || {};
        const totalVendido = sizes.reduce((acc, size) => acc + (Number(vendasRow[size]) || 0), 0);
        const totalPARef = sizes.reduce((acc, size) => acc + (row.data[size]?.pa || 0), 0);
        return { ref: row.ref, cor: row.cor, totalVendido, totalPA: totalPARef };
      })
      .sort((a, b) => b.totalVendido - a.totalVendido || a.totalPA - b.totalPA)
      .slice(0, 5);

    const programacaoHojePesponto = programacaoPesponto?.diasProgramados?.[0];
    const programacaoHojeMontagem = programacaoMontagem?.diasProgramados?.[0];

    return {
      tempoTotal,
      totalPA,
      totalEst,
      totalMontagemAtual,
      totalPespontoAtual,
      vendaMensalTotal,
      vendaDiariaTotal,
      hoje,
      hojeTexto,
      feriadosLista,
      diasUteisNoMes,
      diasUteisSafe,
      mediaFinalizadaPesponto,
      mediaFinalizadaMontagem,
      finalizadoHojePesponto,
      finalizadoHojeMontagem,
      menorCobertura,
      topModelos,
      programacaoHojePesponto,
      programacaoHojeMontagem,
    };
  }, [
    rowsNormalized,
    vendas,
    tempoProducao,
    feriadosTexto,
    pespontoLancamentos,
    montagemLancamentos,
    programacaoPesponto,
    programacaoMontagem,
  ]);

  const irControleGeral = useCallback((ref, cor, numero) => {
    setControleFiltroRef(ref != null && ref !== "" ? ref : "TODAS");
    setControleFiltroCor(cor != null && cor !== "" ? cor : "TODAS");
    setControleFiltroNumero(numero != null && numero !== "" ? String(numero) : "TODAS");
    setActive("Controle Geral");
  }, []);

  useEffect(() => {
    setDraftMinimos(minimos);
  }, [minimos]);

  useEffect(() => {
    setVendasDraft(vendas);
  }, [vendas]);

useEffect(() => {
  const carregarDadosIniciais = async () => {
    const estoqueBanco = await carregarEstoqueDoBanco();
    const minimosBanco = await carregarMinimosDoBanco();
    const vendasBanco = await carregarVendasDoBanco();
    const movimentacoesBanco = await carregarMovimentacoesDoBanco();
    const configProducao = await carregarConfiguracoesProducaoDoBanco();

    if (estoqueBanco) {
      setRows(estoqueBanco);
    }

    if (minimosBanco) {
      setMinimos(minimosBanco);
      setDraftMinimos(minimosBanco);
    }

    if (vendasBanco) {
      setVendas(vendasBanco);
      setVendasDraft(vendasBanco);
    }

    if (movimentacoesBanco) {
      setPespontoLancamentos(movimentacoesBanco.pesponto || []);
      setMontagemLancamentos(movimentacoesBanco.montagem || []);
      setAjustesEst(movimentacoesBanco.ajustesEst || []);
    }

    if (configProducao) {
      setCapacidadePespontoDia(Number(configProducao.capacidade_pesponto_dia) || 396);
      setCapacidadeMontagemDia(Number(configProducao.capacidade_montagem_dia) || 396);

      setTempoProducao({
        pesponto: Number(configProducao.dias_pesponto) || 3,
        montagem: Number(configProducao.dias_montagem) || 2,
      });

      setTempoProducaoDraft({
        pesponto: Number(configProducao.dias_pesponto) || 3,
        montagem: Number(configProducao.dias_montagem) || 2,
      });

      if (Object.prototype.hasOwnProperty.call(configProducao, "reserva_top_pct")) {
        setProgramacaoReservaTopPct(Math.min(100, Math.max(0, Number(configProducao.reserva_top_pct) || 0)));
      }
      if (Object.prototype.hasOwnProperty.call(configProducao, "top_n")) {
        setProgramacaoTopN(Math.min(200, Math.max(1, Math.round(Number(configProducao.top_n) || 10))));
      }
      if (Object.prototype.hasOwnProperty.call(configProducao, "top_mode")) {
        setProgramacaoTopModo(String(configProducao.top_mode).trim().toLowerCase() === "manual" ? "manual" : "auto");
      }
      if (Object.prototype.hasOwnProperty.call(configProducao, "top_manual_keys")) {
        setProgramacaoTopManualKeys(parseTopManualKeysFromDb(configProducao.top_manual_keys));
      }
      if (Object.prototype.hasOwnProperty.call(configProducao, "valor_par_weverton")) {
        setProgramacaoValoresTerceiros((curr) => ({
          ...curr,
          weverton: String(configProducao.valor_par_weverton ?? ""),
        }));
      }
      if (Object.prototype.hasOwnProperty.call(configProducao, "valor_par_romulo")) {
        setProgramacaoValoresTerceiros((curr) => ({
          ...curr,
          romulo: String(configProducao.valor_par_romulo ?? ""),
        }));
      }
    }
  };

  carregarDadosIniciais();
}, []);

useEffect(() => {
  if (!refs.length) return;

  setPespontoForm((current) => {
    const isValid = refs.includes(`${current.ref}__${current.cor}`);
    if (isValid) return current;
    return { ...current, ref: firstRef, cor: firstCor };
  });

  setMontagemForm((current) => {
    const isValid = refs.includes(`${current.ref}__${current.cor}`);
    if (isValid) return current;
    return { ...current, ref: firstRef, cor: firstCor };
  });
}, [refs, firstRef, firstCor]);

  useEffect(() => {
    if (!printRelatorioData) return undefined;

    const runPrint = () => {
      window.print();
    };

    const clearAfterPrint = () => {
      setPrintRelatorioData(null);
    };

    const timer = window.setTimeout(runPrint, 200);
    window.addEventListener("afterprint", clearAfterPrint);

    return () => {
      window.clearTimeout(timer);
      window.removeEventListener("afterprint", clearAfterPrint);
    };
  }, [printRelatorioData]);

  useEffect(() => {
    setProgramacaoFichaSelecao({});
  }, [programacaoSubAba, programacaoModoVisual, programacaoDias, programacaoReservaTopPct, programacaoTopN, programacaoTopModo, programacaoTopManualKeys]);

  useEffect(() => {
    try {
      localStorage.setItem(PROG_RESERVA_TOP_PCT_KEY, String(programacaoReservaTopPct));
    } catch (_) {
      /* ignore */
    }
  }, [programacaoReservaTopPct]);

  useEffect(() => {
    try {
      localStorage.setItem(PROG_TOP_N_KEY, String(programacaoTopN));
    } catch (_) {
      /* ignore */
    }
  }, [programacaoTopN]);

  useEffect(() => {
    try {
      localStorage.setItem(PROG_TOP_MODE_KEY, String(programacaoTopModo));
    } catch (_) {
      /* ignore */
    }
  }, [programacaoTopModo]);

  useEffect(() => {
    try {
      localStorage.setItem(PROG_TOP_MANUAL_KEYS_KEY, JSON.stringify(programacaoTopManualKeys));
    } catch (_) {
      /* ignore */
    }
  }, [programacaoTopManualKeys]);

  useEffect(() => {
    try {
      localStorage.setItem(
        PROG_VALORES_PAGAMENTO_KEY,
        JSON.stringify({
          weverton: String(programacaoValoresTerceiros.weverton ?? ""),
          romulo: String(programacaoValoresTerceiros.romulo ?? ""),
        })
      );
    } catch (_) {
      /* ignore */
    }
  }, [programacaoValoresTerceiros]);

  useEffect(() => {
    setMovListPage((prev) => {
      const next = { ...prev };
      let changed = false;
      const fix = (key, len) => {
        const tp = Math.max(1, Math.ceil(len / 15));
        if ((prev[key] || 1) > tp) {
          next[key] = tp;
          changed = true;
        }
      };
      fix("Pesponto", pespontoLancamentos.length);
      fix("Montagem", montagemLancamentos.length);
      return changed ? next : prev;
    });
  }, [pespontoLancamentos.length, montagemLancamentos.length]);

  const mapRowsToEstoquePayload = (rowsData) =>
    rowsData
      .map((row) =>
        Object.entries(row.data).map(([numero, valores]) => ({
          ref: row.ref,
          cor: row.cor,
          numero: Number(numero),
          pa: valores.pa || 0,
          est: valores.est || 0,
          m: valores.m || 0,
          p: valores.p || 0,
        }))
      )
      .flat();

  const persistRowsToSupabase = async (rowsData) => {
    const { error } = await salvarEstoqueNoBanco(mapRowsToEstoquePayload(rowsData));
    if (error) {
      console.error("persistRowsToSupabase:", error);
      return { ok: false, error };
    }
    return { ok: true, error: null };
  };

  const normalizeLancamentoItems = (lancamento) => {
    const raw = Array.isArray(lancamento?.items) ? lancamento.items : [];
    return raw.filter((item) => Number(item?.size) > 0 && Number(item?.qtd) > 0);
  };

  const cloneRowDataDeep = (row) =>
    Object.fromEntries(
      sizes.map((s) => [
        s,
        {
          pa: Number(row.data[s]?.pa) || 0,
          est: Number(row.data[s]?.est) || 0,
          m: Number(row.data[s]?.m) || 0,
          p: Number(row.data[s]?.p) || 0,
        },
      ])
    );

  /** Aplica finalização com validação de saldo (p ou m) e retorna erro se inválido. */
  const computeFinalizacaoNextRows = (rowsData, tipo, lancamentosAbertos) => {
    const alvo = lancamentosAbertos
      .map((l) => ({ ...l, items: normalizeLancamentoItems(l) }))
      .filter((l) => l.items.length > 0);

    if (!alvo.length) {
      return {
        error:
          "Não há itens válidos nesta programação (grade vazia ou numerações inválidas). Lance novamente ou verifique o Supabase.",
      };
    }

    const nextRows = rowsData.map((row) => ({
      ...row,
      data: cloneRowDataDeep(row),
    }));

    const getMutableRow = (ref, cor) => nextRows.find((r) => r.ref === ref && r.cor === cor);

    for (const lancamento of alvo) {
      const row = getMutableRow(lancamento.ref, lancamento.cor);
      if (!row) {
        return {
          error: `Produto não encontrado no estoque: ${lancamento.ref} • ${lancamento.cor}.`,
        };
      }

      for (const item of lancamento.items) {
        const atual = row.data[item.size] || { pa: 0, est: 0, m: 0, p: 0 };

        if (tipo === "Pesponto") {
          const disponivel = atual.p || 0;
          if (disponivel < item.qtd) {
            return {
              error: `Não dá para finalizar: no pesponto faltam pares em ${lancamento.ref} • ${lancamento.cor}, num. ${item.size} (em aberto: ${disponivel}, necessário: ${item.qtd}).`,
            };
          }
          row.data[item.size] = {
            ...atual,
            p: disponivel - item.qtd,
            est: (atual.est || 0) + item.qtd,
          };
        } else {
          const disponivel = atual.m || 0;
          if (disponivel < item.qtd) {
            return {
              error: `Não dá para finalizar: na montagem faltam pares em ${lancamento.ref} • ${lancamento.cor}, num. ${item.size} (em aberto: ${disponivel}, necessário: ${item.qtd}).`,
            };
          }
          row.data[item.size] = {
            ...atual,
            m: disponivel - item.qtd,
            pa: (atual.pa || 0) + item.qtd,
          };
        }
      }
    }

    return { nextRows };
  };

  const applyLancamentoDeltaToRows = (rowsData, tipo, lancamento, direction = "revert") => {
    const multiplier = direction === "revert" ? -1 : 1;
    return rowsData.map((row) => {
      if (row.ref !== lancamento.ref || row.cor !== lancamento.cor) return row;
      const nextData = { ...row.data };
      lancamento.items.forEach((item) => {
        const atual = nextData[item.size] || { pa: 0, est: 0, m: 0, p: 0 };
        if (tipo === "Pesponto") {
          nextData[item.size] = {
            ...atual,
            p: Math.max(0, (atual.p || 0) + item.qtd * multiplier),
          };
        } else {
          nextData[item.size] = {
            ...atual,
            est: Math.max(0, (atual.est || 0) - item.qtd * multiplier),
            m: Math.max(0, (atual.m || 0) + item.qtd * multiplier),
          };
        }
      });
      return { ...row, data: nextData };
    });
  };

  const executeMov = async (tipo, form, force = false, progFichaStorageKey) => {
    console.log("EXECUTE MOV FOI CHAMADO", { tipo, form });
    const items = sizes
      .map((size) => ({ size, qtd: Number(form.grid[size]) || 0 }))
      .filter((x) => x.qtd > 0);

    const programacaoNome = String(form.programacao || "").trim();
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const refNorm = String(form.ref || "").trim();
    const corNorm = String(form.cor || "").trim();
    const programacaoDuplicada = programacaoNome
      ? source.some(
          (item) =>
            String(item.programacao || "").trim().toUpperCase() === programacaoNome.toUpperCase() &&
            String(item.ref || "").trim() === refNorm &&
            String(item.cor || "").trim() === corNorm
        )
      : false;

    if (!items.length) {
      setMovError((curr) => ({ ...curr, [tipo]: "" }));
      return;
    }

    if (!programacaoNome) {
      setMovError((curr) => ({ ...curr, [tipo]: "Informe o nome da programação antes de lançar." }));
      return;
    }

    if (programacaoDuplicada) {
      setMovError((curr) => ({
        ...curr,
        [tipo]: `Já existe uma programação com o nome "${programacaoNome}" para esta ref e cor em ${tipo}. Use outro nome.`,
      }));
      return;
    }

    const invalidos =
      tipo === "Pesponto" || tipo === "Montagem"
        ? items.filter((item) => item.qtd % 12 !== 0)
        : [];

    const totalLancamento = items.reduce((acc, item) => acc + item.qtd, 0);
    const excedeLimite = (tipo === "Pesponto" || tipo === "Montagem") && totalLancamento > 396;

    if ((tipo === "Pesponto" || tipo === "Montagem") && (invalidos.length || excedeLimite) && !force) {
      setConfirmMov({ tipo, form, progFichaStorageKey });
      return;
    }

    setMovError((curr) => ({ ...curr, [tipo]: "" }));

 const nextRows = rows.map((row) => {
  if (row.ref !== form.ref || row.cor !== form.cor) return row;

  const nextData = { ...row.data };

  items.forEach((item) => {
    const atual = nextData[item.size] || { pa: 0, est: 0, m: 0, p: 0 };

    if (tipo === "Pesponto") {
      nextData[item.size] = {
        ...atual,
        p: (atual.p || 0) + item.qtd,
      };
    } else {
      nextData[item.size] = {
        ...atual,
        est: Math.max(0, (atual.est || 0) - item.qtd),
        m: (atual.m || 0) + item.qtd,
      };
    }
  });

  return {
    ...row,
    data: nextData,
  };
});

const persistLaunch = await persistRowsToSupabase(nextRows);
    if (!persistLaunch.ok) {
      setMovError((curr) => ({
        ...curr,
        [tipo]: `Erro ao salvar estoque no Supabase: ${persistLaunch.error?.message || "tente novamente"}.`,
      }));
      return;
    }

    setRows(nextRows);

    const payload = {
      id: `${tipo}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      programacao: programacaoNome,
      ref: form.ref,
      cor: form.cor,
      items,
      total: totalLancamento,
      status: "Em aberto",
      dataLancamento: new Date().toLocaleDateString("pt-BR"),
    };

    console.log("PAYLOAD CRIADO", payload);

    const movsParaSalvar = items.map((item) => ({
      tipo,
      ref: form.ref,
      cor: form.cor,
      numero: item.size,
      quantidade: item.qtd,
      programacao: programacaoNome,
      status: "Em aberto",
    }));

    if (tipo === "Pesponto") {
      setPespontoLancamentos((c) => [payload, ...c]);
      setPespontoForm((f) => ({ ...f, grid: makeEmptyGrid() }));
    } else {
      setMontagemLancamentos((c) => [payload, ...c]);
      setMontagemForm((f) => ({ ...f, grid: makeEmptyGrid() }));
    }

    console.log("ANTES DE SALVAR MOVIMENTACAO", movsParaSalvar);
    const movSave = await salvarMovimentacao(movsParaSalvar);
    console.log("DEPOIS DE SALVAR");
    if (!movSave?.error && progFichaStorageKey) {
      addLaunchedProgFichaKey(progFichaStorageKey);
    }
    setConfirmMov(null);
    setLancarFichaDaProgramacao(null);
  };

  const getMovErrorMessage = (tipo, form) => {
    if (tipo !== "Pesponto" && tipo !== "Montagem") return "";

    const programacaoNome = String(form.programacao || "").trim();
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const refNorm = String(form.ref || "").trim();
    const corNorm = String(form.cor || "").trim();
    const programacaoDuplicada = programacaoNome
      ? source.some(
          (item) =>
            String(item.programacao || "").trim().toUpperCase() === programacaoNome.toUpperCase() &&
            String(item.ref || "").trim() === refNorm &&
            String(item.cor || "").trim() === corNorm
        )
      : false;

    const invalidos = sizes
      .map((size) => ({ size, qtd: Number(form.grid[size]) || 0 }))
      .filter((item) => item.qtd > 0 && item.qtd % 12 !== 0);

    const totalLancamento = sizes.reduce((acc, size) => acc + (Number(form.grid[size]) || 0), 0);
    const mensagens = [];

    if (!programacaoNome) {
      mensagens.push("informe o nome da programação");
    }

    if (programacaoDuplicada) {
      mensagens.push(`já existe uma programação com esse nome para esta ref e cor em ${tipo}`);
    }

    if (invalidos.length) {
      const lista = invalidos.map((item) => `${item.size} (${item.qtd})`).join(", ");
      mensagens.push(`no ${tipo}, o padrão é lançar múltiplos de 12. Fora da regra: ${lista}`);
    }

    if (totalLancamento > 396) {
      mensagens.push(`o total do lançamento está em ${totalLancamento} pares e não pode passar de 396`);
    }

    if (!mensagens.length) return "";

    return `Atenção: ${mensagens.join(". ")}.`;
  };

  const excluirMovimentacoesNoBanco = async (tipo, alvo, { suppressAlert = false } = {}) => {
    try {
      const { error } = await supabase
        .from("movimentacoes")
        .delete()
        .eq("tipo", tipo)
        .eq("programacao", alvo.programacao)
        .eq("ref", alvo.ref)
        .eq("cor", alvo.cor)
        .eq("status", "Em aberto");

      if (error) {
        console.log("ERRO AO EXCLUIR MOVIMENTACAO:", error);
        if (!suppressAlert) alert("Erro ao excluir movimentação no banco.");
        return false;
      }

      console.log("MOVIMENTACOES EXCLUIDAS DO BANCO");
      return true;
    } catch (err) {
      console.log("ERRO GERAL AO EXCLUIR MOVIMENTACAO:", err);
      if (!suppressAlert) alert("Erro geral ao excluir movimentação.");
      return false;
    }
  };

  const deleteLancamento = (tipo, lancamentoId) => {
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const alvo = source.find((item) => item.id === lancamentoId);
    if (!alvo || alvo.status === "Finalizado") return;

    setConfirmAction({
      kind: "delete",
      tipo,
      lancamentoId,
      titulo: `Excluir programação em ${tipo}`,
      mensagem: `Deseja realmente excluir a programação "${alvo.programacao}"? Essa ação não pode ser desfeita.`,
    });
  };

  const startEditLancamento = async (tipo, lancamento) => {
    if (lancamento.status === "Finalizado") return;
    const nextRows = applyLancamentoDeltaToRows(rows, tipo, lancamento, "revert");
    const persist = await persistRowsToSupabase(nextRows);
    if (!persist.ok) {
      alert(
        `Não foi possível salvar o estoque ao preparar a edição. Tente novamente.\n\n${persist.error?.message || persist.error || ""}`
      );
      return;
    }

    const okDb = await excluirMovimentacoesNoBanco(tipo, lancamento, { suppressAlert: true });
    if (!okDb) {
      await persistRowsToSupabase(rows);
      alert(
        "Não foi possível remover o lançamento antigo no Supabase. O estoque foi mantido como antes. Tente novamente."
      );
      return;
    }

    setRows(nextRows);
    setEditingMov({
      ...lancamento,
      tipo,
      grid: sizes.reduce((acc, size) => {
        const found = lancamento.items.find((item) => item.size === size);
        acc[size] = found?.qtd || 0;
        return acc;
      }, {}),
    });

    if (tipo === "Pesponto") {
      setPespontoLancamentos((curr) => curr.filter((item) => item.id !== lancamento.id));
      setPespontoForm({
        ref: lancamento.ref,
        cor: lancamento.cor,
        programacao: lancamento.programacao,
        grid: sizes.reduce((acc, size) => {
          const found = lancamento.items.find((item) => item.size === size);
          acc[size] = found?.qtd || 0;
          return acc;
        }, {}),
      });
      setActive("Pesponto");
    } else {
      setMontagemLancamentos((curr) => curr.filter((item) => item.id !== lancamento.id));
      setMontagemForm({
        ref: lancamento.ref,
        cor: lancamento.cor,
        programacao: lancamento.programacao,
        grid: sizes.reduce((acc, size) => {
          const found = lancamento.items.find((item) => item.size === size);
          acc[size] = found?.qtd || 0;
          return acc;
        }, {}),
      });
      setActive("Montagem");
    }
  };

  const finalizarProgramacao = (tipo, programacao) => {
    setConfirmAction({
      kind: "finalizar",
      tipo,
      programacao,
      titulo: `Finalizar programação em ${tipo}`,
      mensagem: `Deseja realmente finalizar a programação "${programacao}"? Ao confirmar, o estoque será atualizado.`,
    });
  };

  const confirmDeleteLancamento = async ({ tipo, lancamentoId }) => {
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const alvo = source.find((item) => item.id === lancamentoId);

    if (!alvo || alvo.status === "Finalizado") {
      setConfirmAction(null);
      return;
    }

    const nextRows = applyLancamentoDeltaToRows(rows, tipo, alvo, "revert");
    const persist = await persistRowsToSupabase(nextRows);
    if (!persist.ok) {
      alert(
        `Não foi possível atualizar o estoque ao excluir. Nada foi removido.\n\n${persist.error?.message || persist.error || ""}`
      );
      return;
    }

    const ok = await excluirMovimentacoesNoBanco(tipo, alvo);
    if (!ok) {
      await persistRowsToSupabase(rows);
      return;
    }

    setRows(nextRows);

    if (tipo === "Pesponto") {
      setPespontoLancamentos((curr) => curr.filter((item) => item.id !== lancamentoId));
    } else {
      setMontagemLancamentos((curr) => curr.filter((item) => item.id !== lancamentoId));
    }

    setConfirmAction(null);
  };

  const confirmFinalizarProgramacao = async ({ tipo, programacao }) => {
    const lancamentos = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;

    const alvo = lancamentos.filter(
      (l) => l.programacao === programacao && l.status !== "Finalizado"
    );

    if (!alvo.length) {
      alert("Não há lançamentos em aberto para essa programação.");
      setConfirmAction(null);
      return;
    }

    const computed = computeFinalizacaoNextRows(rows, tipo, alvo);
    if (computed.error) {
      alert(computed.error);
      return;
    }

    const { nextRows } = computed;

    const persist = await persistRowsToSupabase(nextRows);
    if (!persist.ok) {
      alert(
        `Não foi possível salvar o estoque no Supabase. A Costura Pronta / PA não será atualizada até o salvamento funcionar.\n\n${persist.error?.message || persist.error || "Erro desconhecido"}`
      );
      return;
    }

    const { error: statusError } = await atualizarStatusMovimentacoesNoBanco(tipo, programacao);
    if (statusError) {
      const rollback = await persistRowsToSupabase(rows);
      if (!rollback.ok) {
        alert(
          `Estoque pode estar inconsistente: o status não foi atualizado e a reversão falhou. Verifique o Supabase.\n\n${rollback.error?.message || rollback.error}`
        );
      } else {
        alert(
          `Não foi possível marcar a programação como finalizada no banco. O estoque foi mantido como antes.\n\n${statusError.message || statusError}`
        );
      }
      return;
    }

    const dataFinalizacao = new Date().toLocaleDateString("pt-BR");

    setRows(nextRows);

    if (tipo === "Pesponto") {
      setPespontoLancamentos((curr) =>
        curr.map((l) =>
          l.programacao === programacao ? { ...l, status: "Finalizado", dataFinalizacao } : l
        )
      );
    } else {
      setMontagemLancamentos((curr) =>
        curr.map((l) =>
          l.programacao === programacao ? { ...l, status: "Finalizado", dataFinalizacao } : l
        )
      );
    }

    setConfirmAction(null);
  };

function parseGcmSheet(sheet) {
  const rowsSheet = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const resultado = [];

  const toText = (v) => String(v ?? "").trim();
  const upper = (v) => toText(v).toUpperCase();

  for (let i = 0; i < rowsSheet.length; i += 1) {
    const row = Array.isArray(rowsSheet[i]) ? rowsSheet[i] : [];
    const linha1 = row.map(toText);

    if (!linha1.length) continue;

    const cabecalho = upper(linha1[0]);

    // precisa ser algo tipo "BTCV010 - AZUL BB"
    if (!cabecalho.includes("-")) continue;
    if (cabecalho.startsWith("ESTOQUE")) continue;
    if (cabecalho.includes("QTDE")) continue;

    const partes = linha1[0].split("-");
    if (partes.length < 2) continue;

    const ref = toText(partes[0]).toUpperCase();
    const cor = toText(partes.slice(1).join("-")).toUpperCase();

    // tamanhos ficam na mesma linha a partir da coluna 2
    const tamanhos = linha1
      .slice(1)
      .map((v) => Number(v))
      .filter((n) => sizes.includes(n));

    if (!tamanhos.length) continue;

    // próxima linha tem que começar com ESTOQUE
    const prox = Array.isArray(rowsSheet[i + 1]) ? rowsSheet[i + 1] : [];
    const linha2 = prox.map(toText);

    if (!upper(linha2[0]).startsWith("ESTOQUE")) continue;

    const quantidades = linha2
      .slice(1)
      .map((v) => Number(v))
      .filter((n) => !Number.isNaN(n));

    const data = Object.fromEntries(sizes.map((s) => [s, 0]));

    tamanhos.forEach((size, idx) => {
      data[size] = quantidades[idx] || 0;
    });

    const total = sizes.reduce((acc, s) => acc + (data[s] || 0), 0);

    resultado.push({
      ref,
      cor,
      data,
      total,
    });
  }

  return resultado;
}

  const handleFileUpload = async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawText = XLSX.utils.sheet_to_txt(firstSheet);
    const parsed = parseGcmSheet(firstSheet);

    setImportText(rawText);
    setImportFileName(file.name);
    setImportPreview(parsed);
    setImportFeedback(`Arquivo carregado com ${parsed.length} bloco(s) do GCM.`);
  } catch (err) {
    console.log("ERRO AO LER GCM:", err);
    setImportFeedback("Não consegui ler esse arquivo.");
    setImportPreview([]);
  }
};



  const salvarProdutosNoBanco = async (produtos) => {
    try {
      const { data, error } = await supabase
        .from("produtos")
        .upsert(produtos, { onConflict: "ref,cor" })
        .select();

      console.log("PRODUTOS ENVIADOS:", produtos);
      console.log("RETORNO SUPABASE:", data);
      console.log("ERRO AO SALVAR NO SUPABASE:", error);

      if (!error) {
        console.log("PRODUTOS SALVOS NO SUPABASE");
      }

      return { data, error };
    } catch (err) {
      console.log("ERRO GERAL AO SALVAR PRODUTOS:", err);
      return { data: null, error: err };
    }
  };

  const salvarEstoqueNoBanco = async (estoque) => {
    try {
      const { data, error } = await supabase
        .from("estoque")
        .upsert(estoque, { onConflict: "ref,cor,numero" })
        .select();

      console.log("ESTOQUE ENVIADO:", estoque);
      console.log("RETORNO ESTOQUE:", data);
      console.log("ERRO ESTOQUE:", error);

      if (!error) {
        console.log("ESTOQUE SALVO NO SUPABASE");
      }

      return { data, error };
    } catch (err) {
      console.log("ERRO GERAL ESTOQUE:", err);
      return { data: null, error: err };
    }
  };

  const salvarMovimentacao = async (movs) => {
    try {
      const batchTs = new Date().toISOString();
      const movsComData = (movs || []).map((m) => ({
        ...m,
        data_lancamento: m.data_lancamento ?? batchTs,
      }));
      console.log("MOVIMENTACOES ENVIADAS:", movsComData);

      const { data, error } = await supabase
        .from("movimentacoes")
        .insert(movsComData)
        .select();

      console.log("RETORNO MOV:", data);
      console.log("ERRO MOV:", error);

      if (error) {
        alert("Erro ao salvar movimentação. Veja o console.");
      } else {
        alert("Movimentação salva no banco.");
      }

      return { data, error };
    } catch (err) {
      console.log("ERRO GERAL MOV:", err);
      alert("Erro geral ao salvar movimentação.");
      return { data: null, error: err };
    }
  };

const salvarMinimosNoBanco = async (minimosData) => {
  try {
    const linhas = [];

    Object.keys(minimosData || {}).forEach((ref) => {
      Object.keys(minimosData[ref] || {}).forEach((cor) => {
        Object.keys(minimosData[ref][cor] || {}).forEach((numero) => {
          const item = minimosData[ref][cor][numero] || {};
          linhas.push({
            ref,
            cor,
            numero: Number(numero),
            min_pa: Number(item.pa) || 0,
            min_prod: Number(item.prod) || 0,
          });
        });
      });
    });

    const { data, error } = await supabase
      .from("minimos")
      .upsert(linhas, { onConflict: "ref,cor,numero" })
      .select();

    console.log("MINIMOS ENVIADOS:", linhas);
    console.log("RETORNO MINIMOS:", data);
    console.log("ERRO MINIMOS:", error);

    return { data, error };
  } catch (err) {
    console.log("ERRO GERAL MINIMOS:", err);
    return { data: null, error: err };
  }
};

const salvarConfiguracoesProducaoNoBanco = async ({
  capacidadePespontoDia,
  capacidadeMontagemDia,
  diasPesponto,
  diasMontagem,
  reservaTopPct,
  topN,
  topMode,
  topManualKeys,
  valorParWeverton,
  valorParRomulo,
}) => {
  try {
    const atual = await carregarConfiguracoesProducaoDoBanco();
    const parsedWeverton = parseDecimalInput(valorParWeverton);
    const parsedRomulo = parseDecimalInput(valorParRomulo);
    const payloadBase = {
      capacidade_pesponto_dia: Number(capacidadePespontoDia) || 396,
      capacidade_montagem_dia: Number(capacidadeMontagemDia) || 396,
      dias_pesponto: Number(diasPesponto) || 3,
      dias_montagem: Number(diasMontagem) || 2,
    };
    const payload = {
      ...payloadBase,
      reserva_top_pct: Math.min(100, Math.max(0, Number(reservaTopPct) || 0)),
      top_n: Math.min(200, Math.max(1, Math.round(Number(topN) || 10))),
      top_mode: topMode === "manual" ? "manual" : "auto",
      top_manual_keys: [...new Set(Array.isArray(topManualKeys) ? topManualKeys : [])]
        .filter((x) => typeof x === "string" && x.includes("__")),
      valor_par_weverton: Number.isFinite(parsedWeverton) ? parsedWeverton : (Number(atual?.valor_par_weverton) || 0),
      valor_par_romulo: Number.isFinite(parsedRomulo) ? parsedRomulo : (Number(atual?.valor_par_romulo) || 0),
    };

    if (atual?.id) {
      let { data, error } = await supabase
        .from("configuracoes_producao")
        .update(payload)
        .eq("id", atual.id)
        .select()
        .single();
      if (error) {
        const msg = String(error?.message || "").toLowerCase();
        const missingTopColumn =
          msg.includes("reserva_top_pct") ||
          msg.includes("top_n") ||
          msg.includes("top_mode") ||
          msg.includes("top_manual_keys") ||
          msg.includes("valor_par_weverton") ||
          msg.includes("valor_par_romulo");
        if (missingTopColumn) {
          ({ data, error } = await supabase
            .from("configuracoes_producao")
            .update(payloadBase)
            .eq("id", atual.id)
            .select()
            .single());
        }
      }

      console.log("CONFIG PRODUCAO ATUALIZADA:", data);
      console.log("ERRO CONFIG PRODUCAO:", error);

      return { data, error };
    }

    let { data, error } = await supabase
      .from("configuracoes_producao")
      .insert(payload)
      .select()
      .single();
    if (error) {
      const msg = String(error?.message || "").toLowerCase();
      const missingTopColumn =
        msg.includes("reserva_top_pct") ||
        msg.includes("top_n") ||
        msg.includes("top_mode") ||
        msg.includes("top_manual_keys") ||
        msg.includes("valor_par_weverton") ||
        msg.includes("valor_par_romulo");
      if (missingTopColumn) {
        ({ data, error } = await supabase
          .from("configuracoes_producao")
          .insert(payloadBase)
          .select()
          .single());
      }
    }

    console.log("CONFIG PRODUCAO CRIADA:", data);
    console.log("ERRO CONFIG PRODUCAO:", error);

    return { data, error };
  } catch (err) {
    console.log("ERRO GERAL AO SALVAR CONFIG PRODUCAO:", err);
    return { data: null, error: err };
  }
};

const salvarVendasNoBanco = async (vendasData) => {
  try {
    const linhas = [];

    Object.keys(vendasData || {}).forEach((ref) => {
      Object.keys(vendasData[ref] || {}).forEach((cor) => {
        Object.keys(vendasData[ref][cor] || {}).forEach((numero) => {
          linhas.push({
            ref,
            cor,
            numero: Number(numero),
            qtd: Number(vendasData[ref][cor][numero]) || 0,
          });
        });
      });
    });

    const { data, error } = await supabase
      .from("vendas")
      .upsert(linhas, { onConflict: "ref,cor,numero" })
      .select();

    console.log("VENDAS ENVIADAS:", linhas);
    console.log("RETORNO VENDAS:", data);
    console.log("ERRO VENDAS:", error);

    return { data, error };
  } catch (err) {
    console.log("ERRO GERAL VENDAS:", err);
    return { data: null, error: err };
  }
};

const atualizarStatusMovimentacoesNoBanco = async (tipo, programacao) => {
  try {
    const { data, error } = await supabase
      .from("movimentacoes")
      .update({
        status: "Finalizado",
        data_finalizacao: new Date().toISOString(),
      })
      .eq("tipo", tipo)
      .eq("programacao", programacao)
      .eq("status", "Em aberto")
      .select();

    console.log("STATUS MOVIMENTACOES ATUALIZADO:", data);
    console.log("ERRO AO ATUALIZAR STATUS:", error);

    return { data, error };
  } catch (err) {
    console.log("ERRO GERAL AO ATUALIZAR STATUS DAS MOVIMENTACOES:", err);
    return { data: null, error: err };
  }
};

const carregarMinimosDoBanco = async () => {
  try {
    const { data, error } = await supabase
      .from("minimos")
      .select("*")
      .order("ref", { ascending: true })
      .order("cor", { ascending: true })
      .order("numero", { ascending: true });

    if (error) {
      console.log("ERRO AO CARREGAR MINIMOS:", error);
      return null;
    }

    if (!data || !data.length) {
      return null;
    }

    const estruturado = {};

    data.forEach((item) => {
      if (!estruturado[item.ref]) estruturado[item.ref] = {};
      if (!estruturado[item.ref][item.cor]) estruturado[item.ref][item.cor] = {};

      estruturado[item.ref][item.cor][item.numero] = {
        pa: Number(item.min_pa) || 0,
        prod: Number(item.min_prod) || 0,
      };
    });

    console.log("MINIMOS CARREGADOS:", estruturado);
    return estruturado;
  } catch (err) {
    console.log("ERRO GERAL AO CARREGAR MINIMOS:", err);
    return null;
  }
};

const carregarVendasDoBanco = async () => {
  try {
    const { data, error } = await supabase
      .from("vendas")
      .select("*")
      .order("ref", { ascending: true })
      .order("cor", { ascending: true })
      .order("numero", { ascending: true });

    if (error) {
      console.log("ERRO AO CARREGAR VENDAS:", error);
      return null;
    }

    if (!data || !data.length) {
      return null;
    }

    const estruturado = {};

    data.forEach((item) => {
      if (!estruturado[item.ref]) estruturado[item.ref] = {};
      if (!estruturado[item.ref][item.cor]) estruturado[item.ref][item.cor] = {};

      estruturado[item.ref][item.cor][item.numero] = Number(item.qtd) || 0;
    });

    console.log("VENDAS CARREGADAS:", estruturado);
    return estruturado;
  } catch (err) {
    console.log("ERRO GERAL AO CARREGAR VENDAS:", err);
    return null;
  }
};

const carregarEstoqueDoBanco = async () => {
  try {
    const { data, error } = await supabase
      .from("estoque")
      .select("*")
      .order("ref", { ascending: true })
      .order("cor", { ascending: true })
      .order("numero", { ascending: true });

    if (error) {
      console.log("ERRO AO CARREGAR ESTOQUE:", error);
      return null;
    }

    if (!data || !data.length) {
      return null;
    }

    const agrupado = {};

    data.forEach((item) => {
      const key = `${item.ref}__${item.cor}`;

      if (!agrupado[key]) {
        agrupado[key] = {
          ref: item.ref,
          cor: item.cor,
          data: {},
        };
      }

      agrupado[key].data[item.numero] = {
        pa: Number(item.pa) || 0,
        est: Number(item.est) || 0,
        m: Number(item.m) || 0,
        p: Number(item.p) || 0,
      };
    });

    const rowsEstruturadas = Object.values(agrupado).map((row) => ({
      ref: row.ref,
      cor: row.cor,
      data: normalizeProductData(row.data),
    }));

    console.log("ESTOQUE CARREGADO:", rowsEstruturadas);
    return rowsEstruturadas;
  } catch (err) {
    console.log("ERRO GERAL AO CARREGAR ESTOQUE:", err);
    return null;
  }
};

const carregarMovimentacoesDoBanco = async () => {
  try {
    const { data, error } = await supabase
      .from("movimentacoes")
      .select("*")
      .order("data_lancamento", { ascending: false });

    if (error) {
      console.log("ERRO AO CARREGAR MOVIMENTACOES:", error);
      return { pesponto: [], montagem: [], ajustesEst: [] };
    }

    if (!data || !data.length) {
      return { pesponto: [], montagem: [], ajustesEst: [] };
    }

    const agrupar = (lista, tipo) => {
      const agrupado = {};

      lista
        .filter((item) => String(item.tipo || "").trim() === tipo)
        .forEach((item, index) => {
          const programacao = String(item.programacao || "Sem programação").trim();
          const ref = String(item.ref || "").trim();
          const cor = String(item.cor || "").trim();
          const status = String(item.status || "Em aberto").trim();
          /** Mesma programação/ref/cor/status = um lançamento; não usar data (cada linha do insert pode ter ms diferentes e quebrar o grupo). */
          const key = `${tipo}__${programacao}__${ref}__${cor}__${status}`;

          if (!agrupado[key]) {
            agrupado[key] = {
              id: String(item.id || `${key}__${index}`),
              programacao,
              ref,
              cor,
              items: [],
              total: 0,
              status,
              dataLancamento: item.data_lancamento
                ? new Date(item.data_lancamento).toLocaleDateString("pt-BR")
                : "",
              dataFinalizacao: item.data_finalizacao
                ? new Date(item.data_finalizacao).toLocaleDateString("pt-BR")
                : "",
            };
          }

          const size = Number(item.numero) || 0;
          const qtd = Number(item.quantidade) || 0;
          if (size > 0 && qtd > 0) {
            agrupado[key].items.push({ size, qtd });
            agrupado[key].total += qtd;
          }
        });

      return Object.values(agrupado).map((lancamento) => ({
        ...lancamento,
        items: lancamento.items.sort((a, b) => a.size - b.size),
      }));
    };

    console.log("TOTAL BRUTO MOVIMENTACOES", data.length);
    console.log(
      "PESPONTO BRUTO",
      data.filter((item) => String(item.tipo || "").trim() === "Pesponto")
    );
    console.log(
      "MONTAGEM BRUTO",
      data.filter((item) => String(item.tipo || "").trim() === "Montagem")
    );

    const pesponto = agrupar(data, "Pesponto");
    const montagem = agrupar(data, "Montagem");

    const ajustesEst = data
      .filter((item) => String(item.tipo || "").trim() === "Costura Pronta")
      .filter((item) => String(item.status || "").trim() === "Lançado")
      .map((item, index) => {
        const prog = String(item.programacao || "").trim();
        const isEntrada = prog.toLowerCase().includes("entrada");
        return {
          id: item.id || `ajuste-cp-${index}`,
          ref: String(item.ref || "").trim(),
          cor: String(item.cor || "").trim(),
          size: Number(item.numero) || 0,
          qtd: Number(item.quantidade) || 0,
          motivo: prog || "Ajuste manual na costura pronta",
          dataLancamento: item.data_lancamento
            ? new Date(item.data_lancamento).toLocaleDateString("pt-BR")
            : "",
          tipo: isEntrada ? "entrada" : "saida",
        };
      });

    return { pesponto, montagem, ajustesEst };
  } catch (error) {
    console.log("ERRO GERAL AO CARREGAR MOVIMENTACOES:", error);
    return { pesponto: [], montagem: [], ajustesEst: [] };
  }
};



const carregarConfiguracoesProducaoDoBanco = async () => {
  try {
    const { data, error } = await supabase
      .from("configuracoes_producao")
      .select("*")
      .order("id", { ascending: false })
      .limit(1)
      .maybeSingle();

    if (error) {
      console.log("ERRO AO CARREGAR CONFIGURACOES DE PRODUCAO:", error);
      return null;
    }

    return data;
  } catch (err) {
    console.log("ERRO GERAL AO CARREGAR CONFIGURACOES DE PRODUCAO:", err);
    return null;
  }
};

  const executeImport = async () => {
    if (importMode === "reset") {
      setRows((current) =>
        current.map((row) => ({
          ...row,
          data: Object.fromEntries(
            sizes.map((size) => [
              size,
              {
                ...row.data[size],
                pa: 0,
              },
            ])
          ),
        }))
      );
    }

    const parsed = importPreview;

    if (!parsed.length) {
      setImportFeedback("Não encontrei blocos válidos do GCM.");
      setImportPreview([]);
      setConfirmImport(false);
      return;
    }

    const produtosParaSalvar = parsed.map((item) => ({
      ref: item.ref,
      cor: item.cor,
    }));

    const parsedByKey = new Map(
      parsed.map((item) => [
        `${normalizeKey(item.ref)}__${normalizeKey(item.cor)}`,
        item,
      ])
    );

    const parsedByRef = parsed.reduce((acc, item) => {
      const key = normalizeKey(item.ref);
      if (!acc[key]) acc[key] = [];
      acc[key].push(item);
      return acc;
    }, {});

    const encontrarLinhaEstoqueAtual = (item) => {
      const exata = rows.find(
        (r) =>
          normalizeKey(r.ref) === normalizeKey(item.ref) &&
          normalizeKey(r.cor) === normalizeKey(item.cor)
      );
      if (exata) return exata;
      const candidatos = rows.filter((r) => normalizeKey(r.ref) === normalizeKey(item.ref));
      if (candidatos.length === 1) return candidatos[0];
      return candidatos.find(
        (r) =>
          normalizeKey(r.cor).includes(normalizeKey(item.cor)) ||
          normalizeKey(item.cor).includes(normalizeKey(r.cor))
      );
    };

    /** GCM só altera PA; costura pronta (est) e fluxo (m, p) mantêm-se do estoque atual. */
    const estoqueParaSalvar = [];
    parsed.forEach((item) => {
      const linhaAtual = encontrarLinhaEstoqueAtual(item);
      sizes.forEach((numero) => {
        const quantidade = Number(item.data?.[numero] || 0);
        if (quantidade <= 0) return;
        const cell = linhaAtual?.data?.[numero] || { pa: 0, est: 0, m: 0, p: 0 };
        const novoPa =
          importMode === "sum"
            ? (Number(cell.pa) || 0) + quantidade
            : Number(item.data?.[numero] || 0);
        estoqueParaSalvar.push({
          ref: item.ref,
          cor: item.cor,
          numero,
          pa: novoPa,
          est: Number(cell.est) || 0,
          m: Number(cell.m) || 0,
          p: Number(cell.p) || 0,
        });
      });
    });

    let atualizados = 0;
    const usados = new Set();

    setRows((current) => {
      const nextRows = current.map((row) => {
        const exactKey = `${normalizeKey(row.ref)}__${normalizeKey(row.cor)}`;
        let found = parsedByKey.get(exactKey);

        if (!found) {
          const candidates = parsedByRef[normalizeKey(row.ref)] || [];
          if (candidates.length === 1) {
            found = candidates[0];
          } else {
            found = candidates.find(
              (item) =>
                normalizeKey(item.cor).includes(normalizeKey(row.cor)) ||
                normalizeKey(row.cor).includes(normalizeKey(item.cor))
            );
          }
        }

        if (!found) return row;

        usados.add(`${normalizeKey(found.ref)}__${normalizeKey(found.cor)}`);
        atualizados += 1;

        const nextData = { ...row.data };
        sizes.forEach((size) => {
          nextData[size] = {
            ...nextData[size],
            pa:
              importMode === "sum"
                ? (nextData[size]?.pa || 0) + Number(found.data[size] || 0)
                : Number(found.data[size] || 0),
          };
        });

        return { ...row, data: nextData };
      });

      const novos = parsed.filter(
        (item) => !usados.has(`${normalizeKey(item.ref)}__${normalizeKey(item.cor)}`)
      );

      const novosRows = novos.map((item) => ({
        ref: item.ref,
        cor: item.cor,
        data: Object.fromEntries(
          sizes.map((size) => [
            size,
            {
              pa: Number(item.data[size] || 0),
              est: 0,
              m: 0,
              p: 0,
            },
          ])
        ),
      }));

      if (novosRows.length) atualizados += novosRows.length;

      return [...nextRows, ...novosRows];
    });

    setMinimos((curr) => {
      const next = { ...curr };
      parsed.forEach((item) => {
        if (!next[item.ref]) next[item.ref] = {};
        if (!next[item.ref][item.cor]) {
          next[item.ref][item.cor] = Object.fromEntries(
            sizes.map((size) => [size, { pa: size <= 39 ? 12 : 8, prod: 24 }])
          );
        }
      });
      return next;
    });

    setVendas((curr) => {
      const next = { ...curr };
      parsed.forEach((item) => {
        if (!next[item.ref]) next[item.ref] = {};
        if (!next[item.ref][item.cor]) {
          next[item.ref][item.cor] = Object.fromEntries(sizes.map((size) => [size, 0]));
        }
      });
      return next;
    });

    setImportFeedback(`${atualizados} item(ns) atualizado(s) no Produto Acabado.`);
    setUltimaImportacaoGcm({
      arquivo: importFileName || "Importação manual",
      dataHora: new Date().toLocaleString("pt-BR"),
      modo: importMode,
      itens: parsed.length,
      atualizados,
      totalPares: parsed.reduce((acc, item) => acc + (item.total || 0), 0),
      preview: parsed.slice(0, 8),
    });

    await salvarProdutosNoBanco(produtosParaSalvar);
    await salvarEstoqueNoBanco(estoqueParaSalvar);

    setConfirmImport(false);
  };

  const applyImport = () => {
  const parsed = importPreview;
    if (!parsed.length) {
      setImportFeedback("Não encontrei blocos válidos do GCM.");
      setImportPreview([]);
      return;
    }
    setImportPreview(parsed);
    setConfirmImport(true);
  };

  const updateMin = (ref, cor, size, key, value) => {
    setMinimos((curr) => ({
      ...curr,
      [ref]: {
        ...curr[ref],
        [cor]: {
          ...curr[ref][cor],
          [size]: {
            ...curr[ref][cor][size],
            [key]: Number(value) || 0,
          },
        },
      },
    }));
  };

  const updateVenda = (ref, cor, size, value) => {
    setVendasDraft((curr) => ({
      ...curr,
      [ref]: {
        ...(curr?.[ref] || {}),
        [cor]: {
          ...(curr?.[ref]?.[cor] || makeEmptyGrid()),
          [size]: Number(value) || 0,
        },
      },
    }));
    setVendasDirty(true);
  };

const salvarVendasManuais = async () => {
  setVendas(vendasDraft);
  await salvarVendasNoBanco(vendasDraft);

  const resumo = sortedRowsByRefCor
    .map((row) => {
      const atual = vendasDraft?.[row.ref]?.[row.cor] || makeEmptyGrid();
      const total = sizes.reduce((acc, size) => acc + (Number(atual[size]) || 0), 0);
      return total > 0 ? { ref: row.ref, cor: row.cor, total } : null;
    })
    .filter(Boolean)
    .sort((a, b) => b.total - a.total)
    .slice(0, 8);

  setHistoricoVendasManuais((curr) => [
    {
      id: `vendas-manual-${Date.now()}`,
      dataHora: new Date().toLocaleString("pt-BR"),
      itens: resumo.length,
      totalPares: resumo.reduce((acc, item) => acc + item.total, 0),
      resumo,
    },
    ...curr,
  ]);

  setVendasDirty(false);
  setSalesImportFeedback("Vendas manuais salvas com sucesso.");
};

  const handleSalesFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const grouped = {};
      let processedRows = 0;
      data.forEach((row) => {
        const produtoRaw = String(row["PRODUTO"] || row["Produto"] || "").trim();
        const padRaw = String(row["PADRONIZAÇÃO"] || row["PADRONIZACAO"] || row["Padronização"] || row["Padronizacao"] || "").trim();
        const qtd = Number(row["QTDE"] || row["QTD"] || row["Quantidade"] || row["QUANTIDADE"] || 0);
        const size = Number(String(padRaw).match(/\d+/)?.[0]);
        if (!produtoRaw || !sizes.includes(size) || !qtd) return;

        const produto = normalizeKey(produtoRaw);
        const found = rows.find((r) => {
          const refKey = normalizeKey(r.ref);
          const corKey = normalizeKey(r.cor);
          const together = normalizeKey(`${r.ref} ${r.cor}`);
          return produto.includes(refKey) || produto.includes(corKey) || produto.includes(together) || together.includes(produto);
        });
        if (!found) return;
        if (!grouped[found.ref]) grouped[found.ref] = {};
        if (!grouped[found.ref][found.cor]) grouped[found.ref][found.cor] = makeEmptyGrid();
        grouped[found.ref][found.cor][size] += qtd;
        processedRows += 1;
      });

      const preview = Object.keys(grouped).flatMap((ref) =>
        Object.keys(grouped[ref]).map((cor) => ({ ref, cor, data: grouped[ref][cor] }))
      );
      setSalesImportFileName(file.name);
      if (!preview.length) {
        setSalesImportPreview([]);
        setSalesImportFeedback("Li o arquivo, mas não consegui relacionar os produtos.");
        return;
      }
      setVendas((curr) => {
        const next = { ...curr };
        Object.keys(grouped).forEach((ref) => {
          if (!next[ref]) next[ref] = {};
          Object.keys(grouped[ref]).forEach((cor) => {
            next[ref][cor] = { ...(next[ref][cor] || makeEmptyGrid()) };
            sizes.forEach((size) => {
              next[ref][cor][size] = grouped[ref][cor][size] || 0;
            });
          });
        });
        return next;
      });
      setSalesImportPreview(preview);
      setSalesImportFeedback(`Arquivo lido com sucesso. ${processedRows} linha(s) aproveitadas.`);
    } catch {
      setSalesImportFeedback("Erro ao importar vendas.");
      setSalesImportPreview([]);
    }
  };

  const renderDashboard = () => {
    const d = dashboardData;
    const tempoTotal = d.tempoTotal;
    const capP = Number(capacidadePespontoDia) || 396;
    const capM = Number(capacidadeMontagemDia) || 396;
    const progP = Math.min(100, ((d.programacaoHojePesponto?.totalProgramado || 0) / capP) * 100);
    const progM = Math.min(100, ((d.programacaoHojeMontagem?.totalProgramado || 0) / capM) * 100);
    const dashTabs = [
      { id: "visao", label: "Visão" },
      { id: "calendario", label: "Calendário" },
      { id: "riscos", label: "Riscos" },
      { id: "vendas", label: "Vendas" },
    ];
    const tabOn = (id) => dashboardMobileTab === id;
    const showMobile = (id) => `${tabOn(id) ? "" : "hidden"} lg:block`;

    const calendarioBlock = (
      <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5 sm:p-6">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
          <div>
            <h2 className="font-bold text-lg">Calendário útil do mês</h2>
            <p className="text-sm text-slate-500 mt-1">
              Os cards de média finalizada usam apenas dias úteis transcorridos, desconsiderando sábados, domingos e os feriados informados abaixo.
            </p>
          </div>
          <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-right shrink-0">
            <div className="text-xs text-slate-500">Dias úteis considerados</div>
            <div className="text-2xl font-bold text-slate-900">{d.diasUteisNoMes}</div>
          </div>
        </div>
        <div className="mt-4 grid grid-cols-1 xl:grid-cols-[1fr_280px] gap-4">
          <label className="text-sm font-medium text-slate-700">
            Feriados do mês atual
            <textarea
              value={feriadosTexto}
              onChange={(e) => setFeriadosTexto(e.target.value)}
              placeholder="Ex.: 01/05/2026&#10;11/06/2026"
              className="mt-2 w-full min-h-[96px] rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
            />
            <div className="mt-2 text-xs text-slate-500">Digite um feriado por linha. Formato: dd/mm/aaaa</div>
          </label>
          <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
            <div className="text-sm font-semibold text-slate-900">Resumo do cálculo</div>
            <div className="mt-3 space-y-2 text-sm text-slate-600">
              <div className="flex items-center justify-between">
                <span>Dias corridos no mês</span>
                <span className="font-semibold">{d.hoje.getDate()}</span>
              </div>
              <div className="flex items-center justify-between">
                <span>Feriados informados</span>
                <span className="font-semibold">{d.feriadosLista.length}</span>
              </div>
              <div className="flex items-center justify-between">
                <span>Dias úteis usados</span>
                <span className="font-semibold text-[#8B1E2D]">{d.diasUteisNoMes}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    );

    return (
      <PageShell title="Dashboard" subtitle="Visão executiva da operação com foco em giro, risco de ruptura, programação e andamento da produção.">
        <div className="lg:hidden sticky top-1 z-10 mb-3 flex gap-1 overflow-x-auto rounded-2xl border border-slate-200 bg-white/95 p-1 shadow-sm backdrop-blur [-webkit-overflow-scrolling:touch]">
          {dashTabs.map((t) => (
            <button
              key={t.id}
              type="button"
              onClick={() => setDashboardMobileTab(t.id)}
              className={`shrink-0 rounded-xl px-3 py-2 text-xs font-semibold transition sm:text-sm ${
                dashboardMobileTab === t.id ? "bg-[#0F172A] text-white shadow-sm" : "bg-slate-50 text-slate-700 hover:bg-slate-100"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>

        {/* KPIs principais + opcionais */}
        <div className={`space-y-3 ${showMobile("visao")}`}>
          <section className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4">
            <SummaryCard title="PA total" value={d.totalPA} subtitle="Pares em produto acabado" />
            <SummaryCard title="Na montagem" value={d.totalMontagemAtual} subtitle="Fluxo atual de montagem" />
            <SummaryCard title="No pesponto" value={d.totalPespontoAtual} subtitle="Fluxo atual de pesponto" />
            <SummaryCard
              title="Venda diária (aprox.)"
              value={d.vendaDiariaTotal.toFixed(1)}
              subtitle={`Soma das vendas do mês (${d.vendaMensalTotal.toFixed(0)}) ÷ 30 dias corridos — referência rápida, não é previsão.`}
            />
          </section>
          <button
            type="button"
            onClick={() => setDashboardMaisKpisAberto((v) => !v)}
            className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold text-[#0F172A] shadow-sm hover:bg-slate-50 lg:max-w-md"
          >
            {dashboardMaisKpisAberto ? "Ocultar mais indicadores" : "Mostrar mais indicadores (costura e médias do mês)"}
          </button>
          {dashboardMaisKpisAberto ? (
            <section className="grid grid-cols-1 md:grid-cols-3 gap-3 sm:gap-4">
              <SummaryCard title="Costura pronta" value={d.totalEst} subtitle="Pares prontos para montagem" />
              <SummaryCard
                title="Média finalizada • Pesponto"
                value={d.mediaFinalizadaPesponto.toFixed(1)}
                subtitle={`Finalizados no mês ÷ ${d.diasUteisNoMes} dias úteis • Hoje ${d.finalizadoHojePesponto}`}
              />
              <SummaryCard
                title="Média finalizada • Montagem"
                value={d.mediaFinalizadaMontagem.toFixed(1)}
                subtitle={`Finalizados no mês ÷ ${d.diasUteisNoMes} dias úteis • Hoje ${d.finalizadoHojeMontagem}`}
              />
            </section>
          ) : null}
        </div>

        {/* Calendário: colapsável no desktop; aba no mobile */}
        <div className={`mt-6 ${showMobile("calendario")}`}>
          <div className="hidden lg:block">
            {!dashboardFeriadosAberto ? (
              <div className="flex flex-col gap-3 rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h2 className="font-bold text-lg">Calendário útil e feriados</h2>
                  <p className="text-sm text-slate-500 mt-1">
                    Ajuste feriados para o cálculo das médias. Dias úteis no mês: <strong className="text-[#8B1E2D]">{d.diasUteisNoMes}</strong>
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setDashboardFeriadosAberto(true)}
                  className="shrink-0 rounded-2xl bg-[#0F172A] px-5 py-3 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"
                >
                  Mostrar calendário e feriados
                </button>
              </div>
            ) : (
              <div>
                <div className="mb-3 flex justify-end lg:justify-end">
                  <button
                    type="button"
                    onClick={() => setDashboardFeriadosAberto(false)}
                    className="text-sm font-semibold text-slate-600 hover:text-[#8B1E2D]"
                  >
                    Recolher painel
                  </button>
                </div>
                {calendarioBlock}
              </div>
            )}
          </div>
          <div className="lg:hidden">{calendarioBlock}</div>
        </div>

        {/* Radar + programação */}
        <div className={`mt-6 ${showMobile("visao")}`}>
          <section className="grid grid-cols-1 xl:grid-cols-[1.2fr_0.8fr] gap-6">
            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5 sm:p-6 xl:col-span-2">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
                <div>
                  <h2 className="font-bold text-lg">Radar da operação</h2>
                  <p className="text-sm text-slate-500 mt-1">Indicadores principais para decidir a prioridade do dia.</p>
                  <p className="mt-2 text-xs text-slate-500 leading-relaxed max-w-xl">
                    <strong>Cobertura</strong> (em “Próximas ações”) estima quantos dias o PA dura face à venda. Valores abaixo do <strong>lead time</strong> de {tempoTotal} dia(s) indicam maior urgência de reabastecer o fluxo.
                  </p>
                </div>
                <span className="shrink-0 px-3 py-1 rounded-full border text-xs font-semibold bg-[#FFF7F8] text-[#8B1E2D] border-[#E7C7CC]">
                  Lead time {tempoTotal} dia(s)
                </span>
              </div>

              <div className="mt-5 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={() => setActive("Programação do Dia")}
                  className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-semibold text-[#0F172A] hover:bg-slate-100"
                >
                  Programação do dia
                </button>
                <button
                  type="button"
                  onClick={() => setActive("Pesponto")}
                  className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-semibold text-[#0F172A] hover:bg-slate-100"
                >
                  Pesponto
                </button>
                <button
                  type="button"
                  onClick={() => setActive("Montagem")}
                  className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-semibold text-[#0F172A] hover:bg-slate-100"
                >
                  Montagem
                </button>
              </div>

              <div className="mt-5 grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="rounded-[24px] border border-slate-200 bg-slate-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">Situação do estoque</div>
                  <div className="mt-4 space-y-3">
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-slate-500">Itens críticos</span>
                      <span className="text-2xl font-black text-red-600">{metrics.criticos}</span>
                    </div>
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-slate-500">Atenção PA</span>
                      <span className="text-xl font-bold text-amber-600">{metrics.atencaoPA}</span>
                    </div>
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-slate-500">Atenção produção</span>
                      <span className="text-xl font-bold text-sky-700">{metrics.atencaoProd}</span>
                    </div>
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-slate-500">Itens OK</span>
                      <span className="text-xl font-bold text-emerald-600">{metrics.ok}</span>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => setActive("Controle Geral")}
                    className="mt-4 w-full rounded-xl border border-[#8B1E2D]/30 bg-white py-2 text-xs font-semibold text-[#8B1E2D] hover:bg-[#FFF7F8]"
                  >
                    Abrir Controle geral
                  </button>
                </div>

                <div className="rounded-[24px] border border-slate-200 bg-slate-50 p-5">
                  <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">Programação do dia</div>
                  <div className="mt-4 space-y-4">
                    <div>
                      <div className="flex items-center justify-between text-sm">
                        <span className="text-slate-500">Pesponto hoje</span>
                        <span className="text-2xl font-black text-[#0F172A]">{d.programacaoHojePesponto?.totalProgramado || 0}</span>
                      </div>
                      <div className="mt-2 h-2 overflow-hidden rounded-full bg-slate-200">
                        <div className="h-full rounded-full bg-[#8B1E2D] transition-all" style={{ width: `${progP}%` }} />
                      </div>
                      <div className="mt-1 text-[11px] text-slate-500">Capacidade diária: {capP} pares</div>
                    </div>
                    <div>
                      <div className="flex items-center justify-between text-sm">
                        <span className="text-slate-500">Montagem hoje</span>
                        <span className="text-2xl font-black text-[#0F172A]">{d.programacaoHojeMontagem?.totalProgramado || 0}</span>
                      </div>
                      <div className="mt-2 h-2 overflow-hidden rounded-full bg-slate-200">
                        <div className="h-full rounded-full bg-[#0F172A] transition-all" style={{ width: `${progM}%` }} />
                      </div>
                      <div className="mt-1 text-[11px] text-slate-500">Capacidade diária: {capM} pares</div>
                    </div>
                    <div className="flex items-center justify-between text-sm pt-1 border-t border-slate-200">
                      <span className="text-slate-500">Fichas pesponto</span>
                      <span className="text-xl font-bold text-[#8B1E2D]">{d.programacaoHojePesponto?.fichas?.length || 0}</span>
                    </div>
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-slate-500">Fichas montagem</span>
                      <span className="text-xl font-bold text-[#8B1E2D]">{d.programacaoHojeMontagem?.fichas?.length || 0}</span>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </section>
        </div>

        {/* Próximas ações */}
        <div className={`mt-6 ${showMobile("riscos")}`}>
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5 sm:p-6">
            <div className="flex items-center justify-between gap-3">
              <div>
                <h2 className="font-bold text-lg">Próximas ações</h2>
                <p className="text-sm text-slate-500 mt-1">Menor cobertura em dias — priorize reabastecimento.</p>
              </div>
            </div>
            <div className="mt-4 space-y-3">
              {d.menorCobertura.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                  Sem vendas suficientes para montar alertas.
                </div>
              ) : (
                d.menorCobertura.map((item, idx) => (
                  <div key={`${item.ref}-${item.cor}-${item.size}-dash-${idx}`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                      <div className="min-w-0">
                        <div className="font-semibold text-slate-900">
                          {item.ref} • {item.cor}
                        </div>
                        <div className="text-xs text-slate-500 mt-1">
                          Numeração {item.size} • PA {item.pa}
                        </div>
                      </div>
                      <div className="flex flex-wrap items-center gap-2">
                        <span className={`px-3 py-1 rounded-full border text-xs font-semibold ${coberturaBadgeClass(item.cobertura, tempoTotal)}`}>
                          {item.cobertura?.toFixed(1)} dia(s)
                        </span>
                        <button
                          type="button"
                          onClick={() => irControleGeral(item.ref, item.cor, item.size)}
                          className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs font-semibold text-blue-800 hover:bg-blue-100"
                        >
                          Abrir no Controle
                        </button>
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

        {/* Top vendas + sugestões */}
        <div className={`mt-6 ${showMobile("vendas")}`}>
          <section className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5 sm:p-6">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <h2 className="font-bold text-lg">Top modelos por venda</h2>
                <span className="px-3 py-1 rounded-full border text-xs font-semibold bg-emerald-100 text-emerald-700 border-emerald-200">Mensal</span>
              </div>
              <div className="mt-4 space-y-3">
                {d.topModelos.length === 0 ? (
                  <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                    Sem vendas lançadas ainda.
                  </div>
                ) : (
                  d.topModelos.map((item, idx) => (
                    <div key={`${item.ref}-${item.cor}-dashboard-top`} className="rounded-2xl border border-slate-200 px-4 py-3">
                      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                        <div>
                          <div className="font-semibold text-slate-900">
                            #{idx + 1} {item.ref} • {item.cor}
                          </div>
                          <div className="text-xs text-slate-500 mt-1">PA atual {item.totalPA}</div>
                        </div>
                        <div className="flex flex-wrap items-center gap-3">
                          <div className="text-right">
                            <div className="text-2xl font-black text-[#0F172A]">{item.totalVendido}</div>
                            <div className="text-xs text-slate-500">vendidos/mês</div>
                          </div>
                          <button
                            type="button"
                            onClick={() => irControleGeral(item.ref, item.cor, null)}
                            className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs font-semibold text-blue-800 hover:bg-blue-100"
                          >
                            Controle
                          </button>
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>

            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5 sm:p-6">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <h2 className="font-bold text-lg">Resumo das sugestões</h2>
                <span className="px-3 py-1 rounded-full border text-xs font-semibold bg-[#FFF7F8] text-[#8B1E2D] border-[#E7C7CC]">Planejamento</span>
              </div>
              <button
                type="button"
                onClick={() => setActive("Sugestoes")}
                className="mt-4 w-full rounded-2xl border border-[#8B1E2D]/30 bg-[#FFF7F8] px-4 py-3 text-sm font-semibold text-[#8B1E2D] hover:bg-[#fce8ec] sm:w-auto"
              >
                Abrir tela Sugestões
              </button>
              <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <div className="text-sm text-slate-500">Sugestões de pesponto</div>
                  <div className="mt-2 text-3xl font-black text-[#0F172A]">{suggestions.pesponto.length}</div>
                  <div className="mt-2 text-xs text-slate-500">
                    Total sugerido: {suggestions.pesponto.reduce((acc, item) => acc + item.total, 0)} pares
                  </div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <div className="text-sm text-slate-500">Sugestões de montagem</div>
                  <div className="mt-2 text-3xl font-black text-[#0F172A]">{suggestions.montagem.length}</div>
                  <div className="mt-2 text-xs text-slate-500">
                    Total sugerido: {suggestions.montagem.reduce((acc, item) => acc + item.total, 0)} pares
                  </div>
                </div>
              </div>
            </div>
          </section>
        </div>
      </PageShell>
    );
  };

  const renderControle = () => {
    const tempoTotal = (Number(tempoProducao?.pesponto) || 0) + (Number(tempoProducao?.montagem) || 0);

    const coberturaCritica = controleRows
      .flatMap((row) => sizes.map((size) => {
        const vendaMes = Number(vendas?.[row.ref]?.[row.cor]?.[size]) || 0;
        const cobertura = coberturaDias(row.data[size]?.pa || 0, vendaMes);
        return { row, size, cobertura, vendaMes };
      }))
      .filter((item) => item.cobertura !== null)
      .sort((a, b) => a.cobertura - b.cobertura)
      .slice(0, 6);

    return (
    <PageShell title="Controle Geral" subtitle="Visão consolidada por referência, cor e numeração com foco em decisão rápida.">
      <section className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-5 gap-4">
        <SummaryCard title="Itens críticos" value={metrics.criticos} subtitle="Ação imediata" />
        <SummaryCard title="Atenção PA" value={metrics.atencaoPA} subtitle="PA abaixo do mínimo" />
        <SummaryCard title="Atenção produção" value={metrics.atencaoProd} subtitle="Fluxo abaixo do mínimo" />
        <SummaryCard title="Itens OK" value={metrics.ok} subtitle="Saudável" />
        <SummaryCard title="Costura pronta" value={metrics.costura} subtitle="Pares disponíveis" />
      </section>

      <section className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <label className="text-sm font-medium text-slate-700">
            Ref
            <select value={controleFiltroRef} onChange={(e) => { setControleFiltroRef(e.target.value); setControleFiltroCor("TODAS"); }} className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm">
              {controleRefs.map((ref) => <option key={ref} value={ref}>{ref}</option>)}
            </select>
          </label>
          <label className="text-sm font-medium text-slate-700">
            Cor
            <select value={controleFiltroCor} onChange={(e) => setControleFiltroCor(e.target.value)} className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm">
              {controleCores.map((cor) => <option key={cor} value={cor}>{cor}</option>)}
            </select>
          </label>
          <label className="text-sm font-medium text-slate-700">
            Numeração
            <select value={controleFiltroNumero} onChange={(e) => setControleFiltroNumero(e.target.value)} className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm">
              <option value="TODAS">TODAS</option>
              {sizes.map((size) => <option key={size} value={size}>{size}</option>)}
            </select>
          </label>
        </div>
      </section>

      <section className="space-y-6">
        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-hidden">
          <div className="px-6 py-5 border-b border-slate-200">
            <h2 className="font-bold text-lg">Mapa da produção</h2>
            <p className="text-sm text-slate-500 mt-1">Cada bloco mostra PA, EST PTO, M, P e TOTAL.</p>
          </div>
          <div className="overflow-auto">
            <table className="min-w-[2100px] w-full border-collapse">
              <thead>
                <tr className="bg-slate-50">
                  <th rowSpan={2} className="sticky left-0 z-20 bg-slate-50 border-b border-r border-slate-200 px-4 py-4 text-left text-sm font-bold min-w-[120px]">Ref</th>
                  <th rowSpan={2} className="sticky left-[120px] z-20 bg-slate-50 border-b border-r border-slate-200 px-4 py-4 text-left text-sm font-bold min-w-[120px]">Cor</th>
                  {visibleSizesControle.map((size) => <th key={size} colSpan={5} className="border-b border-r border-slate-200 px-4 py-3 text-center text-sm font-bold">{size}</th>)}
                </tr>
                <tr className="bg-slate-50">
                  {visibleSizesControle.flatMap((size) => ["PA", "EST PTO", "M", "P", "TOTAL"].map((label) => <th key={`${size}-${label}`} className="border-b border-r border-slate-200 px-3 py-2 text-[11px] font-bold tracking-wide text-slate-500">{label}</th>))}
                </tr>
              </thead>
              <tbody>
                {controleRows.map((row) => (
                  <tr key={`${row.ref}-${row.cor}`} className="hover:bg-slate-50/70">
                    <td className="sticky left-0 z-10 bg-white border-b border-r border-slate-200 px-4 py-4 font-semibold">{row.ref}</td>
                    <td className="sticky left-[120px] z-10 bg-white border-b border-r border-slate-200 px-4 py-4">{row.cor}</td>
                    {visibleSizesControle.flatMap((size) => {
                      const item = row.data[size] || { pa: 0, est: 0, m: 0, p: 0 };
                      const minimo = minimos?.[row.ref]?.[row.cor]?.[size] || { pa: 0, prod: 0 };
                      const st = statusFor(item, minimo);
                      return [
                        <td key={`${row.ref}-${size}-pa`} className={`border-b border-r border-slate-200 px-3 py-3 text-center text-sm ${tone(st)}`}>{item.pa}</td>,
                        <td key={`${row.ref}-${size}-est`} className={`border-b border-r border-slate-200 px-3 py-3 text-center text-sm ${tone(st)}`}>{item.est}</td>,
                        <td key={`${row.ref}-${size}-m`} className={`border-b border-r border-slate-200 px-3 py-3 text-center text-sm ${tone(st)}`}>{item.m}</td>,
                        <td key={`${row.ref}-${size}-p`} className={`border-b border-r border-slate-200 px-3 py-3 text-center text-sm ${tone(st)}`}>{item.p}</td>,
                        <td key={`${row.ref}-${size}-t`} className={`border-b border-r border-slate-200 px-3 py-3 text-center text-sm font-bold ${tone(st)}`}>{calcTotal(item)}</td>,
                      ];
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5">
          <h2 className="font-bold text-lg">Prioridades</h2>
          <div className="mt-4 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
            {controleRows
              .flatMap((row) => sizes.map((size) => ({ row, size, status: statusFor(row.data[size], minimos?.[row.ref]?.[row.cor]?.[size] || { pa: 0, prod: 0 }) })))
              .filter((x) => x.status !== "OK")
              .slice(0, 9)
              .map((item, idx) => {
                const vendaMes = Number(vendas?.[item.row.ref]?.[item.row.cor]?.[item.size]) || 0;
                const cobertura = coberturaDias(item.row.data[item.size]?.pa || 0, vendaMes);
                return (
                  <div key={`${item.row.ref}-${item.size}-${idx}`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex items-center justify-between gap-3">
                      <span className="text-sm font-medium">{item.row.ref} {item.row.cor} {item.size}</span>
                      <span className={`px-3 py-1 rounded-full border text-xs font-semibold capitalize ${badge(item.status)}`}>{item.status}</span>
                    </div>
                    <div className="mt-2 text-xs text-slate-500">
                      Cobertura: {cobertura == null ? "sem giro" : `${cobertura.toFixed(1)} dia(s)`}
                    </div>
                  </div>
                );
              })}
          </div>
        </div>

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-5">
          <div className="flex items-center justify-between gap-4">
            <h2 className="font-bold text-lg">Cobertura de estoque</h2>
            <span className="text-sm text-slate-500">Lead time: {tempoTotal} dia(s)</span>
          </div>
          <div className="mt-4 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
            {coberturaCritica.length === 0 ? (
              <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500 xl:col-span-3">
                Ainda não há vendas suficientes para calcular cobertura.
              </div>
            ) : (
              coberturaCritica.map((item, idx) => (
                <div key={`${item.row.ref}-${item.row.cor}-${item.size}-cob-${idx}`} className="rounded-2xl border border-slate-200 px-4 py-3">
                  <div className="flex items-center justify-between gap-3">
                    <div>
                      <div className="font-semibold text-slate-900">{item.row.ref} • {item.row.cor}</div>
                      <div className="text-sm text-slate-500 mt-1">Numeração {item.size}</div>
                    </div>
                    <span className={`px-3 py-1 rounded-full border text-xs font-semibold ${coberturaBadgeClass(item.cobertura, tempoTotal)}`}>
                      {coberturaLabel(item.cobertura, tempoTotal)}
                    </span>
                  </div>
                  <div className="mt-3 flex items-center justify-between text-sm text-slate-600">
                    <span>PA atual</span>
                    <span className="font-semibold">{item.row.data[item.size]?.pa || 0}</span>
                  </div>
                  <div className="mt-1 flex items-center justify-between text-sm text-slate-600">
                    <span>Venda diária</span>
                    <span className="font-semibold">{vendaDiaFromMes(item.vendaMes).toFixed(1)}</span>
                  </div>
                  <div className="mt-1 flex items-center justify-between text-sm text-slate-600">
                    <span>Cobertura</span>
                    <span className="font-semibold">{item.cobertura == null ? "sem giro" : `${item.cobertura.toFixed(1)} dia(s)`}</span>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </section>
    </PageShell>
    );
  };

  const renderImport = () => (
    <PageShell
      title="Importar GCM"
      subtitle="Importe o arquivo do GCM para atualizar o Produto Acabado (PA)."
      action={
        <button onClick={applyImport} className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold shadow-sm">
          Aplicar importação
        </button>
      }
    >
      <div className="grid grid-cols-1 xl:grid-cols-[1fr_340px] gap-6">
        <div className="space-y-4">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="mb-4">
              <div className="text-sm font-semibold text-slate-700 mb-2">Modo de importação</div>
              <div className="flex flex-col md:flex-row gap-2">
                <label className="flex items-center gap-2 text-sm">
                  <input type="radio" name="importMode" value="replace" checked={importMode === "replace"} onChange={(e) => setImportMode(e.target.value)} />
                  Substituir estoque
                </label>
                <label className="flex items-center gap-2 text-sm">
                  <input type="radio" name="importMode" value="sum" checked={importMode === "sum"} onChange={(e) => setImportMode(e.target.value)} />
                  Somar ao estoque
                </label>
                <label className="flex items-center gap-2 text-sm">
                  <input type="radio" name="importMode" value="reset" checked={importMode === "reset"} onChange={(e) => setImportMode(e.target.value)} />
                  Zerar e importar
                </label>
              </div>
            </div>

            <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
              <div>
                <div className="font-semibold">Importar arquivo do GCM</div>
                <div className="text-sm text-slate-500 mt-1">Aceita .xls e .xlsx exportados do seu sistema.</div>
              </div>
              <label className="inline-flex cursor-pointer items-center justify-center rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white shadow-sm">
                Escolher arquivo
                <input type="file" accept=".xls,.xlsx" onChange={handleFileUpload} className="hidden" />
              </label>
            </div>

            {importFileName && (
              <div className="mt-4 rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-sm">
                Arquivo carregado: <span className="font-semibold">{importFileName}</span>
              </div>
            )}
            {importFeedback && <div className="mt-3 text-sm text-slate-600">{importFeedback}</div>}
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="font-semibold mb-3">Conteúdo bruto</div>
            <textarea value={importText} onChange={(e) => setImportText(e.target.value)} className="w-full h-80 rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm outline-none" />
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div>
              <div className="font-semibold">Histórico da última importação</div>
              <div className="text-sm text-slate-500 mt-1">Sempre exibe a importação mais recente aplicada no produto acabado.</div>
            </div>

            {!ultimaImportacaoGcm ? (
              <div className="mt-4 rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                Nenhuma importação aplicada ainda.
              </div>
            ) : (
              <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-3 text-sm">
                <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                  <div className="text-slate-500">Arquivo</div>
                  <div className="font-semibold text-slate-900 mt-1">{ultimaImportacaoGcm.arquivo}</div>
                </div>
                <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                  <div className="text-slate-500">Data / hora</div>
                  <div className="font-semibold text-slate-900 mt-1">{ultimaImportacaoGcm.dataHora}</div>
                </div>
                <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                  <div className="text-slate-500">Modo</div>
                  <div className="font-semibold text-slate-900 mt-1">{ultimaImportacaoGcm.modo === "replace" ? "Substituir estoque" : ultimaImportacaoGcm.modo === "sum" ? "Somar ao estoque" : "Zerar e importar"}</div>
                </div>
                <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                  <div className="text-slate-500">Itens / pares</div>
                  <div className="font-semibold text-slate-900 mt-1">{ultimaImportacaoGcm.itens} item(ns) • {ultimaImportacaoGcm.totalPares} pares</div>
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex items-center justify-between gap-4">
            <div>
              <div className="font-semibold">Preview da importação</div>
              <div className="text-sm text-slate-500 mt-1">{importPreview.length} item(ns) reconhecido(s)</div>
            </div>
            {importPreview.length > 0 && (
              <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-right">
                <div className="text-xs text-slate-500">Total PA lido</div>
                <div className="text-2xl font-bold text-slate-900">{importPreview.reduce((acc, item) => acc + (item.total || 0), 0)}</div>
              </div>
            )}
          </div>

          {importPreview.length === 0 ? (
            <div className="mt-4 rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
              Carregue o arquivo para visualizar o preview antes de aplicar.
            </div>
          ) : (
            <div className="mt-4 space-y-4 max-h-[640px] overflow-auto pr-1">
              {importPreview.map((item, idx) => (
                <div key={`${item.ref}-${item.cor}-${idx}`} className="rounded-2xl border border-slate-200 overflow-hidden">
                  <div className="bg-slate-50 border-b border-slate-200 px-4 py-3 flex items-center justify-between gap-4">
                    <div>
                      <div className="font-semibold text-slate-900">{item.ref}</div>
                      <div className="text-sm text-slate-500 mt-1">{item.cor}</div>
                    </div>
                    <div className="px-3 py-2 rounded-xl bg-white border border-slate-200 text-sm font-semibold">{item.total || 0} pares</div>
                  </div>
                  <div className="p-4">
                    <div className="grid grid-cols-3 gap-2">
                      {sizes.map((size) => (
                        <div key={size} className="rounded-xl bg-slate-50 border border-slate-200 px-3 py-2 text-center">
                          <div className="text-xs text-slate-500">{size}</div>
                          <div className="text-sm font-bold text-slate-900 mt-1">{item.data[size] || 0}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </PageShell>
  );

  const renderMovPage = (title, form, setForm, subtitle) => {
    const MOV_PAGE_SIZE = 15;
    const selectionPreview = previewBySelection(form);
    const liveError = getMovErrorMessage(title, form);
    const lancamentos = title === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const lancamentosAbertosCor = lancamentos.filter((item) => {
      if (String(item.status || "").trim() === "Finalizado") return false;
      return (
        normalizeKey(String(item.ref || "")) === normalizeKey(String(form.ref || "")) &&
        normalizeKey(String(item.cor || "")) === normalizeKey(String(form.cor || ""))
      );
    });
    const totalAtual = sizes.reduce((acc, size) => acc + (Number(form.grid[size]) || 0), 0);
    const isPespontoPage = title === "Pesponto";
    const tituloFolhaMov = (String(programacaoEtiquetaFicha || "").trim() || `${title} - Ficha`);
    const nomeProgramacaoMov = String(programacaoNomeLoteImpressao || "").trim() || String(programacaoEtiquetaFicha || "").trim();
    const selecaoMovAtual = movImpressaoSelecao[title] || {};
    const lancamentosSelecionadosMov = lancamentos.filter((item) => selecaoMovAtual[item.id] === true);
    const buildFichaFromLancamento = (item, idx) => {
      const grid = Object.fromEntries(sizes.map((size) => [size, 0]));
      (item.items || []).forEach((entry) => {
        const sizeNum = Number(entry.size);
        if (Number.isFinite(sizeNum) && Object.prototype.hasOwnProperty.call(grid, sizeNum)) {
          grid[sizeNum] = Number(entry.qtd) || 0;
        }
      });
      return {
        ficha: {
          ref: item.ref,
          cor: item.cor,
          nome: item.programacao || `${item.ref} - ${item.cor}`,
          sizes: grid,
          total: Number(item.total) || sizes.reduce((acc, s) => acc + (Number(grid[s]) || 0), 0),
        },
        dia: 1,
        ordem: idx + 1,
        programacaoNome: nomeProgramacaoMov || item.programacao || "",
      };
    };
    const itensMovBase =
      lancamentosSelecionadosMov.length > 0
        ? lancamentosSelecionadosMov.map(buildFichaFromLancamento)
        : [
            {
              ficha: {
                ref: form.ref,
                cor: form.cor,
                nome: form.programacao || `${form.ref} - ${form.cor}`,
                sizes: Object.fromEntries(sizes.map((size) => [size, Number(form.grid?.[size] || 0)])),
                total: totalAtual,
              },
              dia: 1,
              ordem: 1,
              programacaoNome: nomeProgramacaoMov,
            },
          ];
    const buildItensMovComCopias = (copias) => {
      const lista = [];
      for (let copia = 1; copia <= copias; copia += 1) {
        itensMovBase.forEach((item) => {
          lista.push({ ...item, copia });
        });
      }
      return lista;
    };
    const itensMovFolha1 = buildItensMovComCopias(3);
    const itensMovFolha2 = buildItensMovComCopias(Math.max(1, Math.min(4, Number(programacaoCopiasPorPagina) || 1)));
    const itensMovImpressao =
      programacaoTipoFolha === "folha1"
        ? itensMovFolha1
        : programacaoTipoFolha === "folha2"
          ? itensMovFolha2
          : [...itensMovFolha1, ...itensMovFolha2];
    const totalPages = Math.max(1, Math.ceil(lancamentos.length / MOV_PAGE_SIZE));
    const currentPage = Math.min(movListPage[title] || 1, totalPages);
    const startIdx = (currentPage - 1) * MOV_PAGE_SIZE;
    const lancamentosPagina = lancamentos.slice(startIdx, startIdx + MOV_PAGE_SIZE);
    const agrupados = lancamentosPagina.reduce((acc, item) => {
      if (!acc[item.programacao]) acc[item.programacao] = [];
      acc[item.programacao].push(item);
      return acc;
    }, {});
    const setMovPage = (p) => {
      const next = Math.max(1, Math.min(totalPages, p));
      setMovListPage((prev) => ({ ...prev, [title]: next }));
    };

    return (
      <PageShell title={title} subtitle={subtitle}>
        <div className="grid grid-cols-1 xl:grid-cols-[440px_1fr] gap-6 items-start">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6 space-y-4">
            <label className="text-sm font-medium">
              Programação
              <input
                value={form.programacao}
                onChange={(e) => setForm({ ...form, programacao: e.target.value })}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              />
            </label>

            <label className="text-sm font-medium">
              Referência / Cor
              <select
                value={`${form.ref}__${form.cor}`}
                onChange={(e) => {
                  const [ref, cor] = e.target.value.split("__");
                  setForm({ ...form, ref, cor });
                }}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                {refs.map((item) => (
                  <option key={item} value={item}>
                    {item.replace("__", " • ")}
                  </option>
                ))}
              </select>
            </label>

            <div>
              <div className="text-sm font-medium mb-2">Grade completa</div>
              <div className="grid grid-cols-4 gap-2">
                {sizes.map((size) => (
                  <div key={size} className="bg-slate-50 rounded-xl p-2 border border-slate-200">
                    <div className="text-xs text-slate-500 text-center">{size}</div>
                    <input
                      value={form.grid[size]}
                      onChange={(e) =>
                        setForm({
                          ...form,
                          grid: { ...form.grid, [size]: Number(e.target.value) || 0 },
                        })
                      }
                      className="w-full text-center mt-1 rounded-lg border border-slate-200 p-1.5 text-sm"
                    />
                  </div>
                ))}
              </div>

              <div className="mt-4 flex items-center justify-between gap-3">
                <div className="text-sm font-semibold">TOTAL: {totalAtual} pares</div>
                <button
                  onClick={() => executeMov(title, form)}
                  className="rounded-2xl bg-[#8B1E2D] text-white px-5 py-3 text-sm font-semibold shadow-sm hover:bg-[#6F1421]"
                >
                  Lançar
                </button>
              </div>
              {isPespontoPage && (
                <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-3 space-y-3">
                  <div className="text-xs font-semibold uppercase tracking-wide text-slate-600">Impressão da ficha</div>
                  <div className="grid grid-cols-1 gap-3">
                    <label className="text-xs font-medium text-slate-700">
                      Texto no cabeçalho de cada ficha (opcional)
                      <input
                        type="text"
                        value={programacaoEtiquetaFicha}
                        onChange={(e) => setProgramacaoEtiquetaFicha(e.target.value)}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm"
                        placeholder="Ex.: FICHA 26-102 - CANO ALTO PRETO"
                      />
                    </label>
                    <label className="text-xs font-medium text-slate-700">
                      Tipo de impressão
                      <select
                        value={programacaoTipoFolha}
                        onChange={(e) => {
                          const v = e.target.value;
                          setProgramacaoTipoFolha(v === "folha2" || v === "ambas" ? v : "folha1");
                        }}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm"
                      >
                        <option value="folha1">Folha 1 - Terceirizados</option>
                        <option value="folha2">Folha 2 - Pesponto</option>
                        <option value="ambas">Ambas - Folha 1 + Folha 2</option>
                      </select>
                    </label>
                    <label className="text-xs font-medium text-slate-700">
                      Cópias da Folha 2
                      <select
                        value={programacaoCopiasPorPagina}
                        onChange={(e) => setProgramacaoCopiasPorPagina(Math.min(4, Math.max(1, Number(e.target.value) || 1)))}
                        disabled={programacaoTipoFolha === "folha1"}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm"
                      >
                        <option value={1}>1 cópia</option>
                        <option value={2}>2 cópias</option>
                        <option value={3}>3 cópias</option>
                        <option value={4}>4 cópias</option>
                      </select>
                    </label>
                    <label className="text-xs font-medium text-slate-700">
                      Cabeçalho da folha (impressão)
                      <select
                        value={programacaoCabecalhoFolha}
                        onChange={(e) => {
                          const v = e.target.value;
                          setProgramacaoCabecalhoFolha(v === "minimo" || v === "oculto" ? v : "completo");
                        }}
                        className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm"
                      >
                        <option value="completo">Completo (logo, título e resumo)</option>
                        <option value="minimo">Mínimo (uma linha)</option>
                        <option value="oculto">Oculto (só as fichas)</option>
                      </select>
                    </label>
                    <button
                      type="button"
                      onClick={() => {
                        const nenhumSelecionado = lancamentosSelecionadosMov.length === 0;
                        if (nenhumSelecionado && (Number(totalAtual) || 0) <= 0) {
                          alert("Selecione pelo menos uma ficha ou informe a grade antes de imprimir.");
                          return;
                        }
                        if (itensMovBase.length > 1 && !nomeProgramacaoMov) {
                          alert("Informe o nome da programação para imprimir várias fichas como programação única.");
                          return;
                        }
                        window.print();
                      }}
                      className="rounded-xl bg-[#0F172A] text-white px-4 py-2.5 text-sm font-semibold"
                    >
                      Imprimir ficha (Folha 1/2)
                    </button>
                    <div className="text-[11px] text-slate-500">
                      {lancamentosSelecionadosMov.length > 0
                        ? `${lancamentosSelecionadosMov.length} ficha(s) selecionada(s) para impressão`
                        : "Sem seleção: usa a ficha atual do formulário"}
                    </div>
                  </div>
                </div>
              )}

              {(liveError || movError[title]) && (
                <div className="mt-3 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-medium text-red-700">
                  {liveError || movError[title]}
                </div>
              )}
            </div>
          </div>

          <div className="space-y-6">
            {selectionPreview && (
              <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
                <div className="text-sm font-bold uppercase tracking-wide text-[#8B1E2D]">Prévia da cor</div>
                <div className="mt-3 text-2xl font-semibold text-slate-900">
                  {selectionPreview.row.ref} • {selectionPreview.row.cor}
                </div>

                <div className="mt-6 grid grid-cols-2 lg:grid-cols-4 gap-4 text-sm text-slate-700">
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <div>
                      <div className="text-2xl font-black text-[#8B1E2D] tracking-tight">PA</div>
                      <div className="text-xs text-slate-500 mt-0.5">Pronto acabado</div>
                    </div>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalPA}</span>
                  </div>
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <div>
                      <div className="text-2xl font-black text-[#8B1E2D] tracking-tight">P</div>
                      <div className="text-xs text-slate-500 mt-0.5">Pesponto</div>
                    </div>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalPesponto}</span>
                  </div>
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <div>
                      <div className="text-2xl font-black text-[#8B1E2D] tracking-tight">EST</div>
                      <div className="text-xs text-slate-500 mt-0.5">Costura pronta</div>
                    </div>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalEst}</span>
                  </div>
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <div>
                      <div className="text-2xl font-black text-[#8B1E2D] tracking-tight">M</div>
                      <div className="text-xs text-slate-500 mt-0.5">Montagem</div>
                    </div>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalMontagem}</span>
                  </div>
                </div>

                <div className="mt-6 grid grid-cols-4 md:grid-cols-6 xl:grid-cols-8 gap-3 text-center text-sm text-slate-600">
                  {sizes.map((size) => (
                    <div key={size} className="rounded-2xl border border-sky-100 bg-slate-50 px-3 py-3">
                      <div className="font-semibold text-slate-900">{size}</div>
                      <div className="mt-2 space-y-0.5 text-xs text-slate-700">
                        <div>
                          <span className="font-bold text-[#8B1E2D]">PA</span> {selectionPreview.row.data[size]?.pa || 0}
                        </div>
                        <div>
                          <span className="font-bold text-[#8B1E2D]">P</span> {selectionPreview.row.data[size]?.p || 0}
                        </div>
                        <div>
                          <span className="font-bold text-[#8B1E2D]">EST</span> {selectionPreview.row.data[size]?.est || 0}
                        </div>
                        <div>
                          <span className="font-bold text-[#8B1E2D]">M</span> {selectionPreview.row.data[size]?.m || 0}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>

                <div className="mt-6 border-t border-slate-200 pt-6">
                  <div className="text-sm font-bold uppercase tracking-wide text-slate-700">Lançamentos em aberto nesta cor</div>
                  <p className="text-xs text-slate-500 mt-1">
                    Mesma referência e cor do formulário · ainda não finalizados no {title}.
                  </p>
                  {lancamentosAbertosCor.length === 0 ? (
                    <div className="mt-3 rounded-2xl border border-dashed border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-500">
                      Nenhum lançamento em aberto para {selectionPreview.row.ref} • {selectionPreview.row.cor}.
                    </div>
                  ) : (
                    <div className="mt-3 space-y-3">
                      {lancamentosAbertosCor.map((item) => (
                        <div
                          key={item.id}
                          className="rounded-2xl border border-amber-200 bg-amber-50/60 px-4 py-3"
                        >
                          <div className="flex flex-wrap items-start justify-between gap-2">
                            <div className="font-semibold text-slate-900">{item.programacao}</div>
                            <span className="px-2.5 py-0.5 rounded-full border border-amber-300 bg-amber-100 text-xs font-semibold text-amber-900">
                              Em aberto
                            </span>
                          </div>
                          <div className="text-xs text-slate-600 mt-1">
                            {item.total ?? 0} pares · Lançado {item.dataLancamento || "—"}
                          </div>
                          <div className="mt-2 flex flex-wrap gap-1.5">
                            {(item.items || []).map((entry) => (
                              <span
                                key={`${item.id}-${entry.size}`}
                                className="inline-flex px-2 py-1 rounded-lg bg-white border border-amber-100 text-xs text-slate-800"
                              >
                                {entry.size} → {entry.qtd}
                              </span>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}

            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between mb-4">
                <div>
                  <h2 className="font-bold text-lg">Lançamentos enviados</h2>
                  <p className="text-sm text-slate-500 mt-1">Finalize para atualizar o estoque.</p>
                </div>
                <div className="text-sm text-slate-500 text-left sm:text-right shrink-0">
                  <div>{lancamentos.length} lançamento(s) no total</div>
                  {lancamentos.length > 0 && (
                    <div className="text-xs text-slate-400 mt-0.5">
                      Mostrando {startIdx + 1}–{Math.min(startIdx + lancamentosPagina.length, lancamentos.length)} · {MOV_PAGE_SIZE} por página
                    </div>
                  )}
                </div>
              </div>
              {isPespontoPage && (
                <div className="mb-4 flex flex-wrap items-center gap-2 rounded-2xl border border-slate-200 bg-slate-50 px-3 py-2">
                  <button
                    type="button"
                    onClick={() =>
                      setMovImpressaoSelecao((prev) => ({
                        ...prev,
                        [title]: Object.fromEntries(lancamentos.map((item) => [item.id, true])),
                      }))
                    }
                    className="px-3 py-1.5 text-xs font-semibold rounded-xl border border-slate-200 bg-white hover:bg-slate-100"
                  >
                    Marcar todas p/ impressão
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      setMovImpressaoSelecao((prev) => ({
                        ...prev,
                        [title]: {},
                      }))
                    }
                    className="px-3 py-1.5 text-xs font-semibold rounded-xl border border-slate-200 bg-white hover:bg-slate-100"
                  >
                    Limpar seleção
                  </button>
                  <span className="text-xs text-slate-600 ml-auto">
                    {lancamentosSelecionadosMov.length} selecionada(s) para impressão
                  </span>
                </div>
              )}

              {!Object.keys(agrupados).length ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500">
                  Nenhum lançamento ainda.
                </div>
              ) : (
                <div className="space-y-5">
                  {Object.entries(agrupados).map(([programacao, items]) => (
                    <div key={programacao} className="rounded-2xl border overflow-hidden border-slate-200">
                      <div className="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-3">
                        <div>
                          <div className="font-semibold">{programacao}</div>
                        </div>
                        <div className="flex items-center gap-2">
                          {isPespontoPage && (
                            <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer mr-1">
                              <input
                                type="checkbox"
                                className="h-4 w-4 rounded border-slate-300 text-[#8B1E2D]"
                                checked={items.every((item) => selecaoMovAtual[item.id] === true)}
                                onChange={(e) =>
                                  setMovImpressaoSelecao((prev) => {
                                    const base = { ...(prev[title] || {}) };
                                    items.forEach((item) => {
                                      base[item.id] = e.target.checked;
                                    });
                                    return {
                                      ...prev,
                                      [title]: base,
                                    };
                                  })
                                }
                              />
                            </label>
                          )}
                          <div className="text-sm text-slate-500">
                            {items.reduce((acc, item) => acc + item.total, 0)} pares
                          </div>
                          <button
                            onClick={() => finalizarProgramacao(title, programacao)}
                            className="px-3 py-1.5 text-xs font-semibold rounded-xl bg-emerald-100 text-emerald-700 border border-emerald-200"
                          >
                            Finalizar programação
                          </button>
                        </div>
                      </div>

                      <div className="p-4 space-y-4">
                        {items.map((item) => (
                          <div key={item.id} className="rounded-xl border border-slate-200 p-4">
                            <div className="flex items-start justify-between gap-4">
                              <div>
                                <div className="font-medium">{item.ref} • {item.cor}</div>
                                <div className="text-xs text-slate-500 mt-1">Lançado em {item.dataLancamento}</div>
                              </div>
                              <span
                                className={`px-3 py-1 rounded-full border text-xs font-semibold ${
                                  item.status === "Finalizado"
                                    ? "bg-emerald-100 text-emerald-700 border-emerald-200"
                                    : "bg-amber-100 text-amber-700 border-amber-200"
                                }`}
                              >
                                {item.status}
                              </span>
                            </div>

                            <div className="mt-3 flex flex-wrap gap-2">
                              {item.items.map((entry) => (
                                <span
                                  key={`${item.id}-${entry.size}`}
                                  className="px-3 py-1.5 rounded-xl bg-slate-50 border border-slate-200 text-sm"
                                >
                                  {entry.size} → {entry.qtd}
                                </span>
                              ))}
                            </div>

                            {item.status !== "Finalizado" && (
                              <div className="mt-4 flex gap-2">
                                <button
                                  onClick={() => startEditLancamento(title, item)}
                                  className="px-3 py-1.5 text-xs font-semibold rounded-xl bg-[#FCECEE] text-[#8B1E2D] border border-[#E7C7CC]"
                                >
                                  Alterar
                                </button>
                                <button
                                  onClick={() => deleteLancamento(title, item.id)}
                                  className="px-3 py-1.5 text-xs font-semibold rounded-xl bg-red-100 text-red-700 border border-red-200"
                                >
                                  Excluir
                                </button>
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {lancamentos.length > MOV_PAGE_SIZE && (
                <div className="mt-6 flex flex-col items-stretch gap-4 sm:flex-row sm:items-center sm:justify-between pt-4 border-t border-slate-200">
                  <div className="text-sm text-slate-600 text-center sm:text-left">
                    Página <span className="font-semibold text-slate-900">{currentPage}</span> de{" "}
                    <span className="font-semibold text-slate-900">{totalPages}</span>
                  </div>
                  <div className="flex flex-wrap items-center justify-center gap-2">
                    <button
                      type="button"
                      onClick={() => setMovPage(currentPage - 1)}
                      disabled={currentPage <= 1}
                      className="px-3 py-2 rounded-xl text-sm font-semibold border border-slate-200 bg-white disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-50"
                    >
                      Anterior
                    </button>
                    <div className="flex flex-wrap items-center justify-center gap-1 max-w-full">
                      {totalPages <= 12
                        ? Array.from({ length: totalPages }, (_, i) => i + 1).map((num) => (
                            <button
                              key={num}
                              type="button"
                              onClick={() => setMovPage(num)}
                              className={`min-w-[2.25rem] px-2 py-2 rounded-xl text-sm font-semibold border ${
                                num === currentPage
                                  ? "bg-[#0F172A] text-white border-[#0F172A]"
                                  : "bg-white text-slate-700 border-slate-200 hover:bg-slate-50"
                              }`}
                            >
                              {num}
                            </button>
                          ))
                        : (
                            <span className="text-sm text-slate-500 px-2">
                              Use Anterior / Próxima ({totalPages} páginas)
                            </span>
                          )}
                    </div>
                    <button
                      type="button"
                      onClick={() => setMovPage(currentPage + 1)}
                      disabled={currentPage >= totalPages}
                      className="px-3 py-2 rounded-xl text-sm font-semibold border border-slate-200 bg-white disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-50"
                    >
                      Próxima
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
        {isPespontoPage && (
          <div id="print-mov-root" className="hidden print:block programacao-print-sheet">
            {programacaoTipoFolha === "ambas" ? (
              <>
                <ProgramacaoDiaFolhaImpressao
                  titulo={`${tituloFolhaMov} - Folha 1`}
                  logoSrc={programacaoLogoImpressao?.trim() || "/logo-rockstar-bandeira.png"}
                  setor="Pesponto"
                  modoLabel="Ficha direta da aba Pesponto"
                  diasCount={1}
                  dataImpressao={new Date().toLocaleString("pt-BR")}
                  observacoes={programacaoObsImpressao}
                  itens={itensMovFolha1}
                  sizesList={sizes}
                  copiasPorPagina={3}
                  etiquetaFichaCustom={programacaoEtiquetaFicha}
                  cabecalhoFolha={programacaoCabecalhoFolha}
                  valoresParTerceiros={programacaoValoresTerceiros}
                  tipoFolhaImpressao="folha1"
                />
                <div className="programacao-print-item-break" />
                <ProgramacaoDiaFolhaImpressao
                  titulo={`${tituloFolhaMov} - Folha 2`}
                  logoSrc={programacaoLogoImpressao?.trim() || "/logo-rockstar-bandeira.png"}
                  setor="Pesponto"
                  modoLabel="Ficha direta da aba Pesponto"
                  diasCount={1}
                  dataImpressao={new Date().toLocaleString("pt-BR")}
                  observacoes={programacaoObsImpressao}
                  itens={itensMovFolha2}
                  sizesList={sizes}
                  copiasPorPagina={programacaoCopiasPorPagina}
                  etiquetaFichaCustom={programacaoEtiquetaFicha}
                  cabecalhoFolha={programacaoCabecalhoFolha}
                  valoresParTerceiros={programacaoValoresTerceiros}
                  tipoFolhaImpressao="folha2"
                />
              </>
            ) : (
              <ProgramacaoDiaFolhaImpressao
                titulo={tituloFolhaMov}
                logoSrc={programacaoLogoImpressao?.trim() || "/logo-rockstar-bandeira.png"}
                setor="Pesponto"
                modoLabel="Ficha direta da aba Pesponto"
                diasCount={1}
                dataImpressao={new Date().toLocaleString("pt-BR")}
                observacoes={programacaoObsImpressao}
                itens={itensMovImpressao}
                sizesList={sizes}
                copiasPorPagina={programacaoCopiasPorPagina}
                etiquetaFichaCustom={programacaoEtiquetaFicha}
                cabecalhoFolha={programacaoCabecalhoFolha}
                valoresParTerceiros={programacaoValoresTerceiros}
                tipoFolhaImpressao={programacaoTipoFolha}
              />
            )}
          </div>
        )}
      </PageShell>
    );
  };

  const aplicarAjusteEst = async () => {
  const qtd = Number(ajusteEstForm.qtd) || 0;
  if (!qtd) {
    setAjusteEstErro("Informe uma quantidade válida para o ajuste.");
    return;
  }

  const row = rows.find((item) => item.ref === ajusteEstForm.ref && item.cor === ajusteEstForm.cor);
  const atual = row?.data?.[ajusteEstForm.size]?.est || 0;

  if (ajusteEstForm.tipo === "saida" && qtd > atual) {
    setAjusteEstErro(`A saída não pode ser maior que o saldo atual da costura pronta (${atual}).`);
    return;
  }

  setAjusteEstErro("");

  const delta = ajusteEstForm.tipo === "entrada" ? qtd : -qtd;

  const nextRows = rows.map((item) => {
    if (item.ref !== ajusteEstForm.ref || item.cor !== ajusteEstForm.cor) return item;

    const nextData = { ...item.data };
    const atualNumero = nextData[ajusteEstForm.size] || { pa: 0, est: 0, m: 0, p: 0 };

    nextData[ajusteEstForm.size] = {
      ...atualNumero,
      est: Math.max(0, (atualNumero.est || 0) + delta),
    };

    return {
      ...item,
      data: nextData,
    };
  });

  const persistAjuste = await persistRowsToSupabase(nextRows);
  if (!persistAjuste.ok) {
    setAjusteEstErro(
      `Erro ao salvar o ajuste no Supabase: ${persistAjuste.error?.message || persistAjuste.error || "tente novamente"}.`
    );
    return;
  }

  setRows(nextRows);

  const ajuste = {
    id: `ajuste-est-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    data: new Date().toLocaleDateString("pt-BR"),
    dataLancamento: new Date().toLocaleDateString("pt-BR"),
    ref: ajusteEstForm.ref,
    cor: ajusteEstForm.cor,
    tipo: ajusteEstForm.tipo,
    size: ajusteEstForm.size,
    qtd,
    motivo: ajusteEstForm.motivo || "Sem motivo informado",
  };

  setAjustesEst((curr) => [ajuste, ...curr]);

  await salvarMovimentacao([
    {
      tipo: "Costura Pronta",
      ref: ajusteEstForm.ref,
      cor: ajusteEstForm.cor,
      numero: ajusteEstForm.size,
      quantidade: qtd,
      programacao: `Ajuste manual - ${ajusteEstForm.tipo}`,
      status: "Lançado",
    },
  ]);

  setAjusteEstForm((curr) => ({ ...curr, qtd: 0, motivo: "" }));
};

  const renderCosturaPronta = () => (
    <PageShell title="Costura Pronta" subtitle="Etapa intermediária alimentada pelo Pesponto finalizado.">
      <div className="grid grid-cols-1 xl:grid-cols-[1fr_380px] gap-6 items-start">
        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-auto p-6">
          <table className="min-w-[1200px] w-full border-collapse text-sm">
            <thead>
              <tr className="bg-slate-50 text-slate-500">
                <th className="border px-4 py-3 text-left">Ref</th>
                <th className="border px-4 py-3 text-left">Cor</th>
                {sizes.map((s) => <th key={s} className="border px-3 py-3 text-center">{s}</th>)}
                <th className="border px-4 py-3 text-center">Total Est Pto</th>
              </tr>
            </thead>
            <tbody>
              {sortedRowsByRefCor.map((row) => (
                <tr key={`cp-${row.ref}-${row.cor}`}>
                  <td className="border px-4 py-3 font-semibold">{row.ref}</td>
                  <td className="border px-4 py-3">{row.cor}</td>
                  {sizes.map((size) => (
                    <td key={size} className="border px-3 py-3 text-center">{row.data[size]?.est ?? 0}</td>
                  ))}
                  <td className="border px-4 py-3 text-center font-bold">
                    {sizes.reduce((acc, size) => acc + (row.data[size]?.est ?? 0), 0)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div className="space-y-6">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="font-bold text-lg">Ajuste manual</div>
            <p className="text-sm text-slate-500 mt-1">Use para adicionar ou retirar pares da costura pronta.</p>

            <div className="mt-4 space-y-4">
              <label className="text-sm font-medium">
                Referência / Cor
                <select
                  value={`${ajusteEstForm.ref}__${ajusteEstForm.cor}`}
                  onChange={(e) => {
                    const [ref, cor] = e.target.value.split("__");
                    setAjusteEstForm((curr) => ({ ...curr, ref, cor }));
                  }}
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                >
                  {refs.map((item) => (
                    <option key={item} value={item}>{item.replace("__", " • ")}</option>
                  ))}
                </select>
              </label>

              <div className="grid grid-cols-2 gap-3">
                <label className="text-sm font-medium">
                  Numeração
                  <select
                    value={ajusteEstForm.size}
                    onChange={(e) => setAjusteEstForm((curr) => ({ ...curr, size: Number(e.target.value) }))}
                    className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                  >
                    {sizes.map((size) => <option key={size} value={size}>{size}</option>)}
                  </select>
                </label>

                <label className="text-sm font-medium">
                  Quantidade
                  <input
                    type="number"
                    min="0"
                    value={ajusteEstForm.qtd}
                    onChange={(e) => setAjusteEstForm((curr) => ({ ...curr, qtd: Number(e.target.value) || 0 }))}
                    className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                  />
                </label>
              </div>

              <label className="text-sm font-medium">
                Motivo
                <input
                  value={ajusteEstForm.motivo}
                  onChange={(e) => setAjusteEstForm((curr) => ({ ...curr, motivo: e.target.value }))}
                  placeholder="Ex.: avaria, perda, reposição, contagem"
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                />
              </label>

              <div className="grid grid-cols-2 gap-3">
                <button
                  onClick={() => setAjusteEstForm((curr) => ({ ...curr, tipo: "entrada" }))}
                  className={`rounded-2xl px-4 py-3 text-sm font-semibold border ${ajusteEstForm.tipo === "entrada" ? "bg-emerald-100 text-emerald-700 border-emerald-200" : "bg-slate-50 text-slate-700 border-slate-200"}`}
                >
                  Adição manual
                </button>
                <button
                  onClick={() => setAjusteEstForm((curr) => ({ ...curr, tipo: "saida" }))}
                  className={`rounded-2xl px-4 py-3 text-sm font-semibold border ${ajusteEstForm.tipo === "saida" ? "bg-red-100 text-red-700 border-red-200" : "bg-slate-50 text-slate-700 border-slate-200"}`}
                >
                  Saída manual
                </button>
              </div>

              {ajusteEstErro && (
                <div className="rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-medium text-red-700">
                  {ajusteEstErro}
                </div>
              )}

              <button
                onClick={aplicarAjusteEst}
                className="w-full rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold"
              >
                Aplicar ajuste
              </button>
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <div>
                <div className="font-bold text-lg">Histórico de ajustes manuais</div>
                <div className="text-sm text-slate-500 mt-1">Entradas e saídas registradas como movimentações Costura Pronta no banco.</div>
              </div>
              <span className="text-sm text-slate-500">{ajustesEst.length} ajuste(s)</span>
            </div>

            <div className="mt-4 space-y-3 max-h-[420px] overflow-auto pr-1">
              {ajustesEst.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                  Nenhum ajuste manual registrado ainda.
                </div>
              ) : (
                ajustesEst.map((ajuste) => (
                  <div key={ajuste.id} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex items-start justify-between gap-3">
                      <div>
                        <div className="font-semibold">{ajuste.ref} • {ajuste.cor}</div>
                        <div className="text-xs text-slate-500 mt-1">{ajuste.dataLancamento || ajuste.data} • Numeração {ajuste.size}</div>
                      </div>
                      <span className={`px-3 py-1 rounded-full border text-xs font-semibold ${ajuste.tipo === "entrada" ? "bg-emerald-100 text-emerald-700 border-emerald-200" : "bg-red-100 text-red-700 border-red-200"}`}>
                        {ajuste.tipo === "entrada" ? `+${ajuste.qtd}` : `-${ajuste.qtd}`}
                      </span>
                    </div>
                    <div className="text-sm text-slate-600 mt-2">Motivo: {ajuste.motivo}</div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      </div>

      <section className="mt-8 bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
        <div className="flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between">
          <div>
            <h2 className="font-bold text-lg">Pesponto finalizado → costura pronta</h2>
            <p className="text-sm text-slate-500 mt-1">
              Histórico das programações de pesponto já finalizadas (transferência para a coluna EST na grade acima).
            </p>
          </div>
          <span className="text-sm font-medium text-slate-500 shrink-0">
            {pespontoFinalizadosHistorico.length} finalização(ões)
          </span>
        </div>
        <div className="mt-4 space-y-3 max-h-[400px] overflow-auto pr-1">
          {pespontoFinalizadosHistorico.length === 0 ? (
            <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
              Nenhuma programação de pesponto finalizada encontrada. Finalize uma programação na aba Pesponto para ver o registro aqui.
            </div>
          ) : (
            pespontoFinalizadosHistorico.map((item) => (
              <div key={item.id} className="rounded-2xl border border-slate-200 px-4 py-3 bg-slate-50/80">
                <div className="flex flex-wrap items-start justify-between gap-2">
                  <div>
                    <div className="font-semibold text-slate-900">
                      {item.ref} • {item.cor}
                    </div>
                    <div className="text-sm text-slate-600 mt-0.5">{item.programacao}</div>
                  </div>
                  <span className="px-3 py-1 rounded-full border text-xs font-semibold bg-emerald-100 text-emerald-800 border-emerald-200">
                    Finalizado {item.dataFinalizacao || "—"}
                  </span>
                </div>
                <div className="mt-2 text-xs text-slate-500">
                  Lançado em {item.dataLancamento || "—"} • Total {item.total ?? "—"} pares
                </div>
                <div className="mt-2 text-xs text-slate-600 leading-relaxed">
                  Grade:{" "}
                  {(item.items || []).length
                    ? item.items.map((it) => `${it.size}: ${it.qtd}`).join(" · ")
                    : "—"}
                </div>
              </div>
            ))
          )}
        </div>
      </section>
    </PageShell>
  );

  const renderConfig = () => {
  const updateLocal = (ref, cor, size, key, value) => {
    setDraftMinimos((curr) => ({
      ...curr,
      [ref]: {
        ...(curr?.[ref] || {}),
        [cor]: {
          ...(curr?.[ref]?.[cor] || {}),
          [size]: {
            ...((curr?.[ref]?.[cor]?.[size]) || { pa: 0, prod: 0 }),
            [key]: Number(value) || 0,
          },
        },
      },
    }));
    setDirtyMinimos(true);
  };

  const salvar = async () => {
  const valorWevertonValido = parseDecimalInput(programacaoValoresTerceiros.weverton);
  const valorRomuloValido = parseDecimalInput(programacaoValoresTerceiros.romulo);
  if (String(programacaoValoresTerceiros.weverton ?? "").trim() && !Number.isFinite(valorWevertonValido)) {
    alert("Valor de Weverton inválido. Use número com ponto ou vírgula (ex.: 0,70).");
    return;
  }
  if (String(programacaoValoresTerceiros.romulo ?? "").trim() && !Number.isFinite(valorRomuloValido)) {
    alert("Valor de Romulo inválido. Use número com ponto ou vírgula (ex.: 0,50).");
    return;
  }
  setMinimos(draftMinimos);
  setTempoProducao(tempoProducaoDraft);

  await salvarMinimosNoBanco(draftMinimos);

  await salvarConfiguracoesProducaoNoBanco({
    capacidadePespontoDia,
    capacidadeMontagemDia,
    diasPesponto: tempoProducaoDraft.pesponto,
    diasMontagem: tempoProducaoDraft.montagem,
    reservaTopPct: programacaoReservaTopPct,
    topN: programacaoTopN,
    topMode: programacaoTopModo,
    topManualKeys: programacaoTopManualKeys,
    valorParWeverton: programacaoValoresTerceiros.weverton,
    valorParRomulo: programacaoValoresTerceiros.romulo,
  });

  setDirtyMinimos(false);
};

  return (
    <PageShell title="Minimos" subtitle="Aba equivalente aos mínimos da planilha, com PA e PROD por referência, cor e numeração.">
      <div className="space-y-4">

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex items-center justify-between gap-3 mb-4">
            <div>
              <h2 className="font-bold text-lg">Capacidade diária de produção</h2>
              <p className="text-sm text-slate-500 mt-1">
                Ajuste quantos pares cada setor consegue produzir por dia.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <label className="text-sm font-medium text-slate-700">
              Pesponto por dia
              <input
                type="number"
                min="0"
                value={capacidadePespontoDia}
                onChange={(e) => setCapacidadePespontoDia(Number(e.target.value) || 0)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
              />
            </label>

            <label className="text-sm font-medium text-slate-700">
              Montagem por dia
              <input
                type="number"
                min="0"
                value={capacidadeMontagemDia}
                onChange={(e) => setCapacidadeMontagemDia(Number(e.target.value) || 0)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
              />
            </label>
          </div>
        </div>

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex items-center justify-between gap-4 mb-4">
            <div>
              <div className="font-bold text-lg">Tempo de Produção</div>
              <div className="text-sm text-slate-500 mt-1">Esses tempos entram na lógica das sugestões para antecipar risco de falta antes de virar PA.</div>
            </div>
            <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-right">
              <div className="text-xs text-slate-500">Tempo total</div>
              <div className="text-2xl font-bold text-slate-900">
                {(Number(tempoProducaoDraft.pesponto) || 0) + (Number(tempoProducaoDraft.montagem) || 0)} dias
              </div>
            </div>
          </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <label className="text-sm font-medium">
                Tempo de Pesponto (dias)
                <input
                  type="number"
                  min="0"
                  value={tempoProducaoDraft.pesponto}
                  onChange={(e) => {
                    setTempoProducaoDraft((curr) => ({ ...curr, pesponto: Number(e.target.value) || 0 }));
                    setDirtyMinimos(true);
                  }}
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                />
              </label>

              <label className="text-sm font-medium">
                Tempo de Montagem (dias)
                <input
                  type="number"
                  min="0"
                  value={tempoProducaoDraft.montagem}
                  onChange={(e) => {
                    setTempoProducaoDraft((curr) => ({ ...curr, montagem: Number(e.target.value) || 0 }));
                    setDirtyMinimos(true);
                  }}
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                />
              </label>
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-4 mb-4">
              <div>
                <div className="font-bold text-lg">Configuração de pagamentos</div>
                <div className="text-sm text-slate-500 mt-1">
                  Valores por par usados na Folha 1 (terceirizados). O GU segue regra fixa por referência.
                </div>
              </div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <label className="text-sm font-medium text-slate-700">
                Weverton (valor por par)
                <input
                  type="text"
                  inputMode="decimal"
                  value={programacaoValoresTerceiros.weverton}
                  onChange={(e) => {
                    setProgramacaoValoresTerceiros((curr) => ({ ...curr, weverton: e.target.value }));
                    setDirtyMinimos(true);
                  }}
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
                />
              </label>
              <label className="text-sm font-medium text-slate-700">
                Romulo (valor por par)
                <input
                  type="text"
                  inputMode="decimal"
                  value={programacaoValoresTerceiros.romulo}
                  onChange={(e) => {
                    setProgramacaoValoresTerceiros((curr) => ({ ...curr, romulo: e.target.value }));
                    setDirtyMinimos(true);
                  }}
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
                />
              </label>
            </div>
          </div>

          <div className="flex justify-end">
            <button
              onClick={salvar}
              disabled={!dirtyMinimos}
              className={`px-5 py-3 rounded-2xl text-sm font-semibold ${dirtyMinimos ? "bg-[#8B1E2D] text-white hover:bg-[#6F1421]" : "bg-slate-200 text-slate-500 cursor-not-allowed"}`}
            >
              Salvar alterações
            </button>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-auto p-6">
            <table className="min-w-[1700px] w-full border-collapse">
              <thead>
                <tr className="bg-slate-50">
                  <th rowSpan={2} className="border border-slate-200 px-4 py-3 text-left">Ref</th>
                  <th rowSpan={2} className="border border-slate-200 px-4 py-3 text-left">Cor</th>
                  {sizes.map((size) => (
                    <th key={size} colSpan={2} className="border border-slate-200 px-4 py-3 text-center">{size}</th>
                  ))}
                </tr>
                <tr className="bg-slate-50 text-xs text-slate-500">
                  {sizes.flatMap((size) => [
                    <th key={`${size}-pa`} className="border border-slate-200 px-3 py-2">PA</th>,
                    <th key={`${size}-prod`} className="border border-slate-200 px-3 py-2">PROD</th>,
                  ])}
                </tr>
              </thead>
              <tbody>
                {sortedRowsByRefCor.map((row) => (
                  <tr key={`min-${row.ref}-${row.cor}`}>
                    <td className="border border-slate-200 px-4 py-3 font-semibold">{row.ref}</td>
                    <td className="border border-slate-200 px-4 py-3">{row.cor}</td>
                    {sizes.flatMap((size) => [
                      <td key={`${row.ref}-${size}-pa`} className="border border-slate-200 px-2 py-2">
                        <input
                          value={draftMinimos?.[row.ref]?.[row.cor]?.[size]?.pa ?? minimos?.[row.ref]?.[row.cor]?.[size]?.pa ?? 0}
                          onChange={(e) => updateLocal(row.ref, row.cor, size, "pa", e.target.value)}
                          className="w-16 rounded-lg border border-slate-200 p-2 text-center text-sm"
                        />
                      </td>,
                      <td key={`${row.ref}-${size}-prod`} className="border border-slate-200 px-2 py-2">
                        <input
                          value={draftMinimos?.[row.ref]?.[row.cor]?.[size]?.prod ?? minimos?.[row.ref]?.[row.cor]?.[size]?.prod ?? 0}
                          onChange={(e) => updateLocal(row.ref, row.cor, size, "prod", e.target.value)}
                          className="w-16 rounded-lg border border-slate-200 p-2 text-center text-sm"
                        />
                      </td>,
                    ])}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </PageShell>
    );
  };

  const renderVendas = () => (
    <PageShell title="Vendas" subtitle="Os produtos desta tela seguem a mesma base importada no GCM e alimentam a prioridade da produção.">
      <div className="space-y-6">
        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex items-center justify-between">
            <div>
              <div className="font-semibold">Importar vendas (.xls)</div>
              <div className="text-sm text-slate-500">Arquivo com PRODUTO, PADRONIZAÇÃO e QTDE</div>
            </div>
            <label className="cursor-pointer bg-slate-950 text-white px-4 py-3 rounded-2xl text-sm font-semibold">
              Escolher arquivo
              <input type="file" accept=".xls,.xlsx" onChange={handleSalesFileUpload} className="hidden" />
            </label>
          </div>
          <div className="mt-4 rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-sm text-slate-600">
            Produtos disponíveis em vendas: <span className="font-semibold">{rows.length}</span> item(ns), seguindo as referências e cores importadas no GCM.
          </div>
          {salesImportFileName && <div className="mt-4 text-sm">Arquivo: <span className="font-semibold">{salesImportFileName}</span></div>}
          {salesImportFeedback && <div className="mt-3 text-sm text-slate-600">{salesImportFeedback}</div>}
        </div>

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex items-center justify-between gap-4 mb-4">
            <div>
              <div className="font-semibold">Resumo do que foi importado</div>
              <div className="text-sm text-slate-500 mt-1">Consolidação por referência, cor e numeração.</div>
            </div>
            <div className="text-sm text-slate-500">{salesImportPreview.length} item(ns)</div>
          </div>
          {salesImportPreview.length === 0 ? (
            <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500">Nenhum resumo de vendas importado ainda.</div>
          ) : (
            <div className="overflow-auto">
              <table className="min-w-[1200px] w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-slate-50 text-slate-500">
                    <th className="border px-4 py-3 text-left">Ref</th>
                    <th className="border px-4 py-3 text-left">Cor</th>
                    {sizes.map((s) => <th key={s} className="border px-3 py-3 text-center">{s}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {salesImportPreview.map((item, idx) => (
                    <tr key={`${item.ref}-${item.cor}-${idx}`}>
                      <td className="border px-4 py-3 font-semibold">{item.ref}</td>
                      <td className="border px-4 py-3">{item.cor}</td>
                      {sizes.map((size) => <td key={size} className="border px-3 py-3 text-center">{item.data[size]}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div className="grid grid-cols-1 xl:grid-cols-[1fr_360px] gap-6 items-start">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-auto p-6">
            <div className="flex items-center justify-between gap-4 mb-4">
              <div>
                <div className="font-semibold">Lançamento manual de vendas</div>
                <div className="text-sm text-slate-500 mt-1">Mesma base de referências e cores vindas do Importar GCM.</div>
              </div>
              <div className="flex items-center gap-3">
                <div className="text-sm text-slate-500">{rows.length} produto(s)</div>
                <button
                  onClick={salvarVendasManuais}
                  disabled={!vendasDirty}
                  className={`px-5 py-3 rounded-2xl text-sm font-semibold ${vendasDirty ? "bg-[#8B1E2D] text-white hover:bg-[#6F1421]" : "bg-slate-200 text-slate-500 cursor-not-allowed"}`}
                >
                  Salvar vendas
                </button>
              </div>
            </div>
            <table className="min-w-[1200px] w-full border-collapse text-sm">
              <thead>
                <tr className="bg-slate-50 text-slate-500">
                  <th className="border px-4 py-3 text-left">Ref</th>
                  <th className="border px-4 py-3 text-left">Cor</th>
                  {sizes.map((s) => <th key={s} className="border px-3 py-3 text-center">{s}</th>)}
                </tr>
              </thead>
              <tbody>
                {sortedRowsByRefCor.map((row) => {
                  const vendaGrid = vendaGridForRow(row);
                  return (
                    <tr key={`vendas-${row.ref}-${row.cor}`}>
                      <td className="border px-4 py-3 font-semibold">{row.ref}</td>
                      <td className="border px-4 py-3">{row.cor}</td>
                      {sizes.map((size) => (
                        <td key={size} className="border px-2 py-2 text-center">
                          <input
                            value={vendaGrid[size]}
                            onChange={(e) => updateVenda(row.ref, row.cor, size, e.target.value)}
                            className="w-16 rounded-lg border border-slate-200 p-2 text-center text-sm"
                          />
                        </td>
                      ))}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="space-y-6">
            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
              <div className="flex items-center justify-between gap-3">
                <div>
                  <div className="font-bold text-lg">Top vendas manual</div>
                  <div className="text-sm text-slate-500 mt-1">
                    Marque aqui as ref+cor usadas no modo manual da Programação do Dia.
                  </div>
                </div>
                <span className="text-sm text-slate-500">{programacaoTopManualKeys.length} selecionada(s)</span>
              </div>
              <div className="mt-4 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={() => setProgramacaoTopManualKeys(topVendasCadastroOpcoes.map((item) => item.key))}
                  className="px-3 py-1.5 rounded-xl text-xs font-semibold border border-slate-200 bg-white hover:bg-slate-100"
                >
                  Marcar todas
                </button>
                <button
                  type="button"
                  onClick={() => setProgramacaoTopManualKeys([])}
                  className="px-3 py-1.5 rounded-xl text-xs font-semibold border border-slate-200 bg-white hover:bg-slate-100"
                >
                  Limpar todas
                </button>
              </div>
              <div className="mt-3 max-h-72 overflow-auto space-y-2 pr-1">
                {topVendasCadastroOpcoes.map((item) => {
                  const checked = programacaoTopManualKeys.includes(item.key);
                  return (
                    <label
                      key={`top-cadastro-${item.key}`}
                      className="flex items-center justify-between gap-3 rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm"
                      title={`${item.ref} • ${item.cor}`}
                    >
                      <span className="flex items-start gap-2 min-w-0">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() =>
                            setProgramacaoTopManualKeys((prev) =>
                              prev.includes(item.key) ? prev.filter((k) => k !== item.key) : [...prev, item.key]
                            )
                          }
                          className="mt-0.5"
                        />
                        <span className="font-medium text-slate-800 whitespace-normal break-words leading-snug">{item.ref} • {item.cor}</span>
                      </span>
                      <span className="text-xs text-slate-500 tabular-nums">{item.vendaTotal} venda(s)</span>
                    </label>
                  );
                })}
              </div>
            </div>

            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
              <div className="flex items-center justify-between gap-3">
                <div>
                  <div className="font-bold text-lg">Histórico de vendas manuais</div>
                  <div className="text-sm text-slate-500 mt-1">Sempre mostra o último salvamento manual.</div>
                </div>
                <span className="text-sm text-slate-500">{historicoVendasManuais.length} registro(s)</span>
              </div>

              <div className="mt-4 space-y-4">
                {historicoVendasManuais.length === 0 ? (
                  <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                    Nenhum salvamento manual realizado ainda.
                  </div>
                ) : (
                  historicoVendasManuais.map((registro) => (
                    <div key={registro.id} className="rounded-2xl border border-slate-200 overflow-hidden">
                      <div className="bg-slate-50 border-b border-slate-200 px-4 py-3">
                        <div className="font-semibold text-slate-900">Salvamento manual</div>
                        <div className="text-sm text-slate-500 mt-1">{registro.dataHora}</div>
                      </div>
                      <div className="p-4 space-y-3">
                        <div className="grid grid-cols-2 gap-3 text-sm">
                          <div className="rounded-xl bg-slate-50 border border-slate-200 px-3 py-2">
                            <div className="text-slate-500">Itens com venda</div>
                            <div className="font-semibold text-slate-900 mt-1">{registro.itens}</div>
                          </div>
                          <div className="rounded-xl bg-slate-50 border border-slate-200 px-3 py-2">
                            <div className="text-slate-500">Total de pares</div>
                            <div className="font-semibold text-slate-900 mt-1">{registro.totalPares}</div>
                          </div>
                        </div>

                        <div className="space-y-2">
                          {registro.resumo.length === 0 ? (
                            <div className="text-sm text-slate-500">Nenhum item com venda registrado.</div>
                          ) : (
                            registro.resumo.map((item, idx) => (
                              <div key={`${registro.id}-${item.ref}-${item.cor}-${idx}`} className="rounded-xl border border-slate-200 px-3 py-2 flex items-center justify-between gap-3 text-sm">
                                <div>
                                  <div className="font-semibold text-slate-900">{item.ref} • {item.cor}</div>
                                </div>
                                <div className="font-semibold text-slate-900">{item.total} pares</div>
                              </div>
                            ))
                          )}
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </PageShell>
  );

  const renderSugestoes = () => {
    const tempoTotal = (Number(tempoProducao?.pesponto) || 0) + (Number(tempoProducao?.montagem) || 0);
    const analise = rowsNormalized.map((row) => {
      const vendasRow = vendas[row.ref]?.[row.cor] || {};
      const totalVendido = sizes.reduce((acc, size) => acc + (Number(vendasRow[size]) || 0), 0);
      const totalPA = sizes.reduce((acc, size) => acc + (row.data[size]?.pa || 0), 0);
      const totalProd = sizes.reduce((acc, size) => acc + (row.data[size]?.est || 0) + (row.data[size]?.m || 0) + (row.data[size]?.p || 0), 0);
      const ruptura = sizes.some((size) => {
        const item = row.data[size];
        const minimo = minimos?.[row.ref]?.[row.cor]?.[size] || { pa: 0, prod: 0 };
        return item.pa < minimo.pa;
      });
      return {
        ref: row.ref,
        cor: row.cor,
        totalVendido,
        totalPA,
        totalProd,
        ruptura,
      };
    });

    const topVendas = [...analise]
      .filter((item) => item.totalVendido > 0)
      .sort((a, b) => b.totalVendido - a.totalVendido)
      .slice(0, 5);

    const alertaRuptura = [...analise]
      .filter((item) => item.ruptura)
      .sort((a, b) => b.totalVendido - a.totalVendido || a.totalPA - b.totalPA)
      .slice(0, 5);

    const encalhados = [...analise]
      .filter((item) => item.totalVendido === 0 && (item.totalPA > 0 || item.totalProd > 0))
      .sort((a, b) => b.totalPA + b.totalProd - (a.totalPA + a.totalProd))
      .slice(0, 5);

    const coberturaAnalise = rowsNormalized
      .flatMap((row) => sizes.map((size) => {
        const vendaMes = Number(vendas?.[row.ref]?.[row.cor]?.[size]) || 0;
        const cobertura = coberturaDias(row.data[size]?.pa || 0, vendaMes);
        return {
          ref: row.ref,
          cor: row.cor,
          size,
          pa: row.data[size]?.pa || 0,
          vendaMes,
          cobertura,
        };
      }))
      .filter((item) => item.cobertura !== null)
      .sort((a, b) => a.cobertura - b.cobertura);

    const coberturaCritica = coberturaAnalise.filter((item) => item.cobertura < tempoTotal);
    const menorCobertura = coberturaAnalise.slice(0, 5);

    return (
      <PageShell title="Sugestões" subtitle={`Planejamento automático com base em vendas, estoque atual, mínimos, produção em andamento e tempo de produção (${tempoProducao.pesponto}d pesponto + ${tempoProducao.montagem}d montagem).`}>
        <section className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
          <SummaryCard title="Sugestões de Montagem" value={suggestions.montagem.length} subtitle="Ordens sugeridas" />
          <SummaryCard title="Sugestões de Pesponto" value={suggestions.pesponto.length} subtitle="Ordens sugeridas" />
          <SummaryCard title="Produtos com ruptura" value={alertaRuptura.length} subtitle="PA abaixo do mínimo" />
          <SummaryCard title="Produtos encalhados" value={encalhados.length} subtitle="Sem venda no mês" />
          <SummaryCard title="Cobertura crítica" value={coberturaCritica.length} subtitle={`Abaixo de ${tempoTotal} dia(s)`} />
        </section>

        <section className="grid grid-cols-1 xl:grid-cols-4 gap-6">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-bold text-lg">TOP 5 do mês</h2>
              <span className="text-xs font-semibold px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 border border-emerald-200">Mais vendidos</span>
            </div>
            <div className="mt-4 space-y-3">
              {topVendas.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">Sem vendas lançadas ainda.</div>
              ) : (
                topVendas.map((item, idx) => (
                  <div key={`${item.ref}-${item.cor}-top`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex items-center justify-between gap-3">
                      <div>
                        <div className="font-semibold">#{idx + 1} {item.ref} • {item.cor}</div>
                        <div className="text-xs text-slate-500 mt-1">PA: {item.totalPA} • Produção: {item.totalProd}</div>
                      </div>
                      <div className="text-right">
                        <div className="text-2xl font-bold text-slate-900">{item.totalVendido}</div>
                        <div className="text-xs text-slate-500">vendidos</div>
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-bold text-lg">Alerta de ruptura</h2>
              <span className="text-xs font-semibold px-3 py-1 rounded-full bg-red-100 text-red-700 border border-red-200">Ação rápida</span>
            </div>
            <div className="mt-4 space-y-3">
              {alertaRuptura.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">Nenhum item com ruptura agora.</div>
              ) : (
                alertaRuptura.map((item) => (
                  <div key={`${item.ref}-${item.cor}-ruptura`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="font-semibold">{item.ref} • {item.cor}</div>
                    <div className="mt-2 flex items-center justify-between text-sm text-slate-600">
                      <span>Vendas do mês</span>
                      <span className="font-semibold">{item.totalVendido}</span>
                    </div>
                    <div className="mt-1 flex items-center justify-between text-sm text-slate-600">
                      <span>PA atual</span>
                      <span className="font-semibold">{item.totalPA}</span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-bold text-lg">Produtos parados</h2>
              <span className="text-xs font-semibold px-3 py-1 rounded-full bg-amber-100 text-amber-700 border border-amber-200">Sem giro</span>
            </div>
            <div className="mt-4 space-y-3">
              {encalhados.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">Nenhum encalhado identificado.</div>
              ) : (
                encalhados.map((item) => (
                  <div key={`${item.ref}-${item.cor}-encalhado`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="font-semibold">{item.ref} • {item.cor}</div>
                    <div className="mt-2 flex items-center justify-between text-sm text-slate-600">
                      <span>PA</span>
                      <span className="font-semibold">{item.totalPA}</span>
                    </div>
                    <div className="mt-1 flex items-center justify-between text-sm text-slate-600">
                      <span>Produção</span>
                      <span className="font-semibold">{item.totalProd}</span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
                  <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-bold text-lg">Menor cobertura</h2>
              <span className="text-xs font-semibold px-3 py-1 rounded-full bg-sky-100 text-[#8B1E2D] border border-sky-200">Dias de estoque</span>
            </div>
            <div className="mt-4 space-y-3">
              {menorCobertura.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">Sem vendas suficientes para cobertura.</div>
              ) : (
                menorCobertura.map((item) => (
                  <div key={`${item.ref}-${item.cor}-${item.size}-cob`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex items-center justify-between gap-3">
                      <div>
                        <div className="font-semibold">{item.ref} • {item.cor}</div>
                        <div className="text-xs text-slate-500 mt-1">Numeração {item.size}</div>
                      </div>
                      <span className={`px-3 py-1 rounded-full border text-xs font-semibold ${coberturaBadgeClass(item.cobertura, tempoTotal)}`}>
                        {item.cobertura.toFixed(1)} dia(s)
                      </span>
                    </div>
                    <div className="mt-2 flex items-center justify-between text-sm text-slate-600">
                      <span>PA</span>
                      <span className="font-semibold">{item.pa}</span>
                    </div>
                    <div className="mt-1 flex items-center justify-between text-sm text-slate-600">
                      <span>Venda diária</span>
                      <span className="font-semibold">{vendaDiaFromMes(item.vendaMes).toFixed(1)}</span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </section>

        <section className="grid grid-cols-1 xl:grid-cols-2 gap-6">
          {[
            { title: "Montagem", list: suggestions.montagem },
            { title: "Pesponto", list: suggestions.pesponto },
          ].map((block) => (
            <div key={block.title} className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
              <div className="flex items-center justify-between">
                <h2 className="font-bold text-lg">{block.title}</h2>
                <span className="text-sm text-slate-500">{block.list.length} sugestões</span>
              </div>
              <div className="mt-4 space-y-4">
                {block.list.length === 0 ? (
                  <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500">Nenhuma sugestão no momento.</div>
                ) : (
                  block.list.map((s, idx) => (
                    <div key={`${block.title}-${s.ref}-${s.cor}`} className="rounded-2xl bg-slate-50 border border-slate-200 p-4">
                      <div className="flex items-center justify-between gap-3">
                        <div>
                          <div className="font-semibold">{s.ref} • {s.cor}</div>
                          <div className="text-sm text-slate-500 mt-1">Total sugerido: {s.total} pares</div>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-1 rounded-full text-[11px] font-semibold border ${idx === 0 ? "bg-red-100 text-red-700 border-red-200" : idx < 3 ? "bg-amber-100 text-amber-700 border-amber-200" : "bg-emerald-100 text-emerald-700 border-emerald-200"}`}>
                            {idx === 0 ? "Prioridade máxima" : idx < 3 ? "Alta prioridade" : "Prioridade normal"}
                          </span>
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          ))}
        </section>
      </PageShell>
    );
  };

  const renderFichas = () => (
    <PageShell title="Gerador de Fichas" subtitle="Estrutura baseada nas abas GERADOR_PESPONTO e GERADOR_MONTAGEM da planilha.">
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
        {[
          {
            key: "montagem",
            title: "GERADOR_MONTAGEM",
            subtitle: "Fichas geradas a partir das sugestões de montagem.",
            list: fichasMontagem,
          },
          {
            key: "pesponto",
            title: "GERADOR_PESPONTO",
            subtitle: "Fichas geradas a partir das sugestões de pesponto.",
            list: fichasPesponto,
          },
        ].map((block) => {
          const grupos = Object.entries(
            block.list.reduce((acc, ficha) => {
              const key = `${ficha.ref}__${ficha.cor}`;
              if (!acc[key]) acc[key] = [];
              acc[key].push(ficha);
              return acc;
            }, {})
          );

          return (
            <div key={block.key} className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-hidden">
              <div className="px-6 py-5 border-b border-slate-200 bg-slate-50">
                <div className="flex items-center justify-between gap-4">
                  <div>
                    <div className="text-xs uppercase tracking-[0.22em] text-slate-500">Planilha</div>
                    <h2 className="text-xl font-bold mt-1">{block.title}</h2>
                    <p className="text-sm text-slate-500 mt-1">{block.subtitle}</p>
                  </div>
                  <div className="rounded-2xl bg-white border border-slate-200 px-4 py-3 text-right min-w-[120px]">
                    <div className="text-xs text-slate-500">Fichas</div>
                    <div className="text-2xl font-bold text-slate-900">{block.list.length}</div>
                  </div>
                </div>
              </div>

              <div className="p-6 space-y-4">
                {block.list.length === 0 ? (
                  <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500">
                    Nenhuma ficha gerada no momento.
                  </div>
                ) : (
                  grupos.map(([grupoKey, fichasGrupo]) => {
                    const { cor } = fichasGrupo[0];
                    const aberto = Boolean(fichasAbertas[`${block.key}__${grupoKey}`]);
                    return (
                      <div key={grupoKey} className="rounded-2xl border border-slate-200 overflow-hidden">
                        <button
                          type="button"
                          onClick={() => setFichasAbertas((curr) => ({
                            ...curr,
                            [`${block.key}__${grupoKey}`]: !curr[`${block.key}__${grupoKey}`],
                          }))}
                          className="w-full bg-slate-50 px-5 py-4 flex items-center justify-between gap-4 text-left"
                        >
                          <div>
                            <div className="text-lg font-bold text-slate-900">{cor}</div>
                            <div className="text-sm text-slate-500 mt-1">{fichasGrupo.length} ficha{fichasGrupo.length > 1 ? "s" : ""}</div>
                          </div>
                          <div className="px-3 py-1.5 rounded-xl border border-slate-200 bg-white text-sm font-semibold text-slate-700">
                            {aberto ? "Fechar" : "Abrir"}
                          </div>
                        </button>

                        {aberto && (
                          <div className="p-4 space-y-3 bg-white border-t border-slate-200">
                            {fichasGrupo.map((ficha, idx) => (
                              <button
                                key={`${block.key}-${ficha.nome}`}
                                type="button"
                                onClick={() => setPreviewFicha(ficha)}
                                className="w-full rounded-2xl border border-slate-200 px-4 py-4 flex items-center justify-between gap-4 text-left hover:bg-slate-50"
                              >
                                <div>
                                  <div className="font-semibold text-slate-900">Ficha {String(idx + 1).padStart(2, "0")}</div>
                                  <div className="text-sm text-slate-500 mt-1">{ficha.ref} • {ficha.cor}</div>
                                </div>
                                <div className="flex items-center gap-3">
                                  <div className="px-3 py-1.5 rounded-xl bg-slate-950 text-white text-sm font-semibold">{ficha.total} pares</div>
                                  <div className="text-sm font-semibold text-blue-700">Visualizar</div>
                                </div>
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          );
        })}
      </div>
    </PageShell>
  );

  const renderProgramacaoDia = () => {
    const quickDays = [1, 3, 7, 15];
    const capPProg = Number(capacidadePespontoDia) || 396;
    const capMProg = Number(capacidadeMontagemDia) || 396;
    const fichasIncompativeisPesponto = fichasPesponto.filter((ficha) => Number(ficha.total) > capPProg);
    const fichasIncompativeisMontagem = fichasMontagem.filter((ficha) => Number(ficha.total) > capMProg);
    const fichasIncompativeisAtivas =
      programacaoSubAba === "Pesponto" ? fichasIncompativeisPesponto : fichasIncompativeisMontagem;
    const pctReservaTop = Math.min(100, Math.max(0, Number(programacaoReservaTopPct) || 0));
    const reservaParesPespontoDia = Math.round((capPProg * pctReservaTop) / 100);
    const reservaParesMontagemDia = Math.round((capMProg * pctReservaTop) / 100);

    const planoPespontoUI = programacaoPlanoCongelado?.pesponto ?? programacaoPesponto;
    const planoMontagemUI = programacaoPlanoCongelado?.montagem ?? programacaoMontagem;
    const diasPlanoExibicao = programacaoPlanoCongelado?.dias ?? programacaoDias;

    const subAbas = [
      {
        key: "Pesponto",
        titulo: "Programação de Pesponto",
        descricao: "Veja apenas a programação do pesponto no período selecionado.",
        programacao: planoPespontoUI,
        corTag: "bg-amber-100 text-amber-700 border-amber-200",
      },
      {
        key: "Montagem",
        titulo: "Programação de Montagem",
        descricao: "Veja apenas a programação da montagem no período selecionado.",
        programacao: planoMontagemUI,
        corTag: "bg-sky-100 text-[#8B1E2D] border-sky-200",
      },
    ];

    const subAbaAtiva = subAbas.find((item) => item.key === programacaoSubAba) || subAbas[0];

    const abrirLancarFichaProgramacao = (ficha, tipoMov, diaNumero, fichaStorageKey) => {
      const grid = makeEmptyGrid();
      sizes.forEach((s) => {
        grid[s] = Number(ficha.sizes?.[s]) || 0;
      });
      const baseNome = String(ficha.nome || "").trim();
      let programacaoNome = baseNome;
      if (diaNumero != null && diaNumero !== undefined && !Number.isNaN(Number(diaNumero))) {
        programacaoNome = baseNome ? `${baseNome} · Dia ${diaNumero}` : `Dia ${diaNumero} · ${ficha.ref} · ${ficha.cor}`;
      }
      if (!String(programacaoNome || "").trim()) {
        programacaoNome = `${ficha.ref} · ${ficha.cor}`;
      }
      setMovError((curr) => ({ ...curr, [tipoMov]: "" }));
      setLancarFichaDaProgramacao({
        tipo: tipoMov,
        ref: ficha.ref,
        cor: ficha.cor,
        programacao: programacaoNome,
        grid,
        fichaStorageKey: fichaStorageKey || null,
      });
    };

    const isProgFichaLancada = (key) => fichasProgramacaoLancadas.includes(key);

    const todasFichas = subAbaAtiva.programacao.diasProgramados.flatMap((dia) =>
      dia.fichas.map((ficha) => ({
        ...ficha,
        dia: dia.dia,
      }))
    );

    const isFichaSel = (key) => programacaoFichaSelecao[key] !== false;
    const toggleFichaSel = (key) =>
      setProgramacaoFichaSelecao((p) => ({ ...p, [key]: !isFichaSel(key) }));

    const allKeysProg = [];
    if (programacaoModoVisual === "normal") {
      subAbaAtiva.programacao.diasProgramados.forEach((dia) => {
        dia.fichas.forEach((ficha, idx) => {
          allKeysProg.push(`n-${programacaoSubAba}-${dia.dia}-${idx}-${ficha.nome}`);
        });
      });
    } else {
      todasFichas.forEach((ficha, idx) => {
        allKeysProg.push(`c-${programacaoSubAba}-${ficha.dia}-${idx}-${ficha.nome}`);
      });
    }

    const marcarTodasProg = () => {
      setProgramacaoFichaSelecao(Object.fromEntries(allKeysProg.map((k) => [k, true])));
    };
    const desmarcarTodasProg = () => {
      setProgramacaoFichaSelecao(Object.fromEntries(allKeysProg.map((k) => [k, false])));
    };

    const itensImpressao = [];
    const nomeProgramacaoLote = String(programacaoNomeLoteImpressao || "").trim() || String(programacaoEtiquetaFicha || "").trim();
    if (programacaoModoVisual === "normal") {
      subAbaAtiva.programacao.diasProgramados.forEach((dia) => {
        dia.fichas.forEach((ficha, idx) => {
          const key = `n-${programacaoSubAba}-${dia.dia}-${idx}-${ficha.nome}`;
          if (isFichaSel(key)) itensImpressao.push({ ficha, dia: dia.dia, ordem: idx + 1, programacaoNome: nomeProgramacaoLote });
        });
      });
    } else {
      todasFichas.forEach((ficha, idx) => {
        const key = `c-${programacaoSubAba}-${ficha.dia}-${idx}-${ficha.nome}`;
        if (isFichaSel(key)) itensImpressao.push({ ficha, dia: ficha.dia, ordem: idx + 1, programacaoNome: nomeProgramacaoLote });
      });
    }
    itensImpressao.sort((a, b) => a.dia - b.dia || a.ordem - b.ordem);
    const copiasPorPaginaEfetivo = Math.max(1, Math.min(4, Math.round(Number(programacaoCopiasPorPagina) || 1)));
    const buildItensComCopias = (copias) => {
      const lista = [];
      for (let copia = 1; copia <= copias; copia += 1) {
        itensImpressao.forEach((item) => {
          lista.push({ ...item, copia });
        });
      }
      return lista;
    };
    const itensImpressaoFolha1 = buildItensComCopias(3);
    const itensImpressaoFolha2 = buildItensComCopias(copiasPorPaginaEfetivo);
    const tituloFolhaImpressao = (String(programacaoEtiquetaFicha || "").trim() || "Programação do Dia");
    const itensImpressaoComCopias =
      programacaoTipoFolha === "folha1"
        ? itensImpressaoFolha1
        : programacaoTipoFolha === "folha2"
          ? itensImpressaoFolha2
          : [...itensImpressaoFolha1, ...itensImpressaoFolha2];
    const copiasResumo =
      programacaoTipoFolha === "folha1"
        ? "3 (Folha 1)"
        : programacaoTipoFolha === "folha2"
          ? `${copiasPorPaginaEfetivo} (Folha 2)`
          : `3 (Folha 1) + ${copiasPorPaginaEfetivo} (Folha 2)`;
    const blocosPorPaginaEfetivo = programacaoTipoFolha === "folha1" ? 3 : copiasPorPaginaEfetivo;
    const paginasEstimadas =
      programacaoTipoFolha === "ambas"
        ? ((itensImpressaoFolha1.length ? Math.ceil(itensImpressaoFolha1.length / 3) : 0) +
          (itensImpressaoFolha2.length ? Math.ceil(itensImpressaoFolha2.length / copiasPorPaginaEfetivo) : 0))
        : (itensImpressaoComCopias.length
          ? Math.ceil(itensImpressaoComCopias.length / blocosPorPaginaEfetivo)
          : 0);

    const imprimirProgramacaoDia = () => {
      if (itensImpressao.length === 0) {
        alert("Selecione pelo menos uma ficha para imprimir.");
        return;
      }
      if (itensImpressao.length > 1 && !nomeProgramacaoLote) {
        alert("Informe o nome da programação para imprimir várias fichas como uma programação única.");
        return;
      }
      window.print();
    };

    const gerarPdfWhatsappProgramacao = async () => {
      if (itensImpressao.length === 0) {
        alert("Selecione pelo menos uma ficha para gerar o PDF.");
        return;
      }
      if (itensImpressao.length > 1 && !nomeProgramacaoLote) {
        alert("Informe o nome da programação para gerar o PDF em lote.");
        return;
      }
      const el = programacaoPrintSheetRef.current;
      if (!el) {
        alert("Não foi possível localizar a folha de impressão. Abra a Programação do Dia e tente de novo.");
        return;
      }
      const root = document.documentElement;
      setProgramacaoPdfBusy(true);
      root.classList.add("programacao-pdf-capture");
      el.classList.remove("hidden");
      el.classList.add("programacao-print-sheet--capture");
      try {
        try {
          if (document.fonts?.ready) await document.fonts.ready;
        } catch (_) {
          /* ignore */
        }
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
        const blob = await buildProgramacaoPrintSheetPdfBlobFromElement(el);
        const fname = `programacao-${programacaoSubAba}-${new Date().toISOString().slice(0, 10)}.pdf`;
        const file = new File([blob], fname, { type: "application/pdf" });
        const titulo = tituloFolhaImpressao;
        const texto = `${titulo} · ${programacaoSubAba}`;
        const payload = { files: [file], title: titulo, text: texto };
        if (typeof navigator !== "undefined" && navigator.canShare && navigator.canShare(payload)) {
          await navigator.share(payload);
        } else {
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = fname;
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(url);
          alert("PDF descarregado. Abra o WhatsApp e anexe este ficheiro (📎 Documento).");
        }
      } catch (e) {
        if (e && e.name === "AbortError") {
          /* cancelado */
        } else {
          console.error(e);
          alert("Não foi possível gerar o PDF. Tente novamente ou use Imprimir seleção.");
        }
      } finally {
        el.classList.remove("programacao-print-sheet--capture");
        el.classList.add("hidden");
        root.classList.remove("programacao-pdf-capture");
        setProgramacaoPdfBusy(false);
      }
    };

    const renderBlocoProgramacao = (programacao, corTag, fichasIncompativeis = []) => (
      <section className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <SummaryCard title={`Capacidade ${programacao.tipo}`} value={programacao.capacidadeDia * programacao.dias} subtitle={`${programacao.dias} dia(s) planejados`} />
          <SummaryCard title="Programado" value={programacao.totalProgramado} subtitle="Pares encaixados no período" />
          <SummaryCard title="Saldo" value={programacao.totalRestante} subtitle="Capacidade ainda livre" />
          <SummaryCard title="Fichas" value={programacao.totalFichas} subtitle="Fichas distribuídas no período" />
        </div>

        <div className="grid grid-cols-1 xl:grid-cols-[1fr_360px] gap-6 items-start">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-4 mb-4">
              <div>
                <h2 className="font-bold text-lg">Programação de {programacao.tipo}</h2>
                <p className="text-sm text-slate-500 mt-1">Período separado por setor para ajudar no planejamento de materiais e execução.</p>
                {programacao.reservaTopAtiva ? (
                  <p className="text-xs text-amber-800 mt-2 rounded-xl border border-amber-200 bg-amber-50 px-3 py-2">
                    Reserva top vendas: <span className="font-semibold">{programacao.reservaTopVendasPct}%</span> do dia para{" "}
                    {programacao.reservaTopModo === "manual" ? (
                      <>
                        as <span className="font-semibold">{programacao.reservaTopManualSelecionados}</span> ref+cor selecionadas manualmente
                      </>
                    ) : (
                      <>
                        as <span className="font-semibold">{programacao.reservaTopN}</span> ref+cor com maior venda
                      </>
                    )}
                    ; o restante do dia segue o rodízio.
                  </p>
                ) : null}
                {fichasIncompativeis.length > 0 ? (
                  <p className="text-xs text-rose-800 mt-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2">
                    Há <span className="font-semibold">{fichasIncompativeis.length}</span> ficha(s) acima da capacidade diária de{" "}
                    <span className="font-semibold">{programacao.capacidadeDia}</span> pares e elas não podem ser alocadas neste plano.
                  </p>
                ) : null}
              </div>
              <span className={`px-3 py-1.5 rounded-xl border text-sm font-semibold ${corTag}`}>{programacao.tipo}</span>
            </div>

            {programacao.diasProgramados.every((dia) => dia.fichas.length === 0) ? (
              <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500">
                Nenhuma ficha entrou na programação de {programacao.tipo} para este período.
              </div>
            ) : (
              <div className="space-y-5">
                {programacao.diasProgramados.map((dia) => (
                  <div key={`${programacao.tipo}-dia-${dia.dia}`} className="rounded-2xl border border-slate-200 overflow-hidden">
                    <div className="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between gap-4">
                      <div>
                        <div className="font-semibold">Dia {String(dia.dia).padStart(2, "0")}</div>
                        <div className="text-sm text-slate-500 mt-1">{dia.totalProgramado} programados • {dia.restante} livres</div>
                      </div>
                      <div className="text-sm text-slate-500">{dia.fichas.length} ficha(s)</div>
                    </div>

                    <div className="p-4 space-y-3">
                      {dia.fichas.length === 0 ? (
                        <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                          Sem fichas para este dia.
                        </div>
                      ) : (
                        dia.fichas.map((ficha, idx) => {
                          const fKey = `n-${programacaoSubAba}-${dia.dia}-${idx}-${ficha.nome}`;
                          return (
                            <div
                              key={`${programacao.tipo}-${dia.dia}-${ficha.nome}-${idx}`}
                              className="rounded-2xl border border-slate-200 p-4 flex items-stretch justify-between gap-3"
                            >
                              <label className="print:hidden flex min-h-[44px] min-w-[44px] shrink-0 cursor-pointer items-center justify-center sm:min-h-0 sm:min-w-0 sm:items-start sm:justify-start sm:pt-1">
                                <input
                                  type="checkbox"
                                  className="h-6 w-6 rounded border-slate-300 text-[#8B1E2D] focus:ring-2 focus:ring-[#8B1E2D] focus:ring-offset-0 sm:h-5 sm:w-5"
                                  checked={isFichaSel(fKey)}
                                  onChange={() => toggleFichaSel(fKey)}
                                />
                              </label>
                              <div className="flex-1 min-w-0 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                                <div>
                                  <div className="text-xs uppercase tracking-wide text-slate-500">Ordem {String(idx + 1).padStart(2, "0")}</div>
                                  <div className="font-semibold text-slate-900 mt-1">{ficha.cor}</div>
                                  <div className="text-sm text-slate-500 mt-1">{ficha.ref} • {ficha.nome}</div>
                                  {ficha.programacaoReservaTop ? (
                                    <span className="inline-block mt-2 px-2 py-0.5 rounded-lg text-[10px] font-semibold bg-amber-100 text-amber-900 border border-amber-200">
                                      Reserva top vendas
                                    </span>
                                  ) : null}
                                </div>
                                <div className="flex flex-wrap items-center gap-3 shrink-0 justify-end">
                                  <div className="text-right">
                                    <div className="text-xs text-slate-500">Prioridade</div>
                                    <div className="font-semibold text-slate-900">{Math.round(ficha.prioridade || 0)}</div>
                                  </div>
                                  <div className="px-3 py-1.5 rounded-xl bg-slate-950 text-white text-sm font-semibold">{ficha.total} pares</div>
                                  {isProgFichaLancada(fKey) ? (
                                    <span className="print:hidden px-3 py-1.5 text-xs font-semibold rounded-xl bg-emerald-100 text-emerald-800 border border-emerald-200">
                                      Lançada
                                    </span>
                                  ) : null}
                                  <button type="button" onClick={() => setPreviewFicha(ficha)} className="print:hidden px-3 py-1.5 text-xs font-semibold rounded-xl bg-[#FCECEE] text-[#8B1E2D] border border-[#E7C7CC] hover:bg-[#F7DDE1]">Visualizar</button>
                                  <button
                                    type="button"
                                    title={`Lançar esta ficha em ${programacao.tipo}`}
                                    onClick={() => abrirLancarFichaProgramacao(ficha, programacao.tipo, dia.dia, fKey)}
                                    className={`print:hidden px-3 py-1.5 text-xs font-semibold rounded-xl border ${
                                      isProgFichaLancada(fKey)
                                        ? "bg-slate-100 text-slate-500 border-slate-200 cursor-not-allowed"
                                        : "bg-[#0F172A] text-white border-[#0F172A] hover:bg-slate-800"
                                    }`}
                                    disabled={isProgFichaLancada(fKey)}
                                  >
                                    Lançar
                                  </button>
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <h2 className="font-bold text-lg">Resumo por cor / modelo</h2>
            <div className="mt-4 space-y-3">
              {Object.values(programacao.diasProgramados.flatMap((dia) => dia.fichas).reduce((acc, ficha) => {
                const key = `${ficha.ref}__${ficha.cor}`;
                if (!acc[key]) acc[key] = { ref: ficha.ref, cor: ficha.cor, fichas: 0, pares: 0 };
                acc[key].fichas += 1;
                acc[key].pares += ficha.total;
                return acc;
              }, {})).length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
                  Sem programação montada ainda.
                </div>
              ) : (
                Object.values(programacao.diasProgramados.flatMap((dia) => dia.fichas).reduce((acc, ficha) => {
                  const key = `${ficha.ref}__${ficha.cor}`;
                  if (!acc[key]) acc[key] = { ref: ficha.ref, cor: ficha.cor, fichas: 0, pares: 0 };
                  acc[key].fichas += 1;
                  acc[key].pares += ficha.total;
                  return acc;
                }, {}))
                  .sort((a, b) => b.pares - a.pares)
                  .map((item) => (
                    <div key={`${programacao.tipo}-${item.ref}-${item.cor}`} className="rounded-2xl border border-slate-200 px-4 py-3">
                      <div className="font-semibold text-slate-900">{item.cor}</div>
                      <div className="text-sm text-slate-500 mt-1">{item.ref}</div>
                      <div className="mt-2 flex items-center justify-between text-sm text-slate-600">
                        <span>Fichas</span>
                        <span className="font-semibold">{item.fichas}</span>
                      </div>
                      <div className="mt-1 flex items-center justify-between text-sm text-slate-600">
                        <span>Pares</span>
                        <span className="font-semibold">{item.pares}</span>
                      </div>
                    </div>
                  ))
              )}
            </div>
          </div>
        </div>
      </section>
    );

    const renderModoCompleto = () => (
  <section className="space-y-6">
    <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
      <SummaryCard
        title={`Capacidade ${subAbaAtiva.programacao.tipo}`}
        value={subAbaAtiva.programacao.capacidadeDia * subAbaAtiva.programacao.dias}
        subtitle={`${subAbaAtiva.programacao.dias} dia(s) planejados`}
      />
      <SummaryCard
        title="Programado"
        value={subAbaAtiva.programacao.totalProgramado}
        subtitle="Pares encaixados no período"
      />
      <SummaryCard
        title="Saldo"
        value={subAbaAtiva.programacao.totalRestante}
        subtitle="Capacidade ainda livre"
      />
      <SummaryCard
        title="Fichas"
        value={subAbaAtiva.programacao.totalFichas}
        subtitle="Fichas abertas na tela"
      />
    </div>

    {todasFichas.length === 0 ? (
      <div className="bg-white rounded-[28px] border border-dashed border-slate-300 shadow-sm p-10 text-center text-sm text-slate-500">
        Nenhuma ficha disponível para exibição completa.
      </div>
    ) : (
      <div className="space-y-5">
        {todasFichas.map((ficha, idx) => {
          const fKey = `c-${programacaoSubAba}-${ficha.dia}-${idx}-${ficha.nome}`;
          return (
          <div
            key={`${ficha.nome}-${idx}`}
            className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6 break-inside-avoid"
          >
            <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 mb-5">
              <label className="print:hidden flex cursor-pointer items-start gap-3 md:items-center">
                <span className="inline-flex min-h-[44px] min-w-[44px] shrink-0 items-center justify-center sm:min-h-0 sm:min-w-0 sm:inline-flex sm:mt-1 md:mt-0">
                  <input
                    type="checkbox"
                    className="h-6 w-6 rounded border-slate-300 text-[#8B1E2D] focus:ring-2 focus:ring-[#8B1E2D] focus:ring-offset-0 sm:h-5 sm:w-5"
                    checked={isFichaSel(fKey)}
                    onChange={() => toggleFichaSel(fKey)}
                  />
                </span>
                <div>
                <div className="text-xs uppercase tracking-[0.18em] text-slate-400 font-semibold">
                  {programacaoSubAba} • Dia {String(ficha.dia).padStart(2, "0")}
                </div>
                <div className="text-xl font-black text-slate-900 mt-1">{ficha.nome}</div>
                <div className="text-sm text-slate-500 mt-1">
                  <span className="font-semibold text-slate-800">{ficha.ref}</span> • {ficha.cor}
                </div>
                {ficha.programacaoReservaTop ? (
                  <span className="inline-block mt-2 px-2 py-0.5 rounded-lg text-[10px] font-semibold bg-amber-100 text-amber-900 border border-amber-200">
                    Reserva top vendas
                  </span>
                ) : null}
                </div>
              </label>

              <div className="flex flex-wrap items-center gap-3 print:ml-auto">
                <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-right min-w-[120px]">
                  <div className="text-xs text-slate-500">Total da ficha</div>
                  <div className="text-2xl font-black text-slate-900">{ficha.total}</div>
                </div>

                <button
                  type="button"
                  onClick={() => setPreviewFicha(ficha)}
                  className="print:hidden px-4 py-2 rounded-2xl text-sm font-semibold border bg-[#FCECEE] text-[#8B1E2D] border-[#E7C7CC] hover:bg-[#F7DDE1]"
                >
                  Visualizar
                </button>
                {isProgFichaLancada(fKey) ? (
                  <span className="print:hidden px-4 py-2 rounded-2xl text-sm font-semibold border bg-emerald-100 text-emerald-800 border-emerald-200">
                    Lançada
                  </span>
                ) : null}
                <button
                  type="button"
                  title={`Lançar esta ficha em ${subAbaAtiva.programacao.tipo}`}
                  onClick={() => abrirLancarFichaProgramacao(ficha, subAbaAtiva.programacao.tipo, ficha.dia, fKey)}
                  className={`print:hidden px-4 py-2 rounded-2xl text-sm font-semibold border ${
                    isProgFichaLancada(fKey)
                      ? "bg-slate-100 text-slate-500 border-slate-200 cursor-not-allowed"
                      : "bg-[#0F172A] text-white border-[#0F172A] hover:bg-slate-800"
                  }`}
                  disabled={isProgFichaLancada(fKey)}
                >
                  Lançar
                </button>
              </div>
            </div>

            <div className="overflow-auto">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-slate-50">
                    {sizes.map((size) => (
                      <th
                        key={`head-${ficha.nome}-${size}`}
                        className="border border-slate-200 px-3 py-2 text-center font-bold text-slate-700"
                      >
                        {size}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    {sizes.map((size) => (
                      <td
                        key={`body-${ficha.nome}-${size}`}
                        className="border border-slate-200 px-3 py-3 text-center font-semibold text-slate-900"
                      >
                        {Number(ficha.sizes?.[size] || 0)}
                      </td>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>

            <div className="mt-4 flex flex-wrap gap-2">
              {sizes
                .filter((size) => Number(ficha.sizes?.[size] || 0) > 0)
                .map((size) => (
                  <span
                    key={`tag-${ficha.nome}-${size}`}
                    className="px-3 py-1 rounded-full bg-slate-100 text-slate-700 text-xs font-semibold"
                  >
                    {size}: {ficha.sizes?.[size]}
                  </span>
                ))}
            </div>
          </div>
        );
        })}
      </div>
    )}
  </section>
);

    return (
      <PageShell title="Programação do Dia" subtitle="Defina quantos dias quer programar. O sistema monta períodos separados de Montagem e Pesponto para antecipar matérias-primas e execução.">
        <div id="print-programacao-root" className="space-y-6">
          <div className="print:hidden space-y-6">
          <div className="rounded-2xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-950 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
            <p className="min-w-0">
              <span className="font-semibold">Plano congelado.</span>{" "}
              Lançamentos no Pesponto/Montagem não alteram fichas e ordem desta página até você clicar em{" "}
              <span className="font-semibold">Recalcular plano</span>. Mudar dias, capacidade ou reserva top vendas atualiza o plano automaticamente.
            </p>
            <button
              type="button"
              disabled={fichasProgramacaoLancadas.length === 0}
              onClick={() => {
                if (fichasProgramacaoLancadas.length === 0) return;
                if (window.confirm("Limpar todas as marcas de fichas já lançadas (memória deste navegador)?")) {
                  limparMarcasFichasLancadasProgramacao();
                }
              }}
              className="print:hidden shrink-0 rounded-xl border border-emerald-300 bg-white px-3 py-2 text-xs font-semibold text-emerald-900 hover:bg-emerald-100 disabled:opacity-40 disabled:cursor-not-allowed"
            >
              Limpar marcas de lançadas
            </button>
          </div>
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-4 md:p-5">
            <div className="font-bold text-lg text-slate-900">Impressão (A4)</div>
            <p className="text-xs text-slate-500 mt-1">Configuração rápida da folha de impressão.</p>
            <div className="mt-3 grid grid-cols-1 lg:grid-cols-3 gap-3">
              <label className="text-sm font-medium text-slate-700 block">
                Caminho da logo (ficheiro em <code className="text-xs bg-slate-100 px-1 rounded">public/</code>)
                <input
                  type="text"
                  value={programacaoLogoImpressao}
                  onChange={(e) => setProgramacaoLogoImpressao(e.target.value)}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm"
                  placeholder="/logo-rockstar-bandeira.png"
                />
              </label>
              <label className="text-sm font-medium text-slate-700 block">
                Tipo de impressão
                <select
                  value={programacaoTipoFolha}
                  onChange={(e) => {
                    const v = e.target.value;
                    setProgramacaoTipoFolha(v === "folha2" || v === "ambas" ? v : "folha1");
                  }}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm"
                >
                  <option value="folha1">Folha 1 - Terceirizados</option>
                  <option value="folha2">Folha 2 - Pesponto</option>
                  <option value="ambas">Ambas - Folha 1 + Folha 2</option>
                </select>
              </label>
              <label className="text-sm font-medium text-slate-700 block lg:col-span-3">
                Observações na folha
                <textarea
                  value={programacaoObsImpressao}
                  onChange={(e) => setProgramacaoObsImpressao(e.target.value)}
                  rows={1}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm resize-y"
                  placeholder="Notas para a equipe (aparecem na impressão)..."
                />
              </label>
              <label className="text-sm font-medium text-slate-700 block lg:col-span-3">
                Nome da programação (lote)
                <input
                  type="text"
                  value={programacaoNomeLoteImpressao}
                  onChange={(e) => setProgramacaoNomeLoteImpressao(e.target.value)}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm"
                  placeholder="Ex.: PROG 29/04 - LOTE A (obrigatório ao imprimir mais de 1 ficha)"
                />
              </label>
              <label className="text-sm font-medium text-slate-700 block">
                Cópias por página
                <select
                  value={programacaoCopiasPorPagina}
                  onChange={(e) => setProgramacaoCopiasPorPagina(Math.min(4, Math.max(1, Number(e.target.value) || 1)))}
                  disabled={programacaoTipoFolha === "folha1"}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm"
                >
                  <option value={1}>1 cópia por página</option>
                  <option value={2}>2 cópias por página</option>
                  <option value={3}>3 cópias por página</option>
                  <option value={4}>4 cópias por página</option>
                </select>
                {programacaoTipoFolha === "folha1" ? (
                  <p className="mt-1 text-[11px] text-slate-500">Folha 1 usa 3 cópias fixas.</p>
                ) : null}
              </label>
              <label className="text-sm font-medium text-slate-700 block">
                Cabeçalho da folha (impressão)
                <select
                  value={programacaoCabecalhoFolha}
                  onChange={(e) => {
                    const v = e.target.value;
                    setProgramacaoCabecalhoFolha(v === "minimo" || v === "oculto" ? v : "completo");
                  }}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm"
                >
                  <option value="completo">Completo (logo, título e resumo)</option>
                  <option value="minimo">Mínimo (uma linha)</option>
                  <option value="oculto">Oculto (só as fichas — mais espaço no A4)</option>
                </select>
              </label>
              <label className="text-sm font-medium text-slate-700 block lg:col-span-3">
                Texto no cabeçalho de cada ficha (opcional)
                <input
                  type="text"
                  value={programacaoEtiquetaFicha}
                  onChange={(e) => setProgramacaoEtiquetaFicha(e.target.value)}
                  className="mt-1.5 w-full rounded-xl border border-slate-200 bg-slate-50 p-2.5 text-sm"
                  placeholder='Ex.: FICHA 26-102 — CANO ALTO PRETO (texto exibido tal como está; vazio = padrão do sistema)'
                />
              </label>
              <div className={`lg:col-span-3 rounded-2xl border border-slate-200 bg-slate-50/70 p-3 ${programacaoTipoFolha === "folha2" ? "opacity-60" : ""}`}>
                <div className="text-sm font-semibold text-slate-800">Valor por par (Folha 1 · Terceirizados)</div>
                <p className="mt-0.5 text-[11px] text-slate-500">
                  Weverton e Romulo usam valor fixo. Gu segue regra fixa por referencia: BTCV010/TNCV010 = R$ 0,40 e CRVTNCV = R$ 0,30.
                </p>
                <div className="mt-3 grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <label className="text-xs font-medium text-slate-700">
                    Weverton (fixo)
                    <input
                      type="text"
                      inputMode="decimal"
                      value={programacaoValoresTerceiros.weverton}
                      disabled={programacaoTipoFolha === "folha2"}
                      onChange={(e) =>
                        setProgramacaoValoresTerceiros((curr) => ({ ...curr, weverton: e.target.value }))
                      }
                      className="mt-1 w-full rounded-xl border border-slate-200 bg-white p-2 text-sm"
                      placeholder="0,00"
                    />
                  </label>
                  <label className="text-xs font-medium text-slate-700">
                    Romulo (fixo)
                    <input
                      type="text"
                      inputMode="decimal"
                      value={programacaoValoresTerceiros.romulo}
                      disabled={programacaoTipoFolha === "folha2"}
                      onChange={(e) =>
                        setProgramacaoValoresTerceiros((curr) => ({ ...curr, romulo: e.target.value }))
                      }
                      className="mt-1 w-full rounded-xl border border-slate-200 bg-white p-2 text-sm"
                      placeholder="0,00"
                    />
                  </label>
                </div>
              </div>
            </div>
            <div className="mt-3 flex flex-wrap items-center gap-2">
              <button type="button" onClick={marcarTodasProg} className="px-4 py-2 rounded-2xl text-sm font-semibold border border-slate-200 bg-white hover:bg-slate-50">
                Marcar todas as fichas
              </button>
              <button type="button" onClick={desmarcarTodasProg} className="px-4 py-2 rounded-2xl text-sm font-semibold border border-slate-200 bg-white hover:bg-slate-50">
                Desmarcar todas
              </button>
              <span className="text-xs text-slate-500 self-center ml-auto">
                {itensImpressao.length} bloco(s) selecionado(s) · {copiasResumo} ·{" "}
                {blocosPorPaginaEfetivo} bloco(s)/página · {paginasEstimadas} página(s) estimada(s)
              </span>
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="font-bold text-lg text-slate-900">Reserva para top vendas</div>
            <p className="text-sm text-slate-500 mt-1">
              Parte da capacidade de <strong>cada dia</strong> fica reservada para as ref+cor com maior volume de vendas (top N). O sistema preenche primeiro essa fatia com fichas dessas cores; em seguida usa o restante do dia com o rodízio habitual. A sobra da reserva (se nada couber) volta para o bloco geral no mesmo dia. Use <span className="font-semibold">0%</span> para desligar.
            </p>
            <div className="mt-4 grid grid-cols-1 sm:grid-cols-2 gap-4">
              <label className="text-sm font-medium text-slate-700">
                Modo de seleção do top vendas
                <select
                  value={programacaoTopModo}
                  onChange={(e) => setProgramacaoTopModo(e.target.value === "manual" ? "manual" : "auto")}
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
                >
                  <option value="auto">Automático (ranking por vendas)</option>
                  <option value="manual">Manual (eu escolho as ref+cor)</option>
                </select>
              </label>
              <label className="text-sm font-medium text-slate-700">
                % do dia para o top vendas
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={programacaoReservaTopPct}
                  onChange={(e) =>
                    setProgramacaoReservaTopPct(Math.min(100, Math.max(0, Number(e.target.value) || 0)))
                  }
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
                />
              </label>
              <label className="text-sm font-medium text-slate-700">
                Quantidade no ranking (top N)
                <input
                  type="number"
                  min="1"
                  max="200"
                  value={programacaoTopN}
                  onChange={(e) =>
                    setProgramacaoTopN(Math.min(200, Math.max(1, Math.round(Number(e.target.value) || 1))))
                  }
                  className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm"
                />
              </label>
            </div>
            <p className="mt-3 text-xs text-slate-500">
              Com a capacidade atual: Pesponto —{" "}
              <span className="font-semibold text-slate-700">{reservaParesPespontoDia} pares</span> reserva +{" "}
              <span className="font-semibold text-slate-700">{Math.max(0, capPProg - reservaParesPespontoDia)} pares</span> geral (por dia). Montagem —{" "}
              <span className="font-semibold text-slate-700">{reservaParesMontagemDia} pares</span> reserva +{" "}
              <span className="font-semibold text-slate-700">{Math.max(0, capMProg - reservaParesMontagemDia)} pares</span> geral.
            </p>
            {programacaoTopModo === "manual" ? (
              <p className="mt-3 text-xs text-slate-500">
                Seleção manual ativa: <span className="font-semibold text-slate-700">{programacaoTopManualKeys.length}</span> ref+cor marcada(s).
                Para editar a lista, use a aba <span className="font-semibold">Vendas</span> no card <span className="font-semibold">Top vendas manual</span>.
              </p>
            ) : null}
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
              <div>
                <div className="font-bold text-lg">Período da programação</div>
                <div className="text-sm text-slate-500 mt-1">Escolha quantos dias quer planejar: 1, 3, 7, 15 ou outro período personalizado.</div>
              </div>
              <div className="flex flex-col gap-3 lg:items-end print:hidden">
                <div className="flex flex-wrap gap-2">
                  {quickDays.map((day) => (
                    <button
                      key={day}
                      type="button"
                      onClick={() => setProgramacaoDias(day)}
                      className={`px-4 py-2 rounded-2xl text-sm font-semibold border ${programacaoDias === day ? "bg-[#0F172A] text-white border-[#0F172A]" : "bg-white text-[#0F172A] border-slate-200"}`}
                    >
                      {day} dia{day > 1 ? "s" : ""}
                    </button>
                  ))}
                </div>
                <label className="text-sm font-medium text-slate-700">
                  Quantidade personalizada de dias
                  <input
                    type="number"
                    min="1"
                    value={programacaoDias}
                    onChange={(e) => setProgramacaoDias(Math.max(1, Number(e.target.value) || 1))}
                    className="mt-2 w-40 rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
                  />
                </label>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-4 md:p-5">
            <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
              <div>
                <div className="font-bold text-lg">Visualização da programação</div>
                <div className="text-sm text-slate-500 mt-1">
                  Alterne entre a visão por dia e a visão com todas as fichas abertas.
                </div>
              </div>

              <div className="flex flex-wrap gap-2 print:hidden">
                <button
                  type="button"
                  onClick={() => setProgramacaoModoVisual("normal")}
                  className={`px-4 py-2.5 rounded-2xl text-sm font-semibold border transition ${
                    programacaoModoVisual === "normal"
                      ? "bg-[#0F172A] text-white border-[#0F172A]"
                      : "bg-white text-[#0F172A] border-slate-200 hover:bg-slate-50"
                  }`}
                >
                  Por dia
                </button>

                <button
                  type="button"
                  onClick={() => setProgramacaoModoVisual("completo")}
                  className={`px-4 py-2.5 rounded-2xl text-sm font-semibold border transition ${
                    programacaoModoVisual === "completo"
                      ? "bg-[#8B1E2D] text-white border-[#8B1E2D]"
                      : "bg-white text-[#8B1E2D] border-[#E7C7CC] hover:bg-[#FFF7F8]"
                  }`}
                >
                  Todas as fichas
                </button>

                <button
                  type="button"
                  onClick={imprimirProgramacaoDia}
                  className="px-4 py-2.5 rounded-2xl text-sm font-semibold border bg-slate-950 text-white border-slate-950"
                >
                  Imprimir seleção
                </button>
                <button
                  type="button"
                  onClick={gerarPdfWhatsappProgramacao}
                  disabled={programacaoPdfBusy}
                  className="px-4 py-2.5 rounded-2xl text-sm font-semibold border border-[#128C7E] bg-[#25D366] text-white shadow-sm hover:bg-[#20BD5A] hover:border-[#128C7E] disabled:opacity-60 disabled:pointer-events-none"
                >
                  {programacaoPdfBusy ? "A gerar PDF…" : "PDF p/ WhatsApp"}
                </button>
                <button
                  type="button"
                  onClick={recalcularProgramacaoCongelada}
                  className="px-4 py-2.5 rounded-2xl text-sm font-semibold border border-slate-300 bg-white text-slate-900 hover:bg-slate-50"
                >
                  Recalcular plano
                </button>
              </div>
            </div>
            <p className="mt-2 text-xs text-slate-500 print:hidden max-w-xl">
              Gera PDF a partir da mesma folha que a impressão (layout, cabeçalho, cópias por página e fichas). Partilhe ou descarregue e anexe no WhatsApp como documento.
            </p>

            <div className="mt-4 flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
              <div>
                <div className="font-bold text-lg">Setor da programação</div>
                <div className="text-sm text-slate-500 mt-1">
                  Use as sub abas para alternar entre Pesponto e Montagem sem poluir a tela.
                </div>
              </div>

              <div className="flex flex-wrap gap-2 print:hidden">
                {subAbas.map((aba) => (
                  <button
                    key={aba.key}
                    type="button"
                    onClick={() => setProgramacaoSubAba(aba.key)}
                    className={`px-4 py-2.5 rounded-2xl text-sm font-semibold border transition ${
                      programacaoSubAba === aba.key
                        ? "bg-[#8B1E2D] text-white border-[#8B1E2D]"
                        : "bg-[#FFF7F8] text-[#0F172A] border-slate-200 hover:bg-white"
                    }`}
                  >
                    {aba.key}
                  </button>
                ))}
              </div>
            </div>

            <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-600">
              <span className="font-semibold text-slate-900">{subAbaAtiva.titulo}</span>
              <span className="ml-2">{subAbaAtiva.descricao}</span>
            </div>
          </div>

          {programacaoModoVisual === "normal"
            ? renderBlocoProgramacao(subAbaAtiva.programacao, subAbaAtiva.corTag, fichasIncompativeisAtivas)
            : renderModoCompleto()}
          </div>

          <div
            id="programacao-print-sheet-root"
            ref={programacaoPrintSheetRef}
            className="hidden print:block programacao-print-sheet"
          >
            {programacaoTipoFolha === "ambas" ? (
              <>
                <ProgramacaoDiaFolhaImpressao
                  titulo={`${tituloFolhaImpressao} - Folha 1`}
                  logoSrc={programacaoLogoImpressao?.trim() || "/logo-rockstar-bandeira.png"}
                  setor={programacaoSubAba}
                  modoLabel={programacaoModoVisual === "normal" ? "Visão: por dia" : "Visão: todas as fichas"}
                  diasCount={diasPlanoExibicao}
                  dataImpressao={new Date().toLocaleString("pt-BR")}
                  observacoes={programacaoObsImpressao}
                  itens={itensImpressaoFolha1}
                  sizesList={sizes}
                  copiasPorPagina={3}
                  etiquetaFichaCustom={programacaoEtiquetaFicha}
                  cabecalhoFolha={programacaoCabecalhoFolha}
                  valoresParTerceiros={programacaoValoresTerceiros}
                  tipoFolhaImpressao="folha1"
                />
                <div className="programacao-print-item-break" />
                <ProgramacaoDiaFolhaImpressao
                  titulo={`${tituloFolhaImpressao} - Folha 2`}
                  logoSrc={programacaoLogoImpressao?.trim() || "/logo-rockstar-bandeira.png"}
                  setor={programacaoSubAba}
                  modoLabel={programacaoModoVisual === "normal" ? "Visão: por dia" : "Visão: todas as fichas"}
                  diasCount={diasPlanoExibicao}
                  dataImpressao={new Date().toLocaleString("pt-BR")}
                  observacoes={programacaoObsImpressao}
                  itens={itensImpressaoFolha2}
                  sizesList={sizes}
                  copiasPorPagina={programacaoCopiasPorPagina}
                  etiquetaFichaCustom={programacaoEtiquetaFicha}
                  cabecalhoFolha={programacaoCabecalhoFolha}
                  valoresParTerceiros={programacaoValoresTerceiros}
                  tipoFolhaImpressao="folha2"
                />
              </>
            ) : (
              <ProgramacaoDiaFolhaImpressao
                titulo={tituloFolhaImpressao}
                logoSrc={programacaoLogoImpressao?.trim() || "/logo-rockstar-bandeira.png"}
                setor={programacaoSubAba}
                modoLabel={programacaoModoVisual === "normal" ? "Visão: por dia" : "Visão: todas as fichas"}
                diasCount={diasPlanoExibicao}
                dataImpressao={new Date().toLocaleString("pt-BR")}
                observacoes={programacaoObsImpressao}
                itens={itensImpressaoComCopias}
                sizesList={sizes}
                copiasPorPagina={programacaoCopiasPorPagina}
                etiquetaFichaCustom={programacaoEtiquetaFicha}
                cabecalhoFolha={programacaoCabecalhoFolha}
                valoresParTerceiros={programacaoValoresTerceiros}
                tipoFolhaImpressao={programacaoTipoFolha}
              />
            )}
          </div>
        </div>
      </PageShell>
    );
  };

  const renderRelatorioProducao = () => {
    const relatorioRefs = ["TODAS", ...Array.from(new Set(rows.map((row) => row.ref)))];
    const relatorioCores = ["TODAS", ...Array.from(new Set(rows.map((row) => row.cor))).sort((a, b) => a.localeCompare(b, "pt-BR", { sensitivity: "base" }))];
    const parseDateBr = (value) => {
      if (!value) return null;
      const parts = String(value).split("/");
      if (parts.length !== 3) return null;
      const [dd, mm, yyyy] = parts;
      return new Date(`${yyyy}-${mm}-${dd}T00:00:00`);
    };

    const parseDateInput = (value, endOfDay = false) => {
      if (!value) return null;
      return new Date(`${value}T${endOfDay ? "23:59:59" : "00:00:00"}`);
    };

    const inicio = parseDateInput(relatorioDataInicial, false);
    const fim = parseDateInput(relatorioDataFinal, true);

    const base = [
      ...pespontoLancamentos.map((item) => ({ ...item, setor: "Pesponto" })),
      ...montagemLancamentos.map((item) => ({ ...item, setor: "Montagem" })),
    ];

    const linhas = base.filter((item) => {
      if (relatorioRef !== "TODAS" && item.ref !== relatorioRef) return false;
      if (relatorioCor !== "TODAS" && item.cor !== relatorioCor) return false;
      if (relatorioSetor !== "TODOS" && item.setor !== relatorioSetor) return false;
      if (relatorioStatus !== "TODOS") {
        if (relatorioStatus === "EM_ABERTO" && item.status !== "Em aberto") return false;
        if (relatorioStatus === "FINALIZADO" && item.status !== "Finalizado") return false;
      }

      const dataReferencia = item.status === "Finalizado" && item.dataFinalizacao ? item.dataFinalizacao : item.dataLancamento;
      const dataItem = parseDateBr(dataReferencia);
      if (inicio && dataItem && dataItem < inicio) return false;
      if (fim && dataItem && dataItem > fim) return false;
      return true;
    });

    const totalPares = linhas.reduce((acc, item) => acc + (item.total || 0), 0);
    const emAberto = linhas.filter((item) => item.status === "Em aberto");
    const finalizados = linhas.filter((item) => item.status === "Finalizado");

    const abrirPreviaImpressaoRelatorio = () => {
      if (linhas.length === 0) {
        alert("Nenhum registro para imprimir. Ajuste os filtros.");
        return;
      }
      const periodo =
        relatorioDataInicial || relatorioDataFinal
          ? [relatorioDataInicial, relatorioDataFinal].filter(Boolean).join(" → ")
          : "—";
      const filtros = {
        periodo,
        ref: relatorioRef,
        cor: relatorioCor,
        setor: relatorioSetor === "TODOS" ? "Todos" : relatorioSetor,
        status:
          relatorioStatus === "TODOS"
            ? "Todos"
            : relatorioStatus === "EM_ABERTO"
              ? "Em aberto"
              : "Finalizado",
      };
      const resumo = {
        programacoes: linhas.length,
        totalPares: linhas.reduce((acc, item) => acc + (item.total || 0), 0),
        emAberto: linhas.filter((item) => item.status === "Em aberto").length,
        finalizados: linhas.filter((item) => item.status === "Finalizado").length,
      };
      const gruposPorSetor = ["Pesponto", "Montagem"].map((setor) => ({
        setor,
        linhas: linhas.filter((item) => item.setor === setor),
      }));
      setRelatorioPrintDraft({
        dataGeracao: new Date().toLocaleString("pt-BR"),
        filtros,
        resumo,
        gruposPorSetor,
      });
      setRelatorioPrintTitulo("RELATÓRIO DE PRODUÇÃO");
      setRelatorioPrintObs("");
      setRelatorioPrintModalOpen(true);
    };

    return (
      <PageShell title="Relatório de Produção" subtitle="Filtre por período, setor e status para acompanhar tudo que está em aberto ou que foi finalizado.">
        <section className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-6 gap-4">
            <label className="text-sm font-medium text-slate-700">
              Data inicial
              <input
                type="date"
                value={relatorioDataInicial}
                onChange={(e) => setRelatorioDataInicial(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              />
            </label>
            <label className="text-sm font-medium text-slate-700">
              Data final
              <input
                type="date"
                value={relatorioDataFinal}
                onChange={(e) => setRelatorioDataFinal(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              />
            </label>
            <label className="text-sm font-medium text-slate-700">
              REF
              <select
                value={relatorioRef}
                onChange={(e) => setRelatorioRef(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                {relatorioRefs.map((ref) => (
                  <option key={ref} value={ref}>{ref}</option>
                ))}
              </select>
            </label>
            <label className="text-sm font-medium text-slate-700">
              Cor
              <select
                value={relatorioCor}
                onChange={(e) => setRelatorioCor(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                {relatorioCores.map((cor) => (
                  <option key={cor} value={cor}>{cor}</option>
                ))}
              </select>
            </label>
            <label className="text-sm font-medium text-slate-700">
              Setor
              <select
                value={relatorioSetor}
                onChange={(e) => setRelatorioSetor(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                <option value="TODOS">Todos</option>
                <option value="Pesponto">Pesponto</option>
                <option value="Montagem">Montagem</option>
              </select>
            </label>
            <label className="text-sm font-medium text-slate-700">
              Status
              <select
                value={relatorioStatus}
                onChange={(e) => setRelatorioStatus(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                <option value="TODOS">Todos</option>
                <option value="EM_ABERTO">Em aberto</option>
                <option value="FINALIZADO">Finalizado</option>
              </select>
            </label>
          </div>
        </section>

        <section className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
          <SummaryCard title="Programações" value={linhas.length} subtitle="Registros no filtro" />
          <SummaryCard title="Total de pares" value={totalPares} subtitle="Volume total" />
          <SummaryCard title="Em aberto" value={emAberto.length} subtitle="Programações pendentes" />
          <SummaryCard title="Finalizados" value={finalizados.length} subtitle="Programações concluídas" />
        </section>

        <section className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex items-center justify-between gap-4 mb-4">
            <div>
              <h2 className="font-bold text-lg">Resultado do período</h2>
              <p className="text-sm text-slate-500 mt-1">A data usada no filtro é a de finalização para itens finalizados e a de lançamento para itens em aberto.</p>
            </div>
            <div className="flex flex-col items-end gap-1">
              <span className="text-sm text-slate-500">{linhas.length} linha(s)</span>
              <button
                type="button"
                onClick={abrirPreviaImpressaoRelatorio}
                className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold shadow-sm hover:bg-slate-800 transition"
              >
                Prévia e impressão
              </button>
            </div>
          </div>

          {linhas.length === 0 ? (
            <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center text-sm text-slate-500">
              Nenhum registro encontrado para o filtro selecionado.
            </div>
          ) : (
            <div className="overflow-auto">
              <table className="min-w-[1200px] w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-slate-50 text-slate-500">
                    <th className="border px-4 py-3 text-left">Data</th>
                    <th className="border px-4 py-3 text-left">Setor</th>
                    <th className="border px-4 py-3 text-left">Programação</th>
                    <th className="border px-4 py-3 text-left">Ref</th>
                    <th className="border px-4 py-3 text-left">Cor</th>
                    <th className="border px-4 py-3 text-center">Pares</th>
                    <th className="border px-4 py-3 text-center">Status</th>
                    <th className="border px-4 py-3 text-left">Detalhe</th>
                  </tr>
                </thead>
                <tbody>
                  {linhas.map((item) => {
                    const dataRef = item.status === "Finalizado" && item.dataFinalizacao ? item.dataFinalizacao : item.dataLancamento;
                    return (
                      <tr key={`${item.setor}-${item.id}`}>
                        <td className="border px-4 py-3">{dataRef || "-"}</td>
                        <td className="border px-4 py-3 font-semibold">{item.setor}</td>
                        <td className="border px-4 py-3">{item.programacao}</td>
                        <td className="border px-4 py-3">{item.ref}</td>
                        <td className="border px-4 py-3">{item.cor}</td>
                        <td className="border px-4 py-3 text-center font-semibold">{item.total}</td>
                        <td className="border px-4 py-3 text-center">
                          <span className={`px-3 py-1 rounded-full border text-xs font-semibold ${item.status === "Finalizado" ? "bg-emerald-100 text-emerald-700 border-emerald-200" : "bg-amber-100 text-amber-700 border-amber-200"}`}>
                            {item.status}
                          </span>
                        </td>
                        <td className="border px-4 py-3">
                          <div className="flex flex-wrap gap-2">
                            {item.items.map((entry) => (
                              <span key={`${item.id}-${entry.size}`} className="px-2 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs">
                                {entry.size}: {entry.qtd}
                              </span>
                            ))}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </section>
      </PageShell>
    );
  };

  const renderActivePage = () => {
    switch (active) {
      case "Dashboard":
        return renderDashboard();
      case "Controle Geral":
        return renderControle();
      case "Importar GCM":
        return renderImport();
      case "Pesponto":
        return renderMovPage("Pesponto", pespontoForm, setPespontoForm, "Lance pares em produção no pesponto com grade completa.");
      case "Montagem":
        return renderMovPage("Montagem", montagemForm, setMontagemForm, "Lance pares em montagem com grade completa.");
      case "Costura Pronta":
        return renderCosturaPronta();
      case "Minimos":
        return renderConfig();
      case "Vendas":
        return renderVendas();
      case "Sugestoes":
        return renderSugestoes();
      case "Gerador de Fichas":
        return renderFichas();
      case "Programação do Dia":
        return renderProgramacaoDia();
      case "Relatório de Produção":
        return renderRelatorioProducao();
      default:
        return renderControle();
    }
  };

  const fecharPreviaRelatorio = () => {
    setRelatorioPrintModalOpen(false);
    setRelatorioPrintDraft(null);
  };

  const confirmarImpressaoRelatorio = () => {
    if (!relatorioPrintDraft) return;
    setPrintRelatorioData({
      ...relatorioPrintDraft,
      titulo: relatorioPrintTitulo.trim() || "RELATÓRIO DE PRODUÇÃO",
      observacoes: relatorioPrintObs.trim(),
    });
    setRelatorioPrintModalOpen(false);
    setRelatorioPrintDraft(null);
  };

  const menuLabels = {
    "Dashboard": "Painel",
    "Controle Geral": "Painel",
    "Importar GCM": "Entrada",
    "Pesponto": "Operação",
    "Montagem": "Operação",
    "Costura Pronta": "Operação",
    "Minimos": "Planejamento",
    "Vendas": "Planejamento",
    "Sugestoes": "Planejamento",
    "Gerador de Fichas": "Planejamento",
    "Programação do Dia": "Execução",
    "Relatório de Produção": "Relatórios",
  };

  return (
    <div className="min-h-screen min-h-[100dvh] bg-[radial-gradient(circle_at_top,_#FDF2F4_0%,_#FFFFFF_30%,_#F8FAFC_100%)] text-slate-900">
      <style>{`
        @page {
          size: A4;
          margin: 10mm;
        }
        html.programacao-pdf-capture .programacao-print-sheet.programacao-print-sheet--capture {
          display: block !important;
          position: fixed !important;
          left: -12000px !important;
          top: 0 !important;
          width: 190mm !important;
          max-width: 190mm !important;
          background: #ffffff !important;
          padding: 0 !important;
          z-index: 2147483646 !important;
          overflow: visible !important;
          visibility: visible !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-ficha {
          break-inside: avoid;
          page-break-inside: avoid;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-doc {
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-doc--cabecalho-oculto {
          margin-top: 0 !important;
          padding-top: 0 !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-grid-economico {
          display: grid !important;
          grid-template-columns: 1fr !important;
          gap: 4mm !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-ficha-economica {
          padding: 2mm !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-ficha-economica table {
          font-size: 7pt !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-economico-3 .programacao-print-ficha-economica {
          padding: 1.6mm !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-economico-3 .programacao-print-ficha-economica table {
          font-size: 6.4pt !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-economico-4 .programacao-print-ficha-economica {
          padding: 1.2mm !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-economico-4 .programacao-print-ficha-economica table {
          font-size: 5.8pt !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-grade-table th,
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-grade-table td {
          vertical-align: middle !important;
          box-sizing: border-box !important;
          overflow: visible !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-grade-table thead th {
          color: #ffffff !important;
          background-color: #475569 !important;
          border-color: #64748b !important;
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }
        html.programacao-pdf-capture .programacao-print-sheet--capture .programacao-print-grade-table thead tr {
          background-color: #475569 !important;
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }
        @media print {
          body * { visibility: hidden !important; }
          #print-root, #print-root *,
          #print-programacao-root, #print-programacao-root *,
          #print-mov-root, #print-mov-root * { visibility: visible !important; }
          #print-root, #print-programacao-root, #print-mov-root {
            position: absolute;
            left: 0;
            top: 0;
            width: 100%;
            background: white;
            padding: 0;
            color: #0f172a;
          }
          #print-root {
            padding: 24px;
          }
          .programacao-print-sheet {
            padding: 0;
            max-width: 100%;
          }
          #print-root .print-section,
          #print-programacao-root .print-section { page-break-inside: avoid; }
          .programacao-print-ficha {
            break-inside: avoid;
            page-break-inside: avoid;
            break-inside: avoid-page;
          }
          .programacao-print-ref-block {
            break-inside: avoid;
            page-break-inside: avoid;
            break-inside: avoid-page;
          }
          .programacao-print-doc {
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          .programacao-print-doc--cabecalho-oculto {
            margin-top: 0;
            padding-top: 0;
          }
          .programacao-print-grid-economico {
            display: grid;
            grid-template-columns: 1fr;
            gap: 4mm;
          }
          .programacao-print-ficha-economica {
            padding: 2mm;
          }
          .programacao-print-ficha-economica table {
            font-size: 7pt !important;
          }
          .programacao-print-economico-3 .programacao-print-ficha-economica {
            padding: 1.6mm;
          }
          .programacao-print-economico-3 .programacao-print-ficha-economica table {
            font-size: 6.4pt !important;
          }
          .programacao-print-economico-4 .programacao-print-ficha-economica {
            padding: 1.2mm;
          }
          .programacao-print-economico-4 .programacao-print-ficha-economica table {
            font-size: 5.8pt !important;
          }
          .programacao-print-grade-table th,
          .programacao-print-grade-table td {
            vertical-align: middle !important;
            box-sizing: border-box !important;
            overflow: visible !important;
          }
          /* Cabeçalho da grade: #print-programacao-root usa color escuro; força branco no fundo cinza (P&B). */
          .programacao-print-grade-table thead th {
            color: #ffffff !important;
            background-color: #475569 !important;
            border-color: #64748b !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .programacao-print-grade-table thead tr {
            background-color: #475569 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .programacao-print-item-break {
            break-after: page;
            page-break-after: always;
          }
        }
      `}</style>
      <div className="flex min-h-[100dvh] min-h-screen">
        <aside className="hidden lg:flex lg:w-[280px] lg:shrink-0 lg:flex-col lg:justify-between lg:border-r lg:border-white/50 lg:bg-[#0F172A] lg:p-5 lg:text-white lg:shadow-[20px_0_60px_rgba(15,23,42,0.18)] xl:w-[310px] xl:p-6">
          <div>
            <div className="rounded-[30px] border border-white/10 bg-white/5 p-5 shadow-inner shadow-white/5">
              <img src="/logo-rockstar.png" alt="Rock Star" className="h-10 object-contain" />
              <div className="mt-3 text-3xl font-black tracking-tight">Módulo Produção</div>
              <p className="mt-3 text-sm leading-6 text-slate-300">
                Painel inteligente para estoque, programação, cobertura e execução da produção.
              </p>
            </div>

            <div className="mt-6 space-y-5">
              {Array.from(new Set(navItems.map((item) => menuLabels[item]))).map((grupo) => (
                <div key={grupo}>
                  <div className="mb-2 px-2 text-[11px] font-semibold uppercase tracking-[0.22em] text-slate-500">{grupo}</div>
                  <nav className="space-y-1.5">
                    {navItems.filter((item) => menuLabels[item] === grupo).map((item) => (
                      <button
                        key={item}
                        onClick={() => setActive(item)}
                        className={`group flex w-full items-center justify-between rounded-2xl px-4 py-3 text-left text-sm font-semibold transition-all ${
                          active === item
                            ? "bg-white text-[#0F172A] shadow-[0_10px_25px_rgba(255,255,255,0.12)]"
                            : "text-slate-300 hover:bg-white/8 hover:text-white"
                        }`}
                      >
                        <span>{item}</span>
                        <span className={`h-2.5 w-2.5 rounded-full ${active === item ? "bg-[#8B1E2D]" : "bg-slate-700 group-hover:bg-slate-500"}`} />
                      </button>
                    ))}
                  </nav>
                </div>
              ))}
            </div>
          </div>

          <div className="rounded-[26px] border border-white/10 bg-white/5 p-4">
            <div className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-500">Tela ativa</div>
            <div className="mt-2 text-lg font-bold text-white">{active}</div>
            <div className="mt-2 text-sm text-slate-400">Layout refinado para leitura mais rápida e operação mais segura.</div>
          </div>
        </aside>

        <main className="app-main-shell flex-1 min-w-0 p-3 sm:p-4 lg:p-6 xl:p-8 max-[1023px]:landscape:py-2 max-[1023px]:landscape:px-3">
          <div className="mx-auto max-w-[1720px] space-y-3 sm:space-y-4 max-[1023px]:landscape:space-y-2">
            <div className="app-mobile-nav-card lg:hidden rounded-[22px] sm:rounded-[26px] border border-slate-200 bg-white p-3 sm:p-4 shadow-[0_10px_30px_rgba(15,23,42,0.06)] backdrop-blur max-[1023px]:landscape:p-3 max-[1023px]:landscape:rounded-2xl">
              <img src="/logo-rockstar.png" alt="Rock Star" className="h-7 sm:h-8 object-contain max-[1023px]:landscape:h-6" />
              <div className="mt-1.5 sm:mt-2 text-xl sm:text-2xl font-black tracking-tight text-slate-950 max-[1023px]:landscape:text-lg max-[1023px]:landscape:mt-1">Módulo Produção</div>
              <div className="mt-3 sm:mt-4 flex gap-2 overflow-x-auto overflow-y-hidden pb-1 scroll-smooth touch-pan-x [-webkit-overflow-scrolling:touch] snap-x snap-mandatory">
                {navItems.map((item) => (
                  <button
                    key={item}
                    onClick={() => setActive(item)}
                    className={`snap-start shrink-0 whitespace-nowrap rounded-2xl border px-3 py-2 sm:px-4 sm:py-2.5 text-xs sm:text-sm font-semibold transition max-[1023px]:landscape:px-3 max-[1023px]:landscape:py-2 max-[1023px]:landscape:text-[13px] ${
                      active === item
                        ? "border-[#0F172A] bg-[#0F172A] text-white"
                        : "border-[#E5E7EB] bg-[#FFF7F8] text-[#0F172A]"
                    }`}
                  >
                    {item}
                  </button>
                ))}
              </div>
            </div>

            <div className="max-w-[1720px]">{renderActivePage()}</div>
          </div>
        </main>
      </div>

      {lancarFichaDaProgramacao && (
        <div className="fixed inset-0 z-[55] flex items-center justify-center bg-slate-950/50 p-4 overflow-auto max-[1023px]:landscape:items-start max-[1023px]:landscape:py-4">
          <div className="w-full max-w-4xl rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6 max-h-[min(92dvh,900px)] overflow-y-auto">
            <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-4">
              <div>
                <div className="text-lg font-bold">Lançar ficha no sistema</div>
                <div className="text-sm text-slate-500 mt-1">
                  {lancarFichaDaProgramacao.tipo} · {lancarFichaDaProgramacao.ref} • {lancarFichaDaProgramacao.cor}
                </div>
              </div>
              <button
                type="button"
                onClick={() => setLancarFichaDaProgramacao(null)}
                className="rounded-xl border border-slate-200 px-3 py-2 text-sm font-semibold shrink-0"
              >
                Fechar
              </button>
            </div>

            <label className="mt-6 block text-sm font-medium text-slate-700">
              Nome da programação
              <input
                type="text"
                value={lancarFichaDaProgramacao.programacao}
                onChange={(e) =>
                  setLancarFichaDaProgramacao((m) =>
                    m ? { ...m, programacao: e.target.value } : null
                  )
                }
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm font-semibold"
                placeholder="Ex.: Programação semana 12"
              />
            </label>

            <div className="mt-6 overflow-x-auto">
              <table className="w-full border-collapse text-sm min-w-[640px]">
                <thead>
                  <tr className="bg-slate-50">
                    {sizes.map((size) => (
                      <th key={`lanc-prog-h-${size}`} className="border border-slate-200 px-2 py-2 text-center font-bold text-slate-700">
                        {size}
                      </th>
                    ))}
                    <th className="border border-slate-200 px-3 py-2 text-center font-bold text-slate-700">Total</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    {sizes.map((size) => (
                      <td key={`lanc-prog-c-${size}`} className="border border-slate-200 px-1 py-2 text-center">
                        <input
                          type="number"
                          min={0}
                          value={lancarFichaDaProgramacao.grid[size] ?? 0}
                          onChange={(e) => {
                            const v = Math.max(0, Number(e.target.value) || 0);
                            setLancarFichaDaProgramacao((m) =>
                              m ? { ...m, grid: { ...m.grid, [size]: v } } : null
                            );
                          }}
                          className="w-full max-w-[4.5rem] mx-auto rounded-xl border border-slate-200 bg-white px-2 py-2 text-center text-sm font-semibold"
                        />
                      </td>
                    ))}
                    <td className="border border-slate-200 px-3 py-2 text-center font-bold text-slate-900">
                      {sizes.reduce((acc, s) => acc + (Number(lancarFichaDaProgramacao.grid[s]) || 0), 0)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {movError[lancarFichaDaProgramacao.tipo] ? (
              <div className="mt-4 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-medium text-red-700">
                {movError[lancarFichaDaProgramacao.tipo]}
              </div>
            ) : null}

            <div className="mt-6 flex flex-wrap gap-3 justify-end">
              <button
                type="button"
                onClick={() => setLancarFichaDaProgramacao(null)}
                className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold bg-white"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={() => {
                  const m = lancarFichaDaProgramacao;
                  if (!m) return;
                  executeMov(
                    m.tipo,
                    {
                      ref: m.ref,
                      cor: m.cor,
                      programacao: m.programacao,
                      grid: { ...m.grid },
                    },
                    false,
                    m.fichaStorageKey || undefined
                  );
                }}
                className="rounded-2xl bg-[#8B1E2D] text-white px-4 py-3 text-sm font-semibold hover:bg-[#6F1421]"
              >
                Confirmar lançamento
              </button>
            </div>
          </div>
        </div>
      )}

      {previewFicha && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 overflow-auto max-[1023px]:landscape:items-start max-[1023px]:landscape:py-4">
          <div className="w-full max-w-3xl rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6 max-h-[min(92dvh,900px)] overflow-y-auto">
            <div className="flex items-center justify-between gap-4">
              <div>
                <div className="text-lg font-bold">Pré-visualização da ficha</div>
                <div className="text-sm text-slate-500 mt-1">{previewFicha.nome}</div>
                <div className="text-sm text-slate-500">{previewFicha.ref} • {previewFicha.cor}</div>
              </div>
              <button onClick={() => setPreviewFicha(null)} className="rounded-xl border border-slate-200 px-3 py-2 text-sm font-semibold">Fechar</button>
            </div>

            <div className="mt-6 overflow-auto">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-slate-50">
                    {sizes.map((size) => (
                      <th key={size} className="border border-slate-200 px-4 py-3 text-center">{size}</th>
                    ))}
                    <th className="border border-slate-200 px-4 py-3 text-center">Total</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    {sizes.map((size) => (
                      <td key={size} className="border border-slate-200 px-4 py-3 text-center">{previewFicha.sizes[size] || 0}</td>
                    ))}
                    <td className="border border-slate-200 px-4 py-3 text-center font-bold">{previewFicha.total}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {confirmImport && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 max-[1023px]:landscape:items-start max-[1023px]:landscape:py-4">
          <div className="w-full max-w-md max-h-[min(88dvh,800px)] overflow-y-auto rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
            <div className="text-lg font-bold">Confirmar importação do GCM</div>
            <p className="text-sm text-slate-600 mt-3 leading-relaxed">Essa ação vai atualizar o Produto Acabado (PA) com base no arquivo carregado.</p>
            <div className="mt-6 flex gap-3 justify-end">
              <button onClick={() => setConfirmImport(false)} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold bg-white">Cancelar</button>
              <button onClick={executeImport} className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold">Confirmar importação</button>
            </div>
          </div>
        </div>
      )}

      {confirmMov && (() => {
        const invalidos = sizes
          .map((size) => ({ size, qtd: Number(confirmMov.form.grid[size]) || 0 }))
          .filter((item) => item.qtd > 0 && item.qtd % 12 !== 0);
        const lista = invalidos.map((item) => `${item.size} (${item.qtd})`).join(", ");
        const totalLancamento = sizes.reduce((acc, size) => acc + (Number(confirmMov.form.grid[size]) || 0), 0);
        const mensagens = [];

        if (invalidos.length) {
          mensagens.push(`No ${confirmMov.tipo}, o padrão é trabalhar em múltiplos de 12. Fora da regra em: ${lista}.`);
        }

        if (totalLancamento > 396) {
          mensagens.push(`O total informado é ${totalLancamento} pares e não pode passar de 396.`);
        }

        return (
          <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/50 p-4 max-[1023px]:landscape:items-start max-[1023px]:landscape:py-4">
            <div className="w-full max-w-lg max-h-[min(88dvh,800px)] overflow-y-auto rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
              <div className="text-lg font-bold">Lançamento fora da regra</div>
              <p className="text-sm text-slate-600 mt-3 leading-relaxed">
                {mensagens.join(" ")}
              </p>
              <p className="text-sm text-slate-600 mt-3 leading-relaxed">Deseja realmente continuar com esse lançamento?</p>
              <div className="mt-6 flex gap-3 justify-end">
                <button onClick={() => setConfirmMov(null)} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold bg-white">Cancelar</button>
                <button
                  onClick={() =>
                    executeMov(confirmMov.tipo, confirmMov.form, true, confirmMov.progFichaStorageKey || undefined)
                  }
                  className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold"
                >
                  Lançar mesmo assim
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {confirmAction && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 max-[1023px]:landscape:items-start max-[1023px]:landscape:py-4">
          <div className="w-full max-w-lg max-h-[min(88dvh,800px)] overflow-y-auto rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
            <div className="text-lg font-bold">{confirmAction.titulo}</div>
            <p className="text-sm text-slate-600 mt-3 leading-relaxed">{confirmAction.mensagem}</p>
            <div className="mt-6 flex gap-3 justify-end">
              <button onClick={() => setConfirmAction(null)} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold bg-white">Cancelar</button>
              <button
                onClick={() => {
                  if (confirmAction.kind === "delete") {
                    confirmDeleteLancamento(confirmAction);
                  } else if (confirmAction.kind === "finalizar") {
                    confirmFinalizarProgramacao(confirmAction);
                  }
                }}
                className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold"
              >
                Confirmar
              </button>
            </div>
          </div>
        </div>
      )}

      {relatorioPrintModalOpen && relatorioPrintDraft && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/60 p-4 overflow-y-auto max-[1023px]:landscape:items-start max-[1023px]:landscape:py-3">
          <div className="w-full max-w-4xl rounded-[28px] bg-white shadow-2xl border border-slate-200 my-4 sm:my-8 flex flex-col max-h-[min(92dvh,90vh)]">
            <div className="px-6 py-5 border-b border-slate-200 shrink-0">
              <div className="text-lg font-bold text-slate-900">Prévia do relatório</div>
              <p className="text-sm text-slate-500 mt-1">Ajuste o título e as observações; a prévia atualiza em tempo real.</p>
              <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4">
                <label className="text-sm font-medium text-slate-700 block">
                  Título na impressão
                  <input
                    type="text"
                    value={relatorioPrintTitulo}
                    onChange={(e) => setRelatorioPrintTitulo(e.target.value)}
                    className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm font-semibold text-[#0F172A]"
                    placeholder="RELATÓRIO DE PRODUÇÃO"
                  />
                </label>
                <div className="md:col-span-2">
                  <label className="text-sm font-medium text-slate-700 block">
                    Observações (aparecem no relatório impresso)
                    <textarea
                      value={relatorioPrintObs}
                      onChange={(e) => setRelatorioPrintObs(e.target.value)}
                      rows={3}
                      className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm resize-y min-h-[80px]"
                      placeholder="Ex.: conferir prioridades da semana, observações para o setor..."
                    />
                  </label>
                </div>
              </div>
            </div>
            <div className="px-6 py-4 overflow-y-auto flex-1 min-h-0 border-b border-slate-100 bg-slate-100/80">
              <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-inner">
                <RelatorioProducaoPrintDocument
                  data={{
                    ...relatorioPrintDraft,
                    titulo: relatorioPrintTitulo,
                    observacoes: relatorioPrintObs,
                  }}
                />
              </div>
            </div>
            <div className="px-6 py-4 flex flex-wrap justify-end gap-3 shrink-0 bg-white rounded-b-[28px]">
              <button
                type="button"
                onClick={fecharPreviaRelatorio}
                className="rounded-2xl border border-slate-200 px-5 py-3 text-sm font-semibold bg-white hover:bg-slate-50"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={confirmarImpressaoRelatorio}
                className="rounded-2xl bg-[#8B1E2D] text-white px-5 py-3 text-sm font-semibold hover:bg-[#7a1a28] shadow-sm"
              >
                Imprimir
              </button>
            </div>
          </div>
        </div>
      )}

      {printRelatorioData && (
        <div id="print-root" className="hidden print:block">
          <RelatorioProducaoPrintDocument data={printRelatorioData} />
        </div>
      )}
    </div>
  );
}
