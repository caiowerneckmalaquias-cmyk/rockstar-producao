/** Etiqueta de caixa 10×6 cm para Elgin L42 Pro (ZPL, 203 dpi). */

import { buildBlingSku } from "./constants/bling";
import { sizes } from "./constants/production";

export const ETIQUETA_DPI = 203;
/** 100 mm */
export const ETIQUETA_LARGURA_DOTS = 800;
/** 60 mm */
export const ETIQUETA_ALTURA_DOTS = 480;

export const CONTADOR_STORAGE_KEY = "rockstar-etiqueta-contador-v1";

const pctX = (p) => Math.round((p / 100) * ETIQUETA_LARGURA_DOTS);
const pctY = (p) => Math.round((p / 100) * ETIQUETA_ALTURA_DOTS);

function zplSafe(value) {
  return String(value ?? "")
    .replace(/[\^~\\]/g, " ")
    .trim();
}

export function formatContador(n, width = 7) {
  const num = Math.max(0, Math.floor(Number(n) || 0));
  const w = Math.max(1, Math.min(12, Number(width) || 7));
  return String(num).padStart(w, "0");
}

export function parseContadorInicial(value) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!digits) return { n: 1, width: 7 };
  return { n: Number(digits) || 1, width: Math.max(digits.length, 7) };
}

export function readStoredContador() {
  try {
    const raw = localStorage.getItem(CONTADOR_STORAGE_KEY);
    if (!raw) return "0000001";
    const { n, width } = parseContadorInicial(raw);
    return formatContador(n, width);
  } catch {
    return "0000001";
  }
}

export function writeStoredContador(value) {
  try {
    localStorage.setItem(CONTADOR_STORAGE_KEY, String(value));
  } catch {
    /* ignore */
  }
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error(`Não foi possível carregar ${url}`));
    img.src = url;
  });
}

/**
 * Converte PNG/JPEG em ^GFA (1-bit) para ZPL.
 */
export async function imageUrlToZplGfa(url, destWidthDots, destHeightDots, threshold = 168) {
  const img = await loadImage(url);
  const width = Math.max(8, Math.ceil(destWidthDots / 8) * 8);
  const height = Math.max(1, Math.round(destHeightDots));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.drawImage(img, 0, 0, width, height);
  const { data } = ctx.getImageData(0, 0, width, height);
  const bytesPerRow = width / 8;
  const total = bytesPerRow * height;
  let hex = "";
  for (let y = 0; y < height; y += 1) {
    for (let xb = 0; xb < bytesPerRow; xb += 1) {
      let byte = 0;
      for (let bit = 0; bit < 8; bit += 1) {
        const x = xb * 8 + bit;
        const i = (y * width + x) * 4;
        const lum = data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114;
        if (lum < threshold) byte |= 1 << (7 - bit);
      }
      hex += byte.toString(16).toUpperCase().padStart(2, "0");
    }
  }
  return { hex, bytesPerRow, total, width, height };
}

/** Uma etiqueta 10×6. Sempre ^PQ1 — o lote repete ^XA...^XZ. */
export function buildEtiquetaCaixaZpl(data) {
  const nomeLinha = zplSafe(data.nomeLinha).toUpperCase();
  const corLinha = zplSafe(data.corLinha).toUpperCase();
  const tamanho = zplSafe(data.tamanho);
  const prog = zplSafe(data.prog);
  const sku = zplSafe(data.sku);
  const contador = zplSafe(data.contador);

  const logoX = pctX(8);
  const logoY = pctY(1.1);
  const sizeX = pctX(72);
  const sizeY = pctY(3.3);
  const sizeW = pctX(22);
  const sizeH = pctY(33.3);
  const progX = pctX(32.7);
  const progY = pctY(12);
  const progW = pctX(32.5);
  const titleY1 = pctY(44);
  const titleY2 = pctY(52.5);
  const barY = pctY(66);
  const barH = pctY(22);
  const skuX = pctX(10);
  const ctrX = pctX(66.8);

  // Fontes maiores (203 dpi) — alinhadas ao visual da web / PDF.
  const progLabelFont = 52;
  const progValueFont = 56;
  const titleFont = 40;
  const sizeFont = tamanho.length >= 3 ? 90 : 120;

  const lines = [];
  lines.push("^XA");
  lines.push("^CI28");
  lines.push(`^PW${ETIQUETA_LARGURA_DOTS}`);
  lines.push(`^LL${ETIQUETA_ALTURA_DOTS}`);
  lines.push("^LH0,0");
  lines.push("^LS0");
  lines.push("^MD20");
  lines.push("^PR4");
  lines.push("^FWN");

  if (data.logoGfa?.hex) {
    const g = data.logoGfa;
    lines.push(`^FO${logoX},${logoY}^GFA,${g.total},${g.total},${g.bytesPerRow},${g.hex}^FS`);
  } else {
    lines.push(`^FO${logoX},${logoY + 20}^A0N,36,36^FDROCK^FS`);
    lines.push(`^FO${logoX},${logoY + 60}^A0N,36,36^FDSTAR^FS`);
  }

  lines.push(
    `^FO${progX},${progY}^A0N,${progLabelFont},${progLabelFont}^FB${progW},1,0,C,0^FDPROG:^FS`
  );
  lines.push(
    `^FO${progX},${progY + progLabelFont + 6}^A0N,${progValueFont},${progValueFont}^FB${progW},1,0,C,0^FD${prog || "-"}^FS`
  );

  // Caixa preta + numeração branca centralizada (H via ^FB, V via offset).
  lines.push(`^FO${sizeX},${sizeY}^GB${sizeW},${sizeH},${sizeH}^FS`);
  const sizeTextY = sizeY + Math.max(0, Math.round((sizeH - sizeFont) / 2));
  lines.push(
    `^FO${sizeX},${sizeTextY}^FR^A0N,${sizeFont},${sizeFont}^FB${sizeW},1,0,C,0^FD${tamanho}^FS`
  );

  lines.push(
    `^FO0,${titleY1}^A0N,${titleFont},${titleFont}^FB${ETIQUETA_LARGURA_DOTS},1,0,C,0^FD${nomeLinha || "-"}^FS`
  );
  lines.push(
    `^FO0,${titleY2}^A0N,${titleFont},${titleFont}^FB${ETIQUETA_LARGURA_DOTS},1,0,C,0^FD${corLinha || "-"}^FS`
  );

  lines.push(`^BY2,2,${barH}`);
  lines.push(`^FO${skuX},${barY}^BCN,${barH},Y,N,N^FD${sku}^FS`);
  lines.push(`^FO${ctrX},${barY}^BCN,${barH},Y,N,N^FD${contador}^FS`);

  lines.push("^PQ1,0,1,Y");
  lines.push("^XZ");
  return `${lines.join("\n")}\n`;
}

/**
 * Expande linhas do lote: cada unidade da grade = 1 etiqueta; contador +1.
 */
export function expandLoteToLabels(lote, contadorInicial) {
  const { n: startN, width } = parseContadorInicial(contadorInicial);
  let n = startN;
  const labels = [];

  (lote || []).forEach((linha) => {
    const nomeBase = String(linha.nome || "").trim().toUpperCase();
    const nomeLinha = nomeBase ? (nomeBase.endsWith("-") ? nomeBase : `${nomeBase} -`) : "—";
    const corLinha = String(linha.cor || "").trim().toUpperCase() || "—";
    const prog = String(linha.programacao || "").trim();
    const ref = String(linha.ref || "").trim();
    const sigla = String(linha.sigla || "").trim();

    sizes.forEach((size) => {
      const qtd = Math.max(0, Math.floor(Number(linha.grid?.[size]) || 0));
      if (!qtd) return;
      const sku = buildBlingSku(ref, sigla, size);
      if (!sku) return;
      for (let i = 0; i < qtd; i += 1) {
        labels.push({
          nomeLinha,
          corLinha,
          tamanho: size,
          prog,
          sku,
          contador: formatContador(n, width),
          ref,
          cor: corLinha,
        });
        n += 1;
      }
    });
  });

  return {
    labels,
    nextContador: formatContador(n, width),
    total: labels.length,
  };
}

export function buildEtiquetaLoteZpl(labels, logoGfa = null) {
  return (labels || [])
    .map((label) => buildEtiquetaCaixaZpl({ ...label, logoGfa }))
    .join("");
}

export function downloadTextFile(filename, text, mime = "text/plain;charset=utf-8") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export async function buildAndDownloadEtiquetaLoteZpl({
  lote,
  contadorInicial,
  logoUrl = "/logo-rockstar-etiqueta.png",
  filenameBase = "",
}) {
  const { labels, nextContador, total } = expandLoteToLabels(lote, contadorInicial);
  if (!total) {
    throw new Error("Nenhuma etiqueta no lote (grade vazia ou sem sigla Bling).");
  }

  const logoW = pctX(22);
  const logoH = pctY(38.3);
  let logoGfa = null;
  try {
    logoGfa = await imageUrlToZplGfa(logoUrl, logoW, logoH);
  } catch {
    logoGfa = null;
  }

  const zpl = buildEtiquetaLoteZpl(labels, logoGfa);
  const prog = zplSafe(lote?.[0]?.programacao || "").replace(/\s+/g, "") || "lote";
  const base = zplSafe(filenameBase).replace(/\s+/g, "") || prog;
  downloadTextFile(`${base}_${total}etq.zpl`, zpl);
  writeStoredContador(nextContador);
  return { zpl, labels, nextContador, total };
}
