/** Etiqueta de caixa 10×6 cm para Elgin L42 Pro (ZPL, 203 dpi). */

export const ETIQUETA_DPI = 203;
/** 100 mm */
export const ETIQUETA_LARGURA_DOTS = 800;
/** 60 mm */
export const ETIQUETA_ALTURA_DOTS = 480;

const pctX = (p) => Math.round((p / 100) * ETIQUETA_LARGURA_DOTS);
const pctY = (p) => Math.round((p / 100) * ETIQUETA_ALTURA_DOTS);

function zplSafe(value) {
  return String(value ?? "")
    .replace(/[\^~\\]/g, " ")
    .trim();
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
 * @returns {{ hex: string, bytesPerRow: number, total: number, width: number, height: number }}
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

/**
 * Monta ZPL de uma etiqueta 10×6 com o layout do preview / PDF.
 * @param {{
 *   nomeLinha: string,
 *   corLinha: string,
 *   tamanho: string|number,
 *   prog: string,
 *   sku: string,
 *   contador: string,
 *   copias?: number,
 *   logoGfa?: { hex: string, bytesPerRow: number, total: number, width: number, height: number } | null,
 * }} data
 */
export function buildEtiquetaCaixaZpl(data) {
  const nomeLinha = zplSafe(data.nomeLinha).toUpperCase();
  const corLinha = zplSafe(data.corLinha).toUpperCase();
  const tamanho = zplSafe(data.tamanho);
  const prog = zplSafe(data.prog);
  const sku = zplSafe(data.sku);
  const contador = zplSafe(data.contador);
  const copias = Math.max(1, Math.min(9999, Number(data.copias) || 1));

  const logoX = pctX(8);
  const logoY = pctY(1.1);
  const logoW = pctX(22);
  const logoH = pctY(38.3);

  const sizeX = pctX(72);
  const sizeY = pctY(3.3);
  const sizeW = pctX(22);
  const sizeH = pctY(33.3);

  const progX = pctX(32.7);
  const progY = pctY(14);
  const progW = pctX(32.5);

  const titleY1 = pctY(46);
  const titleY2 = pctY(53);

  const barY = pctY(66);
  const barH = pctY(22);
  const skuX = pctX(10);
  const skuW = pctX(41.3);
  const ctrX = pctX(66.8);
  const ctrW = pctX(23.8);

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
    // Fallback textual se o logo não carregar
    lines.push(`^FO${logoX},${logoY + 20}^A0N,36,36^FDROCK^FS`);
    lines.push(`^FO${logoX},${logoY + 60}^A0N,36,36^FDSTAR^FS`);
  }

  // PROG centrado na faixa do meio-topo
  const progLabel = "PROG:";
  const progValue = prog || "-";
  lines.push(`^FO${progX},${progY}^A0N,32,32^FB${progW},1,0,C,0^FD${progLabel}^FS`);
  lines.push(`^FO${progX},${progY + 40}^A0N,36,36^FB${progW},1,0,C,0^FD${progValue}^FS`);

  // Caixa preta + numeração branca (field reverse)
  lines.push(`^FO${sizeX},${sizeY}^GB${sizeW},${sizeH},${sizeH}^FS`);
  const sizeFont = tamanho.length >= 3 ? 70 : 100;
  const sizeTextX = sizeX + Math.round(sizeW * 0.12);
  const sizeTextY = sizeY + Math.round(sizeH * 0.22);
  lines.push(`^FO${sizeTextX},${sizeTextY}^FR^A0N,${sizeFont},${sizeFont}^FD${tamanho}^FS`);

  // Título centralizado
  lines.push(`^FO0,${titleY1}^A0N,28,28^FB${ETIQUETA_LARGURA_DOTS},1,0,C,0^FD${nomeLinha || "-"}^FS`);
  lines.push(`^FO0,${titleY2}^A0N,28,28^FB${ETIQUETA_LARGURA_DOTS},1,0,C,0^FD${corLinha || "-"}^FS`);

  // Code 128 — SKU (esteira) e contador; mesma altura de barra
  // ^BCN,height,printInterpretation,printAbove,checkDigit
  lines.push(`^BY2,2,${barH}`);
  lines.push(`^FO${skuX},${barY}^BCN,${barH},Y,N,N^FD${sku}^FS`);
  lines.push(`^FO${ctrX},${barY}^BCN,${barH},Y,N,N^FD${contador}^FS`);

  lines.push(`^PQ${copias},0,1,Y`);
  lines.push("^XZ");
  return `${lines.join("\n")}\n`;
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

export async function buildAndDownloadEtiquetaZpl({
  nomeLinha,
  corLinha,
  tamanho,
  prog,
  sku,
  contador,
  copias = 1,
  logoUrl = "/logo-rockstar-etiqueta.png",
}) {
  const logoW = pctX(22);
  const logoH = pctY(38.3);
  let logoGfa = null;
  try {
    logoGfa = await imageUrlToZplGfa(logoUrl, logoW, logoH);
  } catch {
    logoGfa = null;
  }

  const zpl = buildEtiquetaCaixaZpl({
    nomeLinha,
    corLinha,
    tamanho,
    prog,
    sku,
    contador,
    copias,
    logoGfa,
  });

  const safeSku = zplSafe(sku) || "etiqueta";
  const safeProg = zplSafe(prog).replace(/\s+/g, "") || "prog";
  downloadTextFile(`${safeSku}_${safeProg}.zpl`, zpl);
  return zpl;
}
