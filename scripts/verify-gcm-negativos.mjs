/**
 * Regressão: negativo no meio da linha GCM vira 0 só naquela numeração.
 * Uso: node scripts/verify-gcm-negativos.mjs [caminho.xls]
 */
import path from "path";
import { fileURLToPath } from "url";
import XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const sizes = [33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44];

function clampGcmQty(n) {
  const v = Number(n);
  if (!Number.isFinite(v)) return 0;
  return Math.max(0, v);
}

function lerQuantidadesPorColuna(linhaCabecalho, linhaQtd) {
  const data = Object.fromEntries(sizes.map((s) => [s, 0]));
  const cab = (linhaCabecalho || []).slice(1);
  const qtd = (linhaQtd || []).slice(1);
  cab.forEach((cell, idx) => {
    const size = Number(cell);
    if (!sizes.includes(size)) return;
    data[size] = clampGcmQty(qtd[idx]);
  });
  return data;
}

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
    if (!cabecalho.includes("-")) continue;
    if (cabecalho.startsWith("ESTOQUE") || cabecalho.startsWith("OVERLOQUE")) continue;
    if (!linha1.slice(1).some((v) => sizes.includes(Number(v)))) continue;

    const partes = linha1[0].split("-");
    if (partes.length < 2) continue;
    const ref = toText(partes[0]).toUpperCase();
    const cor = toText(partes.slice(1).join("-")).toUpperCase();

    const prox = Array.isArray(rowsSheet[i + 1]) ? rowsSheet[i + 1] : [];
    const linha2 = prox.map(toText);
    if (!upper(linha2[0]).startsWith("ESTOQUE")) continue;

    const data = lerQuantidadesPorColuna(linha1, linha2);
    const prox2 = Array.isArray(rowsSheet[i + 2]) ? rowsSheet[i + 2] : [];
    const linha3 = prox2.map(toText);
    const temOverloque = upper(linha3[0]).startsWith("OVERLOQUE");
    const dataEst = temOverloque
      ? lerQuantidadesPorColuna(linha1, linha3)
      : Object.fromEntries(sizes.map((s) => [s, 0]));

    resultado.push({ ref, cor, data, dataEst, linha2Raw: prox, linha3Raw: prox2, linha1 });
  }
  return resultado;
}

/** Extrai qty bruta (sem clamp) na coluna do tamanho — para achar negativos no Excel. */
function qtyBrutaNaColuna(linhaCabecalho, linhaQtd, sizeAlvo) {
  const cab = (linhaCabecalho || []).slice(1);
  const qtd = (linhaQtd || []).slice(1);
  for (let idx = 0; idx < cab.length; idx += 1) {
    if (Number(cab[idx]) !== sizeAlvo) continue;
    const n = Number(qtd[idx]);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

const filePath =
  process.argv[2] ||
  path.join(__dirname, "fixtures", "COURINO_ALL_STAR.xls");

const wb = XLSX.readFile(filePath);
const sheet = wb.Sheets[wb.SheetNames[0]];
const parsed = parseGcmSheet(sheet);

assert(parsed.length > 0, "Nenhum bloco GCM encontrado");

let encontradosNegativos = 0;

for (const item of parsed) {
  for (const size of sizes) {
    const rawPa = qtyBrutaNaColuna(item.linha1, item.linha2Raw, size);
    if (rawPa != null && rawPa < 0) {
      encontradosNegativos += 1;
      assert(
        item.data[size] === 0,
        `${item.cor} PA ${size}: esperado 0 após clamp, veio ${item.data[size]}`
      );
      const depois = sizes.filter((s) => s > size);
      const algumPositivoNoExcel = depois.some((s) => {
        const r = qtyBrutaNaColuna(item.linha1, item.linha2Raw, s);
        return r != null && r > 0;
      });
      if (algumPositivoNoExcel) {
        const todosZerados = depois.every((s) => (item.data[s] || 0) === 0);
        assert(
          !todosZerados,
          `${item.cor} PA: após negativo no ${size}, restante da linha foi zerado`
        );
        for (const s of depois) {
          const r = qtyBrutaNaColuna(item.linha1, item.linha2Raw, s);
          if (r == null || r < 0) continue;
          assert(
            item.data[s] === clampGcmQty(r),
            `${item.cor} PA ${s}: esperado ${clampGcmQty(r)}, veio ${item.data[s]}`
          );
        }
      }
      console.log(`OK clamp PA: ${item.cor} size ${size} (${rawPa} → 0), restante intacto`);
    }

    const rawEst = qtyBrutaNaColuna(item.linha1, item.linha3Raw, size);
    if (rawEst != null && rawEst < 0) {
      encontradosNegativos += 1;
      assert(
        item.dataEst[size] === 0,
        `${item.cor} EST ${size}: esperado 0 após clamp, veio ${item.dataEst[size]}`
      );
      const depois = sizes.filter((s) => s > size);
      const algumPositivoNoExcel = depois.some((s) => {
        const r = qtyBrutaNaColuna(item.linha1, item.linha3Raw, s);
        return r != null && r > 0;
      });
      if (algumPositivoNoExcel) {
        const todosZerados = depois.every((s) => (item.dataEst[s] || 0) === 0);
        assert(
          !todosZerados,
          `${item.cor} EST: após negativo no ${size}, restante da linha foi zerado`
        );
        for (const s of depois) {
          const r = qtyBrutaNaColuna(item.linha1, item.linha3Raw, s);
          if (r == null || r < 0) continue;
          assert(
            item.dataEst[s] === clampGcmQty(r),
            `${item.cor} EST ${s}: esperado ${clampGcmQty(r)}, veio ${item.dataEst[s]}`
          );
        }
      }
      console.log(`OK clamp EST: ${item.cor} size ${size} (${rawEst} → 0), restante intacto`);
    }
  }
}

// Texto: "-7" não pode virar 7
const extrairNumeros = (texto) =>
  (String(texto || "").match(/-?\d+/g) || []).map(Number);
const nums = extrairNumeros("ESTOQUE 32 25 14 -7 15 28");
assert(nums.includes(-7), `extrairNumeros deveria incluir -7, veio ${JSON.stringify(nums)}`);
assert(!nums.includes(7) || nums.indexOf(-7) >= 0, "sinal perdido em -7");
assert(clampGcmQty(-7) === 0, "clampGcmQty(-7) deveria ser 0");
assert(clampGcmQty(15) === 15, "clampGcmQty(15) deveria ser 15");

assert(
  encontradosNegativos >= 1,
  "Fixture deveria ter ao menos um negativo mid-line"
);

console.log(
  `PASS: ${encontradosNegativos} negativo(s) clampados só na célula; ${parsed.length} bloco(s) em ${filePath}`
);
