import React, { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import JsBarcode from "jsbarcode";
import { sizes, refAntigaDe } from "./constants/production";
import { buildBlingSku, normalizeBlingSigla } from "./constants/bling";
import {
  buildAndDownloadEtiquetaLoteZpl,
  expandLoteToLabels,
  readStoredContador,
  writeStoredContador,
} from "./etiquetaZpl";

const makeEmptyGrid = () => Object.fromEntries(sizes.map((s) => [s, 0]));

const normalizeKey = (value) =>
  String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim();

function nomeModeloDefault(ref) {
  const codigo = String(ref || "").trim().toUpperCase();
  const legado = refAntigaDe(codigo) || codigo;
  if (legado === "TNCV010" || codigo === "RCCB010") return "TENIS ROCK STAR CANO BAIXO";
  if (legado === "BTCV010" || codigo === "RCCA010") return "TENIS ROCK STAR CANO ALTO";
  if (legado === "CRVTNCV" || codigo === "RCCOUR01") return "TENIS ROCK STAR COURINO";
  return "TENIS ROCK STAR";
}

function RockStarLogo({ className = "" }) {
  return (
    <img
      src="/logo-rockstar-etiqueta.png"
      alt="Rock Star"
      className={`h-full w-full object-contain ${className}`}
      draggable={false}
    />
  );
}

function BarcodeSvg({
  value,
  className = "",
  height = 42,
  fontSize = 11,
  width = 1.4,
  displayValue = true,
}) {
  const svgRef = useRef(null);
  const safe = String(value || "").trim();

  useEffect(() => {
    if (!svgRef.current) return;
    if (!safe) {
      svgRef.current.innerHTML = "";
      return;
    }
    try {
      JsBarcode(svgRef.current, safe, {
        format: "CODE128",
        lineColor: "#000",
        background: "#fff",
        width,
        height,
        displayValue,
        fontSize,
        margin: 0,
        textMargin: 2,
        font: "Arial",
        textAlign: "center",
        fontOptions: "",
      });
      const intrinsicH = svgRef.current.height?.baseVal?.value || height + fontSize + 4;
      svgRef.current.setAttribute("height", String(intrinsicH));
      svgRef.current.removeAttribute("width");
      svgRef.current.style.width = "100%";
      svgRef.current.style.height = `${intrinsicH}px`;
      svgRef.current.style.maxHeight = "100%";
      svgRef.current.style.display = "block";
    } catch {
      svgRef.current.innerHTML = "";
    }
  }, [safe, height, fontSize, width, displayValue]);

  return <svg ref={svgRef} className={className} style={{ display: "block", width: "100%" }} />;
}

function EtiquetaFace({ nomeLinha, corLinha, tamanho, prog, sku, contador, variant = "preview" }) {
  const isPrint = variant === "print";
  return (
    <div
      className={
        isPrint
          ? "etiqueta-print-page relative overflow-hidden bg-white text-black"
          : "relative w-full max-w-[560px] overflow-hidden bg-white text-black shadow-[0_8px_28px_rgba(0,0,0,0.18)]"
      }
      style={
        isPrint
          ? {
              width: "100mm",
              height: "60mm",
              fontFamily: "Arial, Helvetica, sans-serif",
              WebkitPrintColorAdjust: "exact",
              printColorAdjust: "exact",
            }
          : { aspectRatio: "10 / 6", fontFamily: "Arial, Helvetica, sans-serif" }
      }
      aria-label="Preview da etiqueta 10 por 6 centímetros"
    >
      <div className="absolute left-[8%] top-[1.1%] h-[38.3%] w-[22%]">
        <RockStarLogo />
      </div>
      <div className="absolute left-[32.7%] top-[10.5%] flex h-[24%] w-[32.5%] flex-col items-center justify-center leading-none">
        <div className="text-[clamp(15px,3.1vw,20px)] font-bold uppercase tracking-wide text-black">PROG:</div>
        <div className="mt-[0.3em] text-[clamp(16px,3.3vw,22px)] font-bold tabular-nums tracking-tight text-black">
          {prog || "—"}
        </div>
      </div>
      <div className="absolute left-[72%] top-[3.3%] flex h-[33.3%] w-[22%] items-center justify-center bg-black">
        <span className="text-[clamp(40px,9.5vw,72px)] font-bold leading-none tabular-nums tracking-tight text-white">
          {tamanho}
        </span>
      </div>
      <div className="absolute inset-x-[2%] top-[39%] bottom-[34%] flex flex-col items-center justify-center text-center">
        <div className="text-[clamp(15px,3.15vw,21px)] font-bold uppercase leading-tight tracking-tight text-black">
          {nomeLinha}
        </div>
        <div className="mt-[0.12em] text-[clamp(15px,3.15vw,21px)] font-bold uppercase leading-tight tracking-tight text-black">
          {corLinha}
        </div>
      </div>
      <div className="absolute inset-x-[10%] bottom-[5%] flex h-[29%] items-end justify-between gap-[10%]">
        <div className="flex w-[54%] shrink-0 items-end justify-center">
          <BarcodeSvg value={sku} height={50} fontSize={13} width={1.5} />
        </div>
        <div className="flex w-[32%] shrink-0 items-end justify-center">
          <BarcodeSvg value={contador} height={50} fontSize={13} width={1.5} />
        </div>
      </div>
    </div>
  );
}

/**
 * Aba Etiquetas — fluxo estilo Pesponto: programação, ref/cor, grade, lote e ZPL sequencial.
 */
export default function EtiquetaCaixaPreview({
  refsUnicas = [],
  coresPorRef = new Map(),
  getBlingSigla = () => "",
  startPrintWithTarget,
}) {
  const firstRef = refsUnicas[0] || "RCCA010";
  const firstCores = coresPorRef.get?.(firstRef) || [];
  const firstCor = firstCores[0] || "BRANCO";

  const [programacao, setProgramacao] = useState("");
  const [ref, setRef] = useState(firstRef);
  const [cor, setCor] = useState(firstCor);
  const [nome, setNome] = useState(nomeModeloDefault(firstRef));
  const [grid, setGrid] = useState(makeEmptyGrid);
  const [contadorInicial, setContadorInicial] = useState(() => readStoredContador());
  const [lote, setLote] = useState([]);
  const [previewSize, setPreviewSize] = useState(34);
  const [zplBusy, setZplBusy] = useState(false);
  const [printBusy, setPrintBusy] = useState(false);
  const [printLabels, setPrintLabels] = useState([]);
  const [error, setError] = useState("");
  const pendingNextContadorRef = useRef("");

  const sigla = useMemo(
    () => normalizeBlingSigla(getBlingSigla(ref, cor) || ""),
    [getBlingSigla, ref, cor]
  );

  const coresDaRef = coresPorRef.get?.(ref) || [];

  useEffect(() => {
    if (!refsUnicas.length) return;
    if (!refsUnicas.includes(ref)) {
      const nextRef = refsUnicas[0];
      const cores = coresPorRef.get?.(nextRef) || [];
      setRef(nextRef);
      setCor(cores[0] || "");
      setNome(nomeModeloDefault(nextRef));
    }
  }, [refsUnicas, ref, coresPorRef]);

  useEffect(() => {
    if (!coresDaRef.length) return;
    if (!coresDaRef.some((c) => normalizeKey(c) === normalizeKey(cor))) {
      setCor(coresDaRef[0]);
    }
  }, [coresDaRef, cor]);

  const totalForm = useMemo(
    () => sizes.reduce((acc, s) => acc + (Number(grid[s]) || 0), 0),
    [grid]
  );

  const totalLote = useMemo(
    () =>
      lote.reduce(
        (acc, linha) =>
          acc + sizes.reduce((sum, s) => sum + (Number(linha.grid?.[s]) || 0), 0),
        0
      ),
    [lote]
  );

  const previewSku = useMemo(
    () => buildBlingSku(ref, sigla, previewSize),
    [ref, sigla, previewSize]
  );

  const nomeLinha = useMemo(() => {
    const n = String(nome || "").trim().toUpperCase();
    return n ? `${n} -` : "—";
  }, [nome]);

  const corLinha = useMemo(
    () => String(cor || "").trim().toUpperCase() || "—",
    [cor]
  );

  const adicionarAoLote = () => {
    setError("");
    if (!String(programacao || "").trim()) {
      setError("Informe o nome da programação.");
      return;
    }
    if (!ref || !cor) {
      setError("Selecione referência e cor.");
      return;
    }
    if (!sigla) {
      setError(`Sem sigla Bling para ${ref} • ${cor}. Cadastre no Controle Geral.`);
      return;
    }
    if (!totalForm) {
      setError("Preencha a grade (quantidade de etiquetas por numeração).");
      return;
    }

    setLote((curr) => [
      ...curr,
      {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        programacao: String(programacao).trim(),
        ref,
        cor,
        sigla,
        nome: String(nome || "").trim().toUpperCase() || nomeModeloDefault(ref),
        grid: { ...grid },
        total: totalForm,
      },
    ]);
    setGrid(makeEmptyGrid());
  };

  const removerDoLote = (id) => {
    setLote((curr) => curr.filter((item) => item.id !== id));
  };

  const limparLote = () => {
    setLote([]);
    setError("");
  };

  const baixarZplLote = async () => {
    setError("");
    if (!lote.length) {
      setError("Adicione pelo menos uma linha ao lote.");
      return;
    }
    setZplBusy(true);
    try {
      const result = await buildAndDownloadEtiquetaLoteZpl({
        lote,
        contadorInicial,
        filenameBase: String(programacao || lote[0]?.programacao || "etiquetas").trim(),
      });
      setContadorInicial(result.nextContador);
    } catch (err) {
      setError(err?.message || "Falha ao gerar o ZPL.");
    } finally {
      setZplBusy(false);
    }
  };

  const imprimirLote = () => {
    setError("");
    if (!lote.length) {
      setError("Adicione pelo menos uma linha ao lote.");
      return;
    }
    const { labels, nextContador, total } = expandLoteToLabels(lote, contadorInicial);
    if (!total) {
      setError("Nenhuma etiqueta no lote (grade vazia ou sem sigla Bling).");
      return;
    }
    pendingNextContadorRef.current = nextContador;
    setPrintBusy(true);
    setPrintLabels(labels);
  };

  useEffect(() => {
    if (!printLabels.length) return undefined;

    let finished = false;
    const finishPrint = () => {
      if (finished) return;
      finished = true;
      const next = pendingNextContadorRef.current;
      if (next) {
        writeStoredContador(next);
        setContadorInicial(next);
      }
      pendingNextContadorRef.current = "";
      setPrintLabels([]);
      setPrintBusy(false);
    };

    // Aguarda barcodes (JsBarcode) montarem antes do diálogo de impressão.
    const timer = window.setTimeout(() => {
      window.addEventListener("afterprint", finishPrint, { once: true });
      if (typeof startPrintWithTarget === "function") {
        startPrintWithTarget("etiqueta");
      } else {
        document.body.classList.add("print-target-etiqueta");
        const clearClass = () => document.body.classList.remove("print-target-etiqueta");
        window.addEventListener("afterprint", clearClass, { once: true });
        window.print();
      }
    }, 350);

    return () => {
      window.clearTimeout(timer);
      window.removeEventListener("afterprint", finishPrint);
    };
  }, [printLabels, startPrintWithTarget]);

  const previewContador = useMemo(() => {
    const { labels } = expandLoteToLabels(
      [
        {
          programacao,
          ref,
          cor,
          sigla,
          nome,
          grid: { ...makeEmptyGrid(), [previewSize]: Math.max(1, Number(grid[previewSize]) || 1) },
        },
      ],
      contadorInicial
    );
    return labels[0]?.contador || contadorInicial;
  }, [programacao, ref, cor, sigla, nome, previewSize, grid, contadorInicial]);

  return (
    <div className="grid grid-cols-1 gap-6 xl:grid-cols-[440px_1fr] items-start">
      <div className="space-y-4">
        <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-sm space-y-4">
          <div>
            <div className="text-sm font-semibold text-slate-800">Montar lote de etiquetas</div>
            <p className="mt-1 text-xs text-slate-500">
              Grade = quantidade de etiquetas por numeração. SKU se repete; contador sobe 1 a 1.
            </p>
          </div>

          <label className="text-sm font-medium block">
            Programação
            <input
              value={programacao}
              onChange={(e) => setProgramacao(e.target.value)}
              placeholder="Ex.: FICHA 220 / 26-215"
              className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
            />
          </label>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <label className="text-sm font-medium block">
              Referência
              <select
                value={ref}
                onChange={(e) => {
                  const nextRef = e.target.value;
                  const cores = coresPorRef.get?.(nextRef) || [];
                  setRef(nextRef);
                  setCor(cores[0] || "");
                  setNome(nomeModeloDefault(nextRef));
                }}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                {(refsUnicas.length ? refsUnicas : [ref]).map((r) => (
                  <option key={r} value={r}>
                    {r}
                  </option>
                ))}
              </select>
            </label>
            <label className="text-sm font-medium block">
              Cor
              <select
                value={cor}
                onChange={(e) => setCor(e.target.value)}
                className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
              >
                {(coresDaRef.length ? coresDaRef : [cor]).map((c) => (
                  <option key={`${ref}__${c}`} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <label className="text-sm font-medium block">
            Nome / modelo
            <input
              value={nome}
              onChange={(e) => setNome(e.target.value.toUpperCase())}
              className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 text-sm"
            />
          </label>

          <div className="rounded-2xl border border-slate-200 bg-slate-50 px-3 py-3 text-sm">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Sigla Bling</span>
              <span className={`font-mono font-bold ${sigla ? "text-slate-900" : "text-red-600"}`}>
                {sigla || "não cadastrada"}
              </span>
            </div>
            {sigla ? (
              <div className="mt-1 text-xs text-slate-500">
                Ex. SKU {previewSize}: <span className="font-mono font-semibold text-slate-800">{previewSku}</span>
              </div>
            ) : null}
          </div>

          <div>
            <div className="mb-2 flex items-center justify-between gap-2">
              <div className="text-sm font-medium">Grade completa</div>
              <div className="text-xs font-semibold text-slate-500">Total: {totalForm}</div>
            </div>
            <div className="grid grid-cols-4 gap-2">
              {sizes.map((size) => (
                <div key={size} className="rounded-xl border border-slate-200 bg-slate-50 p-2">
                  <button
                    type="button"
                    onClick={() => setPreviewSize(size)}
                    className={`mx-auto block text-xs text-center ${
                      previewSize === size ? "font-black text-[#8B1E2D]" : "text-slate-500"
                    }`}
                    title="Usar no preview"
                  >
                    {size}
                  </button>
                  <input
                    value={grid[size]}
                    onChange={(e) =>
                      setGrid((curr) => ({
                        ...curr,
                        [size]: Number(e.target.value) || 0,
                      }))
                    }
                    className="mt-1 w-full rounded-lg border border-slate-200 p-1.5 text-center text-sm"
                  />
                </div>
              ))}
            </div>
          </div>

          <label className="text-sm font-medium block">
            Contador inicial (direita)
            <input
              value={contadorInicial}
              onChange={(e) => setContadorInicial(e.target.value.replace(/[^\d]/g, ""))}
              className="mt-2 w-full rounded-2xl border border-slate-200 bg-slate-50 p-3 font-mono text-sm"
              placeholder="0487692"
            />
            <span className="mt-1 block text-xs text-slate-500">
              Cada etiqueta do lote recebe o próximo número. Após baixar, o próximo fica salvo.
            </span>
          </label>

          <button
            type="button"
            onClick={adicionarAoLote}
            className="w-full rounded-2xl bg-[#8B1E2D] px-4 py-3 text-sm font-bold text-white shadow-sm transition hover:bg-[#6f1824]"
          >
            Adicionar ao lote
          </button>
          {error ? <p className="text-sm font-semibold text-red-600">{error}</p> : null}
        </div>

        <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-sm space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div>
              <div className="text-sm font-semibold text-slate-800">Lote para imprimir</div>
              <div className="text-xs text-slate-500">
                {lote.length} linha(s) · {totalLote} etiqueta(s)
              </div>
            </div>
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                onClick={limparLote}
                disabled={!lote.length || printBusy || zplBusy}
                className="rounded-xl border border-slate-200 px-3 py-1.5 text-xs font-semibold text-slate-600 disabled:opacity-40"
              >
                Limpar
              </button>
              <button
                type="button"
                onClick={imprimirLote}
                disabled={printBusy || zplBusy || !lote.length}
                className="rounded-xl bg-[#8B1E2D] px-3 py-1.5 text-xs font-bold text-white disabled:opacity-40"
              >
                {printBusy ? "Preparando…" : "Imprimir"}
              </button>
              <button
                type="button"
                onClick={baixarZplLote}
                disabled={zplBusy || printBusy || !lote.length}
                className="rounded-xl bg-[#0F172A] px-3 py-1.5 text-xs font-bold text-white disabled:opacity-40"
              >
                {zplBusy ? "Gerando…" : "Baixar ZPL"}
              </button>
            </div>
          </div>

          {lote.length === 0 ? (
            <p className="rounded-2xl border border-dashed border-slate-200 bg-slate-50 px-4 py-6 text-center text-sm text-slate-500">
              Nenhuma linha no lote. Preencha a grade e clique em Adicionar.
            </p>
          ) : (
            <div className="space-y-2">
              {lote.map((item) => (
                <div
                  key={item.id}
                  className="rounded-2xl border border-slate-200 bg-slate-50 px-3 py-3 text-sm"
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <div className="font-bold text-slate-900">
                        {item.ref} • {item.cor}
                      </div>
                      <div className="text-xs text-slate-500">
                        {item.programacao} · sigla {item.sigla} · {item.total} etq
                      </div>
                      <div className="mt-1 flex flex-wrap gap-1">
                        {sizes
                          .filter((s) => Number(item.grid?.[s]) > 0)
                          .map((s) => (
                            <span
                              key={`${item.id}-${s}`}
                              className="rounded-md bg-sky-100 px-1.5 py-0.5 text-[11px] font-semibold text-sky-900"
                            >
                              {s} → {item.grid[s]}
                            </span>
                          ))}
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => removerDoLote(item.id)}
                      className="shrink-0 rounded-lg bg-red-100 px-2 py-1 text-xs font-bold text-red-700"
                    >
                      Excluir
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-4 flex flex-wrap items-end justify-between gap-2">
          <div>
            <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
              Preview 10 × 6 cm
            </div>
            <p className="mt-1 text-sm text-slate-500">
              Clique num tamanho na grade para atualizar o preview. ZPL do lote usa contador sequencial.
            </p>
          </div>
          <div className="rounded-full border border-slate-200 bg-slate-50 px-3 py-1 text-[11px] font-semibold text-slate-600">
            Elgin L42 Pro · Code 128 · 203 dpi
          </div>
        </div>

        <div className="flex justify-center overflow-x-auto rounded-[20px] bg-[#e8e8e8] p-4 sm:p-10">
          <EtiquetaFace
            nomeLinha={nomeLinha}
            corLinha={corLinha}
            tamanho={previewSize}
            prog={programacao || "—"}
            sku={previewSku || "—"}
            contador={previewContador}
          />
        </div>
      </section>

      {typeof document !== "undefined" &&
        createPortal(
          <div id="print-etiqueta-root" className="hidden">
            {printLabels.map((label, idx) => (
              <div
                key={`etq-print-${label.contador}-${label.sku}-${idx}`}
                className={idx < printLabels.length - 1 ? "etiqueta-print-break" : undefined}
              >
                <EtiquetaFace
                  variant="print"
                  nomeLinha={label.nomeLinha}
                  corLinha={label.corLinha}
                  tamanho={label.tamanho}
                  prog={label.prog}
                  sku={label.sku}
                  contador={label.contador}
                />
              </div>
            ))}
          </div>,
          document.body
        )}
    </div>
  );
}
