import React, { useEffect, useMemo, useRef, useState } from "react";
import JsBarcode from "jsbarcode";
import { sizes } from "./constants/production";
import { buildBlingSku, normalizeBlingSigla } from "./constants/bling";

function RockStarLogoBw({ className = "" }) {
  return (
    <div className={`relative flex flex-col items-center justify-center leading-none ${className}`}>
      <svg
        viewBox="0 0 64 64"
        className="pointer-events-none absolute inset-0 m-auto h-[90%] w-[90%] opacity-90"
        aria-hidden
      >
        <polygon
          points="32,4 39,24 60,24 43,37 50,58 32,45 14,58 21,37 4,24 25,24"
          fill="none"
          stroke="#000"
          strokeWidth="2.2"
        />
      </svg>
      <div className="relative z-[1] border-y-2 border-black px-1 py-0.5 text-center">
        <div className="text-[10px] font-black tracking-[0.08em] text-black sm:text-xs">ROCK</div>
        <div className="text-[10px] font-black tracking-[0.08em] text-black sm:text-xs">STAR</div>
      </div>
    </div>
  );
}

function BarcodeSvg({ value, className = "", height = 42, fontSize = 11, displayValue = true }) {
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
        width: 1.4,
        height,
        displayValue,
        fontSize,
        margin: 0,
        textMargin: 2,
      });
    } catch {
      svgRef.current.innerHTML = "";
    }
  }, [safe, height, fontSize, displayValue]);

  return <svg ref={svgRef} className={className} />;
}

/** Preview visual da etiqueta de caixa 10×6 cm (P&B), sem ZPL. */
export default function EtiquetaCaixaPreview() {
  const [ref, setRef] = useState("RCCA010");
  const [sigla, setSigla] = useState("BC");
  const [nome, setNome] = useState("TENIS ROCK STAR CANO ALTO");
  const [cor, setCor] = useState("BRANCO");
  const [tamanho, setTamanho] = useState(34);
  const [prog, setProg] = useState("26-215");
  const [contador, setContador] = useState("0487692");

  const sku = useMemo(
    () => buildBlingSku(ref, sigla, tamanho),
    [ref, sigla, tamanho]
  );

  const tituloLinha = useMemo(() => {
    const n = String(nome || "").trim().toUpperCase();
    const c = String(cor || "").trim().toUpperCase();
    if (n && c) return `${n} - ${c}`;
    return n || c || "—";
  }, [nome, cor]);

  const inputClass =
    "w-full rounded-2xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-semibold text-slate-800 outline-none focus:border-[#8B1E2D] focus:ring-2 focus:ring-[#8B1E2D]/15";

  return (
    <div className="grid grid-cols-1 gap-6 xl:grid-cols-[minmax(280px,360px)_1fr]">
      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm space-y-4">
        <div>
          <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">Dados da etiqueta</div>
          <p className="mt-1 text-sm text-slate-500">Altere os campos — o preview atualiza na hora.</p>
        </div>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Referência</span>
          <input
            className={inputClass}
            value={ref}
            onChange={(e) => setRef(e.target.value.toUpperCase())}
            placeholder="RCCA010"
          />
        </label>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Sigla (cor Bling)</span>
          <input
            className={inputClass}
            value={sigla}
            onChange={(e) => setSigla(normalizeBlingSigla(e.target.value))}
            placeholder="BC"
          />
        </label>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Nome / modelo</span>
          <input
            className={inputClass}
            value={nome}
            onChange={(e) => setNome(e.target.value.toUpperCase())}
            placeholder="TENIS ROCK STAR CANO ALTO"
          />
        </label>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Cor (texto)</span>
          <input
            className={inputClass}
            value={cor}
            onChange={(e) => setCor(e.target.value.toUpperCase())}
            placeholder="BRANCO"
          />
        </label>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Tamanho</span>
          <select
            className={inputClass}
            value={tamanho}
            onChange={(e) => setTamanho(Number(e.target.value))}
          >
            {sizes.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
        </label>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">PROG</span>
          <input
            className={inputClass}
            value={prog}
            onChange={(e) => setProg(e.target.value)}
            placeholder="26-215"
          />
        </label>

        <label className="block space-y-1">
          <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Contador (direita)</span>
          <input
            className={inputClass}
            value={contador}
            onChange={(e) => setContador(e.target.value.replace(/[^\d]/g, ""))}
            placeholder="0487692"
          />
        </label>

        <div className="rounded-2xl border border-slate-200 bg-slate-50 px-3 py-3">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">SKU (esteira)</div>
          <div className="mt-1 font-mono text-lg font-black tracking-tight text-slate-900">
            {sku || "—"}
          </div>
        </div>
      </section>

      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-4 flex flex-wrap items-end justify-between gap-2">
          <div>
            <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">Preview 10 × 6 cm</div>
            <p className="mt-1 text-sm text-slate-500">Preto e branco — layout da etiqueta de caixa (sem ZPL ainda).</p>
          </div>
          <div className="rounded-full border border-slate-200 bg-slate-50 px-3 py-1 text-[11px] font-semibold text-slate-600">
            Elgin L42 Pro · Code 128
          </div>
        </div>

        <div className="flex justify-center overflow-x-auto rounded-[20px] bg-slate-100/80 p-4 sm:p-8">
          <div
            className="w-full max-w-[520px] border-2 border-black bg-white text-black shadow-md"
            style={{ aspectRatio: "10 / 6" }}
            aria-label="Preview da etiqueta 10 por 6 centímetros"
          >
            <div className="flex h-full flex-col px-[3.5%] py-[3%]">
              <div className="flex min-h-0 flex-1 gap-2">
                <div className="flex w-[22%] shrink-0 items-start justify-center pt-1">
                  <RockStarLogoBw className="w-full max-w-[72px]" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="text-[clamp(9px,2.1vw,13px)] font-bold uppercase leading-tight tracking-tight">
                    {tituloLinha}
                  </div>
                  <div className="mt-[2%] flex items-end justify-between gap-2">
                    <div className="text-[clamp(42px,9vw,72px)] font-black leading-none tracking-tight tabular-nums">
                      {tamanho}
                    </div>
                    <div className="pb-1 text-right text-[clamp(10px,2vw,13px)] font-bold uppercase leading-tight">
                      <span className="text-slate-700">PROG:</span>{" "}
                      <span className="tabular-nums text-black">{prog || "—"}</span>
                    </div>
                  </div>
                </div>
              </div>

              <div className="mt-auto grid grid-cols-[1.15fr_0.85fr] items-end gap-3 border-t border-black/10 pt-[2%]">
                <div className="min-w-0 overflow-hidden">
                  <BarcodeSvg
                    value={sku}
                    className="mx-auto h-auto max-h-[58px] w-full max-w-full"
                    height={48}
                    fontSize={12}
                  />
                </div>
                <div className="min-w-0 overflow-hidden">
                  <BarcodeSvg
                    value={contador}
                    className="mx-auto h-auto max-h-[52px] w-full max-w-full"
                    height={40}
                    fontSize={10}
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}
