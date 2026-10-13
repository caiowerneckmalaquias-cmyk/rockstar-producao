import React, { useEffect, useMemo, useRef, useState } from "react";
import JsBarcode from "jsbarcode";
import { sizes } from "./constants/production";
import { buildBlingSku, normalizeBlingSigla } from "./constants/bling";

function RockStarLogoBw({ className = "" }) {
  return (
    <img
      src="/logo-rockstar-etiqueta.png"
      alt="Rock Star"
      className={`h-auto w-full object-contain grayscale contrast-150 brightness-90 ${className}`}
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
        fontOptions: "bold",
      });
    } catch {
      svgRef.current.innerHTML = "";
    }
  }, [safe, height, fontSize, width, displayValue]);

  return <svg ref={svgRef} className={className} />;
}

/** Preview visual da etiqueta de caixa 10×6 cm (P&B), fiel ao PDF TESTE ETIQUETA. */
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

  const nomeLinha = useMemo(() => {
    const n = String(nome || "").trim().toUpperCase();
    return n ? `${n} -` : "—";
  }, [nome]);

  const corLinha = useMemo(
    () => String(cor || "").trim().toUpperCase() || "—",
    [cor]
  );

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
            <p className="mt-1 text-sm text-slate-500">Layout alinhado ao PDF TESTE ETIQUETA (P&B · sem ZPL ainda).</p>
          </div>
          <div className="rounded-full border border-slate-200 bg-slate-50 px-3 py-1 text-[11px] font-semibold text-slate-600">
            Elgin L42 Pro · Code 128
          </div>
        </div>

        <div className="flex justify-center overflow-x-auto rounded-[20px] bg-[#e8e8e8] p-4 sm:p-10">
          {/* Etiqueta: proporção 10×6, sem borda grossa de “card” */}
          <div
            className="w-full max-w-[560px] bg-white text-black shadow-[0_8px_28px_rgba(0,0,0,0.18)]"
            style={{ aspectRatio: "10 / 6" }}
            aria-label="Preview da etiqueta 10 por 6 centímetros"
          >
            <div className="box-border flex h-full flex-col px-[3.2%] pb-[2.2%] pt-[2.8%]">
              {/* Topo: logo + título em 2 linhas */}
              <div className="flex items-start gap-[2.5%]">
                <div className="w-[16%] shrink-0 pt-[0.5%]">
                  <RockStarLogoBw />
                </div>
                <div className="min-w-0 flex-1 pt-[0.5%]">
                  <div className="text-[clamp(11px,2.35vw,15px)] font-bold uppercase leading-[1.15] tracking-tight text-black">
                    {nomeLinha}
                  </div>
                  <div className="text-[clamp(11px,2.35vw,15px)] font-bold uppercase leading-[1.15] tracking-tight text-black">
                    {corLinha}
                  </div>
                </div>
              </div>

              {/* Meio: tamanho grande + PROG empilhado */}
              <div className="mt-[1%] flex min-h-0 flex-1 items-center gap-2 pl-[2%]">
                <div className="flex-1 text-[clamp(56px,12vw,96px)] font-black leading-none tracking-tight tabular-nums text-black">
                  {tamanho}
                </div>
                <div className="shrink-0 self-center pr-[2%] text-right leading-none">
                  <div className="text-[clamp(11px,2.1vw,14px)] font-bold uppercase tracking-wide text-black">
                    PROG:
                  </div>
                  <div className="mt-1 text-[clamp(14px,2.6vw,18px)] font-black tabular-nums tracking-tight text-black">
                    {prog || "—"}
                  </div>
                </div>
              </div>

              {/* Base: SKU largo + contador compacto */}
              <div className="mt-auto grid grid-cols-[1.35fr_0.75fr] items-end gap-[3%] px-[0.5%]">
                <div className="min-w-0 overflow-hidden">
                  <BarcodeSvg
                    value={sku}
                    className="mx-auto block h-auto max-h-[72px] w-full"
                    height={56}
                    fontSize={13}
                    width={1.55}
                  />
                </div>
                <div className="min-w-0 overflow-hidden pb-[1%]">
                  <BarcodeSvg
                    value={contador}
                    className="mx-auto block h-auto max-h-[56px] w-[92%]"
                    height={42}
                    fontSize={11}
                    width={1.25}
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
