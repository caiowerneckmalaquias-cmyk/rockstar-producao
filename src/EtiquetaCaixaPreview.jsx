import React, { useEffect, useMemo, useRef, useState } from "react";
import JsBarcode from "jsbarcode";
import { sizes } from "./constants/production";
import { buildBlingSku, normalizeBlingSigla } from "./constants/bling";

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
    } catch {
      svgRef.current.innerHTML = "";
    }
  }, [safe, height, fontSize, width, displayValue]);

  return <svg ref={svgRef} className={className} />;
}

/** Preview visual da etiqueta de caixa 10×6 cm, geometria do PDF TESTE ETIQUETA. */
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
            <p className="mt-1 text-sm text-slate-500">Layout idêntico ao PDF TESTE ETIQUETA (sem ZPL ainda).</p>
          </div>
          <div className="rounded-full border border-slate-200 bg-slate-50 px-3 py-1 text-[11px] font-semibold text-slate-600">
            Elgin L42 Pro · Code 128
          </div>
        </div>

        <div className="flex justify-center overflow-x-auto rounded-[20px] bg-[#e8e8e8] p-4 sm:p-10">
          {/*
            Geometria do PDF MediaBox 283.50 × 170.25 pt (10×6 cm).
            Posições em % da etiqueta (origem topo-esquerda).
          */}
          <div
            className="relative w-full max-w-[560px] overflow-hidden bg-white text-black shadow-[0_8px_28px_rgba(0,0,0,0.18)]"
            style={{ aspectRatio: "10 / 6", fontFamily: "Arial, Helvetica, sans-serif" }}
            aria-label="Preview da etiqueta 10 por 6 centímetros"
          >
            {/* Logo — PDF: x8% y1% w22% h38% */}
            <div className="absolute left-[8%] top-[1.1%] h-[38.3%] w-[22%]">
              <RockStarLogo />
            </div>

            {/* PROG empilhado — faixa central-topo do PDF */}
            <div className="absolute left-[32.7%] top-[10.5%] flex h-[24%] w-[32.5%] flex-col items-center justify-center leading-none">
              <div className="text-[clamp(15px,3.1vw,20px)] font-bold uppercase tracking-wide text-black">
                PROG:
              </div>
              <div className="mt-[0.3em] text-[clamp(16px,3.3vw,22px)] font-bold tabular-nums tracking-tight text-black">
                {prog || "—"}
              </div>
            </div>

            {/* Tamanho branco em caixa preta — PDF: x72% y3% w22% h33% */}
            <div className="absolute left-[72%] top-[3.3%] flex h-[33.3%] w-[22%] items-center justify-center bg-black">
              <span className="text-[clamp(40px,9.5vw,72px)] font-bold leading-none tabular-nums tracking-tight text-white">
                {tamanho}
              </span>
            </div>

            {/* Nome + cor — centralizado na faixa entre topo (logo/PROG/nº) e barcodes */}
            <div className="absolute left-[2%] right-[2%] top-[38%] bottom-[34%] flex flex-col items-center justify-center px-[2%] text-center">
              <div className="text-[clamp(15px,3.15vw,21px)] font-bold uppercase leading-tight tracking-tight text-black">
                {nomeLinha}
              </div>
              <div className="mt-[0.12em] text-[clamp(15px,3.15vw,21px)] font-bold uppercase leading-tight tracking-tight text-black">
                {corLinha}
              </div>
            </div>

            {/* Barcodes na mesma altura / mesma linha de base */}
            <div className="absolute bottom-[5%] left-[10%] right-[9.4%] flex h-[30%] items-end justify-between gap-[8%]">
              <div className="flex h-full w-[55%] items-end justify-center overflow-hidden">
                <BarcodeSvg
                  value={sku}
                  className="block h-auto max-h-full w-full"
                  height={52}
                  fontSize={13}
                  width={1.55}
                />
              </div>
              <div className="flex h-full w-[32%] items-end justify-center overflow-hidden">
                <BarcodeSvg
                  value={contador}
                  className="block h-auto max-h-full w-full"
                  height={52}
                  fontSize={13}
                  width={1.35}
                />
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}
