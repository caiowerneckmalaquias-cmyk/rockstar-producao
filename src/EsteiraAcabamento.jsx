import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  liberarBipPorContador,
  normalizeContadorScan,
  reservarBipPorContador,
} from "./etiquetasCaixaDb";

/**
 * Esteira / Acabamento — bip do contador sequencial (1 bip = 1 par).
 * Anti-duplicidade no Supabase; confirmar lote aplica M→PA + Bling (via callback).
 */
export default function EsteiraAcabamento({ onConfirmLote }) {
  const inputRef = useRef(null);
  const [scan, setScan] = useState("");
  const [busy, setBusy] = useState(false);
  const [confirmBusy, setConfirmBusy] = useState(false);
  const [feedback, setFeedback] = useState({ type: "", text: "" });
  const [sessao, setSessao] = useState([]);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  const totalPares = sessao.length;

  const agrupado = useMemo(() => {
    const map = new Map();
    sessao.forEach((item) => {
      const key = `${item.ref}__${item.cor}__${item.tamanho}`;
      const prev = map.get(key) || {
        ref: item.ref,
        cor: item.cor,
        tamanho: item.tamanho,
        sku: item.sku,
        qtd: 0,
      };
      prev.qtd += 1;
      map.set(key, prev);
    });
    return [...map.values()].sort((a, b) =>
      `${a.ref}${a.cor}${a.tamanho}`.localeCompare(`${b.ref}${b.cor}${b.tamanho}`)
    );
  }, [sessao]);

  const focar = () => {
    window.setTimeout(() => inputRef.current?.focus(), 30);
  };

  const processarBip = async (raw) => {
    const codigo = normalizeContadorScan(raw);
    if (!codigo) {
      setFeedback({ type: "err", text: "Bipe o código sequencial da etiqueta (direita)." });
      focar();
      return;
    }
    if (busy) return;
    setBusy(true);
    setFeedback({ type: "", text: "" });
    try {
      if (sessao.some((s) => normalizeContadorScan(s.contador) === codigo || s.contador_n === Number(codigo))) {
        setFeedback({ type: "err", text: `Caixa ${codigo} já está neste lote.` });
        return;
      }

      const result = await reservarBipPorContador(codigo);
      if (!result.ok) {
        setFeedback({ type: "err", text: result.error || "Falha no bip." });
        return;
      }

      const row = result.row;
      setSessao((curr) => [
        {
          contador: row.contador,
          contador_n: row.contador_n,
          sku: row.sku,
          ref: row.ref,
          cor: row.cor,
          tamanho: Number(row.tamanho),
          programacao: row.programacao || "",
          bipadoEm: row.bipado_em,
        },
        ...curr,
      ]);
      setFeedback({
        type: "ok",
        text: `OK ${row.contador} · ${row.sku} · ${row.ref} • ${row.cor} · nº ${row.tamanho}`,
      });
    } finally {
      setScan("");
      setBusy(false);
      focar();
    }
  };

  const onKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      processarBip(scan);
    }
  };

  const desfazerUltimo = async () => {
    if (!sessao.length || busy || confirmBusy) return;
    const [ultimo, ...rest] = sessao;
    setBusy(true);
    try {
      const result = await liberarBipPorContador(ultimo.contador);
      if (!result.ok) {
        setFeedback({ type: "err", text: result.error || "Não foi possível desfazer." });
        return;
      }
      setSessao(rest);
      setFeedback({ type: "ok", text: `Desfeito ${ultimo.contador}.` });
    } finally {
      setBusy(false);
      focar();
    }
  };

  const limparSessaoLocal = async () => {
    if (!sessao.length || confirmBusy) return;
    if (!window.confirm(`Liberar ${sessao.length} bip(s) sem confirmar estoque?`)) return;
    setBusy(true);
    try {
      for (const item of sessao) {
        await liberarBipPorContador(item.contador);
      }
      setSessao([]);
      setFeedback({ type: "ok", text: "Lote local liberado." });
    } finally {
      setBusy(false);
      focar();
    }
  };

  const confirmar = async () => {
    if (!sessao.length || confirmBusy) return;
    if (typeof onConfirmLote !== "function") {
      setFeedback({ type: "err", text: "Confirmação de estoque não configurada." });
      return;
    }
    if (
      !window.confirm(
        `Confirmar ${sessao.length} par(es)?\nIsso faz M → PA no Rock Star e sobe esses SKUs no Bling.`
      )
    ) {
      return;
    }
    setConfirmBusy(true);
    setFeedback({ type: "", text: "" });
    try {
      const result = await onConfirmLote(sessao);
      if (!result?.ok) {
        setFeedback({ type: "err", text: result?.error || "Falha ao confirmar lote." });
        return;
      }
      setSessao([]);
      setFeedback({
        type: "ok",
        text: result.message || `Lote confirmado: ${result.total || 0} par(es) → PA + Bling.`,
      });
    } catch (err) {
      setFeedback({ type: "err", text: err?.message || "Erro ao confirmar lote." });
    } finally {
      setConfirmBusy(false);
      focar();
    }
  };

  return (
    <div className="grid gap-5 lg:grid-cols-[minmax(0,1.1fr)_minmax(0,0.9fr)]">
      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
        <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
          Leitor — código sequencial
        </div>
        <p className="mt-1 text-sm text-slate-500">
          Bipe o código da <strong>direita</strong> da etiqueta (contador). 1 bip = 1 par. Caixa já
          lida = erro.
        </p>

        <label className="mt-4 block">
          <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            Bip / Enter
          </span>
          <input
            ref={inputRef}
            className="mt-1 w-full rounded-2xl border border-slate-300 bg-slate-50 px-4 py-4 font-mono text-2xl font-bold tracking-wider text-slate-900 outline-none ring-rose-300 focus:bg-white focus:ring-2"
            value={scan}
            onChange={(e) => setScan(e.target.value)}
            onKeyDown={onKeyDown}
            disabled={busy || confirmBusy}
            placeholder="0000001"
            autoComplete="off"
            inputMode="numeric"
          />
        </label>

        <div className="mt-3 flex flex-wrap gap-2">
          <button
            type="button"
            onClick={() => processarBip(scan)}
            disabled={busy || confirmBusy || !scan.trim()}
            className="rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white disabled:opacity-40"
          >
            Registrar bip
          </button>
          <button
            type="button"
            onClick={desfazerUltimo}
            disabled={busy || confirmBusy || !sessao.length}
            className="rounded-2xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 disabled:opacity-40"
          >
            Desfazer último
          </button>
          <button
            type="button"
            onClick={limparSessaoLocal}
            disabled={busy || confirmBusy || !sessao.length}
            className="rounded-2xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 disabled:opacity-40"
          >
            Limpar lote
          </button>
        </div>

        {feedback.text ? (
          <div
            className={`mt-4 rounded-2xl px-4 py-3 text-sm font-medium ${
              feedback.type === "ok"
                ? "bg-emerald-50 text-emerald-800"
                : feedback.type === "err"
                  ? "bg-rose-50 text-rose-800"
                  : "bg-slate-50 text-slate-700"
            }`}
          >
            {feedback.text}
          </div>
        ) : null}

        <div className="mt-5 rounded-2xl border border-dashed border-slate-200 bg-slate-50 px-4 py-3">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            Total do lote
          </div>
          <div className="mt-1 text-3xl font-black tabular-nums text-slate-900">{totalPares}</div>
          <div className="text-sm text-slate-500">par(es) bipados nesta sessão</div>
        </div>

        <button
          type="button"
          onClick={confirmar}
          disabled={confirmBusy || busy || !sessao.length}
          className="mt-4 w-full rounded-2xl bg-rose-600 px-4 py-3.5 text-base font-bold text-white shadow-sm disabled:opacity-40"
        >
          {confirmBusy ? "Confirmando…" : `Confirmar lote → PA + Bling (${totalPares})`}
        </button>
      </section>

      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-3 flex items-end justify-between gap-2">
          <div>
            <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
              Resumo por SKU
            </div>
            <p className="mt-1 text-sm text-slate-500">Agrupado para a baixa M → PA.</p>
          </div>
        </div>

        {agrupado.length === 0 ? (
          <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-10 text-center text-sm text-slate-500">
            Nenhum bip ainda.
          </div>
        ) : (
          <div className="overflow-hidden rounded-2xl border border-slate-200">
            <table className="min-w-full text-left text-sm">
              <thead className="bg-slate-50 text-[11px] uppercase tracking-wide text-slate-500">
                <tr>
                  <th className="px-3 py-2">Ref / cor</th>
                  <th className="px-3 py-2">Nº</th>
                  <th className="px-3 py-2">SKU</th>
                  <th className="px-3 py-2 text-right">Qtd</th>
                </tr>
              </thead>
              <tbody>
                {agrupado.map((row) => (
                  <tr key={`${row.ref}-${row.cor}-${row.tamanho}`} className="border-t border-slate-100">
                    <td className="px-3 py-2 font-medium text-slate-800">
                      {row.ref} · {row.cor}
                    </td>
                    <td className="px-3 py-2 tabular-nums">{row.tamanho}</td>
                    <td className="px-3 py-2 font-mono text-xs text-slate-600">{row.sku}</td>
                    <td className="px-3 py-2 text-right font-bold tabular-nums">{row.qtd}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {sessao.length > 0 ? (
          <div className="mt-4">
            <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
              Últimos bips
            </div>
            <ul className="mt-2 max-h-56 space-y-1 overflow-auto text-sm">
              {sessao.slice(0, 40).map((item) => (
                <li
                  key={item.contador}
                  className="flex items-center justify-between gap-2 rounded-xl bg-slate-50 px-3 py-1.5 font-mono text-xs"
                >
                  <span className="font-bold text-slate-900">{item.contador}</span>
                  <span className="truncate text-slate-600">{item.sku}</span>
                </li>
              ))}
            </ul>
          </div>
        ) : null}
      </section>
    </div>
  );
}
