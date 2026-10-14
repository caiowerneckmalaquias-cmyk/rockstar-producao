import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  liberarBipPorContador,
  normalizeContadorScan,
  reservarBipPorContador,
} from "./etiquetasCaixaDb";

/**
 * Esteira / Acabamento — cada bip sobe 1 par no PA na hora.
 * Bling sincroniza em background (status discreto). Desfazer último reverte o bip.
 */
export default function EsteiraAcabamento({
  onBipConfirmado,
  onDesfazerBip,
  blingQueueStatus = { state: "idle", text: "" },
  onFlushBling,
}) {
  const inputRef = useRef(null);
  const [scan, setScan] = useState("");
  const [busy, setBusy] = useState(false);
  const [feedback, setFeedback] = useState({ type: "", text: "" });
  const [sessao, setSessao] = useState([]);

  const flushRef = useRef(onFlushBling);
  flushRef.current = onFlushBling;

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  useEffect(() => {
    return () => {
      flushRef.current?.();
    };
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
    if (typeof onBipConfirmado !== "function") {
      setFeedback({ type: "err", text: "Handler de bip não configurado." });
      return;
    }

    setBusy(true);
    setFeedback({ type: "", text: "" });
    let reservado = null;
    try {
      if (
        sessao.some(
          (s) =>
            normalizeContadorScan(s.contador) === codigo || s.contador_n === Number(codigo)
        )
      ) {
        setFeedback({ type: "err", text: `Caixa ${codigo} já está nesta sessão.` });
        return;
      }

      const reserved = await reservarBipPorContador(codigo);
      if (!reserved.ok) {
        setFeedback({ type: "err", text: reserved.error || "Falha no bip." });
        return;
      }
      reservado = reserved.row;

      const item = {
        contador: reservado.contador,
        contador_n: reservado.contador_n,
        sku: reservado.sku,
        ref: reservado.ref,
        cor: reservado.cor,
        tamanho: Number(reservado.tamanho),
        programacao: reservado.programacao || "",
      };

      const applied = await onBipConfirmado(item);
      if (!applied?.ok) {
        await liberarBipPorContador(reservado.contador);
        setFeedback({ type: "err", text: applied?.error || "Falha ao subir PA." });
        return;
      }

      setSessao((curr) => [
        {
          ...item,
          tirouDeM: Boolean(applied.tirouDeM),
          bipadoEm: new Date().toISOString(),
        },
        ...curr,
      ]);
      setFeedback({
        type: "ok",
        text: `OK ${item.contador} · ${item.sku} · PA +1${applied.tirouDeM ? " (baixou M)" : ""}`,
      });
    } catch (err) {
      if (reservado?.contador) {
        await liberarBipPorContador(reservado.contador);
      }
      setFeedback({ type: "err", text: err?.message || "Erro no bip." });
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
    if (!sessao.length || busy) return;
    if (typeof onDesfazerBip !== "function") {
      setFeedback({ type: "err", text: "Desfazer não configurado." });
      return;
    }
    const [ultimo, ...rest] = sessao;
    setBusy(true);
    try {
      const result = await onDesfazerBip(ultimo);
      if (!result?.ok) {
        setFeedback({ type: "err", text: result?.error || "Não foi possível desfazer." });
        return;
      }
      setSessao(rest);
      setFeedback({ type: "ok", text: `Desfeito ${ultimo.contador} · PA −1.` });
    } finally {
      setBusy(false);
      focar();
    }
  };

  const blingText =
    blingQueueStatus?.text ||
    (blingQueueStatus?.state === "syncing"
      ? "Bling: sincronizando…"
      : blingQueueStatus?.state === "ok"
        ? "Bling: ok"
        : blingQueueStatus?.state === "erro"
          ? "Bling: falhou (tenta de novo)"
          : blingQueueStatus?.state === "pending"
            ? "Bling: na fila…"
            : "");

  return (
    <div className="grid gap-5 lg:grid-cols-[minmax(0,1.1fr)_minmax(0,0.9fr)]">
      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
        <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
          Leitor — código sequencial
        </div>
        <p className="mt-1 text-sm text-slate-500">
          Bipe o código da <strong>direita</strong>. Cada bip sobe <strong>1 par no PA</strong> na
          hora. Caixa já lida = erro. Bling sincroniza sozinho.
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
            disabled={busy}
            placeholder="0000001"
            autoComplete="off"
            inputMode="numeric"
          />
        </label>

        <div className="mt-3 flex flex-wrap gap-2">
          <button
            type="button"
            onClick={() => processarBip(scan)}
            disabled={busy || !scan.trim()}
            className="rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white disabled:opacity-40"
          >
            Registrar bip
          </button>
          <button
            type="button"
            onClick={desfazerUltimo}
            disabled={busy || !sessao.length}
            className="rounded-2xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 disabled:opacity-40"
          >
            Desfazer último
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

        {blingText ? (
          <div className="mt-3 text-xs font-medium text-slate-500">{blingText}</div>
        ) : null}

        <div className="mt-5 rounded-2xl border border-dashed border-slate-200 bg-slate-50 px-4 py-3">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            Total da sessão
          </div>
          <div className="mt-1 text-3xl font-black tabular-nums text-slate-900">{totalPares}</div>
          <div className="text-sm text-slate-500">par(es) já no PA nesta sessão</div>
        </div>
      </section>

      <section className="rounded-[28px] border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-3 flex items-end justify-between gap-2">
          <div>
            <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
              Resumo por SKU
            </div>
            <p className="mt-1 text-sm text-slate-500">Já contados no PA (esta sessão).</p>
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
