import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
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
    .map((line) => line.replace(/\t+/g, " ").replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

function buildSuggestions(rows, minimos, vendas, tempoProducao) {
  const montagem = [];
  const pesponto = [];

  const diasPesponto = Number(tempoProducao?.pesponto) || 0;
  const diasMontagem = Number(tempoProducao?.montagem) || 0;
  const diasTotal = diasPesponto + diasMontagem;

  rows.forEach((row) => {
    const mins = minimos[row.ref]?.[row.cor];
    const sales = vendas[row.ref]?.[row.cor];
    if (!mins || !sales) return;

    const montSizes = makeEmptyGrid();
    const pespSizes = makeEmptyGrid();
    let montTotal = 0;
    let pespTotal = 0;
    let prioridade = 0;

    sizes.forEach((size) => {
      const item = row.data[size];
      const minimo = mins[size];
      const vendaMes = Number(sales[size]) || 0;
      const vendaDia = vendaMes / 30;

      const consumoDuranteMontagem = Math.ceil(vendaDia * diasMontagem);
      const consumoDuranteCicloTotal = Math.ceil(vendaDia * diasTotal);

      const needPA = Math.max(0, (minimo.pa + consumoDuranteMontagem) - item.pa);
      const prodAtual = item.est + item.m + item.p;
      const needProd = Math.max(0, (minimo.prod + consumoDuranteCicloTotal) - prodAtual);

      const mont = Math.min(item.est, round12(needPA));
      const pesp = round12(needProd);

      montSizes[size] = mont;
      pespSizes[size] = pesp;
      montTotal += mont;
      pespTotal += pesp;

      const riscoMontagem = Math.max(0, consumoDuranteMontagem - item.pa);
      const riscoProducao = Math.max(0, consumoDuranteCicloTotal - prodAtual);
      prioridade += vendaDia + needPA + needProd + (riscoMontagem * 2) + (riscoProducao * 2);
    });

    if (montTotal > 0) montagem.push({ tipo: "Montagem", ref: row.ref, cor: row.cor, sizes: montSizes, total: montTotal, prioridade });
    if (pespTotal > 0) pesponto.push({ tipo: "Pesponto", ref: row.ref, cor: row.cor, sizes: pespSizes, total: pespTotal, prioridade });
  });

  montagem.sort((a, b) => b.prioridade - a.prioridade || b.total - a.total);
  pesponto.sort((a, b) => b.prioridade - a.prioridade || b.total - a.total);
  return { montagem, pesponto };
}

function splitIntoFichas(list, vendas) {
  const fichas = [];
  const MAX_TOTAL_FICHA = 396;
  const MAX_POR_NUMERO_FICHA = 84;
  const LOTE = 12;
  const GRADE_BAIXA = [34, 35, 36, 37, 38, 39];
  const GRADE_ALTA = [40, 41, 42, 43, 44];

  const distribuirGrupo = (entry, grupoSizes, sufixo = "") => {
    const totalGrupo = grupoSizes.reduce((acc, size) => acc + (entry.sizes[size] || 0), 0);
    if (!totalGrupo) return [];

    const minFichasPorTotal = Math.ceil(totalGrupo / MAX_TOTAL_FICHA);
    const minFichasPorNumero = Math.max(
      1,
      ...grupoSizes.map((size) => Math.ceil((entry.sizes[size] || 0) / MAX_POR_NUMERO_FICHA))
    );
    const numFichas = Math.max(1, minFichasPorTotal, minFichasPorNumero);

    const fichasTemp = Array.from({ length: numFichas }, () => ({
      sizes: makeEmptyGrid(),
      total: 0,
    }));

    const vendasRef = vendas?.[entry.ref]?.[entry.cor] || makeEmptyGrid();
    const itens = grupoSizes
      .map((size) => ({
        size,
        qtd: entry.sizes[size] || 0,
        venda: Number(vendasRef[size]) || 0,
        lotesRestantes: Math.floor((entry.sizes[size] || 0) / LOTE),
        current: 0,
      }))
      .filter((item) => item.lotesRestantes > 0);

    if (!itens.length) return [];

    const totalPeso = itens.reduce((acc, item) => acc + (item.venda > 0 ? item.venda : Math.max(1, item.lotesRestantes)), 0) || 1;

    // Entrada mínima: se precisa de um número, tenta colocar 12 em alguma ficha.
    // A prioridade aqui já respeita vendas maiores primeiro.
    [...itens]
      .sort((a, b) => {
        if (b.venda !== a.venda) return b.venda - a.venda;
        if (b.qtd !== a.qtd) return b.qtd - a.qtd;
        return a.size - b.size;
      })
      .forEach((item) => {
        if (item.lotesRestantes <= 0) return;

        const candidatos = fichasTemp
          .map((ficha, fichaIdx) => ({
            ficha,
            fichaIdx,
            total: ficha.total,
            noNumero: ficha.sizes[item.size] || 0,
          }))
          .filter((slot) => slot.total + LOTE <= MAX_TOTAL_FICHA && slot.noNumero + LOTE <= MAX_POR_NUMERO_FICHA)
          .sort((a, b) => {
            if (a.total !== b.total) return a.total - b.total;
            if (a.noNumero !== b.noNumero) return a.noNumero - b.noNumero;
            return a.fichaIdx - b.fichaIdx;
          });

        if (!candidatos.length) return;

        candidatos[0].ficha.sizes[item.size] += LOTE;
        candidatos[0].ficha.total += LOTE;
        item.lotesRestantes -= 1;
      });

    // Distribuição principal: ponderada por vendas, balanceando ficha leve primeiro.
    let lotesPendentes = itens.reduce((acc, item) => acc + item.lotesRestantes, 0);

    while (lotesPendentes > 0) {
      const ativos = itens.filter((item) => item.lotesRestantes > 0);
      if (!ativos.length) break;

      ativos.forEach((item) => {
        const peso = item.venda > 0 ? item.venda : Math.max(1, item.lotesRestantes);
        item.current += peso;
      });

      ativos.sort((a, b) => {
        if (b.current !== a.current) return b.current - a.current;
        if (b.venda !== a.venda) return b.venda - a.venda;
        if (b.qtd !== a.qtd) return b.qtd - a.qtd;
        return a.size - b.size;
      });

      const escolhido = ativos[0];
      escolhido.current -= totalPeso;

      const candidatos = fichasTemp
        .map((ficha, fichaIdx) => ({
          ficha,
          fichaIdx,
          total: ficha.total,
          noNumero: ficha.sizes[escolhido.size] || 0,
        }))
        .filter((slot) => slot.total + LOTE <= MAX_TOTAL_FICHA && slot.noNumero + LOTE <= MAX_POR_NUMERO_FICHA)
        .sort((a, b) => {
          if (a.total !== b.total) return a.total - b.total;
          if (a.noNumero !== b.noNumero) return a.noNumero - b.noNumero;
          return a.fichaIdx - b.fichaIdx;
        });

      if (!candidatos.length) {
        escolhido.lotesRestantes = 0;
        continue;
      }

      candidatos[0].ficha.sizes[escolhido.size] += LOTE;
      candidatos[0].ficha.total += LOTE;
      escolhido.lotesRestantes -= 1;
      lotesPendentes -= 1;
    }

    return fichasTemp
      .map((ficha, idx) => ({ ficha, idx }))
      .filter(({ ficha }) => ficha.total > 0)
      .map(({ ficha, idx }) => ({
        nome: `${entry.tipo} ${entry.cor}${sufixo ? ` • ${sufixo}` : ""} • Ficha ${String(idx + 1).padStart(2, "0")}`,
        ref: entry.ref,
        cor: entry.cor,
        sizes: ficha.sizes,
        total: ficha.total,
      }));
  };

  list.forEach((entry) => {
    const total = sizes.reduce((acc, size) => acc + (entry.sizes[size] || 0), 0);
    if (!total) return;

    // Se couber em uma ficha, pode misturar grade baixa e alta.
    if (total <= MAX_TOTAL_FICHA) {
      fichas.push(...distribuirGrupo(entry, sizes));
      return;
    }

    // Quando o volume é alto, separa grade baixa e grade alta para balancear melhor.
    fichas.push(...distribuirGrupo(entry, GRADE_BAIXA, "Grade Baixa"));
    fichas.push(...distribuirGrupo(entry, GRADE_ALTA, "Grade Alta"));
  });

  return fichas;
}

function buildProgramacaoPeriodo(fichasBase, suggestionsBase, capacidadeDia = 396, dias = 1, tipo = "") {
  const prioridadeMap = new Map();

  (suggestionsBase || []).forEach((item) => {
    prioridadeMap.set(`${item.ref}__${item.cor}`, item.prioridade || 0);
  });

  const grupos = {};
  (fichasBase || []).forEach((ficha) => {
    const key = `${ficha.ref}__${ficha.cor}`;
    if (!grupos[key]) {
      grupos[key] = {
        key,
        tipo,
        ref: ficha.ref,
        cor: ficha.cor,
        prioridade: prioridadeMap.get(key) || 0,
        peso: Math.max(1, prioridadeMap.get(key) || 0),
        current: 0,
        fichas: [],
      };
    }
    grupos[key].fichas.push({ ...ficha, tipo });
  });

  Object.values(grupos).forEach((grupo) => {
    grupo.fichas.sort((a, b) => a.total - b.total || a.nome.localeCompare(b.nome, "pt-BR"));
  });

  const ativos = Object.values(grupos).filter((grupo) => grupo.fichas.length > 0);
  const totalPeso = ativos.reduce((acc, grupo) => acc + grupo.peso, 0) || 1;
  const diasProgramados = [];
  let ultimoGrupoGlobal = "";

  const escolherGrupo = (gruposDisponiveis, restante, ultimoGrupoDia) => {
    const candidatos = gruposDisponiveis.filter(
      (grupo) => grupo.fichas.length > 0 && grupo.fichas.some((ficha) => ficha.total <= restante)
    );

    if (!candidatos.length) return null;

    candidatos.forEach((grupo) => {
      grupo.current += grupo.peso;
    });

    const ordenados = [...candidatos].sort((a, b) => {
      const penalA = (a.key === ultimoGrupoDia ? 1 : 0) + (a.key === ultimoGrupoGlobal ? 1 : 0);
      const penalB = (b.key === ultimoGrupoDia ? 1 : 0) + (b.key === ultimoGrupoGlobal ? 1 : 0);
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
    let restante = capacidadeDia;
    let ultimoGrupoDia = "";
    const selecionadas = [];

    while (restante >= 12) {
      const grupoEscolhido = escolherGrupo(ativos, restante, ultimoGrupoDia);
      if (!grupoEscolhido) break;

      const fichaEscolhida = [...grupoEscolhido.fichas]
        .filter((ficha) => ficha.total <= restante)
        .sort((a, b) => b.total - a.total || a.nome.localeCompare(b.nome, "pt-BR"))[0];

      if (!fichaEscolhida) {
        grupoEscolhido.current -= totalPeso;
        break;
      }

      grupoEscolhido.fichas = grupoEscolhido.fichas.filter((f) => f.nome !== fichaEscolhida.nome);
      selecionadas.push({
        ...fichaEscolhida,
        prioridade: grupoEscolhido.prioridade,
        grupoKey: grupoEscolhido.key,
      });
      restante -= fichaEscolhida.total;
      ultimoGrupoDia = grupoEscolhido.key;
      ultimoGrupoGlobal = grupoEscolhido.key;
    }

    diasProgramados.push({
      dia,
      capacidadeDia,
      totalProgramado: capacidadeDia - restante,
      restante,
      fichas: selecionadas,
    });
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
  };
}

function SummaryCard({ title, value, subtitle }) {
  return (
    <div className="relative overflow-hidden rounded-[28px] border border-[#E5E7EB] bg-white p-5 shadow-[0_12px_30px_rgba(15,23,42,0.08)] backdrop-blur">
      <div className="absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-[#8B1E2D] via-[#6F1421] to-[#0F172A]" />
      <div className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-400">{title}</div>
      <div className="mt-3 text-3xl font-black tracking-tight text-[#0F172A]">{value}</div>
      <div className="mt-2 text-sm text-slate-500">{subtitle}</div>
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
  const [rows, setRows] = useState(initialRows);
  const [minimos, setMinimos] = useState(initialMinimos);
  const [vendas, setVendas] = useState(initialVendas);
  const [importText, setImportText] = useState("");
  const [importFileName, setImportFileName] = useState("");
  const [importFeedback, setImportFeedback] = useState("");
  const [importPreview, setImportPreview] = useState([]);
  const [ultimaImportacaoGcm, setUltimaImportacaoGcm] = useState(null);
  const [salesImportFileName, setSalesImportFileName] = useState("");
  const [salesImportFeedback, setSalesImportFeedback] = useState("");
  const [salesImportPreview, setSalesImportPreview] = useState([]);
  const [vendasDraft, setVendasDraft] = useState(initialVendas);
  const [vendasDirty, setVendasDirty] = useState(false);
  const [historicoVendasManuais, setHistoricoVendasManuais] = useState([]);
  const [pespontoForm, setPespontoForm] = useState({ ref: "BTCV010", cor: "Preto", grid: makeEmptyGrid(), programacao: "Programação A" });
  const [montagemForm, setMontagemForm] = useState({ ref: "BTCV010", cor: "Preto", grid: makeEmptyGrid(), programacao: "Programação A" });
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
    ref: "BTCV010",
    cor: "Preto",
    tipo: "entrada",
    size: 34,
    qtd: 0,
    motivo: "",
  });
  const [ajustesEst, setAjustesEst] = useState([]);
  const [ajusteEstErro, setAjusteEstErro] = useState("");
  const [draftMinimos, setDraftMinimos] = useState(initialMinimos);
  const [dirtyMinimos, setDirtyMinimos] = useState(false);
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
  const [programacaoSubAba, setProgramacaoSubAba] = useState("Pesponto");
  const [feriadosTexto, setFeriadosTexto] = useState("");

  const refs = useMemo(() => rows.map((r) => `${r.ref}__${r.cor}`), [rows]);

  const previewBySelection = (form) => {
    const row = rows.find((item) => item.ref === form.ref && item.cor === form.cor);
    if (!row) return null;

    const totalPesponto = sizes.reduce((acc, size) => acc + (row.data[size]?.p || 0), 0);
    const totalMontagem = sizes.reduce((acc, size) => acc + (row.data[size]?.m || 0), 0);
    const totalEst = sizes.reduce((acc, size) => acc + (row.data[size]?.est || 0), 0);

    return {
      row,
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
    return rows.filter((row) => {
      if (controleFiltroRef !== "TODAS" && row.ref !== controleFiltroRef) return false;
      if (controleFiltroCor !== "TODAS" && row.cor !== controleFiltroCor) return false;
      return true;
    });
  }, [rows, controleFiltroRef, controleFiltroCor]);

  const visibleSizesControle = useMemo(() => {
    if (controleFiltroNumero === "TODAS") return sizes;
    return [Number(controleFiltroNumero)];
  }, [controleFiltroNumero]);

  const sortedRowsByRefCor = useMemo(() => {
    return [...rows].sort((a, b) => {
      const refCompare = String(a.ref).localeCompare(String(b.ref), "pt-BR", { numeric: true, sensitivity: "base" });
      if (refCompare !== 0) return refCompare;
      return String(a.cor).localeCompare(String(b.cor), "pt-BR", { sensitivity: "base" });
    });
  }, [rows]);

  const metrics = useMemo(() => {
    let criticos = 0;
    let atencaoPA = 0;
    let atencaoProd = 0;
    let ok = 0;
    let costura = 0;
    rows.forEach((row) => {
      sizes.forEach((size) => {
        const item = row.data[size];
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
  }, [rows, minimos]);

  const suggestions = useMemo(() => buildSuggestions(rows, minimos, vendas, tempoProducao), [rows, minimos, vendas, tempoProducao]);
  const fichasMontagem = useMemo(() => splitIntoFichas(suggestions.montagem, vendas), [suggestions, vendas]);
  const fichasPesponto = useMemo(() => splitIntoFichas(suggestions.pesponto, vendas), [suggestions, vendas]);
  const programacaoMontagem = useMemo(
    () => buildProgramacaoPeriodo(fichasMontagem, suggestions.montagem, 396, programacaoDias, "Montagem"),
    [fichasMontagem, suggestions, programacaoDias]
  );
  const programacaoPesponto = useMemo(
    () => buildProgramacaoPeriodo(fichasPesponto, suggestions.pesponto, 396, programacaoDias, "Pesponto"),
    [fichasPesponto, suggestions, programacaoDias]
  );

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
    }
  };

  carregarDadosIniciais();
}, []);

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

  const applyGridDelta = (ref, cor, updates) => {
    setRows((current) =>
      current.map((row) => {
        if (row.ref !== ref || row.cor !== cor) return row;

        const nextData = { ...row.data };
        updates.forEach(({ size, field, delta }) => {
          nextData[size] = {
            ...nextData[size],
            [field]: Math.max(0, (nextData[size]?.[field] || 0) + delta),
          };
        });

        return {
          ...row,
          data: nextData,
        };
      })
    );
  };

  const executeMov = async (tipo, form, force = false) => {
    console.log("EXECUTE MOV FOI CHAMADO", { tipo, form });
    const items = sizes
      .map((size) => ({ size, qtd: Number(form.grid[size]) || 0 }))
      .filter((x) => x.qtd > 0);

    const programacaoNome = String(form.programacao || "").trim();
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const programacaoDuplicada = programacaoNome
      ? source.some(
          (item) =>
            String(item.programacao || "").trim().toUpperCase() === programacaoNome.toUpperCase()
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
        [tipo]: `Já existe uma programação com o nome "${programacaoNome}" em ${tipo}. Use outro nome.`,
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
      setConfirmMov({ tipo, form });
      return;
    }

    setMovError((curr) => ({ ...curr, [tipo]: "" }));

    applyGridDelta(
      form.ref,
      form.cor,
      items.map((item) => ({
        size: item.size,
        field: tipo === "Pesponto" ? "p" : "m",
        delta: item.qtd,
      }))
    );

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
    await salvarMovimentacao(movsParaSalvar);
    console.log("DEPOIS DE SALVAR");
    setConfirmMov(null);
  };

  const getMovErrorMessage = (tipo, form) => {
    if (tipo !== "Pesponto" && tipo !== "Montagem") return "";

    const programacaoNome = String(form.programacao || "").trim();
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const programacaoDuplicada = programacaoNome
      ? source.some(
          (item) =>
            String(item.programacao || "").trim().toUpperCase() === programacaoNome.toUpperCase()
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
      mensagens.push(`já existe uma programação com esse nome em ${tipo}`);
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

  const revertLancamentoImpacto = (tipo, lancamento) => {
    const updates = lancamento.items.map((item) => ({
      size: item.size,
      field: tipo === "Pesponto" ? "p" : "m",
      delta: -item.qtd,
    }));
    applyGridDelta(lancamento.ref, lancamento.cor, updates);
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

  const startEditLancamento = (tipo, lancamento) => {
    if (lancamento.status === "Finalizado") return;
    revertLancamentoImpacto(tipo, lancamento);
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

  const confirmDeleteLancamento = ({ tipo, lancamentoId }) => {
    const source = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const alvo = source.find((item) => item.id === lancamentoId);
    if (!alvo || alvo.status === "Finalizado") {
      setConfirmAction(null);
      return;
    }

    revertLancamentoImpacto(tipo, alvo);

    if (tipo === "Pesponto") {
      setPespontoLancamentos((curr) => curr.filter((item) => item.id !== lancamentoId));
    } else {
      setMontagemLancamentos((curr) => curr.filter((item) => item.id !== lancamentoId));
    }

    setConfirmAction(null);
  };

  const confirmFinalizarProgramacao = ({ tipo, programacao }) => {
    const lancamentos = tipo === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const alvo = lancamentos.filter((l) => l.programacao === programacao && l.status !== "Finalizado");

    alvo.forEach((lancamento) => {
      const updates = lancamento.items.flatMap((item) =>
        tipo === "Pesponto"
          ? [
              { size: item.size, field: "p", delta: -item.qtd },
              { size: item.size, field: "est", delta: item.qtd },
            ]
          : [
              { size: item.size, field: "m", delta: -item.qtd },
              { size: item.size, field: "pa", delta: item.qtd },
            ]
      );

      applyGridDelta(lancamento.ref, lancamento.cor, updates);
    });

    const dataFinalizacao = new Date().toLocaleDateString("pt-BR");

    if (tipo === "Pesponto") {
      setPespontoLancamentos((curr) => curr.map((l) => (l.programacao === programacao ? { ...l, status: "Finalizado", dataFinalizacao } : l)));
    } else {
      setMontagemLancamentos((curr) => curr.map((l) => (l.programacao === programacao ? { ...l, status: "Finalizado", dataFinalizacao } : l)));
    }

    setConfirmAction(null);
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawText = XLSX.utils.sheet_to_txt(firstSheet);
      const parsed = parseGcmRawText(rawText);
      setImportText(rawText);
      setImportFileName(file.name);
      setImportPreview(parsed);
      setImportFeedback(`Arquivo carregado com ${parsed.length} bloco(s) do GCM.`);
    } catch {
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
      console.log("MOVIMENTACOES ENVIADAS:", movs);

      const { data, error } = await supabase
        .from("movimentacoes")
        .insert(movs)
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

    const rowsEstruturadas = Object.values(agrupado).map((row) => {
      const dataCompleta = {};

      sizes.forEach((size) => {
        dataCompleta[size] = row.data[size] || {
          pa: 0,
          est: 0,
          m: 0,
          p: 0,
        };
      });

      return {
        ref: row.ref,
        cor: row.cor,
        data: dataCompleta,
      };
    });

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
      .order("created_at", { ascending: false });

    if (error) {
      console.log("ERRO AO CARREGAR MOVIMENTACOES:", error);
      return { pesponto: [], montagem: [] };
    }

    if (!data || !data.length) {
      return { pesponto: [], montagem: [] };
    }

    const agrupar = (lista, tipo) => {
      const mapa = new Map();

      lista
        .filter((item) => item.tipo === tipo)
        .forEach((item) => {
          const programacao = item.programacao || "Sem programação";
          const ref = item.ref || "";
          const cor = item.cor || "";
          const status = item.status || "Em aberto";
          const dataLancamento =
            item.created_at
              ? new Date(item.created_at).toLocaleDateString("pt-BR")
              : new Date().toLocaleDateString("pt-BR");

          const chave = `${tipo}__${programacao}__${ref}__${cor}__${status}__${dataLancamento}`;

          if (!mapa.has(chave)) {
            mapa.set(chave, {
              id: chave,
              programacao,
              ref,
              cor,
              items: [],
              total: 0,
              status,
              dataLancamento,
            });
          }

          const grupo = mapa.get(chave);
          const qtd = Number(item.quantidade) || 0;
          const numero = Number(item.numero) || 0;

          grupo.items.push({
            size: numero,
            qtd,
          });

          grupo.total += qtd;
        });

      return Array.from(mapa.values());
    };

    const pesponto = agrupar(data, "Pesponto");
    const montagem = agrupar(data, "Montagem");

    return { pesponto, montagem };
  } catch (err) {
    console.log("ERRO GERAL AO CARREGAR MOVIMENTACOES:", err);
    return { pesponto: [], montagem: [] };
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

    const parsed = parseGcmRawText(importText);

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

    const estoqueParaSalvar = [];
    parsed.forEach((item) => {
      sizes.forEach((numero) => {
        const quantidade = Number(item.data?.[numero] || 0);
        if (quantidade > 0) {
          estoqueParaSalvar.push({
            ref: item.ref,
            cor: item.cor,
            numero,
            pa: quantidade,
            est: 0,
            m: 0,
            p: 0,
          });
        }
      });
    });

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
    const parsed = parseGcmRawText(importText);
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
    const tempoTotal = (Number(tempoProducao?.pesponto) || 0) + (Number(tempoProducao?.montagem) || 0);

    const totalPA = rows.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.pa || 0), 0), 0);
    const totalEst = rows.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.est || 0), 0), 0);
    const totalMontagemAtual = rows.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.m || 0), 0), 0);
    const totalPespontoAtual = rows.reduce((acc, row) => acc + sizes.reduce((sum, size) => sum + (row.data[size]?.p || 0), 0), 0);
    const vendaMensalTotal = rows.reduce((acc, row) => {
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
    const mesAtual = hoje.getMonth();
    const anoAtual = hoje.getFullYear();
    const hojeTexto = hoje.toLocaleDateString("pt-BR");

    const finalizadosMesPesponto = pespontoLancamentos.filter((item) => {
      if (item.status !== "Finalizado" || !item.dataFinalizacao) return false;
      const data = parseDateBrToDate(item.dataFinalizacao)
      return data && data.getMonth() === mesAtual && data.getFullYear() === anoAtual;
    });

    const finalizadosMesMontagem = montagemLancamentos.filter((item) => {
      if (item.status !== "Finalizado" || !item.dataFinalizacao) return false;
      const data = parseDateBrToDate(item.dataFinalizacao)
      return data && data.getMonth() === mesAtual && data.getFullYear() === anoAtual;
    });

    const totalFinalizadoMesPesponto = finalizadosMesPesponto.reduce((acc, item) => acc + (item.total || 0), 0);
    const totalFinalizadoMesMontagem = finalizadosMesMontagem.reduce((acc, item) => acc + (item.total || 0), 0);
    const mediaFinalizadaPesponto = totalFinalizadoMesPesponto / diasUteisNoMes;
    const mediaFinalizadaMontagem = totalFinalizadoMesMontagem / diasUteisNoMes;

    const finalizadoHojePesponto = finalizadosMesPesponto
      .filter((item) => item.dataFinalizacao === hojeTexto)
      .reduce((acc, item) => acc + (item.total || 0), 0);
    const finalizadoHojeMontagem = finalizadosMesMontagem
      .filter((item) => item.dataFinalizacao === hojeTexto)
      .reduce((acc, item) => acc + (item.total || 0), 0);

    const menorCobertura = rows
      .flatMap((row) => sizes.map((size) => {
        const vendaMes = Number(vendas?.[row.ref]?.[row.cor]?.[size]) || 0;
        const cobertura = coberturaDias(row.data[size]?.pa || 0, vendaMes);
        return { ref: row.ref, cor: row.cor, size, cobertura, pa: row.data[size]?.pa || 0, vendaMes };
      }))
      .filter((item) => item.cobertura !== null)
      .sort((a, b) => a.cobertura - b.cobertura)
      .slice(0, 5);

    const topModelos = rows
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

    return (
      <PageShell title="Dashboard" subtitle="Visão executiva da operação com foco em giro, risco de ruptura, programação e andamento da produção.">
        <section className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-7 gap-4">
          <SummaryCard title="PA total" value={totalPA} subtitle="Pares em produto acabado" />
          <SummaryCard title="Costura pronta" value={totalEst} subtitle="Pares prontos para montagem" />
          <SummaryCard title="Na montagem" value={totalMontagemAtual} subtitle="Fluxo atual de montagem" />
          <SummaryCard title="No pesponto" value={totalPespontoAtual} subtitle="Fluxo atual de pesponto" />
          <SummaryCard title="Venda diária" value={vendaDiariaTotal.toFixed(1)} subtitle="Base mensal ÷ 30 dias" />
          <SummaryCard title="Média finalizada • Pesponto" value={mediaFinalizadaPesponto.toFixed(1)} subtitle={`Finalizados no mês ÷ ${diasUteisNoMes} dias úteis • Hoje ${finalizadoHojePesponto}`} />
          <SummaryCard title="Média finalizada • Montagem" value={mediaFinalizadaMontagem.toFixed(1)} subtitle={`Finalizados no mês ÷ ${diasUteisNoMes} dias úteis • Hoje ${finalizadoHojeMontagem}`} />
        </section>

        <section className="grid grid-cols-1 xl:grid-cols-[1.2fr_0.8fr] gap-6">
          <div className="xl:col-span-2 bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
              <div>
                <h2 className="font-bold text-lg">Calendário útil do mês</h2>
                <p className="text-sm text-slate-500 mt-1">Os cards de média finalizada usam apenas dias úteis transcorridos, desconsiderando sábados, domingos e os feriados informados abaixo.</p>
              </div>
              <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-right">
                <div className="text-xs text-slate-500">Dias úteis considerados</div>
                <div className="text-2xl font-bold text-slate-900">{diasUteisNoMes}</div>
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
                  <div className="flex items-center justify-between"><span>Dias corridos no mês</span><span className="font-semibold">{hoje.getDate()}</span></div>
                  <div className="flex items-center justify-between"><span>Feriados informados</span><span className="font-semibold">{feriadosLista.length}</span></div>
                  <div className="flex items-center justify-between"><span>Dias úteis usados</span><span className="font-semibold text-[#8B1E2D]">{diasUteisNoMes}</span></div>
                </div>
              </div>
            </div>
          </div>
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <div>
                <h2 className="font-bold text-lg">Radar da operação</h2>
                <p className="text-sm text-slate-500 mt-1">Indicadores principais para decidir a prioridade do dia.</p>
              </div>
              <span className="px-3 py-1 rounded-full border text-xs font-semibold bg-[#FFF7F8] text-[#8B1E2D] border-[#E7C7CC]">Lead time {tempoTotal} dias</span>
            </div>

            <div className="mt-5 grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="rounded-[24px] border border-slate-200 bg-slate-50 p-5">
                <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">Situação do estoque</div>
                <div className="mt-4 space-y-3">
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Itens críticos</span><span className="text-2xl font-black text-red-600">{metrics.criticos}</span></div>
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Atenção PA</span><span className="text-xl font-bold text-amber-600">{metrics.atencaoPA}</span></div>
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Atenção produção</span><span className="text-xl font-bold text-sky-700">{metrics.atencaoProd}</span></div>
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Itens OK</span><span className="text-xl font-bold text-emerald-600">{metrics.ok}</span></div>
                </div>
              </div>

              <div className="rounded-[24px] border border-slate-200 bg-slate-50 p-5">
                <div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">Programação do dia</div>
                <div className="mt-4 space-y-3">
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Pesponto hoje</span><span className="text-2xl font-black text-[#0F172A]">{programacaoHojePesponto?.totalProgramado || 0}</span></div>
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Montagem hoje</span><span className="text-2xl font-black text-[#0F172A]">{programacaoHojeMontagem?.totalProgramado || 0}</span></div>
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Fichas pesponto</span><span className="text-xl font-bold text-[#8B1E2D]">{programacaoHojePesponto?.fichas?.length || 0}</span></div>
                  <div className="flex items-center justify-between text-sm"><span className="text-slate-500">Fichas montagem</span><span className="text-xl font-bold text-[#8B1E2D]">{programacaoHojeMontagem?.fichas?.length || 0}</span></div>
                </div>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <div>
                <h2 className="font-bold text-lg">Próximas ações</h2>
                <p className="text-sm text-slate-500 mt-1">Os pontos que merecem atenção imediata.</p>
              </div>
            </div>
            <div className="mt-4 space-y-3">
              {menorCobertura.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">Sem vendas suficientes para montar alertas.</div>
              ) : (
                menorCobertura.map((item, idx) => (
                  <div key={`${item.ref}-${item.cor}-${item.size}-dash-${idx}`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex items-center justify-between gap-3">
                      <div>
                        <div className="font-semibold text-slate-900">{item.ref} • {item.cor}</div>
                        <div className="text-xs text-slate-500 mt-1">Numeração {item.size} • PA {item.pa}</div>
                      </div>
                      <span className={`px-3 py-1 rounded-full border text-xs font-semibold ${coberturaBadgeClass(item.cobertura, tempoTotal)}`}>
                        {item.cobertura?.toFixed(1)} dia(s)
                      </span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </section>

        <section className="grid grid-cols-1 xl:grid-cols-2 gap-6">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-bold text-lg">Top modelos por venda</h2>
              <span className="px-3 py-1 rounded-full border text-xs font-semibold bg-emerald-100 text-emerald-700 border-emerald-200">Mensal</span>
            </div>
            <div className="mt-4 space-y-3">
              {topModelos.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">Sem vendas lançadas ainda.</div>
              ) : (
                topModelos.map((item, idx) => (
                  <div key={`${item.ref}-${item.cor}-dashboard-top`} className="rounded-2xl border border-slate-200 px-4 py-3">
                    <div className="flex items-center justify-between gap-3">
                      <div>
                        <div className="font-semibold text-slate-900">#{idx + 1} {item.ref} • {item.cor}</div>
                        <div className="text-xs text-slate-500 mt-1">PA atual {item.totalPA}</div>
                      </div>
                      <div className="text-right">
                        <div className="text-2xl font-black text-[#0F172A]">{item.totalVendido}</div>
                        <div className="text-xs text-slate-500">vendidos/mês</div>
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-bold text-lg">Resumo das sugestões</h2>
              <span className="px-3 py-1 rounded-full border text-xs font-semibold bg-[#FFF7F8] text-[#8B1E2D] border-[#E7C7CC]">Planejamento</span>
            </div>
            <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                <div className="text-sm text-slate-500">Sugestões de pesponto</div>
                <div className="mt-2 text-3xl font-black text-[#0F172A]">{suggestions.pesponto.length}</div>
                <div className="mt-2 text-xs text-slate-500">Total sugerido: {suggestions.pesponto.reduce((acc, item) => acc + item.total, 0)} pares</div>
              </div>
              <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                <div className="text-sm text-slate-500">Sugestões de montagem</div>
                <div className="mt-2 text-3xl font-black text-[#0F172A]">{suggestions.montagem.length}</div>
                <div className="mt-2 text-xs text-slate-500">Total sugerido: {suggestions.montagem.reduce((acc, item) => acc + item.total, 0)} pares</div>
              </div>
            </div>
          </div>
        </section>
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
                      const item = row.data[size];
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
    const selectionPreview = previewBySelection(form);
    const liveError = getMovErrorMessage(title, form);
    const lancamentos = title === "Pesponto" ? pespontoLancamentos : montagemLancamentos;
    const totalAtual = sizes.reduce((acc, size) => acc + (Number(form.grid[size]) || 0), 0);
    const agrupados = lancamentos.reduce((acc, item) => {
      if (!acc[item.programacao]) acc[item.programacao] = [];
      acc[item.programacao].push(item);
      return acc;
    }, {});

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

              {(liveError || movError[title]) && (
                <div className="mt-3 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-medium text-red-700">
                  {liveError || movError[title]}
                </div>
              )}
            </div>
          </div>

          <div className="space-y-6">
            <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <h2 className="font-bold text-lg">Lançamentos enviados</h2>
                  <p className="text-sm text-slate-500 mt-1">Finalize para atualizar o estoque.</p>
                </div>
                <span className="text-sm text-slate-500">{lancamentos.length} lançamentos</span>
              </div>

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
            </div>

            {selectionPreview && (
              <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
                <div className="text-sm font-bold uppercase tracking-wide text-[#8B1E2D]">Prévia da cor</div>
                <div className="mt-3 text-2xl font-semibold text-slate-900">
                  {selectionPreview.row.ref} • {selectionPreview.row.cor}
                </div>

                <div className="mt-6 grid grid-cols-1 md:grid-cols-3 gap-4 text-sm text-slate-700">
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <span>No Pesponto</span>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalPesponto}</span>
                  </div>
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <span>Costura pronta</span>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalEst}</span>
                  </div>
                  <div className="flex items-center justify-between rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3">
                    <span>Na Montagem</span>
                    <span className="text-3xl font-bold text-slate-900">{selectionPreview.totalMontagem}</span>
                  </div>
                </div>

                <div className="mt-6 grid grid-cols-4 md:grid-cols-6 xl:grid-cols-8 gap-3 text-center text-sm text-slate-600">
                  {sizes.map((size) => (
                    <div key={size} className="rounded-2xl border border-sky-100 bg-slate-50 px-3 py-3">
                      <div className="font-semibold text-slate-900">{size}</div>
                      <div className="mt-1">P: {selectionPreview.row.data[size]?.p || 0}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      </PageShell>
    );
  };

  const aplicarAjusteEst = () => {
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
    applyGridDelta(ajusteEstForm.ref, ajusteEstForm.cor, [
      {
        size: ajusteEstForm.size,
        field: "est",
        delta: ajusteEstForm.tipo === "entrada" ? qtd : -qtd,
      },
    ]);

    setAjustesEst((curr) => [
      {
        id: `ajuste-est-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        data: new Date().toLocaleDateString("pt-BR"),
        ref: ajusteEstForm.ref,
        cor: ajusteEstForm.cor,
        tipo: ajusteEstForm.tipo,
        size: ajusteEstForm.size,
        qtd,
        motivo: ajusteEstForm.motivo || "Sem motivo informado",
      },
      ...curr,
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
                  {sizes.map((size) => <td key={size} className="border px-3 py-3 text-center">{row.data[size].est}</td>)}
                  <td className="border px-4 py-3 text-center font-bold">{sizes.reduce((acc, size) => acc + row.data[size].est, 0)}</td>
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
                <div className="font-bold text-lg">Histórico de ajustes</div>
                <div className="text-sm text-slate-500 mt-1">Registro dos últimos movimentos manuais.</div>
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
                        <div className="text-xs text-slate-500 mt-1">{ajuste.data} • Numeração {ajuste.size}</div>
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
  setMinimos(draftMinimos);
  setTempoProducao(tempoProducaoDraft);
  await salvarMinimosNoBanco(draftMinimos);
  setDirtyMinimos(false);
};

    return (
      <PageShell title="Minimos" subtitle="Aba equivalente aos mínimos da planilha, com PA e PROD por referência, cor e numeração.">
        <div className="space-y-4">
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
            <div className="flex items-center justify-between gap-4 mb-4">
              <div>
                <div className="font-bold text-lg">Tempo de Produção</div>
                <div className="text-sm text-slate-500 mt-1">Esses tempos entram na lógica das sugestões para antecipar risco de falta antes de virar PA.</div>
              </div>
              <div className="rounded-2xl bg-slate-50 border border-slate-200 px-4 py-3 text-right">
                <div className="text-xs text-slate-500">Tempo total</div>
                <div className="text-2xl font-bold text-slate-900">{(Number(tempoProducaoDraft.pesponto) || 0) + (Number(tempoProducaoDraft.montagem) || 0)} dias</div>
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
    </PageShell>
  );

  const renderSugestoes = () => {
    const tempoTotal = (Number(tempoProducao?.pesponto) || 0) + (Number(tempoProducao?.montagem) || 0);
    const analise = rows.map((row) => {
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

    const coberturaAnalise = rows
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

    const renderBlocoProgramacao = (programacao, corTag) => (
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
                        dia.fichas.map((ficha, idx) => (
                          <div key={`${programacao.tipo}-${dia.dia}-${ficha.nome}-${idx}`} className="rounded-2xl border border-slate-200 p-4 flex items-center justify-between gap-4">
                            <div>
                              <div className="text-xs uppercase tracking-wide text-slate-500">Ordem {String(idx + 1).padStart(2, "0")}</div>
                              <div className="font-semibold text-slate-900 mt-1">{ficha.cor}</div>
                              <div className="text-sm text-slate-500 mt-1">{ficha.ref} • {ficha.nome}</div>
                            </div>
                            <div className="flex items-center gap-3">
                              <div className="text-right">
                                <div className="text-xs text-slate-500">Prioridade</div>
                                <div className="font-semibold text-slate-900">{Math.round(ficha.prioridade || 0)}</div>
                              </div>
                              <div className="px-3 py-1.5 rounded-xl bg-slate-950 text-white text-sm font-semibold">{ficha.total} pares</div>
                              <button onClick={() => setPreviewFicha(ficha)} className="px-3 py-1.5 text-xs font-semibold rounded-xl bg-[#FCECEE] text-[#8B1E2D] border border-[#E7C7CC] hover:bg-[#F7DDE1]">Visualizar</button>
                            </div>
                          </div>
                        ))
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

    const subAbas = [
      {
        key: "Pesponto",
        titulo: "Programação de Pesponto",
        descricao: "Veja apenas a programação do pesponto no período selecionado.",
        programacao: programacaoPesponto,
        corTag: "bg-amber-100 text-amber-700 border-amber-200",
      },
      {
        key: "Montagem",
        titulo: "Programação de Montagem",
        descricao: "Veja apenas a programação da montagem no período selecionado.",
        programacao: programacaoMontagem,
        corTag: "bg-sky-100 text-[#8B1E2D] border-sky-200",
      },
    ];

    const subAbaAtiva = subAbas.find((item) => item.key === programacaoSubAba) || subAbas[0];

    return (
      <PageShell title="Programação do Dia" subtitle="Defina quantos dias quer programar. O sistema monta períodos separados de Montagem e Pesponto para antecipar matérias-primas e execução.">
        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm p-6">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <div className="font-bold text-lg">Período da programação</div>
              <div className="text-sm text-slate-500 mt-1">Escolha quantos dias quer planejar: 1, 3, 7, 15 ou outro período personalizado.</div>
            </div>
            <div className="flex flex-col gap-3 lg:items-end">
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
              <div className="font-bold text-lg">Setor da programação</div>
              <div className="text-sm text-slate-500 mt-1">Use as sub abas para alternar entre Pesponto e Montagem sem poluir a tela.</div>
            </div>
            <div className="flex flex-wrap gap-2">
              {subAbas.map((aba) => (
                <button
                  key={aba.key}
                  type="button"
                  onClick={() => setProgramacaoSubAba(aba.key)}
                  className={`px-4 py-2.5 rounded-2xl text-sm font-semibold border transition ${programacaoSubAba === aba.key ? "bg-[#8B1E2D] text-white border-[#8B1E2D]" : "bg-[#FFF7F8] text-[#0F172A] border-slate-200 hover:bg-white"}`}
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

        {renderBlocoProgramacao(subAbaAtiva.programacao, subAbaAtiva.corTag)}
      </PageShell>
    );
  };

  const renderRelatorioProducao = () => {
    const handlePrintRelatorio = (linhasRelatorio) => {
      const conteudo = document.createElement("div");

      const totalPares = linhasRelatorio.reduce((acc, item) => acc + (item.total || 0), 0);
      const emAberto = linhasRelatorio.filter((item) => item.status === "Em aberto").length;
      const finalizados = linhasRelatorio.filter((item) => item.status === "Finalizado").length;

      const dataGeracao = new Date().toLocaleString("pt-BR");

      conteudo.innerHTML = `
        <div style="font-family: Arial; padding:20px;">
          <h1 style="margin:0;">ROCK STAR</h1>
          <h2 style="margin:0;">RELATÓRIO DE PRODUÇÃO</h2>
          <p>Gerado em: ${dataGeracao}</p>

          <hr/>

          <h3>Resumo</h3>
          <p><b>Programações:</b> ${linhasRelatorio.length}</p>
          <p><b>Total de pares:</b> ${totalPares}</p>
          <p><b>Em aberto:</b> ${emAberto}</p>
          <p><b>Finalizados:</b> ${finalizados}</p>

          <hr/>

          ${["Pesponto", "Montagem"].map(setor => {
            const linhas = linhasRelatorio.filter(item => item.setor === setor);

            return `
              <h3>${setor}</h3>
              <table border="1" cellspacing="0" cellpadding="5" width="100%">
                <thead>
                  <tr>
                    <th>Data</th>
                    <th>Programação</th>
                    <th>REF</th>
                    <th>Cor</th>
                    <th>Pares</th>
                    <th>Status</th>
                    <th>Detalhe</th>
                  </tr>
                </thead>
                <tbody>
                  ${linhas.map(item => {
                    const detalhe = (item.items || []).map(e => `${e.size}: ${e.qtd}`).join(" | ");
                    return `
                      <tr>
                        <td>${item.dataLancamento || "-"}</td>
                        <td>${item.programacao || "-"}</td>
                        <td>${item.ref || "-"}</td>
                        <td>${item.cor || "-"}</td>
                        <td>${item.total || 0}</td>
                        <td>${item.status || "-"}</td>
                        <td>${detalhe || "-"}</td>
                      </tr>
                    `;
                  }).join("")}
                </tbody>
              </table>
            `;
          }).join("")}

          <br/><br/>
          <div style="display:flex; justify-content:space-between;">
            <div>_________________________<br/>Responsável Produção</div>
            <div>_________________________<br/>Conferência</div>
          </div>
        </div>
      `;

      const janela = window.open("", "_blank");
      janela.document.write(conteudo.innerHTML);
      janela.document.close();
      janela.print();
    };

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
            <div className="flex items-center gap-3">
              <span className="text-sm text-slate-500">{linhas.length} linha(s)</span>
              <button
                onClick={() => handlePrintRelatorio(linhas)}
                className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold shadow-sm"
              >
                Imprimir relatório
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
    <div className="min-h-screen bg-[radial-gradient(circle_at_top,_#FDF2F4_0%,_#FFFFFF_30%,_#F8FAFC_100%)] text-slate-900">
      <style>{`
        @media print {
          body * { visibility: hidden !important; }
          #print-root, #print-root * { visibility: visible !important; }
          #print-root {
            position: absolute;
            left: 0;
            top: 0;
            width: 100%;
            background: white;
            padding: 24px;
            color: #0f172a;
          }
          #print-root .print-section { page-break-inside: avoid; }
        }
      `}</style>
      <div className="flex min-h-screen">
        <aside className="hidden xl:flex xl:w-[310px] xl:flex-col xl:justify-between xl:border-r xl:border-white/50 xl:bg-[#0F172A] xl/p-6 xl:text-white xl:shadow-[20px_0_60px_rgba(15,23,42,0.18)]">
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

        <main className="flex-1 p-4 lg:p-6 xl:p-8">
          <div className="mx-auto max-w-[1720px] space-y-4">
            <div className="xl:hidden rounded-[26px] border border-slate-200 bg-white p-4 shadow-[0_10px_30px_rgba(15,23,42,0.06)] backdrop-blur">
              <img src="/logo-rockstar.png" alt="Rock Star" className="h-8 object-contain" />
              <div className="mt-2 text-2xl font-black tracking-tight text-slate-950">Módulo Produção</div>
              <div className="mt-4 flex gap-2 overflow-x-auto pb-1">
                {navItems.map((item) => (
                  <button
                    key={item}
                    onClick={() => setActive(item)}
                    className={`whitespace-nowrap rounded-2xl border px-4 py-2.5 text-sm font-semibold transition ${
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

      {previewFicha && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 overflow-auto">
          <div className="w-full max-w-3xl rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
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
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4">
          <div className="w-full max-w-md rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
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
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4">
            <div className="w-full max-w-lg rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
              <div className="text-lg font-bold">Lançamento fora da regra</div>
              <p className="text-sm text-slate-600 mt-3 leading-relaxed">
                {mensagens.join(" ")}
              </p>
              <p className="text-sm text-slate-600 mt-3 leading-relaxed">Deseja realmente continuar com esse lançamento?</p>
              <div className="mt-6 flex gap-3 justify-end">
                <button onClick={() => setConfirmMov(null)} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold bg-white">Cancelar</button>
                <button onClick={() => executeMov(confirmMov.tipo, confirmMov.form, true)} className="rounded-2xl bg-slate-950 text-white px-4 py-3 text-sm font-semibold">Lançar mesmo assim</button>
              </div>
            </div>
          </div>
        );
      })()}

      {confirmAction && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4">
          <div className="w-full max-w-lg rounded-[28px] bg-white shadow-2xl border border-slate-200 p-6">
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

      {printRelatorioData && (
        <div id="print-root" className="hidden print:block">
          <div className="flex items-start justify-between gap-4 border-b-2 border-slate-900 pb-3 print-section">
            <div className="text-xl font-bold tracking-[0.08em]">ROCK STAR</div>
            <div className="text-center flex-1">
              <div className="text-2xl font-bold">RELATÓRIO DE PRODUÇÃO</div>
              <div className="text-xs text-slate-500 mt-1">Gerado pelo Módulo Produção</div>
            </div>
            <div className="text-xs text-right min-w-[180px]">
              <div><strong>Gerado em:</strong> {printRelatorioData.dataGeracao}</div>
            </div>
          </div>

          <div className="mt-4 border border-slate-300 bg-slate-50 p-3 text-xs print-section">
            <div><strong>Filtros aplicados:</strong></div>
            <div className="mt-1">
              Período: {printRelatorioData.filtros.periodo} | REF: {printRelatorioData.filtros.ref} | Cor: {printRelatorioData.filtros.cor} | Setor: {printRelatorioData.filtros.setor} | Status: {printRelatorioData.filtros.status}
            </div>
          </div>

          <div className="mt-4 grid grid-cols-4 gap-3 print-section">
            <div className="border border-slate-300 p-3 text-center">
              <div className="text-sm">Programações</div>
              <div className="text-2xl font-bold mt-1">{printRelatorioData.resumo.programacoes}</div>
            </div>
            <div className="border border-slate-300 p-3 text-center">
              <div className="text-sm">Total de pares</div>
              <div className="text-2xl font-bold mt-1">{printRelatorioData.resumo.totalPares}</div>
            </div>
            <div className="border border-slate-300 p-3 text-center">
              <div className="text-sm">Em aberto</div>
              <div className="text-2xl font-bold mt-1">{printRelatorioData.resumo.emAberto}</div>
            </div>
            <div className="border border-slate-300 p-3 text-center">
              <div className="text-sm">Finalizados</div>
              <div className="text-2xl font-bold mt-1">{printRelatorioData.resumo.finalizados}</div>
            </div>
          </div>

          {printRelatorioData.gruposPorSetor.map((grupo) => (
            <div key={`print-${grupo.setor}`} className="mt-6 print-section">
              <div className="text-base font-bold border-b border-slate-300 pb-2">{grupo.setor.toUpperCase()}</div>
              {grupo.linhas.length === 0 ? (
                <div className="mt-3 p-4 border border-dashed border-slate-300 text-sm text-slate-500 text-center">Sem registros para este setor no período selecionado.</div>
              ) : (
                <table className="w-full border-collapse mt-3 text-[11px]">
                  <thead>
                    <tr className="bg-slate-50">
                      <th className="border border-slate-300 px-2 py-2 text-left">Data</th>
                      <th className="border border-slate-300 px-2 py-2 text-left">Programação</th>
                      <th className="border border-slate-300 px-2 py-2 text-left">REF</th>
                      <th className="border border-slate-300 px-2 py-2 text-left">Cor</th>
                      <th className="border border-slate-300 px-2 py-2 text-center">Pares</th>
                      <th className="border border-slate-300 px-2 py-2 text-center">Status</th>
                      <th className="border border-slate-300 px-2 py-2 text-left">Detalhe</th>
                    </tr>
                  </thead>
                  <tbody>
                    {grupo.linhas.map((item) => {
                      const dataRef = item.status === "Finalizado" && item.dataFinalizacao ? item.dataFinalizacao : item.dataLancamento;
                      return (
                        <tr key={`print-${grupo.setor}-${item.id}`}>
                          <td className="border border-slate-300 px-2 py-2">{dataRef || "-"}</td>
                          <td className="border border-slate-300 px-2 py-2">{item.programacao || "-"}</td>
                          <td className="border border-slate-300 px-2 py-2">{item.ref || "-"}</td>
                          <td className="border border-slate-300 px-2 py-2">{item.cor || "-"}</td>
                          <td className="border border-slate-300 px-2 py-2 text-center">{item.total || 0}</td>
                          <td className="border border-slate-300 px-2 py-2 text-center">{item.status || "-"}</td>
                          <td className="border border-slate-300 px-2 py-2">{(item.items || []).map((entry) => `${entry.size}: ${entry.qtd}`).join(" | ") || "-"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          ))}

          <div className="mt-10 flex justify-between gap-6 print-section">
            <div className="flex-1 text-center pt-8 border-t border-slate-900 text-xs">Responsável Produção</div>
            <div className="flex-1 text-center pt-8 border-t border-slate-900 text-xs">Conferência</div>
          </div>
        </div>
      )}
    </div>
  );
}
