export const sizes = [34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44];

export const navItems = [
  "Dashboard",
  "Controle Geral",
  "Importar GCM",
  "Pesponto",
  "Montagem",
  "Costura Pronta",
  "Minimos",
  "Vendas",
  "Sugestoes",
  "Gerador de Fichas",
  "Programação do Dia",
  "Relatório de Produção",
];

export const initialRows = [
  {
    ref: "BTCV010",
    cor: "Preto",
    data: {
      34: { pa: 10, est: 20, m: 5, p: 15 },
      35: { pa: 12, est: 10, m: 3, p: 8 },
      36: { pa: 4, est: 18, m: 6, p: 12 },
      37: { pa: 2, est: 6, m: 3, p: 9 },
      38: { pa: 1, est: 0, m: 0, p: 12 },
      39: { pa: 8, est: 12, m: 6, p: 0 },
      40: { pa: 15, est: 8, m: 0, p: 0 },
      41: { pa: 14, est: 12, m: 0, p: 0 },
      42: { pa: 16, est: 0, m: 0, p: 12 },
      43: { pa: 9, est: 0, m: 12, p: 0 },
      44: { pa: 5, est: 0, m: 0, p: 12 },
    },
  },
  {
    ref: "TNCV010",
    cor: "Branco",
    data: {
      34: { pa: 18, est: 24, m: 0, p: 0 },
      35: { pa: 9, est: 12, m: 0, p: 12 },
      36: { pa: 6, est: 0, m: 12, p: 12 },
      37: { pa: 3, est: 0, m: 0, p: 12 },
      38: { pa: 14, est: 12, m: 12, p: 0 },
      39: { pa: 10, est: 0, m: 0, p: 12 },
      40: { pa: 22, est: 24, m: 0, p: 0 },
      41: { pa: 18, est: 12, m: 0, p: 0 },
      42: { pa: 6, est: 12, m: 12, p: 0 },
      43: { pa: 3, est: 0, m: 0, p: 12 },
      44: { pa: 12, est: 12, m: 0, p: 0 },
    },
  },
  {
    ref: "CRVTNCV",
    cor: "All Black",
    data: {
      34: { pa: 6, est: 24, m: 12, p: 0 },
      35: { pa: 7, est: 12, m: 12, p: 0 },
      36: { pa: 4, est: 12, m: 0, p: 12 },
      37: { pa: 2, est: 0, m: 12, p: 12 },
      38: { pa: 9, est: 0, m: 0, p: 12 },
      39: { pa: 12, est: 12, m: 0, p: 0 },
      40: { pa: 18, est: 12, m: 0, p: 0 },
      41: { pa: 14, est: 0, m: 12, p: 0 },
      42: { pa: 8, est: 0, m: 12, p: 0 },
      43: { pa: 6, est: 12, m: 0, p: 0 },
      44: { pa: 4, est: 0, m: 0, p: 12 },
    },
  },
];

export const initialMinimos = {
  BTCV010: {
    Preto: Object.fromEntries(sizes.map((s) => [s, { pa: s <= 38 ? 12 : 8, prod: 24 }])),
  },
  TNCV010: {
    Branco: Object.fromEntries(sizes.map((s) => [s, { pa: s <= 39 ? 12 : 8, prod: 24 }])),
  },
  CRVTNCV: {
    "All Black": Object.fromEntries(sizes.map((s) => [s, { pa: 8, prod: 24 }])),
  },
};

export const initialVendas = {
  BTCV010: { Preto: Object.fromEntries(sizes.map((s) => [s, 0])) },
  TNCV010: { Branco: Object.fromEntries(sizes.map((s) => [s, 0])) },
  CRVTNCV: { "All Black": Object.fromEntries(sizes.map((s) => [s, 0])) },
};

export const initialTempoProducao = {
  pesponto: 3,
  montagem: 2,
};

export const LIMITE_PROGRAMACAO_DIA = 396;