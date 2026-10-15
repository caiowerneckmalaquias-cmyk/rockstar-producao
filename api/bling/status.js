import { getConexao, json } from "../_lib/bling.js";

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const row = await getConexao();
    if (!row?.access_token) {
      return json(res, 200, {
        connected: false,
        depositoPaId: null,
        depositoEstId: null,
        depositoMultiId: null,
        depositoPaNome: null,
        depositoEstNome: null,
        depositoMultiNome: null,
        expiresAt: null,
        connectedAt: null,
        stockOffset: Number(process.env.BLING_STOCK_OFFSET) || 1000,
        phase2Note: "Sugestões usam DISP (Multiempresa). Esteira/push usam PA.",
      });
    }

    const expiresAt = row.expires_at ? new Date(row.expires_at).getTime() : 0;
    const expired = expiresAt > 0 && expiresAt < Date.now();

    return json(res, 200, {
      connected: true,
      expired,
      depositoPaId: row.deposito_pa_id ?? null,
      depositoEstId: row.deposito_est_id ?? null,
      depositoMultiId: row.deposito_multi_id ?? null,
      depositoPaNome: row.deposito_pa_nome ?? null,
      depositoEstNome: row.deposito_est_nome ?? null,
      depositoMultiNome: row.deposito_multi_nome ?? null,
      expiresAt: row.expires_at ?? null,
      connectedAt: row.connected_at ?? row.updated_at ?? null,
      stockOffset: Number(process.env.BLING_STOCK_OFFSET) || 1000,
      phase2Note: "Sugestões usam DISP (Multiempresa). Esteira/push usam PA.",
    });
  } catch (e) {
    console.error("bling status", e);
    return json(res, 500, {
      connected: false,
      error: e.message || "Erro ao ler status Bling",
    });
  }
}
