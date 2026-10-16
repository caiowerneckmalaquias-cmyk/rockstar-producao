import {
  BLING_SECUNDARIA_SLOTS,
  getConexao,
  json,
  listContasSecundarias,
} from "../_lib/bling.js";

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const row = await getConexao();
    let secundarias = [];
    try {
      const rows = await listContasSecundarias();
      const bySlot = new Map((rows || []).map((r) => [Number(r.slot), r]));
      secundarias = BLING_SECUNDARIA_SLOTS.map((slot) => {
        const r = bySlot.get(slot);
        const expiresAt = r?.expires_at ? new Date(r.expires_at).getTime() : 0;
        const expired = expiresAt > 0 && expiresAt < Date.now();
        return {
          slot,
          label: r?.label || `Conta secundária ${slot}`,
          connected: !!r?.access_token,
          expired: !!r?.access_token && expired,
          depositoGeralId: r?.deposito_geral_id ?? null,
          depositoGeralNome: r?.deposito_geral_nome ?? null,
          connectedAt: r?.connected_at ?? r?.updated_at ?? null,
        };
      });
    } catch (e) {
      // Tabela ainda não criada no Supabase
      secundarias = BLING_SECUNDARIA_SLOTS.map((slot) => ({
        slot,
        label: `Conta secundária ${slot}`,
        connected: false,
        expired: false,
        depositoGeralId: null,
        depositoGeralNome: null,
        connectedAt: null,
        error: e.message,
      }));
    }

    const secundariasProntas = secundarias.filter(
      (s) => s.connected && !s.expired && s.depositoGeralId != null
    ).length;

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
        secundarias,
        secundariasProntas,
        phase2Note:
          "DISP = PA principal + Geral das contas secundárias. Esteira/push usam PA.",
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
      secundarias,
      secundariasProntas,
      phase2Note:
        "DISP = PA principal + Geral das contas secundárias. Esteira/push usam PA.",
    });
  } catch (e) {
    console.error("bling status", e);
    return json(res, 500, {
      connected: false,
      error: e.message || "Erro ao ler status Bling",
    });
  }
}
