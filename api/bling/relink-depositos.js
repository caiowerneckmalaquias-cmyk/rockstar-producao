import {
  getValidAccessToken,
  json,
  listDepositos,
  matchDepositos,
  summarizeDepositos,
  upsertConexao,
} from "../_lib/bling.js";

/**
 * Religa PA / EST / Multiempresa com o token atual (sem OAuth).
 * Devolve a lista bruta de depósitos para debug na UI.
 */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return json(res, 405, { error: "Method not allowed" });
  }

  try {
    const { accessToken, conexao } = await getValidAccessToken();

    let depositos = [];
    try {
      depositos = await listDepositos(accessToken);
    } catch (e) {
      return json(res, 502, {
        error: `Falha ao listar depósitos no Bling: ${e.message || e}`,
        depositos: [],
      });
    }

    const list = Array.isArray(depositos) ? depositos : [];
    const matched = matchDepositos(list);
    const summary = summarizeDepositos(list);

    const updated = await upsertConexao({
      access_token: conexao.access_token,
      refresh_token: conexao.refresh_token,
      token_type: conexao.token_type || "Bearer",
      expires_at: conexao.expires_at,
      scopes: conexao.scopes,
      deposito_pa_id: matched.depositoPaId,
      deposito_est_id: matched.depositoEstId,
      deposito_multi_id: matched.depositoMultiId,
      deposito_pa_nome: matched.depositoPaNome,
      deposito_est_nome: matched.depositoEstNome,
      deposito_multi_nome: matched.depositoMultiNome,
      connected_at: conexao.connected_at,
    });

    return json(res, 200, {
      ok: true,
      matched: {
        depositoPaId: matched.depositoPaId,
        depositoEstId: matched.depositoEstId,
        depositoMultiId: matched.depositoMultiId,
        depositoPaNome: matched.depositoPaNome,
        depositoEstNome: matched.depositoEstNome,
        depositoMultiNome: matched.depositoMultiNome,
      },
      depositos: summary,
      multiVinculado: matched.depositoMultiId != null,
      warning:
        matched.depositoMultiId != null
          ? null
          : summary.length
            ? `Multiempresa não encontrado entre ${summary.length} depósito(s) da API. Nomes: ${summary
                .map((d) => d.nome)
                .join(", ")}`
            : "API /depositos retornou lista vazia. Verifique o escopo do app Bling.",
      conexaoId: updated?.id ?? 1,
    });
  } catch (e) {
    console.error("bling relink-depositos", e);
    return json(res, 500, { error: e.message || "Falha ao religar depósitos" });
  }
}
