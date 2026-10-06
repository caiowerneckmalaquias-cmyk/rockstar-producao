-- Rodar uma vez no SQL Editor do Supabase (projeto hnjawylcwfhxjhkvtyyc).
-- Guarda tokens OAuth e IDs dos depósitos PA / EST.

create table if not exists public.bling_conexao (
  id integer primary key default 1 check (id = 1),
  access_token text,
  refresh_token text,
  token_type text default 'Bearer',
  expires_at timestamptz,
  scopes text,
  deposito_pa_id bigint,
  deposito_est_id bigint,
  deposito_pa_nome text,
  deposito_est_nome text,
  connected_at timestamptz,
  updated_at timestamptz default now()
);

comment on table public.bling_conexao is
  'Conexão OAuth Bling (singleton id=1). PA=PRODUTO ACABADO, EST=OVERLOQUE.';

alter table public.bling_conexao enable row level security;

-- Sem policies para anon/authenticated: só service_role (API Vercel) acessa.
drop policy if exists bling_conexao_deny_all on public.bling_conexao;

grant select, insert, update, delete on public.bling_conexao to service_role;
