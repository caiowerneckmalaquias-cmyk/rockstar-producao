-- Rodar no SQL Editor do Supabase.
-- Contas Bling secundárias (slots 1–4): depósito Geral entra no DISP.
-- DISP = PA (conta principal) + soma(Geral das secundárias).

create table if not exists public.bling_conta_secundaria (
  slot smallint primary key check (slot >= 1 and slot <= 4),
  label text,
  access_token text,
  refresh_token text,
  token_type text default 'Bearer',
  expires_at timestamptz,
  scopes text,
  deposito_geral_id bigint,
  deposito_geral_nome text,
  connected_at timestamptz,
  updated_at timestamptz default now()
);

comment on table public.bling_conta_secundaria is
  'OAuth Bling das contas secundárias (slots 1-4). Geral soma no DISP junto com PA da principal.';

comment on column public.bling_conta_secundaria.deposito_geral_id is
  'ID do depósito Geral na conta secundária.';

alter table public.bling_conta_secundaria enable row level security;

drop policy if exists bling_conta_secundaria_deny_all on public.bling_conta_secundaria;

grant select, insert, update, delete on public.bling_conta_secundaria to service_role;

-- Garante as 4 linhas de slot (vazias até conectar).
insert into public.bling_conta_secundaria (slot, label)
values
  (1, 'Conta secundária 1'),
  (2, 'Conta secundária 2'),
  (3, 'Conta secundária 3'),
  (4, 'Conta secundária 4')
on conflict (slot) do nothing;
