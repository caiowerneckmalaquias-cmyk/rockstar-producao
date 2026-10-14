-- Rodar uma vez no SQL Editor do Supabase (projeto hnjawylcwfhxjhkvtyyc).
-- Etiquetas de caixa 10×6: contador sequencial único + anti-duplicidade na esteira.

create table if not exists public.etiquetas_caixa (
  contador text primary key,
  contador_n bigint not null unique,
  sku text not null,
  ref text not null,
  cor text not null,
  tamanho integer not null,
  sigla text,
  programacao text,
  origem text,
  impresso_em timestamptz not null default now(),
  bipado_em timestamptz,
  confirmado_em timestamptz
);

comment on table public.etiquetas_caixa is
  'Etiquetas impressas (contador único). bipado_em = lido na esteira; confirmado_em = lote subiu PA/Bling.';

create index if not exists etiquetas_caixa_sku_idx
  on public.etiquetas_caixa (sku);

create index if not exists etiquetas_caixa_pendente_idx
  on public.etiquetas_caixa (bipado_em)
  where bipado_em is not null and confirmado_em is null;

alter table public.etiquetas_caixa enable row level security;

drop policy if exists etiquetas_caixa_anon_all on public.etiquetas_caixa;
create policy etiquetas_caixa_anon_all
  on public.etiquetas_caixa
  for all
  to anon, authenticated
  using (true)
  with check (true);

grant select, insert, update, delete on public.etiquetas_caixa to anon, authenticated, service_role;
