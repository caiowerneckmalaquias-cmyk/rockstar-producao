-- Rodar no SQL Editor do Supabase.
-- Multiempresa = disponível pra venda (Sugestões/Programação).
-- PA continua sendo fábrica / esteira / push Bling.

alter table public.estoque
  add column if not exists disp integer not null default 0;

comment on column public.estoque.disp is
  'Disponível pra venda (saldo depósito Multiempresa do Bling). Usado nas Sugestões.';

alter table public.bling_conexao
  add column if not exists deposito_multi_id bigint,
  add column if not exists deposito_multi_nome text;

comment on column public.bling_conexao.deposito_multi_id is
  'ID do depósito virtual Multiempresa no Bling.';
comment on column public.bling_conexao.deposito_multi_nome is
  'Nome do depósito Multiempresa vinculado.';

-- Preenche disp com PA atual onde ainda estiver zerado (fallback até o 1º pull).
update public.estoque
set disp = coalesce(pa, 0)
where coalesce(disp, 0) = 0
  and coalesce(pa, 0) <> 0;
