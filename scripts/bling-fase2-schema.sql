-- Fase 2: sigla Bling por cor (SKU = REF + SIGLA + TAMANHO).
-- Rodar no SQL Editor do Supabase após scripts/bling-schema.sql.

alter table public.produtos
  add column if not exists bling_sigla text;

comment on column public.produtos.bling_sigla is
  'Sigla da cor no SKU Bling (ex.: RS em TNCV010RS34).';
