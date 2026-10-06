# Rock Star Produção — pronto para Vercel

## Rodar localmente
```bash
npm install
npm run dev
```

## Build de produção
```bash
npm run build
```

## Subir na Vercel
- Envie esta pasta para um repositório GitHub
- Importe o repositório na Vercel
- A Vercel deve detectar **Vite** automaticamente
- Se pedir configuração manual:
  - Build Command: `npm run build`
  - Output Directory: `dist`

## Integração Bling (fase 1 — OAuth + fase 2 — sync)
1. No Supabase SQL Editor, rode [`scripts/bling-schema.sql`](scripts/bling-schema.sql) e [`scripts/bling-fase2-schema.sql`](scripts/bling-fase2-schema.sql).
2. Na Vercel → Environment Variables, configure conforme [`.env.example`](.env.example):
   - `BLING_CLIENT_ID`, `BLING_CLIENT_SECRET`
   - `BLING_REDIRECT_URI=https://rockstar-producao.vercel.app/bling/callback`
   - `APP_URL=https://rockstar-producao.vercel.app`
   - `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`
   - `BLING_STOCK_OFFSET=1000` (opcional)
3. No app Bling, o link de redirecionamento deve ser exatamente o `BLING_REDIRECT_URI` acima.
4. Em **Importar GCM** → **Conectar Bling**. Depósitos `PRODUTO ACABADO` (PA) e `OVERLOQUE` (EST) são vinculados automaticamente.
5. Cadastre **siglas** (Controle Geral ou Nova referência `COR|SIGLA`). SKU = `REF+SIGLA+tamanho` (ex.: `TNCV010RS34`).
6. **Puxar estoque** (−1000) / **Subir estoque** (+1000). Se Bling ≤ 1000, zera no Bling.

## Arquivos principais
- `src/App.jsx` → seu código enviado
- `src/main.jsx` → ponto de entrada React
- `src/styles.css` → Tailwind v4 + estilos base
- `vite.config.js` → Vite + React + Tailwind
- `api/bling/*` → OAuth Bling (serverless Vercel)
