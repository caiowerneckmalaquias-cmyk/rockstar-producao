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

## Estoque e etiquetas
- PA, EST, DISP, Pesponto e Montagem ficam no app (Supabase).
- A sigla da cor (Controle Geral ou Nova referência `COR|SIGLA`) monta o SKU da etiqueta: `REF+SIGLA+tamanho` (ex.: `TNCV010RS34`).
- O estoque é lido e gravado no Supabase pelo navegador.

## Arquivos principais
- `src/App.jsx` → seu código enviado
- `src/main.jsx` → ponto de entrada React
- `src/styles.css` → Tailwind v4 + estilos base
- `vite.config.js` → Vite + React + Tailwind
