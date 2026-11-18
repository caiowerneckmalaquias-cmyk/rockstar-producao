# AGENTS.md

## Cursor Cloud specific instructions

This is a single-product **Vite + React 18** SPA ("Rock Star Produção" — a Portuguese shoe/footwear production management app). Package manager is **npm** (`package-lock.json`). There is no backend in this repo.

### Services
- **Vite dev server** (the entire product): `npm run dev` → serves on `http://localhost:5173`. This is the only service to run for development. Commands are in `package.json` (`dev`, `build`, `preview`).
- **Supabase (hosted, remote)**: the app talks directly to a hosted Supabase project from the browser. Credentials are hardcoded in `src/supabase.js` (publishable anon key) — there is no local database to start and no `.env` needed. Data reads/writes hit that live shared project, so avoid creating throwaway test records; prefer non-destructive actions (filtering, report/PDF generation) when demonstrating functionality.

### Lint / test / build
- There is **no lint script and no test suite** in this repo (no ESLint config, no test runner). Do not expect `npm run lint`/`npm test` to exist.
- Build: `npm run build` (outputs to `dist/`). The build emits a large-chunk warning (>500 kB) — this is expected and not an error.

### Notes
- Almost all logic lives in one large file, `src/App.jsx` (~8k lines).
- `VITE_SENHA_ZERAR_EST` (optional) overrides the stock-reset password; defaults to a hardcoded value when unset.
- The empty 0-byte files `git` and `main` at the repo root are stray artifacts, not used by the build.
