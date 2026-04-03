import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");

const src =
  process.argv[2] ||
  "C:/Users/cwern/Downloads/Supabase Snippet Movimentações do dia anterior por tipo.csv";
const dest =
  process.argv[3] || path.join(root, "movimentacoes_import_supabase.csv");

const raw = fs.readFileSync(src, "utf8");
const lines = raw.split(/\r\n|\n|\r/).filter((l) => l.trim().length);

const out = lines.map((line) => {
  let s = line.replace(/^[^,]+,/, "");
  s = s.replace(/,null\s*$/i, ",");
  return s;
});

fs.writeFileSync(dest, out.join("\n") + "\n", "utf8");
console.log(`OK: ${lines.length} linhas -> ${dest}`);
