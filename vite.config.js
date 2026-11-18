import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

function gcmApiDev() {
  return {
    name: "gcm-api-dev",
    configureServer(server) {
      server.middlewares.use("/api/gcm/pull-estoque", async (req, res) => {
        if (req.method !== "POST") {
          res.statusCode = 405;
          res.setHeader("content-type", "application/json");
          res.end(JSON.stringify({ error: "Method not allowed" }));
          return;
        }
        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        let body = {};
        try {
          const raw = Buffer.concat(chunks).toString("utf8");
          body = raw ? JSON.parse(raw) : {};
        } catch {
          body = {};
        }
        try {
          const mod = await server.ssrLoadModule("/api/gcm/pull-estoque.js");
          const data = await mod.pullEstoqueGcm(body);
          res.statusCode = 200;
          res.setHeader("content-type", "application/json");
          res.end(JSON.stringify(data));
        } catch (error) {
          res.statusCode = error.statusCode || 500;
          res.setHeader("content-type", "application/json");
          res.end(JSON.stringify({ error: error.message || "Falha ao puxar estoque do GCM" }));
        }
      });
    },
  };
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  if (env.MPLUS_INTEGRATION_KEY) {
    process.env.MPLUS_INTEGRATION_KEY = env.MPLUS_INTEGRATION_KEY;
  }
  return {
    plugins: [react(), tailwindcss(), gcmApiDev()],
  };
});
