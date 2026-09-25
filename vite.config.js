import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [
    react(),
    {
      name: "local-api-dev-middleware",
      configureServer(server) {
        let memorySheets = [
          {
            id: "default-batch-28",
            title: "Batch 28 Sheet (Default)",
            url: "https://docs.google.com/spreadsheets/d/1d3d2CF4HzIpDyz2I4QRf4zeNGLH7wL2H-mOUPS8BQww/edit?gid=0#gid=0",
            isDefault: true,
            createdAt: new Date().toISOString()
          }
        ];

        server.middlewares.use(async (req, res, next) => {
          if (req.url.startsWith("/api/fetch-sheet")) {
            const urlObj = new URL(req.url, "http://localhost");
            const targetUrl = urlObj.searchParams.get("url");
            if (!targetUrl) {
              res.statusCode = 400;
              res.setHeader("Content-Type", "application/json");
              return res.end(JSON.stringify({ error: "Missing url parameter" }));
            }
            try {
              let fetchUrl = String(targetUrl).trim();
              const idMatch = fetchUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
              if (idMatch) {
                const sheetId = idMatch[1];
                let gid = "0";
                const gidMatch = fetchUrl.match(/[?&#]gid=([0-9]+)/);
                if (gidMatch) gid = gidMatch[1];
                fetchUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
              }

              const fetchRes = await fetch(fetchUrl, {
                headers: {
                  "User-Agent":
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                  Accept: "text/csv,text/plain,*/*"
                }
              });

              if (!fetchRes.ok) {
                res.statusCode = fetchRes.status;
                res.setHeader("Content-Type", "application/json");
                return res.end(JSON.stringify({ error: `HTTP ${fetchRes.status}: ${fetchRes.statusText}` }));
              }

              const text = await fetchRes.text();
              res.statusCode = 200;
              res.setHeader("Access-Control-Allow-Origin", "*");
              res.setHeader("Content-Type", "text/csv; charset=utf-8");
              return res.end(text);
            } catch (err) {
              res.statusCode = 500;
              res.setHeader("Content-Type", "application/json");
              return res.end(JSON.stringify({ error: err.message }));
            }
          }

          if (req.url.startsWith("/api/sheets")) {
            res.setHeader("Access-Control-Allow-Origin", "*");
            res.setHeader("Content-Type", "application/json");

            if (req.method === "GET") {
              return res.end(JSON.stringify({ sheets: memorySheets, storage: "local-dev" }));
            }

            if (req.method === "POST") {
              let body = "";
              req.on("data", (chunk) => { body += chunk; });
              req.on("end", () => {
                try {
                  const parsed = JSON.parse(body || "{}");
                  if (Array.isArray(parsed.sheets)) {
                    memorySheets = parsed.sheets;
                  }
                  return res.end(JSON.stringify({ success: true, sheets: memorySheets, storage: "local-dev" }));
                } catch (e) {
                  res.statusCode = 400;
                  return res.end(JSON.stringify({ error: "Invalid JSON" }));
                }
              });
              return;
            }
          }

          next();
        });
      }
    }
  ]
});
