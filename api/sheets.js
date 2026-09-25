const DEFAULT_SHEETS = [
  {
    id: "default-batch-28",
    title: "Batch 28 Sheet (Default)",
    url: "https://docs.google.com/spreadsheets/d/1d3d2CF4HzIpDyz2I4QRf4zeNGLH7wL2H-mOUPS8BQww/edit?gid=0#gid=0",
    isDefault: true,
    createdAt: new Date().toISOString()
  }
];

let memorySheets = [...DEFAULT_SHEETS];

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,DELETE,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version"
  );

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const hasKv = Boolean(process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN);
  const hasPostgres = Boolean(process.env.POSTGRES_URL || process.env.DATABASE_URL);

  // Helper for Upstash / Vercel KV REST
  async function kvCommand(cmd, ...args) {
    const url = `${process.env.KV_REST_API_URL}/${cmd}/${args.map(encodeURIComponent).join("/")}`;
    const r = await fetch(url, {
      headers: { Authorization: `Bearer ${process.env.KV_REST_API_TOKEN}` }
    });
    return r.json();
  }

  // GET: retrieve all saved sheet links
  if (req.method === "GET") {
    try {
      if (hasKv) {
        const data = await kvCommand("get", "google_sheets_list");
        let list = DEFAULT_SHEETS;
        if (data && data.result) {
          try {
            list = JSON.parse(data.result);
          } catch (e) {
            list = DEFAULT_SHEETS;
          }
        }
        return res.status(200).json({ sheets: list, storage: "kv" });
      }

      if (hasPostgres) {
        // Dynamic import to avoid crash if pg not installed
        try {
          const { sql } = await import("@vercel/postgres");
          await sql`CREATE TABLE IF NOT EXISTS google_sheets (
            id TEXT PRIMARY KEY,
            title TEXT,
            url TEXT,
            is_default BOOLEAN,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
          );`;
          const { rows } = await sql`SELECT id, title, url, is_default as "isDefault", created_at as "createdAt" FROM google_sheets ORDER BY created_at ASC;`;
          if (rows.length === 0) {
            // Seed default
            await sql`INSERT INTO google_sheets (id, title, url, is_default) VALUES ('default-batch-28', 'Batch 28 Sheet (Default)', 'https://docs.google.com/spreadsheets/d/1d3d2CF4HzIpDyz2I4QRf4zeNGLH7wL2H-mOUPS8BQww/edit?gid=0#gid=0', true) ON CONFLICT DO NOTHING;`;
            return res.status(200).json({ sheets: DEFAULT_SHEETS, storage: "postgres" });
          }
          return res.status(200).json({ sheets: rows, storage: "postgres" });
        } catch (dbErr) {
          console.warn("Postgres error, fallback to memory:", dbErr.message);
        }
      }

      return res.status(200).json({ sheets: memorySheets, storage: "local-sync" });
    } catch (err) {
      return res.status(500).json({ error: err.message, sheets: memorySheets, storage: "local-sync" });
    }
  }

  // POST: save/update sheet list
  if (req.method === "POST") {
    try {
      const body = req.body || {};
      const { sheets } = body;

      if (!Array.isArray(sheets)) {
        return res.status(400).json({ error: "Expected an array of 'sheets'" });
      }

      if (hasKv) {
        await kvCommand("set", "google_sheets_list", JSON.stringify(sheets));
        return res.status(200).json({ success: true, sheets, storage: "kv" });
      }

      if (hasPostgres) {
        try {
          const { sql } = await import("@vercel/postgres");
          await sql`CREATE TABLE IF NOT EXISTS google_sheets (
            id TEXT PRIMARY KEY,
            title TEXT,
            url TEXT,
            is_default BOOLEAN,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
          );`;
          await sql`DELETE FROM google_sheets;`;
          for (const s of sheets) {
            await sql`INSERT INTO google_sheets (id, title, url, is_default) VALUES (${s.id}, ${s.title || ""}, ${s.url}, ${Boolean(s.isDefault)});`;
          }
          return res.status(200).json({ success: true, sheets, storage: "postgres" });
        } catch (dbErr) {
          console.warn("Postgres save error, falling back to memory:", dbErr.message);
        }
      }

      memorySheets = sheets;
      return res.status(200).json({ success: true, sheets: memorySheets, storage: "local-sync" });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
}
