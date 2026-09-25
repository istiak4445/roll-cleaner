export default async function handler(req, res) {
  // Set CORS headers
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version"
  );

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  const { url } = req.query;
  if (!url) {
    return res.status(400).json({ error: "Missing required query parameter: url" });
  }

  try {
    // Normalise Google Sheet URL if needed
    let fetchUrl = String(url).trim();
    const idMatch = fetchUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (idMatch) {
      const sheetId = idMatch[1];
      let gid = "0";
      const gidMatch = fetchUrl.match(/[?&#]gid=([0-9]+)/);
      if (gidMatch) {
        gid = gidMatch[1];
      }
      fetchUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15000);

    const response = await fetch(fetchUrl, {
      signal: controller.signal,
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        Accept: "text/csv,text/plain,*/*"
      }
    });

    clearTimeout(timeout);

    if (!response.ok) {
      return res.status(response.status).json({
        error: `Google Sheets returned HTTP ${response.status}: ${response.statusText}`
      });
    }

    const csvText = await response.text();
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    return res.status(200).send(csvText);
  } catch (error) {
    return res.status(500).json({
      error: error.name === "AbortError" ? "Request to Google Sheets timed out." : error.message
    });
  }
}
