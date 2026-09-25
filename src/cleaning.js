export const EMAIL_KEEP_START = 3;
export const EMAIL_KEEP_END = 3;
export const PHONE_KEEP_START = 3;
export const PHONE_KEEP_END = 3;

export const isBlankOrNA = (value) =>
  value == null ||
  String(value).trim() === "" ||
  String(value).trim().toUpperCase() === "N/A" ||
  String(value).trim().toUpperCase() === "NULL" ||
  String(value).trim().toUpperCase() === "UNDEFINED";

export function maskEmail(value, keepStart = EMAIL_KEEP_START, keepEnd = EMAIL_KEEP_END) {
  if (isBlankOrNA(value)) return value;
  const email = String(value);
  const at = email.indexOf("@");
  if (at <= 0 || at !== email.lastIndexOf("@")) return value;
  const local = email.slice(0, at);
  const domain = email.slice(at);
  if (!domain.includes(".") || local.length <= keepStart + keepEnd) return value;
  return `${local.slice(0, keepStart)}${"*".repeat(local.length - keepStart - keepEnd)}${local.slice(-keepEnd)}${domain}`;
}

export function maskPhone(value, keepStart = PHONE_KEEP_START, keepEnd = PHONE_KEEP_END) {
  if (isBlankOrNA(value)) return value;
  const phone = String(value).trim();
  if (!/^\d+$/.test(phone) || phone.length <= keepStart + keepEnd) return value;
  return `${phone.slice(0, keepStart)}${"*".repeat(phone.length - keepStart - keepEnd)}${phone.slice(-keepEnd)}`;
}

/**
 * Normalizes roll numbers:
 * - Strips hyphens and whitespace (e.g. 28-28508-1 -> 28285081)
 * - If 6 digits, prepends "00" to make 8 digits (e.g. 260401 -> 00260401)
 * - If less than 8 numeric digits, pads with leading zeros (e.g. 880011 -> 00880011)
 */
export function normalizeRoll(value) {
  if (value == null) return "";
  let s = String(value).replace(/[-\s]/g, "").trim();
  if (!s) return "";
  if (/^\d{6}$/.test(s)) {
    return "00" + s;
  }
  if (/^\d+$/.test(s) && s.length < 8) {
    return s.padStart(8, "0");
  }
  return s;
}

export function cleanRoll(value) {
  if (isBlankOrNA(value)) return { value, warning: null };
  const normalized = normalizeRoll(value);
  if (!/^\d+$/.test(normalized)) return { value, warning: null };
  if (normalized.length > 8) {
    return { value: normalized, warning: `Roll ${normalized} has more than 8 digits and was not truncated.` };
  }
  return { value: normalized, warning: null };
}

export function cleanMobile(val) {
  if (isBlankOrNA(val)) return "";
  let cleaned = String(val).trim();
  // Remove country prefixes like +880, 880, +88, 88
  cleaned = cleaned.replace(/^(\+?880|^\+?88)/, "");
  // Strip all non-digit characters
  cleaned = cleaned.replace(/[^0-9]/g, "");
  // If 10 digits starting with 1, 3-9, prepend 0
  if (/^[13456789]\d{9}$/.test(cleaned)) {
    cleaned = "0" + cleaned;
  }
  // Keep first 11 digits
  if (cleaned.length > 11) {
    cleaned = cleaned.slice(0, 11);
  }
  return cleaned;
}

export function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];
    if (inQuotes) {
      if (ch === "\"" && next === "\"") {
        cell += "\"";
        i++;
      } else if (ch === "\"") {
        inQuotes = false;
      } else {
        cell += ch;
      }
    } else {
      if (ch === "\"") {
        inQuotes = true;
      } else if (ch === ",") {
        row.push(cell.trim());
        cell = "";
      } else if (ch === "\n" || (ch === "\r" && next === "\n")) {
        row.push(cell.trim());
        rows.push(row);
        row = [];
        cell = "";
        if (ch === "\r") i++;
      } else {
        cell += ch;
      }
    }
  }
  if (cell || row.length > 0) {
    row.push(cell.trim());
    rows.push(row);
  }
  return rows;
}

export function toGoogleSheetCsvUrl(rawUrl) {
  if (!rawUrl) return "";
  const str = String(rawUrl).trim();
  const idMatch = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) return str;
  const sheetId = idMatch[1];
  let gid = "0";
  const gidMatch = str.match(/[?&#]gid=([0-9]+)/);
  if (gidMatch) {
    gid = gidMatch[1];
  }
  return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
}

export function detectGoogleSheetColumns(headers) {
  const normalized = headers.map((h) => String(h || "").toLowerCase().replace(/\s+/g, " ").trim());

  // WEB Roll* (Col U / Col 20 usually)
  let webRollIdx = normalized.findIndex((h) => /web\s*roll/i.test(h));
  if (webRollIdx < 0) webRollIdx = 20;

  // Secondary Phone (Col J / Col 9 usually)
  let secPhoneIdx = normalized.findIndex((h) => /secondary/i.test(h) && /(phone|mobile|contact|num)/i.test(h));
  if (secPhoneIdx < 0) secPhoneIdx = normalized.findIndex((h) => /secondary/i.test(h));
  if (secPhoneIdx < 0) secPhoneIdx = 9;

  // Primary Phone (Col I / Col 8 usually)
  let priPhoneIdx = normalized.findIndex((h, idx) => idx !== secPhoneIdx && /(^|\b)(phone|mobile|contact)(\b|$)/i.test(h));
  if (priPhoneIdx < 0) priPhoneIdx = 8;

  // Internal Roll (Col E / Col 4 usually)
  let rollIdx = normalized.findIndex((h, idx) => idx !== webRollIdx && /(^|\b)roll(\b|$)/i.test(h));
  if (rollIdx < 0) rollIdx = 4;

  // Name (Col F / Col 5 usually)
  let nameIdx = normalized.findIndex((h) => /name/i.test(h));
  if (nameIdx < 0) nameIdx = 5;

  return { webRollIdx, secPhoneIdx, priPhoneIdx, rollIdx, nameIdx };
}

/**
 * Builds a fast lookup map from Google Sheet CSV text.
 * Rule:
 *  - Primary key is WEB Roll (normalized).
 *  - Mobile is taken from Secondary Phone (Col J) first.
 *  - If Secondary Phone is blank, falls back to Primary Phone (Col I).
 */
export function buildGoogleSheetLookup(csvText, sheetSource = "") {
  const rows = parseCSV(csvText);
  if (!rows || rows.length < 2) return new Map();

  const cols = detectGoogleSheetColumns(rows[0]);
  const lookup = new Map();

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const rawWebRoll = row[cols.webRollIdx];
    const normWebRoll = normalizeRoll(rawWebRoll);

    const priPhone = cleanMobile(row[cols.priPhoneIdx]);
    const secPhone = cleanMobile(row[cols.secPhoneIdx]);
    const studentName = row[cols.nameIdx] ? String(row[cols.nameIdx]).trim() : "";

    // Secondary phone is main; fallback to primary phone if secondary is blank
    const chosenMobile = secPhone || priPhone;
    const phoneSource = secPhone ? "gsheet_secondary" : (priPhone ? "gsheet_primary" : "blank");

    const record = {
      chosenMobile,
      secPhone,
      priPhone,
      phoneSource,
      studentName,
      rawWebRoll,
      normWebRoll,
      sheetSource,
      rowNumber: r + 1
    };

    if (normWebRoll) {
      lookup.set(normWebRoll, record);
    }

    // Also index internal roll as secondary fallback
    const rawRoll = row[cols.rollIdx];
    const normRoll = normalizeRoll(rawRoll);
    if (normRoll && !lookup.has(normRoll)) {
      lookup.set(normRoll, record);
    }
  }

  return lookup;
}

export const normalizeHeader = (value) => String(value ?? "").trim().toLowerCase();

export function cleanGrid(grid) {
  const rows = grid.map((row) => [...row]);
  const headerIndex = rows.findIndex((row) => row.some((cell) => ["email", "phone", "roll", "paid amount"].includes(normalizeHeader(cell))));
  if (headerIndex < 0) return { rows, warnings: ["No supported headers were found on this sheet."], stats: { emails: 0, phones: 0, rolls: 0, removed: false } };
  const headers = rows[headerIndex].map(normalizeHeader);
  const emailIndex = headers.indexOf("email");
  const phoneIndex = headers.indexOf("phone");
  const rollIndex = headers.indexOf("roll");
  const paidIndex = headers.indexOf("paid amount");
  const warnings = [];
  const stats = { emails: 0, phones: 0, rolls: 0, removed: paidIndex >= 0 };
  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    if (emailIndex >= 0) { const next = maskEmail(rows[r][emailIndex]); if (next !== rows[r][emailIndex]) stats.emails += 1; rows[r][emailIndex] = next; }
    if (phoneIndex >= 0) { const next = maskPhone(rows[r][phoneIndex]); if (next !== rows[r][phoneIndex]) stats.phones += 1; rows[r][phoneIndex] = next; }
    if (rollIndex >= 0) {
      const result = cleanRoll(rows[r][rollIndex]);
      if (String(result.value) !== String(rows[r][rollIndex])) stats.rolls += 1;
      rows[r][rollIndex] = result.value;
      if (result.warning) warnings.push(`Row ${r + 1}: ${result.warning}`);
    }
  }
  if (paidIndex >= 0) rows.forEach((row) => row.splice(paidIndex, 1));
  return { rows, warnings, stats, headerIndex };
}

export function safeBaseName(name) {
  const cleaned = String(name || "cleaned_students").replace(/\.(xlsx|pdf|csv)$/i, "").replace(/[\\/:*?"<>|]/g, "-").trim();
  return cleaned || "cleaned_students";
}

/**
 * Formats grid for Web Portal:
 * Columns: ["roll_number", "name", "mobile", "batch"]
 * Roll: normalized to 8 digits
 * Name: stripped trailing ' undefined'
 * Mobile:
 *   If googleSheetMap has student's roll:
 *     - Takes Secondary Phone from Google Sheet first
 *     - If blank, takes Primary Phone from Google Sheet
 *   If not found in Google Sheet:
 *     - Falls back to cleaned mobile number from uploaded sheet
 */
export function formatForWebPortal(grid, defaultBatch = "", googleSheetMap = null) {
  const rows = grid.map((row) => [...row]);
  const headerIndex = rows.findIndex((row) => row.some((cell) => {
    const norm = normalizeHeader(cell);
    return norm.includes("roll") || norm.includes("id") || norm.includes("name") || norm.includes("phone") || norm.includes("mobile") || norm.includes("contact");
  }));

  if (headerIndex < 0) {
    return {
      rows: [["roll_number", "name", "mobile", "batch"]],
      warnings: ["No supported headers (roll, name, or phone) were found on this sheet."],
      stats: { emails: 0, phones: 0, rolls: 0, removed: false, matchedCount: 0, unmatchedCount: 0, secondaryCount: 0, primaryCount: 0 },
      phoneSources: [],
      headerIndex: 0
    };
  }

  const headers = rows[headerIndex].map(normalizeHeader);
  const rollColIndex = headers.findIndex((h) => h.includes("roll") || h.includes("id"));
  const nameColIndex = headers.findIndex((h) => h.includes("name"));
  const mobileColIndex = headers.findIndex((h) => h.includes("phone") || h.includes("mobile") || h.includes("contact"));

  const warnings = [];
  if (rollColIndex < 0) warnings.push("Could not find a Roll/ID column.");
  if (nameColIndex < 0) warnings.push("Could not find a Name column.");
  if (mobileColIndex < 0) warnings.push("Could not find a Phone/Mobile column.");

  const webRows = [["roll_number", "name", "mobile", "batch"]];
  const phoneSources = ["header"];
  const stats = {
    emails: 0,
    phones: 0,
    rolls: 0,
    removed: false,
    matchedCount: 0,
    unmatchedCount: 0,
    secondaryCount: 0,
    primaryCount: 0
  };

  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const rawRoll = rollColIndex >= 0 ? rows[r][rollColIndex] : "";
    const rawName = nameColIndex >= 0 ? rows[r][nameColIndex] : "";
    const rawMobile = mobileColIndex >= 0 ? rows[r][mobileColIndex] : "";

    // Normalize roll to 8 digits
    const rollVal = normalizeRoll(rawRoll);
    if (String(rollVal) !== String(rawRoll)) {
      stats.rolls += 1;
    }

    // Clean name: strip trailing " undefined" if any
    const nameVal = rawName != null ? String(rawName).replace(/\s+undefined$/i, "").trim() : "";

    // Mobile number lookup against Google Sheet
    let mobileVal = "";
    let source = "unmatched";

    const gMatch = googleSheetMap ? googleSheetMap.get(rollVal) : null;

    if (gMatch && gMatch.chosenMobile) {
      mobileVal = gMatch.chosenMobile;
      stats.matchedCount += 1;
      if (gMatch.phoneSource === "gsheet_secondary") {
        stats.secondaryCount += 1;
        source = "gsheet_secondary";
      } else {
        stats.primaryCount += 1;
        source = "gsheet_primary";
      }
    } else {
      // Fallback to uploaded sheet phone
      mobileVal = cleanMobile(rawMobile);
      stats.unmatchedCount += 1;
      source = googleSheetMap && googleSheetMap.size > 0 ? "uploaded_fallback" : "uploaded_sheet";
      if (googleSheetMap && googleSheetMap.size > 0 && rollVal) {
        warnings.push(`Row ${r + 1}: Roll ${rollVal} not found in Google Sheet(s); used uploaded sheet phone.`);
      }
    }

    phoneSources.push(source);
    webRows.push([rollVal, nameVal, mobileVal, defaultBatch]);
  }

  return {
    rows: webRows,
    warnings,
    stats,
    phoneSources,
    headerIndex: 0
  };
}
