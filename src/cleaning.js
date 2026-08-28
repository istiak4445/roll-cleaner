export const EMAIL_KEEP_START = 3;
export const EMAIL_KEEP_END = 3;
export const PHONE_KEEP_START = 3;
export const PHONE_KEEP_END = 3;

const isBlankOrNA = (value) => value == null || String(value).trim() === "" || String(value).trim().toUpperCase() === "N/A";

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

export function cleanRoll(value) {
  if (isBlankOrNA(value)) return { value, warning: null };
  const normalized = String(value).replace(/[-\s]/g, "");
  if (!/^\d+$/.test(normalized)) return { value, warning: null };
  if (normalized.length > 8) return { value: normalized, warning: `Roll ${normalized} has more than 8 digits and was not truncated.` };
  return { value: normalized.padStart(8, "0"), warning: null };
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
  const cleaned = String(name || "cleaned_students").replace(/\.(xlsx|pdf)$/i, "").replace(/[\\/:*?"<>|]/g, "-").trim();
  return cleaned || "cleaned_students";
}

export function formatForWebPortal(grid, defaultBatch = "") {
  const rows = grid.map((row) => [...row]);
  const headerIndex = rows.findIndex((row) => row.some((cell) => {
    const norm = normalizeHeader(cell);
    return norm.includes("roll") || norm.includes("id") || norm.includes("name") || norm.includes("phone") || norm.includes("mobile") || norm.includes("contact");
  }));

  if (headerIndex < 0) {
    return {
      rows: [["roll_number", "name", "mobile", "batch"]],
      warnings: ["No supported headers (roll, name, or phone) were found on this sheet."],
      stats: { emails: 0, phones: 0, rolls: 0, removed: false },
      headerIndex: 0
    };
  }

  const headers = rows[headerIndex].map(normalizeHeader);
  const rollColIndex = headers.findIndex(h => h.includes("roll") || h.includes("id"));
  const nameColIndex = headers.findIndex(h => h.includes("name"));
  const mobileColIndex = headers.findIndex(h => h.includes("phone") || h.includes("mobile") || h.includes("contact"));

  const warnings = [];
  if (rollColIndex < 0) warnings.push("Could not find a Roll/ID column.");
  if (nameColIndex < 0) warnings.push("Could not find a Name column.");
  if (mobileColIndex < 0) warnings.push("Could not find a Phone/Mobile column.");

  const webRows = [["roll_number", "name", "mobile", "batch"]];
  const stats = { emails: 0, phones: 0, rolls: 0, removed: false };

  for (let r = headerIndex + 1; r < rows.length; r += 1) {
    const rawRoll = rollColIndex >= 0 ? rows[r][rollColIndex] : "";
    const rawName = nameColIndex >= 0 ? rows[r][nameColIndex] : "";
    const rawMobile = mobileColIndex >= 0 ? rows[r][mobileColIndex] : "";

    // Clean roll number (do not mask, pad to 8 digits)
    const cleanedRollResult = cleanRoll(rawRoll);
    let rollVal = cleanedRollResult.value;
    if (String(rollVal) !== String(rawRoll)) {
      stats.rolls += 1;
    }
    if (cleanedRollResult.warning) {
      warnings.push(`Row ${r + 1}: ${cleanedRollResult.warning}`);
    }

    // Clean name: strip trailing " undefined" if any
    const nameVal = rawName != null ? String(rawName).replace(/\s+undefined$/i, "").trim() : "";

    // Clean mobile number (do not mask)
    let mobileVal = "";
    if (rawMobile != null && String(rawMobile).trim() !== "" && String(rawMobile).trim().toUpperCase() !== "N/A") {
      let cleaned = String(rawMobile).trim();
      
      // Remove leading country code: +880, 880, +88, 88
      cleaned = cleaned.replace(/^(\+?880|^\+?88)/, "");
      
      // Strip all non-digit characters
      cleaned = cleaned.replace(/[^0-9]/g, "");
      
      // If it's a 10-digit number starting with 1, 3, 4, 5, 6, 7, 8, or 9, pad it with a leading "0"
      if (/^[13456789]\d{9}$/.test(cleaned)) {
        cleaned = "0" + cleaned;
      }
      
      // Keep only the first 11 digits
      if (cleaned.length > 11) {
        cleaned = cleaned.slice(0, 11);
      }
      
      mobileVal = cleaned;
    }

    webRows.push([
      rollVal,
      nameVal,
      mobileVal,
      defaultBatch
    ]);
  }

  return { rows: webRows, warnings, stats, headerIndex: 0 };
}

