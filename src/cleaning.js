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
