// ==================== FORMATTING HELPERS ====================
export const formatNumber = (num) => (Number(num) || 0).toLocaleString('en-US');

export const formatSize = (bytes) => {
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  return `${(bytes / 1024).toFixed(1)} KB`;
};

// 1 -> "A", 27 -> "AA"
export const getColumnLetter = (num) => {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result || 'A';
};

// "A" -> 1, "AA" -> 27
export const getColumnIndex = (letters) => {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
};

const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// Parse "202602" (or 202602 as number) into { year: 2026, month: 2 }.
export const parseProductionMonth = (value) => {
  const s = String(value ?? '').trim();
  const m = s.match(/^(\d{4})(\d{2})$/);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  if (month < 1 || month > 12) return null;
  return { year, month };
};

// { year: 2026, month: 2 } + offset 1 -> "Mar'26"
export const formatMonth = (pm, offset = 0) => {
  if (!pm) return '';
  const idx = pm.month - 1 + offset;
  const year = pm.year + Math.floor(idx / 12);
  const month = ((idx % 12) + 12) % 12;
  return `${MONTH_SHORT[month]}'${String(year).slice(-2)}`;
};

// Labels for N..N+3 given the production month (null -> plain N labels).
export const monthLabels = (pm) =>
  ['N', 'N+1', 'N+2', 'N+3'].map((h, i) => (pm ? `${h} (${formatMonth(pm, i)})` : h));

export const todayStamp = () => new Date().toISOString().slice(0, 10);

// Excel sheet names: max 31 chars, none of []*?/\: and no leading/trailing apostrophe.
export const safeSheetName = (name) =>
  String(name ?? '')
    .replace(/[[\]*?/\\:]/g, ' ')
    .replace(/^'+|'+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 31)
    .trim() || 'Sheet';
