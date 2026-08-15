/**
 * `dd-mm-yyyy` typed dates.
 *
 * The client asked for the Document Date field itself to read `dd-mm-yyyy` (2026-08-15). A native
 * `<input type="date">` cannot do that — it renders from the BROWSER LOCALE — so the field becomes a
 * text input, and everything a text input has to get right lives here, under test.
 *
 * THE STORED VALUE IS UNCHANGED. Internally the form still holds ISO `YYYY-MM-DD`, and `toSpDate`
 * still sends `M/D/YYYY` at the API boundary (gotcha #1). This module converts only at the edge, so a
 * bug here shows up as a rejected keystroke rather than a document filed under the wrong date.
 *
 * Never `new Date("2026-08-02")` anywhere in here: that constructor reads a bare ISO date as UTC, so
 * it renders as the previous day for anyone west of Greenwich — a wrong date, shown confidently.
 */

/** What a typed value means. `empty` is separate from `invalid`: a blank field is not an error yet. */
export type DateParse =
  | { state: "empty" }
  | { state: "invalid"; reason: string }
  | { state: "ok"; iso: string };

const isLeap = (y: number): boolean => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

/** Days in a month, so `31-02` and `31-04` are rejected rather than rolled silently forward. */
export function daysInMonth(month: number, year: number): number {
  if (month === 2) return isLeap(year) ? 29 : 28;
  return [4, 6, 9, 11].indexOf(month) !== -1 ? 30 : 31;
}

const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

/**
 * Parse `dd-mm-yyyy`.
 *
 * `/` and `.` are accepted as separators and normalised — people paste from Excel and from email, and
 * refusing `02/08/2026` over one character reads as the field being broken. The ORDER is never
 * guessed: it is always day, month, year. Guessing is exactly what makes `02/08` unreadable.
 */
export function parseDmy(raw: string): DateParse {
  const text = (raw ?? "").trim();
  if (text.length === 0) return { state: "empty" };

  const m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(text);
  if (!m) return { state: "invalid", reason: "Use dd-mm-yyyy, for example 02-08-2026." };

  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);

  if (month < 1 || month > 12) return { state: "invalid", reason: "Month must be between 01 and 12." };
  if (day < 1) return { state: "invalid", reason: "Day must be 01 or later." };
  const max = daysInMonth(month, year);
  if (day > max) {
    // Names the actual limit. "Invalid date" leaves someone retyping 31-02 with no idea what is wrong.
    return { state: "invalid", reason: `That month has only ${max} days.` };
  }
  return { state: "ok", iso: `${year}-${pad2(month)}-${pad2(day)}` };
}

/** ISO `YYYY-MM-DD` → `dd-mm-yyyy`. Returns "" for anything that is not a plain ISO date. */
export function isoToDmy(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((iso ?? "").trim());
  return m ? `${m[3]}-${m[2]}-${m[1]}` : "";
}

/**
 * Insert the dashes as the user types, without fighting them.
 *
 * Only ever APPENDS a separator after a complete field, and never while deleting — auto-formatting
 * that reinserts the character just removed makes backspace impossible. Anything that is not a digit
 * or a separator is dropped, so a stray letter cannot land mid-value.
 */
export function autoFormatDmy(raw: string, deleting: boolean): string {
  const cleaned = (raw ?? "").replace(/[^\d\-/.]/g, "").replace(/[/.]/g, "-");
  if (deleting) return cleaned;

  const digits = cleaned.replace(/-/g, "").slice(0, 8);
  if (digits.length === 0) return "";
  const parts: string[] = [digits.slice(0, 2)];
  if (digits.length > 2) parts.push(digits.slice(2, 4));
  if (digits.length > 4) parts.push(digits.slice(4, 8));
  let out = parts.join("-");
  // A trailing dash after exactly 2 or 4 digits, so the caret sits where the next field begins.
  if (digits.length === 2 || digits.length === 4) out += "-";
  return out;
}

/** Is this ISO date after today? The form has always refused a document dated in the future. */
export function isFuture(iso: string, today: Date): boolean {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec((iso ?? "").trim());
  if (!m) return false;
  // Compared as strings against a locally-built ISO for `today` — no Date parsing, no timezone shift.
  const t = `${today.getFullYear()}-${pad2(today.getMonth() + 1)}-${pad2(today.getDate())}`;
  return `${m[1]}-${m[2]}-${m[3]}` > t;
}
