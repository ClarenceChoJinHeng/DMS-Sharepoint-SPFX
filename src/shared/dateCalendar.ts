// Pure date-grid math behind `DatePicker` — no DOM, no SharePoint, so it is the part of a custom
// calendar popup that can actually be unit tested (the popup itself cannot: this project has no UI
// tests, per its own standing note on `filterSelect.tsx`).
//
// ⚠ EVERYTHING HERE IS LOCAL-TIME, DELIBERATELY. `new Date("2026-09-13")` (a date-only ISO string)
// parses as UTC MIDNIGHT per the ES spec, which then prints a day EARLIER in any timezone west of
// UTC — the exact trap this project's own gotcha #1 warns about for a different endpoint. Every date
// here is built from `new Date(year, monthIndex, day)`, the LOCAL-time constructor, so "today" and a
// chosen day agree with the browser's own clock.

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const MONTH_ABBR = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

/** Sunday-first, matching `Date.prototype.getDay()` (0 = Sunday) — the grid below relies on that. */
export const WEEKDAY_HEADERS = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];

export interface CalendarCell {
  /** `YYYY-MM-DD` — the same shape `parseIso` reads and the picker's `value` carries. */
  iso: string;
  day: number;
  /** False for a leading/trailing cell borrowed from the previous or next month, to fill the grid. */
  inMonth: boolean;
}

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/** `month` is 1–12, matching every other function here — never the 0-based JS `Date` convention. */
export function isoOf(year: number, month: number, day: number): string {
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

/**
 * Strict `YYYY-MM-DD` → `{ year, month, day }` (`month` 1–12), or `undefined`.
 *
 * ⚠ NEVER GUESSES. A malformed or out-of-range value — `2026-13-01`, `2026-02-30` — reads as "no
 * date chosen", not as some nearby real date. Silently normalising `Feb 30` to `Mar 2` would put a
 * date on screen the admin never typed, which is worse than refusing it.
 */
export function parseIso(
  value: string,
): { year: number; month: number; day: number } | undefined {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
  if (!m) return undefined;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  const d = new Date(year, month - 1, day);
  // `new Date` rolls an invalid day/month forward (Feb 30 → Mar 2) instead of erroring — reading the
  // parts back is what catches that roll-forward and refuses it.
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) {
    return undefined;
  }
  return { year, month, day };
}

/** `YYYY-MM-DD` for right now, in the browser's own local time. */
export function isoToday(now: Date = new Date()): string {
  return isoOf(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

export function monthLabel(year: number, month: number): string {
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

/** `DD/MMM/YYYY` — this project's agreed display format (CLAUDE.md gotcha #1). Falls back to the raw
 * value for anything `parseIso` refuses, same "never invent a date" rule. */
export function displayIso(value: string): string {
  const p = parseIso(value);
  if (!p) return value;
  return `${pad2(p.day)}/${MONTH_ABBR[p.month - 1]}/${p.year}`;
}

/** Steps a (year, month) pair by `delta` months, wrapping the year in either direction. */
export function shiftMonth(
  year: number,
  month: number,
  delta: number,
): { year: number; month: number } {
  const total = year * 12 + (month - 1) + delta;
  return { year: Math.floor(total / 12), month: (((total % 12) + 12) % 12) + 1 };
}

function daysInMonth(year: number, month: number): number {
  // Day 0 of the FOLLOWING month is the last day of THIS one — a one-line way to get 28/29/30/31
  // right, leap years included, with no lookup table to keep in sync with `MONTH_NAMES`.
  return new Date(year, month, 0).getDate();
}

/**
 * Always exactly 42 cells (six full weeks), so the popup's height never jumps between a short
 * February and a long October — the panel is a fixed shape whatever month is showing. Weeks start
 * Sunday, matching `WEEKDAY_HEADERS`.
 */
export function calendarGrid(year: number, month: number): CalendarCell[] {
  const first = new Date(year, month - 1, 1);
  const startWeekday = first.getDay(); // 0 = Sunday
  const cells: CalendarCell[] = [];

  const prev = shiftMonth(year, month, -1);
  const prevDays = daysInMonth(prev.year, prev.month);
  for (let i = startWeekday - 1; i >= 0; i--) {
    const day = prevDays - i;
    cells.push({ iso: isoOf(prev.year, prev.month, day), day, inMonth: false });
  }

  const thisDays = daysInMonth(year, month);
  for (let day = 1; day <= thisDays; day++) {
    cells.push({ iso: isoOf(year, month, day), day, inMonth: true });
  }

  const next = shiftMonth(year, month, 1);
  let nextDay = 1;
  while (cells.length < 42) {
    cells.push({ iso: isoOf(next.year, next.month, nextDay), day: nextDay, inMonth: false });
    nextDay += 1;
  }

  return cells;
}
