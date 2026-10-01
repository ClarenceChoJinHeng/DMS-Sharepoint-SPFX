/**
 * CSV export of the DMS Group Map + each group's SharePoint members, for handing
 * to the client as an access list. Pure string building — no SharePoint calls — so
 * the escaping rules stay easy to reason about and test.
 *
 * CSV (not .xlsx) on purpose: Excel opens it directly and it needs no bundled
 * writer library in the SPFx package.
 */

export interface GroupExportRow {
  group: string;
  segment: string;
  /**
   * The tier chain, split one column per tier — NOT one "Tier" cell (client, 2026-08-18).
   *
   * A row stores only the leaf term, and `Tax`, `Legal` and `PM` each exist under several
   * departments, so a single label is genuinely ambiguous about which folder is being granted. A
   * department-scope row carries `tier1` alone and leaves `tier2` blank, which is also what tells it
   * apart from a unit row — previously they were indistinguishable in both the table and this file.
   */
  tier1: string;
  tier2: string;
  role: string;
  memberName: string;
  memberEmail: string;
}

export const GROUP_EXPORT_HEADERS = [
  "Group",
  "Segment",
  "Tier 1",
  "Tier 2",
  "Role",
  "Member Name",
  "Member Email",
];

/** Placeholder in the member columns for a mapped group nobody is in yet. */
export const NO_MEMBERS = "(no members)";

/**
 * Escape one cell. Two separate jobs:
 *  1. CSV quoting — anything with a comma, quote, or newline gets wrapped and its
 *     quotes doubled. Unit names contain commas ("Group Legal, Risk ＆ Compliance").
 *  2. Formula injection — Excel executes a cell starting with = + - @ (or a tab /
 *     carriage return before one). Prefix a single quote so it stays literal.
 *     These names are client-authored text, so this is not hypothetical.
 */
export function csvCell(value: string): string {
  const v = value ?? "";
  const risky = /^[=+\-@\t\r]/.test(v) ? `'${v}` : v;
  return /[",\r\n]/.test(risky) ? `"${risky.replace(/"/g, '""')}"` : risky;
}

export function toCsv(rows: GroupExportRow[]): string {
  const lines = [GROUP_EXPORT_HEADERS.map(csvCell).join(",")];
  for (const r of rows) {
    lines.push(
      [r.group, r.segment, r.tier1, r.tier2, r.role, r.memberName, r.memberEmail]
        .map(csvCell)
        .join(","),
    );
  }
  // CRLF: what Excel expects, and safe everywhere else.
  return lines.join("\r\n");
}

/** `CRS-group-members-2026-07-31.csv` — the client's prefix; a downloaded file is user-facing. */
export function exportFileName(now: Date): string {
  // No String.padStart — the SPFx tsconfig doesn't target a lib that has it
  // (same constraint as Promise.allSettled).
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `CRS-group-members-${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())}.csv`;
}

/**
 * Trigger a browser download. The BOM is required: without it Excel reads the file
 * as ANSI and mangles non-ASCII characters — including the FULLWIDTH ＆ that the
 * "Group Legal, Risk ＆ Compliance"-style unit names carry.
 */
export function downloadCsv(csv: string, fileName: string): void {
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Revoke on the next tick — revoking synchronously can cancel the download in Edge.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}
