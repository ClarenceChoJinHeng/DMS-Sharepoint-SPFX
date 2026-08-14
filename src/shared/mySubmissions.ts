/**
 * My Submissions — turning two libraries' rows into one uploader-readable list.
 *
 * Spec: docs/superpowers/specs/2026-08-14-my-submissions-design.md
 *
 * Pure and SPFx-free: the folder trail and the ordering are what an uploader actually reads, and
 * both have a wrong-but-plausible version. A mis-built trail sends someone to the wrong folder to
 * look for their file; the wrong sort buries the row they opened the page for.
 *
 * The moderation mapping is NOT redefined here — it is `statusToDecision` from approvalQueue.ts.
 * Two mappings of one SharePoint field is how the approver's screen and the uploader's screen end
 * up disagreeing about whether a document was approved.
 */

import { Decision, statusToDecision } from "./approvalQueue";

/** Pending / Approved / Rejected, from the approver's own vocabulary. */
export type SubmissionStatus = Decision;

export { statusToDecision };

export interface Submission {
  /** SharePoint item id, unique only WITHIN its library — see `submissionKey`. */
  itemId: number;
  /** Which library the row came from. Part of the identity; not shown to the uploader. */
  library: string;
  name: string;
  /** Server-relative path of the file itself. */
  fileRef: string;
  status: SubmissionStatus;
  /** When it was uploaded. Auto-route preserves `Created`, so this survives the move. */
  created?: Date;
  documentType: string;
  /** The approver's rejection comment; "" when there is none or it could not be read. */
  comment: string;
}

/**
 * A stable row key.
 *
 * Item ids repeat ACROSS libraries — an approved file in `Documents` can share an id with a pending
 * one in the approval library — so a list keyed on the id alone silently drops a row. Combined with
 * the library, which is what makes them distinct.
 */
export function submissionKey(s: Submission): string {
  return `${s.library}#${s.itemId}`;
}

/**
 * The folder trail a file sits in, without the file itself.
 *
 * `/sites/X/ApprovalDocument/NBPOLHO/CDS/UPSUPPORT/2024/Tax Return/q1.pdf`
 *   → `["NBPOLHO", "CDS", "UPSUPPORT", "2024", "Tax Return"]`
 *
 * The library segment goes with everything above it. An uploader does not think about which of the
 * two libraries their file currently lives in — that is what the status badge is for — and showing
 * `ApprovalDocument` on one row and `Documents` on the next invites precisely the question this
 * page exists to stop them asking.
 *
 * Matched on the library's URL SEGMENT, never its title: the two differ ("Approval Document" vs
 * "/ApprovalDocument"), and a title match would never fire — leaving the library name in the trail,
 * wrongly and silently (gotcha #12).
 */
export function folderTrail(fileRef: string, librarySegments: readonly string[]): string[] {
  const parts = (fileRef ?? "").split("/").filter((p) => p.trim().length > 0);
  if (parts.length === 0) return [];
  // Drop the file itself — the trail is where it IS, not what it is.
  parts.pop();
  const wanted = (librarySegments ?? [])
    .map((l) => l.trim().toLowerCase())
    .filter((l) => l.length > 0);
  for (let i = 0; i < parts.length; i++) {
    if (wanted.indexOf(parts[i].toLowerCase()) !== -1) return parts.slice(i + 1);
  }
  // No library segment recognised: return everything below /sites/<site>/ rather than nothing. A
  // half-recognised path is still a useful hint; an empty cell tells the uploader nothing at all.
  return parts[0]?.toLowerCase() === "sites" ? parts.slice(2) : parts;
}

/** The trail as one string, in the separator the client's screens already use. */
export function trailText(trail: readonly string[]): string {
  return (trail ?? []).join(" › ");
}

/**
 * Count by status, for the tab labels.
 *
 * Counts what is SHOWN. A tally taken from a different set than the rows beneath it is worse than
 * no tally: the reader cannot reconcile it, and assumes rows are missing.
 */
export function countByStatus(rows: readonly Submission[]): Record<string, number> {
  const out: Record<string, number> = { All: 0, Pending: 0, Approved: 0, Rejected: 0 };
  for (const r of rows ?? []) {
    out.All++;
    out[r.status] = (out[r.status] ?? 0) + 1;
  }
  return out;
}

/** Filter to a tab. `All` really is all — no row is invisible in every tab. */
export function filterByTab(rows: readonly Submission[], tab: string): Submission[] {
  if (tab === "All") return (rows ?? []).slice();
  return (rows ?? []).filter((r) => r.status === tab);
}

/**
 * Newest first.
 *
 * The question this page answers is "what happened to the thing I just sent", so the answer is
 * almost always the most recent row. A row with no date sorts LAST rather than first: an unknown
 * date is not evidence of recency, and putting it on top would displace the row being looked for.
 */
export function sortNewestFirst(rows: readonly Submission[]): Submission[] {
  return (rows ?? []).slice().sort((a, b) => {
    const at = a.created ? a.created.getTime() : Number.NEGATIVE_INFINITY;
    const bt = b.created ? b.created.getTime() : Number.NEGATIVE_INFINITY;
    return bt - at;
  });
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * `DD/MMM/YYYY` — the agreed client display format (gotcha #1).
 *
 * Display only. Nothing here is sent back to SharePoint, so the `M/D/YYYY` and ISO write formats
 * are somebody else's problem — and a date SharePoint has already formatted must never be re-parsed
 * here, which is exactly how the `M/D/YYYY` trap gets reintroduced.
 */
export function formatSubmittedOn(d: Date | undefined): string {
  if (!d || isNaN(d.getTime())) return "—";
  // Hand-rolled pad, not String.padStart: this tsconfig targets below ES2017, the same reason
  // Promise.allSettled is unavailable here (CLAUDE.md #3). It compiles in an editor and fails in
  // the build, so the substitute is worth the two extra characters.
  const n = d.getDate();
  const day = n < 10 ? `0${n}` : `${n}`;
  return `${day}/${MONTHS[d.getMonth()]}/${d.getFullYear()}`;
}
