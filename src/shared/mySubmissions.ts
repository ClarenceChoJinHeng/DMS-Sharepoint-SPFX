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
  /** Which library the row came from — its URL SEGMENT. Part of the identity, and since
   *  2026-08-21 also shown to the uploader when it is an HC library (see `isHcRow`). */
  library: string;
  name: string;
  /** Server-relative path of the file itself. */
  fileRef: string;
  /**
   * The file's `UniqueId` — the key a deletion or share request is raised against.
   *
   * A GUID rather than the path, because a request may be decided days later and a file renamed or
   * moved in the meantime must still resolve. Undefined when the read fell back to the minimal
   * `$select`, which is why the request buttons check for it rather than assuming it.
   */
  uniqueId?: string;
  /**
   * The upload this file arrived in, and the destination batch within it — stamped by `Form.tsx`
   * (2026-08-22). Both UNDEFINED for every file uploaded before that, and for anything Bulk Upload
   * wrote: `groupSubmissions` presents those as single-file submissions with no reference rather than
   * guessing which belonged together.
   */
  submissionId?: string;
  batchId?: string;
  /**
   * The per-FILE stamp, and **the key the submission record joins on** (2026-08-27).
   * Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
   *
   * ⚠ NOT `uniqueId`, which cannot do this job: Auto-route copies then deletes, so an approved file
   * carries a NEW `UniqueId` and the recorded one resolves to nothing. Undefined for every file
   * uploaded before this existed, and for any library that lacked the column at upload time.
   */
  submissionFileId?: string;
  /**
   * Has this document passed seven years and moved to the archive? (2026-08-22)
   *
   * Derived from the LIBRARY the row came from — see `isArchivedRow` — not from the `Archived`
   * column, so it needs no extra field in the `$select` and cannot disagree with where the file
   * actually is. Undefined is "not established", never "no".
   */
  archived?: boolean;
  status: SubmissionStatus;
  /** When it was uploaded. Auto-route preserves `Created`, so this survives the move. */
  created?: Date;
  /** The approver's rejection comment; "" when there is none or it could not be read. */
  comment: string;
  /**
   * File size in BYTES, as text — `File/Length` is a string, and a raw one reaches the screen as
   * `1483776`. Formatted by `formatBytes` at render. Undefined when it could not be read.
   */
  size?: string;
  /**
   * Last modified. Distinct from `created`, and that difference is the useful part on this page:
   * `created` is when the uploader sent it, `modified` is when the approver acted on it.
   */
  modified?: Date;
  // NOTE: no metadata here, deliberately. The full field set is only readable through
  // FieldValuesAsText, which is a PER-ITEM endpoint — one request per row would mean hundreds on
  // a list this long, and the raw $select returns a lookup id (a bare `15` reached the screen on
  // 2026-08-14). Metadata belongs to the detail view, which reads one item at a time.
}

/**
 * Coerce whatever SharePoint returned into text.
 *
 * NOT a nicety — this is the fix for a live crash on 2026-08-14. `Document Type` is a managed
 * metadata (taxonomy) column: values are WRITTEN as `"Label|GUID"` (gotcha #5) and read back under
 * `odata=nometadata` as an OBJECT, `{ Label, TermGuid, WssId }`. So `(value ?? "").trim()` called
 * `.trim` on an object and took the whole page down with *"(intermediate value).trim is not a
 * function"* — an error naming no field at all. The typed `string` in the row interface was actively
 * misleading here: TypeScript was satisfied and the runtime was not.
 *
 * Handles every shape a SharePoint field can arrive in, not just the one that broke, so the next
 * taxonomy column added to this page cannot repeat it. Anything unrecognised becomes "" — a blank
 * cell, never a crash.
 */
export function textOf(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return `${v}`;
  // Taxonomy single-value is { Label, TermGuid, WssId }; multi-value is an array of them.
  if (Array.isArray(v)) return v.map((x) => textOf(x)).filter((x) => x.length > 0).join("; ");
  const label = (v as { Label?: unknown }).Label;
  if (typeof label === "string") return label.trim();
  return "";
}

/**
 * Read a field that may arrive under a double-encoded key.
 *
 * SharePoint's OData layer double-encodes underscores in property names, so a column whose internal
 * name already contains an encoded character comes back renamed: `Document_x0020_Type` (where
 * `_x0020_` is a space) arrives as `Document_x005f_x0020_x005f_Type`. Reading only the name you
 * wrote yields `undefined` — a permanently blank column, with no error anywhere to explain it.
 *
 * The same shape as `pick` in ApprovalDocument.tsx, which met this first.
 */
export function pickField(row: Record<string, unknown>, ...keys: string[]): string {
  for (const k of keys) {
    const hit = textOf((row ?? {})[k]);
    if (hit.length > 0) return hit;
  }
  return "";
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
/* Generic over the row type so a caller holding a WIDER row keeps it. `MergedRow` (a Submission plus
   `recordState`) goes through here, and returning the narrow `Submission[]` would erase the field the
   caller then has to branch on — forcing a cast, which is exactly how a record row ends up rendered
   as an ordinary live document. Runtime behaviour is unchanged. */
export function filterByTab<T extends Submission>(rows: readonly T[], tab: string): T[] {
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
/* Generic for the same reason as `filterByTab`: the page sorts `MergedRow`s and must keep them. */
export function sortNewestFirst<T extends Submission>(rows: readonly T[]): T[] {
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
  /* ⚠ `27 Aug 2026`, NOT `27/Aug/2026` — changed 2026-08-30 at the client's request, and it is a
     change to the format CLAUDE.md records as agreed (`DD/MMM/YYYY`). Spaces, and NO leading zero:
     that is how a person writes a date, which is the whole point of not showing `8/17/2026`.

     ⚠ SharePoint COLUMN FORMATTING in the library views still produces `DD/MMM/YYYY` — that is
     JSON on the column, not code, and it was not touched. The two now differ, and if the client
     wants them to match, the views are a separate job. */
  return `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

/**
 * The same date, with the time — `27/Aug/2026 21:50`.
 *
 * Client, 2026-08-27: *"also a timer of what time it was uploaded"*. On a day when somebody uploads
 * more than once, the date alone cannot tell two submissions apart or say which came first — every
 * row reads `27/Aug/2026` and the list looks unordered even though it is sorted correctly.
 *
 * `HH:mm`, 24-hour, appended to the agreed `DD/MMM/YYYY` — the same shape the audit log viewer
 * already uses, so two screens do not display one moment two different ways.
 *
 * ⚠ FOR UPLOAD MOMENTS ONLY, never for request dates. A request is decided over days and its date is
 * what matters; an upload is one instant, and the time is what distinguishes it from the next one.
 *
 * Hand-rolled pad, not `padStart`, for the same reason as above: this tsconfig targets below ES2017,
 * so that compiles in an editor and fails in the build.
 */
export function formatSubmittedAt(d: Date | undefined): string {
  if (!d || isNaN(d.getTime())) return "—";
  const h = d.getHours();
  const m = d.getMinutes();
  const hh = h < 10 ? `0${h}` : `${h}`;
  const mm = m < 10 ? `0${m}` : `${m}`;
  return `${formatSubmittedOn(d)} ${hh}:${mm}`;
}

/**
 * Did this row come from a Highly Confidential library?
 *
 * ⚠ WHY THE UPLOADER NEEDS TO KNOW (2026-08-21, found on site). The same document name can exist in
 * BOTH the normal and the HC library — one upload classified Highly Confidential, one not — and My
 * Submissions showed them identically: same name, same tier path, same status badge. An uploader
 * raising a deletion could not tell which of the two they were acting on, and the only way to find
 * out afterwards was the Original location column in the recycle bin. For HC documents that is the
 * wrong place to leave a coin-toss.
 *
 * The row is keyed on `uniqueId`, so the ACTION was always unambiguous — this is about the person.
 *
 * Compared on the URL SEGMENT, which is what `Submission.library` holds, trimmed and lower-cased
 * (SharePoint URLs are case-insensitive). `hc` comes from `cachedHcLibraries()`.
 *
 * Returns FALSE when the HC pair is unresolved rather than guessing from a name prefix: a wrong
 * `true` would brand an ordinary document as confidential, and a site with no HC libraries has no
 * HC rows to mislabel. An uncleared uploader cannot see the HC libraries and therefore has no HC
 * rows either, so the unresolved case cannot hide a tag that was needed.
 */
export function isHcRow(
  librarySegment: string,
  hc: { approval?: string; documents?: string } | undefined,
): boolean {
  const key = (librarySegment ?? "").trim().toLowerCase();
  if (key.length === 0 || !hc) return false;
  const names = [hc.approval, hc.documents]
    .map((n) => (n ?? "").trim().toLowerCase())
    .filter((n) => n.length > 0);
  return names.indexOf(key) !== -1;
}

/**
 * Did this row come from an archive library? (2026-08-22)
 *
 * ⚠ WHY IT MUST BE SHOWN. Without the tag an uploader opens their own history and finds a document
 * sitting in a library they have never heard of, with no explanation — and then presses a deletion
 * or share button that cannot work, because every group holds only Read on the archive. The badge
 * and the refusal are two halves of one answer.
 *
 * ⚠ DERIVED FROM THE LIBRARY, NOT FROM THE `Archived` COLUMN, and that is deliberate. The column is
 * a convenience for search and for views; the LIBRARY is where the file actually is, and it is the
 * thing that decides what the reader can do with it. Reading the column would let a stamp that was
 * never written — or one written and later cleared by hand — contradict the permissions, and the
 * permissions would win silently. It also costs nothing: `library` is already on every row.
 *
 * Same shape and same rules as `isHcRow`, deliberately: compared on the URL SEGMENT trimmed and
 * lower-cased, and FALSE when the archive is unresolved rather than guessed from a name prefix. A
 * wrong `true` would tell an uploader their live document is frozen, and a site with no archive has
 * no archived rows to mislabel.
 */
export function isArchivedRow(
  librarySegment: string,
  archive: { normal?: string; hc?: string } | undefined,
): boolean {
  const key = (librarySegment ?? "").trim().toLowerCase();
  if (key.length === 0 || !archive) return false;
  const names = [archive.normal, archive.hc]
    .map((n) => (n ?? "").trim().toLowerCase())
    .filter((n) => n.length > 0);
  return names.indexOf(key) !== -1;
}

/**
 * Can this viewer carry the action out themselves, making a request pointless?
 *
 * Client, 2026-08-30: *"it doesn't make sense for HOU or HOD to approve their own request when they
 * can delete or share, only PIC will need to raise a request because they can't delete or share
 * without someone's approval."*
 *
 * ⚠ MATCHED AGAINST THE DOCUMENT'S WHOLE TIER CHAIN, never its unit alone. A Head of Department's
 * `DEL`/`SHARE` row sits on the DEPARTMENT term and fans down to every unit beneath it — and that
 * department term is already one of the document's own tiers, so comparing the chain expresses the
 * fan-out exactly, with no term-store expansion and no extra request. Comparing the unit alone
 * would have been wrong for every Head of Department and for any future segment-tier grant.
 *
 * ⚠ FAILS OPEN: an empty `held` list answers false, so the request stays offered. A failed read must
 * never HIDE the button — a PIC who cannot ask is stuck with no route at all, where a Head of Unit
 * shown a needless button raises one request their own screen then decides. Understating somebody's
 * rights costs a redundant request; overstating them costs somebody their only way to act.
 *
 * Comparison is case-insensitive and brace-tolerant, because a term GUID reaches this from two
 * different places — a Group Map row and a document's `<Base>Tid` column — and they do not agree
 * about either.
 */
export function canActDirectly(tierGuids: string[], held: string[]): boolean {
  const norm = (v: string): string => (v ?? "").replace(/[{}]/g, "").trim().toLowerCase();
  const mine: { [key: string]: true } = {};
  for (const h of held ?? []) {
    const k = norm(h);
    if (k.length > 0) mine[k] = true;
  }
  for (const t of tierGuids ?? []) {
    const k = norm(t);
    if (k.length > 0 && mine[k] === true) return true;
  }
  return false;
}
