/**
 * The submission RECORD — what survives after the document does not.
 *
 * Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
 *
 * Client, 2026-08-27: *"if I deleted the file the my submission file list also disappears, can we
 * ensure that doesn't happen? … just show this specific file is deleted but what matters is its
 * recorded for that submission."*
 *
 * My Submissions derives everything from the libraries, so a deleted file does not become "deleted"
 * — it stops existing, and its batch quietly shrinks. This module is the other half: a list row per
 * uploaded file, and the rules for joining those rows back to whatever is still live.
 *
 * Pure and SPFx-free. Every rule here has a wrong-but-plausible version, and two of them would tell
 * an uploader their documents had been destroyed when they had not.
 *
 * ⚠⚠ THE JOIN IS ON THE STAMPED `SubmissionFileId`, NEVER ON `UniqueId`. The agreed design said
 * `UniqueId`, and it would have marked EVERY APPROVED FILE as deleted: Auto-route is
 * copy-stamp-delete, so the routed copy in `Documents` carries a new `UniqueId` and the source
 * holding the recorded one is deleted. A stamped column survives the copy — which is exactly why
 * `SubmissionId`/`BatchId` had to exist in all four libraries (2026-08-22) — so the record is keyed
 * on our data rather than on SharePoint's identity. See §2 of the spec.
 *
 * ⚠ 2026-09-20 — `UniqueId` IS NOW USED TOO, BUT ONLY AS A SECONDARY, NARROWER JOIN, AND THIS DOES
 * NOT REOPEN THE PARAGRAPH ABOVE. Since 2026-09-18, `SubmissionFileId` is written onto the live
 * document by `CRS — Apply pending tags` on a poll, not at upload time — so for a few minutes after
 * every upload the primary join genuinely finds nothing, even though nothing is wrong. `UniqueId`
 * fills exactly that gap: it is present on the document from the moment of creation, before it has
 * moved anywhere, so it can join a still-settling record to its live row. It stops being tried the
 * instant the primary join succeeds, and it naturally stops MATCHING the moment a file is routed
 * (its `UniqueId` changes), so it can never resurrect the false-"still there" case the paragraph
 * above rejected. See `indexLiveByUniqueId`.
 */

import { Submission } from "./mySubmissions";

/**
 * Internal names on the submissions list. Space-free, as with the audit log and requests.
 *
 * ⚠ `SubmissionFileId` IS A LITERAL HERE AND A CONSTANT IN `optionalColumns.ts`, and that is not an
 * oversight. This module is unit-tested, and `optionalColumns.ts` imports `@microsoft/sp-http`, which
 * the test environment cannot resolve — the same constraint that keeps `PRIMED_SUFFIXES` in
 * `naming.ts` rather than in `spNaming.ts`. The two are pinned equal by test instead.
 */
export const RECORD_COLUMNS: Array<{ name: string; type: number }> = [
  { name: "SubmissionRef", type: 2 },
  { name: "BatchRef", type: 2 },
  // The join key. Every read of this list filters on it.
  { name: "SubmissionFileId", type: 2 },
  // Secondary only. Useful while the file is still where it was uploaded; never decides Deleted.
  { name: "ItemUniqueId", type: 2 },
  { name: "FileName", type: 2 },
  { name: "ItemPath", type: 3 },
  { name: "LibraryTitle", type: 2 },
  { name: "UploadedBy", type: 2 },
  { name: "UploadedAt", type: 4 },
  { name: "MetadataSnapshot", type: 3 },
  // `Form` or `BulkUpload`. TEXT, never Choice: a value absent from a Choice column's `Choices`
  // fails the whole write, so the day a third writer appears every row from it is lost silently.
  { name: "Source", type: 2 },
  /* ── Replacement (2026-08-28) ────────────────────────────────────────────────
     Set when a LATER upload overwrote this file. Their PRESENCE is what makes the record
     `cancelled` rather than `deleted` — see `RecordState`.

     ⚠ TWO COLUMNS, NOT A YES/NO FLAG. "Who replaced it and when" is the whole reason the client
     asked for this state instead of leaving it as Deleted: a deleted document is a loss, a replaced
     one is a newer version somebody deliberately filed. A boolean records the first half and throws
     away the half that answers the question the uploader actually has. */
  { name: "ReplacedAt", type: 4 },
  { name: "ReplacedBy", type: 2 },
  /* ── Archiving (2026-09-03) ──────────────────────────────────────────────────
     Set by the seven-year archive movers when they move this file out of the approved-side library.
     Its PRESENCE is what makes the record `archived` rather than `deleted`.

     ⚠ THIS EXISTS BECAUSE THIS PAGE CANNOT SEE THE ARCHIVE AND MUST NOT. Archive access narrowed to
     the two C-Level roles on 2026-09-02, and My Submissions is a PIC/HoU page — so an archived file
     resolves in NO library the viewer can read, and `mergeRecords` reported it as `deleted`. Client,
     2026-09-03: *"why all the archive files consider as deleted, can we ensure if its archive it
     shows archive instead?"*

     ⚠ THERE IS NO CLIENT-SIDE ALTERNATIVE, and that is worth stating so it is not re-litigated: the
     viewer holds nothing on Archive/ArchiveHC, so a probe by path, title or `GetFileById` is
     security-trimmed to 404 — indistinguishable from deletion. Only something that WATCHED the move
     can record it, exactly as `ReplacedAt` is written by the flow that watched the overwrite.

     ⚠ NO `ArchivedBy`. The movers are SCHEDULED flows: there is no person, and a column naming the
     service account would read as somebody having done it deliberately to that file. */
  { name: "ArchivedAt", type: 4 },
  /* ── Direct staging delete (2026-09-11) ──────────────────────────────────────
     Client: "now pic can delete without approval on staging... the status in my submission will
     show cancelled." Set when the UPLOADER themselves recycled a pending/rejected file directly on
     the approval library — no request, no approver. Their PRESENCE is what makes the record
     `withdrawn` rather than `deleted`; see `RecordState`.

     ⚠ NOT THE SAME AS `ReplacedAt`. A replace means somebody uploaded a NEWER document over this
     one; a withdrawal means the uploader deleted it with nothing behind it. Different facts, and
     `RECORD_STATE_LABEL.withdrawn` is deliberately "cancelled" — the ONE record state this codebase
     shows that literal word for, since `cancelled` (the type value) already reads "replaced" on
     screen and reusing it here would have shown the wrong badge for what the client asked for. */
  { name: "WithdrawnAt", type: 4 },
  { name: "WithdrawnBy", type: 2 },
  /* ── Tag-by-proxy (2026-09-18) ─────────────────────────────────────────────
     Client: PIC and Approver lose direct Edit rights on the document libraries, so they can no
     longer retag a file by hand in the library view. `TagPayload` is what `CRS — Apply pending
     tags` (Power Automate, running as crs@sdguthrie.com) reads to perform the write the uploader's
     own browser can no longer make directly.

     ⚠ NOT THE SAME FIELD AS `MetadataSnapshot`, DELIBERATELY. `MetadataSnapshot` stores LABELS
     ONLY, is missing Document Type/Year/Keyword/Business Segment, and exists purely for DISPLAY —
     it is read by My Submissions' detail panel for a deleted/archived/replaced record, and its
     correctness bar is "good enough for a person to read." `TagPayload` is the EXACT write
     payload — the same `{FieldName, FieldValue}` array `Form.tsx`/`BulkUpload.tsx` used to POST
     to `validateUpdateListItem` directly, taxonomy fields already in `Label|GUID` form — and its
     correctness bar is "a flow can replay this and get the same result the uploader chose."
     Reusing `MetadataSnapshot` for this would entangle a display concern with a write concern that
     have different tolerances for being slightly wrong. See the 2026-09-18 design doc. */
  { name: "TagPayload", type: 3 },
  /** blank/absent ⇒ not yet processed. "Tagged" ⇒ applied. "Failed" ⇒ see TagError. Same
   * three-state shape as every other status field in this codebase: blank is not a failure. */
  { name: "TagStatus", type: 2 },
  { name: "TagError", type: 3 },
];

/** The one column name this module and `optionalColumns.ts` must agree on. Pinned by test. */
export const RECORD_JOIN_COLUMN = "SubmissionFileId";

/** Which screen wrote the row. */
export type RecordSource = "Form" | "BulkUpload";

/** One recorded upload. Nothing here describes the document's STATE — see `mergeRecords`. */
export interface SubmissionRecord {
  /** The record row's own list item id. */
  itemId: number;
  submissionRef: string;
  batchRef: string;
  /** `SubmissionFileId` — the join key. A record with none can never resolve (see `isJoinable`). */
  fileId: string;
  uniqueId?: string;
  fileName: string;
  /** Server-relative path at upload time. Where it WAS, not necessarily where it is. */
  itemPath: string;
  libraryTitle: string;
  uploadedBy: string;
  uploadedAt?: Date;
  /** The tier and metadata values as filled in, so a deleted row still shows what was submitted. */
  metadata: Record<string, string>;
  source: string;
  /** When a later upload overwrote this file. Present ⇒ the record is `cancelled`, not `deleted`. */
  replacedAt?: Date;
  /** Who overwrote it. Display only — never a matching key. */
  replacedBy?: string;
  /** When the archive mover moved it. Present ⇒ the record is `archived`, not `deleted`. */
  archivedAt?: Date;
  /** When the uploader deleted it THEMSELVES, direct on staging. Present ⇒ `withdrawn`, not `deleted`. */
  withdrawnAt?: Date;
  /** Who deleted it. Always the uploader — this action exists only for a PIC's own pending/rejected
   * file — but stored rather than assumed, for the same reason `replacedBy` is. */
  withdrawnBy?: string;
  /** The exact `{FieldName, FieldValue}[]` payload for `CRS — Apply pending tags` to replay,
   * JSON-encoded. Absent on a row this flow has not been asked to act on (a display-only write, or
   * a site predating this feature). */
  tagPayload?: string;
  /** blank/undefined ⇒ not yet processed by the flow. "Tagged" ⇒ applied. "Failed" ⇒ see tagError. */
  tagStatus?: string;
  tagError?: string;
}

/**
 * What became of a recorded file.
 *
 * ⚠ `unknown` IS THE ONE THAT MATTERS. A record that did not resolve means "deleted" ONLY if the
 * libraries were actually read. If a read failed, an unresolved record proves nothing — and reporting
 * Deleted from a transient error tells an uploader their document was destroyed, which is the worst
 * false positive this page could produce. Empty ≠ unknown, in the place where the cost is somebody
 * believing their work is gone.
 *
 * `cancelled` (2026-08-28, client: *"The older submission under My Submission will change to
 * Cancelled status if replaced"*, then *"I know its weird but client want it to be Cancelled"*) is a
 * REPLACED file — somebody uploaded a newer document over it deliberately.
 *
 * ⚠ IT IS NOT A KIND OF `deleted`, AND THE DIFFERENCE IS THE POINT. `deleted` is derived — a record
 * that failed to resolve, judged against whether the read was complete. `cancelled` is OBSERVED: we
 * watched the replacement happen and wrote it on the record, so it needs no library read to be true
 * and cannot be produced by a failed one. That is why it is a separate state and not a label applied
 * to `deleted` after the fact.
 *
 * `archived` (2026-09-03) is a file the seven-year mover moved into the archive. Like `cancelled` it
 * is OBSERVED — written on the record by the flow that performed the move — and for the same reason:
 * this page deliberately does not read the archive (access is C-Level only since 2026-09-02), so an
 * archived file resolves nowhere and would otherwise be reported as a loss. It is NOT a kind of
 * `deleted`: the document exists, is retained, and is simply somewhere this viewer cannot go.
 *
 * `withdrawn` (2026-09-11, client: PICs can now delete a pending/rejected file directly on staging,
 * with no approval) is the uploader's OWN deliberate deletion of their own unreviewed draft — the
 * approval-library equivalent of a request-based `Documents` delete, but with nobody to decide it.
 * OBSERVED, exactly like `cancelled`/`archived`: the in-app action that recycles the file also stamps
 * the record, so it needs no library read and cannot be produced by a failed one. It is NOT a kind of
 * `cancelled` (that means REPLACED by a newer upload, a different fact) even though its on-screen
 * label happens to be the literal word "Cancelled" — see `RECORD_STATE_LABEL`.
 */
/**
 * `settling` (2026-09-19, client: *"why is it showing not checked? can we not show pending
 * instead? would that not make sense"*) — a record whose libraries were all read SUCCESSFULLY but
 * which is still within its `RECENT_UPLOAD_GRACE_MS` window (see `recentlyUploaded`). It is NOT the
 * same fact as `unknown`, and keeping them apart is the point: `unknown` means a READ FAILED, so
 * the true status could genuinely be anything, including something bad. `settling` means every read
 * SUCCEEDED and simply found nothing yet, on a record too young for that to be surprising — nothing
 * could have approved, rejected or deleted it in the last few minutes when the approver's own queue
 * cannot even resolve it yet. So "Pending" is not a guess here, it is the only status the document
 * could honestly hold — and telling an uploader "not checked" over a plain library lag reads as a
 * fault when there is none. Like `deleted`, it is DERIVED rather than observed, and resolves itself
 * on the very next read: either the join succeeds (becomes an ordinary live row) or the grace
 * period lapses (becomes `deleted`, if it is still unresolved by then). */
export type RecordState =
  | "live"
  | "deleted"
  | "cancelled"
  | "archived"
  | "withdrawn"
  | "settling"
  | "unknown";

/**
 * A row as the page renders it: either a live document, or a record of one that is gone.
 *
 * Extends `Submission` so `groupSubmissions`, `sortNewestFirst` and the existing rendering keep
 * working unchanged — a deleted file must still appear inside its own batch, which is the entire
 * point of recording it.
 */
export interface MergedRow extends Submission {
  /**
   * Absent on an ordinary live row. `deleted` greys the row and removes preview and Open file;
   * `unknown` says the check could not be made, and must not read as either of the others.
   */
  recordState?: RecordState;
  /** The record behind this row, when there is one. Carries the metadata snapshot. */
  record?: SubmissionRecord;
}

/**
 * The `Source` each writer stamps on a record, named ONCE.
 *
 * ⚠ `Form.tsx` ALSO writes `source: "UploadForm"` — to the AUDIT LOG, via `writeAudit`, which is a
 * different list and a different field. Only the two below ever reach `CRS Submissions`. Comparing
 * against the wrong one would classify every ordinary upload as something else, and the symptom
 * would be an empty tab rather than an error.
 */
export const RECORD_SOURCE = { form: "Form", bulk: "BulkUpload" } as const;

/**
 * Did this row come from Bulk Upload?
 *
 * ⚠ MATCHED POSITIVELY ON `BulkUpload`, never by excluding `Form`. A row with no record at all —
 * anything uploaded before the record feature, or a file whose record write was refused — carries no
 * source, and an exclusion test would sweep every one of those into the Bulk Upload tab, where they
 * do not belong. Matching positively means an unknown source stays exactly where it already was, so
 * this can only ever move rows a writer explicitly marked.
 *
 * Trimmed and case-insensitive: the value survives a round trip through a SharePoint text column,
 * and a stray space there would silently empty the tab.
 */
export function isBulkUploadRow(r: MergedRow): boolean {
  return (r.record?.source ?? "").trim().toLowerCase() === RECORD_SOURCE.bulk.toLowerCase();
}

export interface MergeResult {
  rows: MergedRow[];
  /** Records that resolved to a live document. */
  live: number;
  /** Records whose file is gone — only ever non-zero when the live read was complete. */
  deleted: number;
  /** Records replaced by a later upload. Observed, so it needs no live read to be true. */
  cancelled: number;
  /** Records moved to the archive. Observed, like `cancelled` — never derived from a failed read. */
  archived: number;
  /** Records the uploader deleted themselves, direct on staging. Observed, like `cancelled`. */
  withdrawn: number;
  /** Records still within their upload grace period — every read succeeded, just too soon to
   * expect a join yet. Shown as "Pending", never "not checked" — see `RecordState`'s own comment. */
  settling: number;
  /** Records that could not be checked, because a library read failed. */
  unknown: number;
}

/** Trim and lower-case. The stamp is ours, so there are no braces to strip — unlike a term GUID. */
export function normaliseFileId(v: string | undefined): string {
  return (v ?? "").trim().toLowerCase();
}

/**
 * Can this record ever be joined to a live document?
 *
 * A row written without a stamp is unresolvable for ever, so it would sit on the page as a permanent
 * false "Deleted". The writers refuse to create such a row; this is the belt to that braces — an
 * unjoinable record already in the list is dropped rather than displayed as a loss.
 */
export function isJoinable(r: SubmissionRecord): boolean {
  return normaliseFileId(r.fileId).length > 0;
}

/** Live rows indexed by their stamp. Rows with no stamp are absent — they predate the column. */
export function indexLiveByFileId(live: readonly Submission[]): Record<string, Submission> {
  const out: Record<string, Submission> = {};
  for (const s of live ?? []) {
    const k = normaliseFileId(s.submissionFileId);
    if (k.length === 0) continue;
    // FIRST WINS. A stamp should be unique, but Replace (`nameConflictBehavior: 1`, 2026-08-27) can
    // put one file's columns onto another's, and the approval-library copy briefly coexists with the
    // routed one mid-run. Either way, one live document is enough to answer "not deleted".
    if (out[k] === undefined) out[k] = s;
  }
  return out;
}

/**
 * Live rows indexed by their own `UniqueId` — a SECONDARY join, used only while the primary one (the
 * `SubmissionFileId` stamp) has not landed on the document yet.
 *
 * ⚠⚠ THIS DOES NOT CONTRADICT THE WARNING AT THE TOP OF THIS FILE — it answers a narrower question.
 * That warning is about using `UniqueId` to decide whether a document has been DELETED, which breaks
 * the moment a file is approved and routed (Auto-route copies it to a NEW `UniqueId` and deletes the
 * source holding the recorded one). This index is never used for that.
 *
 * It exists only to close the window BETWEEN "the record was written" and "`CRS — Apply pending
 * tags` stamped the live document" (2026-09-18 onward: that flow, not the upload itself, is what now
 * writes `SubmissionFileId`). During that window the document has not moved — it is still in the
 * same library it was uploaded to, under the SAME `UniqueId` SharePoint assigned at creation, which
 * (unlike the custom stamp) is present on every read from the first moment. Once a file is approved
 * and routed its `UniqueId` changes and this index can no longer find it — the safe, narrow failure
 * mode: the fix helps during settling and steps aside once routing makes it inapplicable, by which
 * point the primary join (which DOES survive routing) has normally already taken over.
 */
export function indexLiveByUniqueId(live: readonly Submission[]): Record<string, Submission> {
  const out: Record<string, Submission> = {};
  for (const s of live ?? []) {
    const k = normaliseFileId(s.uniqueId);
    if (k.length === 0) continue;
    if (out[k] === undefined) out[k] = s;
  }
  return out;
}

/**
 * Turn a record into a renderable row, for a file that is no longer live.
 *
 * ⚠ THE STATUS IS NOT GUESSED. `status` is required by `Submission`, and every value it can take
 * ("Pending", "Approved", "Rejected") would be a claim about a document nobody can read any more. It
 * is set to `Pending` only because the type demands a value; `recordState` is what the screen keys
 * on, so the caller must branch on `recordState` FIRST and never fall through to the badge. The
 * status counts exclude these rows for the same reason — see `liveRowsOnly`.
 */
export function rowFromRecord(r: SubmissionRecord, state: RecordState): MergedRow {
  return {
    // Not a document id — the record's own. Never used to address a file; `mergedKey` keys the row,
    // and it uses the stamp.
    itemId: r.itemId,
    library: r.libraryTitle,
    name: r.fileName,
    fileRef: r.itemPath,
    uniqueId: r.uniqueId,
    submissionId: r.submissionRef,
    batchId: r.batchRef,
    submissionFileId: r.fileId,
    status: "Pending",
    created: r.uploadedAt,
    comment: "",
    recordState: state,
    record: r,
  };
}

/**
 * A stable React key across both kinds of row.
 *
 * `submissionKey` is `library#itemId`, and for a record row that id belongs to a DIFFERENT list — so
 * two records could collide with each other and with a live document. The stamp is unique per file
 * and is what these rows are actually about.
 */
export function mergedKey(row: MergedRow): string {
  const stamp = normaliseFileId(row.submissionFileId);
  /* ⚠ ANY record-backed state, NOT a list of them. This read `deleted || unknown` and MISSED
     `cancelled` when that state was added on 2026-08-28 — so a replaced row fell through to
     `library#itemId`, where the id belongs to the RECORD list rather than the document, and could
     collide with a live document's key. React drops a duplicate-keyed row silently.

     Tested by asking what the row IS (does it come from a record?) rather than enumerating which
     kinds there are, so the next state added cannot reintroduce this. */
  if (row.recordState !== undefined && row.recordState !== "live") {
    return `rec#${stamp.length > 0 ? stamp : row.itemId}`;
  }
  return `${row.library}#${row.itemId}`;
}

/**
 * How long a record is allowed to sit unresolved before it may honestly be called "deleted".
 *
 * Set well above `CRS — Apply pending tags`'s own poll interval and the ~90s
 * `pollForTagStatus` budget self-approve waits on for the same stamp — the two exist for
 * different reasons (one gates a permission-sensitive auto-approve, this one gates a display
 * label) but are racing the identical flow, so this window is deliberately the more generous
 * of the two. Understating "deleted" for a few extra minutes costs nothing; getting it wrong in
 * the other direction is the failure this whole module exists to prevent.
 */
export const RECENT_UPLOAD_GRACE_MS = 5 * 60 * 1000;

/**
 * True while a record is still within its settling window — i.e. it may not yet have been
 * stamped onto its live document by `CRS — Apply pending tags`, so an unresolved join says
 * nothing about whether the document still exists.
 *
 * `now.getTime() - uploadedAt.getTime()` deliberately catches BOTH directions: a genuinely
 * recent upload (small positive difference) and one whose `uploadedAt` reads slightly in the
 * FUTURE relative to `now` from ordinary clock skew between the browser that wrote it and
 * whatever machine is merging later (a negative difference, which is still `< GRACE_MS`) — both
 * mean "just uploaded", and neither should assert a deletion.
 */
function recentlyUploaded(r: SubmissionRecord, now: Date): boolean {
  if (r.uploadedAt === undefined) return false;
  return now.getTime() - r.uploadedAt.getTime() < RECENT_UPLOAD_GRACE_MS;
}

/**
 * Merge recorded uploads with whatever is still live.
 *
 * The union, deliberately:
 * - every LIVE row is kept, recorded or not — files uploaded before this existed must keep behaving
 *   exactly as they do now, and dropping them would delete a year of history from the page;
 * - a record that resolves adds NOTHING, or every file would appear twice;
 * - a record that does not resolve becomes a row of its own, which is the feature.
 *
 * ⚠ `liveReadComplete` IS LOAD-BEARING AND FAILS SAFE. Unresolved records come back `unknown` rather
 * than `deleted` when it is false. A page that tells an uploader their documents were destroyed
 * because a library was throttled for a second is far worse than one that says it could not check.
 *
 * ⚠ IT TAKES A PREDICATE AS WELL AS A BOOLEAN, AND ON A REAL SITE THE PREDICATE IS THE CORRECT ONE.
 * My Submissions reads up to six libraries and SWALLOWS failures on the HC pair and the archive —
 * deliberately, because an unreadable HC library is the NORMAL case for everyone without clearance
 * and must not blank the page. So a single boolean would be false for most users on most sites, and
 * nothing would ever be reported as deleted: the feature would silently not work for the majority.
 *
 * The honest question is per record: *was the chain of libraries THIS file could be living in fully
 * read?* A normal-library record needs the normal chain (approval → documents → archive); an HC one
 * needs the HC chain. The caller knows which reads succeeded, so it supplies the rule.
 */
export function mergeRecords(
  records: readonly SubmissionRecord[],
  live: readonly Submission[],
  liveReadComplete: boolean | ((r: SubmissionRecord) => boolean),
  now: Date = new Date(),
): MergeResult {
  const complete = typeof liveReadComplete === "function"
    ? liveReadComplete
    : (): boolean => liveReadComplete === true;
  const index = indexLiveByFileId(live);
  // Secondary join — see `indexLiveByUniqueId`'s own comment for why this is safe alongside the
  // warning at the top of this file about never using `UniqueId` to decide DELETED.
  const uniqueIndex = indexLiveByUniqueId(live);

  /* ── The record wins on WHEN, and carries itself onto the live row ──────────────
     ⚠⚠ `Created` IS A PROPERTY OF THE FILE; `uploadedAt` IS A PROPERTY OF THE SUBMISSION — and this
     page groups by submission. They agree within seconds for an ordinary upload and disagree badly
     in three known cases, all found on 2026-08-28:

       1. A REPLACEMENT. `Files/Add(overwrite=true)` keeps the existing list item and adds a version,
          so the replacing file INHERITS the original's `Created`. Verified live: two submissions ten
          minutes apart both displayed 16:21, and the displaced record (stamped a few seconds AFTER
          the original upload finished) therefore sorted ABOVE its own replacement.
       2. A ROUTED file. Both routing flows stamp `Created` ~8 hours early — a UTC value written in a
          locale format SharePoint reads as local — so a document uploaded before 08:00 lands on the
          previous day. Long known, never fixed, and this is what fixes it.
       3. `isBulkUploadRow` reads `row.record`, which only GONE rows carried — so a live bulk-uploaded
          file could never appear in the Bulk Upload tab. It read `(0)` on a site full of them.

     `uploadedAt` is written client-side at the moment of submission, so it is subject to none of
     those. Falls back to `created` when absent, which is every row uploaded before records existed.

     ⚠ `recordState` IS DELIBERATELY NOT SET HERE. Several call sites test it for TRUTHINESS
     (`row.recordState ? goneRow : fileCard`), so `"live"` would grey out every recorded row and
     strip its Open file button. Absent means live, and that is the contract. */
  const recordByFileId: { [key: string]: SubmissionRecord } = {};
  // Same idea, keyed by `ItemUniqueId` instead — the fallback for a record whose live document has
  // not been stamped with `SubmissionFileId` yet. See `indexLiveByUniqueId`'s comment for the scope.
  const recordByUniqueId: { [key: string]: SubmissionRecord } = {};
  for (const r of records ?? []) {
    if (!isJoinable(r)) continue;
    const key = normaliseFileId(r.fileId);
    if (recordByFileId[key] === undefined) recordByFileId[key] = r;
    const uk = normaliseFileId(r.uniqueId);
    if (uk.length > 0 && recordByUniqueId[uk] === undefined) recordByUniqueId[uk] = r;
  }
  const rows: MergedRow[] = (live ?? []).map((s) => {
    /* ⚠ THE `SubmissionFileId` MATCH IS TRIED FIRST AND ALWAYS WINS. Only when it misses — which,
       since 2026-09-18, is expected for a few minutes on every fresh Bulk Upload/Form upload while
       `CRS — Apply pending tags` has not yet run — do we fall back to `UniqueId`. Found live
       2026-09-20: without this fallback, a batch mid-tagging showed up TWICE in My Submissions —
       once as its own untagged live row (which, having no `SubmissionId` of its own yet, fell into
       the "grouped by folder and date" bucket meant for pre-tracking files) and once as a synthetic
       `settling` placeholder built from the record. Matching via `UniqueId` here lets the real live
       row absorb the record directly, so only one entry — correctly grouped — ever exists. */
    const rec = recordByFileId[normaliseFileId(s.submissionFileId)]
      ?? recordByUniqueId[normaliseFileId(s.uniqueId)];
    if (rec === undefined) return { ...s };
    return {
      ...s,
      record: rec,
      created: rec.uploadedAt ?? s.created,
      // Fills in the reference `groupSubmissions` keys on when the document's own column has not
      // been stamped yet — harmless (a no-op via `??`) once it has, since both then agree.
      submissionId: s.submissionId ?? rec.submissionRef,
      batchId: s.batchId ?? rec.batchRef,
    };
  });
  let liveCount = 0;
  let deleted = 0;
  let cancelled = 0;
  let archived = 0;
  let withdrawn = 0;
  let settling = 0;
  let unknown = 0;

  for (const r of records ?? []) {
    // Unjoinable rows are dropped, never shown: they can never resolve, so displaying one asserts a
    // loss that was never established.
    if (!isJoinable(r)) continue;
    // Resolved either way — by the stamp, or (while it is still missing) by `UniqueId`. Checking
    // both here is what stops the loop below from ALSO pushing a `rowFromRecord` placeholder for a
    // record the `rows` map above has already attached to its real live row.
    const resolvedById = index[normaliseFileId(r.fileId)] !== undefined;
    const resolvedByUniqueId =
      normaliseFileId(r.uniqueId).length > 0 &&
      uniqueIndex[normaliseFileId(r.uniqueId)] !== undefined;
    if (resolvedById || resolvedByUniqueId) {
      liveCount += 1;
      continue;
    }
    /* ⚠ ANY throw from the caller's predicate is read as NOT complete. A rule that cannot answer
       must never license the claim that a document was deleted. */
    let judgeable = false;
    try {
      judgeable = complete(r) === true;
    } catch {
      judgeable = false;
    }
    /* ⚠ THE ORDER OF THESE THREE IS LOAD-BEARING, and `live` winning is why the check above runs
       first: a record that STILL RESOLVES describes a document that exists, so calling it cancelled
       would claim a loss that did not happen — which is exactly what happens if the replacement's
       re-stamp failed and the old id is still on the file.

       `cancelled` then beats both `deleted` and `unknown` because it is OBSERVED rather than
       derived: `replacedAt` is a fact we wrote when we watched the overwrite succeed, so it holds
       whether or not the libraries could be read this time. Deriving it from `judgeable` would make
       a transient read failure downgrade a known replacement to "we are not sure", which is less
       true than what we already know.

       `archived` sits between them for the SAME reason it beats `deleted` and `unknown` — it is
       observed, not derived — and BELOW `cancelled` only because the two cannot honestly co-occur:
       a replaced file's old content never reaches the archive under this stamp, and an archived file
       cannot be uploaded over. If both are somehow set, the replacement is the later fact about what
       became of this record, so it wins.

       ⚠ AND `live` STILL BEATS ARCHIVED, which is why the resolve check runs first: a stamp left
       behind by a mover that moved something else, or a file restored from the archive, must not be
       reported as gone while the viewer can plainly see it.

       `withdrawn` CHECKED FIRST, above all three. It is the uploader's own deliberate, terminal act
       on their own pending/rejected draft — nothing produced by ANY of the other three columns can
       apply to a file that no longer exists to be replaced, archived (archiving only ever touches an
       APPROVED document, which a withdrawal can never be — the two cannot co-occur at all) or
       re-derived as merely "deleted". Checking it first costs nothing on a record where it is
       absent, and guarantees a deliberate self-delete is never masked by a stale value elsewhere.

       ⚠⚠ `judgeable` ALONE IS NOT ENOUGH — A FRESH UPLOAD CAN LOOK DELETED FOR A FEW MINUTES BY
       DESIGN, AND `settling` IS WHAT STOPS THAT READING AS "DELETED" (OR AS THE ALARMING "NOT
       CHECKED"). Since the 2026-09-18 tag/approve-by-proxy change, the join key
       (`SubmissionFileId`) is applied to the LIVE document only by `CRS — Apply pending tags`,
       which fires on a poll — not at the moment `writeSubmissionRecord` creates this row. So for
       the stretch between "the record exists" and "the flow has stamped the live item", a
       FULLY-COMPLETED library read genuinely finds no matching live row, even though nothing is
       actually wrong. `recentlyUploaded` catches exactly that window (anything within
       `RECENT_UPLOAD_GRACE_MS` of `uploadedAt`) and reports it as `settling` — distinct from
       `unknown`, which means a READ FAILED and the true status could genuinely be anything.
       `settling` means every read succeeded and simply found nothing yet, which is the one case
       where "Pending" is not a guess: nothing could have decided this file in the last few minutes
       when it cannot even be found yet. Fails the SAME safe direction either way: understating
       costs a glance at a badge that clears itself; overstating "deleted" tells an uploader their
       just-submitted document was destroyed. */
    const state: RecordState = r.withdrawnAt !== undefined
      ? "withdrawn"
      : r.replacedAt !== undefined
        ? "cancelled"
        : r.archivedAt !== undefined
          ? "archived"
          : !judgeable
            ? "unknown"
            : recentlyUploaded(r, now)
              ? "settling"
              : "deleted";
    if (state === "deleted") deleted += 1;
    else if (state === "cancelled") cancelled += 1;
    else if (state === "archived") archived += 1;
    else if (state === "withdrawn") withdrawn += 1;
    else if (state === "settling") settling += 1;
    else unknown += 1;
    rows.push(rowFromRecord(r, state));
  }

  return { rows, live: liveCount, deleted, cancelled, archived, withdrawn, settling, unknown };
}

/**
 * How many rows are live / gone / unchecked.
 *
 * Separate from `statusCounts` in `submissionGroups.ts`, which counts approval outcomes. A deleted
 * file HAS no approval outcome any more, and folding it into one of those three would report it as
 * pending or approved — a statement about a document that is not there.
 */
export function recordCounts(
  rows: readonly MergedRow[],
): { live: number; deleted: number; cancelled: number; archived: number; withdrawn: number; settling: number; unknown: number } {
  let deleted = 0;
  let cancelled = 0;
  let archived = 0;
  let withdrawn = 0;
  let settling = 0;
  let unknown = 0;
  let liveCount = 0;
  for (const r of rows ?? []) {
    if (r.recordState === "deleted") deleted += 1;
    else if (r.recordState === "cancelled") cancelled += 1;
    else if (r.recordState === "archived") archived += 1;
    else if (r.recordState === "withdrawn") withdrawn += 1;
    else if (r.recordState === "settling") settling += 1;
    else if (r.recordState === "unknown") unknown += 1;
    else liveCount += 1;
  }
  return { live: liveCount, deleted, cancelled, archived, withdrawn, settling, unknown };
}

/* ── The status line's record half ─────────────────────────────────────────────
 *
 * ⚠ THREE STATES HAVE BEEN LEFT OUT OF THAT LINE, ONE PER STATE ADDED. `cancelled` on 2026-08-28
 * and `archived` on 2026-09-03 each shipped with `recordCounts` counting them and the status column
 * never mentioning them — so a submission whose every file was replaced, and later one whose every
 * file was archived, rendered a **blank Status cell**. On the page whose whole job is saying what
 * became of a file, a blank cell reads as the page being broken.
 *
 * A `Record` over the union is what stops a fourth: adding a member to `RecordState` is a COMPILE
 * ERROR here until it is given a word. That is the same guard `RequestStatus` uses for its pill
 * colours, and it is worth more than a comment asking the next person to remember.
 */
export const RECORD_STATE_LABEL: Record<Exclude<RecordState, "live">, string> = {
  deleted: "deleted",
  // Not "cancelled": the file was superseded by a newer upload, and "cancelled" reads as withdrawn.
  cancelled: "replaced",
  archived: "archived",
  // THE ONE state that literally says "cancelled" — see the `RecordState` doc comment for why this
  // is deliberately not the SAME word as the `cancelled` type value two lines above, which means
  // "replaced" on screen. Client's own wording for a PIC's direct staging delete.
  withdrawn: "cancelled",
  // Deliberately the SAME word a live join would show. See `RecordState`'s own comment: every read
  // succeeded, the record is simply too young for a join to be expected yet, and nothing could have
  // decided it in that time — "pending" is the honest answer, not an optimistic guess.
  settling: "pending",
  // Worded differently on purpose — it means the libraries could not all be read, so those files may
  // be perfectly fine. Never "missing".
  unknown: "not checked",
};

/** The order they read in — worst news first, doubt last, `settling` last of all because it is not
 * really doubt at all, just a join that has not caught up yet. */
const RECORD_STATE_ORDER: Array<Exclude<RecordState, "live">> = [
  "deleted", "withdrawn", "cancelled", "archived", "unknown", "settling",
];

/** One number per state. `MergeResult` satisfies it, so a merge's own totals can be passed straight in. */
export type RecordStateCounts = Record<RecordState, number>;

/**
 * The non-approval half of a submission's status line, e.g. `["2 replaced", "1 archived"]`.
 *
 * Zero counts are dropped, so a submission with nothing gone says nothing — the approval counts
 * beside it carry the line. Both are empty only when the submission holds no files at all.
 */
export function recordStateParts(counts: RecordStateCounts): string[] {
  return RECORD_STATE_ORDER
    .filter((k) => counts[k] > 0)
    .map((k) => `${counts[k]} ${RECORD_STATE_LABEL[k]}`);
}

/** Only live rows have an approval outcome — see `recordCounts`. */
export function liveRowsOnly(rows: readonly MergedRow[]): MergedRow[] {
  return (rows ?? []).filter((r) => r.recordState === undefined || r.recordState === "live");
}

/**
 * Only the rows that were moved to the archive — the `Archive` tab (client, 2026-09-03:
 * *"Add another tab call Archive so they can tell filter to archive files."*).
 *
 * ⚠ IT CANNOT GO THROUGH `filterByTab`, and that is why this exists rather than a fourth string
 * comparison there. That function matches `r.status === tab` — Pending / Approved / Rejected — and
 * `archived` is a RECORD state, not an approval outcome. A record row carries `status: "Pending"`
 * only because the type demands a value, so a status-based Archive tab would have matched nothing
 * while a `Pending` tab that forgot to exclude them would list a file nobody can open as awaiting
 * approval. The two live in different modules for the same reason: `Submission` knows nothing about
 * records, and importing `MergedRow` into `mySubmissions.ts` would be circular.
 */
export function archivedRowsOnly(rows: readonly MergedRow[]): MergedRow[] {
  return (rows ?? []).filter((r) => r.recordState === "archived");
}

/* ── The metadata snapshot, split the way the batch view is split ──────────────
 *
 * ⚠ A GONE FILE'S DETAILS ARE THE ONLY DETAILS IT HAS, so they must land in the SAME two places a
 * live file's do: the folder tiers in the "Document folder information" card at the top, and the
 * per-document fields on the file's own card. Rendering the whole snapshot on the file card — which
 * is what happened until 2026-09-03 — put Business Segment, Department and Unit inside the file card
 * and left the folder card holding nothing but Location. Client: *"the archived files wont show
 * Document folder information like a normal file does."*
 *
 * The partition is by KEY, and it is exact rather than a guess: both upload screens write the
 * per-document fields under these fixed labels (`Form.tsx` ~2823, `BulkUpload.tsx` ~2482) and every
 * OTHER key in a snapshot is a tier, written as `snapshot[level.column] = level.label`. So "anything
 * not in this list is part of the folder" is true by construction, and a segment onboarded years
 * from now with tiers nobody has written code for partitions correctly with no change here.
 *
 * ⚠ THE LIST IS THE UNION OF BOTH WRITERS, `Bulk import` included. Leaving one screen's key out
 * would file it as a folder tier — so a bulk-imported document would claim its destination folder
 * had a level called "Bulk import".
 */
export const SNAPSHOT_FILE_KEYS: string[] = [
  "Document name",
  "Project name",
  "Vendor/Customer",
  "Document date",
  "Confidentiality",
  "Remark",
  "Legally privileged",
  "Bulk import",
];

/** A snapshot row, shaped like `DetailRow` so both views render it with no conversion. */
export interface SnapshotRow {
  label: string;
  value: string;
}

const filledKeys = (snap: Record<string, string> | undefined): string[] =>
  Object.keys(snap ?? {}).filter((k) => ((snap ?? {})[k] ?? "").trim().length > 0);

/**
 * The FOLDER half — the tiers, in the order they were written, which is chain order.
 *
 * Insertion order is deliberate and is what makes this read like the upload form: the writers loop
 * `levelSelections` top down, so Business Segment comes before Department before Unit with nothing
 * here having to know their names.
 */
export function snapshotFolderRows(snap: Record<string, string> | undefined): SnapshotRow[] {
  return filledKeys(snap)
    .filter((k) => SNAPSHOT_FILE_KEYS.indexOf(k) === -1)
    .map((k) => ({ label: k, value: (snap ?? {})[k] }));
}

/**
 * The FILE half — the per-document fields, in `SNAPSHOT_FILE_KEYS` order.
 *
 * Fixed order rather than insertion order, because these are the form's own fields and an uploader
 * checking a record years later should meet them in the order they filled them in. Blank ones are
 * dropped, so an optional Project Name that was never given does not leave an empty row.
 */
export function snapshotFileRows(snap: Record<string, string> | undefined): SnapshotRow[] {
  const filled = filledKeys(snap);
  return SNAPSHOT_FILE_KEYS
    .filter((k) => filled.indexOf(k) > -1)
    .map((k) => ({ label: k, value: (snap ?? {})[k] }));
}

/**
 * The columns added on 2026-08-28, named once.
 *
 * Used to build the legacy `$select` by subtraction, and by the reader to decide what a 400 means.
 */
export const REPLACEMENT_COLUMNS = ["ReplacedAt", "ReplacedBy"];

/** The column added on 2026-09-03, named once. Its own rung on the ladder — see below. */
export const ARCHIVE_COLUMNS = ["ArchivedAt"];

/** The columns added on 2026-09-11, named once. The NEWEST-until-tagging, so the FIRST of the
 * pre-2026-09-18 group to drop on a 400. */
export const WITHDRAWAL_COLUMNS = ["WithdrawnAt", "WithdrawnBy"];

/** The columns added on 2026-09-18, named once. NOW the newest — a site that has not run Folder
 * Reconciliation since tag-by-proxy shipped (i.e. every currently-live site, the day this ships)
 * has no `TagPayload`/`TagStatus`/`TagError` columns at all, so asking for them unconditionally
 * 400s the WHOLE read (gotcha #11) and takes every record off My Submissions with it — the exact
 * failure this module exists to prevent, cited in the file's own header. `RECORD_READ_SELECT_NO_
 * TAGGING`, immediately below, is that rung, and it is the FIRST one dropped: see its own comment
 * for why it has to sit above `RECORD_READ_SELECT_NO_WITHDRAWAL` in the chain rather than beside it. */
export const TAG_COLUMNS = ["TagPayload", "TagStatus", "TagError"];

/** Everything the record read asks for, including the built-ins it needs. */
export const RECORD_READ_SELECT = [
  "Id", "SubmissionRef", "BatchRef", "SubmissionFileId", "ItemUniqueId", "FileName", "ItemPath",
  "LibraryTitle", "UploadedBy", "UploadedAt", "MetadataSnapshot", "Source",
  /* ⚠ ADDED 2026-08-28, AND THE READ MUST RETRY WITHOUT THEM. One unknown field name fails the
     WHOLE $select (gotcha #11), so on any site whose CRS Submissions list predates this change,
     asking for them unconditionally would take EVERY record off My Submissions — not just the
     replacement state. `readSubmissionRecords` drops them and re-asks on a 400, exactly as the
     Requests page does for `Stage`. */
  ...REPLACEMENT_COLUMNS,
  ...ARCHIVE_COLUMNS,
  ...WITHDRAWAL_COLUMNS,
  ...TAG_COLUMNS,
].join(",");

/**
 * The same read with only the 2026-09-18 columns dropped — the NEWEST rung, and now the FIRST one
 * tried on a 400.
 *
 * ⚠ WHY TAGGING HAS TO BE THE TOP OF THE CHAIN, NOT FOLDED IN BESIDE ANY OTHER RUNG. Every rung
 * below this one is reached by a site that is MISSING some OLDER optional column — one that
 * predates 2026-09-18 and therefore predates this rung too. Such a site is, by definition, also
 * missing `TagPayload`/`TagStatus`/`TagError`, since those did not exist yet either. So EVERY
 * lower rung must drop the tag columns as well, or the retry ladder just re-asks for a field that
 * is still going to 400 — which is exactly the bug this rung fixes: `RECORD_READ_SELECT_NO_
 * WITHDRAWAL`, `_NO_ARCHIVE` and `_LEGACY` used to derive straight from `RECORD_READ_SELECT`, so
 * all three still carried `TAG_COLUMNS` and would 400 again on any site not yet reconciled for
 * tagging — which, the day this shipped, was every live site. Fixed by re-pointing `RECORD_READ_
 * SELECT_NO_WITHDRAWAL`'s own source to THIS rung instead, so the drop propagates transitively
 * down the whole chain with no other rung's derivation line needing to change.
 */
export const RECORD_READ_SELECT_NO_TAGGING = RECORD_READ_SELECT.split(",")
  .filter((c) => TAG_COLUMNS.indexOf(c) === -1)
  .join(",");

/**
 * The same read with only the 2026-09-11 columns dropped too — a NEW top rung.
 *
 * ⚠ NEWEST DROPS FIRST, same reasoning as the archive rung below it: a site that has `ReplacedAt`
 * and `ArchivedAt` but not `WithdrawnAt`/`WithdrawnBy` must not lose all three at once. Derived by
 * subtraction so the ladder cannot drift.
 *
 * ⚠ DERIVED FROM `RECORD_READ_SELECT_NO_TAGGING`, NOT FROM `RECORD_READ_SELECT` DIRECTLY (as of
 * 2026-09-18) — deriving from the full select would silently reintroduce `TAG_COLUMNS` into every
 * rung below this one. See `RECORD_READ_SELECT_NO_TAGGING`'s own comment for why the tag drop has
 * to sit above this rung rather than be repeated in it.
 */
export const RECORD_READ_SELECT_NO_WITHDRAWAL = RECORD_READ_SELECT_NO_TAGGING.split(",")
  .filter((c) => WITHDRAWAL_COLUMNS.indexOf(c) === -1)
  .join(",");

/**
 * The same read with only the 2026-09-03 column ALSO dropped — the MIDDLE rung.
 *
 * ⚠ THIS RUNG EXISTS SO A SITE THAT HAS `ReplacedAt` BUT NOT `ArchivedAt` DOES NOT LOSE BOTH. A
 * two-rung ladder would drop straight to the 2026-08-28 fallback, and every replaced file on such a
 * site would silently revert to reading "Deleted" — a state the client specifically asked us to stop
 * showing. Exactly the trap `RevokedBy` taught on `CRS Requests` (2026-08-30), where a single retry
 * covering `Stage` was not enough once a second optional column arrived.
 *
 * Derived from the rung ABOVE it, never from the full select — dropping straight from `RECORD_READ_
 * SELECT` would reintroduce the 2026-09-11 columns this rung is supposed to have already lost. That
 * chain also means this rung — and every one below it — transitively excludes the 2026-09-18 tag
 * columns too, since `RECORD_READ_SELECT_NO_WITHDRAWAL` itself now derives from
 * `RECORD_READ_SELECT_NO_TAGGING` rather than from the full select directly.
 */
export const RECORD_READ_SELECT_NO_ARCHIVE = RECORD_READ_SELECT_NO_WITHDRAWAL.split(",")
  .filter((c) => ARCHIVE_COLUMNS.indexOf(c) === -1)
  .join(",");

/**
 * The same read with the 2026-08-28 columns dropped too — every optional column gone.
 *
 * ⚠ THE FALLBACK IS THE WHOLE POINT. A site whose `CRS Submissions` list predates that change has
 * none of these columns, and one unknown field name fails the ENTIRE `$select` — so without this,
 * adding optional columns would empty My Submissions of every record on every unreconciled site. The
 * `Stage` column on `CRS Requests` taught this on 2026-08-20; the difference is that this list is
 * read by an UPLOADER rather than an approver, so the blast radius is everybody.
 *
 * DERIVED by subtraction, never written out again: a second literal list is how the two drift, and
 * the drifting one would be the rarely-exercised fallback.
 */
export const RECORD_READ_SELECT_LEGACY = RECORD_READ_SELECT_NO_ARCHIVE.split(",")
  .filter((c) => REPLACEMENT_COLUMNS.indexOf(c) === -1)
  .join(",");

/** What a record row is written as. One field per column, plus `Title`. */
export type RecordPayload = Record<string, string>;

/**
 * Build the write payload.
 *
 * ⚠ `UploadedAt` IS ISO, and gotcha #1 does NOT apply here. The `M/D/YYYY` locale format belongs to
 * `validateUpdateListItem`, which parses in the site's locale; a plain `/items` POST goes through the
 * OData layer and answers a locale string with *"Cannot convert a primitive value to the expected
 * type 'Edm.DateTime'"* — a 400 naming the type but not the field. A later `$filter` needs ISO too,
 * so the write and the filter share ONE format and cannot be mismatched. Learned on the audit log.
 *
 * `Title` carries the file name, so the list view is readable without anyone building one — the row
 * is a record somebody may have to read in five years.
 *
 * Blank optional values are OMITTED rather than sent as `""`. A DateTime column rejects an empty
 * string and takes the WHOLE row with it, which is how a blank expiry once lost a request row.
 */
export function buildRecordPayload(r: Omit<SubmissionRecord, "itemId">): RecordPayload {
  const out: RecordPayload = {
    Title: (r.fileName ?? "").slice(0, 255),
    SubmissionRef: r.submissionRef ?? "",
    BatchRef: r.batchRef ?? "",
    SubmissionFileId: r.fileId ?? "",
    FileName: r.fileName ?? "",
    ItemPath: r.itemPath ?? "",
    LibraryTitle: r.libraryTitle ?? "",
    UploadedBy: (r.uploadedBy ?? "").toLowerCase(),
    MetadataSnapshot: encodeMetadata(r.metadata ?? {}),
    Source: r.source ?? "",
  };
  const uid = (r.uniqueId ?? "").trim();
  if (uid.length > 0) out.ItemUniqueId = uid;
  if (r.uploadedAt instanceof Date && !isNaN(r.uploadedAt.getTime())) {
    out.UploadedAt = r.uploadedAt.toISOString();
  }
  // Present-only, same as every other optional field above: a row must never claim an empty
  // TagPayload was a deliberate answer, and `TagStatus`/`TagError` unset simply means "not yet
  // processed" rather than "processed with nothing to say."
  if (r.tagPayload !== undefined) out.TagPayload = r.tagPayload;
  if (r.tagStatus !== undefined) out.TagStatus = r.tagStatus;
  if (r.tagError !== undefined) out.TagError = r.tagError;
  return out;
}

/**
 * Read one row back.
 *
 * ⚠ EVERY FIELD IS TREATED AS UNTRUSTED. SharePoint returns `null` for an empty column, a Text
 * column can hold anything, and this project has already taken a page down because a taxonomy value
 * arrived as an OBJECT and `.trim()` was called on it (2026-08-14). Nothing here throws: a record
 * whose fields read badly is still evidence that an upload happened, and `isJoinable` decides
 * whether it is usable.
 */
export function parseRecordRow(raw: Record<string, unknown>): SubmissionRecord {
  const row = raw ?? {};
  const str = (k: string): string => {
    const v = row[k];
    if (typeof v === "string") return v.trim();
    if (typeof v === "number" || typeof v === "boolean") return `${v}`;
    return "";
  };
  const idRaw = row.Id;
  const at = str("UploadedAt");
  const parsed = at.length > 0 ? new Date(at) : undefined;
  return {
    itemId: typeof idRaw === "number" ? idRaw : Number(str("Id")) || 0,
    submissionRef: str("SubmissionRef"),
    batchRef: str("BatchRef"),
    fileId: str("SubmissionFileId"),
    uniqueId: str("ItemUniqueId") || undefined,
    fileName: str("FileName"),
    itemPath: str("ItemPath"),
    libraryTitle: str("LibraryTitle"),
    uploadedBy: str("UploadedBy"),
    // An unparseable date is undefined, never `Invalid Date`: that formats on screen as
    // "Invalid Date" and sorts unpredictably.
    uploadedAt: parsed instanceof Date && !isNaN(parsed.getTime()) ? parsed : undefined,
    metadata: decodeMetadata(str("MetadataSnapshot")),
    source: str("Source"),
    /* ⚠ AN UNPARSEABLE DATE IS `undefined`, NOT `Invalid Date` — and here that is load-bearing
       rather than cosmetic. `replacedAt !== undefined` is what decides `cancelled`, so an
       `Invalid Date` would be truthy and mark a record replaced on the strength of a value nobody
       could read. Absent, blank and malformed all mean "not replaced". */
    replacedAt: (() => {
      const raw = str("ReplacedAt");
      if (raw.length === 0) return undefined;
      const d = new Date(raw);
      return isNaN(d.getTime()) ? undefined : d;
    })(),
    replacedBy: str("ReplacedBy") || undefined,
    /* Same rule as `replacedAt`, and load-bearing for the same reason: `archivedAt !== undefined` is
       what decides `archived`, so an `Invalid Date` would be truthy and report a file as archived on
       the strength of a value nobody could read. Absent, blank and malformed all mean "not
       archived" — which degrades to the previous behaviour rather than inventing a state. */
    archivedAt: (() => {
      const raw = str("ArchivedAt");
      if (raw.length === 0) return undefined;
      const d = new Date(raw);
      return isNaN(d.getTime()) ? undefined : d;
    })(),
    /* Same rule again: `withdrawnAt !== undefined` is what decides `withdrawn`, so an `Invalid Date`
       would be truthy and mark a record withdrawn on the strength of a value nobody could read. */
    withdrawnAt: (() => {
      const raw = str("WithdrawnAt");
      if (raw.length === 0) return undefined;
      const d = new Date(raw);
      return isNaN(d.getTime()) ? undefined : d;
    })(),
    withdrawnBy: str("WithdrawnBy") || undefined,
    tagPayload: str("TagPayload") || undefined,
    tagStatus: str("TagStatus") || undefined,
    tagError: str("TagError") || undefined,
  };
}

/**
 * Serialise the metadata snapshot.
 *
 * Blank values are dropped: a snapshot exists so a deleted row can still show what was submitted, and
 * a wall of empty fields is not that. Stored as JSON in a Note column, so it needs no schema and a
 * segment with different tiers cannot outgrow it.
 */
export function encodeMetadata(meta: Record<string, string>): string {
  const out: Record<string, string> = {};
  for (const k of Object.keys(meta ?? {})) {
    const v = (meta[k] ?? "").trim();
    if (v.length > 0) out[k] = v;
  }
  return JSON.stringify(out);
}

/**
 * Read the snapshot back.
 *
 * ⚠ ANY failure answers `{}` — never throws. This is decoration on a row whose PURPOSE is to survive;
 * a malformed snapshot must cost the metadata table and nothing else. A row that vanished because its
 * JSON was bad would defeat the entire feature.
 */
export function decodeMetadata(raw: string | undefined): Record<string, string> {
  const text = (raw ?? "").trim();
  if (text.length === 0) return {};
  try {
    const parsed: unknown = JSON.parse(text);
    if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) return {};
    const out: Record<string, string> = {};
    const obj = parsed as Record<string, unknown>;
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      if (typeof v === "string") out[k] = v;
      else if (typeof v === "number" || typeof v === "boolean") out[k] = `${v}`;
    }
    return out;
  } catch {
    return {};
  }
}
