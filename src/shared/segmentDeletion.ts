/**
 * Deleting a segment — the rules that decide what may be offered and what must be confirmed.
 *
 * Spec: docs/superpowers/specs/2026-08-14-delete-segment-design.md
 *
 * Pure and SPFx-free, the same split as newSegment.ts / segmentReadiness.ts: everything that
 * decides whether a destructive option APPEARS is unit-tested, because the failure here is not a
 * wrong number on a screen — it is an archive nobody could confirm was empty being deleted.
 *
 * The module rests on one distinction: **"we counted, and it is zero" is not the same as "we could
 * not count"**. That is the `unknown` ≠ empty rule used everywhere else in this codebase, but with
 * the polarity deliberately reversed — see `canOfferFolderDelete`.
 */

/** Did the count succeed? `unknown` means a read failed, NOT that there is nothing there. */
export type CountState = "counted" | "unknown";

export interface SegmentCounts {
  state: CountState;
  /** Folders beneath the segment's top folder, across both libraries. 0 when `unknown`. */
  folders: number;
  /** Documents beneath it, across both libraries. 0 when `unknown`. */
  documents: number;
  /**
   * Group Map rows carrying this segment. Counted separately, because a segment can have readable
   * rows while its folders are unreadable, and the other way round.
   */
  groupMapRows: number;
  /**
   * `CRS Term Abbreviation` rows whose live term tree still resolves under this segment. Retiring
   * ALWAYS deletes these now (2026-09-11, client: "no need checkbox as an option, make it
   * mandatory") — there is no opt-out, so this count exists purely so the confirmation screen can
   * say what is about to go. See docs/superpowers/specs/2026-09-11-retire-deletes-abbreviations-design.md.
   */
  abbreviationRows: number;
  /**
   * Documents in Archive/ArchiveHC for this segment, counted SEPARATELY from `documents` — see
   * `canDeleteArchive`'s own comment for why the two must never be merged into one number.
   * `undefined` means the count could not be established (including "there is no archive on this
   * site at all", which the caller distinguishes separately via `archiveAvailable()`) — never
   * assume empty from a missing value. Only an explicit `0` means confirmed empty.
   */
  archiveDocuments?: number;
  /** Set when `state === "unknown"`: what to tell the admin, instead of a silent grey checkbox. */
  reason?: string;
}

/** Nothing counted yet — the dialog's opening state, which offers no destructive option at all. */
export function unknownCounts(reason: string): SegmentCounts {
  return { state: "unknown", folders: 0, documents: 0, groupMapRows: 0, abbreviationRows: 0, reason };
}

/**
 * May the "also delete the folders" option be offered?
 *
 * ONLY when the count succeeded. This is the one place in this codebase where an unreadable result
 * fails CLOSED rather than open. Elsewhere — the provisioned-segment filters, the allowed file
 * types, the Folder Map read — a failed read must offer everything, because failing closed there
 * takes a working form out of service site-wide.
 *
 * Here the cost runs the other way. Offering the option means an admin can delete a segment's whole
 * folder tree while the screen was unable to tell them it still held ten thousand documents.
 * Withholding it costs one manual delete in SharePoint, where the confirmation and the recycle bin
 * are somebody else's problem.
 */
export function canOfferFolderDelete(counts: SegmentCounts): boolean {
  return counts.state === "counted";
}

/**
 * May Archive/ArchiveHC be deleted TOO, alongside the operational folders?
 *
 * Spec: docs/superpowers/specs/2026-09-12-empty-archive-deletion-on-retire-design.md
 *
 * Only when `archiveDocuments` is the EXPLICIT number `0` — a confirmed-empty count, not merely a
 * falsy or absent one. `undefined` (unread, unreadable, or genuinely has content — see the field's
 * own comment) answers `false` here by construction, because `undefined === 0` is false in
 * JavaScript; there is no separate branch to get wrong.
 *
 * ⚠ THIS IS DELIBERATELY A SEPARATE QUESTION FROM `canOfferFolderDelete`. That one gates whether
 * the operational libraries (Staging/Documents/HC pair) may be touched at all, counted from a
 * DIFFERENT set of libraries than this one. Merging the two counts into one number was considered
 * and rejected: a segment can have real archived records while its operational folders are already
 * empty (the common, correct case today — records already archived, shell ready to retire), and
 * folding Archive's count into the same total would make that segment fail `documents === 0` and
 * silently lose the ability to clean up its empty operational folders at all. The two libraries'
 * emptiness are independent facts and must stay independent fields.
 *
 * Archive and ArchiveHC are asked about TOGETHER by whoever populates `archiveDocuments` (both
 * summed, or `undefined` if either could not be read) — see that field's own comment. This function
 * does not itself distinguish the two; it only ever sees one combined number, by design, so it
 * cannot delete one half of the pair while leaving the other.
 */
export function canDeleteArchive(counts: SegmentCounts): boolean {
  return counts.archiveDocuments === 0;
}

/**
 * Does this deletion need the segment's name typed out?
 *
 * Yes whenever something is genuinely at stake: documents present, mappings present, folders about
 * to be deleted, or a count that failed so nothing can be ruled out. A brand-new segment with
 * nothing under it gets a plain confirm — a gate that fires on the harmless case is one people
 * learn to type through without reading, which is worse than no gate at all.
 */
export function needsTypedConfirmation(counts: SegmentCounts, deleteFolders: boolean): boolean {
  if (counts.state === "unknown") return true;
  if (deleteFolders) return true;
  return counts.documents > 0 || counts.groupMapRows > 0 || counts.abbreviationRows > 0;
}

/**
 * Does what the admin typed match the segment's label?
 *
 * Trimmed and case-insensitive: this is an "are you looking at the segment you think you are"
 * check, not a spelling test. A label like `Upstream Operations Malaysia` is long enough that
 * demanding exact case would only teach people to copy and paste it, which defeats the check.
 *
 * A blank label can never be confirmed. Otherwise a mode row with an empty `ModeLabel` would be
 * deletable by typing nothing at all — in the very dialog that exists to slow the admin down.
 */
export function confirmationMatches(typed: string, label: string): boolean {
  const a = (typed ?? "").trim().toLowerCase();
  const b = (label ?? "").trim().toLowerCase();
  if (a.length === 0 || b.length === 0) return false;
  return a === b;
}

/**
 * What this deletion will NOT touch, in the words the dialog uses.
 *
 * Its own list, rather than left implied: the misreading that matters is an admin believing the
 * documents are gone when they are not — and then reporting a data-loss incident, or worse, not
 * reporting one.
 */
export function survivorLines(
  deleteFolders: boolean,
  hasArchive?: boolean,
  archiveDocuments?: number,
): string[] {
  const lines = [
    "The tier columns stay, with every document's metadata intact.",
    // ⚠ REVERSED 2026-09-11 — this line used to promise abbreviations stay. Retire now deletes
    // them unconditionally (client: "no need checkbox as an option, make it mandatory"), so the
    // dialog has to say the opposite: re-creating this segment means retyping every code.
    "This segment's folder abbreviations are deleted too — re-creating it later means retyping every code.",
    deleteFolders
      ? "The folders go to the recycle bin and can be restored for 93 days."
      : "Every folder and every document stays exactly where it is.",
  ];
  /* Said ONLY when the folders are actually going and this site HAS an archive (client's decision,
     2026-08-26: 7-year retained records outlive the segment that produced them). Stating it
     unconditionally would name a library that does not exist on most sites — and saying nothing at
     all leaves an admin who ticked "delete the folders" believing everything went, then finding
     Archive/<SEG> still standing in the next reconciliation log. `hasArchive` is optional so every
     existing caller and test is unchanged.
     ⚠ SIMPLIFIED 2026-09-13, DELIBERATELY DROPPING THE "IF RECREATED..." CONSEQUENCE — client:
     "just tell them empty folders will be removed for this segment but folders with files wont be
     removed." That consequence (recreating the segment leaves an unresolvable stray that blocks
     "Move existing folders") was true and worth saying UNTIL THIS SAME DAY, when the archive
     code-reuse guard shipped on segment CREATION (see
     docs/superpowers/specs/2026-09-13-archive-code-reuse-guard-design.md): a new segment can no
     longer be created reusing a code Archive already holds files under, so the scenario this
     warning described is now structurally impossible, not merely unlikely. Warning about a
     consequence that can no longer happen is worse than silence — it makes the admin hunt for a
     "Move existing folders" problem that will never arise from this. Down to the plain, current
     fact only: what does or does not happen to THIS segment's archive folders right now.
     `undefined` (could not be confirmed) is folded into the same "left alone" sentence as `> 0`
     (confirmed has content) — the DISPLAYED text no longer distinguishes them, though the
     underlying value still does, and `canDeleteArchive` still refuses on anything but an explicit
     `0`. Only a confirmed `0` gets the "removed too" sentence. */
  if (deleteFolders && hasArchive) {
    if (archiveDocuments === 0) {
      lines.push("This segment's archive folders are empty too, so they are removed as well.");
    } else {
      lines.push(
        "This segment's archive folders are not removed — only empty ones are removed automatically.",
      );
    }
  }
  return lines;
}

/**
 * One-line summary of what is about to be removed, for the dialog and the audit row.
 *
 * Never says "nothing": the mode row always goes, and that IS the deletion. A summary reading
 * "nothing will be deleted" for an empty segment would be untrue, and confusing next to a button
 * labelled Delete.
 */
export function deletionSummary(counts: SegmentCounts, deleteFolders: boolean): string {
  const parts: string[] = [];
  if (counts.groupMapRows > 0) {
    parts.push(`${counts.groupMapRows} folder-access mapping${counts.groupMapRows === 1 ? "" : "s"}`);
  }
  // Unconditional, unlike the folder/document lines below: abbreviation rows go whether or not
  // "also delete the folders" was ticked, so this must not be gated on `deleteFolders` too.
  if (counts.state === "counted" && counts.abbreviationRows > 0) {
    parts.push(`${counts.abbreviationRows} abbreviation${counts.abbreviationRows === 1 ? "" : "s"}`);
  }
  if (deleteFolders && counts.state === "counted") {
    parts.push(`${counts.folders} folder${counts.folders === 1 ? "" : "s"}`);
    if (counts.documents > 0) {
      parts.push(`${counts.documents} document${counts.documents === 1 ? "" : "s"}`);
    }
    // Only when CONFIRMED empty (see canDeleteArchive) — an unread or non-empty archive is left
    // alone and must not be claimed here, or an audit row would record a deletion that did not
    // happen.
    if (counts.archiveDocuments === 0) {
      parts.push("its empty archive");
    }
  }
  if (parts.length === 0) return "Deletes the segment only.";
  const last = parts.pop() as string;
  return parts.length === 0
    ? `Deletes the segment and ${last}.`
    : `Deletes the segment, ${parts.join(", ")} and ${last}.`;
}
