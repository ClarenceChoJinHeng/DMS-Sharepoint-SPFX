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
  /** Set when `state === "unknown"`: what to tell the admin, instead of a silent grey checkbox. */
  reason?: string;
}

/** Nothing counted yet — the dialog's opening state, which offers no destructive option at all. */
export function unknownCounts(reason: string): SegmentCounts {
  return { state: "unknown", folders: 0, documents: 0, groupMapRows: 0, reason };
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
  return counts.documents > 0 || counts.groupMapRows > 0;
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
export function survivorLines(deleteFolders: boolean): string[] {
  return [
    "The tier columns stay, with every document's metadata intact.",
    "The folder abbreviations stay, so re-creating this segment keeps the same folder names.",
    deleteFolders
      ? "The folders go to the recycle bin and can be restored for 93 days."
      : "Every folder and every document stays exactly where it is.",
  ];
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
  if (deleteFolders && counts.state === "counted") {
    parts.push(`${counts.folders} folder${counts.folders === 1 ? "" : "s"}`);
    if (counts.documents > 0) {
      parts.push(`${counts.documents} document${counts.documents === 1 ? "" : "s"}`);
    }
  }
  if (parts.length === 0) return "Deletes the segment only.";
  const last = parts.pop() as string;
  return parts.length === 0
    ? `Deletes the segment and ${last}.`
    : `Deletes the segment, ${parts.join(", ")} and ${last}.`;
}
