/**
 * Re-coding a segment's TOP FOLDER — a physical rename, offered only for an empty segment.
 *
 * Spec: docs/superpowers/specs/2026-09-11-segment-recode-design.md
 *
 * Pure and SPFx-free, the same split as `segmentDeletion.ts` and `newSegment.ts`: everything that
 * decides whether this destructive-adjacent action may proceed is unit-tested, because the failure
 * here is not a wrong number on a screen — it is a document tree reconciliation stops recognising,
 * left behind under a folder name nothing points at any more.
 *
 * Deliberately reuses `SegmentCounts` from `segmentDeletion.ts` rather than a second count type: the
 * count is the SAME count (across the same non-archive libraries), and the fail-closed instinct on an
 * unreadable result is the same instinct for the same reason.
 */

import { sanitizeFolderSegment } from "./formModel";
import { ExistingSegment } from "./newSegment";
import { SegmentCounts } from "./segmentDeletion";

/**
 * May the re-code action actually run?
 *
 * ONLY when the count succeeded AND it found zero documents. `folders` may still be non-zero — an
 * empty tree that reconciliation has already built for a segment nobody has uploaded into yet is
 * exactly the ordinary case this exists for — but `documents` must be zero, because a document left
 * behind under the OLD folder name becomes invisible the moment reconciliation stops walking it.
 *
 * Fails CLOSED on an unreadable count, the same direction as `canOfferFolderDelete` and for the
 * identical reason: an unconfirmed "probably empty" here strands a document tree, and the person who
 * discovers it later is an uploader, not an admin.
 */
export function canOfferRecode(counts: SegmentCounts): boolean {
  return counts.state === "counted" && counts.documents === 0;
}

/**
 * Why the action is refused, when it is. Blank means `canOfferRecode` is true and the form may show.
 *
 * Never a partial path: a segment holding documents is refused outright, with the real route named,
 * rather than offered a "recode but leave the documents behind" option that would abandon them under
 * a name reconciliation no longer resolves.
 */
export function recodeRefusalReason(counts: SegmentCounts): string {
  if (counts.state === "unknown") {
    return (
      `Could not confirm this segment is empty — ${counts.reason}. Re-coding is refused rather ` +
      `than guessed: reconciliation would build a brand-new tree under the new name and leave ` +
      `anything already there behind, in a folder nothing recognises any more.`
    );
  }
  if (counts.documents > 0) {
    return (
      `This segment holds ${counts.documents} document${counts.documents === 1 ? "" : "s"}, so ` +
      `its top folder cannot be re-coded. Move or archive the documents first — re-coding a ` +
      `segment in use would leave them behind under a name reconciliation no longer walks.`
    );
  }
  return "";
}

/**
 * Does the sanitized new folder collide with ANOTHER segment's top folder?
 *
 * Case-insensitive, matching the same normalisation `fieldConflicts` applies when creating a new
 * segment — a clash here means the exact same thing it means there. The segment being recoded is
 * excluded by its own KEY, never by comparing folder values: comparing the new value against its own
 * CURRENT folder would refuse "no change at all" as though it were a clash with someone else, which
 * `folderRecodeIsNoOp` handles separately and more usefully.
 */
export function folderRecodeConflict(
  newFolder: string,
  existing: ExistingSegment[],
  selfKey: string,
): string {
  const folder = sanitizeFolderSegment(newFolder).trim();
  if (!folder) return "";
  const clash = (existing ?? []).filter(
    (e) => e.key !== selfKey && e.stagingFolder.trim().toLowerCase() === folder.toLowerCase(),
  )[0];
  return clash ? `"${folder}" is already the top folder for "${clash.label}".` : "";
}

/**
 * Is the typed value, once sanitized, the same folder the segment already has?
 *
 * A same-value "rename" would still recycle the old (empty) tree for nothing and force a
 * reconciliation run that changes nothing — so it is refused as a no-op rather than silently allowed
 * to do pointless work.
 */
export function folderRecodeIsNoOp(newFolder: string, currentFolder: string): boolean {
  return (
    sanitizeFolderSegment(newFolder).trim().toLowerCase() ===
    (currentFolder ?? "").trim().toLowerCase()
  );
}

/**
 * One-line summary of what a completed recode did, for the dialog and the audit row.
 *
 * Always names the NEXT step — reconciliation — because nothing in this action builds the new
 * folder tree; leaving that unsaid is how "it did not work" gets reported the moment someone checks
 * the library and finds no folder there yet.
 */
export function recodeSummary(oldFolder: string, newFolder: string, counts: SegmentCounts): string {
  const folder = sanitizeFolderSegment(newFolder).trim();
  const folderCount = counts.state === "counted" ? counts.folders : 0;
  return (
    `Top folder renamed from "${oldFolder}" to "${folder}". ${folderCount} empty folder` +
    `${folderCount === 1 ? "" : "s"} recycled. Run Folder Reconciliation to rebuild the tree under ` +
    `the new name.`
  );
}
