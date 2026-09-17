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
 * May the SIMPLE re-code action run — recycle the (confirmed-empty) old tree, write the new name?
 *
 * ONLY when the count succeeded AND it found zero documents. `folders` may still be non-zero — an
 * empty tree that reconciliation has already built for a segment nobody has uploaded into yet is
 * exactly the ordinary case this exists for — but `documents` must be zero, because a document left
 * behind under the OLD folder name becomes invisible the moment reconciliation stops walking it.
 *
 * ⚠ THIS IS NO LONGER THE ONLY WAY A SEGMENT CAN BE RE-CODED. Since 2026-09-16, a segment WITH
 * documents can be re-coded too — `performLiveSegmentRecode` (spSegmentRecode.ts) renames the top
 * folder IN PLACE in every library instead, the same proven mechanism reconciliation already uses
 * for Department/Unit folders, which preserves every document, version, approval status and
 * permission inside. `canOfferAnyRecode` below is the gate for THAT path. This function keeps its
 * original, narrower meaning — "is the CHEAP recycle-and-rebuild path available" — because the two
 * write sequences are genuinely different and a caller must know which one to run.
 *
 * Fails CLOSED on an unreadable count, the same direction as `canOfferFolderDelete` and for the
 * identical reason: an unconfirmed "probably empty" here strands a document tree, and the person who
 * discovers it later is an uploader, not an admin.
 */
export function canOfferRecode(counts: SegmentCounts): boolean {
  return counts.state === "counted" && counts.documents === 0;
}

/**
 * May EITHER re-code action run — the count merely needs to have succeeded, whatever it found.
 *
 * The UI checks this FIRST to decide whether to show the recode form at all; `canOfferRecode` then
 * decides, only once the form is showing, WHICH write sequence a submit should run.
 */
export function canOfferAnyRecode(counts: SegmentCounts): boolean {
  return counts.state === "counted";
}

/**
 * Why the action is refused, when it is. Blank means `canOfferAnyRecode` is true and the form may
 * show — for EITHER path, empty or live.
 *
 * ⚠ NO LONGER FIRES FOR "this segment holds documents" — that used to be a refusal and is now a
 * VALID, DIFFERENT path (the live in-place rename), so a non-zero count is not refused here any
 * more. The only thing this function still refuses is not knowing the count at all: guessing either
 * way — empty or live — risks running the wrong write sequence for what the segment actually holds.
 */
export function recodeRefusalReason(counts: SegmentCounts): string {
  if (counts.state === "unknown") {
    return (
      `Could not confirm what this segment holds — ${counts.reason}. Re-coding is refused rather ` +
      `than guessed: which write sequence is safe (an empty-tree recycle, or an in-place rename ` +
      `that preserves real documents) depends on knowing that first.`
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

/**
 * One-line summary for the LIVE (in-place) rename — a genuinely different action from
 * `recodeSummary` above, so it says a different thing: nothing was recycled or rebuilt, the same
 * folders were simply renamed in place, contents intact.
 */
export function recodeSummaryLive(oldFolder: string, newFolder: string, counts: SegmentCounts): string {
  const folder = sanitizeFolderSegment(newFolder).trim();
  const docCount = counts.state === "counted" ? counts.documents : 0;
  return (
    `Top folder renamed from "${oldFolder}" to "${folder}" in every library, including the archive. ` +
    `${docCount} document${docCount === 1 ? "" : "s"} moved with it — nothing was recycled or ` +
    `re-created, only relabelled. Run Folder Reconciliation to confirm everything still resolves.`
  );
}
