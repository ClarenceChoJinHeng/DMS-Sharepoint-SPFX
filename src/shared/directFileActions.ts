import { FileRights } from "./dmsFolderMap";

/**
 * Should THIS viewer see a direct Delete/Share button for THIS document, reached by clicking a
 * file's Name column in Documents/HC Documents/Archive/HC Archive?
 *
 * Spec: docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md
 *
 * ⚠ ARCHIVE OVERRIDES THE PROBE OUTRIGHT, REGARDLESS OF WHAT IT ANSWERS. `probeFileRights`'s own
 * documentation says it answers permission only, never appropriateness — "an archived document is
 * read-only to everybody by design, and the caller must enforce that." A system admin's Full
 * Control would otherwise make the probe say "granted" on an archived file; this function is what
 * stops that from ever reaching a button.
 *
 * Only a literal "granted" offers an action — "denied", "missing", and "unknown" are all treated
 * identically as "no", never as "probably fine." Understating what someone can do costs them one
 * extra click through the ordinary request flow elsewhere; overstating it would offer a delete or
 * share that then fails at the write, which is the worse of the two directions.
 */
export interface DirectActions {
  canDelete: boolean;
  canShare: boolean;
}

export function directActionsFor(
  rights: FileRights,
  isArchived: boolean,
): DirectActions {
  if (isArchived) return { canDelete: false, canShare: false };
  return {
    canDelete: rights.remove === "granted",
    canShare: rights.share === "granted",
  };
}
