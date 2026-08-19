// Which Folder Map row wins when a term has more than one — and which may be deleted.
//
// ⚠ THE BUG THIS EXISTS FOR, found live 2026-08-19. Reconciliation indexed the Folder Map with
//
//     for (const r of mapRows) if (r.termGuid) mapByTerm.set(r.termGuid.toLowerCase(), r);
//
// and `Map.set` OVERWRITES. So a term with two rows collapsed to the last one: the repair pass
// repointed that row at the rebuilt folder and never saw the other, which went on pointing at a
// folder the same run had just deleted. Nothing reported it, and nothing could ever fix it.
//
// The cost is a SILENT UPLOAD REFUSAL. `lookupFolderMapping` reads `$top=1`, so it takes whichever
// row comes first; when that is the stale one, `GetFolderById` answers 404, `probeFolderUploadAccess`
// returns "missing" — deliberately conclusive, because security trimming also answers 404 — the path
// is dropped, and the uploader is told "your unit isn't ready to receive uploads yet". Their folder
// exists, their ACL is correct, and the form disagrees with both. On the rehearsal site 56 of 67
// terms were in this state and every uploader was refused.
//
// Pure, because it decides what to DELETE. Rows are derived data reconciliation can rebuild, but
// deleting the wrong one strands a unit until the next run — and the person who finds out is an
// uploader, not an administrator.

import { normalizeTermGuid } from "./segmentReadiness";

/** The fields this module needs. Structurally satisfied by `FolderMapRow`. */
export interface MapRowLite {
  itemId: number;
  termGuid: string;
  folderUniqueId: string;
  title: string;
}

export interface DuplicateGroup {
  /** Normalised term GUID the rows share. */
  termGuid: string;
  /** Every row for that term, input order preserved. */
  rows: MapRowLite[];
}

/**
 * Group rows by term, returning only the terms that have more than one.
 *
 * Keyed with `normalizeTermGuid`, NOT a bare `toLowerCase`. The upload form's own filter already
 * normalises — stripping braces and whitespace as well as case — so indexing differently here means
 * the two halves of the system can disagree about whether a term is mapped: a `{GUID}` row would
 * miss the lookup, reconciliation would create a second row, and we arrive at this same bug by
 * another route. That is a LATENT second cause, not the one that fired on 2026-08-19 — there every
 * stored GUID was already bare and lower-case.
 */
export function groupDuplicateRows(rows: MapRowLite[] | null | undefined): DuplicateGroup[] {
  const byTerm: Record<string, MapRowLite[]> = {};
  const order: string[] = [];
  for (const r of rows ?? []) {
    const key = normalizeTermGuid(r.termGuid);
    if (key.length === 0) continue;
    if (!byTerm[key]) { byTerm[key] = []; order.push(key); }
    byTerm[key].push(r);
  }
  const out: DuplicateGroup[] = [];
  for (const key of order) if (byTerm[key].length > 1) out.push({ termGuid: key, rows: byTerm[key] });
  return out;
}

export interface KeeperVerdict {
  /** The row to keep. */
  keep: MapRowLite;
  /** Rows safe to delete. EMPTY when the keeper could not be established from evidence. */
  remove: MapRowLite[];
  /**
   * True when the keeper was chosen because its folder positively resolved. False means nothing
   * resolved — the group is reported and left entirely alone.
   */
  confident: boolean;
}

/**
 * Choose which row survives, given the set of folder UniqueIds that positively resolved.
 *
 * ⚠ FAILS CLOSED, against this codebase's usual habit. Elsewhere an unreadable input fails OPEN,
 * because the cost is a form out of service for a minute. Here the cost is deleting the only row a
 * unit has, taking its upload path away until someone runs reconciliation again.
 *
 *  - at least one row resolves ⇒ keep the FIRST resolving row, delete every other row in the group;
 *  - nothing resolves ⇒ keep the LAST row and delete NOTHING. Last, because the repair pass indexes
 *    with `Map.set` and therefore repoints the last row — keeping the same one means the survivor
 *    reported here is the one that will actually be healed.
 *
 * `liveIds` must contain only ids that answered positively. A throttled or forbidden probe must NOT
 * put an id in this set: unknown is not alive, and treating it as alive would delete a good row on
 * the strength of a failed request.
 */
export function chooseKeeper(
  group: MapRowLite[] | null | undefined,
  liveIds: Set<string> | null | undefined,
): KeeperVerdict | undefined {
  const rows = group ?? [];
  if (rows.length === 0) return undefined;
  const norm = (id: string): string => (id ?? "").trim().toLowerCase();
  const liveNorm: Record<string, true> = {};
  (liveIds ?? new Set<string>()).forEach((v) => { liveNorm[norm(v)] = true; });

  for (const r of rows) {
    if (liveNorm[norm(r.folderUniqueId)]) {
      return { keep: r, remove: rows.filter((x) => x.itemId !== r.itemId), confident: true };
    }
  }
  return { keep: rows[rows.length - 1], remove: [], confident: false };
}
