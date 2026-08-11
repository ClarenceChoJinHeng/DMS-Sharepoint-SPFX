// Subtree migration: deciding which below-Unit folders sit at the wrong depth after a
// structure change, and where each one belongs.
//
// Pure and SPFx-free — no imports from @microsoft/*, so every decision here is testable
// in plain Jest without a tenant. Same approach as folderChain.ts and approvalQueue.ts.
//
// Spec: docs/superpowers/specs/2026-08-11-subtree-migration-design.md
//
// The whole point of this module is that the tool DERIVES what moved rather than asking.
// An admin asked "which level is new?" can answer wrongly, and a wrong answer here files a
// unit's entire history under a value nobody chose.
import { sanitizeFolderSegment } from "./formModel";

/** Case-insensitive name comparison — SharePoint folder names preserve case but do not distinguish it. */
function sameName(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

/**
 * One below-Unit tier as it applies to ONE unit, after the tiers that do not apply have
 * been dropped.
 *
 * `chainIndex` is the tier's position in the mode row's full below-Unit chain; `options`
 * are the sanitized-comparable folder names valid for THIS unit. The two differ because a
 * cascading tier (no `termSet`) draws its values from the term above, so a tier can apply
 * to one unit and not another — the client confirmed on 2026-08-10 that not every Unit has
 * SubUnits.
 */
export interface EffectiveTier {
  chainIndex: number;
  options: string[];
}

/**
 * Drop the tiers that do not apply to this unit.
 *
 * A cascading tier with no terms under this unit does not apply, and its files legitimately
 * sit one level shallower — `decideTier` in folderChain.ts already skips it at upload time.
 * If this did not drop it too, every one of that unit's folders would look misplaced and
 * the tool would offer to move a correctly filed history into a folder whose name it would
 * have to invent.
 *
 * A tier whose options could not be LOADED must never reach here as an empty list: the
 * caller blocks that unit instead. Empty means "no such values exist"; unknown means "do
 * not touch this unit", and conflating them is the difference between skipping a tier and
 * relocating a unit's documents because of a network error.
 */
export function effectiveTiers(optionsByTier: string[][]): EffectiveTier[] {
  const out: EffectiveTier[] = [];
  (optionsByTier ?? []).forEach((options, chainIndex) => {
    if ((options ?? []).length > 0) out.push({ chainIndex, options });
  });
  return out;
}

/** Where a folder found directly under a Unit actually belongs. */
export type Placement =
  /** Its name is a valid value for the first tier — already in the right place. */
  | { kind: "ok" }
  /** Its name belongs `levels` tiers deeper; that many new ancestor folders are needed. */
  | { kind: "misplaced"; levels: number }
  /** Matches no tier at any depth. Reported, never moved. */
  | { kind: "stray" };

/**
 * Classify one folder sitting directly under a Unit.
 *
 * A name valid at tier 0 is `ok` even when it is ALSO valid deeper. That ambiguity is real
 * — `2026` could plausibly be a SubUnit and a Year — and resolving it as "already correct"
 * is the only safe direction: the alternative moves a correctly filed folder on no evidence.
 */
export function classifyChild(name: string, tiers: EffectiveTier[]): Placement {
  const list = tiers ?? [];
  if (list.length === 0) return { kind: "stray" };
  const validAt = (i: number): boolean =>
    ((list[i] ?? { options: [] }).options ?? []).filter((o) => sameName(sanitizeFolderSegment(o), name)).length > 0;
  if (validAt(0)) return { kind: "ok" };
  for (let k = 1; k < list.length; k++) {
    if (validAt(k)) return { kind: "misplaced", levels: k };
  }
  return { kind: "stray" };
}

/** A destination value chosen by the administrator for one tier. */
export interface Destination {
  /** Term label as chosen. The folder name is the sanitized form of this. */
  label: string;
  /** Term GUID, written to the tier's `tidCol`. */
  id: string;
}

/** One folder to relocate, fully resolved. Nothing here needs deciding at run time. */
export interface MoveOp {
  /** Current server-relative path of the folder. */
  from: string;
  /** Destination server-relative path, including the folder's own name. */
  to: string;
  /** The folder's own name, unchanged by the move. */
  name: string;
  /**
   * Ancestor folders to ensure above it, outermost first. Ensuring is idempotent, so
   * prefixes shared by several moves in the same unit cost nothing extra.
   */
  ancestors: Array<{ path: string; chainIndex: number; destination: Destination }>;
}

/** Everything the plan knows about one unit. */
export interface UnitPlan {
  /** Server-relative path of the Unit folder. */
  unitPath: string;
  moves: MoveOp[];
  /** Folders matching no tier — surfaced so an admin can look, never touched. */
  strays: string[];
  /** Chain indexes of the tiers needing a destination, for the pickers. */
  neededTiers: number[];
  /** Set when the unit is skipped, with the reason to show. */
  skipped?: string;
}

/**
 * Build the moves for one unit.
 *
 * `destinations` is indexed by position in `tiers` (the EFFECTIVE list), not by chain
 * index — the caller holds the admin's choices per unit and must index them the same way.
 * A tier with no chosen destination stops the unit: there is no safe default, because
 * picking the first option files a unit's whole history under a value nobody chose.
 */
export function planUnit(
  unitPath: string,
  children: string[],
  tiers: EffectiveTier[],
  destinations: Array<Destination | undefined>,
): UnitPlan {
  const plan: UnitPlan = { unitPath, moves: [], strays: [], neededTiers: [] };
  const misplaced: Array<{ name: string; levels: number }> = [];

  for (const name of children ?? []) {
    const where = classifyChild(name, tiers);
    if (where.kind === "stray") plan.strays.push(name);
    else if (where.kind === "misplaced") misplaced.push({ name, levels: where.levels });
  }
  if (misplaced.length === 0) return plan;

  // The deepest displacement decides how many destinations must be chosen. A unit adrift by
  // different amounts is unusual but possible — two structure edits before one migration —
  // and the shallower folders simply use a shorter prefix of the same choices.
  let deepest = 0;
  for (const m of misplaced) deepest = Math.max(deepest, m.levels);
  for (let i = 0; i < deepest; i++) plan.neededTiers.push(tiers[i].chainIndex);

  for (let i = 0; i < deepest; i++) {
    if (!destinations[i]) {
      plan.skipped = "no destination chosen";
      return plan;
    }
  }

  for (const item of misplaced) {
    const ancestors: MoveOp["ancestors"] = [];
    let path = unitPath;
    for (let i = 0; i < item.levels; i++) {
      const dest = destinations[i] as Destination;
      const segment = sanitizeFolderSegment(dest.label);
      // A destination whose sanitized label is empty would build `unit//2024`, which
      // SharePoint collapses to `unit/2024` — the move would do nothing and report success.
      // Refuse the unit rather than perform a no-op that looks like a migration.
      if (!segment) {
        return {
          ...plan,
          moves: [],
          skipped: `the chosen value for level ${tiers[i].chainIndex + 1} is not a usable folder name`,
        };
      }
      path = `${path}/${segment}`;
      ancestors.push({ path, chainIndex: tiers[i].chainIndex, destination: dest });
    }
    plan.moves.push({
      from: `${unitPath}/${item.name}`,
      to: `${path}/${item.name}`,
      name: item.name,
      ancestors,
    });
  }
  return plan;
}

/** A file as read back from the library, reduced to what the backfill compares. */
export interface FileRow {
  /** Server-relative path including the file name. */
  path: string;
  /** Current value of each tier's label column, keyed by internal name. */
  values: Record<string, string>;
}

/** One file's missing or contradicted tier metadata. */
export interface StampNeed {
  path: string;
  /** Tiers whose column disagrees with the path, with the value the path implies. */
  fields: Array<{ chainIndex: number; label: string }>;
}

/**
 * Find files whose tier columns disagree with where they actually sit.
 *
 * The PATH wins. It is where the document is; the column only describes it, and a
 * description that disagrees is the thing to fix. This is what makes the migration
 * re-runnable: a run that moves folders and then fails while stamping leaves work a second
 * run finds by itself, with no record kept anywhere.
 *
 * It also repairs files uploaded in the window between saving a structure and migrating it,
 * and files left mis-tagged by a hand-edited chain — the `Function` column that received
 * `Finance|cfc4837f…` because its JSON had no `tidCol` is exactly this shape of damage.
 *
 * A path segment matching no option for its tier is IGNORED rather than written: that
 * folder is a stray, `planUnit` already reports it, and stamping would put a value in the
 * column that no term backs.
 */
export function backfillNeeds(
  unitPath: string,
  files: FileRow[],
  tiers: EffectiveTier[],
  labelColFor: (chainIndex: number) => string | undefined,
): StampNeed[] {
  const out: StampNeed[] = [];
  const prefix = `${unitPath}/`;
  for (const file of files ?? []) {
    if (file.path.indexOf(prefix) !== 0) continue;
    const parts = file.path.slice(prefix.length).split("/");
    // The last part is the file name; everything before it is a folder tier.
    const folders = parts.slice(0, parts.length - 1);
    const fields: StampNeed["fields"] = [];
    for (let i = 0; i < folders.length && i < tiers.length; i++) {
      const tier = tiers[i];
      const match = tier.options.filter((o) => sameName(sanitizeFolderSegment(o), folders[i]))[0];
      if (!match) continue;
      const col = labelColFor(tier.chainIndex);
      if (!col) continue;
      if (!sameName(file.values[col] ?? "", match)) {
        fields.push({ chainIndex: tier.chainIndex, label: match });
      }
    }
    if (fields.length > 0) out.push({ path: file.path, fields });
  }
  return out;
}

/** Totals for the dry run. Counted here so the summary cannot drift from the plan. */
export function planTotals(plans: UnitPlan[]): {
  unitsWithMoves: number;
  moves: number;
  strays: number;
  skippedUnits: number;
} {
  let moves = 0;
  let strays = 0;
  let skippedUnits = 0;
  let unitsWithMoves = 0;
  for (const p of plans ?? []) {
    moves += p.moves.length;
    strays += p.strays.length;
    if (p.skipped) skippedUnits++;
    if (p.moves.length > 0) unitsWithMoves++;
  }
  return { unitsWithMoves, moves, strays, skippedUnits };
}
