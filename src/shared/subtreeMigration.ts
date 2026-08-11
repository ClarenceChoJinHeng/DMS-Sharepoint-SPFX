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

/* ===================================================================================
 * The unified model — spec §3.0.
 *
 * Add, reorder and remove are ONE operation: work out which tier each path segment
 * belongs to, then rebuild the path in tier order. The shallow `classifyChild` model
 * above only ever detected an insertion, and it reported "nothing to move" for a
 * reorder — activating a new structure against folders still in the old shape.
 * =================================================================================== */

/** One path segment, resolved to the tier it belongs to. */
export interface TierAssignment {
  /** Position of this segment in the EXISTING path, below the unit. */
  at: number;
  /** Chain index of the tier whose options contain it. */
  chainIndex: number;
  /** The folder name as it appears on disk. */
  name: string;
}

/** What `assignSegments` could and could not resolve. */
export interface SegmentAssignment {
  assigned: TierAssignment[];
  /**
   * Segments matching no tier at any depth. Non-empty means the folder is a STRAY and must never
   * be moved — a hand-made folder or a deleted term, where a confident guess does damage.
   */
  strays: string[];
  /**
   * Segments belonging to a tier NO LONGER in the chain — a removed tier. Dropped from the
   * destination path, which is what makes sibling subtrees collapse together.
   */
  dropped: string[];
}

/**
 * Work out which tier each segment of an existing path belongs to.
 *
 * `tiers` are the target chain's tiers for this unit; `removedOptions` are the option lists of
 * tiers that USED to exist and no longer do, so a segment belonging to one can be told apart from
 * a segment belonging to nothing. That distinction is the entire difference between a deliberate
 * removal (drop the segment, collapse) and an unrecognised folder (refuse to touch it) — and
 * getting it wrong either strands a removal or silently relocates somebody's documents.
 *
 * Where a name is valid at more than one tier, **the tier it is currently at wins**. Same
 * "already-correct beats speculative" rule as `classifyChild`, applied per segment: `2026` could
 * be a Year and a SubUnit, and the reading that requires no movement is the safe one.
 */
export function assignSegments(
  segments: string[],
  tiers: EffectiveTier[],
  removedOptions?: string[][],
): SegmentAssignment {
  const out: SegmentAssignment = { assigned: [], strays: [], dropped: [] };
  const list = tiers ?? [];
  (segments ?? []).forEach((name, at) => {
    const matches: number[] = [];
    for (let i = 0; i < list.length; i++) {
      if (list[i].options.filter((o) => sameName(sanitizeFolderSegment(o), name)).length > 0) {
        matches.push(i);
      }
    }
    if (matches.length > 0) {
      // Identity first: a segment already sitting at a tier that accepts it belongs there. Only
      // then fall back to the shallowest tier that does.
      const identity = matches.filter((i) => i === at)[0];
      const chosen = identity === undefined ? matches[0] : identity;
      out.assigned.push({ at, chainIndex: list[chosen].chainIndex, name });
      return;
    }
    const removed =
      (removedOptions ?? []).filter(
        (opts) => (opts ?? []).filter((o) => sameName(sanitizeFolderSegment(o), name)).length > 0,
      ).length > 0;
    if (removed) out.dropped.push(name);
    else out.strays.push(name);
  });
  return out;
}

/** One leaf folder on disk, with the files it holds. */
export interface LeafFolder {
  /** Server-relative path. */
  path: string;
  /** Path segments below the unit, in order. */
  segments: string[];
  /** File names directly inside it. */
  files: string[];
}

/** What should happen to one leaf folder. */
export interface LeafPlan {
  leaf: LeafFolder;
  /** Destination path when one could be resolved. Equal to `leaf.path` means leave it alone. */
  to?: string;
  /** Ancestor folders to ensure, outermost first, with the tier each represents. */
  ancestors: Array<{ path: string; chainIndex: number; name: string }>;
  /** Chain indexes with no value in the existing path — the admin must choose one for each. */
  missingTiers: number[];
  /** Segments matching no tier. Non-empty means this folder is reported and never moved. */
  strays: string[];
  /** Segments whose tier was removed. Non-empty means this folder collapses into another. */
  dropped: string[];
}

/**
 * Rebuild one leaf folder's path in tier order.
 *
 * This single function is the add, the reorder AND the remove:
 *   - a tier with no segment yields `missingTiers`, filled from `chosen` (an ADD);
 *   - segments out of tier order come back sorted (a REORDER);
 *   - a segment whose tier is gone is simply absent from the result (a REMOVE, and therefore a
 *     collapse — two leaves can then share one destination, which is where filename collisions
 *     come from).
 *
 * A stray blocks the folder entirely: with one segment unexplained, any destination is a guess
 * about where somebody's documents live.
 */
export function planLeaf(
  unitPath: string,
  leaf: LeafFolder,
  tiers: EffectiveTier[],
  chosen: Record<number, Destination | undefined>,
  removedOptions?: string[][],
): LeafPlan {
  const { assigned, strays, dropped } = assignSegments(leaf.segments, tiers, removedOptions);
  const plan: LeafPlan = { leaf, ancestors: [], missingTiers: [], strays, dropped };
  if (strays.length > 0) return plan;

  // A leaf with nothing recognisable cannot be placed at all. Happens when every segment belonged
  // to a removed tier: the only destination would be the unit root, which contradicts the rule
  // below, so it is reported and left alone rather than having a whole path invented for it.
  if (assigned.length === 0) return plan;

  const byTier: Record<number, string> = {};
  for (const a of assigned) byTier[a.chainIndex] = a.name;

  // EVERY tier is built, including ones below the leaf's current depth.
  //
  // Client rule, 2026-08-12: "the file should always land at the end of the folder" — documents
  // live at the full depth of the chain, wherever a new level is inserted. This replaced a cap at
  // the leaf's deepest existing tier, which made adding a level at the BOTTOM a silent no-op: the
  // new tier sat below every leaf, so nothing looked misplaced, the change activated anyway, and
  // old documents stayed a level shallower than new ones in the same folder. "Add a level" now
  // behaves the same wherever it is added, which is what anyone using the page already assumes.
  //
  // A tier that does not apply to this unit never reaches here — `effectiveTiers` drops one whose
  // options are empty — so the optional-SubUnit case is unaffected. A genuinely shallow folder now
  // asks for a value, and "leave this unit alone" remains the way to decline.
  const wanted: Array<{ chainIndex: number; name: string }> = [];
  for (const tier of tiers ?? []) {
    const existing = byTier[tier.chainIndex];
    if (existing !== undefined) {
      wanted.push({ chainIndex: tier.chainIndex, name: existing });
      continue;
    }
    const pick = chosen[tier.chainIndex];
    const segment = pick ? sanitizeFolderSegment(pick.label) : "";
    // An unusable name would build `unit//2024`, which SharePoint collapses to `unit/2024` — a
    // move that does nothing and reports success. Treated as "not chosen".
    if (!segment) {
      plan.missingTiers.push(tier.chainIndex);
      continue;
    }
    wanted.push({ chainIndex: tier.chainIndex, name: segment });
  }
  if (plan.missingTiers.length > 0) return plan;

  let path = unitPath;
  for (let i = 0; i < wanted.length; i++) {
    path = `${path}/${wanted[i].name}`;
    // Every level except the last is an ancestor to ensure; the last IS the folder being moved.
    if (i < wanted.length - 1) {
      plan.ancestors.push({ path, chainIndex: wanted[i].chainIndex, name: wanted[i].name });
    }
  }
  plan.to = path;
  return plan;
}

/** Two or more files competing for one path after a collapse. */
export interface Collision {
  /** Destination folder path. */
  folder: string;
  /** The contested file name. */
  name: string;
  /**
   * Every claimant, in path order. `existing` marks a file ALREADY at the destination — it is not
   * moving, but it owns the name, and missing it is how a collapse overwrites a document nobody
   * was migrating.
   */
  claimants: Array<{ path: string; existing: boolean }>;
}

/**
 * Find every filename collision a set of leaf plans would produce.
 *
 * Counts files already at each destination as claimants. A collapse into an OCCUPIED folder is
 * both the case most easily missed and the most destructive: SharePoint's instinct on a name
 * clash is to overwrite, which reports success while destroying a document.
 *
 * `filesAt` answers "what is already in this destination folder", keyed by folder path; a folder
 * that does not exist yet simply has no entry.
 */
export function findCollisions(plans: LeafPlan[], filesAt: Record<string, string[]>): Collision[] {
  const claims: Record<string, Array<{ path: string; existing: boolean }>> = {};
  // Tab as the separator: legal in neither a SharePoint folder path nor a file name, so it cannot
  // appear inside either half and split the key in the wrong place.
  const key = (folder: string, name: string): string => `${folder}\t${name.toLowerCase()}`;

  for (const plan of plans ?? []) {
    if (!plan.to) continue;
    for (const file of plan.leaf.files) {
      const k = key(plan.to, file);
      if (!claims[k]) claims[k] = [];
      claims[k].push({ path: `${plan.leaf.path}/${file}`, existing: false });
    }
  }
  // Existing occupants, added only where something is actually arriving.
  for (const k of Object.keys(claims)) {
    const cut = k.indexOf("\t");
    const folder = k.slice(0, cut);
    const lower = k.slice(cut + 1);
    for (const there of filesAt[folder] ?? []) {
      if (there.toLowerCase() !== lower) continue;
      // A file whose own leaf is being moved is already a claimant; do not count it twice.
      const alreadyClaimed = claims[k].filter((c) => c.path === `${folder}/${there}`).length > 0;
      if (!alreadyClaimed) claims[k].push({ path: `${folder}/${there}`, existing: true });
    }
  }

  const out: Collision[] = [];
  for (const k of Object.keys(claims)) {
    if (claims[k].length < 2) continue;
    const cut = k.indexOf("\t");
    const folder = k.slice(0, cut);
    const sorted = claims[k].slice().sort((a, b) => a.path.localeCompare(b.path));
    out.push({
      folder,
      name: sorted[0].path.slice(sorted[0].path.lastIndexOf("/") + 1),
      claimants: sorted,
    });
  }
  return out.sort((a, b) => `${a.folder}/${a.name}`.localeCompare(`${b.folder}/${b.name}`));
}

/**
 * Suggest a unique name for a file losing a collision, carrying the value that made it distinct.
 *
 * `a (testig).pdf`, never `a (2).pdf`. The removed tier's value is precisely WHY these two files
 * were different; a numeric suffix throws that away and leaves two unrelated documents looking
 * like versions of one another.
 *
 * The extension is preserved from the original, as the upload form's rename does — one renaming
 * behaviour in the product, not two.
 */
/** Characters SharePoint rejects in a file name — the same set the upload form's rename strips. */
const ILLEGAL_FILE_CHARS = /[\\/:*?"<>|#%]/;

/**
 * Check a name an administrator typed to settle a collision.
 *
 * Returns an admin-facing message, or `undefined` when the name is usable. Validated rather than
 * silently corrected: a name quietly rewritten under someone's fingers is worse than one refused,
 * because they carry on believing they chose it.
 *
 * The extension rule is the one that earns its place. Losing `.pdf` while renaming leaves a file
 * that opens as nothing, and the mistake is invisible in a form whose entire job is renaming — one
 * keystroke in the wrong place does it. So the extension must survive, exactly as the upload form's
 * rename preserves it.
 */
export function validateRename(originalName: string, proposed: string): string | undefined {
  const name = (proposed ?? "").trim();
  if (!name) return "cannot be blank";
  if (ILLEGAL_FILE_CHARS.test(name)) return 'cannot contain \\ / : * ? " < > | # %';
  const dot = (originalName ?? "").lastIndexOf(".");
  if (dot > 0) {
    const ext = originalName.slice(dot).toLowerCase();
    if (name.toLowerCase().slice(-ext.length) !== ext) return `must still end in ${ext}`;
    if (name.length === ext.length) return "needs a name before the extension";
  }
  return undefined;
}

export function suggestRename(fileName: string, distinguisher: string): string {
  const dot = fileName.lastIndexOf(".");
  const stem = dot > 0 ? fileName.slice(0, dot) : fileName;
  const ext = dot > 0 ? fileName.slice(dot) : "";
  const tag = sanitizeFolderSegment(distinguisher);
  if (!tag) return fileName;
  return `${stem} (${tag})${ext}`;
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
