// The folder-structure chain: which tiers of a mode's `Levels` carry an ACL, and
// which are ensure-created on demand beneath the Unit folder.
//
// Pure and SPFx-free — no imports from @microsoft/*, so the ordering rules are
// unit-testable in plain Jest without a tenant. Same approach as approvalQueue.ts.
//
// Spec: docs/superpowers/specs/2026-08-06-configurable-folder-structure-chain-design.md
//
// Both upload web parts held near-duplicate copies of this logic, each hardcoding
// exactly `Year` then `Document Type`. That duplication is the reason this module
// exists: two copies of "what shape is the path" is how a site ends up with two
// trees per unit, both populated and neither complete.
import { Level, sanitizeFolderSegment } from "./formModel";

/**
 * A tier is permissioned unless it says otherwise. ABSENT MEANS TRUE — see the
 * `permissioned` comment on `Level`. Exported because reconciliation needs the
 * same predicate, and a second copy of this default would be a silent-ACL bug.
 */
export function isPermissioned(level: Level): boolean {
  return level.permissioned !== false;
}

/** A chain split into the part reconciliation builds and the part upload builds. */
export interface ChainSplit {
  /** Tiers with their own ACL; the deepest of them is the Unit folder. */
  permissioned: Level[];
  /** Tiers ensure-created at upload time, inheriting the Unit's ACL. In path order. */
  onDemand: Level[];
}

/** Why a chain was rejected. `undefined` from `validateChain` means it is usable. */
export interface ChainError {
  code: "empty" | "prefix-not-contiguous" | "duplicate-column";
  /** Admin-facing. Names the offending tier so a config row can be fixed without a developer. */
  message: string;
}

/**
 * Reject a chain rather than repair it.
 *
 * The contiguity rule is the load-bearing one. Every position the client asked for
 * — before Year, between Year and Document Type, after Document Type — is BELOW
 * Unit, so every new tier is non-permissioned and the permissioned tiers always
 * form a prefix. A permissioned entry appearing after a non-permissioned one would
 * mean an ACL'd folder nested inside an inheriting one: reconciliation never walks
 * there so it would never be built, and the client explicitly declined breaking
 * inheritance below Unit (2026-08-06).
 *
 * Sorting the chain into shape instead would move a tier the author meant to be
 * permissioned down into the inheriting suffix — a silent permissions widening,
 * which is precisely what the `absent means true` default exists to avoid. So a
 * malformed chain is reported and upload is blocked.
 */
export function validateChain(levels: Level[]): ChainError | undefined {
  if (!levels || levels.length === 0) {
    return { code: "empty", message: "The folder structure for this segment is empty." };
  }

  let firstOnDemand: string | undefined;
  for (const lvl of levels) {
    if (!isPermissioned(lvl)) {
      if (firstOnDemand === undefined) firstOnDemand = lvl.label;
      continue;
    }
    if (firstOnDemand !== undefined) {
      return {
        code: "prefix-not-contiguous",
        message:
          `"${lvl.label}" is a permissioned tier but sits below "${firstOnDemand}", which is not. ` +
          `Every permissioned tier must come first in the chain.`,
      };
    }
  }

  // Two tiers writing the same column overwrite each other in the item's metadata,
  // and the stored value ends up describing whichever tier wrote last while both
  // folder segments remain in the path. Cheap to catch here; near-impossible to
  // spot in a live library.
  const seen: Record<string, true> = {};
  for (const lvl of levels) {
    const key = (lvl.labelCol ?? lvl.column ?? "").trim().toLowerCase();
    if (!key) continue;
    if (seen[key]) {
      return {
        code: "duplicate-column",
        message:
          `Two tiers both write the column "${lvl.labelCol ?? lvl.column}". ` +
          `Each tier needs its own column.`,
      };
    }
    seen[key] = true;
  }

  return undefined;
}

/**
 * Split a chain at the permissioned boundary. Assumes `validateChain` passed —
 * callers must check first, or a malformed chain yields a truncated path.
 */
export function splitChain(levels: Level[]): ChainSplit {
  const permissioned: Level[] = [];
  const onDemand: Level[] = [];
  for (const lvl of levels ?? []) {
    if (isPermissioned(lvl)) permissioned.push(lvl);
    else onDemand.push(lvl);
  }
  return { permissioned, onDemand };
}

/**
 * The compatibility bridge.
 *
 * A chain with no non-permissioned entries is every mode row that exists today
 * ([Department, Unit]). Those sites keep the hardcoded `Year -> Document Type`
 * behaviour until their row is migrated, so a site can move one mode at a time
 * instead of all at once, and a half-configured site keeps routing files to a
 * complete path rather than a truncated one.
 *
 * Remove this once every site's rows carry an explicit below-Unit chain.
 */
export function needsLegacyBelowUnit(levels: Level[]): boolean {
  return splitChain(levels).onDemand.length === 0;
}

/**
 * The below-Unit tiers a site behaves as if it had configured.
 *
 * Every live mode row is [Department, Unit] with nothing below it, so without this
 * the upload form would need two code paths — one walking a configured chain, one
 * running the old hardcoded Year -> Document Type pair. Two paths is exactly the
 * duplication that let the form and reconciliation drift into building different
 * shapes. Synthesising the legacy pair as a chain gives one walk for both.
 *
 * The synthetic tiers carry the same `column` keys and internal names the form has
 * always written (`Year`, `Document_x0020_Type`), so nothing about an unmigrated
 * site's metadata or folder path changes.
 *
 * Delete this once every site's mode rows carry an explicit chain; the caller then
 * just uses `splitChain(levels).onDemand`.
 */
export function effectiveOnDemandTiers(
  levels: Level[],
  legacyYearTermSet: string,
  legacyDocTypeTermSet: string,
): Level[] {
  const { onDemand } = splitChain(levels ?? []);
  if (onDemand.length > 0) return onDemand;
  return builtInOnDemandTiers(legacyYearTermSet, legacyDocTypeTermSet);
}

/**
 * The two below-Unit tiers every segment runs on before anyone edits its structure.
 *
 * ONE definition, because their shape is load-bearing in a way that is invisible when
 * wrong: neither carries a `tidCol`, and `Document Type`'s real internal name is the
 * ENCODED `Document_x0020_Type`, not the `DocumentType` a name-sanitizer would derive.
 */
export function builtInOnDemandTiers(
  legacyYearTermSet: string,
  legacyDocTypeTermSet: string,
): Level[] {
  return [
    { label: "Year", column: "Year", labelCol: "Year", termSet: legacyYearTermSet, permissioned: false },
    {
      label: "Document Type", column: "DocumentType", labelCol: "Document_x0020_Type",
      termSet: legacyDocTypeTermSet, permissioned: false,
    },
  ];
}

/**
 * The built-in tier an admin has just re-created BY NAME, or undefined for a genuinely new one.
 *
 * **Removing `Year` or `Document Type` and adding it back must not produce an ordinary level**,
 * and before 2026-08-20 it did. The add form derives `tidCol: "<Column>Tid"` for every tier it
 * creates, but these two are MANAGED METADATA columns that already exist:
 *   - A `tidCol` is what tells every writer the column is plain text. With one attached, the
 *     migrator's backfill and the upload form both write the bare label `2024` into the taxonomy
 *     `Year` column, which SharePoint rejects with *"The data returned from the tagging UI was not
 *     formatted correctly"*. Seen live on 18 of 18 documents.
 *   - `Document Type` is worse: the derived column `DocumentType` does not exist, and ONE unknown
 *     field name fails the WHOLE `validateUpdateListItem` call (gotcha #4) — so a single re-added
 *     tier silently costs every other tier's metadata too.
 *
 * The admin's own term set still wins when they supplied one: that drives which options the
 * dropdown offers, and is theirs to choose. What they may NOT choose is the column shape, because
 * the column already exists and its type is not a matter of opinion.
 */
export function builtInTierFor(
  label: string,
  legacyYearTermSet: string,
  legacyDocTypeTermSet: string,
): Level | undefined {
  const key = (label ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  if (!key) return undefined;
  for (const tier of builtInOnDemandTiers(legacyYearTermSet, legacyDocTypeTermSet)) {
    const byLabel = tier.label.toLowerCase();
    // `documenttype` as well as `document type` — the sanitized column name is what an admin
    // reading the old chain in a list view would most likely retype.
    const byColumn = (tier.column ?? "").toLowerCase();
    if (key === byLabel || key === byColumn) return { ...tier };
  }
  return undefined;
}

/** Whether a below-Unit tier applies to the path currently being built. */
export type TierDecision = "keep" | "skip" | "unresolved";

/**
 * Decide whether a below-Unit tier applies, given how many options it actually has.
 *
 * Exists because the client confirmed on 2026-08-10 that **not every Unit has SubUnits**.
 * A tier that cascades from the term above therefore applies only where that term has
 * children: `CORU` has subunits and must use one; `GTAX` has none, and its files sit
 * directly under the unit.
 *
 * This keeps the original "every tier is required" rule intact where it mattered. That
 * rule existed so files could not scatter across different depths *within one unit* — and
 * under this decision each unit stays internally consistent, because the answer comes from
 * the unit's own terms rather than from an uploader's judgement.
 *
 * `optionCount === undefined` means NOT KNOWN — not loaded yet, or the term-store call
 * failed. That returns `unresolved`, and callers MUST block the upload rather than treat
 * it as "no subunits". Conflating the two would file a document one tier too shallow, into
 * a folder that exists and looks correct, on nothing worse than a transient network error.
 *
 * A tier with its own `termSet` is always kept: an empty flat set is a configuration fault
 * to surface, not a signal that the tier does not apply here.
 */
export function decideTier(level: Level, optionCount: number | undefined): TierDecision {
  if ((level.termSet ?? "").trim()) return "keep";
  if (optionCount === undefined) return "unresolved";
  return optionCount === 0 ? "skip" : "keep";
}

/** What reconciliation needs to know before pre-creating a grid of below-Unit folders. */
export interface GridPlan {
  /** Total folders across every tier — what the progress estimate counts. */
  total: number;
  /**
   * The deepest path of "last label at each tier". The grid is built in order, so
   * if this one folder exists the whole grid exists — one probe instead of a
   * round-trip per folder per leaf per library.
   */
  lastPath: string[];
  /** The tiers actually buildable, truncated at the first one with no terms. */
  tiers: string[][];
}

/**
 * Size a Year × Document Type × … grid of any depth.
 *
 * `total` is the sum of the prefix products, not the product: 3 years × 20 doc
 * types is 3 year folders PLUS 60 document-type folders = 63, because every tier
 * above the deepest is itself a folder. Getting this wrong does not break the
 * build — it makes the progress estimate lie, which is how a long run comes to
 * look like a hang.
 *
 * A tier with no terms truncates everything below it. There is nothing to nest
 * inside a folder that cannot be created, and the tiers above it are still worth
 * building — the upload form ensure-creates the rest on demand.
 */
export function gridPlan(tiers: string[][]): GridPlan {
  const usable: string[][] = [];
  for (const tier of tiers ?? []) {
    const names = (tier ?? []).filter(Boolean);
    if (names.length === 0) break;
    usable.push(names);
  }

  let total = 0;
  let running = 1;
  for (const tier of usable) {
    running *= tier.length;
    total += running;
  }

  return { total, tiers: usable, lastPath: usable.map((t) => t[t.length - 1]) };
}

/** One tier's chosen term, as the UI holds it. */
export interface TierSelection {
  /** Term GUID. Written to the tier's `tidCol`. */
  id: string;
  /** Term label. Becomes the folder segment name AND the `labelCol` value. */
  label: string;
}

/** A resolved below-Unit path, or the reason it could not be resolved. */
export interface SegmentResult {
  /** Folder names in path order, ready to walk with ensureFolder. */
  segments: string[];
  /** Labels of tiers with no selection. Non-empty means do not upload. */
  missing: string[];
}

/**
 * Turn the on-demand suffix plus the user's selections into ordered folder names.
 *
 * Names come from the sanitized TERM LABEL, not the abbreviation list — deliberately.
 * Abbreviations keep permissioned paths short and stable across renames because those
 * paths carry ACLs and appear in the Folder Map; below-Unit folders carry neither, and
 * `2026` / `Invoice` / `Human Resource` is more use to someone browsing than a code.
 *
 * Every tier is required. An optional tier left blank would file documents at
 * inconsistent depths inside one unit, defeating the point of having the tier — so a
 * blank is reported in `missing` rather than skipped, and the caller blocks the upload.
 * Skipping silently would route the file to a shallower path that still exists and
 * still looks correct.
 */
export function buildOnDemandSegments(
  onDemand: Level[],
  selections: Record<string, TierSelection | undefined>,
): SegmentResult {
  const segments: string[] = [];
  const missing: string[] = [];
  for (const lvl of onDemand ?? []) {
    const picked = selections[lvl.column];
    const name = sanitizeFolderSegment(picked?.label ?? "");
    if (!name) {
      missing.push(lvl.label);
      continue;
    }
    segments.push(name);
  }
  return { segments, missing };
}
