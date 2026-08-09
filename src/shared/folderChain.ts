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
