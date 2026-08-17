// Planning a bulk group run — which groups a segment needs, derived from its abbreviation rows.
//
// Spec: docs/superpowers/specs/2026-08-18-group-creation-and-bulk-provisioning-design.md §5
//
// Provisioning the client's site needs a group per persona per unit: 132 units, ~703 groups. One at a time
// is days of clicking, and the mistake it invites is worse than the tedium — a wrong persona picked once
// and then repeated by hand sixty times.
//
// THE SOURCE IS THE ABBREVIATION ROWS, and that is the whole reason this can be derived at all: a row with
// a non-blank code is exactly a term that will get a folder, which is exactly a term that needs groups.
// Reconciliation skips a term with no code — no folder, no error — so a group for such a term would grant
// access to nothing.
//
// Pure and SPFx-free: every rule here decides what gets created on a real tenant, and the cost of being
// wrong is hundreds of groups that have to be deleted by hand.
import { AbbrevRowDraft, folderNameFor } from "./abbreviationDraft";
import { namingRoleFor, PERSONAS, suggestGroupName } from "./groupMapModel";

/** A segment, as the planner needs it. */
export interface BulkSegment {
  /** StagingFolder — the first part of every name. */
  code: string;
  termSetGuid: string;
  /** Permissioned tier names, shallowest first. Decides which depth counts as the unit tier. */
  levelNames: string[];
}

/** One group the run would create or map. */
export interface PlannedGroup {
  name: string;
  personaKey: string;
  /** The term its Group Map rows sit on. Equals `termSetGuid` for a segment-scope persona. */
  tierGuid: string;
  scope: "segment" | "department" | "unit";
  /** True when a site group of this name already exists — it is mapped, never re-created. */
  exists: boolean;
}

/** A term the run deliberately passes over, and why. */
export interface SkippedTerm {
  label: string;
  level: string;
  reason: string;
}

export interface BulkPlan {
  groups: PlannedGroup[];
  skipped: SkippedTerm[];
}

/** Case-insensitive, because SharePoint is not consistent about GUID casing between stores. */
function key(guid: string): string {
  return (guid ?? "").trim().toLowerCase();
}

/**
 * The chain of codes from the shallowest tier down to this row, or `undefined` if any link is missing.
 *
 * **A UNIT WHOSE DEPARTMENT HAS NO CODE MUST BE SKIPPED TOO**, and that is the subtle one. Reconciliation
 * builds `GHO/<dept>/<unit>`, so a unit with a code under a department without one has no folder to be
 * granted on — its group would look correct and grant nothing. Checking only the row's own code would
 * create exactly those.
 *
 * Walks `parentGuid` rather than assuming two tiers, so a family with a different depth needs no change
 * here. Bounded by the row count, so a cyclic parent link cannot spin.
 */
function codeChain(
  row: AbbrevRowDraft,
  byGuid: Record<string, AbbrevRowDraft>,
): string[] | undefined {
  const chain: string[] = [];
  let cur: AbbrevRowDraft | undefined = row;
  let guard = 0;
  const limit = Object.keys(byGuid).length + 1;
  while (cur && guard <= limit) {
    const code = folderNameFor(cur.abbreviation);
    if (!code) return undefined;
    chain.unshift(code);
    const parentKey = key(cur.parentGuid);
    if (!parentKey) break;
    cur = byGuid[parentKey];
    guard++;
  }
  return chain;
}

/**
 * Plan a bulk run.
 *
 * `rows` are the segment's abbreviation rows — permissioned tiers only; a below-Unit tier carries no group
 * and must never reach here. `personaKeys` is what the admin ticked. `existingTitles` are the site's
 * current group names, compared case-insensitively.
 *
 * Ordering is deliberate: segment-scope groups first, then the rows in the order they arrive, which is the
 * term store's own order. A preview an admin has to scan should read top-down like the tree it mirrors.
 */
export function planBulkGroups(
  segment: BulkSegment,
  rows: AbbrevRowDraft[],
  personaKeys: string[],
  existingTitles: string[],
): BulkPlan {
  const groups: PlannedGroup[] = [];
  const skipped: SkippedTerm[] = [];
  const have: Record<string, true> = {};
  for (const t of existingTitles ?? []) have[(t ?? "").trim().toLowerCase()] = true;
  const byGuid: Record<string, AbbrevRowDraft> = {};
  for (const r of rows ?? []) byGuid[key(r.termGuid)] = r;

  const personas = PERSONAS.filter((p) => (personaKeys ?? []).indexOf(p.key) !== -1);
  const deepest = (segment.levelNames ?? []).length;
  const planned: Record<string, true> = {};

  const push = (
    name: string,
    personaKey: string,
    tierGuid: string,
    scope: PlannedGroup["scope"],
  ): void => {
    // A name can only be planned ONCE. Two personas sharing a naming role would otherwise plan the same
    // title twice, and the second create would fail as a duplicate mid-run — noise that reads as an error.
    const k = name.toLowerCase();
    if (planned[k]) return;
    planned[k] = true;
    groups.push({ name, personaKey, tierGuid, scope, exists: have[k] === true });
  };

  // ── Segment scope: one group per persona, no tier ─────────────────────────────
  for (const p of personas.filter((x) => x.scope === "segment")) {
    const name = suggestGroupName(segment.code, [], namingRoleFor(p.key));
    if (name) push(name, p.key, segment.termSetGuid, "segment");
  }

  // ── Department and unit scope, from the rows themselves ───────────────────────
  // A row's DEPTH is its position in the chain, so this needs no per-family special case: the deepest
  // tier is the unit tier and everything above it is a department tier.
  for (const r of rows ?? []) {
    const chain = codeChain(r, byGuid);
    if (!chain) {
      skipped.push({
        label: r.label,
        level: r.level,
        reason: folderNameFor(r.abbreviation)
          ? "a tier above it has no folder code, so it has no folder to grant on"
          : "no folder code, so reconciliation creates no folder for it",
      });
      continue;
    }
    const scope: PlannedGroup["scope"] =
      deepest > 0 && chain.length >= deepest ? "unit" : "department";
    for (const p of personas.filter((x) => x.scope === scope)) {
      const name = suggestGroupName(segment.code, chain, namingRoleFor(p.key));
      if (name) push(name, p.key, r.termGuid, scope);
    }
  }

  return { groups, skipped };
}

/** How many of a plan's groups would actually be created. The rest are mapped only. */
export function toCreateCount(plan: BulkPlan): number {
  return (plan?.groups ?? []).filter((g) => !g.exists).length;
}
