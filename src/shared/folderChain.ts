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
  code:
    | "empty"
    | "prefix-not-contiguous"
    | "duplicate-column"
    | "per-unit-not-contiguous";
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

  /* ⚠ PER-UNIT TIERS MUST BE CONTIGUOUS TOO — the same rule as the permissioned prefix, one level
     down, and its absence cost a whole segment's migration on 2026-09-07.

     A below-Unit tier with NO `termSet` is a PER-UNIT tier: its options are the children of the term
     ABOVE it, cascading down the segment's own term tree from the unit's term. That only means
     anything while every tier above it in the below-Unit run is also per-unit. Put one beneath a
     SHARED-LIST tier and the lookup asks the segment's term set for the children of a term that
     lives in a different set entirely — it 404s, the option list comes back `undefined`, and the
     migrator reports every unit as "some of its folder values could not be read from the term
     store" while naming nothing.

     That is exactly what Buah did: `Clarence Kiwi` (no termSet) sat fourth, below `State`, `Year`
     and `Document Type`, all of which have one. Seven units x six libraries, all unreadable, and the
     Folder levels screen had saved it without a word.

     ⚠ REFUSED RATHER THAN REORDERED. Moving the tier would change where documents are filed, and an
     admin who put it fourth may have meant a shared list and forgotten the ID — the two repairs are
     opposite, so this names the problem and lets them choose.

     ⚠ THE PERMISSIONED TIERS ARE SKIPPED ENTIRELY. They have no `termSet` either — they draw from
     the segment set by definition — so counting them here would reject every valid chain. */
  let firstShared: string | undefined;
  for (const lvl of levels) {
    if (isPermissioned(lvl)) continue;
    const hasSet = (lvl.termSet ?? "").trim().length > 0;
    if (hasSet) {
      if (firstShared === undefined) firstShared = lvl.label;
      continue;
    }
    if (firstShared !== undefined) {
      return {
        code: "per-unit-not-contiguous",
        message:
          `"${lvl.label}" takes its values from the terms under each unit, but it sits below ` +
          `"${firstShared}", which uses its own term set. A per-unit level only works directly ` +
          `under Unit, or under other per-unit levels — move it up, or give it a term set of its own.`,
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

/**
 * Is this below-Unit tier one of the two the client has fixed — Year or Document Type?
 *
 * Client, 2026-09-06 and again 2026-09-07: *"an update that we enforce Year and DocType to not move
 * or be deleted"*, then *"We just need to cater for folder structure that will be infront of year
 * and in betwen and after"*. So the pair is immovable and undeletable, and every OTHER tier may be
 * added, moved or removed around them — before, between, or after.
 *
 * ⚠ MATCHED THE SAME WAY `builtInTierFor` MATCHES, and deliberately so: by LABEL or by sanitized
 * COLUMN, case-insensitively, with runs of whitespace folded. Two different answers to "is this the
 * built-in Year tier?" is how one screen would lock it and another let it be deleted.
 *
 * ⚠ IT DOES NOT LOOK AT `termSet`. A tier removed and re-added through the form comes back with the
 * admin's own term set if they typed one, and it is STILL the fixed Year tier — the identity is the
 * name, which is what the column behind it is keyed on.
 *
 * This is a UI rule, not a validity rule: `validateChain` does NOT reject a chain without them.
 * Segments predating the decision exist, and refusing to load their chain would take their uploads
 * down to enforce a preference.
 */
export function isFixedBelowUnitTier(level: Level): boolean {
  const key = (level?.label ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  if (!key) return false;
  const col = (level?.column ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  for (const tier of builtInOnDemandTiers("", "")) {
    if (key === tier.label.toLowerCase()) return true;
    if (col && col === (tier.column ?? "").toLowerCase()) return true;
  }
  return false;
}

/**
 * How many tiers at the START of the below-Unit run are PER-UNIT (no term set of their own).
 *
 * Takes the below-Unit run — `splitChain().onDemand` — not a whole chain. The permissioned tiers
 * have no `termSet` either, because they draw from the segment's own set by definition, so counting
 * them would report every chain as leading with per-unit tiers.
 */
export function leadingPerUnitCount(onDemand: Level[]): number {
  let n = 0;
  for (const lvl of onDemand) {
    if ((lvl.termSet ?? "").trim().length > 0) break;
    n++;
  }
  return n;
}

/**
 * Which insertion slots a new below-Unit tier may take, given what is already there and which kind
 * of tier it is. A slot `p` means "at index p of the below-Unit run".
 *
 * ⚠ THIS IS `per-unit-not-contiguous` SAID BEFORE THE FACT INSTEAD OF AFTER IT. That rule already
 * refuses a per-unit tier sitting below a shared-list one — it is what caught Buah's
 * `Clarence Kiwi` — but it refuses at SAVE and UPLOAD time, so the Position dropdown went on
 * offering slots that produce a chain the same code then rejects. Three options, two of them
 * traps, with the WORST of them (last) pre-selected.
 *
 * The rule removes the choice rather than warning about it afterwards, which is the same reasoning
 * that took the per-unit/shared toggle away from `Year` and `Document Type`: an admin cannot act on
 * a refusal they meet two screens later, and a chain saved in that shape reads as working right up
 * until a migration reports every unit as unreadable.
 *
 * ⚠ IT NEVER RETURNS AN EMPTY LIST. A per-unit tier can always take slot 0, and a shared-list tier
 * can always go at the end, so the dropdown can never render with nothing in it — which would be a
 * dead end rather than a guard.
 */
export function allowedTierPositions(onDemand: Level[], fromUnit: boolean): number[] {
  const n = leadingPerUnitCount(onDemand);
  const out: number[] = [];
  if (fromUnit) {
    // Anywhere inside the per-unit run, or immediately after it — every one of those keeps the run
    // contiguous with Unit. Slot `n` is "last of the per-unit tiers", not "after the shared ones".
    for (let p = 0; p <= n; p++) out.push(p);
  } else {
    // At or after the end of the per-unit run. Going above one would put a shared list between Unit
    // and a tier that cascades from it, which is the failure this exists to prevent.
    for (let p = n; p <= onDemand.length; p++) out.push(p);
  }
  return out;
}

/**
 * Snap a chosen slot into the ones this kind of tier may actually take.
 *
 * ⚠ HIDING THE INVALID OPTIONS IS NOT ENOUGH ON ITS OWN, and leaving this out would have shipped a
 * worse bug than the one it fixes. A `<select>` whose `value` matches no option renders showing the
 * FIRST one while the state behind it keeps the old number — so the screen would say "Before Year"
 * and Add would still insert after Document Type. Three routes can leave a stale slot behind: the
 * form's own default, and each of the two Folder-setup radios.
 *
 * The allowed slots are always a CONTIGUOUS range, so clamping into it is exact rather than a
 * guess, and it keeps the admin's intent — "as deep as you are allowed" — instead of resetting to
 * the top.
 */
export function clampTierPosition(onDemand: Level[], fromUnit: boolean, position: number): number {
  const ok = allowedTierPositions(onDemand, fromUnit);
  const first = ok[0];
  const last = ok[ok.length - 1];
  if (!Number.isFinite(position)) return first;
  return Math.min(Math.max(Math.floor(position), first), last);
}
