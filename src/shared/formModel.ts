// Pure, SPFx-free model helpers for the multi-segment upload form.
// No imports from @microsoft/* here — keep this unit-testable in plain Jest.

/** One cascade level within a mode: a labelled dropdown that writes one column. */
export interface Level {
  label: string;   // shown to the user, e.g. "Estate/Mill"
  column: string;  // logical column key, mapped to a real field in Form.tsx LEVEL_COLUMNS
  // Config-driven override: the REAL Staging internal names for this level's
  // label + term-GUID columns. When present (from DMS Config Levels JSON) these
  // win over the LEVEL_COLUMNS fallback — so a client tenant needs no code change,
  // just the right internal names in config.
  labelCol?: string;
  tidCol?: string;
}

/** A configured upload mode (one term set + its level chain). */
export interface ModeV2 {
  key: string;
  label: string;
  side: "BusinessSegment" | "Project";
  termSetGuid: string;
  stagingFolder: string;
  levels: Level[];
  sortOrder: number;
}

/**
 * One row of the DMS Group Map list. A row maps a security group (by Entra
 * Object ID) to a term at some tier of a mode's chain.
 *  - Segment-tier rows: `termGuid` == the term-set GUID (segment membership).
 *  - Intermediate/leaf rows: `termGuid` == the term's GUID.
 */
export interface GroupMapRow {
  groupId: string; // Entra Object ID — the match key
  groupName: string;
  segment: string; // term set GUID (which mode this row scopes to)
  termGuid: string; // term GUID, or the term-set GUID for a Segment-tier row
  role: string; // "MEMBER" | "UPL" | "APR" | "GLOBAL" | ""
}

/**
 * The two roles that permit uploading. They are not ranked.
 *
 * UPL uploads ordinary documents to Staging; HC uploads Highly Confidential ones
 * to HC Approval. Holding HC does not imply UPL — that is the whole point of the
 * compartment, and it is what lets a PIC upload and read Highly Confidential
 * documents while Confidential ones stay invisible to them.
 */
export type UploadRole = "UPL" | "HC";

/** A candidate leaf the user may upload to, and in which capacity. */
export interface UploaderLeaf {
  termGuid: string;
  segment: string; // term-set GUID of the leaf's mode
  role: UploadRole;
}

/** What the user's group memberships resolve to, before term-tree validation. */
export interface Membership {
  memberTerms: Set<string>; // every term (or term-set) GUID the user is a member of, normalised
  uploaderLeaves: UploaderLeaf[]; // the user's UPL- and HC-role leaves
}

const normGuid = (g: string): string => (g ?? "").trim().toLowerCase();

/** Safely parse a DMS Config `Levels` JSON string into Level[]. Never throws. */
export function parseLevels(json: string): Level[] {
  if (!json || typeof json !== "string") return [];
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return [];
  }
  if (!Array.isArray(raw)) return [];
  return raw
    .filter(
      (e): e is Level =>
        !!e &&
        typeof (e as Level).label === "string" &&
        typeof (e as Level).column === "string",
    )
    .map((e) => {
      const lvl: Level = { label: e.label, column: e.column };
      if (typeof e.labelCol === "string") lvl.labelCol = e.labelCol;
      if (typeof e.tidCol === "string") lvl.tidCol = e.tidCol;
      return lvl;
    });
}

/** Raw DMS Config `mode` row (only the fields the reconciliation provisioner needs). */
export interface RawModeRow {
  TermSetGuid?: string;
  StagingFolder?: string;
  Levels?: string;
  SortOrder?: number;
}

/** A reconciliation target segment: the term set + its top-level Staging folder. */
export interface ReconMode {
  termSetGuid: string;
  stagingFolder: string;
  sortOrder: number;
}

/**
 * Slim the DMS Config `mode` rows down to what the reconciliation provisioner needs
 * (term set + Staging folder), so seeding a segment is data-only — no code change.
 * Trims stray whitespace (a trailing space silently breaks GUID matching), drops rows
 * missing a term set or folder, and — like the upload form — ignores rows with an
 * empty/old-schema `Levels` (so retired department/project rows never provision).
 * Sorted by SortOrder.
 */
export function parseReconModes(rows: RawModeRow[]): ReconMode[] {
  const out: ReconMode[] = [];
  for (const r of rows ?? []) {
    const termSetGuid = (r.TermSetGuid ?? "").trim();
    const stagingFolder = (r.StagingFolder ?? "").trim();
    if (!termSetGuid || !stagingFolder) continue;
    if (parseLevels(r.Levels ?? "").length === 0) continue;
    out.push({
      termSetGuid,
      stagingFolder,
      sortOrder: typeof r.SortOrder === "number" ? r.SortOrder : 0,
    });
  }
  return out.sort((a, b) => a.sortOrder - b.sortOrder);
}

/**
 * Reduce the DMS Group Map to the current user's memberships:
 *  - `memberTerms`: every term/term-set GUID whose group the user is in (any role);
 *  - `uploaderLeaves`: the leaves the user holds the UPL role for.
 * `GLOBAL` is a read-only super-viewer role (reads all of Documents + Staging, enforced
 * by SharePoint library-level Read grants — see the site-entry-access-layer spec). It grants
 * NO upload capability, so GLOBAL rows are ignored here entirely (they carry no term).
 * Upload-anywhere is a separate site-admin concern, checked outside this function.
 * Matching is case-insensitive and trimmed on both sides.
 */
export function collectMembership(
  rows: GroupMapRow[],
  userGroupIds: string[],
): Membership {
  const wanted = new Set(userGroupIds.map(normGuid));
  const memberTerms = new Set<string>();
  const uploaderLeaves: UploaderLeaf[] = [];
  for (const r of rows) {
    if (!wanted.has(normGuid(r.groupId))) continue;
    const role = (r.role ?? "").trim().toUpperCase();
    if (role === "GLOBAL") continue; // read-only role, no term, no upload
    if (r.termGuid) memberTerms.add(normGuid(r.termGuid));
    // HC counts as an upload role in its own right, not as an extra on top of
    // UPL. A user holding only HC must still be offered their unit — omitting
    // them here is what would tell a correctly provisioned PIC that their
    // account "isn't fully provisioned to upload".
    if (role === "UPL" || role === "HC") {
      uploaderLeaves.push({ termGuid: r.termGuid, segment: r.segment, role });
    }
  }
  return { memberTerms, uploaderLeaves };
}

/**
 * The confidentiality levels a user may choose, given the upload roles they hold
 * at the leaf they are uploading to.
 *
 * Hiding the option is the enforcement, not a courtesy: an HC upload has to land
 * in a pre-secured folder, so offering the level to someone without the HC role
 * would produce a failure at submit time instead of a choice they never see.
 *
 * The inverse matters just as much — a UPL-only user must not be offered Highly
 * Confidential, and an HC-only user must not be offered Confidential or
 * Restricted, because those live in a library they cannot reach.
 */
export function filterConfidentiality<T extends { id: string }>(
  options: readonly T[],
  rolesHeld: readonly UploadRole[],
  hcTermGuid: string,
): T[] {
  const hc = normGuid(hcTermGuid);
  const isHcOption = (o: T): boolean => normGuid(o.id) === hc;
  const hasHc = rolesHeld.indexOf("HC") !== -1;
  const hasUpl = rolesHeld.indexOf("UPL") !== -1;
  // An unconfigured hcTermGuid must not silently strip the option for HC holders
  // or leak it to everyone: with nothing to match, no option is the HC one, so
  // UPL sees every level and HC sees none. The form blocks on the missing config
  // separately rather than guessing here.
  if (hasHc && hasUpl) return options.slice();
  if (hasHc) return options.filter(isHcOption);
  if (hasUpl) return options.filter((o) => !isHcOption(o));
  return [];
}

/**
 * Validate a candidate path STRUCTURALLY: the chain must have resolved, and it
 * must terminate at the uploader leaf it was derived from.
 *
 * Authorisation lives in the leaf alone. The chain is not user input — callers
 * build it with `loadTermPath(termSetGuid, leaf.termGuid)`, i.e. it is the term
 * store's own ancestry for a term the user already holds the UPL role on. There
 * is no way to supply a chain you are not entitled to, so re-checking every
 * ancestor against the Group Map only asks that list to restate what the unit
 * row plus the term store already say. See the leaf-only-upload-authorization
 * spec (2026-07-29).
 *
 * The predecessor (`isChainAuthorized`) demanded membership at EVERY tier plus
 * the segment. That belonged to the retired per-tier group model; under
 * one-group-per-unit it rejects correctly provisioned uploaders unless hundreds
 * of derivable MEMBER rows are maintained by hand.
 *
 * The leaf anchor is worth keeping: `loadTermPath` is a network read wrapped in
 * `.catch(() => [])`, and a partial or unrelated chain would otherwise point an
 * upload at a folder the user has no claim on.
 */
export function isLeafChainValid(
  chainTermGuids: string[],
  leafTermGuid: string,
): boolean {
  if (chainTermGuids.length === 0) return false;
  const last = chainTermGuids[chainTermGuids.length - 1];
  return normGuid(last) === normGuid(leafTermGuid);
}

const ILLEGAL_FOLDER_CHARS = /[\\/:*?"<>|#%]/g;

/** Make a term label safe to use as a single folder name. Returns "" if nothing remains. */
export function sanitizeFolderSegment(name: string): string {
  if (!name) return "";
  return name.replace(ILLEGAL_FOLDER_CHARS, "").replace(/\s+/g, " ").trim();
}

export interface LevelSelection {
  column: string; // logical key
  label: string;
  id: string;
  // Optional explicit internal names (from DMS Config Levels JSON). When set,
  // they override the columnMap lookup for this selection.
  labelCol?: string;
  tidCol?: string;
}
export interface SpFormValue {
  FieldName: string;
  FieldValue: string;
}

/** Maps a logical column key to the two real Staging field internal names. */
export interface ColumnPair {
  label: string; // internal name of the text label column
  tid: string;   // internal name of the term-GUID column
}

/**
 * Build validateUpdateListItem field pairs for the chosen levels:
 * one text pair for the label column, one for its term-GUID (Tid) column.
 *
 * Internal names resolve config-first: a selection's own `labelCol`/`tidCol`
 * (supplied by DMS Config Levels JSON) win; otherwise fall back to `columnMap`
 * keyed by the logical `column`. SharePoint does NOT follow a `<label>_Tid`
 * convention (e.g. label `Business_x0020_Segment` pairs with tid
 * `BusinessSegmentTid`), so both names are always explicit. Skips selections
 * with no resolvable label column or a blank label.
 */
export function buildLevelFormValues(
  columnMap: Record<string, ColumnPair>,
  selections: LevelSelection[],
): SpFormValue[] {
  const out: SpFormValue[] = [];
  for (const s of selections) {
    const fallback = columnMap[s.column];
    const labelCol = s.labelCol || fallback?.label;
    const tidCol = s.tidCol || fallback?.tid;
    if (!labelCol || !s.label) continue;
    out.push({ FieldName: labelCol, FieldValue: s.label });
    if (tidCol) out.push({ FieldName: tidCol, FieldValue: s.id });
  }
  return out;
}
