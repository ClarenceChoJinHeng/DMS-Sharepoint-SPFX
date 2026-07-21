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

/** A candidate leaf the user holds the UPL (uploader) role for. */
export interface UploaderLeaf {
  termGuid: string;
  segment: string; // term-set GUID of the leaf's mode
}

/** What the user's group memberships resolve to, before term-tree validation. */
export interface Membership {
  memberTerms: Set<string>; // every term (or term-set) GUID the user is a member of, normalised
  uploaderLeaves: UploaderLeaf[]; // the user's UPL-role leaves
  isGlobalUploader: boolean; // user holds a GLOBAL-role group -> upload anywhere
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
 *  - `uploaderLeaves`: the leaves the user holds the UPL role for;
 *  - `isGlobalUploader`: true if the user is in any GLOBAL-role group.
 * Matching is case-insensitive and trimmed on both sides.
 */
export function collectMembership(
  rows: GroupMapRow[],
  userGroupIds: string[],
): Membership {
  const wanted = new Set(userGroupIds.map(normGuid));
  const memberTerms = new Set<string>();
  const uploaderLeaves: UploaderLeaf[] = [];
  let isGlobalUploader = false;
  for (const r of rows) {
    if (!wanted.has(normGuid(r.groupId))) continue;
    const role = (r.role ?? "").trim().toUpperCase();
    if (role === "GLOBAL") {
      isGlobalUploader = true;
      continue; // GLOBAL rows carry no term
    }
    if (r.termGuid) memberTerms.add(normGuid(r.termGuid));
    if (role === "UPL") {
      uploaderLeaves.push({ termGuid: r.termGuid, segment: r.segment });
    }
  }
  return { memberTerms, uploaderLeaves, isGlobalUploader };
}

/**
 * Hard-check a candidate path: the user must be a member of every term in the
 * chain [top … leaf], and — for Business Segment modes — of the segment
 * (term set) itself. Returns true only when the whole chain is authorised.
 */
export function isChainAuthorized(
  chainTermGuids: string[],
  termSetGuid: string,
  memberTerms: Set<string>,
  requireSegmentMembership: boolean,
): boolean {
  if (chainTermGuids.length === 0) return false;
  if (requireSegmentMembership && !memberTerms.has(normGuid(termSetGuid))) {
    return false;
  }
  return chainTermGuids.every((g) => memberTerms.has(normGuid(g)));
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
