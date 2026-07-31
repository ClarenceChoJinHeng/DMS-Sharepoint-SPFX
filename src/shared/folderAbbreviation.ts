/**
 * Per-term folder abbreviations, from the `DMS Term Abbreviation` list.
 *
 * Folder names used to come from term labels, which the term store guarantees are
 * unique among siblings — you cannot create two children with the same name under
 * one parent. Abbreviations are hand-maintained and carry no such guarantee, so
 * `findCollisions` restores it: two siblings resolving to one folder would mean a
 * single ACL over two units' documents, breaking the isolation the whole
 * permission model rests on.
 *
 * See docs/superpowers/specs/2026-07-30-folder-abbreviation-naming-design.md section 4.
 */
export interface AbbrevRow {
  termGuid: string;
  abbreviation: string;
}

/** A folder target awaiting its name, with enough context to detect a clash. */
export interface AbbrevTarget {
  parentPath: string;
  termGuid: string;
  abbreviation: string;
  label: string;
}

export interface AbbrevCollision {
  parentPath: string;
  abbreviation: string;
  labels: string[];
}

/**
 * Name of the SharePoint list holding the mappings.
 *
 * The list READ lives in FolderManager.tsx, not here: importing `@microsoft/sp-http`
 * into a shared module makes its Jest suite fail to run ("Cannot find module
 * '@msinternal/ecs-flight'"), which is why dmsFolderMap.ts is untested while this
 * module, formModel and pathEncoding are. Keeping this file free of SPFx imports is
 * what keeps the collision logic covered by tests.
 */
export const ABBREV_LIST = "DMS Term Abbreviation";

/** Term GUIDs are compared lowercased, matching DMS Group Map and DMS Folder Map. */
export function buildAbbrevIndex(rows: readonly AbbrevRow[]): Map<string, string> {
  const ix = new Map<string, string>();
  rows.forEach((r) => {
    const guid = (r.termGuid ?? "").trim().toLowerCase();
    const abbrev = (r.abbreviation ?? "").trim();
    if (guid.length === 0 || abbrev.length === 0) return;
    ix.set(guid, abbrev);
  });
  return ix;
}

export function lookupAbbrev(
  index: Map<string, string>,
  termGuid: string,
): string | undefined {
  return index.get((termGuid ?? "").trim().toLowerCase());
}

/**
 * Siblings sharing an abbreviation, compared case-insensitively because SharePoint
 * folder names are not unique by case — `CORS` and `cors` collide in one parent.
 */
export function findCollisions(
  targets: readonly AbbrevTarget[],
): AbbrevCollision[] {
  const groups = new Map<string, AbbrevTarget[]>();
  targets.forEach((t) => {
    const key = `${t.parentPath} ${t.abbreviation.toLowerCase()}`;
    const bucket = groups.get(key);
    if (bucket) bucket.push(t);
    else groups.set(key, [t]);
  });
  const out: AbbrevCollision[] = [];
  groups.forEach((bucket) => {
    if (bucket.length < 2) return;
    out.push({
      parentPath: bucket[0].parentPath,
      abbreviation: bucket[0].abbreviation,
      labels: bucket.map((b) => b.label),
    });
  });
  return out;
}
