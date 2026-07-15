// Pure, SPFx-free model helpers for the multi-segment upload form.
// No imports from @microsoft/* here — keep this unit-testable in plain Jest.

/** One cascade level within a mode: a labelled dropdown that writes one column. */
export interface Level {
  label: string;   // shown to the user, e.g. "Estate/Mill"
  column: string;  // logical column key, mapped to a real field in Form.tsx LEVEL_COLUMNS
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

/** One row of the DMS Group Map list. */
export interface GroupMapRow {
  groupId: string;      // Entra Object ID — the match key
  groupName: string;
  segment: string;      // term set GUID
  unitTermGuid: string;
  role: string;         // "UPL" | "APR" | ""
}

/** A user's resolved path from a matched group. */
export interface UserPath {
  segment: string;      // term set GUID (identifies the mode)
  unitTermGuid: string;
  role: string;
}

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
    .map((e) => ({ label: e.label, column: e.column }));
}

/** Return the UserPath for every group the user belongs to, in row order. */
export function matchUserPaths(rows: GroupMapRow[], userGroupIds: string[]): UserPath[] {
  const wanted = new Set(userGroupIds.map((id) => (id ?? "").trim().toLowerCase()));
  return rows
    .filter((r) => wanted.has((r.groupId ?? "").trim().toLowerCase()))
    .map((r) => ({ segment: r.segment, unitTermGuid: r.unitTermGuid, role: r.role }));
}
