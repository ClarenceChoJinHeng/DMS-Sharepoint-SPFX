// Pure, SPFx-free helpers for the Group Map Builder. No @microsoft/* imports —
// keep this unit-testable in plain Jest (same pattern as formModel / shareGuard).

export type GroupMapRole = "MEMBER" | "UPL" | "APR" | "GLOBAL";

/** The exact field set POSTed to the DMS Group Map list. */
export interface GroupMapWriteRow {
  GroupId: string;
  GroupName: string;
  Segment: string;       // "" for GLOBAL
  UnitTermGuid: string;  // "" for GLOBAL; equals Segment for a segment-tier row
  Role: GroupMapRole;
}

/** In-progress selections from the builder UI. */
export interface GroupMapDraft {
  groupId: string;
  groupName: string;
  role: GroupMapRole;
  segmentGuid?: string;   // required unless GLOBAL
  tierGuid?: string;      // required unless GLOBAL; may equal segmentGuid
}

const norm = (s: string): string => (s ?? "").trim();

/**
 * Build the row to POST. Trims every field (kills the trailing-space bug that
 * silently dropped DMS_GHO_LRC). A GLOBAL row carries no term, so Segment and
 * UnitTermGuid are forced empty regardless of any stale draft selections.
 */
export function buildGroupMapRow(draft: GroupMapDraft): GroupMapWriteRow {
  const isGlobal = draft.role === "GLOBAL";
  return {
    GroupId: norm(draft.groupId),
    GroupName: norm(draft.groupName),
    Segment: isGlobal ? "" : norm(draft.segmentGuid ?? ""),
    UnitTermGuid: isGlobal ? "" : norm(draft.tierGuid ?? ""),
    Role: draft.role,
  };
}

/** True if an equivalent row exists: GroupId + UnitTermGuid + Role, normalised. */
export function isDuplicateRow(
  existing: GroupMapWriteRow[],
  candidate: GroupMapWriteRow,
): boolean {
  const key = (r: GroupMapWriteRow): string =>
    `${norm(r.GroupId).toLowerCase()}|${norm(r.UnitTermGuid).toLowerCase()}|${norm(r.Role).toUpperCase()}`;
  const k = key(candidate);
  return existing.some((r) => key(r) === k);
}

/**
 * Derive the intended role from a group name's suffix (CTO naming convention:
 * DMS_<seg>_<dept>_<unit>[_UPL|_APR]). Case-insensitive, trims. GLOBAL is a
 * privileged bypass the admin picks by hand — it is never auto-derived, so any
 * non-suffixed name (including a base viewer group) defaults to MEMBER.
 */
export function roleFromGroupName(name: string): GroupMapRole {
  const n = norm(name).toUpperCase();
  if (n.endsWith("_APR")) return "APR";
  if (n.endsWith("_UPL")) return "UPL";
  return "MEMBER";
}

/** Field-level validation for enabling the Add button. Returns [] when valid. */
export function validateDraft(draft: GroupMapDraft): string[] {
  const errors: string[] = [];
  if (!norm(draft.groupId)) errors.push("Select a group.");
  if (!draft.role) errors.push("Select a role.");
  if (draft.role !== "GLOBAL") {
    if (!norm(draft.segmentGuid ?? "")) errors.push("Select a segment.");
    if (!norm(draft.tierGuid ?? "")) errors.push("Select a tier.");
  }
  return errors;
}
