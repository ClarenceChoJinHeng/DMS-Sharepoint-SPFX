// Pure, SPFx-free helpers for the Group Map Builder. No @microsoft/* imports —
// keep this unit-testable in plain Jest (same pattern as formModel).

/**
 * DEL and HC are library-scoped like the rest: DEL grants delete on Documents, HC
 * grants Contribute on HC Approval and Read on HC Library.
 *
 * HC is in the type but nothing in Phase 1 assigns it — Highly Confidential left
 * scope on 2026-08-01 and the term is deleted from the term store, so the level
 * cannot even be selected. It is carried here so `feat/hc-libraries` merges later
 * without touching this model again, and is deliberately absent from the role
 * picker: an HC row created today is filtered out of every library and would
 * silently do nothing.
 *
 * See docs/superpowers/specs/2026-08-03-role-personas-and-department-fanout-design.md
 * and 2026-07-16-highly-confidential-securing-design.md §4.2.
 */
export type GroupMapRole =
  | "MEMBER"
  | "UPL"
  | "APR"
  | "DEL"
  | "HC"
  | "GLOBAL";

/**
 * The roles an administrator may pick in the builder. Excludes HC (Phase 2, see
 * above). Keeping the list here rather than in the component means the type and
 * the picker cannot drift apart.
 */
export const SELECTABLE_ROLES: GroupMapRole[] = ["MEMBER", "UPL", "APR", "DEL", "GLOBAL"];

/**
 * The client's personas, as membership combinations of atomic per-unit groups.
 * No bundle group is ever created: a persona is a SET OF ROWS, not a group,
 * because a ninth group name would be a fourth list to keep joined through a
 * term rename.
 *
 * Head of Department and Head of Unit hold IDENTICAL group sets and differ only
 * in which tier the rows are written at — department vs unit. That is why scope
 * is a separate choice in the UI rather than eight more personas.
 *
 * PIC #2 (`_HC` only) is absent: it is the one persona that cannot exist in
 * Phase 1, since Highly Confidential has no term to select. Listing it with no
 * way to provision it would read as a bug.
 */
export interface Persona {
  key: string;
  label: string;
  /** Atomic roles this persona holds. Order is display order. */
  roles: GroupMapRole[];
  /** Shown under the label, so an admin can tell #1 from #4 without the spec open. */
  summary: string;
}

export const PERSONAS: Persona[] = [
  { key: "headof1",  label: "Head of — approve + upload",          roles: ["MEMBER", "UPL", "APR"],        summary: "Reads approved documents, uploads, and approves." },
  { key: "headof2",  label: "Head of — approve only",              roles: ["MEMBER", "APR"],               summary: "Approves but cannot upload. Needs the DMS Approve permission level." },
  { key: "headof3",  label: "Head of — approve + delete",          roles: ["MEMBER", "APR", "DEL"],        summary: "Approves and deletes approved documents; cannot upload." },
  { key: "headof4",  label: "Head of — approve + upload + delete", roles: ["MEMBER", "UPL", "APR", "DEL"], summary: "Uploads, approves, and deletes approved documents." },
  { key: "pic1",     label: "PIC — upload + read",                 roles: ["MEMBER", "UPL"],               summary: "Uploads, and reads approved documents." },
  { key: "pic3",     label: "PIC — upload only",                   roles: ["UPL"],                         summary: "Uploads and sees only their own pending items. No access to Documents." },
  { key: "employee", label: "SDG Employee",                        roles: ["MEMBER"],                      summary: "Reads approved documents. Cannot upload." },
];

/** Look up a persona by key. Returns undefined rather than throwing — the caller renders nothing. */
export function personaByKey(key: string): Persona | undefined {
  for (const p of PERSONAS) if (p.key === key) return p;
  return undefined;
}

/**
 * The site-entry SP group. Grants Read on the web + Home + Documents library so any DMS user
 * can actually open the site (folder-group Limited Access alone cannot — see the
 * site-entry-access-layer spec). NOT a Group Map row: it is a pure SharePoint permission, not
 * an upload-routing role. Adding a user to any DMS group via the Members modal also adds them
 * here so nobody is left locked out of the site.
 */
export const SITE_ENTRY_GROUP_NAME = "DMS_SITE_MEMBERS";

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
  if (n.endsWith("_DEL")) return "DEL";
  if (n.endsWith("_HC")) return "HC";
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

/**
 * Suggest a site-group title from the builder's current selections, e.g.
 * "DMS_Group Head Office_Group Finance_UPL". Purely a SUGGESTION — the admin
 * edits it freely before create (no abbreviation algorithm by design; see the
 * native-sharepoint-groups spec). GLOBAL is a fixed name; MEMBER has no suffix.
 */
export function suggestGroupName(
  segmentLabel: string,
  tierLabels: string[],
  role: GroupMapRole | "",
): string {
  if (role === "GLOBAL") return "DMS_GLOBAL";
  const parts = ["DMS", norm(segmentLabel), ...tierLabels.map(norm)].filter(Boolean);
  // MEMBER is the base group and carries NO suffix — that is the convention
  // roleFromGroupName reads back, so adding one here would make every base group
  // parse as an unknown role. An unpicked role ("") is the same: the admin is
  // still choosing, and "DMS_GHO_" is not a name to hand them.
  const suffix = role === "MEMBER" || role === "" ? "" : `_${role}`;
  return parts.join("_") + suffix;
}
