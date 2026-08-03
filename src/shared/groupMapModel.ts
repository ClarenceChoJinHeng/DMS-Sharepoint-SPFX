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
  /** Family heading in the picker, matching the client's own document. */
  family: "Head of Department" | "Head of Unit" | "PIC" | "SDG Employee";
  label: string;
  /** Atomic roles this persona holds. Order is display order. */
  roles: GroupMapRole[];
  /**
   * Which tier the rows belong on.
   *
   * "department" — a non-leaf term. The mapping reaches every unit beneath it via
   * reconciliation's fan-out, which must be switched ON
   * (`recon_departmentFanOut`) or the person gets the department folder and no
   * units at all.
   * "unit" — the leaf term; their own unit only.
   *
   * This is the ONLY difference between the two Head-of families: identical
   * roles, different tier. Getting it backwards hands a unit head an entire
   * department.
   */
  scope: "department" | "unit";
  /** Shown under the label, so an admin can tell #1 from #4 without the spec open. */
  summary: string;
  /**
   * Set when the persona cannot be provisioned on this build. Listed anyway, and
   * greyed out: silently omitting one of the client's own numbered groups reads
   * as an oversight, and an admin would go hunting for it.
   */
  unavailable?: string;
}

/**
 * The client's twelve groups, from their document of 2026-08-03, as membership
 * combinations of atomic per-unit groups. No bundle group is ever created.
 *
 * The client's document repeats "see all unit under the department" under Head of
 * UNIT as well — confirmed 2026-07-31 as a copy-paste error, so Head of Unit is
 * scoped to its own unit here. That one phrase is the difference between a unit
 * head seeing one unit and seeing twelve, so it is recorded rather than assumed
 * silently. See 2026-07-16-highly-confidential-securing-design.md §4.4.
 *
 * "see all unit" also confirms every Head-of group holds the base MEMBER role:
 * approvers get view rights by the client's model. That is a separate question
 * from whether the approval page REQUIRES Documents access in order to approve —
 * it does not; see documentsUnitFolderReady in ApprovalDocument.tsx.
 */
export const PERSONAS: Persona[] = [
  // ── Head of Department — the four bundles, at the DEPARTMENT tier ──────────
  { key: "hod1", family: "Head of Department", scope: "department", label: "1 — approve + upload",          roles: ["MEMBER", "UPL", "APR"],        summary: "Sees and uploads to every unit under the department, and approves." },
  { key: "hod2", family: "Head of Department", scope: "department", label: "2 — approve only",              roles: ["MEMBER", "APR"],               summary: "Sees every unit and approves, but cannot upload. Needs the DMS Approve level." },
  { key: "hod3", family: "Head of Department", scope: "department", label: "3 — approve + delete",          roles: ["MEMBER", "APR", "DEL"],        summary: "Approves and deletes approved documents. Cannot upload." },
  { key: "hod4", family: "Head of Department", scope: "department", label: "4 — approve + upload + delete", roles: ["MEMBER", "UPL", "APR", "DEL"], summary: "Uploads, approves, and deletes approved documents." },

  // ── Head of Unit — identical roles, at the UNIT tier ───────────────────────
  { key: "hou1", family: "Head of Unit", scope: "unit", label: "1 — approve + upload",          roles: ["MEMBER", "UPL", "APR"],        summary: "Sees and uploads to their own unit, and approves." },
  { key: "hou2", family: "Head of Unit", scope: "unit", label: "2 — approve only",              roles: ["MEMBER", "APR"],               summary: "Sees their unit and approves, but cannot upload. Needs the DMS Approve level." },
  { key: "hou3", family: "Head of Unit", scope: "unit", label: "3 — approve + delete",          roles: ["MEMBER", "APR", "DEL"],        summary: "Approves and deletes approved documents. Cannot upload." },
  { key: "hou4", family: "Head of Unit", scope: "unit", label: "4 — approve + upload + delete", roles: ["MEMBER", "UPL", "APR", "DEL"], summary: "Uploads, approves, and deletes approved documents." },

  // ── PIC ────────────────────────────────────────────────────────────────────
  { key: "pic1", family: "PIC", scope: "unit", label: "1 — upload + view", roles: ["MEMBER", "UPL"], summary: "Uploads, and reads their unit's approved documents." },
  {
    key: "pic2", family: "PIC", scope: "unit", label: "2 — Highly Confidential only", roles: ["HC"],
    summary: "Uploads and views Highly Confidential, and must NOT see Confidential or Restricted.",
    // Not a missing feature so much as a missing term: HC left Phase 1 on
    // 2026-08-01 and its term was deleted, so the level cannot be chosen on an
    // upload and the two HC libraries do not exist. The code is on
    // feat/hc-libraries; its §14 lists the four restore steps, two of which fail
    // silently.
    unavailable: "Phase 2 — Highly Confidential is out of scope and its term has been deleted from the term store.",
  },
  { key: "pic3", family: "PIC", scope: "unit", label: "3 — upload only", roles: ["UPL"], summary: "Uploads and sees only their own pending items. No access to Documents at all." },

  // ── SDG Employee ───────────────────────────────────────────────────────────
  { key: "employee", family: "SDG Employee", scope: "unit", label: "No power — view only", roles: ["MEMBER"], summary: "Reads their own unit's approved documents. Cannot upload or approve." },
];

/** Family headings in display order, so the picker cannot drift from the model. */
export const PERSONA_FAMILIES: Array<Persona["family"]> =
  ["Head of Department", "Head of Unit", "PIC", "SDG Employee"];

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
