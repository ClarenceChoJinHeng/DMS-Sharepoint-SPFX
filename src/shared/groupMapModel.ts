// Pure, SPFx-free helpers for the Group Map Builder. No @microsoft/* imports —
// keep this unit-testable in plain Jest (same pattern as formModel).

// naming.ts is pure too, so importing it here keeps that property.
import { cachedSiteEntryName } from "./naming";

/**
 * DEL, DELS and HC are library-scoped like the rest: DEL grants delete on
 * Documents, DELS grants delete on STAGING, HC grants Contribute on HC Approval
 * and Read on HC Library.
 *
 * DELS exists because UPL stopped carrying delete on 2026-08-04. Contribute's
 * Delete Items is folder-scoped, never author-scoped, so an uploader could delete
 * a colleague's pending file — which made the client's "a PIC needs the head of
 * unit's approval before deleting" rule unenforceable. UPL now maps to DMS Upload
 * (Contribute minus Delete Items) and PICs raise a deletion request instead;
 * DELS hands Staging delete back to the head-of groups that upload (Head of
 * Department/Unit 1 and 4).
 *
 * A separate role rather than a second upload level, so the capability is carried
 * by an explicit `_DELS` group name. Two upload suffixes differing only in whether
 * delete came along would be picked wrongly eventually, and nothing would surface
 * it until a file went missing.
 *
 * DEL and DELS share the "DMS Delete" permission level and are kept apart by the
 * library rule in FolderManager: DELS on Staging only, DEL on Documents only.
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
/**
 * GLOBAL is the C-level view role: reads every segment and everything beneath it, in
 * the DOCUMENTS library only. A C-level on Staging would be reading everyone's
 * unapproved drafts, which is the isolation rule the whole model rests on.
 *
 * It fans downward WITHOUT the recon_departmentFanOut switch, unlike a department-tier
 * row. That switch exists because a leftover tier row is indistinguishable, BY TIER,
 * from a deliberate wide grant, so tier-based fanning must be opt-in. A GLOBAL row
 * carries no tier at all and cannot be a leftover — the role has never granted
 * anything until now. The explicit name is the consent.
 *
 * SEGVIEW is RETIRED, one day after it was added: it scoped a C-level to a single
 * segment, and the client settled on global-only on 2026-08-04. It stays in the union,
 * out of SELECTABLE_ROLES, and deliberately absent from ROLE_TO_PERMISSION so
 * reconciliation grants it nothing. An inert role, not a deleted one.
 *
 * Deleting it outright would be worse than keeping it: roleFromGroupName would stop
 * recognising the `_SEGVIEW` suffix, so any group already created would fall through to
 * MEMBER and be granted Read at whatever tier its row sits on. A retired role that
 * grants nothing fails safe; a mis-parsed one grants access nobody asked for.
 */
/**
 * ENTRY means "this group may OPEN this thing" — a library or a page. It grants plain Read on
 * the list or the Site Pages item, and exists because nothing else can express that safely.
 *
 * One role for both scopes rather than two, because the grant is identical (Read) and the
 * TARGET already says which object it applies to. A second role would differ only in a field
 * the row already carries.
 *
 * Why it is not just the group's own role: a `_UPL` group granted `DMS Upload` at LIBRARY
 * level could add and edit anywhere in the library that inherits, not merely in its unit
 * folder. Library entry has to be Read, whatever the group does at folder level.
 *
 * Why it is not MEMBER: MEMBER means Read too, but LIBRARY_ROLES confines MEMBER to
 * Documents precisely because a MEMBER row on Staging at FOLDER scope hands a viewer other
 * people's pending drafts. Writing MEMBER-on-Staging rows for library entry would normalise
 * the exact row shape that is a security bug one scope down, and an admin skim-reading the
 * list could not tell the safe one from the dangerous one.
 *
 * It is absent from LIBRARY_ROLES, which is what keeps it library-only: the folder loop
 * grants a role only if that table lists it, so an ENTRY row can never reach a folder even
 * if someone writes one at Folder scope by hand.
 *
 * Absent from SELECTABLE_ROLES for the same reason HC is — the folder picker must never
 * offer it. The Staging Library Access tab writes it directly.
 */
export type GroupMapRole =
  | "MEMBER"
  | "UPL"
  | "APR"
  | "DEL"
  | "DELS"
  | "SEGVIEW"
  | "HC"
  | "ENTRY"
  | "GLOBAL";

/** The role a library-entry row carries. Named so callers never spell it as a literal. */
export const LIBRARY_ENTRY_ROLE: GroupMapRole = "ENTRY";

/**
 * Roles that legitimately act on Staging, and therefore the only groups the Staging Library
 * Access tab offers entry to. A viewer-only group reaching Staging would be reading other
 * people's unapproved drafts.
 *
 * GLOBAL is deliberately ABSENT, and it was briefly added on 2026-08-04 before being reverted
 * the same day. It is the C-level view role and it is DOCUMENTS-ONLY: "sees everything" means
 * every approved document in every segment, never the unapproved drafts. Putting it here would
 * hand a C-level every unit's pending work — the isolation rule the whole model rests on, and
 * the widest accidental grant available in this system.
 */
export const STAGING_FACING_ROLES: GroupMapRole[] = ["UPL", "APR", "DELS"];

/**
 * The roles an administrator may pick in the builder. Excludes HC (Phase 2, see
 * above). Keeping the list here rather than in the component means the type and
 * the picker cannot drift apart.
 */
// SEGVIEW rejoined 2026-08-07, when the client asked for a per-segment C-Level alongside the
// global one. HC stays out: it is Phase 2, its term is deleted, and offering it would let an
// admin author a row that can never be granted.
export const SELECTABLE_ROLES: GroupMapRole[] =
  ["MEMBER", "UPL", "APR", "DEL", "DELS", "GLOBAL", "SEGVIEW"];

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
  family: "C-Level" | "Head of Department" | "Head of Unit" | "PIC" | "SDG Employee";
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
  scope: "segment" | "department" | "unit";
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
  // ── C-Level — view-only, above the department ──────────────────────────────
  // The client's "Head of everything". Named C-Level rather than "Head of …" on
  // purpose: every Head-of family carries APR by definition, and these two carry
  // no power at all. A family whose name began "Head of" would be asserting a
  // capability it does not have.
  //
  // Both are marked unavailable until reconciliation fans a segment-tier grant
  // downward. Listed anyway, because the capability is agreed and an admin who
  // cannot find it will assume it was forgotten.
  {
    key: "clevel_global", family: "C-Level", scope: "segment", label: "Global — view everything",
    roles: ["GLOBAL"],
    summary: "Reads every segment, department and unit in the Documents library, down to Document Type. No Staging access, no upload, no approve.",
  },

  // One segment, not all twelve. GLOBAL narrowed by a term — see the SEGVIEW note in
  // ROLE_TO_PERMISSION. Documents only, for the same reason as GLOBAL: a C-Level on Staging
  // would be reading an entire segment's unapproved drafts.
  {
    key: "clevel_segment", family: "C-Level", scope: "segment", label: "Segment — view one business segment",
    roles: ["SEGVIEW"],
    summary: "Reads one business segment and every department and unit under it, in the Documents library. No Staging access, no upload, no approve.",
  },

  // ── Head of Department — view + delete, at the DEPARTMENT tier ─────────────
  //
  // ONE persona, not four. The four bundles existed to combine approve/upload/delete at
  // department tier; the client's 2026-08-07 restatement moved approval to Head of Unit and
  // left HoD with breadth as its only distinction.
  //
  // DEL is Documents-only by LIBRARY_ROLES, so an HoD group never reaches Staging. Note the
  // grant FANS OUT: department scope plus recon_departmentFanOut puts delete on every unit
  // folder under the department, not just the department folder. Wide, and intended.
  {
    key: "hod", family: "Head of Department", scope: "department", label: "View + delete, department-wide",
    roles: ["MEMBER", "DEL"],
    summary: "Reads every unit under their department in Documents, and can delete approved documents there. Cannot upload, cannot approve, and has no Staging access.",
  },

  // ── Head of Unit — the approver, at the UNIT tier ──────────────────────────
  //
  // ONE persona, not four, and the only family that approves. MEMBER rides with APR because
  // the client's model gives every head-of group view rights; it is not a technical
  // requirement of approving (see documentsUnitFolderReady in ApprovalDocument.tsx).
  {
    key: "hou", family: "Head of Unit", scope: "unit", label: "Approve + view own unit",
    roles: ["MEMBER", "APR"],
    summary: "Approves every file in their own unit and reads the unit's approved documents. Cannot upload or delete, and sees no sibling unit.",
  },

  // ── PIC ────────────────────────────────────────────────────────────────────
  //
  // ONE upload persona, not the client document's three. Corrected 2026-08-04, after the
  // three collapsed into each other:
  //
  // PIC 1 was modelled as MEMBER + UPL, reading "and view" as Documents access. It is not:
  // the client's line is "can see the files in the unit", meaning the unit's STAGING files,
  // which UPL already provides. A PIC's job is to submit documents; whether they may also
  // read the approved archive is a separate decision, granted by adding them to the unit's
  // base group — which is the SDG Employee role. Bundling it here made every PIC a
  // Documents reader by default, and that is the wrong default for a permission.
  //
  // PIC 3 was "upload only, own files only". Once "own files only" is dropped as
  // unachievable (below) it is UPL alone — identical to a corrected PIC 1. Two identical
  // entries in a picker are a trap, not a choice, so there is one.
  //
  // "Sees only their own pending items" is deliberately absent from the summary. View Items
  // and Delete Items are folder-scoped, never author-scoped; the one author-scoped
  // SharePoint feature (Item-level Permissions under Advanced Settings) exists for lists,
  // not document libraries; and Staging's Draft Item Security is deliberately "any user who
  // can read items", because the stricter value hid pending folders from readers and 403'd
  // Contribute uploaders (memory dms-content-approval-blocks-uploader). It is a
  // LIBRARY-wide setting, so it could not vary per persona even if we wanted it to.
  //
  // CONFIDENTIALITY IS NOT A PERMISSION. Confirmed 2026-08-07: any PIC may upload at any
  // confidentiality level, all three. So there is no HC persona and no confidentiality-scoped
  // upload role — the level is metadata the uploader chooses on the form, and it drives
  // labelling, not access. The old "Highly Confidential only" persona is removed rather than
  // left greyed out: a picker entry the client will never choose is a question they have to
  // ask once and get answered every time.
  //
  // The HC role string itself is KEPT and simply unused, for the same reason as DELS — the
  // Phase 2 HC-library work on feat/hc-libraries still refers to it, and removing it here
  // means re-adding suffix parsing and permission mapping the day it comes back.
  {
    key: "pic", family: "PIC", scope: "unit", label: "Upload",
    roles: ["UPL"],
    summary: "Uploads to their unit at any confidentiality level, and sees the unit's pending files. Cannot approve or delete, and has no Documents access unless also added to the unit's base group.",
  },

  // ── SDG Employee ───────────────────────────────────────────────────────────
  { key: "employee", family: "SDG Employee", scope: "unit", label: "No power — view only", roles: ["MEMBER"], summary: "Reads their own unit's approved documents. Cannot upload or approve." },
];

/** Family headings in display order, so the picker cannot drift from the model. */
export const PERSONA_FAMILIES: Array<Persona["family"]> =
  ["C-Level", "Head of Department", "Head of Unit", "PIC", "SDG Employee"];

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
/**
 * A FUNCTION, not a constant, because the name is discovered per site — CRS_SITE_MEMBERS on a
 * renamed site, DMS_SITE_MEMBERS otherwise. Resolved by primeNames(); see cachedSiteEntryName.
 */
export function siteEntryGroupTitle(): string {
  return cachedSiteEntryName();
}

/**
 * What a Group Map row grants access to. See the access-scope-mapping spec.
 *
 * `Folder` is everything this list held before 2026-08-04, and `Segment`/`UnitTermGuid`
 * are meaningful ONLY for it. `Site` and `Library` grants were previously made by hand in
 * native SharePoint across four separate screens with nothing recording that the step had
 * been done — which is why the site-entry group goes missing so reliably.
 */
export type GroupMapScope = "Site" | "Library" | "Folder" | "Page";

export const GROUP_MAP_SCOPES: GroupMapScope[] = ["Site", "Library", "Folder", "Page"];

/**
 * A blank Scope reads as `Folder`.
 *
 * Every row that exists today is a folder row and the column arrives after them, so the
 * default has to be the historical meaning. Treating blank as anything else — or as
 * invalid — would make an existing site's folder permissions silently stop being applied
 * the moment the column is added. That is the worst shape of failure available here: the
 * run still reports success, and nobody discovers it until a new unit is onboarded.
 */
export function normalizeScope(raw: string | string[] | undefined): GroupMapScope {
  // Accepts an ARRAY as well as a string, because Scope is a Choice column and nobody can
  // tell single- from multi-select by looking at the form: both render the value as a pill.
  // A multi-select column returns ["Library"] (and null when emptied — CLAUDE.md #11), so a
  // string-only parser would throw "raw.trim is not a function" and take the whole tab down
  // over a checkbox in the column settings. Taking the first entry degrades to the right
  // answer instead; a row has exactly one scope, so there is never a second to lose.
  const first = Array.isArray(raw) ? raw[0] : raw;
  const v = (first ?? "").trim().toLowerCase();
  if (v === "site") return "Site";
  if (v === "library") return "Library";
  if (v === "page") return "Page";
  return "Folder";
}

/** The exact field set POSTed to the DMS Group Map list. */
export interface GroupMapWriteRow {
  GroupId: string;
  GroupName: string;
  Segment: string;       // "" unless Scope is Folder
  UnitTermGuid: string;  // "" unless Scope is Folder; equals Segment for a segment-tier row
  Role: GroupMapRole;
  Scope: GroupMapScope;
  Target: string;        // library title or page file name; "" for Site and Folder
}

/** In-progress selections from the builder UI. */
export interface GroupMapDraft {
  groupId: string;
  groupName: string;
  role: GroupMapRole;
  scope?: GroupMapScope;  // absent reads as Folder
  target?: string;        // required for Library and Page; ignored otherwise
  segmentGuid?: string;   // required for Folder, unless GLOBAL
  tierGuid?: string;      // required for Folder, unless GLOBAL; may equal segmentGuid
}

/**
 * The one page a Page row must never target.
 *
 * Restricting the site's welcome page makes site entry worthless: the user is granted Read
 * on the web and then gets 403 on the only thing they can navigate to. Refused at write
 * time rather than reported after a run, because by then the grant exists and nothing in
 * this system removes one.
 */
export const FORBIDDEN_PAGE_TARGETS = ["home.aspx", "sitepages/home.aspx"];

export function isForbiddenPageTarget(target: string): boolean {
  const t = (target ?? "").trim().toLowerCase().replace(/^\/+/, "");
  return FORBIDDEN_PAGE_TARGETS.indexOf(t) !== -1;
}

const norm = (s: string): string => (s ?? "").trim();

/**
 * Build the row to POST. Trims every field (kills the trailing-space bug that
 * silently dropped DMS_GHO_LRC). A GLOBAL row carries no term, so Segment and
 * UnitTermGuid are forced empty regardless of any stale draft selections.
 */
export function buildGroupMapRow(draft: GroupMapDraft): GroupMapWriteRow {
  const scope = normalizeScope(draft.scope);
  // Terms describe a folder and nothing else. A Site, Library or Page row that carried a
  // leftover segment/tier from the form's previous state would look, to reconciliation and
  // to the orphan-repair pass, exactly like a folder row — so it would be matched against
  // the term store, reported as an orphan when that term was renamed, and potentially
  // re-pointed. Forced empty here rather than filtered later, for the same reason GLOBAL
  // already was.
  const termless = scope !== "Folder" || draft.role === "GLOBAL";
  // Target belongs to Library and Page only. Left on a Site row it would read as a
  // library title and break that library's inheritance.
  const targeted = scope === "Library" || scope === "Page";
  return {
    GroupId: norm(draft.groupId),
    GroupName: norm(draft.groupName),
    Segment: termless ? "" : norm(draft.segmentGuid ?? ""),
    UnitTermGuid: termless ? "" : norm(draft.tierGuid ?? ""),
    Role: draft.role,
    Scope: scope,
    Target: targeted ? norm(draft.target ?? "") : "",
  };
}

/** True if an equivalent row exists: GroupId + UnitTermGuid + Role, normalised. */
export function isDuplicateRow(
  existing: GroupMapWriteRow[],
  candidate: GroupMapWriteRow,
): boolean {
  // Scope and Target join the key. Without them a Library row for "Documents" and one for
  // "Staging" are the same row (both termless), so the second would be silently rejected
  // as a duplicate — and the admin would be told the mapping already existed while the
  // library they actually wanted stayed unmapped.
  const key = (r: GroupMapWriteRow): string =>
    [
      norm(r.GroupId).toLowerCase(),
      normalizeScope(r.Scope),
      norm(r.Target).toLowerCase(),
      norm(r.UnitTermGuid).toLowerCase(),
      norm(r.Role).toUpperCase(),
    ].join("|");
  const k = key(candidate);
  return existing.some((r) => key(r) === k);
}

/**
 * Every name suffix that maps to a role, longest first.
 *
 * A TABLE rather than a chain of `endsWith` calls, and the sort is the whole point.
 * Suffixes collide by prefix — "_DEL" is a prefix of "_DELS", and "_UPL" is a prefix of
 * "_UPLOADER" — so a hand-ordered chain is only correct while every future editor honours
 * a comment telling them which line must come first. Sorting by length makes the correct
 * order a property of the data instead.
 *
 * The failure this prevents is not cosmetic: an unmatched suffix falls through to MEMBER,
 * which LIBRARY_ROLES puts on Documents. A group the admin named as a Staging uploader
 * would silently become a Documents reader — wrong library, wrong permission, no error.
 *
 * Both forms are accepted forever. The client moved to long-form names on 2026-08-04
 * (GHO_GF_CORU_UPLOADER), but groups already exist on the test site with _UPL/_APR and
 * stranding them would break live mappings. Only NEW names are written long-form.
 *
 * GLOBAL is absent deliberately: it is a privileged bypass the admin picks by hand, never
 * derived from a name. MEMBER is absent because it IS the no-suffix base group.
 */
/**
 * Normalise a `Role` COLUMN value to the short code the permission tables key on.
 *
 * The group NAME suffixes were changed to long form in 2026-08-05 (`_UPLOADER`, `_APPROVER`),
 * so an admin filling in the Group Map by hand naturally types "UPLOADER" in the Role column
 * too — and on 2026-08-07 one did, on the only row on the test site. Nothing normalised it, so
 * `accepts()` found no permission level and the row was SKIPPED. The log then said "no
 * group-map groups for this unit (locked admin-only)", which reads as a missing row rather than
 * a rejected one: the admin goes looking for a row that is sitting right there.
 *
 * Accepting both forms costs nothing and is unambiguous — no long form is a prefix of another
 * short code. Returns "" for anything unrecognised, which callers already treat as skip.
 */
const ROLE_ALIASES: Record<string, GroupMapRole> = {
  UPLOADER: "UPL",
  APPROVER: "APR",
  DELETER_DOCUMENTS: "DEL",
  DELETER_STAGING: "DELS",
  MEMBER: "MEMBER",
  VIEWER: "MEMBER",
  SEGMENTVIEW: "SEGVIEW",
};

/** Every short code the permission tables key on. Listed, not derived — the union is a type. */
const SHORT_ROLE_CODES = ["MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "HC", "GLOBAL", "ENTRY"];

export function normalizeRoleValue(raw: string): string {
  const v = (raw ?? "").trim().toUpperCase().replace(/[\s-]+/g, "_");
  if (v.length === 0) return "";
  // A short code wins outright, so this can never re-map an already-valid value.
  if (SHORT_ROLE_CODES.indexOf(v) !== -1) return v;
  // Unrecognised values pass through unchanged, so they still fail accepts() and are skipped
  // rather than silently becoming some role nobody asked for.
  return ROLE_ALIASES[v] ?? v;
}

export type RoleSuffix = { suffix: string; role: GroupMapRole };

const ROLE_SUFFIXES_UNSORTED: RoleSuffix[] = [
  { suffix: "_DELETER_DOCUMENTS", role: "DEL" },
  { suffix: "_DELETER_STAGING", role: "DELS" },
  { suffix: "_APPROVER", role: "APR" },
  { suffix: "_UPLOADER", role: "UPL" },
  { suffix: "_SEGVIEW", role: "SEGVIEW" },
  { suffix: "_DELS", role: "DELS" },
  { suffix: "_APR", role: "APR" },
  { suffix: "_UPL", role: "UPL" },
  { suffix: "_DEL", role: "DEL" },
  { suffix: "_HC", role: "HC" },
];

// Copied before sorting: sort() mutates, and mutating the source array would make the
// declaration order above meaningless to read.
export const ROLE_SUFFIXES: RoleSuffix[] = ROLE_SUFFIXES_UNSORTED.slice().sort(
  (a, b) => b.suffix.length - a.suffix.length,
);

/**
 * The suffix written into a NEW group name for a role — the long form, since that is the
 * convention the client asked for. MEMBER and GLOBAL return "" (base group / hand-picked).
 */
export function suffixForRole(role: GroupMapRole | ""): string {
  if (role === "MEMBER" || role === "GLOBAL" || role === "") return "";
  for (const s of ROLE_SUFFIXES) if (s.role === role) return s.suffix;
  return `_${role}`;
}

/**
 * Derive the intended role from a group name's suffix (naming convention:
 * <seg>_<dept>_<unit>[_UPLOADER|_APPROVER|…]). Case-insensitive, trims.
 *
 * Any name with no recognised suffix — including a base viewer group — is MEMBER.
 */
export function roleFromGroupName(name: string): GroupMapRole {
  const n = norm(name).toUpperCase();
  for (const { suffix, role } of ROLE_SUFFIXES) {
    if (n.endsWith(suffix)) return role;
  }
  return "MEMBER";
}

/** Field-level validation for enabling the Add button. Returns [] when valid. */
export function validateDraft(draft: GroupMapDraft): string[] {
  const errors: string[] = [];
  if (!norm(draft.groupId)) errors.push("Select a group.");
  if (!draft.role) errors.push("Select a role.");
  const scope = normalizeScope(draft.scope);
  // Each scope requires a different set of fields, which is why the picker asks for the
  // scope FIRST. Validating them all unconditionally would demand a segment for a Site row
  // and a target for a Folder row, and an admin cannot satisfy both.
  if (scope === "Folder") {
    if (draft.role !== "GLOBAL") {
      if (!norm(draft.segmentGuid ?? "")) errors.push("Select a segment.");
      if (!norm(draft.tierGuid ?? "")) errors.push("Select a tier.");
    }
  } else if (scope === "Library") {
    if (!norm(draft.target ?? "")) errors.push("Enter the library name.");
  } else if (scope === "Page") {
    const target = norm(draft.target ?? "");
    if (!target) errors.push("Enter the page file name, e.g. Upload.aspx.");
    // Refused here, before the row is ever written. Reconciliation refusing it later
    // would be too late: the row would sit in the list looking valid, and every run
    // would report the same rejection with nobody able to tell whether it had ever
    // applied.
    else if (isForbiddenPageTarget(target)) {
      errors.push("The site home page cannot be restricted — users would be granted site access and then blocked from the only page they can reach.");
    }
  }
  return errors;
}

/**
 * Suggest a site-group title from the builder's current selections, e.g.
 * "GHO_GF_CORU_UPLOADER". Purely a SUGGESTION — the admin edits it freely before create
 * (no abbreviation algorithm by design; see the native-sharepoint-groups spec).
 *
 * NO "DMS_" prefix as of 2026-08-04, at the client's request, and the long-form role
 * suffix for the same reason. Nothing identifies our groups by title any more — the picker
 * excludes built-ins by id and the site-entry sync reads Group Map's GroupId column — so
 * the prefix had no remaining job. See the 2026-08-04 amendment in
 * docs/superpowers/specs/2026-07-22-role-from-group-name-suffix.md.
 */
export function suggestGroupName(
  segmentLabel: string,
  tierLabels: string[],
  role: GroupMapRole | "",
): string {
  if (role === "GLOBAL") return "GLOBAL";
  const parts = [norm(segmentLabel), ...tierLabels.map(norm)].filter(Boolean);
  // MEMBER is the base group and carries NO suffix — that is the convention
  // roleFromGroupName reads back, so adding one here would make every base group
  // parse as an unknown role. An unpicked role ("") is the same: the admin is
  // still choosing, and a trailing "GHO_" is not a name to hand them.
  return parts.join("_") + suffixForRole(role);
}
