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
 * HC IS BACK, AND IT IS NOT THE OLD `HC` ROLE (2026-08-15). Highly Confidential
 * now has its own LIBRARY PAIR — `HC Approval Document` and `HC Documents` — so
 * clearance is a grant on those libraries, not a permission break on a folder.
 * The old single `HC` role is replaced by `UPLHC` and `APRHC`, because with two
 * libraries and two sides "cleared" no longer says enough: cleared to upload, or
 * cleared to approve?
 *
 * Both are SUPERSETS — an HC uploader files at any level, an HC approver approves
 * in either library. What keeps HC separate is the reverse: plain UPL and APR are
 * absent from the HC rows of LIBRARY_ROLES entirely.
 *
 * See docs/superpowers/specs/2026-08-03-role-personas-and-department-fanout-design.md
 * and 2026-08-15-highly-confidential-library-design.md. The 2026-07-16 HC spec is
 * SUPERSEDED — do not implement its elevated-flow design.
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
  // Highly Confidential, 2026-08-15. HC lives in its own LIBRARY PAIR, so clearance is a grant on
  // those libraries rather than a property of a folder — see
  // docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md.
  //
  // BOTH ARE SUPERSETS: UPLHC uploads to the normal approval library as well, APRHC approves in
  // both. The client's rule, stated for uploaders and applied symmetrically. The separation runs the
  // other way — plain UPL and APR reach the HC libraries NOT AT ALL, and that single absence from
  // LIBRARY_ROLES is the whole feature.
  //
  // These REPLACE the old `HC` role, which belonged to the superseded 2026-07-16 design where HC was
  // a secured subfolder and one group was simply "cleared". With two libraries and two sides,
  // "cleared" is ambiguous — cleared to upload, or to approve? — so the role says which. The legacy
  // `_HC` group-name suffix now parses as UPLHC, which is what such a group meant.
  | "UPLHC"
  | "APRHC"
  | "ENTRY"
  | "GLOBAL"
  // SHARE, 2026-08-15. The right to grant SOMEONE ELSE access to a document — which is what makes a
  // share-request approver able to carry out what they approve. Granting access needs Manage
  // Permissions, and no role above carries it.
  //
  // Its OWN role, never a property of APR, because it is the widest thing in this table: a holder can
  // give any access, to anyone the tenant allows, anywhere their grant reaches. Folded into an
  // existing role it would spread silently the next time that role joined a persona.
  | "SHARE";

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
// UPLHC and APRHC joined 2026-08-15. They face BOTH approval libraries, so a persona carrying one
// belongs under the approval-library half of the Folder Access picker exactly as UPL and APR do.
// Being derived rather than listed is what stops the HC personas being filed under Documents, which
// would present the widest grant in the system as an ordinary one.
export const STAGING_FACING_ROLES: GroupMapRole[] = ["UPL", "APR", "DELS", "UPLHC", "APRHC"];

/**
 * What to SHOW instead of the role code. Display only — the stored `Role` value stays the short
 * code, because reconciliation keys on it and every existing row already carries it.
 *
 * Added 2026-08-09 after an admin nearly deleted a live mapping: two rows read
 * `GHO_GF_CORU_HOU … APR` and `… DELS`, and "DELS" meant nothing to them. Codes are fine in a
 * table their author reads daily, and wrong in one an administrator meets twice a year.
 *
 * **`DEL` and `DELS` are named by LIBRARY, and never abbreviated toward each other.** The first
 * instinct was to rename `DELS` to `DEL`, which would have left two roles both shown as "DEL" —
 * and they are not variations of one capability: `DEL` removes an APPROVED document from
 * Documents, `DELS` removes a PENDING file from the approval library. Telling those apart is the
 * entire reason LIBRARY_ROLES exists, so the labels have to do it too. One letter of difference
 * is what made this confusing in the first place.
 */
export const ROLE_LABEL: Record<string, string> = {
  MEMBER:  "View only",
  UPL:     "Uploader",
  APR:     "Approver",
  DEL:     "Delete approved documents",
  DELS:    "Delete pending files",
  GLOBAL:  "C-Level — all segments",
  SEGVIEW: "C-Level — one segment",
  ENTRY:   "Library / page entry",
  // Named for the CLEARANCE, not the library, because these roles grant in the normal libraries too.
  // "HC uploader" would read as "uploads only to HC", which is the opposite of what they do.
  UPLHC:   "Uploader — Highly Confidential cleared",
  APRHC:   "Approver — Highly Confidential cleared",
};

/** The label for a role code, falling back to the code so an unknown value stays visible. */
export function roleLabel(role: string): string {
  const r = (role ?? "").trim().toUpperCase();
  return ROLE_LABEL[r] ?? r;
}

/**
 * The canonical set of roles that can be DERIVED FROM A GROUP NAME — not, since 2026-08-09, a
 * list of anything the UI offers.
 *
 * The Folder Access page no longer has a role picker: an admin chooses a PERSONA and its roles
 * are applied. The name is kept because the group-name round-trip tests iterate this list, and
 * because it still answers "which roles are legal in a Group Map row". Read "selectable" as
 * "authorable", not "on screen" — the picker it was named for no longer exists.
 *
 * Excludes HC (Phase 2, see above) and ENTRY (library-scope only, written by another tab).
 */
// SEGVIEW rejoined 2026-08-07, when the client asked for a per-segment C-Level alongside the
// global one. HC stays out: it is Phase 2, its term is deleted, and offering it would let an
// admin author a row that can never be granted.
export const SELECTABLE_ROLES: GroupMapRole[] =
  ["MEMBER", "UPL", "APR", "DEL", "DELS", "GLOBAL", "SEGVIEW", "UPLHC", "APRHC"];

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
    key: "clevel_global", family: "C-Level", scope: "segment", label: "Global — view, delete + share everything",
    // DEL and SHARE added 2026-08-15. NO LONGER VIEW-ONLY: a global C-Level deletes or shares any
    // approved document anywhere, without asking anyone. The widest grant in this file.
    roles: ["GLOBAL", "DEL", "SHARE"],
    summary: "Reads every segment, department and unit in the Documents library, down to Document Type. No Staging access, no upload, no approve.",
  },

  // One segment, not all twelve. GLOBAL narrowed by a term — see the SEGVIEW note in
  // ROLE_TO_PERMISSION. Documents only, for the same reason as GLOBAL: a C-Level on Staging
  // would be reading an entire segment's unapproved drafts.
  {
    key: "clevel_segment", family: "C-Level", scope: "segment", label: "Segment — view, delete + share one segment",
    // As clevel_global, narrowed to one segment by the row's term.
    roles: ["SEGVIEW", "DEL", "SHARE"],
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
  //
  // MEMBER DROPPED 2026-08-09, and unlike Head of Unit this one was redundant all along:
  // DEL maps to "CRS Delete", which is Read + Delete Items. The read was already there; the
  // second membership only made it look as though it were not.
  {
    key: "hod", family: "Head of Department", scope: "department", label: "View, delete + share, department-wide",
    // SHARE added 2026-08-15 — HoD shares without approval, like C-Level. DEL was already here.
    roles: ["DEL", "SHARE"],
    summary: "Reads every unit under their department in Documents, and can delete approved documents there. Cannot upload, cannot approve, and has no Staging access.",
  },

  // ── Head of Unit — the approver, at the UNIT tier ──────────────────────────
  //
  // ONE persona, not four, and the only family that approves.
  //
  // MEMBER DROPPED 2026-08-09. It used to ride along purely to grant Documents read, which
  // APR now carries by itself (LIBRARY_ROLES lists APR under Documents, downgraded to Read).
  // Keeping it would leave the admin two memberships to maintain and no way to tell which one
  // was load-bearing. See 2026-08-09-persona-driven-folder-access-design.md.
  //
  // DELS ADDED the same day. It is approval-library delete — the power to remove a pending or
  // rejected file — and it belonged to NO persona before, surviving only so a hand-authored
  // row would still work. That is the Head of Unit's job by the client's own description.
  // It is deliberately NOT Documents delete: LIBRARY_ROLES keeps DELS off Documents, so a
  // Head of Unit still cannot remove an APPROVED document. That stays with Head of Department.
  {
    key: "hou", family: "Head of Unit", scope: "unit", label: "Approve, upload, delete + share own unit",
    // Three roles added 2026-08-15, correcting the model and enabling the request workflow.
    //
    // UPL — the client had said a Head of Unit cannot upload; they can. Consequence, accepted
    // explicitly: a HoU approves their OWN uploads, so the approval step is skippable by the person
    // who files most often. The alternatives were a deadlock in any single-HoU unit, or giving Heads
    // of Department Approve, which would also hand them every draft in the department.
    //
    // DEL and SHARE — because AN APPROVER CAN ONLY APPROVE WHAT THEY CAN PERFORM. Without DEL an
    // approved deletion request fails at the last step; without SHARE an approved share does. SHARE
    // is the wide one: it makes a HoU the sharing AUTHORITY for their unit, able to share directly
    // without any screen, not merely an approver of other people's requests.
    roles: ["APR", "DELS", "UPL", "DEL", "SHARE"],
    summary: "Approves every file in their own unit, can delete pending or rejected files there, and reads the unit's approved documents. Cannot upload, cannot delete approved documents, and sees no sibling unit.",
  },

  // ── Head of Unit, Highly Confidential ──────────────────────────────────────
  //
  // The client's reason for a separate approver group, in their words: "if client wants to ensure
  // not all approver can see HC files". So this is the persona a unit's HC approver gets, and the
  // ordinary `hou` persona reaches no HC library at all.
  //
  // Carries UPLHC as well as APRHC for the same reason `hou` carries UPL: a Head of Unit uploads,
  // and an HC-cleared one uploads HC. The self-approval consequence noted on `hou` therefore applies
  // to HC too, and is worse there — the person approving an HC document may be the person who filed
  // it, in the one place where a second pair of eyes matters most. Stated, not hidden; the client
  // separates the groups if they want it separated.
  {
    key: "hou_hc", family: "Head of Unit", scope: "unit", label: "Approve incl. Highly Confidential, upload, delete + share",
    roles: ["APRHC", "UPLHC", "DELS", "DEL", "SHARE"],
    summary: "Everything a Head of Unit does, plus approving and reading Highly Confidential documents for this unit. A Head of Unit without this clearance cannot see them at all.",
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
  //
  // 2026-08-09: UPL now grants Documents READ as well, so a PIC no longer needs the base
  // group. The paragraph above still holds — whether a PIC may read the approved archive was
  // a separate DECISION — but the client has now made it, for every PIC. What changed is the
  // answer, not the reasoning: it is granted by the same group rather than by a second one.
  {
    key: "pic", family: "PIC", scope: "unit", label: "Upload + delete own pending",
    // DELS added 2026-08-15, correcting the model: the client had said a PIC cannot delete, and the
    // rule is the opposite in the APPROVAL LIBRARY.
    //
    // It needs no "own files only" rule, and that is the point. DELS is Staging-only by
    // LIBRARY_ROLES, and Draft Item Security already hides a peer's pending work from a PIC — so
    // "delete what you can see" IS "delete your own". The permission and the visibility are the same
    // boundary, which means there is nothing extra to enforce and nothing that can drift apart.
    //
    // In DOCUMENTS a PIC still deletes nothing and shares nothing: both are requests the Head of
    // Unit approves. See 2026-08-15-deletion-and-share-requests-design.md.
    roles: ["UPL", "DELS"],
    summary: "Uploads to their unit at any confidentiality level, sees the unit's pending files, and reads the unit's approved documents. Cannot approve or delete.",
  },

  // ── PIC, Highly Confidential ───────────────────────────────────────────────
  //
  // The persona the 2026-08-04 note said "cannot exist in Phase 1". It can now: HC has its own
  // library pair, so clearance is a grant rather than a term nobody can select.
  //
  // It REPLACES the plain PIC for a cleared person — it does not accompany it. UPLHC is a superset
  // that covers the normal approval library too, so putting someone in both groups adds nothing and
  // leaves two rows to keep in step through a term rename. One person, one group, as everywhere else
  // since 2026-08-09.
  {
    key: "pic_hc", family: "PIC", scope: "unit", label: "Upload incl. Highly Confidential + delete own pending",
    roles: ["UPLHC", "DELS"],
    summary: "Everything a PIC does, plus filing and reading Highly Confidential documents for this unit. A PIC without this clearance is never even shown the Highly Confidential level.",
  },

  // ── SDG Employee ───────────────────────────────────────────────────────────
  //
  // Now the ONLY persona whose whole purpose is MEMBER. Since 2026-08-09 every other persona
  // carries its own Documents read, so this one means exactly what its name says: someone who
  // views and does nothing else.
  { key: "employee", family: "SDG Employee", scope: "unit", label: "No power — view only", roles: ["MEMBER"], summary: "Reads their own unit's approved documents. Cannot upload or approve." },
];

/** Family headings in display order, so the picker cannot drift from the model. */
export const PERSONA_FAMILIES: Array<Persona["family"]> =
  ["C-Level", "Head of Department", "Head of Unit", "PIC", "SDG Employee"];

/**
 * Does this persona act on the APPROVAL LIBRARY at all? Drives the Folder Access page's
 * library toggle (2026-08-09).
 *
 * DERIVED from STAGING_FACING_ROLES rather than listed as a per-persona field, because a hand-
 * maintained list is a second place for the same fact to be wrong — and being wrong here is not
 * cosmetic. The client's first draft of this toggle filed the C-Level personas under the
 * approval library; presenting GLOBAL/SEGVIEW as an approval-library shape is the widest
 * accidental grant this system can suggest, since a segment-wide viewer there reads every
 * unapproved draft in the segment. Deriving it means the grouping cannot disagree with the
 * roles the persona actually carries.
 *
 * Note "at all": since 2026-08-09 Head of Unit and PIC grant in BOTH libraries. They are filed
 * under the approval library because that is what distinguishes them — every other persona is
 * Documents-only.
 */
export function personaTouchesStaging(p: Persona): boolean {
  for (const r of p.roles) if (STAGING_FACING_ROLES.indexOf(r) > -1) return true;
  return false;
}

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
 * Is this the site-entry group?
 *
 * Case-insensitive and trimmed, because the title arrives from three different places — a search
 * result, a Group Map row, the name box on a create form — and only one of them carries the exact
 * stored casing. An exact-match comparison fails in the dangerous direction: it would answer "no"
 * for the entry group itself, and the caller would then try to add the entry group to itself.
 */
export function isSiteEntryGroupTitle(title: string | undefined): boolean {
  return (title ?? "").trim().toLowerCase() === siteEntryGroupTitle().trim().toLowerCase();
}

/**
 * Find the site-entry group among a list of site groups.
 *
 * `undefined` means ABSENT — never "could not tell". The caller must have read the group list
 * successfully before calling, because absent and unreadable need different messages: one is a
 * setup step the admin can take, the other is a transient error they cannot.
 */
export function findSiteEntryGroup<T extends { title: string }>(
  groups: readonly T[] | undefined,
): T | undefined {
  for (const g of groups ?? []) if (isSiteEntryGroupTitle(g.title)) return g;
  return undefined;
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
const SHORT_ROLE_CODES = ["MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "UPLHC", "APRHC", "GLOBAL", "ENTRY"];

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
  // Highly Confidential, 2026-08-15. BOTH spellings are accepted: the long one is what the client
  // wrote, the short one is what an administrator types when the name is already long at four tiers.
  //
  // Precedence is handled by the length sort below, not by the order here — `_UPL_HC` (7) is tested
  // before `_UPL` (4) and before `_HC` (3), so a group named `..._UPL_HC` can never parse as a plain
  // uploader. Worth stating because getting it wrong would grant HC clearance to nobody while
  // looking entirely correct in the list.
  { suffix: "_UPL_HIGHLY_CONFIDENTIAL", role: "UPLHC" },
  { suffix: "_APR_HIGHLY_CONFIDENTIAL", role: "APRHC" },
  { suffix: "_UPL_HC", role: "UPLHC" },
  { suffix: "_APR_HC", role: "APRHC" },
  // LEGACY. Under the superseded 2026-07-16 design a `..._HC` group was the unit's HC-cleared
  // members — its uploaders. Kept so such a group still parses as something sensible rather than
  // falling through to MEMBER, which would grant plain Read at whatever tier its row sits on.
  { suffix: "_HC", role: "UPLHC" },
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

/**
 * Characters SharePoint refuses in a site-group title.
 *
 * Checked here rather than left to the server because the server's rejection is an HTTP 500
 * naming none of them, arriving after the admin has typed a name and staged members.
 */
const ILLEGAL_GROUP_NAME_CHARS = ['"', "#", "%", "&", "*", ":", "<", ">", "?", "\\", "/", "{", "|", "}", "~"];

/** SharePoint's title limit. */
const GROUP_NAME_MAX = 255;

/**
 * Validate a new group's name. Returns [] when it can be created.
 *
 * `existingTitles` catches the duplicate BEFORE the request, so the message can say which name
 * clashes. `createSiteGroup` still throws DUPLICATE_GROUP — this check is a courtesy, not the
 * guard: another admin may create the same name between the page loading and the button being
 * pressed, and only the server sees that.
 *
 * Compared case-insensitively because SharePoint group titles are — "gho_gf_coru_uploader" and
 * "GHO_GF_CORU_UPLOADER" cannot coexist, and reporting them as available would produce a failure
 * the admin cannot explain from what is on screen.
 */
export function validateGroupName(name: string, existingTitles?: readonly string[]): string[] {
  const errors: string[] = [];
  const n = norm(name);
  if (!n) {
    errors.push("Enter a group name.");
    return errors;
  }
  if (n.length > GROUP_NAME_MAX) errors.push(`Too long — ${GROUP_NAME_MAX} characters at most.`);
  const bad = ILLEGAL_GROUP_NAME_CHARS.filter((c) => n.indexOf(c) !== -1);
  if (bad.length > 0) {
    errors.push(`SharePoint does not allow ${bad.join(" ")} in a group name.`);
  }
  const clash = (existingTitles ?? []).find((t) => norm(t).toLowerCase() === n.toLowerCase());
  if (clash !== undefined) errors.push(`"${clash}" already exists on this site.`);
  return errors;
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
