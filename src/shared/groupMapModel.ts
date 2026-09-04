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
  // APRHC RETIRED 2026-08-17, REVIVED 2026-08-24 AS A NAMING ROLE ONLY. The retirement (spec
  // 2026-08-17-hc-clearance-and-role-revision-design.md) followed the client's "any approver which
  // is HOU can see Highly Confidential files as well. So we do not need a dedicated HOU". On
  // 2026-08-24 they reversed the UPLOAD half of that: "we will need to make another highly
  // confidential group for HOU, so not every HOU can upload into Highly Confidential, it will be
  // the same pattern as PIC."
  //
  // A FULL GRANTING ROLE AGAIN, hours after being revived as naming-only, because the client
  // clarified the rule the same day: "a normal HOU cannot see HC Approval Document and HC Documents
  // … only HC HOU can see and approve and go into HC libraries." Seeing had to move with filing —
  // and it cannot move while plain APR sits on the HC rows of LIBRARY_ROLES, because a role held by
  // two personas cannot grant to one and withhold from the other. So APR left both HC rows and
  // APRHC carries the HC approver:
  //   • A SUPERSET, like UPLHC: it approves in BOTH approval libraries and reads both approved-side
  //     ones (downgraded to Read there, exactly as APR is on Documents). hou_hc holds it INSTEAD of
  //     APR — both would be two rows granting the same thing on the normal library.
  //   • It also names hou_hc's group (`_APR_HIGHLY_CONFIDENTIAL`): UPLHC already names pic_hc's,
  //     and two personas sharing a naming role plan one title — planBulkGroups' name-once guard
  //     would silently create only one of them.
  //   • ⚠ THE STAFFING RULE THIS CREATES: a unit that files HC documents MUST have someone in its
  //     HC Head-of-Unit group, or its HC uploads sit pending forever, visible only to their author,
  //     with no error anywhere. That is the direct meaning of the client's instruction, not a bug.
  | "APRHC"
  // DELHC and SHAREHC, 2026-08-24 — the HC twins of DEL and SHARE, following DELSHC's precedent
  // exactly: DEL and SHARE are held by the PLAIN Head of Unit (for Documents), so leaving them on
  // the DocumentsHC row is what kept a normal HoU inside HC Documents. The HC twins are listed
  // there instead, held by hou_hc and — the client's explicit choice, 2026-08-24 — by hod, whose
  // department-wide HC delete/share (2026-08-20) survives the split. A department-tier DELHC row
  // fans down with recon_departmentFanOut exactly as DEL does; the fan-out is role-generic.
  | "DELHC"
  | "SHAREHC"
  //
  // HC clearance is now a dedicated group for the two roles at the BOTTOM of the hierarchy only: the
  // uploader (UPLHC + DELSHC) and the plain viewer (MEMBERHC). Every management role reaches HC
  // through the role it already holds.
  //
  // DELSHC exists because DELS CANNOT be listed on the HC approval library: the plain PIC holds DELS
  // too, so listing it there gave every plain PIC CRS Delete on HC — a live leak until today. Draft
  // Item Security narrowed it (a non-approver sees only their own items, and a plain PIC has no HC
  // files) but an APPROVED HC file sitting there before Auto-route moves it is visible to every
  // reader, and they held delete on it.
  | "DELSHC"
  // MEMBERHC — the SDG Employee who may read HC. Named for the CLEARANCE, and it cannot reuse the
  // `_MEMBER_HC` suffix because MEMBER is the BASE group and carries no suffix at all.
  | "MEMBERHC"
  // DEPTVIEW, 2026-08-17. Head of Department became view-only ("HOD no need deletion power, he only
  // view" / "do not have share functionality that is HOU, just needs View"), so DEL and SHARE left
  // that persona.
  //
  // It needs its OWN role and cannot reuse MEMBER — the obvious move, and wrong. MEMBER is the SDG
  // Employee role and is deliberately absent from HC Documents, so reusing it would either lose
  // HoD's HC read, or, if MEMBER were added to HC Documents to restore it, hand every SDG Employee
  // HC read and destroy the dedicated HC viewer group in the same stroke. One role cannot grant HC
  // to one persona and withhold it from another.
  //
  // What SEGVIEW is to a segment, one tier down: plain Read, fanning DOWN to every unit beneath the
  // department, on Documents and HC Documents. Absent from both approval libraries, like every view
  // role — a viewer there would be reading other people's unapproved drafts.
  | "DEPTVIEW"
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
// DELSHC replaced APRHC here 2026-08-17. DEPTVIEW is deliberately ABSENT for the same reason GLOBAL
// and SEGVIEW are: it is a view role, and a viewer reaching an approval library would be reading
// other people's unapproved drafts.
export const STAGING_FACING_ROLES: GroupMapRole[] = ["UPL", "APR", "DELS", "UPLHC", "APRHC", "DELSHC"];

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
  // Was MISSING until 2026-08-18, so it rendered as the bare code "SHARE" on Folder Access while
  // every role beside it read as a sentence — the exact thing this table exists to prevent. It is a
  // real role: Head of Unit and both C-Level personas carry it, which makes a Head of Unit the
  // SHARING AUTHORITY for their unit rather than merely an approver of requests.
  SHARE:   "Share approved documents",
  DELS:    "Delete pending files",
  GLOBAL:  "C-Level — all segments",
  SEGVIEW: "C-Level — one segment",
  ENTRY:   "Library / page entry",
  // Named for the CLEARANCE, not the library, because these roles grant in the normal libraries too.
  // "HC uploader" would read as "uploads only to HC", which is the opposite of what they do.
  UPLHC:   "Uploader — Highly Confidential cleared",
  APRHC:   "Approver — Highly Confidential cleared",
  DELSHC:  "Delete pending files — Highly Confidential cleared",
  DELHC:   "Delete approved documents — Highly Confidential cleared",
  SHAREHC: "Share approved documents — Highly Confidential cleared",
  MEMBERHC: "View only — Highly Confidential cleared",
  // Named for the SCOPE, not "View only", because that is what distinguishes it from MEMBER — the
  // level is identical and only the reach differs. An admin choosing between two rows both reading
  // "View only" has nothing to go on.
  DEPTVIEW: "View only — whole department",
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
// 2026-08-17: APRHC removed (retired). DELSHC, MEMBERHC and DEPTVIEW added — each is a role an admin
// can legitimately author, and omitting one would make its persona unauthorable by hand.
export const SELECTABLE_ROLES: GroupMapRole[] =
  ["MEMBER", "UPL", "APR", "DEL", "DELS", "GLOBAL", "SEGVIEW", "DEPTVIEW", "UPLHC", "DELSHC", "MEMBERHC"];

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
  /**
   * The role whose suffix names this persona's group. DECLARED, never derived.
   *
   * "Its first role" is wrong, and wrong in the dangerous direction: `employee_hc` is
   * `["MEMBER", "MEMBERHC"]`, so first-role naming would call it `<unit>_EMPLOYEE`, which parses back as
   * plain `MEMBER` — presenting an HC-CLEARED VIEWER GROUP AS HAVING NO CLEARANCE. Silent, and exactly
   * what the suffix table exists to prevent.
   *
   * A test asserts every persona declares one and that each round-trips through `roleFromGroupName` back
   * to itself, so a new persona cannot be added without answering this.
   */
  namingRole: GroupMapRole;
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
    key: "clevel_global", namingRole: "GLOBAL", family: "C-Level", scope: "segment", label: "Global — view everything",
    /* VIEW-ONLY AGAIN SINCE 2026-08-20 (client: "C level's Share and Delete power is now suppose to
       be pass on to HOD, no more C level, C level is read only"). DEL and SHARE, added 2026-08-15
       and fanned down the segment on 2026-08-19, are GONE — they moved to `hod`.

       This narrows `segmentFanRoles()` to GLOBAL + SEGVIEW on its own, because that function is
       DERIVED from this array. Nothing in the reconciliation gate needs editing, which is the whole
       point of deriving it: a literal list would have gone on fanning DEL and SHARE from every
       segment-tier row while this file said otherwise.

       MIGRATION: an existing C-Level mapping must be RE-CREATED with this persona. Its rows still
       say DEL/SHARE and keep granting them across the whole segment — the change does not fail, it
       silently retains the power the client removed. Reconciliation must then be re-run. Same shape
       as the HoD note of 2026-08-17, now in reverse. */
    roles: ["GLOBAL"],
    summary: "Reads every segment, department and unit in the Documents library, down to Document Type. No Staging access, no upload, no approve.",
  },

  // One segment, not all twelve. GLOBAL narrowed by a term — see the SEGVIEW note in
  // ROLE_TO_PERMISSION. Documents only, for the same reason as GLOBAL: a C-Level on Staging
  // would be reading an entire segment's unapproved drafts.
  {
    key: "clevel_segment", namingRole: "SEGVIEW", family: "C-Level", scope: "segment", label: "Segment — view one segment",
    // As clevel_global, narrowed to one segment by the row's term — and view-only again since
    // 2026-08-20. See the note there; DEL and SHARE moved to `hod`.
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
  //
  // MEMBER DROPPED 2026-08-09, and unlike Head of Unit this one was redundant all along:
  // DEL maps to "CRS Delete", which is Read + Delete Items. The read was already there; the
  // second membership only made it look as though it were not.
  {
    key: "hod", namingRole: "DEPTVIEW", family: "Head of Department", scope: "department", label: "View, delete + share, department-wide",
    /* VIEW-ONLY SINCE 2026-08-17 (client: "HOD no need deletion power, he only view" and "Head of
       Department do not have share functionality that is HOU, just needs View"). Both DEL and SHARE
       left; deletion authority is now entirely the Head of Unit's, which is where the request
       workflow already put the performing half of it.

       DEPTVIEW, not MEMBER — see the role's own comment. MEMBER is the SDG Employee role and is
       absent from HC Documents, so reusing it here would either lose this persona's HC read or,
       if MEMBER were added there to restore it, hand every SDG Employee HC read.

       MIGRATION: an existing HoD mapping must be RE-CREATED with this persona. Its rows still say
       DEL/SHARE and will keep granting delete — the change does not fail, it silently retains the
       power the client removed. Reconciliation must then be re-run. */
    /* DEL AND SHARE RETURNED 2026-08-20 (client: "C level's Share and Delete power is now suppose
       to be pass on to HOD, no more C level, C level is read only"). This reverses the view-only
       rule of 2026-08-17 above, which is left in place as the record of why they went.

       ⚠ THE FAN-OUT IS WHAT MAKES THIS REACH ANYTHING. Every unit folder under the department has
       unique permissions, so a grant on the DEPARTMENT folder alone reaches no unit — the exact
       defect found in `clevel_segment` on 2026-08-19. A department-tier row fans to every unit
       below it, but ONLY while `recon_departmentFanOut` is on. It defaults on and must stay on;
       with it off, reconciliation REPORTS these two roles and grants neither, and a Head of
       Department silently ends up able to delete nothing.

       Not exempted from that switch the way SEGVIEW and the old C-Level roles were. The exemption
       rests on "the role name is the consent" — true for SEGVIEW, which granted nothing anywhere
       before 2026-08-07, and false here: department-tier DEL rows already exist on provisioned
       sites from the HoD persona as it stood before 2026-08-17, so a leftover row and a deliberate
       one are once again indistinguishable by tier. The switch is the gate that reads them.

       DEL is Documents-only and SHARE never touches an approval library (LIBRARY_ROLES), so this
       widens what a HoD can do with APPROVED documents and gives them nothing over drafts. */
    /* DELHC and SHAREHC added 2026-08-24, and they are how the HoD KEEPS a power rather than
       gaining one. DEL and SHARE left the DocumentsHC row that day — they are also held by the
       plain Head of Unit, who must no longer reach HC at all — which would silently have taken the
       HoD's department-wide HC delete and share (the client's 2026-08-20 grant) with it. The client
       confirmed on 2026-08-24 that management oversight of HC stays, so the HC twins land here too.
       Department-tier rows for them fan down with recon_departmentFanOut exactly as DEL does.

       ⚠ MIGRATION FOR HoD IS ADDITIVE, unlike the approver groups: the _HOD groups keep their name
       and their DEPTVIEW/DEL/SHARE rows, so re-running bulk provisioning maps the two NEW rows onto
       the EXISTING groups (name check first), then reconciliation grants them. No deletion. */
    roles: ["DEPTVIEW", "DEL", "SHARE", "DELHC", "SHAREHC"],
    summary: "Reads every unit under their department, in Documents and in HC Documents, and can delete or share an approved document anywhere in that department without asking — Highly Confidential ones included. Cannot upload, cannot approve, and has no access to either approval library.",
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
    key: "hou", namingRole: "APR", family: "Head of Unit", scope: "unit", label: "Approve, upload, delete + share own unit",
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
    /* NO HC ROLE OF ANY KIND SINCE 2026-08-24 — the client, twice in one day, each time narrower:
       first "not every HOU can upload into Highly Confidential, it will be the same pattern as
       PIC" (which removed UPLHC), then the clarification that decides this whole persona: "a
       normal HOU cannot see HC Approval Document and HC Documents … only HC HOU can see and
       approve and go into HC libraries."

       So: UPL (not UPLHC), APR (which no longer reaches any HC library — it left both HC rows of
       LIBRARY_ROLES the same day), and NO DELSHC — that role grants delete, and therefore read, on
       the HC approval library, which is exactly the door this persona must not hold a key to. A
       plain Head of Unit approves, deletes and shares NORMAL documents in their unit and touches
       nothing HC. HC work belongs to `hou_hc` below.

       ⚠ THE STAFFING RULE THIS CREATES, said plainly because nothing on screen will: a unit that
       files HC documents MUST have someone in its HC Head-of-Unit group, or its HC uploads sit
       pending forever — visible only to their author, with no error anywhere. That is the direct
       meaning of the client's instruction, chosen with the deadlock stated.

       ⚠ MIGRATION IS THE WHOLE JOB (the 2026-08-20 lesson): every _APPROVER group already
       provisioned carries UPLHC/APR rows and reconciliation has GRANTED them on the HC folders.
       Editing this array changes what NEW mappings write and nothing else — deleting a Group Map
       row does not revoke a folder grant. The verified repair path, per segment: export the Group
       Management CSV (members), DELETE every _APPROVER group, bulk provision, reconcile, re-add
       people. Until a unit's group is re-created, its HoU keeps every HC power this change
       removes. */
    roles: ["APR", "DELS", "DEL", "SHARE", "UPL"],
    // SUMMARY CORRECTED 2026-08-17. It still read "Cannot upload, cannot delete approved
    // documents" — written for the pre-2026-08-15 role set and never updated when UPL, DEL and
    // SHARE were added directly above. The persona picker therefore told an administrator, on
    // screen, the opposite of what the persona does, in the same words the `hod` entry uses
    // correctly. Cost real time on 2026-08-17: a Head of Unit could not upload, and this line
    // read as confirmation that they were never meant to. A roles array and its description
    // drifting apart is invisible to every test, because nothing asserts on prose.
    summary: "Approves every ORDINARY file in their own unit, uploads to it, and deletes pending or rejected files there. Reads the unit's approved documents and can delete or share them — the unit's sharing authority for ordinary documents. Touches nothing Highly Confidential: cannot see, approve or enter the HC libraries at all (that is the HC Head of Unit below). Sees no sibling unit.",
  },

  // ── Head of Unit, Highly Confidential ────────────────────────────────────────
  //
  // The client's rule, clarified twice on 2026-08-24 and stricter the second time: "a normal HOU
  // cannot see HC Approval Document and HC Documents … only HC HOU can see and approve and go into
  // HC libraries." So this persona is not merely hou-plus-filing — it is the ONLY Head of Unit that
  // exists inside the HC libraries at all.
  //
  // Its HC access rides on three HC-specific roles, because each plain counterpart is held by the
  // plain HoU (and DEL/SHARE by the HoD too) and a role held by two personas cannot grant to one
  // and withhold from the other:
  //   APRHC   — approve, both approval libraries (superset of APR, exactly as UPLHC is of UPL)
  //   DELHC   — delete approved HC documents (DocumentsHC; DEL stays Documents-only)
  //   SHAREHC — share approved HC documents (DocumentsHC; SHARE stays Documents-only)
  // DELSHC (pending-HC delete) moved here from the plain persona, whose HC access ended.
  //
  // namingRole APRHC, NOT UPLHC: UPLHC already names pic_hc's group, and planBulkGroups plans each
  // NAME once — two personas deriving one title means the second is silently never created.
  {
    key: "hou_hc", namingRole: "APRHC", family: "Head of Unit", scope: "unit",
    label: "Approve, upload, delete + share own unit — incl. Highly Confidential",
    /* Deliberately a SUPERSET of the plain persona's capabilities: an HC Head of Unit is still the
       unit's ordinary approver (APRHC covers the normal approval library; DELS/DEL/SHARE are the
       plain ones), so a unit needs ONE Head of Unit group per person, never both. */
    roles: ["APRHC", "DELS", "DEL", "SHARE", "UPLHC", "DELSHC", "DELHC", "SHAREHC"],
    summary: "Everything the plain Head of Unit can do, plus the Highly Confidential side: sees, approves, files, deletes and shares HC documents in their own unit. The ONLY Head of Unit role with any HC access — a unit that files HC documents must have someone in this group, or those uploads wait forever with nobody able to approve them. Give a Head of Unit THIS group or the plain one, never both.",
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
    key: "pic", namingRole: "UPL", family: "PIC", scope: "unit", label: "Upload only",
    // DELS added 2026-08-15, correcting the model: the client had said a PIC cannot delete, and the
    // rule is the opposite in the APPROVAL LIBRARY.
    //
    // A PIC deletes in the APPROVAL LIBRARY directly. In DOCUMENTS they delete nothing and share
    // nothing without permission: both are requests the Head of Unit decides.
    // See 2026-08-15-deletion-and-share-requests-design.md.
    //
    // DELS is confined to the approval library by LIBRARY_ROLES, and to the unit by the folder ACL.
    // It covers any pending or rejected file there, not only the PIC's own — since 2026-08-19, when
    // draft security was opened to "any user who can read items", peers can see each other's drafts.
    //
    // In DOCUMENTS a PIC still deletes nothing and shares nothing: both are requests the Head of
    // Unit approves. See 2026-08-15-deletion-and-share-requests-design.md.
    /* DELS REMOVED 2026-08-20 (client: "For Staging PIC should not be able to delete, they have to
       request from HOU"). This reverses the 2026-08-15 correction directly above, which is left as
       the record of why it was added.

       ⚠ A PIC NOW HAS NO ROUTE TO REMOVE A PENDING OR REJECTED FILE AT ALL. The requests workflow
       is raised from My Submissions on an APPROVED file only — precisely because a PIC used to hold
       delete in the approval library and needed no permission there. Removing DELS closes that door
       without opening the other one, so until requests cover pending files the answer is a Head of
       Unit deleting it for them, out of band. Stated here because nothing on screen says it.

       Draft Item Security is back to "only users who can approve items (and the author)" as of
       2026-08-20, so a PIC sees only their own drafts. That bounds what they would have deleted; it
       is not a substitute for the permission. */
    roles: ["UPL"],
    // SUMMARY CORRECTED 2026-08-17, same drift as `hou`: it read "Cannot approve or delete"
    // while DELS had been added directly above on 2026-08-15. "Cannot delete" is the sentence a
    // PIC would be shown to explain why they must raise a request — and it is wrong in the
    // library where they hold delete outright.
    summary: "Uploads to their unit at any confidentiality level. Reads the unit's approved documents but cannot approve; deleting anything — a pending file or an approved document — is a request the Head of Unit decides.",
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
    key: "pic_hc", namingRole: "UPLHC", family: "PIC", scope: "unit", label: "Upload only, incl. Highly Confidential",
    /* DELS → DELSHC, 2026-08-17, and this closed a LIVE LEAK rather than tidying a name.
       DELS is held by the PLAIN PIC too, so while DELS was listed on the HC approval library every
       plain PIC held CRS Delete there. Draft Item Security narrowed it — a non-approver sees only
       their own items, and a plain PIC has no HC files — but an APPROVED HC file sitting in that
       library before Auto-route moves it is visible to every reader, and they held delete on it.
       A role held by two personas cannot grant to one and withhold from the other, so the HC twin
       is its own role. See the spec §3.3. */
    /* DELSHC REMOVED 2026-08-20, following the plain PIC losing DELS the same day. The client's
       rule is about the ROLE, not the library: "For Staging PIC should not be able to delete, they
       have to request from HOU", and HC Approval Document is a staging-side library. Leaving it
       here would mean a cleared PIC could delete pending HC files while an uncleared one could
       delete nothing — clearance deciding a DELETE right it has never decided anywhere else.

       DELSHC now belongs to `hou` alone, exactly as DELS does. */
    roles: ["UPLHC"],
    summary: "Everything a PIC does, plus filing and reading Highly Confidential documents for this unit. A PIC without this clearance is never even shown the Highly Confidential level.",
  },

  // ── SDG Employee, Highly Confidential ──────────────────────────────────────
  //
  // New 2026-08-17, on the client's instruction: "there also should be a separate viewer group, the
  // SDG employee who can only see needs another group."
  //
  // Read on HC Documents and nothing else. It REPLACES the plain `employee` persona for a cleared
  // person rather than accompanying it: MEMBERHC is not a superset, so a cleared viewer needs the
  // plain MEMBER row as well to read the normal library — which is why this persona carries BOTH.
  // One person, one group, as everywhere else since 2026-08-09.
  //
  // Approval-side absent, like every view role: a viewer there would be reading unapproved drafts.
  {
    key: "employee_hc", namingRole: "MEMBERHC", family: "SDG Employee", scope: "unit", label: "View only, incl. Highly Confidential",
    roles: ["MEMBER", "MEMBERHC"],
    summary: "Reads their own unit's approved documents, including Highly Confidential ones. Cannot upload, approve, delete or share. An SDG Employee without this clearance cannot see Highly Confidential documents at all.",
  },

  // ── SDG Employee ───────────────────────────────────────────────────────────
  //
  // Now the ONLY persona whose whole purpose is MEMBER. Since 2026-08-09 every other persona
  // carries its own Documents read, so this one means exactly what its name says: someone who
  // views and does nothing else.
  { key: "employee", namingRole: "MEMBER", family: "SDG Employee", scope: "unit", label: "No power — view only", roles: ["MEMBER"], summary: "Reads their own unit's approved documents. Cannot upload or approve." },
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
  // The APRHC→APR alias (2026-08-17, while APRHC was retired) is GONE: APRHC is a real granting
  // role again since 2026-08-24, in SHORT_ROLE_CODES, so a short code wins outright and the alias
  // would be dead code that misleads. Request routing now matches APR OR APRHC explicitly in
  // Requests.tsx and MySubmissions.tsx — the pre-retirement design CLAUDE.md always described:
  // "a unit whose approver is the HC one has no plain APR row, and matching APR alone leaves that
  // queue permanently empty while requests pile up behind it."
  DEPARTMENTVIEW: "DEPTVIEW",
  VIEWER_HC: "MEMBERHC",
};

/** Every short code the permission tables key on. Listed, not derived — the union is a type. */
// APRHC is BACK here since 2026-08-24 (it was absent while retired, so its alias could fire).
const SHORT_ROLE_CODES = ["MEMBER", "UPL", "APR", "DEL", "DELS", "SEGVIEW", "UPLHC", "APRHC", "DELSHC", "DELHC", "SHAREHC", "MEMBERHC", "DEPTVIEW", "SHARE", "GLOBAL", "ENTRY"];

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
  // `_C_LEVEL` is the current spelling (2026-09-02); `_SEGVIEW` is LEGACY and stays so an
  // already-provisioned `GHO_SEGVIEW` keeps parsing rather than falling through to MEMBER. Same
  // additive pattern as `_HOD`/`_DEPARTMENT_VIEWER` and `_VIEWER`/`_EMPLOYEE` above and below.
  { suffix: "_C_LEVEL", role: "SEGVIEW" },
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
  { suffix: "_UPL_HC", role: "UPLHC" },
  { suffix: "_DELS_HIGHLY_CONFIDENTIAL", role: "DELSHC" },
  { suffix: "_DELS_HC", role: "DELSHC" },
  // MEMBERHC cannot use `_MEMBER_HC`: MEMBER is the BASE group and carries no suffix at all
  // (suffixForRole returns "" for it), so there is no `_MEMBER` to extend. `_VIEWER_*` instead,
  // matching the ROLE_LABEL wording an admin sees.
  { suffix: "_VIEWER_HIGHLY_CONFIDENTIAL", role: "MEMBERHC" },
  // Added 2026-08-18. `_EMPLOYEE` gives the base group a parseable name; `_HOD` is the client's
  // preferred spelling for Head of Department. Both are ADDITIVE — the older `_DEPARTMENT_VIEWER`
  // and `_DEPTVIEW` below still parse, so no existing group is stranded.
  { suffix: "_EMPLOYEE", role: "MEMBER" },
  { suffix: "_HOD", role: "DEPTVIEW" },
  { suffix: "_VIEWER_HC", role: "MEMBERHC" },
  { suffix: "_DEPARTMENT_VIEWER", role: "DEPTVIEW" },
  { suffix: "_DEPTVIEW", role: "DEPTVIEW" },
  // `_APR_HIGHLY_CONFIDENTIAL` RESOLVES TO APRHC AGAIN (2026-08-24) — it is `hou_hc`'s naming
  // suffix, and the namingRole round-trip (suffixForRole → roleFromGroupName) must return the role
  // it started from. Between 2026-08-17 and today it mapped to plain APR, because APRHC was retired
  // and any group so named simply meant "approver"; with the HC Head-of-Unit persona back, a name
  // carrying this suffix means exactly that persona. Parsing to APRHC stays harmless for rows —
  // normalizeRoleValue aliases APRHC to APR, and ROLE_TO_PERMISSION carries no APRHC entry — while
  // deleting the suffixes would make such a name fall through to MEMBER: an approver group silently
  // reclassified as view-only, invisible in the mapping list.
  { suffix: "_APR_HIGHLY_CONFIDENTIAL", role: "APRHC" },
  { suffix: "_APR_HC", role: "APRHC" },
  /* STANDARDISED SPELLINGS, 2026-08-26 (client, looking at Group Management): the HC variants
     abbreviated where the plain ones spelled out — `..._APPROVER` beside `..._APR_HIGHLY_CONFIDENTIAL`
     — and the plain viewer said EMPLOYEE while the HC viewer already said VIEWER. These three are
     what NEW names use; every spelling above stays so no group already named is stranded, exactly as
     `_DEPARTMENT_VIEWER` and `_DEPTVIEW` were kept when `_HOD` arrived.

     ⚠ THE LENGTH SORT BELOW IS WHAT KEEPS THESE APART, and it holds:
       `_APPROVER_HIGHLY_CONFIDENTIAL` (29) is tested before `_APPROVER` (9)
       `_UPLOADER_HIGHLY_CONFIDENTIAL` (29) before `_UPLOADER` (9)
       `_VIEWER_HIGHLY_CONFIDENTIAL` (27) and `_VIEWER_HC` (10) before `_VIEWER` (7)
     A plain `_VIEWER` can never swallow the HC spellings anyway — matching is on the END of the
     name, and those end in `_CONFIDENTIAL` / `_HC`. */
  { suffix: "_APPROVER_HIGHLY_CONFIDENTIAL", role: "APRHC" },
  { suffix: "_UPLOADER_HIGHLY_CONFIDENTIAL", role: "UPLHC" },
  { suffix: "_VIEWER", role: "MEMBER" },
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
/**
 * The ONE suffix written into a new name per role, stated explicitly.
 *
 * Derived-by-search was wrong and broke on 2026-08-17: ROLE_SUFFIXES is sorted LONGEST-FIRST (the
 * length sort is load-bearing for PARSING), and once `_APR_HIGHLY_CONFIDENTIAL` was re-pointed at
 * APR, a search for "the first suffix whose role is APR" returned that 24-character legacy spelling
 * instead of `_APPROVER`. Every new approver group would have been suggested with an HC name.
 *
 * Reading order out of a table sorted for a different purpose is the bug; this map is the fix.
 * Several spellings are ACCEPTED for parsing, exactly one is WRITTEN.
 */
const CANONICAL_SUFFIX: Record<string, string> = {
  UPL: "_UPLOADER",
  APR: "_APPROVER",
  DEL: "_DELETER_DOCUMENTS",
  DELS: "_DELETER_STAGING",
  // `_C_LEVEL`, not `_SEGVIEW` (client, 2026-09-02: "Group naming for C-level should not be called
  // Segview but rather GHO_C_LEVEL"). "SEGVIEW" was always the internal ROLE code, never meant to be
  // read by the client — this is the group NAME suffix only, and it now says what the group is for
  // rather than the code that decides its permissions. The old suffix stays in ROLE_SUFFIXES below so
  // an already-provisioned `GHO_SEGVIEW` still parses; only NEW names use `_C_LEVEL`.
  SEGVIEW: "_C_LEVEL",
  // `_HOD`, not `_DEPARTMENT_VIEWER` (client, 2026-08-18). Note HOU is Head of UNIT — a department
  // group must not borrow it. Both older spellings stay in ROLE_SUFFIXES so groups already named
  // that way keep parsing; only new names use this one.
  DEPTVIEW: "_HOD",
  // SPELLED OUT since 2026-08-26, to match `_UPLOADER` / `_APPROVER` on the plain groups. The
  // abbreviated forms still PARSE (see ROLE_SUFFIXES), so existing groups are untouched.
  UPLHC: "_UPLOADER_HIGHLY_CONFIDENTIAL",
  // hou_hc's name (2026-08-24). Written for that persona's groups alone — APRHC appears in no
  // roles array and no LIBRARY_ROLES row, so nothing else ever asks for this suffix.
  APRHC: "_APPROVER_HIGHLY_CONFIDENTIAL",
  DELSHC: "_DELS_HIGHLY_CONFIDENTIAL",
  MEMBERHC: "_VIEWER_HIGHLY_CONFIDENTIAL",
  /**
   * MEMBER now carries a suffix, and this closes a hole rather than tidying one.
   *
   * The base group used to be bare (`GHO_GF_TAX`), and `roleFromGroupName` FALLS THROUGH to MEMBER for
   * any name it does not recognise — so a bare name and a typo were the same answer, which is how a
   * mistyped approver group reads as view-only. With every generated name carrying a suffix, that
   * fallback can later be made loud. (Not changed here: it would reclassify groups on existing sites.)
   */
  // `_VIEWER`, not `_EMPLOYEE` (client, 2026-08-26). The HC counterpart was already
  // `_VIEWER_HIGHLY_CONFIDENTIAL`, so `_EMPLOYEE` was the odd one out — the same role reading as two
  // different words depending on clearance. `_EMPLOYEE` still parses, so groups already named that
  // way keep resolving to MEMBER rather than falling through to it by accident.
  MEMBER: "_VIEWER",
};

/**
 * The role whose suffix names a persona's group, or "" for an unknown persona.
 *
 * A lookup rather than a second table: `Persona.namingRole` is declared beside the roles it names, so the
 * two cannot drift. "" is returned for an unknown key so a caller falls back to an unsuffixed name rather
 * than throwing mid-form — the name is a suggestion, and a missing suffix is visible while an exception is
 * a blank screen.
 */
export function namingRoleFor(personaKey: string): GroupMapRole | "" {
  const p = PERSONAS.filter((x) => x.key === personaKey)[0];
  return p ? p.namingRole : "";
}

export function suffixForRole(role: GroupMapRole | ""): string {
  // GLOBAL alone stays suffix-less: `suggestGroupName` returns the literal "GLOBAL" for it, so there is
  // no <seg>_<tier> stem for a suffix to hang off. MEMBER was in this guard until 2026-08-18 and is now
  // `_EMPLOYEE` — see the note on CANONICAL_SUFFIX.MEMBER for why a bare base-group name was a hole.
  if (role === "GLOBAL" || role === "") return "";
  const canonical = CANONICAL_SUFFIX[role];
  if (canonical !== undefined) return canonical;
  for (const s of ROLE_SUFFIXES) if (s.role === role) return s.suffix;
  return `_${role}`;
}

/**
 * The canonical name for an EXISTING group, or `undefined` if it should be left alone.
 *
 * Client, 2026-08-27: *"you forgot to rename the APR to APPROVAL and UPL to UPLOADER in Groups on
 * this site."* 1.0.261.0 changed what NEW names are WRITTEN and deliberately left existing groups
 * untouched, so a provisioned site ends up MIXED - `GHO_GCA_EG_APR_HIGHLY_CONFIDENTIAL` beside
 * `GHO_GCA_EG_VIEWER_HIGHLY_CONFIDENTIAL`. This is the other half: same role, same stem, canonical
 * suffix.
 *
 * WARN: IT ACTS ONLY ON A NAME WHOSE SUFFIX WAS ACTUALLY MATCHED, NEVER ON THE FALLBACK, and that
 * is the whole safety of it. `roleFromGroupName` FALLS THROUGH TO MEMBER for anything unrecognised
 * - so driving a rename from it would rewrite every hand-named group on the site into a
 * `_VIEWER`, `CRS_SITE_MEMBERS` included. Matching the suffix list directly is what distinguishes
 * "this is an uploader group spelled the old way" from "this is not one of our groups at all".
 *
 * Returns `undefined` for: an unmatched name, a name already canonical, and a stem that would be
 * left empty (a group called exactly `_UPLOADER` has nothing to rename).
 *
 * SAFE TO APPLY: a SharePoint group rename PRESERVES ITS ID, and every Group Map row is keyed on
 * `GroupId`, so mappings, grants and folder ACLs all survive. The stored `GroupName` on those rows
 * goes stale - it is a label, and Folder Access shows it - which is why this is worth doing in one
 * pass rather than leaving a site half-renamed for ever.
 */
export function canonicalGroupRename(name: string): string | undefined {
  const raw = norm(name);
  if (raw.length === 0) return undefined;
  const upper = raw.toUpperCase();
  // The global C-Level group is a LITERAL, not a stem+suffix, so it is handled directly rather than
  // through the suffix table below. "GLOBAL" is the pre-2026-09-02 name and IS offered a rename, to
  // the current literal; a group already named "C_LEVEL_GLOBAL" has nothing left to rename.
  if (upper === "GLOBAL") return "C_LEVEL_GLOBAL";
  if (upper === "C_LEVEL_GLOBAL") return undefined;
  // ROLE_SUFFIXES is sorted LONGEST-FIRST, so `_UPL_HIGHLY_CONFIDENTIAL` cannot be read as `_UPL`.
  for (const s of ROLE_SUFFIXES) {
    const suffix = s.suffix.toUpperCase();
    if (suffix.length === 0 || upper.length <= suffix.length) continue;
    if (upper.slice(upper.length - suffix.length) !== suffix) continue;
    const stem = raw.slice(0, raw.length - suffix.length);
    if (stem.length === 0) return undefined;
    const canonical = CANONICAL_SUFFIX[s.role];
    if (canonical === undefined || canonical.length === 0) return undefined;
    const next = `${stem}${canonical}`;
    return next.toUpperCase() === upper ? undefined : next;
  }
  // No recognised suffix - NOT ours to rename. See the warning above.
  return undefined;
}

/**
 * Derive the intended role from a group name's suffix (naming convention:
 * <seg>_<dept>_<unit>[_UPLOADER|_APPROVER|…]). Case-insensitive, trims.
 *
 * Any name with no recognised suffix — including a base viewer group — is MEMBER.
 */
export function roleFromGroupName(name: string): GroupMapRole {
  const n = norm(name).toUpperCase();
  // The global C-Level group is named "C_LEVEL_GLOBAL" outright (2026-09-02; "GLOBAL" bare is the
  // LEGACY name and still parses) — no stem, so no suffix to match. Without this it fell through to
  // MEMBER, meaning the WIDEST grant in the model parsed back as the narrowest. Exact match only: a
  // group merely ENDING in either word would be a different group.
  if (n === "GLOBAL" || n === "C_LEVEL_GLOBAL") return "GLOBAL";
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
/**
 * Characters SharePoint refuses in a GROUP name.
 *
 * ⚠ THIS LIST WAS WRONG UNTIL 2026-08-19, and the gap was a COMMA. It held SharePoint's
 * FILE/FOLDER illegal set — `" # % & * : < > ? \ / { | } ~` — which omits `, + = ; ' [ ]`. So
 * `validateGroupName` passed a name SharePoint then refused, and bulk provisioning reported fifteen
 * individual HTTP 500s that each looked like a one-off:
 *
 *   The group name is empty, or you are using one or more of the following invalid characters:
 *   " / \ [ ] : | < > + = ; , ? * '
 *
 * Found by a term called `HR Rewards, Services, and Performance`. Invisible with short codes —
 * `GHO_GF_TAX_APPROVER` has nothing to strip — so it only appears through the **"Same as term
 * name"** button, where the folder code IS the term's label, punctuation and all.
 *
 * The union of both sets, deliberately: SharePoint's message is the authority for what it refuses,
 * and the extra file-name characters have never been wanted in a group name either. Being stricter
 * on a NEW name costs nothing; being lax produces a 500 the admin cannot explain from the screen.
 */
const ILLEGAL_GROUP_NAME_CHARS = [
  '"', "#", "%", "&", "*", ":", "<", ">", "?", "\\", "/", "{", "|", "}", "~",
  ",", "+", "=", ";", "'", "[", "]",
];

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
/**
 * Characters SharePoint refuses in a GROUP name.
 *
 * ⚠ THIS IS A DIFFERENT SET FROM THE FOLDER ONE, and that difference is the bug (found live
 * 2026-08-19). `sanitizeFolderSegment` cleans an abbreviation for use as a FOLDER name; the group
 * name is then derived from that same abbreviation chain and was never cleaned again. A COMMA is
 * legal in a folder name and illegal in a group name, so a term called
 * `HR Rewards, Services, and Performance` produced a folder happily and then failed group creation
 * with HTTP 500 — fifteen groups on one segment, each reported individually and each looking like a
 * one-off.
 *
 * Invisible with short codes, which is why it took this long: `GHO_GF_TAX_APPROVER` has nothing to
 * strip. It appears only through the **"Same as term name"** button — a supported feature, and the
 * client's own idea — where the folder code IS the term's label, punctuation and all.
 *
 * Verbatim from SharePoint's own error: `" / \ [ ] : | < > + = ; , ? * '`
 */
/**
 * Clean one part of a group name.
 *
 * REMOVES rather than substitutes, matching `sanitizeFolderSegment`. Substituting an underscore
 * would be worse here than in a folder name: `_` is the SEPARATOR this convention is built on, so
 * `Health, Safety` becoming `Health__Safety` would read as an extra empty tier to anybody parsing
 * the name by eye. Runs of whitespace left behind are collapsed, because `Health  Safety` is just
 * untidy.
 */
export function sanitizeGroupNameSegment(raw: string): string {
  let out = raw ?? "";
  for (const ch of ILLEGAL_GROUP_NAME_CHARS) out = out.split(ch).join("");
  return out.replace(/\s+/g, " ").trim();
}

export function suggestGroupName(
  segmentLabel: string,
  tierLabels: string[],
  role: GroupMapRole | "",
): string {
  // "C_LEVEL_GLOBAL" since 2026-09-02, replacing the bare "GLOBAL" (client: renaming SEGVIEW's
  // suffix to `_C_LEVEL` at the same time made the two read as unrelated names for the same family).
  // Still bare — no segment stem — because this ONE group is site-wide by definition; a segment
  // prefix would misstate what it grants. See `roleFromGroupName`/`canonicalGroupRename` below for
  // the matching parse and rename rules.
  if (role === "GLOBAL") return "C_LEVEL_GLOBAL";
  const parts = [norm(segmentLabel), ...tierLabels.map(norm)]
    .map(sanitizeGroupNameSegment)
    .filter(Boolean);
  // MEMBER is the base group and carries NO suffix — that is the convention
  // roleFromGroupName reads back, so adding one here would make every base group
  // parse as an unknown role. An unpicked role ("") is the same: the admin is
  // still choosing, and a trailing "GHO_" is not a name to hand them.
  return parts.join("_") + suffixForRole(role);
}

/**
 * The roles a SEGMENT-tier Group Map row is allowed to fan down onto every folder beneath it.
 *
 * ⚠ DERIVED FROM THE PERSONAS, never hand-listed. The rule it encodes is exactly the meaning of
 * `scope: "segment"`: a persona scoped to a whole business segment must reach the whole business
 * segment. A literal list here would be a second definition of a C-Level's powers, free to drift
 * from `PERSONAS` — and the drifting copy would be this one, because it is read by reconciliation
 * and never by a screen anybody looks at.
 *
 * Until 2026-08-19 only `SEGVIEW` fanned down, so `clevel_segment` (`SEGVIEW + DEL + SHARE`) landed
 * its delete and share on the SEGMENT folder alone. Every folder below has broken inheritance, so
 * those two permissions reached **nothing**: the persona picker offered *"view, delete + share one
 * segment"* and granted view. The client's decision, 2026-08-19: *"allow Clevel to delete and
 * share."*
 *
 * ⚠ THE SAFETY IS THE TIER, NOT THE ROLE. `DEL` and `SHARE` are also held by `hou`, whose rows sit
 * at the UNIT tier and are therefore never fanned. Only a C-Level persona puts either role on a
 * segment-tier row, and neither was ever auto-created there — so such a row cannot be the
 * pre-2026-07-29 leftover the fan-out gate exists to refuse. Callers must check the row's tier
 * before consulting this list; using it on a department-tier row would re-open exactly that hole.
 */
export function segmentFanRoles(): GroupMapRole[] {
  const out: GroupMapRole[] = [];
  for (const p of PERSONAS) {
    if (p.scope !== "segment") continue;
    for (const r of p.roles) if (out.indexOf(r) === -1) out.push(r);
  }
  return out;
}

/**
 * Convenience for the reconciliation gate: may this role fan down from a segment-tier row?
 *
 * Takes a plain `string`, not `GroupMapRole`, because a stored Group Map row's `Role` is list text
 * and can hold a value outside the union — a hand-edited row, or one written by an older build. An
 * unrecognised role answers **false**, which refuses the fan-down. That is the safe direction: the
 * cost is a C-Level missing a grant an admin can see refused in the log, against a wide delete
 * right handed out on a value nothing in this file defines.
 */
export function fansFromSegmentTier(role: string): boolean {
  return segmentFanRoles().indexOf(role as GroupMapRole) > -1;
}
