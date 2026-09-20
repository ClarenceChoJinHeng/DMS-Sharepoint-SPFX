// Which groups may be offered access to which page.
//
// The Page Access tab used to list every group on the site for every page, which is how an
// upload form ended up offering itself to viewer groups and an admin tool offering itself to
// everyone. The list is now filtered by what the page is FOR.
//
// This is a UI filter, not a security boundary — the boundary is the SharePoint grant itself.
// It exists to stop a wrong grant being made by accident, and it can be overridden in the tab
// for the case nobody anticipated. A filter that cannot be overridden becomes a reason to go
// around the tool entirely.
//
// SPFx-free and pure, so the page→role rules are unit-testable without a tenant.

import { GroupMapRole } from "./groupMapModel";

export interface PagePolicy {
  /**
   * Roles whose groups may be granted this page. EMPTY means administrators only: site owners
   * and site collection admins keep access automatically when inheritance is broken, so an
   * admin page needs no group at all.
   */
  roles: GroupMapRole[];
  /** One line for the UI, explaining who is eligible and why the rest are hidden. */
  reason: string;
  /** True when the page is for administrators only — the empty-roles case, named. */
  adminOnly: boolean;
}

/**
 * View-only roles, never offered any page.
 *
 * MEMBER, GLOBAL and SEGVIEW all mean "can read approved documents". None has business on an
 * upload form, an approval screen or an admin tool — and a reader who opened the upload form
 * could not upload anyway, holding no Staging permission. Listing them only invites a grant that
 * looks like it does something and does not.
 */
export const VIEW_ONLY_ROLES: GroupMapRole[] = ["MEMBER", "GLOBAL", "SEGVIEW"];

/** Roles that act on documents — the default for a page with no specific rule. */
export const ACTION_ROLES: GroupMapRole[] = ["UPL", "APR", "DELS"];
// APRHC joined every people-facing rule on 2026-08-24: the HC Head of Unit (`hou_hc`) holds APRHC
// INSTEAD of APR, so any rule keyed on APR alone would AccessDeny the exact person the HC vertical
// exists for — the fourth instance of "a page keyed on a role its audience does not literally hold"
// (upload form → APR, Requests → DEPTVIEW then UPL/UPLHC again, My Submissions → UPLHC, and this).

/**
 * File-name rules, FIRST MATCH WINS, and the order is load-bearing.
 *
 * "bulk-upload" contains "upload", so the bulk rule must be tested before the upload rule or
 * every admin-only bulk page would be offered to uploaders. Exactly the trap that made the
 * role-suffix table sort longest-first; here the rules are few enough to order by hand, so the
 * ordering is pinned by a test instead.
 *
 * Matched loosely on the file name, because the same page ships under different names across
 * sites (Upload.aspx / Upload-Form.aspx / UploadForm.aspx).
 */
const RULES: Array<{ match: RegExp; policy: PagePolicy }> = [
  {
    // ⚠ THE VIEWER/C-LEVEL/HEAD-OF-DEPARTMENT PAGE (2026-09-21), added at the TOP so nothing later
    // in this array can accidentally claim it first.
    //
    // Client, correcting the same-day plan that put this feature on Requests.aspx for everyone:
    // "I dont want VIEWER, HOD, C LEVEL, or PIC to touch CRS-Requet.aspx... for safety purposes, I am
    // afraid they can by pass it." Three destinations resulted: File Permission (Requests.aspx)
    // stays Approver-exclusive; My Submissions gets PIC (unchanged, plus an "any document" fallback
    // for a file that is not one of their own); and Viewer/C-Level/HOD land HERE.
    //
    // ⚠ `UPL`/`UPLHC`/`APR`/`APRHC` ARE LISTED TOO, and that is not a mistake — SharePoint column
    // formatting cannot route by VIEWER ROLE, only by the row's own data, so the SAME link a PIC or
    // Approver clicks in a library also points here. This page's own component checks the viewer's
    // role on load and silently redirects a PIC to My Submissions and an Approver to File Permission
    // BEFORE rendering anything — but that redirect logic is React code, which cannot run at all if
    // the page-access check above it already refused the page. Listing them here only lets the page
    // LOAD long enough to bounce them; they see no content on it either way.
    //
    // HOD is NOT redirected. HOD's own SHARE grant (Manage Permissions on their department's
    // approved documents) is a real, standing SharePoint right, unrelated to the Share-REQUEST queue
    // that used to live on Requests.aspx — client, correcting a wrong assumption: "HOD can still
    // share, HOD stays in that page to share no redirection." That request-DECIDING queue (approving
    // someone else's raised Share request) is what moved off this page's audience entirely, with no
    // replacement built yet.
    // Spec: docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md
    match: /document.?view|file.?view/i,
    policy: {
      roles: [
        "UPL",
        "UPLHC",
        "APR",
        "APRHC",
        "MEMBER",
        "MEMBERHC",
        "GLOBAL",
        "SEGVIEW",
        "DEPTVIEW",
      ],
      adminOnly: false,
      reason:
        "Every role that can browse Documents/HC Documents/Archive/HC Archive is listed — PIC and " +
        "Approver are redirected to their own page on load, so this is really for Viewer, C-Level " +
        "and Head of Department.",
    },
  },
  {
    /* ⚠ NO LONGER ADMIN-ONLY (2026-08-22, client: "Bulk upload is now allowed for all uploaders to be
       used, client doesnt want admin to do the job"). Spec
       `2026-08-22-bulk-upload-for-uploaders-design.md`.

       The SAME three roles as the upload form, and for the same two reasons that were learned there:
       `hou` and `pic_hc` carry no literal `UPL` — `UPLHC` is a superset that `LIBRARY_ROLES.Staging`
       lists on the normal approval library too — and a Head of Unit group mapped before the
       2026-08-15 persona correction holds only `APR` and `DELS`, so keying on the upload roles alone
       would strand every already-provisioned HoU.

       Widening this cannot grant anyone a new place to file. The page grant opens the FORM; the
       folder ACL decides what can be written, and Bulk Upload now probes `AddListItems` on the
       destination exactly as the upload form does. A role listed here that cannot write sees an
       empty cascade — the safe direction.

       This rule must STILL be tested before /upload/i: "bulk-upload" contains "upload". Pinned by
       test, as is the fact that it is no longer adminOnly. */
    match: /bulk/i,
    policy: {
      roles: ["UPL", "UPLHC", "APR", "APRHC"],
      adminOnly: false,
      reason: "Bulk upload files historical documents that were already approved elsewhere, so uploaders and their Head of Unit reach it.",
    },
  },
  {
    // FIVE ADMIN PAGES WERE NOT MATCHED BY THIS RULE UNTIL 2026-08-17, and every one of them fell
    // through to a policy that offered it to ordinary users. Found by the coverage test added with
    // the admin-page lockdown, not by inspection:
    //
    //   Group-Management.aspx      — "group.?manager" requires "manager"; the page is "Management".
    //   Site-Access.aspx           — matched nothing, took DEFAULT_POLICY (UPL, APR, DELS).
    //   Page-Access.aspx           — same.
    //   CRS-Audit-Log.aspx         — same. Its web part has an IsSiteAdmin gate; the page grant did not.
    //   Approval-Library-Access    — WORSE: matched /approv/i below, so the screen that grants library
    //                                permissions was classified as an approver page.
    //
    // This mattered little while page restriction was a manual step nobody performed. It matters
    // now: reconciliation locks exactly the pages this rule marks adminOnly, so a name that does not
    // match is not a mislabel — it is a page left open, silently, for ever.
    //
    // `group.?manag` not `group.?manager`, so Manager and Management both match. `access` covers the
    // four access screens; Folder-Access already matched via "folder", and keeping "access" general
    // means a fifth access screen is covered on the day it is created rather than the day someone
    // notices. Checked against every non-admin page name: Upload-Form, Approval-Document,
    // My-Submissions, CRS-Requests and Home match none of these tokens.
    match: /folder|group.?manag|config|setting|mapping|admin|access|audit/i,
    policy: {
      roles: [],
      adminOnly: true,
      reason: "This is an administrator tool. Site owners and site collection administrators keep access automatically — no group needs to be added.",
    },
  },
  {
    // BEFORE the approver rule, deliberately: a name like "Approval-Requests.aspx" would otherwise
    // land on /approv/i, and this page's audience is wider than that rule's by one role. It would
    // otherwise fall to DEFAULT_POLICY, which is APR + UPL + DELS; explicit and narrow beats
    // right-by-accident.
    //
    // ⚠⚠ SUPERSEDES A ONE-DAY-OLD CHANGE (2026-09-20 -> 2026-09-21). For less than a day this rule
    // ALSO listed `DEPTVIEW`/`UPL`/`UPLHC`, because the ?file=<UniqueId> feature was first built to
    // reach everyone through THIS page. The client corrected it the same day it shipped: "I dont
    // want VIEWER, HOD, C LEVEL, or PIC to touch CRS-Requet.aspx... for safety purposes, I am afraid
    // they can by pass it." So this page is now Approver-EXCLUSIVE again — File Permission is the
    // request-deciding queue AND, for Approver only, the ?file= read-only/direct-action view too.
    // PIC uses My Submissions for that same feature (with a new "any document, not just mine"
    // fallback); Viewer/C-Level/HOD land on a brand-new, separate page. See the very first RULES
    // entry in this file for that page's own reasoning.
    // Spec: docs/superpowers/specs/2026-09-20-direct-share-delete-from-libraries-design.md
    //
    // ⚠ THIS ALSO MEANS HEAD OF DEPARTMENT NO LONGER OPENS THIS PAGE AT ALL — a real, deliberate loss
    // of their ability to DECIDE another person's raised Share request (the queue), confirmed
    // explicitly by the client. NOT a loss of HOD's own DIRECT share right (Manage Permissions on
    // their department's approved documents) — that is unrelated to this page and lives on, via the
    // new Viewer/C-Level/HOD page instead: "HOD can still share, HOD stays in that page to share no
    // redirection." Nowhere currently rebuilds the request-deciding queue for HOD; that is explicitly
    // deferred, not silently dropped.
    match: /request/i,
    policy: {
      roles: ["APR", "APRHC"],
      adminOnly: false,
      reason: "Only approver groups are listed — this page is the deletion/share request queue, and (for approvers) the read-only detail view reached by clicking a file directly in a library.",
    },
  },
  {
    match: /approv/i,
    policy: {
      roles: ["APR", "APRHC"],
      adminOnly: false,
      reason: "Only approver groups are listed — this page is where pending documents are approved.",
    },
  },
  {
    // BEFORE the generic upload rule, and it must stay there. "My-Submissions.aspx" matches no
    // other rule and would otherwise take DEFAULT_POLICY — UPL, APR and DELS — offering an
    // uploader's own-files page to approvers and Staging-deleters. Naming it "My-Uploads.aspx"
    // would land on UPL by luck, through a rule whose reason says the page is where documents are
    // submitted — untrue of a page that only shows what already was.
    // Spec: docs/superpowers/specs/2026-08-14-my-submissions-design.md §3 D3.
    match: /submission|my.?upload|my.?file/i,
    policy: {
      /* ⚠ `UPLHC` AND `APR` ADDED 2026-08-21 — WITHOUT THEM A HEAD OF UNIT COULD UPLOAD AND THEN NOT
         SEE WHAT THEY HAD UPLOADED.
         `hou` is `APR, DELS, DEL, SHARE, UPLHC, DELSHC` — it carries **no literal `UPL`**, because
         `UPLHC` is a superset and `LIBRARY_ROLES.Staging` lists it on the NORMAL approval library too.
         So a HoU holds CRS Upload on Approval Document and reaches the upload form (via APR), files a
         document, and was then denied the only page that lists it — and with it the only route to a
         deletion or share request about their own file. Third page keyed on a role its intended
         audience does not literally hold (upload form → APR, 2026-08-17; Requests → DEPTVIEW, and
         this).
         `APR` as well as `UPLHC`, for the same reason the upload form lists it: a HoU group mapped
         before the 2026-08-15 persona correction carries `APR` + `DELS` and no upload role at all, so
         keying on the upload roles alone would strand every already-provisioned Head of Unit.
         ⚠ WIDENING THIS CANNOT EXPOSE ANYONE'S FILES TO ANYONE ELSE. The page reads both libraries
         `AuthorId eq <me>`, so every role listed here sees only their OWN submissions — the grant
         opens the page, and authorship decides the rows. The earlier comment here claimed approvers
         were "the people it is private from", which was wrong: it is private from them for everyone
         else's files whether or not they can open it. */
      roles: ["UPL", "UPLHC", "APR", "APRHC"],
      adminOnly: false,
      reason: "Uploader and Head of Unit groups are listed — this page shows a person their own submissions, and a Head of Unit uploads too.",
    },
  },
  {
    // APPROVERS BELONG HERE TOO (2026-08-17, client: "client wants HOU to be able to get into
    // upload form to upload, basically apr can upload").
    //
    // A Head of Unit uploads as well as approves — the `hou` persona has carried UPL since the
    // 2026-08-15 role correction. But a group mapped BEFORE that correction holds only APR and
    // DELS rows, so keying this page on UPL alone locked out every Head of Unit provisioned
    // earlier. Found live 2026-08-17: an APR account got AccessDenied on the upload form while
    // the run log showed the page granted to the UPL group alone.
    //
    // Listing APR fixes both cases with no data migration — the older HoU groups need no new
    // Group Map row, and a pure approver group reaches the page too, which is what was asked.
    //
    // The page grant only opens the FORM. Whether anything can be filed from it is decided by
    // the folder ACL: the form probes AddListItems on the destination and shows the "not ready"
    // empty state where that fails. So a role listed here that cannot write sees an empty
    // cascade rather than an upload that fails at the end — the safe direction, and the reason
    // widening this list cannot grant anyone the ability to upload somewhere new.
    // ⚠ `UPLHC` ADDED 2026-08-31, AND ITS ABSENCE LOCKED EVERY HC-CLEARED PIC OUT OF THIS PAGE.
    // `pic_hc` is `["UPLHC"]` — no literal `UPL`, because `UPLHC` is a superset that
    // `LIBRARY_ROLES.Staging` lists on the NORMAL approval library too. So an HC uploader held no
    // qualifying role here: never granted the page, and REMOVED by the page pass if they ever were.
    // Found live — a guest in `GHO_GCA_GCBC_UPLOADER_HIGHLY_CONFIDENTIAL` got AccessDenied on
    // Upload-Form.aspx while the page's ACL, its publish state and the plain uploader group's Read
    // binding all read back perfectly correct.
    //
    // ⚠ FOURTH PAGE KEYED ON A ROLE ITS AUDIENCE DOES NOT LITERALLY HOLD (upload form needed `APR`
    // 2026-08-17, Requests needed `DEPTVIEW` 2026-08-21, My Submissions needed `UPLHC` 2026-08-21).
    // Both SIBLING rules in this same file — bulk upload and my-submissions — already listed
    // `UPLHC`; it was added to them when `APRHC` arrived on 2026-08-24 and missed here. **When a role
    // is added to one page rule, check every rule whose audience overlaps.**
    match: /upload/i,
    policy: {
      roles: ["UPL", "UPLHC", "APR", "APRHC"],
      adminOnly: false,
      reason: "Uploader and approver groups are listed — this page is where documents are submitted, and a Head of Unit uploads as well as approves.",
    },
  },
];

const DEFAULT_POLICY: PagePolicy = {
  roles: ACTION_ROLES,
  adminOnly: false,
  reason: "Uploader, approver and Staging-deleter groups are listed. View-only groups are never offered a page.",
};

/** The policy for a page, by file name. Unknown pages get the action-roles default. */
export function policyForPage(fileName: string): PagePolicy {
  const n = (fileName ?? "").trim().toLowerCase();
  for (const r of RULES) if (r.match.test(n)) return r.policy;
  return DEFAULT_POLICY;
}

/** Is this role offerable on this page? Always false for a view-only role. */
export function isRoleEligibleForPage(fileName: string, role: GroupMapRole): boolean {
  if (VIEW_ONLY_ROLES.indexOf(role) !== -1) return false;
  return policyForPage(fileName).roles.indexOf(role) !== -1;
}

/**
 * Did an explicit rule match this page name, or did it fall through to the default?
 *
 * `policyForPage` cannot answer this — it returns the same shape either way — and the distinction
 * decides whether reconciliation may DERIVE grants for the page. See
 * docs/superpowers/specs/2026-08-19-derived-page-access-design.md §2.1.
 *
 * Reads the same RULES array, so there is no second list of page names to drift. A hardcoded list
 * of file names inside FolderManager.tsx was the alternative, and its drift is silent: a page
 * renamed in one place and not the other either stops being granted or starts being locked, with
 * nothing reported.
 */
export function pageMatchedRule(fileName: string): boolean {
  const n = (fileName ?? "").trim().toLowerCase();
  for (const r of RULES) if (r.match.test(n)) return true;
  return false;
}

/**
 * Roles that reconciliation may grant this page AUTOMATICALLY, from the roles groups already hold
 * at folder scope. `[]` means derive nothing and leave the page alone.
 *
 * ⚠ EMPTY FOR AN UNMATCHED PAGE, and that is the whole safety of it. Granting at page scope
 * requires breaking the page's inheritance first, after which ONLY the listed groups can open it.
 * Deriving from DEFAULT_POLICY (UPL, APR, DELS) would therefore break the site home page,
 * CollabHome.aspx and anything the client authored, and lock out every role not in that list —
 * SDG Employee, Head of Department and C-Level hold none of them. Reconciliation would take the
 * site away from most of its users on a run that reported success.
 *
 * Empty for an adminOnly page too: those have `roles: []` already and belong to the lockdown pass,
 * which strips every non-Owners assignment. Two passes asserting opposite states on one page would
 * fight every run, for ever.
 *
 * View-only roles are filtered out for the same reason they are never OFFERED one: a reader who
 * opened the upload form could not upload anyway, holding no Staging permission.
 */
export function derivedRolesForPage(fileName: string): GroupMapRole[] {
  if (!pageMatchedRule(fileName)) return [];
  const p = policyForPage(fileName);
  if (p.adminOnly) return [];
  return p.roles.filter((r) => VIEW_ONLY_ROLES.indexOf(r) === -1);
}
