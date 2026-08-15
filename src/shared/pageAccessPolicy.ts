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
    match: /bulk/i,
    policy: {
      roles: [],
      adminOnly: true,
      reason: "Bulk upload is an administrator tool. Site owners and site collection administrators keep access automatically — no group needs to be added.",
    },
  },
  {
    match: /folder|group.?manager|config|setting|mapping|admin/i,
    policy: {
      roles: [],
      adminOnly: true,
      reason: "This is an administrator tool. Site owners and site collection administrators keep access automatically — no group needs to be added.",
    },
  },
  {
    // BEFORE the approver rule, deliberately. This page has TWO audiences — uploaders raise a
    // deletion or share request, the Head of Unit decides it — so landing on the APR-only rule (via a
    // name like "Approval-Requests.aspx") would leave the people who raise requests unable to open the
    // page their own requests are listed on. It would otherwise fall to DEFAULT_POLICY, which is the
    // same set plus DELS; explicit and narrow beats right-by-accident.
    // Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
    match: /request/i,
    policy: {
      roles: ["UPL", "APR"],
      adminOnly: false,
      reason: "Uploader and approver groups are listed — uploaders raise deletion and share requests here, and the Head of Unit decides them.",
    },
  },
  {
    match: /approv/i,
    policy: {
      roles: ["APR"],
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
      roles: ["UPL"],
      adminOnly: false,
      reason: "Only uploader groups are listed — this page shows a person their own submissions.",
    },
  },
  {
    match: /upload/i,
    policy: {
      roles: ["UPL"],
      adminOnly: false,
      reason: "Only uploader groups are listed — this page is where documents are submitted.",
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
