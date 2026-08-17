import {
  policyForPage,
  isRoleEligibleForPage,
  VIEW_ONLY_ROLES,
  ACTION_ROLES,
} from "./pageAccessPolicy";
import { GroupMapRole, SELECTABLE_ROLES } from "./groupMapModel";

describe("policyForPage — the client's four rules (2026-08-05)", () => {
  // Was uploaders-only until 2026-08-17. The client's rule is that a Head of Unit uploads as
  // well as approves ("basically apr can upload"), and an HoU group mapped before the
  // 2026-08-15 persona correction carries no UPL row — so keying the page on UPL alone denied
  // the upload form to every Head of Unit already provisioned.
  it("offers Upload-Form to uploaders AND approvers", () => {
    const p = policyForPage("Upload-Form.aspx");
    expect(p.roles).toEqual(["UPL", "APR"]);
    expect(p.adminOnly).toBe(false);
  });

  it("still never offers Upload-Form to a view-only role", () => {
    // Widening to APR must not widen to the oversight roles: a C-Level or Head of Department
    // holds no permission in the approval library at all, so the page would open onto an empty
    // cascade and imply an ability they do not have.
    for (const role of VIEW_ONLY_ROLES) {
      expect(isRoleEligibleForPage("Upload-Form.aspx", role)).toBe(false);
    }
  });

  it("offers ApprovalDocument to approvers only", () => {
    const p = policyForPage("ApprovalDocument.aspx");
    expect(p.roles).toEqual(["APR"]);
    expect(p.adminOnly).toBe(false);
  });

  it("makes Bulk-Upload administrators only", () => {
    const p = policyForPage("Bulk-Upload.aspx");
    expect(p.adminOnly).toBe(true);
    expect(p.roles).toEqual([]);
  });

  it("makes Folder-Manager administrators only", () => {
    expect(policyForPage("Folder-Manager.aspx").adminOnly).toBe(true);
  });

  it("makes the settings and mapping pages administrators only", () => {
    expect(policyForPage("CRS-Settings.aspx").adminOnly).toBe(true);
    expect(policyForPage("CRS-Mapping.aspx").adminOnly).toBe(true);
    expect(policyForPage("DMS-Config.aspx").adminOnly).toBe(true);
  });
});

// The regression this file exists for. "bulk-upload" CONTAINS "upload", so a first-match rule
// list in the wrong order offers the admin-only bulk tool to every uploader group — the same
// prefix-collision trap as _DEL inside _DELS, and just as silent.
describe("rule ordering", () => {
  it("treats bulk upload as admin-only, never as an upload page", () => {
    for (const name of ["Bulk-Upload.aspx", "BulkUpload.aspx", "bulk upload.aspx"]) {
      expect(policyForPage(name).adminOnly).toBe(true);
      expect(policyForPage(name).roles).toEqual([]);
    }
  });

  it("still treats a plain upload page as an uploader page", () => {
    for (const name of ["Upload.aspx", "Upload-Form.aspx", "UploadForm.aspx"]) {
      expect(policyForPage(name).roles).toEqual(["UPL", "APR"]);
    }
  });
});

// My Submissions — an uploader's view of their own files.
// Spec: docs/superpowers/specs/2026-08-14-my-submissions-design.md §3 D3.
describe("the My Submissions page is for UPLOADERS only", () => {
  const NAMES = ["My-Submissions.aspx", "MySubmissions.aspx", "My-Uploads.aspx", "My-Files.aspx"];

  it("offers it to uploader groups", () => {
    for (const name of NAMES) expect(policyForPage(name).roles).toEqual(["UPL"]);
  });

  it("does NOT offer it to approvers or Staging-deleters", () => {
    // Without its own rule this page matches nothing and takes DEFAULT_POLICY, which is
    // [UPL, APR, DELS] — an uploader's own-files page offered to the people it is private from.
    for (const name of NAMES) {
      expect(isRoleEligibleForPage(name, "APR")).toBe(false);
      expect(isRoleEligibleForPage(name, "DELS")).toBe(false);
      expect(isRoleEligibleForPage(name, "UPL")).toBe(true);
    }
  });

  it("never offers it to a view-only role", () => {
    for (const role of VIEW_ONLY_ROLES) {
      expect(isRoleEligibleForPage("My-Submissions.aspx", role)).toBe(false);
    }
  });

  it("is NOT reached through the generic upload rule — the reason line differs", () => {
    // Ordering is what this asserts. Until 2026-08-17 both rules yielded exactly ["UPL"], so the
    // reason line was the ONLY thing that could tell them apart. It no longer is — the generic
    // upload rule now also lists APR — which makes the roles assertion above a second, harder
    // guard: if the generic rule were hit first, My Submissions would be offered to approvers,
    // the people it is private from. The reason check stays, because a wrong label is what
    // produces a wrong grant later even when the roles happen to match.
    expect(policyForPage("My-Submissions.aspx").reason).toContain("their own submissions");
    expect(policyForPage("Upload-Form.aspx").reason).toContain("where documents are submitted");
  });

  it("does not capture the ADMIN pages that also mention uploads", () => {
    // "Bulk-Upload" contains "upload" and must stay admin-only; the bulk rule precedes them both.
    expect(policyForPage("Bulk-Upload.aspx").adminOnly).toBe(true);
    expect(policyForPage("Bulk-Upload.aspx").roles).toEqual([]);
  });
});

describe("view-only roles are never offered a page", () => {
  it("excludes MEMBER, GLOBAL and SEGVIEW everywhere", () => {
    const pages = ["Upload-Form.aspx", "ApprovalDocument.aspx", "Bulk-Upload.aspx", "CollabHome.aspx"];
    for (const page of pages) {
      for (const role of VIEW_ONLY_ROLES) {
        expect(isRoleEligibleForPage(page, role)).toBe(false);
      }
    }
  });

  it("does not leak a view-only role in through the default policy", () => {
    // CollabHome matches no rule, so it takes the default — which must still be action roles.
    for (const role of VIEW_ONLY_ROLES) {
      expect(ACTION_ROLES.indexOf(role)).toBe(-1);
    }
    expect(policyForPage("CollabHome.aspx").roles).toEqual(ACTION_ROLES);
  });
});

describe("isRoleEligibleForPage", () => {
  it("allows the matching role and refuses the others", () => {
    // THE ASYMMETRY IS DELIBERATE and is the point of this case. Approvers reach the upload
    // form (2026-08-17 — a Head of Unit uploads as well as approves), but uploaders must never
    // reach the approval queue. Making these mirror each other would hand every PIC the power
    // to approve their own documents, which is the one thing the whole approval flow exists to
    // prevent.
    expect(isRoleEligibleForPage("Upload-Form.aspx", "UPL")).toBe(true);
    expect(isRoleEligibleForPage("Upload-Form.aspx", "APR")).toBe(true);
    expect(isRoleEligibleForPage("ApprovalDocument.aspx", "APR")).toBe(true);
    expect(isRoleEligibleForPage("ApprovalDocument.aspx", "UPL")).toBe(false);
  });

  it("refuses every role on an admin-only page", () => {
    // DERIVED, not hand-listed. The literal list this replaced had gone stale twice: it still
    // named the retired "HC" role after it was split into UPLHC/APRHC, and it had never been
    // updated for SHARE — so "every role" was quietly testing all but one of them, and the one it
    // missed is the widest in the system.
    const all: GroupMapRole[] = [...SELECTABLE_ROLES, "ENTRY", "SHARE"];
    for (const role of all) {
      expect(isRoleEligibleForPage("Bulk-Upload.aspx", role)).toBe(false);
    }
  });

  it("is safe on empty and unknown input", () => {
    expect(policyForPage("").roles).toEqual(ACTION_ROLES);
    expect(policyForPage(undefined as unknown as string).roles).toEqual(ACTION_ROLES);
    expect(isRoleEligibleForPage("", "UPL")).toBe(true);
  });
});

describe("the Requests page", () => {
  it("offers both audiences — uploaders raise, the Head of Unit decides", () => {
    expect(policyForPage("Requests.aspx").roles).toEqual(["UPL", "APR"]);
    expect(policyForPage("CRS-Requests.aspx").roles).toEqual(["UPL", "APR"]);
  });

  it("beats the approver rule, so a name carrying 'approval' still reaches uploaders", () => {
    // Ordering, pinned: on the /approv/ rule this page would list approver groups only, and the
    // people who RAISE requests could not open the page their own requests appear on.
    expect(policyForPage("Approval-Requests.aspx").roles).toEqual(["UPL", "APR"]);
  });

  it("does not disturb the approver's own page", () => {
    expect(policyForPage("ApprovalDocument.aspx").roles).toEqual(["APR"]);
  });

  it("is not an administrator tool — an uploader must be able to be granted it", () => {
    expect(policyForPage("Requests.aspx").adminOnly).toBe(false);
  });
});
