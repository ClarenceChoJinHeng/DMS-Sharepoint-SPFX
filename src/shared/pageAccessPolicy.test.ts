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
  it("offers Upload-Form to uploaders AND approvers, incl. the HC approver", () => {
    // APRHC joined 2026-08-24 — it is the HC Head of Unit's approver role, held INSTEAD of APR.
    // UPLHC joined 2026-08-31, after an HC-cleared PIC was locked out of this page entirely.
    const p = policyForPage("Upload-Form.aspx");
    expect(p.roles).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
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

  it("offers ApprovalDocument to approvers only — plain and HC", () => {
    const p = policyForPage("ApprovalDocument.aspx");
    expect(p.roles).toEqual(["APR", "APRHC"]);
    expect(p.adminOnly).toBe(false);
  });

  /* ⚠ REVERSED 2026-08-22 (client: "Bulk upload is now allowed for all uploaders to be used").
     The same three roles as the upload form: `UPLHC` because `hou`/`pic_hc` carry no literal `UPL`,
     and `APR` because a HoU group mapped before 2026-08-15 holds only `APR` and `DELS`. */
  it("offers Bulk-Upload to uploaders and their Head of Unit, not just administrators", () => {
    const p = policyForPage("Bulk-Upload.aspx");
    expect(p.adminOnly).toBe(false);
    expect(p.roles).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
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
  /* The ordering still matters, and MORE subtly than before: both rules now yield a non-admin
     policy, so a wrong order no longer shows up as an admin page being handed out. It shows up only
     in the REASON — which is what reconciliation prints when it grants and removes. Assert the
     reason, not just the roles. */
  it("matches bulk upload on the BULK rule, never the upload rule", () => {
    for (const name of ["Bulk-Upload.aspx", "BulkUpload.aspx", "bulk upload.aspx"]) {
      expect(policyForPage(name).adminOnly).toBe(false);
      expect(policyForPage(name).roles).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
      expect(policyForPage(name).reason).toContain("historical documents");
    }
  });

  it("still treats a plain upload page as an uploader page", () => {
    for (const name of ["Upload.aspx", "Upload-Form.aspx", "UploadForm.aspx"]) {
      expect(policyForPage(name).roles).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
    }
  });
});

// My Submissions — an uploader's view of their own files.
// Spec: docs/superpowers/specs/2026-08-14-my-submissions-design.md §3 D3.
describe("the My Submissions page is for everyone who UPLOADS", () => {
  const NAMES = ["My-Submissions.aspx", "MySubmissions.aspx", "My-Uploads.aspx", "My-Files.aspx"];

  it("offers it to every role that can file a document", () => {
    for (const name of NAMES) expect(policyForPage(name).roles).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
  });

  it("REACHES A HEAD OF UNIT, who uploads through UPLHC and not UPL", () => {
    // 2026-08-21. `hou` carries no literal UPL — UPLHC is a superset and LIBRARY_ROLES lists it on
    // the NORMAL approval library too. Keying on UPL alone let a Head of Unit upload a document and
    // then denied them the only page that lists it, and with it the only route to a deletion or
    // share request about their own file. APR is listed as well, because a HoU group mapped before
    // the 2026-08-15 persona correction holds APR + DELS and no upload role whatsoever.
    for (const name of NAMES) {
      expect(isRoleEligibleForPage(name, "UPLHC")).toBe(true);
      expect(isRoleEligibleForPage(name, "APR")).toBe(true);
    }
  });

  it("still offers nothing to a role that cannot upload", () => {
    // Widening this list cannot expose one person's files to another — the page reads both libraries
    // `AuthorId eq <me>`, so a granted role sees only its own rows. It is still kept off the roles
    // that file nothing, because a page that can only ever be empty is a dead link on their menu.
    for (const name of NAMES) {
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

  it("does not capture bulk upload, which has its own rule ahead of them", () => {
    // "Bulk-Upload" contains "upload". Both rules now grant, so the reason is the only tell.
    expect(policyForPage("Bulk-Upload.aspx").reason).toContain("historical documents");
    expect(policyForPage("Upload-Form.aspx").reason).toContain("where documents are submitted");
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
      // Bulk-Upload was the example here until 2026-08-22, when it stopped being admin-only.
      expect(isRoleEligibleForPage("Folder-Administration.aspx", role)).toBe(false);
    }
  });

  it("is safe on empty and unknown input", () => {
    expect(policyForPage("").roles).toEqual(ACTION_ROLES);
    expect(policyForPage(undefined as unknown as string).roles).toEqual(ACTION_ROLES);
    expect(isRoleEligibleForPage("", "UPL")).toBe(true);
  });
});

describe("the Requests page", () => {
  // Spec: docs/superpowers/specs/2026-08-21-requests-page-hod-access-design.md
  it("offers approvers and Heads of Department — the two roles that can carry a request out", () => {
    expect(policyForPage("Requests.aspx").roles).toEqual(["APR", "APRHC", "DEPTVIEW"]);
    expect(policyForPage("CRS-Requests.aspx").roles).toEqual(["APR", "APRHC", "DEPTVIEW"]);
  });

  it("NO LONGER offers uploaders — their view moved to My Submissions on 2026-08-20", () => {
    // Regression guard for the stale-reasoning bug this rule carried for six days: a PIC granted
    // this page gets a queue filtered to units where they hold APR, i.e. none. An empty page on
    // their menu, granted by a comment describing a design that had been superseded.
    for (const name of ["Requests.aspx", "CRS-Requests.aspx", "Approval-Requests.aspx"]) {
      expect(policyForPage(name).roles).not.toContain("UPL");
    }
  });

  it("beats the approver rule, so a name carrying 'approval' still reaches a Head of Department", () => {
    // Ordering, pinned: on the /approv/ rule this page would list approver groups ONLY, and a Head
    // of Department — who holds DEL and SHARE and can act on approved documents — could not open it.
    expect(policyForPage("Approval-Requests.aspx").roles).toEqual(["APR", "APRHC", "DEPTVIEW"]);
  });

  it("does not disturb the approver's own page — DEPTVIEW must never reach the approval queue", () => {
    // The asymmetry is deliberate. A Head of Department is view-and-act-on-approved-documents;
    // letting them into the queue would let them publish into a unit they do not run.
    expect(policyForPage("ApprovalDocument.aspx").roles).toEqual(["APR", "APRHC"]);
    expect(policyForPage("ApprovalDocument.aspx").roles).not.toContain("DEPTVIEW");
  });

  it("is not an administrator tool — a Head of Unit must be able to be granted it", () => {
    expect(policyForPage("Requests.aspx").adminOnly).toBe(false);
  });
});

// ── Every page reconciliation must lock, and every page it must not ──────────
//
// Reconciliation's admin-page lockdown (2026-08-17) decides what to restrict by running each
// Site Pages file name through policyForPage and acting on `adminOnly`. So a pattern that stops
// matching does not fail — it SILENTLY STOPS PROTECTING that page, which is the exact failure the
// pass was built to remove.
//
// These are the page names in the migration runbook §8. If a page is renamed, this list and the
// runbook change together or the rename ships a hole.
// Spec: docs/superpowers/specs/2026-08-17-admin-page-lockdown-design.md
describe("admin page lockdown coverage", () => {
  const MUST_LOCK = [
    "CRS-Settings.aspx",
    "Folder-Administration.aspx",
    "Group-Management.aspx",
    "Site-Access.aspx",
    "Approval-Library-Access.aspx",
    "Page-Access.aspx",
    "Folder-Access.aspx",
    "CRS-Audit-Log.aspx",
  ];

  // The working pages. Locking any of these to owners takes the system away from the people it
  // exists for — a far louder failure than the one above, but worth pinning in the same place.
  const MUST_NOT_LOCK = [
    "Upload-Form.aspx",
    "Approval-Document.aspx",
    "My-Submissions.aspx",
    "CRS-Requests.aspx",
    "Home.aspx",
    /* ⚠ MOVED HERE FROM MUST_LOCK, 2026-08-22, and the move is REQUIRED rather than cosmetic. Bulk
       Upload is now an uploader tool, so the derived-page pass grants it to the uploader groups — and
       if the lockdown pass still claimed it, the two would fight on every run: one granting, the
       other stripping, for ever. The row pass already REFUSES a row targeting an adminOnly page for
       exactly this reason; this is the same conflict arriving from the derived side. */
    "Bulk-Upload.aspx",
  ];

  for (const name of MUST_LOCK) {
    it(`${name} is adminOnly, so reconciliation locks it`, () => {
      expect(policyForPage(name).adminOnly).toBe(true);
    });
  }

  for (const name of MUST_NOT_LOCK) {
    it(`${name} is NOT adminOnly, so reconciliation leaves it alone`, () => {
      expect(policyForPage(name).adminOnly).toBe(false);
    });
  }
});

/* ⚠ THE FOUR-TIME BUG: A PAGE KEYED ON A ROLE ITS AUDIENCE DOES NOT LITERALLY HOLD.
   Upload form needed `APR` (2026-08-17), Requests needed `DEPTVIEW` (2026-08-21), My Submissions
   needed `UPLHC` (2026-08-21), and the upload form needed `UPLHC` too — found live on 2026-08-31,
   when a guest in `GHO_GCA_GCBC_UPLOADER_HIGHLY_CONFIDENTIAL` got AccessDenied while the page's
   ACL, its publish state and every other check read back correct.

   The cause each time: the rule names the role the persona is THOUGHT of as having, while
   `PERSONAS` gives it a superset instead. `pic_hc` is `["UPLHC"]` with no literal `UPL`.

   This pins the three pages an uploader must reach as ONE set, so adding a role to one of them and
   missing the others fails here rather than in a support call. */
describe("every page an uploader must reach lists the HC upload role", () => {
  const UPLOADER_PAGES = ["Upload-Form.aspx", "Bulk-Upload.aspx", "My-Submissions.aspx"];

  it("lists UPLHC on all of them", () => {
    for (const name of UPLOADER_PAGES) {
      expect(policyForPage(name).roles).toContain("UPLHC");
    }
  });

  // An HC-cleared PIC holds UPLHC and nothing else. If any of these pages stops accepting it,
  // that persona silently loses the page - never granted, and stripped by the page pass if it was.
  it("admits a persona holding UPLHC alone", () => {
    for (const name of UPLOADER_PAGES) {
      expect(isRoleEligibleForPage(name, "UPLHC" as GroupMapRole)).toBe(true);
    }
  });

  // The three are deliberately identical: an uploader reaches the form, the bulk form and their
  // own submissions with the same roles. Divergence is what produced this bug.
  it("keeps the three lists identical", () => {
    const [first, ...rest] = UPLOADER_PAGES.map((n) => policyForPage(n).roles);
    for (const roles of rest) expect(roles).toEqual(first);
  });
});
