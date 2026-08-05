import {
  policyForPage,
  isRoleEligibleForPage,
  VIEW_ONLY_ROLES,
  ACTION_ROLES,
} from "./pageAccessPolicy";
import { GroupMapRole } from "./groupMapModel";

describe("policyForPage — the client's four rules (2026-08-05)", () => {
  it("offers Upload-Form to uploaders only", () => {
    const p = policyForPage("Upload-Form.aspx");
    expect(p.roles).toEqual(["UPL"]);
    expect(p.adminOnly).toBe(false);
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
      expect(policyForPage(name).roles).toEqual(["UPL"]);
    }
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
    expect(isRoleEligibleForPage("Upload-Form.aspx", "UPL")).toBe(true);
    expect(isRoleEligibleForPage("Upload-Form.aspx", "APR")).toBe(false);
    expect(isRoleEligibleForPage("ApprovalDocument.aspx", "APR")).toBe(true);
    expect(isRoleEligibleForPage("ApprovalDocument.aspx", "UPL")).toBe(false);
  });

  it("refuses every role on an admin-only page", () => {
    const all: GroupMapRole[] = ["MEMBER", "UPL", "APR", "DEL", "DELS", "GLOBAL", "SEGVIEW", "HC", "ENTRY"];
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
