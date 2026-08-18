import {
  groupRolesById,
  intendedPageGroups,
  groupsToRemove,
  GroupMapReadRow,
  CurrentAssignment,
} from "./pageGrants";
import { derivedRolesForPage, pageMatchedRule } from "./pageAccessPolicy";

const folder = (id: string, name: string, role: string): GroupMapReadRow => ({
  GroupId: id, GroupName: name, Role: role, Scope: "Folder", Target: "",
});

describe("groupRolesById — collapsing folder rows", () => {
  it("collapses a group's rows across tiers into one entry", () => {
    const out = groupRolesById([
      folder("217", "GHO_GF_TAX_APPROVER", "APR"),
      folder("217", "GHO_GF_TAX_APPROVER", "DELS"),
      folder("217", "GHO_GF_TAX_APPROVER", "UPLHC"),
    ]);
    expect(out.length).toBe(1);
    expect(out[0].groupId).toBe("217");
    expect(out[0].roles.slice().sort()).toEqual(["APR", "DELS", "UPLHC"]);
  });

  it("keys on GroupId, not the name — two groups renamed alike stay separate", () => {
    const out = groupRolesById([
      folder("217", "SAME_NAME", "UPL"),
      folder("218", "SAME_NAME", "APR"),
    ]);
    expect(out.length).toBe(2);
  });

  it("ignores Library and Page rows", () => {
    const out = groupRolesById([
      { GroupId: "217", Role: "ENTRY", Scope: "Page", Target: "Upload-Form.aspx" },
      { GroupId: "218", Role: "ENTRY", Scope: "Library", Target: "Staging" },
      folder("219", "G", "UPL"),
    ]);
    expect(out.map((g) => g.groupId)).toEqual(["219"]);
  });

  it("treats a BLANK Scope as Folder — the column was added later", () => {
    const out = groupRolesById([{ GroupId: "217", GroupName: "G", Role: "UPL" }]);
    expect(out.length).toBe(1);
    expect(out[0].roles).toEqual(["UPL"]);
  });

  it("normalizes a long-form Role value — the bug that refused a real uploader", () => {
    const out = groupRolesById([folder("217", "G", "Uploader")]);
    expect(out[0].roles).toEqual(["UPL"]);
  });

  it("drops rows with no GroupId or no Role rather than inventing a group", () => {
    const out = groupRolesById([
      { GroupId: "", Role: "UPL", Scope: "Folder" },
      { GroupId: "217", Role: "", Scope: "Folder" },
    ]);
    expect(out).toEqual([]);
  });

  it("never overwrites a real name with a blank one from a later row", () => {
    const out = groupRolesById([
      folder("217", "GHO_GF_TAX_UPLOADER", "UPL"),
      { GroupId: "217", GroupName: "", Role: "DELS", Scope: "Folder" },
    ]);
    expect(out[0].groupName).toBe("GHO_GF_TAX_UPLOADER");
  });

  it("survives an empty and an undefined input", () => {
    expect(groupRolesById([])).toEqual([]);
    expect(groupRolesById(undefined as unknown as GroupMapReadRow[])).toEqual([]);
  });
});

describe("derivedRolesForPage — the gate that stops a site-wide lockout", () => {
  it("derives UPL and APR for the upload form", () => {
    expect(derivedRolesForPage("Upload-Form.aspx").slice().sort()).toEqual(["APR", "UPL"]);
  });

  it("derives APR only for the approval queue — an uploader must never reach it", () => {
    expect(derivedRolesForPage("ApprovalDocument.aspx")).toEqual(["APR"]);
  });

  it("derives UPL for My Submissions", () => {
    expect(derivedRolesForPage("My-Submissions.aspx")).toEqual(["UPL"]);
  });

  /**
   * THE ONE THAT MATTERS. An unmatched page falls to DEFAULT_POLICY (UPL, APR, DELS). Deriving from
   * it would break the page's inheritance and grant those three, locking out every other role —
   * SDG Employee, Head of Department and C-Level hold none of them. On the site home page or a page
   * the client authored, that is a site-wide lockout on a run that reports success.
   */
  it("derives NOTHING for a page that matched no rule", () => {
    for (const name of ["CollabHome.aspx", "Home.aspx", "Team-News.aspx", "Whatever.aspx"]) {
      expect(pageMatchedRule(name)).toBe(false);
      expect(derivedRolesForPage(name)).toEqual([]);
    }
  });

  it("derives NOTHING for an administrator-only page — the lockdown pass owns those", () => {
    for (const name of [
      "Folder-Administration.aspx", "Group-Management.aspx", "Folder-Access.aspx",
      "Page-Access.aspx", "Site-Access.aspx", "CRS-Settings.aspx", "CRS-Audit-Log.aspx",
      "Bulk-Upload.aspx", "Approval-Library-Access.aspx", "DMS-Config.aspx",
    ]) {
      expect(derivedRolesForPage(name)).toEqual([]);
    }
  });

  it("never derives a view-only role", () => {
    for (const name of ["Upload-Form.aspx", "ApprovalDocument.aspx", "Requests.aspx"]) {
      const roles = derivedRolesForPage(name) as string[];
      expect(roles.indexOf("MEMBER")).toBe(-1);
      expect(roles.indexOf("GLOBAL")).toBe(-1);
      expect(roles.indexOf("SEGVIEW")).toBe(-1);
    }
  });

  it("handles a blank name without matching anything", () => {
    expect(derivedRolesForPage("")).toEqual([]);
    expect(derivedRolesForPage(undefined as unknown as string)).toEqual([]);
  });
});

describe("intendedPageGroups", () => {
  const groups = groupRolesById([
    folder("101", "GHO_GF_TAX_UPLOADER", "UPL"),
    folder("102", "GHO_GF_TAX_APPROVER", "APR"),
    folder("102", "GHO_GF_TAX_APPROVER", "DELS"),
    folder("103", "GHO_GF_TAX_EMPLOYEE", "MEMBER"),
    folder("104", "GHO_HOD", "DEPTVIEW"),
  ]);

  it("gives the upload form to the uploader and the approver, and nobody else", () => {
    const out = intendedPageGroups("Upload-Form.aspx", groups, []);
    expect(out.map((g) => g.groupId).slice().sort()).toEqual(["101", "102"]);
    expect(out.every((g) => g.source === "derived")).toBe(true);
  });

  it("gives the approval queue to the approver ONLY — never the uploader", () => {
    const out = intendedPageGroups("ApprovalDocument.aspx", groups, []);
    expect(out.map((g) => g.groupId)).toEqual(["102"]);
  });

  it("excludes view-only groups from every page", () => {
    for (const page of ["Upload-Form.aspx", "ApprovalDocument.aspx", "My-Submissions.aspx"]) {
      const ids = intendedPageGroups(page, groups, []).map((g) => g.groupId);
      expect(ids.indexOf("103")).toBe(-1);
      expect(ids.indexOf("104")).toBe(-1);
    }
  });

  it("returns nothing for an unmatched or admin-only page, even with groups present", () => {
    expect(intendedPageGroups("CollabHome.aspx", groups, [])).toEqual([]);
    expect(intendedPageGroups("Folder-Access.aspx", groups, [])).toEqual([]);
  });

  it("MERGES a hand-made Page row for a group the policy does not imply", () => {
    const out = intendedPageGroups("Upload-Form.aspx", groups, [
      { GroupId: "103", GroupName: "GHO_GF_TAX_EMPLOYEE", Role: "ENTRY", Scope: "Page", Target: "Upload-Form.aspx" },
    ]);
    expect(out.map((g) => g.groupId).slice().sort()).toEqual(["101", "102", "103"]);
    expect(out.filter((g) => g.groupId === "103")[0].source).toBe("row");
  });

  it("does not double-list a group that is both derived and hand-rowed", () => {
    const out = intendedPageGroups("Upload-Form.aspx", groups, [
      { GroupId: "101", Role: "ENTRY", Scope: "Page", Target: "Upload-Form.aspx" },
    ]);
    expect(out.filter((g) => g.groupId === "101").length).toBe(1);
    expect(out.filter((g) => g.groupId === "101")[0].source).toBe("derived");
  });

  it("ignores a hand row aimed at a different page", () => {
    const out = intendedPageGroups("Upload-Form.aspx", groups, [
      { GroupId: "103", Role: "ENTRY", Scope: "Page", Target: "My-Submissions.aspx" },
    ]);
    expect(out.map((g) => g.groupId).slice().sort()).toEqual(["101", "102"]);
  });

  it("matches the hand row's Target case-insensitively", () => {
    const out = intendedPageGroups("Upload-Form.aspx", [], [
      { GroupId: "103", Role: "ENTRY", Scope: "Page", Target: "upload-form.ASPX" },
    ]);
    expect(out.map((g) => g.groupId)).toEqual(["103"]);
  });

  it("honours a hand row even on a page that derives nothing, so exceptions still work", () => {
    const out = intendedPageGroups("CollabHome.aspx", groups, [
      { GroupId: "103", Role: "ENTRY", Scope: "Page", Target: "CollabHome.aspx" },
    ]);
    expect(out.map((g) => g.groupId)).toEqual(["103"]);
  });

  it("records the role that earned the page, for the log", () => {
    const out = intendedPageGroups("Upload-Form.aspx", groups, []);
    expect(out.filter((g) => g.groupId === "101")[0].via).toBe("UPL");
  });

  it("is empty when no group holds a qualifying role — the warned case", () => {
    const viewersOnly = groupRolesById([folder("103", "E", "MEMBER")]);
    expect(intendedPageGroups("Upload-Form.aspx", viewersOnly, [])).toEqual([]);
    // ...and the caller can still tell this apart from "leave the page alone".
    expect(derivedRolesForPage("Upload-Form.aspx").length).toBeGreaterThan(0);
  });
});

describe("groupsToRemove", () => {
  const intended = [
    { groupId: "101", groupName: "UPLOADER", source: "derived" as const, via: "UPL" },
    { groupId: "102", groupName: "APPROVER", source: "derived" as const, via: "APR" },
  ];
  const asg = (principalId: number, title: string, isGroup = true): CurrentAssignment =>
    ({ principalId, title, isGroup });

  it("removes a group that holds the page but no qualifying role", () => {
    const out = groupsToRemove(intended, [asg(101, "UPLOADER"), asg(999, "EMPLOYEE")], [3]);
    expect(out.map((a) => a.principalId)).toEqual([999]);
  });

  it("removes nothing when the page already matches", () => {
    expect(groupsToRemove(intended, [asg(101, "U"), asg(102, "A")], [3])).toEqual([]);
  });

  it("NEVER removes a protected principal — Owners, Members, Visitors, site entry", () => {
    const out = groupsToRemove(
      intended,
      [asg(3, "Owners"), asg(4, "Members"), asg(7, "CRS_SITE_MEMBERS")],
      [3, 4, 5, 7],
    );
    expect(out).toEqual([]);
  });

  it("NEVER removes a user principal, however unexpected", () => {
    const out = groupsToRemove(intended, [asg(555, "Some Person", false)], []);
    expect(out).toEqual([]);
  });

  it("cannot keep a group whose GroupId is not an integer", () => {
    // A legacy Entra GUID row cannot be converted to a principal id, so the grant it describes
    // could not have been made by this pass. Whatever is on the page is therefore unexplained.
    const guidRow = [{ groupId: "8f2c-guid-row", groupName: "legacy", source: "row" as const, via: "" }];
    const out = groupsToRemove(guidRow, [asg(101, "UPLOADER")], []);
    expect(out.map((a) => a.principalId)).toEqual([101]);
  });

  it("survives empty and undefined inputs", () => {
    expect(groupsToRemove([], [], [])).toEqual([]);
    expect(groupsToRemove(
      undefined as unknown as typeof intended,
      undefined as unknown as CurrentAssignment[],
      undefined as unknown as number[],
    )).toEqual([]);
  });
});
