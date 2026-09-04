import {
  groupRolesById,
  intendedPageGroups,
  groupsToRemove,
  groupsForRequestLists,
  REQUEST_LIST_ROLES,
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
  it("derives UPL, UPLHC, APR and APRHC for the upload form", () => {
    // APRHC joined 2026-08-24: it is the HC Head of Unit's approver role, held INSTEAD of APR, so
    // omitting it here would AccessDeny the exact persona the HC vertical exists for.
    // UPLHC joined 2026-08-31: `pic_hc` is `["UPLHC"]` with no literal `UPL`, so without it every
    // HC-cleared PIC was denied this page — and the page pass would REMOVE their group if it had
    // ever been granted. Found live, with the page ACL and publish state both reading correct.
    expect(derivedRolesForPage("Upload-Form.aspx").slice().sort()).toEqual(["APR", "APRHC", "UPL", "UPLHC"]);
  });

  it("derives APR and APRHC for the approval queue — an uploader must never reach it", () => {
    expect(derivedRolesForPage("ApprovalDocument.aspx")).toEqual(["APR", "APRHC"]);
  });

  it("derives every uploading role for My Submissions, Head of Unit included", () => {
    // UPLHC and APR joined UPL on 2026-08-21: `hou` carries no literal UPL (UPLHC is a superset that
    // LIBRARY_ROLES lists on the normal approval library too), so a Head of Unit could upload and
    // then not open the page listing what they had uploaded. Safe to widen — the page reads
    // `AuthorId eq <me>`, so a granted role sees only its own rows.
    // APRHC joined 2026-08-24, same reason as APR: the HC Head of Unit uploads and must see their
    // own submissions too.
    expect(derivedRolesForPage("My-Submissions.aspx")).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
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
      "Approval-Library-Access.aspx", "DMS-Config.aspx",
    ]) {
      expect(derivedRolesForPage(name)).toEqual([]);
    }
  });

  /* ⚠ Bulk-Upload LEFT the list above on 2026-08-22. It is an uploader tool now, so the derived pass
     owns it and the lockdown pass must not — otherwise the two fight on every run, one granting and
     the other stripping. Asserted positively here so the change cannot be reverted silently. */
  it("derives the uploader roles for Bulk-Upload, which is no longer an admin page", () => {
    expect(derivedRolesForPage("Bulk-Upload.aspx")).toEqual(["UPL", "UPLHC", "APR", "APRHC"]);
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

/* ---------------------------------------------------------------------------
 * Access to the request and submission LISTS — option 2, 2026-08-27.
 * Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md §5
 * ------------------------------------------------------------------------- */
describe("groupsForRequestLists", () => {
  it("includes an uploader, an approver and a Head of Department", () => {
    const out = groupsForRequestLists(groupRolesById([
      folder("1", "GHO_GF_TAX_UPLOADER", "UPL"),
      folder("2", "GHO_GF_TAX_APPROVER", "APR"),
      folder("3", "GHO_GF_HOD", "DEPTVIEW"),
    ]));
    expect(out.map((g) => g.groupId).sort()).toEqual(["1", "2", "3"]);
  });

  it("includes the HC personas' own roles", () => {
    // `hou_hc` holds APRHC and no plain APR, and `pic_hc` holds UPLHC and no plain UPL. Matching
    // only the plain roles would leave every HC unit unable to raise or decide a request.
    const out = groupsForRequestLists(groupRolesById([
      folder("4", "GHO_GF_TAX_APPROVER_HIGHLY_CONFIDENTIAL", "APRHC"),
      folder("5", "GHO_GF_TAX_UPLOADER_HIGHLY_CONFIDENTIAL", "UPLHC"),
    ]));
    expect(out.map((g) => g.groupId).sort()).toEqual(["4", "5"]);
  });

  it("excludes a plain viewer", () => {
    // MEMBER cannot upload, so has nothing to raise a request about. Including them would
    // re-create the site-wide read exposure option 2 exists to remove.
    expect(groupsForRequestLists(groupRolesById([
      folder("6", "GHO_GF_TAX_VIEWER", "MEMBER"),
    ]))).toEqual([]);
  });

  it("excludes C-Level, who hold nothing in either approval library", () => {
    expect(groupsForRequestLists(groupRolesById([
      folder("7", "GHO_SEGVIEW", "SEGVIEW"),
      folder("8", "CRS_GLOBAL", "GLOBAL"),
    ]))).toEqual([]);
  });

  it("names the role that earned it, for the run log", () => {
    const out = groupsForRequestLists(groupRolesById([
      folder("9", "G", "APR"),
      folder("9", "G", "DELS"),
    ]));
    // DELS is not a qualifying role, so it must not appear as the reason.
    expect(out[0].via).toBe("APR");
  });

  it("counts a group once however many qualifying roles it holds", () => {
    const out = groupsForRequestLists(groupRolesById([
      folder("10", "G", "APR"),
      folder("10", "G", "UPL"),
    ]));
    expect(out).toHaveLength(1);
  });

  it("ignores Page and Library rows, as the folder derivation does", () => {
    expect(groupsForRequestLists(groupRolesById([
      { GroupId: "11", Role: "UPL", Scope: "Page", Target: "Upload-Form.aspx" },
      { GroupId: "12", Role: "UPL", Scope: "Library", Target: "Staging" },
    ]))).toEqual([]);
  });

  it("answers empty for no groups — the caller must NOT revoke on that", () => {
    // An empty intended set on a site with no Group Map is legitimate. The reconciliation pass is
    // required to leave the site-entry grant alone in that case, or it locks everybody out.
    expect(groupsForRequestLists([])).toEqual([]);
    expect(groupsForRequestLists(undefined as unknown as [])).toEqual([]);
  });

  it("carries no upload-or-approve role it should not", () => {
    // Pinned so a later edit cannot widen this to a role that would expose the lists further.
    expect(REQUEST_LIST_ROLES.slice().sort())
      .toEqual(["APR", "APRHC", "DEPTVIEW", "UPL", "UPLHC"]);
  });
});
