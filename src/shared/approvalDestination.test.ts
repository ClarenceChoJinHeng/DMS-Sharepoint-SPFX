import { unitFolderPath, permissionedTierCount, isProblem, UnitFolderResult } from "./approvalDestination";

const WEB = "/sites/CRS";
const SRC = "ApprovalDocument";
const DST = "Shared Documents";

function pathOf(r: UnitFolderResult): string {
  if (isProblem(r)) throw new Error(`expected a path, got: ${r.error}`);
  return r.path;
}

function errorOf(r: UnitFolderResult): string {
  if (!isProblem(r)) throw new Error(`expected a problem, got: ${r.path}`);
  return r.error;
}

describe("permissionedTierCount", () => {
  it("counts the tiers of a head-office chain", () => {
    const chain = JSON.stringify([
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ]);
    expect(permissionedTierCount(chain)).toBe(2);
  });

  it("ignores below-Unit tiers, which is the whole point", () => {
    const chain = JSON.stringify([
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
      { label: "Year", column: "Year", permissioned: false },
      { label: "Document Type", column: "DocumentType", permissioned: false },
      { label: "Archive", column: "Archive", permissioned: false },
    ]);
    expect(permissionedTierCount(chain)).toBe(2);
  });

  it("counts Upstream Ops' own tier names without knowing them", () => {
    const chain = JSON.stringify([
      { label: "Region", column: "Region" },
      { label: "Estate/Mill", column: "EstateMill" },
    ]);
    expect(permissionedTierCount(chain)).toBe(2);
  });

  /* `permissioned` absent means true — matching folderChain.isPermissioned. A tier wrongly counted
     makes the check land one level too DEEP, which refuses loudly; the opposite default would land
     one level shallow on a folder that looks plausible and passes. */
  it("treats an absent `permissioned` as true", () => {
    expect(permissionedTierCount(JSON.stringify([{ label: "A" }, { label: "B" }]))).toBe(2);
  });

  it('does not let the STRING "false" demote a tier', () => {
    const chain = JSON.stringify([{ label: "A" }, { label: "B", permissioned: "false" }]);
    expect(permissionedTierCount(chain)).toBe(2);
  });

  /* undefined, never 0: the caller refuses on both, but only one of them is an administrator's
     problem to go and fix. */
  it("answers undefined — not 0 — for absent, blank or unparseable JSON", () => {
    expect(permissionedTierCount(undefined)).toBeUndefined();
    expect(permissionedTierCount("")).toBeUndefined();
    expect(permissionedTierCount("   ")).toBeUndefined();
    expect(permissionedTierCount("{not json")).toBeUndefined();
    expect(permissionedTierCount(JSON.stringify({ label: "A" }))).toBeUndefined();
    expect(permissionedTierCount(JSON.stringify(["Department", "Unit"]))).toBeUndefined();
  });

  it("answers 0 for an empty array, which the caller refuses separately", () => {
    expect(permissionedTierCount("[]")).toBe(0);
  });
});

describe("unitFolderPath", () => {
  /* THE REGRESSION. This path has three below-Unit tiers, and the old three-levels-up walk landed on
     `2024` — a folder that inherits by design — so every approval in the segment was refused with
     "it is not locked down". Found live 2026-08-24. */
  it("finds the unit under a THREE-tier below-Unit chain", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GCA/EG/2024/Agreement/Archive 1/Proposal - Vendor A - Sample - 19-08-26.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/GHO/GCA/EG");
  });

  it("finds the unit under the built-in two-tier chain", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/Agreement/file.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/GHO/GF/TAX");
  });

  /* The other direction of the same bug: with ONE below-Unit tier, three levels up landed on the
     DEPARTMENT, which does have unique permissions — so it passed, having checked a folder that says
     nothing about whether the unit folder is locked. */
  it("finds the unit under a ONE-tier below-Unit chain, where the old walk passed on the department", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/file.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/GHO/GF/TAX");
  });

  it("handles a file sitting directly in the unit folder", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/file.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/GHO/GF/TAX");
  });

  it("is depth-agnostic below the unit — an added tier changes nothing", () => {
    const base = { webSru: WEB, sourceSegment: SRC, destSegment: DST, permissionedTiers: 2 };
    const shallow = unitFolderPath({ ...base, fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/f.pdf` });
    const deep = unitFolderPath({
      ...base,
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/Agreement/Sub/Deeper/f.pdf`,
    });
    expect(pathOf(shallow)).toBe(pathOf(deep));
  });

  it("routes an HC document to the HC documents library, not the open one", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/HCApprovalDocument/GHO/GCA/EG/2024/Agreement/secret.pdf`,
      webSru: WEB,
      sourceSegment: "HCApprovalDocument",
      destSegment: "HCDocuments",
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/HCDocuments/GHO/GCA/EG");
  });

  /* An unresolved HC pair must never fall back to the open library: that would file a Highly
     Confidential document where the whole unit can read it. */
  it("refuses a blank destination segment rather than falling back", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/HCApprovalDocument/GHO/GCA/EG/2024/secret.pdf`,
      webSru: WEB,
      sourceSegment: "HCApprovalDocument",
      destSegment: "",
      permissionedTiers: 2,
    });
    expect(errorOf(r)).toMatch(/libraries could not be resolved/);
  });

  it("refuses when the tier count is unknown, and says so", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: undefined,
    });
    expect(errorOf(r)).toMatch(/folder structure could not be read/);
  });

  it("refuses a zero tier count as a misread, not as a flat segment", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 0,
    });
    expect(errorOf(r)).toMatch(/no permissioned levels/);
  });

  it("refuses a path outside the approval library", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/SomeOtherLibrary/GHO/GF/TAX/2024/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(errorOf(r)).toMatch(/not inside the approval library/);
  });

  /* A file filed above the unit folder has no unit folder to check, and approving would let
     Auto-route create a path nobody has ACL'd. */
  it("refuses a file filed shallower than the unit folder", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(errorOf(r)).toMatch(/filed only 1 level\(s\) below the business segment/);
  });

  it("compares the library prefix case-insensitively, as SharePoint does", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/approvaldocument/GHO/GF/TAX/2024/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/GHO/GF/TAX");
  });

  it("tolerates stray slashes on the segments and the web url", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/GHO/GF/TAX/2024/f.pdf`,
      webSru: "/sites/CRS/",
      sourceSegment: "/ApprovalDocument/",
      destSegment: "/Shared Documents/",
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/GHO/GF/TAX");
  });

  it("keeps spaces in folder names — encoding is the caller's job", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/UPOPSMY/Central Region/Bukit Mill/2024/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 2,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/UPOPSMY/Central Region/Bukit Mill");
  });

  it("handles a three-tier permissioned family", () => {
    const r = unitFolderPath({
      fileSru: `${WEB}/${SRC}/SEG/A/B/C/2024/Agreement/f.pdf`,
      webSru: WEB,
      sourceSegment: SRC,
      destSegment: DST,
      permissionedTiers: 3,
    });
    expect(pathOf(r)).toBe("/sites/CRS/Shared Documents/SEG/A/B/C");
  });
});
