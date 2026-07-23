import { filterDmsGroups, SpGroup } from "./spGroupsFilter";

const g = (id: number, title: string): SpGroup => ({ id, title });

describe("filterDmsGroups", () => {
  const all = [
    g(3, "SPFX Sandbox Owners"),
    g(4, "SPFX Sandbox Members"),
    g(5, "SPFX Sandbox Visitors"),
    g(27, "DMS_GHO_Group Finance_UPL"),
    g(28, "DMS_GHO_Group Finance_APR"),
    g(29, "dms_gho_group finance"),
  ];

  it("keeps only DMS_-prefixed groups (built-ins excluded)", () => {
    expect(filterDmsGroups(all, "").map((x) => x.id)).toEqual([27, 28, 29]);
  });

  it("prefix match is case-insensitive", () => {
    expect(filterDmsGroups(all, "").some((x) => x.id === 29)).toBe(true);
  });

  it("query filters by substring, case-insensitive", () => {
    expect(filterDmsGroups(all, "_upl").map((x) => x.id)).toEqual([27]);
    expect(filterDmsGroups(all, "FINANCE").map((x) => x.id)).toEqual([27, 28, 29]);
  });

  it("no match returns empty, never throws", () => {
    expect(filterDmsGroups(all, "zzz")).toEqual([]);
    expect(filterDmsGroups([], "x")).toEqual([]);
  });
});
