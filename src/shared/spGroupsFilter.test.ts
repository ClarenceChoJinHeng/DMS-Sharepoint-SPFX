import { filterSelectableGroups, SpGroup } from "./spGroupsFilter";

const g = (id: number, title: string): SpGroup => ({ id, title });

describe("filterSelectableGroups", () => {
  // Built-in titles derive from the SITE title, which is exactly why they are excluded by
  // id here and not by any name pattern.
  const OWNERS = 3, MEMBERS = 4, VISITORS = 5, ENTRY = 9;
  const builtIns = [OWNERS, MEMBERS, VISITORS, ENTRY];

  const all = [
    g(OWNERS, "SPFX Sandbox Owners"),
    g(MEMBERS, "SPFX Sandbox Members"),
    g(VISITORS, "SPFX Sandbox Visitors"),
    g(ENTRY, "DMS_SITE_MEMBERS"),
    g(27, "GHO_GF_CORU_UPLOADER"),
    g(28, "GHO_GF_CORU_APPROVER"),
    g(29, "gho_gf_coru"),
    g(30, "DMS_GHO_Group Finance_UPL"), // legacy prefixed name — must still be selectable
  ];

  it("excludes the supplied built-in ids and keeps everything else", () => {
    expect(filterSelectableGroups(all, "", builtIns).map((x) => x.id)).toEqual([27, 28, 29, 30]);
  });

  // The regression this file exists for: the old filter kept only "DMS_"-prefixed titles,
  // so the client's prefix-less convention would have shown an empty picker.
  it("keeps prefix-less names — the 2026-08-04 convention", () => {
    const ids = filterSelectableGroups(all, "", builtIns).map((x) => x.id);
    expect(ids).toContain(27);
    expect(ids).toContain(29);
  });

  it("keeps legacy DMS_-prefixed names too, so existing groups stay mappable", () => {
    expect(filterSelectableGroups(all, "", builtIns).map((x) => x.id)).toContain(30);
  });

  it("query filters by substring, case-insensitive", () => {
    expect(filterSelectableGroups(all, "_uploader", builtIns).map((x) => x.id)).toEqual([27]);
    expect(filterSelectableGroups(all, "CORU", builtIns).map((x) => x.id)).toEqual([27, 28, 29]);
  });

  // Deliberate: a caller that cannot read the associated groups shows built-ins (visible)
  // rather than silently falling back to a prefix guess.
  it("excludes nothing when no ids are supplied", () => {
    expect(filterSelectableGroups(all, "").length).toBe(all.length);
    expect(filterSelectableGroups(all, "", []).length).toBe(all.length);
  });

  it("no match returns empty, never throws", () => {
    expect(filterSelectableGroups(all, "zzz", builtIns)).toEqual([]);
    expect(filterSelectableGroups([], "x", builtIns)).toEqual([]);
  });
});
