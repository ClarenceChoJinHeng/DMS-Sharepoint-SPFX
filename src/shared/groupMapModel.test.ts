import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  GroupMapWriteRow,
  GroupMapDraft,
} from "./groupMapModel";

describe("buildGroupMapRow", () => {
  it("builds a trimmed row for a unit-tier UPL mapping", () => {
    const draft: GroupMapDraft = {
      groupId: "  00000000-aaaa-bbbb-cccc-000000000001 ",
      groupName: " DMS_LRC_GCO_UPL ",
      role: "UPL",
      segmentGuid: " efa87c6a-9536-4f7c-910f-011bf7413b80 ",
      tierGuid: " 479a74d8-4126-40f0-9b1f-3374de74152b ",
    };
    expect(buildGroupMapRow(draft)).toEqual({
      GroupId: "00000000-aaaa-bbbb-cccc-000000000001",
      GroupName: "DMS_LRC_GCO_UPL",
      Segment: "efa87c6a-9536-4f7c-910f-011bf7413b80",
      UnitTermGuid: "479a74d8-4126-40f0-9b1f-3374de74152b",
      Role: "UPL",
    });
  });

  it("leaves Segment and UnitTermGuid empty for a GLOBAL row", () => {
    const draft: GroupMapDraft = {
      groupId: "g-1",
      groupName: "DMS_GLOBAL_UPLOADERS",
      role: "GLOBAL",
      segmentGuid: "should-be-ignored",
      tierGuid: "should-be-ignored",
    };
    expect(buildGroupMapRow(draft)).toEqual({
      GroupId: "g-1",
      GroupName: "DMS_GLOBAL_UPLOADERS",
      Segment: "",
      UnitTermGuid: "",
      Role: "GLOBAL",
    });
  });

  it("supports a segment-tier row where the tier GUID equals the segment GUID", () => {
    const row = buildGroupMapRow({
      groupId: "g-2",
      groupName: "DMS_GHO",
      role: "MEMBER",
      segmentGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
      tierGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
    });
    expect(row.Segment).toBe(row.UnitTermGuid);
  });
});

describe("isDuplicateRow", () => {
  const existing: GroupMapWriteRow[] = [
    { GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-gco", Role: "UPL" },
  ];

  it("matches on GroupId + UnitTermGuid + Role, case/space-insensitive", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: " G-1 ", GroupName: "A", Segment: "s", UnitTermGuid: "T-GCO", Role: "UPL",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(true);
  });

  it("allows the same group with a different role", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-gco", Role: "APR",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(false);
  });

  it("allows the same group mapping a different tier", () => {
    const candidate: GroupMapWriteRow = {
      GroupId: "g-1", GroupName: "A", Segment: "s", UnitTermGuid: "t-risk", Role: "UPL",
    };
    expect(isDuplicateRow(existing, candidate)).toBe(false);
  });
});

describe("validateDraft", () => {
  it("returns [] for a complete non-global draft", () => {
    expect(
      validateDraft({ groupId: "g", groupName: "n", role: "MEMBER", segmentGuid: "s", tierGuid: "t" }),
    ).toEqual([]);
  });

  it("requires a group and a role", () => {
    const errs = validateDraft({ groupId: "  ", groupName: "", role: undefined as unknown as "MEMBER" });
    expect(errs).toContain("Select a group.");
    expect(errs).toContain("Select a role.");
  });

  it("requires segment + tier unless GLOBAL", () => {
    expect(validateDraft({ groupId: "g", groupName: "n", role: "UPL" })).toEqual([
      "Select a segment.",
      "Select a tier.",
    ]);
  });

  it("does not require segment/tier for GLOBAL", () => {
    expect(validateDraft({ groupId: "g", groupName: "n", role: "GLOBAL" })).toEqual([]);
  });
});
