import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  roleFromGroupName,
  suggestGroupName,
  GroupMapWriteRow,
  GroupMapDraft,
} from "./groupMapModel";

describe("roleFromGroupName", () => {
  it("derives APR from an _APR suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_APR")).toBe("APR");
  });

  it("derives UPL from an _UPL suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_UPL")).toBe("UPL");
  });

  it("defaults a base (no-suffix) group to MEMBER", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU")).toBe("MEMBER");
  });

  it("matches the suffix case-insensitively and trims", () => {
    expect(roleFromGroupName("  dms_gho_gf_coru_apr  ")).toBe("APR");
    expect(roleFromGroupName("DMS_GHO_GF_CORU_Upl")).toBe("UPL");
  });

  it("never auto-derives GLOBAL and is safe on empty/undefined", () => {
    expect(roleFromGroupName("DMS_GLOBAL_UPLOADERS")).toBe("MEMBER");
    expect(roleFromGroupName("")).toBe("MEMBER");
    expect(roleFromGroupName(undefined as unknown as string)).toBe("MEMBER");
  });
});

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

describe("suggestGroupName", () => {
  it("joins DMS + segment + tier labels + role suffix", () => {
    expect(
      suggestGroupName("Group Head Office", ["Group Finance", "Corporate Reporting"], "UPL"),
    ).toBe("DMS_Group Head Office_Group Finance_Corporate Reporting_UPL");
  });

  it("APR suffix works the same way", () => {
    expect(suggestGroupName("Minamas Head Office", ["Finance"], "APR")).toBe(
      "DMS_Minamas Head Office_Finance_APR",
    );
  });

  it("MEMBER gets no suffix (base/viewer group)", () => {
    expect(suggestGroupName("Group Head Office", ["Group Finance"], "MEMBER")).toBe(
      "DMS_Group Head Office_Group Finance",
    );
  });

  it("GLOBAL is a fixed name regardless of labels", () => {
    expect(suggestGroupName("anything", ["x", "y"], "GLOBAL")).toBe("DMS_GLOBAL");
  });

  it("skips empty labels and trims the rest", () => {
    expect(suggestGroupName(" GHO ", ["", "  Unit A "], "APR")).toBe("DMS_GHO_Unit A_APR");
  });

  it("empty role and empty labels degrade gracefully", () => {
    expect(suggestGroupName("GHO", [], "")).toBe("DMS_GHO");
    expect(suggestGroupName("", [], "")).toBe("DMS");
  });
});
