import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  roleFromGroupName,
  suggestGroupName,
  GroupMapWriteRow,
  GroupMapDraft,
  PERSONAS,
  SELECTABLE_ROLES,
  personaByKey,
} from "./groupMapModel";

describe("personas", () => {
  it("never invents a bundle group — every persona is a set of atomic roles", () => {
    const atomic = ["MEMBER", "UPL", "APR", "DEL", "HC"];
    for (const p of PERSONAS) {
      expect(p.roles.length).toBeGreaterThan(0);
      for (const r of p.roles) expect(atomic).toContain(r);
    }
  });

  it("gives every persona a distinct key and a distinct role set", () => {
    const keys = PERSONAS.map((p) => p.key);
    expect(new Set(keys).size).toBe(keys.length);
    // Two personas with identical roles would be indistinguishable once provisioned —
    // the 2x2 of Head-of bundles collapses if approve+upload equals approve+upload+delete.
    const sets = PERSONAS.map((p) => p.roles.slice().sort().join("+"));
    expect(new Set(sets).size).toBe(sets.length);
  });

  it("gives every Head-of persona the approve role — that is what makes it a Head-of", () => {
    for (const p of PERSONAS) {
      if (p.key.indexOf("headof") === 0) expect(p.roles).toContain("APR");
    }
  });

  it("keeps PIC 3 off Documents by omitting the base group", () => {
    expect(personaByKey("pic3")?.roles).toEqual(["UPL"]);
  });

  it("offers no HC persona or HC role in Phase 1 — the term does not exist", () => {
    for (const p of PERSONAS) expect(p.roles).not.toContain("HC");
    expect(SELECTABLE_ROLES).not.toContain("HC");
  });

  it("returns undefined for an unknown persona key rather than throwing", () => {
    expect(personaByKey("nope")).toBeUndefined();
  });
});

describe("roleFromGroupName — DEL and HC suffixes", () => {
  it("derives DEL from a _DEL suffix", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_DEL")).toBe("DEL");
  });

  it("derives HC from an _HC suffix, so a Phase 2 group is not misread as MEMBER", () => {
    expect(roleFromGroupName("DMS_GHO_GF_CORU_HC")).toBe("HC");
  });

  it("does not mistake a unit whose name merely ends in those letters", () => {
    // No underscore before the suffix — this is a unit called "MODEL", not a DEL group.
    expect(roleFromGroupName("DMS_GHO_GF_MODEL")).toBe("MEMBER");
  });
});

describe("suggestGroupName — new role suffixes", () => {
  it("suffixes DEL", () => {
    expect(suggestGroupName("GHO", ["Group Finance"], "DEL")).toBe("DMS_GHO_Group Finance_DEL");
  });

  it("leaves MEMBER unsuffixed — the base group is identified by having no suffix", () => {
    expect(suggestGroupName("GHO", ["Group Finance"], "MEMBER")).toBe("DMS_GHO_Group Finance");
  });

  it("round-trips: a suggested name parses back to the role it was built for", () => {
    for (const role of SELECTABLE_ROLES) {
      if (role === "GLOBAL") continue; // fixed name, no tier
      expect(roleFromGroupName(suggestGroupName("GHO", ["GF", "CORU"], role))).toBe(role);
    }
  });
});

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
