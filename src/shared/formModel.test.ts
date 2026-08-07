import {
  parseLevels,
  Level,
  collectMembership,
  isLeafChainValid,
  GroupMapRow,
  sanitizeFolderSegment,
  buildLevelFormValues,
  ColumnPair,
  parseReconModes,
  RawModeRow,
} from "./formModel";

describe("parseLevels", () => {
  it("parses a valid Levels JSON array", () => {
    const json = '[{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]';
    const result: Level[] = parseLevels(json);
    expect(result).toEqual([
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ]);
  });

  it("returns [] for empty, null, or malformed JSON", () => {
    expect(parseLevels("")).toEqual([]);
    expect(parseLevels(undefined as unknown as string)).toEqual([]);
    expect(parseLevels("{not json}")).toEqual([]);
    expect(parseLevels('"a string"')).toEqual([]);
  });

  it("drops entries missing label or column", () => {
    const json = '[{"label":"Region","column":"Region"},{"label":"x"},{"column":"y"}]';
    expect(parseLevels(json)).toEqual([{ label: "Region", column: "Region" }]);
  });

  it("carries optional labelCol/tidCol internal-name overrides", () => {
    const json =
      '[{"label":"Group Project Name","column":"GroupProjectName","labelCol":"GroupProjectName","tidCol":"GroupProjectNameTid"}]';
    expect(parseLevels(json)).toEqual([
      {
        label: "Group Project Name",
        column: "GroupProjectName",
        labelCol: "GroupProjectName",
        tidCol: "GroupProjectNameTid",
      },
    ]);
  });
});

describe("collectMembership", () => {
  // GHO chain: segment(set-gho) -> dept(t-lrc) -> unit(t-gco).
  const rows: GroupMapRow[] = [
    { groupId: "g-seg", groupName: "GHO Segment", segment: "set-gho", termGuid: "set-gho", role: "MEMBER" },
    { groupId: "g-lrc", groupName: "LRC Dept", segment: "set-gho", termGuid: "t-lrc", role: "MEMBER" },
    { groupId: "g-gco-upl", groupName: "GCO Uploaders", segment: "set-gho", termGuid: "t-gco", role: "UPL" },
    { groupId: "g-risk-upl", groupName: "Risk Uploaders", segment: "set-gho", termGuid: "t-risk", role: "UPL" },
    { groupId: "g-global", groupName: "Global Reader", segment: "", termGuid: "", role: "GLOBAL" },
  ];

  it("collects member terms and uploader leaves for the user's groups", () => {
    const m = collectMembership(rows, ["g-seg", "g-lrc", "g-gco-upl"]);
    expect(m.memberTerms).toEqual(new Set(["set-gho", "t-lrc", "t-gco"]));
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho" }]);
  });

  it("ignores GLOBAL rows entirely — a read-only role, never an uploader", () => {
    const m = collectMembership(rows, ["g-global"]);
    expect(m.uploaderLeaves).toEqual([]);
    expect(m.memberTerms.size).toBe(0);
  });

  it("gives a GLOBAL member no upload paths even alongside a real upl group", () => {
    // GLOBAL adds nothing; the UPL leaf is the only thing that grants upload.
    const m = collectMembership(rows, ["g-global", "g-gco-upl"]);
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho" }]);
    expect(m.memberTerms).toEqual(new Set(["t-gco"]));
  });

  it("is case-insensitive and trims group ids on both sides", () => {
    const m = collectMembership(rows, [" G-GCO-UPL "]);
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho" }]);
  });

  it("returns empty membership when nothing matches", () => {
    const m = collectMembership(rows, ["none"]);
    expect(m.memberTerms.size).toBe(0);
    expect(m.uploaderLeaves).toEqual([]);
  });

  // The regression this block exists for. Every test above uses the SHORT code "UPL", which is
  // why the bug survived: the Group Map's Role is a Choice column holding LONG FORM values, so
  // live rows read "UPLOADER" and the old `role === "UPL"` compare never matched one.
  //
  // The failure was silent and asymmetric. Reconciliation normalised the role and granted the
  // folder ACL, so permissions looked correct in SharePoint — while the form collected no
  // uploader leaf and told a correctly provisioned user "your account isn't fully provisioned
  // to upload". Verified live 2026-08-07.
  describe("long-form role values (the Choice column's actual values)", () => {
    const longForm: GroupMapRow[] = [
      { groupId: "53", groupName: "GHO_GF_CORU_UPL", segment: "set-gho", termGuid: "t-coru", role: "UPLOADER" },
      { groupId: "60", groupName: "GHO_GF_CORU_APR", segment: "set-gho", termGuid: "t-coru", role: "APPROVER" },
    ];

    it("treats UPLOADER as UPL", () => {
      const m = collectMembership(longForm, ["53"]);
      expect(m.uploaderLeaves).toEqual([{ termGuid: "t-coru", segment: "set-gho" }]);
    });

    it("does not treat APPROVER as an uploader", () => {
      // Normalising must not turn every long-form role into an upload grant.
      const m = collectMembership(longForm, ["60"]);
      expect(m.uploaderLeaves).toEqual([]);
      expect(m.memberTerms).toEqual(new Set(["t-coru"]));
    });

    it("still honours the short code, so existing rows keep working", () => {
      const m = collectMembership(rows, ["g-gco-upl"]);
      expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho" }]);
    });

    it("ignores an unknown role rather than inventing one", () => {
      const odd: GroupMapRow[] = [
        { groupId: "99", groupName: "Odd", segment: "set-gho", termGuid: "t-x", role: "SOMETHING_ELSE" },
      ];
      expect(collectMembership(odd, ["99"]).uploaderLeaves).toEqual([]);
    });
  });
});

describe("isLeafChainValid", () => {
  // Regression: an uploader holds ONE Group Map row — the unit row. Its
  // ancestors (segment, department) have no rows and must not need any.
  // The predecessor check demanded membership at every tier and refused a
  // correctly provisioned uploader. See the leaf-only-upload-authorization
  // spec (2026-07-29).
  it("accepts a chain ending at the leaf with no rows for the ancestors", () => {
    expect(isLeafChainValid(["t-lrc", "t-gco"], "t-gco")).toBe(true);
  });

  it("accepts a deeper chain (Project mode: project → dept → unit)", () => {
    expect(isLeafChainValid(["p-1", "d-1", "u-1"], "u-1")).toBe(true);
  });

  it("rejects an empty chain (loadTermPath failed)", () => {
    expect(isLeafChainValid([], "t-gco")).toBe(false);
  });

  it("rejects a chain that does not terminate at the leaf", () => {
    // a truncated ancestry would otherwise point the upload at the department
    expect(isLeafChainValid(["t-lrc"], "t-gco")).toBe(false);
    expect(isLeafChainValid(["t-lrc", "t-other"], "t-gco")).toBe(false);
  });

  // normGuid trims + lowercases only — it does not strip braces.
  it("normalises case and surrounding whitespace on both sides", () => {
    expect(isLeafChainValid(["T-LRC", "T-GCO"], "t-gco")).toBe(true);
    expect(isLeafChainValid(["t-lrc", " t-gco "], "T-GCO")).toBe(true);
  });
});

describe("sanitizeFolderSegment", () => {
  it("removes illegal folder characters and trims", () => {
    expect(sanitizeFolderSegment('In*voice:2026?')).toBe("Invoice2026");
    expect(sanitizeFolderSegment("  2026  ")).toBe("2026");
  });
  it("collapses internal whitespace to single spaces", () => {
    expect(sanitizeFolderSegment("Board   Papers")).toBe("Board Papers");
  });
  it("returns '' for empty/whitespace input", () => {
    expect(sanitizeFolderSegment("   ")).toBe("");
    expect(sanitizeFolderSegment("")).toBe("");
  });
});

describe("buildLevelFormValues", () => {
  const columnMap: Record<string, ColumnPair> = {
    Department: { label: "Department", tid: "DepartmentTid" },
    Unit: { label: "Unit", tid: "UnitTid" },
    BusinessSegment: { label: "Business_x0020_Segment", tid: "BusinessSegmentTid" },
  };
  it("emits a label value and a tid value per selection using explicit tid names", () => {
    const result = buildLevelFormValues(
      columnMap,
      [
        { column: "Department", label: "Group Finance", id: "t-gf" },
        { column: "Unit", label: "Tax", id: "t-tax" },
      ],
    );
    expect(result).toEqual([
      { FieldName: "Department", FieldValue: "Group Finance" },
      { FieldName: "DepartmentTid", FieldValue: "t-gf" },
      { FieldName: "Unit", FieldValue: "Tax" },
      { FieldName: "UnitTid", FieldValue: "t-tax" },
    ]);
  });
  it("uses the mapped internal names and skips unmapped columns", () => {
    const result = buildLevelFormValues(columnMap, [
      { column: "BusinessSegment", label: "Group Head Office", id: "set-gho" },
      { column: "Nope", label: "x", id: "y" },
    ]);
    expect(result).toEqual([
      { FieldName: "Business_x0020_Segment", FieldValue: "Group Head Office" },
      { FieldName: "BusinessSegmentTid", FieldValue: "set-gho" },
    ]);
  });
  it("skips selections with a blank label", () => {
    expect(buildLevelFormValues(columnMap, [{ column: "Unit", label: "", id: "t" }])).toEqual([]);
  });

  it("prefers a selection's own labelCol/tidCol over the columnMap", () => {
    const result = buildLevelFormValues(columnMap, [
      {
        column: "Department",
        label: "Group Finance",
        id: "t-gf",
        labelCol: "ClientDept",
        tidCol: "ClientDeptTid",
      },
    ]);
    expect(result).toEqual([
      { FieldName: "ClientDept", FieldValue: "Group Finance" },
      { FieldName: "ClientDeptTid", FieldValue: "t-gf" },
    ]);
  });

  it("uses labelCol/tidCol even when the column isn't in the columnMap", () => {
    const result = buildLevelFormValues(columnMap, [
      {
        column: "Region",
        label: "East",
        id: "t-east",
        labelCol: "Region",
        tidCol: "RegionTid",
      },
    ]);
    expect(result).toEqual([
      { FieldName: "Region", FieldValue: "East" },
      { FieldName: "RegionTid", FieldValue: "t-east" },
    ]);
  });

  it("resolves a Group Project Name level to its label and Tid columns", () => {
    const columnMap: Record<string, ColumnPair> = {
      GroupProjectName: {
        label: "GroupProjectName",
        tid: "GroupProjectNameTid",
      },
    };
    const result = buildLevelFormValues(columnMap, [
      {
        column: "GroupProjectName",
        label: "Blue Sky",
        id: "3f2a1c8e-0000-0000-0000-000000000001",
      },
    ]);
    expect(result).toEqual([
      { FieldName: "GroupProjectName", FieldValue: "Blue Sky" },
      {
        FieldName: "GroupProjectNameTid",
        FieldValue: "3f2a1c8e-0000-0000-0000-000000000001",
      },
    ]);
  });
});

describe("parseReconModes", () => {
  const levels = '[{"label":"Department","column":"Department"}]';

  it("returns term set + staging folder for valid mode rows, sorted by SortOrder", () => {
    const rows: RawModeRow[] = [
      { TermSetGuid: "set-upstream", StagingFolder: "Group Upstream Operations", Levels: levels, SortOrder: 2 },
      { TermSetGuid: "set-gho", StagingFolder: "Group Head Office", Levels: levels, SortOrder: 1 },
    ];
    expect(parseReconModes(rows)).toEqual([
      { termSetGuid: "set-gho", stagingFolder: "Group Head Office", sortOrder: 1 },
      { termSetGuid: "set-upstream", stagingFolder: "Group Upstream Operations", sortOrder: 2 },
    ]);
  });

  it("trims whitespace so a trailing space in the GUID or folder can't break matching", () => {
    const rows: RawModeRow[] = [
      { TermSetGuid: "  set-gho  ", StagingFolder: " Group Head Office ", Levels: levels, SortOrder: 1 },
    ];
    expect(parseReconModes(rows)[0]).toEqual({
      termSetGuid: "set-gho",
      stagingFolder: "Group Head Office",
      sortOrder: 1,
    });
  });

  it("ignores rows with an empty/old-schema Levels (e.g. retired department/project rows)", () => {
    const rows: RawModeRow[] = [
      { TermSetGuid: "old-dept", StagingFolder: "Departments", Levels: "", SortOrder: 1 },
      { TermSetGuid: "old-proj", StagingFolder: "Projects", SortOrder: 2 },
      { TermSetGuid: "set-gho", StagingFolder: "Group Head Office", Levels: levels, SortOrder: 3 },
    ];
    expect(parseReconModes(rows)).toEqual([
      { termSetGuid: "set-gho", stagingFolder: "Group Head Office", sortOrder: 3 },
    ]);
  });

  it("drops rows missing a term set or staging folder", () => {
    const rows: RawModeRow[] = [
      { TermSetGuid: "", StagingFolder: "X", Levels: levels, SortOrder: 1 },
      { TermSetGuid: "set-x", StagingFolder: "  ", Levels: levels, SortOrder: 2 },
    ];
    expect(parseReconModes(rows)).toEqual([]);
  });

  it("defaults a missing SortOrder to 0", () => {
    const rows: RawModeRow[] = [
      { TermSetGuid: "set-gho", StagingFolder: "Group Head Office", Levels: levels },
    ];
    expect(parseReconModes(rows)[0].sortOrder).toBe(0);
  });
});
