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
  filterConfidentiality,
} from "./formModel";

// Real GUIDs from the test site's Confidentiality Level set, so a copy-paste
// slip in the fixture cannot pass by matching itself.
const CONF = [
  { id: "87f8481b-dbe1-4d60-bbc8-3d111e6e0e5b", label: "Confidential" },
  { id: "420d75d5-f3b7-4525-9eb7-ec06590e7f22", label: "Highly Confidential" },
  { id: "3eccca8e-1617-4713-b923-988dd285c47d", label: "Restricted" },
];
const HC_GUID = "420d75d5-f3b7-4525-9eb7-ec06590e7f22";

describe("filterConfidentiality", () => {
  it("offers everything except Highly Confidential to a UPL-only user", () => {
    expect(filterConfidentiality(CONF, ["UPL"], HC_GUID).map((o) => o.label)).toEqual([
      "Confidential",
      "Restricted",
    ]);
  });

  it("offers ONLY Highly Confidential to an HC-only user", () => {
    // PIC #2 reaches the HC library and nothing else, so Confidential and
    // Restricted would point at a library they cannot open.
    expect(filterConfidentiality(CONF, ["HC"], HC_GUID).map((o) => o.label)).toEqual([
      "Highly Confidential",
    ]);
  });

  it("offers all three when both roles are held", () => {
    expect(filterConfidentiality(CONF, ["UPL", "HC"], HC_GUID)).toHaveLength(3);
  });

  it("offers nothing when neither upload role is held", () => {
    expect(filterConfidentiality(CONF, [], HC_GUID)).toEqual([]);
  });

  it("matches the HC term case-insensitively and ignores surrounding space", () => {
    // The GUID arrives from a DMS Config cell, so stray whitespace and casing are
    // both realistic, and a miss would route HC documents into Staging.
    expect(filterConfidentiality(CONF, ["HC"], `  ${HC_GUID.toUpperCase()}  `)).toHaveLength(1);
  });

  it("does not leak Highly Confidential to a UPL user when hcTermGuid is unset", () => {
    // With nothing to match, no option is the HC one: UPL keeps all three rather
    // than losing one, and HC gets none. The form blocks on the missing config
    // separately instead of this function guessing.
    expect(filterConfidentiality(CONF, ["UPL"], "")).toHaveLength(3);
    expect(filterConfidentiality(CONF, ["HC"], "")).toEqual([]);
  });
});

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
    { groupId: "g-gco-hc", groupName: "GCO HC", segment: "set-gho", termGuid: "t-gco", role: "HC" },
  ];

  it("collects member terms and uploader leaves for the user's groups", () => {
    const m = collectMembership(rows, ["g-seg", "g-lrc", "g-gco-upl"]);
    expect(m.memberTerms).toEqual(new Set(["set-gho", "t-lrc", "t-gco"]));
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho", role: "UPL" }]);
  });

  it("treats HC as an upload role, so an HC-only member gets a path", () => {
    // PIC #2 holds nothing but the HC group. Before HC counted here they were
    // told their account "isn't fully provisioned to upload" — correctly
    // provisioned, and refused.
    const m = collectMembership(rows, ["g-gco-hc"]);
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho", role: "HC" }]);
  });

  it("reports both roles when a user holds UPL and HC on the same unit", () => {
    const m = collectMembership(rows, ["g-gco-upl", "g-gco-hc"]);
    expect(m.uploaderLeaves.map((l) => l.role).sort()).toEqual(["HC", "UPL"]);
  });

  it("ignores GLOBAL rows entirely — a read-only role, never an uploader", () => {
    const m = collectMembership(rows, ["g-global"]);
    expect(m.uploaderLeaves).toEqual([]);
    expect(m.memberTerms.size).toBe(0);
  });

  it("gives a GLOBAL member no upload paths even alongside a real upl group", () => {
    // GLOBAL adds nothing; the UPL leaf is the only thing that grants upload.
    const m = collectMembership(rows, ["g-global", "g-gco-upl"]);
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho", role: "UPL" }]);
    expect(m.memberTerms).toEqual(new Set(["t-gco"]));
  });

  it("is case-insensitive and trims group ids on both sides", () => {
    const m = collectMembership(rows, [" G-GCO-UPL "]);
    expect(m.uploaderLeaves).toEqual([{ termGuid: "t-gco", segment: "set-gho", role: "UPL" }]);
  });

  it("returns empty membership when nothing matches", () => {
    const m = collectMembership(rows, ["none"]);
    expect(m.memberTerms.size).toBe(0);
    expect(m.uploaderLeaves).toEqual([]);
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
