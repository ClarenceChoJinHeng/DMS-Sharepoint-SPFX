import {
  parseLevels,
  Level,
  matchUserPaths,
  GroupMapRow,
  sanitizeFolderSegment,
  buildLevelFormValues,
  ColumnPair,
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
      '[{"label":"Project Name","column":"ProjectName","labelCol":"Project_x0020_Name","tidCol":"ProjectNameTid"}]';
    expect(parseLevels(json)).toEqual([
      {
        label: "Project Name",
        column: "ProjectName",
        labelCol: "Project_x0020_Name",
        tidCol: "ProjectNameTid",
      },
    ]);
  });
});

describe("matchUserPaths", () => {
  const rows: GroupMapRow[] = [
    { groupId: "g-tax", groupName: "DMS_GF_TAX_UPL", segment: "set-gho", unitTermGuid: "t-tax", role: "UPL" },
    { groupId: "g-treas", groupName: "DMS_GF_TREAS_UPL", segment: "set-gho", unitTermGuid: "t-treas", role: "UPL" },
    { groupId: "g-est", groupName: "DMS_UP_EAST_UPL", segment: "set-up", unitTermGuid: "t-east", role: "UPL" },
  ];

  it("returns the paths for the user's matched group ids", () => {
    expect(matchUserPaths(rows, ["g-tax", "g-est", "unrelated"])).toEqual([
      { segment: "set-gho", unitTermGuid: "t-tax", role: "UPL" },
      { segment: "set-up", unitTermGuid: "t-east", role: "UPL" },
    ]);
  });

  it("is case-insensitive and trims group ids on both sides", () => {
    expect(matchUserPaths(rows, [" G-TAX "])).toEqual([
      { segment: "set-gho", unitTermGuid: "t-tax", role: "UPL" },
    ]);
  });

  it("returns [] when nothing matches", () => {
    expect(matchUserPaths(rows, ["none"])).toEqual([]);
    expect(matchUserPaths([], ["g-tax"])).toEqual([]);
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
});
