import { parseLevels, Level, matchUserPaths, GroupMapRow, sanitizeFolderSegment } from "./formModel";

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
