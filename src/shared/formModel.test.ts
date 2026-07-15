import { parseLevels, Level } from "./formModel";

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
