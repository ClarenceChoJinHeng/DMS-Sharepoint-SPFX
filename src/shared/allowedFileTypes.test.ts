import { normalizeFileTypes } from "./allowedFileTypes";

describe("normalizeFileTypes", () => {
  // Regression: a bare "png" typed into DMS Config produced an invalid `accept`
  // token, greying PNG out of the file picker with no error. 2026-07-30.
  it("prepends a missing leading dot", () => {
    expect(normalizeFileTypes(["png"])).toEqual([".png"]);
  });

  it("lowercases", () => {
    expect(normalizeFileTypes([".PNG", ".Docx"])).toEqual([".png", ".docx"]);
  });

  it("trims surrounding whitespace", () => {
    expect(normalizeFileTypes(["  .pdf  "])).toEqual([".pdf"]);
  });

  it("drops empty and whitespace-only entries", () => {
    expect(normalizeFileTypes([".pdf", "", "   ", ".doc"])).toEqual([".pdf", ".doc"]);
  });

  it("de-duplicates, including entries that only differ by dot or case", () => {
    expect(normalizeFileTypes([".pdf", "pdf", ".PDF"])).toEqual([".pdf"]);
  });

  it("preserves the order the values arrived in", () => {
    expect(normalizeFileTypes([".xls", ".doc", ".pdf"])).toEqual([".xls", ".doc", ".pdf"]);
  });

  it("returns an empty array for an empty input", () => {
    expect(normalizeFileTypes([])).toEqual([]);
  });
});
