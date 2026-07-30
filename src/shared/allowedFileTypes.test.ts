import { normalizeFileTypes, readChoiceArray } from "./allowedFileTypes";

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

describe("readChoiceArray", () => {
  it("reads the plain array shape (odata=nometadata / minimalmetadata)", () => {
    expect(readChoiceArray([".pdf", ".doc"])).toEqual([".pdf", ".doc"]);
  });

  it("reads the wrapped shape (odata=verbose)", () => {
    expect(readChoiceArray({ results: [".pdf", ".doc"] })).toEqual([".pdf", ".doc"]);
  });

  it("reads an empty selection in both shapes as an empty array, NOT as absent", () => {
    expect(readChoiceArray([])).toEqual([]);
    expect(readChoiceArray({ results: [] })).toEqual([]);
  });

  it("returns undefined when the field is absent", () => {
    expect(readChoiceArray(undefined)).toBeUndefined();
    expect(readChoiceArray(null)).toBeUndefined();
  });

  it("returns undefined for shapes that are not a choice array", () => {
    expect(readChoiceArray(".pdf,.doc")).toBeUndefined();
    expect(readChoiceArray({ notResults: [".pdf"] })).toBeUndefined();
  });
});
