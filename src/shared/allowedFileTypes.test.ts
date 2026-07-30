import {
  normalizeFileTypes,
  readAllowedFileTypesField,
  readChoiceArray,
  resolveAllowedFileTypes,
  FALLBACK_FILE_TYPES,
  NO_TYPES_MESSAGE,
  CONFIG_UNREADABLE_MESSAGE,
} from "./allowedFileTypes";

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

  // A verbose-mode response can carry a null `results`; it must read as absent,
  // not throw and not look like an empty selection.
  it("returns undefined when the wrapped results field is null or undefined", () => {
    expect(readChoiceArray({ results: null })).toBeUndefined();
    expect(readChoiceArray({ results: undefined })).toBeUndefined();
  });

  it("returns undefined for shapes that are not a choice array", () => {
    expect(readChoiceArray(".pdf,.doc")).toBeUndefined();
    expect(readChoiceArray({ notResults: [".pdf"] })).toBeUndefined();
  });
});

describe("resolveAllowedFileTypes", () => {
  it("resolves a populated selection to configured, normalized", () => {
    expect(resolveAllowedFileTypes(["png", ".PDF"])).toEqual({
      kind: "configured",
      types: [".png", ".pdf"],
    });
  });

  // Spec §3: empty is a hard block, never a silent fallback and never an empty
  // allowlist that some caller might read as permissive.
  it("resolves an empty selection to none", () => {
    expect(resolveAllowedFileTypes([])).toEqual({ kind: "none" });
  });

  it("resolves a selection of only blanks to none", () => {
    expect(resolveAllowedFileTypes(["", "  "])).toEqual({ kind: "none" });
  });

  it("resolves an absent field to unknown, carrying the fallback types", () => {
    expect(resolveAllowedFileTypes(undefined)).toEqual({
      kind: "unknown",
      types: FALLBACK_FILE_TYPES,
    });
  });

  // The two failure states must never be confusable — different people fix them.
  it("distinguishes none from unknown", () => {
    expect(resolveAllowedFileTypes([]).kind).not.toBe(
      resolveAllowedFileTypes(undefined).kind,
    );
  });

  it("never reports configured with an empty type list", () => {
    const resolved = resolveAllowedFileTypes([""]);
    expect(resolved.kind).toBe("none");
    expect(resolved).not.toHaveProperty("types");
  });
});

describe("FALLBACK_FILE_TYPES", () => {
  it("ships fallback types that are already normalized", () => {
    expect(normalizeFileTypes(FALLBACK_FILE_TYPES)).toEqual(FALLBACK_FILE_TYPES);
  });
});

describe("messages", () => {
  it("tells the client which column and list to fix", () => {
    expect(NO_TYPES_MESSAGE).toContain("Allowed File Types");
    expect(NO_TYPES_MESSAGE).toContain("DMS Config");
  });

  it("uses a different message for an unreadable config", () => {
    expect(CONFIG_UNREADABLE_MESSAGE).not.toBe(NO_TYPES_MESSAGE);
    expect(CONFIG_UNREADABLE_MESSAGE).toContain("DMS Config");
    expect(CONFIG_UNREADABLE_MESSAGE).toContain("built-in");
  });
});

describe("readAllowedFileTypesField", () => {
  // Regression, verified against the live list 2026-07-30: unticking every choice
  // makes SharePoint send null, NOT []. Reading the value alone made "emptied"
  // indistinguishable from "absent", so the hard block was unreachable and an
  // emptied column silently fell back to the built-in types.
  it("treats a present-but-null field as an empty selection, not as absent", () => {
    expect(readAllowedFileTypesField({ AllowedFileTypes: null })).toEqual([]);
  });

  it("treats a missing key as absent", () => {
    expect(readAllowedFileTypesField({})).toBeUndefined();
  });

  // The two must resolve to different states — none blocks, unknown falls back.
  it("distinguishes an emptied column from an absent one", () => {
    expect(
      resolveAllowedFileTypes(readAllowedFileTypesField({ AllowedFileTypes: null })).kind,
    ).toBe("none");
    expect(resolveAllowedFileTypes(readAllowedFileTypesField({})).kind).toBe(
      "unknown",
    );
  });

  it("reads a populated selection", () => {
    expect(
      readAllowedFileTypesField({ AllowedFileTypes: [".pdf", ".doc"] }),
    ).toEqual([".pdf", ".doc"]);
  });

  it("reads the verbose wrapped shape", () => {
    expect(
      readAllowedFileTypesField({ AllowedFileTypes: { results: [".pdf"] } }),
    ).toEqual([".pdf"]);
  });

  it("still reads a literal empty array as an empty selection", () => {
    expect(readAllowedFileTypesField({ AllowedFileTypes: [] })).toEqual([]);
  });

  // An unparseable shape must not take every upload down site-wide.
  it("treats an unrecognised shape as absent, not as empty", () => {
    expect(
      readAllowedFileTypesField({ AllowedFileTypes: ".pdf,.doc" }),
    ).toBeUndefined();
  });
});
