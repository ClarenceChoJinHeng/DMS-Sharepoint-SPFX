import {
  AbbrevRowDraft,
  changedRows,
  folderNameFor,
  hasBlockingProblem,
  LONG_NAME_THRESHOLD,
  renamingRows,
  validateRows,
} from "./abbreviationDraft";

const GF = "aaaa1111-1111-4111-8111-aaaaaaaaaaaa"; // Group Finance
const MHO_GA = "bbbb2222-2222-4222-8222-bbbbbbbbbbbb"; // Minamas Group Accounts

const row = (
  termGuid: string,
  label: string,
  abbreviation: string,
  parentGuid = GF,
  original?: string,
): AbbrevRowDraft => ({ termGuid, label, level: "Unit", parentGuid, abbreviation, original });

describe("folderNameFor", () => {
  it("is the name SharePoint will actually create", () => {
    expect(folderNameFor("CORU")).toBe("CORU");
    expect(folderNameFor("  GC EP  ")).toBe("GC EP");
    // Collapsed whitespace: this is why comparison happens on the folder name, not the raw text.
    expect(folderNameFor("GC  EP")).toBe("GC EP");
  });

  it("is empty when nothing usable survives", () => {
    expect(folderNameFor("///")).toBe("");
    expect(folderNameFor("")).toBe("");
  });
});

describe("validateRows — siblings", () => {
  it("passes distinct codes under one parent", () => {
    const p = validateRows([row("t1", "Compliance", "CORU"), row("t2", "Treasury", "TREASURY")]);
    // A clean row gets NO entry at all, rather than an entry with empty fields — so the UI can
    // treat "present in the map" as "has something to say".
    expect(p.t1).toBeUndefined();
    expect(p.t2).toBeUndefined();
    expect(hasBlockingProblem(p)).toBe(false);
  });

  it("flags BOTH rows of a collision and names the other term", () => {
    const p = validateRows([
      row("t1", "Compliance & Operational Risk (CORU)", "CORU"),
      row("t2", "Corporate Reporting", "CORU"),
    ]);
    expect(p.t1.error).toContain("Corporate Reporting");
    expect(p.t2.error).toContain("Compliance & Operational Risk (CORU)");
    expect(hasBlockingProblem(p)).toBe(true);
  });

  it("catches a collision differing only by CASE — SharePoint folders are case-insensitive", () => {
    const p = validateRows([row("t1", "Compliance", "CORU"), row("t2", "Reporting", "coru")]);
    expect(hasBlockingProblem(p)).toBe(true);
  });

  it("catches a collision differing only by collapsed whitespace", () => {
    // findCollisions lowercases but does not sanitize, so this pair only collides because the page
    // feeds it sanitized names. Being stricter than reconciliation is the safe direction.
    const p = validateRows([row("t1", "Events", "GC EP"), row("t2", "Protocol", "GC  EP")]);
    expect(hasBlockingProblem(p)).toBe(true);
  });

  it("ALLOWS the same code under different parents — Tax repeats across departments", () => {
    const p = validateRows([row("t1", "Tax", "TAX", GF), row("t2", "Tax", "TAX", MHO_GA)]);
    expect(p.t1).toBeUndefined();
    expect(p.t2).toBeUndefined();
    expect(hasBlockingProblem(p)).toBe(false);
  });

  it("names all clashing siblings when three share a code", () => {
    const p = validateRows([row("t1", "One", "X"), row("t2", "Two", "X"), row("t3", "Three", "X")]);
    expect(p.t1.error).toContain("Two");
    expect(p.t1.error).toContain("Three");
  });
});

describe("validateRows — missing and unusable", () => {
  it("WARNS on a blank code, saying the unit cannot upload", () => {
    const p = validateRows([row("t1", "Binuang Estate", "")]);
    expect(p.t1.warn).toContain("No folder will be created");
    expect(p.t1.error).toBeUndefined();
    // A blank is a legitimate in-progress state, so it must not block the save.
    expect(hasBlockingProblem(p)).toBe(false);
  });

  it("does not treat two blanks as colliding with each other", () => {
    const p = validateRows([row("t1", "A", ""), row("t2", "B", "")]);
    expect(hasBlockingProblem(p)).toBe(false);
  });

  it("ERRORS when every character typed would be stripped", () => {
    // Looks filled in, creates nothing — indistinguishable from blank in the folder tree.
    const p = validateRows([row("t1", "Legal", "///")]);
    expect(p.t1.error).toContain("nothing usable");
    expect(hasBlockingProblem(p)).toBe(true);
  });
});

describe("validateRows — length", () => {
  it("warns above the threshold without blocking", () => {
    const long = "Global Marketing and Branding";
    expect(long.length).toBeGreaterThan(LONG_NAME_THRESHOLD);
    const p = validateRows([row("t1", "GMB", long)]);
    expect(p.t1.warn).toContain(String(long.length));
    expect(hasBlockingProblem(p)).toBe(false);
  });

  it("says nothing about a short code", () => {
    expect(validateRows([row("t1", "Perak", "Perak")]).t1).toBeUndefined();
  });

  it("reports a collision rather than the length when a long code also clashes", () => {
    // The collision blocks the save; the length only advises. Reporting the advice would bury it.
    const long = "Global Marketing and Branding";
    const p = validateRows([row("t1", "A", long), row("t2", "B", long)]);
    expect(p.t1.error).toBeDefined();
    expect(p.t1.warn).toBeUndefined();
  });
});

describe("changedRows / renamingRows", () => {
  const rows = [
    row("t1", "Untouched", "CORU", GF, "CORU"),
    row("t2", "Filled in", "TAX", GF, ""),
    row("t3", "Renamed", "TREAS", GF, "TREASURY"),
    row("t4", "Cleared", "", GF, "OLD"),
  ];

  it("returns only the rows whose code actually differs", () => {
    expect(changedRows(rows).map((r) => r.termGuid)).toEqual(["t2", "t3", "t4"]);
  });

  it("ignores whitespace-only differences", () => {
    expect(changedRows([row("t1", "A", "  CORU  ", GF, "CORU")])).toEqual([]);
  });

  it("separates a RENAME from a first-time fill", () => {
    // t3 and t4 had a code and now have a different one, so a live folder gets renamed. t2 is new.
    expect(renamingRows(rows).map((r) => r.termGuid)).toEqual(["t3", "t4"]);
  });
});
