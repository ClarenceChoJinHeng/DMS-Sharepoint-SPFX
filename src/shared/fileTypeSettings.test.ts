import {
  badgeFor,
  blockedButTicked,
  classifyAddition,
  CUSTOM_TYPE_DESCRIPTION,
  describeExtension,
  isBlockedType,
  nearMiss,
  normalizeExtension,
  panelState,
} from "./fileTypeSettings";

const CHOICES = [".pdf", ".docx", ".xlsx"];
const TICKED = [".pdf", ".xlsx"];

describe("normalizeExtension", () => {
  // The dot is ENFORCED, not required: a dotless token is silently dropped by the file picker as an
  // invalid MIME type while the JS validator accepts it — one typo, two opposite behaviours.
  it("forces a single leading dot and lowercases", () => {
    expect(normalizeExtension("png")).toEqual({ ok: true, ext: ".png" });
    expect(normalizeExtension(".PNG")).toEqual({ ok: true, ext: ".png" });
    expect(normalizeExtension("  .Pdf  ")).toEqual({ ok: true, ext: ".pdf" });
  });

  it("accepts digits, which real extensions carry", () => {
    expect(normalizeExtension("7z")).toEqual({ ok: true, ext: ".7z" });
    expect(normalizeExtension(".mp4")).toEqual({ ok: true, ext: ".mp4" });
  });

  it("rejects blanks, spaces, a second dot and punctuation", () => {
    expect(normalizeExtension("").ok).toBe(false);
    expect(normalizeExtension("   ").ok).toBe(false);
    expect(normalizeExtension(".").ok).toBe(false);
    expect(normalizeExtension("my file").ok).toBe(false);
    expect(normalizeExtension("tar.gz").ok).toBe(false);
    expect(normalizeExtension("pd/f").ok).toBe(false);
  });

  it("explains WHY, since the admin has to fix the input", () => {
    const spaced = normalizeExtension("two words");
    expect(spaced.ok).toBe(false);
    if (!spaced.ok) expect(spaced.reason).toMatch(/spaces/i);
  });
});

describe("describeExtension", () => {
  it("names known types", () => {
    expect(describeExtension(".xlsx")).toBe("Microsoft Excel Workbook");
    expect(describeExtension(".PDF")).toBe("Portable Document Format");
  });

  // The label is the typo signal: a row meant to be Excel that reads "Custom file type" is a warning.
  it("labels anything unknown as custom", () => {
    expect(describeExtension(".xlxs")).toBe(CUSTOM_TYPE_DESCRIPTION);
    expect(describeExtension(".zzz")).toBe(CUSTOM_TYPE_DESCRIPTION);
  });
});

describe("isBlockedType", () => {
  it("blocks executables and scripts, case-insensitively", () => {
    expect(isBlockedType(".exe")).toBe(true);
    expect(isBlockedType(".PS1")).toBe(true);
    expect(isBlockedType(" .bat ")).toBe(true);
  });

  // Client decision 2026-08-13: archives are allowed, which makes the list advisory. Pinned so the
  // decision cannot be quietly reversed by adding one line.
  it("does NOT block archives", () => {
    expect(isBlockedType(".zip")).toBe(false);
    expect(isBlockedType(".7z")).toBe(false);
    expect(isBlockedType(".rar")).toBe(false);
  });

  it("leaves ordinary documents alone", () => {
    expect(isBlockedType(".pdf")).toBe(false);
    expect(isBlockedType(".docx")).toBe(false);
  });
});

describe("nearMiss", () => {
  it("catches a transposed known type", () => {
    expect(nearMiss(".xlxs")).toBe(".xlsx");
  });

  it("says nothing about a type that is already known", () => {
    expect(nearMiss(".pdf")).toBeUndefined();
  });

  it("says nothing about input resembling nothing", () => {
    expect(nearMiss(".zzz")).toBeUndefined();
  });

  // A same-length candidate wins a tie: that is a substitution or transposition, which is how
  // extensions are actually mistyped. `.xlxs` is also one DELETION from `.xls`.
  it("prefers the same-length candidate over a dropped-character one", () => {
    expect(nearMiss(".tiif")).toBe(".tiff");
  });

  // Guessing between two equally plausible substitutions would put a WRONG extension one click away.
  it("withholds a suggestion when two same-length types are one edit away", () => {
    expect(nearMiss(".dox")).toBeUndefined(); // `.doc` and `.dot` are both one substitution away
  });
});

describe("classifyAddition", () => {
  it("treats an unknown, well-formed extension as new", () => {
    expect(classifyAddition("txt", CHOICES, TICKED)).toEqual({ kind: "new", ext: ".txt" });
  });

  it("attaches a suggestion to a new extension that looks like a typo", () => {
    expect(classifyAddition(".xlxs", CHOICES, TICKED)).toEqual({
      kind: "new",
      ext: ".xlxs",
      suggestion: ".xlsx",
    });
  });

  // The verdict that stops a DUPLICATE choice: two `.docx` entries render two identical rows with
  // independent toggles, one of which appears to do nothing.
  it("reports an existing but disabled choice without enabling it", () => {
    expect(classifyAddition("DOCX", CHOICES, TICKED)).toEqual({
      kind: "exists-disabled",
      ext: ".docx",
    });
  });

  it("reports an existing and enabled choice", () => {
    expect(classifyAddition(".pdf", CHOICES, TICKED)).toEqual({
      kind: "exists-enabled",
      ext: ".pdf",
    });
  });

  it("refuses executables before considering anything else", () => {
    expect(classifyAddition(".exe", CHOICES, TICKED)).toEqual({ kind: "blocked", ext: ".exe" });
  });

  it("refuses malformed input with a reason", () => {
    expect(classifyAddition("a b", CHOICES, TICKED).kind).toBe("invalid");
  });

  it("matches existing choices regardless of stored case or padding", () => {
    expect(classifyAddition(".pdf", [" .PDF "], [" .PDF "])).toEqual({
      kind: "exists-enabled",
      ext: ".pdf",
    });
  });
});

describe("panelState", () => {
  // These two look identical in a list view and mean opposite things: a silent default versus a
  // deliberate block. Only the presence of the column separates them.
  it("distinguishes an absent column from an empty selection", () => {
    expect(panelState(false, [])).toBe("no-column");
    expect(panelState(true, [])).toBe("all-blocked");
  });

  it("is normal as soon as one type is ticked", () => {
    expect(panelState(true, [".pdf"])).toBe("normal");
  });

  // An absent column stays `no-column` even if values somehow arrive: there is nothing to edit.
  it("prefers no-column over any ticked values", () => {
    expect(panelState(false, [".pdf"])).toBe("no-column");
  });
});

describe("blockedButTicked", () => {
  it("finds blocked types that are currently allowed", () => {
    expect(blockedButTicked([".pdf", ".EXE", ".docx"])).toEqual([".exe"]);
  });

  it("de-duplicates and returns nothing for a clean list", () => {
    expect(blockedButTicked([".exe", " .exe "])).toEqual([".exe"]);
    expect(blockedButTicked(TICKED)).toEqual([]);
  });
});

describe("badgeFor", () => {
  it("strips the dot, uppercases and truncates", () => {
    expect(badgeFor(".pdf")).toBe("PDF");
    expect(badgeFor(".xlsx")).toBe("XLSX");
    expect(badgeFor(".heic")).toBe("HEIC");
  });
});
