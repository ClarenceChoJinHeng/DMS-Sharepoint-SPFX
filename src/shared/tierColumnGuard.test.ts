import {
  existingColumnReason,
  reservedColumnReason,
  RESERVED_TIER_COLUMNS,
  tierColumnRefusal,
} from "./tierColumnGuard";

describe("reservedColumnReason", () => {
  it("refuses a name that belongs to a field the upload form writes", () => {
    const why = reservedColumnReason("Remark");
    expect(why).toContain("Remark");
    expect(why).toContain("Give the level a different name");
  });

  /* ⚠ BOTH SPELLINGS, and this is the case a single entry would miss. `columnNameFor` strips illegal
     characters, so a level typed "Vendor/Customer Name" derives `VendorCustomerName` — while the
     real column is `Vendor_x002f_CustomerName`. Neither may be used. */
  it("catches an encoded column under its derived spelling too", () => {
    expect(reservedColumnReason("VendorCustomerName")).toBeDefined();
    expect(reservedColumnReason("Vendor_x002f_CustomerName")).toBeDefined();
    expect(reservedColumnReason("Full_x0020_Name")).toBeDefined();
    expect(reservedColumnReason("FullName")).toBeDefined();
  });

  it("is case-insensitive and ignores surrounding space", () => {
    expect(reservedColumnReason("  keyword  ")).toBeDefined();
    expect(reservedColumnReason("APPROVEDBY")).toBeDefined();
  });

  it("allows an ordinary level name", () => {
    expect(reservedColumnReason("BuahNumber")).toBeUndefined();
    expect(reservedColumnReason("State")).toBeUndefined();
    expect(reservedColumnReason("SubUnit")).toBeUndefined();
  });

  /* The submission stamp and the bulk-import marker are the two that read as harmless — they are
     invisible on every screen — and are exactly the ones whose loss would be hardest to trace:
     grouping silently stops, or every bulk import stops auto-approving. */
  it("covers the invisible bookkeeping columns", () => {
    for (const n of ["SubmissionId", "BatchId", "SubmissionFileId", "BulkImport", "Archived"]) {
      expect(reservedColumnReason(n)).toBeDefined();
    }
  });
});

describe("existingColumnReason", () => {
  it("allows a NEW column — nothing to collide with", () => {
    expect(existingColumnReason("BuahNumber", undefined)).toBeUndefined();
  });

  /* Sharing a plain-text tier column is how `Department` and `Unit` work across every segment, and
     how a removed level is restored under its own name. Refusing it would break both. */
  it("allows binding to an existing plain-text column", () => {
    expect(existingColumnReason("State", "Text")).toBeUndefined();
    expect(existingColumnReason("State", "text")).toBeUndefined();
  });

  /* THE ONE THAT MATTERS. A taxonomy field wants `Label|GUID`; a tier writer sends a bare folder
     name. Nothing errors at save — the segment simply stops tagging, which is the 2026-08-20
     defect. */
  it("refuses a taxonomy column, naming the type", () => {
    const why = existingColumnReason("Year", "TaxonomyFieldType");
    expect(why).toContain("TaxonomyFieldType");
    expect(why).toContain("plain text");
  });

  it("refuses every other type, including Note", () => {
    for (const t of ["Note", "DateTime", "Boolean", "Lookup", "Choice", "User"]) {
      expect(existingColumnReason("Something", t)).toBeDefined();
    }
  });
});

describe("tierColumnRefusal", () => {
  /* Reserved is answered first so the verdict is identical on every site — those names are wrong
     whether or not the column has been created yet. */
  it("reports a reserved name even when the column does not exist", () => {
    expect(tierColumnRefusal("Keyword", undefined)).toBeDefined();
  });

  it("reports a reserved name in preference to its type", () => {
    const why = tierColumnRefusal("Remark", "Note");
    expect(why).toContain("already used for");
  });

  it("passes an ordinary new level", () => {
    expect(tierColumnRefusal("BuahNumber", undefined)).toBeUndefined();
  });

  it("passes an ordinary level sharing an existing text column", () => {
    expect(tierColumnRefusal("State", "Text")).toBeUndefined();
  });

  /* ⚠ EVERY RESERVED ENTRY MUST CARRY A REASON. The message names the job the column already does,
     and a blank one would render as "is already used for ." — pinned because the map is
     hand-written and will be added to. */
  it("gives every reserved column a stated job", () => {
    const names = Object.keys(RESERVED_TIER_COLUMNS);
    expect(names.length).toBeGreaterThan(0);
    for (const n of names) expect(RESERVED_TIER_COLUMNS[n].trim().length).toBeGreaterThan(0);
  });
});
