import { pickFullNameField, SpFieldLite } from "./folderFullName";

const field = (over: Partial<SpFieldLite>): SpFieldLite => ({
  Title: "Full Name",
  InternalName: "Full_x0020_Name",
  TypeAsString: "Text",
  ReadOnlyField: false,
  ...over,
});

describe("pickFullNameField", () => {
  it("returns the internal name, which is not derivable from the display name", () => {
    // The whole reason this function exists: the same display name can encode either way
    // depending on how the column was created. Both are real, so we read, never guess.
    expect(pickFullNameField([field({})])).toBe("Full_x0020_Name");
    expect(pickFullNameField([field({ InternalName: "FullName" })])).toBe("FullName");
  });

  it("returns undefined when the column is absent", () => {
    // Soft state — reconciliation carries on without writing a full name rather than
    // aborting a run whose real job is folders and ACLs.
    expect(pickFullNameField([])).toBeUndefined();
    expect(pickFullNameField([field({ Title: "Title", InternalName: "Title" })])).toBeUndefined();
  });

  it("ignores case and surrounding whitespace in the display name", () => {
    expect(pickFullNameField([field({ Title: " full name " })])).toBe("Full_x0020_Name");
  });

  it("rejects a read-only field of the same name", () => {
    // A calculated column would accept the MERGE and silently discard the value, so it
    // must not be selected at all — a silent no-op is worse than no column.
    expect(pickFullNameField([field({ ReadOnlyField: true })])).toBeUndefined();
  });

  it("rejects non-text types", () => {
    // A Choice column cannot hold an arbitrary term label; picking it would turn a
    // config mistake into a per-folder write failure repeated hundreds of times.
    expect(pickFullNameField([field({ TypeAsString: "Choice" })])).toBeUndefined();
    expect(pickFullNameField([field({ TypeAsString: "Lookup" })])).toBeUndefined();
  });

  it("accepts multiple lines of text", () => {
    expect(pickFullNameField([field({ TypeAsString: "Note", InternalName: "FullName" })])).toBe("FullName");
  });

  it("prefers single-line text when a site has both", () => {
    expect(
      pickFullNameField([
        field({ TypeAsString: "Note", InternalName: "FullNameNote" }),
        field({ TypeAsString: "Text", InternalName: "Full_x0020_Name" }),
      ]),
    ).toBe("Full_x0020_Name");
  });
});
