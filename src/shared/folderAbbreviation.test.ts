import {
  buildAbbrevIndex,
  lookupAbbrev,
  findCollisions,
  AbbrevRow,
} from "./folderAbbreviation";

const rows: AbbrevRow[] = [
  { termGuid: "AAA", abbreviation: "GCA" },
  { termGuid: "BBB", abbreviation: "GMB_STRATCOMMS" },
];

describe("buildAbbrevIndex / lookupAbbrev", () => {
  it("looks up by term guid, case-insensitively", () => {
    const ix = buildAbbrevIndex(rows);
    expect(lookupAbbrev(ix, "aaa")).toBe("GCA");
    expect(lookupAbbrev(ix, "AAA")).toBe("GCA");
  });

  it("returns undefined for an unmapped term", () => {
    expect(lookupAbbrev(buildAbbrevIndex(rows), "ZZZ")).toBeUndefined();
  });

  it("ignores rows with a blank guid or blank abbreviation", () => {
    const ix = buildAbbrevIndex([
      { termGuid: "", abbreviation: "X" },
      { termGuid: "CCC", abbreviation: "   " },
    ]);
    expect(lookupAbbrev(ix, "CCC")).toBeUndefined();
  });

  it("trims surrounding whitespace on the abbreviation", () => {
    const ix = buildAbbrevIndex([{ termGuid: "DDD", abbreviation: "  GF  " }]);
    expect(lookupAbbrev(ix, "DDD")).toBe("GF");
  });
});

describe("findCollisions", () => {
  // The failure this whole guard exists to prevent: two units under one parent
  // resolving to the same folder means one ACL over two units' documents.
  // Real case in the client's data — Group Corporate Affairs holds both
  // "Global Marketing & Branding - Strategic Communications" and
  // "Group Communications - Strategic Communications".
  it("reports two siblings sharing an abbreviation", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "STRATCOMMS", label: "GMB - Strategic Communications" },
        { parentPath: "/GHO/GCA", termGuid: "B", abbreviation: "STRATCOMMS", label: "GC - Strategic Communications" },
      ]),
    ).toEqual([
      {
        parentPath: "/GHO/GCA",
        abbreviation: "STRATCOMMS",
        labels: ["GMB - Strategic Communications", "GC - Strategic Communications"],
      },
    ]);
  });

  it("allows the same abbreviation under different parents", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "PM", label: "x" },
        { parentPath: "/MHO/GCA", termGuid: "B", abbreviation: "PM", label: "y" },
      ]),
    ).toEqual([]);
  });

  it("compares case-insensitively — SharePoint folder names are not case-unique", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "cors", label: "x" },
        { parentPath: "/GHO/GCA", termGuid: "B", abbreviation: "CORS", label: "y" },
      ]).length,
    ).toBe(1);
  });

  it("returns an empty array when everything is unique", () => {
    expect(
      findCollisions([
        { parentPath: "/GHO/GCA", termGuid: "A", abbreviation: "IR", label: "x" },
        { parentPath: "/GHO/GCA", termGuid: "B", abbreviation: "YG", label: "y" },
      ]),
    ).toEqual([]);
  });
});
