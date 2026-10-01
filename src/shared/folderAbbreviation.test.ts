import {
  buildAbbrevIndex,
  lookupAbbrev,
  findCollisions,
  planOrphanRepairs,
  AbbrevRow,
  OrphanAbbrevRow,
  UnclaimedTerm,
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

describe("planOrphanRepairs", () => {
  const orphan = (
    itemId: number,
    title: string,
    level: string,
    abbreviation = "AB",
  ): OrphanAbbrevRow => ({
    itemId,
    termGuid: `dead-${itemId}`,
    title,
    level,
    abbreviation,
  });
  const term = (termGuid: string, label: string, level: string): UnclaimedTerm => ({
    termGuid,
    label,
    level,
  });

  it("repairs an orphan with exactly one match in both directions", () => {
    const o = orphan(1, "Strategy Partner PNG ＆ SI", "Unit", "SP_PNGSI");
    const t = term("new-guid", "Strategy Partner PNG ＆ SI", "Unit");
    const plan = planOrphanRepairs([o], [t]);
    expect(plan.repairs).toEqual([{ orphan: o, term: t }]);
    expect(plan.ambiguous).toEqual([]);
    expect(plan.unmatched).toEqual([]);
  });

  it("does not match across levels — a Department never repairs from a Unit", () => {
    const o = orphan(1, "Legal", "Department");
    const plan = planOrphanRepairs([o], [term("new", "Legal", "Unit")]);
    expect(plan.repairs).toEqual([]);
    expect(plan.unmatched).toEqual([o]);
  });

  // The reason a bare label is not enough: "Tax" exists under GHO/GF, MHO/GA and
  // NBPOLHO/FIN. Re-pointing the wrong one propagates to the Group Map, which is a
  // permissions grant to another department's folder.
  it("refuses to choose when two orphans share a title and level", () => {
    const a = orphan(1, "Tax", "Unit");
    const b = orphan(2, "Tax", "Unit");
    const t = term("new", "Tax", "Unit");
    const plan = planOrphanRepairs([a, b], [t]);
    expect(plan.repairs).toEqual([]);
    expect(plan.ambiguous.length).toBe(2);
    expect(plan.ambiguous[0].orphanCandidates).toEqual([a, b]);
    expect(plan.ambiguous[0].termCandidates).toEqual([t]);
  });

  it("refuses to choose when two live terms share a label and level", () => {
    const o = orphan(1, "Tax", "Unit");
    const t1 = term("new-1", "Tax", "Unit");
    const t2 = term("new-2", "Tax", "Unit");
    const plan = planOrphanRepairs([o], [t1, t2]);
    expect(plan.repairs).toEqual([]);
    expect(plan.ambiguous).toEqual([
      { orphan: o, orphanCandidates: [o], termCandidates: [t1, t2] },
    ]);
  });

  it("reports an orphan with no candidate term as unmatched — a true deletion", () => {
    const o = orphan(1, "Retired Unit", "Unit");
    const plan = planOrphanRepairs([o], [term("new", "Something Else", "Unit")]);
    expect(plan.repairs).toEqual([]);
    expect(plan.ambiguous).toEqual([]);
    expect(plan.unmatched).toEqual([o]);
  });

  it("compares case-insensitively and trims, on both title and level", () => {
    const o = orphan(1, "  group compliance ", "unit");
    const t = term("new", "Group Compliance", "Unit");
    expect(planOrphanRepairs([o], [t]).repairs).toEqual([{ orphan: o, term: t }]);
  });

  // The term store requires FULLWIDTH ＆; both sides store that form, so it is
  // compared as-is. A halfwidth & is a different character and must not match —
  // treating them as equal would let a hand-typed row claim the wrong term.
  it("does not treat a halfwidth & as a fullwidth ＆", () => {
    const o = orphan(1, "Group Legal, Risk & Compliance", "Department");
    const t = term("new", "Group Legal, Risk ＆ Compliance", "Department");
    const plan = planOrphanRepairs([o], [t]);
    expect(plan.repairs).toEqual([]);
    expect(plan.unmatched).toEqual([o]);
  });

  it("never matches a blank title, even against a blank-labelled term", () => {
    const o = orphan(1, "   ", "Unit");
    const plan = planOrphanRepairs([o], [term("new", "", "Unit")]);
    expect(plan.repairs).toEqual([]);
    expect(plan.unmatched).toEqual([o]);
  });

  it("returns an empty plan when there is nothing to do", () => {
    expect(planOrphanRepairs([], [])).toEqual({
      repairs: [],
      ambiguous: [],
      unmatched: [],
    });
  });

  it("repairs several unrelated orphans in one pass", () => {
    const a = orphan(1, "Treasury", "Unit");
    const b = orphan(2, "Internal Audit", "Unit");
    const ta = term("new-a", "Treasury", "Unit");
    const tb = term("new-b", "Internal Audit", "Unit");
    const plan = planOrphanRepairs([a, b], [tb, ta]);
    expect(plan.repairs).toEqual([
      { orphan: a, term: ta },
      { orphan: b, term: tb },
    ]);
  });
});
