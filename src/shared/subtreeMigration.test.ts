import {
  assignSegments,
  backfillNeeds,
  classifyChild,
  effectiveTiers,
  findCollisions,
  LeafPlan,
  planLeaf,
  planTotals,
  planUnit,
  suggestRename,
  validateRename,
  EffectiveTier,
} from "./subtreeMigration";

const UNIT = "/sites/S/ApprovalDocument/GHO/GF/CORU";

/** Chain: 0 = SubUnit (cascading), 1 = Year, 2 = Document Type. */
const TIERS: EffectiveTier[] = [
  { chainIndex: 0, options: ["testig", "Finance"] },
  { chainIndex: 1, options: ["2024", "2025"] },
  { chainIndex: 2, options: ["Tax Return", "Invoice"] },
];

describe("effectiveTiers", () => {
  it("keeps chain indexes when a middle tier does not apply", () => {
    // The unit has no SubUnits: tier 0 drops out, but Year must still know it is chain
    // index 1, or its destination would be written to the wrong column.
    const tiers = effectiveTiers([[], ["2024"], ["Invoice"]]);
    expect(tiers.map((t) => t.chainIndex)).toEqual([1, 2]);
  });

  it("drops every tier when nothing resolved", () => {
    expect(effectiveTiers([[], []])).toEqual([]);
  });

  it("tolerates holes in the input", () => {
    expect(effectiveTiers([undefined as unknown as string[], ["2024"]]).length).toBe(1);
  });
});

describe("classifyChild", () => {
  it("calls a valid first-tier name correctly placed", () => {
    expect(classifyChild("testig", TIERS)).toEqual({ kind: "ok" });
  });

  it("reports how many levels deep a misplaced folder belongs", () => {
    expect(classifyChild("2024", TIERS)).toEqual({ kind: "misplaced", levels: 1 });
    expect(classifyChild("Invoice", TIERS)).toEqual({ kind: "misplaced", levels: 2 });
  });

  it("prefers 'already correct' when a name is valid at two depths", () => {
    // The safe direction: moving a correctly filed folder needs evidence, and a name valid
    // at tier 0 is evidence that it is where it belongs.
    const ambiguous: EffectiveTier[] = [
      { chainIndex: 0, options: ["2026"] },
      { chainIndex: 1, options: ["2026"] },
    ];
    expect(classifyChild("2026", ambiguous)).toEqual({ kind: "ok" });
  });

  it("matches case-insensitively", () => {
    expect(classifyChild("TAX RETURN", TIERS)).toEqual({ kind: "misplaced", levels: 2 });
  });

  it("compares against the sanitized label, not the raw one", () => {
    // Folder names are written sanitized, so an option carrying an illegal character must
    // still match the folder created from it — otherwise a correctly filed folder reads as
    // a stray and an admin is shown a problem that does not exist.
    const tiers: EffectiveTier[] = [{ chainIndex: 0, options: ["Legal: Tax"] }];
    expect(classifyChild("Legal Tax", tiers)).toEqual({ kind: "ok" });
  });

  it("calls an unknown folder a stray", () => {
    expect(classifyChild("Old Stuff", TIERS)).toEqual({ kind: "stray" });
  });

  it("calls everything a stray when there are no tiers", () => {
    expect(classifyChild("2024", [])).toEqual({ kind: "stray" });
  });
});

describe("planUnit", () => {
  const dest = { label: "testig", id: "t-1" };

  it("moves a year folder under the chosen destination, keeping its name", () => {
    const plan = planUnit(UNIT, ["2024", "2025"], TIERS, [dest]);
    expect(plan.skipped).toBeUndefined();
    expect(plan.moves.map((m) => m.from)).toEqual([`${UNIT}/2024`, `${UNIT}/2025`]);
    expect(plan.moves.map((m) => m.to)).toEqual([`${UNIT}/testig/2024`, `${UNIT}/testig/2025`]);
    expect(plan.moves[0].ancestors).toEqual([
      { path: `${UNIT}/testig`, chainIndex: 0, destination: dest },
    ]);
  });

  it("leaves correctly placed folders alone", () => {
    const plan = planUnit(UNIT, ["testig"], TIERS, [dest]);
    expect(plan.moves).toEqual([]);
    expect(plan.strays).toEqual([]);
  });

  it("reports strays without moving them and without blocking the rest", () => {
    const plan = planUnit(UNIT, ["2024", "Old Stuff"], TIERS, [dest]);
    expect(plan.strays).toEqual(["Old Stuff"]);
    expect(plan.moves.length).toBe(1);
  });

  it("builds two ancestor folders when a folder is adrift by two levels", () => {
    const d0 = { label: "testig", id: "t-1" };
    const d1 = { label: "2024", id: "y-1" };
    const plan = planUnit(UNIT, ["Invoice"], TIERS, [d0, d1]);
    expect(plan.moves[0].to).toBe(`${UNIT}/testig/2024/Invoice`);
    expect(plan.moves[0].ancestors.map((a) => a.path)).toEqual([
      `${UNIT}/testig`,
      `${UNIT}/testig/2024`,
    ]);
    expect(plan.moves[0].ancestors.map((a) => a.chainIndex)).toEqual([0, 1]);
  });

  it("uses a shorter prefix for shallower folders in the same unit", () => {
    // Two structure edits before one migration: Invoice is 2 deep and 2024 is 1 deep, and
    // each must land at its own depth rather than all at the deepest.
    const plan = planUnit(UNIT, ["2024", "Invoice"], TIERS, [
      { label: "testig", id: "t-1" },
      { label: "2024", id: "y-1" },
    ]);
    expect(plan.moves.filter((m) => m.name === "2024")[0].to).toBe(`${UNIT}/testig/2024`);
    expect(plan.moves.filter((m) => m.name === "Invoice")[0].to).toBe(
      `${UNIT}/testig/2024/Invoice`,
    );
  });

  it("skips the unit when a destination is missing rather than defaulting", () => {
    const plan = planUnit(UNIT, ["2024"], TIERS, []);
    expect(plan.skipped).toBe("no destination chosen");
    expect(plan.moves).toEqual([]);
  });

  it("still reports which tiers need a destination when it skips", () => {
    // The UI renders one picker per needed tier, so this must survive the skip.
    const plan = planUnit(UNIT, ["Invoice"], TIERS, []);
    expect(plan.neededTiers).toEqual([0, 1]);
  });

  it("needs no destination and stays quiet when nothing is misplaced", () => {
    const plan = planUnit(UNIT, ["testig"], TIERS, []);
    expect(plan.skipped).toBeUndefined();
    expect(plan.neededTiers).toEqual([]);
  });

  it("refuses a destination whose sanitized name is empty", () => {
    // Would build `unit//2024`, which SharePoint collapses to `unit/2024`: a move that does
    // nothing and reports success.
    const plan = planUnit(UNIT, ["2024"], TIERS, [{ label: '###"', id: "x" }]);
    expect(plan.moves).toEqual([]);
    expect(plan.skipped).toMatch(/not a usable folder name/);
  });

  it("sanitizes the destination folder name", () => {
    const plan = planUnit(UNIT, ["2024"], TIERS, [{ label: "R&D: Core", id: "x" }]);
    expect(plan.moves[0].to).toBe(`${UNIT}/R&D Core/2024`);
  });

  it("does nothing for a unit with no children", () => {
    expect(planUnit(UNIT, [], TIERS, [dest]).moves).toEqual([]);
  });
});

describe("backfillNeeds", () => {
  const col = (i: number): string | undefined =>
    ({ 0: "Testing", 1: "Year", 2: "Document_x0020_Type" } as Record<number, string>)[i];

  it("flags a file whose tier column is blank", () => {
    const needs = backfillNeeds(
      UNIT,
      [
        {
          path: `${UNIT}/testig/2024/Tax Return/a.pdf`,
          values: { Year: "2024", Document_x0020_Type: "Tax Return" },
        },
      ],
      TIERS,
      col,
    );
    expect(needs.length).toBe(1);
    expect(needs[0].fields).toEqual([{ chainIndex: 0, label: "testig" }]);
  });

  it("trusts the path over a column that disagrees", () => {
    const needs = backfillNeeds(
      UNIT,
      [
        {
          path: `${UNIT}/testig/2024/Tax Return/a.pdf`,
          values: { Testing: "Finance", Year: "2024", Document_x0020_Type: "Tax Return" },
        },
      ],
      TIERS,
      col,
    );
    expect(needs[0].fields).toEqual([{ chainIndex: 0, label: "testig" }]);
  });

  it("returns nothing when every column already matches the path", () => {
    const needs = backfillNeeds(
      UNIT,
      [
        {
          path: `${UNIT}/testig/2024/Tax Return/a.pdf`,
          values: { Testing: "testig", Year: "2024", Document_x0020_Type: "Tax Return" },
        },
      ],
      TIERS,
      col,
    );
    expect(needs).toEqual([]);
  });

  it("ignores a path segment that matches no term", () => {
    // A stray folder. planUnit already reports it; writing its name into the column would
    // record a value no term backs.
    const needs = backfillNeeds(UNIT, [{ path: `${UNIT}/Old Stuff/a.pdf`, values: {} }], TIERS, col);
    expect(needs).toEqual([]);
  });

  it("infers nothing for a file whose folder belongs to a deeper tier", () => {
    // `2024` sits at position 0, which is SubUnit — it is not a SubUnit value, so nothing
    // is inferred. Relocating this file's folder is the move plan's job, not the stamp's.
    const needs = backfillNeeds(UNIT, [{ path: `${UNIT}/2024/a.pdf`, values: {} }], TIERS, col);
    expect(needs).toEqual([]);
  });

  it("ignores a file directly under the unit", () => {
    expect(backfillNeeds(UNIT, [{ path: `${UNIT}/loose.pdf`, values: {} }], TIERS, col)).toEqual([]);
  });

  it("ignores files outside the unit", () => {
    const needs = backfillNeeds(
      UNIT,
      [{ path: "/sites/S/ApprovalDocument/GHO/GF/OTHER/testig/2024/Tax Return/a.pdf", values: {} }],
      TIERS,
      col,
    );
    expect(needs).toEqual([]);
  });

  it("skips a tier with no label column configured", () => {
    const needs = backfillNeeds(
      UNIT,
      [{ path: `${UNIT}/testig/2024/Tax Return/a.pdf`, values: {} }],
      TIERS,
      (i) => (i === 0 ? undefined : col(i)),
    );
    expect(needs[0].fields.map((f) => f.chainIndex)).toEqual([1, 2]);
  });

  it("stamps only the tiers the caller is willing to write", () => {
    // The real incident, 2026-08-11: Year and Document Type are MANAGED METADATA, and the
    // caller signals that by returning undefined for them (they carry no tidCol). Writing a bare
    // label to a taxonomy field fails with "the data returned from the tagging UI was not
    // formatted correctly" — and since one bad field fails the whole validateUpdateListItem
    // call, it took a perfectly valid write down with it. Reading such a field back as a string
    // also yields "" (the value is an object), so every file looked like it needed a stamp it
    // did not need.
    const needs = backfillNeeds(
      UNIT,
      [{ path: `${UNIT}/testig/2024/Tax Return/a.pdf`, values: {} }],
      TIERS,
      (i) => (i === 0 ? "Testing" : undefined),
    );
    expect(needs[0].fields.map((f) => f.chainIndex)).toEqual([0]);
  });

  it("maps positions to the effective tiers, not the chain, when a tier does not apply", () => {
    // This unit has no SubUnits, so its first folder is Year. Indexing against the full
    // chain instead would read "2024" as a SubUnit and stamp the wrong column.
    const tiers = effectiveTiers([[], ["2024"], ["Tax Return"]]);
    const needs = backfillNeeds(
      UNIT,
      [{ path: `${UNIT}/2024/Tax Return/a.pdf`, values: {} }],
      tiers,
      col,
    );
    expect(needs[0].fields).toEqual([
      { chainIndex: 1, label: "2024" },
      { chainIndex: 2, label: "Tax Return" },
    ]);
  });
});

/* ---- the unified model: add, reorder and remove are one operation (spec §3.0) ---- */

/** Target chain: 0 = Credit_Card, 1 = Testing, 2 = Year, 3 = Document Type. */
const CHAIN: EffectiveTier[] = [
  { chainIndex: 0, options: ["Credit1", "Credit2"] },
  { chainIndex: 1, options: ["testig", "live"] },
  { chainIndex: 2, options: ["2024", "2025"] },
  { chainIndex: 3, options: ["Tax Return", "Invoice"] },
];

const leaf = (segments: string[], files: string[]): { path: string; segments: string[]; files: string[] } => ({
  path: `${UNIT}/${segments.join("/")}`,
  segments,
  files,
});

describe("assignSegments", () => {
  it("keeps a segment at the tier it already occupies when several would accept it", () => {
    // `2026` could be a SubUnit or a Year. The reading that requires no movement is the safe one.
    const tiers: EffectiveTier[] = [
      { chainIndex: 0, options: ["2026"] },
      { chainIndex: 1, options: ["2026"] },
    ];
    expect(assignSegments(["x", "2026"], tiers).assigned).toEqual([{ at: 1, chainIndex: 1, name: "2026" }]);
  });

  it("resolves a reordered path to the tiers the names belong to", () => {
    const { assigned, strays, dropped } = assignSegments(["Credit2", "2024", "testig"], CHAIN);
    expect(assigned).toEqual([
      { at: 0, chainIndex: 0, name: "Credit2" },
      { at: 1, chainIndex: 2, name: "2024" },
      { at: 2, chainIndex: 1, name: "testig" },
    ]);
    expect(strays).toEqual([]);
    expect(dropped).toEqual([]);
  });

  it("separates a removed tier's value from a folder nobody recognises", () => {
    // The distinction decides between collapsing deliberately and relocating someone's documents
    // on a guess, so it must come from data rather than from position.
    const withoutTesting = CHAIN.filter((t) => t.chainIndex !== 1);
    const r = assignSegments(["Credit2", "testig", "Old Stuff"], withoutTesting, [["testig", "live"]]);
    expect(r.dropped).toEqual(["testig"]);
    expect(r.strays).toEqual(["Old Stuff"]);
  });

  it("treats an unknown segment as a stray when no removed tier claims it", () => {
    expect(assignSegments(["Credit2", "mystery"], CHAIN).strays).toEqual(["mystery"]);
  });
});

describe("planLeaf — reorder", () => {
  it("re-nests a swapped pair", () => {
    const plan = planLeaf(UNIT, leaf(["Credit2", "2024", "testig", "Tax Return"], ["a.pdf"]), CHAIN, {});
    expect(plan.to).toBe(`${UNIT}/Credit2/testig/2024/Tax Return`);
    expect(plan.missingTiers).toEqual([]);
  });

  it("names every ancestor of the rebuilt path", () => {
    const plan = planLeaf(UNIT, leaf(["Credit2", "2024", "testig", "Tax Return"], ["a.pdf"]), CHAIN, {});
    expect(plan.ancestors.map((a) => a.path)).toEqual([
      `${UNIT}/Credit2`,
      `${UNIT}/Credit2/testig`,
      `${UNIT}/Credit2/testig/2024`,
    ]);
  });

  it("leaves an already-correct path exactly where it is", () => {
    const l = leaf(["Credit2", "testig", "2024", "Tax Return"], ["a.pdf"]);
    expect(planLeaf(UNIT, l, CHAIN, {}).to).toBe(l.path);
  });
});

describe("planLeaf — add", () => {
  it("asks for the missing tier rather than guessing one", () => {
    const plan = planLeaf(UNIT, leaf(["testig", "2024", "Tax Return"], ["a.pdf"]), CHAIN, {});
    expect(plan.missingTiers).toEqual([0]);
    expect(plan.to).toBeUndefined();
  });

  it("fills the gap from the chosen destination", () => {
    const plan = planLeaf(UNIT, leaf(["testig", "2024", "Tax Return"], ["a.pdf"]), CHAIN, {
      0: { label: "Credit1", id: "c-1" },
    });
    expect(plan.to).toBe(`${UNIT}/Credit1/testig/2024/Tax Return`);
  });

  it("extends a shallow leaf DOWN to the full chain depth", () => {
    // Client rule 2026-08-12: a document always lands at the end of the chain. The previous cap at
    // the leaf's own depth is what made a level added at the BOTTOM a silent no-op.
    const plan = planLeaf(UNIT, leaf(["Credit2", "testig"], ["a.pdf"]), CHAIN, {});
    expect(plan.missingTiers).toEqual([2, 3]);
    expect(plan.to).toBeUndefined();
  });

  it("places a shallow leaf once the deeper values are chosen", () => {
    const plan = planLeaf(UNIT, leaf(["Credit2", "testig"], ["a.pdf"]), CHAIN, {
      2: { label: "2025", id: "y-2" },
      3: { label: "Invoice", id: "d-2" },
    });
    expect(plan.to).toBe(`${UNIT}/Credit2/testig/2025/Invoice`);
  });

  it("treats an unusable chosen name as no choice at all", () => {
    // Would build `unit//2024`, which SharePoint collapses — a move that does nothing and reports
    // success.
    const plan = planLeaf(UNIT, leaf(["testig", "2024", "Tax Return"], ["a.pdf"]), CHAIN, {
      0: { label: '###"', id: "x" },
    });
    expect(plan.missingTiers).toEqual([0]);
    expect(plan.to).toBeUndefined();
  });
});

describe("planLeaf — a level appended at the BOTTOM", () => {
  /** CHAIN plus a fifth tier below Document Type. */
  const withStage: EffectiveTier[] = CHAIN.concat([{ chainIndex: 4, options: ["S1", "S2"] }]);
  const full = ["Credit2", "testig", "2024", "Tax Return"];

  it("asks for the new deepest value instead of reporting nothing to do", () => {
    // The bug this rule fixes: the new tier sat below every leaf, so nothing looked misplaced, the
    // structure activated anyway, and old documents stayed a level shallower than new ones.
    const plan = planLeaf(UNIT, leaf(full, ["a.pdf"]), withStage, {});
    expect(plan.missingTiers).toEqual([4]);
    expect(plan.to).toBeUndefined();
  });

  it("moves the documents down into the chosen value", () => {
    const plan = planLeaf(UNIT, leaf(full, ["a.pdf"]), withStage, { 4: { label: "S1", id: "s-1" } });
    expect(plan.to).toBe(`${UNIT}/Credit2/testig/2024/Tax Return/S1`);
    // The old leaf becomes an ancestor rather than the bottom.
    expect(plan.ancestors.map((a) => a.path).slice(-1)).toEqual([
      `${UNIT}/Credit2/testig/2024/Tax Return`,
    ]);
  });
});

describe("planLeaf — remove", () => {
  const withoutTesting = CHAIN.filter((t) => t.chainIndex !== 1);
  const removed = [["testig", "live"]];

  it("drops the removed tier from the path so siblings collapse together", () => {
    const a = planLeaf(UNIT, leaf(["Credit2", "testig", "2024", "Tax Return"], ["a.pdf"]), withoutTesting, {}, removed);
    const b = planLeaf(UNIT, leaf(["Credit2", "live", "2024", "Tax Return"], ["a.pdf"]), withoutTesting, {}, removed);
    expect(a.to).toBe(`${UNIT}/Credit2/2024/Tax Return`);
    expect(b.to).toBe(`${UNIT}/Credit2/2024/Tax Return`);
    expect(a.dropped).toEqual(["testig"]);
  });

  it("refuses to move a folder holding an unrecognised segment", () => {
    const plan = planLeaf(UNIT, leaf(["Credit2", "mystery", "2024", "Tax Return"], ["a.pdf"]), withoutTesting, {}, removed);
    expect(plan.to).toBeUndefined();
    expect(plan.strays).toEqual(["mystery"]);
  });

  it("leaves a folder alone when every segment belonged to the removed tier", () => {
    // Nothing recognisable is left, so the only destination would be the unit root — which
    // contradicts "documents land at the end of the chain". Reported, never moved.
    const plan = planLeaf(UNIT, leaf(["testig"], ["a.pdf"]), withoutTesting, {}, removed);
    expect(plan.to).toBeUndefined();
    expect(plan.dropped).toEqual(["testig"]);
    expect(plan.strays).toEqual([]);
  });
});

describe("findCollisions", () => {
  const withoutTesting = CHAIN.filter((t) => t.chainIndex !== 1);
  const removed = [["testig", "live"]];
  const collapse = (files: string[][]): LeafPlan[] => [
    planLeaf(UNIT, leaf(["Credit2", "testig", "2024", "Tax Return"], files[0]), withoutTesting, {}, removed),
    planLeaf(UNIT, leaf(["Credit2", "live", "2024", "Tax Return"], files[1]), withoutTesting, {}, removed),
  ];

  it("reports two files of the same name landing in one folder", () => {
    const found = findCollisions(collapse([["a.pdf"], ["a.pdf"]]), {});
    expect(found.length).toBe(1);
    expect(found[0].folder).toBe(`${UNIT}/Credit2/2024/Tax Return`);
    expect(found[0].claimants.map((c) => c.path)).toEqual([
      `${UNIT}/Credit2/live/2024/Tax Return/a.pdf`,
      `${UNIT}/Credit2/testig/2024/Tax Return/a.pdf`,
    ]);
  });

  it("says nothing when the names differ", () => {
    expect(findCollisions(collapse([["a.pdf"], ["b.pdf"]]), {})).toEqual([]);
  });

  it("counts a file ALREADY at the destination as a claimant", () => {
    // The most destructive case and the easiest to miss: the occupant is not migrating, so it
    // appears nowhere in the plan, yet it owns the name and would be overwritten.
    const found = findCollisions(collapse([["a.pdf"], ["b.pdf"]]), {
      [`${UNIT}/Credit2/2024/Tax Return`]: ["a.pdf"],
    });
    expect(found.length).toBe(1);
    expect(found[0].claimants.filter((c) => c.existing).map((c) => c.path)).toEqual([
      `${UNIT}/Credit2/2024/Tax Return/a.pdf`,
    ]);
  });

  it("ignores an occupant of a folder nothing is arriving in", () => {
    expect(findCollisions(collapse([["a.pdf"], ["b.pdf"]]), {
      [`${UNIT}/Credit2/2025`]: ["a.pdf", "b.pdf"],
    })).toEqual([]);
  });

  it("matches names case-insensitively, as SharePoint does", () => {
    expect(findCollisions(collapse([["A.pdf"], ["a.pdf"]]), {}).length).toBe(1);
  });

  it("ignores plans that resolved to no destination", () => {
    const blocked = planLeaf(UNIT, leaf(["testig", "2024"], ["a.pdf"]), CHAIN, {});
    expect(findCollisions([blocked], {})).toEqual([]);
  });

  it("does not double-count a moving file that shares its destination folder", () => {
    // The leaf's destination IS its current folder, which therefore already lists the same file —
    // itself. Counting it twice would invent a collision between a file and itself.
    const l = leaf(["Credit2", "2024", "Tax Return"], ["a.pdf"]);
    const plan = planLeaf(UNIT, l, withoutTesting, {}, removed);
    expect(plan.to).toBe(l.path);
    expect(findCollisions([plan], { [l.path]: ["a.pdf"] })).toEqual([]);
  });
});

describe("suggestRename", () => {
  it("carries the value that made the file distinct, not a number", () => {
    expect(suggestRename("a.pdf", "testig")).toBe("a (testig).pdf");
  });

  it("preserves the extension", () => {
    expect(suggestRename("Tax Return 2024.docx", "live")).toBe("Tax Return 2024 (live).docx");
  });

  it("handles a name with no extension", () => {
    expect(suggestRename("README", "live")).toBe("README (live)");
  });

  it("keeps a leading-dot name intact rather than treating it as all extension", () => {
    expect(suggestRename(".hidden", "live")).toBe(".hidden (live)");
  });

  it("returns the name unchanged when the distinguisher sanitizes to nothing", () => {
    expect(suggestRename("a.pdf", "##")).toBe("a.pdf");
  });

  it("strips characters SharePoint would reject from the distinguisher", () => {
    expect(suggestRename("a.pdf", "R&D: Core")).toBe("a (R&D Core).pdf");
  });
});

describe("validateRename", () => {
  it("accepts a normal rename", () => {
    expect(validateRename("a.pdf", "a (Credit2).pdf")).toBeUndefined();
  });

  it("rejects a blank name", () => {
    expect(validateRename("a.pdf", "   ")).toBe("cannot be blank");
  });

  it("rejects characters SharePoint will not accept", () => {
    expect(validateRename("a.pdf", "a/b.pdf")).toMatch(/cannot contain/);
    expect(validateRename("a.pdf", "a#b.pdf")).toMatch(/cannot contain/);
  });

  it("insists the extension survives", () => {
    // The invisible mistake: one keystroke in a renaming form leaves a file that opens as
    // nothing, and nothing on screen would have said so.
    expect(validateRename("a.pdf", "a")).toBe("must still end in .pdf");
    expect(validateRename("a.pdf", "a.pd")).toBe("must still end in .pdf");
  });

  it("accepts a differently-cased extension", () => {
    expect(validateRename("a.pdf", "b.PDF")).toBeUndefined();
  });

  it("rejects a name that is only the extension", () => {
    expect(validateRename("a.pdf", ".pdf")).toBe("needs a name before the extension");
  });

  it("imposes no extension rule when the original had none", () => {
    expect(validateRename("README", "README (Credit2)")).toBeUndefined();
  });

  it("treats a leading-dot original as having no extension", () => {
    expect(validateRename(".hidden", "renamed")).toBeUndefined();
  });
});

describe("planTotals", () => {
  it("counts moves, strays and skipped units separately", () => {
    const totals = planTotals([
      planUnit(UNIT, ["2024", "2025", "Old"], TIERS, [{ label: "testig", id: "t" }]),
      planUnit(`${UNIT}2`, ["2024"], TIERS, []),
      planUnit(`${UNIT}3`, ["testig"], TIERS, []),
    ]);
    expect(totals).toEqual({ unitsWithMoves: 1, moves: 2, strays: 1, skippedUnits: 1 });
  });

  it("returns zeroes for an empty plan", () => {
    expect(planTotals([])).toEqual({ unitsWithMoves: 0, moves: 0, strays: 0, skippedUnits: 0 });
  });
});
