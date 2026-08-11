import {
  backfillNeeds,
  classifyChild,
  effectiveTiers,
  planTotals,
  planUnit,
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
