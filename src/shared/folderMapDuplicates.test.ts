import { groupDuplicateRows, chooseKeeper, MapRowLite } from "./folderMapDuplicates";

const row = (itemId: number, termGuid: string, folderUniqueId: string, title = "t"): MapRowLite =>
  ({ itemId, termGuid, folderUniqueId, title });

const TAX = "5aae2799-035d-4548-85bb-0d999e9fd1b2";
const TREASURY = "81388291-3695-4136-a610-0fad82c31290";
const DEAD = "a0596d7f-815d-4201-ae68-bdad3e74d383";
const LIVE = "676c5737-f64a-42dc-a2b3-b787ce6438ad";

describe("groupDuplicateRows", () => {
  it("returns only the terms with more than one row", () => {
    const out = groupDuplicateRows([
      row(233, TAX, DEAD), row(234, TAX, LIVE), row(235, TREASURY, LIVE),
    ]);
    expect(out.length).toBe(1);
    expect(out[0].termGuid).toBe(TAX);
    expect(out[0].rows.map((r) => r.itemId)).toEqual([233, 234]);
  });

  it("returns nothing when every term has exactly one row", () => {
    expect(groupDuplicateRows([row(1, TAX, LIVE), row(2, TREASURY, LIVE)])).toEqual([]);
  });

  /**
   * Braces and case are the LATENT second cause. The upload form normalises with
   * normalizeTermGuid; reconciliation used a bare toLowerCase. A `{GUID}` row would miss the
   * lookup, a second row would be created, and the same silent refusal follows by another route.
   */
  it("treats braced, padded and upper-case GUIDs as the SAME term", () => {
    const out = groupDuplicateRows([
      row(1, TAX, DEAD),
      row(2, `{${TAX.toUpperCase()}}`, LIVE),
      row(3, `  ${TAX}  `, LIVE),
    ]);
    expect(out.length).toBe(1);
    expect(out[0].rows.map((r) => r.itemId)).toEqual([1, 2, 3]);
  });

  it("ignores rows with a blank term GUID rather than grouping them together", () => {
    expect(groupDuplicateRows([row(1, "", DEAD), row(2, "   ", LIVE)])).toEqual([]);
  });

  it("preserves input order, so the last row is identifiable", () => {
    const out = groupDuplicateRows([row(9, TAX, DEAD), row(4, TAX, LIVE)]);
    expect(out[0].rows.map((r) => r.itemId)).toEqual([9, 4]);
  });

  it("survives empty and undefined input", () => {
    expect(groupDuplicateRows([])).toEqual([]);
    expect(groupDuplicateRows(undefined)).toEqual([]);
  });
});

describe("chooseKeeper", () => {
  it("keeps the row whose folder resolved and removes the rest — the live case", () => {
    const v = chooseKeeper([row(233, TAX, DEAD), row(234, TAX, LIVE)], new Set([LIVE]));
    expect(v?.keep.itemId).toBe(234);
    expect(v?.remove.map((r) => r.itemId)).toEqual([233]);
    expect(v?.confident).toBe(true);
  });

  it("keeps the FIRST resolving row when several resolve", () => {
    const v = chooseKeeper(
      [row(1, TAX, "aaa"), row(2, TAX, "bbb"), row(3, TAX, DEAD)],
      new Set(["aaa", "bbb"]),
    );
    expect(v?.keep.itemId).toBe(1);
    expect(v?.remove.map((r) => r.itemId)).toEqual([2, 3]);
  });

  /**
   * THE FAIL-CLOSED RULE, and the reason this module is pure. Deleting on a failed probe takes a
   * unit's upload path away, and the person who discovers it is an uploader.
   */
  it("removes NOTHING when no row resolved, and says it is not confident", () => {
    const v = chooseKeeper([row(1, TAX, DEAD), row(2, TAX, "also-dead")], new Set([]));
    expect(v?.remove).toEqual([]);
    expect(v?.confident).toBe(false);
  });

  it("keeps the LAST row when nothing resolved — the one the repair pass will heal", () => {
    const v = chooseKeeper([row(1, TAX, DEAD), row(2, TAX, "also-dead")], new Set([]));
    expect(v?.keep.itemId).toBe(2);
  });

  it("treats an undefined live set as nothing alive, never as everything alive", () => {
    const v = chooseKeeper([row(1, TAX, DEAD), row(2, TAX, LIVE)], undefined);
    expect(v?.remove).toEqual([]);
    expect(v?.confident).toBe(false);
  });

  it("matches folder ids case-insensitively and ignores padding", () => {
    const v = chooseKeeper(
      [row(1, TAX, DEAD), row(2, TAX, `  ${LIVE.toUpperCase()}  `)],
      new Set([LIVE]),
    );
    expect(v?.keep.itemId).toBe(2);
    expect(v?.confident).toBe(true);
  });

  it("handles a single-row group without removing it", () => {
    const v = chooseKeeper([row(1, TAX, LIVE)], new Set([LIVE]));
    expect(v?.keep.itemId).toBe(1);
    expect(v?.remove).toEqual([]);
  });

  it("returns undefined for an empty group", () => {
    expect(chooseKeeper([], new Set([LIVE]))).toBeUndefined();
    expect(chooseKeeper(undefined, new Set([LIVE]))).toBeUndefined();
  });
});
