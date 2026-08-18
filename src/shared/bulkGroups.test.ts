import { AbbrevRowDraft } from "./abbreviationDraft";
import { BulkSegment, planBulkGroups, splitPlannedRows, toCreateCount } from "./bulkGroups";
import { GroupMapRole, GroupMapWriteRow } from "./groupMapModel";

const SET = "cccc3333-3333-4333-8333-cccccccccccc";
const GF = "aaaa1111-1111-4111-8111-aaaaaaaaaaaa";
const GHR = "dddd4444-4444-4444-8444-dddddddddddd";

const seg: BulkSegment = { code: "GHO", termSetGuid: SET, levelNames: ["Department", "Unit"] };

const dept = (guid: string, label: string, code: string): AbbrevRowDraft => ({
  termGuid: guid, label, level: "Department", parentGuid: "", abbreviation: code,
});
const unit = (guid: string, label: string, code: string, parent: string): AbbrevRowDraft => ({
  termGuid: guid, label, level: "Unit", parentGuid: parent, abbreviation: code,
});

/** The client's five per unit, plus the two tiers above. */
const ALL = ["pic", "hou", "employee", "pic_hc", "employee_hc", "hod", "clevel_segment"];

const rows: AbbrevRowDraft[] = [
  dept(GF, "Group Finance", "GF"),
  unit("u1", "Tax", "TAX", GF),
  unit("u2", "Treasury", "TREASURY", GF),
  dept(GHR, "Group Human Resources", "GHR"),
  unit("u3", "Rewards", "REWARDS", GHR),
];

describe("planBulkGroups", () => {
  it("plans the client's five personas for every unit", () => {
    const plan = planBulkGroups(seg, rows, ["pic", "hou", "employee", "pic_hc", "employee_hc"], []);
    const tax = plan.groups.filter((g) => g.name.indexOf("_TAX") !== -1).map((g) => g.name).sort();
    expect(tax).toEqual([
      "GHO_GF_TAX_APPROVER",
      "GHO_GF_TAX_EMPLOYEE",
      "GHO_GF_TAX_UPLOADER",
      "GHO_GF_TAX_UPL_HIGHLY_CONFIDENTIAL",
      "GHO_GF_TAX_VIEWER_HIGHLY_CONFIDENTIAL",
    ]);
    // 3 units x 5 = 15, and nothing else, since no department or segment persona was ticked.
    expect(plan.groups.length).toBe(15);
  });

  it("plans the DEPARTMENT and SEGMENT tiers too, which a unit-only preset would silently omit", () => {
    // If the preset only covered units, the Head of Department groups would never exist — and that fails
    // the way this system always fails: nothing errors, and a HoD later opens a department folder that
    // looks empty.
    const plan = planBulkGroups(seg, rows, ["hod", "clevel_segment"], []);
    expect(plan.groups.map((g) => g.name)).toEqual(["GHO_SEGVIEW", "GHO_GF_HOD", "GHO_GHR_HOD"]);
    expect(plan.groups[0].scope).toBe("segment");
    expect(plan.groups[0].tierGuid).toBe(SET); // the segment IS the tier
    expect(plan.groups[1].scope).toBe("department");
    expect(plan.groups[1].tierGuid).toBe(GF);
  });

  it("totals 5 per unit + 1 per department + 1 per segment", () => {
    const plan = planBulkGroups(seg, rows, ALL, []);
    expect(plan.groups.length).toBe(3 * 5 + 2 + 1);
  });

  it("SKIPS a term with no folder code, and says why", () => {
    // Reconciliation skips such a term silently, so its group would grant access to nothing.
    const withGap = [...rows, unit("u4", "Nameless", "", GF)];
    const plan = planBulkGroups(seg, withGap, ["pic"], []);
    expect(plan.groups.filter((g) => g.name.indexOf("NAMELESS") !== -1)).toEqual([]);
    expect(plan.skipped.length).toBe(1);
    expect(plan.skipped[0].label).toBe("Nameless");
    expect(plan.skipped[0].reason).toContain("no folder code");
  });

  it("SKIPS a coded unit whose DEPARTMENT has no code — the subtle one", () => {
    // The unit has a perfectly good code, so a check of the row alone would create its groups. But
    // reconciliation builds GHO/<dept>/<unit>, so with no department code there is no folder beneath which
    // that unit can exist — the group would look correct and grant nothing.
    const orphanDept = "eeee5555-5555-4555-8555-eeeeeeeeeeee";
    const withGap: AbbrevRowDraft[] = [
      dept(orphanDept, "Uncoded Department", ""),
      unit("u9", "Perfectly Coded", "PC", orphanDept),
    ];
    const plan = planBulkGroups(seg, withGap, ["pic"], []);
    expect(plan.groups).toEqual([]);
    expect(plan.skipped.map((x) => x.label).sort()).toEqual(["Perfectly Coded", "Uncoded Department"]);
    expect(plan.skipped.filter((x) => x.label === "Perfectly Coded")[0].reason)
      .toContain("a tier above it");
  });

  it("marks an existing group as exists, case-insensitively, and excludes it from the create count", () => {
    // Idempotent: an existing group is mapped, never re-created. Titles are compared without case because
    // SharePoint does not preserve it reliably across reads.
    const plan = planBulkGroups(seg, rows, ["pic"], ["gho_gf_tax_uploader"]);
    const tax = plan.groups.filter((g) => g.name === "GHO_GF_TAX_UPLOADER")[0];
    expect(tax.exists).toBe(true);
    expect(toCreateCount(plan)).toBe(plan.groups.length - 1);
  });

  it("never plans the same NAME twice", () => {
    // Two personas sharing a naming role would plan one title twice, and the second create would fail as a
    // duplicate mid-run — noise that reads as a real error.
    const plan = planBulkGroups(seg, rows, ALL, []);
    const names = plan.groups.map((g) => g.name.toLowerCase());
    const unique: Record<string, true> = {};
    for (const n of names) unique[n] = true;
    expect(Object.keys(unique).length).toBe(names.length);
  });

  it("survives a cyclic parent link instead of spinning", () => {
    // Two rows naming each other as parent is not a state the UI can produce, but the walk must terminate
    // regardless — a hung page during a 703-group plan is indistinguishable from a crash.
    const a: AbbrevRowDraft = { termGuid: "x", label: "X", level: "Unit", parentGuid: "y", abbreviation: "X" };
    const b: AbbrevRowDraft = { termGuid: "y", label: "Y", level: "Unit", parentGuid: "x", abbreviation: "Y" };
    expect(() => planBulkGroups(seg, [a, b], ["pic"], [])).not.toThrow();
  });

  it("plans nothing when no persona is ticked, rather than defaulting to some", () => {
    expect(planBulkGroups(seg, rows, [], []).groups).toEqual([]);
  });

  it("survives empty rows and a segment with no levels", () => {
    expect(planBulkGroups(seg, [], ALL, []).groups.map((g) => g.name)).toEqual(["GHO_SEGVIEW"]);
    // No levelNames means no tier is the "unit" tier, so every row reads as a department tier. Better a
    // department-scope group than a crash — and the depth check on the New segment form is what stops this
    // state existing in the first place.
    const noLevels: BulkSegment = { ...seg, levelNames: [] };
    expect(() => planBulkGroups(noLevels, rows, ["hod"], [])).not.toThrow();
  });

  it("uses the CODES, never the term labels — the names must match the folders", () => {
    // Reconciliation names every folder from these codes. A name built from labels lines up with nothing,
    // which is the bug fixed in 1.0.143.0 for the single-group builder.
    const plan = planBulkGroups(seg, rows, ["pic"], []);
    for (const g of plan.groups) {
      expect(g.name).not.toContain(" ");
      expect(g.name.indexOf("GHO_")).toBe(0);
    }
  });
});

describe("splitPlannedRows", () => {
  const row = (
    groupId: string,
    role: GroupMapRole,
    tier = "u1",
  ): GroupMapWriteRow => ({
    GroupId: groupId,
    GroupName: `G${groupId}`,
    Segment: SET,
    UnitTermGuid: tier,
    Role: role,
    Scope: "Folder",
    Target: "",
  });

  it("keeps a row the list does not hold", () => {
    const out = splitPlannedRows([], [row("7", "UPL")]);
    expect(out.fresh.map((r: GroupMapWriteRow) => r.Role)).toEqual(["UPL"]);
    expect(out.duplicate).toEqual([]);
  });

  it("drops a row the list already holds — the second-press bug", () => {
    // THE DEFECT: the run wrote unconditionally, so a second press re-wrote every mapping.
    // 642 rows on the rehearsal site came from exactly this.
    const existing = [row("7", "UPL"), row("7", "DELS")];
    const out = splitPlannedRows(existing, [row("7", "UPL"), row("7", "DELS")]);
    expect(out.fresh).toEqual([]);
    expect(out.duplicate.length).toBe(2);
  });

  it("keeps the rows that are new when only some of a group's rows exist", () => {
    // A run killed part-way (defect 2) leaves exactly this state: the group made, one of its
    // two rows written. The re-run must finish it, not skip the group wholesale.
    const out = splitPlannedRows([row("7", "UPL")], [row("7", "UPL"), row("7", "DELS")]);
    expect(out.fresh.map((r: GroupMapWriteRow) => r.Role)).toEqual(["DELS"]);
    expect(out.duplicate.map((r: GroupMapWriteRow) => r.Role)).toEqual(["UPL"]);
  });

  it("de-duplicates WITHIN the candidates, not only against the list", () => {
    // Nothing in the planner produces two identical rows today. If anything ever does, both would
    // be written — the same silent doubling, from a different direction.
    const out = splitPlannedRows([], [row("7", "UPL"), row("7", "UPL")]);
    expect(out.fresh.length).toBe(1);
    expect(out.duplicate.length).toBe(1);
  });

  it("treats a different group, tier or role as a different row", () => {
    const existing = [row("7", "UPL", "u1")];
    expect(splitPlannedRows(existing, [row("8", "UPL", "u1")]).fresh.length).toBe(1);
    expect(splitPlannedRows(existing, [row("7", "UPL", "u2")]).fresh.length).toBe(1);
    expect(splitPlannedRows(existing, [row("7", "APR", "u1")]).fresh.length).toBe(1);
  });

  it("matches case-insensitively, because term GUID casing differs between stores", () => {
    const existing = [row("7", "UPL", "AAAA-BBBB")];
    expect(splitPlannedRows(existing, [row("7", "UPL", "aaaa-bbbb")]).fresh).toEqual([]);
  });
});
