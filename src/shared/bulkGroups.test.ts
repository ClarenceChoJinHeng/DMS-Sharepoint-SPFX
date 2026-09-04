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
    // Spellings standardised 2026-08-26: the HC uploader spells out `_UPLOADER_…` to match the plain
    // group, and the base group is `_VIEWER` rather than `_EMPLOYEE` so it matches its own HC twin.
    expect(tax).toEqual([
      "GHO_GF_TAX_APPROVER",
      "GHO_GF_TAX_UPLOADER",
      "GHO_GF_TAX_UPLOADER_HIGHLY_CONFIDENTIAL",
      "GHO_GF_TAX_VIEWER",
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
    expect(plan.groups.map((g) => g.name)).toEqual(["GHO_C_LEVEL", "GHO_GF_HOD", "GHO_GHR_HOD"]);
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
    expect(planBulkGroups(seg, [], ALL, []).groups.map((g) => g.name)).toEqual(["GHO_C_LEVEL"]);
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

describe("planBulkGroups recognises a unit whose code was renamed", () => {
  // One PIC group for Tax, provisioned under the OLD code and still named for it.
  // The persona's CURRENT role set. `roleSetKey` is an exact fingerprint, so this must track
  // PERSONAS — see the note on the second test for what happens when a live site's rows do not.
  const picRowsFor = (groupId: string, groupName: string, term: string): GroupMapWriteRow[] =>
    (["UPL"] as GroupMapRole[]).map((r) => ({
      GroupId: groupId,
      GroupName: groupName,
      Segment: SET,
      UnitTermGuid: term,
      Role: r,
      Scope: "Folder",
      Target: "",
    }));

  const renamed: AbbrevRowDraft[] = [
    dept(GF, "Group Finance", "GF"),
    unit("u1", "Tax", "TRC", GF),
  ];

  it("matches on the TERM, not the name, so a rename creates nothing", () => {
    // THE BUG THIS CLOSES: names are derived from the code, so a renamed unit looked new and a bulk
    // run created a second full set of groups plus another 13 rows on the same term.
    const plan = planBulkGroups(
      seg, renamed, ["pic"], ["GHO_GF_TAX_UPLOADER"],
      picRowsFor("7", "GHO_GF_TAX_UPLOADER", "u1"),
    );
    const g = plan.groups.filter((x) => x.personaKey === "pic")[0];
    expect(g.name).toBe("GHO_GF_TRC_UPLOADER");
    expect(g.exists).toBe(true);
    expect(g.existingName).toBe("GHO_GF_TAX_UPLOADER");
    expect(toCreateCount(plan)).toBe(0);
  });

  it("does NOT recognise a renamed unit whose rows predate a persona's role change", () => {
    /**
     * ⚠ PINS A KNOWN LIMIT, not a desired behaviour. `roleSetKey` is an EXACT fingerprint, so when
     * a persona sheds a role — PIC lost DELS on 2026-08-20 — every group provisioned before that
     * still carries the old set and no longer matches. Only the RENAME path is affected: the name
     * check runs first, so an un-renamed group is still recognised.
     *
     * Subset matching was considered and REJECTED: `employee` is ["MEMBER"] and `employee_hc` is
     * ["MEMBER","MEMBERHC"], so a subset rule would let a cleared viewer group answer for the plain
     * one — a false match that reuses the wrong group, which is worse than a duplicate.
     *
     * The operational answer is the one the persona change needs anyway: re-create the affected
     * mappings, which rewrites the rows to the new set.
     */
    const stale = (["UPL", "DELS"] as GroupMapRole[]).map((r) => ({
      GroupId: "7", GroupName: "GHO_GF_TAX_UPLOADER", Segment: SET,
      UnitTermGuid: "u1", Role: r, Scope: "Folder", Target: "",
    })) as GroupMapWriteRow[];
    const plan = planBulkGroups(seg, renamed, ["pic"], ["GHO_GF_TAX_UPLOADER"], stale);
    const g = plan.groups.filter((x) => x.personaKey === "pic")[0];
    expect(g.exists).toBe(false);
    expect(toCreateCount(plan)).toBe(1);
  });

  it("still matches on the NAME, so an interrupted run is finished rather than duplicated", () => {
    // A run stopped part-way leaves the group made with only some of its rows. Its role set does not
    // equal the persona's, so role matching alone would call it new — the regression this guards.
    const partial = picRowsFor("7", "GHO_GF_TRC_UPLOADER", "u1").slice(0, 1);
    const plan = planBulkGroups(
      seg, renamed, ["pic"], ["GHO_GF_TRC_UPLOADER"], partial,
    );
    const g = plan.groups.filter((x) => x.personaKey === "pic")[0];
    expect(g.exists).toBe(true);
    expect(g.existingName).toBeUndefined();
  });

  it("does NOT let a Head of Unit group satisfy the HC PIC", () => {
    // The tempting shortcut is to match a persona's NAMING role. `hou` carries UPLHC among its six,
    // and UPLHC is pic_hc's naming role — so that shortcut would present an approver group as the
    // HC uploader group and leave the real one uncreated.
    const houRoles = ["APR", "DELS", "DEL", "SHARE", "UPLHC", "DELSHC"] as GroupMapRole[];
    const houRows: GroupMapWriteRow[] = houRoles.map((r) => ({
      GroupId: "9", GroupName: "GHO_GF_TAX_APPROVER", Segment: SET,
      UnitTermGuid: "u1", Role: r, Scope: "Folder", Target: "",
    }));
    const plan = planBulkGroups(
      seg, renamed, ["pic_hc"], ["GHO_GF_TAX_APPROVER"], houRows,
    );
    const g = plan.groups.filter((x) => x.personaKey === "pic_hc")[0];
    expect(g.exists).toBe(false);
    expect(g.existingName).toBeUndefined();
  });

  it("ignores rows whose group no longer exists on the site", () => {
    // Rows outlive a deleted group. Treating those as provisioned would report the unit as done while
    // nothing grants anything — understating the work, which is the dangerous direction here.
    const plan = planBulkGroups(
      seg, renamed, ["pic"], [], picRowsFor("7", "GHO_GF_TAX_UPLOADER", "u1"),
    );
    expect(plan.groups.filter((x) => x.personaKey === "pic")[0].exists).toBe(false);
  });

  it("reads a long-form Role value, which a hand-authored row can carry", () => {
    const rows: GroupMapWriteRow[] = [
      { GroupId: "7", GroupName: "GHO_GF_TAX_EMPLOYEE", Segment: SET, UnitTermGuid: "u1",
        Role: "MEMBER" as GroupMapRole, Scope: "Folder", Target: "" },
    ];
    const long = rows.map((r) => ({ ...r, Role: "MEMBER" as GroupMapRole }));
    const plan = planBulkGroups(
      seg, renamed, ["employee"], ["GHO_GF_TAX_EMPLOYEE"], long,
    );
    expect(plan.groups.filter((x) => x.personaKey === "employee")[0].existingName)
      .toBe("GHO_GF_TAX_EMPLOYEE");
  });

  it("plans normally when no rows are supplied at all", () => {
    // Every existing caller passes four arguments; the fifth defaults to none, and the old
    // name-only behaviour must be exactly what they still get.
    const plan = planBulkGroups(seg, rows, ["pic"], []);
    expect(toCreateCount(plan)).toBe(plan.groups.length);
  });
});
