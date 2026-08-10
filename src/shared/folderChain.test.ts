import { Level, parseLevels } from "./formModel";
import {
  buildOnDemandSegments,
  decideTier,
  effectiveOnDemandTiers,
  gridPlan,
  isPermissioned,
  needsLegacyBelowUnit,
  splitChain,
  validateChain,
} from "./folderChain";

const YEAR_SET = "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf";
const DOCTYPE_SET = "866c5754-258e-401f-8685-03d20ae59b1d";

const dept: Level = { label: "Department", column: "Department", labelCol: "Department", tidCol: "DepartmentTid" };
const unit: Level = { label: "Unit", column: "Unit", labelCol: "Unit", tidCol: "UnitTid" };
const fn: Level = {
  label: "Function", column: "Function", labelCol: "Function", tidCol: "FunctionTid",
  termSet: "11111111-1111-1111-1111-111111111111", permissioned: false,
};
const year: Level = { label: "Year", column: "Year", labelCol: "Year", termSet: YEAR_SET, permissioned: false };
const docType: Level = {
  label: "Document Type", column: "DocumentType", labelCol: "Document_x0020_Type",
  termSet: DOCTYPE_SET, permissioned: false,
};

describe("isPermissioned — absent means true", () => {
  it("treats a missing flag as permissioned", () => {
    expect(isPermissioned(unit)).toBe(true);
  });

  it("demotes only on a literal false", () => {
    expect(isPermissioned(year)).toBe(false);
  });

  // The fail-safe direction. A hand-authored row can easily contain the STRING
  // "false", and reading that as a demotion would move a tier that needs an ACL
  // into the inheriting suffix — a permissions widening nobody would see.
  it("keeps a tier permissioned when the flag is a truthy non-boolean", () => {
    expect(isPermissioned({ ...unit, permissioned: "false" as unknown as boolean })).toBe(true);
  });
});

describe("parseLevels carries the new fields", () => {
  // Without this the whole feature parses away to nothing: the chain reads as
  // all-permissioned with no term sets, and the compatibility bridge then hides
  // it by falling back to Year -> Document Type.
  it("keeps termSet and permissioned:false", () => {
    const parsed = parseLevels(JSON.stringify([unit, year]));
    expect(parsed).toHaveLength(2);
    expect(parsed[1].termSet).toBe(YEAR_SET);
    expect(parsed[1].permissioned).toBe(false);
  });

  it("leaves permissioned undefined when the flag is absent", () => {
    const parsed = parseLevels(JSON.stringify([unit]));
    expect(parsed[0].permissioned).toBeUndefined();
    expect(isPermissioned(parsed[0])).toBe(true);
  });

  it("ignores a permissioned flag that is not a literal false", () => {
    const parsed = parseLevels('[{"label":"Unit","column":"Unit","permissioned":"false"}]');
    expect(isPermissioned(parsed[0])).toBe(true);
  });
});

describe("splitChain", () => {
  it("splits at the permissioned boundary, preserving order", () => {
    const { permissioned, onDemand } = splitChain([dept, unit, fn, year, docType]);
    expect(permissioned.map((l) => l.label)).toEqual(["Department", "Unit"]);
    expect(onDemand.map((l) => l.label)).toEqual(["Function", "Year", "Document Type"]);
  });

  it("handles a chain that is entirely permissioned", () => {
    const { permissioned, onDemand } = splitChain([dept, unit]);
    expect(permissioned).toHaveLength(2);
    expect(onDemand).toHaveLength(0);
  });
});

describe("validateChain", () => {
  it("accepts a permissioned prefix followed by on-demand tiers", () => {
    expect(validateChain([dept, unit, fn, year, docType])).toBeUndefined();
  });

  it("accepts today's live chain unchanged", () => {
    expect(validateChain([dept, unit])).toBeUndefined();
  });

  it("rejects an empty chain", () => {
    expect(validateChain([])?.code).toBe("empty");
  });

  // The rule that stops an ACL'd folder being nested inside an inheriting one.
  // Reconciliation never walks there, so such a folder would never be created
  // and the grant would silently not exist.
  it("rejects a permissioned tier below a non-permissioned one", () => {
    const err = validateChain([dept, year, unit]);
    expect(err?.code).toBe("prefix-not-contiguous");
    expect(err?.message).toContain("Unit");
    expect(err?.message).toContain("Year");
  });

  it("rejects two tiers writing the same column", () => {
    const err = validateChain([unit, { ...year, labelCol: "Unit" }]);
    expect(err?.code).toBe("duplicate-column");
  });
});

describe("needsLegacyBelowUnit — the compatibility bridge", () => {
  // The regression the spec asks for: every live mode row today is [Department,
  // Unit], and those sites must keep producing exactly Year -> Document Type
  // until their row is migrated.
  it("is true for the three live mode rows' shape", () => {
    expect(needsLegacyBelowUnit([dept, unit])).toBe(true);
  });

  it("is false once any below-Unit tier is configured", () => {
    expect(needsLegacyBelowUnit([dept, unit, year, docType])).toBe(false);
  });
});

describe("permissioned depth — what caps reconciliation's term-tree walk", () => {
  const subUnit: Level = {
    label: "SubUnit", column: "SubUnit", labelCol: "SubUnit", tidCol: "SubUnitTid",
    permissioned: false, // cascades: no termSet, terms live under each Unit
  };

  // FolderManager derives its walk depth from the permissioned tiers. If SubUnit ever
  // counted here, reconciliation would create an ACL'd folder per subunit carrying only
  // the owners group — invisible to the people who need it — and Units would stop being
  // leaves, moving the Year x Document Type grid onto subunits.
  it("counts Department and Unit but not SubUnit", () => {
    expect(splitChain([dept, unit, subUnit, year, docType]).permissioned).toHaveLength(2);
  });

  it("is unchanged by adding below-Unit tiers", () => {
    const before = splitChain([dept, unit]).permissioned.length;
    const after = splitChain([dept, unit, subUnit, year, docType]).permissioned.length;
    expect(after).toBe(before);
  });

  // A cascading tier is legal: termSet absent means "children of the tier above",
  // which is exactly how each Unit gets its own SubUnits.
  it("accepts a below-Unit tier with no termSet", () => {
    expect(validateChain([dept, unit, subUnit, year, docType])).toBeUndefined();
    expect(subUnit.termSet).toBeUndefined();
  });

  it("still rejects a permissioned tier below SubUnit", () => {
    expect(validateChain([dept, subUnit, unit])?.code).toBe("prefix-not-contiguous");
  });
});

describe("effectiveOnDemandTiers — one walk for migrated and unmigrated sites", () => {
  // The regression the spec asks for: every live mode row is [Department, Unit],
  // and those sites must keep producing exactly Year -> Document Type, with the
  // same column internal names they have always written.
  it("synthesises the legacy pair for an unmigrated chain", () => {
    const tiers = effectiveOnDemandTiers([dept, unit], YEAR_SET, DOCTYPE_SET);
    expect(tiers.map((t) => t.label)).toEqual(["Year", "Document Type"]);
    expect(tiers.map((t) => t.labelCol)).toEqual(["Year", "Document_x0020_Type"]);
    expect(tiers.map((t) => t.termSet)).toEqual([YEAR_SET, DOCTYPE_SET]);
  });

  it("produces the same path as the configured equivalent", () => {
    const legacy = buildOnDemandSegments(effectiveOnDemandTiers([dept, unit], YEAR_SET, DOCTYPE_SET), {
      Year: { id: "y", label: "2026" }, DocumentType: { id: "d", label: "Invoice" },
    });
    expect(legacy.segments).toEqual(["2026", "Invoice"]);
  });

  it("uses the configured chain once one exists", () => {
    const tiers = effectiveOnDemandTiers([dept, unit, fn, year, docType], YEAR_SET, DOCTYPE_SET);
    expect(tiers.map((t) => t.label)).toEqual(["Function", "Year", "Document Type"]);
  });

  // A configured chain wins even when it omits Year entirely — otherwise the
  // legacy pair would reappear underneath and silently deepen every path.
  it("does not append the legacy pair to a configured chain", () => {
    const tiers = effectiveOnDemandTiers([dept, unit, fn], YEAR_SET, DOCTYPE_SET);
    expect(tiers.map((t) => t.label)).toEqual(["Function"]);
  });
});

describe("decideTier — not every Unit has SubUnits", () => {
  const subUnit: Level = {
    label: "SubUnit", column: "SubUnit", labelCol: "SubUnit", tidCol: "SubUnitTid",
    permissioned: false,
  };

  it("skips a cascading tier when the term above has no children", () => {
    expect(decideTier(subUnit, 0)).toBe("skip");
  });

  it("keeps a cascading tier when the term above has children", () => {
    expect(decideTier(subUnit, 2)).toBe("keep");
  });

  // THE important case. "Not loaded" and "no subunits" both produce an empty list, and the
  // shorter path lands in a folder that exists and looks correct — so a transient failure
  // must never be read as "this tier does not apply".
  it("reports unresolved when the option count is unknown", () => {
    expect(decideTier(subUnit, undefined)).toBe("unresolved");
  });

  // A flat set that came back empty is a config fault to surface, not a signal that the
  // tier is inapplicable here — Year and Document Type apply to every unit.
  it("always keeps a tier with its own term set, even with no options", () => {
    expect(decideTier(year, 0)).toBe("keep");
    expect(decideTier(docType, undefined)).toBe("keep");
  });
});

describe("gridPlan", () => {
  // Matches the shipped two-tier arithmetic exactly: years + years*docTypes.
  // 3 + 60, not 60 — every tier above the deepest is itself a folder.
  it("counts every tier, not just the leaves", () => {
    const p = gridPlan([["2024", "2025", "2026"], Array.from({ length: 20 }, (_, i) => `dt${i}`)]);
    expect(p.total).toBe(63);
  });

  it("sizes a three-tier grid", () => {
    // 2 functions + (2*2) years + (2*2*2) doc types = 2 + 4 + 8
    const p = gridPlan([["HR", "Finance"], ["2025", "2026"], ["Invoice", "Receipt"]]);
    expect(p.total).toBe(14);
  });

  it("gives the last-of-each path for the fast-path probe", () => {
    const p = gridPlan([["2025", "2026"], ["Invoice", "Receipt"]]);
    expect(p.lastPath).toEqual(["2026", "Receipt"]);
  });

  // Today's behaviour when the Document Type set is empty: build the year folders
  // and stop. Nothing can nest inside a folder that cannot be created.
  it("truncates at the first tier with no terms", () => {
    const p = gridPlan([["2025", "2026"], [], ["Invoice"]]);
    expect(p.total).toBe(2);
    expect(p.tiers).toEqual([["2025", "2026"]]);
    expect(p.lastPath).toEqual(["2026"]);
  });

  it("plans nothing for an empty chain", () => {
    expect(gridPlan([]).total).toBe(0);
    expect(gridPlan([[]]).total).toBe(0);
  });
});

describe("buildOnDemandSegments", () => {
  const pick = { Year: { id: "y1", label: "2026" }, DocumentType: { id: "d1", label: "Invoice" } };

  it("produces folder names in chain order", () => {
    const r = buildOnDemandSegments([year, docType], pick);
    expect(r.segments).toEqual(["2026", "Invoice"]);
    expect(r.missing).toEqual([]);
  });

  it("follows the chain when a tier is inserted at the top", () => {
    const r = buildOnDemandSegments([fn, year, docType], {
      ...pick, Function: { id: "f1", label: "Human Resource" },
    });
    expect(r.segments).toEqual(["Human Resource", "2026", "Invoice"]);
  });

  it("follows the chain when a tier is inserted at the bottom", () => {
    const r = buildOnDemandSegments([year, docType, fn], {
      ...pick, Function: { id: "f1", label: "Human Resource" },
    });
    expect(r.segments).toEqual(["2026", "Invoice", "Human Resource"]);
  });

  // Reporting the gap rather than skipping it. A skipped tier routes the file to
  // a shallower path that exists and looks correct.
  it("reports a missing selection instead of shortening the path", () => {
    const r = buildOnDemandSegments([fn, year, docType], pick);
    expect(r.missing).toEqual(["Function"]);
    expect(r.segments).not.toContain("");
  });

  it("sanitizes illegal characters out of a term label", () => {
    const r = buildOnDemandSegments([docType], { DocumentType: { id: "d1", label: "Invoice/Receipt" } });
    expect(r.segments).toEqual(["InvoiceReceipt"]);
  });

  it("treats a whitespace-only label as missing", () => {
    const r = buildOnDemandSegments([year], { Year: { id: "y1", label: "   " } });
    expect(r.missing).toEqual(["Year"]);
    expect(r.segments).toEqual([]);
  });
});
