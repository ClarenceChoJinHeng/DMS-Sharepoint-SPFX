import { Level } from "./formModel";
import {
  buildSegmentLevels,
  columnNameFor,
  columnsForDraft,
  depthVerdict,
  ExistingSegment,
  isGuid,
  modeKeyFor,
  NewSegmentDraft,
  nextSortOrder,
  normalizeGuid,
  validateNewSegment,
} from "./newSegment";

const SET = "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf";
const YEAR_SET = "11111111-1111-4111-8111-111111111111";
const DOCTYPE_SET = "22222222-2222-4222-8222-222222222222";

const year: Level = {
  label: "Year",
  column: "Year",
  labelCol: "Year",
  termSet: YEAR_SET,
  permissioned: false,
};
const docType: Level = {
  label: "Document Type",
  column: "DocumentType",
  labelCol: "Document_x0020_Type",
  termSet: DOCTYPE_SET,
  permissioned: false,
};

const draft = (over: Partial<NewSegmentDraft> = {}): NewSegmentDraft => ({
  label: "Upstream Operations",
  family: "BusinessSegment",
  termSetGuid: SET,
  stagingFolder: "UPOPS",
  permissioned: [{ label: "Region" }, { label: "Estate/Mill" }],
  below: [year, docType],
  ...over,
});

const existing: ExistingSegment[] = [
  { key: "mode_gho", label: "Group Head Office", stagingFolder: "GHO", sortOrder: 1 },
  { key: "mode_nbpol_head_office", label: "NBPOL Head Office", stagingFolder: "NBPOLHO", sortOrder: 3 },
];

describe("columnNameFor", () => {
  it("strips everything SharePoint would encode", () => {
    expect(columnNameFor("Estate/Mill")).toBe("EstateMill");
    expect(columnNameFor("I&T Operating Unit")).toBe("ITOperatingUnit");
    expect(columnNameFor("Region")).toBe("Region");
  });

  it("returns empty when nothing usable is left, rather than inventing a name", () => {
    expect(columnNameFor("///")).toBe("");
    expect(columnNameFor("   ")).toBe("");
  });
});

describe("modeKeyFor", () => {
  it("derives a lowercase underscored key", () => {
    expect(modeKeyFor("Upstream Operations")).toBe("mode_upstream_operations");
    expect(modeKeyFor("NBPOL Head Office")).toBe("mode_nbpol_head_office");
  });

  it("collapses punctuation instead of carrying it into a key", () => {
    // sanitizeFolderSegment REMOVES an illegal path character rather than replacing it with a
    // separator, so the two words join. Documented here because it is the difference between
    // this and modeKeyFor("Estate Mill"), which does keep them apart.
    expect(modeKeyFor("Estate/Mill North")).toBe("mode_estatemill_north");
    expect(modeKeyFor("Estate Mill North")).toBe("mode_estate_mill_north");
  });

  it("returns empty for a name with nothing to slug, so the caller can refuse", () => {
    expect(modeKeyFor("///")).toBe("");
  });
});

describe("nextSortOrder", () => {
  it("is one past the highest existing", () => {
    expect(nextSortOrder(existing)).toBe(4);
  });

  it("starts at 1 when there are no segments", () => {
    expect(nextSortOrder([])).toBe(1);
  });

  it("ignores a blank or non-numeric SortOrder instead of returning NaN", () => {
    // A NaN written back would unsort every segment at once.
    const rows: ExistingSegment[] = [
      { key: "a", label: "A", stagingFolder: "A" },
      { key: "b", label: "B", stagingFolder: "B", sortOrder: undefined },
      { key: "c", label: "C", stagingFolder: "C", sortOrder: 2 },
    ];
    expect(nextSortOrder(rows)).toBe(3);
  });
});

describe("normalizeGuid / isGuid", () => {
  it("accepts the braced and padded forms people paste", () => {
    expect(normalizeGuid(`  {${SET.toUpperCase()}}  `)).toBe(SET);
    expect(isGuid(`{${SET}}`)).toBe(true);
  });

  it("rejects anything that is not a GUID", () => {
    expect(isGuid("")).toBe(false);
    expect(isGuid("not-a-guid")).toBe(false);
    expect(isGuid(SET.slice(0, -1))).toBe(false);
  });
});

describe("buildSegmentLevels", () => {
  it("puts the permissioned tiers first and marks them explicitly", () => {
    const chain = buildSegmentLevels(draft());
    expect(chain.map((l) => l.label)).toEqual(["Region", "Estate/Mill", "Year", "Document Type"]);
    expect(chain[0].permissioned).toBe(true);
    expect(chain[1].permissioned).toBe(true);
    expect(chain[2].permissioned).toBe(false);
  });

  it("derives labelCol and tidCol, and gives a permissioned tier NO termSet", () => {
    const chain = buildSegmentLevels(draft());
    expect(chain[1].labelCol).toBe("EstateMill");
    expect(chain[1].tidCol).toBe("EstateMillTid");
    // termSet is the discriminator that would turn this into a flat list instead of a cascade
    // within the segment's own term tree.
    expect(chain[1].termSet).toBeUndefined();
  });

  it("copies the below-Unit tiers rather than aliasing them", () => {
    const d = draft();
    const chain = buildSegmentLevels(d);
    chain[2].label = "mutated";
    expect(d.below[0].label).toBe("Year");
  });
});

describe("columnsForDraft", () => {
  it("returns a label and Tid column per permissioned tier", () => {
    expect(columnsForDraft(draft()).map((c) => c.internal)).toEqual([
      "Region",
      "RegionTid",
      "EstateMill",
      "EstateMillTid",
    ]);
  });

  it("names the Tid column's display so it reads as an ID, not a duplicate", () => {
    expect(columnsForDraft(draft())[3].display).toBe("Estate/Mill ID");
  });

  it("skips a tier whose name yields no column instead of creating a broken one", () => {
    expect(columnsForDraft(draft({ permissioned: [{ label: "///" }] }))).toEqual([]);
  });
});

describe("validateNewSegment", () => {
  it("accepts a well-formed draft", () => {
    expect(validateNewSegment(draft(), existing)).toEqual([]);
  });

  it("accepts I&T's TWO tiers — the slashed name is one tier, not two", () => {
    // Replaces an assertion that a single tier was valid "because I&T needs exactly one". The client
    // corrected that on 2026-08-17: `I&T Operating Units/Department` is ONE TIER NAME containing a
    // slash, and `Unit` sits below it. So I&T is two tiers like every other family, and the old test
    // was pinning the very misreading that produced it — counting the business segment as a tier and
    // therefore having to drop a real one to make the arithmetic work.
    const d = draft({
      label: "I&T Malaysia",
      stagingFolder: "ITMY",
      permissioned: [{ label: "I&T Operating Units/Department" }, { label: "Unit" }],
    });
    expect(validateNewSegment(d, existing)).toEqual([]);
  });

  it("refuses a duplicate segment name", () => {
    const errs = validateNewSegment(draft({ label: "NBPOL Head Office" }), existing);
    expect(errs.join(" ")).toContain("already exists");
  });

  it("refuses a name that slugs onto an existing KEY even when it reads differently", () => {
    // A different string with the same key — extra whitespace collapses away. Without this
    // check it would present as the new segment shadowing the old one, not as a naming clash.
    const errs = validateNewSegment(draft({ label: "NBPOL  Head   Office" }), existing);
    expect(errs.join(" ")).toContain("mode_nbpol_head_office");
  });

  it("refuses a StagingFolder another segment already uses, naming that segment", () => {
    const errs = validateNewSegment(draft({ stagingFolder: "GHO" }), existing);
    expect(errs.join(" ")).toContain("Group Head Office");
    expect(errs.join(" ")).toContain("permissions");
  });

  it("matches an existing StagingFolder case-insensitively", () => {
    expect(validateNewSegment(draft({ stagingFolder: "gho" }), existing).length).toBe(1);
  });

  it("requires a name, a folder, a term set and the permissioned tiers", () => {
    const errs = validateNewSegment(
      draft({ label: "", stagingFolder: "", termSetGuid: "", permissioned: [] }),
      existing,
    );
    expect(errs.length).toBe(4);
  });

  it("REFUSES a single permissioned tier — two is the floor", () => {
    // The client's instruction 2026-08-17, and not redundant with depthVerdict: one declared tier
    // against a ONE-deep set passes the depth check, and a set is one-deep precisely when the
    // departments have been authored and the units have not. That saves with the DEPARTMENT holding
    // the access, so every unit under it shares one folder and one ACL.
    const errs = validateNewSegment(draft({ permissioned: [{ label: "Department" }] }), existing);
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("at least two");
    // Names the tier back, so the message is about THIS draft rather than a rule recital.
    expect(errs[0]).toContain("Department");
  });

  it("accepts exactly two permissioned tiers", () => {
    const errs = validateNewSegment(
      draft({ permissioned: [{ label: "Region" }, { label: "Estate/Mill" }] }),
      existing,
    );
    expect(errs).toEqual([]);
  });

  it("accepts every intended segment family's tier pair, I&T's slashed tier name included", () => {
    // All thirteen intended segments are two tiers. I&T reads as three only if the business segment
    // is counted, and `I&T Operating Units/Department` is ONE tier name containing a slash — the
    // misreading that flipped this spec twice on 2026-08-17.
    const families: Array<[string, string]> = [
      ["Department", "Unit"],
      ["Region", "Estate/Mill"],
      ["Refinery", "Department"],
      ["I&T Operating Units/Department", "Unit"],
    ];
    for (const [a, b] of families) {
      const errs = validateNewSegment(
        draft({ permissioned: [{ label: a }, { label: b }] }),
        existing,
      );
      expect(errs).toEqual([]);
    }
  });

  it("says what a single tier COSTS, not merely that it is disallowed", () => {
    // An admin refused without a reason retypes the same thing or gives up. The message has to name
    // the consequence — everything beneath shares one set of access — because that is the fact that
    // makes the second tier obviously worth adding.
    const errs = validateNewSegment(draft({ permissioned: [{ label: "Unit" }] }), existing);
    expect(errs[0]).toContain("shares one set of access");
  });

  it("refuses two tiers that would need the same column", () => {
    const errs = validateNewSegment(
      draft({ permissioned: [{ label: "Estate/Mill" }, { label: "Estate Mill" }] }),
      existing,
    );
    expect(errs.join(" ")).toContain("EstateMill");
  });

  it("refuses a tier whose name yields no column at all", () => {
    const errs = validateNewSegment(draft({ permissioned: [{ label: "///" }] }), existing);
    expect(errs.join(" ")).toContain("cannot become a column");
  });

  it("refuses a below-Unit tier that claims to be permissioned — the prefix must be contiguous", () => {
    // The rogue flag has to sit BELOW an inheriting tier to break contiguity. A permissioned
    // Year directly after the named tiers is simply a three-tier prefix, which is legal.
    const rogue: Level = { ...docType, permissioned: true };
    const errs = validateNewSegment(draft({ below: [year, rogue] }), existing);
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("Document Type");
  });

  it("reports the field problems without also reporting a chain problem caused by them", () => {
    // Chain validation runs only once the fields are sound, so an empty draft returns the four
    // things to fix rather than five, one of which the admin cannot act on.
    const errs = validateNewSegment(
      draft({ label: "", stagingFolder: "", termSetGuid: "", permissioned: [] }),
      existing,
    );
    expect(errs.join(" ")).not.toContain("structure for this segment is empty");
  });
});

describe("depthVerdict", () => {
  it("passes when the term set is exactly as deep as the named tiers", () => {
    expect(depthVerdict(2, 2)).toEqual({ ok: true });
    expect(depthVerdict(1, 1)).toEqual({ ok: true });
  });

  it("REFUSES a set deeper than the tiers, naming both numbers", () => {
    const v = depthVerdict(3, 2);
    expect(v.ok).toBe(false);
    expect(v.error).toContain("3 levels");
    expect(v.error).toContain("2 levels");
    expect(v.error).toContain("Name another level");
  });

  it("REFUSES a set shallower than the tiers and suggests the opposite fix", () => {
    const v = depthVerdict(2, 3);
    expect(v.ok).toBe(false);
    expect(v.error).toContain("Remove a level");
  });

  it("refuses an empty term set with its own message — nothing would be created", () => {
    const v = depthVerdict(0, 2);
    expect(v.ok).toBe(false);
    expect(v.error).toContain("no terms in it yet");
  });

  it("WARNS but allows when the depth could not be measured — unknown is not a mismatch", () => {
    const v = depthVerdict(undefined, 2);
    expect(v.ok).toBe(true);
    expect(v.warn).toContain("Could not measure");
    expect(v.error).toBeUndefined();
  });

  it("uses singular wording for one level", () => {
    expect(depthVerdict(2, 1).error).toContain("1 level");
  });
});
