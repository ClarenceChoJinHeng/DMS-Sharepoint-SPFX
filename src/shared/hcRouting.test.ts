import {
  RoutingContext,
  canOfferHc,
  isHcLevel,
  refuseReason,
  routeFor,
  selectableLevels,
} from "./hcRouting";

/** The levels as the client's term set orders them. */
const LEVELS = ["Public", "Internal", "Restricted", "Confidential", "Highly Confidential"];

/** A site with HC set up, and a user cleared for it. */
const CLEARED: RoutingContext = {
  hcLevel: "Highly Confidential",
  hcAvailable: true,
  canWriteHc: true,
};

/** The same site, and a user who is not cleared. */
const UNCLEARED: RoutingContext = { ...CLEARED, canWriteHc: false };

/** A site that has never configured HC — which is every site, today. */
const NO_HC: RoutingContext = { hcLevel: "", hcAvailable: false };

describe("isHcLevel", () => {
  it("matches the configured label", () => {
    expect(isHcLevel("Highly Confidential", CLEARED)).toBe(true);
    expect(isHcLevel("Confidential", CLEARED)).toBe(false);
  });

  it("ignores case and surrounding space — the label is typed by hand into a config row", () => {
    expect(isHcLevel("  highly confidential ", CLEARED)).toBe(true);
    expect(isHcLevel("Highly Confidential", { ...CLEARED, hcLevel: " HIGHLY CONFIDENTIAL " })).toBe(true);
  });

  it("matches NOTHING when no level is configured, including a blank level", () => {
    // Without this guard, a document with no confidentiality set routes to HC on every site that
    // has not configured the row.
    expect(isHcLevel("Highly Confidential", NO_HC)).toBe(false);
    expect(isHcLevel("", NO_HC)).toBe(false);
    expect(isHcLevel(undefined, NO_HC)).toBe(false);
  });

  it("does not treat a blank document level as the configured level", () => {
    expect(isHcLevel("", CLEARED)).toBe(false);
    expect(isHcLevel(undefined, CLEARED)).toBe(false);
  });
});

describe("routeFor", () => {
  it("sends an ordinary level to the normal pair", () => {
    expect(routeFor("Confidential", CLEARED)).toEqual({ approval: "Staging", documents: "Documents" });
  });

  it("sends the HC level to the HC pair", () => {
    expect(routeFor("Highly Confidential", CLEARED)).toEqual({
      approval: "StagingHC",
      documents: "DocumentsHC",
    });
  });

  it("REFUSES rather than falling back when HC is unavailable", () => {
    // The single worst thing this module could do: routing a Highly Confidential document into the
    // library its whole unit reads, silently, under a green success toast.
    expect(routeFor("Highly Confidential", { ...CLEARED, hcAvailable: false })).toBeUndefined();
  });

  it("still routes ordinary levels normally on a site with no HC at all", () => {
    expect(routeFor("Internal", NO_HC)).toEqual({ approval: "Staging", documents: "Documents" });
  });

  it("returns the pair together, so no caller can take half of it", () => {
    const r = routeFor("Highly Confidential", CLEARED);
    expect(r && r.approval).toBe("StagingHC");
    expect(r && r.documents).toBe("DocumentsHC");
  });
});

describe("canOfferHc", () => {
  it("is true only when configured, available and writable", () => {
    expect(canOfferHc(CLEARED)).toBe(true);
  });

  it("is false when the user cannot write, even though they may hold the group", () => {
    // Reconciliation grants folder ACLs in a pass separate from group creation, so a group can
    // exist for days before it grants anything.
    expect(canOfferHc(UNCLEARED)).toBe(false);
  });

  it("counts an INCONCLUSIVE probe as no", () => {
    expect(canOfferHc({ ...CLEARED, canWriteHc: undefined })).toBe(false);
  });

  it("is false when the libraries did not resolve", () => {
    expect(canOfferHc({ ...CLEARED, hcAvailable: false })).toBe(false);
  });

  it("is false when no level is configured, however permissive everything else is", () => {
    expect(canOfferHc({ hcLevel: "", hcAvailable: true, canWriteHc: true })).toBe(false);
  });
});

describe("selectableLevels", () => {
  it("hides Highly Confidential from someone who is not cleared", () => {
    expect(selectableLevels(LEVELS, UNCLEARED)).toEqual([
      "Public", "Internal", "Restricted", "Confidential",
    ]);
  });

  it("offers it to someone who is", () => {
    expect(selectableLevels(LEVELS, CLEARED)).toEqual(LEVELS);
  });

  it("leaves every other level alone on a site with no HC", () => {
    // Nothing is hidden merely because HC is unconfigured — the levels are still ordinary labels.
    expect(selectableLevels(LEVELS, NO_HC)).toEqual(LEVELS);
  });

  it("keeps the term set's own order rather than sorting", () => {
    const odd = ["Highly Confidential", "Public", "Confidential"];
    expect(selectableLevels(odd, CLEARED)).toEqual(odd);
  });

  it("hides it however it is cased in the term set", () => {
    expect(selectableLevels(["Public", "HIGHLY confidential"], UNCLEARED)).toEqual(["Public"]);
  });

  it("survives an empty or missing level list", () => {
    expect(selectableLevels([], CLEARED)).toEqual([]);
    expect(selectableLevels(undefined as unknown as string[], CLEARED)).toEqual([]);
  });
});

describe("refuseReason", () => {
  it("says nothing for an ordinary level", () => {
    expect(refuseReason("Confidential", UNCLEARED)).toBeUndefined();
  });

  it("says nothing when the user is cleared", () => {
    expect(refuseReason("Highly Confidential", CLEARED)).toBeUndefined();
  });

  it("names the missing library when HC did not resolve", () => {
    const msg = refuseReason("Highly Confidential", { ...CLEARED, hcAvailable: false });
    expect(msg).toContain("library");
    expect(msg).toContain("Nothing has been uploaded");
  });

  it("names the clearance when the probe said no", () => {
    const msg = refuseReason("Highly Confidential", UNCLEARED);
    expect(msg).toContain("not cleared");
    expect(msg).toContain("Nothing has been uploaded");
  });

  it("ALWAYS states that nothing was uploaded", () => {
    // The alternative reading — that a Highly Confidential file is now sitting somewhere unknown —
    // sends someone hunting through libraries they should not be opening.
    const cases: RoutingContext[] = [
      { ...CLEARED, hcAvailable: false },
      UNCLEARED,
      { ...CLEARED, canWriteHc: undefined },
    ];
    for (const ctx of cases) {
      expect(refuseReason("Highly Confidential", ctx)).toContain("Nothing has been uploaded");
    }
  });
});
