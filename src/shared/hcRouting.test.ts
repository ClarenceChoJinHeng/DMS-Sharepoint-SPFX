import {
  RoutingContext,
  canOfferHc,
  effectiveHcLevel,
  isHcLevel,
  refuseReason,
  routeFor,
  selectableLevels,
  swapLibrarySegment,
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

describe("effectiveHcLevel", () => {
  it("routes nothing on a site with no HC libraries", () => {
    // The level is an ordinary metadata label there, and always has been. Defaulting to
    // "Highly Confidential" regardless would HIDE an existing option on every site that never
    // asked for this feature.
    expect(effectiveHcLevel("", false)).toBe("");
    expect(effectiveHcLevel(undefined, false)).toBe("");
  });

  it("routes the default label once the libraries exist, config row or not", () => {
    // The other half of the trap: defaulting to blank would let a site that created the libraries
    // but forgot the config row file Highly Confidential documents into the NORMAL library.
    expect(effectiveHcLevel("", true)).toBe("Highly Confidential");
    expect(effectiveHcLevel(undefined, true)).toBe("Highly Confidential");
  });

  it("lets the config row rename the level", () => {
    expect(effectiveHcLevel("Top Secret", true)).toBe("Top Secret");
    expect(effectiveHcLevel("  Top Secret  ", true)).toBe("Top Secret");
  });

  it("answers which label routes, never whether anyone may use it", () => {
    // Kept honest rather than clever: a configured level with no libraries still resolves, and
    // canOfferHc is what refuses. Two questions, two functions.
    expect(effectiveHcLevel("Top Secret", false)).toBe("Top Secret");
    expect(canOfferHc({ hcLevel: "Top Secret", hcAvailable: false, canWriteHc: true })).toBe(false);
  });
});

describe("swapLibrarySegment", () => {
  const P = "/sites/CRS/ApprovalDocument/GHO/GCA/CORU";

  it("swaps the library segment and leaves the tree alone", () => {
    expect(swapLibrarySegment(P, "ApprovalDocument", "HCApprovalDocument")).toBe(
      "/sites/CRS/HCApprovalDocument/GHO/GCA/CORU",
    );
  });

  it("matches on a slash-delimited segment, so a folder of the same name is safe", () => {
    // A unit abbreviated "ApprovalDocument" is absurd but a segment folder sharing the library's
    // name is not, and a bare indexOf would rewrite the wrong one.
    const odd = "/sites/CRS/ApprovalDocument/GHO/ApprovalDocument/CORU";
    expect(swapLibrarySegment(odd, "ApprovalDocument", "HCApprovalDocument")).toBe(
      "/sites/CRS/HCApprovalDocument/GHO/ApprovalDocument/CORU",
    );
  });

  it("refuses a path that does not contain the library at all", () => {
    expect(swapLibrarySegment("/sites/CRS/Shared Documents/GHO", "ApprovalDocument", "HCApprovalDocument"))
      .toBeUndefined();
  });

  it("refuses when either segment is unresolved — naming.ts has no HC fallback", () => {
    expect(swapLibrarySegment(P, "ApprovalDocument", "")).toBeUndefined();
    expect(swapLibrarySegment(P, "", "HCApprovalDocument")).toBeUndefined();
    expect(swapLibrarySegment(undefined, "ApprovalDocument", "HCApprovalDocument")).toBeUndefined();
  });

  it("refuses a no-op swap rather than returning a path that changed nothing", () => {
    // Returning the same path would read as success and file an HC document in the normal library.
    expect(swapLibrarySegment(P, "ApprovalDocument", "ApprovalDocument")).toBeUndefined();
  });

  it("handles a library whose URL segment contains a space", () => {
    expect(swapLibrarySegment("/sites/CRS/Shared Documents/GHO", "Shared Documents", "HCDocuments"))
      .toBe("/sites/CRS/HCDocuments/GHO");
  });
});
