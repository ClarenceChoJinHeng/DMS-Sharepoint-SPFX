import {
  AccessVerdict,
  filterProvisionedPaths,
  filterReachablePaths,
  mappedTermGuidSet,
  normalizeTermGuid,
} from "./segmentReadiness";

const CORU = "aaaaaaaa-1111-4111-8111-aaaaaaaaaaaa";
const GMB = "bbbbbbbb-2222-4222-8222-bbbbbbbbbbbb";
const DEPT = "cccccccc-3333-4333-8333-cccccccccccc";

type Path = { modeKey: string; chain: { id: string; label: string }[] };

const path = (modeKey: string, ...leafIds: string[]): Path => ({
  modeKey,
  chain: [
    { id: DEPT, label: "Group Finance" },
    ...leafIds.map((id) => ({ id, label: id.slice(0, 4) })),
  ],
});

describe("normalizeTermGuid", () => {
  it("lowercases, trims and strips braces", () => {
    expect(normalizeTermGuid(`  {${CORU.toUpperCase()}}  `)).toBe(CORU);
  });

  it("treats null/undefined/blank as empty", () => {
    expect(normalizeTermGuid(undefined)).toBe("");
    expect(normalizeTermGuid(null)).toBe("");
    expect(normalizeTermGuid("   ")).toBe("");
  });
});

describe("mappedTermGuidSet", () => {
  it("returns null for a failed read, which is NOT the same as no rows", () => {
    expect(mappedTermGuidSet(null)).toBeNull();
    expect(mappedTermGuidSet(undefined)).toBeNull();
    expect(mappedTermGuidSet([])).toEqual(new Set());
  });

  it("indexes normalised GUIDs and drops blanks", () => {
    const set = mappedTermGuidSet([
      { termGuid: CORU.toUpperCase() },
      { termGuid: "" },
      { termGuid: undefined },
      { termGuid: `{${GMB}}` },
    ]);
    expect(set).toEqual(new Set([CORU, GMB]));
  });
});

describe("filterProvisionedPaths", () => {
  it("keeps a path whose leaf has a Folder Map row", () => {
    const paths = [path("mode_gho", CORU)];
    const out = filterProvisionedPaths(paths, new Set([CORU]));
    expect(out.paths).toEqual(paths);
    expect(out.withheld).toBe(0);
    expect(out.known).toBe(true);
  });

  it("withholds a path whose leaf folder does not exist yet", () => {
    const out = filterProvisionedPaths(
      [path("mode_gho", CORU), path("mode_gho", GMB)],
      new Set([CORU]),
    );
    expect(out.paths.map((p) => p.chain[1].id)).toEqual([CORU]);
    expect(out.withheld).toBe(1);
    expect(out.known).toBe(true);
  });

  it("empties a half-built segment entirely — that is how it disappears", () => {
    const out = filterProvisionedPaths(
      [path("mode_upops", CORU), path("mode_upops", GMB)],
      new Set(),
    );
    expect(out.paths).toEqual([]);
    expect(out.withheld).toBe(2);
    expect(out.known).toBe(true);
  });

  it("matches across GUID case and braces — a case-sensitive compare would hide every unit", () => {
    const out = filterProvisionedPaths(
      [path("mode_gho", CORU.toUpperCase())],
      new Set([CORU]),
    );
    expect(out.paths).toHaveLength(1);
  });

  it("offers EVERYTHING when the Folder Map could not be read (empty is not unknown)", () => {
    const paths = [path("mode_gho", CORU), path("mode_gho", GMB)];
    const out = filterProvisionedPaths(paths, null);
    expect(out.paths).toEqual(paths);
    expect(out.withheld).toBe(0);
    // The caller must not tell the uploader anything was withheld.
    expect(out.known).toBe(false);
  });

  it("gates on the LEAF, not on an ancestor that happens to be mapped", () => {
    // Department folders are mapped too. Matching any chain entry would offer every
    // unit under a provisioned department, which is the bug this exists to stop.
    const out = filterProvisionedPaths(
      [path("mode_gho", CORU)],
      new Set([DEPT]),
    );
    expect(out.paths).toEqual([]);
    expect(out.withheld).toBe(1);
  });

  it("withholds — and counts — a path with an empty chain", () => {
    const out = filterProvisionedPaths(
      [{ modeKey: "mode_gho", chain: [] as { id: string; label: string }[] }],
      new Set([CORU]),
    );
    expect(out.paths).toEqual([]);
    expect(out.withheld).toBe(1);
  });

  it("leaves an empty input empty without claiming anything was withheld", () => {
    const out = filterProvisionedPaths([] as Path[], new Set([CORU]));
    expect(out.paths).toEqual([]);
    expect(out.withheld).toBe(0);
    expect(out.known).toBe(true);
  });
});

describe("filterReachablePaths", () => {
  const two = [path("mode_nbpol", CORU), path("mode_nbpol", GMB)];

  it("keeps a granted path and withholds a denied one", () => {
    const out = filterReachablePaths(two, ["granted", "denied"]);
    expect(out.paths.map((p) => p.chain[1].id)).toEqual([CORU]);
    expect(out.withheld).toBe(1);
    expect(out.known).toBe(true);
  });

  it("withholds a unit whose folder EXISTS but was never granted to the user", () => {
    // The live failure of 2026-08-12: two new groups whose unit folders already existed
    // from an earlier reconciliation run. Existence said yes; the ACL did not exist.
    const out = filterReachablePaths(two, ["denied", "denied"]);
    expect(out.paths).toEqual([]);
    expect(out.withheld).toBe(2);
    expect(out.known).toBe(true);
  });

  it("treats a missing folder exactly like a denied one — same fix, same message", () => {
    const out = filterReachablePaths(two, ["missing", "granted"]);
    expect(out.paths.map((p) => p.chain[1].id)).toEqual([GMB]);
    expect(out.withheld).toBe(1);
  });

  it("KEEPS an unknown verdict — a throttle is not evidence of denial", () => {
    const out = filterReachablePaths(two, ["unknown", "granted"]);
    expect(out.paths).toHaveLength(2);
    expect(out.withheld).toBe(0);
  });

  it("clears `known` when any verdict is uncertain, so nothing is claimed about folders", () => {
    const out = filterReachablePaths(two, ["unknown", "denied"]);
    expect(out.paths.map((p) => p.chain[1].id)).toEqual([CORU]);
    expect(out.withheld).toBe(1);
    expect(out.known).toBe(false);
  });

  it("fails OPEN on a short verdict array rather than hiding the tail", () => {
    const out = filterReachablePaths(two, ["granted"] as AccessVerdict[]);
    expect(out.paths).toHaveLength(2);
    expect(out.known).toBe(false);
  });

  it("handles an empty input without claiming uncertainty", () => {
    const out = filterReachablePaths([] as Path[], []);
    expect(out.paths).toEqual([]);
    expect(out.withheld).toBe(0);
    expect(out.known).toBe(true);
  });
});
