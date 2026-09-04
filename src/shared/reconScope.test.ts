import { resolveRunScope, coversEverySegment, ALWAYS_FULL_PASSES, ScopeSegment } from "./reconScope";

const seg = (key: string, stagingFolder: string, label?: string): ScopeSegment =>
  ({ key, stagingFolder, label });

const ALL: ScopeSegment[] = [
  seg("mode_gho", "GHO", "Group Head Office"),
  seg("mode_minamas_ho", "MHO", "Minamas Head Office"),
  seg("mode_nbpol_ho", "NBPOLHO", "NBPOL Head Office"),
];

describe("resolveRunScope", () => {
  /**
   * The safe default is TODAY'S behaviour. A default of "only what changed" would make the first run
   * after an unseen manual edit skip the one segment that needed it.
   */
  it("runs everything when nothing has been chosen", () => {
    const s = resolveRunScope(ALL, undefined);
    expect(s.segments.map((x) => x.key)).toEqual(["mode_gho", "mode_minamas_ho", "mode_nbpol_ho"]);
    expect(s.refused).toBeUndefined();
    expect(s.label).toBe("all 3 segments");
  });

  it("runs everything when every segment is ticked, and says so as 'all'", () => {
    const s = resolveRunScope(ALL, new Set(["mode_gho", "mode_minamas_ho", "mode_nbpol_ho"]));
    expect(s.segments.length).toBe(3);
    expect(s.label).toBe("all 3 segments");
  });

  it("runs a subset and NAMES it, for the log and the audit row", () => {
    const s = resolveRunScope(ALL, new Set(["mode_gho", "mode_minamas_ho"]));
    expect(s.segments.map((x) => x.key)).toEqual(["mode_gho", "mode_minamas_ho"]);
    expect(s.label).toBe("2 of 3 segments (Group Head Office, Minamas Head Office)");
  });

  it("keeps the given order, not the tick order", () => {
    const s = resolveRunScope(ALL, new Set(["mode_nbpol_ho", "mode_gho"]));
    expect(s.segments.map((x) => x.key)).toEqual(["mode_gho", "mode_nbpol_ho"]);
  });

  /** A mis-click, not an instruction to do nothing and report success. */
  it("REFUSES an empty selection, with a reason", () => {
    const s = resolveRunScope(ALL, new Set([]));
    expect(s.segments).toEqual([]);
    expect(s.refused).toBe("Tick at least one business segment to run.");
  });

  it("refuses when the site has no segments at all, and says that instead", () => {
    const s = resolveRunScope([], undefined);
    expect(s.refused).toContain("No business segments are configured");
  });

  /**
   * A segment retired between the page loading and Run being pressed. Covering one segment fewer
   * beats refusing the whole run.
   */
  it("drops a ticked key that matches no segment, silently", () => {
    const s = resolveRunScope(ALL, new Set(["mode_gho", "mode_retired_last_week"]));
    expect(s.segments.map((x) => x.key)).toEqual(["mode_gho"]);
    expect(s.refused).toBeUndefined();
  });

  it("refuses when every ticked key has gone", () => {
    const s = resolveRunScope(ALL, new Set(["mode_gone"]));
    expect(s.refused).toBe("Tick at least one business segment to run.");
  });

  it("falls back to the staging folder when a segment has no label", () => {
    const s = resolveRunScope([seg("a", "GHO"), seg("b", "MHO")], new Set(["a"]));
    expect(s.label).toBe("1 of 2 segments (GHO)");
  });

  it("ignores rows with a blank key rather than counting them", () => {
    const s = resolveRunScope([seg("", "GHO"), seg("b", "MHO")], undefined);
    expect(s.segments.map((x) => x.stagingFolder)).toEqual(["MHO"]);
    expect(s.label).toBe("all 1 segment");
  });

  it("survives null and undefined inputs", () => {
    expect(resolveRunScope(undefined, undefined).refused).toBeDefined();
    expect(resolveRunScope(null, null).refused).toBeDefined();
  });
});

describe("ALWAYS_FULL_PASSES", () => {
  /**
   * ⚠ THE LIST MUST NOT QUIETLY SHRINK. These passes assert state that has NO segment. Scoping them
   * would recreate the "mechanism driven by grant rows, protected thing has no row" bug this
   * codebase has already found four times.
   */
  it("names every pass that must run in full regardless of scope", () => {
    expect(ALWAYS_FULL_PASSES).toEqual([
      "site entry",
      "library state",
      "page access",
      "admin page lockdown",
      "HC gating check",
    ]);
  });
});

describe("coversEverySegment", () => {
  /**
   * ⚠ REGRESSION GUARD. An MHO-only run on 2026-08-19 pruned all 67 of GHO's Folder Map rows,
   * because the orphan passes judged every row against only the terms that run had walked.
   */
  it("is FALSE for a scoped run — the condition that must block orphan pruning", () => {
    expect(coversEverySegment(ALL, resolveRunScope(ALL, new Set(["mode_gho"])))).toBe(false);
    expect(coversEverySegment(ALL, resolveRunScope(ALL, new Set(["mode_gho", "mode_minamas_ho"])))).toBe(false);
  });

  it("is TRUE only when every segment was walked", () => {
    expect(coversEverySegment(ALL, resolveRunScope(ALL, undefined))).toBe(true);
    expect(
      coversEverySegment(ALL, resolveRunScope(ALL, new Set(["mode_gho", "mode_minamas_ho", "mode_nbpol_ho"]))),
    ).toBe(true);
  });

  /** Fails CLOSED: an unknown or refused scope never licenses a deletion. */
  it("is FALSE when the scope is refused, empty or unreadable", () => {
    expect(coversEverySegment(ALL, resolveRunScope(ALL, new Set([])))).toBe(false);
    expect(coversEverySegment([], resolveRunScope([], undefined))).toBe(false);
    expect(coversEverySegment(undefined, undefined)).toBe(false);
    expect(coversEverySegment(ALL, undefined)).toBe(false);
  });

  /** A retired segment dropped from the tick list still leaves the run short of full coverage. */
  it("is FALSE when a ticked key matched nothing, so fewer segments ran", () => {
    expect(coversEverySegment(ALL, resolveRunScope(ALL, new Set(["mode_gho", "mode_gone"])))).toBe(false);
  });
});
