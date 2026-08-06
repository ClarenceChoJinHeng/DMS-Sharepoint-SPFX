import {
  CANDIDATE_PREFIXES,
  LEGACY_PREFIX,
  LIST_SUFFIX,
  candidateTitles,
  clearNameCache,
  resolveListTitle,
  resolveWritePrefix,
  folderContentTypeName,
  siteEntryGroupName,
  groupPrefix,
  matchesAnyGroupPrefix,
  ListProbe,
  LIBRARY_CANDIDATES,
  LEGACY_LIBRARY,
  cachedLibrary,
  libraryTitle,
  libraryUrlSegment,
  setLibraryNames,
  clearLibraryCache,
} from "./naming";

/** A probe that says yes only to the titles it is given. Records calls, to prove caching. */
function probeFor(existing: string[]): { probe: ListProbe; calls: () => string[] } {
  const seen: string[] = [];
  const probe: ListProbe = async (title) => {
    seen.push(title);
    return existing.indexOf(title) !== -1;
  };
  return { probe, calls: () => seen };
}

beforeEach(() => clearNameCache());

describe("candidateTitles", () => {
  it("offers CRS before DMS", () => {
    expect(candidateTitles("Group Map")).toEqual(["CRS Group Map", "DMS Group Map"]);
  });

  it("keeps the candidate list and the legacy prefix consistent", () => {
    // If LEGACY_PREFIX ever falls out of CANDIDATE_PREFIXES, the fallback title becomes a
    // name that is never probed — so a site using it resolves nothing and is then handed
    // that same unprobed name anyway. Silent, and permanent.
    expect(CANDIDATE_PREFIXES).toContain(LEGACY_PREFIX);
  });
});

describe("resolveListTitle", () => {
  it("resolves a renamed list to its CRS title", async () => {
    const { probe } = probeFor(["CRS Group Map"]);
    expect(await resolveListTitle("Group Map", probe)).toBe("CRS Group Map");
  });

  it("resolves an un-renamed list to its DMS title", async () => {
    const { probe, calls } = probeFor(["DMS Group Map"]);
    expect(await resolveListTitle("Group Map", probe)).toBe("DMS Group Map");
    // CRS was tried first and missed — the one extra request a legacy site pays.
    expect(calls()).toEqual(["CRS Group Map", "DMS Group Map"]);
  });

  it("resolves each list independently on a HALF-renamed site", async () => {
    // The case this design exists for, and not hypothetical: observed live on 2026-08-04,
    // when Folder Map, Group Map and Term Abbreviation were renamed and Config was not. A
    // single global prefix taken from Config would have resolved "DMS" and then 404d on all
    // three renamed lists — a total outage caused by a partial rename.
    const { probe } = probeFor([
      "DMS Config",
      "CRS Group Map",
      "CRS Folder Map",
      "CRS Term Abbreviation",
    ]);
    expect(await resolveListTitle(LIST_SUFFIX.config, probe)).toBe("DMS Config");
    expect(await resolveListTitle(LIST_SUFFIX.groupMap, probe)).toBe("CRS Group Map");
    expect(await resolveListTitle(LIST_SUFFIX.folderMap, probe)).toBe("CRS Folder Map");
    expect(await resolveListTitle(LIST_SUFFIX.abbreviation, probe)).toBe("CRS Term Abbreviation");
  });

  it("caches a hit so six callers do not each pay for the probe", async () => {
    const { probe, calls } = probeFor(["CRS Group Map"]);
    await resolveListTitle("Group Map", probe);
    await resolveListTitle("Group Map", probe);
    await resolveListTitle("Group Map", probe);
    expect(calls()).toEqual(["CRS Group Map"]);
  });

  it("falls back to the legacy title when nothing exists", async () => {
    const { probe } = probeFor([]);
    expect(await resolveListTitle("Group Map", probe)).toBe("DMS Group Map");
  });

  it("does NOT cache the fallback, so a throttled probe is retried", async () => {
    // A failed probe is not evidence of absence. Caching the fallback would pin the wrong
    // name for the rest of the session on the strength of one 429 — and the symptom would
    // be a web part reading a list that does not exist, with no error explaining why.
    let firstRound = true;
    const probe: ListProbe = async (title) => {
      if (firstRound) throw new Error("429 throttled");
      return title === "CRS Group Map";
    };
    expect(await resolveListTitle("Group Map", probe)).toBe("DMS Group Map");
    firstRound = false;
    expect(await resolveListTitle("Group Map", probe)).toBe("CRS Group Map");
  });

  it("treats a thrown probe as a miss rather than propagating it", async () => {
    // One absent list must not take down a whole web part.
    const probe: ListProbe = async () => { throw new Error("boom"); };
    await expect(resolveListTitle("Config", probe)).resolves.toBe("DMS Config");
  });
});

describe("resolveWritePrefix", () => {
  it("takes the prefix from whichever Config list answered", async () => {
    const { probe } = probeFor(["CRS Config"]);
    expect(await resolveWritePrefix(probe)).toBe("CRS");
  });

  it("falls back to the legacy prefix when no Config list is found", async () => {
    const { probe } = probeFor([]);
    expect(await resolveWritePrefix(probe)).toBe("DMS");
  });
});

describe("derived names", () => {
  it("builds the folder content type, site-entry group and group prefix", () => {
    expect(folderContentTypeName("CRS")).toBe("CRS Folder");
    expect(siteEntryGroupName("CRS")).toBe("CRS_SITE_MEMBERS");
    expect(groupPrefix("CRS")).toBe("CRS_");
    expect(siteEntryGroupName("DMS")).toBe("DMS_SITE_MEMBERS");
  });
});

describe("matchesAnyGroupPrefix", () => {
  it("accepts BOTH prefixes, whatever resolved", () => {
    // Load-bearing for the site-entry membership sync. Accepting only the resolved prefix
    // on a half-renamed site would sync one half of the users and silently ignore the
    // other — "only SOME people cannot open the site", reported as a successful run.
    expect(matchesAnyGroupPrefix("CRS_GHO_GF_CORU_UPL")).toBe(true);
    expect(matchesAnyGroupPrefix("DMS_GHO_GF_CORU_UPL")).toBe(true);
  });

  it("is case-insensitive and tolerates stray whitespace", () => {
    // The trailing-space bug that silently dropped DMS_GHO_LRC is why this is asserted.
    expect(matchesAnyGroupPrefix("  crs_gho_gf_coru_apr  ")).toBe(true);
  });

  it("rejects anything that is not ours", () => {
    expect(matchesAnyGroupPrefix("Site Owners")).toBe(false);
    expect(matchesAnyGroupPrefix("CRSGHO")).toBe(false);      // no underscore — not our shape
    expect(matchesAnyGroupPrefix("MYCRS_GHO")).toBe(false);   // prefix must be at the start
    expect(matchesAnyGroupPrefix("")).toBe(false);
    expect(matchesAnyGroupPrefix(undefined as unknown as string)).toBe(false);
  });
});

describe("library name pair", () => {
  beforeEach(() => clearLibraryCache());

  it("defaults to the legacy name for BOTH halves until resolved", () => {
    // A site that has not been migrated must behave exactly as it did before this existed.
    expect(libraryTitle()).toBe(LEGACY_LIBRARY);
    expect(libraryUrlSegment()).toBe(LEGACY_LIBRARY);
  });

  it("keeps a renamed title and its unchanged url segment apart", () => {
    // The live 2026-08-06 case: created as ApprovalDocument, retitled to "Approval Document".
    setLibraryNames("Approval Document", "/sites/Example/ApprovalDocument");
    expect(libraryTitle()).toBe("Approval Document");
    expect(libraryUrlSegment()).toBe("ApprovalDocument");
    expect(cachedLibrary()).toEqual({ title: "Approval Document", urlSegment: "ApprovalDocument" });
  });

  it("takes only the last path segment, never the whole url", () => {
    // Using the full server-relative url as a split token would match nothing.
    setLibraryNames("Approval Document", "/sites/Deep/Nested/Web/ApprovalDocument");
    expect(libraryUrlSegment()).toBe("ApprovalDocument");
  });

  it("tolerates a trailing slash on the root url", () => {
    setLibraryNames("Approval Document", "/sites/Example/ApprovalDocument/");
    expect(libraryUrlSegment()).toBe("ApprovalDocument");
  });

  it("trims the title but leaves the segment exact", () => {
    setLibraryNames("  Approval Document  ", "/sites/Example/ApprovalDocument");
    expect(libraryTitle()).toBe("Approval Document");
  });

  it("refuses a blank title or an unusable url, keeping the legacy pair", () => {
    // A segment of "" would reduce split("//") to matching every separator in the tree, so a
    // half-resolved pair must never be stored — absent is recoverable, wrong is not.
    setLibraryNames("", "/sites/Example/ApprovalDocument");
    expect(cachedLibrary()).toEqual({ title: LEGACY_LIBRARY, urlSegment: LEGACY_LIBRARY });
    setLibraryNames("Approval Document", "");
    expect(cachedLibrary()).toEqual({ title: LEGACY_LIBRARY, urlSegment: LEGACY_LIBRARY });
    setLibraryNames("Approval Document", "///");
    expect(cachedLibrary()).toEqual({ title: LEGACY_LIBRARY, urlSegment: LEGACY_LIBRARY });
  });

  it("survives undefined from a malformed response", () => {
    setLibraryNames(undefined as unknown as string, undefined as unknown as string);
    expect(cachedLibrary()).toEqual({ title: LEGACY_LIBRARY, urlSegment: LEGACY_LIBRARY });
  });

  it("probes the current name first and the legacy name last", () => {
    // Order is the whole contract: a site mid-migration can answer to more than one, and the
    // newest must win. Staging last so an un-migrated site still costs the fewest requests.
    expect(LIBRARY_CANDIDATES[0]).toBe("Approval Document");
    expect(LIBRARY_CANDIDATES[LIBRARY_CANDIDATES.length - 1]).toBe(LEGACY_LIBRARY);
    expect(LIBRARY_CANDIDATES.indexOf("ApprovalDocument")).toBeGreaterThan(0);
  });
});
