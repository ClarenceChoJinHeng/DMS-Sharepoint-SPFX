import {
  CANDIDATE_PREFIXES,
  LEGACY_PREFIX,
  LIST_SUFFIX,
  candidateTitles,
  cachedListTitle,
  clearNameCache,
  noteCreatedList,
  resolvedPrefix,
  titleForNewList,
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
  PRIMED_SUFFIXES,
  setHcLibraryNames,
  clearHcLibraryNames,
  setArchiveLibraryNames,
  clearArchiveLibraryNames,
  cachedArchiveLibraries,
  archiveAvailable,
  archiveTargetFor,
  allLibraryTitles,
  libraryTargets,
  libApiTitle,
  documentsLibraryTitle,
  setDocumentsLibraryName,
  clearDocumentsLibraryName,
  DOCUMENTS_CANDIDATES,
  DOCUMENTS_URL_SEGMENT,
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
    expect(LIBRARY_CANDIDATES[0]).toBe("Approval for Document");
    expect(LIBRARY_CANDIDATES[LIBRARY_CANDIDATES.length - 1]).toBe(LEGACY_LIBRARY);
    expect(LIBRARY_CANDIDATES.indexOf("ApprovalDocument")).toBeGreaterThan(0);
  });

  /* ⚠ EVERY NAME THIS LIBRARY HAS EVER HAD MUST STAY LISTED, and dropping one is the easy mistake
     when a rename arrives: the newest goes in at the top and an older entry looks redundant. It is
     not. Sites migrate at different times — SDG was still on `Approval Document` when
     ClarenceDMSTesting moved to `Approval for Document` — and a name removed from this array stops
     resolving SILENTLY, taking uploads down on whichever site still uses it. */
  it("keeps every historical name, so a site that has not been renamed still resolves", () => {
    for (const past of ["Approval for Document", "Approval Document", "ApprovalDocument", "Staging"]) {
      expect(LIBRARY_CANDIDATES).toContain(past);
    }
  });
});

describe("resolvedPrefix / titleForNewList — creating a list on a renamed site (2026-08-20)", () => {
  beforeEach(() => clearNameCache());

  it("falls back to the legacy prefix when NOTHING has resolved", () => {
    // A fresh site with no CRS lists at all. Same name the code produced before this module existed.
    expect(resolvedPrefix()).toBe("DMS");
    expect(titleForNewList(LIST_SUFFIX.requests)).toBe(`DMS ${LIST_SUFFIX.requests}`);
  });

  it("learns CRS from any ONE resolved list, so a new list joins the others", () => {
    // The live bug: the Requests list does not exist, so it can never be in the cache, so
    // cachedListTitle answered "DMS Requests" on a site where everything else is CRS.
    noteCreatedList(LIST_SUFFIX.config, `CRS ${LIST_SUFFIX.config}`);
    expect(resolvedPrefix()).toBe("CRS");
    expect(titleForNewList(LIST_SUFFIX.requests)).toBe(`CRS ${LIST_SUFFIX.requests}`);
    // The READ path is deliberately unchanged — it still answers the legacy name for an unresolved
    // suffix, because every caller handles its own 404 and that is the pre-existing behaviour.
    expect(cachedListTitle(LIST_SUFFIX.requests)).toBe(`DMS ${LIST_SUFFIX.requests}`);
  });

  it("prefers the RESOLVED title over any derived name", () => {
    // A list that already exists is never renamed by this: creating is then a no-op the caller sees.
    noteCreatedList(LIST_SUFFIX.requests, "Something Entirely Else");
    expect(titleForNewList(LIST_SUFFIX.requests)).toBe("Something Entirely Else");
  });

  it("ignores a resolved title that does not END with its suffix", () => {
    // Guessing a prefix from a name the probe matched some other way would INVENT one, and the
    // invented prefix would then name every list created afterwards.
    noteCreatedList(LIST_SUFFIX.config, "Configuration Of Things");
    expect(resolvedPrefix()).toBe("DMS");
  });

  it("survives a title that is nothing but the suffix", () => {
    // `cut > 1` guards this: a bare "Config" has no prefix, and slicing would yield "".
    noteCreatedList(LIST_SUFFIX.config, LIST_SUFFIX.config);
    expect(resolvedPrefix()).toBe("DMS");
  });
});

describe("primeNames covers every list suffix (2026-08-20)", () => {
  it("probes EVERY suffix in LIST_SUFFIX, so none silently keeps the legacy name", () => {
    // `LIST_SUFFIX.requests` was absent for the whole life of the requests feature. An unprimed
    // suffix never enters the cache, so cachedListTitle answers "DMS <suffix>" for ever — and on a
    // CRS-renamed site that 404s on a list which exists, with no error to explain it.
    const all = Object.keys(LIST_SUFFIX).map((k) => (LIST_SUFFIX as Record<string, string>)[k]);
    const missing = all.filter((sfx) => PRIMED_SUFFIXES.indexOf(sfx) === -1);
    // NAMED. "one suffix is unprimed" is not something anyone can act on.
    expect(missing).toEqual([]);
  });

  it("probes nothing that is not a real suffix", () => {
    const all = Object.keys(LIST_SUFFIX).map((k) => (LIST_SUFFIX as Record<string, string>)[k]);
    expect(PRIMED_SUFFIXES.filter((sfx) => all.indexOf(sfx) === -1)).toEqual([]);
  });
});

/* ---------------------------------------------------------------------------
 * The seven-year archive pair.
 * Spec: docs/superpowers/specs/2026-08-22-seven-year-archive-design.md
 * ------------------------------------------------------------------------- */
describe("the archive pair (2026-08-22)", () => {
  beforeEach(() => {
    clearArchiveLibraryNames();
    clearHcLibraryNames();
    clearLibraryCache();
  });

  it("is unavailable until it resolves, and NEVER falls back to a legacy name", () => {
    // The whole safety. An unresolved archive that fell back to `Documents` would let the mover
    // "archive" a document by writing it into the live library it was leaving — and then delete
    // the source. Same no-fallback rule as HC, for a related reason.
    expect(archiveAvailable()).toBe(false);
    expect(cachedArchiveLibraries()).toBeUndefined();
  });

  it("an unresolved key falls through to ITSELF, so the request 404s loudly", () => {
    expect(libApiTitle("Archive")).toBe("Archive");
    expect(libApiTitle("ArchiveHC")).toBe("ArchiveHC");
  });

  describe("the pairing rule mirrors the site's HC state", () => {
    it("a NON-HC site needs only the normal archive", () => {
      // Requiring `HC Archive` here would disable archiving on a site that is correctly
      // provisioned — which is why HC's flat both-or-neither rule was NOT copied.
      setArchiveLibraryNames("Archive", "/sites/X/Archive", false);
      expect(archiveAvailable()).toBe(true);
      expect(cachedArchiveLibraries()?.hc).toBeUndefined();
    });

    it("an HC site REFUSES a half-provisioned archive", () => {
      // Accepting it would leave HC documents silently never archived, with nothing reporting it.
      setArchiveLibraryNames("Archive", "/sites/X/Archive", true);
      expect(archiveAvailable()).toBe(false);
    });

    it("an HC site accepts the full pair", () => {
      setArchiveLibraryNames("Archive", "/sites/X/Archive", true, "HC Archive", "/sites/X/HCArchive");
      expect(cachedArchiveLibraries()?.normal.title).toBe("Archive");
      expect(cachedArchiveLibraries()?.hc?.title).toBe("HC Archive");
      // The segment comes from the RootFolder, not the title — gotcha #12.
      expect(cachedArchiveLibraries()?.hc?.urlSegment).toBe("HCArchive");
    });

    it("refuses a blank title or a blank url segment", () => {
      // A blank segment reduces split("//") to matching every separator in the document tree.
      setArchiveLibraryNames("", "/sites/X/Archive", false);
      expect(archiveAvailable()).toBe(false);
      setArchiveLibraryNames("Archive", "", false);
      expect(archiveAvailable()).toBe(false);
    });
  });

  describe("archiveTargetFor", () => {
    it("routes each approved-side library to its own archive, never across", () => {
      setArchiveLibraryNames("Archive", "/sites/X/Archive", true, "HC Archive", "/sites/X/HCArchive");
      expect(archiveTargetFor("Documents")).toBe("Archive");
      expect(archiveTargetFor("DocumentsHC")).toBe("ArchiveHC");
    });

    it("FAILS CLOSED for HC when only the normal archive resolved", () => {
      // The one outcome the pair exists to prevent: a Highly Confidential document moved into the
      // open archive. undefined means the caller must refuse, never improvise a destination.
      setArchiveLibraryNames("Archive", "/sites/X/Archive", false);
      expect(archiveTargetFor("DocumentsHC")).toBeUndefined();
    });

    it("gives an APPROVAL library no archive at all", () => {
      // A document still in an approval library was never approved, and an unapproved draft is not
      // a record.
      setArchiveLibraryNames("Archive", "/sites/X/Archive", true, "HC Archive", "/sites/X/HCArchive");
      expect(archiveTargetFor("Staging")).toBeUndefined();
      expect(archiveTargetFor("StagingHC")).toBeUndefined();
    });

    it("answers undefined for everything while unresolved", () => {
      expect(archiveTargetFor("Documents")).toBeUndefined();
    });
  });

  describe("the archive joins the derived library lists", () => {
    it("allLibraryTitles includes it, so a moved document keeps its metadata", () => {
      // A move carries over only columns that EXIST at the destination; the rest vanish with no
      // error. That is the Remark/LegallyPrivileged gap of 2026-08-10, and worse here because an
      // archived document is a record nobody opens for years.
      setLibraryNames("Approval Document", "/sites/X/ApprovalDocument");
      setHcLibraryNames("HC Approval Document", "/sites/X/HCApprovalDocument",
        "HC Documents", "/sites/X/HCDocuments");
      setArchiveLibraryNames("Archive", "/sites/X/Archive", true, "HC Archive", "/sites/X/HCArchive");
      expect(allLibraryTitles()).toEqual([
        "Approval Document", "Documents", "HC Approval Document", "HC Documents",
        "Archive", "HC Archive",
      ]);
    });

    it("libraryTargets includes it, so a structure change re-shapes it too", () => {
      // Register #15: SubtreeMigrator carried a two-element literal and left every HC document in
      // the old shape for four days, reporting success — because a scan that never looks at a
      // library always finds no drift there. An omitted archive repeats that exactly.
      setLibraryNames("Approval Document", "/sites/X/ApprovalDocument");
      setArchiveLibraryNames("Archive", "/sites/X/Archive", false);
      expect(libraryTargets().map((t) => t.key)).toEqual(["Staging", "Documents", "Archive"]);
    });

    it("neither list gains anything while the archive is unresolved", () => {
      setLibraryNames("Approval Document", "/sites/X/ApprovalDocument");
      expect(allLibraryTitles()).toEqual(["Approval Document", "Documents"]);
      expect(libraryTargets().map((t) => t.key)).toEqual(["Staging", "Documents"]);
    });
  });
});

/* ── The Documents library resolves like every other one (2026-08-28) ──────────
 * It was the ONLY library with no probe: a bare `DOCUMENTS_LIBRARY = "Documents"` constant used
 * directly, so when the client retitled it to `Restricted & Confidential Document` there was no
 * candidate array to extend. Worse, the first fix updated `libraryTargets()` and `allLibraryTitles()`
 * and MISSED `libApiTitle` — which is the API boundary every `getbytitle()` goes through — so the
 * two repaired functions reported the right name while every request still 404ed.
 *
 * These tests pin the boundary, not the display. */
describe("documentsLibraryTitle / libApiTitle", () => {
  beforeEach(() => clearDocumentsLibraryName());
  afterEach(() => clearDocumentsLibraryName());

  it("falls back to the legacy literal before anything is probed", () => {
    expect(documentsLibraryTitle()).toBe("Documents");
  });

  it("⚠ FALLS BACK RATHER THAN GOING UNDEFINED, unlike the HC and archive pairs", () => {
    // Their absence is meaningful, so they fail closed. This library always exists, so the useful
    // failure is a loud 404 — `undefined` would drop the approved side out of `libraryTargets` and
    // silently take reconciliation, the migrator and CRS Search with it.
    expect(documentsLibraryTitle().length).toBeGreaterThan(0);
  });

  it("takes the resolved title", () => {
    setDocumentsLibraryName("Restricted & Confidential Document");
    expect(documentsLibraryTitle()).toBe("Restricted & Confidential Document");
  });

  it("⚠⚠ libApiTitle('Documents') RETURNS THE RESOLVED TITLE, not the logical key", () => {
    // The regression. `libApiTitle` used to fall through to `return lib` for this key, which was
    // right only while the library happened to be titled the same as its key.
    setDocumentsLibraryName("Restricted & Confidential Document");
    expect(libApiTitle("Documents")).toBe("Restricted & Confidential Document");
  });

  it("ignores a blank title rather than blanking the name", () => {
    setDocumentsLibraryName("Restricted & Confidential Document");
    setDocumentsLibraryName("   ");
    expect(documentsLibraryTitle()).toBe("Restricted & Confidential Document");
  });

  it("carries the resolved title into libraryTargets, which reconciliation walks", () => {
    setDocumentsLibraryName("Restricted & Confidential Document");
    const docs = libraryTargets().filter((t) => t.key === "Documents")[0];
    expect(docs.title).toBe("Restricted & Confidential Document");
    // ⚠ AND THE URL SEGMENT IS UNCHANGED. A rename never touches the URL (gotcha #12), which is
    // what left all 13 DOCUMENTS_URL_SEGMENT call sites correct through three renames.
    expect(docs.urlSegment).toBe(DOCUMENTS_URL_SEGMENT);
  });

  it("lists every historical name, so an un-renamed site still resolves first", () => {
    expect(DOCUMENTS_CANDIDATES[0]).toBe("Documents");
    expect(DOCUMENTS_CANDIDATES).toContain("Restricted & Confidential Document");
  });
});
