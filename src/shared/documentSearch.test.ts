import {
  buildKql,
  buildListFilter,
  buildRecentFilter,
  CRAWLED_LIBS,
  emptyCriteria,
  failedLibraries,
  hasCriteria,
  hasMetadataFilter,
  isHcLibrary,
  metadataFilterMatches,
  kqlDate,
  kqlPathScope,
  kqlPhrase,
  LIVE_LIBS,
  managedProperty,
  mergeHits,
  odataDate,
  odataLiteral,
  recencyCutoff,
  searchState,
  searchWords,
  sortByModified,
  type LibraryResult,
  type SearchCriteria,
  type SearchHit,
} from "./documentSearch";

const criteria = (over: Partial<SearchCriteria>): SearchCriteria => ({
  ...emptyCriteria(),
  ...over,
});

const hit = (over: Partial<SearchHit>): SearchHit => ({
  uniqueId: "00000000-0000-0000-0000-000000000001",
  itemId: 1,
  library: "Documents",
  name: "file.pdf",
  path: "/sites/x/Shared Documents/file.pdf",
  modified: "2026-08-15T09:00:00Z",
  author: "A. Person",
  size: "1024",
  ...over,
});

describe("library sets", () => {
  it("splits the four libraries between the two engines, with no overlap", () => {
    for (const lib of CRAWLED_LIBS) expect(LIVE_LIBS.indexOf(lib)).toBe(-1);
    expect(CRAWLED_LIBS.length + LIVE_LIBS.length).toBe(4);
  });

  it("puts the growing libraries on Search and the small ones on a live read", () => {
    // The whole reason for the split: Documents has no ceiling, the approval side must be instant.
    expect(CRAWLED_LIBS).toEqual(["Documents", "DocumentsHC"]);
    expect(LIVE_LIBS).toEqual(["Staging", "StagingHC"]);
  });

  it("identifies both HC libraries and neither plain one", () => {
    expect(isHcLibrary("DocumentsHC")).toBe(true);
    expect(isHcLibrary("StagingHC")).toBe(true);
    expect(isHcLibrary("Documents")).toBe(false);
    expect(isHcLibrary("Staging")).toBe(false);
  });
});

describe("hasCriteria", () => {
  it("is false for the resting state, so an empty box never runs a query", () => {
    expect(hasCriteria(emptyCriteria())).toBe(false);
  });

  it("is false for whitespace only", () => {
    expect(hasCriteria(criteria({ text: "   " }))).toBe(false);
  });

  it.each([
    ["text", { text: "tax" }],
    ["documentType", { documentType: "Tax Return" }],
    ["year", { year: "2024" }],
    ["confidentiality", { confidentiality: "Restricted" }],
    ["segment", { segment: "Group Head Office" }],
    ["documentDateFrom", { documentDateFrom: "2026-01-01" }],
    ["documentDateTo", { documentDateTo: "2026-01-01" }],
    ["uploadedFrom", { uploadedFrom: "2026-01-01" }],
    ["uploadedTo", { uploadedTo: "2026-01-01" }],
  ])("is true when %s is set", (_label, over) => {
    expect(hasCriteria(criteria(over as Partial<SearchCriteria>))).toBe(true);
  });

  it("is true for a tier value but false for a tier column with no value", () => {
    expect(hasCriteria(criteria({ tiers: [{ column: "Department", value: "Group Finance" }] }))).toBe(true);
    expect(hasCriteria(criteria({ tiers: [{ column: "Department", value: "" }] }))).toBe(false);
  });

  it("survives a missing object rather than throwing", () => {
    expect(hasCriteria(undefined as unknown as SearchCriteria)).toBe(false);
  });
});

describe("managedProperty", () => {
  it("appends the text suffix", () => {
    expect(managedProperty("ProjectName")).toBe("ProjectNameOWSTEXT");
  });

  it("appends the date suffix when asked", () => {
    expect(managedProperty("DocumentDate", "date")).toBe("DocumentDateOWSDATE");
  });

  it("KEEPS the encoding of an encoded internal name", () => {
    // Decoding would name a property that does not exist — and a KQL clause naming a nonexistent
    // property matches nothing rather than erroring, which is a silently empty filter.
    expect(managedProperty("Business_x0020_Segment")).toBe("Business_x0020_SegmentOWSTEXT");
  });

  it("is blank for a blank name, so no clause is built from nothing", () => {
    expect(managedProperty("")).toBe("");
    expect(managedProperty("   ")).toBe("");
  });
});

describe("kql escaping and dates", () => {
  it("removes double quotes rather than escaping them", () => {
    expect(kqlPhrase('a "b" c')).toBe("a b c");
  });

  it("accepts only YYYY-MM-DD", () => {
    expect(kqlDate("2026-08-16")).toBe("2026-08-16");
    expect(kqlDate("16/08/2026")).toBe("");
    expect(kqlDate("8/16/2026")).toBe("");
    expect(kqlDate("")).toBe("");
  });

  it("splits words and drops empties", () => {
    expect(searchWords("  tax   return ")).toEqual(["tax", "return"]);
    expect(searchWords("   ")).toEqual([]);
  });
});

describe("kqlPathScope", () => {
  it("scopes one library", () => {
    expect(kqlPathScope("https://x.sharepoint.com/sites/s", ["Shared Documents"])).toBe(
      'Path:"https://x.sharepoint.com/sites/s/Shared Documents/*"',
    );
  });

  it("ORs several libraries", () => {
    const q = kqlPathScope("https://x.sharepoint.com/sites/s", ["Shared Documents", "HCDocuments"]);
    expect(q).toBe(
      '(Path:"https://x.sharepoint.com/sites/s/Shared Documents/*" OR ' +
        'Path:"https://x.sharepoint.com/sites/s/HCDocuments/*")',
    );
  });

  it("tolerates a trailing slash on the web url", () => {
    expect(kqlPathScope("https://x.sharepoint.com/sites/s/", ["Docs"])).toBe(
      'Path:"https://x.sharepoint.com/sites/s/Docs/*"',
    );
  });

  it("returns BLANK when there are no segments, never an unscoped query", () => {
    // An unscoped KQL query searches the whole tenant.
    expect(kqlPathScope("https://x.sharepoint.com/sites/s", [])).toBe("");
    expect(kqlPathScope("https://x.sharepoint.com/sites/s", ["", "  "])).toBe("");
  });

  it("returns blank with no web url", () => {
    expect(kqlPathScope("", ["Docs"])).toBe("");
  });
});

describe("buildKql", () => {
  const scope = 'Path:"https://x/sites/s/Shared Documents/*"';

  it("returns blank with no scope, so a caller cannot run an unscoped search", () => {
    expect(buildKql(criteria({ text: "tax" }), "")).toBe("");
  });

  it("returns blank when nothing was asked for", () => {
    expect(buildKql(emptyCriteria(), scope)).toBe("");
  });

  it("always excludes folders", () => {
    // A folder's link IS a library link; 443 of them reached the My Submissions list.
    expect(buildKql(criteria({ text: "tax" }), scope)).toContain("IsDocument:true");
  });

  it("matches a word against contents and the free-text columns", () => {
    const q = buildKql(criteria({ text: "tax" }), scope);
    expect(q).toContain("(tax* OR ProjectNameOWSTEXT:tax*");
    expect(q).toContain("RemarkOWSTEXT:tax*");
  });

  it("ANDs multiple words as separate clauses", () => {
    const q = buildKql(criteria({ text: "tax return" }), scope);
    expect(q).toContain("(tax* OR");
    expect(q).toContain("(return* OR");
  });

  it("adds exact clauses for the fixed metadata", () => {
    const q = buildKql(
      criteria({ documentType: "Tax Return", year: "2024", confidentiality: "Restricted" }),
      scope,
    );
    expect(q).toContain('Document_x0020_TypeOWSTEXT:"Tax Return"');
    expect(q).toContain('YearOWSTEXT:"2024"');
    expect(q).toContain('Confidentiality_x0020_LevelOWSTEXT:"Restricted"');
  });

  it("adds a clause per tier, using the tier's own column name", () => {
    // Upstream Ops names its tiers Region and Estate/Mill — nothing here knows those names.
    const q = buildKql(
      criteria({
        tiers: [
          { column: "Region", value: "Johor" },
          { column: "EstateMill", value: "Bukit Benut" },
        ],
      }),
      scope,
    );
    expect(q).toContain('RegionOWSTEXT:"Johor"');
    expect(q).toContain('EstateMillOWSTEXT:"Bukit Benut"');
  });

  it("skips a tier with no value and one with no column", () => {
    const q = buildKql(
      criteria({
        text: "tax",
        tiers: [
          { column: "Region", value: "" },
          { column: "", value: "Johor" },
        ],
      }),
      scope,
    );
    expect(q).not.toContain("RegionOWSTEXT");
    expect(q).not.toContain('OWSTEXT:"Johor"');
  });

  it("builds both halves of a date range and omits an absent half", () => {
    const both = buildKql(
      criteria({ documentDateFrom: "2026-01-01", documentDateTo: "2026-12-31" }),
      scope,
    );
    expect(both).toContain("DocumentDateOWSDATE>=2026-01-01");
    expect(both).toContain("DocumentDateOWSDATE<=2026-12-31");

    const from = buildKql(criteria({ documentDateFrom: "2026-01-01" }), scope);
    expect(from).toContain("DocumentDateOWSDATE>=2026-01-01");
    expect(from).not.toContain("<=");
  });

  it("ignores a malformed date rather than emitting an invalid clause", () => {
    const q = buildKql(criteria({ text: "tax", documentDateFrom: "01/01/2026" }), scope);
    expect(q).not.toContain("DocumentDateOWSDATE");
  });

  it("uses the built-in Created property for the uploaded range", () => {
    const q = buildKql(criteria({ uploadedFrom: "2026-01-01" }), scope);
    expect(q).toContain("Created>=2026-01-01");
  });

  it("leads with the scope", () => {
    expect(buildKql(criteria({ text: "tax" }), scope).indexOf(scope)).toBe(0);
  });
});

describe("odata escaping and dates", () => {
  it("DOUBLES a single quote rather than stripping it", () => {
    // Apostrophes are ordinary in real names; an unescaped one ends the literal and 400s.
    expect(odataLiteral("O'Brien")).toBe("O''Brien");
  });

  it("emits ISO, not the M/D/YYYY of gotcha #1", () => {
    expect(odataDate("2026-08-16")).toBe("2026-08-16T00:00:00Z");
  });

  it("ends the day at 23:59:59 for a range's upper bound", () => {
    expect(odataDate("2026-08-16", true)).toBe("2026-08-16T23:59:59Z");
  });

  it("rejects anything that is not YYYY-MM-DD", () => {
    expect(odataDate("8/16/2026")).toBe("");
    expect(odataDate("")).toBe("");
  });
});

describe("buildListFilter", () => {
  it("always excludes folders, with the name a $filter accepts", () => {
    // FileSystemObjectType is rejected inside a $filter.
    const f = buildListFilter(emptyCriteria());
    expect(f).toContain("FSObjType eq 0");
    expect(f).not.toContain("FileSystemObjectType");
  });

  it("ORs a word across the fields and ANDs separate words", () => {
    const f = buildListFilter(criteria({ text: "tax return" }));
    expect(f).toContain("(substringof('tax',FileLeafRef) or substringof('tax',ProjectName)");
    expect(f).toContain("substringof('return',Remark)");
    expect(f.split(" and ").length).toBeGreaterThan(2);
  });

  it("escapes a quote inside a searched word", () => {
    expect(buildListFilter(criteria({ text: "O'Brien" }))).toContain("substringof('O''Brien',FileLeafRef)");
  });

  /* ⚠ TWO FAILURES, IN OPPOSITE DIRECTIONS, BOTH FROM ONE COMBINED TEST.
     2026-09-02: all four Advanced Filters were set at once, every library answered HTTP 400, and the
     conclusion drawn was that all four columns are taxonomy — so every one went through
     `substringof('v',TaxCatchAllLabel)`. 2026-09-03: all four answered HTTP **500**, because
     `TaxCatchAllLabel` is a hidden NOTE field and SharePoint cannot filter one at all.
     With four clauses in one `$filter`, ONE bad clause fails the whole request, so that first test
     could never say WHICH column was at fault. Direct inspection had already answered it:
     `SP.Taxonomy.TaxonomyField` on the libraries' own `/fields` shows only THREE of the five are
     taxonomy. These tests pin the split so a third guess cannot be made. */
  it("filters Business Segment with a plain eq — it is a TEXT column, not taxonomy", () => {
    const f = buildListFilter(criteria({ segment: "Group Head Office" }));
    expect(f).toContain("Business_x0020_Segment eq 'Group Head Office'");
    expect(f).not.toContain("TaxCatchAll");
  });

  it("filters a tier on its own COLUMN — tier columns are text, each with a text Tid twin", () => {
    const f = buildListFilter(criteria({ tiers: [{ column: "EstateMill", value: "Bukit Benut" }] }));
    expect(f).toContain("EstateMill eq 'Bukit Benut'");
    expect(f).not.toContain("TaxCatchAll");
  });

  it("NEVER sends the three real taxonomy filters to $filter — that is the 400 and the 500", () => {
    // They are narrowed on the rows instead (`metadataFilterMatches`). One bad clause fails the
    // WHOLE request, so leaking any of them here would take the filters that DO work down with it.
    const f = buildListFilter(criteria({
      documentType: "Letters with Counterparties",
      year: "2024",
      confidentiality: "Confidential",
    }));
    expect(f).not.toContain("Letters with Counterparties");
    expect(f).not.toContain("2024");
    expect(f).not.toContain("Confidential");
    expect(f).not.toContain("TaxCatchAll");
    // Still a valid filter — the folder/file predicate survives on its own.
    expect(f).toContain("FSObjType eq 0");
  });

  it("keeps the text filters when a taxonomy filter is set alongside them", () => {
    const f = buildListFilter(criteria({ segment: "Group Head Office", year: "2024" }));
    expect(f).toContain("Business_x0020_Segment eq 'Group Head Office'");
    expect(f).not.toContain("2024");
  });

  it("builds datetime literals for both date ranges", () => {
    const f = buildListFilter(
      criteria({ documentDateFrom: "2026-01-01", uploadedTo: "2026-12-31" }),
    );
    expect(f).toContain("DocumentDate ge datetime'2026-01-01T00:00:00Z'");
    expect(f).toContain("Created le datetime'2026-12-31T23:59:59Z'");
  });
});

describe("buildRecentFilter", () => {
  it("adds the Modified window on top of the ordinary filter", () => {
    const f = buildRecentFilter(criteria({ text: "tax" }), "2026-08-15T09:00:00Z");
    expect(f).toContain("substringof('tax',FileLeafRef)");
    expect(f).toContain("Modified ge datetime'2026-08-15T09:00:00Z'");
  });

  it("falls back to the plain filter when there is no cutoff", () => {
    expect(buildRecentFilter(criteria({ text: "tax" }), "")).toBe(buildListFilter(criteria({ text: "tax" })));
  });
});

describe("recencyCutoff", () => {
  const now = new Date("2026-08-16T12:00:00Z");

  it("defaults to 24 hours back", () => {
    expect(recencyCutoff(now)).toBe("2026-08-15T12:00:00Z");
  });

  it("honours an explicit window", () => {
    expect(recencyCutoff(now, 2)).toBe("2026-08-16T10:00:00Z");
  });

  it("ignores a nonsense window rather than looking into the future", () => {
    expect(recencyCutoff(now, 0)).toBe("2026-08-15T12:00:00Z");
    expect(recencyCutoff(now, -5)).toBe("2026-08-15T12:00:00Z");
  });

  it("emits no milliseconds, so it can sit in an OData literal", () => {
    expect(recencyCutoff(now)).not.toContain(".");
  });
});

describe("mergeHits", () => {
  it("keeps the first occurrence of a document present in both sets", () => {
    const a = hit({ uniqueId: "A", name: "from-search.pdf" });
    const b = hit({ uniqueId: "A", name: "from-recent.pdf" });
    const out = mergeHits([a], [b]);
    expect(out.length).toBe(1);
    expect(out[0].name).toBe("from-search.pdf");
  });

  it("dedupes case-insensitively, as SharePoint GUIDs vary in case", () => {
    expect(mergeHits([hit({ uniqueId: "abc" })], [hit({ uniqueId: "ABC" })]).length).toBe(1);
  });

  it("dedupes on UniqueId, NOT on path — a moved file is still one file", () => {
    const a = hit({ uniqueId: "A", path: "/old/path.pdf" });
    const b = hit({ uniqueId: "A", path: "/new/path.pdf" });
    expect(mergeHits([a], [b]).length).toBe(1);
  });

  it("keeps distinct documents that happen to share a name", () => {
    expect(mergeHits([hit({ uniqueId: "A" })], [hit({ uniqueId: "B" })]).length).toBe(2);
  });

  it("KEEPS a hit with no UniqueId rather than hiding a document", () => {
    const out = mergeHits([hit({ uniqueId: "" })], [hit({ uniqueId: "" })]);
    expect(out.length).toBe(2);
  });

  it("survives empty and missing sets", () => {
    expect(mergeHits([], undefined as unknown as SearchHit[]).length).toBe(0);
  });
});

describe("sortByModified", () => {
  it("puts the newest first", () => {
    const out = sortByModified([
      hit({ uniqueId: "old", modified: "2026-01-01T00:00:00Z" }),
      hit({ uniqueId: "new", modified: "2026-08-01T00:00:00Z" }),
    ]);
    expect(out.map((h) => h.uniqueId)).toEqual(["new", "old"]);
  });

  it("sorts a blank date last instead of scrambling the order", () => {
    const out = sortByModified([
      hit({ uniqueId: "blank", modified: "" }),
      hit({ uniqueId: "dated", modified: "2026-01-01T00:00:00Z" }),
    ]);
    expect(out.map((h) => h.uniqueId)).toEqual(["dated", "blank"]);
  });

  it("does not mutate its input", () => {
    const input = [hit({ uniqueId: "a", modified: "2026-01-01T00:00:00Z" }), hit({ uniqueId: "b" })];
    sortByModified(input);
    expect(input[0].uniqueId).toBe("a");
  });
});

describe("searchState", () => {
  const res = (over: Partial<LibraryResult>): LibraryResult => ({
    library: "Documents",
    outcome: "ok",
    hits: [],
    ...over,
  });

  it("is idle before anything is searched, even with no results", () => {
    expect(searchState(false, [])).toBe("idle");
  });

  it("is empty when the query ran and matched nothing", () => {
    expect(searchState(true, [res({})])).toBe("empty");
  });

  it("is error when a read FAILED and there is nothing to show", () => {
    // "No results" for an unreachable library made someone upload a second copy in My Submissions.
    expect(searchState(true, [res({ outcome: "failed", status: 404 })])).toBe("error");
  });

  it("treats an HC REFUSAL as empty, not as an error", () => {
    // Every uncleared user would otherwise see a permanent warning and learn to ignore it.
    expect(searchState(true, [res({ library: "DocumentsHC", outcome: "refused" })])).toBe("empty");
  });

  it("is results when one library succeeded even though another failed", () => {
    const out = searchState(true, [
      res({ hits: [hit({})] }),
      res({ library: "Staging", outcome: "failed", status: 403 }),
    ]);
    expect(out).toBe("results");
  });
});

describe("failedLibraries", () => {
  const res = (over: Partial<LibraryResult>): LibraryResult => ({
    library: "Documents",
    outcome: "ok",
    hits: [],
    ...over,
  });

  it("names only genuine failures, never a refusal", () => {
    const out = failedLibraries([
      res({}),
      res({ library: "DocumentsHC", outcome: "refused" }),
      res({ library: "Staging", outcome: "failed", status: 404 }),
    ]);
    expect(out.map((r) => r.library)).toEqual(["Staging"]);
    expect(out[0].status).toBe(404);
  });

  it("survives a missing list", () => {
    expect(failedLibraries(undefined as unknown as LibraryResult[]).length).toBe(0);
  });
});

/* ── The three filters REST cannot apply ───────────────────────────────────────
 *
 * Narrowed on the rows, because `$filter` answers 400 (`eq`) or 500 (`TaxCatchAllLabel`) for a
 * genuine taxonomy column. Selecting them is a different operation and is supported, so the labels
 * arrive with the row.
 */
describe("metadataFilterMatches / hasMetadataFilter", () => {
  type Labels = { documentType: string; year: string; confidentiality: string };
  const row = (over: Partial<Labels>): Labels => ({
    documentType: "", year: "", confidentiality: "", ...over,
  });

  it("passes a row that matches every set filter", () => {
    expect(metadataFilterMatches(
      { documentType: "Tax Return", year: "2024", confidentiality: "Confidential" },
      row({ documentType: "Tax Return", year: "2024", confidentiality: "Confidential" }),
    )).toBe(true);
  });

  it("ignores a filter that is not set", () => {
    expect(metadataFilterMatches(
      { documentType: "", year: "2024", confidentiality: "" },
      row({ documentType: "anything", year: "2024" }),
    )).toBe(true);
  });

  it("rejects a row that misses one of several set filters", () => {
    expect(metadataFilterMatches(
      { documentType: "Tax Return", year: "2024", confidentiality: "" },
      row({ documentType: "Tax Return", year: "2023" }),
    )).toBe(false);
  });

  it("compares trimmed and case-insensitively — these are picked labels, not keys", () => {
    expect(metadataFilterMatches(
      { documentType: "tax return", year: "", confidentiality: "" },
      row({ documentType: "  Tax Return  " }),
    )).toBe(true);
  });

  it("fails a set filter on a row with no value — a blank cannot be shown to match", () => {
    expect(metadataFilterMatches(
      { documentType: "Tax Return", year: "", confidentiality: "" },
      row({}),
    )).toBe(false);
  });

  it("hasMetadataFilter is true only for the three it cannot filter server-side", () => {
    expect(hasMetadataFilter(emptyCriteria())).toBe(false);
    // Business Segment IS filtered server-side, so it must not arm the client-side narrowing.
    expect(hasMetadataFilter({ ...emptyCriteria(), segment: "Group Head Office" })).toBe(false);
    expect(hasMetadataFilter({ ...emptyCriteria(), year: "2024" })).toBe(true);
    expect(hasMetadataFilter({ ...emptyCriteria(), documentType: "Tax Return" })).toBe(true);
    expect(hasMetadataFilter({ ...emptyCriteria(), confidentiality: "Confidential" })).toBe(true);
  });
});

describe("buildListFilter - the Keyword column", () => {
  const words = { ...emptyCriteria(), text: "audit" };

  /* ⚠ THE DEFAULT IS OFF, AND THAT IS THE SAFETY. `Keyword` is created by reconciliation, so a
     library provisioned earlier does not have it - and a $filter naming an absent column returns 400
     and fails the WHOLE read, which would take free-text search down entirely rather than just
     losing one field. */
  it("is absent unless the caller confirms the column", () => {
    expect(buildListFilter(words).indexOf("Keyword")).toBe(-1);
    expect(buildListFilter(words, false).indexOf("Keyword")).toBe(-1);
  });

  it("is matched when the column is confirmed", () => {
    expect(buildListFilter(words, true)).toContain("substringof('audit',Keyword)");
  });

  /* It joins the OR-group rather than adding a clause of its own: a word may match the filename OR
     the keyword OR any other text field, and requiring all of them would match nothing. */
  it("joins the same OR-group as the other text fields", () => {
    const f = buildListFilter(words, true);
    expect(f).toContain("substringof('audit',FileLeafRef) or");
    expect(f).toContain("or substringof('audit',Keyword)");
  });

  it("carries through the recency top-up filter, and defaults off there too", () => {
    const since = "2026-09-01T00:00:00Z";
    expect(buildRecentFilter(words, since).indexOf("Keyword")).toBe(-1);
    expect(buildRecentFilter(words, since, true)).toContain("substringof('audit',Keyword)");
  });
});
