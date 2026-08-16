/**
 * Building the queries behind CRS Search, and reconciling their results.
 *
 * Spec: docs/superpowers/specs/2026-08-16-document-search-design.md
 *
 * Pure and SPFx-free. Everything here is string construction and list arithmetic; the web part
 * supplies the HTTP.
 *
 * ── The one thing to understand before changing anything here ────────────────
 * NOTHING IN THIS FILE ENFORCES WHO MAY SEE WHAT, and nothing here should ever try.
 *
 * Both engines are security-trimmed by SharePoint before a row reaches this code: the Search API
 * applies ACLs at query time, and a REST list read returns only items the caller can open. So the
 * hierarchy the client described — PIC sees their unit, Head of Department sees the department,
 * C-Level sees the segment — falls out of the folder ACLs reconciliation already grants, with no
 * code here.
 *
 * That is the SAFE direction. A bug in this file can only ever return FEWER rows than the person is
 * entitled to. A permission check written here could return more, and would be a second copy of a
 * model whose source of truth is `groupMapModel.ts` and the live ACLs.
 *
 * ── Why there are two engines ────────────────────────────────────────────────
 * `Documents` grows without bound, and REST `$filter` cannot search it: SharePoint throws the
 * 5,000-item list view threshold on any filter it cannot serve from an index, and `substringof`
 * — contains-text, which is most of a search box — cannot use an index at all. Past 5,000 items
 * free-text search would fail entirely rather than return fewer results.
 *
 * The approval libraries are the opposite case. They are small by construction (a file LEAVES on
 * approval) and they hold the just-uploaded file, in front of the one person most certain it
 * exists. Crawl latency there reads as a broken search, so they stay on `$filter`.
 */

/** The four libraries a search can cover. Logical keys, translated at the API boundary. */
export type SearchLibrary = "Documents" | "DocumentsHC" | "Staging" | "StagingHC";

/** The libraries served by the Search API. */
export const CRAWLED_LIBS: SearchLibrary[] = ["Documents", "DocumentsHC"];

/** The libraries served by a live REST read. */
export const LIVE_LIBS: SearchLibrary[] = ["Staging", "StagingHC"];

/** True for the two Highly Confidential libraries. */
export function isHcLibrary(lib: SearchLibrary): boolean {
  return lib === "DocumentsHC" || lib === "StagingHC";
}

/** One tier filter: the column's internal name and the value chosen. */
export interface TierFilter {
  /** Internal name, e.g. `Department` or `EstateMill`. */
  column: string;
  /** The term label as filed. */
  value: string;
}

/** Everything the filter panel can ask for. Every field is optional; blank means "no constraint". */
export interface SearchCriteria {
  /** The free-text box. Multiple words are ANDed. */
  text: string;
  documentType: string;
  year: string;
  confidentiality: string;
  /** The selected segment's label, matched against `Business Segment`. */
  segment: string;
  /** Tier values, only ever for the selected segment's own tiers. */
  tiers: TierFilter[];
  /** `YYYY-MM-DD`, or blank. */
  documentDateFrom: string;
  documentDateTo: string;
  uploadedFrom: string;
  uploadedTo: string;
}

/** An empty criteria object — the resting state, which must not be run as a query. */
export function emptyCriteria(): SearchCriteria {
  return {
    text: "",
    documentType: "",
    year: "",
    confidentiality: "",
    segment: "",
    tiers: [],
    documentDateFrom: "",
    documentDateTo: "",
    uploadedFrom: "",
    uploadedTo: "",
  };
}

const clean = (v: string | undefined): string => (v ?? "").trim();

/**
 * True when the user has actually asked for something.
 *
 * Guards the difference between "nothing searched yet" and "nothing matched" — the first two of the
 * three empty states. Running an unconstrained query would return the whole repository and present
 * it as a search result.
 */
export function hasCriteria(c: SearchCriteria): boolean {
  if (!c) return false;
  if (clean(c.text).length > 0) return true;
  if (clean(c.documentType).length > 0) return true;
  if (clean(c.year).length > 0) return true;
  if (clean(c.confidentiality).length > 0) return true;
  if (clean(c.segment).length > 0) return true;
  if (clean(c.documentDateFrom).length > 0 || clean(c.documentDateTo).length > 0) return true;
  if (clean(c.uploadedFrom).length > 0 || clean(c.uploadedTo).length > 0) return true;
  for (const t of c.tiers ?? []) {
    if (clean(t?.value).length > 0) return true;
  }
  return false;
}

/* ─────────────────────────────── Managed properties ─────────────────────────────── */

/**
 * The managed property name for a column's internal name.
 *
 * KQL cannot filter on an internal name — only on a managed property. SharePoint Online
 * auto-creates queryable ones for list columns by appending a type suffix to the crawled property,
 * so `ProjectName` becomes `ProjectNameOWSTEXT`.
 *
 * THIS IS THE ONE ASSUMPTION IN THE FEATURE THAT HAS TO BE VERIFIED ON A LIVE SITE, and it is
 * isolated in this single function for exactly that reason. If the derivation is wrong the fallback
 * is mapping crawled properties to `RefinableString00`-`99` in Site Settings -> Search Schema, which
 * a site collection admin can do WITHOUT tenant access — the same route as the in-site term store
 * pivot, and the client has already refused tenant access once.
 *
 * Encoded internal names keep their encoding: the crawled property is named after the internal name
 * verbatim, so `Business_x0020_Segment` yields `Business_x0020_SegmentOWSTEXT`. Decoding it here
 * would produce a property that does not exist, and a KQL clause naming a nonexistent property
 * matches NOTHING rather than erroring — a filter that silently returns an empty list.
 */
export function managedProperty(internalName: string, kind?: "text" | "date"): string {
  const name = clean(internalName);
  if (name.length === 0) return "";
  return `${name}OWS${kind === "date" ? "DATE" : "TEXT"}`;
}

/** The free-text columns a typed word is matched against, over and above filename and contents. */
export const TEXT_COLUMNS = ["ProjectName", "Vendor_x002f_CustomerName", "Remark"];

/* ─────────────────────────────── KQL ─────────────────────────────── */

/**
 * A value made safe to sit inside a KQL phrase.
 *
 * Double quotes are REMOVED rather than escaped: KQL's escaping rules differ by context, and a
 * mis-escaped quote does not error — it silently changes what the query means. Term labels do not
 * contain quotes, so removing them costs nothing real.
 */
export function kqlPhrase(value: string): string {
  return clean(value).replace(/"/g, "");
}

/** `prop:"value"`, or blank when there is no value to constrain on. */
function kqlEquals(property: string, value: string): string {
  const v = kqlPhrase(value);
  if (v.length === 0 || property.length === 0) return "";
  return `${property}:"${v}"`;
}

/**
 * One typed word, matched across filename, document contents and the free-text columns.
 *
 * The bare term covers the full-text index — filename and the contents of PDFs and Word files. The
 * property clauses are added explicitly because auto-created managed properties are queryable but
 * are not necessarily part of the full-text index, so a word that appears only in `Remark` would
 * otherwise not be found.
 */
function kqlWordClause(word: string): string {
  const w = kqlPhrase(word);
  if (w.length === 0) return "";
  const parts = [`${w}*`];
  for (const col of TEXT_COLUMNS) {
    parts.push(`${managedProperty(col)}:${w}*`);
  }
  return `(${parts.join(" OR ")})`;
}

/** Split a typed string into words. Multiple words are ANDed: each must match somewhere. */
export function searchWords(text: string): string[] {
  return clean(text)
    .split(/\s+/)
    .map((w) => w.trim())
    .filter((w) => w.length > 0);
}

/** A `YYYY-MM-DD` date, or blank if it is not one. KQL rejects anything else. */
export function kqlDate(value: string): string {
  const v = clean(value);
  return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : "";
}

/** `prop>=from prop<=to`, omitting whichever half is absent. */
function kqlRange(property: string, from: string, to: string): string[] {
  const out: string[] = [];
  const a = kqlDate(from);
  const b = kqlDate(to);
  if (property.length === 0) return out;
  if (a.length > 0) out.push(`${property}>=${a}`);
  if (b.length > 0) out.push(`${property}<=${b}`);
  return out;
}

/**
 * The `Path:` clauses scoping a query to the libraries it may read.
 *
 * Built from RESOLVED url segments, never from a literal. A list's title and its URL are
 * independent and only one of them fails loudly (gotcha #12) — and `Documents` has the URL segment
 * `Shared Documents`, the exact trap that put "Shared Documents" at the head of every folder trail
 * in My Submissions.
 *
 * An empty list of segments returns blank rather than an unscoped query: a query with no `Path:`
 * clause searches the entire tenant, which would show a user documents from sites this system has
 * nothing to do with.
 */
export function kqlPathScope(webAbsoluteUrl: string, urlSegments: string[]): string {
  const base = clean(webAbsoluteUrl).replace(/\/+$/, "");
  const segs = (urlSegments ?? []).map(clean).filter((s) => s.length > 0);
  if (base.length === 0 || segs.length === 0) return "";
  const paths = segs.map((s) => `Path:"${base}/${kqlPhrase(s)}/*"`);
  return paths.length === 1 ? paths[0] : `(${paths.join(" OR ")})`;
}

/**
 * The whole KQL query.
 *
 * Returns blank when the scope is empty or nothing was asked for — the caller must treat blank as
 * "do not run", not as "match everything".
 */
export function buildKql(c: SearchCriteria, scope: string): string {
  if (clean(scope).length === 0) return "";
  if (!hasCriteria(c)) return "";

  const parts: string[] = [scope];
  // Folders carry metadata too, and a folder's link IS a library link — clicking one appears to
  // "open the document library", which is how 443 folders reached the My Submissions list.
  parts.push("IsDocument:true");

  for (const word of searchWords(c.text)) {
    const clause = kqlWordClause(word);
    if (clause.length > 0) parts.push(clause);
  }

  const exact = [
    kqlEquals(managedProperty("Document_x0020_Type"), c.documentType),
    kqlEquals(managedProperty("Year"), c.year),
    kqlEquals(managedProperty("Confidentiality_x0020_Level"), c.confidentiality),
    kqlEquals(managedProperty("Business_x0020_Segment"), c.segment),
  ];
  for (const clause of exact) {
    if (clause.length > 0) parts.push(clause);
  }

  for (const tier of c.tiers ?? []) {
    const clause = kqlEquals(managedProperty(clean(tier?.column)), clean(tier?.value));
    if (clause.length > 0) parts.push(clause);
  }

  const ranges = [
    ...kqlRange(managedProperty("DocumentDate", "date"), c.documentDateFrom, c.documentDateTo),
    ...kqlRange("Created", c.uploadedFrom, c.uploadedTo),
  ];
  for (const clause of ranges) parts.push(clause);

  return parts.join(" ");
}

/* ─────────────────────────────── REST $filter ─────────────────────────────── */

/**
 * A value made safe to sit inside an OData string literal.
 *
 * A single quote is doubled — the OData escape. Unlike the KQL case this one MUST be escaped rather
 * than stripped, because an unescaped quote terminates the literal early and produces a 400 that
 * names a parse error rather than the field, and because apostrophes are ordinary in real names.
 */
export function odataLiteral(value: string): string {
  return clean(value).replace(/'/g, "''");
}

/** `substringof('v',Field)`, or blank. */
function odataContains(field: string, value: string): string {
  const v = odataLiteral(value);
  if (v.length === 0 || field.length === 0) return "";
  return `substringof('${v}',${field})`;
}

/** `Field eq 'v'`, or blank. */
function odataEquals(field: string, value: string): string {
  const v = odataLiteral(value);
  if (v.length === 0 || field.length === 0) return "";
  return `${field} eq '${v}'`;
}

/**
 * `YYYY-MM-DDTHH:MM:SSZ` for a `YYYY-MM-DD` input, or blank.
 *
 * ISO, NOT the `M/D/YYYY` of gotcha #1. That locale format belongs to `validateUpdateListItem`,
 * which parses in the site's locale; an OData `$filter` answers a locale string with a 400 naming
 * the type and not the field.
 */
export function odataDate(value: string, endOfDay?: boolean): string {
  const v = clean(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return "";
  return `${v}T${endOfDay ? "23:59:59" : "00:00:00"}Z`;
}

/**
 * The `$filter` for one of the approval libraries.
 *
 * `FSObjType eq 0` excludes folders. Note the name: the REST spelling `FileSystemObjectType` is
 * rejected inside a `$filter`.
 *
 * Free text is matched on the metadata columns and the filename only — there is no document-contents
 * search here, and that is the accepted cost of being instant on the library where instant matters.
 */
export function buildListFilter(c: SearchCriteria): string {
  const parts: string[] = ["FSObjType eq 0"];

  for (const word of searchWords(c.text)) {
    const anyField = [
      odataContains("FileLeafRef", word),
      odataContains("ProjectName", word),
      odataContains("Vendor_x002f_CustomerName", word),
      odataContains("Remark", word),
    ].filter((p) => p.length > 0);
    // Every word must match SOMEWHERE, but may match a different field than its neighbour.
    if (anyField.length > 0) parts.push(`(${anyField.join(" or ")})`);
  }

  const exact = [
    odataEquals("Business_x0020_Segment", c.segment),
    odataEquals("Year", c.year),
    odataEquals("Confidentiality_x0020_Level", c.confidentiality),
    odataEquals("Document_x0020_Type", c.documentType),
  ];
  for (const clause of exact) {
    if (clause.length > 0) parts.push(clause);
  }

  for (const tier of c.tiers ?? []) {
    const clause = odataEquals(clean(tier?.column), clean(tier?.value));
    if (clause.length > 0) parts.push(clause);
  }

  const ranges: string[] = [];
  const dFrom = odataDate(c.documentDateFrom);
  const dTo = odataDate(c.documentDateTo, true);
  if (dFrom.length > 0) ranges.push(`DocumentDate ge datetime'${dFrom}'`);
  if (dTo.length > 0) ranges.push(`DocumentDate le datetime'${dTo}'`);
  const uFrom = odataDate(c.uploadedFrom);
  const uTo = odataDate(c.uploadedTo, true);
  if (uFrom.length > 0) ranges.push(`Created ge datetime'${uFrom}'`);
  if (uTo.length > 0) ranges.push(`Created le datetime'${uTo}'`);
  for (const clause of ranges) parts.push(clause);

  return parts.join(" and ");
}

/**
 * The `$filter` for the recency top-up read against the approved-side libraries.
 *
 * ── Why this exists ─────────────────────────────────────────────────────────
 * A file approved two minutes ago is in NEITHER engine. Auto-route has moved it out of the approval
 * library and deleted the source — that delete is load-bearing for security, so it is not
 * negotiable — and it is not yet crawled into `Documents`. The window is the crawl interval, minutes
 * to hours, and during it the document is simply unfindable with nothing on screen to say so.
 *
 * `Modified` is indexable and the window keeps the result set tiny, so this stays threshold-safe at
 * any library size — which is the whole reason the main read could not be done this way.
 *
 * It runs on EVERY search rather than only when the main read comes back empty: the gap is per
 * document, not per query, so a search with plenty of results can still be missing the specific
 * recent one somebody is looking for.
 */
export function buildRecentFilter(c: SearchCriteria, sinceIso: string): string {
  const since = clean(sinceIso);
  const base = buildListFilter(c);
  if (since.length === 0) return base;
  return `${base} and Modified ge datetime'${since}'`;
}

/** The cutoff for the top-up read: `now` less `hours`, as ISO. Generous by design. */
export function recencyCutoff(now: Date, hours?: number): string {
  const span = typeof hours === "number" && hours > 0 ? hours : 24;
  return new Date(now.getTime() - span * 3600 * 1000).toISOString().replace(/\.\d{3}Z$/, "Z");
}

/* ─────────────────────────────── Results ─────────────────────────────── */

/** One row of the result list. */
export interface SearchHit {
  /** `UniqueId` — the only stable identity across the two engines. */
  uniqueId: string;
  /** Item id, valid ONLY within `library`. */
  itemId: number;
  /** Which library it came from. Travels with the row; item ids repeat across lists. */
  library: SearchLibrary;
  name: string;
  /** Server-relative path of the file. */
  path: string;
  /** ISO, or blank. */
  modified: string;
  author: string;
  /** Byte count as text, as SharePoint returns it. */
  size: string;
}

const keyOf = (h: SearchHit): string => clean(h?.uniqueId).toLowerCase();

/**
 * Merge result sets, keeping the FIRST occurrence of each document.
 *
 * Deduped on `UniqueId`, never on path or name. The same document legitimately appears in both the
 * crawled result and the recency top-up, and its path may have changed between the crawl and now —
 * so a path key would keep two copies of one file, which is the failure this exists to prevent.
 *
 * A hit with no `UniqueId` is kept rather than dropped: it cannot be deduped, but discarding it
 * would hide a document because of a missing field, and showing it twice is the lesser fault.
 */
export function mergeHits(...sets: SearchHit[][]): SearchHit[] {
  const seen: Record<string, true> = {};
  const out: SearchHit[] = [];
  for (const set of sets) {
    for (const hit of set ?? []) {
      if (!hit) continue;
      const key = keyOf(hit);
      if (key.length === 0) {
        out.push(hit);
        continue;
      }
      if (seen[key]) continue;
      seen[key] = true;
      out.push(hit);
    }
  }
  return out;
}

/** Newest first. A blank date sorts last rather than throwing the order. */
export function sortByModified(hits: SearchHit[]): SearchHit[] {
  return [...(hits ?? [])].sort((a, b) => {
    const av = clean(a?.modified);
    const bv = clean(b?.modified);
    if (av.length === 0 && bv.length === 0) return 0;
    if (av.length === 0) return 1;
    if (bv.length === 0) return -1;
    return av < bv ? 1 : av > bv ? -1 : 0;
  });
}

/* ─────────────────────────────── Outcome ─────────────────────────────── */

/**
 * What happened to one library's read.
 *
 * `refused` is its own state and is NOT a failure. An uncleared user reading an HC library is
 * refused by SharePoint, and that is the correct answer to their query — counting it as an error
 * would put a permanent warning in front of every user who is not HC-cleared, training them to
 * ignore the banner on the day a real failure happens.
 */
export type ReadOutcome = "ok" | "refused" | "failed";

export interface LibraryResult {
  library: SearchLibrary;
  outcome: ReadOutcome;
  /** HTTP status when the read failed. Named on screen: a 404 and a 403 have opposite fixes. */
  status?: number;
  hits: SearchHit[];
}

/** The three empty states, kept distinct as everywhere else in this codebase. */
export type SearchState = "idle" | "results" | "empty" | "error";

/**
 * The state of a completed search.
 *
 * A read that FAILED can never render as a clean empty. Someone told "no results" when a library was
 * unreachable concludes the document does not exist — in My Submissions that made a person upload a
 * second copy; here it would make them conclude a record was never filed.
 *
 * Partial success is `results`, not `error`: there are rows to show. The caller must still surface
 * `failedLibraries` alongside them, or the list silently understates what is there.
 */
export function searchState(searched: boolean, results: LibraryResult[]): SearchState {
  if (!searched) return "idle";
  const list = results ?? [];
  const hits = list.reduce((n, r) => n + (r?.hits?.length ?? 0), 0);
  if (hits > 0) return "results";
  return list.some((r) => r?.outcome === "failed") ? "error" : "empty";
}

/** The libraries whose read failed — never those that merely refused. */
export function failedLibraries(results: LibraryResult[]): LibraryResult[] {
  return (results ?? []).filter((r) => r?.outcome === "failed");
}
