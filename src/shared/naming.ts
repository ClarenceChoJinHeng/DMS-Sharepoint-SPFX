// Resolves the DMS/CRS name prefix for lists, content types and SharePoint groups.
//
// The client rebrands every DMS-named artefact to CRS when the solution reaches their
// site. Hardcoding either prefix is what this module exists to stop: a missed literal is a
// RUNTIME failure (getbytitle returns 404) rather than a compile error, so it surfaces
// during the client's UAT instead of in our build.
//
// SPFx-free by design — the SharePoint lookup is injected as a ListProbe — so the probe
// ordering, the caching and the group matching are all unit-testable in plain Jest, the
// same pattern as formModel and groupMapModel.
//
// See docs/superpowers/specs/2026-08-04-configurable-name-prefix-design.md
// and memory dms-to-crs-rename-pending.

/**
 * Candidate prefixes, in probe order. CRS first, so a fully renamed site costs one request
 * and a legacy site costs two. Adding a third client's prefix here is the whole onboarding
 * change.
 */
export const CANDIDATE_PREFIXES = ["CRS", "DMS"];

/** The historical prefix. Used when nothing resolves, so behaviour matches pre-2026-08-04. */
export const LEGACY_PREFIX = "DMS";

/**
 * List name suffixes. The prefix is discovered; these never change.
 *
 * Staging and Documents are absent on purpose: they are library names read from the config
 * list, and were never prefixed.
 */
export const LIST_SUFFIX = {
  config: "Config",
  groupMap: "Group Map",
  folderMap: "Folder Map",
  abbreviation: "Term Abbreviation",
  deletionLog: "Deletion Log",
  auditLog: "Audit Log",
  // Deletion and share requests a Head of Unit decides — 2026-08-15. NOT the same thing as
  // `deletionLog`, which was never built: that was a record of deletions that had happened, this is
  // the queue of ones being asked for.
  requests: "Requests",
  // One row per uploaded file, so My Submissions can still show a file after it is deleted.
  // Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
  submissions: "Submissions",
};

/**
 * Every list suffix `primeNames` probes. **Must cover all of `LIST_SUFFIX`** — pinned by test.
 *
 * ⚠ `LIST_SUFFIX.requests` WAS MISSING, from the day the requests feature was written until
 * 2026-08-20, and the failure was completely silent. An unprimed suffix never enters the name cache,
 * so `cachedListTitle` answers the LEGACY `DMS <suffix>` for ever — which on a CRS-renamed site 404s
 * on a list that demonstrably exists. Seen live: the admin created `CRS Requests`, and both the
 * Requests page and My Submissions went on reporting *"the requests list does not exist yet"*, which
 * then invited a second press of Create and an HTTP 500.
 *
 * Lives HERE rather than in `spNaming` so it can be tested: that module imports `@microsoft/sp-http`,
 * which the test environment cannot resolve.
 *
 * A suffix added here costs one probe per page load. A suffix left out costs a feature, silently.
 */
export const PRIMED_SUFFIXES: string[] = [
  LIST_SUFFIX.config,
  LIST_SUFFIX.groupMap,
  LIST_SUFFIX.folderMap,
  LIST_SUFFIX.abbreviation,
  LIST_SUFFIX.deletionLog,
  LIST_SUFFIX.auditLog,
  LIST_SUFFIX.requests,
  LIST_SUFFIX.submissions,
];

/** Returns true if a list with this exact title exists. Supplied by the caller. */
export type ListProbe = (candidateTitle: string) => Promise<boolean>;

/** Every title a given suffix could have, in probe order. */
export function candidateTitles(suffix: string): string[] {
  return CANDIDATE_PREFIXES.map((p) => `${p} ${suffix}`);
}

/**
 * Resolved titles, cached per suffix for the page session.
 *
 * Module-level rather than per-caller on purpose: six modules ask for "Group Map", and
 * none of them should each pay two probe requests for it.
 */
const cache = new Map<string, string>();

/** Test seam, and the escape hatch after a rename mid-session. */
export function clearNameCache(): void {
  cache.clear();
}

/**
 * Resolve one list's live title by probing its candidates in order.
 *
 * Resolution is PER SUFFIX, never one global prefix derived from the config list. The
 * client renames lists by hand, one at a time, so there is always a window in which
 * "CRS Config" exists and "CRS Group Map" does not — observed live on 2026-08-04, when
 * three lists were renamed and Config was not. A single global prefix would have resolved
 * "DMS" from Config and then 404d on all three renamed lists. Per-suffix resolution turns
 * a total outage into a partial rename, and makes the rename and the deployment
 * order-independent.
 *
 * When nothing resolves it returns the LEGACY title rather than throwing: every caller
 * already handles its own 404, and throwing here would take down a whole web part over one
 * absent list.
 */
export async function resolveListTitle(suffix: string, probe: ListProbe): Promise<string> {
  const cached = cache.get(suffix);
  if (cached !== undefined) return cached;
  for (const title of candidateTitles(suffix)) {
    let exists = false;
    try {
      exists = await probe(title);
    } catch {
      // A throttled or failed probe is NOT evidence that the list is absent. Try the next
      // candidate; if every probe fails we fall through to the legacy name, which is the
      // exact title the code used before this module existed.
      exists = false;
    }
    if (exists) {
      cache.set(suffix, title);
      return title;
    }
  }
  // NOT cached. A failed probe must not pin the legacy name for the rest of the session —
  // the next call gets to try again, which matters when the first attempt failed because
  // of throttling rather than absence.
  return `${LEGACY_PREFIX} ${suffix}`;
}

/**
 * The resolved title for a suffix, WITHOUT probing.
 *
 * Exists so call sites stay synchronous. Threading an `await` through every getbytitle() in the
 * codebase — around twenty-five of them — is a large diff in the most load-bearing files, and the
 * day before a deployment that is the wrong kind of change. Instead each web part primes the
 * cache once inside the async mount work it already does, and every call site after that is a
 * plain string swap.
 *
 * Returns the LEGACY title when the cache is cold, which is exactly the behaviour before this
 * module existed: a component that forgets to prime still works on a DMS-named site rather than
 * failing outright. A deliberate soft landing — `namesPrimed()` is how a caller checks.
 */
export function cachedListTitle(suffix: string): string {
  return cache.get(suffix) ?? `${LEGACY_PREFIX} ${suffix}`;
}

/**
 * The prefix this site actually uses, learned from whatever has already resolved.
 *
 * `cachedListTitle` answers `DMS <suffix>` for anything not in the cache — correct for READING,
 * because the legacy name is what the code used before this module existed and a caller handles its
 * own 404. It is WRONG for CREATING. A list that does not exist yet can never be in the cache, so a
 * self-provisioning page on a CRS-renamed site was about to create **`DMS Requests`** beside `CRS
 * Config`, `CRS Group Map` and the rest (seen live 2026-08-20).
 *
 * That is not cosmetic: the client renames every list at import, the `DMS_` prefix filter fails
 * SILENTLY (memory `dms-to-crs-rename-pending`), and a list nobody can find is a feature nobody can
 * use. It also strands the list on the next rename sweep, because it was created after it.
 *
 * Derived from a RESOLVED title rather than configured, so it needs no new setting and cannot
 * disagree with what is on the site. Falls back to `LEGACY_PREFIX` when nothing has resolved — a
 * fresh site with no CRS lists at all, where the legacy name is exactly what the old code produced.
 */
export function resolvedPrefix(): string {
  for (const suffix of Object.keys(LIST_SUFFIX)) {
    const key = (LIST_SUFFIX as Record<string, string>)[suffix];
    const title = cache.get(key);
    if (!title) continue;
    const cut = title.length - key.length;
    // Only a title that genuinely ENDS with its suffix tells us anything; anything else is a name
    // the probe matched some other way and guessing a prefix from it would invent one.
    if (cut > 1 && title.slice(cut) === key) return title.slice(0, cut - 1);
  }
  return LEGACY_PREFIX;
}

/**
 * The title a page should use when it CREATES a list for the first time.
 *
 * Prefers the resolved live title when one exists — creating is then a no-op the caller can detect —
 * and otherwise builds the name from the prefix this site is actually using.
 */
export function titleForNewList(suffix: string): string {
  return cache.get(suffix) ?? `${resolvedPrefix()} ${suffix}`;
}

/**
 * Record a title this session just CREATED, so later reads find it without re-probing.
 *
 * Without it, everything after a provision run keeps answering the legacy name until the page is
 * reloaded — the list exists and the very code that made it cannot see it.
 */
export function noteCreatedList(suffix: string, title: string): void {
  if (suffix && title) cache.set(suffix, title);
}

/** True once at least one title has resolved — i.e. the cache has been primed. */
export function namesPrimed(): boolean {
  return cache.size > 0;
}

/**
 * The prefix to use when WRITING a new name — a group title, or a content type.
 *
 * Derived from whichever config list answered, because that is the one artefact that must
 * exist for anything else to work at all.
 */
export async function resolveWritePrefix(probe: ListProbe): Promise<string> {
  const title = await resolveListTitle(LIST_SUFFIX.config, probe);
  return title.split(" ")[0] || LEGACY_PREFIX;
}

/** `<P> Folder` — the content type folders carry so Full Name reaches the details pane. */
export function folderContentTypeName(prefix: string): string {
  return `${prefix} Folder`;
}

/** `<P>_SITE_MEMBERS` — the site-entry group. */
export function siteEntryGroupName(prefix: string): string {
  return `${prefix}_SITE_MEMBERS`;
}

/** Every title the site-entry group could have, in preference order. */
export const SITE_ENTRY_CANDIDATES = CANDIDATE_PREFIXES.map(siteEntryGroupName);

/**
 * The site-entry group's live title.
 *
 * Resolved separately from the lists, and NOT derived from the list prefix, because the two
 * diverge in practice: verified live 2026-08-05, a site with CRS lists and CRS permission levels
 * still had DMS_SITE_MEMBERS.
 *
 * Getting this wrong is worse than a 404. The site-entry pass CREATES the group when it cannot
 * find one — so a stale name means a second, empty group is created, granted Read on the web, and
 * then fed every member by the self-heal, while everyone's real access sits in the original. Two
 * groups that both look correct, and no error anywhere.
 *
 * Defaults to the LEGACY name until resolved, matching the behaviour before this existed.
 */
let siteEntryName: string = siteEntryGroupName(LEGACY_PREFIX);

export function cachedSiteEntryName(): string {
  return siteEntryName;
}

/**
 * Record which candidate actually exists on this site.
 *
 * Takes the titles rather than performing the lookup, so this module stays SPFx-free. A set
 * matching nothing leaves the legacy name: creating `DMS_SITE_MEMBERS` on a site that has neither
 * is what the code did before, and is recoverable with a rename.
 */
export function setSiteEntryName(existingGroupTitles: string[]): void {
  const titles = (existingGroupTitles ?? []).map((t) => (t ?? "").trim().toLowerCase());
  for (const candidate of SITE_ENTRY_CANDIDATES) {
    if (titles.indexOf(candidate.toLowerCase()) !== -1) {
      siteEntryName = candidate;
      return;
    }
  }
}

/** `<P>_` — the prefix new group names are built from. */
export function groupPrefix(prefix: string): string {
  return `${prefix}_`;
}

/**
 * Does this group title belong to us?
 *
 * Accepts EVERY candidate prefix, always — never just the resolved one. This drives group
 * search and, more importantly, the site-entry membership sync. Honouring only the resolved
 * prefix would mean a half-renamed site (some groups CRS_, some still DMS_) syncs one half
 * and silently ignores the other, producing "only SOME users cannot open the site" — the
 * hardest version of this bug to diagnose, and one that reports success while doing it.
 *
 * Matching both costs nothing: a group named for another client's prefix will not exist on
 * this site.
 */
export function matchesAnyGroupPrefix(title: string): boolean {
  const t = (title ?? "").trim().toUpperCase();
  return CANDIDATE_PREFIXES.some((p) => t.indexOf(`${p}_`) === 0);
}

/* ---------------------------------------------------------------------------
 * The upload/approval library.
 *
 * Unlike the config lists this needs TWO names, and they are not interchangeable:
 *
 *   title       -> getbytitle('...')      — 4 call sites in ApprovalDocument, 1 in FolderManager
 *   urlSegment  -> split('/<segment>/')   — 5 call sites that slice a server-relative path
 *
 * They were identical while the library was called "Staging", which is exactly why the old
 * code could hardcode one literal for both. Verified live 2026-08-06 they now differ: the
 * library was created as `ApprovalDocument` for a clean URL and then retitled to
 * `Approval Document`, giving title "Approval Document" and URL ".../ApprovalDocument".
 *
 * Keeping them as one resolved PAIR is deliberate. A wrong title 404s — loud, obvious. A
 * wrong url segment does not: `split("/Staging/")` on a path with no such segment yields a
 * one-element array, so the caller reads `undefined` and carries on, routing a file nowhere
 * while reporting success. The pair can only be right or absent, never half-right.
 * ------------------------------------------------------------------------- */

/**
 * Candidate titles, in probe order: the current name, the pre-retitle name, then the
 * historical one. `Staging` last so an un-migrated site still works untouched.
 */
export const LIBRARY_CANDIDATES = [
  /* ⚠ THE CLIENT RENAMES THESE LIBRARIES AS A MATTER OF COURSE - twice in one afternoon on
     2026-08-27, and again on 2026-08-28. Verified live each time with
     `/_api/web/lists?$select=Title&$expand=RootFolder&$filter=BaseTemplate eq 101`, which is the
     only authority: a rename NEVER changes the URL, so `RootFolder/ServerRelativeUrl` is what
     identifies a library through any number of retitles.
     A title outside this array fails SILENTLY - the library simply stops resolving. */
  "Approval for Document",
  "Approval Document",
  "ApprovalDocument",
  "Staging",
];

/** The historical name, used for both halves until a probe resolves them. */
export const LEGACY_LIBRARY = "Staging";

export interface LibraryNames {
  /** List title, for getbytitle(). */
  title: string;
  /** Final path segment of the library root, for path splitting. No slashes. */
  urlSegment: string;
}

let libraryNames: LibraryNames = { title: LEGACY_LIBRARY, urlSegment: LEGACY_LIBRARY };

export function cachedLibrary(): LibraryNames {
  return libraryNames;
}

/** Convenience readers, so call sites stay short. */
export const libraryTitle = (): string => libraryNames.title;
export const libraryUrlSegment = (): string => libraryNames.urlSegment;

/**
 * Translate a LOGICAL LibTarget key into the title the SharePoint API answers to.
 *
 * `"Staging"` is a logical key, not a name: it is the stored `Target` value on Group Map
 * library-scope rows and the key in TARGETS, and it must keep being written and filtered on
 * unchanged. But the library it names is titled `Approval Document` on this site, so any URL built
 * from the key 404s. Only the API boundary translates; stored data never does.
 *
 * Lived as a private const in FolderManager until 2026-08-14, when StagingAccess was found issuing
 * `getbytitle('Staging')` for its live-ACL reads — leaving every "Access now" cell reading
 * "unknown" and both Allow and Remove unable to apply a permission. That is the failure gotcha #12
 * describes: a wrong title 404s, which at least fails loudly, but only if somebody is looking at
 * the column it fails in. One copy now, because the next screen to need it would have repeated the
 * same mistake.
 *
 * Four keys since 2026-08-15. `StagingHC` and `DocumentsHC` are logical keys in exactly the same
 * way `Staging` is, and are translated here and nowhere else. An HC key with no resolved pair falls
 * through to the key itself, which 404s LOUDLY — the deliberate choice over returning the normal
 * library's title, which would resolve happily and route a Highly Confidential document into the
 * library its whole unit reads.
 */
export const libApiTitle = (lib: string): string => {
  if (lib === "Staging") return libraryTitle();
  /* ⚠ ADDED 2026-08-28, AND ITS ABSENCE WAS THE WHOLE BUG. `Documents` used to fall through to
     `return lib` at the bottom — correct for as long as the library was actually TITLED `Documents`,
     and silently wrong the moment the client retitled it to `Restricted & Confidential Document`.

     This function IS the API boundary: every `getbytitle()` for that library goes through it. So
     fixing `libraryTargets()` and `allLibraryTitles()` without fixing this left reconciliation
     404ing on every folder and grant in the approved side — while the two functions that had been
     fixed reported the right name, which made it look repaired.

     The tell was on screen: the reconciliation progress panel titles its sections with
     `libApiTitle`, and that one still read `DOCUMENTS` where the other five read their live names. */
  if (lib === "Documents") return documentsLibraryTitle();
  const hc = cachedHcLibraries();
  if (lib === "StagingHC") return hc ? hc.approval.title : lib;
  if (lib === "DocumentsHC") return hc ? hc.documents.title : lib;
  // The archive pair, 2026-08-22 — same rule and same reason: an unresolved key falls through to
  // ITSELF and 404s loudly. Never to `Documents`, which would resolve happily and move a document
  // out of the archive workflow into the live library it was supposed to be leaving.
  const arc = cachedArchiveLibraries();
  if (lib === "Archive") return arc ? arc.normal.title : lib;
  if (lib === "ArchiveHC") return arc && arc.hc ? arc.hc.title : lib;
  return lib;
};

/**
 * Record the resolved pair. Takes the values rather than performing the lookup, keeping this
 * module free of SPFx imports — the same split as setSiteEntryName.
 *
 * The url segment is derived from the library's RootFolder rather than from its title,
 * because the title is the half that changes. A blank or slash-only segment is rejected: it
 * would reduce `split("//")` to matching every separator in the document tree.
 */
export function setLibraryNames(title: string, serverRelativeUrl: string): void {
  const t = (title ?? "").trim();
  const segment = (serverRelativeUrl ?? "").split("/").filter(Boolean).pop() ?? "";
  if (t.length === 0 || segment.length === 0) return;
  libraryNames = { title: t, urlSegment: segment };
}

/** Test seam, and the escape hatch after a rename mid-session. */
export function clearLibraryCache(): void {
  libraryNames = { title: LEGACY_LIBRARY, urlSegment: LEGACY_LIBRARY };
}

/* ---------------------------------------------------------------------------
 * The Highly Confidential pair.
 *
 * Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
 *
 * Everything above resolves ONE library and falls back to a legacy name when nothing answers.
 * This resolves TWO, and falls back to NOTHING — that difference is the whole safety of the
 * feature and is not an oversight to be tidied up later.
 *
 * An unresolved normal library costs a screen. An unresolved HC library, if it fell back to
 * anything, would file a Highly Confidential document into the library every PIC in the unit reads
 * once it is approved — silently, under a green success toast. So `undefined` here means the level
 * is not offered at all, and callers must treat it that way.
 * ------------------------------------------------------------------------- */

/**
 * Probe order: the retitled name first, then the name it is created under.
 *
 * WARN: `Highly Confidential Document` IS A REAL LIVE TITLE, verified by REST on ClarenceDMSTesting
 * 2026-08-27 - and it was in NEITHER list, so resolution depended on `getbytitle` happening to match
 * the no-space creation name (which is also the URL name). It works today; it is not something to
 * rely on. Listing the live title makes the match explicit.
 *
 * THE STANDING RULE THIS EXPOSES: **this client renames libraries, and a title outside these arrays
 * fails SILENTLY** - `hcAvailable()` goes false, the Highly Confidential level stops being offered,
 * and HC uploads simply stop with no error anywhere. Check the live titles with
 * `/_api/web/lists?$select=Title&$filter=BaseTemplate eq 101` when provisioning ANY site, SDG
 * included, and add whatever is there.
 *
 * Appended rather than prepended, so the probe count on sites already using the short names is
 * unchanged.
 */
export const HC_APPROVAL_CANDIDATES = [
  // Live on ClarenceDMSTesting since 2026-08-28, at the unchanged URL /HCApprovalDocument.
  "Approval for Highly Confidential Document",
  "HC Approval Document",
  "HC Approval Documents",
  "HCApprovalDocument",
  "Highly Confidential Approval Document",
];
export const HC_DOCUMENTS_CANDIDATES = [
  "HC Documents",
  /* SINGULAR, and it is the live title on ClarenceDMSTesting since 2026-08-27. It is also the house
     style - `Approval Document` and `HC Approval Document` are both singular; `Documents` is plural
     only because SharePoint named it. Listed second so the common spelling is still found first. */
  "HC Document",
  "HCDocuments",
  "Highly Confidential Document",
  "Highly Confidential Documents",
];

/**
 * The normal approved-side library.
 *
 * Its title is `Documents` while its URL segment is `Shared Documents` — the same title/URL split as
 * the approval library, and the reason "Shared Documents" once appeared at the head of every approved
 * file's folder trail. Named here so the four library keys read from one place.
 */
export const DOCUMENTS_LIBRARY = "Documents";

/**
 * Probe order for the normal approved-side library.
 *
 * ⚠⚠ THIS LIBRARY WAS THE ONE THAT WAS NEVER PROBED, AND THAT WAS A LATENT BUG THE 2026-08-28
 * RENAME FINALLY FIRED. Every other library resolves through `probeLibrary`; this one was the bare
 * constant above, used directly in `libraryTargets`, `allLibraryTitles` and `DocumentSearch`. So the
 * day the client retitled it to `Restricted & Confidential Document`, every `getbytitle('Documents')`
 * began 404ing — with no candidate array to extend, because there was no probe to feed.
 *
 * `Documents` stays FIRST: SDG's live site still uses it, and probing in this order means an
 * un-renamed site resolves on the first request exactly as before.
 *
 * ⚠ `Shared Documents` is the URL name, not a title, and it is listed LAST as a deliberate
 * long-stop: `getbytitle` matches the TITLE, so it can only ever hit on a site where somebody
 * retitled the library to match its own URL. Harmless where it does not apply.
 */
export const DOCUMENTS_CANDIDATES = [
  "Documents",
  // Live on ClarenceDMSTesting since 2026-08-28, at the unchanged URL /Shared Documents — which is
  // what PROVED it is this library renamed rather than a new confidentiality tier.
  "Restricted & Confidential Document",
  "Restricted and Confidential Document",
  "Shared Documents",
];

/**
 * The resolved title, or the legacy literal until a probe settles it.
 *
 * ⚠ FALLS BACK TO `Documents` RATHER THAN TO NOTHING, unlike the HC and archive pairs. Those fail
 * closed because their ABSENCE is meaningful — a site may legitimately have no HC libraries, and
 * guessing one into existence would file a secret in the open. This library always exists, so the
 * useful failure is the loud one: a wrong title 404s, where `undefined` would silently drop the
 * approved side out of `libraryTargets` and take reconciliation, the migrator and CRS Search with it.
 */
let documentsLibraryName: string = DOCUMENTS_LIBRARY;

/**
 * The live title of the normal approved-side library.
 *
 * ⚠ NEVER CAPTURE THIS IN A RENDER-TIME `const`. It reads a module cache primed by `primeNames`,
 * which has NOT run on the first render — so the value a first render closes over is the legacy
 * literal, and a later render showing the right name does not re-issue the request that used the
 * wrong one. That is the 2026-08-14 `StagingAccess` defect exactly (`listBase` had to become a
 * function). Fine for DISPLAY; call it at request time for anything that builds a URL.
 */
export function documentsLibraryTitle(): string {
  return documentsLibraryName;
}

/** Set by `primeDocumentsLibrary`. A blank title is ignored — see the fallback note above. */
export function setDocumentsLibraryName(title: string): void {
  const t = (title ?? "").trim();
  if (t.length > 0) documentsLibraryName = t;
}

/** Test seam, and the escape hatch after a rename mid-session. */
export function clearDocumentsLibraryName(): void {
  documentsLibraryName = DOCUMENTS_LIBRARY;
}

/** Its URL segment, which is NOT its title — gotcha #12, and the reason both are named here. */
export const DOCUMENTS_URL_SEGMENT = "Shared Documents";

/**
 * Probe order for the seven-year archive pair. Created without a space so the URL stays clean, then
 * retitled — gotcha #12 — which is why each has two candidates and the retitled name is tried first.
 *
 * Spec: docs/superpowers/specs/2026-08-22-seven-year-archive-design.md
 */
export const ARCHIVE_CANDIDATES = [
  "Archive",
  // Live on ClarenceDMSTesting since 2026-08-28, at the unchanged URL /Archive. A plain `&`, NOT
  // the fullwidth `＆` the term store requires - confirmed from the REST payload (`&amp;`).
  "Archive Restricted & Confidential Document",
  "CRSArchive",
];
/** `Highly Confidential Archive Document` is the live title on ClarenceDMSTesting - see the note on
 *  HC_DOCUMENTS_CANDIDATES for why an unlisted title fails silently. */
export const ARCHIVE_HC_CANDIDATES = [
  /* ⚠ NOTE THE WORD ORDER. `Archive Highly Confidential Document` is the live title;
     `Highly Confidential Archive Document` below is a DIFFERENT string and was the 2026-08-27
     name. Same words, different order is a MISS, not a near-miss the probe can absorb. */
  "Archive Highly Confidential Document",
  "HC Archive",
  "HC Archives",
  "HCArchive",
  "Highly Confidential Archive Document",
  "Highly Confidential Archive",
];

/**
 * The logical library keys. `Staging` is the approval library — a LOGICAL name kept because it
 * is also the stored `Target` value on Group Map library-scope rows; the real title is resolved at
 * the API boundary by `libApiTitle`.
 *
 * `Archive`/`ArchiveHC` joined them 2026-08-22. They are logical keys in the same way, and like the
 * HC pair they resolve to NOTHING when the libraries are absent rather than to a legacy literal.
 *
 * ⚠ WIDENING THIS UNION IS A CHANGE SITE FOR EVERY EXHAUSTIVE READER. `Record<LibTarget, …>` fails
 * to compile until it gains the new keys, which is the loud and safe direction — the alternative,
 * a partial map, would silently skip a library in whichever pass indexed it.
 */
export type LibTarget =
  | "Staging"
  | "Documents"
  | "StagingHC"
  | "DocumentsHC"
  | "Archive"
  | "ArchiveHC";

/** A logical key with the two real names it resolves to. */
export interface LibraryTarget extends LibraryNames {
  key: LibTarget;
}

/**
 * Every library this site actually has, resolved — TWO without the HC pair, FOUR with it.
 *
 * ⚠ EXISTS BECAUSE A TWO-ELEMENT LITERAL WAS THE BUG (register #15, found 2026-08-19).
 * `SubtreeMigrator` carried `[{ key: "Staging" … }, { key: "Documents" … }]` written on 2026-08-11,
 * four days before the HC pair existed, and was never revisited. So a structure change migrated the
 * normal libraries, applied the new chain, and left every HC document in the old shape — reporting
 * success, because the guard that refuses to apply a pending chain re-scans only the libraries it
 * knows about, and a scan that never looks at HC always finds no drift there.
 *
 * Derived, never literal, and shared so the migrator and reconciliation cannot disagree about which
 * libraries exist. `hcAvailable()` fails CLOSED by design — `naming.ts` has no HC fallback, so an
 * unresolved HC library is `undefined` rather than a legacy literal, and a site without HC gets two.
 */
export function libraryTargets(): LibraryTarget[] {
  const base: LibraryTarget[] = [
    { key: "Staging", title: libraryTitle(), urlSegment: libraryUrlSegment() },
    { key: "Documents", title: documentsLibraryTitle(), urlSegment: DOCUMENTS_URL_SEGMENT },
  ];
  const hc = cachedHcLibraries();
  const withHc = hc
    ? base.concat([
        { key: "StagingHC", title: hc.approval.title, urlSegment: hc.approval.urlSegment },
        { key: "DocumentsHC", title: hc.documents.title, urlSegment: hc.documents.urlSegment },
      ])
    : base;

  /* ⚠ THE ARCHIVE IS INCLUDED, AND REGISTER #15 IS WHY. `SubtreeMigrator` reads this to decide which
     libraries a structure change must re-shape, and the HC pair was omitted from a two-element
     literal for four days — so a migration reported success while every HC document stayed in the
     old shape, because a scan that never looks at a library always finds no drift there. An archive
     left out of this list would repeat that exactly: the archive keeps the pre-change folder shape
     for ever, and nothing reports it. Archived files are read-only for GROUPS; the migrator runs as
     an admin, so it can still move them. */
  const arc = cachedArchiveLibraries();
  if (!arc) return withHc;
  const archive: LibraryTarget[] = [
    { key: "Archive", title: arc.normal.title, urlSegment: arc.normal.urlSegment },
  ];
  if (arc.hc) archive.push({ key: "ArchiveHC", title: arc.hc.title, urlSegment: arc.hc.urlSegment });
  return withHc.concat(archive);
}

export interface HcLibraries {
  approval: LibraryNames;
  documents: LibraryNames;
}

/** `undefined` until BOTH halves resolve. Never a partial pair — see `setHcLibraryNames`. */
let hcLibraries: HcLibraries | undefined;

export function cachedHcLibraries(): HcLibraries | undefined {
  return hcLibraries;
}

/**
 * Is Highly Confidential usable on this site?
 *
 * The single question every caller should ask before offering the level. False covers both "not on
 * this site" and "not probed yet", deliberately: from the point of view of whether it is safe to
 * offer the level, those are the same answer, and a caller that distinguished them would eventually
 * offer the level during the window before the probe finished.
 */
export function hcAvailable(): boolean {
  return hcLibraries !== undefined;
}

/**
 * Record the resolved HC pair. Both halves, or nothing.
 *
 * BOTH LIBRARIES OR NEITHER. An HC approval library with no HC documents library is worse than no HC
 * at all: uploads succeed, approval succeeds, and then auto-route has nowhere to put the file — so a
 * document that everyone believes is filed and protected is sitting in a queue nobody watches. The
 * pair can be right or absent, never half-right, exactly as `title` and `urlSegment` can.
 */
export function setHcLibraryNames(
  approvalTitle: string,
  approvalUrl: string,
  documentsTitle: string,
  documentsUrl: string,
): void {
  const pair = (title: string, url: string): LibraryNames | undefined => {
    const t = (title ?? "").trim();
    // Same rule as setLibraryNames: a blank segment would reduce split("//") to matching every
    // separator in the document tree.
    const segment = (url ?? "").split("/").filter(Boolean).pop() ?? "";
    return t.length === 0 || segment.length === 0 ? undefined : { title: t, urlSegment: segment };
  };
  const approval = pair(approvalTitle, approvalUrl);
  const documents = pair(documentsTitle, documentsUrl);
  if (!approval || !documents) return;
  hcLibraries = { approval, documents };
}

/** Test seam, and the escape hatch after a rename mid-session. */
export function clearHcLibraryNames(): void {
  hcLibraries = undefined;
}

/* ---------------------------------------------------------------------------
 * The seven-year archive pair.
 *
 * Spec: docs/superpowers/specs/2026-08-22-seven-year-archive-design.md
 *
 * Same no-fallback rule as HC, for a related reason: an unresolved archive must never resolve to
 * `Documents`, or the mover would "archive" a document by writing it back into the live library it
 * was leaving — and then delete the source.
 *
 * ⚠ BUT THE PAIRING RULE IS DIFFERENT FROM HC's, AND DELIBERATELY SO. HC is both-or-neither because
 * either half alone is a trap. Here the archive MIRRORS WHATEVER HC STATE THE SITE HAS:
 *
 *   · a site WITH HC needs BOTH `Archive` and `HC Archive` — resolving only the first would leave
 *     HC documents silently never archived, with nothing reporting it
 *   · a site WITHOUT HC needs only `Archive`, and requiring `HC Archive` there would disable
 *     archiving on a site that is correctly provisioned
 *
 * Copying HC's flat both-or-neither rule would have been the obvious thing and is wrong in the
 * second case.
 * ------------------------------------------------------------------------- */

export interface ArchiveLibraries {
  normal: LibraryNames;
  /** Present only on a site that has the HC pair. `undefined` is correct there, not a partial read. */
  hc?: LibraryNames;
}

/** `undefined` until the pair resolves to whatever shape this site requires. */
let archiveLibraries: ArchiveLibraries | undefined;

export function cachedArchiveLibraries(): ArchiveLibraries | undefined {
  return archiveLibraries;
}

/**
 * Is the seven-year archive usable on this site?
 *
 * As with `hcAvailable`, false covers both "not on this site" and "not probed yet" — from the point
 * of view of whether it is safe to move a document, those are the same answer.
 */
export function archiveAvailable(): boolean {
  return archiveLibraries !== undefined;
}

/**
 * Record the resolved archive pair.
 *
 * `hcRequired` is passed in rather than read from `hcAvailable()` so the rule is testable without
 * priming the HC cache, and so the caller's view of HC and this one's cannot differ mid-probe.
 *
 * Refuses a partial pair on an HC site. A blank title or segment is refused for the same reason as
 * everywhere else here: a blank segment reduces `split("//")` to matching every separator in the
 * document tree.
 */
export function setArchiveLibraryNames(
  normalTitle: string,
  normalUrl: string,
  hcRequired: boolean,
  hcTitle?: string,
  hcUrl?: string,
): void {
  const pair = (title: string | undefined, url: string | undefined): LibraryNames | undefined => {
    const t = (title ?? "").trim();
    const segment = (url ?? "").split("/").filter(Boolean).pop() ?? "";
    return t.length === 0 || segment.length === 0 ? undefined : { title: t, urlSegment: segment };
  };
  const normal = pair(normalTitle, normalUrl);
  if (!normal) return;
  const hc = pair(hcTitle, hcUrl);
  // On an HC site, half an archive is no archive: the missing half is where the confidential
  // documents would have gone, and the failure would be that they quietly never move.
  if (hcRequired && !hc) return;
  archiveLibraries = hcRequired ? { normal, hc } : { normal };
}

/** Test seam, and the escape hatch after a rename mid-session. */
export function clearArchiveLibraryNames(): void {
  archiveLibraries = undefined;
}

/**
 * Where an archived document belongs, given the library it is leaving.
 *
 * ⚠ FAILS CLOSED — `undefined` rather than a fallback, exactly as `hcRouting.ts` does. A fallback
 * here would move a Highly Confidential document into the open archive, which is the one outcome
 * this whole pair exists to prevent. The caller must refuse the move, never improvise a destination.
 *
 * Only the two APPROVED-side libraries have an archive. A document still in an approval library was
 * never approved, and an unapproved draft is not a record — §10 of the spec.
 */
export function archiveTargetFor(lib: LibTarget): LibTarget | undefined {
  const arc = cachedArchiveLibraries();
  if (!arc) return undefined;
  if (lib === "Documents") return "Archive";
  if (lib === "DocumentsHC") return arc.hc ? "ArchiveHC" : undefined;
  return undefined;
}

/**
 * Every library title a metadata COLUMN must exist in — two, or four with HC.
 *
 * The standing warning, now with twice the surface: a column present in one library and absent from
 * another fails in two ways that name no column. Bulk upload writes its fields unconditionally and
 * ONE unknown field name fails the WHOLE `validateUpdateListItem` call, so every column is lost
 * rather than the missing one; and auto-route's copy carries over only columns that EXIST at the
 * destination, so the rest vanish with no error and a green run. `Documents` was missing four
 * columns for months on ClarenceDMSTesting with nothing reporting it.
 *
 * A FUNCTION, so the HC pair is included only once primeNames has resolved it. Callers that create
 * columns must use this rather than a two-element literal — that literal is the bug, waiting.
 */
export function allLibraryTitles(): string[] {
  const hc = cachedHcLibraries();
  const base = [libraryTitle(), documentsLibraryTitle()];
  const withHc = hc ? [...base, hc.approval.title, hc.documents.title] : base;

  /* ⚠ THE ARCHIVE NEEDS THE FULL COLUMN SET, NOT A SUBSET. A move carries over only columns that
     EXIST at the destination; the rest vanish with no error and a green run. That is the
     Remark/LegallyPrivileged gap of 2026-08-10, which stripped a legal marker off approved documents
     for months — and it would be worse here, because an archived document is a RECORD and nobody
     opens it for years. Derived from `libraryTargets()` so this list and that one cannot disagree. */
  const arc = cachedArchiveLibraries();
  if (!arc) return withHc;
  return arc.hc ? [...withHc, arc.normal.title, arc.hc.title] : [...withHc, arc.normal.title];
}
