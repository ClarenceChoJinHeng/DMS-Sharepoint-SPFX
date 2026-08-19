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
};

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
export const LIBRARY_CANDIDATES = ["Approval Document", "ApprovalDocument", "Staging"];

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
  const hc = cachedHcLibraries();
  if (lib === "StagingHC") return hc ? hc.approval.title : lib;
  if (lib === "DocumentsHC") return hc ? hc.documents.title : lib;
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

/** Probe order: the retitled name first, then the name it is created under. */
export const HC_APPROVAL_CANDIDATES = ["HC Approval Document", "HCApprovalDocument"];
export const HC_DOCUMENTS_CANDIDATES = ["HC Documents", "HCDocuments"];

/**
 * The normal approved-side library.
 *
 * Its title is `Documents` while its URL segment is `Shared Documents` — the same title/URL split as
 * the approval library, and the reason "Shared Documents" once appeared at the head of every approved
 * file's folder trail. Named here so the four library keys read from one place.
 */
export const DOCUMENTS_LIBRARY = "Documents";

/** Its URL segment, which is NOT its title — gotcha #12, and the reason both are named here. */
export const DOCUMENTS_URL_SEGMENT = "Shared Documents";

/**
 * The four logical library keys. `Staging` is the approval library — a LOGICAL name kept because it
 * is also the stored `Target` value on Group Map library-scope rows; the real title is resolved at
 * the API boundary by `libApiTitle`.
 */
export type LibTarget = "Staging" | "Documents" | "StagingHC" | "DocumentsHC";

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
    { key: "Documents", title: DOCUMENTS_LIBRARY, urlSegment: DOCUMENTS_URL_SEGMENT },
  ];
  const hc = cachedHcLibraries();
  if (!hc) return base;
  return base.concat([
    { key: "StagingHC", title: hc.approval.title, urlSegment: hc.approval.urlSegment },
    { key: "DocumentsHC", title: hc.documents.title, urlSegment: hc.documents.urlSegment },
  ]);
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
  const base = [libraryTitle(), DOCUMENTS_LIBRARY];
  return hc ? [...base, hc.approval.title, hc.documents.title] : base;
}
