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
