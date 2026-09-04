// The SharePoint half of name resolution. Kept separate from naming.ts so that module stays
// free of @microsoft/* imports and remains unit-testable in plain Jest — the same split as
// spGroupsFilter.ts (pure) vs spGroups.ts (REST).
//
// See docs/superpowers/specs/2026-08-04-configurable-name-prefix-design.md

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  LIST_SUFFIX,
  ListProbe,
  resolveListTitle,
  resolveWritePrefix,
  folderContentTypeName,
  siteEntryGroupName,
  setSiteEntryName,
  SITE_ENTRY_CANDIDATES,
  LIBRARY_CANDIDATES,
  setLibraryNames,
  HC_APPROVAL_CANDIDATES,
  HC_DOCUMENTS_CANDIDATES,
  DOCUMENTS_CANDIDATES,
  setDocumentsLibraryName,
  setHcLibraryNames,
  hcAvailable,
  ARCHIVE_CANDIDATES,
  ARCHIVE_HC_CANDIDATES,
  setArchiveLibraryNames,
  PRIMED_SUFFIXES,
} from "./naming";

/**
 * A probe that asks SharePoint whether a list with this exact title exists.
 *
 * `$select=Id` keeps the response tiny — existence is the only question. A non-ok response
 * means "not under this title", which is all the caller needs: the distinction between 404
 * and 403 does not change the answer, since a list we cannot read is a list we cannot use.
 */
export function makeListProbe(sp: SPHttpClient, siteUrl: string): ListProbe {
  return async (title: string): Promise<boolean> => {
    const res: SPHttpClientResponse = await sp.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    return res.ok;
  };
}

/**
 * Resolve every list title once, so later call sites can read them synchronously.
 *
 * Call this at the START of a web part's existing mount work, before any list read. Sequential
 * rather than parallel: Promise.allSettled is unavailable on this tsconfig target (CLAUDE.md #3),
 * five suffixes is at most ten cheap existence probes, and doing them in order keeps the failure
 * of one from being confused with the failure of another.
 *
 * Never throws. A probe that fails leaves that suffix unresolved, and `cachedListTitle` then
 * returns the legacy name — the behaviour the codebase had before any of this existed. A naming
 * failure must not stop a web part from loading.
 */
/**
 * Resolved once per page session. Unlike the list titles there is no per-suffix cache to lean on,
 * and primeNames() is called from several loaders, so without this the group query would repeat.
 */
let siteEntryLookup: Promise<void> | undefined;

/**
 * Which of CRS_SITE_MEMBERS / DMS_SITE_MEMBERS exists on this site.
 *
 * Folded into primeNames so every component that already primes gets it for free — the site-entry
 * name is consulted in five places across four components, and a component that resolved its lists
 * but not this would create a duplicate group.
 *
 * Memoised on the PROMISE, not a boolean. primeNames is called from several loaders that start
 * concurrently, and a flag set after the await would let two of them issue the request before
 * either recorded that it had. Sharing the promise makes concurrent callers await one lookup.
 */
async function primeSiteEntry(sp: SPHttpClient, siteUrl: string): Promise<void> {
  if (!siteEntryLookup) {
    siteEntryLookup = (async () => {
      try {
        /* WARN: FILTERED SERVER-SIDE, NOT PAGED-THEN-SEARCHED. This read was `$top=500` on a site
           that holds 684 groups, and `setSiteEntryName` leaves the LEGACY `DMS_SITE_MEMBERS` when it
           finds no candidate - so a truncated read is indistinguishable from the group being absent,
           and every consumer of `siteEntryGroupTitle()` then names a group that does not exist. That
           includes reconciliation's "NOBODY CAN OPEN THE SITE HOME PAGE" warning, which would tell an
           admin to grant Read to `DMS_SITE_MEMBERS` on a site whose group is `CRS_SITE_MEMBERS`.

           FIFTH instance of the capped-read trap (memory `sp-capped-read-reads-as-absent`), after
           `fetchAllSiteGroups`, `libraryHasRefColumns`, `ensureColumn` and `BulkGroupProvisioner`.
           At most two rows come back, so this one cannot be truncated at all. */
        const clauses = SITE_ENTRY_CANDIDATES.map(
          (c) => `Title eq '${c.replace(/'/g, "''")}'`,
        ).join(" or ");
        let titles: string[] | undefined;
        const res: SPHttpClientResponse = await sp.get(
          `${siteUrl}/_api/web/sitegroups?$select=Title&$filter=${encodeURIComponent(clauses)}`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (res.ok) {
          const data = await res.json();
          titles = ((data.value ?? []) as Array<{ Title?: string }>).map((g) => g.Title ?? "");
        }
        if (titles === undefined) {
          /* Insurance against the FIX. If `$filter` on Title were ever rejected, falling through to
             a capped read would restore the original silent failure - so the fallback is 5000, and
             `undefined` (not read) is kept distinct from `[]` (read, no candidate). */
          const all: SPHttpClientResponse = await sp.get(
            `${siteUrl}/_api/web/sitegroups?$select=Title&$top=5000`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (!all.ok) return;
          const data = await all.json();
          titles = ((data.value ?? []) as Array<{ Title?: string }>).map((g) => g.Title ?? "");
        }
        setSiteEntryName(titles);
      } catch {
        // Leave the legacy name; a failed probe must not stop a web part loading.
      }
    })();
  }
  return siteEntryLookup;
}

/**
 * Which of the candidate titles the upload/approval library has, and the URL segment that
 * goes with it.
 *
 * Both halves come from ONE request per candidate, because they must agree: the response
 * carries the Title and RootFolder/ServerRelativeUrl together, so there is no window in
 * which a resolved title is paired with a stale segment.
 *
 * Memoised on the PROMISE for the same reason as primeSiteEntry — several loaders call
 * primeNames concurrently, and a flag set after the await would let two of them issue the
 * probes before either recorded that it had.
 */
let libraryLookup: Promise<void> | undefined;

async function primeLibrary(sp: SPHttpClient, siteUrl: string): Promise<void> {
  if (!libraryLookup) {
    libraryLookup = (async () => {
      for (const candidate of LIBRARY_CANDIDATES) {
        try {
          const res: SPHttpClientResponse = await sp.get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(candidate)}')` +
              `?$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (!res.ok) continue; // not under this title; try the next
          const data = await res.json();
          setLibraryNames(data?.Title ?? candidate, data?.RootFolder?.ServerRelativeUrl ?? "");
          return;
        } catch {
          // A network failure on one candidate must not stop the others being tried.
        }
      }
      // Nothing matched: leave the legacy pair, which is what the code used before this existed.
    })();
  }
  return libraryLookup;
}

/**
 * The Highly Confidential pair, if this site has one.
 *
 * Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md
 *
 * UNLIKE primeLibrary, absence is a normal outcome rather than a failure: most sites will not have
 * HC libraries, and `hcAvailable()` false simply means the level is never offered. So this resolves
 * nothing and leaves the cache undefined, with no fallback of any kind.
 *
 * BOTH halves are probed and BOTH must answer. `setHcLibraryNames` enforces it, but the pairing
 * matters enough to be visible here too: an HC approval library with no HC documents library accepts
 * uploads and approvals and then has nowhere to route them.
 */
/**
 * Try each candidate title in turn; return the first that answers, with its title AND url from the
 * SAME response.
 *
 * That pairing is the point: a resolved title must never be matched with a stale segment, because a
 * wrong title 404s loudly while a wrong URL segment fails SILENTLY (gotcha #12).
 *
 * ONE implementation, shared by the HC pair and the archive pair. It began as a closure inside
 * `primeHcLibraries`; the archive needed exactly the same probe, and a second copy is how the two
 * would come to disagree about what "resolved" means.
 */
async function probeLibrary(
  sp: SPHttpClient,
  siteUrl: string,
  candidates: string[],
): Promise<{ title: string; url: string } | undefined> {
  for (const candidate of candidates) {
    try {
      const res: SPHttpClientResponse = await sp.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(candidate)}')` +
          `?$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) continue; // not under this title; try the next
      const data = await res.json();
      return { title: data?.Title ?? candidate, url: data?.RootFolder?.ServerRelativeUrl ?? "" };
    } catch {
      // A network failure on one candidate must not stop the others being tried.
    }
  }
  return undefined;
}

/**
 * The normal approved-side library's TITLE.
 *
 * ⚠ ITS URL SEGMENT IS NOT PROBED, AND MUST NOT BE. `Shared Documents` is fixed by SharePoint at
 * creation and a rename never changes it (gotcha #12) — verified again on 2026-08-28, when the
 * title became `Restricted & Confidential Document` while the URL stayed put. Thirteen call sites
 * use `DOCUMENTS_URL_SEGMENT` and every one of them is still correct.
 *
 * ⚠ UNLIKE THE HC AND ARCHIVE PROBES, FAILURE IS NOT A NORMAL OUTCOME. Those may legitimately find
 * nothing; this library always exists, so a failed probe leaves the legacy `Documents` literal and
 * the next `getbytitle` 404s loudly. That is the right way round — silently dropping the approved
 * side would take reconciliation, the migrator and CRS Search with it.
 */
let documentsLookup: Promise<void> | undefined;

async function primeDocumentsLibrary(sp: SPHttpClient, siteUrl: string): Promise<void> {
  if (!documentsLookup) {
    documentsLookup = (async () => {
      const found = await probeLibrary(sp, siteUrl, DOCUMENTS_CANDIDATES);
      if (found) setDocumentsLibraryName(found.title);
    })();
  }
  return documentsLookup;
}

let hcLookup: Promise<void> | undefined;

async function primeHcLibraries(sp: SPHttpClient, siteUrl: string): Promise<void> {
  if (!hcLookup) {
    hcLookup = (async () => {
      const probe = (candidates: string[]): Promise<{ title: string; url: string } | undefined> =>
        probeLibrary(sp, siteUrl, candidates);
      const approval = await probe(HC_APPROVAL_CANDIDATES);
      // Short-circuit: with no approval library there is nothing to pair, and the second probe would
      // spend two requests on every page load of every site that has no HC at all.
      if (!approval) return;
      const documents = await probe(HC_DOCUMENTS_CANDIDATES);
      if (!documents) return;
      setHcLibraryNames(approval.title, approval.url, documents.title, documents.url);
    })();
  }
  return hcLookup;
}

/**
 * The seven-year archive pair, if this site has one.
 *
 * Spec: docs/superpowers/specs/2026-08-22-seven-year-archive-design.md
 *
 * Like the HC probe, absence is a normal outcome: a site with no archive libraries simply never
 * archives, and `archiveAvailable()` false is the whole answer.
 *
 * ⚠ MUST RUN AFTER `primeHcLibraries`, and the ordering is load-bearing. The archive's pairing rule
 * mirrors the site's HC state — an HC site needs BOTH archive libraries, a non-HC site needs only
 * `Archive` — so reading `hcAvailable()` before the HC probe has settled would classify an HC site
 * as non-HC and accept a half-provisioned archive, leaving HC documents silently never archived.
 */
let archiveLookup: Promise<void> | undefined;

async function primeArchiveLibraries(sp: SPHttpClient, siteUrl: string): Promise<void> {
  if (!archiveLookup) {
    archiveLookup = (async () => {
      const normal = await probeLibrary(sp, siteUrl, ARCHIVE_CANDIDATES);
      // Short-circuit: with no normal archive there is nothing to pair, and the second probe would
      // spend requests on every page load of every site that has no archive at all.
      if (!normal) return;
      const needsHc = hcAvailable();
      const hc = needsHc ? await probeLibrary(sp, siteUrl, ARCHIVE_HC_CANDIDATES) : undefined;
      setArchiveLibraryNames(normal.title, normal.url, needsHc, hc?.title, hc?.url);
    })();
  }
  return archiveLookup;
}

export async function primeNames(sp: SPHttpClient, siteUrl: string): Promise<void> {
  await primeSiteEntry(sp, siteUrl);
  await primeLibrary(sp, siteUrl);
  // Independent of the HC and archive probes — no ordering constraint, unlike those two.
  await primeDocumentsLibrary(sp, siteUrl);
  await primeHcLibraries(sp, siteUrl);
  // AFTER the HC probe — see primeArchiveLibraries. Not concurrent with it, deliberately.
  await primeArchiveLibraries(sp, siteUrl);
  const probe = makeListProbe(sp, siteUrl);
  for (const suffix of PRIMED_SUFFIXES) {
    try {
      await resolveListTitle(suffix, probe);
    } catch {
      // Leave it unresolved; cachedListTitle falls back to the legacy title.
    }
  }
}

/** Live title of one of our lists, e.g. listTitle(sp, url, LIST_SUFFIX.groupMap). */
export function listTitle(sp: SPHttpClient, siteUrl: string, suffix: string): Promise<string> {
  return resolveListTitle(suffix, makeListProbe(sp, siteUrl));
}

/** URL-encoded live title, for dropping straight into a getbytitle() call. */
export async function listTitleEncoded(sp: SPHttpClient, siteUrl: string, suffix: string): Promise<string> {
  return encodeURIComponent(await listTitle(sp, siteUrl, suffix));
}

/**
 * The prefix in force on this site, derived from whichever Config list answered.
 *
 * Used for the names that cannot be probed the way a list can — the folder content type, the
 * site-entry group, and the three custom permission levels.
 */
export function writePrefix(sp: SPHttpClient, siteUrl: string): Promise<string> {
  return resolveWritePrefix(makeListProbe(sp, siteUrl));
}

/** `<P> Folder` content type name for this site. */
export async function folderContentType(sp: SPHttpClient, siteUrl: string): Promise<string> {
  return folderContentTypeName(await writePrefix(sp, siteUrl));
}

/** `<P>_SITE_MEMBERS` for this site. */
export async function siteEntryGroup(sp: SPHttpClient, siteUrl: string): Promise<string> {
  return siteEntryGroupName(await writePrefix(sp, siteUrl));
}

/**
 * The three custom permission levels, named for this site.
 *
 * Missed by the original spec and added on discovery: ROLE_TO_PERMISSION hardcoded
 * "DMS Upload" / "DMS Approve" / "DMS Delete", so a client who renames those levels along
 * with everything else gets `no "DMS Approve" role definition on site` for every approver
 * grant. That failure is at least visible in the log — unlike the list reads — but it grants
 * nothing, so an approver simply never gains the level.
 *
 * Note these are levels the CLIENT creates by hand. If they prefix them differently from
 * their lists this resolves wrongly, and the log names the level it could not find — the best
 * available outcome, since nothing can guess what they typed.
 */
export async function permissionLevelNames(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<{ upload: string; approve: string; del: string }> {
  const p = await writePrefix(sp, siteUrl);
  return { upload: `${p} Upload`, approve: `${p} Approve`, del: `${p} Delete` };
}

export { LIST_SUFFIX };
