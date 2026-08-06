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
  LIBRARY_CANDIDATES,
  setLibraryNames,
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
        const res: SPHttpClientResponse = await sp.get(
          `${siteUrl}/_api/web/sitegroups?$select=Title&$top=500`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!res.ok) return;
        const data = await res.json();
        setSiteEntryName(((data.value ?? []) as Array<{ Title?: string }>).map((g) => g.Title ?? ""));
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

export async function primeNames(sp: SPHttpClient, siteUrl: string): Promise<void> {
  await primeSiteEntry(sp, siteUrl);
  await primeLibrary(sp, siteUrl);
  const probe = makeListProbe(sp, siteUrl);
  for (const suffix of [
    LIST_SUFFIX.config,
    LIST_SUFFIX.groupMap,
    LIST_SUFFIX.folderMap,
    LIST_SUFFIX.abbreviation,
    LIST_SUFFIX.deletionLog,
  ]) {
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
