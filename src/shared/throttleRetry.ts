import { SPHttpClientResponse } from "@microsoft/sp-http";

/**
 * Retry a SharePoint request that was THROTTLED or hit a momentarily unavailable service.
 *
 * ⚠ 429 AND 503 ARE THE ONLY STATUSES RETRIED, and that list is the whole judgement in this file.
 * Every other failure means something specific and permanent, and asking again would ask the same
 * unanswerable question twice while hiding the real status: 400 is a malformed request, 403 is
 * permissions, 404 is the thing not being there. **503 is "Service Unavailable" — transient by
 * definition** — and SharePoint expects the caller to come back.
 *
 * `Retry-After` is honoured whenever SharePoint sends it, because it knows how long it wants; the
 * exponential fallback covers the case where it does not, and is capped at 30s so a page cannot
 * appear to hang indefinitely.
 *
 * The request is passed as a THUNK rather than as a URL, so the retry cannot drift from the original
 * call — a second copy of the request would be a second thing to keep correct.
 *
 * ⚠ IT RETURNS THE LAST RESPONSE AND NEVER THROWS. A caller that has exhausted the attempts still
 * gets a real `SPHttpClientResponse` to read a status off, so existing error handling is unchanged:
 * this only removes the case where a single blip was treated as a permanent failure.
 *
 * **Why it lives here.** It was a private const in `spGroups.ts` for group reads, and CRS Search had
 * none at all — so one 503 from `_api/search/query` put a red *"One library could not be searched"*
 * banner in front of a user, over a status SharePoint had asked us to retry (reported 2026-09-07).
 * The Search REST surface is throttled harder than ordinary list reads and is called on every press
 * of Search, so it is the most exposed caller in the project.
 *
 * ⚠ SEVEN COPIES OF THIS LOGIC EXIST AND ONLY TWO CALLERS USE THIS ONE. `dmsFolderMap.ts` holds SIX
 * inline retry loops and `FolderManager.tsx` a seventh, each written where it was needed. They are
 * deliberately NOT converted here: those two files carry the folder-creation and permission-granting
 * code that is the most site-verified in the project, several of their loops do more than a plain
 * retry, and rewriting them to save duplication is the wrong risk to take inside a change about a
 * search banner. **Worth doing as its own reviewable change** — and until it is, a change to the
 * retry rule has to be made in three places.
 */
export const withThrottleRetry = async (
  send: () => Promise<SPHttpClientResponse>,
): Promise<SPHttpClientResponse> => {
  let res = await send();
  for (
    let attempt = 0;
    (res.status === 429 || res.status === 503) && attempt < 5;
    attempt++
  ) {
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(30000, 1000 * 2 ** attempt);
    await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
    res = await send();
  }
  return res;
};
