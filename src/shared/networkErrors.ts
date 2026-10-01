/**
 * Turns a raw browser-level fetch failure into a message that names the actual cause.
 *
 * ── What this exists for (found live, 2026-08-24) ───────────────────────────────────────────────
 * `fetch()` (and SPFx's `SPHttpClient`, which is built on it) throws a bare `"Failed to fetch"`
 * when a request is stopped BEFORE it reaches the server — a browser extension or content blocker
 * intercepting it, not SharePoint refusing it. Confirmed live: an upload's `Files/Add` POST was
 * blocked outright by Brave Shields, and every OTHER request on the same page (folder lookups,
 * config reads, the clash checks) went through the identical connection and worked — only the one
 * carrying an actual file body was stopped, which is exactly the shape of a blocker rule keyed on
 * request size or type, not a server-side or permissions problem.
 *
 * Both upload web parts caught this and showed the raw string verbatim: `"Failed to fetch"`, with
 * nothing to tell an uploader it has nothing to do with their document, their permissions, or the
 * DMS at all. That reads as "the system is broken" — the one class of error this codebase's own
 * rule (name the cause AND the fix) had not yet reached.
 *
 * ── Why this substitutes rather than appending ───────────────────────────────────────────────────
 * The raw message is developer noise, not a second fact worth keeping beside the friendly one — an
 * uploader does not need "Failed to fetch (this usually means…)"; they need the one sentence that
 * tells them what to click. `console.error` at each call site still logs the original error for
 * whoever has to actually debug it.
 *
 * ── Deliberately NARROW matching ─────────────────────────────────────────────────────────────────
 * Only the handful of literal strings browsers themselves use for "this request never reached the
 * network" are matched. A genuine SharePoint error — a 403, a malformed payload, a throttle — never
 * produces these exact phrases, so this can never mask a real, actionable server response behind a
 * generic "check your extensions" message. Case-insensitive substring match, because Firefox's
 * wording carries extra words around the same core phrase.
 */
const BLOCKED_BEFORE_NETWORK_SIGNATURES = [
  "failed to fetch", // Chrome / Edge / Brave / any Chromium browser
  "networkerror when attempting to fetch resource", // Firefox
  "load failed", // Safari
];

/**
 * ⚠ EXPORTED BECAUSE THE XHR PATH CANNOT REACH `friendlyUploadError` (2026-08-27).
 *
 * The 2026-08-24 fix covered `fetch`, which THROWS a recognisable string. Bulk Upload writes the
 * file through XMLHttpRequest instead — the only way to get real upload progress, since
 * `spHttpClient` exposes none — and a blocked XHR does not throw at all: `xhr.onerror` fires and the
 * request resolves with `status === 0`. None of the signatures below ever appear, so that screen went
 * on showing the raw `HTTP 0 — Network error during upload.` for the very failure this module exists
 * to explain. Seen live 2026-08-27, and it had been that way since the module was written: the same
 * "a fix applied to one upload path is not applied to the other" trap, now four times in two days.
 *
 * One constant, two entry points — the wording must not be copied.
 */
export const BLOCKED_UPLOAD_MESSAGE =
  "Blocked by your browser or an extension (Brave Shields, an ad blocker, a privacy extension) " +
  "before it reached SharePoint — not a problem with the document or your permissions. If you use " +
  "Brave, click the shield icon in the address bar and turn Shields off for this site, then try " +
  "again. Otherwise, check for a blocking extension and disable it for this site.";

export function friendlyUploadError(err: unknown, fallback: string): string {
  const raw = err instanceof Error ? err.message : typeof err === "string" ? err : "";
  const lower = raw.trim().toLowerCase();
  const blocked = BLOCKED_BEFORE_NETWORK_SIGNATURES.some((sig) => lower.indexOf(sig) !== -1);
  if (blocked) return BLOCKED_UPLOAD_MESSAGE;
  return raw.length > 0 ? raw : fallback;
}

/**
 * Did this request never reach the network?
 *
 * ⚠ FOR XHR ONLY, where a blocked request resolves rather than throwing. `status === 0` means the
 * browser produced no response at all: an extension or content blocker stopped it, the connection
 * dropped, or it was aborted. **The caller must exclude its own aborts first** — an abort is also
 * status 0 and is the user's own doing, so telling them to check their extensions would be wrong.
 *
 * Deliberately narrow, in the same spirit as the signature list above: a genuine SharePoint refusal
 * always carries a real status (403, 400, 429), so this can never mask an actionable server
 * response. It cannot distinguish a blocker from a dropped connection — nothing client-side can —
 * so the message names the likely cause and the fix without asserting it is the only one.
 */
export function blockedBeforeNetwork(status: number): boolean {
  return status === 0;
}
