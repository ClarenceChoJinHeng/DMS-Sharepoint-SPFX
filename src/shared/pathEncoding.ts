// Pure path-encoding helper. Kept SPFx-free so jest can test it directly (importing
// dmsFolderMap.ts would pull in @microsoft/sp-http, which the test runner can't resolve).

/**
 * Encode a server-relative path for use inside an OData `@alias='...'` literal WITHOUT
 * turning the slashes into `%2F`. Encoding the whole path with `encodeURIComponent`
 * produces a "%2F flood" that SharePoint rejects with HTTP 400 once the path is deep
 * (~6+ levels — e.g. the unit/Year/DocType grid). Encoding each segment while keeping
 * the literal "/" avoids that, and still escapes spaces, unicode (e.g. fullwidth ＆),
 * "&", "#", etc. per segment.
 *
 * Single quotes are then DOUBLED, which is how OData escapes a quote inside a string
 * literal. `encodeURIComponent` leaves "'" untouched (it is an unreserved character),
 * so without this a folder like "President's Office" closes the literal early and
 * SharePoint answers HTTP 400 — `The query string "DecodedUrl" is missing or invalid`.
 * Percent-encoding it as %27 does NOT help: the server decodes before the OData parser
 * runs, so the bare quote comes back. Doubling is the only correct fix.
 *
 * Real folders that hit this: "President's Office" (Group Head Office) and
 * "Chief Operating Officer's Office" (NBPOL Head Office).
 */
export function encodeServerRelativePath(path: string): string {
  return path.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''");
}
