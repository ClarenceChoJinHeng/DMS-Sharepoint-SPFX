// Pure path-encoding helper. Kept SPFx-free so jest can test it directly (importing
// dmsFolderMap.ts would pull in @microsoft/sp-http, which the test runner can't resolve).

/**
 * Encode a server-relative path for use inside an OData `@alias='...'` literal WITHOUT
 * turning the slashes into `%2F`. Encoding the whole path with `encodeURIComponent`
 * produces a "%2F flood" that SharePoint rejects with HTTP 400 once the path is deep
 * (~6+ levels — e.g. the unit/Year/DocType grid). Encoding each segment while keeping
 * the literal "/" avoids that, and still escapes spaces, unicode (e.g. fullwidth ＆),
 * "&", "#", etc. per segment.
 */
export function encodeServerRelativePath(path: string): string {
  return path.split("/").map(encodeURIComponent).join("/");
}
