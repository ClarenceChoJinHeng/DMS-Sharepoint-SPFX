// Pure logic for Share Guard: parse the item context passed from the command-set
// button, map access level to a SharePoint role name, and build the request list
// item body. No SPHttpClient here — kept pure for jest (mirrors siteMap.ts).

export type AccessLevel = "Read" | "Edit";
export type ItemType = "File" | "Folder";
export type LibraryName = "Staging" | "Documents";

export interface ShareTarget {
  itemUrl: string; // server-relative path of the file/folder
  itemUniqueId: string; // rename-proof handle used for the grant
  itemType: ItemType;
  library: LibraryName;
}

/** Access level -> SharePoint role definition name (looked up by name at runtime). */
export function roleNameForAccess(level: AccessLevel): string {
  return level === "Edit" ? "Edit" : "Read";
}

/** Parse the item context the command-set button puts on the query string. */
export function parseShareTarget(q: URLSearchParams): ShareTarget | null {
  const itemUrl = q.get("itemUrl") ?? "";
  const itemUniqueId = q.get("itemId") ?? "";
  if (!itemUrl || !itemUniqueId) return null;
  const itemType: ItemType = q.get("itemType") === "Folder" ? "Folder" : "File";
  const library: LibraryName = q.get("library") === "Staging" ? "Staging" : "Documents";
  return { itemUrl, itemUniqueId, itemType, library };
}

/** Last path segment (no trailing slash) — the display name / Title. */
export function itemLeafName(serverRelativeUrl: string): string {
  const trimmed = serverRelativeUrl.replace(/\/+$/, "");
  const idx = trimmed.lastIndexOf("/");
  return idx >= 0 ? trimmed.slice(idx + 1) : trimmed;
}

/** Build the DMS Share Requests list item body. RecipientId = ensured SP user id.
 *  NOTE: field internal names below are the DEFAULT-cased names; if the live list
 *  reveals different frozen internal names, they will be updated later. */
export function buildRequestPayload(
  target: ShareTarget,
  recipientId: number,
  reason: string,
  accessLevel: AccessLevel,
): Record<string, unknown> {
  return {
    Title: itemLeafName(target.itemUrl),
    ItemUrl: target.itemUrl,
    ItemUniqueId: target.itemUniqueId,
    ItemType: target.itemType,
    Library: target.library,
    RecipientId: recipientId,
    AccessLevel: accessLevel,
    Reason: reason,
    Status: "Pending",
  };
}

/** Absolute click-through URL for an item, origin + encoded server-relative path. */
export function itemAbsoluteUrl(origin: string, serverRelativeUrl: string): string {
  return `${origin}${encodeURI(serverRelativeUrl)}`;
}

/** Location breadcrumb: the decoded path segments of an item below the site's web
 *  root. e.g. ("/sites/X/Staging/GHO/Treasury", "/sites/X") -> ["Staging","GHO","Treasury"].
 *  Lets the approver see WHERE an item lives without opening it. Matches the web
 *  root case-insensitively; if it isn't a prefix, returns every segment. */
export function itemBreadcrumb(serverRelativeUrl: string, webServerRelativeUrl: string): string[] {
  const web = webServerRelativeUrl.replace(/\/+$/, "");
  let rest = serverRelativeUrl;
  if (web && rest.toLowerCase().indexOf(web.toLowerCase()) === 0) {
    rest = rest.slice(web.length);
  }
  return rest
    .split("/")
    .filter(Boolean)
    .map((s) => {
      try { return decodeURIComponent(s); } catch { return s; }
    });
}
