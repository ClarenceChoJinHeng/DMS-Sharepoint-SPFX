import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { encodeServerRelativePath } from "./pathEncoding";

// Re-exported so existing consumers (FolderManager) can keep importing it from here.
export { encodeServerRelativePath };

/** The rename-proof lookup list. Field internal names have no spaces. */
export const FOLDER_MAP_LIST = "DMS Folder Map";

export interface FolderMapping {
  termGuid: string;
  folderUniqueId: string;
  title: string;
  folderUrl: string;
  section: string;
}

/** Read the mapping row for a single term. Returns null if the term is not mapped. */
export async function lookupFolderMapping(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  termGuid: string,
): Promise<FolderMapping | null> {
  const url =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FOLDER_MAP_LIST)}')/items` +
    `?$select=Title,TermGuid,FolderUniqueId,FolderUrl,Section&$filter=TermGuid eq '${termGuid}'&$top=1`;
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) throw new Error(`Folder map lookup failed: HTTP ${res.status}`);
  const data = await res.json();
  const row = (data.value ?? [])[0];
  if (!row) return null;
  return {
    termGuid: row.TermGuid,
    folderUniqueId: row.FolderUniqueId,
    title: row.Title,
    folderUrl: row.FolderUrl,
    section: row.Section,
  };
}

/** Read every mapping row's TermGuid — used to skip already-mapped terms. */
export async function loadMappedTermGuids(
  spHttpClient: SPHttpClient,
  siteUrl: string,
): Promise<Set<string>> {
  const guids = new Set<string>();
  let url: string | null =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FOLDER_MAP_LIST)}')/items` +
    `?$select=TermGuid&$top=5000`;
  while (url) {
    const res: SPHttpClientResponse = await spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`Folder map read failed: HTTP ${res.status}`);
    const data = await res.json();
    (data.value ?? []).forEach((r: { TermGuid?: string }) => {
      if (r.TermGuid) guids.add(r.TermGuid);
    });
    url = data["@odata.nextLink"] ?? null;
  }
  return guids;
}

/** Outcome of a folder existence probe — see `probeFolderByPath`. */
export interface FolderProbe {
  folder: { uniqueId: string; serverRelativeUrl: string } | null;
  /**
   * True ONLY when SharePoint positively reported the folder is not there (404).
   * A throttle, a permission error or a malformed request all leave this false —
   * callers must not tell the user "this folder does not exist" unless it is true.
   */
  confirmedMissing: boolean;
  /** Last HTTP status seen, for diagnostics. */
  status: number;
  /** Response body snippet on a non-404 failure, for diagnostics. */
  detail?: string;
}

/**
 * Probe a server-relative folder path, distinguishing "definitely not there" from
 * "could not tell". The old collapse-everything-to-null behaviour caused a real
 * misdiagnosis: a bulk upload run DURING reconciliation reported "the folder does
 * not exist in Documents — ask an administrator to create it" for a folder that
 * existed. Reconciliation is exactly the process that provokes this — it creates
 * folders, breaks inheritance and rewrites role assignments on the very folder
 * being probed, and it generates enough traffic to get throttled. So:
 *   - 404          → confirmedMissing (a real, actionable "not there")
 *   - 429/503      → retried with Retry-After, same as ensureFolder's POST
 *   - 403/400/5xx  → NOT missing; reported with status + body per CLAUDE.md gotcha #9
 */
export async function probeFolderByPath(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  serverRelativePath: string,
): Promise<FolderProbe> {
  const url =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeServerRelativePath(serverRelativePath)}'` +
    `&$select=UniqueId,ServerRelativeUrl`;
  let res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  // A throttle is not an answer about existence — wait it out before deciding.
  for (
    let attempt = 0;
    (res.status === 429 || res.status === 503) && attempt < 5;
    attempt++
  ) {
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(30000, 1000 * 2 ** attempt);
    await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
    res = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata" },
    });
  }
  if (res.ok) {
    const d = await res.json();
    return {
      folder: { uniqueId: d.UniqueId, serverRelativeUrl: d.ServerRelativeUrl },
      confirmedMissing: false,
      status: res.status,
    };
  }
  if (res.status === 404) {
    return { folder: null, confirmedMissing: true, status: 404 };
  }
  // Anything else: log loudly. A 400 means a malformed request, not missing data.
  const detail = await res.text().catch(() => "");
  console.error(
    `Folder probe could not determine existence. HTTP ${res.status} for path:`,
    serverRelativePath,
    detail.slice(0, 300),
  );
  return {
    folder: null,
    confirmedMissing: false,
    status: res.status,
    detail: detail.slice(0, 300),
  };
}

/**
 * Resolve a server-relative folder path to its stable UniqueId. Returns null if the
 * folder does not exist — or if existence could not be determined. Prefer
 * `probeFolderByPath` when the difference matters to the user.
 */
export async function resolveFolderByPath(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  serverRelativePath: string,
): Promise<{ uniqueId: string; serverRelativeUrl: string } | null> {
  const probe = await probeFolderByPath(
    spHttpClient,
    siteUrl,
    serverRelativePath,
  );
  return probe.folder;
}

/** Create one mapping row. */
export async function writeFolderMapping(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  m: FolderMapping,
): Promise<void> {
  const res: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(FOLDER_MAP_LIST)}')/items`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
      },
      body: JSON.stringify({
        Title: m.title,
        TermGuid: m.termGuid,
        FolderUniqueId: m.folderUniqueId,
        FolderUrl: m.folderUrl,
        Section: m.section,
      }),
    },
  );
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Write mapping failed: HTTP ${res.status} ${body.slice(0, 200)}`);
  }
}

/** Get a folder's CURRENT server-relative URL from its stable UniqueId. */
export async function resolveFolderServerUrl(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  uniqueId: string,
): Promise<string | null> {
  const res: SPHttpClientResponse = await spHttpClient.get(
    `${siteUrl}/_api/web/GetFolderById(guid'${uniqueId}')?$select=ServerRelativeUrl`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) return null;
  const d = await res.json();
  return d.ServerRelativeUrl ?? null;
}

/**
 * Ensure a subfolder `name` exists directly under `parentServerRelativeUrl`.
 * Idempotent: creates it, or resolves the existing one on 409/exists.
 * Returns the child's UniqueId + ServerRelativeUrl, or null on failure.
 */
export async function ensureFolder(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  parentServerRelativeUrl: string,
  name: string,
): Promise<{
  uniqueId: string;
  serverRelativeUrl: string;
  /**
   * True only when THIS call created the folder. Callers that pace their writes
   * (reconciliation's throttle) must not charge a delay for a folder that already
   * existed — that is what made a no-op re-run take as long as a real one.
   */
  created: boolean;
} | null> {
  const childPath = `${parentServerRelativeUrl}/${name}`;
  // Check first: on re-runs the folder usually already exists, so skip the create.
  // This is one GET instead of a POST that 400s ("already exists") followed by a
  // resolve GET — faster on re-runs and no noisy 400s in the console.
  const found = await resolveFolderByPath(spHttpClient, siteUrl, childPath);
  if (found) return { ...found, created: false };
  // Retry on SharePoint throttling (429/503), honoring Retry-After — the reconciliation
  // grid can be large, so a transient throttle here must not silently drop a folder.
  let addRes: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl=@u)?@u='${encodeServerRelativePath(childPath)}'&$select=UniqueId,ServerRelativeUrl`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  for (let attempt = 0; (addRes.status === 429 || addRes.status === 503) && attempt < 5; attempt++) {
    const ra = Number(addRes.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(30000, 1000 * 2 ** attempt);
    await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
    addRes = await spHttpClient.post(
      `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl=@u)?@u='${encodeServerRelativePath(childPath)}'&$select=UniqueId,ServerRelativeUrl`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
  }
  if (addRes.ok) {
    const d = await addRes.json();
    return {
      uniqueId: d.UniqueId,
      serverRelativeUrl: d.ServerRelativeUrl,
      created: true,
    };
  }
  // Lost a race, or a transient error — resolve it by path once more. Someone else
  // created it, so this call did not: created stays false.
  const existing = await resolveFolderByPath(spHttpClient, siteUrl, childPath);
  return existing ? { ...existing, created: false } : null;
}
