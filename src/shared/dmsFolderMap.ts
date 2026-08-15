import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { encodeServerRelativePath } from "./pathEncoding";
import { cachedListTitle, LIST_SUFFIX } from "./naming";
import { primeNames } from "./spNaming";

// Re-exported so existing consumers (FolderManager) can keep importing it from here.
export { encodeServerRelativePath };

/**
 * The rename-proof lookup list. Field internal names have no spaces.
 *
 * LEGACY DEFAULT only — the live title is resolved per call, because the client renames this
 * list to "CRS Folder Map" at import (confirmed on their site 2026-08-05).
 */
export const FOLDER_MAP_LIST = "DMS Folder Map";

/**
 * The live, URL-encoded title of the folder-map list.
 *
 * Note the list's ENTITY TYPE does NOT follow the rename: SharePoint derives
 * SP.Data.DMS_x0020_Folder_x0020_MapListItem from the name the list was created with, and a
 * title change never moves it — verified live 2026-08-05 on a fully CRS-renamed site. So
 * anything writing __metadata must keep the DMS form even there.
 */
const mapList = async (spHttpClient: SPHttpClient, siteUrl: string): Promise<string> => {
  await primeNames(spHttpClient, siteUrl);
  return encodeURIComponent(cachedListTitle(LIST_SUFFIX.folderMap));
};

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
    `${siteUrl}/_api/web/lists/getbytitle('${await mapList(spHttpClient, siteUrl)}')/items` +
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

/** A mapping row including its list item Id, so it can be updated or deleted. */
export interface FolderMapRow extends FolderMapping {
  itemId: number;
}

/**
 * Read every mapping row, with its list item Id.
 *
 * Replaces the old `loadMappedTermGuids`, which returned only the term GUIDs. That
 * was enough to answer "is this term mapped?" but not "is the mapping still CORRECT?"
 * — and reconciliation treated the presence of a row as proof it was correct, so a
 * folder deleted and recreated (new UniqueId) left a permanently broken row that no
 * re-run could repair. Callers now get the stored UniqueId and can verify it.
 */
export async function loadFolderMapRows(
  spHttpClient: SPHttpClient,
  siteUrl: string,
): Promise<FolderMapRow[]> {
  const rows: FolderMapRow[] = [];
  let url: string | null =
    `${siteUrl}/_api/web/lists/getbytitle('${await mapList(spHttpClient, siteUrl)}')/items` +
    `?$select=Id,Title,TermGuid,FolderUniqueId,FolderUrl,Section&$top=5000`;
  while (url) {
    const res: SPHttpClientResponse = await spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`Folder map read failed: HTTP ${res.status}`);
    const data = await res.json();
    (data.value ?? []).forEach(
      (r: {
        Id: number;
        Title?: string;
        TermGuid?: string;
        FolderUniqueId?: string;
        FolderUrl?: string;
        Section?: string;
      }) => {
        rows.push({
          itemId: r.Id,
          title: r.Title ?? "",
          termGuid: r.TermGuid ?? "",
          folderUniqueId: r.FolderUniqueId ?? "",
          folderUrl: r.FolderUrl ?? "",
          section: r.Section ?? "",
        });
      },
    );
    url = data["@odata.nextLink"] ?? null;
  }
  return rows;
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
    `${siteUrl}/_api/web/lists/getbytitle('${await mapList(spHttpClient, siteUrl)}')/items`,
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

/**
 * Probe a folder by its stable UniqueId, distinguishing "definitely gone" from
 * "could not tell" — the by-id twin of `probeFolderByPath`, with the same rules.
 *
 * This is the check that decides whether a mapping row is still trustworthy, so it
 * MUST NOT collapse a throttle or a permission error into "missing": treating an
 * uncertain answer as a dead folder would rewrite a perfectly good row.
 *
 * Note this asks by ID, never by path. A folder that was renamed or moved keeps its
 * UniqueId and still resolves here — which is exactly the point, and why the caller
 * must not "verify" a row by comparing it against whatever sits at the term-label
 * path instead. (Doing that would repoint the row at a freshly created empty folder
 * and abandon the real one, documents and all.)
 */
export async function probeFolderById(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  uniqueId: string,
): Promise<FolderProbe> {
  const url =
    `${siteUrl}/_api/web/GetFolderById(guid'${encodeURIComponent(uniqueId)}')` +
    `?$select=UniqueId,ServerRelativeUrl`;
  let res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
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
  // A deleted folder answers 404 here. SharePoint also answers 404 for a folder sitting
  // in the Recycle Bin — restoring it brings the SAME UniqueId back, which is why a
  // caller must never delete a row on this signal alone without the guards in the
  // folder-map-integrity spec.
  if (res.status === 404) {
    return { folder: null, confirmedMissing: true, status: 404 };
  }
  const detail = await res.text().catch(() => "");
  console.error(
    `Folder probe by id could not determine existence. HTTP ${res.status} for UniqueId:`,
    uniqueId,
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
 * Can the CURRENT user actually file a document into this folder?
 *
 * Spec: docs/superpowers/specs/2026-08-12-provisioned-segment-visibility-design.md §2.1
 *
 * "Does the folder exist?" cannot answer this, because existence is not per-user.
 * Reconciliation creates folders by walking the term tree (keyed on abbreviations) and
 * assigns group ACLs in a SEPARATE pass (keyed on Group Map) — so a folder can exist
 * for months while a newly created group has no access to it at all. That gap is what
 * produced the HTTP 403 in the upload form on 2026-08-11: the form offered a unit whose
 * folder was real and whose ACL had never been granted.
 *
 * So this asks the folder what THIS user may do with it, and tests AddListItems —
 * upload, not read. Read is not enough, and is genuinely reachable without upload: a PIC
 * who is also in their unit's base group holds Read on the unit folder through that
 * group while their `*_UPL` group is still ungranted (spec 2026-08-08 §5.8). Probing for
 * mere visibility would call that ready and hand them the same 403.
 *
 * Never collapses an uncertain answer into "denied" — see `UploadAccess`.
 */
export type UploadAccess = "granted" | "denied" | "missing" | "unknown";

/**
 * Bit position of AddListItems in SharePoint's `PermissionKind` enum, whose members are
 * bit POSITIONS rather than masks: viewListItems = 1, addListItems = 2, editListItems =
 * 3. The mask is therefore `1 << (kind - 1)`, which puts AddListItems at bit index 1.
 */
const ADD_LIST_ITEMS_BIT = 1;

/**
 * Test one bit of the low 32 bits of an EffectiveBasePermissions mask.
 *
 * Deliberately arithmetic, not `&`. JavaScript's bitwise operators coerce to a SIGNED
 * 32-bit int, and Full Control returns `Low = "4294967295"` — so `low & mask` reasons
 * about a negative number. It happens to give the right answer for a single-bit test,
 * but it is the kind of correct-by-accident that breaks the moment someone extends this
 * to a two-bit check. Division cannot misread the sign.
 */
function hasPermissionBit(low: number, bitIndex: number): boolean {
  if (!isFinite(low) || low < 0) return false;
  return Math.floor(low / Math.pow(2, bitIndex)) % 2 === 1;
}

export async function probeFolderUploadAccess(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  uniqueId: string,
): Promise<UploadAccess> {
  const url =
    `${siteUrl}/_api/web/GetFolderById(guid'${encodeURIComponent(uniqueId)}')` +
    `/ListItemAllFields/EffectiveBasePermissions`;
  const get = async (): Promise<SPHttpClientResponse> =>
    spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata" },
    });

  // Fewer retries and a shorter ceiling than probeFolderById: this runs on the upload
  // form's mount path, once per authorised unit. A throttled site must not hold the
  // form on a spinner for half a minute — "unknown" already fails open.
  let res: SPHttpClientResponse = await get();
  for (
    let attempt = 0;
    (res.status === 429 || res.status === 503) && attempt < 3;
    attempt++
  ) {
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(8000, 500 * 2 ** attempt);
    await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
    res = await get();
  }

  if (res.status === 401 || res.status === 403) return "denied";
  // Security trimming answers 404 for some folders the user cannot see, so 404 is NOT
  // proof of deletion here. It does not matter: both mean this user cannot upload, and
  // both carry the same fix. Distinguished only so the console line is truthful.
  if (res.status === 404) return "missing";
  if (!res.ok) {
    console.warn(
      `Upload-access probe was inconclusive (HTTP ${res.status}) for folder ${uniqueId} — treating as reachable.`,
    );
    return "unknown";
  }
  const d = await res.json().catch(() => null);
  if (!d || d.Low === undefined) return "unknown";
  return hasPermissionBit(Number(d.Low), ADD_LIST_ITEMS_BIT)
    ? "granted"
    : "denied";
}

/**
 * The same question, asked by PATH instead of by UniqueId.
 *
 * Exists for the Highly Confidential libraries, whose folders are NOT in the Folder Map: that list
 * records one folder per term, in the normal approval library. The HC tree mirrors it exactly — same
 * abbreviations, same shape — so the HC folder is found by swapping the library segment of a path
 * already resolved, and there is nothing to look up.
 *
 * A MISSING FOLDER IS THE NORMAL "NOT SET UP YET" ANSWER HERE, not an anomaly. Reconciliation builds
 * the HC tree in its own pass; before it runs, a cleared uploader's HC folder simply does not exist.
 * Returning "missing" so the caller hides the level is the whole point — the folder must never be
 * ensure-created by an uploader, because a folder created that way INHERITS the library root's
 * permissions instead of carrying the unit's, which is how a Highly Confidential document ends up
 * readable by exactly the people the unit ACL exists to exclude.
 *
 * THE PATH GOES IN AS AN ODATA PARAMETER ALIAS, never as an inline quoted literal (gotcha #9). An
 * inline path returns HTTP 400 — not 404 — once it is deep enough, and a 400 reads as "malformed
 * request" rather than "no such folder", which is the wrong conclusion to draw here.
 */
export async function probeFolderUploadAccessByPath(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  serverRelativeUrl: string,
): Promise<UploadAccess> {
  const path = (serverRelativeUrl ?? "").trim();
  if (path.length === 0) return "missing";
  const url =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/EffectiveBasePermissions` +
    `?@f='${encodeServerRelativePath(path)}'`;
  const get = async (): Promise<SPHttpClientResponse> =>
    spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata" },
    });

  // The same short retry budget as the UniqueId probe: this runs on the upload form's interaction
  // path, and holding the form on a spinner is worse than an inconclusive answer.
  let res: SPHttpClientResponse = await get();
  for (
    let attempt = 0;
    (res.status === 429 || res.status === 503) && attempt < 3;
    attempt++
  ) {
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(8000, 500 * 2 ** attempt);
    await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
    res = await get();
  }

  if (res.status === 401 || res.status === 403) return "denied";
  // Security trimming answers 404 for folders the user cannot see, so this is not proof of absence.
  // It does not matter: both readings mean the same thing here, and neither may offer the level.
  if (res.status === 404) return "missing";
  if (!res.ok) {
    console.warn(
      `HC upload-access probe was inconclusive (HTTP ${res.status}) for ${path} — the level stays hidden.`,
    );
    return "unknown";
  }
  const d = await res.json().catch(() => null);
  if (!d || d.Low === undefined) return "unknown";
  // Arithmetic, never `&`: JS bitwise coerces to a SIGNED 32-bit int, and Full Control returns
  // Low = "4294967295", which as a signed int is -1.
  return hasPermissionBit(Number(d.Low), ADD_LIST_ITEMS_BIT) ? "granted" : "denied";
}

/** Repoint an existing mapping row at a different folder. */
export async function updateFolderMapping(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  itemId: number,
  patch: { folderUniqueId: string; folderUrl: string; title?: string },
): Promise<void> {
  const body: Record<string, string> = {
    FolderUniqueId: patch.folderUniqueId,
    FolderUrl: patch.folderUrl,
  };
  if (patch.title !== undefined) body.Title = patch.title;
  const res: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('${await mapList(spHttpClient, siteUrl)}')/items(${itemId})`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify(body),
    },
  );
  if (!res.ok) {
    const b = await res.text().catch(() => "");
    throw new Error(`Update mapping failed: HTTP ${res.status} ${b.slice(0, 200)}`);
  }
}

/**
 * Delete one mapping row. Removes the LIST ROW ONLY — never the folder it points at.
 * A map row is derived data that reconciliation can rebuild from the term store plus
 * the folder tree; the folder holds documents that exist nowhere else.
 */
export async function deleteFolderMapRow(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  itemId: number,
): Promise<void> {
  const res: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('${await mapList(spHttpClient, siteUrl)}')/items(${itemId})`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    },
  );
  if (!res.ok) {
    const b = await res.text().catch(() => "");
    throw new Error(`Delete mapping row failed: HTTP ${res.status} ${b.slice(0, 200)}`);
  }
}

/** Outcome of resolving a mapped folder — see `resolveMappedFolder`. */
export interface MappedFolderResolution {
  /** The folder's current path, when it could be resolved at all. */
  serverRelativeUrl?: string;
  /** True ONLY when SharePoint positively said the folder is not there (404). */
  confirmedMissing: boolean;
  /** Last status seen on the by-id lookup, for the message and the console. */
  status: number;
  /** True when the by-id lookup failed but the stored path worked. */
  usedStoredPath: boolean;
}

/**
 * Resolve a mapped folder, distinguishing "deleted" from "this user cannot look it up".
 *
 * `resolveFolderServerUrl` collapses every failure to null, and both upload web parts then
 * told the user *"The mapped unit folder no longer exists — ask an administrator to re-run
 * reconciliation."* That sentence is a CONCLUSION, and a 403, a throttle or a moderation
 * trim produces it just as readily as a deletion. It sends an administrator to re-run
 * reconciliation over a tree that is perfectly intact, and it blocks an uploader whose
 * folder is right there.
 *
 * Same failure shape as `probeFolderByPath` (a bulk upload during reconciliation reporting a
 * folder as absent) and as gotcha #9 — a request that could not be answered is not evidence
 * of missing data.
 *
 * So: try the UniqueId, which is rename-proof and the right primary key. If that returns
 * anything OTHER than a clean 404, fall back to the path stored on the map row. That path is
 * the STAGING path — reconciliation writes the row only on the Staging pass — which is
 * exactly what both callers expect, so the fallback cannot route a file into the wrong
 * library. It can be stale after a rename, which is why it is the fallback and never the
 * primary, and it is verified before being returned.
 */
export async function resolveMappedFolder(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  uniqueId: string,
  storedUrl?: string,
): Promise<MappedFolderResolution> {
  const byId = await probeFolderById(spHttpClient, siteUrl, uniqueId);
  if (byId.folder) {
    return {
      serverRelativeUrl: byId.folder.serverRelativeUrl,
      confirmedMissing: false,
      status: byId.status,
      usedStoredPath: false,
    };
  }
  if (byId.confirmedMissing) {
    return { confirmedMissing: true, status: 404, usedStoredPath: false };
  }
  if (storedUrl) {
    const byPath = await probeFolderByPath(spHttpClient, siteUrl, storedUrl);
    if (byPath.folder) {
      console.warn(
        `Folder ${uniqueId} could not be resolved by id (HTTP ${byId.status}); using the stored path instead.`,
      );
      return {
        serverRelativeUrl: byPath.folder.serverRelativeUrl,
        confirmedMissing: false,
        status: byId.status,
        usedStoredPath: true,
      };
    }
  }
  return { confirmedMissing: false, status: byId.status, usedStoredPath: false };
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
/** Outcome of a folder rename — see `renameFolder`. */
export interface RenameResult {
  ok: boolean;
  /** New server-relative url, on success. */
  serverRelativeUrl?: string;
  /**
   * True when the rename failed because a sibling already holds the target name.
   * Distinct from a generic failure because it needs a human decision rather
   * than a retry: two terms are competing for one folder name.
   */
  conflict: boolean;
  status: number;
  detail?: string;
}

/**
 * Rename a folder in place, keeping its UniqueId, contents and ACL.
 *
 * `MoveTo` within the same parent is SharePoint's rename. The UniqueId survives,
 * which is what lets the Folder Map row stay valid — only its `FolderUrl` needs
 * refreshing afterwards.
 *
 * A name collision is reported, never forced. Overwriting would merge two units'
 * documents behind one ACL — the exact isolation failure `findCollisions` exists
 * to prevent — while silently skipping would leave a folder whose name disagrees
 * with the abbreviation list and nothing saying why. The caller surfaces it so an
 * administrator fixes the abbreviation.
 *
 * Paths go in as OData parameter aliases, never inline literals: CLAUDE.md
 * gotcha #9 — a long encoded path in a quoted literal returns HTTP 400, not 404,
 * and reads like a missing folder.
 */
export async function renameFolder(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  currentServerRelativeUrl: string,
  newName: string,
): Promise<RenameResult> {
  const parent = currentServerRelativeUrl.slice(
    0,
    currentServerRelativeUrl.lastIndexOf("/"),
  );
  return moveFolderTo(spHttpClient, siteUrl, currentServerRelativeUrl, `${parent}/${newName}`);
}

/**
 * Move a FILE, optionally renaming it in the same call.
 *
 * `flags=0` means "fail if something is already there". Deliberately not `1` (overwrite): a name
 * clash during subtree migration means two documents want one path, and overwriting reports
 * success while destroying one of them. Collision detection is the primary guard (spec §4.1);
 * this flag is the second one beneath it, for the case detection missed.
 *
 * The destination carries the file NAME, so a resolved collision is applied as part of the move
 * rather than as a separate rename — there is then no moment at which two files compete for one
 * path.
 *
 * The item survives the move, so approval status, version history and Created By travel with it.
 * That is what makes migrating an approval library safe; a copy-then-delete resets all three,
 * which is precisely why the Auto-route flow has to restamp them.
 */
export async function moveFileTo(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  currentServerRelativeUrl: string,
  targetServerRelativeUrl: string,
): Promise<RenameResult> {
  const res: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/MoveTo(newUrl=@d,flags=0)` +
      `?@f='${encodeServerRelativePath(currentServerRelativeUrl)}'` +
      `&@d='${encodeServerRelativePath(targetServerRelativeUrl)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (res.ok) {
    return {
      ok: true,
      serverRelativeUrl: targetServerRelativeUrl,
      conflict: false,
      status: res.status,
    };
  }
  const detail = await res.text().catch(() => "");
  return {
    ok: false,
    conflict: /already exists|destination file/i.test(detail),
    status: res.status,
    detail: detail.slice(0, 300),
  };
}

/** Why a folder was not deleted. `deleted: false` with no reason means it was already gone. */
export interface DeleteFolderResult {
  deleted: boolean;
  /** Present when the folder still holds something — the caller reports it rather than forcing. */
  reason?: string;
  status: number;
}

/**
 * Delete a folder ONLY if it holds neither files nor folders.
 *
 * Subtree migration empties source folders as it moves files out, and an empty husk left in the old
 * shape would be reported as drift by every future scan. But the check happens HERE, immediately
 * before the delete, rather than being inferred from what the run moved: a file uploaded during the
 * run lands in a folder the plan believes it emptied, and deleting that folder would destroy a
 * document that was never part of the migration.
 *
 * Anything still inside means the folder is REPORTED, never forced. Deleting a folder is the one
 * operation in this migration with no cheap undo — and DELETE sends it to the recycle bin, which is
 * the only reason this is acceptable at all.
 */
export async function deleteFolderIfEmpty(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  serverRelativeUrl: string,
): Promise<DeleteFolderResult> {
  const check: SPHttpClientResponse = await spHttpClient.get(
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?$select=Folders,Files&$expand=Folders,Files` +
      `&@f='${encodeServerRelativePath(serverRelativeUrl)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (check.status === 404) return { deleted: false, status: 404 };
  if (!check.ok) {
    return { deleted: false, reason: `could not be read (HTTP ${check.status})`, status: check.status };
  }
  const data = (await check.json()) as { Folders?: unknown[]; Files?: unknown[] };
  const folders = (data.Folders ?? []).length;
  const files = (data.Files ?? []).length;
  if (folders > 0 || files > 0) {
    return {
      deleted: false,
      reason: `still holds ${files} file(s) and ${folders} folder(s)`,
      status: check.status,
    };
  }
  const del: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeServerRelativePath(serverRelativeUrl)}'`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    },
  );
  if (!del.ok) {
    const detail = (await del.text().catch(() => "")).slice(0, 200);
    return { deleted: false, reason: `HTTP ${del.status} ${detail}`, status: del.status };
  }
  return { deleted: true, status: del.status };
}

/**
 * Move a folder to a different parent, keeping its UniqueId, contents and ACL.
 *
 * The same `MoveTo` call as a rename — a rename IS a move whose destination shares the
 * source's parent — which is why `renameFolder` delegates here instead of the two keeping
 * separate copies of the request.
 *
 * Used by subtree migration (spec `2026-08-11-subtree-migration-design.md`) to relocate
 * below-Unit folders after a structure change. Three properties make that safe, and all
 * three belong to this call rather than to the caller:
 *   - the whole subtree travels with the folder, in one request;
 *   - approval status survives (verified live 2026-08-10), so nothing needs re-approving;
 *   - UniqueId survives, so Folder Map rows keyed on it stay valid.
 *
 * A name collision at the destination is REPORTED, never forced. Merging two folders would
 * merge two sets of documents behind one ACL, and afterwards nothing records which files
 * came from where.
 */
export async function moveFolderTo(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  currentServerRelativeUrl: string,
  targetServerRelativeUrl: string,
): Promise<RenameResult> {
  const target = targetServerRelativeUrl;
  const res: SPHttpClientResponse = await spHttpClient.post(
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/MoveTo(newUrl=@d)` +
      `?@f='${encodeServerRelativePath(currentServerRelativeUrl)}'` +
      `&@d='${encodeServerRelativePath(target)}'`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (res.ok) {
    return { ok: true, serverRelativeUrl: target, conflict: false, status: res.status };
  }
  const detail = await res.text().catch(() => "");
  return {
    ok: false,
    // SharePoint reports a collision as a 400 with the reason only in the body,
    // so the status alone cannot tell it from a malformed request.
    conflict: /already exists/i.test(detail),
    status: res.status,
    detail: detail.slice(0, 300),
  };
}

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
