import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

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

/** Resolve a server-relative folder path to its stable UniqueId. Returns null if the folder does not exist. */
export async function resolveFolderByPath(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  serverRelativePath: string,
): Promise<{ uniqueId: string; serverRelativeUrl: string } | null> {
  const url =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeURIComponent(serverRelativePath)}'` +
    `&$select=UniqueId,ServerRelativeUrl`;
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!res.ok) return null;
  const d = await res.json();
  return { uniqueId: d.UniqueId, serverRelativeUrl: d.ServerRelativeUrl };
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
