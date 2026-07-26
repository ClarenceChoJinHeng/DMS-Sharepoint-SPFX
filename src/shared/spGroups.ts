// Native SharePoint site-group access — the ONLY module that talks to
// /_api/web/sitegroups. Replaces Graph group search/membership everywhere
// (see docs/superpowers/specs/2026-07-23-native-sharepoint-groups-design.md).
// Site-collection scoped by nature: these groups exist only on this site.
// Pure helpers/types live in spGroupsFilter.ts (Jest-testable, no @microsoft
// import) and are re-exported here so consumers import everything from one place.
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  SpGroup,
  SpGroupMember,
  PersonPick,
  DUPLICATE_GROUP,
  filterDmsGroups,
} from "./spGroupsFilter";

export { SpGroup, SpGroupMember, PersonPick, DUPLICATE_GROUP, filterDmsGroups } from "./spGroupsFilter";

const GET_HEADERS = { Accept: "application/json;odata=nometadata" };
const POST_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
};

const fail = async (label: string, res: SPHttpClientResponse): Promise<never> => {
  const body = await res.text().catch(() => "");
  throw new Error(`${label} HTTP ${res.status} ${body.slice(0, 200)}`);
};

export async function fetchAllSiteGroups(sp: SPHttpClient, siteUrl: string): Promise<SpGroup[]> {
  const res = await sp.get(
    `${siteUrl}/_api/web/sitegroups?$select=Id,Title&$top=500`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (!res.ok) return fail("sitegroups", res);
  const data = await res.json();
  return ((data.value ?? []) as Array<{ Id: number; Title: string }>).map((g) => ({
    id: g.Id,
    title: g.Title,
  }));
}

export async function searchSiteGroups(
  sp: SPHttpClient,
  siteUrl: string,
  q: string,
): Promise<SpGroup[]> {
  return filterDmsGroups(await fetchAllSiteGroups(sp, siteUrl), q);
}

/**
 * Create a site group with NO permissions (Reconciliation grants folder access
 * later, so creating a group is a safe, reversible act). Throws Error(DUPLICATE_GROUP)
 * if the title already exists in the site collection.
 */
export async function createSiteGroup(
  sp: SPHttpClient,
  siteUrl: string,
  title: string,
): Promise<SpGroup> {
  const res = await sp.post(`${siteUrl}/_api/web/sitegroups`, SPHttpClient.configurations.v1, {
    headers: POST_HEADERS,
    body: JSON.stringify({ Title: title.trim() }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    if (body.toLowerCase().indexOf("already in use") !== -1) throw new Error(DUPLICATE_GROUP);
    throw new Error(`create group HTTP ${res.status} ${body.slice(0, 200)}`);
  }
  const data = await res.json();
  return { id: data.Id as number, title: data.Title as string };
}

export async function getGroupMembers(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
): Promise<SpGroupMember[]> {
  const res = await sp.get(
    `${siteUrl}/_api/web/sitegroups(${groupId})/users?$select=Id,Title,Email,LoginName`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (!res.ok) return fail("group members", res);
  const data = await res.json();
  return ((data.value ?? []) as Array<{ Id: number; Title: string; Email?: string; LoginName: string }>).map(
    (u) => ({ id: u.Id, title: u.Title, email: u.Email ?? "", loginName: u.LoginName }),
  );
}

/** Add by login name (claim). SharePoint ensures the user on the site automatically; guests work. */
export async function addGroupMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  loginName: string,
): Promise<void> {
  const res = await sp.post(
    `${siteUrl}/_api/web/sitegroups(${groupId})/users`,
    SPHttpClient.configurations.v1,
    { headers: POST_HEADERS, body: JSON.stringify({ LoginName: loginName }) },
  );
  if (!res.ok) return fail("add member", res);
}

export async function removeGroupMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  userId: number,
): Promise<void> {
  const res = await sp.post(
    `${siteUrl}/_api/web/sitegroups(${groupId})/users/removebyid(${userId})`,
    SPHttpClient.configurations.v1,
    { headers: POST_HEADERS },
  );
  if (!res.ok) return fail("remove member", res);
}

/**
 * Delete a site group entirely. DESTRUCTIVE: removes the group and every role
 * assignment it holds across the site. The member users themselves are not
 * deleted — only the group and its grants. Caller must confirm with the user first.
 */
export async function deleteSiteGroup(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
): Promise<void> {
  const res = await sp.post(
    `${siteUrl}/_api/web/sitegroups/removebyid(${groupId})`,
    SPHttpClient.configurations.v1,
    { headers: POST_HEADERS },
  );
  if (!res.ok) return fail("delete group", res);
}

/** Tenant-wide people search via SharePoint's own picker service — no Graph. */
export async function searchTenantPeople(
  sp: SPHttpClient,
  siteUrl: string,
  q: string,
): Promise<PersonPick[]> {
  const res = await sp.post(
    `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`,
    SPHttpClient.configurations.v1,
    {
      headers: POST_HEADERS,
      body: JSON.stringify({
        queryParams: {
          QueryString: q,
          MaximumEntitySuggestions: 10,
          AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false,
          PrincipalType: 1, // users only
          PrincipalSource: 15, // all sources
        },
      }),
    },
  );
  if (!res.ok) return fail("people search", res);
  const data = await res.json();
  // The endpoint returns its result as a JSON *string*.
  const raw = (data.value ?? data.ClientPeoplePickerSearchUser ?? "[]") as string;
  const entries = JSON.parse(raw) as Array<{
    Key: string;
    DisplayText: string;
    EntityData?: { Email?: string };
  }>;
  return entries.map((e) => ({
    loginName: e.Key,
    displayName: e.DisplayText,
    email: e.EntityData?.Email ?? "",
  }));
}
