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
  filterSelectableGroups,
} from "./spGroupsFilter";

export { SpGroup, SpGroupMember, PersonPick, DUPLICATE_GROUP, filterSelectableGroups } from "./spGroupsFilter";

const GET_HEADERS = { Accept: "application/json;odata=nometadata" };
const POST_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
};

/**
 * A readable failure.
 *
 * SharePoint answers a throttle with an HTML page, not JSON — so the raw body put a wall of
 * `<!DOCTYPE html><html xml:lang="en"…` in front of the user where a sentence belongs, and said
 * nothing about what to do about it. Seen live 2026-08-16 removing one account from eight groups.
 *
 * HTML is therefore never shown: it is never the explanation, only the transport's error page.
 */
const fail = async (label: string, res: SPHttpClientResponse): Promise<never> => {
  if (res.status === 429 || res.status === 503) {
    throw new Error(
      `${label}: SharePoint is busy and refused the request (HTTP ${res.status}). ` +
        "It was retried and stayed busy. Wait a minute and try again — nothing was changed.",
    );
  }
  const body = await res.text().catch(() => "");
  const looksLikeHtml = /^\s*<(?:!doctype|html)/i.test(body);
  const detail = looksLikeHtml
    ? "(SharePoint returned an error page rather than a message)"
    : body.slice(0, 200);
  throw new Error(`${label} HTTP ${res.status} ${detail}`);
};

/**
 * Run a write, retrying while SharePoint says it is busy.
 *
 * ── Why this exists ──────────────────────────────────────────────────────────
 * `dmsFolderMap.ts` has retried 429/503 in six places since reconciliation was built, because that
 * code makes hundreds of writes and a dropped one silently loses a folder. Group operations were
 * the ONE bulk-write path without it — and creating groups and adding members in bursts is exactly
 * what migrating to a new tenant does. It failed live on the eighth removal of eight.
 *
 * `Retry-After` is honoured whenever SharePoint sends it, because it knows how long it wants; the
 * exponential fallback covers the case where it does not, and is capped so a page cannot appear to
 * hang indefinitely.
 *
 * The request is passed as a THUNK rather than as a URL, so the retry cannot drift from the
 * original call — a second copy of the request would be a second thing to keep correct.
 */
const withThrottleRetry = async (
  send: () => Promise<SPHttpClientResponse>,
): Promise<SPHttpClientResponse> => {
  let res = await send();
  for (let attempt = 0; (res.status === 429 || res.status === 503) && attempt < 5; attempt++) {
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(30000, 1000 * 2 ** attempt);
    await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
    res = await send();
  }
  return res;
};

export async function fetchAllSiteGroups(sp: SPHttpClient, siteUrl: string): Promise<SpGroup[]> {
  const res = await sp.get(
    /* ⚠ 5000, NOT 500 — the cap was reached on a live site (2026-08-21) and the failure is SILENT.
     ClarenceDMSTesting holds 585 groups, so a 500-cap read simply omitted the rest: Group Management
     reported `MHO_SEGVIEW` as not existing while it sat at id 396, and the admin was one step from
     creating a duplicate. Every access screen reads through here — Group Management, Folder Access,
     Site Access, Page Access — so a truncated read makes live groups invisible on all four at once,
     with no error anywhere. Reconciliation's own group check already avoided this reader for exactly
     this reason (FolderManager.tsx ~3352). A provisioned segment is ~324 groups on its own, so 500 is
     under two segments. */
    `${siteUrl}/_api/web/sitegroups?$select=Id,Title&$top=5000`,
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

/**
 * Every group's members, in ONE request.
 *
 * Client, 2026-08-23, on the group list and its CSV: *"the Excel doesnt show who is inside each
 * group"* and *"instead of showing mappings only, show how many users are there in each group,
 * client doesnt understand what is mapping"*. Both need the same data for all 586 groups at once,
 * and `getGroupMembers` is per group — 586 requests, which is not an export, it is an outage.
 *
 * `$expand=Users` answers it in one. The payload is large but bounded by the site's real membership,
 * and this is read once per page load rather than per keystroke.
 *
 * ⚠ RETURNS `undefined` ON FAILURE, never `{}`. A caller must be able to tell "this group has
 * nobody in it" from "we could not find out" — the first is a fact worth showing in amber, the
 * second must not be rendered as `0 people`, which is the same trap as the capped `$top` read above.
 *
 * `$top=5000` for the same reason as `fetchAllSiteGroups`: 500 was reached live on 2026-08-21 and a
 * truncated read reports live groups as having no members.
 */
export async function fetchAllGroupMembers(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<Record<number, SpGroupMember[]> | undefined> {
  const out: Record<number, SpGroupMember[]> = {};
  /* ⚠ PAGED, and it was not before — which is the likeliest reason the CSV export came back with
     "not known" in every People and Members cell on 2026-08-30 while the group list on the same
     screen showed the members perfectly.

     `$top=5000` caps ONE page; it does not promise one page. Expanding `Users` across ~700 groups
     produces a large response, and SharePoint answers a partial set with an `odata.nextLink`
     rather than an error — or refuses the whole thing. Following the link is the only way to be
     sure the index describes the whole site, and an INCOMPLETE index here is worse than none: the
     export would state a group has nobody in it, and a spreadsheet saying a unit's approver group
     is empty gets acted on. Same family as memory `sp-capped-read-reads-as-absent`.

     Bounded at 20 pages so a paging bug cannot spin for ever; hitting the cap returns `undefined`
     rather than a short answer, for the reason above. */
  let url: string | undefined =
    `${siteUrl}/_api/web/sitegroups?$select=Id,Users/Id,Users/Title,Users/Email,Users/LoginName`
    + `&$expand=Users&$top=500`;
  let pages = 0;
  try {
    while (url && pages < 20) {
      pages += 1;
      const next: string = url;
      const res = await withThrottleRetry(() =>
        sp.get(next, SPHttpClient.configurations.v1, { headers: GET_HEADERS }),
      );
      if (!res.ok) {
        /* ⚠ SAID OUT LOUD. This used to fail silently, so the only symptom was a column reading
           "not known" with nothing anywhere explaining it — and the two causes need opposite
           fixes: a 403 is permissions on the site groups, a 429/503 is throttling that a retry
           fixes, a 500 is the expand itself being too large. */
        console.error(
          `[CRS] fetchAllGroupMembers: HTTP ${res.status} on page ${pages}. `
          + "Member counts and the CSV's People/Members columns will read \"not known\".",
        );
        return undefined;
      }
      const data = await res.json();
      for (const g of (data.value ?? []) as Array<{
        Id: number;
        Users?: Array<{ Id: number; Title: string; Email?: string; LoginName: string }>;
      }>) {
        out[g.Id] = (g.Users ?? []).map((u) => ({
          id: u.Id,
          title: u.Title,
          email: u.Email ?? "",
          loginName: u.LoginName,
        }));
      }
      url = typeof data["odata.nextLink"] === "string" ? data["odata.nextLink"] : undefined;
    }
    if (url) {
      console.error("[CRS] fetchAllGroupMembers: more than 20 pages of site groups — refusing a partial index.");
      return undefined;
    }
    return out;
  } catch (e) {
    console.error("[CRS] fetchAllGroupMembers threw:", e);
    return undefined;
  }
}

/**
 * Members for a SET of groups, read with bounded concurrency — neither one request per group
 * sequentially (an 8-minute load across ~700 groups, Site Access, 2026-09-02) nor one giant
 * `$expand=Users` request across the whole site (`fetchAllGroupMembers` above; still 1-2 minutes,
 * confirmed live — SharePoint is slow to COMPUTE that response, not merely slow to send N of them).
 *
 * `concurrency` requests in flight at once, chunk by chunk. Kept well under the point this
 * codebase's own testing found unsafe ("ten simultaneous calls is how a tenant starts returning
 * 429s") by defaulting to 6.
 *
 * A per-group failure is recorded as `undefined` for that group, NEVER as `[]` — empty ≠ unknown,
 * same rule as everywhere else in this codebase. One group that cannot be read must not blank the
 * other 699 that could be (so it never fails the whole call), but it must also never read as "this
 * group genuinely has nobody in it" to a caller deciding something on the strength of that.
 */
export async function fetchGroupMembersBounded(
  sp: SPHttpClient,
  siteUrl: string,
  groupIds: readonly number[],
  concurrency = 6,
): Promise<Record<number, SpGroupMember[] | undefined>> {
  const out: Record<number, SpGroupMember[] | undefined> = {};
  for (let i = 0; i < groupIds.length; i += concurrency) {
    const chunk = groupIds.slice(i, i + concurrency);
    const results = await Promise.all(
      chunk.map(async (gid) => {
        try {
          return { gid, members: await getGroupMembers(sp, siteUrl, gid) as SpGroupMember[] | undefined };
        } catch {
          return { gid, members: undefined as SpGroupMember[] | undefined };
        }
      }),
    );
    for (const r of results) out[r.gid] = r.members;
  }
  return out;
}

/**
 * The site's three built-in association groups (Owners / Members / Visitors).
 *
 * These are what the old "DMS_" title filter existed to keep out of the group picker: an
 * admin mapping "Site Owners" onto a unit folder would grant that folder to everyone with
 * site ownership. With the prefix gone (2026-08-04) they have to be excluded by id.
 *
 * Cached per site for the page session — three requests, asked for on every keystroke of
 * the group search otherwise.
 *
 * A group that fails to resolve is simply omitted rather than throwing. Losing one id means
 * one built-in becomes pickable, which is visible in the list; throwing would take out the
 * group search entirely.
 */
const builtInIdCache = new Map<string, number[]>();

export async function fetchBuiltInGroupIds(sp: SPHttpClient, siteUrl: string): Promise<number[]> {
  const cached = builtInIdCache.get(siteUrl);
  if (cached) return cached;
  const ids: number[] = [];
  // Sequential, not Promise.all: three cheap requests, and per-call try/catch is required
  // anyway because Promise.allSettled is unavailable on this tsconfig target (CLAUDE.md #3).
  for (const assoc of ["AssociatedOwnerGroup", "AssociatedMemberGroup", "AssociatedVisitorGroup"]) {
    try {
      const res = await sp.get(
        `${siteUrl}/_api/web/${assoc}?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: GET_HEADERS },
      );
      if (!res.ok) continue;
      const data = await res.json();
      if (typeof data?.Id === "number") ids.push(data.Id);
    } catch {
      // Omit this one; see the note above on why this is not fatal.
    }
  }
  builtInIdCache.set(siteUrl, ids);
  return ids;
}

/**
 * Promote or demote a person as SITE COLLECTION ADMINISTRATOR.
 *
 * Client, 2026-08-27: *"its ok to allow System Admin to have SCA, their version of system admin means
 * they have access to everything."* So membership of the OWNERS group carries SCA with it - the two
 * were offered as separate decisions and the client chose to couple them, because in their model the
 * role IS "access to everything".
 *
 * WARN: THIS IS THE WIDEST GRANT IN THE SYSTEM. An SCA bypasses every ACL, reads every Highly
 * Confidential document and every pending draft, and can delete the site collection. It is also the
 * only thing that satisfies SharePoint's own site-scoped surfaces (Site Settings, Term Store Manager)
 * and the only way a non-member can read a group's membership.
 *
 * WARN: ONLY AN EXISTING SCA CAN CALL THIS. SharePoint refuses otherwise, so the FIRST administrator
 * is always set by hand - this cannot bootstrap itself. A refusal must be surfaced, never swallowed:
 * the alternative is somebody added to the group, not promoted, with nothing saying so.
 *
 * WARN: A GUEST MAY BE INELIGIBLE on tenants that block external accounts from being site collection
 * administrators. That refusal looks identical to any other and is reported the same way.
 */
export async function setSiteAdmin(
  sp: SPHttpClient,
  siteUrl: string,
  loginName: string,
  isAdmin: boolean,
): Promise<void> {
  const res = await withThrottleRetry(() =>
    sp.post(
      `${siteUrl}/_api/web/siteusers(@v)?@v='${encodeURIComponent(loginName)}'`,
      SPHttpClient.configurations.v1,
      {
        headers: { ...POST_HEADERS, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        body: JSON.stringify({ IsSiteAdmin: isAdmin }),
      },
    ),
  );
  if (res.ok) return;
  const body = await res.text().catch(() => "");
  throw new Error(`${isAdmin ? "promote" : "demote"} site admin HTTP ${res.status} ${body.slice(0, 200)}`);
}

/**
 * Whether the SIGNED-IN user is a site collection administrator - NOT merely an owner.
 *
 * Distinct from `isSystemAdmin`, which answers true for owners-group membership too. This is the
 * narrower fact, and it is the one that decides whether `setSiteAdmin` can work at all: only an
 * existing SCA may promote anyone.
 *
 * WARN: `undefined` MEANS NOT READ, and callers must not render a warning on it. Telling an
 * administrator on every visit that they cannot promote people - because one read failed - trains them
 * to ignore the line on the day it is true. Same rule as the ACL banner on Approval Library Access.
 */
export async function isSiteCollectionAdmin(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<boolean | undefined> {
  try {
    const res = await sp.get(
      `${siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!res.ok) return undefined;
    const me = await res.json();
    return me.IsSiteAdmin === true;
  } catch {
    return undefined;
  }
}

/**
 * How many site collection administrators this site has, or `undefined` if it could not be counted.
 *
 * WARN: THE GUARD AGAINST LOCKING EVERYONE OUT OF THE SITE, so `undefined` must never be read as a
 * number. Demoting the last administrator leaves a site collection nobody can administer, and no
 * screen in this app could then put it back. An uncountable answer REFUSES the demotion.
 */
export async function countSiteAdmins(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<number | undefined> {
  try {
    const res = await sp.get(
      `${siteUrl}/_api/web/siteusers?$select=Id&$filter=IsSiteAdmin eq true&$top=100`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!res.ok) return undefined;
    const data = await res.json();
    return ((data.value ?? []) as unknown[]).length;
  } catch {
    return undefined;
  }
}

/**
 * Whether this user is a SYSTEM ADMINISTRATOR of the CRS system.
 *
 * TRUE for a Site Collection Administrator **or** any member of the site OWNERS group.
 *
 * Client, 2026-08-27: *"What I essentially want is to have a group system admin where they can do
 * anything."* **`IsSiteAdmin` CANNOT BE GRANTED BY A GROUP** - Site Collection Administrator is a
 * per-USER flag and a SharePoint group can never confer it. So the group half has to be the OWNERS
 * group, which already holds Full Control on the web and is already what reconciliation's admin-page
 * lockdown grants and protects. Putting somebody in `CRS Owners` now makes them an administrator
 * everywhere THIS CODE decides, with no per-person SharePoint setting.
 *
 * WARN: THIS WAS ALREADY THE RULE IN ONE PLACE AND NOWHERE ELSE. `GroupManager.loadCanManage` had
 * checked both since it was written; the audit log, the upload form and Bulk Upload each checked
 * `IsSiteAdmin` alone - so an Owner was an administrator on one screen out of four, which is how
 * "she is in the group and still cannot get in" happens with every permission reading back correct.
 *
 * WARN: WHAT THIS CANNOT REACH. Anything SHAREPOINT itself gates on Site Collection Administrator -
 * Site Settings, Term Store Manager, site collection features - is unaffected, because that check is
 * not ours to change. Owners hold Full Control on the web, so every CRS page, library and folder is
 * covered; the site-scoped admin surfaces are not.
 *
 * FAILS CLOSED on any read failure, matching every call site it replaces: this gates controls, and
 * the pages' own permissions are what actually keep non-admins out.
 */
export async function isSystemAdmin(sp: SPHttpClient, siteUrl: string): Promise<boolean> {
  try {
    const meRes = await sp.get(
      `${siteUrl}/_api/web/currentuser?$select=Id,IsSiteAdmin`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!meRes.ok) return false;
    const me = await meRes.json();
    if (me.IsSiteAdmin === true) return true;
    if (typeof me.Id !== "number") return false;
    const ownRes = await sp.get(
      `${siteUrl}/_api/web/AssociatedOwnerGroup/Users?$filter=Id eq ${me.Id}&$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!ownRes.ok) return false;
    const own = await ownRes.json();
    return ((own.value ?? []) as unknown[]).length > 0;
  } catch {
    return false;
  }
}

/**
 * The site's OWNERS group - the one that actually confers system-administrator rights here.
 *
 * WARN: DELIBERATELY NOT PART OF `searchSiteGroups`, which excludes it. That exclusion exists so
 * nobody maps "Site Owners" onto a unit folder, and it must stay. This is the other job: Group
 * Management has to be able to put a person INTO Owners, because that is what makes them an
 * administrator of the CRS system - reconciliation's admin-page lockdown strips every non-Owners
 * assignment from every admin page, so Owners is the only membership that opens them.
 *
 * Client, 2026-08-27: *"we did not include a group as the System Admin group? they should have all
 * the power."* There was no gap in the MODEL - SharePoint's Owners group has always been it - but
 * there was no route to it in our UI, so an admin had to leave for Site Permissions to find one.
 *
 * `undefined` on any failure, never a guess. A wrong id here would show the wrong group's members
 * under a heading promising full control.
 */
export async function fetchOwnersGroup(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<SpGroup | undefined> {
  try {
    const res = await sp.get(
      `${siteUrl}/_api/web/AssociatedOwnerGroup?$select=Id,Title`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!res.ok) return undefined;
    const data = await res.json();
    if (typeof data?.Id !== "number") return undefined;
    return { id: data.Id, title: (data.Title ?? "") as string };
  } catch {
    return undefined;
  }
}

/**
 * Groups an admin may map, matching q. Built-ins are resolved and excluded HERE rather than
 * by the caller, so no call site can forget to pass them — the previous title-prefix filter
 * was unmissable by construction, and this keeps that property.
 */
export async function searchSiteGroups(
  sp: SPHttpClient,
  siteUrl: string,
  q: string,
  extraExcludeIds?: number[],
): Promise<SpGroup[]> {
  const [all, builtIns] = [
    await fetchAllSiteGroups(sp, siteUrl),
    await fetchBuiltInGroupIds(sp, siteUrl),
  ];
  return filterSelectableGroups(all, q, [...builtIns, ...(extraExcludeIds ?? [])]);
}

/**
 * Create a site group with NO permissions (Reconciliation grants folder access
 * later, so creating a group is a safe, reversible act). Throws Error(DUPLICATE_GROUP)
 * if the title already exists in the site collection.
 */
/**
 * Rename an existing site group in place.
 *
 * WARN: THE ID SURVIVES A RENAME, and that is the only reason this is safe to do in bulk. Every
 * Group Map row is keyed on `GroupId`, so mappings, folder grants and page ACLs are untouched; only
 * the stored `GroupName` label on those rows goes stale. `planBulkGroups` also matches on the TERM
 * GUID before the name (1.0.162.0), so a renamed group is still recognised as provisioned.
 *
 * A duplicate title is reported as `DUPLICATE_GROUP`, exactly as creation is - renaming
 * `X_APR_HIGHLY_CONFIDENTIAL` when `X_APPROVER_HIGHLY_CONFIDENTIAL` already exists must not silently
 * do nothing.
 */
export async function renameSiteGroup(
  sp: SPHttpClient,
  siteUrl: string,
  id: number,
  title: string,
): Promise<void> {
  const res = await withThrottleRetry(() =>
    sp.post(`${siteUrl}/_api/web/sitegroups/getbyid(${id})`, SPHttpClient.configurations.v1, {
      headers: { ...POST_HEADERS, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
      body: JSON.stringify({ Title: title.trim() }),
    }),
  );
  if (res.ok) return;
  if (res.status === 429 || res.status === 503) return fail("rename group", res);
  const body = await res.text().catch(() => "");
  if (body.toLowerCase().indexOf("already in use") !== -1) throw new Error(DUPLICATE_GROUP);
  throw new Error(`rename group HTTP ${res.status} ${body.slice(0, 200)}`);
}

export async function createSiteGroup(
  sp: SPHttpClient,
  siteUrl: string,
  title: string,
): Promise<SpGroup> {
  const res = await withThrottleRetry(() =>
    sp.post(`${siteUrl}/_api/web/sitegroups`, SPHttpClient.configurations.v1, {
      headers: POST_HEADERS,
      body: JSON.stringify({ Title: title.trim() }),
    }),
  );
  if (!res.ok) {
    // A throttle is not a duplicate, and its HTML body must not be pattern-matched for
    // "already in use" — hand it to `fail`, which says so in a sentence.
    if (res.status === 429 || res.status === 503) return fail("create group", res);
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

/**
 * Resolve a raw email address to a site principal, creating the site user if needed.
 *
 * Client, 2026-08-27: *"I cannot add new email at all, is there no button for this."* The member
 * editor could only add somebody the PEOPLE PICKER returned, and `searchTenantPeople` suggests
 * nothing for an external address this tenant's picker will not surface - so an address like a
 * `hotmail.com` guest was simply unaddable, with no button and no Enter handler to fall back on.
 *
 * `ensureuser` is the mechanism SharePoint's own Share dialog uses: it returns the canonical claim
 * for an address, inviting an external as a guest where the tenant allows it.
 *
 * WARN: A REFUSAL HERE IS MEANINGFUL AND MUST REACH THE ADMIN. It means the address belongs to
 * nobody, or external sharing forbids it, or it matches more than one account - and a guest CAN exist
 * twice for one address on this tenant (2026-08-18). `explainAddFailure` already translates the
 * message SharePoint returns; do not swallow it.
 */
export async function ensureSiteUser(
  sp: SPHttpClient,
  siteUrl: string,
  logonName: string,
): Promise<PersonPick> {
  const res = await withThrottleRetry(() =>
    sp.post(`${siteUrl}/_api/web/ensureuser`, SPHttpClient.configurations.v1, {
      headers: POST_HEADERS,
      body: JSON.stringify({ logonName: logonName.trim() }),
    }),
  );
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`HTTP ${res.status} ${body.slice(0, 300)}`);
  }
  const d = await res.json();
  return {
    loginName: (d.LoginName ?? "") as string,
    displayName: (d.Title ?? logonName) as string,
    email: (d.Email ?? "") as string,
  };
}

/* ⚠⚠ `inviteToGroup` WAS HERE AND IS DELETED, NOT PARKED (2026-08-27, 1.0.302.0).
   Client, within the hour of it being built: *"client doesn't want to allow us to add outsiders, so
   we got to revert again."* So the ability to invite an external address into a SharePoint group is
   now against policy, and this codebase's habit of leaving unreachable code in place would be wrong
   here: a parked EXPORTED function that does exactly the forbidden thing is an invitation for the
   next session to wire it back up, and lint cannot flag an unused export. Same reasoning that fully
   removed `requestCommand` and `overwrite=true` rather than leaving them dormant.

   THE SHAPE IS RECORDED IN CLAUDE.md so it is recoverable in minutes if the client reverses again:
   `POST {web}/_api/SP.Web.ShareObject` with `roleValue: "group:<spGroupId>"`, `groupId: 0`, the
   address as `peoplePickerInput: JSON.stringify([{ Key: address }])`, and the per-recipient verdict
   read from the BODY because a refusal arrives as `Status: false` behind an HTTP 200.

   WHAT REMAINS TRUE AND MATTERS: `ensureuser` RESOLVES, it does not INVITE. An address nobody has
   invited has no principal, so it answers `"could not be found"` — which is now the correct and
   final answer for an outsider, and `explainAddFailure` states the policy rather than a fault. */

/** Add by login name (claim). SharePoint ensures the user on the site automatically; guests work. */
export async function addGroupMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  loginName: string,
): Promise<void> {
  const res = await withThrottleRetry(() =>
    sp.post(`${siteUrl}/_api/web/sitegroups(${groupId})/users`, SPHttpClient.configurations.v1, {
      headers: POST_HEADERS,
      body: JSON.stringify({ LoginName: loginName }),
    }),
  );
  if (!res.ok) return fail("add member", res);
}

export async function removeGroupMember(
  sp: SPHttpClient,
  siteUrl: string,
  groupId: number,
  userId: number,
): Promise<void> {
  const res = await withThrottleRetry(() =>
    sp.post(
      `${siteUrl}/_api/web/sitegroups(${groupId})/users/removebyid(${userId})`,
      SPHttpClient.configurations.v1,
      { headers: POST_HEADERS },
    ),
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
  const res = await withThrottleRetry(() =>
    sp.post(
      `${siteUrl}/_api/web/sitegroups/removebyid(${groupId})`,
      SPHttpClient.configurations.v1,
      { headers: POST_HEADERS },
    ),
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
