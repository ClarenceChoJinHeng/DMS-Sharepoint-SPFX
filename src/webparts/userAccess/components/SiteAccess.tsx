// Site Access tab — who can open the site at all, and the membership that makes it work.
//
// This is the layer everything else sits on. A folder grant gives a user Limited Access up the
// parent chain, which resolves a DIRECT link but confers no View Pages — so without site-level
// Read a correctly provisioned uploader still cannot open the site's home page.
//
// READ-ONLY since 2026-09-02 (docs/superpowers/specs/2026-09-02-access-pages-read-only-design.md).
// The client's own words: "the group assignment is all done by Folder recon... so I am see there
// is no point for all this pages to allow client to manually strip or assign group to a page or
// library." Reconciliation now asserts the site-entry grant and its membership is managed on
// Group Management, so this page only shows the three facts that used to be actionable here:
//   1. Does the site-entry group EXIST and hold a role on the web?
//   2. Who is in it?
//   3. Who is in a mapped group but NOT in it? That set is the silent failure — those users
//      have folder permissions and cannot reach the site. Only reconciliation heals it now.
//
// See docs/superpowers/specs/2026-08-03-access-scope-mapping-design.md and memory
// dms-two-layer-access-site-plus-folder.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { findSiteEntryGroup, siteEntryGroupTitle } from "../../../shared/groupMapModel";
import { clearSiteEntryCache } from "../../../shared/siteEntryGroup";
import {
  fetchAllSiteGroups,
  fetchGroupMembersBounded,
  getGroupMembers,
  SpGroup,
} from "../../../shared/spGroups";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { CARDS, resolveLink } from "../../../shared/adminPages";
import { readSitePages } from "../../../shared/backToSettings";
import { MemberCountWithPopup, useGroupMembers } from "./accessMemberUi";

type Props = { context: WebPartContext; siteUrl: string };

// Resolved, not hardcoded — the client renames this to "CRS Group Map" at import (confirmed on
// their site 2026-08-05). Read from the cache primed in reload(), which runs on mount before the
// Group Map is read.
const groupMapList = (): string => cachedListTitle(LIST_SUFFIX.groupMap);

type WebGrant = { principalId: number; title: string; levels: string[] };
/** A user who holds a mapped group but is missing site entry. */
type Missing = { loginName: string; title: string; via: string };

const s: Record<string, React.CSSProperties> = {
  wrap:     { fontSize: 13, color: "#242424", lineHeight: 1.5 },
  intro:    { fontSize: 13, color: "#444", margin: "0 0 16px" },
  card:     { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  head:     { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  table:    { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  th:       { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:       { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "middle" },
  okBox:    { marginBottom: 16, padding: "10px 12px", border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 4, fontSize: 12, color: "#0f6c3f", lineHeight: 1.5 },
  warnBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  dangerBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  hint:     { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
  no:       { color: "#8a8886" },
  mono:     { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
};

/* Shared filter-bar look, matching the Requests page (Requests.tsx FIL_*) — one set of styles for
   every filter box on this project rather than a fourth independent implementation. */
const FIL_WRAP: React.CSSProperties = {
  display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", margin: "8px 0 0",
};
const FIL_INPUT: React.CSSProperties = {
  padding: "6px 8px", border: "1px solid #c8c6c4", borderRadius: 2, fontSize: 13,
  flex: "1 1 220px", minWidth: 180,
};
const FIL_CLEAR: React.CSSProperties = {
  padding: "6px 10px", border: "1px solid #c8c6c4", borderRadius: 2, background: "#fff",
  fontSize: 13, cursor: "pointer",
};
/* ⚠ 60vh. Safe because nothing in either section is absolutely positioned — the one popup here
   (`MemberCountWithPopup`'s modal, on group rows in "Everything with permission") is
   `position: fixed`, which escapes an ancestor's `overflow` clipping regardless of DOM nesting.
   An absolutely-positioned popover would NOT be safe here (it has already broken Group
   Management's people picker, the upload form's info panels and a comment-box dropdown) —
   re-check before adding one. */
const SCROLLER: React.CSSProperties = {
  maxHeight: "60vh", overflowY: "auto", overflowX: "hidden", paddingRight: 4, marginTop: 8,
};

// ⚠ `simplifyLevels` (the "Full access" / "Reaches folder/page, no site-wide access" three-bucket
// permission display, built earlier the same day) is GONE — the "Permission" column it fed was
// removed on the client's explicit instruction ("remove the permission, they dont want that"), and
// it had no other caller. Unlike the create-one-group form elsewhere in this project, there is no
// stated ongoing need for this one, so it is deleted rather than kept-but-unreachable.

export default function SiteAccess({ context, siteUrl }: Props): React.ReactElement {
  const [entry, setEntry]     = useState<SpGroup | undefined>(undefined);
  const [entryHasRole, setEntryHasRole] = useState<boolean | undefined>(undefined);
  const [webGrants, setWebGrants] = useState<WebGrant[] | undefined>(undefined);
  const [missing, setMissing] = useState<Missing[]>([]);
  // Empty ≠ unknown: a failed `fetchGroupMembersBounded` read on any one group must never be
  // silently read as "no gaps" — that would report every genuinely missing person as found, on
  // the one banner whose whole job is surfacing a silent failure.
  const [missingUnknown, setMissingUnknown] = useState(false);
  // Every site group, kept for render-time lookup: which "Everything with permission" rows are
  // groups (worth a member popup) versus individual people (nothing to show).
  const [allGroups, setAllGroups] = useState<SpGroup[] | undefined>(undefined);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  // Filter for "Everything with permission on the site itself" only (per the design spec) — the
  // site-entry member list has no equivalent control, just the scroll container.
  const [webFilterText, setWebFilterText] = useState("");

  // Where "Go to Group Management" points — resolved from Site Pages, never hardcoded, because the
  // client renames every page at import. Same lookup and the same `resolveLink` as the retired
  // Folder Access signpost (FolderAccessPage.tsx), which answers the identical question.
  const [groupManagementHref, setGroupManagementHref] = useState<string | undefined>(undefined);

  // Member data for every site group, for the popup on "Everything with permission" rows that are
  // groups. `useGroupMembers` is ONE batched request regardless of how many ids are passed, so
  // there is no cost to asking for all of them rather than filtering to just the visible rows.
  const { members: groupMembers } = useGroupMembers(
    context.spHttpClient,
    siteUrl,
    (allGroups ?? []).map((g) => g.id),
  );

  const GET = { Accept: "application/json;odata=nometadata" };

  const loadWebGrants = async (): Promise<WebGrant[] | undefined> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings` +
        `&$select=PrincipalId,Member/Title,RoleDefinitionBindings/Name`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return undefined;
    const j = await res.json();
    return ((j.value ?? []) as Array<{ PrincipalId?: number; Member?: { Title?: string }; RoleDefinitionBindings?: Array<{ Name?: string }> }>)
      .map((ra) => ({
        principalId: ra.PrincipalId ?? 0,
        title: ra.Member?.Title ?? "",
        levels: (ra.RoleDefinitionBindings ?? []).map((b) => b.Name ?? "").filter(Boolean),
      }));
  };

  /**
   * Every group id the Group Map references.
   *
   * GroupId ONLY: that column predates Scope/Target, so this cannot hit the whole-request 400
   * that naming a nonexistent $select column causes (CLAUDE.md #11), and it covers every scope
   * in one read.
   */
  const loadMappedGroupIds = async (): Promise<number[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(groupMapList())}')/items?$select=GroupId&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return [];
    const j = await res.json();
    const ids = new Set<number>();
    for (const r of (j.value ?? []) as Array<{ GroupId?: string }>) {
      const n = Number((r.GroupId ?? "").toString().trim());
      if (n > 0 && n % 1 === 0) ids.add(n);
    }
    return Array.from(ids);
  };

  const reload = async (firstLoad = false): Promise<void> => {
    if (firstLoad) setLoading(true);
    try {
      // Before the Group Map read: resolves "CRS Group Map" vs "DMS Group Map" once per session.
      await primeNames(context.spHttpClient, siteUrl);
      const all = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
      setAllGroups(all);
      const eg = findSiteEntryGroup(all);
      setEntry(eg);
      // This page is where the entry group gets created, so its cached lookup elsewhere on the
      // site would otherwise stay "absent" for the rest of the session.
      clearSiteEntryCache(siteUrl);
      const grants = await loadWebGrants();
      setWebGrants(grants);
      setEntryHasRole(
        eg === undefined ? false : grants === undefined ? undefined : grants.some((g) => g.principalId === eg.id),
      );

      if (!eg) { setMissing([]); setLoadError(undefined); return; }
      const mem = await getGroupMembers(context.spHttpClient, siteUrl, eg.id);

      // Who holds a mapped group but is missing from site entry? Computed here rather than
      // waiting for a reconciliation run, because this is the state that presents as "only some
      // users cannot open the site" and it is invisible everywhere else.
      const have = new Set(mem.map((m) => m.loginName.toLowerCase()));
      const mappedIds = await loadMappedGroupIds();
      // ⚠ WAS (twice): one `getGroupMembers` request PER mapped group, sequentially — ~700 round
      // trips, an 8-minute load. Then ONE `fetchAllGroupMembers` request across the whole site —
      // still 1-2 minutes, confirmed live (2026-09-02): SharePoint is slow to COMPUTE that one huge
      // `$expand=Users` response, not merely slow to send many small ones. `fetchGroupMembersBounded`
      // reads only the groups actually needed here, with several in flight at once rather than one
      // request or all of them — the fix that actually worked.
      const allMembers = await fetchGroupMembersBounded(context.spHttpClient, siteUrl, mappedIds);
      const gap = new Map<string, Missing>();
      // A group whose read genuinely failed (`undefined`, never `[]` — empty ≠ unknown) makes the
      // WHOLE check unknown rather than silently under-reporting: a "no gaps" banner built from a
      // partial read is a false all-clear on the one screen whose job is catching exactly this.
      let anyFailed = false;
      for (const gid of mappedIds) {
        if (gid === eg.id) continue;
        const g = all.find((x) => x.id === gid);
        const gm = allMembers[gid];
        if (gm === undefined) { anyFailed = true; continue; }
        for (const m of gm) {
          const key = m.loginName.toLowerCase();
          if (have.has(key) || gap.has(key)) continue;
          gap.set(key, { loginName: m.loginName, title: m.title, via: g?.title ?? String(gid) });
        }
      }
      setMissing(Array.from(gap.values()));
      setMissingUnknown(anyFailed);
      setLoadError(undefined);
    } catch (e) {
      setLoadError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { reload(true).catch(() => undefined); }, []);

  /**
   * Resolve "Go to Group Management" from Site Pages, never hardcoded — the client renames every
   * page at import, so a literal `Group-Management.aspx` would fail as a dead link. Same lookup,
   * same `resolveLink`, as the retired Folder Access signpost.
   *
   * A failed read simply leaves the link out; the banner's own text still names the page.
   */
  useEffect(() => {
    const link = CARDS
      .filter((c) => c.key === "access")[0]
      ?.links.filter((l) => l.key === "groups")[0];
    if (!link) return;
    readSitePages(context, siteUrl)
      .then((pages) => {
        const t = resolveLink(link, pages);
        if (t.state !== "missing") setGroupManagementHref(t.url);
      })
      .catch(() => undefined);
  }, []);

  const setUpDone = entry !== undefined && entryHasRole === true;

  return (
    <div style={s.wrap}>
      <p style={s.intro}>
        Who can <strong>open this site</strong>. Folder and library permissions are not enough on
        their own — without site access a user can only reach a folder by direct link, and the home
        page refuses them. Everyone who uses the system needs to be here.
      </p>

      <div style={s.warnBox}>
        <strong>This page is read-only.</strong> Group membership — and with it, who can open the
        site — is now set entirely by Folder Reconciliation from the CRS Group Map, so there is
        nothing to strip or assign here by hand. To change who has access,{" "}
        {groupManagementHref !== undefined
          ? <a href={groupManagementHref} style={{ fontWeight: 600 }}>go to Group Management</a>
          : <strong>go to the Group Management page</strong>}.
      </div>

      {loadError && <div style={s.dangerBox}>Could not load: {loadError}</div>}

      {loading ? (
        <div style={s.card}><span style={s.no}>Loading&hellip;</span></div>
      ) : (
        <>
          {!setUpDone && (
            <div style={s.dangerBox}>
              {!entry ? (
                <>
                  <strong>{siteEntryGroupTitle()} does not exist.</strong> Until it does, nobody can
                  open this site except owners and administrators.
                </>
              ) : entryHasRole === undefined ? (
                <>
                  <strong>{siteEntryGroupTitle()} exists</strong>, but its site permissions could
                  not be read, so whether it works is unknown.
                </>
              ) : (
                <>
                  <strong>{siteEntryGroupTitle()} exists but holds no permission on the site</strong>,
                  so its members still cannot open it.
                </>
              )}
            </div>
          )}

          {setUpDone && (
            <div style={s.okBox}>
              <strong>{siteEntryGroupTitle()}</strong> holds <strong>Read</strong> on the site.
              Anyone in it can open the site; what they see inside is decided by their folder and
              library permissions.
            </div>
          )}

          {missingUnknown && entry && (
            <div style={s.warnBox}>
              <strong>Could not read every mapped group&rsquo;s members</strong>, so whether anyone is
              missing site access could not be checked this time. This is not the same as
              &ldquo;nobody is missing&rdquo; — try refreshing the page.
            </div>
          )}

          {!missingUnknown && missing.length > 0 && entry && (
            <div style={s.warnBox}>
              <strong>{missing.length} user(s) have folder permissions but cannot open the
              site.</strong> They were added to a group directly in SharePoint, which does not grant
              site access on its own — so for them the system looks broken while everyone else is
              fine.
              <div style={s.mono}>
                {missing.slice(0, 12).map((m) => `${m.title} (via ${m.via})`).join(", ")}
                {missing.length > 12 ? `, +${missing.length - 12} more` : ""}
              </div>
            </div>
          )}

          {/* ⚠ "People who can open the site" — REMOVED FROM DISPLAY (2026-09-02, client's request).
              `members` (the site-entry group's own membership) is still read and still feeds the
              "missing from site entry" check above — that is a different feature the client did not
              ask to remove, and dropping the fetch would break it. Only this table went. */}

          <div style={s.card}>
            <div style={s.head}>Everything with permission on the site itself</div>
            <div style={s.hint}>
              Read-only. Site-level permissions reach every library and folder that still inherits,
              so anything unexpected here is worth understanding before it is removed — removing the
              wrong entry can lock administrators out of the site.
            </div>
            {webGrants === undefined ? (
              <div style={s.no}>Could not read the site&rsquo;s permissions.</div>
            ) : (
              <>
                <div style={FIL_WRAP}>
                  <input
                    value={webFilterText}
                    onChange={(e) => setWebFilterText(e.target.value)}
                    placeholder="Filter by name…"
                    style={FIL_INPUT}
                  />
                  {webFilterText !== "" && (
                    <button onClick={() => setWebFilterText("")} style={FIL_CLEAR}>Clear</button>
                  )}
                </div>
                {(() => {
                  // ⚠ GROUPS ONLY (client, 2026-09-02: "dont show individual person under the
                  // name only show the group"). Fails OPEN when `allGroups` has not loaded yet
                  // (shown, not hidden) — a row missing here because a read is still in flight is
                  // worse than one extra row for a moment; `allGroups` loads in the same `reload()`
                  // pass as `webGrants`, so this is normally never reached in practice.
                  const isGroup = (pid: number): boolean =>
                    allGroups === undefined || allGroups.some((sg) => sg.id === pid);
                  const q = webFilterText.trim().toLowerCase();
                  const shown = webGrants
                    .filter((g) => isGroup(g.principalId))
                    .filter((g) => q === "" || (g.title || String(g.principalId)).toLowerCase().indexOf(q) !== -1);
                  return shown.length === 0 ? (
                    <div style={s.no}>
                      {q === "" ? "No groups hold permission on the site." : <>No match for &ldquo;{webFilterText}&rdquo;.</>}
                    </div>
                  ) : (
                    <div style={SCROLLER}>
                      <table style={s.table}>
                        <thead>
                          <tr>
                            <th style={s.th}>Name</th>
                            <th style={s.th}>Members</th>
                          </tr>
                        </thead>
                        <tbody>
                          {shown.map((g) => {
                            const asGroup = allGroups?.find((sg) => sg.id === g.principalId);
                            return (
                              <tr key={g.principalId}>
                                <td style={s.td}>{g.title || g.principalId}</td>
                                <td style={s.td}>
                                  {asGroup
                                    ? <MemberCountWithPopup group={asGroup} members={groupMembers} />
                                    : <span style={s.no}>—</span>}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  );
                })()}
              </>
            )}
          </div>
        </>
      )}
    </div>
  );
}
