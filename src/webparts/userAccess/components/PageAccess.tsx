// Page Access tab — who may open each page in Site Pages.
//
// READ-ONLY since 2026-09-02 (spec docs/superpowers/specs/2026-09-02-access-pages-read-only-design.md).
// Client: "the group assignment is all done by Folder recon, and how a group is have access to
// anything is also done by folder recon, so I am see there is no point for all this pages to allow
// client to manually strip or assign group to a page or library." Every write action that used to
// live here — allow/remove a group, restrict an unlocked page, reopen a locked one, and the
// per-person removal shared with Approval Library Access — is gone. Group Management (creating a
// group, deleting it, and adding/removing its members) is the one remaining route for all of it;
// deleting a group's mapping rows and re-running reconciliation is the standing repair pattern for
// narrowing what a page or library grants.
//
// ⚠ THE CLIENT WAS EXPLICIT THAT A LOCKED PAGE MUST NOT GET AN UNLOCK BUTTON BACK: "when client
// locks a page, they won't want the page to be reopen for normal clients to see, that beats the
// whole purpose of ensuring normal user wont mess with or gain access into CRS Settings." So
// "Reopen to everyone" (resetroleinheritance) has no replacement of any kind — not even a guarded
// or admin-only version of it.
//
// What stays: the live-ACL-vs-Group-Map drift note per row, member counts (now with a popup to see
// who), and the policy explanation of who a page is FOR. Those are read-only findings, not actions,
// and they are still exactly what an admin needs to notice a page whose SharePoint permissions have
// quietly drifted from what Folder Reconciliation believes it granted.
//
// ⚠ USABILITY PASS, 2026-09-02 (spec `2026-09-02-access-pages-usability-pass-design.md`), driven by
// the client's own walkthrough of the read-only page above. Three changes worth recording:
//   1. The "This hides the page, not the documents on it" explainer paragraph is GONE from the UI.
//      It is still TRUE — a restricted page's web parts stay reachable by direct link, and
//      SharePoint does not security-trim nav links — but a long paragraph of caveats on a page
//      that no longer lets anyone act on them reads as noise, not a warning. The home-page exclusion
//      is still stated, in the hint under the page picker.
//   2. The table now answers "which groups CURRENTLY have access", not "which groups match this
//      page's naming policy" — every `_APPROVER`-suffixed group site-wide used to be listed for a
//      page like Approval Document, most reading "no members" / "granted in SharePoint, not
//      mapped", which is noise on a page that can no longer fix a mismatch. Only groups holding a
//      REAL binding on the page's live ACL are shown now (a genuine permission level, never a bare
//      "Limited Access"/"Web-Only Limited Access" automatic traversal entry — this project's own
//      convention treats those as granting nothing). The "Show all N groups" escape hatch is gone
//      with the policy-derived list it existed to override.
//   3. "N members" is now clickable — a popup lists who, reusing the same `members` map the
//      collapsed count already read from `useGroupMembers`, so there is still only one fetch.
//
// See docs/superpowers/specs/2026-08-03-access-scope-mapping-design.md §5 and §5a,
// docs/superpowers/specs/2026-09-02-access-pages-read-only-design.md, and
// docs/superpowers/specs/2026-09-02-access-pages-usability-pass-design.md §3 for this change.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { isForbiddenPageTarget, siteEntryGroupTitle } from "../../../shared/groupMapModel";
import { fetchAllSiteGroups, fetchBuiltInGroupIds, SpGroup } from "../../../shared/spGroups";
import { policyForPage, VIEW_ONLY_ROLES } from "../../../shared/pageAccessPolicy";
import { roleFromGroupName } from "../../../shared/groupMapModel";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { CARDS, resolveLink } from "../../../shared/adminPages";
import { readSitePages } from "../../../shared/backToSettings";
import { MemberCountWithPopup, useGroupMembers } from "./accessMemberUi";

type Props = { context: WebPartContext; siteUrl: string };

// Resolved, not hardcoded — the client renames this to "CRS Group Map" at import (confirmed on
// their site 2026-08-05). Read from the cache primed in reloadAll(), which runs on mount before
// any read this component makes.
const groupMapList = (): string => cachedListTitle(LIST_SUFFIX.groupMap);
const PAGES_LIST = "Site Pages";

type PageItem = { itemId: number; fileName: string; title: string; unique: boolean };
type EntryRow = { itemId: number; groupId: string; groupName: string; target: string };
type LiveGrant = { principalId: number; title: string; levels: string[] };

/**
 * Permission levels that mean "reaches one thing further down, nothing site-wide" — SharePoint's
 * automatic traversal grants, never something handed out by hand and never something that lets a
 * principal browse or list anything on its own. A group holding ONLY these on a page's ACL does not
 * currently have access to the PAGE; it merely has access to something below it in the site.
 */
const NON_REAL_LEVELS = ["Limited Access", "Web-Only Limited Access"];

/** A real, currently-in-force grant — at least one binding beyond an automatic traversal entry. */
function hasRealGrant(grant: LiveGrant | undefined): boolean {
  return grant !== undefined && grant.levels.some((n) => NON_REAL_LEVELS.indexOf(n) === -1);
}

const s: Record<string, React.CSSProperties> = {
  wrap:     { fontSize: 13, color: "#242424", lineHeight: 1.5 },
  intro:    { fontSize: 13, color: "#444", margin: "0 0 16px" },
  card:     { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  head:     { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  label:    { display: "block", fontWeight: 600, fontSize: 12, margin: "0 0 4px" },
  select:   { width: "100%", maxWidth: 420, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  table:    { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  th:       { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:       { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "middle" },
  yes:      { color: "#0f6c3f", fontWeight: 600 },
  no:       { color: "#8a8886", fontWeight: 600 },
  mono:     { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  drift:    { fontSize: 11, color: "#8a4b00", marginTop: 2 },
  policyBox:{ marginBottom: 12, padding: "8px 10px", border: "1px solid #d6e8dc", background: "#f6fbf8", borderRadius: 4, fontSize: 12, color: "#265", lineHeight: 1.5 },
  adminBox: { marginBottom: 12, padding: "8px 10px", border: "1px solid #cfd8e3", background: "#f4f7fb", borderRadius: 4, fontSize: 12, color: "#2b3f56", lineHeight: 1.5 },
  warnBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  dangerBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  openBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #c7c7c7", background: "#fff", borderRadius: 4, fontSize: 12, color: "#444", lineHeight: 1.5 },
  hint:     { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
  // Filter box, matching Group Management's / Approval Library Access's "Filter by name…" input.
  filterRow:{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 },
  input:    { width: "100%", maxWidth: 320, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  // Inner scroll for the group table — see the check above the table for why nothing here clips.
  scroller: { maxHeight: "60vh", overflowY: "auto" },
  // Member popup. Same fixed-overlay pattern as Group Management's delete confirmation
  // (GroupManager.tsx `s.modalBg`/`s.modal`) — reused rather than a fourth definition of a modal.
  modalBg:  { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" },
  modal:    { background: "#fff", borderRadius: 6, padding: 20, width: "min(420px, 92vw)", maxHeight: "76vh", overflowY: "auto", boxShadow: "0 8px 32px rgba(0,0,0,.25)" },
  closeBtn: { padding: "6px 14px", border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", fontSize: 12, cursor: "pointer" },
  memberList:{ margin: "10px 0 0", padding: "0 0 0 18px", fontSize: 12.5, lineHeight: 1.8 },
};

export default function PageAccess({ context, siteUrl }: Props): React.ReactElement {
  const [pages, setPages]     = useState<PageItem[]>([]);
  const [target, setTarget]   = useState<string>("");
  const [groups, setGroups]   = useState<SpGroup[]>([]);
  /* ⚠ THE GROUP MAP ROWS ARE NO LONGER STORED (2026-09-03). They fed the "Current access" column's
     drift note ("granted in SharePoint, not mapped"), and that column is gone at the client's
     request — but `loadRows()` STILL RUNS, because `setScopeMissing` is set inside it: a `$select`
     naming Scope/Target failing with 400 is the only reliable signal those columns are absent
     (CLAUDE.md #11). So the read stays for its side effect; only the result is discarded. */

  const [live, setLive]       = useState<LiveGrant[] | undefined>(undefined);
  const [welcome, setWelcome] = useState<string>("");
  const [loading, setLoading] = useState(true);
  const [scopeMissing, setScopeMissing] = useState(false);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);

  // Search/filter over the group list — it can run to 60+ groups per segment. Reset whenever the
  // page changes, so a filter typed for one page cannot silently hide rows on the next.
  const [filter, setFilter] = useState("");

  // "Go to Group Management", resolved from Site Pages rather than hardcoded — the client renames
  // every page at import, so a literal `Group-Management.aspx` would fail as a dead link. Same
  // lookup, same `resolveLink`, as Site Access and the retired Folder Access signpost.
  const [groupManagementHref, setGroupManagementHref] = useState<string | undefined>(undefined);

  const GET = { Accept: "application/json;odata=nometadata" };
  const pagesBase = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PAGES_LIST)}')`;

  // ── Derived ────────────────────────────────────────────────────────────────
  // ⚠ ADMIN-ONLY PAGES DROPPED FROM THE PICKER TOO (client's mockup, 2026-09-03) — CRS Settings,
  // Approval Library Access, CRS Audit Log, CRS Configuration, Folder Access, Folder Manager, Page
  // Access itself and Site Access. Reuses `policyForPage(...).adminOnly`, the SAME flag reconciliation
  // already locks these pages by — never a second name list here, which would drift from the one that
  // actually decides who can open them. `isForbiddenPageTarget` is untouched: it also gates the
  // site-entry page for reconciliation's own lockdown pass, and widening it here would risk unlocking
  // pages it currently locks, not just hiding them from this dropdown.
  const selectablePages = pages.filter(
    (p) =>
      !isForbiddenPageTarget(p.fileName) &&
      p.fileName.toLowerCase() !== welcome &&
      !policyForPage(p.fileName).adminOnly,
  );
  const page = pages.find((p) => p.fileName === target);

  const liveByPid = new Map<number, LiveGrant>();
  for (const l of live ?? []) liveByPid.set(l.principalId, l);

  const policy = policyForPage(target);

  /**
   * Groups worth ever considering for this page. Two exclusions are absolute:
   *   - the site-entry group, because it contains everyone;
   *   - view-only roles (MEMBER / GLOBAL / SEGVIEW), which never open an upload/approval-shaped page.
   * This is a pool to filter FROM, not the answer — the answer is `granted`, below.
   */
  const allGroups = groups
    .filter((g) => g.title.trim().toLowerCase() !== siteEntryGroupTitle().toLowerCase())
    .filter((g) => VIEW_ONLY_ROLES.indexOf(roleFromGroupName(g.title)) === -1)
    .sort((a, b) => a.title.localeCompare(b.title));

  /**
   * Groups that ACTUALLY currently have access to this page — a real binding on the live ACL, never
   * a policy guess (2026-09-02 usability pass). Before this, every group whose NAME matched the
   * page's derived policy was listed regardless of whether it held any grant at all — on a page like
   * Approval Document that is every `_APPROVER`-suffixed group site-wide, most reading "no members"
   * / "granted in SharePoint, not mapped". `live === undefined` means the ACL could not be read (or
   * has not been read yet for this page), which must never be shown as "nobody has access" — it
   * means "unknown", so nothing is listed as granted until it genuinely is.
   */
  const liveKnown = live !== undefined;
  const granted = liveKnown ? allGroups.filter((g) => hasRealGrant(liveByPid.get(g.id))) : [];

  const visible = granted.filter(
    (g) => !filter.trim() || g.title.toLowerCase().indexOf(filter.trim().toLowerCase()) !== -1,
  );

  /**
   * Member counts for every group on screen, fetched one at a time by the shared hook.
   *
   * Read-only: this is a COUNT, not the per-person removal UI that used to live beside it. Keyed
   * on the granted groups rather than only the mapped ones, so the count is there for every row the
   * table can actually show.
   */
  const { members } = useGroupMembers(
    context.spHttpClient,
    siteUrl,
    granted.map((g) => g.id),
  );

  /**
   * The site's welcome page, read rather than assumed.
   *
   * isForbiddenPageTarget knows "home.aspx", which covers the default. A site whose welcome
   * page was renamed would slip past that constant, so the real value is read and excluded too.
   */
  const loadWelcome = async (): Promise<string> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/RootFolder?$select=WelcomePage`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return "";
    const j = await res.json();
    // "SitePages/Home.aspx" -> "home.aspx"
    return ((j.WelcomePage ?? "") as string).split("/").pop()?.toLowerCase() ?? "";
  };

  const loadPages = async (): Promise<PageItem[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${pagesBase}/items?$select=Id,FileLeafRef,Title,HasUniqueRoleAssignments&$top=500`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) throw new Error(`${PAGES_LIST} HTTP ${res.status}`);
    const j = await res.json();
    return ((j.value ?? []) as Array<{ Id: number; FileLeafRef?: string; Title?: string; HasUniqueRoleAssignments?: boolean }>)
      .map((p) => ({
        itemId: p.Id,
        fileName: p.FileLeafRef ?? "",
        title: p.Title || (p.FileLeafRef ?? ""),
        unique: p.HasUniqueRoleAssignments === true,
      }))
      .filter((p) => p.fileName.toLowerCase().indexOf(".aspx") !== -1)
      .sort((a, b) => a.title.localeCompare(b.title));
  };

  const loadRows = async (): Promise<EntryRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(groupMapList())}')/items` +
        `?$select=Id,GroupId,GroupName,Role,Scope,Target&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    // A $select naming a column that does not exist fails the WHOLE request with 400
    // (CLAUDE.md #11), so this is the only reliable signal that Scope/Target are absent.
    if (!res.ok) { setScopeMissing(true); return []; }
    setScopeMissing(false);
    const j = await res.json();
    return ((j.value ?? []) as Array<{ Id: number; GroupId?: string; GroupName?: string; Scope?: string; Target?: string }>)
      .filter((r) => (r.Scope ?? "").trim().toLowerCase() === "page")
      .map((r) => ({
        itemId: r.Id,
        groupId: (r.GroupId ?? "").trim(),
        groupName: r.GroupName ?? "",
        target: (r.Target ?? "").trim(),
      }));
  };

  const loadLiveFor = async (list: PageItem[], fileName: string): Promise<LiveGrant[] | undefined> => {
    const p = list.find((x) => x.fileName === fileName);
    if (!p) return undefined;
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${pagesBase}/items(${p.itemId})/roleassignments?$expand=Member,RoleDefinitionBindings` +
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

  const reloadAll = async (firstLoad = false): Promise<void> => {
    if (firstLoad) setLoading(true);
    try {
      // Before any list read: resolves "CRS Group Map" vs "DMS Group Map" once per session.
      await primeNames(context.spHttpClient, siteUrl);
      const w = await loadWelcome();
      const p = await loadPages();
      const g = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
      const builtIns = await fetchBuiltInGroupIds(context.spHttpClient, siteUrl);
      // Called for its side effect only — it sets `scopeMissing`. See the note by the state above.
      await loadRows();
      setWelcome(w);
      setPages(p);
      setGroups(g.filter((x) => builtIns.indexOf(x.id) === -1));
      setLoadError(undefined);
    } catch (e) {
      setLoadError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { reloadAll(true).catch(() => undefined); }, []);

  // Refresh the selected page's live ACL when the selection changes.
  useEffect(() => {
    if (!target) { setLive(undefined); return; }
    loadLiveFor(pages, target).then(setLive).catch(() => setLive(undefined));
  }, [target, pages]);

  /**
   * Resolve "Go to Group Management" from Site Pages, never hardcoded — the client renames every
   * page at import, so a literal `Group-Management.aspx` would fail as a dead link. Same lookup,
   * same `resolveLink`, as Site Access and the retired Folder Access signpost.
   *
   * A failed read simply leaves the link out; the banner's own text still names the page.
   */
  useEffect(() => {
    const link = CARDS
      .filter((c) => c.key === "access")[0]
      ?.links.filter((l) => l.key === "groups")[0];
    if (!link) return;
    readSitePages(context, siteUrl)
      .then((pages_) => {
        const t = resolveLink(link, pages_);
        if (t.state !== "missing") setGroupManagementHref(t.url);
      })
      .catch(() => undefined);
  }, []);

  return (
    <div style={s.wrap}>
      {/* ⚠ BOTH SHORTENED (client's mockup, 2026-09-03). The intro used to explain the mechanism
          ("Until a page is restricted, everyone with access to the site can open it"); the client's
          copy just names what the screen shows. The warn box dropped the "which groups... is now set
          entirely by Folder Reconciliation from the CRS Group Map" explanation for the same reason —
          the read-only note plus the one link is what an admin acts on; the mechanism is prose. */}
      <p style={s.intro}>View page access permissions.</p>

      <div style={s.warnBox}>
        <strong>This page is read-only.</strong> To change who can access a page,{" "}
        {groupManagementHref !== undefined
          ? <a href={groupManagementHref} style={{ fontWeight: 600 }}>go to Group Management</a>
          : <strong>go to the Group Management page</strong>}.
      </div>

      {scopeMissing && (
        <div style={s.dangerBox}>
          The <strong>Scope</strong> and <strong>Target</strong> columns are missing from{" "}
          <strong>{groupMapList()}</strong>.
        </div>
      )}
      {loadError && <div style={s.dangerBox}>Could not load: {loadError}</div>}

      <div style={s.card}>
        <label style={s.label}>Page</label>
        {loading ? (
          <div style={s.no}>Loading&hellip;</div>
        ) : (
          <>
            <select
              style={s.select}
              value={target}
              onChange={(e) => { setTarget(e.target.value); setFilter(""); }}
            >
              <option value="">Select a page&hellip;</option>
              {selectablePages.map((p) => (
                <option key={p.itemId} value={p.fileName}>
                  {p.title} ({p.fileName}){p.unique ? " — restricted" : ""}
                </option>
              ))}
            </select>
            {/* ⚠ "The site home page is not listed..." REMOVED (client's mockup, 2026-09-03), and
                `excludedCount` went with it — its only reader. Removing rather than parking: it is a
                one-line derivation, trivially re-added if this note is ever wanted back. */}
            <div style={s.hint}>
              &ldquo;restricted&rdquo; means the page already has its own permissions.
            </div>
          </>
        )}
      </div>

      {page && (
        <div style={s.card}>
          <div style={s.head}>{page.title} <span style={s.mono}>{page.fileName}</span></div>
          {!page.unique ? (
            <div style={s.openBox}>
              <strong>Open to everyone with site access.</strong> Nothing here restricts it.
            </div>
          ) : (
            <div style={s.openBox}>
              <strong>Restricted.</strong> Only the groups below, plus site owners, can open this
              page.
            </div>
          )}

          {/* Who this page is FOR, still worth stating even though the table below now shows
              actual grants rather than a policy guess — it is the context that explains why the
              table holds the groups it does. No escape hatch any more: the old "Show all N groups"
              button existed to override a policy-derived guess, and there is no guess left to
              override — the table already shows every group truly granted, whatever its name. */}
          <div style={policy.adminOnly ? s.adminBox : s.policyBox}>{policy.reason}</div>

          {!page.unique ? (
            <p style={s.hint}>
              Since this page is open to everyone with site access, no specific group holds an
              explicit grant on it — there is nothing to list here.
            </p>
          ) : !liveKnown ? (
            <div style={s.warnBox}>
              Could not read this page&rsquo;s current permissions, so it is not possible to say
              which groups currently have access.
            </div>
          ) : (
            <>
              {/* ⚠ THE "Only groups that currently hold a real grant..." EXPLAINER IS GONE (client's
                  mockup, 2026-09-03). The RULE it described is unchanged — `granted` still filters on
                  `hasRealGrant`, so a bare Limited Access entry is still never listed — only the
                  sentence explaining why is gone from the screen. */}

              {/* Search box — case-insensitive substring match, same rule as Group Management's and
                  Approval Library Access's own filter. This list can run to dozens of rows. */}
              {granted.length > 0 && (
                <div style={s.filterRow}>
                  <input
                    style={s.input}
                    value={filter}
                    placeholder="Filter by name…"
                    onChange={(e) => setFilter(e.target.value)}
                  />
                  {filter.trim() && (
                    <span style={s.hint}>
                      {visible.length} of {granted.length}
                    </span>
                  )}
                </div>
              )}

              {granted.length === 0 ? (
                <div style={s.no}>No group currently holds an explicit grant on this page.</div>
              ) : (
                /* Inner scroll: nothing in this table is absolutely positioned — the member popup
                   below is `position: fixed`, which escapes a scrolling ancestor's clipping rather
                   than being clipped by it, unlike Group Management's people picker or the upload
                   form's info tooltips. Checked before capping the height. */
                <div style={s.scroller}>
                  <table style={s.table}>
                    <thead>
                      <tr>
                        <th style={s.th}>Group</th>
                        <th style={s.th}>People</th>
                      </tr>
                    </thead>
                    <tbody>
                      {visible.length === 0 ? (
                        <tr>
                          <td style={s.td} colSpan={2}>
                            <span style={s.no}>No group matches that filter.</span>
                          </td>
                        </tr>
                      ) : (
                        visible.map((g) => (
                          <tr key={g.id}>
                            <td style={s.td}>{g.title}</td>
                            {/* Member COUNT, click to see who — `MemberCountWithPopup` reuses
                                the same `members` map already read from `useGroupMembers`, so
                                there is still only one fetch per group. */}
                            <td style={s.td}>
                              <MemberCountWithPopup group={g} members={members} />
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}
        </div>
      )}

    </div>
  );
}
