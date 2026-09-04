// Staging Library Access tab — which groups may OPEN the Staging library.
//
// The problem it explains: folder grants alone are not enough to reach a library. SharePoint
// auto-grants *Limited Access* up the parent chain when a child folder is assigned, which
// lets a DIRECT folder URL through but confers no View Items on the list — so an uploader
// clicking "Staging" in the left nav gets Access Denied. Library-level Read is the missing
// piece.
//
// It does NOT widen what they can see. Every folder reconciliation provisions has broken
// inheritance and its own ACL, so library Read does not flow into the segment or department
// folders — those stay security-trimmed. The uploader gets into the library and still cannot
// see GHO or GHR.
//
// ⚠ READ-ONLY SINCE 2026-09-02 (spec `2026-09-02-access-pages-read-only-design.md`). Client:
// "the group assignment is all done by Folder recon, and how a group is have access to
// anything is also done by folder recon, so I am see there is no point for all this pages to
// allow client to manually strip or assign group to a page or library." Bulk provisioning
// creates groups and writes their Group Map rows; reconciliation grants every library, folder
// and page ACL from those rows. A manual grant/revoke here either duplicates that or drifts
// from it and gets silently fought (or ignored) on the next reconciliation run.
//
// This page now only ever READS: the live ACL, the Group Map rows, and the site's group list —
// to show where they agree and where they do not. Narrowing access is done by deleting the
// group in Group Management and letting reconciliation re-grant from what remains; widening it
// is done by mapping/creating the right group there. Both routes keep the Group Map as the one
// source of truth this whole model depends on.
//
// See docs/superpowers/specs/2026-08-03-access-scope-mapping-design.md (the original read model)
// and docs/superpowers/specs/2026-09-02-access-pages-read-only-design.md (why the writes left).
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  GroupMapRole,
  roleFromGroupName,
  STAGING_FACING_ROLES,
  siteEntryGroupTitle,
  normalizeRoleValue,
} from "../../../shared/groupMapModel";
import { fetchAllSiteGroups, SpGroup } from "../../../shared/spGroups";
import { cachedListTitle, LIST_SUFFIX, libApiTitle } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { MemberCountWithPopup, useGroupMembers } from "./accessMemberUi";
import { CARDS, resolveLink } from "../../../shared/adminPages";
import { readSitePages } from "../../../shared/backToSettings";

type Props = { context: WebPartContext; siteUrl: string; library: string };

// Resolved, not hardcoded: the client renames this to "CRS Group Map" at import (confirmed on
// their site 2026-08-05). Read from the cache primed in reload(), which always runs on mount
// before any request this component makes.
const groupMapList = (): string => cachedListTitle(LIST_SUFFIX.groupMap);

/** A Group Map row granting library entry. */
type EntryRow = { itemId: number; groupId: string; groupName: string; role: GroupMapRole };

/** What the library's ACL actually says right now, per principal. */
type LiveGrant = { principalId: number; title: string; levels: string[] };

const s: Record<string, React.CSSProperties> = {
  wrap:     { fontSize: 13, color: "#242424", lineHeight: 1.5 },
  intro:    { fontSize: 13, color: "#444", margin: "0 0 16px" },
  card:     { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  head:     { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  table:    { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  th:       { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:       { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "middle" },
  yes:      { color: "#0f6c3f", fontWeight: 600 },
  no:       { color: "#8a8886" },
  // Filter box, matching Group Management's "Filter by name…" input (GroupManager.tsx `s.input`).
  input:    { width: "100%", maxWidth: 320, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  filterRow:{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 },
  warnBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  dangerBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  readOnlyBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #cfd8e3", background: "#f4f7fb", borderRadius: 4, fontSize: 12, color: "#2b3f56", lineHeight: 1.5 },

  // ── The designer's header block (2026-08-15) ──────────────────────────────────
  // Explanation on the left, callouts on the right. `auto-fit` rather than a fixed pair of columns,
  // so it stacks on a narrow window instead of squeezing the copy to one word a line.
  headGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 320px), 1fr))", gap: 20, alignItems: "start", marginBottom: 20 },
  headLeft: { display: "flex", gap: 14, alignItems: "flex-start" },
  headIcon: { flexShrink: 0, width: 44, height: 44, borderRadius: "50%", background: "#eef7f1", display: "flex", alignItems: "center", justifyContent: "center", color: "#0f6c3f" },
  headTitle:{ fontSize: 15, fontWeight: 600, margin: "0 0 6px", color: "#242424" },
  noteCard: { display: "flex", gap: 10, alignItems: "flex-start", padding: "12px 14px", borderRadius: 6, fontSize: 12, lineHeight: 1.5, marginBottom: 10 },
  noteOk:   { border: "1px solid #b7dcc4", background: "#f3faf5", color: "#1c4d33" },
  // RED, not amber (client's mockup, 2026-09-03: "Change yellow to RED") — a failed permissions read
  // is the state where the "Current access" column cannot be trusted, and this is the one warning on
  // this page that says so; red matches this codebase's established danger palette (`#a4262c` etc).
  noteWarn: { border: "1px solid #f1b0b3", background: "#fdf3f4", color: "#a4262c" },
  noteTitle:{ fontWeight: 600, margin: "0 0 3px", fontSize: 12.5 },
  hint:     { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
  mono:     { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  folderOnly:{ color: "#8a8886", fontStyle: "italic" },
};

export default function StagingAccess({ context, siteUrl, library }: Props): React.ReactElement {
  const [groups, setGroups]   = useState<SpGroup[]>([]);
  const [rows, setRows]       = useState<EntryRow[]>([]);
  const [live, setLive]       = useState<LiveGrant[] | undefined>(undefined);
  const [loading, setLoading] = useState(true);
  // Filter box over the group list — it can run to 60+ groups per segment. Case-insensitive
  // substring match against the group's title, same rule as Group Management's own filter.
  const [filter, setFilter] = useState("");
  const [scopeMissing, setScopeMissing] = useState(false);
  // Why the ACL read failed, if it did. Kept so the banner can name a status rather than only saying
  // "could not read" — see loadLive.
  const [liveError, setLiveError] = useState<string | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  // Where "Group Management" lives, resolved from Site Pages rather than hardcoded — this client
  // renames every page at import. `undefined` while unresolved (or on a failed read), in which case
  // the banner still names the page without linking it.
  const [groupsHref, setGroupsHref] = useState<string | undefined>(undefined);

  /**
   * ⚠ REAL BUG, FOUND LIVE 2026-09-02: this used to be `rows.map((r) => Number(r.groupId))` — the
   * Group Map's LIBRARY-scope entries — while the table below renders `candidates`, a completely
   * DIFFERENT set derived from `groups` (every site group) filtered by NAME PATTERN
   * (`roleFromGroupName`). The two sets do not have to overlap, and on this site they evidently do
   * not: `useGroupMembers` was fetching members for groups nobody was looking at, so every row the
   * table actually displays sat on "loading…" forever — not because the fetch was slow (it was
   * fixed for that, twice), but because it was never asked about the right groups at all.
   *
   * Fixed by feeding the hook the SAME ids the table shows — `candidates`, moved up from the
   * "Derived view" section below so it exists before this call. `groups.filter(...)` runs once
   * either way; only its position changed.
   */
  const candidates = groups
    .filter((g) => g.title.trim().toLowerCase() !== siteEntryGroupTitle().toLowerCase())
    .filter((g) => STAGING_FACING_ROLES.indexOf(roleFromGroupName(g.title)) !== -1)
    .sort((a, b) => a.title.localeCompare(b.title));

  const { members } = useGroupMembers(context.spHttpClient, siteUrl, candidates.map((g) => g.id));

  const GET = { Accept: "application/json;odata=nometadata" };
  /**
   * The library's API URL. **A FUNCTION, deliberately — never a render-time const.**
   *
   * `library` is the LOGICAL key ("Staging"), which is what Group Map rows store and what this
   * component filters. A URL needs the live TITLE, and on this site that is "Approval Document",
   * so it must go through `libApiTitle` (gotcha #12: translate at the API boundary, never in stored
   * data).
   *
   * But `libApiTitle` reads a module-level cache primed by `primeNames`, and that has NOT happened on
   * the first render — where it still answers with the legacy `Staging`. As a const, the value
   * captured by that first render was the one `loadLive` used from the mount effect, so the read
   * 404'd; later renders computed the right URL but nothing re-read the ACL, because the effect only
   * re-runs on `[library]` and that never changes.
   *
   * Called at request time it always sees the primed cache. Do not turn this back into a const.
   */
  const listBase = (): string =>
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(library))}')`;
  // What the client is told the library is called. `library` is an internal key and must not surface
  // in a sentence — the page heading already resolves the live title for the same reason. Safe as a
  // render value: it is only read after a state change has re-rendered with the primed cache.
  const libLabel = libApiTitle(library);

  /**
   * Library-entry rows from the Group Map.
   *
   * Scope/Target are newer than this list, and a $select naming a column that does not exist
   * fails the WHOLE request with HTTP 400 — not a null, not a missing key (CLAUDE.md #11). So
   * the absence of the columns is detected by the request being rejected, and reported rather
   * than degraded: without them no library row is readable, and a tab that silently showed
   * nothing would be blamed on the groups.
   */
  const loadRows = async (): Promise<EntryRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(groupMapList())}')/items` +
        `?$select=Id,GroupId,GroupName,Role,Scope,Target&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) {
      setScopeMissing(true);
      return [];
    }
    setScopeMissing(false);
    const data = await res.json();
    return ((data.value ?? []) as Array<{ Id: number; GroupId?: string; GroupName?: string; Role?: string; Scope?: string; Target?: string }>)
      .filter((r) => (r.Scope ?? "").trim().toLowerCase() === "library")
      .filter((r) => (r.Target ?? "").trim().toLowerCase() === library.toLowerCase())
      .map((r) => ({
        itemId: r.Id,
        groupId: (r.GroupId ?? "").trim(),
        groupName: r.GroupName ?? "",
        role: normalizeRoleValue(r.Role ?? "") as GroupMapRole,
      }));
  };

  /**
   * The library's live ACL.
   *
   * Read separately from the Group Map because the two can disagree, and the disagreement is
   * the thing worth showing: a row with no live grant means reconciliation has not run yet, and
   * a live grant with no row means someone granted it by hand in SharePoint. A tab that showed
   * only its own rows would report access that does not exist, or hide access that does.
   */
  const loadLive = async (): Promise<LiveGrant[] | undefined> => {
    const url =
      `${listBase()}/roleassignments?$expand=Member,RoleDefinitionBindings` +
      `&$select=PrincipalId,Member/Title,RoleDefinitionBindings/Name`;
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) {
      // RECORD THE STATUS. This failed silently for weeks behind a banner saying only "could not
      // read", and a 404 (wrong library title) and a 403 (no Enumerate Permissions) are the same
      // sentence with completely different fixes. Gotcha #9 states the rule and this is the second
      // place to have learned it: log the actual status before assuming a naming or data problem.
      setLiveError(`HTTP ${res.status}`);
      console.warn(`[ApprovalLibraryAccess] could not read the library ACL: HTTP ${res.status} — ${url}`);
      return undefined;
    }
    setLiveError(undefined);
    const data = await res.json();
    return ((data.value ?? []) as Array<{ PrincipalId?: number; Member?: { Title?: string }; RoleDefinitionBindings?: Array<{ Name?: string }> }>)
      .map((ra) => ({
        principalId: ra.PrincipalId ?? 0,
        title: ra.Member?.Title ?? "",
        levels: (ra.RoleDefinitionBindings ?? []).map((b) => b.Name ?? "").filter(Boolean),
      }));
  };

  /** Re-read everything: the site's groups, this library's Group Map rows, and its live ACL. */
  const reload = async (): Promise<void> => {
    setLoading(true);
    try {
      // Before any list read: resolves "CRS Group Map" vs "DMS Group Map" once per session.
      await primeNames(context.spHttpClient, siteUrl);
      const g = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
      const r = await loadRows();
      const l = await loadLive();
      setGroups(g);
      setRows(r);
      setLive(l);
      setLoadError(undefined);
    } catch (e) {
      setLoadError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { reload().catch(() => undefined); }, [library]);

  /**
   * Where "Group Management" lives, resolved from Site Pages the same way the CRS Settings landing
   * page and the retired Folder Access signpost do — never hardcoded, because this client renames
   * every page at import. A failed read (or no match) just leaves the banner unlinked; it still
   * names the page.
   */
  useEffect(() => {
    const link = CARDS
      .filter((c) => c.key === "access")[0]
      ?.links.filter((l) => l.key === "groups")[0];
    if (!link) return;
    readSitePages(context, siteUrl)
      .then((pages) => {
        const t = resolveLink(link, pages);
        if (t.state !== "missing") setGroupsHref(t.url);
      })
      .catch(() => undefined);
  }, [siteUrl]);

  // ── Derived view ───────────────────────────────────────────────────────────
  const rowByGroupId = new Map<string, EntryRow>();
  for (const r of rows) rowByGroupId.set(r.groupId, r);

  const liveByPid = new Map<number, LiveGrant>();
  for (const l of live ?? []) liveByPid.set(l.principalId, l);

  // `candidates` (groups whose NAME says they act on Staging, site-entry group excluded) now
  // computed earlier, before the `useGroupMembers` call — see the comment there for why.

  // Filtered view for display only — `candidates` above stays the full set everything else (the
  // empty-state message, counts) reasons about.
  const visible = candidates.filter(
    (g) => !filter.trim() || g.title.toLowerCase().indexOf(filter.trim().toLowerCase()) !== -1,
  );

  // Rows pointing at a group that is NOT Staging-facing. Shown rather than filtered out: a
  // hand-written row granting a viewer group entry to Staging is exactly what this tab exists
  // to flag, and hiding it would leave it in force and invisible.
  const unexpected = rows.filter((r) => {
    const g = groups.find((x) => String(x.id) === r.groupId);
    const role = roleFromGroupName(g?.title ?? r.groupName);
    return STAGING_FACING_ROLES.indexOf(role) === -1;
  });

  // Live grants with no row behind them, ignoring the ones that must be there. Surfaced so a
  // hand-made grant in SharePoint cannot sit unexplained in the library ACL.
  const unmanaged = (live ?? []).filter((l) => {
    if (rowByGroupId.has(String(l.principalId))) return false;
    if (l.title.trim().toLowerCase() === siteEntryGroupTitle().toLowerCase()) return false;
    // Owners/admins and anyone with Full Control are meant to be here.
    if (l.levels.indexOf("Full Control") !== -1) return false;
    // LIMITED ACCESS IS NOT LIBRARY ACCESS, and listing it was this panel's first bug.
    //
    // SharePoint auto-grants Limited Access on a list to EVERY principal holding a grant on
    // any folder inside it — so every uploader group, and every individual user with a folder
    // grant, appears here automatically. It confers no View Items: it exists only so a direct
    // URL to the child can resolve. Reporting it as an unexplained direct grant listed the
    // entire user base as a finding and buried the one or two entries that matter.
    if (l.levels.every((n) => n === "Limited Access")) return false;
    return true;
  });

  return (
    <div style={s.wrap}>
      {/* Header block to the designer's layout (2026-08-15): the explanation on the left, the two
          callouts on the right.

          The green card states a permanent fact and is always shown. The amber one is a STATE and is
          shown only when the ACL read actually failed — the mock drew it as a static sibling, which
          would tell an admin on every visit that permissions cannot be verified, and train them to
          ignore the message on the day it is true.

          Its copy also drops the mock's "until permissions are synchronized": there is no sync, and
          nothing improves by waiting. The read fails for two reasons with OPPOSITE fixes — a 404 is
          the wrong library title, a 403 is not holding Full Control — so the status stays named. */}
      <div style={s.headGrid}>
        <div style={s.headLeft}>
          <div style={s.headIcon} aria-hidden="true">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
              <path d="M16 20v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
              <circle cx="9" cy="7" r="4" />
              <path d="M22 20v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" />
            </svg>
          </div>
          <div>
            <p style={s.headTitle}>Who can access the {libLabel} library?</p>
            <p style={{ ...s.intro, margin: 0 }}>
              Which groups may <strong>open the {libLabel} library</strong>. Folder permissions alone
              are not enough — without an entry here, an uploader clicking <strong>{libLabel}</strong>{" "}
              in the left navigation gets <em>Access Denied</em>, even though a direct link to their
              own folder works.
            </p>
          </div>
        </div>

        <div>
          <div style={{ ...s.noteCard, ...s.noteOk }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, marginTop: 1 }} aria-hidden="true">
              <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
              <path d="M9 12l2 2 4-4" />
            </svg>
            <div>
              <p style={s.noteTitle}>Library access does not expand document access</p>
              People still see only the folders and documents they already have permission to open.
              Each folder keeps its own permissions, so other units stay hidden — this only lets them
              reach the library.
            </div>
          </div>

          {live === undefined && !loading && (
            <div style={{ ...s.noteCard, ...s.noteWarn }}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, marginTop: 1 }} aria-hidden="true">
                <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
                <path d="M12 9v4M12 17h.01" />
              </svg>
              <div>
                <p style={s.noteTitle}>
                  Unable to verify current permissions{liveError ? ` (${liveError})` : ""}
                </p>
                We could not read {libLabel}&rsquo;s current permissions, so the{" "}
                <strong>Current access</strong> column reads <em>unknown</em>. The list below may not
                reflect what the library actually says right now.
                {liveError === "HTTP 404" && (
                  <div style={{ marginTop: 6 }}>
                    A <strong>404</strong> means no library by that name — check it is still titled{" "}
                    <strong>{libLabel}</strong>.
                  </div>
                )}
                {(liveError === "HTTP 403" || liveError === "HTTP 401") && (
                  <div style={{ marginTop: 6 }}>
                    A <strong>{liveError === "HTTP 403" ? "403" : "401"}</strong> means your account
                    cannot read this library&rsquo;s permissions. Reading them needs Full Control on
                    the library.
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Read-only since 2026-09-02: reconciliation is what grants and revokes every one of these,
          from the Group Map. Adding or removing access here would either duplicate that or drift
          from it and get silently overwritten (or ignored) on the next run. */}
      <div style={s.readOnlyBox}>
        <strong>This page is read-only.</strong> To add or remove access, use{" "}
        {groupsHref
          ? <a href={groupsHref} style={{ fontWeight: 600 }}>Group Management</a>
          : <strong>Group Management</strong>}
        {" "}— create or map a group there to grant access, or delete the group there to take it away;
        Folder Reconciliation applies the result.
      </div>

      {scopeMissing && (
        <div style={s.dangerBox}>
          The <strong>Scope</strong> and <strong>Target</strong> columns are missing from{" "}
          <strong>{groupMapList()}</strong>, so a library mapping cannot be determined at all.
        </div>
      )}

      {loadError && (
        <div style={s.dangerBox}>Could not load: {loadError}</div>
      )}

      {unexpected.length > 0 && (
        <div style={s.dangerBox}>
          <strong>{unexpected.length} mapping(s) grant {libLabel} access to a group that is not an
          uploader, approver or Staging deleter.</strong> A viewer group here can read other
          people&rsquo;s unapproved documents.
          <div style={s.mono}>{unexpected.map((r) => r.groupName || r.groupId).join(", ")}</div>
        </div>
      )}

      {unmanaged.length > 0 && (
        <div style={s.warnBox}>
          <strong>{unmanaged.length} group(s) already hold permissions on {libLabel} without a
          mapping here</strong> — granted directly in SharePoint. They are not managed by
          reconciliation and it will not remove them.
          <div style={s.mono}>
            {unmanaged.map((l) => `${l.title || l.principalId} (${l.levels.join(", ") || "no level"})`).join(", ")}
          </div>
        </div>
      )}

      <div style={s.card}>
        <div style={s.head}>Uploader, approver and Staging-deleter groups</div>
        {loading ? (
          <div style={s.no}>Loading&hellip;</div>
        ) : candidates.length === 0 ? (
          <div style={s.no}>
            No uploader or approver groups found on this site. Create them on the{" "}
            <strong>Group Management</strong> page first — this tab only shows access for groups
            that already exist.
          </div>
        ) : (
          <>
            {/* ABOVE the table, not below it. This legend used to sit under the last row — on this
                site that is 17 rows down, so the one explanation was somewhere the client would
                never scroll to. Now says what "Current access" alone means — the "should have
                access" comparison it used to draw is gone with that column: on a page that can no
                longer act on a mismatch, showing what the Group Map says access "should" be beside
                what it actually is reads as an open discrepancy waiting to be fixed by hand, which
                this page can no longer do. */}
            <div style={{ ...s.hint, margin: "0 0 10px" }}>
              <strong>Current access</strong> is a live read of what the {libLabel} library&rsquo;s
              permissions actually say right now, for each uploader, approver or Staging-deleter
              group. Add or remove a group on Group Management; Folder Reconciliation applies it here.
            </div>
            {/* Search box — case-insensitive substring match, same rule as Group Management's own
                "Filter by name…" filter. This list runs to 60+ rows per segment. */}
            <div style={s.filterRow}>
              <input
                style={s.input}
                value={filter}
                placeholder="Filter by name…"
                onChange={(e) => setFilter(e.target.value)}
              />
              {filter.trim() && (
                <span style={s.hint}>
                  {visible.length} of {candidates.length}
                </span>
              )}
            </div>
            {/* Inner scroll: `MemberCountWithPopup`'s modal is `position: fixed`, not absolute, so it
                is unaffected by this container's `overflow` — unlike Group Management's own list,
                whose member editor's people picker IS absolutely positioned and clipped by a scroll
                cap for exactly that reason. No such case exists here. */}
            <div style={{ maxHeight: "60vh", overflowY: "auto" }}>
              <table style={s.table}>
                <thead>
                  <tr>
                    <th style={s.th}>Group</th>
                    <th style={s.th}>Role</th>
                    <th style={s.th}>People</th>
                    <th style={s.th} title="What the library's permissions actually say right now.">
                      Current access
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {visible.length === 0 ? (
                    <tr>
                      <td style={s.td} colSpan={4}>
                        <span style={s.no}>No group matches that filter.</span>
                      </td>
                    </tr>
                  ) : (
                    visible.map((g) => {
                      const liveGrant = liveByPid.get(g.id);
                      // Limited Access alone is not library access — it is auto-granted because
                      // the group holds a folder inside. Saying "Limited Access" here invited
                      // exactly the wrong conclusion, so it is named for what it means.
                      const folderOnly =
                        liveGrant !== undefined && liveGrant.levels.every((n) => n === "Limited Access");
                      return (
                        <tr key={g.id}>
                          <td style={s.td}>{g.title}</td>
                          <td style={s.td}>{roleFromGroupName(g.title)}</td>
                          {/* A read-only count, click to see who — kept from the removed
                              per-person editor because knowing an "allowed" group is empty is worth
                              seeing on its own. Adding or removing a member is done on Group
                              Management now. */}
                          <td style={s.td}>
                            <MemberCountWithPopup group={g} members={members} />
                          </td>
                          <td style={s.td}>
                            {live === undefined
                              ? <span style={s.no}>unknown</span>
                              : liveGrant === undefined
                                ? <span style={s.no}>none</span>
                                : folderOnly
                                  ? <span style={s.folderOnly}>their folders only</span>
                                  : <span style={s.yes}>{liveGrant.levels.filter((n) => n !== "Limited Access").join(", ") || "granted"}</span>}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
