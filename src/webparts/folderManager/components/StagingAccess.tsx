// Staging Library Access tab — which groups may OPEN the Staging library.
//
// The problem it solves: folder grants alone are not enough to reach a library. SharePoint
// auto-grants *Limited Access* up the parent chain when a child folder is assigned, which
// lets a DIRECT folder URL through but confers no View Items on the list — so an uploader
// clicking "Staging" in the left nav gets Access Denied. Library-level Read is the missing
// piece, and until now it could only be added by hand-writing a Group Map row.
//
// It does NOT widen what they can see. Every folder reconciliation provisions has broken
// inheritance and its own ACL, so library Read does not flow into the segment or department
// folders — those stay security-trimmed. The uploader gets into the library and still cannot
// see GHO or GHR.
//
// Only UPL/APR/DELS groups are offered: a viewer-only group on Staging would be reading other
// people's unapproved drafts, which is the isolation rule the whole model rests on.
//
// See docs/superpowers/specs/2026-08-03-access-scope-mapping-design.md.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  GroupMapRole,
  roleFromGroupName,
  buildGroupMapRow,
  LIBRARY_ENTRY_ROLE,
  STAGING_FACING_ROLES,
  SITE_ENTRY_GROUP_NAME,
} from "../../../shared/groupMapModel";
import { fetchAllSiteGroups, SpGroup } from "../../../shared/spGroups";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

type Props = { context: WebPartContext; siteUrl: string; library: string };

// Resolved, not hardcoded: the client renames this to "CRS Group Map" at import (confirmed on
// their site 2026-08-05). Read from the cache primed in reload(), which always runs on mount
// before any write this component can make.
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
  addBtn:   { padding: "4px 12px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  delBtn:   { padding: "4px 12px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  off:      { padding: "4px 12px", fontSize: 12, background: "#c7c7c7", color: "#fff", border: "none", borderRadius: 4, cursor: "not-allowed" },
  yes:      { color: "#0f6c3f", fontWeight: 600 },
  no:       { color: "#8a8886" },
  warnBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  dangerBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  okBox:    { marginBottom: 16, padding: "10px 12px", border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 4, fontSize: 12, color: "#0f6c3f", lineHeight: 1.5 },
  hint:     { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
  toast:    { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 6, color: "#fff", fontSize: 13, zIndex: 50, boxShadow: "0 4px 14px rgba(0,0,0,.18)" },
  modalOverlay:{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 },
  modalBox: { background: "#fff", borderRadius: 8, width: "100%", maxWidth: 480, boxShadow: "0 8px 30px rgba(0,0,0,.25)" },
  modalHead:{ padding: "12px 16px", borderBottom: "1px solid #eee", fontWeight: 600, fontSize: 14 },
  modalBody:{ padding: "12px 16px" },
  modalFoot:{ padding: "10px 16px", borderTop: "1px solid #eee", display: "flex", justifyContent: "flex-end", gap: 8 },
  ghost:    { padding: "5px 14px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  mono:     { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  // Selection bar. Sticky so it stays reachable with twelve-plus groups listed — the whole
  // point of bulk select is not having to scroll back to a button.
  bulkBar:  { position: "sticky", top: 0, zIndex: 10, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", padding: "10px 12px", marginBottom: 12, border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 6 },
  bulkCount:{ fontSize: 12, fontWeight: 600, color: "#0f6c3f" },
  check:    { width: 15, height: 15, cursor: "pointer", verticalAlign: "middle" },
  checkCell:{ padding: "6px 8px", borderBottom: "1px solid #f0f0f0", width: 28, verticalAlign: "middle" },
  spinner:  { display: "inline-block", width: 11, height: 11, marginRight: 6, border: "2px solid rgba(255,255,255,.45)", borderTopColor: "#fff", borderRadius: "50%", verticalAlign: "-1px", animation: "sa-spin .6s linear infinite" },
  spinnerDark:{ display: "inline-block", width: 11, height: 11, marginRight: 6, border: "2px solid rgba(164,38,44,.3)", borderTopColor: "#a4262c", borderRadius: "50%", verticalAlign: "-1px", animation: "sa-spin .6s linear infinite" },
  working:  { padding: "4px 12px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "wait", opacity: .85 },
  folderOnly:{ color: "#8a8886", fontStyle: "italic" },
};

// Keyframes cannot be expressed in an inline style object, and this component deliberately
// carries no .module.scss — one <style> element keeps the spinner self-contained.
const SPIN_KEYFRAMES = "@keyframes sa-spin{to{transform:rotate(360deg)}}";

export default function StagingAccess({ context, siteUrl, library }: Props): React.ReactElement {
  const [groups, setGroups]   = useState<SpGroup[]>([]);
  const [rows, setRows]       = useState<EntryRow[]>([]);
  const [live, setLive]       = useState<LiveGrant[] | undefined>(undefined);
  const [busy, setBusy]       = useState(false);
  const [loading, setLoading] = useState(true);
  // Group ids with a write in flight. A SET rather than a single boolean so one row's spinner
  // does not appear on every button, and so a bulk run can show progress per row.
  const [pending, setPending] = useState<number[]>([]);
  // Ticked rows, for the bulk Allow. Ids, not indexes — the table re-sorts on reload.
  const [selected, setSelected] = useState<number[]>([]);
  const [bulk, setBulk] = useState<{ done: number; total: number } | undefined>(undefined);
  // Which bulk action is running, so the progress count appears on the button that was
  // pressed rather than on both.
  const [pendingAction, setPendingAction] = useState<"allow" | "remove" | undefined>(undefined);
  const [scopeMissing, setScopeMissing] = useState(false);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [toast, setToast]     = useState<{ message: string; error: boolean } | undefined>(undefined);
  // A LIST, so the single-row Remove and the bulk Remove share one confirm dialog. Two
  // dialogs saying almost the same thing is how their wording drifts apart, and the
  // destructive one is the worst place for that.
  const [confirmRemove, setConfirmRemove] = useState<EntryRow[] | undefined>(undefined);

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 6000);
  };

  const GET = { Accept: "application/json;odata=nometadata" };
  const listBase = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(library)}')`;

  /**
   * Library-entry rows from the Group Map.
   *
   * Scope/Target are newer than this list, and a $select naming a column that does not exist
   * fails the WHOLE request with HTTP 400 — not a null, not a missing key (CLAUDE.md #11). So
   * the absence of the columns is detected by the request being rejected, and reported rather
   * than degraded: without them no library row can be written at all, and a tab that silently
   * did nothing would be blamed on the groups.
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
        role: (r.Role ?? "").toUpperCase() as GroupMapRole,
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
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${listBase}/roleassignments?$expand=Member,RoleDefinitionBindings` +
        `&$select=PrincipalId,Member/Title,RoleDefinitionBindings/Name`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return undefined;
    const data = await res.json();
    return ((data.value ?? []) as Array<{ PrincipalId?: number; Member?: { Title?: string }; RoleDefinitionBindings?: Array<{ Name?: string }> }>)
      .map((ra) => ({
        principalId: ra.PrincipalId ?? 0,
        title: ra.Member?.Title ?? "",
        levels: (ra.RoleDefinitionBindings ?? []).map((b) => b.Name ?? "").filter(Boolean),
      }));
  };

  /**
   * Re-read everything.
   *
   * `firstLoad` decides whether the table is replaced by "Loading…". After an Allow it must
   * NOT be: blanking the whole table for a round-trip reads as a page refresh, loses the
   * reader's place, and makes one row's change look like the whole list rebuilt.
   */
  const reload = async (firstLoad = false): Promise<void> => {
    if (firstLoad) setLoading(true);
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

  useEffect(() => { reload(true).catch(() => undefined); }, [library]);

  /** Grant Read on the list immediately, so the tab reflects reality without a recon run. */
  const grantLive = async (principalId: number): Promise<void> => {
    const defs: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roledefinitions?$select=Id,Name`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!defs.ok) throw new Error(`roledefinitions HTTP ${defs.status}`);
    const dj = await defs.json();
    const readId = ((dj.value ?? []) as Array<{ Id: number; Name: string }>)
      .find((d) => d.Name === "Read")?.Id;
    if (readId === undefined) throw new Error('no "Read" permission level on this site');
    const res = await context.spHttpClient.post(
      `${listBase}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${readId})`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) throw new Error(`addroleassignment HTTP ${res.status}`);
  };

  /**
   * Remove the live grant as well as the row.
   *
   * Reconciliation only ever ADDS, so deleting the row alone would leave the access in place
   * while the tab showed none — the "why do people still have access" complaint. Removing both
   * is the only version of this button that matches what it says.
   */
  const revokeLive = async (principalId: number): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${listBase}/roleassignments/removeroleassignment(principalid=${principalId})`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) throw new Error(`removeroleassignment HTTP ${res.status}`);
  };

  const postRow = async (groupId: string, groupName: string): Promise<void> => {
    const row = buildGroupMapRow({
      groupId,
      groupName,
      role: LIBRARY_ENTRY_ROLE,
      scope: "Library",
      target: library,
    });
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(groupMapList())}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { ...GET, "Content-Type": "application/json;odata=nometadata" },
        body: JSON.stringify({ Title: groupName || groupId, ...row }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  const deleteRow = async (itemId: number): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(groupMapList())}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      { headers: { ...GET, "IF-MATCH": "*", "X-HTTP-Method": "DELETE" } },
    );
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
  };

  /**
   * Grant one group entry. Returns true on success so the bulk runner can count.
   *
   * Does NOT reload or toast — the caller decides, so a bulk run produces one refresh and one
   * summary rather than N of each.
   */
  const allowOne = async (g: SpGroup): Promise<boolean> => {
    setPending((p) => [...p, g.id]);
    try {
      await postRow(String(g.id), g.title);
      // Row first, then the grant. If the grant fails the row still stands and the next
      // reconciliation applies it; the reverse order could grant access with nothing
      // recording why, which is the state nobody can audit.
      try { await grantLive(g.id); } catch { return false; }
      return true;
    } finally {
      setPending((p) => p.filter((id) => id !== g.id));
    }
  };

  const onAdd = async (g: SpGroup): Promise<void> => {
    setBusy(true);
    try {
      const granted = await allowOne(g);
      await reload();
      showToast(
        granted
          ? `${g.title} can now open ${library}.`
          : `${g.title} mapped, but the permission could not be applied now — run Folder Reconciliation.`,
        !granted,
      );
    } catch (e) {
      showToast(`Could not add ${g.title}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /**
   * Allow every ticked group.
   *
   * SEQUENTIAL, deliberately. Each group costs two writes, and firing twelve in parallel is
   * how reconciliation earned its throttle-retry logic — a 429 here would half-apply the
   * selection with no record of which half. Slower and completely determinate.
   */
  const onAllowSelected = async (targets: SpGroup[]): Promise<void> => {
    if (targets.length === 0) return;
    setBusy(true);
    setPendingAction("allow");
    setBulk({ done: 0, total: targets.length });
    let ok = 0;
    const failed: string[] = [];
    for (let i = 0; i < targets.length; i++) {
      const g = targets[i];
      try {
        if (await allowOne(g)) ok++; else failed.push(g.title);
      } catch {
        failed.push(g.title);
      }
      setBulk({ done: i + 1, total: targets.length });
    }
    setBulk(undefined);
    setPendingAction(undefined);
    setSelected([]);
    await reload();
    setBusy(false);
    showToast(
      failed.length === 0
        ? `${ok} group(s) can now open ${library}.`
        : `${ok} granted, ${failed.length} failed: ${failed.join(", ")}`,
      failed.length > 0,
    );
  };

  /**
   * Revoke one group's entry. Returns true only if BOTH the row and the live permission went.
   *
   * A half-success is reported as a failure on purpose: the mapping being gone while the
   * permission remains is precisely the state that produces "why do these people still have
   * access", and calling it success is what would hide it.
   */
  const removeOne = async (row: EntryRow): Promise<boolean> => {
    const pid = Number(row.groupId);
    setPending((p) => (pid > 0 ? [...p, pid] : p));
    try {
      await deleteRow(row.itemId);
      if (pid > 0) { try { await revokeLive(pid); } catch { return false; } }
      return true;
    } finally {
      setPending((p) => p.filter((id) => id !== pid));
    }
  };

  /** Revoke one or many. Sequential, for the same throttling reason as the bulk Allow. */
  const onRemoveMany = async (targets: EntryRow[]): Promise<void> => {
    if (targets.length === 0) return;
    setBusy(true);
    setConfirmRemove(undefined);
    setPendingAction("remove");
    setBulk(targets.length > 1 ? { done: 0, total: targets.length } : undefined);
    let ok = 0;
    const failed: string[] = [];
    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      try {
        if (await removeOne(t)) ok++; else failed.push(t.groupName || t.groupId);
      } catch {
        failed.push(t.groupName || t.groupId);
      }
      if (targets.length > 1) setBulk({ done: i + 1, total: targets.length });
    }
    setBulk(undefined);
    setPendingAction(undefined);
    setSelected([]);
    await reload();
    setBusy(false);
    if (failed.length === 0) {
      showToast(
        targets.length === 1
          ? `${targets[0].groupName || targets[0].groupId} can no longer open ${library}.`
          : `${ok} group(s) can no longer open ${library}.`,
        false,
      );
    } else {
      showToast(
        `${ok} removed. Could not fully revoke ${failed.length}: ${failed.join(", ")} — check the library's permissions.`,
        true,
      );
    }
  };

  // ── Derived view ───────────────────────────────────────────────────────────
  const rowByGroupId = new Map<string, EntryRow>();
  for (const r of rows) rowByGroupId.set(r.groupId, r);

  const liveByPid = new Map<number, LiveGrant>();
  for (const l of live ?? []) liveByPid.set(l.principalId, l);

  // Candidates: groups whose NAME says they act on Staging. The site-entry group is excluded
  // explicitly — it exists to let people onto the site, and putting it here would give every
  // member of the site access to Staging, which is the one thing the model forbids.
  const candidates = groups
    .filter((g) => g.title.trim().toLowerCase() !== SITE_ENTRY_GROUP_NAME.toLowerCase())
    .filter((g) => STAGING_FACING_ROLES.indexOf(roleFromGroupName(g.title)) !== -1)
    .sort((a, b) => a.title.localeCompare(b.title));

  // Every row is now selectable, and the SELECTION decides which action applies to it rather
  // than the tick meaning different things in different rows. A group either has access or it
  // does not, so the two subsets never overlap and no row can be both allowed and removed.
  const allowable = candidates.filter((g) => !rowByGroupId.has(String(g.id)));
  const removable = candidates.filter((g) => rowByGroupId.has(String(g.id)));

  const pickedAllowable = allowable.filter((g) => selected.indexOf(g.id) !== -1);
  const pickedRemovable = removable
    .filter((g) => selected.indexOf(g.id) !== -1)
    .map((g) => rowByGroupId.get(String(g.id)) as EntryRow);

  // Rows pointing at a group that is NOT Staging-facing. Shown rather than filtered out: a
  // hand-written row granting a viewer group entry to Staging is exactly what this tab exists
  // to prevent, and hiding it would leave it in force and invisible.
  const unexpected = rows.filter((r) => {
    const g = groups.find((x) => String(x.id) === r.groupId);
    const role = roleFromGroupName(g?.title ?? r.groupName);
    return STAGING_FACING_ROLES.indexOf(role) === -1;
  });

  // Live grants with no row behind them, ignoring the ones that must be there. Surfaced so a
  // hand-made grant in SharePoint cannot sit unexplained in the library ACL.
  const unmanaged = (live ?? []).filter((l) => {
    if (rowByGroupId.has(String(l.principalId))) return false;
    if (l.title.trim().toLowerCase() === SITE_ENTRY_GROUP_NAME.toLowerCase()) return false;
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
      <style>{SPIN_KEYFRAMES}</style>
      <p style={s.intro}>
        Which groups may <strong>open the {library} library</strong>. Folder permissions alone are
        not enough — without an entry here, an uploader clicking <strong>{library}</strong> in the
        left navigation gets <em>Access Denied</em>, even though a direct link to their own folder
        works.
      </p>

      <div style={s.okBox}>
        This does <strong>not</strong> let them see other units. Each folder has its own
        permissions, so the segment and department folders stay hidden — they simply reach the
        library, and see only what they already had access to.
      </div>

      {scopeMissing && (
        <div style={s.dangerBox}>
          The <strong>Scope</strong> and <strong>Target</strong> columns are missing from{" "}
          <strong>{groupMapList()}</strong>. Add them (both single line of text) before using this
          tab — without them a library mapping cannot be saved at all.
        </div>
      )}

      {loadError && (
        <div style={s.dangerBox}>Could not load: {loadError}</div>
      )}

      {live === undefined && !loading && (
        <div style={s.warnBox}>
          Could not read the current permissions of <strong>{library}</strong>, so the
          &ldquo;Access now&rdquo; column below is unknown. The mappings themselves are still
          accurate.
        </div>
      )}

      {unexpected.length > 0 && (
        <div style={s.dangerBox}>
          <strong>{unexpected.length} mapping(s) grant {library} access to a group that is not an
          uploader, approver or Staging deleter.</strong> A viewer group here can read other
          people&rsquo;s unapproved documents. Remove them below.
          <div style={s.mono}>{unexpected.map((r) => r.groupName || r.groupId).join(", ")}</div>
        </div>
      )}

      {unmanaged.length > 0 && (
        <div style={s.warnBox}>
          <strong>{unmanaged.length} group(s) already hold permissions on {library} without a
          mapping here</strong> — granted directly in SharePoint. They are not managed by this
          tab and reconciliation will not remove them.
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
            No uploader or approver groups found on this site. Create them in the{" "}
            <strong>User Access</strong> tab first — this tab only grants library entry to groups
            that already exist.
          </div>
        ) : (
          <>
            <div style={s.bulkBar}>
              <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
                <input
                  type="checkbox"
                  style={s.check}
                  disabled={busy}
                  // Indeterminate is not expressible as a prop — set it on the node so a
                  // partial selection does not read as "none selected".
                  ref={(el) => {
                    if (el) el.indeterminate = selected.length > 0 && selected.length < candidates.length;
                  }}
                  checked={selected.length === candidates.length && candidates.length > 0}
                  onChange={(e) => setSelected(e.target.checked ? candidates.map((g) => g.id) : [])}
                />
                <span style={s.bulkCount}>
                  {selected.length > 0
                    ? `${selected.length} of ${candidates.length} selected`
                    : `Select all ${candidates.length}`}
                </span>
              </label>
              {/* Both buttons are always present, each labelled with how many of the
                  selection it would actually touch. A single "Apply" would be ambiguous the
                  moment a selection mixes groups that have access with groups that do not. */}
              <button
                style={busy || scopeMissing || pickedAllowable.length === 0 ? s.off : s.addBtn}
                disabled={busy || scopeMissing || pickedAllowable.length === 0}
                onClick={() => { onAllowSelected(pickedAllowable).catch(() => undefined); }}
              >
                {bulk && pendingAction === "allow"
                  ? <><span style={s.spinner} />Allowing {bulk.done} of {bulk.total}&hellip;</>
                  : `Allow${pickedAllowable.length ? ` ${pickedAllowable.length}` : ""}`}
              </button>
              <button
                style={busy || pickedRemovable.length === 0 ? s.off : s.delBtn}
                disabled={busy || pickedRemovable.length === 0}
                onClick={() => setConfirmRemove(pickedRemovable)}
              >
                {bulk && pendingAction === "remove"
                  ? <><span style={s.spinnerDark} />Removing {bulk.done} of {bulk.total}&hellip;</>
                  : `Remove${pickedRemovable.length ? ` ${pickedRemovable.length}` : ""}`}
              </button>
              {selected.length > 0 && !busy && (
                <button style={s.ghost} onClick={() => setSelected([])}>Clear</button>
              )}
            </div>
            <table style={s.table}>
              <thead>
                <tr>
                  <th style={{ ...s.th, width: 28 }} />
                  <th style={s.th}>Group</th>
                  <th style={s.th}>Role</th>
                  <th style={s.th}>Mapped here</th>
                  <th style={s.th}>Access now</th>
                  <th style={s.th} />
                </tr>
              </thead>
              <tbody>
                {candidates.map((g) => {
                  const row = rowByGroupId.get(String(g.id));
                  const liveGrant = liveByPid.get(g.id);
                  const inFlight = pending.indexOf(g.id) !== -1;
                  // Limited Access alone is not library access — it is auto-granted because
                  // the group holds a folder inside. Saying "Limited Access" here invited
                  // exactly the wrong conclusion, so it is named for what it means.
                  const folderOnly =
                    liveGrant !== undefined && liveGrant.levels.every((n) => n === "Limited Access");
                  return (
                    <tr key={g.id}>
                      <td style={s.checkCell}>
                        <input
                          type="checkbox"
                          style={s.check}
                          // scopeMissing blocks Allow but not Remove — an existing row can
                          // always be taken away, and if the columns went missing after rows
                          // were written, removing them is exactly what an admin needs.
                          disabled={busy || (scopeMissing && !row)}
                          checked={selected.indexOf(g.id) !== -1}
                          onChange={(e) => setSelected((sel) =>
                            e.target.checked ? [...sel, g.id] : sel.filter((id) => id !== g.id),
                          )}
                        />
                      </td>
                      <td style={s.td}>{g.title}</td>
                      <td style={s.td}>{roleFromGroupName(g.title)}</td>
                      <td style={s.td}>
                        {row ? <span style={s.yes}>Yes</span> : <span style={s.no}>No</span>}
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
                      <td style={s.td}>
                        {row ? (
                          <button
                            style={inFlight ? s.working : busy ? s.off : s.delBtn}
                            disabled={busy}
                            onClick={() => setConfirmRemove([row])}
                          >
                            {inFlight ? <><span style={s.spinnerDark} />Removing&hellip;</> : "Remove"}
                          </button>
                        ) : (
                          <button
                            style={inFlight ? s.working : busy || scopeMissing ? s.off : s.addBtn}
                            disabled={busy || scopeMissing}
                            onClick={() => { onAdd(g).catch(() => undefined); }}
                          >
                            {inFlight ? <><span style={s.spinner} />Allowing&hellip;</> : "Allow"}
                          </button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </>
        )}
        <div style={s.hint}>
          &ldquo;Mapped here&rdquo; is what this tab records; &ldquo;Access now&rdquo; is what the
          library&rsquo;s permissions actually say. They differ until Folder Reconciliation runs —
          or if someone changed permissions directly in SharePoint.
        </div>
      </div>

      {confirmRemove && confirmRemove.length > 0 && (
        <div style={s.modalOverlay} onClick={() => setConfirmRemove(undefined)}>
          <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
            <div style={s.modalHead}>
              {confirmRemove.length === 1
                ? `Remove ${library} access?`
                : `Remove ${library} access from ${confirmRemove.length} groups?`}
            </div>
            <div style={s.modalBody}>
              {confirmRemove.length === 1 ? (
                <>
                  <strong>{confirmRemove[0].groupName || confirmRemove[0].groupId}</strong> will no
                  longer be able to open the <strong>{library}</strong> library.
                </>
              ) : (
                <>
                  These groups will no longer be able to open the <strong>{library}</strong>{" "}
                  library:
                  {/* Named, not counted. A count alone cannot be checked against what the
                      admin meant to select, and this is the destructive button. */}
                  <ul style={{ margin: "8px 0 0", paddingLeft: 20 }}>
                    {confirmRemove.map((r) => (
                      <li key={r.itemId}>{r.groupName || r.groupId}</li>
                    ))}
                  </ul>
                </>
              )}
              <div style={{ marginTop: 8 }}>
                Members keep their folder permissions, so direct links to their own folders still
                work — they just lose the library itself.
              </div>
              <div style={{ marginTop: 8 }}>
                This removes the permission immediately, not only the mapping.
              </div>
            </div>
            <div style={s.modalFoot}>
              <button style={s.ghost} onClick={() => setConfirmRemove(undefined)}>Cancel</button>
              <button
                style={s.delBtn}
                onClick={() => { onRemoveMany(confirmRemove).catch(() => undefined); }}
              >
                {confirmRemove.length === 1 ? "Remove access" : `Remove ${confirmRemove.length}`}
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>
          {toast.message}
        </div>
      )}
    </div>
  );
}
