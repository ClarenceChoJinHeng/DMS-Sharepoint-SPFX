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
  siteEntryGroupTitle,
  normalizeRoleValue,
} from "../../../shared/groupMapModel";
import { fetchAllSiteGroups, SpGroup, SpGroupMember } from "../../../shared/spGroups";
import { AuditOutcome, EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX, libApiTitle } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { MemberRemoval, memberLabel, removalVerdict } from "../../../shared/accessMembers";
import {
  MemberRemovalDialog,
  MemberRows,
  MemberSummary,
  doRemoveMember,
  removalFor,
  useGroupMembers,
} from "./accessMemberUi";

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

  // ── The designer's header block (2026-08-15) ──────────────────────────────────
  // Explanation on the left, callouts on the right. `auto-fit` rather than a fixed pair of columns,
  // so it stacks on a narrow window instead of squeezing the copy to one word a line.
  headGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 320px), 1fr))", gap: 20, alignItems: "start", marginBottom: 20 },
  headLeft: { display: "flex", gap: 14, alignItems: "flex-start" },
  headIcon: { flexShrink: 0, width: 44, height: 44, borderRadius: "50%", background: "#eef7f1", display: "flex", alignItems: "center", justifyContent: "center", color: "#0f6c3f" },
  headTitle:{ fontSize: 15, fontWeight: 600, margin: "0 0 6px", color: "#242424" },
  noteCard: { display: "flex", gap: 10, alignItems: "flex-start", padding: "12px 14px", borderRadius: 6, fontSize: 12, lineHeight: 1.5, marginBottom: 10 },
  noteOk:   { border: "1px solid #b7dcc4", background: "#f3faf5", color: "#1c4d33" },
  noteWarn: { border: "1px solid #f2c9a0", background: "#fff8f0", color: "#7a4300" },
  noteTitle:{ fontWeight: 600, margin: "0 0 3px", fontSize: 12.5 },
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
  // Why the ACL read failed, if it did. Kept so the banner can name a status rather than only saying
  // "could not read" — see loadLive.
  const [liveError, setLiveError] = useState<string | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [toast, setToast]     = useState<{ message: string; error: boolean } | undefined>(undefined);
  // A LIST, so the single-row Remove and the bulk Remove share one confirm dialog. Two
  // dialogs saying almost the same thing is how their wording drifts apart, and the
  // destructive one is the worst place for that.
  const [confirmRemove, setConfirmRemove] = useState<EntryRow[] | undefined>(undefined);
  // ── Per-person removal (2026-08-14) ────────────────────────────────────────
  // Group ids whose member list is expanded. A SET, not a single id: comparing two units' member
  // lists is the reason an admin opens this at all, and an accordion that closes the previous row
  // makes that impossible.
  const [expanded, setExpanded] = useState<number[]>([]);
  // USER ids with a removal in flight — a different namespace from `pending`, which holds GROUP
  // ids. Sharing one array would spin a group's button because a person inside it is being removed.
  const [pendingUsers, setPendingUsers] = useState<number[]>([]);
  const [confirmMember, setConfirmMember] = useState<
    { removal: MemberRemoval; group: SpGroup } | undefined
  >(undefined);

  /**
   * The groups that actually grant entry to this library right now.
   *
   * Membership only means anything relative to THESE: being in a group that grants nothing here is
   * not access, and counting it would put an "also in …" warning on nearly every row — at which
   * point nobody reads the one that matters.
   */
  const allowedGroupIds = rows.map((r) => Number(r.groupId)).filter((id) => id > 0);

  const { members, refresh: refreshMembers } = useGroupMembers(
    context.spHttpClient,
    siteUrl,
    allowedGroupIds,
  );

  // Resolved from the live group list, so a Group Map row naming a group that has since been
  // deleted contributes no title — and is therefore never listed as a reason access survives.
  const titleOf = (groupId: number): string => groups.find((g) => g.id === groupId)?.title ?? "";

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 6000);
  };

  const GET = { Accept: "application/json;odata=nometadata" };
  /**
   * The library's API URL. **A FUNCTION, deliberately — never a render-time const.**
   *
   * `library` is the LOGICAL key ("Staging"), which is what Group Map rows store and what this
   * component filters and writes. A URL needs the live TITLE, and on this site that is
   * "Approval Document", so it must go through `libApiTitle` (gotcha #12: translate at the API
   * boundary, never in stored data).
   *
   * But `libApiTitle` reads a module-level cache primed by `primeNames`, and that has NOT happened on
   * the first render — where it still answers with the legacy `Staging`. As a const, the value
   * captured by that first render was the one `loadLive` used from the mount effect, so the read
   * 404'd; later renders computed the right URL but nothing re-read the ACL, because the effect only
   * re-runs on `[library]` and that never changes. The visible result was the giveaway (found
   * 2026-08-14, second site test): the warning banner named "Approval Document" correctly while the
   * request had asked for "Staging".
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

  /**
   * Record an access change. Fire-and-forget by design.
   *
   * No caller handles the result: the grant or revoke has already happened, and refusing to report a
   * completed permission change because the log was unreachable would be worse than the gap. The
   * console line in writeAudit is the trail when this fails.
   */
  const logAccess = (
    event: string,
    summary: string,
    details: string[],
    outcome: AuditOutcome,
  ): void => {
    writeAudit(context.spHttpClient, siteUrl, {
      event,
      outcome,
      source: "ApprovalLibraryAccess",
      at: new Date(),
      actorName: context.pageContext.user.displayName,
      actorEmail: context.pageContext.user.email,
      // The live title, not the internal key: the log is read by people, and "Staging" is a name
      // the client retired.
      library: libLabel,
      summary,
      details,
    }).catch(() => undefined);
  };

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
      `${listBase()}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${readId})`,
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
      `${listBase()}/roleassignments/removeroleassignment(principalid=${principalId})`,
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
          ? `${g.title} can now open ${libLabel}.`
          : `${g.title} mapped, but the permission could not be applied now — run Folder Reconciliation.`,
        !granted,
      );
      // A half-success is recorded AS a half-success: the row exists, the permission does not. That
      // is exactly the state that later reads as "why can't they open it", so it is not flattened
      // into a plain grant.
      logAccess(
        EVENT.accessGranted,
        granted
          ? `Library access granted — ${g.title} can open ${libLabel}`
          : `Library access mapped but NOT applied — ${g.title}`,
        granted
          ? [`Granted Read on ${libLabel} to: ${g.title}`, "Group Map row written (Library scope)."]
          : [
              `Group Map row written for ${g.title}, but the live permission could not be applied.`,
              "The next Folder Reconciliation will apply it.",
            ],
        granted ? "Success" : "Failed",
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
        ? `${ok} group(s) can now open ${libLabel}.`
        : `${ok} granted, ${failed.length} failed: ${failed.join(", ")}`,
      failed.length > 0,
    );
    // ONE row for the run, not one per group — the same rule reconciliation follows. A twelve-group
    // grant would otherwise push everything else in the log off the first page.
    logAccess(
      EVENT.accessGranted,
      `Library access granted — ${ok} group(s) can open ${libLabel}` +
        (failed.length > 0 ? `, ${failed.length} not applied` : ""),
      [
        `Granted Read on ${libLabel} to ${ok} group(s).`,
        `Groups: ${targets.map((t) => t.title).join(", ")}`,
        failed.length > 0
          ? `Could not fully apply: ${failed.join(", ")} — a Group Map row exists, so reconciliation will apply it.`
          : "All permissions applied live.",
      ],
      failed.length > 0 ? "Failed" : "Success",
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
          ? `${targets[0].groupName || targets[0].groupId} can no longer open ${libLabel}.`
          : `${ok} group(s) can no longer open ${libLabel}.`,
        false,
      );
    } else {
      showToast(
        `${ok} removed. Could not fully revoke ${failed.length}: ${failed.join(", ")} — check the library's permissions.`,
        true,
      );
    }
    // A PARTIAL revoke is the dangerous one — the mapping is gone while the access remains, so
    // nothing on screen shows who still holds it. Recorded as Failed, naming them.
    logAccess(
      EVENT.accessRevoked,
      `Library access revoked — ${ok} group(s) can no longer open ${libLabel}` +
        (failed.length > 0 ? `, ${failed.length} NOT fully revoked` : ""),
      [
        `Removed the Group Map row and the live Read grant for ${ok} group(s).`,
        `Groups: ${targets.map((t) => t.groupName || t.groupId).join(", ")}`,
        failed.length > 0
          ? `NOT fully revoked: ${failed.join(", ")} — the mapping is gone but the permission may remain. Check the library's permissions.`
          : "All live permissions removed.",
      ],
      failed.length > 0 ? "Failed" : "Success",
    );
  };

  /**
   * Remove ONE PERSON from ONE GROUP — the client's actual request (2026-08-14).
   *
   * The narrowest lever this model has, and still not narrow: the grant belongs to the group, so the
   * only way to take one person's access away is to take them out of it, which costs them everything
   * else that group grants — their unit folder included. The dialog says so before this runs; this
   * function's job is to be honest about what happened afterwards.
   *
   * Refreshes ONLY the affected group. A full reload would re-read every group to show one name
   * disappearing, and would lose which rows the admin had expanded.
   */
  const onRemoveMember = async (removal: MemberRemoval): Promise<void> => {
    const uid = removal.user.id;
    setBusy(true);
    setPendingUsers((p) => [...p, uid]);
    setConfirmMember(undefined);
    const who = memberLabel(removal.user);
    const verdict = removalVerdict(removal);
    try {
      const err = await doRemoveMember(context.spHttpClient, siteUrl, removal.groupId, uid);
      await refreshMembers(removal.groupId);
      if (err) {
        showToast(`Could not remove ${who} from ${removal.groupName}: ${err}`, true);
        logAccess(
          EVENT.membersChanged,
          `FAILED to remove ${who} from ${removal.groupName}`,
          [
            `Attempted to remove ${who} (${removal.user.email || `user ${uid}`}) from ${removal.groupName}.`,
            `The removal did not go through: ${err}`,
            "They still hold whatever that group grants.",
          ],
          "Failed",
        );
        return;
      }
      // The toast repeats the verdict rather than saying "removed". "Removed" on a person who still
      // has access through another group is true and useless, and it is the sentence that stops an
      // admin finishing the job.
      showToast(
        verdict === "survives"
          ? `${who} removed from ${removal.groupName}, but still has access via ${removal.survivingGroups.join(", ")}.`
          : verdict === "unknown"
            ? `${who} removed from ${removal.groupName}. Could not confirm whether they still have access via ${removal.unreadable.join(", ")}.`
            : `${who} can no longer open ${libLabel}.`,
        verdict === "survives" || verdict === "unknown",
      );
      logAccess(
        EVENT.membersChanged,
        `${who} removed from ${removal.groupName}` +
          (verdict === "survives" ? " — access RETAINED via another group" : ""),
        [
          `Removed ${who} (${removal.user.email || `user ${uid}`}) from ${removal.groupName}.`,
          `That group grants entry to ${libLabel}, and its folder permissions.`,
          verdict === "survives"
            ? `Access to ${libLabel} REMAINS: they are also in ${removal.survivingGroups.join(", ")}.`
            : verdict === "unknown"
              ? `Whether access remains is unconfirmed — could not read the members of ${removal.unreadable.join(", ")}.`
              : `No other group listed here grants them ${libLabel}.`,
          removal.lastMember
            ? `${removal.groupName} now has no members, but keeps its grant.`
            : "",
        ].filter((l) => l.length > 0),
        // A removal that leaves the access in place is not a clean outcome. Recorded the same way a
        // half-applied revoke is, so the log never reads as "handled" when it is not.
        verdict === "ends" ? "Success" : "Failed",
      );
    } finally {
      setPendingUsers((p) => p.filter((id) => id !== uid));
      setBusy(false);
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
    .filter((g) => g.title.trim().toLowerCase() !== siteEntryGroupTitle().toLowerCase())
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
      <style>{SPIN_KEYFRAMES}</style>
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
                <strong>Access now</strong> column reads <em>unknown</em>. The mappings below are
                still accurate, and everything on this page still saves.
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

      {/* Says the thing the client could not work out from the screen: access is held by groups,
          so removing one person means taking them out of a group, and that is wider than this
          page. Stated up front rather than only in the confirm dialog — by then the admin has
          already decided, and the honest answer to "can I remove just this person?" is "yes, but
          it costs them their folder too", which changes what they do next. */}
      <div style={s.warnBox}>
        <strong>Access is granted to groups, not to individuals.</strong> Expand{" "}
        <strong>People</strong> on any row to see who is in a group and remove one of them — that
        takes them out of the group, so they also lose the folder permissions it grants. To take
        access away from everybody in a group at once, use <strong>Remove</strong> on the group row.
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

      {unexpected.length > 0 && (
        <div style={s.dangerBox}>
          <strong>{unexpected.length} mapping(s) grant {libLabel} access to a group that is not an
          uploader, approver or Staging deleter.</strong> A viewer group here can read other
          people&rsquo;s unapproved documents. Remove them below.
          <div style={s.mono}>{unexpected.map((r) => r.groupName || r.groupId).join(", ")}</div>
        </div>
      )}

      {unmanaged.length > 0 && (
        <div style={s.warnBox}>
          <strong>{unmanaged.length} group(s) already hold permissions on {libLabel} without a
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
            {/* ABOVE the table, not below it. This legend used to sit under the last row — on this
                site that is 17 rows down, so the one explanation of the two columns was somewhere the
                client would never scroll to, and "Mapped here" had to carry the meaning alone. */}
            <div style={{ ...s.hint, margin: "0 0 10px" }}>
              <strong>Should have access</strong> is what you have recorded on this page.{" "}
              <strong>Access now</strong> is what the library&rsquo;s permissions actually say. They
              differ until Folder Reconciliation runs — or if someone changed permissions directly in
              SharePoint.
            </div>
            <table style={s.table}>
              <thead>
                <tr>
                  <th style={{ ...s.th, width: 28 }} />
                  <th style={s.th}>Group</th>
                  <th style={s.th}>Role</th>
                  <th style={s.th}>People</th>
                  {/* "Mapped here" meant nothing to the client (2026-08-15). It named the mechanism —
                      a Group Map row — rather than what the column tells you. These two columns are
                      the intent and the reality, and the pair is only readable if both say so. */}
                  <th style={s.th} title="Yes = you have recorded on this page that this group should be able to open the library.">
                    Should have access
                  </th>
                  <th style={s.th} title="What the library's permissions actually say right now.">
                    Access now
                  </th>
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
                  // Expandable only where there is per-person access to manage. A group with no
                  // mapping grants nobody anything here, so its member list has nothing to remove
                  // FROM — the action on that row is Allow. The count still shows, because an
                  // empty group is worth knowing about before allowing it.
                  const isExpandable = row !== undefined;
                  const isExpanded = isExpandable && expanded.indexOf(g.id) !== -1;
                  return (
                    <React.Fragment key={g.id}>
                    <tr>
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
                      {/* Who is actually in the group — and, for a mapped group, the way in to
                          removing one of them. "Remove GHO_GF_CORU_UPL" is a decision about
                          PEOPLE, and the group name only names them if you already know the
                          convention, which the client by their own account does not. */}
                      <td style={s.td}>
                        <MemberSummary
                          group={g}
                          members={members}
                          expanded={isExpanded}
                          expandable={isExpandable}
                          onToggle={() => setExpanded((ids) =>
                            ids.indexOf(g.id) !== -1 ? ids.filter((i) => i !== g.id) : [...ids, g.id],
                          )}
                        />
                      </td>
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
                    {isExpanded && (
                      <MemberRows
                        group={g}
                        members={members}
                        allowedGroupIds={allowedGroupIds}
                        titleOf={titleOf}
                        // 7 columns: tick, Group, Role, People, Should have access, Access now, action.
                        colSpan={7}
                        busy={busy}
                        pending={pendingUsers}
                        onRetry={() => { refreshMembers(g.id).catch(() => undefined); }}
                        onRemove={(u: SpGroupMember) => setConfirmMember({
                          group: g,
                          removal: removalFor({
                            user: u, group: g, allowedGroupIds, members, titleOf,
                          }),
                        })}
                      />
                    )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </>
        )}
      </div>

      {confirmRemove && confirmRemove.length > 0 && (
        <div style={s.modalOverlay} onClick={() => setConfirmRemove(undefined)}>
          <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
            <div style={s.modalHead}>
              {confirmRemove.length === 1
                ? `Remove ${libLabel} access?`
                : `Remove ${libLabel} access from ${confirmRemove.length} groups?`}
            </div>
            <div style={s.modalBody}>
              {confirmRemove.length === 1 ? (
                <>
                  <strong>{confirmRemove[0].groupName || confirmRemove[0].groupId}</strong> will no
                  longer be able to open the <strong>{libLabel}</strong> library.
                </>
              ) : (
                <>
                  These groups will no longer be able to open the <strong>{libLabel}</strong>{" "}
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

      {confirmMember && (
        <MemberRemovalDialog
          removal={confirmMember.removal}
          busy={busy}
          // The wider effect, in this screen's own terms. The same group grants the unit folder,
          // so this is never "remove from the library" however the button is labelled — and an
          // admin who believes otherwise has quietly revoked someone's ability to upload.
          scopeWarning={
            <>
              This takes them out of <strong>{confirmMember.removal.groupName}</strong>{" "}
              <strong>everywhere</strong>, not just here. That group also grants their folder
              permissions, so they lose the ability to upload to{" "}
              {confirmMember.removal.groupName.length > 0 ? "that unit" : "their unit"} as well.
              {" "}To take away only library entry, remove the whole group with{" "}
              <strong>Remove</strong> on its row instead.
            </>
          }
          onCancel={() => setConfirmMember(undefined)}
          onConfirm={() => { onRemoveMember(confirmMember.removal).catch(() => undefined); }}
        />
      )}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>
          {toast.message}
        </div>
      )}
    </div>
  );
}
