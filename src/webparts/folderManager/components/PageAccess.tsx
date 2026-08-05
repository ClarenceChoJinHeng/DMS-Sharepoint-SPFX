// Page Access tab — who may open each page in Site Pages.
//
// Breaks inheritance on ONE Site Pages item and grants Read. Never the Site Pages library:
// locking the library takes the home page with it, and a user granted site entry who then
// gets 403 on the only page they can navigate to has no way into the site at all.
//
// THREE LIMITS, all stated in the UI because each one is a guarantee somebody will otherwise
// assume (spec §5a):
//   1. It hides the PAGE, not the data. A web part's documents stay reachable by direct URL —
//      folder permissions are the document boundary, page permissions never are.
//   2. Navigation links are NOT security-trimmed. A restricted page keeps its left-nav link
//      and gives Access Denied when clicked. The remedy that needs no groups is removing the
//      link; the client accepted the un-trimmed link on 2026-08-04.
//   3. The home page is excluded outright.
//
// What it DOES hold: item-level permissions on a Site Pages item are enforced server-side. A
// non-member gets Access Denied, the page is trimmed from their search results, and it
// disappears from their Site Pages listing — so its existence is hidden too.
//
// See docs/superpowers/specs/2026-08-03-access-scope-mapping-design.md §5 and §5a.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  buildGroupMapRow,
  LIBRARY_ENTRY_ROLE,
  isForbiddenPageTarget,
  SITE_ENTRY_GROUP_NAME,
} from "../../../shared/groupMapModel";
import { fetchAllSiteGroups, fetchBuiltInGroupIds, SpGroup } from "../../../shared/spGroups";
import { policyForPage, VIEW_ONLY_ROLES } from "../../../shared/pageAccessPolicy";
import { roleFromGroupName } from "../../../shared/groupMapModel";

type Props = { context: WebPartContext; siteUrl: string };

const GROUP_MAP_LIST = "DMS Group Map";
const PAGES_LIST = "Site Pages";

type PageItem = { itemId: number; fileName: string; title: string; unique: boolean };
type EntryRow = { itemId: number; groupId: string; groupName: string; target: string };
type LiveGrant = { principalId: number; title: string; levels: string[] };

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
  checkCell:{ padding: "6px 8px", borderBottom: "1px solid #f0f0f0", width: 28, verticalAlign: "middle" },
  addBtn:   { padding: "4px 12px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  delBtn:   { padding: "4px 12px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  off:      { padding: "4px 12px", fontSize: 12, background: "#c7c7c7", color: "#fff", border: "none", borderRadius: 4, cursor: "not-allowed" },
  working:  { padding: "4px 12px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "wait", opacity: .85 },
  ghost:    { padding: "5px 14px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  yes:      { color: "#0f6c3f", fontWeight: 600 },
  no:       { color: "#8a8886" },
  check:    { width: 15, height: 15, cursor: "pointer", verticalAlign: "middle" },
  bulkBar:  { position: "sticky", top: 0, zIndex: 10, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", padding: "10px 12px", marginBottom: 12, border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 6 },
  bulkCount:{ fontSize: 12, fontWeight: 600, color: "#0f6c3f" },
  warnBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5 },
  dangerBox:{ marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  openBox:  { marginBottom: 16, padding: "10px 12px", border: "1px solid #c7c7c7", background: "#fff", borderRadius: 4, fontSize: 12, color: "#444", lineHeight: 1.5 },
  hint:     { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
  toast:    { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 6, color: "#fff", fontSize: 13, zIndex: 50, boxShadow: "0 4px 14px rgba(0,0,0,.18)" },
  modalOverlay:{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 },
  modalBox: { background: "#fff", borderRadius: 8, width: "100%", maxWidth: 520, boxShadow: "0 8px 30px rgba(0,0,0,.25)" },
  modalHead:{ padding: "12px 16px", borderBottom: "1px solid #eee", fontWeight: 600, fontSize: 14 },
  modalBody:{ padding: "12px 16px" },
  modalFoot:{ padding: "10px 16px", borderTop: "1px solid #eee", display: "flex", justifyContent: "flex-end", gap: 8 },
  spinner:  { display: "inline-block", width: 11, height: 11, marginRight: 6, border: "2px solid rgba(255,255,255,.45)", borderTopColor: "#fff", borderRadius: "50%", verticalAlign: "-1px", animation: "pa-spin .6s linear infinite" },
  spinnerDark:{ display: "inline-block", width: 11, height: 11, marginRight: 6, border: "2px solid rgba(164,38,44,.3)", borderTopColor: "#a4262c", borderRadius: "50%", verticalAlign: "-1px", animation: "pa-spin .6s linear infinite" },
  mono:     { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  drift:    { fontSize: 11, color: "#8a4b00", marginTop: 2 },
  policyBox:{ marginBottom: 12, padding: "8px 10px", border: "1px solid #d6e8dc", background: "#f6fbf8", borderRadius: 4, fontSize: 12, color: "#265", lineHeight: 1.5 },
  adminBox: { marginBottom: 12, padding: "8px 10px", border: "1px solid #cfd8e3", background: "#f4f7fb", borderRadius: 4, fontSize: 12, color: "#2b3f56", lineHeight: 1.5 },
  linkBtn:  { border: "none", background: "transparent", padding: 0, color: "#0f6c3f", fontSize: 12, textDecoration: "underline", cursor: "pointer" },
};

const SPIN_KEYFRAMES = "@keyframes pa-spin{to{transform:rotate(360deg)}}";

export default function PageAccess({ context, siteUrl }: Props): React.ReactElement {
  const [pages, setPages]     = useState<PageItem[]>([]);
  const [target, setTarget]   = useState<string>("");
  const [groups, setGroups]   = useState<SpGroup[]>([]);
  const [rows, setRows]       = useState<EntryRow[]>([]);
  const [live, setLive]       = useState<LiveGrant[] | undefined>(undefined);
  const [welcome, setWelcome] = useState<string>("");
  const [loading, setLoading] = useState(true);
  const [busy, setBusy]       = useState(false);
  const [pending, setPending] = useState<number[]>([]);
  const [selected, setSelected] = useState<number[]>([]);
  const [bulk, setBulk]       = useState<{ done: number; total: number } | undefined>(undefined);
  const [pendingAction, setPendingAction] = useState<"allow" | "remove" | undefined>(undefined);
  const [scopeMissing, setScopeMissing] = useState(false);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [toast, setToast]     = useState<{ message: string; error: boolean } | undefined>(undefined);
  const [confirmRemove, setConfirmRemove] = useState<EntryRow[] | undefined>(undefined);
  const [confirmFirstLock, setConfirmFirstLock] = useState<SpGroup[] | undefined>(undefined);
  // Escape hatch for the page-specific filter. Off by default and reset whenever the page
  // changes, so it can never silently stay on from a previous selection.
  const [showAllGroups, setShowAllGroups] = useState(false);

  const GET = { Accept: "application/json;odata=nometadata" };
  const pagesBase = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PAGES_LIST)}')`;

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 7000);
  };

  // ── Derived ────────────────────────────────────────────────────────────────
  // Above the handlers, not below: allowOne reads rowByGroupId to stay idempotent, and a
  // forward reference into the same render scope works at runtime but reads as an accident.
  const selectablePages = pages.filter(
    (p) => !isForbiddenPageTarget(p.fileName) && p.fileName.toLowerCase() !== welcome,
  );
  const excludedCount = pages.length - selectablePages.length;
  const page = pages.find((p) => p.fileName === target);

  const rowsForPage = rows.filter((r) => r.target.toLowerCase() === target.toLowerCase());
  const rowByGroupId = new Map<string, EntryRow>();
  for (const r of rowsForPage) rowByGroupId.set(r.groupId, r);

  const liveByPid = new Map<number, LiveGrant>();
  for (const l of live ?? []) liveByPid.set(l.principalId, l);

  const policy = policyForPage(target);

  /**
   * Groups offered for this page, filtered by what the page is FOR.
   *
   * Two exclusions are absolute, and neither is overridable:
   *   - the site-entry group, because it contains everyone — granting it a page grants the page
   *     to the whole site, which is never what an admin picking one group means;
   *   - view-only roles (MEMBER / GLOBAL / SEGVIEW), because a reader on an upload form cannot
   *     upload anyway and the grant only looks like it does something.
   *
   * The page-specific filter (uploaders on the upload form, approvers on the approval page) IS
   * overridable, because it is a guess about intent rather than a safety property — and a filter
   * with no escape hatch just sends the admin to SharePoint to do it unsupervised.
   */
  const allGroups = groups
    .filter((g) => g.title.trim().toLowerCase() !== SITE_ENTRY_GROUP_NAME.toLowerCase())
    .filter((g) => VIEW_ONLY_ROLES.indexOf(roleFromGroupName(g.title)) === -1)
    .sort((a, b) => a.title.localeCompare(b.title));

  const eligible = allGroups.filter(
    (g) => policy.roles.indexOf(roleFromGroupName(g.title)) !== -1,
  );

  // A group already mapped is always shown, whatever the policy says. Hiding an existing grant
  // because it no longer fits the rule would leave it in force and invisible — the same mistake
  // the Staging tab's "unexpected mapping" banner exists to avoid.
  const mappedButIneligible = allGroups.filter(
    (g) => rowByGroupId.has(String(g.id)) && eligible.indexOf(g) === -1,
  );

  const candidates = showAllGroups
    ? allGroups
    : eligible.concat(mappedButIneligible).sort((a, b) => a.title.localeCompare(b.title));

  const hiddenCount = allGroups.length - candidates.length;

  /**
   * The site's welcome page, read rather than assumed.
   *
   * isForbiddenPageTarget knows "home.aspx", which covers the default. A site whose welcome
   * page was renamed would slip past that constant and could be locked — the one mistake in
   * this tab that cannot be undone from inside the site. So the real value is read and excluded
   * as well.
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
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST)}')/items` +
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
      const w = await loadWelcome();
      const p = await loadPages();
      const g = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
      const builtIns = await fetchBuiltInGroupIds(context.spHttpClient, siteUrl);
      const r = await loadRows();
      setWelcome(w);
      setPages(p);
      setGroups(g.filter((x) => builtIns.indexOf(x.id) === -1));
      setRows(r);
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

  /** Re-read after a write WITHOUT blanking the table — see StagingAccess for the reasoning. */
  const refreshAfterWrite = async (): Promise<void> => {
    const p = await loadPages().catch(() => pages);
    setPages(p);
    setRows(await loadRows().catch(() => rows));
    setLive(await loadLiveFor(p, target).catch(() => undefined));
  };

  const readLevelId = async (name: string): Promise<number | undefined> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roledefinitions?$select=Id,Name`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!res.ok) return undefined;
    const j = await res.json();
    return ((j.value ?? []) as Array<{ Id: number; Name: string }>).find((d) => d.Name === name)?.Id;
  };

  /**
   * Break inheritance on the page, once.
   *
   * copyRoleAssignments=FALSE, always. With true, every inherited grant is copied forward, so
   * the page stays visible to exactly the same people and the run reports success — a failure
   * invisible from the outside. Site Owners go back on afterwards with Full Control: with
   * nothing copied the only remaining access is site collection administrators, and an owner who
   * is not also an SCA would lose the page they are meant to manage.
   */
  const ensureBroken = async (p: PageItem): Promise<void> => {
    if (p.unique) return;
    const broke = await context.spHttpClient.post(
      `${pagesBase}/items(${p.itemId})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!broke.ok) throw new Error(`breakroleinheritance HTTP ${broke.status}`);
    const ownerRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/AssociatedOwnerGroup?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    const fullId = await readLevelId("Full Control");
    if (!ownerRes.ok || fullId === undefined) return;
    const oj = await ownerRes.json();
    if (typeof oj.Id !== "number") return;
    await context.spHttpClient.post(
      `${pagesBase}/items(${p.itemId})/roleassignments/addroleassignment(principalid=${oj.Id},roledefid=${fullId})`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    ).catch(() => undefined);
  };

  const postRow = async (g: SpGroup): Promise<void> => {
    const row = buildGroupMapRow({
      groupId: String(g.id),
      groupName: g.title,
      role: LIBRARY_ENTRY_ROLE,
      scope: "Page",
      target,
    });
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST)}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { ...GET, "Content-Type": "application/json;odata=nometadata" },
        body: JSON.stringify({ Title: g.title || String(g.id), ...row }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  const deleteRow = async (itemId: number): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST)}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      { headers: { ...GET, "IF-MATCH": "*", "X-HTTP-Method": "DELETE" } },
    );
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
  };

  /**
   * Make a group's access match its mapping: write the row if missing, grant if missing.
   *
   * IDEMPOTENT in both halves, which is what lets one button serve two situations — a brand new
   * grant, and repairing a page somebody edited in SharePoint. Posting a second row for a group
   * that already has one would leave a duplicate no pass would ever clean up, so the row is
   * written only when absent.
   */
  const allowOne = async (g: SpGroup, p: PageItem, readId: number): Promise<boolean> => {
    setPending((q) => [...q, g.id]);
    try {
      if (!rowByGroupId.has(String(g.id))) await postRow(g);
      const res = await context.spHttpClient.post(
        `${pagesBase}/items(${p.itemId})/roleassignments/addroleassignment(principalid=${g.id},roledefid=${readId})`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      return res.ok;
    } finally {
      setPending((q) => q.filter((id) => id !== g.id));
    }
  };

  const runAllow = async (targets: SpGroup[]): Promise<void> => {
    const p = pages.find((x) => x.fileName === target);
    if (!p || targets.length === 0) return;
    setBusy(true);
    setPendingAction("allow");
    setConfirmFirstLock(undefined);
    if (targets.length > 1) setBulk({ done: 0, total: targets.length });
    let ok = 0;
    const failed: string[] = [];
    try {
      const readId = await readLevelId("Read");
      if (readId === undefined) throw new Error('no "Read" permission level on this site');
      // The break happens ONCE, before any grant. Per-group would leave the first grant sitting
      // on an inheriting page — where it means nothing — if a later one failed.
      await ensureBroken(p);
      const broken = { ...p, unique: true };
      for (let i = 0; i < targets.length; i++) {
        try {
          if (await allowOne(targets[i], broken, readId)) ok++; else failed.push(targets[i].title);
        } catch { failed.push(targets[i].title); }
        if (targets.length > 1) setBulk({ done: i + 1, total: targets.length });
      }
      await refreshAfterWrite();
      showToast(
        failed.length === 0
          ? `${ok} group(s) can now open ${p.title}. Everyone else is denied.`
          : `${ok} granted, ${failed.length} failed: ${failed.join(", ")}`,
        failed.length > 0,
      );
    } catch (e) {
      showToast(`Could not restrict ${p.title}: ${(e as Error).message}`, true);
    } finally {
      setBulk(undefined);
      setPendingAction(undefined);
      setSelected([]);
      setBusy(false);
    }
  };

  /**
   * Restrict a page with NO groups on it — the admin-only case.
   *
   * Separate from runAllow because that one returns early on an empty selection, and rightly so:
   * "allow nobody" is a different intent from "allow these", and it has to be asked for
   * explicitly rather than being what happens when a selection is empty.
   */
  const onRestrictAdminOnly = async (p: PageItem): Promise<void> => {
    setBusy(true);
    try {
      await ensureBroken(p);
      await refreshAfterWrite();
      showToast(`${p.title} is now restricted to administrators.`, false);
    } catch (e) {
      showToast(`Could not restrict ${p.title}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /**
   * Allow, with a one-time confirm on the FIRST group for a page.
   *
   * That first grant is the moment the page stops inheriting, and everyone not explicitly
   * granted loses it in that instant. Later grants are additive and unremarkable, so only the
   * first one asks.
   */
  const onAllow = (targets: SpGroup[]): void => {
    const p = pages.find((x) => x.fileName === target);
    if (!p || targets.length === 0) return;
    if (!p.unique) setConfirmFirstLock(targets);
    else runAllow(targets).catch(() => undefined);
  };

  const onRemoveMany = async (targets: EntryRow[]): Promise<void> => {
    const p = pages.find((x) => x.fileName === target);
    if (!p || targets.length === 0) return;
    setBusy(true);
    setConfirmRemove(undefined);
    setPendingAction("remove");
    if (targets.length > 1) setBulk({ done: 0, total: targets.length });
    let ok = 0;
    const failed: string[] = [];
    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      const pid = Number(t.groupId);
      setPending((q) => (pid > 0 ? [...q, pid] : q));
      try {
        await deleteRow(t.itemId);
        if (pid > 0) {
          const res = await context.spHttpClient.post(
            `${pagesBase}/items(${p.itemId})/roleassignments/removeroleassignment(principalid=${pid})`,
            SPHttpClient.configurations.v1,
            { headers: GET },
          );
          if (res.ok) ok++; else failed.push(t.groupName || t.groupId);
        } else ok++;
      } catch {
        failed.push(t.groupName || t.groupId);
      } finally {
        setPending((q) => q.filter((id) => id !== pid));
      }
      if (targets.length > 1) setBulk({ done: i + 1, total: targets.length });
    }
    setBulk(undefined);
    setPendingAction(undefined);
    setSelected([]);
    await refreshAfterWrite();
    setBusy(false);
    showToast(
      failed.length === 0
        ? `${ok} group(s) can no longer open ${p.title}.`
        : `${ok} removed. Could not fully revoke ${failed.length}: ${failed.join(", ")} — check the page's permissions.`,
      failed.length > 0,
    );
  };

  /**
   * Give the page back to everyone by restoring inheritance.
   *
   * The only way out of a restricted page, and offered here because an admin who locked the
   * wrong one would otherwise have to find item-level permissions inside the Site Pages library
   * to undo it. Also deletes the rows, so the tab never claims grants that no longer apply.
   */
  const onResetInheritance = async (p: PageItem): Promise<void> => {
    setBusy(true);
    try {
      const res = await context.spHttpClient.post(
        `${pagesBase}/items(${p.itemId})/resetroleinheritance`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!res.ok) throw new Error(`resetroleinheritance HTTP ${res.status}`);
      for (const r of rows.filter((x) => x.target.toLowerCase() === p.fileName.toLowerCase())) {
        await deleteRow(r.itemId).catch(() => undefined);
      }
      await refreshAfterWrite();
      showToast(`${p.title} is open to everyone with site access again.`, false);
    } catch (e) {
      showToast(`Could not restore inheritance: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /**
   * Can this group open the page RIGHT NOW?
   *
   * On an inheriting page the answer is yes for everyone, so nothing here is treated as drift —
   * that keeps an unrestricted page behaving exactly as it did before this distinction existed.
   */
  const canOpen = (id: number): boolean => (page?.unique !== true ? true : liveByPid.has(id));

  // Four states, from two independent facts: is there a row, and is there a grant. The buttons
  // key off BOTH, because keying off the row alone cannot express the one state an admin
  // actually needs to act on — mapped but denied, i.e. someone changed the page in SharePoint.
  //
  //   row + grant   → in sync                → Remove only
  //   no row + none → not granted            → Allow only
  //   row + none    → DRIFT, access lost     → Allow (re-apply) and Remove
  //   no row + grant→ DRIFT, granted by hand → Allow (adopts it into the mapping)
  const allowable = candidates.filter((g) => !rowByGroupId.has(String(g.id)) || !canOpen(g.id));
  const removable = candidates.filter((g) => rowByGroupId.has(String(g.id)));
  const pickedAllowable = allowable.filter((g) => selected.indexOf(g.id) !== -1);
  const pickedRemovable = removable
    .filter((g) => selected.indexOf(g.id) !== -1)
    .map((g) => rowByGroupId.get(String(g.id)) as EntryRow);

  return (
    <div style={s.wrap}>
      <style>{SPIN_KEYFRAMES}</style>
      <p style={s.intro}>
        Who may <strong>open each page</strong> on this site. Until a page is restricted, everyone
        with access to the site can open it.
      </p>

      <div style={s.warnBox}>
        <strong>This hides the page, not the documents on it.</strong> A restricted page is
        genuinely enforced — non-members get Access Denied, and it disappears from their search
        results and page list. But anything a web part displays stays reachable by direct link, so{" "}
        <strong>folder permissions remain the document boundary</strong>. Also, SharePoint does not
        hide navigation links: a restricted page keeps its left-hand link and gives Access Denied
        when clicked. Remove the link from the navigation if you do not want it seen.
      </div>

      {scopeMissing && (
        <div style={s.dangerBox}>
          The <strong>Scope</strong> and <strong>Target</strong> columns are missing from{" "}
          <strong>{GROUP_MAP_LIST}</strong>. Add them before using this tab.
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
              disabled={busy}
              onChange={(e) => { setTarget(e.target.value); setSelected([]); setShowAllGroups(false); }}
            >
              <option value="">Select a page&hellip;</option>
              {selectablePages.map((p) => (
                <option key={p.itemId} value={p.fileName}>
                  {p.title} ({p.fileName}){p.unique ? " — restricted" : ""}
                </option>
              ))}
            </select>
            <div style={s.hint}>
              {excludedCount > 0 && (
                <>The site home page is not listed — restricting it would grant people the site and
                then deny them the only page they can reach.{" "}</>
              )}
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
              <strong>Open to everyone with site access.</strong> Allowing your first group below
              will restrict it — from that moment only the groups you list here, plus site owners,
              can open it.
            </div>
          ) : (
            <div style={s.openBox}>
              <strong>Restricted.</strong> Only the groups below, plus site owners, can open this
              page.{" "}
              <button
                style={s.ghost}
                disabled={busy}
                onClick={() => { onResetInheritance(page).catch(() => undefined); }}
              >
                Open to everyone again
              </button>
            </div>
          )}

          {/* Who this page is for, and why the rest of the groups are not listed. Stated for
              every page, not only the filtered ones — an admin who cannot see why a group is
              missing assumes the tool is broken. */}
          <div style={policy.adminOnly ? s.adminBox : s.policyBox}>
            {policy.reason}
            {hiddenCount > 0 && !showAllGroups && (
              <>
                {" "}
                <button style={s.linkBtn} onClick={() => setShowAllGroups(true)}>
                  Show all {allGroups.length} groups
                </button>
              </>
            )}
            {showAllGroups && (
              <>
                {" "}
                <button style={s.linkBtn} onClick={() => setShowAllGroups(false)}>
                  Show only the groups for this page
                </button>
              </>
            )}
          </div>

          {policy.adminOnly && candidates.length === 0 && !showAllGroups && (
            <div style={s.hint}>
              Nothing to add here. Restrict the page and leave it with no groups — owners and site
              collection administrators keep access, everyone else is denied.
              {!page.unique && (
                <>
                  {" "}
                  <button
                    style={s.addBtn}
                    disabled={busy}
                    onClick={() => { onRestrictAdminOnly(page).catch(() => undefined); }}
                  >
                    Restrict to administrators
                  </button>
                </>
              )}
            </div>
          )}

          <div style={s.bulkBar}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}>
              <input
                type="checkbox"
                style={s.check}
                disabled={busy}
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
            <button
              style={busy || scopeMissing || pickedAllowable.length === 0 ? s.off : s.addBtn}
              disabled={busy || scopeMissing || pickedAllowable.length === 0}
              onClick={() => onAllow(pickedAllowable)}
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
                <th style={s.th}>Mapped here</th>
                <th style={s.th}>Access now</th>
                <th style={s.th} />
              </tr>
            </thead>
            <tbody>
              {candidates.map((g) => {
                const row = rowByGroupId.get(String(g.id));
                const grant = liveByPid.get(g.id);
                const inFlight = pending.indexOf(g.id) !== -1;
                // Only meaningful on a restricted page with a readable ACL: on an inheriting
                // page everyone can open it, and with an unreadable ACL we do not know.
                const outOfSync =
                  page.unique && live !== undefined && (row !== undefined) !== (grant !== undefined);
                const needsAllow = !row || !canOpen(g.id);
                return (
                  <tr key={g.id}>
                    <td style={s.checkCell}>
                      <input
                        type="checkbox"
                        style={s.check}
                        disabled={busy || (scopeMissing && !row)}
                        checked={selected.indexOf(g.id) !== -1}
                        onChange={(e) => setSelected((sel) =>
                          e.target.checked ? [...sel, g.id] : sel.filter((id) => id !== g.id),
                        )}
                      />
                    </td>
                    <td style={s.td}>{g.title}</td>
                    <td style={s.td}>{row ? <span style={s.yes}>Yes</span> : <span style={s.no}>No</span>}</td>
                    <td style={s.td}>
                      {!page.unique
                        ? <span style={s.no}>inherits — can open</span>
                        : live === undefined
                          ? <span style={s.no}>unknown</span>
                          : grant
                            ? <span style={s.yes}>{grant.levels.join(", ") || "granted"}</span>
                            : <span style={s.no}>denied</span>}
                      {/* Named, not just implied by two disagreeing columns. Across twenty rows
                          the mismatch has to be scannable or nobody spots it. */}
                      {outOfSync && (
                        <div style={s.drift}>
                          {row ? "out of sync — removed in SharePoint" : "granted in SharePoint, not mapped"}
                        </div>
                      )}
                    </td>
                    <td style={s.td}>
                      {/* Both buttons can appear at once, and that is the point: a drifted row
                          has two legitimate repairs — restore the access, or accept the removal
                          and drop the mapping. Choosing for the admin would be guessing. */}
                      {needsAllow && (
                        <button
                          style={inFlight ? s.working : busy || scopeMissing ? s.off : s.addBtn}
                          disabled={busy || scopeMissing}
                          onClick={() => onAllow([g])}
                        >
                          {inFlight
                            ? <><span style={s.spinner} />Working&hellip;</>
                            : row ? "Re-apply" : "Allow"}
                        </button>
                      )}
                      {row && (
                        <button
                          style={{
                            ...(inFlight && !needsAllow ? s.working : busy ? s.off : s.delBtn),
                            ...(needsAllow ? { marginLeft: 6 } : {}),
                          }}
                          disabled={busy}
                          onClick={() => setConfirmRemove([row])}
                        >
                          {inFlight && !needsAllow
                            ? <><span style={s.spinnerDark} />Removing&hellip;</>
                            : "Remove"}
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          <div style={s.hint}>
            &ldquo;Mapped here&rdquo; is what this tab records; &ldquo;Access now&rdquo; is what the
            page&rsquo;s permissions actually say.
          </div>
        </div>
      )}

      {confirmFirstLock && page && (
        <div style={s.modalOverlay} onClick={() => setConfirmFirstLock(undefined)}>
          <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
            <div style={s.modalHead}>Restrict {page.title}?</div>
            <div style={s.modalBody}>
              This page is currently open to everyone with access to the site. Restricting it means{" "}
              <strong>only these groups, plus site owners, can open it</strong> — everyone else gets
              Access Denied immediately:
              <ul style={{ margin: "8px 0 0", paddingLeft: 20 }}>
                {confirmFirstLock.map((g) => <li key={g.id}>{g.title}</li>)}
              </ul>
              <div style={{ marginTop: 8 }}>
                You can reopen it to everyone later with <strong>Open to everyone again</strong>.
              </div>
            </div>
            <div style={s.modalFoot}>
              <button style={s.ghost} onClick={() => setConfirmFirstLock(undefined)}>Cancel</button>
              <button
                style={s.addBtn}
                onClick={() => { runAllow(confirmFirstLock).catch(() => undefined); }}
              >
                Restrict page
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmRemove && confirmRemove.length > 0 && page && (
        <div style={s.modalOverlay} onClick={() => setConfirmRemove(undefined)}>
          <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
            <div style={s.modalHead}>
              {confirmRemove.length === 1
                ? `Remove access to ${page.title}?`
                : `Remove ${confirmRemove.length} groups from ${page.title}?`}
            </div>
            <div style={s.modalBody}>
              {confirmRemove.length === 1 ? (
                <>
                  <strong>{confirmRemove[0].groupName || confirmRemove[0].groupId}</strong> will get
                  Access Denied on <strong>{page.title}</strong>.
                </>
              ) : (
                <>
                  These groups will get Access Denied on <strong>{page.title}</strong>:
                  <ul style={{ margin: "8px 0 0", paddingLeft: 20 }}>
                    {confirmRemove.map((r) => <li key={r.itemId}>{r.groupName || r.groupId}</li>)}
                  </ul>
                </>
              )}
              {removable.length === confirmRemove.length && (
                <div style={{ marginTop: 8, color: "#a4262c" }}>
                  This removes <strong>every</strong> group from the page. It stays restricted, so
                  only site owners will be able to open it — use <strong>Open to everyone
                  again</strong> if that is not what you want.
                </div>
              )}
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
