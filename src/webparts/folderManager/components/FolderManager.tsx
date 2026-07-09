import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from "@microsoft/sp-http";
import { IFolderManagerProps } from "./IFolderManagerProps";

const MODES = ["Departments", "Projects"] as const;
type Mode      = (typeof MODES)[number];
type LibTarget = "Staging" | "Documents";

const sanitize = (str: string): string => str.replace(/[\\/:*?"<>|#%]/g, "").trim();
const uid      = (): string => Math.random().toString(36).slice(2, 9);

/* ── Types ─────────────────────────────────────────────────────────────────── */

type RoleDef        = { id: number; name: string };
type GroupPick      = { id: string; displayName: string; mail?: string; isUnified: boolean };
type ExistingAssign = { uid: string; principalId: number; title: string; roleDefId: number; kept: boolean };
type PendingAssign  = { uid: string; group: GroupPick; roleDefId: number };
type LogEntry       = { msg: string; ok: boolean };

// Permission state now lives directly on the folder node so a single "Update"
// commit can apply renames, new-folder creation, and permission edits together.
type PermDraft = {
  existing: ExistingAssign[];
  pending: PendingAssign[];
  loaded: boolean;
  loading: boolean;
  isUnique: boolean | null;
  open: boolean;
};

// Arbitrary-depth folder tree. `isNew` folders don't exist in SharePoint yet —
// `path` stays null for them until the Update commit creates them. Existing
// folders always recompute their real path from parentPath + name at commit
// time (never trust a stored path directly) so an ancestor rename earlier in
// the same commit correctly cascades to every descendant.
type FolderNode = {
  id: string;
  name: string;
  newName: string;
  path: string | null;
  isNew: boolean;
  // Existing (non-new) folders only: staged for deletion but not yet recycled.
  // confirmingDelete gates a second click before isDeleted is actually set.
  isDeleted: boolean;
  confirmingDelete: boolean;
  childrenLoaded: boolean;
  children: FolderNode[];
  perm: PermDraft;
};

/* ── Styles ─────────────────────────────────────────────────────────────────── */

const s: Record<string, React.CSSProperties> = {
  wrap:          { maxWidth: 880, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif" },
  h2:            { fontSize: 22, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle:      { fontSize: 13, color: "#666", margin: "0 0 24px" },
  toggleWrap:    { display: "flex", justifyContent: "center", marginBottom: 24 },
  seg:           { display: "flex", border: "1px solid #0f6c3f", borderRadius: 8, overflow: "hidden" },
  segBtn:        { padding: "8px 22px", fontSize: 13, fontFamily: "'Segoe UI', sans-serif", fontWeight: 600, cursor: "pointer", background: "#fff", color: "#0f6c3f", border: "none", borderRight: "1px solid #0f6c3f" },
  segActive:     { background: "#0f6c3f", color: "#fff" },
  secHeader:     { display: "flex", alignItems: "center", gap: 8, cursor: "pointer", userSelect: "none", margin: "0 0 10px", padding: "4px 0" },
  secTitle:      { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", color: "#0f6c3f", margin: 0 },
  ico:           { fontSize: 11, color: "#0f6c3f", lineHeight: 1, flexShrink: 0 },
  badge:         { fontSize: 11, color: "#888", background: "#f3f3f3", borderRadius: 10, padding: "1px 7px", flexShrink: 0 },
  scrollPane:    { maxHeight: 560, overflowY: "auto", border: "1px solid #e0e0e0", borderRadius: 6, padding: "12px 16px", marginBottom: 8 },
  parentRow:     { display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #f0f0f0" },
  childRow:      { display: "flex", alignItems: "center", gap: 8, padding: "4px 0" },
  chevBtn:       { background: "none", border: "none", cursor: "pointer", padding: "2px 4px", fontSize: 11, color: "#666", lineHeight: 1, flexShrink: 0 },
  renameIn:      { padding: "5px 9px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 13, width: 200, boxSizing: "border-box" },
  wasLabel:      { fontSize: 11, color: "#aaa", fontStyle: "italic", whiteSpace: "nowrap" },
  permBtn:       { marginLeft: "auto", background: "none", border: "1px solid #c8c8c8", borderRadius: 4, padding: "3px 10px", fontSize: 11, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", color: "#444", flexShrink: 0, whiteSpace: "nowrap" },
  permBtnOpen:   { borderColor: "#0f6c3f", color: "#0f6c3f" },
  childrenPane:  { marginLeft: 40, borderLeft: "2px solid #e8f5ee", paddingLeft: 12, marginBottom: 4 },
  permPanel:     { margin: "2px 0 8px", background: "#f8faf8", border: "1px solid #d0e8d8", borderRadius: 6, padding: "10px 12px" },
  permTitle:     { fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em", color: "#0f6c3f", margin: "0 0 8px" },
  assignRow:     { display: "flex", alignItems: "center", gap: 6, marginBottom: 5 },
  chip:          { display: "inline-flex", alignItems: "center", gap: 5, color: "#0f6c3f", fontWeight: 600, fontSize: 11, border: "1px solid #cfe8da", borderRadius: 4, padding: "3px 7px", background: "#fff", maxWidth: 190, boxSizing: "border-box", flexShrink: 0 },
  chipName:      { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  chipX:         { background: "none", border: "none", color: "#888", cursor: "pointer", fontSize: 11, lineHeight: 1, padding: 0, flexShrink: 0 },
  roleSelect:    { flex: "0 0 130px", padding: "4px 6px", border: "1px solid #c8c8c8", borderRadius: 4, fontSize: 12, fontFamily: "'Segoe UI', sans-serif", background: "#fff" },
  undoLink:      { background: "none", border: "none", color: "#0f6c3f", cursor: "pointer", fontSize: 11, padding: 0, fontFamily: "'Segoe UI', sans-serif" },
  searchWrap:    { position: "relative", marginBottom: 6 },
  searchIn:      { padding: "5px 9px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 12, width: "100%", boxSizing: "border-box" },
  dropdown:      { position: "absolute", top: 30, left: 0, right: 0, background: "#fff", border: "1px solid #d0d0d0", borderRadius: 4, boxShadow: "0 6px 18px rgba(0,0,0,.14)", zIndex: 100, maxHeight: 180, overflowY: "auto" },
  dropItem:      { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f2f2f2", fontSize: 12 },
  addFolderBtn:  { background: "none", border: "1px dashed #0f6c3f", color: "#0f6c3f", borderRadius: 4, padding: "4px 12px", fontSize: 12, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", marginTop: 8 },
  actions:       { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 16 },
  btn:           { padding: "8px 22px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 13 },
  logBox:        { marginTop: 20, background: "#f5f5f5", borderRadius: 6, padding: "12px 16px" },
  logTitle:      { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#555", margin: "0 0 8px" },
  toast:         { position: "fixed", top: 24, right: 24, color: "#fff", padding: "14px 44px 14px 16px", borderRadius: 6, fontSize: 13, zIndex: 9999, minWidth: 280, maxWidth: 420, boxShadow: "0 4px 16px rgba(0,0,0,.18)" },
  toastClose:    { position: "absolute", top: 10, right: 12, background: "none", border: "none", cursor: "pointer", color: "#fff", fontSize: 16, opacity: .7, lineHeight: "1" },
  confirmBar:    { display: "flex", alignItems: "center", gap: 10, margin: "4px 0 8px", padding: "8px 10px", background: "#fdf3f3", border: "1px solid #f1c0c0", borderRadius: 4, fontSize: 12, color: "#a4262c" },
  dangerBtn:     { padding: "5px 14px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 12, border: "none", background: "#a4262c", color: "#fff", flexShrink: 0 },
  ghostBtn:      { padding: "5px 14px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 12, border: "1px solid #d0d0d0", background: "#fff", color: "#333", flexShrink: 0 },
};

/* ── GroupSearch ─────────────────────────────────────────────────────────────── */

const GroupSearch: React.FC<{
  disabled: boolean;
  placeholder?: string;
  onSearch: (q: string) => Promise<GroupPick[]>;
  onPick: (g: GroupPick) => void;
}> = ({ disabled, placeholder = "Search for a group…", onSearch, onPick }) => {
  const [q,         setQ]         = useState("");
  const [results,   setResults]   = useState<GroupPick[]>([]);
  const [open,      setOpen]      = useState(false);
  const [searching, setSearching] = useState(false);
  const [focused,   setFocused]   = useState(false);

  useEffect(() => {
    if (!focused) return undefined;
    const h = setTimeout(() => {
      setSearching(true);
      onSearch(q.trim())
        .then(r => { setResults(r); setOpen(true); setSearching(false); })
        .catch(() => { setResults([]); setOpen(false); setSearching(false); });
    }, q.trim().length === 0 ? 0 : 300);
    return () => clearTimeout(h);
  }, [q, focused]);

  return (
    <div style={s.searchWrap}>
      <input
        style={s.searchIn}
        placeholder={placeholder}
        value={q}
        disabled={disabled}
        onFocus={() => { setFocused(true); setOpen(true); }}
        onBlur={() => { window.setTimeout(() => { setFocused(false); setOpen(false); }, 150); }}
        onChange={e => setQ(e.target.value)}
      />
      {searching && <span style={{ position: "absolute", right: 8, top: 7, fontSize: 11, color: "#aaa" }}>Searching…</span>}
      {open && (
        <div style={s.dropdown}>
          {results.length > 0 ? results.map(g => (
            <div key={g.id} style={s.dropItem}
              onMouseDown={() => { onPick(g); setOpen(false); setFocused(false); setResults([]); setQ(""); }}
            >
              <div style={{ fontWeight: 600 }}>{g.displayName}</div>
              {g.mail && <div style={{ fontSize: 11, color: "#888" }}>{g.mail}</div>}
            </div>
          )) : (
            <div style={{ ...s.dropItem, color: "#888", cursor: "default" }}>
              {searching ? "Searching…" : "No groups found"}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

/* ── Main ────────────────────────────────────────────────────────────────────── */

export default function FolderManager({ context }: IFolderManagerProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;

  const [libTarget,    setLibTarget]    = useState<LibTarget>("Staging");
  const [tree,         setTree]         = useState<Record<Mode, FolderNode[]>>({ Departments: [], Projects: [] });
  const [libRoot,      setLibRoot]      = useState<string | null>(null);
  const [loading,      setLoading]      = useState(true);
  const [busy,         setBusy]         = useState(false);
  const [roleDefs,     setRoleDefs]     = useState<RoleDef[]>([]);
  const [ownerGroupId, setOwnerGroupId] = useState<number | null>(null);
  const [log,          setLog]          = useState<LogEntry[]>([]);
  const [toast,        setToast]        = useState<{ message: string; error: boolean } | null>(null);
  const [expandedIds,  setExpandedIds]  = useState<Record<string, boolean>>({});
  const [modeOpen,     setModeOpen]     = useState<Record<Mode, boolean>>({ Departments: true, Projects: true });

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(null), 5000);
  };

  const roleName = (id: number): string => roleDefs.find(r => r.id === id)?.name ?? String(id);

  /* ── REST ────────────────────────────────────────────────────────────────────── */

  const getLibraryRoot = async (lib: string): Promise<string | null> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lib)}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return null;
    const data = await res.json();
    return data.ServerRelativeUrl ?? null;
  };

  const getFolders = async (folderPath: string): Promise<Array<{ Name: string; ServerRelativeUrl: string }>> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Folders?$select=Name,ServerRelativeUrl&$orderby=Name`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return data.value ?? [];
  };

  const moveFolder = async (oldPath: string, newPath: string): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@s)/MoveTo(newUrl=@d)?@s='${encodeURIComponent(oldPath)}'&@d='${encodeURIComponent(newPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const json = await res.json();
        const sp = json?.error?.message?.value ?? json?.error?.message ?? json?.["odata.error"]?.message?.value;
        if (sp) msg += ` — ${sp}`;
      } catch {
        msg += ` — ${(await res.text().catch(() => "")).slice(0, 200)}`;
      }
      throw new Error(msg);
    }
  };

  const getRoleAssignments = async (folderPath: string): Promise<ExistingAssign[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments?$expand=Member,RoleDefinitionBindings&@f='${encodeURIComponent(folderPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    const result: ExistingAssign[] = [];
    for (const ra of (data.value ?? [])) {
      if (ownerGroupId !== null && ra.PrincipalId === ownerGroupId) continue;
      const rawBindings = ra.RoleDefinitionBindings;
      const bindings: Array<{ RoleTypeKind: number; Id: number }> = Array.isArray(rawBindings) ? rawBindings : (rawBindings?.value ?? rawBindings?.results ?? []);
      const valid = bindings.find(b => b.RoleTypeKind !== 1 && b.RoleTypeKind !== 7);
      if (!valid) continue;
      result.push({ uid: uid(), principalId: ra.PrincipalId, title: ra.Member?.Title ?? String(ra.PrincipalId), roleDefId: valid.Id, kept: true });
    }
    return result;
  };

  const getHasUniquePerms = async (folderPath: string): Promise<boolean | null> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields?$select=HasUniqueRoleAssignments&@f='${encodeURIComponent(folderPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return null;
    const data = await res.json();
    return typeof data.HasUniqueRoleAssignments === "boolean" ? data.HasUniqueRoleAssignments : null;
  };

  const folderExists = async (path: string): Promise<boolean> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(path)}')`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    return res.ok;
  };

  const createFolder = async (path: string): Promise<void> => {
    if (await folderExists(path)) return;
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl=@d,overwrite=false)?@d='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" } },
    );
    if (!res.ok) throw new Error(`create folder HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
  };

  const breakInheritance = async (path: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`breakroleinheritance HTTP ${res.status}`);
  };

  // Recycles the folder (and everything inside it) to the site Recycle Bin —
  // recoverable, not a permanent delete.
  const deleteFolder = async (path: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/recycle()?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`delete folder HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
  };

  const ensureGroupPrincipal = async (group: GroupPick): Promise<number> => {
    const logonName = group.isUnified
      ? `c:0o.c|federateddirectoryclaimprovider|${group.id}`
      : `c:0t.c|tenant|${group.id}`;
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" }, body: JSON.stringify({ logonName }) },
    );
    if (!res.ok) throw new Error(`ensureuser HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
    const data = await res.json();
    return data.Id as number;
  };

  const addRoleAssignment = async (path: string, principalId: number, roleDefId: number): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`addroleassignment HTTP ${res.status}`);
  };

  const searchGroups = async (query: string): Promise<GroupPick[]> => {
    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");
    const q = query.trim();
    let req = client.api("/groups").header("ConsistencyLevel", "eventual").count(true).select("id,displayName,mail,groupTypes").top(25);
    if (q.length >= 1) req = req.search(`"displayName:${q}"`);
    const res = await req.get();
    const groups = (res as { value?: Array<{ id: string; displayName: string; mail?: string; groupTypes?: string[] }> }).value ?? [];
    return groups.map(g => ({ id: g.id, displayName: g.displayName, mail: g.mail, isUnified: (g.groupTypes ?? []).indexOf("Unified") !== -1 }));
  };

  /* ── Init ────────────────────────────────────────────────────────────────────── */

  useEffect(() => {
    const loadRoleDefs = async (): Promise<void> => {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/roledefinitions?$select=Id,Name,Hidden,RoleTypeKind`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return;
      const data = await res.json();
      const all = (data.value ?? []) as Array<{ Id: number; Name: string; Hidden: boolean; RoleTypeKind: number }>;
      setRoleDefs(all.filter(r => !r.Hidden && r.RoleTypeKind !== 1 && r.RoleTypeKind !== 7).map(r => ({ id: r.Id, name: r.Name })));
    };
    const loadOwnerGroup = async (): Promise<void> => {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/AssociatedOwnerGroup?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return;
      const data = await res.json();
      if (data.Id != null) setOwnerGroupId(data.Id as number);
    };
    Promise.all([loadRoleDefs(), loadOwnerGroup()]).catch(() => undefined);
  }, []);

  /* ── Tree ────────────────────────────────────────────────────────────────────── */

  const makeNode = (name: string, path: string): FolderNode => ({
    id: uid(), name, newName: name, path, isNew: false, isDeleted: false, confirmingDelete: false, childrenLoaded: false, children: [],
    perm: { existing: [], pending: [], loaded: false, loading: false, isUnique: null, open: false },
  });

  const makeNewNode = (): FolderNode => ({
    id: uid(), name: "", newName: "", path: null, isNew: true, isDeleted: false, confirmingDelete: false, childrenLoaded: true, children: [],
    perm: { existing: [], pending: [], loaded: true, loading: false, isUnique: null, open: true },
  });

  const loadTree = async (): Promise<void> => {
    setLoading(true);
    setExpandedIds({});
    const root = await getLibraryRoot(libTarget);
    setLibRoot(root);
    if (!root) {
      setTree({ Departments: [], Projects: [] });
      setLoading(false);
      showToast(`Could not find the "${libTarget}" library.`, true);
      return;
    }
    const result: Record<Mode, FolderNode[]> = { Departments: [], Projects: [] };
    for (const mode of MODES) {
      const parents = await getFolders(`${root}/${mode}`);
      result[mode] = parents.map(p => makeNode(p.Name, p.ServerRelativeUrl));
    }
    setTree(result);
    setLoading(false);
  };

  useEffect(() => {
    loadTree().catch(() => { setLoading(false); showToast("Could not load folders. Check your permissions.", true); });
  }, [libTarget]);

  /* ── Generic tree mutation helpers (recursive, keyed by node id) ───────────────── */

  const mapTree = (nodes: FolderNode[], id: string, updater: (n: FolderNode) => FolderNode): FolderNode[] =>
    nodes.map(n => n.id === id ? updater(n) : (n.children.length > 0 ? { ...n, children: mapTree(n.children, id, updater) } : n));

  const updateNode = (id: string, updater: (n: FolderNode) => FolderNode): void =>
    setTree(prev => ({ Departments: mapTree(prev.Departments, id, updater), Projects: mapTree(prev.Projects, id, updater) }));

  const filterTree = (nodes: FolderNode[], id: string): FolderNode[] =>
    nodes.filter(n => n.id !== id).map(n => n.children.length > 0 ? { ...n, children: filterTree(n.children, id) } : n);

  const discardNode = (id: string): void =>
    setTree(prev => ({ Departments: filterTree(prev.Departments, id), Projects: filterTree(prev.Projects, id) }));

  /* ── Expand / lazy-load children ────────────────────────────────────────────── */

  const toggleExpand = async (node: FolderNode): Promise<void> => {
    const isOpen = !!expandedIds[node.id];
    if (isOpen) { setExpandedIds(prev => ({ ...prev, [node.id]: false })); return; }
    setExpandedIds(prev => ({ ...prev, [node.id]: true }));
    if (node.childrenLoaded || node.isNew || !node.path) return;
    const kids = await getFolders(node.path).catch(() => []);
    updateNode(node.id, n => ({ ...n, childrenLoaded: true, children: kids.map(c => makeNode(c.Name, c.ServerRelativeUrl)) }));
  };

  /* ── Add / discard folders (staged in-memory until Update) ─────────────────────── */

  const addNewChild = (mode: Mode, parentId: string | null): void => {
    const node = makeNewNode();
    if (parentId === null) {
      setTree(prev => ({ ...prev, [mode]: [...prev[mode], node] }));
      return;
    }
    setExpandedIds(prev => ({ ...prev, [parentId]: true }));
    updateNode(parentId, n => ({ ...n, children: [...n.children, node] }));
  };

  /* ── Delete existing folders (staged in-memory, requires confirmation, applied on Update) ── */

  const requestDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, confirmingDelete: true }));

  const cancelDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, confirmingDelete: false }));

  const confirmDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, confirmingDelete: false, isDeleted: true }));

  const undoDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, isDeleted: false }));

  /* ── Permission draft mutations (existing + new nodes share the same shape) ────── */

  const togglePermPanel = async (node: FolderNode): Promise<void> => {
    if (node.isNew || node.perm.loaded) {
      updateNode(node.id, n => ({ ...n, perm: { ...n.perm, open: !n.perm.open } }));
      return;
    }
    updateNode(node.id, n => ({ ...n, perm: { ...n.perm, open: true, loading: true } }));
    const path = node.path as string;
    const [assignments, isUnique] = await Promise.all([
      getRoleAssignments(path).catch(() => [] as ExistingAssign[]),
      getHasUniquePerms(path).catch(() => null as boolean | null),
    ]);
    updateNode(node.id, n => ({ ...n, perm: { ...n.perm, loading: false, loaded: true, existing: assignments, isUnique } }));
  };

  const permToggleKept = (id: string, aUid: string): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, existing: n.perm.existing.map(a => a.uid === aUid ? { ...a, kept: !a.kept } : a) } }));

  const permSetExistingRole = (id: string, aUid: string, roleDefId: number): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, existing: n.perm.existing.map(a => a.uid === aUid ? { ...a, roleDefId } : a) } }));

  const permAddGroup = (id: string, group: GroupPick): void =>
    updateNode(id, n => {
      if (n.perm.pending.some(a => a.group.id === group.id)) return n;
      const def = roleDefs.find(r => r.name === "Read") ?? roleDefs[0];
      return { ...n, perm: { ...n.perm, pending: [...n.perm.pending, { uid: uid(), group, roleDefId: def?.id ?? 0 }] } };
    });

  const permSetPendingRole = (id: string, aUid: string, roleDefId: number): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, pending: n.perm.pending.map(a => a.uid === aUid ? { ...a, roleDefId } : a) } }));

  const permRemovePending = (id: string, aUid: string): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, pending: n.perm.pending.filter(a => a.uid !== aUid) } }));

  /* ── Change detection ────────────────────────────────────────────────────────── */

  const nodeHasChanges = (node: FolderNode): boolean => {
    if (node.isNew) return true;
    if (node.isDeleted) return true;
    const renamed = node.newName.trim() !== "" && node.newName.trim() !== node.name;
    const permDirty = node.perm.loaded && (node.perm.existing.some(a => !a.kept) || node.perm.pending.length > 0);
    return renamed || permDirty || node.children.some(nodeHasChanges);
  };

  const hasChanges = MODES.some(mode => tree[mode].some(nodeHasChanges));

  /* ── Update (rename + create + permissions, all in one commit) ─────────────────── */

  // Validation walks the WHOLE tree regardless of what's currently expanded on
  // screen, so an error can point at a folder buried inside a collapsed section.
  // Each error carries the mode + ancestor chain needed to reveal it in the UI.
  type ValidationError = { message: string; mode: Mode; ancestorIds: string[] };

  const validateTree = (): ValidationError[] => {
    const errors: ValidationError[] = [];
    const walk = (nodes: FolderNode[], mode: Mode, ancestorIds: string[], parentLabel: string): void => {
      for (const n of nodes) {
        // A folder marked for deletion is recycled whole — its children go with
        // it, so any staged edits inside it are moot and don't need validating.
        if (!n.isNew && n.isDeleted) continue;
        if (n.isNew) {
          const nm = n.newName.trim();
          if (!nm) errors.push({ message: `A new folder under "${parentLabel}" is missing a name.`, mode, ancestorIds });
          else if (n.perm.pending.length === 0) errors.push({ message: `"${nm}" (under "${parentLabel}") needs at least one group assigned.`, mode, ancestorIds });
        } else if (n.perm.loaded) {
          const dirty = n.perm.existing.some(a => !a.kept) || n.perm.pending.length > 0;
          if (dirty) {
            const remaining = n.perm.existing.filter(a => a.kept).length + n.perm.pending.length;
            if (remaining === 0) errors.push({ message: `"${n.newName.trim() || n.name}" would end up with no groups assigned.`, mode, ancestorIds });
          }
        }
        if (n.children.length > 0) walk(n.children, mode, [...ancestorIds, n.id], n.newName.trim() || n.name || "(unnamed folder)");
      }
    };
    MODES.forEach(mode => walk(tree[mode], mode, [], mode));
    return errors;
  };

  const handleUpdate = async (): Promise<void> => {
    const problems = validateTree();
    if (problems.length > 0) {
      const first = problems[0];
      setModeOpen(prev => ({ ...prev, [first.mode]: true }));
      if (first.ancestorIds.length > 0) {
        setExpandedIds(prev => {
          const next = { ...prev };
          first.ancestorIds.forEach(id => { next[id] = true; });
          return next;
        });
      }
      showToast(first.message, true);
      return;
    }
    if (!libRoot) { showToast("Library root not found.", true); return; }

    setBusy(true);
    const entries: LogEntry[] = [];
    const fullCtrlId = roleDefs.find(r => r.name === "Full Control")?.id;
    const principalCache = new Map<string, number>();
    const ensurePrincipal = async (group: GroupPick): Promise<number> => {
      const cached = principalCache.get(group.id);
      if (cached !== undefined) return cached;
      const pid = await ensureGroupPrincipal(group);
      principalCache.set(group.id, pid);
      return pid;
    };

    // Pre-order walk: parents are created/renamed before their children are
    // processed, so each child always receives its parent's up-to-date path.
    const processNode = async (node: FolderNode, parentPath: string): Promise<string | null> => {
      const trimmedNew = node.newName.trim();

      if (!node.isNew && node.isDeleted) {
        const currentPath = `${parentPath}/${node.name}`;
        try {
          await deleteFolder(currentPath);
          entries.push({ msg: `"${node.name}" — deleted ✓`, ok: true });
        } catch (e) {
          entries.push({ msg: `"${node.name}" — delete FAILED: ${(e as Error).message}`, ok: false });
        }
        // Recycling the folder takes its whole subtree with it — nothing left to descend into.
        return null;
      }

      if (node.isNew) {
        const fp = `${parentPath}/${sanitize(trimmedNew)}`;
        try {
          await createFolder(fp);
          await breakInheritance(fp);
          if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(fp, ownerGroupId, fullCtrlId);
          for (const a of node.perm.pending) {
            const pid = await ensurePrincipal(a.group);
            await addRoleAssignment(fp, pid, a.roleDefId);
          }
          const summary = node.perm.pending.map(a => `${a.group.displayName}=${roleName(a.roleDefId)}`).join(", ");
          entries.push({ msg: `"${trimmedNew}" — created · ${summary}`, ok: true });
          for (const child of node.children) await processNode(child, fp);
          return fp;
        } catch (e) {
          entries.push({ msg: `"${trimmedNew}" — create FAILED: ${(e as Error).message}`, ok: false });
          return null;
        }
      }

      // Existing node: always derive the CURRENT real path from parentPath + the
      // original name, since an ancestor rename earlier in this same commit
      // shifts every descendant's real URL — never trust the stored node.path here.
      const currentPath = `${parentPath}/${node.name}`;
      let resolvedPath = currentPath;

      if (trimmedNew && trimmedNew !== node.name) {
        const newPath = `${parentPath}/${sanitize(trimmedNew)}`;
        try {
          await moveFolder(currentPath, newPath);
          entries.push({ msg: `"${node.name}" → "${trimmedNew}" ✓`, ok: true });
          resolvedPath = newPath;
        } catch (e) {
          entries.push({ msg: `"${node.name}" → "${trimmedNew}" FAILED: ${(e as Error).message}`, ok: false });
        }
      }

      const perm = node.perm;
      const permDirty = perm.loaded && (perm.existing.some(a => !a.kept) || perm.pending.length > 0);
      if (permDirty) {
        const kept = perm.existing.filter(a => a.kept);
        try {
          await breakInheritance(resolvedPath);
          if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(resolvedPath, ownerGroupId, fullCtrlId);
          for (const a of kept) await addRoleAssignment(resolvedPath, a.principalId, a.roleDefId);
          for (const a of perm.pending) {
            const pid = await ensurePrincipal(a.group);
            await addRoleAssignment(resolvedPath, pid, a.roleDefId);
          }
          entries.push({ msg: `"${trimmedNew || node.name}" — permissions updated ✓`, ok: true });
        } catch (e) {
          entries.push({ msg: `"${trimmedNew || node.name}" — permissions FAILED: ${(e as Error).message}`, ok: false });
        }
      }

      for (const child of node.children) await processNode(child, resolvedPath);
      return resolvedPath;
    };

    for (const mode of MODES) {
      const modeRootPath = `${libRoot}/${mode}`;
      for (const node of tree[mode]) {
        await processNode(node, modeRootPath);
      }
    }

    setLog(entries);
    setBusy(false);
    const failed = entries.filter(e => !e.ok).length;
    const ok     = entries.filter(e => e.ok).length;
    showToast(
      failed > 0 ? `${ok} update${ok !== 1 ? "s" : ""} applied, ${failed} failed — see log.` :
      ok > 0     ? `${ok} update${ok !== 1 ? "s" : ""} applied.` :
                   "No changes to update.",
      failed > 0,
    );
    await loadTree();
  };

  /* ── Render helpers ──────────────────────────────────────────────────────────── */

  const renderPermPanel = (node: FolderNode): React.ReactElement | null => {
    const panel = node.perm;
    if (!panel.open) return null;
    return (
      <div style={s.permPanel}>
        <p style={s.permTitle}>{node.isNew ? "Assign groups" : "Permissions"}</p>
        {panel.loading ? (
          <p style={{ fontSize: 12, color: "#666", margin: 0 }}>Loading…</p>
        ) : (
          <>
            {!node.isNew && (
              <>
                <p style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase" as const, letterSpacing: ".05em", color: "#555", margin: "0 0 6px" }}>
                  Current assignments
                </p>
                {panel.existing.length === 0 ? (
                  <p style={{ fontSize: 12, color: "#888", margin: "0 0 10px", fontStyle: "italic" }}>
                    {panel.isUnique === false
                      ? "This folder inherits permissions from its parent — no unique assignments are set at the folder level. Add groups below to break inheritance and assign explicit access."
                      : "No groups assigned to this folder yet. Add groups below."}
                  </p>
                ) : (
                  panel.existing.map(a => (
                    <div key={a.uid} style={{ ...s.assignRow, opacity: a.kept ? 1 : .45 }}>
                      <span style={{ ...s.chip, ...(a.kept ? {} : { textDecoration: "line-through" } as React.CSSProperties) }}>
                        <span style={s.chipName} title={a.title}>{a.title}</span>
                      </span>
                      <select style={s.roleSelect} value={a.roleDefId} disabled={busy || !a.kept}
                        onChange={e => permSetExistingRole(node.id, a.uid, Number(e.target.value))}>
                        {roleDefs.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
                      </select>
                      {a.kept
                        ? <button style={s.chipX} disabled={busy} title="Remove" onClick={() => permToggleKept(node.id, a.uid)}>✕</button>
                        : <button style={s.undoLink} disabled={busy} onClick={() => permToggleKept(node.id, a.uid)}>Undo</button>
                      }
                    </div>
                  ))
                )}
              </>
            )}

            <p style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase" as const, letterSpacing: ".05em", color: "#555", margin: "10px 0 6px" }}>
              {node.isNew ? "Groups (required)" : "Add groups"}
            </p>
            {panel.pending.map(a => (
              <div key={a.uid} style={s.assignRow}>
                <span style={s.chip}><span style={s.chipName} title={a.group.displayName}>{a.group.displayName}</span></span>
                <select style={s.roleSelect} value={a.roleDefId} disabled={busy}
                  onChange={e => permSetPendingRole(node.id, a.uid, Number(e.target.value))}>
                  {roleDefs.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
                </select>
                <button style={s.chipX} disabled={busy} onClick={() => permRemovePending(node.id, a.uid)}>✕</button>
              </div>
            ))}
            <GroupSearch disabled={busy} placeholder={node.isNew ? "Add group (required)…" : "Search for a group to add…"} onSearch={searchGroups} onPick={g => permAddGroup(node.id, g)} />

            <p style={{ fontSize: 11, color: "#888", marginTop: 8, marginBottom: 0 }}>
              {node.isNew
                ? "This folder and its permissions will be created when you click Update."
                : "Changes are applied when you click Update below."}
            </p>
          </>
        )}
      </div>
    );
  };

  const renderAddForm = (mode: Mode, parentNode: FolderNode | null): React.ReactElement => (
    <div style={{ paddingTop: 8 }}>
      <button style={s.addFolderBtn} disabled={busy}
        onClick={() => addNewChild(mode, parentNode ? parentNode.id : null)}>
        {`+ ${parentNode ? "Add subfolder" : mode === "Departments" ? "Add department folder" : "Add project folder"}`}
      </button>
    </div>
  );

  const renderNode = (node: FolderNode, mode: Mode, depth: number): React.ReactElement => {
    const isOpen = !!expandedIds[node.id];
    const changed = !node.isNew && !node.isDeleted && node.newName.trim() !== "" && node.newName.trim() !== node.name;
    const permCount = node.perm.existing.filter(a => a.kept).length + node.perm.pending.length;
    const noPendingChanges = !node.perm.existing.some(a => !a.kept) && node.perm.pending.length === 0;
    const permLabel = (() => {
      const arrow = node.perm.open ? "▴" : "▾";
      if (node.isNew) return `Groups ${arrow}`;
      if (!node.perm.loaded) return `Permissions ${arrow}`;
      if (node.perm.isUnique === false && noPendingChanges) return `Permissions (inherited) ${arrow}`;
      return `Permissions (${permCount} group${permCount !== 1 ? "s" : ""}) ${arrow}`;
    })();

    return (
      <div key={node.id}>
        <div style={depth === 0 ? s.parentRow : s.childRow}>
          <button style={s.chevBtn}
            onClick={() => { toggleExpand(node).catch(() => undefined); }}
            title={isOpen ? "Collapse" : "Expand subfolders"}
          >
            {isOpen ? "▾" : "▸"}
          </button>
          <span style={{ fontSize: depth === 0 ? 16 : 14, flexShrink: 0 }}>
            {node.isNew ? "🆕" : depth === 0 ? "📁" : "📂"}
          </span>
          <input className="fm-in" type="text" value={node.newName} disabled={busy || node.isDeleted}
            placeholder={node.isNew ? "New folder name" : undefined}
            onChange={e => updateNode(node.id, n => ({ ...n, newName: e.target.value }))}
            style={{
              ...s.renameIn,
              fontWeight: depth === 0 && !node.isNew ? 600 : 400,
              borderColor: (changed || node.isNew) ? "#0f6c3f" : "#c8c8c8",
              textDecoration: node.isDeleted ? "line-through" : undefined,
              color: node.isDeleted ? "#a4262c" : undefined,
              opacity: node.isDeleted ? .6 : 1,
            }}
          />
          {node.childrenLoaded && node.children.length > 0 && (
            <span style={s.badge}>{node.children.length} subfolder{node.children.length !== 1 ? "s" : ""}</span>
          )}
          {changed && <span style={s.wasLabel}>was: {node.name}</span>}
          {node.isNew && <span style={{ ...s.wasLabel, color: "#0f6c3f" }}>new — not yet created</span>}
          {node.isDeleted && <span style={{ ...s.wasLabel, color: "#a4262c" }}>marked for deletion</span>}

          {!node.isDeleted && (
            <button
              style={{ ...s.permBtn, ...(node.perm.open ? s.permBtnOpen : {}) }}
              onClick={() => { togglePermPanel(node).catch(() => undefined); }}
            >
              {permLabel}
            </button>
          )}
          {node.isNew && (
            <button style={{ ...s.permBtn, color: "#a4262c", borderColor: "#a4262c" }} disabled={busy}
              onClick={() => discardNode(node.id)}>
              Discard
            </button>
          )}
          {!node.isNew && !node.isDeleted && !node.confirmingDelete && (
            <button style={{ ...s.permBtn, color: "#a4262c", borderColor: "#a4262c" }} disabled={busy}
              onClick={() => requestDelete(node.id)}>
              Delete
            </button>
          )}
          {!node.isNew && node.isDeleted && (
            <button style={s.undoLink} disabled={busy} onClick={() => undoDelete(node.id)}>
              Undo
            </button>
          )}
        </div>

        {node.confirmingDelete && (
          <div style={s.confirmBar}>
            <span>Delete &quot;{node.name}&quot; and everything inside it? This moves it to the site Recycle Bin.</span>
            <button style={s.dangerBtn} disabled={busy} onClick={() => confirmDelete(node.id)}>Yes, delete</button>
            <button style={s.ghostBtn} disabled={busy} onClick={() => cancelDelete(node.id)}>Cancel</button>
          </div>
        )}

        {!node.isDeleted && renderPermPanel(node)}

        {!node.isDeleted && (
          <div style={s.childrenPane}>
            {isOpen && node.children.map(child => renderNode(child, mode, depth + 1))}
            {renderAddForm(mode, node)}
          </div>
        )}
      </div>
    );
  };

  /* ── Render ──────────────────────────────────────────────────────────────────── */

  return (
    <section style={s.wrap}>
      <style>{`.fm-in:focus { outline: none; box-shadow: 0 0 0 2px rgba(15,108,63,.18); }`}</style>

      <h2 style={s.h2}>Manage Folders</h2>
      <p style={s.subtitle}>Rename, create, and assign permissions to folders at any depth — then apply it all at once.</p>

      {/* Library toggle */}
      <div style={s.toggleWrap}>
        <div style={s.seg}>
          {(["Staging", "Documents"] as LibTarget[]).map((t, i) => (
            <button key={t}
              onClick={() => { setLibTarget(t); setExpandedIds({}); }}
              style={{ ...s.segBtn, ...(i === 1 ? { borderRight: "none" } : {}), ...(libTarget === t ? s.segActive : {}) }}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      {loading ? (
        <p style={{ fontSize: 13, color: "#666" }}>Loading folders…</p>
      ) : (
        MODES.map(mode => (
          <div key={mode} style={{ marginBottom: 28 }}>
            <div style={s.secHeader} onClick={() => setModeOpen(prev => ({ ...prev, [mode]: !prev[mode] }))}>
              <span style={s.ico}>{modeOpen[mode] ? "▾" : "▸"}</span>
              <p style={s.secTitle}>{mode}</p>
              {tree[mode].length > 0 && <span style={s.badge}>{tree[mode].length}</span>}
            </div>

            {modeOpen[mode] && (
              <>
                {tree[mode].length === 0 ? (
                  <p style={{ fontSize: 13, color: "#999" }}>No folders found under {libTarget}/{mode}.</p>
                ) : (
                  <div style={s.scrollPane}>
                    {tree[mode].map(node => renderNode(node, mode, 0))}
                  </div>
                )}

                {libRoot && renderAddForm(mode, null)}
              </>
            )}
          </div>
        ))
      )}

      <div style={s.actions}>
        <button onClick={() => loadTree().catch(() => undefined)} disabled={busy || loading}
          style={{ ...s.btn, background: "#fff", color: "#0f6c3f", border: "1px solid #0f6c3f" }}>
          Refresh
        </button>
        <button onClick={() => { handleUpdate().catch(() => undefined); }}
          disabled={busy || loading || !hasChanges}
          style={{ ...s.btn, background: !busy && !loading && hasChanges ? "#0f6c3f" : "#9bbfaa", color: "#fff", border: "none", cursor: !busy && !loading && hasChanges ? "pointer" : "default" }}>
          {busy ? "Updating…" : "Update"}
        </button>
      </div>

      {log.length > 0 && (
        <div style={s.logBox}>
          <p style={s.logTitle}>Log</p>
          {log.map((entry, i) => (
            <div key={i} style={{ fontSize: 12, color: entry.ok ? "#0f6c3f" : "#d13438", marginBottom: 4, wordBreak: "break-word" }}>
              {entry.ok ? "✓" : "✗"} {entry.msg}
            </div>
          ))}
        </div>
      )}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#d13438" : "#0f6c3f" }}>
          {toast.message}
          <button onClick={() => setToast(null)} style={s.toastClose}>✕</button>
        </div>
      )}
    </section>
  );
}
