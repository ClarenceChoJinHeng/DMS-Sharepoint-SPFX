import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from "@microsoft/sp-http";
import { IOnboardingProps } from "./IOnboardingProps";

const TARGETS = ["Staging", "Documents"] as const;
type Target = (typeof TARGETS)[number];
const SECTIONS = ["Departments", "Projects"] as const;
type Section = (typeof SECTIONS)[number];

interface GroupPick {
  id: string;
  displayName: string;
  mail?: string;
  isUnified: boolean; // true = M365 group, false = security/other AAD group
}
interface RoleDef {
  id: number;
  name: string;
}
interface Assignment {
  id: string;
  group: GroupPick;
  roleId: number;
}
interface FolderRow {
  key: string;
  name: string;
  assignments: Assignment[];
}
interface RowCallbacks {
  onName: (v: string) => void;
  onAddGroup: (g: GroupPick) => void;
  onLevel: (aid: string, roleId: number) => void;
  onRemoveGroup: (aid: string) => void;
  onRemoveFolder?: () => void;
}
type LogEntry = { msg: string; ok: boolean };

// strip characters SharePoint forbids in folder names
const sanitize = (s: string): string => s.replace(/[\\/:*?"<>|#%]/g, "").trim();
const uid = (): string => Math.random().toString(36).slice(2, 9);

/* ── Styles ─────────────────────────────────────────────────────────────────── */

const styles: Record<string, React.CSSProperties> = {
  wrap: { maxWidth: 880, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif" },
  h2: { fontSize: 22, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle: { fontSize: 13, color: "#666", margin: "0 0 24px" },
  label: { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em", color: "#0f6c3f", margin: "0 0 8px" },
  block: { marginBottom: 24 },
  seg: { display: "flex", border: "1px solid #d0d0d0", borderRadius: 6, overflow: "hidden", width: "fit-content" },
  segBtn: { padding: "8px 18px", fontSize: 13, fontFamily: "'Segoe UI', sans-serif", cursor: "pointer", background: "#fff", color: "#555", border: "none", borderRight: "1px solid #d0d0d0" },
  segActive: { background: "#0f6c3f", color: "#fff", fontWeight: 600 },
  card: { border: "1px solid #e0e0e0", borderRadius: 8, padding: "16px 18px", marginBottom: 12 },
  rowHead: { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#888", margin: "0 0 10px" },
  nameRow: { display: "flex", alignItems: "center", gap: 10, marginBottom: 10 },
  assignRow: { display: "flex", alignItems: "center", gap: 10, marginBottom: 8 },
  input: { padding: "7px 10px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 13, boxSizing: "border-box", width: "100%" },
  nameInput: { flex: "0 0 260px" },
  select: { flex: "0 0 150px", padding: "7px 8px", border: "1px solid #c8c8c8", borderRadius: 4, fontSize: 13, fontFamily: "'Segoe UI', sans-serif", background: "#fff", height: 34 },
  removeFolderBtn: { marginLeft: "auto", background: "none", border: "none", color: "#a4262c", cursor: "pointer", fontSize: 12, fontFamily: "'Segoe UI', sans-serif" },
  greenChip: { display: "inline-flex", alignItems: "center", gap: 8, color: "#0f6c3f", fontWeight: 600, fontSize: 13, border: "1px solid #cfe8da", borderRadius: 4, padding: "5px 10px", background: "#fff", maxWidth: "100%", boxSizing: "border-box" },
  chipName: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  chipX: { background: "none", border: "none", color: "#0f6c3f", cursor: "pointer", fontSize: 12, lineHeight: 1, padding: 0, flexShrink: 0 },
  addBtn: { background: "none", border: "1px dashed #0f6c3f", color: "#0f6c3f", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif" },
  searchHint: { position: "absolute", right: 10, top: 9, fontSize: 11, color: "#aaa" },
  dropdown: { position: "absolute", top: 38, left: 0, right: 0, background: "#fff", border: "1px solid #d0d0d0", borderRadius: 4, boxShadow: "0 6px 18px rgba(0,0,0,.14)", zIndex: 50, maxHeight: 220, overflowY: "auto" },
  dropdownItem: { padding: "8px 10px", cursor: "pointer", borderBottom: "1px solid #f2f2f2", fontSize: 13 },
  actions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 16 },
  btn: { padding: "9px 26px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 13 },
  note: { fontSize: 11, color: "#a07b00", background: "#fff8e1", border: "1px solid #ffe2a8", borderRadius: 4, padding: "8px 12px", margin: "0 0 20px" },
  logBox: { marginTop: 20, background: "#f5f5f5", borderRadius: 6, padding: "12px 16px" },
  logTitle: { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#555", margin: "0 0 8px" },
  toast: { position: "fixed", top: 24, right: 24, color: "#fff", padding: "14px 44px 14px 16px", borderRadius: 6, fontSize: 13, zIndex: 9999, minWidth: 280, maxWidth: 440, boxShadow: "0 4px 16px rgba(0,0,0,.18)" },
  toastClose: { position: "absolute", top: 10, right: 12, background: "none", border: "none", cursor: "pointer", color: "#fff", fontSize: 16, opacity: .7, lineHeight: "1" },
  overlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9000 },
  modal: { background: "#fff", borderRadius: 8, padding: "28px 32px", maxWidth: 520, width: "90%", boxShadow: "0 8px 40px rgba(0,0,0,.22)", fontFamily: "'Segoe UI', sans-serif" },
  modalTitle: { fontSize: 18, fontWeight: 700, color: "#1b1b1b", margin: "0 0 8px" },
  modalSub: { fontSize: 13, color: "#555", margin: "0 0 20px", lineHeight: 1.5 },
  folderTree: { background: "#f8f8f8", borderRadius: 6, padding: "14px 16px", marginBottom: 20 },
  folderLine: { fontSize: 13, color: "#1b1b1b", display: "flex", alignItems: "center", gap: 5, padding: "2px 0" },
  groupTag: { fontSize: 11, color: "#0f6c3f", background: "#eaf5ee", borderRadius: 3, padding: "2px 8px", display: "inline-block", marginBottom: 3 },
};

/* ── Group search field (live Microsoft Graph search) ────────────────────────── */

const GroupSearch: React.FC<{
  disabled: boolean;
  placeholder: string;
  onSearch: (q: string) => Promise<GroupPick[]>;
  onPick: (g: GroupPick) => void;
}> = ({ disabled, placeholder, onSearch, onPick }) => {
  const [q, setQ] = useState("");
  const [results, setResults] = useState<GroupPick[]>([]);
  const [open, setOpen] = useState(false);
  const [searching, setSearching] = useState(false);
  const [focused, setFocused] = useState(false);

  useEffect(() => {
    if (!focused) {
      return undefined;
    }
    const query = q.trim();
    const h = setTimeout(() => {
      setSearching(true);
      onSearch(query)
        .then((r) => { setResults(r); setOpen(true); setSearching(false); })
        .catch(() => { setResults([]); setOpen(true); setSearching(false); });
    }, query.length === 0 ? 0 : 300);
    return () => clearTimeout(h);
  }, [q, focused]);

  return (
    <div style={{ position: "relative", flex: 1 }}>
      <input
        className="ob-input"
        style={styles.input}
        placeholder={placeholder}
        value={q}
        disabled={disabled}
        onFocus={() => { setFocused(true); setOpen(true); }}
        onBlur={() => { window.setTimeout(() => { setFocused(false); setOpen(false); }, 150); }}
        onChange={(e) => setQ(e.target.value)}
      />
      {searching && <span style={styles.searchHint}>Searching…</span>}
      {open && (
        <div style={styles.dropdown}>
          {results.length > 0 ? (
            results.map((g) => (
              <div
                key={g.id}
                style={styles.dropdownItem}
                onMouseDown={() => { onPick(g); setOpen(false); setFocused(false); setResults([]); setQ(""); }}
              >
                <div style={{ fontWeight: 600 }}>{g.displayName}</div>
                {g.mail && <div style={{ fontSize: 11, color: "#888" }}>{g.mail}</div>}
              </div>
            ))
          ) : (
            <div style={{ ...styles.dropdownItem, color: "#888", cursor: "default" }}>
              {searching ? "Searching…" : "No groups found"}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

/* ── Main ────────────────────────────────────────────────────────────────────── */

export default function Onboarding({ context }: IOnboardingProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;

  const [target, setTarget] = useState<Target>("Staging");
  const [section, setSection] = useState<Section>("Departments");
  const [roleDefs, setRoleDefs] = useState<RoleDef[]>([]);
  const [ownerGroupId, setOwnerGroupId] = useState<number | null>(null);
  const [parent, setParent] = useState<FolderRow>({ key: "parent", name: "", assignments: [] });
  const [subs, setSubs] = useState<FolderRow[]>([]);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<LogEntry[]>([]);
  const [toast, setToast] = useState<{ message: string; error: boolean } | null>(null);
  const [confirmOpen, setConfirmOpen] = useState(false);

  /* ── Load the site's own permission levels (role definitions) ─────────────── */

  const loadRoleDefs = async (): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/roledefinitions?$select=Id,Name,Hidden,RoleTypeKind`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return;
    const data = await res.json();
    const all = (data.value ?? []) as Array<{ Id: number; Name: string; Hidden: boolean; RoleTypeKind: number }>;
    // Drop hidden levels, Limited Access (RoleTypeKind 1) and System (7) — keep the
    // rest (Read, Contribute, Edit, Design, Full Control + any custom site levels).
    const usable = all
      .filter((r) => !r.Hidden && r.RoleTypeKind !== 1 && r.RoleTypeKind !== 7)
      .map((r) => ({ id: r.Id, name: r.Name }));
    setRoleDefs(usable);
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

  useEffect(() => {
    loadRoleDefs().catch(() => undefined);
    loadOwnerGroup().catch(() => undefined);
  }, []);

  const defaultRoleId = (): number => (roleDefs.find((r) => r.name === "Read") ?? roleDefs[0])?.id ?? 0;

  /* ── Graph helpers ──────────────────────────────────────────────────────── */

  const searchGroups = async (query: string): Promise<GroupPick[]> => {
    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");
    const q = query.trim();
    let req = client
      .api("/groups")
      .header("ConsistencyLevel", "eventual")
      .count(true)
      .select("id,displayName,mail,groupTypes")
      .top(25);
    if (q.length >= 1) req = req.search(`"displayName:${q}"`);
    const res = await req.get();
    const groups = (res as { value?: Array<{ id: string; displayName: string; mail?: string; groupTypes?: string[] }> }).value ?? [];
    return groups.map((g) => ({
      id: g.id,
      displayName: g.displayName,
      mail: g.mail,
      isUnified: (g.groupTypes ?? []).indexOf("Unified") !== -1,
    }));
  };

  const groupExists = async (id: string): Promise<boolean> => {
    try {
      const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient("3");
      await client.api(`/groups/${id}`).select("id").get();
      return true;
    } catch {
      return false;
    }
  };

  /* ── SharePoint REST helpers ────────────────────────────────────────────── */

  const getLibraryRoot = async (libraryTitle: string): Promise<string | null> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryTitle)}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return null;
    const data = await res.json();
    return data.ServerRelativeUrl ?? null;
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
    if (!res.ok) throw new Error(`create folder: HTTP ${res.status} ${(await res.text().catch(() => "")).slice(0, 200)}`);
  };

  const breakInheritance = async (path: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`breakroleinheritance: HTTP ${res.status}`);
  };

  const ensureGroupPrincipal = async (group: GroupPick): Promise<number> => {
    const logonName = group.isUnified
      ? `c:0o.c|federateddirectoryclaimprovider|${group.id}`
      : `c:0t.c|tenant|${group.id}`;
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
        body: JSON.stringify({ logonName }),
      },
    );
    if (!res.ok) throw new Error(`ensureuser: HTTP ${res.status} ${(await res.text().catch(() => "")).slice(0, 200)}`);
    const data = await res.json();
    return data.Id as number;
  };

  const addRoleAssignment = async (path: string, principalId: number, roleDefId: number): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`addroleassignment: HTTP ${res.status}`);
  };

  /* ── Submit ─────────────────────────────────────────────────────────────── */

  const validateFolder = (row: FolderRow, label: string): string | null => {
    if (!sanitize(row.name)) return `${label} needs a name.`;
    if (row.assignments.length === 0) return `${label} needs at least one group.`;
    for (const a of row.assignments) {
      if (!a.roleId) return `${label}: every group needs a permission level.`;
    }
    return null;
  };

  const handleSubmitClick = (): void => {
    if (roleDefs.length === 0) { setToast({ message: "Permission levels haven't loaded yet — try again in a moment.", error: true }); return; }
    const perr = validateFolder(parent, "Parent folder");
    if (perr) { setToast({ message: perr, error: true }); return; }
    for (let i = 0; i < subs.length; i++) {
      const serr = validateFolder(subs[i], `Subfolder #${i + 1}`);
      if (serr) { setToast({ message: serr, error: true }); return; }
    }
    setConfirmOpen(true);
  };

  const handleConfirm = async (): Promise<void> => {
    setConfirmOpen(false);
    setBusy(true);
    setLog([]);
    const entries: LogEntry[] = [];

    try {
      // Validate every selected group still exists before touching anything.
      const allRows: FolderRow[] = [parent, ...subs];
      const allGroups: GroupPick[] = [];
      allRows.forEach((r) => r.assignments.forEach((a) => allGroups.push(a.group)));
      const uniqueIds = Array.from(new Set(allGroups.map((g) => g.id)));
      for (const gid of uniqueIds) {
        if (!(await groupExists(gid))) {
          const gname = allGroups.find((g) => g.id === gid)?.displayName ?? gid;
          setToast({ message: `Group "${gname}" no longer exists — nothing was created.`, error: true });
          setBusy(false);
          return;
        }
      }
      entries.push({ msg: `Validated ${uniqueIds.length} group(s) ✓`, ok: true });

      const libs: string[] = [target];
      const principalCache = new Map<string, number>();

      const getPrincipal = async (g: GroupPick): Promise<number> => {
        const cached = principalCache.get(g.id);
        if (cached !== undefined) return cached;
        const pid = await ensureGroupPrincipal(g);
        principalCache.set(g.id, pid);
        return pid;
      };

      const roleName = (roleId: number): string => roleDefs.find((r) => r.id === roleId)?.name ?? String(roleId);

      const fullCtrlId = roleDefs.find((r) => r.name === "Full Control")?.id;

      const provision = async (path: string, row: FolderRow, libLabel: string): Promise<void> => {
        await createFolder(path);
        await breakInheritance(path);
        // Always keep site owners visible after breaking inheritance
        if (ownerGroupId !== null && fullCtrlId !== undefined) {
          await addRoleAssignment(path, ownerGroupId, fullCtrlId);
        }
        for (const a of row.assignments) {
          const pid = await getPrincipal(a.group);
          await addRoleAssignment(path, pid, a.roleId);
        }
        const leaf = path.slice(path.lastIndexOf("/") + 1);
        const summary = row.assignments.map((a) => `${a.group.displayName}=${roleName(a.roleId)}`).join(", ");
        entries.push({ msg: `${libLabel}: "${leaf}" — created · inheritance broken · ${summary}`, ok: true });
      };

      for (const lib of libs) {
        const root = await getLibraryRoot(lib);
        if (!root) { entries.push({ msg: `${lib}: library not found — skipped`, ok: false }); continue; }

        const pName = sanitize(parent.name);
        const parentPath = `${root}/${section}/${pName}`;
        try {
          await provision(parentPath, parent, lib);
        } catch (e) {
          entries.push({ msg: `${lib}: parent "${pName}" FAILED — ${(e as Error).message}`, ok: false });
          continue;
        }

        for (const s of subs) {
          const subName = sanitize(s.name);
          const subPath = `${root}/${section}/${pName}/${subName}`;
          try {
            await provision(subPath, s, lib);
          } catch (e) {
            entries.push({ msg: `${lib}: subfolder "${subName}" FAILED — ${(e as Error).message}`, ok: false });
          }
        }
      }

      setLog(entries);
      const failed = entries.filter((e) => !e.ok).length;
      setToast({
        message: failed > 0 ? `Done with ${failed} issue(s) — see log below.` : "Onboarding complete — folders created and groups assigned.",
        error: failed > 0,
      });
      if (failed === 0) {
        setParent({ key: "parent", name: "", assignments: [] });
        setSubs([]);
      }
    } catch (e) {
      setLog(entries);
      setToast({ message: `Unexpected error: ${(e as Error).message}`, error: true });
    } finally {
      setBusy(false);
    }
  };

  /* ── Row helpers ────────────────────────────────────────────────────────── */

  // Parent
  const addParentGroup = (g: GroupPick): void =>
    setParent((p) => (p.assignments.some((a) => a.group.id === g.id) ? p : { ...p, assignments: [...p.assignments, { id: uid(), group: g, roleId: defaultRoleId() }] }));
  const setParentLevel = (aid: string, roleId: number): void =>
    setParent((p) => ({ ...p, assignments: p.assignments.map((a) => (a.id === aid ? { ...a, roleId } : a)) }));
  const removeParentGroup = (aid: string): void =>
    setParent((p) => ({ ...p, assignments: p.assignments.filter((a) => a.id !== aid) }));

  // Subfolders
  const addSubGroup = (key: string, g: GroupPick): void =>
    setSubs((prev) => prev.map((s) => (s.key === key && !s.assignments.some((a) => a.group.id === g.id) ? { ...s, assignments: [...s.assignments, { id: uid(), group: g, roleId: defaultRoleId() }] } : s)));
  const setSubLevel = (key: string, aid: string, roleId: number): void =>
    setSubs((prev) => prev.map((s) => (s.key === key ? { ...s, assignments: s.assignments.map((a) => (a.id === aid ? { ...a, roleId } : a)) } : s)));
  const removeSubGroup = (key: string, aid: string): void =>
    setSubs((prev) => prev.map((s) => (s.key === key ? { ...s, assignments: s.assignments.filter((a) => a.id !== aid) } : s)));
  const setSubName = (key: string, name: string): void =>
    setSubs((prev) => prev.map((s) => (s.key === key ? { ...s, name } : s)));
  const addSub = (): void => setSubs((prev) => [...prev, { key: uid(), name: "", assignments: [] }]);
  const removeSub = (key: string): void => setSubs((prev) => prev.filter((s) => s.key !== key));

  const renderRow = (row: FolderRow, cb: RowCallbacks): React.ReactElement => (
    <div>
      <div style={styles.nameRow}>
        <input
          className="ob-input"
          style={{ ...styles.input, ...styles.nameInput }}
          placeholder="Folder name"
          value={row.name}
          disabled={busy}
          onChange={(e) => cb.onName(e.target.value)}
        />
        {cb.onRemoveFolder && (
          <button style={styles.removeFolderBtn} disabled={busy} onClick={cb.onRemoveFolder}>Remove folder ✕</button>
        )}
      </div>

      {row.assignments.map((a) => (
        <div key={a.id} style={styles.assignRow}>
          <div style={{ flex: 1 }}>
            <span style={styles.greenChip}>
              <span style={styles.chipName} title={a.group.mail ?? a.group.displayName}>{a.group.displayName}</span>
              <button style={styles.chipX} disabled={busy} title="Remove group" onClick={() => cb.onRemoveGroup(a.id)}>✕</button>
            </span>
          </div>
          <select
            style={styles.select}
            value={a.roleId}
            disabled={busy || roleDefs.length === 0}
            onChange={(e) => cb.onLevel(a.id, Number(e.target.value))}
          >
            {roleDefs.length === 0 && <option value={0}>Loading…</option>}
            {roleDefs.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </select>
        </div>
      ))}

      <div style={{ marginTop: row.assignments.length ? 4 : 0 }}>
        <GroupSearch
          disabled={busy}
          placeholder="Add a group — click to pick or type to filter…"
          onSearch={searchGroups}
          onPick={cb.onAddGroup}
        />
      </div>
    </div>
  );

  /* ── Confirm modal ──────────────────────────────────────────────────────── */

  const roleLabel = (roleId: number): string => roleDefs.find((r) => r.id === roleId)?.name ?? String(roleId);
  const confirmLibs: string[] = [target];
  const confirmParentName = sanitize(parent.name) || "(unnamed)";

  /* ── Render ─────────────────────────────────────────────────────────────── */

  return (
    <section style={styles.wrap}>
      <style>{`.ob-input:focus { outline: none; box-shadow: 0 0 0 2px rgba(15,108,63,.18); }`}</style>

      <h2 style={styles.h2}>Folder Onboarding</h2>
      <p style={styles.subtitle}>
        Create a parent folder and its subfolders, break permission inheritance, and assign one or
        more M365 groups to each — in one submit.
      </p>

      <div style={styles.note}>
        Group search needs the <strong>Group.Read.All</strong> Graph permission approved in
        SharePoint Admin → API access. If searches return nothing, that approval is likely pending.
      </div>

      {/* Library target */}
      <div style={styles.block}>
        <p style={styles.label}>Target library</p>
        <div style={styles.seg}>
          {TARGETS.map((t, i) => (
            <button
              key={t}
              disabled={busy}
              onClick={() => setTarget(t)}
              style={{ ...styles.segBtn, ...(i === TARGETS.length - 1 ? { borderRight: "none" } : {}), ...(target === t ? styles.segActive : {}) }}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      {/* Section */}
      <div style={styles.block}>
        <p style={styles.label}>Section</p>
        <div style={styles.seg}>
          {SECTIONS.map((sct, i) => (
            <button
              key={sct}
              disabled={busy}
              onClick={() => setSection(sct)}
              style={{ ...styles.segBtn, ...(i === SECTIONS.length - 1 ? { borderRight: "none" } : {}), ...(section === sct ? styles.segActive : {}) }}
            >
              {sct}
            </button>
          ))}
        </div>
      </div>

      {/* Parent folder */}
      <div style={styles.card}>
        <p style={styles.rowHead}>Parent folder · name + one or more groups</p>
        {renderRow(parent, {
          onName: (v) => setParent((p) => ({ ...p, name: v })),
          onAddGroup: addParentGroup,
          onLevel: setParentLevel,
          onRemoveGroup: removeParentGroup,
        })}
      </div>

      {/* Subfolders */}
      <div style={styles.card}>
        <p style={styles.rowHead}>Subfolders</p>
        {subs.length === 0 && <p style={{ fontSize: 13, color: "#999", margin: "0 0 10px" }}>No subfolders yet.</p>}
        {subs.map((s, idx) => (
          <div key={s.key} style={{ paddingTop: idx === 0 ? 0 : 12, marginTop: idx === 0 ? 0 : 4, borderTop: idx === 0 ? "none" : "1px solid #f0f0f0" }}>
            {renderRow(s, {
              onName: (v) => setSubName(s.key, v),
              onAddGroup: (g) => addSubGroup(s.key, g),
              onLevel: (aid, roleId) => setSubLevel(s.key, aid, roleId),
              onRemoveGroup: (aid) => removeSubGroup(s.key, aid),
              onRemoveFolder: () => removeSub(s.key),
            })}
          </div>
        ))}
        <button style={{ ...styles.addBtn, marginTop: 12 }} disabled={busy} onClick={addSub}>+ Add subfolder</button>
      </div>

      {/* Submit */}
      <div style={styles.actions}>
        <button
          onClick={handleSubmitClick}
          disabled={busy}
          style={{ ...styles.btn, background: busy ? "#9bbfaa" : "#0f6c3f", color: "#fff", border: "none", cursor: busy ? "default" : "pointer" }}
        >
          {busy ? "Provisioning…" : "Submit"}
        </button>
      </div>

      {/* Log */}
      {log.length > 0 && (
        <div style={styles.logBox}>
          <p style={styles.logTitle}>Onboarding Log</p>
          {log.map((entry, i) => (
            <div key={i} style={{ fontSize: 12, color: entry.ok ? "#0f6c3f" : "#d13438", marginBottom: 4, wordBreak: "break-word" }}>
              {entry.ok ? "✓" : "✗"} {entry.msg}
            </div>
          ))}
        </div>
      )}

      {/* Confirm modal */}
      {confirmOpen && (
        <div style={styles.overlay} onClick={() => setConfirmOpen(false)}>
          <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
            <h3 style={styles.modalTitle}>Confirm folder creation</h3>
            <p style={styles.modalSub}>
              The following structure will be created in{" "}
              <strong>{target}</strong>{" "}
              under the <strong>{section}</strong> section:
            </p>

            <div style={styles.folderTree}>
              {confirmLibs.map((lib) => (
                <div key={lib} style={{ marginBottom: confirmLibs.length > 1 ? 16 : 0 }}>
                  {confirmLibs.length > 1 && (
                    <div style={{ fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#888", marginBottom: 8 }}>{lib}</div>
                  )}

                  {/* Parent */}
                  <div style={styles.folderLine}>
                    <span>📁</span>
                    <strong>{confirmParentName}</strong>
                  </div>
                  <div style={{ paddingLeft: 24, marginBottom: 4 }}>
                    {parent.assignments.map((a) => (
                      <div key={a.id} style={{ ...styles.groupTag, marginRight: 4 }}>
                        {a.group.displayName} — {roleLabel(a.roleId)}
                      </div>
                    ))}
                  </div>

                  {/* Subfolders */}
                  {subs.map((s) => {
                    const sName = sanitize(s.name) || "(unnamed)";
                    return (
                      <div key={s.key} style={{ marginTop: 6 }}>
                        <div style={{ ...styles.folderLine, paddingLeft: 16 }}>
                          <span style={{ color: "#bbb", fontSize: 11 }}>└──</span>
                          <span>📁</span>
                          <span>{sName}</span>
                        </div>
                        <div style={{ paddingLeft: 46, marginTop: 2 }}>
                          {s.assignments.map((a) => (
                            <div key={a.id} style={{ ...styles.groupTag, marginRight: 4 }}>
                              {a.group.displayName} — {roleLabel(a.roleId)}
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              ))}
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <button
                onClick={() => setConfirmOpen(false)}
                style={{ ...styles.btn, background: "#fff", color: "#333", border: "1px solid #d0d0d0" }}
              >
                Cancel
              </button>
              <button
                onClick={() => { handleConfirm().catch(() => undefined); }}
                style={{ ...styles.btn, background: "#0f6c3f", color: "#fff", border: "none" }}
              >
                Create folders
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div style={{ ...styles.toast, background: toast.error ? "#d13438" : "#0f6c3f" }}>
          {toast.message}
          <button onClick={() => setToast(null)} style={styles.toastClose}>✕</button>
        </div>
      )}
    </section>
  );
}
