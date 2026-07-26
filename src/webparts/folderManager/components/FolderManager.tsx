import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { searchSiteGroups } from "../../../shared/spGroups";
import { IFolderManagerProps } from "./IFolderManagerProps";
import GroupMapBuilder from "./GroupMapBuilder";
import {
  loadMappedTermGuids,
  resolveFolderByPath,
  writeFolderMapping,
  ensureFolder,
  encodeServerRelativePath,
} from "../../../shared/dmsFolderMap";
import { sanitizeFolderSegment, parseReconModes, RawModeRow } from "../../../shared/formModel";

// A "mode" is a top-level container folder under the library root. These used to
// be hardcoded (Departments / Projects); they are now discovered dynamically so
// the tool works with the multi-segment model (Group Head Office, Group Upstream
// Operations, …) or any future top-level folder naming.
type Mode      = string;
type LibTarget = "Staging" | "Documents";
// The three top-level tabs. The two library tabs drive the folder tree; the
// Reconciliation tab is the term-store-driven provisioner (create + lock + map).
type Tab       = LibTarget | "Reconciliation" | "GroupMap";

// Reconciliation "modes" — mirror Form.tsx / the retired Reconciliation web part.
// Each maps a term set to the segment container folder its terms live under.
// Pilot slice: the four Head Office segments (2026); add the 2027 segments once
// their term sets are onboarded (term-set GUID + its container folder name).
type ReconMode = { key: string; termSetGuid: string; stagingFolder: string };
const RECON_MODES: ReconMode[] = [
  { key: "gho", termSetGuid: "df4b9afa-d3b9-4c04-9097-50dcaf5d8036", stagingFolder: "Group Head Office" },
  { key: "upstream_my_ho", termSetGuid: "16a52947-57a3-4217-9a49-b48cb8b0dd31", stagingFolder: "Upstream Malaysia Head Office" },
  { key: "minamas_ho", termSetGuid: "c6b26d32-1c3e-441f-b4ea-78f053e12990", stagingFolder: "Minamas Head Office" },
  { key: "nbpol_ho", termSetGuid: "303f2c38-086a-46ba-8ee0-85445f6bfa3a", stagingFolder: "NBPOL Head Office" },
];
type TermLite = { id: string; label: string };

// Year/Period and Document Type term sets. Under each leaf (unit) folder the
// provisioner pre-creates the full Year × Document Type grid (path order matches
// the upload form: Unit / Year / Document Type). These inherit the unit's ACL.
// Offline fallback only — the grid term sets are read at runtime from the DMS Config
// `setting` rows (termSet_yearPeriod / termSet_documentType) via loadReconGridTermSets,
// so a different tenant needs no code edit. These GUIDs are the sandbox values.
const YEAR_TERMSET    = "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf";
const DOCTYPE_TERMSET = "866c5754-258e-401f-8685-03d20ae59b1d";

// One folder the provisioner will ensure exists + lock. termGuid is null for the
// segment container folder (not a term); term folders (department, unit, …) carry
// their GUID so Staging can be mapped for rename-proof routing. isLeaf marks the
// deepest terms (upload targets) that get the Year × Document Type grid beneath.
// assignTerm is the term used to look up DMS Group Map rows for THIS folder's tier:
// the term-set GUID for the segment container, the term GUID for dept/unit folders.
type ProvTarget = { termGuid: string | null; assignTerm: string; relPath: string; label: string; section: string; isLeaf: boolean };

// DMS Group Map role → SharePoint permission level. GLOBAL is a privileged
// uploader bypass (not folder-scoped) and is never assigned to a folder.
const ROLE_TO_PERMISSION: Record<string, string> = {
  MEMBER: "Read",
  UPL: "Contribute",
  APR: "Design",
};

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

// Throttle-safety tuning for reconciliation (tunable; could move to DMS Config).
// Grid is created sequentially now (not in parallel) — that parallel burst was what
// tripped SharePoint's 429 throttle. A short delay between writes + an automatic
// cooldown every N ops keeps a large run under the limit.
const RECON_WRITE_DELAY_MS = 500;    // pause between folder/permission writes
const RECON_BATCH_SIZE      = 150;   // writes before an automatic cooldown
const RECON_COOLDOWN_MS     = 4000;  // cooldown length (masked in the UI as "work")
// Rotating status text shown during the cooldown so the pause reads as progress.
const COOLDOWN_MESSAGES = ["Discombobulating…", "Generating folders…"];

// Retry a SharePoint write on HTTP 429/503, honoring Retry-After (seconds) with
// exponential backoff. Returns the final response (ok or not) after up to 5 retries.
async function withThrottleRetry(
  doPost: () => Promise<SPHttpClientResponse>,
  onWait?: (ms: number) => void,
): Promise<SPHttpClientResponse> {
  for (let attempt = 0; ; attempt++) {
    const res = await doPost();
    if ((res.status !== 429 && res.status !== 503) || attempt >= 5) return res;
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(30000, 1000 * 2 ** attempt);
    if (onWait) onWait(waitMs);
    await sleep(waitMs);
  }
}

// Live reconciliation progress feed item (rendered in the two-panel progress view).
type ProgItem = { text: string; status: "run" | "ok" | "skip" | "fail" | "warn" | "admin" };

type GroupMapRow = { groupId: string; groupName: string; role: string };

// SharePoint document libraries keep a system "Forms" folder (and other names
// starting with "_") at the root — never show those as manageable sections.
const isSystemFolder = (name: string): boolean =>
  name === "Forms" || name.startsWith("_");

const sanitize = (str: string): string => str.replace(/[\\/:*?"<>|#%]/g, "").trim();
const uid      = (): string => Math.random().toString(36).slice(2, 9);

// Reserved tree key holding NEW top-level folders staged for creation directly
// under the library root. "*" is illegal in SharePoint folder names, so this key
// can never collide with a real discovered section (top-level folder) name.
const NEW_TOP_LEVEL = "*pending-top-level*";

/* ── Types ─────────────────────────────────────────────────────────────────── */

type RoleDef        = { id: number; name: string };
type GroupPick      = { id: string; displayName: string };
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

  // Active tab. The two library tabs keep libTarget in sync (drives the folder
  // tree); the Reconciliation tab shows the provisioner instead.
  const [tab,          setTab]          = useState<Tab>("Staging");
  const [libTarget,    setLibTarget]    = useState<LibTarget>("Staging");
  // Top-level container folders discovered under the library root, in display order.
  const [sections,     setSections]     = useState<Mode[]>([]);
  const [tree,         setTree]         = useState<Record<Mode, FolderNode[]>>({});
  const [libRoot,      setLibRoot]      = useState<string | null>(null);
  const [loading,      setLoading]      = useState(true);
  const [busy,         setBusy]         = useState(false);
  const [roleDefs,     setRoleDefs]     = useState<RoleDef[]>([]);
  const [ownerGroupId, setOwnerGroupId] = useState<number | null>(null);
  const [log,          setLog]          = useState<LogEntry[]>([]);
  const [toast,        setToast]        = useState<{ message: string; error: boolean } | null>(null);
  const [expandedIds,  setExpandedIds]  = useState<Record<string, boolean>>({});
  // Section collapse state, keyed by section name; sections default to open.
  const [modeOpen,     setModeOpen]     = useState<Record<Mode, boolean>>({});
  // Confirm gate for the Reconciliation tab's provisioning run.
  const [reconConfirm, setReconConfirm] = useState(false);
  // Live reconciliation progress (two-panel view + rotating cooldown text).
  const [reconRunning, setReconRunning] = useState(false);
  const [reconPhase,   setReconPhase]   = useState("");
  const [folderFeed,   setFolderFeed]   = useState<ProgItem[]>([]);
  const [assignFeed,   setAssignFeed]   = useState<ProgItem[]>([]);
  const [reconCounts,  setReconCounts]  = useState<{ folders: number; assigns: number }>({ folders: 0, assigns: 0 });

  // Keep the live feeds short so hundreds of ops don't flood the DOM.
  const FEED_CAP = 40;
  const pushFolder = (text: string, status: ProgItem["status"]): void =>
    setFolderFeed((f) => [...f.slice(-(FEED_CAP - 1)), { text, status }]);
  const setLastFolder = (text: string, status: ProgItem["status"]): void =>
    setFolderFeed((f) => (f.length ? [...f.slice(0, -1), { text, status }] : [{ text, status }]));
  const pushAssign = (text: string, status: ProgItem["status"]): void =>
    setAssignFeed((a) => [...a.slice(-(FEED_CAP - 1)), { text, status }]);
  // Surfaces a 429 backoff wait in the rotating status line (safety-net path).
  const reconWaitNote = (ms: number): void =>
    setReconPhase(`Easing off — SharePoint is busy (${Math.round(ms / 1000)}s)…`);

  // Status glyph for a progress-feed row (spinner while running, else a colored mark).
  const progIcon = (status: ProgItem["status"]): React.ReactElement => {
    if (status === "run") {
      return <span style={{ display: "inline-block", width: 11, height: 11, border: "2px solid #cfe4d8", borderTopColor: "#0f6c3f", borderRadius: "50%", animation: "fmspin 0.8s linear infinite", flexShrink: 0 }} />;
    }
    const marks: Record<string, [string, string]> = {
      ok: ["✓", "#0f6c3f"], skip: ["○", "#999"], fail: ["✗", "#c0392b"], admin: ["⚠", "#b45309"], warn: ["⚠", "#b45309"],
    };
    const [ch, color] = marks[status] ?? ["•", "#666"];
    return <span style={{ color, fontSize: 12, width: 11, textAlign: "center", flexShrink: 0 }}>{ch}</span>;
  };

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
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields?$select=HasUniqueRoleAssignments&@f='${encodeServerRelativePath(folderPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return null;
    const data = await res.json();
    return typeof data.HasUniqueRoleAssignments === "boolean" ? data.HasUniqueRoleAssignments : null;
  };

  const folderExists = async (path: string): Promise<boolean> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    return res.ok;
  };

  const createFolder = async (path: string): Promise<void> => {
    if (await folderExists(path)) return;
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl=@d,overwrite=false)?@d='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" } },
    ), reconWaitNote);
    if (!res.ok) throw new Error(`create folder HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
  };

  const breakInheritance = async (path: string): Promise<void> => {
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ), reconWaitNote);
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

  // SP site group: the group's integer Id IS the role-assignment principal id.
  // No ensureuser, no federateddirectoryclaimprovider claim. A GUID here means a
  // legacy Entra row that must be recreated via the Group Map tab.
  const spGroupPrincipalId = (groupId: string): number => {
    const n = Number((groupId ?? "").trim());
    if (!(n > 0) || n % 1 !== 0) {
      throw new Error(
        `"${groupId}" is not a SharePoint site-group id — recreate this Group Map row with the Group Map tab`,
      );
    }
    return n;
  };

  const addRoleAssignment = async (path: string, principalId: number, roleDefId: number): Promise<void> => {
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})?@f='${encodeURIComponent(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ), reconWaitNote);
    if (!res.ok) throw new Error(`addroleassignment HTTP ${res.status}`);
  };

  const searchGroups = async (query: string): Promise<GroupPick[]> => {
    const groups = await searchSiteGroups(context.spHttpClient, siteUrl, query);
    return groups.map((g) => ({ id: String(g.id), displayName: g.title }));
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
      setSections([]);
      setTree({});
      setLoading(false);
      showToast(`Could not find the "${libTarget}" library.`, true);
      return;
    }
    // Discover the top-level container folders under the library root (e.g. the
    // segment folders) instead of assuming a fixed Departments/Projects layout.
    const topFolders = (await getFolders(root))
      .filter(f => !isSystemFolder(f.Name));
    const sectionNames = topFolders.map(f => f.Name);
    const result: Record<Mode, FolderNode[]> = {};
    for (const f of topFolders) {
      const parents = await getFolders(f.ServerRelativeUrl);
      result[f.Name] = parents.map(p => makeNode(p.Name, p.ServerRelativeUrl));
    }
    setSections(sectionNames);
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
    setTree(prev => {
      const next: Record<Mode, FolderNode[]> = {};
      for (const k of Object.keys(prev)) next[k] = mapTree(prev[k], id, updater);
      return next;
    });

  const filterTree = (nodes: FolderNode[], id: string): FolderNode[] =>
    nodes.filter(n => n.id !== id).map(n => n.children.length > 0 ? { ...n, children: filterTree(n.children, id) } : n);

  const discardNode = (id: string): void =>
    setTree(prev => {
      const next: Record<Mode, FolderNode[]> = {};
      for (const k of Object.keys(prev)) next[k] = filterTree(prev[k], id);
      return next;
    });

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
      setTree(prev => ({ ...prev, [mode]: [...(prev[mode] ?? []), node] }));
      return;
    }
    setExpandedIds(prev => ({ ...prev, [parentId]: true }));
    updateNode(parentId, n => ({ ...n, children: [...n.children, node] }));
  };

  // Stage a brand-new TOP-LEVEL folder, created directly under the library root
  // on Update. Works even when the library currently has no folders at all.
  const addTopLevel = (): void =>
    setTree(prev => ({ ...prev, [NEW_TOP_LEVEL]: [...(prev[NEW_TOP_LEVEL] ?? []), makeNewNode()] }));

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

  const hasChanges =
    sections.some(mode => (tree[mode] ?? []).some(nodeHasChanges)) ||
    (tree[NEW_TOP_LEVEL] ?? []).some(nodeHasChanges);

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
    sections.forEach(mode => walk(tree[mode] ?? [], mode, [], mode));
    // New top-level folders live under the library root; label them by library.
    walk(tree[NEW_TOP_LEVEL] ?? [], NEW_TOP_LEVEL, [], libTarget);
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
    const ensurePrincipal = async (group: GroupPick): Promise<number> =>
      spGroupPrincipalId(group.id);

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

    for (const mode of sections) {
      const modeRootPath = `${libRoot}/${mode}`;
      for (const node of (tree[mode] ?? [])) {
        await processNode(node, modeRootPath);
      }
    }
    // New top-level folders are created directly under the library root.
    for (const node of (tree[NEW_TOP_LEVEL] ?? [])) {
      await processNode(node, libRoot);
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

  /* ── Folder Reconciliation (term-store provisioner) ────────────────────────────── */

  // Term-store readers (v2.1 taxonomy API) — same calls the retired Reconciliation
  // web part used. Top-level terms = the first level under the set (departments for
  // GHO); children recurse to the leaf (unit) terms.
  const loadReconTops = async (termSetId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error(`Term set ${termSetId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({ id: t.id, label: t.labels[0].name }));
  };

  const loadReconChildren = async (termSetId: string, parentId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${parentId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({ id: t.id, label: t.labels[0].name }));
  };

  // Load DMS Group Map keyed by term (lowercased UnitTermGuid) → its group rows.
  // Same list/fields the upload form reads. A term can have several rows (one per
  // role), so a folder gets every mapped group at its role's permission level.
  const loadGroupMapForAssign = async (): Promise<Map<string, GroupMapRow[]>> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Group%20Map')/items?$select=GroupName,GroupId,UnitTermGuid,Role&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    const map = new Map<string, GroupMapRow[]>();
    if (!res.ok) return map;
    const data = await res.json();
    for (const r of (data.value ?? []) as Array<{ GroupName?: string; GroupId?: string; UnitTermGuid?: string; Role?: string }>) {
      const term = (r.UnitTermGuid ?? "").toLowerCase();
      if (!term || !r.GroupId) continue;
      const arr = map.get(term) ?? [];
      arr.push({ groupId: r.GroupId, groupName: r.GroupName ?? r.GroupId, role: (r.Role ?? "").toUpperCase() });
      map.set(term, arr);
    }
    return map;
  };

  // Segments to provision come from the SAME DMS Config `mode` rows the upload form
  // reads, so onboarding a segment is data-only (add a mode row → Run) — no redeploy.
  // Falls back to the built-in RECON_MODES (GHO) if the config is empty/unreachable.
  const loadReconModes = async (): Promise<Array<{ termSetGuid: string; stagingFolder: string }>> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=TermSetGuid,StagingFolder,Levels,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return RECON_MODES;
      const data = await res.json();
      const modes = parseReconModes((data.value ?? []) as RawModeRow[]);
      return modes.length > 0 ? modes : RECON_MODES;
    } catch {
      return RECON_MODES;
    }
  };

  // The Year × Document Type grid term sets come from the SAME DMS Config `setting`
  // rows the upload form reads (termSet_yearPeriod / termSet_documentType), so the
  // grid matches the form on any tenant with no code edit. Falls back to the built-in
  // YEAR_TERMSET / DOCTYPE_TERMSET constants per-key if the row or the list is missing.
  const loadReconGridTermSets = async (): Promise<{ year: string; docType: string }> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return { year: YEAR_TERMSET, docType: DOCTYPE_TERMSET };
      const data = await res.json();
      const map: Record<string, string> = {};
      ((data.value ?? []) as Array<{ Title: string; SettingValue: string }>).forEach(
        (item) => { map[item.Title] = (item.SettingValue ?? "").trim(); },
      );
      return {
        year: map.termSet_yearPeriod || YEAR_TERMSET,
        docType: map.termSet_documentType || DOCTYPE_TERMSET,
      };
    } catch {
      return { year: YEAR_TERMSET, docType: DOCTYPE_TERMSET };
    }
  };

  // Flatten the term store into the folders to provision, parent-before-child so a
  // parent always exists before we create its child. The segment container folder
  // (mode.stagingFolder) has no term; every term below it carries its GUID.
  const buildProvisionTargets = async (): Promise<ProvTarget[]> => {
    const out: ProvTarget[] = [];
    const modes = await loadReconModes();
    for (const mode of modes) {
      // Segment container: not a mapped term, but groups target it via the term-set GUID.
      out.push({ termGuid: null, assignTerm: mode.termSetGuid, relPath: `/${mode.stagingFolder}`, label: mode.stagingFolder, section: mode.stagingFolder, isLeaf: false });
      const tops = await loadReconTops(mode.termSetGuid).catch(() => [] as TermLite[]);
      for (const top of tops) {
        const topTarget: ProvTarget = { termGuid: top.id, assignTerm: top.id, relPath: `/${mode.stagingFolder}/${top.label}`, label: `${mode.stagingFolder} > ${top.label}`, section: mode.stagingFolder, isLeaf: false };
        out.push(topTarget);
        // Recurse; returns whether the term had children. A term with no children
        // is a leaf (the upload target) and gets the Year × Document Type grid.
        const walk = async (parentId: string, ancestors: string[]): Promise<boolean> => {
          const children = await loadReconChildren(mode.termSetGuid, parentId);
          for (const child of children) {
            const chain = [...ancestors, child.label];
            const childTarget: ProvTarget = {
              termGuid: child.id,
              assignTerm: child.id,
              relPath: `/${mode.stagingFolder}/${top.label}/${chain.join("/")}`,
              label: `${mode.stagingFolder} > ${top.label} > ${chain.join(" > ")}`,
              section: mode.stagingFolder,
              isLeaf: false,
            };
            out.push(childTarget);
            childTarget.isLeaf = !(await walk(child.id, chain));
          }
          return children.length > 0;
        };
        topTarget.isLeaf = !(await walk(top.id, []));
      }
    }
    return out;
  };

  // Provision from the term store into BOTH libraries in one run: create each folder
  // if missing and break inheritance at every level (segment → department → unit).
  // Owner group is re-added as Full Control so admins keep access. Then AUTO-ASSIGN
  // the DMS Group Map groups to each folder by role. Staging gets UPL→Contribute and
  // APR→Design only (MEMBER is Documents-only for isolation); Documents gets MEMBER→Read
  // only; GLOBAL is skipped everywhere. Staging term folders are also mapped (term → UniqueId) for
  // rename-proof upload routing; Documents is not mapped. Idempotent: folders/maps
  // are not duplicated; role assignments merge (re-adding an existing one is a no-op),
  // and manual extra grants survive. Folders whose term has no group-map rows are
  // logged as a warning (locked admin-only until groups are added).
  const runReconciliation = async (): Promise<void> => {
    setReconConfirm(false);
    setBusy(true);
    setReconRunning(true);
    setFolderFeed([]);
    setAssignFeed([]);
    setReconCounts({ folders: 0, assigns: 0 });
    setReconPhase("Generating folders…");
    const entries: LogEntry[] = [];
    // Throttle governor: pause between writes, and auto-cooldown every N writes. The
    // cooldown is invisible — the spinner keeps running and the status text rotates so
    // it reads as continuous work (no user click needed to resume).
    let opCount = 0;
    const tick = async (): Promise<void> => {
      opCount++;
      await sleep(RECON_WRITE_DELAY_MS);
      if (opCount % RECON_BATCH_SIZE === 0) {
        for (const msg of COOLDOWN_MESSAGES) {
          setReconPhase(msg);
          await sleep(RECON_COOLDOWN_MS / COOLDOWN_MESSAGES.length);
        }
        setReconPhase("Generating folders…");
      }
    };
    const bumpFolders = (): void => setReconCounts((c) => ({ ...c, folders: c.folders + 1 }));
    const bumpAssigns = (): void => setReconCounts((c) => ({ ...c, assigns: c.assigns + 1 }));
    try {
      const fullCtrlId = roleDefs.find(r => r.name === "Full Control")?.id;
      const readId = roleDefs.find(r => r.name === "Read")?.id;
      if (readId === undefined) entries.push({ msg: `⚠ "Read" role definition not found — ancestor browse access will be skipped`, ok: false });
      // Ancestor rel-paths of a target, excluding the folder itself and the library root.
      // e.g. "/A/B/C" -> ["/A", "/A/B"]. Used to grant each unit group Read up its own path
      // so members can browse down to their folder; siblings without a grant stay
      // security-trimmed (invisible) in the view.
      const ancestorRelPaths = (relPath: string): string[] => {
        const parts = relPath.split("/").filter(Boolean);
        const out: string[] = [];
        let cur = "";
        for (let i = 0; i < parts.length - 1; i++) { cur += `/${parts[i]}`; out.push(cur); }
        return out;
      };
      const mapped = await loadMappedTermGuids(context.spHttpClient, siteUrl);
      const groupMap = await loadGroupMapForAssign();
      const targets = await buildProvisionTargets();
      if (targets.length === 0) {
        showToast("No terms found in the term store to provision.", false);
        setBusy(false);
        return;
      }
      // Year × Document Type grid labels (from their term sets) — pre-created under
      // each leaf/unit folder. Both are flat term sets, so top-level terms suffice.
      // Sanitize labels to safe folder names with the SAME helper the upload form uses
      // (formModel.sanitizeFolderSegment), so reconciliation and Form.tsx always agree on
      // the folder name — e.g. a Document Type term containing illegal chars like "/".
      const gridSets = await loadReconGridTermSets();
      const yearLabels    = (await loadReconTops(gridSets.year).catch(() => [] as TermLite[])).map(y => sanitizeFolderSegment(y.label)).filter(Boolean);
      const docTypeLabels = (await loadReconTops(gridSets.docType).catch(() => [] as TermLite[])).map(d => sanitizeFolderSegment(d.label)).filter(Boolean);
      for (const lib of ["Staging", "Documents"] as LibTarget[]) {
        const root = await getLibraryRoot(lib);
        if (!root) {
          entries.push({ msg: `${lib}: library root not found — skipped`, ok: false });
          continue;
        }
        for (const t of targets) {
          const full = `${root}${t.relPath}`;
          const folderLabel = `${lib}${t.relPath}`;
          try {
            pushFolder(`${folderLabel} — creating…`, "run");
            const existed = await folderExists(full);
            if (!existed) {
              await createFolder(full);
              await breakInheritance(full);
              if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(full, ownerGroupId, fullCtrlId);
              entries.push({ msg: `${folderLabel} — created + locked ✓`, ok: true });
              setLastFolder(`${folderLabel} — created + locked`, "ok");
              bumpFolders();
              await tick();
            } else {
              const isUnique = await getHasUniquePerms(full);
              if (isUnique === false) {
                await breakInheritance(full);
                if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(full, ownerGroupId, fullCtrlId);
                entries.push({ msg: `${folderLabel} — existed, locked ✓`, ok: true });
                setLastFolder(`${folderLabel} — existed, locked`, "ok");
                bumpFolders();
                await tick();
              } else {
                entries.push({ msg: `${folderLabel} — already locked, skipped`, ok: true });
                setLastFolder(`${folderLabel} — already there`, "skip");
              }
            }
            // Map Staging term folders only (skip the segment container + already-mapped).
            if (lib === "Staging" && t.termGuid && !mapped.has(t.termGuid)) {
              const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, full);
              if (resolved) {
                await writeFolderMapping(context.spHttpClient, siteUrl, {
                  termGuid: t.termGuid,
                  folderUniqueId: resolved.uniqueId,
                  title: t.label,
                  folderUrl: resolved.serverRelativeUrl,
                  section: t.section,
                });
                mapped.add(t.termGuid);
                entries.push({ msg: `  ↳ mapped ${t.label} → ${resolved.uniqueId}`, ok: true });
              }
            }
            // Auto-assign DMS Group Map groups to this folder by role. Staging gets
            // UPL/APR only; Documents gets MEMBER (viewer) groups only. Idempotent
            // (add-role merges). A folder whose term has no rows is flagged.
            const groupRows = groupMap.get(t.assignTerm.toLowerCase()) ?? [];
            // Isolation rule: MEMBER (base/viewer) groups are Documents-only and must
            // NEVER land on Staging (else a viewer could see pending docs). Staging gets
            // UPL/APR only; Documents gets MEMBER only.
            const applicable = groupRows.filter(g =>
              ROLE_TO_PERMISSION[g.role] !== undefined &&
              (lib === "Staging" ? g.role !== "MEMBER" : g.role === "MEMBER"),
            );
            if (applicable.length === 0) {
              entries.push({ msg: `  ⚠ ${folderLabel} — no group-map groups for this tier (locked admin-only)`, ok: true });
            }
            const grantedPids: Array<{ groupName: string; pid: number }> = [];
            for (const g of applicable) {
              const roleDefId = roleDefs.find(r => r.name === ROLE_TO_PERMISSION[g.role])?.id;
              if (roleDefId === undefined) {
                entries.push({ msg: `  ⚠ ${g.groupName} — no "${ROLE_TO_PERMISSION[g.role]}" role definition on site`, ok: false });
                pushAssign(`${t.label}: "${ROLE_TO_PERMISSION[g.role]}" role missing on site`, "warn");
                continue;
              }
              try {
                const pid = spGroupPrincipalId(g.groupId);
                await addRoleAssignment(full, pid, roleDefId);
                grantedPids.push({ groupName: g.groupName, pid });
                entries.push({ msg: `  ↳ ${g.groupName} → ${ROLE_TO_PERMISSION[g.role]}`, ok: true });
                pushAssign(`${t.label} → ${g.groupName} (${ROLE_TO_PERMISSION[g.role]})`, "ok");
                bumpAssigns();
                await tick();
              } catch (e) {
                entries.push({ msg: `  ✗ ${g.groupName} → ${ROLE_TO_PERMISSION[g.role]} FAILED: ${(e as Error).message}`, ok: false });
                // Most common cause: the SP group doesn't exist yet (or is a legacy Entra
                // row). Surface the admin-needs-to-create-it message in the right panel.
                pushAssign(`Group "${g.groupName}" not found — ask an administrator to create it`, "admin");
              }
            }
            // Browse access: grant each just-assigned group Read on every ANCESTOR folder on
            // its own path (same library), so members can navigate down to their unit.
            // Ancestors get Read only (never edit) — pass-through, not write targets. Siblings
            // get no grant, so SharePoint security-trims them: a user sees only their corridor.
            if (grantedPids.length > 0 && readId !== undefined) {
              for (const anc of ancestorRelPaths(t.relPath)) {
                const ancFull = `${root}${anc}`;
                for (const gp of grantedPids) {
                  try {
                    await addRoleAssignment(ancFull, gp.pid, readId);
                    entries.push({ msg: `  ↳ ${gp.groupName} → Read (browse) on ${anc}`, ok: true });
                  } catch (e) {
                    entries.push({ msg: `  ✗ ${gp.groupName} → Read (browse) on ${anc} FAILED: ${(e as Error).message}`, ok: false });
                  }
                }
              }
            }
            // Under leaf (unit) folders, pre-create the Year × Document Type grid.
            // These inherit the unit's ACL (no lock, no map). Idempotent via
            // ensureFolder. Logged as a per-unit count, not one line per folder.
            if (t.isLeaf && yearLabels.length > 0) {
              const gridTotal = yearLabels.length * (1 + docTypeLabels.length);
              let grid = 0;
              pushFolder(`${t.label} grid: 0 / ${gridTotal}`, "run");
              for (const yr of yearLabels) {
                const yearFolder = await ensureFolder(context.spHttpClient, siteUrl, full, yr);
                if (yearFolder) { grid++; bumpFolders(); }
                setLastFolder(`${t.label} grid: ${grid} / ${gridTotal}`, "run");
                await tick();
                if (!yearFolder) continue;
                // Sequential (NOT parallel) — the old parallel burst was what tripped the
                // 429 throttle. One folder at a time + the inter-write delay keeps it safe.
                for (const dt of docTypeLabels) {
                  const made = await ensureFolder(context.spHttpClient, siteUrl, yearFolder.serverRelativeUrl, dt);
                  if (made) { grid++; bumpFolders(); }
                  setLastFolder(`${t.label} grid: ${grid} / ${gridTotal}`, "run");
                  await tick();
                }
              }
              entries.push({ msg: `  ↳ ${lib}${t.relPath} — Year × Document Type grid: ${grid} folder(s) ensured`, ok: true });
              setLastFolder(`${t.label} grid: ${grid} / ${gridTotal} ✓`, "ok");
            }
          } catch (e) {
            entries.push({ msg: `${lib}${t.relPath} — FAILED: ${(e as Error).message}`, ok: false });
          }
        }
      }
      setLog(entries);
      const failed = entries.filter(e => !e.ok).length;
      showToast(
        failed > 0
          ? `Reconciled with ${failed} error(s) — see log.`
          : `Reconciled ${targets.length} folder(s) across Staging + Documents — locked + groups assigned from DMS Group Map.`,
        failed > 0,
      );
    } catch (e) {
      showToast(`Reconciliation failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
      setReconRunning(false);
      setReconPhase("");
    }
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
        {`+ ${parentNode ? "Add subfolder" : "Add folder"}`}
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

      {/* Tab bar: two library views + the term-store reconciliation provisioner */}
      <div style={s.toggleWrap}>
        <div style={s.seg}>
          {(["Staging", "Documents", "Reconciliation", "GroupMap"] as Tab[]).map((t, i, arr) => (
            <button key={t}
              onClick={() => {
                setTab(t);
                setReconConfirm(false);
                if (t !== "Reconciliation" && t !== "GroupMap") { setLibTarget(t as LibTarget); setExpandedIds({}); }
              }}
              style={{ ...s.segBtn, ...(i === arr.length - 1 ? { borderRight: "none" } : {}), ...(tab === t ? s.segActive : {}) }}
            >
              {t === "Reconciliation" ? "Folder Reconciliation" : t === "GroupMap" ? "Group Map" : t}
            </button>
          ))}
        </div>
      </div>

      {tab === "GroupMap" ? (
        <GroupMapBuilder context={context} siteUrl={siteUrl} />
      ) : tab === "Reconciliation" ? (
        <div>
          <p style={{ fontSize: 13, color: "#444", lineHeight: 1.5, margin: "0 0 16px" }}>
            Build the folder tree from the <strong>term store</strong> in both{" "}
            <strong>Staging</strong> and <strong>Documents</strong>. Every level
            (segment → department → unit) is created if missing and{" "}
            <strong>locked</strong> (inheritance broken) so nobody can see a folder
            until its group is assigned. Groups are then{" "}
            <strong>auto-assigned from DMS Group Map</strong> by role (MEMBER → Read,
            UPL → Contribute, APR → Design; Documents gets viewer/MEMBER groups only).
            Under each unit the full <strong>Year × Document Type</strong> grid is
            pre-created (inheriting the unit folder permissions). Staging folders are
            also mapped for rename-proof upload routing. Safe to re-run: existing
            folders and mappings are not duplicated, and role assignments merge. A
            tier with no group-map rows is flagged in the log. Note: a full run can
            take a minute or two.
          </p>
          {reconConfirm ? (
            <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "4px 0 8px", padding: "8px 10px", background: "#f0f7f2", border: "1px solid #cfe4d8", borderRadius: 4, fontSize: 12, color: "#0f6c3f" }}>
              <span>Create + lock the term-store folder tree in <strong>Staging</strong> and <strong>Documents</strong>, then auto-assign groups from DMS Group Map by role. Safe to re-run.</span>
              <button style={{ ...s.btn, padding: "5px 14px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", flexShrink: 0 }} disabled={busy} onClick={() => { runReconciliation().catch(() => undefined); }}>
                {busy ? "Running…" : "Yes, run reconciliation"}
              </button>
              <button style={s.ghostBtn} disabled={busy} onClick={() => setReconConfirm(false)}>Cancel</button>
            </div>
          ) : (
            <button
              onClick={() => setReconConfirm(true)}
              disabled={busy}
              style={{ ...s.btn, background: "#0f6c3f", color: "#fff", border: "none" }}
            >
              Run reconciliation
            </button>
          )}

          {(reconRunning || folderFeed.length > 0 || assignFeed.length > 0) && (
            <div style={{ marginTop: 18 }}>
              <style>{"@keyframes fmspin{to{transform:rotate(360deg)}}"}</style>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                {reconRunning && (
                  <span style={{ display: "inline-block", width: 16, height: 16, border: "2px solid #cfe4d8", borderTopColor: "#0f6c3f", borderRadius: "50%", animation: "fmspin 0.8s linear infinite" }} />
                )}
                <strong style={{ fontSize: 13, color: "#0f6c3f" }}>
                  {reconRunning ? (reconPhase || "Working…") : "Reconciliation complete"}
                </strong>
                <span style={{ fontSize: 12, color: "#666" }}>
                  {reconCounts.folders} folders · {reconCounts.assigns} group assignments
                </span>
              </div>
              <div style={{ display: "flex", gap: 12 }}>
                {([
                  { title: "Folders", feed: folderFeed },
                  { title: "Group assignments", feed: assignFeed },
                ] as const).map((panel) => (
                  <div key={panel.title} style={{ flex: 1, minWidth: 0, border: "1px solid #e5e5e5", borderRadius: 4, overflow: "hidden" }}>
                    <div style={{ padding: "6px 10px", background: "#f7f7f7", fontSize: 12, fontWeight: 600, color: "#444", borderBottom: "1px solid #eee" }}>{panel.title}</div>
                    <div style={{ maxHeight: 260, overflowY: "auto", padding: "4px 0" }}>
                      {panel.feed.length === 0 ? (
                        <div style={{ padding: "6px 10px", fontSize: 12, color: "#aaa" }}>—</div>
                      ) : (
                        panel.feed.map((it, i) => (
                          <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "3px 10px", fontSize: 12, color: it.status === "admin" ? "#b45309" : "#333" }}>
                            {progIcon(it.status)}
                            <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{it.text}</span>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      ) : loading ? (
        <p style={{ fontSize: 13, color: "#666" }}>Loading folders…</p>
      ) : (
        <>
          {sections.length === 0 && (tree[NEW_TOP_LEVEL] ?? []).length === 0 && (
            <p style={{ fontSize: 13, color: "#999" }}>
              No top-level folders under {libTarget} yet — create the first one below.
            </p>
          )}

          {sections.map(mode => {
            const open = modeOpen[mode] !== false; // sections default to open
            const nodes = tree[mode] ?? [];
            return (
              <div key={mode} style={{ marginBottom: 28 }}>
                <div style={s.secHeader} onClick={() => setModeOpen(prev => ({ ...prev, [mode]: open ? false : true }))}>
                  <span style={s.ico}>{open ? "▾" : "▸"}</span>
                  <p style={s.secTitle}>{mode}</p>
                  {nodes.length > 0 && <span style={s.badge}>{nodes.length}</span>}
                </div>

                {open && (
                  <>
                    {nodes.length === 0 ? (
                      <p style={{ fontSize: 13, color: "#999" }}>No folders found under {libTarget}/{mode}.</p>
                    ) : (
                      <div style={s.scrollPane}>
                        {nodes.map(node => renderNode(node, mode, 0))}
                      </div>
                    )}

                    {libRoot && renderAddForm(mode, null)}
                  </>
                )}
              </div>
            );
          })}

          {/* New top-level folders staged for creation directly under the library root */}
          {(tree[NEW_TOP_LEVEL] ?? []).length > 0 && (
            <div style={{ marginBottom: 20 }}>
              <p style={s.secTitle}>New top-level folder{(tree[NEW_TOP_LEVEL] ?? []).length !== 1 ? "s" : ""}</p>
              <div style={s.scrollPane}>
                {(tree[NEW_TOP_LEVEL] ?? []).map(node => renderNode(node, NEW_TOP_LEVEL, 0))}
              </div>
            </div>
          )}

          {libRoot && (
            <div style={{ paddingTop: 4 }}>
              <button style={s.addFolderBtn} disabled={busy} onClick={addTopLevel}>
                + Add top-level folder
              </button>
            </div>
          )}
        </>
      )}

      {tab !== "Reconciliation" && (
        <div style={s.actions}>
          <button onClick={() => loadTree().catch(() => undefined)} disabled={busy || loading}
            style={{ ...s.btn, marginRight: "auto", background: "#fff", color: "#0f6c3f", border: "1px solid #0f6c3f" }}>
            Refresh
          </button>
          <button onClick={() => { handleUpdate().catch(() => undefined); }}
            disabled={busy || loading || !hasChanges}
            style={{ ...s.btn, background: !busy && !loading && hasChanges ? "#0f6c3f" : "#9bbfaa", color: "#fff", border: "none", cursor: !busy && !loading && hasChanges ? "pointer" : "default" }}>
            {busy ? "Updating…" : "Update"}
          </button>
        </div>
      )}

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
