import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { searchSiteGroups, fetchAllSiteGroups, getGroupMembers, addGroupMember } from "../../../shared/spGroups";
import { SITE_ENTRY_GROUP_NAME } from "../../../shared/groupMapModel";
import { IFolderManagerProps } from "./IFolderManagerProps";
import GroupMapBuilder from "./GroupMapBuilder";
import {
  loadFolderMapRows,
  FolderMapRow,
  resolveFolderByPath,
  probeFolderByPath,
  probeFolderById,
  writeFolderMapping,
  updateFolderMapping,
  renameFolder,
  deleteFolderMapRow,
  ensureFolder,
  encodeServerRelativePath,
} from "../../../shared/dmsFolderMap";
import { sanitizeFolderSegment, parseReconModes, RawModeRow } from "../../../shared/formModel";
import {
  ABBREV_LIST,
  AbbrevCollision,
  AbbrevRow,
  AbbrevTarget,
  buildAbbrevIndex,
  findCollisions,
  lookupAbbrev,
} from "../../../shared/folderAbbreviation";
import { FULL_NAME_COLUMN_TITLE, pickFullNameField, SpFieldLite } from "../../../shared/folderFullName";

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
  { key: "gho", termSetGuid: "08dd94cb-f76c-431c-9b37-e9c98f739ffc", stagingFolder: "Group Head Office" },
  // PLACEHOLDER — Upstream Malaysia is NOT onboarded (client scope 2026-07-29 is the other
  // three head offices). This GUID pre-dates the 2026-07-29 term-set rebuild and is stale.
  // Inert: with no `mode` row in DMS Config the segment is never offered. Replace the GUID
  // when the client creates the term set — nothing else needs to change.
  { key: "upstream_my_ho", termSetGuid: "16a52947-57a3-4217-9a49-b48cb8b0dd31", stagingFolder: "Upstream Malaysia Head Office" },
  { key: "minamas_ho", termSetGuid: "9ad00b00-a43c-4a8b-a39a-d0efa89ba706", stagingFolder: "Minamas Head Office" },
  { key: "nbpol_ho", termSetGuid: "77c3993b-0c3c-4a18-89d9-d69209886322", stagingFolder: "NBPOL Head Office" },
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

/**
 * How much of the Year × Document Type grid to pre-create under each unit folder.
 * Set live from DMS Config (`recon_gridMode` setting row) — no redeploy to change.
 *
 *   off         — create nothing. The upload form ensure-creates Year/Document Type on
 *                 first upload, and the Auto-route flow creates them in Documents as
 *                 approved files land. Viewers then see only folders that hold documents.
 *   currentYear — pre-create the current year's row only.
 *   full        — every year × every document type (the original behaviour).
 *
 * Why `off` is the default: the grid is ~98% of a full run. With 3 years and 20 document
 * types it is 63 folders per unit per library — 16,758 operations for 133 units, turning a
 * ~6 minute run into ~4h20m, to produce folders that are overwhelmingly empty and that
 * make browsing NOISIER for viewers. The structural folders that actually carry
 * permissions are only ~360 operations.
 */
type GridMode = "off" | "currentYear" | "full";
const DEFAULT_GRID_MODE: GridMode = "off";

// One folder the provisioner will ensure exists + lock. termGuid is null for the
// segment container folder (not a term); term folders (department, unit, …) carry
// their GUID so Staging can be mapped for rename-proof routing. isLeaf marks the
// deepest terms (upload targets) that get the Year × Document Type grid beneath.
// assignTerm is the term used to look up DMS Group Map rows for THIS folder's tier:
// the term-set GUID for the segment container, the term GUID for dept/unit folders.
// fullName is the RAW term label for this one folder, written to the Full Name column
// so the abbreviated folder still says what it is. It is deliberately not `label`:
// label is a breadcrumb ("GHO > Group Finance > Treasury") built for the log and the
// map row Title, and putting a breadcrumb in a per-folder column would repeat the
// whole path on every row.
type ProvTarget = { termGuid: string | null; assignTerm: string; relPath: string; label: string; fullName: string; section: string; isLeaf: boolean };

// DMS Group Map role → SharePoint permission level. GLOBAL is a privileged
// uploader bypass (not folder-scoped) and is never assigned to a folder.
const ROLE_TO_PERMISSION: Record<string, string> = {
  MEMBER: "Read",
  UPL: "Contribute",
  APR: "Design",
};

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

// Throttle-safety tuning for reconciliation. These are the DEFAULTS/fallbacks —
// they can be overridden live via DMS Config `setting` rows (recon_writeDelayMs /
// recon_batchSize / recon_cooldownMs) so the client can tune speed without a redeploy.
// withThrottleRetry() is the safety net: any write that still hits a 429/503 backs
// off and retries, so a shorter delay is safe — an occasional throttle self-heals.
// Lower delay = faster; too low risks 429s whose Retry-After penalty can be large,
// so ~150ms is a sensible aggressive floor. Grid is sequential (the old parallel
// burst was what tripped the throttle originally).
const RECON_WRITE_DELAY_MS = 200;    // pause between folder/permission writes (was 500)
const RECON_BATCH_SIZE      = 300;   // writes before an automatic cooldown (was 150)
const RECON_COOLDOWN_MS     = 2500;  // cooldown length, masked in the UI as "work" (was 4000)
const RECON_EST_HTTP_MS     = 250;   // rough per-write network+server time, on top of the delay
                                     // (used only for the up-front estimate before a live rate exists)

// Human-friendly duration: "45s" or "3m 07s".
const fmtDur = (ms: number): string => {
  const s = Math.max(0, Math.round(ms / 1000));
  if (s < 60) return `${s}s`;
  const sec = s % 60;
  return `${Math.floor(s / 60)}m ${sec < 10 ? "0" : ""}${sec}s`;
};
// Rotating status text shown during the cooldown so the pause reads as progress.
const COOLDOWN_MESSAGES = ["Discombobulating…", "Generating folders…"];

// Retry a SharePoint request on transient throttling, honoring Retry-After (seconds)
// with exponential backoff. Returns the final response (ok or not) after up to 5 retries.
// Beyond the obvious 429/503, SharePoint Online's anti-abuse system can surface a
// sustained WRITE burst as `403 E_ACCESSDENIED` / UnauthorizedAccessException even when
// the caller has full permissions — indistinguishable at the status line from a real
// denial except by the burst context. We treat that signature as retryable but only for
// the first few attempts, so a GENUINE 403 still fails fast (after ~3 backoffs) instead
// of masquerading as success forever.
async function withThrottleRetry(
  doPost: () => Promise<SPHttpClientResponse>,
  onWait?: (ms: number) => void,
): Promise<SPHttpClientResponse> {
  for (let attempt = 0; ; attempt++) {
    const res = await doPost();
    const throttled = res.status === 429 || res.status === 503;
    let burst403 = false;
    if (res.status === 403 && attempt < 3) {
      // Peek at a CLONE so the original body stays readable for the caller.
      const body = await res.clone().text().catch(() => "");
      burst403 = /E_ACCESSDENIED|UnauthorizedAccessException/i.test(body);
    }
    if ((!throttled && !burst403) || attempt >= 5) return res;
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
  // ETA: planned total throttled ops, run start time, and a ticking "now" so the
  // elapsed/remaining estimate repaints every second even between ops.
  const [reconPlanned, setReconPlanned] = useState(0);
  const [reconStartMs, setReconStartMs] = useState<number | undefined>(undefined);
  const [reconNow,     setReconNow]     = useState(0);

  useEffect(() => {
    if (!reconRunning) return;
    const id = setInterval(() => setReconNow(Date.now()), 1000);
    return () => clearInterval(id);
  }, [reconRunning]);

  // Guard against losing a run to a stray refresh or tab close. Reconciliation executes
  // entirely in this page — there is no server-side job — so navigating away stops it
  // mid-operation. Work already committed survives and a re-run resumes safely, but a
  // folder interrupted between "created" and "inheritance broken" is briefly left
  // INHERITING its parent's permissions until the next run repairs it. Worth a prompt.
  useEffect(() => {
    if (!reconRunning) return;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      // Browsers show their own wording and ignore ours, but a non-empty returnValue is
      // still what triggers the prompt at all.
      e.returnValue = "Reconciliation is still running. Leaving now will stop it.";
      return e.returnValue;
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [reconRunning]);

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
    // Retry-wrapped: a single throttled GET here used to null out the WHOLE library
    // pass in reconciliation (e.g. Documents skipped → no MEMBER grants land). Log the
    // real HTTP status on genuine failure rather than reporting a false "not found"
    // (CLAUDE.md gotcha #9 — a transient status is not missing data).
    const res: SPHttpClientResponse = await withThrottleRetry(() => context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lib)}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    ), reconWaitNote);
    if (!res.ok) {
      // eslint-disable-next-line no-console
      console.warn(`getLibraryRoot('${lib}') failed: HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
      return null;
    }
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
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments?$expand=Member,RoleDefinitionBindings&@f='${encodeServerRelativePath(folderPath)}'`,
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
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ), reconWaitNote);
    if (!res.ok) throw new Error(`breakroleinheritance HTTP ${res.status}`);
  };

  // Recycles the folder (and everything inside it) to the site Recycle Bin —
  // recoverable, not a permanent delete.
  const deleteFolder = async (path: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/recycle()?@f='${encodeServerRelativePath(path)}'`,
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
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})?@f='${encodeServerRelativePath(path)}'`,
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
    // THROWS on failure — do NOT soften this to `return []`. A silently empty child
    // list truncates the term tree, and prune reads "term not enumerated" as "term
    // deleted". Swallowing this error would let a throttle wipe a segment's map rows.
    if (!res.ok) throw new Error(`Term children ${parentId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({ id: t.id, label: t.labels[0].name }));
  };

  // Name of a term SET (not a term), used as the full name of the segment container
  // folder. Soft-fails to undefined: a missing label is cosmetic, and unlike the term
  // TREE this value feeds nothing that prune or routing depends on.
  const loadTermSetName = async (termSetId: string): Promise<string | undefined> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return undefined;
      const data = await res.json();
      const name = (data.localizedNames ?? [])[0]?.name;
      return typeof name === "string" && name.trim() !== "" ? name : undefined;
    } catch {
      return undefined;
    }
  };

  // Resolve the Full Name column's INTERNAL name on a library, live. Never guessed —
  // see the folderFullName module for why the display name does not determine it.
  // Returns undefined when the column has not been added to that library yet, which
  // is a soft state: the run proceeds and simply writes no full names.
  const loadFullNameField = async (lib: LibTarget): Promise<string | undefined> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lib)}')/fields` +
          `?$select=Title,InternalName,TypeAsString,ReadOnlyField` +
          `&$filter=Title eq '${FULL_NAME_COLUMN_TITLE.replace(/'/g, "''")}'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return undefined;
      const data = await res.json();
      return pickFullNameField((data.value ?? []) as SpFieldLite[]);
    } catch {
      return undefined;
    }
  };

  // Read the current Full Name on a folder's list item, so an unchanged value costs a
  // cheap GET instead of a throttled write. Undefined means "unknown" (empty, or the
  // read failed) and the caller writes.
  const getFolderFullName = async (
    serverRelativeUrl: string,
    internalName: string,
  ): Promise<string | undefined> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
          `?@f='${encodeServerRelativePath(serverRelativeUrl)}'&$select=${internalName}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return undefined;
      const data = await res.json();
      const v = data[internalName];
      return typeof v === "string" && v !== "" ? v : undefined;
    } catch {
      return undefined;
    }
  };

  // Write the full term label onto a folder's list item. Throws on failure so the
  // caller can report it against the folder rather than losing it silently.
  const setFolderFullName = async (
    serverRelativeUrl: string,
    internalName: string,
    value: string,
  ): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
        `?@f='${encodeServerRelativePath(serverRelativeUrl)}'`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          // MERGE leaves every other field alone; IF-MATCH:* skips the etag round-trip.
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify({ [internalName]: value }),
      },
    );
    // validateUpdateListItem returns 200 on field errors, but a plain MERGE does not —
    // it 4xxs. Still surface the body: "field does not exist" and "read-only" read
    // identically as a bare status code.
    if (!res.ok) {
      const detail = (await res.text().catch(() => "")).slice(0, 200);
      throw new Error(`HTTP ${res.status} ${detail}`);
    }
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

  /**
   * Term GUID → folder-name abbreviation.
   *
   * Lives here rather than in the shared module because a shared file that imports
   * `@microsoft/sp-http` cannot be unit tested — see the note on ABBREV_LIST.
   *
   * Deliberately NOT wrapped in a catch returning an empty map: an unreadable list
   * must abort the run. An empty index makes every term look unmapped, which would
   * skip every folder and report 175 false "needs attention" rows.
   */
  const loadAbbreviations = async (): Promise<Map<string, string>> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(ABBREV_LIST)}')/items?$select=TermGuid,Abbreviation&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`${ABBREV_LIST} read failed: HTTP ${res.status}. ${body}`);
    }
    const data = await res.json();
    const rows: AbbrevRow[] = ((data.value ?? []) as Array<{ TermGuid?: string; Abbreviation?: string }>).map(
      (r) => ({ termGuid: r.TermGuid ?? "", abbreviation: r.Abbreviation ?? "" }),
    );
    return buildAbbrevIndex(rows);
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
  const loadReconGridTermSets = async (): Promise<{ year: string; docType: string; gridMode: GridMode }> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return { year: YEAR_TERMSET, docType: DOCTYPE_TERMSET, gridMode: DEFAULT_GRID_MODE };
      const data = await res.json();
      const map: Record<string, string> = {};
      ((data.value ?? []) as Array<{ Title: string; SettingValue: string }>).forEach(
        (item) => { map[item.Title] = (item.SettingValue ?? "").trim(); },
      );
      const raw = (map.recon_gridMode || "").toLowerCase();
      const gridMode: GridMode =
        raw === "full" ? "full"
        : raw === "currentyear" ? "currentYear"
        : raw === "off" ? "off"
        : DEFAULT_GRID_MODE;
      return {
        year: map.termSet_yearPeriod || YEAR_TERMSET,
        docType: map.termSet_documentType || DOCTYPE_TERMSET,
        gridMode,
      };
    } catch {
      return { year: YEAR_TERMSET, docType: DOCTYPE_TERMSET, gridMode: DEFAULT_GRID_MODE };
    }
  };

  // Reconciliation speed knobs, tunable live from DMS Config `setting` rows without a
  // redeploy: recon_writeDelayMs / recon_batchSize / recon_cooldownMs. Falls back to
  // the module defaults. Lets the client dial the delay down to find their tenant's
  // throttle floor (withThrottleRetry catches any 429 that slips through).
  const loadReconThrottle = async (): Promise<{ delayMs: number; batchSize: number; cooldownMs: number }> => {
    const fallback = { delayMs: RECON_WRITE_DELAY_MS, batchSize: RECON_BATCH_SIZE, cooldownMs: RECON_COOLDOWN_MS };
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return fallback;
      const data = await res.json();
      const map: Record<string, string> = {};
      ((data.value ?? []) as Array<{ Title: string; SettingValue: string }>).forEach(
        (item) => { map[item.Title] = (item.SettingValue ?? "").trim(); },
      );
      const num = (v: string | undefined, d: number): number => {
        const n = Number(v);
        return Number.isFinite(n) && n > 0 ? n : d;
      };
      return {
        delayMs: num(map.recon_writeDelayMs, fallback.delayMs),
        batchSize: num(map.recon_batchSize, fallback.batchSize),
        cooldownMs: num(map.recon_cooldownMs, fallback.cooldownMs),
      };
    } catch {
      return fallback;
    }
  };

  // Flatten the term store into the folders to provision, parent-before-child so a
  // parent always exists before we create its child. The segment container folder
  // (mode.stagingFolder) has no term; every term below it carries its GUID.
  /**
   * Enumerate every folder target from the term store.
   *
   * Also reports which segments could NOT be fully enumerated. This matters far more
   * than it looks: prune decides a row is orphaned by asking "is its term in the
   * enumerated set?", so a segment that silently came back empty would make every one
   * of its rows look deleted. Callers MUST treat a non-empty `incomplete` as a reason
   * to skip pruning.
   */
  const buildProvisionTargets = async (): Promise<{
    targets: ProvTarget[];
    incomplete: string[];
    missingAbbrev: Array<{ termGuid: string; label: string }>;
    collisions: AbbrevCollision[];
  }> => {
    const out: ProvTarget[] = [];
    const incomplete: string[] = [];
    // Terms with no abbreviation are skipped, not guessed at, and reported here.
    const missingAbbrev: Array<{ termGuid: string; label: string }> = [];
    // Every named target, so siblings sharing an abbreviation can be caught before
    // a single folder is created.
    const abbrevTargets: AbbrevTarget[] = [];
    const abbrevIndex = await loadAbbreviations();
    const modes = await loadReconModes();
    for (const mode of modes) {
      // Segment container: not a mapped term, but groups target it via the term-set GUID.
      // Its full name is the TERM SET's name, read live. mode.stagingFolder cannot serve
      // here — in DMS Config it already holds the abbreviation ("GHO"), which is exactly
      // the string Full Name exists to explain. Falls back to the folder name if the read
      // fails, so a term-store hiccup costs a label, not the run.
      const segmentFullName = (await loadTermSetName(mode.termSetGuid)) ?? mode.stagingFolder;
      out.push({ termGuid: null, assignTerm: mode.termSetGuid, relPath: `/${mode.stagingFolder}`, label: mode.stagingFolder, fullName: segmentFullName, section: mode.stagingFolder, isLeaf: false });
      // A failure anywhere in this segment's tree marks the WHOLE segment incomplete.
      // Targets gathered before the failure are kept (creating a subset of folders is
      // harmless and idempotent) — but prune must not run against a partial picture.
      try {
      const tops = await loadReconTops(mode.termSetGuid);
      for (const top of tops) {
        // Term labels are NOT safe as folder names. A "/" is the worst case — it is both
        // rejected by SharePoint (HTTP 400, SPException -2130575245) and read as a path
        // separator, so it silently implies an extra folder level. Real example: the
        // Minamas unit "Value Creation / Value Transformation".
        // Path segments are sanitised; the DISPLAY label and the map row Title keep the
        // raw term text, so the log and the index still read like the term store.
        // Folder names come from the abbreviation list, NOT the term label. The
        // labels are long and their fullwidth ampersands cost 9 encoded characters
        // each, which put the worst-case path within 73 characters of the ~330
        // limit where GetFolderByServerRelativeUrl starts returning 400.
        // See the 2026-07-30 folder-abbreviation-naming spec.
        //
        // sanitizeFolderSegment still applies: a "/" in a name is both rejected by
        // SharePoint and read as a path separator.
        //
        // A term with no abbreviation is SKIPPED and reported, never guessed at.
        // Falling back to the label would create a folder at a path the next run
        // does not expect, and uploads resolve by UniqueId so nobody would notice.
        const seg = (termGuid: string, label: string): string | undefined => {
          const abbrev = lookupAbbrev(abbrevIndex, termGuid);
          if (abbrev === undefined) {
            missingAbbrev.push({ termGuid, label });
            return undefined;
          }
          return sanitizeFolderSegment(abbrev) || abbrev;
        };
        const topSeg = seg(top.id, top.label);
        if (topSeg === undefined) continue; // reported; its children are unreachable
        abbrevTargets.push({ parentPath: `/${mode.stagingFolder}`, termGuid: top.id, abbreviation: topSeg, label: top.label });
        const topTarget: ProvTarget = { termGuid: top.id, assignTerm: top.id, relPath: `/${mode.stagingFolder}/${topSeg}`, label: `${mode.stagingFolder} > ${top.label}`, fullName: top.label, section: mode.stagingFolder, isLeaf: false };
        out.push(topTarget);
        // Recurse; returns whether the term had children. A term with no children
        // is a leaf (the upload target) and gets the Year × Document Type grid.
        // `ancestors` carries raw labels for display, `pathAncestors` the sanitised
        // segments for the folder path — they can differ and must not be conflated.
        const walk = async (parentId: string, ancestors: string[], pathAncestors: string[]): Promise<boolean> => {
          const children = await loadReconChildren(mode.termSetGuid, parentId);
          for (const child of children) {
            const childSeg = seg(child.id, child.label);
            if (childSeg === undefined) continue; // reported; skip this subtree
            const chain = [...ancestors, child.label];
            const pathChain = [...pathAncestors, childSeg];
            const parentPath = `/${mode.stagingFolder}/${topSeg}${pathAncestors.length > 0 ? "/" + pathAncestors.join("/") : ""}`;
            abbrevTargets.push({ parentPath, termGuid: child.id, abbreviation: childSeg, label: child.label });
            const childTarget: ProvTarget = {
              termGuid: child.id,
              assignTerm: child.id,
              relPath: `/${mode.stagingFolder}/${topSeg}/${pathChain.join("/")}`,
              label: `${mode.stagingFolder} > ${top.label} > ${chain.join(" > ")}`,
              fullName: child.label,
              section: mode.stagingFolder,
              isLeaf: false,
            };
            out.push(childTarget);
            childTarget.isLeaf = !(await walk(child.id, chain, pathChain));
          }
          return children.length > 0;
        };
        topTarget.isLeaf = !(await walk(top.id, [], []));
      }
      } catch (e) {
        incomplete.push(`${mode.stagingFolder} — ${(e as Error).message}`);
      }
    }
    return { targets: out, incomplete, missingAbbrev, collisions: findCollisions(abbrevTargets) };
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
    setReconPlanned(0);
    setReconStartMs(undefined);
    setReconPhase("Generating folders…");
    const entries: LogEntry[] = [];
    // Throttle governor: pause between writes, and auto-cooldown every N writes. The
    // cooldown is invisible — the spinner keeps running and the status text rotates so
    // it reads as continuous work (no user click needed to resume). Speed knobs come
    // from DMS Config (recon_writeDelayMs / recon_batchSize / recon_cooldownMs) so the
    // client can tune without a redeploy; withThrottleRetry catches any 429 that slips.
    const throttle = await loadReconThrottle();
    let opCount = 0;
    const tick = async (): Promise<void> => {
      opCount++;
      await sleep(throttle.delayMs);
      if (opCount % throttle.batchSize === 0) {
        for (const msg of COOLDOWN_MESSAGES) {
          setReconPhase(msg);
          await sleep(throttle.cooldownMs / COOLDOWN_MESSAGES.length);
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
      // Full rows, not just term GUIDs: the presence of a row is NOT proof the row is
      // still correct, so each one gets verified below. See the folder-map-integrity spec.
      const mapRows = await loadFolderMapRows(context.spHttpClient, siteUrl);
      const mapByTerm = new Map<string, FolderMapRow>();
      for (const r of mapRows) {
        if (r.termGuid) mapByTerm.set(r.termGuid.toLowerCase(), r);
      }
      // Each folder's name BEFORE this run, snapshotted once.
      //
      // Read live from the row instead and the Staging pass — which rewrites
      // FolderUrl after a successful rename — would make the later Documents pass
      // believe the name was already correct. Documents would skip the rename and
      // the create step would make an empty folder at the new name, stranding the
      // real one under the old name. Snapshotting keeps every library renaming
      // from the same origin regardless of pass order.
      const oldNameByTerm = new Map<string, string>();
      for (const r of mapRows) {
        const name = (r.folderUrl ?? "").split("/").pop() ?? "";
        if (r.termGuid && name) oldNameByTerm.set(r.termGuid.toLowerCase(), name);
      }
      const groupMap = await loadGroupMapForAssign();
      const { targets, incomplete: incompleteSegments, missingAbbrev, collisions } = await buildProvisionTargets();
      // A collision aborts BEFORE anything is created. Two siblings resolving to
      // one path means one folder, one ACL, and two units' documents inside it —
      // the isolation the whole permission model rests on. A partial run would
      // create that merged folder before anyone read the log.
      if (collisions.length > 0) {
        for (const c of collisions) {
          entries.push({ msg: `✖ COLLISION in ${c.parentPath}: "${c.abbreviation}" is used by ${c.labels.join(" | ")}`, ok: false });
        }
        entries.push({ msg: `Nothing was created. Give each of these a distinct abbreviation in ${ABBREV_LIST}, then run again.`, ok: false });
        setLog((prev) => [...prev, ...entries]);
        showToast(`${collisions.length} abbreviation collision(s) — nothing was created.`, false);
        setBusy(false);
        return;
      }
      // Not fatal: every other term still provisions. But it must be loud, because
      // a unit with no folder has no map row and its uploaders are blocked.
      for (const m of missingAbbrev) {
        entries.push({ msg: `⚠ SKIPPED (no abbreviation): ${m.label} — add a row to ${ABBREV_LIST} for term ${m.termGuid}`, ok: false });
      }
      // Surface enumeration failures as real errors. Without this they were invisible:
      // the segment just produced no targets, the run looked clean, and prune would
      // then delete every map row for it. See the prune guard below.
      for (const seg of incompleteSegments) {
        entries.push({ msg: `⚠ could not fully read the term store for ${seg} — folders may be missing and pruning is disabled for this run`, ok: false });
      }
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
      let yearLabels: string[] = [];
      let docTypeLabels: string[] = [];
      if (gridSets.gridMode === "off") {
        // Nothing to pre-create. Year/Document Type folders are made on demand by the
        // upload form, and in Documents by the Auto-route flow as approved files land.
        entries.push({ msg: `Year × Document Type grid: skipped (recon_gridMode = off) — folders are created on first use`, ok: true });
      } else {
        yearLabels    = (await loadReconTops(gridSets.year).catch(() => [] as TermLite[])).map(y => sanitizeFolderSegment(y.label)).filter(Boolean);
        docTypeLabels = (await loadReconTops(gridSets.docType).catch(() => [] as TermLite[])).map(d => sanitizeFolderSegment(d.label)).filter(Boolean);
        if (gridSets.gridMode === "currentYear" && yearLabels.length > 0) {
          const thisYear = String(new Date().getFullYear());
          const match = yearLabels.filter((y) => y === thisYear);
          // No term for the current year (e.g. the client has not added 2027 yet) — fall
          // back to the LAST year in the set rather than silently building all of them.
          yearLabels = match.length > 0 ? match : yearLabels.slice(-1);
        }
        entries.push({
          msg: `Year × Document Type grid: ${gridSets.gridMode} — ${yearLabels.length} year(s) × ${docTypeLabels.length} document type(s) per unit`,
          ok: true,
        });
      }
      // Estimate the workload up front: count every throttled op (each incurs the
      // inter-write delay). Structural folder (1) + Year×DocType grid per leaf +
      // applicable group grants per lib. Ancestor browse grants aren't throttled, so
      // they're excluded — the live rate absorbs their real time. Worst case (assumes
      // nothing exists yet); re-runs finish faster as existing folders skip.
      const gridPerLeaf = yearLabels.length > 0 ? yearLabels.length * (1 + docTypeLabels.length) : 0;
      // Resolved once per library, before the estimate, because whether the column
      // exists changes the op count. Absent on a library = that library gets no full
      // names and is told so once, rather than once per folder.
      const fullNameFields = new Map<LibTarget, string>();
      for (const lib of ["Staging", "Documents"] as LibTarget[]) {
        const f = await loadFullNameField(lib);
        if (f) fullNameFields.set(lib, f);
        // ok:true deliberately. This is a warning, not an error: `errorsBeforePrune`
        // counts !ok entries and blocks the orphan prune, and that guard exists because
        // a partial TERM STORE read returns a short target list that makes healthy map
        // rows look deleted. A missing display column cannot shorten the target list, so
        // gating prune on it would silently disable self-healing over a cosmetic column.
        // The ⚠ still puts it in "Needs attention" where an admin will see it.
        else entries.push({ msg: `⚠ ${lib}: no usable "${FULL_NAME_COLUMN_TITLE}" column — folders will show only their abbreviation`, ok: true });
      }
      let plannedOps = 0;
      for (const lib of ["Staging", "Documents"] as LibTarget[]) {
        for (const t of targets) {
          plannedOps += 1;
          if (fullNameFields.has(lib)) plannedOps += 1;
          if (t.isLeaf) plannedOps += gridPerLeaf;
          const rows = groupMap.get(t.assignTerm.toLowerCase()) ?? [];
          plannedOps += rows.filter(g =>
            ROLE_TO_PERMISSION[g.role] !== undefined &&
            (lib === "Staging" ? g.role !== "MEMBER" : g.role === "MEMBER"),
          ).length;
        }
      }
      const reconStart = Date.now();
      setReconPlanned(plannedOps);
      setReconStartMs(reconStart);
      setReconNow(reconStart);
      // ── RENAME PASS ────────────────────────────────────────────────────────
      // Runs to completion across EVERY library before the create/lock loop, and
      // before anything writes to the Folder Map.
      //
      // Two earlier attempts put this inside the library loop and both failed the
      // same way. The Staging map row is the only record of a folder's previous
      // name, so renaming Staging and refreshing that row destroyed the very
      // information the Documents pass needed: Documents then saw "already
      // correct", skipped the rename, and the create step made an empty folder at
      // the new name while the real one kept the old one. Persisted, so re-running
      // could not repair it either.
      //
      // Renaming first, then updating the row once, removes the ordering entirely.
      // It also has to precede the create step regardless: `full` points at the new
      // name, so creating first would put an empty folder exactly where the rename
      // needs to land, and the rename would report a collision against it.
      //
      // Note this reverts a folder renamed by hand. Deliberate — the abbreviation
      // list is the single source of truth and every library must agree — and
      // harmless, because uploads resolve by UniqueId rather than by path.
      const renameLibs = ["Staging", "Documents"] as LibTarget[];
      const renameRoots = new Map<string, string>();
      for (const lib of renameLibs) {
        const r = await getLibraryRoot(lib);
        if (r) renameRoots.set(lib, r);
      }
      for (const t of targets) {
        if (!t.termGuid) continue; // segment containers have no term, so no abbreviation
        const wantName = t.relPath.split("/").pop() ?? "";
        const oldName = oldNameByTerm.get(t.termGuid.toLowerCase()) ?? "";
        // Case-insensitive: SharePoint treats sibling names as case-insensitive for
        // uniqueness, so CORU → Coru would collide with itself and report a
        // conflict that is not one.
        if (!oldName || !wantName || oldName.toLowerCase() === wantName.toLowerCase()) continue;
        const parentRel = t.relPath.slice(0, t.relPath.lastIndexOf("/"));
        let newStagingUrl = "";
        for (const lib of renameLibs) {
          const libRoot = renameRoots.get(lib);
          if (!libRoot) continue;
          const oldProbe = await probeFolderByPath(
            context.spHttpClient,
            siteUrl,
            `${libRoot}${parentRel}/${oldName}`,
          );
          if (!oldProbe.folder) continue; // nothing under the old name here
          const renamed = await renameFolder(
            context.spHttpClient,
            siteUrl,
            oldProbe.folder.serverRelativeUrl,
            wantName,
          );
          if (renamed.ok) {
            entries.push({ msg: `  ✎ ${lib}${parentRel}: renamed ${oldName} → ${wantName}`, ok: true });
            if (lib === "Staging") newStagingUrl = renamed.serverRelativeUrl ?? "";
            await tick();
          } else if (renamed.conflict) {
            // Two terms want one folder name. Forcing it would merge two units'
            // documents behind a single ACL — the isolation failure findCollisions
            // exists to prevent — so this needs a person.
            entries.push({
              msg: `  ⚠ ${t.label} — cannot rename "${oldName}" to "${wantName}" in ${lib}: a folder of that name is already there. Fix the abbreviation, then re-run.`,
              ok: false,
            });
          } else {
            entries.push({
              msg: `  ✗ ${t.label} — rename "${oldName}" → "${wantName}" in ${lib} FAILED (HTTP ${renamed.status}) ${renamed.detail ?? ""}`,
              ok: false,
            });
          }
        }
        // Refresh the stored path only now, with every library already handled. The
        // UniqueId is untouched by a rename, so the verification step later treats
        // the row as valid and would otherwise leave FolderUrl permanently stale —
        // it self-heals only when the folder is missing.
        const row = mapByTerm.get(t.termGuid.toLowerCase());
        if (newStagingUrl && row) {
          await updateFolderMapping(context.spHttpClient, siteUrl, row.itemId, {
            folderUniqueId: row.folderUniqueId,
            folderUrl: newStagingUrl,
            title: t.label,
          });
          row.folderUrl = newStagingUrl;
        }
      }

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
            // Full Name: the abbreviation is the folder's name, so the raw term label
            // goes on the item to keep the tree readable in the details pane.
            //
            // Written on every run, not only at creation, for two reasons: it backfills
            // the trees that already exist (this column arrives after they were built),
            // and it is how a RENAMED term's label catches up — the rename pass above
            // moves the folder when the ABBREVIATION changes, but a term relabelled in
            // the term store with its abbreviation untouched changes nothing on disk and
            // would otherwise keep its stale full name forever.
            //
            // The read guard keeps that cheap: an unchanged value costs one GET, which
            // is not throttled, instead of a write that is. On a settled tree this is
            // the normal case, so the extra pass is close to free.
            const fullNameField = fullNameFields.get(lib);
            if (fullNameField) {
              try {
                const current = await getFolderFullName(full, fullNameField);
                if (current !== t.fullName) {
                  await setFolderFullName(full, fullNameField, t.fullName);
                  entries.push({ msg: `  ↳ ${FULL_NAME_COLUMN_TITLE} = ${t.fullName}`, ok: true });
                  await tick();
                }
              } catch (e) {
                // Never fatal — a label failure must not stop the folder's ACL work,
                // which is the part that actually controls access. ok:true for the same
                // reason as the missing-column notice above: it must not gate the prune.
                entries.push({ msg: `  ⚠ ${folderLabel} — could not set ${FULL_NAME_COLUMN_TITLE}: ${(e as Error).message}`, ok: true });
              }
            }
            // Map Staging term folders only (the segment container has no term).
            if (lib === "Staging" && t.termGuid) {
              const existingRow = mapByTerm.get(t.termGuid.toLowerCase());
              if (!existingRow) {
                // Unmapped term → create the row.
                const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, full);
                if (resolved) {
                  await writeFolderMapping(context.spHttpClient, siteUrl, {
                    termGuid: t.termGuid,
                    folderUniqueId: resolved.uniqueId,
                    title: t.label,
                    folderUrl: resolved.serverRelativeUrl,
                    section: t.section,
                  });
                  entries.push({ msg: `  ↳ mapped ${t.label} → ${resolved.uniqueId}`, ok: true });
                }
              } else {
                // Mapped already — VERIFY the stored UniqueId instead of assuming it is
                // right. Ask whether THAT folder still exists; do NOT compare against
                // whatever sits at the term-label path. A renamed or moved folder keeps
                // its UniqueId, so it still resolves and the row is left alone — that is
                // the rename-proofing working. Comparing by path would see a difference
                // and repoint the row at a freshly created empty folder, abandoning the
                // real one and its documents.
                const probe = await probeFolderById(
                  context.spHttpClient,
                  siteUrl,
                  existingRow.folderUniqueId,
                );
                if (probe.folder) {
                  // Row still valid. Nothing to do — any rename already happened
                  // before the folder was created, further up this iteration.
                } else if (probe.confirmedMissing) {
                  // The mapped folder is gone (deleted, then recreated by this run or by
                  // hand). Repoint the row at the folder that is actually there now —
                  // this is the self-heal that the old skip-if-mapped logic prevented.
                  const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, full);
                  if (resolved) {
                    await updateFolderMapping(context.spHttpClient, siteUrl, existingRow.itemId, {
                      folderUniqueId: resolved.uniqueId,
                      folderUrl: resolved.serverRelativeUrl,
                      title: t.label,
                    });
                    existingRow.folderUniqueId = resolved.uniqueId;
                    existingRow.folderUrl = resolved.serverRelativeUrl;
                    entries.push({
                      msg: `  ↻ remapped ${t.label} — stale folder id replaced with ${resolved.uniqueId}`,
                      ok: true,
                    });
                  } else {
                    entries.push({
                      msg: `  ⚠ ${t.label} — mapped folder is gone and no folder found at ${full}; row left as-is`,
                      ok: false,
                    });
                  }
                } else {
                  // Throttle, permission error, malformed request — NOT evidence of
                  // deletion. Changing the row on this would be acting on a guess.
                  entries.push({
                    msg: `  ⚠ ${t.label} — could not verify mapped folder (HTTP ${probe.status}); row left unchanged`,
                    ok: false,
                  });
                }
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
            // Only a LEAF (unit) folder is expected to carry Group Map rows — the model is
            // leaf-only by design (see CLAUDE.md / the leaf-only authorization spec), so the
            // segment and department tiers having none is the correct state, not a problem.
            // Warning on them buried the real warnings under ~260 structurally unfixable
            // ones. Parent tiers stay admin-only and silent; members reach their unit
            // through the ancestor Read grants below.
            if (applicable.length === 0 && t.isLeaf) {
              entries.push({ msg: `  ⚠ ${folderLabel} — no group-map groups for this unit (locked admin-only)`, ok: true });
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
              // FAST PATH. The grid is built in order, so if the LAST year's LAST
              // document-type folder is there, the whole grid is there. One probe
              // replaces 63 round-trips per leaf per library. Without this a no-op
              // re-run cost the same as a first run (measured: 6,090 folders in
              // 111m for 48 units, of which only 6 were real group assignments).
              // Trade-off: a folder hand-deleted from the middle of a COMPLETE grid
              // is not restored here — the upload form ensure-creates it on demand.
              // probeFolderByPath retries 429s, so a throttle cannot fake "incomplete"
              // and trigger a needless full rebuild.
              const lastYear = yearLabels[yearLabels.length - 1];
              const lastDt = docTypeLabels.length > 0 ? docTypeLabels[docTypeLabels.length - 1] : undefined;
              const gridProbe = await probeFolderByPath(
                context.spHttpClient,
                siteUrl,
                lastDt ? `${full}/${lastYear}/${lastDt}` : `${full}/${lastYear}`,
              );
              if (gridProbe.folder) {
                entries.push({ msg: `  ↳ ${lib}${t.relPath} — Year × Document Type grid already complete (${gridTotal}), skipped`, ok: true });
                setLastFolder(`${t.label} grid: already complete, skipped`, "skip");
              } else {
              pushFolder(`${t.label} grid: 0 / ${gridTotal}`, "run");
              for (const yr of yearLabels) {
                const yearFolder = await ensureFolder(context.spHttpClient, siteUrl, full, yr);
                if (yearFolder) { grid++; bumpFolders(); }
                setLastFolder(`${t.label} grid: ${grid} / ${gridTotal}`, "run");
                // Only pace REAL writes. Charging the throttle delay to a folder that
                // already existed is what made a no-op re-run as slow as a first run.
                if (yearFolder?.created) await tick();
                if (!yearFolder) continue;
                // Sequential (NOT parallel) — the old parallel burst was what tripped the
                // 429 throttle. One folder at a time + the inter-write delay keeps it safe.
                for (const dt of docTypeLabels) {
                  const made = await ensureFolder(context.spHttpClient, siteUrl, yearFolder.serverRelativeUrl, dt);
                  if (made) { grid++; bumpFolders(); }
                  setLastFolder(`${t.label} grid: ${grid} / ${gridTotal}`, "run");
                  if (made?.created) await tick();
                }
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
      // Site-entry self-heal: ensure every member of any DMS_* group is also in
      // DMS_SITE_MEMBERS, so users added the native way (bypassing the web part's
      // auto-add) can still open the site. See site-entry-access-layer spec §7a.
      try {
        setReconPhase("Syncing site-entry group…");
        const allGroups = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
        const entryGroup = allGroups.find(
          (g) => g.title.trim().toLowerCase() === SITE_ENTRY_GROUP_NAME.toLowerCase(),
        );
        if (!entryGroup) {
          entries.push({ msg: `⚠ ${SITE_ENTRY_GROUP_NAME} not found — run "Set up site entry" first (site-entry sync skipped)`, ok: false });
        } else {
          const entryMembers = await getGroupMembers(context.spHttpClient, siteUrl, entryGroup.id);
          const already = new Set(entryMembers.map((m) => m.loginName.toLowerCase()));
          const dmsGroups = allGroups.filter(
            (g) => g.title.toUpperCase().indexOf("DMS_") === 0 && g.id !== entryGroup.id,
          );
          let healed = 0;
          for (const g of dmsGroups) {
            const members = await getGroupMembers(context.spHttpClient, siteUrl, g.id).catch(() => []);
            for (const m of members) {
              if (already.has(m.loginName.toLowerCase())) continue;
              try {
                await addGroupMember(context.spHttpClient, siteUrl, entryGroup.id, m.loginName);
                already.add(m.loginName.toLowerCase());
                healed++;
              } catch { /* skip a member that can't be added; not fatal */ }
            }
          }
          entries.push({
            msg: healed > 0
              ? `${SITE_ENTRY_GROUP_NAME}: added ${healed} member(s) missing site entry ✓`
              : `${SITE_ENTRY_GROUP_NAME}: all DMS group members already have site entry ✓`,
            ok: true,
          });
        }
      } catch (e) {
        entries.push({ msg: `Site-entry sync skipped — ${(e as Error).message}`, ok: false });
      }

      // ── Prune orphaned Folder Map rows ────────────────────────────────────────
      // A row whose TERM no longer exists is never visited by the loop above, so it
      // would linger forever. This is the only place that can see them: the run has
      // just enumerated every valid term.
      //
      // Deletes LIST ROWS ONLY, never folders. A row is derived data that a later run
      // can rebuild from the term store + the folder tree; the folder holds documents
      // that exist nowhere else. A folder left behind is reported so a human can decide.
      const errorsBeforePrune = entries.filter((e) => !e.ok).length;
      if (incompleteSegments.length > 0) {
        // Checked separately from the error count even though an entry was already
        // pushed above — this is THE condition prune must never run under, and it
        // should not depend on that entry still being there.
        entries.push({
          msg: `Prune skipped — the term store could not be fully read (${incompleteSegments.length} segment(s)); rows are never pruned against a partial term list`,
          ok: true,
        });
      } else if (errorsBeforePrune > 0) {
        // Hard guard. A half-failed term-store read returns a SHORT target list, and a
        // short list makes perfectly healthy rows look orphaned — pruning on that would
        // wipe the map wholesale. No clean run, no prune.
        entries.push({
          msg: `Prune skipped — run had ${errorsBeforePrune} error(s); orphaned map rows are only removed after a clean run`,
          ok: true,
        });
      } else {
        try {
          setReconPhase("Pruning orphaned folder map rows…");
          const enumerated = new Set(
            targets
              .filter((t) => t.termGuid)
              .map((t) => (t.termGuid as string).toLowerCase()),
          );
          let pruned = 0;
          for (const row of mapRows) {
            if (!row.termGuid) continue;
            if (enumerated.has(row.termGuid.toLowerCase())) continue; // term alive — handled above
            const probe = await probeFolderById(
              context.spHttpClient,
              siteUrl,
              row.folderUniqueId,
            );
            if (!probe.folder && !probe.confirmedMissing) {
              // Could not tell. Keep the row; never delete on a guess.
              entries.push({
                msg: `  ⚠ orphaned row "${row.title}" — could not verify its folder (HTTP ${probe.status}); row kept`,
                ok: false,
              });
              continue;
            }
            try {
              await deleteFolderMapRow(context.spHttpClient, siteUrl, row.itemId);
              pruned++;
              entries.push({
                msg: probe.folder
                  ? `  ✂ pruned "${row.title}" — term deleted from the term store. Its FOLDER still exists at ${probe.folder.serverRelativeUrl} and was NOT touched — delete it manually if that is intended.`
                  : `  ✂ pruned "${row.title}" — term and folder both gone`,
                ok: true,
              });
            } catch (e) {
              entries.push({
                msg: `  ✗ could not prune "${row.title}": ${(e as Error).message}`,
                ok: false,
              });
            }
          }
          entries.push({
            msg: pruned > 0
              ? `Folder Map: pruned ${pruned} orphaned row(s) ✓`
              : `Folder Map: no orphaned rows ✓`,
            ok: true,
          });
        } catch (e) {
          entries.push({ msg: `Prune skipped — ${(e as Error).message}`, ok: false });
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

      <h2 style={s.h2}>Folder &amp; Group Manager</h2>
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
              <span>
                Create + lock the term-store folder tree in <strong>Staging</strong> and <strong>Documents</strong>, then auto-assign groups from DMS Group Map by role. Safe to re-run.
                <br />
                <strong>Keep this tab open until it finishes.</strong> The run happens in your browser — refreshing, closing the tab or navigating away stops it partway. Nothing is lost and you can simply run it again, but do re-run before letting users in: a folder interrupted at the wrong moment stays unlocked until the next run.
              </span>
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
                  {reconRunning && (
                    <span style={{ marginLeft: 8, fontWeight: 400, color: "#8a5a00" }}>
                      — keep this tab open; leaving stops the run
                    </span>
                  )}
                </strong>
                <span style={{ fontSize: 12, color: "#666" }}>
                  {reconCounts.folders} folders · {reconCounts.assigns} group assignments
                </span>
                {(() => {
                  const completed = reconCounts.folders + reconCounts.assigns;
                  const elapsedMs = reconStartMs ? Math.max(0, reconNow - reconStartMs) : 0;
                  if (!reconRunning) {
                    // Final line after a run completes.
                    return reconStartMs ? (
                      <span style={{ fontSize: 12, color: "#666" }}>· took {fmtDur(elapsedMs)}</span>
                    ) : null;
                  }
                  const remainingOps = Math.max(0, reconPlanned - completed);
                  // Use the live rate once we have a few samples; before that, a static
                  // model from the delay + cooldown so a number shows immediately.
                  const remainMs = completed >= 5
                    ? remainingOps * (elapsedMs / completed)
                    : remainingOps * (RECON_WRITE_DELAY_MS + RECON_EST_HTTP_MS)
                      + Math.floor(reconPlanned / RECON_BATCH_SIZE) * RECON_COOLDOWN_MS;
                  return (
                    <span style={{ fontSize: 12, color: "#0f6c3f", fontWeight: 600 }}>
                      · ~{fmtDur(remainMs)} left
                      <span style={{ color: "#999", fontWeight: 400 }}> ({fmtDur(elapsedMs)} elapsed{reconPlanned > 0 ? `, ${completed}/${reconPlanned}` : ""})</span>
                    </span>
                  );
                })()}
              </div>
              {/* flexWrap + flex-basis makes the two panels sit side-by-side on wide
                  screens and stack on narrow (mobile) — no media query needed. */}
              <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                {([
                  { title: "Folders", feed: folderFeed },
                  { title: "Group assignments", feed: assignFeed },
                ] as const).map((panel) => (
                  <div key={panel.title} style={{ flex: "1 1 280px", minWidth: 0, border: "1px solid #e5e5e5", borderRadius: 4, overflow: "hidden" }}>
                    <div style={{ padding: "6px 10px", background: "#f7f7f7", fontSize: 12, fontWeight: 600, color: "#444", borderBottom: "1px solid #eee" }}>{panel.title}</div>
                    {/* overflowX:auto lets the client slide left/right to read full paths;
                        rows keep nowrap (no ellipsis clip) so the whole message is reachable. */}
                    <div style={{ maxHeight: 260, overflowY: "auto", overflowX: "auto", padding: "4px 0" }}>
                      {panel.feed.length === 0 ? (
                        <div style={{ padding: "6px 10px", fontSize: 12, color: "#aaa" }}>—</div>
                      ) : (
                        panel.feed.map((it, i) => (
                          <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "3px 10px", fontSize: 12, color: it.status === "admin" ? "#b45309" : "#333" }}>
                            {progIcon(it.status)}
                            <span style={{ whiteSpace: "nowrap" }}>{it.text}</span>
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

      {tab !== "Reconciliation" && tab !== "GroupMap" && (
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

      {log.length > 0 && (() => {
        // Split the log so the client sees clearly what succeeded vs what still needs
        // action. "Needs attention" = failures and folders created without a group
        // (the "no group-map groups (locked admin-only)" warnings + missing/failed grants).
        const isAttention = (e: LogEntry): boolean =>
          !e.ok || /no group-map groups|not found|FAILED|✗|⚠/.test(e.msg);
        const attention = log.filter(isAttention);
        const done = log.filter((e) => !isAttention(e));
        return (
          <div style={s.logBox}>
            <p style={{ ...s.logTitle, color: "#0f6c3f" }}>
              ✓ Completed — created &amp; assigned ({done.length})
            </p>
            {done.length === 0 ? (
              <div style={{ fontSize: 12, color: "#999", marginBottom: 6 }}>None yet.</div>
            ) : (
              done.map((entry, i) => (
                <div key={i} style={{ fontSize: 12, color: "#0f6c3f", marginBottom: 4, wordBreak: "break-word" }}>
                  ✓ {entry.msg}
                </div>
              ))
            )}
            <p style={{ ...s.logTitle, marginTop: 14, color: attention.length > 0 ? "#b45309" : "#0f6c3f" }}>
              ⚠ Needs attention — created without a group / errors ({attention.length})
            </p>
            {attention.length === 0 ? (
              <div style={{ fontSize: 12, color: "#0f6c3f" }}>None — every folder got a group. 🎉</div>
            ) : (
              attention.map((entry, i) => (
                <div key={i} style={{ fontSize: 12, color: entry.ok ? "#b45309" : "#d13438", marginBottom: 4, wordBreak: "break-word" }}>
                  {entry.ok ? "⚠" : "✗"} {entry.msg}
                </div>
              ))
            )}
          </div>
        );
      })()}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#d13438" : "#0f6c3f" }}>
          {toast.message}
          <button onClick={() => setToast(null)} style={s.toastClose}>✕</button>
        </div>
      )}
    </section>
  );
}
