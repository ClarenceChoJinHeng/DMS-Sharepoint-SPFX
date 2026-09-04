import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  roleFromGroupName,
  GroupMapRole,
  GroupMapDraft,
  GroupMapWriteRow,
  normalizeScope,
  PERSONAS,
  PERSONA_FAMILIES,
  personaByKey,
  personaTouchesStaging,
  roleLabel,
  Persona,
  normalizeRoleValue,
} from "../../../shared/groupMapModel";
import {
  searchSiteGroups,
  getGroupMembers,
  SpGroupMember,
} from "../../../shared/spGroups";
import {
  toCsv,
  downloadCsv,
  exportFileName,
  GroupExportRow,
  NO_MEMBERS,
} from "../../../shared/groupExportCsv";
import { parseLevels } from "../../../shared/formModel";
import { chainFor, tierChains, TermNode } from "../../../shared/termChains";
import { groupMappingsByGroup } from "../../../shared/groupMappings";
import GroupMembersEditor from "./GroupMembersEditor";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { RolesReference } from "../../../shared/rolesReference";
import { writeAudit } from "../../../shared/spAuditLog";

type Props = {
  context: WebPartContext;
  siteUrl: string;
  /**
   * Which half of this component to render.
   *
   * `members` (Folder Access) is the group list, its mappings and its people. `form` (Group
   * Management) is ONLY the map-a-group-by-hand card. `all` keeps both, for any caller that has not
   * chosen — nothing uses it today.
   *
   * TWO MOUNT POINTS OF ONE COMPONENT, never a copy. The form and the list share `existing`,
   * `postRow` and `isDuplicateRow`; splitting them into separate components would give this site two
   * definitions of what a mapping row is, and the one that drifted would be the rarely-used one.
   *
   * The form left Folder Access on the client's instruction (2026-08-18): mapping is not work done
   * there. It lands on Group Management because that is where BOTH cases needing it originate — a
   * group created with a free-typed name and no persona gets no rows at all, and a group needing a
   * second tier is created there too.
   */
  show?: "all" | "members" | "form";
};
type GroupPick = { id: string; displayName: string };
/** What a segment-scope row shows where a tier would be — it grants across the whole segment. */
const SEGMENT_LEVEL = "(segment level)";
/**
 * Shown when the term tree could not be walked. Deliberately not blank and deliberately not the
 * GUID: blank reads as "this row has no department", and a GUID reads as data corruption. Neither
 * is true — the row is fine and the lookup failed.
 */
const TIER_UNKNOWN = "Tier not known";

/** Case-insensitive GUID comparison, at module scope so the walk can use it before render. */
function guidEq(a: string, b: string): boolean {
  return (a ?? "").trim().toLowerCase() === (b ?? "").trim().toLowerCase();
}

type ModePick = {
  label: string;
  termSetGuid: string;
  /**
   * Permissioned tier names, shallowest first. Carried so the ancestry walk knows how deep to go —
   * a walk past the permissioned depth would descend into below-Unit tiers (SubUnit terms are
   * authored under each unit), which is hundreds of requests for terms no Group Map row can name.
   */
  levelNames: string[];
};
type TermLite = { id: string; label: string };
type ExistingRow = GroupMapWriteRow & { itemId: number };

// The role list that used to feed the picker is gone with the picker itself (2026-08-09).
// SELECTABLE_ROLES still exists in the model as the set of roles a Group Map row may legally
// carry — it just no longer drives anything on screen. ROLE_HINT below is still used: the
// persona checklist shows each role's plain-English meaning next to it.

// What each role actually gets you, in the admin's words rather than the
// permission level's. "APR" alone tells an administrator nothing about whether
// an approver can also upload, which is the entire distinction between the
// Head-of bundles.
// DEL and DELS are one letter apart and grant delete in DIFFERENT libraries, so
// the hints name the library rather than the action. "Delete documents" vs
// "Delete files" would be indistinguishable at a glance, and picking the wrong one
// grants delete over the wrong set of documents.
const ROLE_HINT: Record<string, string> = {
  MEMBER: "Read approved documents",
  // UPL and APR gained Documents READ on 2026-08-09, which is the whole point of the
  // one-group-per-person change — so the hints have to say it. "Upload to Staging" alone
  // now understates the role, and an admin reading it would still add the base group.
  UPL:    "Upload to the approval library, and read the unit's approved documents",
  APR:    "Approve pending items, and read the unit's approved documents",
  DEL:    "Delete approved documents — Documents library",
  DELS:   "Delete pending files — Staging library",
  // C-level. Reads wide, writes nothing, and Documents only — the hint says "Documents"
  // out loud because "view everything" would imply Staging too, and a C-level reading
  // other people's unapproved drafts is the one thing this role must not do.
  GLOBAL:  "C-level view — every segment, Documents only, read only",
  // Was MISSING until 2026-08-09, so the C-Level (segment) persona rendered a blank line
  // where its one capability should have been. A persona that describes itself as nothing
  // is worse than one with a terse label: the admin cannot tell whether it does nothing or
  // whether the page is broken.
  SEGVIEW: "C-level view — one business segment, Documents only, read only",
};
// Resolved, not hardcoded — the client renames both to CRS at import (confirmed on their site
// 2026-08-05). Read from the cache primed in the mount effect, which runs before any write.
const GROUP_MAP_LIST = (): string => cachedListTitle(LIST_SUFFIX.groupMap);
const CONFIG_LIST = (): string => cachedListTitle(LIST_SUFFIX.config);

const s: Record<string, React.CSSProperties> = {
  wrap:       { fontSize: 13, color: "#242424", lineHeight: 1.5 },
  intro:      { fontSize: 13, color: "#444", margin: "0 0 16px" },
  card:       { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  label:      { display: "block", fontWeight: 600, fontSize: 12, margin: "12px 0 4px" },
  input:      { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  select:     { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  ddwrap:     { position: "relative" },
  dd:         { position: "absolute", top: "100%", left: 0, right: 0, zIndex: 20, background: "#fff", border: "1px solid #c7c7c7", borderRadius: 4, maxHeight: 220, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,.12)" },
  ddItem:     { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f0f0f0" },
  roleRow:    { display: "flex", gap: 6, flexWrap: "wrap", marginTop: 4 },
  roleBtn:    { padding: "5px 12px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  roleActive: { background: "#0f6c3f", color: "#fff", borderColor: "#0f6c3f" },
  roleHint:   { fontSize: 12, color: "#605e5c", marginTop: 4 },
  personaBox: { border: "1px solid #e1e1e1", borderRadius: 4, padding: 10, marginTop: 6, background: "#fff", fontSize: 12 },
  personaRow: { display: "flex", gap: 8, alignItems: "center", marginBottom: 4, flexWrap: "wrap" },
  personaBlocked: { border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, padding: "6px 8px", marginBottom: 6, color: "#8a4b00" },
  pickedChip: { display: "inline-flex", alignItems: "center", gap: 8, padding: "5px 10px", background: "#eef6f0", border: "1px solid #b7dcc4", borderRadius: 4, fontSize: 12 },
  chipX:      { border: "none", background: "transparent", cursor: "pointer", color: "#0f6c3f", fontWeight: 700 },
  seglvl:     { marginTop: 6, fontSize: 12, color: "#0f6c3f", background: "transparent", border: "none", cursor: "pointer", textDecoration: "underline", padding: 0 },
  preview:    { marginTop: 14, padding: "10px 12px", background: "#fff", border: "1px dashed #b7dcc4", borderRadius: 4, fontSize: 12 },
  addBtn:     { marginTop: 12, padding: "7px 18px", fontSize: 13, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  addBtnOff:  { marginTop: 12, padding: "7px 18px", fontSize: 13, background: "#c7c7c7", color: "#fff", border: "none", borderRadius: 4, cursor: "not-allowed" },
  // Same footprint as addBtn (padding + fontSize) so the two buttons match; ghost colours.
  req:        { color: "#a4262c", marginLeft: 2 },
  missing:    { marginTop: 8, fontSize: 12, color: "#a4262c" },
  dangerBox:  { marginTop: 8, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  modalOverlay:{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 },
  modalBox:   { background: "#fff", borderRadius: 8, width: "100%", maxWidth: 480, maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 30px rgba(0,0,0,.25)" },
  modalHead:  { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", borderBottom: "1px solid #eee", fontWeight: 600, fontSize: 14 },
  modalBody:  { padding: "12px 16px", overflowY: "auto" },
  table:      { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  gList:      { border: "1px solid #ececec", borderRadius: 6, background: "#fff" },
  gRow:       { display: "flex", alignItems: "center", gap: 10, padding: "8px 10px", borderBottom: "1px solid #f0f0f0" },
  gName:      { fontWeight: 600, fontSize: 13, flex: 1, textAlign: "left", background: "none", border: "none", cursor: "pointer", padding: 0, color: "#1b1b1b", fontFamily: "inherit" },
  gPill:      { fontSize: 11, padding: "1px 7px", borderRadius: 10, background: "#eef4ff", border: "1px solid #cfe0ff", color: "#1b4b8a", whiteSpace: "nowrap" },
  gTier:      { fontSize: 11.5, color: "#605e5c", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 320 },
  gBody:      { padding: "4px 10px 10px 24px", background: "#fcfcfc" },
  // `display: block` is load-bearing: a <button> is inline by default, so the two disclosures and the
  // "Groups (308)" heading all rendered on ONE line, reading as a single run of words.
  disc:       { display: "block", background: "none", border: "none", padding: "0 0 10px", cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 13, fontWeight: 600, color: "#0f6c3f", textAlign: "left" },
  th:         { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:         { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "top" },
  delBtn:     { padding: "3px 10px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  ghost:      { padding: "3px 10px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  toast:      { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 6, color: "#fff", fontSize: 13, zIndex: 50, boxShadow: "0 4px 14px rgba(0,0,0,.18)" },
  mono:       { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  staleBadge: { display: "inline-block", marginLeft: 8, padding: "1px 7px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em", color: "#a4262c", background: "#fde7e9", border: "1px solid #f1b0b3", borderRadius: 10, verticalAlign: "middle" },
};

export default function GroupMapBuilder({
  context,
  siteUrl,
  show = "all",
}: Props): React.ReactElement {
  const [modes, setModes]       = useState<ModePick[]>([]);
  // True when DMS Config returned no usable `mode` rows at all — see the mount effect.
  const [modesUnreadable, setModesUnreadable] = useState(false);
  const [existing, setExisting] = useState<ExistingRow[]>([]);
  const [busy, setBusy]         = useState(false);
  // Persona dropdown open state + its container, for the click-away close. A ref on the
  // wrapper rather than a blur handler on the button: clicking an OPTION blurs the
  // button, so a blur-close would dismiss the list before the click registered and
  // nothing could ever be selected.
  const [personaOpen, setPersonaOpen] = useState(false);
  const personaRef = useRef<HTMLDivElement>(null);
  // Set when the Scope/Target columns are absent from the Group Map list. Without them every
  // row can only ever be a Folder row, so an admin who picks Site or Library writes a mapping
  // that reconciliation reads as a folder row with no term — and blames the term store for it.
  //
  // The detection has always run (loadExisting falls back to the six-field $select when the
  // eight-field one is rejected) but the result was never rendered, so the "surfaced rather
  // than silently degraded" promise in this comment was not kept. Rendered from 2026-08-09;
  // the unused-variable warning was the only sign that half the safeguard was missing.
  const [scopeColumnsMissing, setScopeColumnsMissing] = useState(false);

  // Close on a click outside the dropdown, or on Escape. Bound only while open, so the
  // page carries no listeners the rest of the time. `mousedown` rather than `click`:
  // a click that starts outside and ends inside would otherwise leave the list open.
  useEffect(() => {
    if (!personaOpen) return undefined;
    const onDocMouseDown = (e: MouseEvent): void => {
      if (!personaRef.current?.contains(e.target as Node)) setPersonaOpen(false);
    };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape") setPersonaOpen(false);
    };
    document.addEventListener("mousedown", onDocMouseDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDocMouseDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [personaOpen]);
  const [toast, setToast]       = useState<{ message: string; error: boolean } | undefined>(undefined);
  // NOTE: the Full Control check that used to live here went with group creation and member
  // editing (2026-08-14). What is left — writing and deleting Group Map rows — is governed by
  // permissions on that LIST, not by site ownership, so a "you need Full Control" banner here
  // would warn the wrong people and reassure the wrong people. The list write reports its own
  // failure. Group Management keeps the check, because creating a group genuinely needs it.
  // The three custom permission levels, named for THIS site. Hardcoding "DMS Upload" here while
  // reconciliation grants "CRS Upload" tells an admin to create a level nothing will ever use.
  /**
   * The group whose people and mappings are open. ONE at a time: the member editor issues a read
   * per group, and letting several stand open would fire 308 of them on a site this size.
   */
  const [openGroupId, setOpenGroupId] = useState<string | undefined>(undefined);
  /** Free-text filter over group names — 308 rows is not a list you scroll to find Tax. */
  const [groupFilter, setGroupFilter] = useState("");
  /**
   * The add-a-mapping-by-hand form, CLOSED by default (2026-08-18, client: Folder Access should not
   * let them create mappings at all).
   *
   * Kept rather than deleted, because two narrow cases have no other route: a group created with the
   * advanced free-text name and no persona gets no rows at all, and one group covering two tiers
   * cannot be expressed anywhere else. Mapping the first by any other means would require deleting
   * and re-creating the group, which loses its members. A disclosure answers the actual complaint —
   * the client never meets it — without stranding either case.
   */
  const [formOpen, setFormOpen] = useState(false);

  // Draft selections
  const [group, setGroup]         = useState<GroupPick | undefined>(undefined);
  const [role, setRole]           = useState<GroupMapRole | "">("");
  const [mode, setMode]           = useState<ModePick | undefined>(undefined);
  const [cascade, setCascade]     = useState<TermLite[][]>([]); // options per level
  const [chosen, setChosen]       = useState<TermLite[]>([]);   // picked term per level
  // The persona being provisioned. Purely a GUIDE: it never writes rows by itself.
  // A persona is 2–4 separate groups, and creating them in one click means a
  // failure halfway leaves a half-provisioned person whose access nobody can read
  // off the screen. The panel instead shows which rows exist and which are still
  // missing, and the admin adds them one at a time — slower, but the state is
  // always true.
  const [persona, setPersona]     = useState<string>("");
  const [tierGuid, setTierGuid]   = useState<string>("");

  /**
   * Which library's personas to offer (2026-08-09). Documents first: five of the six personas
   * live there, and it is the safer default to land on.
   *
   * The split is DERIVED — personaTouchesStaging() reads the persona's own roles — so the two
   * lists cannot disagree with what the personas actually grant. That matters more than it
   * looks. The client's first sketch of this toggle had the lists swapped, which would have
   * filed the C-Level personas under the approval library; GLOBAL/SEGVIEW there reads every
   * unapproved draft in a segment, so a UI that merely SUGGESTS it is a defect in itself.
   */
  const [libraryTab, setLibraryTab] = useState<"Documents" | "Staging">("Documents");

  /**
   * Is `recon_departmentFanOut` on? THREE states, and the third is the point.
   *
   * `undefined` = not read yet, or the read failed. Only an explicit `false` warns, so a
   * slow or broken config read stays silent rather than telling the admin their department
   * mapping is broken when it is not.
   */
  const [fanOutOn, setFanOutOn] = useState<boolean | undefined>(undefined);

  useEffect(() => {
    const read = async (): Promise<void> => {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG_LIST())}')/items`
          + `?$select=Title,SettingValue&$filter=Title eq 'recon_departmentFanOut'&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return;
      const rows = (await res.json()).value as Array<{ SettingValue?: string }>;
      // No row at all reads as ON, matching reconciliation's own default. A missing row is
      // the normal state on a site that never customised it, not an opt-out.
      if (rows.length === 0) { setFanOutOn(true); return; }
      const v = (rows[0].SettingValue ?? "").trim().toLowerCase();
      setFanOutOn(!(v === "off" || v === "false" || v === "0" || v === "no"));
    };
    read().catch(() => undefined);
  }, []);

  // Group search box
  const [query, setQuery]         = useState("");
  const [results, setResults]     = useState<GroupPick[]>([]);
  const [searching, setSearching] = useState(false);

  // Cache of tier term GUID -> its chain of labels, so the Existing mappings table shows
  // "Group Finance / Tax" instead of a raw GUID. Populated lazily as rows load.
  const [tierChainMap, setTierChains] = useState<Record<string, string[]>>({});
  /** Segments whose tree has been walked, so a second load does not re-walk them. */
  const walkedSegments = useRef<Record<string, true>>({});

  // NOTE (2026-08-14): creating a group, editing its members and deleting it all left this
  // screen for the Group Management web part — spec
  // 2026-08-14-group-management-separation-design.md. This page now does ONE thing: map an
  // EXISTING group to a segment, tier and role. See §5 of that spec for what was removed.

  const [confirmDel, setConfirmDel]   = useState<number | undefined>(undefined);
  const [selected, setSelected]       = useState<Set<number>>(new Set());
  const [confirmBulk, setConfirmBulk] = useState(false);

  // CSV export progress — "n/total" while member lists are being fetched.
  const [exporting, setExporting] = useState<{ done: number; total: number } | undefined>(undefined);

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 5000);
  };

  /* ── Data access ───────────────────────────────────────────────────────── */

  const searchGroups = async (q: string): Promise<GroupPick[]> => {
    const groups = await searchSiteGroups(context.spHttpClient, siteUrl, q);
    return groups.map((g) => ({ id: String(g.id), displayName: g.title }));
  };

  const loadModes = async (): Promise<ModePick[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG_LIST())}')/items?$select=ModeLabel,TermSetGuid,Levels&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ ModeLabel?: string; TermSetGuid?: string; Levels?: string }>)
      .filter((r) => (r.TermSetGuid ?? "").trim() && (r.Levels ?? "").trim())
      .map((r) => ({
        label: (r.ModeLabel ?? "").trim() || (r.TermSetGuid ?? "").trim(),
        termSetGuid: (r.TermSetGuid ?? "").trim(),
        // PERMISSIONED tiers only, and the filter is load-bearing: `parseLevels` keeps a
        // `"permissioned": false` entry, and walking to that depth would descend into the
        // below-Unit tiers (SubUnit terms are authored under each unit) — hundreds of requests
        // for terms no Group Map row can ever name.
        levelNames: parseLevels(r.Levels ?? "")
          .filter((l) => l.permissioned !== false)
          .map((l) => l.label),
      }));
  };

  const loadTerms = async (url: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      url, SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ id: string; labels: Array<{ name: string }> }>)
      .map((t) => ({ id: t.id, label: t.labels[0].name }));
  };

  const loadTops = (termSetGuid: string): Promise<TermLite[]> =>
    loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/children`);

  const loadChildren = (termSetGuid: string, parentId: string): Promise<TermLite[]> =>
    loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${termSetGuid}/terms/${parentId}/children`);

  const loadExisting = async (): Promise<ExistingRow[]> => {
    const base = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items`;
    const get = (select: string): Promise<SPHttpClientResponse> => context.spHttpClient.get(
      `${base}?$select=${select}&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    // Scope and Target are newer than this list. A $select naming a column that does not
    // exist fails the WHOLE request with HTTP 400 — not a null value, not a missing key —
    // so asking for them unconditionally would blank the mappings table on every site that
    // has not had the columns added, and the symptom reads as "my mappings disappeared"
    // rather than "a column is missing".
    //
    // So: ask for them, and fall back to the original field set when rejected. Same
    // reasoning as readAllowedFileTypesField (CLAUDE.md #11) — the only reliable signal
    // for "column absent" is the request failing, never the value that comes back.
    let res = await get("Id,GroupId,GroupName,Segment,UnitTermGuid,Role,Scope,Target");
    const hasScopeColumns = res.ok;
    if (!res.ok) res = await get("Id,GroupId,GroupName,Segment,UnitTermGuid,Role");
    setScopeColumnsMissing(!hasScopeColumns);
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ Id: number; GroupId?: string; GroupName?: string; Segment?: string; UnitTermGuid?: string; Role?: string; Scope?: string; Target?: string }>)
      .map((r) => ({
        itemId: r.Id,
        GroupId: r.GroupId ?? "",
        GroupName: r.GroupName ?? "",
        Segment: r.Segment ?? "",
        UnitTermGuid: r.UnitTermGuid ?? "",
        // Long-form values ("UPLOADER") normalise to the short code the tables key on, so a
        // hand-authored row shows as the role it means rather than as an unknown.
        Role: normalizeRoleValue(r.Role ?? "") as GroupMapRole,
        // Blank reads as Folder — every row written before the column existed is one.
        Scope: normalizeScope(r.Scope),
        Target: r.Target ?? "",
      }));
  };

  const postRow = async (row: GroupMapWriteRow): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata" },
        // Title is mandatory on a default SP list; set it so the create never 400s.
        body: JSON.stringify({ Title: row.GroupName || row.GroupId, ...row }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  const deleteRow = async (itemId: number): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "IF-MATCH": "*", "X-HTTP-Method": "DELETE" } },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  /**
   * Resolve tier terms for the Existing mappings table — the whole CHAIN, not just the leaf label.
   *
   * A row stores only the leaf GUID, so a unit row never said which department it was in and a
   * department row was indistinguishable from a unit one (client, 2026-08-18). `Tax`, `Legal` and
   * `PM` each exist under several departments on the client's tree, so the leaf label alone is
   * ambiguous about which folder a grant reaches — not merely terse.
   *
   * WALKS THE TREE ONCE PER SEGMENT rather than reading each term. That is what makes the ancestry
   * affordable at all, and it is also far cheaper than what it replaces: one read per distinct unit
   * term was ~130 requests for a provisioned segment, against ~8 for the walk (one per department,
   * plus the top level).
   *
   * Bounded by the segment's PERMISSIONED depth. Below-Unit tiers are authored under each unit, so
   * an unbounded walk would fetch every SubUnit on the site to answer a question about departments.
   */
  const loadTierChains = async (rows: ExistingRow[], modeList: ModePick[]): Promise<void> => {
    // Distinct segments that actually appear in a folder row, resolved up front so the async loop
    // never re-reads its own accumulator across an await (avoids a race-condition lint error).
    const wanted: string[] = [];
    for (const r of rows) {
      const g = (r.Segment ?? "").trim();
      if (!g || !r.UnitTermGuid) continue;
      if (wanted.indexOf(g) === -1 && !walkedSegments.current[g.toLowerCase()]) wanted.push(g);
    }
    if (wanted.length === 0) return;

    const nodes: TermNode[] = [];
    for (const setGuid of wanted) {
      const mode = modeList.filter((m) => guidEq(m.termSetGuid, setGuid))[0];
      // No mode row means no declared depth. Walking blind is the one thing not to do here — it is
      // unbounded — so the segment is left unresolved and its rows say so, which is honest: a
      // mapping under a segment with no mode row is already flagged `stale` beside it.
      const depth = mode ? mode.levelNames.length : 0;
      if (depth <= 0) continue;
      const walk = async (parentId: string, level: number): Promise<void> => {
        if (level > depth) return;
        const url = parentId
          ? `${siteUrl}/_api/v2.1/termStore/sets/${setGuid}/terms/${parentId}/children?$select=id,labels`
          : `${siteUrl}/_api/v2.1/termStore/sets/${setGuid}/children?$select=id,labels`;
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          url, SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } },
        );
        // THROWS rather than returning empty. A branch that could not be read must never be
        // reported as a term with no children — that presents a unit as though it were a
        // department, which is the exact wrong statement these columns exist to prevent. One bad
        // branch abandons the whole segment, whose rows then say "tier not known".
        if (!res.ok) throw new Error(`term store HTTP ${res.status}`);
        const kids = (
          ((await res.json()).value ?? []) as Array<{ id?: string; labels?: Array<{ name?: string }> }>
        );
        for (const k of kids) {
          if (!k.id) continue;
          nodes.push({ id: k.id, label: (k.labels ?? [])[0]?.name ?? "", parentId });
          await walk(k.id, level + 1);
        }
      };
      const ok = await walk("", 1).then(() => true).catch(() => false);
      // Marked walked only when the walk completed, so a transient failure is retried on the next
      // load rather than cached as "this segment has no terms".
      if (ok) walkedSegments.current[setGuid.toLowerCase()] = true;
    }

    const chains = tierChains(nodes);
    if (Object.keys(chains).length) setTierChains((prev) => ({ ...prev, ...chains }));
  };

  /* ── Mount ─────────────────────────────────────────────────────────────── */

  useEffect(() => {
    // Names FIRST. Every read below goes through cachedListTitle, and an unprimed cache resolves
    // to the legacy DMS titles — which 404 on a CRS-renamed site and would present as "the config
    // could not be read" rather than "the list is called something else".
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      /* The permission-level names are no longer read here — `RolesReference` resolves them itself,
         which is what lets it be mounted on two screens without either one having to fetch on its
         behalf. */
      .then(() => {
        // An empty mode list is not a neutral state: every Segment cell falls back to a
        // raw GUID and the segment picker offers nothing, which reads as "the data is
        // wrong" rather than "the config could not be read". Say which it is.
        loadModes()
          .then((m) => {
            setModes(m);
            if (m.length === 0) setModesUnreadable(true);
          })
          .catch(() => { setModes([]); setModesUnreadable(true); });
        loadExisting().then(setExisting).catch(() => setExisting([]));
      })
      .catch(() => undefined);
  }, []);

  // Whenever the mappings change, resolve any new tier chains. Keyed on `modes` too: the walk
  // needs a segment's declared depth, and the two loads race on mount — without it the first
  // render after `existing` lands would abandon every segment for want of a mode row.
  useEffect(() => {
    if (existing.length) loadTierChains(existing, modes).catch(() => undefined);
  }, [existing, modes]);

  /* ── Group search (debounced) ──────────────────────────────────────────── */

  useEffect(() => {
    if (group) return; // already picked
    const q = query.trim();
    if (q.length < 1) { setResults([]); return; }
    let cancelled = false;
    setSearching(true);
    const t = setTimeout(() => {
      searchGroups(q)
        .then((r) => { if (!cancelled) setResults(r); })
        .then(() => { if (!cancelled) setSearching(false); })
        .catch(() => { if (!cancelled) { setResults([]); setSearching(false); } });
    }, 350);
    return () => { cancelled = true; clearTimeout(t); };
  }, [query, group]);

  /* ── Selection handlers ────────────────────────────────────────────────── */

  const pickGroup = (g: GroupPick): void => {
    setGroup(g);
    setQuery(g.displayName);
    setResults([]);
    // Pre-select the role implied by the name suffix (_UPL/_APR); admin can override.
    // This is the one thing that still depends on the naming convention, which is why the
    // Group Management page keeps offering the name builder that produces it.
    setRole(roleFromGroupName(g.displayName));
  };

  const clearGroup = (): void => {
    setGroup(undefined);
    setQuery("");
    setResults([]);
  };

  /**
   * Every row the current draft should produce — one per role of the chosen persona.
   *
   * A persona used to be added a role at a time, on the reasoning that a half-failed
   * multi-write leaves someone half-provisioned. That was right while the admin could SEE
   * the roles: the checklist showed which were done. With the role chips gone (2026-08-09)
   * the roles are an implementation detail, and asking an admin to notice that Head of Unit
   * needs a second click is asking them to remember something the tool already knows —
   * the client's words: "the HoU actually will have delete by default".
   *
   * The half-provisioned risk is answered by REPORTING instead: rows go one at a time and
   * the caller names how many landed and which role failed, so the screen still tells the
   * truth after a partial failure.
   *
   */
  const draftRows = (d: GroupMapDraft): GroupMapWriteRow[] => {
    const p = personaByKey(persona);
    const roles: GroupMapRole[] =
      p && !p.unavailable && p.roles.length > 0 ? p.roles : ([d.role].filter(Boolean) as GroupMapRole[]);
    return roles.map((r) => buildGroupMapRow({ ...d, role: r }));
  };


  const pickRole = (r: GroupMapRole): void => {
    setRole(r);
    if (r === "GLOBAL") { setMode(undefined); setCascade([]); setChosen([]); setTierGuid(""); }
  };

  const pickMode = async (termSetGuid: string): Promise<void> => {
    const m = modes.find((x) => x.termSetGuid === termSetGuid);
    setMode(m);
    setChosen([]);
    setCascade([]);
    // A SEGMENT-scope persona (C-Level — one business segment) has no tier to pick: the
    // segment IS the tier. A segment-tier row is written with UnitTermGuid equal to the
    // TERM SET guid — see GroupMapWriteRow — so set it here rather than making the admin
    // choose a department they are not scoped to. Read via personaByKey, not the derived
    // `chosenPersona` further down, so this does not depend on declaration order.
    const p = personaByKey(persona);
    setTierGuid(m && p?.scope === "segment" ? m.termSetGuid : "");
    if (m) {
      const tops = await loadTops(m.termSetGuid).catch(() => [] as TermLite[]);
      setCascade([tops]);
    }
  };

  const pickTerm = async (levelIdx: number, termId: string): Promise<void> => {
    if (!mode) return;
    const term = (cascade[levelIdx] ?? []).find((t) => t.id === termId);
    if (!term) return;
    const newChosen = chosen.slice(0, levelIdx);
    newChosen[levelIdx] = term;
    setChosen(newChosen);
    setTierGuid(term.id); // assign at this tier by default
    const kids = await loadChildren(mode.termSetGuid, term.id).catch(() => [] as TermLite[]);
    const newCascade = cascade.slice(0, levelIdx + 1);
    if (kids.length) newCascade[levelIdx + 1] = kids;
    setCascade(newCascade);
  };

  /* ── Derived ───────────────────────────────────────────────────────────── */

  const draft: GroupMapDraft = {
    groupId: group?.id ?? "",
    groupName: group?.displayName ?? "",
    role: (role || undefined) as GroupMapRole,
    segmentGuid: mode?.termSetGuid,
    tierGuid,
  };
  const draftErrors = validateDraft(draft);
  const canAdd = draftErrors.length === 0 && !busy;

  // Term GUIDs are compared LOWERCASED, matching DMS Group Map, DMS Folder Map and
  // buildAbbrevIndex. SharePoint is not consistent about GUID casing between what a
  // text column stores and what the term store returns, and a case-sensitive compare
  // here degraded the Segment column to a raw GUID — and, worse, made isStaleRow flag
  // a perfectly healthy row as stale.
  const sameGuid = (a: string, b: string): boolean =>
    (a ?? "").trim().toLowerCase() === (b ?? "").trim().toLowerCase();

  const segmentLabelFor = (guid: string): string => {
    const m = modes.find((x) => sameGuid(x.termSetGuid, guid));
    return m ? m.label : guid;
  };

  // A mapping is stale when its Segment GUID no longer matches any current mode — e.g.
  // the term set was recreated (new GUID) or deleted. Such a row won't reconcile and
  // should be deleted + recreated. (Only meaningful once modes have loaded; GLOBAL rows
  // carry no segment and are never stale.)
  const isStaleRow = (r: ExistingRow): boolean =>
    modes.length > 0 && !!r.Segment && !modes.some((m) => sameGuid(m.termSetGuid, r.Segment));

  /**
   * The tier chain a row points at, or `undefined` when it was never resolved.
   *
   * `undefined` is NOT the same as a chain of one, and conflating them is the misreading these
   * columns exist to end: a chain of one means "this row is on a department", while unresolved
   * means "we do not know which tier this is". Presenting the second as the first would label a
   * unit row with a department heading — a wrong statement about who a grant reaches.
   */
  const tierChainFor = (r: ExistingRow): string[] | undefined => {
    if (!r.UnitTermGuid) return [];
    if (sameGuid(r.UnitTermGuid, r.Segment)) return [SEGMENT_LEVEL];
    return chainFor(tierChainMap, r.UnitTermGuid);
  };

  /**
   * The two tier cells, for the screen and for the CSV — one function, so they cannot disagree.
   *
   * A chain deeper than two is JOINED into the second cell rather than truncated: every family
   * today is exactly two permissioned tiers, but silently dropping a third would hide part of the
   * path a grant reaches, and a wrong-looking cell is recoverable where a missing one is not.
   */
  const tierCells = (r: ExistingRow): { tier1: string; tier2: string } => {
    const chain = tierChainFor(r);
    if (chain === undefined) return { tier1: TIER_UNKNOWN, tier2: "" };
    return { tier1: chain[0] ?? "", tier2: chain.slice(1).join(" › ") };
  };

  /**
   * Export every mapping with its group's members as a CSV the client can open in
   * Excel — one line per member, so a group with 4 people takes 4 lines and a group
   * with none still gets one line marked "(no members)".
   *
   * Members are fetched live rather than from any cache: the point of the export is
   * to be an accurate snapshot at the moment it is taken. Distinct group ids are
   * fetched once even when a group appears in several mappings.
   */
  const onExportCsv = async (): Promise<void> => {
    if (existing.length === 0) { showToast("There are no mappings to export.", true); return; }

    const groupIds: string[] = [];
    for (const r of existing) {
      if (r.GroupId && groupIds.indexOf(r.GroupId) === -1) groupIds.push(r.GroupId);
    }

    setExporting({ done: 0, total: groupIds.length });
    const byGroup: Record<string, SpGroupMember[]> = {};
    let failed = 0;
    for (let i = 0; i < groupIds.length; i++) {
      const id = groupIds[i];
      try {
        byGroup[id] = await getGroupMembers(context.spHttpClient, siteUrl, Number(id));
      } catch {
        // A deleted group (or one we can't read) must not abort the whole export —
        // it is flagged in its own row instead.
        byGroup[id] = [];
        failed++;
      }
      setExporting({ done: i + 1, total: groupIds.length });
    }

    const rows: GroupExportRow[] = [];
    for (const r of existing) {
      const cells = tierCells(r);
      const base = {
        group: r.GroupName || r.GroupId,
        segment: r.Segment ? segmentLabelFor(r.Segment) : "",
        tier1: cells.tier1,
        tier2: cells.tier2,
        role: r.Role,
      };
      const members = byGroup[r.GroupId] ?? [];
      if (members.length === 0) {
        rows.push({ ...base, memberName: NO_MEMBERS, memberEmail: "" });
      } else {
        for (const m of members) {
          rows.push({ ...base, memberName: m.title, memberEmail: m.email });
        }
      }
    }

    downloadCsv(toCsv(rows), exportFileName(new Date()));
    setExporting(undefined);
    showToast(
      failed === 0
        ? `Exported ${existing.length} mapping(s) across ${groupIds.length} group(s).`
        : `Exported, but ${failed} group(s) could not be read — their members are blank.`,
      failed > 0,
    );
  };

  const onAdd = async (): Promise<void> => {
    const rows = draftRows(draft);
    const fresh = rows.filter((r) => !isDuplicateRow(existing, r));
    if (fresh.length === 0) { showToast("This exact mapping already exists.", true); return; }
    setBusy(true);
    const added: string[] = [];
    let failedRole = "";
    let failedWhy = "";
    try {
      for (const r of fresh) {
        try {
          await postRow(r);
          added.push(r.Role);
        } catch (e) {
          failedRole = r.Role;
          failedWhy = (e as Error).message;
          break; // stop at the first failure; the toast names what did and did not land
        }
      }
      setExisting(await loadExisting());
      if (added.length > 0) {
        // reset the draft (keep the picked group for quick multi-role mapping)
        setRole(""); setMode(undefined); setCascade([]); setChosen([]); setTierGuid("");
      }
      if (failedRole) {
        showToast(
          added.length > 0
            ? `Added ${added.join(", ")}, but ${failedRole} failed: ${failedWhy}. Re-add from this page.`
            : `Add failed: ${failedWhy}`,
          true,
        );
      } else {
        showToast(
          `${added.length} row(s) added (${added.join(", ")}) — re-run Folder Reconciliation to apply permissions.`,
          false,
        );
      }
    } finally {
      setBusy(false);
    }
  };

  /**
   * Delete a mapping row. THE ROW ONLY.
   *
   * Until 2026-08-14 this also deleted the SharePoint group when it was that group's last
   * row, on the reasoning that nothing here revokes a role assignment, so a group with no
   * mappings keeps every grant it already holds. That reasoning was sound and the rule still
   * had to go: with group creation moved to its own page, "created but not yet assigned" is a
   * legitimate state, and the rule would silently destroy a group an admin made minutes
   * earlier. Groups are deleted deliberately, on the Group Management page, which states the
   * same caveat before it does so.
   *
   * The residual case is unchanged and still stated rather than hidden: deleting a row leaves
   * that grant on the folders until Folder Reconciliation runs. The toast says so rather than
   * implying the change has already taken effect.
   */
  const onDelete = async (itemId: number): Promise<void> => {
    const row = existing.find((r) => r.itemId === itemId);
    const gid = Number(row?.GroupId ?? NaN);
    const siblings = existing.filter((r) => Number(r.GroupId) === gid && r.itemId !== itemId);
    setBusy(true);
    try {
      await deleteRow(itemId);
      setExisting(await loadExisting());
      setSelected((prev) => { const n = new Set(prev); n.delete(itemId); return n; });
      setConfirmDel(undefined);
      showToast(
        `Row deleted. The group "${row?.GroupName}" is kept` +
        (siblings.length > 0 ? ` and still has ${siblings.length} other mapping(s).` : " with no mappings left.") +
        " Run Folder Reconciliation to take the folder permission away.",
        false,
      );
      // The reconciliation caveat is part of the RECORD, not just the toast: the grant outlives
      // the row, and an entry reading only "row deleted" would have someone believe otherwise.
      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.groupMapChanged,
        outcome: "Success",
        source: "FolderAccess",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        summary: `Group Map row deleted — ${row?.GroupName ?? `item ${itemId}`}${row?.Role ? ` (${row.Role})` : ""}`,
        details: [
          `Row: ${row?.GroupName ?? `item ${itemId}`}, role ${row?.Role ?? "(unknown)"}, scope ${row?.Scope ?? "(unknown)"}, target ${row?.Target ?? "(unknown)"}`,
          `The SharePoint group is KEPT. ${siblings.length} other mapping(s) remain.`,
          // Recorded accurately, because the row is the only record anyone will find later:
          // reconciliation does not revoke folder grants, so this deletion withdrew nothing.
          "The folder grant from this row REMAINS. Reconciliation does not remove folder grants — " +
            "reset the folder's permissions and reconcile, or delete the group, to withdraw it.",
        ],
      }).catch(() => undefined);
    } catch (e) {
      showToast(`Delete failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  const toggleSel = (itemId: number): void => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(itemId)) next.delete(itemId); else next.add(itemId);
      return next;
    });
  };

  // NOTE: `allSelected` is defined below, after existingForDisplay — the checkbox must reflect
  // the rows on screen, not the full set held for the delete paths.

  /**
   * The mappings table, ordered so a persona's rows sit together.
   *
   * A COPY — `existing` keeps its load order, which the selection and delete paths key on by
   * itemId; sorting in place would be fine today and a trap the first time something assumes
   * index order.
   *
   * Group first, then tier, then role. One person's mappings then read as a block instead of
   * being scattered by whatever order the list returned, which is how "Head of Unit needs two
   * rows" becomes visible rather than something you have to already know.
   */
  const existingForDisplay = existing
    // FOLDER rows only. This is the Folder Access page; a library, site or page mapping belongs
    // to Site Access / Approval Library Access / Page Access, and showing it here invited the
    // question "why is a library row in my folder list" — with no answer available on this page.
    //
    // DISPLAY only. `existing` still holds every row, and must: onDeleteGroup removes a group's
    // mappings before deleting the SharePoint group, and onDelete decides whether a row is the
    // group's LAST one. Filtering the source would leave library and page rows behind pointing
    // at a group that no longer exists, and would make a row look like the last when it is not.
    .filter((r) => normalizeScope(r.Scope) === "Folder")
    .sort((a, b) => {
      const g = (a.GroupName || a.GroupId).localeCompare(b.GroupName || b.GroupId);
      if (g !== 0) return g;
      const t = (a.UnitTermGuid ?? "").localeCompare(b.UnitTermGuid ?? "");
      if (t !== 0) return t;
      return (a.Role ?? "").localeCompare(b.Role ?? "");
    });

  /**
   * The visible rows collapsed to one entry per GROUP, then filtered by name.
   *
   * Filtered AFTER grouping, so a match keeps the whole group rather than a subset of its mappings:
   * a group showing 2 of its 13 rows would read as half-provisioned, which is the misreading the
   * mapping count exists to prevent.
   */
  const mappingGroups = groupMappingsByGroup(existingForDisplay).filter((g) => {
    const q = groupFilter.trim().toLowerCase();
    return q === "" || g.label.toLowerCase().indexOf(q) !== -1;
  });

  // Both keyed on what is VISIBLE. "Select all" reaching rows the admin cannot see, and then
  // deleting them, is the kind of surprise this page must not have.
  // NO select-all since the list collapsed to one row per group (2026-08-18). There is nothing left
  // for it to attach to, and a control that selected 790 rows behind 308 collapsed headers would put
  // "delete every mapping on the site" two clicks away. Per-row checkboxes inside an expanded group
  // still feed the "Delete selected" button.

  // Delete every checked row, then reload once. Reloads even on failure so the list
  // reflects any rows that were removed before the error.
  const onDeleteSelected = async (): Promise<void> => {
    setBusy(true);
    try {
      for (const id of Array.from(selected)) await deleteRow(id);
      setExisting(await loadExisting());
      setSelected(new Set());
      setConfirmBulk(false);
      showToast("Selected rows deleted — re-run Folder Reconciliation to apply the change.", false);
    } catch (e) {
      setExisting(await loadExisting());
      setSelected(new Set());
      setConfirmBulk(false);
      showToast(`Delete failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /* ── Render ────────────────────────────────────────────────────────────── */

  // Which roles the chosen tier already has a mapping for. Drives the persona
  // checklist. Keyed on the TIER, not the group: a persona is "this person, on
  // this unit", and it is satisfied by whichever groups carry those roles there.
  const rolesMappedAtTier: Set<string> = (() => {
    const out = new Set<string>();
    if (!tierGuid) return out;
    const want = tierGuid.trim().toLowerCase();
    existing.forEach((r) => {
      if ((r.UnitTermGuid ?? "").trim().toLowerCase() === want) out.add(normalizeRoleValue(r.Role ?? ""));
    });
    return out;
  })();

  /** Personas belonging to the selected library tab. Derived, never hand-listed — see libraryTab. */
  const personaInTab = (p: Persona): boolean =>
    personaTouchesStaging(p) === (libraryTab === "Staging");

  /**
   * Selecting a persona now SELECTS ITS ROLE too (2026-08-09). The role chips are gone, so
   * without this the admin would pick a persona and have no way to start a mapping.
   *
   * First role, not "first unmapped": tierGuid is usually still empty at this point, so
   * rolesMappedAtTier has nothing to say yet. The checklist below shows what is already done
   * and lets them switch to the second role — Head of Unit is the only persona with one.
   */
  const choosePersona = (p: Persona): void => {
    setPersona(p.key);
    setPersonaOpen(false);
    if (!p.unavailable && p.roles.length > 0) pickRole(p.roles[0]);
    // A department mapping is INERT unless fan-out is on, and the failure is silent: the
    // department folder gets the grant, no unit does. Raised only when it is actually off —
    // as standing panel text it warned about a setting that is on, every single time.
    // `false` specifically, not falsy: undefined means the config read failed, and guessing
    // "off" there would cry wolf on every site where the read is merely slow.
    if (p.scope === "department" && fanOutOn === false) {
      showToast(
        "recon_departmentFanOut is OFF in Config, so this mapping will stop at the department folder and reach no units. Switch it on, then run Folder Reconciliation.",
        true,
      );
    }
  };

  const chosenPersona = personaByKey(persona);

  // Is the picked tier a leaf (a unit)? pickTerm only appends the next cascade
  // level when the chosen term HAS children, so an equal length means the deepest
  // pick has none. Derived rather than counting mode.levels, because a term store
  // can be shallower than its configured Levels chain and the folder tree follows
  // the terms, not the config.
  const tierIsLeaf = chosen.length > 0 && cascade.length === chosen.length;

  /**
   * How many tier dropdowns the chosen persona is allowed to see (2026-08-09).
   *
   * Previously every level the term store could load was rendered, so a Head of Department
   * was walked all the way down to a unit and then TOLD OFF for picking one. The scope is
   * known before the admin touches anything, so the deeper levels should not exist rather
   * than be a mistake waiting to be made.
   *
   *   segment    → 0. The segment IS the tier; the Segment dropdown already chose it.
   *   department → 1. Top-level terms of a segment's set ARE its departments.
   *   unit       → all of them.
   *
   * Without a persona (no longer reachable from this page, but the state is still
   * expressible) the full chain shows, which is the old behaviour.
   */
  const maxTierLevels =
    !chosenPersona ? cascade.length
    : chosenPersona.scope === "segment" ? 0
    : chosenPersona.scope === "department" ? 1
    : cascade.length;
  const visibleCascade = cascade.slice(0, maxTierLevels);

  // Only meaningful once a tier is picked; until then neither warning applies. A
  // segment-scope persona can no longer mis-pick — it has no tier dropdown at all — so it
  // is exempt rather than being given a third message it can never act on.
  const scopeMismatch: string | undefined =
    !chosenPersona || !tierGuid || chosenPersona.scope === "segment" ? undefined
    : chosenPersona.scope === "department" && tierIsLeaf
      ? `${chosenPersona.family} covers every unit under a department, but you have picked a unit. Pick the department instead, or use the matching Head of Unit group.`
    : chosenPersona.scope === "unit" && !tierIsLeaf
      ? `${chosenPersona.family} covers one unit, but you have picked a tier that has units beneath it. This would reach all of them.`
    : undefined;

  const personaPanel = (
    <>
      <label style={s.label}>Persona (optional)</label>
      {/* A custom dropdown, not a native <select>.
          The native one could not be made to behave: which direction its popup opens,
          and when it dismisses, are decided by the browser and OS — so a twelve-item
          list on a scrolling page opened UPWARD and closed on the smallest stray
          movement. No CSS or attribute changes either behaviour. This opens downward,
          stays open until a choice is made or the user clicks away, and scrolls
          internally rather than growing past the page. */}
      <div ref={personaRef} style={{ position: "relative" }}>
        <button
          type="button"
          disabled={busy}
          aria-haspopup="listbox"
          aria-expanded={personaOpen}
          style={{ ...s.select, textAlign: "left", cursor: busy ? "default" : "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}
          onClick={() => setPersonaOpen(!personaOpen)}
        >
          {/* "pick a role directly" described the role chips, which are gone. The persona is
              the only way in now, so the placeholder has to say that. */}
          <span>{chosenPersona ? `${chosenPersona.family} — ${chosenPersona.label}` : "— select a persona —"}</span>
          <span style={{ marginLeft: 8, color: "#605e5c" }}>{personaOpen ? "▲" : "▼"}</span>
        </button>
        {personaOpen && (
          <div
            role="listbox"
            style={{
              position: "absolute", top: "100%", left: 0, right: 0, zIndex: 30,
              marginTop: 2, maxHeight: 320, overflowY: "auto",
              background: "#fff", border: "1px solid #c8c6c4", borderRadius: 4,
              boxShadow: "0 4px 12px rgba(0,0,0,0.18)",
            }}
          >
            <div
              role="option"
              aria-selected={!persona}
              style={{ padding: "7px 10px", cursor: "pointer", color: "#605e5c" }}
              onClick={() => { setPersona(""); setRole(""); setPersonaOpen(false); }}
            >
              — none —
            </div>
            {/* Grouped exactly as the client's own document lists them, so an admin
                reading from that document finds the same groups here — then filtered to
                the selected library, so the two tabs never offer the same persona twice. */}
            {PERSONA_FAMILIES.map((fam) => (
              <div key={fam}>
                {PERSONAS.filter(personaInTab).some((p) => p.family === fam) && (
                  <div style={{ padding: "6px 10px", background: "#f3f2f1", fontWeight: 600, color: "#323130", position: "sticky", top: 0 }}>
                    {fam}
                  </div>
                )}
                {PERSONAS.filter((p) => p.family === fam && personaInTab(p)).map((p) => {
                  const selected = p.key === persona;
                  return (
                    <div
                      key={p.key}
                      role="option"
                      aria-selected={selected}
                      title={p.summary}
                      style={{
                        padding: "7px 10px 7px 20px", cursor: "pointer",
                        background: selected ? "#deecf9" : undefined,
                        color: p.unavailable ? "#a19f9d" : "#242424",
                      }}
                      onClick={() => choosePersona(p)}
                    >
                      {p.label}{p.unavailable ? " (not available)" : ""}
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        )}
      </div>
      {chosenPersona && (
        <div style={s.personaBox}>
          <div style={{ marginBottom: 6, color: "#444" }}>
            <strong>{chosenPersona.family}</strong> — {chosenPersona.summary}
          </div>

          {chosenPersona.unavailable ? (
            // Shown rather than hidden: this is one of the client's numbered
            // groups, and an admin who cannot find it will assume it was missed.
            <div style={s.personaBlocked}>{chosenPersona.unavailable}</div>
          ) : (
            <>
              <div style={{ marginBottom: 6, color: "#444" }}>
                Needs <strong>{chosenPersona.roles.length}</strong> mapping
                {chosenPersona.roles.length === 1 ? "" : "s"} at the{" "}
                <strong>{chosenPersona.scope}</strong> tier
                {chosenPersona.roles.length === 1
                  ? ", written when you add it."
                  : `, both written together when you add it.`}
              </div>
              {!tierGuid && <div style={{ color: "#8a6d00", marginBottom: 8 }}>Choose where this group applies to see what is already in place.</div>}
              {scopeMismatch && <div style={s.personaBlocked}>{scopeMismatch}</div>}
              {/* A READ-ONLY list since 2026-08-09, not a set of buttons. Adding now writes
                  every row the persona needs in one action, so a clickable role would be
                  offering a choice the tool has already made — and the client read that
                  choice as "the Head of Unit might not get delete", which is exactly what
                  it must not imply. Still shows ✓ per role so a partial add is visible. */}
              {chosenPersona.roles.map((r) => {
                const done = rolesMappedAtTier.has(r);
                return (
                  <div key={r} style={s.personaRow}>
                    <span style={{ width: 16, color: done ? "#107c10" : "#a19f9d" }}>{done ? "✓" : "○"}</span>
                    <span style={{ color: "#605e5c" }}>{ROLE_HINT[r]}</span>
                    {done && <span style={{ color: "#107c10" }}>already in place</span>}
                  </div>
                );
              })}
              {/* The recon_departmentFanOut caveat used to live here as permanent text.
                  Removed 2026-08-09: the client will not touch that config, so a standing
                  warning about a setting that is ON described a problem nobody had and
                  trained admins to read past the panel. It is now a TOAST, raised only
                  when the setting is actually off — see choosePersona. */}
            </>
          )}
        </div>
      )}
    </>
  );

  const selectionFields = (
    <>
      {personaPanel}
      {/* The free-choice role grid was removed 2026-08-09. Picking a persona already sets the
          roles, so seven chips underneath were either redundant or an invitation to build a
          combination the model never intended — the client's words were that the role naming
          "is confusing for users". The persona checklist above is now the only way to start a
          mapping, and DELS — previously the one role no persona carried, and so the reason to
          keep a manual path — now belongs to Head of Unit. */}
      {role && <div style={s.roleHint}>{ROLE_HINT[role]}</div>}

      {role && role !== "GLOBAL" && (
        <>
          <label style={s.label}>Segment <span style={s.req}>*</span></label>
          <select
            style={s.select}
            disabled={busy}
            value={mode?.termSetGuid ?? ""}
            onChange={(e) => { pickMode(e.target.value).catch(() => undefined); }}
          >
            <option value="">— select a segment —</option>
            {modes.map((m) => <option key={m.termSetGuid} value={m.termSetGuid}>{m.label}</option>)}
          </select>

          {mode && visibleCascade.length > 0 && (
            <>
              <label style={s.label}>Tier (where this group applies) <span style={s.req}>*</span></label>
              {visibleCascade.map((opts, i) => (
                <select
                  key={i}
                  style={{ ...s.select, marginBottom: 6 }}
                  disabled={busy}
                  value={chosen[i]?.id ?? ""}
                  onChange={(e) => { pickTerm(i, e.target.value).catch(() => undefined); }}
                >
                  <option value="">— select —</option>
                  {opts.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
                </select>
              ))}
            </>
          )}
        </>
      )}
    </>
  );


  return (
    <div style={s.wrap}>
      {/* ONE paragraph. Mounted inside the guided flow this sits directly under the step's own hint,
          which says the same thing — three stacked paragraphs of near-identical text is what the
          client saw, and it reads as a page repeating itself. */}
      {show !== "form" && (
        <p style={s.intro}>
          Expand a group to add or remove people — access applies immediately, and nothing here needs
          Folder Reconciliation. Adding someone is a <strong>group</strong> change: they get that
          group&apos;s folders in every library it is mapped to. Groups and their mappings are created
          on the <strong>Group Management</strong> page.
        </p>
      )}

      {/* ⚠ THE ROLES REFERENCE MOVED TO `shared/rolesReference.tsx` (2026-08-30) and is RENDERED
          here rather than duplicated. It had been unreachable since this component's `show="form"`
          mount was retired, and it is now also mounted on the guided flows' Group Management step —
          two copies of a reference nobody reads often are exactly the pair that drifts, and the
          drifting one is always the rarely-seen one.

          Still collapsed here: on this screen it is reference an admin needs ONCE, when the
          permission levels are first created or when choosing between DEL and DELS. The flow step
          opens it, because there it is the point of the step. */}
      {show !== "members" && (
        <RolesReference
          spHttpClient={context.spHttpClient}
          siteUrl={siteUrl}
          defaultOpen={false}
        />
      )}

      {/* The half of the safeguard that was missing until 2026-08-09 — the detection ran, the
          result went nowhere. Named as a PROVISIONING problem, not a page problem: everything
          here still works, and the damage is done later and elsewhere, by reconciliation. */}
      {scopeColumnsMissing && (
        <div style={{ ...s.card, borderColor: "#d13438", background: "#fdf3f4" }}>
          <strong>This site&rsquo;s Group Map list is missing the Scope and Target columns.</strong>{" "}
          Every mapping is therefore read as a <em>folder</em> mapping. Site, library and page
          access cannot be granted, and a row written for one of those is treated as a folder row
          with no term — which surfaces later as a term-store complaint rather than a missing
          column. Add both columns to the list, then re-run Folder Reconciliation.
        </div>
      )}

      {/* ⚠ NOT MOUNTED ANYWHERE AS OF 2026-08-23 (client: "I think you can remove map a group by
          hand"). It left Folder Access on 2026-08-18 and Group Management today, so nothing passes
          `show="form"` and this whole branch is unreachable.

          KEPT rather than deleted, because it is still the only route for the two cases bulk
          provisioning cannot express: a group created with the advanced free-text name and no
          persona gets NO Group Map rows at all (`rowsForNewGroup` returns []), and one group covering
          two tiers cannot be stated anywhere else. Both are now fixed only by deleting and
          re-creating the group. Re-mount by passing show="form" if that bites. */}
      {show !== "members" && (
      <button style={s.disc} onClick={() => setFormOpen((v) => !v)}>
        {formOpen ? "▾" : "▸"} Map a group by hand — rarely needed
      </button>
      )}

      {formOpen && show !== "members" && (
      <div style={s.card}>
        <div style={{ fontSize: 11.5, color: "#8a4b00", background: "#fff8f0", border: "1px solid #f2c9a0", borderRadius: 4, padding: "8px 10px", marginBottom: 12, lineHeight: 1.5 }}>
          You should not normally need this. Creating a group on <strong>Group Management</strong>
          {" "}writes its mappings, and <strong>Create all groups for a segment</strong> writes every
          mapping a segment needs. Use this only for a group that follows no naming convention, or to
          restore a row that was deleted.
        </div>
        {/* Library toggle — decides which personas are on offer. Above the group field
            because it frames everything below it: an admin picks the library they are
            granting in, THEN who they are granting to. */}
        <label style={s.label}>Library</label>
        <div style={s.roleRow}>
          {([
            { key: "Documents" as const, label: "Documents", hint: "The approved archive — C-Level, Head of Department, SDG Employee" },
            { key: "Staging" as const, label: "Approval Document", hint: "Upload and approval — Head of Unit, PIC (these also read their unit in Documents)" },
          ]).map((t) => (
            <button
              key={t.key}
              disabled={busy}
              title={t.hint}
              style={{ ...s.roleBtn, ...(libraryTab === t.key ? s.roleActive : {}), minWidth: 150 }}
              onClick={() => {
                if (libraryTab === t.key) return;
                setLibraryTab(t.key);
                // The selected persona belongs to the tab being left, so it would sit there
                // describing a library the admin is no longer looking at. Clearing the role
                // with it matters more: a stale role is what actually gets WRITTEN.
                setPersona("");
                setRole("");
              }}
            >
              {t.label}
            </button>
          ))}
        </div>
        <div style={s.roleHint}>
          {libraryTab === "Documents"
            ? "Who may read (and delete) approved documents."
            : "Who may upload and approve. Head of Unit and PIC also read their own unit in Documents — one group covers both."}
        </div>

        {/* Group */}
        <label style={s.label}>Group</label>
        {group ? (
          <span style={s.pickedChip}>
            {group.displayName}
            <button style={s.chipX} disabled={busy} title="Change" onClick={clearGroup}>✕</button>
          </span>
        ) : (
          /* Search ONLY. Creating a group moved to the Group Management page on 2026-08-14 —
             spec §4.4. The dropdown's old "Create a new group" row is replaced by a SIGNPOST:
             with five separate access pages, nothing else on screen would tell an admin that
             groups are made somewhere else first. A dead end became a direction. */
          <div style={s.ddwrap}>
            <input
              style={s.input}
              placeholder="Type to search this site's groups…"
              value={query}
              disabled={busy}
              onChange={(e) => setQuery(e.target.value)}
            />
            {(searching || results.length > 0 || query.trim()) && (
              <div style={s.dd}>
                {searching && <div style={s.ddItem}>Searching…</div>}
                {!searching && results.map((g) => (
                  <div key={g.id} style={s.ddItem} onClick={() => pickGroup(g)}>{g.displayName}</div>
                ))}
                {!searching && results.length === 0 && query.trim() && (
                  <div style={{ ...s.ddItem, color: "#666" }}>
                    No group by that name. Groups are created on the{" "}
                    <strong>Group Management</strong> page.
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* Pick a group, pick a persona, pick where. That is the whole page now. */}
        {selectionFields}

        {group && role && draftErrors.length > 0 && (
          <div style={s.missing}>Before adding: {draftErrors.join(" ")}</div>
        )}

        {/* Scope mismatch BLOCKS, it does not merely warn. A Head of Department mapped at
            the unit tier is the worst kind of wrong: the row is valid, reconciliation
            grants it happily, and the head silently ends up seeing one unit instead of the
            whole department. Nothing downstream can detect that — the row looks exactly
            like a legitimate unit mapping — so this is the only place it can be caught. */}
        <button
          style={canAdd && !scopeMismatch ? s.addBtn : s.addBtnOff}
          disabled={!canAdd || !!scopeMismatch}
          onClick={() => { onAdd().catch(() => undefined); }}
        >
          Add mapping
        </button>
      </div>
      )}

      {modesUnreadable && (
        <div style={{ fontSize: 12, color: "#b45309", background: "#fff8e1", border: "1px solid #f0c000", borderRadius: 4, padding: "8px 12px", marginBottom: 12 }}>
          Could not read any <strong>mode</strong> rows from <strong>{CONFIG_LIST()}</strong>, so the
          Segment column below shows raw term-set GUIDs and the segment picker is empty. The
          mappings themselves are fine — this is a config read, not your data.
        </div>
      )}

      {show !== "form" && (
      <>
      {/* Counts GROUPS now, with the mapping total beside it — the list is one row per group, and a
          count you cannot reconcile with the rows in front of you is worse than none. Both are of
          what is SHOWN: library, site and page rows live on their own pages. */}
      {/* Wrapped in a block, so the inline-block underline still hugs its own text instead of the
          heading joining the disclosure buttons above it on one line. */}
      <div>
        <div style={{ fontSize: 13, fontWeight: 600, color: "#0f6c3f", margin: "0 0 12px", paddingBottom: 8, borderBottom: "2px solid #0f6c3f", display: "inline-block" }}>
          Groups ({mappingGroups.length}) · {existingForDisplay.length} mapping
          {existingForDisplay.length === 1 ? "" : "s"}
        </div>
      </div>

      {/* Existing rows */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, margin: "0 0 6px" }}>
        <button
          style={s.ghost}
          disabled={busy || exporting !== undefined || existing.length === 0}
          title="Download every mapping and its group members as a CSV (opens in Excel)"
          onClick={() => { onExportCsv().catch(() => { setExporting(undefined); showToast("Export failed.", true); }); }}
        >
          {exporting ? `Exporting… ${exporting.done}/${exporting.total}` : "Export to Excel (CSV)"}
        </button>
        {selected.size > 0 && (
          confirmBulk ? (
            <span style={{ display: "inline-flex", gap: 6, alignItems: "center" }}>
              <span style={{ fontSize: 12, color: "#a4262c" }}>Delete {selected.size} selected row(s)?</span>
              <button style={s.delBtn} disabled={busy} onClick={() => { onDeleteSelected().catch(() => undefined); }}>Yes, delete</button>
              <button style={s.ghost} disabled={busy} onClick={() => setConfirmBulk(false)}>Cancel</button>
            </span>
          ) : (
            <button style={s.delBtn} disabled={busy} onClick={() => setConfirmBulk(true)}>
              Delete selected ({selected.size})
            </button>
          )
        )}
      </div>
      {/* Filter first: 308 groups is not a list you scroll to find Tax. Matches the group NAME,
          which is what an admin knows — the folder codes are in it (GHO_GF_TAX_UPLOADER). */}
      <input
        style={{ ...s.input, marginBottom: 8 }}
        value={groupFilter}
        placeholder="Filter by group name, e.g. TAX or _APPROVER&hellip;"
        onChange={(e) => setGroupFilter(e.target.value)}
      />

      {mappingGroups.length === 0 && (
        <div style={{ ...s.td, borderBottom: "none" }}>
          {existingForDisplay.length === 0 ? (
            <>
              No folder mappings yet.
              {existing.length > 0 && (
                // Says where the rows went. Without this, an admin who can see mappings exist
                // elsewhere reads an empty list as the page being broken.
                <span style={{ color: "#605e5c" }}>
                  {" "}This site has {existing.length} mapping(s) for libraries, the site or pages
                  — those are managed on Site Access, Approval Library Access and Page Access.
                </span>
              )}
            </>
          ) : (
            <span style={{ color: "#605e5c" }}>No group matches that filter.</span>
          )}
        </div>
      )}

      {/* ONE ROW PER GROUP, not per mapping (2026-08-18). Membership is a property of the GROUP, and
          a unit's approver group carries six mapping rows — so a per-mapping member editor would
          have shown the same people thirteen times per unit, with thirteen Add boxes doing one
          thing. 790 rows became 308.

          Scroll capped only while nothing is expanded: the people picker inside an open group is
          absolutely positioned, and a scroll container clips it for any group near the bottom. */}
      <div
        style={{
          ...s.gList,
          ...(openGroupId === undefined ? { maxHeight: "60vh", overflowY: "auto" as const } : {}),
        }}
      >
        {mappingGroups.map((mg) => {
          const isOpen = openGroupId === mg.groupId && mg.groupId !== "";
          // The tiers this group reaches, deduped. Normally one; more than one means somebody mapped
          // the same group twice deliberately, and the collapsed row should say so, not hide it.
          const tiers: string[] = [];
          for (const r of mg.rows) {
            const c = tierCells(r);
            const text = c.tier2 ? c.tier1 + " › " + c.tier2 : c.tier1;
            if (text && tiers.indexOf(text) === -1) tiers.push(text);
          }
          const numericId = Number(mg.groupId);
          return (
            <div key={mg.groupId || "blank-" + mg.rows[0].itemId}>
              <div style={s.gRow}>
                <button
                  type="button"
                  style={s.gName}
                  onClick={() => setOpenGroupId(isOpen ? undefined : mg.groupId)}
                >
                  {isOpen ? "▾" : "▸"} {mg.label}
                </button>
                {mg.rows.some((r) => isStaleRow(r)) && (
                  <span style={s.staleBadge} title="A mapping here points at a term set or term that no longer exists (likely recreated with a new GUID). Delete it, map it again, then re-run Folder Reconciliation.">stale</span>
                )}
                <span style={s.gTier} title={tiers.join(" · ")}>{tiers.join(" · ")}</span>
                {/* The count is what tells a fully provisioned unit group (13) from a _HOD (1) at a
                    glance — the same distinction the Tier columns were added to make. */}
                <span style={s.gPill}>
                  {mg.rows.length} mapping{mg.rows.length === 1 ? "" : "s"}
                </span>
              </div>

              {isOpen && (
                <div style={s.gBody}>
                  <table style={s.table}>
                    <thead>
                      <tr>
                        <th style={{ ...s.th, width: 28 }} />
                        <th style={s.th}>Segment</th>
                        <th style={s.th}>Tier 1</th>
                        <th style={s.th}>Tier 2</th>
                        <th style={s.th}>Role</th>
                        <th style={s.th} />
                      </tr>
                    </thead>
                    <tbody>
                      {mg.rows.map((r) => (
                        <tr key={r.itemId}>
                          <td style={s.td}>
                            <input type="checkbox" checked={selected.has(r.itemId)} disabled={busy} onChange={() => toggleSel(r.itemId)} />
                          </td>
                          {/* A blank Segment/Tier is never a MISSING value, so it must not read like
                              one. Two things arrive here termless and they mean opposite extremes:
                                GLOBAL      — no term because it reaches EVERY segment. The widest
                                              grant here; "not assigned" would read as broken.
                                non-Folder  — Site/Library/Page rows have no folder to carry a
                                              term. Shown as the target instead. */}
                          <td style={s.td}>
                            {r.Segment
                              ? segmentLabelFor(r.Segment)
                              : r.Role === "GLOBAL"
                                ? <span title="A C-Level global row carries no term: it reaches every segment.">All segments</span>
                                : <span style={{ color: "#605e5c" }}>Not applicable</span>}
                          </td>
                          {(() => {
                            const cells = tierCells(r);
                            const known = cells.tier1 !== TIER_UNKNOWN;
                            const first = r.UnitTermGuid
                              ? (known
                                  ? cells.tier1
                                  : <span
                                      style={{ color: "#8a4b00" }}
                                      title="The segment's term tree could not be read, so this row's tier is not known. The mapping itself is unaffected."
                                    >{TIER_UNKNOWN}</span>)
                              : r.Role === "GLOBAL"
                                ? <span title="Reaches every folder in every segment, at Read, in Documents only.">Every folder</span>
                                : r.Target
                                  ? <span style={s.mono} title="This row grants entry to a library or page, not a folder.">{r.Target}</span>
                                  : <span style={{ color: "#605e5c" }}>Not applicable</span>;
                            return (
                              <>
                                <td style={s.td}>{first}</td>
                                {/* A department-scope row legitimately has no Tier 2, and that
                                    emptiness is the fact identifying it — a dash, not a blank. */}
                                <td style={s.td}>
                                  {cells.tier2 || <span style={{ color: "#605e5c" }}>&mdash;</span>}
                                </td>
                              </>
                            );
                          })()}
                          {/* The label, not the code. "DELS" told an administrator nothing, and being
                              one letter from "DEL" — a DIFFERENT role, in the other library —
                              made it worse than uninformative. The code is still what is stored. */}
                          <td style={s.td}>{roleLabel(r.Role)}</td>
                          <td style={s.td}>
                            <button style={s.delBtn} disabled={busy} onClick={() => setConfirmDel(r.itemId)}>Delete</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  {/* WHO IS IN IT — the job this page exists for since 2026-08-18. Mounted only for
                      a real group id: a row left behind by a deleted group has nothing to read
                      members from, and an empty editor there would read as "nobody is in it". */}
                  {mg.groupId && !isNaN(numericId) ? (
                    <GroupMembersEditor
                      context={context}
                      siteUrl={siteUrl}
                      group={{ id: numericId, title: mg.groupName || mg.label }}
                      showToast={showToast}
                    />
                  ) : (
                    <p style={{ fontSize: 11.5, color: "#a4262c", margin: "8px 0 0 24px" }}>
                      This mapping carries no usable group id, so its members cannot be read. The
                      group was probably deleted — delete the row above and map it again.
                    </p>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
      </>
      )}

      {/* Delete-mapping confirmation — a modal, not an inline Yes/Cancel in the table cell.
          Between 2026-08-04 and 2026-08-14 removing a group's LAST mapping also deleted the
          SharePoint group, so this modal had a second, far more alarming form. That rule is
          GONE (spec §2 D4): "created but not yet assigned" is now a legitimate state, and the
          old behaviour would have destroyed a group an admin made minutes earlier. Deleting a
          row deletes a row. Groups are deleted on the Group Management page. */}
      {confirmDel !== undefined && (() => {
        const row = existing.find((r) => r.itemId === confirmDel);
        const gid = Number(row?.GroupId ?? NaN);
        const others = existing.filter((r) => Number(r.GroupId) === gid && r.itemId !== confirmDel);
        return (
          <div style={s.modalOverlay} onClick={() => setConfirmDel(undefined)}>
            <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
              <div style={s.modalHead}>
                <span>Delete mapping?</span>
                <button style={s.chipX} onClick={() => setConfirmDel(undefined)} title="Close">✕</button>
              </div>
              <div style={s.modalBody}>
                <p style={{ margin: "0 0 8px" }}>
                  Remove the <strong>{roleLabel(row?.Role as GroupMapRole)}</strong> mapping for{" "}
                  <strong>{row?.GroupName}</strong>?
                </p>
                <p style={{ margin: "0 0 8px" }}>
                  The group itself is <strong>kept</strong>
                  {others.length > 0
                    ? ` — it still has ${others.length} other mapping${others.length === 1 ? "" : "s"}.`
                    : ", with no mappings left. Delete it on the Group Management page if it is no longer needed."}
                </p>
                {/* ⚠ THIS SAID RECONCILIATION WOULD REMOVE THE GRANT. IT DOES NOT (corrected
                    2026-08-21, while an admin was following it to take CRS Delete off a PIC group).
                    `groupsToRemove` asserts a full ACL at PAGE scope ONLY — at folder scope a
                    library root also carries SharePoint's automatic Limited Access entries for every
                    principal granted below it, so asserting there would strip every group's access
                    on a run that reported success.

                    So deleting a row removes the row and NOTHING else, for ever. The old wording
                    would have led someone to delete a mapping, run reconciliation, and believe an
                    access had been withdrawn while the group still held it — with the list no longer
                    showing that it exists. A silent gap, produced by following the instructions. */}
                <p style={{ margin: 0, color: "#8a6d00" }}>
                  Note: this removes the row, <strong>not the permission</strong>. The group keeps
                  this access on the folder, and Folder Reconciliation will <strong>not</strong> take
                  it away — it only adds grants. To actually withdraw it, open the folder in
                  SharePoint, choose <strong>Manage Access → Advanced → Delete unique permissions</strong>,
                  then run Folder Reconciliation to rebuild the folder from the rows that remain.
                  Deleting the group instead removes all of its access at once.
                </p>
              </div>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", padding: 12 }}>
                <button style={s.ghost} disabled={busy} onClick={() => setConfirmDel(undefined)}>Cancel</button>
                <button style={s.delBtn} disabled={busy} onClick={() => { onDelete(confirmDel).catch(() => undefined); }}>
                  Delete mapping
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>{toast.message}</div>
      )}
    </div>
  );
}
