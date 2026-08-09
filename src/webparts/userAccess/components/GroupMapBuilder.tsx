import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  buildGroupMapRow,
  isDuplicateRow,
  validateDraft,
  roleFromGroupName,
  suggestGroupName,
  GroupMapRole,
  GroupMapDraft,
  GroupMapWriteRow,
  normalizeScope,
  siteEntryGroupTitle,
  PERSONAS,
  PERSONA_FAMILIES,
  personaByKey,
  personaTouchesStaging,
  Persona,
  normalizeRoleValue,
} from "../../../shared/groupMapModel";
import {
  searchSiteGroups,
  fetchAllSiteGroups,
  createSiteGroup,
  deleteSiteGroup,
  getGroupMembers,
  addGroupMember,
  removeGroupMember,
  searchTenantPeople,
  DUPLICATE_GROUP,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
import {
  toCsv,
  downloadCsv,
  exportFileName,
  GroupExportRow,
  NO_MEMBERS,
} from "../../../shared/groupExportCsv";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames, permissionLevelNames } from "../../../shared/spNaming";

type Props = { context: WebPartContext; siteUrl: string };
type GroupPick = { id: string; displayName: string };
type ModePick = { label: string; termSetGuid: string };
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
  ddCreate:   { color: "#0f6c3f", fontWeight: 600, borderBottom: "none", borderTop: "1px solid #e1e1e1", background: "#f6fbf8" },
  createPanel:{ border: "1px solid #b7dcc4", borderRadius: 6, padding: 14, background: "#f3faf5" },
  createHead: { fontWeight: 600, fontSize: 13, marginBottom: 2, color: "#0f6c3f" },
  hint:       { fontSize: 11, color: "#666", marginTop: 8, lineHeight: 1.45 },
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
  secondaryBtn:{ marginTop: 12, padding: "7px 18px", fontSize: 13, background: "#fff", color: "#242424", border: "1px solid #c7c7c7", borderRadius: 4, cursor: "pointer" },
  req:        { color: "#a4262c", marginLeft: 2 },
  missing:    { marginTop: 8, fontSize: 12, color: "#a4262c" },
  dangerBox:  { marginTop: 8, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 4, fontSize: 12, color: "#a4262c", lineHeight: 1.5 },
  modalOverlay:{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 },
  modalBox:   { background: "#fff", borderRadius: 8, width: "100%", maxWidth: 480, maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 30px rgba(0,0,0,.25)" },
  modalHead:  { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", borderBottom: "1px solid #eee", fontWeight: 600, fontSize: 14 },
  modalBody:  { padding: "12px 16px", overflowY: "auto" },
  modalFoot:  { padding: "10px 16px", borderTop: "1px solid #eee", display: "flex", justifyContent: "flex-end" },
  table:      { width: "100%", borderCollapse: "collapse", fontSize: 12, marginTop: 8 },
  th:         { textAlign: "left", padding: "6px 8px", borderBottom: "2px solid #e1e1e1", fontWeight: 600, color: "#555" },
  td:         { padding: "6px 8px", borderBottom: "1px solid #f0f0f0", verticalAlign: "top" },
  delBtn:     { padding: "3px 10px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  ghost:      { padding: "3px 10px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  toast:      { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 6, color: "#fff", fontSize: 13, zIndex: 50, boxShadow: "0 4px 14px rgba(0,0,0,.18)" },
  mono:       { fontFamily: "Consolas, monospace", fontSize: 11, color: "#666" },
  staleBadge: { display: "inline-block", marginLeft: 8, padding: "1px 7px", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em", color: "#a4262c", background: "#fde7e9", border: "1px solid #f1b0b3", borderRadius: 10, verticalAlign: "middle" },
};

export default function GroupMapBuilder({ context, siteUrl }: Props): React.ReactElement {
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
  const [canManage, setCanManage] = useState<boolean | undefined>(undefined); // undefined = still checking
  // The three custom permission levels, named for THIS site. Hardcoding "DMS Upload" here while
  // reconciliation grants "CRS Upload" tells an admin to create a level nothing will ever use.
  const [levels, setLevels] = useState<{ upload: string; approve: string; del: string }>({
    upload: "Upload", approve: "Approve", del: "Delete",
  });
  // The role reference below is four dense paragraphs. Collapsed by default so the FORM is the
  // first thing on the page — the client's "too complicated" was partly this wall of text.
  const [rolesOpen, setRolesOpen] = useState(false);

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

  // Inline group create
  const [creating, setCreating]   = useState(false);
  const [newName, setNewName]     = useState("");

  // Delete-group confirmation (destructive — two-step).
  const [confirmDeleteGroup, setConfirmDeleteGroup] = useState(false);

  // People staged to be added as members during a one-shot create (added after the
  // group is created). Kept separate from the existing-group member editor.
  const [stagedMembers, setStagedMembers] = useState<PersonPick[]>([]);

  // Cache of tier term GUID -> label, so the Existing mappings table shows the unit
  // name instead of a raw GUID. Populated lazily as rows load.
  const [tierLabels, setTierLabels] = useState<Record<string, string>>({});

  // Which of the two views is showing: the mapping table, or people-by-group.
  const [accessView, setAccessView] = useState<"mappings" | "people">("mappings");
  // Members per group id for the People view. Loaded lazily — one request per group,
  // and a site can carry a hundred groups, so fetching them all on tab open would
  // stall the view before the client had chosen anything to look at.
  const [peopleByGroup, setPeopleByGroup] = useState<Record<string, SpGroupMember[]>>({});
  const [peopleLoading, setPeopleLoading] = useState<Set<string>>(new Set());
  const [peopleOpen, setPeopleOpen] = useState<Set<string>>(new Set());
  const [peopleFilter, setPeopleFilter] = useState("");
  const [loadingAllPeople, setLoadingAllPeople] = useState<{ done: number; total: number } | undefined>(undefined);

  // Member-management modal (opened from an existing-mapping row). Self-contained so
  // it never touches the add-mapping form's segment/tier/role state.
  const [memberModal, setMemberModal]       = useState<GroupPick | null>(null);
  const [mmMembers, setMmMembers]           = useState<SpGroupMember[] | undefined>(undefined);
  const [mmBusy, setMmBusy]                 = useState(false);
  const [mmQuery, setMmQuery]               = useState("");
  const [mmResults, setMmResults]           = useState<PersonPick[]>([]);
  const [mmSearching, setMmSearching]       = useState(false);
  const [mmConfirmRemove, setMmConfirmRemove] = useState<number | undefined>(undefined);

  // Member editor
  const [membersOpen, setMembersOpen]         = useState(false);
  const [members, setMembers]                 = useState<SpGroupMember[] | undefined>(undefined);
  const [memberBusy, setMemberBusy]           = useState(false);
  const [peopleQuery, setPeopleQuery]         = useState("");
  const [peopleResults, setPeopleResults]     = useState<PersonPick[]>([]);
  const [peopleSearching, setPeopleSearching] = useState(false);
  const [confirmRemove, setConfirmRemove]     = useState<number | undefined>(undefined);

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
      .map((r) => ({ label: (r.ModeLabel ?? "").trim() || (r.TermSetGuid ?? "").trim(), termSetGuid: (r.TermSetGuid ?? "").trim() }));
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

  // Create/add/remove need Full Control. Detect up front so the controls are
  // disabled with an explanation instead of failing with a 403 on click.
  const loadCanManage = async (): Promise<boolean> => {
    const meRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/currentuser?$select=Id,IsSiteAdmin`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!meRes.ok) return false;
    const me = await meRes.json();
    if (me.IsSiteAdmin === true) return true;
    const ownRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/AssociatedOwnerGroup/Users?$filter=Id eq ${me.Id}&$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!ownRes.ok) return false;
    const own = await ownRes.json();
    return ((own.value ?? []) as unknown[]).length > 0;
  };

  // Resolve tier term GUIDs to labels for the Existing mappings table (raw GUIDs are
  // meaningless to the client). One read per distinct unit term; cached.
  const loadTierLabels = async (rows: ExistingRow[]): Promise<void> => {
    // Dedupe to distinct, not-yet-resolved unit terms up front so the async loop never
    // re-reads its own accumulator across an await (avoids a race-condition lint error).
    const seen = new Set<string>();
    const need = rows.filter((r) => {
      if (!r.UnitTermGuid || !r.Segment || r.UnitTermGuid === r.Segment) return false;
      if (tierLabels[r.UnitTermGuid] || seen.has(r.UnitTermGuid)) return false;
      seen.add(r.UnitTermGuid);
      return true;
    });
    const map: Record<string, string> = {};
    for (const r of need) {
      try {
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/v2.1/termStore/sets/${r.Segment}/terms/${r.UnitTermGuid}?$select=labels`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json" } },
        );
        if (res.ok) {
          const d = await res.json();
          const nm = (d.labels ?? [])[0]?.name as string | undefined;
          if (nm) map[r.UnitTermGuid] = nm;
        }
      } catch { /* leave as GUID */ }
    }
    if (Object.keys(map).length) setTierLabels((prev) => ({ ...prev, ...map }));
  };

  /* ── Mount ─────────────────────────────────────────────────────────────── */

  useEffect(() => {
    // Names FIRST. Every read below goes through cachedListTitle, and an unprimed cache resolves
    // to the legacy DMS titles — which 404 on a CRS-renamed site and would present as "the config
    // could not be read" rather than "the list is called something else".
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      .then(() => permissionLevelNames(context.spHttpClient, siteUrl))
      .then((n) => setLevels(n))
      .catch(() => undefined)
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
        loadCanManage().then(setCanManage).catch(() => setCanManage(false));
      })
      .catch(() => undefined);
  }, []);

  // Whenever the mappings change, resolve any new tier labels.
  useEffect(() => {
    if (existing.length) loadTierLabels(existing).catch(() => undefined);
  }, [existing]);

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

  /* ── People search for the member editor (debounced) ───────────────────── */

  useEffect(() => {
    const q = peopleQuery.trim();
    // Runs for BOTH the existing-group member editor (membersOpen) and the create
    // panel's "add members" box (creating) — the two are never open at once.
    if ((!membersOpen && !creating) || q.length < 2) { setPeopleResults([]); setPeopleSearching(false); return; }
    let cancelled = false;
    setPeopleSearching(true);
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then((r) => { if (!cancelled) { setPeopleResults(r); setPeopleSearching(false); } })
        .catch(() => { if (!cancelled) { setPeopleResults([]); setPeopleSearching(false); } });
    }, 350);
    return () => { cancelled = true; clearTimeout(t); };
  }, [peopleQuery, membersOpen, creating]);

  // People search for the member modal (debounced).
  useEffect(() => {
    const q = mmQuery.trim();
    if (!memberModal || q.length < 2) { setMmResults([]); setMmSearching(false); return; }
    let cancelled = false;
    setMmSearching(true);
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then((r) => { if (!cancelled) { setMmResults(r); setMmSearching(false); } })
        .catch(() => { if (!cancelled) { setMmResults([]); setMmSearching(false); } });
    }, 350);
    return () => { cancelled = true; clearTimeout(t); };
  }, [mmQuery, memberModal]);

  /* ── Selection handlers ────────────────────────────────────────────────── */

  // Wipe the member-editor state so a freshly picked/cleared group never shows
  // the previous group's members.
  const resetMemberState = (): void => {
    setMembersOpen(false);
    setMembers(undefined);
    setPeopleQuery("");
    setPeopleResults([]);
    setConfirmRemove(undefined);
  };

  const pickGroup = (g: GroupPick): void => {
    setGroup(g);
    setQuery(g.displayName);
    setResults([]);
    setCreating(false);
    setConfirmDeleteGroup(false);
    resetMemberState();
    // Pre-select the role implied by the name suffix (_UPL/_APR); admin can override.
    setRole(roleFromGroupName(g.displayName));
  };

  const clearGroup = (): void => {
    setGroup(undefined);
    setQuery("");
    setResults([]);
    setCreating(false);
    setConfirmDeleteGroup(false);
    resetMemberState();
  };

  // Delete the whole SharePoint group. Also removes any DMS Group Map rows that
  // reference it, so no orphan mappings are left behind. Destructive — gated behind
  // a two-step confirm in the UI.
  const onDeleteGroup = async (): Promise<void> => {
    if (!group) return;
    setBusy(true);
    try {
      const gid = Number(group.id);
      const name = group.displayName;
      const rows = existing.filter((r) => Number(r.GroupId) === gid);
      for (const r of rows) await deleteRow(r.itemId);
      await deleteSiteGroup(context.spHttpClient, siteUrl, gid);
      clearGroup();
      setExisting(await loadExisting());
      showToast(
        `Group "${name}" deleted${rows.length ? ` (+${rows.length} mapping row(s))` : ""} — re-run Folder Reconciliation to refresh folder permissions.`,
        false,
      );
    } catch (e) {
      showToast(`Delete group failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  // Member management modal — fully separate from the add-mapping form, so editing a
  // group's members never disturbs the segment/tier/role selections below.
  const openMemberModal = (r: ExistingRow): void => {
    const g = { id: r.GroupId, displayName: r.GroupName || r.GroupId };
    setMemberModal(g);
    setMmMembers(undefined);
    setMmQuery("");
    setMmResults([]);
    setMmConfirmRemove(undefined);
    getGroupMembers(context.spHttpClient, siteUrl, Number(g.id))
      .then(setMmMembers)
      .catch(() => { setMmMembers([]); showToast("Could not load members.", true); });
  };
  const closeMemberModal = (): void => {
    setMemberModal(null);
    setMmMembers(undefined);
    setMmQuery("");
    setMmResults([]);
    setMmConfirmRemove(undefined);
  };
  const reloadMm = async (): Promise<void> => {
    if (memberModal) setMmMembers(await getGroupMembers(context.spHttpClient, siteUrl, Number(memberModal.id)));
  };
  // Session cache for the site-entry group lookup. Cache the PROMISE (assigned
  // synchronously) so there is no read-modify-write of a ref across an await.
  // Resolves to the group's id, or null if DMS_SITE_MEMBERS doesn't exist yet.
  const entryGroupPromiseRef = useRef<Promise<number | null> | undefined>(undefined);
  const resolveEntryGroupId = (): Promise<number | null> => {
    if (!entryGroupPromiseRef.current) {
      entryGroupPromiseRef.current = fetchAllSiteGroups(context.spHttpClient, siteUrl)
        .catch(() => [])
        .then((all) => {
          const hit = all.find(
            (g) => g.title.trim().toLowerCase() === siteEntryGroupTitle().toLowerCase(),
          );
          return hit ? hit.id : null;
        });
    }
    return entryGroupPromiseRef.current;
  };

  const mmAdd = async (p: PersonPick): Promise<void> => {
    if (!memberModal) return;
    setMmBusy(true);
    try {
      const targetId = Number(memberModal.id);
      await addGroupMember(context.spHttpClient, siteUrl, targetId, p.loginName);
      // Also grant site entry: every DMS group member needs to be able to open the
      // site (folder-group Limited Access alone can't — site-entry-access-layer spec).
      // Skip when the group being edited IS the entry group; re-adding is idempotent.
      let entryNote = "";
      const isEntryGroup =
        memberModal.displayName.trim().toLowerCase() === siteEntryGroupTitle().toLowerCase();
      if (!isEntryGroup) {
        const entryId = await resolveEntryGroupId();
        if (entryId === null) {
          entryNote = ` (note: ${siteEntryGroupTitle()} not found — create it so they can open the site)`;
        } else {
          await addGroupMember(context.spHttpClient, siteUrl, entryId, p.loginName).catch(
            () => {
              entryNote = ` (warning: could not add to ${siteEntryGroupTitle()})`;
            },
          );
        }
      }
      await reloadMm();
      setMmQuery("");
      setMmResults([]);
      showToast(`${p.displayName} added — access is immediate.${entryNote}`, !!entryNote);
    } catch (e) {
      showToast(`Add member failed: ${(e as Error).message}`, true);
    } finally {
      setMmBusy(false);
    }
  };
  const mmRemove = async (userId: number): Promise<void> => {
    if (!memberModal) return;
    setMmBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, Number(memberModal.id), userId);
      await reloadMm();
      setMmConfirmRemove(undefined);
      showToast("Member removed — access revoked immediately.", false);
    } catch (e) {
      showToast(`Remove failed: ${(e as Error).message}`, true);
    } finally {
      setMmBusy(false);
    }
  };

  /* ── Inline group create ───────────────────────────────────────────────── */

  const startCreate = (): void => {
    setCreating(true);
    setStagedMembers([]);
    setPeopleQuery("");
    setPeopleResults([]);
    setNewName(
      query.trim() ||
        suggestGroupName(mode?.label ?? "", chosen.map((t) => t.label), (role || "") as GroupMapRole | ""),
    );
  };

  const cancelCreate = (): void => {
    setCreating(false);
    setNewName("");
    setStagedMembers([]);
    setPeopleQuery("");
    setPeopleResults([]);
  };

  // Stage / unstage a person to be added as a member when the group is created.
  const stageMember = (p: PersonPick): void => {
    setStagedMembers((prev) => prev.some((m) => m.loginName === p.loginName) ? prev : [...prev, p]);
    setPeopleQuery("");
    setPeopleResults([]);
  };
  const unstageMember = (loginName: string): void => {
    setStagedMembers((prev) => prev.filter((m) => m.loginName !== loginName));
  };

  // Warn (never block) when the typed name's suffix disagrees with the selected Role.
  const nameRoleMismatch = (): boolean => {
    if (!newName.trim() || !role || role === "GLOBAL") return false;
    return roleFromGroupName(newName) !== role;
  };

  // What the create panel still needs before its bottom button lights up. Mirrors
  // validateDraft, but keyed on the typed name (the group doesn't exist yet).
  const createErrors: string[] = [];
  if (!newName.trim()) createErrors.push("Enter a group name.");
  if (!role) createErrors.push("Select a role.");
  if (role && role !== "GLOBAL") {
    if (!mode) createErrors.push("Select a segment.");
    // A segment-scope persona has no tier control — pickMode sets tierGuid to the term-set
    // GUID for it. Telling the admin to "select a tier" would name a dropdown that is not
    // on the screen, which reads as the page being broken rather than as a missing step.
    if (!tierGuid && personaByKey(persona)?.scope !== "segment") createErrors.push("Select a tier.");
  }

  // One-shot create: make the site group AND write its mapping row in a single
  // action. Called from the create panel's bottom button, which is only enabled
  // once the name + role + segment + tier are all chosen. On success we land on
  // the freshly created group so the admin can add its members next.
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
   * Declared above both call sites deliberately — onCreateAndMap is the first.
   */
  const draftRows = (d: GroupMapDraft): GroupMapWriteRow[] => {
    const p = personaByKey(persona);
    const roles: GroupMapRole[] =
      p && !p.unavailable && p.roles.length > 0 ? p.roles : ([d.role].filter(Boolean) as GroupMapRole[]);
    return roles.map((r) => buildGroupMapRow({ ...d, role: r }));
  };

  const onCreateAndMap = async (): Promise<void> => {
    const title = newName.trim();
    if (!title || createErrors.length > 0) return;
    setBusy(true);
    try {
      const created = await createSiteGroup(context.spHttpClient, siteUrl, title);
      const newGroup = { id: String(created.id), displayName: created.title };
      // The group now exists. If writing the mapping fails, roll it back so a failed
      // "one shot" never leaves an orphan group behind (which would then collide with
      // a retry as "already exists"). Keeps the action atomic.
      try {
        // EVERY row the persona needs, not just the first. A Head of Unit is APR + DELS,
        // and the admin never sees those names — see draftRows. If any row fails the whole
        // create is rolled back below, which is the right trade here and not in onAdd: the
        // group did not exist a moment ago, so removing it restores the exact prior state
        // rather than destroying something.
        for (const row of draftRows({
          groupId: newGroup.id,
          groupName: newGroup.displayName,
          role: role as GroupMapRole,
          segmentGuid: mode?.termSetGuid,
          tierGuid,
        })) {
          await postRow(row);
        }
      } catch (mapErr) {
        await deleteSiteGroup(context.spHttpClient, siteUrl, Number(newGroup.id)).catch(() => undefined);
        throw mapErr;
      }
      setExisting(await loadExisting());
      // Add any staged members now that the group exists. Failures are counted, not fatal.
      // Each member is also added to the site-entry group so they can open the site
      // (site-entry-access-layer spec). A new DMS group is never the entry group itself.
      const entryId: number | null = stagedMembers.length > 0 ? await resolveEntryGroupId() : null;
      let addedMembers = 0;
      let failedMembers = 0;
      for (const p of stagedMembers) {
        try { await addGroupMember(context.spHttpClient, siteUrl, Number(newGroup.id), p.loginName); addedMembers++; }
        catch { failedMembers++; }
        if (entryId !== null) {
          await addGroupMember(context.spHttpClient, siteUrl, entryId, p.loginName).catch(() => undefined);
        }
      }
      const staged = stagedMembers.length;
      const createdName = newGroup.displayName;
      // Reset the whole form back to the default (empty search) state so the admin can
      // immediately create another group — no leftover group/role/segment selections.
      clearGroup();
      setNewName("");
      setStagedMembers([]);
      // Persona resets with the rest. Leaving it selected was worse than untidy: the
      // panel kept showing "needs N mappings at the <tier> tier" with its ✓/○ list,
      // now describing a group that is no longer selected — so the next group created
      // would be checked off against the previous group's progress.
      setRole(""); setMode(undefined); setCascade([]); setChosen([]); setTierGuid(""); setPersona("");
      const memberNote = staged === 0
        ? " — edit members from its row, then run Folder Reconciliation."
        : failedMembers === 0
          ? ` with ${addedMembers} member(s) — run Folder Reconciliation.`
          : ` — ${addedMembers} member(s) added, ${failedMembers} failed. Run Folder Reconciliation.`;
      showToast(`Group "${createdName}" created & mapped${memberNote}`, failedMembers > 0);
    } catch (e) {
      const msg = (e as Error).message;
      if (msg === DUPLICATE_GROUP) {
        showToast("A group with that name already exists — use Back to search and select it instead.", true);
      } else if (msg.indexOf("does not exist") !== -1) {
        showToast(`Couldn't write the mapping: the '${GROUP_MAP_LIST()}' list is missing. No group was created — create/rename that list, then try again.`, true);
      } else {
        showToast(`Create failed: ${msg} — no group was left behind.`, true);
      }
    } finally {
      setBusy(false);
    }
  };

  /* ── Member editor ─────────────────────────────────────────────────────── */

  const groupIdNum = (): number => Number(group?.id ?? 0);

  const reloadMembers = async (): Promise<void> => {
    setMembers(await getGroupMembers(context.spHttpClient, siteUrl, groupIdNum()));
  };

  const toggleMembers = (): void => {
    const opening = !membersOpen;
    setMembersOpen(opening);
    if (opening && members === undefined) {
      reloadMembers().catch(() => { setMembers([]); showToast("Could not load members.", true); });
    }
  };

  const onAddMember = async (p: PersonPick): Promise<void> => {
    setMemberBusy(true);
    try {
      await addGroupMember(context.spHttpClient, siteUrl, groupIdNum(), p.loginName);
      await reloadMembers();
      setPeopleQuery("");
      setPeopleResults([]);
      showToast(`${p.displayName} added — access is immediate.`, false);
    } catch (e) {
      showToast(`Add member failed: ${(e as Error).message}`, true);
    } finally {
      setMemberBusy(false);
    }
  };

  const onRemoveMember = async (userId: number): Promise<void> => {
    setMemberBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, groupIdNum(), userId);
      await reloadMembers();
      setConfirmRemove(undefined);
      showToast("Member removed — access revoked immediately.", false);
    } catch (e) {
      showToast(`Remove failed: ${(e as Error).message}`, true);
    } finally {
      setMemberBusy(false);
    }
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

  /** Same text the Tier column shows, reused so the CSV and the screen never disagree. */
  const tierLabelFor = (r: ExistingRow): string => {
    if (!r.UnitTermGuid) return "";
    if (r.UnitTermGuid === r.Segment) return "(segment level)";
    return tierLabels[r.UnitTermGuid] ?? r.UnitTermGuid;
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
      const base = {
        group: r.GroupName || r.GroupId,
        segment: r.Segment ? segmentLabelFor(r.Segment) : "",
        tier: tierLabelFor(r),
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
   * Delete a mapping row — and, when it was that group's LAST row, the SharePoint group
   * with it.
   *
   * Why the group goes too: nothing in this system revokes a role assignment, so a group
   * left behind with no mappings keeps every grant it already holds on every folder. The
   * Group Map then says the group has no access while SharePoint says it has all of it,
   * and the question that arrives later is "why can this person still see it". Deleting
   * the group is what actually removes the access, because the principal ceases to exist.
   *
   * Why only on the LAST row: a group is normally several rows — a Head of Unit 4 group is
   * five. Removing the DELS row because that one capability is no longer wanted must not
   * take the group and its members with it.
   *
   * The residual case, which this cannot fix and does not pretend to: deleting ONE of
   * several rows leaves that specific grant on the folders, because the group survives.
   * The toast says so rather than implying the change has taken effect.
   */
  const onDelete = async (itemId: number): Promise<void> => {
    const row = existing.find((r) => r.itemId === itemId);
    const gid = Number(row?.GroupId ?? NaN);
    const siblings = existing.filter((r) => Number(r.GroupId) === gid && r.itemId !== itemId);
    const lastRowForGroup = !!row && !isNaN(gid) && gid > 0 && siblings.length === 0;
    setBusy(true);
    try {
      await deleteRow(itemId);
      let groupDeleted = false;
      if (lastRowForGroup) {
        try {
          await deleteSiteGroup(context.spHttpClient, siteUrl, gid);
          groupDeleted = true;
        } catch (e) {
          // The row is already gone. Report the group failure plainly instead of
          // rolling back — re-creating the row would leave a mapping the admin just
          // deleted, and silence here is what produces the leftover-access complaint.
          showToast(`Row deleted, but the group "${row?.GroupName}" could not be removed: ${(e as Error).message} — delete it from Site permissions, or it keeps its folder access.`, true);
        }
      }
      setExisting(await loadExisting());
      setSelected((prev) => { const n = new Set(prev); n.delete(itemId); return n; });
      setConfirmDel(undefined);
      if (groupDeleted) {
        showToast(`Row deleted, and group "${row?.GroupName}" removed with it — that was its last mapping, so its folder access is gone.`, false);
      } else if (lastRowForGroup) {
        // Toast already shown by the catch above.
      } else {
        showToast(`Row deleted. "${row?.GroupName}" still has ${siblings.length} other mapping(s), so the group is kept — it also keeps the folder grant from this row until it is removed by hand.`, false);
      }
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

  const allSelected = existing.length > 0 && selected.size === existing.length;

  const toggleAll = (): void => {
    setSelected(allSelected ? new Set() : new Set(existing.map((r) => r.itemId)));
  };

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

  // Role + Segment + Tier selectors. Shared by both flows: rendered INSIDE the
  // create panel (one-shot create) and below an already-picked group (add mapping).
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

  /* ── People view ──────────────────────────────────────────────────────────────
     One entry per GROUP, not per mapping row: a group commonly carries several
     mappings (one per tier it covers), and listing the same people once per row is
     what makes "who can reach this" hard to read in the table. */
  type GroupEntry = {
    groupId: string;
    groupName: string;
    /** Every mapping this group appears in, shown as chips on the header row. */
    grants: Array<{ segment: string; tier: string; role: string }>;
  };
  const distinctGroups: GroupEntry[] = (() => {
    const byId = new Map<string, GroupEntry>();
    existing.forEach((r) => {
      if (!r.GroupId) return;
      const entry = byId.get(r.GroupId) ?? {
        groupId: r.GroupId,
        groupName: r.GroupName || r.GroupId,
        grants: [],
      };
      entry.grants.push({
        segment: r.Segment ? segmentLabelFor(r.Segment) : "",
        tier: tierLabelFor(r),
        role: normalizeRoleValue(r.Role ?? ""),
      });
      byId.set(r.GroupId, entry);
    });
    const out: GroupEntry[] = [];
    byId.forEach((v) => out.push(v));
    return out.sort((a, b) => a.groupName.localeCompare(b.groupName));
  })();

  const fetchGroupPeople = async (groupId: string): Promise<void> => {
    if (peopleByGroup[groupId]) return; // cached
    setPeopleLoading((prev) => { const n = new Set(prev); n.add(groupId); return n; });
    try {
      const m = await getGroupMembers(context.spHttpClient, siteUrl, Number(groupId));
      setPeopleByGroup((prev) => ({ ...prev, [groupId]: m }));
    } catch {
      // Cached as empty so a failed group does not re-request on every repaint.
      setPeopleByGroup((prev) => ({ ...prev, [groupId]: [] }));
      showToast("Could not load members for that group.", true);
    } finally {
      setPeopleLoading((prev) => { const n = new Set(prev); n.delete(groupId); return n; });
    }
  };

  const toggleGroupOpen = (groupId: string): void => {
    setPeopleOpen((prev) => {
      const n = new Set(prev);
      if (n.has(groupId)) n.delete(groupId);
      else { n.add(groupId); fetchGroupPeople(groupId).catch(() => undefined); }
      return n;
    });
  };

  /** Sequential, not parallel — a burst of member reads is what trips the 429 throttle. */
  const loadAllPeople = async (): Promise<void> => {
    const need = distinctGroups.filter((g) => !peopleByGroup[g.groupId]);
    setLoadingAllPeople({ done: 0, total: need.length });
    for (let i = 0; i < need.length; i++) {
      await fetchGroupPeople(need[i].groupId);
      setLoadingAllPeople({ done: i + 1, total: need.length });
    }
    setLoadingAllPeople(undefined);
    setPeopleOpen(new Set(distinctGroups.map((g) => g.groupId)));
  };

  /** Case-insensitive substring match, with the hit wrapped so the eye can find it. */
  const highlight = (text: string, q: string): React.ReactNode => {
    if (q.length === 0) return text;
    const at = text.toLowerCase().indexOf(q);
    if (at === -1) return text;
    return (
      <>
        {text.slice(0, at)}
        <mark style={{ background: "#fff3bf", color: "inherit", padding: "0 1px" }}>
          {text.slice(at, at + q.length)}
        </mark>
        {text.slice(at + q.length)}
      </>
    );
  };

  const renderPeopleView = (): React.ReactElement => {
    const q = peopleFilter.trim().toLowerCase();
    const personMatches = (p: SpGroupMember): boolean =>
      (p.title ?? "").toLowerCase().indexOf(q) !== -1 ||
      (p.email ?? "").toLowerCase().indexOf(q) !== -1;
    // A person filter can only see groups already read, so it says so rather than
    // silently reporting "no match" for groups it never looked inside.
    const unloaded = distinctGroups.filter((g) => !peopleByGroup[g.groupId]).length;
    const shown = q.length === 0
      ? distinctGroups
      : distinctGroups.filter((g) => {
          if (g.groupName.toLowerCase().indexOf(q) !== -1) return true;
          const m = peopleByGroup[g.groupId];
          return !!m && m.some(personMatches);
        });
    // How many people the query found overall — searching a name and being told
    // "3 groups" answers a question nobody asked. The count is what confirms the
    // search ran at all.
    const matchedPeople = q.length === 0
      ? 0
      : shown.reduce((n, g) => n + (peopleByGroup[g.groupId] ?? []).filter(personMatches).length, 0);
    return (
      <div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
          <input
            style={{ ...s.input, maxWidth: 320 }}
            placeholder="Filter by person or group name…"
            value={peopleFilter}
            onChange={(e) => setPeopleFilter(e.target.value)}
          />
          <button
            style={s.ghost}
            disabled={busy || loadingAllPeople !== undefined || unloaded === 0}
            title="Read the members of every mapped group, one at a time"
            onClick={() => { loadAllPeople().catch(() => setLoadingAllPeople(undefined)); }}
          >
            {loadingAllPeople
              ? `Loading… ${loadingAllPeople.done}/${loadingAllPeople.total}`
              : unloaded === 0 ? "All members loaded" : `Load all members (${unloaded})`}
          </button>
          {q.length > 0 && unloaded > 0 && (
            <span style={{ fontSize: 12, color: "#b45309" }}>
              {unloaded} group(s) not read yet — load all to search inside them.
            </span>
          )}
        </div>
        {q.length > 0 && shown.length > 0 && (
          <p style={{ fontSize: 12, color: "#0f6c3f", margin: "0 0 10px", fontWeight: 600 }}>
            {matchedPeople > 0
              ? `${matchedPeople} person match${matchedPeople === 1 ? "" : "es"} across ${shown.length} group${shown.length === 1 ? "" : "s"}`
              : `${shown.length} group${shown.length === 1 ? "" : "s"} match by name`}
          </p>
        )}
        {distinctGroups.length === 0 ? (
          <p style={{ fontSize: 13, color: "#999" }}>No groups are mapped yet.</p>
        ) : shown.length === 0 ? (
          <p style={{ fontSize: 13, color: "#999" }}>No group or person matches &ldquo;{peopleFilter}&rdquo;.</p>
        ) : (
          shown.map((g) => {
            const loading = peopleLoading.has(g.groupId);
            const members = peopleByGroup[g.groupId];
            const nameHit = q.length > 0 && g.groupName.toLowerCase().indexOf(q) !== -1;
            // A collapsed group that matched on a PERSON showed a header and nothing
            // else — the search looked broken because the thing it found was hidden.
            const open = q.length > 0 || peopleOpen.has(g.groupId);
            // Searching a person and being handed the whole roster to scan is the
            // other half of that. When the query matched people rather than the group
            // NAME, show the people it matched and count the rest.
            const hits = members && q.length > 0 && !nameHit
              ? members.filter(personMatches)
              : undefined;
            const listed = hits ?? members ?? [];
            const hidden = hits && members ? members.length - hits.length : 0;
            return (
              <div key={g.groupId} style={{ border: "1px solid #e5e5e5", borderRadius: 6, marginBottom: 8, overflow: "hidden" }}>
                <div
                  role="button"
                  tabIndex={0}
                  onClick={() => toggleGroupOpen(g.groupId)}
                  onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggleGroupOpen(g.groupId); } }}
                  style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", cursor: "pointer", background: open ? "#f4f8f5" : "#fff" }}
                >
                  <span style={{ color: "#0f6c3f", fontSize: 11 }}>{open ? "▾" : "▸"}</span>
                  <span style={{ fontWeight: 600, fontSize: 13 }}>{highlight(g.groupName, q)}</span>
                  <span style={{ fontSize: 11, color: "#666" }}>
                    {!members
                      ? "—"
                      : hits
                        ? `${hits.length} of ${members.length} match`
                        : `${members.length} member${members.length === 1 ? "" : "s"}`}
                  </span>
                  <span style={{ marginLeft: "auto", display: "flex", gap: 4, flexWrap: "wrap", justifyContent: "flex-end" }}>
                    {g.grants.map((gr, i) => (
                      <span key={i} style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".04em", color: "#0f6c3f", background: "#e8f5ee", border: "1px solid #b3d9c4", borderRadius: 10, padding: "2px 8px" }}>
                        {gr.role}{gr.tier ? ` · ${gr.tier}` : ""}
                      </span>
                    ))}
                  </span>
                </div>
                {open && (
                  <div style={{ borderTop: "1px solid #eee", padding: "6px 12px 10px 30px" }}>
                    {loading ? (
                      <div style={{ fontSize: 12, color: "#666", padding: "4px 0" }}>Loading members&hellip;</div>
                    ) : !members || members.length === 0 ? (
                      <div style={{ fontSize: 12, color: "#b45309", padding: "4px 0" }}>
                        No members — anyone relying on this group has no access.
                      </div>
                    ) : (
                      <>
                        {listed.map((p) => (
                          <div key={p.id} style={{ display: "flex", gap: 10, alignItems: "baseline", fontSize: 12, padding: "3px 0" }}>
                            <span style={{ fontWeight: 600 }}>{highlight(p.title ?? "", q)}</span>
                            <span style={{ color: "#666" }}>{highlight(p.email ?? "", q)}</span>
                          </div>
                        ))}
                        {hidden > 0 && (
                          <div style={{ fontSize: 11, color: "#999", padding: "4px 0 0" }}>
                            + {hidden} other member{hidden === 1 ? "" : "s"} in this group
                          </div>
                        )}
                      </>
                    )}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    );
  };

  return (
    <div style={s.wrap}>
      <p style={s.intro}>
        Map a <strong>native SharePoint site group</strong> to a segment, tier, and role — or create
        the group right here and add its members. This writes a clean row into the{" "}
        <strong>{GROUP_MAP_LIST()}</strong> list.
      </p>

      {/* Collapsed by default. Everything below is reference an admin needs ONCE (when the
          permission levels are first created, or when choosing between DEL and DELS) and never
          again — but it used to sit above the form on every visit. */}
      <button
        onClick={() => setRolesOpen((v) => !v)}
        style={{
          background: "none", border: "none", padding: "0 0 10px", cursor: "pointer",
          fontFamily: "'Segoe UI', sans-serif", fontSize: 13, fontWeight: 600, color: "#0f6c3f",
        }}
      >
        {rolesOpen ? "▾" : "▸"} How roles and permission levels work
      </button>

      {rolesOpen && (
        <>
          <p style={s.intro}>
            Member changes take effect <strong>immediately</strong>; new or deleted{" "}
            <em>rows</em> need a <strong>Folder Reconciliation</strong> run to apply folder
            permissions (MEMBER → Read, UPL → {levels.upload}, APR → {levels.approve},
            DEL → {levels.del} on Documents, DELS → {levels.del} on the approval library).
          </p>
          {/* Named on the screen because the failure is quiet: reconciliation warns and
              skips the assignment, so an approver simply never gains the level and the
              run still reports success overall.

              Upload is called out separately from the other two because its failure
              mode is worse. APR and DEL previously had no level at all, so a missing one
              only meant "not yet in effect". UPL used to point at Contribute, which always
              exists — so once it points at the custom level, a site without that level gives
              every NEWLY provisioned unit no uploader grant whatsoever. Existing uploaders
              keep the Contribute grant already on their folder (nothing revokes), which is
              exactly why nobody notices until a new unit is onboarded.

              The names are RESOLVED, not hardcoded: telling an admin to create "DMS Upload"
              on a site where reconciliation grants "CRS Upload" sends them to build a level
              nothing will ever use, and the resulting log line looks identical. */}
          <p style={s.intro}>
            <strong>{levels.upload}</strong>, <strong>{levels.approve}</strong> and{" "}
            <strong>{levels.del}</strong> are custom permission levels an administrator creates
            once per site (Site settings → Site permissions → Permission levels): copy{" "}
            <em>Contribute</em> and untick Delete Items; copy <em>Contribute</em>, tick Approve
            Items and untick Add Items, Delete Items and Delete Versions; copy <em>Read</em> and
            tick Delete Items. Until they exist, reconciliation reports{" "}
            <em>no &quot;{levels.approve}&quot; role definition on site</em> and skips those
            grants — nobody loses access, but approve-only and delete do not take effect, and a
            newly provisioned unit gets no uploader grant at all.
          </p>
          {/* DEL vs DELS is one letter for two different libraries, so it is spelled out
              here as well as in the role tooltips. An admin who picks DEL intending
              "can clear out junk in the approval library" grants delete over APPROVED
              documents instead, and the run log looks identical either way. */}
          <p style={s.intro}>
            <strong>DEL</strong> deletes <em>approved</em> documents in the Documents library.{" "}
            <strong>DELS</strong> deletes <em>pending</em> files in the approval library. They
            share one permission level and differ only in which library they are allowed to
            reach, so picking the wrong one grants delete over the wrong set of documents.
            Uploaders (<strong>UPL</strong>) cannot delete at all — that is deliberate, so a PIC
            must ask a head of unit.
          </p>
        </>
      )}

      {canManage === false && (
        <div style={{ ...s.card, borderColor: "#f0c000", background: "#fff8e1" }}>
          Read-only: creating groups and editing members needs <strong>Full Control (site owner)</strong>{" "}
          on this site. You can still view mappings.
        </div>
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

      <div style={s.card}>
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
          <>
            <span style={s.pickedChip}>
              {group.displayName}
              <button style={s.chipX} disabled={busy} title="Change" onClick={clearGroup}>✕</button>
            </span>
            <button style={{ ...s.seglvl, marginLeft: 14 }} disabled={busy} onClick={toggleMembers}>
              {members === undefined ? "members" : `${members.length} member(s)`} {membersOpen ? "▴" : "▾"}
            </button>
            {canManage === true && (
              confirmDeleteGroup ? (
                <div style={s.dangerBox}>
                  <span>
                    ⚠ Permanently delete the SharePoint group <strong>{group.displayName}</strong>? This
                    removes the group and any of its DMS Group Map rows and folder permissions across the
                    site. Its <em>members</em> (the users) are not deleted. This cannot be undone.
                  </span>
                  <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                    <button style={s.delBtn} disabled={busy} onClick={() => { onDeleteGroup().catch(() => undefined); }}>Yes, delete group</button>
                    <button style={s.ghost} disabled={busy} onClick={() => setConfirmDeleteGroup(false)}>Cancel</button>
                  </div>
                </div>
              ) : (
                <button style={{ ...s.seglvl, marginLeft: 14, color: "#a4262c", textDecoration: "underline" }} disabled={busy} onClick={() => setConfirmDeleteGroup(true)}>
                  delete group
                </button>
              )
            )}
            {membersOpen && (
              <div style={{ ...s.preview, borderStyle: "solid", marginTop: 8 }}>
                {members === undefined && <div>Loading members…</div>}
                {members !== undefined && members.length === 0 && <div>No members yet.</div>}
                {(members ?? []).map((m) => (
                  <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "3px 0" }}>
                    <span style={{ flex: 1 }}>{m.title}</span>
                    <span style={s.mono}>{m.email}</span>
                    {canManage === true && (confirmRemove === m.id ? (
                      <span style={{ display: "inline-flex", gap: 6 }}>
                        <button style={s.delBtn} disabled={memberBusy} onClick={() => { onRemoveMember(m.id).catch(() => undefined); }}>Remove</button>
                        <button style={s.ghost} disabled={memberBusy} onClick={() => setConfirmRemove(undefined)}>Cancel</button>
                      </span>
                    ) : (
                      <button style={s.chipX} disabled={memberBusy} title="Remove from group" onClick={() => setConfirmRemove(m.id)}>✕</button>
                    ))}
                  </div>
                ))}
                {canManage === true && (
                  <div style={{ ...s.ddwrap, marginTop: 8 }}>
                    <input
                      style={s.input}
                      placeholder="Search people in the tenant to add…"
                      value={peopleQuery}
                      disabled={memberBusy}
                      onChange={(e) => setPeopleQuery(e.target.value)}
                    />
                    {peopleQuery.trim().length >= 2 && (
                      <div style={s.dd}>
                        {peopleSearching && <div style={s.ddItem}>Searching…</div>}
                        {!peopleSearching && peopleResults.map((p) => (
                          <div key={p.loginName} style={s.ddItem} onClick={() => { onAddMember(p).catch(() => undefined); }}>
                            {p.displayName} <span style={s.mono}>{p.email}</span>
                          </div>
                        ))}
                        {!peopleSearching && peopleResults.length === 0 && (
                          <div style={{ ...s.ddItem, color: "#666" }}>No matching people.</div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
          </>
        ) : creating ? (
          /* Create mode — one shot. Name + role + segment + tier all live in this
             panel, with a single button at the bottom that creates the group AND
             writes its mapping. Search box/dropdown are not rendered, so there is
             never a duplicate name field. */
          <div style={s.createPanel}>
            <div style={s.createHead}>Create a new group</div>
            <label style={s.label}>New group name</label>
            <input
              style={s.input}
              value={newName}
              disabled={busy}
              autoFocus
              onChange={(e) => setNewName(e.target.value)}
              placeholder="e.g. GHO_GF_CORU_UPLOADER"
            />
            {nameRoleMismatch() && (
              <div style={{ fontSize: 12, color: "#a4262c", marginTop: 4 }}>
                Warning: the name suffix doesn&rsquo;t match the selected role ({role}). You can still create it.
              </div>
            )}

            {/* Role + Segment + Tier — chosen here, above the button. */}
            {selectionFields}

            {/* Members (optional) — staged now, added when the group is created. */}
            <label style={s.label}>Members (optional)</label>
            {stagedMembers.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 6 }}>
                {stagedMembers.map((p) => (
                  <span key={p.loginName} style={s.pickedChip}>
                    {p.displayName}
                    <button style={s.chipX} disabled={busy} title="Remove" onClick={() => unstageMember(p.loginName)}>✕</button>
                  </span>
                ))}
              </div>
            )}
            <div style={s.ddwrap}>
              <input
                style={s.input}
                placeholder="Search people in the tenant to add…"
                value={peopleQuery}
                disabled={busy}
                onChange={(e) => setPeopleQuery(e.target.value)}
              />
              {peopleQuery.trim().length >= 2 && (
                <div style={s.dd}>
                  {peopleSearching && <div style={s.ddItem}>Searching…</div>}
                  {!peopleSearching && peopleResults.map((p) => (
                    <div key={p.loginName} style={s.ddItem} onClick={() => stageMember(p)}>
                      {p.displayName} <span style={s.mono}>{p.email}</span>
                    </div>
                  ))}
                  {!peopleSearching && peopleResults.length === 0 && (
                    <div style={{ ...s.ddItem, color: "#666" }}>No matching people.</div>
                  )}
                </div>
              )}
            </div>

            {createErrors.length > 0 && (
              <div style={s.missing}>Before creating: {createErrors.join(" ")}</div>
            )}
            {/* The reason, restated beside the button. The scope warning already appears up
                in the persona panel, but that is far enough away that a disabled button
                reads as "the tool is broken" rather than "you picked the wrong tier" —
                which is exactly how it got reported. */}
            {scopeMismatch && (
              <div style={s.missing}>Before creating: {scopeMismatch}</div>
            )}
            <div style={{ display: "flex", gap: 6, marginTop: 10 }}>
              <button
                // The disabled STYLE must track the disabled STATE. Styling on createErrors
                // alone left this looking clickable while scopeMismatch silently blocked it,
                // and a button that appears to do nothing is worse than the mis-scoped
                // mapping it was added to prevent.
                style={createErrors.length === 0 && !busy && !scopeMismatch ? s.addBtn : s.addBtnOff}
                disabled={createErrors.length > 0 || busy || !!scopeMismatch}
                onClick={() => { onCreateAndMap().catch(() => undefined); }}
              >
                Create group &amp; add mapping
              </button>
              <button style={s.secondaryBtn} disabled={busy} onClick={cancelCreate}>Back to search</button>
            </div>
            <div style={s.hint}>
              Creates the native SharePoint group, its mapping, and any members above — all in
              one step. It gets no permissions until you run Folder Reconciliation.
            </div>
          </div>
        ) : (
          /* Search mode — search box + results dropdown. "Create a new group" is
             an action inside the dropdown that switches to create mode. */
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
                  <div style={{ ...s.ddItem, color: "#666" }}>No matching group.</div>
                )}
                {!searching && canManage === true && (
                  <div style={{ ...s.ddItem, ...s.ddCreate }} onClick={startCreate}>
                    ➕ Create a new group{query.trim() ? ` “${query.trim()}”` : ""}…
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* Existing-group flow: for a group picked via search, choose role/segment/
            tier and Add mapping. Hidden during create — the panel above owns that. */}
        {!creating && (
          <>
            {selectionFields}

            {group && role && draftErrors.length > 0 && (
              <div style={s.missing}>Before adding: {draftErrors.join(" ")}</div>
            )}

            {/* Scope mismatch BLOCKS, it does not merely warn. A Head of Department
                mapped at the unit tier is the worst kind of wrong: the row is valid,
                reconciliation grants it happily, and the head silently ends up seeing
                one unit instead of the whole department. Nothing downstream can detect
                that — the row looks exactly like a legitimate unit mapping — so this
                is the only place it can be caught. */}
            <button
              style={canAdd && !scopeMismatch ? s.addBtn : s.addBtnOff}
              disabled={!canAdd || !!scopeMismatch}
              onClick={() => { onAdd().catch(() => undefined); }}
            >
              Add mapping
            </button>
          </>
        )}
      </div>

      {/* Two views of the same data. "Mappings" answers "which group is on which
          folder"; "People" answers "who is in those groups", which is the question
          the client actually asks and could previously only reach one group at a
          time through the Members modal. */}
      {modesUnreadable && (
        <div style={{ fontSize: 12, color: "#b45309", background: "#fff8e1", border: "1px solid #f0c000", borderRadius: 4, padding: "8px 12px", marginBottom: 12 }}>
          Could not read any <strong>mode</strong> rows from <strong>{CONFIG_LIST()}</strong>, so the
          Segment column below shows raw term-set GUIDs and the segment picker is empty. The
          mappings themselves are fine — this is a config read, not your data.
        </div>
      )}

      <div style={{ display: "flex", gap: 6, marginBottom: 12, borderBottom: "1px solid #e1e1e1" }}>
        {([
          { key: "mappings" as const, label: `Existing mappings (${existing.length})` },
          { key: "people" as const, label: `People (${distinctGroups.length} groups)` },
        ]).map((t) => (
          <button
            key={t.key}
            onClick={() => setAccessView(t.key)}
            style={{
              padding: "8px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer",
              fontFamily: "inherit", background: "none", border: "none",
              borderBottom: `2px solid ${accessView === t.key ? "#0f6c3f" : "transparent"}`,
              color: accessView === t.key ? "#0f6c3f" : "#666",
              marginBottom: -1,
            }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {accessView === "people" ? renderPeopleView() : (
      <>
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
      <table style={s.table}>
        <thead>
          <tr>
            <th style={{ ...s.th, width: 28 }}>
              <input type="checkbox" checked={allSelected} disabled={busy || existing.length === 0} onChange={toggleAll} title="Select all" />
            </th>
            <th style={s.th}>Group</th>
            <th style={s.th}>Segment</th>
            <th style={s.th}>Tier</th>
            <th style={s.th}>Role</th>
            <th style={s.th} />
          </tr>
        </thead>
        <tbody>
          {existing.length === 0 && (
            <tr><td style={s.td} colSpan={6}>No mappings yet.</td></tr>
          )}
          {existing.map((r) => (
            <tr key={r.itemId}>
              <td style={s.td}>
                <input type="checkbox" checked={selected.has(r.itemId)} disabled={busy} onChange={() => toggleSel(r.itemId)} />
              </td>
              <td style={s.td}>
                {r.GroupName || <span style={s.mono}>{r.GroupId}</span>}
                {isStaleRow(r) && (
                  <span style={s.staleBadge} title="This mapping points at a term set/term that no longer exists (likely recreated with a new GUID). Delete it and recreate the mapping, then re-run Folder Reconciliation.">stale</span>
                )}
              </td>
              <td style={s.td}>{r.Segment ? segmentLabelFor(r.Segment) : "—"}</td>
              <td style={s.td}>
                {!r.UnitTermGuid
                  ? "—"
                  : r.UnitTermGuid === r.Segment
                    ? "(segment level)"
                    : tierLabels[r.UnitTermGuid] ?? <span style={s.mono}>{r.UnitTermGuid}</span>}
              </td>
              <td style={s.td}>{r.Role}</td>
              <td style={s.td}>
                {/* Confirmation moved OUT of the row and into a modal (below). Since
                    2026-08-04 deleting the last row for a group also deletes the
                    SharePoint group and its membership, and site groups are NOT
                    recoverable from the recycle bin. A two-button inline confirm in a
                    table cell gave that no more weight than removing a mapping. */}
                <span style={{ display: "inline-flex", gap: 6 }}>
                  {canManage === true && (
                    <button style={s.ghost} disabled={busy} onClick={() => openMemberModal(r)} title="Add or remove members of this group">Members</button>
                  )}
                  <button style={s.delBtn} disabled={busy} onClick={() => setConfirmDel(r.itemId)}>Delete</button>
                </span>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      </>
      )}

      {/* Delete-mapping confirmation — a modal, not an inline Yes/Cancel in the table cell.
          Since 2026-08-04 removing a group's LAST mapping also deletes the SharePoint group
          and its membership, and a deleted site group is NOT recoverable from the recycle
          bin. The wording therefore has to differ between the two cases: an admin who
          thinks they are tidying a mapping must not silently destroy a group and everyone's
          access with it, and an admin who IS clearing out a group should be told the access
          goes too — that is the whole reason the behaviour exists. */}
      {confirmDel !== undefined && (() => {
        const row = existing.find((r) => r.itemId === confirmDel);
        const gid = Number(row?.GroupId ?? NaN);
        const others = existing.filter((r) => Number(r.GroupId) === gid && r.itemId !== confirmDel);
        const lastRow = !!row && !isNaN(gid) && gid > 0 && others.length === 0;
        return (
          <div style={s.modalOverlay} onClick={() => setConfirmDel(undefined)}>
            <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
              <div style={s.modalHead}>
                <span>{lastRow ? "Delete mapping and group?" : "Delete mapping?"}</span>
                <button style={s.chipX} onClick={() => setConfirmDel(undefined)} title="Close">✕</button>
              </div>
              <div style={s.modalBody}>
                {lastRow ? (
                  <div style={s.dangerBox}>
                    <p style={{ margin: "0 0 8px" }}>
                      ⚠ This is the <strong>last mapping</strong> for{" "}
                      <strong>{row?.GroupName}</strong>, so the SharePoint group will be
                      deleted along with it.
                    </p>
                    <ul style={{ margin: "0 0 8px 18px", padding: 0 }}>
                      <li>The group&apos;s <strong>folder access is removed</strong> — that is the point: a group left behind keeps every permission it already holds.</li>
                      <li>Its <strong>membership list is lost</strong>. The user accounts are not touched, but who was in the group is not recoverable.</li>
                      <li>A deleted SharePoint group is <strong>not in the recycle bin</strong>. This cannot be undone.</li>
                    </ul>
                    <p style={{ margin: 0 }}>
                      Keep the group? Add another mapping for it first, then delete this row.
                    </p>
                  </div>
                ) : (
                  <div>
                    <p style={{ margin: "0 0 8px" }}>
                      Remove the <strong>{row?.Role}</strong> mapping for{" "}
                      <strong>{row?.GroupName}</strong>?
                    </p>
                    <p style={{ margin: "0 0 8px" }}>
                      The group is <strong>kept</strong> — it still has {others.length} other
                      mapping{others.length === 1 ? "" : "s"}.
                    </p>
                    {/* Stated because it is the one thing that surprises people: removing a
                        row does not remove the permission it created. Nothing in this tool
                        revokes a role assignment, so the group keeps this grant on the
                        folder until someone strips it by hand — and the list will no longer
                        show that it exists. */}
                    <p style={{ margin: 0, color: "#8a6d00" }}>
                      Note: the folder permission this row created stays in place until it is
                      removed by hand in SharePoint. Deleting the row stops it being
                      re-applied; it does not take the access away.
                    </p>
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", padding: 12 }}>
                <button style={s.ghost} disabled={busy} onClick={() => setConfirmDel(undefined)}>Cancel</button>
                <button style={s.delBtn} disabled={busy} onClick={() => { onDelete(confirmDel).catch(() => undefined); }}>
                  {lastRow ? "Delete mapping and group" : "Delete mapping"}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Member-management modal — separate from the add-mapping form. */}
      {memberModal && (
        <div style={s.modalOverlay} onClick={closeMemberModal}>
          <div style={s.modalBox} onClick={(e) => e.stopPropagation()}>
            <div style={s.modalHead}>
              <span>Members — {memberModal.displayName}</span>
              <button style={s.chipX} onClick={closeMemberModal} title="Close">✕</button>
            </div>
            <div style={s.modalBody}>
              {mmMembers === undefined && <div style={{ fontSize: 12, color: "#666" }}>Loading members…</div>}
              {mmMembers !== undefined && mmMembers.length === 0 && (
                <div style={{ fontSize: 12, color: "#888", fontStyle: "italic" }}>No members yet.</div>
              )}
              {(mmMembers ?? []).map((m) => (
                <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", borderBottom: "1px solid #f4f4f4" }}>
                  <span style={{ flex: 1 }}>{m.title}</span>
                  <span style={s.mono}>{m.email}</span>
                  {canManage === true && (mmConfirmRemove === m.id ? (
                    <span style={{ display: "inline-flex", gap: 6 }}>
                      <button style={s.delBtn} disabled={mmBusy} onClick={() => { mmRemove(m.id).catch(() => undefined); }}>Remove</button>
                      <button style={s.ghost} disabled={mmBusy} onClick={() => setMmConfirmRemove(undefined)}>Cancel</button>
                    </span>
                  ) : (
                    <button style={s.chipX} disabled={mmBusy} title="Remove from group" onClick={() => setMmConfirmRemove(m.id)}>✕</button>
                  ))}
                </div>
              ))}
              {canManage === true && (
                <div style={{ ...s.ddwrap, marginTop: 12 }}>
                  <input
                    style={s.input}
                    placeholder="Search people in the tenant to add…"
                    value={mmQuery}
                    disabled={mmBusy}
                    onChange={(e) => setMmQuery(e.target.value)}
                  />
                  {mmQuery.trim().length >= 2 && (
                    <div style={s.dd}>
                      {mmSearching && <div style={s.ddItem}>Searching…</div>}
                      {!mmSearching && mmResults.map((p) => (
                        <div key={p.loginName} style={s.ddItem} onClick={() => { mmAdd(p).catch(() => undefined); }}>
                          {p.displayName} <span style={s.mono}>{p.email}</span>
                        </div>
                      ))}
                      {!mmSearching && mmResults.length === 0 && (
                        <div style={{ ...s.ddItem, color: "#666" }}>No matching people.</div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>
            <div style={s.modalFoot}>
              <button style={s.secondaryBtn} onClick={closeMemberModal}>Done</button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>{toast.message}</div>
      )}
    </div>
  );
}
