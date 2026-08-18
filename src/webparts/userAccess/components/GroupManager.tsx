// Group Management — the whole lifecycle of a SharePoint group, and nothing about folders.
//
// Spec: docs/superpowers/specs/2026-08-14-group-management-separation-design.md
//
// WHY THIS PAGE EXISTS. Creating a group used to be reachable only as half of Folder Access's
// "Create group & add mapping" — one indivisible action that rolled the group back if any mapping
// row failed, and whose mirror deleted the group when its last mapping row went. The client called
// it confusing, and the confusion was a MISSING STATE: the system could not express "this group
// exists but is not assigned to a folder yet", which is the state an admin actually works in.
//
// So this page owns groups and the people in them. Folder Access owns which folder a group can
// reach. That sentence is the whole mental model, and it only works because the split is complete.
//
// Creating a group here grants NOTHING anywhere: createSiteGroup assigns no permission level, and
// folder ACLs come from Folder Reconciliation reading the Group Map. That is what makes "create
// now, assign later" safe rather than a window in which something is half-granted.
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { parseLevels } from "../../../shared/formModel";
import { abbrevListTitle } from "../../../shared/folderAbbreviation";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  GroupMapRole,
  GroupMapWriteRow,
  PERSONAS,
  buildGroupMapRow,
  namingRoleFor,
  normalizeRoleValue,
  roleLabel,
  suggestGroupName,
  validateGroupName,
} from "../../../shared/groupMapModel";
import { addMemberWithSiteEntry } from "../../../shared/siteEntryGroup";
import {
  searchSiteGroups,
  createSiteGroup,
  deleteSiteGroup,
  searchTenantPeople,
  DUPLICATE_GROUP,
  SpGroup,
  PersonPick,
} from "../../../shared/spGroups";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";

type Props = {
  context: WebPartContext;
  siteUrl: string;
  /**
   * Hide the single-group create form, leaving the group LIST and its members.
   *
   * Set by the guided flow, which offers one creation control with a mode switch instead of stacking two
   * cards that do the same job (client, 2026-08-18). Absent means shown, so the standalone page is
   * unaffected. The list is never hidden — it is how an admin sees what already exists.
   */
  hideCreateForm?: boolean;
  /**
   * Bumped by a host when something outside this component has written Group Map rows — a bulk run.
   *
   * The mapping badges are read once at mount, so after a run wrote 790 rows every group still read
   * `not mapped`, which is the reading that makes an admin press Run a second time. Absent means
   * never refreshed, so the standalone page is unaffected unless it opts in.
   */
  refreshKey?: number;
};

/** Resolved per site — the client renames these to "CRS …" at import. */
const GROUP_MAP_LIST = (): string => cachedListTitle(LIST_SUFFIX.groupMap);
const CONFIG_LIST = (): string => cachedListTitle(LIST_SUFFIX.config);

/**
 * A segment as this page needs it.
 *
 * `code` is the StagingFolder and `levelNames` come from the mode row's `Levels` chain. Both were added
 * 2026-08-17 for the client: the suggested name must use FOLDER CODES (`GHO_GF_COR`), and the cascade
 * must be labelled with the segment's own tier names rather than "Level 1" / "Level 2".
 */
type ModePick = { label: string; termSetGuid: string; code: string; levelNames: string[] };
type TermLite = { id: string; label: string };
/** A Group Map row, only as much of it as the delete dialog has to describe. */
type MapRow = { itemId: number; groupId: string; role: GroupMapRole; segment: string; tier: string };

const s: Record<string, React.CSSProperties> = {
  card:      { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  head:      { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  label:     { display: "block", fontSize: 12, fontWeight: 600, color: "#333", margin: "12px 0 4px" },
  input:     { width: "100%", maxWidth: 460, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  select:    { width: "100%", maxWidth: 460, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  ddwrap:    { position: "relative", maxWidth: 460 },
  dd:        { position: "absolute", top: "100%", left: 0, right: 0, zIndex: 20, background: "#fff", border: "1px solid #c7c7c7", borderRadius: 4, maxHeight: 240, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,.12)" },
  ddItem:    { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f0f0f0", fontSize: 13 },
  btn:       { padding: "6px 16px", fontSize: 13, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  btnOff:    { padding: "6px 16px", fontSize: 13, background: "#c7c7c7", color: "#fff", border: "none", borderRadius: 4, cursor: "not-allowed" },
  ghost:     { padding: "5px 12px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  danger:    { padding: "5px 12px", fontSize: 12, color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, background: "#fff", cursor: "pointer" },
  chip:      { display: "inline-flex", alignItems: "center", gap: 6, padding: "3px 8px", margin: "0 6px 6px 0", fontSize: 12, background: "#eef4ff", border: "1px solid #cfe0ff", borderRadius: 12 },
  row:       { display: "flex", alignItems: "center", gap: 10, padding: "8px 10px", borderBottom: "1px solid #f0f0f0" },
  groupName: { fontWeight: 600, fontSize: 13, flex: 1, textAlign: "left", background: "none", border: "none", cursor: "pointer", padding: 0, color: "#1b1b1b" },
  badge:     { fontSize: 11, padding: "2px 8px", borderRadius: 10, background: "#fff4ce", border: "1px solid #f2d98c", color: "#7a5c00" },
  mapped:    { fontSize: 11, padding: "2px 8px", borderRadius: 10, background: "#e7f4ec", border: "1px solid #b7dcc4", color: "#0f6c3f" },
  err:       { fontSize: 12, color: "#a4262c", margin: "6px 0 0" },
  hint:      { fontSize: 11, color: "#666", marginTop: 6, lineHeight: 1.45 },
  warnBox:   { padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5, marginBottom: 12 },
  okBox:     { padding: "10px 12px", border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 4, fontSize: 12, color: "#0f6c3f", lineHeight: 1.5, marginBottom: 12 },
  modalBg:   { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" },
  modal:     { background: "#fff", borderRadius: 6, padding: 20, width: "min(560px, 92vw)", maxHeight: "84vh", overflowY: "auto", boxShadow: "0 8px 32px rgba(0,0,0,.25)" },
  toast:     { position: "fixed", bottom: 20, right: 20, padding: "10px 16px", borderRadius: 4, fontSize: 13, color: "#fff", zIndex: 200, maxWidth: 460, lineHeight: 1.45 },
};

/**
 * The tier a persona's rows sit on, or "unit" for an unknown key.
 *
 * Defaults to the NARROWEST scope on purpose: an unknown persona naming a group after a unit describes
 * less reach than it has, which is the harmless direction. Defaulting to "segment" would name a
 * unit-scoped group as though it covered everything.
 */
function personaScope(key: string): "segment" | "department" | "unit" {
  const p = PERSONAS.filter((x) => x.key === key)[0];
  return p ? p.scope : "unit";
}

/** Personas grouped by family, in PERSONAS order, for an <optgroup> list. */
function personaFamilies(): Array<{ family: string; items: typeof PERSONAS }> {
  const out: Array<{ family: string; items: typeof PERSONAS }> = [];
  for (const p of PERSONAS) {
    const last = out[out.length - 1];
    if (last && last.family === p.family) last.items.push(p);
    else out.push({ family: p.family, items: [p] });
  }
  return out;
}

export default function GroupManager({ context, siteUrl, hideCreateForm, refreshKey }: Props): React.ReactElement {
  const [canManage, setCanManage] = useState<boolean | undefined>(undefined);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [groups, setGroups] = useState<SpGroup[]>([]);
  const [mapRows, setMapRows] = useState<MapRow[] | undefined>(undefined);
  const [filter, setFilter] = useState("");
  const [busy, setBusy] = useState(false);
  const [toast, setToast] = useState<{ message: string; error: boolean } | undefined>(undefined);

  // Create form
  const [newName, setNewName] = useState("");
  const [nameTouched, setNameTouched] = useState(false);
  const [modes, setModes] = useState<ModePick[]>([]);
  const [mode, setMode] = useState<ModePick | undefined>(undefined);
  const [levels, setLevels] = useState<TermLite[][]>([]);
  /**
   * Folder code by term GUID, from the CRS Term Abbreviation list.
   *
   * The suggested name is built from CODES, not labels (client, 2026-08-17: *"client would want the group
   * name to be GHO_GF_COR"*). That is not only shorter — it is the only form that MATCHES ANYTHING.
   * Reconciliation names every folder from these codes, and FolderAdmin's `groupsExist` check looks for
   * groups whose name starts with the segment's StagingFolder — so a group called
   * `Group Head Office_Group Finance_…` matched neither its own folder nor that check.
   *
   * Read once for the whole list rather than per segment: 167 rows on this tenant, one request, and the
   * builder is used repeatedly in a sitting.
   */
  const [codes, setCodes] = useState<Record<string, string>>({});
  const [chosen, setChosen] = useState<TermLite[]>([]);
  /**
   * The chosen PERSONA key, not a raw role (spec 2026-08-18 §3).
   *
   * Folder Access dropped its role chips so nobody hand-assembles a wrong combination; this screen kept
   * them, which was the same trap in a second place — nobody should have to choose between "Delete
   * pending files" and "View only — whole department". The persona also decides the NAME, via its
   * declared `namingRole`.
   */
  const [builderPersona, setBuilderPersona] = useState<string>("");
  /**
   * Free-text naming, off by default (client, 2026-08-18: *"client doesn't care about naming convention
   * and they will simple type it in for no reason"*).
   *
   * It cannot be removed outright: `CRS_SITE_MEMBERS` follows no convention, and there will always be a
   * one-off. But the admin has to opt in, which is the opposite of the old default.
   */
  const [advanced, setAdvanced] = useState(false);
  const [staged, setStaged] = useState<PersonPick[]>([]);
  const [peopleQuery, setPeopleQuery] = useState("");
  const [peopleResults, setPeopleResults] = useState<PersonPick[]>([]);

  // Members

  // Delete
  const [deleting, setDeleting] = useState<SpGroup | undefined>(undefined);

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(undefined), 7000);
  };

  /* ── Data access ───────────────────────────────────────────────────────── */

  const loadGroups = async (): Promise<SpGroup[]> =>
    // Empty query = every group an admin may manage, built-ins already excluded by id.
    searchSiteGroups(context.spHttpClient, siteUrl, "");

  /**
   * Group Map rows, for the "not mapped" badge and the delete dialog.
   *
   * `undefined` means UNREADABLE and is not the same as an empty array. A failed read must not
   * badge every group "not mapped" — that reads as a data-loss event — nor let the delete dialog
   * claim a group holds no mappings when it may hold several.
   */
  const loadMapRows = async (): Promise<MapRow[] | undefined> => {
    const base = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items`;
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${base}?$select=Id,GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return undefined;
      const data = await res.json();
      return ((data.value ?? []) as Array<{ Id: number; GroupId?: string; Segment?: string; UnitTermGuid?: string; Role?: string }>)
        .map((r) => ({
          itemId: r.Id,
          groupId: (r.GroupId ?? "").trim(),
          role: normalizeRoleValue(r.Role ?? "") as GroupMapRole,
          segment: (r.Segment ?? "").trim(),
          tier: (r.UnitTermGuid ?? "").trim(),
        }));
    } catch {
      return undefined;
    }
  };

  /**
   * Folder codes by term GUID.
   *
   * Fails to `{}` rather than throwing: with no codes the builder falls back to term labels, which still
   * produces a usable suggestion. Blocking group creation because an unrelated list could not be read
   * would be the wrong trade — nothing on this screen writes a permission, and the admin may type any
   * name they like.
   */
  const loadCodes = async (): Promise<Record<string, string>> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(abbrevListTitle())}')/items` +
          `?$select=TermGuid,Abbreviation&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return {};
      const data = await res.json();
      const out: Record<string, string> = {};
      for (const r of (data.value ?? []) as Array<{ TermGuid?: string; Abbreviation?: string }>) {
        const key = (r.TermGuid ?? "").trim().toLowerCase();
        const code = (r.Abbreviation ?? "").trim();
        // A row with a BLANK code is the same as no row: that term gets no folder, so there is no code to
        // put in a name. Skipped rather than stored as "", which would beat the label fallback and yield
        // a name with an empty segment in it.
        if (key && code) out[key] = code;
      }
      return out;
    } catch {
      return {};
    }
  };

  /** One Group Map row. Same endpoint, headers and Title fallback as Folder Access — one shape, not two. */
  const postRow = async (row: GroupMapWriteRow): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
        },
        // Title is mandatory on a default SP list; set it so the create never 400s.
        body: JSON.stringify({ Title: row.GroupName || row.GroupId, ...row }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`);
    }
  };

  /**
   * The Group Map rows a newly created group needs — ONE PER ROLE in its persona.
   *
   * Spec 2026-08-18 §4. The dropdowns already collected segment, tier and persona to build the name, so
   * writing the rows here costs nothing extra and takes the PARSE off the critical path: the role is
   * recorded from the persona that was chosen, never recovered from the group's name.
   *
   * Returns [] when there is nothing safe to write — no persona, no segment, or a free-typed name with no
   * cascade behind it. A group with no rows is inert and mappable later on Folder Access; a GUESSED row is
   * a grant nobody chose.
   */
  const rowsForNewGroup = (groupId: string, groupName: string): GroupMapWriteRow[] => {
    const p = PERSONAS.filter((x) => x.key === builderPersona)[0];
    if (!p || advanced || !mode) return [];
    const scope = personaScope(p.key);
    // A SEGMENT-scope row carries UnitTermGuid === the term-set guid; that is how Folder Access records
    // "the segment IS the tier", and reconciliation reads it the same way.
    const tierGuid =
      scope === "segment"
        ? mode.termSetGuid
        : scope === "department"
          ? chosen[0]?.id ?? ""
          : chosen[chosen.length - 1]?.id ?? "";
    if (!tierGuid) return [];
    return p.roles.map((r) =>
      buildGroupMapRow({
        groupId,
        groupName,
        role: r,
        segmentGuid: mode.termSetGuid,
        tierGuid,
        // No `scope` and no `target`: absent scope reads as Folder, and buildGroupMapRow sets Target "" for
        // it. These are FOLDER grants — library and page entry rows are a different screen's job.
      }),
    );
  };

  const loadModes = async (): Promise<ModePick[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG_LIST())}')/items?$select=ModeLabel,TermSetGuid,StagingFolder,Levels&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{
      ModeLabel?: string; TermSetGuid?: string; StagingFolder?: string; Levels?: string;
    }>)
      .filter((r) => (r.TermSetGuid ?? "").trim() && (r.Levels ?? "").trim())
      .map((r) => ({
        label: (r.ModeLabel ?? "").trim() || (r.TermSetGuid ?? "").trim(),
        termSetGuid: (r.TermSetGuid ?? "").trim(),
        code: (r.StagingFolder ?? "").trim(),
        // PERMISSIONED tiers only — those are the levels this cascade walks. A below-Unit tier such as
        // SubUnit carries no group and must not appear as a dropdown here.
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

  // Full Control is needed to create, delete, and change membership. Detected up front so the
  // controls are explained rather than failing with a 403 on click.
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

  const reload = async (): Promise<void> => {
    setLoading(true);
    try {
      await primeNames(context.spHttpClient, siteUrl);
      setGroups(await loadGroups());
      setMapRows(await loadMapRows());
      setLoadError(undefined);
    } catch (e) {
      setLoadError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  /* ── Mount ─────────────────────────────────────────────────────────────── */

  useEffect(() => {
    loadCanManage().then(setCanManage).catch(() => setCanManage(false));
    reload().catch(() => undefined);
    loadModes().then(setModes).catch(() => undefined);
    loadCodes().then(setCodes).catch(() => undefined);
  }, []);

  /**
   * Re-read when the host says something outside this component wrote Group Map rows.
   *
   * The PREVIOUS value lives in a ref rather than the effect firing on every render: a re-mount
   * carrying a non-zero key would otherwise repeat the read for nothing, and on mount the effect
   * above has already done it.
   */
  const lastRefresh = useRef(refreshKey);
  useEffect(() => {
    if (refreshKey === undefined || refreshKey === lastRefresh.current) return;
    lastRefresh.current = refreshKey;
    reload().catch(() => undefined);
  }, [refreshKey]);

  // Debounced people search, the same shape as the other access pages.
  useEffect(() => {
    const q = peopleQuery.trim();
    if (q.length < 2) { setPeopleResults([]); return; }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q).then(setPeopleResults).catch(() => setPeopleResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [peopleQuery]);

  /* ── Audit ─────────────────────────────────────────────────────────────── */

  /**
   * Record a group-lifecycle event. Fire-and-forget: the act has already happened, and no admin
   * action should fail because the log was unreachable.
   */
  const log = (
    event: string,
    summary: string,
    details: string[],
    outcome: "Success" | "Failed" = "Success",
  ): void => {
    writeAudit(context.spHttpClient, siteUrl, {
      event,
      outcome,
      source: "GroupManagement",
      at: new Date(),
      actorName: context.pageContext.user.displayName,
      actorEmail: context.pageContext.user.email,
      summary,
      details,
    }).catch(() => undefined);
  };

  /* ── The name builder ──────────────────────────────────────────────────── */

  /**
   * Writes NOTHING but the name box.
   *
   * The convention (GHO_GF_CORU_UPLOADER) is derived from segment + tier + role, and Folder Access
   * reads the suffix back to pre-select a role. Creating the group before any of that exists would
   * lose it, so the same cascade is offered here purely as a name generator. The admin may ignore
   * it entirely and type any name — it always was a suggestion, never a rule.
   */
  const applyBuilder = (m: ModePick | undefined, path: TermLite[], persona: string): void => {
    if (!m) return;
    // Code per tier, falling back to the LABEL when a term has no code yet. A fallback rather than a
    // refusal because this box only suggests a name — but such a group will not line up with its folder,
    // so the hint below the field says which tiers are missing a code.
    // TRUNCATED BY THE PERSONA'S SCOPE. A Head of Department group is named `GHO_GF_HOD`, not
    // `GHO_GF_TAX_HOD` — its row sits on the DEPARTMENT term and reaches the units beneath by fan-out, so
    // a unit in its name would describe a narrower grant than it has. C-Level is segment-wide and takes no
    // tier at all.
    const scope = personaScope(persona);
    const depth = scope === "segment" ? 0 : scope === "department" ? 1 : path.length;
    const built = suggestGroupName(
      m.code || m.label,
      path.slice(0, depth).map((t) => codes[t.id.toLowerCase()] || t.label),
      namingRoleFor(persona),
    );
    // Never silently overwrite something the admin typed. The name is the one field here they may
    // have composed by hand for a reason.
    if (nameTouched && newName.trim() && newName.trim() !== built) {
      if (!window.confirm(`Replace the name you typed with "${built}"?`)) return;
    }
    setNewName(built);
    setNameTouched(false);
  };

  const pickMode = async (termSetGuid: string): Promise<void> => {
    const m = modes.find((x) => x.termSetGuid === termSetGuid);
    setMode(m);
    setChosen([]);
    setLevels([]);
    if (!m) return;
    const tops = await loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${m.termSetGuid}/children`);
    setLevels(tops.length ? [tops] : []);
    applyBuilder(m, [], builderPersona);
  };

  const pickTerm = async (levelIndex: number, termId: string): Promise<void> => {
    if (!mode) return;
    const options = levels[levelIndex] ?? [];
    const term = options.find((t) => t.id === termId);
    const path = term ? [...chosen.slice(0, levelIndex), term] : chosen.slice(0, levelIndex);
    setChosen(path);
    const nextLevels = levels.slice(0, levelIndex + 1);
    if (term) {
      const kids = await loadTerms(
        `${siteUrl}/_api/v2.1/termStore/sets/${mode.termSetGuid}/terms/${term.id}/children`,
      );
      if (kids.length) nextLevels.push(kids);
    }
    setLevels(nextLevels);
    applyBuilder(mode, path, builderPersona);
  };

  /* ── Create ────────────────────────────────────────────────────────────── */

  const nameErrors = validateGroupName(newName, groups.map((g) => g.title));
  const canCreate = canManage === true && !busy && nameErrors.length === 0;

  const onCreate = async (): Promise<void> => {
    if (!canCreate) return;
    setBusy(true);
    try {
      const made = await createSiteGroup(context.spHttpClient, siteUrl, newName.trim());

      // No rollback, deliberately. Folder Access deleted the group when a mapping row failed,
      // because a group created FOR a mapping that does not exist was garbage. Here the group is
      // the deliverable — a member who fails to add is reported, and the group is kept.
      let added = 0;
      let failed = 0;
      const notes: string[] = [];
      for (const p of staged) {
        try {
          const r = await addMemberWithSiteEntry(
            context.spHttpClient, siteUrl, { id: made.id, title: made.title }, p.loginName,
          );
          added++;
          if (r.note) notes.push(`${p.displayName}: ${r.note}`);
        } catch (e) {
          failed++;
          notes.push(`${p.displayName}: could not be added — ${(e as Error).message}`);
        }
      }

      /* THE MAPPING ROWS — spec 2026-08-18 §4.
         Written AFTER the group exists, because a row needs its integer GroupId.

         The group is NOT rolled back on a row failure. Deleting it would destroy a group that may already
         have the members added above, and re-running is safe because the group is then found by name. The
         2026-08-14 design had the opposite rule and that is exactly what made "created but not assigned"
         inexpressible. So a half-done write is REPORTED as half-done rather than flattened either way:
         claiming success would leave a group granting nothing, and claiming failure would send someone
         hunting for a group that exists. */
      // String(): the SP group id is an INTEGER and the Group Map column is Text. Grants key on this
      // value, so it has to match what Folder Access writes — a number here would serialise the same
      // but the type would drift the moment anything compared them.
      const wanted = rowsForNewGroup(String(made.id), made.title);
      let mapped = 0;
      for (const row of wanted) {
        try {
          await postRow(row);
          mapped++;
        } catch (e) {
          notes.push(`mapping ${row.Role}: ${(e as Error).message}`);
        }
      }

      // Captured before the reset below, which is about to clear all of it.
      const createdName = made.title;
      const mapNote =
        wanted.length === 0
          ? " It grants nothing yet — assign it on the Folder Access page."
          : mapped === wanted.length
            ? ` Mapped at ${personaScope(builderPersona)} level with ${mapped} role${mapped === 1 ? "" : "s"} — run Folder Reconciliation to apply it.`
            : ` ⚠ Only ${mapped} of ${wanted.length} mappings were written — finish it on Folder Access before reconciling.`;
      setNewName("");
      setNameTouched(false);
      setStaged([]);
      setPeopleQuery("");
      setPeopleResults([]);
      setMode(undefined);
      setChosen([]);
      setLevels([]);
      setBuilderPersona("");
      // `advanced` is deliberately NOT reset: someone creating one-off groups is usually creating several,
      // and silently snapping the name field back to read-only between them would read as the page
      // fighting them.

      await reload();

      showToast(
        `"${createdName}" created${added ? ` with ${added} member${added === 1 ? "" : "s"}` : ""}.` +
        mapNote +
        (notes.length ? ` ${notes.join(" ")}` : ""),
        // Amber/red whenever anything is outstanding, INCLUDING a partial mapping — a green toast over a
        // half-written mapping is the one outcome nobody would go back and check.
        failed > 0 || notes.length > 0 || mapped < wanted.length,
      );
      log(
        EVENT.groupCreated,
        `Group created — ${createdName}`,
        [
          `Group: ${createdName}`,
          `Members added: ${added}${failed ? `, failed: ${failed}` : ""}`,
          // In the record, not only on screen: a group with no mapping grants nothing, and a
          // reader months later needs to know that creating it was not provisioning it.
          "The group holds no permission level. Folder access needs a Group Map row plus a Folder Reconciliation run.",
          ...notes,
        ],
        failed > 0 ? "Failed" : "Success",
      );
    } catch (e) {
      const msg = (e as Error).message;
      showToast(
        msg === DUPLICATE_GROUP
          ? `A group called "${newName.trim()}" already exists on this site.`
          : `Could not create the group: ${msg}`,
        true,
      );
    } finally {
      setBusy(false);
    }
  };

  /* ── Members ───────────────────────────────────────────────────────────── */


  /* ── Delete ────────────────────────────────────────────────────────────── */

  /** Rows this group holds. `undefined` = the Group Map could not be read. */
  const rowsFor = (g: SpGroup): MapRow[] | undefined =>
    mapRows === undefined ? undefined : mapRows.filter((r) => r.groupId === String(g.id));

  const onDeleteGroup = async (g: SpGroup): Promise<void> => {
    setBusy(true);
    try {
      const rows = rowsFor(g) ?? [];
      // Rows first: a deleted group whose rows survive leaves reconciliation trying to grant a
      // principal that no longer exists, on every run.
      let rowsRemoved = 0;
      for (const r of rows) {
        const res: SPHttpClientResponse = await context.spHttpClient.post(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items(${r.itemId})`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata", "IF-MATCH": "*", "X-HTTP-Method": "DELETE" } },
        );
        if (res.ok) rowsRemoved++;
      }
      await deleteSiteGroup(context.spHttpClient, siteUrl, g.id);
      setDeleting(undefined);
      await reload();
      showToast(
        `"${g.title}" deleted.` +
        (rowsRemoved
          ? ` ${rowsRemoved} mapping${rowsRemoved === 1 ? "" : "s"} removed — run Folder Reconciliation to take the folder permissions away.`
          : ""),
        false,
      );
      log(
        EVENT.groupDeleted,
        `Group deleted — ${g.title}`,
        [
          `Group: ${g.title} (id ${g.id})`,
          `Group Map rows removed: ${rowsRemoved}${rows.length !== rowsRemoved ? ` of ${rows.length} — the rest FAILED` : ""}`,
          // The distinction that makes this row worth reading. Deleting the mapping is not
          // removing the access: the ACL survives on the folder until reconciliation runs.
          rowsRemoved > 0
            ? "FOLDER PERMISSIONS ARE STILL IN PLACE until Folder Reconciliation runs."
            : "The group held no folder mappings.",
        ],
        rows.length !== rowsRemoved ? "Failed" : "Success",
      );
    } catch (e) {
      showToast(`Could not delete "${g.title}": ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  /* ── Render ────────────────────────────────────────────────────────────── */

  const visible = groups.filter(
    (g) => !filter.trim() || g.title.toLowerCase().indexOf(filter.trim().toLowerCase()) !== -1,
  );

  if (canManage === false) {
    return (
      <p style={s.warnBox}>
        Managing groups needs <strong>Full Control</strong> on this site. Ask a site owner to add
        you, or to make the change for you.
      </p>
    );
  }

  return (
    <div>
      <p style={s.okBox}>
        Creating a group here grants <strong>nothing</strong> — a new group holds no permission
        level anywhere. Once it exists, assign it to a segment and tier on the{" "}
        <strong>Folder Access</strong> page, then run <strong>Folder Reconciliation</strong>.
      </p>

      {/* ── Create ─────────────────────────────────────────────────────── */}
      {!hideCreateForm && (
      <div style={s.card}>
        <p style={s.head}>Create a group</p>

        {/* THE NAME IS AN OUTPUT, NOT AN INPUT (client, 2026-08-18: *"client doesn't care about naming
            convention and they will simple type it in for no reason"*).

            The name is load-bearing: Folder Access recovers a group's role by PARSING ITS SUFFIX, so a
            hand-typed `GHO_GF_TAX_HC_UPLOADER` — suffix in the middle — parses as a PLAIN uploader and
            pre-selects the wrong role, which an admin then accepts because it looks right. Deriving it
            from the persona removes the guess.

            Read-only rather than hidden: the admin must be able to SEE what will be created, and a
            disabled-looking field with the value in it says "this is decided" better than no field. */}
        <label style={s.label} htmlFor="gm-name">Group name</label>
        <input
          id="gm-name"
          style={advanced ? s.input : { ...s.input, background: "#f3f2f1", color: "#323130" }}
          value={newName}
          readOnly={!advanced}
          placeholder={advanced ? "e.g. CRS_SITE_MEMBERS" : "Choose a segment, tier and persona below"}
          onChange={(e) => { setNewName(e.target.value); setNameTouched(true); }}
        />
        {newName.trim() !== "" && nameErrors.length > 0 && (
          <p style={s.err}>{nameErrors.join(" ")}</p>
        )}

        {/* The escape hatch, and it cannot be removed: `CRS_SITE_MEMBERS` follows no convention, and there
            will always be a one-off. But it is opt-in, which is the opposite of the old default. */}
        <label style={{ ...s.hint, display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
          <input
            type="checkbox"
            checked={advanced}
            onChange={(e) => { setAdvanced(e.target.checked); setNameTouched(false); }}
          />
          Type the name myself (advanced) — for one-offs like the site-entry group
        </label>
        <p style={s.hint}>
          Nothing is granted here. Once the group exists, reconciliation applies its folder access.
        </p>

        {/* Always shown — this IS the form now, not an optional helper beside a text box. */}
        {!advanced && (
          <div style={{ paddingLeft: 12, borderLeft: "2px solid #e1e1e1", marginTop: 10 }}>
            <label style={s.label} htmlFor="gm-seg">Segment</label>
            <select
              id="gm-seg"
              style={s.select}
              value={mode?.termSetGuid ?? ""}
              onChange={(e) => { pickMode(e.target.value).catch(() => undefined); }}
            >
              <option value="">— select a segment —</option>
              {modes.map((m) => <option key={m.termSetGuid} value={m.termSetGuid}>{m.label}</option>)}
            </select>

            {levels.map((options, i) => (
              <div key={i}>
                {/* The segment's OWN tier name — Department, Unit, Region, Estate/Mill (client, 2026-08-17:
                    "show Segment, Department, Unit instead of level 1 and level 2"). "Level 2" is a
                    position, not a thing anyone recognises, and this page already knows the names.
                    Falls back to the position when the chain is shorter than the tree, which is the
                    depth mismatch the New segment form refuses — better an unhelpful label than a
                    blank one. */}
                <label style={s.label} htmlFor={`gm-lvl-${i}`}>
                  {(mode ? mode.levelNames[i] : "") || `Level ${i + 1}`}
                </label>
                <select
                  id={`gm-lvl-${i}`}
                  style={s.select}
                  value={chosen[i]?.id ?? ""}
                  onChange={(e) => { pickTerm(i, e.target.value).catch(() => undefined); }}
                >
                  <option value="">— select —</option>
                  {options.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
                </select>
              </div>
            ))}

            {/* PERSONA, not a raw role (client, 2026-08-18). The role list offered things like "Delete
                pending files" and "View only — whole department", which are combinations nobody should
                have to assemble; Folder Access dropped its role chips for this reason and this screen was
                the same trap in a second place. The persona also decides the NAME, through its declared
                `namingRole` — so name and capability can no longer disagree. */}
            <label style={s.label} htmlFor="gm-persona">Persona</label>
            <select
              id="gm-persona"
              style={s.select}
              value={builderPersona}
              onChange={(e) => {
                const k = e.target.value;
                setBuilderPersona(k);
                applyBuilder(mode, chosen, k);
              }}
            >
              <option value="">— select a persona —</option>
              {personaFamilies().map((f) => (
                <optgroup key={f.family} label={f.family}>
                  {f.items.map((p) => (
                    <option key={p.key} value={p.key}>{p.label}</option>
                  ))}
                </optgroup>
              ))}
            </select>
            {builderPersona !== "" && (
              <p style={s.hint}>
                {/* The roles are SHOWN rather than chosen. An admin should be able to see what a persona
                    carries — that is how a wrong pick gets caught before 60 groups exist — without being
                    able to assemble one by hand. */}
                Grants:{" "}
                <strong>
                  {(PERSONAS.filter((x) => x.key === builderPersona)[0]?.roles ?? [])
                    .map((r) => roleLabel(r))
                    .join(" · ")}
                </strong>
                . Mapped at <strong>{personaScope(builderPersona)}</strong> level.
              </p>
            )}
            {/* Says WHY a name came out with a full term label in it. Without this the fallback is
                invisible: the box just reads `GHO_GF_Compliance & Operational Risk` and looks like a bug
                rather than a missing code. It matters beyond tidiness — reconciliation names the folder
                from the code, so a group named off a label lines up with nothing. */}
            {chosen.filter((t) => !codes[t.id.toLowerCase()]).length > 0 && (
              <p style={{ ...s.hint, color: "#7a4f00" }}>
                No folder code yet for{" "}
                <strong>
                  {chosen.filter((t) => !codes[t.id.toLowerCase()]).map((t) => t.label).join(", ")}
                </strong>
                , so the name uses the full term name instead. Give it a code on{" "}
                <strong>CRS Term Abbreviations</strong> first — reconciliation names the folder from that
                code, and a group named after the term will not line up with it.
              </p>
            )}
          </div>
        )}

        <label style={s.label} htmlFor="gm-people">Members (optional)</label>
        <div style={s.ddwrap}>
          <input
            id="gm-people"
            style={s.input}
            value={peopleQuery}
            placeholder="Search for a person…"
            onChange={(e) => setPeopleQuery(e.target.value)}
          />
          {peopleResults.length > 0 && (
            <div style={s.dd}>
              {peopleResults.map((p) => (
                <div
                  key={p.loginName}
                  style={s.ddItem}
                  onClick={() => {
                    if (!staged.some((x) => x.loginName === p.loginName)) setStaged([...staged, p]);
                    setPeopleQuery("");
                    setPeopleResults([]);
                  }}
                >
                  {p.displayName} <span style={{ color: "#666" }}>{p.email}</span>
                </div>
              ))}
            </div>
          )}
        </div>
        {staged.length > 0 && (
          <div style={{ marginTop: 8 }}>
            {staged.map((p) => (
              <span key={p.loginName} style={s.chip}>
                {p.displayName}
                <button
                  type="button"
                  style={{ border: "none", background: "none", cursor: "pointer", color: "#a4262c" }}
                  onClick={() => setStaged(staged.filter((x) => x.loginName !== p.loginName))}
                >
                  ✕
                </button>
              </span>
            ))}
          </div>
        )}

        <div style={{ marginTop: 14 }}>
          <button
            type="button"
            style={canCreate ? s.btn : s.btnOff}
            disabled={!canCreate}
            onClick={() => { onCreate().catch(() => undefined); }}
          >
            {busy ? "Working…" : "Create group"}
          </button>
        </div>
      </div>
      )}

      {/* ── The groups ─────────────────────────────────────────────────── */}
      <div style={s.card}>
        <p style={s.head}>Groups on this site ({groups.length})</p>
        <input
          style={s.input}
          value={filter}
          placeholder="Filter by name…"
          onChange={(e) => setFilter(e.target.value)}
        />

        {loadError !== undefined && (
          <p style={s.err}>Could not read this site&apos;s groups — {loadError}</p>
        )}
        {mapRows === undefined && !loading && loadError === undefined && (
          // Unreadable ≠ empty. Without this the badges below would silently claim every group is
          // unmapped, which reads as data loss.
          <p style={s.hint}>
            The {GROUP_MAP_LIST()} list could not be read, so this page cannot show which groups are
            mapped. Creating and deleting still work.
          </p>
        )}

        {/* SCROLLS, once the list is long enough to be a page of its own (defect 4, 2026-08-18):
            303 groups pushed everything else off the screen, including the create form above.

            Capped only while every group is COLLAPSED. The member editor inside an expanded group
            carries an absolutely-positioned people picker, and a scroll container would clip its
            results for any group near the bottom — trading a long page for a control that silently
            cannot be used. When one group is open the admin is working inside it, not scanning the
            list, so the cap has nothing to do. */}
        <div
          style={{
            marginTop: 10,
            maxHeight: "60vh",
            overflowY: "auto",
          }}
        >
          {loading && <p style={s.hint}>Loading…</p>}
          {!loading && visible.length === 0 && (
            <p style={s.hint}>
              {groups.length === 0 ? "No groups on this site yet." : "No group matches that filter."}
            </p>
          )}
          {visible.map((g) => {
            const rows = rowsFor(g);
            return (
              <div key={g.id}>
                <div style={s.row}>
                  {/* PLAIN TEXT, not an expander. Membership moved to Folder Access on 2026-08-18 —
                      one editor, one mount point, because two lists of the same people drift and
                      re-merging these pages is what the client called confusing. This page answers
                      "what groups exist"; who is in them is the other page's question. */}
                  <span style={{ ...s.groupName, cursor: "default" }}>{g.title}</span>
                  {rows !== undefined && (
                    rows.length === 0
                      ? <span style={s.badge}>not mapped</span>
                      : <span style={s.mapped}>{rows.length} mapping{rows.length === 1 ? "" : "s"}</span>
                  )}
                  <button type="button" style={s.danger} disabled={busy} onClick={() => setDeleting(g)}>
                    Delete
                  </button>
                </div>

              </div>
            );
          })}
        </div>
      </div>

      {/* ── Delete confirmation ────────────────────────────────────────── */}
      {deleting !== undefined && (
        <div style={s.modalBg} onClick={() => setDeleting(undefined)}>
          <div style={s.modal} onClick={(e) => e.stopPropagation()}>
            <p style={{ ...s.head, fontSize: 15 }}>Delete &quot;{deleting.title}&quot;?</p>

            {(() => {
              const rows = rowsFor(deleting);
              if (rows === undefined) {
                // Unreadable ≠ "no mappings". Deleting is still allowed — the group may genuinely
                // need to go — but the admin is told the list could not be checked.
                return (
                  <p style={s.warnBox}>
                    The {GROUP_MAP_LIST()} list could not be read, so it is not known whether this
                    group holds folder mappings. Any rows it has will be left behind, and
                    reconciliation will report them.
                  </p>
                );
              }
              if (rows.length === 0) {
                return (
                  <p style={{ fontSize: 13 }}>
                    This group holds no folder mappings. Its members lose it immediately.
                  </p>
                );
              }
              return (
                <>
                  <p style={{ fontSize: 13 }}>
                    This group holds <strong>{rows.length}</strong> folder mapping
                    {rows.length === 1 ? "" : "s"}, which will be deleted with it:
                  </p>
                  <ul style={{ fontSize: 12, color: "#444", lineHeight: 1.6 }}>
                    {rows.map((r) => (
                      <li key={r.itemId}>
                        {roleLabel(r.role)} —{" "}
                        <span style={{ fontFamily: "Consolas, monospace", fontSize: 11 }}>
                          {r.tier || r.segment || "(no tier)"}
                        </span>
                      </li>
                    ))}
                  </ul>
                  {/* The sentence this dialog exists for. Removing the mapping is not removing the
                      access — the ACL survives on the folder until reconciliation runs, and an
                      admin who believes otherwise stops looking. */}
                  <p style={s.warnBox}>
                    Deleting this group removes its mappings.{" "}
                    <strong>
                      The folder permissions it was granted stay in place until Folder
                      Reconciliation runs.
                    </strong>
                  </p>
                </>
              );
            })()}

            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button
                type="button"
                style={s.danger}
                disabled={busy}
                onClick={() => { onDeleteGroup(deleting).catch(() => undefined); }}
              >
                {busy ? "Deleting…" : "Delete group"}
              </button>
              <button type="button" style={s.ghost} onClick={() => setDeleting(undefined)}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {toast !== undefined && (
        <div style={{ ...s.toast, background: toast.error ? "#a4262c" : "#0f6c3f" }}>{toast.message}</div>
      )}
    </div>
  );
}
