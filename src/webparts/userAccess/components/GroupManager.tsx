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
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  GroupMapRole,
  SELECTABLE_ROLES,
  normalizeRoleValue,
  roleLabel,
  suggestGroupName,
  validateGroupName,
  siteEntryGroupTitle,
  isSiteEntryGroupTitle,
} from "../../../shared/groupMapModel";
import { addMemberWithSiteEntry } from "../../../shared/siteEntryGroup";
import {
  searchSiteGroups,
  createSiteGroup,
  deleteSiteGroup,
  getGroupMembers,
  removeGroupMember,
  searchTenantPeople,
  DUPLICATE_GROUP,
  SpGroup,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";

type Props = { context: WebPartContext; siteUrl: string };

/** Resolved per site — the client renames these to "CRS …" at import. */
const GROUP_MAP_LIST = (): string => cachedListTitle(LIST_SUFFIX.groupMap);
const CONFIG_LIST = (): string => cachedListTitle(LIST_SUFFIX.config);

type ModePick = { label: string; termSetGuid: string };
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

export default function GroupManager({ context, siteUrl }: Props): React.ReactElement {
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
  const [builderOpen, setBuilderOpen] = useState(false);
  const [modes, setModes] = useState<ModePick[]>([]);
  const [mode, setMode] = useState<ModePick | undefined>(undefined);
  const [levels, setLevels] = useState<TermLite[][]>([]);
  const [chosen, setChosen] = useState<TermLite[]>([]);
  const [builderRole, setBuilderRole] = useState<GroupMapRole | "">("");
  const [staged, setStaged] = useState<PersonPick[]>([]);
  const [peopleQuery, setPeopleQuery] = useState("");
  const [peopleResults, setPeopleResults] = useState<PersonPick[]>([]);

  // Members
  const [openGroup, setOpenGroup] = useState<SpGroup | undefined>(undefined);
  const [members, setMembers] = useState<SpGroupMember[] | undefined>(undefined);
  const [memberQuery, setMemberQuery] = useState("");
  const [memberResults, setMemberResults] = useState<PersonPick[]>([]);
  const [confirmRemove, setConfirmRemove] = useState<number | undefined>(undefined);

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
  }, []);

  // Debounced people search, the same shape as the other access pages.
  useEffect(() => {
    const q = peopleQuery.trim();
    if (q.length < 2) { setPeopleResults([]); return; }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q).then(setPeopleResults).catch(() => setPeopleResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [peopleQuery]);

  useEffect(() => {
    const q = memberQuery.trim();
    if (q.length < 2) { setMemberResults([]); return; }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q).then(setMemberResults).catch(() => setMemberResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [memberQuery]);

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
  const applyBuilder = (m: ModePick | undefined, path: TermLite[], role: GroupMapRole | ""): void => {
    if (!m) return;
    const built = suggestGroupName(m.label, path.map((t) => t.label), role);
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
    applyBuilder(m, [], builderRole);
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
    applyBuilder(mode, path, builderRole);
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

      // Captured before the reset below, which is about to clear all of it.
      const createdName = made.title;
      setNewName("");
      setNameTouched(false);
      setStaged([]);
      setPeopleQuery("");
      setPeopleResults([]);
      setBuilderOpen(false);
      setMode(undefined);
      setChosen([]);
      setLevels([]);
      setBuilderRole("");
      await reload();

      showToast(
        `"${createdName}" created${added ? ` with ${added} member${added === 1 ? "" : "s"}` : ""}.` +
        " It grants nothing yet — assign it on the Folder Access page." +
        (notes.length ? ` ${notes.join(" ")}` : ""),
        failed > 0 || notes.length > 0,
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

  const openMembers = async (g: SpGroup): Promise<void> => {
    if (openGroup && openGroup.id === g.id) { setOpenGroup(undefined); return; }
    setOpenGroup(g);
    setMembers(undefined);
    setMemberQuery("");
    setMemberResults([]);
    setConfirmRemove(undefined);
    try {
      setMembers(await getGroupMembers(context.spHttpClient, siteUrl, g.id));
    } catch {
      setMembers([]);
      showToast("Could not read this group's members.", true);
    }
  };

  const onAddMember = async (p: PersonPick): Promise<void> => {
    if (!openGroup) return;
    setBusy(true);
    try {
      const res = await addMemberWithSiteEntry(
        context.spHttpClient, siteUrl, { id: openGroup.id, title: openGroup.title }, p.loginName,
      );
      setMembers(await getGroupMembers(context.spHttpClient, siteUrl, openGroup.id));
      setMemberQuery("");
      setMemberResults([]);
      // res.note is never dropped: it means they ARE in the group but may not be able to open the
      // site, which neither they nor the admin would discover until they tried.
      showToast(`${p.displayName} added — access is immediate.${res.note ? ` ${res.note}` : ""}`, !!res.note);
      log(
        EVENT.membersChanged,
        `Member added — ${p.displayName} → ${openGroup.title}`,
        [
          `Group: ${openGroup.title}`,
          `Added: ${p.displayName}${p.email ? ` <${p.email}>` : ""}`,
          isSiteEntryGroupTitle(openGroup.title)
            ? "This IS the site-entry group."
            : res.note || `Also added to ${siteEntryGroupTitle()}.`,
        ],
        res.note ? "Failed" : "Success",
      );
    } catch (e) {
      showToast(`Could not add ${p.displayName}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  const onRemoveMember = async (m: SpGroupMember): Promise<void> => {
    if (!openGroup) return;
    setBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, openGroup.id, m.id);
      setMembers(await getGroupMembers(context.spHttpClient, siteUrl, openGroup.id));
      setConfirmRemove(undefined);
      showToast(`${m.title} removed — access is revoked immediately.`, false);
      log(
        EVENT.membersChanged,
        `Member removed — ${m.title} from ${openGroup.title}`,
        [
          `Group: ${openGroup.title}`,
          `Removed: ${m.title}${m.email ? ` <${m.email}>` : ""}`,
          // Removing from one group is not removing from the site. Saying so stops a reader
          // concluding the person was de-provisioned.
          `This removes ONE group. Their membership of ${siteEntryGroupTitle()} and any other group is unchanged.`,
        ],
      );
    } catch (e) {
      showToast(`Could not remove ${m.title}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

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
      if (openGroup && openGroup.id === g.id) setOpenGroup(undefined);
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
      <div style={s.card}>
        <p style={s.head}>Create a group</p>

        <label style={s.label} htmlFor="gm-name">Group name</label>
        <input
          id="gm-name"
          style={s.input}
          value={newName}
          placeholder="e.g. GHO_GF_CORU_UPLOADER"
          onChange={(e) => { setNewName(e.target.value); setNameTouched(true); }}
        />
        {newName.trim() !== "" && nameErrors.length > 0 && (
          <p style={s.err}>{nameErrors.join(" ")}</p>
        )}

        <p style={s.hint}>
          <button
            type="button"
            style={{ ...s.ghost, padding: "3px 10px" }}
            onClick={() => setBuilderOpen(!builderOpen)}
          >
            {builderOpen ? "▾" : "▸"} Build the name from a segment and unit
          </button>{" "}
          Optional — it only fills the box above. Nothing is mapped here.
        </p>

        {builderOpen && (
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
                <label style={s.label} htmlFor={`gm-lvl-${i}`}>Level {i + 1}</label>
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

            <label style={s.label} htmlFor="gm-role">Role</label>
            <select
              id="gm-role"
              style={s.select}
              value={builderRole}
              onChange={(e) => {
                const r = e.target.value as GroupMapRole | "";
                setBuilderRole(r);
                applyBuilder(mode, chosen, r);
              }}
            >
              <option value="">— select a role —</option>
              {SELECTABLE_ROLES.map((r) => <option key={r} value={r}>{roleLabel(r)}</option>)}
            </select>
            <p style={s.hint}>
              The role here only shapes the <strong>name</strong>. What the group can actually do is
              decided by the persona you give it on the Folder Access page.
            </p>
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

        <div style={{ marginTop: 10 }}>
          {loading && <p style={s.hint}>Loading…</p>}
          {!loading && visible.length === 0 && (
            <p style={s.hint}>
              {groups.length === 0 ? "No groups on this site yet." : "No group matches that filter."}
            </p>
          )}
          {visible.map((g) => {
            const rows = rowsFor(g);
            const isOpen = openGroup !== undefined && openGroup.id === g.id;
            return (
              <div key={g.id}>
                <div style={s.row}>
                  <button
                    type="button"
                    style={s.groupName}
                    onClick={() => { openMembers(g).catch(() => undefined); }}
                  >
                    {isOpen ? "▾" : "▸"} {g.title}
                  </button>
                  {rows !== undefined && (
                    rows.length === 0
                      ? <span style={s.badge}>not mapped</span>
                      : <span style={s.mapped}>{rows.length} mapping{rows.length === 1 ? "" : "s"}</span>
                  )}
                  <button type="button" style={s.danger} disabled={busy} onClick={() => setDeleting(g)}>
                    Delete
                  </button>
                </div>

                {isOpen && (
                  <div style={{ padding: "8px 10px 14px 24px", background: "#fff", borderBottom: "1px solid #f0f0f0" }}>
                    {members === undefined && <p style={s.hint}>Loading members…</p>}
                    {members !== undefined && members.length === 0 && (
                      <p style={s.hint}>Nobody is in this group.</p>
                    )}
                    {(members ?? []).map((m) => (
                      <div key={m.id} style={{ ...s.row, borderBottom: "none", padding: "4px 0" }}>
                        <span style={{ flex: 1, fontSize: 13 }}>
                          {m.title} <span style={{ color: "#666", fontSize: 12 }}>{m.email}</span>
                        </span>
                        {confirmRemove === m.id ? (
                          <>
                            <button type="button" style={s.danger} disabled={busy} onClick={() => { onRemoveMember(m).catch(() => undefined); }}>
                              Confirm remove
                            </button>
                            <button type="button" style={s.ghost} onClick={() => setConfirmRemove(undefined)}>Cancel</button>
                          </>
                        ) : (
                          <button type="button" style={s.ghost} disabled={busy} onClick={() => setConfirmRemove(m.id)}>Remove</button>
                        )}
                      </div>
                    ))}

                    <div style={{ ...s.ddwrap, marginTop: 10 }}>
                      <input
                        style={s.input}
                        value={memberQuery}
                        placeholder="Add a person…"
                        onChange={(e) => setMemberQuery(e.target.value)}
                      />
                      {memberResults.length > 0 && (
                        <div style={s.dd}>
                          {memberResults.map((p) => (
                            <div key={p.loginName} style={s.ddItem} onClick={() => { onAddMember(p).catch(() => undefined); }}>
                              {p.displayName} <span style={{ color: "#666" }}>{p.email}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                    <p style={s.hint}>
                      Everyone added here also joins <strong>{siteEntryGroupTitle()}</strong>, without
                      which they cannot open the site at all.
                    </p>
                  </div>
                )}
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
