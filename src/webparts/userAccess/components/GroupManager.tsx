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
  canonicalGroupRename,
} from "../../../shared/groupMapModel";
import { addMemberWithSiteEntry } from "../../../shared/siteEntryGroup";
import GroupMembersEditor from "./GroupMembersEditor";
import {
  fetchOwnersGroup,
  isSystemAdmin,
  renameSiteGroup,
  isSiteCollectionAdmin,
} from "../../../shared/spGroups";
import {
  searchSiteGroups,
  fetchAllGroupMembers,
  createSiteGroup,
  deleteSiteGroup,
  searchTenantPeople,
  DUPLICATE_GROUP,
  SpGroup,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
import {
  UserGroupRef,
  personDisplay,
  summarizeUserAccess,
  describeGroupAccess,
} from "../../../shared/userAccess";
import { csvCell, downloadCsv } from "../../../shared/groupExportCsv";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { PAGE_ACCESS_LINK, resolveLink } from "../../../shared/adminPages";
import { readSitePages } from "../../../shared/backToSettings";

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
type ModePick = {
  label: string;
  termSetGuid: string;
  code: string;
  levelNames: string[];
};
type TermLite = { id: string; label: string };
/** A Group Map row, only as much of it as the delete dialog has to describe. */
type MapRow = {
  itemId: number;
  groupId: string;
  role: GroupMapRole;
  segment: string;
  tier: string;
};

const s: Record<string, React.CSSProperties> = {
  card: {
    border: "1px solid #e1e1e1",
    borderRadius: 6,
    padding: 16,
    marginBottom: 20,
    background: "#fafafa",
  },
  head: { fontWeight: 700, fontSize: 20, margin: "0 0 8px" },
  label: {
    display: "block",
    fontSize: 12,
    fontWeight: 600,
    color: "#333",
    margin: "12px 0 4px",
  },
  input: {
    width: "100%",
    maxWidth: 460,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
  },
  select: {
    width: "100%",
    maxWidth: 460,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
  },
  ddwrap: { position: "relative", maxWidth: 460 },
  dd: {
    position: "absolute",
    top: "100%",
    left: 0,
    right: 0,
    zIndex: 20,
    background: "#fff",
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    maxHeight: 240,
    overflowY: "auto",
    boxShadow: "0 4px 12px rgba(0,0,0,.12)",
  },
  ddItem: {
    padding: "7px 10px",
    cursor: "pointer",
    borderBottom: "1px solid #f0f0f0",
    fontSize: 13,
  },
  btn: {
    padding: "6px 16px",
    fontSize: 13,
    background: "#0f6c3f",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    cursor: "pointer",
  },
  btnOff: {
    padding: "6px 16px",
    fontSize: 13,
    background: "#c7c7c7",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    cursor: "not-allowed",
  },
  ghost: {
    padding: "5px 12px",
    fontSize: 12,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  danger: {
    padding: "5px 12px",
    fontSize: 12,
    color: "#a4262c",
    border: "1px solid #a4262c",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  chip: {
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    padding: "3px 8px",
    margin: "0 6px 6px 0",
    fontSize: 12,
    background: "#eef4ff",
    border: "1px solid #cfe0ff",
    borderRadius: 12,
  },
  row: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "8px 10px",
    borderBottom: "1px solid #f0f0f0",
  },
  groupName: {
    fontWeight: 600,
    fontSize: 13,
    flex: 1,
    textAlign: "left",
    background: "none",
    border: "none",
    cursor: "pointer",
    padding: 0,
    color: "#1b1b1b",
  },
  badge: {
    fontSize: 11,
    padding: "2px 8px",
    borderRadius: 10,
    background: "#fff4ce",
    border: "1px solid #f2d98c",
    color: "#7a5c00",
  },
  mapped: {
    fontSize: 11,
    padding: "2px 8px",
    borderRadius: 10,
    background: "#e7f4ec",
    border: "1px solid #b7dcc4",
    color: "#0f6c3f",
  },
  err: { fontSize: 12, color: "#a4262c", margin: "6px 0 0" },
  hint: { fontSize: 11, color: "#666", marginTop: 6, lineHeight: 1.45 },
  link: { color: "#0f6c3f", fontWeight: 600, textDecoration: "underline" },
  /* The user lookup's own row. It used to borrow `row`, which is `display:flex` for the LIST's
     name + badge + Delete line — so five paragraphs sat side by side at ragged widths, vertically
     centred, with the long ones stretching the row (client, 2026-08-23: "the design is not tidy").
     A lookup row is a STACK: who, what it gives, then the roles. */
  lkSearchRow: { display: "flex", gap: 8, alignItems: "center" },
  lkRow: { padding: "10px 2px", borderBottom: "1px solid #ececec" },
  lkTop: {
    display: "flex",
    alignItems: "baseline",
    justifyContent: "space-between",
    gap: 12,
    flexWrap: "wrap",
  },
  lkWho: { fontWeight: 600, fontSize: 13, margin: "0 0 2px", color: "#1b1b1b" },
  lkName: { fontWeight: 600, fontSize: 13, margin: 0, color: "#1b1b1b" },
  lkSeg: { fontSize: 11, color: "#767676", margin: 0 },
  lkPersona: { fontSize: 12, color: "#1b1b1b", margin: "3px 0 0" },
  lkSum: {
    fontSize: 11,
    color: "#666",
    margin: "2px 0 0",
    lineHeight: 1.45,
    maxWidth: 760,
  },
  lkChips: { margin: "6px 0 0" },
  warnBox: {
    padding: "10px 12px",
    border: "1px solid #f2c9a0",
    background: "#fff8f0",
    borderRadius: 4,
    fontSize: 12,
    color: "#8a4b00",
    lineHeight: 1.5,
    marginBottom: 12,
  },
  okBox: {
    padding: "10px 12px",
    border: "1px solid #b7dcc4",
    background: "#f3faf5",
    borderRadius: 4,
    fontSize: 12,
    color: "#0f6c3f",
    lineHeight: 1.5,
    marginBottom: 12,
  },
  modalBg: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,.4)",
    zIndex: 100,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  modal: {
    background: "#fff",
    borderRadius: 6,
    padding: 20,
    width: "min(560px, 92vw)",
    maxHeight: "84vh",
    overflowY: "auto",
    boxShadow: "0 8px 32px rgba(0,0,0,.25)",
  },
  toast: {
    position: "fixed",
    bottom: 20,
    right: 20,
    padding: "10px 16px",
    borderRadius: 4,
    fontSize: 13,
    color: "#fff",
    zIndex: 200,
    maxWidth: 460,
    lineHeight: 1.45,
  },
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

export default function GroupManager({
  context,
  siteUrl,
  hideCreateForm,
  refreshKey,
}: Props): React.ReactElement {
  const [canManage, setCanManage] = useState<boolean | undefined>(undefined);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [groups, setGroups] = useState<SpGroup[]>([]);
  /**
   * The site OWNERS group - system administrators.
   *
   * WARN: FETCHED SEPARATELY BECAUSE `searchSiteGroups` EXCLUDES IT ON PURPOSE, so nobody maps
   * "Site Owners" onto a unit folder. That exclusion stays; this is the other job - putting a person
   * INTO Owners, which is what makes them an administrator of the CRS system. Reconciliation's
   * lockdown pass strips every non-Owners assignment from every admin page, so Owners membership is
   * the ONLY thing that opens them.
   *
   * `undefined` means the read failed and the card says so, rather than implying the site has no
   * administrators.
   */
  const [owners, setOwners] = useState<SpGroup | undefined>(undefined);
  const [ownersOpen, setOwnersOpen] = useState<boolean>(false);
  /**
   * Is the VIEWER a site collection administrator - not merely an owner?
   *
   * Only an SCA may promote anyone, so without this the card would let an owner add people and every
   * promotion would fail with an error they could not interpret. `undefined` means the read failed and
   * shows NOTHING: a false warning on every visit is how the real one gets ignored.
   */
  const [viewerIsSca, setViewerIsSca] = useState<boolean | undefined>(
    undefined,
  );
  /**
   * Links to the three now-read-only reports (2026-09-02) — Site Access, Approval Library Access,
   * Page Access. Resolved from Site Pages, never hardcoded, same lookup and same `resolveLink` every
   * other cross-page link in this project already uses (the client renames pages at import). A
   * `missing` page leaves its href undefined rather than a dead link; the button is simply hidden.
   */
  const [pageAccessHref, setPageAccessHref] = useState<string | undefined>(
    undefined,
  );
  /** Progress of a standardise-names run. Shares `bulkRunning` and the unload guard with delete. */
  const [renameProgress, setRenameProgress] = useState<
    { done: number; total: number; failed: number } | undefined
  >(undefined);
  const [mapRows, setMapRows] = useState<MapRow[] | undefined>(undefined);
  const [filter, setFilter] = useState("");
  const [busy, setBusy] = useState(false);
  const [toast, setToast] = useState<
    { message: string; error: boolean } | undefined
  >(undefined);
  /**
   * Which group's people are open, by ID. ONE at a time.
   *
   * Client, 2026-08-23: *"all we have to do is to add a way to let client to add user inside the
   * Group Management"*. This reverses the 2026-08-18 move to Folder Access — asked for by the same
   * client, whose reading now is that Folder Access, Site Access and Approval Library Access are all
   * redundant once bulk provisioning writes the groups AND their mappings. Putting a person into
   * their unit's group is the one recurring job in this system, so it belongs on the page they will
   * actually be on.
   *
   * `GroupMembersEditor` is MOUNTED, never copied — it is still the single implementation, and it
   * reads its own members on mount, so only the open group costs a request.
   */
  const [openGroup, setOpenGroup] = useState<number | undefined>(undefined);
  /**
   * Every group's members, keyed by group id, read in ONE request.
   *
   * `undefined` is NOT "no members" — it is "the read failed", and the badge falls back to the
   * mapping count rather than claiming every group is empty. Same rule as `mapRows`.
   */
  const [memberIndex, setMemberIndex] = useState<
    Record<number, SpGroupMember[]> | undefined
  >(undefined);

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
  /**
   * Ticked groups, by id. BY ID and not by name, for the same reason every join in this file is:
   * a rename keeps the id, and two groups can end up named alike.
   *
   * Selection deliberately SURVIVES a filter change — an admin narrowing to `APPROVER`, ticking
   * them, then narrowing to another segment and ticking more is the actual workflow. The bar
   * therefore reports the total selected AND how many of them are currently out of view, because a
   * count that silently ignored the hidden ones would delete more than the screen shows.
   */
  const [selected, setSelected] = useState<Record<number, boolean>>({});
  const [bulkOpen, setBulkOpen] = useState(false);
  /** Typed confirmation for the bulk delete. Nothing else in this file needs one; this does. */
  const [bulkTyped, setBulkTyped] = useState("");
  const [bulkProgress, setBulkProgress] = useState<
    { done: number; total: number; failed: number } | undefined
  >(undefined);
  /**
   * Belt to `busy`'s braces, and a REF because state cannot do this job.
   *
   * `disabled={busy}` is set from state, so it only takes effect on the next render — a fast double
   * click, or a click landing while React is still committing, can enter the runner twice. Twice
   * through this one means deleting a group that no longer exists (harmless) and, worse, a second
   * progress counter fighting the first. The ref flips SYNCHRONOUSLY, so the second entry returns
   * before it does anything.
   */
  const bulkRunning = useRef(false);

  /* ── User lookup ────────────────────────────────────────────────────────
     `lookupGroups` is THREE-STATE and must stay so: `undefined` = not looked up or the read failed,
     `[]` = looked up and they are in nothing, a list = these. Collapsing the first two would render
     a failed read as "this person has no access", which is the answer that stops an admin looking. */
  const [lookupQuery, setLookupQuery] = useState("");
  const [lookupResults, setLookupResults] = useState<PersonPick[]>([]);
  const [lookupPerson, setLookupPerson] = useState<PersonPick | undefined>(
    undefined,
  );
  const [lookupGroups, setLookupGroups] = useState<UserGroupRef[] | undefined>(
    undefined,
  );
  const [lookupNote, setLookupNote] = useState<string | undefined>(undefined);
  const [lookupBusy, setLookupBusy] = useState(false);

  /**
   * Segment filter for the group list. "" = every segment, "__none__" = groups with no mapping rows.
   *
   * The unmapped bucket is the point of it as much as the segments are: on a provisioned site the
   * only groups an admin needs to FIND are the ones that grant nothing, and they are otherwise
   * scattered through several hundred rows.
   */
  const [segFilter, setSegFilter] = useState("");

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
      return (
        (data.value ?? []) as Array<{
          Id: number;
          GroupId?: string;
          Segment?: string;
          UnitTermGuid?: string;
          Role?: string;
        }>
      ).map((r) => ({
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
      // Primes first, for the same reason as loadModes: `abbrevListTitle()` reads the same module
      // cache, and unprimed it answers the legacy DMS title, which 404s on a renamed site.
      await primeNames(context.spHttpClient, siteUrl);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(abbrevListTitle())}')/items` +
          `?$select=TermGuid,Abbreviation&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return {};
      const data = await res.json();
      const out: Record<string, string> = {};
      for (const r of (data.value ?? []) as Array<{
        TermGuid?: string;
        Abbreviation?: string;
      }>) {
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
  const rowsForNewGroup = (
    groupId: string,
    groupName: string,
  ): GroupMapWriteRow[] => {
    const p = PERSONAS.filter((x) => x.key === builderPersona)[0];
    if (!p || advanced || !mode) return [];
    const scope = personaScope(p.key);
    // A SEGMENT-scope row carries UnitTermGuid === the term-set guid; that is how Folder Access records
    // "the segment IS the tier", and reconciliation reads it the same way.
    const tierGuid =
      scope === "segment"
        ? mode.termSetGuid
        : scope === "department"
          ? (chosen[0]?.id ?? "")
          : (chosen[chosen.length - 1]?.id ?? "");
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
    /* ⚠ PRIMES FIRST, and this was a live bug until 2026-08-23. `CONFIG_LIST()` reads a module cache
       that `primeNames` fills; unprimed it answers the LEGACY "DMS Config", which 404s on a
       CRS-renamed site. The mount effect starts `reload()` (which primes) and then calls this in the
       SAME TICK, and priming makes up to ten sequential probes — so this always lost the race,
       returned [] on `!res.ok`, and the Segment dropdown on the create form was permanently empty.
       Silent in every direction: no error, no console line, just no segments.

       Found because the user lookup showed "segment not known" on every row. THIS IS THE 1.0.207.0
       BUG IN A SECOND COMPONENT — the same fix (await priming INSIDE the loader, rather than gating
       on a `namesReady` flag) was applied to `loadReconModes` and never to this call site. */
    await primeNames(context.spHttpClient, siteUrl);
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(CONFIG_LIST())}')/items?$select=ModeLabel,TermSetGuid,StagingFolder,Levels&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (
      (data.value ?? []) as Array<{
        ModeLabel?: string;
        TermSetGuid?: string;
        StagingFolder?: string;
        Levels?: string;
      }>
    )
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
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (
      (data.value ?? []) as Array<{
        id: string;
        labels: Array<{ name: string }>;
      }>
    ).map((t) => ({ id: t.id, label: t.labels[0].name }));
  };

  // Full Control is needed to create, delete, and change membership. Detected up front so the
  // controls are explained rather than failing with a 403 on click.
  /* This page's own rule, extracted to `isSystemAdmin` on 2026-08-27 and now shared with the audit
     log, the upload form and Bulk Upload. It was right here and nowhere else, which is exactly why
     an Owner was an administrator on this screen and on no other. */
  const loadCanManage = async (): Promise<boolean> =>
    isSystemAdmin(context.spHttpClient, siteUrl);

  const reload = async (): Promise<void> => {
    setLoading(true);
    try {
      await primeNames(context.spHttpClient, siteUrl);
      setGroups(await loadGroups());
      // Cannot fail the reload: the rest of the page works without it, and it answers `undefined`.
      setOwners(await fetchOwnersGroup(context.spHttpClient, siteUrl));
      setViewerIsSca(
        await isSiteCollectionAdmin(context.spHttpClient, siteUrl),
      );
      setMapRows(await loadMapRows());
      // AFTER the groups, and never allowed to fail the reload: the list and its create form work
      // perfectly without member counts, so a heavy $expand that throttles must not blank the page.
      // `fetchAllGroupMembers` answers `undefined` rather than throwing.
      setMemberIndex(await fetchAllGroupMembers(context.spHttpClient, siteUrl));
      setLoadError(undefined);
    } catch (e) {
      setLoadError((e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  /**
   * Re-read the member counts alone, after the editor below adds or removes somebody.
   *
   * Not `reload()`: that re-reads 586 groups and every Group Map row to answer a question about one
   * person. The badge is the only thing that went stale.
   */
  const reloadMembers = (): void => {
    fetchAllGroupMembers(context.spHttpClient, siteUrl)
      .then((m) => {
        if (m !== undefined) setMemberIndex(m);
      })
      .catch(() => undefined);
  };

  /* ── Mount ─────────────────────────────────────────────────────────────── */

  useEffect(() => {
    loadCanManage()
      .then(setCanManage)
      .catch(() => setCanManage(false));
    reload().catch(() => undefined);
    loadModes()
      .then(setModes)
      .catch(() => undefined);
    loadCodes()
      .then(setCodes)
      .catch(() => undefined);
  }, []);

  /**
   * Resolve the Page Access report link from Site Pages — same lookup and `resolveLink` every other
   * cross-page link in this project already uses. A `missing` page just leaves the button out rather
   * than a dead link.
   *
   * Site Access and Approval Library Access are GONE (2026-09-02, client: *"Approval Library Access,
   * Site Access is not needed, its confusing them as they already know that adding a user in the
   * group from Group Management it will automatically allow them have access to site and library.
   * Just keep Page access."*). Both pages themselves are retired as signposts pointing back HERE
   * (`SiteAccessPage.tsx`, `ApprovalLibraryAccessPage.tsx`), so a link to either from this screen
   * would only lead an admin in a circle.
   */
  useEffect(() => {
    readSitePages(context, siteUrl)
      .then((pages) => {
        const page = resolveLink(PAGE_ACCESS_LINK, pages);
        setPageAccessHref(page.state === "missing" ? undefined : page.url);
      })
      .catch(() => undefined);
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
    if (q.length < 2) {
      setPeopleResults([]);
      return;
    }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then(setPeopleResults)
        .catch(() => setPeopleResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [peopleQuery]);

  // Same debounce for the lookup's own picker. Deliberately SEPARATE state from the create form's:
  // one picker feeding two jobs would mean choosing someone to inspect also stages them for a group
  // nobody asked to create.
  useEffect(() => {
    const q = lookupQuery.trim();
    if (q.length < 2) {
      setLookupResults([]);
      return;
    }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then(setLookupResults)
        .catch(() => setLookupResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [lookupQuery]);

  /* ── "What can this person reach?" ──────────────────────────────────────
     The question none of the access pages answered. Client, 2026-08-23: "if the group is already auto
     assign then what is the point of the pages?" — bulk provisioning writes the groups AND their
     mapping rows, so the jobs that remain are putting someone in a group and asking what that gives
     them. */

  /** Back to an empty panel: the box, the candidates, the person and their answer. */
  const clearLookup = (): void => {
    setLookupQuery("");
    setLookupResults([]);
    setLookupPerson(undefined);
    setLookupGroups(undefined);
    setLookupNote(undefined);
  };

  const runLookup = async (p: PersonPick): Promise<void> => {
    setLookupPerson(p);
    setLookupGroups(undefined);
    setLookupNote(undefined);
    setLookupResults([]);
    setLookupQuery("");
    setLookupBusy(true);
    try {
      // Matched on LOGIN NAME, not email: a guest can exist TWICE on one site for one address
      // (2026-08-18), and an account can carry no email at all — in which case an email filter finds
      // nobody, which would read as "this person has no access".
      const login = (p.loginName ?? "").trim();
      if (!login) {
        setLookupNote(
          "This account has no login name, so its site membership cannot be looked up.",
        );
        return;
      }
      const usersRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/siteusers?$select=Id,Title,LoginName&$filter=LoginName eq '${encodeURIComponent(login.replace(/'/g, "''"))}'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!usersRes.ok) {
        // Names the status. 403 is not holding rights to enumerate site users, 404 is a wrong URL —
        // opposite fixes, the same rule as the library ACL banner.
        setLookupNote(
          `Could not read this site's users — HTTP ${usersRes.status}.`,
        );
        return;
      }
      const users = ((await usersRes.json()).value ?? []) as Array<{
        Id: number;
        Title?: string;
      }>;
      if (users.length === 0) {
        // A REAL answer, not a failure: they exist in the directory and have never been added here.
        setLookupGroups([]);
        // ⚠ TWO CAUSES, AND THE FIRST IS FAR MORE LIKELY (client, 2026-08-23, having searched
        // `clarencechojinheng@gmai.com`): the address was mistyped, or the person genuinely has no
        // account on this site. The picker resolves a typed address to ITSELF —
        // `searchTenantPeople` runs with `AllowEmailAddresses: true`, so anything shaped like an
        // email comes back as a candidate whether or not it belongs to anybody. Reporting only
        // "not a member yet" invites an admin to go and add a person who does not exist.
        setLookupNote(
          "No account on this site matches that address. Check the spelling — an address typed in " +
            "full is offered in the list whether or not it belongs to anybody. If the spelling is " +
            "right, they have simply never been given access to this site.",
        );
        return;
      }
      if (users.length > 1) {
        setLookupNote(
          `⚠ ${users.length} accounts on this site share this login. Their groups are merged below — ` +
            `a guest can exist twice for one address, and the two can hold different access.`,
        );
      }
      const all: UserGroupRef[] = [];
      for (const u of users) {
        // $top well past any real membership: a truncated read is indistinguishable from a group the
        // person is not in, which here would UNDERSTATE their access.
        const gr: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetUserById(${u.Id})/groups?$select=Id,Title&$top=5000`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!gr.ok) throw new Error(`groups HTTP ${gr.status}`);
        for (const x of ((await gr.json()).value ?? []) as Array<{
          Id: number;
          Title?: string;
        }>) {
          if (!all.some((a) => a.id === x.Id))
            all.push({ id: x.Id, title: (x.Title ?? "").trim() });
        }
      }
      setLookupGroups(all);
    } catch (e) {
      // `undefined`, never [] — "could not read" must not render as "no access", which is the answer
      // that stops an admin looking.
      setLookupGroups(undefined);
      setLookupNote(`Could not read this person's groups — ${String(e)}`);
    } finally {
      setLookupBusy(false);
    }
  };

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
  const applyBuilder = (
    m: ModePick | undefined,
    path: TermLite[],
    persona: string,
  ): void => {
    if (!m) return;
    // Code per tier, falling back to the LABEL when a term has no code yet. A fallback rather than a
    // refusal because this box only suggests a name — but such a group will not line up with its folder,
    // so the hint below the field says which tiers are missing a code.
    // TRUNCATED BY THE PERSONA'S SCOPE. A Head of Department group is named `GHO_GF_HOD`, not
    // `GHO_GF_TAX_HOD` — its row sits on the DEPARTMENT term and reaches the units beneath by fan-out, so
    // a unit in its name would describe a narrower grant than it has. C-Level is segment-wide and takes no
    // tier at all.
    const scope = personaScope(persona);
    const depth =
      scope === "segment" ? 0 : scope === "department" ? 1 : path.length;
    const built = suggestGroupName(
      m.code || m.label,
      path.slice(0, depth).map((t) => codes[t.id.toLowerCase()] || t.label),
      namingRoleFor(persona),
    );
    // Never silently overwrite something the admin typed. The name is the one field here they may
    // have composed by hand for a reason.
    if (nameTouched && newName.trim() && newName.trim() !== built) {
      if (!window.confirm(`Replace the name you typed with "${built}"?`))
        return;
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
    const tops = await loadTerms(
      `${siteUrl}/_api/v2.1/termStore/sets/${m.termSetGuid}/children`,
    );
    setLevels(tops.length ? [tops] : []);
    applyBuilder(m, [], builderPersona);
  };

  const pickTerm = async (
    levelIndex: number,
    termId: string,
  ): Promise<void> => {
    if (!mode) return;
    const options = levels[levelIndex] ?? [];
    const term = options.find((t) => t.id === termId);
    const path = term
      ? [...chosen.slice(0, levelIndex), term]
      : chosen.slice(0, levelIndex);
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

  const nameErrors = validateGroupName(
    newName,
    groups.map((g) => g.title),
  );
  const canCreate = canManage === true && !busy && nameErrors.length === 0;

  const onCreate = async (): Promise<void> => {
    if (!canCreate) return;
    setBusy(true);
    try {
      const made = await createSiteGroup(
        context.spHttpClient,
        siteUrl,
        newName.trim(),
      );

      // No rollback, deliberately. Folder Access deleted the group when a mapping row failed,
      // because a group created FOR a mapping that does not exist was garbage. Here the group is
      // the deliverable — a member who fails to add is reported, and the group is kept.
      let added = 0;
      let failed = 0;
      const notes: string[] = [];
      for (const p of staged) {
        try {
          const r = await addMemberWithSiteEntry(
            context.spHttpClient,
            siteUrl,
            { id: made.id, title: made.title },
            p.loginName,
          );
          added++;
          if (r.note) notes.push(`${p.displayName}: ${r.note}`);
        } catch (e) {
          failed++;
          notes.push(
            `${p.displayName}: could not be added — ${(e as Error).message}`,
          );
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
      // ⚠ NEITHER MESSAGE MAY NAME FOLDER ACCESS ANY MORE. It stopped creating mappings on
      // 2026-08-18 (client, twice), and the hand-mapping form left this page on 2026-08-23 — so
      // "assign it on Folder Access" now sends an admin to a screen that CANNOT do it, and they
      // would go on believing the group is one step from working. With no persona there is no
      // mapping screen left anywhere: delete it and re-create it with one.
      const mapNote =
        wanted.length === 0
          ? " It grants nothing, and no persona was chosen — so no mappings were written and there is" +
            " no screen left that can add them. To give it access, delete it and create it again" +
            " with a persona."
          : mapped === wanted.length
            ? ` Mapped at ${personaScope(builderPersona)} level with ${mapped} role${mapped === 1 ? "" : "s"} — run Folder Reconciliation to apply it.`
            : ` ⚠ Only ${mapped} of ${wanted.length} mappings were written. Press Create again with the` +
              ` same persona to finish the rest before reconciling.`;
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
    mapRows === undefined
      ? undefined
      : mapRows.filter((r) => r.groupId === String(g.id));

  /**
   * Delete one group and its mapping rows. NO toast, NO reload, NO audit row — the caller owns
   * those, because deleting 130 groups must not mean 130 toasts and 130 full re-reads.
   *
   * THROWS on a failed group delete so a bulk run can count it and carry on; a failed ROW delete is
   * counted instead, since a group whose rows partly survive is still better deleted than left
   * behind (reconciliation names the leftovers on its next run).
   */
  const deleteGroupCore = async (
    g: SpGroup,
  ): Promise<{ rowsRemoved: number; rowsTotal: number }> => {
    const rows = rowsFor(g) ?? [];
    // Rows first: a deleted group whose rows survive leaves reconciliation trying to grant a
    // principal that no longer exists, on every run.
    let rowsRemoved = 0;
    for (const r of rows) {
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(GROUP_MAP_LIST())}')/items(${r.itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
          },
        },
      );
      if (res.ok) rowsRemoved++;
    }
    await deleteSiteGroup(context.spHttpClient, siteUrl, g.id);
    return { rowsRemoved, rowsTotal: rows.length };
  };

  const onDeleteGroup = async (g: SpGroup): Promise<void> => {
    setBusy(true);
    try {
      const rows = rowsFor(g) ?? [];
      const core = await deleteGroupCore(g);
      const rowsRemoved = core.rowsRemoved;
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
          // The distinction that makes this row worth reading — and it was recorded BACKWARDS
          // until 2026-08-25. Deleting the GROUP withdraws its access immediately, because
          // SharePoint drops a deleted principal's role assignments; only deleting a mapping ROW
          // leaves the ACL standing.
          rowsRemoved > 0
            ? "Folder permissions withdrawn with the group. Re-create it and reconcile to grant again."
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

  /**
   * Closing or reloading the tab mid-run stops it part-way, and would otherwise say nothing.
   *
   * Keyed on the LONG-RUN progress states, NOT on `busy`: `busy` is also true for a single delete and
   * for creating a group, and a browser prompt on a one-second operation is noise that teaches people
   * to dismiss it. This fires only while a long run is actually in flight. Same guard as the bulk
   * group provisioner and the staged batches in BulkUpload.
   *
   * WARN: `renameProgress` HAD TO BE ADDED HERE EXPLICITLY (2026-08-27). The standardise-names run
   * was written "sharing the unload guard" and did not - the guard names its states one by one, so a
   * new long run is NOT covered until it is listed. Anything else added here must be added here too.
   */
  useEffect(() => {
    if (bulkProgress === undefined && renameProgress === undefined)
      return undefined;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [bulkProgress === undefined, renameProgress === undefined]);

  /** The ticked groups, in list order. Read from `groups`, never from `visible` — see `selected`. */
  const selectedGroups = (): SpGroup[] => groups.filter((g) => selected[g.id]);

  /**
   * Delete every ticked group, one at a time.
   *
   * WHY THIS EXISTS (2026-08-25, client mid-migration: "Can you add like select multiple checkboxes?
   * that would be easier for me to deelte."). Withdrawing a permission from a persona needs the
   * GROUP deleted — reconciliation adds folder grants and never removes them — so the 2026-08-24 HC
   * split left ~130 `_APPROVER` groups per site to delete by hand. A hundred and thirty individual
   * confirm dialogs is not a workflow, it is an invitation to click the wrong row.
   *
   * SEQUENTIAL, not parallel: these are permission writes against a throttling tenant, and a burst
   * of 130 deletes is the shape that gets a run cut off half way.
   *
   * CARRIES ON PAST A FAILURE and reports the names at the end. Stopping on the first error would
   * leave the admin to work out which of 130 had gone, and the operation is idempotent anyway — a
   * group already deleted is simply absent from the next read.
   */
  /**
   * Groups on this site whose name uses an OLD suffix spelling, with the name they should have.
   *
   * Derived at render from the loaded list - no extra request. `canonicalGroupRename` returns
   * `undefined` for anything it does not recognise, which is what keeps `CRS_SITE_MEMBERS`, the
   * owners group and every hand-named group out of this.
   */
  const renameTargets = (): { id: number; from: string; to: string }[] => {
    const out: { id: number; from: string; to: string }[] = [];
    for (const g of groups) {
      const to = canonicalGroupRename(g.title);
      if (to !== undefined) out.push({ id: g.id, from: g.title, to });
    }
    return out;
  };

  /**
   * Rename every group still using an old suffix to the canonical one.
   *
   * Client, 2026-08-27: *"you forgot to rename the APR to APPROVAL and UPL to UPLOADER."* 1.0.261.0
   * changed what NEW names are written and left existing groups alone, so a provisioned site is
   * MIXED. This is the migration.
   *
   * WARN: SAFE BECAUSE A RENAME PRESERVES THE GROUP ID. Group Map rows are keyed on `GroupId`, so
   * every mapping, folder grant and page ACL survives untouched; only the stored `GroupName` label
   * goes stale, and `planBulkGroups` matches on the TERM GUID before the name anyway.
   *
   * Sequential and carries on past failures, for the same reasons as the bulk delete beside it: a
   * burst of ~500 permission writes against a throttling tenant is the shape that gets cut off half
   * way, and stopping on the first error leaves the admin to work out how far it got. Shares
   * `bulkRunning` with delete, and `renameProgress` is listed in the unload guard.
   */
  const runStandardiseNames = async (): Promise<void> => {
    const targets = renameTargets();
    if (targets.length === 0) return;
    if (bulkRunning.current) return;
    bulkRunning.current = true;
    setBusy(true);
    setRenameProgress({ done: 0, total: targets.length, failed: 0 });
    let done = 0;
    let failed = 0;
    const failures: string[] = [];
    for (const t of targets) {
      try {
        await renameSiteGroup(context.spHttpClient, siteUrl, t.id, t.to);
      } catch (e) {
        failed++;
        // A DUPLICATE is the one worth naming: it means both spellings exist and one must be
        // merged or deleted by hand. Renaming cannot resolve that on its own.
        failures.push(`${t.from} → ${t.to} — ${(e as Error).message}`);
      }
      done++;
      setRenameProgress({ done, total: targets.length, failed });
    }
    log(
      EVENT.groupMapChanged,
      `${done - failed} group(s) renamed to the standard suffixes`,
      [
        `Renamed: ${done - failed} of ${targets.length}`,
        ...(failures.length > 0 ? [`FAILED: ${failures.join(" | ")}`] : []),
        "A rename preserves the group id, so mappings, folder grants and page access are unchanged.",
      ],
      failed === 0 ? "Success" : "Failed",
    );
    setRenameProgress(undefined);
    bulkRunning.current = false;
    setBusy(false);
    showToast(
      failed === 0
        ? `${done} group(s) renamed. Access is unchanged — a rename keeps the group id.`
        : `${done - failed} renamed, ${failed} failed. See the audit log for the names.`,
      failed !== 0,
    );
    await reload();
  };

  const runBulkDelete = async (): Promise<void> => {
    const targets = selectedGroups();
    if (targets.length === 0) return;
    if (bulkRunning.current) return;
    bulkRunning.current = true;
    setBusy(true);
    setBulkProgress({ done: 0, total: targets.length, failed: 0 });
    let done = 0;
    let failed = 0;
    let rowsRemoved = 0;
    const failures: string[] = [];
    const deletedNames: string[] = [];
    for (const g of targets) {
      try {
        const r = await deleteGroupCore(g);
        rowsRemoved += r.rowsRemoved;
        deletedNames.push(g.title);
      } catch (e) {
        failed++;
        failures.push(`${g.title} — ${(e as Error).message}`);
      }
      done++;
      setBulkProgress({ done, total: targets.length, failed });
    }
    // ONE audit row for the run, not one per group — the same rule reconciliation follows, and for
    // the same reason: 130 rows would bury every other event on the day it is used.
    log(
      EVENT.groupDeleted,
      `${deletedNames.length} group(s) deleted in one action`,
      [
        `Deleted: ${deletedNames.length} of ${targets.length}`,
        `Group Map rows removed: ${rowsRemoved}`,
        ...(failures.length > 0 ? [`FAILED: ${failures.join(" | ")}`] : []),
        "Deleting a group DOES withdraw its folder permissions immediately — SharePoint drops a " +
          "deleted principal's role assignments. Re-create the groups and reconcile to grant again.",
        ...deletedNames.slice(0, 200),
      ],
      failures.length > 0 ? "Failed" : "Success",
    );
    setSelected({});
    setBulkOpen(false);
    setBulkTyped("");
    setBulkProgress(undefined);
    setBusy(false);
    bulkRunning.current = false;
    await reload();
    showToast(
      `${deletedNames.length} group(s) deleted` +
        (rowsRemoved > 0 ? `, ${rowsRemoved} mapping row(s) removed` : "") +
        (failed > 0 ? `. ${failed} FAILED — see the audit log.` : ".") +
        " Re-create them with Bulk provisioning, then run Folder Reconciliation to grant them again.",
      failed > 0,
    );
  };

  /* ── Render ────────────────────────────────────────────────────────────── */

  /**
   * Which segment a group belongs to, from its MAPPING ROWS — not from its name.
   *
   * The name usually carries the segment code (`GHO_…`), but a hand-named group need not, and the
   * abbreviation can be re-coded while the rows stay put. The rows are what reconciliation reads, so
   * they are what this filter should agree with.
   *
   * `undefined` when the Group Map could not be read: the filter then offers nothing rather than
   * quietly reporting every group as belonging to no segment.
   */
  const segmentsFor = (g: SpGroup): string[] | undefined => {
    const rows = rowsFor(g);
    if (rows === undefined) return undefined;
    const out: string[] = [];
    for (const r of rows) {
      const k = (r.segment ?? "").trim().toLowerCase();
      if (k && out.indexOf(k) === -1) out.push(k);
    }
    return out;
  };

  /** Ticked count — see `selected`. */
  const selectedCount = groups.filter((g) => selected[g.id]).length;
  const visible = groups
    .filter((g) => {
      if (segFilter) {
        const segs = segmentsFor(g);
        // Unreadable map ⇒ show everything. Hiding rows because a list could not be read is how a
        // filter comes to look like "these groups do not exist".
        if (segs !== undefined) {
          if (segFilter === "__none__") {
            if (segs.length > 0) return false;
          } else if (segs.indexOf(segFilter.toLowerCase()) === -1) {
            return false;
          }
        }
      }
      return true;
    })
    .filter(
      (g) =>
        !filter.trim() ||
        g.title.toLowerCase().indexOf(filter.trim().toLowerCase()) !== -1,
    );

  /**
   * Ticked but filtered out of view. Named in the bar, because a selection that quietly reaches
   * past the screen is how somebody deletes a segment they were not looking at.
   */
  const hiddenSelected =
    selectedCount - visible.filter((g) => selected[g.id]).length;
  /**
   * Selected groups that still have members — the ONLY irreversible part of a bulk delete.
   *
   * Empty when the member index could not be read, which is the safe direction here: it says
   * nothing rather than claiming every group is empty, and the dialog's standing warning about
   * membership loss covers the unknown case.
   */
  const peopleAtRisk =
    memberIndex === undefined
      ? []
      : groups
          .filter((g) => selected[g.id] && (memberIndex[g.id] ?? []).length > 0)
          .map((g) => g.title);

  /**
   * CSV of the groups CURRENTLY SHOWN — an inventory, deliberately not the full access list.
   *
   * ⚠ It does NOT reuse `GroupExportRow`, and that is the point. That shape carries `tier1`/`tier2`,
   * and a blank `tier2` MEANS "department-scope row" by its own contract — so emitting blanks here,
   * where the tier chain is not resolved, would state something false about every unit row. Folder
   * Access already exports the full mapping detail WITH tiers and members, because it walks each
   * segment's tree to get them; this page answers "what groups exist", and its export says that and
   * no more. The header row is what tells the two apart when both files are on someone's desktop.
   */
  const exportShown = (): void => {
    // "People" is the column the client asked for by name (2026-08-23: *"the Excel doesnt show who
    // is inside each group"*) — the export listed what a group GRANTS and never who held it, so it
    // could not answer "who has access to Tax", which is the only question anyone opens it to ask.
    // Count first so the sheet can be sorted on it without splitting the names.
    const header = [
      "Group",
      "People",
      "Members",
      "Segment",
      "Roles",
      "Mappings",
    ];
    const lines = [header.map(csvCell).join(",")];
    for (const g of visible) {
      const rows = rowsFor(g);
      const segs =
        rows === undefined
          ? "not known"
          : (segmentsFor(g) ?? [])
              .map(
                (k) =>
                  modes.filter((m) => m.termSetGuid.toLowerCase() === k)[0]
                    ?.label ?? k,
              )
              .join("; ");
      const roles =
        rows === undefined
          ? "not known"
          : rows
              .map((r) => roleLabel(r.role))
              .filter((v, i, a) => a.indexOf(v) === i)
              .join("; ");
      // undefined — the read failed — must NOT export as 0 and an empty cell: a spreadsheet saying
      // a unit's approver group has nobody in it is acted on.
      const people =
        memberIndex === undefined ? undefined : (memberIndex[g.id] ?? []);
      lines.push(
        [
          csvCell(g.title),
          csvCell(people === undefined ? "not known" : String(people.length)),
          csvCell(
            people === undefined
              ? "not known"
              : people
                  .map((m) => (m.email ? `${m.title} <${m.email}>` : m.title))
                  .join("; "),
          ),
          csvCell(rows === undefined ? "not known" : segs || "(none)"),
          csvCell(rows === undefined ? "not known" : roles || "(none)"),
          csvCell(rows === undefined ? "not known" : String(rows.length)),
        ].join(","),
      );
    }
    // Its OWN filename. `exportFileName()` returns "CRS-group-members-<date>.csv", which names a
    // report this file is not — it carries no members. No String.padStart: the SPFx tsconfig does
    // not target a lib that has it (the same constraint noted in groupExportCsv.ts).
    const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
    const d = new Date();
    downloadCsv(
      lines.join("\r\n"),
      `CRS-groups-${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.csv`,
    );
  };

  if (canManage === false) {
    return (
      <p style={s.warnBox}>
        Managing groups needs <strong>Full Control</strong> on this site. Ask a
        site owner to add you, or to make the change for you.
      </p>
    );
  }

  return (
    <div>
      {/* ⚠ THE INTRO BANNER THAT USED TO SIT HERE ("Creating a group here grants nothing… Choosing
          a persona below…") IS GONE (2026-09-02, client's confirmed call). It referred to the
          create-one-group form's persona picker "below" — and with `hideCreateForm` now `true`
          unconditionally at both mount points (the standalone page and the guided flow), that
          picker never renders anywhere, so the banner had gone from a real instruction to a
          reference to nothing. Removed rather than reworded, since the client's request was to
          remove it, not to describe the list-and-lookup screen this now is. */}

      {/* ── Create ─────────────────────────────────────────────────────────
          ONE of the two creation paths is visible at a time, chosen by the toggle the PAGE renders
          above — the same pattern the guided flow already uses for this step (client, 2026-08-23:
          "can we not follow the same design that is used in the Create New Segment"). The switch
          reaches this card through `hideCreateForm`, which is the flow's existing contract, so
          FolderAdmin needs no change and the Add-a-new-segment flow is untouched. */}
      {!hideCreateForm && (
        <div style={s.card}>
          <p style={s.head}>Create one group by hand — rarely needed</p>

          {/* THE NAME IS AN OUTPUT, NOT AN INPUT (client, 2026-08-18: *"client doesn't care about naming
            convention and they will simple type it in for no reason"*).

            The name is load-bearing: Folder Access recovers a group's role by PARSING ITS SUFFIX, so a
            hand-typed `GHO_GF_TAX_HC_UPLOADER` — suffix in the middle — parses as a PLAIN uploader and
            pre-selects the wrong role, which an admin then accepts because it looks right. Deriving it
            from the persona removes the guess.

            Read-only rather than hidden: the admin must be able to SEE what will be created, and a
            disabled-looking field with the value in it says "this is decided" better than no field. */}
          <label style={s.label} htmlFor="gm-name">
            Group name
          </label>
          <input
            id="gm-name"
            style={
              advanced
                ? s.input
                : { ...s.input, background: "#f3f2f1", color: "#323130" }
            }
            value={newName}
            readOnly={!advanced}
            placeholder={
              advanced
                ? "e.g. CRS_SITE_MEMBERS"
                : "Choose a segment, tier and persona below"
            }
            onChange={(e) => {
              setNewName(e.target.value);
              setNameTouched(true);
            }}
          />
          {newName.trim() !== "" && nameErrors.length > 0 && (
            <p style={s.err}>{nameErrors.join(" ")}</p>
          )}

          {/* The escape hatch, and it cannot be removed: `CRS_SITE_MEMBERS` follows no convention, and there
            will always be a one-off. But it is opt-in, which is the opposite of the old default. */}
          <label
            style={{
              ...s.hint,
              display: "flex",
              alignItems: "center",
              gap: 6,
              cursor: "pointer",
            }}
          >
            <input
              type="checkbox"
              checked={advanced}
              onChange={(e) => {
                setAdvanced(e.target.checked);
                setNameTouched(false);
              }}
            />
            Type the name myself (advanced) — for one-offs like the site-entry
            group
          </label>
          <p style={s.hint}>
            Nothing is granted here. Once the group exists, reconciliation
            applies its folder access.
          </p>

          {/* Always shown — this IS the form now, not an optional helper beside a text box. */}
          {!advanced && (
            <div
              style={{
                paddingLeft: 12,
                borderLeft: "2px solid #e1e1e1",
                marginTop: 10,
              }}
            >
              <label style={s.label} htmlFor="gm-seg">
                Segment
              </label>
              <select
                id="gm-seg"
                style={s.select}
                value={mode?.termSetGuid ?? ""}
                onChange={(e) => {
                  pickMode(e.target.value).catch(() => undefined);
                }}
              >
                <option value="">— select a segment —</option>
                {modes.map((m) => (
                  <option key={m.termSetGuid} value={m.termSetGuid}>
                    {m.label}
                  </option>
                ))}
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
                    onChange={(e) => {
                      pickTerm(i, e.target.value).catch(() => undefined);
                    }}
                  >
                    <option value="">— select —</option>
                    {options.map((t) => (
                      <option key={t.id} value={t.id}>
                        {t.label}
                      </option>
                    ))}
                  </select>
                </div>
              ))}

              {/* PERSONA, not a raw role (client, 2026-08-18). The role list offered things like "Delete
                pending files" and "View only — whole department", which are combinations nobody should
                have to assemble; Folder Access dropped its role chips for this reason and this screen was
                the same trap in a second place. The persona also decides the NAME, through its declared
                `namingRole` — so name and capability can no longer disagree. */}
              <label style={s.label} htmlFor="gm-persona">
                Persona
              </label>
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
                      <option key={p.key} value={p.key}>
                        {p.label}
                      </option>
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
                    {(
                      PERSONAS.filter((x) => x.key === builderPersona)[0]
                        ?.roles ?? []
                    )
                      .map((r) => roleLabel(r))
                      .join(" · ")}
                  </strong>
                  . Mapped at <strong>{personaScope(builderPersona)}</strong>{" "}
                  level.
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
                    {chosen
                      .filter((t) => !codes[t.id.toLowerCase()])
                      .map((t) => t.label)
                      .join(", ")}
                  </strong>
                  , so the name uses the full term name instead. Give it a code
                  on <strong>CRS Term Abbreviations</strong> first —
                  reconciliation names the folder from that code, and a group
                  named after the term will not line up with it.
                </p>
              )}
            </div>
          )}

          <label style={s.label} htmlFor="gm-people">
            Members (optional)
          </label>
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
                      if (!staged.some((x) => x.loginName === p.loginName))
                        setStaged([...staged, p]);
                      setPeopleQuery("");
                      setPeopleResults([]);
                    }}
                  >
                    {personDisplay(p.displayName, p.email)}
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
                    style={{
                      border: "none",
                      background: "none",
                      cursor: "pointer",
                      color: "#a4262c",
                    }}
                    onClick={() =>
                      setStaged(
                        staged.filter((x) => x.loginName !== p.loginName),
                      )
                    }
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
              onClick={() => {
                onCreate().catch(() => undefined);
              }}
            >
              {busy ? "Working…" : "Create group"}
            </button>
          </div>
        </div>
      )}

      {/* ── What can this person reach? ─────────────────────────────────
          Placed ABOVE the group list on purpose: it answers a question about a PERSON, and the list
          below answers one about the site. An admin arrives here because somebody cannot get in. */}
      <div style={s.card}>
        {/* "Quick Search" / shorter hint (client's mockup, 2026-09-03) — was "What can this person
            reach?" with a longer explanation. The screen still does exactly what it did; only the
            heading and the one-line description are shorter. */}
        <p style={s.head}>Quick Search</p>
        <p style={s.hint}>
          Search anyone in the directory to see which groups they are in on this
          site.
        </p>

        <div style={s.ddwrap}>
          {/* Clear empties the BOX AND THE ANSWER (client, 2026-08-23: *"I cannot clear once I
              entered, would be confusing UX design for client, put a clear button"*). Selecting a
              person clears the box on purpose — the result below is the answer — but that left no
              way to get back to an empty panel, so the last person looked up stayed on screen for
              the rest of the session. Rendered whenever there is either a query or a result. */}
          <div style={s.lkSearchRow}>
            <input
              style={s.input}
              value={lookupQuery}
              placeholder="Search by name or email…"
              onChange={(e) => setLookupQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Escape") clearLookup();
              }}
            />
            {(lookupQuery.length > 0 || lookupPerson !== undefined) && (
              <button type="button" style={s.ghost} onClick={clearLookup}>
                Clear
              </button>
            )}
          </div>
          {lookupResults.length > 0 && (
            <div style={s.dd}>
              {lookupResults.map((p) => (
                <button
                  key={p.loginName}
                  type="button"
                  style={s.ddItem}
                  onClick={() => {
                    runLookup(p).catch(() => undefined);
                  }}
                >
                  {personDisplay(p.displayName, p.email)}
                </button>
              ))}
            </div>
          )}
        </div>

        {lookupPerson && (
          <div style={{ marginTop: 12 }}>
            {/* THE EMAIL ALONE (client, 2026-08-23: "I think just showing the email is good enough").
                `name · email` read as two different people when the display name is a mangled
                directory form of the same address. The display name is the fallback rather than a
                blank line — a guest account can legitimately have no email. */}
            <p style={s.lkWho}>
              {personDisplay(lookupPerson.displayName, lookupPerson.email)}
            </p>

            {lookupBusy && <p style={s.hint}>Reading their groups&hellip;</p>}

            {/* The note carries BOTH kinds of message — a real answer ("not a member yet") and a
                failed read. They are told apart by whether `lookupGroups` is undefined, which is why
                that state is three-way. */}
            {!lookupBusy && lookupNote !== undefined && (
              <p style={lookupGroups === undefined ? s.err : s.hint}>
                {lookupNote}
              </p>
            )}

            {!lookupBusy &&
              lookupGroups !== undefined &&
              lookupGroups.length > 0 &&
              (() => {
                if (mapRows === undefined) {
                  // The Group Map is what turns a group into an ANSWER. Without it we can list the
                  // groups honestly but must not imply we know what they grant.
                  return (
                    <>
                      <p style={s.err}>
                        The Group Map could not be read, so what these groups
                        GRANT is not known. The groups themselves are listed
                        below.
                      </p>
                      {lookupGroups.map((g) => (
                        <p key={g.id} style={s.row}>
                          <span style={s.chip}>{g.title}</span>
                        </p>
                      ))}
                    </>
                  );
                }
                const summary = summarizeUserAccess(
                  lookupGroups,
                  mapRows.map((r) => ({
                    groupId: Number(r.groupId),
                    segment: r.segment,
                    unitTermGuid: r.tier,
                    role: r.role,
                  })),
                  {
                    segmentLabel: (guid) =>
                      modes.filter(
                        (m) =>
                          m.termSetGuid.toLowerCase() === guid.toLowerCase(),
                      )[0]?.label ?? "",
                    // v1 does not resolve the tier chain: the GROUP NAME already carries it in the
                    // same abbreviations the folders use (GHO_GF_TAX_UPLOADER is GHO / GF / TAX), and
                    // walking each segment's tree costs ~8 requests per segment for a second copy of
                    // what the title already says. `summarizeUserAccess` still returns the field, so
                    // rendering it later needs no change here.
                    tierChain: () => [],
                  },
                );
                return (
                  <>
                    <p style={s.hint}>
                      In {summary.groups.length} group
                      {summary.groups.length === 1 ? "" : "s"} on this site
                      {summary.unmappedCount > 0
                        ? ` · ${summary.unmappedCount} of them grant nothing`
                        : ""}
                    </p>
                    {summary.groups.map((g) => (
                      <div key={g.groupId} style={s.lkRow}>
                        {/* Name left, segment right — not a column of its own. */}
                        <div style={s.lkTop}>
                          <p style={s.lkName}>{g.groupTitle}</p>
                          {g.places.length > 0 && (
                            <p style={s.lkSeg}>
                              {g.places
                                .map(
                                  (pl) =>
                                    pl.segmentLabel || "segment not known",
                                )
                                .filter((v, i, a) => a.indexOf(v) === i)
                                .join(", ")}
                            </p>
                          )}
                        </div>

                        {/* The site-entry group is a NORMAL state, so it is a hint, not the red
                          "grants nothing" — it was rendering as an error because it has no rows,
                          which reads as something being wrong with the person's access. */}
                        {g.siteEntry && (
                          <p style={s.lkSum}>{describeGroupAccess(g)}</p>
                        )}
                        {!g.siteEntry && g.unmapped && (
                          <p style={s.err}>{describeGroupAccess(g)}</p>
                        )}
                        {!g.siteEntry && !g.unmapped && g.persona && (
                          <p style={s.lkPersona}>{g.persona.label}</p>
                        )}
                        {/* ⚠ NEVER `describeGroupAccess` here. With no persona it returns the joined
                          role labels — exactly what the chips below already say, which is what made
                          these rows look like they were repeating themselves. Say instead why there
                          is no summary: an unmatched role set is a real finding (a group holding
                          ENTRY alongside an approver's roles matches no persona exactly). */}
                        {!g.siteEntry && !g.unmapped && !g.persona && (
                          <p style={s.lkSum}>
                            These roles match no persona exactly — the group was
                            mapped by hand, or carries an extra role.
                          </p>
                        )}
                        {/* ⚠ `g.persona.summary` (the long role explanation, e.g. "Uploads to their
                          unit at any confidentiality level...") REMOVED (client's mockup, 2026-09-03)
                          — the short persona LABEL above (`g.persona.label`, "Upload only") and the
                          role chips below already say what this row grants; the paragraph was the
                          same fact a third time. */}

                        {g.roleLabels.length > 0 && (
                          <p style={s.lkChips}>
                            {g.roleLabels.map((r) => (
                              <span key={r} style={s.chip}>
                                {r}
                              </span>
                            ))}
                          </p>
                        )}
                      </div>
                    ))}
                    {/* ⚠ "Roles and scope come from the CRS Group Map..." REMOVED (client's mockup,
                      2026-09-03). The fact is unchanged — this screen still reads live SharePoint
                      state alongside Group Map rows exactly as before — only the explanatory footer
                      sentence is gone. */}
                  </>
                );
              })()}
          </div>
        )}
      </div>

      {/* ── The groups ─────────────────────────────────────────────────── */}
      <div style={s.card}>
        <p style={s.head}>
          Groups on this site ({visible.length}
          {visible.length === groups.length ? "" : ` of ${groups.length}`})
        </p>
        <div
          style={{
            display: "flex",
            gap: 8,
            flexWrap: "wrap",
            alignItems: "center",
          }}
        >
          <input
            style={{ ...s.input, flex: "1 1 220px" }}
            value={filter}
            placeholder="Filter by name…"
            onChange={(e) => setFilter(e.target.value)}
          />
          {/* Segment comes from the MAPPING ROWS, so it agrees with what reconciliation reads rather
              than with the name. Offered only when the Group Map was readable — a filter built from
              a failed read would silently claim every group belongs to no segment. */}
          {mapRows !== undefined && (
            <select
              style={{ ...s.select, flex: "0 1 220px" }}
              value={segFilter}
              onChange={(e) => setSegFilter(e.target.value)}
            >
              <option value="">Every segment</option>
              {modes.map((m) => (
                <option key={m.termSetGuid} value={m.termSetGuid.toLowerCase()}>
                  {m.label}
                </option>
              ))}
              <option value="__none__">Not mapped to anything</option>
            </select>
          )}
          {/* "What is shown", not "everything" — the same promise the audit log makes, and the only
              one that stays true next to a filter. */}
          <button
            type="button"
            style={s.ghost}
            disabled={visible.length === 0}
            onClick={exportShown}
          >
            Export what is shown
          </button>
        </div>

        {/* Selection bar. Shown only once something is ticked, so the ordinary "I came here to look
            at one group" visit is unchanged. `Select all shown` respects BOTH filters, which is the
            whole point: filter to APPROVER, tick all, delete. */}
        <div
          style={{
            display: "flex",
            gap: 8,
            flexWrap: "wrap",
            alignItems: "center",
            marginTop: 8,
          }}
        >
          <button
            type="button"
            style={s.ghost}
            disabled={visible.length === 0 || busy}
            onClick={() => {
              const next = { ...selected };
              const allOn = visible.every((g) => selected[g.id]);
              visible.forEach((g) => {
                if (allOn) {
                  delete next[g.id];
                } else {
                  next[g.id] = true;
                }
              });
              setSelected(next);
            }}
          >
            {visible.length > 0 && visible.every((g) => selected[g.id])
              ? `Clear the ${visible.length} shown`
              : `Select all ${visible.length} shown`}
          </button>
          {/* Page Access is the one report left (2026-09-02) — Site Access and Approval Library
              Access are retired signposts now, so linking to them here would only lead in a circle.
              Hidden when the page could not be resolved, rather than a dead link. */}
          {pageAccessHref && (
            <a href={pageAccessHref} style={s.ghost}>
              Page Access →
            </a>
          )}
          {selectedCount > 0 && (
            <>
              <span style={{ fontSize: 12, color: "#444" }}>
                <strong>{selectedCount}</strong> selected
                {/* A count that ignored the out-of-view ones would delete more than the screen
                    shows — the one number an admin must not be surprised by. */}
                {hiddenSelected > 0
                  ? ` (${hiddenSelected} not currently shown)`
                  : ""}
              </span>
              <button
                type="button"
                style={s.ghost}
                disabled={busy}
                onClick={() => setSelected({})}
              >
                Clear selection
              </button>
              <button
                type="button"
                style={s.danger}
                disabled={busy}
                onClick={() => {
                  setBulkTyped("");
                  setBulkOpen(true);
                }}
              >
                Delete {selectedCount} selected
              </button>
            </>
          )}
        </div>

        {/* ── System administrators ──────────────────────────────────────────
            Client, 2026-08-27: *"we did not include a group as the System Admin group? they should
            have all the power."* There was never a gap in the MODEL - SharePoint's site Owners group
            has always conferred exactly that, and reconciliation's lockdown pass grants the admin
            pages to Owners and nobody else. What was missing was a ROUTE to it here: `loadGroups`
            excludes the built-ins by id, so an admin had to leave for Site Permissions to add anyone.

            WARN: SEPARATED FROM THE LIST BELOW, NOT MERGED INTO IT. Owners must never appear among
            the mappable groups - mapping it onto a unit folder is exactly what the id exclusion
            exists to prevent - and it must never be offered a Delete button. Keeping it in its own
            card makes both true by construction rather than by remembering a flag on every row. */}
        <div
          style={{
            ...s.card,
            borderColor: "#a4262c",
            marginTop: 16,
            border: "1px solid #f1b0b3",
            background: "#fdf3f4",
          }}
        >
          <h1 style={{ ...s.head, marginTop: 0 }}>System administrators</h1>
          {owners === undefined ? (
            <p style={s.hint}>
              The site&apos;s owners group could not be read, so administrators
              cannot be managed here. Add them from Site settings → Site
              permissions instead.
            </p>
          ) : (
            <>
              {/* ⚠ COPY REPLACED VERBATIM (client's mockup, 2026-09-03: "Replace all"). The FACTS
                  are unchanged — {owners.title} still holds full control of every segment, every
                  library including Highly Confidential, and every admin page, unrestricted by any
                  folder permission and never removed by reconciliation — the client's shorter
                  wording just states the same thing without re-deriving the reasoning on screen. */}
              <p style={s.hint}>
                <strong>{owners.title}</strong> have full administrator access.
                Add users here only if they need to manage the entire system.
              </p>
              {/* SHOWN ONLY WHEN IT APPLIES. The first site collection administrator can only be set
                  by hand, because SharePoint lets only an existing one promote anybody — so an owner
                  who is not one would add people here and watch every promotion fail with an error
                  they could not act on. `undefined` (read failed) shows nothing, for the reason on
                  `viewerIsSca`. */}
              {viewerIsSca === false && (
                <p style={{ ...s.hint, color: "#8a4b00" }}>
                  <strong>You are not a site collection administrator.</strong>{" "}
                  Anyone you add here will join the group and get the CRS
                  access, but SharePoint will not let you promote them — only an
                  existing site collection administrator can do that. Set the
                  first one by hand at{" "}
                  <a
                    style={s.link}
                    href={`${siteUrl}/_layouts/15/mngsiteadmin.aspx`}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    Site collection administrators
                  </a>
                  . After that this card can do the rest.
                </p>
              )}
              {/* ⚠ COPY REPLACED VERBATIM (client's mockup, 2026-09-03: "Replace all"). Same three
                  facts the removed comment explained — a per-user SCA flag no group confers, no
                  report if it changes outside this card, a literal (never resolved) `_layouts` link
                  — stated now in the client's own words rather than this file's usual reasoning. */}
              <p style={s.hint}>
                <strong>Warning:</strong> Adding a user also makes them a{" "}
                <strong>Site Collection Administrator</strong>, giving them
                access to Site Settings, Term Store, group memberships, and site
                deletion.
              </p>
              <p style={s.hint}>
                The following cases must be managed manually through{" "}
                <a
                  style={s.link}
                  href={`${siteUrl}/_layouts/15/mngsiteadmin.aspx`}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Site Collection Administrators
                </a>
                :
                <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                  <li>
                    <strong>Your own account</strong> — removing yourself may
                    prevent this page from working correctly.
                  </li>
                  <li>
                    <strong>The only remaining administrator</strong> — a site
                    must always have at least one administrator.
                  </li>
                  <li>
                    <strong>Administrators removed outside this page</strong> —
                    for example, through SharePoint&apos;s Site Permissions
                    page. These changes must be removed manually.
                  </li>
                </ul>
              </p>
              <button
                type="button"
                style={s.ghost}
                onClick={() => setOwnersOpen(!ownersOpen)}
              >
                {ownersOpen ? "Hide administrators" : "Manage administrators"}
              </button>
              {ownersOpen && (
                <GroupMembersEditor
                  context={context}
                  siteUrl={siteUrl}
                  group={{ id: owners.id, title: owners.title }}
                  showToast={showToast}
                  /* WARN: THE ONLY MOUNT THAT MAY PASS THIS. Client, 2026-08-27 - their System Admin
                     means "access to everything", so this group carries SITE COLLECTION ADMINISTRATOR
                     with it. The same component is mounted for every OTHER group in the list below and
                     again on Folder Access; passing it there would promote a unit's uploader to site
                     collection administrator. */
                  alsoSiteAdmin={true}
                />
              )}
            </>
          )}
        </div>

        {/* STANDARDISE NAMES (client, 2026-08-27). 1.0.261.0 changed what NEW names are written and
            left existing groups alone, so a provisioned site is MIXED - `_APR_HIGHLY_CONFIDENTIAL`
            beside `_VIEWER_HIGHLY_CONFIDENTIAL`. Shown only when there is something to do, so it
            disappears once the site is consistent rather than sitting there inviting a pointless run. */}
        {canManage === true &&
          renameTargets().length > 0 &&
          renameProgress === undefined && (
            <div style={{ ...s.card, borderColor: "#8a4b00", marginTop: 16 }}>
              <p style={{ ...s.head, marginTop: 0 }}>
                Group names are not consistent
              </p>
              <p style={s.hint}>
                <strong>{renameTargets().length}</strong> group
                {renameTargets().length === 1 ? " uses" : "s use"} an older
                spelling (<code>_APR_…</code>, <code>_UPL_…</code>,{" "}
                <code>_EMPLOYEE</code>). Renaming them to <code>_APPROVER</code>
                , <code>_UPLOADER</code> and <code>_VIEWER</code> changes{" "}
                <strong>nothing about access</strong>: a rename keeps the
                group&apos;s id, so every mapping, folder permission and page
                grant stays exactly as it is.
              </p>
              <button
                type="button"
                style={s.btn}
                disabled={busy}
                onClick={() => {
                  const n = renameTargets().length;
                  if (
                    !window.confirm(
                      `Rename ${n} group${n === 1 ? "" : "s"} to the standard suffixes? ` +
                        `Nobody gains or loses access — a rename keeps the group id, so mappings and ` +
                        `folder permissions are untouched. Do not close this tab while it runs.`,
                    )
                  )
                    return;
                  runStandardiseNames().catch(() => undefined);
                }}
              >
                Rename {renameTargets().length} group
                {renameTargets().length === 1 ? "" : "s"}
              </button>
            </div>
          )}
        {renameProgress !== undefined && (
          <p style={s.hint}>
            Renaming {renameProgress.done} of {renameProgress.total}…
            {renameProgress.failed > 0
              ? ` (${renameProgress.failed} failed)`
              : ""}
          </p>
        )}

        {loadError !== undefined && (
          <p style={s.err}>
            Could not read this site&apos;s groups — {loadError}
          </p>
        )}
        {mapRows === undefined && !loading && loadError === undefined && (
          // Unreadable ≠ empty. Without this the badges below would silently claim every group is
          // unmapped, which reads as data loss.
          <p style={s.hint}>
            The {GROUP_MAP_LIST()} list could not be read, so this page cannot
            show which groups are mapped. Creating and deleting still work.
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
            // ⚠ The cap is LIFTED while a group is open, and that is not cosmetic: the member
            // editor's people picker is absolutely positioned, so a scroll container clips its
            // results for any group near the bottom — trading a long page for a control that
            // silently cannot be used. Same rule Folder Access follows.
            maxHeight: openGroup === undefined ? "60vh" : undefined,
            overflowY: openGroup === undefined ? "auto" : undefined,
          }}
        >
          {loading && <p style={s.hint}>Loading…</p>}
          {!loading && visible.length === 0 && (
            <p style={s.hint}>
              {groups.length === 0
                ? "No groups on this site yet."
                : "No group matches that filter."}
            </p>
          )}
          {visible.map((g) => {
            const rows = rowsFor(g);
            return (
              <div key={g.id}>
                <div style={s.row}>
                  {/* Ticking is the ONLY thing on this row that does not open or change anything —
                      it just marks. The destructive step is behind the bar above, with a count. */}
                  <input
                    type="checkbox"
                    aria-label={`Select ${g.title}`}
                    checked={selected[g.id] === true}
                    disabled={busy}
                    onChange={(e) => {
                      const next = { ...selected };
                      if (e.target.checked) {
                        next[g.id] = true;
                      } else {
                        delete next[g.id];
                      }
                      setSelected(next);
                    }}
                  />
                  {/* AN EXPANDER AGAIN (2026-08-23). Opening it is the whole membership feature:
                      one group at a time, so 586 groups cost 586 rows and ONE member read. */}
                  <button
                    type="button"
                    style={s.groupName}
                    aria-expanded={openGroup === g.id}
                    onClick={() =>
                      setOpenGroup(openGroup === g.id ? undefined : g.id)
                    }
                  >
                    {openGroup === g.id ? "▾" : "▸"} {g.title}
                  </button>
                  {/* PEOPLE, not mappings (client, 2026-08-23: *"instead of showing mappings only,
                      show how many users are there in each group, client doesnt understand what is
                      mapping"*). A mapping count is an internal fact about the Group Map; the
                      question an admin actually arrives with is who is in the group.

                      ⚠ An UNREADABLE member index falls back to the mapping badge rather than
                      showing `0 people`, which would report every group on the site as empty. */}
                  {memberIndex !== undefined &&
                    (() => {
                      const n = (memberIndex[g.id] ?? []).length;
                      // "No member" (client's mockup, 2026-09-03), was "nobody in it yet".
                      return n === 0 ? (
                        <span style={s.badge}>No member</span>
                      ) : (
                        <span style={s.mapped}>
                          {n} {n === 1 ? "person" : "people"}
                        </span>
                      );
                    })()}
                  {/* Kept, but only in the state that MATTERS and in plain words: a group with no
                      rows grants nothing, whoever is in it. The count is dropped from the happy
                      path — it is in the CSV for anyone who needs the number. */}
                  {rows !== undefined && rows.length === 0 && (
                    <span style={s.badge}>grants nothing yet</span>
                  )}
                  {memberIndex === undefined &&
                    rows !== undefined &&
                    rows.length > 0 && (
                      <span style={s.mapped}>
                        {rows.length} mapping{rows.length === 1 ? "" : "s"}
                      </span>
                    )}
                  <button
                    type="button"
                    style={s.danger}
                    disabled={busy}
                    onClick={() => setDeleting(g)}
                  >
                    Delete
                  </button>
                </div>

                {openGroup === g.id && (
                  <GroupMembersEditor
                    context={context}
                    siteUrl={siteUrl}
                    group={{ id: g.id, title: g.title }}
                    showToast={showToast}
                    onChanged={reloadMembers}
                  />
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
            <p style={{ ...s.head, fontSize: 15 }}>
              Delete &quot;{deleting.title}&quot;?
            </p>

            {(() => {
              const rows = rowsFor(deleting);
              if (rows === undefined) {
                // Unreadable ≠ "no mappings". Deleting is still allowed — the group may genuinely
                // need to go — but the admin is told the list could not be checked.
                return (
                  <p style={s.warnBox}>
                    The {GROUP_MAP_LIST()} list could not be read, so it is not
                    known whether this group holds folder mappings. Any rows it
                    has will be left behind, and reconciliation will report
                    them.
                  </p>
                );
              }
              if (rows.length === 0) {
                return (
                  <p style={{ fontSize: 13 }}>
                    This group holds no folder mappings. Its members lose it
                    immediately.
                  </p>
                );
              }
              return (
                <>
                  <p style={{ fontSize: 13 }}>
                    This group holds <strong>{rows.length}</strong> folder
                    mapping
                    {rows.length === 1 ? "" : "s"}, which will be deleted with
                    it:
                  </p>
                  <ul style={{ fontSize: 12, color: "#444", lineHeight: 1.6 }}>
                    {rows.map((r) => (
                      <li key={r.itemId}>
                        {roleLabel(r.role)} —{" "}
                        <span
                          style={{
                            fontFamily: "Consolas, monospace",
                            fontSize: 11,
                          }}
                        >
                          {r.tier || r.segment || "(no tier)"}
                        </span>
                      </li>
                    ))}
                  </ul>
                  {/* CORRECTED 2026-08-25. This said the folder permissions SURVIVE until
                      reconciliation runs, which is true of deleting a mapping ROW and false of
                      deleting the GROUP: SharePoint drops a deleted principal's role assignments
                      with the principal. The distinction stopped being academic the day an admin
                      deleted ~130 groups specifically to withdraw an obsolete HC grant — this text
                      would have told them it had not worked. */}
                  <p style={s.warnBox}>
                    Deleting this group removes its mappings AND{" "}
                    <strong>
                      withdraws the folder permissions it held, immediately
                    </strong>{" "}
                    — a deleted group keeps no access anywhere. To give it
                    access again, re-create it with Bulk provisioning and run
                    Folder Reconciliation.
                  </p>
                </>
              );
            })()}

            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button
                type="button"
                style={s.danger}
                disabled={busy}
                onClick={() => {
                  onDeleteGroup(deleting).catch(() => undefined);
                }}
              >
                {busy ? "Deleting…" : "Delete group"}
              </button>
              <button
                type="button"
                style={s.ghost}
                onClick={() => setDeleting(undefined)}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Bulk delete confirmation */}
      {bulkOpen && (
        <div
          style={s.modalBg}
          onClick={() => {
            if (!busy) setBulkOpen(false);
          }}
        >
          <div style={s.modal} onClick={(e) => e.stopPropagation()}>
            <p style={{ ...s.head, fontSize: 15 }}>
              Delete {selectedCount} group(s)?
            </p>

            {/* THE COST IS MEMBERSHIP, and it is the only irreversible part. Groups with nobody in
                them cost nothing to delete and re-create; groups with people lose those people, and
                there is no undo. So the ones that WILL cost something are named, not counted. */}
            {peopleAtRisk.length > 0 ? (
              <p style={s.warnBox}>
                <strong>{peopleAtRisk.length}</strong> of these still have
                people in them. Deleting a group removes its members permanently
                — export the list first if you have not.
                <br />
                <span style={{ fontSize: 12 }}>
                  {peopleAtRisk.slice(0, 8).join(", ")}
                  {peopleAtRisk.length > 8
                    ? ` and ${peopleAtRisk.length - 8} more`
                    : ""}
                </span>
              </p>
            ) : (
              <p style={{ fontSize: 13 }}>
                None of them have any members, so nothing is lost that Bulk
                provisioning cannot re-create.
              </p>
            )}

            <p style={s.warnBox}>
              Deleting a group <strong>does</strong> withdraw its folder
              permissions immediately — SharePoint drops a deleted
              principal&apos;s role assignments. That is the point of doing
              this, and it is also why reconciliation alone cannot undo it.
              Afterwards: <strong>Bulk provisioning</strong> to re-create them,
              then <strong>Folder Reconciliation</strong> to grant them again.
            </p>

            <p style={{ fontSize: 13, marginBottom: 4 }}>
              Type <strong>DELETE</strong> to confirm:
            </p>
            <input
              style={s.input}
              value={bulkTyped}
              disabled={busy}
              placeholder="DELETE"
              onChange={(e) => setBulkTyped(e.target.value)}
            />

            {bulkProgress !== undefined && (
              <p style={s.hint}>
                Deleting {bulkProgress.done} of {bulkProgress.total}…
                {bulkProgress.failed > 0
                  ? ` ${bulkProgress.failed} failed so far.`
                  : ""}{" "}
                Leave this tab open — there is no resume.
              </p>
            )}

            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button
                type="button"
                style={s.danger}
                disabled={busy || bulkTyped.trim().toUpperCase() !== "DELETE"}
                onClick={() => {
                  runBulkDelete().catch(() => {
                    bulkRunning.current = false;
                    setBusy(false);
                  });
                }}
              >
                {busy ? (
                  <span
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 6,
                    }}
                  >
                    {/* An SVG with animateTransform rather than a CSS spin: these styles are inline
                        objects, and inline styles cannot carry @keyframes. No stylesheet needed. */}
                    <svg
                      width="12"
                      height="12"
                      viewBox="0 0 50 50"
                      aria-hidden="true"
                    >
                      <circle
                        cx="25"
                        cy="25"
                        r="20"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="6"
                        strokeLinecap="round"
                        strokeDasharray="80 40"
                      >
                        <animateTransform
                          attributeName="transform"
                          type="rotate"
                          from="0 25 25"
                          to="360 25 25"
                          dur="0.9s"
                          repeatCount="indefinite"
                        />
                      </circle>
                    </svg>
                    {bulkProgress
                      ? `Deleting ${bulkProgress.done} of ${bulkProgress.total}…`
                      : "Deleting…"}
                  </span>
                ) : (
                  `Delete ${selectedCount} group(s)`
                )}
              </button>
              <button
                type="button"
                style={s.ghost}
                disabled={busy}
                onClick={() => setBulkOpen(false)}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {toast !== undefined && (
        <div
          style={{
            ...s.toast,
            background: toast.error ? "#a4262c" : "#0f6c3f",
          }}
        >
          {toast.message}
        </div>
      )}
    </div>
  );
}
