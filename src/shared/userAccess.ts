/**
 * "What can this person actually reach?" — answered from their group memberships.
 *
 * Client, 2026-08-23, on why the access pages feel redundant: *"if the group is already auto assign
 * then what is the point of the pages?"* They are right. Bulk provisioning creates the groups AND
 * writes their Group Map rows; reconciliation grants the folder ACLs from those rows. So Folder
 * Access, Site Access and Approval Library Access are mostly vestigial — the two jobs that remain
 * are putting a person in a group, and answering this question, which had no home anywhere.
 *
 * ⚠ THIS MODULE REPORTS PERSONAS AND PLACES, NEVER PERMISSION LEVELS. Also saying "CRS Upload on
 * Approval Document" was considered and rejected: that answer lives in `LIBRARY_ROLES` /
 * `ROLE_TO_PERMISSION` inside `FolderManager.tsx`, and `ROLE_TO_PERMISSION` is MUTATED at runtime by
 * `applyPermissionPrefix()` to re-point `DMS Upload` at the site's real prefix. A second copy here
 * would drift; importing the live one would make the answer depend on whether `FolderManager` had
 * mounted yet — the render-time `libApiTitle()` trap in a new place. An admin asking this question
 * wants "PIC in GHO / Group Finance / Tax", which is what it gives.
 *
 * SPFx-free and clock-free by design, like the rest of `shared/`: everything is decided from its
 * arguments, so it can be tested without a tenant.
 */

import {
  PERSONAS,
  Persona,
  roleLabel,
  isSiteEntryGroupTitle,
  normalizeRoleValue,
  suggestGroupName,
} from "./groupMapModel";
import { roleSetKey } from "./bulkGroups";

/** One SharePoint group the person belongs to. */
export interface UserGroupRef {
  id: number;
  title: string;
}

/**
 * One Group Map row, already read.
 *
 * `groupId` is the join key and the ONLY one — never the name. A group rename keeps its id, so a
 * stored `GroupName` goes stale (the 1.0.162.0 rename bug), and two groups can be renamed alike.
 */
export interface AccessRow {
  groupId: number;
  segment?: string;
  unitTermGuid?: string;
  role?: string;
  scope?: string;
  target?: string;
}

/** One distinct place a group is mapped at. */
export interface AccessPlace {
  /** Human segment name, or "" when it could not be resolved. */
  segmentLabel: string;
  /**
   * Tier labels from the segment down to the mapped term — ["Group Finance", "Tax"].
   *
   * EMPTY means NOT RESOLVED, and must be rendered as "not known" rather than as a term with no
   * parents. Reporting an unresolved chain as a bare leaf presents a UNIT as a DEPARTMENT, which is
   * the difference between granting one unit and granting all of them — the same reasoning behind
   * Folder Access's `Tier not known`.
   */
  tierChain: string[];
  scope: string;
  target: string;
}

/** What one of the person's groups gives them. */
export interface GroupAccess {
  groupId: number;
  groupTitle: string;
  /** Roles this group holds, deduped, upper-cased, in first-seen order. */
  roles: string[];
  /** Readable role names, same order. */
  roleLabels: string[];
  /** The persona whose role set this group matches EXACTLY, when one does. */
  persona?: Persona;
  places: AccessPlace[];
  /**
   * No Group Map rows at all. The group exists and grants NOTHING.
   *
   * A legitimate state since 2026-08-14, when creating a group stopped implying a mapping — so this
   * is a fact, not an error. It is also the single most useful thing this page can tell an admin who
   * is asking why somebody cannot get in.
   */
  unmapped: boolean;
  /**
   * The site-entry group (`CRS_SITE_MEMBERS`). It holds Read on the WEB and nothing in any library,
   * so it is what lets someone open the site at all — and it is NOT a folder grant. Flagged because
   * an admin seeing it listed would otherwise count it as access to documents.
   */
  siteEntry: boolean;
  /**
   * The site's OWNERS group.
   *
   * ⚠ IT HAS NO GROUP MAP ROWS AND IT NEVER WILL — SharePoint grants it **Full Control on the WEB**,
   * which is a site permission rather than a folder grant, so nothing in `CRS Group Map` describes
   * it. Without this flag it fell to `unmapped` and the lookup said *"Grants nothing yet — this
   * group has no access rows"*, **in red**, about the widest access anybody on the site can hold.
   * Reported by the client 2026-09-07.
   *
   * Same reasoning as `siteEntry` beside it: a group whose access comes from somewhere other than a
   * mapping row is a NORMAL state, not a finding, and must not be rendered as one.
   */
  owners: boolean;
}

/** Everything known about one person's access. */
export interface UserAccessSummary {
  groups: GroupAccess[];
  /** How many of those groups grant nothing. They are still in `groups`. */
  unmappedCount: number;
  /** True when the person is in no groups at all. */
  none: boolean;
}

function clean(v: string | undefined): string {
  return (v ?? "").trim();
}

function upper(v: string | undefined): string {
  return clean(v).toUpperCase();
}

/**
 * How to print a person: `Name · email`, or just the email when the "name" is only the address
 * again.
 *
 * Client, 2026-08-23: *"I notice the dropdown is showing the duplicated clarencechojinheng again"* —
 * a guest account's display name is very often the local part of its own address, so
 * `clarencechojinheng · clarencechojinheng@gmail.com` reads as two different people, and
 * `test@gmail.com test@gmail.com` reads as a rendering bug. Real names (`Wyanet Falasifa`
 * `<wyanet@trinergydigital.com>`) still show both, because there the name is the useful half.
 *
 * PURE and shared: the people picker, the member rows and the lookup header all print a person, and
 * three copies of this rule would drift — the first version of it lived only in the lookup header.
 */
export function personDisplay(displayName: string | undefined, email: string | undefined): string {
  const name = clean(displayName);
  const mail = clean(email);
  if (!mail) return name;
  if (!name) return mail;
  const lowerName = name.toLowerCase();
  const lowerMail = mail.toLowerCase();
  // The name IS the address, or the address's local part — nothing is added by printing both.
  if (lowerName === lowerMail) return mail;
  if (lowerMail.indexOf(`${lowerName}@`) === 0) return mail;
  return `${name} · ${mail}`;
}

/**
 * The persona whose role set matches these roles EXACTLY.
 *
 * ⚠ EXACT, never a subset — the rule `planBulkGroups` follows, for the same reason: `employee` is
 * `["MEMBER"]` and `employee_hc` is `["MEMBER","MEMBERHC"]`, so a subset match would let an
 * HC-cleared viewer group answer as the plain one. Presenting a CLEARED group as UNCLEARED is the
 * dangerous direction, so an unrecognised set answers `undefined` and the caller shows the raw roles.
 *
 * Reuses `roleSetKey` rather than re-fingerprinting: two definitions of "the same role set" is how
 * this page and the provisioner would come to disagree about what a group is.
 */
export function personaForRoles(roles: readonly string[]): Persona | undefined {
  const key = roleSetKey(roles as string[]);
  if (key.length === 0) return undefined;
  for (const p of PERSONAS) {
    if (roleSetKey(p.roles as unknown as string[]) === key) return p;
  }
  return undefined;
}

/** A CRS Group Map row as read back from SharePoint, already scoped to one segment. */
export interface GroupMapRenameRow {
  GroupId?: string;
  GroupName?: string;
  Role?: string;
  UnitTermGuid?: string;
}

/**
 * Which SharePoint groups now carry a STALE code in their own name — after a Department/Unit
 * abbreviation rename or a segment top-folder recode — and what they should be called instead.
 *
 * Spec: docs/superpowers/specs/2026-09-15-segment-rename-and-group-rename-design.md (Part 2)
 *
 * A rename NEVER changes what a group grants — Group Map rows are keyed on `GroupId`, which a
 * SharePoint rename preserves — so this is always optional and always safe to skip (see 1.0.261.0's
 * "safe because a rename preserves the group id"). It exists only because `suggestGroupName` derives
 * a group's NAME from the same codes the Abbreviations screen edits (`GHO_GCA_TAX_APPROVER` from the
 * segment code plus the term chain's codes), so a code rename leaves every group naming that term
 * pointing at a folder name that no longer exists.
 *
 * ⚠ `rows` MUST ALREADY BE SCOPED TO THE ONE SEGMENT BEING EDITED (the caller's own
 * `Segment eq '<guid>'` read) — this never widens that scope itself, so a GLOBAL group (blank
 * `Segment`, site-wide by definition) is never a candidate here, matching `suggestGroupName`'s own
 * special-casing of it.
 *
 * `codeChainFor` is caller-supplied, because building it needs the abbreviation rows the caller has
 * already loaded — see `codeChain` in `bulkGroups.ts`, reused verbatim by both callers of this
 * function rather than re-fingerprinted here. `undefined` means the chain could not be resolved (a
 * broken parent link), and a group on such a chain is SKIPPED rather than guessed at — a wrong
 * suggested rename is worse than none, since acting on it points the group at a name nobody chose.
 */
export function plannedGroupRenames(
  renamedTermGuids: readonly string[],
  segmentRecoded: boolean,
  segmentCode: string,
  rows: readonly GroupMapRenameRow[],
  codeChainFor: (termGuid: string) => string[] | undefined,
  /**
   * ⚠⚠ FIXED 2026-09-16 — a second real bug found live right after the first, on the very same GHO
   * recode: the segment's own C-Level group (`GHO_C_LEVEL`, role SEGVIEW) was ALSO silently skipped.
   *
   * `GroupMapWriteRow.UnitTermGuid` is documented (and, checked directly in `bulkGroups.ts`'s
   * `rowsFor`/`push`, actually written) as `"" unless Scope is Folder; equals Segment for a
   * segment-tier row` — so a SEGVIEW row's `UnitTermGuid` is NOT blank in production, it is the
   * SEGMENT's own term SET guid. `codeChainFor` only knows Department/Unit TERM guids walked from
   * inside that set (from `loadPermissionedAbbreviationRows`), never the set's own id — so without
   * being told which guid IS the segment, `codeChainFor(entry.unitGuid)` returns `undefined` for a
   * segment-tier row and it was dropped by the very "never guess without a resolvable chain" guard
   * that exists to protect ordinary unit/department rows. (A GLOBAL row is unaffected — its
   * `UnitTermGuid` is forced blank by `buildGroupMapRow`'s termless rule, and blank already resolved
   * to `chain = []` before this parameter existed.)
   *
   * Optional, defaulting to `""` — no supplied guid can ever equal a non-empty `entry.unitGuid`, so
   * every existing caller and test that never passed one keeps its old (blank ⇒ `[]`, anything else
   * ⇒ ask `codeChainFor`) behaviour exactly.
   */
  segmentTermSetGuid: string = "",
): Array<{ groupId: number; from: string; to: string }> {
  const renamed = new Set(renamedTermGuids.map((g) => (g ?? "").trim().toLowerCase()));

  // Group Map rows are one per (group, term, role), so a group holding six roles carries six rows —
  // all sharing the same GroupId, GroupName and UnitTermGuid. Collapsed to one entry per group
  // BEFORE any name is computed, or `suggestGroupName` would be asked for six different suffixes on
  // the same group and whichever role was seen last would silently win.
  const byGroup: Record<string, { name: string; unitGuid: string; roles: string[] }> = {};
  for (const row of rows) {
    const groupId = (row.GroupId ?? "").trim();
    if (!groupId) continue;
    const unitGuid = (row.UnitTermGuid ?? "").trim();
    // ⚠⚠ FIXED 2026-09-16 — a real bug, found live on GHO's own recode. The STEM every group name is
    // built from (`suggestGroupName(segmentCode, …)`) is the SEGMENT's code for every row, unit-tier
    // included — so recoding the segment goes stale for ALL of them, not only the segment-tier ones.
    // The line this replaces read `unitGuid ? renamed.has(unitGuid) : segmentRecoded` — which checked
    // `segmentRecoded` ONLY for a row with NO unit guid (SEGVIEW/GLOBAL), and for every ordinary
    // unit/department row asked SOLELY whether that row's own term had just been renamed, ignoring
    // `segmentRecoded` entirely. So a pure segment recode (no term renamed, `renamedTermGuids` empty)
    // offered to fix `GHO_SEGVIEW` and left every `GHO_<dept>_<unit>_<role>` group silently stale —
    // exactly what the caller then went and re-provisioned as a parallel `GHOS_*` set, because bulk
    // provisioning had no way to know the old names were meant to become these.
    // `segmentRecoded` now makes EVERY row a candidate, tier notwithstanding — a TERM rename still
    // narrows to just that term's own row via the second half. Neither direction widens who gets
    // RENAMED beyond who was already a candidate: `to === entry.name` still drops a row unchanged
    // two lines down, so this only ever adds rows whose computed name has actually gone stale.
    const affected = segmentRecoded || (unitGuid !== "" && renamed.has(unitGuid.toLowerCase()));
    if (!affected) continue;
    const entry = byGroup[groupId] ?? { name: "", unitGuid, roles: [] };
    if (!entry.name) entry.name = (row.GroupName ?? "").trim();
    const role = normalizeRoleValue(row.Role ?? "");
    if (role) entry.roles.push(role);
    byGroup[groupId] = entry;
  }

  const out: Array<{ groupId: number; from: string; to: string }> = [];
  // `Object.entries` is unavailable at this tsconfig's target (the same ES-level limitation as
  // `Promise.allSettled` — see CLAUDE.md gotcha #3), so `Object.keys` is the house pattern.
  for (const groupId of Object.keys(byGroup)) {
    const entry = byGroup[groupId];
    const id = Number(groupId);
    if (!Number.isFinite(id)) continue;
    // Matched EXACTLY, same rule `planBulkGroups`'s rename recovery already follows — a role set
    // that matches no persona cannot be confidently named, so it is skipped rather than guessed at.
    const persona = personaForRoles(entry.roles);
    if (!persona) continue;
    // A row whose `unitGuid` IS the segment's own term-set guid is the segment-tier (SEGVIEW) group
    // itself — it has no unit chain to resolve, exactly like a blank `unitGuid` (GLOBAL). Checked
    // BEFORE calling `codeChainFor`, which only knows terms walked from INSIDE the set and would
    // otherwise answer `undefined` for the set's own id and drop this row via the very next guard.
    const isSegmentTierGuid =
      entry.unitGuid !== "" &&
      segmentTermSetGuid !== "" &&
      entry.unitGuid.toLowerCase() === segmentTermSetGuid.trim().toLowerCase();
    const chain = entry.unitGuid === "" || isSegmentTierGuid ? [] : codeChainFor(entry.unitGuid);
    if (entry.unitGuid !== "" && !isSegmentTierGuid && chain === undefined) continue;
    const to = suggestGroupName(segmentCode, chain ?? [], persona.namingRole);
    if (!to || entry.name.toUpperCase() === to.toUpperCase()) continue;
    out.push({ groupId: id, from: entry.name, to });
  }
  return out.sort((a, b) => a.from.localeCompare(b.from));
}

/**
 * Fold a person's groups and the Group Map into one answer.
 *
 * `segmentLabel` and `tierChain` are supplied by the caller because resolving them needs the term
 * store and the config list — this module stays testable without either. Both may legitimately
 * answer "" / [], which is reported as "not known" rather than guessed.
 */
export function summarizeUserAccess(
  groups: readonly UserGroupRef[],
  rows: readonly AccessRow[],
  resolve: {
    segmentLabel: (termSetGuid: string) => string;
    tierChain: (termGuid: string) => string[];
  },
  /**
   * The site owners group's id, when the caller could resolve it.
   *
   * ⚠ AN ID, NEVER THE TITLE. This client renames everything at import — the owners group reads
   * "Guthrie Central Repository System Owners" on the live site — so a title match would silently
   * stop recognising it and the group would go back to reading "grants nothing" in red.
   * `fetchOwnersGroup` resolves it from `AssociatedOwnerGroup`, which a rename never touches.
   *
   * `undefined` when that read failed, and then no group is flagged — the group simply reports as
   * unmapped again, which is what it did before this existed. Understating is the safe direction:
   * flagging the WRONG group as owners would tell an admin a unit's uploader group holds full
   * control of the site.
   */
  ownerGroupId?: number,
): UserAccessSummary {
  const byGroup: Record<number, AccessRow[]> = {};
  for (const r of rows ?? []) {
    if (typeof r?.groupId !== "number") continue;
    if (!byGroup[r.groupId]) byGroup[r.groupId] = [];
    byGroup[r.groupId].push(r);
  }

  const out: GroupAccess[] = [];
  for (const g of groups ?? []) {
    const mine = byGroup[g.id] ?? [];

    const roles: string[] = [];
    for (const r of mine) {
      const role = upper(r.role);
      if (role && roles.indexOf(role) === -1) roles.push(role);
    }

    // One entry per distinct PLACE, not per row: a unit's approver group carries six role rows at
    // one term, and listing that term six times says nothing the first line did not.
    const places: AccessPlace[] = [];
    const seen: Record<string, true> = {};
    for (const r of mine) {
      const segment = clean(r.segment);
      const term = clean(r.unitTermGuid);
      const scope = clean(r.scope) || "Folder";
      const target = clean(r.target);
      const k = `${segment}|${term}|${scope}|${target}`;
      if (seen[k]) continue;
      seen[k] = true;
      places.push({
        segmentLabel: segment ? resolve.segmentLabel(segment) : "",
        tierChain: term ? resolve.tierChain(term) : [],
        scope,
        target,
      });
    }

    out.push({
      groupId: g.id,
      groupTitle: g.title,
      roles,
      roleLabels: roles.map((r) => roleLabel(r)),
      persona: personaForRoles(roles),
      places,
      unmapped: mine.length === 0,
      siteEntry: isSiteEntryGroupTitle(g.title),
      owners: typeof ownerGroupId === "number" && g.id === ownerGroupId,
    });
  }

  return {
    groups: out,
    /* ⚠ THE SITE-ENTRY AND OWNERS GROUPS ARE EXCLUDED, and leaving them in is how this line came to
       contradict the rows beneath it. Both legitimately hold no mapping rows — one grants Read on
       the web, the other Full Control — so counting them produced "3 groups on this site · 2 of them
       grant nothing" about a person who could open the site AND administer it. The count exists to
       answer "why can this person not get in", and neither of those two is ever the reason. */
    unmappedCount: out.filter((x) => x.unmapped && !x.siteEntry && !x.owners).length,
    none: out.length === 0,
  };
}

/**
 * A one-line description of what a group gives, for the collapsed row.
 *
 * Persona label when the role set is recognised, the roles themselves when it is not — never a
 * guess, and never blank: a row with no description reads as a failed load.
 */
export function describeGroupAccess(g: GroupAccess): string {
  /* ⚠ OWNERS IS TESTED FIRST, AHEAD OF `unmapped`, AND THE ORDER IS THE FIX. This group holds no
     mapping rows and never will — its Full Control comes from SharePoint at web scope — so the
     `unmapped` branch below would otherwise claim the widest access on the site grants nothing. */
  if (g.owners) return "Full control of this site, granted by SharePoint rather than by a mapping row.";
  if (g.siteEntry) return "Opens the site. No document access on its own.";
  if (g.unmapped) return "Grants nothing yet — this group has no access rows.";
  if (g.persona) return g.persona.label;
  if (g.roleLabels.length > 0) return g.roleLabels.join(", ");
  return "No roles recorded on this group's rows.";
}
