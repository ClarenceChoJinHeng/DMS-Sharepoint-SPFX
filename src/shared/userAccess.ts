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

import { PERSONAS, Persona, roleLabel, isSiteEntryGroupTitle } from "./groupMapModel";
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
    });
  }

  return {
    groups: out,
    unmappedCount: out.filter((x) => x.unmapped).length,
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
  if (g.siteEntry) return "Opens the site. No document access on its own.";
  if (g.unmapped) return "Grants nothing yet — this group has no access rows.";
  if (g.persona) return g.persona.label;
  if (g.roleLabels.length > 0) return g.roleLabels.join(", ");
  return "No roles recorded on this group's rows.";
}
