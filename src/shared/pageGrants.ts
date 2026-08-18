// Which groups should be able to open which page — derived from the roles they already hold.
//
// Spec: docs/superpowers/specs/2026-08-19-derived-page-access-design.md
//
// ⚠ THE BUG THIS EXISTS FOR: reconciliation granted a page only to groups with a `Scope = Page`
// row, and NOTHING creates those rows. Bulk provisioning writes Folder-scope rows, the guided flows
// never mention page access, and Group Management does not offer it. So on a site with 308 groups
// and 790 mapping rows, EIGHT groups could open the upload form — the four units an administrator
// had unblocked by hand while diagnosing an AccessDenied. The run log said `8 mappings applied`,
// which was true and told nobody that 56 units were locked out of the form they exist to use.
//
// Fourth instance of one structural gap: the mechanism is driven by grant ROWS, and the thing that
// needs the grant has none. Fixed the same way as the other three — ASSERT the required state every
// run, rather than deriving it from the presence of a row.
//
// Pure and SPFx-free, because it decides who can open a restricted page and the failure is silent
// in both directions: too few and the people who need the page are denied it, too many and a
// restricted page is not restricted.

import { GroupMapRole, normalizeRoleValue } from "./groupMapModel";
import { derivedRolesForPage } from "./pageAccessPolicy";

/** A CRS Group Map row as read back from SharePoint. Every field optional — SharePoint's choice. */
export interface GroupMapReadRow {
  GroupId?: string;
  GroupName?: string;
  Role?: string;
  Scope?: string;
  Target?: string;
}

/** One group and every role it holds, collapsed across segments and tiers. */
export interface GroupRoles {
  /** The SP site group's integer id, as text. It IS the role-assignment principal id. */
  groupId: string;
  /** Last name seen for it. Display only — it can be stale, so nothing keys on it. */
  groupName: string;
  roles: GroupMapRole[];
}

/** A group that should hold Read on a page, and why. */
export interface IntendedPageGroup {
  groupId: string;
  groupName: string;
  /** `derived` from a folder role; `row` from a hand-made Scope=Page row. Shown in the log. */
  source: "derived" | "row";
  /** The role that earned it, for the log line. Empty for a hand-made row. */
  via: string;
}

/** One role assignment currently on the page. */
export interface CurrentAssignment {
  principalId: number;
  title: string;
  /** True for a SharePoint group. A user principal is never a removal candidate — see spec §4.2. */
  isGroup: boolean;
}

const isFolderScope = (raw: string | undefined): boolean => {
  const v = (raw ?? "").trim().toLowerCase();
  // Blank counts as Folder. The Scope column was added later, so early rows have none, and a row
  // with no scope has always been treated as a folder grant by the folder pass. Reading blank as
  // "not a folder" here would silently under-derive on exactly the oldest sites.
  return v === "" || v === "folder";
};

/**
 * Collapse Folder-scope rows into one entry per group.
 *
 * TIER IS DELIBERATELY IGNORED. A group holding `UPL` anywhere, at any tier, in any segment, needs
 * the upload form: the page is not scoped to a unit and cannot be. What the group may actually
 * FILE is decided by the folder ACL, which the upload form probes before offering a destination —
 * so a group listed here that cannot write anywhere sees an empty cascade rather than a page it
 * should not have had.
 *
 * Keyed on `GroupId`, never the name: two groups can be renamed alike, and a stored `GroupName`
 * goes stale the moment a group is renamed — which is deliberately allowed, because the Id survives
 * a rename and every mapping row with it.
 *
 * Role values go through `normalizeRoleValue`, so a long-form `Uploader` written by an older
 * version counts as `UPL`. That exact mismatch had already made `collectMembership` refuse a
 * correctly provisioned uploader once.
 */
export function groupRolesById(rows: GroupMapReadRow[]): GroupRoles[] {
  const out: GroupRoles[] = [];
  const index: Record<string, number> = {};
  for (const r of rows ?? []) {
    if (!isFolderScope(r.Scope)) continue;
    const id = (r.GroupId ?? "").trim();
    if (id.length === 0) continue;
    const role = normalizeRoleValue(r.Role ?? "").toUpperCase() as GroupMapRole;
    if ((role as string).length === 0) continue;
    const at = index[id];
    if (at === undefined) {
      index[id] = out.length;
      out.push({ groupId: id, groupName: (r.GroupName ?? "").trim(), roles: [role] });
    } else {
      const g = out[at];
      if (g.roles.indexOf(role) === -1) g.roles.push(role);
      // A later row's name wins only if the earlier one was blank — never overwrite a real name
      // with an empty cell, which some rows carry.
      if (g.groupName.length === 0) g.groupName = (r.GroupName ?? "").trim();
    }
  }
  return out;
}

/**
 * The complete set of groups that should hold Read on one page.
 *
 * Derived groups, then hand-made `Scope = Page` rows merged on top. Rows are NOT replaced by
 * derivation: a Page row is the only way to grant a page to a group the policy does not imply, and
 * removing that capability silently would be a regression on a screen the client already uses.
 *
 * Deduped on `groupId`. A group that is both derived and hand-rowed appears once, as `derived` —
 * the row is then redundant rather than wrong, and saying `derived` explains why it would still
 * have the page if the row were deleted.
 *
 * Returns `[]` for a page that matched no rule, is administrator-only, or has no qualifying group.
 * The caller must distinguish the first two (leave the page ALONE) from the third (assert Owners
 * only, and warn) — `derivedRolesForPage(file).length` is what separates them.
 */
export function intendedPageGroups(
  fileName: string,
  groups: GroupRoles[],
  manualPageRows: GroupMapReadRow[],
): IntendedPageGroup[] {
  const roles = derivedRolesForPage(fileName);
  const out: IntendedPageGroup[] = [];
  const seen: Record<string, true> = {};
  if (roles.length > 0) {
    for (const g of groups ?? []) {
      const hit = g.roles.filter((r) => roles.indexOf(r) !== -1);
      if (hit.length === 0) continue;
      seen[g.groupId] = true;
      out.push({ groupId: g.groupId, groupName: g.groupName, source: "derived", via: hit.join(", ") });
    }
  }
  const target = (fileName ?? "").trim().toLowerCase();
  for (const r of manualPageRows ?? []) {
    if ((r.Target ?? "").trim().toLowerCase() !== target) continue;
    const id = (r.GroupId ?? "").trim();
    if (id.length === 0 || seen[id]) continue;
    seen[id] = true;
    out.push({ groupId: id, groupName: (r.GroupName ?? "").trim(), source: "row", via: "" });
  }
  return out;
}

/**
 * Which of a page's current assignments should come off.
 *
 * ⚠ SAFE ONLY AT PAGE SCOPE. A Site Pages item is a LEAF, so its role assignments contain only what
 * someone deliberately granted. A library or folder is not: SharePoint auto-creates a Limited
 * Access assignment at every parent scope for any principal holding a grant further down, so an
 * approval library's root carries an entry for every one of the ~308 groups with a folder grant
 * inside it. Applying this rule there would strip them all and every group would lose its folder
 * access, on a run that reported success. Do not reuse this function above a leaf.
 *
 * User principals are never candidates. Individual grants were explicitly rejected as a mechanism
 * in 2026-08-14-per-person-access-removal-design.md, and removing them here would quietly implement
 * the thing that spec declined; Page Access already flags one in red for a human to decide.
 *
 * `protectedIds` carries the web's associated groups and the site-entry group. Removing Owners is
 * how a page becomes unreachable by the people who administer it.
 */
export function groupsToRemove(
  intended: IntendedPageGroup[],
  current: CurrentAssignment[],
  protectedIds: number[],
): CurrentAssignment[] {
  const keep: Record<number, true> = {};
  for (const g of intended ?? []) {
    const n = Number((g.groupId ?? "").trim());
    if (n > 0 && n % 1 === 0) keep[n] = true;
  }
  for (const p of protectedIds ?? []) if (typeof p === "number" && p > 0) keep[p] = true;
  return (current ?? []).filter((a) => a.isGroup && !keep[a.principalId]);
}
