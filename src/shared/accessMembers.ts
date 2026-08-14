/**
 * Per-person access on the Approval Library Access and Page Access screens.
 *
 * Spec: docs/superpowers/specs/2026-08-14-per-person-access-removal-design.md
 *
 * Both screens grant access to GROUPS. A person has access because they are IN one of those
 * groups, so the only per-person lever is group membership — and that is a wider change than the
 * screen it sits on. This module holds the rules that make the wider change legible before it is
 * made, and nothing else: it is pure, SPFx-free and testable, because every one of these rules has
 * a plausible wrong version that reads as correct on screen.
 *
 * The three that matter:
 *   1. A person in TWO allowed groups keeps access after one removal. A per-row Remove with no
 *      cross-group view makes a correct removal look like a failed one.
 *   2. Removing the last member does NOT remove the grant. The group stays allowed and empty —
 *      which reads as configured.
 *   3. An unreadable member list is NOT an empty one. This distinction was advisory when the
 *      members column was decoration; it is load-bearing now that a removal decision rests on it.
 *
 * `SpGroupMember` comes from spGroupsFilter (the pure half), never from spGroups (the REST half):
 * importing the latter would pull @microsoft/sp-http in here and take this module's plain-Jest
 * tests with it.
 */

import { SpGroupMember } from "./spGroupsFilter";

/**
 * What is known about one group's members.
 *
 * A three-state union rather than `SpGroupMember[] | undefined`, because "not read yet" and "could
 * not be read" must not collapse into each other, and neither may be shown as "no members". The
 * empty-vs-unknown rule this codebase applies everywhere else (gotchas 10b/11, the provisioned
 * segment filters) applies here for a sharper reason: an admin told a group is empty stops looking
 * for the person they came to remove.
 */
export type MemberLoad =
  | { state: "loading" }
  | { state: "loaded"; members: SpGroupMember[] }
  | { state: "error"; message: string };

/** Members per group id, as the screens hold it. A missing key means loading has not started. */
export type MemberMap = Record<number, MemberLoad | undefined>;

export function loadOf(map: MemberMap, groupId: number): MemberLoad {
  return (map ?? {})[groupId] ?? { state: "loading" };
}

/** The members of a group, or [] when they are unknown. Never use this to decide emptiness. */
export function membersOf(map: MemberMap, groupId: number): SpGroupMember[] {
  const load = loadOf(map, groupId);
  return load.state === "loaded" ? load.members : [];
}

/**
 * Which of the ALLOWED groups each user belongs to.
 *
 * Keyed on the SharePoint user id, which is stable within a site — not on email, which is blank for
 * some principals, and not on title, which is not unique.
 *
 * Only allowed groups are counted. A person's membership of a group that grants nothing here is not
 * a reason to warn that access survives, and including it would put a warning on almost every row —
 * at which point the warning stops being read, and the one case that matters is lost with it.
 */
export function membershipIndex(
  allowedGroupIds: readonly number[],
  map: MemberMap,
): Map<number, number[]> {
  const out = new Map<number, number[]>();
  for (const gid of allowedGroupIds ?? []) {
    for (const m of membersOf(map, gid)) {
      const seen = out.get(m.id);
      if (seen) { if (seen.indexOf(gid) === -1) seen.push(gid); }
      else out.set(m.id, [gid]);
    }
  }
  return out;
}

/**
 * The OTHER allowed groups that would still let this person in.
 *
 * The answer to "did my removal work?". Empty means the removal genuinely ends their access at this
 * scope; non-empty means it does not, and saying so BEFORE the click is the whole point.
 *
 * Groups whose members could not be read cannot appear here, so this list is a floor and never a
 * ceiling — `unreadableGroups` is what stops that being read as a guarantee.
 */
export function otherAllowedGroups(
  userId: number,
  thisGroupId: number,
  index: Map<number, number[]>,
  titleOf: (groupId: number) => string,
): string[] {
  return (index.get(userId) ?? [])
    .filter((gid) => gid !== thisGroupId)
    .map((gid) => titleOf(gid))
    .filter((t) => t.trim().length > 0)
    .sort((a, b) => a.localeCompare(b));
}

/**
 * Allowed groups whose membership could not be read.
 *
 * Named in the removal dialog. Without it, "this ends their access" would be asserted from an
 * incomplete picture — and the incomplete half is invisible, because a failed read renders as a
 * collapsed row like any other.
 */
export function unreadableGroups(
  allowedGroupIds: readonly number[],
  map: MemberMap,
  titleOf: (groupId: number) => string,
): string[] {
  return (allowedGroupIds ?? [])
    .filter((gid) => loadOf(map, gid).state === "error")
    .map((gid) => titleOf(gid))
    .filter((t) => t.trim().length > 0)
    .sort((a, b) => a.localeCompare(b));
}

/** One pending per-person removal, and everything the dialog needs to describe it honestly. */
export interface MemberRemoval {
  user: SpGroupMember;
  groupId: number;
  groupName: string;
  /** Other allowed groups that keep this person in at this scope. */
  survivingGroups: string[];
  /** Allowed groups whose members could not be read, so `survivingGroups` may be short. */
  unreadable: string[];
  /** True when this is the group's last remaining member. */
  lastMember: boolean;
}

export function buildRemoval(args: {
  user: SpGroupMember;
  groupId: number;
  groupName: string;
  allowedGroupIds: readonly number[];
  map: MemberMap;
  titleOf: (groupId: number) => string;
}): MemberRemoval {
  const index = membershipIndex(args.allowedGroupIds, args.map);
  const load = loadOf(args.map, args.groupId);
  return {
    user: args.user,
    groupId: args.groupId,
    groupName: args.groupName,
    survivingGroups: otherAllowedGroups(args.user.id, args.groupId, index, args.titleOf),
    unreadable: unreadableGroups(args.allowedGroupIds, args.map, args.titleOf),
    // Only claimable from a LOADED list. On an unreadable one this stays false, which understates
    // rather than asserting — the safe direction for a sentence about what is left behind.
    lastMember: load.state === "loaded" && load.members.length === 1,
  };
}

/**
 * The sentence the dialog leads with: does this actually end their access here?
 *
 * Deliberately not a boolean. "Ends it", "does not end it" and "cannot tell" are three different
 * things to tell an admin, and flattening the third into either of the others is how a screen ends
 * up making a promise it has not checked.
 */
export function removalVerdict(r: MemberRemoval): "ends" | "survives" | "unknown" {
  if (r.survivingGroups.length > 0) return "survives";
  if (r.unreadable.length > 0) return "unknown";
  return "ends";
}

/**
 * Display name for a member, never blank.
 *
 * A row with no visible label cannot be acted on with any confidence, and a SharePoint principal
 * can legitimately have an empty Title — some system accounts and guests do.
 */
export function memberLabel(m: SpGroupMember): string {
  const title = (m?.title ?? "").trim();
  if (title.length > 0) return title;
  const email = (m?.email ?? "").trim();
  if (email.length > 0) return email;
  return `User ${m?.id ?? "?"}`;
}

/** Members in a stable, readable order. Sorted by label so the same list never reshuffles. */
export function sortMembers(members: readonly SpGroupMember[]): SpGroupMember[] {
  return (members ?? []).slice().sort((a, b) => memberLabel(a).localeCompare(memberLabel(b)));
}

/**
 * The count shown on a collapsed row.
 *
 * "No members" is only ever said about a list that was actually read. The other two states say so,
 * because a collapsed row is where an admin decides whether it is worth expanding — and "no
 * members" on a group that merely failed to load is what stops them opening the row that holds the
 * person they came for.
 */
export function memberCountLabel(load: MemberLoad): string {
  if (load.state === "loading") return "loading…";
  if (load.state === "error") return "could not read members";
  const n = load.members.length;
  if (n === 0) return "no members";
  return `${n} member${n === 1 ? "" : "s"}`;
}
