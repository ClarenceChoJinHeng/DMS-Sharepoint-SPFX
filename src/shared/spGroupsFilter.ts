// Pure, SPFx-free half of the native SP site-group module. No @microsoft/*
// imports so it stays unit-testable in plain Jest (same pattern as formModel /
// groupMapModel). The REST half lives in spGroups.ts, which re-exports these.
// See docs/superpowers/specs/2026-07-23-native-sharepoint-groups-design.md.

export type SpGroup = { id: number; title: string };
export type SpGroupMember = { id: number; title: string; email: string; loginName: string };
export type PersonPick = { loginName: string; displayName: string; email: string };

/** Thrown by createSiteGroup when the title is already taken on the site collection. */
export const DUPLICATE_GROUP = "DUPLICATE_GROUP";

/**
 * Site groups an admin may map to a folder, matching q (case-insensitive).
 *
 * Excludes by **id**, not by title. The old filter kept only "DMS_"-prefixed titles, but
 * the client dropped that prefix on 2026-08-04 (GHO_GF_CORU_UPLOADER), so there is no
 * longer anything in a title that marks a group as ours — and there cannot be, because
 * this picker's job is to show groups that are NOT yet mapped.
 *
 * What the prefix was really doing was keeping the built-ins out, so an admin could not
 * map "Site Owners" onto a unit folder. `excludeIds` does that directly: the caller passes
 * the site's associated Owner/Member/Visitor group ids plus the site-entry group.
 *
 * Ids rather than a title blocklist because built-in titles are derived from the SITE title
 * ("Clarence DMS Testing Owners"), so any list of names is correct on exactly one site.
 *
 * A missing/empty excludeIds is honoured as "exclude nothing" rather than falling back to
 * a prefix guess — the caller failing to read the associated groups is a visible problem
 * (built-ins appear in the list), not a silent one.
 */
export function filterSelectableGroups(
  all: SpGroup[],
  q: string,
  excludeIds?: number[],
): SpGroup[] {
  const query = (q ?? "").trim().toLowerCase();
  const excluded = new Set(excludeIds ?? []);
  return (all ?? [])
    .filter((g) => !excluded.has(g.id))
    .filter((g) => !query || (g.title ?? "").toLowerCase().indexOf(query) !== -1);
}
