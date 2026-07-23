// Pure, SPFx-free half of the native SP site-group module. No @microsoft/*
// imports so it stays unit-testable in plain Jest (same pattern as formModel /
// groupMapModel). The REST half lives in spGroups.ts, which re-exports these.
// See docs/superpowers/specs/2026-07-23-native-sharepoint-groups-design.md.

export type SpGroup = { id: number; title: string };
export type SpGroupMember = { id: number; title: string; email: string; loginName: string };
export type PersonPick = { loginName: string; displayName: string; email: string };

/** Thrown by createSiteGroup when the title is already taken on the site collection. */
export const DUPLICATE_GROUP = "DUPLICATE_GROUP";

/** DMS_-prefixed site groups matching q (both case-insensitive). Built-in
 * site groups (Owners/Members/Visitors) are excluded by the prefix so an admin
 * cannot accidentally map them to a unit folder. */
export function filterDmsGroups(all: SpGroup[], q: string): SpGroup[] {
  const query = (q ?? "").trim().toLowerCase();
  return all
    .filter((g) => g.title.toUpperCase().indexOf("DMS_") === 0)
    .filter((g) => !query || g.title.toLowerCase().indexOf(query) !== -1);
}
