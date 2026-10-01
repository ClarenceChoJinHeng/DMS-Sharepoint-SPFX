// Collapsing Group Map rows into the groups that hold them.
//
// Folder Access shows one row per GROUP rather than one per mapping, and that is forced rather than
// cosmetic: membership is a property of the group, and a unit's approver group carries six mapping
// rows — so a per-mapping member editor would show the same people thirteen times per unit, with
// thirteen Add boxes that all did the same thing. 790 rows become 308.
//
// Spec: docs/superpowers/specs/2026-08-18-folder-access-membership-design.md §2.1
//
// Pure, because it decides what an admin reads about who holds a permission.

/** The fields the grouping needs. Callers pass their own richer row type through. */
export interface MappingRow {
  itemId: number;
  GroupId: string;
  GroupName: string;
  Role: string;
}

export interface MappingGroup<T extends MappingRow = MappingRow> {
  groupId: string;
  /** As stored on the rows. Blank when the group was deleted in SharePoint. */
  groupName: string;
  /** What to show: the name, or the id when there is no name. Never blank. */
  label: string;
  rows: T[];
}

/**
 * One entry per group, sorted by name, each carrying that group's mapping rows in the order given.
 *
 * KEYED ON `GroupId`, never on the name. Two groups can be renamed to the same title, a stored
 * `GroupName` can be stale, and the id is what the grant actually hangs off — it is also what
 * `removeGroupMember` needs, so keying on anything else would offer a member editor pointed at the
 * wrong group.
 *
 * A blank id gets its own bucket rather than merging with every other blank: such a row is broken
 * data, and merging them would present unrelated grants as one group's mappings. Listed separately,
 * they can be found and deleted.
 */
export function groupMappingsByGroup<T extends MappingRow>(rows: T[]): Array<MappingGroup<T>> {
  const out: Array<MappingGroup<T>> = [];
  const byId: Record<string, MappingGroup<T>> = {};
  let blanks = 0;
  for (const r of rows ?? []) {
    if (!r) continue;
    const id = (r.GroupId ?? "").trim();
    // Blank ids are never shared, so each gets a key nothing else can collide with.
    const key = id ? `id:${id.toLowerCase()}` : `blank:${blanks++}`;
    let entry = byId[key];
    if (!entry) {
      const name = (r.GroupName ?? "").trim();
      entry = { groupId: id, groupName: name, label: name || id || "(no group)", rows: [] };
      byId[key] = entry;
      out.push(entry);
    }
    entry.rows.push(r);
  }
  return out.sort((a, b) => a.label.localeCompare(b.label));
}
