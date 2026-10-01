import { groupMappingsByGroup, MappingRow } from "./groupMappings";

const row = (itemId: number, groupId: string, groupName: string, role: string): MappingRow => ({
  itemId, GroupId: groupId, GroupName: groupName, Role: role,
});

describe("groupMappingsByGroup", () => {
  it("puts every row of one group under one entry", () => {
    const out = groupMappingsByGroup([
      row(1, "7", "GHO_GF_TAX_APPROVER", "APR"),
      row(2, "7", "GHO_GF_TAX_APPROVER", "DELS"),
      row(3, "8", "GHO_GF_TAX_UPLOADER", "UPL"),
    ]);
    expect(out.map((g) => g.groupName)).toEqual(["GHO_GF_TAX_APPROVER", "GHO_GF_TAX_UPLOADER"]);
    expect(out[0].rows.length).toBe(2);
    expect(out[1].rows.length).toBe(1);
  });

  it("keys on GroupId, never on the name", () => {
    // Two groups can be renamed to the same title, and a stored GroupName can be stale — the id is
    // what the grant actually hangs off, and what removeGroupMember needs.
    const out = groupMappingsByGroup([
      row(1, "7", "SAME_NAME", "UPL"),
      row(2, "9", "SAME_NAME", "UPL"),
    ]);
    expect(out.length).toBe(2);
    expect(out.map((g) => g.groupId).sort()).toEqual(["7", "9"]);
  });

  it("sorts by name so the list reads like the tree", () => {
    const out = groupMappingsByGroup([
      row(1, "9", "GHO_GS_CB_UPLOADER", "UPL"),
      row(2, "7", "GHO_GCA_EG_APPROVER", "APR"),
    ]);
    expect(out.map((g) => g.groupName)).toEqual(["GHO_GCA_EG_APPROVER", "GHO_GS_CB_UPLOADER"]);
  });

  it("keeps a row whose group has no name, under its id", () => {
    // A group deleted in SharePoint leaves its rows behind. They must still be listed — that is how
    // an admin finds a grant nobody can see or manage — so a blank name falls back to the id.
    const out = groupMappingsByGroup([row(1, "7", "", "UPL")]);
    expect(out.length).toBe(1);
    expect(out[0].groupName).toBe("");
    expect(out[0].label).toBe("7");
  });

  it("returns nothing for nothing, and never throws on rubbish", () => {
    expect(groupMappingsByGroup([])).toEqual([]);
    expect(groupMappingsByGroup(undefined as unknown as MappingRow[])).toEqual([]);
  });

  it("treats a blank GroupId as its own bucket rather than merging unrelated rows", () => {
    // A row with no group id is broken data, and merging every such row into one entry would present
    // them as one group's mappings. They are listed, separately, so they can be deleted.
    const out = groupMappingsByGroup([row(1, "", "A", "UPL"), row(2, "", "B", "UPL")]);
    expect(out.length).toBe(2);
  });
});
