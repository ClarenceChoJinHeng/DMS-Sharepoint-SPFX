import { chainFor, tierChains, TermNode } from "./termChains";

const nodes: TermNode[] = [
  { id: "d1", label: "Group Finance", parentId: "" },
  { id: "u1", label: "Tax", parentId: "d1" },
  { id: "u2", label: "Treasury", parentId: "d1" },
  { id: "d2", label: "Group Legal, Risk ＆ Compliance", parentId: "" },
  { id: "u3", label: "Legal", parentId: "d2" },
];

describe("tierChains", () => {
  it("gives a unit its department, and a department itself", () => {
    const c = tierChains(nodes);
    expect(c.u1).toEqual(["Group Finance", "Tax"]);
    expect(c.d1).toEqual(["Group Finance"]);
  });

  it("keeps two terms with the same label apart by their parent", () => {
    // "Tax", "Legal" and "PM" each exist under several departments. A map keyed on the LABEL
    // would collapse them, and the whole point of these columns is telling them apart.
    const dup: TermNode[] = [
      ...nodes,
      { id: "u4", label: "Tax", parentId: "d2" },
    ];
    const c = tierChains(dup);
    expect(c.u1).toEqual(["Group Finance", "Tax"]);
    expect(c.u4).toEqual(["Group Legal, Risk ＆ Compliance", "Tax"]);
  });

  it("matches parent ids case-insensitively — GUID casing differs between stores", () => {
    const c = tierChains([
      { id: "AAAA", label: "Dept", parentId: "" },
      { id: "bbbb", label: "Unit", parentId: "aaaa" },
    ]);
    expect(c.bbbb).toEqual(["Dept", "Unit"]);
  });

  it("returns the term alone when its parent is not in the set", () => {
    // A partial walk. The leaf is still the row's own term, so it is reported; what is NOT
    // reported is a department it might sit under, because that is the thing we do not know.
    const c = tierChains([{ id: "orphan", label: "Somewhere", parentId: "gone" }]);
    expect(c.orphan).toEqual(["Somewhere"]);
  });

  it("terminates on a cyclic parent link instead of spinning", () => {
    // Not a state the term store can produce, but a hung mappings table is indistinguishable
    // from a crashed page — the same guard planBulkGroups carries.
    const cyclic: TermNode[] = [
      { id: "a", label: "A", parentId: "b" },
      { id: "b", label: "B", parentId: "a" },
    ];
    expect(() => tierChains(cyclic)).not.toThrow();
    expect(tierChains(cyclic).a.length).toBeLessThanOrEqual(2);
  });

  it("survives an empty walk without inventing anything", () => {
    expect(tierChains([])).toEqual({});
    expect(chainFor({}, "anything")).toBeUndefined();
  });

  it("looks a chain up case-insensitively too", () => {
    const c = tierChains(nodes);
    expect(chainFor(c, "U1")).toEqual(["Group Finance", "Tax"]);
    // Absent is UNDEFINED, never []. An empty chain would render as a term with no label at
    // all, which reads as a blank cell rather than as "this was never resolved".
    expect(chainFor(c, "nope")).toBeUndefined();
  });
});
