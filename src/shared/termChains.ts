// Term ancestry — turning a walked term tree into "which department is this unit in".
//
// A Group Map row stores only the LEAF term GUID, which is all reconciliation needs and not nearly
// enough to read a mappings table by: a unit row does not say which department it sits under, and a
// department row is indistinguishable from a unit one. `Tax`, `Legal` and `PM` each exist under
// several departments on the client's tree, so the leaf label alone is genuinely ambiguous — not
// merely terse (client's request, 2026-08-18).
//
// The ancestry is not stored anywhere and must not be: a term can be re-parented in the term store,
// and a copy in the list would then describe a shape that no longer exists. It is derived from a walk
// of the segment's tree, which the bulk provisioner already performs for a different reason.
//
// Pure and SPFx-free, so the rule that decides what an admin reads about a permission grant is
// testable without a tenant.

/** One term as a walk returns it. `parentId` is "" for a top-level term. */
export interface TermNode {
  id: string;
  label: string;
  parentId: string;
}

/** Case-insensitive: SharePoint is not consistent about GUID casing between stores. */
function key(guid: string): string {
  return (guid ?? "").trim().toLowerCase();
}

/**
 * Every walked term's chain of LABELS, shallowest first, keyed by lower-cased term GUID.
 *
 * A term whose parent is absent from the walk yields a chain of one — itself. That is deliberately
 * not the same as claiming it is top-level: the caller cannot distinguish the two from the chain
 * alone, so a partial walk must be reported as partial by the caller rather than presented as a
 * department. Chain LENGTH is what says which tier a term sits on, so it is never padded.
 */
export function tierChains(nodes: TermNode[]): Record<string, string[]> {
  const byId: Record<string, TermNode> = {};
  for (const n of nodes ?? []) {
    if (n && key(n.id)) byId[key(n.id)] = n;
  }
  const out: Record<string, string[]> = {};
  for (const id of Object.keys(byId)) {
    const chain: string[] = [];
    // Stops the moment a term repeats, rather than after some generous count. A cyclic parent link
    // is not a state the term store can produce, but a hung mappings table reads as a crashed page —
    // and a mere counter would still let a two-term cycle report a chain deeper than the tree, which
    // is worse than stopping: chain LENGTH is what says which tier a term sits on.
    const walked: Record<string, true> = {};
    let cur: TermNode | undefined = byId[id];
    while (cur && !walked[key(cur.id)]) {
      walked[key(cur.id)] = true;
      chain.unshift(cur.label);
      const parent = key(cur.parentId);
      if (!parent) break;
      cur = byId[parent];
    }
    out[id] = chain;
  }
  return out;
}

/**
 * One term's chain, or `undefined` when the walk never covered it.
 *
 * `undefined`, never `[]`: an empty chain renders as a term with no label, which reads as a blank
 * cell — "this term has no department" — when the truth is "this was never resolved". Empty is not
 * unknown, here as everywhere else in this codebase.
 */
export function chainFor(
  chains: Record<string, string[]>,
  guid: string,
): string[] | undefined {
  return (chains ?? {})[key(guid)];
}
