/**
 * Display order for the Confidential Level dropdown — purely cosmetic, unlike `hcRouting.ts` beside
 * it, which decides what is SAFE to offer. This only decides what ORDER the safe-to-offer options
 * appear in.
 *
 * Client, 2026-09-18: "It should follow the sequence: 'Highly Confidential, Confidential &
 * Restricted,' to ease the user when making selection." The term store returns these alphabetically
 * (Confidential, Highly Confidential, Restricted), which is what the dropdown showed before this.
 *
 * Kept OUT of `hcRouting.ts` deliberately — that module's whole reason for existing is that every
 * rule in it has a wrong version that looks identical on screen, and mixing a cosmetic sort into a
 * module about what may be disclosed is how the two kinds of rule end up read as equally important.
 */

const RANK: Record<string, number> = {
  "highly confidential": 0,
  confidential: 1,
  restricted: 2,
};

/**
 * Rank for a confidentiality label, for use as a sort key. Unrecognised labels (a segment that has
 * renamed or added a level in its own term set) all rank equally, past the three known ones — and
 * `Array.prototype.sort` is stable, so they keep whatever order the term store gave them rather than
 * being shuffled.
 */
export function confidentialityRank(label: string): number {
  const key = (label ?? "").trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(RANK, key)
    ? RANK[key]
    : Number.MAX_SAFE_INTEGER;
}

/**
 * Sorts a list of `{ label }`-shaped items into the client's requested order. Returns a NEW array —
 * never mutates the input, since both call sites pass a list read from live term-store options that
 * other code may still hold a reference to.
 */
export function sortByConfidentialityOrder<T extends { label?: string }>(
  items: readonly T[],
): T[] {
  return items
    .slice()
    .sort(
      (a, b) => confidentialityRank(a.label ?? "") - confidentialityRank(b.label ?? ""),
    );
}
