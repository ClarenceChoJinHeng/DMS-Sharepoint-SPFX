/**
 * Which authorised upload paths are actually PROVISIONED — i.e. their leaf folder
 * exists.
 *
 * Spec: docs/superpowers/specs/2026-08-12-provisioned-segment-visibility-design.md
 *
 * The upload form used to decide what to offer from `DMS Group Map` + the term tree
 * alone, and only discovered at write time whether the folder existed. That produced
 * two symptoms with one cause: a brand-new segment appeared before it had any folders,
 * and a user added to a unit group could pick their unit before reconciliation had run
 * (the origin of the HTTP 403 seen on 2026-08-11).
 *
 * So the gate is the same condition the write already checks — "is there a Folder Map
 * row for this leaf term?" — applied earlier. It is DERIVED, never stored: a `Ready`
 * flag on the mode row could be ticked before it was true, and this cannot disagree
 * with reality because it reads the very row the upload will look for.
 *
 * Pure and shared so both upload web parts gate identically and this is unit-testable
 * without SharePoint.
 */

/** The minimum shape this needs: a chain whose LAST entry is the leaf term. */
export interface LeafPath {
  chain: { id: string }[];
}

/**
 * Normalise a term GUID for comparison.
 *
 * SharePoint hands back GUIDs in more than one skin — braces from some endpoints,
 * mixed case between the term store and a hand-edited list cell. A case-sensitive
 * compare here would find no matches at all and hide EVERY unit on the site, which
 * presents as a total outage rather than a comparison bug.
 */
export function normalizeTermGuid(guid: string | undefined | null): string {
  return (guid ?? "")
    .trim()
    .replace(/^\{|\}$/g, "")
    .toLowerCase();
}

/** Build the lookup set from Folder Map rows. Blank GUIDs are dropped, not indexed. */
export function mappedTermGuidSet(
  rows: { termGuid?: string }[] | null | undefined,
): Set<string> | null {
  if (!rows) return null;
  const set = new Set<string>();
  for (const r of rows) {
    const key = normalizeTermGuid(r.termGuid);
    if (key) set.add(key);
  }
  return set;
}

export interface ProvisionedPaths<P> {
  /** The paths to offer. */
  paths: P[];
  /** How many were withheld because their leaf folder does not exist yet. */
  withheld: number;
  /**
   * False when provisioning could not be determined (an unreadable Folder Map), in
   * which case `paths` is the input untouched. Callers must not describe a withheld
   * path to the user when this is false — nothing was withheld.
   */
  known: boolean;
}

/**
 * Keep only the paths whose LEAF term has a Folder Map row.
 *
 * `mapped === null` means UNKNOWN — the read failed — and returns every path
 * untouched. Empty is not unknown (the same rule as gotcha 11's AllowedFileTypes and
 * the stale-chain guard): a transient error must never empty every dropdown on the
 * site. The upload-time check remains the backstop that makes that degraded path safe.
 */
export function filterProvisionedPaths<P extends LeafPath>(
  paths: P[],
  mapped: Set<string> | null,
): ProvisionedPaths<P> {
  if (!mapped) return { paths, withheld: 0, known: false };

  const kept: P[] = [];
  for (const p of paths) {
    const leaf = p.chain[p.chain.length - 1];
    // A path with no chain cannot be resolved to a folder and cannot be uploaded to,
    // so it is withheld rather than offered — but it is still COUNTED, because the
    // uploader needs to be told something is missing either way.
    if (leaf && mapped.has(normalizeTermGuid(leaf.id))) kept.push(p);
  }
  return { paths: kept, withheld: paths.length - kept.length, known: true };
}
