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

/**
 * Whether a SEGMENT has any folders at all — the admin-facing counterpart of the path
 * filters below. Spec §9.
 *
 * Admins are exempt from the gate so they can build and test a segment, which means a
 * half-built one sits in their dropdown looking exactly like a live one. This is what
 * lets the picker say "not fully set up yet" beside it.
 *
 * Read from the Folder Map rows the form ALREADY loads: a row's `section` holds the
 * segment's top folder, written from `mode.stagingFolder` by reconciliation. So this
 * costs no request and adds nothing new to keep in sync.
 *
 * `unknown` is not `unprovisioned`, and that difference is the whole point. An unreadable
 * Folder Map would otherwise mark EVERY segment — including live ones — as not set up, and
 * send an admin to reconcile an intact tree. Same rule as `filterProvisionedPaths`: empty
 * is not unknown. A blank `stagingFolder` is unknown too: with nothing to match on, an
 * absence of rows proves nothing.
 *
 * Compared case-insensitively because SharePoint folder names are — `upopsmy` and
 * `UPOPSMY` are one folder, and a hand-typed StagingFolder differing only in case must not
 * read as unprovisioned.
 */
export type SegmentProvisionState = "provisioned" | "unprovisioned" | "unknown";

export function segmentProvisionState(
  stagingFolder: string | undefined | null,
  rows: { section?: string }[] | null | undefined,
): SegmentProvisionState {
  if (!rows) return "unknown";
  const want = (stagingFolder ?? "").trim().toLowerCase();
  if (!want) return "unknown";
  for (const r of rows) {
    if ((r.section ?? "").trim().toLowerCase() === want) return "provisioned";
  }
  return "unprovisioned";
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

/**
 * What an upload-access probe concluded about one folder. Mirrors `UploadAccess` in
 * dmsFolderMap so this module needs no SharePoint imports and stays testable alone.
 */
export type AccessVerdict = "granted" | "denied" | "missing" | "unknown";

/**
 * Keep only the paths this user can actually upload into.
 *
 * The existence gate above is not enough, and finding that out cost a live test on
 * 2026-08-12: reconciliation creates folders from the TERM TREE but assigns group ACLs
 * in a separate pass, so creating a group creates no folder and removes none. Two new
 * groups whose units already had folders sailed straight through an existence check and
 * were offered to an uploader with no access to either — the same HTTP 403 the gate was
 * built to prevent, arriving by a route existence cannot see. Existence is a property of
 * the folder; being able to upload is a property of the folder AND the user.
 *
 * `verdicts` is positional — `verdicts[i]` belongs to `paths[i]`. A missing entry counts
 * as `unknown`, so a short array fails OPEN rather than silently hiding the tail.
 *
 * `unknown` is KEPT. A probe that could not reach the server is not evidence of denial,
 * and blocking on it would let a throttle empty the form (§5 of the spec). It also clears
 * `known`, because once any verdict is uncertain the caller can no longer honestly tell
 * the user their folders are missing.
 */
export function filterReachablePaths<P>(
  paths: P[],
  verdicts: AccessVerdict[],
): ProvisionedPaths<P> {
  const kept: P[] = [];
  let certain = true;
  paths.forEach((p, i) => {
    const verdict = verdicts[i] ?? "unknown";
    if (verdict === "unknown") certain = false;
    if (verdict === "granted" || verdict === "unknown") kept.push(p);
  });
  return { paths: kept, withheld: paths.length - kept.length, known: certain };
}
