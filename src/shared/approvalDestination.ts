/**
 * Where an approved document is going, and which folder decides whether that is safe.
 *
 * The approval web part refuses to approve until the destination UNIT folder exists and has unique
 * permissions. That guard is sound; how it located the unit folder was not.
 *
 * ── The bug this replaces (found live 2026-08-24, blocked every approval on a segment) ───────────
 * It walked UP three levels from the file, on the assumption that a path is
 * `…/Unit/Year/Document Type/file`. Its comment claimed this "holds for segments with a deeper
 * Levels chain too", which is exactly backwards: a deeper chain is what breaks it. The moment a
 * below-Unit tier was added through the Structure Manager, three levels up landed on the YEAR
 * folder — which inherits the unit's ACL by design, so `HasUniqueRoleAssignments` read false, and
 * the guard refused every approval in that segment while reporting "it is not locked down".
 *
 *   ApprovalDocument / GHO / GCA / EG / 2024 / Agreement / Archive 1 / file.pdf
 *                      seg   dept  UNIT   ...three levels up is 2024, not EG
 *
 * One below-Unit tier is just as wrong in the other direction: three up lands on the DEPARTMENT,
 * which does have unique permissions, so the guard passes having checked the wrong folder.
 *
 * ── Why the depth is counted from the TOP, and passed in ─────────────────────────────────────────
 * The unit folder sits at `<segment folder>/<tier 1>/…/<tier N>`, where N is the number of
 * PERMISSIONED tiers. Counted from the library root that is exact whatever hangs below it; counted
 * from the file it depends on how many below-Unit tiers a segment happens to have today.
 *
 * N cannot be derived from the document's own fields, and the near-miss is worth stating because it
 * looks correct: every tier column has a `<Base>Tid` twin, so `discoverTierFields` finds them — but
 * `documentDetails.ts` notes that a below-Unit tier such as SubUnit carries a Tid column exactly
 * like a permissioned one. Counting those would land on the SubUnit folder, which inherits, and the
 * guard would refuse all over again for a new reason. Only the mode row's `Levels` knows which tiers
 * are permissioned, so the caller reads it and passes the count here.
 *
 * ── Failure direction ───────────────────────────────────────────────────────────────────────────
 * Every failure REFUSES and names itself, against this codebase's usual fail-open habit. That is the
 * existing guard's own reasoning and it still holds: a refused approval costs the approver a retry,
 * while a wrong "proceed" publishes a unit's documents to everyone who can open the site and nobody
 * finds out. `permissionedTiers: undefined` is therefore a refusal, not a wave-through.
 */

export interface UnitFolderPath {
  /** Server-relative path of the unit folder in the DESTINATION library. */
  path: string;
}

export interface UnitFolderProblem {
  /** Reads after "the document was NOT approved (" — lower case, no trailing stop. */
  error: string;
}

export type UnitFolderResult = UnitFolderPath | UnitFolderProblem;

export function isProblem(r: UnitFolderResult): r is UnitFolderProblem {
  return (r as UnitFolderProblem).error !== undefined;
}

function trimSlashes(s: string): string {
  return (s ?? "").replace(/^\/+/, "").replace(/\/+$/, "");
}

/**
 * Permissioned tiers below the segment folder, from a mode row's `Levels` JSON.
 *
 * `undefined` when the JSON is absent or unparseable — NOT 0. The caller refuses on either, but the
 * two mean different things and only one of them is worth telling an administrator about.
 *
 * `permissioned` ABSENT MEANS TRUE, matching `folderChain.isPermissioned`: only a literal `false`
 * demotes a tier, and the string `"false"` does not. That default fails loudly (a tier counted that
 * should not have been, so the check lands one level too deep and refuses) rather than silently (a
 * tier skipped, landing one level shallow on a folder that looks plausible and passes).
 */
export function permissionedTierCount(levelsJson: string | undefined): number | undefined {
  const raw = (levelsJson ?? "").trim();
  if (raw.length === 0) return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return undefined;
  }
  if (!Array.isArray(parsed)) return undefined;
  let n = 0;
  for (const entry of parsed) {
    if (entry === null || typeof entry !== "object") return undefined;
    if ((entry as { permissioned?: unknown }).permissioned === false) continue;
    n += 1;
  }
  return n;
}

/**
 * The unit folder in the destination library, for the file being approved.
 *
 * `sourceSegment` and `destSegment` are URL segments, never titles — the two differ throughout this
 * project (`Approval Document` at `/ApprovalDocument`, `Documents` at `/Shared Documents`) and only
 * one of them fails loudly when confused. Both are resolved by the caller, so the HC pair is handled
 * by passing HC segments rather than by a branch in here.
 */
export function unitFolderPath(args: {
  /** e.g. `/sites/CRS/ApprovalDocument/GHO/GCA/EG/2024/Agreement/Archive 1/file.pdf` */
  fileSru: string;
  /** e.g. `/sites/CRS` */
  webSru: string;
  /** URL segment of the library the file is in now, e.g. `ApprovalDocument`. */
  sourceSegment: string;
  /** URL segment of the library it will be copied to, e.g. `Shared Documents`. */
  destSegment: string;
  /** Permissioned tiers BELOW the segment folder — 2 for `Department, Unit`. */
  permissionedTiers: number | undefined;
}): UnitFolderResult {
  const { fileSru, webSru, sourceSegment, destSegment, permissionedTiers } = args;

  /* An unresolved HC pair arrives here as a blank segment and is refused, never quietly redirected
     at the normal Documents library — that would file a Highly Confidential document where the whole
     unit can read it. Same reason `hcRouting.ts` returns undefined rather than falling back. */
  if (trimSlashes(sourceSegment).length === 0 || trimSlashes(destSegment).length === 0) {
    return { error: "the libraries could not be resolved on this site" };
  }

  if (permissionedTiers === undefined) {
    return {
      error: "this segment's folder structure could not be read, so its unit folder could not be located",
    };
  }
  if (permissionedTiers < 1) {
    /* Every family in this system has at least two permissioned tiers and `validateNewSegment`
       refuses fewer, so zero means the Levels JSON was misread rather than that a segment is flat. */
    return {
      error: "this segment records no permissioned levels, which cannot be right — check its Levels configuration",
    };
  }

  const prefix = `${trimSlashes(webSru)}/${trimSlashes(sourceSegment)}/`;
  const flat = trimSlashes(fileSru);
  if (flat.toLowerCase().indexOf(prefix.toLowerCase()) !== 0) {
    return { error: "its path is not inside the approval library, so the destination could not be worked out" };
  }

  const parts = flat.slice(prefix.length).split("/").filter((p) => p.length > 0);
  // 1 for the segment folder (GHO), then one per permissioned tier, then the file itself.
  const need = 1 + permissionedTiers;
  if (parts.length < need + 1) {
    /* The file sits at or above the unit folder — filed too shallow, or the chain is wrong. Refusing
       is right either way: there is no unit folder to check, and approving would let Auto-route
       create a path nobody has ACL'd. */
    /* `- 2` drops the segment folder AND the file itself, so the number counts TIERS. Counting the
       file made `GHO/GF/file.pdf` report two levels when it is filed one below the segment — the
       kind of off-by-one that sends an administrator looking at the wrong folder. */
    const depth = Math.max(0, parts.length - 2);
    return {
      error:
        `it is filed only ${depth} level(s) below the business segment, but this segment's ` +
        `unit folder is ${permissionedTiers} level(s) down`,
    };
  }

  return {
    path: `/${trimSlashes(webSru)}/${trimSlashes(destSegment)}/${parts.slice(0, need).join("/")}`,
  };
}
