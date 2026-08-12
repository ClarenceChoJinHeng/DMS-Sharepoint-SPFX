import { Level, sanitizeFolderSegment } from "./formModel";
import { validateChain } from "./folderChain";

/**
 * Adding a whole business segment — the decision layer.
 *
 * Spec: docs/superpowers/specs/2026-08-12-add-segment-design.md (slice B of the Structure
 * Manager; slice A edits the levels of a segment that already exists).
 *
 * Onboarding a segment by hand means authoring a `DMS Config` mode row — a Title, a Category,
 * a term set GUID, a StagingFolder, a SortOrder and a `Levels` JSON chain — and creating the
 * matching columns in both libraries under names whose internal form is frozen at creation.
 * The client cannot edit JSON, and after handover nobody can do it for them.
 *
 * Everything here is pure so the rules are testable without a tenant. The I/O — measuring the
 * term set's depth, creating the columns, writing the row — lives in SegmentCreator.tsx.
 */

/** One permissioned tier the admin named, e.g. "Region" then "Estate/Mill". */
export interface PermissionedTierDraft {
  label: string;
}

export interface NewSegmentDraft {
  /** ModeLabel — what uploaders see in the Segment dropdown. */
  label: string;
  /** Category: the top-level family, which drives which tab it appears under. */
  family: "BusinessSegment" | "Project";
  termSetGuid: string;
  /** StagingFolder — the top folder reconciliation creates, e.g. "UPOPS". */
  stagingFolder: string;
  /** ACL'd tiers, in order. Their columns are created in both libraries. */
  permissioned: PermissionedTierDraft[];
  /** Below-Unit tiers, already in `Level` form (the slice A editor's output). */
  below: Level[];
}

/** An existing mode row, reduced to what a new one must not collide with. */
export interface ExistingSegment {
  key: string;
  label: string;
  stagingFolder: string;
  sortOrder?: number;
}

/**
 * Derive the column internal name from a tier label.
 *
 * A SharePoint column's internal name is fixed at creation, permanently, from the title it is
 * created with. Letting an untrained admin type it offers a decision that cannot be undone and
 * whose consequence is invisible until a metadata write silently fails. So strip everything
 * SharePoint would encode, create under that, then set the display title to what they typed.
 */
export function columnNameFor(label: string): string {
  return sanitizeFolderSegment(label).replace(/[^A-Za-z0-9]/g, "");
}

/**
 * Derive the mode row's `Title` — the key — from the segment name. Never typed.
 *
 * It is a key: permanent, and referenced by folders already filed. Offering it as free text
 * invites a duplicate or a typo whose only symptom is a segment that never appears in the form,
 * with nothing anywhere saying why.
 */
export function modeKeyFor(label: string): string {
  const slug = sanitizeFolderSegment(label)
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase();
  return slug ? `mode_${slug}` : "";
}

/**
 * The next SortOrder. Its only job is ordering a dropdown, so it is derived rather than asked.
 *
 * Guards against a non-numeric or absent cell: a hand-authored row with SortOrder blank reads as
 * NaN, and `NaN + 1` would write NaN into the list and unsort every segment at once.
 */
export function nextSortOrder(existing: ExistingSegment[]): number {
  let max = 0;
  for (const e of existing) {
    const n = Number(e.sortOrder);
    if (isFinite(n) && n > max) max = Math.floor(n);
  }
  return max + 1;
}

/**
 * Assemble the `Levels` chain: permissioned tiers first, then the below-Unit ones.
 *
 * Permissioned tiers are written with an EXPLICIT `permissioned: true` even though absent means
 * true. The default is right, but this row is read by reconciliation to decide what gets its own
 * ACL, and a chain where the ACL'd tiers are the ones with no flag reads as an oversight to the
 * next person editing it by hand.
 *
 * They carry no `termSet`: a permissioned tier cascades within the segment's own term tree, and
 * the presence of `termSet` is exactly the discriminator that would turn it into a flat list.
 */
export function buildSegmentLevels(draft: NewSegmentDraft): Level[] {
  const permissioned: Level[] = draft.permissioned.map((t) => {
    const col = columnNameFor(t.label);
    return {
      label: t.label.trim(),
      column: col,
      labelCol: col,
      tidCol: `${col}Tid`,
      permissioned: true,
    };
  });
  return [...permissioned, ...draft.below.map((l) => ({ ...l }))];
}

/** Every column internal name the new segment needs created, label + Tid per tier. */
export function columnsForDraft(
  draft: NewSegmentDraft,
): Array<{ internal: string; display: string }> {
  const out: Array<{ internal: string; display: string }> = [];
  for (const tier of draft.permissioned) {
    const col = columnNameFor(tier.label);
    if (!col) continue;
    out.push({ internal: col, display: tier.label.trim() });
    out.push({ internal: `${col}Tid`, display: `${tier.label.trim()} ID` });
  }
  return out;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

/** Accepts the braced/whitespaced forms people paste out of SharePoint UI and URLs. */
export function normalizeGuid(raw: string): string {
  return raw
    .trim()
    .replace(/^[{(]|[)}]$/g, "")
    .trim()
    .toLowerCase();
}

export function isGuid(raw: string): boolean {
  return GUID_RE.test(normalizeGuid(raw));
}

/**
 * Everything that must be true before a single column is created.
 *
 * Nothing here touches the term store — depth is measured separately because it costs requests,
 * and there is no point spending them on a draft that fails these.
 */
export function validateNewSegment(
  draft: NewSegmentDraft,
  existing: ExistingSegment[],
): string[] {
  const errors: string[] = [];
  const label = draft.label.trim();
  const key = modeKeyFor(label);
  const folder = sanitizeFolderSegment(draft.stagingFolder).trim();
  const norm = (v: string): string => v.trim().toLowerCase();

  if (!label) errors.push("Give the segment a name — it is what uploaders pick from.");
  else if (!key) {
    errors.push(
      `"${label}" has no letters or numbers in it, so no configuration key can be derived from it.`,
    );
  }

  if (label && existing.filter((e) => norm(e.label) === norm(label)).length > 0) {
    errors.push(`A segment called "${label}" already exists.`);
  }
  // Checked separately from the name: two different names can slug to one key ("R&D" and "R D"),
  // and that collision would present as the new segment shadowing the old one rather than as a
  // naming clash.
  const keyClash = existing.filter((e) => key && norm(e.key) === norm(key))[0];
  if (keyClash) {
    errors.push(
      `That name produces the configuration key "${key}", which "${keyClash.label}" already ` +
        `uses. Choose a name that differs by more than punctuation.`,
    );
  }

  if (!folder) {
    errors.push(
      "Give the segment a top folder name, e.g. UPOPS — it is the folder every document in this segment sits under.",
    );
  } else {
    const folderClash = existing.filter((e) => norm(e.stagingFolder) === norm(folder))[0];
    if (folderClash) {
      // Two segments sharing a top folder would merge their trees under one set of ACLs — every
      // unit of one reachable by the other's groups. The loudest possible refusal.
      errors.push(
        `The top folder "${folder}" is already used by "${folderClash.label}". Two segments ` +
          `cannot share one folder — their units would merge and so would their permissions.`,
      );
    }
  }

  if (!isGuid(draft.termSetGuid)) {
    errors.push(
      "Paste the segment's term set ID. It looks like 023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf.",
    );
  }

  if (draft.permissioned.length === 0) {
    errors.push(
      "Name at least one permissioned level — the level whose folders carry the permissions, " +
        "usually Unit. Without one there is nothing for reconciliation to grant access to.",
    );
  }
  const seen: Record<string, string> = {};
  for (const tier of draft.permissioned) {
    const name = tier.label.trim();
    if (!name) {
      errors.push("One of the permissioned levels has no name.");
      continue;
    }
    const col = columnNameFor(name);
    if (!col) {
      errors.push(`"${name}" has no letters or numbers in it, so it cannot become a column.`);
      continue;
    }
    const prior = seen[col.toLowerCase()];
    if (prior) {
      errors.push(
        `"${name}" and "${prior}" would both need the "${col}" column. Give them names that ` +
          `differ by more than punctuation.`,
      );
      continue;
    }
    seen[col.toLowerCase()] = name;
  }

  // The chain rule that outranks the rest: permissioned tiers must be a contiguous prefix.
  // Building the chain here rather than trusting the form's ordering means a below-Unit tier that
  // somehow carries permissioned:true is caught before any column exists.
  if (errors.length === 0) {
    const chainError = validateChain(buildSegmentLevels(draft));
    if (chainError) errors.push(chainError.message);
  }

  return errors;
}

/**
 * Compare the term set's measured depth against the number of permissioned tiers.
 *
 * THE check this screen exists to make, and the one that is not guessable from the UI.
 * Reconciliation walks the segment's TERM TREE — not `Levels` — and caps its depth at the
 * permissioned tier count. So naming two tiers (Region → Estate) against a three-deep set makes
 * reconciliation stop one level early: the folders it ACLs are Regions while the Group Map rows
 * point at Estates, and every unit-level grant lands on the wrong folder. Nothing errors. It
 * provisions the wrong tree, and it surfaces weeks later as "this person can see too much".
 *
 * `undefined` depth means the walk could not complete (throttled, or capped). That is NOT a
 * mismatch and must not refuse — an unreadable term store says nothing about the draft — so it
 * warns instead. Same rule as everywhere else here: empty is not unknown.
 */
export function depthVerdict(
  measured: number | undefined,
  tierCount: number,
): { ok: boolean; warn?: string; error?: string } {
  const levels = (n: number): string => `${n} ${n === 1 ? "level" : "levels"}`;
  if (measured === undefined) {
    return {
      ok: true,
      warn:
        "Could not measure how deep that term set is. Check it has exactly " +
        `${levels(tierCount)} of terms before you run reconciliation — if it has more or fewer, ` +
        "the folders that carry permissions will be the wrong ones.",
    };
  }
  if (measured === 0) {
    return {
      ok: false,
      error:
        "That term set has no terms in it yet. Add the segment's structure to the term store " +
        "first — reconciliation builds folders from the terms, so an empty set creates nothing.",
    };
  }
  if (measured !== tierCount) {
    return {
      ok: false,
      error:
        `That term set is ${levels(measured)} deep, but you have named ${levels(tierCount)} of ` +
        `permissions. They must match: reconciliation builds one folder level per term level ` +
        `and puts the permissions on the deepest one, so a mismatch silently protects the wrong ` +
        `folders. ` +
        (measured > tierCount
          ? "Name another level, or point at a different term set."
          : "Remove a level, or point at a different term set."),
    };
  }
  return { ok: true };
}
