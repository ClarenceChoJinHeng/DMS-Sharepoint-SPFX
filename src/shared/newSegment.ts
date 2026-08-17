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
  /**
   * The Config list item id. Optional because creation never needs it — only deletion does, and
   * only a reader that asked for `Id` will have it.
   */
  itemId?: number;
  /**
   * The segment's term set. Optional for the same reason: deletion matches Group Map rows on it
   * (a folder row's `Segment` holds the term-set GUID), and nothing in the creation path reads it.
   */
  termSetGuid?: string;
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
 * The three collisions that can be shown BESIDE the field that caused them, as they are typed.
 *
 * Client, 2026-08-17: *"the way to show an error is weirdly not noticeable. I type in the same GUID and
 * yes it shows but when I click on Create Segment it doesn't push me up but only show a dialog which
 * only I can see once I scroll up manually."* The form is taller than the viewport, so a summary at the
 * top is invisible at the moment it is needed — and the admin's reading is that Create did nothing.
 *
 * Deliberately NOT a second rule set: each message is the same collision `validateNewSegment` refuses,
 * so this can never permit what Create would reject. It returns only the three that BELONG to a single
 * field; everything else (a name with no letters, a bad tier list) stays in the summary, because a
 * message has to sit next to the thing it is about or it is just a differently-placed summary.
 *
 * Empty strings mean "nothing to say", so a caller can render unconditionally.
 */
export function fieldConflicts(
  draft: NewSegmentDraft,
  existing: ExistingSegment[],
): { label: string; folder: string; termSet: string } {
  const norm = (v: string): string => v.trim().toLowerCase();
  const rows = existing ?? [];
  const label = draft.label.trim();
  const key = modeKeyFor(label);
  const folder = sanitizeFolderSegment(draft.stagingFolder).trim();

  let labelMsg = "";
  if (label) {
    const byName = rows.filter((e) => norm(e.label) === norm(label))[0];
    // The KEY clash is reported on the name field too, because the key is derived from the name and is
    // not a field anyone can edit — telling someone their key collides while showing them no key would
    // leave them nothing to change.
    const byKey = rows.filter((e) => key && norm(e.key) === norm(key))[0];
    if (byName) labelMsg = `"${byName.label}" already exists.`;
    else if (byKey) {
      labelMsg =
        `This name produces the key "${key}", which "${byKey.label}" already uses. ` +
        `Differ by more than punctuation.`;
    }
  }

  let folderMsg = "";
  if (folder) {
    const clash = rows.filter((e) => norm(e.stagingFolder) === norm(folder))[0];
    if (clash) folderMsg = `"${folder}" is already the top folder for "${clash.label}".`;
  }

  let setMsg = "";
  if (isGuid(draft.termSetGuid)) {
    const clash = rows.filter(
      (e) => normalizeGuid(e.termSetGuid ?? "") === normalizeGuid(draft.termSetGuid),
    )[0];
    if (clash) setMsg = `This term set already belongs to "${clash.label}".`;
  }

  return { label: labelMsg, folder: folderMsg, termSet: setMsg };
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
  } else {
    // TWO SEGMENTS MUST NOT SHARE A TERM SET, and until 2026-08-17 nothing checked it — found while
    // the client was deliberately re-entering an existing GUID to see the error.
    //
    // It is not merely redundant data. **Group Map rows key on `Segment` = the term-set GUID** (a
    // folder-access row stores the set, not the mode key), so two segments sharing a set share each
    // other's access rows: granting a group in one silently grants it in the other, and deleting one
    // segment removes the other's mappings. They would also build identical trees from identical
    // abbreviations under two top folders, doubling every folder and splitting the documents.
    const setClash = existing.filter(
      (e) => normalizeGuid(e.termSetGuid ?? "") === normalizeGuid(draft.termSetGuid),
    )[0];
    if (setClash) {
      errors.push(
        `That term set is already used by "${setClash.label}". Two segments cannot share one term ` +
          `set — they would build the same folders twice and share each other's folder-access rows.`,
      );
    }
  }

  // TWO is the floor, not one (client's instruction 2026-08-17: "I think best to force them to
  // create two not one. atleast two"). Every one of the thirteen intended segments is two tiers —
  // head offices Department/Unit, Upstream Ops Region/Estate·Mill, SDGI Refinery/Department, I&T
  // `I&T Operating Units/Department`/Unit (that slash is ONE tier name, not two tiers).
  //
  // This is NOT redundant with the term-set depth check, which is the reason it is worth having. That
  // check compares the declared count to the set's depth, so ONE declared tier against a ONE-deep set
  // passes it cleanly — and a set is one-deep exactly when someone has authored the departments but
  // not yet the units. The segment then saves with the DEPARTMENT holding the access, so every unit
  // in a department shares one ACL and one folder. Nothing fails; it surfaces later as "this person
  // can see another unit's documents".
  //
  // Deliberately a refusal rather than a warning: the depth check's own get-out is warn-and-allow
  // when depth is UNKNOWN, and stacking a second soft signal on the one shape that reads as complete
  // would leave the dangerous case advisory in both places.
  if (draft.permissioned.length < 2) {
    errors.push(
      draft.permissioned.length === 0
        ? "Name the levels that carry permissions — at least two, e.g. Department then Unit. " +
            "Without them there is nothing for reconciliation to grant access to."
        : "Name at least two levels that carry permissions, e.g. Department then Unit. With only " +
            `"${draft.permissioned[0]?.label?.trim() || "one level"}" the permissions land on that ` +
            "level's folders, so everything beneath it shares one set of access.",
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
