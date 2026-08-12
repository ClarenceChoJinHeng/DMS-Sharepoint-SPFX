import { sanitizeFolderSegment } from "./formModel";
import { AbbrevTarget, findCollisions } from "./folderAbbreviation";

/**
 * Editing a segment's folder abbreviations — the decision layer.
 *
 * Spec: docs/superpowers/specs/2026-08-12-term-abbreviation-page-design.md
 *
 * Folder names come from `DMS Term Abbreviation`, keyed by term GUID. A term with no abbreviation
 * is SKIPPED by every reconciliation run, so that unit gets no folder and cannot upload — and the
 * thing that must never happen is two SIBLINGS sharing a code, because both then resolve to one
 * folder with one ACL and each unit reads the other's documents.
 *
 * Reconciliation already aborts on that. This moves the failure to typing time, which is cheap: the
 * person who runs reconciliation is usually not the person who typed the code.
 *
 * The collision rule itself is NOT reimplemented here — it delegates to `findCollisions`, the same
 * function reconciliation uses. Two copies of that rule would eventually disagree, and the way they
 * would disagree is one of them permitting a pair that merges two units into one folder.
 *
 * Pure, so the rules are testable without a tenant.
 */

/** One editable row: a term, its place in the tree, and the code being typed for it. */
export interface AbbrevRowDraft {
  termGuid: string;
  /** The term's live label — also what "Same as term name" copies in. */
  label: string;
  /** The level NAME, e.g. "Department" / "Unit" / "Region". Stored, not derived. */
  level: string;
  /** Parent term GUID; "" for a top-level term. Siblings share this. */
  parentGuid: string;
  /** What the admin has typed. */
  abbreviation: string;
  /** The row's existing saved value, so a CHANGE can be told from a first-time fill. */
  original?: string;
}

/**
 * The folder name a code will actually produce.
 *
 * Comparison and length both happen on this rather than the raw text, because this is what
 * SharePoint will be asked to create. `GC EP` and `GC  EP` collapse to one segment here — a pair
 * `findCollisions` would miss on its own, since it lowercases but does not sanitize. Feeding it
 * sanitized names makes the page stricter than reconciliation, which is the safe direction: it can
 * block something reconciliation would have accepted, never the reverse.
 */
export function folderNameFor(abbreviation: string): string {
  return sanitizeFolderSegment(abbreviation ?? "");
}

/**
 * Above this many characters, a folder segment gets a warning.
 *
 * Not a refusal: the client may legitimately want a long name, and "Same as term name" is one click
 * away from a 60-character label. But abbreviations exist to keep paths short — a deep path returned
 * HTTP 400 at roughly 330 characters (gotcha #9) — so the cost has to be visible where it is chosen.
 */
export const LONG_NAME_THRESHOLD = 15;

export interface RowProblem {
  /** Blocks saving. A collision would provision two units into one folder with one ACL. */
  error?: string;
  /** Does not block. A missing code and a long name are legitimate, deliberate states. */
  warn?: string;
}

/**
 * Validate every row against its SIBLINGS, returning problems keyed by term GUID.
 *
 * Siblings only, deliberately: `Tax` under Group Finance and `Tax` under Minamas GA are different
 * paths and must both be allowed to be `TAX`. Not an edge case — `Tax`, `Legal` and `PM` each repeat
 * under several parents in the client's data, so a global uniqueness rule would reject correct input.
 *
 * BOTH rows of a collision are flagged, not just the later one: which of the two should change is
 * not knowable from here, and flagging only one implies the other is correct.
 */
export function validateRows(rows: AbbrevRowDraft[]): Record<string, RowProblem> {
  const out: Record<string, RowProblem> = {};

  // Parent term GUID stands in for `parentPath`, which findCollisions treats as an opaque grouping
  // key. Names are sanitized first — see folderNameFor.
  const targets: AbbrevTarget[] = rows
    .filter((r) => folderNameFor(r.abbreviation) !== "")
    .map((r) => ({
      parentPath: (r.parentGuid ?? "").toLowerCase(),
      termGuid: r.termGuid,
      abbreviation: folderNameFor(r.abbreviation),
      label: r.label,
    }));

  // termGuid -> the labels it clashes with. Built from findCollisions' own grouping so the page and
  // reconciliation always agree on what a clash is.
  const clashLabels: Record<string, string[]> = {};
  for (const c of findCollisions(targets)) {
    const members = targets.filter(
      (t) =>
        t.parentPath === c.parentPath &&
        t.abbreviation.toLowerCase() === c.abbreviation.toLowerCase(),
    );
    for (const m of members) {
      clashLabels[m.termGuid] = members
        .filter((o) => o.termGuid !== m.termGuid)
        .map((o) => o.label);
    }
  }

  for (const r of rows) {
    const typed = (r.abbreviation ?? "").trim();
    const name = folderNameFor(typed);

    if (!typed) {
      out[r.termGuid] = {
        warn: `No folder will be created for "${r.label}" — that unit cannot upload until this has a name.`,
      };
      continue;
    }

    // Typed something, but every character was one SharePoint strips. Creating nothing would read
    // exactly like the blank case while the box looks filled in.
    if (!name) {
      out[r.termGuid] = {
        error: `"${typed}" leaves nothing usable once the characters SharePoint cannot put in a folder name are removed.`,
      };
      continue;
    }

    const others = clashLabels[r.termGuid];
    if (others && others.length > 0) {
      out[r.termGuid] = {
        error:
          `"${name}" is already used by ${others.map((o) => `"${o}"`).join(", ")} under the same ` +
          `parent. Both would share one folder and one set of permissions, so each unit would see ` +
          `the other's documents. Choose a different code.`,
      };
      continue;
    }

    if (name.length > LONG_NAME_THRESHOLD) {
      out[r.termGuid] = {
        warn:
          `"${name}" is ${name.length} characters. It will work, but long names at this level make ` +
          `deep paths that SharePoint can reject — short codes are what abbreviations are for.`,
      };
    }
  }

  return out;
}

/** True when any row holds a blocking problem. Saving is refused while this is true. */
export function hasBlockingProblem(problems: Record<string, RowProblem>): boolean {
  for (const key of Object.keys(problems)) {
    if (problems[key].error) return true;
  }
  return false;
}

/** Rows whose code differs from what is saved — the ones a save has to write. */
export function changedRows(rows: AbbrevRowDraft[]): AbbrevRowDraft[] {
  return rows.filter((r) => (r.abbreviation ?? "").trim() !== (r.original ?? "").trim());
}

/**
 * Rows that already had a code and are now being given a DIFFERENT one.
 *
 * Called out separately because these RENAME a live folder on the next reconciliation run, while a
 * first-time fill only creates one. The rename is safe — UniqueId, contents, ACL and approval status
 * all survive, in both libraries — but it is not what someone filling in blanks expects to have done.
 */
export function renamingRows(rows: AbbrevRowDraft[]): AbbrevRowDraft[] {
  return changedRows(rows).filter((r) => (r.original ?? "").trim() !== "");
}
