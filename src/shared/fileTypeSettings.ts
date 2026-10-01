/**
 * Rules for the file type settings page — pure, no SharePoint imports.
 *
 * Spec: docs/superpowers/specs/2026-08-13-file-type-settings-page-design.md
 *
 * The READING half of this feature already lives in `allowedFileTypes.ts` and is not restated here:
 * that module owns `FALLBACK_FILE_TYPES`, the three-state resolution, and `normalizeFileTypes` for a
 * whole list. This module owns what the EDITOR needs — validating one typed extension, classifying it
 * against what already exists, and describing it. Two definitions of "a valid extension" would drift,
 * and the drift would be invisible.
 */

/**
 * Extensions that cannot be added OR enabled.
 *
 * The list filters a file picker and validates in JS, so it is not a security boundary and this is not
 * protection — see spec §3. It stops a well-meaning admin from advertising a document repository as a
 * place to share executables, which is hard to un-decide once files exist.
 *
 * THE RULE MUST COVER ENABLING, NOT JUST ADDING. Guarding only the add path leaves the route that
 * matters unguarded: a column already carrying `.exe` from a hand edit or a migration.
 *
 * Archives are deliberately absent (client, 2026-08-13). A `.zip` can carry any of these, which makes
 * the whole list advisory — recorded, not enforced.
 */
export const BLOCKED_TYPES: readonly string[] = Object.freeze([
  ".exe", ".com", ".bat", ".cmd", ".msi", ".dll", ".scr", ".pif", ".cpl",
  ".ps1", ".psm1", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh", ".hta",
  ".reg", ".lnk", ".jar", ".app", ".sh",
]);

/**
 * Extension -> human description, for the Description column.
 *
 * DERIVED, never stored: the Choice column holds only the extension string, so there is nowhere to put
 * a description. An extension missing from this map is also the page's typo signal — a row meant to be
 * Excel that reads "Custom file type" is its own warning.
 */
const KNOWN_TYPES: { readonly [ext: string]: string } = {
  ".pdf": "Portable Document Format",
  ".doc": "Word 97-2003 Document",
  ".docx": "Microsoft Word Document",
  ".dot": "Word 97-2003 Template",
  ".dotx": "Word Template",
  ".rtf": "Rich Text Format",
  ".txt": "Plain Text",
  ".xls": "Microsoft Excel 97-2003 Workbook",
  ".xlsx": "Microsoft Excel Workbook",
  ".xlsm": "Excel Macro-Enabled Workbook",
  ".xlsb": "Excel Binary Workbook",
  ".csv": "Comma-Separated Values",
  ".ppt": "PowerPoint 97-2003 Presentation",
  ".pptx": "Microsoft PowerPoint Presentation",
  ".png": "PNG Image",
  ".jpg": "JPEG Image",
  ".jpeg": "JPEG Image",
  ".gif": "GIF Image",
  ".bmp": "Bitmap Image",
  ".tif": "TIFF Image",
  ".tiff": "TIFF Image",
  ".heic": "HEIC Image",
  ".svg": "Scalable Vector Graphics",
  ".msg": "Outlook Message",
  ".eml": "Email Message",
  ".zip": "Compressed Archive",
  ".7z": "7-Zip Archive",
  ".rar": "RAR Archive",
  ".xml": "XML Document",
  ".json": "JSON Document",
  ".mp4": "MP4 Video",
  ".mov": "QuickTime Video",
  ".dwg": "AutoCAD Drawing",
  ".dxf": "AutoCAD Exchange Drawing",
};

/** Shown for anything not in KNOWN_TYPES. */
export const CUSTOM_TYPE_DESCRIPTION = "Custom file type";

export function describeExtension(ext: string): string {
  return KNOWN_TYPES[(ext ?? "").trim().toLowerCase()] ?? CUSTOM_TYPE_DESCRIPTION;
}

export function isBlockedType(ext: string): boolean {
  return BLOCKED_TYPES.indexOf((ext ?? "").trim().toLowerCase()) !== -1;
}

/** A short badge for the row, e.g. `.xlsx` -> `XLSX`. Used instead of icon assets. */
export function badgeFor(ext: string): string {
  return (ext ?? "").replace(/^\./, "").toUpperCase().slice(0, 4);
}

export type ExtensionInput =
  | { ok: true; ext: string }
  | { ok: false; reason: string };

/**
 * trim -> lowercase -> force ONE leading dot.
 *
 * The dot is enforced, not required, because a dotless token is the single input with two opposite
 * behaviours: the file picker parses `png` as a MIME type, finds no `/`, and silently drops it —
 * greying that type out of the dialog — while the JS validator's endsWith check accepts it. One typo,
 * half a day (2026-07-30, CLAUDE.md gotcha #10).
 */
export function normalizeExtension(raw: string): ExtensionInput {
  const trimmed = (raw ?? "").trim().toLowerCase();
  if (trimmed.length === 0) {
    return { ok: false, reason: "Type a file extension, for example .pdf" };
  }
  if (/\s/.test(trimmed)) {
    return { ok: false, reason: "A file extension cannot contain spaces." };
  }
  const body = trimmed.charAt(0) === "." ? trimmed.slice(1) : trimmed;
  if (body.length === 0) {
    return { ok: false, reason: "Type a file extension, for example .pdf" };
  }
  if (body.indexOf(".") !== -1) {
    return {
      ok: false,
      reason: "Enter one extension only, without a second dot — for example .docx",
    };
  }
  if (!/^[a-z0-9]+$/.test(body)) {
    return { ok: false, reason: "A file extension can only contain letters and numbers." };
  }
  return { ok: true, ext: `.${body}` };
}

/**
 * Optimal string alignment distance, bailing out as soon as it exceeds `limit`.
 *
 * NOT plain Levenshtein, and the difference is the whole feature: Levenshtein scores an adjacent SWAP
 * as two substitutions, so `.xlxs` sits distance 2 from `.xlsx` and the motivating suggestion could
 * never fire. Transposition is the commonest typing error, so it counts as one edit here.
 */
function distanceWithin(a: string, b: string, limit: number): number {
  if (Math.abs(a.length - b.length) > limit) return limit + 1;
  // Three rows: the one before last (for transposition), the last, and the one being built.
  let beforePrev: number[] = [];
  let prev: number[] = [];
  for (let j = 0; j <= b.length; j++) prev.push(j);
  for (let i = 1; i <= a.length; i++) {
    const row: number[] = [i];
    let best = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      let value = Math.min(row[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
      if (
        i > 1 &&
        j > 1 &&
        a.charAt(i - 1) === b.charAt(j - 2) &&
        a.charAt(i - 2) === b.charAt(j - 1)
      ) {
        value = Math.min(value, beforePrev[j - 2] + 1);
      }
      row.push(value);
      if (value < best) best = value;
    }
    if (best > limit) return limit + 1;
    beforePrev = prev;
    prev = row;
  }
  return prev[b.length];
}

/**
 * The known extension one edit away from `ext`, if there is exactly one.
 *
 * USABILITY, NOT SECURITY. A near miss fails CLOSED — `.xlxs` blocks a legitimate file type, which is
 * annoying and never dangerous. It earns its place by stopping an admin from believing a policy is in
 * force when it is not: the same failure mode as a missing column.
 *
 * Ties prefer a candidate of the SAME LENGTH, because that is a substitution or transposition — how
 * extensions are actually mistyped. Without this the motivating case suppresses itself: `.xlxs` is one
 * edit from `.xlsx` (transposed) AND from `.xls` (dropped character), so the suggestion worth making
 * would never appear.
 *
 * A tie that survives that rule resolves to NO suggestion. `.dox` sits one substitution from both
 * `.doc` and `.dot`, and guessing between them would put a wrong extension one click away.
 */
export function nearMiss(ext: string): string | undefined {
  const target = (ext ?? "").trim().toLowerCase();
  if (KNOWN_TYPES[target] !== undefined) return undefined;
  const hits: string[] = [];
  Object.keys(KNOWN_TYPES).forEach((known) => {
    if (distanceWithin(target, known, 1) <= 1) hits.push(known);
  });
  if (hits.length === 0) return undefined;
  const sameLength = hits.filter((h) => h.length === target.length);
  if (sameLength.length === 1) return sameLength[0];
  if (sameLength.length > 1) return undefined;
  return hits.length === 1 ? hits[0] : undefined;
}

export type AdditionVerdict =
  | { kind: "invalid"; reason: string }
  | { kind: "blocked"; ext: string }
  | { kind: "exists-enabled"; ext: string }
  | { kind: "exists-disabled"; ext: string }
  | { kind: "new"; ext: string; suggestion?: string };

/**
 * What "Add" should do with a typed extension.
 *
 * The two `exists-*` verdicts are why this exists rather than a bare `indexOf`: adding a DUPLICATE
 * choice is the bad outcome. SharePoint holds two `.docx` entries happily, the page renders two
 * identical rows with independent toggles, and one of them appears to do nothing.
 *
 * `exists-disabled` does NOT enable it (client, 2026-08-13) — the page reports the truth and leaves the
 * decision to the admin.
 */
export function classifyAddition(
  raw: string,
  choices: readonly string[],
  ticked: readonly string[],
): AdditionVerdict {
  const parsed = normalizeExtension(raw);
  if (!parsed.ok) return { kind: "invalid", reason: parsed.reason };
  const ext = parsed.ext;
  if (isBlockedType(ext)) return { kind: "blocked", ext };

  const lower = (list: readonly string[]): string[] =>
    list.map((c) => (c ?? "").trim().toLowerCase());
  if (lower(choices).indexOf(ext) !== -1) {
    return lower(ticked).indexOf(ext) !== -1
      ? { kind: "exists-enabled", ext }
      : { kind: "exists-disabled", ext };
  }
  const suggestion = nearMiss(ext);
  return suggestion ? { kind: "new", ext, suggestion } : { kind: "new", ext };
}

export type PanelState = "no-column" | "all-blocked" | "normal";

/**
 * Which of the three screens to render.
 *
 * `no-column` and `all-blocked` look identical in a list view and mean opposite things — a silent
 * default versus a deliberate block — so they must never collapse into each other. Only the PRESENCE
 * of the column separates them; an emptied multi-Choice column returns `null`, not `[]` (gotcha #11).
 *
 * Choices do not enter into it: a column with choices and nothing ticked is `all-blocked`, because what
 * blocks an upload is the ticked value, not the availability of a choice.
 */
export function panelState(columnPresent: boolean, ticked: readonly string[]): PanelState {
  if (!columnPresent) return "no-column";
  return ticked.length === 0 ? "all-blocked" : "normal";
}

/**
 * Blocked extensions that are currently ALLOWED.
 *
 * Reported loudly rather than hidden: this state cannot be produced on this page, so it arrived from a
 * hand edit or a migration and nobody chose it deliberately here.
 */
export function blockedButTicked(ticked: readonly string[]): string[] {
  const out: string[] = [];
  ticked.forEach((t) => {
    const ext = (t ?? "").trim().toLowerCase();
    if (isBlockedType(ext) && out.indexOf(ext) === -1) out.push(ext);
  });
  return out;
}
