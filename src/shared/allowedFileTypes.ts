/**
 * Allowed upload file types, resolved from the `AllowedFileTypes` multi-select
 * Choice column on the `DMS Config` list.
 *
 * This is a discriminated union rather than a bare `string[]` on purpose. Three
 * outcomes have to stay distinguishable, because different people fix them:
 *
 *   configured — the client ticked types. Normal operation.
 *   none       — the client unticked everything. Hard block; the CLIENT fixes it.
 *   unknown    — the config could not be read, or the column does not exist on
 *                this site. Fall back to `FALLBACK_FILE_TYPES`; an ADMIN fixes it.
 *
 * Collapsing `none` and `unknown` into an empty array is the specific mistake
 * this type exists to prevent. See
 * docs/superpowers/specs/2026-07-30-allowed-file-types-dropdown-design.md §3, §6.
 */
export type AllowedFileTypes =
  | { kind: "configured"; types: string[] }
  | { kind: "none" }
  | { kind: "unknown"; types: string[] };

/**
 * Last-resort list used only when `DMS Config` cannot be read at all. Kept equal
 * to the column's configured choices so that even the fallback agrees with the
 * dropdown (spec §7).
 */
export const FALLBACK_FILE_TYPES: string[] = [
  ".pdf",
  ".doc",
  ".docx",
  ".xls",
  ".xlsx",
];

/**
 * trim -> lowercase -> prepend "." if missing -> drop empties -> de-duplicate.
 *
 * Defence in depth. The dropdown removes typos from the common path, but an
 * owner editing the column's *Choices* still types the extension by hand, so a
 * fumbled dot is repaired here instead of silently breaking the picker.
 */
export function normalizeFileTypes(raw: readonly string[]): string[] {
  const out: string[] = [];
  raw.forEach((entry) => {
    const trimmed = (entry ?? "").trim().toLowerCase();
    if (trimmed.length === 0) return;
    const dotted = trimmed.charAt(0) === "." ? trimmed : `.${trimmed}`;
    // indexOf, not includes: the SPFx tsconfig does not target ES2017 libs
    // (CLAUDE.md gotcha #3).
    if (out.indexOf(dotted) === -1) out.push(dotted);
  });
  return out;
}

/**
 * Reads a SharePoint multi-value Choice field, tolerating both JSON shapes:
 * a plain array (odata=nometadata / minimalmetadata) and `{ results: [...] }`
 * (odata=verbose).
 *
 * Returns `undefined` only when the field is genuinely absent. An empty
 * selection returns `[]`, which callers must treat as "nothing ticked" rather
 * than "not configured".
 */
export function readChoiceArray(value: unknown): string[] | undefined {
  if (Array.isArray(value)) return value as string[];
  if (value !== null && typeof value === "object") {
    const wrapped = (value as { results?: unknown }).results;
    if (Array.isArray(wrapped)) return wrapped as string[];
  }
  return undefined;
}

/**
 * Shown when `AllowedFileTypes` is present but nothing is ticked. Names the
 * column and the list because the client is the person who fixes it, in one
 * click, and a vague message would send them to us instead.
 */
export const NO_TYPES_MESSAGE =
  "No file types are configured for upload. Ask your DMS administrator to set Allowed File Types in DMS Config.";

/**
 * Shown when `DMS Config` could not be read at all. Deliberately different from
 * NO_TYPES_MESSAGE: this one is an admin or permissions problem, and uploads
 * continue on built-in defaults rather than stopping.
 */
export const CONFIG_UNREADABLE_MESSAGE =
  "Could not load the upload configuration — using built-in file types. Ask your DMS administrator to check DMS Config.";

/**
 * Turns the raw `AllowedFileTypes` value into one of the three states.
 *
 * `undefined` (field absent — unprovisioned site, or a read that fell back)
 * becomes `unknown`. An empty or all-blank selection becomes `none`.
 */
export function resolveAllowedFileTypes(
  raw: readonly string[] | undefined,
): AllowedFileTypes {
  if (raw === undefined) {
    return { kind: "unknown", types: normalizeFileTypes(FALLBACK_FILE_TYPES) };
  }
  const types = normalizeFileTypes(raw);
  if (types.length === 0) return { kind: "none" };
  return { kind: "configured", types };
}
