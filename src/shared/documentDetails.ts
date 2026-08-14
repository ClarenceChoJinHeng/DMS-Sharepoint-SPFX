/**
 * The metadata panel of a document detail view — which rows, in which order, under which labels.
 *
 * Spec: docs/superpowers/specs/2026-08-14-document-details-panel-design.md
 *
 * Pure and SPFx-free. It exists because both detail views (My Submissions, and the approver's
 * ApprovalDocument) each carried a HARDCODED list of fields naming `Department` and `Unit`
 * literally — so on Upstream Operations Malaysia, whose tiers are `Region` and `Estate/Mill`, both
 * rows read blank, blank rows are dropped, and the two values that decide where the file lives
 * vanished from the panel. Found live 2026-08-14 on a file filed at
 * `UPOPSMY › JHR › BKB › 2024 › Working File`, where the panel showed the segment and nothing below
 * it. Every segment onboarded from here has different tier names, so a list of names could only ever
 * be right for the segments it was written for.
 *
 * The fix is to DERIVE the tier rows. Two facts make that possible with no extra request:
 *   1. Every tier column has a `<Base>Tid` sibling holding the term GUID — created that way by
 *      SegmentCreator and StructureManager. So a field is a tier field iff its Tid twin is present.
 *   2. `FieldValuesAsText` returns EVERY field on the item, Tid columns included, so the evidence
 *      is already in the response the panel has.
 *
 * `Year` and `Document Type` deliberately have no Tid column — they are managed metadata — which is
 * exactly why they are listed explicitly as fixed rows rather than discovered.
 */

/** One row of the panel. */
export interface DetailRow {
  label: string;
  value: string;
}

/**
 * The key SharePoint's OData layer actually returns for an internal name.
 *
 * It double-encodes underscores in property names, so `Business_x0020_Segment` (where `_x0020_` is
 * a space) comes back as `Business_x005f_x0020_x005f_Segment`. Derived rather than hand-written:
 * both detail views listed the two spellings per field by hand, which is a silently blank row per
 * typo, and is why a stale `Year_x002f_Period` fallback was still being carried long after the
 * client renamed that column.
 */
export function encodedKey(internalName: string): string {
  return (internalName ?? "").replace(/_/g, "_x005f_");
}

/** Both spellings, encoded first — that is the one a live response uses. */
export function fieldKeys(internalName: string): string[] {
  const plain = (internalName ?? "").trim();
  if (plain.length === 0) return [];
  const enc = encodedKey(plain);
  return enc === plain ? [plain] : [enc, plain];
}

/**
 * The sanitized base of an internal name, which is what a Tid column is named after.
 *
 * `Business_x0020_Segment` → `BusinessSegment` (→ `BusinessSegmentTid`). `EstateMill` is already
 * bare because `sanitizeFolderSegment` REMOVES illegal characters rather than substituting one, so
 * `Estate/Mill` yielded the column `EstateMill` and the pair `EstateMillTid`.
 */
export function tidBase(internalName: string): string {
  return (internalName ?? "").replace(/_x[0-9a-fA-F]{4}_/g, "");
}

/**
 * A readable label for an internal name.
 *
 * Encoded characters decode to a SPACE, which recovers the real display name exactly for the
 * columns that have one (`Business_x0020_Segment` → `Business Segment`). A name sanitized at
 * creation cannot be recovered — `EstateMill` was `Estate/Mill` and the slash is gone — so it is
 * split on camel case into `Estate Mill`. Readable and honest; guessing the separator back would put
 * a punctuation mark on screen that nobody typed, and would be wrong for any name that never had one.
 */
export function labelFromInternalName(internalName: string): string {
  const decoded = (internalName ?? "").replace(/_x[0-9a-fA-F]{4}_/g, " ").trim();
  if (decoded.length === 0) return "";
  // Already spaced by decoding — leave it alone rather than re-splitting "Business Segment".
  if (decoded.indexOf(" ") !== -1) return decoded;
  return decoded
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2");
}

/** Read a field under whichever spelling the response used. Blank when absent. */
export function readField(fieldText: Record<string, string>, internalName: string): string {
  for (const k of fieldKeys(internalName)) {
    const v = (fieldText ?? {})[k];
    if (typeof v === "string" && v.trim().length > 0) return v.trim();
  }
  return "";
}

/** True when the response carries this name at all, under either spelling. */
function hasKey(fieldText: Record<string, string>, internalName: string): boolean {
  for (const k of fieldKeys(internalName)) {
    if ((fieldText ?? {})[k] !== undefined) return true;
  }
  return false;
}

/**
 * The tier columns present on this item, in response order.
 *
 * A field qualifies when its `<Base>Tid` twin exists — the signature every tier column carries and
 * one no ordinary column has. So `Region`/`RegionTid` and `EstateMill`/`EstateMillTid` are found on
 * a segment nobody wrote code for, and a tier added by the Structure Manager next year appears with
 * no change here.
 *
 * Tid columns themselves are excluded: they hold GUIDs, which are exactly what the term store exists
 * to keep off the screen.
 *
 * ORDER comes from the response, which follows the library's field order and therefore the order the
 * tiers were created in — chain order, because SegmentCreator creates them walking the chain.
 * Business Segment is pinned first regardless: it is the top of the hierarchy, the one position that
 * is semantically load-bearing rather than cosmetic. If the rest ever came back out of order the
 * cost would be a reordered panel, never a missing value.
 */
export function discoverTierFields(fieldText: Record<string, string>): string[] {
  const out: string[] = [];
  const seenBase: Record<string, true> = {};
  for (const rawKey of Object.keys(fieldText ?? {})) {
    // Work in plain internal names, so the two spellings cannot both be listed.
    const plain = rawKey.replace(/_x005f_/g, "_");
    const base = tidBase(plain);
    if (base.length === 0) continue;
    if (/Tid$/.test(base)) continue;
    if (seenBase[base]) continue;
    if (!hasKey(fieldText, `${base}Tid`)) continue;
    seenBase[base] = true;
    out.push(plain);
  }
  // Pinned, not sorted: everything else keeps response order.
  const isSegment = (n: string): boolean => tidBase(n) === "BusinessSegment";
  return out.filter(isSegment).concat(out.filter((n) => !isSegment(n)));
}

/** The tier rows, blanks dropped. A tier with no value does not apply to this document. */
export function tierRows(fieldText: Record<string, string>): DetailRow[] {
  return discoverTierFields(fieldText)
    .map((name) => ({ label: labelFromInternalName(name), value: readField(fieldText, name) }))
    .filter((r) => r.value.length > 0);
}

/**
 * The fixed fields, in the order the client's own library views use (memory
 * `dms-staging-column-order`): what it is, then who it is for, then how sensitive, then free text.
 *
 * Fixed because none of them is a tier: `Year` and `Document Type` are managed metadata with no Tid
 * twin, and the rest are plain columns. Listed by INTERNAL NAME only — the spelling variants are
 * derived, so there is nothing here to mistype.
 */
export const FIXED_FIELDS: Array<{ label: string; field: string }> = [
  { label: "Document Type", field: "Document_x0020_Type" },
  { label: "Year", field: "Year" },
  { label: "Document Date", field: "DocumentDate" },
  { label: "Confidentiality", field: "Confidentiality_x0020_Level" },
  { label: "Legally Privileged", field: "LegallyPrivileged" },
  { label: "Project Name", field: "ProjectName" },
  { label: "Vendor / Customer", field: "Vendor_x002f_CustomerName" },
  { label: "Details", field: "_ExtendedDescription" },
  { label: "Remark", field: "Remark" },
];

/**
 * Every metadata row for one document, in reading order: where it lives, then what it is.
 *
 * `leading` and `trailing` are the caller's own rows — the folder path, the file size, whatever that
 * screen knows and this module does not. Parameters rather than fields, so this module never needs to
 * know which screen it is on.
 *
 * Blank values are DROPPED throughout. On a library serving twelve segments most columns are empty
 * for any given document, and a panel padded with em dashes for `Region` on a Group Head Office file
 * would bury the six rows that mean something. Caller-supplied rows pass through as given: a screen
 * that says "unknown" means it.
 */
export function buildDetailRows(args: {
  fieldText: Record<string, string>;
  leading?: DetailRow[];
  trailing?: DetailRow[];
}): DetailRow[] {
  const ft = args.fieldText ?? {};
  const fixed = FIXED_FIELDS
    .map((f) => ({ label: f.label, value: readField(ft, f.field) }))
    .filter((r) => r.value.length > 0);
  return [...(args.leading ?? []), ...tierRows(ft), ...fixed, ...(args.trailing ?? [])];
}

/**
 * A byte count as text.
 *
 * `File/Length` is a STRING of bytes, so a panel that shows it raw shows `1483776`. Formatted here
 * rather than trusting `FileSizeDisplay`, which SharePoint only sometimes returns.
 */
export function formatBytes(raw: string | number | undefined): string {
  const text = (raw ?? "").toString().trim();
  if (text.length === 0) return "";
  const n = Number(text);
  if (!isFinite(n) || n < 0) return "";
  if (n < 1024) return `${n} B`;
  const units = ["KB", "MB", "GB", "TB"];
  let v = n / 1024;
  let i = 0;
  while (v >= 1024 && i < units.length - 1) { v = v / 1024; i++; }
  // One decimal below 10, none above: "1.4 MB" is useful, "847 KB" is, "847.3 KB" is noise.
  return `${v < 10 ? Math.round(v * 10) / 10 : Math.round(v)} ${units[i]}`;
}
