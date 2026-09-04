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
/**
 * The fixed fields that describe the DESTINATION rather than the document.
 *
 * `Year` and `Document Type` are folder tiers — the built-in below-Unit pair — so on the upload form
 * they belong to the BATCH, alongside the permissioned tiers, and are chosen once for every file
 * going to that folder. Kept separate from the per-file fields for exactly that reason: the read-only
 * batch view in My Submissions renders them under the destination, where the uploader set them.
 *
 * They have no Tid twin (they are managed metadata), which is why `discoverTierFields` cannot find
 * them and they must be listed. That absence is load-bearing, not an oversight.
 */
export const BATCH_FIXED_FIELDS: Array<{ label: string; field: string }> = [
  { label: "Document Type", field: "Document_x0020_Type" },
  { label: "Year", field: "Year" },
];

/**
 * The fixed fields the uploader fills in PER FILE, in the order the client's own library views use
 * (memory `dms-staging-column-order`): what it is for, then how sensitive, then free text.
 *
 * Listed by INTERNAL NAME only — the spelling variants are derived, so there is nothing to mistype.
 */
export const FILE_FIXED_FIELDS: Array<{ label: string; field: string }> = [
  { label: "Document Date", field: "DocumentDate" },
  { label: "Confidentiality", field: "Confidentiality_x0020_Level" },
  { label: "Legally Privileged", field: "LegallyPrivileged" },
  { label: "Project Name", field: "ProjectName" },
  { label: "Vendor / Customer", field: "Vendor_x002f_CustomerName" },
  { label: "Details", field: "_ExtendedDescription" },
  { label: "Remark", field: "Remark" },
];

/**
 * Every fixed field, destination first — the order the single-document panel has always used.
 *
 * ⚠ DERIVED from the two halves, never listed again. A second copy is how the batch view and the
 * detail panel would come to disagree about what a document carries, and the drifting one would be
 * whichever is looked at less.
 */
export const FIXED_FIELDS: Array<{ label: string; field: string }> =
  BATCH_FIXED_FIELDS.concat(FILE_FIXED_FIELDS);
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

/** One tier of a document's path, with the term GUID behind the label. */
export interface TierValue {
  /** The readable tier name — `Department`, `Estate Mill`. */
  label: string;
  /** The term's label, as filed. */
  value: string;
  /** The term's GUID, from the `<Base>Tid` twin. Blank when the pair was written incomplete. */
  guid: string;
}

/** Where a document sits, derived from its own fields. */
export interface DocumentUnit {
  /** The business segment's label. */
  segment: string;
  /** The deepest tier's label — the unit, on a segment whose tiers stop at Unit. */
  unit: string;
  /** The deepest tier's term GUID. */
  unitTermGuid: string;
  /** Every tier that carries a GUID, shallowest first. */
  tiers: TierValue[];
}

/**
 * The tier path of one document, GUIDs included.
 *
 * Built for the request workflow, which has to route a request to whoever approves for the unit — and
 * the only durable key for that is the unit TERM GUID, since labels get renamed and folder names are
 * abbreviations. `FieldValuesAsText` already carries it in the `<Base>Tid` twin, so this costs no
 * extra request.
 *
 * IT RETURNS THE WHOLE CHAIN, not just the deepest tier, and that is deliberate. Nothing in an item's
 * fields says which tiers are PERMISSIONED: a below-Unit tier such as SubUnit is created with a Tid
 * column exactly like a permissioned one, so "deepest" can be a tier no group is ever mapped to.
 * Routing on it would put the request in nobody's queue, with nothing on screen to say so. The caller
 * matches the chain against the Group Map and takes the deepest tier an approver actually exists for;
 * `unitTermGuid` is the fallback for when that list cannot be read.
 */
export function documentUnit(fieldText: Record<string, string>): DocumentUnit {
  const tiers: TierValue[] = [];
  let segment = "";
  for (const name of discoverTierFields(fieldText)) {
    const value = readField(fieldText, name);
    const guid = readField(fieldText, `${tidBase(name)}Tid`);
    if (tidBase(name) === "BusinessSegment") {
      segment = value;
      continue;
    }
    if (value.length === 0 && guid.length === 0) continue;
    tiers.push({ label: labelFromInternalName(name), value, guid });
  }
  const withGuid = tiers.filter((t) => t.guid.length > 0);
  const deepest = withGuid.length > 0 ? withGuid[withGuid.length - 1] : undefined;
  return {
    segment,
    unit: deepest ? deepest.value : "",
    unitTermGuid: deepest ? deepest.guid : "",
    tiers: withGuid,
  };
}

/**
 * The deepest tier someone is recorded as approving for.
 *
 * `approverUnits` is every unit GUID carrying an `APR` mapping — not one person's. Searched
 * DEEPEST-FIRST, because a document filed three tiers down belongs to the unit at the bottom, and a
 * department-level match would hand the decision to the wrong queue.
 *
 * Returns `undefined` when nothing matches, which the caller must NOT read as "no approver exists": an
 * unreadable Group Map produces exactly the same empty set.
 */
export function routeToApprover(unit: DocumentUnit, approverUnits: string[]): TierValue | undefined {
  const known = (approverUnits ?? []).map((u) => (u ?? "").trim().toLowerCase()).filter((u) => u.length > 0);
  if (known.length === 0) return undefined;
  const tiers = unit?.tiers ?? [];
  for (let i = tiers.length - 1; i >= 0; i--) {
    if (known.indexOf(tiers[i].guid.toLowerCase()) !== -1) return tiers[i];
  }
  return undefined;
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


/**
 * The rows that describe a BATCH's destination: the permissioned tiers, then Year and Document Type.
 *
 * Client, 2026-08-22: opening a batch in My Submissions should look like the upload form they filled
 * in — the folder chosen once at the top, then each file's own details beneath it. This is the top
 * half. Blanks are dropped for the same reason as everywhere else here: a library serving twelve
 * segments leaves most tier columns empty on any given document.
 *
 * Taken from ONE file in the batch, which is sound because a batch IS one destination folder — every
 * file in it was filed under the same tiers by construction.
 */
export function buildBatchRows(fieldText: Record<string, string>): DetailRow[] {
  const ft = fieldText ?? {};
  return [
    ...tierRows(ft),
    ...BATCH_FIXED_FIELDS
      .map((f) => ({ label: f.label, value: readField(ft, f.field) }))
      .filter((r) => r.value.length > 0),
  ];
}

/**
 * The rows an uploader filled in for ONE file — everything the batch does not decide.
 *
 * The lower half of the read-only batch view. `trailing` carries rows only the calling screen knows
 * (file size, last updated) and passes through blank-or-not, as in `buildDetailRows`.
 */
export function buildFileRows(args: {
  fieldText: Record<string, string>;
  trailing?: DetailRow[];
}): DetailRow[] {
  const ft = args.fieldText ?? {};
  return [
    ...FILE_FIXED_FIELDS
      .map((f) => ({ label: f.label, value: readField(ft, f.field) }))
      .filter((r) => r.value.length > 0),
    ...(args.trailing ?? []),
  ];
}
