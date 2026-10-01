/**
 * Is the column a new folder level would bind to safe to bind to?
 *
 * ⚠ THE DANGER IS NEVER "THE COLUMN EXISTS" — IT IS "IT EXISTS AS SOMETHING ELSE", and getting that
 * distinction wrong in either direction breaks something real.
 *
 * `ensureTextColumn` SKIPS a column that already exists **whatever its type**, so a level whose name
 * derives an existing column silently binds to that column. Nothing errors, the structure saves, and
 * it surfaces later as a whole segment that has stopped tagging.
 *
 * The obvious guard — refuse when the column exists — would break three things that are supposed to
 * work, which is why this checks the TYPE instead:
 *
 *   1. **Re-adding `Year` or `Document Type`.** Those columns exist and MUST be reused;
 *      `builtInTierFor` restores their exact shape on purpose.
 *   2. **Re-adding a level that was removed.** Columns are never deleted (deleting takes the data
 *      with it and does NOT reach the recycle bin), so the column outlives the level — and
 *      remove-and-re-add is the ONLY way to edit a level, since editing in place is not supported.
 *   3. **Two segments sharing a level name.** Tier columns are shared: `Department` and `Unit` are
 *      one column across every segment, so a second segment adding `State` SHOULD bind to the
 *      existing one.
 *
 * So a plain-text column is a legitimate share and is allowed. Anything else is refused.
 *
 * Client, 2026-09-06, on being shown the silent bind: *"Why not jst.... put an error? force them to
 * create a new column with new name instead of binding to the same one if it already exist?"* — this
 * is that, narrowed to the cases where binding is actually wrong.
 *
 * No `@microsoft/*` imports, which is what keeps this covered by plain Jest — the same constraint
 * that keeps `PRIMED_SUFFIXES` in `naming.ts` rather than `spNaming.ts`.
 */

/**
 * Columns that already do a job, keyed by INTERNAL name.
 *
 * ⚠ THESE ARE NOT TIER COLUMNS AND NEVER SHOULD BE. A level deriving one of these would put two
 * writers in one column: the tier writer stamping a folder name, and the upload form stamping what
 * the uploader typed. The last write wins, and neither writer is wrong from where it stands.
 *
 * Keyed on the INTERNAL name, because that is what a derived name is compared against —
 * `columnNameFor` strips illegal characters, so "Vendor/Customer Name" derives `VendorCustomerName`
 * while the real column is `Vendor_x002f_CustomerName`. BOTH spellings are listed for exactly that
 * reason: a level called "Vendor/Customer Name" must be refused even though its derived name differs
 * from the column's own.
 */
export const RESERVED_TIER_COLUMNS: Readonly<Record<string, string>> = {
  projectname: "the Project Name typed on the upload form",
  remark: "the Remark typed on the upload form",
  keyword: "the Keyword typed on the upload form",
  vendorcustomername: "the Vendor/Customer Name typed on the upload form",
  vendor_x002f_customername: "the Vendor/Customer Name typed on the upload form",
  documentdate: "the Document Date typed on the upload form",
  legallyprivileged: "the Legally Privileged marker",
  confidentialitylevel: "the Confidentiality Level",
  confidentiality_x0020_level: "the Confidentiality Level",
  fullname: "the term's full name, shown on folders",
  full_x0020_name: "the term's full name, shown on folders",
  archived: "the seven-year archive marker",
  submissionid: "the submission reference",
  batchid: "the set reference",
  submissionfileid: "the per-file submission stamp",
  bulkimport: "the bulk-import auto-approve marker",
  approvedby: "who approved the document",
};

/**
 * SharePoint field types a tier column may safely be.
 *
 * ONLY plain text. A tier stores a folder name and, for a shared-list tier, a term id in its `Tid`
 * twin — both plain strings. Everything else is another kind of column wearing the same name.
 *
 * ⚠ `Note` IS EXCLUDED DELIBERATELY though it also holds text: `Remark` and `ItemPath` are Note
 * columns, and a multi-line field is never what a folder level wants.
 */
const SAFE_FIELD_TYPES: Readonly<Record<string, true>> = { text: true };

/** How a name is compared: case-insensitive, trimmed. */
function key(name: string): string {
  return (name ?? "").trim().toLowerCase();
}

/**
 * Why this derived column name must not be used, or `undefined` when it is fine.
 *
 * Checked BEFORE any request, because it needs none — the answer is the same on every site.
 */
export function reservedColumnReason(internalName: string): string | undefined {
  const job = RESERVED_TIER_COLUMNS[key(internalName)];
  if (job === undefined) return undefined;
  return (
    `"${internalName}" is already used for ${job}. Two things would write to one column and the ` +
    `last one would win. Give the level a different name.`
  );
}

/**
 * Why binding to an EXISTING column of this type must be refused, or `undefined` when it is safe.
 *
 * @param typeAsString SharePoint's own `TypeAsString`, e.g. `Text`, `TaxonomyFieldType`, `Note`.
 *   Pass `undefined` when the column does not exist — a NEW column is always safe, so that answers
 *   `undefined` too.
 *
 * ⚠ AN UNREADABLE FIELD LIST MUST NOT REACH HERE AS `undefined`. "Not read" and "not there" are
 * different facts and only one of them is safe; the caller has to keep them apart, because this
 * function cannot. See the call site in `StructureManager`.
 */
export function existingColumnReason(
  internalName: string,
  typeAsString: string | undefined,
): string | undefined {
  if (typeAsString === undefined) return undefined;
  if (SAFE_FIELD_TYPES[key(typeAsString)]) return undefined;
  return (
    `A column called "${internalName}" already exists on this site as ${typeAsString}, and a folder ` +
    `level needs plain text. Binding to it would not fail — it would quietly stop this segment ` +
    `tagging documents. Give the level a different name.`
  );
}

/**
 * The one call the screen makes: reserved name first, then the live type.
 *
 * Reserved is checked first so the answer is instant and identical everywhere — those names are
 * wrong on any site, whether or not the column happens to have been created yet.
 */
export function tierColumnRefusal(
  internalName: string,
  typeAsString: string | undefined,
): string | undefined {
  return (
    reservedColumnReason(internalName) ??
    existingColumnReason(internalName, typeAsString)
  );
}
