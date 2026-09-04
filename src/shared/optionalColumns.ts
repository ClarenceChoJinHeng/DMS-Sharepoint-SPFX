import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/**
 * Is it safe to write these columns to this library?
 *
 * ⚠ THE WHOLE REASON THIS EXISTS. `validateUpdateListItem` fails the ENTIRE call on one unknown field
 * name (gotcha #4) — every column lost, not just the missing one. So a column added after a site was
 * provisioned may not be written blindly: doing so would strip Remark, Confidentiality, the legal
 * marker and every tier field off every upload, reported only as the generic "Uploaded, but tagging
 * metadata failed".
 *
 * The rule this enforces: **write only what is confirmed present; degrading to yesterday's behaviour
 * is always correct, losing a document's metadata never is.**
 *
 * ONE implementation, shared by both upload web parts. Two copies deciding whether a field is safe to
 * write is how one of them starts stripping metadata the other preserves.
 *
 * ⚠ FILTERED SERVER-SIDE, NEVER PAGED AND SEARCHED. The first version of this check read
 * `?$select=InternalName&$top=500` and looked through the result — the `fetchAllSiteGroups` defect of
 * 2026-08-21 by another route, and it cost the submission stamp on 2026-08-22: a document library
 * carries hundreds of fields, **a truncated read is indistinguishable from an absent column**, and the
 * newest columns sit at the END of the collection, so they are the first thing a cap removes. The
 * symptom is silent and total. See memory `sp-capped-read-reads-as-absent`.
 *
 * Not to be confused with `spColumns.ts`, which CREATES a column on the admin path. This only ever
 * asks, and never writes schema from an uploader's session.
 *
 * A FAILED read answers FALSE: unknown must not license the write.
 */

/**
 * The optional columns reconciliation asserts on every CRS library, named ONCE.
 *
 * Both upload web parts write them and reconciliation creates them, so a literal in three places is
 * three chances for a typo whose only symptom is a feature that silently never works.
 */
/** One per press of Upload, and one per destination within it — My Submissions groups on these. */
export const REF_COLUMNS: string[] = ["SubmissionId", "BatchId"];
/**
 * The per-FILE reference, and **the join key of the submission record** (2026-08-27).
 * Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
 *
 * ⚠ DELIBERATELY NOT PART OF `REF_COLUMNS`, and that is not tidiness. `libraryHasColumns` demands
 * EVERY name it is given, so folding this in would make a library reconciled before today fail the
 * check for the pair as well — silently switching OFF the grouping that works there now. Two asks,
 * two cache keys, two independent degradations: a library with the pair alone still groups; only a
 * library with all three also gets a submission record.
 *
 * ⚠ THIS EXISTS BECAUSE `UniqueId` DOES NOT SURVIVE APPROVAL. Auto-route is copy-stamp-delete, so
 * the routed file in `Documents` carries a NEW `UniqueId` and the source holding the old one is
 * deleted. A record keyed on `UniqueId` would therefore report every successfully approved document
 * to its own uploader as DELETED. A stamped column survives the copy — which is exactly why the pair
 * above had to exist in all four libraries — so the record is keyed on our data, not SharePoint's
 * identity. See §2 of the spec.
 */
export const SUBMISSION_FILE_COLUMN = "SubmissionFileId";
/**
 * Set by Bulk Upload alone. The auto-approve flow keys its trigger condition on this, so it is the
 * one thing separating a historical import from every other file in the approval library.
 */
export const BULK_IMPORT_COLUMN = "BulkImport";
/**
 * Set by the seven-year archive mover. Spec:
 * docs/superpowers/specs/2026-08-22-seven-year-archive-design.md
 *
 * Not decoration. My Submissions and CRS Search read it to BADGE a row — without it an uploader sees
 * one of their own documents sitting in a library they have never heard of, with nothing to explain
 * why. It is also what lets a deletion or share request be refused from the row already loaded,
 * rather than costing another read: an approver holds only Read on the archive, so approving such a
 * request would fail in their own session AFTER the requester was told it was being handled.
 */
export const ARCHIVED_COLUMN = "Archived";

/**
 * The approver's email, stamped by BOTH approval routes at the moment of approval (2026-09-01).
 *
 * ⚠ EXISTS BECAUSE SHAREPOINT RECORDS NO "APPROVED BY" FIELD, AND `Editor` IS NOT A STAND-IN FOR ONE.
 * Proven live: after `clarencechojinheng` approved a document uploaded by `chocheetuck4` — through
 * BOTH the approval page's `File.Approve()` and the bulk panel's MERGE of `OData__ModerationStatus`
 * — `Editor` still read `chocheetuck4`. Neither approval route is an EDIT, so neither restamps it.
 * `Editor` means "last person to touch this item", and approving never touches it that way.
 *
 * The consequence this closes: `Audit — approval activity` named the uploader as the approver on
 * every single approval, and the QA#5 self-approval suppression — which compares `Author` to
 * `Editor` to catch the one case of someone approving their own upload — matched on EVERY approval,
 * because `Author` and `Editor` are always the uploader. No approver was ever notified. See CLAUDE.md
 * "`Editor` IS NOT THE APPROVER" (2026-09-01) for the full trigger-payload proof.
 *
 * ⚠ ONLY THE TWO APPROVAL-SIDE LIBRARIES NEED IT — `Staging`/`StagingHC`, never `Documents` or the
 * archive. It answers "who approved THIS pending item", which is meaningless once the item has been
 * routed and deleted; Auto-route reads it from the SOURCE item, before the copy, so it never has to
 * survive the move the way `SubmissionFileId` does.
 *
 * ⚠ BLANK MEANS "SEND", NEVER "SUPPRESS". The native Approve/Reject command bypasses the app
 * entirely and leaves this column empty, and so does every document approved before it existed.
 * Reading a blank value as "self-approved" would silently revert to the exact bug this column fixes.
 */
export const APPROVED_BY_COLUMN = "ApprovedBy";

/**
 * Free text the uploader types so a document can be found by a word that is nowhere else on it
 * (client, 2026-09-04: *"Add a new field call Keyword add it in upload form and bulk upload - free
 * text field, it will be used to search throough the home page."*).
 *
 * ⚠ ON EVERY LIBRARY, and for the reason the three columns above already exist everywhere:
 * Auto-route's copy carries over ONLY columns that EXIST at the destination. Absent from the
 * approved side, an uploader's keywords would be stripped from the document the moment it was
 * approved — silently, on a green run — and the archive would hold nothing to search on at all. That
 * is the `Remark`/`LegallyPrivileged` gap of 2026-08-10, and it would be worse here because the whole
 * point of the field is being findable years later.
 *
 * ⚠ SINGULAR. SharePoint's own enterprise keywords field is `TaxKeyword`, displayed *"Enterprise
 * Keywords"*, and some templates carry a `Keywords` column — so the plural risks colliding with a
 * built-in whose type is managed metadata, not text. `ensureColumn` SKIPS a name that already exists
 * WHATEVER ITS TYPE (the standing rule from 1.0.195.0), so a collision would not fail: it would leave
 * a taxonomy field in place and every write of a plain string to it would be refused.
 *
 * ⚠ SEARCHABILITY IS ONLY HALF-ANSWERED BY THIS COLUMN. CRS Search reads the two approval libraries
 * through REST `$filter`, where `substringof` on a text column works immediately; the approved side
 * goes through the Search API, where a column is only queryable once it has a MANAGED PROPERTY — the
 * `<Name>OWSTEXT` assumption that was verified NOT to hold on this tenant (2026-08-23). Free text may
 * still find it through the crawled full-text index, which is untested. See the note at the call site
 * in `DocumentSearch`.
 */
export const KEYWORD_COLUMN = "Keyword";

/** Keyed by library title AND the columns asked about, so two callers wanting different sets both cache. */
const cache: Record<string, boolean> = {};

/** For tests, and for a screen that has just created the columns and must ask again. */
export function clearColumnCache(): void {
  for (const k of Object.keys(cache)) delete cache[k];
}

/** `InternalName eq 'A' or InternalName eq 'B'`. The names are ours, but quote-escape regardless. */
export function columnFilter(names: string[]): string {
  return (names ?? [])
    .map((n) => `InternalName eq '${(n ?? "").replace(/'/g, "''")}'`)
    .join(" or ");
}

export async function libraryHasColumns(
  client: SPHttpClient,
  siteUrl: string,
  libraryTitle: string,
  internalNames: string[],
): Promise<boolean> {
  const wanted = (internalNames ?? []).map((n) => (n ?? "").trim()).filter((n) => n.length > 0);
  // An empty ask is not "everything is present" — it is a caller with nothing to write, and
  // answering true would invite a payload built from an empty list.
  if (wanted.length === 0) return false;

  const key = `${(libraryTitle ?? "").toLowerCase()}|${wanted.join(",").toLowerCase()}`;
  // Captured BEFORE the awaits. The object identity never changes, so mutating this local is the same
  // store — and it satisfies `require-atomic-updates`, which fires on a property write to anything
  // re-read after an await.
  const store = cache;
  const known = store[key];
  if (known !== undefined) return known;

  const fieldsUrl =
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryTitle)}')/fields`;

  /* `undefined` is NOT-READ and `[]` is read-and-empty. The caller below needs the difference to
     decide whether the fallback is worth spending a request on. Empty ≠ unknown, again. */
  const read = async (query: string): Promise<string[] | undefined> => {
    const res: SPHttpClientResponse = await client.get(
      `${fieldsUrl}?${query}`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return undefined;
    const data = await res.json();
    return ((data.value ?? []) as Array<{ InternalName?: string }>)
      .map((f) => (f.InternalName ?? "").toLowerCase());
  };

  let present = false;
  try {
    let names = await read(
      `$select=InternalName&$filter=${encodeURIComponent(columnFilter(wanted))}`,
    );
    /* ⚠ INSURANCE AGAINST THE FIX ITSELF. If `$filter` on `InternalName` were ever rejected, the
       failure would be the ORIGINAL silent one — no write, no error, for ever — and we would have
       replaced one invisible cause with another. So a failed filter re-reads the collection at a cap
       far above any real library rather than concluding the columns are absent. Two requests in a
       case that should never happen, once per library per page. */
    if (names === undefined) names = await read("$select=InternalName&$top=5000");
    const got = names;
    present = got !== undefined && wanted.every((n) => got.indexOf(n.toLowerCase()) > -1);
  } catch {
    present = false;
  }
  store[key] = present;
  return present;
}
