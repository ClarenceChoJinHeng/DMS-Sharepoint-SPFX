import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  SubmissionRecord,
  RECORD_READ_SELECT,
  RECORD_READ_SELECT_NO_ARCHIVE,
  RECORD_READ_SELECT_LEGACY,
  REPLACEMENT_COLUMNS,
  buildRecordPayload,
  parseRecordRow,
} from "./submissionRecords";
import { cachedListTitle, LIST_SUFFIX } from "./naming";

/**
 * The submission record — the write and the read.
 *
 * Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
 *
 * The SPFx half of the pair; `submissionRecords.ts` holds everything testable. Split the same way as
 * `naming.ts` / `spNaming.ts` and `auditLog.ts` / `spAuditLog.ts` — and for a concrete reason here:
 * the pure half is unit-tested, and importing `@microsoft/sp-http` into it would make it untestable,
 * exactly as it does for `PRIMED_SUFFIXES`.
 */

/**
 * ⚠ ALL THREE HEADERS, AND THE THIRD IS THE LOAD-BEARING ONE.
 *
 * JSON light needs no `__metadata` envelope and no entity type, so `ListItemEntityTypeFullName` is
 * deliberately NOT read for this list — do not "restore" it per gotcha #12, which applies to the
 * verbose call sites in `FolderMap.tsx`.
 *
 * SPFx's `SPHttpClient` injects `odata-version: 4.0` on every request, and SharePoint cannot infer
 * the entity set for a JSON-light ENTRY payload under OData 4 — so a row POST returns 400 while
 * `/_api/web/lists` tolerates the identical headers and creates a list quite happily. That asymmetry
 * is what made this look like three unrelated bugs across 2026-08-13, and it cost the file-type page
 * again on 2026-08-27 (1.0.249.0). An empty value removes the header.
 */
const WRITE_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
  "odata-version": "",
};

// ⚠ NO-CACHE, same reason as `dmsFolderMap.ts`'s folder-permission probes and the Requests/My
// Submissions fix (2026-08-30): a stale response here would misjudge whether a displaced record
// should read Cancelled/Deleted for whoever is currently signed in.
const GET_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
};

const listBase = (siteUrl: string): string =>
  `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.submissions))}')`;

/**
 * Record one uploaded file. **Never throws, and never blocks the upload.**
 *
 * ⚠ THE UPLOAD HAS ALREADY SUCCEEDED BY THE TIME THIS RUNS. A failure here must cost the RECORD and
 * nothing else — the document is uploaded, tagged and routed regardless, and the page simply behaves
 * as it did before the feature existed. The same rule as the `SubmissionId`/`BatchId` stamp and as
 * `writeAudit`: degrading to yesterday's behaviour is always correct.
 *
 * ⚠ AND IT IS NEVER SILENT. It returns false and logs the status and the body. A record gap somebody
 * knows about is worth far more than one nobody does — and on this list the only symptom of a missing
 * row is a file that will one day disappear from its own submission.
 *
 * ⚠ THE CALLER MUST NOT CALL THIS WITHOUT A STAMP. A row whose `SubmissionFileId` never made it onto
 * the document can never be joined back to it, so it would sit on My Submissions as a permanent false
 * "Deleted" — worse than no row at all. Refused here as well as at the call site, because this is the
 * invariant the whole feature rests on.
 */
export async function writeSubmissionRecord(
  sp: SPHttpClient,
  siteUrl: string,
  record: Omit<SubmissionRecord, "itemId">,
): Promise<boolean> {
  const name = record?.fileName ?? "";
  try {
    if ((record?.fileId ?? "").trim().length === 0) {
      console.warn(`[submissions] not recorded (no SubmissionFileId): ${name}`);
      return false;
    }
    // Straight to the POST. There is no entity type to look up under JSON light, so an unprovisioned
    // list simply 404s here — one request instead of two, and no state that can go stale.
    const res: SPHttpClientResponse = await sp.post(
      `${listBase(siteUrl)}/items`,
      SPHttpClient.configurations.v1,
      { headers: WRITE_HEADERS, body: JSON.stringify(buildRecordPayload(record)) },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      /* 404 is the ORDINARY case on a site where reconciliation has not run since this shipped, and
         it is not a defect — the uploader is unaffected. Logged more quietly so it cannot drown the
         failures that do need attention. */
      if (res.status === 404) {
        console.info(`[submissions] list not provisioned yet — "${name}" not recorded (harmless)`);
      } else {
        console.warn(`[submissions] not recorded (HTTP ${res.status}): ${name}. ${body.slice(0, 300)}`);
      }
      return false;
    }
    return true;
  } catch (e) {
    console.warn(`[submissions] not recorded: ${name} — ${(e as Error).message}`);
    return false;
  }
}

/**
 * Every record this person's uploads produced.
 *
 * ⚠ `undefined` MEANS NOT-READ, AND THE CALLER MUST NOT TREAT IT AS `[]`. An unreadable or
 * unprovisioned list is not evidence that nothing was uploaded; My Submissions falls back to showing
 * live rows only, which is exactly its behaviour before this feature. `[]` means read-and-empty,
 * which is the normal state for an uploader whose files all still exist. Empty is not unknown, in the
 * place where the cost is claiming a document was deleted.
 *
 * ⚠ FILTERED ON `AuthorId`, NOT on `UploadedBy`. The record is written in the uploader's own session,
 * so its author IS them — the same mechanism the two library reads already use, so a person sees
 * their own records exactly as they see their own files. An email column would additionally have to
 * survive case and alias differences to be right.
 *
 * ⚠ `$top=5000` DOES NOT LIFT THE 5,000-ITEM LIST THRESHOLD. The filter is per person so the returned
 * set is small, but the LIST will grow past that in normal use — so **`Author` must be indexed on
 * this list**, the same provisioning note `Created By` on `Documents` already carries. A truncated
 * read here omits records rather than inventing them, so it under-reports deletions instead of
 * claiming false ones; that is the safe direction, and it is still worth avoiding.
 */
/**
 * Mark the record of a file that a later upload has just overwritten.
 *
 * Client, 2026-08-28: *"The older submission under My Submission will change to Cancelled status if
 * replaced"* — and, when the wording was queried, *"I know its weird but client want it to be
 * Cancelled"*.
 *
 * ⚠ THIS IS WHY THE REPLACED RECORD IS NOT SIMPLY "DELETED". Overwriting a file destroys the old
 * `SubmissionFileId` stamp — the new upload writes its own — so the displaced record stops resolving
 * and `mergeRecords` would otherwise call it `deleted`, telling the original uploader their document
 * was destroyed when in fact somebody filed a newer version of it.
 *
 * ⚠ **NEVER THROWS, AND NEVER BLOCKS THE UPLOAD.** Same rule as `writeSubmissionRecord`: the
 * replacement has already succeeded by the time this runs. A failure costs the *label* — the record
 * falls back to reading `deleted`, which is what it did before this existed — and nothing else.
 *
 * ⚠ NOT FILTERED BY AUTHOR. The person replacing a file is very often NOT the person who uploaded
 * it, so the row being updated belongs to somebody else. That is only possible because the
 * reconciliation ACL pass grants the uploader / approver / HoD groups the `CRS Request` level
 * (Add + Edit + View, no Delete) on this list — the same coarse grant already accepted for
 * `CRS Requests`, and the reason the exposure note in that design applies here too.
 */
export async function markRecordReplaced(
  sp: SPHttpClient,
  siteUrl: string,
  fileId: string,
  actor: string,
): Promise<boolean> {
  const id = (fileId ?? "").trim();
  try {
    if (id.length === 0) return false;

    /* Find the displaced record by its stamp. `$top=5` rather than 1: one stamp should identify one
       row, and if it somehow named two, silently updating whichever came first would leave the other
       reading `deleted` for ever. Cheap to be certain. */
    const find: SPHttpClientResponse = await sp.get(
      `${listBase(siteUrl)}/items?$select=Id&$filter=SubmissionFileId eq '${encodeURIComponent(id.replace(/'/g, "''"))}'&$top=5`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!find.ok) {
      console.info(`[submissions] replaced record not found (HTTP ${find.status}) for ${id}`);
      return false;
    }
    const found = ((await find.json()).value ?? []) as Array<{ Id?: number }>;
    if (found.length === 0) {
      // Ordinary on a site where the file predates the record feature. Not a defect.
      console.info(`[submissions] no record for the replaced file (${id}) — nothing to mark`);
      return false;
    }

    const body: Record<string, string> = {};
    // ISO, never `M/D/YYYY`. That locale format belongs to `validateUpdateListItem`; a plain items
    // MERGE goes through the OData layer, which answers a locale string with "Cannot convert a
    // primitive value to the expected type 'Edm.DateTime'" — a 400 naming the type but not the field.
    body[REPLACEMENT_COLUMNS[0]] = new Date().toISOString();
    body[REPLACEMENT_COLUMNS[1]] = (actor ?? "").trim();

    let ok = true;
    for (const row of found) {
      if (typeof row.Id !== "number") continue;
      const upd: SPHttpClientResponse = await sp.post(
        `${listBase(siteUrl)}/items(${row.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: { ...WRITE_HEADERS, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
          body: JSON.stringify(body),
        },
      );
      if (!upd.ok) {
        ok = false;
        /* A 400 here is the columns not existing — a site that has not reconciled since they
           shipped. Logged quietly, because the outcome is simply the old behaviour: the record
           reads `deleted` instead of `cancelled`, and no uploader is blocked. */
        const detail = await upd.text().catch(() => "");
        console.info(
          `[submissions] could not mark ${id} replaced (HTTP ${upd.status}). ${detail.slice(0, 200)}`,
        );
      }
    }
    return ok;
  } catch (e) {
    console.info(`[submissions] could not mark ${id} replaced — ${(e as Error).message}`);
    return false;
  }
}

export async function readSubmissionRecords(
  sp: SPHttpClient,
  siteUrl: string,
  userId: number,
): Promise<SubmissionRecord[] | undefined> {
  try {
    if (!(userId > 0)) return undefined;
    const ask = async (select: string): Promise<SPHttpClientResponse> =>
      sp.get(
        `${listBase(siteUrl)}/items` +
          `?$select=${select}` +
          `&$filter=AuthorId eq ${userId}` +
          `&$orderby=Id desc&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: GET_HEADERS },
      );

    let res: SPHttpClientResponse = await ask(RECORD_READ_SELECT);
    /* ⚠ A THREE-RUNG LADDER SINCE 2026-09-03, AND THE MIDDLE RUNG IS THE POINT. `ArchivedAt` arrived
       after `ReplacedAt`/`ReplacedBy`, so a site carrying only the older pair must lose ONLY the new
       column — dropping straight to the legacy select would take the replacement state with it, and
       every replaced file there would silently go back to reading "Deleted". `RevokedBy` taught this
       on `CRS Requests` (2026-08-30): one retry is not enough once a second optional column exists. */
    if (res.status === 400) {
      console.info("[submissions] no ArchivedAt column on this list — reading without it");
      res = await ask(RECORD_READ_SELECT_NO_ARCHIVE);
    }
    /* ⚠ RETRIED WITHOUT THE 2026-08-28 COLUMNS ON **400 ONLY**, and the status is the whole rule.
       400 means the list does not have those columns — a site that has not reconciled since they
       shipped — and one unknown field name fails the WHOLE $select, so asking unconditionally would
       take EVERY record off this page rather than merely the replacement state.
       A 404 is the LIST missing and a 403 is the ACL pass not having run; retrying either asks the
       same unanswerable question twice and reports an unprovisioned list as an unreadable one. */
    if (res.status === 400) {
      console.info("[submissions] no replacement columns on this list — reading without them");
      res = await ask(RECORD_READ_SELECT_LEGACY);
    }
    if (!res.ok) {
      /* Reported once, and NOT as an error the uploader has to act on. A 404 is a site that has not
         reconciled since this shipped; a 400 is a list missing a column, which reconciliation
         repairs; a 403 is the ACL pass not having run. All three degrade the page identically and
         none of them is the uploader's problem. */
      console.info(`[submissions] records not read (HTTP ${res.status}) — showing live files only`);
      return undefined;
    }
    const data = await res.json();
    return ((data.value ?? []) as Array<Record<string, unknown>>).map((r) => parseRecordRow(r));
  } catch (e) {
    console.info(`[submissions] records not read — showing live files only (${(e as Error).message})`);
    return undefined;
  }
}
