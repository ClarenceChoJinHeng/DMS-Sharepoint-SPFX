# Runbook — the archive movers stamp `ArchivedAt` on the submission record

**Date:** 2026-09-03 · **Code half:** built (1.0.380.0) · **Flow half:** NOT BUILT
**Flows to change:** `CRS — Archive after seven years` **and** its HC clone.

---

## 0. Why this exists

Client, 2026-09-03: *"CLient is complaining why all the archive files consider as deleted, can we
ensure if its archive it shows archive instead?"*

My Submissions builds its rows from the libraries it can read, and **deliberately does not read the
archive** — access narrowed to the two C-Level roles on 2026-09-02, and this is a PIC/HoU page. So a
file the seven-year mover moved out resolves in **no** library the viewer can reach, and
`mergeRecords` reported it as `deleted`.

> ⚠ **THERE IS NO CLIENT-SIDE FIX, and this should not be re-attempted.** The viewer holds nothing on
> `Archive` / `Archive Highly Confidential Document`, so a probe by path, by title, or by
> `GetFileById` is **security-trimmed to 404** — byte-for-byte identical to the file having been
> deleted. Only something that WATCHED the move can record that it happened.

So the mover records it, exactly as `Auto-route` already records a replacement in `ReplacedAt`.

---

## 1. What is already built (code side)

- `ArchivedAt` (DateTime) added to `RECORD_COLUMNS` — **reconciliation creates it** on
  `CRS Submissions` on its next run, alongside every other record column.
- `RecordState` gained `"archived"`. Order: `live` > `cancelled` > `archived` > `deleted`/`unknown`.
- The record read is a **three-rung ladder** (full → without `ArchivedAt` → without the 2026-08-28
  replacement columns either), so a site that has not reconciled degrades one column at a time.
- My Submissions renders a slate **Archived** badge, `in the archive` in the action cell, and a line
  saying it is retained under the seven-year rule and only C-level users can open the archive.

**Until the flow half is built nothing writes the column, so nothing changes on screen.** That is a
silent no-op by design, not a failure — the page behaves exactly as it did before.

---

## 2. The change, per flow

Add **three actions** inside the existing `Apply to each` that moves a file, **after `MoveTo`
succeeds** and after `Find_the_moved_item` (which already runs, for `Stamp_Archived` and
`Reset_inheritance`).

### 2.1 `GetArchivedSubmissionFileId` — read the stamp off the moved item

`Send an HTTP request to SharePoint`, `GET`:

```
_api/web/lists(guid'<ARCHIVE LIBRARY LIST GUID>')/items(@{first(body('Find_the_moved_item')?['value'])?['ID']})?$select=SubmissionFileId
```

Headers: `Accept: application/json;odata=nometadata`.

> ⚠ **`MoveTo` PRESERVES COLUMNS**, which is the whole reason this works — the stamp travelled with
> the file. Read it **after** the move, from the archive library, not before: reading it first and
> holding it across the move adds a failure mode for no benefit.

> ⚠ **USE THE LIST GUID, NOT `getbytitle`.** These libraries have been renamed three times (see
> CLAUDE.md, 2026-08-28) and a title-based reference breaks silently on the next rename. The GUID is
> on the library's Settings page URL (`List=%7B<GUID>%7D`).

### 2.2 `GetArchivedRecordId` — find the record row

`Send an HTTP request to SharePoint`, `GET`:

```
_api/web/lists/getbytitle('CRS Submissions')/items?$select=Id&$filter=SubmissionFileId eq '@{body('GetArchivedSubmissionFileId')?['SubmissionFileId']}'
```

### 2.3 `StampArchivedRecord` — write the date

`Send an HTTP request to SharePoint`, **`POST`** (inside an `Apply to each` over
`body('GetArchivedRecordId')?['value']`):

```
Uri:    _api/web/lists/getbytitle('CRS Submissions')/items(@{items('Apply_to_each_2')?['ID']})
Method: POST
Headers:
  Accept:        application/json;odata=nometadata
  Content-Type:  application/json;odata=nometadata
  X-HTTP-Method: MERGE
  IF-MATCH:      *
Body:   { "ArchivedAt": "@{utcNow()}" }
```

> ⚠ **`POST`, NOT `GET`.** `X-HTTP-Method: MERGE` only takes effect on a real POST — a `GET` with a
> body saves cleanly, runs green, and writes nothing. That exact mistake was made building
> `StampReplacedRecord` on 2026-09-02.

> ⚠ **ISO, via `utcNow()`.** Gotcha #1's `M/D/YYYY` belongs to `validateUpdateListItem`; a plain
> `/items` MERGE goes through the OData layer and answers a locale string with *"Cannot convert a
> primitive value to the expected type 'Edm.DateTime'"*.

---

## 3. Rules that are not optional

- **Every one of the three actions gets `Run after` = *is successful* AND *has failed* on whatever
  follows it.** A failed record stamp must never stop a file being archived, and must never fail the
  run. Worst case the record reads `Deleted` — yesterday's behaviour.
- **A record row that is not found is a normal outcome, not an error.** Anything uploaded before the
  record feature carries no stamp, so `GetArchivedRecordId` legitimately returns `[]` and the
  `Apply to each` runs zero times. Do not add a failure branch for it.
- **Build it as the SERVICE ACCOUNT on SDG's tenant**, before the first action — the connection is
  baked in for life.

---

## 4. ⚠ The HC clone changes exactly ONE value, and the trap is swapping two

Every HC clone in this project has shipped with a library reference nobody swapped. Here the risk is
the opposite one:

1. `GetArchivedSubmissionFileId`'s **list GUID** → the HC archive library. **This is the only swap.**
2. `CRS Submissions` in 2.2 and 2.3 is **SHARED and must NOT be swapped.** There is one submissions
   list for both verticals.

Repointing `CRS Submissions` at some HC-sounding list would find nothing and stamp nothing, silently,
and the symptom would be identical to not having built the flow at all.

---

## 5. Testing

1. Reconcile first, or `ArchivedAt` does not exist and the MERGE 400s.
2. Upload a file as an ordinary uploader, approve it, let it route.
3. Run the mover against a cutoff that includes it.
4. My Submissions, as that uploader: the row reads **Archived**, slate, with *"Moved to the archive on
   YYYY-MM-DD under the seven-year retention rule"* — **not** Deleted.
5. Repeat once through the HC vertical, and confirm the stamp landed in the SHARED `CRS Submissions`
   list rather than nowhere.

⚠ **Files already archived get no stamp.** Nothing backfills, so every document the movers already
moved (57 + 17 on ClarenceDMSTesting, 2026-09-02) will go on reading **Deleted** unless `ArchivedAt`
is set on those rows by hand. Decide with the client whether that matters; on SDG nothing has been
archived yet, so it starts clean.
