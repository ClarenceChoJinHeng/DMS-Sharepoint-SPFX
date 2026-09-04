# My Submissions groups by SUBMISSION, then batch, then file

**Date:** 2026-08-22 · **Status:** designed, implementing
**Client:** *"they want My Submission to detect or to record the submission of the uploader themselves… 
if they click on the batch it will show the details of the batch in upload form with the details but 
view only."* Clarified to three levels: **submission → batches → each file's upload-form detail.**

## 1. The blocker: a batch leaves no trace

`Batch.id` in `shared/uploadBatches.ts` is documented as *"stable within a session"*. Batches exist in
the browser while the form is open and are **never written to SharePoint**. So today My Submissions has
nothing to group by, which is exactly why it lists loose files.

Grouping therefore requires stamping identity onto every uploaded file at upload time.

## 2. Two ids, because the client asked for two levels

| Column | Written | Meaning |
|---|---|---|
| `SubmissionId` | once per **press of Upload** | everything sent together |
| `BatchId` | once per **batch** | one destination folder within that submission |

Both **plain text**, both in **all four libraries** (`allLibraryTitles()`).

⚠ **THE COLUMNS MUST EXIST IN `Documents` AND `HC Documents` TOO, NOT ONLY THE APPROVAL SIDE.**
SharePoint's copy carries over only columns that EXIST at the destination — so a file routed by
Auto-route would arrive with its reference silently stripped, and the submission would lose its
approved files one by one as they were approved. Same trap as `Remark`/`LegallyPrivileged` in 2026-08-10.

**Reference format:** `SUB-YYYYMMDD-XXXX` / `BAT-YYYYMMDD-XXXX`, uppercase base32 suffix. Human-readable
because it appears on screen as a reference the way the client's own claim system shows one; a bare GUID
is unreadable and unspeakable over the phone. Uniqueness comes from the random suffix, not the date.

## 3. ⚠ THE WRITE MUST BE CONDITIONAL, OR IT DESTROYS ALL METADATA

`validateUpdateListItem` fails the **WHOLE call** on one unknown field name (gotcha #4) — every column
lost, not just the missing one. So adding `SubmissionId` to the payload on a site whose libraries lack
the column would mean **every upload silently loses Remark, Confidentiality, the tier fields and
everything else**, reported only as the generic "Uploaded, but tagging metadata failed".

Therefore:
- `Form.tsx` includes the two fields **only when it has confirmed they exist** on the target library.
- The check is a `/fields` read, cached per library for the page's life — the same shape as
  `pickFullNameField`, which exists for this class of problem.
- Absent columns mean the upload proceeds **exactly as it does today**, ungrouped. Degrading to the
  current behaviour is always correct; losing a document's metadata never is.

## 4. Reconciliation asserts the columns

Added to the four-library pass that already asserts inheritance. `ensureColumn` is idempotent, so this
costs one `/fields` read per library per run and creates nothing on a provisioned site.

**Assert, never rely on a one-off provisioning step** — the fifth instance of that rule. A library
added later, or a site provisioned from an older runbook, otherwise silently stops grouping.

## 5. Grouping rules — `shared/submissionGroups.ts`, pure

- Rows sharing a `SubmissionId` form one submission; within it, rows sharing a `BatchId` form one batch.
- **A row with no `SubmissionId` is its own submission of one file, with no reference.** Every file
  uploaded before today is in this state (client's choice). Nothing is inferred from timestamps and
  folder: two unrelated files a minute apart would be presented as a submission that never happened.
- A submission's **date** is its earliest `created`; its **status summary** counts its files by status.
  Never a single status — a submission can be part approved and part pending, and flattening that to one
  badge is how an uploader concludes a rejected file was fine.
- Sort newest first, reusing the existing comparator rather than a second definition.

## 6. Levels 2 and 3 are VIEWS, not new reads

The rows are already loaded. Drilling in filters what is on screen; it issues no request except the
per-item `FieldValuesAsText` the existing detail view already does for one file at a time.

Level 3 lays the fields out in the upload form's own order — folder information, then per-file document
details — so the uploader recognises what they filled in. **Read-only, with no controls**: this page has
never written a document and must not start.

## 7. Out of scope

- **`BulkUpload.tsx` is not stamped.** It is admin-only, writes straight to the approved side, and is not
  "the submission of the uploader themselves". Its files fall into the no-reference case, correctly.
- Retro-fitting references to historical files (see §5).
- Any change to what My Submissions can DO — the request buttons and their rules are untouched.

---

## Revision — level 3 is the upload form, read only (2026-08-22, 1.0.222.0)

The first build made level 3 a file LIST, opened one file at a time through the existing detail
panel. The client's correction: *"When I say upload Form I literally meant a copy of the upload form
with the details in it but view based only."*

That is a different shape, and the form itself defines it. An uploader picks a destination **once**
and fills in each file's details beneath it, so the read-only view renders one **Document folder
information** card followed by one expanded card per file — nothing to click through.

### What changed

- `documentDetails.ts` splits the fixed fields into `BATCH_FIXED_FIELDS` (Document Type, Year — the
  built-in below-Unit folder tiers, chosen once per destination) and `FILE_FIXED_FIELDS` (everything
  the uploader types per file). `FIXED_FIELDS` is now their concatenation, **derived**, so the
  single-document panel is unchanged and cannot drift from the batch view. `buildBatchRows` and
  `buildFileRows` render the halves. A test pins the partition as exact.
- Level 1 gains a **Batches** column, per the client's sketch.
- The tab count is computed from `groupSubmissions` rather than the status-tab `counts` map, which
  had it reading `Submissions (0)` beside a table of fourteen.

### The cost, and why it is acceptable

Level 3 issues **one `FieldValuesAsText` request per file**. That endpoint is per-ITEM — it is why
the flat list shows no metadata at all, where hundreds of rows would be hundreds of requests. Here
the set is a batch (a handful of files) and the view is opened deliberately. `Promise.allSettled` is
unavailable on this tsconfig (gotcha #3), so each read carries its own catch inside `Promise.all`:
one unreadable file must never blank the batch.

The destination is read from whichever file in the batch **could** be read, never simply the first —
a batch is one folder so any of them answers, but the first may be the one that failed, and falling
back to it would show an empty destination for a batch that has one. A file whose read returned `{}`
says so, rather than rendering as a document with no metadata: empty and unknown are
indistinguishable from that response, so the panel states both.

## Revision 2 — old files are grouped, and rows are labelled by destination (1.0.222.0)

The client, on seeing revision 1: *"Can we not show per file instead… client doesn't want to go
through each file one by one figuring out and noting down which one."*

**The refusal to infer is overruled.** One row per referenceless file is what the data honestly
supports, and it is also exactly the flat list this feature exists to replace — on a site with a year
of history, every row. Files with no `SubmissionId` are now grouped by **folder and day**.

Folder *and* day, never folder alone: a unit files into the same destination every month, so the
folder by itself would present a year of unrelated uploads as one submission. Same destination on the
same day is the strongest evidence available and is the shape a real batch has. A row with no
timestamp groups with nothing — unknown must never widen a group.

The cost is that two people filing into one folder on one day appear as one group. `inferred` carries
that to the screen, which labels the row *"Grouped by folder and date — uploaded before submissions
were recorded"*. Nothing is hidden and no file moves; only how the rows are stacked. A guess presented
as a record is the failure mode the original rule guarded against — a guess **labelled** is not.

Rows are also labelled by their **destination path** rather than a reference or a file name, per the
client's sketch. A reference is what you quote to an approver; a path is what you recognise, and
recognising the one you want is the whole job of the list. The reference stays beneath it, and a
submission spanning several destinations says `+ N more destinations`.
