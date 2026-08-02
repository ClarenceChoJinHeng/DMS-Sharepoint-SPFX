# Bulk Upload — Single Selection — Design

**Date:** 2026-08-03
**Status:** Implemented in 1.0.71.0 (`feat/folder-abbreviations`) — awaiting the tenant
verification in §7
**Supersedes:** `2026-07-24-bulk-upload-two-batch-design.md` (batching deferred to Phase 2)
**Related:** `2026-07-28-bulk-upload-progress-and-scroll-design.md`, and the Form relayout
shipped in 1.0.62.0 – 1.0.68.0

---

## 1. What changes

Bulk Upload currently models **batches**: up to `MAX_BATCHES = 2` groups, each with its own
metadata selection, edited through a side panel. `Batch`, `BatchSelection`, `BatchOutcome`,
`LiveBatch`, the panel and the editing state account for 169 references across a
2,586-line file that is already over the 2,000-line lint limit.

Phase 1 drops batching. One folder destination, one metadata set, up to 50 files.

The screen's purpose is **migrating the client's historical documents** into the Documents
library — not day-to-day filing, which is what the Form is for. That is why one shared
metadata set is sufficient, and why the files arrive already named.

This is a **net deletion**. The mechanism the new screens need already exists —
`postFileWithProgress`, `UploadingBar` and the per-file `FileState` union all stay.

## 2. Layout

Two cards, matching the Form's order and styling.

**DOCUMENTS FOLDER INFORMATION** — the same shape as the Form's folder card: Upload to,
Segment, Department, Unit | Year | Document Type on one row, Remark spanning the row.

**DOCUMENTS DETAILS** — the file area, then Document Date | Confidential Level, with the
Legally Privileged tick beside Confidential Level when it applies.

Actions sit bottom-right: Cancel, Upload.

## 3. The three states of the file area

**Empty.** A drop zone reading *"Choose multiple documents or drop them here"* with
*"Max. 50 files."* beneath it. Same component and styling as the Form's single-file zone;
`multiple` on the input is the only functional difference.

**Selected.** A green summary bar — *"5 document selected"* — with an **Add more** link on
the right. Below it, one row per file: name, size, and a `✕` that removes just that file.

Add more **appends**; it does not replace the selection. The single-file picker replaces,
so this is the one behaviour a user coming from the Form would guess wrong, and it is why
Add more is a distinct control rather than a second use of the drop zone.

**Uploading.** A header — *"3 of 5 documents processed"* — with an overall percentage and
bar. Each row shows `READY` (done) or `LOADING` with its own bar and percentage. Rows keep
their order, and completed rows stay visible rather than disappearing, so the header count
can be checked against the list.

## 4. Rules

- **50 files**, enforced on the combined selection after Add more, not per pick.
- **Extension checks go through one gate**, as in the Form. `accept` filters the file
  dialog only, and a drop bypasses it entirely — validating in the picker alone would let
  a dropped `.exe` through. Rejected files are reported by name and dropped from the
  selection; the rest proceed.
- **One metadata set applies to every file.** That is the point of the screen.
- **Files keep their own names.** This screen exists to bring the client's **historical
  documents** into the Documents library, and in Phase 1 they name those files themselves
  before uploading. So no composition here — the Form's
  `[Project] - [Vendor] - [Document Name] - [Date]` rule does not apply. Applying it would
  also be impossible: there is one metadata set for the whole selection, so all 50 files
  would be handed the same name.
- **Per-file failure is not fatal.** A file that fails upload or tagging is marked failed
  and the run continues. The existing `Outcome` union (`uploaded` / `skipped` / `failed` /
  `tagFailed`) already covers this and stays.
- **Replace prompts, one file at a time.** A same-named file raises the existing "Replace
  Existing File" dialog, and the run waits for the answer before continuing. Yes overwrites
  that file, No skips it and records `skipped`; either way the remaining files carry on.
  Historical uploads are exactly where a genuine re-upload and an accidental duplicate look
  alike from the outside, and only the person doing it can tell them apart — so this is a
  decision the screen must not make on its own, even at the cost of an unattended run.

## 5. Removed

`MAX_BATCHES`, `Batch`, `BatchSelection`, `BatchOutcome`, `LiveBatch`, the `batches` state,
the batch side panel, `panelOpen`, `editingId`, and the per-batch summary. Anything
reachable only from those goes with them.

`MAX_FILES` stays at 50 and stops meaning "per batch".

## 6. Out of scope — Phase 2

Several batches with differing metadata in one run. The superseded spec still describes
that model; if it returns, it should return as "add another destination" layered over this
single-selection screen, not as the panel that exists today.

## 7. Verification

Not automatable — needs the tenant.

1. Pick 3 files, confirm the count, remove one, confirm it reads 2.
2. Add more; confirm it appends rather than replaces.
3. Try to exceed 50; confirm the message names the limit and nothing is silently dropped.
4. Drop a disallowed type; confirm it is rejected by name and the others survive.
5. Upload; confirm the per-file bars, the header count, and that every file lands in
   `…/Unit/Year/DocumentType` carrying the shared metadata.
6. Re-upload a selection containing two same-named files; confirm the dialog appears for
   each in turn, that Yes overwrites and No records `skipped`, and that files after the
   prompt still upload.
