# Bulk Upload — Single Selection — Design

**Date:** 2026-08-03
**Status:** Agreed, not yet implemented
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
- **Document naming.** The Form composes
  `[Project] - [Vendor] - [Document Name] - [Date]`. Bulk Upload has no per-file Document
  Name and no Project/Vendor fields in these mockups, so each file **keeps its original
  name**. Renaming 50 files identically would produce 50 collisions.
- **Per-file failure is not fatal.** A file that fails upload or tagging is marked failed
  and the run continues. The existing `Outcome` union (`uploaded` / `skipped` / `failed` /
  `tagFailed`) already covers this and stays.
- **Replace prompts.** The single-file "Replace Existing File" dialog does not scale to 50.
  Phase 1: a same-named file is **skipped** and reported as such in the summary, with no
  prompt. Prompting per file would make a 50-file run unattendable.

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
6. Re-upload one same-named file; confirm it is skipped and reported, with no prompt.
