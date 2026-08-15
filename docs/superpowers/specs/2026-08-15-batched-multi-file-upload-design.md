# Batched multi-file upload (the Form) — design

**Status:** DESIGN — not yet built, awaiting review.
**Date:** 2026-08-15
**Web part:** `src/webparts/form/components/Form.tsx` (the approval-library upload form).

**Client ask:** *"the current upload form, its only allowing one file per upload, client wants each
file to go to different places, so multiple upload."* Refined the same day: *"they want batches, so
this first few one can keep adding and renaming and more, then after that save batches and add more."*

**Related, and load-bearing:** `2026-07-24-bulk-upload-two-batch-design.md` (built) and
`2026-08-03-bulk-upload-single-selection-design.md` (**removed it again**). See §8 — this project has
built batching once and torn it out, and the reason it was torn out applies to `Form.tsx` today.

---

## 1. What a batch is

**A batch is one destination folder plus the files that belong in it.**

That is what reconciles the client's two sentences. "Each file goes to different places" cannot mean
per-file tier pickers — the designer's mockup keeps Segment/Department/Unit at the top — and it cannot
mean one shared folder either. The destination is chosen **per batch**, and an upload carries several.

```
BATCH 1  ->  GHO / Group Finance / Corporate / 2026 / Tax Return
   Rewards B            Project-A_Vendor-B_Rewards.pdf   2MB
   HRIS Report          Project-A_Vendor-B_HRIS.pdf      2MB

BATCH 2  ->  GHO / Group Finance / Corporate / 2025 / Invoice
   Vendor A Invoice     Project-A_Vendor-B_TS.pdf        2MB

                                             [ + New batch ]   [ Upload all 3 ]
```

**Save batch stages; nothing is sent until Upload.** The client chose this over committing each batch
as it is saved, for the reason that makes it worth its cost: they see the whole plan, and can still
edit an earlier batch, before anything exists in SharePoint.

### 1.1 The cost of staging, stated plainly

**Staged batches live only in browser memory and cannot be persisted.** A `File` is not serialisable —
there is no localStorage draft and no recovery. A closed tab, a crash or a session timeout loses every
staged batch.

Mitigations, and their limits:

- A `beforeunload` guard whenever anything is staged — the pattern `SegmentCreator` already uses. It
  cannot survive a crash, and the browser owns the wording.
- The staging area is headed **"Nothing has been uploaded yet"**, and the button carries a live count.
  The one belief a user must never form is that saving a batch sent it.
- **Cancel confirms, and names how many files it discards.**

This is a genuine regression against uploading one file at a time, and it is the price of the model the
client chose. Written down so a later lost-work report is recognised as the known cost, not a bug.

---

## 2. What varies per file, and what does not

| Field | Scope | Why |
|---|---|---|
| Segment / Department / Unit (permissioned tiers) | **Batch** | They decide who can see the file. |
| Below-Unit tiers (SubUnit, Year, Document Type, any configured level) | **Batch** | They are folder levels, and the batch *is* the folder. |
| Document Name (rename) | **File** | Per the mockup. Blank keeps the original filename. |
| Project Name | **File** | |
| Vendor / Customer Name | **File** | |
| Document Date | **File** | |
| Confidentiality Level | **File** | Metadata, never a permission — any PIC may upload at any level. |
| Legally Privileged | **File** | Offered only where *that file's* confidentiality equals the `legallyPrivilegedFor` config row, so the tick appears and disappears per row. |
| Details / Remark | **File** | Existing fields, same treatment. |

**Year and Document Type are path levels, so they belong to the batch.** They are still written as
metadata, derived from the batch — unchanged from today.

### 2.1 A new file inherits the last values typed in its own batch

Adding a fifth file pre-fills Project, Vendor, Date and Confidentiality from the file most recently
edited **in that same batch**. This is not invented data: it is copied from a sibling the user typed, it
is visible, and it is editable. Without it a ten-file batch means typing one vendor ten times, and
effort is the entire complaint.

**Document Name is never inherited** — two files sharing a name is the collision in §5.

---

## 3. The stale-chain guard is the dangerous part

`Form.tsx` already re-reads the mode row's `Levels` immediately before writing and refuses with *"reload
the page"* if the chain moved (gotcha 10b, `chainSignature` / `freshChainFor` at :737 and :758).

Batching makes that guard **more likely to fire and far more expensive**: a staged session lasts minutes
rather than seconds, and *"reload the page"* would destroy every staged batch.

1. `chainSignature` is captured **per batch, at save time**, beside the chosen path.
2. At Upload the chain is re-read **once per distinct segment**, not once per batch.
3. If a segment's chain moved, only the batches on that segment are affected. They **stay staged** and
   are marked *"the folder structure for this segment changed — choose this batch's destination again"*,
   with the pickers refreshed **in place**. Never a page reload; never a discard.
4. Batches on other segments upload normally.
5. **A failed read never blocks.** It proves nothing, and taking the form down over a transient error is
   worse than the risk it guards against. Unchanged from today.

Point 3 is not a nicety. Today the refusal costs one file's worth of typing; here it would cost
everything, so the guard has to become gentler as it becomes more likely.

---

## 4. Upload

Sequential, in order: batch 1's files, then batch 2's. SharePoint has no batch transaction, so this is
a loop with a progress line — **"Batch 2 of 3 · file 1 of 4 · Vendor A Invoice"**.

- **Folders are ensure-created once per batch**, before its first file — not per file.
- Each file is uploaded as a raw `File`/`Blob` (gotcha #6), then tagged via `validateUpdateListItem`
  with `HasException` checked per result (gotcha #4). `DocumentDate` stays `M/D/YYYY` (gotcha #1).
- Every file arrives Pending. Approval, Draft Item Security and Auto-route are unchanged.

### 4.1 Partial failure: keep the successes

The client's choice, and the only defensible one — rollback would mean deleting files that already
uploaded, which can itself fail, and could delete a document an approver has already opened.

- A file that uploads is **removed from its batch**; an emptied batch disappears.
- A file that fails **stays staged with its reason on the row**, so what is left on screen is exactly
  what still needs doing. **Retry** re-runs only those.
- The result line is counts, never a verdict: *"Uploaded 4 of 6. 2 could not be uploaded."*
- **A retry can never upload a file twice**, because success is what removes it from the list.

---

## 5. Collisions

Two failure modes, caught in different places, because only one is knowable offline.

**Within a batch — blocked at save time.** Two files whose resolved upload name is identical would land
one on top of the other. `buildUploadName` is pure, so this is checked as the client types: both rows
flag, each naming the other, and **Save batch is refused**. Compared case-insensitively, because
SharePoint file names are.

**Against the destination folder — a per-file failure at upload time.** Whether a name already exists
there cannot be known without a read, and a read at save time would be stale by Upload anyway.

**Never overwrite.** A conflict fails that one file — *"a document called X is already in this folder"* —
and the rest of the batch continues. This follows the subtree migrator's precedent: a collision gets a
rename, never a silent overwrite. It matters more here, because the existing file may be Approved and
already routed to `Documents`.

---

## 6. Limits

- **File type** — the existing `allowedFileTypes` check runs **as each file is added**, so a rejected
  file never enters a batch. A refusal writes `UploadRefused` to the audit log, as today.
- **File size** — unchanged.
- **Staged total** — a **warning, not a block**, past 25 files or 200 MB across all batches: the browser
  holds every byte, and a client queueing a hundred scans should hear about it before the tab dies.
- **The same file in two batches** — warned, not blocked. Filing one document in two places is unusual
  but legitimate, and blocking it invents a rule nobody asked for.

---

## 7. Where the code goes

**`src/shared/uploadBatches.ts` — pure, Jest-tested.** Everything decidable without the network:

- the batch/file model and its ids
- `collisionsWithin(batch)` — §5
- `inheritDefaults(batch)` — §2.1
- `applyUploadResults(batches, results)` — §4.1, the part that must never lose or duplicate a file
- `stagedTotals(batches)` — §6
- `batchesNeedingRepick(batches, movedSegments)` — §3

`Form.tsx` keeps the UI, the requests and the existing per-file upload call — the code that is
site-verified and must not be rewritten for what is a UI change.

### 7.1 Two facts found while reading `Form.tsx`, which shape the integration

**The uploaded name is NOT the typed name.** `handleUpload` builds it as
`buildUploadName(file.name, composeUploadBase(projectName, vendor, docName, documentDate))` (:1461) —
the `[Project] - [Vendor] - [Document Name] - [Date]` convention. So "Document Name" is one *segment*
of the result, and **two files with different typed names still collide** when project, vendor and date
match. Comparing typed names would miss precisely that case.

`StagedFile.finalName` therefore holds the composed name, computed by the form and stored, and
`collisionsWithin` prefers it. Not re-derived in `uploadBatches.ts`: a second copy of the naming
convention would be free to drift from the one that actually performs the upload.

**A batch must snapshot its destination, not reference the pickers.** Today the handler reads
`levelValues`, `levelChoices`, `tierValues` and `tierPlan()` live, because there is one destination and
it is the one on screen. With several batches, the pickers describe whichever batch is being edited
*now* — so uploading batch 1 would use batch 3's department, silently, into a real folder that looks
correct.

`Batch.destination` is therefore captured at **Save batch**, when the pickers are current and validated,
and holds:

| Key | Purpose |
|---|---|
| `unitSru` | the permissioned leaf folder — where ensure-creation starts |
| `segments` | below-Unit folder names in chain order |
| `levelSelections` | the label/GUID pairs `buildLevelFormValues` needs |
| `tierFormValues` | the already-built `FieldName`/`FieldValue` pairs for below-Unit tiers |

Upload then needs **no picker state at all**: ensure `segments` under `unitSru`, then per file write
`tierFormValues` + `levelSelections` + that file's own metadata.

### 7.2 The remaining integration, in order

1. `batches: Batch[]` (saved) and `draftFiles: StagedFile[]` (the batch being built); `activeFileId`
   selects which per-file panel is open.
2. The existing field block becomes the editor for the active file. Switching panels captures the
   editor into that file and loads the next — the existing `onChange` handlers are left alone.
3. **Save batch** captures the editor, runs `canSaveBatch`, snapshots the destination per §7.1, and
   pushes a `Batch`. **New batch** clears the pickers via the existing `resetForm`.
4. `handleUpload` becomes: re-read the chain **once per distinct segment** → `batchesNeedingRepick` →
   for each uploadable batch, ensure folders once → per file, the existing upload + tag code with its
   values taken from `sf.meta` instead of component state → collect `UploadResult[]` →
   `applyUploadResults`.
5. The per-file tail of today's handler (:1636–:1846) is extracted as
   `uploadStagedFile(dest, sf): Promise<UploadResult>` — it returns a result instead of calling
   `showToast` and `return`, which is the only behavioural change to that code.

---

## 8. Why this is factored this way: batching has been removed from this project once

`2026-07-24` built batching in **Bulk Upload** — `MAX_BATCHES = 2`, independent destination per batch,
sequential "Upload all". `2026-08-03` **removed it**, and the stated reason is the whole argument for §7:

> `Batch`, `BatchSelection`, `BatchOutcome`, `LiveBatch`, the panel and the editing state account for
> **169 references across a 2,586-line file** that is already over the 2,000-line lint limit.

`Form.tsx` is **2,653 lines** and over the same limit. Repeating the shape that was torn out, in a file
of the same size, would produce the same outcome.

Two things follow:

- **The rules leave the component.** A pure module with tests is what makes the model survivable, and it
  is the difference between this attempt and the one that was reverted.
- **Bulk Upload's Phase 2 is the same model.** That removal was *"batching deferred to Phase 2"*, not
  *"batching was wrong"*. `uploadBatches.ts` should be written so Bulk Upload can adopt it — batch-level
  metadata there, per-file here, which is a difference in what a batch *contains*, not in what a batch
  *is*.

One inherited decision is **not** copied: Bulk Upload dropped rename because one shared metadata set
would give all 50 files the same name. Per-file metadata removes that objection, and rename is in the
client's mockup, so the Form keeps it.

---

## 9. Out of scope, deliberately

- **Drag and drop.** Neither upload web part has it; adding it here would mean two new ways to get a
  file wrong in one release.
- **Persisting staged batches.** Not possible for `File` objects — §1.1.
- **Per-file destinations.** Not what the client asked for once batches were explained, and it would let
  one screen mix units with different audiences.
- **Changing `BulkUpload.tsx`.** Different library, different rules (`UPL` is Read-only in `Documents`),
  different gating. Adopting `uploadBatches.ts` there is a later, separate change.
- **A batch cap.** Bulk Upload capped at 2; the client said "save batches and add more", so the Form is
  uncapped and governed by the §6 warning instead.

---

## 10. Test plan

1. One batch, one file — today's behaviour, unchanged.
2. Two batches, different `Year`, same unit → two folders, both files tagged.
3. Two batches on different segments → both upload; the chain is re-read **twice, not four times**.
4. Structure change mid-session on segment A → the batch on A asks for a re-pick, the batch on B uploads.
5. Chain read fails → everything uploads (fail-open).
6. Duplicate resolved name inside one batch → Save refused, both rows flagged.
7. Name already in the destination → that file fails, its siblings succeed, nothing is overwritten.
8. Force a 403 on one file → successes leave the list, the failure stays with its reason, Retry sends
   only it, and nothing uploads twice.
9. Disallowed extension → never enters the batch; `UploadRefused` written.
10. Close the tab with staged batches → the browser warns.
11. Cancel with staged batches → the confirmation names the file count.
12. `Estate/Mill` segment (Upstream Ops) → the batch's tier labels read `Region` / `Estate/Mill`, not
    `Department` / `Unit`.
