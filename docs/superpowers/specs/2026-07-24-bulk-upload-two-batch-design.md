# Bulk Upload — Two-Batch Design

**Date:** 2026-07-24
**Status:** Approved for implementation
**Web part:** `src/webparts/bulkUpload/components/BulkUpload.tsx`
**Supersedes UX of:** `2026-07-23-bulk-upload-direct-to-documents-design.md` (same upload semantics, new multi-batch UI)

## Problem

Historical-document migration needs more than 50 files per sitting, and the 50-file
single-submission cap makes the current bulk-upload tool slow for large backfills. The
client wants to queue **two** independent batches — each with its own destination folder
and metadata — and upload them in one pass, unattended.

## Goals

- Up to **2 batches** per session, **50 files each** (100 total).
- Each batch has an **independent destination** (Segment → level cascade → Year → Document
  Type) and metadata. Two batches may target the same folder or different folders.
- After configuring a batch it **collapses into a folder tile** (📁 + destination path +
  `Year · DocType · N files`) instead of listing 50 files.
- **Rename functionality removed** — files upload under their original names.
- One **Upload all** button processes batch 1 then batch 2 sequentially.

## Non-Goals

- More than 2 batches (hard cap at 2 for now).
- Per-file metadata within a batch (metadata still applies to every file in a batch).
- Any change to upload semantics: still direct-to-Documents, skipping Staging/approval.

## UX

### Layout
```
Batches
  📁 Batch 1 — GHO › Group Legal › Group Compliance · 2026 · Contracts · 42 files  [Edit] [Remove]
  📁 Batch 2 — Minamas HO › … · 2026 · Invoices · 50 files                         [Edit] [Remove]
  [ + Add batch ]            (hidden or disabled once 2 batches exist)

  ── inline config panel (shown only while adding/editing a batch) ──
  Files picker (up to 50; file list shown only while configuring)
  Upload into: (Business Segment | Project) · Segment select · level cascade
  Year / Period · Document Type · Document Date · Confidentiality · Vendor
  [ Save batch ]  [ Cancel ]

[ Clear ]                                         [ Upload all (N files) ]
```

### Flow
1. Click **+ Add batch** → the inline config panel expands with an empty draft (privileged
   users) or the restricted auto-locked cascade (non-privileged users, unchanged).
2. Pick files (≤ 50, disallowed extensions rejected with the existing toast) and set all
   required metadata.
3. **Save batch** → validates the draft (see Validation); on success, snapshots it into a
   folder tile and **resets the draft panel to empty**, then collapses it.
4. Repeat for batch 2. **+ Add batch** is disabled/hidden once 2 batches exist.
5. **Upload all** processes each batch sequentially and shows per-batch results.

### Edit / Remove
- **Edit** loads the batch back into the draft panel and removes its tile until re-saved
  (Save re-adds it). Only one batch can be edited/added at a time — while the draft panel is
  open, **+ Add batch** and **Upload all** are disabled.
- **Remove** deletes the batch (and its queued files) with no confirmation.

### Reset behaviour
On **Save batch** the draft resets: file picker cleared, every dropdown back to `--`, date
cleared. For **restricted** users the level cascade resets to their allowed default
(auto-selected when only one authorized path exists) — identical to today's single-submission
form. **Clear** empties all batches and the draft.

## Data Model

A saved batch snapshots everything needed to upload independently — no reliance on live
cascade state after Save:

```ts
type BatchSelection = {
  column: string;
  label: string;
  id: string;
  labelCol?: string;
  tidCol?: string;
};

type Batch = {
  id: string;                    // stable key, derived from a monotonic counter (no Date.now in render)
  files: File[];                 // ≤ 50, no per-file rename
  modeKey: string;
  modeLabel: string;             // segment label, for the tile + BusinessSegment column
  termSetGuid: string;           // for the BusinessSegment selection value
  levelSelections: BatchSelection[];   // resolved level terms (label + id)
  docTypeId: string;   docTypeLabel: string;
  yearId: string;      yearLabel: string;
  confId: string;      confLabel: string;
  vendorId: string;    vendorLabel: string;   // optional; empty when unset
  documentDate: string;          // ISO yyyy-mm-dd as entered; converted at upload
  destinationLabel: string;      // "GHO › Group Legal › Group Compliance" for the tile
};
```

Storing resolved labels+ids (not the `levelChoices` matrix) keeps each batch self-contained
for both tile rendering and `formValues` construction.

## Upload Core (refactor)

Extract the existing per-submission logic in `handleUpload` into a pure-ish helper:

```ts
uploadBatch(batch: Batch): Promise<FileResult[]>
```

`uploadBatch` performs, unchanged from today:
1. `lookupFolderMapping` on the leaf term's id → `folderUniqueId`.
2. `resolveFolderServerUrl` (rename-proof Staging path) → `toDocumentsPath` library swap.
3. `resolveFolderByPath` for the Documents unit folder (must already exist — never
   auto-created, to preserve ACLs).
4. `ensureFolder` Year then Document Type subfolders (safe — inherit unit ACL).
5. Per-file: existence probe (`overwrite=false` → skip), upload raw `File` body, fetch list
   item, `validateUpdateListItem` with `formValues`, `HasException` check.

`formValues` is built from the batch's stored selections via `buildLevelFormValues` +
`LEVEL_COLUMNS`, exactly as now, with the `BusinessSegment` row unshifted from
`modeLabel`/`termSetGuid`.

`handleUpload` becomes: for each batch in order, `setStatus("Batch i of n …")`, call
`uploadBatch`, and collect results into a per-batch group. Batch failures never stop the
next batch.

### Preserved gotchas
US date format via `toSpDate`; `HasException` on each `validateUpdateListItem` result; raw
`File`/`Blob` body (never FormData); `GetFolderByServerRelativeUrl` alias form; the
folder-map "not mapped yet" and "folder missing in Documents" error toasts.

## Validation (at Save batch)

A batch cannot be queued incomplete. On **Save batch**, require: ≥ 1 file, all level values,
Year, Document Type, Document Date, Confidentiality. Duplicate filenames **within the
batch** are rejected (files upload under original names, so the check is on `file.name`,
case-insensitive). Non-privileged users' chosen path must be authorized (existing
`resolveValidPaths` logic already constrains the cascade, so this holds by construction).

Cross-batch duplicate filenames are **allowed** — the two batches may legitimately target
different folders.

## Results

Results are grouped per batch:

```
Batch 1 — 48 uploaded, 2 skipped
  <per-file outcome rows, same as today>
Batch 2 — 50 uploaded
  <per-file outcome rows>
```

Per-file outcomes (`uploaded` / `skipped` / `failed` / `tagFailed`) and their remediation
copy are unchanged. Fully-succeeded batches collapse to just their summary line; failed/
skipped rows stay expanded.

## Restricted vs Privileged

Unchanged. The draft panel runs the existing restricted auto-lock (`applyRestrictedMode`) or
privileged manual cascade (`initCascade`). Each batch is validated against the user's
authorized paths at Save time.

## Testing

- Manual, in the workbench, both privileged and restricted accounts:
  - 1 batch, 2 batches, same folder twice, two different folders.
  - Duplicate filename within a batch (rejected at Save); same filename across two batches
    to different folders (allowed, both upload).
  - A batch whose Documents unit folder is missing (error surfaced, other batch still runs).
  - Partial per-file failure (one bad file) — batch continues, results correct.
- Existing shared-module unit tests (`formModel`, `dmsFolderMap`) remain green; no signature
  changes to those modules.

## Out of Scope / Future

- 3+ batches (raise `MAX_BATCHES` if the client asks).
- Per-file metadata; drag-and-drop; progress bar per file.
