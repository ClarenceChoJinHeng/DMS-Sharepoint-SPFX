# Project Name + Vendor metadata, and single-file upload form redesign

**Date:** 2026-07-28
**Status:** Approved, pending implementation
**Affects:** `src/webparts/form/components/Form.tsx`,
`src/webparts/approvalDocument/components/ApprovalDocument.tsx`,
Staging library schema, `DMS Config` list

---

## Problem

Two gaps, one visible and one silent.

The **visible** gap: the single-file upload form's layout has drifted from the approved
design. Folder-selection fields sit above document metadata, labels differ from the
mockup, and there is no field for a project name on an ordinary Business Segment upload.

The **silent** gap: the Approval Document page already renders a `Vendor` row, but it
always shows "—". The form captures Vendor as free text and then deliberately discards
it — `Form.tsx:1054-1056` skips the `Vendor` `FieldValue` on upload. Nothing writes the
column, so there is nothing to read back. Project Name has the same problem in advance:
there is no column at all.

The round-trip only works when all three links hold: **form writes column → SharePoint
stores it on the item → Approval page reads it from `FieldValuesAsText`**. The typed text
lives nowhere else — not in the filename, not in the folder path, and the form's React
state is discarded the moment the upload completes.

## Goals

1. Free-text **Project Name** and **Vendor/Customer Name** are stored on the uploaded
   file and displayed on the Approval Document page.
2. The single-file upload form matches the approved mockup, excluding the action buttons.
3. No regression to the folder-routing cascade, term-store lookups, or the auto-composed
   document name.

## Non-goals

- Backfilling Project Name / Vendor onto documents already in Staging. Existing items
  will show "—". A backfill, if wanted, is separate work.
- Restyling the Cancel / Upload buttons. Explicitly out of scope.
- Creating the `GroupProjectName` columns. Wired in code, created later.
- Any change to `BulkUpload.tsx`. This spec covers single-file upload only.

---

## Decisions

| Question | Decision |
|---|---|
| Mockup fidelity | Full pixel match, buttons excluded |
| Project Name type | Free text (single line of text), **not** managed metadata |
| Project Name column | New `ProjectName` column in Staging |
| Group-led Projects level | Renamed "Project Name" → **"Group Project Name"** |
| Both project fields on one screen | Show both, always |
| Document Date | **Required.** Briefly made optional to match the mockup, then reverted 2026-07-28 on an updated client requirement |
| Date input format | Keep native `<input type="date">`; format follows browser locale |

---

## 1. Data model

### Staging library columns

Both columns exist. Verified against the live `/fields` API on 2026-07-28 — these are
facts, not inferences:

| Display name | Internal name | Type |
|---|---|---|
| `ProjectName` | `ProjectName` | Text |
| `Vendor/CustomerName` | `Vendor_x002f_CustomerName` | Text |

**The old `Vendor` column was deleted** on 2026-07-28 and replaced by
`Vendor/CustomerName`. The slash in the display name encodes to `_x002f_` in the internal
name — the same pattern as the existing `Year_x002f_Period` elsewhere in the tenant. It
held no data, since nothing ever wrote to it.

### Read keys on the Approval page

`FieldValuesAsText` double-encodes the underscores in its response keys (see the
`sp-fieldvaluesastext-underscore-encoding` memory), so the two columns read differently:

| Column | Read key |
|---|---|
| `ProjectName` | `ProjectName` — no encoded characters, so no double-encoding |
| `Vendor_x002f_CustomerName` | `Vendor_x005f_x002f_x005f_CustomerName`, falling back to the plain name |

The plain-name fallback follows the existing `pick(...)` convention already used for
Document Type and Confidentiality Level.

### Stale references to the deleted `Vendor`

Four places still name the deleted column and must be updated together. None is failing
today only because both write paths are currently disabled — the Form skips Vendor
deliberately, and Bulk Upload's field is hidden so its guard never fires. **This work
re-enables the Form write path, which is exactly when a stale name starts throwing
`HasException`.**

| Location | Current | Change to |
|---|---|---|
| `Form.tsx` `FIELDS.vendor` | `Vendor` | `Vendor_x002f_CustomerName` |
| `BulkUpload.tsx` `FIELDS.vendor` | `Vendor` | `Vendor_x002f_CustomerName` |
| `ApprovalDocument.tsx` `pick("Vendor")` | `Vendor` | double-encoded key + fallback |
| `DMS Config` `col_vendor` row | `Vendor` | `Vendor_x002f_CustomerName` |

### Column visibility

Both columns are **hidden from the Staging library views** and surface only on the
Approval Document page. Hide them by removing them from the view (Library settings → the
view → untick the column). The column still exists, is still written, and is still
returned by the API — this is purely presentational and carries no risk.

`ApprovalDocument.tsx:136` calls `FieldValuesAsText` with no `$select`, so it retrieves
the entire field bag regardless of which columns any view displays. View membership and
the Approval page are completely decoupled.

Two variants to avoid:

- **Column-level `Hidden=true`** ("Hidden (Will not appear in forms)"). Reads and writes
  still work, but the column disappears from the forms UI too, which makes a missing
  value much harder to diagnose. Prefer view-level hiding.
- **Read-only.** `validateUpdateListItem` refuses read-only fields, so every upload would
  fail with `HasException` on that field. Never mark these read-only.

Because the columns are invisible in Staging, the Approval Document page is the **only**
place a wrong value is observable. The UAT steps below are the sole verification path —
do not skip them.

### Two distinct "project name" concepts

These must not be conflated. They are different fields with different sources, storage,
and effects:

| | Free-text `ProjectName` | Level `GroupProjectName` |
|---|---|---|
| Source | User types anything | Picked from the Group-led Projects term set |
| Storage | One text column | Label column + Tid column |
| Applies to | Every upload, any segment | Only the Group-led Projects mode |
| Affects folder path | No | **Yes** — it is a folder level |
| Column exists? | Created as part of this work | Not yet; deferred |

Levels write a label + term-GUID pair via `buildLevelFormValues` in
`src/shared/formModel.ts`, so `GroupProjectName` resolves to two columns:
`GroupProjectName` (label) and `GroupProjectNameTid` (term GUID). This is the same
pattern already used by Department and Unit.

Both fields are visible simultaneously on a Group-led Projects upload. This is the
accepted decision. The distinct labels — "Project Name" versus "Group Project Name",
the latter rendered as a dropdown — are what disambiguate them for the user.

> **Supersedes earlier specs.** `2026-07-21-segment-onboarding-plan.md` (line 105) and
> `2026-07-16-tenant-seed-data-runbook.md` (line 132) both reserve
> `ProjectName`/`ProjectNameTid` for the Group-led Projects *level*. This spec reassigns
> the bare `ProjectName` name to the new free-text column and moves the level to
> `GroupProjectName`/`GroupProjectNameTid`. Those two documents must be amended so a
> future onboarding pass does not recreate the collision. Neither column exists yet, so
> nothing is migrated — this is a naming change on paper only.

### DMS Config

| Setting row | Value | Status |
|---|---|---|
| `col_projectName` | `ProjectName` | **New row** |
| `col_vendor` | `Vendor_x002f_CustomerName` | Row exists; **value must be corrected** |

Add `projectName` to the `columns` block of the `DmsSettings` type and to
`DEFAULT_SETTINGS.columns` in `Form.tsx`, following the existing pattern, plus a
`get("col_projectName")` read alongside the other column settings.

The Group-led Projects mode row's `Levels` JSON changes its entry from
`{"label":"Project Name","column":"ProjectName"}` to
`{"label":"Group Project Name","column":"GroupProjectName"}`, with matching
`labelCol`/`tidCol` once those columns exist. The same rename applies to `DEFAULT_MODES`
in `Form.tsx` so the offline fallback stays consistent.

> The `column` key rename is required, not cosmetic: leaving it as `ProjectName` would
> point the level at the new free-text column, and a Group-led Projects upload would
> overwrite the user's typed value with the term label.

---

## 2. Form layout

### Current structure

Two cards — **File** (file picker, document name) and **Document Information**, the
latter containing the folder cascade *above* the document metadata.

### Target structure

The mockup inverts this: document metadata first, folder selection second.

**Card 1 — "Upload a Document"**

```
Document Name                                        [full width]
  ↳ Max. 30 character
Upload Document                                      [full width]
Project Name        | Vendor/Customer Name | Document Date
  ↳ Max. 30 char      ↳ Max. 30 char
Year*               | Document Type*       | Confidential Level*
```

**Card 2 — "DOCUMENT FOLDER INFORMATION"**

```
Upload to  ◉ Business Segment
Segment*
Department*
Unit*
```

The existing `.dms-grid` is already a three-column grid, and `.dms-field small` already
styles helper hints — both reusable as-is.

### Label changes

| Current | Target |
|---|---|
| `Year / Period` | `Year` |
| `Confidentiality Level` | `Confidential Level` |
| `Vendor (if applicable)` | `Vendor/Customer Name` |
| `Upload into:` | `Upload to` |
| `--` select placeholder | `Select Segment` / `Select Department` / `Select Unit` |

These are display labels only. The `FIELDS` / `settings.columns` internal names are
untouched.

### Character limits

`maxLength={30}` plus a `<small>Max. 30 character</small>` hint on Document Name,
Project Name, and Vendor/Customer Name.

### Date input

The mockup shows a `dd/mm/yyyy` placeholder. A native `<input type="date">` renders its
format from the browser's locale and ignores markup — on an en-US machine it displays
`mm/dd/yyyy`. The native input is kept; the displayed format follows each user's locale.
This does not affect the stored value, which is still converted by `toSpDate()`. A custom
control was considered and rejected as disproportionate.

---

## 3. Form logic

Three changes. Everything else — the cascade, term-store loading, restricted-user path
resolution, the replace-file guard — is untouched.

### 3.1 Project Name state

New `projectName` state, cleared by `resetForm()` alongside the other fields. Written
unconditionally as a `FieldValue` — see the replace-file section below for why a blank
value must be sent rather than skipped.

Value is trimmed. The `ILLEGAL_NAME_CHARS` filter is **not** applied — that rule exists
for filenames, and this is a metadata value.

### 3.2 Vendor is written

Remove the carve-out at `Form.tsx:1054-1056` and push
`{ FieldName: settings.columns.vendor, FieldValue: vendor.trim() }`, again
unconditionally.

Vendor keeps feeding the auto-composed document name via `onVendorChange` /
`composeDocName`. That behaviour is unchanged — it now additionally persists.

### 3.3 Document Date stays required

The mockup shows no asterisk on Document Date, so it was briefly made optional. **An
updated client requirement on 2026-07-28 reverted this** — the field keeps its validation
check and its required asterisk. The mockup's missing asterisk is treated as an oversight.

One artefact of the round trip remains deliberately: the `FieldValue` is pushed behind an
`if (documentDate)` guard rather than sitting unconditionally in the `formValues` array
literal. Validation now guarantees a value, so the guard is redundant — it is kept as a
safety net because `toSpDate("")` produces the malformed `NaN/NaN/`, which SharePoint
rejects with a `HasException` surfacing as a confusing "Uploaded, but a field failed"
far from its cause.

---

## 4. Approval Document page

Add one row to the `metadata` array (`ApprovalDocument.tsx:376-382`) and let the existing
Vendor row start resolving:

```ts
["Project Name",         pick("ProjectName")],
["Vendor/Customer Name", pick("Vendor_x005f_x002f_x005f_CustomerName",
                              "Vendor_x002f_CustomerName")],
```

`ProjectName` has no encoded characters, so a single plain lookup suffices.
`Vendor_x002f_CustomerName` does, so it needs the double-encoded key first with the plain
name as a fallback. The existing `pick` helper already returns "—" when a value is absent.

The row label also changes from "Vendor" to "Vendor/Customer Name" to match the form.

### Existing documents

Files already in Staging carry neither value and will show "—" on both rows. Only uploads
made after this ships will populate them. This is expected and is not a defect.

---

## Testing

**Unit** — `src/shared/formModel.test.ts` is the existing suite. Add coverage for the
level-column rename resolving `GroupProjectName` to its label/Tid pair.

**Manual, against `/sites/ClarenceDMSTesting`:**

1. Business Segment upload with Project Name and Vendor filled → both appear on the
   Approval Document page.
2. Business Segment upload with both blank → Approval page shows "—" for both; no field
   error toast.
3. Upload with Document Date blank → **blocked** by validation, with "Document Date"
   named in the missing-fields toast.
4. Upload with Document Date filled → date stored correctly, name composes as
   `<Vendor>-<DD-MM-YY>`.
5. **Replace an existing file**, leaving Project Name and Vendor blank → both columns are
   cleared; the Approval page shows "—" for each, not the replaced file's values.
5. 31st character rejected in all three limited fields.
6. Group-led Projects upload → both "Project Name" (text) and "Group Project Name"
   (dropdown) visible; folder path uses the term, free text stored separately.
7. Restricted (non-privileged) user → cascade and "Uploading to:" breadcrumb still
   resolve correctly under the new card ordering.
8. Layout check at desktop and narrow widths against the mockup.

**Regression watch:** the card restructure moves a large block of JSX. Confirm the
loading state, the not-provisioned error, the segment radio group, and the empty-level
warning all still render in the folder card.

---

## Rollout

1. ~~Create `ProjectName` and `Vendor/CustomerName` in Staging.~~ **Done 2026-07-28.**
2. Remove both from the Staging library views (view-level hiding only — see Column
   visibility above).
3. Add the `col_projectName` row to `DMS Config`, and correct `col_vendor` to
   `Vendor_x002f_CustomerName`.
4. Rename the Group-led Projects level in the mode row's `Levels` JSON.
5. Amend `2026-07-21-segment-onboarding-plan.md` and
   `2026-07-16-tenant-seed-data-runbook.md` for the `GroupProjectName` rename.
6. Ship the web part.

Step 3 must precede step 6. The code fallbacks carry the correct names, so a missed
config row degrades to the fallback rather than failing — but a `col_vendor` row left
pointing at the deleted `Vendor` column overrides the fallback and breaks every upload.

## Replace-file behaviour — blank clears

Found in review after implementation; neither this spec nor the plan originally
considered it. **Resolved 2026-07-28: a blank field clears the column.**

`Form.tsx` uploads with `Files/Add(..., overwrite=true)` when the user confirms replacing
a same-named file. That swaps the file's *content* but reuses the **same list item**, so
any column not written survives from the previous upload.

The original conditional writes therefore left stale values: upload `Invoice.pdf` with
Vendor `Acme Trading`, later replace it leaving Vendor blank, and the Approval Document
page still showed `Acme Trading` — metadata describing the superseded file.

Project Name and Vendor are now written **unconditionally**; a blank field sends `""`,
which clears the Text column. The rationale is that a replaced file is a new document, so
its metadata should describe that document and nothing else.

Document Date needs no equivalent handling: it is required, so it is always written and
always refreshed. It also could not use the same fix — an empty value would reach
`toSpDate` and produce `NaN/NaN/`.

## Correction to an existing fixture

The contract fixture in `2026-07-28-dms-uat-test-plan.md` (§2.1) is wrong on three rows,
confirmed against the live `/fields` API:

| Column | Fixture claims | Actual |
|---|---|---|
| `Document_x0020_Type` | Text | `TaxonomyFieldType` |
| `Confidentiality_x0020_Level` | Text | `TaxonomyFieldType` |
| Year | `Year_x002f_Period`, Text | internal name `Year`, `TaxonomyFieldType` |

`Form.tsx` already uses the correct `Year`. The fixture should be corrected so it does not
fail its own test; out of scope for this change but worth doing alongside.
