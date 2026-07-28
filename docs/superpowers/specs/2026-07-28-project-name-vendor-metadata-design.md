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
| Document Date | Becomes **optional** (mockup shows no asterisk) |
| Date input format | Keep native `<input type="date">`; format follows browser locale |

---

## 1. Data model

### Staging library columns

| Display name | Internal name | Type | Status |
|---|---|---|---|
| Project Name | `ProjectName` | Single line of text | **Must be created before ship** |
| Vendor | `Vendor` | Single line of text | Exists; starts being written |

`Vendor` is confirmed as a **Text** column by the UAT contract fixture in
`2026-07-28-dms-uat-test-plan.md`, so writing a plain string to it is safe. The older
`2026-07-24-documents-metadata-parity-design.md` lists it as Taxonomy; that entry is
stale and the newer UAT contract governs.

`ProjectName` has no spaces in its display name, so SharePoint will not encode the
internal name. Both columns are read back from `FieldValuesAsText` under their plain
internal name — the `_x005f_x0020_x005f_` double-encoding problem documented in the
`sp-fieldvaluesastext-underscore-encoding` memory only affects columns whose display
name contains a space (`Document Type`, `Confidentiality Level`).

> **Blocking prerequisite:** if `ProjectName` does not exist when this ships, the
> `validateUpdateListItem` call returns `HasException` for that field and the user sees
> "Uploaded, but a field failed". Create the column first.

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

| Setting row | Fallback | Status |
|---|---|---|
| `col_projectName` | `ProjectName` | **New** |
| `col_vendor` | `Vendor` | Already exists (see the migration runbook) |

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

New `projectName` state, cleared by `resetForm()` alongside the other fields. Pushed as a
`FieldValue` only when non-blank, so an untouched field does not write an empty string.

Value is trimmed. The `ILLEGAL_NAME_CHARS` filter is **not** applied — that rule exists
for filenames, and this is a metadata value.

### 3.2 Vendor is written

Remove the carve-out at `Form.tsx:1054-1056` and push
`{ FieldName: settings.columns.vendor, FieldValue: vendor.trim() }` when non-blank.

Vendor keeps feeding the auto-composed document name via `onVendorChange` /
`composeDocName`. That behaviour is unchanged — it now additionally persists.

### 3.3 Document Date becomes optional

Remove `Document Date` from the `missing` validation array (`Form.tsx:807`), and drop the
asterisk from its label.

The `FieldValue` must be **skipped entirely when blank**, not sent as an empty string:
`toSpDate("")` splits an empty string and produces the malformed `NaN/NaN/`, which
SharePoint rejects with a `HasException` on that field.

`composeDocName` already tolerates an empty date — it filters falsy parts — so the
auto-composed name degrades to vendor-only with no change needed.

---

## 4. Approval Document page

Add one row to the `metadata` array (`ApprovalDocument.tsx:376-382`) and let the existing
Vendor row start resolving:

```ts
["Project Name", pick("ProjectName")],
["Vendor",       pick("Vendor")],
```

Both are plain-text columns with space-free display names, so a single plain-internal-name
lookup suffices — no double-encoded variant needed. The existing `pick` helper already
falls back to "—" when a value is absent.

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
3. Upload with Document Date blank → succeeds; document name falls back to vendor-only;
   no `NaN/NaN/` error.
4. Upload with Document Date filled → date stored correctly, name composes as
   `<Vendor>-<DD-MM-YY>`.
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

1. Create the `ProjectName` column in the Staging library.
2. Remove `ProjectName` and `Vendor` from the Staging library views (view-level hiding
   only — see Column visibility above).
3. Add the `col_projectName` setting row to `DMS Config`.
4. Rename the Group-led Projects level in the mode row's `Levels` JSON.
5. Amend `2026-07-21-segment-onboarding-plan.md` and
   `2026-07-16-tenant-seed-data-runbook.md` for the `GroupProjectName` rename.
6. Ship the web part.

Step 1 must precede step 6, or every upload reports a field failure.
