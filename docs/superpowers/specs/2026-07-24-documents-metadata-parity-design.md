# Documents Library Metadata Parity — Design

**Date:** 2026-07-24
**Status:** approved design — not yet implemented
**Trigger:** Approved files in the **Documents** library show blank metadata (Business
Segment, Department, Unit, Document Type, Document Date). Staging has the tags; Documents
does not.

> **Note (2026-07-24):** the Document Type column's internal name was subsequently migrated
> from `Department_x0020_Type` to **`Document_x0020_Type`** — see
> `2026-07-24-document-type-internal-name-migration-design.md`. Wherever this spec says
> `Department_x0020_Type`, read `Document_x0020_Type`, and the frozen-internal-name / "create as
> Department Type first" workaround no longer applies (the display name "Document Type" now yields
> the correct internal name directly).

## Root cause (verified against live `/fields`, 2026-07-24)

The Documents library is **missing 8 of the 11** metadata columns that Staging has. It is
not a flow-mapping subtlety — the columns literally do not exist on Documents, so the
metadata has nowhere to land. This affects **both** the Auto-route flow path *and* the
bulk-upload tool (which writes those same fields to Documents via `validateUpdateListItem`
and silently logs `HasException` for the absent columns — file uploads, tags drop).

| Internal name | Title | Type | Staging | Documents |
|---|---|---|---|---|
| `Business_x0020_Segment` | Business Segment | Text | ✅ | ❌ add |
| `BusinessSegmentTid` | Business Segment Tid | Text | ✅ | ❌ add |
| `Department` | Department | Text | ✅ | ❌ add |
| `DepartmentTid` | Department Tid | Text | ✅ | ❌ add |
| `Unit` | Unit | Text | ✅ | ❌ add |
| `UnitTid` | Unit Tid | Text | ✅ | ❌ add |
| `Department_x0020_Type` | Document Type | **Taxonomy** → Document Type set `0540e66e-7cb3-47ac-b0ef-4e3069387394` | ✅ | ❌ add |
| `DocumentDate` | Document Date | DateTime | ✅ | ❌ add |
| `Year_x002f_Period` | Year | Taxonomy | ✅ | ✅ |
| `Confidentiality_x0020_Level` | Confidentiality Level | Taxonomy | ✅ | ✅ |
| `Vendor` | Vendor | Taxonomy | ✅ | ✅ |

**Decision (client-confirmed):** Staging and Documents metadata columns must be **mirrored** —
identical internal names and types. Level columns stay **plain text** (label + `…Tid` pair),
not taxonomy, matching Staging and what the flow / bulk-upload already produce.

## Fix — three parts

### 1. Add the 8 missing columns to Documents (scripted, not hand-created)

Create the 8 columns on the Documents library with the **exact internal names and types**
above. **Scripted, not via the SharePoint UI**, because:

- Internal names are frozen at creation from the *display* name. `Department_x0020_Type` only
  arises if the column is first created display-named **"Department Type"** (CLAUDE.md gotcha:
  "Document Type" was created as "Department Type"). Creating it as "Document Type" in the
  UI yields `Document_x0020_Type` — a *different* column the flow won't recognise. A script
  sets the internal name directly and sidesteps the rename dance.
- The taxonomy column (`Department_x0020_Type`) must be **bound to the same term set**
  (`0540e66e-…`) and term store as Staging; taxonomy fields also carry a hidden note field.
  This is fiddly by hand and trivial with PnP.

**Chosen tool: PnP PowerShell** (one-off provisioning). Sketch:

```powershell
Connect-PnPOnline -Url "https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground" -Interactive
$list = "Documents"

# 6 plain-text level columns (internal name set explicitly to match Staging)
Add-PnPField -List $list -DisplayName "Business Segment" -InternalName "Business_x0020_Segment" -Type Text
Add-PnPField -List $list -DisplayName "Business Segment Tid" -InternalName "BusinessSegmentTid" -Type Text
Add-PnPField -List $list -DisplayName "Department" -InternalName "Department" -Type Text
Add-PnPField -List $list -DisplayName "Department Tid" -InternalName "DepartmentTid" -Type Text
Add-PnPField -List $list -DisplayName "Unit" -InternalName "Unit" -Type Text
Add-PnPField -List $list -DisplayName "Unit Tid" -InternalName "UnitTid" -Type Text

# Document Date
Add-PnPField -List $list -DisplayName "Document Date" -InternalName "DocumentDate" -Type DateTime

# Document Type — taxonomy, same term set as Staging. InternalName set explicitly to match.
Add-PnPTaxonomyField -List $list -DisplayName "Document Type" -InternalName "Department_x0020_Type" `
  -TermSetId "0540e66e-7cb3-47ac-b0ef-4e3069387394"
```

> Verify after running with the same `/fields` query used in diagnosis — all 11 present on
> Documents, internal names + `TypeAsString` matching Staging exactly (esp. `Department_x0020_Type`,
> not `Document_x0020_Type`).

Save the script under `scripts/` and keep it: Phase 2 segment sites need the identical Documents
schema (the column-parity risk already flagged in `2026-07-20-multi-site-storage-phase2-design.md`).

### 2. Auto-route flow must write metadata after the copy

The Power Automate **Copy file** action copies content only, never library metadata — so even
with the columns present the flow must set them explicitly. Add a metadata-write step after the
file lands in Documents, mirroring exactly what the web parts already do
(`BulkUpload.tsx:1064` / `Form.tsx`):

- **Send an HTTP request to SharePoint** → `POST .../Documents/items({id})/validateUpdateListItem`
  with a `formValues` array.
- **Format per field type** (the web part's proven format — do not invent another):
  - Taxonomy (`Department_x0020_Type`, `Year_x002f_Period`, `Confidentiality_x0020_Level`,
    `Vendor`): value = `"Label|GUID"` (single). Read the label + term GUID from the source
    Staging item.
  - Level text columns (`Business_x0020_Segment`+`BusinessSegmentTid`, `Department`+`DepartmentTid`,
    `Unit`+`UnitTid`): plain label into the label column, term GUID into the `…Tid` column.
  - `DocumentDate`: `M/D/YYYY` (US site locale — ISO is rejected; CLAUDE.md gotcha #1).
- **Check `HasException`** on each returned result (HTTP 200 is returned even on field errors —
  CLAUDE.md gotcha #4).

Source values come from the Staging item's fields (the flow already reads the item to route it).
Simplest: read the Staging item's `FieldValuesAsText` + the raw term GUIDs and re-emit them.

> Cannot be edited from this repo — the flow lives in Power Automate. Deliverable here is the
> exact `formValues` shape; the flow author wires it up. The bulk-upload tool needs no change —
> once the columns exist, its existing write starts succeeding.

### 3. `Approval Status` column on Documents (default "Approved")

Documents has no content approval (Staging-only), so status is surfaced as a plain column:

- Add **`Approval Status`** — Choice (`Approved`) or single-line text — **default value `Approved`**.
- Covers both paths uniformly: flow-routed files (already approved to get there) and
  bulk-uploaded historical files (pre-approved by definition).
- **Caveats:** the default applies only to **new** items, so files already in Documents need a
  **one-time backfill** (Edit in grid view, or a `validateUpdateListItem` loop in the same script).
  Confirm the flow's Copy action triggers the default; if not, have the flow set it explicitly.

## Verification

1. `/fields` on Documents returns all 11 metadata columns, names + types matching Staging.
2. Approve a new file in Staging → after the flow runs, the Documents copy shows **every** tag
   populated (Business Segment → Vendor + Document Date), not just Year/Confidentiality/Vendor.
3. Bulk-upload a test file → its Documents item is fully tagged (no `tagFailed` in the results).
4. Every Documents item shows `Approval Status = Approved` (new ones by default; existing ones
   after backfill).

## Future segments — same rule applies

This fix creates only the **head-office pilot** columns (`Business Segment`/`Department`/`Unit`
+ Tids, `Document Type`, `Document Date`). The plain-text label + `…Tid` design is deliberately
segment-agnostic — segments that share level names (e.g. Projects reuse `Department`/`Unit`, SDGI
reuses `Department`) reuse these columns. But a 2027 segment with **new** level names introduces
new columns — `Region`, `EstateMill`, `Refinery`, `ITOperatingUnit`, `ProjectName` (+ their Tids)
— and each must be created in **both Staging and Documents** at onboarding, or this exact blank-
metadata bug returns on that segment. The onboarding plan's column step
(`2026-07-21-segment-onboarding-plan.md` §5) now states this explicitly.

## Out of scope

- Making the level columns taxonomy (deferred; text mirrors Staging and is enough now).
- Changing the flow's file-copy / rename-proof routing logic (untouched).
- Retroactively re-tagging historical approved files beyond the `Approval Status` backfill
  (their source metadata may not exist to copy — separate cleanup if needed).

## Sequencing

1. Run the column script on Documents (+ `Approval Status`). Verify `/fields`.
2. Backfill `Approval Status` on existing Documents items.
3. Add the metadata-write step to the Auto-route flow; test one approval end-to-end.
4. Re-test bulk-upload tagging (should now succeed with no code change).
