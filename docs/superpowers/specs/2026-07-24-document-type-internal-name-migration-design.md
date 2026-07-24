# Document Type Internal-Name Migration — Design

**Date:** 2026-07-24
**Status:** approved design — not yet implemented
**Amends:** `2026-07-24-documents-metadata-parity-design.md` (that spec created Documents columns
using the OLD name `Department_x0020_Type`; this migration changes the target to
`Document_x0020_Type` on both libraries).

## Goal

Rename the Document Type column's **internal name** from the frozen accident
`Department_x0020_Type` (created as "Department Type") to the clean `Document_x0020_Type`,
consistently across **Staging** and **Documents**. Display name stays "Document Type"
throughout — users see no change.

**Scope is Document Type only.** Document Date is already clean (`DocumentDate` in Staging and
in the code); it needs no migration — just delete the stray `Document_x0020_Date` look-alike on
Documents, leaving the correct `DocumentDate` the browser snippet created.

> SharePoint internal names are immutable, so "rename" = create-new + migrate-data + repoint-
> code/flow + delete-old. The ordering below avoids any window where uploads write to a column
> that doesn't exist.

## Code footprint (verified by grep, 2026-07-24)

| File / line | Current | After |
|---|---|---|
| `src/webparts/form/components/Form.tsx:29` | `documentType: "Department_x0020_Type"` | `"Document_x0020_Type"` |
| `src/webparts/bulkUpload/components/BulkUpload.tsx:46` | `documentType: "Department_x0020_Type"` | `"Document_x0020_Type"` |
| `src/webparts/approvalDocument/components/ApprovalDocument.tsx:323` | `pick("Department_x005f_x0020_x005f_Type", "Department_x0020_Type")` | `pick("Document_x005f_x0020_x005f_Type", "Document_x0020_Type")` |
| `CLAUDE.md`, `.claude/backlog.md`, older specs | prose notes | update to the new name |

The `_x005f_x0020_x005f_` form in ApprovalDocument is FieldValuesAsText's double-encoding of the
`_x0020_` (space) in the internal name — it becomes `Document_x005f_x0020_x005f_Type` for the new
name (see `sp-fieldvaluesastext-underscore-encoding` memory). The comment at
`ApprovalDocument.tsx:18-19` should be updated too.

**Unaffected:** the Document Type *term set* GUID (`0540e66e-…`, `DEFAULT_SETTINGS.termSets.documentType`)
is separate from the column internal name — the new column binds to the same term set. No DMS
Config change.

## Migration order (no upload-breakage window)

The new column must exist **before** the code switches to it, and the old column lingers until
**after** cutover. Two columns can't share the display name "Document Type" in one list, so the
old one is renamed out of the way first.

1. **Free the display name.** In **Staging**, rename the old column's *display* to
   `Document Type (legacy)` (internal stays `Department_x0020_Type` — code still works, it uses
   the internal name).
2. **Create the new column** `Document Type` (internal `Document_x0020_Type`), Managed Metadata
   bound to term set `0540e66e-…`, on **Staging AND Documents**.
   - On **Documents**: if the original `Document_x0020_Type` look-alike was *not* deleted, it is
     already the target — just bind it to the term set and ensure its display is "Document Type".
     If it was deleted, create it fresh.
3. **Backfill Staging data** — copy every item's value from `Department_x0020_Type` →
   `Document_x0020_Type` (validateUpdateListItem in `"Label|GUID"` form, or Edit-in-grid for the
   handful of Staging items). Documents items were blank, so nothing to backfill there.
4. **Deploy the code change** (3 files above) — rebuild `.sppkg`, bump version, redeploy. New
   uploads now write `Document_x0020_Type`, which already exists.
5. **Update the Auto-route flow** — change its Document Type field reference to
   `Document_x0020_Type` (part of the flow's metadata-write step from the parity spec).
6. **Verify** (below).
7. **Delete** the old `Department_x0020_Type` (`Document Type (legacy)`) column from Staging.

> Because this is the sandbox with a small Staging item count, a simpler brief-maintenance-window
> order (swap column, then deploy) is also acceptable — but the order above is the safe default
> and the one to use if this is ever run against a populated/production site.

## Verification

1. `/fields` on **both** libraries shows `Document_x0020_Type` (Taxonomy, bound to the Document
   Type term set) and **no** `Department_x0020_Type`.
2. Upload via the form → the file's Document Type tag lands (check `HasException` is clean).
3. Approve a file → Auto-route copy in Documents shows Document Type populated.
4. ApprovalDocument web part still displays Document Type for existing and new items.
5. `grep -rn "Department_x0020_Type" src/` returns nothing.

## Rollback

If a step fails before deletion (step 7), the old `Department_x0020_Type` column still holds the
original data and the previous package still targets it — redeploy the previous `.sppkg` and
rename the legacy column's display back to "Document Type". Nothing is destructive until step 7.

## Out of scope

- Any other column's internal name (Business Segment/Department/Unit + Tids, Year, Confidentiality,
  Vendor, DocumentDate all stay as-is — they are already sensible).
- The parity spec's other 7 columns (unchanged by this migration).
