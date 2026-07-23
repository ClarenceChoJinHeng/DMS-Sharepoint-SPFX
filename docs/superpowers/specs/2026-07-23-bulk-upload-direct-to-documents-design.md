# Bulk Upload — direct to Documents (temporary tool)

**Date:** 2026-07-23
**Status:** Approved, implementing
**Web part:** `bulkUpload` / "Bulk Upload"

## Purpose

A duplicate of the `form` upload web part for **temporary** use. Two behaviour changes:

1. Accepts **up to 50 files** per submission instead of one.
2. Writes **directly into the Documents library**, bypassing Staging entirely — so
   no content approval and no Auto-route flow involvement.

Everything else is a full clone: metadata tagging, Entra-group mode detection,
term-store cascade, per-Unit folder routing, Year/Document-Type subfolders.

## Decisions

| Question | Decision |
|---|---|
| Fidelity | Full clone, retarget only — keep tagging, group detection, term paths, Unit routing |
| Documents folder routing | **Path-swap** from the Staging map (Documents mirrors Staging exactly) |
| Rename | **Per-file rename rows** — one name box per selected file |
| Failure handling | **Continue on failure**, per-file success/failure report at the end |
| Max files | 50 (hard cap, enforced at selection) |

## Architecture

New SPFx web part cloned from `form`, its own bundle so it deploys and is added to
a page independently:

```
src/webparts/bulkUpload/
  BulkUploadWebPart.ts
  BulkUploadWebPart.manifest.json   (new component GUID + alias)
  components/BulkUpload.tsx
  components/IBulkUploadProps.ts
  loc/{en-us.js, mystrings.d.ts}
```

- Registered in `config/config.json` as bundle `bulk-upload-web-part` +
  localized resource `BulkUploadWebPartStrings`.
- `config/package-solution.json` version bumped.
- **`src/shared/*` is reused, not forked.** `dmsFolderMap` and `formModel` keep
  their single implementation. Only the component/flow is duplicated.

## Documents-library routing (the one real design change)

The `DMS Folder Map` list maps a term GUID to a **Staging** folder's
`FolderUniqueId`. The Documents library's Unit folder is a *different* folder with
a *different* UniqueId, so retargeting is not a library-name swap.

Resolution order:

1. `lookupFolderMapping(leafTermGuid)` → Staging Unit folder `FolderUniqueId`.
2. `resolveFolderServerUrl(uniqueId)` → the Unit folder's **current** Staging path
   (rename-proof).
3. **Path-swap** the library segment: `/<web>/Staging/<rest>` → `/<web>/Shared Documents/<rest>`.
   `Shared Documents` is the Documents library's real URL segment (its display
   title is `Documents`) — matching the Auto-route flow's own
   `concat('/Shared Documents/', …)` destination.
4. `resolveFolderByPath(swappedPath)` → the Documents Unit folder's UniqueId.
   - If this returns `null`, **abort with a clear error**. The Unit folder is
     *not* auto-created: creating it here would produce a folder inheriting the
     Documents root ACL, silently widening access. Missing folder = admin task.
5. `ensureFolder(Year)` then `ensureFolder(Document Type)` under it — these are
   safe to create, they inherit the Unit folder's ACL exactly as in Staging.
6. Upload each file into the resulting folder by UniqueId.
7. Metadata written via `validateUpdateListItem` against the **`Documents`** list
   (display title), not `Staging`.

All folder/file lookups use the OData parameter-alias form
(`GetFolderByServerRelativeUrl(@f)?@f='…'`) per the project gotcha — inline quoted
literals 400 on deep paths.

## Multi-file UI

- File input gains `multiple`. Selection over 50 is rejected with a toast; the
  extension allow-list is applied per file at selection time.
- Selected files render as a list; each row shows the original name, size, a
  **rename box**, and a remove (✕) button.
- The resolved final name (`buildUploadName(original, typed)`) previews per row,
  so the extension-preserving behaviour stays visible.
- One shared metadata form (segment/levels/year/doc type/date/confidentiality/
  vendor) applies to the whole batch — these drive the folder path, so the batch
  necessarily lands in one folder.

## Batch upload behaviour

Files upload **sequentially** (not parallel) to keep SharePoint throttling and
progress reporting predictable. For each file:

1. Duplicate probe → if a same-name file exists, record `skipped (duplicate)` and
   continue. `overwrite=false` is kept, so an existing file is never clobbered.
2. `Files/Add` → on non-OK, record `failed` with the HTTP status + SharePoint
   error message, and continue.
3. Fetch the new item id, then `validateUpdateListItem` with the batch metadata.
   A metadata failure records `uploaded, tagging failed` — the file is not deleted.

Progress shows `Uploading 7 of 50…`. On completion a **result panel** lists every
file with its outcome, and a summary line (`45 uploaded, 3 skipped, 2 failed`).
Successful rows are removed from the pending list; failed/skipped rows remain so
the user can rename and retry without re-picking everything.

## Explicitly out of scope

- No changes to `form`, `Staging`, the Auto-route flow, or `src/shared/*` logic.
- No parallel/chunked upload for large files (the existing single-shot
  `Files/Add` limit applies — fine for the document types allowed).
- No new folder map data and no Documents-side map list.

## Risk note

This tool writes into Documents **without approval**. That is the explicit intent
("temporary"), but it means uploads bypass the review gate the rest of the DMS
depends on. It should be surfaced on a restricted page and removed when no longer
needed. The web part title and on-page subtitle state this plainly.
