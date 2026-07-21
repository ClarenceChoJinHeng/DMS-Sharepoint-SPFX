# Spec 1: Flat, Folder-Free Staging Review View

**Status:** Design pivoted by user (2026-07-10). Not yet implemented. Supersedes the original custom-web-part design (see "Superseded design" section at the bottom).

## Problem

Today, uploads land in a real nested folder tree inside the `Staging` library:

```
Staging/{Department|Projects}/{DeptLabel}/{SubTeam1}/{SubTeam2}/.../file.ext
```

(built in `handleUpload()`, [Form.tsx:568-633](../../../src/webparts/form/components/Form.tsx#L568-L633))

Approvers currently have to click through this folder tree to find items awaiting review. For a department with several sub-teams, reviewing 10 pending files can mean navigating 10 different folders. The ask: an Approver should go into the **native Staging library itself** and see files sitting there flat, with no folder navigation — not a separate custom page.

## Constraints carried over from existing architecture

- Folders must **stay physically as-is** — they are the unit SharePoint permissions are broken on (Approver = Design, Uploader = Contribute, scoped per-department folder in Staging), and the Auto-route automated flow moves approved files by **live folder path**. This spec only changes how Approvers *view* Staging, not how files are stored or routed.
- No Power Automate — stays consistent with the rest of this app's direct-REST architecture.
- Reuse the existing `ApprovalDocument` page/web part for the actual approve/reject action — this spec is a queue/index in front of it, not a replacement.

## All approaches considered, with drawbacks

Six distinct ways to solve "Approvers shouldn't have to click through folders" were identified across this design discussion. Ranked by how much they preserve the existing architecture (folders for permissions/Auto-route stay physically real) at the top, down to the most disruptive at the bottom.

### A. Native flat view + Command Set — **chosen, see Design below**
Turn on SharePoint's built-in "Show all items without folders" view setting on Staging, add a small List View Command Set extension so clicking a row still opens `ApprovalDocument`.
- **Drawback:** The "Folders" section that exposes this setting is missing from this tenant's classic View Edit page. Requires a REST workaround (PATCH the view's `Scope`/`ViewQuery` properties directly) instead of a UI toggle — **confirmed working and durable, see "REST workaround" section below.**
- **Drawback:** Still requires one small custom extension (the Command Set) to be built and maintained, even though it's much smaller than a full web part.

### B. Custom SPFx web part with a CAML query (superseded design, already built once)
A dedicated page/web part runs `GetItems` with `Scope="RecursiveAll"` and renders its own flat table.
- **Drawback:** Full custom code to build and maintain — data fetching, rendering, click-through, pagination, empty states — to reproduce something native SharePoint already does.
- **Drawback:** It's a *separate page*, not the native library — already rejected once by the user for this reason (2026-07-10 pivot). Only relevant again if the Command Set in Option A turns out to be infeasible.

### C. Lean into Metadata Navigation and Filtering (native metadata-tree browsing instead of folder-tree browsing)
Let SharePoint present a metadata-attribute tree (e.g. by Department, then Document Type) instead of the folder tree, using the same native feature that's suspected of hiding the Folders UI in Option A.
- **Drawback:** Doesn't produce a true flat list — still a tree, just organized by metadata instead of folders, so it only partially solves "no clicking through nested navigation."
- **Drawback:** Changes navigation library-wide (Uploaders would see it too, not just Approvers), and its interaction with the "Show all items without folders" view setting is not well documented — could conflict with Option A rather than complement it.

### D. Power Automate–maintained denormalized "index" list mirroring metadata
Keep the real files in folders untouched; a flow copies/mirrors metadata into a separate flat SharePoint list that Approvers browse, then click through to the real item.
- **Drawback:** Reintroduces Power Automate into a codepath that's currently direct-REST-only — against this project's established architecture ([[dms-architecture-direct-rest]]).
- **Drawback:** Two copies of the truth (real item + index row) that can drift out of sync (timing lag between upload and the flow mirroring it); more moving parts than Option A for the same end result.

### E. Search-driven view (PnP Modern Search / Content Search Web Part) over the search index
Query SharePoint's search index, filtered to the Staging library, as the flat listing mechanism.
- **Drawback:** Search index crawl latency — a newly uploaded file may not appear for minutes, which is a real problem for an approval queue where Approvers expect to see new items immediately.
- **Drawback:** More setup (managed properties, search schema) and still effectively another custom web part to build.

### F. Drop folders entirely — upload flat into `Staging/` root, organize purely by metadata columns
No physical folder structure at all; department/sub-team become just column values.
- **Drawback (permissions):** SharePoint permission scoping only works on securable objects (site/list/folder/item) — never on metadata values. Removing folders means the only way to keep per-department access control is unique permissions on every single item, which is a well-documented SharePoint performance anti-pattern at scale (hundreds/thousands of items).
- **Drawback (routing):** The Auto-route automated flow currently determines each approved file's destination by its live folder path. Without folders, that flow's logic would need a full rewrite to derive destination from metadata instead — a change to the Power Automate flow itself, out of scope for this SPFx app, and not yet designed.
- **Drawback (biggest):** Unlike A-E, this doesn't just change how Approvers *view* Staging — it changes how every uploader submits and how the whole system stores data, touching upload, permissions, and downstream routing simultaneously.

## Design (current — native flat view + Command Set)

### 1. Native flat view on the Staging library

Modern SharePoint document libraries support a built-in **"Show all items without folders"** option:

> List/Library → View settings (create or edit a view) → **Folders** section → select **"Show all items without folders"** (instead of the default "Show items inside folders").

This recursively flattens the display to every file across all folder levels, with **no code and no custom web part** — the physical folder structure underneath is completely untouched.

- Create a new view, e.g. **"Approver Flat View"**, on the `Staging` library.
- Folders setting: **Show all items without folders**.
- Columns to add: Department, Document Type, Sub-team/Project (`Department_x0020_Team` or `Project_x0020_Name` depending on upload mode), Document Date, Uploader (Author), Approval Status (SharePoint's built-in content-approval column, automatically present on any library with content approval enabled).
- Set as the default view for this library, or link Approvers directly to it via `...?viewid={flatViewGuid}`.

**Permission trimming is automatic and requires no extra work:** SharePoint always filters list/library views by the viewing user's ACLs, regardless of folder vs. flat view. Two layers already in place apply here exactly as they do to folder browsing:
- Per-department folder permission breaks (Approver = Design, Uploader = Contribute, scoped per department folder).
- **Draft Item Security = "Only users who can approve items (and the author of the item)"** — Approvers (has Approve Items) see every status in their permitted folders; Uploaders (Contribute, no Approve) see only Approved items (any author) + their own items regardless of status.

So an Approver opening this one flat view sees exactly the files across all their department's sub-team folders, any status, with folders never shown — for free, via native SharePoint behavior.

### 2. List View Command Set extension for click-through

Native views don't support "click a row → open a custom page with that item's ID" out of the box — clicking a file row in a normal library view opens the file's default preview/edit experience, not the existing `ApprovalDocument` approve/reject page.

To preserve the existing approve/reject workflow, add a small **SPFx List View Command Set** extension scoped to the Staging library:
- Registers a custom command (e.g. "Review") that appears when a row is selected, or overrides the default row click/open behavior.
- On invoke, navigates to the existing `ApprovalDocument` page/web part, passing the selected item's `Id` — reusing the approve/reject UI already built there. This spec adds no new approval logic.
- The command operates on the real underlying list item — its real `FileRef`/server-relative URL is unchanged by the flat view; flattening is purely a display setting on the view, not a change to storage.

This is the only piece of custom code this design needs — much smaller in scope than a full custom web part (no CAML query, no data-fetch/rendering/filtering logic to build or maintain).

### 3. Native library access for Approvers

Same decision as before: hide/restrict navigation so the flat view is the realistic entry point Approvers use day-to-day (the same audience-targeting pattern already used to hide the Upload Document nav link). This is a UX/discoverability measure, not a new security boundary — the underlying permission model is unchanged either way.

## REST workaround (confirmed working, 2026-07-11)

The classic View Edit page (`listedit.aspx`) on this tenant is missing its "Folders" section entirely — no "Show all items without folders" checkbox exists in the UI. Metadata Navigation and Filtering was suspected as the cause; **ruled out** (checked List Settings for a "Metadata navigation settings" link — zero matches via in-page search). Root cause of the missing UI section remains undiagnosed, but is irrelevant since the fix below bypasses the UI entirely via the same official `SP.View` REST API the UI itself calls.

### The fix

A view must have **both** properties set together — neither alone works:
- `Scope = 2` (`RecursiveAll`) — recurses through all nested folders. (`Scope = 3`/`FilesOnly` was tried first and confirmed **not** recursive — it only queries the current/root folder level, so on Staging, which has no files directly in its root, it returns an empty result.)
- `ViewQuery = "<Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>"` — filters out folder rows, leaving files only.

This combination is what the classic UI checkbox applies under the hood.

### Working script (run in browser console, on the site)

```js
const digest = await fetch("https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground/_api/contextinfo", {
  method: "POST",
  headers: { Accept: "application/json;odata=nometadata" }
}).then(r => r.json()).then(d => d.FormDigestValue);

fetch(`https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground/_api/web/lists/getbytitle('Staging')/views('${viewId}')`, {
  method: "POST",
  headers: {
    "Accept": "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=verbose",
    "X-RequestDigest": digest,
    "X-HTTP-Method": "MERGE",
    "IF-MATCH": "*"
  },
  body: JSON.stringify({
    __metadata: { type: "SP.View" },
    Scope: 2,
    ViewQuery: "<Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>"
  })
}).then(async r => { console.log(r.status); console.log(await r.text()); });
```

Replace `${viewId}` with the target view's GUID (find via `_api/web/lists/getbytitle('Staging')/views` and match on `Title`). A `204` response means success.

**Gotcha:** with `Content-Type: application/json;odata=verbose`, the body must include `__metadata: { type: "SP.View" }` or the request 400s with `"An entry without a type name was found, but no expected type was specified."`

### Durability — confirmed permanent, not fragile

Tested 4 ways against the "Test Flat" view; `Scope`/`ViewQuery` survived all of them unchanged:
1. Editing an unrelated setting via the classic Edit View page (ticking a Description checkbox) and saving.
2. Applying a filter via the modern UI's Filter funnel icon without saving.
3. Applying and **saving** a filter via the modern UI's Filter funnel icon.
4. Uploading a genuinely new file into a 5-level-deep nested folder — it appeared automatically in the flat view, confirming live recursion still works.

This is an official, supported technique — `Scope`/`ViewQuery` are documented public properties of Microsoft's `SP.View` REST/CSOM API, the same API the SharePoint UI calls internally. Using the browser console for a one-off admin action like this is a normal, commonly-used approach (not the standard for repeatable automation, which would use PnP PowerShell or a versioned script instead, but fine for a single tenant-side fix).

### Deployment status

| View | Library | Status |
|---|---|---|
| "Test Flat" (throwaway, sandbox tenant) | Staging | Fixed — `Scope=2` + `ViewQuery` applied, confirmed durable |
| "Approve/reject Items" (production-analogous, sandbox tenant) | Staging | Fixed — `204` confirmed applied |
| "All Documents" (sandbox tenant) | Staging | **Deliberately left untouched** — user decided (2026-07-11) to keep this view folder-based so other roles (e.g. Uploaders) can still browse folders; only the Approver-facing view should be flat |

**Not yet done:** reproduce this same PATCH against the real SD Guthrie production tenant (`dcidigitalcom.sharepoint.com` sandbox ≠ SDG's actual site) — the user noted this is tenant-specific and will need to be redone there once they have access, using the same script above against SDG's own "Approve/reject Items"-equivalent view GUID.

## Open risks / to confirm during implementation

- **Command Set scoping**: confirm a List View Command Set extension can be scoped to fire only on the Staging library (not site-wide) and only overrides/adds behavior for file rows (not folder rows, though folders won't appear in the flat view anyway).
- **Column availability in flat view**: confirm the Approval Status column and all custom metadata columns (Department, Document Type, etc.) render correctly in a flattened (non-grouped-by-folder) view — expected to work since these are just list columns, not folder-dependent, but not yet tested live.
- **Pagination**: native views handle large item counts with their own built-in paging (unlike the CAML `RowLimit` concern that applied to the custom web part approach) — lower risk here, but worth confirming behavior with a large item count.
- **Empty-state UI**: native views show SharePoint's standard "no items" message when a department has nothing pending — likely sufficient without custom empty-state work, but worth confirming it reads clearly to an Approver.

## Superseded design (parked, not deleted — kept for the confirmed facts below)

A custom `stagingReview` SPFx web part (`src/webparts/stagingReview/`) was scaffolded and partially tested this session — a separate page with a CAML `GetItems` query (`Scope="RecursiveAll"`), client-side status filter, and its own row rendering/click-through. **This was abandoned and deleted 2026-07-10** once the user clarified the actual intent: Approvers should land on the *native* Staging library itself, not a separate custom page. The native flat-view + Command Set design above replaces it entirely.

The CAML query work is still valuable and confirmed live, in case a future scenario needs a custom CAML query against Staging again:

### Confirmed facts (verified against the live tenant)

1. **Content approval is enabled** on the Staging library (List Settings -> Versioning Settings -> "Require content approval for submitted items?" = Yes).
2. **Draft Item Security = "Only users who can approve items (and the author of the item)."**
3. **REST/CAML field name dual-naming, confirmed live** (see [[sp-caml-internal-vs-rest-field-names]] memory):
   - CAML `FieldRef Name="FSObjType"` (not `FileSystemObjectType`, which is REST-only and throws `-2130575340` if used in CAML).
   - CAML `FieldRef Name="_ModerationStatus"` (not `OData__ModerationStatus`, which is REST-only, same error if used in CAML).
   - The JSON response is always keyed by the REST names (`FileSystemObjectType`, `OData__ModerationStatus`) regardless of which name was used to request the field.
4. **Moderation status enum**: `0` = Approved, `1` = Rejected (Denied), `2` = Pending, `3` = Draft/Scheduled.
5. **Working CAML query** (files only, recursive, with moderation status): `Scope="RecursiveAll"`, `Where FSObjType = 0`, `ViewFields` including `_ModerationStatus` — confirmed live returning correct results (item 354: `FileSystemObjectType: 0`, `OData__ModerationStatus: 2`).
6. **Exact field internal names** for department/sub-team metadata, read from `FIELDS` in [Form.tsx:10-20](../../../src/webparts/form/components/Form.tsx#L10-L20): `Department`, `Department_x0020_Type`, `Department_x0020_Team` (dept mode) / `Project_x0020_Name` (project mode), `Year_x002f_Period`, `Confidentiality_x0020_Level`, `DocumentDate`.

## Out of scope (parked as Spec 2)

Making Department<->Project term-set linking dynamic (arbitrary naming, no forced label match) and supporting deeper arbitrary-depth nesting — separate, unresolved design discussion, not blocking this spec.
