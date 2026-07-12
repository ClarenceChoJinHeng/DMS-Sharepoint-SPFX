# Spec: Rename-Proof Uploads (address folders by UniqueId, not name)

**Status:** Design approved by user (2026-07-12). Not yet implemented.
**Related:** [staging-review-flat-list.md](./staging-review-flat-list.md) — the "Spec 2 / arbitrary naming decoupling" parked there is what this spec delivers. [[dms-architecture-direct-rest]], [[dms-onboarding-webpart]], [[dms-autoroute-path-based-routing]].

## Problem

The upload form finds the destination folder by **name**. It builds a path from the *labels* of the term-store values the user selected, then requires that exact path to already exist before uploading:

- Path built from labels: [Form.tsx:615-633](../../../src/webparts/form/components/Form.tsx#L615-L633)
- Folder-exists check that aborts the upload on any mismatch: [Form.tsx:664-739](../../../src/webparts/form/components/Form.tsx#L664-L739)
- Upload POST into the name-built path: [Form.tsx:742](../../../src/webparts/form/components/Form.tsx#L742)

Because the destination is addressed by name, **renaming any folder in the path breaks every upload that targets it** — the built path no longer matches, the existence check fails, and the file is never uploaded. `normalizePathSegment` ([Form.tsx:53-54](../../../src/webparts/form/components/Form.tsx#L53-L54)) is effectively a no-op and provides no protection.

The client requires:
- Folders may be **renamed freely** with **zero intervention** — uploads must keep working.
- Folder names may be **fully arbitrary** — no requirement that a folder name match its term label, even at first setup.
- Folders may be **created natively** (SharePoint "New → Folder"), not only through the Onboarding web part.
- Very large existing structure (hundreds of thousands of files, years of history). **No files or folders may be moved.**

## Core principle

Stop addressing folders by name. Address them by their permanent **`UniqueId` (GUID)**.

Every SharePoint folder is a list item with a `UniqueId` and integer `Id` that **do not change on rename or move**. REST resolves a folder from it via `GetFolderById(guid'<UniqueId>')`, which always returns the folder at its *current* location. This is the documented, sanctioned way to reference a folder durably.

The term↔folder link is arbitrary (names need not match), so it must be **recorded once by a human** and then stored as `TermGuid → FolderUniqueId`. Nothing can infer it automatically when names are arbitrary.

## What does NOT change

- **Files and folders never move.** No content is relocated, renamed, or created by this feature.
- **Permissions** (per-department inheritance breaks: Approver = Design, Uploader = Contribute) are untouched — they live on the physical folders, which are unchanged.
- **Auto-route** flow is untouched — it still routes approved files by live folder path.
- **Metadata tagging** on upload (Department, sub-team/project, Year/Period, etc.) is unchanged.
- The Onboarding web part is **not** a dependency (folders may be created natively). It may optionally write config rows later, but the system does not rely on it.

## Deliverables

Three pieces.

### 1. Config list — `DMS Folder Map` (new SharePoint list)

The rename-proof lookup table. One row per **selectable leaf term** (the term whose folder a file lands in). Row count is bounded by the term store, **not** by file count.

| Column | Type | Purpose |
|--------|------|---------|
| `Title` | Single line text | Human-readable label, e.g. `Account > AI Engineer` — **display only** |
| `TermGuid` | Single line text (**indexed**) | The GUID of the destination term — the **lookup key** |
| `FolderUniqueId` | Single line text | The folder's `UniqueId` GUID — the **resolution target** |
| `FolderUrl` | Single line text | Cached server-relative path — **display only, never used to resolve** (may go stale after a rename; that is acceptable) |
| `Section` | Choice: `Departments` / `Projects` | Which upload mode / term set — display + disambiguation |

- `TermGuid` is the join key because term GUIDs are globally unique and a selected leaf term uniquely identifies one destination folder (the leaf encodes its full ancestry). The upload uses `selectedChoice.term.id` as this key.
- `FolderUrl` is a convenience for admins reading the list; the system never reads it for resolution. Only `FolderUniqueId` is authoritative.
- **"Who mapped it / when" is free** — every SharePoint list automatically tracks built-in `Author` (Created By), `Created`, `Editor` (Modified By), and `Modified` on each row. No custom column needed; the reconciliation tool surfaces these in its UI so admins can see who registered each folder and when. No custom audit column is added.
- **Only leaf terms need rows.** Uploads always target a leaf sub-team/project folder — the sub-team/project field is required ([Form.tsx:573-576](../../../src/webparts/form/components/Form.tsx#L573-L576)), so a file can never land directly in a department root folder. Department-level terms therefore do **not** need rows.

Sample row (synthetic): `Title="Account > AI Engineer"`, `TermGuid="94ce322b-...-leaf"`, `FolderUniqueId="8f3a...c1"`, `FolderUrl="/sites/.../Staging/Departments/Acct-01/Team-AI"`, `Section="Departments"`.

### 2. Reconciliation web part (new small admin tool)

**Read-only against the folder tree** — it never creates, moves, renames, or deletes any folder or file. Its only writes are rows in `DMS Folder Map`.

Purpose: populate/maintain the config list by capturing each folder's `UniqueId` and pairing it to a term.

**Two entry modes:**

1. **Name auto-match (convenience).** For each leaf term, compute the expected folder path using the same logic the form uses today ([Form.tsx:615-633](../../../src/webparts/form/components/Form.tsx#L615-L633)), resolve it to a folder, capture its `UniqueId`, and pre-fill the row. Because today's uploads only work when names match, everything that currently works auto-matches for free. Admin reviews and confirms.
2. **Two-pane manual pairing (always works, v1 core).** A two-column screen:
   - **Left:** leaf terms that are not yet mapped (read from the term store).
   - **Right:** a dropdown of folders that actually exist in the target library (read live from SharePoint).
   - Admin picks the matching folder -> tool captures its `UniqueId` -> writes the row -> term drops off the "unmapped" list.
   - A running counter shows `mapped / remaining`.

**Behaviour:**
- **Idempotent** — re-runnable anytime. Skips terms already mapped; surfaces only new/unmapped terms and folders.
- **Handles native folder creation** — a folder created natively (or any way) needs one row before uploads target it. The admin registers it here: auto-match if the name lines up (confirm-only), otherwise pick it manually. Never requires typing a raw GUID.
- Terms it cannot auto-match are listed for manual pairing — this is the only manual effort, and only at add time, **never on rename**.

**Out of scope for v1 (fast-follow):** spreadsheet/CSV import for bulk seeding. Same output (rows in the config list), added only if the initial manual seed proves painful. The two-pane UI is required regardless (it is the ongoing "register a new folder" tool), so it ships first.

### 3. Upload form change — `Form.tsx`

Replace name-based resolution with ID-based resolution.

**Remove/replace:** the label->path build ([615-633](../../../src/webparts/form/components/Form.tsx#L615-L633)) and the folder-exists probe/abort block ([664-739](../../../src/webparts/form/components/Form.tsx#L664-L739)) as the *primary* resolution mechanism.

**New flow on upload:**
1. User selects term(s) as today; take the destination leaf term GUID (`selectedChoice.term.id`).
2. Look it up in `DMS Folder Map` (`?$filter=TermGuid eq '<guid>'&$select=FolderUniqueId,Title`).
3. If **no row** -> stop with a clear message: *"This folder hasn't been mapped yet. Ask an administrator to register it in the reconciliation tool."* (replaces the old "hasn't been set up" message).
4. If a row exists -> read `FolderUniqueId`.
5. **Duplicate check:** `GetFolderById(guid'<FolderUniqueId>')/Files('<finalName>')` — 200 = exists (block), 404 = clear to upload.
6. **Upload:** `POST GetFolderById(guid'<FolderUniqueId>')/Files/Add(url='<finalName>',overwrite=false)` with the raw `File` body (per existing gotcha #6).
7. Set metadata via `validateUpdateListItem` on the returned item exactly as today (unchanged).

Because every REST call targets the folder by `UniqueId`, a rename or move of that folder does not affect the upload.

**Keep** the detailed HTTP-status/error-body logging already added ([Form.tsx:669-693](../../../src/webparts/form/components/Form.tsx#L669-L693)) — adapt it to the new calls so failures still surface the real status/body, not a generic message.

## Lifecycle summary

| Event | Admin action |
|-------|--------------|
| Folder **renamed** | **None** — uploads keep working (the point of this spec) |
| Folder **moved** | **None** — `UniqueId` unchanged |
| Folder **created** (native or onboarding) | Register once in the reconciliation tool (auto-match confirm, or manual pick) |
| Folder **deleted** | Optional: remove its now-stale row |
| Files uploaded daily | None |

## Open risks / to confirm during implementation

- **`GetFolderById` + `/Files/Add` and `/Files('name')`** — confirm both compose correctly against this tenant (they are standard `SP.Folder` REST surface, expected to work).
- ~~**Terms with no leaf selection**~~ — **resolved**: the sub-team/project field is required ([Form.tsx:573-576](../../../src/webparts/form/components/Form.tsx#L573-L576)), so uploads always target a leaf folder. Only leaf terms need rows.
- **Config list lookup performance** — `TermGuid` must be indexed; the list stays small (bounded by term count), so this is low risk.
- **Stale `FolderUrl`** — display-only, so staleness is cosmetic; confirm no code path reads it for resolution.
- **Deleted-folder rows** — a stale row pointing at a deleted `UniqueId` will fail `GetFolderById`; upload should surface a clear "folder no longer exists" message rather than a generic error.
- **Auto-route interaction** — Auto-route reads live folder path and is unaffected, but confirm no assumption there depends on the folder name matching a term label.
