# Reconciliation as Term-Store Provisioner — Design

**Status:** proposed. Extends the Folder Reconciliation web part from a read-only
term→folder mapper into a term-store-driven provisioner that creates, locks, and
secures the folder tree in one pass. Implement after Phase 1 is verified.

## Goal

Eliminate manual folder creation + permissioning. The **term store is the structure
authority**, and reconciliation already computes the exact folder path for every term.
So in one walk it can: create the folder, break inheritance, assign the right groups
from DMS Group Map, and write the DMS Folder Map row.

## Behaviour — per term, walking each mode's term set

For every term (at every level — segment, department, unit):
1. **Create** the folder at its computed path if missing (idempotent; skip if it exists).
2. **Break inheritance** on it (every level, so nothing is visible until groups are set).
3. **Assign groups** mapped to *that level's term GUID* in DMS Group Map, by role.
4. **Write** the DMS Folder Map row (term → folder UniqueId) — Staging only.

Runs against **both** libraries (Staging + Documents), with a **confirm step**, a full
**log**, and safe **re-runs** (existing + already-unique folders are left untouched so a
re-run never clobbers manually assigned groups).

## Role → SharePoint permission convention

| Group Map role | Staging | Documents |
|---|---|---|
| `MEMBER` | **Read** | **Read** |
| `UPL` (uploader) | **Contribute** | *skipped* |
| `APR` (approver) | **Design** | *skipped* |
| `GLOBAL` | skipped (not folder-scoped) | skipped |

- **Staging** = working/approval library → all roles apply.
- **Documents** = read-only approved store → only the **MEMBER (viewer)** groups get
  **Read**; uploaders/approvers get no special access there.
- Segment/dept folders (MEMBER rows only) end up Read-locked to their tier; unit folders
  get Read + Contribute + Design in Staging, Read in Documents.

## Traversal note

Breaking inheritance at every level and assigning only that tier's group means a pure
segment member cannot see a department's folders, etc. — matching the tiered model.
SharePoint grants **Limited Access** up the parent chain automatically for anyone assigned
on a deeper folder, so an authorised unit user can still traverse to reach it.

## Folder Map scope

Map **Staging only** (that is what the upload form routes by, via UniqueId). Documents
folders are created + secured but not mapped — the view web part reads them by path and
the Phase 2 auto-route flow routes by path. (Revisit if a Documents map is ever needed.)

## Idempotency + safety

- Folder exists + already has unique permissions → **skipped** (no re-break, no clobber).
- Folder exists but still inheriting → break + assign.
- Folder missing → create + break + assign.
- **Confirm step** before any writes; optional dry-run listing what *would* change.
- Log **warnings** for folders that got locked with **no group-map rows** for their term
  (locked admin-only until an admin assigns groups) so nothing is silently orphaned.

## Multi-segment: config-driven, data-gated

The tool is **config-driven** over the modes in `DEFAULT_MODES` / DMS Config — it handles
all five segments (GHO, Upstream, SDGI, I&T, Group-led Projects) with **no code change**.
What activates each segment is **data**:

- its **real term-set GUID** (the non-GHO modes currently hold `REPLACE-*-TERMSET-GUID`
  placeholders), and
- its **DMS Group Map rows** (groups per tier).

Only **GHO** is seeded today, so a run now provisions GHO only. Non-GHO modes with a
placeholder term set return no terms (`loadTops(...).catch(() => [])`) and are skipped
safely. Onboarding a segment later = capture its term-set GUID → add to config → seed its
group-map rows → run reconciliation.

## Overlap with the Folder Manager "Copy structure → Documents" button

This provisioner supersedes that button for Documents (it now creates + locks + assigns
reader groups on Documents from the term store). Keep the copy button for now as a lighter,
group-map-independent option; flag it for removal once the provisioner is proven.

## Implementation notes (reuse existing patterns)

Reconciliation will need the machinery FolderManager/Onboarding already use:
- load **DMS Group Map** (term GUID → {groupId, role}); load **role definitions** (name→Id);
- `folders/AddUsingPath` (create), `breakroleinheritance`, `ensureuser`
  (claim `c:0o.c|federateddirectoryclaimprovider|{GroupId}`), `addroleassignment`;
- reuse `writeFolderMapping` / `resolveFolderByPath` for the map step;
- keep the site-collection owner group as Full Control on each broken folder (as
  FolderManager does) so admins retain access.

## Risks / watch-items

- **Requires Manage Permissions** — must be run by an admin (site owner/SCA); uploaders
  can't break inheritance.
- **Large trees / throttling** — sequential REST per folder; add small backoff if needed.
- **Role-convention assumptions** — MEMBER→Read etc. are fixed; if a client wants different
  levels, this becomes config.
- **Documents reader = MEMBER groups** assumes the tier MEMBER groups are the correct
  viewers of approved docs; revisit if a dedicated reader group is introduced.

## Sequencing

1. Phase 1 verified live first.
2. Build the provisioner; test on GHO (fully seeded) end-to-end.
3. Seed the other four segments' term-set GUIDs + group-map rows as they onboard, then run.
