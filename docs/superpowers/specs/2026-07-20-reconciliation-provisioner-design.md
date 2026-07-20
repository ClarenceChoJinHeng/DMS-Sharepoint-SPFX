# Reconciliation as Term-Store Provisioner — Design

**Status:** in progress. Turns Folder Reconciliation from a read-only term→folder
mapper into a term-store-driven provisioner that **creates folders and breaks
inheritance** in one pass. It lives as a **third tab inside the Folder Manager
("Manage Folders") web part** — Staging / Documents / Folder Reconciliation — so
users have one tool, not two pages. The standalone Reconciliation web part is
retired, and the Folder Manager "Copy structure → Documents" button is removed
(this tab supersedes it).

## Decisions (2026-07-20) — supersede the earlier draft below

- **Group assignment is now IMPLEMENTED (2026-07-20c).** Reconciliation auto-assigns
  the DMS Group Map groups to each folder by role using the table below
  (MEMBER→Read, UPL→Contribute, APR→Design; GLOBAL skipped). Documents gets
  MEMBER/viewer groups only. The join key is **DMS Group Map `UnitTermGuid`** →
  the folder's tier term: for the segment container that term is the **term-set
  GUID** (`assignTerm`), for dept/unit folders it is the term GUID. A term can have
  several rows (one per role); the folder gets every mapped group at its level.
  Assignment is idempotent — `addroleassignment` merges, so re-adding an existing
  binding is a no-op and manual extra grants survive. A folder whose tier term has
  no group-map rows is logged as a warning (locked admin-only until rows are added).
  (Superseded the earlier "scratched for now" decision after the 403-on-locked-folder
  finding made per-folder manual assignment clearly untenable at scale.)
- **Both libraries in one run.** A single Run provisions Staging (create + lock +
  map) and Documents (create + lock, no map).
- **Lives in Folder Manager** as a tab; the standalone Reconciliation web part is
  removed from the bundle/componentId; the copy button is deleted.

## Decision (2026-07-20b) — pre-create the Year × Document Type grid

Under every **leaf (unit)** folder — the actual upload targets — reconciliation
also pre-creates the full **Year × Document Type** grid: each Year folder, and
inside it each Document Type folder. Both come from Term Store term sets
(`yearPeriod` = `f7c578a1-…`, `documentType` = `0540e66e-…`), so they are finite
and walkable just like the segment/dept/unit terms.

- Path order matches the upload form: **Unit / Year / Document Type** (Form.tsx
  creates the Year folder first, then the Document Type folder inside it).
- These grid folders **inherit the unit's ACL** — no `breakroleinheritance`, no
  folder-map row. They are organizational, not security boundaries.
- Created in **both** libraries (Staging + Documents).
- Idempotent via `ensureFolder` (create-or-resolve); a re-run after a new Year term
  is added simply fills in the new folders. Logged as a per-unit summary count,
  not one line per folder, to keep the log readable.
- Cost: units × years × doc types × 2 libraries folder ensures — acceptable for a
  one-time (occasionally re-run) admin action, but the slowest part of a run.

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
