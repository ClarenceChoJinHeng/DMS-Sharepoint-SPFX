# SDG DMS — Backlog

> Update this file as work is completed or new tasks are added.
> Last updated: 2026-07-06.

---

## DONE

- [x] SPFx project scaffolded (Heft toolchain, Node 22, React 17, TS ~5.8)
- [x] Form.tsx: File picker (single file only)
- [x] Form.tsx: File type restriction — allowed list enforced before upload
- [x] Form.tsx: Document Name free-text rename input (extension preserved, illegal chars stripped)
- [x] Form.tsx: All 7 metadata dropdowns loaded from Term Store via direct REST
  - Document Type, Department, Year/Period, Confidentiality Level, Vendor (flat term sets)
  - Project Name (cascade — loads child terms of selected Department on change)
- [x] Form.tsx: Document Date field — input type="date" converted to M/D/YYYY for US site locale
- [x] Form.tsx: Details textarea wired to _ExtendedDescription (built-in Note field)
- [x] Form.tsx: File upload via /Files/Add + metadata tagging via validateUpdateListItem
- [x] Form.tsx: validateUpdateListItem HasException check (HTTP 200 can still carry field errors)
- [x] Form.tsx: Resilient term set loading — per-set try/catch so one 404 cannot blank all dropdowns
- [x] Term Store GUIDs verified against live tenant
- [x] Column internal names verified against live /fields API
  - DocumentType frozen as Department_x0020_Type (created as "Department Type", renamed)
  - DocumentDate has no _x0020_ encoding (created without space)
- [x] Project Name column re-bound to Department term set (eaba82e5) in Staging library settings
- [x] Approver SharePoint group created with Design permission level
- [x] Staging library: Content approval enabled (Versioning settings > Require content approval = Yes)
- [x] ApprovalDocument SPFx web part built (`ApprovalDocument.tsx`)
  - Loads item metadata (FieldValuesAsText) and renders document preview via iframe
  - Shows Submitted by, all metadata fields, PDF/Office preview
  - Pre-selects current approval status from `OData__ModerationStatus` on load (0=Approved, 1=Rejected, 2=Pending)
  - Approve: uses `/approve()` file REST endpoint
  - Reject: uses MERGE on list item with `OData__ModerationStatus: 1` (NOT `/reject()` — it fails on already-Approved files)
  - Pending: uses MERGE on list item with `OData__ModerationStatus: 2`
  - Reject + Pending: search-then-delete from Documents library (queries by FileLeafRef, deletes by real FileRef)
  - Comment passed as `OData__ModerationComments` on all three actions
- [x] Auto-route PA flow: Condition (Approved branch) working — copies file to dept library + sends approval email
- [x] Auto-route PA flow: Condition 1 (Rejected/Denied branch) fixed
  - PA connector returns `"Denied"` (NOT `"Rejected"`) for `{ModerationStatus}` on rejected items
  - Condition uses OR: `{ModerationStatus}` equals `"Denied"` OR equals `1`
  - Reject email: comment field uses `body('Get_item')?['{ModerationComment}']` — singular, no 's', with curly braces
- [x] Auto-route PA flow: copy destination is now **path-based, not metadata-based** (2026-06-26)
  - Derives destination from the file's live folder path (`{Path}`) → rename-proof, no stale-metadata failures
  - Added `Create new folder` (idempotent) before Copy file → fixes "destination doesn't exist" + brand-new team first-approval
  - Gotchas: split on `Staging/` (no leading slash); Copy file rejects trailing slash; Documents real URL = `Shared Documents`
  - See requirements.md §8 for the full action chain
- [x] Approval/rejection emails sent from shared mailbox `dms-noreply@trinergydigital.com` (DMS No Reply)
  - Swapped to "Send an email from a shared mailbox (V2)"; granted Send as delegation to `clarence@trinergydigital.com`
- [x] Rejection email **To** field wrapped in `trim(outputs('Get_item')?['body/Author/Email'])`
  - Fixed `string/email` validation failure caused by a trailing newline (`clarence@trinergydigital.com\n`)
- [x] Folder Manager web part: bulk-rename tool that **mirrors each rename to BOTH Staging + Documents** (no toggle)
  - Per-rename log indicator shows which libraries were touched (`Staging ✓, Documents ✓` / `Documents — not present yet` / `FAILED`)
  - Documents skipped gracefully when the folder doesn't exist there yet (only created on first approval)
  - Resolves library roots via `getbytitle/RootFolder` so the Staging → `Shared Documents` mapping is correct
  - Collapsible Departments/Projects sections + per-folder subfolder expand
- [x] ApprovalDocument: "Back to document list" + Cancel now return to the **file's own folder** (`AllItems.aspx?id=<folder>`) instead of the Staging root
- [x] FieldValuesAsText OData double-encoding: keys like `Department_x0020_Type` come back as `Department_x005f_x0020_x005f_Type` — `pick()` helper reads double-encoded key first then falls back

---

## IN PROGRESS

### Nested Subcategory Folder Chain (CLIENT CHANGE — 2026-07-06)
Folder hierarchy is a **nested chain** (Sub-1 contains Sub-2 contains Sub-3, etc.), NOT flat siblings — confirmed by client. See requirements.md §11a-note for full detail.

- [x] **Term Store (Account D):** Nested chain created — `Account D > AI Engineer Sub-1 > Sub-2 > Sub-3 > Sub-4 > Sub-5 > Sub-6`, all within the Department term set (eaba82e5).
- [x] **Staging library column:** `Department_x0020_Team` verified via REST (`InternalName`/`StaticName` both confirmed) — no column change needed, existing column reused.
- [x] **Upload Form (`Form.tsx`):** Cascading nested dropdown picker implemented —
  - One dropdown per depth level; level 0 required, deeper levels optional with a "-- Use this folder --" stop option
  - Each level's children are fetched live via Term Store REST when a term is picked
  - Deepest pick is tagged to `Department_x0020_Team` (or `Project_x0020_Name` in Project mode) as `Label|GUID`
  - Upload path is built by joining every picked level as real nested folders (e.g. `.../Account D/Sub-1/Sub-2/Sub-3/`)
  - New state: `subTeamLevels`, `subTeamPath`; new handler `handleSubTeamLevelChange`; new loader `loadLevelChildren`
  - Build verified clean (`npm run build`, no new errors/warnings)
- [x] **Folders + groups (Account D):** Created manually in both Staging and Documents at every nested level, with permissions assigned.
- [ ] **Folder Manager:** Still assumes 1-level subfolder per department — needs nested create/rename support to match the new chain structure (not yet started).
- [ ] **Onboarding web part:** Subcategory chain creation not supported — folders were created manually this round; needs a proper creation flow for future departments.
- [ ] **Auto-route PA flow:** Needs an end-to-end test now that the destination path is deeper (should be transparent since it derives from the live `{Path}`, but unverified with the extra nesting).
- [ ] **Other departments:** Only Account D has the nested chain today. Repeat term store + folder + group setup for Finance/HR/IT/etc. when needed.
- [ ] **CLAUDE.md:** Update term set section with the nested-chain note once Folder Manager/Onboarding support lands.

---

## PENDING (not started)

### Approval Pipeline
- [ ] Native SP approve/reject buttons: set up a separate PA flow "SDG - Remove from Documents on Rejection" to delete from Documents when approver uses the native library UI (not the web part)
- [ ] Test full pipeline end to end: upload > approve via web part > auto-route fires > file in dept library

### RBAC Groups
- [ ] Create Uploader SP group (Contribute permission on Staging library)
- [ ] Create Reader SP group (Read permission on department destination libraries only)
- [ ] Break permission inheritance on department libraries so Readers see only their department

### Approver Experience
- [x] ApprovalDocument web part built (see DONE above)
- [ ] Restore native hover underline on Staging Name column (PARKED)
  - Current: Name column is custom-formatted so clicking a FILE opens ApprovalDocument. Side effect: folder name has no hover underline, because the folder uses `customRowAction: defaultClick` (no `href`) to keep single-click nav + grid-view rename working, and SP's hover-underline rule only targets anchors WITH an href.
  - Trade-off is unavoidable while the approval link lives in the Name column.
  - Fix when revisited: move the approval trigger to a SEPARATE "Review" column (single line of text, blank), formatted as a button linking to `SitePages/ApprovalDocument.aspx?itemId=[$ID]` shown only when `File_x0020_Type != ''`. Then strip ALL formatting from the Name column → full native behavior returns (single-click folder nav, grid rename, AND hover underline on both files and folders).
  - Cost: approvers click a "Review" button in its own column instead of the file name itself.

### Notifications & Escalation
- [x] Email notification to uploader on approval or rejection (done — auto-route PA flow, Author/Email)
- [ ] Reminder email to Approver after N days of no action on pending item
- [ ] Optional escalation to secondary approver

### Duplicate Detection
- [ ] Pre-upload REST check: does a file with this name already exist in Staging?
- [ ] Prevent overwrite; allow versioning only

### Audit Logging
- [ ] Configure SharePoint / M365 Compliance audit logging
- [ ] Track: uploads, approvals, edits, deletions, access events

### Retention Policy
- [ ] Define retention durations per document type
- [ ] Apply M365 Purview retention labels
- [ ] Prevent unauthorized deletion where required

### Search & Retrieval
- [ ] Pre-configured library views (by Department, Approval Status, Year, Document Type)
- [ ] SharePoint search across document content (PDF/Office)

### Naming Convention (deferred)
- [ ] Auto-generate filename as [Department]_[DocType]_[Project]_[Year].pdf
  - Currently: free-text rename only; user types any name

### Admin Onboarding Automation
- [x] **Folder Onboarding SPFx web part built** (`onboarding` web part, GUID `7b3e9c14-2a8d-4f61-9b0e-5d3c8a1f4e72`) — 2026-06-26
  - UI: Target library toggle (Staging / Documents / Both) + Section toggle (Departments / Projects) + parent folder row + add-as-many subfolder rows. Each row = name + **live Graph group search** + per-row permission-level dropdown (Read / Contribute / Design).
  - On submit: (Step 7) validate every group via Graph `GET /groups/{id}`; (Step 8) create parent + subfolders via REST `folders/AddUsingPath`; (Step 9) `breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)` per folder; (Step 10) `ensureuser` the M365 group claim then `roleassignments/addroleassignment`.
  - Group search: Graph `GET /groups?$search="displayName:…"` (ConsistencyLevel: eventual, $count).
  - M365 group → SP principal: `ensureuser` with `c:0o.c|federateddirectoryclaimprovider|{groupId}`.
  - **Requires `Group.Read.All` approved in SharePoint Admin → API access** (Global Admin — Desmond). Added to package-solution.json webApiPermissionRequests.
- [ ] **NOT included — Term Store child-term creation.** The original plan's step (create the dept/project child term in the Term Store) is NOT in this web part. So a newly-onboarded department won't appear in the Form's metadata dropdowns until its term is added manually (or a future enhancement adds Graph/CSOM term creation).
- [ ] Pending: Group.Read.All admin approval, then end-to-end test (search → submit → verify folders + broken inheritance + group assignment in both libraries).

---

## NOTES

- Power Apps canvas app ("DMS Upload Form") is fully abandoned
- Power Automate instant flows (GetTermSetValues, DepartmentProjectList, UploadToStaging)
  exist in the tenant but are NOT used by the SPFx web part (they were built for Power Apps)
- Only the Auto-route automated flow is still relevant — it is server-side, fires after approval
