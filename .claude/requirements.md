# SDG DMS — Living Requirements

> Update this file whenever scope changes. Last updated: 2026-07-06.
> Sources: sdg-dms-scope + SDG_DMS_Handover_ForClaudeCode (2).pdf + session decisions.

---

## Modules from Scope (sdg-dms-scope)

### 1. Dedicated SharePoint Site
- Site provisioning for DMS
- Site-level security groups: Admin, Approver, Uploader, Reader
- Document libraries: Staging (upload area) + Final/department libraries
- Status: Site exists (sandbox). Approver group created. Uploader + Reader groups pending.

### 2. Metadata Columns on Libraries
Required fields captured at upload:
- Document Type (Term Store)
- Department / Business Unit (Term Store)
- Year / Period (Term Store)
- Document Date (Date field — sent as M/D/YYYY, US locale)
- Project Name (Term Store — cascade child of Department)
- Confidentiality Level (Term Store)
- Vendor (optional, Term Store)
- Details / Description (built-in _ExtendedDescription Note field)
- **[ADDED — see rename section]** Document Name — free-text rename input

Status: All columns exist on Staging. Web part captures all fields. DONE.

### 3. Document Name / Rename (ADDED REQUIREMENT)
User can type a custom name for the uploaded file before submitting:
- Extension is always preserved from the original file (user cannot accidentally strip it)
- Illegal filename chars are stripped automatically
- If left blank, original filename is used unchanged
- Preview shown: "Saved as: X.pdf"
- Input is disabled until a file is selected

Status: DONE in Form.tsx (`buildUploadName()`, `docName` state).

### 4. Controlled Vocabulary (Term Store)
- All dropdowns load from DMS Metadata term group via REST `_api/v2.1/termStore/sets/{id}/children`
- No hardcoded values, no SharePoint Choice columns
- Project Name is a cascade: selecting a Department loads that dept's child terms as Project options
- Resilient loading: each term set wrapped in its own try/catch — one 404 can't blank all dropdowns

Status: DONE.

### 5. Versioning & File Type Restrictions
Allowed extensions: `.pdf .doc .docx .xls .xlsx .ppt .pptx .txt .csv .jpg .jpeg .png`
All others rejected before upload. Library-level versioning config TBD.

Status: File type check DONE (`isAllowedFile()` in Form.tsx). Library versioning TBD.

### 6. Naming Convention Enforcement
- Scope format: [Department]_[DocType]_[Project]_[Year].pdf
- Decision (2026-06-16): Free-text rename is sufficient for now. Auto-naming deferred.

Status: Free-text rename DONE. Auto-convention NOT implemented (future backlog).

### 7. Staging Upload Mechanism
- SPFx web part uploads ONE file to Staging library
- All metadata captured at upload time via validateUpdateListItem
- Uploaders only; no anonymous access

Status: DONE. Direct REST via spHttpClient.

### 8. Approval Workflow Automation
Process: Uploader submits > Approver reviews in Staging > Auto-route fires on approval

If Approved:
- Document copied to dept library under `Documents/Departments/{Dept}/{Project}/`
- Notification sent to uploader (approval email via PA flow)

If Rejected:
- Notification sent to uploader (rejection email via PA flow)
- File deleted from Documents library if previously approved and routed

If Pending:
- File deleted from Documents library (de-published from reader view)

Implementation approach: Native SharePoint content approval + `ApprovalDocument.tsx` SPFx web part + Auto-route PA flow.

**ApprovalDocument web part critical notes:**
- Reject uses `MERGE` (`OData__ModerationStatus: 1`) — NOT `/reject()` (fails on already-Approved files)
- Pending uses `MERGE` (`OData__ModerationStatus: 2`)
- Delete from Documents: search-then-delete by `FileLeafRef` (do NOT guess path — use real `FileRef` from query)
- Status pre-selection on load: maps `OData__ModerationStatus` integer → radio button (0=Approved, 1=Rejected, 2=Pending)
- Back/Cancel links return to the **file's own folder** in Staging (`AllItems.aspx?id=<folder>`, derived from `item.File.ServerRelativeUrl`), not the library root — saves the approver from drilling back in. Helper `backUrl()`.

**PA flow critical notes (Auto-route approved Pending folders to Departments):**
- Outer Condition: `{ModerationStatus}` equals `"Approved"` → True branch copies file + sends approval email
- Condition 1: `{ModerationStatus}` equals `"Denied"` OR equals `1` → True branch sends rejection email
  - PA connector returns `"Denied"` (NOT `"Rejected"`) for rejected items — this is non-obvious
  - Comment in rejection email uses `body('Get_item')?['{ModerationComment}']` — singular "Comment", no 's', with curly braces

**Auto-route copy destination — path-based, NOT metadata-based (redesigned 2026-06-26):**
- The Approved branch now derives the destination from the file's **live folder path** (`{Path}`), not from Department/Project metadata. Rename-proof: files move with their folder, so the path is always current. Metadata-built paths went stale on rename and failed with "destination location does not exist" (0x80070002).
- True-branch actions: `Get item` → `Compose` (= `{Path}`, e.g. `Staging/Departments/Account D/AI Engineer D/`) → `Compose 1` (strip `Staging/` + trailing slash: `substring(split(outputs('Compose'),'Staging/')?[1],0,sub(length(split(outputs('Compose'),'Staging/')?[1]),1))`) → `Create new folder` (Library=Documents, Path=`outputs('Compose_1')`) → `Copy file` (Destination=`concat('/Shared Documents/', outputs('Compose_1'))`) → email.
- Gotchas: split on `Staging/` (NO leading slash — `{Path}` has none); Copy file rejects a **trailing slash** even when the folder exists; Documents real URL = `Shared Documents`.
- `Create new folder` is idempotent (no duplicate if exists) → also fixes first-approval failures for brand-new departments/projects.
- Approval/rejection email sender switched to shared mailbox `dms-noreply@trinergydigital.com` (DMS No Reply) via "Send an email from a shared mailbox (V2)"; needs **Send as** delegation for the flow's account.
- Rejection email **To** must be `trim(outputs('Get_item')?['body/Author/Email'])` — the raw Created By Email carries a trailing newline that fails the `string/email` check.

Status: DONE (web part + PA flow + path-based routing + noreply sender). Native SP approve/reject button → Documents delete flow NOT YET built.

### 9. Escalation & Reminder Logic
- Reminder email after N days if no approval action taken
- Optional escalation to secondary approver
- Email / Teams notifications

Status: NOT STARTED. Needs a Power Automate scheduled/reminder flow.

### 10. Duplicate & Conflict Detection
- Check for duplicate filename before upload
- System prevents overwrite; allows versioning only

Status: NOT STARTED. Needs pre-upload REST check against Staging.

### 11a-note. Folder Structure Change — Nested Subcategory Chain (CLIENT REQUEST 2026-07-06)

**Change:** Departments now have a **nested chain of subcategory folders** beneath them — NOT a flat sibling list. Confirmed by client 2026-07-06: *"the requirements is to make a Sub folder for each not siblings."*

Previous structure (2-level):
```
Staging/Departments/{Department}/{Project}/
Documents/Departments/{Department}/{Project}/
```

New structure (N-level nested chain — each subcategory nests inside the previous one):
```
Staging/Departments/{Department}/{Sub-1}/{Sub-2}/{Sub-3}/.../{Sub-N}/
Documents/Departments/{Department}/{Sub-1}/{Sub-2}/{Sub-3}/.../{Sub-N}/
```

Example (Account D, live in sandbox tenant):
```
Departments/Account D/AI Engineer Sub-1/AI Engineer Sub-2/AI Engineer Sub-3/AI Engineer Sub-4/AI Engineer Sub-5/AI Engineer Sub-6/
```

**Term Store structure:** Each `AI Engineer Sub-N` term is a **child of the previous Sub term**, not a sibling under Account D directly — i.e. `Account D > Sub-1 > Sub-2 > Sub-3 > Sub-4 > Sub-5 > Sub-6`, all within the Department term set (eaba82e5). Because it's a linear chain (no branching at each level, at least for Account D), each level currently has exactly one child term — but the code does not assume this; it supports branching if a future department needs multiple sub-options at any level.

**Upload Form UX (Form.tsx) — cascading dropdown picker, implemented 2026-07-06:**
- One dropdown per depth level. Level 0 is **required** (mirrors the old single sub-team dropdown requirement); every level below it is optional and includes a "-- Use this folder --" option to stop drilling deeper.
- Selecting a term at level *i* fetches its children live via Term Store REST and, if any exist, reveals level *i+1*. Selecting the blank option truncates the chain at that depth.
- The **deepest folder picked** (last entry in the path) is what gets written to the `Department_x0020_Team` / `Project_x0020_Name` metadata column — same single `Label|GUID` format as before.
- The **upload destination path** is built by joining every picked level's label as real nested folders, e.g. picking through Sub-1 → Sub-2 → Sub-3 uploads to `.../Account D/Sub-1/Sub-2/Sub-3/`.
- State: `subTeamLevels: TermOption[][]` (options per level) + `subTeamPath: TermOption[]` (what the user picked at each level). Handler: `handleSubTeamLevelChange(levelIndex, selectedId)`. Loader: `loadLevelChildren(mode, termId)`.

**Affected components:**
| Component | Status |
|---|---|
| Upload Form (`Form.tsx`) | **DONE** — cascading nested dropdown picker, arbitrary depth |
| Staging library `Department_x0020_Team` column | **DONE** — verified internal name matches, no change needed |
| Folder Manager | **NOT YET DONE** — still assumes 1-level subfolder per department; needs nested create/rename support |
| Auto-route PA flow | Likely unaffected — destination is derived from the file's live `{Path}`, which will simply be longer. Needs an end-to-end test once a file is routed through a nested folder. |
| Onboarding web part | **NOT YET DONE** — subcategory chain creation not supported; folders were created manually this round |
| RBAC / department detection | **Confirmed for Account D:** folders + groups were created manually with proper permissions at each nested level. No naming-convention change needed since department detection is unaffected (still keys off the top-level Department term). |

Status: **Account D — DONE end-to-end (term store nested chain + form cascading picker + manual folder/group setup).** Other departments and Folder Manager / Onboarding support: **PENDING**.

---

### 11. Role-Based Access Control (RBAC)
| Role     | SP Permission Level | Scope                      | Can do                          |
|----------|--------------------|-----------------------------|----------------------------------|
| Admin    | Full Control       | Site                        | Everything                      |
| Approver | Design             | Staging library             | Review and approve pending items |
| Uploader | Contribute         | Staging library             | Upload and tag; own pending only |
| Reader   | Read               | Their department library    | Read approved docs only          |

Department isolation: Readers get Read only on their department's destination library — NOT a filter, actual permission boundary.

Status: Approver group created (Design). Uploader + Reader groups NOT YET created.

### 11a. Department Detection — How It Works & Naming Convention

**Flow:**
1. User opens the Upload Form
2. Form calls Microsoft Graph API (`/me/memberOf`) → retrieves the user's M365 group memberships
3. Form calls Term Store API → retrieves all Department term labels (e.g. `Account-Team-1`, `Finance D`)
4. Code filters the user's groups for any group containing `"uploader"` in the name
5. For each uploader group, strips the `SDG-` prefix and `-Uploader` suffix to extract the department core (e.g. `SDG-Account-Team-1-Uploader` → core = `account-team-1`)
6. Cross-references that core against the Term Store labels using bare alphanumeric matching (hyphens, spaces, casing all ignored)
7. If a match is found → form shows that department and restricts uploads to that department's folders
8. If no match → error: "Your account is not assigned to a department group"

**Naming Convention (MUST be consistent across all three):**
| Thing | Example |
|---|---|
| Entra M365 Group | `SDG-Account-Team-1-Uploader` |
| Term Store Department label | `Account-Team-1` |
| Staging / Documents folder | `Account-Team-1` |

The department name between `SDG-` and `-Uploader` in the group name **must match** the Term Store label (hyphens, spaces, and casing are ignored — `Account-Team-1`, `Account Team 1`, and `accountteam1` all match each other).

**Known limitation:** The matching is based on naming convention, not an explicit mapping. If the convention is broken (e.g. group renamed in Entra without updating the Term Store label or folder), the detection fails silently and the user sees the "not assigned" error. Fix: keep all three names in sync when onboarding new departments.

Status: DONE (bare alphanumeric matching in `detectDepartment()`, Form.tsx).

### 12. Audit Logging + Access Matrix
Client ask (2026-07-30) is two things: (1) activity log — download, upload, approve, delete,
share; (2) access matrix — what a given user can reach at folder and file level.
These are different problems: (1) is events over time, (2) is current state. Purview answers
(1) and cannot answer (2).

Uploads/approvals/edits/deletions are deliverable in-app with no new access. Sharing is
**prevented** (owner-only sharing) rather than logged. **Downloads are Purview-only** — and
that needs **Audit Reader**, NOT SharePoint Admin (SP Admin and Site Collection Admin grant
zero audit access).

Status: DESIGNED, not started.
Design: `docs/superpowers/specs/2026-07-30-audit-log-and-access-matrix-design.md`

### 13. Retention Policy
Define retention durations, apply retention labels, prevent unauthorized deletion.
Status: NOT STARTED. M365 Purview / Compliance.

### 14. Search & Retrieval
- Search by title, content (PDF/Office files)
- Filter by metadata (Department, Year, Document Type, etc.)
- Pre-configured library views

Status: NOT STARTED. SharePoint search + custom views.

---

## Key Decisions That Diverge from Handover PDF

| Topic           | Handover PDF (original)                    | Current (actual)                                  |
|-----------------|--------------------------------------------|---------------------------------------------------|
| Form fields     | File name input + file picker + dept only  | ALL metadata fields                               |
| Project Name    | Separate term set (phantom GUID e1163337)  | Cascade child of Department term                  |
| Data layer      | Power Automate flows (V2 trigger)          | Direct SharePoint REST via spHttpClient           |
| Naming          | User types filename manually               | Free-text rename input (auto-convention deferred) |
| Files per submit| One                                        | One (unchanged)                                   |
