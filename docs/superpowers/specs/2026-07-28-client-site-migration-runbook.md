# DMS — Client-Site Migration Runbook

**Date:** 2026-07-28
**From:** `dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting`
**To:** `‹client-tenant›/sites/‹client-site›`
**Companion:** `2026-07-28-dms-uat-test-plan.md` (the acceptance gate)

> **Supersedes `2026-07-16-tenant-seed-data-runbook.md`**, which is materially stale: it still asks
> for Graph API permissions (Graph fully retired), still says `GroupId` = Entra Object ID (it is now
> the SP group **integer** id), still names the columns `TermGuid` and `Side` (actual: `UnitTermGuid`
> and `Category`), and still lists the 5 retired term sets instead of the 4 Head Offices.
> Use this document instead.

---

## The core migration risk, stated once

**Nothing in the code is site-agnostic by default — it is site-agnostic only if `DMS Config` is
complete.** Every `termSet_*`, `col_*`, and mode row falls back to a hard-coded
ClarenceDMSTesting value when absent:

| Fallback | File |
|---|---|
| `DEFAULT_MODES`, `DEFAULT_SETTINGS` | `src/webparts/form/components/Form.tsx` |
| `DEFAULT_MODES` | `src/webparts/bulkUpload/components/BulkUpload.tsx` |
| `RECON_MODES` | `src/webparts/folderManager/components/FolderManager.tsx` |

On the client site those GUIDs point at a **different site collection**, so the term-store call
returns `404 / apiNotFound` and the affected dropdown renders **empty with no error message**.
There is no loud failure. Treat "config complete" as a hard gate, not a nice-to-have — this is
test **T2.2a/T2.2b** in the test plan.

---

## Phase 0 — Pre-flight code work (do before packaging)

| # | Item | Action |
|---|---|---|
| **M-04** | **`uploadCommand` will not deploy.** Its id `c2fd5b59-02ed-4892-87f2-2418fefb4980` is **absent** from the `componentIds` array in `config/package-solution.json` (which lists 6 ids: Form, HideAppBar, ApprovalDocument, FolderManager, Onboarding, BulkUpload). | Add the id, or consciously drop the command set from scope |
| **M-04b** | **No `elements.xml` anywhere in the repo.** Neither the `hideAppBar` application customizer nor the `uploadCommand` command set has any custom-action provisioning. With `skipFeatureDeployment: true` they rely on tenant-wide deployment or manual custom-action creation. | Decide: add `elements.xml` + set `skipFeatureDeployment: false`, or document the manual custom-action step |
| **M-05** | `config/serve.json` still points 3 of 5 serve configurations at the retired `/sites/SPFX-Sandbox-Testing-Ground` (`approvalDocument`, `hideAppBar`, `uploadCommand`). | Repoint to the client site (dev-only, but it will waste your time otherwise) |
| **M-07** | Version is `1.0.47.0`. | Bump to a clean client baseline (suggest `2.0.0.0`) so the client's app catalog history starts clean |
| **M-10** | Working tree has 12 modified files and ~10 untracked (`dev/`, 4 loose `Group *.svg` at repo root, `src/DMS Folder Map.csv`, drawio backups). | Commit or clean. Do **not** package from a dirty tree — and note `src/DMS Folder Map.csv` sits inside `src/` |
| M-11 | 14 lint warnings, incl. `max-lines` on `BulkUpload.tsx` (2430 > 2000). | Optional. Non-blocking, but agree it's accepted debt before handover |
| M-12 | Confirm `webApiPermissionRequests: []` stays empty. | The client must approve **no** API permissions — a genuine selling point of the Graph retirement. Verify with test O-03 |

### Hard-coded titles — the client site has no naming freedom here

These are **not** config-driven. The client site must use these exact list/library titles:

| Title | Where hard-coded |
|---|---|
| `Staging` | `ApprovalDocument.tsx` lines 124, 137, 176 — **ignores the `stagingLibrary` config row** |
| `Documents` | `ApprovalDocument.tsx` lines 199, 229; `BulkUpload.tsx` line 46 (`DOCUMENTS_LIST_TITLE`) |
| `DMS Config` | `Form.tsx` 399/455, `BulkUpload.tsx` 510/566, `FolderManager.tsx` 875/895/922, `FolderMap.tsx` 84, `GroupMapBuilder.tsx` 38 |
| `DMS Group Map` | `GroupMapBuilder.tsx` 37 |
| `DMS Folder Map` | `dmsFolderMap.ts` 8 |

If the client insists on different names, that is a code change — cost it before promising.

---

## Phase 1 — Client-site provisioning (in order; each step feeds the next)

### 1. Site + app catalog
1. Create/identify the client site collection. Record its URL.
2. Decide **tenant** app catalog vs **site-collection** app catalog. Site-collection scope is
   lower-friction and matches how this was tested; tenant scope is required if you want the
   extensions auto-provisioned tenant-wide (interacts with M-04b).
3. Upload the `.sppkg`, deploy, add the app to the site.
4. Confirm no pending API-permission approvals appear (M-12).

### 2. Term store — rebuild in-site (**GUIDs will all be new**)
Term sets are **per site collection**. The ClarenceDMSTesting GUIDs in CLAUDE.md are dead on
arrival at the client site.

1. As Site Collection Admin, open the site's `termstoremanager.aspx`. No Admin Center access needed
   — this is the whole point of the in-site pivot (`2026-07-26-in-site-term-store-management-design.md`).
2. Create a site-collection-local group (e.g. `DMS`).
3. Import the 5 CSVs from [docs/term-store-import/](docs/term-store-import/):
   `01-document-type`, `02-year`, `03-confidentiality`, `04-vendor`, `05-group-head-office`.
4. Create the 3 other Head Office sets (Upstream Malaysia, Minamas, **NBPOL** — not "NBBOL").
   Confirm with the client whether they supply distinct Department/Unit trees; today they share
   GHO's terms via `isAvailableForTagging`.
5. **Capture all 8 new GUIDs**:

   | Set | New GUID |
   |---|---|
   | documentType | `‹…›` |
   | yearPeriod | `‹…›` |
   | confidentiality | `‹…›` |
   | vendor | `‹…›` |
   | Group Head Office | `‹…›` |
   | Upstream Malaysia Head Office | `‹…›` |
   | Minamas Head Office | `‹…›` |
   | NBPOL Head Office | `‹…›` |

6. Verify each: `GET ‹site›/_api/v2.1/termStore/sets/{guid}/terms` → 200 + non-empty `value`.
   Note the working read path is `/terms`; `/children` on a **set** returned `apiNotFound` during
   testing — use `/terms` for sets, `/terms/{id}/children` for descent.
7. **Ampersands must be fullwidth ＆ (U+FF06)** — SharePoint requirement, and plain `&` breaks
   folder-name matching (test T2.3).

### 3. Libraries + columns
1. Create `Staging` and `Documents` (exact titles — see hard-coded table).
2. Provision columns with [docs/column-provisioning/provision-dms-columns.js](docs/column-provisioning/provision-dms-columns.js)
   (browser-console script — preferred over `scripts/Add-DocumentsMetadataColumns.ps1`, since
   PowerShell reliably fails in this environment).
3. Add the Documents metadata-parity columns (`2026-07-24-documents-metadata-parity-design.md`).
4. **Read the internal names back** — they are frozen at creation and unpredictable:
   `GET ‹site›/_api/web/lists/getbytitle('Staging')/fields?$select=Title,InternalName&$top=500`.
   Avoid `/` and `&` in display names at creation time (they force `_x002f_` / `_x0026_`).
5. Record every internal name; they become the `col_*` rows and `labelCol`/`tidCol` values in step 5.
6. Staging settings: content approval **ON**, **Draft Item Security = "Any user who can edit items"**.
   Any other value gives Contribute uploaders a misleading 403 (test F-17).
7. **`Full Name` column — on BOTH libraries.** Single line of text, display name exactly
   `Full Name`. Folder names are abbreviations (`GMB_STRATCOMMS`); this column carries the term's
   real label so the details pane can explain the code. **Read the internal name back** like every
   other column — it is usually `Full_x0020_Name`, but the resolver matches however the space was
   typed, so do not hand-write it anywhere.
   Then **remove it from the default view, not from the form**: the details pane renders form
   fields, so hiding it on the form deletes it from the pane, while merely dropping it from the
   view still shows it.
8. **`DMS Folder` content type — on BOTH libraries.** Create a content type named exactly
   `DMS Folder`, parent **Folder**, add the `Full Name` column to it, and add it to `Staging` and
   `Documents`. Reconciliation stamps it on every folder it creates. Without it the folders keep
   the built-in `Folder` type, `Full Name` never reaches the details pane, and the run logs a
   warning rather than failing — so this is easy to miss until a user asks what `GMB_STRATCOMMS`
   means.

### 4. DMS lists
Create `DMS Config`, `DMS Group Map`, `DMS Folder Map`, `DMS Term Abbreviation` (exact titles).

- **`DMS Config`** columns: `Title`, `ConfigType` (`mode`|`setting`), `ModeLabel`, `Category`
  (`BusinessSegment`|`Project`), `TermSetGuid`, `StagingFolder`, `SortOrder` (Number),
  `Levels` (multi-line plain text), `SettingValue`.
- **`DMS Group Map`** columns: `GroupId` (SP group **integer** id, as text), `GroupName`,
  `Segment` (term-set GUID), `UnitTermGuid`, `Role` (`MEMBER`|`UPL`|`APR`|`GLOBAL`).
- **`DMS Folder Map`**: TermGuid → FolderUniqueId, written by reconciliation. Leave empty.
- **`DMS Term Abbreviation`** columns: `Title` (Text — the term's full label), `TermGuid` (Text),
  `Abbreviation` (Text — the folder-name segment), `Level` (Choice: `Segment`|`Department`|`Unit`).
  **Required, not optional**: folder names come from this list, and a term with no row here is
  skipped by every reconciliation run, so that unit can never upload. Seed it in step 7 — the
  GUIDs are per-site, so it cannot be populated before the term store exists.
  > `Title` is not decoration. It is the only copy of the term's label outside the term store, and
  > it is what lets reconciliation repair a term that was deleted and re-added. Do not blank it.

### 5. Populate DMS Config — **the gate**

**16 `setting` rows** (`ConfigType = setting`, `Title` / `SettingValue`). Write **all of them**
explicitly, even where the value matches a code default — the point is to never depend on a
fallback:

```
termSet_documentType      ‹new documentType guid›
termSet_yearPeriod        ‹new yearPeriod guid›
termSet_confidentiality   ‹new confidentiality guid›
termSet_vendor            ‹new vendor guid›
col_documentType          Document_x0020_Type
col_yearPeriod            Year
col_documentDate          DocumentDate
col_confidentiality       Confidentiality_x0020_Level
col_vendor                Vendor_x002f_CustomerName
col_remark                Remark
col_legallyPrivileged     LegallyPrivileged
col_businessSegment       Business_x0020_Segment
col_businessSegmentTid    BusinessSegmentTid
legallyPrivilegedFor      ‹term guid of the ONE confidentiality level that offers the tick›
stagingLibrary            Staging
allowedExtensions         (leave SettingValue EMPTY — see AllowedFileTypes below)
```

> **`legallyPrivilegedFor`** names the single confidentiality term below which the *Legally
> Privileged* checkbox appears on the upload form. **Leave it blank and the tick is never
> offered** — that is the default, and it is silent, so a site that forgets this row simply never
> collects the flag and nobody notices. The value is re-derived at upload time, so a user who
> ticks the box and then changes the level cannot stamp `true` on a level that does not offer it.
>
> `col_remark` and `col_legallyPrivileged` back the two fields added in 1.0.6x. Both the Form and
> Bulk Upload write them on every upload.

(the `col_*` values above are ClarenceDMSTesting's frozen names — replace with whatever step 3.4
actually returned on the client site)

> ⚠ **REQUIRED: the `AllowedFileTypes` column.** Since 2026-07-30 file types are read from a
> multi-select Choice column on `DMS Config`, **not** from `SettingValue`. Create it before the
> first upload test:
>
> | Property | Value |
> |---|---|
> | Internal name | `AllowedFileTypes` |
> | Type | Choice, **allow multiple selections** |
> | Fill-in choices | **Disabled** (enabling it restores the typo risk it exists to remove) |
> | Choices | `.pdf` `.doc` `.docx` `.xls` `.xlsx` |
>
> Then tick all five on the `allowedExtensions` row. There is **no fallback to
> `SettingValue`** — a site without this column runs on the hardcoded `FALLBACK_FILE_TYPES`
> in `src/shared/allowedFileTypes.ts` and shows admins a warning. Verify with:
>
> ```
> /_api/web/lists/getbytitle('DMS%20Config')/fields?$select=InternalName,FillInChoice,Choices&$filter=InternalName%20eq%20'AllowedFileTypes'
> ```
>
> Expect `"FillInChoice": false` and exactly those five choices, each with a leading dot.
> Config is read once on mount, so **hard-refresh** before testing any change.

**4 `mode` rows** (`ConfigType = mode`), one per Head Office. All pilot segments share
`[Department, Unit]`, so all four take the same `Levels`:

```json
[{"label":"Department","column":"Department","labelCol":"Department","tidCol":"DepartmentTid"},
 {"label":"Unit","column":"Unit","labelCol":"Unit","tidCol":"UnitTid"}]
```

| Title | ModeLabel | Category | TermSetGuid | StagingFolder | SortOrder |
|---|---|---|---|---|---|
| gho | Group Head Office | BusinessSegment | ‹gho› | Group Head Office | 1 |
| upstream-my-ho | Upstream Malaysia Head Office | BusinessSegment | ‹umy› | Upstream Malaysia Head Office | 2 |
| minamas-ho | Minamas Head Office | BusinessSegment | ‹min› | Minamas Head Office | 3 |
| nbpol-ho | NBPOL Head Office | BusinessSegment | ‹nbp› | NBPOL Head Office | 4 |

> `parseReconModes` **drops** any mode row with an empty or old-schema `Levels`, and trims
> whitespace — a trailing space in a GUID silently breaks matching. Paste carefully.

### 6. Security groups + site entry
1. Create the site-entry Read group (`DMS_SITE_MEMBERS`). **Required** — folder-group Limited Access
   only reaches an assigned folder via direct link; the site Home/root still denies
   (memory `dms-two-layer-access-site-plus-folder`).
2. **Break `Staging` inheritance at *library* level** so site-Read cannot leak in.
3. Create the 3 groups per unit — base (Documents view), `_UPL` (Staging Contribute),
   `_APR` (Staging Design). Base is **never** on Staging; that is the isolation mechanism.
   GHO alone is ~36 groups.
   Use the **Group Map Builder** web part (one-shot create + members + role-from-suffix), not
   `scripts/Create-DmsGroups.ps1`.
4. Confirm `GroupId` rows land as SP integer ids (test GM-03).
5. Optional: one `GLOBAL` row for read-only super-viewers (no upload — `collectMembership` skips it).

### 7. Folders + reconciliation

**Seed `DMS Term Abbreviation` FIRST — before the first reconciliation run.** Folder names come
from it, and a term with no row is skipped, so running before seeding creates nothing and produces
one `SKIPPED (no abbreviation)` line per term.

1. Set the **segment** codes: `StagingFolder` on each `mode` row → `GHO`, `MHO`, `NBPOLHO`
   (these are NOT in the abbreviation list). Note this makes `StagingFolder` differ from
   `ModeLabel` — deliberately; the label is what the form shows.
2. Seed the department and unit rows with `scripts/seed-term-abbreviations.js` (paste
   `docs/term-store-import/10-per-level-abbreviations.csv` into `CSV_TEXT`, run in DevTools).
   The CSV is keyed by label **path**, not GUID, so the script walks the live term store to
   resolve GUIDs — which is why it must run after step 2. Expect `wrote 175 rows` and **no
   unresolved label paths**. Any unresolved path means the CSV label does not match the live term
   exactly; fix the term or the term store, **not** the CSV.
3. Use **Folder Manager** to create `Staging/‹SEG›/‹DEPT›/‹UNIT›/` — abbreviated at every level.
   Year + Document Type subfolders are ensure-created on first upload and inherit the Unit ACL.
4. Run reconciliation **dry-run first**, review, then live.
5. Re-run immediately to confirm idempotency (test FM-04).
6. Confirm `DMS Folder Map` is populated. **Uploads cannot succeed before this** — the form
   resolves the target by `FolderUniqueId`.
7. Read the log for these, in the order they hurt:
   - `COLLISION` — two siblings share a code. **Nothing was created.** Fix and re-run; do not
     work around it, this is one ACL over two units' documents.
   - `SKIPPED (no abbreviation)` — that unit cannot upload at all until a row exists.
   - `NO TERM` — a folder no live term claims. Never deleted; review by hand.
8. Confirm `Full Name` is populated on a folder and visible in the details pane (step 3.7-3.8).
   If the grid shows the column, remove it from the **view**, not the form.

> Tell the client, in writing, before handover: **rename terms, never delete and re-add.** A
> rename keeps the term's GUID; a delete-and-re-add issues a new one and orphans the abbreviation,
> the folder mapping and the group permissions at the same moment. Reconciliation repairs that
> only when exactly one dead row and one new term share a name at the same level — and `Tax`,
> `Legal` and `PM` each exist under several parents, so it often cannot. See
> `docs/client/folder-abbreviations-guide.md`.

### 8. Auto-route flow
Rebuild in the client tenant — connections do not migrate.
- Routes by **live folder path** (rename-proof), splits on `Staging/` with **no leading slash**,
  strips trailing slashes (Copy file rejects them).
- "Create new folder" branch for missing targets.
- Sender: the client's equivalent of the `DMS No Reply` shared mailbox — **needs to be created and
  licensed on the client side. Raise this early; it has a lead time you do not control.**

### 9. Pages
Add web parts: Form (upload page), ApprovalDocument (approval page), BulkUpload + FolderManager +
Onboarding (admin-only pages). Restrict the admin pages.

**Publish every page.** An unpublished page is invisible to anyone without edit rights and denies
with the same screen a permission failure produces. Publishing state is not a permission — never
leave a page as a draft to keep it admin-only.

---

## Phase 1b — Access layers, verified on the test site 2026-07-29

Every one of these cost time on `/sites/ClarenceDMSTesting`. They are ordered the way a user hits
them, and each fails with a *different* error — knowing which layer you are on is most of the
diagnosis.

### The three grants a plain uploader needs

| # | Grant | Scope | Failure if missing |
|---|---|---|---|
| 1 | `DMS_SITE_MEMBERS` → **Read** | Site (`user.aspx`) | "You need access to **this site**" |
| 2 | `DMS_SITE_MEMBERS` → **Read** | **Site Pages** library | "You need access to **this list**" |
| 3 | `DMS_UPL` → **Read** | `Upload-Form.aspx` item | Page denies; the site itself works |

(2) is the one that gets missed. Site Pages carries **unique permissions**, so the site-level grant
does not reach it — the user is admitted to the site and refused by the page that *is* the site.
Sweep every list the web parts read and confirm `DMS_SITE_MEMBERS` has Read on each one that has
unique permissions: **Site Pages, Staging, DMS Config, DMS Mapping, DMS Folder Map**.

`Form.tsx` wraps those reads in `.catch(() => [])`, so a 403 on `DMS Group Map` or `DMS Config` is
indistinguishable from "this user has no rows" — it renders "your account isn't fully provisioned
to upload". **A provisioning message is not proof of a provisioning problem.** Have the user open
the list directly before touching Group Map data.

### Content approval hides FOLDERS — set Draft Item Security correctly

Staging: Content Approval = **Yes**, Draft Item Security = **"Any user who can read items"**.

Not "only users who can edit". Reconciliation-created **folders** are themselves Pending, and
nothing in the DMS approves them. With "edit items", a user holding only the Read-browse grant on
the Segment and Department folders — exactly what reconciliation gives unit groups — cannot **see**
those folders. Symptom: **Staging opens but is completely empty, while a direct deep link into the
unit folder works.** Every permission layer checks out, which sends you into the ACL for hours.

Folder approval status is invisible while the Approval Status column carries the DMS column
formatting (`display:none` when `[$File_x0020_Type] == ''`). Remove the formatting before reading
folder statuses. See memory `dms-content-approval-blocks-uploader`.

### One group per unit — do NOT add per-tier Group Map rows

One row per group, at the **leaf**. `isLeafChainValid` only checks that the term-store ancestry
resolved and ends at that leaf. The old `isChainAuthorized` demanded membership at every tier and
refused correctly provisioned uploaders; removed 2026-07-29, spec
`2026-07-29-leaf-only-upload-authorization-design.md`. If someone "fixes" a provisioning error by
adding MEMBER rows for the segment and department tiers, they are reintroducing hundreds of
derivable rows.

### Membership added natively skips site entry

Adding a user through People and Groups does **not** add them to `DMS_SITE_MEMBERS`. The **Folder &
Group Manager** does (`GroupMapBuilder` auto-adds), and reconciliation's self-heal sweeps every
`DMS_*` group's members in on each run. **Tell the client to add members via the web part**, or
every native addition produces a half-provisioned user.

### Site entry must come from exactly one group

If any other `DMS_*` group also holds Read at site level, site entry has two independent doors and
removing someone from `DMS_SITE_MEMBERS` will not revoke it. Keep site-level Read on
`DMS_SITE_MEMBERS` alone; page-gate groups such as `DMS_UPL` get **no** site permission.

### Revoking access is not a Group Map operation

Deleting a `DMS Group Map` row stops reconciliation re-assigning the group and stops the form
offering a path — it does **not** remove the SP group, its members, or the existing role assignment
on the folder. The user keeps full upload access. Reconciliation only ever *adds* role assignments;
there is no revoke anywhere in the run. **Removing the user from the SP group is the only real
revocation.**

### Items with unique permissions bypass library grants

Check the **People** tab and **Links** tab under Manage access on each page. Pages shared
individually carry their own grants (a stray "Edit link for internal users" gave 7 people
Contribute on `Upload-Form.aspx`). Fix: item → Manage access → ⋯ → Advanced settings → **Delete
unique permissions**, then remove the link separately — deleting unique permissions revokes what a
link granted but can leave the link object alive.

### Stale tokens

SharePoint resolves group membership at sign-in. After any permission change a refresh reuses the
cached claims — the user must **fully sign out** (`login.microsoftonline.com/logout`) or open a new
private window. Budget for this in every acceptance test or you will chase phantom failures.

---

## Phase 2 — Acceptance

Run the full test plan against the client site. §2 fixtures first — if T2.2a fails, stop and fix
config; everything downstream is invalid. Sign-off = the test plan's §5 exit criteria.

---

## Rollback

The client site is new, so rollback is retreat rather than restore:
1. Turn the auto-route flow **off first** — otherwise approvals keep copying into `Documents`
   mid-rollback.
2. Retract/remove the app from the site (removes web parts; **leaves lists, libraries, term sets,
   groups, and folders intact**).
3. Term sets, DMS lists and folder ACLs persist and can be reused on retry. There is nothing to
   restore from ClarenceDMSTesting — it is a separate site collection and no data migrates.

**No document content migrates.** ClarenceDMSTesting is a test site; the client site starts empty.
Confirm this in writing with the client — if they expect existing documents carried over, that is
unscoped work.

---

## Open questions to resolve before starting

1. Tenant or site-collection app catalog? (drives M-04b)
2. Are the extensions in scope? If yes, M-04 + M-04b are blockers, not warnings.
3. Do Upstream Malaysia / Minamas / NBPOL get their own Department/Unit trees, or keep sharing GHO's?
4. Who creates the client's `DMS No Reply` equivalent mailbox, and when?
5. Will Clarence hold Site Collection Admin on the client site? Required for the in-site term store,
   reconciliation, and breaking inheritance.
6. Any expectation of migrating existing documents? (assumed **no**)
