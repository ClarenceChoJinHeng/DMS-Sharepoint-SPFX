# SDG DMS — Claude Code Project Context

Auto-loaded every session. Keep this up to date whenever decisions change.
Full requirements: `.claude/requirements.md` | Backlog: `.claude/backlog.md`

---

## Project Identity
- **Client:** SD Guthrie (SDG) | **Agency:** Trinergy Digital | **Dev:** Clarence (Junior Digital Developer)
- **Type:** SPFx 1.23.0 React web part — Document Management System upload form
- **Project folder:** `C:\laragon\www\Work\Projects\sd-gatrie`
- **Tenant:** `dcidigitalcom.sharepoint.com` | **Site:** `/sites/ClarenceDMSTesting` (rebuilding here
  after being locked out of the old `/sites/SPFX-Sandbox-Testing-Ground`; serve.json still points at
  the old sandbox)
- **Staging library:** `Staging`
- **Single-site DMS** — the cross-site/multi-site model was retired 2026-07-26 (a new site draws from
  the same tenant storage quota → no space saved). Package is **Graph-free** (Share Guard retired).
  See memories `dms-single-site-decision`, specs `2026-07-26-cross-site-upload-retirement.md` +
  `2026-07-23-share-guard-retirement.md`.

## Dev Commands
```
nvm use 22          # must be Node 22
npm run start       # Heft dev server — NOT gulp serve (Gulp is NOT used)
npm run build
```
Workbench: `https://dcidigitalcom.sharepoint.com/_layouts/workbench.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true`

## Key Files
| File | Purpose |
|------|---------|
| `src/webparts/form/components/Form.tsx` | THE upload form — all logic lives here |
| `src/webparts/form/FormWebPart.ts` | Entry point — only imports Form.tsx |
| `src/webparts/form/components/IFormProps.ts` | Props interface (context only) |
| `Form.reference.tsx` (project root, outside src/) | Backup of original prototype — NOT compiled |

## Multi-Segment Model (current — replaces the old 2-mode Department/Project cascade)
The form is now **data-driven over "modes"** loaded from the `DMS Config` list. Each mode
has a `Category` (`BusinessSegment` | `Project` — the top-level family, NOT a tier; drives the
form's two top-level tabs and whether a Business Segment is required; internally still the
`side` property), a term-set GUID, and a `Levels` JSON chain
(variable depth, e.g. `[{"label":"Department","column":"Department"},{"label":"Unit","column":"Unit"}]`).
**12 intended modes** (see `docs/superpowers/specs/2026-07-24-twelve-segment-expansion-design.md`
for the full map): 4 Head Offices (Group, Upstream Malaysia, Minamas, **NBPOL** — not "NBBOL"),
3 Upstream Ops, 2 SDGI, 2 I&T, 1 Group-led Project.
- **Pilot scope (2026):** the **4 Head Office segments**, all `Levels = [Department, Unit]`
  (shared columns — zero new columns). All 4 term sets exist flat in the `DMS Metadata` term
  group (GUIDs below) and are in the code fallbacks: `DEFAULT_MODES` in `Form.tsx` +
  `BulkUpload.tsx`, `RECON_MODES` in `FolderManager.tsx`. DMS Config `mode` rows are the
  runtime source of truth — the 3 new rows still need creating. 2027 segments need new level
  columns first (see the spec + `2026-07-21-segment-onboarding-plan.md`).
- **User path auto-detection:** the form reads the user's **SharePoint group** memberships via SP
  REST (`spGroups.ts`, `/_api/web/...` — **no Graph, no admin consent**; Graph was fully retired in
  the native-SP-groups refactor), matched against the **`DMS Group Map`** list (`GroupId` = SP group
  **integer** id, `Segment`, `UnitTermGuid`, `Role`). **Only unit-level rows are needed** — one row
  per group, at the leaf. The form walks the term ancestry (`loadTermPath`) to reconstruct
  Segment→Dept→Unit from just the unit term GUID, then `isLeafChainValid` checks only that the
  chain resolved and terminates at that leaf. Do NOT add `MEMBER` rows for the segment/department
  tiers: they are derivable, and the old `isChainAuthorized` that demanded them was removed
  2026-07-29 (spec `2026-07-29-leaf-only-upload-authorization-design.md`) after it refused a
  correctly provisioned uploader with "your account isn't fully provisioned to upload".
  See memory `dms-group-model-per-role-per-library`.
- **Folder routing:** the deepest (leaf) level = the permissioned **Unit** folder, resolved by
  UniqueId via DMS Folder Map. `Year` + `Document Type` subfolders are **ensure-created on demand**
  under the Unit folder and inherit its ACL.
- **Folder NAMES come from `DMS Term Abbreviation`** (keyed by term GUID), NOT from term labels —
  `GHO/GCA/GMB_STRATCOMMS`. Segment codes come from `StagingFolder` on the DMS Config `mode` row.
  Term labels stay full and drive the upload dropdowns; the full label is written to the
  **`Full Name`** column on every folder, and folders carry the **`DMS Folder`** content type so
  the details pane renders it. Three rules: a term with **no abbreviation is SKIPPED**, never
  guessed (no folder → that unit cannot upload); abbreviations must be **unique among siblings**
  or two units merge into one folder with one ACL (reconciliation aborts before creating
  anything); changing one **renames a live folder** on the next run. Spec
  `2026-07-30-folder-abbreviation-naming-design.md`, client guide
  `docs/client/folder-abbreviations-guide.md`.
- **Deleting a term orphans three lists at once** (Abbreviation, Folder Map, Group Map) — a
  re-created term gets a NEW GUID and nothing joins them. Tell the client to **rename terms, never
  delete and re-add**. Reconciliation repairs an orphan only on a **1:1 match of level + label**
  (the label is the one thing that survives, kept in the abbreviation row's `Title`), then rewrites
  the dead GUID across the Group Map. Ambiguous matches are reported, never guessed: `Tax`, `Legal`
  and `PM` each exist under several parents, and a wrong re-point is a permissions grant to another
  department's folder. Abbreviation rows are **never auto-deleted** (authored data, no other copy);
  Folder Map rows still are (derivable). Folders whose term is gone are reported (`NO TERM`) and
  **never** deleted — deleting a term revokes nobody's access. Both passes sit behind the existing
  `incomplete` + clean-run guards. Spec `2026-08-02-term-guid-orphan-repair-design.md`.

> ✅ **TERM STORE — RESOLVED (2026-07-27):** managed **in-site**, no admin center. Client refuses
> tenant term-store access, so the sets were rebuilt in a **site-collection-local `DMS` group** on
> the test site **`/sites/ClarenceDMSTesting`** (managed as Site Collection Admin; Contributor on a
> group is the alternative). **GUIDs are per-site** — those below are ClarenceDMSTesting's, verified
> via the `/_api/v2.1/termStore/sets/{guid}/terms` read path. Code fallbacks
> (`DEFAULT_MODES`/`DEFAULT_SETTINGS`, `RECON_MODES`) already rewired; **DMS Config must carry these
> too** (setting rows `termSet_*` + mode rows' `TermSetGuid`). See spec
> `2026-07-26-in-site-term-store-management-design.md` + memory `dms-term-store-site-collection-pivot`.
```
documentType:    866c5754-258e-401f-8685-03d20ae59b1d
yearPeriod:      023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf
confidentiality: 0d6d1da8-27e5-477f-8684-e8cf169f8fb9
vendor:          RETIRED — Vendor/Customer Name is FREE TEXT; there is no vendor term set.
                 The set was deleted from the site 2026-07-29 and the code no longer reads one.
Group Head Office (business-segment set): 08dd94cb-f76c-431c-9b37-e9c98f739ffc
  Structure: Set → Department → Unit. Verified 14 terms: Group Finance [9 units],
  "Group Legal, Risk ＆ Compliance" [3 units]. Ampersands are FULLWIDTH ＆ (SharePoint requirement).
  Levels = [Department, Unit]. Rebuilt via CSV import — see docs/term-store-import/.
Upstream Malaysia Head Office: 16a52947-57a3-4217-9a49-b48cb8b0dd31  ⚠ PLACEHOLDER — NOT onboarded.
  Client scope as of 2026-07-29 is the other THREE head offices only. This GUID is stale (pre-dates
  the 2026-07-29 term-set rebuild) and is deliberately left in the code fallbacks as a slot-holder.
  It is inert: with no `mode` row in DMS Config the segment is never offered, and the fallback
  constants only surface if the DMS Config read fails. When the client adds the term set, read its
  GUID off the site and update DEFAULT_MODES (Form.tsx, BulkUpload.tsx) + RECON_MODES
  (FolderManager.tsx) — no other change needed.
Minamas Head Office:           9ad00b00-a43c-4a8b-a39a-d0efa89ba706
NBPOL Head Office:             77c3993b-0c3c-4a18-89d9-d69209886322
  (same structure/Levels as GHO; own Department/Unit trees pending client — currently share GHO's
  terms via isAvailableForTagging). Old tenant GUIDs (0540e66e/f7c578a1/032534ab/cb3c0ab7/efa87c6a/
  5ab1c7c4/6ba9a64c/21d7e6fe) are RETIRED for this site.
```
> The old single `department` term set (`eaba82e5-…`) and the `project` term set
> (`94ce322b-…`, lookupStyle "parentMatch") are **retired** by the multi-segment model.
> WARN: Handover PDF lists e1163337-93bb-4b6e-847e-346f51cc6806 for Project Name — this GUID does NOT exist in the tenant. Do not use it.

## Staging Library — Column Internal Names
Verified against live `/fields` API. Do NOT guess from display names.
```
Document_x0020_Type        <- "Document Type" (migrated 2026-07-24 from the old frozen Department_x0020_Type)
Year                       <- plain "Year", NOT Year_x002f_Period (client renamed it; matches FIELDS in Form.tsx)
DocumentDate               <- DateTime, no space encoding
Confidentiality_x0020_Level
LegallyPrivileged          <- Yes/No, written as the STRING "true"/"false"; shown only for the level
                              named by the `legallyPrivilegedFor` DMS Config row (blank = never offered)
Remark                     <- a DEDICATED column, not the built-in _ExtendedDescription
Full_x0020_Name            <- on FOLDERS, not files: the term's real label behind the abbreviated
                              folder name. Read the internal name back — the resolver matches
                              however the space was typed. Needs the `DMS Folder` content type to
                              reach the details pane.
Vendor_x002f_CustomerName  <- "Vendor/CustomerName"; the old plain `Vendor` column was DELETED 2026-07-28
_ExtendedDescription       <- built-in doc Description (Note) — used for "Details" field

# Multi-segment level columns (plain text, label + term-GUID pairs). Written via
# buildLevelFormValues + LEVEL_COLUMNS in Form.tsx. GUID cols have NO "_Tid" suffix —
# SharePoint stripped the spaces (verified against /fields):
Business_x0020_Segment / BusinessSegmentTid   <- "Business Segment" + its Tid
Department             / DepartmentTid         <- clean "Department" name (old taxonomy col removed)
Unit                  / UnitTid
# Add Region/Estate·Mill/Refinery/I&T·Operating·Unit/Project·Name (+ their Tid) when
# the other 4 segments are onboarded; keep LEVEL_COLUMNS in sync.
```
> The old taxonomy `Department` (TaxonomyFieldTypeMulti) and `Project_x0020_Name` columns
> are retired by the multi-segment model. The plain-text `Department` above is a NEW column
> that reuses the freed internal name.

## Critical Rules / Gotchas
1. **DocumentDate** — send as `M/D/YYYY` (US site locale). ISO `YYYY-MM-DD` is rejected. Use `toSpDate()` in Form.tsx.
2. **Project Name column** — must be bound to the **Project** term set (94ce322b-4515-4fda-8f50-35709f1f521d). Re-bind: Staging library settings > Project Name column > Term Set. NOT the Department term set.
3. **`Promise.allSettled` unavailable** — SPFx tsconfig doesn't target ES2020. Use per-set `try/catch` inside `Promise.all`.
4. **`validateUpdateListItem` returns HTTP 200 even on field errors.** Check `HasException` on each result.
5. **Managed metadata value format:** `"Label|GUID"` single / `"Label1|GUID1;Label2|GUID2"` multi. NOT legacy `-1;#Label|GUID`.
6. **File upload body:** raw `File`/`Blob` — NOT FormData (FormData corrupts the upload).
7. **Do NOT use `gulp serve`** — uses Heft toolchain.
8. **Term Store required** for ALL metadata dropdowns — do NOT use SharePoint Choice columns.
9. **`GetFolderByServerRelativeUrl`/`GetFileByServerRelativeUrl` — always pass the path as an OData parameter alias, never as an inline quoted literal.** `getfolderbyserverrelativeurl('<long encoded path>')` returns **HTTP 400** (not 404) once the path is long/deep enough (many `%2F` from encoded slashes in one literal — hit this at ~330 chars / 6 nested levels, well under SharePoint's general 400-char item-path limit, so don't assume length limits rule this out). Correct form: `GetFolderByServerRelativeUrl(@f)?@f='<encoded path>'`. `FolderManager.tsx` already used the alias form everywhere and never hit this; `Form.tsx` used the inline-literal form and silently failed on deep nested subcategory paths — a real folder, correct name, wrongly reported as "not found." If a folder/file check fails unexpectedly, log the actual HTTP status + response body (a 400 means malformed request, not missing data) before assuming a naming/data problem.

10. **`accept` tokens need a leading dot.** `<input type="file" accept="png">` is **not**
    a filter for `.png` — a token without a leading `.` is parsed as a MIME type, is
    invalid without a `/`, and is silently dropped, greying that type out of the file
    dialog. Neither upload web part has a drag-and-drop handler, so the picker is the only
    way in and this is a hard block with no workaround. Worse, the JS validator used
    `endsWith(ext)`, which *accepts* a dotless `png` — one typo, two opposite behaviours.
    Now normalized in `src/shared/allowedFileTypes.ts` (trim → lowercase → prepend the
    dot → drop empties → de-duplicate). Cost half a day on 2026-07-30, made worse by a
    stale page: settings are read once in a mount-time `useEffect`, so **every config
    change needs a hard refresh** before it can be tested.

11. **An emptied multi-value Choice field returns `null`, not `[]`.** Untick every choice on a
    multi-select Choice column and SharePoint returns
    `{"Title":"allowedExtensions","AllowedFileTypes":null}` — verified live 2026-07-30. So the
    *value* cannot distinguish "user cleared it" from "column doesn't exist"; only the **presence of
    the key** can (a nonexistent column makes the whole `$select` return HTTP 400, so the key is
    absent). Reading the value alone made a deliberate hard-block state unreachable and let it
    silently degrade to fallback defaults. See `readAllowedFileTypesField` in
    `src/shared/allowedFileTypes.ts`. Also note the two JSON shapes: `odata=verbose` wraps arrays as
    `{ "results": [...] }` while `nometadata`/`minimalmetadata` return a plain array — pin the
    `Accept` header rather than relying on the default.

## Architecture: Direct REST (not Power Automate)
The 3 instant flows (GetTermSetValues, DepartmentProjectList, UploadToStaging) use "When Power Apps calls a flow (V2)" trigger — not callable from SPFx. Form.tsx uses `context.spHttpClient` directly.

The **Auto-route** automated flow fires server-side after content approval to move approved files to department libraries. Web part job ends at "upload to Staging."

## RBAC
| Group | SP Permission Level | Scope | Can do |
|-------|--------------------|----|------|
| Admin | Full Control | Site | Everything |
| Approver | **Design** | Department Folder in Staging Library (their dept only) | Review & approve pending items for their dept |
| Approver | **Read** | Department Folder in Documents Library (their dept only) | Read approved docs for their dept |
| Uploader | **Contribute** | Department Folder in Staging Library (their dept only) | Upload + tag; sees only their own pending items |
| Uploader | **Read** | Department Folder in Documents Library (their dept only) | Read approved docs for their dept |
| Reader | **Read** | Department Folder in Documents Library (their dept only) | Read approved docs only |

> Design is the only non-Full-Control SP level that includes "approve items".
> Inheritance is broken per department folder in BOTH libraries — not at library level.

## Allowed File Types
Driven by the **`AllowedFileTypes`** multi-select Choice column on `DMS Config`
(row `allowedExtensions`) — the **single source of truth** since 2026-07-30.
`SettingValue` is no longer read for this setting; clear that cell once a site is verified.

Current choices: `.pdf .doc .docx .xls .xlsx` (client policy as of 2026-07-30 — the
12-type list previously documented here was never the client's actual policy).
Everything else (`.exe`, `.bat`, etc.) rejected before upload.

- **Nothing ticked = hard block**, with a message naming the column and the list. The
  client fixes it themselves in one click. Never a silent fallback.
- **Column absent / config unreadable** = code fallback `FALLBACK_FILE_TYPES` in
  `src/shared/allowedFileTypes.ts`, plus an admin-only warning toast and a console line.
  Adding the column is a **required** step when provisioning any new site — there is no
  `SettingValue` bridge any more.
- Adding a type the client cannot already tick means editing the column's **Choices**,
  which needs Manage Lists (site Owner/admin). Ticking needs only item edit rights.
- Keep `FillInChoice` **disabled** — enabling it restores the free-text typo risk the
  column exists to remove.

See spec `docs/superpowers/specs/2026-07-30-allowed-file-types-dropdown-design.md`
and plan `docs/superpowers/plans/2026-07-30-allowed-file-types-dropdown.md`.

## Document Rename Feature
User types a custom name — extension auto-preserved from original file. Illegal chars stripped. Blank = keep original filename. See `buildUploadName()` in Form.tsx.
