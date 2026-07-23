# SDG DMS — Claude Code Project Context

Auto-loaded every session. Keep this up to date whenever decisions change.
Full requirements: `.claude/requirements.md` | Backlog: `.claude/backlog.md`

---

## Project Identity
- **Client:** SD Guthrie (SDG) | **Agency:** Trinergy Digital | **Dev:** Clarence (Junior Digital Developer)
- **Type:** SPFx 1.23.0 React web part — Document Management System upload form
- **Project folder:** `C:\laragon\www\Work\Projects\sd-gatrie`
- **Tenant:** `dcidigitalcom.sharepoint.com` | **Site:** `/sites/SPFX-Sandbox-Testing-Ground`
- **Staging library:** `Staging`

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
has a `Side` (`BusinessSegment` | `Project`), a term-set GUID, and a `Levels` JSON chain
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
- **User path auto-detection:** the form reads the user's Entra group Object IDs via Graph
  `/me/memberOf`, matched against the **`DMS Group Map`** list (`GroupId`, `Segment`, `UnitTermGuid`,
  `Role`). Names are cosmetic — matching is by Object ID (so existing `SDG-*` groups are reused
  without renaming). Only unit-level rows are needed; the form walks the term ancestry
  (`loadTermPath`) to reconstruct Segment→Dept→Unit from just the unit term GUID.
- **Folder routing:** the deepest (leaf) level = the permissioned **Unit** folder, resolved by
  UniqueId via DMS Folder Map. `Year` + `Document Type` subfolders are **ensure-created on demand**
  under the Unit folder and inherit its ACL.

```
documentType:    0540e66e-7cb3-47ac-b0ef-4e3069387394
yearPeriod:      f7c578a1-e0e5-42ff-9e0c-d748cba42ede
confidentiality: 032534ab-9285-4b42-98c6-5c7b0df1f066
vendor:          cb3c0ab7-a959-4200-9b7b-d1e13397d240
Group Head Office (business-segment set): efa87c6a-9536-4f7c-910f-011bf7413b80
  Structure: Set → Department terms (e.g. Group Legal, Risk & Compliance) → Unit terms
  (Group Compliance/GCO, Group Risk, Group Legal). Levels = [Department, Unit].
Upstream Malaysia Head Office: 5ab1c7c4-78d2-43b4-869f-3eab4b1c375c  <- created as "Upstream
  Head Office" — rename in term store to match client list
Minamas Head Office:           6ba9a64c-a363-48fd-afd1-324897df781c
NBPOL Head Office:             21d7e6fe-8f71-4a56-bd2e-e4a2176995a7
  (all 3 same structure/Levels as GHO; Department/Unit terms pending client trees)
```
> The old single `department` term set (`eaba82e5-…`) and the `project` term set
> (`94ce322b-…`, lookupStyle "parentMatch") are **retired** by the multi-segment model.
> WARN: Handover PDF lists e1163337-93bb-4b6e-847e-346f51cc6806 for Project Name — this GUID does NOT exist in the tenant. Do not use it.

## Staging Library — Column Internal Names
Verified against live `/fields` API. Do NOT guess from display names.
```
Department_x0020_Type      <- "Document Type" (created as "Department Type", frozen internal name)
Year_x002f_Period
DocumentDate               <- DateTime, no space encoding
Confidentiality_x0020_Level
Vendor
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
`.pdf .doc .docx .xls .xlsx .ppt .pptx .txt .csv .jpg .jpeg .png`
Everything else (`.exe`, `.bat`, etc.) rejected before upload.

## Document Rename Feature
User types a custom name — extension auto-preserved from original file. Illegal chars stripped. Blank = keep original filename. See `buildUploadName()` in Form.tsx.
