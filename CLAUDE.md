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

## Term Set GUIDs (DMS Metadata group, sandbox tenant)
```
documentType:    0540e66e-7cb3-47ac-b0ef-4e3069387394
department:      eaba82e5-3e5f-4719-9a76-091f034ad407
yearPeriod:      f7c578a1-e0e5-42ff-9e0c-d748cba42ede
confidentiality: 032534ab-9285-4b42-98c6-5c7b0df1f066
vendor:          cb3c0ab7-a959-4200-9b7b-d1e13397d240
project:         94ce322b-4515-4fda-8f50-35709f1f521d  ← SEPARATE term set under DMS Metadata group.
                 Structure: top-level terms mirror dept names (Account, Finance, HR, IT)
                 and sub-terms are the actual projects (e.g. Project > Account > AI Engineer).
                 DMS Config uses lookupStyle "parentMatch": pick dept → match parent label → load children.
```
> WARN: Handover PDF lists e1163337-93bb-4b6e-847e-346f51cc6806 for Project Name — this GUID does NOT exist in the tenant. Do not use it.

## Staging Library — Column Internal Names
Verified against live `/fields` API. Do NOT guess from display names.
```
Department_x0020_Type      <- "Document Type" (created as "Department Type", frozen internal name)
Department                 <- TaxonomyFieldTypeMulti
Year_x002f_Period
DocumentDate               <- DateTime, no space encoding
Project_x0020_Name         <- bound to the Project term set (verify GUID in Term Store)
Confidentiality_x0020_Level
Vendor
_ExtendedDescription       <- built-in doc Description (Note) — used for "Details" field
```

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
