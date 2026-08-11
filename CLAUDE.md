# SDG DMS — Claude Code Project Context

Auto-loaded every session. Keep this up to date whenever decisions change.
Full requirements: `.claude/requirements.md` | Backlog: `.claude/backlog.md`

> 📍 **Current site state, and what is outstanding: `docs/2026-08-07-project-state.md`.**
> Read it before acting on anything in this file. THIS file holds the rules; that one holds what
> is actually true on `/sites/ClarenceDMSTesting` right now — which lists exist, that the library
> was recreated as `Approval Document`, what has been deployed, and the open items in order.

---

## Project Identity
- **Client:** SD Guthrie (SDG) | **Agency:** Trinergy Digital | **Dev:** Clarence (Junior Digital Developer)
- **Type:** SPFx 1.23.0 React web part — Document Management System upload form
- **Project folder:** `C:\laragon\www\Work\Projects\sd-gatrie`
- **Tenant:** `dcidigitalcom.sharepoint.com` | **Site:** `/sites/ClarenceDMSTesting` (rebuilding here
  after being locked out of the old `/sites/SPFX-Sandbox-Testing-Ground`; serve.json still points at
  the old sandbox)
- **Upload/approval library:** title **`Approval Document`**, URL **`/ApprovalDocument`** — they
  DIFFER, and both are needed. Recreated 2026-08-06 (the old `Staging` library was deleted, losing
  its 182 test items and folder ACLs; reconciliation rebuilds the folders). Created without the
  space so the URL stays clean, then retitled. **Never hardcode either half:** `libraryTitle()` and
  `libraryUrlSegment()` in `src/shared/naming.ts` resolve them as one pair via `primeNames()`,
  probing `Approval Document` → `ApprovalDocument` → `Staging`. A wrong title 404s loudly; a wrong
  URL segment fails SILENTLY (`split("/Staging/")` returns a 1-element array, so the caller reads
  `undefined` and routes the file nowhere while reporting success). `LibTarget`/`TARGETS` keep
  `"Staging"` as a LOGICAL key — it is also the stored `Target` value on Group Map library-scope
  rows — and map to the real title at the API boundary via `libApiTitle()`.
  The `stagingLibrary` **DMS Config row is superseded**: it still reads `Staging` on migrated sites
  and would 404, so the live title wins unless that row holds some other non-legacy name.
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
  runtime source of truth — **3 rows now exist, verified live 2026-08-03**: `mode_gho`,
  `mode_minamas_ho`, `mode_nbpol_ho`, all carrying the shared `[Department, Unit]` Levels JSON.
  Upstream Malaysia has **no** mode row, deliberately (its GUID above is the stale placeholder),
  so that segment is never offered and reconciliation never walks it. 2027 segments need new level
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
- **Folder routing:** the deepest **permissioned** level = the **Unit** folder, resolved by
  UniqueId via DMS Folder Map. Everything below it is **ensure-created on demand** and inherits the
  Unit's ACL — nothing below Unit breaks inheritance (client decision, 2026-08-06).
- **Below-Unit structure is CONFIGURABLE since 2026-08-09** (spec
  `2026-08-06-configurable-folder-structure-chain-design.md`; built, **not yet site-tested**).
  A `Levels` entry with `"permissioned": false` is a below-Unit tier: add one and both upload web
  parts render a dropdown and create a folder level, in chain position, with no redeploy.
  - **`permissioned` ABSENT MEANS TRUE**, and only a literal `false` demotes a tier — the string
    `"false"` does not. The default fails loudly (an extra ACL'd folder) rather than silently (a
    tier that needed an ACL inheriting instead).
  - **Permissioned tiers must be a contiguous prefix.** `validateChain` REJECTS a chain with a
    permissioned tier below a non-permissioned one and never sorts it into shape — sorting would
    demote a tier the author meant to protect. A rejected chain BLOCKS upload; it must never route
    to a partial path, which lands one tier shallow in a folder that looks correct.
  - **`UploadMode.levels` is the permissioned prefix only**; the full chain is `UploadMode.chain`.
    `levelValues`/`levelChoices` are indexed against `levels` and the cascade only walks the segment
    term tree, so a below-Unit entry in that array shifts every index.
  - **Reconciliation does NOT walk `Levels`** — it walks the segment TERM TREE. `Levels` supplies
    tier names for reports and, via the chain, the below-Unit grid shape when `recon_gridMode` is on.
    The spec claimed otherwise until it was corrected 2026-08-09.
  - Below-Unit folder names come from the **sanitized term label** (`2026`, `Human Resource`), never
    an abbreviation. No abbreviation row, no Folder Map row, no group needed.
  - **Piece 2, the Structure Manager UI, is BUILT and site-verified (2026-08-11)** — spec
    `2026-08-10-structure-manager-ui-design.md`. Its own web part (`Folder Structure`, GUID
    `ba145053-…`); the client edits a list of levels and never sees the JSON. Verified on NBPOL:
    added a level above Year, columns created in BOTH libraries, `Levels` rewritten, upload landed at
    `NBPOLHO/CDS/UPSUPPORT/testig/2024/Tax Return`.
    - It **seeds the editor with `effectiveOnDemandTiers`**, not the raw chain. A segment with
      nothing below Unit is silently running on the built-in `Year → Document Type` pair, so an empty
      list would make the first added level REPLACE that pair — dropping Year and Document Type from
      every future path, with no error. The seeded pair carries no `tidCol` on purpose: those two are
      managed metadata, and a derived `YearTid` would start writing a bare label into a taxonomy field.
    - The term set ID is **resolved as it is typed** and the set's NAME shown back. Malformed / 404
      block Add; a resolvable-but-empty set and an unreachable term store warn only. A wrong ID here
      is this screen's worst failure — reconciliation ignores below-Unit tiers, so it saves clean and
      surfaces later as one permanently empty dropdown that blocks upload, for someone else.
    - Columns are created `Options: 8`, so they are **not added to any view** — the client switches
      them on per view. Deliberate: `12` would reshape every view they arranged, once per level.
    - Still OUT of scope: adding a whole new segment (**slice B**, not built).
  - **A DOCUMENT ALWAYS LANDS AT THE END OF THE CHAIN** (client rule, 2026-08-12). Every tier is
    built, and a tier with no value in a folder's path is a gap the admin fills — wherever it sits.
    Before this, `planLeaf` capped at the leaf's own depth, which made **adding a level at the BOTTOM
    a silent no-op**: nothing looked misplaced, the change activated, and old documents sat a level
    shallower than new ones. Exceptions: a tier that does not apply to a unit never enters the
    calculation (optional SubUnit), and a folder with nothing recognisable left is reported rather
    than placed. Cost, accepted: a genuinely shallow folder now asks for its deeper values too.
  - **Piece 3, subtree migration, is BUILT and VERIFIED LIVE — all five positions** (add mid-chain,
    reorder, remove-clean, remove-with-collision, **add at the bottom**), 2026-08-11/12 — spec
    `2026-08-11-subtree-migration-design.md`, tab 2 of the same web part, logic in
    `shared/subtreeMigration.ts` (71 tests). Three results worth keeping: a REORDER re-stamped **0**
    documents (it changes no tier's value, so anything else would be a bug); the collision run ended
    with the **document count unchanged** — both files side by side under different names, neither
    overwritten; and the bottom-add tidied **0** folders, because documents moving *down* vacate
    nothing. Note the bug in the bottom-add case survived four passing tests, all of which happened
    to use full-depth folders — the untested position was the broken one.
    - **ADD, REORDER and REMOVE are ONE operation**, and treating them separately is what produced a
      tool that could only do the first: work out which TIER each path segment belongs to, then
      rebuild the path in tier order. A tier with no segment is a gap the admin fills (add);
      segments out of order come back sorted (reorder); a segment whose tier is gone is absent from
      the result (remove — and therefore a COLLAPSE). Combinations work, which is what a client
      editing a live structure actually produces.
    - **The scan walks every LEAF folder, not just children of the Unit.** The shallow version
      reported "nothing to move" for a reorder — `Credit2` was still a valid first tier — then
      activated the new structure against folders still in the old order. Silent wrong success.
    - **FILES move; folders are ensure-created.** Collapse forces it: two `2024` folders cannot both
      move into one place, a resolved collision renames a file, and a destination already holding
      documents is a merge rather than a move. One request per file beats two code paths where the
      rarely-tested one handles the dangerous case.
    - **A filename collision gets a rename FORM, not a refusal** (client, 2026-08-11). Files already
      at the destination count as claimants — not migrating, but they own the name, and that is the
      case that silently overwrites. Suggested names carry the removed tier's value
      (`a (testig).pdf`), never a number: the value is *why* the files differed.
    - **A verified fact the design rests on: a FILE move preserves approval status** (2026-08-11,
      alongside the folder-move check of 2026-08-10).
    - **A STRAY does not block activation.** The tool cannot resolve one, and letting it hold a
      structure change hostage forever leaves the client stuck; it is counted in the result instead.
    - Ambiguity resolves to "already correct" — moving a correctly filed folder needs evidence.
    - **A tier with no values for a unit does not apply to that unit** (optional SubUnit), so its
      files legitimately sit shallower and nothing moves. But options that could not be READ make
      every folder look misplaced — so unreadable skips the whole unit. Empty ≠ unknown, again.
    - **Moves are safe because of three facts verified 2026-08-10:** the subtree travels with the
      folder, approval status survives, and UniqueId survives — so Folder Map needs no repair and
      nothing needs re-approving. **Turn Auto-route and the folder-approval flow OFF first.**
    - **Destination folders it creates must be stamped Approved** where the library moderates, or
      every document moved into them is invisible to the whole unit.
    - **Metadata is stamped from the PATH, after the moves, re-read fresh.** That is what makes a
      half-finished run safe to repeat: no progress is recorded anywhere, so a second run
      re-derives what is left. It also repairs files a hand-edited chain mis-tagged.
- **A STRUCTURE CHANGE ON A SEGMENT IN USE IS STAGED, NOT LIVE (2026-08-11, client's request).**
  The Structure Manager writes to a **`PendingLevels`** Note column on the mode row; `Levels` — the
  only one the upload form and reconciliation read — is untouched until the migration finishes and
  **applies it as its own last step**. Before this, saving went live immediately while existing
  folders stayed a tier above, so a unit had two shapes at once and the people who met that state
  first were uploaders nobody had told: it reads as a broken system, not an unfinished admin task.
  - **Applying is never a separate button.** Between "moved" and "applied", every upload lands in
    the old shape again and re-creates the drift just cleaned up.
  - **It applies only when a FRESH scan finds no drift left** — never a tally of what the run
    attempted. A partially migrated segment stays pending and reports how many units remain.
  - **Empty segments activate immediately**, and a staged-then-emptied one has its pending chain
    cleared on save. A staged change with nothing to move still gets an explicit Apply button, or
    it could never go live.
  - Editing a segment that already has a pending change edits **the pending chain**. Otherwise a
    second edit silently discards the first and the migration applies a shape nobody reviewed.
  - The typed-`CHANGE` gate on save is **gone**: it warned about the side-by-side state that no
    longer happens, and saving is now inert. The gate that matters is the typed `MOVE`.
- **`SubUnit` (client, 2026-08-10) is a FIXED below-Unit tier that INHERITS — not a new permissioned
  tier.** Structure becomes `Business Segment → Department → Unit → SubUnit → Year → Document Type`,
  and the client may not reorder or remove any of the first four. But SubUnit is authored
  `"permissioned": false`, so it needs **no groups, no Group Map rows, no abbreviation rows and no
  code change** — just a term set, `SubUnit`/`SubUnitTid` columns in BOTH libraries, and one
  `Levels` entry before `Year`. Subunit data not yet supplied by the client.
  - **THE UNIT IS STILL THE SMALLEST CONFIDENTIALITY BOUNDARY.** Two SubUnits under one Unit see
    each other's documents completely — the ACL is on the Unit folder and SubUnit inherits it.
    Confirmed by the client 2026-08-10 when asked directly. Everything in
    `docs/client/document-visibility-within-a-unit.md` stands unchanged, and "put them in different
    subunits" is NOT a way to separate two people.
  - **"Fixed" and "permissioned" are now different things.** The Structure Manager currently locks
    only the PERMISSIONED prefix (Segment/Department/Unit), so SubUnit renders as an editable
    below-Unit level and the client could reorder or remove it. Acceptable while SubUnit is unbuilt
    client data; if they are told SubUnit is fixed, lockedness must stop being derived from
    `permissioned`.
- **Folder NAMES come from `DMS Term Abbreviation`** (keyed by term GUID), NOT from term labels —
  `GHO/GCA/GMB_STRATCOMMS`. Segment codes come from `StagingFolder` on the DMS Config `mode` row.
  Term labels stay full and drive the upload dropdowns; the full label is written to the
  **`Full Name`** column on every folder, and folders carry the **`CRS Folder`** content type so
  the details pane renders it. That name is resolved, not hardcoded: `FOLDER_CONTENT_TYPE_CANDIDATES`
  in `FolderManager.tsx` probes `CRS Folder` then `DMS Folder`, preferring the first, so a library
  mid-rename is not stamped with the type being retired. `CRS Folder` was created 2026-08-06 (site
  content type, group `CRS Content Types`, **parent group `Folder Content Types`, parent `Folder`** —
  a Document/Item parent will not attach to a folder). It must be added to BOTH libraries with the
  `Full Name` column on it, and hidden from the New button. `DMS Folder` still exists and cannot be
  deleted until reconciliation has re-stamped the `Documents` folders. Three rules: a term with **no abbreviation is SKIPPED**, never
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

> ⚠ **`Documents` MUST carry the same columns under the same internal names.** It was missing
> `Remark`, `LegallyPrivileged`, `ProjectName` and `Vendor/CustomerName` on ClarenceDMSTesting until
> 2026-08-10 — for months, with nothing reporting it. Two silent failures came from that one gap,
> and neither presents as a column problem:
> 1. **Bulk upload tags NOTHING.** It writes `Remark` + `LegallyPrivileged` unconditionally, and one
>    unknown field name fails the WHOLE `validateUpdateListItem` call — every column lost, not just
>    the missing ones. Shows only as a "No tags" badge per file.
> 2. **Auto-route drops metadata on every approved document.** SharePoint's copy carries over only
>    columns that EXIST at the destination; the rest vanish with no error and a green run. Approved
>    files were arriving with no Remark, no ProjectName and no `LegallyPrivileged` — a legal marker,
>    absent from the library where the documents actually live.
>
> Diff `/fields?$select=Title,InternalName` on BOTH libraries when provisioning a site. A matching
> display name over a DIFFERENT internal name fails exactly like an absent column, and looks right.

## Approval Document Library — Column Internal Names
> Re-verified 2026-08-06 against the recreated library (`/fields`). All 15 present and correct.
> Recreating is what makes this section load-bearing: internal names are derived from the title a
> column is CREATED with, once, permanently — so a column must be created under the name that
> yields the required internal name, then renamed. `DocumentDate` displays as "Document Date" but
> has no `_x0020_` precisely because it was made that way.
Verified against live `/fields` API. Do NOT guess from display names.
```
Document_x0020_Type        <- "Document Type" (migrated 2026-07-24 from the old frozen Department_x0020_Type)
Year                       <- plain "Year", NOT Year_x002f_Period (client renamed it; matches FIELDS in Form.tsx)
DocumentDate               <- DateTime, no space encoding
Confidentiality_x0020_Level
LegallyPrivileged          <- Yes/No, written as the STRING "true"/"false"; shown only for the level
                              named by the `legallyPrivilegedFor` DMS Config row (blank = never offered)
Remark                     <- a DEDICATED column, not the built-in _ExtendedDescription
Full_x0020_Name            <- on the recreated library. The OLD Staging library had `FullName0`
                              (created as "FullName", then renamed) and `Documents` still does —
                              which is why pickFullNameField matches on DISPLAY TITLE or internal
                              name, never internal name alone: nameKey keeps digits, so
                              "FullName0" would never match "fullname".
                           <- on FOLDERS, not files: the term's real label behind the abbreviated
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
   **Display is a separate concern:** every user-facing date reads **`DD/MMM/YYYY`** (`08/Jan/2035`) — the
   agreed client format. In code that's `formatDate()` in `ApprovalDocument.tsx`; in SharePoint views it's
   column formatting (there is no month-name token, so build it with
   `substring('JanFebMarAprMayJunJulAugSepOctNovDec', getMonth(@currentField) * 3, getMonth(@currentField) * 3 + 3)`
   — `getMonth()` is **0-based**, which is exactly why the `* 3` offset lands right; do not "fix" it by adding 1).
   Formatting is display-only: sort, filter and the Auto-route flow all still use the stored value. Never
   re-parse a date SharePoint already formatted (see the `Document Date` row in ApprovalDocument) — that
   reintroduces the M/D/YYYY trap.
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

10b. **A STALE UPLOAD PAGE FILES INTO THE OLD FOLDER SHAPE, and nothing about it looks wrong.**
    Settings and modes are read once in a mount-time `useEffect` (see #10), so a tab left open
    across a structure change keeps building the OLD path — and every upload from it re-creates the
    two-shapes state a migration just cleaned up. Found live 2026-08-11: a file filed with no
    `Credit_Card` level hours after that level went live, noticed only because the migrator listed
    the folder. The upload succeeds, the file lands somewhere plausible, and no error exists
    anywhere. **Both upload web parts now re-read the mode row's `Levels` immediately before writing
    and refuse with "reload the page" if the chain moved** (`chainSignature` + `freshChainFor` in
    Form.tsx and BulkUpload.tsx). A read FAILURE must never block — it proves nothing, and taking
    the form down over a transient error is worse than the risk it guards against. This closes the
    last route by which folder drift can reappear on its own after a migration.

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

12. **A list's title and its URL are independent, and only one of them fails loudly.** SharePoint
    fixes a list's URL at CREATION and never moves it on rename — which is how `CRS Config` sits at
    `/Lists/Config` and `Approval Document` at `/ApprovalDocument`. Use that deliberately: create a
    library without spaces, then retitle. But it means **one literal can no longer serve both**.
    `getbytitle('<wrong>')` returns 404 — obvious. `path.split("/<wrong>/")` returns a 1-element
    array, so `[1]` is `undefined` and the code carries on: the file routes nowhere, the log says
    success. Always resolve the pair together from `RootFolder/ServerRelativeUrl` on the response
    that also carried the Title (`primeNames` → `libraryTitle()` / `libraryUrlSegment()`), never
    derive one from the other. Same class of bug as #9: the loud failure is the safe one.
    Recreating a list ALSO changes `ListItemEntityTypeFullName` (renaming does not) — read it off
    the list at runtime, as `FolderMap.tsx` does, or `__metadata` writes break.

## Architecture: Direct REST (not Power Automate)
The 3 instant flows (GetTermSetValues, DepartmentProjectList, UploadToStaging) use "When Power Apps calls a flow (V2)" trigger — not callable from SPFx. Form.tsx uses `context.spHttpClient` directly.

The **Auto-route** automated flow fires server-side after content approval and **MOVES** approved
files to `Documents` — copy, stamp, delete source. Web part job ends at "upload to the approval
library."

> 📘 **Full flow configuration + migration runbook:
> `docs/superpowers/specs/2026-08-08-auto-route-flow-and-draft-isolation.md`.**
> Power Automate config is NOT in source control — that spec is the only record of it. Rebuild
> from it verbatim; several settings look cosmetic and are not.
> **Sign in as the SERVICE ACCOUNT before creating either flow (§0).** A flow runs under its
> connection, and the connection is created implicitly by the first action — so the builder's
> account is baked in for life. Built as a person, both flows stop **silently** when that password
> changes, and present as "my colleague's folder doesn't exist". On the test site this is knowingly
> unfixed; on SDG's tenant it must be right from the first action.

Working as of 2026-08-08, verified with a guest uploader and an admin approver:
- Path split is on **`ApprovalDocument/`** (the URL segment, no space) — was `Staging/`.
- The copy keeps the **uploader** in `Created By`, `Modified By` and `Created`, via
  `validateUpdateListItem` with `bNewDocumentUpdate: true` (`Author`, `Editor`, `Created`).
  `Created` must be `M/d/yyyy h:mm tt` on that endpoint; a MERGE to the same item wants ISO.
- The source item is then **DELETED by item id** (`X-HTTP-Method: DELETE`), gated on the stamp
  succeeding. The connector's `Delete file` action does NOT work here — its `File Identifier`
  never resolves, because the trigger is a *list* trigger pointed at a library.
- **The delete is load-bearing for security, not housekeeping:** an approved file left behind is
  visible to every PIC in the unit, which defeats the draft isolation below.

> ⚠ **Power Automate header keys must NOT include the colon.** The key box wants `Accept`, not
> `Accept:`. With the colon the header does not exist, and the symptoms look unrelated to each
> other: responses silently come back `odata=verbose` (so every `body('X')?['Field']` is null →
> *"Not well formatted JSON stream"*), `validateUpdateListItem` reports `HasException: false`
> and changes nothing, and a MERGE goes as a plain POST (*"The parameter AuthorId does not exist
> in method GetById"*). This cost two hours on 2026-08-08 and produced a false conclusion that
> `Author` was unwritable. Read the action's raw **Inputs** — the colon is visible there and
> nowhere else. And **never** conclude a field is unwritable from a clean response: re-read the item.

## Per-uploader file isolation — approval library YES, Documents NO (2026-08-08)
Spec `2026-08-08-auto-route-flow-and-draft-isolation.md`. This **reverses** the 2026-08-06
"out of scope" decision for the approval library, and **confirms** it for Documents.

**Approval Document — works, built, verified.** Content approval **on**, Draft Item Security =
*"Only users who can approve items (and the author)"* (`DraftVersionVisibility = 2`). A PIC sees
their own pending/rejected files; a peer PIC sees nothing; the Head of Unit sees all because they
hold `ApproveItems`. Approved files leave via Auto-route. The **author exception works for FILES**
(verified with a guest account) but NOT for a folder created by someone else — hence folder approval.

**Folders must be Approved or navigation breaks.** The upload form ensure-creates `Year` and
`Document Type` folders **as the uploader**, so they arrive Pending and are invisible to every
other PIC — who then cannot reach their own file inside. Uploaders cannot self-fix (needs
`ApproveItems`). Backlog: `Downloads/approve-folders.js`. Ongoing: a **separate** flow with trigger
condition `@equals(triggerOutputs()?['body/{IsFolder}'], true)` that MERGEs
`{"OData__ModerationStatus": 0}`. That `{IsFolder}` guard is the whole safety of it — approving a
FILE would skip human approval and hand it straight to Auto-route.

**`Documents` must have content approval OFF — verify it on every site.** If it is on, every item
the Auto-route flow creates there (the `Year`/`Document Type` folders AND the copied file) arrives
**Pending** and is invisible to every read-only viewer: the uploader's approval email link is
denied, and the unit sees a folder that looks empty or absent. It presents as a permissions bug and
is not one. Tell-tale without a query: the view bar shows `Approve/reject Items` + `Show All Files`,
views SharePoint adds only when moderation is on. Fix = Versioning settings → *Require content
approval* = No; everything Pending becomes visible at once. Nothing in the code enables it —
reconciliation only READS `EnableModeration` to decide whether to stamp new folders Approved.
Spec §5.7.

**A PIC needs the base `_MEMBER` group to open the emailed link.** `*_UPL` grants upload on the
approval library and NOTHING in `Documents` — so an upload-only PIC gets "your document is now
available" and an AccessDenied. Working as designed (the denial URL still carries `listItemId`, so
the file resolved and was then refused), but the client must decide: PICs who should read their
unit's approved documents go in the unit base group too. Spec §5.8.

**Documents — NOT achievable, do not re-attempt.** Hiding an item from someone with Read requires
moderation; seeing a hidden item requires `Approve Items`; and **`Approve Items` cannot be
separated from `Edit Items`** (verified in the permission-level editor — ticking one ticks the
other). So view-only oversight roles cannot see other people's files. `ReadSecurity = 2` is also
out: its exemption is `Manage Lists` at LIBRARY scope, so folder-scoped approvers would be
restricted too, and it hides every reconciliation-created folder from uploaders. Per-file ACLs hit
the 50k scope ceiling. A per-uploader folder tier would work but changes the client's fixed folder
architecture. **The unit folder remains the smallest confidentiality boundary — the answer to
"these two must not see each other" is separate units.** A `Created By = [Me]` view is a
convenience only and must never be described as isolation.

## RBAC — six personas (2026-08-07, spec `2026-08-07-role-model-simplification-design.md`)
Twelve personas collapsed to six when the client moved approval to Head of Unit and made Head of
Department view-and-delete only. `PERSONAS` in `groupMapModel.ts` is the source of truth.

| Persona | Roles | Scope | Documents | Approval Document (Staging) |
|---|---|---|---|---|
| C-Level (global) | `GLOBAL` | all segments | view everything | **none** |
| C-Level (segment) | `SEGVIEW` | one segment | view that segment, all the way down | **none** |
| Head of Department | `DEL` | department | view **every unit under the dept**, + delete approved docs | **none** — does NOT approve |
| Head of Unit | `APR`, `DELS` | unit | view own path | **approve + view every file in their unit, + delete pending/rejected** |
| PIC | `UPL` | unit | **view own path** | upload at **any** confidentiality level |
| SDG Employee | `MEMBER` | unit | view own path | none |

- **Navigation starts at the business segment in BOTH libraries.** Reconciliation grants each
  group `Read` up its own path (the "ancestor browse" corridor); siblings get nothing and stay
  security-trimmed. This was deleted on 2026-08-04 for a requirement the client then withdrew, and
  **restored 2026-08-07** — its absence is invisible on an existing site but renders a NEW library
  empty for every non-admin.
- **`recon_departmentFanOut` now defaults ON.** Unit folders have unique permissions, so without
  it a Head of Department sees a department folder that appears to contain no units.
- **`recon_revokeAncestorRead` is hard-off**, config row ignored — it would fight the restored
  grant and break navigation.
- **Confidentiality is metadata, not a permission.** Any PIC may upload at any level; there is no
  HC persona and `HC` is offered nowhere.
- **ONE GROUP PER PERSON (2026-08-09, spec `2026-08-09-persona-driven-folder-access-design.md`).**
  `UPL` and `APR` now appear in **both** libraries, so a PIC or Head of Unit no longer needs the
  unit's base group to read Documents — hence `MEMBER` dropped from HoU, and from HoD where `DEL`
  (= Read + Delete Items) already covered it. **The level is library-dependent:**
  `permissionForRole(lib, role)` downgrades `UPL`/`APR` to **`Read`** on Documents. Never read
  `ROLE_TO_PERMISSION` directly — the flat table would grant `CRS Upload` there, letting every PIC
  edit approved documents. `DELS` stays Staging-only, so a Head of Unit still cannot delete an
  approved document. Migration = redeploy + **re-run reconciliation**; no row changes.
  Consequence to keep stating: **every PIC now reads every approved document in their unit**, at any
  confidentiality level.
- **`DELS` now belongs to Head of Unit** (was: no persona). It is delete on the **approval library
  only** — pending and rejected files.
- **Folder Access is persona-only.** The role chips are gone: picking a persona applies its roles.
  A library toggle above the group field splits the personas — Documents (C-Level ×2, HoD, SDG
  Employee) vs Approval Document (HoU, PIC). That grouping is **derived** via
  `personaTouchesStaging()`, never hand-listed: filing C-Level under the approval library would
  present the widest accidental grant in the system as normal.
- **`SEGVIEW`/`GLOBAL` must NEVER appear in `LIBRARY_ROLES.Staging`** — a segment-wide viewer
  there reads every unapproved draft in the segment.

> Design is the only non-Full-Control SP level that includes "approve items" — but this site uses
> the custom `CRS Approve` level, so **verify it actually contains Approve Items**: draft-item
> security and the whole approval flow key off that permission, not off a group name.
> Inheritance is broken per department folder in BOTH libraries — not at library level.

> ⚠ **PARTIALLY SUPERSEDED 2026-08-08** — see the per-uploader isolation section above and spec
> `2026-08-08-auto-route-flow-and-draft-isolation.md`. In the **approval library** isolation IS
> achievable and is now built, using content approval + Draft Item Security (approver + author),
> because the only roles there are uploader and approver. Everything below still stands for
> **Documents**, where view-only oversight roles make it impossible.
>
> ❌ **PER-UPLOADER FILE ISOLATION IN `Documents` IS OUT OF SCOPE (2026-08-06, re-confirmed
> 2026-08-08).** The client asked for
> uploaders in the same `*_UPL` group to see only their own files inside the shared unit folder
> (Clarence not seeing Bayajit's). **The unit folder is the smallest confidentiality boundary this
> system has** — do not re-propose any of these mechanisms:
> - **Moderation in `Documents`** — would require every viewer (HoD, HoU, SEGVIEW, GLOBAL) to hold
>   `Approve Items` just to read, and **`Approve Items` cannot be separated from `Edit Items`**
>   (verified in the permission-level editor 2026-08-08: ticking one ticks the other). The client
>   requires those roles to be view-only, so this is a dead end, not a configuration problem.
> - **List item-level permissions (`ReadSecurity=2`)** exempts only holders of **Manage Lists**, which
>   is evaluated at LIBRARY scope. Approvers hold Design on the FOLDER and nothing at library root,
>   so they would be restricted too and approval would break. Granting them library-level rights to
>   fix it re-opens cross-department visibility — the exact thing the folder ACLs exist to prevent.
>   Also library-wide only (no per-unit opt-out) and it risks hiding admin-created folders from
>   uploaders, breaking navigation and the form's folder resolution.
> - **Per-file unique ACLs at upload** works and scales to the pending-queue depth, but the client
>   expects volumes that approach SharePoint's **50,000 unique-permission-scope ceiling** per list,
>   with performance degrading well before it. Rejected as a long-term design.
>
> The structural answer if a client genuinely needs two people isolated: **they belong in different
> units.** Adding a unit costs one term + one abbreviation + 3 groups and stays inside the model.
>
> 📄 **Client-facing answer, ready to show: `docs/client/document-visibility-within-a-unit.md`.**
> Plain language, no jargon — use it instead of re-deriving the argument. Its framing is the one
> that lands: *drafts and rejections are already private; approved documents are shared because the
> unit must be able to find its own records.* It also carries the honest caveat: this IS possible if
> the folder structure changes (a per-person folder under each unit), so never say "impossible"
> flat. Say **impossible without changing the filing from unit-based to person-based** — the cost
> that usually decides it is that "all of CORU's 2026 Tax Returns" stops being a folder you open and
> becomes a search.

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
