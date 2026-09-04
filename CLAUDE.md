# SDG DMS — Claude Code Project Context

> ⛔ **BEFORE YOU CHANGE ANYTHING: READ THE CURRENT SOURCE, AND CONFIRM THE LIVE STATE.** Standing rule
> from the client, 2026-08-18 — *"always check the code and current situation first before we make a
> changes or not due to old file or old comments new session claude will destroy the progress."*
>
> **This file goes stale, and it is detailed enough to be convincing when it does.** In one day it was
> found claiming SubUnit was unbuilt (it was built), that the abbreviation count was wired up (it was
> computed and never handed over), and that `MEMBER` groups carry no suffix (they carry `_EMPLOYEE`). The
> migration runbook recorded 71 abbreviation codes as saved after the client had cleared the list.
>
> So: **grep the file before editing it** — a comment is evidence of intent, not of behaviour. **Check the
> INSTALLED app version in Site Contents** before diagnosing any UI report; the catalog listing is not the
> same thing, and a site-collection catalog beats the tenant one. **Ask which site** — this project spans
> `/sites/CRS` on SDG's tenant and `/sites/ClarenceDMSTesting` on ours, with different term sets, GUIDs and
> counts; never check one's data against the other's reference files. **One fix, one deploy, one check.**
> See memory `feedback-verify-live-state-before-changing`.

Auto-loaded every session. Keep this up to date whenever decisions change.
Full requirements: `.claude/requirements.md` | Backlog: `.claude/backlog.md`

> ✅ **READ FIRST: `docs/2026-08-18-bulk-groups-handoff.md`.** dcistaging is fully provisioned — 308
> groups, 790 Group Map rows, folders and ACLs in all four libraries — and **UPLOAD → APPROVE → ROUTE
> IS VERIFIED END TO END there (2026-08-18)** with two guest accounts: `Created By` survives the copy,
> the metadata survives it, the approval email carries the right path, and a NON-ADMIN uploader sees the
> routed file in `Documents`. The bulk run is idempotent (second press wrote nothing).
> Still untested anywhere: **anything HC** (no `CRS Folder` content type on the HC pair, and neither HC
> Power Automate flow exists) and the **deletion/share request** flows.

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
**13 intended modes** (see `docs/superpowers/specs/2026-07-24-twelve-segment-expansion-design.md`
for the full map): 4 Head Offices (Group, Upstream Malaysia, Minamas, **NBPOL** — not "NBBOL"),
3 Upstream Ops, 2 SDGI, 2 I&T, 1 Group-led Project, 1 Group Business Ventures & Transformation
(**added by the client 2026-08-17**; head-office shape, no term set yet, in no code fallback).
- **ALL THIRTEEN FAMILIES ARE EXACTLY TWO PERMISSIONED LEVELS, and `validateNewSegment` REFUSES fewer
  than two** (client, 2026-08-17: *"I think best to force them to create two not one. atleast two"*).
  Head offices `Department → Unit`, Upstream Ops `Region → Estate/Mill`, SDGI `Refinery → Department`,
  I&T `I&T Operating Units/Department → Unit`, Group-led Project `Department → Unit`.
  - **THE RECURRING MISREADING IS COUNTING THE BUSINESS SEGMENT AS A LEVEL.** It is `StagingFolder`,
    above the chain — which is why "Head Office is three levels" feels right and is wrong by one. It
    corrupted three separate rows of the segment map and flipped I&T twice in one day: Group-led Project
    was recorded as **three** because `Project Name` (its top folder name) was counted, and I&T as
    **one** because someone had to drop a real tier to make three arithmetic work.
  - **`I&T Operating Units/Department` is ONE TIER NAME containing a slash**, with `Unit` below it —
    clarified by the client after it had been read as two tiers and then as one. A `/` inside a level
    name is the one punctuation mark here that means nothing structural, which is precisely why it
    misleads. The derived column is `I&TOperatingUnitsDepartment` (illegal characters are removed, not
    substituted — as `Estate/Mill` → `EstateMill`).
  - **ONE TIER TOO MANY IS THE SILENT DIRECTION.** The depth check demands the term set's depth equal
    the tier count, and reconciliation walks the term tree capped at that count — so an extra tier moves
    every unit's ACL a level DOWN onto a folder no Group Map row points at. It grants successfully and
    looks correct. One too FEW is loud: the form refuses and writes nothing.
  - **The two-tier floor is NOT redundant with the depth check**, which is the reason it exists. One
    declared tier against a **one-deep** set passes that check cleanly — and a set is one-deep exactly
    when the departments have been authored and the units have not. The segment saves with the
    **department** holding the access, so every unit under it shares one folder and one ACL. It refuses
    rather than warns because the depth check already warns-and-allows when depth is unknown, and a
    second soft signal would leave the one shape that reads as complete advisory in both places.
  - `SegmentCreator`'s hint copy asserted "I&T uses a single level" while the spec said two; it now
    names the SDGI pair, says **at least two are required**, and says outright that the business segment
    is the Top folder name and not one of these. That sentence is the only thing standing between a 2027
    onboarder and this bug.
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
- **⚠ ONE TERM, ONE FOLDER MAP ROW — enforced since 1.0.173.0, and its absence took a whole site's
  uploads down (2026-08-19).** Reconciliation indexed the map with
  `mapByTerm.set(r.termGuid.toLowerCase(), r)`, and **`set` OVERWRITES** — so a term with two rows
  collapsed to the last one. The repair pass repointed THAT row at the rebuilt folder and never saw
  the other, which went on pointing at a folder the same run had just deleted.
  - **The symptom is a silent upload REFUSAL, and it names the wrong cause.** `lookupFolderMapping`
    reads `$top=1` and takes whichever row comes first; when that is the stale one, `GetFolderById`
    answers 404, `probeFolderUploadAccess` returns `missing` — deliberately CONCLUSIVE, because
    security trimming answers 404 too — the path is dropped, and the uploader is told *"your unit
    isn't ready to receive uploads yet… run folder reconciliation"*. Reconciliation is what created
    the state. **56 of 67 terms were like this and every uploader on the site was refused**, while
    the folder existed and `EffectiveBasePermissions` on it read `AddListItems` granted.
  - **`writeFolderMapping` UPSERTS now.** It was an unconditional POST called only when the
    in-memory index missed, so one missed lookup — a braced GUID, an interrupted run — created a
    permanent duplicate. **A failed existence check REFUSES rather than creating:** a run that
    writes nothing is repaired by the next one; a run that writes a second row is repaired by
    nothing.
  - **Reconciliation now reports and removes duplicates**, probing each row's folder and deleting
    only those whose folder is gone — **naming each one**. `chooseKeeper` in
    `shared/folderMapDuplicates.ts` (pure, 14 tests) **fails CLOSED against this codebase's habit**:
    if NO row resolves it deletes nothing and says so, because a wrong deletion takes a unit's
    upload path away and the person who discovers it is an uploader.
  - **Keys are `normalizeTermGuid`, never `toLowerCase`.** The upload form already normalised
    (braces and whitespace as well as case); reconciliation did not, so the two halves could
    disagree about whether a term was mapped — a second route to the same duplication. Not what
    fired on 2026-08-19, where every stored GUID was already bare.
  - **DELETE-AND-REBUILD IS A SUPPORTED RECOVERY PATH** — deleting every folder and re-running is
    how a site is cleaned of strays — so it must not end with an upload form that refuses everyone.
- **Folder routing:** the deepest **permissioned** level = the **Unit** folder, resolved by
  UniqueId via DMS Folder Map. Everything below it is **ensure-created on demand** and inherits the
  Unit's ACL — nothing below Unit breaks inheritance (client decision, 2026-08-06).
- **A PATH IS OFFERED ONLY IF THIS USER CAN ACTUALLY UPLOAD INTO IT (2026-08-12, client's
  instruction** — *"I do not want the half built segment to show… once the Segment and structure and
  group is properly assigned only then the uploader should be able to upload"*). Spec
  `2026-08-12-provisioned-segment-visibility-design.md`. `Form.tsx` reads **DMS Folder Map** at mount,
  drops any authorised path whose **leaf** term has no row, then **probes the survivors' folders for
  `AddListItems`** (`filterProvisionedPaths` + `filterReachablePaths` in `shared/segmentReadiness.ts`,
  `probeFolderUploadAccess` in `dmsFolderMap.ts`).
  - This closed **two symptoms with one cause**: a new segment appeared before it had folders, and a
    user added to a unit group could pick that unit before reconciliation had run (the origin of the
    upload form's HTTP 403). The form offered intent, and only asked whether the destination was
    usable at *write* time.
  - **EXISTENCE IS NOT THE QUESTION — the first build got this wrong and a live test caught it.**
    Reconciliation creates folders from the **term tree** (keyed on abbreviations) and grants group
    ACLs in a **SEPARATE pass** (keyed on Group Map). So creating a group creates no folder: two new
    NBPOL groups whose units already had folders passed an existence check and were offered to an
    uploader with no access to either. **Existence is a property of the folder; being able to upload
    is a property of the folder AND the user**, and no amount of Folder Map reading answers the second.
  - **The probe tests UPLOAD, not read.** A PIC also in their unit's base group holds Read on the
    folder while their `*_UPL` group is still ungranted (2026-08-08 §5.8) — a visibility probe would
    call that ready and deliver the same 403. Read `Low` **arithmetically, never with `&`**: JS
    bitwise coerces to a SIGNED 32-bit int and Full Control returns `Low = "4294967295"`.
  - **`BulkUpload.tsx` is existence-gated only, deliberately.** It writes into **Documents**, where
    `UPL` is Read-only by design (2026-08-09), so an `AddListItems` probe there would empty that form
    for every PIC — accurate, but a client decision, not a side effect of fixing the upload form.
  - The gate is **DERIVED, never a `Ready` flag.** A flag can be ticked before it is true; this asks
    the same question the upload will. It needs no migration, and a segment-level flag could not
    express "three units reachable, one not".
  - **Gating the LEAF is the whole mechanism** — the tiers above are derived from surviving paths, so
    an unreachable segment empties itself and vanishes with no per-segment rule.
  - **An unreadable Folder Map or an inconclusive probe must offer EVERYTHING** (`unknown` ≠ empty, as
    with gotcha 11 and 10b; a short verdict array fails OPEN too). A transient error that empties every
    dropdown takes the form down site-wide; the upload-time checks stay as the backstop that makes the
    degraded path safe.
  - **Site admins are unaffected** and issue **no probes** — `privileged` users get the full cascade
    and no `validPaths`, so they can still see and test a segment they are building. The Structure
    Manager checklist is what tells them it is not live yet; the uploader no longer will.
  - **…but the admin's dropdown now SAYS so (2026-08-12, client's request, spec §9).** That exemption
    meant a half-built segment sat in an admin's picker looking identical to a live one — found live the
    hour Upstream Ops was created. The option reads `<label> — not fully set up yet`, and selecting it
    shows an amber banner naming the fix order (abbreviation → Folder Access rows → reconciliation) and
    that **uploaders cannot see it at all**.
    - **LABELLED, never disabled.** Greying it out is the obvious reading of the request and destroys
      the only reason for the exemption: the same admin had just used the unready segment to verify the
      Region → Estate/Mill cascade before any folder existed. Say it plainly, block nothing.
    - Readiness comes from `segmentProvisionState()` in `shared/segmentReadiness.ts` over the Folder Map
      rows **already loaded at mount** — a row's `Section` holds the segment's top folder, so it costs no
      request. `unknown` (unreadable list, or a blank `StagingFolder`) shows **nothing**: one transient
      error would otherwise brand every segment, live ones included. Empty ≠ unknown, again. Compared
      case-insensitively, because SharePoint folder names are.
    - Gated on `privileged` even though non-admins can never reach the state, so a later change to the
      filters cannot leak the label into an uploader's dropdown.
  - The empty state **names the right fix**: "isn't ready to receive uploads yet… run folder
    reconciliation — and if the unit has no folder at all, give it an abbreviation first" when the user
    IS authorised, vs the membership message when they are not. Reconciliation leads because the common
    case is an ungranted ACL. Shown only when BOTH filters report `known`.
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
    - **Slice B — adding a whole new segment — is BUILT and SITE-VERIFIED 2026-08-12.** Onboarded
      **Upstream Operations Malaysia** (`UPOPSMY`, `Region → Estate/Mill`, 15 terms) end to end: the
      depth check REFUSED 3 named tiers against the 2-deep set and wrote nothing; the real run derived
      `mode_upstream_operations_malaysia` and `SortOrder` **5** (max+1 of 1, 3, 4 — not count+1),
      created **8 columns across BOTH libraries** (`Region`/`RegionTid`, `EstateMill`/`EstateMillTid`),
      kept the label `Estate/Mill` while sanitizing the column to `EstateMill`, and ended on the
      checklist. Reconciliation then built the whole tree, and the upload form's cascade rendered
      `Region → Estate/Mill` with no code change. Spec
      `2026-08-12-add-segment-design.md`; rules in `shared/newSegment.ts` (pure, 33 tests), UI in
      `userAccess/components/SegmentCreator.tsx` as **tab 3** of the same web part. It writes the
      `mode` row (`ConfigType`, `Title`, `ModeLabel`, `Category`, `TermSetGuid`, `StagingFolder`,
      `SortOrder`, `Levels`) and the tier columns in BOTH libraries — and **nothing else**.
      - **THE TIER LIST NOW STARTS EMPTY (2026-08-15, client: *"it is always showing department and
        Unit this will confuse the client"*).** It was seeded `["Department","Unit"]`, which was not
        just confusing — it was a trap the depth check **cannot** catch, because that check compares
        the term set's depth to the tier **count**. A 2-deep Upstream Ops set matched the seed
        perfectly (right count, wrong names) and would create `Department`/`Unit` columns for a
        segment whose tiers are Region and Estate/Mill: nothing fails, and it surfaces later as a
        detail panel labelled in another segment's vocabulary. The admin NAMES the tiers, so a
        default is a name nobody chose. Examples stay in the hint (Head office → Department, Unit ·
        Upstream Ops → Region, Estate/Mill · I&T → one level), an empty row reads *"No levels yet"*
        so it cannot be mistaken for a failed load, and **Create is disabled with the reason beside
        it** — `validateNewSegment` already refuses zero tiers, but that message arrived after a
        click, and "no levels" is now the state every first-time user starts in.
      - **The admin NAMES the permissioned tiers.** A fixed `Department/Unit` prefix cannot onboard
        the eight remaining segments: Upstream Ops needs `Region → Estate/Mill`, I&T needs a SINGLE
        tier. So the form creates whatever columns those names imply.
      - **TERM-SET DEPTH MUST EQUAL THE PERMISSIONED TIER COUNT, and the form refuses on mismatch.**
        Reconciliation walks the TERM TREE and caps at that count, so 2 tiers against a 3-deep set
        ACLs **Regions** while the Group Map points at **Estates** — every unit grant on the wrong
        folder, no error, surfacing weeks later as "this person can see too much". The walk is
        bounded (400 requests, 8 at a time, stops one level past the tier count; GHO needs ~115).
        Cap or failure ⇒ depth **unknown** ⇒ **warn and allow**, never refuse. One unreadable branch
        makes the whole depth unknown on purpose: counting it as "no children" would report a
        SHALLOWER set than exists, and too-deep is the silent direction.
      - **`Title` and `SortOrder` are DERIVED** (`mode_<slug>`, max+1) — they are keys, and free text
        invites a typo whose only symptom is a segment that never appears.
      - **It ends on a CHECKLIST, not a success message** — abbreviations, groups, reconciliation,
        then verify. The abbreviation step is the one that silently creates nothing when missed.
      - `ensureColumn` now lives in `shared/spColumns.ts`, shared with slice A: the internal-name
        trick and `Options: 8` are load-bearing and invisible when wrong, so there is ONE copy.
      - **`sanitizeFolderSegment` REMOVES illegal characters rather than substituting a separator**,
        so `Estate/Mill` yields the column `EstateMill` and the key `mode_estatemill`. Pinned by
        test, because it decides derived column and key names.
      - Still OUT of scope, deliberately: creating the term set or its terms, abbreviation rows,
        groups/Group Map rows, reordering segments.
    - **DELETING a segment is BUILT (2026-08-14)**, on the same tab — now retitled **Segments** —
      spec `2026-08-14-delete-segment-design.md`, rules in `shared/segmentDeletion.ts` (pure, 23
      tests). The 2026-08-12 spec excluded it as having "no safe meaning"; this answers that rather
      than overruling it, because **Delete RETIRES**: it removes the `mode` row and the segment's
      Group Map rows, and touches **no document and no column**.
      - **Deleting the folders is a separate opt-in**, behind a typed confirmation of the segment's
        own label — added on the client's reasoning that once the files have been moved out with the
        migrator, removing the empty shell should be easy. It **recycles** rather than purges, so it
        is restorable for 93 days, and the dialog says so: that is what makes the option acceptable
        at all. ⚠ **"off by default" was true until 2026-08-26 and is now CONDITIONAL** — pre-ticked
        for an empty segment, unticked when it holds documents. See the 1.0.258.0 section below.
      - **An UNCOUNTABLE segment is not offered the folder delete.** Everywhere else in this
        codebase a failed read fails OPEN (gotchas 10b/11, the provisioned-segment filters), because
        the cost is a form that takes itself out of service. Here the cost is deleting an archive
        nobody could confirm was empty, so `canOfferFolderDelete` fails **CLOSED** — and the dialog
        names the failed read, because an unexplained missing checkbox reads as a broken page.
      - **Retire runs BEFORE the folders**, and a failed `mode`-row delete STOPS the run. The other
        order can leave a live segment with no folders, where every upload fails or silently
        re-creates a shell.
      - **Abbreviation rows are never deleted** — authored data with no other copy, and re-creating
        the segment without them renames every folder. **Columns are never deleted** either: the
        documents may have been MOVED elsewhere, which is exactly this workflow, and a deleted
        column takes its data with it and does NOT go to the recycle bin.
      - Group Map rows go WITH the segment: a row under a segment that no longer exists is a grant
        nobody can see or manage while reconciliation keeps re-applying it. Folder Map rows go only
        when the folders do — they are derivable, and rows without folders point at dead UniqueIds.
      - Reversible by design: re-creating the row with the same `Title`, `TermSetGuid`,
        `StagingFolder` and `Levels` restores the segment over folders that never moved.
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
  code change** — just `SubUnit`/`SubUnitTid` columns in BOTH libraries and one `Levels` entry
  before `Year`. Subunit data not yet supplied by the client.
  - **⚠ SUBUNIT TERMS ARE AUTHORED UNDER THE UNIT TERM, NOT IN A TERM SET OF THEIR OWN
    (2026-08-17, client's data).** Some units under a department have subunits and some have none,
    and the subunits are unit-SPECIFIC — which a flat standalone set cannot express: it offers every
    unit the same list, so a Finance Operations uploader would be shown Finance Land's subunits.
    Authoring under the unit makes optionality fall out of the data (no children ⇒ the tier is not
    offered) with nothing to configure and no "N/A" terms.
  - **This is what the architecture already assumes.** Reconciliation's walk stops at
    `permissionedDepth`, and its own comment says *"anything deeper in the term tree is a below-Unit
    tier: created on demand by the upload form, inheriting this folder's ACL, never provisioned
    here"* (`FolderManager.tsx` ~2046). So terms under a unit are **safe to author now**: no folders,
    no ACLs, no group-map warnings.
  - **BUILT (2026-08-17) — this section previously said "NOT YET BUILT", and that was stale.** The
    cascade exists in `Form.tsx` and `BulkUpload.tsx` (`tierPlan` → `childCache` → `decideTier`), and
    `StructureManager` has always stored it. **`termSet` ABSENT on a below-Unit tier IS the
    discriminator** — options come from the children of the term above, starting at the permissioned
    leaf. So SubUnit terms authored under each unit work with no code change. Not yet site-tested.
    - **The button the client asked for exists now.** Adding the tier previously meant *leaving a
      text box blank*, which is undiscoverable and whose wrong guess is silent: paste any other set's
      ID and every unit is offered every other unit's subunits, and it reads as working. The Folder
      levels screen now offers an explicit **"Under each Unit"** vs **"One shared list"** choice,
      defaulting to per-unit, and switching to per-unit **CLEARS** the GUID rather than ignoring it —
      a half-typed ID that survived the toggle would be saved the moment anyone switched back.
      `canAddTier` no longer lets the term-set verdict gate a per-unit tier (a stale `notfound` would
      disable Add over a field the admin can no longer see), and now REFUSES a shared-list tier with
      a blank ID — blank is not neutral there, it is the other kind of tier.
    - **`DocumentSearch` was the one real gap and is fixed.** A non-permissioned tier with no
      `termSet` fell through to a blank URL, so a SubUnit filter rendered with **no options** — and a
      filter that cannot be used reads as "this metadata was never captured". A cascading tier with a
      genuinely empty option list is now HIDDEN, matching the upload form; `[]` only, never
      `undefined`, because a failed read leaves it undefined and hiding on that would drop a filter
      whose options merely failed to load. Empty ≠ unknown, again.
    - **Optionality needs no configuration.** A unit with no child terms simply never shows the
      dropdown (`decideTier` → skip), and a tier that does not apply legitimately files one level
      shallower. But an **unresolved** option list blocks the upload naming the tier — treating
      unknown as "does not apply" would file a document a level too shallow, silently.
    - **A stale SubUnit selection cannot leak across units.** `tierSelections` keeps only selections
      still present in their tier's CURRENT options, so changing the Unit drops a SubUnit belonging to
      another unit and the upload is refused naming it — never written into the wrong unit's folder.
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
  anything); changing one **renames a live folder** on the next run.
  - **A RENAME AFTER THE GROUPS EXIST IS SAFE SINCE 1.0.162.0, and it was not before.** Renaming a
    code renames the FOLDER and touches no group, but `planBulkGroups` derives every group NAME from
    the abbreviation chain, so `Trace` to `TRC` made the planner look for `GHO_GS_TRC_UPLOADER`, find
    only `GHO_GS_Trace_UPLOADER`, and plan the unit as **unprovisioned**: a bulk run then created five
    more groups and thirteen more rows at the same term. `isDuplicateRow` cannot catch those (a
    different `GroupId` is a legitimately different row), so the unit ended with ten groups and
    twenty-six rows, all granted, for ever. Access stayed correct throughout; idempotence did not.
    - **It now matches on the TERM GUID**, which a rename never touches, using the Group Map rows the
      provisioner already loads for the dedupe. A unit is provisioned when a group holds a persona
      **role SET** at that term, whatever it is called.
    - **Never match a naming role alone**: `hou` carries `UPLHC` among its six and `UPLHC` names
      `pic_hc`, so that shortcut presents an approver group as the HC uploader group.
    - **The name check still runs FIRST**, because a run stopped part-way leaves a group holding only
      some of its rows: its role set matches no persona, and term-only matching would duplicate it.
    - **Rows whose group was deleted are ignored**, or the unit reads as done while nothing grants.
    - The preview and the log say `already there as <old name>`. **Renaming the groups is now optional
      tidy-up, not a required repair**: a group rename preserves its Id, so every mapping row survives,
      though `GroupName` on those rows goes stale and Folder Access labels them by it.
    - Still the better habit: **settle abbreviations before creating groups.**
  - **A CASE-ONLY change is deliberately NOT a rename** (`FolderManager.tsx` ~3233 compares
    lower-cased, because SharePoint sibling names are case-insensitive and `Trace` → `TRACE` would
    collide with itself). But `changedRows` compares case-SENSITIVELY, so such an edit saves, shows the
    `✎ renamed` warning, and then renames nothing — leaving the list and the folder disagreeing. Not
    harmful (folders resolve by UniqueId), and it reads as the rename being broken when it is not. Spec
  `2026-07-30-folder-abbreviation-naming-design.md`, client guide
  `docs/client/folder-abbreviations-guide.md`.
  - **The list now has an EDITOR — BUILT and SITE-VERIFIED 2026-08-12.** Spec
    `2026-08-12-term-abbreviation-page-design.md`; rules in `shared/abbreviationDraft.ts` (pure, with
    tests), UI in `folderManager/components/AbbreviationManager.tsx` as **tab 1** of the Folder
    Administration web part. The client will not touch a list view or paste a console script after
    handover, and this is the worst data to hand-edit: keyed by **term GUID**, obtainable only from a
    term-store properties panel one term at a time.
    - **THE PAGE NEVER SHOWS A GUID.** An admin picks a segment; the tree is walked and the keys
      resolved behind it. Showing the whole tree at once is also what makes the sibling check free.
    - **The sibling-collision check runs at TYPING time and BLOCKS save**, reusing `findCollisions`
      rather than copying the rule — the page must never save what reconciliation would refuse. Both
      rows are flagged (which one is wrong is not knowable), each message names the other term and
      their shared parent, and the comparison is case-insensitive and post-sanitize because that is
      what the folder will be called.
    - **No exemption flag, no guessed initials.** A per-level exemption would give reconciliation a
      second way to name a folder, invisible in the list — so the list would stop telling you what a
      folder is called. The client's own better idea, a **"Same as term name"** button (per row and per
      tier), stores the literal folder name, so skip-not-guess, sibling uniqueness and rename-on-change
      all keep applying to it for free. A guessed `GMB_PMB2C` looks authoritative and is wrong.
    - **Save writes rows ONLY** — inert and repeatable. Folders change only when someone runs Folder
      Reconciliation, so the order is always: fill in → save → reconcile. An existing row is MERGEd by
      item id; a duplicate row for one term GUID would leave reconciliation choosing between two codes.
    - **The row list SCROLLS at `58vh`** (client, 2026-08-18): 7 departments plus 60 units is 67 rows,
      so Save sat far below the fold and the warning banner scrolled out of sight while an admin
      worked. The warning, the collision banner and Save all stay OUTSIDE the box — only the rows
      move. Capped unconditionally, unlike Folder Access's group list, because these rows hold text
      boxes and buttons and nothing absolutely positioned for a scroll container to clip.
    - Below-Unit tiers are **listed but read-only** (they name folders from the term label), so nobody
      goes hunting for Year.
    - **Blanking a code after its folder exists, then setting a new one, shows NO rename warning** —
      the warning compares against the STORED code (now empty), so it reads as a first-time fill, while
      the Folder Map row still points at the old folder and the next run renames it. Seen live
      2026-08-12 (`BE` → blank → `BNG` → `✎ renamed BE → BNG` in both libraries). The outcome is
      right; only the warning is silent.
- **FOLDER ADMINISTRATION IS ONE TAB BAR (2026-08-12)** — same spec §6, same web part as before
  (`Folder Manager`, GUID `02007994-…`, **retitled** `Folder Administration`; the id is unchanged so
  pages already hosting it keep working). Five tabs, in the order the work
  happens — reordered 2026-08-15 on the client's request so the segment leads: **New segment → Term
  Abbreviations → Folder levels → Move existing folders → Folder Reconciliation**. Nothing else on
  the page can be done until a segment exists, so last was the one position that could not be right;
  it had been put there as "the rarest", which reads the tab bar as a frequency ranking rather than a
  sequence. The label says **New segment** though the tab also DELETES one — retiring has its own
  guided flow, and naming the tab for its destructive half puts "delete" in front of an admin who is
  almost always there to add. **The `newsegment` deep-link slug is UNCHANGED** (a public name once
  shipped; this is a label and a position, not a rename). It still **opens on Term Abbreviations, not
  on the first tab** — position states the sequence, the default states the likely job: segments
  already exist on a site being administered, so landing every visit on a creation form invites a
  duplicate, and reconciliation is what *fails* when the codes are missing.
  - The three `Folder Structure` tabs are **MOUNTED from `userAccess/components`, never copied** —
    `StructureManager` rewrites `Levels` and creates columns in both libraries, `SegmentCreator` writes
    a `mode` row, and drift in either surfaces weeks later as a wrong column or a half-built segment.
    **The `Folder Structure` web part stays registered** — it may already be on a page.
  - **`Staging` and `Documents` are GONE** (client: *"I am honestly not using it"*). Their manual
    folder tree — hand-create, rename, assign per-folder permissions — is covered by reconciliation and
    the Folder Access page. Lost with them: browsing a folder to see who actually holds access, and
    hand-creating a folder (arguably worth losing — reconciliation does not know about a hand-made one).
  - **The tree's CODE stays, unreachable**, so its deletion is its own reviewable change. Two traps
    came with that: its controls were gated on `tab !== "Reconciliation"`, which now also matches all
    four new tabs (it would have drawn a second Refresh/Update bar over screens that each have their
    own Save — now `treeTab`); and its mount-time crawl is **gated off**, since it was dozens of
    requests per page load for a view nobody can open, ending in a permissions toast about a missing tab.
  - A tab switch with unsaved changes is **REFUSED, not confirmed** — Save is a few pixels away, and a
    "discard?" prompt puts losing the work one click behind ordinary-looking navigation.
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
LegallyPrivileged          <- Yes/No, written as the STRING "true"/"false"; shown only for the levels
                              (⚠ THE CONFIG ROW HOLDS LEVEL NAMES, NOT TERM GUIDS, since 2026-08-19 —
                              on the client site it held `87f8481b-…` because the check compared the
                              dropdown's RAW VALUE, which is the term ID. HC routing had always resolved
                              the label first via `confidentialityLabel`; this had not, so the two rules
                              wanted opposite shapes in config and nothing on screen could explain it.
                              A stray `term_highlyConfidential` GUID row on that site is read by NO code.)
                              LISTED by the `legallyPrivilegedFor` DMS Config row (blank = never
                              offered). A LIST since 2026-08-19 (client: the tick must show for Highly
                              Confidential too) — `Confidential;Highly Confidential`, semicolon OR comma
                              separated, matched trimmed and case-insensitively. It was a single `===`
                              compare, so naming HC would have REMOVED the tick from Confidential: the
                              setting could express "one level" and nothing else. Rule lives in
                              `shared/legalPrivilege.ts` (pure, 10 tests) and is shared by FOUR call
                              sites — the write derivation and the render guard in BOTH upload web parts,
                              which were hand-maintained copies of each other deciding a legal marker.
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
- **Auto-route MUST carry the trigger condition `@equals(triggerOutputs()?['body/{IsFolder}'], false)`**
  (added 2026-08-13). Without it every folder the upload form ensure-creates is approved by the
  folder-approval flow, fires Auto-route, and fails `Copy file` with `NotFound` — and one such run had
  already copied a file into `Documents` while it was still **Waiting for Approval**. Approval was
  being bypassed whenever the flow won the race against the upload. **The polarity is the whole
  thing:** `false` here (files only), `true` on the folder-approval flow. Set `true` here and no file
  is ever routed, silently, because a flow that never fires leaves no run history. Both mistakes have
  now been made — see spec `2026-08-08-auto-route-flow-and-draft-isolation.md` §4.

> ⚠ **Power Automate header keys must NOT include the colon.** The key box wants `Accept`, not
> `Accept:`. With the colon the header does not exist, and the symptoms look unrelated to each
> other: responses silently come back `odata=verbose` (so every `body('X')?['Field']` is null →
> *"Not well formatted JSON stream"*), `validateUpdateListItem` reports `HasException: false`
> and changes nothing, and a MERGE goes as a plain POST (*"The parameter AuthorId does not exist
> in method GetById"*). This cost two hours on 2026-08-08 and produced a false conclusion that
> `Author` was unwritable. Read the action's raw **Inputs** — the colon is visible there and
> nowhere else. And **never** conclude a field is unwritable from a clean response: re-read the item.

## ⚠⚠ THE "DRAFT ISOLATION TURNED OFF" SECTION BELOW DOES NOT DESCRIBE THE LIVE SITE (checked 2026-08-22)
**Draft Item Security reads *"Only users who can approve items (and the author of the item)"* on BOTH
`Approval Document` AND `HC Approval Document`.** The setting was never applied, or was applied and
reverted. **The 2026-08-08 design is what is actually running**, everywhere.

Everything the section below asserts as current is therefore false today, and three of its claims were
acted on in one session before anyone opened Versioning Settings:
- **BOTH folder-approval flows must STAY ON.** They are not dead weight; without them a below-Unit
  folder created by one PIC is invisible to the next, who cannot reach their own file inside it and
  cannot self-fix (it needs `ApproveItems`). Bulk Upload ensure-creates those folders too.
- **A PIC does NOT see their colleagues' pending or rejected files.** So `DELS` is once again
  equivalent to "delete your own", and the 2026-08-15 design's silence on that is correct after all.
- **`docs/client/document-visibility-within-a-unit.md` had been rewritten to say drafts are NOT
  private — corrected 2026-08-22.** That is the page shown in meetings, so it is the worst place for
  this to have been wrong.

**How it happened, and the rule:** a change was agreed, written up in detail, and recorded here as
done. Nothing in the repo can observe a SharePoint setting, so "we agreed it" became "it is true".
**Before acting on any claim in this file about a SharePoint SETTING, open the setting.** Caught only
because Clarence questioned an instruction that rested on it. See `feedback-verify-live-state-before-changing`.

## ⚠ DRAFT ISOLATION WAS TURNED OFF ON 2026-08-19, AT THE CLIENT'S REQUEST — NOT ACTUALLY APPLIED, SEE ABOVE
**Draft Item Security was set to *"Any user who can read items"*.** This reverses the approval-library
half of the 2026-08-08 design below, which is otherwise still the record of how the mechanism works —
read it for the reasoning, not for the current state.

> ⚠ **"ALL libraries" WAS WRONG — `HC Approval Document` STILL READS *"Only users who can approve items
> (and the author)"*, verified on site 2026-08-22.** So the two verticals behave differently today: in
> the normal approval library every uploader in a unit sees every pending file; in the HC one a PIC sees
> only their own. Consequences that follow directly:
> - **`HC folder approval` MUST STAY ON.** A below-Unit folder created by one HC-cleared PIC arrives
>   Pending and is invisible to the next, who then cannot reach their own file inside it — and cannot
>   self-fix, because that needs `ApproveItems`. This applies to Bulk Upload too, which ensure-creates
>   the same folders. **The instruction to turn both folder-approval flows off is correct for the
>   NORMAL pair only.**
> - For HC this is arguably the better setting, so it is flagged rather than "fixed" — but it should be
>   a client decision, not an accident of which libraries got visited on 2026-08-19.
> - Caught only because Clarence challenged an instruction of mine that rested on this sentence.
>   **A site-wide claim in this file is a claim about the libraries someone actually opened.**

- **Every uploader in a unit now sees every other uploader's PENDING and REJECTED files.** The
  peer-isolation behaviour verified on site the same afternoon (a peer PIC saw an empty folder) is
  deliberately gone.
- **The boundary the client actually wants is the HC one**, in their words: `GHO_GF_TAX_UPLOADER` must
  not see HC files or reach the HC libraries; only `GHO_GF_TAX_UPLOADER_HIGHLY_CONFIDENTIAL` can. That
  boundary is untouched — it is folder ACLs and `LIBRARY_ROLES`, not draft security — and was verified
  on 2026-08-19.
- ~~**BOTH folder-approval flows are now unnecessary and are being turned OFF.**~~ **❌ FALSE — DO NOT
  ACT ON THIS. BOTH FLOWS MUST STAY ON.** Draft Item Security was never actually changed (verified
  2026-08-22, and again 2026-08-23): it reads *"Only users who can approve items (and the author)"* on
  BOTH approval libraries, so a pending folder created by one PIC IS invisible to the next, who then
  cannot reach their own file inside it and cannot self-fix (that needs `ApproveItems`). These flows
  are what prevent it.
  - **This line has now been acted on TWICE by mistake** — once on 2026-08-19 and again on 2026-08-23,
    when it was quoted as a reason to switch `HC folder approval` off, and Clarence caught it both
    times. It is struck through rather than deleted because the reasoning below it is the record of
    what the reversal WOULD have implied, and a deleted line cannot be recognised when someone finds
    the same argument again.
  - The whole section it sits in describes a change that **was agreed and never applied**. Read the
    correction at the top of this file before anything here.
- **A PIC's `DELS` now covers ANY pending file in their unit, not only their own.** The 2026-08-15
  design wrote no "own files only" rule because draft security made one unnecessary — *"delete what you
  can see" IS "delete your own"*. The permission is unchanged; the visibility that made it equivalent is
  not. Still confined to the **approval library** and the **unit**; approved documents are untouched, so
  deleting or sharing one remains a request the Head of Unit decides. Marked superseded in the
  2026-08-15 and 2026-08-17 specs, where the same sentence was restated (`DELSHC` inherits it too).
- **Exposure is still bounded by the folder ACLs.** Only roles in `LIBRARY_ROLES.Staging` can open an
  approval library at all, so this widens visibility to the unit's uploaders and approver — nobody
  else, and nobody outside the unit.
- **Client-facing doc updated**: `docs/client/document-visibility-within-a-unit.md` no longer says
  drafts are private. That sentence was the one that landed in meetings, so it must not survive.

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
| Head of Department | `DEPTVIEW` | department | view **every unit under the dept**, incl. HC — **nothing else** | **none** — does NOT approve |
| Head of Unit | `APR`, `DEL`, `DELS`, `SHARE`, `UPL`, `DELSHC` (HC-cleared variant `hou_hc`: `UPLHC` for `UPL`, 2026-08-24) | unit | view own path, delete + share | **approve + view every file in their unit (incl. HC), delete pending/rejected, upload — HC filing needs the `hou_hc` group** |
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
- **RECONCILIATION LOCKS THE ADMIN PAGES ITSELF (2026-08-17, client: *"can we not auto restrict the
  admin pages?"*).** Spec `2026-08-17-admin-page-lockdown-design.md`. Closes the last blocker before
  migration. `CRS_SITE_MEMBERS` holds Read on the WEB — it must, or a folder-only user is denied on
  Home — and Site Pages inherits, so **every uploader could open every admin page**.
  - **The existing page pass could not fix it, and that is the point:** it iterates pages that HAVE
    Group Map rows, and an `adminOnly` page has `roles: []` so no row can exist. **Third instance of
    one structural gap** — the mechanism is driven by grant rows and the thing needing protection has
    none (the others: HC libraries inheriting, #10; the inverted site-entry grant, #7). All three are
    fixed the same way: **assert the required state every run** instead of deriving it from rows.
  - **A SEPARATE pass with its own Site Pages read**, not a restructure of the row pass — that block
    builds everything inside `if (pageRows.length > 0)`, and reshaping the most site-verified code in
    the file days before a client migration is the wrong risk. One extra GET per run.
  - **It asserts in BOTH directions.** A page can be **unique and still exposed** — inheritance
    broken with a group granted Read by hand — so checking `HasUniqueRoleAssignments` alone would
    call it locked. Every non-Owners assignment is stripped, and **each removal is logged with the
    principal named**: silently removing a deliberate grant is worse than not removing it, because
    the admin goes on believing it is there.
  - **The row pass now REFUSES a row targeting an `adminOnly` page**, so the two cannot fight —
    otherwise one grants Read and the other strips it, every run, for ever. Order then stops
    mattering, which is a correctness property rather than a sequencing convention.
  - **Fails CLOSED per page, OPEN on the read.** Unreadable Site Pages ⇒ lock nothing; an unreadable
    ACL ⇒ skip **that** page unchanged (stripping what you could not read removes invisible grants);
    a failed break ⇒ reported **still open** with the status, never as locked.
  - **THE COVERAGE TEST FOUND FIVE ADMIN PAGES THE POLICY NEVER MATCHED**, and none by inspection:
    `Group-Management` (`group.?manager` does not match "Management"), `Site-Access`, `Page-Access`
    and `CRS-Audit-Log` (all fell to `DEFAULT_POLICY` = UPL/APR/DELS), and **`Approval-Library-Access`,
    which matched `/approv/i`** — the screen that grants library permissions, classified as an
    approver page. Harmless while restriction was a manual step nobody did; now the rule decides what
    gets locked, **an unmatched name is a page left open silently, for ever**. Pattern is now
    `/folder|group.?manag|config|setting|mapping|admin|access|audit/i`, pinned by a test listing every
    page the runbook creates on both sides of the line.
  - Matched by FILE NAME, so a client page called `Configuration.aspx` would be locked. Accepted:
    every lock is logged by name. An exact allow-list was rejected — this client renames everything
    at import, and the list would stop matching silently.
- **PAGE ACCESS IS DERIVED FROM FOLDER ROLES SINCE 1.0.171.0 — nothing had ever created a Page row.**
  Spec `2026-08-19-derived-page-access-design.md`; rules in `shared/pageGrants.ts` (pure, 34 tests).
  Reconciliation granted a page only to groups holding a **`Scope = Page`** row, and **only
  `PageAccess.tsx` writes one**: bulk provisioning writes FOLDER rows, the guided flows never mention
  page access, Group Management does not offer it. So the rows exist only where a person made them —
  and **a unit or segment added later gets folders, groups and no page**, with nothing reporting it,
  surfacing months on as one person's AccessDenied. `My-Submissions.aspx` and `Requests.aspx` had no
  rows at all on the rehearsal site and were therefore still INHERITING: openable by anyone who could
  open the site.
  - **⚠ THIS SECTION FIRST CLAIMED "8 of 120 groups could open the upload form", AND THAT WAS WRONG.**
    It came from a **progress panel read at 0%** — eight lines of a list of ~120 — taken as the
    finished result. **A mid-run progress panel is not a result.** Kept here because the premise was
    verified only after the code was written.
  - **FOURTH INSTANCE OF ONE STRUCTURAL GAP** — the mechanism is driven by grant ROWS and the thing
    needing the grant has none (the others: the inverted site-entry library grant, the HC libraries
    inheriting, the admin pages readable by uploaders). All four fixed the same way: **assert the
    required state every run.**
  - **The loop is driven by the PAGES now, not the rows**, and that inversion IS the fix. Any group
    holding a qualifying role at ANY tier in ANY segment gets the page: a page is not scoped to a
    unit and cannot be. What they may actually FILE is still decided by the folder ACL, which the
    upload form probes — so a group listed here that cannot write sees an empty cascade.
  - **⚠ DERIVATION FIRES ONLY ON A PAGE WHOSE NAME MATCHED A RULE, and that gate is the whole
    safety.** `policyForPage` answers identically whether `/upload/i` matched or the name fell to
    `DEFAULT_POLICY` (`UPL, APR, DELS`). Granting at page scope requires BREAKING the page's
    inheritance first, after which only the listed groups can open it — so deriving from that default
    would break the site home page, `CollabHome.aspx` and every page the client authored, and lock
    out every role not in the list: SDG Employee, Head of Department and C-Level hold none of them.
    Reconciliation would take the site away from most of its users on a run reporting success. Hence
    `pageMatchedRule` + `derivedRolesForPage`, both reading the existing `RULES` array so there is no
    second list of page names to drift.
  - **Hand-made Page rows SURVIVE and merge**, deduped on `GroupId` — a Page row is the only way to
    grant a page the policy does not imply, and it still works on a page that derives nothing.
  - **The full ACL is ASSERTED: a group with no qualifying role is REMOVED, named in the log.** So
    deleting a group's mapping rows now revokes its page on the next run. ⚠ **Safe at page scope and
    NOWHERE ELSE** — a Site Pages item is a LEAF, while a library or folder root also carries
    SharePoint's automatic **Limited Access** entry for every principal granted further down (~308 of
    them on the approval library), so the same rule there would strip them all and take every group's
    folder access away on a run that reported success. Do not lift `groupsToRemove` above a leaf.
  - **Only site Owners is protected.** The site-entry group is deliberately NOT: it holds Read on the
    WEB, so on an inheriting page it is there by inheritance and vanishes when inheritance breaks — a
    page-SCOPE assignment for it can only have been added by hand, and would hand every plain member
    the upload form. **User principals are never removal candidates**, since individual grants were
    rejected as a mechanism in `2026-08-14-per-person-access-removal-design.md`.
  - **A restricted page with nobody now SAYS SO** (`⚠ restricted, but no group holds UPL or APR`) — a
    warning, not a refusal: an empty set is correct on a site with no groups, and refusing would
    leave the page INHERITING, i.e. readable by every site member.
  - **Fails OPEN on the read, CLOSED per page on the write.** An unreadable Group Map derives
    nothing; an unreadable page ACL removes nothing but still grants. The Group Map read is `$top=5000`
    and does NOT lift the 5,000-item threshold — past ~5 segments it must be paged, or a truncated
    read would under-derive and then strip grants that should stay.
- **⚠ A SCOPED RECONCILIATION RUN PRUNED AN UNCOVERED SEGMENT'S FOLDER MAP ROWS — LIVE DATA LOSS,
  2026-08-19, fixed 1.0.179.0.** Spec `2026-08-19-selective-reconciliation-design.md` §5. The Folder Map
  prune and the term-GUID orphan repair decide a row is dead by asking *"is its term among the terms
  THIS RUN enumerated"*. Scope the run to one segment (#18) and **every row of every uncovered segment
  answers no**. An MHO-only run deleted all **67 of GHO's** rows; every GHO uploader was then refused
  with *"your unit isn't ready to receive uploads yet — run folder reconciliation"* — the message
  reconciliation itself tells them to act on, produced by reconciliation. Second time in one day that a
  Folder Map defect presented as that exact sentence.
  - **THE MIRROR OF `ALWAYS_FULL_PASSES`, and the direction nobody wrote down was the destructive one.**
    Site-wide passes must never be NARROWED by scope; orphan passes must never be WIDENED beyond it. The
    spec stated the first at length and never considered the second.
  - **Skipped ENTIRELY on a scoped run, never filtered to the covered segments.** Filtering needs every
    row to declare its segment reliably, and a row whose segment could not be determined would then be
    deleted by the rule meant to protect it. Orphan cleanup is maintenance — a full run does it, and
    doing nothing is always recoverable.
  - `coversEverySegment` (`shared/reconScope.ts`, pure) fails **CLOSED**: empty, refused or unreadable
    scope ⇒ false, so an unknown scope never licenses a deletion. Regression tests pinned.
  - **Recovery is a FULL run** — rows are derivable from the term tree and the folders, which the prune
    never touches. It deletes list rows only, and that is what made this survivable.
  - **⚠ THE FIX THEN OVER-GATED, and #19's quarantine sat inside the guard for a day (fixed 1.0.186.0).**
    Gating the whole `else` block took the STRAY-FOLDER pass with it, so a scoped run reported no
    strays at all — and scoped runs are now the normal way to work. The two passes ask different
    questions: the prune asks *"was this row's term in THIS RUN's targets"*, which is false for every
    uncovered segment; the quarantine pass **descends FROM `targets`**, so it can only ever see inside
    the segments that were walked. **Scope-safety is a property of the question, not of the block a
    pass happens to sit in.**
  - **The clean-run guard did not help**, and could not: the prune runs BEFORE the abbreviation pass, so
    the `✗ ... 67 orphaned` line that revealed the problem did not exist yet when the guard was checked.
- **A SEGMENT C-LEVEL NOW DELETES AND SHARES INSIDE THEIR SEGMENT (1.0.177.0, client 2026-08-19:
  *"allow Clevel to delete and share"*).** Spec `2026-08-19-clevel-segment-fan-down-design.md`; rule
  is `segmentFanRoles()` / `fansFromSegmentTier()` in `groupMapModel.ts`, **derived from `PERSONAS`**.
  - **`clevel_segment` is `SEGVIEW + DEL + SHARE` and only `SEGVIEW` used to fan down**, so DEL and
    SHARE landed on the SEGMENT folder alone — and every folder below has broken inheritance, so
    **they reached nothing.** The persona picker said *"view, delete + share one segment"* and granted
    view. Not a regression; true since the persona was written, and surfaced only by reading a run log
    where the refusal printed twice on every folder and read as a safety feature.
  - **The exemption argument already existed and simply had not been extended.** The fan-out gate
    refuses a segment/department-tier row because it cannot be told apart BY TIER from a
    pre-2026-07-29 leftover `MEMBER` row. `SEGVIEW` was exempt because it granted nothing on any site
    until 2026-08-07 — *the role name is the consent*. `DEL` and `SHARE` at the SEGMENT tier are in
    the same position: neither was ever auto-created there, and only the two C-Level personas put
    them there.
  - **⚠ THE SAFETY IS THE TIER, NOT THE ROLE.** `DEL` and `SHARE` are also held by `hou`, at the UNIT
    tier, which is never fanned. The caller checks the row's tier BEFORE consulting the list —
    using it on a DEPARTMENT-tier row re-opens the exact hole the gate exists to close. A test pins
    that no upload or approve role is ever in the list: a C-Level has no Staging access, so fanning
    one would invent it.
  - **DERIVED, never listed.** A literal would be a second definition of a C-Level's powers, free to
    drift from `PERSONAS` — and the drifting copy would be the one reconciliation reads and no screen
    shows. An unrecognised `Role` string answers **false** and refuses: a missing grant is visible in
    the log, a wrongly granted delete right is not.
  - **Say the widening plainly to the client:** a segment C-Level can now delete or share ANY approved
    document anywhere in their segment with nobody's approval. Unchanged: they hold nothing in either
    approval library, so unapproved drafts stay out of reach.
  - **Migration = re-run reconciliation.** No row, group or schema change.
- **RECONCILIATION SKIPS A GRANT ALREADY IN PLACE SINCE 1.0.177.0 — before that it re-issued EVERY
  folder grant on EVERY run.** Spec `2026-08-19-skip-existing-grants-design.md`; rule in
  `shared/grantSkip.ts` (pure, 10 tests). `addRoleAssignment` was called unconditionally for every
  group on every folder in every library, and SharePoint takes a duplicate as a **no-op** — so nothing
  ever failed. It cost a round trip and a throttle tick each time and logged a line saying the grant
  had been made. For one segment that is ~980 pointless writes, which is most of why re-running a
  provisioned segment cost as long as provisioning it, and the real content of the client's #18 ask
  (*"I added one unit, why is this an hour"*).
  - **THE ASYMMETRY IS WHY IT SURVIVED:** the ancestor-browse grant twenty lines below in the SAME
    loop already skipped what was in place. One loop checked its neighbour's work and the other did
    not, and **the only symptom was time** — which reads as SharePoint being slow, not as a defect.
  - **The ACL is read only where it can pay.** Two of the three folder branches RESET the ACL
    themselves (created, or inheritance broken with `copyRoleAssignments=false`), leaving only site
    Owners — so a read there adds a request per folder to the FIRST run, the expensive one, to save
    nothing. The third branch (`already locked, skipped`) is the whole steady state of a re-run, and
    there one read replaces up to five writes per folder per library.
  - **⚠ UNKNOWN GRANTS.** `needsGrant(undefined, …)` is true. A failed read says nothing about the
    folder, and guessing "already granted" leaves a group silently ungranted on a run reporting
    success — an uploader who cannot upload after the admin did the one thing they would be told to do.
    `getRoleAssignments` folds a non-OK status into `[]` itself, which fails the same safe direction.
    **`[]` is read-and-empty; `undefined` is not-read.** Empty ≠ unknown, where the cost is a missing
    grant.
  - **The match is EXACT on principal AND role definition, never the principal alone.** A group can
    hold two bindings on one folder (a HoD holds `DEPTVIEW` on a department folder plus the ancestor
    browse `Read`), and the reader keeps only the first binding per principal — so a principal holding
    a different level re-grants. One wasted no-op beats a missing permission.
  - **⚠ A SKIPPED GRANT STILL COUNTS AS GRANTED for the ancestor-browse pass.** `grantedPids` drives
    the corridor from the library root down to a member's folder. A skipped group still HOLDS the role,
    so it still needs the corridor; dropping it would render the library empty for its members — the
    2026-08-04 regression by a new route.
  - **The first run of a segment is unchanged in cost, deliberately.** MHO's 71 minutes was honest
    work: 2,441 folders and grants that did not exist. **A known-expensive mechanism being present in
    the code is not evidence it is running** — the grid was the obvious suspect and was already `off`.
    Check the config row and the run's opening lines before optimising against it.
  - Reported **once per FOLDER** (`N group grant(s) already correct, skipped`), never per group: on a
    settled site that is every group on every folder, and a line each buries the run's real findings.
  - **⚠ IT NEEDS ITS OWN READER — `getAllRoleBindings`, NOT `getRoleAssignments`.** SharePoint collects
    every permission level a principal holds on a folder into **ONE** role assignment with several
    `RoleDefinitionBindings`, and the display reader keeps only the **first** (right for its callers,
    which want one level to show). So an approver group holding CRS Approve + CRS Delete + CRS Upload
    read back as holding one, and the rest were re-granted every run. **Measured: 71m → 19m50s with the
    wrong reader, 868 assignments still re-issued** — per unit folder, five re-issued and three skipped,
    the three being whichever level came back first. The spec called this "a rare case"; it is every
    group with more than one role, i.e. every approver group. The new reader **filters nothing** (the
    display one drops `RoleTypeKind` 1 and 7): a dropped binding can only cause a needed skip to be
    MISSED, while an extra one can only cause a wasted write.
  - **A grant made during the run is pushed into the in-memory list**, so a second role resolving to the
    SAME level is neither re-written nor re-printed. `APR`/`UPL` are both Read on Documents and
    `DELS`/`DELSHC` are both CRS Delete on the approval library, so the log used to show one group
    against one level **twice on one folder** — which reads as a double grant (client, 2026-08-19:
    *"its basically confusing the client"*).
- **THE UPLOAD FORM PAGE ACCEPTS `APR` AS WELL AS `UPL` (2026-08-17, client: *"client wants HOU to
  be able to get into upload form to upload, basically apr can upload"*).** `pageAccessPolicy.ts`'s
  `/upload/i` rule listed `["UPL"]`, so an approver got **AccessDenied on the upload form** — found
  live with a real second account, while the run log showed the page granted to the UPL group alone.
  - **The `hou` persona has carried `UPL` since the 2026-08-15 correction, and that is not enough.**
    A Head of Unit group mapped BEFORE that correction holds only `APR` and `DELS` rows, so keying
    the page on `UPL` locked out every HoU already provisioned. Listing `APR` fixes both that group
    and a pure approver group with **no data migration** — no new Group Map row, no re-run.
  - **Widening this list cannot grant anyone a new place to upload.** The page grant opens the FORM;
    the folder ACL decides what can be filed, because the form probes `AddListItems` on the
    destination and shows the "not ready" empty state where that fails. A role listed here that
    cannot write sees an empty cascade — the safe direction.
  - **The asymmetry with `ApprovalDocument.aspx` is deliberate and pinned by test:** approvers reach
    the upload form, uploaders must NEVER reach the approval queue. Mirroring them would let every
    PIC approve their own documents. View-only roles stay excluded from both.
- **Folder Access is persona-only.** The role chips are gone: picking a persona applies its roles.
  A library toggle above the group field splits the personas — Documents (C-Level ×2, HoD, SDG
  Employee) vs Approval Document (HoU, PIC). That grouping is **derived** via
  `personaTouchesStaging()`, never hand-listed: filing C-Level under the approval library would
  present the widest accidental grant in the system as normal.
- **⚠ GROUP NAMES CHANGED 2026-08-18 — `_EMPLOYEE` AND `_HOD` ARE NEW SUFFIXES.** Spec
  `2026-08-18-group-creation-and-bulk-provisioning-design.md`; live state and five open defects in
  **`docs/2026-08-18-bulk-groups-handoff.md`**.
  - **`MEMBER` NO LONGER CARRIES NO SUFFIX.** It was a bare `GHO_GF_TAX`, and `roleFromGroupName` **falls
    through to MEMBER** for anything unrecognised — so a bare name and a typo were the same answer, which
    is how a mistyped approver group reads as view-only. Now `GHO_GF_TAX_EMPLOYEE`. Do NOT restore the
    bare form: making that fallback loud later is only safe *because* every generated name now carries a
    suffix (deliberately not done yet — it would reclassify groups on existing sites).
  - **`DEPTVIEW` is `_HOD`**, the client's spelling. **HOU is Head of UNIT** — a department group must
    never borrow it. `_DEPARTMENT_VIEWER` and `_DEPTVIEW` still parse, so nothing already named is
    stranded; only new names use `_HOD`.
  - **`suffixForRole` no longer short-circuits MEMBER**, and **`roleFromGroupName` now recognises the
    literal `GLOBAL`** — it used to fall through to MEMBER, so the WIDEST grant in the model parsed back
    as the narrowest.
  - **Every persona DECLARES its `namingRole`; never derive it from `roles[0]`.** `employee_hc` is
    `["MEMBER","MEMBERHC"]`, so first-role naming would present an HC-cleared viewer group as having no
    clearance. Pinned by a round-trip test across every persona.
  - **Group Management asks for a PERSONA, and the name is derived and read-only** (free text behind an
    advanced toggle, for `CRS_SITE_MEMBERS` and genuine one-offs). **Creating a group also writes its
    Group Map rows**, which takes the name-parse off the critical path entirely.
  - **Bulk provisioning exists** (`BulkGroupProvisioner`, on Group Management and in the flow's group
    step): 5 personas per unit + `_HOD` per department + `_SEGVIEW` per segment, planned from the
    abbreviation rows.
  - **THE RUN IS IDEMPOTENT FOR ROWS AS WELL AS GROUPS SINCE 1.0.151.0** (`splitPlannedRows` in
    `shared/bulkGroups.ts`), and the asymmetry is what made the bug invisible: an existing group title
    was always mapped rather than re-created, so a second press logged `= already existed (mapping
    only)` on every line **while writing every mapping underneath it again** — 642 rows on the
    rehearsal site. Nothing on any screen shows a row twice, so the only symptom is a count nobody has
    a reference for. It reuses `isDuplicateRow`, never a second definition of "the same mapping".
    - **The partition is per ROW.** A run stopped part-way leaves a group made with one of its rows
      written; skipping the whole group as done would strand it for good. Pressing Run again is the
      only resume this has (defect 2), so completing the gaps IS the recovery path.
    - **The Group Map read is PAGED and fails CLOSED.** `$top` caps a page and does not lift the
      5,000-item threshold, and a truncated read reports unseen rows as absent — the same duplication
      by another route. An unreadable list is `undefined`, never `[]`, and HOLDS the run with the
      reason on screen: this codebase fails open nearly everywhere because the cost is a form out of
      service for a minute; here it is hundreds of duplicate rows nobody would find.
  - **NAVIGATING AWAY FROM A RUN USED TO KILL IT SILENTLY** (fixed 1.0.152.0). The run lives in
    component state with no resume, so a step change unmounts it mid-way — the likely cause of 302 of
    308 groups on the rehearsal site. `onBusyChange` reports it up and `FolderAdmin` holds the rail,
    Back, Next, Finish, the mode switch and the back band, with the reason beside the greyed buttons;
    the `beforeunload` guard sits in the component so the standalone page is covered too.
    - **This is the ONLY padlock in the guided runner**, against its own rule that the flow must not
      stop an admin working — because here navigation destroys work in flight, exactly as a tab switch
      with unsaved abbreviations does. Temporary, self-clearing, and paired with a **Stop** button so
      nobody is held longer than they choose. Stop lands BETWEEN groups, and a stopped run reports as
      a warning however cleanly it stopped.
    - **Stop is only safe because the run is idempotent** — pressing Run again finishes the gaps.
  - **FOLDER ACCESS SHOWS `Tier 1` / `Tier 2`, NOT ONE BARE TERM** (1.0.153.0, client's request).
    A row stores only the LEAF term GUID, so a unit row never said which department it was in and a
    department row looked identical to a unit one — while the difference between them is the
    difference between granting one unit and granting all of them. `Tax`, `Legal` and `PM` each exist
    under several departments, so the leaf label is ambiguous, not merely terse.
    - **Ancestry is DERIVED from a walk, never stored.** A term can be re-parented; a copy in the
      list would then describe a shape that no longer exists. `shared/termChains.ts` (pure) builds
      `guid → [labels]`; `GroupMapBuilder` walks each segment **once**, which also replaced ~130
      per-term reads with ~8.
    - **The walk is bounded by the PERMISSIONED depth**, and `parseLevels` keeps
      `"permissioned": false` entries — so the filter is what stops it descending into every SubUnit
      on the site (they are authored under each unit).
    - **An unreadable branch abandons the whole segment**, which reads `Tier not known`. Reporting it
      as a term with no children would present a unit AS a department. A chain of one legitimately
      means "on a department"; unresolved means "we do not know" — rendered differently on purpose.
    - The CSV split with it, from the SAME function as the screen.
  - The **group list scrolls** at `60vh`, but only while every group is COLLAPSED — an expanded
    group's people picker is absolutely positioned, and a scroll container clips it for any group near
    the bottom, trading a long page for a control that silently cannot be used.
  - **The rail re-reads its facts when a run ENDS** (`onRunBusyChange` → `reload`). `groupsExist` was
    read when the segment was picked, so the step still said *To do* after 300 groups were created. A
    mount-time read reflects nothing a step below it has since written — the same trap as the segment
    picker's Refresh button.
  - **A SCREEN THAT READS A LIST AT MOUNT LIES ABOUT ANY RUN BESIDE IT.** Third instance in this one
    feature: the rail's tick, `GroupManager`'s mapping badges (every group read `not mapped` right
    after a run wrote 790 rows — found on site 2026-08-18), and the segment picker's own Refresh
    button before them. The badges were the dangerous one: an under-reported tick understates
    progress, while `not mapped` states the run did not do what it just did, and invites the second
    press. `onRunComplete` is a SEPARATE signal from `busy` going false — "the run ended" is the fact
    that makes other screens stale.
  - **"STILL READING" IS NOT "UNKNOWN", AND THE NEXT GATE CONFLATED THEM** (fixed 1.0.156.0, found on
    site). The abbreviations step offered **Next** while the panel said *"Reading the term store…"* —
    the count is `undefined` for the whole read and `undefined` never gates. `abbreviationsLoading` is
    a SEPARATE fact, because one value cannot distinguish "not read yet" from "read and failed", and
    only the first should hold a button: it clears itself in seconds, where gating on a failure would
    strand the admin for good. Same split as `subjectGiven` — **fail-open exists for reads that can
    FAIL, not for reads still running.** Pinned by test to gate that step alone.
  - **All five defects are fixed in 1.0.156.0, and only the first has been run on a site** — see
    `docs/2026-08-18-bulk-groups-handoff.md` before running it.
- **MEMBERSHIP MOVED *TO* FOLDER ACCESS, AND MAPPING LEFT IT (2026-08-18, client's instruction).**
  Spec `2026-08-18-folder-access-membership-design.md`. Supersedes the membership half of the
  2026-08-14 separation below; everything else there stands.
  - **Bulk provisioning removed Folder Access's original job.** Creating a group writes its Group Map
    rows and a bulk run writes 790 in one press, so the mapping form was the fallback for a hand-named
    group — while sitting at the top of the page, stating that mapping is manual work done here. And
    the one genuinely recurring job in the system, **putting a person into their unit's group**, was on
    Group Management, a page an admin otherwise visits twice in a segment's life.
  - **Group Management answers "what groups exist"; Folder Access answers "who is in them, and what
    they reach".** The persona stays on Group Management and CANNOT move: a persona decides the group's
    NAME, so it must be known before the group exists. Anything Folder Access does is after the fact.
  - **ONE ROW PER GROUP (308), not per mapping (790), and that is forced.** Membership is a property of
    the GROUP, and a unit's approver group carries six mapping rows — a per-mapping editor would show
    the same people 13 times per unit with 13 Add boxes doing one thing. `groupMappingsByGroup` in
    `shared/groupMappings.ts` (pure, 6 tests) keys on **`GroupId`, never the name**: two groups can be
    renamed alike, a stored name can be stale, and the id is what `removeGroupMember` needs.
  - **`GroupMembersEditor` is MOVED, not copied** — Group Management lost its editor. Two lists of the
    same people drift, and re-merging these pages is what the client called confusing on 2026-08-14.
    **⚠ PARTLY REVERSED 2026-08-23** — the editor is mounted on Group Management again at the client's
    request (see *GROUP MANAGEMENT ANSWERS "WHAT CAN THIS PERSON REACH?"*). Folder Access keeps its
    mount, so the rule stands in the form that mattered: **one component, two mount points.**
    Not `accessMemberUi.tsx`, which is removal-only for library/page grants and gains no Add: adding
    someone there would look scoped to that library or page and never is.
  - **⚠ THE ADD-MAPPING FORM IS NOW MOUNTED NOWHERE (2026-08-23, client: *"I think you can remove
    map a group by hand"*).** It left Folder Access on 2026-08-18 and Group Management today, so
    nothing passes `show="form"` and that whole branch of `GroupMapBuilder` is unreachable. **The code
    is KEPT, not deleted** — re-mount by passing `show="form"`.
    - **TWO CASES NOW HAVE NO ROUTE AT ALL, and both are silent.** A group made with the advanced
      free-text name and no persona gets **no rows whatsoever** (`rowsForNewGroup` returns `[]`), and
      one group covering two tiers cannot be expressed anywhere else. The only repair for either is
      **deleting the group and re-creating it with a persona** — which is also the verified repair
      path for a deleted group, so it is a known road rather than a new one. Acceptable because bulk
      provisioning writes every mapping a segment needs and the persona picker names the group; the
      form was reachable only behind a collapsed disclosure and, on the client's reading, invited
      exactly the hand-editing the persona model exists to remove.
    - **The `How roles and permission levels work` reference went with it**, since it rendered from the
      same mount. That text is the only place naming how to CREATE `CRS Upload` / `CRS Approve` /
      `CRS Delete` (Site settings → Permission levels) and what a missing level costs — a newly
      provisioned unit silently gets no uploader grant. Both live sites already have the levels, so
      nothing is broken today; if a THIRD site is ever provisioned, that guidance must come back
      somewhere or reconciliation's `no "CRS Approve" role definition on site` line is the only clue.
    - A `show` prop (`"members"` | `"form"`) gates the two halves — **two mount points of one
      component, never a copy.** They share `existing`, `postRow` and `isDuplicateRow`; two components
      would mean two definitions of a mapping row, and the drifting one would be the rare one.
    - **Not added to the guided flow's group step**, deliberately: that step creates groups, and
      mounting this there costs a second 790-row read for a rare job.
  - **The select-all checkbox is GONE with the flat table.** Behind 308 collapsed headers it would have
    put "delete every mapping on the site" two clicks away. Per-row checkboxes inside an expanded group
    still feed Delete selected.
  - **`SHARE` HAD NO `ROLE_LABEL` ENTRY** until 2026-08-18, so it rendered as the bare code beside
    five roles reading as sentences. The fallback is deliberate — an unknown value must stay visible —
    which is exactly why a MISSING label reports nothing. Now pinned by a test iterating **`PERSONAS`,
    not `SELECTABLE_ROLES`**: that list is the roles derivable from a group NAME and excludes `SHARE`,
    which is how the gap survived.
  - **FOLDER ACCESS IS NEVER A REQUIRED STEP**, and must never join `NEXT_GATED_STEPS`: membership is
    INTENT, and intent is not checkable — nothing can tell whether the RIGHT people are in a group.
    **Reconciliation grants to a group, not to its members**, so an empty group is a valid end state:
    the grant is in place and applies the moment someone is added, with nothing to re-run.
- **THE GROUP LIFECYCLE LEFT FOLDER ACCESS (2026-08-14, client: *"the group creation is done in
  folder creation and its confusing"*).** Spec `2026-08-14-group-management-separation-design.md`.
  New web part **`Group Management`** (`3f81c6d2-…`, inside the `user-access-web-parts` bundle) owns
  creating a group, its members and deleting it. Folder Access now does ONE thing: map an
  **existing** group to a segment, tier and role.
  - The confusion was a **MISSING STATE**, not a layout problem. Creating a group was reachable only
    as half of "Create group & add mapping" — one indivisible action that rolled the group back if
    any mapping row failed — and its mirror deleted the SP group when its **last** mapping row went.
    So "this group exists but is not assigned yet" could not be expressed.
  - **THE AUTO-DELETE-ON-LAST-ROW RULE IS GONE, and removing it was required, not incidental.** Once
    "created but not yet assigned" is legitimate, that rule silently destroys a group an admin made
    minutes earlier. `onDelete` now deletes A ROW. Deleting a group is deliberate, on the new page,
    and **says that folder permissions survive until reconciliation runs** — removing the mapping is
    not removing the access, and an admin who believes otherwise stops looking.
  - **The naming convention survives as a NAME BUILDER that writes nothing.** `suggestGroupName`
    derives `GHO_GF_CORU_UPLOADER` from segment + tier + role, and Folder Access still reads that
    suffix back via `roleFromGroupName` to pre-select a role. The new page offers the same cascade
    purely to fill the name box, and asks before overwriting a hand-typed name.
  - **Folder Access's "➕ Create a new group" row became a SIGNPOST** ("Groups are created on the
    Group Management page"). With five separate access pages nothing else tells an admin the order;
    that one line is the whole mitigation for not merging them into a tab bar.
  - **The Full Control check left with the features that needed it.** What remains here is a Group
    Map LIST write, governed by list permissions — a "you need Full Control" banner would warn the
    wrong people and reassure the wrong people. `Group Management` keeps the check.
  - Gone from Folder Access with it: the per-row **Members** modal, the inline member editor and the
    **People** tab. The file dropped 2325 → ~1390 lines, back under the lint ceiling.
  - **No list schema change and no migration** — same Group Map rows, same groups.
  - Audit gained `GroupCreated` / `GroupDeleted` / `MembersChanged`, safe because `EventType` is a
    **Text** column. `GroupMapChanged` stays, and stays on Folder Access: it describes a MAPPING.
- **⚠ A NEW GROUP DOES NOT REACH A SESSION ALREADY OPEN (observed 2026-08-18, cost an hour).** A guest
  added to `GHO_GF_TAX_APPROVER` was bounced off `Upload-Form.aspx` to the Documents library while EVERY
  permission read back correct: group membership, the page role assignment, its `Read` binding, and the
  user's own `currentuser/groups` listing the new group. A full sign-out and sign-in fixed it instantly.
  It is the signed-in SESSION that is stale, not the permissions.
  - **The toast must not say "access is immediate."** It did, which is why the first conclusion was that
    the tool had failed, and the diagnosis ran through six API checks before anyone tried the obvious.
  - Diagnostic order for "I added them and they still cannot get in": is the principal the right one (a
    gmail guest can exist TWICE for one address), is the group granted on the page, is the binding
    `Read`, does their own session list the group — then **have them sign out and back in.**
- **REMOVING ONE PERSON'S ACCESS IS BUILT ON BOTH ACCESS SCREENS (2026-08-14, client: the two
  screens *"doesn't make sense in terms of user experience"* — removing someone should remove the
  USER, not their whole group).** Spec `2026-08-14-per-person-access-removal-design.md`; rules in
  `shared/accessMembers.ts` (pure, 28 tests), UI in `userAccess/components/accessMemberUi.tsx`,
  mounted by BOTH `StagingAccess.tsx` and `PageAccess.tsx`. Expand **People** on any granted group
  to list its members and remove one.
  - **IT IS A GROUP-MEMBERSHIP CHANGE, and can only be.** Library and page grants are held by
    GROUPS, so the only per-person lever is taking them out of the group — which is the SAME
    `*_UPL`/`*_APR` group reconciliation maps at folder scope. So this is never scoped to the
    library or the page however the button reads: **they lose their unit folder too.** Stated in the
    intro banner, the confirm dialog and the audit row, because an admin who believes otherwise has
    quietly revoked someone's ability to upload.
  - **Individual direct grants were REJECTED** (a role assignment on a user principal would be
    genuinely narrow). They manufacture the exact "granted directly in SharePoint — reconciliation
    will not remove them" state both screens already flag in red, with no Group Map row explaining
    why. Achievable, so do not call it impossible; re-argue the trade before re-proposing.
  - **A PERSON IN TWO ALLOWED GROUPS KEEPS ACCESS**, so `also in GHO_GF_CORU_APR` is on the member
    row itself — before the click, not only in the dialog, because by the dialog the admin has
    already chosen the row. Without it a CORRECT removal looks like a failed one.
  - **`removalVerdict` has THREE answers, never a boolean:** `ends` / `survives` / `unknown`.
    `unknown` exists because a group whose members failed to load renders like any other collapsed
    row, so "this ends their access" would be asserted from a picture with an invisible hole.
    `survives` beats `unknown`. Unlike the fail-open rules elsewhere, this does not refuse the
    removal — it refuses to CLAIM the removal was sufficient. `lastMember` stays false on an
    unreadable list for the same reason: understate, never assert.
  - **A failed member read is an ERROR state, not `[]`.** `MemberLoad` is a 3-state union and the
    cell says `could not read members`, never `no members` — advisory when the column was
    decoration, load-bearing now a removal is decided from it. An admin told a group is empty stops
    looking for the person they came for.
  - **Removing the last member does NOT remove the grant** — the group stays mapped and allowed,
    just empty, which reads as configured. Dialog says so; the row goes amber.
  - **On an UNRESTRICTED page the control is inert** — everyone with site access can open it, so
    removal changes nothing about who can. Said in four places (row, dialog, toast, audit row):
    this is the case where it looks like it worked and did nothing.
  - Page Access **had no audit trail at all**; it has one now (`MembersChanged`, plus the existing
    `EventType`-is-Text safety). `Outcome` is `Success` ONLY when access genuinely ended — survived,
    unconfirmed, inert and errored all record `Failed`, because a row reading "member removed"
    stops someone looking.
  - PageAccess's read-only members modal is **gone, not kept alongside** — two lists of the same
    people drift after a write. Adding people stays on Group Management; these screens only remove.
- **NEVER BUILD A URL FROM `libApiTitle()`/`libraryTitle()` IN A RENDER-TIME `const` (2026-08-14).**
  They read a module cache primed by `primeNames`, which has NOT run on the first render — where they
  still answer with the legacy `Staging`. StagingAccess's `listBase` was such a const, so the value
  the FIRST render captured was the one the mount effect's `loadLive` used: a 404. Later renders
  computed the right URL but nothing re-read the ACL (the effect keys on `[library]`, which never
  changes), so `Access now` read `unknown` forever, and Allow/Remove could not apply a live
  permission. **The tell was that the banner named "Approval Document" correctly while the request
  had asked for "Staging"** — the displayed value came from a later render than the closure did.
  `listBase` is a FUNCTION now; do not turn it back. Corollary: a name resolved at render time is
  fine for DISPLAY and never for a request.
- **THE HEADER IS THE DESIGNER'S TWO-CARD BLOCK (2026-08-15)** — explanation left, callouts right.
  The green *"Library access does not expand document access"* is a permanent fact and always shows.
  **The amber *"Unable to verify current permissions"* is a STATE and shows only when the ACL read
  actually failed** — the mock drew it as a static sibling, which would tell an admin on every visit
  that permissions cannot be verified and train them to ignore it on the day it is true. Its copy
  drops the mock's *"until permissions are synchronized"*: **there is no sync and waiting fixes
  nothing** — a 404 is the wrong library title, a 403 is not holding Full Control, opposite fixes,
  which is why the status stays named in the heading. The mock's body text also said `Staging` twice;
  every string here resolves the live title.
- **A failed ACL read now REPORTS ITS STATUS**, in the banner and the console. It had said only
  "could not read": `404` (wrong library title) and `403` (no Enumerate Permissions — reading a
  library's permissions needs Full Control on it) are the same sentence with opposite fixes, and the
  banner now names which one happened. Gotcha #9's rule, learned in a second place.
- **`libApiTitle` NOW LIVES IN `shared/naming.ts` (2026-08-14), and StagingAccess was broken without
  it.** It builds URLs from the logical `LibTarget` key, and `getbytitle('Staging')` 404s on a site
  whose library is titled `Approval Document` — so **every "Access now" cell read `unknown` and
  neither Allow nor Remove could apply a permission.** `ApprovalLibraryAccessPage` already claimed
  the translation happened "at the API boundary"; it never did, because the helper was a private
  const in `FolderManager.tsx`. Gotcha #12 again: **translate at the API boundary, never in stored
  data** — the `Target` filter and the written row still use the key `Staging`. The screen also
  showed that retired name in 11 user-facing strings; they now resolve the live title.
- **The site-entry rule has ONE implementation: `shared/siteEntryGroup.ts`** (2026-08-14). It was
  written out three times (GroupMapBuilder, SiteAccess, reconciliation) and this change would have
  made a fourth. `ensureSiteEntryGroup` **throws rather than creating when the group list cannot be
  read**: SiteAccess used to create whenever its `entry` state was empty, which is also what a
  FAILED read looks like — producing a second entry group, both looking correct, with everyone's
  real access sitting in the original. `addMemberWithSiteEntry` adds to the target group and the
  entry group and returns a `note` (it never throws) that every caller must surface. **Absent ≠
  unreadable** here as everywhere: one is a setup step the admin can take, the other is not.
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

## FOLDER MANAGEMENT IS GUIDED FLOWS NOW (2026-08-14, spec `2026-08-14-folder-management-guided-flows-design.md`)
Client: the tab bar *"is confusing and client doesnt know how it works"*. The tabs were already in work
order, which was not enough — five equal doors do not say four of them are steps of one job.
`FolderAdmin.tsx` is now picker → flow runner → **All tools** (the five tabs, unchanged). Rules in
`shared/folderFlows.ts` (pure, 45 tests). BUILT, not yet site-tested.
- **"ALL TOOLS" LEFT THE PICKER, AND RECONCILIATION GAINED A CARD (2026-08-20, client: *"I don't want
  client to touch and mess it all up without following proper flow"*).** The five screens have a real
  ORDER — an abbreviation before reconciliation, a saved pending chain before a migration — and opening
  one directly is how that order gets skipped.
  - **The VIEW stays, reachable by `#tab=<slug>`.** Those links shipped on the CRS Settings page and may
    be bookmarked; a public name that stops working is worse than one nobody clicks. Removing the
    invitation is not removing the capability.
  - **⚠ WHAT MADE IT SAFE WAS THE NEW `runRecon` FLOW.** Reconciliation was the only screen with no flow
    of its own — it is the tail of three others — so without a card an admin needing to re-run it would
    have had to start the flow for ADDING A UNIT and jump the rail. `needsSegment: false`, because the
    reconciliation screen has its own segment tick-list and a picker in front of it would ask twice.
  - It reuses the SAME `RECONCILE` step object, lock included. That lock can never fire there (no
    abbreviations screen ⇒ the count stays `undefined` ⇒ unknown never gates), and one definition beats
    a lock-free copy that could drift.
  - **⚠ THE CARD MUST NOT SAY "after adding people to a group" — corrected by Clarence before it
    shipped.** Reconciliation grants to the GROUP, never to its members, so a new member inherits the
    grant the moment they join and there is nothing to re-run. The real cases: a group deleted or
    re-created, a folder's permissions changed by hand in SharePoint, a group's mappings removed, or
    simple doubt about a segment's state.
- **THE GOVERNING RULE: *"the flow should not stop them from doing the work."* ⚠ ATTRIBUTED TO THE CLIENT
  UNTIL 2026-08-17, WHEN CLARENCE CORRECTED IT — it was never theirs.** Kept, because it earns its place,
  but as OUR design rule, so it may be traded off rather than cited as a constraint handed down. Their
  actual ask is narrower and stronger: *"I just want to ensure this flow is working properly and it should
  make them understand how it work."* So a step **definitively** not done must say so and hold the Next
  button (`blocksNext`), while a step merely unchecked must block nobody — fail-open guards against a
  WRONG check, not against checking. Every
  step is reachable from the rail; the rail SAYS what is outstanding and padlocks almost nothing.
  `isLocked` returns true only when a fact is **known and unmet** — anything `undefined` (unread, or a read
  that failed) is not a lock. **Fail-open by construction, not by remembering to.**
- **Five flows.** Add a new segment · **Add a department or unit** · Change the folder structure · Rename
  or re-code a folder · *Retire a segment* (destructive, styled apart). **Flow 2 is the everyday one** and
  is literally the tail of flow 1 — pinned by a test, because if they diverge one of them is wrong.
  Without it the frequent job has no home, and an admin starts the new-segment flow and skips half of it,
  one wrong click from a duplicate segment.
- **NEXT IS DISABLED UNTIL THE STEP IS DONE — on exactly TWO steps (2026-08-17, client: *"won't it make
  more sense once client finish filling up the New Segment and saved and only next step is
  available?"*).** `blocksNext` in `shared/folderFlows.ts`, gating `createSegment` and `abbreviations`;
  the reason renders **beside** the greyed button, because an unexplained disabled button reads as a
  broken page and the admin's next move is to reload rather than finish the step.
  - **It gates on an explicit ALLOW-LIST, never on "any `todo` step",** because most facts behind
    `stepState` are advisory and gating them would trap someone who did the work another way:
    `groupsExist` is read by the naming convention that `suggestGroupName` only *suggests*, so a
    hand-named group reads as absent; `subjectFound` misses on a genuine spelling difference; mapping
    groups after building folders is a legitimate order. Pinned by a test that walks every step of every
    flow with every fact false and asserts exactly two gate.
  - **UNKNOWN never gates**, same rule as `isLocked`. A throttled config list would otherwise strand an
    admin on step 2 with no way forward.
  - **⚠ BUT A BLANK FORM FIELD IS NOT "UNKNOWN", AND CONFLATING THE TWO SHIPPED A GATE THAT NEVER
    FIRED.** The first build derived `segmentExists` only once a name had been typed, so with the field
    empty — the state every admin starts in — nothing was gated and Next stayed enabled on a step visibly
    not done. The client reported it twice. **Fail-open exists for READS THAT CAN FAIL, not for a text box
    nobody has filled in:** a throttled list cannot blank a local field, so gating on the field strands
    nobody, whereas gating on an unreadable list strands everybody. Hence `FlowFacts.subjectGiven`, set
    **only when the segment list was readable**, which keeps the one genuine fail-open case intact.
  - `blocksNext` therefore answers `createSegment` in a fixed order: **exists ⇒ allow** (a page reopened
    after the work is done has no typed name and must not be told to create a segment that already
    exists), then **name blank ⇒ "type the name first"**, then **typed but unmatched ⇒ "create it
    first"**. A test pins that `subjectGiven: false` gates NO other step, since it is set for the whole
    flow and every step sees it.
  - **Flow 1 CONFIRMS THE CREATED SEGMENT WITH A PICKER (`asksSubject: "newSegment"`), and that is what
    made any of this answerable.** `segmentExists` is only computed for a CHOSEN segment, and flow 1 had
    no picker — the segment does not exist yet — so **the `segmentExists` padlock on its abbreviation step
    could never fire in the one flow it was written for** (found on the client's site 2026-08-17).
    Choosing the segment after Create sets `segKey`, which answers the gate AND carries the segment into
    steps 3-6 instead of leaving each to ask again.
  - **THREE DESIGNS WERE TRIED IN ONE SITTING; the first two are instructive.** *Counting `mode` rows*
    fails on a page refresh — a baseline taken when the flow opens resets, so the flow then refuses work
    already done, worse than no check. *Asking for the name as text* asked for the same value the form
    below already asked for (the client: *"what is What will the new segment be called?"*), and reads as
    a bug rather than a feature. **Picking beats typing on every count:** nothing to spell — so no
    fullwidth-＆ trap — it is the control every other flow already uses, and it is derived from data, so
    it survives a refresh, a second tab and someone else's session.
  - **⚠ THE SEGMENT LIST IS READ ON MOUNT, AND CREATE HAPPENS AFTER IT.** So the segment just made was
    missing from its own confirmation, and the rail said *"Not checked"* beside a panel saying *"is
    configured"* — a screen contradicting itself, which is worse than one that says nothing. Hence the
    `reload` counter and an explicit **Refresh list** button. Do not assume a mount-time read reflects
    anything a step below it has since written.
  - The fact is **derived at render, not in the facts effect**: that effect performs three list reads, so
    keying it on the picked segment would fire all three per change.
  - Both gate messages **name the way out** (*"or jump straight on from the list of steps"*), because the
    gate can legitimately be wrong after a refresh. The rail stays clickable throughout, and a correct
    gate with no stated escape reads as a dead end.
- **THE SEGMENT PICKER DOES NOT STAND IN FRONT OF AN INSTRUCTION-ONLY STEP** (client, 2026-08-19:
  *"weird the first step is asking for the business segment when it is suppose to be showing only a sign
  to users"*). It replaced the content of EVERY step of a segment-scoped flow until a segment was chosen
  — including a `kind: "outside"` step, which renders nothing but instructions for work done elsewhere
  (pause a Power Automate flow, edit the term store) and uses no segment. So the one thing that step
  exists to say was hidden behind a dropdown asking for a value it has no use for. The next step asks,
  which is where the subject is first needed.
- **FOUR locks, each on ONE definitive read:** segment exists (flow 1's later steps, and Retire),
  `PendingLevels` set (before migrating), and **reconciliation blocked while any term lacks a code** — the
  last is the one that matters, because that is the silent failure (recon skips the term, creates no
  folder, no error anywhere).
- **⚠ THE ABBREVIATION COUNT WAS NEVER ACTUALLY WIRED UP UNTIL 2026-08-17.** `AbbreviationManager`
  computed `missing` for its own warning banner and told nobody, so `FlowFacts.abbreviationsMissing`
  stayed `undefined`, unknown gates nothing, and **Next was clickable on a screen covered in "no folder
  will be created" warnings** — the exact silent failure the lock exists to prevent. Fixed with
  `onMissingChange` → `IFolderManagerProps.onAbbreviationsMissingChange` → `FolderAdmin`, the same
  report-upward shape as `onCreated`. It reports the **live** value, not the saved one, so the count on
  screen and the gate can never disagree; and it reports `undefined` — never 0 — while loading, with no
  segment chosen, or for a term set with no terms, because each of those computes 0 from an empty list and
  0 means "all done" to a caller.
- **The abbreviation count IS affordable, contrary to an earlier draft of the spec.**
  `AbbreviationManager` already walks the tree and already counts terms with no code, so the number is
  exact **once the step is open**. Only pre-walking every segment on the picker is unaffordable (~115
  requests for GHO), so the picker says "not checked" and the lock arms after the step is visited.
  **`abbreviationsMissing: undefined` must never be read as zero.**
- **Flows 2 and 4 ASK what you are doing** ("adding Treasury under Group Finance"), which is the only
  thing that makes their term-store step checkable — and it pre-fills the later steps. **Matching folds
  the FULLWIDTH ＆** (GHO's `Group Legal, Risk ＆ Compliance`), case, whitespace and zero-width chars; a
  miss says *check the spelling*, never *you have not done step 1*. The subject is **optional** — blank
  costs help, never progress.
- **It DRIVES `FolderManager` (new props `initialTab`, `hideTabs`) and never dismantles it.**
  Reconciliation is inline in a 4,000-line file and is the most site-verified code here; extracting it to
  make it mountable would risk the wrong thing for a navigation change. Cost: one re-mount per step.
  - **⚠ `PendingLevels` MUST BE READ IN ITS OWN REQUEST, AND `FolderAdmin` WAS NOT DOING IT** (fixed
    1.0.163.0, found 2026-08-18 while explaining why the migrate step was locked). That column is
    created ON DEMAND by `StructureManager` the first time a change is staged, so on any site where
    that has never happened it does not exist — and one unknown name in a `$select` fails the WHOLE
    request with HTTP 400 (gotcha #11). It was in the same `$select` as the segment list, so on such a
    site **every guided flow reported "the segment list could not be read" and offered no segment at
    all** — on a site that is otherwise perfectly provisioned, and on the one screen every folder job
    starts from. `StructureManager` and `SubtreeMigrator` both split this read and say why in a comment;
    this file did not.
    - The flag is now **three-state**: `undefined` when the column is absent or the read failed, which
      locks nothing, versus `false` which locks the migrate step. Same rule as everywhere else here.
    - **Most likely to bite on CRS**, where nothing has ever staged a structure change.
  - **⚠ THE RE-MOUNT NEEDS A `key`, AND ITS ABSENCE WAS A LIVE BUG (found on the client's site
    2026-08-17, first time anyone stepped through a flow).** `FolderManager` resolves `initialTab` in a
    `useState` **initialiser**, which React runs once per mounted instance — so with no key React
    reconciled the same element type across a step change, kept the instance, and never read the new
    prop. Stepping from *New segment* to *CRS Term Abbreviations* moved the heading and the rail
    highlight while **the New segment form stayed on screen**. Worse than a dead link, because it looks
    like it worked: the admin fills in one screen under another step's title. Fixed with
    `key={st.screen.tab}` in `FolderAdmin`.
  - Keyed on the **tab**, not the step id: `ABBREVIATIONS` is a step of four different flows, so keying
    on the tab re-uses the instance where the screen genuinely is the same and re-mounts only when it
    changes. A step-id key would throw away a half-typed screen on any same-tab move.
  - **The general trap: a prop named `initialX` is read ONCE.** Anything driving such a component from
    outside must either key it or the component must watch the prop — and "it renders the right heading"
    is not evidence either happened.
- **Group Management and Folder Access are mounted from `userAccess/components`** inside the flows AND stay
  standalone on the landing page — one component, two mount points, never a copy.
- **What no checking reaches:** the right PEOPLE in a group, the right abbreviation (`TRS` vs `TREAS`), the
  right parent term, whether they finished in the term store, anything in Power Automate. **Existence is
  checkable; intent is not** — every tick means "this exists", never "this is right".
- **A migration does NOT need the Power Automate flows turned off** (Auto-route spec §5.6). Auto-route
  reads moderation status in its body, so a moved pending/rejected file takes the False branch and nothing
  is touched. The step advises **both ways**: leaving them on is safe, pausing avoids hundreds of no-op
  runs burning the daily quota — and an exhausted quota means the next real approval is not routed,
  **silently**. Residual: a stale approved file still in the library gets routed mid-run and can land in
  `Documents` under the pre-migration path.
- **Landing page: Folder Management is now ONE link, not three.** Term Abbreviations / Folder Structure /
  Reconciliation stopped being destinations. `#tab=` links still work — `FolderAdmin` opens **All tools**
  for one — so no bookmark breaks; `#flow=<id>` opens a flow directly.

## A BLANK `termSet` HID YEAR AND DOCUMENT TYPE FOR A WHOLE SEGMENT (2026-08-26, SDG)
Found on SDG's first live upload. The upload form's folder card showed Segment / Department / Unit and
then stopped - no Year, no Document Type - so GHO documents would have filed TWO LEVELS SHALLOWER than
MHO and NBPOLHO, in the same system, with nothing erroring anywhere.
- **THE CAUSE IS ONE UNTOUCHED DEFAULT.** GHO's `Levels` carried
  `{"label":"Year","column":"Year","labelCol":"Year","termSet":"","permissioned":false}` - an EMPTY
  `termSet` - while MHO and NBPOLHO held the real GUIDs. An absent or empty `termSet` on a below-Unit
  tier **IS the discriminator for a per-unit tier**, so the form looked for child terms under the
  selected Unit, found none, and hid the tier.
- **The Folder levels screen defaults the tier source to "Under each Unit", and switching to per-unit
  deliberately CLEARS the GUID.** Editing a level in place is not supported - you remove and re-add -
  so re-adding `Year` without switching the toggle to "One shared list" produces exactly this.
  The no-`tidCol` shape in the stored row is the fingerprint: `builtInTierFor` restored the built-in
  column shape, which only happens on the add-tier path.
- **WARN: FOUR COMPONENTS BEHAVED CORRECTLY AND THE COMBINATION WAS A SILENT MISFILE.**
  `validateChain` only checks that permissioned tiers form a contiguous prefix. The depth check
  compares the SEGMENT term set's depth to the permissioned tier count and never looks below Unit.
  **Reconciliation does not read `Levels` at all** - it walks the term tree. And the upload form hides
  a cascading tier with an empty option list ON PURPOSE, so an optional SubUnit needs no
  configuration. No layer validates a below-Unit term set, and each has a good reason not to.
- **THE FIX THAT REMOVES THE CHOICE BEATS THE FIX THAT WARNS ABOUT IT.** `Year` and `Document Type`
  are built-in managed-metadata columns bound to SITE-WIDE term sets; a per-unit Year is meaningless.
  The toggle should not be offered for those two names, and `builtInTierFor` - which already restores
  their column shape verbatim - should restore their term set too. A reconciliation warning
  (*"below-Unit tier X has no term set and no unit has child terms"*) is the weaker second option.
  NEITHER IS BUILT.
- **Repaired by editing the `mode_gho` row's `Levels` directly**, safe because SDG's segments were all
  `EMPTY` so no documents sat in the old shape. On a segment in use this would be a migration.
- **The tell, if it recurs: compare segments against each other.** One segment filing shallower than
  its siblings is invisible in any single screen and obvious the moment two `Levels` rows are read
  side by side.

## `HC folder approval` WAS ON AND FAILING EVERY RUN FOR SIX DAYS (2026-08-25)
Found because an HC upload refused with *"Could not create the 2025 folder in the Highly Confidential
library"*. The flow's status read **On**; every run since it was created on 19 August had **Failed** in
about two seconds.
- **THE CAUSE IS THE SEVENTH HC-CLONE LIBRARY REFERENCE.** Its MERGE read
  `_api/web/lists/getbytitle('Approval Document')/items(<ID>)` - the NORMAL library - while the trigger
  correctly fired on `HC Approval Document`. **Item ids are per-LIST**, so the HC folder's id did not
  exist there and SharePoint answered *"Item does not exist. It may have been deleted by another
  user."* One string: `'Approval Document'` -> `'HC Approval Document'`.
- **WARN: IT FAILED SAFELY ONLY BY LUCK.** That MERGE sets `OData__ModerationStatus: 0` - **Approved**.
  Had any id existed in the normal library, the flow would have approved whatever held it; if that were
  a pending FILE, Auto-route would then have routed an unreviewed document. **Approval bypass, on a
  flow nobody was watching.** Check `All runs` for any `Succeeded` before assuming a broken flow of
  this shape was harmless.
- **THE SYMPTOM APPEARS TWO PEOPLE LATER AND NAMES THE WRONG THING.** A below-Unit folder created by
  one HC PIC stays Pending, so it is invisible to the NEXT HC PIC - who cannot see it, tries to create
  it, and is refused for a DUPLICATE NAME. The error blames folder creation; the fault is a flow that
  failed days earlier for someone else. Verified: `2024` and `2025` under
  `HCApprovalDocument/GHO/GCA/EG` both read moderation `2`, created by two different people.
- **WARN: WHY IT SURVIVED SIX DAYS - THE STRUCK-THROUGH LINE IN THIS FILE.** *"Both folder-approval
  flows are now unnecessary and are being turned OFF"* was acted on twice by mistake and struck through
  both times - but its lasting damage was different and worse. **A flow believed redundant never has
  its run history read.** Nobody asked whether it worked, only whether it should be on. A wrong claim
  that something is unnecessary hides its failures as effectively as deleting it would.
- **WARN: THE FIRST FIX ATTEMPT TYPED `'HCApproval Document'`** - a hybrid of the TITLE (`HC Approval
  Document`) and the URL segment (`HCApprovalDocument`), matching neither. Gotcha #12 in a new place,
  and it would have failed as `NotFound` **exactly like the original bug**, sending the diagnosis round
  the same loop a second time. `getbytitle` takes the title; only the URL has no spaces.
- **THE FIX DOES NOT BACKFILL.** The flow fires on creation, so folders already stranded stay Pending
  until someone approves them by hand or the tree is rebuilt. Here GHO's HC departments were deleted
  and reconciliation rebuilt them - a supported recovery path, and reconciliation stamps new folders
  Approved where the library moderates, so the rebuilt tree came back clean.
- **VERIFIED:** eight consecutive `Succeeded` runs, and a folder created by one PIC was immediately
  visible to another.
- **WARN: THE OTHER HC CLONES HAVE NOT BEEN RE-CHECKED THIS WAY** - the three HC audit flows and
  `HC auto-approve`. Same clone, same trap, and the audit ones fail into a log nobody reads.

## APPROVERS ARE EMAILED WHEN A DOCUMENT ARRIVES (2026-08-25) - BOTH FLOWS VERIFIED END TO END
Client: *"can we build the notification to alert approval that the file has been uploaded and in need
of their approval"*. Runbook `2026-08-25-approver-notification-flow-runbook.md` - Power Automate only,
**no code**. **BOTH flows are BUILT AND VERIFIED END TO END on ClarenceDMSTesting.** `NotifyApprovers`:
uploaded into GCA > GCBC, one email arrived naming the document, the unit and the uploader, its link
opened that document on the approval page, and the approval went through. `HCNotifyApprovers`: uploaded
into GCA > EG as an HC-cleared PIC, the HC approver was emailed, approved through the bulk approve
panel, and `HC Auto Route` moved the file to `HC Documents/GHO/GCA/EG/2025/Approval Papers/Archive 2`
with segment, department, unit, document date and `LegallyPrivileged` intact.
- **THE APPROVER IS FOUND THROUGH THE GROUP MAP, NEVER THE FOLDER PATH.** The path implies the group
  name (`/ApprovalDocument/GHO/GCA/GCBC/...` -> `GHO_GCA_GCBC_APPROVER`), and deriving it that way is
  the trap: **renaming an abbreviation renames the FOLDER and leaves the group alone** (1.0.162.0), so
  the derived name would stop matching, `Get group members` would return nothing, and **the flow would
  go on succeeding while nobody was ever notified**. It reads the document's own `UnitTid` and filters
  the Group Map on `Role eq 'APR' and UnitTermGuid eq '<UnitTid>'` - the same join reconciliation uses,
  so the two cannot disagree about who approves a unit. A rename never touches a term GUID.
- **WARN: AN EMPTY APPROVER GROUP FAILS THE RUN UNLESS IT IS GUARDED, AND THE RUNBOOK NAMED THE
  EXPECTED BEHAVIOUR WITHOUT NAMING THE MECHANISM.** Section 9 said *"upload to a unit whose group is
  empty -> no email, no failure"*, and the flow built exactly to spec does the opposite: `Recipients`
  joins to an empty string and `Send an email (V2)` returns **"To Field cannot be null or empty"**.
  Hit on the first live test, against `GHO_GCA_EG_APPROVER` (recreated that morning, nobody in it).
  Fixed with a `HasRecipients` condition - `length(coalesce(body('PickEmails'), createArray()))` is
  greater than `0` - wrapping the email. **An empty approver group is a valid state**, so the run must
  complete quietly rather than sit in the failed list looking like a broken notification system.
  - **The general lesson: a runbook that states an expected OUTCOME without naming the ACTION that
    produces it will be built without that action**, and the gap only shows on the one test case
    nobody reaches first.
  - **WARN: A SKIPPED EMAIL AND A BROKEN FILTER LOOK IDENTICAL.** `GetMembers` -> `MembersWithEmail`
    -> `PickEmails` all ran in 0s and produced nothing, which is exactly what a wrong property name in
    the filter would also produce. Only the populated-group test distinguishes them - so **verifying
    the empty case first proves the guard and NOT the chain.**
- **THE DUPLICATE-EMAIL GUARD IS `HasUnit`, and it means TWO RUNS PER UPLOAD IS CORRECT.** The trigger
  is *created or modified* - it must be, because the upload form uploads and THEN tags, so at creation
  `UnitTid` does not exist yet. The first firing therefore stops at `HasUnit` = False and the second
  sends. **Do not read the second run in the history as a duplicate**, and do not "fix" the trigger to
  created-only: that fires before the metadata the flow needs exists.
- **WARN: POWER AUTOMATE CONDITION TYPE MISMATCH - `0` TYPED INTO THE VALUE BOX IS TEXT.** `length()`
  returns an Integer, and comparing it against a text `0` fails the whole condition with a type error.
  **Enter the `0` through the fx editor**, on every one of these conditions (`HasUnit`, `HasApprover`,
  `HasRecipients`). Cost time on all three today.
- **HC IS THE ONE THAT MATTERS, AND IT IS DONE.** Client: *"Ofcourse it include HC, just make
  sure only HC people get the file"* - a plain approver emailed about an HC document learns its
  filename and its unit, which is the disclosure the whole HC split exists to prevent. The clone
  changes **three** things and nothing else: trigger library -> `HC Approval Document`, the Group Map
  filter -> `Role eq 'APRHC'`, and `&lib=hc` on the approval link (item ids are per-LIST, so without
  it the page resolves the wrong library or a different document sharing that id). **`APR` is held by
  the plain `hou` persona and `APRHC` only by `hou_hc`** - filtering the HC flow on `APR` would email
  every ordinary Head of Unit about Highly Confidential documents. Every HC clone in this project has
  shipped faults from a library reference nobody swapped; six of them in `HC Auto Route`.
  - **WARN: THE CLONE IS FOUR EDITS, NOT THREE - THE LIBRARY IS NAMED IN TWO ACTIONS.** The trigger AND
    `GetDoc` both carry a List Name, and a clone with one swapped and one not fires on HC documents
    while reading the normal library. Counting the swap as a single change is how a half-swapped clone
    gets called done.
  - **THE EMPTY PLAIN GROUP MADE THE TEST SHARPER THAN IT WAS DESIGNED TO BE, and this is reusable.**
    `GHO_GCA_EG_APPROVER` had NOBODY in it while `GHO_GCA_EG_APR_HIGHLY_CONFIDENTIAL` had one member.
    So a filter of `APR` would have matched the empty group, found no recipients and sent NOTHING -
    while `APRHC` matched the HC group and sent one email. **An email arriving is therefore proof the
    filter is `APRHC`**, established by the flow's own behaviour rather than by reading the config
    back. When a negative test is hard to stage, look for a unit whose data already forces the two
    branches apart.
- **Built as `clarence@trinergydigital.com` here. On SDG's tenant, sign in as the SERVICE ACCOUNT
  before the first action** - the connection is baked in for life, and a flow built as a person stops
  silently when that password changes.
- Bulk imports are excluded (client: *"not needed since its a historical file"*) - they auto-approve,
  so nobody has to act, and a 50-file import would send 50 pointless emails. That clause depends on
  **`BulkImport` being VISIBLE on the library**: `Hidden: true` removes a field from the trigger
  payload entirely, so the expression reads null, the clause passes, and bulk imports email everybody.

## ⏭ NEXT UP: EMAIL BUNDLING, 11 CLIENT TEMPLATES, AND TWO NEW FLOWS (2026-08-28)
Spec: **`docs/superpowers/specs/2026-08-28-email-bundling-and-templates-design.md`** — read it rather
than re-deriving; it carries the client's 11 templates verbatim. **Nothing is built.**

- **The client asked for two flows BY NAME:** an approval **reminder after 3 days**, and one to
  **notify the approver of Share and Delete requests**. The second consolidates six of the eleven
  templates into ONE flow on `CRS Requests`, branching `RequestType` × `Status` — the same shape
  `CRS — Audit request activity` already uses. Six separate flows would be six copies of one trigger,
  and the copy that drifts is always the rarely-exercised one.
- **20 uploaded files send 20 emails** because `NotifyApprovers` has `Split on` = On, i.e. one run per
  item. ⚠ **Turning `Split on` off is the wrong fix** — a poll batch is "whatever arrived since the
  last check", not "what the uploader submitted", so it would bundle unrelated uploads AND still split
  one submission across a poll boundary. It would look fixed and be wrong in both directions.
- **`SubmissionId` IS the grouping key, and it already exists** on every uploaded file in all four
  libraries (1.0.221.0). ⚠ **But group on (SubmissionId, APPROVER GROUP), never SubmissionId alone** —
  a submission can span units with different approvers, so bundling on the id alone tells an approver
  the filenames of another unit's documents. A blank id means one email per file, as today: blank is
  unknown, never a value.
- **⚠ THERE IS NO `ApprovalBatchId`.** Bulk Approve writes moderation status per file with no shared
  reference, so the "approved together" half has nothing to group on. Recommended answer is
  `SubmissionId` again; stamping a real batch id is a code change to BOTH approval screens plus a
  column in four libraries.
- **⚠ The dedupe races:** 20 runs all check "already notified?" before any writes a marker, so all 20
  send. Needs `Concurrency control` = 1 (⚠ **cannot be changed after enabling**) AND a marker row
  checked first, reusing the audit flows' `Already logged` shape. **A FAILED check must still send** —
  a duplicate email is read past, a missing one stalls a document with nobody looking.
- **⚠ Four things in the templates cannot be built as written**, all to raise with the client FIRST:
  `[Approve]`/`[Reject]` cannot be one-click (a bearer trigger URL in the bundle was already rejected
  for HC), so they collapse into one link to the page; **template 1 is not a flow at all** — there is
  no SharePoint trigger for group membership, and outsiders cannot be added since 2026-08-27 — so it
  belongs in code; template 10 promises **90 days of recycle bin when it is 93**; and **bundling
  contradicts every template**, since all eleven are written for a single file.

## BULK APPROVE FROM THE LIBRARY COMMAND BAR (2026-08-25, 1.0.248.0)
Client: a side panel in the NATIVE library view to approve several files at once — the approval web
part reviews ONE document, which is right for reading one and wrong for clearing a backlog of twenty.
`extensions/bulkApprove/` (ListViewCommandSet `6d1a9f47-...`), panel in `components/BulkApprovePanel.tsx`.
**BUILT AND VERIFIED END TO END on ClarenceDMSTesting** — 5 documents approved in one press, routed by
Auto-route to `Documents/GHO/GCA/GCBC/2025/Term Sheet/Archive 1` with metadata and `Created By` intact.
- **⚠ THE CHECKS LIVE IN `shared/approvalGuards.ts` AND ARE SHARED WITH THE APPROVAL PAGE.** A second
  approval route with weaker checks is not a convenience, it is a faster way to do the damage they
  prevent: publishing into an unlocked destination folder, silently REPLACING a same-named document
  (Auto-route copied with `nameConflictBehavior: 1` until 2026-08-25 — it renames now), or
  approving something the approver holds no
  `ApproveItems` on. `ApprovalDocument.tsx` was refactored to call the same three functions, so there
  is ONE implementation. Never re-inline them.
- **VISIBILITY IS ABOUT THE SELECTION, NOT THE PERSON.** There is no site-wide "approver":
  `ApproveItems` is granted per UNIT FOLDER, so a list-scope permission check would hide the command
  from nearly everyone. Gate 1 is the library (only the two approval libraries — never Documents, HC
  Documents or the archives); gate 2 probes the selected items. The same person sees it in one unit
  folder and not another, in the same library, which is correct.
- **⚠ ONLY `granted` SHOWS IT — `unknown` HIDES, and the opposite rule shipped and failed the same
  hour.** It was built as "`denied` hides, `unknown` shows", reasoning that a transient read must not
  take the feature from a real approver and that the approval page's panel follows that rule. **The
  probe is ASYNC**, so every selection is `unknown` for a moment, and in that window a PIC who could
  approve nothing saw the button. Found on site immediately.
  - **The asymmetry with the approval page is deliberate.** Hiding THAT panel on doubt strands an
    approver on the one screen built for approving; hiding THIS costs a shortcut while the page
    remains. Cheaper failure, so it fails the other way. A rule that is right most of the time does
    not satisfy "non-approvers never see it".
  - Cost, accepted: the button can appear a beat AFTER the selection. If that ever grates, cache the
    verdict per FOLDER rather than per selection.
- **A MIXED SELECTION IS DONE PER FILE, never refused as a block.** Four files across two units: the
  ones that pass are approved and the rest come back NAMED with the reason ("you are not the approver
  for that unit", "a document of that name is already in the destination library"). Skipped files are
  untouched and still pending. "3 skipped" would be unactionable — the whole point of doing fifty at
  once is not having to work out afterwards which three.
- **Sized for FIFTY** (the client's number): scrolling list, named per-file progress, sequential
  writes, and a failure never stops the run. The unit-folder check is cached **per unit** — fifty files
  in one folder ask once — while the clash check is per FILE, because it is about that document's name
  at that moment.
- **REJECTION SKIPS THE DESTINATION CHECKS.** Nothing is copied anywhere, so there is nothing to
  verify beyond the right to decide. The panel says outright that ONE comment is written to all N
  documents, because a rejection reason is normally per document and an approver who assumes otherwise
  writes "wrong vendor" onto fifty files.
- **⚠ IT DOES NOT CLOSE THE NATIVE Approve/Reject GAP.** That command sets the moderation field
  directly and is unaware any of this exists. This narrows the reason anyone would reach for it; only
  the same check inside Auto-route closes it.
- **⚠ REGISTRATION, AGAIN — `elements.xml` STILL DOES NOT PROVISION ON THIS TENANT.** After deploying,
  `/_api/web/usercustomactions` held only the hand-made New Folder row; the command set had to be
  POSTed by hand (`Location: ClientSideExtension.ListViewCommandSet.CommandBar`, `RegistrationId: 101`,
  `RegistrationType: 1`). **Second time on this site.** Expect to repeat it on SDG with the URL swapped.
- **⚠ AND A NEW COMPONENT ID TAKES MINUTES TO REACH THE SITE, WITH NO UPDATE PROMPT.** Changes to
  EXISTING components flow through without touching Site Contents; a brand-new component id does not —
  the site must actually take the package. Meanwhile Site Contents offers no Update, the `...` menu
  shows only Details/Remove, and the storefront Details page reports a **false version** (`1.0.148.0`,
  the same fiction CLAUDE.md already records). **The tell is the Network tab: zero requests for the
  bundle means the site's package lacks the component.** It resolved itself after several minutes — a
  remove/re-add was queued and proved unnecessary, so try waiting first.
- Registered against ALL document libraries (`RegistrationId 101` — there is no per-library
  registration), with the library check in code. The icon is an inline **base64 data URI**, not a
  packaged file: a real asset would have to provision, and provisioning is exactly what does not work
  here, so the icon would have come back blank.
- The comment box is `resize: none` — the panel is a fixed-width column, and a drag-resized textarea
  pushes the document list and the buttons out of view with no way back short of reopening.

## THE APPROVAL PANEL IS HIDDEN FROM SOMEONE WHO CANNOT APPROVE (2026-08-25, 1.0.243.0)
Client, having watched a Head of Unit for one unit open the approval page on their own upload in
ANOTHER unit: *"is there no way to like to ensure the uploader for that group dont see that page?"*
- **THE PAGE GRANT IS PER PAGE, NOT PER UNIT, and cannot be otherwise.** Holding `APR` in ANY unit
  opens `ApprovalDocument.aspx`; a page is not scoped to a unit (see the derived-page-access section).
  So an approver of GCBC who is also an uploader in EG could open the page on an EG document, see a
  fully live Approval panel, and learn only at submit — from a raw `UnauthorizedAccessException` —
  that they hold nothing there.
- **NOT A LEAK, and saying so precisely matters.** The folder ACL and Draft Item Security decide what
  is READABLE: they saw their OWN upload. Another uploader's pending file in that unit stays invisible
  (no `ApproveItems` on the folder, not the author). What was missing was the pre-check the upload
  form has made since 2026-08-12 — probe first, do not offer what cannot work.
- **THE PROBE ASKS THE ITEM, NEVER THE ROLE TABLE.** `items(N)/EffectiveBasePermissions` answers for
  this user on this document, so it accounts for the folder ACL, inheritance and site admin in one
  read, with no need to know which folder the file sits in or which persona the user holds. A role
  lookup can say "this persona approves" while reconciliation has not granted it, or while the group
  was renamed; an ACL read cannot be wrong that way. `ApproveItems` is PermissionKind 5 so bit index 4,
  read ARITHMETICALLY (Full Control returns `Low = "4294967295"`).
- **THE WHOLE CARD GOES, not a disabled version of it** (client, on seeing an explanatory card first:
  *"I think its best to remove the card"*). For someone who is only ever an uploader here, a panel
  headed "Approval" is a control they can never use, and explaining that on every visit is noise.
  What they came for — document, metadata, status — is all on the left.
- **AND THE GRID DROPS TO TWO COLUMNS.** `s.grid` reserves `280px` for the panel, so hiding the card
  alone left a blank column and a preview stopping short of the page. Overridden at the element, not
  in `s.grid`: that object is shared and static, and the column count is the one part of this layout
  that depends on state.
- **`denied` ONLY — `unknown` keeps the panel.** A failed probe must never take the controls away from
  a real approver, and the 403 at submit is still the backstop. Fail-open, as everywhere the cost of a
  wrong answer is a form out of service.
- **⚠ THE HOOK MUST SIT ABOVE THE RENDER GUARDS, and putting it below blanked the page for an hour.**
  This file returns early three times (`loading`, `fetchError`, `!item`). A `useEffect` declared after
  them does not run while `loading` is true, so the moment loading finishes React sees MORE hooks than
  the previous render, throws *"Rendered more hooks than during the previous render"*, and the
  component renders **nothing** — no error state, no message, a blank web part. Exactly the outcome
  the 1.0.194.0 guards exist to prevent, arriving by a route those guards cannot see.
  - **THE DIAGNOSTIC THAT SETTLED IT:** the item read back `200` with `ModerationStatus: 2`, so it
    existed and was readable — which rules out every guard branch at once and leaves "the component
    crashed" as the only explanation. **A blank SPFx web part with no error UI is almost always a
    hooks-order violation**, since a thrown render normally produces "Something went wrong".
  - Same rule the share-recipient picker already follows in `MySubmissions`. A warning comment now
    sits above the hook.

## GROUPS CAN BE DELETED IN BULK, AND THE OLD DIALOG WAS LYING ABOUT WHAT DELETING DOES (2026-08-25, 1.0.239.0)
Client, facing ~130 `_APPROVER` groups to remove by hand for the HC split: *"Can you add like select
multiple checkboxes? that would be easier for me to deelte."* Checkboxes, `Select all N shown`, and a
`Delete N selected` with a typed `DELETE` confirmation, in `GroupManager.tsx`.
- **⚠ THE CORRECTION THAT MATTERS MOST: the single-delete dialog said *"The folder permissions it was
  granted stay in place until Folder Reconciliation runs."* THAT IS FALSE FOR A GROUP.** It is true of
  deleting a mapping ROW; deleting the GROUP withdraws its access **immediately**, because SharePoint
  drops a deleted principal's role assignments. The admin was deleting 130 groups *specifically* to
  withdraw an obsolete HC grant, and this text would have told them it had not worked. Corrected in the
  dialog and in the audit row.
- **THE ONE-DIRECTIONAL GAP THIS EXISTS FOR: permissions can be WIDENED by a re-run and never NARROWED
  by one.** Reconciliation adds folder grants and never removes them (`groupsToRemove` is page-scope
  only — at folder scope it would strip SharePoint's automatic Limited Access entries and take
  everyone's access away). So taking a role off a persona changes what NEW mappings write and nothing
  else, and the only thing that actually withdraws the old grant is deleting the group.
- **`Select all shown` respects BOTH filters** — filter to `APPROVER`, tick all, delete. Selection
  SURVIVES a filter change, so the bar names how many ticked rows are **not currently shown**: a
  selection reaching past the screen is how somebody deletes a segment they were not looking at.
- **Groups with MEMBERS are named, not counted.** Membership is the only irreversible part; empty
  groups cost nothing to delete and re-create. Export the CSV first if any have people.
- **SEQUENTIAL, carries on past failures, ONE audit row for the run.** A 130-write burst is the shape
  that gets cut off half way; stopping on the first error would leave the admin to work out which of
  130 had gone, and the operation is idempotent anyway.
- **Re-entrancy is guarded by a REF, not by `disabled`.** `disabled={busy}` is set from state and only
  takes effect on the next render; the ref flips synchronously. A `beforeunload` guard fires only while
  a bulk run is in flight — keyed on `bulkProgress`, not `busy`, since a prompt on a one-second single
  delete is noise that teaches people to dismiss it.
- **⚠ DELETING THE GROUP MAP LIST IS NOT DELETING THE GROUPS, and confusing the two cost a day.** On
  2026-08-24 the CRS Group Map's ITEMS were cleared in the belief it would withdraw the HC grants. It
  withdrew nothing — the groups, their members and their ACLs were untouched — and 2,140 mapping rows
  then had to be rebuilt. **Deleting a mapping row removes the INTENT; deleting the group removes the
  PRINCIPAL.** Only the second changes a permission.

## THE `hou` / `hou_hc` SPLIT IS VERIFIED END TO END ON A REAL ACCOUNT (2026-08-25)
ClarenceDMSTesting, one guest account in `GHO_GCA_GCBC_APPROVER` (plain `hou`) and
`GHO_GCA_EG_UPLOADER` (`pic`):
- Approved a GCBC document through the approval page.
- `HCApprovalDocument` and `HCDocuments` both returned **AccessDenied, `Type=list`** — a normal Head of
  Unit cannot enter either HC library, which is the client's rule of 2026-08-24 working on live ACLs.
- Approving an **EG** document was refused. Unit isolation as well as HC isolation.
- The folder ACL confirms it directly: `GHO_GCA_EG_APPROVER` is PRESENT on `Documents/GHO/GCA/EG` and
  **absent** from both HC libraries.
- **THE MIGRATION IS WHAT MADE IT TRUE, not the deploy.** Before deleting and re-creating the
  `_APPROVER` groups, every one of them still held its pre-split HC grant — 111 groups on the test
  site, and the same will be true on any site where the split is deployed without the group delete.

## RECONCILIATION GRANTED NOTHING AND REPORTED ONLY WARNINGS (2026-08-25, 1.0.237.0)
A full two-segment run on ClarenceDMSTesting produced **7,677 warning lines**, every one of them
`no "Read" role definition on site` or `no "CRS Approve" role definition on site`, and **not one
folder or library grant was applied.** `Read` is a BUILT-IN level; it cannot be absent. A direct read
confirmed all of them present: `Full Control | Design | Edit | Contribute | Read | ... | CRS Approve |
CRS Delete | CRS Upload | CRS Share | CRS Request`.
- **THE CAUSE IS A STALE CLOSURE OVER REACT STATE.** `roleDefs` is `useState`, filled by a mount-time
  effect. `runReconciliation` is one long async function and reads `roleDefs` from the closure it was
  created in — so **a run started before that fetch resolves holds `[]` for its entire length**, and
  every `roleDefs.find(...)` misses. Same trap as the run counters and `stopRef`, in the one place
  where the cost is every permission on the site.
- **IT DOES NOT LOOK LIKE A FAILED RUN, AND THAT IS THE DANGEROUS PART.** Folders are still created,
  inheritance is still broken, the progress panel and counts look normal, `failed` stays 0, and the
  ACLs are simply left exactly as the PREVIOUS run left them. Reading a folder afterwards shows a
  correct-looking ACL — from days ago. **The only new groups are the ones missing**, which is what
  finally exposed it: every `_APR_HIGHLY_CONFIDENTIAL` group (created that morning) was absent from
  every folder while everything older looked right.
- **THE TELL, worth reusing: the warnings named `"CRS Approve"`, not `"DMS Approve"`.**
  `applyPermissionPrefix` is a MODULE-LEVEL mutation performed in the same effect, so it had already
  run; `setRoleDefs` had not reached the executing closure. **Module global visible, React state not**
  — that pair distinguishes this from a genuinely missing permission level in one glance.
- **FIXED: the run resolves the levels ITSELF** (`fetchRoleDefs`, which now RETURNS them as well as
  setting state). State is used when already populated; otherwise the run fetches before touching
  anything, and the two remaining in-run lookups read the local `defs`, never `roleDefs`.
- **AND IT NOW REFUSES — the one place in this run that fails CLOSED.** An empty level list does not
  degrade the run, it silently empties it, so the run stops before making changes and says
  `Could not read this site's permission levels, so NOTHING was granted...`. **A run that grants
  nothing while reporting warnings is far worse than one that stops and explains itself.**
- **The mount effect's `if (!res.ok) return;` is gone too** — a failed read left `roleDefs` empty with
  nothing said anywhere, i.e. the identical end state by a second route, equally invisible.
- **PAGE ACLs SURVIVED, which is why nobody was locked out.** The page pass grants `Read` through the
  same lookup, so it granted nothing either — but `groupsToRemove` does not need a role definition, and
  the Group Map was complete by then, so nothing was removed. Verified: `Upload-Form.aspx` 335
  principals, `My-Submissions.aspx` 446, every admin page at 1 (Owners only). **Had the Group Map been
  incomplete at the same moment, this would have stripped page access site-wide instead.**
- **Workaround on any build before 1.0.237.0: load the page, WAIT for it to settle, then press Run.**

## A BULK LIST DELETE KEEPS RUNNING LONG AFTER THE UI GOES QUIET (2026-08-24/25)
The CRS Group Map was cleared from the list UI. It read as empty immediately — while Site Contents
still reported 2,419 items — and **SharePoint went on deleting in the background for well over an
hour**, at roughly 25 rows a minute.
- **EVERY PROVISIONING RUN DURING THAT WINDOW WAS FEEDING A SHREDDER.** Rows were written, read back as
  present (so the run reported them `already there` and declared the segment complete), and deleted
  minutes later. Four separate runs each reported `0 failed` over a list that was quietly draining.
- **THREE READS DISAGREED, AND ALL THREE WERE "RIGHT" AT DIFFERENT MOMENTS:** `ItemCount` 1984, a live
  paged count 1870, then 1821, then 1770. **`ItemCount` is a CACHED aggregate and lags in both
  directions** — it also read 2,419 for a list the view showed as empty.
- **THE LIST VIEW'S FILTER PANE ALSO SERVED STALE RESULTS**, twice, in opposite directions: it showed
  4 rows for a group that had 5, then 1 row for the same group. `/_api/.../items?$filter=...` was
  correct both times. **When a list has just taken thousands of writes, only the API is authoritative**
  — not the view, not Site Contents, not the run log.
- **THE DIAGNOSTIC THAT SETTLED IT:** `/_api/web/recyclebin?$orderby=DeletedDate desc` showed CRS Group
  Map rows being deleted seconds earlier, with `GHO_GS_CCR_APPROVER` appearing four times in ten
  seconds — its `APR, DELS, DEL, SHARE` rows going while `UPL` remained. That matched the "APPROVER
  groups hold only `[UPL]`" pattern exactly. **Ask what the data is doing before theorising about code.**
- **DO NOT RESTORE FROM THE RECYCLE BIN afterwards** — rows re-written in the meantime would become
  duplicates, which no screen displays and reconciliation would act on.
- **WAIT FOR TWO IDENTICAL COUNTS BEFORE WRITING ANYTHING.** The delete finished at 820 surviving rows
  (those written into the higher id range); one Bulk Provisioning pass per segment then restored the
  full 2,140 and the audit came back **696 groups holding rows, 2,140 rows, 0 mismatched**.
- **VERIFY BY ARITHMETIC, NOT BY THE LOG.** A segment's row count is exactly
  `units x 18 + departments x 5 + 1`: GHO 62x18+8x5+1 = **1157**, MHO 49x18+20x5+1 = **983**. Both
  landed on those numbers to the row, which is what finally established the runs were honest.
- **NEVER DO THIS ON A LIVE TENANT.** There is no need to: Bulk Provisioning is idempotent, so a
  segment is repaired by re-running it, never by clearing the list first.

## ONE PRESS OF BULK PROVISIONING NOW FINISHES THE SEGMENT (2026-08-24, 1.0.236.0)
Client, rebuilding GHO after the CRS Group Map list was emptied: *"Can't we keep running untill
eveyrthing is there?"* then *"So you expect that client to click multiple times manually? or is there a
way to automate this"*. `run` in `BulkGroupProvisioner.tsx` is a LOOP now — passes repeat until one
writes nothing.
- **⚠ THE OBSERVATION THAT FORCED IT, AND IT IS NOT EXPLAINED.** Rebuilding GHO took **four** presses to
  settle — **250 mapping rows written, then 18, then 68, then 0** — and every one of those runs reported
  **`0 failed`**. So a pass can complete cleanly while seeing less than the whole list. Worse, run 3's log
  claimed `GHO_PO_SGC_APPROVER` held only `UPL` while the list itself showed `APR, DELS, DEL, SHARE` and
  **no `UPL`** — the log and the list disagreeing in *both* directions on one group.
  - **THROTTLING IS NOT THE ANSWER, though it was the first guess and it was told to the client.**
    `withRetry` already retries 429/503 four times with backoff, and a genuine failure increments `failed`
    and prints `✗`. Neither happened. **A clean `0 failed` is not evidence a pass saw everything.**
  - **NOR IS A DEDUPE COLLISION.** `isDuplicateRow` keys on `GroupId + Scope + Target + UnitTermGuid +
    Role`, so two groups at one term cannot mask each other. Checked, ruled out.
  - **THE CAUSE IS STILL UNKNOWN.** Do not write one into this file later without evidence. The loop
    converges regardless of cause, which is why it was the right thing to build first.
- **THE FINISH LINE IS A PASS THAT WROTE NOTHING, not a pass that completed.** Only a further pass can
  establish it. Leaving that to the admin means a segment is finished when somebody happens to press once
  more, which is not a condition anyone can check — and the failure is silent, because a part-provisioned
  segment reports exactly like a finished one.
- **SAFE BY CONSTRUCTION, not by brute force.** Every group is matched before it is created and every row
  checked against the list before it is written, so an extra pass over finished work writes nothing and
  costs one read. That property already existed; the loop just stops relying on a human to exploit it.
- **⚠ THE DATA IS RE-READ BETWEEN PASSES (`readSegmentData`), never carried over.** The whole reason a
  further pass finds work is that the previous read did not describe the list correctly — reusing it would
  reproduce the same blind spot and converge on the same wrong answer, only faster. This is what forced
  `loadForSegment` to split: the reader RETURNS its data and the loop threads it pass to pass, because a
  value set into state during an async run is not visible to the closure already running.
- **BOUNDED at `MAX_PASSES = 10`, and hitting the cap is a WARNING, never a quiet stop.** A pass that keeps
  finding work forever is a defect, not a big segment. GHO settled on the fourth. The toast and the log both
  say outright that the segment may be incomplete — the one ending whose counts look like a clean run.
- **A failed re-read ENDS the run and says so.** Another pass over data that could not be refreshed is the
  previous pass again, and its "already there" count would be a guess.
- **`already` is the LAST pass's count, never a sum** — it describes the state of the list at the end, and
  adding every pass's view of it reports the same row several times.
- **Stop still lands between groups** and is unaffected: a stopped run breaks the loop immediately.
- **⚠ FIXED IN THE SAME CHANGE: the site-group read was `$top=1000`.** `spGroups.ts` was raised to 5000 on
  2026-08-21 (memory `sp-capped-read-reads-as-absent`) and **this copy was missed** — the fourth instance of
  that trap. One provisioned segment is ~320 groups, so 1000 is under three segments and **SDG's live tenant
  crosses it**. A capped read is indistinguishable from the group being ABSENT, so the run tries to CREATE a
  group that already exists, fails on the duplicate name, and **skips every one of its mapping rows with
  it** — this one WOULD have shown as `failed`, unlike the defect above. Not what bit ClarenceDMSTesting
  (585 groups, under the cap); it would have bitten SDG on the next run.

## A DELETED GROUP IS NOT REPAIRED BY RECONCILIATION, AND NOTHING SAID SO (2026-08-20, 1.0.192.0)
Proven on site by deleting `GHO_GTS_un1_UPLOADER` and re-running. **Reconciliation grants to groups
that EXIST and never creates them.** The Group Map rows survive, point at an id nothing resolves to,
and `toGrant` — built straight from those rows — never checked. The unit lost that role's access on
every library and **the run reported completely clean**.
- **The `no group-map groups for this folder` warning cannot catch it.** That fires only when
  `toGrant` is EMPTY, i.e. every group for the folder is gone. The real case is one of five deleted
  by accident, which leaves it non-empty.
- **Now reported, ONCE per folder and NAMED:** *"mapping row(s) point at group(s) that no longer
  exist: X. Nothing was granted for them. Re-create them with Bulk provisioning on Group Management,
  then reconcile again."* Named rather than counted — the name goes straight into the provisioner,
  while "2 rows" is unactionable. Once per folder because an approver group carries six rows.
- **`$top=5000`, NOT `fetchAllSiteGroups`** — its `$top=500` cap is already at 499 on this site, and a
  TRUNCATED read reports live groups as missing, sending an admin to re-create groups that exist.
- **FAILS OPEN** (`undefined` ⇒ report nothing; an empty result is unreadable, never "no groups"),
  and resolves the principal through the SAME `spGroupPrincipalId` the grant uses — a check that
  disagreed with the grant about which principal a row means would name the wrong group.
- **A SEPARATE "fix missing groups" BUTTON WAS PROPOSED AND REJECTED.** Bulk provisioning already IS
  the repair (`1 to create`, idempotent, writes the rows too); a second button would be a second
  definition of which groups should exist, and the drifting copy is always the rarely-used one. The
  gap was never the repair — it was that nothing DETECTED the problem.
- **VERIFIED REPAIR PATH: Bulk provisioning → reconciliation.** 1 created, 2 mappings written, Folder
  Access showed **2 mappings not 4** (no dead rows left behind), and the grants came back exactly:
  Approval Document 6→8, Documents 11→12, plus the browse corridor and both derived pages.
- **⚠ THE `runRecon` FLOW CARD SAID "after a group was deleted" AND THAT WAS A DEAD END** — written
  the same morning, corrected the same day. It now names the two-step order and states outright that
  reconciliation cannot bring a group back.

## STRUCTURE CHANGE + FILE MOVE — COLLISION VERIFIED IN BOTH LIBRARIES (2026-08-20)
The dangerous position, tested twice on site. **Removing a tier COLLAPSES sibling folders into one,
and it is the only position where a bug could OVERWRITE a document rather than misfile it.**
- **Staging:** two `TEST - TEST - TEST - 20-08-26.pdf` under `Archive 1` and `Archive 2` → clash
  detected, `Rebuild` **disabled** until settled, suggestion `… (Archive 2).pdf`. Moved 15, tidied 35
  empty folders, both files present, neither overwritten, `Created By` preserved.
- **Documents:** the same filename under `2024` and `2025`, both APPROVED, `Year` removed → clash
  detected as `Documents · …`, suggestion `… (2025).pdf`. Moved 15, tidied 28.
- **THE SUFFIX IS THE REMOVED TIER'S VALUE, and it generalises** — `(Archive 2)` when Archive went,
  `(2025)` when Year went. Never a number: the value is *why* the two files differed, and `(1)`/`(2)`
  throws away the only thing that distinguished them.
- **ONE file keeps the original name**; only the others are renamed. Renaming both would be gratuitous.
- **Nothing moves until every clash is settled** — a hard gate on `Rebuild`, not a warning. `Next` is
  held separately (`onMigratePendingChange`), because walking on would leave the segment half-changed
  and the closing step turns uploads back on over it.
- **Positions now verified:** add-at-bottom, remove-clean, remove-with-collision (Staging AND
  Documents). **Still untested: add mid-chain, reorder.**
- **⚠ THE `Uploader` COLUMN DRIFTS ON EVERY MIGRATION.** `Created By` survives a move correctly; the
  `Uploader` column shows whoever RAN the migration, because it is bound to **Editor**, not Author
  (memory `dms-uploader-column-is-author`). No longer theoretical — after two runs, three of four
  files in one folder were mis-attributed. Every structure change makes it worse.

## ⚠ RE-ADDING `Year` OR `Document Type` CORRUPTED THE SEGMENT'S CHAIN (2026-08-20, 1.0.195.0)
Found on the first mid-chain add run: `18 document(s) moved, **0 tagged**, 18 problem(s)`, every one
reading `Year: The data returned from the tagging UI was not formatted correctly`. The folders were
right; the metadata was never written.
- **`addTier` derived `tidCol: "<Column>Tid"` UNCONDITIONALLY.** The seeded built-in pair carries NO
  `tidCol` — deliberately, with a comment saying so — but the moment an admin REMOVES `Year` and adds
  it back through the form, it returns as an ordinary level WITH one.
- **`tidCol` is what tells every writer the column is plain text.** With one attached,
  `backfillMetadata`'s `colFor` stops skipping the tier and writes the bare label `2024` into the
  TAXONOMY `Year` column. `SubtreeMigrator` documents this exact failure and guards against it — the
  guard was simply bypassed by data that no longer looked managed.
- **IT WAS NOT ONLY THE MIGRATION.** `buildLevelFormValues` reads the same `labelCol`/`tidCol`, so
  **every new upload into that segment would fail to tag Year too**, reporting the generic *"Uploaded,
  but tagging metadata failed"*. One bad tier, two independent writers, neither obviously related.
- **`Document Type` would have been worse and is pinned by test.** Its real internal name is the
  encoded `Document_x0020_Type`; `columnNameFor("Document Type")` derives `DocumentType`, which does
  not exist — and ONE unknown field name fails the WHOLE `validateUpdateListItem` call (gotcha #4), so
  a single re-added tier costs every OTHER tier's metadata as well.
- **Fixed by `builtInTierFor` in `shared/folderChain.ts`** (pure, 7 tests): re-adding either tier by
  name restores the built-in shape verbatim. **The admin's own term set still wins** when they typed
  one — that chooses which options the dropdown offers, which is theirs — but the COLUMN SHAPE is not
  a matter of opinion, because the column already exists.
- **ONE definition**, shared with `effectiveOnDemandTiers` via `builtInOnDemandTiers`, and a test
  asserts the seeded pair and the re-added tier are equal. Two copies of this shape would drift, and
  the drifting one would be the rare path nobody looks at.
- **THE STANDING RULE: `ensureTextColumn` skips a column that already exists WHATEVER ITS TYPE.** So a
  chain may name a real column and still be wrong about what it is, and nothing at save time can tell.
  Column creation is not validation.
- **RECOVERY needs no re-migration.** The folders are correct, so re-saving the corrected chain scans
  clean and offers Apply immediately; the standalone **backfill** button then tags the files without
  moving anything. That button exists precisely because this class of failure happened once before
  (2026-08-11) and the only route to the repair was through the thing that broke it.
- **⚠ "Check document tags" CANNOT CONFIRM `Year` OR `Document Type`, AND SAID IT HAD (fixed
  1.0.196.0).** `colFor` excludes any tier with no `tidCol` — which is exactly those two — so on a
  segment whose only below-Unit levels are the built-in pair, `allCols` is empty and the pass returns
  `0 stamped, 0 failed` **having looked at nothing**. The caller read that as agreement and printed
  *"Every document's folder columns already match where it sits"*: an assertion about a check that
  never ran, on the one screen an admin uses to confirm a repair. `backfillMetadata` now returns
  `checkable`, and the message says outright that these two are managed metadata, are never written
  here, and therefore cannot be confirmed here. **Empty ≠ unknown**, in the place where the cost is a
  false all-clear.
- **⚠ REMOVING A TIER DOES NOT CLEAR ITS COLUMN, so folder and metadata can disagree afterwards.**
  The backfill never writes managed metadata, so a file that was tagged `Year = 2025` keeps that value
  while its folder collapses away and is later rebuilt as `2024`. Verified shape on
  `TEST - TEST - TEST - 20-08-26 (2025).pdf`. Nothing detects it and nothing repairs it — the value
  must be corrected on the document. State this to the client alongside "removing a level is not
  reversible": the SHAPE comes back, the values do not, and the two can end up contradicting.
- **⚠ FILES UPLOADED DURING A BROKEN-CHAIN WINDOW ARE PERMANENTLY UNTAGGED, and nothing finds them.**
  A bad value on a KNOWN column fails that field ALONE (`HasException` is per field), so the upload
  succeeds with every other column populated and one blank — the uploader sees *"Uploaded, but a field
  failed"* and clicks past it. The backfill cannot repair it (managed metadata is never written there)
  and no other pass looks. Seen live: one approved document in `Documents/GHO/GF/TAX/2024/Agreement`
  with Year blank and everything else correct. **After any structure change, filter the affected
  column for blanks before turning uploads back on** — that is the only way to find them.
- **A `YearTid` text column now exists in all four libraries** — created by the bad save, harmless,
  never read once the chain is corrected. Not deleted: a column takes its data with it and does NOT
  go to the recycle bin.

## A PAGE THAT RENDERS NOTHING IS THE WORST EMPTY STATE (2026-08-20, 1.0.194.0)
`ApprovalDocument.tsx` ended its render guards with `if (!item) return null;` — **a blank page**,
reported on site. Approving the LAST document in the queue is the ordinary way to reach it: Auto-route
copies the file to `Documents` and **deletes it from the approval library**, so the `?itemId=` the
approver arrived on stops resolving. Loading has finished, no request errored, and there is simply no
item — so an approver pressing Back after every approval met an empty screen.
- **It says NEITHER of the two things this codebase distinguishes everywhere else** — not "could not
  read", not "nothing here". `null` is the one render that answers no question at all.
- Now names the likely cause and the distinction that makes it checkable: an APPROVED document has
  moved to Documents, a REJECTED one would still be here.
- **The general rule: `return null` in a render guard is a bug waiting to be reported.** Every other
  branch in this file already returns a message (`loading`, `fetchError`); this one was the gap.

## ⚠ A 503 ON ONE TERM MADE RECONCILIATION CALL SIX LIVE UNITS STRAYS (2026-08-20, 1.0.191.0)
Found on a routine scoped GHO run. `Term children c76a5886-… returned 503` — one read, one department
(President's Office) — and the stray pass then reported **all six PO units in all four libraries** with
*"no live term maps to this folder… Move them somewhere real, then delete the folder."*
- **THE SAME DEFECT AS THE 2026-08-19 SCOPED PRUNE, ON A DIFFERENT AXIS — and the comment added by
  that fix asserts safety on the wrong one.** 1.0.186.0 says the pass is safe because it *"DESCENDS
  FROM `targets`, so on a scoped run it only ever looks inside the segments that were walked"*. True
  for SCOPE. It says nothing about `targets` being **short INSIDE a segment that WAS walked** — which
  is exactly what an incomplete term read produces, because `walk` keeps the targets gathered before
  the failure. PO stayed in `targets` as a department; its six units did not; the pass descended from
  the department and found six folders it did not recognise.
- **NOTHING WAS DAMAGED, AND ONLY BY LUCK.** Quarantine breaks inheritance and restores site Owners —
  which would have taken those six units' access away. It did not fire, because `already =
  getHasUniquePerms(...) === true` and **every real unit folder already has unique permissions by
  design**, so the `ALREADY QUARANTINED` branch reported and `continue`d. A legitimately locked folder
  and a previously quarantined stray are indistinguishable to that check, so it must never be what
  stands between a term-store hiccup and a wrongly secured folder.
- **`buildProvisionTargets` now returns `incompleteSections` — bare `stagingFolder` keys, separate
  from `incomplete`'s human sentence** (`"GHO — …503"`). Anything that has to ACT on incompleteness
  needs the segment, and parsing it back out would make the guard depend on the wording of an error.
- **Gated PER SEGMENT, not globally**: a 503 in GHO must not stop MHO's strays being quarantined, or
  one flaky read disables the pass site-wide. **The skip is REPORTED**, naming the segment — a pass
  that quietly checked less than the admin believes is how "no strays reported" becomes "no strays".
- **THE GENERAL RULE, now twice learned: there are TWO independent ways a target list can be
  incomplete — the run was SCOPED, or a read FAILED — and a pass that judges "does a live term claim
  this?" must be guarded against BOTH.** The Folder Map prune already was (`fullCoverage` and
  `incompleteSegments` are separate branches); the stray pass was guarded against neither, then
  against one.

## A FLOW ABOUT ONE NEW TERM MUST NOT TICK OFF SEGMENT-LEVEL FACTS (2026-08-20, 1.0.189.0)
Found the first time anyone walked **Add a department or unit** on a provisioned segment. The rail showed
**Group Management, Folder Access and Folder Reconciliation all `Done`** before the admin had done
anything — on the flow whose entire purpose is adding something that does not exist yet.
- **THE FACTS WERE TRUE AT THE WRONG GRANULARITY.** `groupsExist`, `folderAccessRows` and `foldersExist`
  each ask *"does this SEGMENT have any at all"* — any site group with the segment's code prefix, any
  Group Map row, any Folder Map row. GHO has **308 groups, 790 mapping rows and 67 folder rows**, so all
  three answer yes the moment the segment is picked. Right question for flow 1, where the answer starts
  false and turns true as the flow is walked; wrong question for flow 2.
- **SAME SHAPE AS THE SCOPED-RECONCILIATION DATA LOSS** (`coversEverySegment`, 2026-08-19): a question
  that is correct site-wide and wrong the moment something narrows the subject. **Whenever a flow, a run
  or a read narrows what it is about, re-ask whether the existing facts still answer the new question.**
  Memory `feedback-audit-consumers-when-narrowing-data`.
- **DROPPED, not re-scoped**, because the honest answer is available and the correct one is not.
  Scoping `groupsExist` to a new unit needs its abbreviation — which does not exist until a later step of
  this very flow; scoping `foldersExist` needs the new term's GUID, known only when the OPTIONAL subject
  box was filled in. `unknown` renders **"Not checked"**, already this codebase's first-class answer for
  a fact it cannot establish. **Understating costs a glance at a screen; overstating tells an admin they
  are finished when nothing was provisioned.**
- `scopeFactsToFlow` in `shared/folderFlows.ts` (pure, 9 tests), keyed on `asksSubject` being
  `add`/`rename` — no new field, and the discriminator already meant "this flow is about one term inside
  an existing segment". Applied ONCE where the facts reach the rail, so no caller has to remember.
- **`abbreviationsMissing` is deliberately NOT dropped** — the abbreviation screen reports it live over
  the whole tree, so a newly added term with no code makes it non-zero and that step correctly reads
  `todo` **and holds Next**. It is the one fact here that was already subject-aware. `segmentExists` stays
  too: the segment genuinely does exist.
- **⚠ A STALE TICK IS A SEPARATE FAULT AND LOOKS IDENTICAL.** The count is reported UP from the
  abbreviations screen, so it lags a term added since that screen last read. Seen the same afternoon.
  When a tick looks wrong, establish which of the two it is before changing anything — the fix for one
  does nothing for the other.
- **THE SUBJECT BOX ASKS FOR A NAME, NEVER A GUID, and now says so.** It read *"What are you adding?"*
  over a text box on an `outside` step — which reads as *type it here and I'll add it*, so an admin could
  type a name, press Next and believe the term had been created. Now past tense (*"Which department or
  unit did you add to Group Head Office?"*), with **"This box creates nothing"** stated outright.
  - **The segment is NAMED when one is chosen, never ASKED for.** Adding a segment question here walks
    straight back into the 2026-08-19 bug `stepUsesSegment` exists to prevent.
  - **A GUID was considered and REJECTED.** It breaks the abbreviation page's founding rule (the page
    never shows a GUID — it is obtainable only from the term-store properties panel, one term at a time),
    it would put the hardest-to-get value in the system in front of an **optional** convenience, and a
    mistyped GUID is indistinguishable from a nonexistent term with no near-match hint possible.
  - **The `/` worry is unfounded**: `normaliseLabel` folds fullwidth `＆`, case, whitespace runs and
    zero-width characters, and compares `/` literally on both sides — both of which come from what
    SharePoint displays. `Estate/Mill` matches. Sanitizing strips `/` only when deriving a COLUMN name.
- **A CHILDLESS DEPARTMENT IS A LEAF, and the log used to call it a UNIT (1.0.190.0).** `walk` returns
  `children.length > 0` and the caller sets `isLeaf = !that`, so a department with no units yet comes
  back as the upload target. Verified on site: `GHO/GTS` was created in all four libraries, warned
  `no group-map groups`, and `Manage Access` on it showed **only the two Owners groups** — locked
  admin-only, exactly right. The warning now says **"for this folder"**: the tier name is
  segment-specific (Unit / Estate-Mill / Refinery), so naming any one of them is wrong somewhere, and
  it sent an admin hunting for a unit that did not exist.
  - **Group provisioning and reconciliation disagree about what it IS, and both are right.** The
    provisioner plans a childless department as a **department** (one `_HOD`, verified: `1 to create`),
    while reconciliation treats it as a **leaf**. The end state is safe because a folder with no
    mapped groups is locked rather than opened.
- **THE ABBREVIATION FOOTER NAMED NO ROW (fixed 1.0.190.0).** It said *"N existing codes changed —
  live folders will be renamed"* while the `was X → Y` detail appeared **only in the save summary,
  i.e. after committing**. An admin was warned a live folder would be renamed and given no way to find
  which one. Rows now carry `RowProblem.note`, rendered inline as `✎ was TAX — …`.
  - **Attached in a SEPARATE pass**, because every branch of `validateRows` `continue`s — the first
    build skipped the marker on blank, unusable and **colliding** rows, the last being exactly the row
    an admin is staring at when they most need to know it also renames a live folder.
  - **CLEARING A SAVED CODE IS A REMOVAL, NOT A RENAME**, and `changedRows` counts it among the
    renames because it only asks whether the value moved. Consequences are opposite: a rename moves a
    live folder, clearing leaves it where it is and stops reconciliation managing it, so the unit
    silently loses its upload path. Its own message now says so. Found on site 2026-08-20 when a
    cleared code the admin had forgotten about produced an unexplained rename warning.
- **KNOWN, NOT FIXED: `labelMatches` asks "does a term with this name exist in the segment", not "under
  the right parent".** `Tax`, `Legal` and `PM` each exist under several GHO departments, so adding a
  `Tax` unit under a new department ticks step 2 off Group Finance's existing one. Advisory only — it
  gates nothing, so it cannot produce a wrong outcome, only a wrong reassurance. The fix is matching on
  the parent chain, which `shared/termChains.ts` already builds for Folder Access's tier labels.

## THE STRUCTURE FLOW, AS THE CLIENT ACTUALLY WALKED IT (2026-08-19, 1.0.182-183)
First end-to-end run of "Change the folder structure" by a person following the rail. Seven findings,
all reported in one sitting, and none of them visible from the code.
- **⚠ EVERY LINE IN THE MIGRATION SCAN IS A FOLDER, NOT A FILE — and it was read as files.** The
  document count was appended only when NON-ZERO, so an empty leaf showed no count and looked like a
  phantom entry, and nothing on the line said "folder". An empty leaf is the NORMAL case: its
  documents were approved and routed away, leaving the folder behind, which still has to be re-placed
  under the new chain. Every line now states `folder`, the count **including `empty`**, the full
  server-relative path, and the document names.
- **The facts are re-read on STEP CHANGE (`stepIdx` in the effect deps).** Saving on the levels step
  writes `PendingLevels` — exactly the fact the migrate step is locked behind — but the facts were
  read when the flow opened, so step 4 said *"Not ready yet"* beside a screen that had just saved, and
  step 3 stayed *"Not checked"* until a manual refresh. **Fourth instance in this feature of a screen
  reading at mount and lying about a run beside it.**
- **The segment picker no longer stands in front of the Folder levels step** — that screen lists every
  segment with its own Edit button, so the picker asked for a choice the screen then asked again, and
  picking changed nothing visible. Same class as the `outside` fix the day before; the gate is now
  `stepUsesSegment`, a rule with a test, not a list of screen kinds.
- **Both pause steps tick, and in OPPOSITE directions** — `uploadsPaused` true means step 1 is done and
  step 5 is not. A rail that ticked the closing step while the site was still paused would confirm the
  wrong thing. `undefined` leaves both unknown; it must never claim uploads are on.
- **The migration scan now holds the flow's navigation**, via `onMigrateRunningChange` into the SAME
  `runBusy` padlock as reconciliation and bulk group runs. Scan and move both live in the page with no
  resume, so Next during "Checking…" threw the scan away silently. One reason to hold, one
  implementation, one message.
- **Unsaved level edits gate Next, with their OWN reason** (`onStructureDirtyChange`), not folded into
  the run padlock: the fix is a click away on the same screen, and *"a run is in progress"* would send
  the admin hunting for something that is not happening. Worse than the tab-switch case it already
  guarded — the next step reads `PendingLevels`, which an unsaved edit has not written, so walking on
  reports "nothing to move" for a change the admin believes they made.
- **User-facing `DMS` is now `CRS`** in both upload web parts, the allowed-file-types messages and the
  pause message. ⚠ **That pass MISSED four strings, found and fixed 2026-08-23** — the chain-error
  toast in BOTH upload web parts (*"check the DMS Config mode row"*), the upload form's not-ready
  empty state (*"Your DMS administrator… DMS Term Abbreviation list"*) and Bulk Upload's no-folder
  error. `console.error` text still says DMS deliberately: developer-facing, and it matches the
  code's internal naming. **A rename pass that reports itself done is worth re-grepping** — this one
  claimed the strings a client reads were all converted, and four of them were not.
- Still open, deliberately: **editing a level in place** (remove-and-re-add works — a feature, not a
  defect).

## UPLOADS CAN BE PAUSED SITE-WIDE WHILE THE STRUCTURE CHANGES (2026-08-19, 1.0.180.0)
Spec `2026-08-19-upload-pause-design.md`; rule in `shared/uploadPause.ts` (pure, 8 tests); screen is
`userAccess/components/UploadPauseToggle.tsx`. BUILT, **not site-tested**.
- **The flow's only preparation step was "Power Automate (optional)", and Clarence's point stands:**
  *"they are not going to go to Pause power automate."* A file uploaded DURING a migration lands in
  the OLD shape — correctly, since `Levels` is still the old chain — and if it arrives after its
  folder was scanned it is never moved. The staged-apply guard keeps that from going live (the
  pending chain applies only when a FRESH scan finds no drift), so the cost is a migration that will
  not finish rather than documents silently misfiled. On a live site that loops.
- **SITE-WIDE, not per segment** (client's decision): *"after work hours no one is going to upload,
  once its ready and file is correctly moved then they turn it back on manually."* Accepted because
  of WHEN the work happens, not because the blast radius is small — restate that whenever someone
  proposes migrating during the day.
- **TWO checks, and the second is the one that matters.** At mount for the banner; **re-read
  immediately before writing**, because settings are read once in a mount-time effect (gotcha #10)
  and a tab opened before the pause would otherwise upload straight through it. Same shape as the
  stale-chain guard.
- **⚠ IT CANNOT REACH A SESSION THAT NEVER TALKS TO THE SERVER.** The write-time check refuses the
  document rather than misfiling it, but a tab open beforehand holds its own state. **Do not present
  the toggle as complete protection** — the operational answer is the client's, structure changes
  outside working hours. Said on the screen itself.
- **The VALUE fails open; the SCREEN's claim about it fails closed.** Missing row, blank, unrecognised
  or unreadable ⇒ **not paused**, because a wrong `true` blocks every uploader on the site over a
  transient read while a wrong `false` costs one file the fresh-scan guard catches (gotcha 10b again).
  But the toggle screen shows an ERROR rather than "uploads are on" when it cannot read: an admin who
  believes uploads stopped when they did not will migrate on top of live traffic.
- **Turning it back on is its own step at the END of the flow**, because that is the half that gets
  forgotten — and a forgotten pause is a DMS that quietly accepts no documents behind a banner that
  makes it look deliberate. Resuming writes `no`, **never a blank cell**: blank cannot be told from
  never-set, and the point of the step is that an admin can confirm they did it.
- **`FolderManager.tsx` is untouched** — it holds the folder-creation and group-assignment code frozen
  after the 2026-08-19 verification, and this needed nothing from it. The screen mounts as a
  `kind: "component"` step, like Group Management and Folder Access.
- **⚠ THE SEGMENT PICKER STOOD IN FRONT OF IT ON THE FIRST RUN, and that is the SAME BUG AS THE DAY
  BEFORE.** The picker was gated on `st.screen.kind !== "outside"` — the literal added on 2026-08-19
  when the client reported it in front of an instruction-only step. The upload pause is site-wide and
  is a `component`, so the very next step added to a flow walked straight back into it. **The test is
  now `stepUsesSegment` in `folderFlows.ts`** (pure, 3 tests, one pinning both pause steps), because a
  rule survives the next addition and a list of screen kinds does not. Anything that does not consume
  a segment belongs there, whatever it renders.

## APPROVAL LIBRARY ACCESS, SITE ACCESS, PAGE ACCESS ARE NOW READ-ONLY (2026-09-02)
Client: *"the group assignment is all done by Folder recon, and how a group is have access to
anything is also done by folder recon, so I am see there is no point for all this pages to allow
client to manually strip or assign group to a page or library... I wish you can make this pages read
only."* Spec `docs/superpowers/specs/2026-09-02-access-pages-read-only-design.md`. **BUILT, not yet
site-tested.**
- **EVERY WRITE ACTION IS GONE FROM ALL THREE PAGES** — Add/Allow, Remove/Revoke, per-person removal,
  "Reopen to everyone", Site Access's add/remove person. Each keeps its READ-ONLY display (live-ACL-
  vs-Group-Map comparison, drift banners, member counts, policy explanations) and gains a banner
  pointing at Group Management for making changes, resolved via the same `resolveLink`/`readSitePages`
  pattern the retired Folder Access signpost already used — never hardcoded, since this client renames
  every page at import.
- **EACH REMOVED ACTION WAS CHECKED AGAINST WHETHER RECONCILIATION ALREADY DOES IT, AND SOME DO NOT —
  the client's own answer settled every one of them:**
  - **Revoking a live library grant** has no reconciliation equivalent (recon only ever ADDS library
    ACLs) — covered by the standing "delete the group, recreate it" repair pattern on Group
    Management, which the client already uses for persona changes.
  - **"Reopen to everyone" on Page Access** is gone with **no replacement of any kind** — the client's
    own words: *"when client locks a page, they won't want the page to be reopen for normal clients
    to see, that beats the whole purpose."*
  - **Site Access's add/remove person** — `CRS_SITE_MEMBERS` is an ordinary group, already in Group
    Management's own list, whose `GroupMembersEditor` already adds/removes people from it.
  - **Per-person removal (`accessMemberUi`'s dialog)** — same underlying `removeGroupMember` call
    exists on Group Management. The one thing that dialog added — checking whether a person is ALSO
    in a second group reaching the same library/page — is **not replaced with new UI**: removing
    someone from ANY group ends whatever that group granted, and reconciliation never leaves a stale
    grant for a group they are no longer in.
- **`accessMemberUi.tsx` IS TRIMMED, NOT DELETED.** `MemberRemovalDialog`/`MemberRows`/`removalFor`/
  `doRemoveMember` are gone (confirmed unused anywhere else via grep before removal); `MemberSummary`/
  `useGroupMembers` stay, since both pages still show a read-only member COUNT per group. The pure
  logic module behind the removed dialog (`shared/accessMembers.ts`, 28 tests) is left alone
  deliberately — self-contained, SPFx-free, costs nothing to keep, and trimming it would mean deleting
  its own test file too for no benefit to this change.
- **GROUP MANAGEMENT GAINS THREE LINK BUTTONS** beside "Select all N shown" — Site Access →, Approval
  Library Access →, Page Access → — so looking at the group list, an admin can jump straight to the
  detailed library/page/segment breakdown for any group before deleting or editing it. `SITE_ACCESS_
  LINK`/`LIBRARY_ACCESS_LINK`/`PAGE_ACCESS_LINK` are now EXPORTED from `shared/adminPages.ts` (used
  both there, in `CARDS`, and here) rather than existing as a second copy of the same `match` regex.
- **No schema change, no reconciliation change, no migration.** Every Group Map row and every live ACL
  is untouched — this removed UI only. `tsc --noEmit` clean, full suite **1597/0**.

## APPROVAL LIBRARY ACCESS AND SITE ACCESS ARE NOW SIGNPOSTS — ONE DAY AFTER GOING READ-ONLY (2026-09-02)
Client, the same day the read-only pass above shipped: *"Approval Library Access, Site Access is not
needed, its confusing them as they already know that adding a user in the group from Group Management
it will automatically allow them have access to site and library. Just keep Page access."*
- **THE READ-ONLY PASS ANSWERED "CAN THEY BE MADE SAFE" AND THE CLIENT'S REAL QUESTION WAS "ARE THEY
  NEEDED AT ALL."** Removing every write action left two pages that could only ever CONFIRM what Group
  Management already states, which is exactly what made them confusing rather than merely redundant —
  a page an admin cannot act on, showing a fact they already know, reads as broken or as a second
  source of truth to reconcile against. Page Access is kept because it is NOT redundant with anything:
  it is the only place that names which admin PAGES a group can open, which Group Management does not
  show.
- **SAME TREATMENT AS `FolderAccessPage.tsx` (2026-08-23) — kept, not deleted.** Both web parts
  (`ApprovalLibraryAccessPage.tsx`, `SiteAccessPage.tsx`) stay registered in `componentIds`; the
  `Approval-Library-Access.aspx` and `Site-Access.aspx` pages stay on the site. Removing a component
  from `componentIds` is the change that has already bitten this project twice, both silently — CRS
  Requests undeployable for weeks, the `+ New Folder` customizer inert for months — and an admin may
  have either URL bookmarked or in site navigation. `pageAccessPolicy` still locks both, so nothing is
  newly exposed. `StagingAccess.tsx` and `SiteAccess.tsx` themselves are UNTOUCHED — the pages simply
  stop mounting them.
- **BOTH REWRITTEN AS A SHORT STATIC MESSAGE + A RESOLVED LINK TO GROUP MANAGEMENT**, reusing the
  `resolveLink`/`readSitePages`/`CARDS` pattern every other cross-page link in this project already
  uses — never hardcoded, since this client renames every page at import. A failed page-list read
  leaves the link out and falls back to plain text naming the page, rather than a dead `<a>`.
- **`shared/adminPages.ts`: `SITE_ACCESS_LINK`/`LIBRARY_ACCESS_LINK` ARE DELETED, NOT LEFT UNUSED.**
  A parked-but-exported const inviting a future re-wire is the exact trap `spGroups.ts`'s deleted
  `inviteToGroup` comment already warns against — kept alive, it is only a matter of time before
  something links to it again. The "access" card's `links` array narrows to Group Management + Page
  Access, and its `blurb` drops "sites, approval libraries" to describe only what remains.
- **THE THREE LINK BUTTONS `GROUP MANAGEMENT` GAINED ON 2026-09-02 (the read-only pass, immediately
  above) LOSE TWO OF THREE THE SAME DAY.** "Site Access →" and "Approval Library Access →" would now
  point at pages that are themselves signposts BACK to Group Management — a circular link with no
  destination. "Page Access →" survives; it is the one report that still says something Group
  Management does not.
- **TWO EXISTING TESTS IN `adminPages.test.ts` NAMED THE DELETED LINK KEYS AND HAD TO BE REWRITTEN,
  not merely deleted** — one asserted `libraryAccess` resolves to the right page and not the approver's
  queue, folded into a comment recording why that collision existed; the other, testing that
  `pageAccess`/`siteAccess`/`groups` do not collide with each other, keeps only the `pageAccess`/
  `groups` half.
- **No schema change, no reconciliation change, no migration.** Every Group Map row and every live ACL
  is untouched — this removes navigable UI only. `tsc --noEmit` clean, full suite **1596/0**.

## GROUP MANAGEMENT ANSWERS "WHAT CAN THIS PERSON REACH?" (2026-08-23, 1.0.232.0)
Client, on why the access pages feel redundant: *"I notice that we dont even need to give them the Site
Access page and Folder Access and Approval Library Access, since its all done by Group Management."*
They are right — bulk provisioning creates the groups AND writes their Group Map rows, and
reconciliation grants the ACLs from those rows, so nothing about those three pages is on the critical
path any more. Rules in `shared/userAccess.ts` (pure, 18 tests); UI in `GroupManager.tsx`. BUILT and
site-verified.
- **THE LOOKUP REPORTS PERSONAS AND PLACES, NEVER PERMISSION LEVELS.** Answering *"CRS Upload on
  Approval Document"* was considered and rejected twice over: that lives in `LIBRARY_ROLES` /
  `ROLE_TO_PERMISSION` inside `FolderManager.tsx`, and `ROLE_TO_PERMISSION` is **MUTATED at runtime**
  by `applyPermissionPrefix()`. A second copy would drift; importing the live one would make the
  answer depend on whether `FolderManager` had mounted — the render-time `libApiTitle()` trap in a
  new place.
- **`personaForRoles` matches the role set EXACTLY, never a subset**, reusing `roleSetKey` from
  `bulkGroups.ts` rather than re-fingerprinting. `employee` is `["MEMBER"]` and `employee_hc` is
  `["MEMBER","MEMBERHC"]`, so a subset rule would present an **HC-cleared** viewer group as the plain
  one. An unmatched set says so out loud (*"these roles match no persona exactly"*) — a real finding:
  a group holding `ENTRY` alongside an approver's roles matches nothing, and that is worth seeing.
- **Joined on `groupId`, never the group NAME.** A rename keeps the id and leaves the stored
  `GroupName` stale (the 1.0.162.0 bug), and two groups can be renamed alike.
- **A group with no rows is a FACT, not an error** — legitimate since 2026-08-14, and the single most
  useful thing this page can tell someone asking why a person cannot get in. The **site-entry group is
  always in that state** and is flagged separately, so it renders as a hint: red there reads as
  something being wrong with the person's access when nothing is.
- **The tier chain is deliberately NOT resolved (v1).** The group NAME already carries it in the same
  abbreviations the folders use (`GHO_GF_TAX_UPLOADER` is GHO / GF / TAX), and walking each segment's
  tree costs ~8 requests per segment for a second copy of what the title says. `summarizeUserAccess`
  still returns the field, so rendering it later needs no change at the call site.
- **⚠ THE SEGMENT DROPDOWN WAS EMPTY BECAUSE `loadModes` DID NOT AWAIT `primeNames` — the 1.0.207.0
  race again, in a new file.** `cachedListTitle` answers the LEGACY `DMS Config` until priming
  settles, which 404s on a CRS-renamed site. Every screen that builds a list URL must await
  `primeNames` **inside its own reader**, not rely on a mount effect elsewhere having got there first.
- **The segment filter comes from the MAPPING ROWS, not the name**, so it agrees with what
  reconciliation reads; it is offered only when the Group Map was readable, because a filter built
  from a failed read would silently claim every group belongs to no segment.
- **Export writes what is SHOWN**, filters included — an export that ignores the filter is how a
  384-row spreadsheet gets sent as an answer about one segment.
- **⚠ The lookup rows must NOT reuse `s.row`.** That style is `display:flex` for the LIST's
  name + badge + Delete line, so it laid five paragraphs side by side at ragged widths (client:
  *"the design is not tidy"*). A lookup row is a STACK — `lkRow`/`lkTop`/`lkName`/`lkSeg`/`lkPersona`/
  `lkSum`/`lkChips`. **`s` is a `Record<string, CSSProperties>`**, so a key that does not exist yields
  `undefined` and the element renders unstyled with a green build.
- **The person is identified by EMAIL ALONE** (client: *"I think just showing the email is good
  enough"*). `name · email` read as two different people where the display name is a mangled
  directory form of the same address. Display name is the fallback, never a blank line — a guest
  account can have no email.
- **⚠ THE PICKER RESOLVES A TYPED ADDRESS TO ITSELF, so a mistyped one looks like a real person.**
  `searchTenantPeople` runs `AllowEmailAddresses: true` / `AllowOnlyEmailAddresses: false` — needed,
  because a share recipient is usually somebody with no account yet — so anything shaped like an email
  is offered as a candidate whether or not it belongs to anybody. The lookup therefore reported
  `clarencechojinheng@gmai.com` as *"not a member of this site yet"*, which invites an admin to go and
  add a person who does not exist. It now names BOTH causes (spelling, or never given access), because
  from the outside they are indistinguishable. Same trap behind the add box's
  `does not exist or is not unique`.
- **Both search boxes have a Clear button** (client, 2026-08-23, twice: *"I cannot clear once I
  entered, would be confusing UX design for client"*). Selecting a person clears the LOOKUP box by
  design — the result below is the answer — which left no way back to an empty panel, so the last
  person looked up stayed on screen for the session. `clearLookup` empties the box, the candidates, the
  person AND their answer; Escape does the same.
- **Matched on LOGIN NAME, not email**, and merged across duplicate accounts: a guest can exist
  **twice** for one address (2026-08-18), and an account may carry no email at all.
- **⚠ FOLDER ACCESS IS RETIRED (2026-08-23, client: *"I think we can remove folder access and also in
  the step, reason being I separated Folder Access to allow client to add the user separately but since
  the Group Management will be doing most of the job of bulk group and allow client to add user on
  their own, that makes Folder Access redundant."*)** Both its jobs moved: mappings are written when a
  group is created (or in bulk), and membership is edited in Group Management's own group list.
  - **The `folderAccess` FLOW STEP IS GONE** from *Add a new segment* and *Add a department or unit* —
    it existed to "add the people", which is now the same screen a click earlier. `FlowFacts` lost
    `folderAccessRows` with it, and the `screen.id` union no longer carries the value, so a new
    component step is a **compile error** in `FolderAdmin` rather than a silent fallthrough. The Groups
    step's hint absorbed the rule that mattered: **membership must never gate Next** — intent is not
    checkable, and an empty group is a valid end state.
  - **The CRS Settings card entry is gone**, pinned by a test that no card links `folderAccess`.
  - **⚠ THE WEB PART STAYS REGISTERED, AND THE PAGE BECOMES A SIGNPOST** — *"Folder Access has
    moved"*, with a resolved link to Group Management. **Do not delete the component id yet.** Removing
    a component is the change that has bitten this project twice, both times silently (CRS Requests
    undeployable for weeks; the `+ New Folder` customizer inert for months), and a package that deletes
    a web part sitting on a live page renders that page **broken with nothing saying why**. An admin may
    also have the URL bookmarked or in site navigation. `pageAccessPolicy` still locks the page, so
    nothing is newly exposed. Delete the page and the registration together, after the migration, as
    one reviewable change.
  - The signpost resolves the Group Management address **from Site Pages, never hardcoded** — the
    client renames every page at import, and a literal would be a dead link on the one page whose only
    job is saying where to go. `readSitePages` is now EXPORTED from `backToSettings.tsx` and shared,
    rather than copied.
  - **`GroupMapBuilder.tsx` is now mounted NOWHERE** — both its halves are unreachable (`show="form"`
    since this morning, `show="members"` since this change). Kept as the parked implementation of
    hand-mapping; re-mount if the two cases it covers ever bite.
- **THE LIST COUNTS PEOPLE, NOT MAPPINGS (2026-08-23, client: *"instead of showing mappings only,
  show how many users are there in each group, client doesnt understand what is mapping"*).** A
  mapping count is an internal fact about the Group Map; the question an admin arrives with is who is
  in the group. The badge reads `3 people` / `nobody in it yet`, and the mapping count survives only
  in the CSV. **`not mapped` is now `grants nothing yet`** — same fact, said in the words of someone
  who has never read this file, and shown only in the state that matters.
- **The CSV gained `People` and `Members`** (client: *"the Excel doesnt show who is inside each
  group"*). It listed what a group GRANTS and never who held it, so it could not answer *"who has
  access to Tax"* — the only question anyone opens it to ask. Count first, so the sheet sorts on it
  without splitting the names.
- **⚠ BOTH COME FROM ONE REQUEST: `fetchAllGroupMembers`, `sitegroups?$expand=Users`.**
  `getGroupMembers` is per group — 586 requests, which is not an export, it is an outage. It answers
  **`undefined` on failure, never `{}`**: the badge then falls back to the mapping count and the CSV
  writes `not known`, because `0 people` on 586 rows reports the whole site as empty, and a
  spreadsheet saying a unit's approver group has nobody in it **gets acted on**. `$top=5000` for the
  reason `fetchAllSiteGroups` learned on 2026-08-21.
- **The member read never fails the page.** It runs last in `reload` and cannot throw, so a heavy
  `$expand` that throttles costs the counts and nothing else — the list and the create form are
  unaffected.
- **`GroupMembersEditor` gained an OPTIONAL `onChanged`**, so the count re-reads after an add or
  remove. It calls `reloadMembers`, not `reload`: re-reading 586 groups and every Group Map row to
  answer a question about one person is what makes a page feel broken. Folder Access shows no count
  and passes nothing. **Fifth instance of "a screen that reads a list at mount lies about any write
  beside it."**
- **⚠ A FAILED ADD IS AN INLINE ERROR, NOT A TOAST** (client: *"when an email can't be found then show
  an error not a toast with error client cannot understand"*). It was showing SharePoint's own
  `HTTP 404 {"error":{"code":"-2130575276, Microsoft.SharePoint.SPException","message":"The user does
  not exist or is not unique."}}` beside the box the address was typed into, then clearing it after
  seven seconds. `explainAddFailure` translates it: an error about ONE field belongs beside that
  field, and must name the fix rather than the exception.
  - **`does not exist or is not unique` means TWO different things** and the message says both,
    because the admin cannot tell which from outside: the address belongs to nobody (a typo, or
    somebody outside the organisation never invited to the site), or it matches **more than one
    account** — which happens here, since a guest can exist twice for one address (2026-08-18). The
    fixes differ: check the spelling / invite them once / pick from the list instead of typing.
  - **Anything unrecognised keeps the raw text on the end.** An unknown error must stay visible.
- **The add toast now says what they can REACH**, not just which list they joined (client: *"the toast
  should tell them they also have access to the folder"*). `added to <group>` reads as a membership
  change; an admin who does not realise it hands over FOLDER access in every mapped library has just
  granted documents without knowing it. The hint under the box already said so — now it is said at
  the moment the change happens.
- **ADDING A PERSON IS BACK ON THIS PAGE (later the same day).** Expand any group in the list to add
  or remove its members. `GroupMembersEditor` is **MOUNTED, not copied** — Folder Access still mounts
  the same component, so there are two mount points and ONE implementation. That is what the
  2026-08-18 warning was actually about: two *lists* of the same people drift, two mounts of one
  component cannot.
  - **ONE group open at a time**, so 586 rows cost one member read.
  - **⚠ THE 60vh SCROLL CAP IS LIFTED WHILE A GROUP IS OPEN**, and that is not cosmetic: the people
    picker is absolutely positioned, so a scroll container clips its results for any group near the
    bottom — a long page traded for a control that silently cannot be used. The comment describing
    this rule had survived in the file since the editor was removed; the behaviour had not.
  - **Every string naming Folder Access as the place to assign a mapping is GONE** (the green intro
    box and both create toasts). Folder Access stopped creating mappings on 2026-08-18 and the
    hand-mapping form left this page on 2026-08-23, so *"assign it on the Folder Access page"* sent an
    admin to a screen that **cannot do it** — and they would go on believing the group was one step
    from working. With no persona there is now no mapping screen anywhere: the message says outright
    to delete the group and re-create it with one. A partial mapping now says **press Create again**,
    which is true (the write is deduped on the row) where naming another page was not.

## CRS Settings — the admin landing page (2026-08-14, spec `2026-08-14-crs-settings-landing-page-design.md`)
Client: they *"do not know what to do or how to operate"*. There were **14 web parts and no menu**.
Web part **`CRS Settings`** (`4d8b7e21-…`, own bundle), built to the client's mockup; rules in
`shared/adminPages.ts` (pure, 33 tests), icons in `crsSettings/components/icons.tsx`, design sources in
`docs/design-assets/`. BUILT, not yet site-tested.
- **It is a DIRECTORY and nothing more, by the client's decision** (*"don't worry about how to operate
  it, I got an idea"*). A "Common tasks" band carrying the real workflows was offered and declined.
  Worth restating: a grid answers *where*, and the failure that bites is *order* — miss the abbreviation
  step and reconciliation creates nothing, silently. **If their operating idea does not land, the order
  belongs on this page.**
- **Three cards the mockup lacked were added**, and one matters: **Group Management leads the access
  card**, because Folder Access maps an EXISTING group, so an admin starting there has nothing to pick.
  Also **CRS Audit Log** (on no menu at all before — nobody would find it) and **Bulk Upload**
  (`adminOnly` per `pageAccessPolicy`). **CRS Mapping is dropped**; abbreviations live under Folder
  Management, where the Folder Administration tab bar already had them.
- **LINKS ARE RESOLVED FROM SITE PAGES AT RUNTIME, NEVER HARDCODED.** The client renames everything at
  import (memory `dms-to-crs-rename-pending`), and a hardcoded `Folder-Administration.aspx` fails as a
  **dead link** — no error, no clue — on the one page whose job is saying where to go. Uses `FileRef`,
  not a path built from the library name (gotcha #12).
- **The patterns are load-bearing and pinned by test.** A bare `/folder/i` claims **Folder-Access.aspx**,
  so "Folder Reconciliations" would open a permissions screen; a bare `/approv/i` claims
  **ApprovalDocument.aspx**, the approver's queue. Hence `folder.?admin|folder.?manage|folder.?structure`
  and `approval.?library|library.?access`.
- **Three link states.** `resolved`; `ambiguous` → navigates to the SHORTEST file name (a duplicate is
  the longer one) and names the others, because silently picking is how a half-renamed page sends people
  somewhere nobody meant; `missing` → the row is **disabled and names the page to create**, never a dead
  arrow. **An unreadable Site Pages list is NOT "no pages exist"** — it says so, or it would tell the
  admin to create ten pages that already exist.
- **The "create a page called…" advice names the SHARED page** for a row that is only a tab of one
  (`pageName`). Deriving it from the label would say `Folder-Reconciliations.aspx` — a page the pattern
  can never match, leaving the row dead after the admin did exactly as told. A test asserting every
  suggestion resolves caught this on its first run.
- **Folder Management's three rows deep-link with `#tab=<slug>`**, read by `FolderManager` on mount —
  they are TABS of one five-tab page, so without it two of three look broken. `tabFromHash` has ONE
  implementation, shared by the writer and the reader. **A slug is a public name once shipped** (an admin
  may bookmark it), so `DEEP_LINK_TABS` maps slug → `Tab` explicitly; renaming a `Tab` must not break a
  bookmark. An unrecognised hash falls back to the default tab, never a blank screen.
- **EVERY PAGE IT LINKS TO CAN GET BACK (2026-08-15, client: *"I can't go back to CRS Settings"*).**
  A directory whose destinations are one-way trips sends the admin to the browser Back button or the
  site nav. `shared/backToSettings.tsx` owns the band and the destination for all nine pages.
  - **The target is RESOLVED FROM SITE PAGES, reusing `resolveLink`** — same reason as the landing
    page's own links, and the same rule kept in ONE place. Pattern `crs.?settings|^settings\b`, never a
    bare `settings`, which would claim a `Site Settings` page.
  - **A page it cannot find still gets a band**, as a plain `Back` on `history.back()`. The label must
    not promise a destination the code could not locate — but a band that vanished on a failed read
    would leave the admin in the dead end they reported. `missing` ⇒ history; `ambiguous` ⇒ navigate
    anyway.
  - **`withBackToSettings` wraps at the WEB PART boundary (`render()`), never inside a component.**
    That seam is load-bearing: `GroupManager` and `GroupMapBuilder` are ALSO mounted as steps of a
    Folder Management guided flow, and a band offering the way OUT — styled identically to the band
    that goes back one step within it — would read as part of the flow. Wrapping at `render()` means
    an embedded mount cannot inherit it, with no `embedded` prop for a caller to forget.
  - **Folder Administration is the deliberate exception**: `FolderAdmin`'s picker renders
    `BackToSettings` itself, so the deeper views keep their own `Back to Folder Management` rather than
    showing two bands. `BackBand` is shared by both.
  - Not applied to the uploader/approver pages (Upload Form, Approval Document, My Submissions) or to
    `Folder Structure`, which the landing page does not list — a back link there would assert where the
    user came from.
- Property pane carries **one optional address per link** (`link_<key>`, flat, not nested). An override
  **wins outright, unchecked** — the only fix without a redeploy on a site whose names cannot be guessed.
  Placeholders name what auto-detection looks for, so the pane answers "why is this row grey".
- **No access-policy change needed:** a page called `CRS Settings` already matches the `setting` keyword
  in `pageAccessPolicy.ts` → `adminOnly`. Verified, not assumed.

## ⚠ THE APPROVAL GUARD FOUND THE WRONG FOLDER AND BLOCKED EVERY APPROVAL (2026-08-24, 1.0.233.0)
Found on SDG migration day, by two accounts independently. Approving returned *"This unit's folder is
not ready in the Documents library… its permissions could not be read from your account"*, which reads
as a permissions problem and is not one. Spec correction in
`2026-07-29-approval-destination-guard-design.md`; rules now in `shared/approvalDestination.ts` (pure,
22 tests).
- **THE GUARD'S RULE WAS RIGHT; HOW IT LOCATED THE UNIT FOLDER WAS NOT.** It walked **up three
  levels** from the file, assuming `…/Unit/Year/Document Type/file` — and its own comment claimed that
  *"holds for segments with a deeper Levels chain too"*, which is **exactly backwards: a deeper chain
  is what breaks it.** With a below-Unit tier added through the Structure Manager,
  `GHO/GCA/EG/2024/Agreement/Archive 1/file.pdf` puts three-up on **`2024`** — which inherits the
  unit's ACL **by design** — so `HasUniqueRoleAssignments` read `false` and it refused, reporting
  *"it is not locked down"*.
  - **One below-Unit tier is wrong in the OTHER direction and is the dangerous one:** three-up lands
    on the **department**, which does have unique permissions, so it **passed** having checked a
    folder that says nothing about whether the unit folder is locked.
  - **The count is from the TOP now** — segment folder plus one level per PERMISSIONED tier — which
    is exact whatever hangs below it. Counting from the file depends on how many below-Unit tiers a
    segment happens to have *today*, so it was guaranteed to rot the first time anyone edited a chain.
- **⚠ THE TIER COUNT CANNOT COME FROM THE DOCUMENT'S OWN FIELDS, and the near-miss looks correct.**
  Every tier column has a `<Base>Tid` twin, so `discoverTierFields` finds them — but
  `documentDetails.documentUnit` already records that a below-Unit tier such as **SubUnit carries a Tid
  column exactly like a permissioned one**. Counting those lands on the SubUnit folder, which inherits,
  and the guard refuses all over again for a new reason. Only the mode row's `Levels` knows, so it is
  read and passed in.
  - Keyed on the segment **FOLDER** name, which IS the mode row's `StagingFolder` — no term lookup,
    and renaming a segment's label cannot break the match.
  - Read in **its own effect that awaits `primeNames` itself**, not chained behind the document load:
    `cachedListTitle` answers the legacy `DMS Config` until priming settles (the 1.0.207.0 race that
    emptied reconciliation's segment picker). `SortOrder` deliberately not selected (gotcha #11).
  - **Still fails CLOSED, and now names which cause.** A segment whose `Levels` will not parse is
    **left OUT of the map** rather than stored as `0`: absent reads as unknown and refuses, where a
    stored `0` would claim we know the segment is flat. `undefined` is never read as zero.
- **A SECOND BUG IN THE SAME FUNCTION: the destination library was HARDCODED.**
  `DOCUMENTS_URL_SEGMENT` was used unconditionally while `libSeg()` above it is HC-aware — **so
  approving an HC document checked the NORMAL library's folder.** The same HC-clone failure as
  everywhere else in that rollout: one library reference never swapped.
  - `approvedLibSeg()` resolves it and **returns BLANK rather than falling back** when the HC pair is
    unresolved, which `unitFolderPath` refuses by name. Deliberately **unlike `approvedLibTitle()`
    beside it**, whose fallback to `"Documents"` is harmless for a display string and would be the
    worst possible thing here: it would point an HC document's check — and the message an approver
    acts on — at the open library.
  - Both user-facing messages now name the **actual** destination library, so an approver of an HC
    document is not sent to check their access to a library that is working fine.
- **INTERIM WORKAROUND on any site still on an older build: approve from the LIBRARY VIEW**
  (`Approve/reject Items`, or the item's ⋯ → Approve/Reject). That sets moderation status directly,
  Auto-route behaves identically, and this check is not involved — **the guard lives only in the web
  part**, so no flow is affected by any of this.
- **⚠ STILL OPEN, found in the same diagnosis: a HIGHLY CONFIDENTIAL document is sitting in the
  NORMAL `ApprovalDocument` library** on ClarenceDMSTesting (`…/ApprovalDocument/GHO/GCA/EG/2024/
  Agreement/Archive 1/`). Either it predates HC routing or `hcAvailable()` was false at upload.
  Exposure is bounded by draft security (author + approvers only, while pending) but it should not be
  there. **Check whether a NEW HC upload on that site reproduces it** before assuming it is historical.
- **The diagnosis is the reusable part.** Three theories were wrong before anyone looked at the data
  — a missing config row, an HC clearance problem, a response-shape mismatch. One console snippet
  printing the file's real path and each candidate folder's answer settled it in thirty seconds.
  **Ask for the observation before building the explanation.**

## DELETION/SHARE REQUEST AUDIT EVENTS — BUILT AND VERIFIED (2026-08-26)
The deletion-and-share-requests feature (2026-08-15/20/21) writes audit rows through the ordinary
code-side `writeAudit` calls in `Requests.tsx`/`MySubmissions.tsx` — and `CRS Audit Log` restricts
writes to Owners and the service account by design (tamper-resistance). So every request raised or
decided by a PIC or Head of Unit was **silently refused**: the deletion/share itself always worked
(it runs in the approver's own session against the document, unrelated to the audit list's ACL), but
no row was ever written. Confirmed live: an Owner's (Clarence Cho's) own test rows logged correctly;
a guest PIC/HoU pair's identical actions logged nothing.
- **FIXED WITH A NEW FLOW, NOT BY WIDENING THE LIST'S ACL.** `CRS — Audit request activity`, built as
  a **Save As of `Audit — approval activity`** (the create/modify-trigger shape — not
  `Audit — approval deletions`, which is the wrong trigger shape for this and was a false start).
  Widening `CRS Audit Log` write access to uploaders/HoUs was rejected: it would let a non-privileged
  client-side write forge `ActorEmail`, undermining the log's whole tamper-resistant premise. The flow
  runs as the SERVICE ACCOUNT, same as every other audit writer in this system.
- **Trigger:** created or modified, on `CRS Requests`. `EventKind` (Compose) derives the event name
  from `Status`/`RequestType`:
  `if(equals(Status,'Pending'), if(equals(RequestType,'Share'),'ShareRequested','DeletionRequested'),
  if(equals(Status,'Approved'),'RequestApproved','RequestRejected'))`. `ActorEmail` (Compose) picks
  `RequestedBy` on Pending, `DecidedBy` otherwise. Dedupe (`Already logged`, a Get items filtered on
  `ItemUniqueId`+`EventType` within a 5-minute window) guards the same item firing twice, the same
  pattern as the file-lifecycle flows.
- **⚠ THE CLONE CARRIED AN EMPTY `For each` AS DEAD WEIGHT, and it is a silent-write-nothing trap.**
  Power Automate auto-wraps a `For each` around an action the moment ANY of its fields is picked from
  inside an array-typed output (here, `Already logged`'s `value`) — and does **NOT** remove the loop
  when that field is later repointed elsewhere. So after `Segment` was corrected to read from the
  trigger instead of from `Already logged`, the `For each` shell survived with nothing inside it
  referencing the loop item — and a loop over `Already logged`'s array (normally **empty**, since
  empty is exactly what routes into the write branch) runs **zero times**. `Create item` would never
  fire, the run would report **Succeeded**, and no row would ever be written — indistinguishable from
  the original permissions bug except for the green checkmark. Fixed by dragging `Create item` out of
  the `For each` and deleting the empty loop, once every field was confirmed sourced from the trigger,
  `EventKind` or `ActorEmail` and nothing referenced `Already_logged` any more.
  - **THE TELL: check what a `For each` is iterating over even when nothing inside it looks wrong.**
    Every field mapping in `Create item` was individually correct; the defect was structural, one
    level up, invisible unless you click the loop itself and read its "Select an output" value.
- **VERIFIED END TO END, both halves, in two separate runs:** `chocheetuck4` (PIC) raised a Share
  request on an HC document → `ShareRequested` row logged, actor = requester, `Source:
  Flow:RequestActivity`, correct file/library path (resolves the HC library correctly) and correct
  `Details` (requester, reason, recipients). `clarencechojinheng` (HoU) then approved it →
  `RequestApproved` row logged, actor = approver. No duplicates, no stray loop, both event types
  filterable in the viewer's own event-type picker.
- **STILL TO DO, not urgent:** migrate this flow to SDG (export → repoint site URL + the `CRS
  Requests`/`CRS Audit Log` list GUIDs → import as new → test, the same route as the other 13 flows).
  And in a future deploy, **remove the code-side `writeAudit` calls** in `Requests.tsx`/
  `MySubmissions.tsx` — they are now provably a no-op for every non-Owner account and would
  double-write against this flow for an Owner's own actions.

## THE AUDIT FLOWS — THE NORMAL VERTICAL IS COMPLETE AND VERIFIED ON SITE (2026-08-23)
Runbook `2026-08-23-audit-log-flows-runbook.md`. The code-side writers have been live for days; this
is the FILE LIFECYCLE half, which the section below says is missing. **`Uploaded`, `Approved`,
`Rejected` and `Routed` now work end to end on ClarenceDMSTesting** — verified with one document:
three rows, right actors, and `History of this file` resolving across all of them.
- **BUILT: Flow A (`Audit — approval activity`), Flow B (approval deletions), Flow C (Documents
  deletions) and the `Create item` added to Auto-route.**
- **ALL FOUR HC CLONES ARE BUILT (2026-08-24), so HC has parity with the normal libraries.**
  `Audit — HC approval activity`, `Audit — HC approval deletions`, `Audit — HC Documents deletions`,
  and the `Create item` inside `HC Auto Route`.
  - **VERIFIED:** upload → approve → route, one document, three rows, right actor on each, and **no
    spurious `Deleted`** despite Auto-route deleting the source.
  - **VERIFIED TOO: both DELETION clones.** A pending HC file deleted by hand ⇒ one `Deleted` row from
    `Flow:ApprovalDeletionsHC` **with the folder path recovered from the recycle bin**; an approved one
    deleted from `HC Documents` ⇒ `Flow:DocumentsDeletionsHC`. So all five event types now work on
    both verticals, and **`History of this file` threads upload → approve → route on one GUID**.
  - ⚠ **A TRAILING NEWLINE IN `ItemUniqueId` BREAKS THE FILE HISTORY, SILENTLY.** Every filter on that
    column is an EXACT match, so a value ending in `
` never joins the rows holding the same GUID
    without one — the history reads *"No events match these filters"* over a document whose whole
    lifecycle is sitting in the list. It arrives by **pasting an expression into a Power Automate
    field**. **The tell is `xml:space="preserve"` in the REST response**, and it is invisible in a list
    view, in the web part, and in the flow designer. Hit twice on 2026-08-23 (once on a Uri, once
    here); the normal `Auto-route` was checked and is clean.
  - ⚠ **A blank `ItemUniqueId` row poisons every later dedupe.** `Already logged` filters
    `ItemUniqueId eq '…'`, so one row written with a blank id makes that filter MATCH for any later
    event whose id also fails to resolve — the flow then answers *already logged* and **writes nothing
    at all**, reporting success. One such row existed (from a misconfigured `GetSourceUniqueId`) and
    was deleted. Watch for others.
- **⚠ THE DELETE TRIGGER RETURNS A DIFFERENT, MUCH SMALLER SHAPE than the create/modify one** —
  `ID`, `Name`, `FileNameWithExtension`, `DeletedByUserName`, `TimeDeleted`, `IsFolder`, and nothing
  else. **No path and no UniqueId.** The path is recovered from the RECYCLE BIN
  (`_api/web/recyclebin?$filter=LeafName eq '…'` → `DirName`), which also yields the deleter's email —
  a `Deleted` row with no path cannot say WHICH UNIT'S document was destroyed. The `UniqueId` is NOT
  recoverable (the bin's `Id` is the bin entry's own), so **a deletion is the one event that cannot be
  threaded into `History of this file`**.
- **⚠ RUN-AFTER GUARDS THE ACTION; IT DOES NOT GUARD THE EXPRESSION THAT READS ITS OUTPUT.** The
  recycle-bin call went GREEN while `first(body('GetDeletedInfo')?['value'])` threw *"One or more
  fields provided is of type 'Null'"* and failed the whole flow — because the `Accept` header was
  missing, SharePoint answered verbose (`{"d":{"results":[…]}}`, no `value` key). Two lessons in one
  failure: **the `Accept: application/json;odata=nometadata` header is load-bearing on every one of
  these calls**, and every expression reading a response must be wrapped
  (`first(coalesce(…, createArray()))?['DirName']`) because a wrong response SHAPE is not an action
  failure. It is `FileNameWithExtension`, NOT the
  braced `{FilenameWithExtension}` — the braced form resolves to null and writes a row with a blank
  filename. There IS an actor (`DeletedByUserName`), better than the design assumed.
- **⚠ BOTH DELETION FLOWS NEED `@equals(triggerOutputs()?['body/IsFolder'], false)`.** Without it,
  deleting a FOLDER logs as a destroyed document — and reconciliation, subtree migrations and stray
  cleanup all delete folders in bulk, so one structure change buries the real deletions.
- **`EventTime` on a deletion is `TimeDeleted`, never `utcNow()`** — the trigger polls, so `utcNow()`
  records when the flow noticed rather than when the deletion happened.
- **⚠ THE CONDITION FORM DIFFERS BETWEEN A AND B, AND THAT IS NOT TIDINESS.** Flow A's
  `if(empty(...), 'write', 'skip') = 'write'` works there and **FAILED in Flow B** with the array
  verifiably `[]` — False where it should have been True, cause never established (Flow A's code view
  carries a trailing `

`, so invisible whitespace is the suspect). B and C use
  `length(coalesce(..., createArray())) = 0`, which cannot carry whitespace and keeps the fail-open
  rule. **Flow A was left alone deliberately** — verified behaviour is not changed on a theory. Use
  the numeric form for both when rebuilding.
- **Flow C has NO routing check**, because nothing routes files out of `Documents` — **until the
  seven-year archive mover lands**, at which point `MoveTo` fires this trigger and every archived
  document logs as `Deleted`. The mover must write its row before the move and Flow C must gain B's
  check-first branch.
- **⚠ `{ModerationStatus}` FROM THE CONNECTOR IS A STRING, AND REJECTION READS `"Denied"`.** Both
  halves of that were wrong in the first build and both fail SILENTLY. Comparing to the integers
  `0`/`1` — which is what every REST call in this codebase reads — matches nothing, so **every
  approval and rejection logs as `Uploaded` with the approver's name on it**. Fixing that but
  comparing to `'Rejected'` still falls through, and then the dedupe finds the existing `Uploaded`
  row and skips, so **nothing is written at all** and the run reports success. Caught only by reading
  the raw trigger output.
- **⚠ `{Identifier}` IS A DOUBLE-ENCODED PATH, NOT THE FILE GUID.** Every code-side writer stores a
  real `UniqueId`, so keying flow rows on `{Identifier}` would have broken the "history of this file"
  join between flow rows and admin rows — the one feature that makes this list worth reading. The
  flows fetch it with `_api/…/items(ID)/File?$select=UniqueId` instead.
- **TWO PROBLEMS THE DESIGN DID NOT ANTICIPATE, both fixed by ONE mechanism.** Auto-route DELETES the
  approved source, so flow B would log **every approved document as `Deleted`**; and the upload form
  uploads-then-tags, so a single upload fires the create/modify trigger twice. Both are handled by
  writing the `Routed` row **BEFORE** the delete and having every writer check the list first.
  **The ordering is the whole safety** — put that `Create item` after the delete and the deletion flow
  wins the race and logs the spurious `Deleted` anyway.
- **A FAILED DEDUPE CHECK WRITES THE ROW ANYWAY** (`empty()`, not a length test — `empty(null)` is
  true where `length(null)` throws). An unreadable audit list says nothing about whether the event was
  already logged: a duplicate row is something a human reads past, a missing row is gone. The one
  place where this codebase's fail-open habit and "log more" point the same way.
- **⚠ AN AUDIT WRITE MUST NEVER BE ABLE TO STOP A DOCUMENT ROUTING.** `Get source author` — the action
  AFTER the new `Create item` in Auto-route — carries `Run after: has failed`. Note the shape:
  **`Run after` is configured on the DOWNSTREAM action**, and getting that backwards protects nothing
  while looking correct.
- Flow B cannot use the GUID at all: the file is deleted, so nothing remains to fetch it from, and
  `ItemPath` is a **Note** column that SharePoint cannot `$filter`. It matches the `Routed` row on
  **`ItemName` + a 5-minute `EventTime` window** instead.
- **Built as `clarence@trinergydigital.com` here. On SDG's tenant every flow must be created signed in
  as the SERVICE ACCOUNT, from the first action** — the connection is baked in for life.
- The `CRS Audit Log` list's permissions are still inherited. Design §5.2 leaves that manual on purpose
  (breaking inheritance has to name the service account, and a wrong guess locks the flows out of the
  list they write to).

## Audit Log (2026-08-13, spec `2026-08-13-audit-log-design.md`)
Client asked to "track every single thing". **BUILT: the list, the writer, the viewer and four
writers; NOT yet site-tested, and the flows are not built.** Web part **`CRS Audit Log`**
(`9e2c4d17-…`), its own page, admin-only.
- **PURVIEW IS UNREACHABLE — do not re-propose it.** Three independent blockers: the only web-part
  route is Graph's audit-log query API needing **tenant-wide** `AuditLogsQuery.Read.All`; API
  permission requests are processed from the **TENANT** app catalog while this package ships to the
  **site collection** one, so it cannot even ask; and the API is an async job, not a query. It also
  would not answer the event the client named — an approval appears there as a generic item
  modification, with no from/to status. **Reads and downloads are therefore unrecordable by us**, and
  that is the one gap that matters; it is stated on the page, not buried in a doc.
- **`<P> Audit Log`** (`LIST_SUFFIX.auditLog`), self-provisioned by the page — the client cannot run
  PowerShell, so a scripted step would not happen. 14 columns, internal names space-free;
  `EventTime`, `EventType`, `ActorEmail`, `ItemUniqueId` indexed. **Provisioning does NOT set the
  permissions** (breaking inheritance must name the service account, and a wrong guess locks the flows
  out of the list they write to) — the screen states it as a manual step.
- **WRITES ARE JSON LIGHT WITH NO `__metadata`, and both header halves must say
  `odata=nometadata`.** Two failures on 2026-08-13 before this landed: `Accept: verbose` with
  `Content-Type: application/json` → *"The property '__metadata' does not exist on type 'SP.List'"*
  (the body was parsed as non-verbose); then BOTH headers verbose → *"Parsing JSON Light feeds or
  entries in requests without entity set is not supported"*, because SPFx's `SPHttpClient` attaches
  its own OData version header that the OData 3 verbose dialect does not agree with. JSON light needs
  no envelope and **no entity type**, so `ListItemEntityTypeFullName` is not read for this list at all
  — do not "restore" it per gotcha #12, which applies to the verbose call sites in `FolderMap.tsx`.
  The list create failed loudly; the row write carried the same bug and would have failed SILENTLY on
  every event.
- **`EventType` is TEXT, never Choice.** Writing a value absent from a Choice column's `Choices`
  FAILS the whole write, so the day someone adds a type in code, every row of that type is lost
  silently.
- **`EventTime` is written ISO, NOT `M/D/YYYY` — gotcha #1 does NOT apply to this endpoint.** That
  locale format belongs to `validateUpdateListItem`, which parses in the site's locale; a plain
  `/items` POST goes through the OData layer and answers a locale string with *"Cannot convert a
  primitive value to the expected type 'Edm.DateTime'"* — a 400 that names the type but not the field.
  `$filter` needs ISO too, so writes and filters share ONE format and cannot be mismatched. Display is
  `DD/MMM/YYYY HH:mm`, in the viewer only. **Also send `odata-version: ""`**: SPFx's `SPHttpClient`
  injects `4.0`, under which SharePoint cannot infer the entity set for a JSON-light entry payload, so
  a row POST 400s while `/_api/web/lists` tolerates the identical headers.
- **`writeAudit` never throws and never blocks** the action it logs, but is never silent either: it
  returns false, admin screens raise a non-blocking warning, and the console carries the status. An
  audit gap someone knows about is worth far more than one nobody does.
- **ONE ROW PER RUN** for reconciliation and migration, with the log in `Details` — one row per folder
  would bury every other event the first time somebody reconciles. Counts must come from LOCALS, not
  from React state set during the run (a closure reads its render-time value and records zero).
- **"Could not read" and "nothing matched" are separate states everywhere.** Three distinct empty
  states in the viewer, for the same reason `unknown` ≠ empty elsewhere in this codebase.
- **Nobody but Owners and the service account can write**, and that costs nothing: flows write file
  events as the service account and admins write admin events, so no uploader or approver needs
  access. Tamper-resistant by construction rather than by policy. **No auto-delete, ever**; version
  history on.
- **Every code-side writer is wired** (10 event types, 14 call sites): `PolicyChanged`,
  `AbbreviationChanged`, `ReconciliationRun`, `AccessGranted`/`AccessRevoked` (library entry ×3, site
  entry ×4), `UploadRefused`, `SegmentCreated`, `StructureChanged`, `MigrationRun`, `GroupMapChanged`
  (add + delete). **Only the flows remain** — 3 new ones plus the added Auto-route action (spec §8),
  which are what supply `Uploaded`/`Approved`/`Rejected`/`Routed`/`Deleted`. Until they exist the log
  records ADMIN activity only and the file lifecycle is absent, so do not read an empty file history
  as "nothing happened to that file".
- **A HALF-DONE permission change is recorded as half-done, never flattened.** Three places where the
  distinction IS the value of the row: a Group Map row can exist while the live grant failed
  (reconciliation fixes it); a revoke can remove the mapping while the access REMAINS (nothing on
  screen then shows who holds it — `Outcome: "Failed"`, naming them); and deleting one of several
  Group Map rows keeps the group, so that folder grant survives. A row reading only "access revoked"
  would stop someone looking.
- **A staged structure change SAYS staged.** `StructureChanged` distinguishes `PendingLevels` from
  `Levels`, because for a segment in use uploads deliberately have not moved yet — precisely the
  misreading the staging exists to prevent.
- **Site-access removal carries its caveat into the record**, not just the toast: folder ACLs are
  untouched, so it is a lock-out and not a de-provisioning.
- `SubtreeMigrator` mirrors its run log into a **`useRef`**, not state, for the same stale-closure
  reason as the reconciliation counts.
- **Supersedes the never-built `<P> Deletion Log`** — two append-only trails would leave a permanent
  question about which is authoritative. `LIST_SUFFIX.deletionLog` still exists; nothing new reads it.

## C-LEVEL IS READ-ONLY AGAIN; HoD TAKES DELETE AND SHARE; PIC LOSES STAGING DELETE (2026-08-20, 1.0.197.0)
Client: *"C level's Share and Delete power is now suppose to be pass on to HOD, no more C level, C level
is read only"* and *"For Staging PIC should not be able to delete, they have to request from HOU"*.
Reverses the 2026-08-15 C-Level widening, the 2026-08-19 segment fan-down, the 2026-08-17 HoD view-only
rule and the 2026-08-15 PIC correction — four decisions, in one instruction.

| Persona | Was | Now |
|---|---|---|
| `clevel_global` | `GLOBAL + DEL + SHARE` | **`GLOBAL`** |
| `clevel_segment` | `SEGVIEW + DEL + SHARE` | **`SEGVIEW`** |
| `hod` | `DEPTVIEW` | **`DEPTVIEW + DEL + SHARE`** |
| `pic` | `UPL + DELS` | **`UPL`** |
| `pic_hc` | `UPLHC + DELSHC` | **`UPLHC`** |
| `hou` | unchanged | `APR, DELS, DEL, SHARE, UPLHC, DELSHC` (superseded 2026-08-24: `UPL` for `UPLHC`; HC filing moved to `hou_hc`) |

- **THE SEGMENT FAN-DOWN WITHDREW ITSELF, with no edit to reconciliation.** `segmentFanRoles()` is
  DERIVED from `PERSONAS`, so removing DEL/SHARE from the C-Level personas narrows it to
  `GLOBAL + SEGVIEW` automatically. **This is the whole return on deriving it** — a literal list would
  have gone on fanning delete from every segment-tier row while `groupMapModel.ts` said otherwise, and
  the drifting copy is the one reconciliation reads.
- **⚠ HoD's NEW POWERS REACH NOTHING UNLESS `recon_departmentFanOut` IS ON.** Every unit folder has
  unique permissions, so a grant on the DEPARTMENT folder alone reaches no unit — the exact defect
  found in `clevel_segment` on 2026-08-19, now one tier down. It defaults on and must stay on; with it
  off, reconciliation REPORTS both roles and grants neither.
  - **Deliberately NOT exempted from that switch** the way SEGVIEW and the old C-Level roles were. That
    exemption rests on *"the role name is the consent"* — true for SEGVIEW, which granted nothing
    anywhere before 2026-08-07, and **false here**: department-tier `DEL` rows already exist on
    provisioned sites from the pre-2026-08-17 HoD persona, so a leftover row and a deliberate one are
    once again indistinguishable by tier. The switch is what reads them.
- **⚠ A PIC NOW HAS NO ROUTE TO REMOVE A PENDING OR REJECTED FILE AT ALL.** Requests are raised from My
  Submissions on an **approved** file only — precisely because a PIC used to hold delete in the approval
  library and needed no permission there. Removing `DELS` closes that door without opening the other
  one. Until requests cover pending files, the answer is a Head of Unit deleting it for them, out of
  band. **Nothing on screen says this**; it must be told to the client.
- **`DELS` and `DELSHC` now have exactly ONE holder each (`hou`), pinned by test.** `DELSHC` went from
  `pic_hc` for the same reason: the client's rule is about the ROLE, and HC Approval Document is a
  staging-side library — leaving it would mean clearance decided a DELETE right, which it decides
  nowhere else.
- **⚠ MIGRATION IS THE WHOLE JOB, AND NOTHING FAILS IF IT IS SKIPPED.** Existing Group Map rows still
  say `DEL`/`SHARE`/`DELS` and keep granting them. The persona edit changes what NEW mappings write and
  nothing else.
  - **Deleting a Group Map row does NOT revoke the folder grant.** `groupsToRemove` asserts the full ACL
    at **page scope only** — at folder scope a library root carries SharePoint's automatic Limited
    Access entries, so asserting there would strip every group's access. So a removed row leaves
    `CRS Delete` on the folder for ever.
  - **The reliable path is therefore DELETE THE GROUP and re-create it** (Group Management → Bulk
    provisioning → reconcile). SharePoint drops a deleted principal's role assignments, which is the
    only thing that actually withdraws the permission. Verified repair path, 1.0.192.0.
- **⚠ A PERSONA SHEDDING A ROLE BREAKS `planBulkGroups`'s RENAME RECOVERY.** `roleSetKey` is an EXACT
  fingerprint, so a PIC group provisioned as `UPL+DELS` no longer matches the `UPL` persona and a
  RENAMED unit would be planned as new — duplicate groups, the 1.0.162.0 bug by a new route. The name
  check runs first, so an un-renamed group is safe. Pinned by a regression test.
  - **Subset matching was REJECTED**: `employee` is `["MEMBER"]` and `employee_hc` is
    `["MEMBER","MEMBERHC"]`, so a subset rule would let a cleared viewer group answer for the plain one
    — reusing the WRONG group, which is worse than duplicating.
  - Fifth instance of memory `feedback-audit-consumers-when-narrowing-data`: **narrowing a shared
    definition makes every reader of it a change site.**

## THE REPLACE-CLASH POPUPS ARE GONE — MOVED TO STATIC TEXT (2026-09-02)
Client: *"there is two popup for sending for a file replacement, one is custom which is our one and
another is native, the native one is annoying to client... there is no need for two popup as the one
we made is enough."* Two separate native `window.confirm()` calls removed, both replaced with static
text shown BEFORE the deciding click rather than a second popup AFTER it. **Built, not yet
site-tested.**
- **`Form.tsx` — the redundant confirm inside the custom clash dialog is gone.** Clicking "Replace the
  waiting document" or "Send for approval as a replacement" used to fire a native `confirm()` on top
  of the custom dialog already clicked through — two popups for one decision. The facts that confirm
  carried (version history vs the recycle bin, the "someone else's invisible pending file" warning)
  are now static amber text shown above the buttons, always visible before either is clicked. The
  custom dialog's own button click is now the only confirmation.
- **`ApprovalDocument.tsx` — the native "Approving REPLACES it… Approve anyway?" confirm is gone.**
  `documentsFileClash` now runs PROACTIVELY the moment the document loads (`clashCheck` state, a new
  `useEffect` placed directly after the `approveRight` probe — same "must stay above the render
  guards" rule, same reason: a hook added below `loading`/`fetchError`/`!item`'s early returns blanks
  the whole web part, per the 1.0.194.0 lesson this file already carries). The result renders as
  static amber text under the comment box — same two messages the confirm used to show (confirmed
  clash vs. an unanswerable check). **Approve is now a single click.**
  - **`submitDecision` still RE-CHECKS fresh at submit time** — unrelated to display, and unchanged:
    a name landing AFTER page load and BEFORE the click must still be caught. It just no longer
    interrupts; on a confirmed clash it proceeds straight to the existing `markRecordReplaced`
    bookkeeping with no `window.confirm()` gate in between.
- **Neither guard's SAFETY changed, only its UI.** Nothing is skipped, refused less, or approved more
  easily than before — the same facts are just told earlier and without an extra click.
- `tsc --noEmit` clean, full suite **1597/0**, no new lint warnings.

## THE NATIVE APPROVE COMMAND BYPASSES BOTH GUARDS — CLOSED AT THE COPY, 2026-08-25
Confirmed on ClarenceDMSTesting the same day the guards above were built. `documentsFileClash` and
the destination-folder guard both live entirely in `ApprovalDocument.tsx` — client-side logic that
runs only when someone presses Approve **through that web part**. SharePoint's own classic
**Approve/Reject** command — reachable from the `Approve/reject Items` view, or an item's own `⋯`
menu, which exists automatically the moment content approval is on — sets
`OData__ModerationStatus` directly, with no awareness either guard exists. Auto-route then routes it
on its next poll exactly as if nothing had been checked. Until 2026-08-25 that meant a silent
OVERWRITE, because `Copy file` was set to replace; it copies with a new name now (below).
- **Reachable by ANYONE who holds Approve on the folder**, not an obscure admin path — the same
  limitation this project already accepted for the native Share button (spec
  `2026-07-23-share-guard-retirement.md`), for the same reason: a web part cannot intercept a
  SharePoint-native control, only offer its own route beside it.
- **Today's mitigation is PROCESS, not code**: tell every approver to use the Approval Document
  **page**, never the classic library's Approve/Reject command or the `Approve/reject Items` view.
  There is no way to hide or disable that native control while content approval is on.
- **FIXED 2026-08-25 by the CLIENT'S suggestion, and it is better than either option we had.** Both
  `Auto-route` and `HC Auto Route` now set `Copy file` -> "If another file is already there" =
  **Copy with a new name** (`nameConflictBehavior: 2`). Nothing is ever overwritten, from ANY route:
  the native command, our page, the bulk panel, or the bulk-import auto-approve flow.
  - **Two options were on the table and both were worse.** *Fail* protects the existing document but
    leaves the new one approved-and-unrouted with nobody told — a stuck document nobody is looking
    for. A pre-`Copy file` existence CHECK is a whole extra branch and, being a check, still answers
    about a slightly earlier moment.
  - **⚠ THE INSIGHT WORTH KEEPING: the copy is the only place whose answer cannot go stale.** Every
    check the app makes — `approvedClash` at upload, `checkDestinationClash` at approval — is correct
    about the moment it ran and can be overtaken before the copy happens. Adding a third and a fourth
    would not change that. The conflict behaviour acts AT the copy, so no gap is left.
  - **The app-side checks stay, and are still worth having:** they catch a clash EARLY, with a
    message naming the file, instead of leaving somebody to notice `report-1.pdf` weeks later. They
    are simply no longer the last line of defence.
  - **Cost, accepted:** a renamed copy breaks the `[Project] - [Vendor] - [Name] - [Date]` convention
    and nobody is told a collision happened. Two documents to reconcile beats one silently destroyed.
  - **⚠ SET IT ON EVERY NEW ROUTING FLOW.** The connector defaults to *Replace*, so a hand-built flow
    on SDG gets the dangerous behaviour unless someone changes it deliberately. Verify in Code view:
    `"nameConflictBehavior": 2`. `0` is Fail — the one setting that strands documents.
- **THE TWO WAYS A CLASH REACHES THE COPY, now both harmless:**
  1. **Nobody checked** — a file placed into `Documents` outside the upload form (an admin dragging
     one in, OneDrive sync, `Copy to`/`Move to` from elsewhere). No check ever ran.
  2. **The check went stale** — bulk import's second batch checking while the first is still routing
     (~18s per file), something landing in `Documents` between an upload and its approval, or a
     subtree migration re-filing documents into a path a pending file will later use.

## A BROWSER EXTENSION BLOCKING AN UPLOAD NOW SAYS SO (2026-08-24, 1.0.235.0)
Found live, testing today's other changes: an upload's `Files/Add` POST was blocked by **Brave
Shields** — confirmed in DevTools by "Provisional headers are shown" with no response headers and
no status, the exact signature of a request stopped before it reached the network, while every
other request on the same page (folder lookups, config reads, both clash checks) succeeded through
the identical connection. The uploader saw the raw browser string **`"Failed to fetch"`** as the
failure reason — true, and useless: nothing told them it was their browser, not their document,
their permissions, or the DMS.
- **`friendlyUploadError` in `shared/networkErrors.ts`** (pure, 7 tests) substitutes an actionable
  message — names Brave Shields, ad blockers and privacy extensions, and says what to click — for
  the handful of literal strings browsers themselves use when a request never reached the server:
  Chrome/Edge/Brave's `"Failed to fetch"`, Firefox's `"NetworkError when attempting to fetch
  resource"`, Safari's `"Load failed"`.
- **DELIBERATELY NARROW MATCHING.** Only those exact browser-level phrases trigger the substitution.
  A genuine SharePoint error — 403, a malformed payload, a throttle — never produces them, so this
  can never mask a real, actionable server response behind a generic "check your extensions"
  message. `console.error` still logs the original at both call sites for actual debugging.
- **Two call sites, one function**: `Form.tsx`'s `uploadStagedFile` catch (the uploader-facing
  message) and `BulkUpload.tsx`'s per-file catch (the admin-facing `detail` field) — the only two
  places in either upload web part that surfaced a raw caught error message to a screen.

## APPROVING CAN NO LONGER SILENTLY OVERWRITE AN EXISTING DOCUMENT (2026-08-24, 1.0.235.0)
Client, checking whether an earlier flagged gap had actually been fixed: it had not — only noted.
`documentsFileClash` in `ApprovalDocument.tsx`, beside the destination-folder guard it complements.
- **THE RACE: upload-time and approval-time are two different moments, and only one was checked.**
  `Form.tsx`'s `approvedClash` (2026-08-22) refuses an upload whose name already exists in
  `Documents` — but that runs once, before this document is even in the approval library. It cannot
  see a name that lands in `Documents` **between** this upload and this approval (a second uploader,
  a bulk import). **Auto-route's `Copy file` step replaced on a name clash until 2026-08-25, when
  it was changed to "Copy with a new name"** (verified in the flow's
  own Code view), so that race ends in a silent overwrite — no error, no warning, a green run.
- **Re-checked immediately before submitting the approval**, against the file's FULL destination
  path — library segment swapped, everything below it (Year/Document Type/Archive…) preserved
  exactly, mirroring what Auto-route's own `Compose_1` expression preserves. NOT the unit folder the
  neighbouring guard computes, which deliberately stops at Unit.
- **⚠ FAILS CLOSED — the OPPOSITE direction from the upload-time check it complements.** That check
  runs on every upload site-wide, so an unanswerable read there would take the whole form out of
  service, an acceptable trade to avoid. This runs once, at one approval, and its neighbour (the
  destination-folder guard, same file) already established why fail-closed is right here: a retry
  costs the approver seconds, a wrong "proceed" overwrites a document and nobody finds out.
- **A 404 on the below-Unit folder means CLEAN, not inconclusive — a different question from the
  neighbouring guard's 404.** That guard's 404 (on the UNIT folder) means "not ready, refuse." This
  one's 404 is on a below-Unit folder, ensure-created ON DEMAND by Auto-route's own `Create new
  folder` step — normal to be absent before the FIRST approval into a given Year/Document
  Type/Archive combination. Absent folder and absent file mean the same thing: nothing to clash with.

## A HEAD OF UNIT'S OWN UPLOAD CAN AUTO-APPROVE — TOGGLE, OFF BY DEFAULT (2026-08-24, 1.0.234.0/235.0)
Client: *"can we make it that if normal HOU or HC HOU upload, they dont need to do approval but
rather it will automatically approve."* Rules in `shared/selfApprove.ts` (pure, 2 tests); the live
ACL read is `probeFolderApproveAccess` in `shared/dmsFolderMap.ts`. **BUILT, off until a `CRS Config`
row turns it on.**
- **⚠ RE-CHECKS FOR A NAME CLASH IMMEDIATELY BEFORE THE MERGE, ADDED 1.0.235.0 — found by the
  client asking directly whether self-approve had the same overwrite protection just built for
  manual approval.** It did not: the upload-time clash check (`approvedClash`, 2026-08-22) runs once
  at the TOP of `uploadStagedFile`, seconds before the self-approve MERGE at the bottom of the same
  function — a small but real window for a second uploader or a bulk import to file the same name
  into the approved side in between. `checkApprovedClash` (the extracted, now-shared version of that
  check) is called a second time, right before the MERGE; a clash SKIPS self-approve only — never
  fails the upload, since the file is already correctly uploaded and tagged by that point. Still
  fails OPEN, unlike `documentsFileClash`'s approval-time guard: a wrong "no clash" here just leaves
  the item Pending, exactly where a denied probe already leaves it — never an overwrite, because
  self-approve simply does not fire.
  - **The function had to be HOISTED**, not just called twice in place — it was declared inside the
    FIRST of two `try` blocks in `uploadStagedFile` and went out of scope before the second (the
    self-approve one) could see it. Caught by `tsc`, not silently; moved above both blocks.
- **A `CRS Config` row, `Title = "autoApproveOwnUpload"`, `SettingValue = "yes"` turns it on.**
  Anything else — absent, blank, `"no"`, a typo — means OFF, i.e. the existing manual-approve
  behaviour. **Fails CLOSED, the reverse of `uploadsPaused`/`legallyPrivilegedFor`**: those fail open
  because the cost is a form out of service for a minute; here a wrong "on" publishes a document
  nobody reviewed, which is the more expensive failure.
- **THE GATE IS A LIVE ACL READ, NEVER A ROLE LOOKUP.** After tagging succeeds, `Form.tsx` probes
  `EffectiveBasePermissions` on the EXACT folder the file just landed in for the `ApproveItems` bit
  (bit index 4 — `viewListItems=1, addListItems=2, editListItems=3, deleteListItems=4,
  approveItems=5`, so `5-1`). Only `"granted"` self-approves; `"denied"/"missing"/"unknown"` all
  leave the item Pending. A role table can say "this persona approves" while reconciliation has not
  actually granted it yet, or a group was renamed — the ACL read cannot be wrong that way.
  Arithmetic, never `&`, same reason as the upload probe next to it: Full Control returns
  `Low = "4294967295"`, which JS bitwise coerces to a signed 32-bit int.
- **NO PERSONA BRANCH NEEDED — covers `hou` and `hou_hc` for free.** `folderId` and
  `libraryTitleForTagging` already point at whichever library (normal or HC) the caller resolved
  this file into, so the probe answers correctly for either without the code knowing which persona
  is uploading. A PIC's probe always reads `"denied"` — they hold no Approve role on any folder — so
  this can never fire for a PIC, with no explicit check needed to keep it that way.
- **HoU KEEPS THE APPROVAL EMAIL** (client's explicit instruction — do not suppress it). Nothing
  here touches Auto-route: flipping `OData__ModerationStatus` to `0` via a MERGE is indistinguishable
  from a human clicking Approve, so Auto-route's own poll picks it up and sends the same *"Your file
  has been approved"* email to the same person for the same reason, with no code change needed.
- **NEVER BLOCKS OR FAILS THE UPLOAD.** This runs last, after tagging has already succeeded, inside
  its own try/catch — a failed probe, a failed MERGE, or a throttled request all just leave the item
  Pending, which is the upload's existing outcome. The uploader never sees an error from this.
- **Bulk Upload is untouched, deliberately** — it already writes straight to the approved side with
  no approval step at all, so there is nothing here for it to skip.

## A NORMAL HEAD OF UNIT CANNOT REACH THE HC LIBRARIES AT ALL — `hou_hc` (2026-08-24, 1.0.234.0)
Client, twice in one day, each narrower than the last. First: *"we will need to make another highly
confidential group for HOU, so not every HOU can upload into Highly Confidential, it will be the same
pattern as PIC"* — which would have kept APR/DELSHC on the plain persona (clearance gates filing
only). Asked to confirm, the client settled it: *"a normal HOU cannot see HC Approval Document and HC
Documents … only HC HOU can see and approve and go into HC libraries."* **That second sentence is
what shipped.** Spec `2026-08-24-hc-head-of-unit-clearance-design.md` (superseded by this section —
read here, not there). **BUILT, NOT provisioned on either site.**
- **`APR`, `DEL` AND `SHARE` LEFT BOTH HC ROWS OF `LIBRARY_ROLES`.** A plain Head of Unit holds all
  three for their ORDINARY work, and a role held by two personas cannot grant to one and withhold
  from the other — so seeing HC could not be separated from filing it while APR stayed on
  `StagingHC`/`DocumentsHC`. Three HC-specific twins replace them, held only by `hou_hc` (and, for
  DEL/SHARE, by `hod`):
  - **`APRHC`** — approve, both approval libraries. A superset like `UPLHC`, exactly as `hou_hc`
    approves ordinary documents through it too — no separate `APR` row.
  - **`DELHC`** — delete approved HC documents (`DocumentsHC`).
  - **`SHAREHC`** — share approved HC documents (`DocumentsHC`).
  - `DELSHC` (pending-HC delete) stays exactly where it was, now `hou_hc`-exclusive.
- **`hou` = `APR, DELS, DEL, SHARE, UPL`.** No HC role of any kind. **`hou_hc` = `APRHC, DELS, DEL,
  SHARE, UPLHC, DELSHC, DELHC, SHAREHC`** — a superset of the plain persona's ORDINARY powers plus
  every HC one. A Head of Unit gets ONE of the two groups, never both.
- **⚠ THE STAFFING RULE THIS CREATES, because nothing on screen will say it:** a unit that files HC
  documents MUST have someone in its `_APR_HIGHLY_CONFIDENTIAL` group, or its HC uploads sit pending
  forever — invisible to everyone but their author, with no error anywhere. That is the direct
  meaning of the client's instruction, chosen with the deadlock stated plainly and accepted.
- **`hod` GAINS `DELHC`/`SHAREHC` THE SAME DAY, to keep a power it already had.** The Head of
  Department's department-wide HC delete and share (2026-08-20) rode on plain `DEL`/`SHARE`, which
  just left `DocumentsHC`. Confirmed with the client: management oversight of HC stays — a Head of
  Department and C-Level still read HC Documents in their scope, and HoD keeps delete/share over it.
  **HoD's migration is ADDITIVE, unlike the approver groups** — same group name, two new rows mapped
  onto it by a re-run of bulk provisioning; no deletion needed.
- **APRHC IS REVIVED AS A FULL GRANTING ROLE**, not the naming-only shim the first design used for a
  few hours. It sits in `LIBRARY_ROLES`, `ROLE_TO_PERMISSION` (→ `DMS Approve`/`CRS Approve`,
  downgraded to Read on the approved side like `APR`), `STAGING_FACING_ROLES`, and both request-
  routing matches (`Requests.tsx`, `MySubmissions.tsx` now match `APR` **or** `APRHC` — matching APR
  alone would leave every HC unit's queue permanently empty). The `ROLE_ALIASES.APRHC → APR` rewrite
  is GONE: a stored `APRHC` row is a real, distinct grant now, not a stand-in for "approver".
- **⚠ EVERY `_APPROVER` GROUP ALREADY PROVISIONED NOW FINGERPRINTS AS `hou_hc`'s PRE-SPLIT SHAPE** —
  its role set (`APR, DELS, DEL, SHARE, UPL, UPLHC` from the 2026-08-17 widening) matches neither
  persona exactly, so `personaForRoles` reports it unmatched until migrated.
  **⚠ RUNNING BULK PROVISIONING WITHOUT DELETING THE OLD `_APPROVER` GROUPS FIRST MAKES A MESS:** the
  name check finds `GHO_GF_TAX_APPROVER` exists and writes rows onto the OLD group, which keeps its
  HC grant, and no `_APR_HIGHLY_CONFIDENTIAL` group is created. **Migration order is not optional, per
  segment:** export the Group Management CSV (members — deleting a group loses them) → delete the
  segment's `_APPROVER` groups → bulk provision → reconcile → re-add people (uncleared HoU →
  `_APPROVER`; cleared HoU → `_APR_HIGHLY_CONFIDENTIAL` only). `hod`'s groups are untouched by this —
  see the additive note above.
- **Until a unit is migrated, its HoU KEEPS full HC access** — the persona edit changes what new
  mappings write and nothing else; reconciliation has already granted the old rows on the HC
  folders, and deleting the group is the only thing that withdraws them (2026-08-20).
- **Page policies and request routing both updated in the same change** — `Upload-Form.aspx`,
  `ApprovalDocument.aspx`, `My-Submissions.aspx`, `Bulk-Upload.aspx` and `Requests.aspx` all now list
  `APRHC` beside `APR`; omitting it would `AccessDenied` the exact persona the HC vertical exists
  for. `collectMembership` needed no change — it counts UPLOAD roles (`UPL`/`UPLHC`), which this
  did not touch; `hou_hc` still uploads through `UPLHC` exactly as before.

## HC CLEARANCE IS ONLY FOR PIC AND SDG EMPLOYEE (2026-08-17, spec `2026-08-17-hc-clearance-and-role-revision-design.md`)
The client **reversed their 2026-08-15 rule** (*"any approver which is HOU can see Highly Confidential
files as well… we do not need a dedicated HOU but rather a dedicated PIC who is an uploader only and
dedicated viewer who is SDG Employee"*), and separately made **Head of Department view-only** (*"HOD no
need deletion power, he only view"* / *"do not have share functionality that is HOU, just needs View"*).
- **HC clearance is a dedicated group for the two roles at the BOTTOM only** — the uploader and the
  plain viewer. Every management role reaches HC through the role it already holds. Say this to the
  client, because it reads oddly: **a Head of Unit needs no clearance, an SDG Employee does.**
- `APRHC` and the `hou_hc` persona are **RETIRED**. `_APR_HC` group names and stored `APRHC` values
  resolve to plain `APR` (suffix kept + `ROLE_ALIASES`) — dropping them would make such a name fall
  through to **MEMBER**, silently reclassifying an approver group as view-only.
- **THREE COLLISIONS, all the same shape: a role held by two personas cannot grant to one and withhold
  from the other.** Each obvious implementation hands HC to people the client excluded.
  1. **HoU uploads HC via `UPLHC` ON THE PERSONA, never `UPL` on the HC library** — the plain PIC holds
     `UPL`, so that would give *every* PIC HC upload. `UPLHC` **replaces** `UPL` on `hou` (it is a
     superset covering the normal library too, so both would be two rows granting the same thing).
  2. **HoD gets `DEPTVIEW`, never `MEMBER`** — `MEMBER` is the SDG Employee role and is absent from
     `DocumentsHC`; reusing it would either lose HoD's HC read or, once added there, hand every SDG
     Employee HC read. `DEPTVIEW` is what `SEGVIEW` is to a segment, one tier down.
  3. **`DELSHC`, because `DELS` cannot sit on the HC approval library** — the plain PIC holds `DELS`, so
     **until today every plain PIC held `CRS Delete` on HC Approval.** A live leak, narrowed by Draft
     Item Security but real for an approved HC file awaiting Auto-route.
- **⚠ `DEL` IS NOT ON `StagingHC`, AND AN EARLIER DRAFT GOT THIS WRONG.** The reasoning "HoD lost `DEL`,
  so `DEL` is now HoU-exclusive" is false — **C-Level carries `DEL` too** (`clevel_global` =
  `GLOBAL+DEL+SHARE`). `DEL` there would have given a C-Level read and delete on **unapproved HC
  drafts**. HoU's HC pending delete comes from `DELSHC`. *"Role X is exclusive to persona Y"* is a claim
  about the whole `PERSONAS` array; both holders sets are pinned by test.
- **`collectMembership` NOW COUNTS `UPLHC` AS AN UPLOADER ROLE — its absence meant HC upload had NEVER
  WORKED.** `pic_hc` carries `UPLHC` and no `UPL`, so every HC-cleared PIC had zero uploadable paths and
  was told *"your account isn't fully provisioned to upload"* — the folder ACL correct, the form
  disagreeing, exactly as with the long-form `Role` bug of 2026-08-07.
- **`suffixForRole` reads a CANONICAL map, not `ROLE_SUFFIXES`.** That table is sorted longest-first for
  *parsing*; once `_APR_HIGHLY_CONFIDENTIAL` pointed at `APR`, a search returned it instead of
  `_APPROVER` and every new approver group was suggested with an HC name.
- **Live tables:** `Staging: UPL·APR·DELS·UPLHC·DELSHC` · `Documents: MEMBER·DEPTVIEW·GLOBAL·SEGVIEW·UPL·APR·DEL·SHARE·UPLHC`
  · `StagingHC: UPLHC·DELSHC·APR` · `DocumentsHC: UPLHC·MEMBERHC·APR·DEL·DEPTVIEW·GLOBAL·SEGVIEW·SHARE`.
- **MIGRATION: an existing HoD mapping must be RE-CREATED with the persona** — its rows still say
  `DEL`/`SHARE` and keep granting delete. It does not fail; it silently retains the power the client
  removed. **Reconciliation must be re-run**, and this is the one change that *reduces* access.

## HIGHLY CONFIDENTIAL — ITS OWN LIBRARY PAIR (2026-08-15, spec `2026-08-15-highly-confidential-library-design.md`)
> ⚠ **The role/clearance half of this section is SUPERSEDED by 2026-08-17 above** — `APRHC` and `hou_hc`
> no longer exist, `MEMBER`'s absence from `DocumentsHC` is now joined by `DELS`, and HC clearance is
> only for PIC and SDG Employee. **The library-pair ARCHITECTURE below is unchanged and still governs.**
Client: an HC uploader group files at **any** level and reaches **both** approval libraries;
`GHO_GF_CORU_UPL` reaches only the normal one. **SUPERSEDES `2026-07-16-highly-confidential-securing-design.md`**
— do not implement its elevated-flow design. BUILT, **not site-tested**; the libraries, the groups and
the two Power Automate flows do not exist yet.
- **Four libraries, two pairs:** `Approval Document` → `Documents`, and **`HC Approval Document` →
  `HC Documents`**. Same tree, same abbreviations, same content approval, same draft isolation.
  Created without spaces then retitled (gotcha #12), and **resolved at runtime, never hardcoded**.
- **A SEPARATE LIBRARY REMOVES THE PROBLEM THE JULY SPEC SOLVED.** That design kept HC in a secured
  subfolder and therefore needed an elevated service-account flow, a bearer trigger URL inside the
  SPFx bundle, and a window between folder creation and securing where a peer could see the file.
  Ordinary uploaders hold **no permission on the HC library at any moment**, so none of it is needed.
- **`naming.ts` HAS NO HC FALLBACK, and that is the safety.** Every other name falls back to a legacy
  literal; `cachedHcLibraries()` returns `undefined` and `hcAvailable()` is false. `libApiTitle`
  returns the KEY for an unresolved HC target, which **404s loudly** rather than resolving to the
  normal library. **Both halves resolve or neither does** — an HC approval library with no HC
  documents library accepts uploads and approvals and then has nowhere to route them.
- **⚠ THE `hcConfidentialityLevel` ROW IS THE GATE, NOT DECORATION — without it EVERY uploader sees
  the Highly Confidential level (found live 2026-08-19, fixed by adding the row; asserted in
  1.0.174.0).** `effectiveHcLevel` falls back to the LIBRARIES when the row is blank, and
  `hcAvailable()` resolves those libraries **by title** — which SharePoint **security-trims**. So an
  uncleared uploader gets `List 'HC Approval Document' does not exist`, HC reads as *not on this
  site*, `isHcLevel` is false for every level, `selectableLevels` filters nothing, and the one person
  who must not see the level is the one shown it. **A site with no HC at all gives the identical
  answer from the identical probe**, which is why nothing could tell them apart — `empty ≠ unknown`,
  in the place where the cost is disclosure.
  - **Setting the row fixes it outright, with no code change**: `effectiveHcLevel` returns a
    configured value regardless of library visibility, so `canOfferHc` then fails on `hcAvailable`
    and the level is hidden. `Title` = `hcConfidentialityLevel`, `ConfigType` = `setting`,
    `SettingValue` = the confidentiality term's own label.
  - **Reconciliation now ASSERTS it every run** and reports `⚠ HC libraries exist but the
    hcConfidentialityLevel config row is MISSING — the Highly Confidential level is NOT being gated`.
    Reported, never repaired: the label must match the client's own term, and guessing it would hide
    a level that is legitimately selectable on a site whose term is named differently.
  - **The exposure is the LABEL, not the documents.** With `hcLevel` blank nothing routes to HC, so
    an uncleared uploader's file lands in the NORMAL approval library and they gain no HC access.
    Still wrong: it tells them the classification exists and their unit may hold documents under it.
- **THE LIBRARIES DECIDE WHETHER ROUTING IS ON; the `hcConfidentialityLevel` config row only RENAMES
  the level** (`effectiveHcLevel`). Both simpler defaults are wrong in opposite directions: defaulting
  to `Highly Confidential` unconditionally **hides** that level on every site without HC, where it is
  an ordinary label anyone may pick; defaulting to blank lets a site that created the libraries but
  missed the row file HC documents into the **normal** library.
- **`hcRouting.ts` FAILS CLOSED, against this codebase's habit** (gotchas 10b/11, the provisioned-path
  filters). An over-hidden segment blocks an upload someone retries a minute later; an over-offered HC
  level publishes a secret and nobody finds out. `routeFor` returns **undefined** rather than falling
  back — falling back is the single worst thing it could do.
- **CLEARANCE IS A WRITE PROBE, NEVER GROUP MEMBERSHIP.** Reconciliation grants folder ACLs in a pass
  separate from group creation, so a `_UPL_HC` group can exist for days before it grants anything —
  the origin of the upload form's first 403. `probeFolderUploadAccessByPath` asks the HC folder
  directly (HC folders have **no Folder Map rows**; the path comes from `swapLibrarySegment`). Cached
  per leaf term, because clearance is per unit. Read `Low` **arithmetically, never with `&`**.
- **THE HC UNIT FOLDER IS RESOLVED, NEVER CREATED.** One created by an uploader inherits the LIBRARY
  root's permissions instead of carrying the unit's — exactly how an HC document becomes readable by
  the people the unit ACL excludes. Absent ⇒ refuse and name reconciliation. **Below** the unit,
  ensure-creation stays correct: that tier inherits by design.
- **`LIBRARY_ROLES` IS THE FEATURE, and what matters is what is ABSENT.** `UPL` and `APR` appear in
  **neither** HC row. `MEMBER` is absent from `DocumentsHC`; `SEGVIEW`/`GLOBAL` are absent from both
  approval-side rows. `UPLHC`/`APRHC` are **supersets** — they cover the normal libraries too — and
  downgrade to **Read** on both approved-side libraries (`APPROVED_SIDE_LIBS`), or the HC archive
  becomes writable by every cleared uploader.
- **The HC roles reuse the PLAIN permission levels.** There is no `CRS Upload HC` and there must not
  be: the separation is which library the grant lands on, held in ONE table.
- **HoD and C-Level keep their EXISTING groups** (client's instruction), so their existing powers
  travel: a **HoD can delete** an approved HC document and **C-Level can share** one, with no request.
  Consequences of that instruction, not of the code — open questions in the spec §10.
- **Two new personas, `pic_hc` and `hou_hc`**, replacing (never accompanying) their plain
  counterparts. A `hou_hc` **approves their own HC uploads** — the `hou` caveat, worse here.
- **Group names accept BOTH spellings** (`_UPL_HC` and `_UPL_HIGHLY_CONFIDENTIAL`); the builder
  suggests the long form, matching `_UPLOADER`. **Precedence comes from the existing length sort** —
  `_UPL_HC` (7) beats `_UPL` (4) and `_HC` (3), so an HC group can never parse as a plain uploader.
  The legacy bare `_HC` suffix now parses as `UPLHC`. The old `HC` role is **retired**.
- **`allLibraryTitles()` — tier columns go in FOUR libraries.** The two-element literal it replaced
  was the bug waiting: one unknown field name fails the WHOLE `validateUpdateListItem` call, and
  auto-route drops whatever does not exist at the destination.
- **`ApprovalDocument` resolves its library from `?lib=hc`, else tries normal then HC.** Item ids are
  per-LIST, so `?itemId=` alone cannot say which library — the wrong one 404s or opens a DIFFERENT
  document with the same id. **The queue stays inside one library.** Give the HC library's Name-column
  formatting `&lib=hc`.
- **⚠ THE FILE IS PLACED BY FOLDER ID AND TAGGED BY LIBRARY TITLE, AND THOSE TWO MUST AGREE** (fixed
  1.0.170.0; **HC uploads from the form could never tag, on any site**). `GetFolderById` reaches into any
  library, so an HC document landed in `HC Approval Document` correctly — while `validateUpdateListItem`
  was hardcoded to `settings.stagingLibrary`, the NORMAL approval library.
  - **Item ids are per-LIST.** So that call looked for the new item's id in the wrong list: a 404 on a
    good day, and on a bad one it finds a DIFFERENT document holding that id and writes the metadata
    onto it, reporting success. The symptom was the generic *"Uploaded, but tagging metadata failed"*.
  - `uploadStagedFile` now takes the target library title, defaulted so no other caller changes, and the
    HC branch passes `libApiTitle("StagingHC")`.
  - **`BulkUpload` already had this right** (`targetListTitle()` switches on `uploadingHc()`), which is
    the tell: two upload paths, one pattern, one of them written by hand.
  - **The failure message now names the STATUS and the LIBRARY.** "tagging metadata failed" discarded
    both, and the causes need opposite fixes — 404 is the wrong title, 403 is permissions on that list,
    400 is a malformed payload. Distinguishing them by hand cost an hour on 2026-08-19.
- **Bulk Upload gets NO write probe**, deliberately — admin-only, and already existence-gated only
  because an `AddListItems` probe against `Documents` would empty the form for every PIC. It writes
  straight to the approved side, so HC there means **`HC Documents`**, path segment and metadata target.
- **Requests match `APR` *and* `APRHC`.** A unit whose approver is the HC one has no plain `APR` row,
  and matching `APR` alone leaves that queue permanently empty while requests pile up behind it.
- ⚠ **BOTH HC FLOWS WERE BUILT AND THE VERTICAL WAS VERIFIED END TO END ON 2026-08-19** (commit
  `e58e97c`) — upload → HC Approval Document → approve → HC Auto Route → HC Documents, with the tier
  metadata, `Highly Confidential`, `LegallyPrivileged` and `Created By` all preserved. **This line
  used to read "TWO NEW POWER AUTOMATE FLOWS ARE REQUIRED and nothing works without them", and on
  2026-08-22 that stale sentence was repeated three times in one session as a reason HC could not
  work.** Exactly what the warning at the top of this file describes: detailed enough to be convincing
  when wrong. `HC Documents` must still have **content approval OFF**.
  - Getting there found **six faults, all the same shape**: the HC flow is a CLONE of Auto-route and
    every library reference it inherited was wrong. Two would not have failed loudly — item ids are
    per-LIST, so a stamp pointed at the wrong list writes onto a different document, and a **delete**
    pointed at the wrong list removes whatever holds that id there, which is data loss rather than a
    failure.
  - **`HC folder approval` should now be turned OFF** — both folder-approval flows became unnecessary
    on 2026-08-19 when draft isolation was reversed. A flow erroring on every folder creation burns
    the daily quota, and an exhausted quota means the next real approval is not routed, silently.
  - **Still open on the HC flow:** the approval email links to the SOURCE library, which the flow's own
    delete has already emptied, so the link is dead on arrival; and `Created` is stamped ~8 hours early
    on BOTH routing flows (a UTC value written in a locale format SharePoint reads as local time), so a
    document uploaded before 08:00 local lands on the previous day — and `Created` is what My
    Submissions sorts by.
- **MIGRATION IS OUT OF SCOPE AND MUST NOT BE SILENTLY SKIPPED.** Documents already labelled Highly
  Confidential sit in the normal libraries, readable by their whole unit. After deploy the label
  implies a protection they do not have — worse than before. Until a sweep runs, HC protection applies
  only to documents filed **after** deployment.

## ⚠ THE ARCHIVE LIBRARIES' METADATA COLUMNS WERE WRONG-TYPE, HALF-MISSING, AND MISNAMED (2026-09-02)
Found live while preparing to test the archive mover for the first time (see the runbook addendum
below). Someone had partially acted on the "mirror `Documents`' columns" fix this file already flagged
as deferred — but wrongly:
- **`Business_x0020_Segment`/`BusinessSegmentTid` were MISSING, replaced by a stray
  `GroupHeadOfficeArchive`/`GroupHeadOfficeArchiveTid` pair** that matches no naming convention this
  project uses anywhere. Deleted; recreated under the correct names.
- **`Document_x0020_Type` and `Year` existed as plain TEXT, not `TaxonomyFieldType`.** `Documents`' own
  equivalents ARE genuine managed-metadata columns (confirmed live via `SP.Taxonomy.TaxonomyField` on
  `Documents`'s fields, ruling out an earlier wrong guess that ALL of Business Segment/Department/
  Unit/Document Type/Year were taxonomy — only three of the five actually are). A text column with the
  same internal name silently accepts no taxonomy value from a `MoveTo`. Deleted; recreated bound to
  the correct SITE-WIDE term sets (`866c5754-…` / `023a866a-…`).
- **`Confidentiality_x0020_Level` was missing entirely.** Created, bound to `0d6d1da8-…`.
- **`Department`/`Unit`/their `Tid` twins were ALREADY CORRECT** (plain Text, matching `Documents`) —
  an earlier pass over this file wrongly assumed they needed fixing too; they did not.
- **Confirmed identical on BOTH `Archive Restricted & Confidential Document` and `Archive Highly
  Confidential Document`** — same wrong state, same fix, applied and verified on both via a live
  `/fields` read (internal names, types, and single- vs multi-value all confirmed correct afterward).
- **This had to be done by hand, on the site — it cannot be done by code**, per this file's own
  standing rule: a wrong taxonomy binding looks right and silently tags nothing, so `ensureColumn`
  deliberately never creates one. The plain-text pair (`Business_x0020_Segment`/`BusinessSegmentTid`)
  COULD have been code-created but was done by hand alongside the taxonomy fixes for one clean pass.

## ✅ THE ARCHIVE MOVER WORKS, BOTH VERTICALS, VERIFIED LIVE — CLAUDE.md'S OWN "NOT BUILT" WAS STALE (2026-09-02)
The section immediately below this one still says "MOVER NOT BUILT" — that was WRONG the moment this
was checked. **`CRS — Archive after seven years` already existed**, built by Clarence on 2026-08-23,
sitting there the whole time this file kept saying otherwise. It had two real bugs (below), both now
fixed, and both verticals are now verified end to end with real production data.
- **⚠ THE EXISTING FLOW WAS BROKEN BY THE 2026-08-28 LIBRARY RENAME, AND HAD NEVER SUCCEEDED ON REAL
  WORK.** `Get items`'s "List Name" stored the literal title `"Documents"` — broken the day the rename
  landed, producing `List not found` on every scheduled run since. Its `$filter` was ALSO still set to
  a single leftover test filename from manual testing, never restored. The one "Succeeded" run in its
  history (2026-08-28) turned out to have found **zero** matching items — a trivial, meaningless
  success, not evidence the mechanism worked. **No file had ever actually been moved by this flow
  before today.**
- **FIXED: `Get items` re-pointed at the current title, filter set to the test cutoff
  (`Created lt datetime'2026-09-01T00:00:00Z'`), Pagination turned ON with `minimumItemCount: 5000`
  (a SEPARATE setting from `$top` that silently caps a run if left off — easy to miss).** The two
  `getbytitle('Archive')` references inside the loop (`Stamp_Archived`, `Reset_inheritance`) were
  updated to the current title. **`Create new folder` and the `moveto` destination were ALREADY
  correct** — they use the library's internal list GUID and the URL segment respectively, neither of
  which the rename touched. Only title-based references broke; segment/GUID-based ones did not.
- **✅ RUN 1 (`Documents` → `Archive`): 57 of 57 files moved, zero clashes, zero skips** — confirmed by
  cross-checking `Get items`'s count against the library's `Archived = Yes` filtered count, which
  matched exactly. **Every fixed metadata column carried a real value**: Document Type, Year,
  Confidentiality Level (the three taxonomy fields fixed in the section below this one), and Business
  Segment/BusinessSegmentTid all populated correctly on every archived file, alongside the
  already-correct Department/Unit.
- **✅ RUN 2, the HC CLONE (`HC Documents` → `HC Archive`): 17 of 17 files moved, same clean result.**
  Built via Save As from the now-fixed normal flow, per this project's standard HC-clone pattern.
  **⚠ ONE REFERENCE WAS MISSED ON THE FIRST PASS, caught before running** — `Find_the_moved_item`
  still read `/Archive/` and `'Shared Documents/'` after the other five swaps were done. Left
  unfixed, the file would have moved into `HCArchive` correctly but then silently failed to get its
  `Archived` stamp or its permission reset, since both of those actions require `Find_the_moved_item`
  to succeed first — the exact "moved but half-tagged" failure shape this project's HC clones keep
  producing. **Six references need checking on this clone, not five**: `Get items`' library,
  `RelPath`'s replace, `Create new folder`'s library, the `moveto` destination, `Find_the_moved_item`'s
  path, and the two `getbytitle` calls in `Stamp_Archived`/`Reset_inheritance`. Confirmed Highly
  Confidential metadata carried through correctly too (`Confidentiality Level: Highly Confidential`,
  correct segment/department/unit values across both GHO and MHO).
- **Runbook used: `2026-08-22-seven-year-archive-mover-runbook.md` + the test-cutoff addendum
  immediately below.** Both flows are now correctly built against a fixed 2026-09-01 cutoff rather
  than the real rolling 7-year calculation — **do not mistake either for the real scheduled mover.**
  Both were turned back OFF after this test (client: archived files stay archived, this is the testing
  site — no revert needed, but no reason to leave them scheduled against a permanently stale cutoff
  either).
- **STILL NOT DONE: promoting either flow into the real recurring mover** (swap the `CutOff` Compose
  back to the rolling `addDays(utcNow(), -2557, ...)` calculation, restore the 100-file cap, switch
  the trigger philosophy back to genuinely scheduled use, re-verify on SDG's tenant with its own list
  GUIDs). That is a deliberate later step, not implied by today's test.

## ⚠ ARCHIVE ACCESS NARROWED TO C-LEVEL ONLY, AND AUDIT LOGGING FOR IT IS STILL A TODO (2026-09-02, 1.0.363.0)
Client, the same day both archive movers were proven working end to end: *"we need to make sure
Archive is also track in audit log and client update, only C level and system administrator have
access to archive, so no more HOD till Normal viewer have access to Archive Library, and that would
mean the My Submission should not update to archive either."* Runbook:
`docs/superpowers/specs/2026-09-02-archive-audit-log-and-access-narrowing-runbook.md`.
- **REVERSES THE 2026-08-22 RULE DIRECTLY.** `LIBRARY_ROLES.Archive`/`.ArchiveHC` in
  `FolderManager.tsx` went from every persona (MEMBER, DEPTVIEW, DEL, SHARE, UPL, APR, UPLHC, APRHC,
  ...) down to **`["GLOBAL", "SEGVIEW"]`** — the two C-Level roles, and nothing else. "System
  administrator" needed no row: `CRS Owners` holds Full Control on the web regardless of any folder
  ACL, the same reason the admin-page lockdown never has to name it.
- **⚠ THIS DOES NOT REVOKE ANYTHING ALREADY GRANTED, and that matters here specifically because the
  archive was JUST used for real (today's 57 + 17 test files).** Reconciliation only ever ADDS
  folder-scope grants — `groupsToRemove`'s full-ACL assertion is page-scope only, because a folder
  root also carries SharePoint's automatic Limited Access entries for principals granted deeper in
  the tree, and asserting there would strip those too. So ClarenceDMSTesting's already-provisioned
  Archive/ArchiveHC folders keep every HOD/PIC/HOU/MEMBER group's Read from before this change until
  someone removes it by hand; **narrowing `LIBRARY_ROLES` only changes what a FUTURE run grants.**
  SDG's tenant is unaffected by this caveat — nothing is archived there yet, so the narrower rule
  applies cleanly from the first real archive onward.
- **MY SUBMISSIONS NO LONGER READS THE ARCHIVE AT ALL** (was added 2026-08-22, removed the same day
  as the access narrowing). My Submissions is a PIC/HOU page and neither role holds `GLOBAL` or
  `SEGVIEW`, so the read would now just 403/404 for every viewer — removed rather than left to fail
  silently. `cachedArchiveLibraries()` is still read (a synchronous cache lookup, no request) because
  `isHcRecord` still needs the archive HC library's title to classify a `CRS Submissions` record.
- **`EventType: "Archived"` IS REGISTERED IN CODE** (`auditLog.ts` — `EVENT.archived`, `EVENT_LABEL`,
  `ALL_EVENT_TYPES`), so it is filterable in the Audit Log viewer's Action dropdown the moment a row
  with that value exists — no schema change needed, `EventType` is plain Text.
- **⚠ NEITHER ARCHIVE FLOW WRITES THAT ROW YET.** Both `CRS — Archive after seven years` and its HC
  clone move files and stamp `Archived`/reset inheritance, but log nothing to `CRS Audit Log`. The
  runbook above specifies the exact `Create item` action to add to each (placement: after
  `Reset_inheritance` succeeds; fields: `EventType: Archived`, `Source: Flow:ArchiveMover`/
  `Flow:ArchiveMoverHC`, `ItemUniqueId` — confirmed today that `MoveTo` preserves it exactly, so a
  future "History of this file" feature can thread Uploaded → Approved → Routed → Archived on one
  GUID). **Not built. Today's 57 + 17 archived test files produced no audit rows.**
- **✅ THE CATCH-UP FOR ALREADY-GRANTED ARCHIVE ACCESS IS BUILT — a one-time repair button on the
  Reconciliation tab, "Prune archive access to C-Level only".** Client, asked how to clean up
  ClarenceDMSTesting's already-provisioned test archive: *"Build a one-time repair tool."* Walks
  every `Archive`/`Archive Highly Confidential Document` folder up to segment/department/unit depth
  (the only levels reconciliation itself ever grants at — below Unit inherits by design and is
  never walked) and removes any group NOT currently mapped `GLOBAL` (site-wide) or `SEGVIEW` (for
  that folder's own segment, derived from `loadReconModes()` matching the folder's top path segment
  against `StagingFolder`) in `CRS Group Map`.
  - **SAFE BY CONSTRUCTION, reusing `getRoleAssignments`**, which already excludes site Owners and
    Limited Access bindings (`RoleTypeKind` 1/7) — so a navigation-only corridor entry for a group
    still holding a REAL grant elsewhere is never touched by this alone, and because the pass covers
    EVERY folder in both libraries in one run, a group left with no real grant anywhere in
    Archive/ArchiveHC also loses whatever corridor entries existed only to reach it. End state
    matches a from-scratch build under the new rule.
  - **`ownerGroupId` IS RESOLVED FRESH (`resolveOwnerGroupId()`) AND READ FROM THE RETURN VALUE**,
    never the state the same run just set — the identical stale-closure trap `runReconciliation`'s
    own `roleDefs`/`ownerGroupId` fixes exist for (1.0.237.0/1.0.293.0), paid for once already in
    this file and avoided here by construction rather than by remembering.
  - **`removeRoleAssignment` WAS "UNREACHABLE ON PURPOSE" SINCE 2026-08-04-ish** (`revokeAncestorRead`
    stays hard-off for the ORDINARY reconciliation run, deliberately) — this is its first real
    caller. The stale `eslint-disable-next-line @typescript-eslint/no-unused-vars` above it is
    removed, since it is genuinely used now.
  - **A ONE-CLICK CONFIRM BAR, NOT `window.confirm()`** — matching this project's established
    aversion to native browser dialogs (the replace-clash popups were removed for exactly this
    reason). Reports every removal by folder + group name, matching the "named, not counted"
    convention of every other prune pass in this file.
  - **⚠ DOES NOT TOUCH SDG's TENANT** (nothing archived there yet, so nothing to prune) and **does
    not itself re-grant anything** — the report explicitly says to re-run Folder Reconciliation
    afterwards to confirm the surviving C-Level groups are still correctly granted.
  - Rendered only when `cachedArchiveLibraries() !== undefined`, so it is invisible on a site with
    no archive. **Verified**: `tsc --noEmit` clean, full suite 1598/0, 35 warnings (matches the
    established baseline, no new ones). Packaged as `1.0.364.0`.
  - **✅ RUN LIVE ON BOTH SITES, 2026-09-02, AND REMOVED THE SAME DAY.** ClarenceDMSTesting: 97
    grants removed, all on `Archive Highly Confidential Document/MHO`'s segment root (nothing
    below it needed touching — that whole subtree inherits from the segment root rather than
    breaking its own inheritance per department/unit, unlike `Documents`/`HC Documents`). **SDG's
    real tenant: run across all three segments (GHO, MHO, NBPOLHO) and both libraries** — hundreds
    of `<SEG>_..._UPLOADER`/`APPROVER`/`VIEWER`/`HOD` grants removed, **`CRS Owners` never appeared
    in either run's log** (the owner-protection held). **One recurring non-group entry, "Guthrie
    Central Repository System" (the site's own title), was removed once per segment root on SDG**
    — almost certainly a SharePoint auto-created default group from initial site provisioning,
    never renamed the way `CRS_SITE_MEMBERS`/`CRS Owners` deliberately were, holding one blanket
    ancestor-level Read grant per segment. Consistent with a stray leftover, not a mistake — but
    **not independently confirmed** what that principal actually is; if anything about archive
    access looks wrong later, start there.
  - **REMOVED THE SAME DAY, PACKAGED AS `1.0.365.0`** (client: *"we don't want to confuse the
    client with a new feature like this"*). The button, `runArchivePrune`, and its three state
    variables are gone from `FolderManager.tsx`; `removeRoleAssignment` reverted to its prior
    `eslint-disable-next-line` "unreachable on purpose" state, now WITH a note that it has a real
    history (this tool) rather than being purely hypothetical. **Kept, not deleted** — the runbook
    (`docs/superpowers/specs/2026-09-02-archive-audit-log-and-access-narrowing-runbook.md`) has the
    exact shape if this class of repair is ever needed again (e.g., a fourth segment onboarded
    under the old rule before someone remembers to check).
  - **⚠ NEITHER SITE HAS RE-RUN FOLDER RECONCILIATION SINCE THE PRUNE, and that is the outstanding
    step.** The prune only removes; confirming the surviving GLOBAL/SEGVIEW groups (and Owners)
    are still correctly granted requires the ordinary reconciliation run, same as any other
    folder-ACL change in this project.

## ⏭ ARCHIVE MOVER TEST — RUNBOOK ADDENDUM WRITTEN, FLOW NOT YET BUILT (2026-09-02)
Client (Crystal, relayed by Clarence): *"ALL FILES that was uploaded before 1 September 2026, 00:00:00
push it to archive"* — a deliberate bulk test of the mover, not a scheduled run. The base mover runbook
(`2026-08-22-seven-year-archive-mover-runbook.md`, below) already covers the real 7-year flow in full
detail; **`docs/superpowers/specs/2026-09-02-archive-test-cutoff-runbook.md` lists only the six deltas**
for this test variant — read the base runbook first, then the addendum for what differs:
1. The `getbytitle(...)` calls need the POST-RENAME library titles (`Archive Restricted & Confidential
   Document` / `Archive Highly Confidential Document`) — the base runbook predates the 2026-08-28
   rename. Path derivation itself is UNAFFECTED (it uses URL segments, which never changed).
2. `CutOff` is a fixed literal (`2026-09-01T00:00:00Z`), not the rolling 7-year calculation — named
   `CutOff_TEST_20260901` in the flow so a future clone-into-the-real-mover can't silently inherit it.
3. The 100-file weekly cap is REMOVED for this one-off run (confirmed with the client) — **and
   `Get items`' Pagination must be turned ON with a raised Threshold**, or the action silently caps
   itself at its own default page size regardless of `Top Count`.
4. Trigger is "Manually trigger a flow" (instant), not Recurrence — this build's only current purpose
   is the manual test.
5. ✅ **CONFIRMED: no revert needed.** *"It can be stayed as archive, we are still on testing site"* —
   the archived files stay archived after this test. There is still no built-in undo mechanism in this
   project if that ever changes; base runbook §6 is explicit it would be a manual, file-by-file job.
- **Neither the base runbook nor this addendum has been built into an actual flow yet.** Build order:
  read `2026-08-22-seven-year-archive-mover-runbook.md` end to end, then apply the five deltas above
  from the addendum.

## THE SEVEN-YEAR ARCHIVE (2026-08-22, 1.0.231.0) — FOUNDATION BUILT, MOVER NOT BUILT
Client: *"After 7 years the file in the Documents Library should be archive. So we need another
library"*, then *"knowing client they want the same group to have access"*, then *"for all the Archive
files make sure all groups have read only."* Spec
`docs/superpowers/specs/2026-08-22-seven-year-archive-design.md`.
Two libraries, **`Archive`** and **`HC Archive`**, created on ClarenceDMSTesting 2026-08-22 with
content approval OFF. **Built: resolution, reconciliation, permissions, My Submissions, CRS Search,
the request refusal. NOT built: the mover.** Nothing has been archived anywhere.
- ⚠ **NOTHING ON THIS SYSTEM IS CLOSE TO SEVEN YEARS OLD.** The oldest document dates from July 2026,
  so the first file becomes eligible in **2033**. The libraries sit empty until then, and the feature
  can only be tested with a BACKDATED `Created` — `validateUpdateListItem` with
  `bNewDocumentUpdate: true`, which wants `M/d/yyyy h:mm tt`, not ISO. Do not read an empty archive as
  a broken mover.
- **ONE TREE, NO YEAR PARTITION, and a collision is resolved by RENAMING** (`report (2026).pdf`, the
  suffix carrying the creation year). The client wanted a year folder on top and the collision it
  guards against is real — a name freed by one file archiving can be reused and archive years later.
  But **reconciliation cannot build a year tree**: it walks the TERM TREE and there is no year in the
  term store, so a year on top is a complete copy of the whole tree per year, each needing its own
  broken inheritance and grants, growing for ever. Scopes are fine (~130/year against a 50,000
  ceiling); RUN TIME is not. Agreed fallback if the client rejects the rename: notify an administrator
  and move nothing. **Browsing by year is a VIEW grouped on `Created`, not folders.**
- **EVERY ROLE IS READ, via `READ_ONLY_LIBS` in `FolderManager.tsx` — a LIBRARY rule, never an
  extension of `DOCUMENTS_READ_ONLY_ROLES`.** That list applies across `APPROVED_SIDE_LIBS`, so adding
  `DEL`/`SHARE` to it — the obvious-looking edit — would silently strip a Head of Unit's delete and
  share on `Documents`, reversing the client's 2026-08-20 decision as a side effect. The archive is
  read-only because of what the LIBRARY is, not which roles reach it. It still answers `undefined` for
  an unknown role, so `ENTRY` keeps being skipped rather than approximated.
  - **Tell the client: ARCHIVING REMOVES THE ABILITY TO DELETE OR SHARE.** `DEL` and `SHARE` are on
    the rows but resolve to Read. Almost certainly what "archive" means; it should not be discovered.
  - `UPL`/`UPLHC` are on the rows AT READ on purpose — a PIC must still find their own archived files.
    `SEGVIEW`/`GLOBAL` are safe here though absent from both approval-side rows: the archive holds only
    approved documents, so there is no draft for a segment-wide viewer to reach.
- **⚠ A DELETION OR SHARE REQUEST AGAINST AN ARCHIVED FILE IS REFUSED IN `validateDraft`, and this is
  the piece that fails silently if it is missed.** The approval executes in the APPROVER'S OWN
  SESSION, and they hold only Read there — so an approved request would fail at the last step,
  recorded `Failed`, days after the requester was told it was being handled. Refused in the RULES, not
  hidden in the UI, for the same reason as the pending-share refusal beside it. Both types, and it is
  the ONLY message returned.
- **THE PAIRING RULE IS NOT HC's, deliberately.** HC is both-or-neither. The archive **mirrors the
  site's HC state**: an HC site needs BOTH archive libraries (or HC documents silently never archive),
  a non-HC site needs only `Archive`. Copying HC's flat rule would disable archiving on a correctly
  provisioned non-HC site. `primeArchiveLibraries` therefore MUST run after `primeHcLibraries` —
  reading `hcAvailable()` too early classifies an HC site as non-HC and accepts half an archive.
- **`archiveTargetFor` FAILS CLOSED** — `undefined`, never a fallback. A fallback would move a Highly
  Confidential document into the open archive. An approval library gets no archive at all: a document
  that was never approved is not a record.
- **⚠ THERE WERE TWO `LibTarget` UNIONS, and reconciliation read the wrong one.** `FolderManager.tsx`
  declared its own copy beside the one `naming.ts` exports. Both said the same four keys so nothing
  ever failed — but `libApiTitle`, `libraryTargets` and `allLibraryTitles` switch on the SHARED one
  while `LIBRARY_ROLES` and every progress feed indexed the LOCAL one. **Adding a library to
  `naming.ts` would have compiled cleanly and been silently skipped by reconciliation** — no folders,
  no grants, on a run reporting success. Widening a union is only a loud failure if there is ONE
  union. It is an import now; do not re-declare it.
- **`allLibraryTitles()` and `libraryTargets()` both include the archive**, and both must. The first
  because a MOVE carries over only columns that EXIST at the destination — the
  `Remark`/`LegallyPrivileged` gap of 2026-08-10, worse here because an archived document is a record
  nobody opens for years. The second because `SubtreeMigrator` reads it to decide which libraries a
  structure change re-shapes, and an omitted archive repeats register #15 exactly.
- **`Archived` (Yes/No) is asserted on every library every run**, but `isArchivedRow` derives the tag
  from the **LIBRARY**, not the column: the column is a convenience for views and search, the library
  is where the file actually is and is what decides what a reader can do with it.
- **My Submissions and CRS Search both read the archive now.** Missing either is how an uploader's own
  seven-year-old document silently vanishes from their history, or becomes unfindable — and nobody
  reports that as a bug, they conclude it was deleted. The archive segments are also in `libs`, or the
  folder trail reads `Archive › GHO › …` with the library name presented as a folder.
- ⚠ **THE MOVER WORKS AND IS VERIFIED ON SITE (2026-08-23), BUT AN ARCHIVED DOCUMENT LOSES ALL ITS
  METADATA — CONFIRMED, NOT THEORETICAL.** `Archive` has only `Archived`, `SubmissionId`, `BatchId`,
  `BulkImport` and `Full Name`; asking for `Year` returns *"The field or property 'Year' does not
  exist"*. A move carries over only columns that EXIST at the destination, so Document Type, Year,
  Confidentiality Level, LegallyPrivileged, Remark, Vendor/CustomerName, DocumentDate and **every tier
  column and `Tid` field** are stripped. The `Remark`/`LegallyPrivileged` gap of 2026-08-10, at a new
  library.
  - **It is worse than a display problem: CRS Search filters on those columns**, so an archived
    document sits inside the KQL scope and is still unfindable by metadata. In an archive, metadata is
    the only way to find anything.
  - **THE FIX IS TO MIRROR `Documents`, NEVER A HARDCODED LIST.** Tier columns differ per segment and
    new ones appear on every onboarding, so a literal would be right today and silently wrong after the
    next one — the two-element-literal bug that cost the HC pair four days. Read `Documents`' fields and
    create whatever is missing, skipping hidden, read-only and system fields.
  - ⚠ **The three TAXONOMY columns cannot be created that way** (`Document Type`, `Year`,
    `Confidentiality Level`): a taxonomy field needs its hidden note field and a term-set binding, and a
    wrong binding yields a column that looks right and silently tags nothing. **Report them by name and
    add them by hand**, like the `hcConfidentialityLevel` assertion.
  - **DEFERRED until after the SDG migration** (client's call, 2026-08-23): nothing archives until 2033,
    so the loss is not yet reachable. Do it on ClarenceDMSTesting first, then push to SDG.
- ⚠ **BEFORE DEPLOYING TO ANY SITE: CHECK FOR A NAME COLLISION ON `Archive`.** It is a generic name, and
  a pre-existing library called `Archive` or `CRSArchive` would resolve as ours — reconciliation would
  break its inheritance, build ~140 CRS folders inside it and grant every group Read, on a library
  nobody offered. One call answers it:
  `/_api/web/lists?$select=Title&$filter=Title eq 'Archive' or Title eq 'CRSArchive' or Title eq 'HC Archive' or Title eq 'HCArchive'`
  **Verified EMPTY on SDG's tenant 2026-08-23** — they have Documents, Approval and the HC pair only, so
  the archive stays dormant there and 1.0.231.0 is safe to deploy.
- **On a site with no archive libraries the feature is inert**: `cachedArchiveLibraries()` stays
  `undefined`, so `reconLibs()`, `allLibraryTitles()` and `libraryTargets()` are unchanged, My Submissions
  makes no extra reads, and CRS Search adds no scopes. The only footprint is the `Archived` column, created
  with `Options: 8` so it joins no view.
- **THE MOVER IS NOT BUILT — runbook `2026-08-22-seven-year-archive-mover-runbook.md`.** It must use **`MoveTo`, not Auto-route's copy-stamp-delete**: a copy is a
  new file and therefore a NEW `UniqueId`, which breaks the audit trail at the archive boundary.
  ⚠ **Auto-route already does this today** — every document in `Documents` has a different `UniqueId`
  than it had in `Approval Document`. It has not bitten only because the file-lifecycle audit flows
  were never built.
  ✅ **VERIFIED ON SITE 2026-08-22: A CROSS-LIBRARY MOVE PRESERVES `UniqueId`.** `accdbebf-…a992` read
  identical before and after moving from `Documents` into `Archive`; the LIST ITEM id changed
  (6095 → 142, because item ids are per-list) and `Modified` was NOT restamped. So the mover uses plain
  `MoveTo`, needs no GUID-mapping row, and the audit trail stays continuous across the archive boundary.
  ⚠ **Whether an individual SHARE survives is still unknown, and no longer matters:** the mover resets
  the moved file to inherit UNCONDITIONALLY — a no-op if shares die, and the client's stated read-only
  rule if they live. There is no revoke feature, so a surviving share would be permanent and invisible.
  **Tell the client archiving ends individual shares.**

## A HEAD OF UNIT COULD UPLOAD AND NOT SEE WHAT THEY UPLOADED (2026-08-21, 1.0.216.0)
`My-Submissions.aspx` was keyed `["UPL"]`. **`hou` carries no literal `UPL`** — it is
`APR, DELS, DEL, SHARE, UPLHC, DELSHC`, and `UPLHC` is a SUPERSET that
`LIBRARY_ROLES.Staging` lists on the **normal** approval library too. So a Head of Unit held
`CRS Upload` on `Approval Document`, reached the upload form (via `APR`), filed a document — and was
then denied the only page that lists it, **and with it the only route to a deletion or share request
about their own file.** Now `["UPL", "UPLHC", "APR"]`.
- **THIRD PAGE KEYED ON A ROLE ITS AUDIENCE DOES NOT LITERALLY HOLD** — upload form needed `APR`
  (2026-08-17), Requests needed `DEPTVIEW` (2026-08-21), this needed `UPLHC`. The pattern: a page rule
  names the role a persona is *thought* of as having, while `PERSONAS` gives it the superset or the
  management role instead. **Check the persona's actual `roles` array, never the family name.**
- **THREE THINGS MUST LINE UP AND ONLY ONE WAS WRONG:** reaching the form (page policy), writing into
  the folder (`LIBRARY_ROLES`), and seeing your own submissions (page policy). Uploading was never
  broken — which is why nothing failed and nobody noticed.
- **`APR` as well as `UPLHC`**, same argument as the upload form: a HoU group mapped before the
  2026-08-15 persona correction holds `APR` + `DELS` and no upload role at all, so keying on the
  upload roles alone would strand every already-provisioned Head of Unit.
- **⚠ WIDENING THIS CANNOT EXPOSE ONE PERSON'S FILES TO ANOTHER.** The page reads BOTH libraries
  `AuthorId eq <me>`, so a granted role sees only its OWN rows — the grant opens the page, authorship
  decides the content. The previous comment called approvers *"the people it is private from"*, which
  was **wrong**: it is private from them for everyone else's files whether or not they can open it.
  `DELS` stays off, not for safety but because a role that files nothing gets a permanently empty page.
- **MIGRATION = re-run reconciliation.** Page grants are derived and asserted, so the run adds the
  `_APPROVER` groups to that page. No row or schema change.

## MY SUBMISSIONS GROUPS BY SUBMISSION → BATCH → FILE (2026-08-22, 1.0.221.0)
Client: *"they want My Submission to detect or to record the submission of the uploader themselves…
they click the submission and it will bring them in showing the batches and then they can open each
batches and it will show the details of each file."* Spec
`2026-08-22-submission-grouping-design.md`; rules in `shared/submissionGroups.ts` (pure, 14 tests).
- **⚠ THE BLOCKER WAS THAT A BATCH LEFT NO TRACE.** `Batch.id` in `uploadBatches.ts` is documented
  *"stable within a session"* — batches live in the browser and were **never written to SharePoint**.
  That is why the page listed loose files: there was nothing to group by. Two text columns now carry
  it, `SubmissionId` (one per press of Upload) and `BatchId` (one per destination within it).
- **⚠ ALL FOUR LIBRARIES, and this is the trap.** SharePoint's copy carries over only columns that
  EXIST at the destination, so a file routed by Auto-route into `Documents` would arrive with its
  reference stripped and a submission would lose its files one by one **as they were approved**. Same
  shape as the `Remark`/`LegallyPrivileged` gap of 2026-08-10.
- **⚠ THE WRITE IS CONDITIONAL, OR IT DESTROYS ALL METADATA.** `validateUpdateListItem` fails the
  WHOLE call on one unknown field name (gotcha #4) — every column lost, not just the missing one. So
  `Form.tsx` reads `/fields` once per library (`libraryHasRefColumns`, cached, **false on a failed
  read**) and stamps only where both columns are confirmed. Absent ⇒ the upload behaves exactly as it
  did before, ungrouped. **Degrading to yesterday's behaviour is always correct; losing a document's
  metadata never is.**
- **Reconciliation ASSERTS the columns every run** over `allLibraryTitles()`, rather than relying on a
  provisioning step — a library added later would otherwise stop grouping with nothing to say so.
  Fifth instance of that rule. Reported, never fatal.
- **The read asks for them and retries WITHOUT on failure**, like `Stage` on the requests list: asking
  unconditionally would empty this page with a 400 on any site that has not reconciled. The page then
  says why, as a warning rather than an error — grouping is an addition, and the list still works.
- **A ROW WITH NO `SubmissionId` IS ITS OWN SUBMISSION OF ONE FILE** (client's choice). Blank is
  *unknown*, not a value: three referenceless files are three submissions, never one group of three.
  Inferring groups from timestamp and folder was rejected — two unrelated files a minute apart would
  be shown as a submission that never happened. Every pre-2026-08-22 file is in this state.
- **A submission is NEVER collapsed to one status.** It can be part approved and part rejected, and
  one badge for it is how an uploader concludes a rejected file was fine. `statusCounts` returns all
  three and the view shows only the non-zero ones.
- **References are readable — `SUB-20260822-K4P2`** — because an uploader quotes them to their
  approver, as the client's own claim system does. `I L O 0 1` are absent from the alphabet: this gets
  read aloud and written down. A GUID would be neither.
- **Levels 2 and 3 are VIEWS, not reads.** The rows are already loaded; drilling in filters what is on
  screen. Only the existing per-file detail panel reads `FieldValuesAsText`, which is a per-ITEM
  endpoint and can only ever be one file at a time.
- `BulkUpload` is deliberately NOT stamped: admin-only, writes straight to the approved side, and is
  not "the submission of the uploader themselves". Its files fall into the no-reference case.
- ⚠ **`s` in this file is a `Record<string, CSSProperties>`** — a key that does not exist yields
  `undefined`, so the build stays green and the element renders unstyled. Four styles were referenced
  before being declared; any new one must be added to that object.

## ⚠ THE COLUMN CHECK BEHIND THE SUBMISSION STAMP WAS CAPPED AT 500 FIELDS (2026-08-22, 1.0.223.0)
`libraryHasRefColumns` in `Form.tsx` read `/fields?$select=InternalName&$top=500` and searched the
result for `SubmissionId`/`BatchId`. **A truncated read is indistinguishable from an absent column** —
the `fetchAllSiteGroups` defect of 2026-08-21 by another route, one day later.
- **WORSE HERE THAN THERE.** A document library carries hundreds of fields, and these two are the
  NEWEST on the list — so they sit at the END of the collection and are the first thing a cap removes.
- **THE SYMPTOM IS SILENT AND TOTAL:** the check answers false, the reference is deliberately not
  written (it must not be — one unknown field name fails the whole `validateUpdateListItem` call), the
  upload succeeds with every other column correct, and **every submission is grouped by folder and date
  for ever** with nothing anywhere to say why.
- **Now `$filter=InternalName eq 'SubmissionId' or InternalName eq 'BatchId'`** — at most two rows,
  which cannot be truncated. **Filter server-side; never page a collection whose size you do not
  control and then search it.**
- **A FAILED filter falls back to `$top=5000` rather than concluding the columns are absent** — insurance
  against the FIX. If `$filter` on `InternalName` were ever rejected, the failure would be the original
  silent one, and we would have replaced one invisible cause with another. `undefined` (not read) is kept
  distinct from `[]` (read, empty) so the fallback fires only where it can help.
- **CONFIRMED ON SITE 2026-08-22:** `/items?$select=…,SubmissionId,BatchId` returned rows with
  `m:null="true"` — the columns EXIST and the values were never written. That distinction is the whole
  diagnosis: a 400 would have meant an unreconciled library, and `null` means the write was refused.
- ⚠ **A blank stamp is INVISIBLE ON THE UPLOAD FORM.** `refsMissing` on My Submissions only fires when
  the READ 400s, which says the columns are absent from that library — it says nothing about whether a
  given file was stamped when it was written. So a file uploaded before reconciliation ran looks
  identical to one the check wrongly refused to stamp. **Diagnose by reading the item, never by the
  banner:** `/items?$select=Id,FileLeafRef,Created,SubmissionId,BatchId&$orderby=Id desc&$top=10`.

## ⚠ OLD FILES ARE GROUPED BY FOLDER AND DAY — THE REFUSAL TO INFER WAS OVERRULED (2026-08-22, 1.0.222.0)
The spec rejected inferring submissions from timestamp and folder, and the client overruled it on
seeing the result: *"Can we not show per file instead… client doesn't want to go through each file one
by one figuring out and noting down which one."* One row per referenceless file **is** the flat list
this whole feature exists to replace, and on a site with a year of history that is every row.
- **FOLDER *AND* DAY, never folder alone.** A unit files into the same folder every month, so folder
  by itself would present a year of unrelated uploads as one submission. Same destination on the same
  day is the strongest evidence available, and it is the shape a real batch has.
- **The cost is stated on screen, not buried:** two people filing into one folder on one day are shown
  as one group. `SubmissionGroup.inferred` carries it and the row reads *"Grouped by folder and date —
  uploaded before submissions were recorded"*. Nothing is hidden and no file moves — only how the rows
  are stacked. **A guess presented as a record is the failure mode; a guess LABELLED is not.**
- **A row with NO date groups with nothing** — a missing timestamp is unknown, and unknown must never
  widen a group. Empty ≠ unknown, again.
- **Stamped and inferred keys live in separate spaces** (`~` prefix), so a stored id can never collide
  with a derived one, and a real submission is never merged into a guessed one.
- **THE ROW IS LABELLED BY ITS DESTINATION, NOT ITS REFERENCE** (client's own sketch:
  `GHO › GF › TAX › 2024 › Term Sheet › Archive 2`). A reference is what you QUOTE to an approver; a
  path is what you RECOGNISE, and recognising the one you want is the whole job of this list. The
  reference stays beneath it in monospace. A submission spanning several destinations says
  `+ N more destinations` — one that went to three places must not look like one that went to one.

## AN ADDRESS THE PICKER WILL NOT SUGGEST CAN NOW BE ADDED (2026-08-27, 1.0.299.0)
Client: *"I cannot add new email at all, is there no button for this."* Correct - there was none.
`GroupMembersEditor` could add somebody ONLY by clicking a candidate in the people-picker dropdown:
**no Add button, and no Enter handler.** So an address `searchTenantPeople` does not surface -
`reenelow@hotmail.com`, an external guest not yet in the directory - was **unaddable**, with nothing on
screen to say why or what to do.
- **⚠ THE PICKER'S SETTINGS LOOK LIKE THEY SHOULD COVER THIS AND DO NOT.** `searchTenantPeople` runs
  `AllowEmailAddresses: true` and `PrincipalSource: 15`, and the 2026-08-23 note on the lookup box says
  *"anything shaped like an email is offered as a candidate whether or not it belongs to anybody"*.
  That is true of the LOOKUP box; it was not true here for an external address this tenant's picker
  declines to surface. **A documented behaviour of one control is not a behaviour of the other.**
- **FIXED with an `Add` button plus Enter**, both routed through `ensureSiteUser` → `/_api/web/ensureuser`,
  which resolves the address to a canonical claim so the answer about whether an address is real comes
  from SharePoint rather than from us.
  - **⚠⚠ THE TWO CLAIMS THAT USED TO SIT HERE WERE BOTH FALSE, AND THEY PRODUCED A WRONG DIAGNOSIS
    THE NEXT DAY (corrected 2026-08-27).** They read *"the same mechanism SharePoint's own Share dialog
    uses"* and *"invites an external as a guest where the tenant allows it"*.
    - **The Share dialog calls `SP.Web.ShareObject`, NOT `ensureuser`** — `Requests.tsx`'s
      `performShare` is the proof, and it is in this same repo.
    - **`ensureuser` DOES NOT INVITE ANYBODY.** It RESOLVES an address to a principal that already
      exists (a directory user, or a guest somebody already invited). For an external address nobody
      has ever invited there is nothing to resolve, and it answers **`"could not be found"`** — which
      is exactly the error the client hit, twice.
    - **THE COST OF THE WRONG LINE: it sent the diagnosis to the site's sharing setting.** With
      `ensureuser` believed to invite, a refusal can only mean sharing is off — so that was the answer
      given. **The client disproved it in one sentence: `ckyan90@live.com` (a live.com address) was
      added by hand, so external sharing was demonstrably ON the whole time.** A guest invited through
      SharePoint's own UI resolves fine afterwards; one never invited never will.
    - **DIAGNOSTIC THAT DISTINGUISHES THEM, and it is free:** if ANY external address is already a
      member of any group on the site, sharing is on and a `could not be found` is about THAT address
      never having been invited — not about the setting.
- **⚠⚠ AN INVITE PATH WAS BUILT (1.0.301.0) AND REVERTED WITHIN THE HOUR (1.0.302.0) — ADDING
  OUTSIDERS IS NOW AGAINST CLIENT POLICY.** Client: *"client doesn't want to allow us to add
  outsiders, so we got to revert again. So just now was correct when it shows 404, but show an error
  message why it is not allowed."*
  - **SO `could not be found` FOR AN OUTSIDER IS THE CORRECT, FINAL OUTCOME — a refusal to be
    EXPLAINED, not a gap to close.** Do not re-add the invite without the client asking. `1.0.301.0`
    was never deployed anywhere.
  - **`inviteToGroup` IS DELETED, NOT PARKED**, against this codebase's usual habit (`GroupMapBuilder`,
    `inheritDefaults`). A parked EXPORTED function doing exactly the forbidden thing invites the next
    session to wire it back up, and **lint cannot flag an unused export**. Same reasoning that fully
    removed `requestCommand` and `overwrite=true`. A comment at `spGroups.ts` ~502 records the shape so
    it is recoverable in minutes: `POST {web}/_api/SP.Web.ShareObject`, `roleValue: "group:<spGroupId>"`,
    `groupId: 0`, `peoplePickerInput: JSON.stringify([{ Key: address }])`, verdict read from the BODY
    because a refusal arrives as `Status: false` behind an HTTP 200. **The `group:` role value was
    never run live** — if it is ever revived, test it before trusting it.
  - **THE REVERT WAS FOUR EDITS, NOT ONE, and three of them were MESSAGES.** Deleting the call is the
    easy part; what remains otherwise goes on advertising a capability the page does not have:
    **both** branches of `explainAddFailure` told the admin an outsider "has to be invited to the site
    once" (one naming the Add button as how), and the hint under the box said an outside address is
    "invited and put in this group in one step". **A message promising a button will do something it
    cannot is worse than no message** — the admin presses it, gets the same refusal, and reports a bug.
    That is how this arrived twice in one day.
  - **THE MESSAGE NOW STATES THE RULE AND SAYS IT IS DELIBERATE.** That last clause is load-bearing:
    without it the refusal reads as a fault and gets escalated, which is exactly what happened. It also
    **names both causes** (typo, or no account here) because SharePoint answers identically for both and
    a domain check cannot separate them either — `MySubmissions`'s `tenantDomains` counts a guest's own
    domain as internal. Naming one cause sends half the admins the wrong way.
  - **THE ADD BUTTON STAYS.** Its original 1.0.299.0 reason is untouched: an address that EXISTS but
    which the picker declines to surface — an existing guest, or a work account it fails to match — is
    still added by typing it. Only the invite fallback is gone.
  - **✅ THE MESSAGE IS VERIFIED LIVE (2026-08-27)** — the new sentence rendered for
    `minecraftx@gmail.com`, in place of the raw `SPException`.
- **⚠ AND THE SAME TEST FOUND THE DROPDOWN COVERING IT (fixed 1.0.303.0).** Client: *"The dropdown
  keeps showing even after I add already"*, with the first line of the explanation hidden behind the
  candidate list.
  - **THIRD ROUTE-MISMATCH BUG IN THIS ONE CONTROL IN ONE DAY, AND THAT IS THE POINT.**
    `onAdd`'s catch has closed the dropdown since 2026-08-23 (client, then: *"the dropdown doesnt
    clear making it difficult to see"*). The Add button was built on 2026-08-27 as a SECOND route into
    the same failure path and inherited neither that `setResults([])` nor `explainAddFailure`'s
    `ensureuser` wording. **A new route into an existing failure path silently starts from zero** —
    it compiles, it runs, and it is missing every lesson the old route had learned.
  - **FIXED BY EXTRACTING `failAdd`, NOT BY COPYING THE LINE.** A second `setResults([])` would have
    been a second copy of the rule, which is exactly the shape that produced the bug; now there is ONE
    definition of *"the add failed"* — inline reason, dropdown closed — and a fourth route cannot miss
    it. Every other `setAddError` in the file is a CLEAR, so this is the only place a failure is set.
  - The typed text is deliberately kept: a typo is fixed by editing it, not retyping it.
  - **The dropdown is absolutely positioned over the error**, which is why this is a real bug and not
    cosmetic — the same class as the people-picker being clipped by a scroll container in
    `GroupManager` and the info panels clipped by `.dms-staged-row`. **Anything absolutely positioned
    in this codebase has now bitten three different screens.**
  - **⚠ THE SHARE-REQUEST FLOW IS UNTOUCHED AND STILL REACHES EXTERNAL ADDRESSES.**
    `Requests.tsx`'s `performShare` calls `ShareObject` with a `role:` value, gated by the
    `allowExternalSharing` config row — so an approved share request can still send a DOCUMENT outside
    the organisation while this page refuses to add the same person to a group. **Both may well be
    intended** (a one-off document share is not standing membership), but it is an inconsistency the
    client has not been asked about. Raise it rather than discovering it.
- **⚠ ANSWERED THE SAME DAY: A SYSTEM ADMIN DOES NOT NEED `CRS_SITE_MEMBERS`, BUT IS ADDED TO IT
  ANYWAY.** `CRS Owners` holds **Full Control on the web**, which is strictly wider than the Read the
  entry group grants — so site entry is already covered. `addMemberWithSiteEntry` adds it regardless
  (it skips only when the TARGET is the entry group), which is redundant and harmless.
  - **⚠ BUT `onRemove` DOES NOT REMOVE IT, so the two directions are ASYMMETRIC.** Taking somebody out
    of `CRS Owners` leaves them in `CRS_SITE_MEMBERS` with Read on the site, while the toast says that
    group's access is revoked — true of the folder grants, and not of site entry. **Not a leak** (Read
    on the web is what every ordinary member has) but it is not "access removed" either. Known, not
    fixed; state it if anyone asks why a removed owner can still open the site.
- **⚠ THE ONLY ROUTE FOR AN OUTSIDER IS NOW OUTSIDE THIS APP, and that is the point:** somebody with
  tenant rights invites the address through SharePoint's own **Share** or Site permissions → Add
  members, after which it resolves and the Add button works. **Do not present this as a workaround to
  an admin** — the client's decision is that adding outsiders is not this page's job, so the message
  says the account must exist and stops there.
- **Shown ONLY when the text is address-shaped**, so the button never invites a click that cannot work.
  Picking from the dropdown stays the normal path; this is the fallback.
- **A refusal is surfaced, never swallowed**, through the existing `explainAddFailure` - which already
  translates *"The user does not exist or is not unique"* into the two things it can mean, one of which
  (a guest existing TWICE for one address) is real on this tenant.
- **⚠ `onAddTyped` MUST BE DECLARED AFTER `onAdd`** - it calls it, and `no-use-before-define` is on.
  Caught by lint, not by `tsc`.
- **✅ TESTED LIVE, and it found a second gap (1.0.300.0).** The button fired, `ensureuser` answered,
  and the admin was shown **the raw JSON**:
  `HTTP 400 {"error":{"code":"-2146232832, Microsoft.SharePoint.SPException","message":"The specified
  user reenelow@hotmail.com could not be found."}}`
  - **⚠ `ensureuser` USES A DIFFERENT WORDING FROM THE PEOPLE PICKER.** `explainAddFailure` matched
    *"does not exist or is not unique"* and this route says *"could not be found"* - so it fell straight
    through to the fallback that keeps the raw text. Same situation, two phrases. **THE LESSON: a new
    route into a translating function needs its own phrase checked, not assumed** - the whole point of
    that function is that an admin never sees an SPException code, and adding a caller quietly
    reintroduced exactly that.
  - Now translated, and it names **both** causes because they are indistinguishable from outside: the
    address belongs to nobody, or **this site does not allow new guests**
    (Site settings → Site permissions → Sharing → *"New and existing guests"*). `ensureuser` cannot
    invite an unknown external when sharing is restricted, which is the likely cause here.
- **Verified**: `tsc --noEmit` clean, `heft build` clean, full suite green. Packaged as `1.0.300.0`.

## ⚠⚠ THE HC LEVEL APPEARED INTERMITTENTLY - TWO DIFFERENT TIMING BUGS (2026-08-27, 1.0.297.0)
Client: *"the highly confidential is glitchy, even when I am in the highly confidential group the
upload form and bulk upload form sometimes shows and sometimes doesn't even after selecting unit."*
**One symptom, one screen each, two unrelated causes.** Both are timing, which is why it looked random.

**UPLOAD FORM - RE-RENDERS BURNED THE PROBE BUDGET.**
- `HC_PROBE_ATTEMPTS` was 2 and the counter was incremented **in the effect body, before `ask()`
  ran** - so **two RENDERS exhausted it whether or not either probe finished.** The effect depends on
  `levelValues`, which is a fresh ARRAY on every cascade step, so a unit selection followed by any
  other level change spent both attempts in milliseconds.
- **⚠ THE COMMENT ABOVE IT HAS CLAIMED "COUNTED AFTER THE WORK, NOT BEFORE IT" SINCE 2026-08-22 AND
  WAS WRONG.** That fix moved the CONCLUSIVE pinning into `ask()` and left the increment outside it.
  The comment then read as a guarantee for five days. **A comment describing the fix is not the fix.**
- **FIXED**: a `failed()` helper called from EVERY bail path and nowhere else, so an attempt now means
  a genuine failure; `hcInFlight` per leaf so concurrent renders reuse the running probe instead of
  starting another; skip entirely once a leaf answers `granted`; budget raised to 3.
- `hcWrite` added to the deps so a granted answer settles the effect rather than leaving it live.

**BULK UPLOAD - NO `hcReady`, SO REACT NEVER SAW THE PAIR RESOLVE.**
- `hcCtx()` calls `hcAvailable()`, which **reads a MODULE CACHE React cannot observe**. The HC pair
  resolves asynchronously at mount, so the first render answers "no HC" and **nothing re-renders when
  priming lands** - the level then appeared only if some other state change happened to force a
  repaint. `Form.tsx` has had an `hcReady` flag for exactly this since 1.0.229.0; **this screen never
  got one.**
- **FIXED**: an `hcReady` state set after `primeNames`, and `hcCtx()` READS it (`hcReady ||
  hcAvailable()`) rather than leaving it unused - an unused re-render trigger gets deleted by the next
  person, and lint flags it. It can only be true when `hcAvailable()` was true, so it cannot widen the
  answer.
- **⚠ THE RENAME MADE THIS WORSE THE SAME DAY.** `HC_DOCUMENTS_CANDIDATES` now probes
  `HC Documents` (404 on this site) before `HC Document` (200), so priming takes one more sequential
  request than it did - a wider window for exactly this race.
**BULK UPLOAD NOW HAS THE CLEARANCE PROBE TOO (1.0.298.0)**, at the client's request: *"bulk upload
also need HC dropdown, the same thing ensuring only Highly confidential group can see the dropdown."*
- **⚠ IT HAD NONE AT ALL, AND THE COMMENT EXPLAINING WHY WAS STALE.** It said the page was
  **admin-only** - true when written, and **false since 2026-08-22**, when Bulk Upload was opened to
  every uploader. So `canWriteHc: hcAvailable()` offered the Highly Confidential level to anyone who
  could open the page. Nothing was ever filed wrongly (the HC unit folder is RESOLVED, never created,
  so the write refuses), but the level was shown to people who cannot use it.
- **The part of that comment that DOES still hold** is the bit about `Documents`: an `AddListItems`
  probe against the approved side would empty this form for every PIC, because `UPL` holds Read there
  by design. The existence gate on the normal path is unchanged; this probe asks about the HC APPROVAL
  library, which is where an HC upload actually lands.

**AND THE PROBE IS NOW ONE IMPLEMENTATION: `shared/hcClearance.ts` (`useHcClearance`).**
- **⚠ THE TWO BUGS ABOVE EXISTED *BECAUSE* IT WAS INLINE IN ONE SCREEN AND ABSENT FROM THE OTHER.**
  The form gained an `hcReady` trigger on 2026-08-22; Bulk did not. The attempt counter was fixed in
  the form's copy and there was no other copy to fix. **A second copy of this logic will drift again** -
  the client's single sentence covered two unrelated causes, which is what drift looks like from
  outside. Mount the hook; do not re-inline it.
- Both screens already had an identical `permissionedLeafTerm()`, so the hook takes the leaf and the
  `hcReady` flag and owns everything else: the failure counting, the in-flight dedupe, the
  fail-closed verdict handling and the logging.
- `grep probeFolderUploadAccessByPath` now returns the hook only.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.298.0`. **Not yet tested live** - select a unit and change a below-Unit tier several times,
  then reload and repeat, on BOTH screens; and confirm an UNCLEARED uploader is offered no HC level on
  Bulk Upload, which is new behaviour there.

## ⚠ HC UPLOAD THROUGH THE BATCHED FORM HAD NEVER WORKED — FIXED AND VERIFIED (2026-08-22, 1.0.230.0)
**Verified on site the same day:** Highly Confidential + Legally Privileged through the upload form
landed in `HC Approval Document/GHO/GCA/EG/2024/Tax Return/Archive 1`, stayed *Waiting for Approval*,
and the auto-approve flow logged **no run** — correct, a form upload carries no `BulkImport` marker.

Every Highly Confidential file was refused with *"You are not cleared to file Highly Confidential
documents in this unit"* — on a unit the uploader demonstrably could write to. The clearance probe
fired, returned **200** with `Low = 4294705151` (AddListItems set), and the form refused anyway.
- **`hcContext()` read `permissionedLeafTerm()`, which reads the LIVE dropdowns — and SAVING A BATCH
  RESETS THEM.** So at Upload time the leaf was `""`, `hcWrite[""]` was undefined, and clearance read
  false. It now takes the BATCH's `leafTermId`; the live default remains only for the render-time call
  that decides which levels to offer while editing.
- **THE SAME TRAP `uploadBatches.ts` ALREADY NAMES FOR THE DESTINATION** — *the destination is snapshot
  at save, never referenced* — applied to a second field nobody re-checked. **When a form stages work
  and clears its inputs, EVERY write-time read of those inputs is a bug.** Worth grepping for others.
- **It failed CLOSED, which is why it was never reported as a leak** and also why it survived: the
  refusal looks like a permissions problem, so it gets escalated as "give me access" rather than as a
  defect. Bulk Upload was unaffected — it has one selection and never resets.
- **The probe itself was also hardened (1.0.229.0)**: attempts are counted AFTER the work rather than
  before, so a failed folder-map read no longer denies a unit for the life of the page; only a
  conclusive verdict settles it; every bail path is logged (all four were silent); and `hcAvailable()`
  is watched through an `hcReady` state flag, since it reads a module cache React cannot see change.
- **THREE WRONG THEORIES PRECEDED THE FIX** — a missing tier column, the `'Pending'` trigger clause,
  then `hcReady`. Each was killed by a cheap check: a field diff, a one-library unhide, and finally the
  Network tab showing the probe firing and succeeding. **Ask for the observation before building the
  explanation.**

## BULK UPLOAD IS AN UPLOADER TOOL, THROUGH THE APPROVAL LIBRARY (2026-08-22, 1.0.225.0)
Client: *"Bulk upload is now allowed for all uploaders to be used, client doesnt want admin to do the
job"*, and on why approval is skipped: *"the reason for bulk uploads is because it is for client to
upload old documents that is already approved before."* Spec
`2026-08-22-bulk-upload-for-uploaders-design.md`. **BUILT, and the flow EXISTS —
`CRS — Auto-approve bulk imports in Approval Document`, confirmed in the flow list 2026-08-23.**
⚠ The two lines that used to sit here and below — *"the flow does not exist yet"* and *"STILL TO DO:
build the auto-approve flow"* — were **stale, and were repeated as fact on 2026-08-23** while the flow
had been running for a day. The same failure this file's own header warns about.
- **THE CLIENT ASKED FOR A NEW HIDDEN LIBRARY, AND IT ALREADY EXISTS — it is `Approval Document`.**
  Same tree, same per-unit ACLs, same Auto-route. The only difference in the proposal was skipping
  approval, which is a flow and not a library. **"Hidden from everyone" cannot mean flat**: uploaders
  must write into it, so without per-unit ACLs every uploader would see every other unit's files —
  which is what would have made it a duplicate rather than a drop box, at the cost of a fifth and
  sixth library, ~2,400 folders and ~980 grants per segment, and two more flows.
- **REPOINTING WAS DELETING A SWAP.** Bulk Upload already resolved the APPROVAL library's unit folder
  by UniqueId and then called `toDocumentsPath` to mirror it. It now writes into the folder it had
  already resolved; the mirror survives for the clash check alone.
- **⚠ THE DUPLICATE CHECK HAD TO BE PINNED TO THE APPROVED SIDE OR IT WOULD HAVE FOLLOWED THE WRITE**
  (spotted by Clarence). The probe was against `Documents` *by construction*, because that was the
  target. Repointing would silently have repointed the check too, leaving nothing looking at where the
  file lands — and the clash would surface inside Auto-route's `Copy file`, as a document stuck
  approved-but-never-routed or an approved record overwritten. **A check does not move with the code
  that owns it; it must follow the FILE.**
  - **Both web parts now check the approved side**, and the upload form never did at all — so this
    closed a live gap there too, not only a new one here.
  - **Bulk Upload REFUSES rather than prompting**, unlike its approval-library clash: an uploader holds
    Read on `Documents` and could not overwrite anyway, and an admin must not be one click from
    silently replacing an approved record.
  - **Fails OPEN** — a 403, a throttle or a missing folder allows the upload, which is the behaviour
    before the check existed. Paths go through the OData alias form (gotcha #9), or a deep path returns
    400, reads as "no clash", and is the exact silent failure the check exists to prevent.
- **`BulkImport` (Yes/No) is the auto-approve marker**, written by Bulk Upload alone; a third flow
  approves items carrying it and the EXISTING Auto-route does the rest.
- **⚠ THE FLOW'S TRIGGER NEEDS THREE CLAUSES, AND THE FIRST DESIGN HAD ONLY TWO — corrected before
  building.** *Created* alone could never have worked: Bulk Upload uploads the file and THEN tags it,
  so at creation there is no marker, the condition is false, and **a flow that never fires leaves no
  run history**. So it is *created or modified*, which then needs a stop clause or the MERGE re-fires
  it in a loop. `{IsFolder} = false` AND `BulkImport = true` AND `{ModerationStatus} != Approved`.
  **Drop the marker clause and it approves EVERY file in the approval library on touch — abolishing
  approval site-wide, silently, on a flow that reports success every time.**
- **Degrading is safe and VISIBLE here**, unlike the submission stamp: no column ⇒ no marker ⇒ the flow
  never fires ⇒ the files WAIT IN THE APPROVAL QUEUE, where the HoU can see them. **A failure that
  leaves work in a queue is not a silent one.**
- **THE WRITE PROBE IS ON NOW, AND ITS ABSENCE WAS NEVER A BUG.** It was existence-gated only
  *because* it wrote into `Documents`, where a PIC holds Read — an `AddListItems` probe there would
  have correctly emptied the form for every uploader. **That reasoning expired with the target.**
- **`Bulk-Upload.aspx` MOVED FROM `MUST_LOCK` TO `MUST_NOT_LOCK`, and the move is required.** The
  derived-page pass now grants it; if the lockdown pass still claimed it the two would fight on every
  run, one granting and the other stripping, for ever. Pinned in both test files.
- **⚠ `ensureColumn` HAD THE SAME CAPPED READ — third instance in two days** (585 groups, then the
  submission stamp). `/fields?$top=500` searched for the column: on a big library it would conclude a
  column is absent and try to CREATE it, which SharePoint refuses as a duplicate internal name — so
  reconciliation would report a column it cannot create on a library that already has it. Now
  `$filter`ed. **`libraryHasColumns` in `shared/optionalColumns.ts` is the ONE implementation** both
  upload web parts use; two copies deciding whether a field is safe to write is how one starts
  stripping metadata the other preserves.
- **⚠ THE REPOINT SENT HC UPLOADS INTO THE NORMAL APPROVAL LIBRARY — FOUND ON SITE THE SAME DAY,
  FIXED 1.0.227.0.** `stagingSru` comes from the Folder Map, which only ever describes the NORMAL
  library. The old code reached the HC pair solely through `toDocumentsPath`'s segment swap, so
  **deleting the swap deleted HC routing with it** — and nothing in the diff looked like HC.
  - **The dangerous outcome was the silent one.** A Highly Confidential document was filed where every
    uploader in the unit can see it while pending (draft security has been off since 2026-08-19).
  - **The 400 was the LUCKY part.** The file landed in the normal library while
    `validateUpdateListItem` addressed the HC one, and **item ids are per-LIST** — so tagging failed
    loudly instead of writing this document's metadata onto a different one. The 1.0.170.0 defect in
    mirror image.
  - **The HC unit folder is RESOLVED, NEVER CREATED** — one created here inherits the HC library
    ROOT's ACL instead of the unit's, which is exactly how an HC document becomes readable by the
    people the unit ACL excludes. Absent ⇒ refuse and name reconciliation.
  - `writeSideSegment()` returns **undefined** rather than falling back when the HC pair is
    unresolved. `mirrorPath` then yields null and the run refuses. A fallback files a secret in the
    open library; `hcRouting.ts` fails closed for this reason and so does this.
  - **THE GENERAL RULE: when a helper is deleted, ask what ELSE it was doing.** `toDocumentsPath` read
    as "work out the Documents path"; it was also the only thing that knew about HC.
- **⚠⚠ `Hidden: true` ON A FIELD REMOVES IT FROM THE POWER AUTOMATE TRIGGER PAYLOAD — and nothing
  reports it (2026-08-22).** `BulkImport` was hidden on all four libraries to stop an uploader ticking
  it by hand. The WRITE still worked (`validateUpdateListItem` ignores `Hidden`, and a REST read showed
  `BulkImport = true`), but `triggerOutputs()?['body/BulkImport']` came back **null**, so the
  auto-approve condition could never match and **both flows stopped firing**. Silent at both ends: no
  error on the write, and no run to inspect on the flow.
  - **Proven by unhiding ONE library.** With `Approval Document` visible and `HC Approval Document`
    still hidden, the normal bulk import fired and the HC one did not. Three wrong theories preceded
    that — a missing tier column, then the `'Pending'` clause, then "both libraries are broken" — and
    each was settled by a cheap check rather than by reasoning. **Change one thing, test, repeat.**
  - **Keep it VISIBLE on both approval libraries. Hidden is fine on `Documents` / `HC Documents`**,
    which no flow watches.
  - **So the hand-tick exposure cannot be closed with `Hidden`.** The narrower `ShowInEditForm: false`
    is the next candidate and must be verified the same way: set it on ONE library, bulk-upload, confirm
    the flow still fires. If that also breaks the trigger, accept the residual risk — a working
    auto-approve beats a deterrent against a deliberate tick.
- **⚠ A MIXED RUN LEFT THE UPLOADED FILES IN THE LIST, SO PRESSING UPLOAD AGAIN RE-SENT THEM**
  (found on site 2026-08-22, fixed 1.0.228.0). The clear was `if (!anyProblem && okCount > 0)` —
  all-or-nothing — so ONE skipped file kept the whole selection on screen including everything that had
  already gone. A second press raises a replace prompt per file, and someone clicking through
  overwrites a document in the approval library.
  **`uploadBatches.ts` already states the rule the batched form runs on — SUCCESS REMOVES, FAILURE
  STAYS — and this screen never got it.** It is what makes a second press a safe RETRY: an uploaded
  file cannot be sent twice, and what is left on screen is exactly what still needs doing.
  Matched by INDEX, never by name: this screen keeps original filenames, so two identical names in one
  selection would collapse to one match. A missing result counts as not uploaded and stays.
- **⚠ THE TAGGING FAILURE NOW CARRIES THE RESPONSE BODY AND THE LIBRARY.** It said only
  `tagging failed (HTTP 400)`, which is what turned the above into an hour of guessing. 404 is the
  wrong title, 403 is permissions on that list, 400 is a malformed payload or an unknown field name —
  opposite fixes, and SharePoint names the cause in the body. Third time this lesson has been learned
  here (gotcha #9, then 1.0.170.0).
- **⚠ AUTO-ROUTE'S APPROVAL EMAIL MUST BE SUPPRESSED FOR BULK IMPORTS** (found on the first live test).
  It reads *"Great news! Your document has been approved"* — true for a normal upload, false for a bulk
  import where **no human approved anything**, and a 50-file import sends 50 of them. Its final email
  action is now wrapped in a Condition on `body('Get_item')?['BulkImport'] is not equal to true`.
  **The operator was set the wrong way round on the first attempt**, which emails only bulk imports and
  **silences every normal upload** — a failure nobody reports for weeks. Read the condition back as a
  sentence before saving. **The REJECTION email is deliberately untouched: a rejection is always a
  human act**, and a bulk file only reaches that branch when auto-approve did not fire.
- **⚠ A STALE LIST VIEW SHOWED THE ROUTED FILE AS HAVING NO METADATA, AND IT HAD ALL OF IT.** Reported
  as a bug on the first live run; the **Details pane** showed Business Segment, Department, Unit, all
  three Tids, Document Type, Year and Confidentiality present. The list VIEW caches; the Details pane
  reads live. Before diagnosing a metadata loss on a just-routed file, open the pane — the copy takes
  18 seconds and the view renders long before the stamp finishes.
- **THE FLOWS EXIST.** `CRS — Auto-approve bulk imports in Approval Document` and `HC auto-approve`
  are both in the flow list, alongside `HC Auto Route` — so the HC vertical has its mover too, and the
  old claim here that *"HC bulk upload cannot work — the HC Auto-route flow has never been built"* was
  **contradicted by this same file** two sections up, which records that flow as built and verified end
  to end on 2026-08-19. Runbook:
  `docs/superpowers/specs/2026-08-22-bulk-import-auto-approve-flow-runbook.md`.
  - **Built as `clarence@trinergydigital.com` here. On SDG's tenant, sign in as the SERVICE ACCOUNT
    before the first action** — the connection is baked in for life.
  - **STILL OPEN, and it is one field:** `BulkImport` must be **VISIBLE** on BOTH approval libraries.
    `Hidden: true` removes a field from the trigger payload, so the flow never fires and leaves **no
    run history** to inspect. On 2026-08-22 the normal library was unhidden and its bulk import fired;
    **whether `HC Approval Document` was ever unhidden is NOT recorded** — check the field before
    concluding HC bulk import is broken.
  - Reconciliation creates `BulkImport` and grants the page, so re-run it on any site where bulk
    upload has never been used.
- **IT LEFT THE CRS SETTINGS LANDING PAGE, AND ITS BACK BAND WENT WITH IT (1.0.226.0).** The card was
  there because Bulk Upload was an ADMIN tool. That page is itself admin-only (a page called "CRS
  Settings" matches the `setting` keyword), so keeping the card would point administrators at a screen
  that is no longer theirs **while staying invisible to every person who now uses it** — the worst of
  both. The `withBackToSettings` band came off `BulkUploadWebPart` for the mirror reason: it would
  offer every uploader a link that answers **AccessDenied**, which reads as a broken system rather
  than a page not meant for them. Now matching the upload form, the approval queue and My Submissions,
  none of which carry the band. **Put the link where uploaders already go — the home page, beside CRS
  Search.** The `bulk` icon is kept; a future admin-facing import card should not be a wider change
  than adding a card back. Pinned by test.
- **Tell the client:** nothing enforces "historical only"; `Created By` becomes whoever bulk-uploaded,
  not the original author (the real date is `DocumentDate`); and it costs **two flow runs per file**,
  bounded at 100 per press by the existing `MAX_FILES = 50`.

## SUBMISSION GROUPING IS VERIFIED ON SITE (2026-08-22, 1.0.224.0)
One press of Upload spanning two departments produced ONE row — `SUB-20260822-8DJN`, two batches
(`BAT-…-GRWZ`, `BAT-…-HWJ9`), each opening as the read-only upload form. The whole chain works:
stamp on write, group on read, drill down to the details.
- **The diagnosis that got there is the reusable part.** `/items?$select=…,SubmissionId,BatchId`
  returned `m:null="true"` — **the columns EXIST and the value was never written.** A 400 would have
  meant an unreconciled library and a completely different fix. Read the ITEM, never the banner:
  `refsMissing` only fires when the READ 400s, so it says nothing about whether a given file was
  stamped when it was written.
- **⚠ THE STATUS-TAB EMPTY STATES FIRED ON THE SUBMISSIONS TAB** — *"Nothing is submissions right
  now"* printed underneath a table listing seven of them, on every screenshot of this feature for
  three revisions. `filterByTab` has no case for `Submissions` (it renders its own view), so `shown`
  is empty and the three shared empty states, gated only on `tab !== "Requests"`, all applied. Now
  `isStatusTab`, a named predicate rather than a second literal, so the next tab added cannot walk
  back into it.
- **Files uploaded before the stamp landed can never be grouped** — nothing retro-stamps them, and
  they stay as folder-and-date guesses for ever. Expect this on every site at cutover.

## A BATCH OPENS AS THE UPLOAD FORM, READ ONLY (2026-08-22, 1.0.222.0)
Client, correcting the first build: *"When I say upload Form I literally meant a copy of the upload
form with the details in it but view based only."* The first attempt made level 3 another file LIST
that had to be clicked through one file at a time — which is the flat list the whole change was
meant to replace.
- **THE FORM'S OWN SHAPE IS THE SPEC.** The destination is chosen ONCE and every file's details are
  filled in beneath it, so the read-only view renders exactly that: one **Document folder
  information** card (the permissioned tiers, then Document Type and Year) and then one expanded card
  per file. Nothing to click through.
- **`FIXED_FIELDS` IS NOW `BATCH_FIXED_FIELDS.concat(FILE_FIXED_FIELDS)`, derived and never re-listed.**
  Year and Document Type sit on the BATCH because they are the built-in below-Unit folder tiers — the
  uploader picks them once per destination. A second literal list is how the batch view and the
  single-document panel would come to disagree about what a document carries; a test pins the
  partition as exact, nothing lost and nothing shown twice.
- **⚠ THIS IS THE ONE VIEW THAT COSTS REQUESTS — one `FieldValuesAsText` per file.** It is a per-ITEM
  endpoint, which is why the flat list shows no metadata at all. Bounded (a batch is a handful of
  files) and deliberate (they opened it). `Promise.allSettled` is unavailable here (gotcha #3), so
  each read carries its own catch inside `Promise.all`: **one unreadable file must never blank the
  batch.**
- **The destination is read from whichever file COULD be read, not from the first.** A batch is one
  folder so any of them answers — but the first may be the one whose read failed, and falling back to
  it would show an empty destination for a batch that has one.
- **`{}` for a file says "could not be read, or nothing was recorded"**, never an empty card. Empty ≠
  unknown, again — the two are indistinguishable from the response, so the panel says both.
- **Level 1 gained a `Batches` column** (client's own sketch: Submission · Batches · Files ·
  Uploaded), and **`Location` always renders** on the destination card because it comes from the file
  PATH rather than the metadata — the one row that stays right when the per-item read failed.
- **⚠ THE TAB READ `Submissions (0)` OVER A TABLE LISTING FOURTEEN.** `tabCounts` spread `counts`,
  which covers the status tabs only. A page contradicting itself reads as a bug; it now counts from
  the same `groupSubmissions` the tab renders, so the tally and the list cannot disagree.
- **Open file** on each card still reaches the full panel — preview, rejection reason, and the
  deletion and share buttons. This view is read-only by design; that one is where an uploader ACTS.

## REQUEST LISTS REFRESH THEMSELVES NOW (2026-08-21, 1.0.219.0)
Client: *"is it possible to make the request notification react live? I always have to refresh the
page which is troublesome."* Rule and hook in `shared/liveRefresh.ts` (pure predicate, 5 tests);
mounted by `MySubmissions` and `Requests`.
- **⚠ THERE IS NO PUSH CHANNEL, and do not go looking for one.** SPFx runs in the browser with no
  server component, and SharePoint webhooks notify a SERVICE, not a page. Sub-second updates would
  need a service polling server-side and pushing over websockets — out of scope and out of proportion
  for a queue that changes a few times a day.
- **THE FOCUS LISTENER IS THE ONE THAT ACTUALLY SOLVES IT.** The reported workflow is switching to
  the approver's account, deciding, and switching back — so refreshing on `focus` and
  `visibilitychange` makes the list current before the reader has looked at it, with **no timer cost
  while the tab is idle**. The 45-second poll is only the backstop for someone sitting on the page.
- **BLOCKED while a dialog is open or a write is in flight.** A refresh that lands mid-decision can
  replace the row being acted on. Skipping is free — the next focus or tick picks it up, and every
  write path already reloads on its own.
- **My Submissions reloads `loadPolicy` ONLY** — the small requests list. Never the two document
  libraries: those are the expensive reads and they do not change while someone watches this page.
- **`reload` and `blocked` are held in REFS, and the effect has empty deps.** Both are redefined on
  every render, so depending on them would tear down and rebuild the interval several times a second
  — which in practice means the timer never fires at all.
- Errors are swallowed: a background refresh must never put an error in front of someone who did not
  ask for anything.

## A CANCELLED REQUEST WAS STILL CARRIED OUT (2026-08-21, 1.0.220.0)
Found on site within an hour of guarding the other direction. The approver's queue is read at MOUNT,
the requester withdraws in their own session — so a page left open across a cancellation still lists
the row, and pressing Approve ran `performShare` **first**. A share was executed and the recipient
emailed for a request whose row read `Cancelled`.
- **THE RE-READ MUST COME BEFORE THE ACTION, NOT BEFORE THE STATUS WRITE.** `decide()` already
  recorded `Failed` when an action failed, and that is not the same guarantee: the action is the part
  that cannot be taken back. Granting access or recycling a document is not undone by a status field.
- **THE PAIR IS THE LESSON.** Guarding Cancel against a stale approval and leaving Approve unguarded
  against a stale cancellation fixed one direction of the same race and shipped the other. **When a
  row can be written from two sessions, guard EVERY writer, not the one that was reported.**
- Fails CLOSED — an unreadable check decides nothing and says so — and the refusal distinguishes
  *"the uploader withdrew it"* from *"someone else decided it first"*, because those mean different
  things to an approver looking at an empty queue.
- Fourth place this codebase has needed **the state that matters is the one at WRITE time**: the
  stale-chain guard, the upload-pause re-check, Cancel, and now Approve.

## CANCEL COULD OVERWRITE AN APPROVED DECISION (2026-08-21, 1.0.217.0)
Found on site by the client. A requester's My Submissions page reads its rows at MOUNT, and the
approver decides in their own session — so a page left open across an approval **still shows Cancel**,
`canCancel` still passes on the stale `Pending`, and the MERGE (`IF-MATCH: "*"`, no concurrency check)
wrote **`Cancelled` over an `Approved` decision on a file already in the recycle bin**. The audit row
then read *"Withdrawn by the requester before any decision was made"* — false, on the one list whose
whole purpose is being a record.
- **THE ROW IS RE-READ IMMEDIATELY BEFORE THE WRITE**, and the cancel is refused if its CURRENT status
  is not `Pending`. Same shape as the stale-chain guard (gotcha 10b) and the upload-pause re-check:
  **the state that matters is the one at WRITE time, never at render time.** Third place this pattern
  has been needed.
- **FAILS CLOSED, against this codebase's usual direction.** An unreadable check refuses and says so.
  Everywhere else a failed read costs a form for a minute; here it would overwrite a decision that has
  already been carried out, and the user can simply try again.
- **The refusal NAMES the decider and the outcome.** "Too late" without saying what happened leaves the
  requester unsure whether their document still exists.
- **THE REFUSAL NEEDS ITS OWN BANNER (1.0.218.0).** It was first routed through `setLoadError`, to
  avoid a second message channel — which rendered it as *"Could not read your submissions … please
  try again rather than uploading again"* over a page that had loaded perfectly and a request that
  had been decided correctly. **A normal outcome must not borrow the failure channel**: a refused
  cancel is the system working, and telling an uploader their submissions could not be read invites
  the one thing the message warns against. Now `requestNotice`, amber, cleared on each attempt.
- Not fixed: the button is still *shown* on a stale page. That is cosmetic once the write is guarded,
  and polling to hide it would cost a request per row per page.

## MY SUBMISSIONS NOW TAGS HC ROWS (2026-08-21, 1.0.217.0)
The same document name can exist in BOTH the normal and the HC library — one upload classified Highly
Confidential, one not — and the page showed them **identically**: same name, same tier path, same
status badge. An uploader raising a deletion was choosing blind; the only way to learn which one went
was the recycle bin's *Original location*. Confirmed live: `TES - TES - TES - 20-08-26.pdf` existed in
both, and the approved deletion took the HC copy while an identically named file remained.
- `isHcRow` in `shared/mySubmissions.ts` (pure, 6 tests), rendered as an amber **HC** tag on the row
  **and repeated in the request dialog** — by the dialog the uploader has committed to a file, and that
  is the last moment they can notice.
- **The ACTION was never ambiguous** — requests are keyed on `uniqueId`. This is about the person.
- **Returns FALSE when the HC pair is unresolved, rather than guessing from an `HC` name prefix.** A
  wrong `true` brands an ordinary document as confidential, and the unresolved case cannot hide a
  needed tag: an uncleared uploader cannot see the HC libraries, so they have no HC rows.

## RECONCILIATION NOW CHECKS SITE PAGES IS READABLE (2026-08-21, 1.0.217.0)
The mirror of the library assertion it sits beside. That one checks a CRS library does **not** inherit,
because inheriting is a **leak**; this checks Site Pages **is** readable by the site-entry group,
because not being readable is a **lockout** — and the lockout is the one nothing reported.
- **REPORTS, NEVER REPAIRS.** Breaking a library's inheritance only ever REMOVES access; granting Read
  here would ADD it, on the one list that also holds the ten admin pages the same run deliberately
  locks. The message names the exact fix **and warns against the "share everything in this folder"
  tick**, which would unlock those admin pages.
- **Three states, and it says which:** inheriting ⇒ fine; unique and the entry group is listed ⇒ fine;
  unique and absent ⇒ `⚠ NOBODY CAN OPEN THE SITE HOME PAGE`. Every read **fails OPEN** and reports
  "NOT checked" — a false alarm here would train an admin to ignore the line that matters.
- The entry group's principal id is **carried from the site-entry block**, never re-resolved: two
  answers to "which group is the entry group" is how the two checks would come to disagree.

## ⚠ THE SITE PAGES LIST HAD UNIQUE PERMISSIONS AND NOBODY COULD OPEN THE HOME PAGE (2026-08-21)
Both test accounts got `AccessDenied … Type=item` on `CollabHome.aspx` — **the site was unusable for
every non-admin** — while the page itself inherited, was published, and `CRS_SITE_MEMBERS` held Read
on the WEB. The **Site Pages LIST** had `HasUniqueRoleAssignments: true` with only `CRS Owners` on it,
so every page that INHERITS fell through to a list ACL granting nobody.
- **THE TELL IS THE ASYMMETRY: pages with their own ACL worked, inheriting pages did not.**
  `Request.aspx` (unique, derived grants) opened fine; `Home.aspx` and `CollabHome.aspx` (inheriting)
  were denied. When some pages work and others don't, look at the LIST, not the items.
- **Our code did not do it** — verified by grepping every `breakroleinheritance` call: two are
  library-scope (`listBase` from Library-scope rows, `entryListBase` over `reconLibs()`), the rest are
  folder or `items(N)`. Nothing targets Site Pages as a list. It was a manual change at some point.
- **✅ FIXED AND VERIFIED STILL IN PLACE 2026-08-27** — `/sites/ClarenceDMSTesting` Site Pages reads
  `HasUniqueRoleAssignments: true` with principal **1581 (`CRS_SITE_MEMBERS`) → Read**. SDG's
  `/sites/CRS` Site Pages **INHERITS** and needs nothing. Both are correct states; do not "fix" either.
  - **✅ RESTORING LIST INHERITANCE IS SAFE — VERIFIED 2026-08-27, AND THIS FILE PREVIOUSLY SAID
    OTHERWISE.** ClarenceDMSTesting's Site Pages was reset to inherit ("Delete unique permissions"),
    and **every page item kept its own ACL**: 14 pages read `HasUniqueRoleAssignments: true` —
    including all nine admin pages — while only `Home.aspx` and `CollabHome.aspx` inherit, which is
    exactly the shape that needs the list-level Read. `ResetRoleInheritance` acts on the object it is
    invoked on and does **not** cascade to items.
    - **The old caution said a reset "risks" the item ACLs. That was never verified, and it is wrong.**
      It was repeated three times in one session as though it were established, which is the failure
      mode the header of this file warns about — a cautious note being quoted as a finding.
    - **Both fixes work.** The GRANT is the smaller change when a list is already unique and working;
      INHERITING is the tidier end state, because it tracks the web instead of needing a hand-kept
      grant, and it is what SDG has. Either is correct; there is no per-site special case any more.
    - **Reconciliation re-asserts the admin-page locks every run regardless**, so even a genuine loss
      of an item ACL is repaired by a run rather than by hand.
  - **⚠ DO NOT CONCLUDE THE GROUP IS ABSENT FROM A TRUNCATED `roleassignments` DUMP.** The feed is
    ascending by principal id and `CRS_SITE_MEMBERS` is **1581** here — higher than every unit group,
    so it sorts LAST and is the first thing a truncated response loses. Nearly called it missing on
    2026-08-27 for that reason. Ask for it directly:
    `roleassignments/getbyprincipalid(1581)?$expand=RoleDefinitionBindings` — 404 means absent.
  - **⚠ Almost every entry on that list is `Limited Access` (role def `1073741825`) and grants
    NOTHING.** It is SharePoint's automatic entry for a principal granted on a child ITEM, which on
    this list means every group with a page grant. Only a real `Read` binding lets anyone open an
    inheriting page.
- **FIX: grant `CRS_SITE_MEMBERS` → Read on the Site Pages LIST.** Do NOT "Delete unique permissions"
  — 14 page items carry their own ACLs and resetting the list scope risks them.
  - **⚠ UNTICK "Share everything in this folder, even items with unique permissions."** It is ON by
    default in that dialog and would push Read onto all **10 locked admin pages**, handing every
    uploader the admin tools — then the next reconciliation strips them again, producing a hundred
    confusing removal lines. Caught before it was clicked.
- **SIXTH INSTANCE OF THE STRUCTURAL GAP.** Reconciliation asserts every run that the four CRS
  libraries do NOT inherit (an inheriting library is a leak). **Nothing asserts that Site Pages IS
  readable by the site-entry group** — where the failure is a LOCKOUT, equally silent. Worth building:
  one read, reported like the `hcConfidentialityLevel` assertion.
- **`Home.aspx` is a stray.** Created by the System Account at site creation, never published, linked
  from nothing — and it produces the same `AccessDenied`, which sent this diagnosis down a false path
  for half an hour. **The real home page is `CollabHome.aspx`.** Delete `Home.aspx` on any CRS site.

## ⚠ RECONCILIATION ON A STALE APP VERSION STRIPS GRANTS THE CURRENT BUILD WOULD KEEP (2026-08-21)
The site was running an older app than the catalog held, so the page pass asserted the **previous**
version's policy: every `_HOD` group was **removed** from `request.aspx` with the reason *"holds no UPL
or APR role"* — the pre-1.0.212.0 role list. The run reported success.
- **THE REASON LINE IS THE TELL, AND IT IS DERIVED** (`FolderManager.tsx` ~3036 —
  `holds no ${derivedRolesForPage(file).join(" or ")} role`). So the message NAMES the policy the
  running bundle holds. If it lists roles the current source does not, the site is behind. Compare it
  with the source before touching anything else.
- **A ONE-LINE CHECK SETTLES IT:** run `derivedRolesForPage` for that page in a scratch test. Source
  said `["APR","DEPTVIEW"]`, site said `["UPL","APR"]` — code cannot go backwards, so it is a
  deployment state problem, not a defect. That took two minutes and replaced a wrong theory.
- **⚠ NEW CONSEQUENCE OF DERIVED PAGE ACCESS (1.0.171.0).** Before it, running an old bundle cost a
  stale form. Now **reconciliation ASSERTS page ACLs**, so an old bundle actively REMOVES grants — and
  `groupsToRemove` is deliberately full-assertion at page scope, which is what makes it destructive
  here. Recoverable in one run, but every affected group loses the page meanwhile.
- **CHECK THE INSTALLED VERSION IN SITE CONTENTS BEFORE EVERY RUN, NEVER THE CATALOG.** The standing
  rule at the top of this file, learned again. `Site Contents → the app → ⋯ → Details` shows *"There is
  a new version of this app"* when it is behind; the version NUMBER on that classic page is unreliable
  (it read `1.0.148.0` on a site demonstrably running ≥1.0.200.0, since CRS Requests worked), but the
  new-version sentence is not.
- **A stale TAB does the same thing** and is not hypothetical — reconciliation runs in the browser, so
  a page loaded before a deploy holds the old bundle. Close every CRS tab after updating the app.
  - **⚠ AND IT ALSO PRESENTS AS A DEAD WEB PART SAYING `Something went wrong` / `ERROR: [object
    Object]` (seen 2026-08-26 on Group Management).** That panel names nothing, so it reads as a code
    defect and invites a hunt for a React race. **The REAL error is only in the browser console:**
    *"Refused to execute script from `.../ClientSideAssets/.../user-access-web-parts_<hash>.js`
    because its MIME type ('text/html') is not executable"*, followed by *"Could not load … in
    require. Script error for `<web part guid>_0.0.1`"*.
  - **The mechanism is the CONTENT HASH in the bundle filename.** Every build emits a new one and a
    deploy replaces the old file, so a page still open in the browser goes on requesting the OLD hash,
    gets an HTML 404 back, and the MIME check rejects it. The component never instantiates.
  - **SOFT NAVIGATION is what makes it intermittent.** SharePoint swaps modern pages without a full
    reload (`_interceptAnchorClick` → `_onIntercept` → `push` in the stack), so one tab opened before a
    deploy keeps handing out stale references all session — failing on some navigations and not
    others. A hard refresh re-fetches the page, which then points at the current hash.
  - **Nothing to fix in code.** Before debugging a web part showing that panel, check the console for a
    MIME-type refusal: if it is there, the component never ran at all and the bug is in the tab.
- **The 28 removals also blocked the orphan prune** (`Prune skipped — run had 28 error(s)`), so a
  polluted run silently skips its cleanup too. That guard behaved correctly.

## SHARE RECIPIENTS GET A PEOPLE PICKER, AND IT MUST NOT REPLACE THE TEXT BOX (2026-08-21, 1.0.215.0)
Client, after the first successful share test: *"ensure that the email input is like the Group
management where it will pre show the email."* A typo in that box is **silent** — the share goes to
nobody, or to the wrong address, and the requester never finds out.
- **⚠ A PURE PICKER WOULD BREAK THE MAIN CASE.** The point of a share request is usually somebody who
  is **not on the site yet** — the 2026-08-21 test shared with an address that had no access at all,
  which is what made it meaningful. So the picker **APPENDS** to the free-text box; the box stays.
- Reuses `searchTenantPeople` (`shared/spGroups.ts`), already configured `AllowEmailAddresses: true` /
  `AllowOnlyEmailAddresses: false` — it resolves known people and still accepts an unknown address.
- **Only results carrying an email are offered.** `SP.Web.ShareObject` shares with an ADDRESS, so a
  result without one is a row that cannot be acted on.
- **"Nobody matches" is a HINT, not an error**, and points at the box below. For this workflow no match
  is the normal outcome, and an error-styled message would read as a refusal.
- Debounced 300 ms, 3-character minimum, and a **failed search blocks nothing** — the address can
  always be typed.
- Hook declared **above** the detail view's early return, or it would not run on every render.

## DELETION AND SHARE ARE VERIFIED END TO END ON SITE (2026-08-21)
Both flows, with three real accounts (admin, a PIC guest, a HoD guest) — the first live proof of the
2026-08-15 design.
- **Deletion:** PIC raised on an approved document → HoD approved → file recycled → My Submissions
  went 17→16 / Approved 13→12 → request read `Approved by <the HoD>`. Only the requested file moved.
- **Share:** PIC raised → HoD approved → SharePoint emailed the recipient *"<the HoD> invited you to
  view a file"* — **named after the approver, because the action runs in their session with their own
  `CRS Share`**, not as an admin or a service account.
- **THE RECIPIENT GOT EXACTLY ONE FILE.** Signed in as the shared-with account: the file opened in
  `Documents/GHO/GF/TAX/2024/Term Sheet/Archive 2`; the parent `Term Sheet` folder answered *"Unknown
  render failure"*; the `Documents` root read *"This folder is empty"*. Security trimming intact
  around a per-file grant.
- **The HoD scope proved itself unprompted:** a pending-stage deletion existed throughout and never
  appeared in the HoD's queue, while an approved-stage one did.
- ⚠ **A HEAD OF DEPARTMENT CANNOT OPEN My Submissions and cannot upload** — `hod` holds no `UPL`, and
  the page is `["UPL"]`. Correct by the persona model, and worth telling the client: a real HoD who
  needs to file documents needs a `pic` group as well.
- ⚠ **A GUEST'S OWN DOMAIN COUNTS AS INTERNAL.** `MySubmissions.tsx` always pushes the signed-in
  user's domain into `tenantDomains`, so a gmail guest sharing to a gmail address never meets the
  `allowExternalSharing` gate. Right for `@sdguthrie.com` staff; **open for contractors onboarded as
  guests.** Known, not fixed.

## THE SITE-GROUP READER WAS CAPPED AT 500 ON A 585-GROUP SITE (2026-08-21, 1.0.214.0)
`fetchAllSiteGroups` read `$top=500` (`shared/spGroups.ts`). ClarenceDMSTesting holds **585** groups, so
the read was silently truncated and **Group Management reported `MHO_SEGVIEW` as not existing while it
sat at id 396** — the admin was one step from creating a duplicate group with the same name.
- **FOUR SCREENS READ THROUGH IT** — Group Management, Folder Access, Site Access, Page Access — so one
  cap makes live groups invisible on all four at once, **with no error anywhere**. A truncated read is
  indistinguishable from an absent group.
- **Reconciliation's own deleted-group check already avoided this reader for exactly this reason**
  (`FolderManager.tsx` ~3352, 1.0.192.0) and said so in a comment. The lesson was written down and the
  shared helper was left as it was: **fixing the caller that bit you is not fixing the defect.**
- **A provisioned segment is ~324 groups**, so 500 is under two segments — this site crossed the line
  the day MHO was provisioned, months before anyone noticed.
- `BulkGroupProvisioner` reads `$top=1000` independently, which is why bulk runs kept reporting
  `already exists` correctly while the screens did not. That divergence is what made the symptom look
  like a Group Management bug rather than a shared-reader one.

## A MODERN PAGE IN DRAFT DENIES READ-ONLY USERS, AND EVERY PERMISSION CHECK PASSES (2026-08-21)
Cost six REST round-trips before anyone looked at the page. A Head of Department was refused
`Request.aspx` with `AccessDenied.aspx?…&Type=item&listItemId=24` while, verified from the server:
the page's ACL granted his group `Read`, his session's `currentuser/groups` listed that group, and the
item id in the denial matched the item we had read the ACL of. **All three correct, all three
irrelevant** — the page had never been published.
- **The cheap check is the `Draft` badge on the page itself**, plus *"The comments section will be
  displayed after the page is published."* Look BEFORE reaching for the API.
- Same shape as the moderation trap on `Documents` (§ per-uploader isolation): a pending item is
  invisible to readers, every permission reads correct, and it presents as a permissions bug.
- **A newly authored page is the normal case**, not an exotic one — every page the runbook creates
  starts as a draft, so this will happen again on SDG's tenant. Publish every CRS page after placing
  its web part.

## ⚠ THE `CRS Requests` LIST NEEDS A WRITE GRANT OR NO REQUEST CAN BE RAISED (2026-08-21)
Found on the first live test: raising a deletion returned **HTTP 403 `UnauthorizedAccessException`**.
The list is created by `provision` and **inherits site permissions**, where `CRS_SITE_MEMBERS` holds
**Read** — and Read cannot add items. **The entire deletion/share feature is inert until this is fixed
by hand on each site.**
- **Custom permission level `CRS Request`** — a copy of Contribute with **Delete Items** and **Delete
  Versions** unticked — granted to `CRS_SITE_MEMBERS` on that list. Add/Edit/View/Open/View Versions
  only.
  - **Delete Versions matters as much as Delete Items**: leaving it on protects the row while leaving
    its history erasable, which is the half that matters in a dispute.
  - **Editing others' rows is still possible** and cannot be prevented: item-level write security
    ("edit only own") exempts only **Manage Lists** holders, which approvers do not hold — so scoping
    that way would stop an approver recording a decision. Same dead end as the `Documents` isolation.
  - **THE EXPOSURE IS REQUEST METADATA, NEVER DOCUMENTS.** The deletion and the share execute only when
    an approver presses Approve **in their own session with their own permissions**; setting
    `Status = Approved` by hand in the list changes a text field and nothing else. Say this plainly —
    it is what makes the coarse grant acceptable.
- **Scoping to the `_UPLOADER`/`_APPROVER`/`_HOD` groups instead would be tighter and is ~450 grants**,
  and — the stronger objection — **nothing would keep it correct**: a unit added later gets groups,
  folders and page access, and no request access, failing silently months on. **Fifth instance of the
  structural gap**; the right fix is a reconciliation pass that ASSERTS this list's ACL from the Group
  Map, not manual grants. Not built.
- **The list read is separate and still unrestricted** — anyone who can open the site can read every
  request's file name, reason, recipients and decision note.

## A SYSTEM ADMIN NOW DECIDES EVERY REQUEST — AND THE PAGE HAD NO ADMIN CHECK AT ALL (2026-08-27, 1.0.304.0)
Client: *"I notice the system admin don't show the full CRS Request"*, then, offered the choice,
*"yes obviouly 1"* — the option where an admin sees and decides everything. Rule is
`ViewerScope.systemAdmin` in `shared/requests.ts` (pure, 9 new tests, 82 total).
- **THE CAUSE: SCOPE IS BUILT PURELY FROM GROUP MAP ROWS, AND AN ADMIN HOLDS NONE BY DESIGN.**
  `aprUnits` comes from `APR`/`APRHC` rows and `hodUnits` from `DEPTVIEW` rows; `CRS Owners` grants
  through **Full Control**, not through a persona. So both arrays were empty, `isVisibleTo` fell
  through to its requester branch, and an owner saw only the four requests they had raised themselves.
- **⚠ FIFTH SCREEN IN THE `isSystemAdmin` GAP, and the worst kind: it checked NEITHER.** Four screens
  checked `IsSiteAdmin` alone until `isSystemAdmin` was extracted earlier the same day so Owners
  membership counted too — this one had no admin check of any sort, so **a grep for the old check
  could never have found it.** When a rule about "who is an administrator" changes, the screens that
  never asked are invisible to the search.
- **⚠ THE BANNER WAS WORSE THAN THE GAP, and fixing the scope alone would have left the page
  contradicting itself.** It told every owner that *"an administrator adds an `APR` or `Head of
  Department` mapping for your group on the Folder Access page"* — **actively wrong advice**: an admin
  has no unit group, and mapping them would hand them a persona they should not hold and pollute the
  list reconciliation reads. It is now suppressed for an admin (`decidesSomething` includes them), as
  is the HoD-scope-failed warning, which would report a limit they do not have.
- **⚠ WHY THIS IS SAFE HERE WHEN IT IS NOT FOR A HEAD OF DEPARTMENT — the distinction is the whole
  argument.** Approving EXECUTES in the decider's own session with their own permissions. `hod` is
  barred from pending-stage requests because it holds nothing in either approval library, so its
  Approve would fail **after** telling the requester yes. An owner holds Full Control, so the recycle
  or the share genuinely succeeds. **Do not read this as licence to widen `hod`.**
- **The operational need is real, not theoretical:** `GHO_GCA_EG_APPROVER` sat **EMPTY** on 2026-08-27.
  With an approver group empty or its Head of Unit gone, requests pile up with nobody able to decide,
  and the only workaround was adding yourself to a unit's approver group — the pollution above.
- **THE BYPASS LIVES IN `inScope`, NOT IN THE TWO CALLERS**, so visibility and decide-rights cannot
  diverge: the invariant that *a row they cannot decide is a row they do not see* holds for an admin
  exactly as for everybody else. **`canDecide` still requires the row to be PENDING**, so this grants
  no power to re-decide something already settled — pinned by test.
- **`toScope` CANNOT PRODUCE `systemAdmin` FROM A BARE `string[]`**, and it normalises with `=== true`
  so a truthy non-boolean cannot widen it and `undefined` reads as NOT an admin. The widest right in
  that file must never arrive from its oldest call shape. Three tests pin it.
- **`isSystemAdmin` FAILS CLOSED** (false on any read failure) and runs in **its own effect**, never
  folded into `load` — an admin check failing must not leave the page reporting it could not read
  requests. The two facts are independent and their failures mean different things.
- **The admin banner SAYS WHY they see everything**, because otherwise a full queue for someone
  holding no mapping reads as a leak — and it states that their name is recorded as the decider, since
  an approved deletion recycles the file and an approved share grants access.
- **No schema change, no migration, no reconciliation run.** Page ACLs already admit admins.
- **NOT YET TESTED LIVE.** Open `Request.aspx` as a `CRS Owners` member: expect the amber
  *"system administrator"* note, `Waiting for you` listing other people's pending requests, and the
  *"not recorded as the approver"* banner gone.

## THE REQUESTS PAGE IS TABBED, AND A SHARE CAN NOW BE TAKEN BACK (2026-08-28)
Client: *"can we redesign this a bit? Can we separate into like Share Category, Deletion Category and
inside have pending approval and, similar design as the My Submission? So we can track who did what as
well? Also a tab to show which user is sharing which one? And then allow HOU to revoke their sharing
permits for that file?"* Spec `2026-08-28-requests-page-redesign-and-share-revoke-design.md`; rules in
`shared/requests.ts` (pure, 57 new tests, 1509 total). **BUILT, NOT DEPLOYED, NOT SITE-TESTED.**
- **⚠ THE CLIENT'S OWN PREMISE WAS WRONG, and the redesign had to keep all four audiences working.**
  They asked *"only HOU sees their own unit request from PIC right?"* — `inScope` has **four**: HoU
  (own units, BOTH stages), **HoD (their department, APPROVED stage only)**, system admin
  (everything), and the requester (their own rows, always). The HoD stage limit is not policy — a HoD
  holds nothing in either approval library, so they physically cannot action a pending-file request.
- **Four tabs — Deletion · Share · Shared files · Your requests** — each counted in its label. Deletion
  and Share each hold **Waiting for you** and a new **Decided** section naming who decided it, when and
  what they wrote. That history did not exist: a decided request VANISHED from this page unless the
  viewer happened to be the requester, so a Head of Unit could not answer *"who approved that"* without
  opening the list itself.
- **Every tab filters the SAME `visible` array the tally uses**, so a tab and the tally cannot disagree
  — the 2026-08-21 defect, where a HoD read `2 pending` beneath a queue of one and thereby learned of a
  pending-stage row `ViewerScope` deliberately hides from them.
- **⚠ REVOKE NEEDED NO NEW PERMISSION, AND THAT IS THE WHOLE REASON IT WAS CHEAP.** `CRS Share`
  **already contains Manage Permissions** — the entire reason it exists as a separate level
  (`FolderManager.tsx` ~341). So no new role, no new level, no Group Map change, **no reconciliation
  run**. Still verify the level on the site, per the standing warning on that comment: a `CRS Share`
  without Manage Permissions breaks *approvals* too, so it is a pre-existing check.
- **TWO MECHANISMS, and only one of them actually solves the problem §5.3 was written for.**
  `removeroleassignment(principalid=N)` drops one recipient and **keeps the unique scope**;
  `resetroleinheritance` drops everyone **and reclaims it**. Every approved share creates a scope, and
  that is the same 50,000-per-list ceiling that killed per-uploader ACLs on 2026-08-06 — so
  per-recipient revocation ALONE leaves the count monotonic, which is exactly the state the 2026-08-15
  design flagged as *"without a way back out, the count only ever grows"*. The **Revoke all** button
  says on screen that it releases the scope.
  - `resetroleinheritance` returns the file to its **UNIT FOLDER's** ACL, not the library root —
    everyone in the unit keeps what they already had, which is what makes it safe to offer.
- **⚠ ROWS ARE THE INDEX; THE LIVE ACL IS THE VERDICT.** They disagree in BOTH directions and neither
  is rare: a share made through SharePoint's **own Share button** leaves no row (a web part cannot
  intercept it — settled 2026-07-23), and a share revoked in SharePoint leaves its row reading
  `Approved`. Rows-only under-reports AND over-reports. Same shape as `mergeRecords` in My Submissions.
  - **THREE STATES, NEVER TWO** — `live` / `revoked` / **`unknown`**. A throttled or refused read must
    never render as *"access ended"*: an approver told a share is over stops looking, and the document
    stays reachable. `unknown` is the AMBER one, deliberately inverting the usual instinct.
  - **⚠ THE TAB CANNOT SEE A FILE THAT WAS ONLY EVER SHARED NATIVELY**, because nothing indexes it.
    Finding those means crawling every file in `Documents` for unique permissions — unbounded. **Said
    on the screen**: a tab implying completeness about who can reach the company's documents is worse
    than one that admits its edge. It DOES catch extra recipients added natively to an indexed file
    (`granted outside CRS`) and recipients revoked outside the app.
  - **Group principals are SKIPPED.** Breaking inheritance COPIES the unit folder's groups onto the
    file, so counting them would report a unit's ordinary access as a share.
  - **The probe is STICKY** — one request per shared file, on opening the tab only, never on
    `useLiveRefresh`'s 45-second tick, or a background convenience becomes a per-file request storm.
    A **Re-check access** button is the manual refresh; a revoke re-probes its own file regardless.
  - **A revoke re-probes rather than trusting its own write.** Believing the write would make the one
    screen that checks reality the one screen that assumes it.
- **⚠ THE AUDIT ROW CANNOT BE WRITTEN FROM CODE, and this is already established.** `writeAudit` is
  REFUSED for every non-Owner — which is what made the `CRS — Audit request activity` flow necessary on
  2026-08-26 — and a Head of Unit pressing Revoke is such a non-Owner. So the **ROW** is updated
  instead (`Status` → `Revoked` when the last recipient goes; a line APPENDED to `DecisionNote`
  otherwise) and the existing created-or-modified flow picks it up.
  - **⚠ FOLLOW-UP, FLOW-SIDE: `CRS — Audit request activity` needs one branch.** Its `EventKind`
    expression ends `if(equals(Status,'Approved'),'RequestApproved','RequestRejected')`, so **until a
    `Revoked` branch is added a revocation logs as `RequestRejected`** — the actor and the row are
    right, only the label misleads. Written here because a flow change cannot live in source control.
  - The note is **APPENDED**, never overwritten: the approver's original note records WHY access was
    granted, and destroying it to record the revoke loses half the trail.
- **NO SCHEMA CHANGE.** `Status` is a **Text** column, never Choice — the 2026-08-15 design chose that
  deliberately for exactly this moment, since a value absent from a Choice column's `Choices` fails the
  whole write silently. `RequestStatus` gained `Revoked`; `counts` and `PILL` are
  `Record<RequestStatus, …>`, so **the compiler named both sites** rather than letting one drift.
- **`canRevoke` REUSES `inScope`**, so the revoke button and the queue can never disagree about which
  units a HoD may act on. `matchesUnit`/`inScope` were widened to a structural `UnitScoped` type — a
  `RequestRow` still satisfies it, so every existing caller is unchanged. **NOT the requester:** a PIC
  holds no Manage Permissions, so offering them the button would fail in their session AFTER the screen
  offered it.
- **Tell the client:** a revoke is **silent** — SharePoint sends no "access removed" email and any link
  simply starts refusing; **Revoke all** is the one that reclaims the scope; natively-shared files are
  not fully visible here and cannot be; a **Head of Department can revoke too** within their
  department, since they hold `SHARE` under the 2026-08-20 revision.

## THE REQUESTS PAGE IS APPROVERS AND HEADS OF DEPARTMENT ONLY (2026-08-21, 1.0.212.0)
Client: *"make a page for HOU where they can handle the request of deletion and share from the
Uploader themselves which is PIC, that page should only allow HOU and HOD to enter, the reason I also
include HOD is because they literally have Share and Deletion power."* Spec
`2026-08-21-requests-page-hod-access-design.md`; rules in `shared/requests.ts` (`ViewerScope`, pure,
11 new tests). **No new web part** — `Requests.aspx` already IS that page; it had the wrong audience.
- **`UPL` LEFT THE `/request/i` POLICY, AND ITS REMOVAL IS THE POINT.** The rule granted `UPL + APR`
  because the 2026-08-15 design gave the page two audiences — the uploader raising a request, the
  Head of Unit deciding it. **That reasoning went stale on 2026-08-20**, when the requester's own view
  moved to **My Submissions → Requests** with Cancel. A PIC opening the page since then gets a queue
  filtered to units where they hold `APR`, i.e. none: **an empty page on their menu, granted by a
  comment describing a design that no longer existed.** Not a leak — the render already filtered — but
  a policy that described the wrong system. Pinned by a regression test asserting `UPL` is absent.
- **⚠ A HEAD OF DEPARTMENT CANNOT ACTION A PENDING-FILE REQUEST, AND THAT SHAPES THE WHOLE DESIGN.**
  `hod` is `DEPTVIEW + DEL + SHARE` and **none of those three is in `LIBRARY_ROLES.Staging`**, so they
  hold nothing whatsoever in either approval library — they cannot read a pending file, let alone
  recycle one. Since **the approval executes in the approver's own browser session**, pressing Approve
  would fail as them and record `Failed`, with the requester told their request was being handled and
  nothing done.
  - **Pending-stage requests are HIDDEN from a HoD entirely** (client's choice, over showing them
    disabled). `isVisibleTo` and `canDecide` share ONE `inScope` helper so the two can never disagree
    about which rows a HoD may know about — a visible-but-inert row reads as a broken button.
  - Nothing is stranded: the request was routed to a unit with an `APR` mapping in the first place,
    and **`DELS` has exactly one holder (`hou`)**, so there is exactly one person who can carry out a
    pending-file deletion. This page cannot close that gap.
  - A HoD still sees a pending-stage request **they raised themselves** — the requester branch is
    about authorship, not scope, and runs first.
- **`ViewerScope` is TWO SETS, unioned, never one overriding the other.** `aprUnits` (any stage) and
  `hodUnits` (approved stage only). Someone who is Head of Unit in one unit and Head of Department
  over another gets both powers in their own places; the narrower set must never cap the wider one.
  Pinned by test.
- **A bare `string[]` is still accepted and read as `aprUnits`** — that is exactly what the array has
  always meant, so every existing caller and test stays correct. **An array can never grant HoD
  scope**, so no old call site can widen access by accident.
- **THE JOIN NEEDS NO NEW DATA.** A Group Map row already carries `Segment` = the segment's **term
  set GUID** and `UnitTermGuid` = the tier's term — the **department** on a `DEPTVIEW` row
  (`groupMapModel.ts` ~779). So a department expands to its units with ONE request:
  `/_api/v2.1/termStore/sets/{Segment}/terms/{UnitTermGuid}/children?$select=id`. A person heads one
  department in the normal case.
  - **Storing the department on the request row was REJECTED.** It is a new column, every row written
    before today would lack it, and `Stage` two days earlier showed the cost — a read retried without
    the field, a write retried without the field, and a provisioning gap to detect. Ancestry is
    derivable, and a stored copy would describe a shape a re-parented term has since left.
  - **Fails CLOSED per department, and SAYS SO.** A failed expansion shows nothing extra — an
    over-wide fallback would show one department's requests to another's head — but the page states
    the department could not be read rather than reporting an empty queue. `empty ≠ unknown`, again.
    The warning shows only to someone who actually holds a `DEPTVIEW` row.
- **The asymmetry with `ApprovalDocument.aspx` is deliberate and pinned:** `DEPTVIEW` reaches the
  requests page and must NEVER reach the approval queue — letting a Head of Department approve
  documents would let them publish into a unit they do not run.
- **MIGRATION = re-run reconciliation.** Page grants are asserted in full at page scope, so the run
  removes every `_UPLOADER` group from `Requests.aspx` (naming each removal) and adds the `_HOD`
  groups. No row change, no schema change.
- **THE COUNTS LINE IS SCOPED TO THE VISIBLE ROWS, NEVER THE WHOLE LIST (1.0.214.0).** The read
  returns every row — the list is not trimmed per unit — so counting all of them showed a Head of
  Department `2 pending` beneath a queue reading `Waiting for you (1)`. Two faults in one line: the
  page contradicted itself, which reads as a bug, and it disclosed the existence of the pending-stage
  request `ViewerScope` deliberately hides. It reuses `isVisibleTo`, so the tally and the cards cannot
  disagree.
- **⚠ STILL OPEN: the `CRS Requests` LIST inherits site permissions.** This page filters what it
  RENDERS; anyone who can open the site can read the list at its own URL — every request's file name,
  reason, share recipients and decision note, across both segments. Raised with the client
  2026-08-21. Not a document exposure (the files stay behind their folder ACLs), but it is request
  metadata with no boundary at all.

## DELETION REQUESTS NOW COVER PENDING FILES (2026-08-20, 1.0.199.0)
The other half of the PIC losing `DELS`. Until then requests were raised on an **approved** file only
— precisely because a PIC held delete in the approval library and needed no permission there.
Removing `DELS` closed that door without opening this one, leaving a PIC no route to remove their own
pending or rejected file at all.
- **`RequestStage` = `"approved" | "pending"`, OPTIONAL, defaulting to approved.** Read it through
  `stageOf`, never directly: every row written before today predates the field and every one of them
  was an approved document. Same shape as `unitTermGuid`.
- **SHARE STAYS APPROVED-ONLY, and `validateDraft` REFUSES it rather than the form merely hiding it.**
  A pending file has been approved by nobody, and `LIBRARY_ROLES` keeps `SHARE` off both approval
  libraries — so a Head of Unit could not carry out such a share even having agreed to it. It would
  fail in their own session, AFTER telling the requester yes. Hiding the button alone would turn a
  rule into a UI decision. The refusal is also the ONLY message returned (no "add a recipient"
  alongside), because that second one reads as something the requester could fix.
- **The approver's execution needed no change.** `GetFileById(guid'…')/recycle()` is WEB-scoped, so it
  already reached any library; and `hou` holds `DELS`, so the approval succeeds in their own session.
- **⚠ THE BUTTONS BELONG ON MY SUBMISSIONS, AND THE REASON IS STRUCTURAL.** They were moved to the
  library command bar on the client's request and moved **straight back the same day, on the client's
  own correction** — *"you were right and I was wrong… this should help isolate the delete and share
  for both document library and staging library even more."*
  - **The page filter IS the isolation.** Both reads are `AuthorId eq me`, so an uploader can only
    ever raise a request against something they filed — with no author check to write, and therefore
    none to get wrong. It holds in BOTH libraries and does not depend on Draft Item Security staying
    at PIC 1.
  - **A command-bar button cannot have that property.** In `Documents` a PIC sees their whole unit's
    approved files by design (2026-08-09), so the button would have let them ask for a COLLEAGUE's
    document to be deleted. An author check was possible but would have been a rule someone has to
    maintain, in place of a guarantee that maintains itself.
  - The `requestCommand` ListViewCommandSet, its bundle entry, its `componentId` and its
    `elements.xml` registration were all removed rather than left dormant — an unused second route
    into the same workflow is how the two drift. **Do not move these buttons again without replacing
    the author guarantee with a real one.**
- **⚠ THE `Stage` COLUMN DOES NOT EXIST ON A LIST PROVISIONED BEFORE TODAY, and one unknown field name
  fails the WHOLE request** (gotcha #11) — in BOTH directions:
  - **Read** (`Requests.tsx`): asked for, retried without it on **400 only**. A 404 is the LIST missing
    and is answered separately; retrying that would report an unprovisioned list as an unreadable one.
    Asking unconditionally would empty an approver's entire queue on a site that is otherwise fine.
  - **Write** (`MySubmissions.tsx`): attempted first, retried without it on 400. Refusing would stop an
    uploader raising a request at all — worse than an imprecise queue label, since the deletion
    resolves by UniqueId and works either way. **Never the reverse order**, or a correctly provisioned
    site would silently stop recording the stage.
  - **Reported, not silently tolerated**: the Requests page shows the absence and offers
    `addMissingColumns`, which adds only what is missing and reuses `COLUMNS` so there is never a
    second list of what the list should hold. `provision` CREATES the list and would fail on an
    existing one — which is why there had been no route to add a column at all.
- **`Stage` is TEXT, never Choice** — as with `RequestType` and `Status`. A value absent from a Choice
  column's `Choices` fails the whole write, so the day a third stage is added, every row of that stage
  is lost silently.
- **The stage is DERIVED from the row's moderation status**, not from which library the reader came
  from: a `Documents` row carries no status because everything there is approved by definition. That
  is the same discriminator the page already reads, so the two cannot disagree.
- **"YOU ALREADY ASKED" IS SHOWN, and it is the difference between a workflow and a suggestion box.**
  My Submissions reads the requester's OWN request rows (`RequestedBy eq me`, `Id desc`, first row per
  file wins) and, while one is Pending, **replaces the buttons with a banner**. Without it an uploader
  who asked yesterday sees a page identical to one who never asked, and asks again — so the Head of
  Unit gets two rows for one decision and cannot tell a mistake from a reminder. A **decided** request
  does not hide the buttons: a rejected deletion may legitimately be asked for again once the reason
  is fixed, and the approver's note is shown so they know what to fix. A pending request also tags its
  row in the list, beside — never instead of — the document's own status badge, because a file can be
  Approved AND have a deletion pending.

## ⚠ RECONCILIATION OFFERED FOUR SEGMENTS ON A TWO-SEGMENT SITE (2026-08-21, 1.0.207.0)
Client, reading the scope picker: *"umm why is there four business segment?"* — the CRS Config list
holds exactly two `mode` rows (GHO, MHO). The picker was showing the built-in `RECON_MODES`, which
includes **`Upstream Malaysia Head Office`**, a deliberate placeholder whose term-set GUID is stale.
Running it would have walked a term set that does not exist on the site.
- **THE CAUSE WAS A RACE, NOT THE DATA.** `loadReconModes` builds its URL from
  `cachedListTitle(LIST_SUFFIX.config)`, which answers the LEGACY `DMS Config` until `primeNames` has
  filled the cache. Priming makes ten sequential probes; the read fires on a button press. Press
  before priming settles ⇒ 404 ⇒ silent fallback. **Fixed by awaiting `primeNames` inside
  `loadReconModes`** rather than gating on `namesReady`, so it holds whichever screen or flow mounted
  the component.
- **THE TELL WAS THE LABELS, and it is the diagnostic worth reusing.** Config stores `StagingFolder`
  as `GHO` and `MHO`; the screen showed *"Group Head Office"* and *"Minamas Head Office"* — strings
  that exist **only** in the fallback constant. When a screen shows plausible data, check whether the
  VALUES could have come from where you think they did.
- **TWO WRONG THEORIES FIRST, both disproved by cheap checks.** (1) A missing `SortOrder` column
  failing the whole `$select` (gotcha #11) — the column exists, populated 1 and 2. (2) The mount
  effect not priming — it does, first thing. **The decisive step was pasting the exact REST URL into
  a browser**: it returned both rows perfectly, which ruled out the request and left only the timing.
  Ask for the raw response earlier than feels necessary.
- **THE FALLBACK IS NO LONGER SILENT.** `modeSource` distinguishes *read failed* / *no mode rows* /
  *read without SortOrder*, and the picker says outright that these are built-in segments, names the
  placeholder, and says **do not run this**. Nothing on that screen had ever distinguished "your
  configuration" from "a hardcoded list", in the one place where the cost is provisioning a segment
  nobody created. `empty ≠ unknown`, again.
- **`SortOrder` is now retried-without on a 400** anyway — harmless, and it removes a real trap on a
  site whose mode rows were authored by hand rather than by `SegmentCreator`, which is what creates
  that column. **`BulkUpload` and `DocumentSearch` select `SortOrder` too and have NOT been
  hardened** — they fall back to `DEFAULT_MODES`, the same four segments.
- **The fix incidentally protects the whole run**: the picker's read is the first thing that happens
  when the bar opens, so the cache is warm for every Group Map and library read that follows.

## ⚠ CRS REQUESTS WAS BUNDLED BUT NEVER REGISTERED — UNDEPLOYABLE SINCE THE DAY IT WAS WRITTEN (2026-08-20, 1.0.200.0)
Found when the client could not add the web part to a page. `CRS Requests`
(`7a4f1e93-2c58-4d07-b6a1-9e3c85f2d410`) was the **only one of 18 bundled components missing from
`package-solution.json` → `features[0].componentIds`** — 17 registered, 18 bundled. It therefore never
appeared in the web-part toolbox on any site, which is also why the Requests LIST has never existed:
nobody could put the web part on a page to provision it.
- **THE WHOLE DELETION AND SHARE FEATURE HAS BEEN UNREACHABLE**, while the repo, the specs and this
  file all described it as built. Every review of the code would have passed.
- **SECOND INSTANCE OF ONE FAILURE, and the first is documented three sections up.** The `+ New Folder`
  customizer sat correct, deployed and inert for months for want of a `UserCustomAction`.
  **Bundling makes a component AVAILABLE; the feature is what REGISTERS it.** The package deploys
  cleanly, the catalog says `Deployed: Yes`, and every other web part works — there is no error
  anywhere, on any screen or in any log.
- **The diagnosis that worked, and it is worth reusing:** search the toolbox for the shared prefix
  (`CRS`) rather than the exact name. Four other CRS web parts appeared and this one did not, which
  ruled out the app not being updated and the full-width-section restriction in ONE action. Both were
  live hypotheses and both were wrong.
- **NOW UNDER TEST — `shared/solutionPackaging.test.ts` (4 tests).** It cross-checks every manifest in
  `config.json` bundles against `componentIds`, in BOTH directions, and **names** anything missing
  (a count sends someone reading eighteen GUIDs). It also pins this specific id, so a future rewrite of
  `componentIds` cannot drop the one already dropped once, and asserts `skipFeatureDeployment` stays
  `false`. The one filesystem test in the repo: `fs` is `require`d behind a narrow local interface with
  a single-line lint exemption, because adding `@types/node` globally would invite Node APIs into
  browser code where they fail at runtime instead of at compile time.
- **Recovery is deploy-only** — no data change, no reconciliation. Add the id, rebuild, upload, and
  **update the app in Site Contents**.

## DELETION AND SHARE REQUESTS (2026-08-15, spec `2026-08-15-deletion-and-share-requests-design.md`)
Client: a PIC deletes in the approval library but **not** in `Documents` — there they ASK, and the Head
of Unit approves. HoD and C-Level delete and share directly, no request. Web part **`CRS Requests`**
(`7a4f1e93-…`), its own page. BUILT, **not site-tested**; the folder-scope grants and the `CRS Share`
permission level do not exist yet, so **nothing works until reconciliation is re-run after deploy**.
- **THE APPROVAL EXECUTES IN THE APPROVER'S OWN BROWSER SESSION.** Pressing Approve recycles the file
  or grants the access as *them*. No service account, no flow, nothing acting on anyone's behalf — so
  the audit row names who actually did it, and an approval **cannot exceed the approver's own rights**;
  it fails loudly instead. Cost: an approval only completes while they are on the page.
- **An approver can only approve what they can PERFORM.** That is why HoU gains `DEL` and `SHARE` — and
  why granting them sharing rights makes them **the sharing authority for their unit, not merely an
  approver of one**: they can then share directly, with no screen involved. Client decision, stated
  rather than buried.
- **PIC gained `DELS`, HoU gained `UPL`** (2026-08-15 correction — the earlier model had these exactly
  backwards). A HoU therefore **approves their own uploads**; the client accepted this.
- **Deletions RECYCLE, never purge** — restorable for 93 days, which is what makes approving one
  reasonable, and the dialog says so. Resolved by **UniqueId**, so a rename or move since the request
  was raised does not matter.
- **Shares go through `SP.Web.ShareObject`** — the endpoint SharePoint's own Share dialog calls, so it
  honours tenant and site sharing settings rather than working around them. **HTTP 200 does NOT mean it
  worked**: the per-recipient result is in the BODY, exactly as with `validateUpdateListItem`
  (gotcha #4). A tenant refusal arrives there, not as a status code.
- **External sharing is a `DMS Config` row (`allowExternalSharing`), and it FAILS CLOSED.** Absent,
  unreadable or anything but an explicit yes ⇒ internal only. A deliberate exception to this codebase's
  fail-open rule, for the same reason `canOfferFolderDelete` fails closed: elsewhere a failed read costs
  a form, here it would send a document out of the organisation on a setting nobody could confirm.
  `tenantDomains` (also a config row) is what `isExternal` compares against — **no domains supplied
  means every recipient reads as external**, never as internal.
- **A web part CANNOT intercept the native Share button** — settled in
  `2026-07-23-share-guard-retirement.md` and not re-openable. This page is the *sanctioned route*, not a
  gate; SharePoint's own "only site owners can share" lockdown remains the backstop, and a share pressed
  there lands with an administrator instead. Do not re-propose forcing everyone through this page.
- **ROUTING IS ON THE TERM GUID, and "deepest tier" is the wrong key.** `documentUnit()` derives the
  whole tier chain from the document's own `<Base>Tid` fields; `routeToApprover()` then takes the
  **deepest tier that actually has an `APR` mapping**. A below-Unit tier such as SubUnit carries a Tid
  column exactly like a permissioned one, so routing on the deepest alone would file the request where
  **nobody can see it, with nothing on screen to say so**. When nothing matches, the request is still
  sent and the requester is told an APR mapping is missing — an unreadable Group Map and a genuinely
  absent approver look identical from here.
- **`RequestType` and `Status` are TEXT, never Choice** — same reason as `EventType`: a value absent
  from `Choices` fails the whole write, silently. `RequestedAt`/`ExpiresAt` are **ISO**; a blank expiry
  is **omitted**, never sent as `""`, which a DateTime column rejects and takes the whole row with it.
- **A failed action records `Failed`, never `Approved`.** Approving a deletion for a file that has since
  gone, or a share the tenant refuses, must not read as done. `applyDecision` enforces it.
- **Every approved share creates a unique permission scope** — the same 50,000-scope ceiling that killed
  per-uploader ACLs on 2026-08-06. Workable only because shares are rare where those would have been
  universal. **Revoke is spec'd (§5.3) and NOT yet built**, so the count currently only grows.
- Requests are raised from **My Submissions**, on an **approved** file's detail view only — pending and
  rejected files are still in the approval library, where a PIC holds Delete and needs no permission.
- 📄 **Client-facing note, ready to show: `docs/client/deletion-and-share-requests.md`.** Carries the
  six things they must be told (HoU self-approval, HoU as sharing authority, the un-interceptable Share
  button, the external-sharing assumption, shares being permanent with no revoke yet, and that this
  changes deleting and not *seeing*), plus the two prerequisites and four open questions.

## ⚠⚠ EVERY METADATA FILTER 400'D — REST `$filter eq` CANNOT TOUCH A TAXONOMY FIELD (2026-09-02, 1.0.361.0)
The very next live test after the 500 fix below: free-text search worked cleanly (8 results across
all four libraries, `HC`/`Awaiting approval` tags rendering correctly), then setting ANY Advanced
Filter (Document Type / Year / Confidentiality / Business Segment) with no typed word produced
*"4 libraries could not be searched: Approval for Document (HTTP 400), Restricted & Confidential
Document (HTTP 400), Approval for Highly Confidential Document (HTTP 400), Highly Confidential
Document (HTTP 400)."*
- **THE CAUSE: SHAREPOINT'S REST `$filter` REJECTS `eq` AGAINST A MANAGED METADATA COLUMN OUTRIGHT.**
  `buildListFilter`'s `odataEquals("Business_x0020_Segment", c.segment)` (and the Year/Confidentiality
  Level/Document Type/tier equivalents) produced `Field eq 'value'` — syntactically fine REST, and a
  hard 400 anyway, because a taxonomy field is internally a Lookup-derived type and `eq` is not a
  valid operator against it. Confirmed against Microsoft's own documented guidance and community
  reports of the identical error shape.
- **⚠ ALL FOUR NAMED LIBRARIES WERE THE REST SIDE, NOT THE SEARCH API SIDE — including "Restricted &
  Confidential Document", which reads as the crawled `Documents` library and is not.** `runSearch`
  (the KQL/Search API call) returned `outcome: "ok"` the whole time — it was the SEPARATE 24-hour
  "recency top-up" REST read against `Documents`/`DocumentsHC` (`runListRead` + `buildRecentFilter`,
  which reuses `buildListFilter`) that 400'd, and both `LibraryResult`s carry `library: "Documents"`
  so `libraryLabel("Documents")` labelled the failing one with the crawled library's title. Two
  results can share a library key — one `ok`, one `failed` — and only `failedLibraries` distinguishes
  them; the banner names the library, not which of the two reads inside it broke.
- **FIXED: `odataTaxonomyEquals`, filtering `TaxCatchAllLabel` with `substringof` instead of a plain
  `eq`.** This is the standard, widely-documented SharePoint REST workaround for exactly this
  limitation — `TaxCatchAllLabel` is a hidden system NOTE field SharePoint auto-maintains on any list
  carrying a taxonomy column, holding every tagged term's LABEL. Replaces `odataEquals` at all FIVE
  call sites: the four fixed filters and the tier-value filter (every permissioned/below-Unit tier
  column is ALSO managed metadata — that is exactly why each one has a `<Base>Tid` twin — so it needed
  the identical fix, not just the four named columns).
- **⚠ UNLIKE THE KQL SIDE'S STILL-UNVERIFIED `<Field>OWSTEXT` MANAGED PROPERTY (§5/§6 of the search
  spec), THIS NEEDS NO SEARCH SCHEMA SETUP.** `TaxCatchAllLabel` is populated the moment an item is
  tagged, with no crawl involved — this fix is complete and self-contained, unlike the crawled-library
  metadata-filter gap this file already flags as open and unfixed since 2026-08-23.
- **⚠ A FALSE POSITIVE IS THEORETICALLY POSSIBLE, and accepted.** `TaxCatchAllLabel` cannot say WHICH
  taxonomy field a term came from, so a match only requires the item to carry that LABEL somewhere.
  The four vocabularies here (segments, years, confidentiality levels, document types) share no
  labels with each other, and multiple filters AND together, so a false positive would need the same
  coincidence to repeat across every filter set at once.
- **Verified**: `tsc --noEmit` clean, full suite **1598/0** (2 assertions rewritten in
  `documentSearch.test.ts` to match the new clause shape — the old ones literally asserted the broken
  `eq` syntax that was failing live), no new lint warnings. **NOT yet re-tested live** — confirm a
  metadata-only filter (no typed word) now returns clean results across all four libraries.

## THE FILTER GRID COULDN'T FIT 4 COLUMNS — ITS OWN CONTAINER WAS 2px TOO NARROW (2026-09-02)
Same test round. Client: *"I think they meant that the layout should be grid 4 not grid 3"* and *"the
box doesnt extend to the end of the page, not sure how to fix this issue as this is sharepoint native
I think."* **Not a SharePoint constraint — our own CSS.**
- **`heroInner`'s `maxWidth` was 720.** `advPanel`'s grid (`repeat(auto-fit, minmax(170px, 1fr))`, 14px
  gaps) needs `4 × 170 + 3 × 14 = 722px` minimum to hold all four Document Type / Year /
  Confidentiality / Business Segment dropdowns on one row — **two pixels over the cap** — so `auto-fit`
  had no choice but to wrap to 3-then-1 on every screen, however wide. `s.hero` (the green banner
  itself) already spans the section's full width; only this INNER content wrapper was artificially
  capped, and that capping is exactly what "the box doesn't extend to the end of the page" was
  describing — nothing SharePoint controls.
- **720 was sized for a single search bar** — the hero design predates Advanced Filters, and the
  wrapper was never widened when the 4-column grid was added on top of it.
- **Raised to 1100, generously past the 722px threshold** rather than to the exact minimum, so ordinary
  padding/gap rounding on a slightly different viewport can't reopen the same wrap. The grid itself
  (`auto-fit`/`minmax`) is untouched — still responsive, still degrades gracefully on a genuinely
  narrow screen; only the outer cap that was silently defeating it moved.
- **Verified**: `tsc --noEmit` clean, full suite 1598/0, no new lint warnings — a CSS-only change, no
  pure-module tests to update. **NOT yet re-tested live.**

## ⚠⚠ EVERY CRS SEARCH QUERY 500'D — `odata-version: 4.0` AGAIN, SIXTH TIME (2026-09-02, 1.0.360.0)
Client, testing the home page search bar: *"How come it state one library cannot be searched? What
library is that? Also, I tried to search the content inside the document but it shows error."*
Screenshot showed the exact banner this feature's own design names for a failed read: *"One library
could not be searched: Restricted & Confidential Document (HTTP 500)."* — on a query as plain as
`ocbc`.
- **THE CAUSE IS THE SAME MECHANISM THIS PROJECT HAS NOW HIT SIX TIMES, AND THIS IS THE FIRST TIME ON
  A READ.** `SPFx`'s `SPHttpClient` injects `odata-version: 4.0` on EVERY request regardless of what
  `Accept` says. Every prior instance (`spAuditLog.ts`'s `WRITE_HEADERS`, `FileTypeSettings.tsx`,
  `FolderManager.tsx`, `Requests.tsx`, `MySubmissions.tsx`, `UploadPauseToggle.tsx`) was a WRITE that
  needed `odata=verbose` for a `__metadata` envelope. `_api/search/query` is neither a write nor
  verbose — it is the LEGACY Search REST surface, older than the JSON-light conventions the rest of
  this file's `/_api/web/lists` calls tolerate, and it throws a bare, bodyless HTTP 500 under
  `odata=nometadata` + the injected `odata-version: 4.0`, confirmed against Microsoft's own known-issue
  reports for this exact combination.
- **⚠ THE FAILING CALL WAS ALWAYS LABELLED BY THE LIBRARY'S CURRENT TITLE, WHICH IS WHY THE MESSAGE
  NAMED "Restricted & Confidential Document" AND NOT "Documents".** `libraryLabel("Documents")` reads
  `documentsLibraryTitle()` — the live title, correctly resolved through the 2026-09-02 rename. The
  library itself was never the problem; only the request headers were.
- **FIXED WITH A SECOND, NARROWLY-SCOPED HEADER SET (`SEARCH_HEADERS`), NOT A CHANGE TO THE SHARED
  DEFAULT.** `jsonGet` gained an optional `headers` override; every other call in this file (the
  ordinary `/_api/web/lists` reads for segments, term options, list results, `FieldValuesAsText`) is
  untouched, because the injected header is harmless there — this codebase's own habit of a per-call-
  site header override rather than a global one, same as `spAuditLog.ts`'s `WRITE_HEADERS` being its
  own const rather than a change to every other write in the project.
- **THIS LIKELY EXPLAINS BOTH REPORTED SYMPTOMS AT ONCE.** "One library could not be searched" and
  "searching document content shows an error" both run through the same `runSearch` → `_api/search/
  query` call — free-text content search has no separate code path. Confirm both are gone on the next
  live test rather than assuming either is a distinct bug.
- **STILL UNVERIFIED, and unrelated to this fix: whether §6's `<InternalName>OWSTEXT` managed-property
  assumption holds** — the metadata-filter section immediately below this one. A 500 masked whatever
  that would have shown; once free text runs clean, re-test a metadata-only filter (Document Type,
  Year, etc. with no typed word) to see whether it returns results or an empty `TotalRows`.
- **Verified**: `tsc --noEmit` clean, full suite **1596/0**, no new lint warnings on this file
  (`documentSearch.test.ts`'s 72 tests are the pure module and untouched — this was a headers-only
  change in the SPFx component, not the shared rules). Packaged as `1.0.360.0`, shipped bundle grepped
  for the literal `odata-version":""` and confirmed present. **NOT yet re-tested live.**

## ⚠ THE HERO BANNER WAS ONLY VISIBLE TO ADMIN — NOT A CODE BUG, A SITE-ASSETS PERMISSION (2026-09-02)
Client: *"The banner cannot be seen by anyone except for Admin."* Confirmed by reading the render:
`s.hero`'s `backgroundImage` is TWO CSS layers, `url("<heroImageUrl>"), linear-gradient(...)` — the
photo is the front layer and the gradient sits behind it deliberately, "so a banner that 404s degrades
to the plain green rather than to nothing" (comment already in the file, 2026-08-30). **That degrade
is exactly what a non-admin is seeing** in the first screenshot: a clean gradient with no error, no
broken-image icon, nothing in the code path that could distinguish "Admin" from anyone else — there is
no `isSystemAdmin` check anywhere in this file (confirmed by grep).
- **SO THE FAILURE IS ON THE SHAREPOINT SIDE, AT THE IMAGE FILE ITSELF, NOT IN THIS CODE.** The default
  URL is `<site>/SiteAssets/img_kv-banner.jpg` (`DocumentSearchWebPart.ts`), resolved site-relative so
  the same package works on both sites without a literal. An admin loading it successfully while
  everyone else's identical request silently fails to paint is the signature of either:
  1. **Content approval is ON for Site Assets and the file is sitting Pending** — the exact trap this
     project has hit repeatedly on `Documents`/the approval libraries (moderation hides a pending item
     from everyone but its author and approvers, with every permission read looking correct).
  2. **Site Assets (or just this file) has unique permissions that exclude `CRS_SITE_MEMBERS`** — Site
     Assets inherits site permissions by default, so this would mean something was changed by hand.
- **THE CHEAP DIAGNOSTIC, before touching anything in SharePoint:** have a NON-ADMIN account open the
  image URL directly in a browser tab and report the status. `200` with the photo rules out both
  theories above and points somewhere else entirely; `403` confirms a permissions/approval problem;
  `404` means the file genuinely is not there for that account to resolve (also consistent with either
  theory, since SharePoint security-trims to 404 as often as it answers 403 — this project's own
  gotcha, recorded repeatedly elsewhere in this file).
- **No code change here — there is nothing to fix in `DocumentSearch.tsx` or `DocumentSearchWebPart.ts`
  for this.** If the cheap diagnostic confirms content approval, the fix is Site Assets → Library
  Settings → Versioning Settings → *Require content approval* = No (or approve the specific file).
  If it confirms a permissions gap, the fix is restoring Site Assets' inheritance or granting
  `CRS_SITE_MEMBERS` Read on it directly.
- **✅ THE DIAGNOSTIC CAME BACK, AND IT NAMES THEORY 2 — A PERMISSIONS GAP, NOT CONTENT APPROVAL.** A
  genuine uploader-only account hitting the banner URL directly landed on SharePoint's own
  `AccessDenied.aspx` — *"You (clarencechojinheng@gmail.com) don't have access to this item… Request
  access."* That is specifically SharePoint's standard missing-Read-permission flow; a Pending
  (moderation) item produces a different failure shape entirely, not this one. **Still no code
  change** — this confirms the fix is on Site Assets' permissions (restore inheritance, or grant
  `CRS_SITE_MEMBERS` Read directly on the file), not in the web part.

## ⚠ CRS SEARCH'S METADATA FILTERS RETURN NOTHING — VERIFIED ON SITE 2026-08-23, NOT FIXED
The one unverified assumption in the 2026-08-16 design was that SharePoint auto-creates queryable
`<InternalName>OWSTEXT` managed properties. **It does not hold on this tenant.** Checked with four
Search REST calls, in an order where each rules out a different cause:
1. `path:"…/Shared Documents"` → **180 rows**. The library IS crawled, so nothing here is latency.
2. `selectproperties='DepartmentOWSTEXT,UnitOWSTEXT,Business_x0020_SegmentOWSTEXT'` → no error.
3. `DepartmentOWSTEXT:"Group Finance"` → **TotalRows 0**.
4. Same select restricted to `IsDocument:true` → **15 real documents, every one of the three NULL**,
   on documents that demonstrably carry Department and Unit in the library.
- **⚠ TEST 4 IS THE ONE THAT COUNTS, AND TEST 2 ALONE WOULD HAVE MISLED.** An earlier run without
  `IsDocument:true` returned only FOLDERS (`GHO`, `MHO`, `GF`…), which legitimately have no
  Department — so their nulls looked like evidence and were not. **Always exclude folders before
  concluding a metadata property is empty.** The control for test 2 (selecting a deliberately
  nonexistent property, to learn whether `selectproperties` validates at all) was never run, so
  whether the properties EXIST-but-are-unmapped or do not exist is still open.
- **CONSEQUENCE, and it is the failure mode the design named:** a tier filter that matches nothing
  reads as *"there are no such documents"*, not as a broken filter. Do not describe CRS Search's
  metadata filters as working.
- **What still works, and is untouched by this:** the REST `$filter` half over the two approval
  libraries (no KQL involved), and free-text KQL over filename and contents. **Only the
  Documents-side metadata filters are dead.**
- **THE FIX IS NOT SAME-DAY.** Site Settings → Search Schema → map the crawled property (e.g.
  `ows_Department`) to a `RefinableString00`–`99`, THEN re-index — which is asynchronous and
  routinely takes hours. `managedProperty()` in `shared/documentSearch.ts` would also have to emit
  the refinable name instead of `<Name>OWSTEXT`; it is isolated in that one function precisely so
  this swap is cheap. A site collection admin can do all of it without tenant access.
- Paused 2026-08-23 on the client's call, with one day to migration: a change that cannot finish
  before a re-index is not a change worth starting that day.

## CRS SEARCH (2026-08-16, spec `2026-08-16-document-search-design.md`)
Client: *"add a search bar in homepage that allows them to search files via metadata"*. Web part
**`CRS Search`** (`8b2f4a95-…`, own bundle), for the **home page**. Rules in `shared/documentSearch.ts`
(pure, 72 tests). BUILT, **not site-tested**.
- **THE ASK WAS CORRECTED MID-CONVERSATION AND THE SECOND VERSION GOVERNS.** It began *"only user who
  created that file can see it"*; asked what a non-uploader should see, the client replied *"the
  security part follows the today's hierachy of how the group works. PIC sees only their own unit, HOU
  sees the unit, HOD sees the entire unit and segment and glboal you get the gist"*. The first reading
  is My Submissions, which already exists, and would have shown a Head of Department **nothing**.
- **IT ENFORCES NO PERMISSIONS, AND MUST NEVER BE CHANGED TO.** Search applies ACLs at query time and a
  REST list read returns only items the caller can open, so the hierarchy falls out of the folder ACLs
  reconciliation already grants. That is the **safe** direction: a bug here can only return FEWER rows
  than the person is entitled to. A role check written here could return more, and would be a second,
  drifting copy of `groupMapModel.ts`. Say plainly to the client: **a PIC finds their whole unit's
  approved documents**, exactly as browsing already shows them.
- **TWO ENGINES, and the split is not arbitrary.** Search/KQL for `Documents` + `HC Documents`; REST
  `$filter` for the two approval libraries.
  - `$filter` **cannot** search the approved side: SharePoint throws the **5,000-item list view
    threshold** on any filter it cannot serve from an index, and **`substringof` cannot use an index at
    all**. Past 5,000 items free-text search FAILS rather than returning less.
  - Search cannot serve the approval side: crawl latency, and that library holds the just-uploaded file
    in front of the one person certain it exists.
- **THE GAP BETWEEN THEM IS REAL AND IS CLOSED BY A THIRD READ.** A file approved two minutes ago has
  been moved out of the approval library and deleted from the source (that delete is load-bearing for
  security), and is not yet crawled — so it is in **neither** engine. `buildRecentFilter` adds a
  `Modified ge <24h ago>` read of the approved side, merged and **deduped on `UniqueId`, never on path**
  (a moved file is still one file). `Modified` is indexable and the window keeps the set tiny, which is
  exactly why the MAIN read could not be done this way. It runs on every search, not only on an empty
  one — the gap is per document, not per query.
- **THE TIER FILTERS ARE DERIVED FROM THE SEGMENT'S `Levels` CHAIN**, cascading over the term tree. A
  hardcoded `Department`/`Unit` pair is the bug already fixed once in the details panel: on Upstream Ops
  (`Region`, `Estate/Mill`) both rows read blank. Here it would be worse — **a filter that silently
  matches nothing reads as "there are no such documents"**. Tier options are children of the tier above,
  so choosing a new Department **clears everything below it**, or a stale Unit queries a combination
  that cannot exist.
- **⚠ ONE UNVERIFIED ASSUMPTION, isolated in `managedProperty()` for exactly that reason.** KQL filters
  only on MANAGED PROPERTIES, and this assumes SharePoint auto-creates queryable `<InternalName>OWSTEXT`
  / `OWSDATE` ones. **Verify on the live site before trusting the dropdowns.** Fallback: map crawled
  properties to `RefinableString00`–`99` in **Site Settings → Search Schema** — a site collection admin
  can do this **without tenant access**, the same route as the in-site term store pivot. Encoded names
  keep their encoding (`Business_x0020_SegmentOWSTEXT`); decoding names a property that does not exist,
  and **a KQL clause naming a nonexistent property matches nothing rather than erroring**.
- **`kqlPathScope` returns BLANK rather than an unscoped query** — an unscoped KQL query searches the
  whole tenant. `buildKql` returns blank for blank scope or no criteria, and blank means *do not run*,
  never *match everything*. `IsDocument:true` excludes folders (`FSObjType eq 0` on the REST half — the
  REST spelling `FileSystemObjectType` is rejected inside a `$filter`).
- **Dates differ per engine and must not be unified:** KQL takes `2026-08-16`; OData takes ISO
  `datetime'2026-08-16T00:00:00Z'`. Gotcha #1's `M/D/YYYY` belongs to `validateUpdateListItem` and is
  wrong in both. Display is `DD/MMM/YYYY`.
- **`refused` is a THIRD outcome, not a failure.** An uncleared user's HC read 403s, and that is the
  correct answer to their query — counting it as an error puts a permanent warning in front of everyone
  who is not HC-cleared and trains them to ignore it. `failedLibraries` never names a refusal.
- **A partial failure shows its rows AND says what is missing**, naming the library and the status.
  Three empty states as everywhere else; `error` explicitly says it is **not** a statement that nothing
  matched.
- **Not added to the CRS Settings landing page** — that page is admin-only by the client's instruction,
  and this is for everyone. It goes on the home page.

## THE `+ New Folder` BUTTON WAS BUILT AND NEVER RAN (2026-08-16)
Client asked for the library command bar to read **`+ New Folder`** instead of **`+ Create or upload`**.
The code for it had existed for months in `extensions/hideAppBar/HideAppBarApplicationCustomizer.ts`
(`_injectNewFolderButton` hides the native button and drives its "Folder" menu item behind the scenes),
was bundled in `config.json`, and its GUID was in the feature's `componentIds`.
- **BUNDLING A COMPONENT ONLY MAKES IT AVAILABLE.** An application customizer runs only where a
  `UserCustomAction` points at it, and **no registration existed anywhere in this repo** — no
  `sharepoint/assets/`, no script. So the extension was deployed, correct and completely inert, with
  nothing on screen or in any log to suggest the feature had ever been written. Fixed by adding
  `sharepoint/assets/elements.xml` + `features[0].assets.elementManifests` in `package-solution.json`.
- **`skipFeatureDeployment` MUST STAY `false`, and it was `true` — which is why adding `elements.xml`
  alone changed nothing.** That flag tells SharePoint not to activate the solution's feature on the
  site, and the feature is the only thing that provisions `elements.xml`. Silent by construction: the
  package deploys, the bundle loads, every web part works, and the customizer never runs, with nothing
  connecting the two. `false` requires the app to be added per site rather than offered tenant-wide —
  already this project's model (site collection app catalog). **Do NOT add an explanatory `_comment_`
  key to `package-solution.json`**: the schema forbids additional properties and the build fails.
- **AFTER UPLOADING THE PACKAGE, the app must be UPDATED in Site Contents.** Feature elements run on
  install/update only. The bundle refreshes on its own, so web parts appear to update while the
  customizer does not — which reads as "the button fix didn't work" rather than "the app wasn't
  updated".
- **⚠ ON ClarenceDMSTesting THE REGISTRATION IS MANUAL, NOT FROM THE FEATURE (2026-08-16).** Even after
  `skipFeatureDeployment: false` and an app update, the element never provisioned, so the
  `UserCustomAction` was POSTed by hand and the button then worked immediately. **Do not read this
  site's working button as proof that `elements.xml` provisions.** On SDG's tenant, deploy, then run
  the usercustomactions check below; if it returns 0, register it by hand the same way. Verified live:
  the customizer's DOM selectors are current and correct — registration was the only thing missing.
- **CHECK FOR A HAND-MADE REGISTRATION BEFORE DEPLOYING:**
  `/_api/web/usercustomactions?$select=Title,Location,ClientSideComponentId`. Two registrations of one
  component load it twice — two buttons, two MutationObservers.
- **The upload-hiding gate was STALE and silently matching nothing.** It tested
  `pathname.includes("/staging")`, which stopped matching the day the library was recreated as
  `Approval Document` at `/ApprovalDocument`. Now gated on the list TITLE against all CRS libraries
  (`_isCrsLibrary`) — a title is also the only thing that identifies `Documents`, whose URL segment is
  `Shared Documents` (gotcha #12). Applies to **every** CRS library now: every upload is meant to arrive
  through the form, which is what gives a document its metadata, routing and approval trail.

## My Submissions (2026-08-14, spec `2026-08-14-my-submissions-design.md`)
Client, for the uploaders: it is difficult to track what you uploaded. Web part **`My Submissions`**
(`5c9d1a83-…`), its own page, **uploaders only**. BUILT, not yet site-tested. No longer strictly view
only — it now raises the deletion and share requests above, which write a row and change no document.
- **It reads BOTH libraries, because a file's life spans two.** Pending and rejected sit in the
  approval library; an approved file has been MOVED to `Documents` and deleted from the source, so one
  library shows half a lifecycle. Both reads filter `AuthorId eq <me>`, and Auto-route preserving the
  uploader in `Author` (verified 2026-08-08) is what makes the halves joinable at all. `Documents`
  rows need no status read — everything there is approved by definition, and moderation is OFF.
- **IT DOES NOT MAKE APPROVED DOCUMENTS PRIVATE, and must never be described as doing so.** Pending
  and rejected privacy is real, enforced by Draft Item Security. Approved files live in `Documents`,
  where every PIC reads their whole unit's approved documents by design (2026-08-09) — so filtering
  to the signed-in user is a **convenience, not a boundary**. The client was told and **accepted**
  this (2026-08-14: *"as long as the file inside this webpart is not exposed to another uploader that
  is fine, the documents library they accepted the fact its going to shown by everyone else since its
  approved"*). The requirement is that the PAGE shows one person's files — which it does.
- **Access is the PAGE grant, not a code gate.** `pageAccessPolicy.ts` carries a dedicated rule
  (`/submission|my.?upload|my.?file/i` → `["UPL"]`) placed **BEFORE** the generic `upload` rule.
  Without it `My-Submissions.aspx` matches nothing, takes `DEFAULT_POLICY` = `[UPL, APR, DELS]`, and
  offers an uploader's own-files page to the people it is private from. Pinned by tests, the ordering
  included — both rules yield `UPL`, so only the reason line distinguishes them.
- **The moderation mapping is NOT redefined** — `mySubmissions.ts` re-exports `statusToDecision` from
  `approvalQueue.ts`. Two mappings of one field is how the approver's screen and the uploader's screen
  end up disagreeing about whether a document was approved. Draft (3) reads as Pending.
- **The rejection comment is optional by construction.** `OData__ModerationComments` is asked for and
  the read is RETRIED WITHOUT it on failure — a `$select` naming an absent column fails the WHOLE
  request (gotcha #11), so asking unconditionally would blank the page on a site that lacks it. A
  rejected row with no comment says so rather than showing nothing.
- **"Could not read" is never rendered as "you have no files."** An uploader told they have nothing,
  when a library was merely unreachable, uploads the file again — and now there are two. Three empty
  states, as everywhere else in this codebase.
- Row keys are `library#itemId`: item ids repeat ACROSS libraries, and keying on the id alone drops a
  row silently.
- **FIVE bugs found on site test (2026-08-14) — 1-4 on the first pass, 5 on the second, all worth
  remembering:**
  1. **`Document Type` is MANAGED METADATA, so `$select` returns an OBJECT** (`{Label, TermGuid,
     WssId}`) or a bare lookup id. `(v ?? "").trim()` on it threw *"(intermediate value).trim is not
     a function"* and killed the whole page with an error naming no field. Every SharePoint field now
     goes through `textOf()`. **The row interface typed `string` was the real culprit:** it made
     TypeScript vouch for what SharePoint does not guarantee, so the build was green and the runtime
     was not. `RawRow` is an index signature now, deliberately.
  2. **`FSObjType eq 0` was missing, so FOLDERS were listed as submissions** — 443 rows on the test
     site, and clicking one "opened the document library" because a folder's link IS a library link.
     That was the reported symptom; the folder bug was the cause. Note `FSObjType`, not
     `FileSystemObjectType` — the REST name is rejected in a `$filter`.
  3. **`Documents` has the URL segment `Shared Documents`**, so a hardcoded `"Documents"` left
     "Shared Documents" at the head of every approved file's folder trail. Gotcha #12 again; the
     segment is now read off `RootFolder/ServerRelativeUrl`, falling back to the title.
  5. **THE DETAIL PANEL'S FIELD LIST WAS HARDCODED, so a segment's tiers VANISHED from it** (found
     2026-08-14, spec `2026-08-14-document-details-panel-design.md`; rules in
     `shared/documentDetails.ts`, pure, 40 tests). It named `Department` and `Unit` literally, so on
     **Upstream Operations Malaysia** — tiers `Region` and `Estate/Mill` — both rows read blank,
     blank rows are dropped, and the two values that decide where the file lives were simply absent.
     Not segment-specific: the admin NAMES the tiers at onboarding, so a list of field names could
     only ever be right for the segments it was written for, and **every future onboarding would
     silently lose its own tiers**.
     - **THE TIER ROWS ARE NOW DERIVED, with no extra request.** Every tier column has a
       `<Base>Tid` sibling (that is how `ensureColumn` makes them) and `FieldValuesAsText` returns
       EVERY field including the Tid ones — so **a field is a tier field iff its Tid twin is
       present**. `Year`/`Document Type` are managed metadata with **no** Tid column, which is
       exactly why they stay in the fixed list; that absence is load-bearing, not an oversight. Tid
       columns are never displayed — they hold the GUIDs the term store exists to hide.
     - **The double-encoded key is DERIVED too:** `name.replace(/_/g, "_x005f_")`. Both detail views
       listed the two spellings per field BY HAND — a silently blank row per typo, and why a stale
       `Year_x002f_Period` fallback was still being carried after that column was renamed.
     - Order comes from response order (= field creation order = chain order); **Business Segment is
       pinned first** because that position is the only semantically load-bearing one. Labels decode
       `_xNNNN_` to a space, recovering the real name; a name sanitized at creation cannot be
       (`Estate/Mill` → column `EstateMill` → label `Estate Mill`) and the separator is NOT guessed
       back. Exact titles would need a `/fields` read — declined for one character.
     - Also added: **Location**, **Details** (`_ExtendedDescription`), **File size** (`File/Length`
       is a STRING of bytes — raw, it reaches the screen as `1483776`) and **Last updated**
       (`Modified`; the gap from `Created` is when the APPROVER acted). `File/Length` needs
       `$expand=File`, so the read gained a **last-resort retry that drops both new fields** — a
       decorative panel row must never be why an uploader is told they have no files.
     - **`ApprovalDocument.tsx:704-705` HAS THE SAME BUG and is NOT yet fixed.** It matters more
       there: the approver is deciding whether to publish into that unit, and its own comment says
       the full labels exist because the folder path is abbreviated. `buildDetailRows` serves both;
       note its `pick` returns `"—"` where the shared module drops the row.
  4. **`FieldValuesAsText` is a PER-ITEM endpoint** (as ApprovalDocument already knew), so the list
     cannot show metadata at all without one request per row. Metadata therefore lives ONLY in the
     detail view — which is also what turns a taxonomy value into a readable LABEL instead of `15`.
- **Clicking a file opens an IN-PAGE detail view** (2026-08-14, client: *"client wants the file to be
  open inside the page… like the current approvaldocument.aspx"*): preview on the left, metadata on
  the right, Back at the top, rejection reason called out. The preview reuses `previewTarget` from
  `shared/filePreview.ts` rather than re-guessing — the case that bites is invisible until it does,
  since SharePoint serves an Office file as a DOWNLOAD and a raw URL in an iframe renders nothing.
  Images fit to WIDTH and scroll, the same fix the approval page needed.
- **Provisioning:** index **`Created By`** on `Documents`. Past 5,000 items the `AuthorId` filter
  starts failing; the symptom is the error state rather than silence, but it is still a five-minute
  fix that has to be remembered.

## THE UPLOAD FORM IS BATCH CARDS NOW (2026-08-23, client's mockup + five rounds of live feedback)
Client supplied a design: a page of numbered batch cards, each with **1. Document folder information
· 2. Upload documents · 3. Document information for each file**. Built in `Form.tsx`; verified on
site by the client the same day. No spec — the mockup and the exchange below are the record.
- **THE STATE MODEL DID NOT CHANGE, AND THAT WAS THE POINT.** The cards imply per-batch pickers; there
  is still ONE set, and the destination is snapshot into the batch at save. **Only one batch is
  editable at a time.** Per-batch picker state is the exact shape that produced the silent
  HC-clearance failure of 2026-08-22 — a write-time read of the live dropdowns — which refused every
  HC upload for weeks and read as a permissions problem. Do not "improve" this into two open cards.
- **`Batch.editState` IS A SECOND SNAPSHOT, SEPARATE FROM `destination`, ON PURPOSE.** `destination`
  is what the upload WRITES and must never move once taken; `editState` (`uploadMode`, `levelValues`,
  `tierValues`, `remark`) is what the PICKERS need to show the batch again. One field serving both
  would let the act of re-opening a batch change where it files. A batch saved before this field
  existed refuses to re-open and says so, rather than opening a form pre-filled with whatever is on
  screen.
- **⚠ `restoreCascade` IS NOT A LOOP OVER `onLevelChange`.** That reads `levelValues`/`levelChoices`
  out of the render closure, so awaited calls in sequence all see pre-first-call state and the batch
  comes back with only its top level filled in. It builds both arrays locally and sets them once.
  Declared below `applyRestrictedMode` because `no-use-before-define` is on.
- **THE FLOW, as the client settled it over five rounds:**
  - `+ Add another batch` sits ABOVE every card and only OPENS a form. It does not save.
  - `Save batch` lives IN the card. On save the card CLOSES and nothing reopens — an auto-appearing
    next batch made the page look like it held unsaved work and made the Add button pointless.
  - **THERE IS ALWAYS A FORM.** Discarding the open card clears it and only closes it when a saved
    batch remains; deleting the last saved batch reopens one. The client hit both dead ends.
  - **ANY open card blocks Upload once a batch is saved** — not just one holding files. Testing for
    files let a half-filled Batch 2 be left behind while Batch 1 uploaded. Cannot deadlock: the
    always-a-form rule only holds when nothing is saved.
  - Upload **says what is outstanding on click** rather than going grey. A disabled primary button
    states that something is wrong and not what, and its `title` only appears on hover — so the one
    person who needs the message never sees it.
- **REQUIRED FIELDS TURN RED, WITH A LINE UNDER THE BOX** (client: *"the error is so vague I as a user
  did not realize"*). Gated on `showErrors`, set only by a FAILED SAVE: colouring empty fields on load
  makes a blank form look broken and people stop reading red. Each field's error is derived from its
  own value, so it clears itself and the others stay marked. `renderSelect` carries it, so all six
  folder selects got it from one change. `missingForFile` is still the only definition of required.
- **FIELD ORDER IS FILENAME ORDER**: Project Name (full width), then Vendor/Customer + Document Name,
  then Date + Confidentiality — matching `[Project] - [Vendor] - [Document Name] - [Date]`. Document
  Name used to lead, which read as "the name of the file" when it is one SEGMENT of it. The missing-
  field message lists them in the same order. **If the naming convention changes, this layout moves
  with it.**
- **The "Saves as" preview lives on the FILE ROW**, not under Document Name — it is built from four
  fields, so anchoring it to one read as a preview of that box. Open row only: it reflects the live
  editor, which describes whichever row is open.
- **⚠ `has-file` ON THE DROPZONE KEYS ON `draftFiles.length`, NEVER `file`.** `file` holds only the
  last document ADDED, so a batch re-opened for editing has files and no `file` — the zone drew its
  empty stacked layout while saying "1 document in this batch".
- **⚠ `.dms-staged-row` MUST NOT HAVE `overflow: hidden`.** The per-file editor renders inside it and
  the Confidentiality / Legally Privileged info panels are absolutely positioned children — clipping
  the row cut the tooltips off, so the definitions the icon exists to show were unreadable. The
  rounded corners come from `.dms-staged-head`.
- **Legally Privileged has no label above it**, so in the flex row it aligned with the LABELS beside
  it rather than the inputs. `.dms-lp-wrap` adds the 20px label row it does not have.
- **⚠⚠ THE STYLES ARE A JS TEMPLATE LITERAL — A BACKTICK IN A CSS COMMENT ENDS IT.** This broke the
  file THREE times on 2026-08-23, once for 411 parse errors, each time from a comment like
  `overflow: hidden` written with backticks. There is a warning in the block itself. Inside a JSX
  attribute list only `{/* */}` or `//` work, never a bare `/* */`.
- **⚠ Form.tsx WAS TRUNCATED TO ZERO BYTES on 2026-08-23** by a patch script opening it `"w"` (which
  truncates) and then throwing on a lone-surrogate escape. Recovered EXACTLY from
  `lib-commonjs/webparts/form/components/Form.js.map` → `sourcesContent[0]`, which carries the full
  original TSX. **That source map is a real backup after any `heft test` run.** Write to a temp file
  and `os.replace`; never emit unicode escapes into a patch.
- **Deliberately NOT done, and still open:** Remark stays PER FILE (the mockup puts it in the folder
  card, which would make every file in a batch share one — a data-model change touching
  `documentDetails.ts`, the My Submissions batch view and a partition test); the mockup's "up to 100
  files (max 20MB per file)" is neither shown nor enforced; Year and Document Type are NOT
  pre-defaulted, because they are folder tiers and a default files documents silently.

## BATCHED MULTI-FILE UPLOAD (2026-08-15, spec `2026-08-15-batched-multi-file-upload-design.md`)
Client: *"its only allowing one file per upload, client wants each file to go to different places"*,
refined to *"they want batches… save batches and add more"*. BUILT in 1.0.104.0, **not site-tested**.
- **A BATCH IS ONE DESTINATION FOLDER PLUS ITS FILES.** Tiers, Year and Document Type belong to the
  batch (they *are* the folder); name, project, vendor, date, confidentiality and privilege are per
  file. Save stages; one Upload sends everything.
- **THE DESTINATION IS SNAPSHOT AT SAVE, never referenced.** By Upload the pickers describe whichever
  batch is being edited *then*, so a live read writes batch 1 into batch 3's department — silently,
  into a folder that exists and looks correct. It snapshots the **leaf term id, not the folder URL**:
  resolving from the term at upload is what makes this rename-proof, and freezing a URL would undo it.
- **THE UPLOADED NAME IS NOT THE TYPED NAME.** `composeUploadBase` builds
  `[Project] - [Vendor] - [Document Name] - [Date]`, so two files with DIFFERENT typed names collide
  when project, vendor and date match — the common case, not the exotic one. `StagedFile.finalName`
  carries the composed name and `collisionsWithin` compares that; it is never re-derived in
  `shared/uploadBatches.ts`, or the convention would exist twice and be free to drift.
- **SUCCESS REMOVES, FAILURE STAYS.** That one rule makes Retry safe — an uploaded file leaves the
  list, so it cannot be sent twice — and what remains on screen is exactly what still needs doing.
  Rollback was rejected: it means deleting files that already uploaded, which can fail on its own.
- **THE STALE-CHAIN GUARD GOT GENTLER AS IT GOT MORE LIKELY.** Gotcha 10b's *"reload the page"* would
  destroy every staged batch, so a moved segment marks only **its own** batches, which stay staged and
  ask for a re-pick. The chain is re-read **once per distinct segment**. A chain that could not be
  READ marks nothing — unknown is not changed.
- **NEVER OVERWRITE; the "Replace Existing File" prompt is GONE.** A batch runs unattended so nobody
  can answer it, and the file being replaced may already be Approved and routed to `Documents`.
- **STAGED WORK CANNOT BE PERSISTED** — a `File` is not serialisable, so a closed tab loses it. A
  `beforeunload` guard and a *"nothing has been uploaded yet"* header are the whole mitigation; Cancel
  confirms with a count. This is the accepted cost of staging over commit-as-you-go.
- Rules live in **`shared/uploadBatches.ts` (pure, 48 tests)** because batching was built into
  `BulkUpload.tsx` in 2026-07-24 and **torn out** on 2026-08-03 — 169 references in a 2,586-line file
  over the lint limit. `Form.tsx` is the same size. Bulk Upload's deferred Phase 2 is the same model:
  a batch there carries one metadata set, which is a difference in what a batch *contains*, not *is*.

## ⚠ RETIRING A SEGMENT ONLY EVER DELETED THE APPROVAL LIBRARY, AND ONLY EVER COUNTED IT (2026-08-26, 1.0.259.0)
Client, straight after retiring GHO with the folder box ticked: *"I deleted all the GHO but it only
delete approval document."* Correct — and it was **two bugs in one two-element literal**, in BOTH the
delete loop and `countSegment`, which each read `[libraryUrlSegment(), DOCUMENTS_LIST_TITLE]`.
- **1. REGISTER #15, THIRD INSTANCE.** That literal predates the HC pair and the archive and was never
  revisited, so **HC Approval Document, HC Documents, Archive and HC Archive were never counted and
  never deleted.** Exactly the defect `libraryTargets()` was created to end — its own comment says so
  — and exactly what bit `SubtreeMigrator` on 2026-08-19.
- **2. ⚠ GOTCHA #12, AND THIS IS THE ONE THAT HID THE WHOLE THING.** `DOCUMENTS_LIST_TITLE` is the
  library's **TITLE** (`Documents`), while every URL in this file is built from a **URL SEGMENT** — and
  that library's segment is **`Shared Documents`** (`naming.ts` exports `DOCUMENTS_URL_SEGMENT` for
  precisely this). So the normal Documents library resolved to `/sites/…/Documents/<SEG>`, which does
  not exist, `walkFolders` took its deliberate *"404 = this branch is not in this library, which is not
  an error"* branch, and Documents counted as **zero folders and zero documents** — silently, for as
  long as this screen has existed. So the dialog's *"across both libraries"* was reading exactly ONE.
- **⚠ AND IT QUIETLY UNDERMINED 1.0.258.0, SHIPPED AN HOUR EARLIER.** That change pre-ticks the folder
  box when `documents === 0` — computed from this count, i.e. from one library. A segment full of HC
  documents would have read *"(they are empty)"* and pre-ticked. Harmless ONLY because the delete
  missed those libraries too: **fixing the delete without fixing the count would have created a real
  data-loss path.** They had to move together, and that is why one function now feeds both.
- **`retireLibraries()` is the single definition**, `libraryTargets()` filtered by KEY — so a future
  non-archive library arrives automatically (the point of deriving it) while the archive stays out
  deliberately. **`.urlSegment` for every request, `.title` only for the log line** an admin reads.
- **THE ARCHIVE PAIR IS EXCLUDED BY THE CLIENT'S DECISION (2026-08-26):** 7-year retained records
  outlive the segment that produced them. Offered as all-six vs the-four-named; they chose four.
  `survivorLines` gained an optional `hasArchive` flag and now says outright *"The archive folders are
  NOT deleted"* — shown only when folders ARE going and the site actually has an archive, because
  saying nothing leaves an admin who ticked the box finding `Archive/<SEG>` in the next reconciliation
  log. Cost, accepted and worth restating to the client: those folders stay, so reconciliation keeps
  reporting them as strays.
- **The count sentence is derived too** (`N librar(y|ies)`), never the literal "both" that was wrong in
  two directions at once — four on an HC site, two without one.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, `segmentDeletion` **26/26**
  (three new tests pin the archive line: shown on an archive site, absent when the flag is false OR
  omitted, and never shown when no folder is being deleted). Packaged as `1.0.259.0`.
  **NOT yet re-tested live, and this is the one to actually test** — retire a throwaway segment with
  the box ticked and confirm the run log names **four** libraries, including `Documents/<SEG> moved to
  the recycle bin`, which has never once happened before this build.

## THE FOLDER-DELETE CHECKBOX IS PRE-TICKED FOR AN EMPTY SEGMENT ONLY (2026-08-26, 1.0.258.0)
Client, on finding GHO's folders still standing after retiring it: *"it's because I did not select Also
Delete the folders? can we make this as default? The whole point of retiring a segment is to delete
everything."* Correct diagnosis — the box was unticked, and the dialog's own "What survives" text had
said so at the time (*"Every folder and every document stays exactly where it is"*).
- **⚠ TIED TO THE DOCUMENT COUNT, NOT FLAT-ON — and the reason is that the checkbox's LABEL is
  conditional.** With no documents it reads *"Also delete the folders (they are empty)"*; with
  documents it reads *"Also delete the folders — including N documents"*. A flat default-on would
  therefore point the default action of an in-use segment at recycling real files. Retiring normally
  happens AFTER the migrator has moved documents out — which is exactly the `documents === 0` case —
  so the default is given where it is harmless and withheld where it is not. Flat-on was offered and
  declined in favour of this.
- **`state === "counted"` is required as well.** An UNKNOWN count makes `canOfferFolderDelete` refuse
  to render the checkbox at all, and a hidden-but-true flag would be a deletion the admin was never
  shown. `alsoFolders` in `onDelete` re-checks the same guard independently, so the flag being wrong
  could not act on its own.
- **No gate was removed.** `needsTypedConfirmation` returns true whenever `deleteFolders` is true, so
  a pre-ticked box still demands the segment's label typed out — the pre-tick changes the default, not
  the confirmation.
- **GHO hid this all day**: it held 0 documents, so the difference between the two designs was
  invisible on the one segment being tested. SDG's segments will not always be empty.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full `heft test` green
  (`segmentDeletion` 23/23 — the pure module is untouched, the change is in `SegmentCreator`'s
  `openDelete`), packaged as `1.0.258.0`. **Not yet re-tested live** — deploy, open Retire on a
  segment with 0 documents and confirm the box is ticked, then on one WITH documents and confirm it
  is not.

## SWITCHING SEGMENTS FAST ON BULK PROVISIONING RACED, AND THE CREATE BUTTON STAYED CLICKABLE (2026-08-26, 1.0.257.0)
Found live, rebuilding GHO: picking GHO then quickly picking MHO produced a preview reading **"MHO_GCA_HOD
— already there as GHO_GCA_HOD"** — group names correctly prefixed for MHO, built from GHO's department/
unit chain, matched as already-existing against GHO's real groups of a DIFFERENT name. It "eventually
refreshed" to the correct MHO-only preview on its own, which is what made it look harmless — but for the
whole window before it did, the plan on screen was a hybrid of two segments' data.
- **THE CAUSE: `loadForSegment` has no out-of-order guard.** Each segment pick fires an async
  `readSegmentData` call with no cancellation, and the responses can resolve in either order. Picking
  GHO then MHO quickly enough for GHO's (slower) read to resolve AFTER MHO's is exactly what happened:
  MHO's response set `rows`/`existingRows` correctly, then GHO's LATE response overwrote them —
  `seg`/`plan` stayed keyed on the now-current MHO selection, `rows`/`existingRows` reverted to GHO's,
  and `planBulkGroups` built its output from that mismatched pair.
- **⚠ NOT MERELY A DISPLAY GLITCH — THE CREATE BUTTON WAS GENUINELY CLICKABLE ON THE WRONG PLAN.** Its
  disabled condition is `busy || plan.groups.length === 0 || !existingRows` — it checks `existingRows`
  is populated, never whether that population is still fresh. During the contaminated window `busy` is
  false and `existingRows` IS populated (with GHO's stale rows), so the button was live the entire time,
  and clicking it would have written the hybrid plan for real: groups named with MHO's prefix over GHO's
  department/unit codes, "already exists" verdicts pointing at GHO's actual groups.
- **Fixed with a ticket counter (`loadSeq`, a `useRef`).** Every `loadForSegment` call takes the next
  ticket; a response is applied to state only if its ticket is still the most recently issued one when
  it resolves — a superseded (slower, older) response is silently discarded rather than allowed to
  overwrite a newer selection's correct data. `setLoading(false)` in `finally` is gated the same way, so
  a late GHO resolution cannot toggle loading state after MHO's own load already settled it.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.257.0`.
  **Not yet re-tested live** — deploy, deliberately switch segments fast a few times, and confirm the
  preview never shows a name/match mismatch between two segments, even transiently.

## THE MIGRATE SCREEN'S SEGMENT PICKER STILL LOOKED CHOOSABLE EVEN WHEN A FLOW HAD ALREADY PICKED IT (2026-08-26, 1.0.256.0)
Client, on the Retire flow's step 1 ("Move the documents out"), which correctly pre-selected "Group
Head Office" via `initialSegmentKey`: *"You haven't remove the select... just put a message indicator
is enough."* Right — pre-selecting a `<select>` still leaves it looking editable, which reads as being
asked to choose something the flow already decided, even though the value shown is correct.
- **Gated on `initialSegmentKey` being present, not on which flow this is.** `SubtreeMigrator` mounts
  identically from TWO guided flows — Retire's "Move the documents out" and "Change the folder
  structure"'s migrate step — plus the standalone "All tools" Migrate tab, which has no flow feeding it
  a segment and genuinely needs the interactive picker. Gating on the flow id would have fixed one and
  left the other with the same complaint the moment someone tested it; gating on whether a segment was
  already supplied fixes both guided-flow cases in one change and leaves the standalone tab untouched.
- **⚠ `s.h` DOES NOT EXIST IN THIS FILE'S STYLE MAP — caught before shipping, not after.** The first
  draft referenced `style={s.h}`, which is `undefined` in `SubtreeMigrator.tsx`'s own
  `Record<string, CSSProperties>` — the exact "key that does not exist yields undefined, build stays
  green, element renders unstyled" trap this codebase has hit before in this same file's family.
  Reused `s.label` instead — the actual style the "Business segment" text it replaces was using.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.256.0`.
  **Not yet re-tested live** — deploy, open Retire's step 1 and confirm it now reads a plain
  "Segment: Group Head Office" / "Minamas Head Office" line with no dropdown, then confirm the
  standalone Migrate tab (via "All tools" → Move existing folders) still shows the real picker.

## RETIRE FLOW NO LONGER SHOWS THE CREATE-A-SEGMENT FORM (2026-08-26, 1.0.255.0)
Client, having just retired GHO: *"Retiring a segment is just to delete, not to create."* Right —
step 2 of Retire ("Segments -> Delete") mounts the exact same `SegmentCreator` screen the "Add a new
segment" flow uses, and until now that meant the full create form (Segment name, Appears under, Term
set ID, Top folder name, Levels…) sat below the segments list on EVERY mount, including this one,
where none of it applies.
- **Mirrors the existing `allowDelete`/`hideSegmentDelete` pair, inverted.** That pair already hides
  the DELETE button for the create-focused flow, for the identical reason stated in its own comment:
  offering the destructive half beside a creation form with no guard is the wrong default. The missing
  half — hiding the CREATE form for the delete-focused flow — simply had no prop to do it with.
- **`allowCreate` on `SegmentCreator`, `hideSegmentCreate` threaded through `IFolderManagerProps` →
  `FolderManager` → `FolderAdmin`**, set `true` only when `active.id === "retire"`. Defaults to
  shown, so the standalone Segments tab and "Add a new segment" are both untouched.
- **The segments list and its Delete buttons ALWAYS stay**, whatever the flag — seeing what already
  exists, and being able to delete it, is the entire point of this screen for the Retire flow.
- **The wrapped block needed a Fragment, not a bare parenthesis** — the create form is a sequence of
  SIBLING cards (warning box, name, appears-under, term set, top folder, permissioned levels,
  below-Unit levels, the Create button row), not one wrapping element, so `{allowCreate !== false &&
  (...)}}` needs `<>...</>` inside it or TypeScript refuses with "JSX expressions must have one parent
  element."
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.255.0`.
  **Not yet re-tested live** — deploy and open the Retire flow's delete step, confirm only the
  segments list and Delete buttons show, no create form beneath it; then open "Add a new segment" and
  confirm it is unchanged.

## SEGMENT DELETE NOW SHOWS A PROGRESS COUNT — WAS A STATIC "DELETING…" LABEL (2026-08-26, 1.0.254.0)
Client, watching GHO's retire run: *"it would be great if there is a progress bar."* Fair — deleting a
segment is one HTTP DELETE per Group Map row, sequential (no bulk-delete endpoint), and GHO alone was
**1157 mappings**. That can run for minutes with the button reading a static "Deleting…" the whole
time, which is indistinguishable from a hung request.
- **`delProgress` in `SegmentCreator.tsx`** tracks `{done, total}` across the row-deletion loop only —
  the mode-row delete and the optional folder-recycle step that follow are both fast and already have
  their own line in the result log, so they do not need a bar.
- **The button label itself becomes the count** (`Deleting… (342 of 1157)`), plus a thin visual bar and
  a line naming what is actually happening ("Removing folder-access mappings — … This is one request
  per row; do not close this tab until it finishes"), following this codebase's rule that the reason
  sits BESIDE the control, not only in a tooltip.
- **`undefined` before any row count is known** (segment has 0 mappings, or the count read failed) —
  the bar simply does not render rather than showing a division-by-zero or a stuck 0%.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.254.0`.
  **Not yet re-tested live** — deploy and retire a segment with a real mapping count to confirm the bar
  advances and the button's count matches the eventual result log.

## "MOVE THE DOCUMENTS OUT" HAS NO TOOL BEHIND IT, AND THAT IS DELIBERATE (2026-08-26)
Found testing the Retire flow's step 1 on GHO: the hint says "Move anything worth keeping somewhere
else first," and the screen it mounts is `SubtreeMigrator` — the SAME tool "Change the folder
structure" uses. That tool moves documents to match a NEW CHAIN SHAPE **within one segment**; it has
no concept of moving documents OUT to a DIFFERENT segment.
- **THERE IS NO CROSS-SEGMENT MOVE TOOL ANYWHERE IN THIS SYSTEM.** If a segment being retired has
  documents worth keeping, the admin's only route is manual — drag files into a different segment's
  folder in SharePoint directly, or re-upload through the form under a different segment. Step 1's
  hint text IS the mechanism; it is not a placeholder for a feature that exists elsewhere on the
  screen.
- **The mount is not wasted, though its purpose is narrower than the step's label suggests**: the
  "Check for existing files" button can still surface files sitting in an OLD folder shape within GHO
  itself (left over from some earlier structure change), which is worth knowing about even though it
  cannot move them anywhere outside the segment.
- **GHO had 0 documents when retired (2026-08-26)**, so this never had to be exercised for real. If a
  future retirement involves a segment WITH documents worth keeping, tell the client plainly: this
  step will not move them anywhere — they must be relocated by hand, in SharePoint, before opting into
  folder deletion on step 2.

## LEAVING A FLOW WHILE DIRTY BLED A STALE WARNING INTO A DIFFERENT FLOW ENTIRELY (2026-08-26, 1.0.253.0)
Found live, immediately after the fix above shipped: an unsaved abbreviation edit on GHO ("GCA" →
"GCAa", never saved) was left behind by clicking **"< Back to Folder Management"**, and the admin then
opened a completely unrelated flow — **Retire a segment**. That screen's Next button was held with the
message *"Save the abbreviation changes first…"*, about a screen the admin was no longer looking at,
inside a flow that has nothing to do with abbreviations.
- **THE ROOT CAUSE: `openFlow` and `leaveFlow` never reset `structureDirty`/`abbreviationsDirty`.**
  Both flags are set by `onStructureDirtyChange`/`onAbbreviationsDirtyChange` while their screen is
  mounted, and nothing clears them when the ADMIN leaves without saving or discarding — they persist in
  `FolderAdmin`'s own state across a full flow switch, because the state lives in the component that
  hosts every flow, not in the one that was dirtied.
- **THIS WAS ALREADY TRUE OF `structureDirty` SINCE THE STEP-LEVEL FIX SHIPPED, and 1.0.252.0's new
  `abbreviationsDirty` simply gave the same latent defect a second door.** Both flags were already
  correctly held WITHIN one flow (Next stays blocked at every later step, which is right); neither was
  ever checked at the one place that lets an admin leave a flow entirely: the **Back band** and
  **Finish**, both of which call `leaveFlow` unconditionally.
- **NOT a case for clearing the flag on leave.** Resetting `structureDirty`/`abbreviationsDirty` back to
  false the moment someone navigates away would silently let them walk off from the very edit the flag
  exists to protect — exactly the failure the whole mechanism was built to prevent, just moved to a
  different exit. The right fix is the same rule this codebase already applies to `FolderManager`'s own
  internal tab bar: **a tab switch with unsaved changes is REFUSED, not confirmed.**
- **Fixed by extending that exact rule to the flow-level exits.** `BackBand` already had a `disabled`
  prop (added for holding it during a bulk group run) — its call site now also disables on
  `structureDirty || abbreviationsDirty`, with its own reason printed beside it: *"Finish or clear what
  you are editing first — leaving now would lose it, and the guided flow you land on would open holding
  a warning that belongs to this screen, not to it."* **Finish** (only reachable on a flow's last step)
  is gated the same way, for the case where a flow's last step happens to be a dirty screen and Next was
  never the exit tried.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.253.0`.
  **Not yet re-tested live** — deploy, dirty the Abbreviations (or Structure) screen, try the Back band
  and confirm it's held with the reason shown, save or discard, confirm it releases.

## "RENAME OR RE-CODE A FOLDER" IS VERIFIED LIVE, END TO END (2026-08-26)
First full run of this flow by a person following the rail, on ClarenceDMSTesting/MHO. Renamed the
term `Group Compliance` → `Group Compliance TEST` in Term Store Manager (in place, never deleted),
typed that name into step 1 (confirmed it auto-focuses the matching row on step 2, out of 20
departments), changed its abbreviation `Group Compliance` → `GCT` on step 2, saved, ran Folder
Reconciliation scoped to MHO on step 3. Result, confirmed in the library view: the live folder
renamed to **`GCT`**, and its **`Full Name`** column reads **"Group Compliance TEST"** — both halves
of the mechanism (abbreviation → folder name, term label → Full Name) landed correctly in one run.
- **Found and fixed getting here, in order:** the Abbreviations step never inherited the flow's
  already-picked segment (two attempts — 1.0.250.0 didn't work, 1.0.251.0 fixed the real cause, a race
  against an unconditional auto-select-first-segment effect); and an unsaved abbreviation edit could be
  walked past into reconciliation with no warning (1.0.252.0, raised by the client's own question
  about exactly this). See the sections immediately above for both.
- **The unsaved-edit gate (1.0.252.0) was not itself exercised in this run** — the edit was saved
  before advancing, so the new hold never had occasion to fire. Confirm it separately: type an edit,
  do not save, confirm Next is held with its reason, save, confirm it clears.
- **Not yet reverted**: the term is still named "Group Compliance TEST" and the folder is still
  `GCT`. Rename back the same way (term store, then the abbreviation, then reconcile again) before
  relying on this segment for anything client-facing, unless it's being kept deliberately as a demo
  example.

## AN UNSAVED ABBREVIATION EDIT COULD BE WALKED PAST INTO RECONCILIATION (2026-08-26, 1.0.252.0)
Client's own question, testing the rename flow: *"Wouldn't that mean someone can actually [not] need
to save and I can click to next step without saving? Which in terms will bring client asking why it
did not work?"* — right on both counts, and a real gap, not a hypothetical.
- **`structureDirty` ALREADY GATES NEXT ON THE STRUCTURE/LEVELS STEP, and the identical protection was
  never extended to the sibling Abbreviations step.** `StructureManager`'s `onDirtyChange` is forwarded
  up via `onStructureDirtyChange` and holds the guided flow's Next button with its own message ("Save
  the structure first…"). `AbbreviationManager`'s `onDirtyChange` was only ever consumed LOCALLY inside
  `FolderManager` — for the "All tools" tab bar's own switch-tabs guard — and never forwarded to
  `FolderAdmin` at all. So typing a new code, not saving, and clicking Next was silently allowed.
- **The failure is invisible at the moment it happens.** Reconciliation reads the SAVED rows in `CRS
  Term Abbreviation`, so a run against an unsaved edit finds nothing to rename for that term, completes
  cleanly, and reports success. The admin believes the rename happened; the client is the one who
  finds out it did not, days or weeks later, asking why a folder still has its old name.
- **Fixed by copying the exact same pattern**: `onAbbreviationsDirtyChange` added to
  `IFolderManagerProps`, forwarded from `AbbreviationManager`'s existing `onDirtyChange` callback
  inside `FolderManager`, held in an `abbreviationsDirty` state in `FolderAdmin`, and added to the
  Next button's disabled condition and hint message alongside `structureDirty` — same shape, same
  wording style, second instance of the same protection.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.252.0`.
  **Not yet re-tested live** — deploy, type an abbreviation edit, do NOT save, confirm Next is now
  held with the reason shown beside it, save, confirm it clears.

## THE ABBREVIATIONS STEP ASKED THE SEGMENT AGAIN — TWO FIXES NEEDED, NOT ONE (2026-08-26, fixed in 1.0.251.0)
Found live, testing "Rename or re-code a folder" before the client demo. The flow's own header read
**"Subject: Group Head Office TEST"** — the segment was already picked — and step 2 (CRS Term
Abbreviations) still opened on a blank **"Select a segment…"**, forcing a second pick of the same
thing.
- **THIS IS THE EXACT COMPLAINT ALREADY MADE ONCE, for a different screen.** `IFolderManagerProps.ts`
  already carries `migrateInitialSegmentKey`, added 2026-08-20 for precisely this — the client's own
  words are in the comment: *"After selecting I got to select again which is weird."* That fix pre-
  selects the **Migrate** tab's segment picker from the flow's own `segKey`. It was never extended to
  the **Abbreviations** tab, which has no equivalent prop at all — `AbbreviationManagerProps` carried
  no `initialSegmentKey`, unlike `SubtreeMigrator`, which does.
- **Fixed by copying the exact same pattern**, not inventing a new one: `AbbreviationManager` gained
  `initialSegmentKey?: string` and the identical guarded effect `SubtreeMigrator` already uses —
  fires once (`chosen === ""`), never overrides a manual pick, and only applies once the key is
  confirmed to be a real option in the loaded segment list (a stale or mistyped key must never
  silently select nothing, and must never fight a choice already made). Threaded through
  `FolderManager` → `FolderAdmin` as `abbreviationsInitialSegmentKey={segment?.key}`, alongside the
  existing `migrateInitialSegmentKey={segment?.key}` — same source value, two named props, because
  each mount is deliberately named for its own screen rather than one generic prop reused everywhere.
- **This affects every flow that carries `needsSegment: true` and an ABBREVIATIONS step** — Add a new
  segment, Add a department or unit, Change the folder structure, and Rename or re-code a folder —
  not just the one it was found on.
- **⚠ 1.0.250.0 SHIPPED THAT FIX AND IT DID NOTHING — verified live, including after a full
  logout/login (ruling out the stale-tab/soft-navigation explanation entirely).** The real cause was
  a SECOND, pre-existing effect in `AbbreviationManager` that unconditionally auto-selects the FIRST
  segment in the list the moment it loads (`if (opts.length > 0) setChosen(opts[0].key)`), firing in
  the SAME pass as `setSegments`. My new effect's guard — `if (!initialSegmentKey || chosen !== "")
  return;` — was meant to mean "don't override a manual pick," but by the time it ran, `chosen` was
  ALREADY non-empty because the auto-select-first effect had just set it. The guard could not tell
  "the admin already chose one" apart from "this component already picked one FOR them a moment
  ago," and always lost that race.
- **FIXED PROPERLY in 1.0.251.0** by removing the separate effect entirely and folding the decision
  into the SAME pass that builds `opts` and calls `setSegments`: prefer `initialSegmentKey` when it
  matches a real option, fall back to `opts[0]` only when it doesn't. One decision, one place, no
  race — the same shape of fix as the audit request-activity flow's stray `For each` two days
  earlier: the individual pieces were each correct in isolation, and the defect was structural,
  between two mechanisms touching the same state.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.251.0`.
  **Still not re-tested live as of this writing** — deploy, redo the same flow, and confirm the
  segment actually carries over onto step 2 without re-asking. Given 1.0.250.0 was documented as
  fixed and was not, do not treat this entry as done until it is re-confirmed on screen.

## BULK UPLOAD CAN NO LONGER OVERWRITE ANYTHING (2026-08-27, 1.0.271.0/272.0)
The last `overwrite=true` in the system is gone. `BulkUpload.tsx` asked *"We noticed there's a same
name file… Are you sure you want to override this file?"* and, on **Yes**, uploaded with
`overwrite=true` — the ONLY path anywhere that could destroy a pending document.
- **TWO THINGS WERE WRONG WITH IT, and the second is why it was dangerous rather than merely
  permissive.** It destroyed, with no undo; and it asked **per file**, blocking a 50-file import on a
  modal. An admin importing historical documents is exactly the person who clicks Yes through the
  fifth prompt without reading it, and every Yes was somebody's upload gone. The old comment defended
  the prompt on the grounds that *"a pending file has been approved by nobody, so replacing one is
  recoverable"* — it is not recoverable, and the person who uploaded it is not the person clicking.
- **Now it SKIPS and offers a free name, decided ONCE after the run** — the client's choice, and the
  same model as the upload form, so an admin who has met one dialog recognises the other. `overwrite`
  is hardcoded `false`.
- **`nextAvailableName` is IMPORTED from `uploadBatches.ts`, not reimplemented.** One definition of
  what a free name looks like, so the two screens can never disagree about it.
- **⚠ THE RETRY IS KEYED ON THE `File` OBJECT, never an index or a name.** Two independent reasons, and
  either alone would break it: successful files leave the selection, which shifts every index; and this
  screen keeps ORIGINAL filenames, so two identically-named files in one selection would collapse to
  one match. `renameOverrides` is a `Map<File, string>`.
- **⚠ AND THE POST-RUN `setPicked` FILTER HAD TO CHANGE WITH IT.** It matched results to the selection
  BY INDEX — correct while every run sent the whole selection, and wrong the moment a retry sends only
  the clashing files, because an index into that shorter result array means nothing to `picked`. Now
  matched by File identity, which is right for both cases. The old comment explicitly defended index
  over name; identity beats both.
- **The clash check LISTS the folder instead of probing one name** — the same single request the
  `Exists` probe cost, and it yields the names a suggestion needs.
- **Proceed passes the file list EXPLICITLY**, because the successful files were filtered out of
  `picked` moments earlier and that setState has not landed — reading state there would re-send
  everything that just worked. Same trap, same fix, as the upload form's `handleUpload(override)`.
- **A suggestion that has itself been taken converges**: the forced name still goes through the clash
  check, so it is skipped and offered the next number.
- **Unchanged, deliberately: the APPROVED-side check still refuses outright** with no rename offer. An
  admin must not be one click from filing a second copy of an already-approved record — and unlike the
  upload form's uploader, an admin can go and look.
- **⚠ THE INVISIBLE-PENDING CASE HAD TO BE CARRIED OVER TOO, AND IT IS THE COMMON CASE HERE
  (1.0.272.0).** The Add-failure clash detection was added to `Form.tsx` an hour earlier and NOT to
  this screen — caught only because the client said they would test as an **uploader-only** account.
  - **Worse here than in the form.** Bulk Upload has been open to ordinary uploaders since 2026-08-22,
    and Draft Item Security is *"Only users who can approve items (and the author)"* — so an uploader's
    folder listing contains **only their own** pending files. Every pending file belonging to anybody
    else is invisible to the pre-check, and only `Files/Add` knows. For a PIC importing into a shared
    unit folder that is the NORMAL situation, not an edge case.
  - Without it the uploader gets SharePoint's raw *"A later version of this item has already been
    modified…"* — versioning explained to someone who only chose a filename — and no rename offered for
    the exact case the offer exists for.
  - Same loose matching as the form (`later version` / `already been modified` / `already exists` /
    409), same reasoning: a false positive only offers a rename that can be declined, a false negative
    hides the fix behind a 500. Same guarantee of a suggestion, by including the known-taken name when
    the listing came back trimmed to nothing.
  - **THE GENERAL LESSON, and this is the second time today:** a fix applied to one upload path is not
    applied to the other. `Form.tsx` and `BulkUpload.tsx` share `nextAvailableName` but not their write
    paths, so every clash-handling change has to be made twice — deliberately, since their pre-checks
    differ (the form composes a name and probes one file; this screen uses the file's own name and
    lists the folder). Grep the sibling before calling a clash fix done.
- **Verified**: `tsc --noEmit` clean, `heft build` clean, full test suite green, no `overwrite=true` or
  `askReplace`/`answerReplace` left in code (only in the comments recording their removal). Packaged
  as `1.0.272.0`. **Not yet re-tested live** — and the test that matters is as an UPLOADER-ONLY
  account against a folder holding somebody else's pending file, since that is the path the pre-check
  cannot see.

## THE PER-ROW DELETE IS GONE FROM BULK UPLOAD, AND THE CLASH DIALOG LISTS EVERY REFUSAL (2026-08-27, 1.0.274.0)
Two client reports from one live Bulk Upload run, plus one correction of mine.
- **⚠ THE `✕` ON A COMPLETED ROW DELETED THE UPLOADED DOCUMENT, AND IT IS REMOVED** (client: *"pls
  remove the delete functionality … don't let it be there since PIC cannot delete the file until they
  request"*). They prefaced it *"if I am not wrong"* — they were right, and it was worse than they
  thought. **FOUR faults, and only the first is the one they could see:**
  1. **BULK UPLOAD STOPPED BEING ADMIN-ONLY ON 2026-08-22.** Every uploader reaches this screen, and a
     PIC holds no delete anywhere — they lost `DELS` on 2026-08-20 and must raise a REQUEST a Head of
     Unit decides. A delete button on their own screen contradicts the entire workflow.
  2. **IT PURGED RATHER THAN RECYCLED.** A plain `X-HTTP-Method: DELETE` does NOT reach the recycle
     bin, while the approved deletion-request flow deliberately calls `recycle()` — *that* 93-day
     restorability is what makes approving one reasonable. So this was **strictly more destructive
     than the sanctioned route, with no approval in front of it.**
  3. **IT REPORTED SUCCESS AFTER DELETING NOTHING.** It treated a 404 as success and its `sru` points
     into the APPROVAL library, which auto-approve plus Auto-route empties within about a minute. So
     clicking it after the file routed marked the row `Deleted` and toasted *"deleted from
     Documents"* over an untouched document. **A silent lie on a destructive control.**
  4. **Every string said "Documents"**, stale since the 2026-08-22 repoint.
  `sru` is left captured and inert; the `deleted` FileState, its label and its CSS are removed. A long
  comment records the four faults so nobody re-adds it — if a delete is ever wanted here again it must
  **recycle**, be gated on the **persona** rather than the row's state, and re-resolve the file.
- **THE CLASH DIALOG LISTED ONE OF TWO REFUSALS** (client: *"the popup doesn't show two files like a
  normal upload form, it shows only one for bulk upload"*). It was filtered on `suggestedName`, so the
  approved-side refusal — which deliberately had none — never reached it and survived only in the
  results list below. **THIRD INSTANCE TODAY of the standing rule: a fix applied to one upload path is
  not applied to the other.** `Form.tsx` had this exact defect fixed at 1.0.269.0 and the sibling was
  not touched. **Grep the other screen before calling any clash fix done.**
  - `FileResult.nameClash` is now what routes a skip into the dialog, kept SEPARATE from
    `suggestedName` so a clash with **no** free name still appears. A dialog that accounts for SOME of
    the failures is worse than one accounting for none, because it looks complete.
- **AND THE APPROVED SIDE IS OFFERED A RENAME HERE TOO, reversing a deliberate decision from
  1.0.270.0's sibling section.** That refusal argued nobody should be one click from filing a second
  copy of an approved record. **THE ARGUMENT IT MISSED IS THAT AUTO-ROUTE RENAMES IT ANYWAY:** this
  screen writes into the approval library and stamps `BulkImport`, so auto-approve fires and Auto-route
  copies with `nameConflictBehavior: 2`, producing a SharePoint-renamed `TEST1.pdf` at routing time and
  telling nobody. **A visible rename beats an invisible one**, which is the same reasoning that settled
  it for the upload form.
  - **⚠ THE SUGGESTION MUST CLEAR BOTH LIBRARIES**, or a name free only in the approval library just
    moves the collision to Auto-route. `probeApprovedSideFile` became `approvedSideClashInfo`, LISTING
    the folder instead of probing one name — the same single request, and it yields the names.
  - **Offered AND still flagged**: the offer row says *"The original has already been approved and
    filed — check this is not the same document."* A renamed copy of an approved document is a SECOND
    document, which is a different thing from a first one.
  - Proceed renders only when something is renameable; Cancel becomes **Close** when nothing is.
- **⚠ AND MY 1.0.273.0 FRAMING WAS WRONG — corrected by the client.** I called the `(2) (2)` in their
  screenshot a defect; they checked their own machine and found they had genuinely uploaded a file
  named `… (2).pdf`, so the output was **correct for the input**. The increment is still the better
  behaviour and stays, and it does fix a REAL stacking bug on the convergence path (a suggestion fed
  back in as the desired name produced `(2) (2)` instead of `(3)`) — but that is not what the
  screenshot showed. **Ask what the input actually was before calling the output a bug.**
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green (46 suites,
  0 failed). Packaged as `1.0.274.0`. **Not yet re-tested live** — confirm the `✕` is gone from
  completed rows, and that a run clashing on BOTH sides lists two files in the popup.

## ⏸ PARKED, AND URGENT: A ROUTED DOCUMENT WAS MOVED TO THE `Documents` ROOT (2026-08-27)
Parked on the client's call the night before SDG migration — **not resolved.** Run history keeps 28
days, so the evidence is intact until ~2026-09-24.

**One file, `TEST - Clarence - CHO - 27-08-26.pdf`, ended up at `/Shared Documents/` with no folder.**
⚠ **A file at the library root has NO unit ACL** — reconciliation breaks inheritance per DEPARTMENT
folder, not at library level, and the site-entry group holds Read on the root for the browse corridor.
So it was readable by everyone with site access. That is why this matters beyond one test document.

- **⚠ WHAT IS ESTABLISHED, FROM THE RUN ITSELF — THE FLOW'S PATH LOGIC IS CORRECT.**
  - `Compose - 1` output: `MHO/CEO Office Administration/General Administration ＆ Building Management/2024/Agreement` — right.
  - `Copy file` **Inputs**: `destinationFolderPath` = `/Shared Documents/MHO/CEO Office Administration/General Administration ＆ Building Management/2024/Agreement`, `nameConflictBehavior: 1` — right.
  - `Copy file` **Outputs**: `ItemId 6828`, `Path` = that same correct folder, `Size 19716`, at **14:56:45Z**.
  - SharePoint then reports item **6828** at `/Shared Documents/TEST - Clarence - CHO - 27-08-26.pdf`,
    `Modified` **14:56:58Z**.
  - **So something moved it 13 seconds after the copy, INSIDE the same run** (started 14:56:42Z, ran 17s).
- **⚠ TWO THEORIES ARE DEAD — do not re-propose either.**
  - **The fullwidth `＆` is INNOCENT.** It survived the split and the copy intact, proven by the inputs
    and outputs above. (It was proposed twice and killed twice.)
  - **Not systemic.** Exactly ONE file sits at the root, and items 6829/6830/6831 — consecutive ids,
    same destination folder, same minute — landed correctly.
- **NEXT STEP: the actions AFTER `Copy file` in that run.** `GetSourceUniqueId` is ruled out (a read).
  **Prime suspect is the `validateUpdateListItem` stamp** — `FileLeafRef` IS settable through that
  endpoint, which is how a rename is done, so a malformed value there could reposition a file. Get the
  full `formValues` payload. Start by listing the run's action names rather than pasting each one.
- **⚠ MIGRATION IMPLICATION, and this is the part that matters tomorrow:** the cause is unknown, so it
  can recur on SDG with REAL documents. **After the SDG migration, list the `Documents` and
  `HC Documents` ROOTS** and confirm nothing is sitting there:
  `/_api/web/GetFolderByServerRelativeUrl(@f)/Files?$select=Name,TimeCreated&@f='/sites/CRS/Shared Documents'`
- **The routing `Created` bug is now PROVEN, not inferred** (it was previously recorded as suspected):
  source `Created` `14:56:17Z`, copy `Created` `06:56:00Z` — exactly 8 hours early, seconds collapsed to
  `:00` because the flow writes `M/d/yyyy h:mm tt`. That pair is Auto-route's fingerprint on any file.
- **Third independent proof that `UniqueId` does not survive routing:** source `f7d2d18c-…`, copy
  `B04E2544-…`. Same document, two GUIDs.

## A REPLACED SUBMISSION READS `Cancelled`, NOT `Deleted` (2026-08-28)
Client: *"The older submission under My Submission will change to Cancelled status if replaced."* The
word was queried — "Superseded" or "Replaced" says what happened, where "Cancelled" normally means
somebody withdrew something — and they confirmed it: *"I know its weird but client want it to be
Cancelled."* **BUILT, NOT DEPLOYED, NOT SITE-TESTED.**
- **⚠ THE PROBLEM IS THAT A REPLACEMENT LOOKS EXACTLY LIKE A DELETION TO THE RECORD.** Overwriting a
  file destroys the old `SubmissionFileId` stamp — the new upload writes its own — so the displaced
  record stops resolving and `mergeRecords` called it **deleted**, telling the original uploader their
  document was destroyed when somebody had in fact filed a newer version of it. The one false
  positive this whole feature exists to prevent, arriving through a door it did not cover.
- **`RecordState` gains `cancelled`, and it is NOT a kind of `deleted`.** `deleted` is DERIVED — a
  record that failed to resolve, judged against whether the live read was complete. `cancelled` is
  **OBSERVED**: we watched the overwrite succeed and wrote `ReplacedAt`/`ReplacedBy` on the record. So
  it holds whether or not the libraries could be read this time, and a failed read can never produce
  it.
- **⚠ THE PRECEDENCE IS LOAD-BEARING: `live` > `cancelled` > `deleted`/`unknown`.** A record that
  STILL RESOLVES describes a document that exists, so calling it cancelled would claim a loss that did
  not happen — which is exactly the state left behind if a replacement's re-stamp failed. And
  `cancelled` beats `unknown` because a transient read failure must not downgrade a known replacement
  to "we are not sure", which is *less* true than what we already know.
- **TWO COLUMNS, NOT A YES/NO FLAG** — `ReplacedAt` (DateTime, **ISO**) and `ReplacedBy` (Text). "Who
  filed the newer version, and when" is the entire reason this state is worth having over `Deleted`; a
  boolean records the first half and throws away the half the uploader actually needs. Shown on the
  file card, so they can go and find the replacement.
- **⚠ THE READ RETRIES WITHOUT THEM ON HTTP 400, AND THAT IS NOT OPTIONAL.** One unknown field name
  fails the WHOLE `$select` (gotcha #11), so on any site whose `CRS Submissions` predates this,
  asking unconditionally would take **every record off My Submissions** — not merely the replacement
  state. `RECORD_READ_SELECT_LEGACY` is derived by SUBTRACTION from the full list, never written out
  again, so the two cannot drift. **400 only**: a 404 is the list missing and a 403 is the ACL pass not
  having run, and retrying either asks the same unanswerable question twice.
- **⚠ TWO REPLACEMENT PATHS, AND THEY ARE CAUGHT IN DIFFERENT PLACES.**
  1. **The upload form's staging replace** — `Form.tsx` reads the doomed file's `SubmissionFileId`
     **BEFORE** `Files/Add(overwrite=true)`. There is no second chance: the Add replaces the file and
     this upload stamps its own id over the columns, so read it afterwards and there is nothing left.
  2. **An approval that replaces a filed document** — `ApprovalDocument.tsx`, in the confirmed-clash
     branch of the existing warning. **Marked BEFORE the replacement happens**, which is safe ONLY
     because of the precedence above: Auto-route does the copy minutes later, so until it runs the old
     file still resolves and reads `live`; the moment it is replaced the row turns Cancelled. It
     self-corrects, and an approval Auto-route never completes leaves nothing wrongly marked.
     **Confirmed clashes only** — the other branch of that warning is "the check could not be
     answered", and marking on that would assert a replacement that may never happen.
- **Bulk Upload needs nothing**: its `overwrite` is hardcoded `false` (1.0.271.0), so it never
  displaces anything directly.
- **⚠ `markRecordReplaced` IS NOT FILTERED BY AUTHOR**, because the replacer is usually not the person
  who uploaded the file — so it edits somebody else's row. Only possible because the reconciliation ACL
  pass grants the uploader/approver/HoD groups `CRS Request` (Add + Edit + View, no Delete) on this
  list. The coarse-grant exposure note from the `CRS Requests` design applies here too.
- **Never throws, never blocks.** The replacement has already succeeded by the time it runs; a failure
  costs the LABEL and nothing else, and the record falls back to reading `deleted` — yesterday's
  behaviour, which is always the correct way to degrade. Not gated on the new record write succeeding
  either: the displaced record can exist on a site where this upload's own row could not be written.
- **⚠ AN UNPARSEABLE `ReplacedAt` IS `undefined`, NEVER `Invalid Date`** — and here that decides
  behaviour rather than formatting, because `replacedAt !== undefined` is what makes a record
  cancelled and an `Invalid Date` is truthy.
- **Blue badge, deliberately not grey or amber.** Grey is `Deleted` (a loss), amber is `Not checked` (a
  doubt); sharing either would put this state straight back into the one it exists to be told apart
  from.
- **Migration = re-run reconciliation** (it asserts `RECORD_COLUMNS` on `CRS Submissions`). Until then
  the read falls back and every replacement keeps reading `Deleted` — degraded, never broken.

## A SUBMISSION RECORD THAT SURVIVES DELETION — BUILT AND SITE-VERIFIED (2026-08-27, 1.0.310.0)
Spec: `docs/superpowers/specs/2026-08-27-submission-record-design.md`. All six steps built.
**Inert until reconciliation runs**: no list, no columns, no stamp — so an upload writes no record and
the page behaves exactly as it did before. That degradation is the design, not a gap.

- **⚠⚠ THE AGREED KEY WAS `UniqueId` AND IT IS WRONG — IT WOULD HAVE REPORTED EVERY APPROVED FILE AS
  DELETED.** Auto-route is copy-stamp-delete, so the routed copy in `Documents` carries a **NEW**
  `UniqueId` and the source holding the recorded one is deleted. This file already says so twice (the
  archive section, and why the archive mover must use `MoveTo`); the agreed design contradicted it.
  - **Nobody had hit it because `MySubmissions` reads `UniqueId` LIVE, per row, from whichever library
    the row came from** — it never has to survive anything. The comment at `MySubmissions.tsx:356`
    (*"a GUID that survives the rename or move that a path does not"*) is true in its own context and
    is exactly the sentence that makes the key look safe.
  - **THE JOIN IS A STAMPED PER-FILE COLUMN, `SubmissionFileId` (`SFI-<yyyymmdd>-<4>`).** A stamped
    column survives the copy — which is precisely why `SubmissionId`/`BatchId` had to exist in all four
    libraries (2026-08-22) — so the record is keyed on OUR data, not SharePoint's identity.
    `ItemUniqueId` is still stored as a secondary and **never decides Deleted**.
  - **A ROW IS WRITTEN ONLY WHEN THE FILE COULD BE STAMPED.** An unstampable upload recorded anyway
    would be a permanent false "Deleted"; degrading to yesterday's behaviour is always correct.
  - **Replace REASSIGNS the stamp** (the 2026-08-27 `nameConflictBehavior: 1` change): the copy
    overwrites the destination's columns, so the replacing file's stamp lands there and the REPLACED
    file's row stops resolving and reads Deleted. Defensible — the old content is a version now — but
    it is a behaviour to tell the client, not an accident. **Unverified**, as is whether Replace
    versions at all. Supersedes the earlier claim that *"Replace needs nothing extra."*
- **⚠ AN UNRESOLVED RECORD IS `deleted` ONLY IF THE LIBRARIES WERE ACTUALLY READ.** `RecordState` is
  THREE-valued (`live`/`deleted`/`unknown`) and `mergeRecords` takes `liveReadComplete`. A throttled
  read must never tell an uploader their documents were destroyed. Empty ≠ unknown, where the cost is
  somebody believing their work is gone.
- **⚠ RECONCILIATION CREATES THE SUBMISSIONS LIST — a deliberate departure from the agreed note**,
  which said to follow the `CRS Requests` page-provisioning pattern. There is no admin page that owns
  this list: the screen that READS it is My Submissions, which every uploader opens, so a Provision
  button there would be shown to the people who cannot press it. Admin-only, idempotent, and beside
  the pass that already asserts the reference columns. One fewer manual step at cutover.
- **ALL SIX STEPS BUILT (`tsc` clean, 1452 tests, 0 failures):**
  1. `LIST_SUFFIX.submissions` + `PRIMED_SUFFIXES`.
  2. `SUBMISSION_FILE_COLUMN` in `optionalColumns.ts`, asserted by reconciliation on every library.
     **Deliberately NOT part of `REF_COLUMNS`** — folding it in would make a library reconciled before
     today fail the check for the pair too, silently switching off grouping that works there now.
  3. `shared/submissionRecords.ts` + 50 tests (row shape, payload/parse, the join, the verdict);
     `Submission.submissionFileId`; `groupsForRequestLists`/`REQUEST_LIST_ROLES` + 9 tests.
  4. Reconciliation: list creation, column assertion, **`Author` indexed**, and the two-list ACL pass
     with the conditional revoke.
  5. Both writers, via `shared/spSubmissionRecords.ts` (the SPFx half; `writeSubmissionRecord` cannot
     throw and its result is ignored deliberately). **BulkUpload already stamped the pair** — the note
     that it was "deliberately NOT stamped" was stale.
  6. My Submissions reads the records, merges, and renders the Deleted state.
- **⚠⚠ TWO DEFECTS FOUND WHILE BUILDING, EITHER OF WHICH WOULD HAVE BROKEN THIS SILENTLY:**
  - **`readLibrary` NEVER ASKED FOR `SubmissionFileId`.** Live rows carried no stamp, the index was
    empty, and **every record would have been reported as DELETED** — the exact false positive the
    whole design exists to prevent, by a route the design did not cover. The read is a **three-rung
    ladder** now (all three refs → the pair → none), so a library missing only the newer column keeps
    its grouping, and **`stampMissingRef`** tells the merge the stamp was unavailable so it says *"not
    checked"* instead. **It is a REF, not state** — `load` must read it back in the same pass, and a
    `useState` value is invisible to a closure already running (the 1.0.237.0 trap).
  - **A SINGLE `liveReadComplete` BOOLEAN WOULD HAVE DISABLED THE FEATURE FOR MOST USERS.** This page
    reads up to six libraries and deliberately swallows HC and archive failures — an unreadable HC
    library is the NORMAL case for everyone without clearance. One global flag would be false for all
    of them, so nothing would ever read as deleted. `mergeRecords` takes a **predicate**; the caller
    answers **per chain** (a normal-library record needs the normal chain, an HC one the HC chain), and
    a throwing predicate reads as not-complete.
- **WHERE A GONE FILE APPEARS:** the **Submissions** view (inside its own batch — the client's actual
  ask) and **All**. **NEVER** on Pending/Approved/Rejected, because `filterByTab` matches `r.status`
  and a record row's `Pending` is a placeholder — listing it would say a destroyed file awaits
  approval. Never in the status tallies or batch counts either (`liveRowsOnly`).
- **A gone row has no name button, no `Open file`, no preview and NO REQUEST BUTTONS.** A share request
  would be approved by a Head of Unit and then fail in their own session, after the requester was told
  it was being handled. Its details come from the record's **snapshot**; no per-item read is attempted,
  because `FieldValuesAsText` would answer nothing for a deleted item.
- **✅ VERIFIED ON ClarenceDMSTesting 2026-08-27. Test plan:
  `docs/2026-08-27-recent-features-test-plan.md`**, which also covers the 1.0.304.0 handoff items and
  carries the verified log. Reconcile FIRST on any new site — the list, its columns, the `Author`
  index and the ACL pass all come from a run.
  - **✅ §2.3 — THE STAMP SURVIVES AUTO-ROUTE, AND THE AGREED DESIGN WOULD HAVE FAILED HERE.**
    `TEST - Clarence - CHO - 27-08-26.pdf` uploaded as a non-admin PIC (`chocheetuck4`,
    `GHO_GCA_EG_UPLOADER`), recorded as `SFI-20260827-BGWU` against `Approval Document`, approved,
    routed. The file is now `Documents` item **6793** — a NEW item id, proving a COPY rather than a
    move — carrying `SubmissionFileId = SFI-20260827-BGWU`. The record joins; the row reads `Approved`.
    - **On the agreed `UniqueId` key this exact file would have read `Deleted`.** The copy has a new
      `UniqueId` and the source holding the recorded one was deleted. No longer a theoretical
      objection — it is the observed behaviour of the first approved document through this feature.
  - **✅ §2.1** — three records written, `SubmissionRef`/`BatchRef`/`SubmissionFileId` correctly
    formatted, `LibraryTitle` resolved to the LIVE title (not the legacy `Staging`), `UploadedBy`
    lower-cased, `Source` = `Form`. Written from a non-admin uploader's own session.
  - **⚠ A FILE APPROVED BUT NOT YET ROUTED IS A VISIBLE STATE, AND IT LOOKS BROKEN.**
    `statusToDecision(0)` maps Approved, so such a file reads `Approved` in My Submissions **with the
    approval-library path** — and the moment Auto-route deletes the source, that path 404s: blank
    preview, *"No details were recorded for this file"*. Seen live 2026-08-27. Transient and
    self-correcting on reload, but an uploader can hit it, so **do not diagnose it as a record
    failure.**
  - **✅ THE ACL PASS'S GRANT AND REVOKE BOTH FIRED, first live execution anywhere.**
    `roleassignments/getbyprincipalid(1581)` answers `Can not find the principal` on **both**
    `CRS Requests` and `CRS Submissions`, and a non-admin PIC could still raise a deletion request
    afterwards — the proof that mattered. **SDG's run was cancelled half way**, so there the revoke
    never fired and only 13 groups are granted; re-run it to completion.
  - **⚠ AN EMPTY `CRS Submissions` READ IS NOT EVIDENCE OF FAILURE — CHECK THE CLOCK FIRST.** On
    2026-08-27 an empty result was read as the writer being broken; the query had run **four minutes
    before the upload it was meant to check**. The decisive evidence is the Network tab: a POST to
    `items` returning **201**.
- **⚠ `MySubmissions.tsx` IS NOW 2161 LINES**, over the 2000 ceiling — a new warning and the fifth file
  in the project to exceed it. Extracting part of it deserves its own reviewable change.

The agreed design as the client stated it, kept below because the reasoning still governs:

Client: *"if I deleted the file the my submission file list also disappears, can we ensure that doesn't
happen? … Client wants to be able to record that file even if someone delete or replace … I know if
someone deletes the file I cannot open file and show the content but that is fine, then just show this
specific file is deleted but what matters is its recorded for that submission."*

- **THE CAUSE IS STRUCTURAL: MY SUBMISSIONS HAS NO RECORD OF ITS OWN.** It reads the two document
  libraries `AuthorId eq me` and derives the submission grouping from the surviving rows, so a deleted
  file does not become "deleted" - it stops existing, and its batch quietly shrinks.
- **BUILD: a `<P> Submissions` list, one row per uploaded file**, written at upload time by **BOTH**
  upload screens (client's choice - Bulk Upload included). Row carries the submission reference, batch
  reference, filename, destination path, ~~the file's `UniqueId`~~ (**SUPERSEDED — the stamped
  `SubmissionFileId` is the key; `UniqueId` is a secondary that never decides Deleted, see above**),
  and a metadata snapshot.
- **⚠ STATUS IS JOINED LIVE FROM THE LIBRARIES, NEVER STORED ON THE ROW** - the client asked for this
  explicitly (*"make sure if the Approval status changes it will change as well"*). The list decides only
  that a row EXISTS; pending/approved/rejected and the routed path all come from the libraries on every
  load. A row whose ~~`UniqueId`~~ **stamped `SubmissionFileId`** is in NEITHER library renders
  **Deleted** - greyed, no preview, no Open file, name/path/batch/metadata intact.
- ~~**Replace needs nothing extra**: a replaced document keeps its identity and gains a version, so
  the row still resolves.~~ **❌ SUPERSEDED — Replace REASSIGNS the stamp onto the destination file, so
  the REPLACED file's row stops resolving and reads Deleted. See the correction above.**
- **TWO CHEAPER OPTIONS WERE RULED OUT, with reasons, so they are not re-proposed:**
  - **The recycle bin** keeps name and path for 93 days, but an ordinary user sees only items THEY
    deleted - so an approver's or admin's deletion is invisible to the uploader. It is also full of
    noise: Auto-route deletes the approval-library copy of every routed file.
  - **The audit log** already records `Uploaded` per file with the actor, but those rows carry no
    submission/batch reference, so a deleted file could not be put back in its batch - which is the
    whole requirement.
- **PRE-EXISTING FILES GET NO ROWS and keep behaving as they do now** (grouped by folder and date,
  vanishing if deleted). No retro-fix exists - nothing recorded them. Accepted by the client for the
  test site; **expect it at SDG cutover too.**
- **⚠⚠ PERMISSIONS: OPTION 2 WAS CHOSEN, AND IT COVERS BOTH LISTS.** The client's words:
  *"I think best to go to option 2, I don't want to get a lash back later on."*
  - **REJECTED: granting `CRS_SITE_MEMBERS` the `CRS Request` level**, which is what `CRS Requests`
    had **until this change** (the 1.0.292.0 pass). It works, and the client accepted it earlier for
    requests - but it lets ANY site member READ every submission (filenames, paths, metadata, who
    submitted what) and EDIT any row including someone else's. Delete is blocked, which is the core
    requirement, but the read exposure is what they do not want.
  - **BUILT: a reconciliation pass that scopes BOTH `CRS Requests` AND `CRS Submissions` to the
    `_UPLOADER` / `_APPROVER` / `_HOD` groups**, asserted every run so it cannot rot. Named in the
    2026-08-21 requests section as the right answer (~450 grants) and left unbuilt until now.
    **Both lists in one pass** - a tighter level on one list only would be the wrong shape.
    - **Derived from the ROLES** (`UPL, UPLHC, APR, APRHC, DEPTVIEW`) via `groupsForRequestLists`,
      **never the name suffixes** — old spellings, renames and `roleFromGroupName`'s MEMBER
      fall-through make a suffix rule silently wrong in both directions.
    - **⚠ AND IT REVOKES, which the agreed note did not mention.** Granting ~470 groups while leaving
      the site-wide grant in place changes the exposure not at all. Every binding for the site-entry
      principal comes off — not just `CRS Request`, because the break used
      `copyRoleAssignments=true` so that group also carries the web's Read.
    - **LEFT IN PLACE when nothing was derived, when any grant failed, or when the Group Map could
      not be read**, each with its own log line. A wrong removal means nobody can raise a request,
      found by an uploader; a delayed one means one more run of an exposure nobody is watching.
  - Extended the 1.0.292.0 requests-list pass rather than writing a second one.
- **⚠ IMPLEMENTATION TRAPS, all previously paid for:**
  - **`LIST_SUFFIX.submissions` MUST be added to `PRIMED_SUFFIXES`.** `LIST_SUFFIX.requests` was left
    out from the day requests was written until 2026-08-20 and the failure was TOTAL and SILENT - an
    unprimed suffix answers the legacy `DMS <suffix>` for ever, which 404s on a CRS site. There is a
    test pinning that the arrays match; keep it passing.
  - **The write must NEVER fail the upload.** Own try/catch, no row = degrade to today's behaviour.
    Same rule as the `SubmissionId`/`BatchId` stamp, which is conditional for exactly this reason.
  - **Provisioning**: ~~follow the `CRS Requests` pattern (created when an admin opens the page)~~
    **SUPERSEDED — reconciliation creates it; there is no admin page that owns this list. See above.**
    An uploader cannot create a list, so an upload before provisioning writes nothing - which must be
    silent, not an error.
  - Match rows to files on ~~**`UniqueId`**~~ **the stamped `SubmissionFileId`**, never name or path -
    both change, and `UniqueId` does not survive Auto-route.
- **ORDER TO BUILD:** naming + `PRIMED_SUFFIXES` → the list and its provisioning → the reconciliation
  ACL pass (both lists) → the two writers → the My Submissions merge and the Deleted state.
  **Steps 1-4 are done; 5 and 6 remain.**

## ✅ END-OF-DAY VERIFICATION, 2026-08-27 (through 1.0.293.0)
Everything below was confirmed on ClarenceDMSTesting with real accounts, not inferred.
- **`CRS Owners` IS STILL ON `CRS-Audit-Log.aspx`** — so no earlier run stripped it, and the
  `ownerGroupId` stale closure never actually caused a removal. That uncertainty is CLOSED.
- **A non-SCA in `CRS Owners` can now open the Audit Log** — `isSystemAdmin` verified live, which was
  the whole point of 1.0.286.0. The log is also recording correctly (the membership changes made during
  the test appear in it).
- **The group rename ran: 327 groups**, and the list now reads `..._APPROVER_HIGHLY_CONFIDENTIAL`,
  `..._UPLOADER_HIGHLY_CONFIDENTIAL`, `..._VIEWER`. Access unchanged, as designed.
- **DELETION AND SHARE REQUESTS WORK END TO END for an uploader**, after granting
  `CRS_SITE_MEMBERS` → `CRS Request` on the list by hand: the share dialog renders, the request
  submits, and the already-asked guard shows `deletion asked` with the buttons replaced by `asked`.
- **`Folder Map: no orphaned rows ✓`** — the prune ran, first time that day.
- **⚠ THE `MHO_Corporate Services…` HTTP 500 DID NOT RECUR** on a later full run. It was
  TRANSIENT — do not go looking for a cause. The diagnostic order in that section stands if it returns.
- **The three orphaned abbreviation rows (`Group Test Something`, `Unit1`, `Unit2`) were deleted by
  hand**, which is the documented resolution: those rows are never auto-deleted because they hold the
  only copy of an authored abbreviation.
- **⚠ NOT verified, because the site no longer reproduces it: the requests-list GRANT path in
  1.0.292.0.** The grant already existed by hand, so the run took the `already raise requests ✓`
  branch. **SDG will exercise the grant path** — watch that line on its first run there.
- **⚠ STILL UNVERIFIED AND THE MOST IMPORTANT ONE: whether Auto-route's Replace writes a VERSION
  rather than destroying the old content.** Versioning is on for `Documents` (major, 500 kept), so it
  should — but the client is relying on this and it has never been run. Send a document for approval
  as a replacement, approve it, then open Version History and expect **2.0** with the old content at 1.0.

## ⚠⚠ THE OWNERS GROUP ID WAS A STALE CLOSURE TOO - THE 1.0.237.0 BUG, MISSED (2026-08-27, 1.0.293.0)
A run reported `⚠ Administrator pages: site Owners group or "Full Control" not resolved — none
locked`. **Full Control was fine; `ownerGroupId` was `null`.** It is `useState`, filled ONLY by a
mount-time effect, and `runReconciliation` is one long async function reading it from the closure it
was created in - **so a run started before that fetch resolved held `null` for its entire length.**
Identical in shape to the 1.0.237.0 `roleDefs` defect, in the one place that fix did not reach.
- **⚠ THE REPORTED SYMPTOM WAS THE *SAFE* CONSEQUENCE OF THREE.** The admin-page lockdown refuses
  (fail-closed, correct, and the only one that says anything). The other two are silent:
  1. **The site-entry library pass SKIPS restoring Owners after BREAKING a library's inheritance** -
     leaving that library reachable only by site collection administrators. An owner who is not also
     one loses it.
  2. **`protectedIds` in the page pass comes back EMPTY**, and the removal loop's
     `if (ra.PrincipalId === ownerGroupId) continue;` never matches - so the pass that asserts a FULL
     ACL at page scope loses the one principal it exists never to remove. **It would strip site Owners
     from every page it asserts.**
- **⚠ THIS RUN GOT LUCKY: `Prune skipped — run had 1 error(s)` was the abbreviation orphans, and the
  page pass reported no removals.** Whether Owners was actually stripped from any page on an earlier
  affected run is NOT established - check a page's ACL if anything looks wrong.
- **FIXED THE SAME WAY 1.0.237.0 WAS**: `resolveOwnerGroupId` RETURNS the id as well as storing it, the
  run uses state when already populated and fetches otherwise, and **all seven in-run consumers now
  read the local `ownersId`** rather than the state. The `typeof ownerGroupId === "number"` guards that
  existed to survive the null are gone with it - they were treating the symptom.
- **AND THE RUN NOW REFUSES when it cannot be resolved**, the second fail-closed stop in this run after
  the empty-levels one, and for a stronger reason: an unresolvable Owners id can cause REMOVALS that
  the run cannot undo. A run that might strip Owners must not start.
- **THE GENERAL RULE, NOW TWICE: ANY `useState` READ INSIDE `runReconciliation` IS SUSPECT.** State set
  by a mount effect is invisible to a closure already executing. `roleDefs` and `ownerGroupId` are
  fixed; **grep for others before trusting one.**
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.293.0`. **Not yet tested live** - the admin-page pass should now lock or confirm all ten pages
  instead of skipping, on a run started immediately after page load.

## RECONCILIATION NOW GRANTS THE REQUESTS LIST ITSELF (2026-08-27, 1.0.292.0)
Client, having just fixed it by hand: *"how in the world did the CRS_SITE_MEMBERS not applied
automatically for the page and the list?"* **Because nothing ever did.** Two different reasons, and
only one was a defect:
- **Site Pages is REPORT-ONLY BY DESIGN and stays that way.** Granting there would touch the same list
  that holds the ten admin pages the run deliberately locks, so it warns and names the manual steps
  (including *do not tick "share everything in this folder"*). Not a gap.
- **The requests list was never automated at all** - the Requests page's own message says
  *"Its permissions are NOT set automatically"*, recorded as a manual step on 2026-08-21. **SIXTH
  instance of the structural gap**: the mechanism is driven by Group Map rows and this list has none,
  and can have none. Now asserted every run.
- **⚠ AN INHERITING LIST IS ALSO BROKEN, JUST LATER, and that is what makes this worth automating.**
  Inherited Read lets an uploader SEE the list, so the *"requests are not set up on this site yet"*
  message disappears and the button enables - and then the POST 403s, because **Read cannot add
  items**. Both states fail; they fail at different moments, which is harder to diagnose than failing
  consistently.
- **⚠ `copyRoleAssignments=TRUE` HERE, AND `false` EVERYWHERE ELSE IN THE FILE.** On a CRS library,
  inheriting is an EXPOSURE and copying forward would leave it readable by everyone under a green run.
  This pass has the opposite job - ADD one capability, remove nothing - so `false` would strip whatever
  the web granted, including access an administrator set deliberately, to fix a problem that is about a
  MISSING grant. **Do not "make it consistent" with the library pass.**
- **GRANTS, NEVER REMOVES.** The page pass asserts a full ACL because a Site Pages item is a LEAF; a
  list is not the place to start deciding who should *not* have access.
- **A missing `CRS Request` level is REPORTED, never approximated.** Granting Contribute instead would
  hand uploaders **Delete Items** on the one list that IS the record of who asked for what. The message
  names how to create it.
- **A 404 on the list is a `✓`, not a warning** - the list is created when the Requests page is first
  opened, and a site that has never used the feature is a normal state.
- **Accepts `CRS Request` or `DMS Request`.** It is the one level lookup with no `ROLE_TO_PERMISSION`
  entry for `applyPermissionPrefix` to re-point, so both spellings are matched directly.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.292.0`. **Not yet tested live** - on ClarenceDMSTesting the grant now exists by hand, so the
  expected line is `✓ CRS Requests: CRS_SITE_MEMBERS can already raise requests ✓`. The GRANT path is
  what SDG will exercise.

## ⚠⚠ ALL SIX LIBRARIES RENAMED AGAIN, AND `Documents` HAD NO PROBE AT ALL (2026-08-28)
Third rename in two days. **Verified live**, not inferred, with
`/_api/web/lists?$select=Title,ItemCount&$expand=RootFolder&$filter=BaseTemplate eq 101`:

| Logical key | Live title | URL segment (UNCHANGED) | Items |
|---|---|---|---|
| `Staging` | `Approval for Document` | `ApprovalDocument` | 173 |
| `StagingHC` | `Approval for Highly Confidential Document` | `HCApprovalDocument` | 166 |
| `Documents` | **`Restricted & Confidential Document`** | **`Shared Documents`** | 187 |
| `DocumentsHC` | `Highly Confidential Document` | `HCDocuments` | 157 |
| `Archive` | `Archive Restricted & Confidential Document` | `Archive` | 142 |
| `ArchiveHC` | `Archive Highly Confidential Document` | `HCArchive` | 138 |

- **⚠⚠ `Documents` WAS NEVER PROBED — a latent bug this rename finally fired.** Every other library
  resolves through `probeLibrary`; this one was the bare constant `DOCUMENTS_LIBRARY = "Documents"`.
  So there was **no candidate array to extend** — the fix was building the probe
  (`DOCUMENTS_CANDIDATES`, `primeDocumentsLibrary`, `documentsLibraryTitle()`), not adding a string.
- **⚠⚠ IT TOOK THREE PASSES, AND THE SECOND AND THIRD ARE THE LESSON.** That literal was copied into
  FOUR places, and fixing the obvious ones made the remainder LOOK repaired:
  1. `libraryTargets` / `allLibraryTitles` / `DocumentSearch` — the ones a search for
     `DOCUMENTS_LIBRARY` finds.
  2. **`libApiTitle`** — `Documents` fell through to `return lib`, correct only while the library was
     genuinely titled that. **It is the API boundary every `getbytitle` goes through**, so the run
     still 404ed on the entire approved side while the three fixed functions reported the right name.
     Found because the reconciliation panel titles its sections with `libApiTitle`, and that one row
     still read `DOCUMENTS` where the other five read their live names.
  3. **`MySubmissions.tsx` (`const DOCUMENTS`) and `StructureManager.tsx` (`DOCUMENTS_LIST_TITLE`)** —
     private local copies. The first took an uploader's whole history off the page; the second creates
     the tier columns in BOTH libraries, so a 404 there means the approved side silently never gets
     them, which is the `Remark`/`LegallyPrivileged` gap of 2026-08-10 by a new route.
- **⚠ THE MISS IN STEP 3 WAS SELF-INFLICTED AND IS WORTH REMEMBERING: the hunt was
  `grep -rn '"Documents"' src/ | head -20`, and the output was TRUNCATED AT 20 LINES.**
  `src/webparts/...` sorts after `src/shared/...`, so both remaining copies fell off the end and the
  capped list was read as complete — memory `sp-capped-read-reads-as-absent`, committed while fixing a
  bug of exactly that family. **Never `head` a grep whose purpose is to prove absence.**
  - **It falls back to `Documents` rather than to `undefined`, unlike the HC and archive pairs.**
    Those fail closed because their ABSENCE is meaningful; this library always exists, so the useful
    failure is the loud one — a wrong title 404s, where `undefined` would silently drop the approved
    side out of `libraryTargets` and take reconciliation, the migrator and CRS Search with it.
  - `"Documents"` stays FIRST in the array: **SDG's live site has not been renamed**, so it resolves
    on the first request exactly as before.
- **⚠ THE URL IS WHAT IDENTIFIES A LIBRARY THROUGH A RENAME, AND IT SETTLED THE ONE REAL QUESTION.**
  `Restricted & Confidential Document` could have been the renamed `Documents` **or a new
  confidentiality tier** — and mapping a new tier as if it were `Documents` would route approved
  documents into the wrong library. It sits at `/Shared Documents`, so it is the rename. **Always
  `$expand=RootFolder`; the title alone cannot answer this.**
- **⚠ `DOCUMENTS_URL_SEGMENT` NEEDED NO CHANGE** — 13 call sites across 6 files, all still correct,
  because a rename never touches the URL (gotcha #12). Only titles moved.
- **⚠ WORD ORDER IS NOT A NEAR-MISS.** The archive HC title went from
  `Highly Confidential Archive Document` (2026-08-27) to `Archive Highly Confidential Document` —
  same words, different order, and `getbytitle` is exact. Both are listed now.
- **The `&` is a PLAIN `&`, not the fullwidth `＆`** the term store requires — confirmed from the raw
  payload (`&amp;`). `probeLibrary` uses `encodeURIComponent`, so it survives the query string.
- **✅ VERIFIED ON SITE 2026-08-28 (1.0.313.0):** a full run walked all six libraries under their live
  names — 1,643 steps, **0 warnings, 0 errors** — and every line read `already locked, skipped` /
  `already correct, skipped`, which is the proof that a rename touches no ACL and no folder. Being
  clean, it also ran the guarded passes: `Folder Map: no orphaned rows ✓`, `CRS Term Abbreviation: no
  orphaned rows ✓`, `Folders: every folder maps to a live term ✓`.
- **⚠ A TEST NOW PINS THAT EVERY HISTORICAL NAME STAYS LISTED.** Dropping an older entry is the easy
  mistake when a rename arrives — the new one goes on top and the old looks redundant. It is not:
  sites migrate at different times, and a removed name stops resolving **silently**, taking uploads
  down on whichever site still uses it. `naming.test.ts`, alongside the newest-first ordering test
  that caught this change.
- **Both verification scripts were stale too** (`check-hc-setup.js`, `check-library-columns.js`) and
  would have reported healthy libraries as MISSING — worse than not checking, because it sends
  somebody hunting a fault that is in the script. `check-library-columns.js` now RESOLVES each
  library from a candidate list rather than assuming one title.
- **THE STANDING RULE, now paid for three times: after ANY rename, run the REST query above and then
  actually upload one HC document.** A title outside the arrays produces no error on any screen —
  `hcAvailable()` simply goes false and the Highly Confidential level stops being offered.

## ⚠⚠ THE HC LIBRARIES WERE RESOLVING BY LUCK - THEIR REAL TITLES WERE IN NO CANDIDATE LIST (2026-08-27, 1.0.291.0)
Crystal asked whether the libraries could be renamed. Reading the code to answer produced a
contradiction worth the detour, and a REST check settled it. **Verified live titles on
ClarenceDMSTesting**, `/_api/web/lists?$select=Title&$filter=BaseTemplate eq 101`:
`Approval Document` · `Archive` · `Documents` · `HC Approval Document` ·
**`Highly Confidential Document`** · **`Highly Confidential Archive Document`**.
- **THERE IS NO LIST TITLED `HC Documents`, AND NONE TITLED `HC Archive`** - yet HC demonstrably
  works (148 items in the HC approval library, 144 on the approved side) and the run log walks
  `HC Archive/GHO/…` granting ACLs. So `probeLibrary`'s `getbytitle` is matching something other
  than the Title, almost certainly the no-space CREATION name, which is also the URL name
  (`HCDocuments`, `HCArchive`) - the gotcha #12 create-then-retitle pattern. `probeLibrary` returns
  `data.Title`, so everything downstream then uses the live title correctly.
  **⚠ THAT MECHANISM IS INFERRED, NOT READ.** Do not build on it.
- **THE LIVE TITLES ARE NOW EXPLICIT CANDIDATES**, appended so the probe count is unchanged on sites
  using the short names. Resolution no longer depends on an unexplained behaviour.
- **⚠ THE CLIENT RENAMED THEM AGAIN THE SAME DAY, AND THE NEW TITLE WAS ALSO UNLISTED (1.0.294.0).**
  Live titles are now `HC Approval Document`, **`HC Document`** and `HC Archive`. Two of the three
  matched; **`HC Document` is SINGULAR and matched nothing** - so HC was one unexplained URL-name match
  away from being silently dead again, four hours after the last time.
  - **The SINGULAR form is the house style** (`Approval Document`, `HC Approval Document`), so the code
    was changed rather than the client asked to bend. Singular AND plural are now accepted for all
    three HC libraries.
  - **THE REAL LESSON IS THE FREQUENCY.** This client renames libraries as a matter of course, twice in
    one afternoon. **Treat the candidate arrays as data that follows the site, not a fixed contract** -
    and after ANY rename, run
    `/_api/web/lists?$select=Title&$filter=BaseTemplate eq 101` and then **actually upload one HC
    document**, because a wrong title produces no error anywhere: the level simply stops being offered.
- **⚠ THE STANDING RULE THIS EXPOSES: A LIBRARY TITLE OUTSIDE THESE ARRAYS FAILS SILENTLY.**
  `hcAvailable()` goes false, the Highly Confidential level stops being offered on the upload form,
  and HC uploads simply stop - no error, on any screen. **Read the live titles with that REST call when
  provisioning ANY site, SDG included, and add whatever is actually there.** This client renames
  libraries; the candidate arrays have to follow.
- **ANSWER TO THE QUESTION ASKED:** renaming is cheap and touches NO data - Group Map rows hold the
  logical key `Staging`, Folder Map holds UniqueIds, and titles resolve at runtime, so no
  reconciliation is needed. Two routes: add the name to the candidate array (one line, half an hour), or
  rename the library to a listed name (no code). **The only forbidden state is a title matching nothing.**
  And **the nav labels are free text** - the friendly names Crystal wants cost nothing there, which is
  what she had already done.
- **⚠ A RENAME NEVER CHANGES THE URL**, so `Highly Confidential Document` still sits at
  `/HCDocuments` and that old name persists in every folder trail and path - worth saying if the point
  of renaming was consistency.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.291.0`. **STILL OPEN:** confirm the mechanism by probing
  `getbytitle('HCDocuments')` and `getbytitle('HC Documents')` directly - if the former returns
  `Title: Highly Confidential Document`, the URL-name match is confirmed and belongs in this note.

## THE SITE-ENTRY NAME PROBE WAS CAPPED AT 500 ON A 684-GROUP SITE (2026-08-27, 1.0.288.0)
Found while explaining a reconciliation run that named **`DMS_SITE_MEMBERS`** on a CRS site. In that
particular case the message was CORRECT - the group really was still called `DMS_SITE_MEMBERS` and the
client renamed it by hand - but the read behind it is broken and would have produced the same message
falsely.
- **`primeSiteEntry` read `sitegroups?$select=Title&$top=500`** and searched the result.
  ClarenceDMSTesting holds **684**. **FIFTH INSTANCE of the capped-read trap** (memory
  `sp-capped-read-reads-as-absent`), after `fetchAllSiteGroups`, `libraryHasRefColumns`,
  `ensureColumn` and `BulkGroupProvisioner`.
- **⚠ THE FAILURE IS SILENT AND POINTS AT A GROUP THAT DOES NOT EXIST.** `setSiteEntryName` leaves the
  LEGACY name when it finds no candidate, so a truncated read is indistinguishable from the group being
  absent - and every consumer of `siteEntryGroupTitle()` then names `DMS_SITE_MEMBERS`. Worst of them is
  reconciliation's *"NOBODY CAN OPEN THE SITE HOME PAGE"* warning, which would instruct an admin to
  grant Read to a group that is not there, on the one screen they would act on immediately.
- **Now `$filter`ed on the two candidate titles** - at most two rows, so it cannot be truncated - with
  a `$top=5000` fallback as insurance against the FIX, since a rejected `$filter` would otherwise
  restore the original silent failure.
- **⚠ A GROUP RENAME PRESERVES THE ID, so renaming `DMS_SITE_MEMBERS` → `CRS_SITE_MEMBERS` by hand is
  safe** and is what the client did. Nothing was granted or revoked by it.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.288.0`. **Not yet tested live.**

## ⚠ UNEXPLAINED: `addroleassignment HTTP 500 … 0x80131904` ON ONE FOLDER GRANT (2026-08-27)
A full MHO run reported ONE failure among thousands of grants:
`MHO_Corporate Services_Partnership Development_UPLOADER — failed to grant Read on
/MHO/Corporate Services: addroleassignment HTTP 500 (principal 952, role 1073741826)`.
- **`0x80131904` IS A GENERIC SERVER-SIDE EXCEPTION** (SharePoint/SQL), not a permissions or naming
  error. It carries no information about the cause on its own.
- **The run behaved correctly**: one grant failed, it was NAMED with the group, the folder and the
  principal, and the run carried on. That is the design working.
- **The grant is the ANCESTOR-BROWSE Read on a DEPARTMENT folder**, which every unit's groups receive -
  so that one scope takes far more writes than any other in a run. A collision or a throttle there is
  the most plausible cause, and both are transient.
- **DIAGNOSTIC ORDER, none of it done yet:** re-run reconciliation (idempotent - a transient failure
  simply succeeds); if it recurs on the SAME folder, check that principal 952 still resolves
  (`/_api/web/sitegroups/getbyid(952)`) and count the existing role assignments on that folder; only
  then suspect a scope limit.
- **DO NOT write a cause into this file without evidence.** The 1.0.236.0 entry records the last time a
  guess (throttling) was told to the client and was wrong.

## EXISTING GROUPS CAN BE RENAMED TO THE STANDARD SUFFIXES (2026-08-27, 1.0.287.0)
Client: *"you forgot to rename the APR to APPROVAL and UPL to UPLOADER in Groups on this site
(684)."* Not forgotten - 1.0.261.0 changed what NEW names are WRITTEN and deliberately left existing
groups alone, and that section says so in as many words. **The consequence it predicted is exactly
what the client is looking at**: a provisioned site is MIXED, `GHO_GCA_EG_APR_HIGHLY_CONFIDENTIAL`
sitting beside `GHO_GCA_EG_VIEWER_HIGHLY_CONFIDENTIAL`. This is the migration half.
- **A RENAME IS SAFE BECAUSE THE GROUP ID SURVIVES IT.** Every Group Map row is keyed on `GroupId`,
  so mappings, folder grants and page ACLs are untouched - **nobody gains or loses access**. Only the
  stored `GroupName` LABEL goes stale, and `planBulkGroups` matches on the TERM GUID before the name
  anyway (1.0.162.0), so a renamed group is still recognised as provisioned.
- **⚠ THE RULE MUST NOT BE DRIVEN FROM `roleFromGroupName`, AND THAT IS THE WHOLE SAFETY OF IT.**
  That function **FALLS THROUGH TO MEMBER** for anything unrecognised, so a rename derived from it
  would rewrite every hand-named group on the site into a `_VIEWER` - **`CRS_SITE_MEMBERS` and the
  owners group included**. `canonicalGroupRename` matches the SUFFIX LIST directly and returns
  `undefined` when nothing matched, which is what distinguishes *"an uploader group spelled the old
  way"* from *"not one of our groups at all"*. Pinned by a test naming both.
- **`undefined` also for**: a name already canonical, the literal `GLOBAL` (no stem), and a name that
  is nothing but a suffix (`_UPLOADER` alone has nothing to rename).
- **The length sort keeps doing its job**: `_UPL` can never match inside `_UPL_HIGHLY_CONFIDENTIAL`,
  so an HC group is never renamed to the plain form. Pinned.
- **The card appears ONLY when there is something to do** and disappears once the site is consistent,
  rather than sitting there inviting a pointless run. Sequential, carries on past failures, one audit
  row for the whole run - the same shape as the bulk delete beside it, and it shares `bulkRunning`
  and the unload guard.
- **A DUPLICATE TITLE IS REPORTED, NEVER SWALLOWED.** If both spellings already exist, that rename
  fails and is named - the two groups have to be merged or one deleted by hand, which renaming cannot
  resolve on its own.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, `groupMapModel` **159
  passed** (7 new), full suite green. Packaged as `1.0.287.0`. **Not yet tested live** - run it on
  ClarenceDMSTesting, then confirm on a renamed unit that Folder Access still shows its mappings and
  an uploader can still upload.

## OWNERS MEMBERSHIP NOW CARRIES SITE COLLECTION ADMINISTRATOR (2026-08-27, 1.0.295.0)
Client: *"its ok to allow System Admin to have SCA, their version of system admin means they have
access to everything."* Offered as two separate decisions - group membership for the app, SCA for
SharePoint's own surfaces - and they chose to COUPLE them, because in their model the role IS
"access to everything". Adding somebody to the owners group now also promotes them.
- **WHY IT WAS NEEDED AT ALL, and it is not obvious:** Full Control on the WEB does **not** let you
  read a group's membership. SharePoint checks the GROUP's own `OnlyAllowMembersViewMembership`, which
  REST-created groups default to members-only - so an Owner who was not IN a group saw
  *"Could not read this group's members"*. Only the group's owner and **site collection
  administrators** bypass it. Same reason Site Settings and Term Store Manager were out of reach.
- **THE ALTERNATIVE WAS REJECTED, and should stay rejected:** setting
  `OnlyAllowMembersViewMembership: false` on all 684 groups would let EVERY site member see who is in
  every group. The client had just chosen the tighter option for the requests/submissions lists;
  widening this would have contradicted it. SCA keeps membership visible to administrators only.
- **⚠ THE PROMOTION IS REPORTED SEPARATELY AND NEVER FAILS THE ADD.** Only an existing SCA may
  promote, and **a guest may be ineligible** on tenants that block external site collection
  administrators - so refusal is a real outcome, and the person IS in the group either way. The toast
  says which happened and names the manual fix.
- **⚠⚠ THE REMOVAL DIRECTION IS THE DANGEROUS ONE, AND HAS TWO GUARDS.** Leaving SCA behind would mean
  somebody removed from the administrators group keeps full control while the screen says access was
  revoked - so it is removed too, EXCEPT:
  1. **Never the signed-in user.** Demoting yourself removes the very right the call needs, leaving a
     half-changed state.
  2. **Never the last one.** A site collection with no administrator cannot be administered, and
     **nothing in this app could put one back** - only a tenant administrator. `countSiteAdmins`
     answers `undefined` when it could not count, and **an uncountable answer REFUSES the demotion**;
     guessing "there must be others" is how the site is lost.
  Every refusal is stated in the toast, so a skipped demotion is never silent.
- **⚠ `alsoSiteAdmin` IS PASSED ON THE OWNERS MOUNT AND NOWHERE ELSE.** `GroupMembersEditor` is
  mounted for EVERY group in the list and again on Folder Access; passing it elsewhere would promote a
  unit's uploader to site collection administrator. The prop is commented to that effect at both the
  definition and the call site.
- **⚠ IT CANNOT BOOTSTRAP ITSELF: the FIRST site collection administrator is always set by hand**, in
  Site settings, because only an existing SCA may promote.
  - **THE CARD DETECTS THIS AND SAYS SO (1.0.296.0).** Clarence asked whether the manual first step was
    written down anywhere - it was not, and the omission was worse than a missing note: an owner who is
    NOT an SCA would add people and watch **every promotion fail** with a raw SharePoint error they
    could not act on. `isSiteCollectionAdmin` is a SEPARATE read from `isSystemAdmin`, which answers
    true for owners-group membership too and therefore cannot answer this.
  - **Shown only when `viewerIsSca === false`.** `undefined` (read failed) shows NOTHING - a warning on
    every visit that you cannot promote people is how the real one gets ignored on the day it is true.
    Same rule as the ACL banner on Approval Library Access.
- **⚠ STILL OUTSIDE EITHER GRANT, and it matters on SDG: POWER AUTOMATE FLOWS.** They live in the
  owner's account, not the site - an SCA cannot edit or even see a flow they do not own. Fourteen of
  them run Auto-route, HC routing, approver notifications, bulk auto-approve and the audit trail. Build
  them as the SERVICE ACCOUNT and give administrators co-ownership or those credentials, or a "system
  admin" cannot fix half the system. The tenant App Catalog is likewise out of scope.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.296.0`. **NOT tested live** - and the tests that matter are the two guards: try removing
  YOURSELF from the owners group, and try removing the only administrator. Both should keep SCA and say
  so.

## SYSTEM ADMINISTRATORS ARE MANAGED ON GROUP MANAGEMENT NOW (2026-08-27, 1.0.284.0)
Client, urgently: *"we did not include a group as the System Admin group? they should have all the
power. Crystal's second account cannot become system admin to test."*
- **⚠ THERE WAS NEVER A GAP IN THE MODEL - SharePoint's site OWNERS group has always been it.**
  Reconciliation's admin-page lockdown strips **every non-Owners assignment** from every admin page,
  so Owners membership is the ONLY thing that opens them, and Owners hold Full Control on the web
  regardless of any folder ACL. **Do not create a new "System Admin" group** - a second group holding
  the same rights would be invisible to the lockdown pass, which asserts Owners and removes the rest,
  so it would be stripped on the next reconciliation and read as the tool breaking.
- **WHAT WAS MISSING WAS A ROUTE TO IT.** `loadGroups` calls `searchSiteGroups`, which excludes the
  built-in Owners / Members / Visitors **by id** - deliberately, so nobody maps "Site Owners" onto a
  unit folder. Correct for the mapping picker, and it also meant the group never appeared on the one
  page that manages membership. The admin had to leave for Site settings → Site permissions.
- **`fetchOwnersGroup` is SEPARATE from `searchSiteGroups`, and the exclusion there stays.** Two
  different jobs: one asks *"which groups may be mapped"*, the other *"who administers this site"*.
- **⚠ RENDERED AS ITS OWN CARD, NOT MERGED INTO THE GROUP LIST.** That keeps two things true by
  construction rather than by a flag on every row: Owners can never appear among the mappable groups,
  and it can never be offered the **Delete** button. Deleting the site's owners group would take
  every admin page and every Full Control grant with it.
- The card states what the grant actually is - every segment, every library **including HC**, every
  admin page, unrestricted by folder permissions - because "add to Owners" reads like an ordinary
  group membership and is the widest grant in the system.
- **`undefined` says the read failed and points at Site permissions**, rather than implying the site
  has no administrators.
- **⚠ `s.linkBtn` DID NOT EXIST and the build stayed green** - `s` is a `Record<string,
  CSSProperties>`, so a missing key yields `undefined` and the element renders unstyled. Caught by
  listing the keys, not by the compiler. Third file in this project to hit it.
- **⚠⚠ ADDING SOMEBODY TO OWNERS WAS NOT ENOUGH, AND THE CAUSE WAS `IsSiteAdmin` (fixed 1.0.286.0).**
  Crystal's second account was added to `CRS Owners`, signed out and back in, and still could not use
  the Audit Log. **`IsSiteAdmin` IS A PER-USER FLAG AND NO SHAREPOINT GROUP CAN EVER CONFER IT** - it
  means Site Collection Administrator, set one person at a time in Site Settings. So a gate written
  against it can never be satisfied by a group, however much power that group holds.
  - **THE RULE ALREADY EXISTED IN EXACTLY ONE PLACE.** `GroupManager.loadCanManage` had checked
    `IsSiteAdmin` **OR** membership of `AssociatedOwnerGroup` since it was written. The audit log, the
    upload form and Bulk Upload each checked `IsSiteAdmin` alone - so an Owner was an administrator on
    **one screen out of four**, with every permission reading back correct. Classic symptom: *"she is
    in the group and still cannot get in."*
  - **Extracted as `isSystemAdmin` in `spGroups.ts` and now used by all four.** Client:
    *"What I essentially want is to have a group system admin where they can do anything."* Putting
    somebody in `CRS Owners` now makes them an administrator everywhere THIS CODE decides, with no
    per-person SharePoint setting. Fails CLOSED on a read failure, as every call site it replaced did.
  - **⚠ WHAT IT CANNOT REACH, and this is a SharePoint boundary rather than a gap:** anything
    SharePoint itself gates on Site Collection Administrator - **Site Settings, Term Store Manager,
    site collection features** - is unaffected. Owners hold Full Control on the web, so every CRS page,
    library and folder is covered; those site-scoped surfaces are not, and never will be from a group.
  - **A guest may not be eligible to be a real SCA at all** on a tenant that blocks external accounts,
    which is a second reason not to build anything on `IsSiteAdmin`.
- **⚠ THE `A potentially dangerous Request.Path value was detected (<)` ERROR IS UNRELATED.** That is
  ASP.NET rejecting a URL containing a literal `<` before SharePoint sees it - not a permissions
  failure. Nothing in the audit log web part builds such a path (its only request is the admin check
  above), so it points at a mangled link. NOT diagnosed; get the address bar contents if it recurs.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.286.0`. **Not yet tested live** - add a second account to `CRS Owners`, sign out and in, and
  confirm the Audit Log, the upload form's full cascade and Bulk Upload all treat them as an admin.

## ⚠⚠ SUPERSEDED 2026-08-28 — A PENDING FILE **CAN** NOW BE REPLACED. READ THE SPEC FIRST.
**`docs/superpowers/specs/2026-08-28-file-replacement-design.md`.** The client reversed the rule below
one day after making it: the uploader now decides whether to replace a pending draft, and the approver
now decides whether to replace a filed one (*"the overwritten is needed"*). **BUILT — the section
beneath describes the PRE-28-AUGUST code and is history, not behaviour.**
- **⚠ THIS LINE READ "DESIGN ONLY so far — nothing is built" UNTIL 2026-09-04, LONG AFTER IT SHIPPED,
  AND IT COST A ROUND OF CONFUSION.** With that sentence in place, the *"A PENDING ONE NEVER"* heading
  directly below read as the live rule — so the Replace dialog firing on a pending file looked like a
  defect, and the proposed fix was to revert the client's own reversal. Exactly the failure the header
  of this file warns about: **detailed enough to be convincing when stale.**
- **THE LIVE TABLE (upload form, 1.0.396.0), so nobody has to reconstruct it from three sections:**

  | In staging (pending) | In Documents (filed) | What **Yes** does |
  |---|---|---|
  | no | no | uploads normally, no dialog |
  | **yes** | no | `overwrite=true` — **destroys the pending draft NOW** |
  | no | **yes** | ordinary upload; Auto-route replaces the filed copy **only if approved** |
  | **yes** | **yes** | both — the draft now, the filed copy on approval |

  **One dialog covers all three clash rows and Yes means something different in each**, which is the
  real cost of the 2026-09-04 collapse to Yes/No. Row 2 is the dangerous one: draft security means the
  uploader usually cannot SEE the file they are destroying.
- **⚠ THE RENAME ESCAPE IS WHAT LEFT ON 2026-09-04, NOT THE OVERWRITE.** A staging clash offered
  *rename OR replace* from 28 Aug; it now offers Yes/No only. When somebody says "it used to not let
  you overwrite a pending file", they are remembering the **27 August** build (rename only) or
  **Bulk Upload**, which still cannot overwrite staging (`overwrite=false`, hardcoded) and diverges
  deliberately.
- **⚠ RE-CONFIRMED WITH THE CLIENT 2026-09-04** after they asked *"I thought if its sit pending on
  staging it cannot be overwrite?"* — shown the two-way history and the table, they left it as built:
  *"no need to change anything for now."* **Do not revert it on the strength of the 27 August rule.**
- **⚠ The both-libraries rule is VOID, not overruled.** It existed only because staging was
  untouchable — the client's own reason, in the clause after the comma: *"don't allow them to
  overwrite, which is basically the don't allow them to overwrite the staging file."* Remove the
  premise and the rule has no independent motivation. It survives in **Bulk Upload only**, where
  staging genuinely still cannot be overwritten.
- **✅ THE BLOCKING PREREQUISITE IS MET — version history on BOTH approval libraries, verified
  2026-08-31.** Draft security means an uploader cannot SEE a peer's pending file, so Replace destroys
  something they never saw, and versioning is the only recovery. All four CRS libraries read **major
  versions, 500 kept, no time limit**. ⚠ **SDG's own libraries have NOT been checked this way** — do it
  before this reaches their tenant, because the symptom of it being off is a silently destroyed draft
  with no way back.
- `overwrite=true` returns to ONE path (the upload form's explicit Replace) after being deliberately
  deleted in 1.0.271.0. Everywhere else stays `false`.

## A FILED DOCUMENT CAN BE REPLACED THROUGH APPROVAL; A PENDING ONE NEVER (2026-08-27, 1.0.282.0)
**SUPERSEDES 1.0.281.0 ENTIRELY - that build shipped the OPPOSITE scope and existed for about an
hour.** Client's first brief was *"can you add a overwrite function"*; asked to choose a scope they
picked "any pending file in your unit", and then corrected it against the real requirement:
*"if the files is in Document Library approved then allow the client to upload the file to staging to
overwrite so that the approval can approve. Never allow anyone to overwrite a pending file in
staging."* The revised rule is better in every respect and is what is built.
- **THE TWO HALVES:**
  - **A PENDING draft is untouchable, by anyone.** The approval-library existence check is NEVER
    skipped and `overwrite` is hardcoded `false` in both screens. Somebody's unreviewed work - which
    they may not even be able to see - cannot be destroyed by an upload.
  - **A FILED document no longer BLOCKS the upload.** The approved-side clash becomes an explicit
    opt-in instead of a refusal: the file goes to the approval library under the SAME name and the
    approver decides, exactly as the client described.
- **⚠ IT SETS NO `overwrite` FLAG ANYWHERE, AND THAT IS WHY IT IS SAFE.** The name is free in the
  approval library - its only occupant is the filed copy, one library over - so this is an ORDINARY
  upload. All the opt-in does is stop the approved-side check refusing. `grep overwrite=true` returns
  comments only.
- **⚠⚠ THE SECOND HALF IS A FLOW CHANGE, AND THE CLIENT HAS DECIDED IT: `Copy file` GOES BACK TO
  REPLACE.** *"I think the power automate I must allow them to Replace, apparently they say they will
  be careful, but they definitely want to replace the document library file, that is on purpose to
  ensure they update the file."* **This REVERSES the 2026-08-25 change to
  `nameConflictBehavior: 2`** - deliberately, with the trade stated and accepted. Set it on **BOTH
  `Auto-route` AND `HC Auto Route`**; every HC clone in this project has shipped faults from a
  reference nobody swapped. Verify in Code view: **`"nameConflictBehavior": 1`** (0 = Fail, which
  strands documents; 2 = copy with a new name).
- **⚠ VERSION HISTORY ON `Documents` AND `HC Documents` IS WHAT MAKES THIS SOUND, and it should be
  confirmed before the switch.** Replacing a file in a VERSIONED library creates a NEW VERSION and
  keeps the old content - which is precisely *"update the file"*, the client's own words. Without
  versioning the previous content is gone and there is no trail.
  - **✅ VERIFIED ON ClarenceDMSTesting 2026-08-27, `Documents`: version history ON, major versions,
    500 kept, no time limit; content approval OFF (correct for the approved side).** So a Replace
    writes version 2.0 and the old content is recoverable.
  - **✅ APPLIED AND CONFIRMED 2026-08-31.** Both `Copy file` actions now read **Replace**, and
    versioning was verified on BOTH approved-side libraries first: `Restricted & Confidential Document`
    AND `Highly Confidential Document` (list `E4FB3B2C-...`, at `/HCDocuments`) — major versions, 500
    kept, no time limit, content approval OFF. **Checking HC BEFORE flipping its flow was the point:**
    Replace over a library with no version history destroys the previous content outright, and HC is
    the class of document where that matters most.
  - **⚠ SDG's two libraries are STILL unchecked.** Same page, same three values, before either flow
    is imported there.
  - **⚠ STILL AN ASSUMPTION, AND IT IS THE WHOLE SAFETY ARGUMENT: nobody has confirmed that
    `Copy file` with Replace writes a VERSION rather than deleting and recreating the item.** Versioning
    protects the old content only in the first case. Prove it on the first live replacement — Version
    History on the filed document should read **2.0 with the old content at 1.0**. A single `1.0` means
    the item was replaced, the previous content is gone despite versioning being on, and the client's
    decision has to be revisited before SDG.
  - **Worth confirming on the first live replace:** the filed document should KEEP its identity and
    gain a version rather than being deleted and recreated - open Version History in `Documents` and
    expect 2.0 with the old content at 1.0. If instead it comes back as a single 1.0, the copy is
    replacing the ITEM rather than versioning it, and the old content is gone despite versioning
    being on.
- **The app-side checks mean Replace now mostly fires ON PURPOSE**, which is what changed the risk
  calculus. Both upload screens REFUSE an approved-side clash unless the uploader explicitly presses
  *Send for approval as a replacement* - so a routed document only meets an existing name because
  somebody chose it. **This matters most for BULK IMPORT**, where auto-approve puts no human in the
  loop: a 50-file import with colliding names is SKIPPED unless the admin opts in per run.
- **⚠ TWO RESIDUAL PATHS REPLACE WITHOUT ANYONE CHOOSING IT**, and versioning is the only thing behind
  them: the **native Approve/Reject command** in the library view, which bypasses every app-side check
  (2026-08-25), and a **race** - a name that was free at upload and filled before approval.
- A marker column (`ReplacesFiled`, mirroring `BulkImport`) was designed to keep Replace scoped to the
  opt-in only. **No longer needed for the client's feature** now that Replace is global, and it is
  recorded here as the fallback if the residual paths above ever cause a real loss.
- **⚠ A NAME TAKEN IN BOTH LIBRARIES IS RENAME-ONLY, and the ORDER OF THE TWO CHECKS is what enforces
  it** (client: *"if the file is existing in document library and also exist in staging library, don't
  allow them to overwrite, which is basically the don't allow them to overwrite the staging file"*).
  The STAGING check is evaluated FIRST in both screens and returns immediately, so `approvedClash` is
  never set and the dialog cannot offer *Send for approval as a replacement* for a file whose real
  blocker is a pending draft.
  - Checking the approved side first was **safe but wrong-feeling**: the staging check still caught it
    on the retry, so nothing could be overwritten - but the uploader was shown a button that could not
    work, pressed it, and got a second refusal. Safety was never the issue; a dead button was.
  - It also saves a request in the common case: a staging clash returns without asking the approved
    side at all.
- **THE FOUR CASES, AS BUILT (both screens, identical):**
  | in staging | in Documents | what happens |
  |---|---|---|
  | no | no | uploads normally |
  | no | **yes** | offered *Send for approval as a replacement*, OR a free name |
  | **yes** | no | **rename only** - free name offered, no replace button |
  | **yes** | **yes** | **rename only** - the staging draft is the blocker |
- **The sibling / within-batch collision check is UNTOUCHED and still runs first of all.**
  `collisionsWithin` blocks SAVING a batch whose files compose to one name, and
  `duplicateAcrossBatches` covers the cross-batch case - both before any upload happens, so a rename
  offered here can never collide with another file in the same run. Verified live 2026-08-27.
- **The button is offered ONLY for `alreadyFiled` / `approvedSide` offers**, so an approval-library
  clash is still rename-only. The confirm says what actually happens - *"will be replaced only once an
  approver approves it. Nothing is changed now."* - because nothing IS destroyed by pressing it.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green, no live
  `overwrite=true` in either file. Packaged as `1.0.283.0` (check ordering). **NOT tested live.**

## THE UPLOAD FORM'S FILE LIST SCROLLS — BUT ONLY WHILE COLLAPSED (2026-08-27, 1.0.280.0)
Client: *"if the file have 30 files it will be too long to scroll, add the scroll feature for the
upload form."* True the moment `MAX_FILES_PER_BATCH` arrived — thirty collapsed rows is roughly 1,400px
before the Save button is reached. **The cap was lowered to 20 later the same day, which shortens the
list without removing the need for this**: 20 collapsed rows still runs past the fold on a laptop.
- **⚠ THE CAP IS LIFTED WHILE ANY ROW IS OPEN, and that is not a detail.** The per-file editor's
  **Confidential Level** and **Legally Privileged** info panels are ABSOLUTELY POSITIONED children of
  the row, and a scroll container CLIPS them — which is the bug already recorded on `.dms-staged-row`
  (*"the definitions the icon exists to show were unreadable"*), and the same reason Group Management's
  group list lifts its own 60vh cap while a group is expanded. **THIRD instance of this trap**; a
  fourth screen adding a scroll box must check for absolutely-positioned descendants first.
- **The trade, accepted:** with a row open the page is long again, because the other 29 rows are still
  rendered. The alternative — capping always and re-basing the tooltips to `position: fixed` — is the
  "proper" fix and was rejected as the riskier change for a layout complaint.
- **The `N documents selected` header stays OUTSIDE the box**, so the count never scrolls away —
  matching the abbreviation editor, which keeps its warning, collision banner and Save outside its own
  58vh box.
- Applied by inline style rather than a CSS class because the cap is STATE-dependent, and this file's
  styles are a static template literal.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.280.0`. **Not yet tested live** — stage 30 files and confirm the list scrolls, then open one and
  confirm both info tooltips render in full.

## BULK UPLOAD'S APPROVED-SIDE RENAME IS VERIFIED LIVE (2026-08-27)
`TEST - TEST - TEST - 26-08-26 (2) - Copy - Copy.pdf` clashed with an already-filed document and was
offered `… - Copy - Copy - Copy.pdf`, carrying the warning *"The original has already been approved and
filed — check this is not the same document."* Both halves of the 1.0.270.0 decision working: a name IS
offered, and the fact that the original went through approval is still said out loud. The new
de-select `✕` also renders on the skipped row.

## A SKIPPED OR FAILED ROW CAN BE DE-SELECTED AFTER A RUN (2026-08-27, 1.0.279.0)
Client, immediately after the delete button was removed: *"i want it so it will be easier for them to
deselect."* Fair — once a run has happened the progress list REPLACES the picker, so there was no way
to drop a single file without pressing Cancel and re-picking the whole set.
- **⚠ THIS IS NOT THE DELETE COMING BACK.** It touches nothing on the server; it is the picker's own
  de-selection made reachable afterwards. The old `✕` sat on **uploaded** rows and recycled the
  document; this one is offered ONLY on rows that were never uploaded. The `title` says so outright
  (*"nothing is deleted"*), because the two controls look identical and sit in the same column.
- **`done` IS EXCLUDED, `tagFailed` IS OFFERED**, and the asymmetry is deliberate: an uploaded row is
  the only record that the file went, while a `tagFailed` file **stays in `picked`** (only outcome
  `uploaded` is filtered out post-run) and would otherwise be re-sent on the next press.
- **⚠ MATCHED ON THE `File` OBJECT, never index or name** — `LiveFile` now carries it. A run removes
  its successful files from `picked`, so the two lists stop lining up the moment anything succeeds;
  and this screen keeps ORIGINAL filenames, so two files added from different folders can share one.
  Third place this screen has needed identity matching (`renameOverrides`, the post-run `setPicked`
  filter, now this).
- **⚠ EMPTYING THE LIST SETS `live` TO `null`, NOT `[]`.** The render is
  `live ? progress : picked.length === 0 ? dropzone : picker`, and **an empty array is truthy** — so
  dropping the last row would leave a dead empty panel with no dropzone and no *Add more*, i.e. no way
  back to picking files at all.
- A pending rename offer for the removed file is dropped with it, or it would outlive the row and
  re-upload the file on Proceed. Held entirely while a run is in flight.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.279.0`. **Not yet tested live.**

## PROJECT NAME AND VENDOR/CUSTOMER NAME ARE OPTIONAL (2026-08-27, 1.0.278.0)
Client's request. They were required so every saved file carried the full
`[Project] - [Vendor] - [Document Name] - [Date]` shape; the client's position is that not every
document has a project or a vendor, and demanding a placeholder puts junk in a metadata column to
satisfy a naming convention.
- **⚠ THE COST, STATED PLAINLY: FILENAMES ARE NOW LESS UNIQUE.** `composeUploadBase` DROPS blank
  parts, so a document with neither composes to `[Document Name] - [Date]` — and two documents of the
  same kind filed on the same day now collide where they previously could not. **Handled rather than
  prevented**, and only because the clash work landed first: `collisionsWithin` blocks the save, and
  the six clash branches offer a ` - Copy` name at upload. Do not re-tighten without the client asking.
- **A SHORTENED NAME IS AFTERWARDS INDISTINGUISHABLE FROM A DELIBERATE ONE.** That was the original
  argument for requiring them and it is still true; it is now an accepted consequence rather than a
  defect to fix.
- **⚠ `nameSettled` HAD TO CHANGE WITH IT, and missing that would have been silent.** It tested all
  FOUR name parts as a proxy for *"the uploader has finished typing"*. Blank is now a legitimate END
  state, so a file correctly omitting both would NEVER be settled — its row would show the original
  filename for ever, and a genuine clash between two such files would go unseen until Save. It now
  tests only the REQUIRED name parts (Document Name, Document Date).
  - Residual, accepted: two files sharing a Document Name and Date can flag each other for the moment
    between filling those in and typing a Project that would distinguish them. Self-clearing, and
    advisory — `collisionsWithin` at Save is the real gate.
- **`missingForFile` is the ONLY required-field gate** — `canSaveBatch` checks files, path and
  collisions and nothing else — so the rule lives in one place. The `*`, the `invalid` class, the
  `aria-invalid` and the two error lines came off both fields; the hint now reads `Optional · max. 50
  characters`.
- **Bulk Upload is untouched and needed nothing**: its `vendor` is a permanently blank `useState` with
  no setter, and it never collected a project name.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green. Packaged as
  `1.0.278.0`. **Not yet tested live** — save a batch with both fields blank and confirm it uploads as
  `[Document Name] - [Date].pdf`.

## THE INVISIBLE-PENDING CLASH IS VERIFIED END TO END, TWO REAL ACCOUNTS (2026-08-27, 1.0.279.0)
The case the whole clash mechanism exists for, and the one the pre-checks are STRUCTURALLY BLIND to.
Staged on ClarenceDMSTesting with two uploader-only guests in `GHO/GCA/EG/2024/Agreement`, neither an
approver there, so Draft Item Security hides each one's pending files from the other.
- **`clarencechojinheng` uploaded five via the form** — four landed Pending, one was caught by his own
  pre-check because he was the AUTHOR of an identical 9-hour-old file and could therefore see it. Worth
  keeping: **the author exception makes a name look taken to its uploader and free to everybody else.**
- **`chocheetuck4` then uploaded the same five: 0 of 5 written, 5 of 5 caught**, every one through the
  **Add-failure branch** — confirmed by the wording, which only that branch produces: *"A document
  called X already exists in this folder. It may be a pending upload from someone else that you cannot
  see until it is approved."* His pre-check listed the folder and found NOTHING, so `Files/Add(
  overwrite=false)` refusing the write is the only thing that could have caught them.
- **Five distinct ` - Copy` offers, Proceed accepted, all five landed.** The library then held BOTH
  sets side by side — five originals by `clarencechojinheng` and five ` - Copy` by `chocheetuck4`,
  nothing overwritten, `Created By` correct on each.
- **✅ THE WITHIN-BATCH COMPOSED-NAME GUARD IS VERIFIED TOO**, by adding a sixth file whose Project /
  Vendor / Document Name / Date composed to the SAME name as another in the batch. **Both rows** were
  flagged amber (*"Another document in this batch would be saved under this same name… change one of
  those"*) — both, because which one is wrong is not knowable, exactly as with the abbreviation sibling
  check. The row also showed its `Saves as:` line, which is what makes the collision explicable rather
  than mysterious.
- **⚠ `reserved` REMAINS UNEXERCISED, and is now DEFENSIVE DEPTH rather than a live path.** It fires
  only when two files in one run are offered the same free name — which the never-strip rule already
  prevents for different originals, and which `collisionsWithin` / `duplicateAcrossBatches` already
  block for identical composed names. Keep it: it is the only thing standing behind those guards if a
  future change loosens either, and it costs nothing. **Do not read "untested" as "unnecessary."**

## THE CLASH SUFFIX IS ` - Copy`, AND IT NEVER STRIPS (2026-08-27, 1.0.275.0/276.0)
**REVERSES 1.0.273.0 ENTIRELY.** Client, reading a Bulk Upload dialog that offered ` (3)` for TWO
different files: *"look at the second clash checking Documents Library, it increments to (3), so if
both files go through Approval Document to be auto approve for bulk upload, won't both files clash at
the same time?"* Yes. Their proposed fix — *"I think its better if you append - Copy instead"* — is
right, and for a structural reason worth writing down.
- **⚠ THE INCREMENT MERGED TWO NAMESPACES, AND I INTRODUCED IT THE SAME DAY.** 1.0.273.0 stripped a
  trailing ` (n)` and resumed counting, so `X.pdf` and `X (2).pdf` both reduced to the stem `X` and
  drew from ONE counter. Two files, two different original names, one suggestion. **Appending cannot
  do this**: the suffix goes on the WHOLE original, so distinct originals stay distinct
  (`X - Copy.pdf` vs `X (2) - Copy.pdf`). **Never strip is the invariant; the per-file prettiness of
  incrementing is what cost it.**
- **A WORD, NOT A NUMBER, and that also settles the 2026-08-26 objection properly.** The client
  rejected `TEST-2` because a digit reads as one more segment of
  `[Project] - [Vendor] - [Name] - [Date]`. ` - Copy` cannot be mistaken for a segment, and it is what
  Windows itself produces — the client's own file listing showed
  `TEST - TEST - TEST - 26-08-26 (2) - Copy.pdf`, so nobody has to be taught it.
- **⚠ A SUGGESTION MUST ALSO CLEAR THE NAMES THIS RUN HAS ALREADY OFFERED.** `nextAvailableName` sees
  only the folder listing, and NEITHER clashing file has been written yet — so the listing cannot
  separate them however good the suffix scheme is. Both upload screens now thread a `reserved` array
  through the run and feed every offered name back in as taken. Nothing was ever overwritten
  (`overwrite=false`), but the second file was refused a second time, turning one round trip into two.
- **⚠ AND THE APPROVAL-LIBRARY BRANCH WAS CLEARING ONLY ONE LIBRARY — in BOTH files.** The rule
  *"the suggestion must clear BOTH libraries"* was written for the approved-side branch (1.0.270.0)
  and never applied to its sibling twenty lines above it, so a name free in the approval library could
  still collide inside Auto-route's `Copy file` and be renamed silently at routing time. **FOURTH
  instance today of a fix landing on one path and not its sibling** — and this time both siblings were
  in the same function.
- **The three fixes are independent and all three were needed.** The suffix scheme stops different
  originals colliding; `reserved` stops identical originals colliding; clearing both libraries stops
  the collision moving to Auto-route. Any one alone leaves a live path.
- **THREE THINGS A SUGGESTION IS CHECKED AGAINST, and it takes all three** (client: *"so now it will
  check each other siblings and not just the documents library and approval document correct?"*):
  the **approval library** listing, the **approved-side** listing, and the run's own **`reserved`**
  offers. A sibling that already UPLOADED in this run is covered by a fourth thing — Bulk Upload
  re-lists the folder per file, so it sees it.
- **⚠ THE FORM DOES NOT RE-LIST — `takenFor` CACHES PER FOLDER FOR THE WHOLE RUN**, so its listing
  predates anything the run has since written. That is fine for correctness (the live `Exists` probe
  still detects the clash) but it left the approval-library branch able to compute a suggestion from a
  listing that did not contain the clashing name — `nextAvailableName` would then hand back the name
  that just failed, the guard would discard it, and the uploader would get a **refusal with no offer**.
  Fixed in 1.0.276.0 by seeding `taken` with `finalName`, which the probe has just proved is taken.
  Same reasoning as the Add-failure branch, where the listing can come back trimmed to nothing.
- **⚠ EACH SCREEN HAS THREE CLASH BRANCHES AND ALL THREE NEEDED THE SAME TREATMENT.** The two
  pre-checks (approval library, approved side) and the **Add-failure** branch. The form's Add-failure
  branch was missed on the first pass and fixed in 1.0.277.0 — it neither reserved nor cleared the
  approved side. **It is the branch that fires for a CONCURRENT uploader**, i.e. the one case the
  pre-checks are structurally blind to, so it is the last one that should be left behind.
  **When changing clash handling, the unit of work is SIX branches, not one.**
- **⚠ NONE OF THIS COVERS TWO PEOPLE UPLOADING AT ONCE.** `reserved` is per RUN and the listings are a
  moment old; only `overwrite=false` on the write, and Auto-route's `nameConflictBehavior: 2`, stand
  between two simultaneous uploaders and a collision. Neither destroys anything — that is the point.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, `uploadBatches` **61 passed**
  (the describe block was rewritten; new tests pin the never-strip invariant and that two different
  originals are never offered one name), full suite green. Packaged as `1.0.276.0`.
- **✅ VERIFIED LIVE 2026-08-27, Bulk Upload, six clashing files in one run.** All six offers DISTINCT,
  each derived from its own original, and the walk demonstrably correct: `… (2).pdf` was offered
  ` (2) - Copy (2)` because ` (2) - Copy` was itself one of the clashing files, and `… (2) - Copy.pdf`
  was offered ` - Copy - Copy (2)` for the same reason. Reservation and the folder listing cooperating.
- **⚠ THE FIRST RUN OF THIS TEST SHOWED THE OLD SCHEME AND LOOKED LIKE A CODE FAULT.** `1.0.275.0-278.0`
  had not been deployed — the tab was still serving an older bundle. **The tell was in the dialog:** two
  files offered the SAME ` (3)`, which only the pre-`reserved` build can do. One offer coincidentally
  read `… - Copy (2)` and looked like the new scheme; it was the OLD `(2)` appended to a file already
  named `- Copy`. **Read every row before concluding which build is running**, and check Site Contents,
  not the catalog. **Not yet re-tested
  live** — re-run the four-file clash and confirm the four offers are all DIFFERENT.

## AN ALREADY-NUMBERED NAME IS INCREMENTED, NOT SUFFIXED TWICE (2026-08-27, 1.0.273.0)
Client, on the first live Bulk Upload clash test: the dialog offered
`TEST - TEST - TEST - 26-08-26 (2) (2).pdf` — *"why not you increment this? instead of (2) (2), if I
am not wrong usually by default the file will increment the number?"* Right, and it is a defect rather
than a cosmetic one: Windows and SharePoint both increment, so a stacked suffix reads as the feature
being broken on the very screen that exists to reassure the uploader nothing was destroyed.
- **THE CAUSE IS THAT BULK UPLOAD KEEPS THE FILE'S OWN ORIGINAL NAME.** `nextAvailableName` took the
  whole stem — ` (2)` included, because the file on disk had been saved from an EARLIER offer — and
  appended another. It never looked at the number already on the end.
- **The count RESUMES from that number** (`(2)` → probe `(3)`, `(4)`, …), and the probing behaviour is
  unchanged: a resumed number can itself be taken, so it still walks rather than appending blindly.
- **⚠ THE BARE STEM IS NEVER OFFERED BACK, even when it is free.** Stripping ` (2)` and handing back
  `TEST.pdf` renames the uploader's file to something they did not choose — and, in a folder where a
  DIFFERENT document holds that bare name invisibly (draft security), walks straight back into a clash.
  Pinned by a test asserting the bare name is never the answer.
- **Only a trailing bracketed NUMBER counts.** `TEST (final).pdf` is suffixed normally as
  `TEST (final) (2).pdf` — that bracket is part of the name. And a name that is NOTHING but a counter
  (`(2).pdf`) keeps its whole stem, since stripping would leave an empty one; the regex demands a
  non-space character before the bracket, which is what enforces both.
- **The cap is measured from wherever the count resumed**, so a file arriving as ` (998)` still gets a
  full run of attempts rather than falling straight through to the timestamp fallback.
- **⚠ IT ALSO FIXES A CLAIM THIS FILE ALREADY MADE AND WAS WRONG ABOUT.** The 1.0.268.0 section says a
  suggestion colliding with another invisible pending file *"offers the next number"* — it was in fact
  producing `(2) (2)`, because the second pass fed its own suggestion back in as the desired name. That
  path converges properly now. **Form.tsx benefits without being touched**: both screens call the one
  shared rule.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, `uploadBatches` **63 passed**
  (5 new tests), full suite green. Packaged as `1.0.273.0`. **Not yet re-tested live** — re-run the
  Bulk Upload clash with a file already named ` (2)` and confirm the offer reads ` (3)`.

## A NAME CLASH NOW OFFERS A FREE NAME INSTEAD OF ONLY REFUSING (2026-08-26, 1.0.266.0-270.0)
Client: *"when the file in staging is pending which person A uploaded then person B upload the same
file name, they dont want the file to overwrite, they want to allow user to upload the file but with a
different name … it will show that popup warning saying the existing file exist your filename will be
TEST-TEST-TEST-2.pdf, then ask them if they want to proceed with the new file name or cancel and
rename that specific file themselves."*
- **⚠ THIS IS NOT THE OLD REPLACE PROMPT COMING BACK.** *"Replace Existing File?"* was removed on
  2026-08-15 because replacing destroys a record an approver may already have acted on. This asks a
  DIFFERENT question — may we file it under a FREE name — so both answers are safe and the existing
  document is untouched either way. `overwrite=false` is unchanged.
- **`nextAvailableName` in `uploadBatches.ts`** (pure, 10 tests) PROBES ` (2)`, ` (3)`, … rather than
  appending `2`: ` (2)` may itself be taken (the third person to upload that name), and a blind append
  hands back a name that clashes just as hard, turning one refusal into two. Suffix goes BEFORE the
  extension, matching is case-insensitive because SharePoint names are.
- **PARENTHESISED, not hyphenated** (client on seeing it live, 1.0.267.0: *"can you use that (2)
  enquote to make it obvious"*). `TEST-2.pdf` read as part of the name against a convention already
  built from hyphens — `[Project] - [Vendor] - [Name] - [Date]` — so the suffix was indistinguishable
  from a segment. ` (2)` cannot be. It also matches `SubtreeMigrator`, which already suffixes
  collisions as `a (testig).pdf`.
- **The dialog carries the project's AMBER WARNING ICON** — the same `#FF952A` exclamation Bulk
  Upload's own clash prompt uses, not the success popup's teal tick that sits a few lines below it in
  the same file. Two prompts about the same situation should look like the same system. The name list
  stays LEFT-aligned inside the centred popup: a struck-through old name above a new one only reads as
  a before/after pair when both start at the same edge.
- **THE DIALOG COMES AFTER THE RUN, NOT BEFORE IT — and that was a deliberate reversal.** A true
  pre-flight would have to hoist every per-batch step (folder-map lookup, unit resolve, ensure-create,
  HC routing) out of the upload loop and hold that plan in state across a dialog, in the most
  demo-critical function of a 4,200-line file. And `ensureFolder` CREATES folders, so they would be
  created before the question was answered anyway. Running first costs nothing that was not going to
  happen: clashing files are refused, stay staged, and nothing is overwritten.
- **`UploadResult.suggestedName` is the discriminator.** Its PRESENCE is what makes a refusal
  offerable — an APPROVED-side clash deliberately leaves it undefined, because that document has been
  through approval and filing a renamed second copy is not the right answer there.
- **⚠ `handleUpload` TAKES AN OVERRIDE, and it has to.** Proceed applies the names and re-runs, but
  `setBatches` has not landed when the callback fires — reading state would send the OLD names and
  refuse all over again. Every read of the batch list inside it goes through `source`, never `batches`.
- **Offers are matched against files STILL STAGED**, not against the results array alone: a file that
  somehow succeeded must never be offered a rename it no longer needs.
- **⚠ THE FOLDER LISTING CACHES THE PROMISE, NOT THE ARRAY — `require-atomic-updates` failed the build
  on the first attempt and was right to.** Caching the value puts an `await` between the cache check
  and the cache write, so two callers for one folder both miss and both fetch. Uploads happen to be
  sequential today and nothing guarantees they stay that way. Keyed per FOLDER, because one batch can
  write into two libraries (confidentiality is per FILE) and a name must be checked against the folder
  it is actually going into.
- **Fetched LAZILY and only on a clash** — no cost on the overwhelmingly common path — and `undefined`
  on any failure, which means "no suggestion" rather than "no names": an unreadable folder must never
  turn a rename offer into a blocked upload.
- **Verified**: `tsc --noEmit` clean, `heft build` clean, `uploadBatches` **58 passed**, full suite
  green. **Verified live**: the dialog appears with the amber icon, the struck-through old name and
  the offer.
- **⚠ AND THE CLIENT'S ORIGINAL SCENARIO DEFEATED ALL OF IT UNTIL 1.0.268.0 — the checks are BLIND to
  another person's pending file.** Proven on site: `chocheetuck4` had a pending
  `TEST - TEST - TEST - 26-08-26.pdf` in `GHO/GCA/EG/2024/Agreement`; uploading the same name as
  `clarencechojinheng` produced a raw **HTTP 500 — "A later version of this item has already been
  modified. Other users can not edit the item until that version is approved or published"** and **no
  dialog at all**.
  - **Draft Item Security on the approval library is *"Only users who can approve items (and the
    author)"*,** so A's pending file is invisible to B — and therefore to both of our checks.
    `Files('name')?$select=Exists` returns a security-trimmed **404** that reads as *no clash*; the
    folder listing does not contain it, so no suggestion could be computed; only `Files/Add` knew.
  - **So the offer is now ALSO made from the Add failure**, or it could never fire for the exact case
    it was built for. ⚠ **Matched on several signals, not one status** — the moderation conflict is a
    500, an ordinary duplicate can be 400 or 409, and the wording differs. Matched loosely on purpose:
    a false POSITIVE only offers a rename the uploader can decline, while a false negative puts a
    versioning error in front of them and hides the fix.
  - **The suggestion is computed from the visible names PLUS the now-known-taken name**, because the
    listing may be empty or trimmed — that is how we got here — so `finalName` is added explicitly and
    a suggestion is guaranteed even when nothing could be listed at all.
  - **If the suggestion collides with ANOTHER invisible pending file, the same branch fires again and
    offers the next number.** It converges, and every step is visible rather than silently looping.
  - The message no longer describes versioning to someone who only picked a filename: it says the name
    already exists and *"may be a pending upload from someone else that you cannot see until it is
    approved."*
- **VERIFIED LIVE** at 1.0.268.0: uploading the same name as another person's PENDING file now raises
  the dialog and offers ` (2)`, for two files at once.
- **⚠ THE DIALOG THEN LISTED ONE OF TWO FAILURES, WHICH READ AS A BUG (fixed 1.0.269.0).** A batch with
  one approval-library clash and one APPROVED-side clash showed only the first — the second was
  refused, correctly, but silently omitted from the popup, so the client's reaction was *"that is why
  I dont see the first file."* A dialog that accounts for some of the failures is worse than one that
  accounts for none, because it looks complete.
  - `UploadResult.approvedClash` marks it. Kept SEPARATE from `suggestedName` because the two need
    opposite handling and the dialog has to say which: an approval-library clash is offered a free
    name, while a document already filed on the approved side has been through approval, so the useful
    response is *"check whether this was already uploaded"* — not a renamed second copy an approver
    then has to reconcile against the original.
  - The popup now has **two blocks**: the renameable ones (neutral grey, struck-through old name above
    the new one) and the already-filed ones (amber, explained, no offer). Title and intro cover THREE
    shapes — renameable only, filed-already only, and a mix — and the Proceed button renders only when
    there is something renameable, with Cancel becoming **Close** when there is not.
  - The `Documents` / `HC Documents` wording is derived from `hcAvailable()`, so a site without the HC
    pair is not told about a library it does not have.
- **VERIFIED LIVE at 1.0.269.0, whole chain**: the mixed dialog listed both files, Proceed filed
  `TEST - TESTTEST - TEST - 26-08-26 (2).pdf` into the library, the `Documents`-clash file stayed
  behind, and the notes-only case rendered with a single **Close**.
- **THE APPROVED SIDE IS OFFERED A RENAME TOO SINCE 1.0.270.0** (client: *"why not auto rename the file
  for the Documents library version as well?"*). The earlier explain-only stance is **superseded**, and
  the deciding argument is one the original reasoning missed: **Auto-route renames it anyway.** Since
  2026-08-25 `Copy file` uses `nameConflictBehavior: 2`, so approving a duplicate produces a
  SharePoint-renamed copy (`TEST1.pdf`) at routing time and tells nobody. A visible rename now beats an
  invisible one later.
  - **⚠ THE SUGGESTION MUST CLEAR BOTH LIBRARIES.** The file lands in the approval library NOW and is
    routed to the approved side LATER, so a name free only in the approval library just moves the
    collision to Auto-route. Probed against the approved-side names, the approval-library names, and
    the known-taken one.
  - **`checkApprovedClash` became `approvedClashInfo`**, returning `{ clash, taken }` — it LISTS the
    mirrored folder rather than probing one name, which is the same single request and yields the names
    a suggestion needs. Still fails OPEN on any failure.
  - **⚠ `tsc` CAUGHT A SECOND CALLER: the SELF-APPROVE path.** It uses the same check to SKIP
    self-approve on a clash, and breaking it would have auto-approved over an existing document. It
    reads `.clash` only — the names are for building a suggestion, which has no meaning there.
  - **Offered AND still flagged.** `approvedClash` stays true even when a name is offered, and the
    offer row says *"The original has already been approved and filed — check this is not the same
    document."* A renamed copy of an approved document is a SECOND document, which is a different
    thing from a first one, and the uploader should know that before clicking. `clashNotes` now holds
    only the ones no name could be found for, so nothing is ever silently omitted.
- Packaged as `1.0.270.0`. **Still to confirm live** — that a `Documents` clash now appears in the
  OFFER block with its warning rather than the amber block, and that a third upload of one name offers
  ` (3)` rather than ` (2)` again.

## THE PER-FILE PREFILL IS GONE — EVERY ADDED FILE STARTS BLANK (2026-08-26, 1.0.265.0)
Third and final round of the same evening's thread. With the row names fixed, the remaining complaint
was the state underneath them: *"the user input is suddenly halfway automatically filled."*
- **`inheritDefaults` IS NO LONGER CALLED.** It copied the previous file's Project, Vendor, Date and
  Confidentiality into each newly added file — everything except Document Name, which had to stay
  unique. **That asymmetry was the whole problem:** a new row arrived two-thirds filled with values the
  uploader had never typed for THAT document, and nothing on screen said where they came from.
- **Offered the choice between LABELLING the prefill and REMOVING it, the client chose removal.**
  Recorded because the trade is real and the cost lands on them: a ten-file batch to one vendor now
  means typing the same Project, Vendor and Date ten times. `inheritDefaults` was written precisely
  because *"effort is the whole complaint"*, so this may well come back.
- **⚠ IF IT DOES COME BACK, THE ANSWER IS A LABELLED PREFILL OR AN EXPLICIT "COPY FROM PREVIOUS"
  BUTTON — never silently restoring this.** Silent prefill is the behaviour that was just rejected.
  `inheritDefaults` is deliberately LEFT in `uploadBatches.ts` with its tests, unused, so either
  option is a small change rather than a rewrite.
- **The two fixes from earlier in the evening still earn their place**, and this is why they were not
  reverted with the prefill: `nameSettled` gates both the clash check (1.0.263.0) and the row label
  (1.0.264.0), and a HALF-TYPED row is exactly as misleading as a prefilled one. Neither fix depends
  on a prefill that no longer exists — the stale comment saying otherwise was corrected in the same
  change.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of NEW warnings (the pre-existing unused
  `file` warning at Form.tsx:517 predates today), packaged as `1.0.265.0`. **Not yet re-tested live** —
  add several files and confirm every field on every row is empty, and that the first file is
  unaffected (it always started blank, since `inheritDefaults` returned `{}` for an empty batch).

## THE FILE LIST RENAMED EVERY NEW FILE TO THE FIRST ONE'S VALUES (2026-08-26, 1.0.264.0)
Same session, immediately after 1.0.263.0 removed the false clash warnings: *"I select one file and
prefill the file name. Then I select more files with different name and suddenly it name the three
more files test test test."* Four genuinely different documents
(`TPV2409-13-Clarence.pdf`, `TPV2409-13-Clarence-PAID.pdf`, `TPV2412-11-Clarence.pdf`, …) all listed
as `TEST - TEST - 26-08-26.pdf`.
- **THE SAME ROOT CAUSE AS 1.0.263.0, showing through a different symptom — worth noting, because
  fixing the warnings made this one MORE visible rather than less.** `inheritDefaults` copies Project
  and Vendor into each newly added file (right, and what stops a ten-file batch meaning ten retyped
  vendors), so the composed name for a file with no Document Name yet is `TEST - TEST - <date>.pdf`
  for every one of them.
- **That name will NEVER be uploaded** — the blank Document Name blocks the save — so the list was
  displaying a name that could not happen, and in doing so destroyed the only thing that told the
  uploader which row was which file.
- **Fixed by showing the ORIGINAL filename until `nameSettled(sf)`**, then the composed name. The row
  answers *which file is this*; the open row's **"Saves as"** line answers *what will it become*, built
  from the live editor and shown throughout. Two different questions, two different answers — they were
  previously both giving the second one, and only one of them was right to.
- Reuses the `nameSettled` predicate added in 1.0.263.0, so the display and the clash check can never
  disagree about whether a file's name is real yet.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.264.0`.
  **Not yet re-tested live** — pick several differently-named files after filling in the first, and
  confirm each row still shows its own filename until you name it.

## UNFILLED FILES ACCUSED EACH OTHER OF A NAME CLASH (2026-08-26, 1.0.263.0)
Client, after picking four files into one batch and naming the first: *"it automatically force other
files to have a rename when I already rename the first file."* Three of the four showed *"Another
document in this batch would be saved under this same name"* before a character had been typed into
them.
- **THE CAUSE IS A CORRECT RULE MEETING A CORRECT RULE.** `inheritDefaults` copies the previous file's
  Project and Vendor into a newly added file but **never** its typed Document Name — deliberately,
  because two files sharing a name is the exact clash `collisionsWithin` exists to catch. So every new
  file starts with a blank Document Name, `composeUploadBase` drops the blank part, and all of them
  compose to the same SHORTER name (`TEST - TEST - 26-08-26.pdf` against the filled-in
  `TEST - TEST - TEST - 26-08-26.pdf`). Three unfilled files, one identical name, mutual accusation.
- **The message was also the wrong one.** Their real state is *"not filled in yet"*, which the
  incomplete-field marker beside them already reports accurately. Two markers on one row, one of them
  describing a problem the uploader cannot act on yet, reads as the form making demands.
- **Fixed by judging only files whose NAME IS SETTLED** — `nameSettled` requires Project, Vendor,
  Document Name and Date. ⚠ **Filtered on the four NAME parts, NOT on `missingForFile`:** Confidential
  Level is required to save but contributes nothing to the filename, so a blank one must not suppress
  a real clash.
- **⚠ IT CANNOT LET A GENUINE CLASH REACH SAVE.** An incomplete file blocks the save on its own, so by
  the time every file is complete every file is in the set — the check is delayed, never skipped. That
  is the property that makes this safe to filter at all, and it is why the filter belongs here rather
  than inside `collisionsWithin`, which `canSaveBatch` calls unfiltered over the saved batch.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.263.0`.
  **Not yet re-tested live** — pick four files, name only the first, and confirm the other three show
  only the incomplete-field marker; then give two of them the SAME name and confirm the clash still
  fires.

## A REFUSED FILE'S REASON WAS HIDDEN BEHIND A COLLAPSED CARD (2026-08-26, 1.0.262.0)
Client, after a deliberate name-clash test: *"can you automatically make the dropdown show so I can
tell which file is wrong? Right now I cannot tell unless I manually click the dropdown."*
- **THE PAGE TOLD THEM TO ACT ON INFORMATION IT WAS HIDING.** Saved batches are collapsed by default
  — correct, since a saved batch is settled work — but the footer after a partial run reads *"fix the
  reason shown and press Upload again"*, and the per-file reason lives INSIDE the card. So the one
  state where the card's contents matter was the one state where it stayed shut.
- **Fixed at the results fold in `uploadAll`**: any batch left holding a per-file `error` is expanded.
  Reads the same `StagedFile.error` that `applyUploadResults` writes on failure and clears on a
  survivor, so the expansion and the message can never disagree about which batch is at fault.
- **Nothing is ever CLOSED here, and that is deliberate** — the new state is MERGED into
  `openBatches` rather than replacing it, so a batch the uploader opened themselves stays open and a
  clean batch is left exactly as they had it. Replacing the map would have made a successful run
  quietly collapse someone's expanded card.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full test suite green,
  packaged as `1.0.262.0`. **Not yet re-tested live** — stage a deliberate clash, upload, and confirm
  the failing batch opens itself with the reason visible.
- **⚠ NOT changed, on purpose: the CROSS-BATCH clash is still only caught at UPLOAD time.**
  `collisionsWithin` is per-batch by definition, so two separate batches sharing a destination and a
  composed name are not checked at Save. Examined and left alone (2026-08-26) after watching it
  behave: `overwrite=false` plus existence checks on BOTH the approval and approved sides mean
  nothing is overwritten, the refusal names the file and offers two real options, the file stays
  staged, and the footer states nothing is sent twice. Late feedback, not an unsafe outcome — and
  worth DEMOING, since "what if two people upload the same filename" is a question clients ask.

## C-LEVEL GROUP NAMING CHANGED, AND THE GLOBAL ONE HAD NO ROUTE TO EXISTING AT ALL (2026-09-02)
Client, urgently: *"Groupd naming for C-level should not be called Segview but rather GHO_C_LEVEL.
Also there isn't any global viewer for some reason, C_LEVEL_GLOBAL."* Two changes, one of them a real
bug this session's own earlier edit had just made worse.
- **`_SEGVIEW` → `_C_LEVEL` (segment-scoped C-Level).** `CANONICAL_SUFFIX.SEGVIEW` now writes
  `_C_LEVEL`; the ROLE code stays `SEGVIEW` (internal, never client-facing) and `_SEGVIEW` stays in
  `ROLE_SUFFIXES` so an already-provisioned `GHO_SEGVIEW`/`MHO_SEGVIEW` still parses — the same
  additive pattern as every prior suffix rename (`_HOD`, `_VIEWER`, `_UPLOADER`/`_APPROVER`).
  `canonicalGroupRename` picks the new suffix up for free (it derives from `CANONICAL_SUFFIX`), so
  Group Management's existing "Group names are not consistent" bulk-rename card offers `GHO_SEGVIEW`
  → `GHO_C_LEVEL` with no separate UI change.
- **`"GLOBAL"` → `"C_LEVEL_GLOBAL"` (site-wide C-Level).** This is a LITERAL, not a stem+suffix — one
  group, no segment prefix, because GLOBAL by definition reaches every segment and a segment-prefixed
  name would misstate that. Confirmed with the client after their two messages disagreed (first
  `GHO_C_LEVEL_GLOBAL`, then the corrected `C_LEVEL_GLOBAL`). `suggestGroupName`, `roleFromGroupName`
  and `canonicalGroupRename` each special-case the literal separately from the suffix table — the old
  bare `"GLOBAL"` still parses and is now itself offered a rename to `"C_LEVEL_GLOBAL"` through the
  same bulk-rename card.
- **⚠ "THERE ISN'T ANY GLOBAL VIEWER" WAS A REAL BUG, NOT A MISUNDERSTANDING — and this SAME SESSION
  had just caused it.** `clevel_global` was NEVER in `BulkGroupProvisioner.tsx`'s `OFFERED` list, so
  bulk provisioning could never create it; the only other route was the "Create one group" hand-form
  on Group Management, which was removed earlier THIS session (client: *"the group creation is done
  automatically, remove Create one group"*). Together that left `C_LEVEL_GLOBAL` with **no path to
  existing at all** — confirmed by the client's own screenshot, where the persona picker's site-wide
  card explains how administrators (a different mechanism, `CRS Owners`) work but no group for it
  exists anywhere in the 684-group list.
  - **FIXED: `clevel_global` added to `OFFERED`, ticked by default.** No bespoke "create once,
    site-wide" logic was needed — `planBulkGroups`'s existing name-based idempotence already handles
    it for free: `suggestGroupName` returns the same bare `"C_LEVEL_GLOBAL"` regardless of which
    segment's run creates it, so a GHO run creates it and an MHO run afterwards just MAPS the
    already-existing group rather than duplicating it. `buildGroupMapRow`'s termless-forcing for
    `role === "GLOBAL"` (Segment/UnitTermGuid both blank) is likewise already correct and untouched —
    `BulkGroupProvisioner.tsx`'s `rowsFor` already calls it, it was simply never reached because the
    persona was never offered.
- **Verified**: `tsc --noEmit` clean, full suite **1598/0** (2 new tests: `canonicalGroupRename` now
  renaming both the old segment suffix and the old global literal, plus 5 existing assertions in
  `groupMapModel.test.ts`/`bulkGroups.test.ts` updated for the new expected names), no new lint
  warnings. **NOT yet deployed or site-tested** — the next bulk-provisioning run against a real
  segment with `clevel_global` ticked should create `C_LEVEL_GLOBAL` once and map it (not duplicate
  it) on a second segment's run.

## GROUP-NAME SUFFIXES STANDARDISED, AND A 30-FILE CAP PER BATCH (2026-08-26, 1.0.261.0)
Client, looking at Group Management: *"the APR change to APPROVER, the UPL change to UPLOADER, the
Employee change to Viewer"*, plus *"ensure maximum file per batch in upload form is 30 files"*.
- **THE ASK WAS A REAL INCONSISTENCY, NOT A LABEL PREFERENCE.** On one unit the list read
  `..._APPROVER` beside `..._APR_HIGHLY_CONFIDENTIAL`, and `..._UPLOADER` beside
  `..._UPL_HIGHLY_CONFIDENTIAL` — the HC variants abbreviating where the plain ones spelled out — while
  the plain viewer said `_EMPLOYEE` and its HC twin already said `_VIEWER_HIGHLY_CONFIDENTIAL`. One
  role reading as two different words depending on clearance.
- **`CANONICAL_SUFFIX` now writes** `_UPLOADER_HIGHLY_CONFIDENTIAL`, `_APPROVER_HIGHLY_CONFIDENTIAL`
  and `_VIEWER`. **⚠ EVERY OLD SPELLING STILL PARSES** — `_UPL_HIGHLY_CONFIDENTIAL`, `_UPL_HC`,
  `_APR_HIGHLY_CONFIDENTIAL`, `_APR_HC`, `_EMPLOYEE` — exactly as `_DEPARTMENT_VIEWER` and `_DEPTVIEW`
  were kept when `_HOD` arrived. **Dropping one would be silent and severe:** `roleFromGroupName`
  FALLS THROUGH to MEMBER for anything unrecognised, so an existing approver or HC uploader group
  would reclassify as view-only, invisibly, in the mapping list. Two new tests pin both directions.
- **The length sort is what keeps them apart, and it holds:** `_APPROVER_HIGHLY_CONFIDENTIAL` (29)
  and `_UPLOADER_HIGHLY_CONFIDENTIAL` (29) are tested before `_APPROVER`/`_UPLOADER` (9), and
  `_VIEWER_HIGHLY_CONFIDENTIAL` (27) / `_VIEWER_HC` (10) before `_VIEWER` (7). A plain `_VIEWER`
  could not swallow them anyway — matching is on the END of the name.
- **⚠ EXISTING GROUPS KEEP THEIR OLD NAMES, so a site ends up MIXED until someone renames them.**
  This is safe but must be said: bulk provisioning checks the NAME first, misses, then falls back to
  matching on the TERM GUID — finds the `_EMPLOYEE` group holding `[MEMBER]` at that term, and reports
  `already there as GHO_..._EMPLOYEE` rather than creating a duplicate. That is the 1.0.162.0
  rename-recovery path doing its job. Renaming the groups is optional tidy-up; a SharePoint group
  rename preserves its Id, so every mapping row survives, though `GroupName` on those rows goes stale.
- **`MAX_FILES_PER_BATCH` in `Form.tsx`**, enforced in `acceptFiles`. There was **no ceiling at all**
  before — the mockup named one and it was never built. **30 on 2026-08-26, lowered to 20 on
  2026-08-27** at the client's request (*"client wants it to be 20 per batch"*). Every user-facing
  message is derived from the constant, so the number lives in exactly one place — change it there
  and nothing else needs touching.
  - **COUNTED AGAINST WHAT IS ALREADY STAGED**, not against the selection. Checking the pick alone
    would let six picks of twenty through, which is how anyone actually reaches thirty.
  - **TAKES WHAT FITS rather than refusing the lot**: a pick of forty stages thirty and the toast
    names both numbers. Refusing all forty would lose the twenty-nine that were fine and leave the
    uploader re-picking with no idea which mattered. Never silent.
  - Per BATCH, not per upload — the toast says to save and add another batch, so the way on is stated.
  - **NOT shared with `BulkUpload`'s `MAX_FILES` (50)**: different screen, different question, and the
    client set the two numbers separately.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, **full test suite green**
  (`groupMapModel` 152, `bulkGroups` 24, `userAccess` 22 — the five tests that pinned the old
  spellings were updated and two added). Packaged as `1.0.261.0`. **Not yet re-tested live.**

## A FILE TYPE CAN BE REMOVED NOW, NOT ONLY SWITCHED OFF (2026-08-26, 1.0.260.0)
Client, having added `.text` by accident while testing the add fix: *"I think best to give them
delete."* Fair — the page could add a type and toggle it, but never remove one, so a typo was
permanent. Toggling off blocks it, but the junk row stays in the list for ever.
- **⚠ UNTICK FIRST, THEN REMOVE THE CHOICE — the order IS the safety.** The list on screen is the
  column's `Choices`; what actually permits an upload is the ITEM's stored value. Drop a choice that
  is still ticked and SharePoint can leave the stale value behind in the item, so the type goes on
  being **ALLOWED with no row on screen showing it** — on the one page whose whole job is stating
  upload policy. Reversed, the worst case is harmless and visible: the untick lands, the `Choices`
  write fails, and the type is left blocked-but-listed.
- **It is ONE action, deliberately, rather than "switch it off first, then remove".** A two-step an
  admin can get half-way through is exactly how the dangerous ordering happens by hand.
- **A BLOCKED type gets no Remove**, and that is not an oversight: those rows ARE the guard — the
  list is how an admin sees that `.exe` is refused by policy. Removing the row removes the evidence
  while `BLOCKED_TYPES` goes on refusing it anyway.
- **Plain confirm, not a typed one.** Removing a type is reversible in one click (add it again), and
  a typed gate on a reversible action is what teaches people to type through gates that are not.
  The dialog does call out the one consequential case — removing the LAST allowed type blocks every
  upload site-wide — the same fact the toggle's own last-one-off confirm exists for, repeated because
  Remove is a different button and an admin may only ever meet this one.
- **The read-only fallback panel is pinned to three columns**, since the shared grid gained a fourth
  for the action and that panel can never carry one.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as `1.0.260.0`.
  **Not yet re-tested live** — deploy, remove `.text`, and confirm it disappears from the list and
  the audit log records a `PolicyChanged` row naming it.

## ⚠ ADDING A FILE TYPE ON THE CRS CONFIGURATION PAGE 400'D — FIXED (2026-08-26, 1.0.249.0)
Found live while preparing the CRS Settings demo: `+ Add file type` on `FileTypeSettings.tsx` failed
with **HTTP 400** *"Parsing JSON Light feeds or entries in requests without entity set is not
supported"* on every attempt. Toggling an EXISTING type on/off worked fine — only ADDING a new one
was broken.
- **THE SAME TRAP THE AUDIT LOG ALREADY PAID FOR, NEVER PATCHED HERE.** `writeChoices` MERGEs the
  FIELD's `Choices` (not the item), which needs the `__metadata` envelope — verbose-only, so this call
  is genuinely `odata=verbose` on both headers, unlike every other write on this page. SPFx's
  `SPHttpClient` injects `odata-version: 4.0` on EVERY request regardless of what Accept/Content-Type
  say, and OData 4 does not agree with the OData 3 verbose dialect this write needs — the exact
  collision `spAuditLog.ts`'s `WRITE_HEADERS` comment already names in detail (2026-08-13). That fix
  never touched `FileTypeSettings.tsx`, because it is the only OTHER call site in the codebase that
  needs a genuinely verbose write (a Field's `Choices`, not an item or a list) — every other verbose
  call this project makes is a different shape.
- **Fixed with the same override**: `"odata-version": ""` added to `writeChoices`'s headers, removing
  the injected 4.0 that was colliding. Toggling (`writeTicked`) was never affected — it is
  `odata=nometadata` throughout, the same shape as the audit log's write.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, packaged as 1.0.249.0.
  **Not yet re-tested live** — deploy and retry adding a type (e.g. `.txt`) before the client demo.

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

## ⚠⚠ A REVOKED RECIPIENT WENT ON READING `has access` — PRESENCE IS NOT ACCESS (2026-08-28, 1.0.315.0)
Client, having revoked Crystal and confirmed from her own account that she genuinely lost the file:
*"in this page where HOU and HOD sees, it still shows crystal have access and needs to be revoke, it
should remove crystal from the label since she have no access anymore."* Two defects, one of them
the more serious.
- **⚠ `readAcl` NEVER ASKED WHAT LEVEL A PRINCIPAL HELD.** It selected `RoleAssignments/Member/*`
  only, so `mergeShareAcl` marked a recipient `live` purely because their name appeared in
  `RoleAssignments`. **Being listed there is not the same as having access.** SharePoint adds an
  automatic **Limited Access** entry (role def `1073741825`, `RoleTypeKind` 1 = Guest) wherever a
  principal was granted on a child item, and leaves one behind after some revocations — it exists so
  a URL path resolves and it lets nobody open, list or search anything.
  - **THE TRAP WAS ALREADY WRITTEN DOWN IN THIS FILE, for the Site Pages list:** *"almost every entry
    on that list is `Limited Access` and grants NOTHING… only a real `Read` binding lets anyone open
    an inheriting page."* Same mechanism, new screen, and the new screen is the one that answers
    **"who can reach this document"** — so the wrong answer here is a disclosure, not a nuisance.
  - **`holdsRealAccess` in `shared/requests.ts`** now discards a principal whose bindings are ALL
    Limited Access, matched on **both** the role-definition id and `RoleTypeKind` — either signal
    alone would be a single point of failure on a value nobody controls.
  - **⚠ IT FAILS TOWARDS SHOWING THE GRANT, and the asymmetry is deliberate.** `bindings` **absent**
    means the read did not ask (an older caller, an unanticipated response shape) ⇒ answer TRUE,
    exactly as an absent `principalType` is read as a user. **`[]` is treated the same way**, because
    an empty array here is far more likely to be a shape problem than a principal holding literally
    nothing. Only bindings that were READ and are ALL Limited Access answer false. On this screen an
    extra name is a question somebody asks; a missing one is a disclosure nobody notices.
  - It also stops a Limited Access entry being reported as an **`unrecorded`** recipient
    (*"granted outside CRS"*) — which would have invented a share that never happened.
- **REVOKED RECIPIENTS ARE NO LONGER LISTED — `currentlyShared`, and this REVERSES one line of the
  2026-08-28 design.** That line kept them as the record that access had been granted and taken back.
  The client is right that the **Shared files** tab asks a present-tense question, and somebody who
  cannot reach the file is not part of that answer. Nothing is lost: the revoke still appends to the
  request's `DecisionNote`, the row still turns `Revoked` when the last recipient goes, and the
  **Share** tab's Decided section still shows it. History belongs there.
  - **⚠ `unknown` IS KEPT, and that is the whole care in the function.** Hiding a file whose ACL could
    not be read would let one throttled request quietly shorten the list of who can reach the
    company's documents. Only `revoked` — a definite answer — is dropped.
  - A file left with nobody leaves the tab, and the empty state now distinguishes *"nothing was ever
    shared"* from *"everything shared has been revoked"*. One message for both would contradict the
    Share tab sitting beside it.
- **⚠ THE TWO DEFECTS MASKED EACH OTHER, which is why the report read as "revoke does not work".**
  Fixing only the display would have hidden a recipient the code still believed had access; fixing
  only the read would have left a correct-but-unwanted `access ended` row. Both were needed.
- **Verified**: `tsc --noEmit` clean, full suite green (**126 tests in `requests.test.ts`**, 17 new,
  including a regression test for the Limited-Access recipient), 34 warnings — the pre-existing set,
  none new. Packaged as `1.0.315.0`. **NOT yet tested live.**

## `Limited Access — Given directly` ON A LIBRARY IS NOT A LEAK (2026-08-28)
Client, having opened `Highly Confidential Document` → Permissions and found only `CRS Owners`, then
run Check Permissions on Crystal: *"it is suppose to not allow CRS_SITE_MEMBERS to have access but
crystal can have access despite being CRS_SITE_MEMBERS… Is it because crystal is given direct
access?"* **Yes — and the dialog says so exactly: `Limited Access` / `Given directly`.**
- **IT IS NOT A GROUP GRANT, AND THE RECONCILIATION LINE IS CORRECT.** `✓ … CRS_SITE_MEMBERS has no
  access — correct` and a library ACL listing only `CRS Owners` are both true. Limited Access is
  SharePoint's **automatic traversal entry**, created when a principal is granted on an ITEM inside
  the library so the URL path to that item resolves. It cannot be granted by hand and cannot be
  removed while the item grant exists.
- **IT GRANTS NOTHING ON ITS OWN.** No browsing, no file list, no search — the one file, by direct
  URL. Already verified on this site (2026-08-21): the recipient's file opened, the parent folder
  answered *"Unknown render failure"* and the library root read *"This folder is empty"*.
- **⚠ SO THE CHECK IS "WHAT CAN SHE SEE", NEVER "IS SHE LISTED".** *Can open the library page* is not
  *can see documents*. If a shared-with account ever sees a FILE LIST there, that is a different and
  real problem — Limited Access cannot produce it.
- **⚠ AND IT IS THE SAME FACT BEHIND THE REVOKE BUG ABOVE.** A stale Limited Access entry is
  indistinguishable from a real grant unless the role-definition bindings are read. The Check
  Permissions dialog is the cheap live diagnostic: it names the LEVEL, where the library's own
  Permissions page only names groups.
- **⚠ A POLICY POINT FOR THE CLIENT, NOT A DEFECT: an approved share hands a named person one HC
  document regardless of clearance.** Crystal is in `CRS_SITE_MEMBERS` only and holds no HC role, and
  a share request approved by an HC-cleared Head of Unit reaches her anyway — because that is what
  sharing IS. The HC library split governs GROUP access; per-file shares sit outside it by design.
  Worth stating plainly rather than letting it be discovered.

## FINDING AND REMOVING AN INDIVIDUAL'S LEFTOVER ACCESS TO A LIBRARY (2026-08-28)
Crystal could open `Highly Confidential Document` after her share was revoked — page loaded, **no
folders visible**. That is `Limited Access` and it is not a leak; it is a leftover traversal entry.
The route to find and remove it, learned the long way.
- **⚠ "LIST THE ITEMS WITH UNIQUE PERMISSIONS" IS USELESS IN A CRS LIBRARY — this was suggested and was
  wrong.** Reconciliation breaks inheritance per DEPARTMENT and per UNIT folder, so **~130 of the ~130
  folders answer `HasUniqueRoleAssignments: true` by design**. Only the below-Unit tiers (`2024`,
  `Term Sheet`) inherit. The flag cannot separate a share from ordinary provisioning here.
- **THE ROUTE THAT WORKS: library → Permissions → "Show users" in the yellow banner.** That reveals
  the Limited Access principals the Permissions page hides. Equivalent REST, one call:
  `/_api/web/lists(guid'<id>')/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/LoginName,RoleDefinitionBindings/Name`
- **⚠⚠ THE `Type: User` ROWS ARE THE ONLY INDIVIDUAL GRANTS. EVERYTHING ELSE ON THAT LIST IS A GROUP
  AND MUST NOT BE TOUCHED.** Every `*_HIGHLY_CONFIDENTIAL` group shows `Limited Access` too — the
  automatic entry for a principal granted on folders BELOW — so ticking the header checkbox and
  pressing **Remove User Permissions** would take every group's folder access away across the whole
  library, with no confirmation and nothing on screen to say what happened. **Same hazard as
  `groupsToRemove`**, which is page-scope only for exactly this reason (~308 such entries on the
  approval library). A recovery run would repair it, but every uploader is locked out meanwhile.
- **⚠ ORDER MATTERS: REVOKE THE ITEM GRANT FIRST, THEN THE LEFTOVER.** Removing Limited Access while a
  file grant still stands strips traversal and leaves the grant orphaned — invisible from the library
  page, and actively misleading now that the Requests page reads the file's own bindings. Only remove
  it once the share reads `access ended`.
- **⚠ AND THIS IS THE ONLY PLACE A NATIVE SHARE IS VISIBLE.** A document shared with SharePoint's own
  Share button leaves NO request row, so the Requests page cannot list it — settled 2026-07-23 and
  stated on that screen. The `Type: User` rows here are the real index of individual grants, native
  ones included. **Use this list when the question is "who has been given a document in this
  library", not the Requests page.**
- Verify with **Check Permissions** on the person: gone means no row at all, not `Limited Access`.
  That dialog names the LEVEL, where the library's Permissions page names only groups.
- ⚠ The same entries exist at WEB scope (`/_api/web/roleassignments`). Clear them there too if the
  person should be fully gone — but **never remove their `CRS_SITE_MEMBERS` Read**, which is their
  legitimate site entry and taking it removes them from the DMS entirely.

## A REVOKE NOW RECORDS WHO REVOKED IT — `RevokedBy` (2026-08-30, 1.0.317.0)
Found while adding the `Revoked` branch to `CRS — Audit request activity`. The branch alone would
have shipped a row **correctly labelled and wrongly attributed**, which on an audit log is worse than
no row at all: a wrong name gets believed. Client: *"we better track HOD since that is the whole point
of Audit log, we do not want to give a defect to client."*
- **⚠ THE CAUSE IS A DELIBERATE OMISSION MEETING A REASONABLE ASSUMPTION.** `performRevoke` writes
  only `DecisionNote` (appended) and `Status = Revoked` — it leaves `DecidedBy` alone ON PURPOSE,
  because that field records who let the share through and overwriting it destroys half the trail.
  The flow's `ActorEmail` picks `RequestedBy` on Pending and **`DecidedBy` otherwise**, which is right
  for approve and reject and wrong for revoke. **A Head of Department can revoke a share a Head of
  Unit approved**, and the log would have credited the Head of Unit.
- **`RevokedBy` IS A THIRD FIELD, NOT A REUSE.** Who granted access and who withdrew it are different
  facts about one request, and both are now shown — `Approved by X` above `Revoked by Y`.
- **⚠ WRITTEN ON EVERY REVOKE, not only the one that closes the request.** A partial revoke leaves
  `Status = Approved`, so gating the write on `ended` would leave the field blank until the last
  recipient went — and the value would then name whoever happened to remove the final one rather than
  the person who acted. Last writer wins, which is what "who acted last" means.
- **⚠ THE READ IS A THREE-RUNG LADDER NOW** (`Stage,RevokedBy` → `Stage` → neither). One unknown
  name fails the WHOLE `$select` (gotcha #11), so the single retry that covered `Stage` was not
  enough: a site holding `Stage` but not `RevokedBy` would have fallen through both rungs and **lost
  the stage as well**, reporting every pending-file request as an approved-document one. Same shape
  as `readLibrary`'s ladder in My Submissions.
- **⚠ AND THE WRITE IS GATED ON THAT READ.** Attempting a column the list lacks fails the whole
  MERGE — which would lose the revoke's `DecisionNote` line too, and that line is where the
  revoker's name currently lives on older sites. **Degrade to prose; never lose the record.**
  `revokedByMissing` may be OVERSTATED (a site missing only `Stage` lands on rung 3 with both flags
  set); deliberate, because a needless skip costs one line of display where a wrong guess costs the
  whole write.
- **MIGRATION: the column is added by the Requests page's own `addMissingColumns`**, not by
  reconciliation — same as `Stage`. Until an admin opens that page and adds it, revokes behave
  exactly as they did before: the label is right and the actor is the approver.
- **⚠ THE FLOW NEEDS TWO EDITS, NOT ONE.** `EventKind` gains the `Revoked` branch AND `ActorEmail`
  must prefer `RevokedBy`. Doing only the first is the defect this section exists to prevent, dressed
  up as a fix — the row would say `ShareRevoked` and still name the approver.
- **BOTH FLOW EDITS ARE DONE (2026-08-31), and they branch on `RevokedBy`, NOT on `Status`.**
  Branching on `Status = 'Revoked'` is the obvious reading and is WRONG for the **partial** revoke:
  `performRevoke` writes `RevokedBy` on EVERY revoke but only sets `Status = Revoked` when the LAST
  recipient goes. So removing one of three recipients left `Status` at `Approved` and logged a
  **second `RequestApproved`** — a claim that an approval happened, naming the approver — which the
  5-minute dedupe cannot catch, because a revoke normally happens days after the approval.
  `if(not(empty(coalesce(triggerOutputs()?['body/RevokedBy'], ''))), ...)` answers the partial and the
  full case identically. Cost, accepted: `RevokedBy` stays set for ever, so a later unrelated edit to
  that row logs another `ShareRevoked`. **Noise beats a false `RequestApproved`.**
- **VERIFY THE COLUMN EXISTS BEFORE TRUSTING THE FIX.** `RevokedBy` is created by the Requests page's
  own `addMissingColumns`, NOT by reconciliation — so on a site where that has never been pressed the
  expression reads null, `coalesce` makes it `''`, and the flow silently reverts to the old wrong
  behaviour. The cheap check is the Share tab: a decided row rendering `Revoked by ...` proves it.
  Confirmed present on ClarenceDMSTesting 2026-08-31.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings, full suite green (47 suites,
  0 failed; `requests.test.js` 126). Packaged as `1.0.317.0`. **NOT yet deployed or site-tested.**

## THE REQUESTS PAGE FILTERS AND SCROLLS — AND `Rejected` IS RED NOW (2026-08-30, 1.0.318.0)
Client, looking at a Decided list of seven and a Deletion tab of fourteen: *"Shouldn't Rejected label
be red and shouldn't there be a scrollbar to scroll through the tiring list? Also shouldn't there be a
dropdown to filter through rejected, pending, approved list? Also shouldn't there also be an email
input to filter through whose request?"* All four built.
- **`Rejected` IS SOLID RED, MIRRORING `Approved`'s SOLID GREEN.** It was grey, which put a refusal in
  the same visual register as `Cancelled` — something that merely stopped rather than something an
  approver decided against. **`Failed` stays PALE red** and must: *"the approver said no"* and *"the
  action broke"* are different facts, and one is nobody's fault.
- **⚠ THE FILTERS TOUCH THE DECIDED LIST AND NOTHING ELSE. THIS IS A SAFETY RULE, NOT A LAYOUT ONE.**
  *"Waiting for you"* IS the work. An approver who filtered to `Approved` to look something up and
  then forgot would stop seeing their own queue, with nothing on screen saying why — requests would
  sit unanswered behind a control they set themselves. **A display filter must never be able to hide
  outstanding work.**
- **The count reads `N of M` while filtered**, so a short list never reads as a missing one — and the
  empty state says the FILTER is why (`N are hidden — press Clear`), never *"nothing has been
  decided"*, which would be a false statement about the data rather than about the view. Same
  distinction this codebase draws everywhere between empty and unknown.
- **Outcome options are DERIVED from the rows present**, so a filter that yields nothing cannot be
  chosen. `Pending` is absent by construction — those rows are in the queue, never in the decided
  list.
- **Clear is not optional.** With the list scrolled, a filter matching nothing shows an empty box and
  no obvious way back; the client raised exactly that twice about Group Management's search boxes.
- **⚠ THE 60vh SCROLL IS SAFE HERE ONLY BECAUSE NOTHING IN THIS CARD IS ABSOLUTELY POSITIONED.**
  Group Management's people picker, the upload form's Confidentiality/Legally-Privileged panels and
  the member-add dropdown have each been clipped by a scroll container — three screens, same trap.
  **Check for popovers and people pickers BEFORE adding a scroll box**, and lift the cap if one is
  ever added here.
- **Filters sit OUTSIDE the scrolling box**, the same rule the abbreviation editor follows with its
  Save button: a control must not scroll away from the rows it governs.
- **⚠ `&rsquo;` IN A JSX *ATTRIBUTE* IS NOT DECODED** — a placeholder written that way renders the
  literal entity. Entities work in JSX text, not in attribute values; caught before shipping.
- **Verified**: `tsc --noEmit` clean, `heft build` clean of new warnings on this file, full suite green
  (0 failed). Packaged as `1.0.318.0`. **NOT yet deployed or site-tested.**
- **ALL FOUR TABS NOW, not two (1.0.319.0).** Client: *"i notice you only apply it to Share, can you
  apply it to the rest as well?"* — **Deletion already had it**, since it and Share render through the
  same `typeTab`. `Shared files` and `Your requests` are separate blocks and genuinely did not.
  - **⚠ THE CONTROLS ARE TAILORED PER TAB, NOT COPIED.** `Your requests` gets NO requester filter:
    every row there is already yours, so the control would match everything or nothing. **Copying it
    for visual consistency would be noise wearing the costume of a feature.** It gets document name
    instead — the question that tab is actually opened with.
  - **`Shared files` filters on an ACCESS STATE (`live`/`unknown`), not a request status**, so it needs
    its own state: a shared control would carry a value the other tab cannot mean, and a filter
    silently matching nothing is exactly how a list reads as empty.
  - **⚠ ITS TEXT BOX MATCHES DOCUMENT NAME *OR* ANY RECIPIENT EMAIL**, because the question people
    arrive with is *"does <person> still have access to anything?"*, which a filename-only filter
    cannot answer. Labelled to say so — a box matching two things silently is worse than one matching
    neither.
  - **⚠⚠ A MATCHED FILE'S RECIPIENT LIST IS NEVER NARROWED.** Filtering the people INSIDE a file would
    hide somebody who DOES have access, on the one tab whose whole job is answering who can reach a
    document — a disclosure-shaped error, not a display one. The filter picks which FILES show; it
    never edits what a file says about itself.
  - Filter styles and the 60vh scroller are ONE definition (`FIL_*`, `SCROLLER`): three bars that
    drift apart look like three controls doing three different things.
  - **Verified**: `tsc --noEmit` clean, no new lint warnings (the file is 1732 lines, under the 2000
    ceiling), full suite green. Packaged as `1.0.319.0`. **NOT yet deployed or site-tested.**
- **⚠ "WAITING FOR YOU" GOT THE SCROLL AND A TEXT FILTER — BUT NOT THE OUTCOME DROPDOWN (1.0.320.0).**
  Client asked directly whether the queue had them. It had neither; the scroll was an oversight, the
  filters were not.
  - **THE DROPDOWN IS STILL WITHHELD, AND THE REASON IS ARITHMETIC RATHER THAN CAUTION:** every row in
    that queue is `Pending`, so an outcome filter cannot narrow it — it can only **empty it
    wholesale**, which is the one outcome no display control should be able to cause. The bar SAYS
    *"Outcome applies to decided requests only"*, because without that line an approver sets it to
    `Approved`, watches the queue stay put, and reads the control as broken.
  - **The text filter IS applied**, with `N of M` and a Clear beside it — and on this list that count
    is load-bearing in a way it is not on the decided one. **This queue is the WORK**: somebody who
    typed a name, forgot, and saw a short queue would believe they were up to date while requests sat
    unanswered. The count is what makes filtering it safe at all.
  - **The empty state says the hidden rows STILL NEED A DECISION**, never just "no match" — the
    difference between a view being narrow and work being done.
  - **ONE bar above BOTH sections now, not one per card.** Two boxes both labelled "requester email"
    on one screen is a puzzle, and a single Clear puts the whole tab back.
  - The text filter matches **document name OR requester email** on both sections, one rule and one
    label, rather than the earlier requester-only match that differed from the Shared-files box.
  - **Verified**: `tsc --noEmit` clean, full suite green, 1744 lines (under the ceiling). Packaged as
    `1.0.320.0`. **NOT yet deployed or site-tested.**

## MY SUBMISSIONS NOW HIDES THE DELETE/SHARE BUTTONS INSTANTLY, NOT AFTER A REFRESH (2026-09-02)
Client: raising a deletion request left the Delete button showing on that row until a manual page
refresh — it needed a reload to show "deletion asked" and hide the buttons. **Built, not yet
site-tested.**
- **CAUSE: `myRequests` (the per-file index of the requester's own most-recent request, keyed on
  lowercased `uniqueId`) was only ever populated by the mount-time read.** Both the row-list badge/
  button-hiding and the detail view's `openRequest` read this SAME state, so nothing updated either
  until the next full load re-read `CRS Requests` from scratch.
- **Fixed in `submitRequest`**: right after a successful POST, `res.json()` (the created row
  SharePoint's own POST response already returns) supplies the new item's `Id`, and a `MyRequest`
  object is written straight into `myRequests`/`myRequestList` — no extra request, no re-read of the
  whole list. Wrapped in its own try/catch that changes nothing on failure: the next load (or the
  focus-triggered live refresh this page already has) still catches up from the server exactly as it
  did before this existed.
- **One update feeds both places** — the row list and the detail view read the identical state, so
  there is no second definition of "is a request pending" to drift between them.
- `tsc --noEmit` clean, full suite **1597/0** (this file has no unit tests of its own; it's a
  component, not pure logic).

## MY SUBMISSIONS HAS A MANUAL REFRESH BUTTON, NOT JUST THE PAGE RELOAD (2026-09-02)
Client, testing the "Replaced" audit-actor fix below: *"it would be great if you add a refresh
button like the one in audit log to make it easier to refresh instead of refreshing the entire
page."* Mirrors the Audit Log's icon-only header refresh (`↻`), beside the `My Submissions` heading.
- **The mount-time `load` async function was moved OUT of its `useEffect` into a plain component-body
  `const`**, so a new `onRefresh` handler can call the exact same read pipeline (both libraries, the
  HC pair, the archive pair, the `CRS Submissions` record merge, and `loadPolicy` at the end) — one
  definition, two callers, rather than a second copy that drifts.
- **⚠ A FAILED MANUAL REFRESH DOES NOT CLEAR `rows`, UNLIKE THE MOUNT LOAD.** The mount path's
  `setRows(undefined)` on failure exists because there is nothing on screen yet to protect; wiping an
  already-loaded list because one refresh attempt hit a transient error would throw away good data
  over nothing. `onRefresh` only sets `loadError` on failure — the last successfully loaded rows stay
  visible, with the error banner explaining why they might be stale.
- **`.finally` is unavailable on this tsconfig target** (CLAUDE.md #3, same class as
  `Promise.allSettled`) — `setRefreshing(false)` is set on both the `.then` and `.catch` branches
  instead of a single `.finally`.
- `tsc --noEmit` clean, full suite **1597/0**, no new lint warnings (the file's pre-existing
  `max-lines` warning grew from the added lines but was already over the 2000 ceiling before this).

## A REPLACED RECORD'S "Replaced" AUDIT ROW NAMED THE APPROVER, NOT THE REPLACER — FIXED ON BOTH FLOWS (2026-09-02)
Client, reading the audit log: *"I notice that approver is the one who is recorded as the one who
replace, its actually suppose to be the uploader... it should show reene as the one who replace
despite it being ckyan who approve it."*
- **THE CAUSE: `Create_item_2`'s `ActorName`/`ActorEmail` reused the exact same
  `ApprovedBy`-falling-back-to-`Editor` expression the "Approved" row uses** — right for who approved
  it, wrong for who replaced it. The replacer is whoever chose "Send for approval as a replacement" at
  upload, i.e. the SOURCE item's own `Author` — the same field the "Uploaded" row already reads, and
  needing no fallback (an uploaded item always has an Author).
- **FIXED in both `Auto-route` and `HC Auto Route`**: `ActorName`/`ActorEmail` now read
  `triggerOutputs()?['body/Author/DisplayName']` / `.../Author/Email'` directly, with the
  `if(empty(coalesce(...)))` expression removed entirely. The action's position (still after
  `Get_ApprovedBy_HTTP` and `GetSourceUniqueId`, for the `ItemUniqueId` dependency) is unchanged.
- **VERIFIED LIVE on `Auto-route`**: cross-person test (uploader `chocheetuck4`, approver
  `clarencechojinheng`) — "Replaced" correctly named `chocheetuck4`, distinct from "Approved"/"Moved
  to …" (`clarencechojinheng`).
- **VERIFIED LIVE on `HC Auto Route`** the same way, same result.
- **⚠ THE SIBLING BUG WAS ALSO FOUND AND FIXED: `StampReplacedRecord`'s `ReplacedBy` (the field behind
  My Submissions' "Cancelled — Replaced by a newer upload from X" message) had the identical
  ApprovedBy/Editor-fallback expression, live on `Auto-route` already.** Same reasoning, same fix:
  `"ReplacedBy": "@{triggerOutputs()?['body/Author/Email']}"`. Applied to `Auto-route`'s existing
  action and built correctly into `HC Auto Route`'s new one from the start.

## ⚠⚠ EVERY RECONCILIATION RUN CREATED 14 DUPLICATE COLUMNS ON `CRS Submissions` (2026-09-03, 1.0.383.0)
Found by reading the list through REST after a run reported `14 column(s) created` — a count that
should only ever appear once in a list's life. The live field list held **`SubmissionRef` through
`SubmissionRef15`, `SubmissionFileId` through `SubmissionFileId15`, and `ArchivedAt`, `ArchivedAt0`,
`ArchivedAt1`, `ArchivedAt2`** — sixteen runs, fourteen junk columns each, ~224 in total.
- **⚠ THE FALSE ASSUMPTION: "a duplicate column comes back 400 and is success by omission."** The
  comment in that loop said so explicitly and it is WRONG. **SharePoint permits duplicate DISPLAY
  names and silently derives a unique INTERNAL name by appending a number** — so the POST SUCCEEDS,
  `r.ok` is true, and the counter increments. Every run created a fresh set.
- **⚠ IT NEVER BROKE ANYTHING, WHICH IS EXACTLY WHY IT RAN SIXTEEN TIMES.** Reads and writes name the
  first, unsuffixed column, so submission records worked perfectly throughout. **The only symptom was
  a number in the run log that nobody had a reference for** — the same shape as the 2026-08-24 bulk
  provisioning defect, where `0 failed` was reported over work that was quietly incomplete. Left
  alone it ends at SharePoint's per-list column ceiling, where creation starts failing for real.
- **THE FIX READS THE EXISTING FIELDS FIRST AND SKIPS THEM**, which is what `Requests.tsx`'s
  `ensureColumns` has always done via `readFieldNames` — **that list was never affected**. One naive
  loop and one safe loop in the same codebase; the safe one was written second and the first was
  never revisited.
- **⚠ FAILS CLOSED: an unreadable field list creates NOTHING and says so.** A missed create is
  repaired by the next run; a blind create is permanent junk that cannot be undone safely, because
  **deleting a column takes its data with it and does NOT reach the recycle bin.**
- **THE DUPLICATES ARE REPORTED, NEVER DELETED**, for that same reason — the run now names them and
  says they are inert and safe to remove by hand. Matched as "one of our names followed by nothing but
  digits", so a client-authored column can never be counted. **No regex is built at runtime** (lint
  forbids it, and escaping a config-supplied name would be a trap of its own).
- **⚠ SDG HAS THE SAME DUPLICATES, in proportion to how many times it has been reconciled.** Check it
  with the same REST read before its next run; the fix stops new ones either way.
- **THE DIAGNOSTIC THAT FOUND IT IS THE REUSABLE PART.** A Site Contents screenshot showed one list
  with the right rows and looked fine; Clarence asked for an API check instead — *"Screenshot itself
  is not fully reliabel"* — and `/fields?$select=InternalName` showed the whole thing in one response.
  **The standing rule of 2026-08-24 again: when a list has just taken writes, only the API is
  authoritative.**
- **Verified**: `tsc --noEmit` clean, full suite **1608/0**, `FolderManager.tsx` lint unchanged from
  its own baseline (3 `no-new-null`, 1 `max-lines`; Heft's total moved 31 → 35 only because it lints
  incrementally and this file had not changed before). Packaged as `1.0.383.0`, bundle grepped.
  **NOT yet run on a site** — the next run should report **0 columns created** and name the strays.

## THE UPLOAD FORM TITLES ITSELF AGAIN, AND THE INSTRUCTIONS FOLLOW THE OPEN BATCH (2026-09-03, 1.0.382.0)
Client: *"client wants the Upload Document title to be inside the Upload Form now"*, and the
instructions line *"to only appear when a upload form is open for each batch… if batch one is open it
will show above the batch but if its close it will disappear."*
- **⚠⚠ THE HEADING REVERSES ITS OWN REMOVAL, AND THE PAGE MUST BE EDITED OR IT APPEARS TWICE.** It was
  taken out precisely because the page carried its own title web part directly above it. So
  **deleting that text web part is a required deployment step** on every site — `Upload-Form.aspx` on
  ClarenceDMSTesting AND on SDG. **Nothing in code can detect the duplicate**: a text web part is not
  readable from inside a web part, so this will not fail, warn, or look wrong to any check — it will
  just render "Upload Document" twice until somebody opens the page.
- **ONE LINE ACHIEVES THE PER-BATCH RULE, because there is exactly ONE open card.** That is the state
  model, not a layout accident (a single set of pickers, destination snapshot at save — per-batch
  picker state is what caused the silent HC-clearance failure of 2026-08-22). The open card renders
  at one place, after every saved card, so gating the paragraph on the same `draftOpen` puts it above
  whichever batch is being filled in, with nothing to keep in sync per batch.
  - **⚠ IF MORE THAN ONE CARD CAN EVER BE OPEN, OR THE OPEN CARD MOVES, THIS BREAKS SILENTLY** — the
    paragraph would go on rendering at that position while the card it describes is elsewhere. It
    belongs immediately above the card and must move with it. Said in a comment at the site.
- **`.dms-subtitle` IS LEFT IN PLACE, UNUSED-BY-THIS-PATH BUT NOT DELETED** — it is the top-of-page
  variant, and the new `.dms-batch-intro` differs only in spacing because it now follows the
  Add-batch bar rather than a heading.
- **⚠ THE BACKTICK TRAP FIRED AGAIN WHILE WRITING THIS, FOURTH TIME.** A CSS comment referencing
  `` `.dms-subtitle` `` in backticks **ends the `<style>` template literal**, and `tsc` reports it as
  *"Property 'dms' does not exist on type…"* / *"Cannot find name 'subtitle'"* — which names neither
  the cause nor the line that matters. Caught by the typecheck this time rather than at runtime. The
  warning already in that block is now repeated beside the new rule: **no backticks anywhere inside
  it.**
- **Verified**: `tsc --noEmit` clean, full suite **1608/0**, `Form.tsx` lint unchanged from its
  baseline (the pre-existing `file`, `max-lines` and return-type warnings), shipped bundle grepped for
  `dms-page-title` and `dms-batch-intro`. Packaged as `1.0.382.0`. **NOT site-tested.**

## ⚠ AN ARCHIVED FILE READ AS `Deleted` ON MY SUBMISSIONS — CODE HALF BUILT, FLOW HALF NOT (2026-09-03, 1.0.380.0)
Client: *"CLient is complaining why all the archive files consider as deleted, can we ensure if its
archive it shows archive isnteaD?"* Runbook:
`docs/superpowers/specs/2026-09-03-archived-submission-stamp-runbook.md`.
- **THIS IS THE DIRECT COST OF THE 2026-09-02 ACCESS NARROWING, ARRIVING A DAY LATE.** My Submissions
  stopped reading the archive that day — correctly, on the client's own instruction (*"My Submission
  should not update to archive either"*) — so an archived file resolves in **no** library the viewer
  can read, `mergeRecords` judged the record unresolved, and it reported **Deleted**. Nothing was
  wrong with the merge: it was answering honestly with the only evidence it had.
- **⚠ THERE IS NO CLIENT-SIDE FIX AND IT SHOULD NOT BE RE-ATTEMPTED.** A PIC/HoU holds nothing on
  `Archive`/`Archive Highly Confidential Document`, so a probe by path, by title, or by
  `GetFileById` is **security-trimmed to 404 — byte-for-byte identical to deletion**. Only something
  that WATCHED the move can record it. Re-adding an archive read would also reverse the client's
  instruction, and would fail anyway.
- **SO IT IS A STAMP, EXACTLY LIKE `ReplacedAt`.** New `ArchivedAt` (DateTime) on `CRS Submissions`,
  written by the two archive movers; `RecordState` gains `"archived"`, which is **OBSERVED, never
  derived** — it holds whether or not the libraries could be read, and a failed read can never
  produce it. Same shape, same reasoning, second instance.
- **PRECEDENCE: `live` > `cancelled` > `archived` > `deleted`/`unknown`.** `live` first because a
  stale stamp must not hide a document the viewer can plainly see; `cancelled` above `archived`
  because the two cannot honestly co-occur (a replaced file's old content never reaches the archive
  under that stamp) and the replacement is the later fact about the record. Pinned by test.
- **⚠ THE READ IS A THREE-RUNG LADDER NOW** (full → drop `ArchivedAt` → drop the 2026-08-28
  replacement pair as well). A single retry would have taken `ReplacedAt`/`ReplacedBy` down with it
  on any site holding those but not the new column — so **every REPLACED file there would silently
  revert to reading "Deleted"**, the very state this change exists to stop showing. `RevokedBy`
  taught this on `CRS Requests` (2026-08-30); the lesson generalises: **one optional column, one
  rung.**
- **✅ THE FLOW HALF IS BUILT AND VERIFIED END TO END ON ClarenceDMSTesting (2026-09-03).** Four
  actions added to `CRS — Archive after seven years` inside its `For each`, after `Reset_inheritance`:
  `GetArchivedFileId` (reads `SubmissionFileId` off the moved item via the archive library's list
  GUID `8c28ca71-4edb-445b-8ebd-5711dacd71d4`), the `HasStamp` condition, `GetArchivedRecordId`, and
  `StampArchivedRecord` inside a loop named **`StampLoop`**. Result: four rows on My Submissions read
  **Archived** with *"in the archive"*, while genuinely deleted files still read **Deleted** — the two
  states told apart, which is the whole point.
  - **⚠ `HasStamp` IS NOT OPTIONAL AND IS THE ONE THING THAT COULD HAVE DONE REAL DAMAGE.** A file
    uploaded before the record feature carries a BLANK `SubmissionFileId`, so without the guard the
    filter becomes `SubmissionFileId eq ''` — which matches every record row with a blank stamp and
    would mark a pile of unrelated submissions as archived.
  - **⚠ FOUR TRAPS FIRED WHILE BUILDING IT, all previously documented, all silent if missed:**
    1. `StampArchivedRecord` was created as **GET**. `X-HTTP-Method: MERGE` only acts on a real POST —
       it saves clean, runs green and writes nothing. Same mistake as `StampReplacedRecord`.
    2. The `Apply to each` was never added while its Uri already referenced `items('Apply_to_each_2')`.
       **The loop is now named `StampLoop`**, so a future HC clone cannot repoint it at `For_each`.
    3. **`Get items`' `$filter` HELD A HARDCODED DATE AND IGNORED `CutOff` ENTIRELY** — a leftover from
       the 2026-09-02 test. It now reads `datetime'@{trim(outputs('CutOff'))}'`, so that Compose is the
       single place the cutoff lives. **Without this, restoring `CutOff` changes nothing** and the
       mover archives by a fixed date for ever, reporting success every time.
    4. **A TRAILING SPACE IN THE `CutOff` LITERAL** produced `Creating query failed`. The tell was in
       the error string itself: `datetime'2026-09-03T00:00:00Z '`. Fourth instance of invisible
       whitespace typed into a Power Automate field — hence the `trim()`, insurance against a failure
       mode nobody can see in the designer.
  - **⚠ THE CUTOFF MUST BE RESTORED TO `addDays(utcNow(), -2557, 'yyyy-MM-ddTHH:mm:ssZ')` AFTER ANY
    TEST**, and that matters more now than it did: the filter reads `CutOff` live, the flow is On with
    a daily recurrence, and a scheduled run against a stale literal would archive everything before
    that date silently.
  - **✅ THE HC CLONE IS BUILT AND VERIFIED THE SAME DAY.** `CRS — HC Archive after seven years` got
    the identical four actions; **exactly ONE value differs** — the archive GUID in
    `GetArchivedFileId`, `bfcd6735-7605-47e7-a0d6-bbfff41bab7b`. `CRS Submissions` is SHARED by both
    verticals and was deliberately NOT repointed: doing so would find nothing and stamp nothing,
    silently, looking identical to never having built it. Verified live — My Submissions went from
    four archived rows to **seven**, the HC files among them.
    - **⚠ `StampLoop`'s INPUT was set to the Uri expression instead of the array** on the first
      attempt — `"foreach": "@concat('_api/web/lists/...items(', items('StampLoop')...)"`, which also
      self-references the loop from inside itself. Corrected to
      `@body('GetArchivedRecordId')?['value']`. **A loop's input is invisible unless you open the loop
      itself**, and a wrong one either fails circularly or iterates zero times and writes nothing while
      the run reports success. Third instance in this project (the reminder flow's unbound `For each`,
      the audit flow's stray empty one). **Always read the `"foreach":` line in code view after
      building a loop.**
    - **⚠ ITS `Get items` FILTER HELD THE SAME HARDCODED `2026-09-01` DATE** while its `CutOff` was
      already the rolling expression — so the flow archived by a fixed date and `CutOff` did nothing.
      Both flows now read `datetime'@{trim(outputs('CutOff'))}'`.
  - **⚠ A TEST CUTOFF IS UTC AND MY SUBMISSIONS DISPLAYS LOCAL (UTC+8).** Two files shown as uploaded
    `3 Sep 02:51` were archived by a `2026-09-03T00:00:00Z` cutoff, correctly — 02:51 local is 18:51 UTC
    the previous day. Anything uploaded before 08:00 local falls into the previous UTC day, so a cutoff
    reaches roughly a working day further forward than it appears to. Do not read that as a misfire.
- **⚠ AND NOTHING BACKFILLS.** The 57 + 17 documents already archived on ClarenceDMSTesting
  (2026-09-02) carry no stamp and will go on reading **Deleted** unless those rows are set by hand.
  SDG has archived nothing, so it starts clean.
- **⚠ THE HC CLONE CHANGES EXACTLY ONE VALUE, and the trap here is the OPPOSITE of the usual one:**
  only the archive library's list GUID is swapped — **`CRS Submissions` is SHARED by both verticals
  and must NOT be**. Repointing it would find nothing and stamp nothing, silently, looking identical
  to the flow never having been built.
- **No `ArchivedBy`.** The movers are scheduled flows: there is no person, and a column naming the
  service account would read as somebody having done it deliberately to that file.
- **Verified**: `tsc --noEmit` clean, full suite **1608/0** (12 new tests — the merge precedence, the
  observed-not-derived rule, the `Invalid Date` guard, and the middle rung), 31 lint warnings
  (established baseline, zero new). Packaged as `1.0.380.0` and the shipped bundle grepped for
  `ArchivedAt` and `Moved to the archive`. **NOT site-tested, and cannot be until the flows stamp.**

## THE `CRS Submissions` RECORD STAMP (§0's "still to do") IS NOW CLONED INTO `HC Auto Route` (2026-09-02)
Completes the item flagged in the section above (a displaced HC submission previously stayed
`Deleted` instead of turning `Cancelled`, because none of this existed on the HC flow). All five
actions built and verified wired correctly:
- **`GetDestSubmissionFileId`** — inserted as the new first action inside `DestExist`'s True branch
  (before `GetSourceContent`, which was re-pointed to run after it): `GET` against
  `_api/web/lists/getbytitle('Highly Confidential Document')/items(@{first(body('GetDestMatch')?
  ['value'])?['Id']})?$select=SubmissionFileId`. Only the library title differs from the normal flow.
- **`GetReplacedRecordId`** — inside `WasReplacedHC`'s True branch, after `Create item 2`: queries the
  SHARED `CRS Submissions` list by the old `SubmissionFileId`, verbatim, no swap.
- **An `Apply to each` → `StampReplacedRecord`** — writes `ReplacedAt`/`ReplacedBy` (Author-based, per
  the fix above) onto the matched `CRS Submissions` row(s), verbatim.
- **`GetSourceSubmissionIds`** — inserted between the Author/Editor/Created stamp
  (`Send an HTTP request to SharePoint`) and the source-delete step
  (`Send an HTTP request to SharePoint 1`): reads `SubmissionId`/`BatchId`/`SubmissionFileId` off the
  SOURCE item, from the **HC approval library's own list GUID**
  (`2311cd83-90aa-4077-a647-65d251a5f426` on ClarenceDMSTesting — read from the library's Settings
  page URL, `List=%7B<GUID>%7D`, not via any script).
- **`StampDestSubmissionIds`** — runs after it, MERGEs those three fields onto the destination item
  (`outputs('DestItemId')`) in the **`Highly Confidential Document` library's own list GUID**
  (`e4fb3b2c-6f20-4bbb-a4f2-e9a5301d4080`). The source-delete step's `runAfter` was re-pointed from
  the stamp action to this one, so the source cannot be deleted before the transfer completes.
- **⚠ TWO REAL BUGS CAUGHT BEFORE THEY SHIPPED, both invisible from a green save:**
  1. **`StampReplacedRecord` was built as `GET` with no body** — syntactically valid (hence no error),
     semantically a no-op: `X-HTTP-Method: MERGE` only takes effect on an actual `POST`. Fixed by
     changing the Method dropdown to `POST` and adding the `Body` parameter.
  2. **The `Apply to each` foreach's input carried a stray trailing comma** —
     `@body('GetReplacedRecordId')?['value'],` — which Power Automate's own save validation caught
     outright (`"invalid expression(s)"`), forcing the fix before it could ship silently.
- **✅ VERIFIED LIVE 2026-09-02.** HC replacement test (uploader `chocheetuck4`, approver
  `clarencechojinheng`) — the displaced record now correctly reads **Cancelled**, "Replaced by a
  newer upload from chocheetuck4@gmail.com…", matching the normal flow's already-verified behaviour.

## THE THREE ACCESS PAGES GOT A USABILITY PASS ON TOP OF THE READ-ONLY CONVERSION (2026-09-02)
Client, live-testing the deployed read-only pages, per page: search/filter/scroll everywhere,
drop "Should have access" (misleading on a page that can no longer act on a mismatch), rename
"Access now" → "Current access". Spec:
`docs/superpowers/specs/2026-09-02-access-pages-usability-pass-design.md`. **Built via three parallel
agents, one per file — `tsc` clean, full suite 1597/1597, zero new lint warnings. NOT yet
site-tested.**
- **Approval Library Access (`StagingAccess.tsx`)**: filter box, 60vh scroll (safe unconditionally —
  `MemberSummary` is always mounted `expandable={false}` here, nothing absolutely positioned),
  "Should have access" column removed (the `rowByGroupId` map it read stays, since the `unmanaged`
  grants filter still needs it), "Access now" → "Current access" in the header, its tooltip, and one
  prose reference in the ACL-failure banner.
- **Site Access (`SiteAccess.tsx`)**: scroll on both lists, a filter box on "Everything with
  permission" (reusing `Requests.tsx`'s `FIL_*` style constants verbatim rather than a fourth
  implementation), and `simplifyLevels()` — a pure function collapsing raw permission levels into
  three buckets: **Full Control → "Full access"** (supersedes everything else on that row),
  **Limited Access / Web-Only Limited Access → one deduplicated "Reaches folder/page, no site-wide
  access"**, anything else shown as-is. Client's confirmed wording: *"Reaches folder /page no need to
  say something below, no site-wide access is good."*
- **Page Access (`PageAccess.tsx`)**: explainer paragraph removed, filter + 60vh scroll, "Should have
  access" removed, "Access now" → "Current access". **The bigger change: the group list now shows
  ONLY groups that actually hold a real grant on the page** (`hasRealGrant()` against the ACL data
  already fetched — a binding beyond Limited Access/Web-Only Limited Access), replacing the old
  policy-name-matching machinery that listed every `_APPROVER`-suffixed group site-wide regardless of
  whether it was actually granted. This changes what the page answers: from "which groups SHOULD have
  access" to "which groups CURRENTLY have access" — the only reading that makes sense once "Should
  have access" is gone. A member popup was added on "N members" — **no reusable expanded-member-list
  component existed anywhere in this codebase** (`accessMemberUi.tsx`'s `MemberSummary` only ever
  renders a collapsed count), so this is a genuinely new small modal, built reusing the pure
  label/sort helpers from `shared/accessMembers.ts` and `GroupManager.tsx`'s existing modal styling.
- **⚠ A REAL LINT DEFECT SURFACED FROM MY OWN EARLIER EDIT, NOT THE AGENTS' WORK.** The Site Access
  agent flagged three `react/no-unescaped-entities` warnings on lines it didn't touch — from the
  `missingUnknown` banner text I'd added moments earlier in the performance fix below, which I had
  verified with a `grep` too narrow to catch a warning type I wasn't expecting. Fixed
  (`&rsquo;`/`&ldquo;`/`&rdquo;`). **The lesson: after any edit, check the FULL lint output at least
  once, not a filtered grep for the warning types you expect** — a narrow filter can hide a real new
  warning as cleanly as no filter at all would have shown it.

## THE ROLES REFERENCE AND THE CREATE-FORM'S INTRO BANNER ARE BOTH GONE — CLIENT'S CONFIRMED CALL (2026-09-02)
Minutes after asking for `RolesReference` to be added to the standalone Group Management page, the
client's live screenshot showed it sitting directly above a SECOND green banner ("Creating a group
here grants nothing… Choosing a persona below…") and asked to remove both: *"Also client feedback
remove this two in Group Management, check the screenshot also ensure that it is remove in the
Folder Management step."* Confirmed explicitly via a clarifying question — **both** the banner and
the roles box, from **both** mount points (standalone page and the guided flow).
- **⚠ THE ORPHANED BANNER WAS THE REAL TELL, and it long predates today.** `GroupManager.tsx`'s own
  intro ("Choosing a persona below also writes its folder mappings…") referred to the create-one-group
  form's persona picker "below" it — and since `hideCreateForm` is now `true` unconditionally at
  BOTH mount points, that picker never renders anywhere any more. The banner had quietly gone from a
  real instruction into a reference to nothing, sitting outside the `{!hideCreateForm && …}` block so
  it kept showing regardless. Removed outright, not reworded — the client's ask was to remove it.
- **`RolesReference` removed from BOTH `GroupManagementPage.tsx` (added minutes earlier) and
  `FolderAdmin.tsx`'s guided-flow group step (present since 2026-08-30).** ⚠ **Flagged to the client
  before removing, not silently complied with**: it is the only place documenting how to create the
  three custom permission levels (CRS Upload/Approve/Delete) — without it, a newly provisioned unit
  can silently end up with no uploader grant at all, and reconciliation only reports the gap after
  the fact ("no CRS Approve role definition on site"). Removed anyway on the client's explicit,
  confirmed instruction. **The component itself is untouched** — re-adding it anywhere is one JSX
  line, not a rebuild, per both files' own comments recording exactly what to paste back.
- `tsc --noEmit` clean, full suite **1597/1597**, 31 lint warnings (baseline, both hits pre-existing —
  `GroupManager.tsx` max-lines and `FolderAdmin.tsx`'s `mapRows`, just shifted line numbers).

## ⚠⚠ THE PEOPLE COLUMN WAS NEVER SLOW — IT WAS FETCHING THE WRONG GROUPS ENTIRELY (2026-09-02)
Two performance fixes in a row (single-batch, then bounded-concurrency) made zero difference, and
the client kept reporting the same symptom. Live Network-tab diagnosis (three rounds, two of them
wasted on filter typos — `sitegroups(id)/users` typed literally, then a filter that matched nothing)
eventually showed **zero requests to `sitegroups(N)/users` ever fired at all**, on any of the three
implementations. That ruled out "slow" and pointed at "never asked."
- **THE ACTUAL BUG, in `StagingAccess.tsx`: `useGroupMembers` was fed `allowedGroupIds`
  (`rows.map(r => Number(r.groupId))`) — the Group Map's LIBRARY-scope entries — while the TABLE
  renders `candidates`, a COMPLETELY DIFFERENT set derived from `groups` (every site group) filtered
  by NAME PATTERN (`roleFromGroupName`).** The two sets never had to overlap, and evidently did not.
  So the hook dutifully fetched members for groups nobody was looking at, while every group the
  table actually displayed sat on "loading…" forever — not because any read was slow, but because
  it was never asked about the right groups in the first place. Both the Group Map items read (200)
  and the site-groups read (200) were succeeding the whole time; the mismatch was purely in-memory.
- **FIXED: `candidates`' computation moved earlier in the component, and `useGroupMembers` now reads
  `candidates.map(g => g.id)` instead of the unrelated `rows`-derived list.** One `groups.filter(...)`
  call now serves both the fetch and the display, so they cannot diverge again.
- **CHECKED THE SIBLINGS BEFORE CALLING THIS DONE** — the exact same hook is used in `PageAccess.tsx`
  and `SiteAccess.tsx`'s new Members column, and neither has this defect: `PageAccess.tsx` feeds the
  hook `granted` (a superset of what it filters for display), and `SiteAccess.tsx` feeds it the full
  `allGroups` (a superset of `shown`). Only `StagingAccess.tsx` had two INDEPENDENT computations
  standing in for what should have been one.
- **THE DIAGNOSTIC PATH IS WORTH KEEPING**: a mistyped filter (`sitegroups(id)/users` as a literal
  string) hid the real answer on the first attempt; a second filter attempt (`user-access-web-parts_`)
  hid it again. Only once the filter was DROPPED ENTIRELY and the raw request list read by eye did the
  actual picture emerge — followed by one precisely targeted filter (`GroupId`, a fragment unique to
  one specific request) to confirm the Group Map read itself was fine. **When a filter keeps producing
  suspiciously empty results, the filter is a suspect before the code is.**
- `tsc --noEmit` clean, full suite **1597/1597**, 31 lint warnings (baseline, unchanged). **NOT yet
  re-tested live** — this is the fix that should actually make the People column resolve.

## SITE ACCESS TRIMMED DOWN AGAIN — GROUPS ONLY, NO PERMISSION COLUMN, NO PEOPLE LIST (2026-09-02)
Client, on the same page shortly after the performance/removal pass: *"Client wants this 'People who
can open the site' to be removed"* and *"dont show individual person under the name only show the
group, also remove the permission, tehy dont want that."*
- **"People who can open the site" (the site-entry group's own member table) is gone from display.**
  The underlying `getGroupMembers` read for the entry group is UNTOUCHED — it still feeds the
  "N user(s) have folder permissions but cannot open the site" gap check above it, a different
  feature the client did not ask to remove. Only the table went.
  - **`members` state itself went with it**, not just its table — once nothing read the state value
    (the gap check already used the read's own local variable, never the state), keeping it would
    have been dead state invisible until lint caught it. Removed alongside its now-unused
    `SpGroupMember` import.
- **"Everything with permission on the site itself" now lists GROUPS ONLY.** Individual people (the
  seven named accounts in the earlier screenshot) are filtered out entirely, not merely hidden —
  `isGroup(pid)` checks membership against `allGroups` (the same site-group list the Members popup
  column already used), **failing OPEN (shown, not hidden) if `allGroups` has not loaded yet**, since
  a row missing because a read is mid-flight is worse than one extra row for a moment.
- **The "Permission" column is gone entirely, and `simplifyLevels` — the three-bucket permission
  display built earlier the SAME DAY — is deleted, not parked.** Unlike the create-one-group form
  elsewhere in this project, there is no stated ongoing need for it; the client's own words were "they
  dont want that," not "hide it for now." No other file ever called it, confirmed via Grep before
  deleting.
- `tsc --noEmit` clean, full suite **1597/1597**, 31 lint warnings — matches the established
  baseline exactly. **NOT yet site-tested.**

## THE GROUP-MEMBERS PERFORMANCE FIX NEEDED A SECOND PASS — BATCHED WAS STILL SLOW (2026-09-02)
Client, after deploying the `fetchAllGroupMembers` fix and confirming the version (1.0.355.0, "a few
seconds ago" in the App Catalog — not a stale-bundle issue): *"Already updated but still took 1
minutes or 2 minutes and still loading."*
- **THE REAL LESSON: the earlier fix traded a request-COUNT problem for a request-SIZE problem, and
  both were slow for the SAME underlying reason — SharePoint, not the network.** ~700 sequential
  `getGroupMembers` calls was slow because there were 700 round trips. ONE `fetchAllGroupMembers`
  request (`sitegroups?$expand=Users`) across all 684 groups was ALSO slow — 1-2 minutes, confirmed
  live — because SharePoint itself is slow to COMPUTE that one enormous response, not merely slow to
  send many small ones. Neither extreme (1 giant request, 700 tiny ones) was right.
- **FIXED with `fetchGroupMembersBounded`** (new export, `shared/spGroups.ts`): several groups'
  members read in parallel per chunk (default **6** — deliberately under the "ten simultaneous calls
  is how a tenant starts returning 429s" threshold this project's own earlier comment on this exact
  hook already names), chunk by chunk rather than all 684 at once or one at a time. Rows fill in
  progressively per chunk, which the single-giant-request version had also lost (the ORIGINAL
  sequential design's stated benefit — "rows filling in visibly also tells the admin it is still
  working" — regained).
- **`useGroupMembers` (`accessMemberUi.tsx`, shared by all three access pages) now uses this
  bounded-chunk approach inline** rather than `fetchAllGroupMembers`.
- **`SiteAccess.tsx`'s OWN separate `fetchAllGroupMembers` call** (for the "missing from site entry"
  check, unrelated to the member-count column) **had the identical latent problem and was fixed the
  same way**, scoped to only the mapped groups actually needed rather than the whole site.
- **⚠ `fetchGroupMembersBounded` RETURNS `undefined` PER GROUP ON FAILURE, NEVER `[]`** — empty ≠
  unknown, the same rule as everywhere else in this codebase, deliberately NOT the shortcut of
  silently treating a failed group as empty. `SiteAccess.tsx`'s "missing" check sets `missingUnknown`
  true if ANY mapped group's read failed, so a partial failure reports "could not fully check" rather
  than a false "no gaps" all-clear.
- `tsc --noEmit` clean, full suite **1597/1597**, 31 lint warnings (baseline, unchanged). **NOT yet
  re-tested live** — next check is whether Approval Library Access and Page Access now load in
  seconds rather than minutes.

## "CREATE ONE GROUP" IS GONE FROM GROUP MANAGEMENT — KEPT, NOT DELETED — BOTH MOUNT POINTS (2026-09-02)
Client: *"remove the Create one group, since we are automatically creating the group for them."*
First pass only touched `GroupManagementPage.tsx` (the standalone page); the client caught the
guided flow's own copy the same day: *"you forgot to remove it from here. Also add [the roles
reference] for the standalone Group Management as well."* **Both mount points now match.**
- **`FolderAdmin.tsx`'s "Add a new segment" flow group step had its OWN, separate copy of this
  exact mode switch** (`groupMode`/`s.modeBar`/`s.modeOn`/`s.modeOff`) — the standalone page and the
  flow step each mount `GroupManager`/`BulkGroupProvisioner` independently, so removing the toggle
  from one page never touched the other. Removed the same way: `groupMode` state deleted,
  `BulkGroupProvisioner` always renders, `hideCreateForm={true}` unconditionally.
- **⚠ `runBusy` STAYS, UNLIKE THE STANDALONE PAGE'S NOW-DELETED COPY.** On the standalone page
  `runBusy` existed only to disable/style the removed button and nothing else read it, so it went
  with the button. In `FolderAdmin.tsx`, `runBusy` is the flow's OWN rail-lock — it holds Back,
  Next, Finish and the way out of the entire flow for as long as a bulk run is in flight (2026-08-17,
  the same mechanism that prevents a step change from silently killing a run mid-way). Deleting it
  here would have re-opened that exact defect. Kept, still wired to `BulkGroupProvisioner`'s
  `onBusyChange`, untouched.
- **`RolesReference` — the "How roles and permission levels work" box — added to the standalone
  page too**, client's own follow-up request. Same shared component `FolderAdmin.tsx` and
  `GroupMapBuilder.tsx` already mount (`shared/rolesReference.tsx`), now a third mount point rather
  than a copy. **`defaultOpen={false}` here, unlike the flow step's `true`** — the flow step opens
  expanded because provisioning-a-new-segment is exactly when an admin needs it; the standalone page
  is visited at least as often just to check on existing groups, where a wall of permission-level
  setup instructions on every visit is noise.
- **`GroupManager`'s create-one-group FORM is NOT deleted, only permanently hidden** —
  `hideCreateForm={true}` unconditionally, the same "kept, not deleted" pattern this project already
  uses for `GroupMapBuilder`'s form (`show="form"`, unreachable since 2026-08-23) and the retired
  folder-tree code. The form's own header comment already recorded WHY it still matters — the
  site-entry group and genuine one-offs bulk provisioning cannot express — so re-enabling it later is
  passing `false` here again, not rebuilding anything.
- **`runBusy` state and `onBusyChange` went with the button** — both existed only to disable/style
  the now-removed "Create one group" tab while a bulk run was in flight; nothing else in this file
  ever read them, and `BulkGroupProvisioner`'s `onBusyChange` prop is optional.
- **Scoped to this ONE standalone page.** The guided flow's own group step (inside `FolderAdmin`'s
  "Add a new segment" flow) mounts `GroupManager`/`BulkGroupProvisioner` separately and was not
  touched — the client's screenshot and complaint were specifically about the standalone Group
  Management page, not the flow.
- `tsc --noEmit` clean, full suite **1597/1597**, 31 lint warnings — matches the established
  baseline, zero new. **NOT yet site-tested.**

## THE MEMBER POPUP EXTENDED TO ALL THREE ACCESS PAGES, AND THE SAME PERFORMANCE BUG FOUND TWICE MORE (2026-09-02)
Client, after the usability pass above: *"Can you apply the same popup to show the people from Page
Access into Approval Library Access and Site Access? also I notice the Approval Library access is
loading too slow for people same goes for Page access, I think if possible show 150 list of groups
first then add a button below to show more."*
- **PAGINATION WAS THE WRONG FIX — the slowness was the exact same defect already found and fixed
  once, in two more places.** `useGroupMembers` (shared by `StagingAccess.tsx` and `PageAccess.tsx`,
  in `accessMemberUi.tsx`) called `getGroupMembers` once PER GROUP, sequentially — identical shape to
  the Site Access defect fixed minutes earlier. Fixed the same way: `fetchAllGroupMembers` reads
  every group's members in one paged request; the hook's external signature (`{ members, refresh }`)
  is unchanged, so neither consumer needed to change for the performance half of this.
- **`MemberCountWithPopup` — the click-a-count-see-who pattern, extracted into `accessMemberUi.tsx`
  as ONE shared component**, rather than three copies of a modal. It wraps the existing
  `MemberSummary` (expandable mode) with its own open/closed state and the modal rendering, so a
  caller needs nothing beyond the group and its `members` map.
  - **`PageAccess.tsx` — the ONLY place this popup existed — was refactored to use the new shared
    version**, its local `expandedGroup` state and modal JSX deleted. First instance, now the
    reference implementation lives in one place instead of the one screen that happened to build it.
  - **`StagingAccess.tsx`** — swapped its `MemberSummary expandable={false}` (plain count) for
    `MemberCountWithPopup`.
  - **`SiteAccess.tsx`** — did NOT previously use `accessMemberUi.tsx` at all (it manages the
    site-entry group's own members directly). Added a NEW **Members** column to "Everything with
    permission on the site itself", shown only for rows that ARE a SharePoint group (matched against
    `fetchAllSiteGroups`' own result, kept in a new `allGroups` state) — an individual person has
    nothing to pop up. Costs one additional `fetchAllGroupMembers` call (on top of the one `reload()`
    already makes for the "missing from site entry" check) — two batched requests instead of one is
    still vastly cheaper than the ~700 sequential reads this replaced.
- **The `position: fixed` modal is unaffected by any of the three pages' `60vh` scroll containers**,
  regardless of DOM nesting — confirmed and documented at each scroll container's own comment, since
  this project has hit the OPPOSITE case (an absolutely-positioned popover clipped by a scroll cap)
  three times already and the distinction is easy to get backwards.
- `tsc --noEmit` clean, full suite **1597/1597**, **31 lint warnings total — matches the established
  pre-existing baseline exactly, zero new ones** across all four touched files. **NOT yet
  site-tested.**

## ⚠ SITE ACCESS TOOK 8 MINUTES TO LOAD — SEVENTH INSTANCE OF THE SAME PER-GROUP SEQUENTIAL READ (2026-09-02)
Client, live-testing the read-only access pages: *"It takes a really long time to load... took about
8 minutes to load."* Spec: `docs/superpowers/specs/2026-09-02-access-pages-usability-pass-design.md`
(step 1, investigated before any UI change to this page).
- **THE CAUSE: `SiteAccess.tsx`'s `reload()` called `getGroupMembers` ONCE PER MAPPED GROUP,
  SEQUENTIALLY**, to compute "who is missing from site entry." `mappedIds` is every unique `GroupId`
  in the whole Group Map — on this site, ~700 — so the page made ~700 sequential round trips just for
  that one check.
- **THE SAME DEFECT, ALREADY FOUND AND FIXED ONCE IN THIS PROJECT.** Group Management's member export
  hit this exact shape on 2026-08-30 ("not an export, it is an outage") and was fixed with
  `fetchAllGroupMembers` — one paged `sitegroups?$expand=Users` request instead of N. `SiteAccess.tsx`
  never got the same fix.
- **FIXED: replaced the loop's per-group `getGroupMembers` calls with one `fetchAllGroupMembers`
  call**, looked up per group id from the returned map instead of awaited per group.
- **`fetchAllGroupMembers` returning `undefined` (a failed/incomplete read) is NOT read as "no
  gaps."** A new `missingUnknown` flag distinguishes it, with its own banner — silently proceeding
  with an empty gap list on a failed read would report a genuinely missing person as found, on the
  one banner whose entire purpose is catching exactly that.
- `tsc --noEmit` clean, full suite **1597/0**, no new lint warnings. **NOT yet re-tested live** — the
  test is simply: does Site Access now load in roughly the time of one paged read, not ~700 of them.

## ✅ A FOLDER-PERMISSION PROBE WITH NO CACHE-BUSTING HEADERS COULD SELF-APPROVE AN UPLOADER'S OWN REPLACEMENT — FIXED AND VERIFIED (2026-09-02)
Client, on the very next test after the above: *"the approve I see happens the moment I click on
Send for a file replacement, and I go check on staging as clarencechojinheng and I dont see the file
at all but when I go check My Submission for chocheetuck it shows Approved."* `chocheetuck4` holds
**only** `UPLOADER_HIGHLY_CONFIDENTIAL` — no approve role whatsoever — ruling out
`autoApproveOwnUpload` firing *correctly*.
- **THE CAUSE: `probeFolderApproveAccess` (`dmsFolderMap.ts`) — the live ACL read that gates
  self-approve — had NO cache-busting headers**, same gap already fixed twice today in
  `Requests.tsx`/`MySubmissions.tsx`. Its URL
  (`GetFolderById(guid'<folder>')/ListItemAllFields/EffectiveBasePermissions`) is keyed ONLY on the
  folder's UniqueId, never on who is asking.
- **`clarencechojinheng` genuinely holds `ApproveItems` on that GCBC folder**, so an earlier probe
  against that exact URL (their own self-approve attempt, or an upload-readiness check) could cache a
  `"granted"` response for it. `chocheetuck4`, tested in the **same browser profile** shortly after,
  hit the identical URL and — per this theory — was served that stale `"granted"` answer instead of a
  fresh, correctly-`"denied"` one, firing self-approve for someone holding no approve role at all.
  Explains every symptom at once: approval happening instantly (self-approve runs inline, right after
  upload/tagging), the item vanishing from staging (Auto-route routes an Approved item away on its
  next poll), and My Submissions genuinely reading Approved — not a display bug, a real self-approve.
- **`probeFolderUploadAccess` and `probeFolderUploadAccessByPath`** — the two ordinary
  upload-readiness probes sharing the identical URL shape — had the same gap and were fixed alongside
  it, since a stale answer there can equally misjudge an ordinary uploader's access either way.
- **Fixed**: `Cache-Control: no-cache` + `Pragma: no-cache` added to all three probes' `GET`.
- **Also fixed while in the area**: `spSubmissionRecords.ts`'s `GET_HEADERS` (behind
  `readSubmissionRecords`, the Cancelled/Deleted overlay) had the identical gap — unrelated to this
  specific bug (that read only decides Cancelled/Deleted, never Approved) but a confirmed real
  instance of the same class of defect, so fixed at the same time.
- **✅ `autoApproveOwnUpload` CONFIRMED `"yes"` in `CRS Config`** — the theory held together, not
  invalidated by the config being off.
- **✅ VERIFIED LIVE 2026-09-02, on `1.0.351.0`, same reproduction shape as the original defect**
  (`chocheetuck4` uploading a replacement in the same browser profile just used as
  `clarencechojinheng`): "Send for approval as a replacement" now lands the file genuinely
  **Pending** — `SUB-20260902-Y4YG` reads "1 pending" on My Submissions, and the item is visibly
  sitting unapproved in `Approval for Document/GHO/GCA/GCBC/2024/Tax Return`. No self-approve fired.
  `tsc --noEmit` clean, full suite **1597/0**, no new lint warnings.

## A DELETED DOCUMENT LEAVES THE SHARED-FILES TAB (2026-08-30, 1.0.321.0)
Client, on six rows all reading `could not be checked`: *"if they don't have access anymore just
remove it.."*
- **⚠ THE DIAGNOSIS CAME FIRST, AND IT SAVED FIXING THE WRONG THING.** Six of six unknown looks
  exactly like a broken query — and `readAcl`'s `$select` had been changed days earlier to add
  `RoleDefinitionBindings`, so the obvious suspect was mine. **The Network tab settled it in one
  glance: `404 System.IO.FileNotFoundException` on every request.** The query was fine; the documents
  were deleted test files. **Ask for the status code before rewriting the request.**
- **`fileGone` IS SET ON 404 AND ON NOTHING ELSE.** A throttle, a 403 or a malformed response all stay
  `undefined` and still render as *"could not be checked"* — because `fileGone` REMOVES the row, and
  **silently shortening the list of who can reach documents on the strength of a read that merely
  failed is the one error this tab must not make.** Same three-state discipline as `holdsRealAccess`.
- **⚠ A 404 MEANS TWO THINGS AND THE WORDING COVERS BOTH.** `GetFileById` answers 404 for a deleted
  file AND for one outside the viewer's reach — SharePoint security-trims to 404, never 403. That is
  the same trap `probeFolderUploadAccess` records. The message says *"no longer in the library"*
  rather than asserting deletion.
- **Dropped BEFORE `mergeShareAcl` ever sees them**, so a gone file cannot reach `rollUp` and colour
  the tab's state. It also removes a **Revoke button that could never work**.
- **⚠⚠ THE OMISSION WAS ANNOUNCED IN A NOTE, AND THE CLIENT HAD IT REMOVED THE SAME NIGHT
  (1.0.322.0).** It read *"6 shared documents are no longer in the library, so they are not listed
  here — a deleted document cannot be reached by anyone. The requests stay on the Share tab as the
  record."* Client: *"This might make client confuse and wonder."*
  - **THE RULE IT WAS BUILT ON IS THIS CODEBASE'S OWN AND IT IS STILL RIGHT — it was applied to the
    wrong kind of omission.** *A tab about access that quietly gets shorter is a tab people stop
    trusting* holds where the omission is a LIMIT OF THE TAB, which is why the sentence about
    natively-shared files stays: that one names something the page genuinely cannot see, and an
    administrator who does not know it would draw a wrong conclusion. A deleted document is not a
    blind spot — it is a document that no longer exists, and the tab's answer about it (*nobody can
    reach it*) is complete.
  - **A NOTE THAT RAISES A QUESTION THE PAGE CANNOT ANSWER IS WORSE THAN SILENCE.** *"6 documents"*
    invites *which six, deleted by whom, and when* — none of which the Requests page knows or should.
    The audit log answers it; this tab does not.
  - **Nothing is lost by removing it**: the request rows still stand on the **Share** tab, which is
    where the history was always kept.
  - `goneFiles` in `shared/requests.ts` is retained (unused, with its tests) so the count is one line
    away if it is ever wanted. `withoutGoneFiles` — the part that actually matters — is unchanged.
  - **The all-empty case now falls through to *"every share here has been revoked or has ended"***,
    which is true of a deleted document too. Its own sentence would have re-introduced the wondering.
- **This accumulates on a live site**: every deleted document that was ever shared would otherwise sit
  here for ever, reading like a permissions fault.
- **Verified**: `tsc --noEmit` clean, `requests.test.ts` **131 passed** (5 new, pinning that
  `undefined` and a clean read both KEEP the row). Packaged as `1.0.322.0` (the note removed).
  **NOT yet site-tested.**

## THE REVOKE AUDIT ROW LANDED, AND THE TEST THAT PRODUCED IT PROVED HALF OF WHAT IT LOOKED LIKE (2026-08-30, 1.0.323.0)
First live `ShareRevoked` row on ClarenceDMSTesting: right event, right file
(`HCDocuments/GHO/GCA/EG/2024/Tax Return/sd - asd - dasd - 30-08-26.pdf`), right recipients,
`Flow:RequestActivity`. So the `Revoked` branch added to `CRS — Audit request activity`'s `EventKind`
is working.
- **✅ `ActorEmail` DEMONSTRABLY READS `RevokedBy`, NOT `DecidedBy` — proven from the row's own data.**
  On that request `DecidedBy` is `clarencechojinheng@gmail.com` (the Head of Unit who approved) and
  `RevokedBy` is `clarence@trinergydigital.com` (who took it back). The audit row names the SECOND.
  That is the exact case `RevokedBy` was built for, and it was already sitting in the data.
  - **⚠ THIS WAS FIRST RECORDED AS INCONCLUSIVE, on the assumption that one account did both jobs —
    and the next step proposed was to FAKE a second actor by hand-editing `DecidedBy`.** Opening the
    list view settled it instead. **The audit log's actor column alone cannot answer this; the two
    candidate fields have to be READ.** Before staging a test to force two values apart, check
    whether the live data already holds them apart.
- **⚠ THE NOTE'S DATE WAS A DAY BEHIND, AND THE AUDIT ROW BESIDE IT DISAGREED.** Row stamped
  `30/Aug/2026 02:43`; the appended `DecisionNote` read *"revoked by … on 2026-08-29"*.
  `revocationNote` was fed `new Date().toISOString()` — **UTC** — so at UTC+8 anything done before
  08:00 local recorded the previous day. Fixed with `localDateStamp` (local components, 3 tests).
  - **The two dates in one record serve different readers and stay different on purpose:**
    `EventTime` is ISO/UTC because `$filter` needs it; this sentence is a person saying when they did
    something, so it takes their own local day. **Do not "make them consistent" by putting UTC back.**
  - Same family as Auto-route's `Created` stamp landing ~8 hours early.
- **A RECIPIENT NOBODY RECORDED WAS REVOKED TOO, and that is the mechanism working.** The request named
  only the gmail address; the note names `clarence@trinergydigital.com` as well. **Revoke all** resets
  inheritance, so it drops everyone the FILE's ACL holds — including anyone granted outside CRS. Proof
  the tab reads the live ACL rather than the request row.
- **⚠ THE DOCUMENT WAS IN `HCDocuments` AND WENT TO AN OUTSIDE GMAIL ADDRESS.** An approved share hands
  a named person one Highly Confidential document whatever their clearance — the HC library split
  governs GROUP access and per-file shares sit outside it by design. **Tell the client rather than
  letting them find it.**
- **Verified**: `tsc --noEmit` clean, `requests.test.ts` **134 passed** (3 new), full suite 1561/0.
  Packaged as `1.0.323.0`.

## ⚠ `allowExternalSharing` GATES THE REQUEST, NOT THE SHARE — AND A GUEST WALKS PAST IT (2026-08-30)
Verified live: on ClarenceDMSTesting **neither** `allowExternalSharing` nor `tenantDomains` exists in
`CRS Config`, so external sharing is OFF by the fail-closed rule — and a share request to
`clarencemindpalace@gmail.com` was raised, approved and executed anyway.
- **THE BYPASS IS THE DOCUMENTED CAVEAT, NOW CONFIRMED RATHER THAN SUSPECTED.** `MySubmissions.tsx`
  always pushes the SIGNED-IN USER'S OWN DOMAIN into `tenantDomains`, so `chocheetuck4@gmail.com`
  raising a share to a gmail address had `gmail.com` counted as the organisation. `isExternal`
  answered false and the gate never engaged.
- **⚠ AND NOTHING DOWNSTREAM RE-CHECKS IT.** On the approver's side `isExternal` is used for DISPLAY
  ONLY — the row chip, the recipient chip, the dialog warning. There is no `allowExternal` test in
  `decide`/`performShare`. What actually permits the share is **SharePoint's own external sharing
  setting**, which `ShareObject` honours. So the config row reads like a policy switch and is closer
  to a hint: **it stops an honest internal user asking, and stops nobody else.**
- **THE CHIP IS JUDGED BY WHOEVER IS LOOKING.** On one screen `clarencemindpalace@gmail.com` was
  UNFLAGGED while `crystal@trinergydigital.com` was flagged *outside the organisation* — both are
  outside this tenant. A guest viewer sees their own guest domain as internal, so the flag an approver
  reads before deciding can be wrong in **both directions at once**.
- **ON SDG THIS MOSTLY SELF-CORRECTS, WHICH IS WHY IT IS NOT URGENT** (client, 2026-08-30: *"for sdg
  they wont be sharing files with outside organization so its fine I suppose"*). Every staff address is
  `@sdguthrie.com`, so the viewer's own domain IS the tenant domain, the chip is right, and the absent
  row refuses external shares at the point of raising — the wanted behaviour, with no config.
  - **⚠ THE GAP THAT REMAINS: the agency accounts are GUESTS there.** A share raised from
    `@trinergydigital.com` counts that domain as internal and can reach other `@trinergydigital.com`
    addresses past a gate that is supposedly off. **Set `tenantDomains = sdguthrie.com`** — one list
    row — and decide `allowExternalSharing` deliberately rather than leaving it blank.

## ⚠ A REVOKED SHARE TOLD THE UPLOADER THEIR REQUEST HAD *FAILED* (2026-08-30, 1.0.324.0)
Client, reading their own My Submissions after the first live revoke: *"Your share request was
attempted and failed by clarencechojinheng@gmail.com … This message should not be showing, as per UX
design its weird to keep this label stuck here."* The request had been approved, used, and then
deliberately revoked — nothing failed, and the person named had approved it rather than broken it.
- **THE CAUSE IS A TWO-BRANCH CHAIN WITH A CATCH-ALL:** `Approved` → "approved", `Rejected` →
  "rejected", **everything else** → "attempted and failed". `Revoked` was added to `RequestStatus`
  for the Requests page redesign and **this screen was never revisited**, so the widest-reaching
  reader of that union silently reclassified a normal outcome as an error.
  - **`Cancelled` HAD THE SAME DEFECT AND FOR FAR LONGER** — withdrawing your own request came back
    as *"attempted and failed by …"*, i.e. a failure somebody else caused. Nobody reported it, which
    is how a catch-all branch hides: it is only wrong for the statuses nobody tests.
  - **Sixth instance of `feedback-audit-consumers-when-narrowing-data`, in its other direction:
    ADDING a value to a shared union makes every reader a change site, exactly as narrowing one
    does.** `RequestStatus` now has six members and a fallback branch can only ever be right for the
    ones its author was thinking about. **Name every status; leave no `else`.**
- **THE ACTOR IS WITHHELD ON A REVOKE, and that is not cosmetic.** `decidedBy` holds whoever
  APPROVED it — a different person from the revoker by design (that asymmetry is why `RevokedBy`
  exists) — so printing it beside "taken back" credits the wrong act to the wrong person. `Cancelled`
  likewise: nobody decided it but the requester.
- **⚠ AND THE DECISION NOTE IS WITHHELD ON A REVOKE.** After a revoke that note carries the appended
  line naming **every principal dropped from the file**, including anyone granted OUTSIDE CRS whom
  this uploader never named. The live example showed the uploader `clarence@trinergydigital.com`, a
  recipient they had nothing to do with. Kept for `Approved`/`Rejected`/`Failed`, where it is the
  reason they came for.
- A revoke now ends *"You can ask again if it is still needed"*, matching the existing line on a
  rejection — the banner is informational, not a dead end.
- **Verified**: `tsc --noEmit` clean, full suite 1561/0. Packaged as `1.0.324.0`. **NOT yet
  site-tested.**

## ⚠ "APPROVE DOES NOTHING" WAS A STALE TAB — AND THE HUNT FOUND A REAL DEFECT ANYWAY (2026-08-30, 1.0.325.0)
Client: *"I click Approve and the list is still there, I click reject and the list is still there"*,
then, minutes later: *"I didn't deploy anything I have to refresh and clear cache which then works."*
**A hard refresh fixed it. Nothing was wrong with the code.**
- **FOURTH TIME A STALE TAB HAS PRESENTED AS A CODE FAULT IN THIS PROJECT**, and the first three are
  already written up (the dead web part with `ERROR: [object Object]`, the reconciliation run
  asserting an old page policy, and the Bulk Upload clash dialog showing the pre-`reserved` scheme).
  The mechanism is the CONTENT HASH in the bundle filename plus SharePoint's SOFT NAVIGATION, which
  swaps modern pages without a full reload — so one tab opened before a deploy hands out stale
  references all session and the component never instantiates.
  - **⚠ THE DIAGNOSIS GIVEN WAS WRONG, and it was plausible enough to have been believed:** the
    outcome banner really does sit at the top of the page in green, so "the click did nothing" had a
    tidy explanation fitting every symptom. **A theory that explains the symptom is not evidence for
    the cause.** The cheap check comes first: **hard refresh before diagnosing dead controls**,
    exactly as memory `feedback-scale-process-to-change-size` says.
- **THE DEFECT IT DID TURN UP IS REAL AND IS FIXED:** `Requests.tsx` rendered **every** notice in the
  GREEN `s.ok` style — `Recorded as failed — …`, `Could not confirm this request is still open
  (HTTP …)`, `Could not finish: …` and `Nothing was done — this request is no longer open` included.
  **A refusal dressed as a success**, on the screen where the action cannot be undone. `noticeBad`
  now selects `s.warn`, and every failure path sets it.
- **AND THE BANNER SCROLLS ITSELF INTO VIEW.** It sits above the tabs while the row being decided can
  be far below, so a real failure was invisible from where the click happened. `scrollIntoView` on
  the notice, wrapped in try/catch.
  - **⚠ THE EFFECT IS DECLARED WITH THE OTHER HOOKS, ABOVE THE `rows.state` EARLY RETURNS.** Below
    them it would run a different number of times on the render after loading finishes — *"Rendered
    more hooks than during the previous render"* — blanking the whole web part. Same trap as the
    approval panel's probe (1.0.243.0), which cost an hour.
- **Verified**: `tsc --noEmit` clean, full suite 1561/0. Packaged as `1.0.325.0`. **NOT site-tested,
  and it fixes NOTHING the client reported** — it makes the next genuine failure legible.

## ⚠⚠ A REVOKED SHARE CAME BACK TO THE APPROVER'S QUEUE AS `Pending`, FOR EVER (2026-08-30, 1.0.326.0)
Client: *"I click Approve and the list is still there"* — on SHARE requests only; deletions worked.
The banner (scrolled off-screen, see 1.0.325.0) had been saying it all along: **"Nothing was done —
this request is no longer open (revoked). Someone else decided it first."**
- **THE CAUSE: `fromListItem` PARSED `Status` AGAINST A HAND-WRITTEN LIST OF FOUR** —
  `["Pending", "Approved", "Rejected", "Failed"]` — with everything else falling back to **Pending**.
  `Revoked` and `Cancelled` were added to `RequestStatus` for the redesign and never added there.
- **THE FALLBACK'S OWN COMMENT DEFENDED IT, AND WAS EXACTLY BACKWARDS:** *"anything unrecognised
  reads as Pending, so a row written by a newer version is still decidable rather than invisible."*
  Both statuses that reached it are **TERMINAL**, and presenting a terminal row as decidable is what
  produced a queue entry **nobody could ever clear** — Approve and Reject were both refused by the
  pre-write re-read guard, correctly, on every click.
- **⚠ THE GUARD IS THE ONLY REASON THIS WAS COSMETIC.** Without the 1.0.220.0 re-read, Approve would
  have re-shared a document whose access had deliberately been taken back, and Reject would have
  overwritten a revocation. **A fail-closed check twenty lines away turned a data-modelling bug into
  a display one.**
- **`Cancelled` HAD THE SAME DEFECT AND NOBODY HAD NOTICED** — a request the uploader withdrew sat in
  the approver's queue looking like work. That is the worse of the two: an approver could act on
  something the requester had pulled.
- **FIXED WITH A COMPILER-ENFORCED SET, NOT A LONGER LITERAL.** `STATUS_SET` in `shared/requests.ts`
  is a `Record<RequestStatus, true>`, so **adding a member to the union without adding it there is a
  compile error**; `REQUEST_STATUSES` and `parseRequestStatus` derive from it. A `RequestStatus[]`
  literal would have been the same bug with two more entries.
  - `STATUS_ORDER` on the filter is **derived** from it too — the preferred list decides ORDER only,
    membership comes from `REQUEST_STATUSES`, so a status added later lands at the end of the
    dropdown rather than nowhere.
  - **`MySubmissions` was casting the raw text (`as RequestStatus`) and is now parsed as well.** It
    happened to behave correctly, which is exactly why it would have been skipped: a blind cast lets
    a garbage value into the union and onto the screen.
- **SEVENTH INSTANCE OF `feedback-audit-consumers-when-narrowing-data`, AND THE SECOND TONIGHT** —
  the My Submissions banner (1.0.324.0) was the same union, the same omission, a different reader.
  **When a value is added to a shared union, `grep` for every literal that enumerates it.** A type
  cannot find them; only a search can.
- **Verified**: `tsc --noEmit` clean, `requests.test.ts` **141 passed** (7 new, including a
  round-trip over every member and a length assertion), full suite 1568/0. Packaged as `1.0.326.0`.
- **⚠ RECOVERY ON A SITE ALREADY SHOWING PHANTOMS: none needed.** The rows were always correct in
  the list; only the reading was wrong. Deploying moves them straight into **Decided**.

## ⚠ THE SHARE EXPIRY DATE IS DISPLAYED AND NEVER ENFORCED — FLOW DESIGNED, NOT BUILT (2026-08-30)
Runbook: **`docs/superpowers/specs/2026-08-30-share-expiry-enforcement-flow-runbook.md`**. Client, on
being shown the gap: *"2. Build the enforcement, lets build the flow."* **Nothing is built yet.**
- **`performShare` SENDS NO EXPIRATION OF ANY KIND.** Its `SP.Web.ShareObject` body carries `url`,
  `peoplePickerInput`, `roleValue`, `groupId`, `propagateAcl`, `sendEmail`,
  `includeAnonymousLinkInEmail`, `emailSubject`, `emailBody`, `useSimplifiedRoles`. `expiresAt`
  reaches the LIST COLUMN and three display strings — the row line, the Shared-files line and
  `decisionSummary` — and stops. **Access granted through an approved share is permanent.**
- **⚠ AND IT CANNOT BE FIXED IN THE SHARE CALL.** SharePoint has no per-grant expiry for a DIRECT
  USER grant: expiry is a property of an anonymous LINK (which this is not — these are `role:` grants
  to named people) or a tenant-level Entra guest-lifecycle setting. So the sentence in the approval
  dialog — *"until … or 2026-08-30 passes"* — can only be made true by something that comes back
  later and removes the grant.
- **THE TWO HONEST OPTIONS WERE PUT TO THE CLIENT AND THEY CHOSE ENFORCEMENT.** The other was to stop
  promising it: relabel the field *"review by"*. Worth remembering if the flow is ever deferred —
  leaving the label as it is, unenforced, is the one choice that is actively misleading.
- **⚠ THE TWO TRAPS THAT WOULD MAKE THIS FLOW DESTRUCTIVE, both in the runbook:**
  - **`ExpiresAt ne null` in the filter.** Blank means *no expiry*, and an OData comparison against a
    null column is not reliably false — omit the clause and the first run revokes **every share that
    has no expiry date at all**, silently.
  - **Compare against `startOfDay(utcNow())`, not `utcNow()`.** `ExpiresAt` is stored as midnight, so
    a plain `utcNow()` comparison expires a share **at 00:01 on the day the user chose**.
  - And **`removeroleassignment` per recipient, never `resetroleinheritance`** — one file can carry
    two share rows (observed live 2026-08-30), and resetting kills the ones that have not expired.
- Needs no code change and no schema change; it reads and writes columns that already exist, and the
  existing `CRS — Audit request activity` flow logs the result because `Status` changes.

## ⚠⚠ THE BROWSER CACHE ANSWERED "WHO CAN REACH THIS DOCUMENT" — AND SAID NOBODY (2026-08-30, 1.0.327.0)
A share was approved, the recipient was emailed, and they opened the file. **Shared files still read
`(0)` — *"Nobody currently has access through a CRS share"*.** Pressing **Re-check access** changed
nothing. The Network panel had the answer in one column: four ACL reads, three `404` and one
**`304 Not Modified`**.
- **THE 304 IS THE WHOLE BUG.** The request went out, the browser revalidated, and the **cached body
  from before the share existed** was served back. The recipient was absent from those stale
  principals, so `mergeShareAcl` marked them `revoked`, `currentlyShared` dropped the file, and the
  tab reported nobody — about a document somebody had just been given and could demonstrably open.
- **⚠ THE FAILURE IS SILENT AND POINTS THE WRONG WAY.** It does not read as "stale"; it reads as
  *"every share here has been revoked or has ended"*, which is a positive claim about security. On
  the one screen whose entire job is answering who can reach a document, **a cached answer is worse
  than no answer.**
- **⚠ AND IT DEFEATED THE BUTTON BUILT TO ESCAPE IT.** *Re-check access* clears `aclsRead` and
  re-issues every probe — correct, and useless, because the re-issued request was answered from
  cache. **A refresh control is only as good as the request underneath it**; test it against a value
  that has actually changed, not merely that it re-fires.
- **FIXED WITH TWO INDEPENDENT GUARDS, because either alone can be defeated** by a proxy or a service
  worker: `GET_FRESH` sends `Cache-Control: no-cache` and `If-None-Match: *`, and `bust()` appends a
  unique parameter so there is nothing left to match on. **Applied to the ACL read ONLY** — the list
  reads are cheap to be a second stale and this one decides a security answer.
- **A `304` REACHING THE PARSER NOW RETURNS `undefined`, NEVER AN EMPTY PRINCIPAL LIST.** With the
  guards in place it should be unreachable; if one is ever stripped in transit, *"not modified"*
  means we were handed nothing, and nothing is **unknown**. An empty list would read as "everyone has
  been revoked" — the exact wrong answer this defect produced.
- **⚠ THE DIAGNOSTIC ORDER THAT WORKED, THIRD TIME IN TWO DAYS: read the Network panel's STATUS
  column before theorising.** The three preceding hypotheses — a sticky probe, `holdsRealAccess`
  discarding the recipient, and `fileGone` over-removing — were all plausible, all cheap to argue
  for, and all wrong. `404, 404, 404, 304` settled it instantly. Same lesson as the deleted-document
  diagnosis hours earlier and the `Copy file` inputs before that.
- **⚠⚠ THE FIRST FIX (1.0.327.0) SENT `If-None-Match: "*"` AND THEREBY *GUARANTEED* THE 304 —
  corrected in 1.0.328.0.** On a GET that header means *"if any version of this resource exists,
  answer 304"*. It is the read-side inversion of the `IF-MATCH: "*"` idiom used on the MERGEs a few
  hundred lines below in this same file, and it was added while reasoning about caching rather than
  about conditional requests. **Cache-busting a GET is `Cache-Control` + `Pragma` + a unique URL;
  never a conditional header.**
  - **The site test caught it in one screenshot** — the file appeared under Shared files, the
    recipient was listed, and both read `could not be checked`, with a `304` still in the Network
    panel. That is the *new* 304 branch doing its job: **unknown, not "revoked"**. So the guard added
    in the same change is what kept a self-inflicted 304 from once again reporting a live share as
    ended. Fail-safe behaviour is worth most when the bug is your own.
  - Cost, while it stood: no Revoke button, because a recipient whose state is `unknown` is not
    offered one — correct, and the client's report was *"I cannot revoke access??"*.
- **Verified**: `tsc --noEmit` clean, full suite 1568/0. Packaged as `1.0.328.0`. **NOT yet
  site-tested** — deploy, hard refresh, open Shared files, and the recipient should read
  `has access` with Revoke offered.

## ⚠⚠ THE BROWSER CACHED THE REQUEST LIST, AND FOUR DIAGNOSES CHASED THE SYMPTOM (2026-08-30, 1.0.330.0)
Client, after an evening of it: *"I refresh and it shows the request and I click approve and or reject
it doesnt work but when I clear cache and refresh it dissappears? Can we ensure the share system is
working properly? this is so messy."* They were right to be exasperated — **the code was correct
throughout and the browser was replaying an old response.**
- **THE PROOF, IN TWO SCREENSHOTS OF ONE PAGE.** Plain refresh: *Waiting for you (2)*, two `Pending`
  rows with live Approve/Reject. Clear cache, refresh: *Waiting for you (**0**)*, the same two rows in
  **Decided** reading `Revoked`, with both revokers named. A direct REST read confirmed the list had
  held them as `Revoked` the whole time.
- **⚠ A STALE RESPONSE AND A STALE BUNDLE PRESENT ALMOST IDENTICALLY, AND THAT IS WHY THIS TOOK SO
  LONG.** Both look like *"the page is showing something that is not true and the buttons do
  nothing"*. **The discriminator is cheap and should be the FIRST question: does CLEARING THE CACHE
  fix it (stale response) or only a HARD RELOAD (stale bundle)?**
- **FOUR WRONG DIAGNOSES CAME OFF THIS ONE CAUSE**, each plausible, each costing a build:
  1. a stale BUNDLE — true earlier in the evening for a different symptom, which is what made it
     convincing here;
  2. `parseRequestStatus` — a REAL bug (1.0.326.0) fixed on the way past, but not this one;
  3. the counts line miscounting — it was reporting the cached rows faithfully;
  4. `fileGone` over-removing, then the ACL probe being sticky — the same cache, one endpoint over.
- **THE PRE-WRITE RE-READ GUARD IS WHY NONE OF IT WAS DANGEROUS.** Every Approve on a phantom row was
  refused, because that check asks the SERVER rather than the rendered state. **A guard that
  re-reads is worth more than any amount of correct rendering** — it held while four different
  things about the page were wrong. It is now cache-busted itself: a cached answer THERE would
  defeat it silently and look exactly like a pass.
- **FIXED: every read on both screens sends `Cache-Control: no-cache` + `Pragma: no-cache`, and the
  request-list read and the decision guard also carry `bust()`** — a unique URL, the only part a
  proxy or service worker cannot ignore. `GET_FRESH` is now an alias of `GET`, because two header
  objects that must stay identical are how one of them quietly stops being.
- **A PERMISSIONS-AND-APPROVALS SCREEN MUST NEVER RENDER A CACHED ANSWER.** Everything on it is a
  decision somebody is about to act on. The cost is a handful of uncached requests per visit, which
  is not a trade worth thinking about twice.
- **My Submissions took the same treatment** (9 reads): it decides whether to show *"you already
  asked"* and whether to offer Cancel, so a stale answer invites a duplicate request for a decision
  already made, or a Cancel on something already carried out.
- **Verified**: `tsc --noEmit` clean, full suite 1568/0. Packaged as `1.0.330.0`. **Deploy, then test
  with an ORDINARY refresh** — the whole point is that clearing the cache should no longer be
  necessary, so a test that begins by clearing it proves nothing.

## ⚠⚠⚠ `heft package-solution` DOES NOT BUILD — NINE PACKAGES SHIPPED STALE JAVASCRIPT (2026-08-30)
**The single most expensive mistake of the project so far, and it cost an entire evening of wrong
diagnoses.** Versions `1.0.323.0` through `1.0.330.0` were built, tested, packaged, deployed and
installed — and every one of them contained the **1.0.322.0** JavaScript.
- **THE CAUSE: `heft package-solution` ONLY ZIPS `release/assets`. It does not run webpack.** The
  loop used all evening was `npx heft test` then `npx heft package-solution --production`. `heft
  test` compiles TypeScript to `lib-commonjs` and runs jest — which is why `tsc --noEmit` was clean
  and 1568 tests passed every time — but it **never regenerates the hashed browser bundles**.
- **⚠ THE TELL, AND IT IS UNMISSABLE ONCE YOU LOOK:** in `release/assets`, the `*_<hash>.js` files
  are what SharePoint serves, and their timestamps were **hours older** than the un-hashed `*.js`
  beside them. `ls -la release/assets/` shows it in one line.
  ```
  requests-web-part.js                             16:53   <- compiled, NOT shipped
  requests-web-part_6d724879775462323ced.js        02:26   <- SHIPPED, and 8 versions stale
  ```
- **⚠ THE SECOND TELL: PACKAGE SIZE.** A real production package here is **~456 KB**. The
  package-solution-only runs produced **~2.1 MB** — they were re-zipping a debug-shaped asset folder.
  A sudden 4x size change between builds means the build shape changed, not the code.
- **ALWAYS `npm run build`** (`heft test --clean --production && heft package-solution --production`).
  Never `package-solution` alone. **`--clean` is what forces the hashed bundles to be regenerated.**
- **VERIFY THE ARTEFACT, NOT THE BUILD LOG.** Function names are minified away, so grep for a STRING
  LITERAL the change introduced — `grep -c "Revoked:!0"` (a `Record` key survives minification),
  `grep -c "no-cache"`, a new user-facing sentence. **A green build and a passing suite say nothing
  about what is inside the .sppkg.**
- **⚠ WHAT IT COST, and every one of these was written up as a real cause at the time:** a stale
  BUNDLE (right mechanism, wrong reason — the bundle was stale in the PACKAGE, not the browser), the
  browser HTTP cache, a sticky ACL probe, `fileGone` over-removing, a miscounting summary line, and
  a self-inflicted `If-None-Match` 304. **Only three of the nine builds fixed anything the client had
  reported.**
- **⚠ THE CLIENT WAS RIGHT EVERY TIME AND WAS TALKED OUT OF IT.** *"I refresh and it shows the
  request … when I clear cache and refresh it dissappears"*, *"Same issue"*, *"this is so messy"*.
  Each report was answered with another theory and another package containing the same code.
  **When someone says a fix did not take effect, verify the ARTEFACT before proposing a mechanism.**
- **THE DIAGNOSTIC THAT FINALLY WORKED, in order:** the request URL in the Network panel lacked the
  `&_=` parameter added two builds earlier → the running code is older than that build → the app
  catalog says `1.0.330.0` and Site Contents says `1.0.330.0` → therefore the PACKAGE is wrong, not
  the deployment. **A version number is a label somebody typed; the bundle is the fact.**

## THE CLIENT QA PASS — 35 ITEMS, THREE STANDING RULES REVERSED (2026-08-30, 1.0.331.0)
A full QA deck plus seven numbered points. **34 of 35 are built**; the last is a Power Automate
change. Everything below that reads like an omission was a DECISION — read the reasons before
"fixing" any of it back.

- **⚠ MUCH OF THE DECK WAS SHOT AGAINST AN OLDER BUILD, and two items evaporated on inspection.**
  The Requests screenshots predate the tabbed redesign (1.0.318–321, never deployed), so the orange
  `Deletion` chip they asked to move **no longer exists** — the Deletion/Share tabs replaced it, and
  a chip inside a tab would repeat the tab it sits under. And *"bulk upload leaves no record on My
  Submissions"* was disproved by one query: `Source eq 'BulkUpload'` returns rows. **Check which
  build the client is on before building the fix.**

### The three reversals — each deliberate, each traded
1. **THE GUIDED-FLOW RAIL NO LONGER NAVIGATES FORWARD.** Until today the rule was *every step is
   reachable; the rail says what is outstanding and padlocks almost nothing*. The client asked for
   the opposite outright (*"ensure that side panel doesn't allow user to navigate"*) and confirmed it
   knowing the cost: *"just do as the client's requested, if any issue for the flow we will fix that
   flow."* Clarence, on the trade: *"its ok if they can't jump, as long as its not confusing for
   them."*
   - **BACKWARD IS STILL ALLOWED.** Re-reading a step you have passed changes nothing and is how
     somebody checks what they typed; blocking it turns a wizard into a one-way corridor.
   - **The cost, now the client's to carry:** an admin who did step 3 outside the tool last week must
     click through 1 and 2 to reach it. That was the original argument for a clickable rail.
   - **⚠ `stepState` and `remainingCount` ARE UNTOUCHED in `folderFlows.ts` and still tested** — only
     `FolderAdmin` stopped rendering them. `isLocked` and `blocksNext` read the same `FlowFacts`, so
     the four real locks still fire and still explain themselves beside Next. **What went is the
     GRADING of every step, not the checking of the ones that can be checked.**
2. **THE AUDIT LOG NO LONGER SHOWS `Source` UNDER THE ACTOR.** That field is what tells a row a
   PERSON wrote (`CrsConfiguration`) from one a FLOW wrote (`Flow:DocumentsDeletions`) — i.e.
   *"somebody did this"* from *"the system did this"*. Still stored, still in the CSV, so the
   question stays answerable; just not at a glance.
3. **THE UPLOAD FORM NO LONGER SAYS "nothing has been uploaded yet", AND BULK UPLOAD NO LONGER STATES
   THE 50-FILE CAP.** Both facts are unchanged and still enforced — `MAX_FILES` refuses the 51st file,
   and staged batches still die with the tab. Only the sentences are gone, so the `beforeunload`
   guard and the Cancel confirmation are now the ONLY things warning about lost staged work.

### Exclusions that look like oversights and are not
- **`&` — plain and FULLWIDTH `＆` — is NOT blocked by `inputSanitize.ts`.** GHO's own department is
  *Group Legal, Risk ＆ Compliance* and the term store requires the fullwidth form: a blanket reading
  of *"no special characters"* would refuse the client's own data. Pinned by test, as are `'` and `-`.
- **Free-text REASONS and DECISION NOTES are not guarded.** A rejection reading *"Why is this dated
  2025?"* needs `?`; *"see item #4"* needs `#`. Those are prose one person writes for another, never
  a file name. **Raise it with the client rather than widening it quietly.**
- **The `Delete File` TAB still carries the value `Deletion`.** It is compared against `RequestType`
  everywhere and every stored row says `Deletion`; renaming the value would be a data change wearing
  a copy change's clothes. `TAB_LABEL` is the only place the word is displayed.
- **`Approve`/`Reject` on the approval radios are LABELS; the `Decision` values are still
  `Approved`/`Rejected`** — that is what is written to the item and read back everywhere else.
- **`previewTarget.fileUrl` is unchanged and still the DOWNLOAD route.** The new `openUrl` is what
  "Open in a new tab" uses; keeping both means a future Download control has something correct.
- **93 days, not the 90 in the client's mock**, in both delete dialogs. Client: *"stay 93, they might
  not know that is why they say 90."*

### The two real defects fixed, and why they were invisible
- **"Open in a new tab" DOWNLOADED the document.** It pointed at `preview.fileUrl` — the file's own
  URL — and SharePoint SERVES that, which for an Office file means a download. `openUrl` sends Office
  through `WopiFrame…&action=view` (not `embedview`, which is right in an iframe and wrong in a tab)
  and everything else through `?web=1`. **Applied to PDFs too**, though only a `.docx` was reported:
  a PDF downloads on some tenants for the same reason.
- **`fetchAllGroupMembers` FAILED SILENTLY**, which is why the CSV read `not known` in every People
  and Members cell while the same screen listed the members. It did ONE request with `$top=5000` and
  `$expand=Users` and returned `undefined` on any failure with no log. It **pages** now (`$top=500`,
  following `odata.nextLink`, capped at 20) and **logs the status** — 403, 429 and 500 need opposite
  fixes. **Hitting the cap returns `undefined`, never a short answer:** a spreadsheet saying a unit's
  approver group is empty gets acted on. ⚠ Paging is the likeliest cause and is NOT proven — the read
  failed silently, so there is no evidence either way. The log will say next time.

### `#3` — no request where the viewer can already act
Client: *"it doesn't make sense for HOU or HOD to approve their own request when they can delete or
share, only PIC will need to raise a request."* `canActDirectly` in `shared/mySubmissions.ts`.
- **⚠ MATCHED AGAINST THE DOCUMENT'S WHOLE TIER CHAIN, never its unit alone.** A Head of Department's
  `DEL`/`SHARE` sits on the DEPARTMENT term and fans down; that term is already one of the document's
  own tiers, so a chain comparison expresses the fan-out **with no term-store expansion and no extra
  request**. A unit-only compare would have been wrong for every HoD.
- **Split by stage, because the rights are:** an APPROVED document needs `DEL`, one awaiting approval
  needs `DELS` — different roles, different libraries. HoU holds both, HoD only the first.
- **Scoped to the viewer's OWN groups.** Reading every Group Map row would take a PIC's buttons away
  because somebody ELSE holds delete on their unit — precisely the person they should be asking.
- **FAILS OPEN.** Unread chain, empty role list, failed read ⇒ buttons offered. Understating rights
  costs one redundant request; overstating them leaves a PIC with no route to act at all.
- The absence is EXPLAINED (*"You can delete or share this document yourself, in the library"*): a row
  with no controls and no sentence reads as a page that failed to load.
- **"Your requests" is gone from the Requests page** for the same reason — a PIC cannot reach that
  page at all, and an approver's own requests belong on My Submissions.

### Still open
- **QA#5: an approver is emailed that their OWN upload was approved.** Power Automate, not code.
  Runbook: **`docs/superpowers/specs/2026-08-30-suppress-self-approval-email-runbook.md`**. One extra
  row on the Condition that already suppresses bulk imports, comparing `Author.Claims` to
  `Editor.Claims` from `Get item` — SharePoint records no "approved by" field, and `Editor` is the
  best signal there is. **Read it BEFORE the stamp**, which rewrites `Editor` to the uploader on the
  copy; read it after and every uploader stops being notified. Compare CLAIMS, never Email (a guest
  can have none, and two blanks compare equal). ⚠ **It reverses the 2026-08-24 decision that
  self-approve should still email** — confirm that is wanted. `HC Auto Route` needs the same row.
  **Not built.**
- **Date formats now DIFFER between the app and the library views.** `formatSubmittedOn` gives
  `27 Aug 2026`; SharePoint column formatting still gives `27/Aug/2026` — that is JSON on the column,
  not code. Making them match is a separate job.
- **`ApprovalDocument.tsx` still shows the raw `FieldValuesAsText` Document Date** (`8/17/2026 12:00
  AM`). Only My Submissions was corrected, by re-reading the raw `Edm.DateTime` — **never by parsing
  the formatted string back**, which is gotcha #1's trap (`8/9/2026` is ambiguous).
- **Verified**: `tsc --noEmit` clean, full suite **1590/0**, and the shipped bundles grepped for a
  string from this batch before packaging — `npm run build`, never `package-solution` alone.
- **⚠ TEST PLAN: `docs/2026-08-30-qa-batch-test-plan.md`.** — **none of the UI changes have run on a
  site.** The unit tests cover the pure rules (`canActDirectly`, the character list, the preview
  URLs, the date format) and **none of the JSX**, because this project has no UI tests. §4.8/4.9 and
  §7.6–7.8 are the pairs that prove a gate works rather than merely blocks everything.
- **⚠ THE RAIL LOCK NEEDED A SECOND FIX BEFORE IT WAS SAFE (1.0.332.0), found by Clarence asking
  *"We are going to ensure that it wont break right?"*.** Gating on the CURRENT step meant stepping
  BACK threw away the way forward — an admin on step 4 who glanced at step 1 would have had to press
  Next three times through finished screens. `maxIdx` tracks the furthest step REACHED, so going back
  and returning is free and only genuinely-new steps need Next. **That was the "confusing" the client
  was worried about, not the locking.**
  - And the blocked-Next message still ended *"— or jump straight on from the list of steps"*, which
    the lock made **false**. It names **Refresh list** now. **A gate pointing at a control that no
    longer works is worse than a gate with no advice**: they try it, nothing happens, and the page
    reads as broken. Pinned by a test asserting the old sentence is absent.

## THE EMAIL'S `[Approve]` / `[Reject]` LINKS PRE-TICK THE RADIO (2026-08-31, 1.0.335.0)
Client, on the reminder email: *"best follow exactly as crystal gave, is that possible?"* — and for
templates 2 and 3 the only clause that was not buildable was the pair of action links. It is now,
without decoding anything from an inbox. Rule in `decisionFromLink` (`shared/approvalQueue.ts`, pure,
4 tests); applied in `ApprovalDocument.tsx`.
```
.../ApprovalDocument.aspx?itemId=N&decision=approve
.../ApprovalDocument.aspx?itemId=N&decision=reject     (+ &lib=hc for HC)
```
- **⚠ PRE-SELECTING IS NOT DECIDING, and nothing here may ever become an auto-submit.** Everything
  the page does at the moment of submitting — the destination-folder guard, `documentsFileClash`, the
  `ApproveItems` probe — runs on that button press. A link that decided from the email would skip all
  three, and the native Approve/Reject command already proves how expensive that class of bypass is.
- **THE BEARER TRIGGER URL STAYS REJECTED, and is worse here than in the HC design that first refused
  it.** The reminder goes to up to ten approvers, so ten copies of a token that approves a document
  would be sitting in ten mailboxes.
- **⚠ IT APPLIES ONLY WHILE THE DOCUMENT IS STILL PENDING.** An already-decided one shows what
  actually happened to it — pre-ticking `Reject` on something approved an hour ago states a falsehood
  on the one screen that IS the record of the decision. The button is disabled for a decided document
  anyway, so this is about what the approver is TOLD, not what they can do.
- **⚠ READ ONCE INTO A `useRef`, THEN STRIPPED FROM THE ADDRESS BAR.** The queue rewrites `?itemId=`
  with `replaceState` as the approver walks it, so a surviving parameter would pre-tick a decision
  from an email about a **DIFFERENT DOCUMENT** — the only failure here that could genuinely mislead
  somebody. Stripping also stops a refresh re-applying it after they deliberately changed it.
- **Anything unrecognised is `undefined`, never a guess** — a mangled parameter leaves the page
  exactly as it behaves with no parameter at all: nothing ticked, button disabled until they choose.
  `pending` is deliberately unreachable, since it is a system state and would render a dead page.
- **`decisionFromLink` takes `string | undefined`, not `string | null`** (the house lint rule), so the
  `URLSearchParams.get` null is converted at the call site. The `?? ""` inside still absorbs one.
- **Verified**: `tsc --noEmit` clean, full suite **1594/0**, 31 warnings (one FEWER than before — no
  new ones), and the shipped hashed bundle grepped for the minified rule before packaging.
  **NOT yet site-tested.**

## `CRS — Approval reminder` — THE 3-DAY REMINDER FLOW (2026-08-31)
Runbook: the client's template 3, verbatim, in
`docs/superpowers/specs/2026-08-28-email-bundling-and-templates-design.md` §5. **BUILT AND VERIFIED
END TO END on ClarenceDMSTesting 2026-08-31** — one pending document, two approvers, two emails, all
six template fields populated and both action links live. Shape:
`Recurrence → GetStale → Apply to each → GetApproverGroup → HasApprover → GetMembers →
MembersWithEmail → PickEmails → Recipients → HasRecipients → Apply to each 1 → Send an email (V2)`.
- **⚠ FOUR FAULTS CAME OUT OF ONE PASTE, and every one of them was invisible in the designer.**
  Building this cost an evening, almost all of it on paste artefacts rather than logic:
  1. **The `Apply to each` had NO input.** An unbound loop exposes no item, so `items('Apply_to_each')`
     inside it reported *"does not exist in the current workflow"* — which reads as a misspelled action
     name and is not. **Check the loop's own Parameters before doubting an `items()` reference.**
  2. **`body('GetDoc')` — an action from `NotifyApprovers`, which this email body was copied from.**
     Five references to an action this flow does not have. The designer reported it on the SIBLING
     cards, so the card showing the error was not the card containing it.
  3. **Trailing `\r\n` inside TWO condition expressions** (`HasApprover`, `HasRecipients`) and the
     subject line. Same trap as the `ItemUniqueId` newline of 2026-08-23, from the same cause: pasting
     a multi-line expression into a Power Automate field. **Never press Enter in an fx box.**
  4. `Recipients` left over and unused once the send moved inside the per-approver loop.
- **ONE EMAIL PER APPROVER, not one joined To line** — that is what makes `Dear [Approver Name]`
  possible, which the template asks for. Inner `Apply to each 1` over `take(body('MembersWithEmail'),
  10)`; `item()` is the APPROVER and `items('Apply_to_each')` is the DOCUMENT. ⚠ Swap them anywhere
  and the email names the wrong person or the wrong file, silently.
- **⚠ THE `take(…, 10)` CAP IS NOW SILENT** — the "only the first 10 were emailed" note was dropped as
  not being in Crystal's template. Acceptable only because there is one Head of Unit per unit, so ten
  is far beyond any real approver group. It reminds DAILY, so an uncapped 40-member group would send
  40 emails per document per day.
- **`Submission Date` goes through `convertFromUtc(…, 'Singapore Standard Time')`.** `Created` is UTC,
  so at UTC+8 a document uploaded before 08:00 would show the PREVIOUS day — the same 8-hour fault
  already recorded against Auto-route's `Created` stamp.
- **MANAGED METADATA COMES BACK AS `Label|GUID` FROM THE CONNECTOR — gotcha #5, in a new place.**
  `?['Year']?['Value']` reads `2024|f365ff14-4b34-4c1e-a8ac-30acda07c068`, and Document Type the same.
  Both now go through `first(split(coalesce(...?['Value'], ''), '|'))`.
  - **Chosen over hunting for a `Label` property**, because the shape of `Value` is PROVEN by a live
    run while a guess at a second property comes back **blank, not an error** — an empty line reads as
    a formatting slip rather than a wrong path.
  - **IT SHIPPED WRONG AND THE RUN WAS GREEN.** The email sent, the flow reported success, and the
    GUID simply sat in the body. **A green `Send an email` says the send worked, never that the
    content is right** — the only check is opening the message.
- **THE `Send an email (V2)` BODY IS A RICH-TEXT EDITOR, so pasted HTML is ESCAPED, not rendered.**
  The first correct-looking build emailed literal `<p>` and `<b>` tags to two approvers. **Click the
  `</>` button in the body toolbar and paste the HTML in code view.** Applies to any flow here that
  sends formatted mail.
- **`&` IN A URL MUST BE `&amp;` once the body is real HTML** — `...?itemId=1771&amp;decision=approve`.
  Harmless while the tags were being escaped, which is the kind of fault that only appears after an
  unrelated fix.
- **A CORRECT ZERO AND A BROKEN FILTER LOOK IDENTICAL, so the date clause was proven by moving it.**
  With `-3` the only pending document (2 days old) is legitimately excluded, so `GetStale` returns `[]`
  and every run is green and empty — indistinguishable from a filter that matches nothing for ever, and
  **nobody would notice for weeks**. Settled by temporarily setting `addDays(utcNow(), -1)`, watching one
  item come back and an email arrive, then restoring `-3`. **Do the same on the HC clone**; a malformed
  filter 400s, but one that parses and never matches does not.
- **GMAIL THREADS THE REMINDERS.** The subject is identical per document, so a re-test lands inside the
  earlier conversation and reads as *"no email arrived"*. Check the thread before diagnosing the flow.
- **⚠ Also needs `FSObjType eq 0` in the `GetStale` filter** — folders arrive Pending too, so without
  it approvers are reminded to approve folders.
- **STILL TO DO:** clone as `CRS — HC approval reminder` (trigger → the HC approval library, Group Map
  filter → `Role eq 'APRHC'`, `&lib=hc` on both links). ⚠ **`APR` is held by the plain `hou` persona**,
  so an HC clone filtered on `APR` emails every ordinary Head of Unit about HC documents.
- On SDG, **sign in as the SERVICE ACCOUNT before the first action** — the connection is baked in for
  life.

## ✅ REPLACE NOW VERSIONS INSTEAD OF DELETE-AND-RECREATE — BUILT AND VERIFIED ON BOTH FLOWS (2026-09-02)
Follows the runbook `docs/superpowers/specs/2026-08-31-versioned-replace-runbook.md` (§0 explains why
this was needed — read that first). **BOTH `Auto-route` and `HC Auto Route` are done and verified live**,
each with two clean end-to-end tests: an ordinary approval (no clash, unaffected) and a genuine
replacement (Version History showing `2.0` current / `1.0` with the original content, empty recycle
bin, metadata intact, and a new `Replaced` audit row naming the **uploader who sent it as a
replacement** — see the correction below; it first shipped naming the approver instead).
- **THE SHAPE MATCHES THE RUNBOOK EXACTLY**: `GetDestMatch` (`Get files (properties only)`, never `Get
  file metadata using path` — the runbook's own warning about the 404-on-absent-file failure mode) →
  `DestExist` Condition → True: `GetSourceContent` → `UpdateFile` (writes a real version) → False:
  `Copy_file` (unchanged) → `DestItemId` (Compose, normalises which item id downstream code should
  use) → the rest of the flow unchanged, but repointed.
- **A NEW `EventType: "Replaced"` AUDIT ROW TRACKS WHO REPLACED A FILE** — the client's actual ask.
  Fires from a **separate `WasReplaced` Condition placed AFTER `Get_ApprovedBy_HTTP` and
  `GetSourceUniqueId`**, re-testing the same `GetDestMatch`-based boolean `DestExist` already checked
  — **not** inside `DestExist`'s own True branch, because `Get_ApprovedBy_HTTP` (needed for the actor)
  runs downstream of the whole `DestExist` block closing. Putting the audit-row `Create item` inside
  `DestExist`'s branch throws `InvalidTemplate: … cannot reference action 'Get_ApprovedBy_HTTP' …
  must either be in 'runAfter' path or within a scope action on the 'runAfter' path` — hit live, on
  the first attempt, in the normal flow. **Attribution ORIGINALLY reused the exact same
  `ApprovedBy`-falling-back-to-`Editor` pattern the existing "Approved" row uses, and that was
  WRONG — corrected the same day, 2026-09-02.** `ItemUniqueId` deliberately reuses `GetSourceUniqueId`'s
  value (the *source* item's UniqueId) rather than looking up the destination's — matching this
  project's established convention that "History of this file" always threads on the source's GUID,
  even though the destination's real UniqueId differs on an ordinary Copy.
  - **⚠ "WHO REPLACED IT" IS NOT "WHO APPROVED IT", AND THE `ApprovedBy` PATTERN ANSWERS THE WRONG
    QUESTION HERE.** Found live by the client on the first cross-person test: a document uploaded by
    `chocheetuck4` as a deliberate replacement, approved by `clarencechojinheng`, logged the "Replaced"
    row as `clarencechojinheng` — correct for "Approved", backwards for "Replaced". **The replacer is
    whoever chose "Send for approval as a replacement" at upload, i.e. the SOURCE item's own `Author`**
    — the exact same field the "Uploaded" row already reads, and needing none of the Editor-fallback
    complexity `ApprovedBy` needed (an uploaded item always has an Author).
  - **FIXED: `ActorName`/`ActorEmail` on the "Replaced" `Create item` now read
    `triggerOutputs()?['body/Author/DisplayName']` / `.../Author/Email'` directly** — no fallback, no
    reference to `Get_ApprovedBy_HTTP` at all. The action's position in the flow (still after
    `Get_ApprovedBy_HTTP` and `GetSourceUniqueId`, for the `ItemUniqueId` dependency) is unchanged;
    only the two attribution fields moved.
  - **VERIFIED LIVE 2026-09-02, genuine cross-person test** (`chocheetuck4` uploaded and sent as
    replacement, `clarencechojinheng` approved): the "Replaced" row correctly read `chocheetuck4`,
    distinct from "Approved" (`clarencechojinheng`) and "Moved to …" (`clarencechojinheng`, the account
    that ran the copy). **`HC Auto Route`'s identical action still needs the same two-field fix** —
    same pattern, not yet applied there.
- **NO NEW WORK WAS NEEDED FOR "WHO SENT IT AS A REPLACEMENT AT UPLOAD TIME"** — deliberately not
  built, for two reasons. First, it would be pure redundancy: the only way a document can ever reach
  `DestExist`'s True branch is by having been uploaded under an *unrenamed* clashing name, which the
  upload form's clash dialog only permits via "Send for approval as a replacement" — so the `Replaced`
  row **is** the record of that click, confirmed by the actual outcome. Second, it genuinely couldn't
  be built client-side anyway: `CRS Audit Log` write access is Owners/service-account only, and the
  uploader's own browser session is neither — which is exactly why this whole trail runs through flows
  in the first place.
- **✅ A DISPLACED RECORD NOW READS `Cancelled`, NOT `Deleted` — BUILT AND VERIFIED THE SAME DAY, on
  `Auto-route` only (`HC Auto Route` clone still to do).** Runbook:
  `docs/superpowers/specs/2026-09-02-replaced-record-stamp-runbook.md`. The `Replaced` audit row
  above always logged correctly, but nothing told the displaced file's `CRS Submissions` record it was
  superseded — `markRecordReplaced` is only ever called client-side, from the upload form's overwrite
  and the approval page's *confirmed* `documentsFileClash` warning, neither of which fires for a
  replace this flow catches on its own (the native Approve/Reject command, the Bulk Approve panel, or
  a collision appearing after the approval click). Fixed by adding a `GetDestSubmissionFileId` read
  inside `DestExist`'s True branch (captures the OLD stamp before anything overwrites it) and a MERGE
  onto the matching `CRS Submissions` row, alongside the `Replaced` audit-row write in `WasReplaced`'s
  True branch.
  - **⚠ THAT ALONE WASN'T ENOUGH, AND THE REASON IS A SECOND, LARGER GAP: `Update file` never
    transfers ANY list column, not just `SubmissionFileId`.** `Copy file` automatically carries every
    column from source to destination for free; `Update file` only overwrites bytes, and this flow's
    only explicit stamp is Author/Editor/Created. So after a replace, the destination item kept the
    FIRST upload's `SubmissionId`/`BatchId`/`SubmissionFileId` forever — meaning the REPLACER's own
    record could never resolve (permanently `Deleted`, the false message this feature exists to
    prevent, now on the wrong person) and the `Cancelled` stamp on the OLD record was invisible,
    since `mergeRecords`' precedence rule lets a still-resolving record win as `live` regardless of
    `replacedAt`.
  - **Fixed with two more actions, `GetSourceSubmissionIds` + `StampDestSubmissionIds`**, run
    UNCONDITIONALLY on both branches (harmless no-op on the ordinary `Copy file` path) between the
    Author/Editor/Created stamp and the source-delete step, transferring
    `SubmissionId`/`BatchId`/`SubmissionFileId` from source to `DestItemId`.
  - **⚠ THE WIDER METADATA GAP (Document Type, Year, Confidentiality, tier columns, Remark,
    LegallyPrivileged, Vendor, etc.) IS STILL OPEN, DELIBERATELY.** Every one of those ALSO stays
    stale after a replace — client's own scope decision (2026-09-02) was My Submissions only, for now.
  - **VERIFIED LIVE**: a clean replace test showed the displaced record reading **Cancelled** —
    *"Replaced by a newer upload from clarencechojinheng@gmail.com on 2026-09-01. The details below
    are what this submission carried."* — and the replacer's own record resolving live/Approved.
  - **⚠ EVERY NEW ACTION HAD TO BE BUILT VIA THE PARAMETERS TAB, NOT CODE VIEW** — pasting a JSON
    block into Code view repeatedly landed inside the Uri box instead of replacing the action. Build
    field by field (Method, Uri, Headers as rows, Body via "Add new parameter"); use Code view only to
    inspect the result afterward, and watch for trailing newlines picked up from copy-pasted values.
- **⚠⚠ THREE REAL BUGS SURFACED WHILE VERIFYING THE ALREADY-BUILT NORMAL FLOW, NONE OF WHICH WOULD
  HAVE SHOWN UP FROM A GREEN SAVE OR A SUPERFICIAL READ:**
  1. **`DestExist`'s condition carried a trailing `\r\n`** on the `length(...)` expression — the exact
     class of bug this project has hit repeatedly (a `0` typed as text, a pasted expression bringing a
     line break with it). Silently risks the field being read as a string template instead of a pure
     expression, breaking the numeric comparison in an unpredictable direction.
  2. **The Routed audit row's `ItemPath` in `HC Auto Route` read `@body('Copy_file')?['Path']`
     directly** — a genuine divergence from the normal flow, which builds its path independently from
     `Compose_1` + filename. This would have resolved to `null` on every HC replace, corrupting that
     one audit row's path silently. **Never assume the HC clone matches the normal flow just because
     the earlier parts did** — this is the same lesson as every other HC-clone defect in this project,
     just found this time by deliberately checking rather than by a live incident.
  3. **A hand-typed `DestItemId` expression self-referenced its own action** (`outputs('DestItemId')`
     instead of `outputs('Copy_file')?['body/ItemId']`) — caught only because Power Automate's own
     expression editor flagged "This expression has a problem" before save. Worth remembering: the
     function/dynamic-content picker can suggest an action's own name when you're editing that exact
     action, and it's easy to select the wrong one without noticing.
- **`runAfter` VALUES SHOWING `"SUCCEEDED"` (ALL CAPS) TURNED OUT NOT TO BE A REAL PROBLEM**, contrary
  to an initial concern raised mid-build. Confirmed the Settings tab's "Configure run after" checkbox
  was correctly set to "Is successful" only in every case, and — decisively — the normal flow's full
  end-to-end replacement test succeeded completely with this exact casing present in multiple
  `runAfter` clauses. Power Automate's dependency-status matching is evidently case-insensitive here.
  **Do not re-raise this as a concern without new evidence** — it was checked and it does not matter.
- **THE "HC Documents" → "Highly Confidential Document" WORDING FIX (found the same day) is also
  done** — the Routed row's `item/Title` in `HC Auto Route` was hardcoded `"Moved to HC Documents: …"`
  from before the library was last renamed; both HC deletion-audit flows already used the correct
  current name and needed no change.
- **STILL OPEN, not raised with the client yet**: the runbook's §5 `Created By` question (overwrite to
  the replacer, or preserve the original uploader) — client explicitly said *"leave it as it is"* for
  now, meaning the existing stamp behaviour (which already reads `Author` from the *source* item, i.e.
  effectively the replacer) is unchanged and untouched by any of this build.

## ⚠⚠ REPLACE DOES NOT VERSION — `Copy file` DELETES THE FILED DOCUMENT (2026-08-31, 1.0.336.0) — SUPERSEDED ABOVE, kept for the original diagnosis
Proven on site, and it invalidates the premise the client's own Replace decision was made on.
`Restricted & Confidential Document`, one document replaced through the approval page:

| | before | after |
|---|---|---|
| version | `1.0` | **`1.0`** — a NEW one, not `2.0` |
| modified | 8/26 6:32 PM | 8/31 8:03 PM |
| size | 227.3 KB | 117.4 KB |
| uploader | chocheetuck4 | Clarence Cho |
| Submission Id | `SUB-20260826-KJ78` | `SUB-20260831-7X92` |

- **AUTO-ROUTE'S `Copy file` WITH `nameConflictBehavior: 1` DELETES THE DESTINATION ITEM AND CREATES
  A NEW ONE.** Version history is never reached, so the 500-major-version setting on both approved-side
  libraries makes no difference whatsoever. Every claim in this file and in the product that "the
  previous version stays in version history" was **false for the approved side**.
- **THE RECYCLE BIN IS THE ONLY RECOVERY, and it is real:** the deleted item was found at
  `/Shared Documents/GHO/GCA/EG/2024/Term Sheet`, 228.4 KB, **Created By the ORIGINAL uploader**,
  deleted at the moment of routing. 93 days. **But an approver would never think to look there**, and
  nothing tells them to.
- **⚠ THE CLIENT CHOSE REPLACE ON A FALSE PREMISE.** Their words were *"they will be careful"* and
  *"that is on purpose to ensure they update the file"* — decided while our own confirm dialog said the
  old version was kept in version history. **A decision made on a false premise is not a decision;
  put the correction to them before treating Replace as settled.**
- **THREE STRINGS CLAIMED IT AND THEY ARE NOT ALL WRONG THE SAME WAY** — corrected in 1.0.336.0:
  - `ApprovalDocument.tsx` and `BulkUpload.tsx` describe the APPROVED side, replaced by Auto-route.
    **Both were false**; they now say the document moves to the site recycle bin for 93 days.
  - `Form.tsx` describes replacing a PENDING file via `Files/Add(overwrite=true)` — a genuine
    SharePoint overwrite, which **does** version. That half was TRUE. Its dialog also covers the filed
    copy when `where === "both"`, and **that half was false**, so the sentence is now split in two.
- **✅ VERSIONING VERIFIED ON ALL FOUR CRS LIBRARIES (2026-08-31), which is what made the split
  answerable:** `Restricted & Confidential Document`, `Highly Confidential Document`,
  `Approval for Document` and `Approval for Highly Confidential Document` — **major versions, 500 kept,
  no time limit**, every one. The two approval libraries also read content approval **Yes** and Draft
  Item Security **"Only users who can approve items (and the author)"**, which independently confirms
  the 2026-08-22 correction that draft isolation was never actually turned off.
- **⚠ THE REAL FIX IS RUNBOOKED AND NOT BUILT — `docs/superpowers/specs/2026-08-31-versioned-replace-runbook.md`.** Client accepted 2026-08-31 that anyone who can read a document can also open its version history, which is consistent with the model rather than an exception: the unit folder is already the smallest confidentiality boundary, and HC history sits behind the HC ACLs the same way. Branch in Auto-route: `Get file metadata using
  path` on the destination; if it resolves use **`Update file`** (which writes a genuine `2.0`),
  otherwise `Copy file` as today. Three things make it more than an action swap:
  1. **Everything downstream reads `body('Copy_file')`** — `GetSourceUniqueId`, the audit `Create item`,
     the metadata stamp, the source delete. Normalise the destination item id through ONE `Compose`
     both branches feed, or the tail gets duplicated and the rarely-run copy drifts.
  2. **`Created By` becomes a POLICY QUESTION.** On the Update path the destination already holds the
     ORIGINAL uploader in `Author`, and today's stamp would overwrite it with the new one. Version 2 by
     a different person is exactly the case; silently rewriting the original author loses provenance.
     **Client decision, not ours.**
  3. **It must be built TWICE** — Auto-route and HC Auto Route. Every HC clone in this project has
     shipped with a reference nobody swapped.
  4. **⚠ USE `Get files (properties only)`, NOT `Get file metadata using path`.** The latter FAILS
     with a 404 when the file is absent — which is the NORMAL case — so the flow would red-run
     on every ordinary approval. The former answers an empty array.
- **⚠ `2026-08-28-file-replacement-design.md` §7.3 STILL SAYS "version history holds the
  content" AND IS NOW WRONG.** Written before this was tested. Correct it when the runbook
  ships, or it is the next stale line somebody acts on.
- **⚠ THE BULK APPROVE PANEL CANNOT REPLACE AT ALL.** `documentsFileClash` skips a clashing file with
  no override — correct while replacement was forbidden, incomplete now it is a deliberate workflow.
  Approvers will keep falling back to the page for exactly this case until it gains the same choice.

## ⚠⚠ AN HC-CLEARED PIC COULD NOT REACH THE UPLOAD FORM — `UPLHC` WAS MISSING (2026-08-31, 1.0.337.0)
Found live: a guest in `GHO_GCA_GCBC_UPLOADER_HIGHLY_CONFIDENTIAL` got **AccessDenied on
`Upload-Form.aspx`** — `Type=item`, Site Pages item 4. The `/upload/i` rule in `pageAccessPolicy.ts`
read `["UPL", "APR", "APRHC"]`, and **`pic_hc` is `["UPLHC"]` with no literal `UPL`.** So an HC
uploader held no qualifying role: never granted the page, and **REMOVED by the derived page pass if
they ever had been**, since that pass asserts the full ACL at page scope.
- **FOURTH INSTANCE OF ONE PATTERN, and this file already named the first three:** the upload form
  needed `APR` (2026-08-17), Requests needed `DEPTVIEW` (2026-08-21), My Submissions needed `UPLHC`
  (2026-08-21). **A page rule names the role a persona is THOUGHT of as having, while `PERSONAS` gives
  it a superset or a management role instead.** Check the persona's actual `roles` array, never the
  family name.
- **⚠ BOTH SIBLING RULES IN THE SAME FILE ALREADY HAD IT.** `Bulk-Upload.aspx` and
  `My-Submissions.aspx` are `["UPL", "UPLHC", "APR", "APRHC"]`; `UPLHC` went into those two when
  `APRHC` arrived on 2026-08-24 and was missed here. **When a role is added to one page rule, check
  every rule whose audience overlaps** — the three uploader pages are now pinned as ONE set by a test
  asserting their lists are identical, so the next divergence fails in CI rather than in a support call.
- **Widening cannot grant anyone a new place to upload** — the page grant opens the FORM; the folder
  ACL decides what can be filed, and the form probes `AddListItems` and shows the not-ready empty state
  where it fails. Same argument that admitted `APR` in 2026-08-17.
- **MIGRATION = re-run reconciliation.** Page grants are derived and asserted, so the run adds every
  `_UPLOADER_HIGHLY_CONFIDENTIAL` group to the upload form. No row or schema change.
- **VERIFIED LIVE 2026-09-01.** A run on the fixed bundle printed
  `GHO_GCA_GCBC_UPL_HIGHLY_CONFIDENTIAL -> upload-form.aspx -> Read (page, from UPLHC)` for every HC
  uploader group across both segments, zero removals. The guest then opened the form with the cascade
  correctly locked to their single unit. **`(page, from UPLHC)` in the log is the proof the fixed policy
  is the one running** - it cannot be printed by an older bundle.
- **WARN: THE RUN LOG NAMES GROUPS BY A STALE `GroupName`, AND IT MISLEADS DURING EXACTLY THIS KIND OF
  DIAGNOSIS.** The line above says `..._UPL_HIGHLY_CONFIDENTIAL` while the live group is
  `..._UPLOADER_HIGHLY_CONFIDENTIAL` - renamed on 2026-08-27. The grant is CORRECT: it resolves by
  `GroupId`, which a rename never changes. Only the label is old, because it comes from the value
  stored on the mapping row. It cost a wrong turn here, looking for a suffix that no longer exists in
  the log and no longer exists on the group. **Print the live title.** Not fixed.
- **⚠ THE DIAGNOSIS RAN IN THE WRONG ORDER AND COST AN HOUR.** This file's own recorded sequence for
  *"they cannot get in"* is: **is the principal the right one** → is the group granted on the page → is
  the binding `Read` → does their session list the group → sign out and in. Instead it went: is it the
  new package (no — three string changes), is the page a draft (no — published), is the page ACL wrong
  (no — `GHO_GCA_EG_UPLOADER` held a genuine `Read`)... and only then, which group is the person
  actually IN. **The Group Management lookup answers that in one search and would have pointed
  straight at `_UPLOADER_HIGHLY_CONFIDENTIAL`.**
  - **⚠ AND A TRUNCATED ACL DUMP WAS READ AS ABSENCE**, mid-diagnosis: the page's role assignments ran
    7 → 900s and showed no GHO group, so GHO was called missing. GHO's ids are **1215–1220** and the
    paste had simply been cut off. `sp-capped-read-reads-as-absent` in a new guise — **check the id
    range before concluding a principal is not in a list.**

## ⚠⚠ FOLDER MANAGEMENT WENT BLANK ON EVERY FLOW — A HOOK BELOW AN EARLY RETURN (2026-08-31, 1.0.338.0)
The card grid rendered; clicking any card produced a **completely empty page**. Console:
**`Minified React error #310`** — *"Rendered more hooks than during the previous render"* — at
`useEffect` inside `folder-manager-web-part`.
- **`FolderAdmin.tsx` returns early twice** — the picker (`!flow && !allTools`) and All tools — and the
  `maxIdx` effect added with the rail lock sat **below both of them**. So it ran on the flow path and
  not on the picker path, and the render after opening a flow saw more hooks than the one before.
- **⚠ THE COMMENT ON THAT EFFECT ALREADY SAID *"Declared with the other hooks, above every early
  return."*** It was not. **A comment describing the fix is not the fix** — the same failure recorded
  five days earlier against the HC clearance probe, whose comment claimed "counted after the work, not
  before it" for five days while the counter was still incremented before it.
- **THIRD INSTANCE IN THIS PROJECT.** The approval panel's `ApproveItems` probe (1.0.243.0, cost an
  hour) and My Submissions' share-recipient picker were both this. **A `useEffect` below a `return`, in
  a component that returns early, is a blank web part waiting to happen.**
- **⚠ A BLANK SPFx WEB PART HAS NO ERROR UI, so it reads as a DEPLOYMENT problem.** The first
  diagnosis offered here was a stale bundle — plausible, because this project has had exactly that
  (the content-hash MIME refusal) and 17 tabs were open across two deploys. **The console tells them
  apart in one look:** `Refused to execute script … MIME type ('text/html')` is a stale bundle;
  `Minified React error #310` is hooks. Ask for the console before proposing either.
- **The fix moves `idx` up too**, computed defensively (`flow ? flow.steps.length : 0`) because there
  may be no flow yet — `steps.length - 1` on an absent flow is `-1`, and the clamp has to survive that
  without the runner's `active` cast.
- **Introduced by the 2026-08-30 design pass** (card grid, `maxIdx` rail lock, stepper restyle), which
  the QA test plan explicitly flags as **never having run on a site** — 1,597 unit tests cover the pure
  rules and **none of the JSX**. This is what that gap looks like when it lands.
- **Verified**: `tsc --noEmit` clean, full suite 1597/0, no new lint warnings. Packaged as `1.0.338.0`.

## ⚠⚠ `Editor` IS NOT THE APPROVER — QA#5 SUPPRESSES EVERY APPROVAL EMAIL (2026-09-01)
Proven on site from the raw trigger payload of `Audit — approval activity`. `clarencechojinheng`
approved a document uploaded by `chocheetuck4`, and SharePoint reported:
```
Author  -> chocheetuck4
Editor  -> chocheetuck4        <- AFTER the approval
{ModerationStatus} -> Approved
```
- **SHAREPOINT DOES NOT RESTAMP `Editor` ON A MODERATION-STATUS CHANGE.** Not via `File.Approve()`
  (the approval page) and not via a MERGE of `OData__ModerationStatus` (the bulk panel) - both routes
  were tested and both leave `Editor` as whoever last EDITED the item, which is the uploader from the
  upload form's tagging call. Approving is not an edit. Sensible, and fatal to the assumption.
- **THE `2026-08-30-suppress-self-approval-email-runbook.md` SAID THIS AND IT WAS UNDER-WEIGHTED:**
  *"`Editor` MEANS 'LAST PERSON TO TOUCH THIS', NOT 'APPROVER'. They coincide because approving is
  normally the last thing that happens to a pending item."* They never coincide. **A caveat written
  as "this is the best signal available" is not evidence that the signal works.**
- **TWO LIVE CONSEQUENCES, both silent:**
  1. **`Audit — approval activity` names the UPLOADER as the approver** on every approval, because its
     `ActorName` reads `Editor/DisplayName`. Wrong data on the one list that IS the record.
  2. **QA#5 SUPPRESSES 100% OF APPROVAL EMAILS.** `Condition 2`'s second row compares `Author.Claims`
     to `Editor.Claims` to suppress the ONE case of an approver approving their own upload. They are
     always equal, so **no uploader is ever told their document was approved.** A feature meant to trim
     one email switched off the whole notification, and nobody reports a missing email for weeks.
- **INTERIM MITIGATION (30 seconds): delete the second row of `Condition 2`**, keeping the `BulkImport`
  row. Emails resume for everyone; the cost is one redundant email when an approver approves their own
  upload - the pre-QA#5 behaviour, and strictly better than silence.
- **THE REAL FIX IS AN `ApprovedBy` COLUMN, the same shape as `RevokedBy` (1.0.317.0).** SharePoint
  records no actor for this action, so the app must: both approval routes run AS the approver and know
  who it is. Write it on both approval libraries; `Audit — approval activity` reads it for the actor
  and falls back to `Editor` when blank; `Condition 2` compares `Author.Claims` to it.
  - **BLANK MUST MEAN SEND.** The native Approve/Reject command bypasses the app entirely and leaves it
    empty, as does any document approved before the column existed. Degrading to a redundant email is
    always correct; degrading to silence is what this entry is about.
  - **BOTH routes must write it** - `ApprovalDocument.tsx` and `BulkApprovePanel.tsx`. One of them
    missing it is a half-working audit trail, which is worse than a consistently broken one because
    nobody knows which rows to trust.
- **✅ THE CODE HALF IS BUILT (1.0.344.0).** `APPROVED_BY_COLUMN` in `shared/optionalColumns.ts`
  ("ApprovedBy", Text) is asserted by reconciliation on the TWO approval-side libraries only
  (`Staging`/`StagingHC`, skipping the HC one when `hcAvailable()` is false) — never `Documents` or
  the archive, because the question "who approved THIS pending item" is meaningless once the item is
  routed and deleted, and Auto-route reads it from the SOURCE item before the copy.
  - **BOTH routes now write it**, guarded by `libraryHasColumns` so a site where reconciliation has
    not yet run degrades to today's behaviour rather than a failed approval:
    - `ApprovalDocument.tsx` — a SEPARATE MERGE writing only `ApprovedBy`, sent AFTER `approve()` has
      already succeeded, in its own `try/catch` that swallows failure. Same reasoning as the
      `markRecordReplaced` bookkeeping a few lines above it: a failure to stamp must never read back
      as a failed approval.
    - `BulkApprovePanel.tsx` — folded into the EXISTING MERGE of `OData__ModerationStatus`, one extra
      field, checked ONCE per run (`p.listTitle` is fixed for the panel's lifetime) rather than once
      per file. `approverEmail` is a new required prop, threaded from
      `BulkApproveCommandSet.onExecute` → `openBulkApprovePanel` → the component, because the panel
      has no other way to know who is running it — `SPHttpClient` carries no identity of its own.
  - **✅ DONE 2026-09-01: `Audit — approval activity`'s `ActorName`/`ActorEmail` read `ApprovedBy`
    first and fall back to `Editor` only when blank** — never the reverse, or a genuinely blank stamp
    (a document approved via the native Approve/Reject command, or one approved before this shipped)
    would report no actor at all instead of the imperfect-but-present `Editor` guess. **⚠ THAT FLOW NO
    LONGER WRITES THE `Approved` ROW AT ALL — Auto-route does. See the ownership section at the end of
    this file.** The fallback still matters for `Uploaded`/`Rejected`.
  - **⚠ STILL TO DO, and it is Power Automate, not code:**
    - `Condition 2` in BOTH `Auto-route` and `HC Auto Route` should compare `Author.Claims` to
      `ApprovedBy` instead of `Editor.Claims`. **Until this is done, the INTERIM MITIGATION above
      (delete the second row of `Condition 2`) is what is actually keeping emails flowing** — the new
      column exists and is being written, but nothing downstream reads it yet.
  - **NOT YET RE-TESTED LIVE.** The next approval through either route should show `ApprovedBy` =
    the approver's email on the item, distinct from `Editor`, which proves the write independently of
    whatever the flows end up doing with it.


## ⚠⚠ A SECOND MERGE AFTER `approve()` SILENTLY RESET THE ITEM TO PENDING (2026-09-01, 1.0.347.0)
The `ApprovedBy` stamp added in 1.0.344.0 was sent as a SEPARATE MERGE immediately after `approve()`
succeeded. Both calls returned **204**. The item was still **Pending** afterwards.
- **ON A MODERATED LIST, ANY EDIT TO AN ITEM CAN RESET ITS MODERATION STATUS unless that same write
  re-asserts it.** So the stamp silently undid the approval it existed to record. Fixed by folding
  `OData__ModerationStatus: 0` into the SAME MERGE as `ApprovedBy` — one write, one intent.
- **⚠ THE SYMPTOM NAMED EVERYTHING EXCEPT THE CAUSE, AND COST FIVE HOURS.** The page showed *Approval
  Successful*, both REST calls were 204, Auto-route never fired (there was nothing approved to route),
  and the audit flow read `{ModerationStatus}: Pending` and — correctly — logged `Uploaded`. Five
  theories were built and disproven first: a `304` on the config read, a stale bundle, GHO's `Levels`
  row, `Get_item` staleness, trigger polling lag. **The reading that settled it was the cheapest one
  available the whole time: open the `Approve/reject Items` view and see the document still sitting
  there.** When an approval "works" but nothing downstream happens, check the item's actual moderation
  status BEFORE theorising about anything that reads it.
- **`BulkApprovePanel.tsx` NEVER HAD THIS BUG** — it has always written `ApprovedBy` inside the same
  MERGE that sets the status. Only the approval page's separate follow-up write was affected. That
  asymmetry is the tell, and it is worth checking first whenever one approval route misbehaves and the
  other does not.
- **THE STANDING RULE: on a library with content approval on, every write to an item must carry the
  moderation status it intends to leave behind.** There is no such thing as a field-only MERGE there.

## THE `Approved` AUDIT ROW IS WRITTEN BY AUTO-ROUTE, NOT THE AUDIT FLOW (2026-09-01)
`Audit — approval activity` POLLS. The approval event lives on the SOURCE item, and Auto-route
**deletes that item seconds later**. So the flow was racing to poll a row being destroyed, and it
cannot win that race by any margin.
- **BOTH FAILURE DIRECTIONS WERE OBSERVED LIVE:** `GetUniqueId` failing with *"Item does not exist. It
  may have been deleted by another user."* (polled too late), and an approval producing **no run at
  all** (the item was gone before the poll — SharePoint cannot report a change on a row that no longer
  exists, so the event is simply lost).
- **⚠ NO DELAY VALUE CAN WORK, AND 15s WAS TRIED.** Too long and the item is deleted first; too short
  and `ApprovedBy` has not landed. The two constraints pull in opposite directions.
- **⚠ THE DELAY ALSO CREATED A SECOND RACE, WITH A DISTINCTIVE FINGERPRINT.** The UPLOAD-triggered
  instance sleeps, wakes, re-reads the item as now-Approved, and logs the `Approved` row using **its
  own** trigger data — the uploader's. The real approval instance then finds an `Approved` row already
  logged and correctly skips. **The tell is an `Approved` row attributed to the uploader AND NO
  `Uploaded` ROW AT ALL for that document** — one instance did the wrong job and consumed the other's
  slot. Seen on `test - test - 22` and `today - today - today`.
- **THE FIX IS OWNERSHIP, NOT TIMING.** Auto-route triggers on the approval, reads the item fresh
  (`Get_item`, which carries `ApprovedBy`), and already writes an audit row before its own delete. It
  has no polling race with itself, so it is the only component that provably observes the approval.
  - **`Auto-route` gained `Create item 1`**, placed BEFORE the existing `Create item`: EventType
    `Approved`, LibraryName `Approval for Document`, ItemPath `@triggerOutputs()?['body/{FullPath}']`
    (the SOURCE path), ActorName/ActorEmail `body('Get_item')?['ApprovedBy']` with a
    `triggerOutputs()?['body/Editor/...']` fallback. **ItemUniqueId stays
    `body('GetSourceUniqueId')?['UniqueId']`** so *History of this file* still threads Uploaded →
    Approved → Routed on one GUID.
  - **`Audit — approval activity` lost its Delay entirely.** `EventKind` now yields **`Skip`** for
    Approved, and `Not already logged` gained a second condition row: `outputs('EventKind') is not
    equal to Skip`. It owns **`Uploaded` and `Rejected` only** — both are states where the item still
    exists, so neither has a race and neither needs a delay.
- **⚠ BOTH AUDIT WRITES IN AUTO-ROUTE CARRY `Run after` = is successful AND has failed.** Pasting an
  action in front of an existing one defaults to **success-only**, which would let a failed audit write
  block the `Routed` row, the metadata stamp and the approval email. **An audit write must never stop a
  document routing** — the rule already recorded for `Get source author`, now applying to two actions.
- **⚠ `HC Auto Route` HAS NOT BEEN GIVEN THE SAME `Approved` ROW.** Same clone, same trap as every
  other HC clone in this project. Until it is, HC approvals log no `Approved` event.
- **NOT YET TESTED LIVE.** The test is: upload as an uploader, approve as a DIFFERENT approver, expect
  three rows — `Uploaded` (uploader), `Approved` (approver), `Moved to …` — with no delay anywhere and
  no dependence on how fast the two actions follow each other.

## ⚠⚠ `ApprovedBy` NEVER WROTE BECAUSE SHAREPOINT REJECTS THE COMBINED MERGE — FOUND AND FIXED (2026-09-01, 1.0.348.0)
Three genuine two-account tests (different uploader/approver, one via a fresh incognito window)
all showed the `Approved` audit row attributing the **uploader**, not the approver, even after the
`Auto-route` race-condition fix (below) landed correctly. Chased through the flow's own Network tab
during a live approval:
```
POST .../items(1816)  →  500
"message": "You cannot change moderation status and set other item properties at that same time."
```
- **THE 1.0.347.0 "FIX" COULD NEVER HAVE WORKED.** It combined `ApprovedBy` and
  `OData__ModerationStatus: 0` into ONE MERGE specifically to stop a separate MERGE reverting the
  item to Pending — but SharePoint's REST API rejects outright any MERGE that sets moderation status
  alongside any other field, full stop. Every approval 500'd on this call, silently, inside the
  existing `try/catch`. Approval and routing kept working regardless, because the earlier `approve()`
  call had already set the status independently — which is exactly why this was invisible for so long.
- **`OData__ModerationStatus` + `OData__ModerationComments` TOGETHER IS FINE** (both are moderation
  fields) — it is specifically an unrelated field like `ApprovedBy` in the same body that 500s. The
  `Rejected` branch, which only ever sends those two, was never affected.
- **FIXED IN BOTH PLACES THAT HAD THIS SHAPE** — `ApprovalDocument.tsx` and `BulkApprovePanel.tsx`
  (the bulk-approve panel had the identical combined body and would have hit the same 500 the first
  time anyone bulk-approved a document on a library carrying the column; nobody had yet). The fix:
  write `ApprovedBy` **alone** first (an ordinary field edit, does not hit the restriction), then
  re-assert the real outcome — `approve()` again in the approval page (a dedicated method, not a
  field MERGE, so it doesn't hit the restriction either), or the `ModerationStatus`+`Comments` MERGE
  again in the bulk panel. Order no longer matters for the end state; only the field-combination did.
- **VERIFIED LIVE 2026-09-01**, after rebuilding via `npm run build` (not `package-solution` alone —
  see the register a few lines below) and bumping to 1.0.348.0 so the App Catalog update was
  unambiguous: `items(1817)` MERGE returned **204**, and the `Approved` audit row correctly read
  `clarencechojinheng@gmail.com`, distinct from the `Uploaded` row's `chocheetuck4`.
- **THE DIAGNOSIS TOOK FOUR WRONG THEORIES FIRST** — a race condition (real, fixed, but not the whole
  story), connector schema caching, read permissions, a stale package — each ruled out by a cheap,
  direct check (fixing the ordering, re-checking the field as the approver's own account, verifying
  the shipped bundle's hash and content) before the Network tab during an actual live click settled it
  in one screenshot. **When silent failure is this persistent, watch the request itself rather than
  reasoning about it further.**

## ⚠⚠⚠ `heft package-solution` DOES NOT BUILD — SAME LESSON, PAID AGAIN (2026-09-01)
While chasing the above, the deployed `1.0.347.0` package was suspected stale purely from this
project's own extensive prior history of that exact failure mode (see the September 30 register
earlier in this file). It turned out NOT to be the cause this time — the real bug was the combined
MERGE — but the check was worth doing anyway: `npm run build` was run fresh, the resulting bundle
hashes for both changed files were confirmed different from what shipped in `1.0.347.0`, and the
version was bumped to `1.0.348.0` specifically so there was no ambiguity in Site Contents about
whether the new package had actually been picked up. **Bump the version on every fix that touches
already-shipped code, even a same-session one — it costs nothing and removes one entire class of
"did it actually deploy" doubt from the next diagnosis.**

## ⏭ WHERE TO PICK UP (2026-09-01, end of session)
1. ~~Test the two changes above~~ **DONE.** The race condition and the real combined-MERGE bug are
   fixed and verified live.
2. ~~Clone the `Approved` row AND the `ApprovedBy` fix into `HC Auto Route`~~ **DONE.** `Get ApprovedBy
   HTTP`, `Create item 1` (the `Approved` audit row) and the corrected `runAfter` wiring are all cloned
   into `HC Auto Route`, plus the `ActorName` display-name fix (split on `@` rather than showing the
   raw email) applied to BOTH flows. Verified live 2026-09-01 with a genuine two-account HC test:
   `Uploaded` (uploader) / `Approved` (approver, name-only) / `Moved to Highly Confidential Document`,
   all three rows correct.
3. **`Condition 2` (self-approval email suppression) IS FIXED AND VERIFIED IN BOTH `Auto-route` AND
   `HC Auto Route` (2026-09-02).** The normal flow's `Condition 2` now reads: send the "your file has been
   approved" email unless `BulkImport` is true, OR unless `ApprovedBy` is set and matches the uploader's
   own email (case-insensitive, compared on `Author.Email`/`ApprovedBy` directly — never `Editor.Claims`,
   which is never the approver, see the section above). Verified live both ways: different
   uploader/approver → email sent; same person uploads and self-approves → email suppressed, confirmed
   via the audit log AND an empty inbox. **Also required a code fix**: `Form.tsx`'s self-approve MERGE
   (a THIRD approval write path, alongside `ApprovalDocument.tsx` and `BulkApprovePanel.tsx`) never
   stamped `ApprovedBy` at all — fixed the same way, as a separate prior MERGE before the moderation
   status flip. Shipped as `1.0.349.0`.
   - **Built via the row/group builder, not Code view — Code view is read-only for a Condition
     action.** The visual builder also refused to let a group or the top level drop below 2 rows even
     after a real delete, so two harmless filler rows were added instead (`0 is equal to 1` inside the
     `Or` group, `0 is equal to 0` at the outer `And` level) — both no-ops, confirmed via Code view
     after saving. If rebuilding this in `HC Auto Route`, expect the same friction and use the same
     filler-row trick rather than fighting the deletion.
   - **✅ CLONED INTO `HC Auto Route` AND VERIFIED LIVE 2026-09-02.** Built the identical `Or` group
     (`ApprovedBy` blank, OR `toLower(Author.Email)` ≠ `toLower(ApprovedBy)`, plus the same two filler
     no-op rows) against HC's own `Get item` action — confirmed via Code view to match the `Auto-route`
     original exactly. Tested both directions on real accounts: `hga-hga-hga-02-09-26.pdf` (uploaded by
     `chocheetuck4`, approved by `clarencechojinheng` — cross-person) sent the "Your file has been
     approved" email to `chocheetuck4` correctly; `ss-ss-ss-02-09-26.pdf` (uploaded AND self-approved by
     `clarencechojinheng` via `autoApproveOwnUpload`) correctly sent **no** email — confirmed absent
     from a freshly refreshed inbox.
   - **⚠ ONE STEP NEARLY WENT WRONG AND IS WORTH RECORDING: mid-edit, navigation landed back on the
     LIVE, already-verified `Auto-route` flow with the half-built condition still showing in an unsaved
     state.** Caught before saving by checking the breadcrumb (it read "Auto-route approved Pending
     files to Documents Library", not "HC Auto Route"). **Always confirm the flow name in the
     breadcrumb before saving a condition edit** — Power Automate's back/forward navigation between two
     similarly-structured flows does not make clear which one is currently open.
4. **⚠ A NEW, SEPARATE GAP FOUND WHILE TESTING #3 — logged for next session, NOT diagnosed.** When a
   Head of Unit's own upload is instantly self-approved (`autoApproveOwnUpload` on), the Audit Log is
   missing its **`Uploaded`** row entirely — only `Approved` and `Moved to …` appear. A normal upload
   (pending, approved later by someone else) always gets its `Uploaded` row. Working theory, not yet
   confirmed: `Audit — approval activity` polls on created-or-modified, and `EventKind` reads `Skip` for
   an already-Approved item (since 1.0.NNN, `Approved` is logged by Auto-route itself, not this flow) —
   if self-approve flips the status within the same polling window as the item's creation, the flow may
   only ever see the FINAL (Approved) state on its one firing, so it evaluates `Skip` and the `Uploaded`
   event is lost with no second firing to catch it. **Next session: check `Audit — approval activity`'s
   run history for a self-approved document to confirm whether it fired once as `Skip`, fired twice, or
   never fired at all — that will show which of several possible fixes (a distinct trigger condition, a
   separate write inside the self-approve code path itself, etc.) is right.** Do not guess at a fix
   before seeing the actual run history.
5. **⚠ CLAUDE.md is now large enough to be a real per-turn cost** in every session, since it auto-loads
   in full. Much of it is historical incident write-up that could move to `docs/` and be read on demand.
   Worth a deliberate trim as its own reviewable change.

## THE ARCHIVE LIBRARY-ROOT PRUNE TOOL — BUILT, RUN ON BOTH SITES, THEN REMOVED (2026-09-02/03)
Following the folder-level archive access narrowing (client-facing repair tool, built and removed in
an earlier session per *"we don't want to confuse the client with a new feature like this"*), a SECOND
gap was found live: `chocheetuck4` (an ordinary uploader, no C-Level role) could open
`Archive Restricted & Confidential Document` — reach the LIBRARY, though no folders inside it were
visible. **The folder-level prune never checked the library object's own ACL** — a stray direct grant
at the library root, invisible to a folder-scoped pass.
- **A second one-time repair tool was built in `FolderManager.tsx` (`runArchiveRootPrune`), following
  the same pattern**: read the library's LIST-scoped `roleassignments` (not folder-scoped — a
  library root has no `ListItemAllFields`, and the first attempt's `GetFolderByServerRelativeUrl(...)
  /ListItemAllFields/roleassignments` 404'd for exactly that reason), keep only C-Level roles
  (GLOBAL/SEGVIEW derived from the live Group Map, plus site Owners), remove everything else via
  `removeroleassignment(principalid=X, roledefid=Y)` per binding (list-scope removal needs BOTH ids,
  unlike the folder-scope one-call-removes-everything shape).
- **⚠ TWO REAL BUGS CAUGHT DURING BUILD, both self-diagnosed from the tool's own live output before
  the user had to point either out:**
  1. **Wrong REST endpoint** (first version, `1.0.368.0`) — 404 on both libraries, fixed by switching
     to the list-scoped endpoint, resolving the list GUID via `getbytitle(...)?$select=Id` first.
  2. **A false "✓ Done — nothing to remove, already correct" on a run where BOTH library reads had
     just FAILED with 404.** `removed === 0` was being read as "confirmed clean" when it actually meant
     "never looked." Fixed with a separate `anyReadFailed` flag: a failed read now reports "Stopped
     early... nothing was skipped by mistake; try again," never a false all-clear. Same "empty ≠
     unknown" principle this codebase repeats everywhere else.
- **✅ RUN ON BOTH SITES, `1.0.369.0`.** Test site: hundreds of stray grants removed. SDG: **1272**
  stray grants removed across both archive libraries' roots (Restricted & Confidential + Highly
  Confidential), zero errors on either run.
- **✅ BOTH PRUNE TOOLS ARE NOW REMOVED — the library-root one in `1.0.381.0` (2026-09-03), the
  folder-level one in `1.0.365.0`.** Client, on seeing it still sitting on the Reconciliation tab:
  *"umm you might want to remove this"*. Same reasoning both times: a permanent button offering to
  strip permissions is exactly what an admin presses to find out what it does, and these were repairs
  for grants issued under the pre-2026-09-02 rule, not features.
  - `runArchiveRootPrune`, its three `useState` hooks and its JSX are gone; **`removeRoleAssignment`
    returns to its documented "unreachable on purpose" state** with its `eslint-disable` restored, and
    its comment now records that the library-root pass deliberately did NOT use it — a library root has
    no `ListItemAllFields`, so it needs the LIST-scoped
    `lists(guid'…')/roleassignments` and a `roledefid` per removal. That distinction cost a 404 the
    first time and is the one thing worth not rediscovering.
  - **Recoverable in minutes, not lost:** both shapes are in
    `docs/superpowers/specs/2026-09-02-archive-audit-log-and-access-narrowing-runbook.md`. Start there
    if a fourth segment is ever onboarded under the old rule.
  - **⚠ REMOVING THE TOOLS DOES NOT UNDO THEIR WORK, and does not need to.** The prunes only ever
    REMOVED grants, so non-C-Level archive access is already gone on both sites; a reconciliation
    re-run can only re-ADD the mapped C-Level grants, never restore anyone else. The access half of
    the objective therefore stands on its own.
  - **✅ CLARENCEDMSTESTING'S CONFIRMING RUN COMPLETED CLEAN (2026-09-03).** Every folder in both
    archive libraries reported `already locked, skipped` and **`2 group grant(s) already correct,
    skipped`** — exactly `GLOBAL` plus that segment's `SEGVIEW`, and nothing else. So the prune took
    only what it should have and the surviving C-Level grants are intact. The run was clean enough to
    reach its guarded passes: `Folder Map: no orphaned rows ✓`, `CRS Term Abbreviation: no orphaned
    rows ✓`, `Folders: every folder maps to a live term ✓`.
    - **The "2" is the useful number to re-check on any future run.** Three or more on an archive
      folder means something has been granted there that the 2026-09-02 rule excludes.
  - **✅ SDG'S CONFIRMING RUN COMPLETED CLEAN TOO (2026-09-03)** — all three segments, every archive
    folder `already locked, skipped`.
  - **⚠ SDG SHOWS `1 group grant` PER ARCHIVE FOLDER WHERE CLARENCEDMSTESTING SHOWS `2`, AND THAT IS
    NOT A FAULT.** Verified by REST: SDG holds `GHO_SEGVIEW` (19), `MHO_SEGVIEW` (343) and
    `NBPOLHO_SEGVIEW` (609) and **no site-wide C-Level group at all** — no `GLOBAL`, no
    `C_LEVEL_GLOBAL`. So the single grant is that segment's own C-Level, which is correct; the test
    site has two because `clevel_global` was given its first route to existing on 2026-09-02 (added to
    bulk provisioning) and only that site has had a bulk run since. **The two sites diverged that day,
    not before.** Do not "repair" SDG on the strength of the count.
    - **Open, and a client decision rather than a defect:** whether SDG wants a site-wide C-Level. If
      so, one bulk provisioning pass creates and maps `C_LEVEL_GLOBAL`, then reconciliation grants it.
    - **Also open: the `_SEGVIEW` → `_C_LEVEL` rename has NOT been applied on SDG**, though the client
      asked for that naming on 2026-09-02. The persona edit only changes what NEW names are written;
      Group Management's rename card migrates existing groups, and a SharePoint rename preserves the
      group Id so every mapping row, folder grant and page ACL survives untouched.
  - **Verified**: `tsc --noEmit` clean, full suite **1608/0**, this file's lint unchanged from its
    baseline (3 `no-new-null`, 1 `max-lines`), and the shipped bundle grepped for `Prune archive` —
    **0 hits**.

## "GROUP LED PROJECT" — FIRST PROJECT-CATEGORY SEGMENT EVER CREATED, VERIFIED END TO END (2026-09-02)
First test of the `Category = "Project"` half of the segment model (every prior segment has been
`BusinessSegment`). Term set created via CSV import (`docs/term-store-import/13-group-led-project-test.csv`,
placeholder data), segment created through `SegmentCreator`, uploaded, approved, routed — all verified.
- **`SegmentCreator`'s "Segment name" label is now conditional on `family === "Project"`**, reading
  "Project name" (placeholder "Group Led Project") instead of "Segment name" — the field itself was
  always correct data-wise (writes to the same `Business_x0020_Segment`/`BusinessSegmentTid` columns
  regardless of category; there is no separate physical "Project Name" column), only the LABEL was
  generic.
- **⚠ THE UPLOAD FORMS' "Upload to" RADIO AND "Segment" FIELD LABEL ARE HARDCODED TO "Group Led
  Project" / "Project", PER EXPLICIT CLIENT INSTRUCTION** (`Form.tsx`, `BulkUpload.tsx`), not derived
  from the mode row's actual label. The client asked for this literally — *"Rename the Project Button
  to Group Led Project"* — overriding the more scalable option (deriving the label from whichever
  Project-category segment exists). **This WILL need revisiting the day a second Project-category
  segment is onboarded**, since the radio can only ever say one thing.
- **⚠ FLAGGED BY THE CLIENT, NOT YET BUILT: the "Business Segment" label in `ApprovalDocument.tsx`
  and `MySubmissions.tsx`'s detail views is still HARDCODED, unlike the upload forms.** Client's own
  words: *"I notice that the label is not dynamic such as the one showing in ApprovalDocument.aspx or
  My Submission... in the future they might be different structure naming so it needs to be updated as
  well."* Both screens have their OWN fixed metadata field lists (documented gap, `documentDetails.ts`
  vs. `ApprovalDocument.tsx`'s hand-copied array) and would need to read the document's mode/category
  to decide "Business Segment" vs "Project" per row. **Not started.**
- **END-TO-END VERIFIED**: term import → segment created → folders/groups provisioned by reconciliation
  (`GLP_C_LEVEL`/`C_LEVEL_GLOBAL` correctly locked the new segment's folders, same as every other
  segment — proof the archive/C-Level-only rule from the section above applies automatically to a
  segment that did not exist when that rule shipped) → uploaded → approved → routed, metadata intact.

## A UI POLISH BATCH ACROSS THREE SCREENS (2026-09-03)
Client pasted ~14 annotated screenshots covering `BulkUpload.tsx`, `Form.tsx` and `MySubmissions.tsx`
while Folder Reconciliation ran in the background on SDG. Built as `1.0.371.0`.
- **Bulk Upload's Legally Privileged info icon sat 24px from its checkbox** — it was a separate flex
  child of `.dms-detail-row`, inheriting the row's own gap instead of sitting close to what it
  explains (the way Confidential Level's icon does via `.dms-labelrow`). Wrapped both in one flex
  container with a 6px gap.
- **⚠ DRAGGING MORE FILES ONTO AN ALREADY-PICKED LIST DID NOTHING** in Bulk Upload — only the EMPTY
  dropzone (`picked.length === 0`) had `onDrop`/`onDragOver` handlers; once files were staged, "Add
  more" was click-only and a drop just navigated the browser to the file. Client: *"I tried to drag and
  drop by adding more files, but system didn't allow."* Fixed by adding the same handlers (and a subtle
  dashed-outline hover state) to the already-picked list's wrapper, reusing the existing `addFiles`.
- **The upload form's "All fields marked * are mandatory" line was shown TWICE** — once at the top of
  the page, again above the saved-batches list. Removed the second occurrence.
- **"HOD/HOU" → "approver"** in My Submissions' delete/share request dialog copy, matching the client's
  given wording. **93 days KEPT, not the 90 in the client's new mock** — same conflict, same resolution
  as the one already recorded and confirmed with the client on 2026-08-30 (*"stay 93, they might not
  know that is why they say 90"*); applied without re-asking, per that precedent.
  "Head of Unit" → "approver" in the note list at the bottom of the page too.
  "asked" → "requested" throughout (row tag, inline text) — matches the dialog's own wording.
- **"View and edit" removed from the share-permission dropdown** — only "View only" now shown, as
  fixed text rather than a control. `permission` state stays `"View"` always; `SharePermission` itself
  and `Requests.tsx`'s reading of it are untouched, so this is a display-only narrowing on the
  request-RAISING screen alone.
- **The second remaining "Batch reference BAT-..." display removed** (the batch-detail "Document
  folder information" card) — the client asked for this exact string removed once already, 2026-08-30,
  from a table two levels up; this was the one occurrence that survived that pass. Still stored, still
  what groups the files — only the display is gone, on both screens now.
- **Auto-scroll (60vh) added to all three tables on My Submissions** — the submissions list, the
  Requests tab, and the main status tabs (All/Pending/Approved/Rejected) — client: *"it is sooo long of
  a list."* Same pattern this codebase already uses elsewhere; checked first that nothing absolutely
  positioned lives inside any of the three tables (the recurring trap — a scroll cap has clipped a
  popover on three other screens in this project).
- **The clash/rename-offer popup's warning icon changed from amber (`#FF952A`) to RED (`#FF4646`)**,
  matching the client's own reference graphic verbatim (`Group 70858.svg` — three nested circles plus
  an exclamation, geometry unchanged, colour swapped) on BOTH the upload form's dialog and Bulk
  Upload's identical one, for consistency (the existing code comment already stated the two must look
  like the same system). Title text colour moved from amber-brown (`#8a4b00`) to this codebase's
  established danger red (`#a4262c`). **⚠ THIS IS NOT THE OLD "Replace Existing File?" PROMPT** — that
  one was genuinely removed on 2026-08-15 and stays removed; CLAUDE.md's own record of that removal is
  still accurate. The screenshot the client called "the popup mockup for file replacement" was this
  clash/rename-offer dialog, confirmed by matching the icon geometry exactly against the live code.
- **⚠ STILL OPEN, NOT ADDRESSED THIS PASS:**
  - **The pending-request "chip" beside the file status** (`s.pendingChip`, "Deletion pending"/"Share
    pending" in the detail view) — the client's mockup showed a visual redesign for this specific
    element and the exact target pixel design could not be confirmed without re-seeing the image; the
    wording/data flow around it (dialog copy, "requested" vs "asked") was fixed, but the chip's own
    styling was left as-is pending clarification.
  - **"How come no details showing?"** — a reported bug in a batch/file detail drill-down, not yet
    diagnosed. Needs reproduction (likely `batchText`/`ftFor` returning empty for every file in a
    batch, or a per-item `FieldValuesAsText` read failing silently) before a fix.
  - **One ambiguous annotation ("Just remove this")** pointing near either the Remark field or "Open in
    a new tab" in a single-file detail view — left untouched, genuinely unclear which element without
    the source image.
  - A "Remove text" annotation paired with the "share requested"/"deletion requested" wording change
    was not separately actioned — unclear what additional text (beyond the wording already fixed) it
    referred to.
- **Verified**: `tsc --noEmit` clean, full suite **1598/1598**, 31 lint warnings (established baseline,
  zero new — every file's warning count matches its documented pre-existing set). Packaged as
  `1.0.371.0`, and every change spot-checked present in the shipped `.sppkg` bundle (grep for the new
  strings/colours in each affected `ClientSideAssets/*.js`) before reporting done. **NOT yet
  site-tested** — none of this batch has been clicked through on a live site.

## A SECOND UI POLISH ROUND, PLUS THE FIRST REQUIRED-REASON RULE (2026-09-03)
Same day, following up on the batch above. Two more small fixes and one behavioural change, all
shipped and bundle-verified.
- **`ApprovalDocument.tsx`: removed the two supplementary "Open in a new tab" links** (image-preview
  and iframe-preview branches). The `preview.kind === "none"` branch's own link is UNTOUCHED — that
  one is the sole route to the document when nothing can render, not decoration, and removing it
  would have taken away the only way to see certain files.
- **`MySubmissions.tsx`: found and fixed "no details showing" while investigating it.**
  `fetchFieldText`'s library-title resolution checked Documents and the HC pair but never the
  Archive pair — so opening the detail view on any archived file silently queried the wrong list,
  404'd, and showed "no details were recorded." Extended the same resolution ladder to check
  `cachedArchiveLibraries()`, matching the pattern already used elsewhere in the file to TAG a row as
  archived, just never before used to RESOLVE it. Same file's own "Open in a new tab" removed too
  (client screenshot of `My-Submissions.aspx`'s detail view), same reasoning as ApprovalDocument.
- **The pending-request chip on My Submissions redesigned to match the client's exact reference
  image**: status badge and the "Deletion pending"/"Share pending" chip now share their OWN row
  (previously crammed onto one line with "Uploaded ... path"), which moved to a separate line below.
  Chip background is now the client's exact `rgba(255, 225, 159, 1)`. Added 🗑/📤 icons before the
  text (only Deletion was in the reference image; Share's icon is a reasonable guess, flagged as such).
- **⚠ A NOTE IS NOW REQUIRED FOR EVERY DECISION ON `Requests.tsx`, APPROVE OR REJECT** — client's own
  capitals: *"IF APPROVER WANTS TO REJECT A REQUEST ... APPROVER MUST INCLUDE THE REASON ... ELSE
  SYSTEM DOESN'T ALLOW TO PROCEED."* **THIS REVERSES the 2026-08-30 design**, whose own comment
  argued an approval "needs none" — that argument is superseded, not wrong for its time. Gated the
  same way every other required-field check in this codebase is: never on load, only after a
  decision was attempted with nothing typed (`showNoteError`, reset every time the dialog reopens).
  The dialog's copy also now matches the client's exact mockup: title stays "Approve/Reject this
  request?", body for a REJECT reads `"{filename} will not be touched, and {requestedBy} will see
  your note."` (APPROVE keeps the existing `decisionSummary()`, which already states the real
  effect — replacing it with "will not be touched" would be false for an approval), and the label is
  now `"Note — say why; the requester sees this"` for both. **One dialog serves both Delete and
  Share requests**, so this single fix covers the "File Share Request" reject-validation ask too.
- **`ApprovalDocument.tsx`: the "already decided in this session" state is now a clean read-only
  summary, not the same form disabled.** Client: *"having clickable radio buttons and CTA buttons
  (Already Decided & Cancel) displayed out in the layout is really confusing."* Before this, every
  control (radios, comment textarea, Approve/Cancel buttons) stayed on screen, merely `disabled` —
  which still READS as an interactive form. Now: `STATUS` (badge), `COMMENT` (the typed text or "No
  comment was given"), then `"Reviewed by {displayName} on {date} — no further action is needed."`
  Reuses `formatDate` and the existing `s.navBadge`/`s.navBadgeOk`/`s.navBadgeNo` styles — no new
  helpers. The still-undecided branch (radios/textarea/buttons) is completely unchanged, just now
  the `else` half of one conditional instead of the only path.
- **Two items checked and found to need NO code change**: `BulkApprovePanel.tsx`'s "3 selected."
  text and its routing-flow copy — both already match what the client asked for in the CURRENT
  source (confirmed by grep: neither string exists). The client's screenshot is a stale deployed
  package. Their "I cant seem to trigger this" report is most likely the recurring
  `usercustomactions` registration gap this project has hit before with the bulk-approve command set
  — check `/_api/web/usercustomactions` before assuming a code defect.
- **Verified**: `tsc --noEmit` clean, full suite **1598/1598**, **zero lint warnings on every touched
  file** (not just baseline-matching — `Requests.tsx` and `ApprovalDocument.tsx` both lint clean).
  Packaged as `1.0.375.0` across several incremental versions (`.372`–`.375`), each spot-checked
  present in the shipped bundle before moving to the next fix. **NOT yet site-tested.**

## ⏭ THE "DOCUMENT DELETION & SHARING APPROVAL" REDESIGN IS SCOPED, NOT BUILT (2026-09-03)
Client's mockups asked for something much bigger than a copy fix: `Requests.aspx` renamed to
"Document Deletion & Sharing Approval", reachable through a NEW site-nav parent item
("Approval & Request") with three children, Delete File + Share merged from two tabs into one page
with two labelled sections, and "Shared files" demoted from a tab to a third section on the same
page. Client, offered the choice: *"I guess you can scope it out for now since its a big piece of
work for you."*
- **Full scope, what already exists vs. what's new, and four open questions to settle with the
  client before building: `docs/superpowers/specs/2026-09-03-requests-page-redesign-scope.md`.**
- **THE HEADLINE FACT: THE SITE-NAV RESTRUCTURING IS NOT CODE.** Adding a parent nav item with
  children is a manual `Edit navigation` action per site — a web part package cannot do this on
  install. The two "Approval (...)" children need no new page at all, just a nav entry pointing at
  the existing normal/HC `ApprovalDocument.aspx` pages.
- **The reject/approve dialog built minutes earlier THIS SAME SESSION (see above, `1.0.375.0`) needs
  NO further change** — it already matches the mockup's copy and already serves both Delete and
  Share requests from one component. The merged view just needs to call the same `setDeciding({...})`
  handlers from differently-laid-out cards.
- **Nothing in `Requests.tsx` has been touched for this redesign.** The tab bar
  (`tab`/`setTab`/`TABS`/`TAB_LABEL`/`tabCount`), the per-tab filter bars, and `requestCard` are all
  exactly as they were before this session. Do not assume any of this is in progress.

## ⏭ THE CODE HALF OF THE REDESIGN IS BUILT — `1.0.376.0`. THE SITE-NAV HALF IS STILL MANUAL, NOT DONE (2026-09-03)
Client answered the four open questions from the scope doc above within the same session — (1) card
buttons still open the existing decision dialog, (2) one shared filter bar, (3) *"remove the Waiting
for you, but just change the status from pending to Approved"*, (4) yes, the nav rename goes on SDG
too. `Requests.tsx` restructured accordingly; **the site-navigation half described in the scope doc
(`Approval & Request` parent nav item + three children) is NOT done — that is a manual `Edit
navigation` action, per site, and nobody has done it yet.**
- **THE TAB BAR IS GONE.** `TABS`/`TAB_LABEL`/`Tab`/`tab`/`setTab`/`tabCount` all removed. Delete
  Requests and Share Requests now render simultaneously, one below the other, each its own `<div
  style={s.card}>` with a header reading `🗑 Delete Requests` / `📤 Share Requests` plus an amber
  `N pending` pill (reusing `PILL.Pending`'s exact colour) — shown only when `N > 0`, per the mockup.
- **"REMOVE WAITING FOR YOU, JUST CHANGE STATUS FROM PENDING TO APPROVED" MEANT: ONE MERGED LIST PER
  TYPE, NOT TWO STACKED BOXES.** `typeSection(t, emoji, label)` replaces `typeTab`: pending rows
  (from `ofType(t)`, i.e. everything VISIBLE to this viewer, not only what they can decide) sort
  first, decided rows (`decidedOf(t)`, unchanged) follow. A row's own status badge — not which box
  it sits in — is what tells you whether it is still pending.
  - **⚠ `actionable` IS NOW `r.status === "Pending" && canDecide(r, scope)`, computed PER ROW** —
    previously the caller passed a blanket `true`/`false` per BOX (every row in "Waiting for you" was
    `true`). Necessary because the merged list can hold a Pending row a HoD can SEE but cannot
    DECIDE (the 2026-08-21 pending-stage rule) — showing Approve/Reject on that row would offer a
    button that fails at the write.
  - **`waitingOf`/`queue`-per-type is GONE, deliberately** — `queue` (used only for `queue.length` in
    the header badge) stays; the per-type `waitingOf` helper became dead code the moment pending rows
    started coming from `ofType` instead, and was removed rather than left unused.
- **⚠ THE STATUS FILTER GAINED A THIRD CLAUSE TO SURVIVE THE MERGE: `r.status === "Pending" ||
  statusFilter === "All" || r.status === statusFilter`.** Before the merge, "Waiting for you" was a
  SEPARATE box the status dropdown never touched at all — pending rows were structurally elsewhere.
  Now they sit in the SAME array the dropdown filters, so without the first clause, choosing
  `Approved` from one shared dropdown would silently empty every unit's outstanding work — the exact
  "no display control may hide pending work" rule this codebase already states elsewhere, paid again
  in a new shape. The text filter still applies to everything, matching what "Waiting for you" always
  did (with its own warning line, preserved as `hiddenPending`).
- **ONE FILTER BAR (`requestsFilterBar`), computed once, rendered once, above BOTH `typeSection`
  calls** — per the client's own confirmed answer. Its outcome dropdown's options are now the UNION
  of decided statuses across BOTH Deletion and Share (`STATUS_ORDER.filter(v => decidedOf("Deletion")
  ... || decidedOf("Share") ...)`), since one control has to serve two types that used to have their
  own.
- **"Shared files" is a THIRD, ALWAYS-RENDERED SECTION, not a tab** — retitled "Documents with
  Shared Access" per the mockup. Its own filter (`sharedState`/`sharedText`) is DELIBERATELY kept
  separate from the shared Delete/Share bar: it filters an ACCESS STATE (`live`/`unknown`), not a
  request STATUS, and folding the two together would give one control a value the other section
  cannot mean.
  - **⚠ ITS LAZY ACL PROBE'S GATE CHANGED FROM "tab opened" TO "page loaded."** The effect that
    calls `loadAcls` was gated on `tab !== "Shared files"` specifically so 700+ files were never
    probed until someone clicked that tab; with the tab gone, it now runs once on page load (still
    gated on `aclsRead`, so it is still exactly ONE probe per page visit, just no longer deferred).
- **PAGE HEADER RENAMED**: `<h2>Requests</h2>` + a long explanatory paragraph → `<h2>Document
  Deletion &amp; Sharing Approval</h2>` + the client's one-line subtitle. A new top bar shows
  `Requests {queue.length}` and an **Access Audit** link, resolved via the SAME `resolveLink`/
  `readSitePages` pattern every other admin-page directory link in this project already uses (own
  `useEffect`, independent of the requests-list read, so a failed Site Pages lookup cannot take the
  queue down with it — matching the `systemAdmin` effect right beside it).
## A UI POLISH BATCH — FOLDER MANAGEMENT, GROUP MANAGEMENT, PAGE ACCESS, AUDIT LOG (2026-09-03)
Client pasted ~11 annotated mockups. Built as `1.0.377.0`.
- **Folder Management landing page**: "Run folder reconciliation" now spans the full row
  (`gridColumn: "1 / -1"`), rather than sitting card-width alongside the four flow cards — it is the
  one card with no guided-flow sequence behind it, so it read as the odd one out anyway.
- **Group Management's "Create every group for a segment" is now "Create group"**, MOVED to render
  AFTER `GroupManager` (Quick Search + the group list + System administrators) in
  `GroupManagementPage.tsx`, not before it. The client's mockup 1 pointed at the "Create all groups
  for a segment / Create one group" toggle — **already removed in an earlier session this same day**
  (2026-09-02: *"remove the Create one group"*), so the screenshot was stale; the instruction was
  reinterpreted against the ONE create-group heading that still exists.
- **The whole "Create group" card follows the client's exact copy and list** (mockup 2, "Follow the
  exact copy, exact list"): info box rewritten (and its stale "set on Folder Access" reference is
  gone as a side effect — that page was retired 2026-08-23); the checklist relabelled and reordered
  to PIC / PIC (HC) / Head of unit / Head of unit (HC) / Viewer / Viewer (HC) / Head of department /
  C-level, via a LOCAL display map (`ROW_ORDER`/`ROW_DISPLAY` in `BulkGroupProvisioner.tsx`) rather
  than renaming `PERSONAS` in `shared/groupMapModel.ts` — that array is read by Folder Access's tier
  picker, Quick Search's persona summary and CSV headers, so renaming it there would relabel "SDG
  Employee" to "Viewer" everywhere at once, a far wider change than one screen's copy.
  - **`clevel_global` is APPENDED to the display list, not dropped**, even though the client's mockup
    only showed 8 rows — it was given its only route to existing anywhere in an EARLIER session this
    same day (2026-09-02, "there isn't any global viewer"), and dropping it now would silently
    reverse that fix.
  - **A new "View all N groups" / "Hide N groups" toggle** collapses the group-preview table by
    default — it did not exist before (the table always rendered once a plan was computed); the
    count moved off a separate "Preview — N groups..." line and onto the toggle button itself, in the
    SAME row as "Create missing groups" and "Export CSV" (renamed from "Create N group(s) and map M").
- **"What can this person reach?" → "Quick Search"**, with a shortened one-line description — same
  screen, same lookup, just a shorter heading and hint.
- **Two explanatory paragraphs removed** from Quick Search's result cards: the per-group persona
  SUMMARY sentence (the short persona label and the role chips already say the same thing) and the
  page-level footer ("Roles and scope come from the CRS Group Map...").
- **"nobody in it yet" → "No member"** on the group list's member-count badge.
- **System administrators' explanatory copy replaced verbatim** with the client's given text (a
  short intro line, a "Warning:" paragraph, and a bulleted three-cases list) — the underlying rules
  (per-user SCA flag, no report if changed outside this card, literal `_layouts` link) are unchanged,
  only the prose describing them.
- **Page Access**: intro shortened to "View page access permissions."; the read-only warning box
  shortened to drop the "set entirely by Folder Reconciliation from the CRS Group Map" explanation;
  the "site home page is not listed" hint and the "Only groups that currently hold a real grant..."
  footer both removed (the RULES behind both are unchanged, only the sentences are gone); **admin-only
  pages (CRS Settings, Approval Library Access, CRS Audit Log, CRS Configuration, Folder Access,
  Folder Manager, Page Access itself, Site Access) dropped from the page picker**, reusing
  `policyForPage(...).adminOnly` — the SAME flag reconciliation already locks these pages by, not a
  second name list that could drift. `isForbiddenPageTarget` itself is untouched, since it also gates
  reconciliation's own lockdown pass and widening it would risk unlocking pages it currently locks.
- **`StagingAccess.tsx`'s "Unable to verify current permissions" box changed from amber to red**
  (`noteWarn` style) — the one place in the codebase this exact box exists; the client's mockups for
  BOTH Group Management and Page Access referenced it as a shared design pattern.
- **⚠ ONE ITEM FLAGGED, NOT FIXABLE IN CODE**: the Audit Log's "Replaced" event detail text ("This
  upload replaced an existing document of the same name. The previous content is preserved as an
  earlier version in version history.") is written by the `Auto-route`/`HC Auto Route` Power Automate
  flow's `Create item` action, not rendered by `AuditLog.tsx` — that component just displays whatever
  string is stored in the `Details` column verbatim. Trimming the confusing second sentence requires
  editing the flow directly.
- **Verified**: `tsc --noEmit` clean, full suite **1598/1598**, only pre-existing baseline warnings
  (`FolderAdmin.tsx`'s `mapRows`, `GroupManager.tsx`'s max-lines) — no new categories. Packaged as
  `1.0.377.0`, every new string/colour spot-checked present in the shipped bundle. **NOT yet
  site-tested.**

- **✅ THE NAV SIDE WAS ALREADY DONE, ON BOTH SITES — confirmed by the client, 2026-09-03.**
  `Approval & Request` already exists on ClarenceDMSTesting as a parent nav item with three children
  (`Approval for Document`, `Approval for Highly Co...`, `Approval for File Requ...`), matching the
  scope doc's ask exactly, and the client confirmed SDG's nav is "the same." **The entry above
  calling this "still open, not code" was wrong** — written without having seen the client's site.
  **This whole redesign — code and nav, both sites — is DONE.** Only remaining step is a live
  click-through, which has not happened yet.
- **Verified**: `tsc --noEmit` clean, full suite **1598/1598**, **zero lint warnings** on the touched
  file (not merely baseline-matching). Packaged as `1.0.376.0`, every new string spot-checked present
  in the shipped bundle. **NOT yet site-tested** — none of this has been clicked through live.

## ⚠⚠ THE BATCH DETAIL VIEW SHOWED NO DETAILS — FOUR REPORTS, THREE CAUSES (2026-09-03, 1.0.385.0)
Client, testing My Submissions: *"nothing is showing for bulk upload in my submisison"*, *"I notice
that the archived files wont show Document folder information like a normal file does"*, and *"I
upload normally through an upload form but when I open the submission I dont see the details anymore
after approved."* All three are the batch view (level 3), and every one of them rendered as a card
that looked like a document nobody had filled in.
- **⚠ 1. THE READ WAS DRIVEN BY THE CLICK, AND THERE ARE TWO WAYS INTO THAT VIEW.** Level 2's batch
  row called `loadBatchText`; **level 1's row does not** — a bulk run has one destination, so it
  opens the batch directly (client's own ask, 2026-08-28) — and it never loaded anything. So
  `batchText` stayed `undefined` for the whole visit and every file read *"Reading the details…"*
  for ever. **This is the standing rule in this file, paid again: a NEW ROUTE INTO AN EXISTING PATH
  SILENTLY STARTS FROM ZERO** (the third instance, after `GroupMembersEditor`'s two failure paths
  and the six clash branches). Fixed by driving the load from a `useEffect` on
  `[openSubmission, openBatch, rows]` — **one route in, one loader** — so a future third entry point
  cannot reintroduce it.
- **⚠ 2. THE MAP WENT STALE THE MOMENT A FILE WAS APPROVED, and this was the one reported as "I
  don't see the details anymore after approved".** `batchText` is keyed by `mergedKey`, which for a
  live row is `library#itemId` — and **Auto-route COPIES the file into the approved-side library and
  deletes the source**, so it comes back with a different library AND a different item id (the same
  fact this file already records twice: it is why the submission record had to be keyed on a stamped
  `SubmissionFileId` and why the archive mover must use `MoveTo`). The map loaded while the file was
  pending then held a key nothing asks about again.
  - **`ftFor` MADE IT INVISIBLE by answering `{}` for a missing key.** `{}` means *read, and there
    was nothing*; a missing key means *not read*. Two different facts through one value — `empty ≠
    unknown`, in a new place. Worse, `buildFileRows` appends **File size** from the row rather than
    from the read, so the card rendered exactly one row and never tripped the "could not be read"
    branch: a document that looked like it had never carried metadata. It returns `undefined` now.
  - The effect keys on a **SIGNATURE** of the batch (`openBatch` plus every file's `mergedKey`), not
    a boolean, so the ids changing IS what triggers the reload. A superseded read is discarded by
    comparing the signature again before `setBatchText`, or the older one could resolve last and
    install a map keyed on ids that no longer exist.
  - **⚠ CLEARING `batchText` WITHOUT CLEARING THE SIGNATURE STRANDS THE NEXT VISIT** — reopening the
    same batch computes the same signature, the effect decides it is already loaded, and nothing
    fills the map it just emptied. The Back button clears both, together.
- **⚠ 3. A GONE FILE'S `ft` IS `undefined` BECAUSE IT IS NEVER FETCHED — AND THE LOADING MESSAGE WAS
  KEYED ON EXACTLY THAT.** So every archived and deleted file sat under *"Reading the details…"*
  permanently, **directly above its own details**. `readingFt` is now `!gone && ft === undefined`,
  and the "no details were recorded" branch covers a gone row whose snapshot is empty (a record
  written before snapshots existed) rather than leaving a blank card.
- **⚠ 4. THE FOLDER CARD HAD NOTHING TO READ FOR AN ALL-GONE BATCH, which is the archived complaint.**
  `destSource` is the first live `ft` with any keys, and a batch whose every file has been archived
  has none — so the card showed **Location and nothing else** while Business Segment, Department and
  Unit were rendered on the FILE card, one card below, where they do not belong.
  - **THE SNAPSHOT IS NOW SPLIT THE WAY THE VIEW IS SPLIT** — `snapshotFolderRows` /
    `snapshotFileRows` in `shared/submissionRecords.ts` (pure, 7 tests). The partition is **by KEY and
    is exact rather than a guess**: both upload screens write the per-document fields under eight
    fixed labels (`SNAPSHOT_FILE_KEYS`) and every other key is a tier, written as
    `snapshot[level.column] = level.label`. So *"anything not in that list is part of the folder"* is
    true by construction, and a segment onboarded years from now with tiers nobody has written code
    for partitions correctly with no change here.
  - **⚠ THE KEY LIST IS THE UNION OF BOTH WRITERS, `Bulk import` INCLUDED.** Leaving one screen's key
    out files it as a folder tier, so a bulk-imported document would claim its destination folder had
    a level called "Bulk import". Pinned by a test, along with "nothing lost, nothing shown twice".
  - Tier rows keep **insertion order**, which is chain order, because the writers loop
    `levelSelections` top down — so the card reads like the upload form with nothing here knowing any
    tier's name.
- **THE DIAGNOSIS IS THE REUSABLE PART, and it came from the screenshots rather than from theory.**
  Four screenshots of the same page distinguished all three causes: a bulk batch stuck on *"Reading
  the details…"* (cause 1, confirmed by its Back label reading *"Back to the list"*, which only a bulk
  group produces); the same five files stuck before AND after approval (cause 2); an archived file
  showing *"Reading the details…"* **and** its metadata at once (cause 3, which no loading state can
  do); and an approved file showing **File size alone** (cause 2's `{}` fallback, which is the only
  thing that renders exactly one row). **Read what the screen shows against what each branch CAN
  render before proposing a mechanism.**
- **Verified**: `tsc --noEmit` clean, full suite **1615/0** (7 new), 31 lint warnings — the documented
  baseline, zero new. Packaged as `1.0.385.0` via `npm run build`, and the shipped
  `my-submissions-web-part_*.js` grepped for both the new key array and the signature separator.
  **NOT yet site-tested.**

### ⚠ AND THE STATUS COLUMN WAS BLANK FOR AN ALL-ARCHIVED SUBMISSION (1.0.386.0, same day)
Reported the moment the above was deployed: *"the files in Status for Archive is not showing, the rest
is fine."* Everything else in that batch was confirmed working on site — the bulk batch, the folder
card for archived files, and the details surviving approval.
- **`countLine` COUNTED `archived` AND NEVER MENTIONED IT.** `recordCounts` returns all five states;
  the line pushed `deleted`, `cancelled` and `unknown`. So a submission whose every file had been
  archived had no approval counts (correct — `liveRowsOnly` excludes them) and no record count either,
  and rendered a **blank Status cell**.
- **⚠ SECOND TIME, IDENTICALLY.** `cancelled` was left out of the same line when it was added on
  2026-08-28, with the same symptom, and the comment recording that fix is four lines above where
  `archived` was then omitted. **A comment asking the next person to remember is not a guard.**
- **NOW DRIVEN BY `RECORD_STATE_LABEL`, a `Record` over the union** (`shared/submissionRecords.ts`,
  with `recordStateParts`, pure, 5 tests) — so **adding a member to `RecordState` is a COMPILE ERROR
  until it is given a word.** Same guard `RequestStatus` already uses for its pill colours, and the
  reason that union has never shipped a blank pill.
- The wording is deliberate and pinned: **`replaced`, never `cancelled`** (the file was superseded,
  not withdrawn) and **`not checked`, never anything implying loss** (it means a library read failed,
  so those files may be perfectly fine).
- **Verified**: `tsc --noEmit` clean, full suite **1620/0** (5 new), only the pre-existing `max-lines`
  warning on the touched file. Packaged as `1.0.386.0` via `npm run build`; shipped bundle grepped for
  the label map and its hash confirmed changed. **NOT yet site-tested.**

### ⚠ THE BANNER STILL DOES NOT SHOW, AND IT IS NOW SELF-DIAGNOSING (1.0.387.0, 2026-09-03)
Client: *"I notice the background image is still not showing"* — the second report of this, after the
2026-09-02 diagnosis established (from a real uploader-only account hitting SharePoint's own
AccessDenied page) that the cause is on the SharePoint side, not in the web part.
- **THE CODE IS CORRECT AND WAS RE-READ TO CONFIRM IT.** `backgroundImage` is two layers, image first
  and gradient behind, over a site-relative `<web>/SiteAssets/img_kv-banner.jpg`. **What is on screen
  IS the designed degrade**: a plain green banner, no broken-image icon, no error.
- **⚠ THAT SILENCE IS THE REAL PROBLEM, AND IT HAS NOW COST TWO ROUNDS OF GUESSING.** Degrading
  silently is right for a viewer and useless for whoever has to fix it: a **404** (nothing at that
  address) and a **403** (there, unreadable by this account) need opposite fixes, and nothing on the
  page could tell them apart. Nothing in code can fix the underlying cause — the file has to exist and
  the viewer has to be able to read it — so what code CAN do is stop the failure being invisible.
- **AN ADMIN-ONLY LINE NOW NAMES THE STATUS AND THE URL.** `isSystemAdmin` is checked FIRST, so a
  non-admin issues no probe at all and never meets a message about a file they cannot fix — the same
  rule as the unready-segment label. Reads the **status only**; the body is an image and is never
  parsed. A network-level failure reports NOTHING, because a blocked request says nothing about the
  file (the Brave Shields case of 2026-08-24).
- **⚠ IT NAMES BOTH CAUSES AND DELIBERATELY PICKS NEITHER.** SharePoint security-trims to **404** as
  readily as it answers **403**, so a 404 does not rule out permissions — it only establishes that
  this account cannot resolve it. Picking one would send half the admins the wrong way.
- **✅ THE CAUSE IS SETTLED, SAME DAY, AND IT IS THE PERMISSION ONE.** Opened directly as an
  administrator the image **displays**; opened as an approver (`clarencechojinheng@gmail.com`,
  `GHO_GCA_GCBC_APPROVER`) it answers SharePoint's own **You need access** page. So the file exists,
  the URL is right, and **Site Assets does not grant Read to `CRS_SITE_MEMBERS`** — the 2026-09-02
  diagnosis confirmed a second time, from the opposite direction.
  - **✅ FIXED AND VERIFIED THE SAME DAY, ON THE SITE, NOT IN CODE.** Site Assets → Permissions now
    reads `CRS Owners` **Full Control** + `CRS_SITE_MEMBERS` **Read**, and the banner renders for the
    approver account. ⚠ **Untick *"Share everything in this folder, even items with unique
    permissions"*** when granting it — the same trap already recorded for the Site Pages fix, where
    that tick would have pushed Read onto ten locked admin pages.
  - **⚠ THE LIBRARY GRANT DOES NOT COVER AN ITEM WITH ITS OWN ACL.** That Permissions page warns
    *"Some items of this list may have unique permissions which are not controlled from this page"* —
    so if a FUTURE asset fails while the banner works, check the file, not the library.
  - **⚠ THE TEST LEFT A PENDING ACCESS REQUEST.** The AccessDenied page offers **Request access**, and
    the same page now reports *"People are waiting for your approval so that they can access this
    site."* That queue entry is our own diagnostic clicking through, not a real need — dismiss it, and
    do not grant it: it would hand that account site-wide access nobody decided on.
  - **⚠ SDG NEEDS THE SAME CHECK.** Nothing asserts it, and the symptom is a plain green rectangle
    that reads as a design choice.
- **⚠⚠ AND THE ADMIN-ONLY GATE MAKES THAT NOTE BLIND TO EXACTLY THIS CAUSE.** An administrator CAN
  read the file, so they never receive a 403 — the note catches a wrong or missing URL, which affects
  everybody equally, and **cannot catch the permission case at all**. Its silence must not be read as
  "the banner is fine for everyone". Kept, because the missing-file case is real and it costs one
  admin-only request; the comment at the effect now says so outright.
  - **THE CHECK THAT WOULD ACTUALLY CATCH IT ASKS ABOUT A GROUP, NOT ABOUT THIS SESSION**, so it
    belongs in **reconciliation** beside the existing Site Pages assertion (`⚠ NOBODY CAN OPEN THE
    SITE HOME PAGE`) — identical shape, identical fail-open rule, and it would then cover SDG on its
    next run. **Seventh instance of "assert the required state every run". NOT BUILT.**
- **Verified**: `tsc --noEmit` clean, `eslint` clean on the touched file, full suite **1620/0**.
  Packaged as `1.0.387.0` via `npm run build`; shipped `document-search-web-part_*.js` grepped for the
  new message. **NOT yet site-tested.**

## ⚠⚠ THE METADATA FILTERS 500'D — THE 2026-09-02 FIX TRADED A 400 FOR A 500 (2026-09-03, 1.0.389.0)
Client, with all four Advanced Filters set and a nonsense search word: *"4 libraries could not be
searched: Approval for Document (HTTP 500), Restricted & Confidential Document (HTTP 500), Approval
for Highly Confidential Document (HTTP 500), Highly Confidential Document (HTTP 500)"*, and the ask
*"can you ensure it doesnt show this error but isntead notify them file not found."*
- **⚠ RELABELLING WAS REFUSED, AND THAT IS THE FIRST THING TO SAY.** A 500 means the query FAILED; it
  is not a statement that nothing matched. Showing *"file not found"* would tell the client no
  documents match when the search never ran — the one thing this page must never do, and exactly what
  the existing three-way empty state (`idle` / `empty` / `error`) exists to keep apart. The banner was
  right; the query was broken. **With the 500 gone, their own case now shows the existing
  *"No documents matched"* — which is what they were actually asking for.**
- **⚠⚠ TWO FAILURES IN OPPOSITE DIRECTIONS, BOTH INFERRED FROM ONE COMBINED TEST.** On 2026-09-02
  every filter answered **400**, the conclusion drawn was *"all four columns are term-set-bound"*, and
  the fix routed all of them through `substringof('v',TaxCatchAllLabel)`. That answers **500**, because
  **`TaxCatchAllLabel` is a hidden NOTE field and SharePoint cannot filter a Note field at all.**
  - **THE ORIGINAL DIAGNOSIS WAS WRONG AND THE EVIDENCE WAS ALREADY IN THIS FILE.** With four clauses
    in one `$filter`, **ONE bad clause fails the whole request**, so that test could never say WHICH
    column was at fault — "all four" was a formula fitted to a single measurement, the same error this
    file already records against the hero-banner offset. The archive-column work of the SAME DAY had
    checked the libraries directly via `SP.Taxonomy.TaxonomyField` on their `/fields` and found **only
    three of the five are taxonomy**: `Document Type`, `Year`, `Confidentiality Level`. Business
    Segment, Department and Unit are **plain TEXT** — which is why `ensureColumn` gives each a text
    `<Base>Tid` twin instead of relying on a taxonomy field's hidden note field. **Direct inspection
    beats inference from a combined failure.**
- **THE SPLIT, NOW PINNED BY TEST:** Business Segment and every tier → plain `eq` on their own text
  column, indexable, provably fine. Document Type / Year / Confidentiality Level → **never sent to
  `$filter` at all**, because `eq` 400s and `TaxCatchAllLabel` 500s and either takes the working
  filters down with it.
- **THE THREE ARE NARROWED ON THE ROWS INSTEAD** (`metadataFilterMatches`, pure, 6 tests). **Selecting
  a taxonomy column is a different operation from filtering one and is well supported**, so the labels
  travel with the row and no third piece of server-side syntax had to be guessed at — a third blind
  attempt is precisely what produced the 500. `textOf` already handles the `{Label, TermGuid, WssId}`
  shape.
  - **The `$select` retries WITHOUT them on 400 only**, so a library provisioned before those columns
    existed stays searchable. That library then reports **`unnarrowed`** and the page says outright
    that those three filters could not be applied to it — results shown, limitation named. Silently
    showing them would overstate the match; dropping the library would hide documents that may
    qualify. 400 only: a 403 is a permission answer and a 404 is the library.
  - Narrowing happens within the read's existing `$top=200` window, ordered by `Modified desc` — the
    same window the page already promised.
- **⚠ THE 500'S BODY IS NOW LOGGED.** SharePoint names the offending field and the reason in the
  RESPONSE BODY, and nothing was reading it — which is what made two blind fixes possible. Fourth time
  this project has learned it (gotcha #9, the tagging failure, the ACL read). Console, not the banner:
  the banner already names the library and the status.
- **Verified**: `tsc --noEmit` clean, `eslint` clean on both touched files, full suite **1628/0** (10
  new), 35 warnings — the full-lint baseline. Packaged as `1.0.389.0` via `npm run build`; the shipped
  `document-search-web-part_*.js` grepped for **zero** occurrences of `TaxCatchAll` and one each of the
  three column names. **NOT yet site-tested** — the test is the client's own case: all four filters
  set, expecting either results or *"No documents matched"*, and no 500.

## EVERY ATTENTION BANNER IS ONE PALETTE NOW, AND IT IS THE RETIRE-A-SEGMENT RED (2026-09-03, 1.0.391.0)
Client, with Bulk Upload's amber banner beside Folder Management's red *Retire a segment* one:
*"change the background color to follow the Retire a Segment, any banner that uses yellow pls follow
the Retire a Segment color"*, plus new copy for the Bulk Upload warning.
- **⚠ FIFTEEN BANNERS ACROSS THIRTEEN FILES, IN FOUR DIFFERENT AMBERS** (`#fff4e5`, `#fff4ce`,
  `#fff8e6`, `#fff8e1`) with four different text colours — each written where it was needed rather
  than taken from anywhere. So "change the banner colour" was a thirteen-file sweep with **no way to
  tell whether it had been finished.** `shared/noticeStyles.ts` is the one definition now
  (`NOTICE_ATTENTION` for style objects, `NOTICE_ATTENTION_CSS` for the two web parts whose styles
  are a CSS template literal), taken verbatim from `FolderAdmin.tsx`'s `danger` style — the banner
  the client actually pointed at. **This is the SECOND colour change they have asked for** (the
  replace-clash popup's icon went amber to red days earlier), which is why it is worth a module.
- **⚠ BADGES, CHIPS AND PER-ROW TINTS ARE DELIBERATELY EXCLUDED, and must not be pointed at it.**
  On My Submissions an amber `Pending` sits beside a red `Rejected`, a slate `Archived` and a
  pale-red `Failed` — those four colours ARE the information, and `Rejected` (an approver said no)
  must not look like `Failed` (the action broke). Also excluded: `.dms-fp-row.skipped` /
  `.dms-result.tagFailed` (per-row state in the results list), `.dms-admin-badge`, `chipPending`,
  `bPending`, `hcTag`, `badgeUsed`, and **`shared/toast.tsx`** — a toast is transient and already has
  an `error` variant, so reddening every warning toast would leave nothing to separate "this failed"
  from "this worked, with a caveat". A banner is a full-width box explaining a state; a badge is a
  word beside a filename.
- **⚠ THE COST, STATED RATHER THAN DISCOVERED: amber meant DOUBT and red meant FAILURE, and that is
  now gone from banners.** Several of these say *could not be checked* / *could not be read*, where
  this codebase treats an unanswerable read as different from a refused one — `unknown` is the amber
  one on purpose, in a dozen places. The WORDING still distinguishes them and every one of those
  banners already names which it is; only the colour stopped carrying it. Accepted on the client's
  instruction — re-raise it if they ever ask why a "not checked" notice looks alarming.
- **THE BULK UPLOAD COPY IS THE CLIENT'S OWN, VERBATIM.** ⚠ It is shorter than what it replaced, and
  the clause it drops is the one an admin might not infer: the old text said the files *"are visible
  to everyone with access to the destination folder as soon as they upload"*. That consequence is now
  unstated on screen. Their call, recorded here because it is a disclosure property rather than a
  wording preference.
  - ⚠ It still says **"the Documents library"** while the live library is titled
    **Restricted & Confidential Document** — as the previous copy did, so no regression, but it is a
    hardcoded name in a project that resolves every other title at runtime. `MySubmissions`'s footer
    interpolates `documentsTitle()` for exactly this sentence; offer the same here.
- **⚠ ONE GEOMETRY SLIP, CAUGHT AND REVERTED BEFORE SHIPPING.** Rebuilding `.dms-setup-warn` from the
  palette also changed its `border-radius`, `padding` and `margin`. The client asked for COLOUR;
  restoring the spacing was not optional. When replacing a whole CSS rule to change three
  declarations, diff the other declarations.
- **THE BANNER'S ICON IS THE CLIENT'S OWN SVG NOW (1.0.392.0)**, replacing the `⚠` glyph — supplied
  as `Icon.svg` at the project root.
  - **⚠ INLINED AS JSX, NEVER REFERENCED AS A FILE.** Shipping an image asset in this package has
    failed to provision on this tenant **twice, both times silently** (the `+ New Folder` customizer,
    then the bulk-approve command set), so a `src`-referenced icon would render as a broken box with
    nothing explaining it. Same reason the bulk-approve command set carries a base64 data URI. **The
    root `Icon.svg` is therefore the design source and is NOT shipped** — nothing reads it, like
    `Form.reference.tsx` beside it.
  - `flexShrink: 0` because `.dms-warn` is a flex row and the icon is otherwise squeezed as the
    sentence grows. `fill` stays the client's **`#DA3B3B`**, deliberately a shade off the banner's
    `#a4262c` text — switching it to `currentColor` would have quietly redesigned their asset.
  - **It was the ONLY rendered banner icon in the app.** The other `⚠` occurrences are glyphs inside
    reconciliation LOG LINES (`FolderManager.tsx` ~3538, ~6279), which are text in a run report and
    not banners — left alone.
- **Verified**: `tsc --noEmit` clean, `eslint` silent on all ten converted files (`BulkUpload.tsx`
  keeps only its pre-existing `max-lines`), full suite **1628/0**, 35 warnings — the full-lint
  baseline. Packaged as `1.0.392.0` via `npm run build`; the shipped bundles were grepped and **ten
  of them carry `#fdf3f4`**, the new copy and `#DA3B3B` are both present in `bulk-upload-web-part`,
  and the old `⚠` glyph is gone from it. A grep for the four old ambers leaves only badges, per-row
  tints, the toast module and the unmounted `GroupMapBuilder`. **NOT yet site-tested.**

## MY SUBMISSIONS LOSES THE SUMMARY PARAGRAPH AND GAINS AN `Archive` TAB (2026-09-03, 1.0.393.0)
Client, reading the page with real history on it: *"Remove these text in My Submissions"* (quoting the
whole deleted / replaced / archived paragraph), *"Remove 'in the archive' and 'replaced' and 'no longer
here'"*, and *"Add another tab call Archive so they can tell filter to archive files."*
- **THE PARAGRAPH AND THE THREE ACTION-CELL WORDS WERE THE SAME FACT SAID TWICE.** Every row already
  carries its state as a badge — `Deleted`, `Cancelled`, `Archived` — so the summary restated at the
  top of the page what the list says line by line (four sentences before the table on this site), and
  the action cell repeated it a second time in the next column along.
- **⚠ THE `unknown` SENTENCE IS DELIBERATELY KEPT**, and it is the one of the four that is not
  redundant: it means a LIBRARY READ FAILED, so rows may be missing from the page entirely and the
  ones badged `Not checked` may be perfectly fine. The other three describe rows that ARE listed;
  this one describes rows that are not. Dropping it would leave a page quietly incomplete with
  nothing saying so.
- **⚠ THE EMPTY BRANCH IN THE ACTION CELL IS LOAD-BEARING — the words went, the `r.recordState ?`
  test did NOT.** A gone row must get no Delete and no Share button, and for an archived file that is
  not tidiness: `validateDraft` REFUSES a deletion or share request against one outright
  (2026-08-22), because every role holds only Read on the archive and an approved request would fail
  in the approver's own session days later. Rendering `undefined` keeps the guard and drops the text.
- **⚠ THE `Archive` TAB CANNOT GO THROUGH `filterByTab`.** That matches `r.status === tab` —
  Pending / Approved / Rejected — and `archived` is a RECORD state, not an approval outcome; an
  archived row carries `status: "Pending"` only because the type demands a value. So a status-based
  Archive tab would match nothing. `archivedRowsOnly` in `shared/submissionRecords.ts` (pure, 3
  tests) beside `liveRowsOnly`, and the component handles it next to the existing `All` exception
  rather than adding a fourth string comparison to a function about approval outcomes. The two live
  in different modules because `Submission` knows nothing about records and importing `MergedRow`
  into `mySubmissions.ts` would be circular.
- **The tab count comes from the SAME helper the tab renders**, so the tally and the list cannot
  disagree — the defect that once read `Submissions (0)` over a table of fourteen.
- **⚠ ITS EMPTY STATE NEEDED ITS OWN WORDING, TWICE OVER.** The generic message is built from the tab
  name and would have read *"Nothing is archive right now"*; and an empty Archive is the NORMAL case
  for years — nothing on either site is close to seven years old — so it says *"No files of yours have
  been archived yet. Documents move here seven years after they were filed"* rather than implying
  something is missing.
- **⚠ ONE SELF-INFLICTED BREAKAGE WORTH RECORDING: splicing out the paragraph cut at the WRONG closing
  brace** and left an orphaned `</p>)}`, which `tsc` reported as *"Expected corresponding JSX closing
  tag for 'section'"* 200 lines away. When removing a nested JSX block by index, anchor on the OUTER
  closing sequence, not the first `)}` after the inner condition — and read the region back before
  building.
- **Verified**: `tsc --noEmit` clean, `eslint` on the file down to its pre-existing `max-lines` alone,
  full suite **1631/0** (3 new), 35 warnings — the baseline. Packaged as `1.0.393.0` via
  `npm run build`; the shipped `my-submissions-web-part_*.js` was grepped and carries
  `"Rejected","Archive","Requests"` with **zero** occurrences of `no longer here`, `in the archive`
  or the archived summary sentence. **NOT yet site-tested.**

## THE SHARE DIALOG IS A CHIP LIST NOW, AND THE EXTERNAL CHECK STOPPED LYING (2026-09-03, 1.0.394.0)
Five client items in one pass. The last one was a real defect and the third reverses a documented
decision, so both are recorded in full.
- **Title Case on the request chips** — `Share Requested` / `Deletion Requested`.
- **The action cell is EMPTY while a request is pending**, where it read `requested`. The chip beside
  the file name already says it, so that column was the same fact a third time. ⚠ **The BRANCH stays
  and is what matters:** while a request is open, Delete and Share must not be offered, or the
  uploader raises a second row for a decision the approver has not made — which is what the
  2026-08-20 "you already asked" work exists to prevent.
- **⚠⚠ THE FREE-TEXT RECIPIENT BOX IS NOW A READ-ONLY CHIP LIST, REVERSING A DELIBERATE DESIGN.**
  Client: *"Ensure client cannot type directly into the Email Address text box, but rather let them
  type one by one in the share with"*, plus a scroll cap and an X per address. `shareWith` is still
  the single source of truth — the chips and the external check both read one `shareRecipients`
  parse, so what is displayed and what is validated cannot diverge.
  - **WHY IT IS NOW SAFE, since the box existed for a stated reason:** it was kept because *"the
    point of a share request is usually somebody who is NOT on the site yet"* and
    `searchTenantPeople` will not always surface such an address, so typing was the only route to it.
    **That case is closed by POLICY** — adding outsiders was withdrawn on 2026-08-27 — and where
    external sharing is off the gate refuses those addresses anyway. The box had become the one way
    to type an address the form would then refuse.
  - **⚠ IF THE CLIENT EVER RE-ALLOWS OUTSIDE SHARING, THIS IS THE FIRST CONTROL TO REVISIT.** The
    picker alone may not reach every address they would then be permitted to share with.
  - Scrolls at **140px**: a long list would otherwise push Submit off a dialog that has no scroll of
    its own. Safe because nothing inside is absolutely positioned — the people picker sits ABOVE the
    box, not in it — unlike the three screens where a scroll cap has clipped a popover.
- **⚠⚠ THE EXTERNAL WARNING NAMED A COLLEAGUE ON THE VIEWER'S OWN TENANT AS AN OUTSIDER.** Client:
  *"this ... Is showing when I add Taba who is under the same tenant as me."*
  - **THE CAUSE IS THAT ONE SENTENCE SERVED TWO STATES.** `isExternal` fails **CLOSED** by design —
    *no domains supplied means UNKNOWN, and unknown counts as external* — so a site with **no
    `tenantDomains` row** refuses EVERY address while asserting each one is "outside the
    organisation". `MySubmissions` pushes the viewer's own domain in only when the config read
    yielded nothing, so a row that exists and lists other domains, or a `user.email` that is blank
    (a guest can have none), both land in the same place.
  - **THE CHECK IS UNCHANGED — only the reason it gives.** Guessing INTERNAL on an unconfigured site
    is how a document leaves the organisation on the strength of nothing. Both the dialog banner and
    `validateDraft`'s refusal now distinguish *"this site has not been told which domains belong to
    your organisation"* from *"this address is outside <named domains>"*, and the second **names the
    domains it compared against** so the answer is checkable at a glance.
  - **⚠ A BLANK ENTRY COUNTS AS NO DOMAIN.** A `tenantDomains` row of `"  "` parses to one empty
    string; treating that as configured would refuse everybody while claiming the domains were known.
    Filtered in `validateDraft` and pinned by test.
  - **⚠ THE FIX FOR THEIR SITE IS STILL A CONFIG ROW**, and the banner now says so: set
    `tenantDomains` on CRS Config. For SDG that is `sdguthrie.com` — already recorded as needed, and
    now the page will state it rather than mislabelling a colleague.
- **⚠ THIS MATTERS MORE NOW THAT THE BOX IS READ-ONLY.** With typing removed, a wrongly-classified
  internal address has **no route in at all**: the picker adds it, the gate refuses it, and there is
  no longer a text box to work around it. The classification and the recipient control changed in the
  same build for that reason.
- **Verified**: `tsc --noEmit` clean, `eslint` on both files down to `MySubmissions.tsx`'s
  pre-existing `max-lines`, full suite **1634/0** (3 new on `validateDraft`). Packaged as `1.0.394.0`
  via `npm run build`; the shipped `my-submissions-web-part_*.js` carries `Share Requested`,
  `Deletion Requested`, the chip list's empty state and the unconfigured-domains sentence, with the
  old `shareWith` textarea handler gone. **NOT yet site-tested.**

## ⚠⚠ THE AUDIT LOG COULD NOT FILTER TWO EVENT TYPES ITS OWN FLOWS HAVE BEEN WRITING (2026-09-03, 1.0.395.0)
Client: *"I notice the Audit log is missing some filter such as the File Replace filter."* Correct —
and it is a gap no test in this repo could have caught.
- **`Replaced` AND `ShareRevoked` WERE NEVER IN `EVENT`.** `Auto-route` and `HC Auto Route` have been
  writing `Replaced` rows since 2026-09-02, and `CRS — Audit request activity` writes `ShareRevoked`
  on a revoke. Both were stored perfectly — **`EventType` is a TEXT column, deliberately**, so a
  value the code has never heard of can never fail a write — and the viewer's Action dropdown is
  built from `ALL_EVENT_TYPES`, so the events sat in the log **unfilterable**.
- **⚠ THE STRUCTURAL POINT: A FLOW CAN INTRODUCE AN EVENT TYPE WITHOUT TOUCHING THIS CODEBASE, AND
  NOTHING FAILS WHEN IT DOES.** The existing test proves `EVENT` and `ALL_EVENT_TYPES` agree with each
  other; neither knows what Power Automate writes. **Whenever a flow gains a `Create item` with a new
  `EventType`, add it to `EVENT`, `EVENT_LABEL` and `ALL_EVENT_TYPES` in the same breath** — the row
  is written either way and only the FILTER is lost, which is why it goes unnoticed. The two are now
  pinned by name, with that reasoning at the test.
- **`Replaced` reads "Replaced by a newer upload"**, not the bare word: the row records that THIS
  document was superseded, and "Replaced" alone reads as though it did the replacing.
- **✅ CHECKED AGAINST THE LIVE LIST THE SAME DAY, AND THERE IS NO THIRD GAP.** With `Replaced` and
  `ShareRevoked` registered, a query for every row whose `EventType` is NOT one of the 26 known
  values returned an **empty feed** on ClarenceDMSTesting. So the log holds nothing unfilterable —
  established from the data rather than from reading the code.
- **⚠ THE QUERY THAT WORKS IS `ne` CHAINED WITH `and`, NOT `not(...)`.** The obvious form —
  `$filter=not(EventType eq 'A' or EventType eq 'B' or ...)` — answers
  **`-2147024809, System.ArgumentException — Value does not fall within the expected range`**, which
  is SharePoint's generic "I cannot parse this filter" and names nothing. `EventType ne 'A' and
  EventType ne 'B' and ...` over the same 26 values parses fine. Keep the exclusion in ONE filter:
  splitting it into shorter queries does not isolate unknown types, because each subset returns
  everything outside itself.
  - The no-filter fallback, if a longer list ever exceeds what the list will parse, is
    `$select=EventType&$orderby=EventType&$top=5000` — sorted, so identical values are adjacent and
    only the CHANGES have to be read.
- **⚠ RE-RUN IT WHENEVER AN AUDIT FLOW IS EDITED.** It is the only check that catches a flow
  inventing a type: this gap was found by a client reading a dropdown, which is not a control.

## BULK UPLOAD TITLES ITSELF, LIKE THE UPLOAD FORM (2026-09-03, 1.0.395.0)
Client: *"Same like a normal upload form, add the title back into the form for bulk upload."*
`<h1 className="dms-page-title">Bulk Upload</h1>`, with the CSS rule copied from `Form.tsx` verbatim
so the two upload pages carry the same heading — restyle one and restyle both.
- **⚠ THE PAGE'S OWN TITLE WEB PART MUST BE DELETED, or "Bulk Upload" renders TWICE** — the same
  required deployment step the upload form needed, on `Bulk-Upload.aspx` here and on SDG. **Nothing
  in code can detect it:** a text web part is not readable from inside a web part, so this will not
  fail, warn, or look wrong to any check.
- ⚠ `BulkUpload.tsx` is **4292 lines** and carries the client's own edits from this session (the
  banner icon at 35x35 and `.dms-warn { max-width: 626px }`). Both were preserved — every edit made
  to that file here is a string replacement with an assertion, so a changed string fails the edit
  loudly instead of overwriting their work.
- **Verified**: `tsc --noEmit` clean, `eslint` down to the pre-existing `max-lines` alone, full suite
  **1637/0** (3 new). Packaged as `1.0.395.0` via `npm run build`; the shipped bundles carry
  `dms-page-title` in `bulk-upload-web-part` and both new labels in `audit-log-web-part`.
  **NOT yet site-tested.**

## THE CLASH DIALOG IS YES / NO NOW, AND THE REQUESTS PAGE IS THREE ACCORDIONS (2026-09-04, 1.0.396.0)
A client batch of six. One of them was ambiguous in a way that destroys documents, and questioning it
was the right call.
- **⚠⚠ THE REPLACEMENT POPUP: THE MOCKUP CONTRADICTED ITSELF, AND EITHER READING SHIPS A REAL BUG.**
  The new copy says *"proceed to overwrite this file"* while the arrows pointed **Yes** at the
  **rename** button and marked the two REPLACE buttons *Remove*. So Yes meant overwrite under one
  reading and rename under the other. Asked rather than guessed; client: *"Option 1 is the way to go,
  you were right to question this ... Yes is basically Send for approval as a replacement."*
  - **ONE Yes REPLACES EACH ROW BY WHICHEVER MECHANISM THAT ROW NEEDS**, and the two are not the same
    act: an `approved` row goes to the approver like any other upload and **destroys nothing now**
    (Auto-route replaces the filed copy later, only if approved), while `staging`/`hidden`/`both` is
    the ONLY path that sets `overwrite=true` and destroys a pending draft immediately.
    `handleUpload` already took both id sets, so it is one call.
  - **⚠ THE `hidden` WARNING IS NO LONGER ON SCREEN ANYWHERE.** `clashRowLine` carried *"It may be
    someone else's, and replacing it would discard their file"* — a `hidden` clash is a COLLEAGUE'S
    draft that Draft Item Security hides, and Yes now destroys it without saying whose it is. The
    function is deleted with its per-row line; the sentence is preserved as a comment where it was
    removed and the Yes button restates what each case does. **Raised with the client — a known trade,
    not a silent one.**
  - The ` - Copy` suggestion is no longer OFFERED, per Option 1. `nextAvailableName` still runs (Bulk
    Upload uses it, and `reserved` needs it so two files in one run are never offered the same name);
    it simply has no button here.
  - **⚠ BULK UPLOAD'S CLASH DIALOG IS DELIBERATELY UNCHANGED.** Its `overwrite` is hardcoded `false`
    since 1.0.271.0, so it has no replace path at all — a Yes/No overwrite prompt there would offer
    something it cannot do. Rename-or-cancel is genuinely its only pair of options.
- **THE REQUESTS PAGE: THREE ACCORDIONS** — Deletion, Share, Documents with Shared Access.
  - **⚠ ALL THREE OPEN BY DEFAULT, and that is a safety rule rather than a style choice.** A collapsed
    section can HIDE PENDING WORK — the same reason the status filter may never hide a Pending row.
    Closed-by-default would let an approver load the page, see three tidy headers and miss four
    requests waiting on them. The pending count stays in the header while a section is shut.
  - **⚠ THE RE-CHECK BUTTON MOVED OUT OF THE SHARED-FILES HEADER**: that header is now a `<button>`,
    and a button inside a button is invalid HTML — the inner one swallows the click or the outer fires
    with it.
  - **Cards are white with a shadow and spaced apart** (client: *"it is blending in in the background
    making it difficult for them to see"* — they were `#fafafa` on a white card). White-on-white works
    only WITH the shadow; if the shadow ever goes, the background has to come back.
  - **A 4px LEFT BORDER PER STATUS**, from `STATUS_ACCENT` — the strong half of each existing pill, so
    a card's border and its badge cannot disagree. `Pending` uses its pill's TEXT colour (`#8a4b00`),
    because that pill's pale amber background is invisible as a rule. A `Record` over the union, so a
    new `RequestStatus` is a compile error until given a colour.
  - **The status badge moved to the LEFT of the file name** — it was on the far right, so on a wide
    screen the status and the filename it describes were the width of the page apart.
  - **Access Audit link removed**, and `auditLink`, its Site Pages effect and two now-unused imports
    went with it rather than being parked: an unused local is a lint warning and this project tracks a
    zero-new-warning baseline. Re-adding it is the `resolveLink` + `readSitePages` pattern five other
    files use.
- **The delete-request dialog takes the client's two-line copy**, ⚠ **keeping 93 days, not the 90 in
  their mock** — the same conflict resolved on 2026-08-30 when they were asked directly: *"stay 93,
  they might not know that is why they say 90."* The dropped clause (*"so only you and your approver
  can see it"*) is still TRUE, just no longer stated.
- **The dropdown arrow sits further from the border on BOTH upload forms** — `padding-right: 18px` on
  the select. Chromium draws a native select's arrow inside the padding box, so padding is what moves
  it; positioning the arrow directly means replacing it with a background image, which then needs its
  own disabled and dark-mode states.
- **⚠⚠ FIFTH INSTANCE OF THE BACKTICK TRAP, AND THE SECOND IN ONE DAY BY THE SAME HAND.** Writing that
  very rule, the comment said `` `padding-right` `` — a backtick inside the `<style>` template literal
  ENDS it, and `tsc` reports a JSX error naming neither the cause nor the line. A warning now sits on
  the rule itself, in both files.
- **✅ "THE CLASH DIALOG DOES NOT FIRE FOR A SYSTEM ADMIN" — REPORTED, THEN NOT REPRODUCIBLE.**
  Checked the same day with the Network panel open, as an admin, on a name already filed: the dialog
  fired correctly and the two probes answered
  `Files('<name>')?$select=Exists` → **404** (not in the approval library, because it had already
  routed) and both folder listings → **200**. So it took the `approved` branch exactly as designed.
  Client: *"oh huh it shows now... I think its probably me being tired, lets forget about this."*
  - **WHAT THIS RULES OUT if it ever returns:** the probes work for an admin, so the fail-open path in
    `approvedClashInfo` was NOT the cause, and `autoApproveOwnUpload` being `yes` does not suppress
    the dialog — either branch (still in staging, or already routed) produces it.
  - **The reasoning that made it worth checking rather than guessing still stands:** nothing in that
    path is permission-dependent in a way that would skip it, because an admin sees MORE, not less.
    Start again from the Network panel, not from a theory.
- **Verified**: `tsc --noEmit` clean, `eslint` with zero new warnings on all five touched files, full
  suite **1637/0**, 35 warnings — the baseline. Packaged as `1.0.396.0` via `npm run build`; the
  shipped bundles were grepped and carry `Replace Existing File` with **zero** occurrences of
  `Upload with the new name`, the accordion, `padding-right: 18px` on both upload forms, the new
  delete copy, and **zero** occurrences of `Access Audit`. **NOT yet site-tested.**
