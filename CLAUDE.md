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

> 🔴 **IN PROGRESS — READ FIRST: `docs/2026-08-18-bulk-groups-handoff.md`.** Bulk group provisioning is
> built and part-run on the dcistaging rehearsal site. Three of the five defects are **fixed in 1.0.153.0 and
> not yet site-tested**; **two defects remain**, listed there in build order. `CRS Group Map` has been
> cleared and **reconciliation must not be run** until the row count is right.

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
      - **Deleting the folders is a separate opt-in**, off by default, behind a typed confirmation
        of the segment's own label — added on the client's reasoning that once the files have been
        moved out with the migrator, removing the empty shell should be easy. It **recycles** rather
        than purges, so it is restorable for 93 days, and the dialog says so: that is what makes the
        option acceptable at all.
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
  anything); changing one **renames a live folder** on the next run. Spec
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
| Head of Unit | `APR`, `DEL`, `DELS`, `SHARE`, `UPLHC`, `DELSHC` | unit | view own path, delete + share | **approve + view every file in their unit, delete pending/rejected, upload incl. HC** |
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
  - **Two defects remain open — see `docs/2026-08-18-bulk-groups-handoff.md` before running it.**
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
- **Bulk Upload gets NO write probe**, deliberately — admin-only, and already existence-gated only
  because an `AddListItems` probe against `Documents` would empty the form for every PIC. It writes
  straight to the approved side, so HC there means **`HC Documents`**, path segment and metadata target.
- **Requests match `APR` *and* `APRHC`.** A unit whose approver is the HC one has no plain `APR` row,
  and matching `APR` alone leaves that queue permanently empty while requests pile up behind it.
- **TWO NEW POWER AUTOMATE FLOWS ARE REQUIRED and nothing works without them** — HC Auto-route
  (`{IsFolder}` **false**) and HC folder approval (**true**). The polarity is the whole thing, and both
  mistakes have already been made once on the normal pair. Build as the **service account**.
  `HC Documents` must have **content approval OFF**.
- **MIGRATION IS OUT OF SCOPE AND MUST NOT BE SILENTLY SKIPPED.** Documents already labelled Highly
  Confidential sit in the normal libraries, readable by their whole unit. After deploy the label
  implies a protection they do not have — worse than before. Until a sweep runs, HC protection applies
  only to documents filed **after** deployment.

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
