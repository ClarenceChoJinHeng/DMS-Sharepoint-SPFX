# Configurable Folder Structure Chain

**Date:** 2026-08-06
**Revised:** 2026-08-09 — re-verified against the code. One central claim was wrong; see
"The reconciliation hazard". Line references replaced with symbol names, which do not drift.
**Status:** Piece 1 IMPLEMENTED 2026-08-09 — built and unit tested, **not yet tested on a live site**
**Scope:** Piece 1 of 3 — the data model. Not a client-facing deliverable on its own.

> ⚠ **Piece 1 alone gives the client nothing.** The chain lives in `Levels` JSON on a `DMS Config`
> mode row, and the client cannot edit JSON — that is the whole reason this feature exists. Piece 1
> is testable by hand-authoring a row; it is not handover-ready until piece 2 (Structure Manager UI)
> exists. Do not demo it as a finished capability.

---

## Problem

Adding a business segment, or changing the folder shape below Unit, currently requires a developer.
Two separate causes:

1. **Segment onboarding** needs hand-authored `Levels` JSON on a `DMS Config` mode row.
2. **Below-Unit structure is hardcoded.** [Form.tsx](../../../src/webparts/form/components/Form.tsx)
   resolves the Unit folder by UniqueId and then makes exactly two `ensureFolder` calls, `Year` then
   `Document Type`, from `yearPeriod` / `documentType` — each with its own React state, its own entry
   in the required-field check inside `handleUpload`, and its own JSX select beneath the chain
   `.map`. [BulkUpload.tsx](../../../src/webparts/bulkUpload/components/BulkUpload.tsx) — note
   `src/webparts/bulkUpload/`, not under `form/` — holds a near-duplicate copy, including its own
   `LEVEL_COLUMNS`, its own `LevelSelection` type and its own two `ensureFolder` calls.

The client's actual ask, confirmed 2026-08-06: add a new folder tier below Unit — e.g. `Function`
containing `Human Resource` — and place it **anywhere** in the path (before Year, between Year and
Document Type, or after Document Type), then upload into it.

The client cannot edit JSON. This spec defines the model that a later admin UI writes to; the
client never sees the JSON at any point.

## Non-goals

- **The admin UI.** Separate spec (piece 2). Required before client handover — this piece alone
  never ships to them.
- **Migrating existing files** when a tier is inserted above existing content. Separate spec
  (piece 3). Decided 2026-08-06: legacy subtrees move under one admin-chosen value per unit, as a
  folder-subtree move, not a file-by-file move. Findings gathered so far are recorded below under
  "Piece 3 — findings and outstanding tests", so the eventual spec starts from evidence rather
  than assumption.
- **Ordering of non-tier fields** (Document Date, Confidentiality, Remark, filename parts).
- **Breaking inheritance below Unit.** Explicitly declined by the client 2026-08-06 — every folder
  below Unit inherits the Unit's ACL, including new tiers.

---

## Model

One ordered chain per mode, stored in the existing `Levels` column on a `DMS Config` `mode` row.
Below-Unit tiers are entries carrying `permissioned: false`.

```json
[
  {"label":"Department","column":"Department","labelCol":"Department","tidCol":"DepartmentTid"},
  {"label":"Unit","column":"Unit","labelCol":"Unit","tidCol":"UnitTid"},
  {"label":"Function","column":"Function","labelCol":"Function","tidCol":"FunctionTid",
   "termSet":"00000000-0000-0000-0000-000000000000","permissioned":false},
  {"label":"Year","column":"Year","labelCol":"Year",
   "termSet":"023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf","permissioned":false},
  {"label":"Document Type","column":"Document_x0020_Type","labelCol":"Document_x0020_Type",
   "termSet":"866c5754-258e-401f-8685-03d20ae59b1d","permissioned":false}
]
```

Reordering a tier is moving an entry. Inserting one is adding an entry. The form's existing `.map`
over `activeMode()?.levels` renders every entry; the upload path is built by walking the chain in
order.

> ⚠ **`parseLevels` must be extended before any of this works.** It lives in
> [`src/shared/formModel.ts`](../../../src/shared/formModel.ts) and copies exactly `label`, `column`,
> `labelCol`, `tidCol` — every other key is dropped without comment. `termSet` and `permissioned`
> authored on a live row today would be discarded at parse, leaving an all-permissioned chain with no
> term sets, which the compatibility bridge below would then mask by falling back to today's
> behaviour. The feature would appear to be built and would do nothing. Extend the `Level` interface
> and the parser first.

Today's `Level` interface, for reference — `termSet` and `permissioned` are the two additions, and
`termSet` currently lives on the **mode** (`ModeV2.termSetGuid`), never on a level:

```ts
export interface Level {
  label: string;
  column: string;
  labelCol?: string;
  tidCol?: string;
}
```

This supersedes the fixed `{Year}/{Document Type}` model specified in
[2026-07-15-multi-segment-form-flow-design.md](2026-07-15-multi-segment-form-flow-design.md), lines
111–122. That spec's statement that Unit is "the deepest permissioned folder" remains true and
becomes the `permissioned` flag.

### Field semantics

| Field | Meaning |
|---|---|
| `label` | Dropdown caption, and the name used in "please choose X" validation messages |
| `column` | Logical key. Resolves the `LEVEL_COLUMNS` fallback pair when `labelCol`/`tidCol` are absent |
| `labelCol` | Internal name of the column receiving the term's label |
| `tidCol` | Internal name of the column receiving the term's GUID. Optional — omit for tiers with no GUID column |
| `termSet` | Term-set GUID. **Presence decides option loading** (below) |
| `permissioned` | `true` → reconciliation builds it and it gets an ACL. `false` → ensure-created on demand, inherits parent. **Absent means `true`** |

### Two derived behaviours, no extra flags

- **Entry has `termSet`** → options are that set's top terms, loaded independently of the preceding
  tier. Flat. (Function, Year, Document Type.)
- **Entry has no `termSet`** → options cascade from the previously selected term inside the mode's
  own segment term set. (Department, Unit — today's behaviour, unchanged.)

Derivable from the data, so no `source` discriminator is added: a tier bound to its own term set is
by definition not a cascade of the segment tree.

### Folder naming — unchanged, and deliberately split

- **Permissioned tiers** take their folder name from the **term abbreviation list**, resolved at
  runtime via `cachedListTitle` / `LIST_SUFFIX` in `src/shared/naming.ts`. The live list is
  `CRS Term Abbreviation` on renamed sites and `DMS Term Abbreviation` on others — never a literal.
  A term with no abbreviation is skipped, never guessed.
- **Non-permissioned tiers** use the sanitized term label (`2026`, `Invoice`, `Human Resource`) via
  `sanitizeFolderSegment`. No abbreviation lookup.

Rationale: abbreviations keep permissioned paths short and stable across renames, because those
paths carry ACLs and appear in the Folder Map. Below-Unit folders carry neither, and a full readable
label is more useful to someone browsing.

---

## The reconciliation hazard, and the guard

> ❌ **CORRECTED 2026-08-09.** The original text claimed reconciliation "walks `Levels` and, for
> every entry, creates a folder per term and breaks inheritance on it", and that unfiltered
> non-permissioned entries would mint thousands of ACL'd folders. **That is not what the code does**,
> and building the guard as specified would have protected against nothing while leaving the real
> hazard in place. Kept visible rather than deleted, because the mistaken model is the intuitive one.

**What reconciliation actually walks is the term store, not `Levels`.** `walk()` in
`FolderManager.tsx` recurses the *segment term set's own hierarchy* — a term's children are the next
folder tier, to whatever depth the term tree happens to have. `Levels` contributes only **display
names** for those tiers, via `loadReconLevelNames()` (with `FALLBACK_LEVEL_NAMES` behind it), used
when reporting a term that has no abbreviation.

So a `permissioned: false` entry bound to its **own** term set — Function, Year, Document Type — is
not part of the segment tree at all, and reconciliation would never create a folder for it however
the flag is set. The catastrophic outcome the original text feared cannot occur.

Three real hazards remain, and they are what the guard must actually address:

1. **Level names misalign with depth.** `loadReconLevelNames()` returns the labels in chain order and
   they are indexed by tree depth. Add `Function` to the chain and every below-Unit label shifts, so
   a missing-abbreviation report names the wrong tier. Cosmetic, but it is a report an admin acts on.
   **Guard: filter to `permissioned === true` before deriving names.**

2. **The Year × Document Type grid is hardcoded, and it is the one place that WOULD build a wrong
   shape.** When `recon_gridMode` is on, reconciliation pre-creates `Unit / Year / Document Type`
   from `gridSets.year` and `gridSets.docType` — two fixed loops, no reference to the chain. Insert
   `Function` and the form starts writing `Unit / Function / Year / Document Type` while
   reconciliation keeps pre-creating the old shape: **two trees per unit, both populated, neither
   complete.** `recon_gridMode` is currently **off**, which is the only reason this is latent.
   **Guard: the grid loop walks the non-permissioned suffix in chain order.** It must be fixed in the
   same change even though it is off, precisely because turning it on later would look unrelated.

3. **`parseLevels` silently discards unknown keys.** It copies `label`, `column`, `labelCol`,
   `tidCol` and nothing else, so `termSet` and `permissioned` authored today vanish on parse — the
   chain would read as all-permissioned with no term sets, and the compatibility bridge below would
   quietly hide it by falling back to `Year → Document Type`. **`parseLevels` must be extended
   first**, or every other change in this spec appears to work and does nothing.

**A missing `permissioned` key is still treated as `true`**, for the reason that survives the
correction:

1. **Backwards compatibility.** The three live mode rows (`mode_gho`, `mode_minamas_ho`,
   `mode_nbpol_ho`) carry `[Department, Unit]` with no flag, and both tiers are permissioned. They
   behave identically with zero edits.
2. **Fail-safe direction.** A forgotten flag leaves a tier in the permissioned prefix, where it is
   visible and noisy. The opposite default would move a tier that needs an ACL into the on-demand
   suffix, which inherits — a silent permissions widening. Same principle as gotchas #9 and #12 in
   CLAUDE.md: prefer the loud failure.

### The permissioned prefix must be contiguous

Every position the client asked for — before Year, between Year and Document Type, after Document
Type — is **below** Unit, so every new tier is non-permissioned. A permissioned entry appearing
*after* a non-permissioned one would mean an ACL'd folder inside an inheriting one, which
reconciliation cannot build (it never reaches there) and which the client explicitly declined.

`folderChain.ts` therefore **rejects such a chain outright** rather than sorting it into shape.
Silently reordering would move a tier the author meant to be permissioned into the inheriting
suffix — the exact silent widening the default above exists to avoid.

---

## Compatibility bridge

If a mode's chain contains **no** non-permissioned entries, the upload web parts fall back to
today's hardcoded `Year → Document Type` behaviour.

This lets a site migrate one mode row at a time instead of all-or-nothing, and means a partially
configured site keeps uploading correctly rather than routing files to a truncated path. The
fallback is removed once every site is migrated.

---

## Components

### `src/shared/formModel.ts` (extend first)

`Level` gains `termSet?: string` and `permissioned?: boolean`, and `parseLevels` copies them. Nothing
else in this spec functions until this lands — see the warning under "Model". `parseLevels` keeps
dropping entries without a string `label` + `column`; that behaviour is unchanged and still correct.

`LEVEL_COLUMNS` is currently **duplicated** — one copy in `Form.tsx`, an identical one in
`BulkUpload.tsx`. Move it here in the same change. Two copies of a table that maps logical tier keys
to real SharePoint internal names is exactly the shape of bug this spec exists to remove.

### `src/shared/folderChain.ts` (new)

Both upload web parts hold near-duplicate copies of the level logic today. The chain walker goes in
`shared/` and both call it. Pure and SPFx-free where possible, so the ordering rules are unit
testable without a tenant — same approach as `approvalQueue.ts`.

Responsibilities:

- Parse and validate a chain (extends the existing `parseLevels`).
- Split it into permissioned prefix and non-permissioned suffix.
- Given selected terms, produce the ordered list of folder segment names for the suffix.
- Report a malformed chain loudly rather than routing files to a partial path.

### `Form.tsx` / `BulkUpload.tsx`

- Replace `yearPeriod` / `documentType` dedicated state with chain-indexed selections.
- Derive required-field validation from the chain, not hardcoded strings.
- Replace the fixed two-step `ensureFolder` chain with a walk over the non-permissioned suffix.
- Metadata writes: each tier writes its `labelCol` / `tidCol` exactly as levels do now, via
  `buildLevelFormValues`.

### `FolderManager.tsx`

- Filter to `permissioned === true` before deriving tier **names** in `loadReconLevelNames()`, so
  missing-abbreviation reports name the right tier (hazard 1).
- **Rewrite the Year × Document Type grid loop to walk the non-permissioned suffix in chain order**
  (hazard 2). Two hardcoded loops become one walk. Do this even though `recon_gridMode` is off.
- No change to `walk()` or to the create/break-inheritance pass — those follow the term tree and were
  never driven by `Levels`.

---

## What changed between 2026-08-06 and 2026-08-09

Facts established after this spec was written that bear on it. None invalidate the design; two remove
work it would otherwise have needed.

**Below-Unit folders are already created, approved and verified in production shape.** The upload
form ensure-creates them **as the uploader**, so in the moderated approval library they arrive
Pending and are invisible to peer PICs. `CRS — Approve new folders in Approval Document` approves
them, guarded by `@equals(triggerOutputs()?['body/{IsFolder}'], true)`. A new tier is just another
on-demand folder on that path, so **no flow change is required** — the guard is on "is it a folder",
not on which tier it is. See
[2026-08-08-auto-route-flow-and-draft-isolation.md](2026-08-08-auto-route-flow-and-draft-isolation.md).

**The library is `Approval Document` at `/ApprovalDocument`.** Title and URL differ. The Auto-route
flow splits destination paths on `ApprovalDocument/`. Nothing in the chain walker may hardcode
either half — resolve via `libraryTitle()` / `libraryUrlSegment()` (CLAUDE.md gotcha #12).

**Every new tier makes every path one level deeper**, and depth has bitten this codebase before.
CLAUDE.md gotcha #9: an inline quoted path in `GetFolderByServerRelativeUrl` returns **HTTP 400** —
not 404 — at roughly 330 characters and 6 nesting levels, well under SharePoint's 400-character
item-path limit. The current below-Unit path already uses the parameter-alias form throughout
(`ensureFolder` → `AddUsingPath(DecodedUrl=@u)`, `resolveFolderByPath` → `…(@f)?@f='…'`), so this
path is safe **today** — but any new request added by the chain walker must use the alias form too,
and inserting a tier moves every site closer to SharePoint's own limit. Worth stating to the client
when they choose an insertion position, alongside the migration-cost table below.

---

---

## Decisions taken

- **All tiers are required.** An optional tier left blank produces files at inconsistent depths
  within the same unit, defeating the purpose of the tier. Not worth the complexity.
- **Layout: chain entries render three per row, in order.** The current code deliberately pins the
  deepest level onto one row with Year and Document Type; that rule has no meaning once the chain
  varies in length. Flagged to Clarence 2026-08-06; revisit if the client objects to the grouping
  change.
- **Below-Unit folders do not receive the folder content type or `Full Name`.** Unchanged from
  today. Those exist for permissioned folders, where the abbreviated name hides the real label.

## Open questions

None blocking. The layout grouping above is the only item that may need revisiting after the client
sees it.

---

## Error handling

| Condition | Behaviour |
|---|---|
| Chain fails to parse | Upload blocked, admin-facing message naming the mode row. Never a silent fallback to a partial path |
| Tier's term set returns no terms | Existing graceful empty-state, extended to any tier: name the missing tier and tell the user to ask an administrator ([Form.tsx:1668](../../../src/webparts/form/components/Form.tsx#L1668)) |
| `labelCol` names a nonexistent column | `validateUpdateListItem` returns HTTP 200 with `HasException` set — must be checked per gotcha #4, and surfaced rather than swallowed |
| Permissioned tier's term has no abbreviation | Existing behaviour: skipped, reported by reconciliation, no folder created |

## Testing

**Automated — done.** 30 tests in `folderChain.test.ts`, 306 across the suite, build clean with no
new lint warnings. Covers ordering, the permissioned split, `absent means true` (including the
string `"false"`, which a hand-authored row can easily contain), malformed-chain rejection,
`gridPlan` arithmetic at two and three tiers, and the regression that an unconfigured chain produces
exactly today's `Year → Document Type` path and column names.

**Manual — NOT YET RUN.** Nothing below has been exercised against a real site. Do this before
telling the client anything is available:

| Check | Why it is the one that matters |
|---|---|
| Upload with **no** chain configured | Proves the migration is invisible. Path, metadata and folder names must be byte-identical to before. Do this FIRST — it is the regression that affects every existing user |
| Insert a tier **before** Year | The cheap position, and the one the client will most likely want |
| Insert one **between** Year and Document Type, and **after** Document Type | Ordering is the whole feature; a chain that silently sorts itself would pass the first test and fail these |
| Run reconciliation with `recon_gridMode` **on** | The two-trees-per-unit bug. Confirm one tree, matching what the form writes |
| Run reconciliation with a chain configured and `recon_gridMode` **off** | Confirm nothing below Unit is created and nothing below Unit breaks inheritance |
| Author a deliberately malformed chain | Upload must be BLOCKED and name the config row — never routed to a partial path |
| Bulk upload with the same chain | Both web parts must agree on the shape; disagreement is the failure this feature exists to prevent |

**A tier added below Unit needs its term set to exist and its terms to be tagged**, and — unlike
permissioned tiers — needs **no** abbreviation row, no Folder Map row and no group.

---

## Piece 3 — findings and outstanding tests

Recorded 2026-08-07 during design discussion. **Piece 3 has no spec yet.** This section exists so
whoever writes it starts from what is known rather than re-deriving it.

### Established

**The client wants migration, not coexistence.** Confirmed 2026-08-06. When a tier is inserted
above existing content, the existing subtree moves under it, so a unit ends up with one shape, not
two side by side.

**Cost scales with insertion depth**, because a folder move carries its whole subtree:

| Insert position | What moves | Volume per unit |
|---|---|---|
| Above Year (top) | the Year subtrees | a few folder moves |
| Between Year and Document Type | the Document Type subtrees | years × a few |
| Below Document Type (bottom) | the **files** themselves | thousands |

Inserting at the top is cheap. Inserting at the bottom is the expensive case and should be
called out to the client before they choose a position.

**The legacy value is an admin decision, not a derivable one.** Nothing in the system knows whether
a 2026 invoice was Human Resource or Finance. The admin picks one target value per unit when adding
the tier, with a preview of what will move before anything happens.

### Auto-route flow — verified from the live flow, 2026-08-07

Two properties observed directly:

- **Trigger is `When an item is created or modified`.** This is correct and must not be changed —
  Power Automate offers no "when approval status changes" trigger for SharePoint, so
  created-or-modified plus a moderation-status condition is the only way to catch an approval.
- **`Copy file` is set to `If another file is already there: Replace`.** So a re-fire that resolves
  to the *same* destination path overwrites in place. No duplicate files from metadata edits.

Two consequences follow, and both matter to piece 3:

1. **Replace protects only an unchanged path.** The destination is a `concat(...)` built from the
   live folder path, which is what makes the flow rename-proof. A folder *move* changes that path,
   so a re-fire after migration copies to the NEW Documents path while the old copy remains at the
   old path — an orphan, not a replacement.

   **Therefore migration must operate on BOTH libraries as a single run, with the flow disabled for
   its duration.** Migrating `Approval Document` alone leaves `Documents` in the old shape, and the
   next approval in that unit silently starts building a second tree. This is a firm scope
   conclusion, not a suggestion.

2. **The approval email is probably not idempotent.** `Send an email from a shared mailbox` sits in
   the same True branch after `Copy file`. Replace makes the file idempotent; nothing makes the
   email idempotent. If the branch condition tests approval status alone, every metadata edit on an
   already-approved file re-sends the approval notification to the uploader. **Unverified** — the
   condition itself has not been read. If confirmed this is a live bug independent of this feature,
   not something migration introduces.

### Outstanding tests — blocked on site maintenance

To run on `/sites/ClarenceDMSTesting` against a throwaway unit folder, never real data. Set up
`TESTUNIT/2026/Invoice/` holding one approved file (let Auto-route copy it to Documents) and one
pending file, then create `TESTUNIT/Human Resource/` and **Move to** the `2026` folder under it.

| Check | Expectation | Result |
|---|---|---|
| Approval status on the approved file | Still Approved — same list item, path only changed | **CONFIRMED 2026-08-10** — moved `CORU/2024` under `CORU/TESTMOVE/`; file still `Approved` |
| Metadata columns (`Business Segment`, `Department`, `Unit`) | Unchanged — stored values, not derived from path | **CONFIRMED 2026-08-10** — intact after the move |
| Approval status on the pending file | Still Pending | not yet run |
| List item `ID` of both | Unchanged | not yet run |
| Version history | Preserved (Move keeps it; Copy would not) | not yet run |
| Documents library copy | Still at the OLD path — the move does not touch Documents | superseded — the flow re-fires, see below |
| Permissions on moved files | Unchanged — still inheriting from the unit folder | not yet run |
| Read the flow's branch condition | Confirms or clears the duplicate-email question above | not yet run |

**Approval status survives a folder move, so piece 3 is folder moves and nothing else.** No
re-approval pass, no per-file work, no admin process to re-stamp status. This was the single finding
that decided whether the migration was cheap or expensive, and it came back cheap — the
"new uploads only, two shapes side by side" fallback is no longer needed.

**The move re-fires Auto-route — observed in the same test.** A move *modifies* the item, the
trigger is `When an item is created or modified`, and the condition tests only approval status, so
the flow copies the file again to the NEW path in `Documents` while the old copy stays at the old
path. No longer a prediction: it is why a migration run **must disable the flow for its duration and
move both libraries in one pass**. Nothing about the move itself is at fault, and no flow change
would fix it — Power Automate offers no "when approval status changes" trigger for SharePoint.

---

## Related

- [2026-07-15-multi-segment-form-flow-design.md](2026-07-15-multi-segment-form-flow-design.md) — the
  fixed-structure model this supersedes
- [2026-07-24-twelve-segment-expansion-design.md](2026-07-24-twelve-segment-expansion-design.md) —
  the 12-segment map this unblocks
- [2026-07-30-folder-abbreviation-naming-design.md](2026-07-30-folder-abbreviation-naming-design.md)
  — abbreviation rules for permissioned tiers
- Piece 2 (Structure Manager UI) and piece 3 (subtree migration) — specs to follow
