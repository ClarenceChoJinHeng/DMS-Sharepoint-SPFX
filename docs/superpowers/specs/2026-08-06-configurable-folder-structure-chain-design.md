# Configurable Folder Structure Chain

**Date:** 2026-08-06
**Status:** Design agreed, not implemented
**Scope:** Piece 1 of 3 — the data model. Not a client-facing deliverable on its own.

---

## Problem

Adding a business segment, or changing the folder shape below Unit, currently requires a developer.
Two separate causes:

1. **Segment onboarding** needs hand-authored `Levels` JSON on a `DMS Config` mode row.
2. **Below-Unit structure is hardcoded.** `Form.tsx` ensure-creates exactly `Year` then
   `Document Type` ([Form.tsx:1120](../../../src/webparts/form/components/Form.tsx#L1120)), each with
   its own React state, its own required-field check, and its own JSX
   ([Form.tsx:1650](../../../src/webparts/form/components/Form.tsx#L1650),
   [1656](../../../src/webparts/form/components/Form.tsx#L1656)). `BulkUpload.tsx` holds a
   near-duplicate copy.

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
over the chain ([Form.tsx:1632](../../../src/webparts/form/components/Form.tsx#L1632)) renders every
entry; the upload path is built by walking the chain in order.

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

`FolderManager.tsx` reconciliation walks `Levels` and, for every entry, creates a folder per term
**and breaks inheritance on it**. If reconciliation sees the new non-permissioned entries
unfiltered, it will create a permissioned folder for every Year term and every Document Type term
under every unit — thousands of folders with unique ACLs. Unpicking that by hand is days of work,
and it directly contradicts the client's instruction that nothing below Unit breaks inheritance.

**Guard:** reconciliation filters the chain to `permissioned === true` before walking it, and **a
missing `permissioned` key is treated as `true`.**

Two reasons for that default:

1. **Backwards compatibility.** The three live mode rows (`mode_gho`, `mode_minamas_ho`,
   `mode_nbpol_ho`) carry `[Department, Unit]` with no flag, and both tiers are permissioned. They
   behave identically with zero edits.
2. **Fail-safe direction.** A forgotten flag over-builds permissioned folders — visible, noisy, and
   caught by the existing pre-flight checks. The opposite default would silently drop ACLs from
   folders that need them, a permissions failure nobody sees. Same principle as gotchas #9 and #12
   in CLAUDE.md: prefer the loud failure.

---

## Compatibility bridge

If a mode's chain contains **no** non-permissioned entries, the upload web parts fall back to
today's hardcoded `Year → Document Type` behaviour.

This lets a site migrate one mode row at a time instead of all-or-nothing, and means a partially
configured site keeps uploading correctly rather than routing files to a truncated path. The
fallback is removed once every site is migrated.

---

## Components

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

- Apply the `permissioned` filter before walking the chain, in both reconciliation and `RECON_MODES`.

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

- Unit tests on `folderChain.ts`: ordering, the permissioned/non-permissioned split, missing-flag
  defaulting to `true`, malformed-chain rejection.
- A regression test asserting a chain with no non-permissioned entries produces exactly today's
  `Year → Document Type` path.
- Manual: insert a tier before Year, between Year and Document Type, and after Document Type;
  confirm the path, the metadata columns, and that reconciliation creates nothing below Unit.

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
| Approval status on the approved file | Still Approved — same list item, path only changed | |
| Approval status on the pending file | Still Pending | |
| List item `ID` of both | Unchanged | |
| Metadata columns (`Unit`, `UnitTid`, `Year`, `Document_x0020_Type`) | Unchanged — stored values, not derived from path | |
| Version history | Preserved (Move keeps it; Copy would not) | |
| Documents library copy | Still at the OLD path — the move does not touch Documents | |
| Permissions on moved files | Unchanged — still inheriting from TESTUNIT | |
| Read the flow's branch condition | Confirms or clears the duplicate-email question above | |

**If approval status does NOT survive the move**, migration becomes materially more expensive —
every moved file needs re-approving or its status re-stamped by an admin process. That is the point
at which "new uploads only, two shapes side by side" should be reconsidered, despite the client
having rejected it.

---

## Related

- [2026-07-15-multi-segment-form-flow-design.md](2026-07-15-multi-segment-form-flow-design.md) — the
  fixed-structure model this supersedes
- [2026-07-24-twelve-segment-expansion-design.md](2026-07-24-twelve-segment-expansion-design.md) —
  the 12-segment map this unblocks
- [2026-07-30-folder-abbreviation-naming-design.md](2026-07-30-folder-abbreviation-naming-design.md)
  — abbreviation rules for permissioned tiers
- Piece 2 (Structure Manager UI) and piece 3 (subtree migration) — specs to follow
