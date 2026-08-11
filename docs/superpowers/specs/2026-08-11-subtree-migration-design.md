# Subtree Migration — moving documents that are already filed into a new folder shape

**Date:** 2026-08-11
**Status:** built and **VERIFIED LIVE — all four cases** (add, reorder, remove-clean, remove-with-
collision) on ClarenceDMSTesting, 2026-08-11/12. `SubtreeMigrator.tsx` (tab 2 of the Folder Structure
web part) + `shared/subtreeMigration.ts` (67 unit tests).

> **The four runs, each matching its predicted counts exactly.** Baseline
> `Credit_Card / Year / Document Type` with three documents: A and B sharing a filename under
> different Credit_Card values, C distinct.
>
> | Test | Change | Result |
> |---|---|---|
> | 1 | ADD `Testing` between Credit_Card and Year | one picker; 3 moved, 6 tidied, 3 tagged |
> | 2 | REORDER `Year` above `Testing` | **no picker**; 3 moved, 8 tidied, **0 tagged** |
> | 3 | REMOVE `Testing` (no collision) | 3 moved, 6 tidied, 0 tagged |
> | 4 | REMOVE `Credit_Card` (collision) | 1 clash surfaced, renamed, 3 moved, 8 tidied — **count held at 3** |
> | 5 | ADD `Stage` **at the bottom**, below Document Type | 3 moved, **0 tidied**, 3 tagged |
>
> Test 2's **0 tagged** matters: a reorder changes no tier's value, so nothing should be re-stamped.
> Test 4's **count held at 3** matters more: both documents landed side by side under different
> names, neither overwritten. That is the only failure in this feature that loses a document rather
> than misplacing it.
>
> Test 5's **0 tidied** matters for the opposite reason to the others: nothing is vacated when
> documents move *down*, so the old leaves simply become parents. Any deletion there would be wrong.
> That test is also the one that found the depth-cap bug — four passing tests had all used
> full-depth folders, so the only untested position was the one that was broken.
>
> Also confirmed by these runs: **a file move preserves approval status** (every document stayed
> "Waiting for Approval" across five migrations), and the emptied-folder cleanup removed exactly the
> husks each change vacated and nothing else.

> **Verified on ClarenceDMSTesting, 2026-08-11 — reorder.** `Testing` and `Year` swapped on NBPOL
> while 3 documents were filed. `Credit2/testig/2024/Tax Return` → `Credit2/2024/testig/Tax Return`
> in both libraries: 3 documents moved, 6 empty folders tidied (3 per library), 3 re-tagged, and the
> staged structure activated only after the re-scan came back clean.
>
> Two facts confirmed by that run, both previously assumed:
> - **A FILE move preserves approval status.** Both approval-library documents were still "Waiting
>   for Approval" afterwards. Folder moves were verified on 2026-08-10; this is the file case, and
>   the whole file-by-file approach depends on it.
> - **Emptied-folder cleanup removes exactly the husks and nothing else** — the count matched the
>   three tiers per library that the reorder vacated.
**Piece 3** of the configurable folder chain. Piece 1 made the below-Unit chain data-driven;
piece 2 ([2026-08-10-structure-manager-ui-design.md](2026-08-10-structure-manager-ui-design.md))
gave the client an editor for it. Both apply to **new uploads only**, which is what this closes.

---

## 1. The problem this solves

Add a level below Unit to a segment that already holds documents and the library ends up with two
shapes side by side:

```
GHO / GF / CORU / 2024 / Tax Return / a.pdf       <- filed before the change
GHO / GF / CORU / testig / 2025 / Invoice / b.pdf <- filed after it
```

Nothing is broken and nothing is lost — old files stay exactly where they were, and the Structure
Manager warns before saving. But "all of CORU's Tax Returns" has stopped being one folder you open,
which is the thing the folder architecture exists to provide.

The client's instruction (2026-08-10) fixes the shape of the answer:

> the superadmin will choose which old year folder will move into which folder first … it can only
> be one folder to move into, no copying old folder into multiple new folders, only one

So: **an administrator chooses the destination, the tool performs the moves.** No inference about
where a document *should* live, no copying, no fan-out.

## 2. Why moving is cheap, and what makes it safe

Established by live test on 2026-08-10:

- **A folder move carries its whole subtree** — one move relocates every year, document type and
  file beneath it.
- **Approval status survives the move.** A file approved before the move is still approved after it.
  Nothing needs re-approving, and no uploader is asked to resubmit.
- **UniqueId survives.** `DMS Folder Map` keys on UniqueId and only ever holds permissioned folders
  (segment / department / unit) — none of which move here — so the Folder Map needs no repair. State
  this explicitly because the opposite assumption would justify a large amount of pointless work.
- **Below-Unit folders carry no ACL.** They inherit the Unit, and both source and destination sit
  inside the same Unit, so no permission changes and nothing to re-grant. This is the single reason
  this migration is a data move rather than a re-provisioning.

### 2.1 The two flows must be off

Moving a file in the approval library re-fires the **Auto-route** flow. Verified 2026-08-10: it does
not corrupt anything — it fails at `Copy_file` with "File not found", because the source has already
moved — but that is a race, not a guarantee, and a failed run per moved file makes the run history
useless for spotting a real failure.

**Turn off Auto-route and the folder-approval flow before migrating, and back on after.** The UI
states this as a precondition and cannot verify it — Power Automate is not readable from here.

### 2.2 Newly created destination folders must be approved

A folder created in a library with content approval on arrives **Pending**, and a pending folder is
invisible to everyone who cannot see drafts. The approval library has Draft Item Security set to
approver-and-author, so a destination folder created by an admin and left Pending would hide every
document moved into it from the entire unit — presenting as "the migration deleted our files".

So: read `EnableModeration` per library and stamp `OData__ModerationStatus = 0` on every folder this
tool creates, exactly as reconciliation does. Files are never touched by that stamp.

## 3. What counts as misplaced — detection, not guessing

The tool must not ask the admin which level is new; it must show evidence. For each unit folder,
walk the children and test each name against the chain:

- Child at position `i` whose name **is** a valid option for tier `i` → correctly placed.
- Child at position `i` whose name is not valid for tier `i` but **is** valid for tier `i + k`
  (smallest such `k > 0`) → **misplaced by k levels**, and it needs `k` new ancestor folders above it.
- Child matching **no** tier at any depth → **reported, never moved.** A folder someone created by
  hand, or a term that was deleted, is exactly the case where a confident guess does damage.

`k > 1` covers the client's "2 folders before year" case: the admin then picks a target for each of
the `k` new tiers, and the destination is `unit / t1 / t2 / … / oldFolder`.

Deriving `k` from the term data rather than from a question means the tool is also correct when
someone edits the structure twice before migrating, and it cannot be told a wrong answer.

### 3.0 SUPERSEDED 2026-08-11 — one algorithm, not three

Everything above describes the FIRST build, which looked only at folders **directly under the Unit**.
That detects an inserted tier and nothing else, and it failed on the first reorder: the client moved
`Year` above `Testing`, `Credit2` was still a valid first-tier value so the scan reported "nothing to
move", and the pending structure was activated against folders still in the old order. A silent wrong
success — the worst outcome this tool can produce.

The correction is not a special case for reordering. **Add, reorder and remove are the same
operation**, and treating them separately is what produced a tool that could only do one:

> For every folder below a Unit, work out which TIER each of its path segments belongs to, then
> rebuild the path in tier order.

| What the admin did | What that means here |
|---|---|
| **Added** a tier | Some tier has no segment in the existing path → a gap the admin fills with one chosen value |
| **Reordered** tiers | The segments map to tiers in the wrong order → the same segments, re-nested |
| **Removed** a tier | A segment belongs to no tier in the new chain → it is dropped, so sibling subtrees **collapse together** |

One computation covers all three, including combinations — a reorder *and* an insert in the same
edit, which is what a client will actually do.

**A DOCUMENT ALWAYS LANDS AT THE END OF THE CHAIN** (client rule, 2026-08-12). Every tier is built,
including ones below a folder's current depth, and any tier with no value in the path is a gap the
admin fills — exactly like a tier inserted in the middle.

This replaced a cap at the leaf's own deepest tier, and the cap made **adding a level at the bottom a
silent no-op**: the new tier sat below every leaf, so nothing looked misplaced, the pending change
activated anyway, and old documents stayed a level shallower than new ones in the same folder. Found
live 2026-08-12 by testing the one position the earlier four tests had not covered. "Add a level" now
behaves the same wherever it is added, which is what anyone using the page already assumes.

Two exceptions survive, and both are deliberate:

- **A tier that does not apply to a unit never reaches the calculation.** `effectiveTiers` drops a
  cascading tier with no values for that unit, so the optional-SubUnit case is untouched.
- **A folder with nothing recognisable left is reported, not placed.** When every segment belonged to
  a removed tier, the only destination would be the unit root — which contradicts this very rule —
  so it is left alone rather than having a whole path invented for it.

The cost, stated honestly: a genuinely shallow folder — a document filed at Year with no Document
Type — now asks for the deeper values too. That is the same rule applied consistently, and *Leave
this unit alone* remains the way to decline.

**Segment → tier assignment.** A segment's tier is the one whose option list contains its name. Where
a name is valid at more than one tier, the tier it is CURRENTLY at wins — the same
"already-correct beats speculative" rule as §3, applied per segment rather than per folder. A segment
matching no tier at any depth makes the whole folder a **stray**: reported, never moved. That is what
keeps the algorithm safe on a hand-made folder or a deleted term.

**Leaf folders are what gets planned; FILES are what actually move.**

A leaf folder — one with no subfolders — is where documents live in this system, because the upload
form always creates the whole chain before writing the file. So the plan is computed per leaf. But the
move itself is performed **file by file**, and the destination folders are ensure-created rather than
relocated. That is not the obvious choice, and the first draft had it the other way round; collapse is
what settles it:

- Removing a tier sends `Credit2/testig/2024` **and** `Credit2/live/2024` to `Credit2/2024`. As folder
  moves, the first succeeds and the second fails with "already exists" — the two cannot merge.
- A resolved collision **renames a file**, which a folder move cannot express at all.
- A destination folder may already exist and hold documents, so relocating a folder onto it is not a
  move but a merge, which SharePoint does not offer.

Moving files costs one request each instead of one per folder. Worth it: the alternative is two code
paths — folder moves for a reorder, file moves for a collapse — and the folder path would be the one
nobody tests, because collapse is the rarer operation.

A file sitting loose in an intermediate folder is still **reported, not moved**: its path does not say
which tier values it belongs to, so there is no destination to derive.

Source folders emptied by the run are deleted only when they hold **neither files nor folders**, and
only when this run emptied them. Anything else is left in place and reported — an unexpected leftover
is a fact worth showing, and deleting a folder is the one operation here with no cheap undo.

Every file move uses the **non-overwriting** form. If a collision somehow reaches the move stage, it
must fail loudly rather than silently replace a document — the collision detection is the safety net,
and this is the second one under it.

**Collapse is where the danger is.** Removing a tier maps two or more distinct paths onto one:

```
Credit2 / testig / 2024 / Tax Return / a.pdf   ┐
Credit2 / live   / 2024 / Tax Return / a.pdf   ┘ → Credit2 / 2024 / Tax Return / a.pdf
```

Folders merging is harmless. **Two files with the same name are not**, and SharePoint's instinct is to
overwrite — which reports success while destroying a document. So a collapse is planned in full before
anything moves, and every filename collision, including against files ALREADY at the destination, is
surfaced for a human decision (§4.1). Nothing moves until every collision has a resolution.

### 3.1 Valid option names

A tier with a `termSet` has one flat option list, shared by every unit. A **cascading** tier (no
`termSet`) draws its options from the children of the term above — so its valid names differ per
unit, and per department. That is not an inconvenience to work around: it is why the destination is
chosen **per unit** rather than once for the segment (§4).

Comparison is on the **sanitized term label**, because that is what below-Unit folder names are built
from — never an abbreviation (piece 1). Compare with the same `sanitizeFolderSegment` the upload form
uses, or a label containing a character SharePoint strips will read as "matches no tier" and be
reported as a stray folder.

## 4. The destination is chosen per unit

One dropdown per unit per new tier, listing that unit's valid options. Within a unit, every misplaced
folder moves into the same destination — which is the client's rule, and it means the choice count
stays at "one per unit", not one per year per unit.

A segment-wide "apply to all units" convenience is offered **only when the tier is flat** (has a
`termSet`). For a cascading tier there is no such thing as a shared value, and offering one would
either fail per unit or invent a term.

A unit with no destination chosen is **skipped**, not defaulted. There is no safe default: picking
the first option files a unit's entire history under a value nobody chose.

### 4.1 Filename collisions are resolved by a human, in a form

The client's request, 2026-08-11: rather than refusing a collapse, *show which files conflict and let
them rename each one*. Better than either refusing or auto-renaming, because the person deciding is
the only one who knows whether two files with one name are two documents or a duplicate.

**Detection precedes everything.** The full collapse is planned first: every file's destination
computed, grouped by destination folder + filename. A group of two or more is a collision, and files
**already at the destination** count as members — a collapse into an occupied folder is the case most
likely to be missed, and the one where an overwrite destroys a document nobody was even migrating.

**One file per group keeps its name.** The rest need new ones. There is no correct rule for which
keeps it, so the first in path order does, and every row is editable — including that one.

**Suggested names carry the information the removal is about to destroy:**

```
Credit2 / 2024 / Tax Return / a.pdf   ← 2 documents want this name

  from testig    a.pdf  →  [ a (testig).pdf ]
  from live      a.pdf  →  [ a.pdf          ]   keeps its name
```

`a (testig).pdf`, not `a (2).pdf`. The removed tier's value is *why* these two files were distinct;
`(2)` throws that away and leaves two documents looking like versions of one. A **Use suggestions for
all** control makes 200 collisions one decision instead of 200.

**Validation matches the upload form's rename** — illegal characters stripped, **extension preserved
automatically**, blank rejected — so there is one renaming behaviour in the product, not two. Plus
uniqueness within the destination, re-checked as they type, against both the other movers and what is
already there.

**Renaming happens as part of the move**, not before it. A file move can set the new name in the same
call, so no intermediate state exists in which two files compete for one path.

**The rows are rendered from the ORIGINAL names and stay put once shown.** The first build rendered
the *unresolved* list, so the instant a typed name settled a clash its row vanished — taking with it
any way to see or correct what had just been typed. Found live 2026-08-11 on the first real
collision. A mistyped extension would have applied silently, from a form whose entire purpose is
renaming. Two computations now, with two jobs: the original names drive what is DISPLAYED, the
renamed ones decide whether the run MAY PROCEED. Each group is marked *settled* or *needs a different
name*, so progress is visible rather than rows disappearing.

**Every typed name is validated, not silently corrected** (`validateRename`): blank, illegal
characters, and — the one that earns its place — **the extension must survive**. A rename that drops
`.pdf` leaves a file that opens as nothing, and one keystroke does it. Correcting quietly would be
worse than refusing, because the admin carries on believing they chose the name. A name can be unique
and still unusable, so validation gates the run independently of collision detection.

**Lead with the count, not the list.** A collapse producing 1,240 collisions is telling you the tier
being removed carries real meaning; the screen should say so before it renders a form nobody can
work through. The escape hatch stays available: remove that tier's values one at a time, or keep the
tier.

## 5. What the run does, in order

Per library (`Approval Document`, then `Documents`) — both hold the same tree and a unit can be
adrift in one and not the other:

1. **Create the destination folder** (`ensureFolder`, idempotent), then approve it if the library
   moderates (§2.2).
2. **Move the misplaced folder** into it — `MoveTo`, which is the same call `renameFolder` already
   uses; a rename is just a move within one parent.
3. **Stamp the new tier's columns** on every file beneath the moved folder: `labelCol` = the chosen
   label, `tidCol` = its term GUID. Via `validateUpdateListItem`, checking `HasException` per field
   (gotcha #4).
4. **Report** every move as `from → to`, every skip with its reason, every failure with its status.

Order matters: a move into a folder that does not exist fails, and a stamp on a file that has not
moved records a value its path contradicts. Failure at any step abandons **that folder** and
continues with the next — one bad folder must not strand the other forty.

### 5.1 Step 3 is not optional, and it is why the tool is re-runnable

A moved file whose new-tier column is blank has a path saying `testig` and metadata saying nothing.
Every view, filter and search on that column then misses the migrated documents while showing the
new ones — the same class of silent inconsistency as the missing `Documents` columns, which went
unnoticed for months.

So the plan contains **two independent parts**, both computed from what is on disk:

| Part | Finds | Action |
|---|---|---|
| A. Moves | folders sitting at the wrong depth | create destination, move, stamp beneath |
| B. Backfill | files already at the right depth whose tier column is blank | stamp only |

Part B makes the whole operation **idempotent**: a run that fails halfway through stamping can simply
be run again, because the second run sees the moved-but-unstamped files as part B work. It also
repairs files uploaded during the window between saving a structure and migrating it.

Part B derives the expected value from the **path**, and trusts the path over the column. The path is
where the document actually is; the column is a description of it.

**A tier with no `tidCol` is MANAGED METADATA and must be left alone** — found live 2026-08-11, on
the first real run. Two independent reasons, either sufficient:

1. A taxonomy field needs `Label|GUID` (gotcha #5). The bare label fails with *"The data returned
   from the tagging UI was not formatted correctly"* — and because one bad field fails the WHOLE
   `validateUpdateListItem` call (gotcha #4), it took the perfectly valid `CreditCard` write down
   with it. The same failure that once made bulk upload tag nothing at all.
2. Nothing needed writing anyway. A migration INSERTS an ancestor tier; the `2024` and `Tax Return`
   folders keep their names and their values, so those columns were already correct. Reading a
   taxonomy field back as a plain string yields `""` — the value is an object — so **every file
   looked like it needed a stamp it did not need.**

The `tidCol` test is exact rather than a proxy: the plain-text label+GUID pair is what the
multi-segment model writes, and the only tiers without one are the built-in Year / Document Type
pair synthesised by `effectiveOnDemandTiers`, which omits it deliberately.

Stamps are also **grouped per tier and retried individually** when the combined call fails, so one
unwritable column cannot cost the others — and the report names the tier, not just the file.

**Part B must be reachable WITHOUT a move**, through its own *Check document tags* button. The first
live run proved why: the folders moved, the tagging failed, and the retry then found nothing to move
— so the only route to the repair was closed and the metadata could never be fixed. **A repair step
reachable only through the thing that broke it is not a repair step.** Running it is safe at any
time and as often as wanted: it derives its work from the paths, writes only what disagrees, and
touches no folder.

### 5.2 A structure change is STAGED, and applying it is the last step of the migration

Added 2026-08-11 at the client's request, before any of this reached a user. The original build
applied a structure change immediately and left the migration as a later, optional job — so
between the two a segment genuinely has two shapes, and **the people who meet that state first are
uploaders nobody told.** It reads as a broken system rather than an unfinished admin task, and it
generates complaints that cannot be answered honestly.

So on a segment that already holds documents, the Structure Manager writes the new chain to
`PendingLevels`, NOT `Levels`:

| Column | Read by | Meaning |
|---|---|---|
| `Levels` | upload form, reconciliation | what uploads are doing **right now** |
| `PendingLevels` | Structure Manager, this screen | authored and reviewed, **not yet applied** |

The sequence becomes author → migrate → **apply, automatically, as the final step of the run**.
Uploaders meet the new dropdown at the moment the folders are already correct.

Four rules hold it together:

1. **Applying is never a separate button.** Between "folders moved" and "structure applied",
   uploads would land in the old shape again and re-create the drift just cleaned up. The two must
   not be separable in the UI.
2. **It applies only when a FRESH scan finds nothing left in the old shape** — not a tally of what
   the run attempted, because a move that reported success but landed somewhere unexpected can
   only be caught by looking. A partially migrated segment stays pending and says how many units
   remain.
3. **An empty segment activates immediately.** Nothing to move, and making someone run a migration
   over zero folders teaches them to click through it. A staged-then-emptied segment has its
   pending chain cleared on save, or it would offer to apply a change already applied.
4. **A staged change with nothing to move must still be applyable**, or it could never go live —
   which happens whenever the new level only affects units holding no documents. The no-drift
   result therefore carries an explicit *Apply the new structure* button.

`PendingLevels` is a **Note** column, created on demand. Note rather than Text because a five-level
chain with internal names and term-set GUIDs passes 255 characters easily, and a silently truncated
chain would either fail to parse or — far worse — parse as a shorter structure than was authored.
Created on demand rather than documented as a provisioning step because a missing column would fail
the save on an otherwise correct site, and the person hitting it would have no way to know why.

The Structure Manager's typed-`CHANGE` gate was **removed** with this change. It warned about two
folder shapes side by side; staging removed the thing it warned about, and saving is now inert —
nothing moves, uploads are unchanged, and the chain can still be edited. The gate that matters is
the typed `MOVE` on the migration, the step that actually relocates documents. A scary modal on a
harmless action is how someone learns to click through the one that counts.

Editing a segment that already has a pending change edits **the pending chain**, not the live one.
Otherwise a second edit silently discards the first, and the migration would apply a shape nobody
reviewed.

## 6. A dry run is the default and cannot be skipped

The plan is always computed and displayed first: every move as `from → to`, grouped by unit, with
counts of files affected. Nothing is written until the admin confirms.

For a segment that already holds documents the confirmation is the typed word `MOVE` — the same
device the Structure Manager uses for a structure change, and for the same reason: this is the one
action on the page that a second click cannot undo.

## 7. Explicitly out of scope

- **Splitting one old folder across several new ones.** The client ruled it out and it has no correct
  automatic answer — `2024`'s files may belong under several destinations and only a human reading
  each document could say which.
- **Moving individual files.** Only folders move. A file sitting loose directly under a unit is
  reported, not relocated: it has no tier folder to infer a destination from.
- **Deleting anything.** An emptied source folder is left in place. It is harmless, and deleting a
  folder is the one operation here with no cheap undo.
- **Undo.** There is none. The mitigation is the dry run, not a rollback — a reverse-move log sounds
  reassuring and would be wrong the moment anyone uploaded into the new shape.
- **Structural (permissioned) tiers.** Segment, Department and Unit folders never move; they carry
  ACLs and Folder Map rows. Reordering those is not a migration, it is a re-provisioning.

## 8. Errors

| Condition | Behaviour |
|---|---|
| Chain fails `validateChain` | Refuse the whole segment. A malformed chain has no correct destination |
| Folder matches no tier at any depth | Report as a stray. Never moved |
| Destination folder cannot be created | Skip that folder, report the status. Do not move into a guessed parent |
| Move returns "already exists" | Report as a collision, skip. Merging two folders merges two units' history |
| Stamp fails (`HasException`) | Report per folder, keep the move. Part B will retry it on the next run |
| Library has moderation and the approve stamp fails | Report loudly — the folder is invisible to the unit until fixed |
| No destination chosen for a unit | Skip silently in the plan; the unit is listed as "no destination chosen" |

## 9. Testing

- **Insert one tier above Year on a used segment**, migrate: `2024` and `2025` land under the chosen
  destination with their document types and files intact, and approval status unchanged.
- **Insert two tiers**, migrate: destination is `unit / t1 / t2 / 2024`, and the two intermediate
  folders exist and are approved.
- **Insert a tier at the deepest position**: detection reports nothing to move, because files sit in
  the last folder and files are never moved. The plan must say "nothing to move", not show an empty
  success.
- **A cascading tier**: each unit offers its own destinations, and the "apply to all" control is absent.
- **A unit with no destination chosen**: nothing moves for it; the report says so.
- **A stray hand-made folder**: reported, not moved, and does not block the rest of the unit.
- **Re-run immediately after a successful run**: finds nothing — no duplicate folders, no re-stamping.
- **Re-run after killing the browser mid-stamp**: part B finishes the stamping and part A finds nothing.
- **Both libraries**: a unit adrift only in `Documents` is migrated there and untouched in the
  approval library.
- **Moderated approval library**: destination folders come out Approved and a PIC in the unit can see
  the moved documents.

## 10. Implementation notes worth keeping

- **Files are read WHOLE-LIBRARY, once, and filtered per unit in memory** — not one CAML query
  per unit. `GetItems` returns no paging token in that shape, so a per-unit CAML query silently
  stops at one page or repeats page one forever while reporting progress. `/items?$top=5000` +
  `@odata.nextLink` pages correctly, and carries `Id`, which also removes one resolve request
  per file. A column named in `$select` that does not exist fails the WHOLE request with HTTP
  400 (gotcha #11), so that error is reported as "could not read <library>" and names the status.
- **The term cache is a `useRef(Map)`, not state.** The scan runs one long async loop; a state
  update would not be visible to the next iteration, so every unit would re-fetch the same flat
  term set. `Map.set` also sidesteps the `require-atomic-updates` lint that a plain object
  assignment after `await` trips.
- **Both libraries share one destination choice**, keyed on the unit path tail (`gho/gf/coru`),
  which is identical in both. Two separate choices for the same unit is a state nobody wants and
  the UI should not be able to express.
- **Choosing a destination clears every deeper choice.** A deeper cascading tier's options come
  from the tier above, so keeping them would offer a value belonging to a different parent.

## 11. Related

- [2026-08-06-configurable-folder-structure-chain-design.md](2026-08-06-configurable-folder-structure-chain-design.md) — piece 1, the chain model
- [2026-08-10-structure-manager-ui-design.md](2026-08-10-structure-manager-ui-design.md) — piece 2, the editor whose warning this resolves
- [2026-08-08-auto-route-flow-and-draft-isolation.md](2026-08-08-auto-route-flow-and-draft-isolation.md) — the flows that must be off, and why folders must be Approved
