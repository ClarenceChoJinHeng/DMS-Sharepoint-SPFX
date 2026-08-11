# Subtree Migration — moving documents that are already filed into a new folder shape

**Date:** 2026-08-11
**Status:** built, NOT yet site-tested. `SubtreeMigrator.tsx` (tab 2 of the Folder Structure web
part) + `shared/subtreeMigration.ts` (32 unit tests).
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
