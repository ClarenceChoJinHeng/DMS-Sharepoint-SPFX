# The HC libraries are not migrated when the folder structure changes

**Date:** 2026-08-19
**Status:** design agreed, NOT implemented
**Register:** #15
**Extends** `2026-08-11-subtree-migration-design.md`, which is otherwise unchanged and still governs.

---

## 1. The gap, verified against current source

`SubtreeMigrator` moves existing folders into a new shape when a segment's structure changes. It
walks **two** libraries. There are **four**.

| Evidence | |
|---|---|
| `SubtreeMigrator.tsx:100` | `interface LibCtx { key: "Staging" \| "Documents"; … }` — the HC pair cannot even be *expressed* |
| `SubtreeMigrator.tsx:587-590` | the library list is a two-element literal |
| `SubtreeMigrator.tsx:1176` | `for (const lib of ["Staging", "Documents"])` — the conflict and unresolved-count loop |
| `SubtreeMigrator.tsx:1119` | the audit row records `library: "Approval Document + Documents"` |

`allLibraryTitles()` has existed in `shared/naming.ts:400` since the HC work, and **`SegmentCreator`
and `StructureManager` both use it** to reach all four. The migrator was written 2026-08-11, before
the HC pair existed on 2026-08-15, and was never revisited.

## 2. What happens today

A client adds or reorders a level on a segment that holds HC documents:

1. `StructureManager` writes `PendingLevels` and **creates the tier columns in all four libraries**
   (it uses `allLibraryTitles()`), so the HC libraries even have the columns waiting.
2. The migrator scans, finds work in `Approval Document` and `Documents`, moves those files,
   re-stamps their metadata from the path, and **applies `PendingLevels` to `Levels`**.
3. The HC pair is never scanned. Nothing there moves.
4. The structure is now live. Every existing HC document sits in the **old** shape; every new HC
   upload lands in the **new** one.

**It reports success**, and the only place the omission is visible is an audit row naming two
libraries — visible only to someone who already suspects it.

### 2.1 The safety mechanism actively conceals it

Applying the pending chain is deliberately gated on **a fresh scan finding no drift left** (2026-08-11
spec) — never on a tally of what the run attempted, precisely so a partial migration stays pending.

A scan that never looks at the HC libraries always finds no drift there. **The guard confirms the
state is clean because it is not looking.** That is why this is worse than an ordinary "half
migrated": the mechanism designed to catch a half migration is the one reporting it complete.

## 3. The change

### 3.1 `LibCtx.key` becomes `LibTarget`

The two-literal union becomes the project's own `LibTarget`, so `StagingHC` and `DocumentsHC` are
expressible **and the compiler finds every place that assumed two**. Doing this first is what makes
the rest mechanical rather than a hunt.

### 3.2 The library list is derived, never literal

Built from the resolved names as reconciliation builds its own, skipping the HC pair on a site
without one. **`hcAvailable()` is the gate and it fails CLOSED by design**: `naming.ts` has no HC
fallback, so an unresolved HC library is `undefined` rather than quietly resolving to a legacy
literal. A site without HC migrates two libraries and says so in the log.

### 3.3 The conflict loop follows

`SubtreeMigrator.tsx:1176` iterates the same derived list. This is what makes an HC filename
collision reach the rename form. Without it the run would move HC files with **no collision detection
at all** — and collision handling is the one part of this tool standing between a move and an
overwritten document.

### 3.4 The audit row names what it actually walked

A record that says "Approval Document + Documents" on a four-library run is worse than one that names
none: it reads as a decision rather than an oversight.

## 4. What must NOT change

- **The per-library moderation read stays** (`SubtreeMigrator.tsx:591`). `HC Approval Document`
  moderates and `HC Documents` does not, and folders this tool creates in a moderated library must be
  stamped Approved or **every document moved into them is invisible to the whole unit**. The literal
  at 587-590 initialises `moderated: false`, but line 591 immediately asks each library — that is an
  initialiser, not an assumption. Do not "simplify" it to a constant.
- **HC folders have NO Folder Map rows** (2026-08-15 spec §folder resolution); their paths are derived
  by swapping the library segment. So the tail-keyed term lookup will legitimately miss for HC
  folders, and those units must be reported **unresolved**, never guessed. Unresolved blocks Run,
  which is correct: a below-Unit tier whose term cannot be resolved would file HC documents one level
  shallow, silently.
- **The unit grouping by tail stays.** One set of pickers already serves both libraries because a unit
  is adrift the same way in each; that generalises to four with no UI change.
- **The staged-then-applied model stays.** Nothing here touches when `PendingLevels` becomes `Levels`
  — only which libraries are scanned before it does.

## 5. Failure rules

| Condition | Behaviour |
|---|---|
| Site has no HC pair | two libraries, stated in the log and the audit row. Never an error. |
| One HC library resolves and its twin does not | **refuse the run.** Both halves resolve or neither does — migrating an HC approval library whose documents library is missing leaves approved HC files with nowhere to route. |
| An HC folder tree cannot be read | that unit is **unresolved**, which blocks Run. Deliberately unlike the fail-open reads elsewhere in this codebase: there a failed read costs a form for a minute, here it writes documents to wrong paths. |
| An HC collision the admin has not named | blocks Run, exactly as a normal one does. |
| An HC file move fails | counted in `failed`, run outcome `Failed`, pending chain NOT applied. Unchanged behaviour, now reachable for HC. |

## 6. Testing

The 2026-08-11 work verified five structure positions — add mid-chain, reorder, remove-clean,
remove-with-collision, add-at-bottom. **The bug found then had survived four passing tests, because
every one happened to use full-depth folders: the untested position was the broken one.**

So the test is not "does it move HC files" but "does every position still behave with HC present":

1. One HC document in a unit, plus a normal document in the same unit.
2. Add a level mid-chain → both move, both re-tagged, four libraries reported.
3. Reorder → **0** documents re-stamped in every library including HC (a reorder changes no tier's
   value, so anything else is a bug).
4. Remove a tier with two HC files colliding → the rename form lists the HC claimants, and the
   document count is unchanged afterwards.
5. Confirm `Levels` applied only after the HC libraries are also clean.

## 7. Files

| File | Change |
|---|---|
| `src/webparts/userAccess/components/SubtreeMigrator.tsx` | `LibCtx.key` → `LibTarget`; derive the library list; the loop at 1176; the audit row |
| `src/shared/subtreeMigration.ts` | none expected — it is library-agnostic, which is why the fix is contained |
