# Structure Manager — client-editable folder structure

**Date:** 2026-08-10
**Status:** BUILT and SITE-VERIFIED 2026-08-11 on NBPOL — a level added above Year, columns created in both libraries, `Levels` rewritten, upload landed at the new path. Header corrected 2026-08-19.
**Scope:** Piece 2 of 3. Depends on piece 1
([2026-08-06-configurable-folder-structure-chain-design.md](2026-08-06-configurable-folder-structure-chain-design.md),
built 2026-08-09).
**Touches:** a new tab in the admin web part, `DMS Config`, and site columns on both libraries.

---

## 1. Why this exists

Piece 1 made the folder structure data-driven. It did not make it *editable by the client*, and the
client cannot edit JSON — which is the entire reason the feature was requested. Today, adding one
folder tier takes three manual steps:

1. Create the term set (fine — the client is happy to do this)
2. Create a text column in **both** libraries, under a name that yields the right internal name
3. Hand-edit the `Levels` JSON on a `DMS Config` mode row

Steps 2 and 3 are the ones to automate. Step 3 especially: a hand-edited chain is exactly where a
malformed structure comes from, and `validateChain` can only reject it after it has been authored.

**The forcing constraint:** Clarence will not be supporting this client after handover. The client
has supplied Head Office data only, and their own structure table lists five more segments plus
*"2 additional layers for buffer — Trinergy to take note and action on."* Anything not automated is
work nobody will be there to do.

## 2. The shape is uniform, and that is the design

From the client's structure table (2026-08-10):

| Segment | Tier 1 | Tier 2 | Below Unit |
|---|---|---|---|
| Head Offices | Department | Unit | Year, Document Type |
| Upstream Operations | Region | Estate/Mill | Year, Document Type |
| SDGI Operations | Refinery | Department | Year, Document Type |
| Group Business Ventures & Transformation | Department | Unit | Year, Document Type |
| Innovation & Technology | I&T Operating Unit | Department/Unit | Year, Document Type |
| Group-led projects | Department | Unit | Year, Document Type |

**Every segment is `[top] → [tier 1] → [tier 2] → Year → Document Type`.** The shape never varies;
only the labels and the columns behind them do.

So "add a segment" is a **fill-in-the-blanks form, not a structure builder**. That is a materially
smaller surface to hand an untrained admin, and it removes whole classes of mistake: a segment
cannot be created one tier deep, or six, or with its permissioned tiers in an order reconciliation
cannot walk.

Two permissioned tiers is the *default*, not a hard limit — the form offers a third slot, blank by
default. Hardcoding two would bet that no 2027 segment needs a different depth, and that bet is paid
for with a redeploy nobody will be around to do.

## 3. Decisions taken

**Editing a live segment's below-Unit chain is ALLOWED, with a hard warning** (client decision,
2026-08-10, overriding the recommendation to block). The client explicitly wants to add folders
below Unit over time, and blocking it would put the one thing they asked for behind a person who is
no longer engaged. The two-tree consequence is real (§6) and is stated in the confirmation, in plain
language, with the document count.

**"Locked" is NOT the same as "permissioned"** — added 2026-08-10, when the client introduced
`SubUnit`. SubUnit is a **fixed** tier the client may not reorder or remove, but it **inherits** the
Unit's ACL, so it is authored `"permissioned": false`. The fixed prefix is now four tiers deep while
the permissioned prefix is still three.

The editor therefore cannot derive lockedness from `permissioned`, which was the original plan. It
needs its own signal — `locked: true` on the level entry, with permissioned tiers implicitly locked.
Deriving it would let the client drag `SubUnit` around, and since SubUnit sits above Year, moving it
is a migration of every document in the segment.

The two locks mean different things and must not read alike on screen:

- **Permissioned tiers** — "has folder permissions, cannot be changed"
- **SubUnit** — "part of the standard structure, cannot be changed", and explicitly **not** a
  security boundary: everyone in the Unit sees every SubUnit under it

That second caption matters. A tier named SubUnit sitting directly under Unit in a list of locked
tiers looks exactly like a permission boundary, and the client confirmed on 2026-08-10 that it is
not one. Leaving it to be inferred is how someone eventually files something sensitive in a SubUnit
believing it is separated.

**The permissioned prefix is immutable once created.** Segment / Tier 1 / Tier 2 are *set* when a
segment is created and never edited afterwards. This is what Clarence is telling the client, and the
UI must enforce it rather than rely on them remembering. Renaming a permissioned tier renames live
ACL'd folders; reordering one asks reconciliation to walk a tree that does not exist. Neither is a
config edit.

**Term sets stay manual.** The client creates the term set and pastes its GUID. Writing to the term
store is a different API surface, and pasting a GUID forces them to confirm the set exists before
the structure references it. A structure pointing at a nonexistent term set produces an empty
dropdown, which reads as "the system is broken".

**New columns are plain TEXT, never managed metadata.** Already true of every level column
(`Department`/`DepartmentTid` are text pairs), and it is what makes automation a single REST call.
A taxonomy column needs a term-set binding and a hidden note field, and buys nothing — the term GUID
is already stored in the `*Tid` column.

## 4. What the interface does

A **Folder Structure** tab. Two screens.

### 4.1 Segment list

Every `DMS Config` mode row, with its chain rendered as a path and a document count:

```
Group Head Office     GHO / [Department] / [Unit] / [Year] / [Document Type]      1,847 docs
Minamas Head Office   MHO / [Department] / [Unit] / [Year] / [Document Type]         12 docs
NBPOL Head Office     NBPOLHO / [Department] / [Unit] / [Year] / [Document Type]      0 docs
                                                                     [ + Add segment ]
```

The document count is not decoration — it is the number that decides whether an edit is routine or a
migration, and showing it *before* the click is what makes the later warning read as information
rather than boilerplate.

### 4.2 Structure editor

```
Group Head Office                                          1,847 documents filed

  Permissioned tiers        these have folder permissions and cannot be changed
    1  Department           column: Department   + DepartmentTid    [locked]
    2  Unit                 column: Unit         + UnitTid          [locked]

  Folders below Unit        these inherit the unit's permissions
    3  Year                 term set: 023a866a-…            [up] [down] [remove]
    4  Document Type        term set: 866c5754-…            [up] [down] [remove]
                                                        [ + Add folder level ]

  Path preview
    GHO / GF / CORU / 2026 / Invoice

                                                    [ Cancel ]  [ Save structure ]
```

The **path preview updates live** and uses real values from the segment. It is the only
representation of the change the client can actually reason about; the JSON is never shown.

The locks and the "cannot be changed" caption are load-bearing: they teach the rule Clarence is
relying on them to remember, at the moment it applies.

### 4.3 Add folder level

```
  Name          [ Function              ]   shown above the dropdown on the upload form
  Term set ID   [ paste from term store ]   (i) where to find this
  Position      ( ) Before Year
                (o) Between Year and Document Type
                ( ) After Document Type

  Column        Function                    created automatically in both libraries
```

The column name is **derived from the tier name, shown read-only, never typed**. A SharePoint
column's internal name is fixed at creation, permanently, from the title it was created with — so
letting an untrained admin type it offers a decision that cannot be undone and whose consequence is
invisible. Derive it: strip spaces and punctuation, create under that name, then set the display
title to what they typed. (This is exactly why `DocumentDate` displays as "Document Date" and has no
`_x0020_`.)

### 4.4 Add segment

```
  Segment name  [ Upstream Operations   ]
  Term set ID   [ paste from term store ]
  Folder code   [ UPO ]                     the top-level folder name in both libraries

  Tier 1        [ Region      ]   column: Region      + RegionTid
  Tier 2        [ Estate/Mill ]   column: EstateMill  + EstateMillTid
  Tier 3        [             ]   optional — leave blank for a two-tier segment

  Below Unit    Year, Document Type          added automatically; can be changed later
```

## 5. What it writes, and in what order

Every save is a sequence with a defined failure point. Order matters: a column created without a
chain is a harmless orphan, whereas a chain saved against a missing column breaks **every upload in
that segment** — and `validateUpdateListItem` returns HTTP 200 with `HasException` (gotcha #4), so
that failure is not even loud.

1. **Validate** the resulting chain with `validateChain`. Reject before touching anything.
2. **Create missing columns** in `Approval Document` and `Documents` — idempotent, skip if present.
   Create under the derived space-free name, then set the display title.
3. **Write the `Levels` JSON** to the mode row (creating the row itself, for a new segment).
4. **Report** what changed, naming the columns created and the libraries touched.

If step 2 fails on the second library, stop and say so. A column present in one library and absent
from the other routes files correctly and loses their metadata in one library only — the hardest of
these states to notice.

Resolve both library names through `libraryTitle()` / `libraryUrlSegment()`; never hardcode either
half (gotcha #12).

New columns are created with `Options: 8` (AddToAllContentTypes) and are therefore **not added to
any view** — deliberately, and **not** `12`, which also adds each one to the default view and would
reshape every view the client arranged, once per level anyone ever adds. Verified on
ClarenceDMSTesting 2026-08-11: `Testing` / `TestingTid` were created in both libraries, written
correctly on upload, and had to be switched on per view via *Show or hide columns*.

The cost is that a freshly created column is invisible where the client looks for it, which is
indistinguishable from the save having failed — so the success banner states it explicitly. Do not
"fix" this by switching to `12`.

### 5.1 A new segment is NOT finished when this form is submitted

The mode row is one of six things a segment needs. The form must end on an explicit checklist, not a
success toast, or the client will believe they are done:

| Step | Automated here? |
|---|---|
| `DMS Config` mode row + `Levels` | YES |
| Level columns in both libraries | YES |
| Term set + its terms | No — client, in the term store |
| **Abbreviation row for every term** | No — client. **A term with no abbreviation gets NO folder** |
| Security groups per unit + Group Map rows | No — Folder Access tab |
| Folder Reconciliation run | No — button, run after all of the above |

The abbreviation line is the one that silently does nothing: reconciliation skips a term with no
abbreviation rather than guessing, so a fully configured segment with none provisions an empty tree
and reports success. Say this on the screen, not only in a document.

## 6. The two-tree warning

Shown whenever a below-Unit change is saved for a segment whose document count is above zero. The
wording matters more than the mechanism — this is the client's only chance to understand it:

```
!  This changes where new documents are filed

   Group Head Office has 1,847 documents already filed.

   Those documents stay exactly where they are:
       CORU / 2026 / Invoice
   New uploads will go to:
       CORU / Human Resource / 2026 / Invoice

   Both folders will exist side by side. Nothing moves the older
   documents, and "all of CORU's 2026 invoices" will no longer be
   a single folder.

   This is safe on a new segment. On a segment already in use,
   plan it deliberately.

   Type ADD LAYER to confirm:  [__________]
                                    [ Cancel ]  [ Confirm ]
```

The typed confirmation is not friction for its own sake: this is the only action in the tool whose
consequence cannot be undone from the tool.

> **Moving documents already filed is now BUILT** — see
> [2026-08-11-subtree-migration-design.md](2026-08-11-subtree-migration-design.md), tab 2 of this
> same web part. The warning below still stands: this screen changes new uploads only, and the
> migration is a separate, deliberate step.

## 7. Out of scope

- **Moving existing documents** under a newly inserted tier — piece 3, still unspecified. Cost scales
  with insertion depth: above Year is a few folder moves per unit; below Document Type is thousands
  of file moves.
- **Term set and term creation.**
- **Abbreviations, groups, Folder Map rows** — existing tabs own these.
- **Deleting a segment.** Removing a mode row strands its folders and every group pointing at them.
  If a segment is wrong, the answer is a new one.
- **Editing the permissioned prefix** — §3.

## 8. Errors

| Condition | Behaviour |
|---|---|
| Term set ID malformed | Reject on the form. Add stays disabled |
| Term set ID does not resolve on this site (404) | Reject on the form. Covers the two commonest mistakes: an ID copied from another site (term sets are per-site here) and a TERM's ID pasted instead of the set's — well-formed, and 404s |
| Term set resolves | Show its **name** and top-level term count back. The name is what actually catches a wrong-but-valid ID; resolution alone only rules out typos |
| Term set resolves but holds no terms | **Warn, allow.** Authoring the tier before its terms is a legitimate order of work, and the warning states what is still owed. Distinct from the not-found case, which cannot be resolved by adding terms |
| Term store unreachable / HTTP 5xx | **Warn, allow.** Says nothing about the ID; refusing here would block a correct one |
| Column exists in one library only | Create the missing one; report both |
| Column exists with a different type | Refuse; do not attempt conversion. Name the column and the library |
| Chain fails `validateChain` | Refuse, quoting the rule that failed. Never write a partial chain |
| Mode row write fails after columns created | Report the orphan columns explicitly — harmless, but must not be silent |
| Duplicate tier name within a segment | Refuse — two tiers writing one column overwrite each other |

### 8.1 Unsaved changes are guarded at all three exits

Added 2026-08-11 after it happened during testing: a level was added, the page was left, and the
work was gone with **nothing on screen afterwards to say so**. That is the failure worth designing
against — the list looks finished while it is being edited, so an abandoned edit is
indistinguishable from one that was never started.

Three exits, three guards, each as strong as what it can actually protect:

| Exit | Guard | Why this one |
|---|---|---|
| **Cancel** button | modal offering *Keep editing / Save them / Discard* | They asked to leave, so discarding must be available — just not as the default |
| **Tab switch** | **refused**, with a message | The Save button is a few pixels away; offering "discard" here would put losing the work one click behind something that reads as ordinary navigation |
| **Closing the tab / typing a URL** | `beforeunload` | The only exit the page cannot mediate |

Dirtiness is a JSON comparison against **what was seeded into the editor**, not against the mode
row — `startEdit` seeds from `effectiveOnDemandTiers`, so comparing against the row would mark a
segment running on the built-in Year/Document Type pair as dirty the moment it opened, and an
always-on warning is one nobody reads. Reordering a level and putting it back reads as clean for
the same reason.

`beforeunload` is attached **only while dirty**. An unconditional handler prompts on every exit,
and a prompt that always fires gets clicked through — including the once it mattered.

The inline warning sits **beside the Save button**, not at the top of the editor: a banner above a
long form is scrolled out of view at exactly the moment someone is about to leave.

## 9. Testing

- A new segment end to end: form → columns exist in both libraries → mode row correct → upload lands
  at the right path with the right metadata.
- Adding a tier to an **empty** segment: no warning, works.
- Adding a tier to a **populated** segment: warning shows the correct count; cancel leaves everything
  untouched; confirm writes.
- Re-saving an unchanged structure: no duplicate columns, no rewritten JSON.
- Term set ID checks: a wrong-length string is refused; a well-formed but unknown GUID is refused;
  a TERM's GUID is refused (it 404s); a real set shows its **name** back, which is the check an
  admin can actually act on; an empty set warns but is allowed.
- **Verified end to end on ClarenceDMSTesting, 2026-08-11.** `Testing` added to NBPOL (own term set
  `Function2`, one term `testig`), positioned above Year: columns created in both libraries, `Levels`
  rewritten, and an upload landed at `NBPOLHO / CDS / UPSUPPORT / testig / 2024 / Tax Return` with
  Year and Document Type intact below the new level — the case the seeding fix exists for.
- The typed ID is stored **normalized** — a paste with braces or trailing whitespace must not reach
  `Levels`, where it would 404 at upload time on a site nobody is watching.
- Permissioned tiers cannot be renamed, reordered or removed through the UI at all.
- After a new segment: reconciliation provisions it, and a term with no abbreviation is reported and
  skipped rather than guessed.

## 10. Related

- [2026-08-06-configurable-folder-structure-chain-design.md](2026-08-06-configurable-folder-structure-chain-design.md)
  — piece 1, the chain model this edits
- [2026-07-21-segment-onboarding-plan.md](2026-07-21-segment-onboarding-plan.md) — the manual
  onboarding steps this partly replaces
- [2026-07-30-folder-abbreviation-naming-design.md](2026-07-30-folder-abbreviation-naming-design.md)
  — why a term with no abbreviation gets no folder
- Piece 3 (subtree migration) — unspecified
