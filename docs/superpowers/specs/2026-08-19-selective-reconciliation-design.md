# Selective reconciliation — run the segments that changed

**Date:** 2026-08-19
**Status:** HALF A BUILT 1.0.175.0 (the segment picker), NOT site-tested. ⚠ Half B — detecting which segments changed — is deliberately DEFERRED: it needs a per-segment term walk at picker time (~115 requests for GHO), which is the expensive thing this feature exists to avoid. Half A delivers the whole time saving; the marks were advisory.
**Register:** #18
**Client's ask, 2026-08-19:** *"if client only added new subunit or new unit under two business
segment's department and they have 5 Business segment in the future, this will slow them down if
folder recon have to go through everything again, might as well when folder recon is going to run we
ask client which business segment they want to run that have an update, they tick and select then
run… If possible we detect for any changes only then show which business segment that the folder
recon have to go through."*

**Relates to** `2026-07-20-reconciliation-provisioner-design.md` (what a run does) and
`2026-07-27-portable-columns-and-reconciliation-ux-design.md` (progress and throttling *within* a
run). Neither addresses which segments a run covers.

---

## 1. The problem, measured

The from-scratch run on 2026-08-19 produced **4,552 log lines** across four libraries for **one**
segment and took over an hour. A check run — everything already in place — took about forty minutes.

Adding one unit under one department therefore costs that same forty minutes to change **one folder in
four libraries**, and the one line that matters sits among four thousand that say `already there`. At
five segments it is over three hours, and the client will simply stop running it — which is worse
than slow, because reconciliation is the step that makes access real.

## 2. Two halves, and only one of them is safe to trust

**Half A — pick the segments.** Tick which to run. Mechanical: the run already loops over modes and
this filters that loop.

**Half B — detect what changed, and say so.** Compare the term tree and the abbreviation rows against
what the Folder Map records, and mark the segments that look like they have work.

**⚠ DETECTION RECOMMENDS. IT NEVER RESTRICTS.** Every segment stays tickable, always. This is the
whole safety of the feature:

- Detection reads the same lists reconciliation reads. **A failed read, or a shape nobody
  anticipated, produces "no changes"** — and a segment silently excluded leaves folders unbuilt, ACLs
  ungranted and uploads refused, with a green log saying nothing needed doing.
- The failures reconciliation exists to repair are exactly the ones **no comparison can see**: an ACL
  removed by hand in SharePoint, a folder renamed in the library, inheritance broken by an
  administrator, a group deleted and re-created with a new id, a page grant stripped. None of those
  changes a term or a row.

So a "nothing changed" verdict is a statement about the **lists**, never about the **site**. The UI
must never present it as authority: it says *"looks like it needs a run"*, and it never unticks
anything on the user's behalf.

## 3. The picker

```
Which business segments should this run cover?

  [x] Group Head Office            ● 3 units with no folder yet, 1 code changed
  [x] Minamas Head Office          ○ nothing obvious to do
  [x] NBPOL Head Office            ○ nothing obvious to do
  [x] Upstream Operations Malaysia ⚠ could not check — its term set did not answer

  These marks are a hint from the term store and the folder map. They cannot see a permission
  someone changed by hand in SharePoint, so if access is wrong for a segment, run it even when
  it says there is nothing to do.

  [ Select all ]  [ Only the ones with changes ]          [ Run reconciliation ]
```

- **Default: every segment ticked** — today's behaviour. A default of "only what changed" makes the
  first run after an unseen manual edit skip the one segment that needed it.
- **`Only the ones with changes` is one click**, so the fast path costs nothing but is a deliberate
  act rather than a default.
- **An unticked segment is not "clean"**, and no copy may imply it.
- **Running zero segments is refused**, with the reason. An empty tick list is a mis-click.

## 4. What detection can see — and what it cannot

Computed from reads the run already performs:

| Signal | Source | Meaning |
|---|---|---|
| Terms with no Folder Map row | term tree × Folder Map | new unit or department — folders missing |
| Folder Map rows whose term is gone | Folder Map × term tree | deleted or re-parented term; a stray is likely |
| Abbreviation code ≠ the folder's tail name | Abbreviation rows × Folder Map `folderUrl` | a re-code — folders will be renamed |
| **Terms with no abbreviation row** | term tree × Abbreviation rows | **folders will be SKIPPED** — the silent failure the whole system guards against |
| Group Map rows whose group is gone | Group Map × site groups | a mapping pointing at a deleted group |
| `PendingLevels` set | mode row | a staged structure change awaiting migration |

What it **cannot** see, stated in the UI: folder ACLs, broken inheritance, hand-renamed folders, page
grants, content types, `Full Name` values, moderation state. **That is everything reconciliation
asserts** — detection compares *inputs*, reconciliation asserts *outcomes*, and the gap between those
two is the entire reason the marks are advisory.

### 4.1 Cost, and why it is affordable

The term-tree walk is the expensive part (~115 requests for GHO), bounded by `permissionedDepth`. The
picker is affordable **only because it reuses the reads the run performs anyway** — computed once at
mount, reused when Run is pressed. A picker that doubles the requests in order to save time is not a
saving.

**If a walk would exceed its cap, that segment reports `could not check`** — never a partial answer.
A term tree read to 90% is indistinguishable from a segment with fewer units.

## 5. What must not change

- **The run itself is untouched.** Every pass, failure rule and log line stays; only the set of
  segments entering the loop narrows.
- **⚠ THE SITE-WIDE PASSES STILL RUN IN FULL, EVERY TIME.** Site entry, library state, page access and
  the admin-page lockdown assert state that **has no segment**. Scoping them to the ticked segments
  would recreate precisely the class of bug this codebase has now found four times — a mechanism
  driven by rows, and the thing needing protection having none. Their cost is a handful of requests;
  their absence is an exposed library or a locked-out administrator.
- **The log and the audit row must name the scope**: `Reconciliation — 2 of 5 segments (Group Head
  Office, Minamas Head Office)`. A record that does not say what it covered is unreadable a month
  later, and someone will compare a two-segment run with a five-segment one and conclude something
  broke.
- **The guided flows pass their own segment through** — a flow already knows which segment the admin
  is working on, so it pre-ticks that one and leaves the rest tickable.

## 6. Failure rules

| Condition | Behaviour |
|---|---|
| Any detection read fails | that segment shows `could not check`, stays tickable, and **counts as changed** for `Only the ones with changes` — fail towards doing the work. |
| The term store is unreachable entirely | every segment unmarked, detection reported unavailable, Run still offered with all ticked. |
| Zero segments ticked | Run refused, reason stated. |
| A ticked segment has lost its mode row by the time Run is pressed | reported per segment; the rest continue. |

## 7. Why not a "last reconciled" stamp per segment

Rejected. A stamp records that a run *happened*, not that its result is still true — and the failures
that matter happen entirely outside it. It would also need a column, and **a stored flag that governs
whether work is offered can be true before it is true**. Derived, never flagged: the same rule as the
provisioned-path gate, for the same reason.

## 8. Files

| File | Change |
|---|---|
| `src/shared/reconScope.ts` | **new, pure** — `detectSegmentChanges(...)` → per segment `changed` / `clean` / `unknown`, with reasons |
| `src/shared/reconScope.test.ts` | **new** — every verdict, and that an unreadable input yields `unknown`, never `clean` |
| `src/webparts/folderManager/components/FolderManager.tsx` | the picker above Run; filter the mode loop; scope in the log and the audit row |
| `src/webparts/folderManager/components/FolderAdmin.tsx` | pre-tick the flow's segment |
