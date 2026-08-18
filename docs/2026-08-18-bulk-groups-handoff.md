# Bulk group provisioning — state and open defects

**Date:** 2026-08-18. **Read this before touching group provisioning.**
Spec: `docs/superpowers/specs/2026-08-18-group-creation-and-bulk-provisioning-design.md`

---

## Where the rehearsal site stands

**dcistaging** = `dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting`. It has its **own
site-collection app catalog**, which takes precedence over the tenant one — several rounds were lost to
uploading a package to the wrong catalog and diagnosing it as caching. Check the *installed* version in
Site Contents, not the catalog listing.

| | State |
|---|---|
| Term store, GHO | **7 departments, 60 units** (verified by console read) |
| Abbreviation codes | complete for GHO |
| `mode_group_head_office` | created |
| SharePoint groups | **308 of 308** — the preview reported 0 to create on 2026-08-18 |
| `CRS Group Map` | **790 rows**, written by one clean run on 2026-08-18 (was cleared; 642 before that, with duplicates) |
| `CRS Folder Map` | empty |
| Reconciliation | **NOT RUN. Hold until the row count looks right.** |

dcistaging's tree is NOT the client's: it lacks Group Corporate Secretarial and one Group Finance unit.
So **308 is correct there, and CRS should plan 324** (63 × 5 + 8 + 1). A different number on CRS is a
real gap, not a repeat of this.

Deployed there: **1.0.150.0**; `feat/folder-abbreviations` is now at **1.0.156.0** (all five defects, plus two more found on site), 1065
tests, 15 warnings (the baseline).

---

## ⚠ The CLIENT site was also reset on 2026-08-18

Separately from the rehearsal, on **`sdguthrie.sharepoint.com/sites/CRS`** the client emptied
**`CRS Folder Map`, `CRS Group Map` and `CRS Term Abbreviation`**, and deleted the handful of groups
created by hand. So the migration runbook's progress is **stale from §10 onwards**: the 71 GHO
abbreviation codes recorded as saved there were cleared and re-entered, and any group created before that
point is gone.

What still stands on CRS: the four libraries and their columns, the three term sets, the permission
levels, the config rows, and the three `mode` rows. What does not: abbreviations, group map rows, folder
map rows, groups.

Do not read `docs/2026-08-17-sdg-migration-runbook.md` §10–§12 as done. Re-check the lists.

---

## Open defects, in build order — ALL FIVE FIXED 2026-08-18, NONE SITE-TESTED

### 1. The bulk run duplicates Group Map rows — ~~THE BUG~~ **FIXED, 1.0.151.0, not yet site-tested**

`BulkGroupProvisioner.rowsFor` wrote rows **unconditionally**. The run was idempotent for *groups* (an
existing title is mapped, not re-created) and not for *rows*, so a second press wrote every mapping
again. That is where 642 rows came from — and the log read as safe while it happened, every line saying
`= already existed (mapping only)` about the group whose rows were being doubled underneath it.

**Fixed** by `splitPlannedRows` in `shared/bulkGroups.ts` (pure, 6 tests), reusing `isDuplicateRow`
rather than re-deriving what "the same mapping" means — a second definition would be free to drift from
the one an admin sees on Folder Access. The run now reads every Group Map row first, partitions each
group's planned rows into fresh and duplicate, writes only the fresh ones, and grows its local copy as
it goes (so a row is never doubled within one run either).

Three things worth keeping:
- **The partition is per ROW, not per group.** A run stopped part-way (defect 2) leaves a group made
  with one of its two rows written; skipping the whole group as "done" would strand it permanently.
  A press after an interrupted run now completes exactly what is missing.
- **The read is PAGED.** `$top` caps a page, it does not lift the 5,000-item threshold, and one
  segment on CRS is ~790 rows. A truncated read reports the rows it could not see as absent — the
  same bug wearing the fix's clothes.
- **It fails CLOSED.** An unreadable Group Map is `undefined`, never `[]`, and holds the Run button
  with the reason on screen. This codebase fails open nearly everywhere; there the cost is a form out
  of service for a minute, here it is hundreds of duplicate rows that no screen will ever show you.

The button now also states the row count it will leave alone, so "safe to press twice" is visible
before the press rather than discovered after it.

**A single run against an empty Group Map is clean**, which is why the list was cleared rather than
deduplicated by hand.

### 2. Navigation during a run kills it, silently — **FIXED, 1.0.152.0, not yet site-tested**

The run lives in component state with no resume, so changing step, pressing Next/Back or closing the tab
stopped it part-way with no warning. This is the likely cause of 302 of 308.

**Fixed:** `BulkGroupProvisioner` reports `busy` up through a new `onBusyChange` prop (the same
report-upward shape as `onAbbreviationsMissingChange`), and `FolderAdmin` holds the rail, Back, Next,
Finish, the mode switch and the *Back to Folder Management* band for the duration, with the reason
beside the greyed buttons. The `beforeunload` guard lives in the component, so the standalone Group
Management page is covered too.

- **This is the one place in the runner that padlocks navigation, and the exception is argued rather
  than assumed.** The flow's rule is that it must not stop an admin doing the work; here navigation
  destroys work already in flight, which is the same reasoning that makes a tab switch with unsaved
  abbreviations a refusal rather than a "discard?" prompt. It is temporary and self-clearing.
- **A Stop button was added with it**, because holding every exit for the length of a 300-group run
  and offering no way out would be half a feature. It stops **between groups**, never inside one, and
  a stopped run reports as a WARNING however cleanly it stopped — a green toast over a
  half-provisioned segment is how 302 of 308 goes unnoticed a second time.
- **Stopping is only safe because of defect 1.** Pressing Run again finishes what is missing and
  re-writes nothing, so an interrupted run is now a pause rather than a mess.
- The mode switch is held too: *Create one group* unmounts the provisioner exactly as a step change
  does.

### 3. The mappings table shows one bare term — **FIXED, 1.0.153.0, not yet site-tested**

Folder Access showed `Tier` as a single label, so a unit row did not say which department it was in, and a
department row was indistinguishable from a unit one. Client asked for **Tier 1 / Tier 2** columns.

**Fixed:** `shared/termChains.ts` (pure, 7 tests) turns a walked tree into `guid → [labels]`, and
`GroupMapBuilder` walks each distinct segment **once**, bounded by its permissioned depth.

- **It replaces a per-term read and is cheaper than what it replaces** — ~130 requests for a
  provisioned segment against ~8 for the walk (one per department, plus the top level).
- **The depth bound is load-bearing.** SubUnit terms are authored *under each unit*, so an unbounded
  walk would fetch every SubUnit on the site to answer a question about departments. `parseLevels`
  keeps non-permissioned entries, so the `permissioned !== false` filter is what sets the bound.
- **A branch that could not be read abandons the whole segment**, which then reads *Tier not known*.
  Reporting it as a term with no children would present a unit as a department — the exact wrong
  statement these columns exist to prevent. Unknown ≠ empty, again.
- **A chain of one means "this row is on a department"; unresolved means "we do not know".** They
  render differently on purpose, and a department row's blank Tier 2 shows as a dash, because that
  emptiness is the fact identifying it.
- **The CSV split with it** (`Tier 1`, `Tier 2` headers) from the same function as the screen — the
  export is the client's cross-check, and two derivations of one value drift.

Segments with no `mode` row (Upstream Malaysia's placeholder) resolve nothing and say so; those rows
already carry the `stale` badge beside them.

### 4. The group list is not scrollable — **FIXED, 1.0.154.0, not yet site-tested**

303 groups stretched the page, pushing the create form off the top. The list now caps at `60vh` and
scrolls — **but only while every group is collapsed.** An expanded group's member editor carries an
absolutely-positioned people picker, and a scroll container clips it for any group near the bottom:
that trades a long page for a control that silently cannot be used. With one group open the admin is
working inside it rather than scanning the list, so the cap has nothing to do.

### 5c. Next was available while the term store was still being read — **FIXED, 1.0.156.0** (found on site, 2026-08-18)

The abbreviations step showed *"Reading the term store…"* with **Next step** enabled beside it. The
count is `undefined` for the whole read, and `undefined` never gates — by design, because a throttled
list must not strand anyone mid-flow.

**In flight is not unknown**, and that is the same split that made `subjectGiven` right: fail-open
exists for reads that can FAIL, not for reads still running. Gating while loading strands nobody
because it clears itself in seconds; gating on a failed read would strand everybody. So
`abbreviationsLoading` is a **separate fact** from the count — one value cannot tell "not read yet"
from "read and failed", and only the first should hold a button.

Reported up through `onLoadingChange` → `onAbbreviationsLoadingChange` → `FolderAdmin`, the same shape
as the count beside it. Pinned by tests: it gates the abbreviation step **only** (it is set for the
whole flow, so any other step reading it would be held whenever the screen was open), and it stops
gating the moment the read finishes, failure included.

### 5b. The group list still said "not mapped" after a run — **FIXED, 1.0.155.0** (found on site, 2026-08-18)

The first real bulk run wrote all 790 rows and every group in the list beside it still read **not
mapped**. `GroupManager` loads its mapping badges at mount and nothing told it to look again — defect 5
one screen over, and the more dangerous of the two: the rail merely under-reports progress, while this
says the run did not do what it just did, which invites exactly the needless second press.

`BulkGroupProvisioner` now fires `onRunComplete` after its own post-run re-read, and both hosts pass a
`refreshKey` down to `GroupManager`. **A separate signal from `busy` going false**, deliberately: "the
run ended" is the fact that makes other screens stale, and deriving it from the falling edge of a UI
flag ties two unrelated things together.

### 5. Step 4's tick does not refresh after a run — **FIXED, 1.0.154.0, not yet site-tested**

`groupsExist` was read when the segment was picked, so the step still said *To do* after a run created
300 groups. `onRunBusyChange` bumps `reload` on the busy→idle edge and the facts effect keys on it.
Cosmetic in that nothing underneath was wrong; not cosmetic in the way that counts — the rail is what
tells an admin what is left, and one entry known to be lying is enough to stop them trusting the rest.

Found alongside it: the groups read was `$top=500`, and a provisioned segment is ~324 groups on its
own. Now 5000 — a truncated read would have reported an existing segment's groups as absent.

---

## Then, in this order

0. **Deploy 1.0.154.0 to dcistaging first** — every fix below is code, and none has been run on a site.
   Check the INSTALLED version in Site Contents, and remember dcistaging's own site-collection catalog
   takes precedence over the tenant one.
1. ~~One bulk run~~ **DONE 2026-08-18 on dcistaging: 0 created, 790 written, 0 failed** — all 308
   groups already existed, so it was a mapping-only run, and 790 is the exact expected count
   (60 units × 13 rows + 7 departments × 1 + 1 segment × 3). Navigation was correctly held throughout.
   **The second press, which is what actually proves the dedupe, has NOT been done**: expect
   0 created, 0 written, 790 already there. Anything approaching 1,580 rows means stop.
2. Spot-check `CRS Group Map`: a `_HOD` row must carry the **department** term, a unit row the leaf
3. Folder Reconciliation — first time at this scale, ~1 hour, single tab, no resume
4. Two-account verification: upload → approve → route

**Nothing in this system has been verified end to end yet, on any site.** That is still the real
outstanding risk; everything above is provisioning.

---

## Judgement calls worth not re-litigating

- **Five personas per unit** (`_UPLOADER`, `_APPROVER`, `_EMPLOYEE`, `_UPL_HIGHLY_CONFIDENTIAL`,
  `_VIEWER_HIGHLY_CONFIDENTIAL`) plus `_HOD` per department and `_SEGVIEW` per segment. A person is in
  exactly ONE — `UPLHC` and `MEMBERHC` are supersets covering the normal libraries too.
- **`_EMPLOYEE` and `_HOD` are new suffixes** (2026-08-18). HOU is Head of *Unit*; a department group
  must not borrow it.
- **The name is derived and read-only**, with free text behind an advanced toggle, because Folder Access
  recovers a role by parsing the suffix and a hand-typed name pre-selects the wrong one.
- **Creating a group writes its Group Map rows**, which takes that parse off the critical path.
- **A coded unit under an uncoded department is skipped**, because reconciliation builds
  `GHO/<dept>/<unit>` and its group would grant nothing.
