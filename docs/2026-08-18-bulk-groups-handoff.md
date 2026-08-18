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
| SharePoint groups | **~302 of 308** — the run did not finish |
| `CRS Group Map` | **cleared** — was 642 rows, with duplicates |
| `CRS Folder Map` | empty |
| Reconciliation | **NOT RUN. Hold until the row count looks right.** |

dcistaging's tree is NOT the client's: it lacks Group Corporate Secretarial and one Group Finance unit.
So **308 is correct there, and CRS should plan 324** (63 × 5 + 8 + 1). A different number on CRS is a
real gap, not a repeat of this.

Deployed there: **1.0.150.0**; `feat/folder-abbreviations` is now at **1.0.152.0** (defects 1-2), 1056
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

## Open defects, in build order

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

### 3. The mappings table shows one bare term

Folder Access shows `Tier` as a single label, so a unit row does not say which department it is in, and a
department row is indistinguishable from a unit one. Client asked for **Tier 1 / Tier 2** columns.

A row stores only the **leaf** term GUID, so the department must be derived by walking the segment's term
tree — the walk `BulkGroupProvisioner` already performs. Build `guid → {tier1, tier2}` once per segment.

### 4. The group list is not scrollable

303 groups stretch the page; the list needs its own `max-height` and overflow. Same class of problem as
defect 2 — a 300-item run and a 300-item list were both designed as if they would be small.

### 5. Step 4's tick does not refresh after a run

`groupsExist` is read when the segment is picked, so the step still says *To do* once the groups exist.
Cosmetic, but it makes the rail lie.

---

## Then, in this order

1. One bulk run → expect ~6 created, ~790 rows, **0 failed** (the 302 existing groups log as `=`)
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
