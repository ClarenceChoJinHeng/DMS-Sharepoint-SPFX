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

Deployed there: **1.0.150.0**. Latest commit `d0490a6` on `feat/folder-abbreviations`, 1050 tests, 19
warnings (the baseline).

---

## Open defects, in build order

### 1. The bulk run duplicates Group Map rows — THE BUG

`BulkGroupProvisioner.rowsFor` writes rows **unconditionally**. The run is idempotent for *groups* (an
existing title is mapped, not re-created) and not for *rows*, so a second press writes every mapping
again. That is where 642 rows came from.

**Fix:** read the segment's existing rows and filter with `isDuplicateRow` from `groupMapModel.ts` —
`GroupMapBuilder.tsx` already does exactly this. Re-read after the run so a second press sees them.

**A single run against an empty Group Map is clean**, which is why the list was cleared rather than
deduplicated by hand.

### 2. Navigation during a run kills it, silently

The run lives in component state with no resume, so changing step, pressing Next/Back or closing the tab
stops it part-way with no warning. This is the likely cause of 302 of 308.

**Fix:** surface `busy` to the host so Next, Back and the rail are disabled, plus a `beforeunload` guard.
Same reasoning as the staged-batch guard in `BulkUpload.tsx`.

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
