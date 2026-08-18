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

### 1. The bulk run duplicates Group Map rows — ~~THE BUG~~ **FIXED and SITE-VERIFIED 2026-08-18**

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

**Verified on dcistaging, 2026-08-18.** First run: 0 created (all 308 groups already existed), **790
mappings written**, 0 failed. Second press: **0 created, 0 written, 790 already there** — and it
returned instantly, which is the other half of the proof, since a run that writes nothing makes no
requests at all. `CRS Group Map` held at 790. The log also confirmed the per-persona role counts:
3 for `_SEGVIEW` (SEGVIEW, DEL, SHARE), 1 for `_HOD` (DEPTVIEW), and 13 per unit
(APR/DELS/DEL/SHARE/UPLHC/DELSHC = 6, UPL/DELS = 2, UPLHC/DELSHC = 2, MEMBER/MEMBERHC = 2, MEMBER = 1).

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

## Reconciliation HAS RUN on dcistaging (2026-08-18)

First run completed, then a clean re-run confirmed it: **0 folders created, all folders "already
there"**, ~2,706 steps, about 40 minutes. **No `SKIPPED (no abbreviation)` lines and no collision
aborts** — the GHO codes are complete and unique, which is the pair of failures that run had to rule
out.

Everything else verified from the log: the ten admin pages all report **administrative-only, already
locked**; `CRS_SITE_MEMBERS` holds Read on `Documents` and **nothing** on the other three libraries;
departmental fan-out ON; ancestor-browse Read granting; both approval libraries have content approval
on and provisioned folders are stamped Approved.

### One real item outstanding

```
HCApprovalDocument: no "CRS Folder" or "DMS Folder" content type
HCDocuments:        no "CRS Folder" or "DMS Folder" content type
```

The HC pair never had the content type added, so their folders keep the built-in `Folder` type and the
details pane will not show **Full Name** — the term's real label behind the abbreviated folder name.
Harmless to permissions and to routing; it costs the HC libraries the one thing that makes an
abbreviated tree readable. Add `CRS Folder` to both HC libraries with `Full Name` on it, hide it from
the New button, then re-run and open one HC folder's details pane to confirm the type was re-stamped.

### The `SEGVIEW` warnings were noise, and the wording was the bug

Two `⚠ GHO_SEGVIEW ... not granted` lines per folder in `Documents`, immediately followed by
`↳↓ GHO_SEGVIEW → Read (inherited from a parent-tier mapping)` — the log contradicting itself.

Nothing was wrong: `clevel_segment` is **SEGVIEW + DEL + SHARE**, a segment-tier row fans down for the
**SEGVIEW role alone**, and those two lines were the DEL and SHARE rows being correctly withheld. The
message named the GROUP and never the ROLE, so it read as the whole group being refused, and printed
twice identically. Fixed in **1.0.165.0** — it now names the role and the level and says a SEGVIEW row
on the same group is still applied.

---

## Then, in this order

0. ~~Deploy~~ **Deploy 1.0.165.0** — everything from 1.0.159.0 onward is unreleased — every fix below is code, and none has been run on a site.
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

## ✅ VERIFIED END TO END, 2026-08-18 — the first time on any site

`GHO / GF / TAX`, two gmail guest accounts, one JPG.

1. **Upload** — chocheetuck4 (`_UPLOADER`) opened the form, which auto-detected
   `Group Head Office > Group Finance > Tax` from group membership and locked the tiers. File landed in
   `Approval Document/GHO/GF/TAX/2024/Tax Return`, Pending, with every field.
2. **Approval** — the approval screen showed the preview, the full metadata and *Publish to GHO > GF >
   TAX*. Approved.
3. **Auto-route** — the file moved to `Documents/GHO/GF/TAX/2024/Tax Return` and left the approval
   library. **`Created By` is still chocheetuck4**, and Business Segment, Department, Unit, Year,
   Document Type and Confidentiality all survived the copy. Those are the two steps that had never been
   proven, and both are what the `bNewDocumentUpdate: true` stamp and the column-parity rule exist for.
4. **Notification** — the uploader received "Your file has been approved" with the correct location
   (`GHO > GF > TAX > 2024 > Tax Return`) and a working file link.

Also confirmed in passing: an approver group reaches the upload form and gets its own uploadable path
(the 2026-08-17 "APR can upload" decision), and an uploader is correctly DENIED on
`ApprovalDocument.aspx` (the deliberate asymmetry).

**Still unverified:** anything HC (no content type on the HC pair, and neither HC flow exists), the
deletion and share request flows, and whether a PIC who is not an admin can actually SEE the approved
file in `Documents` — see below.

### One thing to check next

The `Documents` view bar shows a **Show All Files** view, which SharePoint adds when content approval is
on. The routed file is visible to an ADMIN, which proves nothing — admins see drafts. **Have
chocheetuck4 open `Documents/GHO/GF/TAX/2024/Tax Return` and confirm they can see it.** If they cannot,
`Documents` still has moderation on and every routed file is invisible to the unit it belongs to.

---

## ⚠ Renaming an abbreviation after the groups exist

`GHO_GS_Trace` was left mixed-case deliberately on 2026-08-18, to be renamed later as a test of the
rename path. Two things must be known before that test is read as a result:

- **A case-only change renames nothing.** Reconciliation compares codes lower-cased (SharePoint
  sibling names are case-insensitive, so `Trace` → `TRACE` would collide with itself), while the
  abbreviation page compares case-sensitively. So `TRACE` saves, warns `✎ renamed`, and leaves the
  folder called `Trace`. Use a real change — `TRC` — to test the path.
- **The groups do NOT follow**, but since **1.0.162.0** that no longer duplicates anything. The planner
  matches on the TERM GUID, so the unit is recognised under its old group names and the preview reads
  `already there as GHO_GS_Trace_UPLOADER`. Before that build it planned the unit as new, and a run
  created five more groups and thirteen more rows on the same term.

So after renaming that code you MAY rename the five `GHO_GS_Trace_*` groups in SharePoint to match (a
rename preserves the group Id, so every mapping row survives). That is tidy-up, not a repair. Leaving
them is now safe, including through a further bulk run.

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
