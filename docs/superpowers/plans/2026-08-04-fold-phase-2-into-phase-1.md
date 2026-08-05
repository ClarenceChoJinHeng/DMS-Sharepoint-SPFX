# Fold Phase 2 (Highly Confidential) into Phase 1 — Plan

**Date:** 2026-08-04
**Status:** DRAFT — for review before any code moves.
**Trigger:** client asked on 2026-08-04 for Phase 2 to ship inside Phase 1.
**Source design:** `2026-07-16-highly-confidential-securing-design.md` — **the 556-line version on
`feat/hc-libraries`**, not the 88-line stub currently on `feat/folder-abbreviations`.

---

## 1. The headline: do not merge the branch

`feat/hc-libraries` diverged before roughly two weeks of work. `git diff HEAD feat/hc-libraries`
reports **37 files, 2,855 insertions, 5,396 deletions** — and the deletions are things Phase 1
gained afterwards:

| Missing from the HC branch | Consequence of merging |
| --- | --- |
| `folderAbbreviation.ts` + tests, `folderFullName.ts`, `fileSize.ts` | abbreviation-named folders regress to term-label names |
| All four `2026-08-03-*` specs | the agreed role model disappears from the repo |
| ~1,280 lines of `FolderManager.tsx` | site-entry pass, library pass, revoke pass, `DELS`, C-Level fan-down |
| ~440 lines of `GroupMapBuilder.tsx` | persona picker, custom dropdown, scope blocking, delete modal |
| ~930 lines of `Form.tsx`, ~1,950 of `BulkUpload.tsx` | single-selection bulk upload, allowed-file-types work |

A merge or rebase would fight over every one of those files, and the resolution would be "keep
Phase 1" in almost every hunk. The HC-specific code is a small fraction of that diff.

**So: treat the branch as a reference implementation, not a source to merge.** Re-apply the four
HC commits' *intent* on top of the current branch, file by file, then delete the branch once its
spec has been brought forward. The commits worth reading:

```
8117616  feat: add DEL/HC roles and make the folder map library-aware
5a85cc2  feat: reconcile the two Highly Confidential libraries
7412f16  feat: treat HC as an upload role and filter confidentiality by it
44da504  feat: route Highly Confidential uploads to the HC Approval library
```

## 2. What Phase 1 already has

Some of the HC groundwork landed independently while HC was parked, so the gap is smaller than
the spec implies:

- `HC` is already in `GroupMapRole`, deliberately absent from `SELECTABLE_ROLES`.
- The `pic2` persona already exists, listed and greyed out with its reason shown.
- `DEL` is implemented — re-added on its own merits, not as part of HC.
- The custom permission levels exist on the test site (`DMS Approve`, `DMS Delete`, `DMS Upload`).
- The library rule (`LIBRARY_ROLES`) and departmental fan-out are in place, and HC plugs into both.

## 3. What has to change, and the one structural item

### 3.1 `LibTarget` stops being a two-value union — this is the real work

Reconciliation is written around `LibTarget = "Staging" | "Documents"`. That type drives the
target loop, the library rule, the log tabs and the folder-map rows. HC adds **two** more
libraries, so it becomes four values and every `Record<LibTarget, …>` grows with it.

Schedule this first and alone. It touches the whole reconciliation loop, and a mistake in it is a
folder provisioned into the wrong library with the wrong ACL.

| Role | Staging | Documents | HC Approval | HC Library |
| --- | --- | --- | --- | --- |
| `UPL` | ✅ | | | |
| `APR` | ✅ | | | |
| `DELS` | ✅ | | | |
| `MEMBER` | | ✅ | | |
| `DEL` | | ✅ | | |
| `GLOBAL` | | ✅ | | |
| `HC` | | | ✅ | ✅ |

`HC` reaching **both** HC libraries is the one row that breaks the existing one-role-one-library
pattern, and it is deliberate: an HC uploader submits to HC Approval and reads from HC Library.

### 3.2 The rest

1. **Reconciliation** provisions the folder tree in both HC libraries, with the same abbreviation
   naming, `Full Name` column and content type as Staging/Documents.
2. **Form** filters the Confidentiality dropdown by whether the user holds `HC`, and routes an HC
   upload to HC Approval instead of Staging.
3. **Folder map** becomes library-aware, so a term maps to four folders rather than two.
4. **Auto-route** gains a second route, HC Approval → HC Library, plus the `UploadedBy` handling
   from spec §9.2.
5. **Columns and views** on the two HC libraries, per spec §10.
6. **`SELECTABLE_ROLES`** gains `HC` and `pic2` loses its `unavailable` flag — **last**, so the
   role cannot be provisioned before it works.

## 4. The four restore steps, three of which fail silently

From spec §14, and the reason this plan leads with them:

| # | Step | If missed |
| - | --- | --- |
| 1 | Recreate the **Highly Confidential** term in the Confidentiality Level set | Level cannot be selected — visible immediately |
| 2 | Update `term_highlyConfidential` in DMS Config to the term's **new** GUID | **Silent** — every HC document routes to Staging with no error |
| 3 | Update `DEFAULT_SETTINGS.hcTermGuid` in `Form.tsx` | **Silent**, and only when config is unreadable — the worst kind of stale fallback |
| 4 | Restore the Highly Confidential tooltip text in `Form.tsx` (removed in `7745d2b`) | Cosmetic |

**The term was deleted, so its GUID is dead.** `420d75d5-f3b7-4525-9eb7-ec06590e7f22` appears in
both the old config row and the branch's code fallback and must not be reused. Read the new one:

```
GET <site>/_api/v2.1/termStore/sets/0d6d1da8-27e5-477f-8684-e8cf169f8fb9/terms
```

Steps 2 and 3 are the dangerous pair. The design matches on GUID rather than label precisely so a
term *rename* cannot reroute confidential documents unnoticed — but a stale GUID causes that same
failure by another route. **If HC uploads land in Staging once this ships, check those two before
anything else.**

## 5. Interaction with work done since HC was parked

Four Phase 1 changes post-date the HC design, which has not been reconciled against any of them.
Each needs an explicit decision rather than an assumption:

**Ancestor Read is gone.** HC library folders will have no readable parents, so an HC uploader can
no more browse to their folder than a normal uploader can. The **router web part must cover the HC
libraries too**, or HC is unreachable except by direct link.

**`UPL` no longer carries delete.** Does `HC` — an upload role — carry delete on HC Approval? By
the same argument that produced `DELS`, it should not. Needs confirming: the HC design predates
that decision.

**PIC collapsed to one persona.** `pic2` is the HC persona, defined as `HC` alone. With PIC 1 now
`UPL` alone that is consistent — but confirm the client still wants an HC-only PIC who cannot see
ordinary Confidential documents, which was the point of it.

**The `Scope` column exists.** HC rows are `Scope = Folder` like the rest. No Library-scope rows
for the HC libraries: that would grant across every unit and defeat the isolation.

## 6. Client and admin prerequisites

None of these can be done by us, and 1–3 block everything:

1. **Create the two HC libraries** with no-space URL names (spec §15).
2. **Recreate the Highly Confidential term** and give us its new GUID.
3. **Decide who holds `HC`** — per unit, as with every other role.
4. Confirm **approvers may read HC documents** (spec §15 assumes yes).
5. Confirm **department-tier fan-out applies to `HC`** (spec §15 assumes yes, consistent with the
   other roles).

## 7. Suggested order

1. Bring the **full HC spec forward** onto this branch — documentation only, so no conflicts, and
   it stops the authoritative design living solely on a stale branch. Then delete the branch.
2. Reconcile that spec against §5 and record the four decisions.
3. **`LibTarget` → four libraries** (§3.1), on its own, with the library-rule table as tests.
4. Reconciliation provisions the HC trees.
5. Form: filtering, then routing.
6. Auto-route second route.
7. Un-gate: `HC` into `SELECTABLE_ROLES`, `pic2` un-greyed.

Steps 1–2 are worth doing even if the client changes their mind again: they cost almost nothing
and they stop the design rotting on a branch nobody is on.

## 8. Known conflicts with the current queue

This lands on top of unfinished work. In dependency order:

| Item | Relationship to HC |
| --- | --- |
| Router web part | **HC depends on it** — HC folders have no browsable parents |
| `naming.ts` wiring (DMS→CRS) | Independent, but HC adds two more names to resolve — cheaper before |
| Deletion request flow | Independent |
| Site / Library / Page tabs | Independent |
| Group model rework (paused 2026-08-04) | **Overlaps.** The client says their group list has changed again, and §5 raises four role questions. Settle the group model first, or HC gets built against a model that is already moving |
