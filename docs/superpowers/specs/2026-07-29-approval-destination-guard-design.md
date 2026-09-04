# Approval Destination Guard — Design

**Date:** 2026-07-29
**Status:** Agreed, implementing
**Touches:** `ApprovalDocument.tsx`

---

## Problem

The Auto-route flow's `Create new folder` action creates whatever path it is given. If the
**Documents unit folder does not exist**, the flow does not fail — it creates the entire
chain from the library root: `Segment / Department / Unit / Year / Document Type`.

Those flow-created folders **inherit the Documents library's permissions**, where
`DMS_SITE_MEMBERS` holds Read (site-entry spec §3). So every DMS user across every segment
can read that unit's approved documents.

It **fails open, silently**: the flow run shows success, nothing is logged as an error, and
the exposure is only visible by inspecting the folder's permissions. This is the opposite of
how reconciliation behaves — it breaks inheritance *before* granting anything, so a
misconfigured folder there is invisible rather than exposed.

Trigger in practice: a term is added, someone uploads and an approver approves, all before
the next reconciliation run creates and locks that unit's folders.

## Guard

Before approving, verify the destination is safe. Approval is **blocked** unless the
Documents unit folder **exists AND has unique permissions**.

### Deriving the unit folder

From the Staging file path, the unit folder is three levels up — the layout is always
`…/Unit/Year/Document Type/file.ext`, because the upload form ensure-creates Year and
Document Type on the way in. Walking up by position is depth-independent, so it holds for
segments whose Levels chain is deeper than Department → Unit.

```
/sites/Example/Staging/Segment/Dept/Unit/2026/Agreement/file.pdf
  drop filename        → …/Unit/2026/Agreement
  drop Document Type   → …/Unit/2026
  drop Year            → …/Unit                    ← the unit folder
```

Then swap the library segment, anchored on the web-relative prefix so a folder named
"Staging" deeper in the tree is not mangled — the same rule `toDocumentsPath` already uses
in `BulkUpload.tsx`:

```
/sites/Example/Staging/…  →  /sites/Example/Shared Documents/…
```

### Outcomes

| Check | Result |
|---|---|
| Folder exists, permissions unique | **Approve proceeds** |
| Folder confirmed missing (404) | **Blocked** — reconciliation has not created it |
| Folder exists but still inherits | **Blocked** — created by the flow or by hand; readable by all DMS users |
| Could not determine (403 / 429 / 5xx) | **Blocked** — never approve on an unverified destination |

Blocking on an uncertain answer is deliberate. A retry costs an approver seconds; a wrong
"proceed" publishes documents to everyone and nobody finds out. This is the opposite call to
the one the folder-map prune makes — there, uncertainty means *don't delete*; here it means
*don't publish*. Both resolve toward the outcome that cannot leak data.

### Message

> **This unit's folder is not ready in the Documents library, so the document was not
> approved.**
> Ask an administrator to run Folder Reconciliation, then approve again.

Names the fix in the words of the tool the admin has to open. No jargon about inheritance or
serial numbers — the approver cannot act on those.

## Rejection is not guarded

Only `Approved` is checked. Rejecting never copies anything to Documents, so there is no
destination to verify and no reason to block it.

## Known limitation

This guards the **Approval Document web part only**. Approving directly through the
SharePoint list UI bypasses it entirely, and the exposure still occurs. A complete fix would
add the same check inside the Auto-route flow — verify the unit folder before
`Create new folder`, and fail the run loudly instead of creating an unlocked chain.

Recommend doing both. The web part guard catches the normal path and gives the approver a
message they can act on; the flow guard is the backstop that no UI can route around.

## Out of scope

Creating or repairing the folder from the approval screen. The approver may not hold
permission to create folders in Documents, and a folder created there would be exactly the
unlocked one this guard exists to prevent. Reconciliation is the only thing that should
create and lock unit folders.

---

## ⚠ CORRECTION, 2026-08-24 (1.0.233.0) — THE UNIT FOLDER WAS LOCATED WRONGLY, AND IT BLOCKED EVERY APPROVAL

The guard's *rule* below is unchanged and still correct. How it FOUND the unit folder was not, and on
migration day it refused every approval in a segment while reporting "it is not locked down".

It walked **up three levels** from the file, assuming a path of `…/Unit/Year/Document Type/file`. The
comment claimed that "holds for segments with a deeper Levels chain too" — exactly backwards: **a
deeper chain is what breaks it.** A below-Unit tier added through the Structure Manager gives

```
ApprovalDocument / GHO / GCA / EG / 2024 / Agreement / Archive 1 / file.pdf
                   seg   dept  UNIT   ...three up is 2024, not EG
```

and `2024` **inherits the unit's ACL by design**, so `HasUniqueRoleAssignments` read `false` and the
guard refused. One below-Unit tier is wrong the other way: three up lands on the DEPARTMENT, which
does have unique permissions, so it PASSED having checked a folder that says nothing about the unit.

| Below-Unit tiers | three-up lands on | Outcome |
|---|---|---|
| Year, Document Type | the unit | correct — the only shape it was ever right for |
| Year, Document Type, Archive | the **Year** folder | **every approval refused** |
| Year only | the **department** | passes, wrong folder checked |

**THE UNIT FOLDER IS NOW COUNTED FROM THE TOP** — `<segment folder>` plus one level per PERMISSIONED
tier — which is exact whatever hangs below it. Rules in `shared/approvalDestination.ts` (pure, 22
tests), including the regression path above.

- **The tier count comes from the mode row's `Levels`, and CANNOT come from the document's fields.**
  The near-miss looks right and is worth stating: every tier column has a `<Base>Tid` twin, so
  `discoverTierFields` finds them — but `documentDetails.documentUnit` records that a below-Unit tier
  such as SubUnit carries a Tid column *exactly like a permissioned one*. Counting those lands on the
  SubUnit folder, which inherits, and the guard refuses again for a new reason.
- **Keyed on the segment FOLDER name**, which is the mode row's own `StagingFolder`, so no term lookup
  is needed and renaming a segment's label cannot break the match.
- **The mode rows are read in their own effect that awaits `primeNames` itself.** `cachedListTitle`
  answers the legacy `DMS Config` until priming settles — the race that emptied reconciliation's
  segment picker on 2026-08-21. `SortOrder` is deliberately not selected (gotcha #11).
- **Still fails CLOSED, and now names which cause.** An unread `Levels`, an unparseable one, or a
  count of zero all refuse. `undefined` is never read as zero, and a segment whose `Levels` would not
  parse is LEFT OUT of the map rather than stored as `0` — absent means unknown, `0` would claim we
  know the segment is flat.

### The second bug in the same function: the destination library was hardcoded

`DOCUMENTS_URL_SEGMENT` (`"Shared Documents"`) was used unconditionally, while `libSeg()` above it is
HC-aware. **So approving an HC document checked the wrong library's folder** — the same HC-clone
failure as everywhere else in that rollout: one library reference never swapped.

`approvedLibSeg()` now resolves it, and **returns BLANK rather than falling back** when the HC pair is
unresolved, which `unitFolderPath` refuses by name. Unlike `approvedLibTitle()` beside it, whose
fallback to `"Documents"` is harmless for a display string and would be the worst possible thing here:
it would point an HC document's check, and the message an approver acts on, at the open library. Same
reasoning as `hcRouting.ts` returning `undefined`.

Both user-facing messages now name the **actual** destination library, so an approver of an HC
document is not sent to check their access to a library that is working fine.

### A SECOND, SEPARATE GUARD ADDED THE SAME DAY: the file-clash check (1.0.235.0)

The unit-folder guard above answers "is the destination locked down." It says nothing about whether
approving would **overwrite** an existing document — a real gap, found by the client asking directly
whether the counter-check they'd been told about earlier had actually shipped. It had not; it had
only been flagged.

**The race it closes:** `Form.tsx` already refuses an upload whose name clashes with something
already in `Documents` (`approvedClash`, 2026-08-22) — but that check runs once, at upload, before
this document is even in the approval library. It cannot see a name that lands in `Documents`
**after** this upload and **before** this approval (a second uploader, a bulk import). Auto-route's
`Copy file` step was configured to **replace** on a name clash (verified in the flow's own Code
view) — changed to **"Copy with a new name"** on 2026-08-25, so an overwrite is no longer possible
from any route,
so that race ends in a silent overwrite — no error, no warning, a green run.

`documentsFileClash` in `ApprovalDocument.tsx` is the second half: re-checked immediately before
submitting the approval, against the file's own full destination path (library segment swapped,
everything below it — Year/Document Type/Archive… — preserved exactly, mirroring what Auto-route's
`Compose_1` expression itself preserves).

**⚠ FAILS CLOSED, the opposite direction from the upload-time check it complements.** That check
runs on every upload site-wide, so an unanswerable read there would take the whole form out of
service — an acceptable trade to avoid. This check runs once, at one approval, and its neighbour
above (the unit-folder guard) already established the reasoning for fail-closed here: a retry costs
the approver seconds, a wrong "proceed" overwrites a document and nobody finds out.

**A 404 on the below-Unit folder is read as CLEAN, not inconclusive — and that is not a relaxation
of the fail-closed rule, it is a different question.** The unit-folder guard's 404 means "not ready,
refuse." This check's 404 means the below-Unit folder (ensure-created on demand by Auto-route's own
`Create new folder` step) simply does not exist yet — the normal state before the first approval into
a given Year/Document Type/Archive combination. Absent folder and absent file mean the same thing
here: nothing to clash with.

### Interim workaround, while a site is on an older build

**Approve from the library view** (`Approve/reject Items`, or the item's ⋯ → Approve/Reject). That
sets the moderation status directly, Auto-route behaves identically, and this pre-flight check is not
involved. The guard lives only in the web part.
