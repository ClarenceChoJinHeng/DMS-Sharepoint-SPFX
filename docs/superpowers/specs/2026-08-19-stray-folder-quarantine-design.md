# Stray folders inherit the browse corridor — quarantine them

**Date:** 2026-08-19
**Status:** BUILT 1.0.175.0, NOT site-tested.
**Register:** #19
**Extends** `2026-08-02-term-guid-orphan-repair-design.md` §8, which established that a folder whose
term is gone is reported and **never deleted**. That still stands. This decides what happens to its
**permissions**, which that spec did not address.

---

## 1. The bug, found live

**2026-08-18:** an HC-cleared uploader in one unit could see a folder called `COSEC` — a unit they
had no mapping to — in `HC Approval Document`. Every permission read back correct. No grant was wrong.

`COSEC` was a folder with **no live term behind it**, left over from a deleted or re-coded term. It
had no unique permissions, so it inherited from the **segment folder** above it. And the segment
folder deliberately grants **Read to every group in the segment** — that is the ancestor-browse
corridor which lets a person navigate down to their own unit (restored 2026-08-07; without it a new
library renders empty for every non-admin).

So: **a folder nobody was granted, readable by everyone in the segment.** Not a permissions failure —
a folder the system did not know about, sitting inside a corridor built for the folders it does.

Reconciliation already detects these, and `FolderManager.tsx:4203` states the problem in its own
words:

> `⚠ NO TERM: … no live term maps to this folder, but its permissions are unchanged and its
> documents are still reachable. Review it; nothing was deleted.`

**"Permissions are unchanged" is the finding.** Unchanged means still inside the corridor.

## 2. Why this recurs, and why the rebuild did not fix it

Deleting the folders and re-running reconciliation cleared the existing strays — verified 2026-08-19,
`COSEC` did not return. That cleaned the site; it did not change the mechanism. A stray appears again
whenever:

- **a term is deleted from the store** — the folder survives, the term does not (which is exactly why
  the client is told to rename terms, never delete and re-add);
- **an abbreviation is re-coded** — the next run renames the folder, but a run stopped part-way, or a
  folder renamed by hand, leaves the old name behind;
- **someone creates a folder by hand** in the library, which reconciliation has never known about.

All three are ordinary administration, not misuse.

## 3. The change: quarantine, don't merely report

For each unclaimed folder, reconciliation:

1. **breaks its inheritance** — `copyRoleAssignments=false`, as everywhere else in this file. With
   `true` every inherited grant is carried forward, so the folder stays readable by exactly the same
   people and the run reports success: a failure invisible from the log.
2. **restores site Owners → Full Control**, so the folder does not become unreachable by the people
   who have to decide what to do with it.
3. **reports it by name, with its document count.**

```
⚠ QUARANTINED: HC Approval Document/GHO/GCA/COSEC — no live term maps to this folder.
  Inheritance broken; only site owners can open it now. IT CONTAINS 14 DOCUMENT(S) —
  move them somewhere real before deleting the folder.
```

### 3.1 Nothing is deleted, ever

Unchanged from 2026-08-02 and worth restating: the folder may hold the only copy of real documents,
and **deleting a term revokes nobody's access** — the term store and the documents are independent.
Quarantine removes the accidental *audience*, never the content.

### 3.2 The document count is load-bearing, not decoration

An empty stray is a ten-second tidy-up. A stray holding documents is a small migration. The
difference decides what the administrator does next, so it belongs in the same line — a count on
another screen is a count nobody reads.

**Unknown ≠ zero.** If the contents cannot be counted the message says the count is unknown and
**still quarantines**: the exposure is real whether or not the count could be read. Reporting an
uncountable folder as empty would invite someone to delete it.

### 3.3 Idempotent by construction

A folder already quarantined has unique permissions, so a later run finds nothing to break and
re-reports it as `already quarantined`. No state is stored anywhere — the folder's own ACL is the
record. Same principle as every other assertion pass here.

## 4. ⚠ What this cannot reach, and why that is deliberate

The unclaimed-folder pass **descends only into non-leaf targets** (`descendFrom`, around
`FolderManager.tsx:4190`), for a good reason: below a leaf sit the on-demand `Year` × `Document Type`
folders, which have **no term by design**. Descending there would report hundreds of false positives.

**Consequence: a stray directly under a UNIT folder is undetectable, and must stay that way.** It is
indistinguishable from a legitimate on-demand folder — both are named from a sanitized label, neither
has a term or a Folder Map row. Quarantining on that basis would break the upload form's own
ensure-created folders, which is a far worse failure than the one being fixed.

State the limit in the log rather than leaving it implied:

```
Unclaimed-folder check: segment and department levels only. A folder created directly inside a
unit cannot be told apart from the Year / Document Type folders the upload form creates.
```

Silence here is how "reconciliation checks for stray folders" comes to be believed to mean *all*
stray folders.

## 5. Failure rules

Consistent with the rest of the file: **fail open on the read, closed per folder on the write.**

| Condition | Behaviour |
|---|---|
| A folder listing fails | that branch is **not** checked, reported as today (`could not list N folder(s)`). Never treated as "no strays here". |
| `breakroleinheritance` fails | reported as **STILL EXPOSED**, with the HTTP status — never as quarantined. Same rule as the admin page lockdown. |
| Owners cannot be restored | reported loudly: the folder is now reachable only by site collection administrators. Recoverable, but must never be silent. |
| The document count fails | quarantine proceeds; the count reads `unknown`. |
| No `Read`/`Full Control` role definition on the site | the pass is skipped and says so, as the other passes do. |

## 6. Why quarantine rather than a louder warning, when this codebase fails open

Most reads here fail open because the cost is a form out of service for a minute. **The cost of the
open direction here is a document readable by dozens of people who were never granted it**, for as
long as it takes someone to notice one line among four thousand. The same reasoning that made
`canOfferFolderDelete` and `hcRouting` fail **closed** applies: where the failure mode is exposure
rather than inconvenience, the safe direction inverts.

The counter-argument, considered and rejected: quarantine could hide a folder somebody is actively
using. It is answered by the folder being **named in the log with its document count**, by nothing
being deleted, and by inheritance being restorable in two clicks. An administrator who loses a folder
they were using will come looking; a group quietly reading another unit's documents will not.

## 7. What it does not fix

- **It cannot say which unit the documents belonged to.** The term is gone and that information left
  with it. The orphan-repair pass re-points a folder whose term was re-created only on a **1:1 match
  of level + label**, and ambiguous matches are still reported rather than guessed — `Tax`, `Legal`
  and `PM` each exist under several departments, and a wrong re-point is a grant to another
  department's folder.
- **It does not stop strays appearing.** "Rename terms, never delete and re-add" remains the
  guidance, and remains unenforceable from here.
- **It does not reach below a unit** — see §4.

## 8. Files

| File | Change |
|---|---|
| `src/webparts/folderManager/components/FolderManager.tsx` (~4175-4215) | quarantine inside the existing unclaimed-folder pass; document count; the scope caveat line |
| `src/shared/` | none — the rule is "no term ⇒ break inheritance", with no branching worth extracting |
