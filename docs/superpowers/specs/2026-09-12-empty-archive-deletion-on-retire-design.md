# Delete an empty Archive on segment retire (2026-09-12)

## Where this came from

Live incident on PCAR (see CLAUDE.md's PCAR troubleshooting log, 2026-09-11/12). Crystal retired an
earlier "Project Cars" with "also delete the folders" ticked. That deleted the folder tree in
`Staging`/`Documents`/`StagingHC`/`DocumentsHC` — but **deliberately left `Archive`/`ArchiveHC`
untouched**, per the 2026-08-26 client decision that 7-year retained records outlive the segment
that produced them.

Clarence recreated the segment later, reusing the same `StagingFolder` code and term set. Because
the segment's **top-level folder has no term behind it** (unlike departments/units, which are
matched by term GUID via Folder Map), reconciliation found the surviving `PCAR` container in
Archive and **reused it** rather than creating a duplicate. But the department/unit folders one
level down — `EB`/`Volvo`/`Honda`, tied to the OLD terms — had no way to reconnect to the NEW term
tree, since their Folder Map rows died along with the recycled Staging folders. They sat there as
unresolvable strays, which blocked `SubtreeMigrator`'s "Move existing folders" step a day later,
for reasons that took most of a session to trace back to this cause.

A warning was added to the retire dialog the same day (`segmentDeletion.ts`,
`survivorLines`) telling the admin this will happen. This spec is the next step: **can the Archive
folders be deleted too, when doing so is provably safe** — closing the gap rather than just warning
about it.

## The proposed rule

When "also delete the folders" is used on retire, **also delete the segment's Archive/ArchiveHC
folder tree — but only when a fresh count confirms it holds zero documents.**

This is not a relaxation of the 2026-08-26 decision. That decision is about **documents**, not
empty containers — "7-year retained records outlive the segment" says nothing about a folder that
was never used, or one whose contents already moved on. Deleting an empty shell protects nothing;
leaving it behind is what created this incident.

## Why the existing document count can't just be reused

`SegmentCounts.documents` already exists and already gates whether "also delete the folders" can be
offered at all (`canOfferFolderDelete`). It is **not** safe to fold Archive into that same number:

- `documents` today counts only the libraries retire actually deletes from
  (`Staging`/`Documents`/`StagingHC`/`DocumentsHC`, via `retireLibraries()`). If Archive's own
  document count were merged into that same total, a segment with real archived records but an
  otherwise empty operational tree would show `documents > 0` and **lose the existing, working
  ability to clean up the empty operational folders while leaving Archive alone** — which is the
  common, correct case today (a segment fully wound down, its records already archived). That would
  be a regression in the middle of fixing something else.
- The two questions are genuinely different: *"can I delete Staging/Documents at all"* and *"can I
  ALSO delete Archive"* are separate decisions with separate consequences, and must stay separate
  fields.

So: a **new, independent** count is needed, specific to Archive/ArchiveHC.

## New shape

```ts
export interface SegmentCounts {
  // ...existing fields unchanged...

  /**
   * Documents in Archive/ArchiveHC for this segment, counted SEPARATELY from `documents`.
   * `undefined` means the count could not be established — never assume empty from a failed read.
   */
  archiveDocuments?: number;
}
```

- **Absent/`undefined`** — Archive wasn't counted, or the count failed. Treated identically to
  "Archive has documents" for the purpose of deletion: **fail closed.** Guessing empty on a failed
  read is exactly the "an archive nobody could confirm was empty being deleted" scenario the original
  delete-segment spec already refuses, for the same reason.
- **`0`** — confirmed empty. Only this value permits deleting Archive's tree.
- **`> 0`** — has content. Archive is left exactly as it is today; the existing warning line applies.

### Archive and ArchiveHC are one combined check, not two independent ones

Matching the project's existing "both-or-neither" treatment of the HC pair everywhere else (the
archive mirrors the site's HC state; `hcRouting.ts` fails closed the same way): if the site has both
`Archive` and `ArchiveHC`, **both must independently confirm zero documents** before either is
deleted. If one has content, or one's count fails, **neither is deleted.** Partially deleting the
pair — removing the plain archive while leaving an orphaned HC one, or the reverse — creates a new,
asymmetric version of the exact problem this spec exists to close.

## Where the counting happens

`segmentDeletion.ts` stays pure and SPFx-free, per its own stated design (same split as
`newSegment.ts`/`segmentReadiness.ts`) — it holds the **decision** logic, never the REST calls. The
actual count is a new async read added alongside the existing `countSegment` in
`SegmentCreator.tsx` (or a sibling function it calls), walking `Archive`/`ArchiveHC` the same way
the existing count walks the other four libraries, and reporting `archiveDocuments` back into
`SegmentCounts`.

## Dialog behaviour

The "also delete the folders" checkbox stays a single control — **no third checkbox.** Its
consequence text picks one of three lines depending on `archiveDocuments`:

1. **`0` (confirmed empty):** *"The archive folders are also empty, so they will be deleted too."*
   Positive confirmation — nothing is being silently swept in, the admin is told the segment's
   archive footprint is genuinely nothing.
2. **`> 0`:** the existing line from 2026-09-12 — *"The archive folders are NOT deleted — archived
   records outlive the segment. If this segment (or a new one using the same code) is ever
   recreated, these will not be recognised and will block 'Move existing folders' until someone
   finds and removes them by hand."*
3. **`undefined` (could not be counted):** a THIRD line, distinct from both — *"The archive folders'
   contents could not be confirmed, so they are being left in place rather than guessed at."* This
   must not reuse either of the other two wordings: saying "not deleted, records outlive the
   segment" implies a confirmed answer that was never reached, and staying silent would look like
   the check was simply skipped.

`deletionSummary` and the audit row both need the same three-way split, for the same reason the
dialog does — an admin reading the audit log afterwards needs to know which of the three actually
happened, not just "folders were deleted: yes/no."

## Non-goals

- **No retroactive cleanup.** This does not touch PCAR's existing strays, or any other segment
  already left in this state by a past retirement. It only changes what happens on a **future**
  retire. PCAR's leftovers still need manual cleanup, as already planned.
- **No standalone "clean up archive" tool.** A prior one-off admin button for a related archive
  cleanup was built and then explicitly removed at the client's request ("we don't want to confuse
  the client with a new feature like this" — 2026-09-02). This capability lives entirely inside the
  existing retire flow's existing checkbox; it introduces no new page, button, or admin-facing
  control.
- **Does not change the 2026-09-11 abbreviation-deletion rule.** Abbreviation rows are still deleted
  unconditionally on retire, independent of anything in this spec.
- **Does not change the segment-create-time check** discussed alongside this (a heads-up when
  creating a segment whose code already has leftover Archive folders from an earlier retirement).
  That is a separate, independent piece of work, not a prerequisite for this one — either could ship
  without the other, since one prevents leftovers being created and the other warns when leftovers
  are about to be reused.

## Open question for the client / whoever owns this decision

Confirm this reading of the 2026-08-26 decision is acceptable: **"records outlive the segment"
means documents outlive it, not empty folders.** If the intent was closer to "never touch anything
under Archive/ArchiveHC for a retired segment, structure included, regardless of content," this
spec's whole premise doesn't hold and the fix should stay at the warning-only stage already shipped.
