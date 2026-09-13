# Re-coding a segment's top folder — physical rename, empty segments only

**2026-09-11.** Client, on the "Rename or re-code a folder" flow: they can rename a Department or a
Unit there, but not the segment itself — *"sometimes client made a mistake on naming the segment and
then they have to delete the entire segment or deal with it."* Confirmed they want the actual folder
physically renamed, not just the display label, after being shown the risk. This was already flagged
as agreed-and-deferred in CLAUDE.md on 2026-09-10 (parked until after the SDG migration); the client
chose to build it now.

## Why there was no box for this already

The segment container folder has **no term** — unlike Department and Unit, it has no abbreviation row
and no term-keyed Folder Map row. It is created **by name**, straight from the mode row's
`StagingFolder` value. Reconciliation resolves and renames Department/Unit folders **by UniqueId**
because a Folder Map row anchors them to a term; a segment folder has nothing anchoring it at all.

So changing `StagingFolder` by itself would:
- make reconciliation build a **brand-new, empty tree** under the new name on its next run, and
- leave the **old tree — documents included — as an unrecognised stray**, because nothing walks it any
  more;
- leave the Folder Map rows for that segment still naming the OLD `Section`, so
  `segmentProvisionState` mismatches and every upload into it is refused;
- leave the approval destination guard resolving the unit folder from the OLD name, so approvals into
  units under it are refused too.

`StagingFolder` is a **key**, referenced across roughly two dozen files, not a display label.

## What is NOT affected, and why the recode is smaller than it sounds

- **Group names are unaffected.** `suggestGroupName` derives a group's stem from the segment's
  **label** (`ModeLabel`), never from `StagingFolder`. Renaming the folder key does not touch group
  names, and group provisioning already matches on the underlying **term GUID**, not the name
  (1.0.162.0's rename-recovery fix), so nothing about group matching cares which folder key a segment
  currently uses.
- **Group Map rows are unaffected.** A row's `Segment` column holds the segment's **term-set GUID**
  and `UnitTermGuid` holds a **term GUID** — neither depends on `StagingFolder` in any way. Every
  existing grant, once reconciliation re-runs, lands on the SAME term-derived folders under the new
  key with no row change needed.
- **Abbreviation rows are unaffected.** They are keyed per-term and hold no reference to the segment's
  top folder at all.

So the only two things that genuinely need touching are the mode row's `StagingFolder` cell and the
Folder Map rows whose `Section` names the OLD value.

## Scope: empty segments only, and refused otherwise

- **Empty and countable** (documents = 0, confirmed by a successful count across every non-archive
  library) ⇒ **allowed**. The old folder tree is recycled, the mode row's `StagingFolder` is updated,
  and the Folder Map rows for the old `Section` are dropped — they are derivable, and a full
  reconciliation run rebuilds them under the new name. Nothing here creates the new folder tree
  itself; reconciliation does that, on its own next run.
- **Holds documents** ⇒ **refused outright, no partial path offered.** Re-coding a segment in use
  would abandon its documents in a folder reconciliation no longer recognises. The message names the
  real route: move or archive the documents first (there is no cross-segment move tool in this
  system, so that is a manual SharePoint step), or retire the segment properly if it is being
  replaced.
- **Uncountable** (a failed read) ⇒ **refused**, fail-closed — the same rule `canOfferFolderDelete`
  already uses for the identical reason: guessing "empty" here strands a document tree, and the
  person who discovers it later is an uploader.

This reuses `SegmentCounts` and `countSegment()`/`retireLibraries()` from the existing delete flow
directly — same count, same libraries, same fail-closed instinct — rather than a second definition
that could drift from it.

## Where it lives

On the **Segments** list (the "New segment" tab's existing "Segments on this site" panel), next to
Delete on each row — never on the abbreviations screen, because a segment is not a term and a field
that can strand a whole document tree does not belong in a column of five fields that simply work.

## The flow

1. Press **Re-code folder** on a segment row → opens a dialog, counts the segment (reusing the exact
   same count the Delete dialog runs).
2. **Counting** → "Counting what this segment holds…"
3. **Holds documents / uncountable** → refusal message naming the count or the read failure, **Close**
   only. No rename controls rendered.
4. **Empty and countable** → a form:
   - New top folder name, pre-filled with the current value, sanitized live the same way `create()`
     sanitizes a new segment's folder.
   - A live conflict check against every OTHER segment's `stagingFolder` (case-insensitive) — reusing
     the same `sanitizeFolderSegment` normalisation the create form uses, so a clash here means the
     exact same thing it means there.
   - Disabled when: the sanitized value is blank, collides with another segment, or is unchanged from
     the current value (a same-value "rename" would recycle an empty tree for nothing).
   - A typed confirmation — type the segment's own label — **always required**, unlike the delete
     dialog's harmless-case exemption: there is no harmless case here, since offering the action at
     all already means a folder tree is about to be recycled.
5. On confirm:
   - Recycle the OLD folder tree across every library `retireLibraries()` names (same recycle call the
     delete flow uses).
   - MERGE the mode row's `StagingFolder` to the new sanitized value.
   - Read the segment's Folder Map rows, keep only those whose `Section` (trimmed, case-insensitive)
     equals the OLD root, delete them — the same block the delete flow already runs when "also delete
     the folders" is ticked.
   - Write an audit row (`EVENT.segmentRecoded`) naming the old and new folder, the folder count
     recycled, and the Folder Map rows removed.
   - Result message tells the admin to run **Folder Reconciliation** next — nothing here builds the
     new tree.

## What this deliberately does NOT do

- Does not create the new folder tree — reconciliation's job, on its own next run.
- Does not touch abbreviation rows, Group Map rows, or any group.
- Does not offer a partial path for a segment holding documents.
- Does not allow editing the segment's **label** — this is the physical folder key only. Renaming the
  label is a separate, much smaller, already-safe edit (a plain `ModeLabel` text change with no folder
  or reconciliation impact) and is not part of this feature.

## Testing status

Built 2026-09-11. Pure module unit-tested (`shared/segmentRecode.ts`). **Not yet exercised on a live
site** — the first real test should be a genuinely empty, freshly-created test segment, recoded, then
reconciled, and its upload path confirmed to resolve correctly under the new key.
