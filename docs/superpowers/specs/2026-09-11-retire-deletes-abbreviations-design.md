# Retiring a segment now deletes its abbreviation rows too

Client's decision, 2026-09-11, made after the fruit/car (`PCAR`) abbreviation rows kept
surfacing as `⚠ ORPHANED` on every reconciliation run long after that segment was retired.

## The change

Until now, Retire a Segment deliberately **kept** `CRS Term Abbreviation` rows for the retired
segment — the documented reasoning was that re-creating a segment with the same name/term set
should bring folder names back automatically, with nothing to retype. That is reversed.

**Retiring a segment now deletes its abbreviation rows unconditionally, as part of the same
action — no checkbox, no opt-out.** Client's own words: *"no need checkbox as an option, make it
mandatory."*

- If a segment is ever re-created after being retired, every department/unit (and any coded
  below-Unit level) needs its abbreviation retyped from scratch. That is the accepted cost.
- This does NOT touch tier columns or documents — unchanged from before. Only the naming rows.

## How a row is matched to the segment being retired

Abbreviation rows carry no segment/section field — only `TermGuid`, `Title`, `Level`,
`Abbreviation`. The only reliable way to know "this term belongs to segment X" is to ask the
LIVE term tree, at retire time, before anything is deleted: walk `TermSetGuid` (BFS, every
depth, same shape as the abbreviation editor's own walk) and collect every term GUID under it.
Any abbreviation row whose `TermGuid` is in that set belongs to this segment.

This only works while the terms still exist. **It does not reach back and repair a segment that
was already retired with its terms already deleted from the term store** (PCAR's own case) —
those rows have to be cleared by hand, once, regardless of this change. Going forward, since
abbreviation deletion now happens automatically the moment a segment is retired, this situation
should not recur: there is no longer a window where a segment is retired but its abbreviations
survive to later become orphaned.

## Fail-closed, not best-effort

The walk is bounded (same request cap as elsewhere in this codebase). If it cannot complete —
a term-store read fails, or the cap is hit — the WHOLE segment count goes `unknown`, exactly
like an unreadable Group Map or an unreadable folder tree already do. `canOfferFolderDelete`
already treats `unknown` as "do not offer the destructive option"; the same now applies here.
Guessing "there are no abbreviation rows" when the walk merely failed would under-report what
retiring is about to remove, and would make the confirmation screen understate the action.

At actual deletion time, the walk runs again (state may have moved between opening the dialog
and pressing the button). A failure there does not stop the retire — it is logged the same way
a failed Group Map row read is: reported, `ok: false`, but the mode row still goes and the run
continues. This matches how Group Map row cleanup already behaves; abbreviation cleanup is not
made a harder blocker than the row cleanup it sits beside.

## What changes where

- `shared/segmentDeletion.ts` — `SegmentCounts` gains `abbreviationRows`; `needsTypedConfirmation`
  gates on it too; `survivorLines` drops the "abbreviations stay" promise and states they are
  removed instead; `deletionSummary` counts them.
- `SegmentCreator.tsx` — a new bounded term-tree walk (`loadSegmentTermGuids`), wired into
  `countSegment` (for the preview) and `onDelete` (for the actual deletion, positioned alongside
  the Group Map row cleanup — before the mode row, non-blocking).
