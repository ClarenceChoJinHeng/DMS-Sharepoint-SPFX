# Segment rename on the Abbreviations screen, and a "rename matching groups" convenience

Client, 2026-09-15: *"the big implementation comes after you finally add the rename a segment term
abbreviation in the same place for the existing renaming term abbreviation. Client wants that. Add a
renaming Group button once a segment or department or unit term is renamed, we had that before."*

Confirmed via clarifying questions:
- Segment rename should be **added to** the CRS Term Abbreviations screen (`AbbreviationManager.tsx`),
  alongside where Department and Unit already get renamed — not moved off the Segments tab, which
  keeps its existing recode form untouched.
- The "rename Group" ask means: after a term's (or a segment's) code is renamed, offer a button that
  renames the matching SharePoint group(s) to use the new code — e.g. `TAX` → `TAXX` renamed offers to
  rename `GHO_GCA_TAX_APPROVER` → `GHO_GCA_TAXX_APPROVER` in one click.

## Part 1 — Segment rename, added to the Abbreviations screen

### What already exists (2026-09-11, `shared/segmentRecode.ts` + `SegmentCreator.tsx`'s Segments tab)
- Pure validators: `canOfferRecode`, `recodeRefusalReason`, `folderRecodeConflict`,
  `folderRecodeIsNoOp`, `recodeSummary` — all in `shared/segmentRecode.ts`, unit-tested, SPFx-free.
- The write sequence lives INLINE as `onRecode` in `SegmentCreator.tsx` (~1429-1554): recycle the old
  (confirmed-empty) folder tree across every `retireLibraries()` target, MERGE the mode row's
  `StagingFolder`, delete the stale `CRS Folder Map` rows for the old `Section`, write one audit row
  (`EVENT.segmentRecoded`), report a line-by-line log.
- `countSegment` (the fresh document-count read that gates the whole action) is ALSO inline, local to
  `SegmentCreator.tsx`.

### The gap
Neither the write sequence nor the count function is exported — they are closures inside one
component. The Abbreviations screen cannot reuse them without either (a) duplicating ~170 lines of
destructive-write logic, which this codebase has repeatedly learned costs a drifting second copy, or
(b) a proper extraction.

### The plan: extract, don't duplicate
1. **New file `src/shared/spSegmentRecode.ts`** (SPFx-boundary, mirrors the existing split between
   `segmentDeletion.ts`'s pure rules and its SPFx-facing callers):
   - `countSegmentDocuments(context, siteUrl, seg): Promise<SegmentCounts>` — moved verbatim from
     `SegmentCreator.tsx`'s local `countSegment`.
   - `performSegmentRecode(context, siteUrl, seg, newFolder, counts): Promise<{ ok: boolean; lines:
     string[] }>` — the write sequence moved verbatim from `onRecode`, MINUS the component-local state
     updates (`setRecodeLog`, `setResult`, `setExisting`, `setRecoding`, `reload()`) and MINUS the
     audit-row write, which the caller issues itself (callers differ on `source:` — `"SegmentCreator"`
     vs `"AbbreviationManager"` — and on what happens to their own list/reload afterward).
2. **`SegmentCreator.tsx`'s `onRecode` becomes a thin wrapper**: calls the two new functions, keeps its
   own state updates and its own audit-row write (`source: "SegmentCreator"`), otherwise unchanged
   behaviour. This is the low-risk half — same writes, same order, just delegated.
3. **`AbbreviationManager.tsx` gains a "Segment" section above Department**, using the SAME two shared
   functions with its own compact form (folder-name input, typed-confirmation of the segment's own
   label — matching `SegmentCreator`'s existing confirmation gate, since this stays "destructive,
   never a harmless default"), its own audit write (`source: "AbbreviationManager"`).
   - Rendered only when a segment is chosen (this screen already requires one) and gated behind the
     SAME `canOfferRecode` check — refused outright with the named reason when the segment holds
     documents or the count could not be confirmed. No partial path, matching the original design.
   - Runs its own fresh count on mount/segment-change (not shared state with the Segments tab — the two
     screens are independent mounts, and a stale count here would be exactly the guessed-at-emptiness
     failure `canOfferRecode` exists to prevent).

### What is explicitly NOT changed
- `SegmentCreator.tsx`'s own recode UI, dialog, and its position on the Segments tab — untouched,
  still there, still works exactly as before. This is an ADDITION, not a relocation.
- The typed-confirmation gate, the empty-only rule, the refusal wording — all reused verbatim via the
  pure module; no new validation logic.

## Part 2 — "Rename matching groups" after any code change

### The mechanism a group name is built from
Confirmed by reading `bulkGroups.ts` and `groupMapModel.ts`: a group's name is
`suggestGroupName(segmentCode, tierCodeChain, role)`, where `tierCodeChain` comes from `codeChain()` —
walking a term's ancestry through the SAME `CRS Term Abbreviation` codes this screen edits. So renaming
`TAX` → `TAXX` (or a segment's `StagingFolder`) leaves every group whose name embeds the OLD code
pointing at a folder that no longer exists under that name — the group itself, its Group Map rows and
its ACLs are all still correct (nothing here ever revokes access), only the NAME goes stale.

### Existing, reusable building blocks — nothing new needs writing for the write side
- `renameSiteGroup(sp, siteUrl, id, newTitle): Promise<void>` — already in `shared/spGroups.ts`,
  already proven (this is what `GroupManager.tsx`'s existing suffix-standardisation bulk-rename button
  already calls). Handles the duplicate-name case by throwing a named error.
- `suggestGroupName` / `codeChain` / `namingRoleFor` — already exist, already tested, already exactly
  what computes a group's canonical name from a segment code + tier code chain + persona.
- Matching WHICH groups are affected: read the CRS Group Map (`GroupId, GroupName, Role, Scope, Target,
  Segment, UnitTermGuid`), filter to rows whose `UnitTermGuid` is one of the renamed term GUIDs
  (Department/Unit case) or whose `Segment` equals the segment's own `termSetGuid` with a BLANK
  `UnitTermGuid` (segment-tier rows, e.g. `SEGVIEW`/`GLOBAL` — though `GLOBAL` is excluded, it is
  site-wide and never carries a segment stem). Group by `GroupId` since one group can carry several
  mapping rows (six for an approver group).

### New pure function needed: `plannedGroupRenames`
In `shared/groupMapModel.ts`, beside `suggestGroupName`/`codeChain`:
```
plannedGroupRenames(
  renamedTermGuids: string[],   // department/unit terms whose code just changed
  segmentRecoded: boolean,       // true when THIS save was a segment top-folder recode
  segmentCode: string,           // the NEW segment code
  rows: GroupMapReadRow[],       // full Group Map, unfiltered
  codeChainFor: (termGuid: string) => string[] | undefined,  // caller-supplied, since chain-walking needs the abbreviation index this screen already has loaded
): Array<{ groupId: number; from: string; to: string }>
```
Returns only entries where `from !== to` (a group whose current name already matches needs no rename —
this can genuinely happen if it was hand-renamed already, or if the rename did not touch a code
position this group's chain passes through).

### UI: offered right after a successful save
- Both screens (Abbreviations' existing Department/Unit rename flow, and the new Segment section) call
  `plannedGroupRenames` immediately after a successful save, using the rows/codes as they now stand.
- If the list is non-empty, a confirm bar appears (same pattern as this project's other bulk-rename
  tools — button first, "Confirm — rename N group(s)" / "Skip" second click, never a native
  `window.confirm`), naming every `from → to` pair before running.
- Sequential, carries on past failures (same shape as `runStandardiseNames`), one audit row for the
  whole batch (reuse `EVENT.groupMapChanged` — this is changing what a mapped group is CALLED, the
  closest existing category; do not invent a new `EventType` unless asked, per the standing "grep
  every reader of `ALL_EVENT_TYPES`" lesson).
- **Explicitly optional, always skippable.** A rename here changes NOTHING about access — Group Map
  rows are keyed on `GroupId`, which a SharePoint rename preserves, so skipping this step is always
  safe and leaves the site in the same "mixed old/new names" state it has lived in before (see
  `1.0.261.0`'s "safe because a rename preserves the group id" note, same reasoning again).

### What this deliberately does NOT do
- It does not touch group MEMBERSHIP, Group Map rows, or any folder ACL — rename only.
- It does not attempt to rename `CRS_SITE_MEMBERS`, the Owners group, or any hand-named group —
  `plannedGroupRenames` only ever proposes names for groups the Group Map maps against the renamed
  term(s)/segment, so anything outside that scope is never a candidate.
- It is NOT the same feature as `GroupManager.tsx`'s existing suffix-standardisation button
  (`runStandardiseNames`/`canonicalGroupRename`) — that one fixes OLD SUFFIX SPELLINGS
  (`_APR` → `_APPROVER`); this one fixes a STALE CODE SEGMENT inside an otherwise-correct name. Both
  can legitimately apply to the same group at different times; neither replaces the other.

## Build order
1. Extract `spSegmentRecode.ts`, refactor `SegmentCreator.tsx` to use it — verify no behaviour change
   (existing tests + a careful read of the diff).
2. Add the Segment section to `AbbreviationManager.tsx`.
3. Add `plannedGroupRenames` to `groupMapModel.ts` with unit tests.
4. Wire the "rename matching groups" confirm bar into both the existing Department/Unit save path and
   the new Segment recode path.
5. `heft test`, `npm run build`, verify the shipped bundle, bump version, ship.
