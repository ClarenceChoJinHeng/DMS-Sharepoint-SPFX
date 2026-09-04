# A segment C-Level could view their segment but not delete or share in it

**Date:** 2026-08-19
**Status:** BUILT 1.0.177.0, NOT site-tested.
**Extends** `2026-08-07-role-model-simplification-design.md` (the personas) and
`2026-08-15-deletion-and-share-requests-design.md` (which assumes C-Level acts directly).

---

## 1. Found by reading a run log

The MHO reconciliation printed two warnings on **every** folder in `Documents`:

```
⚠ MHO_SEGVIEW role DEL (CRS Delete) not applied on Documents/MHO/Treasury/Payment
⚠ MHO_SEGVIEW role SHARE (CRS Share) not applied on Documents/MHO/Treasury/Payment
```

Correct reporting of deliberate behaviour — and the behaviour was wrong.

`clevel_segment` is `["SEGVIEW", "DEL", "SHARE"]`, labelled **"Segment — view, delete + share one
segment"**. Its Group Map rows sit at the **segment** tier. Reconciliation fanned a segment-tier row
down for `SEGVIEW` **alone**, so `DEL` and `SHARE` landed on the segment folder only — and every
folder below it has broken inheritance.

**So those two permissions reached nothing.** A segment C-Level read the whole segment and could not
delete or share a single document in it, while the persona picker said they could and the
deletion/share design routed no request to anyone because C-Level was supposed to act directly.

Not a regression: true since the persona was written. It took a run log to surface, because the only
evidence was a warning that read as a safety feature.

## 2. Why it was written that way, and why the reason does not extend

The fan-out gate exists because a segment- or department-tier row **cannot be told apart by tier**
from a pre-2026-07-29 leftover, when `isChainAuthorized` demanded a `MEMBER` row at every tier.
Honouring such a row would hand a viewer Read across an entire business segment.

`SEGVIEW` was exempted because it granted nothing on any site until 2026-08-07, so every row that
exists was written on purpose — *the role name is the consent*.

**That argument covers `DEL` and `SHARE` at the SEGMENT tier equally.** Neither was ever
auto-created there. The only personas that put either role on a segment-tier row are the two
C-Level ones. What the gate refuses is a legacy `MEMBER` row, and this is not one.

## 3. The change

`segmentFanRoles()` in `shared/groupMapModel.ts`, **derived from `PERSONAS`** — every role held by a
persona with `scope: "segment"`. That is precisely what the scope means: a persona scoped to a whole
business segment must reach the whole business segment.

Derived rather than listed because a literal would be a second definition of a C-Level's powers,
free to drift from `PERSONAS` — and the drifting copy would be this one, read by reconciliation and
never by a screen anyone looks at.

### 3.1 ⚠ The safety is the TIER, not the role

`DEL` and `SHARE` are also held by `hou`, whose rows sit at the **unit** tier and are therefore never
fanned at all. The caller checks the row's tier *before* consulting the list; using it on a
department-tier row would re-open exactly the hole the gate exists to close. Pinned by a test that
asserts no upload or approve role is ever in the list — a C-Level has no Staging access, so fanning
one would invent it.

### 3.2 An unrecognised role refuses

`fansFromSegmentTier` takes a plain `string`, because a stored `Role` is list text and can hold a
value outside the union (a hand-edited row, an older build). Unknown answers **false**. The cost of
refusing is a C-Level missing a grant, visible in the log; the cost of allowing is a wide delete
right handed out on a value nothing defines.

## 4. What this widens, stated plainly

**A segment C-Level can now delete or share any approved document anywhere in their segment,
without asking anyone.** That is the client's instruction of 2026-08-19 (*"allow Clevel to delete and
share"*) and matches what the persona has always claimed. It is a real widening of live access and
must be said to the client in those words rather than as "a fix".

Unchanged: C-Level holds nothing in either approval library, so unapproved drafts stay out of reach.
`SEGVIEW`/`GLOBAL` remain absent from `LIBRARY_ROLES.Staging`.

## 5. Side effect worth noting

The two warnings per folder disappear — they *were* the refusals. On MHO that is ~140 lines of
`Documents` log removed, which matters: warnings that appear on every folder train an admin to skip
warnings.

## 6. Migration

**Re-run reconciliation.** No row changes, no group changes, no schema change. Existing
`clevel_segment` mappings gain the grants on the next run.

## 7. Files

| File | Change |
|---|---|
| `src/shared/groupMapModel.ts` | **new** `segmentFanRoles()` / `fansFromSegmentTier()`, derived from `PERSONAS` |
| `src/shared/groupMapModel.test.ts` | 4 tests — derivation, the refused roles, no undeclared role |
| `src/webparts/folderManager/components/FolderManager.tsx` | both fan-out gates consult the tier, then the list |
