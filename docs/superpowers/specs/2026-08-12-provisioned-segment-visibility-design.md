# A segment is offered only once its folders exist

**Date:** 2026-08-12
**Status:** designed, building
**Client instruction (2026-08-12):** *"I also do not want the half built segment to show, the UI UX
experience should be that way, once the Segment and structure and group is properly assigned only
then the uploader should be able to upload."*

Supersedes the closing paragraph of §5 of
[2026-08-12-add-segment-design.md](2026-08-12-add-segment-design.md), which said a new segment would
be visible immediately and that a "draft" state should not be added. The client wants the opposite,
and — see §2 — the mechanism that delivers it is not a draft state at all.

---

## 1. Two symptoms, one cause

These were tracked as separate problems. They are the same one:

1. **A newly added segment appears in the upload form before it has any folders.** Slice B writes a
   `mode` row; the form reads mode rows live; the segment is offered.
2. **A user added to a unit group can pick their unit before reconciliation has run** (observed
   2026-08-11, and the origin of the HTTP 403 in the upload form). The dropdown offers the unit, the
   upload then fails at the destination.

The cause in both cases: **the form decides what to offer from the `DMS Group Map` and the term
tree, and never asks whether the folder exists.** `resolveValidPaths` walks group membership against
the term ancestry and returns every leaf whose chain resolves. `DMS Folder Map` — the list that
records which folders actually exist — is not read until the user presses Upload.

So the form offers a path built from *intent* and only discovers at write time whether that path is
*real*. Both symptoms are that gap, at different tiers.

## 2. The gate is DERIVED, not a flag

Readiness is **"a `DMS Folder Map` row exists for the leaf term"** — the exact condition the upload
path itself checks (`lookupFolderMapping` in `Form.tsx`). Nothing new is stored.

A `Ready` / `Active` column on the mode row was the obvious alternative and is worse:

- It can be **wrong**. An admin ticks it when they believe the segment is finished; the gate then
  asserts something nobody verified. A derived gate cannot disagree with reality because it *is* the
  reality — the same row the uploader's write will look for.
- It is a **second staging mechanism**. `PendingLevels` already stages structure changes
  ([2026-08-11-subtree-migration-design.md](2026-08-11-subtree-migration-design.md)); a second,
  differently-shaped "not live yet" concept invites the two to disagree.
- It only fixes symptom 1. Symptom 2 is per-unit, not per-segment — a flag at segment level cannot
  express "these three units are provisioned and that one is not".

And it needs no migration: existing sites have Folder Map rows for every provisioned unit already, so
the gate is a no-op everywhere it should be.

## 3. Behaviour

Filtering happens at the **leaf**, and the tiers above disappear as a consequence:

| State | What the uploader sees |
|---|---|
| Some units provisioned | Only those units. The tiers above list only what leads to them |
| No unit in a segment provisioned | **The segment is absent** — from the Segment dropdown and, if it was the only one, from the Business Segment / Project toggle |
| No unit provisioned at all | The empty state, with the message in §4 |
| Folder Map unreadable | **Everything is offered, as before** — see §5 |

This is why no per-segment logic is needed to satisfy the client's instruction. A half-built segment
has no provisioned units, so every one of its paths is filtered out and the segment vanishes on its
own. Hiding it is a *consequence* of the leaf gate, not a second rule.

**Site admins are unaffected.** `privileged` users get the full manual cascade over the term set and
no `validPaths` at all; they are the people building the segment and must be able to see and test it.
The Structure Manager's closing checklist is where an admin is told what remains (slice A §5.1).

## 4. The empty state must name the real cause

Today, an uploader with groups but no folders gets:

> Your account isn't fully provisioned to upload — you need membership at every level plus the unit
> uploader role. Contact your administrator.

Which is **wrong**, and wrong in an expensive direction: it sends the administrator to check group
membership that is already correct. So the two states are now distinguished by whether
`resolveValidPaths` returned anything *before* the folder filter:

- **No paths before filtering** → the existing message. It is accurate: this is a membership problem.
- **Paths existed, all filtered out** → a new message naming the actual fix:

  > Your unit's folders haven't been created yet. Your DMS administrator needs to run Folder
  > Reconciliation (and give every unit an abbreviation in `DMS Term Abbreviation` first).

Both causes are named, in the order they must be fixed, for the same reason the upload-time message
at `Form.tsx:1378` names both: re-running reconciliation does nothing for a unit with no abbreviation
row, so "re-run reconciliation" alone is wrong half the time.

## 5. An unreadable Folder Map must NOT hide anything

If the Folder Map read fails, the gate is skipped entirely and every path is offered.

This is the rule already established for the stale-chain guard (gotcha 10b) and for
`AllowedFileTypes` (gotcha 11): **empty is not unknown.** A read failure is evidence of nothing, and
a transient error that silently empties every dropdown takes the whole form down for every uploader
on the site — far worse than the state the gate exists to prevent. The failure mode we accept is the
old one: the path is offered, and the upload fails loudly at the destination with a message that
names the fix.

The upload-time check at `Form.tsx:1373` therefore **stays**. It is now a backstop rather than the
only line of defence, and it is what makes the degraded path safe.

## 6. Implementation

- **`src/shared/segmentReadiness.ts`** (new, pure, unit-tested) —
  `filterProvisionedPaths(paths, mappedTermGuids)`. Generic over the path shape so both web parts
  share it. `mappedTermGuids === null` means *unknown* and returns the input untouched; that
  distinction is the whole of §5. GUIDs are normalised (lowercased, braces stripped) before
  comparison, because a stored `{ABC…}` and a term store's `abc…` are the same term and a
  case-sensitive compare would hide every unit on the site.
- **`Form.tsx` / `BulkUpload.tsx`** — add `loadFolderMapRows` to the mount-time `Promise.all` with a
  `.catch(() => null)`, apply the filter to `resolveValidPaths`'s output, and record whether anything
  was filtered so §4 can pick its message. One extra request per page load, paged, already used by
  reconciliation.
- Both web parts keep their duplicated `resolveValidPaths`; unifying those is a separate change and
  not one to make while altering behaviour.

## 7. Testing

- A unit with a group but **no** Folder Map row: absent from the Unit dropdown; if it was the only
  one, the segment is absent and the §4 message names reconciliation.
- The same unit **after** reconciliation: present, upload lands correctly.
- A segment with a `mode` row and no folders at all: absent for uploaders, **present** for a site
  admin.
- A user with **no** groups: unchanged membership message, not the new one.
- Folder Map read forced to fail: every path offered, upload still refuses loudly at the destination.
- A Folder Map row whose GUID is stored in a different case: unit still offered.
- A provisioned unit in one segment and an unprovisioned one in another: only the first segment is
  offered, and the toggle hides the emptied side.

## 8. Related

- [2026-08-12-add-segment-design.md](2026-08-12-add-segment-design.md) — slice B, whose §5 this changes
- [2026-08-10-structure-manager-ui-design.md](2026-08-10-structure-manager-ui-design.md) — §5.1, the
  six-step checklist this gate makes enforceable rather than advisory
- [2026-08-11-subtree-migration-design.md](2026-08-11-subtree-migration-design.md) — `PendingLevels`,
  the staging mechanism this deliberately does not duplicate
