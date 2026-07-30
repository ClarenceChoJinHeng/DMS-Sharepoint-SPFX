# Leaf-Only Upload Authorization — Design

**Date:** 2026-07-29
**Status:** Agreed, implementing
**Touches:** `src/shared/formModel.ts`, `Form.tsx`, `BulkUpload.tsx`, `formModel.test.ts`

---

## Problem

A fully provisioned uploader is refused by the form.

Kaisya is a member of `DMS_GHO_GF_CORU_UPL`. That group is the *only* group her unit needs:
reconciliation assigns it Contribute on the CORU folder and Read-browse on both ancestors
(`Group Head Office`, `Group Finance`). `DMS Group Map` carries one row for it — the unit row,
role `UPL`. This is the agreed model: **one group per role per unit**, granted at every level.

The form refuses her with *"Your account isn't fully provisioned to upload — you need membership
at every level plus the unit uploader role."*

The cause is `isChainAuthorized`. `collectMembership` builds `memberTerms` from Group Map rows
matching the user's groups, so hers is `{CORU}`. `isChainAuthorized` then demands `memberTerms`
contain **every** term in the resolved chain *plus* the term-set GUID:

```
requireSegmentMembership && !memberTerms.has(termSetGuid)  → false
chainTermGuids.every(g => memberTerms.has(g))              → false
```

Satisfying it needs two extra `MEMBER` rows per group — one at the segment tier, one per
intermediate tier — for every unit, every role. For the 4 head offices that is hundreds of rows
whose only purpose is to re-assert what the unit row already says.

## Why the check is redundant

The chain is not user input. `resolveValidPaths` derives it:

```
loadTermPath(mode.termSetGuid, leaf.termGuid) → [Department, …, Unit]
```

It is the **ancestry of a term the user already holds the UPL role on**, read from the term store.
There is no path by which a user supplies a chain they are not entitled to — the leaf is the
authorization, and the ancestry is a consequence of it. Re-checking each ancestor asks the Group
Map to restate a fact the term store already owns.

The check made sense under the earlier model where a group was mapped per tier. That model was
replaced by one-group-per-unit (memory `dms-group-model-per-role-per-library`), and this check was
not updated with it. `CLAUDE.md` already documents the current intent — *"Only unit-level rows are
needed; the form walks the term ancestry (`loadTermPath`) to reconstruct Segment→Dept→Unit from
just the unit term GUID"* — so the code is the thing that is out of step, not the docs.

## Change

Replace `isChainAuthorized` with `isLeafChainValid`, a **structural** check rather than a
membership one:

| Condition | Result |
|---|---|
| Chain is empty (`loadTermPath` failed or returned nothing) | reject |
| Chain's last GUID ≠ the uploader leaf's GUID | reject |
| Otherwise | accept |

The second condition is the guard worth keeping. `loadTermPath` is a network read wrapped in
`.catch(() => [])`; if it ever returns a partial or unrelated chain, the form would otherwise
route an upload to a folder the user has no claim on. Anchoring the chain to the leaf makes that
impossible without requiring any Group Map rows.

`requireSegmentMembership` and the `memberTerms` argument are dropped from the signature.
`collectMembership` still returns `memberTerms` — it is cheap, and it stays available for a future
read-side feature — but nothing consumes it for upload authorization.

## What still gates an upload

Removing the ancestor check does not widen access. Four gates remain, and the strongest of them is
SharePoint's own:

1. **Group Map** — a `UPL` row for a group the user is actually in. No row, no path.
2. **Term ancestry** — the chain must resolve and terminate at that leaf.
3. **Folder ACL** — Contribute on the unit folder, granted by reconciliation. A user who somehow
   reached the wrong folder still gets 403 from SharePoint.
4. **Content approval** — uploads land as Pending and require an approver.

Gate 3 is the real boundary. The form's check is a UX affordance — it decides what the cascade
*offers* — not the security perimeter.

## Rejected alternative

**Add the MEMBER rows.** Works today, and it is what the current code asks for. Rejected because
it scales as *(tiers × units × roles)* rows that carry no information — every one of them
derivable from the unit row plus the term store — and because a missing row fails with a message
that names the wrong cause, which is exactly how this cost an afternoon. Data that must be kept
in sync with a hierarchy the system can already read is a defect, not a configuration step.

## Test changes

`formModel.test.ts` — the `isChainAuthorized` describe block is replaced by `isLeafChainValid`:

- accepts a chain ending at the leaf, with **no** Group Map rows for the ancestors ← the
  regression this fixes
- rejects an empty chain
- rejects a chain whose last GUID is not the leaf
- accepts case-insensitively and ignores surrounding whitespace (`normGuid` trims and
  lowercases; it does **not** strip braces)

`collectMembership` tests are unchanged.
