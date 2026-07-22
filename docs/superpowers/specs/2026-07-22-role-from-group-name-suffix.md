# Auto-derive Role from Group-Name Suffix — Design

**Date:** 2026-07-22
**Status:** Approved for implementation
**Component:** Group Map Builder (DMS Admin Tool → Group Map tab)

## Background

The CTO's group model collapses the old stacked-group access (segment + dept + unit
groups) into **one self-contained group per unit**, whose name encodes the full path
and whose role is encoded as a name suffix:

| Group name | Meaning |
|---|---|
| `DMS_GHO_GF_CORU` | viewer / read (base) |
| `DMS_GHO_GF_CORU_UPL` | uploader |
| `DMS_GHO_GF_CORU_APR` | approver |

**Access model (a), confirmed:** a base viewer group per unit, with `_UPL` / `_APR`
role groups added on top. Uploaders/approvers are members of the base group **and**
their role group. (Pending boss confirmation on whether pure-viewer members exist —
non-blocking; model (a) supports both outcomes with no code change.)

## Scope

The names are **cosmetic** — the form still matches groups by Entra Object ID against
`DMS Group Map` rows, and segment/tier still come from the term-store cascade the admin
picks. The path-in-name is human convenience only.

The **only** behavioural change: when the admin picks a group in the Builder, the Role
selector **pre-selects** the role implied by the name suffix, as an editable default.

**Question 1 (one row per group) needs no code.** `isDuplicateRow` already keys on
`GroupId + UnitTermGuid + Role`, so three rows for one unit (base=MEMBER, UPL, APR)
are already allowed and de-duplicated correctly.

## Change 1 — `roleFromGroupName` (pure helper, `src/shared/groupMapModel.ts`)

```
roleFromGroupName(name: string): GroupMapRole
  - trims, compares case-insensitively
  - ends with "_APR"  -> "APR"
  - ends with "_UPL"  -> "UPL"
  - otherwise         -> "MEMBER"
```

`GLOBAL` is **never** auto-derived — it is a privileged bypass the admin selects by hand,
not a name convention. Unit tests cover: `_APR`, `_UPL`, no-suffix, lowercase suffix,
empty/undefined name → `MEMBER`.

## Change 2 — pre-select in `GroupMapBuilder.pickGroup`

When a group is chosen from the search results, call `roleFromGroupName(g.displayName)`
and `setRole(...)`. The existing `pickRole` handler is untouched, so the admin can still
click any other role button — the suffix only sets the default.

## Out of scope / unchanged

- Object-ID matching, term-store cascade, folder ACL assignment.
- `buildGroupMapRow`, `isDuplicateRow`, `validateDraft`.
- The `DMS Group Map` list schema.
