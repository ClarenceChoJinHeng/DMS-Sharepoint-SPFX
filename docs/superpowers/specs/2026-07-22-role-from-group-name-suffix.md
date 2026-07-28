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

**Access model — FINAL, confirmed 2026-07-23. Separate group PER ROLE, isolated PER
LIBRARY. 3 groups per unit:**

| Group | Library assigned to | Permission | Persona |
|---|---|---|---|
| `DMS_<seg>_<dept>_<unit>` (base) | Documents **only** | Read | viewer |
| `..._UPL` | Staging **only** | Contribute | uploader |
| `..._APR` | Staging **only** | Design (= approve) | approver |

**Isolation rule:** the base group is **never** assigned to the Staging library. A
viewer is only in the base group → zero Staging grant → a direct Staging link returns
Access Denied. Staging only ever lists `_UPL`/`_APR`, satisfying "only uploaders and
approvers in Staging."

**Cross-library visibility:** an uploader/approver who also needs to read approved docs
is added to the base group **in addition to** their role group (deliberate double
membership — the price of clean per-library separation). Base membership is optional per
user; the app itself never requires Documents access to upload or approve.

GHO scope = 12 units × 3 = **36 groups** (9 Group Finance + 3 GLRC).

> Decision history: model (a) base-viewer → briefly "model (b) 24 groups, no base"
> (misread of the boss answer, which meant only uploaders/approvers in *Staging*) →
> FINAL: base viewer exists but is Documents-only, so both constraints hold.
> `roleFromGroupName` maps `_UPL`→UPL, `_APR`→APR, base (no suffix)→MEMBER. No code
> change across any of these decisions — group→role→library is configuration only.

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
