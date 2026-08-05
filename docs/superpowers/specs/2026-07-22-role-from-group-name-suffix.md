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

---

# AMENDMENT 2026-08-04 — prefix-less names, long-form suffixes

**Status:** AGREED. Client asked for `GHO_GF_CORU_UPLOADER` on 2026-08-04; raised as
breaking, reaffirmed, implemented.

```
DMS_GHO_GF_CORU_UPL   →   GHO_GF_CORU_UPLOADER
```

Two independent changes in one name: the `DMS_` prefix goes, and role suffixes become
whole words.

## A. Suffix matching becomes a table, sorted longest-first

The original chain hand-ordered `endsWith` tests, with a comment explaining that `_DELS`
must be tested before `_DEL` because one is a prefix of the other. Long-form names add
more collisions — `_UPL` is a prefix of `_UPLOADER` — and the failure mode is severe: an
unmatched suffix falls through to **`MEMBER`**, which the library rule puts on
**Documents**. An intended Staging uploader would silently become a Documents reader.

Replaced by a suffix table **sorted by length descending**, which makes collisions
structurally impossible instead of a comment somebody has to honour:

| Role | Accepted suffixes | Library |
| --- | --- | --- |
| `APR` | `_APPROVER`, `_APR` | Staging |
| `UPL` | `_UPLOADER`, `_UPL` | Staging |
| `DELS` | `_DELETER_STAGING`, `_DELS` | Staging |
| `DEL` | `_DELETER_DOCUMENTS`, `_DEL` | Documents |
| `HC` | `_HC` | HC (Phase 2) |
| `SEGVIEW` | `_SEGVIEW` | retired — parsed only so it cannot read as `MEMBER` |
| `MEMBER` | *(no suffix — the base group)* | Documents |
| `GLOBAL` | *(never derived; admin picks it by hand)* | Documents |

**Short forms stay accepted permanently.** Groups already exist on the test site with
`_UPL`/`_APR`; dropping them would strand those rows. Only *new* names are written
long-form.

`_DELETER_STAGING` / `_DELETER_DOCUMENTS` are **our** long forms — the client specified
only `UPLOADER`. They still need confirming, which is cheap because both short forms keep
working regardless.

## B. The prefix was the discriminator, so it is replaced — not deleted

`DMS_` was never decoration; it was the only test for "is this site group ours?". Three
consumers relied on it, two of them failing *silently* without it:

| Consumer | Purpose | With `GHO_…` |
| --- | --- | --- |
| `filterDmsGroups` | group-picker search | group never appears — visible |
| `FolderManager` site-entry self-heal | add managed groups' members to the site-entry group | **silent** — no site entry, run still reports ✓ |
| `naming.matchesAnyGroupPrefix` | CRS/DMS half-rename tolerance | **silent**, same |

The self-heal case produces "only *some* users cannot open the site" while reporting
success — the hardest version to diagnose.

Two different replacements, because the two consumers ask different questions:

**The picker asks "may an admin map this group?"** It must show groups *not yet* mapped,
so Group Map membership cannot be the test. The real requirement was always *exclude the
built-ins* so nobody maps `Site Owners` to a unit folder. Now excluded by **group id**,
read from `AssociatedOwnerGroup`/`AssociatedMemberGroup`/`AssociatedVisitorGroup` plus the
site-entry group. Ids rather than titles because built-in titles derive from the site
title (`Clarence DMS Testing Owners`) — any title blocklist breaks on the next site.

**The self-heal asks "which groups do we manage?"** Already stored: every managed group
has a **`GroupId`** in the Group Map list. Enumerating those rows is strictly more accurate
than a name filter, and additionally stops syncing a group somebody hand-named with our
prefix but never mapped. This is the better design independent of the rename; the rename
is what forces it.

## C. Blast radius is bounded — the name is a hint, not the authority

- Reconciliation grants permissions from the row's **`Role` column**, never the title.
- `roleFromGroupName` drives only (a) the **pre-selected** role when an admin picks a
  group, and (b) the **mismatch warning** when typed name and picked role disagree.

A missed suffix therefore corrupts nothing — it makes the default wrong and makes the
warning fire on every correctly-named group, training admins to click past it. The
failure being prevented is usability, not security.

## D. `DMS_SITE_MEMBERS` deliberately keeps its name

Not a segment/role group, the client did not ask, and it is live on the test site holding
site-level Read — renaming it removes site access for every user until reconciliation
re-runs. Its eventual `DMS`→`CRS` rename belongs to `naming.ts`. Flagged to the client,
not decided here.

## E. Migration: none required

Existing `DMS_*` groups keep working — short suffixes still parse, and neither new
discriminator looks at the prefix. If the client renames groups in `groups.aspx`, Group Map
rows still join by `GroupId`, so a rename is safe and needs no reconciliation run.
