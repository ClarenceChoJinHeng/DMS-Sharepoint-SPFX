# Term-GUID Orphan Repair — Design

**Date:** 2026-08-02
**Status:** Agreed, not yet implemented
**Supersedes:** Task 6 of `docs/superpowers/plans/2026-07-31-folder-abbreviation-naming.md`
**Related:** `2026-07-30-folder-abbreviation-naming-design.md`, `2026-07-29-folder-map-integrity-design.md`

---

## 1. Problem

Three DMS lists are keyed by term GUID:

| List | Key column | Holds |
|------|-----------|-------|
| `DMS Term Abbreviation` | `TermGuid` | `Abbreviation`, `Title` (the term label), `Level` |
| `DMS Folder Map` | `TermGuid` | `FolderUniqueId`, `FolderUrl`, `Section`, `Title` |
| `DMS Group Map` | `UnitTermGuid` | `GroupName`, `GroupId`, `Role` |

Deleting a term in the term store orphans a row in **all three at once**. Re-creating the
term with an identical name repairs none of them: SharePoint issues a new GUID and there
is no undelete.

Observed 2026-08-02 with `Strategy Partner PNG ＆ SI`: the unit was skipped for want of an
abbreviation, no folder was created, and the only signal was one
`SKIPPED (no abbreviation)` line inside a 550-line log. The Group Map orphan produced no
signal at all — that folder would simply have been locked admin-only with no explanation.

What makes this repairable is that the **label survives**. A re-created term carries the
same name, and `DMS Term Abbreviation` stores it in `Title` (written by
`seed-term-abbreviations.js` as `Title: term.label`).

## 2. Why not move abbreviations into the term store

Considered and rejected on 2026-08-02. Putting the abbreviation on the term as a Shared
custom property would make this class of bug impossible — the abbreviation would die with
the term.

Verified live against `/sites/ClarenceDMSTesting`:

- **Read works.** `/_api/v2.1/termStore/sets/{id}/children?$select=id,labels,properties`
  returns `properties: []`. The collection exists; it is omitted from the default
  response, which is why an earlier probe appeared to show no support at all.
- **The UI can edit it.** Term → **Advanced** → *Shared custom properties* → Edit.
- **The API cannot.** `PATCH /_api/v2.1/termStore/sets/{setId}/terms/{termId}` with
  `{properties:[{key,value}]}` and a valid form digest returns **403 Forbidden**. 403 is
  permission-denied, not unsupported: term-store writes check a taxonomy role separate
  from Site Collection Admin.
- **That role is unobtainable.** The client cannot be made a term store administrator or
  term group admin — it requires their security team's approval.

Without an API write there is no bulk path (CSV import does not carry custom properties),
so seeding 175 abbreviations would be 175 manual UI entries on data where a typo silently
misnames a folder or collides with a sibling. An unverifiable manual seed is worse than
the repair logic below.

> Revisit only if the 2027 segments are onboarded, which needs a fresh seed anyway — and
> only after re-testing the API write.

## 3. Principles

1. **Repair beats delete.** Prune already deletes orphaned Folder Map rows automatically
   today. Automatically *repairing* an orphan is strictly less destructive than behaviour
   already shipped, so automating it does not raise the risk ceiling.
2. **Derivable state may be deleted automatically; authored state may not.** A Folder Map
   row is rebuilt from the term plus the folder on disk. An abbreviation row exists
   nowhere else.
3. **Never guess when a wrong guess grants permissions.** A Group Map re-point changes who
   can reach a folder.
4. **Never act on a partial term-store read.** A throttled read must not look like a mass
   deletion. Everything here sits behind the existing `incomplete` guard.
5. **Never delete a folder.** Documents sit behind it.

## 4. Discovery

An abbreviation row is an **orphan** when its `TermGuid` is absent from the enumerated
live term list for the run.

A live term is **unclaimed** when no abbreviation row carries its GUID. These are already
detected — they are the `missingAbbrev` entries.

Orphans and unclaimed terms are the two halves of a delete-and-re-add.

## 5. Matching

A repair is proposed only on a **1:1 match in both directions**:

- exactly **one** orphaned abbreviation row whose `Title` and `Level` match, **and**
- exactly **one** unclaimed live term whose label and depth match

Both are required. Bare labels are **not** unique in this term store — `Tax` exists under
`GHO/GF`, `MHO/GA` and `NBPOLHO/FIN`; so do `Legal`, `GCA`, `PM`, `Group Compliance`,
`Risk Management` and `Transformation`. The abbreviation list stores `Level` but not the
parent chain, so a label match alone could re-point a row to the wrong department's unit —
and the same repair propagates to the Group Map, making that a permissions grant to the
wrong folder.

Comparison is case-insensitive and trims whitespace. Fullwidth `＆` is compared as-is; it
is the stored form on both sides.

Where either side has more than one candidate, no repair is proposed. The orphan is
reported and left alone.

## 6. Propagation

The join across the three lists cannot *discover* the new GUID — every row holds the same
dead GUID, so joining them yields stale rows agreeing with each other. Only the term store
knows the new GUID, and only the label reaches it.

The join is how the answer is *applied*. Once `oldGuid → newGuid` is established for a
unit, every row carrying `oldGuid` in any list is the same unit by definition:

```
orphaned abbrev row → match Title to an unclaimed live term → newGuid
                            ↓
      oldGuid → update every row holding it, across all three lists
```

`DMS Folder Map` needs no explicit repair: reconciliation already writes a fresh row for
any unmapped term (the `if (!existingRow)` branch in `FolderManager.tsx`). Listed for
completeness only.

## 7. Action policy

| Situation | Action | Automatic |
|-----------|--------|-----------|
| Orphan with a 1:1 match | Re-point abbreviation row + every Group Map row on the old GUID | **Yes**, logged |
| Orphan with an ambiguous match | Report both candidate sets | No |
| Orphan with no match (true deletion) | Report; offer deletion | No — admin clicks |
| Orphaned **Folder Map** row | Delete | **Yes** — unchanged, derivable |
| Orphaned **Group Map** row with no match | Report | No — permissions change |
| Folder no live term claims | Report | **Never** deleted |
| Any of the above, `incomplete` non-empty | Skip the entire pass | — |

This revises the plan's Task 6, which pruned abbreviation rows automatically under the
`incomplete` guard. That is now judged wrong: deleting an abbreviation row destroys the
only copy of authored data, and if the term is later re-created with a *different*
abbreviation, Task 4's rename pass renames the live folder and its documents to match. An
auto-prune can therefore end in a folder rename nobody requested.

## 8. Unclaimed folders

Reconciliation only ever walks from the term store to the libraries, never the reverse. A
permanently deleted term leaves its folder in Staging and Documents — broken inheritance
intact, still granting Contribute and Design to that unit's groups. Deleting a term does
**not** revoke anyone's access; the folder stays reachable by direct link or by browsing
the library.

Add a reporting pass: after the target walk, list folders at term depth that no target
claims. Report only, never delete — there are documents behind them. Cheap to add here
because both sides are already in memory at that point.

## 9. Reporting

Repairs and proposals go in the existing "Needs attention" panel.

An applied repair reads:

```
✎ SP_PNGSI — term re-created; re-pointed to fdf73198… (+2 group-map rows)
```

A proposal names both sides and the group-map blast radius, so an admin can see what a
click would change before clicking it.

## 10. Out of scope

- Preventing term deletion. SharePoint has no protect-from-deletion flag, and anyone with
  Contributor on the term group can delete. Mitigation is procedural and belongs in the
  Task 7 client guide: **rename terms, never delete and re-add** — a rename keeps the GUID
  and orphans nothing.
- Restoring a deleted term. Not possible.
- Reconciling folder *contents*.
