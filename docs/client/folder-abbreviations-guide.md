# Folder Abbreviations — Administrator's Guide

**For:** DMS administrators at SD Guthrie
**Last updated:** 2026-08-03

---

## What this is

Folders in the document libraries are named with **short codes**, not the full
department and unit names:

```
GHO / GCA / GMB_STRATCOMMS
```

rather than

```
Group Head Office / Group Corporate Affairs / GMB - Strategic Communications
```

The full names are still there. They appear in the upload form's dropdowns, and on
each folder in the **Full Name** column — click a folder and open the details pane
to see it.

## Why

SharePoint has a hard limit on how long a file's full path can be, and it counts the
*encoded* length, not what you see. A space costs 3 characters and the ampersand the
term store requires (`＆`) costs 9. With full names the longest path was within about
70 characters of the point where uploads start failing — and every new sub-unit ate
into what was left. Short codes take the same path from 257 characters to 129.

## Where they live

One list: **DMS Term Abbreviation**. One row per department and per unit.

| Column | What it holds |
|---|---|
| `Title` | The full name, so the list is readable. Also what lets us repair a deleted term — see below. |
| `TermGuid` | Which term the row is for. Set automatically when seeded; do not edit by hand. |
| `Abbreviation` | The folder name. This is the part you change. |
| `Level` | `Department` or `Unit`. |

Segment codes (`GHO`, `MHO`, `NBPOLHO`) are not in this list — they are the
`StagingFolder` value on the segment's row in **DMS Config**.

## The four rules

**1. An abbreviation must be unique among its siblings.**

Two units under the same department cannot share a code. If they did, both would
resolve to the *same folder* — one folder, one set of permissions, two units'
documents inside it. That is exactly the separation the whole system is built on.

Reconciliation checks this before it creates anything. If it finds a clash it
**stops and creates nothing**, and the log names both units. Give one of them a
different code and run it again.

**2. A unit's code usually needs its sub-group prefix.**

Group Corporate Affairs has two units both called *Strategic Communications* — one
under GMB, one under GC. `STRATCOMMS` for both would be the clash above. So:

- GMB - Strategic Communications → `GMB_STRATCOMMS`
- GC - Strategic Communications → `GC_STRATCOMMS`

Keep the prefix even when there is currently only one, if the department is likely
to gain a second.

**3. Changing an abbreviation renames a live folder.**

The next reconciliation run renames the folder and everything inside it moves with
it. Existing documents are not lost and existing links keep working, because the
system tracks folders by their internal id rather than their name.

If something already occupies the new name, the run **reports the conflict and does
nothing** rather than merging the two folders. Sort that out by hand before running
again.

**4. A missing abbreviation blocks that unit completely.**

No row means no folder, and no folder means nobody in that unit can upload. The
reconciliation log lists it as `SKIPPED (no abbreviation)`, and anyone who tries to
upload there is told to ask you for an abbreviation.

Reconciliation never invents one. A guessed code would create a folder at a path
nothing else expects.

## Renaming a term: do not delete and re-add

> **This is the one that costs real time. Please read it.**

If a department or unit needs a different name in the term store, **rename the
term**. Do not delete it and create a new one with the new name.

The system tracks every term by an internal id, not by its name. A rename keeps that
id, so nothing else has to change. Deleting and re-adding creates a **new** id, and
at that moment three lists — the abbreviation, the folder mapping and the group
permissions — all still point at the old, now dead, id. The unit's folder stops being
recognised and its uploads stop working.

Reconciliation now tries to repair this automatically. It looks for a deleted row and
a newly appeared term **with the same name at the same level**, and if there is
exactly one of each it re-points everything and tells you it did.

It deliberately gives up when the name is not unique. `Tax` exists under three
different departments; so do `Legal` and `PM`. If two candidates match, guessing
wrong would hand another department's staff access to this folder — so it reports
both and waits for you. You then set `TermGuid` on the correct row by hand.

There is no way to undo a term deletion, so a rename is always the safer move.

## Folders left behind

Deleting a term does **not** delete its folder, and does **not** remove anyone's
access to it. The folder stays in the library with its permissions exactly as they
were, reachable by anyone who has the link.

Reconciliation reports these as `NO TERM`. It never deletes them — there are
documents inside. Review each one and delete it yourself if that is what you intend.

## Quick reference

| You see in the log | What it means | What to do |
|---|---|---|
| `COLLISION` | Two siblings share a code. Nothing was created. | Change one code, run again. |
| `SKIPPED (no abbreviation)` | That unit has no row. | Add a row, run again. |
| `renamed X → Y` | A code changed; the folder was renamed. | Nothing. |
| `CANNOT RENAME` | Something already has that name. | Resolve by hand. |
| `✎ … re-created; re-pointed` | A deleted term was matched to its replacement. | Run again to create the folder. |
| `AMBIGUOUS` | Several units share that name; not repaired. | Set `TermGuid` by hand. |
| `ORPHANED` | Term gone, no match. The row was kept. | Delete the row once you are sure. |
| `NO TERM` | A folder no term points to. | Review; delete yourself if intended. |
