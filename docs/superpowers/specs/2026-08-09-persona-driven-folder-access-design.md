# Persona-driven Folder Access, and one group per person

**Date:** 2026-08-09
**Supersedes parts of:** `2026-08-07-role-model-simplification-design.md`
**Retires the role picker from:** `2026-07-22-group-map-builder-design.md`
**Touches:** `groupMapModel.ts`, `FolderManager.tsx`, `GroupMapBuilder.tsx`

---

## 1. What the client asked for

Two things, on 2026-08-09.

**One group per person.** Today a PIC needs `GHO_GF_CORU_UPL` to upload *and* the base group
`GHO_GF_CORU` to see anything in `Documents` — because `UPL` grants nothing there. The client
called re-assigning people into a second group "tiring" and asked that the **same group** be
granted Read on the Documents path instead. Same for Head of Unit.

The base `_MEMBER` group then means what its name says: **someone who only views** and neither
uploads nor approves.

**A Folder Access page driven by personas, not roles.** The page currently shows a persona dropdown
*and* seven role chips. If picking a persona already sets the roles, the chips are either redundant
or a way to build a combination the model never intended.

## 2. The new persona model

| Persona | Roles | Approval Document | Documents |
|---|---|---|---|
| C-Level (global) | `GLOBAL` | — | read everything |
| C-Level (segment) | `SEGVIEW` | — | read one segment, all the way down |
| Head of Department | `DEL` | — | read + delete approved, department-wide |
| Head of Unit | `APR`, `DELS` | approve, and delete pending/rejected | read own unit |
| PIC | `UPL` | upload | read own unit |
| SDG Employee | `MEMBER` | — | read own unit |

Three changes from 2026-08-07:

- **Head of Unit loses `MEMBER`** — `APR` now carries Documents read on its own.
- **Head of Department loses `MEMBER`** — `DEL` is `CRS Delete` = Read + Delete Items, so `MEMBER`
  was already redundant there. Confirmed by the client in the same conversation.
- **`DELS` joins Head of Unit.** It previously belonged to no persona at all — kept alive only so a
  hand-authored row would still work. It is approval-library delete, which is precisely the Head of
  Unit's job. It is NOT Documents delete, so a Head of Unit still cannot remove an approved
  document; that stays with Head of Department.

## 3. The trap: one role, two libraries, two different levels

`ROLE_TO_PERMISSION` is flat — one permission level per role, whatever the library:

```ts
{ MEMBER: "Read", UPL: "CRS Upload", APR: "CRS Approve", DEL: "CRS Delete", DELS: "CRS Delete", … }
```

So simply adding `UPL` to `LIBRARY_ROLES.Documents` would grant **`CRS Upload`** there — Contribute
minus Delete. Every PIC could add and edit **approved** documents in their unit, with no approval
step and nothing in any log to show for it. That is the opposite of the request.

**The lookup therefore becomes library-aware.** On `Documents`, `UPL` and `APR` resolve to `Read`
regardless of what they mean on the approval library:

```ts
const DOCUMENTS_READ_ONLY_ROLES = ["UPL", "APR"];   // write roles, downgraded to Read here

function permissionForRole(lib: LibTarget, role: string): string | undefined {
  if (lib === "Documents" && DOCUMENTS_READ_ONLY_ROLES.indexOf(role) > -1) return "Read";
  return ROLE_TO_PERMISSION[role];
}
```

A list of exceptions rather than a second full table: the table is what an editor reads to answer
"what does this role do", and two of them would let the answer differ by which one you opened.

`LIBRARY_ROLES` becomes:

```ts
{
  Staging:   ["UPL", "APR", "DELS"],
  Documents: ["MEMBER", "DEL", "GLOBAL", "SEGVIEW", "UPL", "APR"],
}
```

## 4. Invariants this must not break

Asserted in `groupMapModel.test.ts` — each one is a silent grant if it regresses.

| Invariant | Why |
|---|---|
| `SEGVIEW`/`GLOBAL` never in `Staging` | a segment-wide viewer reads every unapproved draft in the segment |
| `MEMBER` never in `Staging` | same, one tier down: a viewer reads other people's pending files |
| `DELS` never in `Documents` | a Head of Unit could delete approved documents |
| `DEL` never in `Staging` | a Documents deleter could destroy other people's pending files |
| `UPL`/`APR` on `Documents` resolve to `Read` **only** | otherwise uploaders can edit approved documents |
| `ENTRY` never in `LIBRARY_ROLES` | it is library-scope by definition; a folder-scope ENTRY row means nothing |

The fifth is new and has no precedent in the codebase, so its test names the level explicitly rather
than asserting "not upload" — the failure mode is a *specific wrong level*, not an absent one.

## 5. Folder Access UI

**Library toggle above the group field.** Personas grouped by the library they are *about*:

```
[ Documents ] [ Approval Document ]

Documents                 Approval Document
  C-Level (global)          Head of Unit
  C-Level (segment)         PIC
  Head of Department
  SDG Employee
```

The client's first draft had these two lists swapped. Building it that way would have filed the
C-Level personas under a library they must never touch, presenting the most dangerous mapping in the
model as the normal one. Corrected on confirmation — recorded here because the grouping looks
arbitrary without knowing why, and the same slip is easy to repeat.

Head of Unit and PIC grant in **both** libraries, so they are filed under the approval library — the
thing that distinguishes them from every other persona, which is Documents-only.

**Tier dropdowns follow the persona's scope.** Previously every level the term store could load was
rendered, so a Head of Department was walked down to a unit and then told off for picking one:

| Scope | Tier dropdowns |
|---|---|
| `segment` (C-Level) | none — the Segment dropdown already chose it |
| `department` (HoD) | one; a segment's top-level terms ARE its departments |
| `unit` | the full chain |

C-Level **Global** shows neither Segment nor Tier — it is termless, and that block already sat
behind `role !== "GLOBAL"`.

**Role chips are removed, and adding a persona writes EVERY row it needs.** Selecting a persona
applies its roles directly; there is no manual role path from this page, and Head of Unit's `APR` +
`DELS` are written together. The old one-row-at-a-time flow existed so a half-failed multi-write
could not leave someone half-provisioned — sound while the admin could SEE the roles, but with the
chips gone the roles are an implementation detail, and the client read the two-step checklist as
"the Head of Unit might not get delete". The risk is answered by reporting instead: rows go one at a
time and the toast names how many landed and which failed. `Create group` rolls the whole group back
instead, which is safe only there — the group did not exist a moment earlier.

The persona checklist is read-only: it ticks each role so a partial add stays visible, but offers no
choice the tool has already made.

### 5.1 Making segment scope actually grant something

A segment-tier row carries `UnitTermGuid = the segment's term-set GUID`, so it was always
*authorable* — but reconciliation deliberately kept that GUID out of every descendant's
`ancestorTerms`, so the row granted on the segment container folder and nowhere beneath it. C-Level
(segment) could be created and did nothing.

The GUID is now included, and the protection it provided moves one layer down rather than
disappearing: the fan-down accepts a segment-tier inheritance for **`SEGVIEW` only**
(`segmentTermSets` in `FolderManager.tsx`). The original hazard was a leftover segment-tier `MEMBER`
row from before 2026-07-29 — indistinguishable BY TIER from a deliberate one, and worth Read across
an entire business segment if honoured. `SEGVIEW` cannot be such a leftover: it granted nothing on
any site until 2026-08-07, so every row that exists was written on purpose. The role name is the
consent, exactly as it is for `GLOBAL`.

Not gated on `recon_departmentFanOut`: that switch is about DEPARTMENT rows, and turning it off must
not silently disable a C-Level.

Verified 2026-08-09 — `↳↓ … (inherited from a parent-tier mapping)` on every department and unit
under GHO, **and no such line anywhere under MHO or NBPOLHO, or anywhere in the approval library.**
Those two absences are the test; the grants are just the feature.

## 5.2 All six personas, verified end to end (2026-08-09)

Read the right-hand column first. For every persona the interesting result is what it did **not**
get — a grant that appears is a feature working, a grant that fails to appear where it must not is
the model holding.

| Persona | Granted | Correctly absent |
|---|---|---|
| PIC (`UPL`) | `CRS Upload` on its unit; `Read` on the same unit in Documents | cannot edit or delete in Documents; peer PICs' pending files invisible |
| SDG Employee (`MEMBER`) | `Read` on its unit in Documents | nothing at all in the approval library |
| Head of Unit (`APR`+`DELS`) | `CRS Approve` + `CRS Delete` on its unit; `Read` in Documents | no `CRS Delete` in **Documents** — approved documents stay with HoD |
| Head of Department (`DEL`) | `CRS Delete` on its department, fanned to all nine units under GF | **no line under GHR, GIGA, GLRC, GS, PO or GCA**; nothing in the approval library |
| C-Level segment (`SEGVIEW`) | `Read` across every department and unit of GHO | **no line under MHO or NBPOLHO**; nothing in the approval library |
| C-Level global (`GLOBAL`) | `Read` across GHO, MHO **and** NBPOLHO | nothing in the approval library |

The last column of the bottom three rows is the whole security model: a viewer role reaching the
approval library would read other people's unapproved drafts, and for the two C-Level rows it would
be an entire segment's — or the whole company's — worth.

**Group naming lags the model.** `GHO_GF_CORU_SEGVIEW` and `GHO_GF_CORU_GLOBAL` are named for a unit
they have nothing to do with; both are segment- or tenant-wide. A name that misdescribes a grant is
how the next admin gets it wrong, so rename them (`GHO_SEGVIEW`, `CRS_GLOBAL`) — the integer id
survives a rename, so every grant follows.

`SELECTABLE_ROLES` survives as the canonical set of name-derivable roles — the group-name round-trip
tests iterate it — but no longer feeds a picker. Its doc comment must say so, or "selectable" will
read as "offered in the UI" to the next person.

## 6. Migration

No data migration. Group Map rows are unchanged: a folder-scope row's ROLE decides which libraries
it reaches, and that table is code.

1. Deploy.
2. **Check the Group Map list has `Scope` and `Target` columns.** Without them every row reads as a
   *folder* row, so site/library/page mappings silently become folder rows with no term — and the
   failure surfaces later as a term-store complaint. Folder Access now shows a red banner when they
   are absent (added 2026-08-09; the detection had always run but its result was never rendered).
3. **Re-run Folder Reconciliation.** Existing `_UPL` and `_APR` rows gain their Documents Read grant
   on that run; nothing needs re-authoring.
3. Optionally remove now-redundant base-group memberships for PICs and Heads of Unit. Leaving them
   costs nothing — a second Read grant on the same path is idempotent — so this is tidying, not a fix.
4. Verify as a PIC who is **not** in the base group: they should now open their unit folder in
   Documents, and the link in the approval email should work. This makes §5.8 of
   `2026-08-08-auto-route-flow-and-draft-isolation.md` obsolete for PICs.

## 7. Consequence to state to the client

**In the `Documents` library, every PIC will read every approved document in their unit, at any
confidentiality level.**

Scope, stated precisely because the two libraries behave differently and that difference IS the
product: in the **approval library** a PIC still sees only their OWN pending and rejected files.
Draft isolation is untouched by this change — re-verified 2026-08-09 with two PICs in the same
`_UPL` group, one of whom saw an empty folder while the other's pending file sat inside it. The
widening applies to approved documents only.

That follows from the request and is consistent with the existing model — confidentiality is
metadata, not a permission, and the unit folder is already the smallest boundary this system has.
But it is a real widening: before this change a PIC outside the base group could upload without ever
seeing what else the unit had filed. It should be the client's explicit decision, not a side effect
of a convenience fix.

The structural answer if two people must not see each other's approved documents is unchanged:
**they belong in different units.**
