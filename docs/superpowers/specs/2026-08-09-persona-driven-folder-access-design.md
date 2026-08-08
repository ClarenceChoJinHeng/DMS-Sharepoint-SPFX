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

**Role chips are removed.** Selecting a persona applies its roles directly; there is no manual role
path from this page. `DELS`, previously the only reason to keep one, now lives inside the Head of
Unit persona.

`SELECTABLE_ROLES` survives as the canonical set of name-derivable roles — the group-name round-trip
tests iterate it — but no longer feeds a picker. Its doc comment must say so, or "selectable" will
read as "offered in the UI" to the next person.

## 6. Migration

No data migration. Group Map rows are unchanged: a folder-scope row's ROLE decides which libraries
it reaches, and that table is code.

1. Deploy.
2. **Re-run Folder Reconciliation.** Existing `_UPL` and `_APR` rows gain their Documents Read grant
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
