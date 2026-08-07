# Role Model Simplification — Design

**Date:** 2026-08-07
**Status:** agreed in outline — two assumptions flagged in §7, implementing now
**Branch:** `feat/folder-abbreviations`
**Supersedes:** `2026-08-03-role-personas-and-department-fanout-design.md` (the twelve-persona
model) and `2026-08-03-visibility-scope-and-view-group-separation-design.md` (DRAFT, never
implemented — its five open decisions are void)

---

## 1. What the client's third restatement changes

Two structural changes, and they pull in the opposite direction from the second restatement:

1. **Navigation is back to normal, in BOTH libraries.** The second restatement said Head of
   Department "cannot see the Segment" and Head of Unit "cannot see the department folder". That
   is withdrawn. Everyone navigates from their business segment downward, in Documents *and* in
   Staging. "View the HOU unit folder ONLY" meant *not sibling units* — not *no corridor*.
2. **Head of Department no longer approves.** Approval belongs to Head of Unit. HoD becomes a
   view-only role whose only distinction is breadth: every unit under their department.

Everything else follows from those two.

### 1.1 The ancestor-Read grant has to be RESTORED, not left alone

An earlier draft of this spec said the corridor still worked and needed no change. **That was
wrong**, and the error mattered: the 2026-08-04 requirement had already been implemented, so the
ancestor-Read **granting code was deleted** and replaced with an opt-in **revoke** pass.

Why it looked harmless at the time: reconciliation had only ever ADDED assignments, so every
already-provisioned site still carried ancestor Read from earlier runs. Removing the grant changed
nothing visible, and the run log reported a clean pass.

It bites on a **new** library. The `Approval Document` library was recreated on 2026-08-06 with
zero folders, so every folder reconciliation creates is new, none receives ancestor Read, and the
library root renders **empty for every non-admin** — the bug from
`dms-content-approval-blocks-uploader`, arrived at by a completely different route.

**Changes:**

- **Restore** the ancestor-Read grant in the assignment loop, idempotently (a second unit under
  one department must not re-grant its parent).
- **Hard-off `recon_revokeAncestorRead`**, ignoring the config row. The requirement it served is
  withdrawn, and with granting restored the two passes would fight — an admin flipping that row
  would break navigation for every user on the next run.

Sibling folders still get nothing and stay security-trimmed, so restoring the corridor exposes
only the ancestor folders a user must traverse — which is exactly what the client now asks for.

## 2. The personas

Six, down from twelve.

| Key | Family | Scope | Roles | Documents | Staging |
| --- | --- | --- | --- | --- | --- |
| `clevel_global` | C-Level | segment | `GLOBAL` | every segment, view | none |
| `clevel_segment` | C-Level | segment | `SEGVIEW` | one segment, view | none |
| `hod` | Head of Department | department | `MEMBER`, `DEL` | segment → dept → **all units under it**, and **delete** | none |
| `hou` | Head of Unit | unit | `MEMBER`, `APR` | own path | **approve + view all files in the unit** |
| `pic` | PIC | unit | `UPL` | none by default | upload + view the unit's pending files |
| `employee` | SDG Employee | unit | `MEMBER` | own path | none |

### 2.1 What was removed, and why

- **HoD 1–4 → one `hod`.** The four bundles existed to combine approve/upload/delete at
  department tier. HoD holds none of those now.
- **HoU 1–4 → one `hou`.** Same collapse; the client describes exactly one Head of Unit shape.
- **PIC 2 (Highly Confidential)** — removed at the client's instruction, 2026-08-07. It was
  already `unavailable` (HC left Phase 1 and its term was deleted). The code remains on
  `feat/hc-libraries`.
- **Head of Sub Unit** — explicitly deferred by the client. It would need a third tier in the
  term store, `SubUnit` + `SubUnitTid` columns, and a `Levels` JSON change. Out of scope.

### 2.2 PIC keeps `UPL` alone — deliberately

A PIC gets no Documents access from this persona. That is unchanged from the 2026-08-04
correction and still right: whether someone may read the approved archive is a separate grant,
made by also placing them in the unit's base group (`employee`). Bundling it made every PIC a
Documents reader by default, which is the wrong default for a permission.

## 3. `SEGVIEW` has to be built, not wired

`clevel_segment` is the only genuinely new capability. `SEGVIEW` exists as a role string and a
group suffix, but it was **retired the day after it was added** and is deliberately absent from
`ROLE_TO_PERMISSION`, so today a `SEGVIEW` row is skipped on every folder in every library.

It is `GLOBAL` narrowed to one segment:

| | `GLOBAL` | `SEGVIEW` |
| --- | --- | --- |
| Term on the row | none — termless by design | the **segment** term |
| Reaches | every segment | one segment and everything under it |
| Library | Documents only | Documents only |
| Permission | `Read` | `Read` |

The fan-**down** machinery already exists for `GLOBAL` (the C-level view fan-down). `SEGVIEW`
reuses it with a filter: instead of every unit folder, only those whose segment matches the row's
term.

**Documents-only is not a detail.** A C-Level on Staging would be reading everyone's unapproved
drafts across an entire segment. `LIBRARY_ROLES.Staging` must never contain `SEGVIEW` or
`GLOBAL`.

## 4. Head of Department needs the fan-out ON

Unit folders have unique permissions, so a grant at department tier stops dead at the department
folder and never flows down. `recon_departmentFanOut` already implements the fan-out but
**defaults to off**.

With it off, an `hod` row grants `Read` on the department folder and nothing else — the HoD sees
a folder containing, as far as SharePoint tells them, no units. That is not a partial result; it
is the feature not working.

**Change:** default `recon_departmentFanOut` to on.

## 5. Delete roles

**`DEL` belongs to Head of Department** (confirmed 2026-08-07). It is Documents-only, which
`LIBRARY_ROLES.Documents` already enforces — so an `hod` group is granted delete on approved
documents and never reaches Staging at all.

Two consequences worth stating plainly, because neither is obvious from "HoD can delete":

- **Delete fans out with the rest of the persona.** `hod` is department-scoped, so with the
  fan-out on (§4) the delete grant lands on **every unit folder under that department**, not just
  the department folder. A Head of Department can delete any approved document in any of their
  units. That follows from the scope, but it is a wide grant and the client should hear it in
  those words.
- **Deleted means recycle bin, not gone.** A site-collection admin can restore for 93 days. Worth
  saying to the client, because "delete" in a document management system tends to be heard as
  irreversible.

**`DELS` (Staging, pending files) belongs to no persona.** HoD has no Staging access, and the
client has not asked for pending-file delete anywhere. The role, its suffix and its permission
mapping are **kept** — an admin can still author a `DELS` row by hand and reconciliation will
honour it — but no persona offers it, so no group is created carrying it.

Keeping the unused role costs nothing; removing it would mean re-adding suffix parsing,
permission mapping and library rules the first time someone asks for it back.

## 6. What does not change

- Sibling folders are never granted, so SharePoint keeps security-trimming them (§1.1)
- Leaf-only upload authorization (`2026-07-29-leaf-only-upload-authorization-design.md`)
- One group per role per unit; no bundle group is ever created; SharePoint groups cannot nest
- The unit folder remains the smallest confidentiality boundary
  (`dms-per-uploader-isolation-rejected`)

## 7. Assumptions — confirm before this ships

- **A1 — RESOLVED 2026-08-07.** `DEL` belongs to Head of Department. `DELS` belongs to no
  persona and is kept unused. See §5.
- **A2 — `recon_departmentFanOut` on by default.** Required by §4. It changes what an existing
  reconciliation run does on any site already configured, so it is called out rather than slipped
  in.

## 8. Not in this spec

Per-uploader file isolation in Staging (Draft Item Security = approver + author, plus
move-on-approve in the Auto-route flow) is a **library and flow** change, not a role change. It
is being decided with the client separately and will get its own spec if adopted.

## 9. Implementation order

1. `groupMapModel.ts` — replace `PERSONAS`; keep `Persona`, `PERSONA_FAMILIES`, the role type
2. `groupMapModel.ts` — add `SEGVIEW` to `SELECTABLE_ROLES`
3. `FolderManager.tsx` — `ROLE_TO_PERMISSION.SEGVIEW = "Read"`; add `SEGVIEW` to
   `LIBRARY_ROLES.Documents` (**never** `Staging`)
4. `FolderManager.tsx` — segment-scoped fan-down for `SEGVIEW`, reusing the `GLOBAL` pass
5. `FolderManager.tsx` — default `recon_departmentFanOut` to on
6. Tests: persona shape, `SEGVIEW` never on Staging, fan-down filters by segment
7. CLAUDE.md — the RBAC table
