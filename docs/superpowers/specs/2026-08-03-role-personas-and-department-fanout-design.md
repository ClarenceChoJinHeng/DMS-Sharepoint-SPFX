# Role Personas — Head of Department, Head of Unit, PIC — Design

**Date:** 2026-08-03
**Status:** agreed in outline, implementing now
**Branch:** `feat/folder-abbreviations`
**Parent design:** `2026-07-16-highly-confidential-securing-design.md` §4–§6 (on
`feat/hc-libraries`). This spec is the **Phase 1 subset** of that role model — the part
that can ship without Highly Confidential.
**Related:** `2026-07-29-leaf-only-upload-authorization-design.md` (the leaf-only rule this
spec extends), `2026-07-22-group-map-builder-design.md`

---

## 1. Why this is not just "add two groups"

The client's eight personas are membership combinations of atomic per-unit groups. Six of
the eight already work on this branch. What does not:

| Persona | Groups | Works today? |
| --- | --- | --- |
| Head-of #1 | `{Unit}` + `_UPL` + `_APR` | yes |
| Head-of #2 | `{Unit}` + `_APR` | **no** — `_APR` maps to Design, which includes Add Items, so an approver can always upload |
| Head-of #3 | `{Unit}` + `_APR` + `_DEL` | **no** — `_DEL` does not exist |
| Head-of #4 | `{Unit}` + `_UPL` + `_APR` + `_DEL` | **no** — same |
| PIC #1 | `{Unit}` + `_UPL` | yes |
| PIC #2 | `_HC` only | **Phase 2** — see §7 |
| PIC #3 | `_UPL` only | yes |
| SDG Employee | `{Unit}` only | yes |

And cutting across all of them: **Head of Department and Head of Unit hold the identical
group sets and differ only in scope.** Head of Unit is their own unit — a unit-tier row,
which is what the system already does. Head of Department is every unit under the
department, and that has never worked. Unit folders have unique permissions, so a
department-tier assignment stops dead at the department folder and never flows down.

So the work is three capabilities, not eight groups: **approve-without-upload**,
**delete-in-Documents**, and **departmental fan-out**.

### 1.1 The client's document, as twelve groups

Restated by the client on 2026-08-03 as four families rather than eight personas. The role
sets are unchanged from the table above; what the document adds is that Head of Department
and Head of Unit are **named separately**, so the UI names them separately too.

| Family | Group | Roles | Tier |
| --- | --- | --- | --- |
| Head of Department | 1 — approve + upload | `MEMBER` `UPL` `APR` | department |
| | 2 — approve only | `MEMBER` `APR` | department |
| | 3 — approve + delete | `MEMBER` `APR` `DEL` | department |
| | 4 — approve + upload + delete | `MEMBER` `UPL` `APR` `DEL` | department |
| Head of Unit | 1–4 | **identical to the above** | unit |
| PIC | 1 — upload + view | `MEMBER` `UPL` | unit |
| | 2 — Highly Confidential only | `HC` | unit — **Phase 2** |
| | 3 — upload only | `UPL` | unit |
| SDG Employee | no power — view only | `MEMBER` | unit |

Three things this settles:

**"see all unit" means every Head-of group holds `MEMBER`.** Approvers get view rights by
the client's model. That is a separate question from whether the approval page *requires*
Documents access in order to approve — it does not, and must not; see §10.

**The document repeats "see all unit under the department" under Head of UNIT.** Confirmed
2026-07-31 as a copy-paste error, so Head of Unit is scoped to its own unit. Recorded rather
than silently assumed, because that one phrase is the difference between a unit head seeing
one unit and seeing twelve. If the client did mean it literally, Head of Unit and Head of
Department are the same thing and four of the twelve groups are redundant — worth confirming
out loud before anyone is provisioned.

**PIC 2 is listed in the picker but greyed out**, with the reason shown. Silently omitting
one of the client's own numbered groups reads as an oversight, and an admin would go hunting
for it.

## 2. No bundle groups

Personas are never groups. `CRS_GHO_GCA_EG_HEAD_OF_DEPT_TYPE_3` would be a fourth name to
keep in sync with a term rename, and reconciliation already has three lists to keep joined.
A persona is a **membership combination**, chosen in the UI, written as one row per atomic
group.

## 3. Roles

`GroupMapRole` gains `DEL` and `HC`. `HC` is added to the **type** now, though nothing in
Phase 1 assigns it, so `feat/hc-libraries` merges later without touching the shared model
again. It is deliberately **not offered in the role picker** — an `HC` row created today
would be filtered out of every library and silently do nothing.

| Role | Library | Permission level | Status |
| --- | --- | --- | --- |
| `MEMBER` | Documents | Read | exists |
| `UPL` | Staging | Contribute | exists |
| `APR` | Staging | **DMS Approve** | exists; **level changes from Design** |
| `DEL` | Documents | **DMS Delete** | new |
| `HC` | — | — | type only; Phase 2 |

### 3.1 The two custom permission levels are a client prerequisite

Created once per site at **Site Settings → Site permissions → Permission Levels**, by
copying an existing level so its supporting permissions come along:

| Level | Copy from | Then |
| --- | --- | --- |
| **DMS Approve** | Contribute | tick **Approve Items**; untick **Add Items**, **Delete Items** |
| **DMS Delete** | Read | tick **Delete Items** |

Copy Contribute rather than Design so Manage Lists is not granted. Approve Items depends on
Edit Items, not Add Items, so approve-without-upload is a legal combination.

> **Operational warning, and the reason this is called out rather than buried.** The moment
> `APR → DMS Approve` ships, every approver assignment warns and skips until that level
> exists on the site — reconciliation already logs `no "<name>" role definition on site`.
> That is the correct failure (visible, and it grants nothing by mistake) but it means
> **existing approvers keep their current Design grant and no new one is added**. Nobody
> loses access; the change simply does not take effect, and "approve only" stays
> unenforced. Create both levels before the next reconciliation run.

## 4. Departmental fan-out

A Group Map row on a **non-leaf** term is applied to that folder *and* to every folder
beneath it, at the row's own permission level.

This is the mirror of the ancestor fan-**up** that already exists: fan-up grants Read to
ancestors so a member can navigate down to their unit; fan-down grants the row's real level
to descendants so a department head can act on every unit.

Rules:

- **Leaf rows never fan.** A unit is a leaf and has nothing beneath it, so unit isolation is
  preserved by construction rather than by a special case that could be got wrong.
- **The library rule still applies to a fanned grant** (§5). A department `_APR` row reaches
  Staging only; a department `MEMBER` row reaches Documents only.
- **Idempotent**, like every other assignment — `addroleassignment` merges.
- **Fanned grants are logged distinctly** (`↳↓`), because "why does this group hold
  DMS Approve on a unit folder that has no row for it" is otherwise unanswerable from the
  log, and an unexplained grant is indistinguishable from a bug.

### 4.1 Deeper chains than [Department, Unit]

Today every onboarded segment has `Levels = [Department, Unit]`, so a unit's only non-leaf
ancestor **is** its department, and "fan down from any ancestor" and "fan down from the
department" are the same rule. The 2027 segments have deeper chains (Region →
Estate/Mill), where a Region-tier row would reach everything below it.

That is the honest reading of what a Region-tier row means, so it is the implemented
behaviour — but it is a wider grant than anyone will picture from the phrase "departmental
fan-out", and it must be reviewed when those segments are onboarded rather than discovered
afterwards. Recorded here and in `2026-07-21-segment-onboarding-plan.md`.

### 4.2 Fan-out is opt-in, and defaults OFF

Added during implementation, against the original intent, because of historical data.

The plan above assumed no department-tier rows exist, since the model is leaf-only. That is
true of rows created *today* — but until 2026-07-29 the old `isChainAuthorized` **required a
`MEMBER` row at every tier**, so any site provisioned before then can still carry leftover
segment- and department-tier `MEMBER` rows that nobody remembers creating.

Fanning one of those would silently grant Read on **every unit folder in Documents** —
precisely the cross-unit leak the isolation model exists to prevent, and invisible because
nothing in the UI shows a row's blast radius.

So a `recon_departmentFanOut` setting row on `DMS Config` gates it:

| Value | Behaviour |
| --- | --- |
| absent / anything else | **off** — every fanned grant is **reported** and none is made |
| `on` | grants applied |

The run states the mode as its first log line, because "off" is the state in which a Head
of Department silently does not work, and that must be visible rather than inferred from an
absence of grants. The intended sequence is: reconcile, read the `would inherit …` warnings,
delete the leftovers, then switch it on.

**Segment-tier rows never fan at all**, at any setting. A segment-tier leftover is the
widest possible accident, and unlike a department row there is no persona that wants one.

## 5. The library rule, restated

`FolderManager.tsx` enforces that Staging receives only `UPL`/`APR` and Documents only
`MEMBER`, which is what keeps viewers off pending documents. `DEL` joins the **Documents**
side. Expressed as a table rather than the current inline ternary, because a two-way
condition that must now express three cases is exactly where the next isolation bug goes.

| Library | Roles accepted |
| --- | --- |
| Staging | `UPL`, `APR` |
| Documents | `MEMBER`, `DEL` |

## 6. Upload authorization must follow the fan-out

The Form resolves a user's permitted upload paths from their Group Map rows, walks each
row's term ancestry, and requires the chain to terminate at that term
(`isLeafChainValid`). A department-tier `_UPL` row breaks this: the chain terminates at the
department, the Unit dropdown has no options, and Head of Department #1 cannot upload —
while their folder permissions say they can. A permission model whose two halves disagree
is worse than either half alone.

So `resolveValidPaths` gains the mirror of §4: **a row on a non-leaf term expands to every
descendant leaf**, one valid path each. A row on a leaf behaves exactly as it does now.

This extends leaf-only authorization rather than retreating from it. The rule removed on
2026-07-29 demanded a `MEMBER` row at *every* tier and refused a correctly provisioned
uploader; this adds only the paths a deliberate department-tier row already grants, and
never infers a row that is not there.

## 7. Out of scope

- **`_HC` and the two HC libraries.** Highly Confidential left Phase 1 on 2026-08-01 and the
  term is deleted from the term store, so the level cannot be selected at all.
  `feat/hc-libraries` stays parked; its §14 lists the four steps to restore it, two of which
  fail silently.
- **Revoking a fanned grant.** Nothing in this system removes a role assignment, so deleting
  a department row leaves its descendant grants in place. Same posture, and same reason, as
  `2026-08-03-access-scope-mapping-design.md` §7.
- **Head of Department across segments.** A department row fans down its own subtree only.

## 8. Verification

Needs the tenant. Both custom permission levels must exist first.

1. **Head-of #2** — `{Unit}` + `_APR`, no `_UPL`. Approve a pending item: works. Try to
   upload: refused. This is the check that proves `DMS Approve` took; on the old Design
   level the upload would succeed and everything else would look identical.
2. **Head-of #3** — add `_DEL`. Delete an approved document in Documents: works. Confirm
   they still cannot upload.
3. **Fan-out off (the default)** — create a department-tier `_APR` row and reconcile with
   `recon_departmentFanOut` absent. Confirm the log opens with `Departmental fan-out: off`,
   reports `would inherit DMS Approve on …` for each unit, and that the group holds **no**
   new grant on any unit folder. Then set the row to `on` for the next check.
4. **Head of Department** — one department-tier `_APR` row, no unit rows. Reconcile.
   Confirm the group holds DMS Approve on **every** unit folder under that department in
   Staging, and on **no** unit under a sibling department.
4. **Isolation** — the same run must not have granted that `_APR` row anything in
   Documents. Check by direct URL, not by browsing; browsing is security-trimmed and looks
   correct either way.
5. **Upload as Head of Department #1** — a department-tier `_UPL` row alone. The Unit
   dropdown offers every unit under that department and nothing else, and an upload lands in
   the chosen unit.
6. **Idempotency** — reconcile twice; the second run reports no new assignments.
7. **Before the levels exist** — confirm the run warns `no "DMS Approve" role definition on
   site` and skips that one assignment, rather than failing the whole reconciliation.

---

## 10. Approving does not require Documents access

Recorded because the code got this wrong twice in one afternoon and the wrong fix was
attractive both times.

The approval page checks that the destination folder in Documents exists and has unique
permissions before approving, so an approval cannot hand a document to the Auto-route flow
which would otherwise create the whole chain from the library root with the *library* ACL.
That check runs **as the signed-in approver** — SPFx has no privilege elevation.

But `APR` is Staging-only (§5), and reconciliation locks each Documents folder to site
Owners plus the unit's `MEMBER` group. So a correctly provisioned approver **cannot see the
destination folder**, and a guard that demands to read it refuses every one of them.

The tempting fix — put approvers in the unit's `MEMBER` group — is wrong. It widens
Documents access for every approver on the site to work around a defect in one function, and
it makes an unrelated group membership load-bearing for a button.

The states are distinguishable, because **the dangerous one is the visible one**:

| Folder state | Seen by an approver | Verdict |
| --- | --- | --- |
| locked (unique perms, approver not on it) | library grant does not reach it → item trimmed → `{"odata.null":true}` | **safe** |
| still inheriting (Auto-route made it) | library Read flows down → item readable → `HasUniqueRoleAssignments: false` | refuse |
| missing | `GetFolderByServerRelativeUrl` → 404 | refuse |

Inheritance is exactly what makes a folder visible, so a trimmed item on a folder that
*resolved* is positive evidence of unique permissions excluding the caller — proof of
locking, not absence of proof.

Assumes only that the approver can resolve the folder, which needs library-level Read from
the site-entry group. Without it they 404 and are refused, which is safe; the 404 message
names that cause alongside a genuinely missing folder.

> Note the recurring trap, third sighting in this codebase after `AllowedFileTypes`
> (CLAUDE.md #11) and `FieldValuesAsText`: **a value cannot distinguish "false" from
> "absent"**. `HasUniqueRoleAssignments` being `undefined` was read as `false` and reported
> as "not locked down" — a false statement about a folder's security, on a screen an
> approver is trusting.
