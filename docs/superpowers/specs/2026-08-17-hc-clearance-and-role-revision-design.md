# HC clearance narrowed to two roles, and Head of Department made view-only

**Date:** 2026-08-17
**Status:** BUILT — `PERSONAS` and `LIBRARY_ROLES` carry it. Partially site-verified 2026-08-19: the from-scratch run granted only `_APPROVER` and `_UPL_HIGHLY_CONFIDENTIAL` on `HC Approval Document`, with no plain uploader or employee on either HC library. An HC upload has still not been tested end to end. Header corrected 2026-08-19.
**Revises:** `2026-08-15-highly-confidential-library-design.md` (the library-pair architecture is
UNCHANGED and still governs), `2026-08-07-role-model-simplification-design.md`,
`2026-08-09-persona-driven-folder-access-design.md`.

---

## 1. What the client changed

Three instructions, given while provisioning their site:

> "any approver which is HOU can see Highly Confidential files as well. So we do not need a dedicated
> HOU but rather a dedicated PIC who is an uploader only and dedicated viewer who is SDG Employee."

> "HOU have the power to delete, PIC can request for deletion. HOD no need deletion power, he only
> view."

> "Head of Department do not have share functionality that is HOU, just needs View. HOU uploads to HC
> yes, any type of HOU can upload to HC, so we do not need a separate HOU for HC"

**This supersedes their 2026-08-15 instruction** that *"only for Approver HOU and PIC uploader needs a
separate group"*. Recorded rather than silently replaced: that earlier rule is why `APRHC` and the
`hou_hc` persona exist at all.

**The resulting principle: HC clearance is a dedicated group for the two roles at the bottom of the
hierarchy — the uploader and the plain viewer. Every management role reaches HC through the role it
already holds.**

Worth stating to the client, because it reads oddly: a Head of Unit needs no HC clearance, an SDG
Employee does. HC is therefore narrower than normal documents only at the bottom, not the top.

---

## 2. The model

| Persona | Roles | Reaches HC by |
|---|---|---|
| C-Level (all segments) | `GLOBAL` | role |
| C-Level (one segment) | `SEGVIEW` | role |
| **Head of Department** | **`DEPTVIEW`** | role |
| Head of Unit | `APR`, `DEL`, `DELS`, `SHARE`, **`UPLHC`**, **`DELSHC`** | role |
| PIC | `UPL`, `DELS` | — |
| PIC (HC) | `UPLHC`, **`DELSHC`** | dedicated group |
| SDG Employee | `MEMBER` | — |
| **SDG Employee (HC)** | **`MEMBERHC`** | dedicated group |

Retired: `APRHC`, and the `hou_hc` persona.

### 2.1 Library access

```
Approval Document :  UPL · APR · DELS · UPLHC · DELSHC
Documents         :  MEMBER · DEPTVIEW · GLOBAL · SEGVIEW · UPL · APR · DEL · SHARE · UPLHC
HC Approval       :  UPLHC · DELSHC · APR
HC Documents      :  UPLHC · MEMBERHC · APR · DEL · DEPTVIEW · GLOBAL · SEGVIEW · SHARE
```

**`DEL` is deliberately NOT on `HC Approval`** — see §3.4, which corrects an earlier draft of this
spec.

`MEMBER` and `DELS` are **absent from both HC rows**, and that absence is the entire feature — it is
what makes a plain PIC and a plain SDG Employee unable to reach HC at all.

---

## 3. Three collisions this had to work around

Each is a case where the obvious implementation hands HC to people the client excluded. All three have
the same shape: **a role held by two personas cannot grant to one of them and not the other.**

### 3.1 A Head of Unit uploads HC — but NOT by putting `UPL` on the HC library

`UPL` is held by the plain PIC as well. Listing it on the HC approval library would give **every PIC**
HC upload and destroy the dedicated HC uploader group the client just asked for.

**Instead the `hou` persona itself gains `UPLHC`.** Same capability, no leak, no new role. A Head of
Unit group carries an HC upload grant; a PIC group does not.

### 3.2 A Head of Department stays view-only — but NOT by swapping to `MEMBER`

`MEMBER` is the SDG Employee role, and it is deliberately absent from `HC Documents`. Reusing it for HoD
would either lose HoD's HC read, or — if `MEMBER` were added to `HC Documents` to fix that — give
**every SDG Employee** HC read, destroying the dedicated HC viewer group.

**Hence `DEPTVIEW`**: Read, department scope, fanning down to every unit beneath, reaching `Documents`
and `HC Documents`. What `SEGVIEW` is to a segment, one tier down.

### 3.3 PIC (HC) deletes their own pending HC files — but `DELS` cannot be on the HC library

**This one is a live bug, not merely a design constraint.** `DELS` is currently listed on the HC
approval library, and **the plain PIC holds `DELS`** — so today every plain PIC holds `CRS Delete` on
the HC approval folder.

Narrow in practice: Draft Item Security shows a non-approver only their own items, and a plain PIC has
no HC files. But an **approved** HC file still sitting in that library before Auto-route moves it is
visible to every reader, and they would hold delete on it.

**Hence `DELSHC`**, and `DELS` comes off the HC approval library.

Which leaves the table symmetric — every HC-clearance role is the plain role's twin:

| Capability | Plain | HC |
|---|---|---|
| Upload | `UPL` | `UPLHC` |
| Read approved | `MEMBER` | `MEMBERHC` |
| Delete own pending | `DELS` | `DELSHC` |

### 3.4 …and one collision that resolves itself

**An earlier draft of this spec was wrong here, and the tests caught it before it shipped.** The claim
was: once HoD loses `DEL`, that role is held *only* by Head of Unit, so `DEL` becomes safe on both HC
libraries — giving a HoU HC delete without reintroducing 3.3.

**`DEL` is not HoU-exclusive.** `clevel_global` is `["GLOBAL","DEL","SHARE"]`, and `clevel_segment` the
same with `SEGVIEW`. So `DEL` on the HC **approval** library would have given a C-Level read and delete
on **unapproved HC drafts** — breaking the oldest rule in this model, that no view role touches an
approval library, and breaking it in the one library where it matters most.

**So `DEL` stays off `HC Approval`, and a Head of Unit's HC pending delete comes from `DELSHC`**, which
they now carry alongside `pic_hc`. `DELSHC` is held by exactly those two unit personas, and that is what
makes it safe where `DEL` is not. `DEL` remains on `HC Documents`, where C-Level reaching it is
pre-existing and intended.

**The transferable lesson:** *"role X is now exclusive to persona Y"* is a claim about the whole
`PERSONAS` array, not about the entry being edited. Both `DEL`'s and `DELSHC`'s holders are pinned as
exact sets by test for precisely that reason.

---

## 4. `DELS` needs no "own files only" rule

Unchanged from 2026-08-15, restated because the client asked directly.

`DELS` is approval-library-only, and Draft Item Security already hides a peer's pending work from a PIC
— so **"delete what you can see" IS "delete your own"**. The permission and the visibility are the same
boundary, so there is nothing extra to enforce and nothing that can drift apart. `DELSHC` inherits the
property for free.

> ⚠ **SUPERSEDED 2026-08-19**, and `DELSHC` inherits that too: a cleared PIC deletes any pending HC file
> in their unit, not only their own. The HC boundary is untouched — an uncleared uploader still reaches
> no HC library at all.

The exception, in both libraries: an **approved** file is visible to every reader until Auto-route moves
it. That window is why the delete inside Auto-route is load-bearing for security rather than
housekeeping.

---

## 5. Group naming

Suffixes follow the existing pattern — a short and a long form, long form suggested:

| Role | Suffixes |
|---|---|
| `DEPTVIEW` | `_DEPTVIEW`, `_DEPARTMENT_VIEWER` |
| `MEMBERHC` | `_VIEWER_HC`, `_VIEWER_HIGHLY_CONFIDENTIAL` |
| `DELSHC` | `_DELS_HC`, `_DELETE_PENDING_HIGHLY_CONFIDENTIAL` |

**`MEMBERHC` cannot use `_MEMBER_HC`**, because `MEMBER` is the *base* group and carries no suffix at
all (`suggestGroupName` returns `GHO_Group Finance` for it). Hence `_VIEWER_HC`.

**Precedence comes from the existing longest-first sort**, so `_DELS_HC` (8) beats `_DELS` (5) beats
`_DEL` (4), and `_VIEWER_HC` (10) beats `_HC` (3). Nothing new is needed — but nothing may be added
that a shorter suffix could swallow.

**An existing `_APR_HC` group name resolves to plain `APR`.** The suffix stays mapped rather than
deleted, so a group someone already created keeps working and means what it now means.

---

## 6. Migration

**No list schema change and no row rewriting**, with two exceptions an administrator must action:

1. **Any existing Head of Department mapping must be re-created with the persona**, because its rows say
   `DEL`/`SHARE` and must say `DEPTVIEW`. A stale `DEL` row keeps granting delete — it does not fail, it
   silently retains the power the client removed.
2. **Any `_APR_HC` group is now redundant.** It resolves to `APR`, which already reaches HC, so it grants
   nothing extra. Harmless, but delete it so nobody maintains it.

**Reconciliation must be re-run.** Nothing changes on a folder until it is — including the removal of
HoD's delete, which is the one change here that *reduces* access and therefore has a clock on it.

---

## 7. What is NOT changing

- The **library pair architecture**. HC is a separate library pair, not a secured subfolder. The
  2026-07-16 elevated-flow design stays superseded: an uploader cannot break inheritance, so securing a
  subfolder needed a service-account flow and left a window where a peer could see the file.
- **The unit is still the smallest confidentiality boundary.** Nothing here changes that two people in
  one unit see each other's approved documents.
- **`UPLHC` remains a superset** — it covers the normal approval library too, so an HC-cleared PIC needs
  one group, not two.
- **Both HC libraries must still resolve or neither does**, and `naming.ts` still has no HC fallback.
- **The two HC Power Automate flows are still required** and nothing HC works without them.

---

## 8. Open

- **Which units need HC at all**, and who is cleared. Only those units need the two HC groups.
- **`hcConfidentialityLevel`** stays blank until the HC groups exist, so `Highly Confidential` is offered
  but refuses — safe, because it fails closed, but visible. Either finish the HC groups or remove the
  term until then.
- **Migration of documents already labelled Highly Confidential** remains out of scope. Until a sweep
  runs, HC protection applies only to documents filed after deployment.
