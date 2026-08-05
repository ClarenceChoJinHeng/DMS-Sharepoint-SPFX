# Visibility Scope + Per-Family View Groups — Design

**Date:** 2026-08-03
**Status:** DRAFT — five decisions open (§7). Nothing implemented.
**Branch:** `feat/folder-abbreviations`
**Parent:** `2026-08-03-role-personas-and-department-fanout-design.md` (the role model this
revises), `2026-08-03-access-scope-mapping-design.md`
**Supersedes in part:** the ancestor Read fan-up relied on by the parent spec §4

---

## 1. What the client's second restatement actually changes

The 2026-08-03 restatement of the group document adds three things to the role model that was
agreed this morning, and none of them is a role change:

1. **A visibility boundary per family, stated negatively.** "Head of Department **cannot see
   the Segment**." "Head of Unit **cannot see the department folder**." "PIC **cannot see
   segment and cannot see department**." This is new, and it contradicts what reconciliation
   does today.
2. **A fifth family, "Head of everything"** — sees a Segment and everything under it.
3. **Separate view groups per family**, and Documents-delete kept in its own group away from
   the Staging groups.

The role sets (`MEMBER`/`UPL`/`APR`/`DEL`) are unchanged. What changes is **which folders a
group is granted anything on**, which is a reconciliation concern, not a persona concern.

## 2. The visibility boundary is the whole job

### 2.1 What happens today

Reconciliation grants each group `Read` on **every ancestor folder on its own path** in the
same library — the "browse access" pass at `FolderManager.tsx` ~2132. Sibling folders get
nothing and SharePoint security-trims them, so a member sees a corridor: the segment folder,
their department folder, their unit, and no siblings at any level.

That corridor exists for one reason: **so a user can open the library and click down to their
unit.** It is the only navigation path the product has.

### 2.2 What the client now requires

| Family | Row tier | May see | Ancestor Read today | Required |
| --- | --- | --- | --- | --- |
| Head of everything | segment | segment and down | (library root only) | unchanged |
| Head of Department | department | department and down | segment folder | **none** |
| Head of Unit | unit | unit and down | segment + department | **none** |
| PIC | unit | unit and down | segment + department | **none** |
| SDG Employee | unit | own unit (Documents) | segment + department | **none** |

Every row reduces to the same rule, worth stating as one line because it is much simpler than
the table suggests:

> **A group is granted access on its own folder and its descendants, and on nothing above it.
> The ancestor Read fan-up is removed entirely.**

Head of everything needs no special case: a segment-tier row's only ancestor *is* the library
root, which is not a folder anyone was granted on anyway.

### 2.3 Removing the fan-up strands every user, and that must be fixed in the same change

This is the part that makes it more than deleting a loop.

With no Read above the unit, a user who opens the Staging library sees **an empty library**.
SharePoint auto-grants *Limited Access* on the parent chain when a child is assigned, so a
**direct URL** to the unit folder still opens — but Limited Access confers no *View Items* on
the parent, so nothing is listable and there is no path to click.

So the change is two-sided:

- **Remove** the ancestor Read pass.
- **Add** an in-product launcher: a "My folders" list of direct links to the folders the
  signed-in user actually holds. The data is already resolved — `resolveValidPaths` in
  `formModel.ts` walks the user's Group Map rows to their leaves, and the Folder Map carries
  each folder's live URL. A read-only list of links, not a new permission surface.

Shipping the removal without the launcher would present as "the DMS is empty" to every
non-admin on the site. The two halves are one change.

**A cheaper answer found 2026-08-04, and it needs no code at all.** Set the library view's
**Folders** option to **"Show all items without folders"**. A flat view lists *files* rather
than folders, and security trimming applies per file — so a user sees their own unit's
documents directly, with no folder to click and no folder *name* on screen. That is stricter
than the corridor it replaces, not looser: the segment and department folders are not merely
unreachable, they are absent from the page.

It does not remove the need for the launcher — a user still cannot navigate *to* a folder to
upload into — but it does make library access usable on day one without it.

**This is a view setting applied by the administrator, and deliberately NOT surfaced in the
Staging Library Access tab.** It is one change per library, done once. A permanent banner
telling an admin to do something they did weeks ago is noise, and noise is what trains people
to click past the banners that matter. It belongs in the provisioning runbook instead, which
is what this note is for.

### 2.4 Nothing in this system revokes a grant — so the boundary needs a revoke pass

Reconciliation only ever calls `addroleassignment`. Deleting the fan-up code stops **new**
ancestor grants; it does not remove the ones already on the segment and department folders of
every provisioned site. Those users keep seeing what the client asked to hide, and the run log
would report a clean pass while the requirement stays unmet.

So a **revoke pass** is required, and it is the first destructive operation in this codebase.
It follows the `recon_departmentFanOut` precedent exactly — report-only until deliberately
switched on:

| `recon_revokeAncestorRead` | Behaviour |
| --- | --- |
| absent / anything else | **off** — every candidate revocation is **reported**, none performed |
| `on` | `deleteroleassignment` performed |

Safety rails, all of them load-bearing:

- Only assignments whose level is exactly `Read` on a **non-leaf** folder are candidates. A
  department-tier group's real grant (`DMS Approve`, `Contribute`, `DMS Delete`) on its own
  department folder is never a candidate.
- **Never** touch site Owners, Full Control, or `DMS_SITE_MEMBERS`. Removing the site-entry
  group is how the whole site becomes unreachable.
- A group holding a legitimate **department-tier view row** keeps its department Read.
  Candidate selection is therefore "Read held by a group whose own row sits *below* this
  folder", not "all Read on this folder".
- The run states the mode as its first log line and reports every revocation with the group,
  the folder, and the row tier that justified it. An unexplained revocation is
  indistinguishable from a bug, same as an unexplained grant.

## 3. Head of everything

A new family, one row: a **segment-tier** view row that reaches every department and unit under
that segment.

The parent spec §4.2 states **segment-tier rows never fan, at any setting** — deliberately,
because a leftover segment-tier `MEMBER` row from before 2026-07-29 is the widest possible
accident and no persona wanted one. This family is that persona, so the guard cannot simply be
dropped: a leftover row and a Head of Everything row are indistinguishable by tier alone.

They are distinguishable by **role**. So Head of Everything gets its own role rather than a
relaxed guard:

| Role | Tier | Library | Level | Fans |
| --- | --- | --- | --- | --- |
| `SEGVIEW` | segment only | Documents | Read | yes, from the segment |

`MEMBER` keeps its guard — a segment-tier `MEMBER` row still never fans, so every historical
leftover stays inert. Only a row an admin created *deliberately as* `SEGVIEW` fans from the
segment. Same reasoning as `DEL` being its own role rather than a wider `Contribute`:
capability is carried by an explicit name, never inferred from position.

Open: whether Head of Everything also approves or uploads, and whether they see Staging pending
items or only approved Documents (§7 D1). The table above assumes **view-only, Documents
only**, the narrowest reading of "can see Segment and onward".

## 4. Group inventory after this change

Four atomic groups per tier that has a head, none shared between families — which is the "do
not mix group" requirement, and it falls out of the tier being part of the group name.

| Family | Tier | Groups |
| --- | --- | --- |
| Head of everything | segment | `DMS_<SEG>_SEGVIEW` |
| Head of Department | department | `DMS_<SEG>_<DEPT>` (view) · `_UPL` · `_APR` · `_DEL` |
| Head of Unit | unit | `DMS_<SEG>_<DEPT>_<UNIT>` (view) · `_UPL` · `_APR` · `_DEL` |
| PIC | unit | reuses the unit view group + `_UPL` |
| SDG Employee | unit | reuses the unit view group |

The four client groups remain **memberships, not groups** (parent spec §2), identically at both
tiers:

| Group | Roles |
| --- | --- |
| 1 — approve + upload | view + `UPL` + `APR` |
| 2 — approve only | view + `APR` |
| 3 — approve + delete | view + `APR` + `DEL` |
| 4 — approve + upload + delete | view + `UPL` + `APR` + `DEL` |

**Cost, stated because it is the client's provisioning work.** GHO today is 2 departments / 12
units. Before: 36 groups. After: 12 × 4 + 2 × 4 + 1 = **57**. Create `_DEL` and `_APR` only
where a real person needs them; the inventory is a ceiling, not a checklist.

## 5. Documents-delete stays in its own group

Already true and unchanged — recorded because the client asked for it explicitly. `DEL` maps to
`DMS Delete` on **Documents**; `UPL`/`APR` map to Staging levels. The library rule in
`FolderManager.tsx` enforces the split, so a `_DEL` group cannot receive a Staging grant even
if someone writes the row at the wrong tier:

| Library | Roles accepted |
| --- | --- |
| Staging | `UPL`, `APR` |
| Documents | `MEMBER`, `SEGVIEW`, `DEL` |

## 6. What cannot be enforced, and the sentence to give the client instead

Five lines of the client's document describe restrictions SharePoint cannot express. Each is
listed with what the platform *does* deliver, because a permission level that claims a
restriction it does not enforce ends up in a handover document as a security guarantee.

**"Cannot Edit files on staging" (both families, groups 2 and 3).** *Approve Items* depends on
*Edit Items*; unticking Edit silently clears Approve, because approving **is** an edit (it
writes `_ModerationStatus`). → *They cannot add a file or delete one — only move one through
the workflow.*

**"PIC 3 can see their own files only."** An uploader seeing only their own pending items is
**Draft Item Security**, a *library-wide* setting — it cannot differ between PIC 1 and PIC 3 in
the same library. The current value is deliberately "Any user who can read items", because the
stricter one hides pending **folders** from readers and 403s Contribute uploaders (memory
`dms-content-approval-blocks-uploader`). → *Enforceable in the web part's pending list only
(filter by `Author`); the native library view still shows the unit's items.*

**"PIC can delete only their own files."** *Delete Items* is **folder-scoped, not
author-scoped**. Anyone with Contribute on a unit folder can delete a colleague's pending file
by direct URL. → *No per-author delete exists at any permission level.*

**"If any PIC wants to delete they need approval from the head of unit first."** A process rule
with no technical hook while PIC holds Contribute. → *Policy only, unenforced — unless PIC
loses Staging delete entirely (§7 D4).*

**"Head of Department 4 — delete files on staging unit level only."** Delete rides the same
Contribute grant as Add, and that grant fans across the department, so delete cannot be
narrower than upload while `UPL` is one role at one level. → *Their Staging delete reaches
every unit they can upload to.*

## 7. Decisions open

Nothing is implemented until these are answered; D2 in particular changes what users see on day
one.

- **D1 — Head of everything.** View-only in Documents (as drafted), or also approve/upload?
  Does "Segment" mean one segment per person, or all twelve — which is the existing `GLOBAL`
  role, and a different thing?
- **D2 — Ancestor visibility.** Confirm that losing in-library click-down navigation is
  acceptable in exchange for hiding the segment/department folders, with a "My folders" link
  list as the replacement. And: revoke the ancestor Read already granted on existing sites
  (destructive, report-first), or only stop granting it from now on?
- **D3 — PIC 3.** Enforce "own files only" in the web part alone, accepting that the native
  library view shows the unit's items, or drop the requirement?
- **D4 — PIC delete.** Keep it (the approval rule is policy only), or remove Staging delete
  from PIC — which needs the delete-free upload level explicitly decided against earlier today?
- **D5 — Head of Department 4.** Accept that Staging delete follows the upload fan-out, or
  split `UPL` into two roles so delete can stay at unit tier?

## 8. Implementation order

Each step is independently shippable and independently verifiable. Nothing before D1–D5.

1. **`SEGVIEW` role** — add to `GroupMapRole`, `SELECTABLE_ROLES`, `roleFromGroupName`
   (`_SEGVIEW` suffix), `ROLE_TO_PERMISSION`, the library rule, and the segment-tier fan guard.
   Add the Head of everything persona. Pure model work, unit-testable, no tenant.
2. **Remove the ancestor Read pass** + **"My folders" launcher**. One change, two halves
   (§2.3).
3. **Revoke pass** behind `recon_revokeAncestorRead`, report-only by default (§2.4).
4. **Personas + picker** — department- and unit-tier variants surfaced per family, so an admin
   cannot pick a Head of Unit at department tier by accident.
5. **Web-part-side PIC 3 filter**, if D3 says so.

## 9. Verification

Needs the tenant, a non-admin test account per family, and both custom permission levels
already created.

1. **Head of Unit sees no department folder.** Sign in as a unit-tier member, open Staging.
   Confirm the library lists nothing, the "My folders" link opens the unit folder, and a
   hand-typed URL to the *department* folder is refused. Check by URL, not by browsing —
   browsing is security-trimmed and looks correct either way.
2. **Head of Department sees the department, not the segment.** Same test one tier up.
3. **Head of everything** — one `SEGVIEW` row, no others. Confirm Read on every unit folder
   under that segment in Documents, **nothing** in Staging, and nothing under a sibling segment.
4. **Leftover rows stay inert** — a segment-tier `MEMBER` row must still fan nowhere, with
   `SEGVIEW` shipped and fan-out ON. This is the check that proves the guard survived.
5. **Revoke pass, off** — confirm the log opens with the mode, reports each candidate, and that
   no assignment was actually removed.
6. **Revoke pass, on** — confirm `DMS_SITE_MEMBERS` and site Owners are untouched, and that a
   department-tier group keeps its own department-folder grant.
7. **Uploading still works** after the revoke — the Form ensures `Year`/`Document Type`
   subfolders under the unit folder, which needs the unit grant only, but this is the path
   Limited Access has to carry and it must be proven, not assumed.
8. **Approval still works** after the revoke — the approval guard resolves the *Documents* unit
   folder as the signed-in approver (parent spec §10). Confirm the three-state verdict still
   reads correctly with no ancestor Read in play.
9. **Idempotency** — reconcile twice; the second run reports no new assignments and no new
   revocations.
