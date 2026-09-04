# Requests page — approvers only, and a Head of Department joins them

**Date:** 2026-08-21 · **Status:** designed, implementing
**Supersedes:** the audience half of `2026-08-15-deletion-and-share-requests-design.md`

Client, 2026-08-21: *"make a page for HOU where they can handle the request of deletion and share
from the Uploader themselves which is PIC, that page should only allow HOU and HOD to enter, the
reason I also include HOD is because they literally have Share and Deletion power."*

## 1. What is wrong today

`pageAccessPolicy.ts` grants `Requests.aspx` to **`UPL` and `APR`**, and the comment above the rule
says why: the page was designed with two audiences, the uploader raising a request and the Head of
Unit deciding it.

**That reasoning went stale on 2026-08-20**, when the requester's own view moved to **My Submissions
→ Requests tab** (with Cancel). The uploader's half of this page no longer exists. A PIC who opens
it today gets a screen filtered to units where they hold `APR` — none — so an empty page.

Meanwhile the **Head of Department**, who since 1.0.197.0 holds `DEPTVIEW + DEL + SHARE` and can
therefore delete and share approved documents outright, cannot open the page at all.

## 2. The change

| | Was | Now |
|---|---|---|
| Page policy | `["UPL", "APR"]` | **`["APR", "DEPTVIEW"]`** |
| Who decides | units where the viewer holds `APR` | those, **plus** units under a department the viewer heads |
| HoD sees | nothing | their department's **approved-stage** requests only |

## 3. ⚠ A HEAD OF DEPARTMENT CANNOT ACTION A PENDING-FILE REQUEST, AND THIS IS THE WHOLE DESIGN

`hod` is `DEPTVIEW + DEL + SHARE`. **None of those three appears in `LIBRARY_ROLES.Staging`** — so a
Head of Department holds nothing whatsoever in either approval library. They cannot read a pending
file, let alone recycle one.

And **the approval executes in the approver's own browser session** (2026-08-15 design, unchanged):
pressing Approve recycles the file as *them*. So a HoD approving a pending-stage deletion would fail
in their own session, `applyDecision` would honestly record `Failed`, and the requester would get
nothing — having been told their request was being handled.

**Pending-stage requests are therefore HIDDEN from a Head of Department entirely** (client's choice,
2026-08-21, over showing-them-disabled). A Head of Unit always exists for a unit that has an `APR`
mapping, and that is who the request was routed to in the first place — so nothing is stranded by
hiding it. The alternative, letting them press a button that cannot work, is the one this codebase
rejects everywhere else.

Consequence to state plainly: **`DELS` has exactly one holder (`hou`)**, so a pending-file deletion
has exactly one person who can carry it out. That is not a gap this page can close.

## 4. Scope is DERIVED, and the join needs no new data

A Group Map row already carries both halves (`groupMapModel.ts` ~779):

- `Segment` = the segment's **term set GUID**
- `UnitTermGuid` = the tier's term GUID — the **department** on a `DEPTVIEW` row

So a department expands to its units with **one** request per department headed:
`/_api/v2.1/termStore/sets/{Segment}/terms/{UnitTermGuid}/children?$select=id`.

No config read, no new column, no migration. A person heads one department in the normal case, so
this is one extra request at mount.

**Rejected: storing the department GUID on the request row.** It is a new column, every row written
before today would lack it, and the `Stage` column two days ago already showed what that costs — a
read that must be retried without the field, a write that must be retried without the field, and a
provisioning gap to detect. Ancestry is derivable; a stored copy would also describe a shape that a
re-parented term has since left.

## 5. Failure directions

- **The department expansion fails ⇒ the HoD sees nothing extra.** Fails CLOSED, matching the
  existing rule that an unreadable Group Map yields an empty queue rather than a broken page. An
  over-wide queue would show one department's requests to another's head.
- **It is reported, not silent.** A Head of Department told "no requests" when the expansion failed
  would stop looking. The page says the department scope could not be read, naming the status.
- **`aprUnits` is unchanged and unaffected** — a HoD who is also a Head of Unit somewhere keeps that
  queue in full, at every stage. The two sets are a union, never an override.

## 6. Rules live in `shared/requests.ts`, pure and tested

`ViewerScope { aprUnits: string[]; hodUnits: string[] }`, and `canDecide` / `isVisibleTo` / `queueFor`
take it. A bare `string[]` is still accepted and read as `aprUnits` — that is exactly what the array
has always meant, so every existing caller and test stays correct rather than being rewritten around
a new shape.

`canDecide`: Pending, then `APR` match at any stage, or `DEPTVIEW` match at **approved** stage only.
`isVisibleTo`: own row, or the same two conditions — hidden means hidden, not merely un-actionable.

## 7. Migration

**Re-run reconciliation.** Page grants are asserted in full at page scope, so the run removes every
`_UPLOADER` group from `Requests.aspx` (naming each removal) and adds the `_HOD` groups. No row
change, no schema change.

## 8. Out of scope

- **The `CRS Requests` LIST still inherits site permissions.** This page filters what it *renders*;
  anyone who can open the site can read the list at its own URL — request reasons, recipients and
  decision notes across every unit. Raised with the client 2026-08-21; not addressed here.
- Revoking an approved share (spec §5.3, still unbuilt).
