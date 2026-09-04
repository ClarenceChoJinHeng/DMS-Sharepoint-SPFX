# Deletion and share requests — design

**Status:** BUILT — web part `CRS Requests` (`7a4f1e93-…`). **NOT site-tested**, and nothing works until the prerequisites exist: the folder-scope grants and the `CRS Share` permission level, then a reconciliation re-run. Header corrected 2026-08-19.
**Date:** 2026-08-15
**Supersedes:** `2026-08-04-deletion-request-approval-design.md` (DRAFT). Its premise was that a PIC
loses `Delete Items` on **Staging**; the client reversed that on 2026-08-15 — a PIC *does* delete in
the approval library. The workflow it described survives, moved to `Documents`, and gets simpler (§3).
**Required reading before touching §5:** `2026-07-23-share-guard-retirement.md`.

**Client, 2026-08-15:** *"PIC is suppose to delete in staging and HOU can upload. PIC in document
library on the other hand cannot delete but can request for a deletion, HOU will be the one who
approves the deletion. HOD and C Level do not need permission to delete files. PIC in document library
can share but if they want to share a file it needs to go through the HOU to approve… HOD and C level
can share without approval."*

---

## 1. The corrected role model

The 2026-08-07 six-persona model had two rules backwards. Corrected:

| Persona | Approval Document | Documents |
|---|---|---|
| PIC | upload **and delete** (own pending/rejected) | read · **request** deletion · **request** share |
| Head of Unit | **upload**, approve, delete pending/rejected | read · **delete** · **share** · approves both request types |
| Head of Department | none | read · delete · share — **no approval needed** |
| C-Level (segment / global) | none | read · **delete** · **share** — no approval needed |

**PIC delete on the approval library needs no scoping rule.** Draft Item Security already hides other
people's pending and rejected files from a PIC, so "delete what you can see" *is* "delete your own".
The permission and the visibility are the same boundary — there is nothing extra to enforce.

> ⚠ **SUPERSEDED 2026-08-19**, when draft security was opened to *"Any user who can read items"*. A PIC
> deletes any pending or rejected file in their unit, not only their own. Still confined to the approval
> library and to the unit.

> ⚠ **THE HoD ROW ABOVE IS SUPERSEDED 2026-08-17.** Head of Department is **view-only** (client: *"HOD
> no need deletion power, he only view"*). C-Level's row stands, and since 2026-08-19 its delete and
> share reach every folder in the segment rather than the segment folder alone.

**A Head of Unit approves their own uploads** (client's choice, asked explicitly). Recorded because it
is a real control gap: the one person who most often files documents can skip the approval step. The
alternatives were a deadlock in any single-HoU unit, or giving Heads of Department read access to every
draft in the department. Named here so it stays a decision rather than becoming an oversight.

**C-Level gains delete**, moving from view-only to delete across a whole segment — or, for `GLOBAL`,
across everything. That is the widest single grant in this change set.

**`SEGVIEW`/`GLOBAL` still never appear in `LIBRARY_ROLES.Staging`.** Unchanged and non-negotiable: a
segment-wide viewer there would read every unapproved draft in the segment.

---

## 2. Two requests, one shape

A deletion request and a share request are the same object: a PIC asks, a Head of Unit decides, and on
approval something happens. **One list, one page, a `RequestType` column.** Built separately they become
two screens that drift, with two audit trails and two ideas of what "pending" means.

### 2.1 The list — `<P> Requests`

Self-provisioned by the page, like the audit log — the client cannot run PowerShell, so a scripted step
would simply not happen.

| Column | Type | Notes |
|---|---|---|
| `RequestType` | **Text** | `Deletion` \| `Share`. **Never Choice** — a value absent from a Choice column fails the whole write, so the day someone adds a type in code, every row of that type is lost silently (audit-log lesson). |
| `Status` | **Text** | `Pending` \| `Approved` \| `Rejected` \| `Failed`. Same reasoning. |
| `ItemUniqueId` | Text, indexed | The file's `UniqueId` — survives rename and move, which a URL does not. |
| `ItemUrl`, `ItemName` | Text | For display and for the audit row. Never the lookup key. |
| `Segment`, `Unit` | Text | Which approver's queue this belongs in. |
| `RequestedBy` / `RequestedAt` | Text / DateTime | `RequestedAt` written **ISO** — a plain `/items` POST goes through the OData layer and rejects a locale string (audit-log lesson; gotcha #1 does NOT apply to this endpoint). |
| `Reason` | Note | Required. An approver deciding blind will just approve. |
| `ShareWith` | Text | Share only. One or more addresses. |
| `SharePermission` | Text | Share only: `View` \| `Edit`. |
| `ExpiresAt` | DateTime | Share only, optional — see §5.3. |
| `DecidedBy` / `DecidedAt` / `DecisionNote` | Text / DateTime / Note | |

**No item-level security on this list**, deliberately. Hiding a PIC's request from their peers buys
nothing: every PIC already reads every approved document in their unit (2026-08-09), so a request row
reveals no filename they could not already see. Page permissions are the control.

---

## 3. The principle that makes this work

**The approval executes in the approver's own browser session.**

When a Head of Unit clicks Approve, *their* session performs the delete or the grant, because they hold
the rights the PIC lacks. Consequences, all of them good:

- **No service account and no Power Automate.** The superseded spec routed deletions through a flow with
  its own credentials — an extra moving part that stops silently when a password changes (the Auto-route
  lesson), and one more thing the client cannot debug.
- **The audit trail is honest.** The person who performed the action *is* the person who authorised it.
  Nothing acts on anyone's behalf.
- **It cannot exceed the approver's rights.** If a HoU cannot delete a file, the approval fails loudly
  rather than escalating privilege to get it done.

The cost: an approval only completes while the approver is on the page. A request approved against a
file that has since moved or been deleted fails, and says so, rather than half-succeeding.

---

## 4. Deletion requests

1. PIC raises a deletion request, with a reason, from the file's row on **My Submissions** or the
   Requests page.
2. The HoU sees it queued, with file name, unit, requester and reason.
3. On **Approve**, their session deletes the file — **recycle, never purge**, so it is restorable for 93
   days. The dialog says so; that is what makes approving a deletion a reasonable thing to do at all.
4. On **Reject**, the row closes with a note and the file is untouched.

**Requires:** Head of Unit gains `Delete Items` on `Documents` — narrow, and the only permission this
half of the change needs.

**A file already gone is not a failure to hide.** If it cannot be found, the request closes as `Failed`
naming why. "Approved" against a file that no longer exists teaches the approver the screen is lying.

---

## 5. Share requests

### 5.1 Read the retirement spec first

`2026-07-23-share-guard-retirement.md` retired a custom sharing gate after finding SharePoint already
provides one: with *"Only site owners can share files, folders, and the site"* enabled, a member's share
attempt is **not granted** — it becomes an Accept/Decline request to the site owner. Verified live
2026-07-21.

**Its premise still stands: a web part cannot intercept the native Share button.** So this page is not a
gate. It is the *sanctioned route*, and the native lockdown remains the backstop for anyone who presses
Share instead — that request simply lands with an administrator rather than the Head of Unit.

Stated plainly because the obvious future question — *"can we force everyone through our page?"* — has
already been answered once: no.

### 5.2 What it costs

**A Head of Unit must gain sharing rights over their unit folder** (client's decision, 2026-08-15),
because an approver can only approve what they can perform. Granting access to a file requires
`Manage Permissions` on it, which a HoU does not have today.

Consequences to state to the client rather than bury:

- A HoU can then **share directly**, with no screen involved. They become the sharing authority for
  their unit, not merely an approver of one.
- They can grant **any** access within their unit, to anyone the tenant allows.

HoD and C-Level need the same rights at department and segment scope in order to share without approval.

### 5.3 Executing a share

- **Internal recipient** — break inheritance on the file and add a role assignment. Direct, revocable,
  and visible in the file's own permissions.
- **External recipient** — a sharing link, which is a different API and additionally requires external
  sharing enabled at **tenant and site** level. Our switch overrides neither; if SDG's IT has it off,
  external requests must fail saying which layer refused, rather than appearing to work.

**External sharing is a `DMS Config` row, not a constant** (client: *"for now External recipients too,
incase client says no we can off it"*). Absent or unreadable ⇒ **internal only**. That is a deliberate
exception to this codebase's fail-open rule, for the same reason `canOfferFolderDelete` fails closed:
everywhere else the cost of a failed read is a form out of service, but here it is a document leaving
the organisation on the strength of a setting nobody could confirm.

**Every approved share creates a unique permission scope on that file.** That is the mechanism that hit
SharePoint's 50,000-scope ceiling when per-uploader ACLs were considered (2026-08-06). Shares are rare
where per-file ACLs would have been universal, so it is workable — but it is the same ceiling, so:

- shares carry an **optional expiry**, and
- the Requests page lists live shares with a **Revoke** action, which removes the scope.

Without a way back out, the count only ever grows.

> ✅ **BUILT 2026-08-28** — see `2026-08-28-requests-page-redesign-and-share-revoke-design.md`. Note
> the correction it carries: only **"Revoke all"** (`resetroleinheritance`) reclaims the scope;
> per-recipient revocation removes the person and leaves the scope in place, so it does not on its own
> solve the problem this paragraph is about.

---

## 6. What changes in code

| Area | Change |
|---|---|
| `shared/groupMapModel.ts` | `PERSONAS` — PIC gains `DELS`; HoU gains `UPL`; C-Level gains delete on Documents; a new role carrying sharing rights. |
| `shared/naming.ts` | `LIST_SUFFIX.requests`. |
| `shared/requests.ts` *(new, pure, tested)* | The request model, its state machine, who may decide what, and what a failed action records. |
| `webparts/requests/` *(new)* | The approver's queue, and the raise-a-request form. |
| `MySubmissions.tsx` | "Request deletion" / "Request share" on a row. |
| `shared/auditLog.ts` | `DeletionRequested`, `ShareRequested`, `RequestApproved`, `RequestRejected`, `ShareRevoked`. `EventType` is Text, so this is additive and safe. |
| Reconciliation | Grants the new permissions. No schema change. |

---

## 7. Out of scope, deliberately

- **Intercepting the native Share button.** Established as impossible — §5.1.
- **Approving a share made through SharePoint's own dialog.** That request goes to site owners and is
  outside this system entirely.
- **Bulk approval.** Each decision is about one document leaving, or one person gaining access.

---

## 8. Open questions for the client

1. **Does SDG's tenant allow external sharing at all?** If not, §5.3's external half is inert and should
   be switched off rather than built against.
2. **Should a share expire by default**, and after how long? A share with no expiry is permanent, and
   nobody reviews it.
3. **C-Level delete is segment-wide, or global for `GLOBAL`.** Confirmed as intended?

---

## 9. Test plan

1. PIC deletes a pending file in the approval library. ⚠ Since 2026-08-19 they CAN see a peer's and
   can delete it — the old expectation ("cannot see a peer's") tested draft security, which is now off.
2. HoU uploads; their own document appears in their own queue and they approve it.
3. PIC has no Delete in `Documents` — the native command is absent, not merely refused.
4. PIC requests deletion → HoU approves → the file is in the **recycle bin**, not purged.
5. Request deletion for a file already deleted → `Failed`, naming why.
6. HoD and C-Level delete directly, with no request.
7. PIC requests an internal share → HoU approves → the recipient opens that file and nothing else.
8. Same, external, config row off → refused, naming the setting.
9. Same, external, tenant sharing off → fails saying which layer refused.
10. Revoke a share → the recipient loses access and the scope is gone.
11. Every path writes its audit row, the failures included.
