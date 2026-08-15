# Deletion Request + Approval — Design

**Date:** 2026-08-04
**Status:** SUPERSEDED 2026-08-15 by `2026-08-15-deletion-and-share-requests-design.md`.

Its premise is reversed: this spec removes `Delete Items` from a PIC on **Staging**, and the client
confirmed on 2026-08-15 that a PIC *should* delete there. The request workflow moved to **Documents**,
where a PIC genuinely has no delete — and got simpler, because the approver can now perform the
deletion in their own session instead of handing it to a flow with its own credentials.

Kept for the §2 permission-level detail and the reasoning in §1, both of which still apply.

**Was:** DRAFT — three decisions open (§9). Depends on the `DMS Upload` level (§2).
**Branch:** `feat/folder-abbreviations`
**Parent:** `2026-08-03-visibility-scope-and-view-group-separation-design.md` §6, which records
that per-author delete cannot exist. This spec is the answer to that limitation rather than a
workaround for it.
**Related:** `2026-08-03-role-personas-and-department-fanout-design.md` (the role model)

---

## 1. What this solves

The client's rule is *"If any PIC wants to delete they need to get an approval from the head of
unit first."* SharePoint has no such gate: a user either holds `Delete Items` on a folder or
does not, and the permission is folder-scoped, never author-scoped.

So the rule is implemented as an **application-level workflow on top of a removed permission**,
not as a permission:

1. PIC loses `Delete Items` on Staging (§2). The rule now has teeth — there is no Delete button
   for them to press.
2. PIC raises a **request** from the library command bar (§3).
3. An approver for that unit **authorises** it on a custom page (§5).
4. An automated flow **performs** the deletion with its own credentials and logs it (§6).

The user-visible effect is "a PIC can delete, with approval," which is what the client
pictures. The permission never leaves the flow.

## 2. Prerequisite: the `DMS Upload` level

Nothing here means anything until PIC cannot delete. That is a third custom permission level,
created once per site alongside `DMS Approve` and `DMS Delete`:

| Level | Copy from | Then |
| --- | --- | --- |
| **DMS Upload** | Contribute | untick **Delete Items** |

`UPL` then maps to `DMS Upload` instead of stock Contribute — a one-line change to `ROLE_LEVEL`
in `FolderManager.tsx`.

**This reverses the 2026-08-03 decision that uploaders keep delete**, and it applies to every
holder of `UPL`, not only PICs — so Head of Department 1/4 and Head of Unit 1/4 lose Staging
delete too. That is the intended consequence, not a side effect: with the flow performing
deletions, nobody needs the raw permission. If the client wants heads to retain direct delete,
`UPL` must split into two roles (§9 D2).

> Same failure mode as the other two levels: until `DMS Upload` exists on the site by that exact
> name, reconciliation logs `no "DMS Upload" role definition on site` and skips the assignment.
> Existing uploaders keep Contribute, so nobody breaks — the restriction simply does not take
> effect, and every PIC keeps a working Delete button. Create the level before the next run.

## 3. The request button — a ListView Command Set

A `ListViewCommandSet` extension puts **"Request deletion"** in the Staging command bar and `...`
menu, beside where Delete sits today.

### 3.1 There is a precedent in this repo, and it must not be copied verbatim

`src/extensions/uploadCommand/UploadCommandCommandSet.ts` is an existing command set. Two things
about it are load-bearing for this design:

- **It is not deployed.** Its component id is absent from `componentIds` in
  `package-solution.json`; the only registered extension is `HideAppBarApplicationCustomizer`. So
  the *shape* is proven, the *deployment path* is not, and §4 exists because of that.
- **It hardcodes `list?.title === "Staging"`** and an absolute path to the retired sandbox site.
  The first is the CRS-rename trap (memory `dms-to-crs-rename-pending`): the client renames every
  DMS-named artefact at import, the title check silently stops matching, and the button quietly
  never appears. The second concatenates a full site path onto the current web URL and yields a
  404.

Both are fixed by the same rule, stated once here: **no library name, list title, or site path is
ever hardcoded.** The Staging library is resolved from `DMS Config` at runtime, the same way every
other component resolves it.

### 3.2 Visibility is driven by permissions, never by persona

The command is visible when the current user **lacks** `DeleteListItems` on the current folder.

Not "when the user is a PIC." A group-name or Group-Map-role check goes stale the moment someone's
membership changes, and a request button offered to someone who can already delete is a confusing
no-op. Reading the effective permission is the only test that stays true, and it degrades
correctly: a PIC promoted to Head of Unit stops seeing the request button and starts seeing the
real Delete, with no configuration change.

Because unit folders have unique permissions, the check is **per folder**, not per library.
`onListViewUpdated` is called frequently and must stay cheap, so: resolve the current folder, look
it up in a `Map<folderPath, boolean>` cache, and on a miss fetch `EffectiveBasePermissions` for
that folder asynchronously, store it, and call `this.raiseOnChange()` to re-evaluate. First render
of a new folder may briefly omit the button; that is preferable to a synchronous fetch on every
selection change.

### 3.3 Only Pending and Rejected items may be requested

An **Approved** file has already been copied to Documents by Auto-route. Deleting the Staging copy
removes the staging artefact only — the published document survives. A PIC who "retracted" a
document and did not is a worse outcome than no button at all.

| Approval state | `_ModerationStatus` | Request allowed |
| --- | --- | --- |
| Approved | 0 | **no** |
| Rejected | 1 | yes — the main housekeeping case |
| Pending | 2 | yes — cancels an upload before anyone acts on it |
| Draft | 3 | yes |

Read `_ModerationStatus` off the selected row, and note it is an **integer, not a label** (memory
`sp-column-formatting-gotchas`). On an Approved selection the command is hidden rather than
disabled-with-a-tooltip: a greyed button invites a support ticket asking how to enable it.

Removing an approved *document* is a `DEL`-on-Documents action performed by a head, and is
deliberately not reachable from this workflow.

### 3.4 The dialog

A `BaseDialog` collecting a required **reason**, showing the file name and the unit it sits in, and
naming who will be asked to approve. On confirm it writes the three columns (§7) in one `MERGE`.
Writing needs `Edit Items` — which `DMS Upload` retains, and which every approver holds whether we
want it or not (`DMS Approve` cannot drop it). The one permission we could not remove is what makes
this workflow possible.

## 4. Registering the command against the library

A `ListViewCommandSet` needs a `CustomAction` bound to the list; the `ClientSideComponent` element
alone does not attach it. Two routes:

| Route | Verdict |
| --- | --- |
| Feature-framework `elements.xml` with `RegistrationId="101"` | **rejected** — 101 is *every* document library on the site, so the button would appear on Documents too, and `skipFeatureDeployment: true` would have to change |
| POST a `UserCustomAction` to the Staging list | **chosen** |

The Folder Manager admin web part already provisions lists, columns and folders over REST, so it
gains one more idempotent step: POST to `/_api/web/lists(guid'…')/UserCustomActions` with
`Location = ClientSideExtension.ListViewCommandSet.CommandBar`, the component id, and the
registration scoped to the Staging list resolved from `DMS Config`.

Three reasons this is the right route: it binds to **one** library rather than all of them; it
survives the CRS rename because the library is resolved at runtime; and it needs no PowerShell or
PnP, which is a standing constraint (memory `feedback-avoid-powershell-scripts`). It must report
"already registered" rather than creating a duplicate action on a second run.

## 5. The approver page

A new web part, sibling to `ApprovalDocument.tsx`, listing items where `DeleteStatus = Requested`
**for units the signed-in user approves for** — resolved from their `_APR` Group Map rows, exactly
as the Form resolves upload paths. Each row shows file name, unit, requester, reason, approval
state, and Approve / Reject.

**Approving writes `DeleteStatus = Approved`. It does not delete.** SPFx has no privilege
elevation, so a page that deleted directly would work only for approvers who hold delete — which,
after §2, is nobody. Writing a column needs `Edit Items`, which every approver has. This is what
makes the workflow independent of who holds what.

It also closes a gap that an email-approval design would have left open: nothing in the data model
designates *the* Head of Unit, so any approver for that unit can authorise. With a page that
self-trims from the Group Map, that is the correct behaviour rather than a compromise — no new
per-unit designated-approver field for the client to maintain.

Rejecting clears `DeleteStatus` back to empty and leaves the file untouched.

## 6. The flow

Automated, on Staging, in the same family as Auto-route:

1. Trigger: item modified, with a **trigger condition** on `DeleteStatus = Approved` so the flow
   does not re-enter on its own writes.
2. Write the audit row to `DMS Deletion Log` (§7) — **before** the delete. If the file goes first
   and the log write fails, the deletion is unevidenced, which defeats the point of the feature.
   Log-then-delete can at worst leave a log row for a file that still exists, which is visible and
   fixable; the reverse is neither.
3. Delete the file. It lands in the recycle bin (93 days), so the log is the durable record, not
   the recovery mechanism.

Failure surfaces as a flow run failure plus the item sitting at `Approved` with no log row —
deliberately not silent.

## 7. Data

Columns on Staging. **Named without spaces on purpose** — `Document_x0020_Type` and
`Vendor_x002f_CustomerName` have cost real debugging time, and a space in a display name is not
worth the encoding guesswork:

| Column | Type | Notes |
| --- | --- | --- |
| `DeleteStatus` | Choice | `Requested` / `Approved` / `Rejected`; empty = no request |
| `DeleteReason` | Text | required by the dialog, not by the column |
| `DeleteRequestedBy` | Text | the requester's email, captured at request time |

`DeleteRequestedBy` is stored rather than derived from `Modified By`, because approving is an edit
and would overwrite `Editor` — the same trap that forced Staging's "Uploader" column to be `Author`
(memory `dms-uploader-column-is-author`).

> ⚠ **SUPERSEDED 2026-08-13 — do NOT build this list.** The audit log
> (`2026-08-13-audit-log-design.md`) covers deletion as a `Deleted` event type, and two overlapping
> append-only trails would leave a permanent question about which one is authoritative. When this
> feature is built it writes a `Deleted` row to the **audit log**, carrying the request reason and the
> approver in `Details`. The rest of this section is kept only to record what the columns were for.

New list **`DMS Deletion Log`** — `Title` (file name), `Unit`, `RequestedBy`, `ApprovedBy`,
`Reason`, `DeletedOn`. `DeletedOn` is written `M/D/YYYY` per the site locale and displayed
`DD/MMM/YYYY` like every other date. The list is append-only; nobody but Owners gets more than
Read, or the audit trail is editable by the people it audits.

Both the columns and the list are renamed by the client at import (`DMS_` → `CRS_`), so the log
list name comes from `DMS Config`, not a constant.

## 8. Out of scope

- **Deleting approved documents from Documents.** That is the `DEL` role, already designed.
- **Bulk requests.** One file per request; a multi-select request with one shared reason makes the
  audit row meaningless.
- **Withdrawing a request.** The PIC re-requests or asks the approver to reject. Adding a withdraw
  path doubles the state machine for a rare case.

## 9. Decisions open

- **D1 — Does the requester get notified of the outcome?** A rejection is currently silent outside
  the page. An email needs a second flow branch; the page alone means the PIC has to look.
- **D2 — Do heads keep direct delete?** §2 removes it from every `UPL` holder. Keeping it for Head
  of Department/Unit 1 and 4 means splitting `UPL` into two roles and two group suffixes, with a
  wrong pick invisible until someone deletes something.
- **D3 — Is the retired `uploadCommand` extension fixed or deleted?** It is undeployed and carries
  a broken URL and the hardcoded title. Deleting it removes a live copy-paste hazard; fixing it
  means owning a second command set.

## 10. Implementation order

1. `DMS Upload` level + `ROLE_LEVEL` change. Small, and required by everything else.
2. Columns + `DMS Deletion Log`, provisioned by Folder Manager.
3. The command set + dialog, with the permission-based visibility and state rules.
4. `UserCustomAction` registration step in Folder Manager.
5. The approver web part.
6. The flow.

Steps 1–2 are shippable and useful on their own: after step 1 the client's rule is enforced, with
the escalation being a conversation rather than a button.

## 11. Verification

Needs the tenant, a PIC account and an approver account, and all three custom levels created.

1. **The button appears for a PIC and not for a head.** Same file, two accounts. This is the check
   that proves visibility is permission-driven; a persona-driven implementation passes a casual
   look and fails here only after a membership change.
2. **The native Delete is gone for the PIC.** Confirm SharePoint hides it rather than showing it
   and failing — two delete-ish buttons is the worst outcome.
3. **Approved items offer no request.** Select an approved file as a PIC; no command.
4. **Cross-unit isolation on the approver page.** An approver for unit A must not see unit B's
   requests. Check by direct URL, not by browsing.
5. **Approval deletes and logs, in that order.** Confirm the log row exists and matches, and that
   the file is in the recycle bin.
6. **A failed log write does not delete.** Temporarily point the flow at a bad list name and confirm
   the file survives.
7. **The flow does not loop.** Confirm one run per approval, not two.
8. **Registration is idempotent.** Run the Folder Manager step twice; the second reports already
   registered and does not create a second button.
9. **Renamed library.** Rename Staging and confirm the button still appears — the check the existing
   `uploadCommand` would fail.
