# Share Guard — Design

**Status:** proposed (approved in brainstorming 2026-07-20). An approval gate for
sharing DMS documents: a member cannot share directly — they submit a **request**,
and only an **admin/owner** can approve it, which then grants the recipient access.

## The hard constraint (read first)

**A web part/extension cannot intercept SharePoint's native Share button or hold a
sharing link for approval.** Native sharing grants access the instant a link is
created; there is no supported "pending" hook. So Share Guard does **not** intercept
native sharing — it **replaces** it:

1. **Enforcement (admin config, not code):** native sharing is locked down so members
   can't share. Site Permissions → *Sharing settings* → **"Only site owners can share
   files, folders, and the site."** This is the single most important step — without
   it the guard is decorative and trivially bypassed via the real Share button.
2. **Compliant path (this solution):** a custom guarded "Request to Share" button that
   feels native, backed by a request list and an admin approval queue.

## Decisions (from brainstorming)

- **Approach A** — SPFx **ListView Command Set** (a button in the library
  toolbar/right-click menu), not a standalone page. Feels native.
- **Recipients: internal only.** On approval the system grants an internal user
  access to the item. No external/guest machinery.
- **Approver: central admin/owner.** One approval queue for the whole DMS (not
  per-department).
- **Access level: requester chooses Read or Edit;** the admin sees and approves it.
- **Share source: both Staging and Documents.** The button appears on both libraries.
- **A member cannot share** — only request. The admin executes the grant.

## Components

1. **Native sharing lockdown** — admin sets "Only site owners can share" (config).
2. **Guarded Share button — SPFx ListView Command Set.** Registered on the Staging
   and Documents libraries. Enabled when exactly one file/folder is selected. Opens a
   dialog: **Recipient** (internal people picker), **Access level** (Read / Edit),
   **Reason**. On submit → writes one `DMS Share Requests` row (Status = Pending). No
   access is granted at this point.
3. **`DMS Share Requests` list** — request store + audit trail (schema below).
4. **Approval Queue web part (admin-only)** — lists Pending requests; **Approve** /
   **Reject**. On Approve: grant the recipient the chosen role on the exact item,
   set Status = Approved, trigger the recipient notification. On Reject: set
   Status = Rejected; nothing is granted.

## End-to-end flow

```
Member selects a file/folder in Staging or Documents → "Request to Share"
  → picks Recipient (internal) + Read/Edit + Reason → Submit
     → DMS Share Requests row, Status = Pending           (nothing granted)
        → Admin opens the Approval Queue web part
           → Approve → addroleassignment(item, recipient, Read|Edit)
                       → Status = Approved → notify recipient (link)
           → Reject  → Status = Rejected                  (nothing granted)
```

## Data model — `DMS Share Requests`

| Field (internal name TBD at creation, read back via /fields) | Type | Purpose |
|---|---|---|
| Title | Text | short label (e.g. "GCO/Invoice.pdf → J. Tan") |
| ItemUrl | Text | server-relative path of the file/folder (display + fallback) |
| ItemUniqueId | Text | the item's UniqueId — rename-proof handle used for the grant |
| ItemType | Choice | File / Folder |
| Library | Choice | Staging / Documents |
| Requester | Person | who asked (set from context) |
| Recipient | Person | internal user to grant access to |
| AccessLevel | Choice | Read / Edit |
| Reason | Note | justification |
| Status | Choice | Pending / Approved / Rejected (default Pending) |
| DecidedBy | Person | admin who actioned it |
| DecidedOn | DateTime | when actioned (M/D/YYYY per site locale) |

**List permissions:** members get **Contribute** but with item-level security =
"read items created by the user" + "edit items created by the user", so a member sees
and edits only **their own** requests. Admins/owners see all. Status/DecidedBy/DecidedOn
are only writable by the approval web part (admin context).

## How the grant works (internal → clean)

On Approve the web part resolves the item by `ItemUniqueId`
(`GetFileById`/`GetFolderById`), ensures the item has unique permissions if needed,
and calls `addroleassignment(item, recipientPrincipalId, roleDefId)` where roleDefId
= **Read** or **Edit**. SharePoint grants the recipient access to that item **plus
Limited Access up the parent chain**, so a direct link opens even though the recipient
is not in the unit's tiered groups. The grant is **item-scoped** — the recipient sees
only that one shared item, nothing else in the unit.

## Notifications

Recipient email on approval (with the item link) is handled by a small **Power
Automate flow** triggered on `DMS Share Requests` (Status → Approved), sender = the
existing **DMS No Reply** shared mailbox — matching the auto-route pattern. (Doing it
in the web part via Graph sendMail is a fallback.) Optional: notify admins on a new
Pending request.

## Security summary

- **Enforcement = the lockdown.** Owners-only sharing removes the native bypass; the
  guarded button is the only path left for members.
- **A member never grants access** — they only create a Pending row. The grant is an
  admin action.
- **Item-scoped grants** — recipients get exactly one item, never the unit.
- **Full audit trail** — every request records requester, recipient, level, reason,
  decision, and who/when.

## Out of scope (YAGNI for v1)

- External/guest sharing (internal-only chosen).
- Per-department approver routing (central admin only).
- Expiry / time-limited access — add `ExpiryDate` + a cleanup flow later if needed.
- Revocation UI — v1 revokes by removing the item role assignment manually; a
  "Revoke" button on the queue can come later.

## Risks / watch-items

- **Lockdown is manual and essential.** If an admin forgets "Only site owners can
  share," members can still share natively and the guard is bypassed. Document it as
  step 1 of deployment and verify it.
- **Command-set dialog complexity.** SPFx custom dialogs (people picker + fields) are
  more involved than a plain web part. Fallback = Approach B (standalone request web
  part) if the dialog fights us.
- **Recipient discoverability.** After a grant, the recipient needs the direct link
  (they won't see the item by browsing). The approval notification must include it.
- **Members need Contribute on the requests list** — a targeted grant (via the global
  site-access group) on just `DMS Share Requests`, with item-level security so they
  see only their own.

## Sequencing

1. Create `DMS Share Requests` list (+ item-level security) and set the native
   sharing lockdown.
2. Build the Approval Queue web part (grant-on-approve) — testable on its own with a
   hand-created Pending row.
3. Build the guarded Share button (Command Set + dialog) that writes Pending rows.
4. Add the notification flow.
5. Verify end-to-end: member requests → admin approves → recipient opens the link;
   confirm a member cannot share natively.
