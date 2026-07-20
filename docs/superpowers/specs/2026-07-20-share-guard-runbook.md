# Share Guard — Admin Runbook

Manual SharePoint/tenant steps that back the Share Guard code. **Do these in order.**
Steps 1–2 unblock the code (field names); **Step 5 is the enforcement — without it the
guard is bypassable.** Plan: `docs/superpowers/plans/2026-07-20-share-guard.md`.

---

## Step 1 — Create the `DMS Share Requests` list

Site contents → New → **List** → name it exactly **DMS Share Requests**. Add columns:

| Column (display name) | Type | Notes |
|---|---|---|
| Title | Single line | default column — keep |
| ItemUrl | Single line of text | server-relative path of the shared file/folder |
| ItemUniqueId | Single line of text | the item's UniqueId (rename-proof handle) |
| ItemType | Choice | choices: **File**, **Folder** |
| Library | Choice | choices: **Staging**, **Documents** |
| Requester | Person | who asked |
| Recipient | Person | internal user to grant access to |
| AccessLevel | Choice | choices: **Read**, **Edit** |
| Reason | Multiple lines (plain text) | justification |
| Status | Choice | choices: **Pending**, **Approved**, **Rejected**; default **Pending** |
| DecidedBy | Person | admin who actioned (set by the queue) |
| DecidedOn | Date and Time | when actioned (site locale, M/D/YYYY) |

## Step 2 — Read back the FROZEN internal field names (the code needs these)

SharePoint freezes internal names unpredictably (spaces → `_x0020_`, etc.). Get the real
ones and record them below; the web parts must use these, not display names.

Browse to:
```
https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground/_api/web/lists/getbytitle('DMS%20Share%20Requests')/fields?$select=Title,InternalName&$filter=Hidden eq false
```

Record the internal names here after creating the list:

| Display | Internal name (fill in) |
|---|---|
| ItemUrl | `____` |
| ItemUniqueId | `____` |
| ItemType | `____` |
| Library | `____` |
| Recipient | `____` (Person → set via `<InternalName>Id`) |
| AccessLevel | `____` |
| Reason | `____` |
| Status | `____` |
| DecidedBy | `____` |
| DecidedOn | `____` |

> If any differ from the plan's default names (`ItemUrl`, `ItemUniqueId`, `ItemType`,
> `Library`, `RecipientId`, `AccessLevel`, `Reason`, `Status`), tell the developer so the
> web part payloads/queries are corrected before testing.

## Step 3 — Item-level security (members see only their own)

List Settings → **Advanced settings** → *Item-level Permissions*:
- **Read access:** "Read items that were created by the user"
- **Create and Edit access:** "Create items and edit items that were created by the user"

## Step 4 — Grant members Contribute on this list only

List Settings → **Permissions for this list** → *Stop Inheriting Permissions* → grant the
**global site-access security group** **Contribute** on this list. (Item-level security
from Step 3 keeps each member limited to their own rows.) Admins/owners keep Full Control.

## Step 5 — NATIVE SHARING LOCKDOWN (the enforcement — do not skip)

Site Settings → **Site permissions** → *Sharing settings* (or the Site permissions panel →
"Change how members can share") → set:

> **"Only site owners can share files, folders, and the site."**

**Verify:** sign in as a member (e.g. Kaisya) and confirm the **Share** / **Copy link**
action is unavailable or errors. If a member can still create a sharing link, the guard is
bypassable — fix this before going live.

## Step 6 — Record the Read/Edit role definition IDs (for verification)

```
https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground/_api/web/roledefinitions?$select=Id,Name
```
- "Read" Id: `____`
- "Edit" Id: `____`

(The Approval Queue web part looks these up by name at runtime; recording them is only for
sanity-checking that both role definitions exist on the site.)

---

## Later (Task 6) — Notification flow

Power Automate → automated cloud flow → trigger **When an item is created or modified** on
`DMS Share Requests`:
- Condition **Status == Approved** → **Send an email (V2)** as the **DMS No Reply** shared
  mailbox → to `Recipient Email`, body includes the absolute `ItemUrl` link.
- (Optional) On create + Status Pending → notify site admins that a request is waiting.

Test by approving a request in the queue and confirming the recipient email arrives with a
working link.
