# Asking to delete or share an approved document

Plain-language notes for SD Guthrie. Written 2026-08-15, when the feature was built.

---

## What was asked for

- A **PIC** can delete their own files in the approval library, but **not** in the main Documents
  library. There, they **ask**, and the **Head of Unit** approves.
- A PIC can also **ask** to share a document with someone. Again, the Head of Unit approves.
- **Heads of Department and C-Level** delete and share directly. They do not ask anyone.

That is what has been built.

---

## How it works, in one paragraph

The uploader opens one of their approved documents on **My Submissions** and presses *Request
deletion* or *Request share*, giving a reason. It appears on the **Requests** page for the Head of
Unit of the unit that document belongs to. If they approve, the action happens straight away. A
deletion sends the file to the **recycle bin**, where it can be restored for **93 days** — nothing is
destroyed. A rejection leaves the document untouched, and the uploader sees the reason.

---

## Six things worth knowing before it goes live

### 1. A Head of Unit approves their own uploads

Heads of Unit can now upload, and they are also the approver for their unit. So their own documents
are approved by them. This was agreed, and it is normal for a small team — but it means **for those
particular documents the approval step is a record rather than a check.** If a genuine second pair of
eyes is needed for a Head of Unit's own documents, that has to be a different person, and it is worth
deciding now rather than after go-live.

### 2. A Head of Unit becomes the sharing *authority* for their unit, not only an approver

To approve a share, a person must be able to perform it. So Heads of Unit are being given sharing
rights over their unit's folder.

The consequence: **they can share directly, without going near the request screen** — with anyone the
tenant permits, at view or edit level, for anything in their unit. The request page is the sanctioned,
recorded route. It is not a gate that stops them.

This is unavoidable rather than an oversight, and it is the same for Heads of Department and C-Level.

### 3. SharePoint's own Share button cannot be intercepted

If someone presses **Share** in SharePoint itself rather than using our page, our system never sees
it. This was investigated fully in July and settled: a web part cannot take over that button.

The protection that does work is SharePoint's own setting, *"Only site owners can share files,
folders, and the site."* With it on, a member pressing Share does **not** grant anything — it becomes
a request to a site owner. **We recommend leaving that setting on.** Our page then becomes the route
that reaches the Head of Unit; the native button reaches an administrator. Neither one leaks.

**One more setting closes the gap you asked about (2026-08-19).** You said approvals should reach the
Head of Unit rather than a site owner or an email. By default SharePoint turns that blocked Share
press into an **access request emailed to the site owner** — which is exactly the route you want gone.

> Site Settings → Site Permissions → **Access Requests** → turn OFF

With it off, a person pressing Share is simply told they cannot share this file. No email, no owner
involved, nothing sitting in an administrator's inbox. The Requests page is then the only route that
*works* — which is a stronger position than blocking the button, because there is nothing left to go
around.

Worth being clear about what this does and does not do:

- It does **not** stop someone pressing the button; it makes pressing it do nothing.
- Nobody loses access they already have. This only affects requests for *new* access.
- If a person genuinely needs something and there is no request route, they will ask a colleague
  instead of the system — so the Requests page needs to be somewhere they can find it.

### 4. External sharing is assumed to be allowed, and can be switched off in one click

You asked for external recipients to be possible for now. Two separate things control it:

- **Your tenant and site settings.** We do not control these. If your IT has external sharing off, an
  approved external share will simply fail, and the screen will say so.
- **A setting in the CRS Config list** (`allowExternalSharing`). This one is ours. **It is off unless
  explicitly switched on** — and if the setting cannot be read for any reason, it stays off. That is
  deliberate: the safe direction for "we are not sure" is *internal only*.

So nothing goes outside the company by accident. But **please confirm whether your tenant allows
external sharing at all** — if it does not, we should turn our half off rather than leave a button
that always fails.

### 5. Every share is permanent unless someone ends it

A share stays in place until it is removed. Requests can carry an **expiry date**, and we recommend
using it — but it is optional today, and a share with no expiry is forever.

**One thing is not built yet: a way to revoke a share from our screens.** It is designed and can be
added; today a share is removed through SharePoint's own permissions on the file. If sharing is going
to be used regularly, this is worth building before go-live rather than after.

### 6. Deleting is a request; *seeing* is not

This changes who can **delete** a document. It changes nothing about who can **read** one. Everyone in
a unit still sees every approved document in that unit — agreed separately, and unchanged here.

---

## Two things must be done before any of this works

Both are one-time setup on the live site. Until they are done the pages load and do nothing useful:

1. **A `CRS Share` permission level must be created**, containing *Manage Permissions*. Without it, a
   Head of Unit cannot carry out a share they have approved, and the approval fails. ⚠ **Check the
   contents, not the name** — a level called `CRS Share` that does not actually contain *Manage
   Permissions* fails at the moment of approval and looks like a bug in the page.
2. **Folder reconciliation must be re-run after the update is deployed.** That is what actually gives
   Heads of Unit their new delete and share rights on the folders. Nothing takes effect until it runs.
3. **Access Requests should be turned off** (see point 3 above), or the native Share button keeps
   sending approvals to a site owner by email — the thing this feature exists to replace.

On the rehearsal site, steps 1 and 2 were completed by the reconciliation run of **2026-08-19**: the
log shows `→ CRS Delete` and `→ CRS Share` granted to every Head of Unit group, which also proves the
permission level exists.

The **Requests** list itself is created automatically the first time an administrator opens the
Requests page.

---

## Questions we need answered

1. **Does your tenant allow external sharing?** If not, we switch our half off.
2. **Should shares expire by default**, and after how long?
3. **Is a Head of Unit approving their own uploads acceptable**, or should those go to someone else?
4. **Should share revocation be built before go-live?** (See point 5.)
