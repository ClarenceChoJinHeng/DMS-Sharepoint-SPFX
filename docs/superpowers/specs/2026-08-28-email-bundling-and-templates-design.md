# Email bundling and the client's 11 templates — design

**Date:** 2026-08-28 · **Status:** DESIGN ONLY, nothing built · **Applies to:** both sites

Client, the night before the SDG migration:

> "the file email, client notice that if they uploaded 5 files it will send them 5 emails but if its
> 20 files then it will be 20 emails, we got to ensure that if uploader uploaders a file it will
> notify the approver in a proper bundle. same goes to when they approve the files. if both of my
> files are approve at the same time, it will send one email containing two files inside."

And two new flows requested by name:

> "a new flow for reminder approver to approver files after 3 days and also another flow to notify
> approver to Approve the Request in Share and Delete"

Plus an 11-template email format, reproduced verbatim in §5 because it is the client's requirement
and must not be paraphrased away.

---

## 1. Why 20 files send 20 emails

`NotifyApprovers` triggers on **item created or modified** in `Approval Document` with **`Split on` =
On** over `@triggerOutputs()?['body/value']`. Split-on spawns **one run per item**, so each file gets
its own run and its own email. It is working exactly as built; the requirement changed.

⚠ **The obvious fix is the wrong one.** Turning `Split on` OFF gives one run per trigger *poll*, and
a poll batch is "whatever arrived since the last check" — not "whatever the uploader submitted". Two
unrelated uploads a minute apart would bundle together, and one submission spanning a poll boundary
would still split. It would look fixed and be wrong in both directions.

## 2. The grouping key already exists

**`SubmissionId`** — one per press of Upload — and **`BatchId`** — one per destination within it —
have been stamped on every uploaded file since 1.0.221.0, in **all four libraries** (they had to be,
or Auto-route's copy would strip them). That is precisely the boundary the client is describing, and
it is already in the data.

⚠ **`SubmissionId` alone is the wrong grain, and getting this wrong is a disclosure.** A submission
can span several destinations, and different destinations can have different approvers. The correct
grouping is **(SubmissionId, approver group)**: one email per approver, listing only the files in
units they actually approve. Bundling on SubmissionId alone would tell an approver the filenames of
another unit's documents — the exact thing the per-unit ACL model exists to prevent.

**Files with no `SubmissionId`** — anything uploaded before 1.0.221.0, and every Bulk Upload file —
fall back to today's behaviour, one email each. Blank is *unknown*, never a value: three
referenceless files are three submissions, never one bundle of three. Same rule My Submissions
already follows.

## 3. The two hard parts

### 3.1 Waiting for the rest of the submission

A run fires on the first file and cannot know how many more are coming.

1. Trigger on created-or-modified, `Split on` **On** (unchanged).
2. Read `SubmissionId`. Blank ⇒ send a single-file email exactly as today.
3. **Delay** — 2–5 minutes — to let the remaining files land.
4. Re-query the approval library for every Pending item carrying that `SubmissionId`.
5. **Dedupe**: only ONE run per (SubmissionId, approver group) may send.
6. Build one email listing every file.

### 3.2 The dedupe, and the race that breaks the naive version

⚠ **With 20 files, 20 runs start within seconds of each other.** All 20 reach the "have we already
notified?" check before any has written a marker, all 20 see nothing, and all 20 send. The
check-then-write is not atomic.

Both mechanisms are needed:

- **`Concurrency control` → `Limit` = 1** on the trigger, so runs execute sequentially. ⚠ Power
  Automate warns this **cannot be changed after enabling** — confirm the intent before setting it.
  It also makes a 20-file submission take 20 × the per-run time, which the Delay already dominates.
- **A marker row, checked first.** Reuse the `Already logged` shape the audit flows already run on:
  filter `CRS Audit Log` for a row with this SubmissionId inside a window, write one after sending. A
  dedicated `EventType` (`ApproverNotified`) keeps it out of the file-lifecycle history. Proven
  pattern here, and no new list.

**A failed dedupe check must SEND anyway** — the same fail-open rule as the audit flows. A duplicate
email is something a person reads past; a missing approval notification stalls a document
indefinitely and nobody knows to look for it.

### 3.3 ⚠ The approval side has NO batch reference, and that is a real gap

The client wants the same bundling when files are approved together. There is nothing to group on:

- **Bulk Approve writes moderation status per file, sequentially, with no shared id.** There is no
  `ApprovalBatchId` anywhere in the system.
- Auto-route's approval email fires per routed file, so N approvals send N emails.

Three options, in order of preference:

1. **Group on `SubmissionId` again.** An approver clearing 3 of 5 files from one submission produces
   one email naming 3 files. Natural, needs no new column, and the recipient is the uploader — who
   thinks in submissions, because that is what My Submissions shows them. **Recommended.**
2. **Stamp an `ApprovalBatchId`** at approval time — a code change to `BulkApprovePanel.tsx` and
   `ApprovalDocument.tsx` plus a column in all four libraries. Exact, and the most work.
3. **Time-window grouping** — every approval for one uploader within N minutes. Closest to the
   client's literal words ("approved at the same time") and the least defensible: two unrelated
   approvals minutes apart would be presented as one action.

⚠ Option 1 has one limitation to state plainly to the client: files approved from **different**
submissions at the same moment still produce separate emails. Arguably correct — they are different
submissions — but it is not what "approved at the same time" literally asks for.

## 4. Consolidate the request emails into ONE flow

Templates 6–11 are six emails about `CRS Requests`. **Do not build six flows.** One flow on
created-or-modified over `CRS Requests`, branching on `RequestType` × `Status`, covers all six —
exactly the shape `CRS — Audit request activity` already uses with its `EventKind` Compose:

```
RequestType = Share    x Status = Pending   -> template 6  (to approver)
RequestType = Share    x Status = Approved  -> template 7  (to PIC)
RequestType = Share    x Status = Rejected  -> template 8  (to PIC)
RequestType = Deletion x Status = Pending   -> template 9  (to approver)
RequestType = Deletion x Status = Approved  -> template 10 (to PIC)
RequestType = Deletion x Status = Rejected  -> template 11 (to PIC)
```

Six flows would be six copies of the same trigger, the same approver lookup and the same recipient
resolution — and the copy that drifts is always the rarely-exercised one. Same reasoning that keeps
`GroupMembersEditor` as one component with two mount points.

⚠ **The approver lookup must filter `Role eq 'APR'` OR `'APRHC'`** according to the document's
library, exactly as `NotifyApprovers` / `HCNotifyApprovers` do. A request on an HC document notified
through an `APR` filter would email every ordinary Head of Unit about a Highly Confidential file.

## 5. The client's templates, verbatim

Reproduced as supplied. Bracketed tokens are theirs.

### 1 — Invitation Notification (to Recipient)
```
Subject: Welcome to [Site/Page Name/Folder Name]

Dear [Recipient Name],

You can now log in and access the Guthrie Central Respository System using the link below:
[Go to [SharePoint Site/Page Name Folder Name]]([Link to Site])

Note: If you experience any login issues, please try clearing your browser cache or opening the
link in an Incognito window

Thank you,
SD Guthrie
```

### 2 — Document Submission Notification (to Approver)
```
Subject: Approval Requested: [File Name]

Dear [Approver Name],

A new document has been submitted for your review and approval in SharePoint.
Document Name: [File Name]
Unit: [Unit Name]
Year: [Year]
Document Type: [Document Type]
Submitted By: [Author Name]
Submission Date: [Date]

Please review the file here: [View Document in SharePoint]([Link to Item])

Please choose an action:
[Approve]([Approval Link / Action])
[Reject]([Rejection Link / Action])

Thank you,
[PIC Name]
```

### 3 — Approval Reminder (to Approver) — NEW FLOW, after 3 days
```
Subject: REMINDER: Approval Required for [File Name]

Dear [Approver Name],

This is a friendly reminder that a document is still awaiting your review and approval in SharePoint.
Document Name: [File Name]
Unit: [Unit Name]
Year: [Year]
Document Type: [Document Type]
Submitted By: [Author Name]
Submission Date: [Date]

Please review the file here: [View Document in SharePoint]([Link to Item])

Please choose an action:
[Approve]([Approval Link / Action])
[Reject]([Rejection Link / Action])

Thank you,
[PIC Name]
```

### 4 — Approved Notification (to PIC)
```
Subject: Approved: [File Name] has been approved

Dear [PIC Name],

Great news! Your document has been approved.
Document Name: [File Name]
File Location: [File Location Path]
Approved By: [Approver Name]
Approval Date: [Approval Date]
Comment: [Approver Comments]
File Link: [File Link]

Thank you,
[Approver Name]
```

### 5 — Rejected Notification (to PIC)
```
Subject: Rejected: [File Name] has been rejected

Dear [PIC Name],

Your submission has been reviewed and was not approved.
Document Name: [File Name]
Rejected By: [Approver Name]
Rejection Date: [Rejection Date]
Reason/Comment: [Approver Comments]
File Link: [Staging File Link]
Please review the comment provided above, update the file in SharePoint as needed, and submit a new
approval request when ready.

Thank you,
[Approver Name]
```

### 6 — File Sharing Request (to Approver) — NEW
```
Subject: File Sharing Request: [File Name]

Dear [Approver Name],

I would like to share [File Name] to [Recipient Name/email].
Document Name: [File Name]
Unit: [Unit Name]
Year: [Year]
Document Type: [Document Type]
Requested By: [PIC Name]
Requested Date: [Date]
File Link: [File Link]

Please choose an action:
[Approve]([Approval Link / Action])
[Reject]([Rejection Link / Action])

Thank you,
[PIC Name]
```

### 7 — Approved File Sharing Request (to PIC) — NEW
```
Subject: Approved File Sharing Request: [File Name]

Dear [PIC Name],

Your request has been approved and the file link has been shared with [Recipient Email].

Thank you,
[Approver Name]
```

### 8 — Rejected File Sharing Request (to PIC) — NEW
```
Subject: Rejected File Sharing Request: [File Name]

Dear [PIC Name],

Your request to share [File Name] to [Recipient Email] has been rejected.
Document Name: [File Name]
Status: Denied
Reason/Comments: [Approver Reason]

Thank you,
[PIC Name]
```

### 9 — File Delete Request (to Approver) — NEW
```
Subject: Approval Needed to Delete File: [File Name]

Dear [Approver Name],

A request has been submitted to permanently delete a file from SharePoint.
Please review the details below to approve or deny this action.
Document Name: [File Name]
Location: [Folder Path Name]
Requested By: [PIC Name]
Requested Date: [Date]
File Link: [File Link]

Please choose an action:
[Approve]([Approval Link / Action])
[Reject]([Rejection Link / Action])

Thank you,
[PIC Name]
```

### 10 — Approved File Delete Request (to PIC) — NEW
```
Subject: Approved File Deletion Request: [File Name]

Dear [PIC Name],

Your request to delete the following file has been approved. The file has been moved to the
Recycle Bin.

File Name: [File Name]
Approved By: [Approval Name]
Deletion Date: [Current Date]
Note: This file will be retrievable from the SharePoint Recycle Bin for the next [90] days if you
need to restore it.

Thank you,
[Approver Name]
```

### 11 — Rejected File Delete Request (to PIC) — NEW
```
Subject: Rejected File Deletion Request: [File Name]

Dear [PIC Name],

Your request to delete [File Name] has been rejected.
The file will remain in its current location.
Document Name: [File Name]
Status: Denied
Reason/Comments: [Approver Reason]

Thank you,
[Approver Name]
```

## 6. ⚠ Things in the templates that cannot be built as written

Raise all of these with the client BEFORE building, not after.

### 6.1 `[Approve]` / `[Reject]` cannot be one-click links
Templates 2, 3, 6 and 9 imply acting from inside the email. A link that actually approves needs
either Power Automate Approvals or an authenticated trigger URL — and **a bearer trigger URL inside
the SPFx bundle was already rejected** in the HC design, because the token ships to the browser.

**What is buildable:** both links point at the relevant page — `ApprovalDocument.aspx?itemId=N` (plus
`&lib=hc` for HC, since item ids are per-list) or `Request.aspx` — where the approver acts in their
own session with their own permissions. That is also what makes the audit row name the real actor.

**⚠ SUPERSEDED 2026-08-31 (1.0.335.0) — BOTH LINKS ARE BUILT, AND THEY DO NOT COLLAPSE.**
The paragraph that stood here said the two buttons must become one honest "Review and decide" link,
because an *Approve* button that merely navigates is worse than none. The client asked to follow
Crystal's templates exactly, and there is a third option the paragraph missed: keep both buttons and
have each one **pre-select its radio** on the approval page.

```
.../ApprovalDocument.aspx?itemId=N&decision=approve
.../ApprovalDocument.aspx?itemId=N&decision=reject     (+ &lib=hc for HC)
```

`decisionFromLink` in `shared/approvalQueue.ts` (pure, 4 tests) parses it;
`ApprovalDocument.tsx` applies it at load.

- **PRE-SELECTING IS NOT DECIDING.** `submitDecision` is still reached only by pressing the button,
  so the destination-folder guard, the name-clash check and the `ApproveItems` probe all still run.
  The original objection — *"the approver believes they have approved"* — is answered by the page
  itself: they arrive on it, with their choice ticked and a button that plainly has not been pressed.
- **⚠ ONLY WHILE THE DOCUMENT IS STILL PENDING.** An already-decided one shows what actually happened
  to it. Pre-ticking *Reject* on a document somebody approved an hour ago states something false on
  the one screen that IS the record of the decision.
- **⚠ THE PARAMETER IS READ ONCE AND THEN STRIPPED FROM THE ADDRESS BAR.** The queue rewrites
  `?itemId=` with `replaceState` as the approver moves through it, so a parameter that survived would
  pre-tick a decision from an email about a **different document** — the one failure here that could
  genuinely mislead. Held in a `useRef`, so a re-render cannot re-read it either.
- **Anything unrecognised is `undefined`, never a guess** — a mangled parameter leaves the page
  exactly as it behaves with no parameter: nothing ticked, button disabled. `pending` is deliberately
  unreachable.
- **The bearer trigger URL stays rejected**, and is worse here than in the HC design: the reminder
  goes to up to ten mailboxes, so ten copies of a token that approves a document would be sitting in
  ten inboxes.

### 6.2 Template 1 is not a flow
There is **no SharePoint trigger for group membership changes.** And since 2026-08-27 the client's own
decision is that **outsiders cannot be added at all**, so this can only ever reach an account that
already exists in the tenant.

Options: send it from code in `GroupMembersEditor`'s add path, which already toasts on success; or a
scheduled flow diffing membership, which is expensive and lags. **Code is the right answer**, making
this the one template that is not Power Automate work.

### 6.3 "[90] days" is wrong — it is **93**
Template 10 promises 90. The recycle-bin retention here is **93 days**, as recorded throughout this
project. A client-facing email must not understate a retention window somebody may rely on. Fix to
93, or make it a config row so the client can set it.

### 6.4 Template 5's subject is malformed as supplied
`Rejected: [​ Name] has been rejected` — the token is missing "File" and contains a zero-width
character. Assume `[File Name]`.

### 6.5 "Respository" is misspelled in template 1
Client's own text. Ask before correcting; almost certainly a typo rather than branding.

### 6.6 ⚠ Bundling and per-file templates directly conflict
Every template is written for ONE file (`Subject: Approval Requested: [File Name]`). Bundling five
files into one email means the subject and body must change shape — e.g. `Approval Requested: 5
documents` with a table beneath. **The client has asked for both bundling and these templates, and
the two cannot both be satisfied literally.**

Proposal to put to them: keep the single-file template exactly as written when a submission has one
file, and use a bundled variant when it has more. Get that agreed before building, because it decides
the shape of every affected email.

## 7. Build order

1. **The two flows the client asked for by name** — the 3-day reminder (§5 template 3) and the request
   notifications (§4, which consolidates six templates into one flow). New capability, no rework.
2. **Reformat the three existing emails** to templates 2, 4 and 5. No structural change, just fields:
   `Unit`, `Year`, `Document Type`, `Author` and the submission date are all already on the item.
3. **Bundling last**, because it is the only part needing a Delay, a one-way concurrency change and a
   dedupe marker — and because the reformatted single-file emails are the prerequisite for designing
   the bundled variant (§6.6).
4. **Template 1 separately, in code.**

⚠ **Every flow built or cloned for this must be created while signed in as the SERVICE ACCOUNT**
(`crs@sdguthrie.com` on SDG). The connection is baked in at the first action, for life — a flow built
as a person stops silently when that password changes.

⚠ **Each HC clone is four edits, not three:** the trigger's List Name, `GetDoc`'s List Name, the Group
Map filter (`APRHC`), and `&lib=hc` on the link. Every HC clone in this project has shipped a fault
from a library reference nobody swapped.
