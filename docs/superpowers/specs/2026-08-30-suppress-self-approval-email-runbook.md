# Suppressing the approval email when the approver uploaded the document

**Date:** 2026-08-30
**Status:** DESIGN — not built.
**Flows to change:** `Auto-route approved Pending folders to Documents Library`, **and its clone**
`HC Auto Route`.
**Client QA #5:** *"MY ROLE IS APPROVER. IF I UPLOAD DOCUMENT, I SHOULD NOT BE RECEIVING AN EMAIL
STATING MY UPLOADED FILES HAS BEEN APPROVED."*

---

## 0. Why this is a flow change and not a code change

Nothing in the SPFx package sends that email. Auto-route's last action —
`Send an email from a shared mailbox (V2)` — notifies the **uploader** that their document was
approved and routed. Correct for a PIC; pointless for a Head of Unit who approved their own upload,
because they are the person who did it.

Two routes end with an approver receiving it, and a fix has to cover both:

1. **They approved their own upload** — through the approval page, the bulk panel, or SharePoint's
   native Approve/Reject command.
2. **Self-approve fired** (1.0.234.0). A Head of Unit's own upload auto-approves where the
   `autoApproveOwnUpload` config row is on, and **the client's explicit instruction at the time was
   that the email should still be sent**. ⚠ **THIS CHANGE REVERSES THAT** — same email, same person,
   opposite answer. Confirm it is wanted in both cases before building; the 2026-08-24 note records
   the previous decision in as many words.

---

## 1. The signal: `Editor`, and its limits

SharePoint's content approval records **no "approved by" field**. What it does is stamp `Editor`
(*Modified By*) with whoever changed the moderation status — the same fact memory
`dms-uploader-column-is-author` records from the other direction, and the reason the `Uploader`
column must be bound to `Author` and never `Editor`.

So at the moment Auto-route's `Get item` runs:

| | |
|---|---|
| `Author` | the uploader |
| `Editor` | **the approver** |

> ⚠ **`Editor` MEANS "LAST PERSON TO TOUCH THIS", NOT "APPROVER".** They coincide because approving
> is normally the last thing that happens to a pending item — but anyone editing the document's
> metadata *after* approval and *before* the flow polls takes that place. The window is one polling
> interval and pending items are not edited by third parties, so this is the best signal available.
> **Do not label it "the approver" in the flow.**

> ⚠ **READ IT FROM `Get item`, WHICH RUNS BEFORE THE STAMP.** Auto-route later writes `Author`,
> `Editor` and `Created` onto the **copy** via `validateUpdateListItem` (§4.3 of the 2026-08-08
> spec), setting `Editor` to the ORIGINAL uploader so attribution survives the move. Read `Editor`
> after that and every document looks self-approved, and **every uploader stops being notified** — a
> silent, total failure of the feature this is meant to trim.

---

## 2. The change

The email is already wrapped in a Condition (added 2026-08-22) suppressing it for bulk imports.
**Add a second row to that same Condition** rather than nesting another one — two nested conditions
around one action is how the second gets deleted by somebody tidying up.

Existing row, unchanged:

```
body('Get_item')?['BulkImport']        is not equal to        true
```

New row, joined with **And** — built in the fx editor, not typed into the value box:

```
@or(
  empty(coalesce(body('Get_item')?['Author']?['Claims'], '')),
  not(equals(
    body('Get_item')?['Author']?['Claims'],
    body('Get_item')?['Editor']?['Claims']
  ))
)
```

compared to `true`.

Read as a sentence: **send if we cannot tell who uploaded it, or if the uploader is not the person
who approved it.**

> ⚠ **COMPARE `Claims`, NOT `Email`.** A guest account can have no email at all (recorded 2026-08-23
> on the Group Management lookup). Two blank emails compare EQUAL, which would suppress the
> notification for every guest uploader — and guests are how this whole system is tested.

> ⚠ **THE `or(empty(...))` CLAUSE IS THE FAIL-OPEN, AND IT IS THE POINT.** If `Author.Claims` cannot
> be read, both sides are null, compare equal, and the email is suppressed. A missing notification is
> far worse than a redundant one: the uploader is left wondering whether their document went through
> and nothing anywhere tells them it did. **When in doubt, send.**

> ⚠ **DO NOT REPLACE THE `BulkImport` ROW.** It suppresses ~50 pointless emails per import, and its
> operator was set the wrong way round once already — which silences every NORMAL upload, a failure
> nobody reports for weeks. Read both rows back as a sentence before saving.

---

## 3. ⚠ The HC clone

`HC Auto Route` is a Save As of Auto-route and needs the **same** row. Every HC clone in this project
has shipped with a reference nobody swapped — six of them in `HC Auto Route` itself, and one in
`HC folder approval` that failed silently for six days.

Nothing library-specific is being added here, so this really is one expression copied — but **check
that `Get item` in the HC flow is still named `Get_item`**. A Save As sometimes renames actions, and
`body('Get_item')` against an action called `Get_item_1` resolves to null, lands in the `empty()`
branch, and sends the email regardless.

> ⚠ **FAILING OPEN MAKES THAT MISTAKE INVISIBLE:** HC approvers keep receiving emails and nobody
> reports a bug, because the old behaviour is what they are used to. Read the expression back in
> Code view.

---

## 4. What is deliberately NOT changed

- **The REJECTION email.** An approver rejecting their own upload is still notified. The client did
  not raise it, rejection is rarer, and the note usually carries a reason worth re-reading.
  ⚠ **Ask** rather than assuming symmetry.
- **`NotifyApprovers`** — the inbound *"a document needs your approval"* mail. A Head of Unit
  uploading into their own unit is emailed asking themselves to approve it, which is arguably the
  same complaint one step earlier. Same fix shape, different flow; raise it.
- **The bulk-import auto-approve flow.** It approves as the service account, so `Author` and `Editor`
  differ and the new clause passes — but the existing `BulkImport` row already stops those emails
  before the send is reached.

---

## 5. Testing

All four, in this order. The first is the change's blast radius; the last two are what it must not
break.

1. **PIC uploads, Head of Unit approves.** Email arrives, as today. *If this fails, revert — it is
   the case the email exists for.*
2. **Head of Unit uploads and approves their own document.** **No email.**
3. **Head of Unit uploads with `autoApproveOwnUpload` on.** No email — same condition, different
   route to it. ⚠ Confirm this is wanted; it reverses the 2026-08-24 decision.
4. **Bulk import of two files.** No emails, and the run shows the condition failing on the
   `BulkImport` row rather than the new one.

Then repeat 1 and 2 against `HC Approval Document` for `HC Auto Route`.

> ⚠ **CHECK THE RUN HISTORY, NOT JUST THE INBOX.** A flow that stops firing leaves no run at all, and
> *"no email arrived"* looks identical to *"the condition worked"*. Open the run and read which
> branch it took.

---

## 6. On SDG's tenant

**Sign in as the SERVICE ACCOUNT before touching either flow.** A flow runs under its connection, and
the connection is created implicitly by the first action — so a flow built by a person stops
**silently** when that password changes. Editing an existing action does not re-create the
connection; any action ADDED does.
