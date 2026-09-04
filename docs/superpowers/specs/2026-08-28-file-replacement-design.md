# Deliberate file replacement, and auditing it — design

**Date:** 2026-08-28 · **Status:** DESIGN ONLY, nothing built
**Supersedes:** the replacement half of 1.0.282.0/283.0 (CLAUDE.md, *"A FILED DOCUMENT CAN BE
REPLACED THROUGH APPROVAL; A PENDING ONE NEVER"*)

Client, 2026-08-28:

> "now they want us to allow them to replace file on staging, so the decision to replace will be on
> their [side]"

> "For the ApprovalDocument.aspx, just drop the guard and let them override, just allow
> ApprovalDocument.aspx to crosscheck and notify them its going to be overwritten and let them do it,
> because the overwritten is needed."

> "For the Bulk Upload, it will not allow client to overwrite the staging"

> "Ensure this is applied to both HC and Non HC Library."

> "Now ensure it also detect it is overwritten in Audit Log"

---

## 1. What this reverses, and why the old rule is void rather than overruled

1.0.282.0 was explicit: **a pending draft is untouchable, by anyone.** That was the client's own
instruction the day before — *"Never allow anyone to overwrite a pending file in staging."*

The both-libraries rule came from the same sentence:

> "if the file is existing in document library and also exist in staging library, don't allow them to
> overwrite, **which is basically the don't allow them to overwrite the staging file**"

⚠ **That rule existed ONLY because staging was untouchable.** Its stated reason is the clause after
the comma. Making staging replaceable removes the premise, so the rule must be re-decided rather than
inherited — it is not independently motivated.

**The new request has real merit.** Today an uploader filing a corrected version of a document that is
already pending can only rename it, which leaves two near-identical drafts and an approver guessing
which is current. Replacing is usually the correct intent.

## 2. ⚠ The danger this creates, and the two things that contain it

Draft Item Security on both approval libraries is *"Only users who can approve items (and the
author)"*. **Uploader B cannot see uploader A's pending file.** The pre-check lists the folder and
finds nothing; only `Files/Add` knows, which is the entire reason the Add-failure branch exists
(1.0.268.0).

So Replace in staging lets B destroy A's unreviewed upload — a file B never saw, from someone B may
not know.

**Two containments, and the first is a hard prerequisite:**

1. ⚠ **Version history MUST be ON for `Approval Document` and `HC Approval Document`.** Then A's
   content survives at 1.0 and is recoverable. Confirmed for `Documents` (major, 500 kept); the
   approval pair has **never been checked**. `scripts/check-hc-setup.js` reports `EnableVersioning`
   for all six libraries — run it before building. **If versioning is off there, this feature destroys
   content with no recovery and must not ship until it is on.**
2. **An audit row naming the replacement** — §7.

**A third, already in place:** A's submission record stops resolving, so My Submissions shows their
file as **Deleted** (1.0.310.0). After the fact, but not silent.

## 3. Upload form — both checks stay, four cases

The client corrected an earlier instruction: the Documents check is **kept**.

| In staging | In Documents | Offer |
|---|---|---|
| no | no | uploads normally |
| **yes** | no | rename, or **replace the pending draft** |
| no | **yes** | rename, or **send for approval as a replacement** (unchanged) |
| **yes** | **yes** | ⚠ **one Replace button** — see below. **Client confirmation still open.** |

### 3.1 ⚠ The both-clash case is ONE action, not two

Replacing the pending draft means that draft goes through approval, and Auto-route then replaces the
filed document anyway (`nameConflictBehavior: 1`). **Choosing "replace staging" IS choosing "replace
Documents, later."** They are not independent, so two buttons would imply a combination that cannot
exist.

One button, with copy naming both consequences:
> *Replaces a pending upload now, and replaces the filed document when it is approved.*

### 3.2 The check order changes meaning

Today the staging check runs **first and returns immediately** — deliberately, so the dialog cannot
offer a replacement for a file whose real blocker is a pending draft (1.0.283.0). That short-circuit
must go: both facts are now actionable, and the uploader's decision depends on both. The dialog must
report **which library or libraries** each file clashed in.

### 3.3 ⚠ `overwrite=true` returns, narrowly

It was deliberately deleted from the codebase (1.0.271.0), and `spGroups.ts:508` records the reasoning
for that class of removal: a parked capability doing exactly the forbidden thing invites the next
session to wire it back up.

It comes back **only** on the upload form's explicit Replace path, reached only after the uploader has
seen the dialog and chosen. Everywhere else stays `overwrite=false`:

- `BulkUpload.tsx:2246` — unchanged
- the upload form's normal path — unchanged

⚠ **Do not reintroduce a per-file "Replace?" prompt.** The 2026-08-15 removal reason still holds: a
batch runs unattended so nobody can answer it, and an admin clicking through the fifth prompt is
somebody's upload gone. The decision is made once, in the dialog, before the run.

## 4. The popup

- **Remove the header subcopy** *"It can be uploaded under a free name instead — nothing already filed
  will be changed or replaced."* It is about to be false.
- ⚠ **That subcopy was carrying the explanation**, so each row must now state its own situation —
  which is more useful anyway, and is where the requested batch and path naturally live:

```
Batch 1 · GHO › GCA › EG › 2024 › Agreement
sd - d - d - 26-08-26.pdf  ->  sd - d - d - 26-08-26 - Copy.pdf
A pending upload with this name already exists.
```

```
Batch 2 · MHO › TREAS › AP › 2025 › Invoice
report.pdf  ->  report - Copy.pdf
Already approved and filed. Check this is not the same document.
```

- **Scrollable list, capped height.** ⚠ Cap the LIST, not the dialog, and keep the buttons outside the
  scroll box — the same rule the abbreviation editor follows, where the warning, collision banner and
  Save all stay outside its 58vh box.
- ⚠ **Check for absolutely-positioned descendants before adding `overflow`.** Three screens here have
  been bitten by a scroll container clipping an absolutely-positioned child (the people-picker in
  `GroupManager`, the info panels in `.dms-staged-row`, the add-box dropdown). If a row gains a
  tooltip later, this breaks silently.

## 5. `ApprovalDocument.aspx` — check, warn, allow

Client: *"just drop the guard and let them override … crosscheck and notify them its going to be
overwritten and let them do it, because the overwritten is needed."*

`documentsFileClash` currently **refuses**, failing closed. It becomes a **warning with a choice**,
not a removal:

- Keep the check, and keep it re-evaluated immediately before submitting — upload-time and
  approval-time are different moments (the 1.0.235.0 race).
- On a clash: state plainly that approving will replace the filed document, name it, require a
  deliberate confirm.
- ⚠ **A failed check must still warn, not silently allow.** "Could not verify" is a third answer.

⚠ **Resolve the destination library through `approvedLibSeg()`, never `DOCUMENTS_URL_SEGMENT`.** The
2026-08-24 defect had it hardcoded, so approving an HC document checked the normal library. That
helper returns **blank** rather than falling back when the HC pair is unresolved — deliberately unlike
`approvedLibTitle()` beside it, because pointing an HC document's check at the open library is the
worst available outcome.

## 6. Bulk Upload — never overwrites staging

| In Documents | In staging | Offer |
|---|---|---|
| yes | no | rename, or send for approval as a replacement |
| yes | **yes** | **rename only** |
| no | **yes** | **rename only** |
| no | no | uploads normally |

Bulk can only offer replacement when the name is free in staging, so it never overwrites a pending
file. **The both-clash rule survives here for its original reason** — Bulk still cannot overwrite
staging — which is why it diverges from the upload form. Pin that asymmetry with a test, or someone
will "make them consistent".

## 7. Auditing a replacement — four paths

⚠ **The constraint that shapes all of them: `writeAudit` from the app is refused for everyone who
matters.** `CRS Audit Log` restricts writes to Owners and the service account, so a PIC's or Head of
Unit's session records **nothing** — proven 2026-08-26, and the reason the request-activity flow had
to exist. **Every path must end in a flow.**

| # | Path | Detection |
|---|---|---|
| 1 | Upload form replaces in staging | App stamps `ReplacedOnUpload`; `Audit — approval activity` reads it and writes `Replaced` instead of `Uploaded` |
| 2 | Approver approves over a filed document | App stamps `ApprovedAsReplacement`; Auto-route reads it |
| 3 | Auto-route's copy replaces, chosen or not | Flow-side HTTP GET on the destination **before** `Copy file` |
| 4 | Someone replaces directly in the library UI | ⚠ **Out of reach.** No app involvement; catching it needs a modify-trigger flow that would also fire on every metadata stamp. Version history is the only record. Say so to the client. |

**Paths 2 and 3 are complementary, not duplicates.** The GET answers *did a replacement happen*; the
stamp answers *did someone choose it*. The second matters for the native Approve/Reject command, which
bypasses every app-side check and would otherwise produce an unattributed replacement.

### 7.1 The two new columns

`ReplacedOnUpload` (Yes/No) and `ApprovedAsReplacement` (Yes/No), with the displaced file's author and
timestamp carried in `Details`.

- ⚠ **In all four libraries**, or Auto-route's copy strips them — the `Remark`/`LegallyPrivileged` gap
  of 2026-08-10 and the `SubmissionId` lesson of 2026-08-22.
- ⚠ **VISIBLE on both approval libraries.** `Hidden: true` removes a field from the Power Automate
  trigger payload entirely, so the flow never fires and leaves **no run history** to inspect — silent
  at both ends. This cost a day on `BulkImport`.
- ⚠ **Written conditionally**, via `libraryHasColumns` in `shared/optionalColumns.ts`. One unknown
  field name fails the WHOLE `validateUpdateListItem` call, losing every column rather than the
  missing one. Degrading to today's behaviour is always correct; losing a document's metadata never is.
- Asserted by reconciliation on every run, like the other optional columns.

### 7.2 `EventType: Replaced`

Safe to add: **`EventType` is Text, never Choice** — precisely so a new value cannot silently lose
every row of that type. It joins the viewer's event-type filter automatically.

`Details` names what was displaced: previous author, previous modified date, and whether the
replacement was chosen (`ReplacedOnUpload`) or incidental (detected by the GET alone).

### 7.3 ⚠ The audit log records that it happened; version history holds the content

Both are needed and neither substitutes for the other. Versioning on the approval libraries (§2) is
the prerequisite for path 1 being survivable at all.

## 7A. My Submissions — five changes (client, 2026-08-28)

### 7A.1 A replaced submission reads **Cancelled**

⚠ **DELETION AND REPLACEMENT ARE INDISTINGUISHABLE TODAY.** A record greys out when its
`SubmissionFileId` stops resolving, and that happens for both. "Cancelled" therefore needs a signal
that does not exist yet.

**Mechanism:** before `Files/Add(overwrite=true)`, read the existing file's `SubmissionFileId`; after
the replace succeeds, mark THAT record `Cancelled`.

- Same person replacing their own earlier upload — clean, and the word is exactly right.
- Someone replacing a colleague's pending draft — this writes to **their** row. The list ACL permits
  it (add/edit for the uploader groups; item-level "edit own only" was rejected because approvers
  must be able to record decisions on rows they did not raise).

⚠ **THE CLIENT CHOSE "Cancelled" WITH THE OBJECTION STATED, so do not "correct" it.** It reads oddly
in the cross-user case — the person cancelled nothing, they were overwritten — and **"Replaced"** was
offered as the more honest word. Their answer: *"I know its weird but client want it to be
Cancelled."*

**Inference was REJECTED** as the mechanism ("another record has the same name and a later date").
It is wrong whenever a file is genuinely deleted and the name later reused, and it would state a
guess as a record — the failure mode this project already names for inferred submission grouping.

### 7A.2 A Bulk Upload tab

⚠ **NO DATA CHANGE NEEDED — this is display only.** Bulk Upload already mints `SUB`, `BAT` and a
per-file `SFI` reference and calls `writeSubmissionRecord`. The CLAUDE.md note saying it was
"deliberately NOT stamped" is stale twice over: that reasoning rested on Bulk Upload being admin-only
and writing straight to the approved side, and neither has been true since 2026-08-22.

Distinguish a bulk run by its record `Source` rather than by guessing from shape.

### 7A.3 A bulk run opens straight to its files

A bulk run has ONE destination, so the batch level is a list of one and clicking through it is a
click that answers nothing. Skip it — level 1 goes directly to the file list.

### 7A.4 Scroll the Document details list

A run can hold 50 files. ⚠ Cap the LIST and keep anything below it outside the scroll box, the same
rule as the clash dialog and the abbreviation editor — and check for absolutely-positioned children
before adding `overflow`, which has bitten three screens here.

### 7A.5 "Head of Unit" → "Approver"

~12 user-facing strings in `MySubmissions.tsx`. ⚠ Change them ALL, not just the footer the client
pointed at: half the page saying "Head of Unit" and half saying "Approver" is worse than either.
Code comments keep the persona name, which is what `PERSONAS` actually calls it.

## 8. Risks to state to the client

1. **A pending file can now be destroyed by someone who cannot see it.** Recoverable only through
   version history.
2. **The approver can replace a filed record.** By explicit instruction, and warned.
3. **Path 4 is unaudited** — a direct replace in the library UI leaves only a version.
4. ⚠ **Version history on the approval libraries is unverified.** Check before building.
5. ⚠ **Auto-route's Replace has never been proven to write a version rather than destroy content.**
   It is on the outstanding list and is the single most load-bearing unverified assumption behind the
   whole replacement model.

## 9. Build order

1. **Verify versioning** on `Approval Document` and `HC Approval Document` (§2). **Blocking.**
2. **The two columns** + reconciliation assertion + `libraryHasColumns` guard.
3. **Upload form**: four-case logic, the narrow `overwrite=true`, the stamp.
4. **The popup**: subcopy removed, per-row context, scroll.
5. **`ApprovalDocument.aspx`**: warn-and-allow, HC-aware destination.
6. **Bulk Upload**: the table in §6, `overwrite=false` unchanged.
7. **Flows**: the pre-copy GET in both routing flows; `Replaced` in the audit flows.
8. **Tests**: the six clash branches, the both-clash asymmetry between the two screens, the fail-open
   directions.

⚠ **Every clash change is SIX branches, not one** — three per screen (staging pre-check, approved-side
pre-check, Add-failure). The Add-failure branch is the one that fires for a concurrent uploader, i.e.
the only case the pre-checks are structurally blind to, and it has been missed twice.
