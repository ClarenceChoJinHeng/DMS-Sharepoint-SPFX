# CRS — Auto-approve bulk imports · Power Automate build runbook

**Date:** 2026-08-22 · **Status:** not yet built
**Belongs with** `2026-08-08-auto-route-flow-and-draft-isolation.md` (the record of the other two
flows — read its §0 and §5 before building anything) and
`2026-08-22-bulk-upload-for-uploaders-design.md` (why this exists).

This is the **third** flow on the approval library. It exists because Bulk Upload became an uploader
tool: an uploader holds Read on `Documents` and cannot file there, so bulk imports go through the
approval library like everything else — and this flow approves them so nobody has to click Approve
fifty times for documents that were approved years ago.

It approves. It does **not** copy anything. The existing, verified **Auto-route** flow does the move.

---

## 0. Sign in as the service account FIRST

⚠ Non-negotiable, and it cannot be repaired afterwards. A flow runs under its **connection**, and the
connection is created implicitly by the **first action** you add — so whoever is signed in at that
moment is baked in for the life of the flow. Built as a person, it stops **silently** the day that
password changes, and presents months later as *"my colleague's bulk import never arrived"*.

Full reasoning in §0 of the Auto-route spec. On the SDG tenant this applies from the first action.

## 1. Prerequisites

| # | Thing | How to confirm |
|---|---|---|
| 1 | App is **1.0.225.0 or later** | Site Contents → the app → ⋯ → Details. ⚠ Read the INSTALLED version, never the catalog listing. |
| 2 | Reconciliation has been run since that deploy | Its log says `Submission reference and bulk import columns: present on all CRS libraries ✓` |
| 3 | `BulkImport` exists on the approval library | `/_api/web/lists/getbytitle('Approval Document')/fields?$select=InternalName,TypeAsString&$filter=InternalName eq 'BulkImport'` → one row, `Boolean` |
| 4 | You have the approval library's **GUID** | `/_api/web/lists/getbytitle('Approval Document')?$select=Id` |
| 5 | Auto-route is ON and healthy | Its run history shows recent successes |

⚠ **If step 2 or 3 fails, stop and reconcile.** Without the column Bulk Upload writes no marker, this
flow never fires, and the files simply wait in the approval queue. That is the designed safe
degradation — but building the flow first and wondering why nothing happens is an hour wasted.

## 2. Create the flow

**Name:** `CRS — Auto-approve bulk imports in Approval Document`

Name it for what it does. The Auto-route flow is called *"…approved Pending **folders**…"* while it
routes **files**, and that name is exactly how folder logic ends up in the wrong flow.

### 2.1 Trigger

**When an item is created or modified** — SharePoint connector.

- **Site Address:** the CRS site
- **List Name:** the library **GUID** as a *custom value*. ⚠ The picker does not list document
  libraries (§5.3 of the Auto-route spec) — typing the title will not work; paste the GUID from
  prerequisite 4.

⚠ **NOT "When an item is created", even though that is what the folder-approval flow uses.**

Bulk Upload uploads the file and *then* tags it in a second call. At creation the item has **no
`BulkImport` value**, so a created-only trigger would test the condition before the marker exists,
never fire, and leave **no run history at all** — a flow that never triggers records nothing, which
is the hardest failure in this system to notice. The marker arrives on the first **modify**.

### 2.2 Trigger condition

Settings → **Trigger conditions** → add one row:

```
@and(equals(triggerOutputs()?['body/{IsFolder}'], false), equals(triggerOutputs()?['body/BulkImport'], true), not(equals(triggerOutputs()?['body/{ModerationStatus}'], 'Approved')))
```

**All three clauses are load-bearing. Removing any one causes a different failure:**

| Clause | Remove it and… |
|---|---|
| `{IsFolder}` is `false` | every folder Bulk Upload ensure-creates is approved by this flow, fires Auto-route, and fails `Copy file` with `NotFound` — the 2026-08-13 defect, verbatim |
| `BulkImport` is `true` | ⚠ **it approves EVERY file in the approval library the moment it is touched — abolishing approval for the entire system, silently, on a flow that reports success every time** |
| `{ModerationStatus}` is not `Approved` | the MERGE below modifies the item, which re-fires this same trigger, whose first two clauses are still true → it loops until Power Automate's loop protection stops it, burning the daily quota |

⚠ **The middle row is the one to re-read before saving.** Nothing on any screen would look wrong. The
files route, the uploaders are happy, and content approval has quietly ceased to exist.

**If `{ModerationStatus}` is not offered on the trigger** — it is available on `Get item` for certain,
and on the trigger it is expected but unverified — do **not** ship without a stop condition. Use the
two-clause trigger condition and add, as the *first* action:

```
Get item                          (same library, Id = triggerOutputs()?['body/ID'])
Condition   body('Get_item')?['{ModerationStatus}']  is not equal to   Approved
└─ True → the MERGE in §2.3
```

That is exactly the shape Auto-route already uses in its §4, so it is known to work here. It costs one
extra action per run.

### 2.3 The one action

**Send an HTTP request to SharePoint**

- **Method:** `POST`
- **Uri:** `_api/web/lists(guid'<library-guid>')/items(@{triggerOutputs()?['body/ID']})`
- **Headers:**

| Key | Value |
|---|---|
| `Accept` | `application/json;odata=nometadata` |
| `Content-Type` | `application/json;odata=nometadata` |
| `X-HTTP-Method` | `MERGE` |
| `IF-MATCH` | `*` |

- **Body:**

```json
{ "OData__ModerationStatus": 0 }
```

⚠ **`0` unquoted — it is an integer.** And ⚠ **the header keys must NOT include a colon.** Typing
`Accept:` creates no header at all, and the symptoms look unrelated to each other: responses come back
`odata=verbose` so every `body(...)?['Field']` is null, `validateUpdateListItem` reports
`HasException: false` and changes nothing, and a MERGE goes as a plain POST. Two hours on 2026-08-08.
The colon is visible in the action's raw **Inputs** and nowhere else. Full account in §5.1 of the
Auto-route spec.

⚠ **Do not set any other property in this call.** SharePoint rejects a merge that changes moderation
status alongside other fields with HTTP 500 — *"You cannot change moderation status and set other item
properties at that same time"* — and in a moderated library any property write **re-pends** the item
(§3 of that spec, where this knocked a whole tree back to Pending).

### 2.4 Split-on tracking ID

If the flow refuses to save with *"The 'clientTrackingId' value is not valid…"*, put this in the
**Split On Tracking Id** box:

```
@{triggerOutputs()?['body/ID']}
```

⚠ **Cleared is not the same as not there** — the UI renders blank and absent identically. And ⚠ the
tracking-ID box and the trigger-condition box sit in the same panel and look alike: the condition is
the row with an **✕** beside it. Pasting the tracking ID over the condition silently removes your
guards, which for this flow means the middle row of the table above.

## 3. Test it, in this order

Do not skip test 2. It is the one that proves approval still exists.

| # | Do | Expect |
|---|---|---|
| 1 | Bulk-upload **one** file to a unit you can reach | approved within seconds; arrives in `Documents` at the mirrored path; `Created By` is you |
| 2 | ⚠ Upload a file through the **normal upload form** | **stays *Waiting for Approval*.** If it is auto-approved, the `BulkImport` clause is missing or wrong — turn this flow OFF immediately and fix the condition |
| 3 | Check this flow's run history after test 1 | **one** succeeded run. Several, or a run every few seconds, means the `{ModerationStatus}` stop clause is not working |
| 4 | Check Auto-route's run history | one succeeded run for the same file |
| 5 | Bulk-upload one file, then a normal upload approved, then a normal upload rejected | **no email** / **approval email** / **rejection email** — see 3.5. The middle one is the regression test |
| 6 | Bulk-upload **five** files | five approved, five routed, and the source folder left empty |
| 7 | Sign in as a **second** uploader in another unit | they see only their own units in the picker, and their import routes to their own folder |

## 3.5 ⚠ SUPPRESS AUTO-ROUTE'S APPROVAL EMAIL FOR BULK IMPORTS

Found on the first live test (2026-08-22), and it is not cosmetic. Auto-route ends by emailing the
uploader *"Great news! Your document has been approved and is now available in the Document library."*

For a normal upload that is correct — a person reviewed it. **For a bulk import nobody approved
anything**, and a fifty-file historical import sends fifty of those emails to the person who just
pressed Upload.

In the **Auto-route** flow (not this one), wrap its final `Send an email from a shared mailbox (V2)`:

1. On the email card: **... -> Cut** — never rebuild it, the body carries six expressions
2. In its place add a **Condition**:
   - left (fx expression): `body('Get_item')?['BulkImport']`
   - operator: **is not equal to**
   - right: `true`
3. **Paste an action** into the **True** branch — the email returns fully configured
4. Leave **False** empty

⚠ **THE OPERATOR IS THE WHOLE THING, AND IT WAS SET THE WRONG WAY ROUND ON THE FIRST ATTEMPT.**
`is equal to` with the email in True does the exact opposite: it emails only for bulk imports and
**silences every normal upload**. Read it back as a sentence — *"if BulkImport is not equal to true,
send the email"* — before saving. An uploader who is never told their document was approved does not
report it for weeks, and by then nobody connects it to this change.

⚠ If the `Get item` action has been renamed, `body('Get_item')` must be renamed with it. Power
Automate builds that reference from the display name with spaces as underscores.

**THE REJECTION EMAIL IS DELIBERATELY LEFT ALONE.** The approval email is suppressed because no human
approved it. **A rejection is always a human act** — if a bulk-imported file is rejected, an approver
opened the queue and did it on purpose, and the uploader needs to know their import did not land. A
bulk file only reaches that branch when auto-approve did not fire (flow off, marker missing, quota
exhausted), which is precisely the designed safe degradation, and the email is right in it.

## 3.6 ⚠ DO NOT SET THE MARKER COLUMN TO `Hidden` ON AN APPROVAL LIBRARY

`BulkImport` is a Yes/No field on a library an uploader can edit, so hiding it from the forms looks
like sensible hardening. **`Hidden: true` breaks the flow.**

The write is unaffected — `validateUpdateListItem` ignores `Hidden`, and a REST read confirms
`BulkImport = true` on the item. But the SharePoint connector omits hidden fields from its trigger
output, so `triggerOutputs()?['body/BulkImport']` is **null** and the condition can never match. No
error on the write; no run on the flow. Both ends silent.

Verified 2026-08-22 by unhiding one library at a time: normal bulk import fired with the column
visible, HC did not with it hidden.

- **Approval libraries: keep `BulkImport` VISIBLE.**
- `Documents` / `HC Documents`: hidden is fine, no flow watches them.
- The narrower `ShowInEditForm: false` may keep it off the edit form without touching the payload —
  **verify it the same way** (one library, bulk-upload, confirm the flow still fires) before applying
  it anywhere else. If it also breaks the trigger, leave the field visible: a working auto-approve
  beats a deterrent against a deliberate tick.

## 4. What this flow does NOT do, and must not be extended to do

- **It does not copy, stamp or delete.** Auto-route owns all of that and is verified; a second copier
  is how two flows race over one file.
- **It does not touch folders.** Approving folders is the other flow, with the opposite `{IsFolder}`
  polarity. Keeping them separate is what makes each one's polarity checkable in isolation.
- **It does not run on `Documents`.** Nothing there needs approving — moderation must be OFF on that
  library — and this flow watching it would be a loop with the copy.

## 5. HC is not covered, and cannot be until the HC pair has flows

⚠ **THIS SECTION SAID "the HC Auto-route flow has never been built". IT WAS WRONG** — corrected
2026-08-22 after the flow list was actually looked at. `HC Auto Route` was built and the HC vertical
verified end to end on **2026-08-19** (commit `e58e97c`). The claim came from a stale line in CLAUDE.md
and was repeated three times in one session before anyone checked. **Read the flow list before
asserting a flow does not exist.**

So HC bulk import needs **one** more flow, not two:

- **HC auto-approve** — a `Save As` of the flow above, with the trigger **List Name** and the HTTP
  action's **Uri** repointed at the `HC Approval Document` GUID. The trigger condition is unchanged,
  byte for byte.

⚠ **HC Auto Route is a CLONE of the normal Auto-route, so it sends its own approval email** — it needs
the same `BulkImport is not equal to true` condition from §3.5, or every HC bulk import emails per file.

⚠ **`HC folder approval` MUST STAY ON — an earlier draft of this section said to turn it off, and that
was wrong.** It rested on CLAUDE.md's claim that draft isolation was reversed on *all* libraries on
2026-08-19. Verified on site 2026-08-22: `HC Approval Document` still reads **"Only users who can
approve items (and the author)"**. So a below-Unit folder created by one HC-cleared PIC arrives Pending
and is invisible to the next, who then cannot reach their own file inside it — and cannot self-fix,
since that needs `ApproveItems`. Bulk Upload ensure-creates those same folders, so it depends on this
flow too. **Turning off the folder-approval flow is correct for the NORMAL pair only.**

Check the setting before acting on any instruction of this kind: Library settings → Versioning
settings → Draft Item Security. A site-wide claim in a doc is a claim about the libraries someone
actually opened.

Two defects already recorded against HC Auto Route, unrelated to bulk import: its approval email links
to the SOURCE library that its own delete has already emptied, and `Created` is stamped ~8 hours early
on both routing flows.

## 6. Cost, and what to tell the client

**Two flow runs per file minimum** — this one, then Auto-route. In practice more: *created or modified*
fires several times per upload, and each fire that fails the trigger condition still counts against the
quota.

`MAX_FILES` is **50**, which bounds one press. A large historical import should be **spread across
days**, and this is not housekeeping advice: ⚠ an exhausted daily quota means the next real approval is
not routed, **silently**, and the person waiting for it has no way to tell.

Two more things the client must be told, neither of which the software enforces:

1. **Bulk Upload has no idea what "historical" means.** A brand-new document put through it reaches
   `Documents` unreviewed. The rule is theirs to communicate to uploaders.
2. **`Created By` becomes whoever bulk-uploaded**, not the original author. Auto-route preserves the
   uploader faithfully; it cannot know who wrote a document years ago. The real date is captured in
   `DocumentDate`, which the form asks for.

## 7. Turning it off

Turn the flow **off**, do not delete it — an errored flow burns quota on every upload, and a deleted one
takes its run history with it. With it off, bulk-uploaded files simply wait in the approval queue for a
Head of Unit, which is a working system and not a broken one.
