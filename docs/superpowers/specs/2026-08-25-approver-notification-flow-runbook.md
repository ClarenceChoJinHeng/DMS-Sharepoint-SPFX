# Approver notification — flow runbook

**Date:** 2026-08-25
**Status:** `CRS - Notify approvers` BUILT AND VERIFIED 2026-08-25 (ClarenceDMSTesting).
`CRS - Notify HC approvers` (built as `HCNotifyApprovers`) BUILT AND VERIFIED 2026-08-25.
See section 11 for what the build changed.
**Applies to:** two flows — `CRS — Notify approvers` and `CRS — Notify HC approvers`
**Touches:** no code. Power Automate only.

Client, 2026-08-25: *"can we build the notification to alert approval that the file has been uploaded
and in need of their approval"*. Today an approver learns about a pending document only by opening
the approval queue; nothing tells them one arrived.

---

## 1. The three decisions the client made

| Question | Answer | Why it matters |
|---|---|---|
| Bulk imports? | **Excluded** — *"not needed since its a historical file"* | They auto-approve. Nobody has to act, and a 50-file import would send 50 pointless emails. |
| Who is emailed? | The unit's approver group members, **capped** — *"there shouldn't be so many people in the approver group but put a limit"* | The cap is a safety valve against a misconfigured group, not a policy. See §4. |
| HC? | **Yes, but only HC-cleared people** — *"just make sure only HC people get the file"* | A plain approver emailed about an HC document learns its filename and its unit. That is the disclosure the whole HC split exists to prevent. |

## 2. Two flows, and the separation is structural

One flow per approval library. **The trigger is the library**, so a document can only ever reach the
flow that watches the library it is in — an HC document cannot be picked up by the normal flow, and
no condition has to be written to keep them apart.

| | Normal | HC |
|---|---|---|
| Trigger library | `Approval Document` | `HC Approval Document` |
| Group Map `Role` filter | `APR` | `APRHC` |
| Group named | `<SEG>_<DEPT>_<UNIT>_APPROVER` | `<SEG>_<DEPT>_<UNIT>_APR_HIGHLY_CONFIDENTIAL` |

⚠ **The `Role` filter is the only thing keeping HC private, so get it right.** `APR` is held by the
plain `hou` persona; `APRHC` only by `hou_hc`. Filtering the HC flow on `APR` would email every
ordinary Head of Unit about Highly Confidential documents.

## 3. Finding the approvers — via the Group Map, NOT the folder path

The tempting shortcut is deriving the group name from the path: every family is exactly two
permissioned tiers, so `/ApprovalDocument/GHO/GF/TAX/2024/...` implies `GHO_GF_TAX_APPROVER`.

⚠ **Do not.** Renaming a term abbreviation renames the FOLDER and leaves the group name alone
(CLAUDE.md, 1.0.162.0). The derived name would stop matching, `Get group members` would return
nothing, and **the flow would carry on succeeding while nobody was ever notified** — a silent failure
with no run-history error to find.

Use the document's own term instead, which no rename touches:

1. **Get item** → read `UnitTid` (the unit's term GUID, e.g. `a9d818bd-a4af-4cf5-afe3-eb3a21edb74c`)
   and `BulkImport`.
2. **Send an HTTP request to SharePoint** — find the approver group for that term:
   ```
   GET _api/web/lists/getbytitle('CRS Group Map')/items
       ?$select=GroupId,GroupName
       &$filter=Role eq 'APR' and UnitTermGuid eq '<UnitTid>'
       &$top=5
   ```
   (`APRHC` in the HC flow.) Same join reconciliation uses, so the two cannot disagree about who
   approves a unit.
3. **Get the members** of that `GroupId`:
   ```
   GET _api/web/sitegroups(<GroupId>)/users?$select=Title,Email,LoginName
   ```

⚠ **`Accept: application/json;odata=nometadata` on BOTH calls.** Without it SharePoint answers
verbose (`{"d":{"results":[…]}}`, no `value` key) and every expression reading the response returns
null — while the action itself goes GREEN. That cost two hours on the audit flows.

⚠ **Wrap every expression that reads a response**, e.g.
`first(coalesce(body('GetApproverGroup')?['value'], createArray()))?['GroupId']`. A wrong response
SHAPE is not an action failure, so `Run after` does not protect against it.

## 4. The recipient cap

Cap the recipient list at **10**. It is a guard against a group nobody has curated — an approver
group is meant to hold one or two people, and a group that has somehow acquired forty is a
configuration problem, not a reason to send forty emails.

**Send to the first 10 rather than skipping the send.** Notifying some of the approvers is better
than notifying none, and the alternative fails in the direction where the document sits unnoticed.
Add a line to the email body when the cap bites, so the state is visible to a human:

> This unit's approver group has 40 members; only the first 10 were emailed. That group probably
> needs reviewing.

Expression: `take(body('GetMembers')?['value'], 10)`.

## 5. Conditions on the trigger

**Trigger:** *When an item is created or modified* on the approval library.

Three clauses, all required:

```
@and(
  equals(triggerOutputs()?['body/{IsFolder}'], false),
  not(equals(triggerOutputs()?['body/BulkImport'], true)),
  equals(triggerOutputs()?['body/{ModerationStatus}'], 'Pending')
)
```

- **`{IsFolder}` false** — the upload form ensure-creates Year / Document Type / below-Unit folders,
  and without this every one of them emails the approvers. Same guard Auto-route needs.
- **`BulkImport` not true** — the client's decision above.
  ⚠ **`BulkImport` must be VISIBLE on the library**: `Hidden: true` removes a field from the trigger
  payload entirely, so the expression reads null, the clause passes, and bulk imports email everybody.
  That is a real trap already hit once (2026-08-22).
- **`{ModerationStatus}` is `Pending`** — a STRING from the connector, not the integer every REST call
  in this codebase reads, and a rejection reads `"Denied"`. Comparing to `0`/`1` matches nothing.

⚠ **Created *or modified*, not created alone.** The upload form uploads and THEN tags, so at creation
the metadata — including `UnitTid`, which this flow needs — is not there yet. But that means the
trigger fires twice per upload, so §6 is not optional.

## 6. Do not email twice

The upload form's upload-then-tag sequence fires the trigger at least twice for one document. Without
a guard the approver gets two emails per upload, which is how people start ignoring them.

Cheapest reliable check: **only send when `UnitTid` is present** — the second firing is the one that
has it, and the first cannot be acted on anyway (no unit, no group, no recipients).

If duplicates still appear in practice, add a `CRS Audit Log` lookup for an existing
`ApproverNotified` row on this `ItemUniqueId`, the same dedupe the audit flows use — and, as there,
**a FAILED dedupe check should still send**. A duplicate email is something a person reads past; a
missing one leaves a document sitting unapproved.

## 7. The email

Keep it short. It is a nudge, not a report.

- **Subject:** `A document needs your approval — <Unit>`
- **Body:** the document name, the unit path, who uploaded it, and a **link to the approval page**
  (`/SitePages/ApprovalDocument.aspx?itemId=<ID>`; add `&lib=hc` for HC — item ids are per-LIST, so
  without it the page resolves the wrong library or a different document sharing that id).
- **Do not attach the file**, and for HC do not quote its metadata beyond the name and unit.

⚠ **Point the link at the approval PAGE, never the library view.** The page runs the destination and
clash guards; the library's own Approve/Reject command runs neither.

## 8. Building it

- **Sign in as the SERVICE ACCOUNT before the first action.** A flow runs under its connection, and
  the connection is created implicitly by the first action — the builder's account is baked in for
  life. Built as a person, the flow stops silently when that password changes.
- Build the normal flow first, verify, then **clone for HC** — and change **three** things: the
  trigger library, the `Role` filter (`APR` → `APRHC`), and the `&lib=hc` on the link. Every HC clone
  in this project has produced faults from a library reference that was never swapped; six of them in
  the HC Auto Route.

## 9. Verifying

1. Upload one document as a PIC → the unit's approver receives one email, with a working link.
2. Upload to a unit whose `_APPROVER` group is **empty** → no email, no failure. (An empty group is a
   valid state; the run should complete quietly.)
3. **Bulk import several files** → no emails at all.
4. Upload an HC document → **only** the `_APR_HIGHLY_CONFIDENTIAL` members are emailed. Confirm with a
   plain approver account that nothing arrives. This is the one that matters.
5. Approve something → no notification (the trigger's `Pending` clause).

## 11. As built - what this runbook got wrong

Built on ClarenceDMSTesting 2026-08-25 as `NotifyApprovers`. The design above is correct except for
one omission that stopped the flow dead, recorded here rather than silently patched into section 5.

### The action chain, as it actually runs

```
Trigger (created or modified, Approval Document, 3 clauses per section 5)
  GetDoc                    Get item -> UnitTid, BulkImport, Created By
  HasUnit                   length(UnitTid) > 0        <- also the duplicate guard, see below
    True:
      GetApproverGroup      HTTP: Group Map, Role eq 'APR' and UnitTermGuid eq '<UnitTid>'
      HasApprover           length(value) > 0
        True:
          GetMembers        HTTP: sitegroups(<GroupId>)/users
          MembersWithEmail  filter to rows carrying an Email
          PickEmails        take(..., 10)              <- the section 4 cap
          Recipients        join(..., ';')
          HasRecipients     length(body('PickEmails')) > 0     <-- NOT IN THE DESIGN
            True:
              Send an email (V2)
```

### WARN: section 9's test 2 is unbuildable from section 5

Section 9 requires that an empty approver group produce *"no email, no failure"*. **Nothing in section
5 produces that, and the flow built exactly to this spec does the opposite:** `Recipients` joins an
empty array to an empty string and the email action returns `Bad Request - To Field cannot be null or
empty`. The run lands in the failed list, which is precisely where a notification system goes to be
ignored.

Fixed with `HasRecipients`, wrapping the email:

```
length(coalesce(body('PickEmails'), createArray()))   is greater than   0
```

**The reusable point: a runbook that states an expected OUTCOME without naming the ACTION that
produces it will be built without that action.** The gap surfaces on the test case nobody reaches
first - here, the empty group, which is the normal state of a group recreated that morning.

### WARN: a skipped email and a broken filter are indistinguishable

When the guard fired, `GetMembers` / `MembersWithEmail` / `PickEmails` all ran green in 0s and produced
nothing - which is exactly what a wrong property name inside `MembersWithEmail` would also produce.
**Verifying the empty-group case proves the guard and says nothing about the chain.** Only section 9's
test 1, against a group with members, distinguishes them. Run it before believing any of this works.

### WARN: `0` typed into a condition's value box is TEXT

`length()` returns an Integer and the comparison fails outright on the type. **Enter the `0` through
the fx editor** on all three conditions. Cost a round of debugging on each.

### Two runs per upload is correct

Section 6's dedupe is `HasUnit`, and it works because the upload form uploads and THEN tags: the first
trigger firing has no `UnitTid` and stops there, the second sends. So the run history shows two runs
and one email per upload. **Do not read the second as a duplicate, and do not narrow the trigger to
created-only** - that fires before the metadata this flow reads exists.

### Verified

Uploaded into GCA > GCBC as another account. One email, subject
`A document needs your approval - Group Communications - Brand Communications`, naming the document,
the unit and the uploader; its link opened that document on the approval page; the approval went
through. Section 9 tests 1 and 2 pass.

**Test 4 - the HC one - also passes.** Uploaded into GCA > EG as an HC-cleared PIC: the HC approver was
emailed, approved through the bulk approve panel, and `HC Auto Route` moved the file into
`HC Documents/GHO/GCA/EG/2025/Approval Papers/Archive 2` with its metadata and `LegallyPrivileged`
intact.

WARN: **the negative half was proved by the DATA, not by a second account.** `GHO_GCA_EG_APPROVER` was
empty while `GHO_GCA_EG_APR_HIGHLY_CONFIDENTIAL` had one member - so a `Role eq 'APR'` filter would
have matched the empty group and sent nothing at all, and only `APRHC` could produce an email. The
email arriving is therefore proof the filter is right. **When a negative test is awkward to stage, look
for a unit whose existing data already forces the two branches to differ.**

**Tests 3 and 5 are still not run** (bulk imports produce no email; approving produces no email).

WARN: **THE CLONE IS FOUR EDITS, NOT THREE.** Section 8 says "the trigger library, the Role filter and
the link". The library is named in **two** actions - the trigger and `GetDoc` - and a clone with one
swapped and one not fires on HC documents while reading the normal library.

## 10. Out of scope

- **Reminders for documents pending more than a week.** Requested separately and deferred (memory
  `crs-approver-reminder-requested`); it is a scheduled flow over the queue, not this trigger.
- **Notifying the uploader.** Auto-route already emails them on approval.
- **A digest instead of one email per document.** Worth revisiting if volume annoys people, but with
  bulk imports excluded the remaining traffic is one email per genuine upload.
