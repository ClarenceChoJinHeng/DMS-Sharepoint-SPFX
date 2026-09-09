# `CRS — Notify request activity` — the six share/delete request emails, in ONE flow

Power Automate only. **No code change.** Builds item 2 of
`2026-08-28-email-bundling-and-templates-design.md` §7. That design carries the client's templates
6–11 verbatim; this is how to build them.

**✅ BUILT AND VERIFIED ON ClarenceDMSTesting, 2026-09-10.** Four emails, in one flow named
`CRS — Notify request activity`, built as a Save As of `NotifyApprovers`. ⚠ **Scope cut to FOUR
emails on 2026-09-09** (client: *"lets build the four emails now"*) — the two Rejected templates are
deferred, not cancelled.

**What was verified live, end to end:** the trigger fires on `CRS Requests`; all six Composes
evaluate; `NeedsApprover` routes pending vs decided; the Group Map approver lookup resolves a real
group; the requester is correctly excluded from the approver list; `HasRecipients` gates an empty
group; the two approver emails arrive with the right subject, body, unit, dates and working links;
the delete-approved email arrives at the requester saying 93 days; a **Cancelled** row correctly
sends nothing; and **`ApproverRole` resolved `APRHC` on an HC document** — the test that mattered.

⚠ **NOT exercised: the `ShareApproved` email (template 7).** Every other branch has run. That one is
structurally identical to the delete-approved email, which works, and its content was read on screen
— but it has never actually fired. Approve a share request to close it out.

⚠ **NOT built: the `Failed` case.** See §9 — it was hit live within the hour and the requester was
told nothing.

⚠ **`sendEmail: true` on `performShare` is UNCHANGED**, deliberately. Microsoft's own invite still
goes to the person being shared with, and still lands in spam for some recipients — that is the
pre-existing deliverability issue, not this flow. The tenant's SPF/DKIM/DMARC is the place to look,
not here.

---

## 1. Six emails, one flow — FOUR of them in scope

Templates 6–11 are six emails about one list. **Do not build six flows** — that is six copies of the
same trigger, the same approver lookup and the same recipient resolution, and the copy that drifts is
always the rarely-exercised one.

| `RequestType` | `Status` | Template | Goes to | In scope |
|---|---|---|---|---|
| Share | Pending | 6 | the unit's approver(s) | **yes** |
| Share | Approved | 7 | the requester | **yes** |
| Share | Rejected | 8 | the requester | deferred |
| Deletion | Pending | 9 | the unit's approver(s) | **yes** |
| Deletion | Approved | 10 | the requester | **yes** |
| Deletion | Rejected | 11 | the requester | deferred |

⚠ **A DEFERRED TEMPLATE MUST REACH `Skip`, NOT A FALL-THROUGH BRANCH.** `Rejected` is a real status
that fires this trigger, so leaving it unhandled in a chain of `if`s is exactly the mistake
`fromListItem` made in code (1.0.326.0). §3's expression already sends it to `Skip`; **keep it that
way rather than deleting the clause** — a rejected request then sends nothing, silently and
deliberately, and the branch is one line away when the client wants it.

⚠ **THE REJECTED GAP IS REAL WHILE THEY ARE DEFERRED: a rejected requester is told nothing at all.**
They raised a request, the page shows the outcome, and no email arrives. Say this to the client
rather than letting them discover it.

Same shape as `CRS — Audit request activity`'s `EventKind` Compose — ⚠ **but do NOT Save As that
flow. Save As `NotifyApprovers`.** See §10, which was corrected 2026-09-09 after building this for
real: the Audit flow lacks the approver lookup, which is the half that is expensive to get right.

---

## 2. Trigger

**When an item is created or modified**, list `CRS Requests`.

⚠ **BOTH, and neither alone is enough.** A request is CREATED as `Pending` (templates 6/9) and
MODIFIED when decided (7/8/10/11).

⚠ **UNLIKE `NotifyApprovers`, THERE IS NO TWO-PHASE WRITE HERE.** That flow fires twice per upload
because the form uploads and *then* tags, so `HasUnit` has to discard the first firing. `Requests.tsx`
POSTs a complete row in one call, so a creation fires once. Do not copy `HasUnit` across.

---

## 3. `EventKind` — and the three statuses that must SKIP

Compose, named `EventKind`:

```
if(equals(triggerOutputs()?['body/Status'],'Pending'),
   if(equals(triggerOutputs()?['body/RequestType'],'Share'),'ShareRequest','DeleteRequest'),
 if(equals(triggerOutputs()?['body/Status'],'Approved'),
   if(equals(triggerOutputs()?['body/RequestType'],'Share'),'ShareApproved','DeleteApproved'),
 if(equals(triggerOutputs()?['body/Status'],'Rejected'),
   if(equals(triggerOutputs()?['body/RequestType'],'Share'),'ShareRejected','DeleteRejected'),
 'Skip')))
```

⚠⚠ **`RequestStatus` HAS SIX MEMBERS AND ONLY THREE HAVE AN EMAIL.** A chain of `if`s with a
fall-through catch-all is exactly how a value nobody was thinking about gets the wrong branch — the
mistake `fromListItem` made in code, where `Revoked` and `Cancelled` fell through to `Pending` and put
un-clearable rows in every approver's queue (1.0.326.0). The three that must reach `'Skip'`:

- **`Revoked`** — `performRevoke` writes `Status` and `RevokedBy`, so a revoke DOES fire this trigger.
- **`Cancelled`** — the requester withdrew it. Emailing them about their own withdrawal is noise.
- **`Failed`** — the action broke *after* approval. ⚠ **Nobody is told, and that is a real gap** — see §9.

Then a Condition **`Should notify`**: `outputs('EventKind')` **is not equal to** `Skip`. Everything
below sits in its True branch.

⚠ **Type any numeric comparison through the fx editor.** A `0` typed into the value box is TEXT while
`length()` returns an Integer, and that mismatch failed all three conditions in the approver
notification flow on 2026-08-25.

---

## 4. Who gets it

### 4.1 Pending → the unit's approver(s)

Same Group Map join `NotifyApprovers` uses, so the two cannot disagree about who approves a unit:

```
CRS Group Map, filter:  Role eq '<APR or APRHC>' and UnitTermGuid eq '<UnitTermGuid>'
```

⚠ **`UnitTermGuid`, NEVER `Unit`.** The row carries both and the label is display-only — *"terms get
renamed"*. A queue keyed on the label stops matching the day someone tidies the term store.

⚠⚠ **CORRECTED 2026-09-09 WHILE BUILDING: THE `Unit` FALLBACK BELOW IS NOT BUILDABLE, AND THE NAIVE
BLANK CASE IS DANGEROUS.** This paragraph used to read *"fall back to `Unit` when `UnitTermGuid` is
blank — `matchesUnit` in `requests.ts` already does exactly this."* Both halves are wrong here:

- **`CRS Group Map` HAS NO `Unit` COLUMN.** Its columns are `GroupId`, `GroupName`, `Segment`,
  `UnitTermGuid`, `Role`, `Scope`, `Target` (`GroupMapWriteRow` in `groupMapModel.ts`). Filtering
  `Unit eq '…'` names a column that does not exist and **fails the run**. `matchesUnit` compares a
  REQUEST against a viewer's scope in code — it never queries this list, so it was the wrong
  precedent to reach for.
- **⚠ AND `UnitTermGuid eq ''` MATCHES THE WRONG ROWS RATHER THAN NO ROWS.** `Segment` and
  `UnitTermGuid` are populated **only on `Folder`-scope rows**; Site, Library and Page rows carry
  blank ones. So a blank guid on the request would match **every library- and page-scope row**,
  including the Library-scope `APR` row that every unit's approver group holds — emailing approvers
  across the whole site off one empty field.

**So it fails CLOSED.** The filter is built as an expression that cannot match when the guid is
blank:

```
if(empty(coalesce(triggerOutputs()?['body/UnitTermGuid'],'')),
   'Role eq ''__none__''',
   concat('Role eq ''', outputs('ApproverRole'),
          ''' and UnitTermGuid eq ''', triggerOutputs()?['body/UnitTermGuid'], ''''))
```

⚠ **Cost, accepted and worth stating: a request row with a blank `UnitTermGuid` notifies NOBODY**,
silently. That is the correct trade against blasting the site, but it means a legacy row predating
that column gets no approver email at all. If those rows turn out to exist in numbers, the fix is to
resolve the term from `Unit` **before** the Group Map read, not to loosen this filter.

⚠ **DO NOT ADD `and Scope eq 'Folder'`.** It reads as free insurance and is not: a blank `Scope`
normalises to `Folder` *in code* (`normalizeScope`), but an OData filter will not match a blank — so
the clause would silently exclude every row written before that column existed.

⚠ **NEVER derive the group name from the path.** Renaming an abbreviation renames the FOLDER and
leaves the group alone (1.0.162.0), so a derived name stops matching, `Get group members` returns
nothing, and **the flow goes on succeeding while nobody is ever notified.**

### 4.2 ⚠⚠ `APR` vs `APRHC` — decided by the LIBRARY, not by the request

**`APR` is held by the plain `hou` persona and `APRHC` only by `hou_hc`.** Filter an HC document's
request on `APR` and **every ordinary Head of Unit is emailed about a Highly Confidential file** —
they learn its filename, its unit and who asked. That disclosure is the entire reason the HC split
exists.

The row has no library field, so derive it from **`ItemUrl`**:

```
HC  when  contains(toLower(triggerOutputs()?['body/ItemUrl']), '/hcapprovaldocument/')
     or   contains(toLower(triggerOutputs()?['body/ItemUrl']), '/hcdocuments/')
     or   contains(toLower(triggerOutputs()?['body/ItemUrl']), '/hcarchive/')
```

⚠ **URL SEGMENTS, NOT TITLES.** A rename never touches the URL (gotcha #12) and this client renames
libraries routinely — three times in two days in August. Matching on `HC Approval Document` would have
broken on 2026-08-27 and again on 2026-08-28.

⚠ **`ItemUrl` IS OPTIONAL ON THE ROW.** If it is blank the library is UNKNOWN — **filter on `APRHC`,
not `APR`.** Over-notifying HC approvers about an ordinary document is a wasted email; the other way
round is a disclosure. Fail towards the narrower audience.

### 4.3 Approved / Rejected → the requester

`triggerOutputs()?['body/RequestedBy']` — a single address, already stored lower-cased. No lookup.

---

## 5. The empty-approver guard

⚠ **AN EMPTY APPROVER GROUP IS A VALID STATE AND FAILS THE RUN WITHOUT THIS.** `Recipients` joins to
`""` and `Send an email (V2)` returns **"To Field cannot be null or empty"**. Hit on the first live
test of `NotifyApprovers`, against a group recreated that morning with nobody in it.

Condition **`HasRecipients`**: `length(coalesce(body('PickEmails'), createArray()))` **is greater
than** `0` — the `0` entered **through the fx editor**.

⚠ **A SKIPPED EMAIL AND A BROKEN FILTER LOOK IDENTICAL** — both produce 0s and nothing sent. Testing
the empty case proves the guard and **not** the chain. Test a populated group too.

---

## 6. What the templates ask for that the row does NOT have

Read this before writing the bodies — two tokens have no source.

| Token | Where it comes from |
|---|---|
| `[File Name]` | `ItemName` |
| `[File Link]` | `ItemUrl` |
| `[Unit Name]` | `Unit` (the label — right for display) |
| `[Requested Date]` | `RequestedAt`, ⚠ through `convertFromUtc(..., 'Singapore Standard Time')` |
| `[Recipient Email]` | `ShareWith` — multi-value, so `join(...,', ')` |
| `[Approver Reason]` | `DecisionNote` |
| `[Approved By]` / `[Approver Name]` | `DecidedBy` — an **address**; `first(split(...,'@'))` for a name |
| `[PIC Name]` | `RequestedBy` — same, an address |
| `[Folder Path Name]` (template 9 "Location") | derive from `ItemUrl`: everything before the last `/` |
| ⚠ **`[Year]`** (template 6) | **NOT ON THE ROW** |
| ⚠ **`[Document Type]`** (template 6) | **NOT ON THE ROW** |

⚠ **Year and Document Type would need a second read of the DOCUMENT**, by `ItemUniqueId`, and the
library depends on `Stage`: a pending-stage request's file is in the approval library, an approved
one's in `Documents`/`HC Documents`. Doable — the flow runs as the service account and can read
either — but it is an extra action, an extra failure mode, and both fields are decoration in a
template whose point is *which document, whose request, what link*.

**✅ SETTLED 2026-09-09: both lines are DROPPED from template 6.** The email still carries which
document, whose request, which unit and the link — Year and Document Type are decoration in a
template whose job is *come and decide this*. Raised with the client as a recommendation and not
objected to; **it is a visible change to their supplied wording, so confirm it at the next
walkthrough rather than assuming it landed.**

If they want them back: add a `Get item` by `ItemUniqueId` inside the Pending branch only, and
⚠ remember managed metadata comes back as `Label|GUID` from the connector —
`first(split(...?['Value'], '|'))`, the trap that shipped a raw GUID into the reminder email on
2026-08-31. ⚠ **And the library depends on `Stage`** (a pending-stage request's file is in the
approval library, an approved one's in `Documents`/`HC Documents`), so it is two reads or a
conditional one, not a single `Get item`.

---

## 7. ⚠ `[Approve]` / `[Reject]` cannot be one-click

Templates 6 and 9 imply acting from inside the email. **They collapse into ONE link to the Requests
page**, exactly as templates 2 and 3 did:

```
<a href="{site}/SitePages/Request.aspx">Open the request</a>
```

A link that actually approves needs a bearer trigger URL in the mail body, and that was **rejected for
HC** and is worse here — a share request decided from an email would run outside the approver's own
session, and the whole approval design rests on the action executing **as them, with their own
permissions**. The Requests page is where the decision belongs.

⚠ **Do not point it at `ApprovalDocument.aspx`.** That is the document approval queue — a different
screen for a different job.

---

## 8. Dedupe

Reuse the `Already logged` shape from `CRS — Audit request activity`: `Get items` on `CRS Audit Log`,
filtered on `ItemUniqueId` + the event, inside a short window.

⚠ **A FAILED DEDUPE CHECK MUST STILL SEND.** Use `empty()`, not a length test — `empty(null)` is true
where `length(null)` throws. An unreadable list says nothing about whether the email went; **a
duplicate is read past, a missing one stalls a request with nobody looking.**

⚠ **A BLANK `ItemUniqueId` POISONS EVERY LATER DEDUPE.** One row written with a blank id makes
`ItemUniqueId eq ''` match for any later blank, and the flow then answers *already sent* and writes
nothing while reporting success. Guard with a `HasKey` condition before the read — the trap that
silently stopped an audit flow on 2026-08-23.

`Concurrency control` = 1 if duplicates ever appear. ⚠ **It cannot be changed after enabling.**

---

## 9. Known gaps to raise with the client

- **⚠⚠ `Failed` notifies nobody — OBSERVED LIVE 2026-09-10, no longer hypothetical.** An approval
  whose share or delete broke records `Outcome: Failed`, `EventKind` maps it to `Skip`, and the
  requester hears nothing at all. The templates have no seventh case; this needs one, or a line in
  template 8/11.
  - **It was hit on the first real failure**, inside an hour of the flow going live on the test site:
    a pending-stage deletion failed because the document had been approved and routed in the meantime
    (its `UniqueId` no longer resolved — see the CLAUDE.md entry of 2026-09-10). The request read
    `Failed` on the page and **no email was sent to anybody**.
  - **⚠ THE TWO GAPS COMPOUND, and that is the part worth stating to the client.** The failure is
    *permanent* — nothing repoints the id, so re-approving keeps failing — **and** it is *silent*. So
    the requester waits on something that can never complete and is never told. Either gap alone is
    tolerable; together they are a request that disappears.
  - **The build is two edits:** `EventKind`'s final `'Skip'` becomes
    `if(equals(triggerOutputs()?['body/Status'],'Failed'), 'RequestFailed', 'Skip')`, and
    `WhichApproved` gains a `RequestFailed` case emailing `RequestedBy`. ⚠ Keep everything else
    falling through to `Skip` — `Revoked`, `Cancelled` and both `Rejected` cases must stay silent.
- **`Revoked` notifies nobody**, and SharePoint sends no "access removed" mail either — a revoked link
  simply starts refusing. Deliberate today; worth a decision.
- **Template 10 says 90 days. It is 93.** Recorded throughout this project and already corrected in
  the delete dialogs on 2026-09-03 (*"stay 93, they might not know that is why they say 90"*).
- **Template 8 signs off `[PIC Name]` — the person it is addressed to.** Templates 7 and 11 sign off
  `[Approver Name]`. Almost certainly a slip in 8; confirm.
- **Templates 6 and 9 also sign off `[PIC Name]`** while being addressed to the approver. That one
  reads correctly — the PIC is the one asking.

---

## 10. Build order and verification

⚠⚠ **STEPS 1–3 WERE REWRITTEN 2026-09-09 AFTER BUILDING THIS FOR REAL. THE ORIGINAL WAS WRONG IN
THREE WAYS AND COST AN HOUR OF SAVE FAILURES.** What it said, and why each was wrong:

- *"Save As `CRS — Audit request activity`"* — **wrong base.** That flow has the trigger and the
  dedupe; it does **not** have the approver lookup, which is the expensive half. Use
  `NotifyApprovers`, which already carries `GetApproverGroup → HasApprover → GetMembers →
  MembersWithEmail → PickEmails → Recipients → HasRecipients` **verified live**. The dedupe is a
  handful of actions; that chain took a day and produced three separate live defects.
- *"repoint the trigger"* — **necessary and nowhere near sufficient.** See step 2.
- It never mentioned the inherited junk, which is what actually blocks the save.

### The order that works

1. **Save As `NotifyApprovers`**, rename to `CRS — Notify request activity`, and on the trigger set
   **List Name** to `CRS Requests`. Leave `Split on` **On** (one run per request row is correct
   here) and `Concurrency control` **Off** — ⚠ that setting cannot be undone once enabled.

2. ⚠⚠ **DELETE ALL THREE INHERITED TRIGGER CONDITIONS** (Settings → Trigger conditions):
   `{IsFolder}`, `BulkImport`, `{ModerationStatus} = 'Pending'`. **Two of them can never be true on a
   LIST item** — `CRS Requests` has no folders and no content approval, so both read null — so the
   flow would have **fired zero times**. And a flow that never fires leaves **no run history**, so it
   presents as a finished, saved, switched-on flow with nothing at all to inspect. This is the worst
   failure shape in this project's Power Automate history and the original runbook did not mention it.

3. **Delete `GetDoc`.** None of the four emails need it — every value comes off the trigger.
   ⚠ **IT IS REFERENCED BY TWO OTHER ACTIONS AND THE FLOW WILL NOT SAVE UNTIL BOTH ARE FIXED.** The
   error names one action at a time, so it surfaces twice and looks like two unrelated problems:
   - **`MembersWithEmail`** — its filter excludes the *document's uploader* from the approver list
     (so a Head of Unit who uploads their own file is not emailed to approve it). **Translate, do not
     delete**: the equivalent here is excluding whoever raised the request, since `hou` holds both
     `APR` and `UPL`. Swap `body('GetDoc')?['Author']?['Email']` for
     `triggerOutputs()?['body/RequestedBy']`, leaving the `'---'` fallback in place so a blank
     address excludes nobody rather than silencing the email. Leave its `From` untouched.
   - **`HasUnit`** — its `length(...) > 0` row is the old "has this document been tagged yet" check.
     **Repurpose the whole condition** (step 4) rather than deleting it; its True branch already
     holds the approver chain, so deleting it means rebuilding that by hand.

4. **Add the six Composes from §11.2** between the trigger and `HasUnit`, in order — `SiteOrigin`,
   `FileLink`, `RequestsPageLink`, `IsHc`, `ApproverRole`, `EventKind`. ⚠ **Adding the action and
   filling its `Inputs` are two separate steps in the new designer**, and an empty Compose fails
   validation with `'Inputs' is required`. ⚠ Paste 2–6 through the **`fx` expression editor**, not the
   plain field — as plain text they are stored as literal strings and the flow will cheerfully email
   somebody the characters `if(equals(...))`.

5. **Rename `HasUnit` to `NeedsApprover`** and replace its condition with **two rows joined by `Or`**:
   `outputs('EventKind')` is equal to `ShareRequest`, and `outputs('EventKind')` is equal to
   `DeleteRequest`.
   - ⚠ **DELETE the original `length(...) > 0` row** — it is the last `GetDoc` reference.
   - ⚠⚠ **DELETE ANY EMPTY "Choose a value" ROW.** An empty row in an `Or` group compares nothing to
     nothing, which evaluates **true**, and one true row makes the whole condition true — so
     `NeedsApprover` would be true for *everything*: the two approved emails would never fire, and
     rejected and cancelled requests would start hitting the approver lookup. It would look like it
     worked while being wrong in three directions.
   - Renaming is safe: a Condition has no output, so nothing can reference it by name in an
     expression.

6. **Why this shape, and the trap that comes with it.** `NeedsApprover` splits the four emails
   exactly where they differ, and **nothing has to be moved**: its True branch already contains the
   approver chain (→ the two "please approve" emails), and its empty False branch takes the two "your
   request was approved" emails, which go straight to `RequestedBy` with no lookup.
   ⚠ **`Skip` IS THEREFORE HANDLED IMPLICITLY** — a skipped event lands in the False branch, where
   the Switch simply has no matching case. **If anyone ever puts an action in that Switch's `Default`
   branch, every rejected, cancelled, revoked and failed request starts emailing people.**

7. **Repoint `GetApproverGroup`** at `CRS Group Map` with the fail-closed filter in §4.1.
   ⚠ Its inherited filter reads the *approval library's* `UnitTid`, which does not exist on
   `CRS Requests` — so left alone it matches nothing, finds no approver, and **sends no email while
   reporting success.**

8. `HasRecipients` — test the EMPTY group first, then a populated one.
4. The six branches. ⚠ Open each email — **a green `Send an email` says the send worked, never that
   the content is right.** A raw GUID in the body and an escaped `<p>` both report success.
5. ⚠ **Paste the bodies through the `</>` code-view button**, not the rich-text box, or the HTML is
   escaped and the approver receives literal tags.
6. ⚠ **`&` in a link must be `&amp;`** once the body is real HTML.
7. ⚠ **Never press Enter in an fx box.** A trailing newline inside an expression is invisible in the
   designer and has broken this project's flows four times.
8. **The HC test is the one that matters**: raise a share request on an HC document and confirm the
   mail reaches `_APR_HIGHLY_CONFIDENTIAL` and **not** the plain `_APPROVER` group.

⚠ **On SDG's tenant, sign in as the SERVICE ACCOUNT before the first action.** The connection is baked
in at creation, and a flow built as a person stops silently when that password changes.

---

## 11. Build sheet — the four emails, paste-ready

⚠ **SAVE AS `NotifyApprovers`, NOT `CRS — Audit request activity`.** §1 and §10 say the latter because
it carries the trigger shape and the dedupe — but **the expensive, error-prone half of this flow is the
approver lookup** (Group Map join → group members → collect addresses → empty guard), and
`NotifyApprovers` already has that chain **verified live**. The dedupe is a handful of actions; the
approver chain took a day and produced three separate live defects. Start from the harder half.

### 11.1 Trigger and the fields it gives you

**When an item is created or modified** → `CRS Requests`.

Everything below reads from that one trigger. No `Get item` is needed for any of the four emails.

`Id` · `RequestType` · `Status` · `ItemUniqueId` · `ItemName` · `ItemUrl` · `Segment` · `Unit` ·
`UnitTermGuid` · `RequestedBy` · `RequestedAt` · `Reason` · `ShareWith` · `SharePermission` ·
`ExpiresAt` · `DecidedBy` · `DecidedAt` · `DecisionNote` · `Stage` · `RevokedBy`

### 11.2 Composes to create first, in this order

**`SiteOrigin`** — a plain literal, no expression:
```
https://dcidigitalcom.sharepoint.com
```
⚠ **Swap this on SDG's tenant.** It is the one value that differs per site, so it lives in exactly
one place rather than inside four email bodies.

**`FileLink`**
```
if(startsWith(coalesce(triggerOutputs()?['body/ItemUrl'],''), 'http'),
   triggerOutputs()?['body/ItemUrl'],
   concat(outputs('SiteOrigin'), triggerOutputs()?['body/ItemUrl']))
```
⚠ **`ItemUrl` IS SERVER-RELATIVE** — `performShare` proves it, prefixing `window.location.origin`
before use. A body that drops it in raw produces a dead link. The `startsWith` arm is insurance
against a row that ever stores an absolute one, so the concat cannot double up.

**`RequestsPageLink`**
```
concat(outputs('SiteOrigin'), '/sites/ClarenceDMSTesting/SitePages/Request.aspx')
```
⚠ **VERIFY THE FILE NAME PER SITE — this client renames every page at import**, and the same rename
already made a dead link of a hardcoded page address once. It is `Request.aspx` on
ClarenceDMSTesting (confirmed 2026-08-21 from a live AccessDenied); **read it off Site Pages on SDG
rather than assuming.** There is no cheap dynamic resolution inside a flow, so this is a per-site
literal that has to be checked.

**`IsHc`**
```
or(or(contains(toLower(coalesce(triggerOutputs()?['body/ItemUrl'],'')), '/hcapprovaldocument/'),
      contains(toLower(coalesce(triggerOutputs()?['body/ItemUrl'],'')), '/hcdocuments/')),
   contains(toLower(coalesce(triggerOutputs()?['body/ItemUrl'],'')), '/hcarchive/'))
```

**`ApproverRole`**
```
if(empty(coalesce(triggerOutputs()?['body/ItemUrl'],'')), 'APRHC',
   if(outputs('IsHc'), 'APRHC', 'APR'))
```
⚠ **A BLANK `ItemUrl` RESOLVES TO `APRHC`, AND THAT DIRECTION IS DELIBERATE.** Unknown library means
unknown confidentiality; over-notifying an HC approver about an ordinary document wastes an email,
while the other way round emails **every plain Head of Unit** the filename and unit of a Highly
Confidential document. Fail towards the narrower audience — §4.2.

**`EventKind`** — §3's expression verbatim. ⚠ Do not simplify the nested `if`s into a flat lookup:
`Revoked`, `Cancelled`, `Failed` and both `Rejected` cases must land on `Skip`.

### 11.3 The four bodies

Tokens map to these, and nothing else is needed:

| Template token | Expression |
|---|---|
| `[File Name]` / Document Name | `triggerOutputs()?['body/ItemName']` |
| `[File Link]` | `outputs('FileLink')` |
| `[Unit Name]` | `triggerOutputs()?['body/Unit']` |
| `[PIC Name]` | `first(split(coalesce(triggerOutputs()?['body/RequestedBy'],''), '@'))` |
| `[Recipient Email]` | `triggerOutputs()?['body/ShareWith']` |
| `[Approver Name]` / `[Approval Name]` | `first(split(coalesce(triggerOutputs()?['body/DecidedBy'],''), '@'))` |
| `[Approver Reason]` | `triggerOutputs()?['body/DecisionNote']` |
| `[Requested Date]` | see below |
| `[Current Date]` (template 10) | see below |
| `[Folder Path Name]` (template 9 Location) | see below |

**`[Requested Date]`** — ⚠ guarded, because `convertFromUtc` **throws on null** and takes the whole
run with it:
```
if(empty(coalesce(triggerOutputs()?['body/RequestedAt'],'')), '',
   formatDateTime(convertFromUtc(triggerOutputs()?['body/RequestedAt'],
     'Singapore Standard Time'), 'dd MMM yyyy HH:mm'))
```
⚠ **`convertFromUtc` IS NOT OPTIONAL.** SharePoint stores UTC and the site runs UTC+8, so anything
raised before 08:00 local renders on the **previous day** — the same 8-hour fault already recorded
against Auto-route's `Created` stamp and fixed in the reminder flow on 2026-08-31.

**`[Current Date]`** — prefer the moment the decision was actually recorded, not the moment the flow
happened to run:
```
formatDateTime(convertFromUtc(coalesce(triggerOutputs()?['body/DecidedAt'], utcNow()),
  'Singapore Standard Time'), 'dd MMM yyyy')
```

**`[Folder Path Name]`** — everything before the last `/`:
```
join(take(split(triggerOutputs()?['body/ItemUrl'], '/'),
     sub(length(split(triggerOutputs()?['body/ItemUrl'], '/')), 1)), '/')
```

### 11.4 Deviations from the client's supplied wording — all four, stated

1. **`[Approve]` / `[Reject]` collapse into ONE link** to the Requests page, per §7. A link that
   decides from the inbox needs a bearer trigger URL in the mail body, which was rejected for HC and
   is worse here: the whole approval design rests on the action running **as the approver, in their
   own session, with their own permissions**.
2. **Template 6 loses `Year` and `Document Type`** — §6, settled 2026-09-09.
3. **Template 10 says 93 days, not 90.** Client confirmed 93 on 2026-09-03 (*"stay 93, they might not
   know that is why they say 90"*). ⚠ **A client-facing email must never understate a retention
   window somebody may rely on to recover a document.**
4. **Template 6 signs off `[PIC Name]` while addressed to the approver — kept as supplied.** That one
   reads correctly: the PIC is the one asking. ⚠ Template 8 has the same shape and is **wrong** there
   (it is addressed to the PIC and signed by them), which is why it is worth confirming before the
   deferred templates get built.

### 11.5 Verification — the order that actually proves something

1. **Deletion request, populated approver group** → one email, right approver, link opens the
   Requests page.
2. **Approve it** → one email to the requester, 93 days in the body.
3. **Share request** → approver email names the recipient and the unit.
4. **Approve it** → requester email. ⚠ **Microsoft's own invite to the recipient still fires** —
   `sendEmail: true` is unchanged, deliberately — so expect **two** emails to two different people,
   and the recipient's one may be in spam. That is the pre-existing issue, not a fault in this flow.
5. **Reject one, and cancel another** → expect runs that complete and send **nothing**. This is the
   test that proves `Skip` works, and it is the one most likely to be skipped.
6. ⚠ **THE HC TEST IS THE ONE THAT MATTERS.** Raise a share request on an HC document and confirm the
   mail reaches `_APR_HIGHLY_CONFIDENTIAL` and **not** the plain `_APPROVER` group.
7. ⚠ **An empty approver group must complete quietly, not fail.** §5. Test it — but note it proves
   the guard and **not** the chain, so test a populated group too.

⚠ **OPEN AND WATCH: duplicates.** The trigger is created-or-modified and nothing here dedupes. A row
touched again while still `Pending` would re-send the approver email. Not known to happen in the
normal flow — the row is created `Pending` and then decided — but if duplicates appear, add §8's
`Already logged` shape. ⚠ `Concurrency control` = 1 **cannot be undone once enabled.**
