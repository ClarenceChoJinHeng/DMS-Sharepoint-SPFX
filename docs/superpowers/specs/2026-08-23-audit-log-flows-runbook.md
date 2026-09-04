# Audit log — the Power Automate flows, step by step

Design: `2026-08-13-audit-log-design.md` §8. That section is the record of WHAT the flows do; this
one is the record of HOW they are built, action by action, because Power Automate config is not in
source control.

**Status (2026-08-23):** **Flows A, B and C and the Auto-route action are ALL BUILT and VERIFIED on
ClarenceDMSTesting.** One document produced `Uploaded` → `Approved` → `Moved to Documents` with the
right actors and `History of this file` resolving across all three; hand-deletions in both the
approval library and `Documents` each produced exactly one `Deleted` row with the deleter named; and
an approval produced **no** spurious `Deleted`, which is what §0.1 exists for. **The four HC clones
(§7) are NOT built.** Nothing is built on SDG's tenant.

> ⚠ **On SDG's tenant, sign in as the SERVICE ACCOUNT before creating any of these.** A flow runs
> under its connection, the connection is created implicitly by the first action, and the builder's
> account is baked in for life. Built as a person, the flow stops **silently** when that password
> changes. On ClarenceDMSTesting this is knowingly unfixed.

---

## 0. Two problems the design did not anticipate

Both were found while writing this runbook. Both produce a log that is actively WRONG rather than
merely incomplete, which is worse — a wrong audit trail is believed.

### 0.1 Auto-route's delete would log every approved document as `Deleted`

Auto-route copies the approved file to `Documents`, stamps it, then **deletes the source** from the
approval library. That delete is load-bearing for security (an approved file left behind is visible
to every PIC in the unit). But it fires the approval-library deletion trigger, so **every approved
document would produce a `Deleted` row** — and an auditor reading the log concludes documents are
being destroyed at the moment of approval.

### 0.2 The upload form fires the create/modify trigger twice

`Form.tsx` uploads the file, then calls `validateUpdateListItem` to tag it. That is a create
followed immediately by a modify, both with moderation status Pending — so a single upload writes
**two `Uploaded` rows**.

### 0.3 One mechanism fixes both: the `Routed` row is written BEFORE the delete, and every writer checks first

- Auto-route's added action writes its `Routed` row carrying the **source file's `UniqueId`**, and
  it sits **before** the delete action. So by the time the deletion trigger fires, the row exists.
- The deletion flow looks for a `Routed` row with that `UniqueId`. Found ⇒ this was a routing
  delete, already logged, **write nothing**. Not found ⇒ a person deleted it, write `Deleted`.
- The activity flow looks for a row with that `UniqueId` **and that `EventType`**. Found ⇒ already
  logged, write nothing.

**The ordering is the whole safety.** If the `Create item` is placed after the delete in Auto-route,
the deletion flow can win the race and log a spurious `Deleted`. Put it before, and check it is
before every time you edit that flow.

**A dedupe check that FAILS must write the row anyway.** An unreadable audit list says nothing about
whether the event was already logged, and a duplicate row is recoverable by a human reading two
identical lines — a missing row is not. Fail open, and this is one of the rare places where "fail
open" and "log more" are the same direction.

---

## 1. Prerequisites, in order

1. **The list must exist.** It is provisioned by the `CRS Audit Log` web part, not by a script —
   create a page, add the web part, open it once. Confirm the list title (it is
   `<Prefix> Audit Log`, so `CRS Audit Log` or `DMS Audit Log` depending on the site's prefix) and
   use that exact title in every action below.
2. **Confirm the fourteen columns exist** with the internal names in §2. The flows write by internal
   name; a display-name match over a different internal name fails exactly like an absent column.
3. **Permissions (design §5.2).** Provisioning deliberately does NOT set them, because breaking
   inheritance has to name the service account and a wrong guess locks the flows out of the list they
   write to. Grant the account the flows run under **Contribute** on this list. On SDG's tenant that
   is the service account.
4. **Publish the page** — a modern page left in Draft denies every read-only user, and every
   permission check still reads correct.

---

## 2. The columns each flow writes

Internal names, from `AUDIT_COLUMNS` in `src/shared/spAuditLog.ts`. Everything except `Title`,
`EventTime`, `EventType` and `Source` is optional; **an absent value is left blank, never omitted as
an event**.

| Internal name | Type | Notes |
|---|---|---|
| `Title` | Text | The summary line. This is what the viewer's feed shows. |
| `EventTime` | DateTime | `utcNow()`. ISO — this is the SharePoint connector, not our JSON-light writer, so gotcha #1's `M/D/YYYY` does NOT apply. |
| `EventType` | Text | One of the values in §3. **Text, never Choice** — a value absent from a Choice column's `Choices` fails the whole write, silently. |
| `Outcome` | Text | `Success` for all flow-written rows. |
| `ActorName` | Text | Display name. |
| `ActorEmail` | Text | Indexed — this is how "everything this person did" is queried. |
| `Source` | Text | `Flow:ApprovalActivity`, `Flow:ApprovalDeletions`, `Flow:DocumentsDeletions`, `Flow:AutoRoute`. |
| `LibraryName` | Text | The library title. |
| `ItemUniqueId` | Text | Indexed. **The join key for "history of this file"** — and what the dedupe checks read. |
| `ItemName` | Text | `FileLeafRef`. |
| `ItemPath` | Note | Full server-relative path. A Note column **because a truncated path cannot be matched back to a folder**. |
| `Segment` | Text | Leave blank in the flows — deriving it means parsing the path, and a wrong segment is worse than none. |
| `UnitPath` | Text | Leave blank, same reason. |
| `Details` | Note | Free lines. |

---

## 3. Flow A — Audit: approval activity

**Trigger:** *When an item is created or modified* (SharePoint), site = the CRS site, list =
**`Approval Document`**.

**Trigger condition** (Settings → Trigger Conditions) — folders are not documents:

```
@equals(triggerOutputs()?['body/{IsFolder}'], false)
```

### A1 — `Compose`, name it `EventKind`

Moderation status decides the event.

```
if(equals(triggerOutputs()?['body/{ModerationStatus}'], 'Approved'), 'Approved', if(or(equals(triggerOutputs()?['body/{ModerationStatus}'], 'Denied'), equals(triggerOutputs()?['body/{ModerationStatus}'], 'Rejected')), 'Rejected', 'Uploaded'))
```

> ⚠ **`{ModerationStatus}` IS A STRING, AND THE REJECTED VALUE IS `"Denied"`.** Verified on site
> 2026-08-23 by reading the raw trigger output. Two separate traps in one field:
>
> 1. **Not the integer.** SharePoint stores `OData__ModerationStatus` as 0/1/2/3 and every REST call
>    in this codebase reads it that way — but the *connector* hands the flow a label. An expression
>    comparing to `0` and `1` never matches, falls through to the `else`, and **logs every approval
>    and every rejection as an upload, with the approver's name on it**. It fails silently: the flow
>    succeeds, a row is written, and it is wrong.
> 2. **Not "Rejected".** The connector says **`Denied`**, the legacy vocabulary. Comparing to
>    `'Rejected'` alone leaves rejections falling through to `Uploaded` — where the A2 dedupe then
>    finds the existing `Uploaded` row and skips, so **nothing is written at all** and the run still
>    reports success. Both spellings are accepted above.
>
> We still write `EventType = "Rejected"`, never `Denied`: that is the value in `EVENT`
> (`shared/auditLog.ts`), the viewer's filter chip, and what `statusToDecision` uses everywhere else.
> Logging the connector's word would give the chip nothing to match.

> Draft is folded into `Uploaded` deliberately. A draft is a file that has been put there and not
> yet submitted; calling it anything else invents a lifecycle stage the rest of the system does not
> have.

### A1b — `Send an HTTP request to SharePoint`, name it `GetUniqueId`

> ⚠ **THE TRIGGER DOES NOT CARRY THE FILE'S GUID.** Verified 2026-08-23: `{Identifier}` is a
> **double-URL-encoded path**, not a UniqueId —
> `ApprovalDocument%252fGHO%252fGF%252f…%252fasd%2b-%2bads.pdf`. Using it as the dedupe key matches
> nothing, so the dedupe never fires and the only symptom is duplicate rows nobody counts.
>
> It also has to be the real GUID rather than the path for a second reason: **every code-side writer
> stores a real `UniqueId`**, so a flow storing paths would break the "history of this file" join
> between flow rows and admin rows — the one feature that makes this list worth reading.

- Site Address: the CRS site · Method: **GET**
- Uri:
  ```
  _api/web/lists/getbytitle('Approval Document')/items(@{triggerOutputs()?['body/ID']})/File?$select=UniqueId
  ```
- Headers: key `Accept`, value `application/json;odata=nometadata`

Returns a bare `{"UniqueId":"<guid>"}`. **If it comes back wrapped in a `d` object the `Accept`
header did not take** — check for a trailing colon in the key — and `body('GetUniqueId')?['UniqueId']`
reads null, writing a blank id with nothing to report it.

### A2 — `Get items`, name it `Already logged`

- List: the audit list.
- **Filter Query:**
  ```
  ItemUniqueId eq '@{body('GetUniqueId')?['UniqueId']}' and EventType eq '@{outputs('EventKind')}'
  ```
- **Top Count: 1** — this only ever asks "does one exist".

### A3 — `Condition`, name it `Not already logged`

- Left — expression:
  ```
  if(empty(body('Already_logged')?['value']), 'write', 'skip')
  ```
- Operator: **is equal to** · Right: `write`

Everything else goes in the **True** branch. Leave **False** empty.

⚠ **`empty()` and a string comparison, not a length test against a boolean.** If the audit-list read
fails, `body()` returns an error object, `?['value']` is null, and `empty(null)` is `true` — so the
row is written anyway (§0.3). `length(null)` would throw instead, losing the event. The string
compare avoids Power Automate's loose boolean-vs-string handling in the Condition control.

⚠ **Set `Configure run after` on this Condition to tick "has failed" as well as "is successful"**,
or a failed `Already logged` stops the flow and the event is lost regardless of the expression.

### A4 — `Create item` (inside **True**)

`Title` is not required on this list, so it sits under **Advanced parameters → Show all** rather than
at the top of the panel. Do not skip it — it is what the viewer's **What** column renders, and a
blank one shows an empty summary beside the admin rows that have one, which reads as the flow being
broken.

| Field | Value |
|---|---|
| Title | `@{outputs('EventKind')}: @{triggerOutputs()?['body/{FilenameWithExtension}']}` |
| EventTime | `@{utcNow()}` |
| EventType | `@{outputs('EventKind')}` |
| Outcome | `Success` |
| ActorName | `@{if(equals(outputs('EventKind'), 'Uploaded'), triggerOutputs()?['body/Author/DisplayName'], triggerOutputs()?['body/Editor/DisplayName'])}` |
| ActorEmail | `@{if(equals(outputs('EventKind'), 'Uploaded'), triggerOutputs()?['body/Author/Email'], triggerOutputs()?['body/Editor/Email'])}` |
| Source | `Flow:ApprovalActivity` |
| LibraryName | `Approval Document` |
| ItemUniqueId | `@{body('GetUniqueId')?['UniqueId']}` |
| ItemName | `@{triggerOutputs()?['body/{FilenameWithExtension}']}` |
| ItemPath | `@{triggerOutputs()?['body/{FullPath}']}` |
| Details | `Moderation status: @{triggerOutputs()?['body/{ModerationStatus}']}` |

> **The actor differs by event and that is the point.** An upload is done by the author; an approval
> or rejection is done by whoever last touched it, which on those events is the approver. Reading
> `Author` for an approval would credit the approval to the uploader — the single most misleading
> row this log could contain.

---

## 4. Flow B — Audit: approval deletions

**Trigger:** *When an item is deleted* (SharePoint), list = **`Approval Document`**.

**Trigger condition** (Settings → Trigger Conditions) — **required**:

```
@equals(triggerOutputs()?['body/IsFolder'], false)
```

> ⚠ **Without it, deleting a FOLDER writes a row claiming a document was destroyed.** Reconciliation,
> subtree migrations and stray-folder cleanup all delete folders in bulk, so one structure change
> buries the real deletions under dozens of false ones.

> ⚠ **THE DELETE TRIGGER RETURNS A COMPLETELY DIFFERENT, MUCH SMALLER SHAPE.** Verified on site
> 2026-08-23 — the whole body is:
>
> ```json
> { "ID": 1411, "Name": "TE - TE - TE - 23-08-26",
>   "FileNameWithExtension": "TE - TE - TE - 23-08-26.pdf",
>   "DeletedByUserName": "Clarence Cho", "TimeDeleted": "2026-08-23T02:34:38Z",
>   "IsFolder": false }
> ```
>
> Three consequences, none of them guessable from the create/modify trigger:
> - **`FileNameWithExtension`, not `{FilenameWithExtension}`** — no braces, lowercase `n` in "Name".
>   The braced token resolves to null, so the filter becomes `ItemName eq ''`, matches nothing, and a
>   row is written with a blank filename. That happened on the first run here.
> - **There is NO path and NO UniqueId.** `ItemPath` and `ItemUniqueId` are left **blank** — the file
>   is gone, so nothing remains to fetch a GUID from. **A deletion is therefore the one event that
>   cannot be threaded into `History of this file`.** State that rather than inventing a key.
> - **There IS an actor — `DeletedByUserName`.** Better than the design assumed. No email, so
>   `ActorEmail` stays blank.

### B0 — `Send an HTTP request to SharePoint`, name it `GetDeletedInfo`

The trigger carries no path, but SharePoint keeps the original location on the recycle-bin entry —
and the deleter's **email**, which the trigger also lacks.

- Method **GET**
- Uri:
  ```
  _api/web/recyclebin?$select=LeafName,DirName,DeletedByEmail,DeletedDate&$filter=LeafName eq '@{triggerOutputs()?['body/FileNameWithExtension']}'&$orderby=DeletedDate desc&$top=1
  ```
- Header: key `Accept`, value `application/json;odata=nometadata`
- On the NEXT action (`Was routed` in B, `Create item` in C): **Run after** → tick **has failed**

> ⚠ **A `Deleted` row with no path cannot say WHICH UNIT'S document was destroyed**, which is most of
> what makes a deletion worth auditing. This recovers it. It does **not** recover the `UniqueId` —
> the recycle bin's `Id` is the bin entry's own — so `History of this file` still cannot include a
> deletion.
>
> ⚠ **THE `Accept` HEADER IS LOAD-BEARING AND ITS ABSENCE FAILED THE WHOLE FLOW.** Without it
> SharePoint answers verbose — `{"d":{"results":[…]}}`, no `value` key — so
> `body('GetDeletedInfo')?['value']` is null and `first(null)` **throws**, failing `Create item` with
> *"One or more fields provided is of type 'Null'"*. Happened here on the first run.
>
> ⚠ **AND THE RUN-AFTER DID NOT SAVE IT, which is the general lesson.** Run-after guards against the
> ACTION failing; here the action went green and the failure moved into the EXPRESSION that read its
> output. **Guarding an action and guarding the value it produces are two different things.** Hence
> `coalesce` in B3 — that is what actually holds when the response shape is wrong rather than the
> request.
>
> Residual, accepted: an apostrophe in a filename breaks the `$filter` (double it with
> `replace(…, '''', '''''')` if it ever bites), and two same-named files deleted seconds apart give
> the most recent one's path. Best effort beats blank.

### B1 — `Get items`, name it `Was routed`

- List: the audit list.
- **Filter Query:**
  ```
  ItemName eq '@{triggerOutputs()?['body/FileNameWithExtension']}' and EventType eq 'Routed' and EventTime ge '@{formatDateTime(addMinutes(utcNow(), -5), 'yyyy-MM-ddTHH:mm:ssZ')}'
  ```
- Top Count: 1.

> ⚠ **Filename plus a five-minute window, and that combination is FORCED.** The obvious key is the
> `UniqueId`, and it does not exist here (above). The next obvious is the path — and `ItemPath` is a
> **Note** column, which SharePoint cannot `$filter` at all. `ItemName` (Text) and `EventTime`
> (DateTime, indexed) are the only filterable pair left. A routing delete lands seconds after its
> `Routed` row, so five minutes is generous.
>
> Residual, accepted: a person deleting a DIFFERENT file of the same name within five minutes of a
> routing of that name would be skipped. The cost is one missing `Deleted` row, never a wrong one.
>
> `formatDateTime(...)` is required — `addMinutes(utcNow(), -5)` alone emits fractional seconds that
> SharePoint rejects.

### B2 — `Condition`, name it `Not a routing delete`

- Left — expression:
  ```
  length(coalesce(body('Was_routed')?['value'], createArray()))
  ```
- Operator: **is equal to** · Right: `0`
- **⋯ → Configure run after** → tick **has failed** as well as **is successful**

> ⚠ **USE THIS NUMERIC FORM, NOT THE STRING COMPARISON IN A3.** Flow A's
> `if(empty(...), 'write', 'skip') = 'write'` works there and **failed here** with `value` verifiably
> `[]` — the condition went False when it should have gone True. The cause was never established;
> Flow A's code view shows a trailing `

` on the left-hand string, so invisible whitespace is the
> likely culprit. Numbers cannot carry it. `coalesce` preserves the fail-open rule: a failed read
> gives null → empty array → length 0 → the row is written anyway.
>
> **Flow A was deliberately left on the string form** — it is verified working, and changing verified
> behaviour on an unproven theory is the wrong risk. **Anyone rebuilding should use the numeric form
> for both.**

### B3 — `Create item` (inside **True**)

| Field | Value |
|---|---|
| Title | `Deleted: @{triggerOutputs()?['body/FileNameWithExtension']}` |
| EventTime | `@{triggerOutputs()?['body/TimeDeleted']}` |
| EventType | `Deleted` |
| Outcome | `Success` |
| ActorName | `@{triggerOutputs()?['body/DeletedByUserName']}` |
| ActorEmail | `@{first(coalesce(body('GetDeletedInfo')?['value'], createArray()))?['DeletedByEmail']}` |
| Source | `Flow:ApprovalDeletions` |
| LibraryName | `Approval Document` |
| ItemUniqueId | *(blank — see above)* |
| ItemName | `@{triggerOutputs()?['body/FileNameWithExtension']}` |
| ItemPath | `@{first(coalesce(body('GetDeletedInfo')?['value'], createArray()))?['DirName']}` |
| Details | `Deleted from the approval library — not a routing delete.` |

> **`EventTime` is `TimeDeleted`, not `utcNow()`.** The trigger POLLS, so `utcNow()` records when the
> flow noticed rather than when the deletion happened — a gap of up to a minute, on the one event
> where the timestamp is the point.

---

## 5. Flow C — Audit: Documents deletions

Trigger *When an item is deleted* on **`Documents`**, the same `IsFolder` trigger condition, and a
single `Create item` — **no `Was routed` check and no Condition at all**, because nothing routes files
*out* of `Documents` today, so every deletion there is real.

It DOES keep `GetDeletedInfo` (B0) — the path and deleter email matter just as much here.

Changes from B: `LibraryName` = `Documents`, `Source` = `Flow:DocumentsDeletions`, `Details` =
`Deleted from Documents.`

⚠ **This changes the day the seven-year archive mover goes live.** `MoveTo` removes the item from
`Documents`, which fires this trigger — so **every archived document would log as `Deleted`**. When
the mover is built it must write its row **before** the move, and this flow must gain the same
check-first branch as Flow B. Recorded here because the mover runbook is a separate file and this is
the flow that breaks.

⚠ **Neither deletion flow dedupes against ITSELF.** Flow A skips a repeat of the same
`UniqueId` + `EventType`; B and C cannot, having no `UniqueId`. If the delete trigger were ever to
re-deliver, you would get two identical `Deleted` rows. Not observed here — two rows for one filename
on 2026-08-23 turned out to be two real deletions — so no guard was added. If duplicates do appear,
extend B1's filter to also match an existing `Deleted` row in the window.

---

## 6. Auto-route — the added action

Open the existing **Auto-route** flow. Add ONE `Create item` action, and **place it before the
delete-source action** (§0.3 — the ordering is the whole safety).

| Field | Value |
|---|---|
| Title | `Moved to Documents: @{triggerOutputs()?['body/{FilenameWithExtension}']}` |
| EventTime | `@{utcNow()}` |
| EventType | `Routed` |
| Outcome | `Success` |
| ActorName / ActorEmail | the approver — the same Editor values Auto-route already reads |
| Source | `Flow:AutoRoute` |
| LibraryName | `Documents` |
| ItemUniqueId | **the SOURCE file's UniqueId**, the same value Flow B will look for |
| ItemName | `@{triggerOutputs()?['body/{FilenameWithExtension}']}` |
| ItemPath | the DESTINATION path — where the document now lives |
| Details | `From: <source path>` then `To: <destination path>` |

> ⚠ **`ItemUniqueId` here is the SOURCE file's, while `ItemPath` is the DESTINATION.** That looks
> inconsistent and is deliberate: the id is the dedupe key Flow B matches on, and the path is where
> a reader needs to go to find the document. **A copy produces a NEW `UniqueId`** — Auto-route
> copies rather than moves — so the id of the file in `Documents` is not the id of the file that was
> approved. This is the known break in the audit trail at the routing boundary
> (see the archive spec: a cross-library **move** preserves `UniqueId`, a copy does not).

**Do not touch the trigger condition** while in here. Auto-route must keep
`@equals(triggerOutputs()?['body/{IsFolder}'], false)`, and the polarity is the opposite of the
folder-approval flow's.

---

## 7. HC — the clones

**STATUS, 2026-08-24: ALL FOUR BUILT. HC HAS PARITY WITH THE NORMAL LIBRARIES.**

| Clone | State |
|---|---|
| Auto-route action (in `HC Auto Route`) | ✅ built + **verified** — `Routed` row with a real `UniqueId` |
| Flow A → `Audit — HC approval activity` | ✅ built + **verified** — `Uploaded` and `Approved`, right actors |
| Flow B → `Audit — HC approval deletions` | ✅ built, Flow checker clean — **not yet fired** |
| Flow C → `Audit — HC Documents deletions` | ✅ built — **not yet fired** |

⚠ **THE TWO DELETION CLONES HAVE NEVER RUN.** A flow that never fires leaves no run history, so an
empty history is not evidence they work — it is equally evidence the trigger can never be true. Test
by deleting a pending HC file by hand (⇒ `Flow:ApprovalDeletionsHC`) and an approved one from
`HC Documents` (⇒ `Flow:DocumentsDeletionsHC`).

**VERIFIED HC LIFECYCLE, 2026-08-24 00:04–00:11** — one document, three rows, right actor on each:
`Uploaded` → `Approved` (both `Flow:ApprovalActivityHC`) → `Moved to Documents`
(`Flow:AutoRouteHC`), with **no spurious `Deleted`** despite Auto-route deleting the source. That
absence is what proves the `Create item`-before-delete ordering.

⚠ **`Details` ON A CLONE IS PROSE AND SAYS THE WRONG LIBRARY.** Flow C's read
*"Deleted from Documents."* on the HC clone — contradicting the `Library` column beside it, on the one
row someone reads when asking which library lost a document. Every free-text field in a clone needs
the same pass as every list reference.

⚠ **`Save As` on the original is the whole method**, and then exactly the references below change.
For Flow A that was **three** edits: the trigger's `List Name`, the `getbytitle('…')` inside
`GetUniqueId`, and `Source` + `LibraryName` on `Create item`. **The audit list GUID
(`84ed065f-…`) in `Already logged` and `Create item` must NOT change** — that is where the rows go.

⚠ **THE `fx` BUTTON DESTROYS A MIXED URI, and it looks like an edit rather than a break.** Clicking
it converts the whole Uri box into ONE expression, so the literal `getbytitle('…')` text and the
`ID` token can no longer coexist; the action goes to **Invalid parameters** and the `Method` box
empties. Recovery is to clear the box and retype it as PLAIN TEXT with the token inserted mid-string.
Cost two round trips on 2026-08-23.

⚠ **A TRAILING NEWLINE IN `ItemUniqueId` BREAKS "History of this file", SILENTLY — found
2026-08-24.** Every filter on that column is an EXACT match (`ItemUniqueId eq '…'`), so a value with
a `
` on the end never joins the rows that have the same GUID without one. The symptom is a history
view that reads *"No events match these filters"* over a document whose upload, approval and routing
are all sitting in the list. Two `Routed` rows were written this way; the response shows it as
`<d:ItemUniqueId xml:space="preserve">` with the newline before the closing tag — **that `xml:space`
attribute IS the tell**, and it is invisible in a list view.

The newline arrives by **pasting an expression out of a document into a Power Automate field** — the
same way the `GetSourceUniqueId` Uri picked one up an hour earlier. **After pasting into any flow
field, press End and check for a stray line break**, and prefer reading the value back with
`?$select=ItemUniqueId` over trusting the designer.

Existing rows are repaired by editing the list item and retyping the GUID; nothing else joins them.

⚠ **A BLANK `ItemUniqueId` ROW IS A LANDMINE FOR EVERY LATER DEDUPE.** `Already logged` filters
`ItemUniqueId eq '…'`, so one row written with a blank id (this happened once, while
`GetSourceUniqueId` was misconfigured) makes the filter MATCH for any later event whose id also fails
to resolve — the condition then answers *already logged* and **writes nothing at all**, on a run that
reports success. Delete such rows from `CRS Audit Log` as soon as they are noticed.

Once A, B and C work, clone them:

| Clone of | Trigger list | `LibraryName` | `Source` |
|---|---|---|---|
| Flow A | `HC Approval Document` | `HC Approval Document` | `Flow:ApprovalActivityHC` |
| Flow B | `HC Approval Document` | `HC Approval Document` | `Flow:ApprovalDeletionsHC` |
| Flow C | `HC Documents` | `HC Documents` | `Flow:DocumentsDeletionsHC` |
| Auto-route action | — | added to **HC Auto Route**, `LibraryName` = `HC Documents` | `Flow:AutoRouteHC` |

⚠ **A clone inherits every library reference from its original, and the HC Auto-route clone of
2026-08-19 found SIX wrong ones.** Two of those would not have failed loudly: item ids are per-LIST,
so a stamp pointed at the wrong list writes onto a different document, and a delete pointed at the
wrong list removes whatever holds that id there. **Check every list reference in a clone before
saving it**, including the audit list itself.

---

## 8. Troubleshooting, from the flows already built here

- **`{Identifier}` is a double-encoded PATH, not the file GUID** — confirmed on this tenant
  2026-08-23. Hence `GetUniqueId` in A1b. **Read the raw trigger output before building anything on
  a trigger field**: two of this flow's three design assumptions were wrong, and one read killed
  both.
- **`{ModerationStatus}` is a STRING and rejection reads `"Denied"`** — see A1. The integer form
  used by every REST call in this codebase does not apply to the connector.
- **Header keys must not include the colon.** `Accept`, never `Accept:`. With the colon the header
  silently does not exist and the symptoms look unrelated to one another. Cost two hours on
  2026-08-08.
- **`Run after` is configured on the DOWNSTREAM action**, not on the one that might fail. Getting
  this backwards is how an action ends up never running at all.
- **Dragging actions out of a loop does NOT rewrite `items('…')` references** — the save fails with
  *"The repetition action(s) … are not defined in the template"*. Fix the expressions by hand.
- **A flow that never fires leaves no run history.** An empty run list is not evidence the flow is
  fine; it is equally evidence the trigger condition can never be true.
- **`Hidden: true` on a field removes it from the trigger payload** and nothing reports it — the
  write still succeeds and the trigger reads null. Do not hide a column any flow reads.

---

## 9. Test plan

Run each in order and confirm exactly ONE row appears, with the right actor:

1. **Upload a file through the form.** ⇒ one `Uploaded`, actor = the uploader. **Two rows here means
   the A2 dedupe query is not matching** — check `{Identifier}`.
2. **Reject it.** ⇒ one `Rejected`, actor = the approver.
3. **Approve it.** ⇒ one `Approved` (actor = approver), then one `Routed` (source
   `Flow:AutoRoute`), and **no `Deleted` row** — that absence is the whole point of §0.1. If a
   `Deleted` appears, the `Create item` is after the delete in Auto-route.
4. **Delete a file from `Documents` by hand.** ⇒ one `Deleted`, source `Flow:DocumentsDeletions`.
5. **Delete a pending file from the approval library by hand.** ⇒ one `Deleted`, source
   `Flow:ApprovalDeletions`.
6. **Open the CRS Audit Log page** and confirm the feed reads correctly, the filters narrow
   server-side, and "history of this file" follows one document across upload → approve → route.

⚠ **Step 3 is the one that matters.** It is the only step that proves the two problems in §0 are
actually solved, and it is the sequence that happens hundreds of times a week in production.
