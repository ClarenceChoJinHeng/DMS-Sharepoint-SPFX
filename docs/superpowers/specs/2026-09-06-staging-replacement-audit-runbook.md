# `CRS — Audit replacements` — runbook (2026-09-06)

**Status: NOT BUILT.** Power Automate only; no code change.

Closes the last gap in the replacement trail: a document replaced **while still pending in the
approval library** is recorded nowhere in `CRS Audit Log`.

---

## 1. Why this cannot be done in code

`writeAudit` is refused for every non-Owner — that is the whole reason
`CRS — Audit request activity` had to exist (2026-08-26). The person pressing *Replace* on the
upload form is a PIC or a Head of Unit, i.e. exactly such a non-Owner. So the row must be written by
a flow running as the service account.

---

## 2. Why the trigger is `CRS Submissions`, not the library

A library trigger **cannot tell a replacement from an ordinary upload**:

- `Files/Add(overwrite=true)` keeps the same item, so there is no create event.
- Every upload has `Modified` a few seconds after `Created` anyway — the form uploads, then tags.

The app already records the fact, in the one place that watched it happen:
`Form.tsx:3141` calls `markRecordReplaced(...)`, which sets **`ReplacedAt`** and **`ReplacedBy`** on
the displaced `CRS Submissions` row. That row also carries `FileName`, `ItemPath`, `LibraryTitle`,
`ItemUniqueId` and `SubmissionFileId` — everything an audit row needs.

**So the flow watches the record, not the file.**

---

## 3. ⚠ ONE WRITER — REMOVE THE `Replaced` ROW FROM BOTH ROUTING FLOWS FIRST

`ReplacedAt` is set on **three** occasions today:

| Where | When |
|---|---|
| `Form.tsx` | an uploader replaces a **pending** draft — **logged nowhere** |
| `ApprovalDocument.tsx` | an approver approves over a **filed** document |
| `Auto-route` / `HC Auto Route` (`StampReplacedRecord`) | the routing copy replaces the filed document |

A flow on `CRS Submissions` fires for **all three**, and `Auto-route` already writes its own
`Replaced` audit row for the second and third. Leave both in place and every approval-time
replacement is logged **twice**.

- ⚠ **`LibraryTitle` cannot separate them.** The record's `LibraryTitle` is where the file was
  *uploaded* — the approval library — for the staging case *and* the approval case alike.
- ⚠ **Do not try to filter by actor either.** The replacer and the approver are frequently the same
  person.

**Required first step:** delete the `Replaced` `Create item` action (the one inside the
`WasReplaced` condition) from **`Auto-route`** and from **`HC Auto Route`**.

- **Nothing is lost.** The new row carries the same file name, path, library and `ItemUniqueId`, and
  it covers a case the routing flows never could.
- **`StampReplacedRecord` STAYS in both flows.** It writes `ReplacedAt`, which is what My
  Submissions reads to show `Cancelled` and what this flow now triggers on. Removing it would
  silently break both.
- ⚠ **Take the row out of `HC Auto Route` too.** Every HC clone in this project has shipped with a
  reference nobody swapped; this is the same shape in reverse.

---

## 4. Build

**Save As** of `CRS — Audit request activity` — same trigger shape, same dedupe pattern, same
`Create item` target.

> On SDG's tenant, **sign in as the SERVICE ACCOUNT before the first action.** The connection is
> baked in for life, and a flow built as a person stops silently when that password changes.

### 4.1 Trigger

**When an item is created or modified** → list **`CRS Submissions`**.

Settings → **Trigger Conditions**, one condition:

```
@not(empty(coalesce(triggerOutputs()?['body/ReplacedAt'], '')))
```

- ⚠ **Type it into the box with no line break.** A trailing `\r\n` pasted into a Power Automate
  field has broken four expressions in this project and is invisible in the designer.
- ⚠ `ReplacedAt` stays set for ever, so **any later edit of that row re-fires the flow.** The dedupe
  in 4.2 is what makes that harmless — it is not optional.

### 4.2 `HasKey` — Condition, BEFORE the dedupe read

```
@not(empty(coalesce(triggerOutputs()?['body/ItemUniqueId'], '')))
```

- ⚠ A blank id makes the dedupe filter `ItemUniqueId eq ''`, which **matches every other blank-id
  row** — the poisoned-dedupe trap that made one audit flow stop writing entirely on 2026-08-23.
- Blank ⇒ do nothing. It happens only for records written before the id was stamped.

### 4.3 `Already logged` — Get items

List **`CRS Audit Log`**, Filter Query:

```
ItemUniqueId eq '@{triggerOutputs()?['body/ItemUniqueId']}' and EventType eq 'Replaced'
```

Top Count `1`.

⚠ **A record is replaced at most once**, so one record = one `Replaced` event. A file replaced twice
produces a *second* record with its own id, which logs its own event correctly.

### 4.4 `Not already logged` — Condition

```
length(coalesce(body('Already logged')?['value'], createArray()))   is equal to   0
```

- ⚠ **Enter the `0` through the fx editor.** Typed into the value box it is TEXT, and comparing an
  Integer against text fails the whole condition with a type error. Cost time on three conditions
  already.
- **Use this numeric form, not `if(empty(...), 'write', 'skip')`.** That form works in
  `Audit — approval activity` and **failed** in the deletion flow with the array verifiably `[]`;
  cause never established, invisible whitespace the suspect.
- **A FAILED check must still write.** `coalesce(..., createArray())` gives length `0` when the read
  answered nothing, so an unreadable audit list produces a duplicate row rather than a missing one.
  A duplicate is read past; a missing row is gone.

### 4.5 `Create item` — inside the True branch

List **`CRS Audit Log`**.

| Field | Value |
|---|---|
| `Title` | `Replaced by a newer upload: @{triggerOutputs()?['body/FileName']}` |
| `EventTime` | `@{triggerOutputs()?['body/ReplacedAt']}` |
| `EventType` | `Replaced` |
| `Outcome` | `Success` |
| `ActorEmail` | `@{triggerOutputs()?['body/ReplacedBy']}` |
| `ActorName` | `@{first(split(coalesce(triggerOutputs()?['body/ReplacedBy'], ''), '@'))}` |
| `Source` | `Flow:Replacements` |
| `LibraryName` | `@{triggerOutputs()?['body/LibraryTitle']}` |
| `ItemUniqueId` | `@{triggerOutputs()?['body/ItemUniqueId']}` |
| `ItemName` | `@{triggerOutputs()?['body/FileName']}` |
| `ItemPath` | `@{triggerOutputs()?['body/ItemPath']}` |
| `Details` | `This document was replaced by a newer upload of the same name. Submission @{triggerOutputs()?['body/SubmissionRef']}, file @{triggerOutputs()?['body/SubmissionFileId']}.` |

- ⚠ **`EventTime` is `ReplacedAt`, NEVER `utcNow()`.** The trigger polls, so `utcNow()` records when
  the flow noticed rather than when the replacement happened. Same rule as both deletion flows.
- ⚠ **`ReplacedAt` is already ISO** (`markRecordReplaced` writes `new Date().toISOString()` —
  e.g. `2026-09-06T02:14:00.000Z`), which is what the audit list's DateTime column wants. Do **not**
  run it through `formatDateTime` into `M/d/yyyy` — that locale format belongs to
  `validateUpdateListItem`, and this is a plain `/items` POST, which answers a locale string with
  *"Cannot convert a primitive value to the expected type 'Edm.DateTime'"*.
- **`EventType` is TEXT on this list**, so `Replaced` cannot fail the write — and the value is
  already registered in `auditLog.ts`, so it is filterable in the viewer's Action dropdown from the
  first row.
- ⚠ **Do NOT wrap `Create item` in an `Apply to each`.** Power Automate auto-adds one the moment any
  field is picked from an array-typed output, and it does not remove it when the field is later
  repointed. `Already logged`'s array is normally **empty**, so such a loop runs **zero times**, the
  row is never written, and the run reports **Succeeded**. Open the loop and read its `"foreach":`
  line if one appears — that defect has shipped twice in this project.

---

## 5. Verify

1. **Deployment first.** Confirm `Auto-route` and `HC Auto Route` no longer hold the `Replaced`
   `Create item`, and that `StampReplacedRecord` is still there in both.
2. **Staging case** — as an uploader, upload a document; as a *different* uploader, upload the same
   filename into the same folder and choose **Yes** on the replace dialog. Expect **one** `Replaced`
   row naming the second uploader, and the first uploader's My Submissions row reading `Cancelled`.
3. **Approval case** — upload a name already filed, approve it. Expect exactly **one** `Replaced`
   row, from this flow, not two.
4. **HC case** — the same through the HC pair.
5. **Re-fire** — edit anything on that `CRS Submissions` row. Expect the flow to run and the dedupe
   to skip; **no second row**.

⚠ **An empty result on step 2 is not proof of anything until step 3 has also run.** A wrong trigger
condition and a correctly-quiet flow look identical in the run history.
