# Auto-route: stamp the DISPLACED record `Cancelled` when a replace fires

**Date:** 2026-09-02
**Status:** BUILT AND VERIFIED on `Auto-route` (ClarenceDMSTesting). **`HC Auto Route` clone still to
do.**
**Flows to change:** `Auto-route approved Pending files to Documents Library`, and its clone
`HC Auto Route`.

---

## VERIFIED 2026-09-02 — and a second gap found and fixed along the way

Built exactly as specced below (§2, §3), then tested end to end with a clean pair of documents
(`test1 - test1 - test1 - 02-09-26.pdf`, uploaded twice, second one via "Send for approval as a
replacement"). Result: the displaced record now reads **Cancelled** — *"Replaced by a newer upload
from clarencechojinheng@gmail.com on 2026-09-01. The details below are what this submission
carried."* — and the replacer's own record resolves as live/Approved. Both halves working.

**⚠ A SECOND, LARGER GAP WAS FOUND WHILE BUILDING THIS, AND HAD TO BE FIXED FIRST.** The first test
(a different document pair) showed the FIRST record still reading "Approved" (not Cancelled) even
after §2/§3 were built and ran — because `Update file` only overwrites file BYTES, never any list
column. On the ordinary `Copy file` path, SharePoint's copy automatically carries over EVERY column
(including `SubmissionId`/`BatchId`/`SubmissionFileId`) for free; there is no equivalent on the
`Update file` path, and this flow's only explicit metadata stamp is Author/Editor/Created (found by
inspecting `Get source author` and both `Send an HTTP request to SharePoint` actions after
`WasReplaced` — neither one stamps business or tracking columns). So after a replace, the destination
item kept the FIRST upload's `SubmissionFileId` forever, which meant:
- The **replacer's own** record could never resolve (nothing on the live item ever matched its stamp)
  — it read `Deleted` permanently, the exact false message this feature exists to prevent, now landing
  on the wrong person.
- The `Cancelled` stamp on the displaced record was invisible, because `mergeRecords`'s precedence
  rule (*"a record that still resolves wins as live regardless of `replacedAt`"*) meant the OLD
  record's still-resolving `SubmissionFileId` kept it reading `Approved`.

**Fixed with two more actions**, `GetSourceSubmissionIds` (reads `SubmissionId`/`BatchId`/
`SubmissionFileId` off the SOURCE item before it's deleted) and `StampDestSubmissionIds` (MERGEs
them onto `DestItemId`) — placed between the existing Author/Editor/Created stamp and the
source-delete step, running **unconditionally** on both branches (harmless no-op on the ordinary
`Copy file` path, since those columns already transferred there automatically). See the note in §4
below — this pair is now part of what must also be cloned into `HC Auto Route`.

**⚠ THE WIDER METADATA GAP (Document Type, Year, Confidentiality, tier columns, Remark,
LegallyPrivileged, Vendor, etc.) IS STILL OPEN, DELIBERATELY NOT FIXED HERE.** Client's own scope
decision (2026-09-02): fix My Submissions only, for now. Every one of those fields ALSO stays as the
first upload's values after a replace — only the three submission-tracking columns were brought into
scope today. If this needs fixing later, the pattern is the same (extend `GetSourceSubmissionIds` /
`StampDestSubmissionIds` — or a renamed, wider version of them — to cover the full field list), but it
was explicitly deferred.

**Building notes worth keeping for the HC clone:** every new action had to be built via the
**Parameters tab**, never Code view — pasting a full JSON block into Code view repeatedly landed inside
the Uri box instead of replacing the action (Code view appears to be read-only for these action types
in this designer). Build each field individually: Method dropdown, Uri box (cleared and typed), Headers
(added as key/value rows), Body (added via "Add new parameter" → Body). Watch for trailing newlines
picked up from copy-pasting values — hit multiple times this session, always traced by re-opening the
action's Code view (read-only, for inspection only) and looking for a stray `\n` at the end of a
string.

---
**Depends on:** `2026-08-31-versioned-replace-runbook.md` (the `GetDestMatch` / `DestExist` /
`UpdateFile` mechanism, and the `WasReplaced` Condition + `Replaced` audit row already built on top of
it, per CLAUDE.md's 2026-09-02 section). This runbook only ADDS to what is already there.

---

## 0. Why — the finding this exists to fix

Verified live 2026-09-02: a document replaced through the versioned-replace mechanism correctly logs
`Replaced` in `CRS Audit Log` — but the DISPLACED file's record in `CRS Submissions` still reads
**Deleted** on My Submissions, not **Cancelled**.

`markRecordReplaced` (`shared/spSubmissionRecords.ts`) is the only thing that ever stamps
`ReplacedAt`/`ReplacedBy` on a record, and it is called from exactly two places, both client-side, both
gated on the APP's own client-side clash check succeeding at the moment a person presses a button:

- `Form.tsx` — the upload form's pending-file overwrite.
- `ApprovalDocument.tsx` — the approval page's `documentsFileClash` check, and only on a **confirmed**
  clash at the moment Approve is pressed.

The flow-level replace (`GetDestMatch` → `DestExist` → `UpdateFile`) runs **later**, server-side, when
Auto-route actually routes the file — and it can fire in cases the client-side check never sees or
never gets a chance to run: the native Approve/Reject command (bypasses the app entirely), the Bulk
Approve panel (currently has no replace path at all), or simply a collision that appears in the gap
between the approval click and the flow actually processing it. In every one of those cases nothing
ever calls `markRecordReplaced`, so the displaced record has no way to know it was superseded and falls
through to the default **Deleted** state — the exact false "your work was destroyed" message this whole
feature exists to prevent.

**The fix: since the flow is now the thing that actually KNOWS a replace happened (it is what decided
`DestExist`), make the flow do the stamping itself**, the same way `RevokedBy` is written by app code
but a flow (`CRS — Audit request activity`) is responsible for reading it into the audit trail. Here
it's the mirror: the flow is the one place that reliably observes every replace, code-side or not, so
it should own writing `ReplacedAt`/`ReplacedBy` directly onto `CRS Submissions` — not just onto the
audit log.

---

## 1. What already exists (context, not to be rebuilt)

Inside the `DestExist` Condition's True branch: `GetSourceContent` → `UpdateFile`.

Later in the flow, after `Get_ApprovedBy_HTTP` and `GetSourceUniqueId`, a `WasReplaced` Condition
re-tests the same boolean (`length(coalesce(body('GetDestMatch')?['value'], createArray())) is greater
than 0`) and, in its True branch, writes the `Replaced` row to `CRS Audit Log` — actor from
`Get_ApprovedBy_HTTP`, falling back to `Editor`, `ItemUniqueId` reused from `GetSourceUniqueId`.

This runbook adds two things around that existing shape.

---

## 2. `GetDestSubmissionFileId` — capture the OLD stamp before anything overwrites it

**Placed inside `DestExist`'s True branch, immediately after `GetDestMatch`** — before
`GetSourceContent`/`UpdateFile`, and well before the metadata stamp near the end of the flow tail,
which is what eventually overwrites the destination item's `SubmissionFileId` to the REPLACING file's
own stamp. If this read happens too late, it reads the wrong (new) value and the wrong record gets
marked.

`GetDestMatch` ("Get files, properties only") does not return list-column values — only file
properties — so this needs a separate read of the matched item's fields, the same shape
`ApprovalDocument.tsx` already uses client-side for the same purpose.

**New action: `GetDestSubmissionFileId`** — *Send an HTTP request to SharePoint*

| Field | Value |
|---|---|
| Site Address | the site |
| Method | GET |
| Uri | `_api/web/lists/getbytitle('<approved-side library title>')/items(@{first(body('GetDestMatch')?['value'])?['Id']})?$select=SubmissionFileId` |
| Headers | `Accept: application/json;odata=nometadata` |

`<approved-side library title>` is the same one `GetDestMatch` itself already targets — `Restricted &
Confidential Document` on `Auto-route`, `Highly Confidential Document` on `HC Auto Route`. Reuse the
exact library-name value already sitting in `GetDestMatch`'s own Library Name field; do not retype it.

> ⚠ **This must run only when `DestExist` is True.** Placing it there, rather than unconditionally
> earlier, means a failed or empty read here can only ever affect the rare replace path, never an
> ordinary approval.

> ⚠ **A failed or malformed read must not stop routing.** Treat a missing `SubmissionFileId` in the
> response the same way the code does — nothing downstream should hard-fail on it; the worst case is
> simply that the later stamp step (§3) finds no record to mark, which is exactly what happens today.

---

## 3. Stamp `CRS Submissions`, alongside the existing `Replaced` audit row

**Placed in the SAME `WasReplaced` True branch that already writes the `Replaced` audit row** — same
reason the audit row itself couldn't sit inside `DestExist`'s own branch: the actor value comes from
`Get_ApprovedBy_HTTP`, which only exists downstream of the whole `DestExist` block closing. Keeping
both writes in the one place also means there is only ONE definition of "the replace happened and here
is who did it" for this flow to maintain.

### 3a. `GetReplacedRecordId` — find the row to update

*Send an HTTP request to SharePoint*, GET:

```
_api/web/lists/getbytitle('CRS Submissions')/items?$select=Id&$top=5
  &$filter=SubmissionFileId eq '@{replace(body('GetDestSubmissionFileId')?['SubmissionFileId'], '''', '''''')}'
```

Headers: `Accept: application/json;odata=nometadata`

- **`$top=5`, not 1** — mirrors `markRecordReplaced`'s own reasoning: one stamp should identify one
  row, but if it somehow named two, silently updating only the first would leave the other reading
  `Deleted` for ever. Cheap to be certain.
- **The `replace(...,"'","''")` is defensive** — `SubmissionFileId` values are always
  `SFI-<yyyymmdd>-<4>` with no apostrophe, but it costs nothing and matches the escaping rule this
  project applies everywhere else a stored string goes into an OData filter.
- **A blank `SubmissionFileId` from step 2 must not blow up this filter.** An empty `$filter` value
  (`SubmissionFileId eq ''`) just returns zero rows — the safe, do-nothing outcome — so no extra
  guarding is needed here specifically, but do not "helpfully" skip the call on blank; let the filter
  answer naturally.

### 3b. `Apply to each` over `GetReplacedRecordId`'s rows

Input: `body('GetReplacedRecordId')?['value']`

Inside the loop — **`StampReplacedRecord`**, *Send an HTTP request to SharePoint*, one MERGE per row
(ordinarily zero or one row; the loop exists only for the two-row edge case above):

| Field | Value |
|---|---|
| Method | POST |
| Uri | `_api/web/lists/getbytitle('CRS Submissions')/items(@{items('Apply_to_each_<n>')?['Id']})` |
| Headers | `Accept: application/json;odata=nometadata`<br>`Content-Type: application/json;odata=nometadata`<br>`X-HTTP-Method: MERGE`<br>`IF-MATCH: *`<br>`odata-version:` *(leave the value blank — see note)* |
| Body | `{ "ReplacedAt": "@{utcNow()}", "ReplacedBy": "<the SAME actor expression already used for the Replaced audit row's ActorEmail>" }` |

> ⚠ **The `odata-version` header must be present with an EMPTY value, not omitted.** SPFx's
> `SPHttpClient` injects `odata-version: 4.0` on browser calls and that is what forced this pattern in
> `WRITE_HEADERS` — a flow's HTTP action does not carry that same default, but keep the header anyway so
> this write behaves identically to every other JSON-light write in this project and is not the one
> exception someone has to remember. If Power Automate's action does not allow an empty header value,
> omit it instead; do not send `4.0`.

> ⚠ **Reuse the exact `ActorEmail`/`ActorName` expression the `Replaced` audit row already uses for
> `ReplacedBy` — do not re-derive it.** Two independent expressions computing "who did the replace" is
> how the audit row and the submission record end up naming different people after the next edit to
> either one.

> ⚠ **`ReplacedAt` is ISO (`utcNow()`), never `M/D/YYYY`.** This is a plain items MERGE through the
> OData layer, not `validateUpdateListItem` — a locale-formatted date here fails with *"Cannot convert a
> primitive value to the expected type 'Edm.DateTime'"*, naming the type but not the field.

> ⚠ **A failed MERGE here must not fail the flow.** Set `Configure run after` on whatever follows this
> loop to `is successful` **and** `has failed` **and** `is skipped`, the same rule already applied to
> every audit write in this flow — the record staying `Deleted` instead of `Cancelled` is a worse
> outcome than it could have been, but it must never be the reason a document fails to route.

---

## 3c. `GetSourceSubmissionIds` + `StampDestSubmissionIds` — the metadata-transfer fix (added after
testing found §2/§3 alone were not enough — see the VERIFIED note at the top)

Placed **between the existing Author/Editor/Created stamp** (`Send an HTTP request to SharePoint`,
the one whose body sets `Author`/`Editor`/`Created` via `body('Get_item')`) **and the source-delete
step** (`Send an HTTP request to SharePoint 1`, the `X-HTTP-Method: DELETE` one). Runs
**unconditionally on both branches** — harmless on the ordinary `Copy file` path, since SharePoint's
copy already carried these columns over there; load-bearing on the `Update file` path, where nothing
else ever would.

**`GetSourceSubmissionIds`** — *Send an HTTP request to SharePoint*, GET, reading the SOURCE item
(same approval-library list, same `triggerOutputs()?['body/ID']` the delete step and `Get source
author` already use) for the three tracking columns:

```
_api/web/lists(guid'<approval-library list guid>')/items(@{triggerOutputs()?['body/ID']})?$select=SubmissionId,BatchId,SubmissionFileId
```
Headers: `Accept: application/json;odata=nometadata`. `runAfter` the Author/Editor/Created stamp,
succeeded.

**`StampDestSubmissionIds`** — *Send an HTTP request to SharePoint*, POST/MERGE onto `DestItemId`:

```
_api/web/lists(guid'<destination-library list guid>')/items(@{outputs('DestItemId')})
```
Headers: `Accept`, `Content-Type` (both `application/json;odata=nometadata`), `X-HTTP-Method: MERGE`,
`IF-MATCH: *`. Body:
```json
{"SubmissionId": "@{body('GetSourceSubmissionIds')?['SubmissionId']}", "BatchId": "@{body('GetSourceSubmissionIds')?['BatchId']}", "SubmissionFileId": "@{body('GetSourceSubmissionIds')?['SubmissionFileId']}"}
```
`runAfter` `GetSourceSubmissionIds`, succeeded.

> ⚠ **Both actions had to be built via the Parameters tab, field by field — Code view repeatedly
> accepted a pasted JSON block into the Uri box instead of replacing the whole action.** Whatever the
> cause in this designer, don't rely on Code view for creating these; use it only to inspect the
> result afterward.

---

## 4. The HC clone

`HC Auto Route` needs all **four** of these additions now, not two: §2's `GetDestSubmissionFileId`,
§3's `GetReplacedRecordId`/`Apply to each`/`StampReplacedRecord`, and §3c's
`GetSourceSubmissionIds`/`StampDestSubmissionIds`. **`CRS Submissions` is ONE list, shared by both
verticals** — `readSubmissionRecords`/`writeSubmissionRecord` are not library-scoped — so §3 and §3c's
writes onto `CRS Submissions` are copied **verbatim**, no library swap needed there.

**Library names to swap for the HC clone:**
- `GetDestSubmissionFileId` (§2) — Uri must point at `Highly Confidential Document`, matching
  `GetDestMatch`'s own Library Name on the HC clone.
- `GetSourceSubmissionIds` (§3c) — its list guid is the **HC approval library's**, not the normal
  one; check against whatever `HC Auto Route`'s own delete step / `Get source author` already use for
  the source item.
- `StampDestSubmissionIds` (§3c) — its list guid is the **HC destination library's** (`Highly
  Confidential Document`), matching the HC `Copy file`/`Update file`'s own destination.

Same single-field trap this project's other HC clones have gotten wrong before; check each explicitly
rather than assuming the normal flow's guids carry over.

---

## 5. Testing

1. **Ordinary approval, no clash.** `DestExist` False ⇒ neither new action runs. Confirm nothing
   changed about a normal approval's audit trail or its My Submissions record.
2. **Replacement, via the app's own confirmed-clash path** (upload a name clash, approve, confirm
   replace). The displaced record should ALREADY read `Cancelled` from the client-side
   `markRecordReplaced` call — this test is to confirm the new flow-level stamp does not fight it or
   write a second, conflicting value. `ReplacedAt` may move slightly later (the flow's write happens
   after the client-side one); that is fine — the record only needs to be marked once, and last-write
   wins on the same two columns either way.
3. **Replacement that the app never saw** — the actual gap this fixes. Reproduce the scenario from the
   screenshot: get a document into the destination library that the client-side check did not catch at
   approval time (e.g. approve through the library's native Approve/Reject command instead of the app's
   own page), so that only the flow's `DestExist` branch detects the collision. Confirm:
   - `CRS Audit Log` shows the `Replaced` row (already working).
   - The displaced file's record on **My Submissions** now reads **Cancelled**, not Deleted, naming who
     replaced it and when.
4. Repeat 1–3 on the HC vertical.

> ⚠ **A green run proves nothing on its own** — the whole reason this gap existed for a day is that
> `Replaced` logging to the audit list worked perfectly while the submission record silently stayed
> wrong. Open My Submissions, not just the audit log.

---

## 6. Rollback

Delete `GetDestSubmissionFileId`, `GetReplacedRecordId`, the `Apply to each` and `StampReplacedRecord`.
Behaviour returns to today's: `Replaced` still logs correctly to the audit list; the displaced record
reads `Deleted` on My Submissions.

---

## 7. On SDG's tenant

**Sign in as the SERVICE ACCOUNT before adding the first of these actions.** Editing an existing action
does not re-create the flow's connection; adding a new one does, and a flow built as a person stops
silently when that password changes.
