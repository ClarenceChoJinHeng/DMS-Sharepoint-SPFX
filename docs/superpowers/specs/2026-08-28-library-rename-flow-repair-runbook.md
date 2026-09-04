# Repairing the Power Automate flows after the 2026-08-28 library rename

**Status: IN PROGRESS.** Flows are not in source control, so this file is the record.

The client renamed all six document libraries on 2026-08-28 (the third rename in two days). The SPFx
package was fixed in `1.0.313.0`. This is the flow half.

---

## 1. The six libraries — title, URL, GUID

Read live from `/_api/web/lists?$select=Title,Id&$expand=RootFolder&$filter=BaseTemplate eq 101`
on `/sites/ClarenceDMSTesting`, 2026-08-28.

| Logical key | Live title | URL segment | **List GUID** |
|---|---|---|---|
| `Staging`     | `Approval for Document` | `/ApprovalDocument` | `a9342528-66be-458c-ba70-a6e3248a5133` |
| `StagingHC`   | `Approval for Highly Confidential Document` | `/HCApprovalDocument` | `2311cd83-90aa-4077-a647-65d251a5f426` |
| `Documents`   | `Restricted & Confidential Document` | `/Shared Documents` | `e322e3a5-3687-4da3-94f8-e0b06c01dd7e` |
| `DocumentsHC` | `Highly Confidential Document` | `/HCDocuments` | `e4fb3b2c-6f20-4bbb-a4f2-e9a5301d4080` |
| `Archive`     | `Archive Restricted & Confidential Document` | `/Archive` | `8c28ca71-4edb-445b-8ebd-5711dacd71d4` |
| `ArchiveHC`   | `Archive Highly Confidential Document` | `/HCArchive` | `bfcd6735-7605-47e7-a0d6-bbfff41bab7b` |

⚠ **`CRS Audit Log` is `84ed065f-…` and MUST NOT be touched.** It was not renamed, and it is where
every audit row goes.

---

## 2. What actually broke — and what did NOT

**A rename changes the TITLE and never the URL** (gotcha #12). So:

| Reference style | Affected? |
|---|---|
| `Copy file`, `Create file`, `Move file` — **paths** | **NO.** URLs are unchanged. |
| `_api/web/lists(guid'…')` | **NO.** |
| Auto-route's `Compose_1`, which splits on `ApprovalDocument/` | **NO** — that is the URL segment. |
| `_api/web/lists/getbytitle('<library>')` | ⚠ **YES — 404.** |
| Connector `List Name` picked from the dropdown | Usually **NO** — see below. |

⚠⚠ **NOR IS THE YELLOW *"Failed to retrieve dynamic outputs … 'GetTable' failed with status code
'NotFound' … List not found"* BANNER.** The designer calls `GetTable` to populate the dynamic-content
picker and often cannot resolve a GUID typed in as a custom value, where the flow engine resolves it
perfectly. The same state can also render a red **`'List Name' is required`** over a box that visibly
holds the GUID. Seen on four flows on 2026-08-29, every one of which then ran green. **The run history
is the only judge — in both directions:** a flow that looks broken can be fine, and `HC folder approval`
read **On** for six days while every run failed.

⚠⚠ **A DROPDOWN SHOWING A RAW GUID IS NOT A BROKEN FLOW.** When a list is picked from the connector's
picker, Power Automate stores the **list GUID** and merely *displays* the title. After a rename the
designer can no longer resolve the title, so it shows the bare GUID — which looks broken and runs
perfectly. **Judge from the run history, never from how the designer renders a box.** Only a box
holding a literal *title* as a custom value is genuinely broken.

---

## 3. The fix: GUIDs, not the new titles

Replace `getbytitle('<title>')` with `lists(guid'<id>')` from the table above.

**This is already the project's own pattern**, not a new idea: Auto-route uses it twice
(`AuthorId` read, delete-by-id) and BOTH bulk-import auto-approve flows use it for the trigger and
the MERGE. Those are precisely the parts that survived this rename untouched.

Same typing as swapping a title, and it is **immune to rename #4** — which, on three renames in two
days, is not hypothetical.

---

## 4. Per-flow checklist

Tick as done. `⬜` = not yet checked.

### Auto-route (normal) — ✅ **REPAIRED AND VERIFIED 2026-08-29**

Both edits made and the chain re-tested end to end: a routed file came back with the real uploader
(`Crystal Kong`) as `Author`, `Created` seconds collapsed to `:00`, and the source deleted from the
approval library. **TWO title references, not one:**
- ⬜ `GetSourceUniqueId` — `getbytitle('Approval Document')` → `lists(guid'a9342528-…')`
  ⚠ **THE FIRST VERSION OF THIS CHECKLIST SAID AUTO-ROUTE WAS ONE EDIT, AND THAT WAS WRONG.** This
  action was added on 2026-08-23 with the audit flows, long after the 2026-08-08 spec this list was
  derived from — so it appears in neither document's action list. It failed on the first repair run.
- ⬜ `validateUpdateListItem` stamp — `getbytitle('Documents')` → `lists(guid'e322e3a5-…')`
  (the only action naming the DESTINATION library; its fingerprint is `outputs('Copy_file')?['body/ItemId']`)
- ✅ `Get source author` — already `lists(guid'…')`
- ✅ delete by item id (`Send an HTTP request to SharePoint 1`) — already `lists(guid'a9342528-…')`
- ✅ `Copy file` — path-based; `If another file is already there` = **Replace** is correct
  (the client's 2026-08-27 decision — do not change it back)
- ⬜ audit `Create item` — `LibraryName` is FREE TEXT, update the prose to the new title
- ⬜ anything inside `Condition 2` — not yet inspected

⚠⚠ **THE RUNBOOKS UNDER-REPORT, SO OPEN EVERY HTTP ACTION.** Actions were added to these flows
after their runbooks were written and nothing updated the runbook. **Do not trust an action list —
this one included. Click through every `Send an HTTP request to SharePoint*` in each flow.**

⚠ **A FAILED `GetSourceUniqueId` POISONS THE AUDIT DEDUPE.** If a run reaches the audit `Create item`
with no id, it writes a row with a **blank `ItemUniqueId`** — and `Already logged` filters on that
column, so the blank row MATCHES for any later event whose id also fails to resolve. The flow then
answers *already logged* and **writes nothing at all, on a run reporting success**. After repairing,
delete every blank-`ItemUniqueId` row from `CRS Audit Log`.

⚠ **THE STAMP FAILING IS NOT COSMETIC.** The source delete is gated on the stamp succeeding, so
while this is broken: approved files are **left in the approval library**, where they are visible to
every PIC in the unit (the delete is load-bearing for security, not housekeeping), and the routed
copy carries the FLOW OWNER in `Created By` instead of the uploader. **Check what has routed since
the rename and repair those before assuming this is only a flow fix.**

### HC Auto Route — ✅ **REPAIRED AND VERIFIED 2026-08-29**

⚠⚠ **IT HAD BEEN BROKEN SINCE 27 AUGUST — A DAY BEFORE THE RENAME THIS RUNBOOK IS ABOUT, AND NOBODY
NOTICED.** The 27 August HC rename (`HC Document`, singular — the one 1.0.294.0 was written for) broke
it; the CODE was fixed that day and the FLOW was not. **Nine HC documents sat approved in the HC
approval library for two days**, visible to every HC-cleared person in their units — exactly what the
routing delete exists to prevent.
- **The seconds dated it precisely.** `mermaid - 27-08-26.docx` stamped `26 Aug 22:12:00` ✅; every
  copy from `27 Aug 09:16:42` onward carried real seconds ❌. Working on the 26th, broken on the 27th.
- **THIRD SILENT HC FAILURE**, after `HC folder approval` (six days red) and the six-fault HC clone.
  **Whenever a library is renamed, read the HC flows' run history the same day** — fixing the SPFx
  package is half the job, and the half that reports nothing is the flow.
- **The mis-attribution was material, not cosmetic**: six documents claimed `Clarence Cho` and
  actually belonged to `Reene Low`.

**Verified**: nine files routed, new item ids, `:00` seconds, real uploaders; both same-named files
survived into different folders; the HC approval library returned to its four genuinely Pending items.

**The three references (all confirmed correct):**
- ⬜ same stamp → `lists(guid'e4fb3b2c-…')` ⚠ **HC, not the normal library**
- ⬜ `LibraryName` free text
- ⚠ This flow's 2026-08-19 clone shipped with **six** wrong library references. Re-read every action.

### Folder approval / HC folder approval — 2026-08-29

✅ **Both trigger conditions verified present and correct**:
`@equals(triggerOutputs()?['body/{IsFolder}'], true)` on each. The approval-bypass risk is closed.
⚠ Trigger conditions live in the trigger's **Settings** tab and are INVISIBLE on the canvas — a
screenshot of the flow diagram tells you nothing about them.

✅ **`CRS — Approve new folders in Approval Document` needed NOTHING** — its trigger `List Name` was
already the GUID `a9342528-…`. That is why the normal folder approval survived the rename untouched.

⚠⚠ **`HC folder approval` COULD NOT EVEN SAVE**, and the error names the cause:
*"'GetTable' failed with status code 'NotFound' … List not found"*. Its trigger `List Name` held the
old TITLE `HC Approval Document`. **So it had been pointed at a nonexistent list since 27 August and
never fired.** Fixed by replacing it with the GUID `2311cd83-…` as a custom value.

⚠⚠ **CORRECTION TO §2 OF THIS RUNBOOK: A CONNECTOR `List Name` IS *NOT* ALWAYS A GUID.** This
document originally said a `List Name` picked from the dropdown stores the GUID and is therefore
safe, and that only a raw-GUID *display* was worth ignoring. Both halves hold — but a box holding a
literal TITLE is as broken as a `getbytitle`, and this flow proves it. **Check every trigger's
`List Name`, not just the HTTP actions.** Two places per flow, not one.

**Consequence found and cleared:** six below-Unit folders (`2024`, `Agreement`, `2024`, `Term Sheet`,
`2024`, `Tax Return` — ids 461, 462, 464, 465, 483, 484) were stuck **Pending** in the HC approval
library. Each is invisible to every HC PIC but its author, so the next PIC filing into that unit
cannot see it, tries to create it, and gets a **duplicate-name error** — which blames folder creation
for a flow failure days earlier. Approved by hand 2026-08-29.

⚠ **FIXING THE FLOW DOES NOT BACKFILL.** These triggers are **When an item is created**, not
created-or-modified, so touching a stranded folder does nothing — its creation event has passed.
Existing Pending folders must be approved BY HAND (native Approve/Reject, or the
`Approve/reject Items` view), or the tree rebuilt.

✅ **VERIFIED CLEAN 2026-08-29** — both flows firing again, and no folder left Pending in either
approval library.

⚠⚠ **AND `CRS — Approve new folders` WAS BROKEN TOO — its HTTP action still said
`getbytitle('Approval Document')`.** Two Failed runs on 28 August at 17:23 and 17:36:
*"List 'Approval Document' does not exist at site with URL …"*. **It had SUCCEEDED at 11:35 and 12:10
the same day**, which dates the NORMAL library's rename to that afternoon — NOT to the 04:41 timestamp
on the library's root folder, which reflects content changes and is not when a title changed. **Do
not date a rename from `RootFolder/TimeLastModified`; date it from the first flow failure.**

⚠⚠ **USE THE FILTERED QUERY TO FIND STRANDED FOLDERS — THE PLAIN LISTING GIVES A FALSE ALL-CLEAR.**

```
/_api/web/lists(guid'<library>')/items?$select=Id,FileLeafRef&$filter=FSObjType eq 1 and OData__ModerationStatus eq 2&$top=500
```

An unfiltered folder listing on this site is truncated at ~50,000 characters, and on 2026-08-29 the
cut fell at item **1421** — hiding **1757 `Approval Papers`** and **1759 `Certificates`**, both still
Pending. Every visible row read `0`, so the truncated response looked like proof the library was
clean. The filtered form returns only the problem rows and cannot be truncated into a false negative.
**Sixth instance of `sp-capped-read-reads-as-absent`.** The same query also caught HC **464 `2024`**,
which had been missed from a batch approve of six.

**The reference to fix:**
- ⬜ the MERGE's `getbytitle(…)` → `a9342528-…` / `2311cd83-…`

⚠⚠ **THIS MERGE SETS `OData__ModerationStatus: 0` — APPROVED.** A wrong library id here does not
merely fail: if any item holds that id in the library it *does* reach, it approves it. That is the
2026-08-25 `HC folder approval` bug, which failed safely only by luck. **After fixing, read `All
runs` and account for every `Succeeded`.**

### Audit — approval activity / deletions / Documents deletions (+ 3 HC clones)

**Trigger `List Name` and the `IsFolder = false` condition are CONFIRMED on all four deletion flows,
2026-08-29:**

| Flow | Trigger | `IsFolder` | `Library` prose |
|---|---|---|---|
| `Audit — approval deletions` | ⚠ held the literal `Approval Document` → `a9342528-…` | ✅ | ⬜ |
| `Audit — HC approval deletions` | ✅ `2311cd83-…` | ✅ | ✅ already new title |
| `Audit — Documents deletions` | ✅ `e322e3a5-…` | ✅ | ✅ — but `Details` still says *"Deleted from Documents."* |
| `Audit — HC Documents deletions` | ✅ `e4fb3b2c-…` | ✅ | ✅ |

- ✅ `Audit — approval activity` trigger `List Name` — already `a9342528-…`
- ✅ `Audit — approval activity` / `Audit — HC approval activity` `GetUniqueId` — GUIDs
- ⬜ `Library` free text on `Create item` — still the OLD title on both activity flows
- ✅ audit list GUID `84ed065f-…` — **do not change**

⚠⚠ **`Audit — approval deletions` WAS THE ONLY DELETION FLOW ACTUALLY BROKEN, AND IT HELD A LITERAL
TITLE — the third `List Name` in this repair to do so.** §2 of this runbook originally claimed a
connector `List Name` is GUID-backed and therefore safe. **Three flows have now disproved it**
(`HC folder approval`, `CRS — Approve new folders`'s HTTP action, and this). The pattern is not
"dropdown vs custom value" — it is simply **unknowable without opening every trigger**.

**Consequence while it was broken:** every deletion in the normal approval library since 28 August went
**unlogged**. Routing deletes are excluded by the `Was routed` branch, so the missing rows are precisely
the genuine ones — a Head of Unit deleting a pending or rejected file. **Nothing backfills**: the
trigger fires on deletion and those events have passed. `CRS Audit Log` has a two-day hole; do not read
it as complete for that window.

⚠ **`Audit — Documents deletions` AND ITS HC TWIN HAVE NO `Was routed` CHECK**, correctly — nothing
routes files OUT of the approved side today. **That changes the day the seven-year archive mover lands**:
`MoveTo` fires this same trigger, so every archived document would log as `Deleted`. The mover must write
its own row first and these two flows must gain the check-first branch.

⚠ **ALL SIX AUDIT FLOWS REFERENCE `CRS Audit Log` BY *TITLE*, NOT BY GUID.** It works only because
that list has not been renamed. If it ever is, **all six break at once**, and the failure is the quiet
kind: no audit rows, on runs that look fine until somebody goes looking for history. Convert them at
the same time as SDG's (§7).

### NotifyApprovers — ✅ **REPAIRED AND VERIFIED 2026-08-29**

Trigger and `GetDoc` both already held `a9342528-…`. **The proof it had been broken was the gap in the
run history: last run 27 August, nothing on the 28th or 29th** — the rename killed it and no approver
had been emailed since, silently.

**Verified**: an upload into `GHO / GCA / Elaeis Garden (Hospitality)` by an uploader-only account
produced one email naming the document, the unit and the real uploader, addressed to both members of
`GHO_GCA_EG_APPROVER`.

⚠ **IT DOES NOT FIRE THE INSTANT A FILE LANDS, AND THAT IS CORRECT.** Two reasons compound, and either
alone makes a working flow look dead if you refresh immediately: the trigger **polls**, and the upload
form uploads *then* tags — so the first firing stops at `HasUnit` before `UnitTid` exists and the
**second** run is the one that sends. Give it a minute before opening the run history.

### HCNotifyApprovers — ✅ **REPAIRED AND VERIFIED 2026-08-29**
- ✅ trigger `List Name` — `2311cd83-…`
- ✅ `GetDoc` `List Name` — `2311cd83-…` ⚠ **the library is named in TWO actions**; a clone with one
  swapped and one not fires on HC documents while reading the normal library
- ✅ `getbytitle('CRS Group Map')` — that list was not renamed

✅✅ **THE RECIPIENT LIST PROVED THE FILTER, WITHOUT READING THE CONFIG BACK.** Two emails, nine minutes
apart, same unit: the NORMAL document went to **both** members of `GHO_GCA_EG_APPROVER`; the HC document
went to **one** address — the single member of `GHO_GCA_EG_APPROVER_HIGHLY_CONFIDENTIAL`. A filter of
`Role eq 'APR'` would have matched the group of two. So the HC flow is demonstrably on `'APRHC'`.

⚠ **THE TECHNIQUE GENERALISES, AND IT IS THE SECOND TIME IT HAS PAID HERE** (see 2026-08-25, where an
empty plain group made the same point in reverse). **When two groups that a filter must distinguish have
DIFFERENT membership counts, who receives the mail identifies which branch ran.** Stage that asymmetry
deliberately rather than testing against two equally-populated groups, where both filters look alike.

⚠⚠ **DO NOT TEST HC AGAINST AN EMPTY APPROVER GROUP.** `GHO_GCA_EG_APPROVER_HIGHLY_CONFIDENTIAL` read
**nobody in it yet** on 2026-08-29. An empty group sends **no email and raises no failure**, which is
indistinguishable from a broken filter — the exact false negative recorded in CLAUDE.md from the
2026-08-25 build. Put somebody in the group first, or pick a unit that already has one.

⚠ **THE THING THE HC TEST IS ACTUALLY CHECKING** is that the mail goes to the
`_APPROVER_HIGHLY_CONFIDENTIAL` group and NOT to the plain `_APPROVER`. If an ordinary Head of Unit
receives it, the Group Map filter is still `Role eq 'APR'` instead of `'APRHC'`, and plain approvers are
being told the filename and unit of Highly Confidential documents — the disclosure the whole HC split
exists to prevent.

### Bulk-import auto-approve ×2
- ✅ **Nothing.** The runbook specified the library GUID as a custom value throughout.

### CRS — Audit request activity
- ✅ Nothing for the RENAME. Triggers on `CRS Requests`, which was not renamed.
- ✅ **BUT TWO EXPRESSIONS CHANGED 2026-08-30** for the revoke work (1.0.317.0):
  - `EventKind` gained a `Revoked` → `ShareRevoked` branch. Ends `)))` now, was `))`.
  - `ActorEmail` prefers `RevokedBy`, falling back to `DecidedBy`. Ends `))`, was `)`.

⚠⚠ **`EventKind` ALONE WOULD HAVE BEEN A DEFECT DRESSED AS A FIX.** A revoke writes `Status` and
`DecisionNote` and deliberately **not** `DecidedBy` — that field records who let the share through.
So with only the label fixed, a `ShareRevoked` row would name **the approver, not the revoker**, and a
Head of Department can revoke a share a Head of Unit approved. On an audit log a wrong name is worse
than no row: it gets believed. Both edits, always.

⚠ **`empty()`, NEVER `coalesce()`, in the ActorEmail fallback.** `coalesce` returns the first
NON-NULL value and an empty string is not null — so every request revoked before the `RevokedBy`
column existed would hand the log a blank actor. `empty()` is true for null and `""` alike.

⚠ **The flow edit is safe BEFORE the column exists.** `?['body/RevokedBy']` is safe navigation, so an
absent column yields null and the expression falls back to `DecidedBy` — today's behaviour. There is
no ordering constraint between this and the deploy.

---

## 4b. Free-text `Library` / `Details` — ✅ DONE 2026-08-29/30

Neither field is wired to anything, so a rename does not break them — it makes them **lie**. Nothing
fails; an admin simply reads an audit row naming a library that has not existed since August and cannot
tell a stale label from a genuinely different library.

| Flow | `Library` | `Details` |
|---|---|---|
| `Audit — approval deletions` | ✅ `Approval for Document` | ✅ rewritten |
| `Audit — Documents deletions` | ✅ `Restricted & Confidential Document` | ⚠ last seen as *"Deleted from Documents."* |
| `Audit — approval activity` | ✅ `Approval for Document` | token |
| `Audit — HC approval activity` | ✅ `Approval for Highly Confidential Document` | token |
| `Auto-route` | ✅ `Restricted & Confidential Document` | token (`From: body/{FullPath}`) |
| `HC Auto Route` | ✅ `Highly Confidential Document` | token |

⚠ **THE `&` IS A PLAIN AMPERSAND, NOT THE FULLWIDTH `＆`** the term store requires. They are different
characters, the wrong one sits there looking correct, and it is easy to carry across from a copied term
label. Type it.

⚠ **`Audit — HC approval activity` REFERENCES `CRS Audit Log` BY TITLE**, where its five siblings use the
GUID `84ed065f-…`. Both work today because that list has not been renamed. **The inconsistency is the
danger**: rename it and that one flow dies while the others carry on, leaving a PARTIAL audit trail —
which reads as a complete one. Convert it with SDG's (§7).

---

## 5. Traps already paid for — do not rediscover these

⚠ **THE `fx` BUTTON DESTROYS A MIXED URI.** Clicking it converts the whole Uri box into ONE
expression, so a literal `lists(guid'…')` and an `@{…}` token can no longer coexist; the action goes
to **Invalid parameters** and the `Method` box empties. Recovery: clear the box and retype as PLAIN
TEXT with the token inserted mid-string. Cost two round trips on 2026-08-23.

⚠ **A TRAILING NEWLINE ARRIVES BY PASTING.** After pasting into any flow field, press End and check
for a stray line break. In `ItemUniqueId` it silently breaks *History of this file*; the tell is
`xml:space="preserve"` in the REST response, invisible in a list view.

⚠ **HEADER KEYS TAKE NO COLON.** `Accept`, not `Accept:`. With the colon the header does not exist
and the symptoms look unrelated to each other.

⚠ **`Details` AND `LibraryName` ARE PROSE AND WILL STILL SAY THE OLD NAME.** Every free-text field
needs the same pass as every list reference, or the audit log names a library that no longer exists.

---

## 5a. ⚠ THE NATIVE Approve/Reject IS THE ONLY ROUTE THAT REACHES A STRANDED FILE

Both the Approval page and the Bulk Approve panel run `documentsFileClash` and REFUSE every stranded
file. **SharePoint's own Approve/Reject command bypasses that guard entirely** — it writes
`OData__ModerationStatus` directly and has no idea the guard exists. CLAUDE.md records that as a
*risk*; for this repair it is the way through, and `Copy file`'s **Replace** then swaps the bad copy
cleanly (verified 2026-08-29: nine files, new ids, no orphans left behind).

⚠ **A STRANDED FILE IS ALREADY `Approved`, SO A PLAIN Approve WRITES NOTHING** and the *created or
modified* trigger never fires. Confirmed live: an Approve pass produced **no new runs at all**.
**Reject, then Approve** — the Reject run takes the False branch and does nothing, the Approve run
routes.

⚠ **THIS IS AN ADMIN REPAIR, NOT A HABIT.** Approvers keep using the Approval page: the guard being
bypassed here is the one that stops an approval silently replacing a different document.

## 5b. Repairing what routed while a flow was broken — the method that worked

Anything approved during the outage was **copied but never deleted** (the delete is gated on the
stamp), so it sits Approved in the approval library with a badly-stamped twin in the destination.

⚠ **THE APPROVAL PAGE AND THE BULK APPROVE PANEL BOTH REFUSE THESE**, with *"a document of that name
is already in the destination library, and approving would replace it"*. That is `documentsFileClash`
working correctly — and the thing it clashes with is **the file's own orphaned copy**. The guard
compares names at a moment in time; it cannot ask *"is the destination file a stale copy of this same
document?"*. (It could: both carry `SubmissionFileId`, so matching stamps would mean replacing is
safe. NOT BUILT.)

**So the repair is:**
1. **Delete the badly-stamped copies** from the destination library — they recycle, and the sources
   still exist, so nothing is at risk.
2. **Touch each source** — Details pane, edit any field. Auto-route triggers on *modified* and the
   items are already Approved, so it takes the routing branch.
3. Verify with the two queries in §6.

⚠ **HOW TO TELL A STAMPED FILE FROM AN UNSTAMPED ONE, FOR FREE: READ THE SECONDS ON `Created`.**
Auto-route writes it as `M/d/yyyy h:mm tt`, which has no seconds field — so a stamped file always
reads `:00`. Real seconds (`:33`, `:01`) mean `Created` is the COPY time and the stamp never ran.
This settled the 2026-08-29 diagnosis in one glance, after a green run had already been mistaken for
proof.

⚠ **THE `Uploader` COLUMN IS THE OTHER FREE TELL.** It is bound to `Author`, so on the SOURCE item it
still shows the real uploader even when the routed copy is mis-attributed to the flow owner.

⚠ **AND THE ~8-HOUR `Created` SKEW IS PRE-EXISTING, NOT A SYMPTOM OF THIS.** Both routing flows write
a UTC value in a locale format SharePoint reads as local, so a 09:08 upload is stamped 01:08. Already
recorded in CLAUDE.md; do not chase it as part of a rename repair.

## 6. Verification, in order

1. Upload one document → it reaches `Approval for Document`.
2. Approve it → **`Auto-route` succeeds**, the file lands in `Restricted & Confidential Document`,
   `Created By` is the UPLOADER, and the source is **gone** from the approval library.
3. Repeat as an HC-cleared PIC through `Approval for Highly Confidential Document`.
4. `CRS Audit Log` shows `Uploaded` → `Approved` → `Routed` on ONE `ItemUniqueId`.
5. Approver was emailed, once, naming the right document.
6. ⚠ Read `All runs` on both folder-approval flows and account for every `Succeeded`.

---

## 7. For SDG

SDG's libraries have **not** been renamed, so its flows are unaffected — but they carry the same
`getbytitle` fragility. **Convert them to GUIDs at migration**, using SDG's own list ids from the
query in §1. A title-based flow is one client rename away from this whole exercise.
