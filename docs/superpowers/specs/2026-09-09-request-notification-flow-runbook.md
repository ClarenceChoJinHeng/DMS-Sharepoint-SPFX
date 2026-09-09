# `CRS — Notify request activity` — the six share/delete request emails, in ONE flow

Power Automate only. **No code change.** Builds item 2 of
`2026-08-28-email-bundling-and-templates-design.md` §7. That design carries the client's templates
6–11 verbatim; this is how to build them.

**NOT BUILT.**

---

## 1. Six emails, one flow

Templates 6–11 are six emails about one list. **Do not build six flows** — that is six copies of the
same trigger, the same approver lookup and the same recipient resolution, and the copy that drifts is
always the rarely-exercised one.

| `RequestType` | `Status` | Template | Goes to |
|---|---|---|---|
| Share | Pending | 6 | the unit's approver(s) |
| Share | Approved | 7 | the requester |
| Share | Rejected | 8 | the requester |
| Deletion | Pending | 9 | the unit's approver(s) |
| Deletion | Approved | 10 | the requester |
| Deletion | Rejected | 11 | the requester |

Same shape as `CRS — Audit request activity`'s `EventKind` Compose. **Build it as a Save As of that
flow** — it already has the trigger, the dedupe and the branch expression.

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

⚠ **AND FALL BACK TO `Unit` WHEN `UnitTermGuid` IS BLANK.** It is optional on the row (added later),
and `matchesUnit` in `requests.ts` already does exactly this. A row predating it must still notify
somebody rather than silently notifying nobody.

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

**Recommendation: drop those two lines from template 6 and confirm it with the client.** If they want
them, add a `Get item` by `ItemUniqueId` inside the Pending branch only, and ⚠ remember managed
metadata comes back as `Label|GUID` from the connector — `first(split(...?['Value'], '|'))`, the trap
that shipped a raw GUID into the reminder email on 2026-08-31.

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

- **⚠ `Failed` notifies nobody.** An approval whose share or delete broke records `Outcome: Failed`
  and there is no template for it — so the requester was told *"being handled"* and never hears again.
  The templates have no seventh case; this needs one, or a line in template 8/11.
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

1. Save As `CRS — Audit request activity`; repoint the trigger, delete its `Create item`.
2. `EventKind` + `Should notify` — **verify `Revoked` and `Cancelled` reach `Skip`** by revoking a
   share and cancelling a request. Expect a run that completes and sends nothing.
3. Approver lookup + `HasRecipients`. Test the EMPTY group first, then a populated one.
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
