# `CRS — Execute approved deletion` — flow runbook

**No file calls this document** — a Power Automate build runbook under `docs/superpowers/specs/`,
matching every other `*-runbook.md` in this project. Read alongside
`2026-09-17-proxy-deletion-via-power-automate-design.md`, which carries the reasoning; this carries
the exact build steps.

**User, 2026-09-17:** confirmed the code-side alignment (PIC/Approver hold no direct delete grant,
HOD holds none and no longer decides a Deletion request either, everything routes through this one
flow) and said **"yes"** to building it.

**Status: NOT BUILT.** The code half (persona roles, `Requests.tsx`, `MySubmissions.tsx`, the
`hodUnits` decision-scope restriction) is built and verified — see CLAUDE.md's entries dated
2026-09-17 for both. **Until this flow exists and is turned on, pressing Delete anywhere in the app
writes a request row and deletes NOTHING.**

---

## 0. Prerequisite — `crs@sdguthrie.com` needs real delete rights

Add `crs@sdguthrie.com` to **`CRS Owners`** before building anything. That group already holds Full
Control on the web regardless of any folder ACL — no new permission level, no new Group Map row, and
consistent with how every other elevated flow in this project is authorised (design spec §4).

⚠ **Sign in as `crs@sdguthrie.com` before creating the flow, from the first action.** The connection
a flow uses is baked in at creation; built as a person, it stops silently the day that person's
password changes — the same caution every other flow in this project's history repeats.

## 1. The flow

**Name:** `CRS — Execute approved deletion`
**Trigger:** **When an item is created or modified** → List Name **`CRS Requests`**.

⚠⚠ **DELETE ANY INHERITED TRIGGER CONDITIONS if this is built as a Save-As of an existing flow**
(Settings → Trigger conditions). The 2026-09-09 request-notification runbook found `{IsFolder}` and
`{ModerationStatus} = 'Pending'` inherited from a document-library flow, both of which read `null` on
a plain list item — a flow with either condition present **fires zero times, silently**, and looks
identical to a finished, working flow with no run history to inspect. **Building this one FRESH
(not as a Save-As) avoids the whole class of problem** and is the recommended route, since nothing
about this flow's action shapes overlaps enough with an existing one to be worth inheriting.

## 2. The condition

One Condition action, two rows joined by the default **And**:

- `RequestType` (`triggerOutputs()?['body/RequestType']`) **is equal to** `Deletion`
- `Status` (`triggerOutputs()?['body/Status']`) **is equal to** `Approved`

⚠ **Enter both through the ordinary field — the `fx` editor is not needed here**, since these are
plain string comparisons against trigger fields, not composed expressions. If either value is typed
with a trailing space or newline (a recurring defect in this project's flow history — invisible
whitespace has broken at least four flows so far), the condition will silently never match. Retype
rather than paste if in doubt.

## 3. True branch — resolve and recycle

**Action: Send an HTTP request to SharePoint**

```
Site Address: <this site>
Method: POST
Uri: _api/web/GetFileById(guid'@{triggerOutputs()?['body/ItemUniqueId']}')/recycle()
Headers:
  Accept: application/json;odata=nometadata
```

⚠⚠ **DO NOT add `odata-version: ""` to this action pre-emptively.** That header exists in this
project's SPFx client code (`spAuditLog.ts`, `Requests.tsx`, etc.) specifically because SPFx's
`SPHttpClient` injects `odata-version: 4.0` on every request — a fact about that ONE client library,
not about Power Automate's own "Send an HTTP request to SharePoint" connector. No flow anywhere in
this project's history has needed it. If this action genuinely comes back with *"Parsing JSON Light
feeds or entries in requests without entity set is not supported"*, THEN try it — treat it as a live
diagnosis, never a pre-emptive cure carried over from a different code path (same caution the
2026-09-17 reindex runbook already states for exactly this reason).

⚠ **`GetFileById` is web-scoped** — it resolves the document by GUID regardless of which library it
currently lives in. This is the entire reason the client-side code (`Requests.tsx`'s
`resolveDeletionTarget`, `MySubmissions.tsx`'s `writeApprovedDeletionRequest`) resolves and writes
the CURRENT `ItemUniqueId` onto the row **before** setting `Status = Approved` — by the time this
flow runs, `ItemUniqueId` is already correct and this action needs no per-library branching of its
own.

## 4. Success branch — log the deletion

**"Configure run after"** on the action below: **is successful** on the recycle action from §3.

**Action: Create item** (native SharePoint connector action, on the **`CRS Audit Log`** list —
matching every other audit-writing flow in this project, e.g. Auto-route's own `Create item` /
`Create item 1`. Never a raw HTTP POST for this one — there is no odata/entity-type subtlety here
that the native action does not already handle).

| Field | Value |
|---|---|
| `Title` | `concat('Deleted — ', triggerOutputs()?['body/ItemName'])` |
| `EventType` | `Deleted` (literal text) |
| `EventTime` | `utcNow()` |
| `Outcome` | `Success` (literal text) |
| `ActorEmail` | see below |
| `ActorName` | see below |
| `Source` | `Flow:ExecuteApprovedDeletion` (literal text) |
| `ItemUniqueId` | `triggerOutputs()?['body/ItemUniqueId']` |
| `ItemName` | `triggerOutputs()?['body/ItemName']` |
| `ItemPath` | `triggerOutputs()?['body/ItemUrl']` |
| `Segment` | `triggerOutputs()?['body/Segment']` |
| `UnitPath` | `triggerOutputs()?['body/Unit']` |
| `Details` | see below |

**`ActorEmail`** — design spec §3.5: *"for now real person… `DecidedBy` for an approved-then-executed
request, `RequestedBy` for a self-approved one — never the flow's own identity."*

```
if(
  empty(triggerOutputs()?['body/DecidedBy']),
  triggerOutputs()?['body/RequestedBy'],
  triggerOutputs()?['body/DecidedBy']
)
```

**`ActorName`** — `DecidedBy`/`RequestedBy` are email addresses, not display names (confirmed in
`RequestRow`'s own source: both are written from `context.pageContext.user.email`). Matching this
project's own precedent for exactly this situation (*"the `ActorName` display-name fix — split on
`@` rather than showing the raw email — applied to both routing flows"*):

```
first(split(
  if(
    empty(triggerOutputs()?['body/DecidedBy']),
    triggerOutputs()?['body/RequestedBy'],
    triggerOutputs()?['body/DecidedBy']
  ),
  '@'
))
```

**`Details`** — mirrors the wording `Requests.tsx`'s own (now-unreachable for this path, but
still-present-elsewhere) `writeAudit` call already used, so the audit trail reads consistently
whichever writer produced the row:

```
concat(
  'Requested by ', triggerOutputs()?['body/RequestedBy'], ' on ', triggerOutputs()?['body/RequestedAt'], '. ',
  'Reason: ', triggerOutputs()?['body/Reason'], '. ',
  if(
    empty(triggerOutputs()?['body/DecidedBy']),
    'No approval needed — carried out automatically.',
    concat('Decided by ', triggerOutputs()?['body/DecidedBy'], ' on ', triggerOutputs()?['body/DecidedAt'], '.')
  )
)
```

⚠ **`ItemPath` may read the PRE-ROUTING path in the rare case where `Requests.tsx` had to fall back
to the `SubmissionFileId` stamp** (a pending-stage request approved after its document was already
routed — the 2026-09-10 "a request survives its document being routed" fix). `resolveDeletionTarget`
updates `ItemUniqueId` in that case but does **not** rewrite `ItemUrl`, so this field can occasionally
name the OLD approval-library path rather than the new `Documents`/`HC Documents` one. Cosmetic only
— `ItemUniqueId` is what actually resolved and was recycled, and is what the reader should trust.

## 5. Failure branch — record what happened, never lose it

**"Configure run after"** on the action below: **has failed** on the recycle action from §3.

**Action: Update item** (native SharePoint connector action, on **`CRS Requests`**, item id
`triggerOutputs()?['body/ID']`):

```
Status: Failed
DecisionNote:
  concat(
    triggerOutputs()?['body/DecisionNote'], ' ',
    'Automated deletion failed (HTTP ', string(outputs('Recycle_the_file')?['statusCode']), ').'
  )
```

⚠ **Reads `outputs('Recycle_the_file')`, so name the recycle action from §3 exactly `Recycle_the_file`
when adding it** (or adjust this expression to match whatever name Power Automate assigns). If the
action was renamed with spaces, the internal reference uses underscores — check the `fx` picker's own
suggestion rather than typing it blind.

⚠⚠ **THIS BRANCH IS WHAT MAKES A FAILED PROXY DELETION VISIBLE AT ALL.** Without it, a recycle that
404s (file already gone) or 403s (a genuine permission gap, e.g. `crs@sdguthrie.com` not yet added to
`CRS Owners` per §0) leaves the row silently `Approved` forever, with nobody told the deletion never
happened — the exact silent-failure shape this project's own house rules exist to prevent.

## 6. Dedupe — deliberately none

**Design spec §3.4, client's own words:** *"I guess we can reuse the current one"* — no dedicated
dedupe marker column. `Status eq 'Approved'` is the trigger condition on its own, and
`GetFileById(...)/recycle()` on a document already in the recycle bin simply errors harmlessly (the
Failure branch above records it, but nothing is double-deleted). **Do not add a marker field
defensively** — if this proves noisy in the run history once live, that is the trigger to revisit
this section, not a reason to build it pre-emptively now.

## 7. What "it worked" looks like

- A PIC deletes their own pending file → `MySubmissions.tsx` writes a `CRS Requests` row with
  `Status = Approved` and no `DecidedBy` → this flow fires, recycles the file, writes a `Deleted` row
  attributed to `RequestedBy` (the PIC).
- A HoU approves a PIC's deletion request on an approved document → `Requests.tsx` writes
  `Status = Approved` with `DecidedBy` set → this flow fires, recycles the file, writes a `Deleted`
  row attributed to `DecidedBy` (the HoU).
- Either way: the document is gone from the library (recoverable for 93 days, in the recycle bin),
  and exactly ONE `Deleted` audit row exists for it — see §8 for why that second part needs a
  follow-up edit to four EXISTING flows before it is actually true.

## 8. ⚠⚠ STILL NEEDED AFTER THIS FLOW WORKS: stop the double-logging on four existing flows

Design spec §3.5. Once this flow performs every proxy deletion as `crs@sdguthrie.com`, the FOUR
existing generic delete-trigger flows — `Audit — approval deletions`, `Audit — Documents deletions`,
and their HC clones — will **also** fire (they trigger unconditionally off the library's own delete
event and read `DeletedByUserName` from the recycle-bin metadata, which will now always read
`crs@sdguthrie.com` for a proxy deletion). Left alone, **every proxy deletion logs twice**: once
accurately (this flow, real actor) and once inaccurately (the old flow, service account as actor).

**Fix, scoped to those four flows only — add ONE condition clause to each, before their own
`Create item`:**

```
DeletedByUserName is not equal to crs@sdguthrie.com
```

⚠⚠ **TREAT THIS WITH THE SAME CARE THIS PROJECT'S OWN HC-CLONE HISTORY REPEATEDLY WARNS ABOUT.**
Read each of the four flows' exported definition before touching it — confirm which condition
action each one actually has today, and which library reference each one carries (the HC pair has a
documented history of exactly one reference per clone being left unswapped). Change nothing else in
any of the four.

**Do this only after §1–7 above are built, verified, and turned on** — editing the four existing,
already-working audit flows before the new flow exists would add a condition that can never matter
yet, and adds risk to four production flows for no immediate benefit.
