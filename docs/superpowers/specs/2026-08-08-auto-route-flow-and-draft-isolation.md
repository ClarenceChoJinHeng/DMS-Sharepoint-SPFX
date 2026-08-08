# Auto-route flow + draft isolation — implementation and migration runbook

**Date:** 2026-08-08
**Status:** IMPLEMENTED and verified end to end on `/sites/ClarenceDMSTesting`
**Supersedes:** the per-uploader-isolation parts of `2026-08-07-role-model-simplification-design.md`
and reverses the "out of scope" half of memory `dms-per-uploader-isolation-rejected`.

> ⚠ **Power Automate configuration is NOT in source control.** This document is the only record
> of it. If a flow is rebuilt on another tenant, rebuild it from §4 exactly — several of the
> settings here look cosmetic and are not.

---

## 1. What was asked for, and what is actually deliverable

The client wants an uploader (PIC) to see only their own files, while Head of Unit, Head of
Department and C-Level see everything.

| Library | Isolation | Deliverable? |
|---|---|---|
| **Approval Document** (approval/staging) | PIC sees own pending + rejected; Head of Unit sees all | **YES — built and verified** |
| **Documents** (post-approval) | PIC sees own approved files only | **NO — not expressible** |

### Why Documents is impossible

Only moderation hides an item from someone holding Read on its folder. To *see* a hidden item
you need **`Approve Items`** — and in SharePoint `Approve Items` **cannot be separated from
`Edit Items`**; ticking it auto-ticks `Edit Items`, and unticking `Edit Items` unticks
`Approve Items`. **Verified in the permission-level editor 2026-08-08.**

So oversight roles would need edit rights on every document just to read them, which the client
has explicitly ruled out. The alternatives were priced and rejected:

- **`ReadSecurity = 2`** (item-level read) — the exemption is `Manage Lists`, evaluated at
  **library** scope. Approvers hold their level on the **folder**, so they would be restricted
  too and approval would break. It also hides every folder from uploaders (folders are items
  authored by reconciliation, not by the PIC), and it is library-wide with no per-unit opt-out.
  It additionally requires PnP/PowerShell plus an Entra app consent this project cannot obtain.
- **Unique ACL per file** — works, but consumes one of the list's 50,000 unique permission
  scopes per document, with degradation well before that.
- **Per-uploader folder tier** — would work and scales (scopes bounded by headcount), but
  changes the folder architecture, which the client has fixed.

**The unit folder is the smallest confidentiality boundary this system has.** If two people's
approved documents must never be mutually visible, they belong in **different units** — one
term, one abbreviation, three groups.

A `Created By = [Me]` filtered view in Documents is fine as a **convenience**, and must never be
described as isolation: the user can switch to All Documents at any time.

## 2. How approval-library isolation works

| Item | Moderation state | Who sees it |
|---|---|---|
| Segment / Department / Unit folders | **Approved** | anyone with Read on the folder |
| Year / Document Type folders | **Approved** (see §3) | anyone with Read on the folder |
| An uploaded file | **Pending** | the **author** + anyone with `ApproveItems` on that folder |
| A rejected file | **Rejected** | same |
| An approved file | Approved → **removed by the flow** | n/a |

Settings on **Approval Document**:

- Require content approval = **Yes**
- Draft Item Security = **Only users who can approve items (and the author of the item)**
  (`DraftVersionVisibility = 2`)

**The author exception works for FILES** — verified 2026-08-08 with a guest account that saw its
own pending upload. It does **not** save an uploader from a pending **folder** created by someone
else, which is what §3 exists for.

**This holds only because there is no view-only role in that library.** `LIBRARY_ROLES.Staging`
is `UPL`, `APR`, `DELS`; Head of Department has no access at all and the base MEMBER group is
excluded. A view-only role added here would see nothing but its own uploads — the same wall as
Documents. `SEGVIEW`/`GLOBAL` are already banned from this library in code for the same reason.

## 3. Folder approval — required, two mechanisms

The upload form ensure-creates `Year` and `Document Type` folders **as the uploader**, so they
arrive **Pending**. Their author can see them; a second PIC cannot — and therefore cannot browse
to their own file inside. Uploaders cannot fix this: setting the status needs `ApproveItems`.

**Backlog:** `C:\Users\clare\Downloads\approve-folders.js` — console script, previews by default,
`await approveFolders({apply:true})` to write. Ran 2026-08-08: 6 folders approved.

**Built and verified 2026-08-09** — flow `CRS — Approve new folders in Approval Document`, separate
from Auto-route.

- Trigger: **When an item is created**, list = the library **GUID** as a custom value
  - **Not** "created or modified": the action below *modifies* the item, so that trigger
    re-fires itself, the condition is still true (still a folder), and it loops until Power
    Automate's loop protection stops it.
  - **Split-on tracking ID must not be empty.** Saving with a present-but-blank value fails with
    *"The 'clientTrackingId' value is not valid. The value cannot be null or all whitespace
    characters."* The UI renders blank and absent identically, so clearing it does not help — put
    `@{triggerOutputs()?['body/ID']}` in it. Same class as the empty trigger-condition row that
    reports "Invalid settings": **cleared is not the same as not there.**
  - The tracking-ID box and the trigger-condition box look alike and sit in the same panel. The
    condition row is the one with an **✕** beside it (conditions are a removable list); the tracking
    ID is a plain field. Pasting the tracking ID over the condition silently removes the
    `{IsFolder}` guard — see the warning below for what that costs.
- **Trigger condition** (Settings → Trigger conditions):
  `@equals(triggerOutputs()?['body/{IsFolder}'], true)`
- One action — `Send an HTTP request to SharePoint`:
  - `POST` · `_api/web/lists(guid'<library-guid>')/items(@{triggerOutputs()?['body/ID']})`
  - Headers: `Accept`, `Content-Type` = `application/json;odata=nometadata`,
    `X-HTTP-Method` = `MERGE`, `IF-MATCH` = `*`
  - Body: `{ "OData__ModerationStatus": 0 }` — integer, **unquoted**

> **The `{IsFolder}` guard is the entire safety of this flow.** If it ever approves a FILE, that
> file skips human approval, becomes visible to everyone with Read, and Auto-route ships it to
> Documents. Test it deliberately: upload a file and confirm it stays *Waiting for Approval*.

**Verified end to end 2026-08-09**, upload as a PIC (`clarencechojinheng`) with a second PIC
(`chocheetuck4`, guest) in the same `_UPL` group:

| Check | Result |
|---|---|
| Flow fires on the new `Document Type` folder | succeeded, 0.4s |
| `{IsFolder}` guard holds | the uploaded FILE stayed *Waiting for Approval* |
| Peer PIC can navigate the new folder | yes — reaches `2024/Tax Return` |
| Peer PIC cannot see the file inside | yes — folder renders **empty** for them |

The last two together are the deliverable: navigation works, isolation holds.

**Also fixed the same day — reconciliation's moderation merge.** It was sending `Full Name`,
`ContentTypeId` and `OData__ModerationStatus` in ONE merge, which SharePoint rejects with HTTP 500
*"You cannot change moderation status and set other item properties at that same time."*
Compounding it: **in a moderated library any property write re-pends the item**, so the earlier runs
that stamped `Full Name` had silently knocked the whole tree back to Pending, and uploaders then hit
*"The mapped unit folder no longer exists"* — the UniqueId lookup 404s for a user who cannot see a
pending folder. Nothing was deleted; the message named the symptom, not the cause. `FolderManager.tsx`
now writes properties first, then moderation status in its **own** merge, and re-approves when either
the folder was pending or the property write just re-pended it.

Reconciliation also has a folder-approval pass for the structural tiers, but it does **not** cover
the Year × Document Type grid (`recon_gridMode` is off, and the grid fast path would skip it
anyway). See `docs/2026-08-07-project-state.md` §3 item 5.

## 4. The Auto-route flow — exact configuration

Name: **Auto-route approved Pending folders to Documents Library**

```
When an item is created or modified        (approval library, GUID as custom value)
Get item                                   (same library, Id = trigger ID)
Condition   ?{ModerationStatus} is equal to  Approved
├─ True
│   Compose        source path bits
│   Compose 1      destination sub-path        <- see 4.1
│   Create new folder                          (Documents, path from Compose 1)
│   Copy file                                  (source -> Documents)
│   Get source author                          <- GET, 4.2
│   Send an HTTP request to SharePoint         <- stamp Author/Editor/Created, 4.3
│   Send an HTTP request to SharePoint         <- DELETE source, 4.4
│   Compose 2      body(...) of the stamp, for inspection
│   Send an email from a shared mailbox (V2)   notify the uploader
└─ False → 1 action (no-op)
```

**No trigger condition on this flow.** One was added by mistake on 2026-08-08
(`{IsFolder}` — it belongs on the folder-approval flow) and it silently stopped every file from
being routed, because the flow then only fired for folders.

### 4.1 Path split — `Compose 1`

```
substring(
  split(body('Get_item')?['{Path}'], 'ApprovalDocument/')?[1], 0,
  sub(length(split(body('Get_item')?['{Path}'], 'ApprovalDocument/')?[1]), 1))
```

Splits on the **URL segment** `ApprovalDocument/` (no space), not the title
`Approval Document`. Then strips the trailing character — the trailing slash, which `Copy file`
rejects. It was `Staging/` before the library was renamed; that failure is **silent** (the split
yields one element, `?[1]` is null) so it must be checked by reading the output, not by waiting
for an error.

### 4.2 `Get source author` — GET

```
GET  _api/web/lists(guid'<library-guid>')/items(@{triggerOutputs()?['body/ID']})?$select=AuthorId
Headers: Accept = application/json;odata=nometadata
```

Kept because `AuthorId` (integer) is occasionally useful, though §4.3 uses the claims string from
`Get item` instead. If the `Accept` header is missing or malformed the response is **verbose** and
the value lives at `d.AuthorId`, not `AuthorId`.

### 4.3 Stamp `Author` / `Editor` / `Created` — POST

```
POST _api/web/lists/getbytitle('Documents')/items(@{outputs('Copy_file')?['body/ItemId']})/validateUpdateListItem
Headers: Accept, Content-Type = application/json;odata=nometadata      (NO X-HTTP-Method, NO IF-MATCH)
```

```json
{
  "formValues": [
    { "FieldName": "Author",  "FieldValue": "[{'Key':'@{body('Get_item')?['Author']?['Claims']}'}]" },
    { "FieldName": "Editor",  "FieldValue": "[{'Key':'@{body('Get_item')?['Author']?['Claims']}'}]" },
    { "FieldName": "Created", "FieldValue": "@{formatDateTime(body('Get_item')?['Created'],'M/d/yyyy h:mm tt')}" }
  ],
  "bNewDocumentUpdate": true
}
```

- **`bNewDocumentUpdate: true` is what permits overriding `Author`/`Created`.** Without it the
  call succeeds and ignores them.
- **`Editor` is set from the AUTHOR's claims deliberately** — the client wants *Modified By* to
  read the uploader, not the flow.
- **`Created` must be `M/d/yyyy h:mm tt`** here (US locale). A MERGE to the same item wants
  **ISO 8601** instead. Two endpoints, two date rules.
- Person value format is `[{'Key':'<claims>'}]` — claims, not email.

**Verified 2026-08-08:** the Documents item read `by chocheetuck4 | mod chocheetuck4` with the
original upload timestamp, for a file uploaded by a guest and approved by an admin.

> **This was believed impossible for two hours.** `validateUpdateListItem` returned
> `HasException: false` with the correct `FieldValue` echoed back, and a `MERGE` on `AuthorId`
> returned `204`, while the item never changed. Both were caused by §5.1. Do not conclude a field
> is unwritable from a clean response — **always re-read the item.**

### 4.4 Delete the source — POST

```
POST _api/web/lists(guid'<library-guid>')/items(@{triggerOutputs()?['body/ID']})
Headers: Accept = application/json;odata=nometadata, X-HTTP-Method = DELETE, IF-MATCH = *
Body: empty
```

**Configure run after → only "Is successful"** on §4.3. A failed stamp must leave the source
intact so the run can be repeated.

**This step is load-bearing for the security model, not housekeeping.** An approved file left in
the approval library is visible to every PIC in the unit, and the isolation ends at the first
approval.

The connector's **`Delete file`** action was tried first and fails with *"The response is not in a
JSON format"* — its `File Identifier` does not resolve here, because the trigger is a **list**
trigger pointed at a library. Delete by item ID instead.

## 5. Gotchas that cost real time

### 5.1 Header keys must not include the colon — THE BIG ONE

Power Automate's header **key** box wants `Accept`, not `Accept:`. Typing the colon makes it part
of the header **name**, so the header does not exist. One typo produced three unrelated-looking
symptoms:

| Symptom | Real cause |
|---|---|
| `Get source author` always returned `odata=verbose` however often the `Accept` header was set | header was named `Accept:` |
| `validateUpdateListItem` reported success and changed nothing | `Content-Type:` meant no content type was sent |
| `The parameter AuthorId does not exist in method GetById` | `X-HTTP-Method:` meant the MERGE went as a plain POST |

**Check the raw Inputs of any HTTP action that misbehaves.** The keys appear verbatim, e.g.
`"Accept:": "application/json;odata=nometadata"` — the colon is visible there and nowhere else.

### 5.2 Every `Send an HTTP request to SharePoint` needs `Accept: application/json;odata=nometadata`

Without it the response is verbose and every property sits under `d`, so
`body('X')?['Field']` is **silently null** — and a null inside a JSON body produces
`{ "AuthorId": , ... }`, i.e. *"Not well formatted JSON stream"*.

### 5.3 The library picker never lists document libraries

`When an item is created or modified` shows only generic lists. Use **Enter custom value** with
the library **GUID**. Read the GUID off the site, do not guess:

```
_api/web/lists/getbytitle('Approval%20Document')?$select=Id,Title,BaseTemplate,RootFolder/ServerRelativeUrl&$expand=RootFolder
```

A dead GUID (the list was deleted and recreated) breaks the picker for the **whole flow**, so
every other action's dropdown also fails until the trigger resolves.

### 5.4 `Copy file` outputs

- `ItemId` — **integer list item ID** in the destination. Use this.
- `Id` — a URL-encoded **path**, not a GUID. `GetFileById` will not work with it.
- `Path` — **site-relative** (`/Shared Documents/...`), so it is rejected by
  `GetFileByServerRelativeUrl`, which wants `/sites/<site>/...`.

One run wrote to `items(3961)` which later 404'd — if a stamp 404s, read `Copy file`'s raw output
for that run rather than assuming `ItemId` is stale.

### 5.5 Documents library: title vs URL

Title is **`Documents`**, URL is **`/Shared Documents`**. Same trap as
`Approval Document` / `ApprovalDocument`. Address items by ID, never by path.

### 5.6 The flow fires several times per upload

`created or modified` fires on the file being created, on each folder being created, and again on
approval. All but one correctly take the False branch. A run whose True branch is **skipped** is
not a failure. Consider a trigger condition on moderation status if the noise becomes a problem.

### 5.7 Content approval on `Documents` hides everything the flow routes

Verified 2026-08-08. If `Documents` has **content approval on**, every item the flow creates there —
the `Year`/`Document Type` folders from `Create new folder` **and the copied file itself** — arrives
**Pending**, and is therefore invisible to every read-only viewer. The uploader gets an approval
email whose link is denied, and the unit sees a folder that appears not to exist. It looks exactly
like a permissions bug and is not one.

**How to spot it without a query:** the library's view bar shows **`Approve/reject Items`** and
**`Show All Files`**. SharePoint adds those two views *only* when moderation is enabled. Confirm with
`getbytitle('Documents')?$select=EnableModeration,DraftVersionVisibility`.

**Fix:** Library settings → Versioning settings → *Require content approval* = **No**. Everything
Pending becomes visible immediately; no re-approval pass is needed.

`Documents` must never be moderated — §1 already rules it out as an isolation mechanism, and this is
a second, independent reason. Approval happens in the approval library; anything reaching
`Documents` is approved by definition.

Nothing in the code enables it — reconciliation only *reads* `EnableModeration`
(`FolderManager.tsx`) to decide whether to stamp new folders Approved. If it is on, a human turned
it on.

### 5.8 A PIC needs the base `_MEMBER` group to open the link in the email

`*_UPL` grants upload on the approval library and **nothing in `Documents`** (see the persona table
in CLAUDE.md). So a PIC who is only in `_UPL` receives "your document is now available" and is
denied on the link. That is the model behaving correctly, not a broken link — note the denial URL
still carries `Type=item&listItemId=…`, i.e. SharePoint *resolved* the file and then refused.

A PIC who should be able to read their unit's approved documents must **also** be a member of the
unit's base group (e.g. `GHO_GF_CORU`). Uploading and viewing are deliberately separate grants.

### 5.9 The notification emails — expressions must track the library rename

Both emails (approved, and rejected in the `Condition 1` True branch) parse the folder path out of
`{Path}` by splitting on the library's **URL segment**:

```
Segment:       @{first(split(split(body('Get_item')?['{Path}'], 'ApprovalDocument/')?[1], '/'))}
File Location: @{replace(substring(split(body('Get_item')?['{Path}'], 'ApprovalDocument/')?[1], 0,
                 sub(length(split(body('Get_item')?['{Path}'], 'ApprovalDocument/')?[1]), 1)), '/', ' > ')}
```

They were authored against `Staging/`. After the rename `?[1]` became **null**, and
`split(null,'/')` / `substring(null,…)` / `length(null)` threw
`InvalidTemplate … the provided value is of type 'Null'` — failing the whole run at the last action.
Same class of failure as §4.1: **one stale token, silent null, loud error somewhere else.**

Two more differences between the branches, both easy to get wrong by copy-paste:

| | Approved | Rejected |
|---|---|---|
| File link | `@{replace(body('Get_item')?['{Link}'], '/ApprovalDocument/', '/Shared%20Documents/')}` | `@{body('Get_item')?['{Link}']}` — **unreplaced** |
| Date label | `Approved on` | `Rejected on` |

A rejected file is **not moved**; it stays in the approval library, where the author exception keeps
it visible to its uploader. Rewriting its link to `/Shared%20Documents/` points at a file that was
never there.

Also: the body must read **`Document_x0020_Type`**. The retired `Department_x0020_Type` fails
*quietly* — `?['Label']` on null just renders blank — so the email sends with an empty Document Type
and nothing in the run history flags it.

## 6. Migration checklist for a new site

1. Read the new library's **GUID** and confirm `BaseTemplate = 101`.
2. **Approval Document**: content approval **on**, Draft Item Security = **approver + author**.
3. **Documents**: content approval **off** — verify, do not assume (§5.7). If the view bar shows
   `Approve/reject Items`, it is on and every routed file will be invisible.
4. Recreate both flows from §3 and §4. Use a **service account** for both connections — with a
   personal account, a password change stops folder approval and routing silently.
5. Header keys with **no colons**; `Accept: application/json;odata=nometadata` on every HTTP action.
6. Run `approve-folders.js` once to clear any pending-folder backlog.
7. Test as a **non-admin**: upload as a PIC, approve as an approver, then verify
   - the Documents copy shows the PIC in `Created By` **and** `Modified By`, with the original `Created`
   - the source is **gone** from the approval library
   - a second PIC cannot see the first PIC's pending file
   - the uploader can **open the link in the email** — needs the base `_MEMBER` group (§5.8)
8. Clear any pre-existing `status=0` files from the approval library by hand — the flow only fires
   on change, so it will never collect them.
9. Check both email bodies for the library's own URL segment and `Document_x0020_Type` (§5.9).

## 7. Known outstanding

- Leftover approved files in the approval library from before §4.4 existed.
- **Rejected email `Reason`** — unverified whether `{ModerationComment}` from `Get item` carries the
  comment, or whether the branch's `Get Reject Comment` output is required. Test one rejection with
  a comment typed in.
- Some SharePoint groups still carry the `DMS_` prefix (e.g. `DMS_GHO_GF_CORU_UPL`) — re-run
  `rename-crs-groups.js` in preview mode to enumerate what the first pass missed.
- Both flow connections still run as a personal account; move them to the service account before
  handover — a password change stops folder approval and routing **silently**.
- ~~Reconciliation logs a moderation-approve failure as `ok: true`~~ — **fixed 2026-08-09.** The
  property write and the approve now have their own `try/catch`: a failed `Full Name` stays
  `ok: true` (cosmetic — the ACLs are still correct), a failed approve is `ok: false` and says the
  folder will be invisible to its uploaders. The approve also runs even when the property write
  threw, so a cosmetic failure can no longer cascade into an access one.
- `recon_gridMode` is **off**; Year/Doc-Type folders are created on demand and rely on the §3 flow,
  which makes that flow a single point of failure for navigation. Turning `gridMode` on would have
  reconciliation pre-create the grid **as an admin** (born Approved, flow off the critical path), at
  the cost of Year × Document Type folders per unit up front. Undecided.
