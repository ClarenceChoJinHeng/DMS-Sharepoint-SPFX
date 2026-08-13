# Audit Log — design

**Date:** 2026-08-13
**Status:** APPROVED — implementing
**Supersedes:** the `<P> Deletion Log` list in `2026-08-04-deletion-request-approval-design.md` §7 (see §9)

---

## 1. What the client asked for

> "client wants us to build audit log into webpart… They want to track every single thing, a file being
> approve or more."
>
> "give all the data we could possible give, if those data we cannot give then we tell the why we can't."

So this is two deliverables, not one:

1. **A record** of everything the system can observe, written as it happens.
2. **An honest statement, on the page itself**, of what is not in the record and why — because an audit
   log that looks complete and isn't is worse than no audit log. Silence must not read as "nothing
   happened".

## 2. Decisions taken

| # | Decision | Rationale |
|---|---|---|
| D1 | **Super admin only** can read it | Matches "everything under CRS Settings is Super Admin only". The rows carry file names and full paths, so any reader-scoping bug is a confidentiality incident. Per-unit visibility stays possible later — the rows already carry `Segment`/`UnitPath` to trim on. |
| D2 | **Hybrid writers** — Power Automate for file events, web part code for admin events | A flow sees an approval done in the library UI; our code never does. Our code sees a permission grant; a flow never does. Neither alone is complete. |
| D3 | **One list, indexed, always date-filtered** | SharePoint lists hold millions of items; the 5,000 figure limits *unindexed view queries*, not the list. The viewer never asks for "everything". |
| D4 | **No auto-delete, ever** | A log that trims itself defeats its purpose. If the client wants retention trimming it is a flow they own, not code we ship. |
| D5 | **Version history ON, and nobody but Owners and the service account can write** | An audit trail editable by the people it audits is not an audit trail. |
| D6 | **Purview is not used** | See §4. Three independent blockers, any one of which is fatal. |

## 3. What gets logged

### 3.1 By flow — the file lifecycle, however it was performed

| Event | `EventType` | Captured |
|---|---|---|
| Uploaded to the approval library | `Uploaded` | uploader, file name, full path, size, every metadata column, confidentiality |
| Approved | `Approved` | approver, approval comment, file, path |
| Rejected | `Rejected` | approver, rejection comment, file, path |
| Approved file moved to `Documents` | `Routed` | source path → destination path (logged by Auto-route, which is already at that moment) |
| Deleted from the approval library | `Deleted` | who, file name, path (Head of Unit's `DELS`) |
| Approved document deleted from `Documents` | `Deleted` | who, path (Head of Department's `DEL`) |

**The previous approval status is not captured at write time.** A create-or-modify trigger receives the
new item, not the old one; there is no "old value" to read. The viewer derives the transition from the
log's own previous row for the same `ItemUniqueId`. Same information, and it degrades honestly — if
there is no earlier row, the viewer says "previous status not recorded" rather than guessing.

### 3.2 By our own code — everything a flow cannot see

| Event | `EventType` | Captured |
|---|---|---|
| Upload refused by file-type policy | `UploadRefused` | who, file name, the extension, the reason |
| Access granted / revoked | `AccessGranted` / `AccessRevoked` | actor, persona, group, folder, library, roles |
| Reconciliation run | `ReconciliationRun` | folders created / renamed / skipped, ACLs granted, orphans repaired, errors |
| Folder structure saved | `StructureChanged` | segment, old chain → new chain, staged or applied |
| Subtree migration run | `MigrationRun` | files moved, folders tidied, collisions and their resolutions, strays |
| Segment created | `SegmentCreated` | segment, term set, tiers, columns created in both libraries |
| Abbreviation saved | `AbbreviationChanged` | term label, old code → new code, and that a rename follows on the next run |
| File-type policy changed | `PolicyChanged` | extension, enabled / disabled / added |
| Group Map row written | `GroupMapChanged` | group, segment, unit, role, added / edited / deleted |

**`UploadRefused` is the one event class Purview also cannot see** — the file never reaches SharePoint,
so no server-side audit exists. Someone repeatedly trying to upload `.exe` is precisely what an audit
log should surface, and here our log is strictly better than the tenant's. Worth saying to the client.

### 3.3 One row per run, not one per item

Reconciliation touches hundreds of folders; a migration moves hundreds of files. **Each run writes ONE
row**, carrying counts in `Title` and the per-item lines in `Details`. One row per folder would bury
every other event in the log the first time someone reconciles, which is the same as losing them.

## 4. What cannot be logged, and why

This text ships **on the page**, not only in this spec.

1. **Who viewed or downloaded a file.** Nothing notifies SharePoint code of a read. It exists only in
   Purview's unified audit log, and we cannot reach it:
   - the only supported route from a web part is Graph's audit-log query API with
     `AuditLogsQuery.Read.All`, which is a **tenant-wide** admin consent;
   - API permission requests are processed from the **tenant** app catalog, and this package ships to
     the **site collection** app catalog on the CRS site only, so it cannot even ask;
   - the API is an async job — submit, poll, read — so it could not drive a live page anyway.

   If the client's real question is *"who saw this document"*, the answer is a Purview report their
   tenant admin runs. That is not a gap we can close from inside the package.
2. **Failed access attempts.** A denied request never reaches our code. Purview only.
3. **Anything before deployment day.** The log starts empty and cannot be back-filled.
4. **Permission changes made in SharePoint's own "Manage access" dialog** rather than our Folder Access
   page. Our code is not in the loop. Mitigation is procedural: admins use the web part.
5. **Term store edits** — adding, renaming or deleting a term. No API notifies us. We log the
   *consequence*: reconciliation records the folder rename, or the orphaned term it found.
6. **Direct edits to the config lists in the list UI.** List version history is the only trace, which
   is why D5 turns it on.

## 5. The list

**`<P> Audit Log`** — resolved through `naming.ts`/`spNaming.ts` like every other list
(`LIST_SUFFIX.auditLog = "Audit Log"`), never a hardcoded literal, because the client renames `DMS_` to
`CRS_` at import and a missed literal is a runtime 404.

Generic list, `BaseTemplate: 100`. Column **internal names carry no spaces** — created under the
internal name, then retitled, using the existing `ensureColumn` in `shared/spColumns.ts`. A
`_x0020_` in an audit column would cost the same debugging time `Vendor_x002f_CustomerName` already has.

| Internal name | Type | Indexed | Purpose |
|---|---|---|---|
| `Title` | Text | — | the one-line human summary, e.g. `Approved — Q1 Tax Return.pdf` |
| `EventTime` | DateTime | **yes** | when the event actually happened |
| `EventType` | Text | **yes** | one of the values in §3 |
| `Outcome` | Text | — | `Success` / `Refused` / `Failed` |
| `ActorName` | Text | — | display name at the time |
| `ActorEmail` | Text | **yes** | the stable identity to filter on |
| `Source` | Text | — | which writer produced the row (`Flow:ApprovalActivity`, `UploadForm`, `FolderAccess`, …) |
| `LibraryName` | Text | — | `Approval Document`, `Documents`, or blank |
| `ItemUniqueId` | Text | **yes** | per-file identity — survives a path change, so it is what "this file's history" groups on |
| `ItemName` | Text | — | file or folder name |
| `ItemPath` | Note | — | server-relative path at the time of the event. `Note`, because paths exceed the 255-char Text limit |
| `Segment` | Text | — | filtering now, reader-trimming later |
| `UnitPath` | Text | — | `GHO/GF/CORU` — the permissioned prefix |
| `Details` | Note | — | the event-specific payload as readable lines |

**`EventTime` is separate from the built-in `Created`** and is the column everything sorts and filters
on. A flow writes minutes after the fact, and on a busy morning `Created` order is not event order.
It is written `M/D/YYYY h:mm tt` per the site locale (gotcha #1) and displayed `DD/MMM/YYYY`.

**`EventType` is Text, not Choice — deliberately.** Writing a value absent from a Choice column's
`Choices` **fails**, so the day someone adds an event type in code, every row of that type is silently
lost. Choice would also need `Manage Lists` to extend, the same trap the file-type panel already
documents. Text cannot fail this way, and it indexes and filters identically.

**No `Person` column.** A flow writing a Person field must resolve the user, and a resolution failure
loses the whole row; a Person field also breaks when someone leaves the tenant. Name and email as text
are what an audit trail wants anyway — the value as it was, not a live lookup.

### 5.1 Provisioning

The page **provisions its own list** on first load by an admin — create list, create columns, set the
indexes, enable version history. The client cannot run PowerShell (standing constraint), so anything
requiring a script would not get done.

Provisioning is idempotent and safe to re-run: every step checks first. A non-admin who opens the page
before it is provisioned sees "not set up yet — ask an administrator to open this page once", never a
failed create.

### 5.2 Permissions

**A manual step, not part of provisioning**, and the page says so on the provisioning screen. Breaking
inheritance needs to name the service account, and the code does not know which account that is — a
guess would either grant the wrong identity or, worse, break inheritance and lock the flows out of the
list they are about to write to. So the code provisions the schema and the admin sets the permissions.

Break inheritance on the list. Owners get Full Control; the service account gets Contribute (it is what
the flows run as); **nobody else is granted anything**. This works because of how the writers split:
file events are written by flows as the service account, and admin events are written by admins. **No
uploader or approver ever needs write access**, so the log is tamper-resistant by construction rather
than by policy.

## 6. The writer

Two modules, matching the `naming.ts` / `spNaming.ts` split already in the codebase:

- **`src/shared/auditLog.ts`** — SPFx-free and unit-tested: the `EventType` constants, the row shape,
  `buildAuditRow()`, and `summarize()` which composes `Title` from the event. Pure, so the summary
  wording and the truncation rules are pinned by tests.
- **`src/shared/spAuditLog.ts`** — the write: `writeAudit(sp, siteUrl, event)`.

### 6.1 Best-effort, but never silent

`writeAudit` **must not fail the action it is logging.** Refusing an upload because the audit write
failed would be a worse outcome than an incomplete log.

But a swallowed failure makes the log quietly wrong, which is the failure mode this whole feature
exists to prevent. So:

- the user's action always proceeds, and its own result is reported normally;
- the failure is logged to the console with the HTTP status and body;
- for **admin** actions the screen shows a non-blocking warning — *"the change was applied but was not
  recorded in the audit log"* — because an admin can act on that, and an audit gap they know about is
  worth far more than one they don't;
- for **uploader** actions there is no message. An uploader cannot fix it and does not need to know.

`Outcome: "Failed"` is for the *audited action* failing, not for the audit write failing. A row that was
never written has no outcome to record — which is the honest limit of any self-hosted log, and is
stated in §4 alongside the rest.

### 6.2 Truncation

`Details` is a `Note` column and large but not unbounded. A run producing thousands of lines is
truncated with an explicit `… N more lines not recorded` marker. Truncation must always announce
itself; a quietly clipped detail block is indistinguishable from a complete one.

## 7. The viewer

Its own web part — **`CRS Audit Log`**, `src/webparts/auditLog/`, its own bundle. Not a tab on CRS
Configuration: a log is not a setting, and it gets its own page and its own Page-scope access row.

- **Feed by default**, newest first, last 7 days.
- **Filters**: date range, event type (multi-select), actor, and free text over file name and path.
  All applied **server-side** via `$filter` on the indexed columns with `$orderby EventTime desc`, so
  the query cost does not grow with the list.
- **Paging** by `$skiptoken`/`$top`, never a client-side slice of everything.
- **"History of this file"** — one click from any row filters to that `ItemUniqueId`, showing the
  file's whole life across path changes. This is the question actually asked during a dispute.
- **CSV export** of the current filtered view, reusing the `groupExportCsv.ts` pattern. Auditors ask
  for a file, and the alternative is them being given list access.
- **A collapsed "What this log cannot tell you" panel** carrying §4 verbatim.
- Dates render `DD/MMM/YYYY` with the time, per the agreed format.

Empty states are distinguished, because they mean different things: **not provisioned** (an admin must
open the page once) vs **no events match these filters** vs **could not read the list** (a permissions
or transient problem — never presented as "nothing happened").

## 8. The flows

Power Automate config is not in source control, so this section is the only record of it. Rebuild
verbatim on SDG's tenant. **Sign in as the service account before creating any of them** — a flow runs
under its connection, the connection is created implicitly by the first action, and a flow built as a
person dies silently when that password changes.

Three new flows, plus one action added to an existing one:

| Flow | Trigger | Writes |
|---|---|---|
| **Audit — approval activity** | When an item is created or modified, on the approval library | `Uploaded` when moderation status is Pending and this is a create; `Approved` / `Rejected` on status 0 / 1 |
| **Audit — approval deletions** | When an item is deleted, on the approval library | `Deleted` |
| **Audit — Documents deletions** | When an item is deleted, on `Documents` | `Deleted` |
| **Auto-route** (existing) | — | one added `Create item` action logging `Routed` with source and destination paths |

Documents *creations* need no flow: they are all made by Auto-route, which logs its own action.

Two carried-over traps that apply here as much as anywhere:

- **Header keys must not include the colon.** The key box wants `Accept`, not `Accept:`. With the colon
  the header silently does not exist, and the symptoms look unrelated to each other.
- **A deletion trigger gives limited item properties.** Capture `ID`, `FileLeafRef` and the path from
  the trigger body; do not expect metadata columns. Where a value is unavailable the row records it as
  blank rather than omitting the event.

## 9. Relationship to the Deletion Log spec

`2026-08-04-deletion-request-approval-design.md` §7 defines a separate `<P> Deletion Log` list with the
same append-only, Owners-only shape. It was never built — nothing reads or writes it, though
`LIST_SUFFIX.deletionLog` is primed in `spNaming.ts:135`.

**The audit log absorbs it.** Two overlapping audit trails would leave a permanent question about which
is authoritative, and deletion is already an `EventType` here. When the deletion-request feature is
built it writes `Deleted` rows to this list, with the request reason and approver in `Details`.
`LIST_SUFFIX.deletionLog` stays in place for now — removing it is a separate, reviewable change — but
nothing new should read it.

## 10. Out of scope

- **Reader scoping below super admin** (D1). The columns support it; the UI does not.
- **Alerting** — no email or Teams notification on an event. A log is pull, not push.
- **Retention trimming** (D4).
- **Anything from Purview** (§4).
- **Editing or deleting rows from the page.** There is deliberately no UI for it at all.

## 11. Implementation order

1. `LIST_SUFFIX.auditLog`, and `ensureColumn` extended with a `DateTime` kind. Small, and everything
   else needs it.
2. `shared/auditLog.ts` + tests — pure, so the row shape and summaries are pinned before anything
   writes them.
3. `shared/spAuditLog.ts` — provisioning and the write.
4. The viewer web part, with the three empty states. Shippable and useful as soon as anything writes.
5. Wire the code-side writers, highest value first: file-type policy, Folder Access, reconciliation,
   abbreviations, structure/migration/segment, Group Map, upload-refused.
6. The flows, from §8.

Steps 1–4 are shippable on their own: the page provisions the list and shows admin events as they are
wired in step 5.

## 12. Test plan

Unit (pure, Jest): summary composition for every `EventType`; truncation announces itself; a missing
actor or path degrades to blank rather than throwing; `buildAuditRow` never emits a field the list does
not have.

Site: provision on a clean site and confirm the list, all fourteen columns' internal names, the four
indexes and version history; re-open and confirm provisioning is a no-op; a policy change writes exactly
one row; a reconciliation run writes exactly one row with counts; filters narrow server-side; "history
of this file" follows a file across a subtree migration; CSV matches the filtered view; the three empty
states each appear for the right reason; a non-admin gets the "ask an administrator" state and no rows.
