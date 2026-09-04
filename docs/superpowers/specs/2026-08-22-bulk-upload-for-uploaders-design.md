# Bulk Upload for uploaders, via the approval library

**Date:** 2026-08-22
**Status:** design agreed, not yet built
**Supersedes the central decision of** `2026-07-23-bulk-upload-direct-to-documents-design.md` — that
spec's premise is an admin tool writing straight to `Documents`. It stays as the record of why.

**Client instruction:** *"Bulk upload is now allowed for all uploaders to be used, client doesnt want
admin to do the job."* And, on why approval is skipped: *"the reason for bulk uploads is because it is
for client to upload old documents that is already approved before."*

---

## 1. What was proposed, and why it is not being built

The client's plan was a NEW library — same folder structure, hidden from everyone, uploaders bulk-upload
into it, a flow auto-approves and pushes to `Documents`.

**That library already exists. It is `Approval Document`.** It has the folder tree, the per-unit ACLs,
content approval, and an Auto-route flow that moves approved files into `Documents` preserving
`Created By` (verified 2026-08-08). The only difference in the proposal is that approval is skipped —
which is a flow, not a library.

⚠ **"Hidden from everyone else" cannot mean flat.** If uploaders write into it they need permission on
it, and without per-unit folder ACLs every uploader would see every other unit's files while they sat
there. So the new library would need the full tree and grants — which is what makes it a duplicate
rather than a drop box. The cost of building it anyway:

- a **fifth and sixth** library (HC needs its own pair), each needing reconciliation to build ~2,400
  folders and ~980 grants per segment — MHO's first run was 71 minutes across four libraries
- an entry in `LIBRARY_ROLES`; `allLibraryTitles()` grows, so tier columns, `SubmissionId`/`BatchId`,
  `Full Name` and content types all follow
- resolution in `naming.ts`, with the both-halves-or-neither rule
- **two more Power Automate flows**, including the `{IsFolder}` polarity trap — both directions of that
  mistake have already been made once each on the existing pair

**Decision: reuse `Approval Document`. No new library.**

## 2. The shape of the change

Bulk Upload already resolves the **approval library's** unit folder — `lookupFolderMapping` on the leaf
term, then `resolveMappedFolder` by UniqueId, which is rename-proof — and *then* swaps the library
segment to reach the mirrored `Documents` folder (`toDocumentsPath`).

**So repointing it is not a rewrite. It is deleting the swap.** The file is written into the folder
that was already resolved, and `validateUpdateListItem` addresses the approval library instead.

### 2.1 Target library

`targetListTitle()` returns `libApiTitle("Staging")`, or `libApiTitle("StagingHC")` when the chosen
confidentiality routes to HC. Never a literal — the title and URL segment differ (gotcha #12), and
`libApiTitle` translates the logical key at the API boundary.

⚠ **The file is placed by folder id and tagged by library title, and those two must agree.** That
mismatch is what made HC uploads unable to tag on any site until 1.0.170.0: item ids are per-LIST, so
tagging the wrong list 404s on a good day and writes onto a *different document* on a bad one.

### 2.2 Page access

`pageAccessPolicy.ts`'s `/bulk/i` rule stops being `adminOnly` and becomes `["UPL", "UPLHC", "APR"]` —
the same three as the upload form, and for the same reasons:

- `UPLHC` because `hou` and `pic_hc` carry no literal `UPL`; it is a superset that
  `LIBRARY_ROLES.Staging` lists on the normal approval library too
- `APR` because a Head of Unit group mapped before the 2026-08-15 persona correction holds only `APR`
  and `DELS`, and keying on upload roles alone would strand every already-provisioned HoU

The `/bulk/i` rule must stay ordered **before** `/upload/i` — "bulk-upload" contains "upload".

**Migration: re-run reconciliation.** Page grants are derived and asserted, so the run adds the
uploader groups to `Bulk-Upload.aspx`.

### 2.3 The write probe is turned ON

Bulk Upload is existence-gated only, deliberately — *because* it wrote into `Documents`, where `UPL` is
Read by design since 2026-08-09, so an `AddListItems` probe would have correctly and silently emptied
the form for every PIC.

**That reasoning expires with this change.** The target is now the approval library, where a PIC holds
`CRS Upload`, so the probe is both correct and necessary: it is what stops an uploader picking a unit
whose ACL has not been granted yet and meeting a 403 at write time. `filterReachablePaths` +
`probeFolderUploadAccess`, exactly as the upload form uses them.

An unreadable Folder Map or an inconclusive probe must still offer EVERYTHING — a transient error that
empties every dropdown takes the form down site-wide.

## 3. Auto-approve

### 3.1 The marker

A `BulkImport` Yes/No column on all four libraries, written `"true"` by Bulk Upload only — the same
string-boolean shape `LegallyPrivileged` already uses through `validateUpdateListItem`.

⚠ **Written CONDITIONALLY, on a confirmed column.** One unknown field name fails the WHOLE
`validateUpdateListItem` call (gotcha #4) — every column lost, not just the missing one. So the same
`libraryHasRefColumns` pattern applies, and it must **filter server-side**, never `$top` a field
collection and search it (memory `sp-capped-read-reads-as-absent` — that exact defect cost the
submission stamp the same day).

**Degrading is safe and VISIBLE here**, unlike the submission stamp: with no marker the flow never
fires and the files sit in the approval queue, where the Head of Unit can see and approve them. A
failure that leaves work in a queue is not a silent one.

Reconciliation asserts the column every run over `allLibraryTitles()`, rather than relying on a
provisioning step — a library added later would otherwise stop auto-approving with nothing to say so.
Sixth instance of that rule.

### 3.2 The flow

A third Power Automate flow, built **as the service account** (§0 of the Auto-route spec — a flow runs
under its connection, and the connection is created implicitly by the first action, so the builder's
account is baked in for life).

⚠ **CORRECTED BEFORE BUILDING — an "item is created" trigger could never have worked.** Bulk Upload
uploads the file and *then* tags it, so at creation the item exists with **no `BulkImport` value**: the
condition would be false, the flow would never fire, and — being a flow that never triggers — it would
leave **no run history to look at**. The same silent-failure shape as setting `{IsFolder}` to `true` on
Auto-route. The marker arrives on the first MODIFY, so the trigger must be *created or modified*.

- Trigger: **When an item is created or modified**, approval library (GUID as a custom value)
- Trigger condition — all three clauses, and each one is load-bearing:
  `@and(equals(triggerOutputs()?['body/{IsFolder}'], false), equals(triggerOutputs()?['body/BulkImport'], true), not(equals(triggerOutputs()?['body/{ModerationStatus}'], 'Approved')))`
- Action: MERGE `{"OData__ModerationStatus": 0}`

⚠ **The third clause is what stops an infinite loop.** The MERGE modifies the item, which re-fires this
same trigger — and `{IsFolder}` and `BulkImport` are both still true. Without a stop clause it runs
until Power Automate's loop protection cuts it off, burning the daily quota. This is not hypothetical:
it is exactly why the folder-approval flow uses *created* rather than *created or modified* (§3), and
that option is unavailable here.

Full build steps, the test order, and what to do if `{ModerationStatus}` is not exposed on the trigger
are in `2026-08-22-bulk-import-auto-approve-flow-runbook.md`.

Approving fires the **existing, verified** Auto-route flow, which copies to `Documents`, stamps
`Author`/`Editor`/`Created`, and deletes the source. Nothing about that flow changes.

⚠ **The `{IsFolder}` polarity is the whole safety.** `false` here (files only). `true` would approve
the folders Bulk Upload ensure-creates.

⚠ **The marker condition is the second half of that safety.** Without it this flow approves EVERY file
in the approval library the moment it is created — abolishing approval for the whole system, silently,
on a run that reports success.

### 3.3 Cost per file

Two flow runs per document — auto-approve, then Auto-route. `MAX_FILES` is already **50**, so one press
is bounded at 100 runs. Worth stating to the client: an exhausted daily quota means the next real
approval is not routed, **silently**, so a large historical import should be spread across days.

## 4. Name clashes against the approved side

⚠ **THE GAP THIS CLOSES, spotted by Clarence.** Both upload web parts check the destination folder for
an existing file name. Once Bulk Upload targets the approval library, that check probes the approval
library — and the file's real destination is `Documents`. A historical document already filed there
would not be seen, and the clash would surface inside Auto-route's `Copy file`, where the outcomes are
a file stuck approved-but-not-routed or an approved record overwritten.

**So both web parts now also check the mirrored approved-side folder**, reusing `swapLibrarySegment`
(`hcRouting.ts`, already tested) to derive the path, and `probeFolderByPath`.

⚠ **The path must go through the OData parameter alias form** — `GetFolderByServerRelativeUrl(@f)?@f='…'`
— never an inline quoted literal. A deep encoded path returns **HTTP 400, not 404** past roughly 330
characters, and `Form.tsx` silently failed exactly this way on nested paths (gotcha #9).
`probeFolderByPath` already does it correctly.

Behaviour on a clash:

- **Upload form: refuses that file**, naming the library it is already in. Consistent with its existing
  rule — never overwrite, fail the one file, let its siblings through.
- **Bulk Upload: skips that file** with the reason, and **does not offer to replace**. Its existing
  "replace?" prompt stays for a clash in the approval library, where overwriting a pending file is
  recoverable. On the approved side it is not: a PIC holds Read on `Documents` and could not overwrite
  anyway, and an admin should not silently replace an approved record.
- ⚠ **Fails OPEN.** An unanswerable check allows the upload, which is exactly today's behaviour — this
  is an additional guard, and a transient error must not become a new way to block uploading.

## 5. Submission references

Bulk Upload was deliberately not stamped with `SubmissionId`/`BatchId` — it was admin-only and wrote
straight to the approved side, so it was not "the submission of the uploader themselves".

**Both halves of that reasoning are gone.** Its files now land in the approval library and are authored
by the uploader, so they appear in My Submissions (which reads `AuthorId eq me` in both libraries). It
stamps them, on the same conditional-write rule as the marker.

## 6. What the client must be told

1. **Bulk Upload is for historical documents that were already approved elsewhere.** Nothing in the
   software enforces that — the client's own instruction is that uploaders are told the rule. A PIC can
   put a brand-new document through it and it will reach `Documents` unreviewed.
2. **`Created By` becomes the person who bulk-uploaded, not the original author.** Auto-route preserves
   the uploader faithfully; it has no way to know who wrote the document years ago. The real date is
   captured in `DocumentDate`, which the form asks for.
3. ⚠ **HC bulk upload does not work yet.** The HC Auto-route flow has never been built, so an HC file
   would be auto-approved and then sit in `HC Approval Document` with nothing to move it.
4. **Two flow runs per file**, bounded at 100 per press by the 50-file cap. Spread a large import.

## 7. Out of scope

- Bulk-approve in the approver's queue. Considered, and it is the better answer if the client ever
  wants bulk uploads REVIEWED — but they do not, for this workflow.
- Retro-stamping documents already in `Documents`.
- Any change to the Auto-route or folder-approval flows.
