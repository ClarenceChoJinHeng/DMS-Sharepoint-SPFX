# A submission record that survives deletion

**Status:** design agreed with the client 2026-08-27. **One correction to the agreed design is made
here and it changes the primary key** — see §2. Nothing built at the time of writing.

Client: *"if I deleted the file the my submission file list also disappears, can we ensure that
doesn't happen? … Client wants to be able to record that file even if someone delete or replace … I
know if someone deletes the file I cannot open file and show the content but that is fine, then just
show this specific file is deleted but what matters is its recorded for that submission."*

---

## 1. Why the page loses a deleted file today

`MySubmissions.tsx` has **no record of its own**. It reads the two document libraries (four with the
HC pair, six with the archive) filtered `AuthorId eq me`, and `groupSubmissions` derives the whole
submission → batch → file shape from the rows that came back.

So a deleted file does not become *"deleted"* — **it stops existing**, and its batch quietly shrinks.
Nothing is wrong with the code; there is simply nothing left to read.

The fix is a list that records the upload as an **event**, independent of the document's survival.

---

## 2. THE AGREED KEY IS WRONG, AND IT WOULD HAVE MARKED EVERY APPROVED FILE "DELETED"

The agreed design says the row carries the file's **`UniqueId`**, that rows are matched to files on
`UniqueId` *"never name or path — both change"*, and that a row whose `UniqueId` is in **neither**
library renders **Deleted**.

**`UniqueId` does not survive approval.** Auto-route is **copy, stamp, delete source** (it has to be:
it recreates authorship on the copy). A copy is a new file, so it gets a **new `UniqueId`**, and the
source — carrying the id we stored — is then deleted. Recorded in two places already:

- `CLAUDE.md`, the archive section: *"a copy is a new file and therefore a NEW `UniqueId` …
  **Auto-route already does this today** — every document in `Documents` has a different `UniqueId`
  than it had in `Approval Document`."*
- `2026-08-22-seven-year-archive-design.md` §6, which is the whole reason the archive mover must use
  `MoveTo` rather than copying.

So on the agreed key the feature fails in its **most common case**:

| moment | stored `UniqueId` resolves? | row would render |
|---|---|---|
| uploaded, pending | yes, in the approval library | Pending, correct |
| **approved and routed** | **no — source deleted, copy has a new id** | **Deleted, wrong** |
| genuinely deleted | no | Deleted, correct |

Every successfully approved document would be reported to its own uploader as deleted, and
*"approved"* would be unreachable — the exact opposite of a record you can trust.

**Why nobody has hit this:** `MySubmissions` reads `UniqueId` live, per row, from whichever library
the row came from. It never has to survive anything. The comment at `MySubmissions.tsx:356` —
*"a GUID that survives the rename or move that a path does not"* — is true in its own context (a
request raised against a row just read) and is precisely the sentence that makes this key look safe.

### The corrected key: a stamped per-file reference

**A stamped column survives routing, and that mechanism is already load-bearing here.** SharePoint's
copy carries over every column that EXISTS at the destination — which is exactly why
`SubmissionId`/`BatchId` had to be created in all four libraries (2026-08-22), or a submission would
*"lose its files one by one as they were approved"*. That is proven; the join should use it.

So: a third reference column, **`SubmissionFileId`** (`SFI-<yyyymmdd>-<4>`), one per FILE, stamped at
upload beside the existing two, created by reconciliation on every library in `allLibraryTitles()`.

- **The join is on `SubmissionFileId`.** It survives Auto-route, the archive `MoveTo`, a rename and a
  move, because it is our data rather than SharePoint's identity.
- **`UniqueId` is still stored, as a secondary.** It is what a deletion or share request is raised
  against, and it is the fast path for a file still sitting where it was uploaded. It is **never** the
  thing that decides Deleted.
- **Deleted means: `SubmissionFileId` found in no library.** One condition, one key.

### Consequences to state plainly

- **A row is only as good as its stamp**, so **a row is written only when the file could be stamped.**
  A library missing the reference columns gets no stamp; recording an upload that can never be
  re-joined would manufacture a permanent false Deleted row, which is worse than recording nothing.
  Degrading to yesterday's behaviour is always correct.
- **Replace (the client's 2026-08-27 flow change, `nameConflictBehavior: 1`) reassigns the stamp.**
  The copy overwrites the destination file's columns, so the *replacing* file's `SubmissionFileId`
  lands on it and the *replaced* file's row stops resolving and reads Deleted. Defensible — the old
  content is a version now, not a live document — but it is a behaviour, not an accident, and the
  client should be told. **Unverified**, along with whether Replace versions at all.
- **Pre-existing files get no rows and keep behaving as they do now** (grouped by folder and date,
  vanishing if deleted). Nothing recorded them and no retro-fix exists. Accepted for the test site;
  **expect it at SDG cutover too.**

---

## 3. Status is joined live, never stored

The client's explicit instruction: *"make sure if the Approval status changes it will change as
well."*

The list decides only that a row **EXISTS**. Pending / Approved / Rejected, the routed path, the
rejection comment and the file size all come from the libraries on every load, exactly as today.

**Nothing about a document's state is written to the submissions list.** A stored status is a second
copy of a fact SharePoint already owns, and the copy is the one that goes stale — a failure this
codebase has paid for repeatedly.

---

## 4. Two cheaper options, ruled out with reasons

Recorded so they are not re-proposed.

- **The recycle bin.** Keeps name and path for 93 days — but an ordinary user sees only items **they**
  deleted, so an approver's or an admin's deletion is invisible to the uploader, who is the person
  asking. It is also full of noise: Auto-route deletes the approval-library copy of **every** routed
  file, so the bin's dominant content is successful approvals.
- **The audit log.** Already records `Uploaded` per file with the actor — but those rows carry **no
  submission or batch reference**, so a deleted file could not be put back in its batch, which is the
  whole requirement. It also cannot record a deletion's `UniqueId` at all: the delete trigger returns
  no path and no id.

---

## 5. Permissions — option 2, and it covers both lists

Client: *"I think best to go to option 2, I don't want to get a lash back later on."*

**Rejected (option 1): granting the site-entry group the `CRS Request` level**, which is what
`CRS Requests` has **today** via the 1.0.292.0 reconciliation pass. It works, and the client accepted
it earlier for requests — but it lets **any site member read every row** (filenames, paths, metadata,
who submitted what) and edit anyone else's. Delete is blocked, which was the core requirement, but the
**read** exposure is what they do not want. On a submissions list that exposure is strictly worse: it
is every upload anyone has ever made.

**Built instead: one reconciliation pass scoping BOTH `CRS Requests` AND `CRS Submissions` to the
groups that actually need them**, asserted every run so it cannot rot. Named as the right answer in
the 2026-08-21 requests section and left unbuilt then.

### Derived from ROLES, never from name suffixes

The agreed wording is *"the `_UPLOADER` / `_APPROVER` / `_HOD` groups"*. **Implemented as the roles
those suffixes stand for** — `UPL`, `UPLHC`, `APR`, `APRHC`, `DEPTVIEW` — read from the Group Map,
the same source reconciliation already uses for folder and page grants.

Matching on the suffix would be wrong for reasons this project has already paid for: names carry
**old spellings** (`_APR_HIGHLY_CONFIDENTIAL`, `_EMPLOYEE`), a group can be **renamed** while its id
and mappings survive (1.0.162.0), and `roleFromGroupName` **falls through to MEMBER** for anything
unrecognised — so a suffix-driven rule would silently miss real approver groups and could sweep in
hand-named ones. Roles are the fact; the name is a label on it.

Same shape as the derived-page-access pass: loop the **thing needing the grant**, derive from roles
held at any tier in any segment.

### THE SITE-ENTRY GRANT MUST BE REMOVED, OR OPTION 2 ACHIEVES NOTHING

The part the agreed note does not spell out. Adding ~470 per-group grants while leaving
`CRS_SITE_MEMBERS` holding `CRS Request` in place **changes nothing about the read exposure** — every
site member still reads every row. The whole point of option 2 is that grant going away.

But removal is the dangerous direction, and the existing pass says in as many words *"GRANTS, never
removes"*. So it is **sequenced and conditional**:

1. Grant every qualifying group first, on both lists.
2. **Only if that pass completed with no failures**, remove the site-entry assignment — and **name it
   in the log**.
3. A single failed grant leaves the site-entry grant **in place** and says so.

Rationale: the cost of a wrong removal is *nobody can raise a request*, discovered by an uploader; the
cost of a delayed removal is *the exposure persists one more run*, discovered by nobody. Fail toward
the second. An over-wide list that still works beats a tight one that refuses everyone.

Site **Owners** is never a removal candidate, and user principals never are — consistent with
`groupsToRemove`.

### Full assertion is safe at page scope and NOT here

A list root is **not** a leaf. Unlike a Site Pages item it can carry SharePoint's automatic
**Limited Access** entries. So this pass removes **only** the one site-entry assignment it knows
about, by principal id — it does **not** assert a full ACL. Lifting `groupsToRemove`'s
full-assertion behaviour to a list scope is how every group's access gets stripped on a run that
reports success.

### Cost

~470 groups times 2 lists, so roughly 940 role assignments on the **first** run. Later runs skip what
is already in place (1.0.177.0), so the steady state is reads and no writes. Reported once per list,
not once per group — a line each would bury the run's real findings.

---

## 6. Traps, all previously paid for

1. **`LIST_SUFFIX.submissions` MUST be added to `PRIMED_SUFFIXES`.** `requests` was left out from the
   day it was written until 2026-08-20 and the failure was **total and silent**: an unprimed suffix
   never enters the name cache, so `cachedListTitle` answers the legacy `DMS <suffix>` for ever, which
   404s on a CRS site. **Self-enforcing** — `naming.test.ts` already pins both directions, so adding
   the suffix without priming it fails the build.
2. **The write must NEVER fail the upload.** Own try/catch; no row means degrade to today's behaviour.
   Same rule as the `SubmissionId`/`BatchId` stamp, which is conditional for exactly this reason.
3. **Provisioning.** An uploader cannot create a list, so an upload before provisioning must write
   nothing, **silently**. **Deviation from the agreed note, stated:** the note says *"follow the
   `CRS Requests` pattern (created when an admin opens the page)"*. There is **no admin page that owns
   this list** — the page that reads it is My Submissions, which every uploader opens, so a Provision
   button there would be shown to the people who cannot press it. **Reconciliation creates it**
   instead: admin-only, idempotent, and already the thing that asserts the reference columns on four
   libraries and the requests-list ACL. One fewer manual step at SDG cutover.
4. **Match on the stamped reference, never name or path** — both change. And never on `UniqueId`
   either, per §2.
5. **Text columns, never Choice.** A value absent from a Choice column's `Choices` fails the whole
   write, so the day a new value is added every row carrying it is lost silently. Same rule as
   `EventType`, `RequestType`, `Status` and `Stage`.
6. **Dates are ISO on a plain `/items` POST**, not gotcha #1's `M/D/YYYY` — that locale format belongs
   to `validateUpdateListItem`. A plain OData write answers a locale string with *"Cannot convert a
   primitive value to the expected type 'Edm.DateTime'"*.
7. **JSON light, no `__metadata`, both header halves `nometadata`, plus `odata-version: ""`.** SPFx
   injects `4.0`, under which SharePoint cannot infer the entity set for a JSON-light entry payload.
   Learned on the audit log, re-learned on the file-type page (1.0.249.0).

---

## 7. The list

`<prefix> Submissions`, `BaseTemplate: 100`. Internal names space-free, as with the audit log and the
requests list.

| Column | FieldTypeKind | Meaning |
|---|---|---|
| `SubmissionRef` | 2 Text | One per press of Upload. Matches `SubmissionId` on the document. |
| `BatchRef` | 2 Text | One per destination within the submission. Matches `BatchId`. |
| `SubmissionFileId` | 2 Text | **The join key.** One per file. Matches `SubmissionFileId`. |
| `ItemUniqueId` | 2 Text | Secondary only — never decides Deleted. |
| `FileName` | 2 Text | As uploaded, after the naming convention was applied. |
| `ItemPath` | 3 Note | Server-relative path at upload time. |
| `LibraryTitle` | 2 Text | Which library it was filed into. |
| `UploadedBy` | 2 Text | Lower-cased email of the uploader. |
| `UploadedAt` | 4 DateTime | **ISO**, per trap 6. |
| `MetadataSnapshot` | 3 Note | JSON of the tier and metadata values as filled in. |
| `Source` | 2 Text | `Form` or `BulkUpload` — which screen wrote the row. |

No status column, per §3. No `NumberOfLines` on the Note columns — it belongs to
`SP.FieldMultiLineText`, not `SP.Field`, and with `odata=nometadata` the extra property is rejected
outright with HTTP 400. That is the exact point every requests-list provisioning run died on
2026-08-20.

---

## 8. Build order — ALL SIX BUILT (1.0.305.0)

1. ✅ `LIST_SUFFIX.submissions` + `PRIMED_SUFFIXES` (the existing test proves it).
2. ✅ `SUBMISSION_FILE_COLUMN` in `optionalColumns.ts`; reconciliation asserts it on every library.
3. ✅ `shared/submissionRecords.ts` — pure: row shape, payload/parse, the live join, the Deleted
   verdict. 50 tests. Plus `groupsForRequestLists` in `pageGrants.ts` (9 tests).
4. ✅ Reconciliation: creates the list, asserts its columns, **indexes `Author`**, and the two-list
   ACL pass with the conditional revoke.
5. ✅ The two writers — `Form.tsx` and `BulkUpload.tsx`, via `shared/spSubmissionRecords.ts`.
6. ✅ My Submissions: reads the records, merges, renders Deleted.

`tsc` clean, **1452 tests, 0 failures**. Not yet run on a site.

### Two things found while building, both of which would have broken the feature silently

- **⚠ `readLibrary` DID NOT ASK FOR `SubmissionFileId`.** The live rows therefore carried no stamp,
  `indexLiveByFileId` was empty, and **every record would have been judged DELETED** — the exact false
  positive §2 exists to prevent, arriving by a route §2 did not cover. The read now has a **three-rung
  ladder** (all three refs → the pair → none) so a library missing only the newer column keeps its
  grouping, and `stampMissingRef` tells the merge the stamp was unavailable so it answers *"not
  checked"* rather than *"deleted"*.
- **⚠ A SINGLE `liveReadComplete` BOOLEAN WOULD HAVE DISABLED THE FEATURE FOR MOST USERS.** My
  Submissions reads up to six libraries and deliberately swallows failures on the HC pair and the
  archive — an unreadable HC library is the normal case for everyone without clearance. One global
  flag would be false for those users, so nothing would ever be reported as deleted. `mergeRecords`
  now takes a **predicate**, and the caller answers per chain: a normal-library record needs the
  normal chain, an HC one needs the HC chain. A throwing predicate reads as not-complete.

### Where a gone file appears, and where it must not

| View | Gone rows | Why |
|---|---|---|
| **Submissions** (grouped) | **yes** | the client's actual requirement — the file still inside its batch |
| **All** | **yes** | the complete history |
| Pending / Approved / Rejected | **no** | `filterByTab` matches `r.status`, and a record row's `Pending` is a placeholder — listing it would say a destroyed file awaits approval |
| status tallies, batch counts | **no** | `liveRowsOnly`; a gone file has no approval outcome |

A gone row has **no** name button, no `Open file`, no preview and **no request buttons** — a share
request would be approved by a Head of Unit and then fail in their own session, after the requester
was told it was being handled. Its details come from the record's **snapshot**, not from
`FieldValuesAsText`, which is a per-item endpoint that would answer nothing; no per-item read is even
attempted for one.

### Still to do

**Test plan: `docs/2026-08-27-recent-features-test-plan.md`** — the ordered, tickable version of this
list, plus the 1.0.304.0 handoff items and a verified log. The one test that proves §2 is its **§2.3**:
an approved file must still read `Approved`, never `Deleted`.

- **Run it on a site.** Reconcile first (the list, the columns, the `Author` index and the ACL pass all
  come from a run), then upload, then delete a file and confirm the row greys out rather than vanishing.
- **The ACL pass has never granted on a live site** — on ClarenceDMSTesting the requests grant already
  exists by hand, so that path takes the `already correct` branch. **SDG will exercise the grant and
  the revoke**; watch those two log lines on its first run.
- **`MySubmissions.tsx` is now 2161 lines**, over the 2000 max-lines ceiling — a new warning, and the
  fifth file in the project to exceed it. Extracting part of it deserves its own change rather than a
  tail-end tidy inside a feature.
