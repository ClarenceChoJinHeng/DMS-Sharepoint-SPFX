# CRS — The seven-year archive mover · Power Automate build runbook

**Date:** 2026-08-22
**Design:** `2026-08-22-seven-year-archive-design.md` — read §6 before building.
**Status:** not built.

[OK] **Section 1.5 IS ANSWERED FOR `UniqueId` - VERIFIED ON SITE 2026-08-22.** A cross-library move
PRESERVES it, so the flow uses plain `MoveTo` and needs no GUID-mapping action. The share half is
still open and no longer gates anything - see 4.5.

---

## 0. Sign in as the SERVICE ACCOUNT first

A flow runs under its connection, and the connection is created implicitly by the **first action** —
so whoever is signed in when you place that action is baked in for life. Built as a person, this flow
stops **silently** when their password changes, and presents months later as "why did nothing archive
this year".

On SDG's tenant this must be right from the first action. On ClarenceDMSTesting it is knowingly not.

## 1. Prerequisites

| Thing | State | How to check |
|---|---|---|
| `Archive` library | exists, content approval **OFF** | Library settings → Versioning |
| `HC Archive` library | exists, content approval **OFF** | as above |
| Both trees built, all groups **Read** | reconciliation has run on 1.0.231.0+ | the run log shows six panels |
| `Archived` column | on all six libraries | reconciliation asserts it every run |
| App version on the site | 1.0.231.0 or later | **Site Contents → the app → ⋯ → Details** — never the catalog |

⚠ **Content approval OFF is load-bearing.** With it on, every folder this flow creates and every file
it moves arrives **Pending** and is invisible to every reader. It presents as a permissions bug and is
not one.

## 1.5 What was verified, and what is still open

**VERIFIED ON SITE, 2026-08-22.** `Shared Documents/GHO/GF/TAX/2024/Term Sheet/Archive 2/TEST - TEST -
Non Highly COnfidential - 19-08-26.pdf` was moved by hand into `Archive/GHO/GF/TAX`:

| | Before | After |
|---|---|---|
| `UniqueId` | `accdbebf-...a992` | **`accdbebf-...a992` - unchanged** |
| list item id | 6095 | 142 - changed, as expected (**item ids are per-LIST**) |
| `Modified` | original | unchanged, not restamped |

**So a cross-library `MoveTo` preserves file identity.** Three consequences, all good:

- the flow uses plain `MoveTo`; the copy-stamp-delete shape Auto-route uses is NOT needed here
- **the audit trail stays continuous** - the audit log keys on `ItemUniqueId`, and a copy would have
  broken the join at the archive boundary. Auto-route already breaks it at the APPROVAL boundary for
  exactly that reason; do not copy that pattern here
- an existing deletion or share request row still resolves after its file archives

**Still open: does an individual SHARE survive the move?** It no longer gates anything - 4.5 resets the
moved file to inherit **unconditionally**, which is a no-op if shares do not survive and is the correct
behaviour if they do.

## 2. Create the flow — `CRS — Archive after seven years`

**Scheduled cloud flow**, not automated. Nothing about a file changes on its seventh birthday, so
there is no trigger to hang this on.

### 2.1 Trigger

**Recurrence** — Frequency `Week`, Interval `1`. Pick a time outside working hours.

⚠ **Weekly, not daily.** A seven-year boundary does not need daily precision, each run is a full
library scan, and every run competes for the same daily flow quota as Auto-route.

### 2.2 Work out the cut-off

**Compose**, named `CutOff`:

```
@{addDays(utcNow(), -2557, 'yyyy-MM-ddTHH:mm:ssZ')}
```

`2557` = 7 × 365 + 2 leap days. **ISO, because this feeds an OData `$filter`** — gotcha #1's
`M/D/YYYY` belongs to `validateUpdateListItem` and is rejected here.

### 2.3 Find the eligible files

**Get items** (not Get files) on `Documents`:

- Filter Query: `Created lt datetime'@{outputs('CutOff')}' and FSObjType eq 0`
- Top Count: **100**
- Order By: `Created`

⚠ **`FSObjType`, not `FileSystemObjectType`** — the REST spelling is rejected inside a `$filter`.
Without it the flow tries to "archive" every folder.

⚠ **`Created` must be indexed on `Documents`**, or past 5,000 items this filter throws the list view
threshold and the run fails rather than returning fewer rows. Library settings → Indexed columns.

⚠ **THE CAP IS DELIBERATE.** 100 files per week, oldest first. A first sweep on a real archive could
be thousands of files, and an exhausted daily quota means **the next real approval is not routed,
silently**. The backlog drains over successive runs; nothing is lost by going slowly.

### 2.4 Apply to each

Over `value` from step 2.3. Set **Concurrency Control → OFF** (sequential): parallel moves into the
same destination folder race on folder creation.

Inside the loop:

#### a. Derive the destination path

**Compose**, `DestPath`:

```
@{replace(items('Apply_to_each')?['{FullPath}'], '/Shared Documents/', '/Archive/')}
```

⚠ **`Shared Documents`, NOT `Documents`.** A list's title and its URL are independent (gotcha #12),
and this library is the standing example: titled `Documents`, addressed as `Shared Documents`. Using
the title here matches nothing and every file silently stays put.

#### b. Ensure the destination folder

**Create new folder** on `Archive`, folder path = the parent of `DestPath`. **Configure run after →
also on `has failed`**: it fails when the folder already exists, which is the normal case.

⚠ **THIS IS SAFE ONLY BELOW THE UNIT.** A **unit** folder this flow creates would inherit the
department folder's ACL instead of carrying the unit's — which is how a unit's documents become
readable by its whole department. Reconciliation builds every folder down to the unit; if a unit
folder is missing, reconciliation has not been run and **the answer is to run it, not to let the flow
improvise**. Add §4.3 if you want that asserted rather than assumed.

#### c. Handle a name clash BEFORE the move

**Get file metadata using path** on `Archive` at `DestPath`, with **run after → also on `has
failed`**.

- Failed (404) ⇒ no clash, move as `DestPath`
- Succeeded ⇒ a file of that name is already archived. Move to
  `<name> (@{formatDateTime(items('Apply_to_each')?['Created'], 'yyyy')})<ext>` instead.

**The suffix is the creation YEAR, never a number.** `(1)` and `(2)` throw away the only thing that
distinguished the two documents. **One file keeps the original name** — the one already there.

If the client rejects renaming (§2.2 of the design), replace this branch with: **move nothing, send
the administrator an email naming the file, both paths and the clash.** A silent skip is not an
option — the file stays in `Documents` looking as though it was never eligible.

#### d. Move the file

**Send an HTTP request to SharePoint** — the connector's own Move action re-creates the file.

- Method: `POST`
- Uri:
  `_api/web/GetFileByServerRelativeUrl(@f)/moveto(newurl=@d,flags=1)?@f='@{encodeUriComponent(items('Apply_to_each')?['{FullPath}'])}'&@d='@{encodeUriComponent(outputs('DestPath'))}'`
- Headers: `Accept` → `application/json;odata=nometadata`

⚠ **THE PARAMETER ALIAS FORM IS REQUIRED, NOT STYLISTIC.** A long encoded path passed as an inline
quoted literal returns **HTTP 400, not 404**, at around 330 characters — well under SharePoint's item
path limit — and `Form.tsx` failed exactly this way on nested paths for weeks (gotcha #9).

⚠ **HEADER KEYS HAVE NO COLON.** `Accept`, never `Accept:`. With the colon the header does not exist,
and the symptoms look unrelated to each other. This cost two hours on 2026-08-08.

⚠ **`flags=1` overwrites at the destination.** Only reachable when §c has already renamed away a
clash, and it is what makes a re-run of a half-finished move safe.

#### e. Stamp `Archived`

**Send an HTTP request to SharePoint**, against the **`Archive`** library:

- Method: `POST`
- Uri: `_api/web/lists/getbytitle('Archive')/items(<new id>)/validateUpdateListItem`
- Headers: `Accept` and `Content-Type` both `application/json;odata=nometadata`
- Body: `{"formValues":[{"FieldName":"Archived","FieldValue":"true"}]}`

⚠ **ITEM IDS ARE PER-LIST.** The id from `Documents` is meaningless in `Archive`; read the new item's
id back from the move response or by path. Stamping the wrong list is a 404 on a good day and, on a
bad one, **writes onto a different document** — the 1.0.170.0 defect.

⚠ **HTTP 200 does not mean it worked.** `validateUpdateListItem` returns 200 with per-field errors in
the body — check `HasException` on each result (gotcha #4).

⚠ The badge on My Submissions does **not** depend on this: `isArchivedRow` derives the tag from the
LIBRARY. The column is for views and search. A failed stamp is worth reporting, not worth failing the
run over — the file is already correctly filed and correctly permissioned.

## 3. HC — a SECOND flow, cloned and every reference changed

`Documents` → `Archive` and `HC Documents` → `HC Archive`, **never crossing**.

⚠ **THE HC AUTO-ROUTE CLONE PRODUCED SIX FAULTS OF EXACTLY THIS SHAPE ON 2026-08-19**, and two of
them would not have failed loudly: a stamp pointed at the wrong list writes onto a different
document, and a **delete** pointed at the wrong list removes whatever holds that id there. Change
every library reference and read the flow back before saving:

- the **Get items** list
- `'/Shared Documents/'` → `'/HC Documents/'` and `'/Archive/'` → `'/HC Archive/'` in `DestPath`
- the **Create new folder** library
- the **getbytitle** in the stamp

## 4. Two more actions in the loop

### 4.3 Assert the unit folder exists (recommended)

Before ensure-creating below it, confirm the unit folder is there and skip the file with a NAMED error
if not. Turns an invisible ACL mistake into a visible skip - see 2.4b for why.

### 4.5 Reset the moved file to inherit - DECIDED, not conditional

After the move, on the archived file:

`_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields/resetroleinheritance?@f='...'`

**Unconditional, and that is what removes the last unanswered question.** If a per-file share does not
survive a cross-library move, this is a harmless no-op. If it does survive, this is what makes the
client's own rule - *"for all the Archive files make sure all groups have read only"* - actually true.

**The asymmetry that decides it: there is no revoke feature.** It was specified and never built, so a
share made in 2026 would still be granting one person access to an archived record in 2033, on a library
where every group is deliberately Read, with nothing anywhere listing it. A wrongly-kept share is
permanent and invisible; a wrongly-removed one is one request away from being re-granted.

Inheriting is the SAFE state here, unlike almost everywhere else in this system: the file inherits its
unit folder, whose ACL reconciliation has just asserted as Read for every mapped group.

**Tell the client: archiving a document ends any individual shares on it.**

## 5. Test it, in this order

**Do not wait until 2033, and do NOT shorten the cut-off to do it.**

### 5.1 The safe lever is the FILTER, not the schedule and not the date

A scheduled flow can always be fired with the **Run** button, so the recurrence interval is irrelevant
to testing. Leave it weekly.

For the run itself, temporarily replace the Filter Query in 2.3 with a FILENAME:

```
FSObjType eq 0 and FileLeafRef eq '<one known test file>'
```

That exercises the whole mechanism - path derivation, folder ensure-creation, clash check, move,
`Archived` stamp, inherit reset - against **exactly one file**, with no possibility of touching
anything else. Put the date filter back afterwards.

**DO NOT test by changing `-2557` to a small number.** It is the obvious shortcut and it is the
dangerous one: every file in `Documents` qualifies at once, and with the 100-file cap a single run
archives **100 real documents**. Undoing that means moving 100 files back by hand and clearing their
`Archived` flags one at a time.

### 5.2 The date arithmetic is proved separately, with no files at all

Run the flow and read the **`CutOff` Compose output**. It should be about seven years ago to the day.
If it is, the expression is right and nothing else about it can be wrong.

### 5.3 The order

1. Filename filter, one file, **Run**. It should land in `Archive` at the mirrored path with metadata,
   `Created By` and `Created` intact, and `Archived` ticked.
2. Open **My Submissions** as the uploader: the row shows an `Archived` badge, and the deletion and
   share buttons are **refused** with the read-only-record message.
3. Put a second file with the SAME NAME in the same destination and re-run: confirm the rename
   (`report (2018).pdf`) and that the first file is untouched.
4. **CRS Search** finds the archived file. Allow for crawl latency - a just-moved file may take time to
   appear, and that is not a bug in the flow.
5. Sign in as a PIC: the archived file is visible and **read-only**.
6. Repeat 1 and 5 on the HC pair, and confirm an **uncleared** uploader sees nothing.
7. Restore the date filter. Run once more: it must report zero moves.
8. **Move the test files back to `Documents`** and clear their `Archived` flag. They are real documents.

### 5.4 If you ever do need a genuinely old file

Backdate `Created` with `validateUpdateListItem` and `bNewDocumentUpdate: true`. That endpoint wants
`M/d/yyyy h:mm tt`, **not** ISO, and not the `$filter` format from 2.2 - two date formats in one
feature, and they are not interchangeable.

## 6. What this flow must NOT be extended to do

- **Deleting anything.** This moves documents; disposal is a separate decision with separate approval.
- **Archiving from either approval library.** A document that was never approved is not a record.
- **Un-archiving.** Moving a file back is a deliberate manual act.
- **Running daily, or without the 100-file cap.** Both trade a quota that a real approval needs.

## 7. Turning it off

Turn the flow **off**, do not delete it — the build is not in source control and this document is the
only record of it. Nothing else changes: the libraries, the folders and the grants stay, and files
already archived stay where they are and stay readable.
