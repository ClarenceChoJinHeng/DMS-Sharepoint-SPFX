# Requests page — categories, decision history, and share revocation

**Date:** 2026-08-28
**Status:** design agreed, not built
**Supersedes:** nothing. **Completes** `2026-08-15-deletion-and-share-requests-design.md` §5.3, which
specified a Revoke action and left it unbuilt (*"Without a way back out, the count only ever grows"*).

Client, on the current page: *"can we redesign this a bit? Can we separate into like Share Category,
Deletion Category and inside have pending approval and, similar design as the My Submission? So we can
track who did what as well? Also a tab to show which user is sharing which one? And then allow HOU to
revoke their sharing permits for that file?"*

---

## 0. One correction the client's question rests on

They asked: *"only HOU sees their own unit request from PIC right?"* — **not quite.** From `inScope`
(`shared/requests.ts`), there are **four** audiences, and the redesign must keep all four working:

| Who | Sees | Stages |
|---|---|---|
| **Head of Unit** (`APR` / `APRHC` mapping) | their own units | pending **and** approved |
| **Head of Department** (`DEPTVIEW`) | every unit under their department | **approved only** |
| **System administrator** (`CRS Owners` / SCA) | every request on the site | both |
| **The requester** | their own rows, always | both |

⚠ **The HoD stage limit is not a policy choice** — a HoD holds nothing in either approval library, so
they physically cannot action a pending-file request. Do not "fix" it by widening the tab filter.

---

## 1. Page shape

Four tabs, counted in the label, the same pattern as My Submissions' `tabCounts`:

```
[ Deletion (3) ]  [ Share (2) ]  [ Shared files (7) ]  [ Your requests (5) ]
```

**Deletion** and **Share** each hold two sections:

- **Waiting for you (N)** — Approve / Reject, exactly as today.
- **Decided (N)** — closed rows, newest first, each stating **who decided it, when, and what they
  wrote**. This is the *"track who did what"* half, and it is currently missing entirely: a decided
  request vanishes from the page unless the viewer happens to be the requester.

**Your requests** keeps today's behaviour unchanged.

**Shared files** is new — see §3.

### 1.1 Why tabs and not one longer page

The client asked for the My Submissions shape, and the reason it works there applies here: the page
has two axes (type × status) and a flat list can only express one. Today's page expresses neither —
it shows *pending for me* and *raised by me*, so a Head of Unit cannot answer "what shares are open
in my unit" without reading every row.

### 1.2 What does NOT change

- `isVisibleTo` / `canDecide` / `inScope` — the visibility rule is untouched. Every tab filters the
  same `visible` array the tally already uses, so **the tabs and the tally cannot disagree**
  (the 2026-08-21 defect, where a HoD read "2 pending" beneath a queue of 1).
- The decision flow, its re-read-before-acting guard, and the confirmation dialog.
- The list schema. **No new column.** See §4.3.

---

## 2. Revoking a share

### 2.1 It needs no new permission — this is the good news

`CRS Share` **already contains Manage Permissions**. That is the entire reason it exists as a
separate level rather than reusing Contribute (`FolderManager.tsx` ~341, which says so and warns to
verify it on the site rather than trusting the name). Manage Permissions is exactly the right a
revoke requires.

So: **no new role, no new permission level, no Group Map change, no reconciliation run.** Every
persona that can approve a share can already perform its reversal.

⚠ **Verify on the live site anyway**, per the standing warning on that comment. A `CRS Share` level
built without Manage Permissions leaves *approvals* failing at the last step too, so this is a
pre-existing check, not a new one.

### 2.2 The mechanism

An approved share breaks inheritance on the **file** and adds the recipient's role assignment.
Two ways back, and **both are offered** (client's choice):

| Action | Call | Effect |
|---|---|---|
| **Revoke one recipient** | `…/roleassignments/removeroleassignment(principalid=N)` | drops that person; the file keeps its unique scope |
| **Revoke all** | `…/ListItemAllFields/resetroleinheritance` | drops every share **and reclaims the scope** |

⚠ **Only the second reclaims the permission scope**, and that is the half that matters: every
approved share creates a unique scope, and it is the same 50,000-per-list ceiling that killed
per-uploader ACLs on 2026-08-06. Per-recipient revocation alone leaves the count monotonic — which is
precisely the problem §5.3 was written to solve. So *"Revoke all"* is not a convenience button; it is
the mechanism.

⚠ **`resetroleinheritance` returns the file to its UNIT FOLDER's ACL**, which is the correct resting
state — not to the library root. Nothing is lost: the unit's own groups reach it exactly as they did
before the share.

### 2.3 Who may press it

**The same rule that decides the request** — `inScope`, unchanged, reused. HoU for their units, HoD
for their department's approved files, system admin everywhere.

All three already hold the right: `hou` and `hod` carry `SHARE`, `hou_hc` and `hod` carry `SHAREHC`.
Reusing `inScope` rather than writing a second predicate keeps the invariant the file already
protects: **a row they cannot act on is a row they do not see.**

⚠ **NOT the original requester.** A PIC holds no Manage Permissions, so offering them the button
would fail in their session *after* the screen had offered it — the "visible but inert" state this
codebase already rejected for pending-stage HoD requests.

---

## 3. The "Shared files" tab

### 3.1 Rows index; the live ACL decides

The request rows and the real permissions **can disagree, in both directions**:

- A share performed through **SharePoint's own Share button** leaves **no request row**. A web part
  cannot intercept that native control — settled in `2026-07-23-share-guard-retirement.md` and not
  re-openable. Rows-only **under**-reports.
- A share revoked directly in SharePoint leaves the row still reading `Approved`. Rows-only
  **over**-reports.

On a screen whose whole job is *"who can currently reach this document"*, both are unacceptable. So:

> **Approved `Share` rows are the INDEX. The file's live ACL is the VERDICT.**

The same shape `mergeRecords` already uses in My Submissions, where the submission record indexes and
the library read decides.

### 3.2 ⚠ Three states, never two

`ShareState = "live" | "revoked" | "unknown"`.

`unknown` exists because a throttled or refused ACL read must **never** render as *"revoked"* — that
would tell an approver access had ended when it had not, and they would stop looking. Same rule as
`removalVerdict` on the access screens and `RecordState` in My Submissions. **Empty ≠ unknown**, in
the place where the cost is somebody believing a document is no longer shared.

The rollup is worst-case: a file with any `unknown` recipient is `unknown`, never `live`.

### 3.3 ⚠ What this tab CANNOT see, and says so

It probes files that **have a request row**. A file that was **only ever** shared natively has no row,
so nothing indexes it and it does not appear at all. Finding those means enumerating every file in
`Documents` with unique permissions — unbounded, and a full-library crawl on a page load.

**Stated on the screen**, not buried here: a tab that implies completeness about *who can reach our
documents* is worse than one that admits its edge.

What it **does** catch on an indexed file: extra recipients added natively (rendered as
*"granted outside CRS"*), and recipients revoked outside the app.

### 3.4 Cost

One request per shared file, on opening the tab only. Shares are rare by design — that is the whole
reason per-file ACLs are workable here where they were not for per-uploader isolation. Reads are
issued in parallel and **failure is per file**: one unreadable ACL marks that file `unknown` and
leaves the rest correct.

Group principals are skipped — when inheritance breaks, SharePoint copies the unit folder's groups
onto the file, and listing them as "shared with" would report the unit's own access as a share.

---

## 4. Recording who revoked

### 4.1 ⚠ The code cannot write the audit row, and this is already established

`writeAudit` is **refused for every non-Owner** — `CRS Audit Log` restricts writes to Owners and the
service account by design, which is what made the whole `CRS — Audit request activity` flow necessary
on 2026-08-26. A Head of Unit pressing Revoke is exactly such a non-Owner. Do not add a `writeAudit`
call here expecting it to work; it will silently do nothing, exactly as the request writers did.

### 4.2 So the revoke is recorded on the ROW, and the flow picks it up

- Revoking the **last remaining** recipient sets the row's `Status` to **`Revoked`**.
- Revoking **one of several** appends a line to `DecisionNote` and leaves `Status` as `Approved` —
  because the request's own outcome has not changed; part of its grant has.

Both are MERGEs on `CRS Requests`, so the existing **created-or-modified** audit flow fires with no
new trigger.

### 4.3 No new column

`Status` is a **Text** column, never Choice — deliberately, and stated in the 2026-08-15 design for
exactly this reason: *a value absent from a Choice column's `Choices` fails the whole write,
silently.* So `Revoked` writes with no schema change, no provisioning step, and none of the
read-retry / write-retry dance that the `Stage` column needed.

`RequestStatus` gains `"Revoked"`; `counts` and `PILL` gain an entry. Grey and quiet, like
`Cancelled` — a revocation is the system working, not a failure.

### 4.4 ⚠ FOLLOW-UP REQUIRED: one branch in the audit flow

`CRS — Audit request activity` derives `EventKind` as:

```
if(equals(Status,'Pending'), …, if(equals(Status,'Approved'),'RequestApproved','RequestRejected'))
```

so **until a branch is added, a revoke logs as `RequestRejected`** — wrong, though not dangerous
(the row is written, the actor is right, only the label misleads). Add:

```
if(equals(Status,'Revoked'),'ShareRevoked', …)
```

**Written down here because a flow change cannot live in source control**, which is the standing
failure mode for every flow in this project.

---

## 5. What changes in code

| File | Change |
|---|---|
| `shared/requests.ts` | `ShareState`, `ShareRecipient`, `SharedFile`, `collectSharedFiles`, `mergeShareAcl`, `canRevoke`, `Revoked` status. `matchesUnit`/`inScope` widened structurally so a `SharedFile` can be scope-checked without a cast. |
| `shared/requests.test.ts` | rules above, incl. the three-state verdict and the group-principal skip |
| `webparts/requests/components/Requests.tsx` | four tabs, Decided sections, Shared files tab, ACL probe, revoke calls |
| `CLAUDE.md` | new section |

**No change** to `pageAccessPolicy`, `groupMapModel`, `LIBRARY_ROLES`, reconciliation, or the list
schema.

---

## 6. What to tell the client

1. **A revoke is immediate and silent to the recipient.** SharePoint sends no "access removed" email.
   If they had the link, it now returns Access Denied with no explanation.
2. **"Revoke all" is the one that matters.** Per-person revocation tidies the list; only revoke-all
   reclaims the permission scope that the 50,000 ceiling counts.
3. **Files shared through SharePoint's own Share button are not fully visible here** (§3.3), and
   cannot be — a web part cannot intercept that control. The tab shows what CRS granted, plus
   anything extra found on those same files.
4. **A Head of Department can revoke too**, within their department's approved documents — they hold
   `SHARE` under the 2026-08-20 role revision.
5. Until the audit flow gains its branch (§4.4), a revocation appears in the log as a rejection.
