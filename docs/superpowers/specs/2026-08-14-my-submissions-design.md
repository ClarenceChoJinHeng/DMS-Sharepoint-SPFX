# My Submissions — an uploader's view of their own files

**Date:** 2026-08-14
**Status:** Agreed, not yet built
**Depends on:** `2026-08-08-auto-route-flow-and-draft-isolation.md` (what an uploader can already see)
**Related:** `docs/client/document-visibility-within-a-unit.md` (why "only their own" has a limit),
`2026-07-31-approval-queue-navigation-design.md` (the APPROVER's equivalent)

---

## 1. Why

The client, on behalf of the uploaders: it is difficult to track what you uploaded. Once a file
leaves the upload form it disappears from the uploader's view — approval happens elsewhere, the
approved copy moves to another library, and a rejection is visible only if somebody emails about it.
No screen answers *"what happened to the thing I sent?"*

This page is that screen. **View only** (client, verbatim: *"its mainly for them to view and see
what is going on with their file currently"*) and reachable **only by uploaders** (verbatim: *"its
only for people who are uploaders who can access into this page"*).

---

## 2. The one thing the client must understand

They asked to *"ensure the uploaders only see their own files"*. Half of that is already true and
enforced; half of it cannot be true, and this page must not pretend otherwise.

| Status | Where the file is | Can a peer uploader see it? |
|---|---|---|
| Pending | `Approval Document` | **No** — enforced by Draft Item Security (approver + author) |
| Rejected | `Approval Document` | **No** — same mechanism |
| Approved | `Documents` | **Yes.** Every PIC reads every approved document in their unit |

Pending and rejected privacy is real and SharePoint enforces it. This page neither creates it nor
can lose it.

Approved files are different. Auto-route MOVES them to `Documents` and deletes the source, and in
`Documents` a unit's approved documents are readable by that whole unit — deliberately, since the
one-group-per-person model of 2026-08-09. So the approved list here is **filtered** to the signed-in
user, and that filter is a convenience, **not a boundary**: the same person can open `Documents` and
see a colleague's approved files.

**The page therefore never claims to hide approved documents.** Making them private per person means
filing per person instead of per unit — the trade already set out for the client in
`docs/client/document-visibility-within-a-unit.md`.

---

## 3. Decisions

| # | Decision | Chosen |
|---|---|---|
| D1 | Scope of action | **View only** — no delete, no resubmit, no edit |
| D2 | Whose files | The signed-in user's own uploads (`Created By = me`). No "my unit" toggle |
| D3 | Access | Page Access, **`UPL` only**, via a dedicated rule in `pageAccessPolicy.ts` |
| D4 | Sources | Both libraries — pending/rejected from `Approval Document`, approved from `Documents` |
| D5 | Rejection comment | Shown when readable; the page degrades rather than fails without it |
| D6 | Unreadable ≠ empty | Three distinct empty states |

### D3 — why a dedicated policy rule

`policyForPage` matches on file name, **first match wins, and the order is load-bearing** — the
`bulk` rule sits before `upload` precisely so a bulk page is never offered to uploaders.

A page called `My-Submissions.aspx` matches no existing rule and would take `DEFAULT_POLICY` —
`UPL`, `APR` and `DELS` — offering it to approvers and Staging-deleters. Naming it `My-Uploads.aspx`
would match `/upload/i` and land on `UPL` by luck, with a reason line reading "this page is where
documents are submitted", which is untrue of a tracking page.

So: an explicit rule, `/submission|my.?upload|my.?file/i` → `["UPL"]`, placed **before** the generic
`upload` rule, with its own reason. Pinned by a test, mirroring the bulk/upload ordering it copies.

The grant itself is made on the Page Access screen, exactly as `upload-form.aspx` is today. This
module is a UI filter, never the boundary — the boundary is the SharePoint grant.

### D4 — two libraries, because a file's life spans two

A file is Pending, then either Rejected (stays put) or Approved (moves). Reading one library would
show half a lifecycle. Both reads filter on `AuthorId eq <me>`; Auto-route preserves the uploader in
`Author` (verified 2026-08-08), which is what makes the second half joinable at all.

Everything in `Documents` is approved by definition — it got there by being approved, and that
library has content approval OFF — so those rows need no status read.

### D5 — the rejection comment

The most useful thing on this page: without it, "why was it rejected" becomes an email. It lives in
`OData__ModerationComments`.

A `$select` naming a column that does not exist fails the **whole request** with HTTP 400 — not a
null, not a missing key (the same trap as `Scope`/`Target` on the Group Map). So the read asks for
it and **falls back to the field set without it** when rejected, rather than blanking the page. When
the comment cannot be read the row still appears and says the reason was not recorded — which is
also the honest answer when an approver rejected without typing one.

---

## 4. The page

**`CRS My Submissions`**, its own page and web part. No admin gate in code: the gate is the page
grant. An admin who opens it sees their own uploads, which is correct and occasionally useful.

- **Tabs: All / Pending / Approved / Rejected**, each with a count.
- One row per file: **name** (opens the file), **status**, **when uploaded**, **where it went**, and
  **document type**.
- The path is shown as a folder trail — `NBPOLHO › CDS › UPSUPPORT › 2024 › Tax Return` — derived
  from `FileRef` with the site and library prefix stripped. An uploader thinks in folders, not URLs.
- A **rejected** row shows the approver's comment where there is one.
- Dates read `DD/MMM/YYYY`, the agreed client format (gotcha #1).
- Sorted newest first: the question being asked is *what happened to the thing I just sent*, and the
  answer is almost always at the top.

### 4.1 Status

From `OData__ModerationStatus` in the approval library: `0` Approved, `1` Rejected, `2` Pending,
`3` Draft. **Draft is shown as Pending** — it is not a state an uploader can act on, and a fourth
badge nobody can explain is worse than a slightly coarse one. An unrecognised value is shown as
itself rather than dropped, so a row never silently disappears.

### 4.2 The three empty states (D6)

1. **"You have not uploaded anything yet."** — both reads succeeded, nothing matched.
2. **"Nothing in this tab."** — reads succeeded; other tabs have rows.
3. **"Could not read your submissions."** — a read failed, naming which library. **Never** shown as
   "you have no files": an uploader told they have nothing when the library was merely unreachable
   will upload it again, and now there are two.

---

## 5. Risks

| Risk | Handling |
|---|---|
| Client believes the page makes approved files private | §2 — in this spec, and stated on the page |
| `Documents` passes 5,000 items and the `AuthorId` filter starts failing | Index `Created By` on `Documents`; on the provisioning checklist. Surfaces as empty state 3, never as silence |
| `OData__ModerationComments` unreadable on a site | D5's fallback — the page loses one column, not its content |
| Page granted to approvers by accident | D3's explicit rule, pinned by a test |
| An uploader in several units sees one long list | Newest first, and the path column names the unit on every row |

---

## 6. Test plan

Pure logic (`src/shared/`):

- the page policy: `My-Submissions.aspx` → `["UPL"]`; `APR`, `DELS` and view-only roles refused
- rule ORDER: the new rule is reached before the generic `upload` rule
- status mapping, including Draft → Pending, and an unknown value shown rather than dropped
- the folder trail: site and library prefix stripped, file name removed, separators normalised

Site verification:

1. As an uploader with a pending file — it appears under Pending, and opens
2. Approve it — it shows as Approved on the next load, still visible to that uploader
3. Reject one with a comment — appears under Rejected, comment shown
4. Reject one **without** a comment — the row still appears, saying none was recorded
5. As a *different* uploader in the same unit — neither of those pending/rejected files appears
6. The count on each tab matches its rows
7. A non-uploader cannot open the page at all (Page Access)

---

## 7. Implementation order

1. The page-policy rule + tests
2. `shared/mySubmissions.ts` — status mapping and the folder trail + tests
3. The web part: the two reads, the tabs, the rows
4. Register the bundle and componentId, bump the version
5. Create the page, grant it to `UPL` groups on Page Access, run §6
