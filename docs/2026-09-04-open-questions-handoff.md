# Open questions and unfinished work — as at 2026-09-04, build 1.0.411.0

Written at the client's request before a context compaction: *"anything you want to ask me later on or
any question from previous task save it now because i dont [want] you to forget."*

⚠ **NOTHING FROM 2026-09-04 HAS BEEN TESTED ON A SITE.** Every check that day was a compile, a unit
test, or a grep of the shipped bundle. This project has no UI tests, so all of `1.0.397.0` →
`1.0.411.0` is unverified in a browser. Deploy and click through before treating any of it as done.

⚠ **NOTHING IS COMMITTED.** The working tree holds every change from the last several sessions,
including the client's own hand edits to `Form.tsx` and `GroupManager.tsx`.

---

## 1. Questions for the client — answers change what gets built

**Answered 2026-09-04 — 1, 2, 3 and 6 are closed and built (`1.0.412.0`). 4 and 5 were not
understood as asked and are restated below; they are cosmetic and nothing is blocked on them.**

| # | Question | Answer | State |
|---|---|---|---|
| 1 | **Audit Log: show a real total?** | *"Add it, just follow what the mockup would want."* | **BUILT.** `countAudit` in `spAuditLog.ts` — a counted walk, NOT `ItemCount`. See the note below. |
| 2 | **"My Request — 3 per page": which list?** | *"The Approver page show 3 list card per dropdown as well."* Plus, unprompted: switching tab must reset paging to page 1. | **BUILT.** Both. |
| 3 | **Pending stripe now pale amber** | *"Keep keep"* | Closed, no change. |
| 4 | **Approve is BLACK in the dialog and GREEN on the cards.** Should they match? | *"leave it, colour is not a complain for the buttons."* | Closed. The mismatch is deliberate and known — do not "tidy" it. |
| 5 | **The Reject dialog shows the filename**, which their mockup omitted. Keep it? | Not answered — they replied about the sign-out caveat instead, and it was not re-raised. | **Kept**, which is the safe default: an approver refusing a request with nothing naming the document is how the wrong one gets refused. Not worth a second ask. |
| 6 | **Requests page padding** — shared page shell, or their `20px 80px`? | *"the same as upload form"* | Closed, already the shell. |

⚠ **THE AUDIT TOTAL IS COUNTED, AND MUST NEVER BE SWITCHED TO `ItemCount`.** That aggregate is
cached and lags both ways — it read **2,419 for a list whose view was empty** (2026-08-24) — and it
cannot answer a FILTERED question at all. `countAudit` walks `$select=Id` pages instead: one extra
request per load for any result under 5,000. Three display states, and they must stay apart: a
counted total; a **floor** (`25,000+`) when the walk hits its five-page cap, because a floor shown
as a total is the one thing this screen must not do; and the old *"and more"* when the count could
not be made at all. Fails open — a failed count never takes the log off the screen.

## 2. Things removed at the client's request that may come back

Each is one line to restore, and each removed a fact nothing else on screen states.

- **The sign-out/sign-in caveat** (Group Management, member add). **DECIDED 2026-09-04 — STAYS
  REMOVED, and this is not an oversight to correct.** Client: *"its ok just follow the client's
  feedback, as of now there isn't any issue where I have to sign out or sign in, if that ever happens
  we ask them to log out and login."* So the mitigation is now a SUPPORT ANSWER rather than a line on
  the screen, and the diagnosis is kept here for whoever fields the call.
  - ⚠ **The failure is real and reads as a permissions fault.** On 2026-08-18 a guest was bounced off
    a page their new group demonstrably granted while EVERY permission read back correct — group
    membership, the page role assignment, its `Read` binding, and the user's own `currentuser/groups`
    listing the new group. A full sign-out and sign-in fixed it instantly. **It is the signed-in
    SESSION that is stale, not the permissions**, so no amount of REST checking will find it.
  - **Diagnostic order for *"I added them and they still cannot get in"*:** is the principal the right
    one (a guest can exist TWICE for one address), is the group granted on the page, is the binding
    `Read`, does their own session list the group — **then have them sign out and back in.**
  - Gone with the same paragraph: that adding someone also joins them to the site-entry group, and
    that membership is a GROUP change reaching every mapped folder. Neither is stated anywhere on
    screen now.
- **The `hidden`-clash warning** in the upload form's Replace dialog — *"It may be someone else's, and
  replacing it would discard their file."* Yes now destroys an invisible colleague's pending draft
  without naming whose. Raised with the client; not yet answered.
- **"Searchable from the home page"** under the Keyword field — the only thing saying why it exists.
- **`policy.reason`** on Page Access — an admin page and an approver page now look identical there.
- **The failed-read message** on a group's member list. A failure is now SILENT rather than
  mislabelled (an empty group still says so) — the safe direction, but not an explained one.

## 3. Flow-side work — Power Automate, not code

- ⚠ **`CRS — Audit request activity` logs a CANCELLED request as `RequestRejected`, with no actor.**
  Its `EventKind` is `if(Status='Pending', …, if(Status='Approved','RequestApproved','RequestRejected'))`,
  so **Cancelled falls into the Rejected branch** — a request the requester WITHDREW is recorded as one
  an approver REFUSED. And `ActorEmail` picks `DecidedBy`, which a cancellation has not got (the person
  is in `RequestedBy`), which is the empty **Who** the client reported. Two edits, same shape as the
  `RevokedBy` fix: add a `Cancelled` branch, and prefer `RequestedBy` for it.
  - **The code half is deliberately NOT done yet:** `RequestCancelled` needs adding to `EVENT`,
    `EVENT_LABEL` and `ALL_EVENT_TYPES` in `shared/auditLog.ts` — but only AFTER the flow writes the
    value, or the filter offers an option that matches nothing.
- **Both archive movers write no audit row.** `EventType: "Archived"` is registered in code and
  filterable; neither `CRS — Archive after seven years` nor its HC clone creates the row. Runbook:
  `docs/superpowers/specs/2026-09-02-archive-audit-log-and-access-narrowing-runbook.md`.
- **Both archive movers still carry a TEST cutoff**, not the rolling 7-year one. `CutOff` must go back
  to `addDays(utcNow(), -2557, 'yyyy-MM-ddTHH:mm:ssZ')` before either is trusted on a schedule.

## 4. Deployment steps code cannot do

- ⚠ **Delete the page's own title web part on `Upload-Form.aspx` AND `Bulk-Upload.aspx`**, or the
  heading renders twice. Nothing in code can detect a text web part.
- **Run Folder Reconciliation** before testing Keyword — it creates the column, and until it exists
  both upload screens silently do not write it (by design, gotcha #4).
- ⚠ **Keyword search will only find PENDING documents at first.** The approval libraries go through
  REST and work immediately; the approved side goes through the Search API, which needs a **managed
  property** mapped and a re-index. That is the `<Name>OWSTEXT` assumption verified NOT to hold on this
  tenant (2026-08-23). Fix: Site Settings → Search Schema → map `ows_Keyword` to a `RefinableString`,
  then re-index (asynchronous, hours).
- **SDG still needs:** Site Assets Read for `CRS_SITE_MEMBERS` (or its banner stays a plain green
  rectangle); `tenantDomains = sdguthrie.com` in CRS Config; and version history confirmed on both
  approved-side libraries before the Replace behaviour is trusted.

## 5. Consequences to tell the client

- **The "Documents with Shared Access" section is hidden** (`SHOW_SHARED_FILES = false`,
  `Requests.tsx`). With it off there is **no index of CRS shares and no in-app way to revoke one** —
  that now goes through the library's Permissions page → *Show users* → remove the `Type: User` rows.
  ⚠ Never the header checkbox there: every group also shows Limited Access, and removing those strips
  every group's folder access across the whole library.
- **Removing a level does not clear its column**, so folder and metadata can disagree after a
  structure change. Nothing detects or repairs it.
- **A PIC still has no route to delete a pending or rejected file** — requests are raised on approved
  documents only. Out of band, a Head of Unit deletes it for them. Nothing on screen says so.

## 6. Technical debt worth an hour

- ⚠ **`Form.tsx`'s `<style>` template literal has been broken SEVEN times by a backtick in a CSS
  comment** — the `tsc` error names neither the cause nor the line. **Offered, not yet taken up:** move
  that stylesheet into a plain string constant so the hazard stops existing. About ten minutes, no
  visual change.
- **Five files are over the 2,000-line lint ceiling**: `Form.tsx` 5,756 · `BulkUpload.tsx` 4,366 ·
  `MySubmissions.tsx` 2,856 · `Requests.tsx` 2,721 · `FolderManager.tsx` 6,362. Warnings only.
- **`s.select` in `AuditLog.tsx` is now unused** (the Action filter became `FilterSelect`). Harmless —
  a `Record` key lint cannot see — and it is the style any future dropdown there would want.
