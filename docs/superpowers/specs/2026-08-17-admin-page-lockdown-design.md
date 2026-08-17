# Admin page lockdown — reconciliation restricts the pages that have no grant rows

**Date:** 2026-08-17
**Status:** Agreed. Closes open finding #1/#6, the last blocker before migration.
**Scope:** one new pass in `FolderManager.tsx` reconciliation, plus one refusal in the existing page pass.

---

## 1. Why

> "can we not auto restrict the admin pages?"

`CRS_SITE_MEMBERS` holds **Read at site level** — it must, or a user granted only a folder is denied
on Home and the site reads as broken. Site Pages inherits from the web. So **every uploader can open
every admin page**: Folder Administration, Group Management, Site Access, Folder Access, Page Access,
the Audit Log and Bulk Upload.

Confirmed live 2026-08-16 with a guest account holding only Read plus Limited Access.

**It is not a code bug.** `pageAccessPolicy.ts` is a UI filter and says so; the real boundary is the
SharePoint page grant. Reconciliation *does* break inheritance on Site Pages and grant from Group Map
— but it iterates only pages that **have** rows. An `adminOnly` page has no eligible roles, so nobody
can create a row for it, so the pass never visits it.

**The mechanism is driven by grants, and the thing that needs protecting has none.** The same
structural gap left both HC libraries readable by every site member (finding #10) and inverted the
site-entry grant (#7). Both were fixed by asserting the required state every run instead of deriving
it from rows. This is the third and last instance.

Until now the mitigation was a manual step in the migration runbook (§8.1). A required manual step
that is invisible when skipped is not a mitigation — it is a deferred incident.

---

## 2. What it does

A new pass, after the row-driven page pass.

For every item in **Site Pages** whose file name satisfies `policyForPage(name).adminOnly`:

1. If the page still inherits → `breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`.
2. Grant the **site Owners** group Full Control.
3. Remove every **other** role assignment on the page.
4. Report each action, naming the page.

Site collection administrators are unaffected — they bypass role assignments entirely, so an admin
cannot lock themselves out with this.

`copyRoleAssignments=false` for the reason it is false everywhere else here: with `true`, every
inherited grant is carried forward, so the page stays visible to exactly the same people and the run
reports success. A failure invisible from the log.

---

## 3. Assert every run, in both directions

Step 3 is what makes this an assertion rather than a one-time fix, and it is the half that is easy to
leave out.

A page can be **unique and still exposed**: inheritance broken, with a group granted Read on it by
hand or by an earlier version of this code. Checking only `HasUniqueRoleAssignments` would call that
page locked. So the pass reads the assignments and strips anything that is not the Owners group — the
same shape as the site-entry library pass, which asserts presence *and* absence every run.

Every removal is logged with the principal named. An administrator who deliberately granted someone
the Audit Log will see it taken away and read why. Removing it silently would be worse than not
removing it at all, because they would go on believing the grant was there.

---

## 4. The two passes must not fight

An `adminOnly` page has `roles: []`, so `GroupMapBuilder` never offers one — but a hand-authored row
can still target one. Without a guard the row pass would grant Read and this pass would immediately
remove it, every run, both logged, for ever.

So **the row pass refuses a row whose target is an `adminOnly` page**, exactly as it already refuses
one targeting the welcome page: the row is left in place (deleting authored data is not this pass's
job), the refusal is logged, and it can be removed deliberately.

Order then stops mattering — which is the point. A correctness property, not a sequencing convention.

---

## 5. What it must never touch

- **The welcome page.** Resolved from `RootFolder/WelcomePage`, not merely the `home.aspx` constant —
  a renamed welcome page slips past the constant, and locking it makes the site unreachable for
  everyone who is not an administrator. The existing `isForbiddenPageTarget` guard is reused, never
  reimplemented.
- **Any page that is not `adminOnly`.** Upload Form, Approval Document, My Submissions and Home keep
  their row-driven grants.

---

## 6. Failure states

**This pass fails CLOSED per page and OPEN on the read**, and the asymmetry is deliberate:

- **Site Pages unreadable** → warn, lock nothing. There is no page list to act on and inventing one
  is impossible. The runbook's manual step remains the backstop.
- **A page's assignments unreadable** → warn and skip **that page**, changing nothing. Stripping
  assignments from a list we could not read would remove grants nobody could see.
- **Break or grant fails** → report the page as **still open**, with the HTTP status. Never as
  locked. A page reported locked that is not is the one outcome worse than today's, because it stops
  anyone looking.

An unreadable ACL is not evidence of absence — gotcha #11, and the site-entry pass.

---

## 7. Pattern matching, and its one risk

Pages are matched by **file name**, through the existing `policyForPage`. On a site containing only
our pages this is exact. On a site with other content, a page called `Configuration.aspx` would match
`/configuration|crs.?config/i` and be locked.

Accepted, with mitigation: **every lock is logged by name**, so a wrong one is visible in the run log
rather than discovered by the person who lost access. The client's site
(`sdguthrie.sharepoint.com/sites/CRS`) was created for this system and holds essentially nothing else.

Deliberately NOT mitigated by an allow-list of exact file names: the landing page resolves links by
pattern precisely because this client renames everything at import, and an exact list would stop
matching on the day of the rename — silently, which is the failure class this pass exists to remove.

---

## 8. Scope

Everything is built from `context.pageContext.web.absoluteUrl`. **This pass cannot reach another
site**, and neither can anything else in the solution: site groups are site-collection scoped, the
term store group is site-collection local, the customizer registers a web-scoped `UserCustomAction`,
and CRS Search never issues an unscoped KQL query (`kqlPathScope` returns blank rather than unscoped,
and blank means *do not run*).

---

## 9. Out of scope

- **An `IsSiteAdmin` gate inside each admin web part** (finding #2). Defence in depth and worth
  doing, but a different change: the boundary is the page grant, and that is what this fixes.
- **Restricting the lists** (`CRS Config`, `CRS Group Map`) as opposed to the pages. A determined
  uploader can still open a list by its own URL. Worth a later pass; it does not block migration, and
  the admin *tools* are what was asked about.
- **Removing the runbook's §8.1.** It stays, rewritten as a verification step — "check that a
  non-admin cannot open these" — because verifying is not the same as performing.
