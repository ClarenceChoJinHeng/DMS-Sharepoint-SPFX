# Access pages — usability pass (Approval Library Access, Site Access, Page Access)

**Date:** 2026-09-02
**Status:** ✅ BUILT — `tsc` clean, full suite 1597/1597, zero new lint warnings. NOT yet site-tested.
**Follows:** `2026-09-02-access-pages-read-only-design.md` (which made all three read-only). This
spec is the second pass, driven by the client's own live walkthrough of the deployed read-only pages.

---

## 0. Context

The client tested the read-only conversion live and came back with per-page feedback plus one
genuine performance complaint (Site Access took ~8 minutes to load). This spec organizes that
feedback into concrete changes, and flags the handful of items that need a decision before writing
code rather than guessing at wording or scope.

---

## 1. Approval Library Access (`StagingAccess.tsx`)

**Confirmed changes:**
- Add a search/filter box over the group list.
- Add an inner scroll container for the group list (it can run to 60+ groups per segment).
- Remove the **"Should have access"** column.
- Rename/reframe **"Access now"** to read as a plain, read-only fact rather than an editable-looking
  column header — proposed: **"Current access"**.

**Reasoning for dropping "Should have access":** on a page that can no longer act on a mismatch,
showing the "should" column beside the "is" column reads as an open discrepancy asking to be
fixed by hand, which the page can no longer do. The client's own words on the Page Access page make
the same point ("if they are listed here of course they have access") — the "should" column implies
doubt that no longer belongs on a read-only screen.

**Open question:** none — this one is unambiguous.

---

## 2. Site Access (`SiteAccess.tsx`)

**Confirmed changes:**
- Scrollable container for "People who can open the site" (the site-entry group's members).
- Scrollable container for "Everything with permission on the site itself" (the full per-principal
  permission dump).
- Add a filter/search box for the "Everything with permission" list.
- Simplify the permission-level display. Raw SharePoint labels currently shown: `Full Control`,
  `Limited Access`, `Web-Only Limited Access`.

**✅ CONFIRMED (2026-09-02).** `Web-Only Limited Access` is a real, distinct SharePoint permission
level (`WebOnlyLimitedAccess`), separate from ordinary `Limited Access`. Both are automatic traversal
grants a principal gets when they hold access on something below the level shown (a folder, a page)
— never something granted by hand and never something that lets someone browse or list anything on
its own. The difference between them (whether the grant also covers "web-only" scenarios like some
modern-page contexts) is not meaningful to a client reading this page.

**Three plain-language buckets**, replacing all raw level names:
| Raw SharePoint level(s) | Shown as |
|---|---|
| `Full Control` | **Full access** |
| `Limited Access`, `Web-Only Limited Access` | **Reaches folder/page, no site-wide access** |
| anything else (a real named permission level, e.g. `Read`, `CRS Approve`) | shown as-is — these are real, deliberate grants worth naming precisely |

This treats the two "Limited Access" variants identically (both mean "can't browse from here, only
reach one specific thing further down") — which is accurate and matches this project's own existing
language on Page Access ("Limited Access... grants NOTHING... automatic entry").

**⚠ THE 8-MINUTE LOAD — investigate before patching over it.**
Not proposing a fix yet. Likely cause: `fetchAllGroupMembers`/`readAcl`-style per-principal reads
running sequentially over hundreds of site groups (this project's own `fetchAllGroupMembers` design
note already says a per-group read is "not an export, it is an outage" at scale — this page may be
doing something in that shape for the full permission dump). Plan: read the actual page code, confirm
what it's doing per-principal, and report back with either (a) a real fix (batch/parallelize/cache),
or (b) confirmation this is inherent to the number of principals and a progress indicator is the
honest answer. **Will investigate as a separate step before touching this page's code.**

---

## 3. Page Access (`PageAccess.tsx`)

**Confirmed changes:**
- Remove the explainer paragraph: *"This hides the page, not the documents on it..."*
- Add filter + search + scroll to the group list.
- **Stop showing every group in the tenant/site.** Currently every `_APPROVER`-suffixed group
  site-wide is listed for a page like `ApprovalDocument.aspx` (correct per the derived-page-access
  design — page grants are NOT scoped to one unit, any approver anywhere can open the shared
  approval page) — but showing all of them, most with "no members" and "granted in SharePoint, not
  mapped", reads as noise. Change: **only list groups that are ACTUALLY currently granted access to
  this page** (i.e., have a real, non-Limited-Access binding on the page's ACL) — drop rows for
  groups that match the page's policy by name/role but hold no actual grant yet.
  - ⚠ This changes what the page answers, worth stating plainly: today it answers "which groups
    SHOULD/COULD have access" (derived from the naming policy); after this change it answers "which
    groups CURRENTLY have access" (derived from the live ACL). The "Should have access" column removal
    (below) makes this the only sensible reading anyway.
- Add a popup showing the member list when clicking "N members" (reuses `MemberSummary`'s existing
  expand behaviour, or a lightweight modal — implementation detail, not a design question).
- Remove the **"Should have access"** column (same reasoning as Approval Library Access).
- Style "Access now" as read-only-flavored text, matching the other two pages' "Current access"
  language.

**Open question:** none on the mechanics — the filtering-to-actually-granted-groups change is
mechanical (already have `readAcl`'s live binding data on this page); it just changes what gets
rendered.

---

## 4. Shared across all three

- "Should have access" is removed everywhere for the same reason: a column implying an open question
  ("should this be different from what it is?") does not belong on a page that can no longer answer
  it by doing anything.
- Filter/search/scroll all reuse the SAME small set of styles this project already has elsewhere
  (the Requests page's filter bar, the Group Management search box) rather than three independent
  implementations.
- **⚠ Scroll containers must be checked for anything absolutely positioned inside them before
  shipping** — this project has hit that exact defect three times already (Group Management's people
  picker, the upload form's info tooltips, a comment-box dropdown). None of these three pages are
  expected to have anything absolutely positioned, but confirm before capping height.

---

## 5. Order of work

1. Investigate the Site Access load-time issue (read the code, report findings) — done BEFORE any UI
   change to that page, since a UI polish pass on a broken page is a wasted diff if the real fix
   reshapes how it reads data.
2. Approval Library Access — search/filter/scroll, remove "Should have access", rename "Access now".
3. Page Access — remove explainer text, search/filter/scroll, member popup, filter list to
   actually-granted groups, remove "Should have access", read-only styling.
4. Site Access — apply whatever the investigation in step 1 recommends, plus scroll/filter/permission
   simplification (pending your confirmation on the three-bucket wording above).
