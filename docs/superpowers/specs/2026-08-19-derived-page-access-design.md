# Derived page access — reconciliation grants the pages a role implies

**Date:** 2026-08-19
**Status:** BUILT 1.0.171.0, NOT site-tested. ⚠ §1's original premise was wrong and is corrected in place — read it before citing this spec.
**Supersedes nothing.** Extends the page pass of the 2026-08-05 page-access work and sits alongside
`2026-08-17-admin-page-lockdown-design.md`.

---

## 1. The problem, as found on site

`Upload-Form.aspx` is restricted. Which groups may open it comes from `policyForPage` →
`["UPL", "APR"]`. But reconciliation only ever granted a page to groups that had a
**`Scope = Page` row** in `CRS Group Map` — and nothing creates those rows automatically. Bulk
provisioning writes **Folder**-scope rows; the guided flows never mention page access; Group
Management does not offer it.

**⚠ THE ORIGINAL FRAMING OF THIS SPEC WAS WRONG, AND THE CORRECTION IS WORTH KEEPING.** It was
written from a progress screenshot showing eight `upload-form.aspx` grants, read as the complete set;
the full log showed ~120 — every `_UPLOADER` and `_APPROVER`. The administrator had granted them on
**Page Access**, selecting all the groups. So nobody was locked out, and **a mid-run progress panel is
not a result.** The premise was checked only after the code was written.

What remains true, and is the actual justification:

- **Nothing writes a `Scope = Page` row except `PageAccess.tsx`.** Bulk provisioning writes Folder
  rows; the guided flows never mention page access; Group Management does not offer it. The rows
  exist only because a person went to that screen and made them.
- **A unit or segment added later therefore gets folders, groups and no page.** Nothing reports it,
  and it appears months later as one person's AccessDenied — exactly how 2026-08-17's hour of
  diagnosis started.
- **`My-Submissions.aspx` and `Requests.aspx` had no rows at all**, so they were still inheriting:
  openable by anyone who can open the site.

**This is the fourth instance of one structural gap**, and the pattern is now unmistakable:

| Instance | The mechanism was driven by | The thing needing protection had |
|---|---|---|
| Inverted site-entry library grant (2026-08-16) | library-scope rows | no row |
| HC libraries inheriting (2026-08-17) | library-scope rows | no row |
| Admin pages readable by uploaders (2026-08-17) | page rows | no row (`roles: []`) |
| **A new unit's page access (2026-08-19)** | **page rows** | **no row until someone makes one** |

Every one was fixed the same way: **assert the required state every run** instead of deriving it
from the presence of a grant row. This does that for the pages that have a role policy.

---

## 2. What is derived

A page's intended access is computed from the roles groups **already hold at folder scope**:

```
intended(page) = { g : roles(g) ∩ derivedRolesForPage(page) ≠ ∅ }
               ∪ { g : a hand-made Scope=Page row targets this page for g }
```

`roles(g)` collapses every Folder-scope row for that group across every segment and tier. Tier is
deliberately ignored: a group holding `UPL` **anywhere** needs the upload form. The page is not
scoped to a unit and cannot be.

For the rehearsal site that yields, from the same 790 rows already read — matching the ~180 grants
an administrator had made by hand, and adding the two pages that had none:

| Page | Roles | Groups |
|---|---|---|
| `Upload-Form.aspx` | `UPL, APR` | ~120 (60 uploader + 60 approver) |
| `ApprovalDocument.aspx` | `APR` | ~60 |
| `My-Submissions.aspx` | `UPL` | ~60 |
| `Requests.aspx` | `UPL, APR` | ~120 |

### 2.1 ⚠ Derivation fires only on a page whose name MATCHED A RULE

`policyForPage` returns the same shape whether a rule matched or the name fell through to
`DEFAULT_POLICY` (`UPL, APR, DELS`). **Deriving from that default would be a site-wide lockout.**

Granting a group Read on a page requires **breaking that page's inheritance first**, and after that
only the listed groups can open it. A page that matched nothing — the site home page,
`CollabHome.aspx`, anything the client authored — would be broken and granted to `UPL/APR/DELS`,
which **removes** every other role: SDG Employee, Head of Department and C-Level hold none of them.
Reconciliation would take the site away from most of its users while reporting success.

So `pageAccessPolicy.ts` gains:

- **`pageMatchedRule(fileName)`** — did an explicit `RULES` entry match?
- **`derivedRolesForPage(fileName)`** — the roles eligible for automatic derivation: `[]` unless a
  rule matched **and** the page is not `adminOnly`.

Both read the existing `RULES` array, so there is no second list of page names. A hardcoded list of
four file names inside `FolderManager.tsx` was rejected for the reason `tabFromHash` and the
abbreviation collision rule are single-sourced: **the drift is silent.** A page renamed in one place
and not the other either stops being granted or starts being locked, and neither reports anything.

Consequences, pinned by test:

- `adminOnly` pages derive **nothing**. They have `roles: []` and are owned by the lockdown pass.
- Unmatched pages are **left inheriting** — not broken, not granted, not logged as an error.
- A page whose eligible roles are view-only derives nothing (`VIEW_ONLY_ROLES` are never offered a
  page).

---

## 3. Hand-made Page rows survive, and merge

They are **not** replaced. Derivation covers the rule; a `Scope = Page` row is the only way to grant
a page to a group the policy does not imply, and removing that capability silently would be a
regression on a screen the client already uses.

Merged and **deduped on `GroupId`**, never on group name: two groups can be renamed alike, and a
stored `GroupName` goes stale the moment a group is renamed (which a rename deliberately allows,
since the Id survives). The existing refusals are untouched and still apply to a hand row:

- the welcome page cannot be restricted (`isForbiddenPageTarget` / `file === welcome`);
- an `adminOnly` page refuses page rows, so the two passes cannot fight;
- a row whose `Target` names no existing page is reported — **but a derived page that does not exist
  is silently skipped**, because it is simply absent from Site Pages and there is no typo to report.

---

## 4. The page's FULL access is asserted — additions and removals

Once derived and hand-made rows together describe the complete intended set, the pass can assert it
rather than only adding to it:

```
page permissions = site Owners (Full Control)
                 + intended groups (Read)
                 + nothing else that is a SharePoint group
```

**Every removal is logged with the principal named:**

```
⚠ Upload-Form.aspx: removed GHO_GF_TAX_EMPLOYEE — holds no UPL or APR role
```

Silently removing a grant someone made deliberately is worse than not removing it, because the
administrator goes on believing it is there. Same rule as the admin lockdown pass.

### 4.1 Why removal is safe HERE and nowhere else

**A Site Pages item is a leaf.** Nothing is scoped beneath it, so its role assignments contain only
what someone deliberately granted.

**A library or folder is not.** SharePoint auto-creates a **Limited Access** assignment at every
parent scope for any principal holding a grant further down — so the Approval Document library
root's permission list contains an entry for **every one of the ~308 groups** holding a folder grant
inside it. None was authored; all are load-bearing. A pass that removed "anything not in my intended
list" at library or folder scope would strip them and **every group would lose its folder access**,
on a run that reported success.

Stated as a hard rule:

| Scope | Grant | Revoke |
|---|---|---|
| **Page** | derived + rows | ✅ asserted, non-intended groups removed |
| **Library** | rows only | ❌ never — Limited Access lives here |
| **Folder** | rows only | ❌ never — same |

### 4.2 What is never removed, even at page scope

- **Site Owners only**, protected by id — matching the lockdown pass, which also keeps exactly that
  one principal. Removing Owners is how a page becomes unreachable by the people who administer it.
  - **The site-entry group is deliberately NOT protected.** It holds Read on the *web*, so on an
    inheriting page it is present by inheritance and vanishes the moment inheritance is broken. A
    page-**scope** assignment for it can therefore only have been added by hand, and it would give
    every plain member the upload form — the precise thing the policy exists to prevent. Members and
    Visitors likewise: a CRS page granted to the site's default members group is not a grant to
    respect.
- **User principals.** Only SharePoint *groups* are candidates. A directly-granted user is already
  flagged in red on Page Access, and individual grants were explicitly rejected as a mechanism in
  `2026-08-14-per-person-access-removal-design.md`; removing them here would quietly implement the
  thing that spec declined.
- **Anything at all, if the current assignments could not be read.** An unreadable ACL means the page
  is skipped unchanged — stripping what you could not read removes grants you never saw.

---

## 5. The warning that was missing

A page that matched a rule and ends with an **empty** intended set logs:

```
⚠ Upload-Form.aspx: restricted, but no group holds UPL or APR — nobody but site owners can open it
```

A **warning, not a refusal.** An empty set is correct on a site with no groups yet, and the pass
still breaks inheritance and restores Owners — the page is simply administrator-only until groups
exist. Refusing would leave it inheriting, i.e. readable by every site member, which is worse.

Nothing reports this today, on any site.

---

## 6. Failure rules

Consistent with the rest of reconciliation: **fail open on the read, closed per page on the write.**

| Condition | Behaviour |
|---|---|
| `CRS Group Map` unreadable | derive nothing; pages left as they are. Never break a page's inheritance off a read that failed — that is a lockout caused by a transient error. |
| Site Pages unreadable | nothing at all (existing behaviour). |
| A page's current assignments unreadable | that page skipped unchanged, reported. |
| `breakroleinheritance` fails | reported `nothing granted`, never as locked. |
| A single `addroleassignment` fails | reported per group; the rest continue. |
| A `GroupId` that is not an integer | reported per row, as today (`spGroupPrincipalId` throws). |

**The Group Map read is the existing one** at `FolderManager.tsx:2632` — `$top=5000`, which does
**not** lift the 5,000-item threshold. At 790 rows there is headroom for five segments; past that
this read must be paged, as `BulkGroupProvisioner`'s already is. Noted, not built: a truncated read
would under-derive and then strip grants that should stay. Backlog.

---

## 7. What this does not fix

- **A group with no rows gets nothing.** Derivation keys on `Role` values, not names, so a
  hand-named group with proper rows works — but one created with the advanced free-text name and no
  persona has no rows at all (`rowsForNewGroup` returns `[]`) and therefore no page. Unchanged.
- **Membership is still intent.** The right *people* being in the group is not checkable here or
  anywhere.
- **A deleted group's grant.** Deleting the SP group drops its role assignments with the principal,
  so this needs no help. Removing only the *mapping rows* while keeping the group is what the
  assertion in §4 cleans up on the next run — at page scope only.

---

## 8. Migration

**None.** No column, no list, no row change. Redeploy and re-run reconciliation.

The ~180 hand-made Page rows may stay or be deleted; the outcome after a run is identical, because
derivation covers those groups and the merge dedupes on `GroupId`.

⚠ **BUT THIS IS NOT A NO-OP ON A LIVE SITE.** `My-Submissions.aspx` and `Requests.aspx` currently
have no rows and therefore still INHERIT — anyone who can open the site can open them. The first run
after deploying breaks their inheritance and restricts them to `UPL` and to `UPL`+`APR`. That is the
policy working as designed, and it is a visible change to announce rather than discover.

---

## 9. Files

| File | Change |
|---|---|
| `src/shared/pageAccessPolicy.ts` | `pageMatchedRule`, `derivedRolesForPage` |
| `src/shared/pageGrants.ts` | **new, pure** — `groupRolesById`, `intendedPageGroups`, `groupsToRemove` |
| `src/shared/pageGrants.test.ts` | **new** |
| `src/shared/pageAccessPolicy.test.ts` | derivation gating, every page the runbook creates |
| `src/webparts/folderManager/components/FolderManager.tsx` | split the Group Map read; merge; assert; log removals; warn on empty |
