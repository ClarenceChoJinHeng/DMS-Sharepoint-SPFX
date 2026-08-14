# CRS Settings — the admin landing page

**Date:** 2026-08-14
**Status:** Built, not yet site-tested
**Web part:** `CRS Settings` (`4d8b7e21-9c65-4a3f-b028-5e71d9c4a836`), own bundle `crs-settings-web-part`
**Rules:** `src/shared/adminPages.ts` (pure, 33 tests)
**UI:** `src/webparts/crsSettings/components/CrsSettings.tsx` + `icons.tsx`
**Design source:** the client's mockup; icons in `docs/design-assets/`

---

## 1. Why

> "Lets design UI UX experience, client is complaining as they do not know what to do or how to
> operate. Look at the first screenshot, this is going to be a landing page for the admin pages, CRS
> settings. There will be Folder Management, CRS Configuration, User Access Management. Ignore CRS
> Mapping as the term abbreviation will be under Folder Management."

There are **fourteen** web parts and no menu. Every admin screen is reached today by typing or
bookmarking a URL, which is most of the complaint.

## 2. It is a DIRECTORY, deliberately

A card grid tells you where things are, not what to do — and "what to do" is the failure that actually
bites, because the **order** matters and nothing enforces it: miss the abbreviation step and
reconciliation creates nothing, silently, with no error anywhere. The client has asked the order
question twice in their own words (*"Where should I go step by step when: Creating new segment / Move
existing folder / When to run Folder Recon"*).

A "Common tasks" band carrying those workflows was offered and **declined** — the client has another
plan for the operating side (2026-08-14: *"don't worry about how to operate it, I got an idea"*). So this
page navigates and nothing more. Recorded here because it is a deliberate limit rather than an omission:
if the operating idea does not land, this page is where the order belongs.

## 3. What it lists, and the three additions

Two columns, exactly as mocked. **CRS Mapping is dropped** — term abbreviations moved under Folder
Management, which is where the Folder Administration tab bar had already put them.

| Card | Rows |
|---|---|
| **Folder Management** (green) | CRS Term Abbreviations · Folder Structure Management · Folder Reconciliations |
| **User Access Management** (purple) | **Group Management** · Site Access · Approval Library Access · Page Access · Folder Access |
| **CRS Configuration** (amber) | card is the link |
| **CRS Audit Log** (blue, new) | card is the link |
| **Bulk Upload** (rose, new) | card is the link |

Three were missing from the mockup and were added on the client's confirmation:

- **Group Management leads the access card, and had to be there.** Folder Access maps an **existing**
  group, so an admin who starts there has nothing to pick. Its absence from the mockup was the
  complaint in miniature.
- **CRS Audit Log** was on no menu at all — if it is not here, nobody finds it.
- **Bulk Upload** is an administrator tool per `pageAccessPolicy` (`adminOnly`), not an uploader page.

Rows are ordered by when the work happens, which is also the order of the tabs behind them.

## 4. Links resolve against the site — never hardcoded

Each row carries a regex, matched against every Site Pages entry's **file name and title**. Hardcoding
`Folder-Administration.aspx` was rejected because the client renames every DMS-named list, group and
content type at import (memory `dms-to-crs-rename-pending`) and page names will not survive either — and
a hardcoded name fails as a **dead link**: no error, no clue, on the one page whose job is telling an
admin where to go.

**The patterns are load-bearing and pinned by test**, because the plausible-looking version is wrong:

- `folder.?admin|folder.?manage|folder.?structure` for the three Folder Management rows. A bare
  `/folder/i` also claims **Folder-Access.aspx**, so "Folder Reconciliations" would open a permissions
  screen.
- `approval.?library|library.?access` for Approval Library Access. A bare `/approv/i` also claims
  **ApprovalDocument.aspx**, the approver's queue.

Three states, not two:

- **resolved** — one match, navigate.
- **ambiguous** — several. Navigates to the **shortest file name** (a duplicate is almost always the
  longer one: `…-Copy.aspx`, `…-old.aspx`) and **says so**, naming the others. A best guess beats a dead
  row, but silently picking one is how a half-renamed page sends people somewhere nobody meant.
- **missing** — the row renders **disabled, naming the page to create**. Never a dead arrow: a
  navigation page whose links go nowhere teaches the client the tool is broken.

The "create a page called…" advice names the **shared** page for a row that is only a tab of one
(`pageName`). Deriving it from the label would tell an admin to create `Folder-Reconciliations.aspx` — a
page the pattern can never match, leaving the row dead after they did exactly as they were told. Caught
by a test asserting every suggestion resolves, and it failed on the first run.

**An unreadable Site Pages list is NOT "no pages exist"** — the same empty ≠ unknown rule as everywhere
else here. It shows a red banner saying so explicitly, because the alternative renders ten rows telling
the admin to create ten pages that already exist. Overrides still apply in that state, so a
fully-overridden page keeps working.

**`FileRef` is used for the URL**, never a path assembled from the library name: a list's URL is
independent of its title (gotcha #12) and `SitePages` is not guaranteed to be the segment.

### The property-pane escape hatch

One optional address per link, stored flat as `link_<key>`. An override **wins outright and is never
pattern-checked** — it is the only fix available without a redeploy on a site whose page names cannot be
guessed, and second-guessing it would remove that. Each field's placeholder names the page
auto-detection is looking for, so the pane doubles as the answer to "why is this row greyed out".

Flat keys rather than a nested `overrides` object: SPFx binds one field to one property path, and a flat
key survives a link being added or removed (an orphaned property is inert, a stale nested shape needs
migrating).

## 5. Tab deep-linking

Folder Management's three rows are **tabs of one page** — `Folder Administration` is a five-tab web
part — so each row appends `#tab=<slug>`, which `FolderManager` reads on mount. Without it all three land
on the same tab and two of the three look broken.

- `tabFromHash` has **one implementation**, in `adminPages.ts`, shared by the page that writes the hash
  and the page that reads it. Two halves of one contract cannot be allowed to drift.
- **A slug is a public name once shipped** — an admin may bookmark the URL — so `DEEP_LINK_TABS` maps
  slug → `Tab` explicitly. Renaming a `Tab` value must not silently break a bookmark.
- An absent or unrecognised hash falls back to the default tab, **never to a blank screen**.

## 6. Icons

The client's three SVGs, converted to JSX **verbatim** — same viewBox, same path data, same colours
(`#00684A`/`#C9E6BD`, `#9561C9`/`#E9D9F9`, `#976B0A`/`#FFE19F`). Only attribute names change, since JSX
needs `clipPath`/`fillRule`/`strokeMiterlimit`. Nothing is re-authored, so the page matches what was
signed off. Originals kept in `docs/design-assets/` under meaningful names.

Two new ones drawn to the same recipe (54×54, `rx=10` pastel tile, one flat glyph in a darker shade of the
hue): **Audit Log** blue, a ruled sheet with a clock — the log answers *when*; **Bulk Upload** rose,
stacked sheets with one up-arrow, the stack being the "many at once".

`clipPath` ids are **document-scoped**, not element-scoped, so the id carries a prefix and must stay
unique across every icon on the page. `CardIcon` returns an empty tile for an unknown name rather than
nothing — a missing icon must not collapse the card header, `undefined` is not a valid JSX component
type, and this codebase avoids `null` (`@rushstack/no-new-null`).

## 7. Access

**No policy change needed.** A page named `CRS Settings` already matches the `setting` keyword in
`pageAccessPolicy.ts`, which resolves to `adminOnly` — verified before building rather than assumed.

## 8. Testing

1. Add the web part to a page named `CRS-Settings.aspx`. Every row should be live, none greyed.
2. Folder Management's three rows must open **Folder Administration on three different tabs** — Term
   Abbreviations, Folder levels, Folder Reconciliation.
3. Folder Access must open Folder Access, **not** Folder Administration. Approval Library Access must
   open the access screen, **not** the approver's queue.
4. Rename a page (or delete one) and confirm the row greys out naming the page to create, rather than
   becoming a dead arrow.
5. Duplicate a page and confirm the row still works but warns that more than one matches.
6. Set one property-pane address by hand and confirm it wins over auto-detection.
7. Confirm a non-admin cannot open the page at all (`adminOnly` by policy, granted at the page).
