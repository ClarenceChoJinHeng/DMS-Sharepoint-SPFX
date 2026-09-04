# Requests page redesign — scope, not a build (2026-09-03)

Client's mockups (pasted 2026-09-03, "Approver – Request for Approval" ×2) ask for a bigger
change than a copy tweak: the approver's "Requests" page becomes **"Document Deletion & Sharing
Approval"**, reachable through a restructured left-nav item **"Approval & Request"** with three
sub-pages, and the current three-tab layout (Delete File / Share / Shared files) becomes two
labelled sections plus a separately-headed shared-access panel.

**Nothing here is built.** This is a scope, written per this project's standing rule to spec
before implementing (`feedback_write_specs_before_implementing`) — and because the client
themselves said this is a big piece of work, worth confirming the shape before spending a session
on it.

## What the mockups actually show

1. **Site navigation**, not a web part: `Approval & Request` as a parent nav item with three
   children —
   - `Approval (Confidential & Restricted)`
   - `Approval (Highly Confidential)`
   - `Approval (Deletion & Sharing)`
2. **A renamed, restructured Requests page**, titled "Document Deletion & Sharing Approval",
   subtitled "Approve requests to delete or share documents." Top bar: `Requests [N]` +
   `Access Audit` (a link, right-aligned).
3. **Delete and Share merged into ONE view**, not two tabs to switch between — two labelled,
   colour-coded sections stacked on one page:
   - `🗑 Delete Requests [N pending]` — red left-border cards
   - `📤 Share Requests [N pending]` — blue left-border cards
   - Each card: filename (bold), `requester · department · date`, "Requester's reason" + text,
     Approve (green) / Reject (red outline) buttons — no intermediate click to a detail view.
4. **"Shared files" stops being a third tab.** It becomes its own section, "Documents with Shared
   Access (N)", inside the same page — same content as today's tab (filter, `Re-check access`,
   per-recipient `has access`/`could not be checked` state, `Revoke`/`Revoke all access`), just no
   longer competing with Delete/Share for a tab click.

## What already exists and does NOT need to change

- **The reject/approve dialog** (redesigned 2026-09-03, shipped `1.0.375.0`) — exact copy match,
  required-note validation, same component for both request types. Reused as-is; the merged view
  just needs to call the same `setDeciding({...})` handlers from card buttons instead of from
  inside a single-type tab.
- **`Access Audit`** — this project already has a resolved-link pattern for exactly this
  (`shared/adminPages.ts`'s `resolveLink`/`CARDS`, used by `crs-settings` and the `backToSettings`
  band on every admin page). The new page's `Access Audit` link is a straightforward reuse, not a
  new mechanism.
- **The "Shared files" section's own logic** (`collectSharedFiles`, `mergeShareAcl`, `withoutGoneFiles`,
  revoke/revoke-all) — untouched. Only its position on the page and its heading change.
- **`ViewerScope`/`canDecide`/`isVisibleTo`** — the visibility and decide-rights rules are
  unaffected by how the page is laid out; a merged view still has to respect who can see what.

## What is genuinely new work

1. **Site navigation restructuring.** Adding a parent nav item with children is a SharePoint
   site-navigation edit (`Edit navigation` in the left rail) — **not code**, and not something a
   web part package can do on install. Someone with site-edit rights does this by hand, once per
   site (ClarenceDMSTesting, then SDG). The two "Approval (...)" children pointing at the existing
   normal/HC `ApprovalDocument.aspx` pages need no new page — just a nav entry with a friendlier
   label than whatever the page is currently called in nav.
2. **`Requests.tsx`'s tab bar goes away**, replaced by a single scrolling page with two (or three,
   counting Shared files) headed sections. This is a real internal restructuring:
   - `tab`/`setTab`/`TABS`/`TAB_LABEL`/`tabCount` — the whole tab-switching mechanism — is removed
     or repurposed as anchor navigation within one page (a `Delete Requests` / `Share Requests` /
     `Shared files` set of headers with counts, all visible without a click).
   - The filter bars (`All outcomes` dropdown, `Filter by document or requester email`) currently
     live once per tab. Need to decide: one shared filter bar for the whole page, or one per
     section (the mockup shows filtering only inside "Waiting for you", not per merged section —
     worth confirming which).
   - `requestCard` already renders Approve/Reject buttons inline (`actionable` param) — the merged
     view can likely reuse it directly under each section, keyed by `r.type` instead of by which
     tab is active.
3. **Page title/rename.** "Requests" → "Document Deletion & Sharing Approval" is a page-title and
   possibly a URL-friendly-name change. Cheap, but it means updating anywhere this page is
   currently linked BY NAME (`adminPages.ts`'s `/request/i` pattern already matches by regex, not
   literal title, so this should resolve automatically — worth a grep-and-check before assuming).

## Open questions to settle with the client before building

- **Does Approve/Reject fire directly from the merged-view card**, or does clicking a card still
  open the (already-redesigned) decision dialog? The mockup's cards show the buttons directly on
  the card with no visible intermediate step — but the required-note field has to live somewhere,
  and putting a note input on every card at once is a different, busier design than opening one
  dialog per decision. Recommend: buttons on the card still open the existing dialog (cheapest,
  reuses everything already built); flag this assumption rather than silently deciding it.
- **One shared filter bar, or one per section?** Affects `whoFilter`/`statusFilter` state shape.
- **Does "Delete Requests"/"Share Requests" replace the queue/decided split, or sit inside it?**
  Today's page has "Waiting for you" (queue) and "Decided" as the two top-level sections; the
  mockup's "Delete Requests"/"Share Requests" split is orthogonal to that (type, not status). Need
  to confirm whether decided requests also get shown split by type on this same page, or whether
  the merged type-split view is queue-only and "Decided" stays a separate, unified list below it
  as it is today.
- **Does the nav rename touch SDG too, or ClarenceDMSTesting only for now?** Nav edits are
  per-site and manual; worth knowing before promising both.

## Recommended order, once confirmed

1. Settle the four open questions above with the client.
2. Restructure `Requests.tsx` internally (remove/repurpose the tab bar, merge Delete+Share into
   one page with two headed sections, keep Shared files as a third section) — this is the real
   code work and can be built and tested without touching navigation at all.
3. Rename the page title.
4. Site-nav edit (`Approval & Request` parent + three children) — manual, on each site, last —
   since it is what a user actually clicks and should point at a finished page.
