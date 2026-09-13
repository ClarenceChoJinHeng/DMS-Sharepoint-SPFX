# Archive code reuse guard — design

2026-09-13. Agreed with the client (relayed by Reene) after a design discussion about the PCAR
incident and how to stop it recurring, without building a new list/mapping system.

## Where this came from

A segment's top folder (`StagingFolder`, e.g. `PCAR`) has no term behind it — it is matched purely
by NAME. Retiring a segment never deletes a non-empty Archive folder (7-year retention). So if a
*different*, unrelated segment is later created reusing the same top-folder code, its documents
land inside the *same* Archive folder as the old, retired segment's leftover records — visibly, at
the top level, with nothing distinguishing them. A client browsing Archive sees one folder named
`PCAR` containing files from two unrelated businesses, with no way to tell them apart at a glance.

This is not a data-safety problem (nothing gets moved or mis-tagged) — it is a **naming collision**
a human notices and cannot explain.

## Rejected: a separate Archive abbreviation + folder mapping

Considered and dropped. It solves the uniqueness problem but:
- Needs a new required step at segment creation (a second, Archive-only code).
- Needs a persistent mapping list, kept in sync forever.
- Needs the 7-year archive-mover flow updated to consult it.
- Creates its own new client-facing confusion: "why doesn't the Archive folder's code match the
  Document Library's code for the same segment?" — the very kind of confusion this whole feature
  exists to prevent, just moved one level over.

## The rule

At segment creation, before anything is written, check whether Archive/ArchiveHC already has any
files under the proposed top-folder code. If it does, **refuse to create the segment with that
code** — a hard block, not a dismissable warning. The message must say explicitly that the code
already exists in the **Archive** library specifically, since Staging and Documents will show no
conflict at all (that is exactly what will confuse an admin if the message doesn't say so).

- **Scope: the segment's own top-folder code only.** Not department/unit abbreviations. Reasoning,
  confirmed with the client: SharePoint does not allow two sibling folders with the same name, so
  once the top-level code is guaranteed unique, nothing beneath it can ever collide with a
  *different* segment's history — it would always be created fresh under a folder that has never
  existed before.
- **Hard refusal, no override.** Confirmed with the client: not needed, because (a) the check
  itself prevents any collision from ever happening, so there is nothing to override *around*, and
  (b) a segment retired with a genuinely empty Archive already has that folder deleted (the
  2026-09-12 empty-archive-deletion feature), so the code becomes available again on its own once
  it is safe — no manual override path is needed.
- **Fails CLOSED on an unreadable check**, matching `canDeleteArchive`'s own reasoning elsewhere in
  this codebase: an unconfirmed archive state must never be treated as "clear", because the cost of
  guessing wrong here is a silent, hard-to-explain collision discovered much later.
- **Checked fresh at submission time**, not merely as a live hint while typing — the same
  "validate first, create nothing until it passes" order the form already follows for its other
  checks (name clash, key clash, folder clash, depth). Given time constraints, no live-as-you-type
  debounced indicator is built for this one (unlike the Term Set field) — the submission-time
  refusal is enough to make the collision structurally impossible, at the cost of the admin only
  finding out on Create rather than while typing.

## What this does NOT fix

Existing test data (PCAR/PCT on SDG's site) already has this exact collision from before this
guard existed — confirmed as test-only content, not a live client concern, and left as-is. This
guard prevents recurrence; it does not clean up what already happened.

The "Move existing folders" tool ignoring already-orphaned archive folders (a complementary
safety net for any *existing* collision) is explicitly OUT OF SCOPE for this change, per the
client's own call given time constraints — the reasoning is that this guard, going forward, makes
new collisions impossible, so that net is not urgently needed.
