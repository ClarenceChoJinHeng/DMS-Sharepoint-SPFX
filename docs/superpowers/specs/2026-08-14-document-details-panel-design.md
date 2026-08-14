# Document details panel — derived tier rows

**Date:** 2026-08-14
**Status:** Built for My Submissions, not yet site-tested. `ApprovalDocument` NOT yet migrated.
**Rules:** `src/shared/documentDetails.ts` (pure, 40 tests)
**Supplements:** `2026-08-14-my-submissions-design.md` — this is the fifth bug found on that page.

---

## 1. What was reported

> "Nice but I notice the Details are so little"

Said of the My Submissions detail view, on a file filed at
`UPOPSMY › JHR › BKB › 2024 › Working File`. The panel showed **Business Segment: Upstream Operations
Malaysia** and then jumped straight to Document Type — `JHR` (Region) and `BKB` (Estate/Mill), the two
values that decide where the file lives, were absent.

## 2. The cause, and why it was guaranteed to happen

The panel was a **hardcoded list of fields** that named `Department` and `Unit` literally:

```ts
["Department", pickField(fieldText, "Department")],
["Unit",       pickField(fieldText, "Unit")],
```

Upstream Operations Malaysia has neither. Its permissioned tiers are `Region` and `EstateMill`
(created by SegmentCreator on 2026-08-12 in both libraries). Both hardcoded rows resolved blank,
blank rows are dropped — correctly, since most columns are empty for any given document on a library
serving twelve segments — and the segment's actual tiers had no row to appear in at all.

**This was not specific to one segment.** The admin NAMES the permissioned tiers when onboarding
(`Region → Estate/Mill`, a single tier for I&T, `Department → Unit` for the head offices), and the
Structure Manager can add below-Unit tiers at any time. A list of field names can only ever be right
for the segments it was written for, so every future onboarding would silently lose its own tiers from
this panel.

The same list, with the same two hardcoded names, is at `ApprovalDocument.tsx:704-705` — see §6.

## 3. The fix: derive the tiers

Two facts make it derivable **with no extra request**:

1. **Every tier column has a `<Base>Tid` sibling** holding the term GUID, because that is how
   `ensureColumn` creates them in `SegmentCreator` and `StructureManager`. So a field is a tier field
   **iff its Tid twin is present**. No ordinary column has that signature.
2. **`FieldValuesAsText` returns EVERY field on the item**, Tid columns included, so the evidence is
   already inside the response the panel has. No `/fields` read, no DMS Config read, no matching a
   segment label to a mode row — and therefore no new failure mode.

`Year` and `Document Type` deliberately have **no** Tid column: they are managed metadata, and a
derived `YearTid` would have started writing a bare label into a taxonomy field (the reason
`effectiveOnDemandTiers` seeds that pair without a `tidCol`). That is precisely why they stay in the
fixed list rather than being discovered — the absence is load-bearing, not an oversight.

Tid columns are excluded from display: they hold GUIDs, which is what the term store exists to keep
off the screen.

### Ordering and labels

**Order** comes from the response, which follows the library's field order and therefore the order the
tiers were created in — chain order, because SegmentCreator creates them walking the chain. Business
Segment is **pinned first** regardless: it is the top of the hierarchy, the one position that is
semantically load-bearing rather than cosmetic. If the rest ever came back out of order, the cost is a
reordered panel and never a missing value.

**Labels** are derived too. `_xNNNN_` decodes to a space, which recovers the display name exactly for
columns that have one (`Business_x0020_Segment` → `Business Segment`). A name sanitized at creation
cannot be recovered — `sanitizeFolderSegment` REMOVES illegal characters, so `Estate/Mill` became the
column `EstateMill` and the slash is gone — so it is split on camel case into `Estate Mill`. Readable
and honest. Guessing the separator back would print punctuation nobody typed, and would be wrong for
any name that never had any.

Exact display titles would need a `/fields` read per library. Deliberately not taken: one more request
and one more failure mode, to recover one character in one label.

### The double-encoded key is now derived, not typed

`encodedKey` is `name.replace(/_/g, "_x005f_")`. Both detail views listed the two spellings per field
**by hand**, which is a silently blank row per typo — and is why a stale `Year_x002f_Period` fallback
was still being carried long after the client renamed that column. One derivation replaces ~20
hand-written key pairs across the two screens.

## 4. What else was added to the panel

The complaint was about volume as well as the missing tiers, so:

- **Location** — the folder trail, first row. Shows `the library root` rather than being omitted when
  a file sits at the top, because an absent row reads as a file with no location.
- **Details** (`_ExtendedDescription`, the built-in Description) — was missing entirely from this
  panel although the approver's screen has always had it.
- **File size** — from `File/Length`, which is a **string of bytes**; a raw one reaches the screen as
  `1483776`, hence `formatBytes`. One decimal below 10, none above.
- **Last updated** — `Modified`. The gap between this and `Created` is the story on this page: created
  is when the uploader sent it, modified is when the approver acted.

`Uploaded by` was deliberately **not** added: on this page it is always the person reading it.

Caller-supplied rows (`leading`/`trailing`) pass through **blank or not**, which is why each says
`unknown` rather than being dropped — this screen can tell the difference between "no size" and "size
unreadable", and the shared module cannot.

## 5. The read got one more fallback

`File/Length` requires `$expand=File`, and a `$select` naming anything unavailable fails the **whole
request** with HTTP 400 rather than omitting a key (gotcha #11). Since these two fields exist only for
a richer panel, a **last-resort retry** was added that drops both and reads the set the page cannot do
without.

**A decorative panel row must never be the reason an uploader is told they have no files** — the
failure this page's design already guards against, because someone told they have nothing uploads the
file again and now there are two.

The empty state also moved: it now keys off `Object.keys(fieldText).length === 0`, not
`details.length === 0`. Location and the file facts always have a value, so `details` is never empty
and could no longer carry that state.

## 6. `ApprovalDocument` has the identical bug — NOT fixed here

`ApprovalDocument.tsx:704-705` hardcodes the same `Department` and `Unit`, so an approver looking at an
Upstream Operations file sees the segment and neither tier. It matters **more** there than on My
Submissions: the approver is deciding whether to publish into that unit, and `ApprovalDocument`'s own
comment says the full labels are shown precisely because the folder path is abbreviated and "an
approver needs the full label to be sure which unit they are publishing to". That guarantee is
currently not met for any segment outside the head offices.

Not migrated in this change because it is a different, site-verified screen and was not what was asked.
`buildDetailRows` was built to serve both, so the migration is small: replace the `metadata` array with
a `buildDetailRows` call whose `leading` row is the existing `orgLocation`. One difference has to be
decided when it happens: its `pick` returns `"—"` for a blank where this module drops the row.

## 7. Testing

1. Open the Upstream Operations file — the panel must now show **Region: Johor** and
   **Estate Mill: …** between Business Segment and Document Type.
2. Open a Group Head Office file — **Department** and **Unit** must still appear, in that order.
3. Confirm no GUID appears anywhere in the panel (the Tid columns must stay hidden).
4. Confirm **Location**, **File size** and **Last updated** are present and plausible; a rejected
   file's Last updated should be later than its Uploaded date.
5. Confirm `Document Date` still reads as SharePoint formatted it — never re-parsed (gotcha #1).
6. A file with no Project Name / Vendor must show **no** row for them, not an em dash.
