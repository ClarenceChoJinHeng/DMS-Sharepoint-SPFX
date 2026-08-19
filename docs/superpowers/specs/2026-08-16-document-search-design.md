# CRS Search — searching the repository by metadata

**Date:** 2026-08-16
**Status:** BUILT — `documentSearch/components/DocumentSearch.tsx`, 72 tests. NOT site-tested, and §6's `OWSTEXT` managed-property assumption is still UNVERIFIED on a live site. Header corrected 2026-08-19.
**Web part:** `CRS Search` — `8b2f4a95-7c31-4d68-9e04-1a6b3f8c25d7`, own bundle, homepage.
**Rules module:** `src/shared/documentSearch.ts` (pure, tested).

---

## 1. What the client asked for

> "client wants us to add a search bar in homepage that allows them to search files via metadata
> and ensuring that the file is filtered by user, so only user who created that file can see it."

and then, when asked what a non-uploader should see, corrected it:

> "wait the search is ofcourse metadata but the security part follows the today's hierachy of how
> the group works. PIC sees only their own unit, HOU sees the unit, HOD sees the entire unit and
> segment and global you get the gist"

**The second statement governs.** The first would have produced an author-filtered list — which already
exists as My Submissions, and which would have shown a Head of Department nothing. What is actually
wanted is a search across the repository, scoped to what each person can already see.

That distinction is the whole design. It is recorded here because the two readings produce completely
different web parts, and the first one is the one written down in the original request.

---

## 2. The security model: we implement none of it

**Both engines are security-trimmed by SharePoint before any row reaches this code.**

- The Search API applies ACLs at query time — a result the caller cannot open is never returned.
- A REST list query returns only items the caller can read.

So the hierarchy the client described is not implemented here at all. It falls out of the folder ACLs
reconciliation already grants:

| Persona | Sees, without one line of code here |
|---|---|
| PIC | their unit's approved documents; their own pending/rejected |
| Head of Unit | their unit, including every pending file in it |
| Head of Department | every unit under the department |
| C-Level (segment) | that whole segment |
| C-Level (global) | everything |

**This is the safe direction, and it is why no role check belongs in this web part.** A bug in our
filtering can only ever return FEWER rows than the person is entitled to. A role check written here
could return more — and would be a second, drifting copy of a model that already has one source of
truth in `groupMapModel.ts` and the folder ACLs.

The approval-library half needs no special handling either: Draft Item Security (`DraftVersionVisibility
= 2`) already hides peers' drafts, so a PIC searching there gets their own pending files and a Head of
Unit gets the unit's.

**Consequence to state to the client, because it differs from their first sentence:** a PIC searching
finds *their unit's* approved documents, not only their own. That is the same visibility they already
have browsing the library, and the same fact recorded in
`docs/client/document-visibility-within-a-unit.md`. If they want an own-files-only view, that is My
Submissions, which already exists.

---

## 3. Two readers, and why the split is not arbitrary

| Reader | Libraries | Reason |
|---|---|---|
| **Search / KQL** | `Documents`, `HC Documents` | Grows without bound. No list view threshold, and full-text inside PDFs and Word files. |
| **REST `$filter`** | `Approval Document`, `HC Approval Document` | Small by construction — a file LEAVES on approval. Instant, no crawl. |

### 3.1 Why not `$filter` everywhere

The engine was first chosen as `$filter` when this was believed to be an own-files search. Scoped to the
hierarchy it is repository-wide, and there `$filter` has a hard ceiling:

- SharePoint throws the **5,000-item list view threshold** on any filter it cannot serve from an index.
- **`substringof` cannot use an index at all** — and contains-text is most of what a search box does.
- A list allows roughly 20 indexes, so indexing does not rescue the free-text case at any volume.

Past 5,000 items in `Documents`, free-text search would stop working *entirely* and return an error,
not fewer results. SDG's expected volumes are the same ones that made per-file ACLs unworkable against
the 50,000-scope ceiling (2026-08-06), so this is not hypothetical.

### 3.2 Why not Search everywhere

Crawl latency. A file uploaded five minutes ago is not findable — and the approval library is exactly
where a just-uploaded file sits, in front of the one person who most expects to find it. That reads as
a broken search, and no disclaimer fixes the impression.

### 3.3 The gap between the two, and the third read that closes it

**A file approved two minutes ago appears in NEITHER reader.** It has been moved out of the approval
library by Auto-route and deleted from the source (that delete is load-bearing for security, so it is
not negotiable), and it is not yet crawled into `Documents`. The window is the crawl interval —
minutes to hours.

Closed with a third, cheap read: `Documents` (and `HC Documents`) filtered `Modified ge <24h ago>`,
merged into the results and deduped by `UniqueId`.

- `Modified` is indexable and the result set is tiny, so this is threshold-safe at any library size.
- 24 hours is generous against a crawl measured in minutes; the cost of being generous is a few extra
  rows to dedupe, and the cost of being tight is the exact invisible gap this exists to close.
- It runs on every search, not only when the main read returns nothing — the gap is per-document, not
  per-query, so a search that already has results can still be missing a specific recent one.

**Dedupe is on `UniqueId`, never on path or name.** The same document legitimately appears in both the
Search result and the recency read, and its path may have changed between the crawl and now.

---

## 4. The filter panel

| Control | Matches | Source |
|---|---|---|
| Free text | filename, Project Name, Vendor/Customer, Remark — plus document CONTENTS on the Search half | typed |
| Document Type | `Document_x0020_Type` | term set `866c5754-…` |
| Year | `Year` | term set `023a866a-…` |
| Confidentiality | `Confidentiality_x0020_Level` | term set `0d6d1da8-…` |
| Segment | the mode row | DMS Config `mode` rows |
| Tier values | that segment's own tiers | **derived from the mode row's `Levels`** |
| Document Date | `DocumentDate` from/to | typed |
| Uploaded | `Created` from/to | typed |

### 4.1 The tier filters are DERIVED, and this is not a stylistic preference

A hardcoded `Department` / `Unit` pair is the bug already found and fixed once, in the document details
panel (2026-08-14): on Upstream Operations Malaysia, whose tiers are `Region` and `Estate/Mill`, both
rows read blank and the two values that decide where a file lives simply vanished.

The same failure here would be worse, because a filter that silently matches nothing looks like "there
are no such documents". Every future segment onboarded through `SegmentCreator` names its own tiers, so
a written-out list can only ever be right for the segments it was written for.

**Source: the `Levels` chain on the selected segment's `mode` row** — the same source the upload form
cascades over. Picking a segment reveals that segment's tiers and no others. With no segment selected,
tier filters are not offered at all, because "Department" means different things in different segments
and an unscoped tier filter would be a promise the query cannot keep.

Note this differs from `discoverTierFields`, which derives tiers from an ITEM's own Tid twins. There is
no item at search time, so the chain has to come from configuration. Both derive; neither hardcodes.

### 4.2 Free text is one box, not one box per field

The client asked for a search bar. Typing `tax` searches every free-text field at once and the contents
of documents. Per-field boxes would be a form, and a person searching from memory does not know which
field the word was in.

---

## 5. Managed properties — the one real setup cost

KQL can only filter on **managed properties**, not on column internal names.

**Assumption to verify before building the tier and metadata filters:** SharePoint Online
auto-creates queryable managed properties for list columns, in the form `<InternalName>OWSTEXT` for
text, `OWSDATE` for dates, and so on. If that holds, tier columns created by `SegmentCreator` become
searchable with **no configuration at all** — which is what makes §4.1 work for segments that do not
exist yet.

**Fallback if it does not hold:** map crawled properties to `RefinableString00`–`99` in
**Site Settings → Search Schema**. A site collection administrator can do this **without tenant
access** — the same constraint and the same route as the in-site term store pivot (2026-07-27), and
the client has already refused tenant-level access once.

**If neither works:** free text and the fixed columns still search; tier dropdowns are not offered.
Say so on the page rather than rendering dropdowns that match nothing.

**Whichever route, it is per-site setup**, so it belongs in the migration runbook alongside the column
diff. A new tier column added by the Structure Manager next year needs no action under the auto route
and needs a new mapping under the fallback route — that difference is the reason to verify rather than
assume.

---

## 6. Scoping the query

Both readers must restrict to this site and to the four libraries.

- KQL: `Path:"<web url>/<segment>"` per library, OR'd — built from `libraryUrlSegment()` and the
  resolved HC pair, **never from a literal**. A list's title and its URL are independent (gotcha #12),
  and `Documents` has the URL segment `Shared Documents` — the exact trap that put "Shared Documents"
  at the head of every folder trail in My Submissions.
- KQL: `IsDocument:true` to exclude folders. My Submissions listed 443 folders as submissions before
  `FSObjType eq 0` was added; a folder's link IS a library link, so clicking one appears to "open the
  document library" and the cause is invisible.
- REST: `FSObjType eq 0` — note the REST name `FileSystemObjectType` is rejected inside a `$filter`.

---

## 7. Highly Confidential

Included **only when both HC halves resolve** (`hcAvailable()`). Consistent with `hcRouting.ts`: there
is no HC fallback anywhere in `naming.ts`, and an unresolved HC target must 404 loudly rather than
resolve to the normal library.

- An uncleared user's HC read is refused by SharePoint. That is **expected, not an error** — it must
  not empty the page, block the search, or show a warning. It is simply fewer results, which is the
  correct answer.
- A cleared user's HC results are **marked as HC in the list**, so nobody copies a path out of a
  result without knowing which library it came from.
- A result's library travels with the row and is passed to the detail view. Item ids are per-list, so
  a row that does not carry its library opens the wrong document or 404s — the bug already fixed in
  My Submissions when HC exposed it.

---

## 8. Failure states

Three, distinguished, as everywhere else in this codebase:

1. **Nothing searched yet** — the resting state. Not "no results".
2. **Nothing matched** — the query ran and the repository has nothing.
3. **Could not read** — one or both engines failed. Names which, and the HTTP status.

A person told "no results" when a library was unreachable concludes the document is not there. In My
Submissions that made someone upload a second copy; here it would make them conclude a record does not
exist. **A partial failure reports partial results and says so** — it never renders as a clean empty.

An HC refusal for an uncleared user is not a failure and is not counted as one (§7).

---

## 9. Clicking a result

Opens the detail view **in the page**, reusing `buildDetailRows` from `shared/documentDetails.ts` and
`previewTarget` from `shared/filePreview.ts`.

Not a third copy of that panel: it has been written twice already, and both copies carried the
hardcoded-tier bug. `ApprovalDocument.tsx:704-705` still does and is fixed separately.

Metadata comes from `FieldValuesAsText`, which is a **per-item** endpoint — so the list itself cannot
show metadata without one request per row, and metadata lives only in the detail view. That is also
what turns a taxonomy value into a readable label instead of a lookup id.

---

## 10. What this is not

- **Not a permission boundary.** It shows what the person can already open. It grants nothing and
  hides nothing that browsing would not also show.
- **Not My Submissions.** That page is author-filtered and stays that way. This one is not.
- **Not a replacement for the library views.** It finds documents; the libraries remain how a unit
  browses its own filing.
- **Not full-text on pending files.** The approval-library half is `$filter`, so it matches metadata
  and filename, not document contents. Deliberate: instant matters more there, and a pending file is
  usually one the searcher just uploaded and can name.

---

## 11. Out of scope

- Saved searches, search history, export of results.
- Refiners with counts (needs `RefinableString` mappings and sortable managed properties).
- Searching the audit log, the config lists, or site pages.
- Any change to who can read what.
