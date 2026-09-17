# Forced reindex of the CRS approved-side libraries, every 6 hours, via Power Automate

**No file calls this document** — a design spec under `docs/superpowers/specs/`, read before code.
**No data file is read or written by this document.** Confirmed via `Glob`: `*reindex*` matches
nothing; `*search*` matches `2026-08-16-document-search-design.md` (the CRS Search feature this
reindex serves, referenced not duplicated) and `2026-09-13-quick-search-remove-button-design.md`
(an unrelated People/Groups screen). No existing spec covers forcing a reindex.

**Client, 2026-09-17:** *"Archive is a problem in retrieving the file for c levels and global but
there is also problem with other libraries… index all."* Then, same day, on review of the runbook:
*"no need to crawl the Staging libraries and also lets crawl every 6 hours."*

**Status: NOT BUILT. Design only.**

---

## 1. The underlying, already-diagnosed problem

Already established live on this project (`2026-09-05` findings, recorded in `CLAUDE.md`): SharePoint
indexes a document at creation/change time, and if that initial crawl is missed for any reason,
**nothing automatically re-triggers it.** Measured coverage on one library: 23 of 264 items indexed.
Search results for a genuinely present document, and metadata filters over it, come back empty —
not because the code is wrong, but because the item was never crawled.

The one lever reachable without tenant admin access is **forcing a full recrawl of a list**, done by
flipping its `NoCrawl` property `true` then `false`. SharePoint's UI exposes this as *"Reindex
Document Library"* (Library settings → Advanced); the same effect is reachable via two REST `MERGE`
calls against the list, which is exactly the "Send an HTTP request to SharePoint" pattern every other
flow in this project already uses — no new connector, no new permission shape.

## 2. Honest limits — answering "does it make it faster" directly

**No — it does not make the crawl itself faster, and it is not instant.** Toggling `NoCrawl` marks
the WHOLE list as needing a fresh crawl; the actual re-crawl still happens on SharePoint's own crawl
schedule, which in SharePoint Online is near-continuous for most tenants but not synchronous with the
toggle. What this buys is **guaranteed eventual coverage** rather than a permanent gap — a document
that somehow missed its own crawl will be picked up by the next toggle, rather than never being
picked up again. Running it **every 6 hours** bounds the worst case to "about 6 hours," not
"forever" (tightened from the original nightly proposal, per the client's own follow-up).

Worth saying plainly so it isn't oversold: this closes the *coverage* gap (documents nobody's search
will ever find), not a *speed* gap (documents that just uploaded and aren't searchable yet — that
delay is normal and unrelated to this fix).

## 3. Scope — four libraries, Staging excluded

**Client, revised same day:** *"no need to crawl the Staging libraries."* So this now covers
`Documents`, `HC Documents`, `Archive`, `HC Archive` only — **not** `Approval Document` /
`HC Approval Document` (whichever of the HC pair actually resolve on a given site, same conditional
pattern every other HC-aware flow already follows).

⚠ **This exclusion lines up with a decision already made elsewhere, worth stating rather than
taking purely on instruction:** CRS Search stopped reading the approval libraries entirely on
2026-09-10 (`2026-09-10` entry, CLAUDE.md — a viewer with no role in Staging gets a security-trimmed
404 there, which was surfacing as a permanent, confusing "could not be searched" banner for most
users). So there is nothing in Staging that CRS Search would ever surface even if it were perfectly
indexed — reindexing it would cost two REST calls per run for zero benefit. Dropping it isn't just
a scope cut, it's removing dead weight.

⚠ **A flow cannot dynamically discover library titles the way this project's own `naming.ts`
probing does.** Each site's flow needs the four (or two, on a non-HC site) library identifiers
hardcoded into it at build time — the same constraint every existing flow in this project already
lives with, and the same reason the HC clones are separate flows rather than one parameterised one.
(As of §4, these are GUIDs rather than titles, so at least a future rename won't be one of the things
that needs re-hardcoding.)

## 4. The flow

New flow, `CRS — Library reindex` (working name — "Nightly" dropped from the name now that it runs
every 6 hours, not once a day). Trigger: **Recurrence**, every 6 hours (site's local time zone set
for correctness, though an interval-based recurrence is largely timezone-agnostic once the first run
has fired — `Singapore Standard Time` is what every other time-sensitive flow here uses, so it stays
the default for consistency).

⚠ **Addressed by list GUID, not by title** — decided after the first live verification query on SDG
(2026-09-17, exact GUIDs recorded in the runbook). A title-addressed flow breaks silently the next
time this client renames one of these libraries (already happened at least three times); a GUID
never changes across a rename. Since the flow has to be built with hardcoded values either way, GUID
is the more durable choice at no extra cost.

For each of the four library GUIDs:

```
PATCH _api/web/lists(guid'<Library Id>')   { "NoCrawl": true }
PATCH _api/web/lists(guid'<Library Id>')   { "NoCrawl": false }
```

Run sequentially per library (not parallel) to avoid throttling simultaneous list-property writes
against the same site in one run — matching this project's own established caution around
concurrent SharePoint REST calls (*"ten simultaneous calls is how a tenant starts returning 429s"*).
With four libraries this is now eight actions total (four, on a non-HC site) rather than twelve.

## 5. What this does NOT do

- **Does not target only what changed since the last run.** `NoCrawl` toggling has no per-item
  granularity — it is always a whole-list operation. A library with 10,000 items gets a full recrawl
  every 6 hours, forever. Accepted cost, per the client's explicit "index all," now run four times a
  day instead of once — a real increase in load on the four remaining libraries, worth watching if
  either of them grows very large.
- **Does not report success/failure anywhere visible.** Worth deciding whether this flow's own run
  history is the only place anyone would ever check it, or whether it should write a line to
  `CRS Audit Log` (a new `EventType`, following the same registration discipline every other event
  type in this project needs — added to `EVENT`, `EVENT_LABEL` and `ALL_EVENT_TYPES` together, or it
  is written correctly and simply invisible to the viewer's Action filter, as has happened twice
  before in this project with `Replaced`/`ShareRevoked`).

## 6. Open questions

1. **Audit visibility** (§5) — silent flow, or a logged event? Leaning toward logging one row per
   run (not per library) so an admin can eventually check "did this actually run" without opening
   Power Automate's own run history — matching this project's existing "one row per run" convention
   for reconciliation/migration, never one row per item. With four runs a day instead of one, this
   also decides whether the audit log gets four new rows a day forever, worth weighing against
   staying silent.
2. **Does the HC pair actually exist on every site this ships to?** If SDG's tenant hasn't had its HC
   pair ported/verified at the time this is built, the flow needs the same conditional shape every
   other HC-aware feature uses rather than hardcoding all four titles unconditionally.
