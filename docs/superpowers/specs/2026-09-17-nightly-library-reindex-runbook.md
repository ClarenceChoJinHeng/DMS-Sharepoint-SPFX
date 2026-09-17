# Library reindex — flow runbook (every 6 hours, Staging excluded)

**No file calls this document** — a Power Automate build runbook under `docs/superpowers/specs/`,
the project's established place for flow build steps (matching `2026-08-25-approver-notification-
flow-runbook.md` and every other `*-runbook.md` in this folder). **No data file is read or written
by this document itself.**

**User's current instruction, verbatim (this session):** *"Ok we can start with 4"* — item 4 being
the reindex flow from the earlier ordering discussion, design agreed in
`2026-09-17-nightly-library-reindex-design.md`. Revised the same day: *"Need some updates, no need
to crawl the Staging libraries and also lets crawl every 6 hours."*

**Date:** 2026-09-17
**Status:** NOT BUILT. This is the build runbook — read it alongside the design doc, which carries
the reasoning; this carries the exact steps.
**Applies to:** one new flow, `CRS — Library reindex`.
**Touches:** no code. Power Automate only, same as every other flow in this project.

---

## 0. Resolve the library GUID for each site FIRST — address lists by GUID, never by title

⚠ **Addressing changed from `getbytitle('<Title>')` to `Lists(guid'<Id>')` after the first verification
run, deliberately.** This client renames these exact libraries routinely (recorded at length in
CLAUDE.md — six libraries were renamed as recently as 2026-09-08, and again) — a title-addressed flow
breaks silently the next time that happens and needs manual repair; a GUID never changes across a
rename. Since a Power Automate flow has to be built with hardcoded values either way (it cannot
dynamically discover a library the way this project's own `naming.ts` probing does), GUID is strictly
the more durable of the two choices for zero extra cost.

**Run this on the target site before building anything**, from a browser console or the address bar
while signed into the site:

```
<site>/_api/web/lists?$select=Id,Title&$filter=BaseTemplate eq 101
```

`BaseTemplate eq 101` filters to document libraries only. Match the returned titles against
`Documents`, `HC Documents`, `Archive`, `HC Archive` (whatever their CURRENT display names actually
are on this site) and record each one's `Id`. **`Approval Document` / `HC Approval Document` are
deliberately excluded** (client, 2026-09-17: *"no need to crawl the Staging libraries"* — CRS Search
stopped reading them entirely on 2026-09-10, so there is nothing there for a reindex to help). Do
this separately per site — GUIDs are never shared across sites even for a library with the same
title.

### Confirmed on SDG's tenant (`/sites/CRS`), 2026-09-17

| Logical | Title | GUID |
|---|---|---|
| Documents | Restricted & Confidential Document | `fbc062dd-fcd0-4458-9161-2bfc19c917fa` |
| HC Documents | Highly Confidential Document | `122a4aa9-65fb-479a-90a7-3866d19f51b4` |
| Archive | Archive Restricted & Confidential Document | `647fd644-ab00-4358-9b3f-49031083fcf0` |
| HC Archive | Archive Highly Confidential Document | `fc301054-abf5-4e53-957a-6f8c58a410b8` |

Also present but irrelevant to this flow: `Approval for Document`, `Approval for Highly Confidential
Document` (Staging, excluded), `Form Templates`, `Site Assets`, `Style Library` (not CRS document
libraries).

**Not yet run for ClarenceDMSTesting** — do the same lookup there before building that site's copy of
the flow; do not assume its GUIDs match SDG's, since they never do across sites even when the titles
happen to be identical.

## 1. The flow

**Name:** `CRS — Library reindex`
**Trigger:** Recurrence
- Interval: `6`, Frequency: `Hour`
- Time zone: `Singapore Standard Time` (confirmed correct for both current sites elsewhere in this
  project's flow history — re-verify against the site's own Regional Settings if this ever ships
  somewhere else; largely academic for an hours-based interval once the first run has fired, but set
  it for consistency with every other time-sensitive flow here)
- Start time: any time — the interval is what matters, not the exact clock minute it first fires

**Connection:** signed in as `crs@sdguthrie.com` (or whichever service account this tenant's other
flows already use) — no elevated permission is needed beyond ordinary Contribute/Manage Lists on
each library, since toggling `NoCrawl` is a list-property write, not a security change.

## 2. One action pair per library, run SEQUENTIALLY

⚠ **Do not run these in parallel.** Simultaneous list-property writes against one site risk
throttling (429s) — this project's own established caution: *"ten simultaneous calls is how a tenant
starts returning 429s."* Build them as ordinary sequential steps, each `runAfter` the one before it
succeeding, not a parallel branch.

For **each** of the (up to four) library GUIDs confirmed in step 0 — `Documents`, `HC Documents`,
`Archive`, `HC Archive`, **never** `Approval Document` / `HC Approval Document` — two consecutive
**"Send an HTTP request to SharePoint"** actions, addressed by GUID:

```
Site Address: <this site>
Method: POST
Uri: _api/web/lists(guid'<Library Id>')
Headers:
  Accept: application/json;odata=verbose
  Content-Type: application/json;odata=verbose
  X-HTTP-Method: MERGE
  IF-MATCH: *
Body:
{
  "__metadata": { "type": "SP.List" },
  "NoCrawl": true
}
```

On SDG, this is four literal URIs — `_api/web/lists(guid'fbc062dd-fcd0-4458-9161-2bfc19c917fa')`,
`_api/web/lists(guid'122a4aa9-65fb-479a-90a7-3866d19f51b4')`,
`_api/web/lists(guid'647fd644-ab00-4358-9b3f-49031083fcf0')`,
`_api/web/lists(guid'fc301054-abf5-4e53-957a-6f8c58a410b8')` — each with its own toggle-true /
toggle-false pair.

then immediately after (same shape, `NoCrawl: false`):

```
Body:
{
  "__metadata": { "type": "SP.List" },
  "NoCrawl": false
}
```

⚠⚠ **NONE OF THIS HAS PRECEDENT IN THIS PROJECT — checked directly before writing this, not assumed.**
Grepping every existing `X-HTTP-Method: MERGE` call in this codebase (18 hits, across `spAuditLog.ts`,
`Requests.tsx`, `FileTypeSettings.tsx`, `spSegmentRecode.ts`, `MySubmissions.tsx`, and every runbook
for an existing Power Automate flow) turns up nothing that updates a LIST's own property. Every
single one updates a list ITEM. This is genuinely new territory here — verbose + `__metadata` above
is the textbook Microsoft REST shape for a list-property update (the same mechanism the "Reindex
Document Library" button in Library settings triggers under the hood), not something re-derived from
this project's own tested history. **Verify against a real run before trusting it** — if it 400s,
read the response body for the actual complaint rather than guessing a second shape blind (per this
project's own standing rule: SharePoint names the real cause in the body).

⚠⚠ **THE `odata-version: ""` FIX BELOW IS DELIBERATELY *NOT* CARRIED OVER FROM THIS PROJECT'S OWN
HISTORY, AND THAT IS THE POINT.** Every one of this project's six-plus "verbose collides with an
injected `odata-version: 4.0` header" incidents comes from **SPFx's own `SPHttpClient`**, in
application code — that header injection is a documented property of SPFx's client, not of Power
Automate's "Send an HTTP request to SharePoint" connector. I found no evidence anywhere in this
project's flow history that Power Automate injects the same header. So: **do not add
`odata-version: ""` pre-emptively** — if the first attempt genuinely comes back with *"Parsing JSON
Light feeds or entries in requests without entity set is not supported"* (the specific symptom that
header caused in SPFx code), THEN it's worth trying the same fix here; but treat it as a live
diagnosis, not a known cure carried over from a different codepath.

Repeat this pair for every confirmed library GUID from step 0 — up to eight actions total (four
libraries × two toggles), all in one straight sequential chain, no branching needed.

## 3. What "it worked" looks like

- **Immediately**: the flow's own run history shows eight successful actions (or four, on a
  non-HC site), each a `200`/`204`, four times a day.
- **The `NoCrawl` toggle itself proves nothing about search results changing right away** — see the
  design doc §2 for why this is a coverage fix on a roughly 6-hour horizon, not an instant one.
- **To actually confirm it worked**, re-run the coverage check this project already has a script for
  (`scripts/dump-segment-folders.js`'s sibling coverage-checking approach, or a plain count of
  indexed vs. actual items via `_api/search/query` against `path:"<library URL>"` vs the library's
  own item count) a day or so after this has run a few times, and confirm the gap has closed rather
  than assuming it from the flow's green run history alone.

## 4. Deliberately not built into this flow (see design doc §5 for why)

- No per-item targeting — every run reindexes the WHOLE library, four times a day, forever. Accepted
  cost, and a real increase in crawl load compared to the original once-nightly proposal — worth
  watching if `Documents`/`HC Documents` grow very large.
- **`Approval Document` / `HC Approval Document` are excluded on purpose** — CRS Search has not read
  them since 2026-09-10, so there is nothing there for this flow to help find.
- No audit-log row. Purely a Power Automate flow with no application code touched. If visibility
  inside `CRS Audit Log` is wanted later, that is a small follow-up (a new `EventType`, registered in
  `EVENT`/`EVENT_LABEL`/`ALL_EVENT_TYPES` together — the exact place two earlier event types,
  `Replaced` and `ShareRevoked`, were silently invisible to the viewer's filter for having been added
  to the flow but not to that registration).
