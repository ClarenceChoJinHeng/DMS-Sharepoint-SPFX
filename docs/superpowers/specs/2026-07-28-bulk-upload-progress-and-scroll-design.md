# Bulk Upload — scrollable file list, hidden Vendor, real byte progress

**Date:** 2026-07-28
**Scope:** `src/webparts/bulkUpload/components/BulkUpload.tsx` only
**Status:** implemented (build + lint + tsc green)
**Supersedes in part:** `2026-07-24-bulk-upload-two-batch-design.md` (Vendor as a panel
field), `2026-07-27-upload-approval-popup-polish.md` §3 (the progress panel this reworks)

## Why

Three complaints from testing the bulk upload tool:

1. The **Selected files** list paginated behind *Show more (N more)* / *Show less*. With
   a 50-file batch that meant repeated clicking to see what was queued.
2. The **Vendor** dropdown wasn't wanted in this tool for now.
3. The per-file progress bar wasn't progress at all — it was an **indeterminate
   looping stripe** (`@keyframes dms-slide`, a 40%-wide block sliding left→right on
   repeat). It conveyed "something is happening" but never how far along, and read as
   a bug ("a one way line that pass by again and again"). It also didn't match the
   agreed row layout.

## Changes

### 1. Selected files — scroll, not paginate

- Deleted `FILE_PAGE`, the `fileShowCount` state and both Show more / Show less buttons.
- `.dms-filelist` gains `max-height: 340px; overflow-y: auto; overscroll-behavior: contain`.
  The list renders `picked` in full and scrolls internally.
- Footer collapses to a single count: `27 files selected (50 max).`
- `.dms-progress-list` (the live panel) got the same treatment — `max-height: 360px`,
  `overscroll-behavior: contain`.

### 2. Vendor — hidden, not removed

Client asked to hide it "for now". Deliberately **not** ripped out:

- The `renderSelect("Vendor (if applicable)", …)` call is commented out in the panel.
- `vendor` / `setVendor` state, the `termSet_vendor` load, `Batch.vendorId/vendorLabel`,
  and the `formValues.push({ FieldName: settings.columns.vendor, … })` guard all stay.
- Because the field can no longer be set, `batch.vendorId` is always `""`, so the guard
  never fires and **nothing is written to the `Vendor` column**. Un-commenting the select
  is the only change needed to restore the old behaviour.

> Note the divergence from `Form.tsx`, which keeps Vendor as **free text** used only to
> compose the document name (`<Vendor>-<DD-MM-YY>`) and deliberately never writes it as
> metadata. Bulk upload has no auto-rename, so Vendor had no naming role here either way.

### 3. Real byte-level progress

`spHttpClient` is fetch-based and exposes **no upload progress events** — that is the
reason the bar was indeterminate in the first place. `XMLHttpRequest` still exposes
`upload.onprogress`, so the file POST (and only the file POST) moved off `spHttpClient`:

- New module-level `postFileWithProgress(url, file, digest, onPct)` — same
  `Files/Add(url='…',overwrite=…)?$select=ServerRelativeUrl` endpoint, same **raw `File`
  body** (never FormData — see CLAUDE.md gotcha #6), resolving
  `{ ok, status, body }` rather than throwing.
- Progress is **capped at 99%** during transfer; 100% is only set after `onload`, since
  the last stretch is SharePoint committing the file, which bytes-sent can't observe.
- `spHttpClient` attached `X-RequestDigest` for us; XHR does not. Added `fetchDigest()` /
  `getDigest(force?)` over `POST /_api/contextinfo`, caching the value against the real
  `FormDigestTimeoutSeconds` and refreshing a minute early. A **403 on upload retries
  once** with `getDigest(true)`, since that's what a rejected/expired digest looks like.
  - `fetchDigest` is the only writer of `digestRef` and never reads it first — keeps
    eslint's `require-atomic-updates` happy across the `await`.
- `digestRef` is a `useRef`, so a long two-batch run survives digest rotation.

Everything after the upload (`ListItemAllFields` fetch, `validateUpdateListItem`, the
`HasException` check per CLAUDE.md gotcha #4) still goes through `spHttpClient` unchanged.

### 4. Progress row layout

`LiveBatch.files` entries gained `size: number` and `pct: number`. Row order is now:

```
[STATUS]  filename.pdf   2.9 MB   [████████░░░░░]  72%   ✕
```

- Status tag moved to the **left** as a fixed 84px uppercase label; the row itself is
  tinted by state (`done` → `#eaf7f0`, `skipped`/`tagFailed` → `#fff4e5`,
  `failed` → `#fdf3f3`, `deleted` → `#f2f2f2`).
- `fmtSize()` renders MB above 1 MiB, else KB.
- The bar + `%` render **only while `state === "uploading"`**; a flex spacer holds the
  column otherwise, so settled rows are clean text. Marked up as
  `role="progressbar"` with `aria-valuenow`.
- `.dms-fp-fill` is width-driven off `pct` — see §4b for how that width is animated. The
  `@keyframes dms-slide` animation is **deleted** (`dms-slidein`, the toast animation,
  is a different rule and stays).
- Labels: `pending` → WAITING, `uploading` → UPLOADING, `done` → **UPLOADED** (was DONE).
- The per-row `✕` (post-hoc delete from Documents) is unchanged — still only offered for
  `done` / `tagFailed` rows that captured a `ServerRelativeUrl`.

### 4b. Making the fill actually animate (`UploadingBar`)

Byte-accurate progress turned out not to be enough on its own. The browser fires
`upload.onprogress` only a handful of times per request — first test run on a 2.1 MB PDF
reported **once**, so the bar snapped to 50%, sat frozen, then vanished into UPLOADED.
Technically correct, but it doesn't read as loading.

New module-private `UploadingBar` component owns the animation:

- A `requestAnimationFrame` loop maintains its own `shown` value:
  - `shown < target` → `shown += (target - shown) * 0.15`. Real measurements win, and
    the bar catches up to them within a few frames.
  - otherwise → creep toward `target + (99 - target) * 0.5`, at `* 0.006` per frame.
    That's the **midpoint of what's unknown**, approached asymptotically, so the bar
    keeps visibly moving without claiming more than half of the unverified remainder.
- Hard-clamped to 99. 100% is never shown by the animation — the row flips to UPLOADED
  only when the upload genuinely resolves.
- The loop writes `style.width` and the `%` text **through refs, not state**. Only one
  file uploads at a time, but 60 `setState` calls a second re-rendering a 50-row batch
  list would be pure waste.
- Mounted only while `state === "uploading"`, so each file's bar starts fresh at 0.
- `aria-valuenow` reports the **real** `target`, not the animated value — assistive tech
  gets the measurement, not the easing.

The `.12s linear` CSS transition on `.dms-fp-fill` was removed: a per-frame width write
plus a transition fight each other and produce visible lag.

### 5. Overall bar

`liveTotals` now tracks a third accumulator, `partial`, adding the in-flight file's
`pct / 100`. The headline bar therefore creeps forward continuously instead of jumping a
whole file at a time. Still clamped to 100.

## Metadata written (unchanged apart from Vendor)

Synthetic values; formats per CLAUDE.md:

| FieldName | Format | Example |
|---|---|---|
| `Document_x0020_Type` | taxonomy `Label\|GUID` | `Invoice\|866c5754-…` |
| `Year` | taxonomy `Label\|GUID` | `2026\|023a866a-…` |
| `Confidentiality_x0020_Level` | taxonomy `Label\|GUID` | `Internal\|0d6d1da8-…` |
| `DocumentDate` | **M/D/YYYY** (US site locale; ISO rejected) | `7/28/2026` |
| `Vendor` | taxonomy `Label\|GUID` | **no longer sent** (was `Acme Sdn Bhd\|eaafd0e5-…`) |
| `Business_x0020_Segment` / `…Tid`, `Department` / `DepartmentTid`, `Unit` / `UnitTid` | plain text label + GUID pair | `Group Finance` / `62d56a2e-…` |

## Verification

- `npx tsc --noEmit` — clean.
- `npx eslint` on the file — 0 errors. Two warnings remain, both pre-existing and
  untouched by this change (`eqeqeq` on an unrelated `!=`, and `max-lines`, which the
  file already exceeded).
- `npm run build` — packaged `sd-gatrie.sppkg` successfully.
- Not yet exercised against a live library; the digest + XHR path needs a real
  multi-file batch run on `/sites/ClarenceDMSTesting` to confirm progress ticks and that
  no 403-retry loop appears.
