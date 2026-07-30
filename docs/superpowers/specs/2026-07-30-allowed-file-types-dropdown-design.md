# Allowed File Types — Dropdown Instead of Free Text — Design

**Date:** 2026-07-30
**Status:** Agreed, not yet implemented
**Touches:** `DMS Config` list schema, `Form.tsx`, `BulkUpload.tsx`, new `src/shared/allowedFileTypes.ts`

---

## Problem

`allowedExtensions` is a comma-separated string typed by hand into the shared `SettingValue`
text column. A typo silently breaks uploads with no error anywhere.

This is not hypothetical — it happened on 2026-07-30. The value read:

```
.pdf,.doc,.docx,.xls,.xlsx,png
```

`png` is missing its leading dot. The string is split on commas and passed verbatim into the file
input's `accept` attribute (`Form.tsx:1299`). Browsers only treat an `accept` token as a file
extension when it starts with `.`; a bare `png` is parsed as a MIME type, has no `/`, is invalid,
and is silently dropped. PNG files were greyed out in the file dialog, and since `Form.tsx` has no
drag-and-drop handler, that picker is the only way in — so PNG was unuploadable.

Two things made it expensive to diagnose:

1. **The gate is inconsistent.** `accept` rejected the bare `png`, but the JS validator uses
   `finalName.toLowerCase().endsWith(ext)` (`Form.tsx:829`, `Form.tsx:1306`), and
   `"photo.png".endsWith("png")` is `true`. So the validator would have accepted the very file the
   picker refused to show. One typo, two different behaviours.
2. **Failures are invisible.** `loadSettings().catch(() => DEFAULT_SETTINGS)` (`Form.tsx:659`)
   discards the error entirely. A misconfigured value and an unreadable list look identical from
   the outside, and the form renders as if hardcoded fallbacks were configured values. Diagnosis
   took four rounds of guesswork before the actual value was read off the API.

## Goal

Make the common act — the client choosing which file types are allowed — structurally
typo-proof, and make the remaining failure modes loud instead of silent.

## Non-goal

Removing config-driven extensions entirely. The client self-serves this setting occasionally, so it
stays live-editable without a redeploy.

## Design

### 1. Schema — one new column on `DMS Config`

| Property | Value |
|---|---|
| Internal name | `AllowedFileTypes` |
| Type | Choice, **allow multiple selections** |
| Fill-in choices | **Disabled** — re-enabling it reintroduces the exact typo this design removes |
| Choices | `.pdf` `.doc` `.docx` `.xls` `.xlsx` |
| Default | none |

Only the `allowedExtensions` row populates it. The other setting rows leave it blank, exactly as
they already leave `ModeLabel`, `Levels`, and `StagingFolder` blank.

The column already exists on `/sites/ClarenceDMSTesting` as of 2026-07-30 but is **unpopulated**.
Before implementation, verify its three properties above — multi-select on, fill-in off, and the
five choices each carrying a leading dot.

### 2. Who can change what

Deliberate split, and the reason a Choice column is the right shape here:

- **Which types are allowed** (ticking boxes on the item) — any client user with edit rights on the
  list. No typing, so no typo.
- **Which types are offerable** (the column's choices) — requires **Manage Lists**, i.e. site
  Owner or admin. Still involves typing the extension once into the Choices box, so the normalizer
  below is what keeps a fumbled dot from becoming another silent outage.

### 3. Read precedence

In `loadSettings`, first non-empty result wins:

1. `AllowedFileTypes` — the array from the Choice column
2. `SettingValue` — the existing free-text string
3. `DEFAULT_SETTINGS.allowedExtensions`

Keeping (2) means every already-seeded site keeps working the moment the code ships, whether or not
its column has been populated yet. Sites migrate by ticking boxes, on their own schedule.

**An empty Choice selection must fall through to (2), not resolve to an empty allowlist.** An empty
array means "not configured", never "allow nothing" — the latter would block all uploads site-wide
from an accidental untick.

### 4. Normalizer

Every path — Choice array, free text, and the hardcoded default — runs through one function:

```
trim → lowercase → prepend "." if missing → drop empties → de-duplicate
```

This is defence in depth, not the primary fix. It would have prevented the 2026-07-30 bug on its
own, and it covers the residual schema-edit typo surface from §2.

Lives in **`src/shared/allowedFileTypes.ts`** with unit tests, following the existing
`formModel.ts` / `pathEncoding.ts` pattern. Not duplicated into two web parts.

### 5. Deployment hazard: the `$select` trap

Adding `AllowedFileTypes` to the `$select` makes the whole request **400 on any site that does not
have the column yet**. Combined with the silent `.catch()`, that would drop *every* setting —
term-set GUIDs, `stagingLibrary`, column names — back to hardcoded defaults across the entire form.
A self-inflicted repeat of the bug this spec exists to fix.

So `loadSettings` **retries once without the new field** if the first call fails, and logs which
path it took. Column-then-code ordering then stops being a deployment requirement.

### 6. Diagnosability (agreed in scope)

The `.catch()` blocks at `Form.tsx:659-661` and the equivalents in `BulkUpload.tsx` log the real
HTTP status and response body instead of discarding the error. Same principle as CLAUDE.md
gotcha #9: log the actual status before assuming a data problem. A fallback to hardcoded values
must be visible in the console, not indistinguishable from success.

### 7. Default alignment

`DEFAULT_SETTINGS.allowedExtensions` is currently `[".pdf", ".xls", ".xlsx"]` (`Form.tsx:281`,
`BulkUpload.tsx:324`) — missing `.doc` and `.docx`. This mismatch is part of why the fallback was
confusing to diagnose. Align both to the agreed five.

## Files touched

| File | Change |
|---|---|
| `src/shared/allowedFileTypes.ts` | **new** — normalizer + precedence resolver |
| `src/shared/allowedFileTypes.test.ts` | **new** — unit tests |
| `src/webparts/form/components/Form.tsx` | `loadSettings` precedence + retry (line 506), `DEFAULT_SETTINGS` (line 281), `.catch` logging (line 659) |
| `src/webparts/bulkUpload/components/BulkUpload.tsx` | same, against its own `loadSettings` (line 681) and `DEFAULT_SETTINGS` (line 324) |

`FolderManager.tsx` reads `SettingValue` but not `allowedExtensions`, so it is unaffected.

## Testing

Unit tests on the normalizer:

- missing leading dot (`png` → `.png`) — the 2026-07-30 regression
- uppercase (`.PNG` → `.png`)
- stray whitespace (`" .pdf "` → `.pdf`)
- empty and whitespace-only entries dropped
- duplicates collapsed
- both input shapes: `string[]` from Choice, `string` from free text

Unit tests on precedence:

- populated Choice column wins over `SettingValue`
- **empty** Choice array falls through to `SettingValue`, and does **not** yield an empty allowlist
- absent `SettingValue` falls through to defaults
- resolved result is never empty

Manual verification on `/sites/ClarenceDMSTesting`:

- tick all five, hard-refresh, confirm the picker offers all five and a `.png` is refused
- untick `.pdf`, hard-refresh, confirm a PDF is rejected and greyed out in the dialog
- confirm the console names the read path taken, and names the HTTP status on a forced failure

Note that config is read once in a mount-time `useEffect`, so **every manual test needs a hard
refresh**. A stale page was half of the 2026-07-30 confusion.

## Rollout

1. Verify/complete the column properties (§1).
2. Populate the `allowedExtensions` row: tick `.pdf .doc .docx .xls .xlsx`.
3. Ship the code. `SettingValue` stays in place as the fallback; delete it only once the client site
   has the column populated and verified.

Order is not load-bearing thanks to §5, but this sequence means no window where the form runs on
defaults.

## Out of scope

- **Showing the accepted types in the form UI** next to the file picker. Proposed, not confirmed —
  worth revisiting, since today the client discovers the allowlist only by being rejected.
- `.ppt .pptx .txt .csv .jpg .jpeg` as offerable choices. CLAUDE.md documents 12 allowed types; the
  client's actual policy as of 2026-07-30 is these five. CLAUDE.md's "Allowed File Types" section
  should be corrected to match.
- **Housekeeping spotted in `DMS Config` on 2026-07-30, not part of this change:** the retired
  `termSet_vendor` row still holds `eaafd0e5-…` though the vendor term set was deleted and
  Vendor is now free text; and `Levels` JSON has been fill-dragged into the `col_businessSegment`
  and `col_businessSegmentTid` setting rows. Both are inert — `loadModes` filters on
  `ConfigType eq 'mode'` — but both are misleading to the next person reading the list.
