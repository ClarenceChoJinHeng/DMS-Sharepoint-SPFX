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

### 3. `AllowedFileTypes` is the single source of truth

`SettingValue` is **not** a fallback for this setting. There is exactly one place to read and one
place to fix. Two states only:

| State | Behaviour |
|---|---|
| `AllowedFileTypes` has ≥1 tick | Those types are the allowlist. Normal operation. |
| `AllowedFileTypes` is empty | **Hard block.** No file can be selected or uploaded. |
| `DMS Config` read fails entirely | `DEFAULT_SETTINGS.allowedExtensions` + loud console log. |

**Empty means empty.** An unticked column blocks all uploads and shows an actionable error rather
than quietly substituting some other set. Decided 2026-07-30: an empty allowlist is a
client-authored state, the client is the one who can fix it in one click, and silently papering
over it is how the 2026-07-30 incident stayed alive for four rounds.

**The read-failure tier is deliberately not a hard block.** When the whole list read fails
(403 / 404 / network), there is no data — `SettingValue` cannot help either, since it lives in the
list that just failed. Taking every upload down site-wide for a transient network blip is neither
the client's fault nor something they can fix by ticking a box, so this tier degrades to code
defaults and logs loudly. §7 aligns those defaults to the same five, so even the last resort agrees
with the dropdown.

**The two failures must not look alike.** "No file types are configured" and "couldn't read the
configuration" require different people to fix them. Each gets its own distinct message; neither is
allowed to render as ordinary operation.

### 3a. Consequence: every site needs the column

With no `SettingValue` bridge, a site whose `DMS Config` lacks `AllowedFileTypes` falls straight to
code defaults. Adding the column becomes a **required provisioning step** for any new or migrated
site, not an optional upgrade. Add it to `2026-07-28-client-site-migration-runbook.md`.

### 4. Normalizer

Both remaining inputs — the Choice array and the hardcoded default — run through one function:

```
trim → lowercase → prepend "." if missing → drop empties → de-duplicate
```

This is defence in depth, not the primary fix. It would have prevented the 2026-07-30 bug on its
own, and it still earns its place after the dropdown lands: per §2, an owner editing the column's
**Choices** types the extension by hand, so a fumbled dot is still possible one level up. The
normalizer repairs it at read time instead of letting it silently grey out the picker.

Note the normalizer alone cannot make an empty list non-empty — §3's hard block is a separate,
explicit check, not an emergent property of normalization.

Lives in **`src/shared/allowedFileTypes.ts`** with unit tests, following the existing
`formModel.ts` / `pathEncoding.ts` pattern. Not duplicated into two web parts.

### 5. Deployment hazard: the `$select` trap

Adding `AllowedFileTypes` to the `$select` makes the whole request **400 on any site that does not
have the column yet**. Combined with the silent `.catch()`, that would drop *every* setting —
term-set GUIDs, `stagingLibrary`, column names — back to hardcoded defaults across the entire form.
A self-inflicted repeat of the bug this spec exists to fix.

So `loadSettings` **retries once without the new field** if the first call fails, and logs which
path it took. Column-then-code ordering then stops being a deployment requirement.

This survives the §3 simplification. It is not about extensions — it protects the *other eight*
setting rows on the same request, which would otherwise be collateral damage from one unknown
field name.

### 6. Diagnosability (agreed in scope)

The `.catch()` blocks at `Form.tsx:659-661` and the equivalents in `BulkUpload.tsx` log the real
HTTP status and response body instead of discarding the error. Same principle as CLAUDE.md
gotcha #9: log the actual status before assuming a data problem. A fallback to hardcoded values
must be visible in the console, not indistinguishable from success.

Three states, three distinct user-visible outcomes — no two may be confusable:

| State | User sees |
|---|---|
| Types configured | Normal form. Picker offers those types. |
| `AllowedFileTypes` empty | Picker disabled + "No file types are configured for upload. Ask your DMS administrator to set Allowed File Types in DMS Config." |
| `DMS Config` unreadable | Normal form on code defaults + visible warning that configuration could not be loaded, and the HTTP status in the console. |

The middle message names the column and the list on purpose: the client is the person who fixes it,
so the message must tell them exactly where to click.

### 7. Default alignment

`DEFAULT_SETTINGS.allowedExtensions` is currently `[".pdf", ".xls", ".xlsx"]` (`Form.tsx:281`,
`BulkUpload.tsx:324`) — missing `.doc` and `.docx`. This mismatch is part of why the fallback was
confusing to diagnose. Align both to the agreed five.

## Files touched

| File | Change |
|---|---|
| `src/shared/allowedFileTypes.ts` | **new** — normalizer + resolver returning one of the three §3 states |
| `src/shared/allowedFileTypes.test.ts` | **new** — unit tests |
| `src/webparts/form/components/Form.tsx` | `loadSettings` read + retry (line 506), `DEFAULT_SETTINGS` (line 281), `.catch` logging (line 659), disabled-picker + message for the blocked state (line 1296) |
| `src/webparts/bulkUpload/components/BulkUpload.tsx` | same, against its own `loadSettings` (line 681), `DEFAULT_SETTINGS` (line 324), file input (line 2155) |

The resolver returns a **state**, not just an array — the two web parts need to tell "these are the
types" apart from "there are none" apart from "we don't know". Returning a bare `string[]` would
collapse the last two into an empty array and re-create the ambiguity §6 exists to remove.

`FolderManager.tsx` reads `SettingValue` but not `allowedExtensions`, so it is unaffected.

## Testing

Unit tests on the normalizer:

- missing leading dot (`png` → `.png`) — the 2026-07-30 regression
- uppercase (`.PNG` → `.png`)
- stray whitespace (`" .pdf "` → `.pdf`)
- empty and whitespace-only entries dropped
- duplicates collapsed
- input arrives as `string[]` from the Choice column

Unit tests on resolution:

- populated `AllowedFileTypes` is used verbatim (after normalization), and `SettingValue` is
  **ignored even when present and different** — this is the regression test for §3
- empty array resolves to the blocked state, **not** to defaults and **not** to an empty allowlist
  silently treated as permissive
- a failed read resolves to defaults, and is reported as a *different* state from empty

Manual verification on `/sites/ClarenceDMSTesting`:

- all five ticked, hard-refresh: picker offers all five, a `.png` is refused
- untick `.pdf`, hard-refresh: a PDF is rejected and greyed out in the dialog
- set `SettingValue` to something contradictory (e.g. `.png`) and confirm it has **no effect** —
  proves single source of truth
- untick everything, hard-refresh: picker disabled, §6 message shown, no upload possible
- force a read failure and confirm the console names the HTTP status, and that the UI shows the
  *unreadable* warning rather than the *empty* message

Note that config is read once in a mount-time `useEffect`, so **every manual test needs a hard
refresh**. A stale page was half of the 2026-07-30 confusion.

## Rollout

1. Verify/complete the column properties (§1). **Done and verified 2026-07-30:**
   `FillInChoice: false`, five choices, multi-select on.
2. Populate the `allowedExtensions` row: tick `.pdf .doc .docx .xls .xlsx`. **Done 2026-07-30.**
3. **Before the code ships**, set `SettingValue` on that row to `.pdf,.doc,.docx,.xls,.xlsx`. It is
   still the live setting until deploy, and currently reads `.doc,.docx,.png` — so right now the
   form rejects PDF and Excel and accepts PNG.
4. Ship the code.
5. **After** verifying, clear the `SettingValue` cell on the `allowedExtensions` row. Keep the row
   — its `Title` is the lookup key and it carries the ticks. Leaving a stale value in a cell nothing
   reads is exactly the `termSet_vendor` problem noted below.

Order is not load-bearing thanks to §5, but this sequence means no window where the form runs on
the wrong set.

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
