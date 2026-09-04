# CRS — Archive test run, fixed cutoff (Crystal's test) · addendum

**Date:** 2026-09-02
**Base runbook:** `2026-08-22-seven-year-archive-mover-runbook.md` — read that one FIRST. Everything
not listed here is unchanged from it: path derivation, the clash-rename rule, the `MoveTo` call, the
`Archived` stamp, `resetroleinheritance`, and §6's "must not be extended to" list.
**Purpose:** Crystal wants to test the mover now, without waiting until 2033. Client's own words,
relayed by Clarence: *"ALL FILES that was uploaded before 1 September 2026, 00:00:00 push it to
archive."*
**Status:** not built. This file lists the six deltas from the base runbook; everything else in that
document applies unchanged.

---

## 0. Why an addendum, not a rewrite

The base runbook is correct and detailed — six months of this project's own hard-won gotchas are
folded into it (the `Shared Documents` URL-segment trap, `FSObjType` vs `FileSystemObjectType`, the
parameter-alias requirement on `moveto`, item ids being per-list). Duplicating it risks the two
copies drifting; this file only states what's DIFFERENT for a deliberate bulk test run.

## 1. Delta — the library TITLES are stale in the base runbook

Both archive libraries were renamed on 2026-08-28 (CLAUDE.md: "ALL SIX LIBRARIES RENAMED AGAIN").
**The URL segments are unchanged** — `/Archive/` and `/HCArchive/` still work exactly as the base
runbook's `DestPath` logic (§2.4a) already assumes, since that logic is built on URL segments, not
titles. Only the `getbytitle(...)` calls — which take the TITLE — need correcting:

| Base runbook says | Now use |
|---|---|
| `getbytitle('Archive')` (§2.4e, the `Archived` stamp) | `getbytitle('Archive Restricted %26 Confidential Document')` |
| `getbytitle('HC Archive')` (§3, the HC clone's stamp) | `getbytitle('Archive Highly Confidential Document')` |

⚠ **The `%26` is the URL-encoded `&`** in the title — required in a REST URI the same way it was
required for the verification queries earlier today. A raw `&` in a Power Automate URI field can be
mangled by the designer's own editor; encode it explicitly.

**Nothing else changes.** `Create new folder` (§2.4b) and `Get file metadata using path` (§2.4c) both
pick their library from a DROPDOWN in the designer, not a typed title — reselect `Archive Restricted &
Confidential Document` / `Archive Highly Confidential Document` from the list, and the connector
handles the correct internal reference itself.

## 2. Delta — the CutOff is a fixed date, not a rolling 7-year calculation

Base runbook §2.2's `Compose` (`CutOff`) becomes a literal instead of `addDays(utcNow(), -2557, ...)`:

```
2026-09-01T00:00:00Z
```

Still **ISO**, still feeds the same `$filter` in §2.3 (`Created lt datetime'@{outputs('CutOff')}' and
FSObjType eq 0`) unchanged. Nothing downstream of the Compose needs to know the value is fixed rather
than computed.

⚠ **This is deliberately named a TEST cutoff, and it should read that way in the flow itself.** Name
the Compose action `CutOff_TEST_20260901` rather than reusing `CutOff` — if this flow is ever cloned
into the real scheduled mover, a stray literal silently surviving a copy-paste is exactly the kind of
defect that reads as "why did this stop archiving new files after September."

## 3. Delta — the 100-file cap is REMOVED for this run

Base runbook §2.3 caps `Top Count` at 100 and §2.4/§6 explain why: 100/week protects Auto-route's
shared daily flow quota on a **recurring schedule**. This is a one-off manual test, not the schedule,
so the cap is dropped — **confirmed with the client**, given the oldest document on site predates the
cutoff by only two months, meaning this will move essentially every current document in `Documents`
and `HC Documents` into their archive counterparts in one run.

⚠ **TWO SEPARATE SETTINGS CONTROL THE ROW LIMIT, and missing the second one silently caps the run
anyway.** `Top Count` in the action's own fields is one; **Settings → Pagination**, OFF by default, is
the other — with Pagination off, `Get items` returns at most the first page (up to `Top Count`, itself
capped at 5000) REGARDLESS of what `Top Count` says below that. Turn **Pagination ON** and set
**Threshold** to something comfortably above the library's current item count (e.g. `5000`) so the
action actually returns everything eligible, not just the first page.

⚠ **PUT THE CAP BACK before this flow (or a clone of it) ever runs on a schedule, and definitely
before it reaches SDG's tenant.** The removal is specific to this one deliberate test run on
ClarenceDMSTesting. Re-read base runbook §6 before reusing this build for anything recurring.

## 4. Delta — trigger is "Manually trigger a flow", not Recurrence

Base runbook §2.1 uses a weekly Recurrence trigger, fired via the designer's Run button for testing.
Since this build's only purpose right now IS the manual test — no schedule is wanted yet — start with
an **instant cloud flow** ("Manually trigger a flow", no inputs needed). This is what lets Crystal (or
Clarence) press Run directly from Power Automate's Run history / the mobile app, without needing to
open the designer.

If this build is later promoted into the real recurring mover, swap the trigger to Recurrence at that
point — do not build both into one flow with a bypass toggle; that is exactly the kind of "test
override left behind in the real thing" this addendum's §2 warns about for the cutoff date.

## 5. Delta — test order, for THIS run specifically

Base runbook §5 builds a careful "one file first" procedure specifically to avoid ever running the
real bulk case by accident. That caution does not apply here — the bulk case IS the point — but the
verification steps after it (§5.3 points 2–6: My Submissions shows the `Archived` badge and refuses
deletion/share, CRS Search finds the file allowing for crawl latency, a PIC sees it read-only, the HC
pair behaves the same and refuses uncleared uploaders) all still apply and should be checked against a
SAMPLE of the moved files, not just one, given the scale of this run.

**Skip §5.1's filename-only smoke test** (the bulk cutoff already covers it) but **keep §5.2** — run
once, read the `CutOff` Compose output, confirm it reads `2026-09-01T00:00:00Z` exactly before letting
the loop run over real files.

✅ **CONFIRMED WITH THE CLIENT (2026-09-02): NO REVERT NEEDED.** *"It can be stayed as archive, we are
still on testing site."* The archived files stay in `Archive`/`HC Archive` after this test — do not
build a revert step, and do not treat the archived state as temporary when reasoning about later
tests on this site. (There is still no built-in undo mechanism in this project if that decision ever
changes — base runbook §6 is explicit that reversing an archive is a manual, file-by-file job.)

## 6. Nothing else changes

Path derivation (§2.4a), the ensure-folder step and its "safe only below the Unit" warning (§2.4b),
the clash-rename-by-year logic (§2.4c), the `moveto` call with its parameter-alias requirement
(§2.4d), the `Archived` stamp's `HasException` check (§2.4e), the HC clone's four reference changes
(§3), and `resetroleinheritance` (§4.5) are all UNCHANGED from the base runbook. Build against that
document; use this one only for the six deltas above.
