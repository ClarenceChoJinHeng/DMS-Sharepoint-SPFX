# Per-file reminder cadence — flow runbook

**Date:** 2026-09-22
**Status:** DESIGNED, NOT BUILT. Power Automate only — no code change anywhere in this repo.
**Applies to:** `CRS — Reminding Approver to Approve` (the normal 3-day reminder) and
`CRS — HC approval reminder` (its HC twin).
**Read this before opening either flow** — it changes the trigger frequency, the `GetStale` filter,
and adds one new action to each. Build one flow, test it, THEN clone the same three changes onto
the other.

---

## 1. What is actually live today, read from the real export

Extracted `PowerAutomateFlowsSDG/CRS—RemindingApprovertoApprove_20260918082053.zip` directly (not
reasoned from an earlier runbook) to confirm the exact current shape before designing on top of it:

```
Recurrence: frequency "Day", interval 1, startTime "2026-08-31T01:00:00.000Z",
            timeZone "Singapore Standard Time"   ← fires once a day, at 9 AM SGT
GetStale ($filter):
  FSObjType eq 0 and OData__ModerationStatus eq 2
    and Created lt '<now minus 3 days>'
Apply_to_each → GetApproverGroup → HasApprover → GetMembers → MembersWithEmail
  → HasRecipients → Apply_to_each_1 → Send an email (V2)
```

**There is no per-file memory of any kind.** `GetStale` re-evaluates fresh on every run, so a
document that is still Pending and older than 3 days matches EVERY single day, forever, until it is
decided. That is what CLAUDE.md already calls *"one gentle nag per day until resolved"* — a
deliberate design, not a bug — but it is why the client's timing observation is correct: the ONLY
thing that decides when an email goes out is the fixed 9 AM daily trigger. It has never had any
relationship to when a given file was actually uploaded.

## 2. What the client wants instead

> *"I think they are expecting those files to be sent in a 3 days cycle but during night time on the
> day it was uploaded."*

Confirmed with Clarence: **reminder timed to each file's own upload time**, not a single global
daily slot. Read literally against the observed screenshot (files uploaded 15:30–16:17, i.e.
afternoon/evening, and the client expecting the reminder to land at "night time" on those same
days) — the simplest, most faithful design is:

- First reminder: **exactly 3 days after the document's own `Created` timestamp**, at the same time
  of day it was originally uploaded.
- If still undecided, remind again **3 days after that**, and so on — a genuine repeating 3-day
  cycle anchored to the file, not to the clock.

Because most uploads in practice happen in the afternoon/evening (per the screenshot), this design
naturally produces "reminders at night" **without hardcoding a night-time window** — the reminder
simply inherits whatever hour the file was uploaded at. If the client instead wants a HARD night-only
window (e.g. never before 8 PM regardless of upload hour), that is a different, smaller addition —
see §6.

## 3. Why this needs a stamped column, and why the trigger must go hourly

Power Automate's `Recurrence` trigger fires the WHOLE flow at one fixed moment — it cannot spawn a
different due-time per row on its own. The only way to track "this specific file is due for a
reminder at this specific future moment" is to **store that moment on the item itself**, and poll
often enough to catch it close to when it actually arrives.

- **New column: `NextReminderAt` (Date and Time), on `Approval for Document` AND
  `Approval for Highly Confidential Document`.** Create this in SharePoint library settings
  directly (Library settings → Create column → Date and Time) — no code touches this list for
  reminders, so this is a plain manual SharePoint step, not a script.
  - ⚠ After creating it, open `GetStale` in the flow and re-select the list (or use "Add new
    parameter" / refresh dynamic content) so the connector's cached schema picks up the new field —
    a freshly created column sometimes needs this before it appears as filterable.
- **Recurrence frequency changes from `Day` to `Hour`, interval `1`.** This is what lets the flow
  catch a file's own due-time to within about an hour, whatever hour that is — the design deliberately
  does NOT restrict polling to a night window, because the due-time itself already encodes the right
  hour (see §2). Keep `timeZone: "Singapore Standard Time"` unchanged; it only affects how a
  `Day`-frequency start time resolves and has no effect once frequency is `Hour`.
  - Cost: 24 runs/day instead of 1, per flow — negligible on typical Power Automate plans, but worth
    saying out loud since it is a real increase in run volume. Applies to BOTH flows once cloned.

## 4. The `GetStale` filter — two cases, combined with OR

A file is due for a reminder in exactly one of two states: it has NEVER been reminded and is now
3+ days old, or it HAS been reminded before and its stored next-due time has arrived.

```
FSObjType eq 0 and OData__ModerationStatus eq 2 and (
  (NextReminderAt eq null and Created le '@{formatDateTime(addDays(utcNow(), -3), 'yyyy-MM-ddTHH:mm:ssZ')}')
  or
  (NextReminderAt ne null and NextReminderAt le '@{formatDateTime(utcNow(), 'yyyy-MM-ddTHH:mm:ssZ')}')
)
```

Everything after `GetStale` — `GetApproverGroup`, `HasApprover`, `GetMembers`, `MembersWithEmail`,
`HasRecipients`, the email itself — is **completely unchanged**. The approver lookup, the recipient
cap, the email body, the link — none of that needs touching.

## 5. The one new action: advance the file's own clock after each check

Add an **Update item** action, sitting AFTER `HasApprover`'s if-block finishes (both its
Succeeded AND Failed paths — the clock must advance whether or not an approver group was actually
found, exactly matching how the CURRENT flow already behaves: it checks daily and emails IF a group
exists, with no guarantee either way. This design keeps that same tolerance, just on a 3-day
cadence instead of a 1-day one).

```
Update item — target: the SAME approval library (Approval for Document /
Approval for Highly Confidential Document), item id = items('Apply_to_each')?['ID']

NextReminderAt = @{addDays(
  if(
    empty(items('Apply_to_each')?['NextReminderAt']),
    addDays(items('Apply_to_each')?['Created'], 3),
    items('Apply_to_each')?['NextReminderAt']
  ),
  3
)}
```

Walked through what this produces:
- File uploaded, `NextReminderAt` blank. First time it is 3+ days old and matched by `GetStale`,
  this expression reads `addDays(Created, 3)` (the due moment that just triggered it) and adds 3
  more days → stores `Created + 6 days`.
- Six days after upload, `NextReminderAt` (`Created + 6`) is now `<= now`, so `GetStale` matches
  again on the second branch. The expression now reads the STORED `NextReminderAt`
  (`Created + 6`) and adds 3 more → stores `Created + 9`.
- And so on, indefinitely, until the document is approved or rejected (at which point
  `OData__ModerationStatus eq 2` stops matching and the cycle ends on its own — no extra logic
  needed for that).

This is drift-free: because every step re-derives from the ANCHOR value already stored (never from
`utcNow()`), an hourly poll running a few minutes late never compounds — each cycle is always
exactly 3 days after the last one, at the same time of day the file was originally uploaded.

**Run this Update item action with `runAfter` set to `["Succeeded", "Failed", "Skipped"]`** on
`HasApprover` (Configure run after) — the clock must advance regardless of outcome, or a unit with a
broken/empty approver group would re-trigger the whole approver lookup on every hourly poll forever
instead of settling into its own 3-day rhythm.

## 6. Not built, and deliberately left as an open question

If the client instead wants reminders to **only ever go out inside a fixed night window** (say,
never before 20:00 local, regardless of what hour a file happened to be uploaded), that is an
ADDITIONAL condition on top of everything above — e.g. gating the actual send (not the due-time
math) on the local hour computed via `convertFromUtc(utcNow(),'Singapore Standard Time')`. Not
built here, because §2's simpler per-file-anchored design already satisfies what was actually
observed and asked for — ask before adding it, since it would mean a file uploaded at 9 AM gets its
first reminder held until that evening rather than 3 days later at 9 AM.

## 7. Build order

1. Create `NextReminderAt` on `Approval for Document`. Build and test everything above on
   `CRS — Reminding Approver to Approve` ALONE first — do not touch the HC twin yet.
2. Verify live: temporarily shorten the `GetStale` clauses' "3 days" to something like 1 hour (the
   same "prove the mechanism with a shortened window, then restore it" technique already used to
   verify this flow's original 3-day build), confirm ONE email arrives, confirm `NextReminderAt`
   gets stamped correctly on the item afterward, then restore the real 3-day values.
3. Only once that is confirmed working, create the SAME `NextReminderAt` column on
   `Approval for Highly Confidential Document` and clone all three changes onto
   `CRS — HC approval reminder`, keeping its existing `Role eq 'APRHC'` filter and `&lib=hc` link
   untouched.
