# Service account migration — crs@sdguthrie.com to gdc@sdguthrie.com

**Date:** 2026-09-19
**Not code — Power Automate only.** This repository cannot perform any of the steps below; it can
only record them precisely so nothing is missed across 23 flows, and so this isn't re-derived from
memory next time. Every flow was built by hand in the designer originally, and this migration should
be done the same way — retype the corrected values, never paste, and verify each flow with a real
test run before moving to the next.

## ⏭ STATUS AS OF 2026-09-19, LATE NIGHT — IMPORTED, FIXED, TURNED ON. NOT YET TEST-VERIFIED.

- **✅ All 23 flows imported** into the `gdc@sdguthrie.com`-owned Power Automate environment
  (`SD Guthrie`), each as a brand-new flow (Power Automate's import always showed "Create as new,"
  never "Update existing" — the old `crs@sdguthrie.com`-owned originals are untouched, off, and
  invisible from this account's "My flows" view since ownership scopes that list per-account).
- **✅ Two accidental duplicate imports found and cleaned up** (`NotifyApprovers` and
  `HCNotifyApprovers` each briefly existed twice) — confirmed resolved, exactly one copy of each of
  the 23 names remains.
- **✅ Two connections created**, both confirmed (via each connection's own Details page, "Owner:
  Guthrie Document Centre") to have authorized correctly as `gdc@sdguthrie.com`:
  `GDC Proxy - SharePoint` and `GDC Proxy - Outlook`.
- **✅ All 8 literal-value fixes applied and independently verified** by extracting the re-exported
  `.zip` for each of the 8 and grepping `definition.json` directly (not just eyeballing the
  designer) — confirmed **zero** remaining occurrences of `crs@sdguthrie.com` or
  `Guthrie Central Repository System` across all 8, and the new values present exactly where
  expected. One real miss was caught and fixed this way: `HCNotifyApprovers`'s `Compose` action
  (recipient-list join) still had the old literal after the first attempt only fixed the
  `HasRecipients` condition beside it — both are now consistent.
  - **Also noted, not a defect:** `NotifyApprovers`'s email sign-off was changed from `"Thank you,
    SD Guthrie"` to `"Thank you, Guthrie Document Centre"` — a deliberate branding update, confirmed
    against the original export to rule out anything unexpected. `HCNotifyApprovers` almost
    certainly has the same old sign-off in its own email body and was NOT updated to match — cosmetic
    only, your call whether to bother.
- **✅ All 23 flows turned On.**
- **⚠⚠ NOT YET DONE: the actual verification/test-run step, per-flow, for any of the 23.** Turning a
  flow on proves nothing on its own — the "run one real test through this specific flow's own
  trigger, and read the run's own output" step in "Per-flow steps" below has not been done for any
  of them yet. This is the next and last remaining step in the whole migration. Recommended order to
  test in is unchanged from "Recommended order" below — low-stakes audit/reindex flows first, then
  folder-approval → Auto-route → deletion execution → bulk auto-approve, then the notification
  flows, then the archive movers last.
- **Not yet done, lower priority, from "What NOT to do":** removing `crs@sdguthrie.com` from
  `CRS Owners` once everything above is confirmed working — leave that until the test-run pass is
  complete.

## Why this is riskier than it looks

Every one of these flows already has a *connection* — a reference to whichever SharePoint/Outlook
account was signed in when that connector was first authorised. Swapping who a flow runs as is
mostly a matter of reconnecting that reference. But **8 of the 23 flows also have the old account's
email address, or its display name, typed as a literal string inside the flow's own logic** — not
in the connection, in an actual condition or a value being written. Those will not update just
because the connection changes, and if they're missed the flow will keep running (green, no error)
while silently doing the wrong thing — mainly, treating the new account's own actions as if they
were a stranger's, and double-logging them. This is exactly the class of bug this project has been
bitten by before with these flows (see CLAUDE.md's "DELETE IS PERFORMED BY PROXY" section, and the
repeated warnings throughout the audit-flow history about not trusting a green run as proof a
comparison value is right).

## The two values

| | Old | New |
|---|---|---|
| Email | `crs@sdguthrie.com` | `gdc@sdguthrie.com` |
| Display name | `Guthrie Central Repository System` | `Guthrie Document Centre` |

**✅ Confirmed already done:** `gdc@sdguthrie.com` is already a member of the site's System
Administrators (`CRS Owners`) — visible directly in the Group Management screenshot used to find the
display name above. So the one prerequisite this migration would otherwise need (the new account
holding the same elevated SharePoint rights the old one has, for delete/write actions) is already
satisfied. Nothing to do there.

## The complete inventory — 23 flows, from the 2026-09-18 export

Every flow needs its **SharePoint connection** reconnected. **7 of them also use a second
connector — Office 365 Outlook** — for sending email, and that needs reconnecting separately, since
"Send an email (V2)" sends *from* whichever mailbox owns that connection, independent of the
SharePoint connection sitting beside it in the same flow.

**8 of the 23 also need a literal value retyped inside the flow itself** — these are listed with
their exact before/after in the next section. Everything not in that list needs the connection swap
only, nothing else.

| Flow | SharePoint | + Outlook | Literal fix needed |
|---|---|---|---|
| Audit — approval activity | ✓ | | no |
| Audit — HC approval activity | ✓ | | no |
| Audit — approval deletions | ✓ | | **yes** |
| Audit — HC approval deletions | ✓ | | **yes** |
| Audit — Documents deletions | ✓ | | **yes (two values)** |
| Audit — HC Documents deletions | ✓ | | **yes (two values)** |
| CRS — Audit replacements | ✓ | | no |
| CRS — Notify request activity | ✓ | ✓ | no |
| Auto-route (approved Pending files to Documents Library) | ✓ | ✓ | no |
| HC Auto Route | ✓ | ✓ | no |
| CRS — Approve new folders in Approval Document | ✓ | | no |
| HC folder approval | ✓ | | no |
| CRS — Auto-approve bulk imports in Approval Document | ✓ | | no |
| HC auto-approve | ✓ | | no |
| CRS — Execute approved deletion | ✓ | | no |
| CRS — Archive after seven years | ✓ | | **yes** |
| CRS — HC Archive after seven years | ✓ | | **yes** |
| CRS — Nightly library reindex | ✓ | | no |
| NotifyApprovers | ✓ | ✓ | **yes** |
| HCNotifyApprovers | ✓ | ✓ | **yes** |
| CRS — Reminding Approver to Approve | ✓ | ✓ | no |
| CRS — HC approval reminder | ✓ | ✓ | no |

That's 22 rows — the 23rd of the exported files, `CRS — Execute approved deletion`, is listed above;
the count in this heading (23) is the number of `.zip` files exported. If a recount ever comes up
short, the exported zip names are the ground truth, not this table.

## The 8 flows needing a literal value fixed — exact before/after

**⚠ Fix these by opening the flow in the designer, finding the exact action/condition named below,
and retyping the value. Do not paste — this project has hit the "invisible trailing whitespace
pasted into a Power Automate expression box" bug more than once, and it produces exactly the kind of
silent failure this migration is trying to avoid.**

### 1. Audit — approval deletions
Action: the `If` condition guarding whether to write the deletion audit row (its `expression`
combines a `Was_routed` count check with this one).
- Find: `equals(...DeletedByEmail..., "crs@sdguthrie.com")` (inside a `not(...)`)
- Retype the literal to: `gdc@sdguthrie.com`

### 2. Audit — HC approval deletions
Same shape as #1, same fix, same value.

### 3. Audit — Documents deletions
Same `If` condition, but this one has **two** literals to fix, both inside the same `and(...)`:
- `equals(triggerOutputs()?['body/DeletedByUserName'], "Guthrie Central Repository System")` —
  retype the literal to `Guthrie Document Centre`
- `equals(...DeletedByEmail..., "crs@sdguthrie.com")` — retype the literal to `gdc@sdguthrie.com`

### 4. Audit — HC Documents deletions
Same shape as #3, same two fixes, same two values.

### 5. CRS — Archive after seven years
Action: the `Create item` step that writes the `Archived` audit row (`item/ActorEmail`).
- Find: `item/ActorEmail` set to the literal `crs@sdguthrie.com`
- Retype to: `gdc@sdguthrie.com`
- (`item/ActorName`, currently the literal `Archive mover (automated)`, is not identity-specific —
  leave it as-is unless you want to reword it for its own reasons.)

### 6. CRS — HC Archive after seven years
Same shape as #5, same fix, same value.

### 7. NotifyApprovers
**Two separate places, same literal, same flow:**
- The `Compose` action building the recipient list —
  `union(body('PickEmails'), createArray('crs@sdguthrie.com'))` — retype the literal inside
  `createArray(...)` to `gdc@sdguthrie.com`.
- The `HasRecipients` condition — `greater(length(union(body('PickEmails'),
  createArray('crs@sdguthrie.com'))), 0)` — same literal, same fix. **Both must be changed
  together**; they're two independent copies of the same expression fragment, not one shared value,
  so fixing only one leaves the other still comparing against the old address.

### 8. HCNotifyApprovers
Same shape as #7, same two spots, same fix.

⚠ **Worth knowing before you fix #7/#8, not fixing to do differently right now:** that
`createArray('crs@sdguthrie.com')` unconditionally adds the proxy account as a recipient on every
approval-needed email, in BOTH the send list and the "are there any recipients" guard — which means
the guard can never actually be empty, since the proxy address is always in the array. That's a
pre-existing property of these two flows, not something this migration introduces or needs to solve;
flagging it here only so the "the `gdc@sdguthrie.com`-shaped condition passing" doesn't get read as
evidence the guard is doing its real job.

## Recommended order

Not required, but this project's own established habit is "one fix, one deploy, one check" rather
than batch-changing everything and hoping — and with 23 flows, doing them in a sensible order means
the process gets validated on something low-stakes before it's trusted on something that actually
deletes or moves documents.

**1. Low-stakes, no literal fix — validate the reconnect process itself:**
`Audit — approval activity`, `Audit — HC approval activity`, `CRS — Audit replacements`,
`CRS — Nightly library reindex`. Nothing here writes anything a person will notice if it's briefly
wrong; good place to confirm the reconnect steps work as expected before anything that matters rides
on them.

**2. Low-stakes, WITH a literal fix — validate that half of the process too:**
`Audit — approval deletions`, `Audit — HC approval deletions`, `Audit — Documents deletions`,
`Audit — HC Documents deletions`. Same reasoning as above, but now also exercising the literal-value
fix on something where a wrong answer is easy to detect (a duplicated audit row) and cheap to fix
(delete the stray row, redo the fix).

**3. Core document flow — the highest-stakes tier, do these carefully, one at a time, with a real
test upload/approval/deletion after each:**
`CRS — Approve new folders in Approval Document`, `HC folder approval` (needed for uploads to work
at all — a folder stuck Pending blocks the uploader inside it), then `Auto-route`, `HC Auto Route`
(these physically move documents), then `CRS — Execute approved deletion` (this physically deletes
documents — recoverable via the recycle bin, but confirm with a throwaway test file, not a real
one), then `CRS — Auto-approve bulk imports in Approval Document`, `HC auto-approve`.

**⚠ FOR THIS TIER, PAUSE UPLOADS SITE-WIDE WHILE EACH ONE IS OFF.** A trigger firing while a flow is
disabled is not queued and replayed once it's re-enabled — it's simply missed, and for this tier
that means a document can get stuck (approved but never routed, a folder left Pending with nobody
able to reach it, a deletion approved but never actually executed). Set `uploadsPaused = yes` on
`CRS Config` (the same site-wide toggle the folder-structure migration flow already uses for
exactly this reason) before starting this tier, work through all six flows in this tier while it's
on, then set it back to `no` once every one of them is confirmed working. Do this outside normal
working hours if at all possible — the same operational answer this project has already settled on
for every other change that needs uploads paused.

**4. Notifications, WITH literal fixes:** `NotifyApprovers`, `HCNotifyApprovers`.

**5. Notifications, no literal fix, lower urgency (these fire every 3 days, not on every
action):** `CRS — Reminding Approver to Approve`, `CRS — HC approval reminder`,
`CRS — Notify request activity`.

**6. Lowest urgency — scheduled, and nothing on either site is close to firing them for real yet
(per CLAUDE.md, the oldest document is nowhere near seven years old):** `CRS — Archive after seven
years`, `CRS — HC Archive after seven years`. Safe to leave for last.

## Per-flow steps

**One flow at a time — turn it off, reconnect it, verify it, turn it back on, then move to the
next.** Never leave a flow off for longer than it takes to do its own steps, and never turn off
more than one flow at once (tier 3's site-wide upload pause above is the one deliberate exception,
and it exists precisely because that tier's flows need to be off for slightly longer while each one
is worked through).

For every flow:
1. **Turn the flow Off first.** If you are updating the SAME existing flow in place, this avoids a
   trigger firing mid-edit while the connection is half-swapped. If the way you are reconnecting it
   produces a SEPARATE new flow instead of updating in place, this is the step that prevents the old
   and new versions ever both being On at once — see "What NOT to do" below for why that matters.
2. Open it in Power Automate, sign in as (or already have an existing connection for)
   `gdc@sdguthrie.com`.
3. Reconnect the SharePoint connection — either via the flow's own Connections panel (if the maker
   portal offers a direct swap) or by re-importing the exported package into this same environment
   and choosing to update the existing flow, selecting/creating the `gdc@sdguthrie.com` connection
   for the SharePoint connection reference.
4. **If this flow also uses Office 365 Outlook** (see the table above): reconnect that connection
   too, the same way. Missing this one is easy to miss since the SharePoint side can look fully
   reconnected while the email side is still sending as the old account.
5. **If this flow is one of the 8 listed above**: make its specific literal-value fix(es) now, by
   hand, per the exact before/after given.
6. Save.
7. **Confirm the connection panel reads "Connected to gdc@sdguthrie.com"** on every action that
   shows a connection — this project's own established check for every flow it has ever built, and
   it catches a reconnect that silently didn't take.
8. Turn the flow back On.
9. Run one real test through this specific flow's own trigger, and read the run's own output —
   never assume from "no error shown" that the right thing happened. For the 8 flows with literal
   fixes, specifically confirm the NEW behaviour: does a `gdc@sdguthrie.com`-performed deletion now
   get correctly excluded (no duplicate audit row)? Does the archived-file row now carry
   `gdc@sdguthrie.com`? Does the notification list now carry the new address, not the old one?
10. **If this reconnect produced a separate new flow rather than updating the old one in place**,
    only now — once the new one is confirmed working — turn the OLD (`crs@sdguthrie.com`-connected)
    version off for good. Do not delete it yet; leave it disabled for a few days in case something
    the test run didn't happen to exercise turns up, then remove it once you're confident.

## What NOT to do

- **Don't batch-reconnect all 23 without testing between them.** If the process has a mistake in it
  (a wrong connector picked, a missed literal), doing all 23 that way means finding out from 23
  simultaneous small failures instead of one, with no way to tell which fix produced which problem.
- **Don't ever have the old and new version of the same flow both turned On at once.** If reconnecting
  produces a separate new flow rather than updating the existing one in place, this is a real,
  previously-hit failure mode in this exact project — CLAUDE.md records "a duplicate flow left
  active" as one of the concrete things that went wrong during the 2026-09-17 deletion-proxy work.
  Two live copies of the same flow both watching the same trigger can both fire on one event and race
  each other — for something like Auto-route, that means two flows both trying to copy/move/delete
  the same file at once. The "turn old off first, only turn new on once verified, only then retire
  old" order in the steps above exists specifically to make this impossible.
- **Don't turn off a tier-3 flow (folder approval, Auto-route, deletion execution, bulk auto-approve)
  without also pausing uploads first.** A trigger event that happens while the flow is off is missed,
  not queued — see the note under "3. Core document flow" above.
- **Don't leave both accounts holding elevated SharePoint rights indefinitely once this is
  done.** Once every flow is confirmed migrated and working, removing `crs@sdguthrie.com` from
  `CRS Owners` (visible on the same System Administrators screen the display name was found on) is
  the natural close-out step — not urgent, but don't forget it exists as a step, since the old
  account otherwise stays a live administrator with nothing depending on it.
- **Don't guess the display name if it's ever unclear for a different flow later.** It was findable
  directly from a live SharePoint screen this time; if another literal string search ever turns up
  something that looks identity-shaped but isn't confirmed, look it up the same way rather than
  assuming.
