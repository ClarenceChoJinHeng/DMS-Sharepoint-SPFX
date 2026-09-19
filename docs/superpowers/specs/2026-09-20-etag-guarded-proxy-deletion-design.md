# ETag-guarded proxy deletion — design

**Client's decision, 2026-09-20: build the structurally correct fix (option 3 of three offered),
not the cheaper "refuse and retry" option.** Confirmed after the race was found to have genuinely
destroyed a document, not merely a theoretical risk.

## The problem, confirmed with evidence

`CRS — Execute approved deletion` resolves its target purely by `ItemUniqueId`, captured in
`triggerOutputs()` at the moment its trigger fires. It is a **polling** trigger — its actual
`recycle()` call can run tens of seconds after the deletion request was approved, not instantly.

Meanwhile, `Form.tsx`'s `deleteClashingDraftByProxy` polls for up to ~30 seconds to confirm the
proxy flow has cleared the old draft. If it gives up, it falls back to `Files/Add(overwrite=true)`
on the SAME name — which SharePoint applies **in place**, keeping the same `UniqueId`.

So: the client can give up and reuse an identity that the (slower, but genuinely still running)
proxy flow is still going to act on. When it eventually fires, it recycles whatever is CURRENTLY
under that GUID — which, if the fallback already landed, is the brand-new content, not the stale
draft the deletion was meant to clear.

**This is not theoretical.** Traced live on 2026-09-20 (see `docs/2026-09-20-session-handoff.md`,
items 23-28): three separate uploads of `test1-test1-test1-20092026.xlsx`
(`17:18:11Z`, `17:46:53Z`, `18:54:27Z`) all shared the identical `ItemUniqueId`, confirming the
fallback fired at least twice. The audit log shows the proxy flow's actual recycle landing 31
seconds after the deletion was raised — past the 30-second budget — and the file's content is now
recoverable only from the recycle bin, not live.

## The fix

Stamp the target's current `ETag` onto the `CRS Requests` row **at the moment the deletion is
authorised** (self-approved at write time, or approved by a human later). Have the flow re-check
that ETag against the file's CURRENT ETag immediately before recycling. If they differ — meaning
the identity has been reused for different content since authorisation — **skip the delete** rather
than blindly acting on whatever is there now.

ETag, not `Modified`/a timestamp: the destructive window observed above was under two seconds
(`18:54:26Z` to `18:54:27Z`), and SharePoint's `Modified` field is second-granularity — two writes
inside one second would read identically. ETag increments on every write regardless of timing.

### New column: `CRS Requests.TargetETag` (Text)

Written, never read/displayed by the app — it exists purely for the flow to compare against. Added
to `Requests.tsx`'s `COLUMNS` array so `ensureColumns`/`addMissingColumns` provision it like every
other optional column in this list (`Stage`, `RevokedBy`, `SubmissionFileId`).

### Code side — capture the ETag at the moment of authorisation, in all three writers

A shared helper, `readFileETag` in `src/shared/deletionGuard.ts` — one small function, reused by
all three call sites rather than three copies that drift:

1. **`Requests.tsx`'s `decide()`** — after `resolveDeletionTarget` resolves the CURRENT
   `ItemUniqueId` (the moment a human approves a Deletion request), read that identity's ETag and
   include it in the same MERGE that writes `Status: Approved`.
2. **`MySubmissions.tsx`'s `writeApprovedDeletionRequest`** — the self-approved path (a PIC's own
   pending/rejected draft, or an approver/admin's direct delete). Read the target's ETag right
   before writing the row, which is self-approved at creation.
3. **`Form.tsx`'s `deleteClashingDraftByProxy`** — the replace-clash path, which is exactly where
   the confirmed incident happened. Read the OLD (about-to-be-displaced) file's ETag alongside its
   already-captured `UniqueId`, and include it on the same self-approved request row.

**Fails open, deliberately.** A failed ETag read must never block writing the deletion request
itself — `readFileETag` returns `undefined` on any failure, and every write site omits the field
rather than failing when that happens. `TargetETag` absent is exactly the state every row written
BEFORE this change is already in, and the flow must treat that identically: no baseline captured
means no guard to apply, proceed as before. Understating the guard costs nothing new; refusing to
write a deletion request over a failed ETag read would be a regression.

**Optional-column read ladder**, matching this list's existing pattern (`Stage`/`RevokedBy`/
`SubmissionFileId`): on a 400, retry the write with `TargetETag` dropped first — it is the newest
addition, so it drops before the older optional columns on a site that has not yet reconciled/added
it.

### Flow side — `CRS — Execute approved deletion` (Power Automate, not code)

Cannot be built here; runbook for the exact steps is in
`docs/superpowers/specs/2026-09-20-etag-guard-execute-approved-deletion-runbook.md`. Summary: before
the existing recycle action, read the file's CURRENT `ETag`, compare to `TargetETag` from the
trigger. Match (or `TargetETag` blank, for backward compatibility with rows written before this
column existed) → proceed exactly as today. Mismatch → skip the recycle, mark the row `Failed` with
a note explaining why, naming that the request can be raised again if the document should still go.

## What this does NOT fix

- **Does not close the window entirely for a row written before this ships** — those carry no
  `TargetETag`, so they get no guard (fail open, matching today's behaviour). Only new requests
  raised after both halves (code + flow) are live are protected.
- **Does not stop the underlying cause** — the client-side poll can still time out and fall back to
  an in-place overwrite. This fix makes that safe rather than eliminating the race itself. If the
  fallback keeps firing often enough to be a nuisance on its own (separate from the safety question),
  that is a reason to revisit the poll budget or one of the other two options later — not something
  this change addresses.
- **Does not repair the already-recycled `test1-test1-test1-20092026.xlsx`** — that is a manual
  restore-from-recycle-bin action if the content is still wanted, unrelated to this fix.
