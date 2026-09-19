# Adding the ETag guard to `CRS — Execute approved deletion` — flow runbook

**No file calls this document** — a Power Automate EDIT runbook, matching every other `*-runbook.md`
in this project. Read alongside `2026-09-20-etag-guarded-proxy-deletion-design.md`, which carries the
reasoning and the evidence this is based on; this carries the exact edit steps.

**Status: NOT DONE.** The code half (all three writers stamping `TargetETag` onto `CRS Requests`,
plus the new column itself) is built and verified — `tsc --noEmit` clean, full suite 1968/1968, zero
new lint warnings. **This flow edit is the only remaining piece.** Until it is applied, the new
column is written and simply ignored — no worse than before, but no safer either.

⚠ This EDITS the EXISTING, already-built `CRS — Execute approved deletion` flow (spec:
`2026-09-17-execute-approved-deletion-flow-runbook.md`). It does not replace it, and every existing
action, field and branch described in that runbook stays exactly as it is — this only inserts two
new actions and a new condition BEFORE the existing recycle step.

## 0. What you're inserting, in plain terms

Right now, the flow trusts the deletion request completely: it reads `ItemUniqueId` off the
triggering row and recycles whatever is under that GUID, whenever it happens to get around to it —
which can be tens of seconds after the request was approved, since this is a polling trigger.

The new step asks one question immediately before deleting anything: **"is this still the same file
I was told to delete, or has something else been uploaded under this identity since?"** — by
comparing the file's current `ETag` (a value that changes on every single edit, unlike a
timestamp) against the `TargetETag` the app captured at the moment the deletion was approved. If
they don't match, something reused this identity in between — most likely a replace-clash fallback
overwriting a name in place — and deleting it now would destroy the WRONG content. So it skips the
delete instead, and records why.

## 1. Open the flow, find where to insert

Open `CRS — Execute approved deletion` in edit mode. You'll see, inside the outer Condition's TRUE
branch (`RequestType eq 'Deletion' and Status eq 'Approved'`):

```
[Recycle_the_file]   ← "Send an HTTP request to SharePoint", recycle() by GUID
    ├─ (run after: is successful) → Create item (logs "Deleted" to CRS Audit Log)
    └─ (run after: has failed)    → Update item (writes Status: Failed)
```

You are inserting two new actions and a new condition **immediately before** `Recycle_the_file`,
inside the same TRUE branch. `Recycle_the_file` itself, and both of its existing branches, are
**not changed** — only where they sit changes (they move one level deeper, inside the new
condition's TRUE branch).

## 2. New action — read the file's current ETag

**Action: Send an HTTP request to SharePoint.** Add it directly above `Recycle_the_file`, inside the
same branch (nothing needs to run-after it yet — it's the first thing in this section).

```
Site Address: <this site>
Method: GET
Uri: _api/web/GetFileById(guid'@{triggerOutputs()?['body/ItemUniqueId']}')?$select=ETag,Exists
Headers:
  Accept: application/json;odata=nometadata
```

Name it **`GetCurrentETag`** (used by name below — rename the expressions if you call it something
else).

⚠ **This can genuinely fail (404) — that's expected, not a bug**, and it's exactly what happened to
`test1-test1-test1-20092026.xlsx`: the target had already been recycled by the time this ran. Do
not add retry logic here; let it fail, and handle that explicitly in step 5 below.

## 3. New condition — does the ETag still match?

**Action: Condition.** "Configure run after" on this action: **is successful** on `GetCurrentETag`
— so it only evaluates when the GET actually returned something to compare.

Two rows joined by **Or**:

- `outputs('GetCurrentETag')?['body/ETag']` **is equal to**
  `triggerOutputs()?['body/TargetETag']`
- `triggerOutputs()?['body/TargetETag']` **is equal to** `` (empty)

⚠ **The second row is what keeps every row written BEFORE this column existed working exactly as
today** — those rows have no `TargetETag` at all, so there is no baseline to check against, and
treating that as "proceed" (rather than "always skip") is the deliberate fail-open choice — see the
design doc's "fails open" section. Enter both rows through the ordinary field editor, not `fx` — they
are plain equality checks against trigger fields, no composed expressions needed.

**TRUE branch:** move the EXISTING `Recycle_the_file` action (and everything already hanging off
it — its success and failure branches) into here, unchanged. In the Power Automate designer this is
a drag: grab `Recycle_the_file` and drop it inside this new condition's Yes branch. Its own two
"Configure run after" branches travel with it automatically.

**FALSE branch:** this is new — see step 4.

## 4. New action — the ETag genuinely didn't match

**Action: Update item** (native SharePoint connector, on **`CRS Requests`**, item id
`triggerOutputs()?['body/ID']` — same target the existing §5 failure branch in the original runbook
already uses). Place it in the Condition's **No** (False) branch from step 3.

```
Status: Failed
DecisionNote:
  concat(
    triggerOutputs()?['body/DecisionNote'], ' ',
    'Automated deletion skipped — the file changed after this deletion was approved (it may have ',
    'been replaced by a newer upload), so nothing was deleted to avoid removing the wrong content. ',
    'If this document should still be deleted, raise the request again.'
  )
```

## 5. New action — the file was already gone before the check even ran

**Action: Update item**, same list and target as step 4. "Configure run after" on this action:
**has failed** on `GetCurrentETag` — this is a SIBLING of the Condition from step 3, not nested
inside it, since a failed GET never reaches that condition at all.

```
Status: Failed
DecisionNote:
  concat(
    triggerOutputs()?['body/DecisionNote'], ' ',
    'Automated deletion could not proceed — the document could not be found (it may already have ',
    'been removed).'
  )
```

⚠ This is a DIFFERENT outcome from step 4's, and the wording says so: step 4 means "something else
now occupies this identity"; step 5 means "there is nothing here at all". Worth keeping distinct in
case either recurs often enough to be worth investigating on its own.

## 6. What "it worked" looks like

- **Ordinary case, unaffected:** a deletion request approved, nothing else touches that file in the
  meantime, `GetCurrentETag` succeeds, the Condition's ETags match, `Recycle_the_file` runs exactly
  as it always has, the existing `Deleted` audit row is written. No visible change from today.
- **The race, now caught instead of destructive:** a replace-clash fallback overwrites the same
  name/identity after the deletion was approved but before the flow got to it. `GetCurrentETag`
  reads the NEW content's ETag, it does not match `TargetETag`, the Condition goes False, step 4's
  Update item marks the request `Failed` with an explanation — **and the new content is never
  touched.**
- **Rows written before this shipped:** `TargetETag` is blank, the Condition's second OR-row
  matches, behaviour is identical to before this change — no regression for old rows.

## 7. Test it

1. Confirm a genuinely ordinary deletion (no clash involved) still recycles correctly and logs
   `Deleted` as before.
2. Reproduce something close to the original incident if you can: replace-clash a file, and if
   possible slow down or delay the flow's own execution relative to the client's ~30-second poll
   (hard to force reliably — even without forcing it, watch the NEXT time a replace-clash fallback
   fires naturally and confirm the row reads `Failed` with the "skipped" wording rather than the
   file disappearing).
3. Confirm an ordinary Share request is completely unaffected — none of this touches the Share
   branch, which still executes directly in the approver's session as before.

## 8. What this does not fix

Same caveats as the design doc: this protects requests raised AFTER both halves are live; it does
not repair `test1-test1-test1-20092026.xlsx` (recycle-bin restore, if still wanted, is separate); and
it does not stop the client-side poll from timing out in the first place — it makes that timeout
safe rather than eliminating it.
