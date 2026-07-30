# Approval Destination Guard — Design

**Date:** 2026-07-29
**Status:** Agreed, implementing
**Touches:** `ApprovalDocument.tsx`

---

## Problem

The Auto-route flow's `Create new folder` action creates whatever path it is given. If the
**Documents unit folder does not exist**, the flow does not fail — it creates the entire
chain from the library root: `Segment / Department / Unit / Year / Document Type`.

Those flow-created folders **inherit the Documents library's permissions**, where
`DMS_SITE_MEMBERS` holds Read (site-entry spec §3). So every DMS user across every segment
can read that unit's approved documents.

It **fails open, silently**: the flow run shows success, nothing is logged as an error, and
the exposure is only visible by inspecting the folder's permissions. This is the opposite of
how reconciliation behaves — it breaks inheritance *before* granting anything, so a
misconfigured folder there is invisible rather than exposed.

Trigger in practice: a term is added, someone uploads and an approver approves, all before
the next reconciliation run creates and locks that unit's folders.

## Guard

Before approving, verify the destination is safe. Approval is **blocked** unless the
Documents unit folder **exists AND has unique permissions**.

### Deriving the unit folder

From the Staging file path, the unit folder is three levels up — the layout is always
`…/Unit/Year/Document Type/file.ext`, because the upload form ensure-creates Year and
Document Type on the way in. Walking up by position is depth-independent, so it holds for
segments whose Levels chain is deeper than Department → Unit.

```
/sites/Example/Staging/Segment/Dept/Unit/2026/Agreement/file.pdf
  drop filename        → …/Unit/2026/Agreement
  drop Document Type   → …/Unit/2026
  drop Year            → …/Unit                    ← the unit folder
```

Then swap the library segment, anchored on the web-relative prefix so a folder named
"Staging" deeper in the tree is not mangled — the same rule `toDocumentsPath` already uses
in `BulkUpload.tsx`:

```
/sites/Example/Staging/…  →  /sites/Example/Shared Documents/…
```

### Outcomes

| Check | Result |
|---|---|
| Folder exists, permissions unique | **Approve proceeds** |
| Folder confirmed missing (404) | **Blocked** — reconciliation has not created it |
| Folder exists but still inherits | **Blocked** — created by the flow or by hand; readable by all DMS users |
| Could not determine (403 / 429 / 5xx) | **Blocked** — never approve on an unverified destination |

Blocking on an uncertain answer is deliberate. A retry costs an approver seconds; a wrong
"proceed" publishes documents to everyone and nobody finds out. This is the opposite call to
the one the folder-map prune makes — there, uncertainty means *don't delete*; here it means
*don't publish*. Both resolve toward the outcome that cannot leak data.

### Message

> **This unit's folder is not ready in the Documents library, so the document was not
> approved.**
> Ask an administrator to run Folder Reconciliation, then approve again.

Names the fix in the words of the tool the admin has to open. No jargon about inheritance or
serial numbers — the approver cannot act on those.

## Rejection is not guarded

Only `Approved` is checked. Rejecting never copies anything to Documents, so there is no
destination to verify and no reason to block it.

## Known limitation

This guards the **Approval Document web part only**. Approving directly through the
SharePoint list UI bypasses it entirely, and the exposure still occurs. A complete fix would
add the same check inside the Auto-route flow — verify the unit folder before
`Create new folder`, and fail the run loudly instead of creating an unlocked chain.

Recommend doing both. The web part guard catches the normal path and gives the approver a
message they can act on; the flow guard is the backstop that no UI can route around.

## Out of scope

Creating or repairing the folder from the approval screen. The approver may not hold
permission to create folders in Documents, and a folder created there would be exactly the
unlocked one this guard exists to prevent. Reconciliation is the only thing that should
create and lock unit folders.
