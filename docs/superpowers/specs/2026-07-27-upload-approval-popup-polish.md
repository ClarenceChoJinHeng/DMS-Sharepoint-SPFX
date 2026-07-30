# Upload & Approval Popup Polish — 2026-07-27

Client-requested UX changes across the upload form, bulk upload, and approval web parts.
Designer icons supplied as `Group 70858 (2).svg` (replace warning), `Group 70890.svg`
(approval success), `Group 70890 (1).svg` (rejected), `Group 70890 (2).svg` (pending) — all
184×184, embedded inline (matches the existing Form.tsx success-popup pattern).

## 1. Replace-file guard (Form.tsx + BulkUpload.tsx)
Today a name clash aborts with an error toast (single) / silent "skipped" (bulk).
Replace with a **"Replace Existing File"** Yes/No modal:
- Title: **Replace Existing File**
- Body: *"We noticed there's a same name file. Are you sure you want to override this file?"*
- **Yes** → re-run `Files/Add` with `overwrite=true`; **No** → cancel that file.
- Icon: orange warning circle (`Group 70858 (2).svg`).

Implementation: a promise-based prompt (`askReplace(name): Promise<boolean>`) backed by a
`replaceAsk` state + a `resolve` ref. Uploads are sequential, so one shared modal is safe.
Both `overwrite=false` literals become an `overwrite` flag flipped to `true` on Yes.
Bulk: a Yes-replaced file is reported as `uploaded`; No keeps the existing `skipped` outcome.

## 2. Rename success button (Form.tsx)
Single-upload success popup button `Track Status on Home` → **`Track File Status`**
(title stays "Upload Successful"; still redirects to the site home).

## 3. Bulk progress relocation + progress bar + X-to-delete (BulkUpload.tsx)
Incremental (keep the batches model — confirmed with client):
- Move the live per-file progress panel to render **above** the batch list / config form
  (currently at the bottom, after the Upload button).
- Add an **overall progress bar** above the panel: `completed / total` files, where completed =
  any terminal state (done/tagFailed/skipped/failed/deleted).
- **X on a completed file deletes it from Documents** (bulk upload writes straight
  to the Documents library, skipping Staging). Track each uploaded file's
  `ServerRelativeUrl` in the live state (`sru`). X shows only for `done`/`tagFailed`
  (files that actually exist). Click → confirm → `DELETE`
  `getfilebyserverrelativeurl(@f)` → mark the row `deleted`.

## 4–6. Approval decision popups (ApprovalDocument.tsx)
Replace the inline success screen with a modal card (matches Form.tsx popup style),
button **Back to Document** → `backUrl()` (the file's Staging folder = "redirect to staging"):
- **Approved** → "Approval Successful" / *Your document has been approved successfully.* / teal check (`Group 70890.svg`)
- **Rejected** → "Document Rejected" / *The document has been returned for revision.* / red doc-x (`Group 70890 (1).svg`)
- **Pending** → "Still pending Approval" / *You can monitor its status anytime.* / orange doc-clock (`Group 70890 (2).svg`)
