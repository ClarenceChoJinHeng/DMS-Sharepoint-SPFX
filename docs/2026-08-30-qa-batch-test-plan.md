# CRS — Test plan for the 2026-08-30 client QA batch

**Build:** `1.0.332.0` — supersedes 322–331. ⚠ **None of those shipped correct JavaScript**; see §0.
**Site:** `/sites/ClarenceDMSTesting`.
**Scope:** the 35-item client QA deck, the share/revoke work from the same day, and the two runbooks
that came out of it. **Not** a full system test — `docs/2026-08-16-full-system-test-plan.md` still
covers that, and `docs/2026-08-27-recent-features-test-plan.md` covers the submission record.

> ⚠ **NOTHING FROM §2 ONWARDS HAS RUN ON A SITE.** 1,590 unit tests and a clean `tsc` cover the pure
> rules — `canActDirectly`, the blocked-character list, the preview URLs, the date format. They cover
> **none of the JSX**: this project has no UI tests, so every layout and copy change below is
> unverified until somebody looks at it.

---

## 0. Before anything — the deploy itself

⚠ On 2026-08-30 nine packages shipped **stale JavaScript**, because `heft package-solution` only zips
`release/assets` and does not run webpack. Every diagnosis that day was poisoned by it, so this
section is not a formality.

1. Upload `sharepoint/solution/sd-gatrie.sppkg`, **Replace** when prompted.
2. Site Contents → the app → **⋯ → Details**. It must NOT offer a new version.
3. **Close every CRS tab.** Not refresh — close. A tab open across a deploy keeps serving the old
   bundle by content hash.
4. Open a CRS page, F12 → Network, and confirm the `items?$select=…` request URL ends with **`&_=`**
   plus random characters. That arrived in 1.0.330.0. **If it is absent the build is not live and
   nothing below is worth testing.**

---

## 1. Carried over — the share and revoke chain

Verified against the list on 2026-08-30, but the phantom-row bug was in play throughout.

| # | Do | Expect |
|---|---|---|
| 1.1 | Open **Requests → Share** | Both `sd - asd - dasd` rows under **Decided**, marked `Revoked`. **`Waiting for you (0)`** |
| 1.2 | Raise a share as a PIC, approve as HoU, open **Shared files** | Recipient reads **`has access`** in green |
| 1.3 | **Revoke all access** | Row → `Revoked`; `DecisionNote` gains a line dated **today** |

⚠ **1.1 is the regression check for the phantom queue.** If decided rows show as `Pending` with live
Approve buttons again, the status parsing is broken and everything else can wait.

---

## 2. Upload form

| # | Do | Expect |
|---|---|---|
| 2.1 | Open the form | Intro reads *"Add folder information… All fields marked \* are mandatory."* — no *"Upload one or more documents in batches"* |
| 2.2 | Save a batch | The line above the cards reads *"All fields marked \* are mandatory"* — **not** *"1 batch ready · 1 document — nothing has been uploaded yet"* |
| 2.3 | Type `Tax: Q1/2026 #4` into **Project Name** | Renders `Tax Q12026 4`; hint under the box turns red and names `: / #` |
| 2.4 | Type `Risk ＆ Compliance` (fullwidth) into **Vendor** | **Unchanged.** ⚠ This is the one that would refuse the client's own data |
| 2.5 | Type `O'Brien` and `Q1-2026` | Unchanged |
| 2.6 | Repeat in **Document Name** and **Remark** | Same behaviour |
| 2.7 | Look at **Confidential Level** and **Legally Privileged** | The ⓘ icons sit closer to their labels |
| 2.8 | Add 21 files to one batch | Refused at 21 with a toast naming the cap. ⚠ The cap is no longer STATED beforehand |

---

## 3. Bulk Upload

| # | Do | Expect |
|---|---|---|
| 3.1 | Open it | Intro is only *"All fields marked \* are required."* — no file count, no *"same folder, same details"* |
| 3.2 | Type blocked characters into **Remark** | Stripped, red hint names them |
| 3.3 | Upload a file whose name clashes on the approved side | Rename still offered |

---

## 4. My Submissions

| # | Do | Expect |
|---|---|---|
| 4.1 | Open the list | **Uploaded** reads `27 Aug 2026` — spaces, no slashes, no leading zero |
| 4.2 | Open a submission | Batch rows show `2 files` with **no `BAT-20260827-…`** |
| 4.3 | Read the legend | Four **bullets**, not a paragraph |
| 4.4 | Open a file | **Document date** reads `17 Aug 2026` — no `12:00 AM`, no `8/17/2026` |
| 4.5 | Open a **.docx** → **Open in a new tab** | ⚠ **Opens in Word Online, does NOT download.** A 404 here means the WopiFrame URL is wrong — a *different* failure from the one it replaced |
| 4.6 | Same on a **.pdf** | Opens in the browser |
| 4.7 | As a **PIC**, open an approved file | **Request deletion** and **Request share** both offered |
| 4.8 | As the **Head of Unit for that unit**, open one of their own approved files | ⚠ **No request buttons**; the line reads *"You can delete or share this document yourself, in the library"* |
| 4.9 | As that same HoU, open a file they uploaded into a unit where they are only a PIC | Buttons **present** |
| 4.10 | Raise a deletion, reopen the file | Chip beside the status reads **`Deletion pending`**; quiet line says *"Deletion request submitted on …"* |
| 4.11 | Open a file whose share was revoked | *"This share was approved, and the access has since been taken back"* — ⚠ **never** *"attempted and failed"* |
| 4.12 | Open the delete dialog | Title *"Delete this file?"*, label **Reason**, button **Submit**, text says **93 days** |

⚠ **4.8 and 4.9 are a pair.** Either alone proves nothing — 4.8 without 4.9 could just mean the
buttons are broken for everybody.

---

## 5. Approval page and the bulk panel

| # | Do | Expect |
|---|---|---|
| 5.1 | Open a pending document | Radios read **Approve** / **Reject**; button reads **Proceed** |
| 5.2 | **Open in a new tab** on a `.docx` | Opens, does not download |
| 5.3 | Select 3 files → bulk approve panel | No *"3 selected."*; routing line reads *"Approved documents are automatically moved to … after approval."* |
| 5.4 | Approve one | Routes as before. ⚠ The three guards still run even though the sentence is shorter |

---

## 6. Requests

| # | Do | Expect |
|---|---|---|
| 6.1 | Open the page | Tabs read **Delete File**, Share, Shared files — **no "Your requests"** |
| 6.2 | Read the intro | *"Uploaders can submit deletion or sharing requests…"* |
| 6.3 | Look at a waiting request | A **Reason** label above the requester's text |
| 6.4 | Press **Reject** | File name on its own line, *"Please provide your reasoning below."*, label **Reason**, button **Reject** |
| 6.5 | Confirm the rejection | Banner reads **"Request has been rejected."** |
| 6.6 | Press **Approve** on a deletion | *"This file will be moved to the recycle bin, where it can be restored for 93 days."*, button **Approve** |
| 6.7 | Force a failure (decide something already decided elsewhere) | ⚠ Banner is **amber**, not green, and **scrolls itself into view** |

---

## 7. CRS Settings and Folder Management

| # | Do | Expect |
|---|---|---|
| 7.1 | Open CRS Settings | Cards read **File Type Management** and **Audit Log**, with the new subcopy |
| 7.2 | Click both | ⚠ They still **open the right pages** — the card titles changed, the page names did not |
| 7.3 | Open Folder Management | Intro *"Select an action below…"*; destructive heading reads **Take note** |
| 7.4 | Read the cards | *"Add another department or unit under an existing segment"*, the long reconciliation text, the new retire text |
| 7.5 | Start **Add a new segment** | Header reads **`Step 1 of 5`** — no *"5 of 5 steps still to check"*, no per-step Done / To do / Not checked |
| 7.6 | Click step 4 in the rail | ⚠ **Refused**, tooltip *"Finish the step you are on first."* |
| 7.7 | Next to step 3, then click step 1 | ⚠ **Allowed** |
| 7.8 | From step 1, click step 3 again | ⚠ **Allowed** — this is the fix; without it, going back stranded you |
| 7.9 | On step 2 with no segment typed | Next disabled, message ends *"press Refresh list if it is not there yet"* — ⚠ **not** *"jump straight on from the list of steps"* |
| 7.10 | Open **File Type Management** | Retitled; no *"File type settings"* heading; no *"Adding needs site owner rights"* |
| 7.11 | Add and remove a file type | Still works |

---

## 8. Group Management and Audit Log

| # | Do | Expect |
|---|---|---|
| 8.1 | Open Group Management | Member counts show real numbers, not `not known` |
| 8.2 | **Export** | People and Members populated. ⚠ If still `not known`, open the console — `fetchAllGroupMembers` now logs the HTTP status and the cause |
| 8.3 | Open the Audit Log | **No source line** under the actor (`Flow:DocumentsDeletions` etc.) |
| 8.4 | Export the audit CSV | `Source` still present — it left the view, not the data |

---

## 9. Regressions — what must NOT have broken

No client-facing change in any of these, which is exactly why a copy pass can damage them unnoticed.

| # | Check |
|---|---|
| 9.1 | An **HC upload** still routes to `HC Approval Document` and tags correctly |
| 9.2 | **Auto-route** still moves an approved document and preserves `Created By` |
| 9.3 | A **PIC** can still reach the upload form and file a document |
| 9.4 | **Reconciliation** completes with no new warnings |
| 9.5 | The upload **clash dialog** still offers a ` - Copy` name |
| 9.6 | **Bulk provisioning** still reports `already there` on a settled segment |

---

## 10. Not testable yet

- **`CRS — Expire shares`** — designed, not built
  (`docs/superpowers/specs/2026-08-30-share-expiry-enforcement-flow-runbook.md`). Until it exists the
  expiry date on a share is **displayed and not enforced**.
- **QA#5, the self-approval email** — designed, not built
  (`docs/superpowers/specs/2026-08-30-suppress-self-approval-email-runbook.md`). ⚠ It reverses the
  2026-08-24 decision; confirm with the client first.
- **`Audit — Documents deletions`** still says *"Deleted from Documents."* — should name the live
  title, `Restricted & Confidential Document`. One field, in Power Automate.
- **`Request.aspx` is still a draft.** Publish it: a modern page in draft denies read-only users with
  an error that looks exactly like a permissions bug.
- **`tenantDomains` config row** — absent, so *"outside the organisation"* is judged by whoever is
  looking. Set it to the real domain on SDG.

---

## 11. How to report a failure

Give the **status code or the exact message**, never the symptom alone. Three separate diagnoses on
2026-08-30 ran the wrong way for hours because a symptom had two opposite causes.

- **A control that does nothing** → hard refresh first. Then: does *clearing the cache* fix it (stale
  response) or only a *reload* (stale bundle)? Those need opposite fixes.
- **A read returning nothing** → the Network **status column**, before any theory.
- **A page showing something untrue** → the **Response** tab, not the screen.
