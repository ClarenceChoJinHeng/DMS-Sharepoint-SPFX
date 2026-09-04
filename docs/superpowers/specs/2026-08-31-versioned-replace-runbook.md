# Auto-route: replace by VERSIONING, not by delete-and-recreate

**Date:** 2026-08-31
**Status:** RUNBOOK — not built.
**Flows to change:** `Auto-route approved Pending files to Documents Library`, **and its clone**
`HC Auto Route`.

---

## 0. Why — the finding this exists to fix

Proven on ClarenceDMSTesting, 2026-08-31. One document replaced through the approval page:

| | before | after |
|---|---|---|
| version | `1.0` | **`1.0`** — a NEW one, not `2.0` |
| size | 227.3 KB | 117.4 KB |
| uploader | chocheetuck4 | Clarence Cho |

**`Copy file` with `nameConflictBehavior: 1` DELETES the destination item and creates a new one.**
Version history is never reached, so the 500-major-version setting on both approved-side libraries
makes no difference. The old document goes to the recycle bin for 93 days — recoverable, but only by
someone who thinks to look there, and gone after 93 days.

> ⚠ **THE CLIENT CHOSE REPLACE BELIEVING VERSION HISTORY KEPT THE OLD FILE**, because our own confirm
> dialog said so (corrected in 1.0.336.0). This runbook makes the behaviour match what they were told.

**The outcome wanted:** ONE document, same name, same link — with the previous content retrievable as
`1.0` behind it. **Not** the `nameConflictBehavior: 2` duplicate (`TEST1.pdf`), which the client has
explicitly rejected.

> The client accepted (2026-08-31) that anyone who can read the document can also open its version
> history. That is consistent with the model rather than an exception to it: the unit folder is already
> this system's smallest confidentiality boundary, and HC version history sits behind the HC ACLs the
> same way.

> ⚠ **`2026-08-28-file-replacement-design.md` §7.3 SAYS "version history holds the content" AND IS NOW
> WRONG.** It was written before this was tested. Correct it when this ships, or it becomes the next
> stale line somebody acts on.

---

## 1. The shape of the change

Today, inside the Condition's True branch:

```
Compose -> Compose 1 -> Create new folder -> Copy file -> GetSourceUniqueId -> Create item -> ...
```

After:

```
Compose -> Compose 1 -> Create new folder
  -> GetDestMatch          (NEW - does a file of this name already exist there?)
  -> DestExists            (NEW - Condition)
       True:  GetSourceContent -> UpdateFile     (writes 2.0)
       False: Copy file                           (unchanged)
  -> DestItemId            (NEW - Compose, normalises the item id)
  -> GetSourceUniqueId -> Create item -> ...      (unchanged, but reading DestItemId)
```

---

## 2. `GetDestMatch` — must not FAIL when the file is absent

**Use `Get files (properties only)`, NOT `Get file metadata using path`.**

`Get file metadata using path` **fails with a 404** when the file is not there, and "absent" is the
normal case — so the flow would red-run on every ordinary approval. Working around that needs
`Configure run after` plus a status test, which is more moving parts than the problem deserves.

`Get files (properties only)` answers an empty `value` array instead. No failure, no run-after
juggling.

| Field | Value |
|---|---|
| Site Address | the site |
| Library Name | the **approved-side** library (`Restricted & Confidential Document` / `Highly Confidential Document`) |
| Filter Query | `FileLeafRef eq '<the file name>'` |
| Limit Entries to Folder | the destination folder — the same value `Copy file`'s **Destination Folder** already computes |

> ⚠ **Reuse the existing `Compose 1` output for the folder** rather than rebuilding the path. There is
> already one expression in this flow that knows where the document goes; a second one is a second
> thing to get wrong, and it would be wrong only for the segments nobody tests.

> ⚠ **The name needs escaping if it can contain an apostrophe.** OData escapes a single quote by
> doubling it, so wrap the name in a `replace()` that turns one quote into two. A document called
> `O'Brien Agreement.pdf` otherwise breaks the filter — and a broken filter returns nothing, which
> reads as "no clash" and silently takes the `Copy file` branch, i.e. today's behaviour. **Failing back
> into the old bug is the worst kind of failure, because nothing reports it.**

---

## 3. `DestExists` — the Condition

```
length(coalesce(body('GetDestMatch')?['value'], createArray()))   is greater than   0
```

⚠ Enter the `0` through the **fx** editor. A `0` typed into the value box is TEXT, `length()` returns
an Integer, and the comparison fails the whole condition with a type error. This has cost time on
three separate flows in this project.

⚠ No trailing newline in the expression. Pasting a multi-line expression into a Power Automate field
brings one with it, and it has broken four fields in this project — twice invisibly.

### True branch — `GetSourceContent` then `UpdateFile`

| Action | Field | Value |
|---|---|---|
| `Get file content` | File Identifier | the SOURCE file — the same identifier `Copy file` uses today |
| `Update file` | File Identifier | the matched file's identifier, from `GetDestMatch`'s first row |
| | File Content | `GetSourceContent`'s body |

`Update file` writes into the existing item, so SharePoint records a **new major version**. One
document, same name, same URL, same `UniqueId`.

### False branch — `Copy file`

Unchanged, including `nameConflictBehavior`. ⚠ **Leave it at `1`.** It is unreachable on this branch
(we only get here when no file of that name exists) — but if the check ever answers wrongly, `1`
overwrites where `2` would silently create the duplicate the client rejected. Fail towards the
behaviour they asked for.

---

## 4. `DestItemId` — the part that is easy to get wrong

**Everything downstream reads `Copy file`'s output** — `GetSourceUniqueId`, the audit `Create item`,
the `validateUpdateListItem` metadata stamp, the source delete, the email condition. On the Update
branch that action never ran, so every one of those references resolves to null.

Add ONE `Compose` immediately after the Condition, whose expression answers *the destination item's
id from whichever branch ran*: the matched row's `ID` when `GetDestMatch` found something, and
`Copy file`'s id otherwise.

Then **repoint every downstream reference to that `Compose`.**

> ⚠ **DO NOT duplicate the tail into both branches.** It is nine actions; two copies drift, and the
> copy that drifts is the rarely-run one — which here is the replacement path, i.e. the one this whole
> change exists for. One `Compose`, one tail.

> ⚠ **Search the flow's Code view for `Copy_file` when you think you are done.** Anything left pointing
> at it is a null on the Update path, and nulls in this flow fail quietly: a missing metadata stamp
> looks like an untagged document, not like a broken flow.

---

## 5. `Created By` — a CLIENT decision, not ours

On the Update path the destination item **already exists**, with the ORIGINAL uploader in `Author`.
Today's stamp sets `Author` to whoever uploaded the replacement. Both answers have a real cost:

| | keeps | loses |
|---|---|---|
| **Overwrite `Author`** with the replacer | the replacement shows in THEIR My Submissions | the original author of the document, permanently |
| **Leave `Author`** as the original | provenance — who first filed this record | ⚠ **the replacer's document does not appear in their My Submissions**, because that page filters on the signed-in user being the author |

Neither is obviously right. Version history records who wrote `2.0` either way, so nothing is lost to
the audit — this is about which name the *document* carries, and whose My Submissions it appears in.

**Put it to the client in one sentence.** Do not decide it inside the flow.

---

## 6. The HC clone

`HC Auto Route` needs the identical change. Every HC clone in this project has shipped with a library
reference nobody swapped — six of them in `HC Auto Route` itself, and one in `HC folder approval` that
failed silently for six days.

**One field in the new actions carries a library name:** `GetDestMatch`'s **Library Name**.
`Update file` and `Get file content` address files by identifier. Swap the one, check the one.

---

## 7. Testing — and the empty case proves nothing

1. **Ordinary approval, no clash.** Must behave exactly as today: `DestExists` False, `Copy file` runs,
   document routes with metadata and `Created By` intact. ⚠ **Run this FIRST** — it is the path every
   approval takes, and breaking it to fix a rare one is the bad trade.
2. **Replacement.** Upload a document composing to an existing filed name, choose *Send for approval as
   a replacement*, approve. Then open **Version history** on the filed document:
   **expect `2.0` current and `1.0` openable with the old content.**
3. **Check the recycle bin is EMPTY of that document.** Nothing should have been deleted. If a copy is
   sitting there, `Copy file` ran and the Condition answered wrongly.
4. **Check the metadata stamped correctly** on the replaced document — that is what proves the
   `DestItemId` Compose was wired through, and a missed reference shows up here as blank columns rather
   than as an error.
5. Repeat 1 and 2 on the HC vertical.

> ⚠ **A green run means nothing on its own.** Both faults this change fixes shipped with green runs:
> the false dialog text, and the delete-and-recreate itself. Open the document.

---

## 8. Rollback

Delete `GetDestMatch`, `DestExists`, `GetSourceContent`, `UpdateFile` and `DestItemId`, and repoint the
tail back to `Copy file`. Behaviour returns to 1.0.336.0's: replace by delete-and-recreate, old
document in the recycle bin for 93 days.

---

## 9. On SDG's tenant

**Sign in as the SERVICE ACCOUNT before adding the first action.** A flow runs under its connection,
the connection is created implicitly by the first action ADDED, and a flow built as a person stops
**silently** when that password changes. Editing an existing action does not re-create the connection;
adding one does — and this change adds five.
