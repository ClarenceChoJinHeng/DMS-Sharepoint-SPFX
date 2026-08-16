# CRS Term Manager — authoring the term tree without the native UI

**Date:** 2026-08-17
**Status:** Design agreed. Not yet built.
**Placement:** a tab in **Folder Administration**, immediately before *Term Abbreviations*.
**Write path:** legacy CSOM `ProcessQuery` — verified live 2026-08-17.

---

## 1. Why

> "We got to create a term store UI for client for an easier way to force them to ensure they know
> how to add since the native UI Term store doesn't tell them that."

The native term store manager is correct and completely uninformed about CRS. Every mistake it lets
an admin make fails **silently**:

| Mistake in the native UI | What actually happens |
|---|---|
| Term added with no abbreviation | Reconciliation SKIPS it. No folder, no error, nobody can upload there. |
| Term deleted and re-added | New GUID. Abbreviation, Folder Map and Group Map rows all orphan at once. |
| Term added one level too deep | Below the permissioned boundary, so no folder and no warning. |
| `&` typed instead of `＆` | Accepted, and inconsistent with every other term. |
| Subunits authored in their own set | Every unit is offered every other unit's subunits. |

None of these produces an error message anywhere in the system. The whole value of this page is that
it **knows the rules the native UI cannot know**.

---

## 2. The write path, and why it is the OLD API

**Reads stay on `/_api/v2.1/termStore/...`** — every existing resolver uses it and none of that
changes.

**Writes MUST use `/_vti_bin/client.svc/ProcessQuery` (CSOM).** Verified 2026-08-17: a `POST` to the
v2.1 endpoint from a signed-in browser is refused with

```
403  {"error":{"code":"notAllowed","message":"OAuth only flow is enabled and the call has not been issued from an app."}}
```

The modern taxonomy API accepts writes only from an app-authenticated OAuth context. SPFx runs on the
user's cookie plus a request digest and can never satisfy that. Graph is not an option either — it
needs `TermStore.ReadWrite.All`, and API permission requests are processed from the **tenant** app
catalog while this package ships to the **site collection** one, so it cannot even ask. The same dead
end as Purview.

CSOM works because it is what SharePoint's own term store page uses.

**Therefore: `shared/taxonomyWrite.ts` is the ONLY module that emits CSOM.** Two protocols against one
store is exactly the asymmetry someone later "tidies up" onto the modern endpoint, producing a 403
whose message explains nothing. The module header states the reason and cites this spec.

Verified shape:

```
GetTaxonomySession → GetDefaultSiteCollectionTermStore → GetTerm(parentGuid)
  → CreateTerm(name, 1033, newGuid) → CommitAll
```

- **`CommitAll` is required.** Without it the write is discarded with no error.
- GUIDs are brace-wrapped: `{fdafa271-…}`.
- **Success is `"ErrorInfo":null` inside an HTTP 200, never the status code alone** — a CSOM failure
  returns 200 with the error in the body. Same class as gotcha #4.

---

## 3. What it shows

The tree for one segment at a time, **labelled by tier from that segment's `Levels` chain** — so GHO
reads Department / Unit and Upstream Ops reads Region / Estate·Mill. Never a hardcoded pair: the admin
names the tiers at onboarding, and a written-out list is only ever right for the segments it was
written for. The same rule as the details panel and the search filters.

Below the permissioned depth the tier is labelled **SubUnit** (or whatever the chain calls it), and
the page says plainly that it needs no group, no abbreviation and no folder of its own.

Each term row shows:

- its label
- its abbreviation, or **`no code — reconciliation will skip this`** in amber, for permissioned tiers
- nothing for below-Unit tiers, which name their folders from the label

That missing-abbreviation warning is the single most valuable thing on the page: it is the failure
that produces no folder, no error, and no way to discover the cause.

---

## 4. What it does

### 4.1 Add a term

At any depth the chain allows. The parent is always explicit — you add *under* something, never into
a blank field, because "which parent" is the thing the native UI makes easiest to get wrong.

- `&` is normalised to the fullwidth **`＆`** SharePoint requires, and the page says it did so.
- A duplicate sibling label is **blocked before the write**, not after: two siblings with one name is
  how two units end up sharing one folder and one ACL.
- After a successful add at a permissioned tier the page offers the abbreviation **inline**. The
  correct next action is always "give it a code", and sending them to another tab is how it gets
  forgotten.

### 4.2 Rename a term

Allowed, encouraged, and safe — worth saying on screen, because it looks dangerous and is not:

- Folder names come from **abbreviations**, not labels, so a rename does not rename a folder.
- The GUID is unchanged, so Abbreviation, Folder Map and Group Map rows stay joined.
- What does change: dropdown labels, and the `Full Name` column on the next reconciliation.

### 4.3 Reorder

Sets the parent's custom sort order. Cosmetic — it changes dropdown order and nothing else. Offered
because term order is the client's first complaint about any picker, and it is otherwise buried.

### 4.4 Delete — NOT OFFERED

Deleting a term orphans three lists at once, and a re-created term gets a **new GUID** that rejoins
nothing. Reconciliation repairs an orphan only on a 1:1 match of level + label and refuses ambiguous
ones — `Tax`, `Legal` and `PM` each exist under several parents, and a wrong re-point is a permissions
grant to another department's folder.

So the page **has no delete button, and explains why in place**. Silence would send them to the native
UI to do it there.

If a term must go out of use the answer is to rename it, never to remove it.

---

## 5. What it refuses

- **Adding below the chain's last tier.** Nothing reads it, so it would create a term that can never
  appear in any dropdown.
- **A blank or whitespace-only label.**
- **A duplicate sibling label**, compared case-insensitively and after `＆` normalisation — because
  that is what a collision actually is.
- **Anything when the term set cannot be read.** Adding into an unknown tree is how duplicates happen.

Every refusal names the conflicting term and its parent. A refusal that only says "invalid" sends them
to the native UI to do the same thing unguarded.

---

## 6. Failure states

Three, distinguished, as everywhere else in this codebase:

1. **Not read yet** — the resting state.
2. **Read, and this parent has no children** — a real and common answer, especially for units with no
   subunits.
3. **Could not read** — names the status. Never rendered as "no terms": an admin told a department is
   empty will add terms that already exist, and duplicate siblings are the one collision that merges
   two units into a single folder.

A CSOM write returning 200 with a non-null `ErrorInfo` is a **failure** and is reported as one.

---

## 7. Placement

A tab in **Folder Administration**, before *Term Abbreviations*, matching the order the work happens:

```
New segment → Term store → Term Abbreviations → Folder levels → Move existing folders → Reconciliation
```

It also fills the guided flows' existing **"Create the term set"** step, which today can only tell the
admin to go and do it somewhere else — the one step in the everyday "add a department or unit" flow
that leaves the tool.

Mounted from one component, as `GroupManager` and `GroupMapBuilder` already are, so the tab and the
flow step cannot drift apart.

---

## 8. Out of scope for the first build

- **Creating term sets or term groups.** Needs `CreateTermSet` on the site-collection group and a
  different CSOM shape; segment onboarding already assumes the set exists. Worth adding once term
  creation is proven in use.
- Moving a term to a different parent — a re-parent is a folder move, and the subtree migrator owns
  that class of problem.
- Deprecating / `isAvailableForTagging`.
- Translations and synonyms.
- Any term group other than the site-collection one.

---

## 9. Risks

- **CSOM is undocumented in modern SharePoint and returns 200 on failure.** Mitigated by isolating it
  in one module, asserting `ErrorInfo === null`, and re-reading the tree after every write rather than
  trusting the response.
- **The term store is site-collection local** (2026-07-27 pivot), so `GetDefaultSiteCollectionTermStore`
  is correct here and would be wrong against a tenant store. Named in the module, never assumed.
- **A write needs term-store contributor rights.** A site collection admin has them; a plain site owner
  may not. A 403 must be reported as *"you do not have permission to change the term store"*, not as a
  generic failure — that distinction is the difference between a five-minute fix and an afternoon.
