# Site Entry Access Layer — Design

**Date:** 2026-07-27
**Status:** Approved, implementing
**Site:** `/sites/ClarenceDMSTesting`
**Related:** memory `dms-two-layer-access-site-plus-folder` (needs correction — see §7),
`dms-group-model-per-role-per-library`, RBAC table in `CLAUDE.md`.

---

## 1. Problem

Folder-scoped SP groups grant only **Limited Access**, which lets a user *traverse to a
specific folder via a direct link* — it does **not** grant Read on the site root / Home
page. Result: a user added to (only) a unit folder group hits **Access Denied** at the site
URL. Verified live: `Bayajit@trinergydigital.com`, added to `DMS_GHO_GF_FO_UPL`, still
denied at `/sites/ClarenceDMSTesting`.

There is currently **no site-entry layer** — nothing grants "open the site + see Home."

## 2. Target access matrix

| Role | Home page | Documents library | Staging library | Upload |
|------|-----------|-------------------|-----------------|--------|
| Normal member (entry only) | ✅ | ✅ (their dept) | ❌ | ❌ |
| Uploader (`_UPL`) | ✅ | ✅ | ✅ (their unit) | ✅ |
| Approver (`_APR`) | ✅ | ✅ | ✅ (their unit) | approve |
| Global reader (`DMS_GLOBAL_READER`) | ✅ | ✅ **all depts** | ✅ **all, read-only** | ❌ |

## 3. Design — three layers

1. **Site entry** (NEW): a plain SP group **`DMS_SITE_MEMBERS`**, granted **Read** on the
   web root + Home page + Documents library. Every DMS user is a member. This is the piece
   that finally lets anyone open the site. **Not** in the DMS Group Map (see §5).
2. **Documents library**: per-department folder isolation stays (base folder group grants
   Read on the user's dept folder). `DMS_SITE_MEMBERS` gets **Read at the library level** so
   the library opens and users can browse to their folder. Trade-off accepted: other-dept
   *folder names* are visible, but opening them is denied by the per-folder broken
   inheritance.
3. **Staging library** — the isolation guard: **break inheritance at the *library* level**
   (not just per-folder) and **remove all site groups** (`DMS_SITE_MEMBERS`, Visitors).
   After that the only way into Staging is the `_UPL` (Contribute) / `_APR` (Design) folder
   groups, plus `DMS_GLOBAL_READER` granted **Read at the Staging library level** for the
   read-only auditor role.

   > ⚠️ **Critical, non-skippable.** Site-level Read inherits *down* into any library that
   > still inherits. If Staging is left inheriting, every `DMS_SITE_MEMBERS` user would
   > *see* (and download) all pending Staging docs. They could never upload (Read ≠
   > Contribute), but the visibility alone breaks isolation. Breaking Staging at the library
   > level is what closes that leak.

## 4. `GLOBAL` role semantics — bug fix

`DMS_GLOBAL_READER` is intended as a **read-only super-viewer** (reads *all* of Documents +
Staging; **cannot upload**). The current code contradicts that:

- `collectMembership` sets `isGlobalUploader = true` for any `Role=GLOBAL` row
  (`src/shared/formModel.ts:138`).
- `isGlobalUploader` → `privileged` → **bypasses tier detection and may upload anywhere**
  (`src/webparts/form/components/Form.tsx:669`, `:277`).

So today a member of `DMS_GLOBAL_READER` is silently an **upload-anywhere** user — the exact
opposite of intent.

**Fix:** `Role=GLOBAL` must confer **no upload capability**. It is a pure read role, enforced
by SharePoint library-level Read grants (§3), invisible to the upload form. Site admins keep
upload-anywhere via the separate `admin` check — unaffected.

- `collectMembership`: drop `isGlobalUploader`; `GLOBAL` rows are ignored (no term, no
  upload). A user whose *only* role is `GLOBAL` has zero upload paths → the form shows its
  normal "you have no upload permissions" state, which is correct for a read-only user.
- `Form.tsx`: `privileged = admin` only.
- Update `Membership` interface + tests.

## 5. Group Map is NOT the site-entry mechanism

The DMS Group Map exists for **one purpose: routing uploads**. Site entry is a pure
SharePoint permission (a Read grant + group membership) that the form never reads. Therefore:

- `DMS_SITE_MEMBERS` is **not** added to the Group Map. Registering it there buys nothing.
- `DMS_GLOBAL_READER` may keep its Group Map row, but after §4 that row is **inert in the
  form** (GLOBAL no longer routes or grants upload). Its actual power comes from the
  SharePoint library-level Read grants. Keeping the row is optional/documentary.

## 6. Automation — auto-add to `DMS_SITE_MEMBERS`

Users get into groups via the **Members modal** in the Group Map Builder
(`GroupMapBuilder.tsx` → `mmAdd`, which calls `addGroupMember`). Change: whenever a user is
added to **any** DMS group via that modal, **also** add them to `DMS_SITE_MEMBERS`, so
onboarding stays a single action and nobody is left unable to enter the site.

- Resolve the entry group's id by name (`fetchAllSiteGroups` → find `DMS_SITE_MEMBERS`),
  cached per session.
- Skip the auto-add when the group being edited **is** `DMS_SITE_MEMBERS` itself.
- SP `addGroupMember` is idempotent for an existing member — re-adding is a no-op, safe.
- If `DMS_SITE_MEMBERS` doesn't exist yet, show a non-blocking toast ("entry group not
  found — create DMS_SITE_MEMBERS") and still complete the primary add.
- Entry group name is a single shared constant `SITE_ENTRY_GROUP_NAME` in `groupMapModel.ts`.

## 7. Memory correction

`dms-two-layer-access-site-plus-folder` says folder Limited Access "is enough to enter the
site." Correct it: Limited Access reaches an **assigned folder via direct link**, but the
**site Home / root still denies** — which is why this entry layer exists.

## 6a. One-click provisioner (designed, then REMOVED 2026-07-28)

> **Status: not shipped.** Built into the Onboarding web part, then removed at the user's
> request — a site-permissions provisioner didn't belong on the Folder Onboarding surface and
> was confusing. Site-entry setup is **manual for now** (see §8). The design below is retained
> for if it's rebuilt later on a more appropriate surface (e.g. its own admin tab).

Rather than manual clicking, the **Onboarding web part** gets a "Set up site entry" panel
(a one-time site-setup action; runs as the current admin, needs Full Control / Manage
Permissions). Steps, each logged, behind a confirm dialog:

1. **Ensure `DMS_SITE_MEMBERS`** — `web/sitegroups/getbyname` → if missing, `createSiteGroup`.
2. **Staging isolation** — if `Staging` library still inherits
   (`HasUniqueRoleAssignments === false`), `breakroleinheritance(copyRoleAssignments=true,
   clearSubscopes=false)`. `copy=true` preserves current access (non-destructive);
   `clearSubscopes=false` leaves the per-unit folder scopes untouched.
3. **Remove the Read leak from Staging** — remove the site **Visitors** group
   (`AssociatedVisitorGroup`) and `DMS_SITE_MEMBERS` (if present) from the Staging library
   scope, so web-level Read cannot surface Staging. Owners/Members retained; final Staging
   assignments are logged so the admin can eyeball.
4. **Grant `DMS_SITE_MEMBERS` Read on the web root** — reaches Home + Site Pages + Documents
   by inheritance; Staging is now excluded (step 2).
5. **Grant `DMS_SITE_MEMBERS` Read on the Documents library** explicitly (idempotent — covers
   the case where Documents has its own unique scope).

Order matters: break/clean Staging **before** granting web Read, so the entry group's Read
never transiently reaches Staging. Re-runnable (idempotent): re-adding a member/assignment is
a no-op; the break is skipped when the library is already unique.

## 7a. Native-add safety (client adds users the SharePoint way)

Native group edits **cannot corrupt the flow** — the form and reconciliation read live group
membership every load, so access is always computed from actual memberships regardless of how
a user was added. The only failure mode is **incompleteness**: adding someone natively to a
role group but **not** `DMS_SITE_MEMBERS` leaves them with folder Limited Access but no site
entry (Access Denied at Home) — the Bayajit case. Guidance: onboard via the web part (one
click does both); if going native, add to `DMS_SITE_MEMBERS` too.

**Self-heal (built 2026-07-28):** Folder Reconciliation now ends with a site-entry sweep —
it reads every `DMS_*` site group (except `DMS_SITE_MEMBERS`) and adds any member missing from
`DMS_SITE_MEMBERS` into it, logging the count. So an incomplete native add self-heals on the
next reconciliation run. If `DMS_SITE_MEMBERS` doesn't exist yet, the sweep logs a warning and
is skipped (run the provisioner first). See `runReconciliation` in `FolderManager.tsx`.

## 8. Manual SharePoint config (fallback if the provisioner isn't used)

1. `⚙ → Site permissions → Advanced → New → Group` → create **`DMS_SITE_MEMBERS`**, grant it
   **Read** at the site level.
2. **Documents** library → Settings → Permissions → grant `DMS_SITE_MEMBERS` **Read** at the
   library level. Keep per-dept folder inheritance broken.
3. **Staging** library → Settings → Permissions → **Stop Inheriting Permissions** (library
   level) → **remove** `DMS_SITE_MEMBERS` / Visitors / all site groups. Confirm only
   `_UPL` / `_APR` folder groups remain (add `DMS_GLOBAL_READER` **Read** here for the
   auditor role).
4. Add every existing DMS user to `DMS_SITE_MEMBERS` (going forward the web part does this).

## 9. Out of scope

- No change to `_UPL` / `_APR` folder-group creation or reconciliation.
- No Graph, no admin center.
