# Pause uploads while the folder structure changes

**Date:** 2026-08-19
**Status:** BUILT 1.0.180.0, NOT site-tested.
**Extends** `2026-08-11-subtree-migration-design.md` and `2026-08-14-folder-management-guided-flows-design.md`.

---

## 1. The gap, and why the flow made it visible

The "Change the folder structure" flow opened with a step headed **Power Automate (optional)** and
said nothing about uploads. Clarence, on seeing it: *"they are not going to go to Pause power
automate, so that is why I was thinking of building that toggle to stop user from uploading."*

A file uploaded **during** a migration lands in the **old** shape — correctly, because `Levels` is
still the old chain until the run's last step. The problem is timing: if it arrives after the
migrator has scanned its folder, it is never moved.

**The staged-apply model already stops that going live silently.** The pending chain is applied only
when a **fresh scan finds no drift**, so a slipped file leaves the migration *pending* rather than
completing over a stray. That bounds the damage to "the migration will not finish", which on a live
site can loop indefinitely — an admin re-running while uploaders keep filing.

## 2. Site-wide, not per segment

Per segment is narrower and was offered. Clarence chose site-wide: *"after work hours no one is going
to upload, once its ready and file is correctly moved then they turn it back on manually."*

The trade accepted: a structure change on one segment stops uploads for all of them. It is acceptable
**because of when the work happens**, not because the blast radius is small — and that is worth
restating whenever someone proposes running a migration during the day.

## 3. Two checks, and the second is the one that matters

- **At mount**, for the banner: an uploader learns before filling the form in, not after choosing a
  file and pressing Upload.
- **Immediately before writing**, re-read: settings are read once in a mount-time effect (gotcha
  #10), so a tab opened *before* the pause would otherwise upload straight through it.

The write-time check is what makes this more than decoration. It is the same shape as the stale-chain
guard, for the same reason.

### 3.1 ⚠ It cannot reach a session that never talks to the server

A tab open before the pause holds its own state, and no config row changes that. The write-time
re-check closes most of the window — the document is **refused rather than misfiled** — but the
operational answer is the client's: make structure changes outside working hours. **Say this to the
client rather than implying the toggle is complete protection.** It is stated on the screen itself.

## 4. Failure rules

| Condition | Behaviour |
|---|---|
| Config row missing | **not paused** — the setting has never been used on this site |
| Row blank or unrecognised | **not paused** |
| Read fails | **not paused** |
| Toggle screen cannot read the row | says so in red, and does NOT claim uploads are on |
| Write fails | reports the status and says **uploads are UNCHANGED** |

**⚠ FAILS OPEN, against the instinct that a safety switch should fail safe.** A wrong `true` blocks
every uploader on the site over a transient read; a wrong `false` costs one file in the old shape,
which the fresh-scan guard already catches. Same reasoning as gotcha 10b.

**But the SCREEN fails closed in what it claims.** An unreadable setting shows an error, never "on"
— because an admin who believes uploads are paused when the row could not be read will migrate on
top of live traffic. The value fails open; the *assertion about* the value does not.

## 5. Where the control lives

Two steps in the structure flow, **one component**, `mode` changing only the wording:

1. **Pause uploads** — now the first step, ahead of the Power Automate one
2. **Turn uploads back on** — a step of its own at the end

The closing step exists because **turning it back on is what gets forgotten**, and a forgotten pause
is a DMS that quietly accepts no documents while showing a banner that makes it look deliberate. The
flow does not end until someone has looked at that screen.

Resuming writes `no`, never a blank cell: blank cannot be told from never-set, and the point of the
step is that an admin can confirm they did it.

## 6. Files

| File | Change |
|---|---|
| `src/shared/uploadPause.ts` | **new, pure** — `uploadsArePaused`, `pauseSettingValue`, the setting name and message |
| `src/shared/uploadPause.test.ts` | **new, 8 tests** — including that missing/blank/unreadable never pauses |
| `src/webparts/userAccess/components/UploadPauseToggle.tsx` | **new** — the screen, mounted twice |
| `src/webparts/form/components/Form.tsx` | live read at mount + before writing; banner |
| `src/webparts/bulkUpload/components/BulkUpload.tsx` | same |
| `src/shared/folderFlows.ts` | two steps; `StepScreen` gains the `pauseUploads` component id |
| `src/webparts/folderManager/components/FolderAdmin.tsx` | mounts the component |

**`FolderManager.tsx` is deliberately untouched** — it holds the folder-creation and group-assignment
code frozen after the 2026-08-19 verification, and this needed nothing from it.
