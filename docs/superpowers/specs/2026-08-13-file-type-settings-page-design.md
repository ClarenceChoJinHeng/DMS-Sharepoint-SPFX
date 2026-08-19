# File type settings — a page for the allowed-extensions policy

**Date:** 2026-08-13
**Status:** BUILT — `crsConfiguration/components/FileTypeSettings.tsx`. Site-test state not recorded here; CLAUDE.md is the live record. Header corrected 2026-08-19.

Builds an admin UI for the data specced in
[2026-07-30-allowed-file-types-dropdown-design.md](2026-07-30-allowed-file-types-dropdown-design.md),
which made the `AllowedFileTypes` column the single source of truth but left it to be ticked in the
SharePoint list view.

---

## 1. Why

The client decides which file types may be uploaded. Today that decision is expressed by ticking boxes
on a multi-select Choice column in a list view, and *adding* a type they cannot already tick means
editing the column's Choices in list settings. The client will do neither after handover.

Two failure modes are invisible in the list view, and both mislead the person looking at it:

- **A typo'd extension fails CLOSED.** `.xlxs` looks enabled; nobody can upload an Excel file, and the
  uploader is told the file type is not permitted — which reads as policy, not as a mistake.
- **A missing column looks like an empty policy.** The upload form silently falls back to a built-in
  list, so the client believes their policy is in force when it is not.

## 2. Scope — file types only

The web part is titled **`CRS Configuration`** because it is a container that may later hold other
settings from `CRS Config` (`recon_gridMode`, `legallyPrivilegedFor`, the term-set GUIDs). It ships with
**one panel: File type settings**, and no tab framework. A second panel becomes a sibling component when
a second panel exists — not before.

**Its own web part**, not a sixth tab on Folder Administration: the audience and the task are different,
`CRS Settings` in the site navigation already implies its own page, and that file is already 4,100 lines.

## 3. THE LIST IS NOT A SECURITY BOUNDARY

Stated first, because designing it as one would be the real mistake. The allowed-extensions list filters
the file picker and validates in JavaScript before upload. It stops nothing determined: a file can be
renamed, and anyone with Contribute on a folder can upload through the SharePoint UI without touching the
form.

What actually protects the repository is already built and untouched by this page: folder ACLs so a PIC
can only write into their own unit, content approval so nothing reaches `Documents` unreviewed, plus
tenant-level blocked file types and Microsoft's own malware scanning.

**Hiding this page protects nothing.** A toggle needs only item-edit on `CRS Config`, so whoever can edit
that list controls upload policy whether or not they can see the page. On `/sites/ClarenceDMSTesting`
non-admins hold **Read** at site level and the four CRS lists inherit it (verified 2026-08-13), so the
boundary holds today — but it is the LIST's permissions holding it, not the page's.

## 4. The data, unchanged

Two reads at mount, after `primeNames()`:

| Read | Answers |
|---|---|
| the **field** `AllowedFileTypes` on `CRS Config` → `Choices` | every extension that can be chosen, and whether the column exists at all |
| the **item** `allowedExtensions` → its value | which of them are currently allowed |

Rows on screen are the union: every choice, with a toggle showing whether it is ticked.

**Names are resolved, never hardcoded** — the list is `CRS Config` here and `DMS Config` on a migrated
site, so the page primes first and reads through `cachedListTitle()`. An unprimed cache resolves to the
legacy title and 404s; that exact bug shipped on the abbreviation editor and surfaced only when a page
load lost the race.

**The column is matched by internal name**, not display name. A matching display name over a different
internal name fails exactly like an absent column and looks correct — the trap that left `Documents`
missing four columns for months.

## 5. Two writes, two permissions

The two actions look alike on screen and touch different things:

| Action | Target | Permission |
|---|---|---|
| Toggle a type on/off | the **item**'s `AllowedFileTypes` value | item edit |
| Add a new type | the **field**'s `Choices` | **Manage Lists** |

The field update needs `__metadata` typed `SP.FieldMultiChoice`, so that one call goes `odata=verbose`
while everything else stays `nometadata`. A 403 on it reports *"adding a new file type needs site owner
rights"* rather than a generic failure — the two operations fail differently, for different reasons.

**`FillInChoice` stays OFF.** Enabling it restores the free-text typo risk this page exists to remove.

## 6. Toggles save immediately, and a failed write flips back

Each toggle writes on click. No Save button, no dirty-state guard: the page has exactly one kind of edit
and each row is independent, so there is nothing to batch. (The abbreviation editor needs a Save because
a collision spans rows; nothing here spans rows.)

**A failed write must revert the toggle.** A switch that stays on after a failed save is a UI lying about
a permission setting, and the lie survives until someone reloads.

## 7. Turning off the last type is allowed, with a warning

All-off is a **supported, deliberate state**: both upload forms hard-block and name the column and list
to fix, rather than silently falling back to built-in defaults. So the page does not forbid it.

Turning off the **last enabled** type opens a confirm naming all three consequences: nobody can upload
anything in any unit, existing files are untouched, and re-enabling any single type restores uploads
immediately.

Considered and rejected: refusing to leave it empty. Six deliberate clicks and one accidental last click
look identical afterwards — but so does a deliberate freeze, and an admin who means it should not have to
hand-edit a list to get there.

Also rejected: the designer's **"Restore defaults"** button. Nobody clicking it could predict what it
would do — the code's fallback list differs from the live policy, so it would have turned `.docx` on and
`.png`/`.jpeg` off. Its actual intent was "disable everything", which the toggles already express.

## 8. Adding a type

Input is normalized before anything else: trim, lowercase, **force a leading dot**. `txt` and `.txt` both
become `.txt`.

The dot is enforced rather than required because a dotless token is the one input with two opposite
behaviours: the file picker parses `png` as a MIME type, finds no `/`, and silently drops it — greying
that type out of the dialog — while the JS validator's `endsWith` check accepts it. One typo, half a day
(2026-07-30).

Rejected outright: spaces, a second dot, an empty result.

Then classified against the existing choices:

| Input | State | Behaviour |
|---|---|---|
| `.txt` | not a choice | add the choice **and enable it** — it was added in order to be allowed |
| `.docx` | a choice, currently **off** | *"`.docx` is already in the list."* No change; the page scrolls to and highlights that row so it need not be hunted for |
| `.pdf` | a choice, currently **on** | *"`.pdf` is already allowed."* No change |

The already-exists cases must never add a second choice. SharePoint will hold two `.docx` entries
happily, the page would render two identical rows with independent toggles, and one of them would appear
to do nothing.

**The middle case does not auto-enable** (client, 2026-08-13). It reports the truth and leaves the
decision to the admin.

## 9. Did-you-mean, for near misses

An unrecognised extension is compared against the known-types map at **edit distance 1**. On a hit:

> `.xlxs` isn't a known file type — did you mean `.xlsx`?

Both options are offered; neither is forced. This is **usability, not security** — a near miss fails
closed, blocking a legitimate type, which is annoying and never dangerous. It earns its place by stopping
an admin from believing a policy is in force when it is not: the same failure mode as §1's missing column.

Nonsense that resembles nothing (`.zzz`) is allowed. There is no list of all real extensions, and an
extension that matches no file simply never matches a file.

**Unrecognised types are labelled.** The description column reads *"Custom file type"*, so a row meant to
be Excel that says *"Custom file type"* is its own warning.

## 10. Executable types are refused — for BOTH actions

A short explicit list cannot be added **or enabled**:

```
.exe .com .bat .cmd .msi .dll .scr .pif .cpl .ps1 .psm1 .vbs .vbe .js .jse
.wsf .wsh .hta .reg .lnk .jar .app .sh
```

- Typing one → refused, with a message saying it cannot be allowed in a document repository.
- One already in the list → the row renders **locked**, marked *"Blocked by policy"*. Visible, not
  quietly absent, so nobody wonders why a toggle is missing.
- One already **ticked** → said loudly, with an offer to untick it. That is a live state nobody chose on
  this screen, and it can arrive from a hand edit or a migration.

**The rule must cover enabling, not just adding.** Blocking only the add path would leave the one route
that matters unguarded: a column that already carries `.exe` from outside this page.

This is defence in depth, not protection — see §3. It stops a well-meaning admin from advertising the
repository as a place to share executables, and that decision is hard to un-make once files exist.

**Archives are NOT blocked** (client, 2026-08-13). `.zip`, `.7z` and `.rar` have legitimate uses. The
consequence is recorded rather than enforced: an archive can carry any of the types above, so allowing one
makes this whole list advisory.

## 11. Three states, and empty is not unknown

| State | Detected by | The page shows |
|---|---|---|
| `no-column` | the `$select` on `AllowedFileTypes` fails (HTTP 400) | *"This site has no `AllowedFileTypes` column, so uploads are running on the built-in list below. To manage them here, add a multi-select Choice column named `AllowedFileTypes` to `CRS Config`, with 'Allow fill-in choices' off."* — plus the built-in list rendered **read-only**, so what is actually in force is visible |
| `all-blocked` | column present, no values ticked | *"No file types are allowed — uploads are blocked for everyone."* No fallback list, because none is in force |
| `normal` | column present, ≥1 ticked | the rows |

The first two look identical in a list view and mean opposite things — one a silent default, the other a
deliberate block. Distinguishing them is only possible because the code tests the **presence of the key**,
not the value: an emptied multi-Choice column returns `null`, not `[]`.

**The page never creates the column** (client, 2026-08-13). It explains what to create, precisely enough
that the message alone suffices, and shows the built-in list standing in for it. Considered: creating it
with one click, since the column has already been missed once on this site and the migration to SDG's own
site is a single high-stakes event. Rejected in favour of a page that never writes schema.

## 12. What was cut from the mockup, and why

| Cut | Reason |
|---|---|
| **Actions** column | empty in the design; toggling off already blocks uploads. Deleting a choice would strand any value still ticked — allowed-looking and never re-selectable |
| **Search box** | 5–15 rows. Costs a control, earns nothing. Trivial to add if the list grows |
| **"Restore defaults"** | §7 |
| **File-type icons** | replaced with a text badge (`PDF`, `XLSX`). The package is deliberately self-contained; icons mean bundled assets or external URLs, for decoration |
| *"Changes take effect immediately"* | untrue — both upload forms read settings once at mount, so an open tab keeps the old list. Reworded: *"Applies to new uploads. Anyone with the upload form open will need to refresh."* |
| "extention", "Files with disable extentions" | spelling |

The **Description** column stays, but is **derived in code**. The column stores only the extension string;
there is nowhere to put a description or an icon. Known types come from a lookup map; everything else
reads *"Custom file type"*.

## 13. Structure

- **`src/shared/fileTypeSettings.ts`** — pure, no SharePoint imports, unit-tested:
  - `normalizeExtension` — trim, lowercase, force leading dot; rejects spaces, a second dot, empty
  - `classifyAddition` → `new` | `exists-enabled` | `exists-disabled` | `blocked` | `invalid`
  - `nearMiss` — edit-distance-1 match against the known-types map
  - `describeExtension` — map lookup, else *"Custom file type"*
  - `isBlockedType` — the §10 list
  - `panelState` → the three states of §11, from `(columnPresent, choices, ticked)`
- **`src/webparts/crsConfiguration/components/FileTypeSettings.tsx`** — the panel: reads, writes,
  confirms, toasts.
- **`src/webparts/crsConfiguration/CrsConfigurationWebPart.ts`** + manifest, title `CRS Configuration`.

Reuses `FALLBACK_FILE_TYPES` and the normalization already in `src/shared/allowedFileTypes.ts` rather than
restating either — two definitions of "what is a valid extension" would drift, and the reading half
already lives there.

**Deployment leaves two manual steps**, both easy to forget because nothing fails visibly without them:

1. The page hosting the web part needs a **Page-scope Group Map row** limiting it to admins, like the
   other `CRS Settings` pages. `Target` matches on `FileLeafRef` — the page's FILE name, fixed at
   creation and unchanged by retitling — so a typo writes a row that is re-asserted forever against
   nothing.
2. `CRS Config` should keep **version history on**, so a change to upload policy is attributable. It is
   the only audit trail this page will have; the page itself records nothing.

## 14. Testing

Unit (pure rules):

- `png` → `.png`; `.PNG` → `.png`; ` .pdf ` → `.pdf`
- `a.b`, `my file`, `.`, `` → invalid
- `.xlxs` → near miss on `.xlsx`; `.zzz` → no near miss, allowed, described as *"Custom file type"*
- `.exe`, `.PS1` → blocked; blocked-and-ticked reported, not silently dropped
- `classifyAddition` returns `exists-disabled` for a choice present and unticked — and never produces a
  second choice
- `panelState`: absent column → `no-column`; present with `null` value → `all-blocked`; present with one
  value → `normal`

Site:

- Toggle a type off, hard-refresh the upload form: that type is gone from the picker and rejected by the
  validator.
- Toggle it back on: accepted again.
- Turn off the last enabled type: the warning names all three consequences; confirming leaves the upload
  form hard-blocking with a message naming the column and list.
- Add `.txt`: appears enabled, uploadable after a refresh, and the field's Choices carries exactly one
  `.txt`.
- Add `.docx` while `.docx` is present but off: message, no change, row highlighted, still one choice.
- Add `.exe`: refused.
- Break the write (remove item-edit rights): the toggle flips back and an error is shown.
- Rename the column, reload: `no-column` state with the built-in list read-only, and the upload form still
  working on the fallback.
- As a non-admin with item edit but not Manage Lists: toggles work, Add reports the rights it needs.

## 15. Related

- [2026-07-30-allowed-file-types-dropdown-design.md](2026-07-30-allowed-file-types-dropdown-design.md) —
  the column this edits: hard-block on empty, fallback on absent, `FillInChoice` off
- [2026-08-12-term-abbreviation-page-design.md](2026-08-12-term-abbreviation-page-design.md) — the sibling
  admin editor whose priming bug §4 avoids
- [2026-07-28-client-site-migration-runbook.md](2026-07-28-client-site-migration-runbook.md) — carries the
  manual "add the AllowedFileTypes column" step that §11 explains rather than automates
