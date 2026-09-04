# Reconciliation re-issues every folder grant on every run

**Date:** 2026-08-19
**Status:** BUILT 1.0.177.0, NOT site-tested.
**Register:** #18b — the half of selective reconciliation that #18's segment picker cannot reach.
**Extends** `2026-08-19-selective-reconciliation-design.md`. That spec narrows a run to the ticked
segments; this one narrows it to the work actually outstanding *within* them.

---

## 1. The measurement that provoked it

Minamas Head Office, first provision, one segment only after #18: **71 minutes, 2,441 steps.** The
client's reaction — *"71 minutes is still long"* — is correct, but the first run is not where the
waste is. Those 2,441 folders and role assignments genuinely did not exist.

The waste is the run **after** it. That is the run the client described when they asked for #18:
adding one unit under one department, and paying the same hour to do it.

### 1.1 The first diagnosis was wrong, and is recorded because the reasoning was plausible

The Year × Document Type grid is the documented 98% of a full run, so it was the obvious suspect. It
was already `off` on this site — `DEFAULT_GRID_MODE` is `off`, and the config row confirmed it. **A
known-expensive mechanism being present in the code is not evidence it is running.** Check the config
row and the run's own opening lines before optimising against it.

## 2. The cause

`FolderManager.tsx`, in the folder pass, called `addRoleAssignment` **unconditionally** for every
group on every folder in every library on every run. SharePoint accepts a duplicate assignment as a
no-op, so nothing ever failed — it cost a round trip and a throttle tick each time and produced a log
line saying the grant had been made.

For MHO that is roughly **49 units × 5 personas × 4 libraries ≈ 980 writes**, plus the department and
segment tiers, re-issued on every future run to set permissions already exactly right.

**The asymmetry is what let this survive.** The ancestor-browse grant, twenty lines below in the same
loop, already skipped what was in place:

```ts
if (existing.some(a => a.principalId === gp.pid && a.roleDefId === readId)) continue;
```

One loop checked its neighbour's work and the other did not. The only symptom was time, and time
reads as SharePoint being slow rather than as a defect.

## 3. The change

Skip a grant when the folder already holds that exact `(principal, role definition)` pair.

### 3.1 Read the existing ACL only where it can pay

Reconciliation resets the ACL itself in two of its three branches — a folder it creates, and a folder
whose inheritance it breaks with `copyRoleAssignments=false`. Both leave **only site Owners**, so
every planned grant is genuinely missing and a read could return nothing useful. Reading there would
add a request per folder to the *first* run — the expensive one — to save nothing.

The third branch, `already locked, skipped`, is the entire steady state of a re-run. There **one read
replaces up to five writes** per folder per library.

`shouldReadExistingAcl(aclWasReset, plannedGrants)` is that rule, pure and pinned by test.

### 3.2 ⚠ Unknown grants. Always.

`needsGrant(undefined, …)` is **true**. A read that failed says nothing about the folder, and the cost
of guessing "already granted" is a group that silently never receives its permission — presenting as
an uploader who cannot upload, on a run that reported success, after an admin has already done the one
thing they would be told to do.

`getRoleAssignments` folds a non-OK status into `[]` on its own, which fails in the same safe
direction (it grants). The wiring only has to catch a throw. **`[]` is a folder read and found empty;
`undefined` is a folder not read.** Empty ≠ unknown, in the place where the cost is a missing grant.

### 3.3 The match is exact on both halves

Never on the principal alone: a principal seen holding a *different* level is treated as ungranted and
re-granted. A wasted no-op write is the cheap error; a missing permission is not.

### 3.4 ⚠ It needs its OWN reader, and the first build did not have one — measured cost, 19m50s

The first build reused `getRoleAssignments`, which keeps only the **first** binding per principal. That
is right for its original callers, which want one level to display, and wrong here: SharePoint collects
all of a principal's permission levels into **one** role assignment with several
`RoleDefinitionBindings`. An approver group holding CRS Approve + CRS Delete + CRS Upload therefore read
back as holding one of them, and the other two were re-granted every run.

This spec first called that "one wasted no-op write in a rare case". **It is not rare — it is every
group with more than one role, which is every approver group.** The live log settled it: per unit folder
in the approval library, **five grants re-issued and three skipped**, the three being whichever level
happened to come back first for each of the three principals.

`getAllRoleBindings` returns every pair and **filters nothing**. The display reader drops
`RoleTypeKind` 1 and 7 as noise; here a dropped binding can only cause a needed skip to be missed, while
keeping every binding can only cause a wasted write. It matches on the raw `Id`.

### 3.5 A grant made in this run counts as held

Two roles can resolve to ONE permission level — `APR` and `UPL` are both Read on Documents, `DELS` and
`DELSHC` are both CRS Delete on the approval library. Each granted pair is pushed into the in-memory
list, so the second role skips.

That fixes a **reporting** problem as much as a performance one. The log prints one line per role and
names only the level, so the same group appeared twice against the same level on one folder — which
reads as a double grant. The client's words, 2026-08-19: *"its basically confusing the client."*

### 3.4 ⚠ A skipped grant still counts as granted for the ancestor-browse pass

`grantedPids` drives the corridor that lets a member click down from the library root to their own
folder. A skipped group **still holds its role**, so it still needs that corridor. Dropping it from
`grantedPids` would silently withdraw ancestor Read and render the library empty for its members —
the 2026-08-04 regression, reached by a new route. It is pushed in both paths.

### 3.5 One log line per folder, not per group

On a settled site every group on every folder is skipped. A line each would bury the run's real
findings under thousands saying nothing happened, which is the failure the run log already fights.
`↳ <folder> — N group grant(s) already correct, skipped`.

## 4. What does not change

- **Every grant is still asserted.** This changes how the required state is reached, never what it is.
  A grant removed by hand in SharePoint is absent from the read and is re-issued — which is the
  failure reconciliation exists to repair, and the reason detection was rejected for #18's Half B.
- **The first run of a segment is unchanged in cost**, deliberately (§3.1).
- **`bumpAssigns()` fires only on a real grant**, so the counter falls on a settled site. That is
  honest: it counts work done, and inflating it with no-ops is what hid this.
- **`step()` is still counted for a skipped grant**, so the progress estimate keeps its denominator.

## 5. Failure rules

| Condition | Behaviour |
|---|---|
| ACL read throws | `undefined` ⇒ **grant everything on that folder**, as today. |
| ACL read returns non-OK | `getRoleAssignments` yields `[]` ⇒ grant everything. Same direction. |
| Principal holds a different level | grant (§3.3). |
| Reconciliation just reset the ACL | no read; grant everything. |
| A grant fails | unchanged — reported, and the group is not added to `grantedPids`. |

## 6. Expected effect

| Run | Before | After |
|---|---|---|
| First provision of a segment | ~2,441 steps | unchanged |
| Re-run, nothing changed | ~2,441 steps (71m) | ~1 read per folder, no grant writes |

Measured on MHO, 2026-08-19: **71m → 19m50s** with the first-binding reader still in place (868
assignments re-issued, of which the `SEGVIEW` delete/share grants were genuinely new). The remaining
assignments are what §3.4 removes.
| Re-run after adding one unit | ~2,441 steps | that unit's folders and grants only |

The third row is the client's actual request.

## 7. Files

| File | Change |
|---|---|
| `src/shared/grantSkip.ts` | **new, pure** — `needsGrant`, `shouldReadExistingAcl` |
| `src/shared/grantSkip.test.ts` | **new, 10 tests** — including that an unreadable ACL grants |
| `src/webparts/folderManager/components/FolderManager.tsx` | `aclWasReset` per folder; one ACL read; the skip; the per-folder log line |
