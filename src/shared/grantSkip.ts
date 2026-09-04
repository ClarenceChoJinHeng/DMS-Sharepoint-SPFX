// Should reconciliation re-issue a folder role assignment it may already have made?
//
// Spec: docs/superpowers/specs/2026-08-19-skip-existing-grants-design.md
//
// ⚠ THE COST THIS EXISTS TO REMOVE IS REAL AND MEASURED. Until 1.0.177.0 the folder pass
// called `addRoleAssignment` UNCONDITIONALLY for every group on every folder in every
// library on every run. SharePoint accepts a duplicate as a no-op, so nothing ever failed —
// it simply cost a round trip and a throttle tick each time. On the rehearsal site that is
// ~980 pointless writes for ONE segment, which is why adding a single unit still cost the
// client the best part of an hour and produced register #18.
//
// Note the asymmetry that made this survive so long: the ancestor-browse grant twenty lines
// below it in the same pass ALREADY skipped what was in place. One loop checked, its
// neighbour did not, and the only symptom was time.
//
// ⚠ UNKNOWN MUST GRANT. A failed or unreadable ACL read means we do not know what the folder
// holds — and the cost of guessing "already granted" is a group that silently never receives
// its permission, which presents as an uploader who cannot upload and an admin who has
// already run reconciliation and been told everything is fine. Skipping is an optimisation;
// granting is the job. Same rule as everywhere else here: empty ≠ unknown.

/** The fields this module needs from a folder's existing role assignment. */
export interface AssignLite {
  principalId: number;
  roleDefId: number;
}

/**
 * Does this (principal, role definition) pair still need granting?
 *
 * `existing` is the folder's current assignments, or `undefined` when they could not be
 * read — which always answers `true`.
 *
 * The match is EXACT on both halves, deliberately. A group can hold more than one binding on
 * one folder (a Head of Department holds `DEPTVIEW` on a department folder and can also hold
 * the ancestor-browse `Read` there), and the reader this feeds keeps only the first binding
 * it finds per principal. So a principal seen holding a DIFFERENT level is treated as not yet
 * granted and the grant is re-issued — a wasted no-op write in a rare case, versus a missing
 * permission if we matched on the principal alone.
 */
export function needsGrant(
  existing: AssignLite[] | null | undefined,
  principalId: number,
  roleDefId: number,
): boolean {
  if (!existing) return true;
  return !existing.some((a) => a && a.principalId === principalId && a.roleDefId === roleDefId);
}

/**
 * Is reading the folder's current assignments worth a request?
 *
 * No, in the two cases where reconciliation has just RESET the ACL itself: a folder it created,
 * and a folder whose inheritance it just broke with `copyRoleAssignments=false`. Both leave
 * only site Owners, so every planned grant is genuinely missing and a read could only ever
 * return nothing useful. Reading there would ADD a request per folder to the first run — the
 * expensive one — to save nothing.
 *
 * Yes, on a folder that was already locked and left alone, which is the whole steady state of
 * a re-run: one read replaces up to five writes per folder per library.
 */
export function shouldReadExistingAcl(aclWasReset: boolean, plannedGrants: number): boolean {
  return !aclWasReset && plannedGrants > 0;
}
