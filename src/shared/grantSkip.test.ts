import { needsGrant, shouldReadExistingAcl, AssignLite } from "./grantSkip";

const a = (principalId: number, roleDefId: number): AssignLite => ({ principalId, roleDefId });

describe("needsGrant", () => {
  it("skips a grant that is already exactly in place", () => {
    expect(needsGrant([a(12, 1073741826)], 12, 1073741826)).toBe(false);
  });

  it("grants when the folder holds nothing", () => {
    expect(needsGrant([], 12, 1073741826)).toBe(true);
  });

  it("grants when a DIFFERENT principal holds the same level", () => {
    expect(needsGrant([a(99, 1073741826)], 12, 1073741826)).toBe(true);
  });

  /**
   * The Head-of-Department case: the same group can hold its own level AND the ancestor-browse
   * Read on one department folder, and the reader keeps only the first binding per principal.
   * Re-issuing a no-op write is the cheap error; skipping a real permission is not.
   */
  it("grants when the principal is present holding a DIFFERENT level", () => {
    expect(needsGrant([a(12, 1073741827)], 12, 1073741826)).toBe(true);
  });

  it("finds the match among several assignments", () => {
    expect(needsGrant([a(1, 5), a(12, 1073741826), a(3, 9)], 12, 1073741826)).toBe(false);
  });

  /**
   * ⚠ UNKNOWN GRANTS. A read that failed says nothing about the folder, and a wrongly skipped
   * grant is an uploader who cannot upload on a run that reported success.
   */
  it("GRANTS when the assignments could not be read", () => {
    expect(needsGrant(undefined, 12, 1073741826)).toBe(true);
    expect(needsGrant(null, 12, 1073741826)).toBe(true);
  });

  it("survives a malformed entry rather than throwing mid-run", () => {
    expect(needsGrant([undefined as unknown as AssignLite, a(12, 7)], 12, 7)).toBe(false);
    expect(needsGrant([undefined as unknown as AssignLite], 12, 7)).toBe(true);
  });
});

describe("shouldReadExistingAcl", () => {
  /**
   * Reconciliation just reset the ACL itself — created the folder, or broke inheritance with
   * copyRoleAssignments=false. Only Owners remain, so a read costs a request per folder on the
   * FIRST run and can never save a write.
   */
  it("does not read when reconciliation just reset the ACL", () => {
    expect(shouldReadExistingAcl(true, 5)).toBe(false);
  });

  /** The steady state of a re-run: the folder was already locked and left alone. */
  it("reads when the folder was left as it was found", () => {
    expect(shouldReadExistingAcl(false, 5)).toBe(true);
  });

  it("does not read when there is nothing to grant anyway", () => {
    expect(shouldReadExistingAcl(false, 0)).toBe(false);
    expect(shouldReadExistingAcl(true, 0)).toBe(false);
  });
});
