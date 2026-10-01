import { shouldAttemptSelfApprove } from "./selfApprove";

describe("shouldAttemptSelfApprove", () => {
  it("is on only for the literal value 'yes', case-insensitively", () => {
    expect(shouldAttemptSelfApprove("yes")).toBe(true);
    expect(shouldAttemptSelfApprove("YES")).toBe(true);
    expect(shouldAttemptSelfApprove("Yes")).toBe(true);
    expect(shouldAttemptSelfApprove(" yes ")).toBe(true);
  });

  // Fails CLOSED: absent, blank or unrecognised all mean "do not self-approve" — the opposite
  // default from uploadsPaused/legallyPrivilegedFor, because the cost here (a document going live
  // with nobody having looked at it) is worse than the cost everywhere else (a form out of service).
  it("is off for anything else, including absent or unreadable config", () => {
    expect(shouldAttemptSelfApprove(undefined)).toBe(false);
    expect(shouldAttemptSelfApprove("")).toBe(false);
    expect(shouldAttemptSelfApprove("no")).toBe(false);
    expect(shouldAttemptSelfApprove("true")).toBe(false);
    expect(shouldAttemptSelfApprove("1")).toBe(false);
    expect(shouldAttemptSelfApprove("yess")).toBe(false);
    expect(shouldAttemptSelfApprove("  ")).toBe(false);
  });
});
