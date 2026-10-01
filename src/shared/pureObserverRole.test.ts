import { isPureObserverRole } from "./pureObserverRole";

describe("isPureObserverRole", () => {
  it("is true for a plain Viewer (MEMBER only)", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "MEMBER" }]),
    ).toBe(true);
  });

  it("is true for a C-Level (GLOBAL) account", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "GLOBAL" }]),
    ).toBe(true);
  });

  it("is true for a segment C-Level (SEGVIEW) account", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "SEGVIEW" }]),
    ).toBe(true);
  });

  it("is true when the viewer holds no Group Map rows at all", () => {
    expect(isPureObserverRole(["10"], [])).toBe(true);
  });

  it("is false for a PIC (UPL)", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "UPL" }]),
    ).toBe(false);
  });

  it("is false for an Approver (APR)", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "APR" }]),
    ).toBe(false);
  });

  it("is false for a Head of Department — DEPTVIEW never travels alone under the current persona model, and is itself in the acting set", () => {
    expect(
      isPureObserverRole(
        ["10"],
        [
          { GroupId: "10", Role: "DEPTVIEW" },
          { GroupId: "10", Role: "SHARE" },
        ],
      ),
    ).toBe(false);
    // Even a bare DEPTVIEW row on its own must not read as a pure observer — see the module's own
    // comment on why DEPTVIEW is in ACTING_ROLES despite the name.
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "DEPTVIEW" }]),
    ).toBe(false);
  });

  it("ignores a row belonging to a group the viewer is NOT in", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "99", Role: "APR" }]),
    ).toBe(true);
  });

  it("is false the moment ANY acting role is found, even mixed with view-only roles", () => {
    expect(
      isPureObserverRole(
        ["10"],
        [
          { GroupId: "10", Role: "MEMBER" },
          { GroupId: "10", Role: "DELS" },
        ],
      ),
    ).toBe(false);
  });

  it("normalises long-form role spellings, same as every other Group Map reader in this project", () => {
    expect(
      isPureObserverRole(["10"], [{ GroupId: "10", Role: "UPLOADER" }]),
    ).toBe(false);
  });
});
