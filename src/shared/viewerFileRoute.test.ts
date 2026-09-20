import { classifyViewerForFileRoute } from "./viewerFileRoute";

describe("classifyViewerForFileRoute", () => {
  it("returns 'approver' when the viewer holds an APR row", () => {
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "10", Role: "APR" }]),
    ).toBe("approver");
  });

  it("returns 'approver' for APRHC too", () => {
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "10", Role: "APRHC" }]),
    ).toBe("approver");
  });

  it("returns 'uploader' when the viewer holds only UPL/UPLHC", () => {
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "10", Role: "UPL" }]),
    ).toBe("uploader");
  });

  it("returns 'other' when the viewer holds neither role", () => {
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "10", Role: "DEPTVIEW" }]),
    ).toBe("other");
  });

  it("returns 'other' when the viewer holds no rows at all", () => {
    expect(classifyViewerForFileRoute(["10"], [])).toBe("other");
  });

  it("ignores a row belonging to a group the viewer is NOT in", () => {
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "99", Role: "APR" }]),
    ).toBe("other");
  });

  it("prefers approver over uploader when the viewer holds both — a Head of Unit's own persona carries UPL too", () => {
    expect(
      classifyViewerForFileRoute(
        ["10"],
        [
          { GroupId: "10", Role: "UPL" },
          { GroupId: "10", Role: "APR" },
        ],
      ),
    ).toBe("approver");
    // Order in the input must not matter — APR appearing FIRST or LAST both resolve the same way.
    expect(
      classifyViewerForFileRoute(
        ["10"],
        [
          { GroupId: "10", Role: "APR" },
          { GroupId: "10", Role: "UPL" },
        ],
      ),
    ).toBe("approver");
  });

  it("normalises long-form role spellings, same as every other Group Map reader in this project", () => {
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "10", Role: "APPROVER" }]),
    ).toBe("approver");
    expect(
      classifyViewerForFileRoute(["10"], [{ GroupId: "10", Role: "UPLOADER" }]),
    ).toBe("uploader");
  });
});
