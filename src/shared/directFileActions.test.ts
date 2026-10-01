import { directActionsFor } from "./directFileActions";
import { FileRights } from "./dmsFolderMap";

const rights = (remove: FileRights["remove"], share: FileRights["share"]): FileRights => ({
  remove,
  share,
});

describe("directActionsFor", () => {
  it("offers both when the probe grants both and the file is not archived", () => {
    expect(directActionsFor(rights("granted", "granted"), false)).toEqual({
      canDelete: true,
      canShare: true,
    });
  });

  it("offers share only when the probe denies delete but grants share (the ordinary Approver)", () => {
    expect(directActionsFor(rights("denied", "granted"), false)).toEqual({
      canDelete: false,
      canShare: true,
    });
  });

  it("offers neither when the probe denies both (a plain uploader)", () => {
    expect(directActionsFor(rights("denied", "denied"), false)).toEqual({
      canDelete: false,
      canShare: false,
    });
  });

  it("offers neither on 'missing' or 'unknown', same as 'denied' — never assume a right exists", () => {
    expect(directActionsFor(rights("missing", "unknown"), false)).toEqual({
      canDelete: false,
      canShare: false,
    });
  });

  it("ARCHIVE OVERRIDES EVERYTHING — both suppressed even when the probe grants both", () => {
    // This is the whole reason the function exists rather than a one-line inline check:
    // probeFileRights only answers permission, not appropriateness, and an archived document
    // is read-only to everybody by design regardless of what Full Control would otherwise allow.
    expect(directActionsFor(rights("granted", "granted"), true)).toEqual({
      canDelete: false,
      canShare: false,
    });
  });
});
