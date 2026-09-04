import {
  friendlyUploadError, blockedBeforeNetwork, BLOCKED_UPLOAD_MESSAGE,
} from "./networkErrors";

describe("friendlyUploadError", () => {
  it("replaces Chrome/Edge/Brave's blocked-request message with the actionable one", () => {
    const msg = friendlyUploadError(new Error("Failed to fetch"), "Upload failed.");
    expect(msg).toContain("Brave Shields");
    expect(msg).toContain("ad blocker");
    expect(msg).not.toContain("Failed to fetch");
  });

  it("matches case-insensitively", () => {
    expect(friendlyUploadError(new Error("FAILED TO FETCH"), "x")).toContain("Brave Shields");
  });

  it("replaces Firefox's wording too", () => {
    const msg = friendlyUploadError(
      new Error("NetworkError when attempting to fetch resource."),
      "x",
    );
    expect(msg).toContain("Brave Shields");
  });

  it("replaces Safari's wording too", () => {
    expect(friendlyUploadError(new Error("Load failed"), "x")).toContain("Brave Shields");
  });

  // The whole point: a real SharePoint error must reach the uploader UNCHANGED, never masked
  // behind a generic "check your extensions" message.
  it("leaves a genuine SharePoint or validation error untouched", () => {
    const msg = friendlyUploadError(new Error("HTTP 403 Forbidden"), "x");
    expect(msg).toBe("HTTP 403 Forbidden");
  });

  it("falls back to the caller's default for a non-Error throw with no message", () => {
    expect(friendlyUploadError(undefined, "Upload failed.")).toBe("Upload failed.");
    expect(friendlyUploadError(null, "Upload failed.")).toBe("Upload failed.");
    expect(friendlyUploadError({}, "Upload failed.")).toBe("Upload failed.");
  });

  it("passes through a thrown plain string", () => {
    expect(friendlyUploadError("something odd happened", "x")).toBe("something odd happened");
  });
});

describe("blockedBeforeNetwork (the XHR path)", () => {
  it("recognises status 0 — a request that never reached the network", () => {
    // Bulk Upload writes through XMLHttpRequest for the progress bar, and a blocked XHR does not
    // throw: it resolves with status 0. That is why `friendlyUploadError` alone could not catch it.
    expect(blockedBeforeNetwork(0)).toBe(true);
  });

  it("leaves every real SharePoint status alone", () => {
    // A genuine refusal always carries a status, so this can never mask an actionable response.
    for (const st of [200, 201, 400, 403, 404, 409, 429, 500, 503]) {
      expect(blockedBeforeNetwork(st)).toBe(false);
    }
  });

  it("shares one message with the fetch path", () => {
    // Two entry points, one wording — the constant exists so it is not copied.
    expect(friendlyUploadError(new Error("Failed to fetch"), "x")).toBe(BLOCKED_UPLOAD_MESSAGE);
  });

  it("names the fix, not just the cause", () => {
    expect(BLOCKED_UPLOAD_MESSAGE).toContain("Shields");
    expect(BLOCKED_UPLOAD_MESSAGE.toLowerCase()).toContain("not a problem with the document");
  });
});
