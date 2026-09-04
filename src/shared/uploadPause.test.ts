import { uploadsArePaused, pauseSettingValue, UPLOAD_PAUSE_SETTING, UPLOAD_PAUSE_MESSAGE } from "./uploadPause";

describe("uploadsArePaused", () => {
  it("pauses on the affirmative spellings an admin might type", () => {
    for (const v of ["yes", "YES", " Yes ", "true", "on", "1", "paused"]) {
      expect(uploadsArePaused(v)).toBe(true);
    }
  });

  it("does not pause on the negative spellings", () => {
    for (const v of ["no", "NO", "false", "off", "0"]) {
      expect(uploadsArePaused(v)).toBe(false);
    }
  });

  /**
   * ⚠ FAILS OPEN. A wrong `true` blocks every uploader on the site over a transient read; a wrong
   * `false` costs one file in the old shape, which the migration's fresh-scan guard catches. Same
   * reasoning as gotcha 10b — a failed read proves nothing.
   */
  it("does NOT pause when the setting is missing, blank or unreadable", () => {
    expect(uploadsArePaused(undefined)).toBe(false);
    expect(uploadsArePaused(null)).toBe(false);
    expect(uploadsArePaused("")).toBe(false);
    expect(uploadsArePaused("   ")).toBe(false);
  });

  /** An unrecognised value is not a licence to block the whole site. */
  it("does not pause on anything unrecognised", () => {
    for (const v of ["maybe", "y", "pause", "ON HOLD", "yes please"]) {
      expect(uploadsArePaused(v)).toBe(false);
    }
  });
});

describe("pauseSettingValue", () => {
  it("writes words a human can read in a list view", () => {
    expect(pauseSettingValue(true)).toBe("yes");
    expect(pauseSettingValue(false)).toBe("no");
  });

  /** Turning it off must WRITE "no", never blank the cell — blank cannot be told from never-set. */
  it("round-trips through the reader", () => {
    expect(uploadsArePaused(pauseSettingValue(true))).toBe(true);
    expect(uploadsArePaused(pauseSettingValue(false))).toBe(false);
  });
});

describe("the constants", () => {
  it("names the config row", () => {
    expect(UPLOAD_PAUSE_SETTING).toBe("uploadsPaused");
  });

  /** The uploader must be told nothing is wrong with THEIR file, or they retry it repeatedly. */
  it("explains itself to an uploader who did nothing wrong", () => {
    expect(UPLOAD_PAUSE_MESSAGE).toContain("paused");
    expect(UPLOAD_PAUSE_MESSAGE.toLowerCase()).toContain("nothing is wrong");
  });
});
