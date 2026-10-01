import {
  BLOCKED_INPUT_CHARS,
  blockedCharsIn,
  stripBlockedChars,
  blockedCharsMessage,
} from "./inputSanitize";

describe("blocked input characters", () => {
  it("blocks every character SharePoint refuses in a file name", () => {
    for (const ch of '"*:<>?/\|') expect(BLOCKED_INPUT_CHARS.indexOf(ch)).toBeGreaterThan(-1);
  });

  it("blocks the client's own additions", () => {
    for (const ch of "![]@#") expect(BLOCKED_INPUT_CHARS.indexOf(ch)).toBeGreaterThan(-1);
  });

  /* ⚠ THE ONE THAT WOULD BREAK THE CLIENT'S OWN DATA. GHO's department is "Group Legal, Risk ＆
     Compliance" and the term store requires the FULLWIDTH ampersand. A blanket "no special
     characters" reading of the request would refuse it. */
  it("does NOT block an ampersand, plain or fullwidth", () => {
    expect(blockedCharsIn("Risk & Compliance")).toEqual([]);
    expect(blockedCharsIn("Risk \uFF06 Compliance")).toEqual([]);
    expect(stripBlockedChars("Risk \uFF06 Compliance")).toBe("Risk \uFF06 Compliance");
  });

  it("does NOT block punctuation that appears in real values", () => {
    expect(blockedCharsIn("O'Brien")).toEqual([]);
    expect(blockedCharsIn("Q1-2026")).toEqual([]);
    expect(blockedCharsIn("v1.2 final")).toEqual([]);
  });

  it("names each offending character once, in the order typed", () => {
    expect(blockedCharsIn("a/b/c?d")).toEqual(["/", "?"]);
  });

  it("strips them and leaves everything else alone", () => {
    expect(stripBlockedChars("Tax: Q1/2026 #4")).toBe("Tax Q12026 4");
  });

  /* SharePoint refuses a name STARTING with a dot, and the failure surfaces as a rejected upload
     rather than as anything about the field. A dot elsewhere is legal and is kept. */
  it("removes a leading dot but keeps one in the middle", () => {
    expect(stripBlockedChars(".hidden")).toBe("hidden");
    expect(stripBlockedChars("...x")).toBe("x");
    expect(stripBlockedChars("v1.2")).toBe("v1.2");
  });

  /* ⚠ SPACING IS LEFT EXACTLY AS TYPED. Trimming while somebody is mid-sentence eats the space
     before their next word, which reads as the field fighting them. */
  it("does not trim or collapse whitespace", () => {
    expect(stripBlockedChars("Group  Finance ")).toBe("Group  Finance ");
  });

  it("says nothing when nothing was blocked", () => {
    expect(blockedCharsMessage("Group Finance")).toBeUndefined();
  });

  it("names the characters rather than saying 'invalid characters'", () => {
    const msg = blockedCharsMessage("a<b");
    expect(msg).toBeDefined();
    expect((msg ?? "").indexOf("<")).toBeGreaterThan(-1);
  });

  it("survives an empty or missing value", () => {
    expect(blockedCharsIn("")).toEqual([]);
    expect(stripBlockedChars("")).toBe("");
    expect(blockedCharsMessage("")).toBeUndefined();
  });
});
