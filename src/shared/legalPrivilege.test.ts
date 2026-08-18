import { offersLegalPrivilege, privilegedLevels } from "./legalPrivilege";

describe("privilegedLevels", () => {
  it("reads a single level exactly as before", () => {
    expect(privilegedLevels("Confidential")).toEqual(["Confidential"]);
  });

  it("reads a list, separated by either character an admin might type", () => {
    expect(privilegedLevels("Confidential;Highly Confidential"))
      .toEqual(["Confidential", "Highly Confidential"]);
    expect(privilegedLevels("Confidential, Highly Confidential"))
      .toEqual(["Confidential", "Highly Confidential"]);
  });

  it("yields nothing for blank, which is what keeps the tick hidden by default", () => {
    expect(privilegedLevels("")).toEqual([]);
    expect(privilegedLevels(undefined)).toEqual([]);
    expect(privilegedLevels("  ;  , ")).toEqual([]);
  });

  it("de-duplicates, so a doubled entry cannot appear twice anywhere downstream", () => {
    expect(privilegedLevels("Confidential;Confidential")).toEqual(["Confidential"]);
  });
});

describe("offersLegalPrivilege", () => {
  it("offers the tick for Highly Confidential when it is listed", () => {
    // THE CLIENT'S ASK, 2026-08-19. Before this, naming HC would have removed the tick from
    // Confidential, because the setting could only ever express one level.
    expect(offersLegalPrivilege("Highly Confidential", "Confidential;Highly Confidential")).toBe(true);
    expect(offersLegalPrivilege("Confidential", "Confidential;Highly Confidential")).toBe(true);
  });

  it("does not offer it for a level that is not listed", () => {
    expect(offersLegalPrivilege("Restricted", "Confidential;Highly Confidential")).toBe(false);
  });

  it("tolerates the whitespace and casing a hand-typed config cell carries", () => {
    // The old `===` would have offered the tick NOWHERE for any of these, while the config row
    // looked correct — and nothing reports a setting that matches nothing.
    expect(offersLegalPrivilege("Highly Confidential", " Highly Confidential ")).toBe(true);
    expect(offersLegalPrivilege("highly confidential", "Highly Confidential")).toBe(true);
    expect(offersLegalPrivilege(" Confidential", "Confidential")).toBe(true);
  });

  it("never offers it when nothing is configured", () => {
    expect(offersLegalPrivilege("Highly Confidential", "")).toBe(false);
    expect(offersLegalPrivilege("Highly Confidential", undefined)).toBe(false);
  });

  it("never offers it for a document with no confidentiality set", () => {
    // A blank level must not match a blank list entry — that would stamp a legal marker on a
    // document whose level nobody chose.
    expect(offersLegalPrivilege("", "Confidential;")).toBe(false);
    expect(offersLegalPrivilege(undefined, "Confidential")).toBe(false);
  });
});
