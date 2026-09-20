import {
  canonicalServiceAccountName,
  displayNameFor,
  nameFromEmail,
} from "./displayName";

describe("nameFromEmail", () => {
  it("title-cases a dot-separated local part", () => {
    expect(nameFromEmail("chiew.wei.chien@sdguthrie.com")).toBe(
      "Chiew Wei Chien",
    );
  });

  it("title-cases an underscore-separated local part", () => {
    expect(nameFromEmail("goh_kheng_wei@sdguthrie.com")).toBe(
      "Goh Kheng Wei",
    );
  });

  it("handles a mix of dots and underscores", () => {
    expect(nameFromEmail("first.last_name@sdguthrie.com")).toBe(
      "First Last Name",
    );
  });

  it("handles a single-word local part", () => {
    expect(nameFromEmail("clarence@trinergydigital.com")).toBe("Clarence");
  });

  it("returns the address itself when there is no local part", () => {
    expect(nameFromEmail("@sdguthrie.com")).toBe("@sdguthrie.com");
  });

  it("returns the address itself when there is no @ at all", () => {
    expect(nameFromEmail("gdc")).toBe("Gdc");
  });
});

describe("displayNameFor", () => {
  it("derives a name for a real address", () => {
    expect(displayNameFor("goh.kheng.wei@sdguthrie.com")).toBe(
      "Goh Kheng Wei",
    );
  });

  it("returns undefined for undefined", () => {
    expect(displayNameFor(undefined)).toBeUndefined();
  });

  it("returns undefined for an empty string", () => {
    expect(displayNameFor("")).toBeUndefined();
  });

  it("returns undefined for a whitespace-only string", () => {
    expect(displayNameFor("   ")).toBeUndefined();
  });

  it("canonicalises the current proxy account over guessing from its address", () => {
    // A real live shape: the flow stored the bare alias, not the full address.
    expect(displayNameFor("gdc")).toBe("Guthrie Document Centre");
    expect(displayNameFor("gdc@sdguthrie.com")).toBe(
      "Guthrie Document Centre",
    );
  });

  it("canonicalises the retired proxy account too, for historical rows", () => {
    expect(displayNameFor("crs@sdguthrie.com")).toBe(
      "Guthrie Central Repository System",
    );
  });
});

describe("canonicalServiceAccountName", () => {
  it("recognises the bare alias, the full address, and mixed case, identically", () => {
    // Client screenshot, 2026-09-20: the SAME account rendered three different ways across
    // adjacent Audit Log rows because three different flow actions stored it three different ways.
    expect(canonicalServiceAccountName("gdc")).toBe(
      "Guthrie Document Centre",
    );
    expect(canonicalServiceAccountName("gdc@sdguthrie.com")).toBe(
      "Guthrie Document Centre",
    );
    expect(canonicalServiceAccountName("GDC@SDGuthrie.com")).toBe(
      "Guthrie Document Centre",
    );
  });

  it("recognises the retired crs proxy account", () => {
    expect(canonicalServiceAccountName("crs@sdguthrie.com")).toBe(
      "Guthrie Central Repository System",
    );
  });

  it("returns undefined for a real person's address", () => {
    expect(
      canonicalServiceAccountName("goh.kheng.wei@sdguthrie.com"),
    ).toBeUndefined();
  });

  it("returns undefined for an empty string", () => {
    expect(canonicalServiceAccountName("")).toBeUndefined();
  });
});
