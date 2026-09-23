import {
  canonicalServiceAccountName,
  displayNameFor,
  nameFromEmail,
  resolveActorDisplay,
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

  it("canonicalises the retired proxy account under the CURRENT name, not its own old one", () => {
    // Changed 2026-09-23, client: "crs is approving but the move to documents is showing crs
    // instead of gdc" — crs and gdc are one identity across time, from the client's point of view.
    expect(displayNameFor("crs@sdguthrie.com")).toBe("Guthrie Document Centre");
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

  it("recognises the retired crs proxy account, under the CURRENT name", () => {
    expect(canonicalServiceAccountName("crs@sdguthrie.com")).toBe(
      "Guthrie Document Centre",
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

describe("resolveActorDisplay", () => {
  it("cleans up a FULL raw email stored as ActorName — CRS — Audit request activity's shape", () => {
    expect(
      resolveActorDisplay(
        "clarence@trinergydigital.com",
        "clarence@trinergydigital.com",
      ),
    ).toBe("Clarence");
  });

  it("cleans up a bare LOCAL PART stored as ActorName — Execute approved deletion's shape", () => {
    expect(
      resolveActorDisplay("goh.kheng.wei@sdguthrie.com", "goh.kheng.wei"),
    ).toBe("Goh Kheng Wei");
  });

  it("leaves a genuine display name essentially unchanged — Auto-route's shape", () => {
    expect(
      resolveActorDisplay("clarence@trinergydigital.com", "Clarence Cho"),
    ).toBe("Clarence Cho");
  });

  it("recognises a known service account over guessing, on EITHER field", () => {
    expect(resolveActorDisplay("gdc@sdguthrie.com", "gdc@sdguthrie.com")).toBe(
      "Guthrie Document Centre",
    );
    expect(resolveActorDisplay(undefined, "gdc")).toBe(
      "Guthrie Document Centre",
    );
  });

  it("falls back to ActorEmail when ActorName is blank", () => {
    expect(resolveActorDisplay("goh.kheng.wei@sdguthrie.com", undefined)).toBe(
      "Goh Kheng Wei",
    );
    expect(resolveActorDisplay("goh.kheng.wei@sdguthrie.com", "")).toBe(
      "Goh Kheng Wei",
    );
  });

  it("returns an em dash when neither field has anything", () => {
    expect(resolveActorDisplay(undefined, undefined)).toBe("—");
    expect(resolveActorDisplay("", "")).toBe("—");
  });
});
