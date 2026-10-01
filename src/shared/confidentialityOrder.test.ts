import {
  confidentialityRank,
  sortByConfidentialityOrder,
} from "./confidentialityOrder";

describe("confidentialityRank", () => {
  it("ranks Highly Confidential first, Confidential second, Restricted third", () => {
    expect(confidentialityRank("Highly Confidential")).toBeLessThan(
      confidentialityRank("Confidential"),
    );
    expect(confidentialityRank("Confidential")).toBeLessThan(
      confidentialityRank("Restricted"),
    );
  });

  it("is case- and whitespace-insensitive", () => {
    expect(confidentialityRank("  highly confidential  ")).toBe(
      confidentialityRank("Highly Confidential"),
    );
  });

  it("ranks an unrecognised label past all three known ones", () => {
    expect(confidentialityRank("Public")).toBeGreaterThan(
      confidentialityRank("Restricted"),
    );
  });
});

describe("sortByConfidentialityOrder", () => {
  it("reorders the term store's alphabetical result into the client's requested order", () => {
    const alphabetical = [
      { label: "Confidential" },
      { label: "Highly Confidential" },
      { label: "Restricted" },
    ];
    expect(sortByConfidentialityOrder(alphabetical).map((o) => o.label)).toEqual([
      "Highly Confidential",
      "Confidential",
      "Restricted",
    ]);
  });

  it("does not mutate the input array", () => {
    const input = [{ label: "Confidential" }, { label: "Highly Confidential" }];
    const copy = input.slice();
    sortByConfidentialityOrder(input);
    expect(input).toEqual(copy);
  });

  it("keeps unrecognised labels in their original relative order, at the end", () => {
    const withExtra = [
      { label: "Confidential" },
      { label: "Public" },
      { label: "Highly Confidential" },
      { label: "Internal" },
    ];
    expect(sortByConfidentialityOrder(withExtra).map((o) => o.label)).toEqual([
      "Highly Confidential",
      "Confidential",
      "Public",
      "Internal",
    ]);
  });

  it("preserves the object identity of each item, only reordering them", () => {
    const a = { label: "Confidential", id: "1" };
    const b = { label: "Highly Confidential", id: "2" };
    const sorted = sortByConfidentialityOrder([a, b]);
    expect(sorted[0]).toBe(b);
    expect(sorted[1]).toBe(a);
  });
});
