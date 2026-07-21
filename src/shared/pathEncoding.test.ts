import { encodeServerRelativePath } from "./pathEncoding";

describe("encodeServerRelativePath", () => {
  it("keeps the slashes literal and encodes spaces per segment", () => {
    expect(encodeServerRelativePath("/sites/X/Staging/Group Head Office/2024"))
      .toBe("/sites/X/Staging/Group%20Head%20Office/2024");
  });

  it("preserves a leading slash", () => {
    expect(encodeServerRelativePath("/a/b")).toBe("/a/b");
  });

  it("encodes a real ampersand and hash within a segment (URL-safe)", () => {
    expect(encodeServerRelativePath("/x/A & B #1/y"))
      .toBe("/x/A%20%26%20B%20%231/y");
  });

  it("encodes a fullwidth ampersand (unicode) within a segment", () => {
    expect(encodeServerRelativePath("/x/Compliance ＆ Operational Risk (CORU)"))
      .toBe("/x/Compliance%20%EF%BC%86%20Operational%20Risk%20(CORU)");
  });

  it("does NOT flood the path with %2F (deep path stays readable)", () => {
    const deep = "/sites/X/Staging/GHO/Group Finance/Unit A/2026/Agreement";
    const out = encodeServerRelativePath(deep);
    expect(out.indexOf("%2F")).toBe(-1);
    expect(out.split("/").length).toBe(deep.split("/").length);
  });
});
