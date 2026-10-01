import { formatFileSize } from "./fileSize";

describe("formatFileSize", () => {
  it("renders whole megabytes without a trailing .0", () => {
    // The mockup shows "2MB", not "2.0MB".
    expect(formatFileSize(2097152)).toBe("2MB");
  });

  it("keeps one decimal for part-megabytes", () => {
    expect(formatFileSize(2621440)).toBe("2.5MB");
  });

  it("falls back to KB below a megabyte", () => {
    // A plain megabyte conversion renders these as "0MB", and most uploads here
    // are small PDFs, so that would be the common case rather than the edge one.
    expect(formatFileSize(524288)).toBe("512KB");
    expect(formatFileSize(1024)).toBe("1KB");
  });

  it("falls back to bytes below a kilobyte", () => {
    expect(formatFileSize(512)).toBe("512B");
    expect(formatFileSize(0)).toBe("0B");
  });

  it("uses binary units, matching what Windows and SharePoint report", () => {
    // 1,000,000 bytes is 0.95 MiB — a decimal-unit implementation would call it
    // 1MB and disagree with the size SharePoint shows for the same file.
    expect(formatFileSize(1000000)).toBe("977KB");
  });

  it("returns an empty string for nonsense rather than NaN", () => {
    expect(formatFileSize(NaN)).toBe("");
    expect(formatFileSize(-1)).toBe("");
    expect(formatFileSize(Infinity)).toBe("");
  });
});
