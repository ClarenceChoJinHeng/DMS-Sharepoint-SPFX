import {
  parseSiteMapRows,
  siteOrigin,
  libraryRootPath,
  SiteMapRow,
} from "./siteMap";

describe("parseSiteMapRows", () => {
  it("keeps valid rows and defaults a blank library to Documents", () => {
    const rows: SiteMapRow[] = [
      { title: "Upstream", siteUrl: "https://t.sharepoint.com/sites/Up", documentsLibrary: "" },
    ];
    expect(parseSiteMapRows(rows)).toEqual([
      { segment: "Upstream", siteUrl: "https://t.sharepoint.com/sites/Up", documentsLibrary: "Documents" },
    ]);
  });

  it("strips a trailing slash from the site URL", () => {
    const rows: SiteMapRow[] = [
      { title: "Up", siteUrl: "https://t.sharepoint.com/sites/Up/", documentsLibrary: "Documents" },
    ];
    expect(parseSiteMapRows(rows)[0].siteUrl).toBe("https://t.sharepoint.com/sites/Up");
  });

  it("skips rows with no site URL", () => {
    const rows: SiteMapRow[] = [
      { title: "Bad", siteUrl: "", documentsLibrary: "Documents" },
      { title: "Good", siteUrl: "https://t.sharepoint.com/sites/Good", documentsLibrary: "Docs" },
    ];
    const out = parseSiteMapRows(rows);
    expect(out).toHaveLength(1);
    expect(out[0].segment).toBe("Good");
    expect(out[0].documentsLibrary).toBe("Docs");
  });

  it("falls back to the site URL as the label when Title is blank", () => {
    const rows: SiteMapRow[] = [
      { title: "", siteUrl: "https://t.sharepoint.com/sites/X", documentsLibrary: "Documents" },
    ];
    expect(parseSiteMapRows(rows)[0].segment).toBe("https://t.sharepoint.com/sites/X");
  });

  it("returns [] for empty/nullish input", () => {
    expect(parseSiteMapRows([])).toEqual([]);
    expect(parseSiteMapRows(undefined as unknown as SiteMapRow[])).toEqual([]);
  });
});

describe("siteOrigin", () => {
  it("returns the origin of a valid URL", () => {
    expect(siteOrigin("https://t.sharepoint.com/sites/Up")).toBe("https://t.sharepoint.com");
  });
  it("returns '' for an invalid URL", () => {
    expect(siteOrigin("not a url")).toBe("");
  });
});

describe("libraryRootPath", () => {
  it("joins the site path and library name", () => {
    expect(libraryRootPath("https://t.sharepoint.com/sites/Up", "Documents")).toBe(
      "/sites/Up/Documents",
    );
  });
  it("defaults a blank library to Documents", () => {
    expect(libraryRootPath("https://t.sharepoint.com/sites/Up", "")).toBe("/sites/Up/Documents");
  });
  it("strips a trailing slash on the site path before joining", () => {
    expect(libraryRootPath("https://t.sharepoint.com/sites/Up/", "Staging")).toBe(
      "/sites/Up/Staging",
    );
  });
  it("returns '' for an invalid URL", () => {
    expect(libraryRootPath("nope", "Documents")).toBe("");
  });
});
