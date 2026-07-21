import {
  roleNameForAccess,
  parseShareTarget,
  buildRequestPayload,
  itemAbsoluteUrl,
  itemBreadcrumb,
  ShareTarget,
} from "./shareGuard";

describe("roleNameForAccess", () => {
  it("maps Read and Edit to SharePoint role names", () => {
    expect(roleNameForAccess("Read")).toBe("Read");
    expect(roleNameForAccess("Edit")).toBe("Edit");
  });
});

describe("parseShareTarget", () => {
  it("reads item context from URLSearchParams", () => {
    const q = new URLSearchParams(
      "itemUrl=%2Fsites%2FX%2FStaging%2FGHO%2Ff.pdf&itemId=abc-123&itemType=File&library=Staging",
    );
    const t = parseShareTarget(q);
    expect(t).toEqual<ShareTarget>({
      itemUrl: "/sites/X/Staging/GHO/f.pdf",
      itemUniqueId: "abc-123",
      itemType: "File",
      library: "Staging",
    });
  });

  it("returns null when required params are missing", () => {
    expect(parseShareTarget(new URLSearchParams("itemUrl=x"))).toBeNull();
  });

  it("defaults itemType to File and library to Documents when unrecognised", () => {
    const t = parseShareTarget(
      new URLSearchParams("itemUrl=%2Fa&itemId=z&itemType=Bogus&library=Bogus"),
    );
    expect(t?.itemType).toBe("File");
    expect(t?.library).toBe("Documents");
  });
});

describe("buildRequestPayload", () => {
  it("builds the list item body with Pending status", () => {
    const target: ShareTarget = {
      itemUrl: "/sites/X/Staging/GHO/f.pdf",
      itemUniqueId: "abc-123",
      itemType: "File",
      library: "Staging",
    };
    const body = buildRequestPayload(target, 42, "Audit review", "Read");
    expect(body).toEqual({
      Title: "f.pdf",
      ItemUrl: "/sites/X/Staging/GHO/f.pdf",
      ItemUniqueId: "abc-123",
      ItemType: "File",
      Library: "Staging",
      RecipientId: 42,
      AccessLevel: "Read",
      Reason: "Audit review",
      Status: "Pending",
    });
  });

  it("derives Title from a folder path without a trailing slash", () => {
    const target: ShareTarget = {
      itemUrl: "/sites/X/Documents/GHO/GCO",
      itemUniqueId: "u",
      itemType: "Folder",
      library: "Documents",
    };
    expect(buildRequestPayload(target, 1, "", "Edit").Title).toBe("GCO");
  });
});

describe("itemAbsoluteUrl", () => {
  it("joins origin and encoded server-relative url", () => {
    expect(itemAbsoluteUrl("https://t.sharepoint.com", "/sites/X/Docs/a b.pdf"))
      .toBe("https://t.sharepoint.com/sites/X/Docs/a%20b.pdf");
  });
});

describe("itemBreadcrumb", () => {
  it("returns the path segments below the web root", () => {
    expect(itemBreadcrumb("/sites/X/Staging/Group Head Office/Treasury", "/sites/X"))
      .toEqual(["Staging", "Group Head Office", "Treasury"]);
  });

  it("tolerates a trailing slash on the web url and the item url", () => {
    expect(itemBreadcrumb("/sites/X/Documents/GHO/GCO/", "/sites/X/"))
      .toEqual(["Documents", "GHO", "GCO"]);
  });

  it("decodes percent-encoded segments", () => {
    expect(itemBreadcrumb("/sites/X/Staging/a%20b/c", "/sites/X"))
      .toEqual(["Staging", "a b", "c"]);
  });

  it("matches the web root case-insensitively", () => {
    expect(itemBreadcrumb("/Sites/X/Staging/A", "/sites/x"))
      .toEqual(["Staging", "A"]);
  });

  it("returns all segments when the web root is not a prefix", () => {
    expect(itemBreadcrumb("/other/Staging/A", "/sites/X"))
      .toEqual(["other", "Staging", "A"]);
  });
});
