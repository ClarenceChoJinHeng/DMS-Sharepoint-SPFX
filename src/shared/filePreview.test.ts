import { previewKind, previewTarget, extensionOf, encodePath } from "./filePreview";

const ROOT = "https://example.sharepoint.com";
const WEB = "https://example.sharepoint.com/sites/Example";
const PATH = "/sites/Example/Staging/GHO/GF/CORU/2026/Invoice/file";

describe("extensionOf", () => {
  it("reads the last extension, lowercased", () => {
    expect(extensionOf("Invoice.PDF")).toBe("pdf");
    expect(extensionOf("report.final.docx")).toBe("docx");
  });

  it("returns empty for names with no usable extension", () => {
    expect(extensionOf("README")).toBe("");
    expect(extensionOf("trailing.")).toBe("");
    // A dotfile is a name, not an extension — ".pdf" alone must not read as a PDF.
    expect(extensionOf(".pdf")).toBe("");
    expect(extensionOf("")).toBe("");
    expect(extensionOf(undefined as unknown as string)).toBe("");
  });
});

describe("previewKind", () => {
  it("keeps PDF on the native browser viewer", () => {
    expect(previewKind("invoice.pdf")).toBe("pdf");
  });

  // The regression this file exists for: only PDF rendered, so an approver was asked to decide
  // on documents they could not see.
  it("routes images to the image strategy", () => {
    for (const n of ["scan.png", "photo.JPG", "a.jpeg", "b.gif", "c.bmp", "d.webp", "e.svg"]) {
      expect(previewKind(n)).toBe("image");
    }
  });

  it("routes Office documents to Office Online", () => {
    for (const n of ["a.doc", "b.docx", "c.xls", "d.xlsx", "e.xlsm", "f.csv", "g.ppt", "h.pptx"]) {
      expect(previewKind(n)).toBe("office");
    }
  });

  it("covers every type the client actually allows for upload", () => {
    // CLAUDE.md: allowed types are .pdf .doc .docx .xls .xlsx as of 2026-07-30. None of these may
    // fall through to "none" — that would be a document nobody can review.
    for (const n of ["a.pdf", "b.doc", "c.docx", "d.xls", "e.xlsx"]) {
      expect(previewKind(n)).not.toBe("none");
    }
  });

  it("shows plain text inline", () => {
    expect(previewKind("notes.txt")).toBe("text");
    expect(previewKind("data.json")).toBe("text");
  });

  // Refusing beats guessing: an unknown type rendered in an iframe is the blank pane this module
  // was written to remove.
  it("refuses anything it does not know", () => {
    for (const n of ["archive.zip", "installer.exe", "clip.mp4", "font.woff2", "README"]) {
      expect(previewKind(n)).toBe("none");
    }
  });
});

describe("encodePath", () => {
  it("encodes segments but keeps the separators", () => {
    expect(encodePath("/sites/Ex/Staging/Group Finance/f.pdf"))
      .toBe("/sites/Ex/Staging/Group%20Finance/f.pdf");
  });

  it("encodes the fullwidth ampersand SharePoint requires in term labels", () => {
    // "Group Legal, Risk ＆ Compliance" is a real folder name on this tenant.
    expect(encodePath("/a/Risk ＆ Co/f.pdf")).toBe("/a/Risk%20%EF%BC%86%20Co/f.pdf");
  });
});

describe("previewTarget", () => {
  it("appends the PDF fit parameter", () => {
    const t = previewTarget("f.pdf", `${PATH}.pdf`, ROOT, WEB);
    expect(t.kind).toBe("pdf");
    expect(t.url).toBe(`${ROOT}${PATH}.pdf#view=FitH`);
  });

  it("builds a WopiFrame embed URL for Office, encoding sourcedoc exactly once", () => {
    const t = previewTarget("f.docx", "/sites/Ex/Staging/Group Finance/f.docx", ROOT, WEB);
    expect(t.kind).toBe("office");
    expect(t.url).toBe(
      `${WEB}/_layouts/15/WopiFrame.aspx?sourcedoc=%2Fsites%2FEx%2FStaging%2FGroup%20Finance%2Ff.docx&action=embedview`,
    );
    // Double-encoding here yields a "file not found" page that reads as a permissions problem.
    expect(t.url).not.toContain("%252F");
  });

  it("gives images the plain file URL, with no PDF parameter", () => {
    const t = previewTarget("scan.png", `${PATH}.png`, ROOT, WEB);
    expect(t.kind).toBe("image");
    expect(t.url).toBe(`${ROOT}${PATH}.png`);
    expect(t.url).not.toContain("#view");
  });

  it("leaves url empty for an unpreviewable type but still offers the file", () => {
    const t = previewTarget("archive.zip", `${PATH}.zip`, ROOT, WEB);
    expect(t.kind).toBe("none");
    expect(t.url).toBe("");
    expect(t.fileUrl).toBe(`${ROOT}${PATH}.zip`);
  });

  it("always populates fileUrl, whatever the kind", () => {
    for (const n of ["a.pdf", "b.docx", "c.png", "d.txt", "e.zip"]) {
      expect(previewTarget(n, `/sites/Ex/${n}`, ROOT, WEB).fileUrl).toBe(`${ROOT}/sites/Ex/${n}`);
    }
  });
});
