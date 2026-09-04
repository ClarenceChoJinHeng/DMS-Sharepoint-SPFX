// How to preview a file, by extension.
//
// The approval page used one strategy for everything: the file's own URL in an iframe. That works
// for PDF, because the browser has a built-in PDF viewer, and for little else — SharePoint serves
// an Office document as a DOWNLOAD, so the iframe silently renders nothing and the approver is
// asked to decide on a document they cannot see.
//
// Four strategies, because SharePoint needs three different endpoints and one honest refusal:
//
//   pdf    — the file URL, with #view=FitH so portrait pages fill the width
//   office — WopiFrame embedview, i.e. Office Online. The only way to render docx/xlsx in place
//   image  — the file URL in an <img>, which fits and centres far better than an iframe
//   none   — no viewer exists; say so and offer to open the file rather than show a blank pane
//
// Pure and SPFx-free so the extension table is testable, and so adding a type is one line here
// rather than a change inside a component.

export type PreviewKind = "pdf" | "office" | "image" | "text" | "none";

/**
 * Office types Office Online can render.
 *
 * `csv` is included: Excel Online opens it, and a CSV shown as a grid is far more reviewable than
 * the same file as raw text. The legacy binary formats are here too — a client running for twenty
 * years has .doc and .xls in circulation.
 */
const OFFICE = ["doc", "docx", "dot", "dotx", "xls", "xlsx", "xlsm", "csv", "ppt", "pptx", "pps", "ppsx"];

/** Raster and vector images browsers render natively. */
const IMAGE = ["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg", "ico"];

/**
 * Plain-text types a browser shows inline in an iframe.
 *
 * Deliberately narrow. Anything unlisted falls through to `none`, which offers the file instead of
 * pretending to display it — a blank iframe is the failure this module exists to remove, so
 * guessing that an unknown extension is text would reintroduce it.
 */
const TEXT = ["txt", "log", "json", "xml", "md"];

/** Lowercase extension without the dot. "" when there is none. */
export function extensionOf(fileName: string): string {
  const n = (fileName ?? "").trim().toLowerCase();
  const dot = n.lastIndexOf(".");
  if (dot <= 0 || dot === n.length - 1) return "";
  return n.slice(dot + 1);
}

/** Which strategy renders this file. */
export function previewKind(fileName: string): PreviewKind {
  const ext = extensionOf(fileName);
  if (ext === "pdf") return "pdf";
  if (OFFICE.indexOf(ext) !== -1) return "office";
  if (IMAGE.indexOf(ext) !== -1) return "image";
  if (TEXT.indexOf(ext) !== -1) return "text";
  return "none";
}

export interface PreviewTarget {
  kind: PreviewKind;
  /** The URL to render. Empty for `none`. */
  url: string;
  /**
   * The file itself. **The download route** — keep it for a Download control, never for "open".
   */
  fileUrl: string;
  /**
   * Where "Open in a new tab" should actually go.
   *
   * ⚠ NOT `fileUrl`. Navigating to a document's own URL makes SharePoint SERVE it, and for an
   * Office file that means a DOWNLOAD — reported by the client 2026-08-30 on both the approval
   * preview and My Submissions: *"the system download the uploaded file, instead of opening it in
   * a new tab."* The link said one thing and did another.
   *
   * Office goes through WopiFrame in `view` mode, which opens the document in the browser editor.
   * Everything else gets `?web=1`, the flag that tells SharePoint to RENDER a file rather than
   * hand it over — without it a PDF downloads on some tenants too, which is the same complaint
   * waiting to happen on a different file type.
   */
  openUrl: string;
}

/** Percent-encode each path segment, leaving the separators intact. */
export function encodePath(serverRelativeUrl: string): string {
  return (serverRelativeUrl ?? "")
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/");
}

/**
 * Build the preview target.
 *
 * `tenantRoot` is the origin with no /sites/… , because a server-relative path already carries the
 * site. `webUrl` is the absolute web URL, which WopiFrame needs.
 *
 * WopiFrame takes the RAW server-relative path in `sourcedoc`, percent-encoded once as a query
 * value — not the already-encoded path. Encoding it twice yields a "file not found" page that
 * looks exactly like a permissions problem.
 */
export function previewTarget(
  fileName: string,
  serverRelativeUrl: string,
  tenantRoot: string,
  webUrl: string,
): PreviewTarget {
  const kind = previewKind(fileName);
  const fileUrl = `${tenantRoot}${encodePath(serverRelativeUrl)}`;
  /* `?web=1` asks SharePoint to render rather than serve. Appended with `?` because `fileUrl` is a
     bare path with no query of its own — and BEFORE any fragment, which is why the PDF branch
     builds its own rather than concatenating onto this. */
  const webOpen = `${fileUrl}?web=1`;
  if (kind === "office") {
    const wopi = `${webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}`;
    return {
      kind,
      url: `${wopi}&action=embedview`,
      fileUrl,
      // `view`, not `embedview`: the embedded chrome is right inside an iframe and wrong in a tab
      // of its own, where the reader expects the full viewer.
      openUrl: `${wopi}&action=view`,
    };
  }
  if (kind === "pdf") {
    // #view=FitH is a PDF Open Parameter the browser's native viewer honours — fits the page to
    // the iframe's width rather than its height, which otherwise leaves gutters on portrait pages.
    return { kind, url: `${fileUrl}#view=FitH`, fileUrl, openUrl: webOpen };
  }
  if (kind === "image" || kind === "text") return { kind, url: fileUrl, fileUrl, openUrl: webOpen };
  return { kind: "none", url: "", fileUrl, openUrl: webOpen };
}
