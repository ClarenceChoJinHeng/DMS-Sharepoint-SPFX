/**
 * The file view — preview on the left, the document's details on the right.
 *
 * Lifted out of My Submissions (client, 2026-09-10: *"Reuse the my submission view file component for
 * the my request"*) so the Requests page shows an approver the SAME view an uploader sees, and the
 * request emails can land on it. ONE component, two mount points — a second copy is how the two views
 * would come to disagree about what a document carries, and the drifting copy would be the one nobody
 * tests.
 *
 * It renders what it is given and fetches nothing: each page resolves the file its own way (My
 * Submissions by library + item id, the Requests page by the request's recorded or stamped identity),
 * and the loaded `fieldText` is handed in.
 */
import * as React from "react";
import { previewTarget } from "./filePreview";
import { buildDetailRows, formatBytes } from "./documentDetails";
import { formatSubmittedOn } from "./mySubmissions";

export interface FileDetailPanelProps {
  /** The file's name, for the image alt text and the iframe title. */
  name: string;
  /** Server-relative path — `/sites/CRS/<Library>/…/<file>`. */
  fileRef: string;
  /** Origin with no path, e.g. `https://tenant.sharepoint.com`. */
  tenantRoot: string;
  /** The web's absolute URL. */
  siteUrl: string;
  /** `FieldValuesAsText` for the item. `undefined` while it loads; `{}` when it could not be read. */
  fieldText: Record<string, string> | undefined;
  /** The folder trail, already made readable by the caller. Blank reads as the library root. */
  location: string;
  size?: string | number;
  modified?: Date;
  /** Renames the top tier row for a Group-Led Project. Omitted keeps `Business Segment`. */
  segmentLabel?: string;
}

const s: Record<string, React.CSSProperties> = {
  detailGrid: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 300px)", gap: 20, alignItems: "start" },
  sectionTitle: { fontSize: 14, fontWeight: 600, color: "#201f1e", marginBottom: 10 },
  detailRow: { marginBottom: 12 },
  detailLabel: { fontSize: 11, fontWeight: 600, color: "#605e5c", textTransform: "uppercase", letterSpacing: 0.3 },
  detailValue: { fontSize: 13, color: "#201f1e", marginTop: 2, overflowWrap: "break-word" },
  imageBox: { display: "flex", alignItems: "flex-start", justifyContent: "center", width: "100%", height: "calc(100vh - 320px)", minHeight: 520, background: "#faf9f8", border: "1px solid #edebe9", borderRadius: 4, overflow: "auto", padding: 12, boxSizing: "border-box" },
  empty: { fontSize: 13, color: "#605e5c", padding: "28px 4px", lineHeight: 1.6 },
  link: { color: "#0f6cbd", textDecoration: "none", fontWeight: 600 },
};

export function FileDetailPanel(props: FileDetailPanelProps): React.ReactElement {
  const { name, fileRef, tenantRoot, siteUrl, fieldText } = props;
  const preview = previewTarget(name, fileRef, tenantRoot, siteUrl);
  /* The metadata rows come from shared/documentDetails.ts, which DERIVES the tier rows instead of
     naming them — a hardcoded Department/Unit pair read blank on every segment that names its tiers
     differently (Region / Estate·Mill on Upstream Ops, reported 2026-08-14). */
  const details = buildDetailRows({
    fieldText: fieldText ?? {},
    segmentLabel: props.segmentLabel,
    leading: [{ label: "Location", value: props.location || "the library root" }],
    trailing: [
      { label: "File size", value: formatBytes(props.size) || "unknown" },
      // Last updated is what changed when the approver acted.
      { label: "Last updated", value: formatSubmittedOn(props.modified) },
    ],
  });

  return (
    <>
      {/* ⚠⚠ NO BACKTICKS INSIDE THIS TEMPLATE LITERAL, not even in a CSS comment — one ends it.
          ⚠ A MEDIA QUERY, NOT a container query: both host pages render fixed dialogs, and container
          containment would shrink them to this grid. !important because the template is inline. */}
      <style>{`
        @media (max-width: 640px) {
          .crs-ms-detail { grid-template-columns: minmax(0, 1fr) !important; }
        }
      `}</style>
      <div className="crs-ms-detail" style={s.detailGrid}>
        <div>
          <div style={s.sectionTitle}>Preview</div>
          {preview.kind === "image" ? (
            // Fit to WIDTH and scroll: fitting both dimensions shrinks a tall screenshot to a sliver.
            <div style={s.imageBox}>
              <img src={preview.url} alt={name} style={{ maxWidth: "100%", height: "auto", display: "block" }} />
            </div>
          ) : preview.kind === "none" ? (
            <div style={{ ...s.imageBox, alignItems: "center", justifyContent: "center", color: "#605e5c" }}>
              <div style={{ textAlign: "center" }}>
                No preview is available for this file type.
                <div style={{ marginTop: 8 }}>
                  {/* The one surviving "open in a new tab" link: here the preview cannot render
                      anything, so it is the only route to the document, not decoration. */}
                  <a style={s.link} href={preview.openUrl} target="_blank" rel="noopener noreferrer">
                    Open it in a new tab
                  </a>
                </div>
              </div>
            </div>
          ) : (
            <iframe
              src={preview.url}
              style={{ width: "100%", height: "calc(100vh - 320px)", minHeight: 520, border: "1px solid #edebe9", borderRadius: 4, display: "block" }}
              title={`Preview of ${name}`}
              allowFullScreen
            />
          )}
        </div>

        <div>
          <div style={s.sectionTitle}>Details</div>
          {fieldText === undefined && <p style={s.empty}>Loading details&hellip;</p>}
          {/* Empty and unreadable look the same from here, so say the honest thing. Location and the
              file facts always have a value, so `details` is never empty and cannot carry this. */}
          {fieldText !== undefined && Object.keys(fieldText).length === 0 && (
            <p style={{ fontSize: 12, color: "#605e5c", lineHeight: 1.5 }}>
              No details were recorded for this file, or they could not be read.
            </p>
          )}
          {details.map(({ label, value }) => (
            <div key={label} style={s.detailRow}>
              <div style={s.detailLabel}>{label}</div>
              <div style={s.detailValue}>{value}</div>
            </div>
          ))}
        </div>
      </div>
    </>
  );
}
