import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { IApprovalDocumentProps } from "./IApprovalDocumentProps";

// ── Types ────────────────────────────────────────────────────────────────────

// Library URL segments used to map a Staging path to its Documents twin. "Shared Documents"
// is the Documents library's URL segment even though its display name is "Documents".
const STAGING_URL_SEGMENT = "Staging";
const DOCUMENTS_URL_SEGMENT = "Shared Documents";

interface IFileItem {
  ID: number;
  FileLeafRef: string;
  OData__ModerationStatus: number;
  Author: { Title: string };
  Created: string;
  File: { Length: string; ServerRelativeUrl: string };
}

// SharePoint's OData layer double-encodes underscores in property names, so a field
// whose internal name already contains an encoded char (e.g. Document_x0020_Type,
// where _x0020_ is a space) comes back as Document_x005f_x0020_x005f_Type.
// Index signature so we can read whichever key SharePoint returns.
interface IFieldText {
  [internalName: string]: string;
}

type Decision = "Approved" | "Rejected" | "Pending";

// ── Helpers ──────────────────────────────────────────────────────────────────

function initials(name: string): string {
  return name.split(" ").slice(0, 2).map(n => n[0] ?? "").join("").toUpperCase();
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  const date = d.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" });
  const time = d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit", hour12: true });
  return `${date} ${time}`;
}

function formatSize(bytes: string): string {
  const b = parseInt(bytes, 10);
  if (isNaN(b)) return "";
  return b >= 1_048_576
    ? `${(b / 1_048_576).toFixed(1)} MB`
    : `${(b / 1_024).toFixed(1)} KB`;
}


// ── Styles ───────────────────────────────────────────────────────────────────

const s = {
  root:        { fontFamily: "'Segoe UI', Tahoma, sans-serif", color: "#323130", background: "#fff", padding: "0 0 40px" } as React.CSSProperties,
  backLink:    { color: "rgba(0, 104, 74, 1)", textDecoration: "none", fontSize: 14, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, marginBottom: 12 } as React.CSSProperties,
  docTitle:    { margin: "0 0 4px", fontSize: 38, fontWeight: 600, color: "#201f1e" } as React.CSSProperties,
  docMeta:     { fontSize: 13, color: "#605e5c", display: "flex", gap: 8, alignItems: "center", marginBottom: 24 } as React.CSSProperties,
  dot:         { color: "#c8c6c4" } as React.CSSProperties,
  // minmax(0, 1fr) on the centre column: a bare 1fr floors at the iframe's
  // min-content width, so the preview could never give ground. Narrower rails +
  // a tighter gap hand ~70px back to the preview.
  grid:        { display: "grid", gridTemplateColumns: "196px minmax(0, 1fr) 280px", gap: 16, alignItems: "start" } as React.CSSProperties,
  sectionTitle:{ fontSize: 14, fontWeight: 600, color: "#201f1e", marginBottom: 12 } as React.CSSProperties,
  avatar:      { width: 36, height: 36, borderRadius: "50%", background: "#0f6cbd", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, flexShrink: 0 } as React.CSSProperties,
  authorName:  { fontWeight: 600, fontSize: 14, color: "#201f1e" } as React.CSSProperties,
  authorDate:  { fontSize: 12, color: "#605e5c", marginTop: 2 } as React.CSSProperties,
  // Details stack vertically — label above value — so a long value (a deep
  // Location path, a full vendor name) wraps into the column's own width
  // instead of being squeezed into whatever the label leaves of a 196px rail.
  detailRow:   { marginBottom: 14 } as React.CSSProperties,
  metaLabel:   { color: "#605e5c", fontSize: 13, marginBottom: 2 } as React.CSSProperties,
  metaValue:   { fontWeight: 500, fontSize: 13, color: "#201f1e", lineHeight: 1.35, overflowWrap: "break-word" as const } as React.CSSProperties,
  panel:       { border: "1px solid #edebe9", borderRadius: 4, padding: 20, position: "sticky" as const, top: 16, background: "#fff", boxShadow: "0 2px 6px rgba(0,0,0,0.08)" } as React.CSSProperties,
  panelTitle:  { fontWeight: 700, fontSize: 16, color: "#201f1e" } as React.CSSProperties,
  panelHint:   { fontSize: 13, color: "#605e5c", marginBottom: 16, lineHeight: 1.4 } as React.CSSProperties,
  textarea:    { width: "100%", height: 122, maxHeight: 122, padding: "6px 8px", fontSize: 13, border: "1px solid #8a8886", borderRadius: 2, resize: "vertical" as const, fontFamily: "inherit", boxSizing: "border-box" as const, color: "#201f1e" } as React.CSSProperties,
  charCount:   { fontSize: 12, color: "#605e5c", textAlign: "right" as const, marginTop: 2, marginBottom: 12 } as React.CSSProperties,
  publishLabel:{ fontSize: 13, fontWeight: 600, marginBottom: 6 } as React.CSSProperties,
  publishRow:  { display: "flex", alignItems: "center", gap: 6, marginBottom: 20 } as React.CSSProperties,
  publishValue:{ color: "#0f6cbd", fontSize: 13 } as React.CSSProperties,
  btnApprove:  { width: "100%", padding: "10px 0", marginBottom: 8, background: "#107c10", color: "#fff", border: "none", borderRadius: 4, fontSize: 14, fontWeight: 600, cursor: "pointer" } as React.CSSProperties,
  btnSendBack: { width: "100%", padding: "10px 0", marginBottom: 8, background: "#fff", color: "#201f1e", border: "1px solid #8a8886", borderRadius: 4, fontSize: 14, fontWeight: 600, cursor: "pointer" } as React.CSSProperties,
  btnReject:   { width: "100%", padding: "10px 0", background: "#fff", color: "#a4262c", border: "1px solid #a4262c", borderRadius: 4, fontSize: 14, fontWeight: 600, cursor: "pointer" } as React.CSSProperties,
  errText:     { color: "#a4262c", fontSize: 13, marginBottom: 10 } as React.CSSProperties,
  popupOverlay:{ position: "fixed" as const, inset: 0, background: "rgba(0,0,0,0.25)", backdropFilter: "blur(6px)", WebkitBackdropFilter: "blur(6px)", zIndex: 9998, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 } as React.CSSProperties,
  popupCard:   { background: "#fff", borderRadius: 16, padding: "40px 40px 32px", textAlign: "center" as const, maxWidth: 420, width: "100%", boxShadow: "0 8px 40px rgba(0,0,0,0.15)" } as React.CSSProperties,
  popupTitle:  { fontSize: 22, fontWeight: 700, color: "#0f6c3f", margin: "0 0 12px" } as React.CSSProperties,
  popupMsg:    { fontSize: 14, color: "#555", margin: "0 0 28px", lineHeight: 1.6 } as React.CSSProperties,
  popupBtn:    { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 6, padding: "12px 28px", fontSize: 14, fontWeight: 600, fontFamily: "inherit", cursor: "pointer", minWidth: 183 } as React.CSSProperties,
};

// ── Component ─────────────────────────────────────────────────────────────────

const ApprovalDocument: React.FC<IApprovalDocumentProps> = ({ context }) => {
  const [item, setItem]           = useState<IFileItem | null>(null);
  const [loading, setLoading]     = useState(true);
  const [decision, setDecision]   = useState<Decision>("Approved");
  const [comments, setComments]   = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState<Decision | null>(null);
  const [fetchError, setFetchError] = useState("");
  const [submitError, setSubmitError] = useState("");
  const [fieldText, setFieldText]   = useState<IFieldText>({});

  const webUrl = context.pageContext.web.absoluteUrl;

  const getItemId = (): number | null => {
    const p = new URLSearchParams(window.location.search);
    const id = p.get("itemId");
    return id ? parseInt(id, 10) : null;
  };

  // Back link → the file's own folder in Staging (not the library root), so the
  // approver lands where the document lives instead of having to drill back in.
  const backUrl = (): string => {
    const base = `${webUrl}/Staging/Forms/AllItems.aspx`;
    const fileRef = item?.File?.ServerRelativeUrl;
    if (!fileRef) return base;
    const folder = fileRef.slice(0, fileRef.lastIndexOf("/"));
    return `${base}?id=${encodeURIComponent(folder)}`;
  };

  const loadItem = async (): Promise<void> => {
    const itemId = getItemId();
    if (!itemId) {
      setFetchError("No document ID provided. The link should include ?itemId=123.");
      setLoading(false);
      return;
    }
    try {
      const url =
        `${webUrl}/_api/web/lists/getbytitle('Staging')/items(${itemId})` +
        `?$expand=File,Author` +
        `&$select=ID,FileLeafRef,OData__ModerationStatus,Created,Author/Title,File/Length,File/ServerRelativeUrl`;
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        const body = await res.text().catch(() => "");
        throw new Error(`HTTP ${res.status}: ${body.slice(0, 300)}`);
      }
      const data: IFileItem = await res.json();
      setItem(data);
      const statusMap: Record<number, Decision> = { 0: "Approved", 1: "Rejected", 2: "Pending" };
      setDecision(statusMap[data.OData__ModerationStatus] ?? "Pending");
      try {
        const textUrl = `${webUrl}/_api/web/lists/getbytitle('Staging')/items(${itemId})/FieldValuesAsText`;
        const textRes = await context.spHttpClient.get(textUrl, SPHttpClient.configurations.v1);
        if (textRes.ok) {
          const ft = await textRes.json() as IFieldText;
          setFieldText(ft);
        }
      } catch { /* labels are non-critical */ }
    } catch (err) {
      setFetchError(err instanceof Error ? err.message : "Could not load this document.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { loadItem().catch(() => undefined); }, []);

  const getDigest = async (): Promise<string> => {
    const res = await context.spHttpClient.post(
      `${webUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } }
    );
    const data = await res.json();
    return data.FormDigestValue as string;
  };

  /**
   * Is this document's destination in the Documents library ready to receive it?
   *
   * The Auto-route flow creates whatever path it is handed. If the unit folder is missing,
   * it silently creates the WHOLE chain from the library root, and those folders inherit
   * the Documents library ACL — where DMS_SITE_MEMBERS holds Read. Every DMS user would
   * then be able to read that unit's approved documents, with nothing logged as an error.
   *
   * So refuse to approve unless the unit folder exists AND has unique permissions.
   * An inconclusive answer also blocks: a retry costs the approver seconds, whereas a wrong
   * "proceed" publishes documents to everyone and nobody finds out.
   *
   * See docs/superpowers/specs/2026-07-29-approval-destination-guard-design.md
   */
  const documentsUnitFolderReady = async (
    stagingFileUrl: string,
  ): Promise<{ ok: boolean; reason?: string }> => {
    const webSru = context.pageContext.web.serverRelativeUrl;
    // …/Unit/Year/Document Type/file.ext → walk up three levels to the unit folder.
    // Position-based, so it holds for segments with a deeper Levels chain too.
    const parts = stagingFileUrl.split("/");
    if (parts.length < 4) return { ok: false, reason: "unexpected file path" };
    const unitStaging = parts.slice(0, parts.length - 3).join("/");

    // Swap the library segment, anchored on the web-relative prefix so a folder that
    // happens to be named "Staging" deeper in the tree is not mangled.
    const prefix = `${webSru}/${STAGING_URL_SEGMENT}/`;
    if (unitStaging.toLowerCase().indexOf(prefix.toLowerCase()) !== 0) {
      return { ok: false, reason: "could not work out the Documents path" };
    }
    const unitDocs = `${webSru}/${DOCUMENTS_URL_SEGMENT}/${unitStaging.slice(prefix.length)}`;
    // Per-segment encoding (no %2F flood) with OData quote doubling — see pathEncoding.ts.
    const encoded = unitDocs.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''");

    const res = await context.spHttpClient.get(
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields?$select=HasUniqueRoleAssignments&@f='${encoded}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (res.status === 404) return { ok: false, reason: "the folder does not exist yet" };
    if (!res.ok) {
      console.error("Approval destination check failed:", res.status, unitDocs);
      return { ok: false, reason: `it could not be verified (HTTP ${res.status})` };
    }
    const d = await res.json();
    if (d.HasUniqueRoleAssignments !== true) {
      return { ok: false, reason: "it is not locked down — it would be readable by every DMS user" };
    }
    return { ok: true };
  };

  const submitDecision = async (action: Decision): Promise<void> => {
    if (!item || submitting) return;
    setDecision(action);
    setSubmitting(true);
    setSubmitError("");
    try {
      // Only approving publishes to Documents. Rejecting copies nothing, so it is not gated.
      if (action === "Approved") {
        const ready = await documentsUnitFolderReady(item.File.ServerRelativeUrl);
        if (!ready.ok) {
          setSubmitError(
            `This unit's folder is not ready in the Documents library, so the document was NOT approved (${ready.reason}). ` +
            `Ask an administrator to run Folder Reconciliation, then approve again.`,
          );
          // Leave `decision` as the approver chose it — the selection is still valid, it is
          // the destination that is not ready.
          setSubmitting(false);
          return;
        }
      }
      const digest      = await getDigest();
      const safeComment = comments.replace(/'/g, "''");
      // Path passed as an OData parameter alias (@f) appended to the query
      // string, not embedded inline in the URL path — inline literals hit
      // IIS's maxUrlLength once the server-relative path gets long/deep
      // (nested subcategory nesting can add up fast). See sp-rest-alias-vs-inline-literal-400-error memory.
      const safeUrl     = item.File.ServerRelativeUrl.replace(/'/g, "''");
      const itemBase    = `${webUrl}/_api/web/lists/getbytitle('Staging')/items(${item.ID})`;
      const headers     = { "X-RequestDigest": digest, Accept: "application/json;odata=nometadata" };

      if (action === "Approved") {
        const res = await context.spHttpClient.post(
          `${webUrl}/_api/web/getfilebyserverrelativeurl(@f)/approve(comment='${safeComment}')?@f='${safeUrl}'`,
          SPHttpClient.configurations.v1,
          { headers },
        );
        if (!res.ok) throw new Error(`Approve returned ${res.status}: ${await res.text().catch(() => "")}`);
      } else if (action === "Rejected") {
        // /reject() fails when the item is already Approved; MERGE works regardless of current state.
        const rejectRes = await context.spHttpClient.fetch(itemBase, SPHttpClient.configurations.v1, {
          method: "POST",
          headers: { ...headers, "X-HTTP-Method": "MERGE", "IF-MATCH": "*", "Content-Type": "application/json;odata=nometadata" },
          body: JSON.stringify({ OData__ModerationStatus: 1, OData__ModerationComments: safeComment }),
        });
        if (!rejectRes.ok) throw new Error(`Reject returned ${rejectRes.status}: ${await rejectRes.text().catch(() => "")}`);

        // Delete the file from the Documents library if it was previously approved and routed there.
        // Search by filename so we don't need to guess the exact folder path.
        const safeLeaf = item.FileLeafRef.replace(/'/g, "''");
        const searchRes = await context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('Documents')/items?$filter=FileLeafRef eq '${safeLeaf}'&$select=FileRef&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (searchRes.ok) {
          const searchData = await searchRes.json() as { value: { FileRef: string }[] };
          const found = searchData.value?.[0];
          if (found?.FileRef) {
            const safeDocUrl = found.FileRef.replace(/'/g, "''");
            const delRes = await context.spHttpClient.fetch(
              `${webUrl}/_api/web/getfilebyserverrelativeurl(@f)?@f='${safeDocUrl}'`,
              SPHttpClient.configurations.v1,
              { method: "POST", headers: { ...headers, "X-HTTP-Method": "DELETE" } }
            );
            if (!delRes.ok && delRes.status !== 404) {
              throw new Error(`Delete from Documents returned ${delRes.status}: ${await delRes.text().catch(() => "")}`);
            }
          }
          // found is empty = file was never routed to Documents (not yet approved) — nothing to delete
        }
      } else if (action === "Pending") {
        const res = await context.spHttpClient.fetch(itemBase, SPHttpClient.configurations.v1, {
          method: "POST",
          headers: { ...headers, "X-HTTP-Method": "MERGE", "IF-MATCH": "*", "Content-Type": "application/json;odata=nometadata" },
          body: JSON.stringify({ OData__ModerationStatus: 2, OData__ModerationComments: safeComment }),
        });
        if (!res.ok) throw new Error(`Pending returned ${res.status}: ${await res.text().catch(() => "")}`);

        // Remove from Documents — file is no longer approved so readers shouldn't see it.
        const safeLeaf = item.FileLeafRef.replace(/'/g, "''");
        const searchRes = await context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('Documents')/items?$filter=FileLeafRef eq '${safeLeaf}'&$select=FileRef&$top=1`,
          SPHttpClient.configurations.v1
        );
        if (searchRes.ok) {
          const searchData = await searchRes.json() as { value: { FileRef: string }[] };
          const found = searchData.value?.[0];
          if (found?.FileRef) {
            const safeDocUrl = found.FileRef.replace(/'/g, "''");
            const delRes = await context.spHttpClient.fetch(
              `${webUrl}/_api/web/getfilebyserverrelativeurl(@f)?@f='${safeDocUrl}'`,
              SPHttpClient.configurations.v1,
              { method: "POST", headers: { ...headers, "X-HTTP-Method": "DELETE" } }
            );
            if (!delRes.ok && delRes.status !== 404) {
              throw new Error(`Delete from Documents returned ${delRes.status}: ${await delRes.text().catch(() => "")}`);
            }
          }
        }
      }
      setSubmitted(action);
    } catch (err) {
      console.error("[ApprovalDoc] submitDecision failed:", err);
      setSubmitError(err instanceof Error ? err.message : "Could not submit your decision. Please try again.");
    } finally {
      setSubmitting(false);
    }
  };

  // ── Loading ────────────────────────────────────────────────────────────────

  if (loading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 200, color: "#605e5c", fontSize: 14 }}>
        Loading document…
      </div>
    );
  }

  if (fetchError) {
    return <div style={{ padding: 32, color: "#a4262c", fontSize: 14 }}>{fetchError}</div>;
  }

  if (!item) return null;

  // ── Success ────────────────────────────────────────────────────────────────

  if (submitted) {
    const popups: Record<
      Decision,
      { title: string; message: string; icon: React.ReactNode }
    > = {
      Approved: {
        title: "Approval Successful",
        message: "Your document has been approved successfully.",
        icon: (
          <svg width="120" height="120" viewBox="0 0 184 184" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle opacity="0.3" cx="92.0001" cy="91.9999" r="75.4872" fill="#14C7A5" />
            <circle cx="92" cy="92" r="92" fill="#14C7A5" fillOpacity="0.2" />
            <circle cx="92.0003" cy="91.9998" r="61.3333" fill="white" stroke="#14C7A5" strokeWidth="3" />
            <path d="M102 106.841L111 114.841L122 99.8413" stroke="#14C7A5" strokeWidth="7" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M98 121H65L65 77.5L82 58H113V91" stroke="#14C7A5" strokeWidth="6" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M80 81H99" stroke="#14C7A5" strokeWidth="6" strokeLinecap="round" />
            <path d="M80 92H87" stroke="#14C7A5" strokeWidth="6" strokeLinecap="round" />
          </svg>
        ),
      },
      Rejected: {
        title: "Document Rejected",
        message: "The document has been returned for revision.",
        icon: (
          <svg width="120" height="120" viewBox="0 0 184 184" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle opacity="0.3" cx="92.0001" cy="91.9999" r="75.4872" fill="#FF4646" />
            <circle cx="92" cy="92" r="92" fill="#FF4646" fillOpacity="0.2" />
            <circle cx="92.0003" cy="91.9998" r="61.3333" fill="white" stroke="#FF4646" strokeWidth="3" />
            <path d="M119.801 98.5981L104.598 113.801" stroke="#FF4646" strokeWidth="7" strokeLinecap="round" />
            <path d="M104.598 98.605L119.801 113.808" stroke="#FF4646" strokeWidth="7" strokeLinecap="round" />
            <path d="M98 121H65L65 77.5L82 58H113V91" stroke="#FF4646" strokeWidth="6" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M80 81H99" stroke="#FF4646" strokeWidth="6" strokeLinecap="round" />
            <path d="M80 92H87" stroke="#FF4646" strokeWidth="6" strokeLinecap="round" />
          </svg>
        ),
      },
      Pending: {
        title: "Still pending Approval",
        message: "You can monitor its status anytime.",
        icon: (
          <svg width="120" height="120" viewBox="0 0 184 184" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle opacity="0.3" cx="92.0001" cy="91.9999" r="75.4872" fill="#FF8800" />
            <circle cx="92" cy="92" r="92" fill="#FF8800" fillOpacity="0.2" />
            <circle cx="92.0003" cy="91.9998" r="61.3333" fill="white" stroke="#FF8800" strokeWidth="3" />
            <path d="M98 121H65L65 77.5L82 58H113V91" stroke="#FF8800" strokeWidth="6" strokeLinecap="round" strokeLinejoin="round" />
            <circle cx="113" cy="110" r="11" stroke="#FF8800" strokeWidth="5" />
            <path d="M113 105V112L118 109.5" stroke="#FF8800" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M80 81H99" stroke="#FF8800" strokeWidth="6" strokeLinecap="round" />
            <path d="M80 92H87" stroke="#FF8800" strokeWidth="6" strokeLinecap="round" />
          </svg>
        ),
      },
    };
    const { title, message, icon } = popups[submitted];
    return (
      <div style={s.popupOverlay} role="dialog" aria-modal="true">
        <div style={s.popupCard}>
          <div style={{ marginBottom: 20 }}>{icon}</div>
          <div style={s.popupTitle}>{title}</div>
          <div style={s.popupMsg}>{message}</div>
          <button style={s.popupBtn} onClick={() => { window.location.href = backUrl(); }}>
            Back to Document
          </button>
        </div>
      </div>
    );
  }

  // ── Main ───────────────────────────────────────────────────────────────────

  const tenantRoot = webUrl.split('/sites/')[0];
  const encodedPath = item.File.ServerRelativeUrl
    .split('/')
    .map((s: string) => encodeURIComponent(s))
    .join('/');
  // #view=FitH is a PDF Open Parameter the browser's native PDF viewer honors —
  // fits the page to the iframe's width instead of its height, which otherwise
  // leaves empty gutters on portrait pages when the viewer defaults to fit-height.
  const previewUrl = `${tenantRoot}${encodedPath}#view=FitH`;

  // Read whichever key SharePoint returns: the double-encoded name (what the OData
  // response actually uses) first, then the plain internal name as a fallback.
  const pick = (...keys: string[]): string => {
    for (const k of keys) {
      const v = fieldText[k];
      if (v) return v;
    }
    return "—";
  };

  // Location = the org folder path (Segment › … › Unit), derived from the file's live
  // Staging path so it is segment-agnostic (GHO, Upstream, Projects — any depth). Drops
  // the deepest three path segments — Year, Document Type, and the filename — which are
  // shown separately below and are not part of the org location.
  const orgLocation = ((): string => {
    const after = item.File.ServerRelativeUrl.split("/Staging/")[1];
    if (!after) return "—";
    const parts = after.split("/");
    const org = parts.slice(0, Math.max(0, parts.length - 3));
    return org.length ? org.join(" › ") : "—";
  })();

  const metadata: [string, string][] = [
    ["Location",           orgLocation],
    ["Document Type",      pick("Document_x005f_x0020_x005f_Type", "Document_x0020_Type")],
    ["Confidential Level", pick("Confidentiality_x005f_x0020_x005f_Level", "Confidentiality_x0020_Level")],
    ["Year",               pick("Year", "Year_x005f_x002f_x005f_Period", "Year_x002f_Period")],
    // ProjectName has no encoded characters, so its response key is unencoded.
    ["Project Name",       pick("ProjectName")],
    // Vendor_x002f_CustomerName does, and FieldValuesAsText double-encodes the
    // underscores in response keys — hence the _x005f_ form first.
    ["Vendor/Customer Name", pick("Vendor_x005f_x002f_x005f_CustomerName", "Vendor_x002f_CustomerName")],
  ];

  // Pending is a system state only — approvers pick Approved or Rejected.
  const radioOptions: { val: Decision; label: string }[] = [
    { val: "Approved", label: "Approved" },
    { val: "Rejected", label: "Rejected" },
  ];

  return (
    <div style={s.root}>

      <a href={backUrl()} style={s.backLink}>
        {/* A bare "<" must be escaped as an expression — JSX reads a literal
            left angle bracket in children as the start of a tag. */}
        <span style={{ fontSize: 16, lineHeight: 1, fontWeight: 700 }}>{"<"}</span>
        Back to document list
      </a>

      <h1 style={s.docTitle}>{item.FileLeafRef}</h1>
      <div style={s.docMeta}>
        <img src={`${webUrl}/_layouts/15/images/ic${(item.FileLeafRef.split('.').pop() ?? 'txt').toLowerCase()}.png`} width={16} height={16} alt="" style={{ flexShrink: 0 }} />
        <span>{(item.FileLeafRef.split('.').pop() ?? 'File').toUpperCase()} document</span>
        <span style={s.dot}>•</span>
        <span>{formatSize(item.File.Length)}</span>
        <span style={s.dot}>•</span>
        <span>Uploaded on {formatDate(item.Created)}</span>
      </div>

      <div style={s.grid}>

        {/* Left — Uploaded by + Details */}
        <div>
          <div style={s.sectionTitle}>Uploaded by</div>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 28 }}>
            <div style={s.avatar}>{initials(item.Author.Title)}</div>
            <div>
              <div style={s.authorName}>{item.Author.Title}</div>
              <div style={s.authorDate}>{formatDate(item.Created)}</div>
            </div>
          </div>

          <hr style={{ border: "none", borderTop: "1px solid #edebe9", margin: "0 0 16px" }} />
          <div style={s.sectionTitle}>Details</div>
          {metadata.map(([label, value]) => (
            <div key={label} style={s.detailRow}>
              <div style={s.metaLabel}>{label}</div>
              <div style={s.metaValue}>{value}</div>
            </div>
          ))}
        </div>

        {/* Center — Preview */}
        <div>
          <div style={s.sectionTitle}>Preview</div>
          <iframe
            src={previewUrl}
            style={{ width: "100%", height: "calc(100vh - 200px)", minHeight: 640, border: "none", borderRadius: 4, display: "block" }}
            title="Document preview"
            allowFullScreen
          />
        </div>

        {/* Right — Approval panel */}
        <div style={s.panel}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <span style={s.panelTitle}>Approval</span>
          </div>
          <p style={s.panelHint}>Please review the document and its details.</p>

          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10 }}>Approval status</div>
            {radioOptions.map(({ val, label }) => (
              <label key={val} style={{ display: "block", marginBottom: 10, cursor: "pointer" }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <input
                    type="radio"
                    name="decision"
                    value={val}
                    checked={decision === val}
                    onChange={() => setDecision(val)}
                    style={{ flexShrink: 0 }}
                  />
                  <div style={{ fontSize: 13, fontWeight: decision === val ? 600 : 400 }}>{label}</div>
                </div>
              </label>
            ))}
          </div>

          <div>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 6 }}>Comment</div>
            <textarea
              value={comments}
              onChange={e => setComments(e.target.value.slice(0, 500))}
              placeholder="Use this field to enter any comments about why the item was approved or rejected."
              style={s.textarea}
            />
            <div style={s.charCount}>{comments.length}/500</div>
          </div>

          <div style={{ marginBottom: 20 }}>
            <div style={s.publishLabel}>Publish to</div>
            <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap" as const }}>
              {(() => {
                // Segment › … › Unit, from the live Staging path (segment-agnostic).
                const after = item.File.ServerRelativeUrl.split('/Staging/')[1];
                const parts = after ? after.split('/') : [];
                const crumbs = parts.slice(0, Math.max(0, parts.length - 3));
                return crumbs.map((crumb, i) => (
                  <React.Fragment key={crumb}>
                    {i > 0 && <span style={{ color: "#c8c6c4", fontSize: 12 }}>›</span>}
                    <span style={{ color: i === crumbs.length - 1 ? "#0f6cbd" : "#605e5c", fontSize: 13 }}>{crumb}</span>
                  </React.Fragment>
                ));
              })()}
            </div>
          </div>

          {submitError && <div style={s.errText}>{submitError}</div>}

          {/* Pending is no longer selectable — require an explicit Approved/Rejected choice. */}
          <button
            onClick={() => { submitDecision(decision).catch(() => undefined); }}
            disabled={submitting || decision === "Pending"}
            style={{ ...s.btnApprove, opacity: submitting || decision === "Pending" ? 0.7 : 1, cursor: decision === "Pending" ? "not-allowed" : "pointer" }}
          >
            {submitting ? "Saving…" : "Ok"}
          </button>
          <button onClick={() => { window.location.href = backUrl(); }} disabled={submitting} style={{ ...s.btnSendBack, opacity: submitting ? 0.7 : 1 }}>
            Cancel
          </button>
        </div>

      </div>
    </div>
  );
};

export default ApprovalDocument;
