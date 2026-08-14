// My Submissions — what happened to the files I uploaded.
//
// Spec: docs/superpowers/specs/2026-08-14-my-submissions-design.md
//
// VIEW ONLY, and uploaders only. The access gate is the PAGE grant (Page Access → UPL groups), not
// code: pageAccessPolicy.ts carries a dedicated rule so this page is never offered to approvers or
// deleters. An admin who opens it sees their own uploads, which is correct.
//
// TWO LIBRARIES, because a file's life spans two. Pending and rejected files sit in the approval
// library; an approved one has been MOVED to Documents by Auto-route and deleted from the source.
// Reading one library would show half a lifecycle. Both reads filter on AuthorId, and Auto-route
// preserves the uploader in Author (verified 2026-08-08), which is what makes the halves joinable.
//
// WHAT THIS PAGE DOES NOT DO: make approved documents private. In Documents a unit's approved files
// are readable by that whole unit, by design (one-group-per-person, 2026-08-09). Filtering to the
// signed-in user here is a convenience, not a boundary — the client was told and accepted this on
// 2026-08-14. Pending and rejected privacy IS real, and comes from Draft Item Security.
import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IMySubmissionsProps } from "./IMySubmissionsProps";
import {
  Submission,
  countByStatus,
  filterByTab,
  folderTrail,
  formatSubmittedOn,
  pickField,
  sortNewestFirst,
  statusToDecision,
  submissionKey,
  textOf,
  trailText,
} from "../../../shared/mySubmissions";
import { libraryTitle, libraryUrlSegment } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
// The preview strategy is ALREADY built and tested for the approval page: PDF, Office Online,
// image, text, and an honest refusal. Reusing it rather than re-guessing, because the case that
// matters is invisible until it bites — SharePoint serves an Office file as a DOWNLOAD, so a raw
// URL in an iframe renders nothing at all.
import { previewTarget } from "../../../shared/filePreview";

const DOCUMENTS = "Documents";
const TABS = ["All", "Pending", "Approved", "Rejected"];

const s: Record<string, React.CSSProperties> = {
  // Capped and centred like the other full-page screens, so a wide monitor does not stretch the
  // rows into unreadable ribbons.
  wrap:     { fontFamily: "'Segoe UI', sans-serif", color: "#1b1b1b", maxWidth: 1100, margin: "32px auto", padding: "0 24px 48px" },
  h2:       { fontSize: 22, fontWeight: 700, margin: "0 0 4px" },
  sub:      { fontSize: 13, color: "#666", margin: "0 0 20px" },
  tabs:     { display: "flex", gap: 4, borderBottom: "1px solid #edebe9", marginBottom: 16, flexWrap: "wrap" },
  tab:      { background: "none", border: "none", borderBottom: "2px solid transparent", color: "#605e5c", fontSize: 13, fontFamily: "inherit", padding: "8px 14px", cursor: "pointer", marginBottom: -1 },
  tabOn:    { borderBottom: "2px solid #0f6c3f", color: "#0f6c3f", fontWeight: 700 },
  table:    { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th:       { textAlign: "left", padding: "8px 10px", borderBottom: "2px solid #edebe9", fontWeight: 600, color: "#555", fontSize: 12 },
  td:       { padding: "9px 10px", borderBottom: "1px solid #f4f4f4", verticalAlign: "top" },
  link:     { color: "#0f6cbd", textDecoration: "none", fontWeight: 600 },
  trail:    { fontSize: 12, color: "#605e5c" },
  badge:    { display: "inline-block", fontSize: 11, fontWeight: 700, padding: "2px 9px", borderRadius: 10, whiteSpace: "nowrap" },
  bPending: { color: "#7a5c00", background: "#fff4ce", border: "1px solid #f2d98c" },
  bOk:      { color: "#0f6c3f", background: "#e7f4ec", border: "1px solid #b7dcc4" },
  bNo:      { color: "#a4262c", background: "#fde7e9", border: "1px solid #f1b0b3" },
  comment:  { fontSize: 12, color: "#a4262c", marginTop: 4, lineHeight: 1.45 },
  empty:    { fontSize: 13, color: "#605e5c", padding: "28px 4px", lineHeight: 1.6 },
  errBox:   { fontSize: 13, padding: "12px 14px", borderRadius: 6, border: "1px solid #f1c9c9", background: "#fdf3f3", color: "#a4262c", lineHeight: 1.55, marginBottom: 16 },
  note:     { fontSize: 12, color: "#605e5c", marginTop: 24, paddingTop: 12, borderTop: "1px solid #f0f0f0", lineHeight: 1.55 },
  // A button styled as a link: it opens the in-page detail view, so it must not look or behave
  // like navigation away from the page.
  nameBtn:  { background: "none", border: "none", padding: 0, font: "inherit", fontWeight: 600, fontSize: 13, color: "#0f6cbd", cursor: "pointer", textAlign: "left" },
  // ── Detail view ──
  backBand: { background: "rgba(15, 108, 63, 0.08)", borderRadius: 4, padding: "10px 16px", marginBottom: 20 },
  backLink: { background: "none", border: "none", padding: 0, font: "inherit", fontSize: 14, fontWeight: 600, color: "#0f6c3f", cursor: "pointer" },
  // minmax(0, 1fr) on the preview column: a bare 1fr floors at the iframe's min-content width, so
  // the preview could never give ground. Learned on the approval page.
  detailGrid:  { display: "grid", gridTemplateColumns: "minmax(0, 1fr) 300px", gap: 20, alignItems: "start" },
  sectionTitle:{ fontSize: 14, fontWeight: 600, color: "#201f1e", marginBottom: 10 },
  detailRow:   { marginBottom: 12 },
  detailLabel: { fontSize: 11, fontWeight: 600, color: "#605e5c", textTransform: "uppercase", letterSpacing: 0.3 },
  detailValue: { fontSize: 13, color: "#201f1e", marginTop: 2, overflowWrap: "break-word" },
  imageBox:    { display: "flex", alignItems: "flex-start", justifyContent: "center", width: "100%", height: "calc(100vh - 320px)", minHeight: 520, background: "#faf9f8", border: "1px solid #edebe9", borderRadius: 4, overflow: "auto", padding: 12, boxSizing: "border-box" },
  rejectBox:   { fontSize: 13, padding: "12px 14px", borderRadius: 6, border: "1px solid #f1b0b3", background: "#fde7e9", color: "#a4262c", lineHeight: 1.55, marginBottom: 16 },
};

function badgeFor(status: string): React.CSSProperties {
  if (status === "Approved") return { ...s.badge, ...s.bOk };
  if (status === "Rejected") return { ...s.badge, ...s.bNo };
  return { ...s.badge, ...s.bPending };
}

/**
 * A row as the reads return it.
 *
 * An INDEX SIGNATURE, not a typed shape. Two reasons, both learned the hard way on 2026-08-14:
 *
 * 1. SharePoint's OData layer double-encodes underscores, so `Document_x0020_Type` can arrive as
 *    `Document_x005f_x0020_x005f_Type` — a key a fixed interface cannot even name.
 * 2. Declaring these fields `string` made TypeScript vouch for something SharePoint does not
 *    guarantee. `Document Type` is managed metadata and arrives as an OBJECT, so `.trim()` on it
 *    crashed the page while the build stayed green. A lie in a type is worse than no type.
 *
 * Everything is therefore read through `textOf` / `pickField`, which coerce what actually arrives.
 */
type RawRow = Record<string, unknown>;

export default function MySubmissions({ context }: IMySubmissionsProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  // Origin with no /sites/… — a server-relative path already carries the site, so previewTarget
  // needs the bare host to build an absolute file URL.
  const tenantRoot = siteUrl.replace(/^(https?:\/\/[^/]+).*$/, "$1");
  const [tab, setTab] = useState("All");
  const [rows, setRows] = useState<Submission[] | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [commentsMissing, setCommentsMissing] = useState(false);
  /**
   * The Documents library's real URL SEGMENT, resolved rather than assumed.
   *
   * Its title is `Documents` and its URL is `/Shared Documents/` — the two differ, exactly as they
   * do for the approval library (gotcha #12). Hardcoding one string for both put "Shared Documents"
   * at the front of every approved file's folder trail. Defaults to the title so a failed read
   * degrades to the old behaviour rather than to nothing.
   */
  const [docsSegment, setDocsSegment] = useState(DOCUMENTS);
  /** The row being examined. `undefined` = the list. */
  const [open, setOpen] = useState<Submission | undefined>(undefined);
  /** Per-item metadata labels for the open row. `undefined` while in flight. */
  const [fieldText, setFieldText] = useState<Record<string, string> | undefined>(undefined);

  const currentUserId = async (): Promise<number> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/currentuser?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`could not identify you (HTTP ${res.status})`);
    return Number((await res.json()).Id);
  };

  /**
   * Read one library, filtered to the signed-in user.
   *
   * `approvalLibrary` decides whether the moderation status is read at all: everything in
   * `Documents` is approved by definition — it arrived by being approved, and that library has
   * content approval OFF — so a status there would be meaningless at best.
   *
   * The moderation comment is asked for and RETRIED WITHOUT on failure. A `$select` naming a column
   * that does not exist fails the whole request with HTTP 400 — not a null, not a missing key — so
   * asking for it unconditionally would blank this page on any site that lacks it. The same trap as
   * `Scope`/`Target` on the Group Map, and as gotcha #11.
   */
  const readLibrary = async (
    listTitle: string,
    urlSegment: string,
    approvalLibrary: boolean,
    userId: number,
  ): Promise<Submission[]> => {
    const base = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`;
    const common = "Id,FileLeafRef,FileRef,Created";
    const withStatus = approvalLibrary ? `${common},OData__ModerationStatus` : common;
    // FSObjType eq 0 = FILES ONLY. Without it every folder the uploader ever caused to be
    // created comes back as a "submission": 443 rows on the test site, and clicking one opened
    // the library rather than a document, because a folder's link IS a library link. The uploader
    // asked what happened to their FILES.
    //
    // FSObjType, not FileSystemObjectType: this is a $filter, and the REST name is rejected there
    // the same way the CAML name is rejected in a $select (memory sp-caml-internal-vs-rest-field-names).
    const get = (select: string): Promise<SPHttpClientResponse> =>
      context.spHttpClient.get(
        `${base}?$select=${select}&$filter=AuthorId eq ${userId} and FSObjType eq 0` +
          "&$top=2000&$orderby=Created desc",
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );

    let res = approvalLibrary
      ? await get(`${withStatus},OData__ModerationComments`)
      : await get(withStatus);
    if (!res.ok && approvalLibrary) {
      // The comment is the only optional field, so a failure earns one retry without it before
      // giving up on the whole library.
      setCommentsMissing(true);
      res = await get(withStatus);
    }
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`${listTitle}: HTTP ${res.status} ${body.slice(0, 140)}`);
    }

    const data = await res.json();
    return ((data.value ?? []) as RawRow[]).map((r) => {
      const createdRaw = textOf(r.Created);
      return {
        itemId: Number(r.Id),
        library: urlSegment,
        name: pickField(r, "FileLeafRef"),
        fileRef: pickField(r, "FileRef"),
        // Documents rows carry no status: they are there because they were approved.
        status: approvalLibrary ? statusToDecision(Number(r.OData__ModerationStatus)) : "Approved",
        created: createdRaw.length > 0 ? new Date(createdRaw) : undefined,
        comment: pickField(r, "OData__ModerationComments", "OData__x005f_ModerationComments"),
      };
    });
  };

  /**
   * The URL segment of a library, read off its own root folder.
   *
   * The only reliable source: a list's title and its URL are independent, and SharePoint never
   * moves the URL on rename. Falls back to the title, which is what the code did before and is
   * right on a site where they happen to match.
   */
  const resolveSegment = async (listTitle: string): Promise<string> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/RootFolder?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return listTitle;
      const url = textOf((await res.json()).ServerRelativeUrl);
      const last = url.split("/").filter((p) => p.length > 0).pop();
      return last && last.length > 0 ? last : listTitle;
    } catch {
      return listTitle;
    }
  };

  /**
   * Metadata labels for one item, from FieldValuesAsText.
   *
   * PER-ITEM by nature, which is why the list does not show metadata at all: one request per row
   * would be hundreds. Here it is one request for the one file being looked at.
   *
   * It also returns LABELS, not raw values — which is what makes a taxonomy field readable. The
   * plain `$select` gives a lookup id, and a bare `15` reached the screen on 2026-08-14.
   */
  const loadFieldText = async (row: Submission): Promise<void> => {
    setFieldText(undefined);
    const listTitle = row.library === docsSegment ? DOCUMENTS : libraryTitle();
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${row.itemId})/FieldValuesAsText`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      // {} rather than undefined on failure: the detail view still shows the file and its status,
      // and says the details could not be read. Labels are never worth blocking the preview for.
      setFieldText(res.ok ? ((await res.json()) as Record<string, string>) : {});
    } catch {
      setFieldText({});
    }
  };

  const openRow = (row: Submission): void => {
    setOpen(row);
    loadFieldText(row).catch(() => setFieldText({}));
  };

  useEffect(() => {
    const load = async (): Promise<void> => {
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const segment = await resolveSegment(DOCUMENTS);
      setDocsSegment(segment);
      const userId = await currentUserId();
      // Sequential, not Promise.all: per-call try/catch is required anyway because
      // Promise.allSettled is unavailable on this tsconfig target (CLAUDE.md #3), and either
      // library failing must produce the "could not read" state rather than a half list.
      const staging = await readLibrary(libraryTitle(), libraryUrlSegment(), true, userId);
      // The resolved segment, not the title — otherwise "Shared Documents" stays in every
      // approved file's folder trail.
      const documents = await readLibrary(DOCUMENTS, segment, false, userId);
      setRows(sortNewestFirst([...staging, ...documents]));
      setLoadError(undefined);
    };
    load().catch((e) => {
      // NEVER fall through to an empty list. An uploader told they have nothing, when a library was
      // merely unreachable, uploads the file again — and now there are two.
      setRows(undefined);
      setLoadError((e as Error).message);
    });
  }, []);

  const counts = countByStatus(rows ?? []);
  const shown = filterByTab(rows ?? [], tab);
  const libs = [libraryUrlSegment(), docsSegment, DOCUMENTS];

  /* ── The detail view ──────────────────────────────────────────────────────────
     An INNER view, not a link out. The client's point: clicking a file used to leave the page for
     the document library, which is the thing this page exists to save them from.

     The preview strategy comes from shared/filePreview.ts — already built and tested for the
     approval page, and already handling the case that matters: SharePoint serves an Office file as
     a DOWNLOAD, so a raw URL in an iframe renders nothing at all. PDF, Office, image, text and an
     honest refusal are each handled there rather than re-guessed here. */
  if (open !== undefined) {
    const preview = previewTarget(open.name, open.fileRef, tenantRoot, siteUrl);
    // FieldValuesAsText keys arrive double-encoded, so each row asks for both spellings — the same
    // reason pickField exists. Blank values are dropped rather than shown as empty rows.
    const details: Array<[string, string]> = [
      ["Document Type", pickField(fieldText ?? {}, "Document_x005f_x0020_x005f_Type", "Document_x0020_Type")],
      ["Year", pickField(fieldText ?? {}, "Year", "Year_x005f_x002f_x005f_Period", "Year_x002f_Period")],
      ["Document Date", pickField(fieldText ?? {}, "DocumentDate")],
      ["Business Segment", pickField(fieldText ?? {}, "Business_x005f_x0020_x005f_Segment", "Business_x0020_Segment")],
      ["Department", pickField(fieldText ?? {}, "Department")],
      ["Unit", pickField(fieldText ?? {}, "Unit")],
      ["Confidentiality", pickField(fieldText ?? {}, "Confidentiality_x005f_x0020_x005f_Level", "Confidentiality_x0020_Level")],
      ["Legally Privileged", pickField(fieldText ?? {}, "LegallyPrivileged")],
      ["Project Name", pickField(fieldText ?? {}, "ProjectName", "Project_x005f_x0020_x005f_Name")],
      ["Vendor / Customer", pickField(fieldText ?? {}, "Vendor_x005f_x002f_x005f_CustomerName", "Vendor_x002f_CustomerName")],
      ["Remark", pickField(fieldText ?? {}, "Remark")],
    ].filter(([, v]) => v.length > 0) as Array<[string, string]>;

    return (
      <section style={s.wrap}>
        <div style={s.backBand}>
          <button style={s.backLink} onClick={() => { setOpen(undefined); setFieldText(undefined); }}>
            ‹ Back to my submissions
          </button>
        </div>

        <h2 style={s.h2}>{open.name}</h2>
        <p style={s.sub}>
          <span style={badgeFor(open.status)}>{open.status}</span>{" "}
          &nbsp;Uploaded {formatSubmittedOn(open.created)} &nbsp;·&nbsp;{" "}
          {trailText(folderTrail(open.fileRef, libs)) || "—"}
        </p>

        {open.status === "Rejected" && (
          <div style={s.rejectBox}>
            <strong>This file was rejected.</strong>{" "}
            {open.comment
              ? open.comment
              : "No reason was recorded. Ask your approver what needs changing."}
          </div>
        )}

        <div style={s.detailGrid}>
          <div>
            <div style={s.sectionTitle}>Preview</div>
            {preview.kind === "image" ? (
              // Fit to WIDTH and scroll, the same fix the approval page needed: fitting BOTH
              // dimensions shrinks a tall screenshot to an unreadable sliver.
              <div style={s.imageBox}>
                <img src={preview.url} alt={open.name} style={{ maxWidth: "100%", height: "auto", display: "block" }} />
              </div>
            ) : preview.kind === "none" ? (
              <div style={{ ...s.imageBox, alignItems: "center", justifyContent: "center", color: "#605e5c" }}>
                <div style={{ textAlign: "center" }}>
                  No preview is available for this file type.
                  <div style={{ marginTop: 8 }}>
                    <a style={s.link} href={preview.fileUrl} target="_blank" rel="noopener noreferrer">
                      Open it in a new tab
                    </a>
                  </div>
                </div>
              </div>
            ) : (
              <iframe
                src={preview.url}
                style={{ width: "100%", height: "calc(100vh - 320px)", minHeight: 520, border: "1px solid #edebe9", borderRadius: 4, display: "block" }}
                title={`Preview of ${open.name}`}
                allowFullScreen
              />
            )}
            <div style={{ marginTop: 6, textAlign: "right" }}>
              <a style={s.link} href={preview.fileUrl} target="_blank" rel="noopener noreferrer">
                Open in a new tab
              </a>
            </div>
          </div>

          <div>
            <div style={s.sectionTitle}>Details</div>
            {fieldText === undefined && <p style={s.empty}>Loading details&hellip;</p>}
            {fieldText !== undefined && details.length === 0 && (
              // Empty and unreadable look the same from here, so say the honest thing: the file and
              // its status are still correct above, which is what the page is for.
              <p style={{ fontSize: 12, color: "#605e5c", lineHeight: 1.5 }}>
                No details were recorded for this file, or they could not be read.
              </p>
            )}
            {details.map(([label, value]) => (
              <div key={label} style={s.detailRow}>
                <div style={s.detailLabel}>{label}</div>
                <div style={s.detailValue}>{value}</div>
              </div>
            ))}
          </div>
        </div>
      </section>
    );
  }

  return (
    <section style={s.wrap}>
      <h2 style={s.h2}>My Submissions</h2>
      <p style={s.sub}>
        Every file you have uploaded, and where it has got to. Only your own files are listed here.
      </p>

      {loadError !== undefined && (
        <div style={s.errBox}>
          <strong>Could not read your submissions.</strong> {loadError}
          <div style={{ marginTop: 6 }}>
            This is not the same as having none — please try again rather than uploading again.
          </div>
        </div>
      )}

      <div style={s.tabs}>
        {TABS.map((t) => (
          <button
            key={t}
            style={{ ...s.tab, ...(tab === t ? s.tabOn : {}) }}
            onClick={() => setTab(t)}
          >
            {t} ({counts[t] ?? 0})
          </button>
        ))}
      </div>

      {rows === undefined && loadError === undefined && <p style={s.empty}>Loading&hellip;</p>}

      {/* Three empty states, kept apart on purpose — spec §4.2. */}
      {rows !== undefined && rows.length === 0 && (
        <p style={s.empty}>
          You have not uploaded anything yet. Files you submit on the upload form appear here, with
          their approval status.
        </p>
      )}
      {rows !== undefined && rows.length > 0 && shown.length === 0 && (
        <p style={s.empty}>
          Nothing {tab === "All" ? "here" : `is ${tab.toLowerCase()}`} right now. Your other files
          are under the tabs above.
        </p>
      )}

      {shown.length > 0 && (
        <table style={s.table}>
          <thead>
            <tr>
              <th style={s.th}>File</th>
              <th style={s.th}>Status</th>
              <th style={s.th}>Uploaded</th>
            </tr>
          </thead>
          <tbody>
            {shown.map((r) => (
              <tr key={submissionKey(r)}>
                <td style={s.td}>
                  {/* A button, not a link. It opens the detail view INSIDE this page — the client's
                      request, and the fix for a click that used to land the uploader in the
                      document library. The file itself is still one click further, from there. */}
                  <button style={s.nameBtn} onClick={() => openRow(r)}>
                    {r.name}
                  </button>
                  <div style={s.trail}>{trailText(folderTrail(r.fileRef, libs)) || "—"}</div>
                  {r.status === "Rejected" && (
                    <div style={s.comment}>
                      {r.comment
                        ? `Reason: ${r.comment}`
                        : "No reason was recorded. Ask your approver what needs changing."}
                    </div>
                  )}
                </td>
                <td style={s.td}>
                  <span style={badgeFor(r.status)}>{r.status}</span>
                </td>
                <td style={s.td}>{formatSubmittedOn(r.created)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      <p style={s.note}>
        <strong>Pending</strong> means waiting for your approver. <strong>Rejected</strong> files stay
        here so you can see why. <strong>Approved</strong> files have moved into the main {DOCUMENTS}{" "}
        library, where the rest of your unit can find them too.
        {commentsMissing && (
          <>
            {" "}
            Rejection reasons cannot be shown on this site — the comment column could not be read.
          </>
        )}
      </p>
    </section>
  );
}
