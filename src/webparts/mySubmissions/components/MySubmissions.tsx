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
  sortNewestFirst,
  statusToDecision,
  submissionKey,
  trailText,
} from "../../../shared/mySubmissions";
import { libraryTitle, libraryUrlSegment } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

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
};

function badgeFor(status: string): React.CSSProperties {
  if (status === "Approved") return { ...s.badge, ...s.bOk };
  if (status === "Rejected") return { ...s.badge, ...s.bNo };
  return { ...s.badge, ...s.bPending };
}

/** A row as the reads return it. Field names are internal names, exactly. */
interface RawRow {
  Id: number;
  FileLeafRef?: string;
  FileRef?: string;
  Created?: string;
  OData__ModerationStatus?: number;
  OData__ModerationComments?: string;
  Document_x0020_Type?: string;
}

export default function MySubmissions({ context }: IMySubmissionsProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const [tab, setTab] = useState("All");
  const [rows, setRows] = useState<Submission[] | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [commentsMissing, setCommentsMissing] = useState(false);

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
    const common = "Id,FileLeafRef,FileRef,Created,Document_x0020_Type";
    const withStatus = approvalLibrary ? `${common},OData__ModerationStatus` : common;
    const get = (select: string): Promise<SPHttpClientResponse> =>
      context.spHttpClient.get(
        `${base}?$select=${select}&$filter=AuthorId eq ${userId}&$top=2000&$orderby=Created desc`,
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
    return ((data.value ?? []) as RawRow[]).map((r) => ({
      itemId: r.Id,
      library: urlSegment,
      name: (r.FileLeafRef ?? "").trim(),
      fileRef: (r.FileRef ?? "").trim(),
      // Documents rows carry no status: they are there because they were approved.
      status: approvalLibrary ? statusToDecision(Number(r.OData__ModerationStatus)) : "Approved",
      created: r.Created ? new Date(r.Created) : undefined,
      documentType: (r.Document_x0020_Type ?? "").trim(),
      comment: (r.OData__ModerationComments ?? "").trim(),
    }));
  };

  useEffect(() => {
    const load = async (): Promise<void> => {
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const userId = await currentUserId();
      // Sequential, not Promise.all: per-call try/catch is required anyway because
      // Promise.allSettled is unavailable on this tsconfig target (CLAUDE.md #3), and either
      // library failing must produce the "could not read" state rather than a half list.
      const staging = await readLibrary(libraryTitle(), libraryUrlSegment(), true, userId);
      const documents = await readLibrary(DOCUMENTS, DOCUMENTS, false, userId);
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
  const libs = [libraryUrlSegment(), DOCUMENTS];

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
              <th style={s.th}>Type</th>
            </tr>
          </thead>
          <tbody>
            {shown.map((r) => (
              <tr key={submissionKey(r)}>
                <td style={s.td}>
                  {/* FileRef is already server-relative, so it is the href as it stands. A pending
                      file is the uploader's OWN, so they can always open it — that is the author
                      exception in Draft Item Security. */}
                  <a style={s.link} href={r.fileRef} target="_blank" rel="noopener noreferrer">
                    {r.name}
                  </a>
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
                <td style={s.td}>{r.documentType || "—"}</td>
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
