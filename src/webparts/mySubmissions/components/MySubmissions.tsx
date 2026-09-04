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
import { useEffect, useRef, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IMySubmissionsProps } from "./IMySubmissionsProps";
import { PersonPick, searchTenantPeople } from "../../../shared/spGroups";
import { useLiveRefresh } from "../../../shared/liveRefresh";
import { groupSubmissions, statusCounts, SubmissionGroup } from "../../../shared/submissionGroups";
import {
  Submission,
  countByStatus,
  filterByTab,
  folderTrail,
  formatSubmittedOn,
  canActDirectly,
  formatSubmittedAt,
  pickField,
  sortNewestFirst,
  statusToDecision,
  submissionKey,
  textOf,
  trailText,
  isHcRow,
  isArchivedRow,
} from "../../../shared/mySubmissions";
/* The submission RECORD — what keeps a deleted file on this page (2026-08-27).
   Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
   `MergedRow` is a `Submission` plus `recordState`, so every existing helper here still applies. */
import {
  MergedRow, mergeRecords, mergedKey, liveRowsOnly, recordCounts, SubmissionRecord,
  isBulkUploadRow, snapshotFolderRows, snapshotFileRows, recordStateParts, archivedRowsOnly,
} from "../../../shared/submissionRecords";
import { readSubmissionRecords } from "../../../shared/spSubmissionRecords";
// The metadata panel's rows. It DERIVES the tier rows from the item's own fields rather than naming
// them, which is what makes Region/Estate·Mill appear on a segment nobody wrote code for.
import {
  buildDetailRows, buildBatchRows, buildFileRows,
  documentUnit, formatBytes, routeToApprover,
} from "../../../shared/documentDetails";
import {
  cachedHcLibraries,
  cachedArchiveLibraries,
  cachedListTitle,
  LIST_SUFFIX,
  libraryTitle,
  documentsLibraryTitle,
  DOCUMENTS_URL_SEGMENT,
  libraryUrlSegment,
} from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { normalizeRoleValue } from "../../../shared/groupMapModel";
// The request rules — validation, recipient parsing and the inside/outside test — live under test in
// shared/requests.ts and are shared with the approver's queue. Two copies of "what counts as external"
// is how one screen ends up permitting what the other refuses.
import {
  RequestDraft,
  RequestRow,
  RequestStage,
  RequestStatus,
  parseRequestStatus,
  canCancel,
  RequestType,
  SharePermission,
  isExternal,
  parseRecipients,
  validateDraft,
} from "../../../shared/requests";
import { writeAudit } from "../../../shared/spAuditLog";
import { EVENT } from "../../../shared/auditLog";
// The preview strategy is ALREADY built and tested for the approval page: PDF, Office Online,
// image, text, and an honest refusal. Reusing it rather than re-guessing, because the case that
// matters is invisible until it bites — SharePoint serves an Office file as a DOWNLOAD, so a raw
// URL in an iframe renders nothing at all.
import { previewTarget } from "../../../shared/filePreview";

/* ⚠ THE LITERAL `"Documents"` USED TO LIVE HERE, AND IT IS WHAT BROKE THIS PAGE ON 2026-08-28.
   The client retitled that library to `Restricted & Confidential Document`, every `getbytitle`
   below 404ed, and the page showed "Could not read your submissions" over an uploader's entire
   history. A LOCAL copy of a name that is resolved everywhere else — the third such copy found
   that day, after `libApiTitle` and `StructureManager`.
   Resolved at CALL time, never captured: `primeNames` has not run on the first render. */
const documentsTitle = (): string => documentsLibraryTitle();
/* "Requests" is not a document status like the other four — it lists what has been ASKED
   about those documents. Last, because it is the rarer errand. */
/* "Submissions" leads and is the default (client, 2026-08-22): what an uploader did is a submission,
   and the flat per-file list is the detail underneath it. The status tabs stay — someone chasing one
   rejected file should not have to remember which submission it was in. */
/* ⚠ THE "Bulk Upload" TAB WAS REMOVED ON 2026-08-30, and the client's correction is worth keeping
   because it reads like a small one and is not: *"I notice you create a literal tab for Bulk Upload,
   but what I meant was under Submission tab put a bulk upload column beside batches to indicate if
   its bulk or batches upload."*

   A separate tab PARTITIONED the list — a run appeared on exactly one of the two, never both — so an
   uploader looking for something they filed had to know HOW it was filed before they could find it,
   which is the one thing they are least likely to remember. A column answers the same question
   without hiding anything: one list, sorted by date, with the method stated per row. */
/**
 * ⚠ NO-CACHE ON EVERY READ. The Requests page was found on 2026-08-30 rendering a CACHED list —
 * decided requests came back as pending with live Approve buttons, and clearing the browser cache
 * was the only thing that fixed it. This page reads the same request rows to decide whether to show
 * *"you already asked"* and to offer Cancel, so a stale answer here invites a second request for a
 * decision that has already been made, or a Cancel on something already carried out.
 *
 * A stale RESPONSE and a stale BUNDLE look almost identical from the outside; clearing the cache
 * separates them. Costing a few uncached requests is the cheaper side of that.
 */
const NO_CACHE = {
  Accept: "application/json;odata=nometadata",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
};

/* ⚠ `Archive` FILTERS ON A RECORD STATE, NOT A STATUS — see `archivedRowsOnly`. It sits after
   the three approval outcomes because it is not one of them: a file reaches it by ageing out,
   not by anybody deciding anything. Client, 2026-09-03: *"Add another tab call Archive so they
   can tell filter to archive files."* */
const TABS = ["Submissions", "All", "Pending", "Approved", "Rejected", "Archive", "Requests"];

const s: Record<string, React.CSSProperties> = {
  // Capped and centred like the other full-page screens, so a wide monitor does not stretch the
  // rows into unreadable ribbons.
  wrap:     { fontFamily: "'Segoe UI', sans-serif", color: "#1b1b1b", maxWidth: 1100, margin: "32px auto", padding: "0 24px 48px" },
  headRow:  { display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 8 },
  h2:       { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  headRefresh: {
    border: "1px solid #c7c7c7", background: "#fff", borderRadius: 4, cursor: "pointer",
    fontSize: 13, lineHeight: 1, padding: "3px 9px", color: "#0f6c3f",
  },
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
  // ── Requests ──
  /* Every key used must EXIST here - `s` is a `Record<string, CSSProperties>`, so a missing one
     yields `undefined` and renders unstyled with a green build. */
  /* The Upload column's two chips. Quiet on purpose — this states HOW a submission was made, which
     matters far less than its status, and a loud badge here would compete with the one that says
     whether the documents were approved. */
  // ⚠ `batchChip` REMOVED 2026-09-03 with the "Upload" column — Batch is the unmarked default now,
  // and only Bulk is tagged (`bulkChip`, rendered inline beside the folder path). Deleted rather
  // than parked: lint cannot flag an unused key on this object, so a dead style stays for ever.
  bulkChip:  { display: "inline-block", background: "#eef3fb", color: "#1d4f91", borderRadius: 999, padding: "2px 10px", fontSize: 11.5, fontWeight: 600 },
  askQuiet: { borderTop: "1px solid #edebe9", paddingTop: 10, marginTop: 12, fontSize: 12.5, color: "#605e5c" },
  // Background is the client's own exact value (2026-09-03 mockup): rgba(255, 225, 159, 1) = #FFE19F.
  pendingChip: { display: "inline-block", background: "rgba(255, 225, 159, 1)", border: "1px solid #f0d5a8", color: "#8a5a00", borderRadius: 999, padding: "2px 10px", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" },
  askPending: { background: "#fff9f0", border: "1px solid #f3e3c3", borderRadius: 4, padding: "10px 12px", fontSize: 12.5, color: "#6b4a12", lineHeight: 1.5, margin: "12px 0" },
  askDecided: { background: "#f3f9f5", border: "1px solid #cde8d8", borderRadius: 4, padding: "10px 12px", fontSize: 12.5, color: "#1c5334", lineHeight: 1.5, margin: "12px 0" },
  rowBtn: { padding: "3px 9px", marginLeft: 6, fontSize: 11.5, background: "#fff", color: "#0f6c3f", border: "1px solid #cde8d8", borderRadius: 3, cursor: "pointer" },
  rowBtnOff: { padding: "3px 9px", marginLeft: 6, fontSize: 11.5, background: "#f3f2f1", color: "#a19f9d", border: "1px solid #edebe9", borderRadius: 3, cursor: "default" },
  /* The recipient chips. `wordBreak` because an address can be longer than the dialog is wide, and a
     chip that cannot wrap would force the whole box to scroll sideways. */
  recipChip: { display: "inline-flex", alignItems: "center", gap: 4, maxWidth: "100%", padding: "2px 4px 2px 8px", borderRadius: 12, fontSize: 12, background: "#eff6f2", color: "#0b4f2e", border: "1px solid #cfe3d8", wordBreak: "break-all" },
  recipX: { border: "none", background: "transparent", color: "#0b4f2e", fontSize: 15, lineHeight: 1, padding: "0 4px", cursor: "pointer", fontFamily: "inherit" },
  askedTag: { display: "inline-block", marginLeft: 8, padding: "1px 6px", borderRadius: 8, fontSize: 10.5, fontWeight: 600, background: "#fdf5e6", color: "#8a5a00", border: "1px solid #f0dcb4" },
  // Client, 2026-09-03: "add auto scroll for All tabs, pending, approved, Rejected and request, it
  // is sooo long of a list." Same 60vh pattern this codebase already uses elsewhere (the abbreviation
  // editor, Group Management's group list) — wraps the table only, so headers stay put while the
  // rows scroll, and nothing absolutely positioned lives inside any of these three tables (unlike
  // the cases in this codebase where a 60vh cap has clipped a popover — checked, none apply here).
  scroller: { maxHeight: "60vh", overflowY: "auto" },
  askBar:   { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 16 },
  askBtn:   { padding: "6px 14px", fontSize: 12.5, fontFamily: "inherit", border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  askOff:   { padding: "6px 14px", fontSize: 12.5, fontFamily: "inherit", border: "1px solid #e6e6e6", borderRadius: 4, background: "#f4f4f4", color: "#9a9a9a", cursor: "not-allowed" },
  okBox:    { fontSize: 13, padding: "12px 14px", borderRadius: 6, border: "1px solid #b7dcc4", background: "#f3faf5", color: "#1c4d33", lineHeight: 1.55, marginBottom: 16 },
  warnBox:  { fontSize: 12.5, padding: "12px 14px", borderRadius: 6, border: "1px solid #f2c9a0", background: "#fff8f0", color: "#8a4b00", lineHeight: 1.55, marginBottom: 12 },
  modalBg:  { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 },
  modal:    { background: "#fff", borderRadius: 8, padding: 20, width: "min(560px, 94vw)", maxHeight: "86vh", overflowY: "auto", fontSize: 13 },
  label:    { display: "block", marginTop: 12, marginBottom: 4, fontSize: 12, fontWeight: 600, color: "#3b3a39" },
  field:    { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13, fontFamily: "inherit", border: "1px solid #c7c7c7", borderRadius: 4 },
  hint:     { fontSize: 12, color: "#605e5c", marginTop: 6, lineHeight: 1.5 },
  // Amber rather than red: this is a CLASSIFICATION, not a problem with the row.
  /* ⚠ VISUALLY DISTINCT FROM `hcTag`, deliberately. They can appear on the SAME row — an archived
     Highly Confidential document — and two amber pills side by side read as one badge that wrapped.
     Slate rather than amber: archived is a state, not a warning. */
  arcTag:   { display: "inline-block", marginLeft: 8, padding: "1px 7px", borderRadius: 10, fontSize: 11, fontWeight: 700, letterSpacing: 0.3, background: "#eef1f5", color: "#3b4a5a", border: "1px solid #c8d2de", verticalAlign: "middle" },
  hcTag:    { display: "inline-block", marginLeft: 8, padding: "1px 7px", borderRadius: 10, fontSize: 11, fontWeight: 700, letterSpacing: 0.3, background: "#fff4ce", color: "#8a5700", border: "1px solid #f2d18b", verticalAlign: "middle" },
  /* ── The submission RECORD's own states (2026-08-27) ──────────────────────────
     A row for a file that is no longer here. GREY, never red: a deleted document is a fact about the
     past, and colouring it as a failure would read as something being wrong with this page.
     `goneName` is a span, not a button — there is nothing to open. */
  goneRow:   { opacity: 0.62 },
  goneName:  { fontSize: 13, color: "#605e5c", fontStyle: "italic" },
  goneBadge: { display: "inline-block", padding: "2px 8px", borderRadius: 10, fontSize: 11.5, fontWeight: 600, background: "#f3f2f1", color: "#605e5c", border: "1px solid #d6d4d2" },
  /* "Not checked" is a THIRD answer and must not read as either of the others: the libraries could
     not all be read, so this row's file may be perfectly fine. Empty ≠ unknown, again. */
  /* ⚠ BLUE, NOT GREY AND NOT AMBER. Grey is `Deleted` (a loss) and amber is `Not checked` (a
     doubt); a REPLACED file is neither — somebody deliberately filed a newer version. Sharing a
     colour with either would put the client's whole reason for asking for this state back into the
     one it was meant to be told apart from. */
  cancelBadge: { display: "inline-block", padding: "2px 8px", borderRadius: 10, fontSize: 11.5, fontWeight: 600, background: "#eff6fc", color: "#005a9e", border: "1px solid #c7e0f4" },
  unsureBadge: { display: "inline-block", padding: "2px 8px", borderRadius: 10, fontSize: 11.5, fontWeight: 600, background: "#fdf5e6", color: "#8a5a00", border: "1px solid #f0dcb4" },
  /* ⚠ SLATE, AND DELIBERATELY THE SAME PALETTE AS `arcTag` — the "Archived" chip this page already
     uses for a row LIVING in the archive. The two say the same word about the same fact and would be
     read as two different things if they were coloured differently. It must not share GREY with
     `Deleted`: an archived document is retained, not lost, and telling those apart is the entire
     reason the state exists (client, 2026-09-03). */
  archivedBadge: { display: "inline-block", padding: "2px 8px", borderRadius: 10, fontSize: 11.5, fontWeight: 600, background: "#eef1f5", color: "#3b4a5a", border: "1px solid #c8d2de" },
  goneNote:  { margin: "0 0 10px", fontSize: 12.5, color: "#605e5c" },
  /* ⚠ ADDED WITH THE SUBMISSIONS VIEW (2026-08-22). `s` is a Record<string, CSSProperties>, so a key
     that does not exist yields `undefined` and React simply renders the element unstyled — the build
     stays green and the page looks broken. Any new style referenced above must be declared here. */
  backBtn:  { background: "none", border: "none", padding: 0, marginBottom: 12, fontSize: 13, fontFamily: "inherit", color: "#0b6a3a", cursor: "pointer" },
  card:     { border: "1px solid #e1dfdd", borderRadius: 6, padding: "12px 14px", marginBottom: 16, background: "#faf9f8" },
  cardHead: { fontSize: 13, fontWeight: 600, color: "#3b3a39", margin: "0 0 6px" },
  refLine:  { fontSize: 12, color: "#605e5c", marginTop: 6, fontFamily: "Consolas, monospace" },
  /* The read-only upload form (2026-08-22). Two columns so a batch's five or six destination fields
     read as a filled-in form rather than a long ladder; one column below 720px via `minmax`. */
  grid2:    { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "10px 24px" },
  fileCard: { border: "1px solid #e1dfdd", borderRadius: 6, padding: "12px 14px", marginBottom: 12, background: "#fff" },
  fileHead: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 10, paddingBottom: 8, borderBottom: "1px solid #f3f2f1" },
  openBtn:  { marginLeft: "auto", background: "none", border: "1px solid #c8c6c4", borderRadius: 4, padding: "3px 10px", fontSize: 12, fontFamily: "inherit", color: "#0f6cbd", cursor: "pointer" },  // Capped and scrolling: the people picker returns up to 10, and an unbounded list pushes the
  // permission dropdown and Send button below the fold inside a dialog.
  pickList: { marginTop: 6, maxHeight: 168, overflowY: "auto", border: "1px solid #e1dfdd", borderRadius: 4 },
  pickRow:  { display: "block", width: "100%", textAlign: "left", padding: "7px 10px", fontSize: 13, fontFamily: "inherit", background: "#fff", border: "none", borderBottom: "1px solid #f3f2f1", cursor: "pointer" },
  primary:  { padding: "7px 16px", fontSize: 12.5, fontFamily: "inherit", background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
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

/**
 * What the request form needs to know before it can offer anything.
 *
 * `known` is separate from the values, because every field's "not yet read" state is identical to a
 * legitimate one: no approver units, no extra domains, external sharing off. Offering a request form
 * built from a failed read would refuse valid recipients while naming a setting nobody had touched.
 */
interface RequestPolicy {
  known: boolean;
  /**
   * FAILS CLOSED, unlike almost everything else in this codebase.
   *
   * Elsewhere the cost of a failed read is a form out of service, so `unknown` means "offer
   * everything". Here it is a document leaving the organisation on the strength of a setting nobody
   * could confirm — the same reason `canOfferFolderDelete` fails closed. Spec §5.3.
   */
  allowExternal: boolean;
  tenantDomains: string[];
  /** Every unit term GUID carrying an `APR` mapping — used to route the request, not to authorise it. */
  approverUnits: string[];
  /**
   * Term GUIDs where THIS VIEWER can already do the thing themselves, so asking is pointless.
   *
   * Client, 2026-08-30: *"it doesn't make sense for HOU or HOD to approve their own request when
   * they can delete or share, only PIC will need to raise a request because they can't delete or
   * share without someone's approval."*
   *
   * ⚠ MATCHED AGAINST THE DOCUMENT'S WHOLE TIER CHAIN, not just its unit — which is what makes
   * this correct for a Head of Department with NO extra requests. Their `DEL`/`SHARE` row sits on
   * the DEPARTMENT term and fans down to every unit under it; the department term is already one of
   * the document's own tiers, so a chain match expresses the fan-out exactly. A unit-only compare
   * would have needed a term-store expansion per department, and would still have been wrong for a
   * segment-tier grant.
   *
   * Split by stage because the RIGHTS are: deleting an APPROVED document needs `DEL`, deleting one
   * still awaiting approval needs `DELS` — different roles on different libraries, and a Head of
   * Unit holds both while a Head of Department holds only the first.
   */
  directDelete: string[];
  directDeleteStaging: string[];
  directShare: string[];
  /** False only when the list answered 404. A different failure leaves this true and fails at the write. */
  listExists: boolean;
}

/** The latest request THIS PERSON raised against one file. */
interface MyRequest {
  /** The list item id — needed to cancel, which is a MERGE on that item. */
  id: number;
  type: "Deletion" | "Share";
  status: RequestStatus;
  at: string;
  decidedBy: string;
  note: string;
  /** Carried so the Requests tab can name the file without re-reading either library. */
  itemName: string;
  /* Stored even though the read is already filtered to this person. Passing the signed-in address in
     its place would make the ownership check circular — it would answer yes by construction, and
     stop being a check at all. */
  requestedBy: string;
}

const BLANK_POLICY: RequestPolicy = {
  known: false, allowExternal: false, tenantDomains: [], approverUnits: [], listExists: true,
  /* ⚠ EMPTY MEANS "OFFER THE REQUEST", and that is the safe direction. A failed read here must
     never HIDE the buttons: a PIC who cannot ask is stuck with no route at all, whereas a Head of
     Unit shown a button they did not need raises one request their own screen then decides. Cost of
     being wrong: a redundant request. Cost of the other way: somebody blocked. */
  directDelete: [], directDeleteStaging: [], directShare: [],
};

/* JSON light with NO `__metadata`, and BOTH header halves saying nometadata — plus `odata-version: ""`
   because SPFx injects 4.0, under which SharePoint cannot infer the entity set for a JSON-light entry
   and a row POST 400s. All learned on the audit log; the same headers the Requests page uses. */
const WRITE_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
  "odata-version": "",
};

export default function MySubmissions({ context }: IMySubmissionsProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  // Origin with no /sites/… — a server-relative path already carries the site, so previewTarget
  // needs the bare host to build an absolute file URL.
  const tenantRoot = siteUrl.replace(/^(https?:\/\/[^/]+).*$/, "$1");
  const [tab, setTab] = useState("Submissions");
  /* `MergedRow`, not `Submission`: a row can now be a RECORD of a file that no longer exists.
     `MergedRow extends Submission`, so every helper on this page still takes them unchanged, and a
     row with no `recordState` is an ordinary live document exactly as before. */
  const [rows, setRows] = useState<MergedRow[] | undefined>(undefined);
  /* How many recorded uploads could not be found, and how many could not be CHECKED. Shown as a
     quiet line rather than an error: a deleted document is a fact about the past, not a failure of
     this page. `undefined` until the first load settles. */
  const [recordNote, setRecordNote] =
    useState<{ deleted: number; cancelled: number; archived: number; unknown: number } | undefined>(undefined);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [commentsMissing, setCommentsMissing] = useState(false);
  /* True when a library lacked the reference columns, so the page can say WHY nothing is
     grouped rather than looking broken. Not an error — grouping is an addition. */
  const [refsMissing, setRefsMissing] = useState(false);
  /* ⚠ A REF, NOT STATE, AND THAT IS LOAD-BEARING. `load` sets this while reading the libraries and
     must read it back in the SAME pass to decide whether a missing file may be called deleted; a
     `useState` value is not visible to the closure already running (the stale-closure trap that cost
     reconciliation every permission on the site in 1.0.237.0). True means at least one library could
     not return `SubmissionFileId`, so no record can be safely judged. */
  const stampMissingRef = React.useRef(false);
  /* Where the reader is in submission → batch → files. Held as the reference strings rather than
     indices: the list is reloaded on focus and by the poll, and an index would silently point at a
     different submission after anything was added. */
  const [openSubmission, setOpenSubmission] = useState<string | undefined>(undefined);
  const [openBatch, setOpenBatch] = useState<string | undefined>(undefined);
  /**
   * The Documents library's real URL SEGMENT, resolved rather than assumed.
   *
   * Its title is `Documents` and its URL is `/Shared Documents/` — the two differ, exactly as they
   * do for the approval library (gotcha #12). Hardcoding one string for both put "Shared Documents"
   * at the front of every approved file's folder trail. Defaults to the title so a failed read
   * degrades to the old behaviour rather than to nothing.
   */
  const [docsSegment, setDocsSegment] = useState(DOCUMENTS_URL_SEGMENT);
  /** The row being examined. `undefined` = the list. */
  const [open, setOpen] = useState<Submission | undefined>(undefined);
  /** Per-item metadata labels for the open row. `undefined` while in flight. */
  const [fieldText, setFieldText] = useState<Record<string, string> | undefined>(undefined);
  /**
   * Per-item metadata for EVERY file in the batch being read, keyed by `submissionKey`.
   *
   * ⚠ ONE REQUEST PER FILE, and that is unavoidable: `FieldValuesAsText` is a per-ITEM endpoint.
   * It is why the flat list shows no metadata at all — hundreds of rows would be hundreds of
   * requests. A batch is a handful of files and is opened deliberately, so the cost is bounded and
   * asked for. `undefined` for the whole map means still reading; `{}` for one file means that
   * file's details could not be read, which the panel says rather than showing an empty card.
   */
  const [batchText, setBatchText] = useState<Record<string, Record<string, string>> | undefined>(undefined);
  /* Which batch `batchText` describes — "this batch, with these item ids". Declared with the state it
     guards, and NOT next to the effect that uses it, because `loadBatchText` reads it too and
     `no-use-before-define` is on. See that effect for why a signature rather than a boolean. */
  const batchSigRef = useRef<string>("");

  /* ── Requests ──────────────────────────────────────────────────────────────
     An uploader cannot delete or share in Documents, so they ask and a Head of Unit decides.
     Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md */
  const [policy, setPolicy] = useState<RequestPolicy>(BLANK_POLICY);
  /* Keyed by lower-cased UniqueId. Empty means "none raised, or the list could not be read" —
     which is safe HERE and only here: the worst outcome is a badge that does not appear, and
     the buttons stay usable either way. */
  const [myRequests, setMyRequests] = useState<Record<string, MyRequest>>({});
  /* The same rows, unindexed and in the order raised — the Requests tab lists every request, while
     `myRequests` answers "what is outstanding on THIS file". One read feeds both. */
  const [myRequestList, setMyRequestList] = useState<MyRequest[]>([]);
  const [cancelling, setCancelling] = useState<number | undefined>(undefined);
  /* ⚠ SEPARATE FROM `loadError`, and that separation is the point. A cancel that is refused because
     the approver got there first is a NORMAL outcome, not a failure to read the page — and routing
     it through the load-error banner produced "Could not read your submissions … please try again
     rather than uploading again" over a page that had loaded perfectly and a request that had been
     decided correctly. Reported on site 2026-08-21, the same afternoon the guard was added. */
  const [requestNotice, setRequestNotice] = useState<string | undefined>(undefined);
  /* WHICH FILE the request dialog is about. Separate from `open`, because the dialog is now reachable
     from a row in the list as well as from the detail view — the client asked for the buttons "beside
     the uploaded for each file so its easier ... instead of going into each file manually". */
  const [askRow, setAskRow] = useState<Submission | undefined>(undefined);
  const [asking, setAsking] = useState<RequestType | undefined>(undefined);
  const [reason, setReason] = useState("");
  const [shareWith, setShareWith] = useState("");
  /* Recipient lookup (2026-08-21, client: *"ensure that the email input is like the Group management
     where it will pre show the email"*). A typo in the address box is silent — the share goes to
     nobody, or to the wrong person, and the requester never finds out. Suggestions remove the typo
     for anyone already known to the site.
     ⚠ IT CANNOT REPLACE THE FREE-TEXT BOX. The main case for a share request is somebody who is NOT
     on the site yet — that is what the whole flow exists for — so the textarea stays and this only
     APPENDS to it. `searchTenantPeople` is already configured `AllowEmailAddresses: true` /
     `AllowOnlyEmailAddresses: false`, so it resolves known people and still accepts an address that
     matches nobody. */
  const [peopleQ, setPeopleQ] = useState("");
  const [peopleHits, setPeopleHits] = useState<PersonPick[]>([]);
  const [peopleBusy, setPeopleBusy] = useState(false);
  const [permission, setPermission] = useState<SharePermission>("View");
  const [expiresAt, setExpiresAt] = useState("");
  const [problems, setProblems] = useState<string[]>([]);
  const [sending, setSending] = useState(false);
  const [sent, setSent] = useState<string | undefined>(undefined);

  const currentUserId = async (): Promise<number> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/currentuser?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: NO_CACHE },
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
    // `File/Length` needs the $expand — a bare `File/Length` in a $select is rejected. Modified is
    // read alongside Created because on this page the gap between them is the story: created is when
    // the uploader sent it, modified is when the approver acted.
    // UniqueId is what a deletion or share request is raised against — a GUID that survives the
    // rename or move that a path does not. It is a built-in field on every list item, so it costs
    // nothing and is safe on both libraries; the minimal fallback below drops it, and the request
    // buttons check for it rather than assuming.
    const common = "Id,FileLeafRef,FileRef,UniqueId,Created,Modified,File/Length";
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
        `${base}?$select=${select}&$expand=File&$filter=AuthorId eq ${userId} and FSObjType eq 0` +
          "&$top=2000&$orderby=Created desc",
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );

    /* ⚠ THE REFERENCE COLUMNS ARE ASKED FOR FIRST AND DROPPED ON FAILURE (2026-08-22).
       `SubmissionId`/`BatchId` are created by reconciliation, so a site that has not run it since
       does not have them — and one unknown name in a `$select` fails the WHOLE request with 400
       (gotcha #11). Asking unconditionally would empty this page on an otherwise healthy site.
       Dropping them costs the grouping and nothing else: `groupSubmissions` then presents every file
       as its own single-file submission, which is exactly the pre-2026-08-22 behaviour. */
    /* ⚠⚠ AND `SubmissionFileId` IS ASKED FOR AHEAD OF THEM, WITH ITS OWN RUNG ON THE LADDER.
       It is the key the submission record joins on, so a live row WITHOUT it cannot be matched to its
       record — and an unmatched record is judged DELETED. Folding it into the same rung as the pair
       would mean a library missing only the newer column loses the grouping too; worse, dropping it
       silently would make this page report every recorded upload as destroyed. Hence three rungs:
       all three, then the pair, then none — and `stampMissingRef` so the merge KNOWS the stamp was
       unavailable and answers "not checked" instead of "deleted". */
    const withStamp = `${withStatus},SubmissionId,BatchId,SubmissionFileId`;
    const withRefs = `${withStatus},SubmissionId,BatchId`;
    let res = approvalLibrary
      ? await get(`${withStamp},OData__ModerationComments`)
      : await get(withStamp);
    if (!res.ok) {
      // The stamp column is absent — reconciliation has not run on this library since 2026-08-27.
      stampMissingRef.current = true;
      res = approvalLibrary
        ? await get(`${withRefs},OData__ModerationComments`)
        : await get(withRefs);
    }
    if (!res.ok) {
      setRefsMissing(true);
      res = approvalLibrary
        ? await get(`${withStatus},OData__ModerationComments`)
        : await get(withStatus);
    }
    if (!res.ok && approvalLibrary) {
      // The comment is optional, so a failure earns a retry without it before giving up.
      setCommentsMissing(true);
      res = await get(withStatus);
    }
    if (!res.ok) {
      // LAST RESORT: drop the two decorative fields — file size and Modified — and read the set the
      // page cannot do without. They were added on 2026-08-14 for a richer detail panel, and a
      // panel row must never be the reason an uploader is told they have no files. `File/Length`
      // in particular needs the $expand, and a $select naming anything unavailable fails the WHOLE
      // request with a 400 rather than omitting a key (gotcha #11).
      const minimal = approvalLibrary
        ? "Id,FileLeafRef,FileRef,Created,OData__ModerationStatus"
        : "Id,FileLeafRef,FileRef,Created";
      res = await context.spHttpClient.get(
        `${base}?$select=${minimal}&$filter=AuthorId eq ${userId} and FSObjType eq 0` +
          "&$top=2000&$orderby=Created desc",
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
    }
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`${listTitle}: HTTP ${res.status} ${body.slice(0, 140)}`);
    }

    const data = await res.json();
    return ((data.value ?? []) as RawRow[]).map((r) => {
      const createdRaw = textOf(r.Created);
      const modifiedRaw = textOf(r.Modified);
      // The expanded File, when the $expand survived. `textOf` handles every shape SharePoint can
      // return for a field, but Length sits one level down, so it is reached explicitly.
      const file = r.File as { Length?: unknown } | undefined;
      return {
        itemId: Number(r.Id),
        library: urlSegment,
        name: pickField(r, "FileLeafRef"),
        fileRef: pickField(r, "FileRef"),
        uniqueId: pickField(r, "UniqueId") || undefined,
        submissionId: pickField(r, "SubmissionId") || undefined,
        batchId: pickField(r, "BatchId") || undefined,
        // The record's join key. Absent on every file uploaded before 2026-08-27, and on any library
        // whose read had to fall back — see `stampMissingRef`.
        submissionFileId: pickField(r, "SubmissionFileId") || undefined,
        // Documents rows carry no status: they are there because they were approved.
        status: approvalLibrary ? statusToDecision(Number(r.OData__ModerationStatus)) : "Approved",
        created: createdRaw.length > 0 ? new Date(createdRaw) : undefined,
        comment: pickField(r, "OData__ModerationComments", "OData__x005f_ModerationComments"),
        size: textOf(file?.Length) || undefined,
        modified: modifiedRaw.length > 0 ? new Date(modifiedRaw) : undefined,
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
        { headers: NO_CACHE },
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
  const fetchFieldText = async (row: Submission): Promise<Record<string, string>> => {
    /* Which library this row came from, by its URL segment.
       A two-way guess (`docsSegment ? Documents : approval`) sent every Highly Confidential row to
       the NORMAL approval library, which answers 404 for an item id it does not have — so the panel
       showed "no details were recorded" for exactly the documents whose metadata matters most.

       ⚠ THE ARCHIVE PAIR WAS MISSING FROM THIS LIST TOO (found 2026-09-03, "how come no details
       showing?" — an archived file's row fell through every branch to the plain `libraryTitle()`
       fallback, 404'd against the wrong list, and silently showed "no details were recorded" for
       the one case where a file legitimately IS somewhere other than the two obvious libraries.
       `arcLibs`/`isArchivedRow` are already used elsewhere in this file to TAG a row as archived —
       this is the same fact, finally used to resolve where to actually read it from. */
    const hc = cachedHcLibraries();
    const arc = cachedArchiveLibraries();
    const listTitle =
      row.library === docsSegment ? documentsTitle()
      : hc && row.library === hc.approval.urlSegment ? hc.approval.title
      : hc && row.library === hc.documents.urlSegment ? hc.documents.title
      : arc && row.library === arc.normal.urlSegment ? arc.normal.title
      : arc && arc.hc && row.library === arc.hc.urlSegment ? arc.hc.title
      : libraryTitle();
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${row.itemId})/FieldValuesAsText`,
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
      // {} rather than undefined on failure: the detail view still shows the file and its status,
      // and says the details could not be read. Labels are never worth blocking the preview for.
      if (!res.ok) return {};
      const text = (await res.json()) as Record<string, string>;

      /* ⚠ DOCUMENT DATE IS RE-READ RAW, AND THE REASON IS THAT IT MUST NOT BE PARSED BACK.
         `FieldValuesAsText` hands back a string SharePoint has ALREADY formatted in the site's
         locale — `8/17/2026 12:00 AM`. The client asked for `17 Aug 2026`, and the obvious route,
         parsing that string, is the exact trap gotcha #1 records: this site is US-locale, so
         `8/9/2026` is 9 August or 9 September depending on an assumption nothing on the page can
         check, and a date silently off by a month in a document record is worse than an ugly one.

         So the raw `Edm.DateTime` is fetched instead — unambiguous ISO — and formatted by the same
         `formatSubmittedOn` every other date on this page uses, which is what keeps them agreeing.

         Costs ONE extra request per OPENED file. The panel already costs one, and it fires only
         when somebody opens a document. */
      try {
        const raw: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${row.itemId})?$select=DocumentDate`,
          SPHttpClient.configurations.v1,
          { headers: NO_CACHE },
        );
        if (raw.ok) {
          const iso = ((await raw.json()) as { DocumentDate?: string }).DocumentDate;
          const d = iso ? new Date(iso) : undefined;
          // Only overwrite on a date we could actually read. A blank column, an unparseable value
          // or a failed request all leave SharePoint's own string standing — uglier, never wrong.
          if (d && !isNaN(d.getTime())) text.DocumentDate = formatSubmittedOn(d);
        }
      } catch {
        /* keep the formatted string — a decoration must never cost the panel its details */
      }
      return text;
    } catch {
      return {};
    }
  };

  /** Fill `fieldText` for one row — the single-document detail view's loader. */
  const loadFieldText = async (row: Submission): Promise<void> => {
    setFieldText(undefined);
    setFieldText(await fetchFieldText(row));
  };

  /**
   * Fill `batchText` for every file in one batch.
   *
   * ⚠ `Promise.allSettled` is unavailable on this tsconfig (gotcha #3), so each read carries its own
   * catch INSIDE `Promise.all` — one unreadable file must never blank the whole batch. `fetchFieldText`
   * already resolves to `{}` rather than rejecting, which the per-file panel reports honestly.
   */
  const loadBatchText = async (files: MergedRow[], sig: string): Promise<void> => {
    setBatchText(undefined);
    const pairs = await Promise.all(
      files.map(async (f) => {
        /* ⚠ A GONE FILE IS NOT FETCHED. `FieldValuesAsText` is a per-ITEM endpoint and the item no
           longer exists, so this would be one guaranteed-failing request per deleted file — and the
           batch card reads its details from the record's own snapshot instead. */
        if (f.recordState) return { key: mergedKey(f), ft: {} as Record<string, string> };
        const ft = await fetchFieldText(f).catch(() => ({} as Record<string, string>));
        return { key: mergedKey(f), ft };
      }),
    );
    const map: Record<string, Record<string, string>> = {};
    for (const pair of pairs) map[pair.key] = pair.ft;
    /* ⚠ A SUPERSEDED READ MUST NOT WIN. Two loads can be in flight when a file routes mid-view (the
       live refresh re-reads the rows, the file's item id changes, and the batch is reloaded under a
       new signature). If the older one resolved last it would install a map keyed on item ids that no
       longer exist, and every card would read as unloaded for ever. */
    if (batchSigRef.current !== sig) return;
    setBatchText(map);
  };

  const openRow = (row: Submission): void => {
    setOpen(row);
    setAsking(undefined);
    setSent(undefined);
    setProblems([]);
    loadFieldText(row).catch(() => setFieldText({}));
  };

  /* ── The batch detail read is driven by STATE, not by the click that opened it ──
   *
   * ⚠ IT WAS DRIVEN BY THE CLICK, AND THERE ARE TWO WAYS INTO THAT VIEW — so the second one loaded
   * nothing and every file read "Reading the details…" for ever. Level 2's batch row called
   * `loadBatchText`; level 1's row does NOT, because a bulk run has one destination and opens the
   * batch directly (2026-08-28). Client, 2026-09-03: *"nothing is showing for bulk upload in my
   * submisison."* Third instance in this project of a new route into an existing path silently
   * starting from zero — the same shape as the member-add failure paths in `GroupMembersEditor`.
   *
   * ⚠ AND IT WENT STALE THE MOMENT A FILE WAS APPROVED. `batchText` is keyed by `mergedKey`, which
   * for a live row is `library#itemId` — and Auto-route COPIES an approved file into the approved-side
   * library and deletes the source, so it comes back with a different library and a different item id.
   * The map loaded while the file was pending then answers for a key nothing asks about any more.
   * Client: *"I upload normally through an upload form but when I open the submission I dont see the
   * details anymore after approved."* Keying the effect on the batch's own signature is what fixes
   * that: the keys change, so the signature changes, so it reloads.
   *
   * Deriving the batch here rather than in the render also means `openBatch` surviving a reload, or
   * being set by any future third route, cannot produce the stuck state again.
   */
  useEffect(() => {
    if (openSubmission === undefined || openBatch === undefined) return;
    const groups = groupSubmissions(rows ?? []);
    const keyFor = (g: SubmissionGroup): string =>
      g.reference || (g.files[0] ? submissionKey(g.files[0]) : "");
    const g = groups.filter((x) => keyFor(x) === openSubmission)[0];
    const b = g ? g.batches.filter((x) => (x.reference || "—") === openBatch)[0] : undefined;
    if (!b) return;
    // The signature is what identifies "this batch, as it stands now" — the files it holds and the
    // ids they hold. `join` on a stable separator: a mergedKey cannot contain one.
    const sig = openBatch + "::" + b.files.map(mergedKey).join("|");
    if (batchSigRef.current === sig) return;
    // Set BEFORE the await, so a re-render mid-flight does not start a second identical load.
    batchSigRef.current = sig;
    loadBatchText(b.files, sig).catch(() => setBatchText({}));
  }, [openSubmission, openBatch, rows]);

  /* ── Requests: what the form needs before it can offer anything ──────────── */

  /**
   * Three independent reads, none of which may take the page down.
   *
   * This whole block is decoration for the primary job — telling an uploader what happened to their
   * files — so every failure is caught individually and `known` simply stays false. The buttons then
   * say why rather than appearing and failing at the write.
   */
  const loadPolicy = async (): Promise<void> => {
    const next: RequestPolicy = { ...BLANK_POLICY, known: true };

    /* ⚠ THE CONFIG ROW WINS OUTRIGHT; the viewer's own domain is only a FALLBACK, applied AFTER
       the read below finds nothing. This used to seed unconditionally on the reasoning that the
       signed-in user "is signed in to this tenant, so it is not a guess" — **false for a guest**,
       and here it decides whether `allowExternalSharing` engages at all. A gmail guest sharing to a
       gmail address therefore walked straight past a gate that was switched off. Confirmed live
       2026-08-30. Same correction as the Requests page. */
    const me = (context.pageContext.user.email ?? "").toLowerCase();

    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Title,SettingValue&$filter=Title eq 'allowExternalSharing' or Title eq 'tenantDomains'&$top=20`,
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
      if (res.ok) {
        for (const r of ((await res.json()).value ?? []) as Array<{ Title?: string; SettingValue?: string }>) {
          const value = (r.SettingValue ?? "").trim();
          if (r.Title === "allowExternalSharing") {
            // Only an explicit yes turns it on. Anything else — blank, "no", a typo — stays off.
            next.allowExternal = /^(true|yes|on|1)$/i.test(value);
          } else if (r.Title === "tenantDomains") {
            for (const d of value.split(/[,;\s]+/)) {
              const clean = d.trim().toLowerCase();
              if (clean && next.tenantDomains.indexOf(clean) === -1) next.tenantDomains.push(clean);
            }
          }
        }
      }
    } catch {
      /* internal-only, and the form says so */
    }

    /* No configured domains means nothing better is available, so the viewer's own is the best
       guess left — which keeps an unconfigured site behaving exactly as it did. With a row set it
       is an ANSWER, and a guest's personal domain must not widen it. */
    if (next.tenantDomains.length === 0) {
      const at = me.lastIndexOf("@");
      if (at > -1) next.tenantDomains.push(me.slice(at + 1));
    }

    /* WHICH SITE GROUPS IS THIS PERSON IN? Needed only to work out what they can already do
       themselves — an empty or failed answer leaves every request button offered, which is the safe
       direction (see `BLANK_POLICY`). */
    const myGroupIds: number[] = [];
    try {
      const gr: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser/groups?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
      if (gr.ok) {
        for (const g of ((await gr.json()).value ?? []) as Array<{ Id?: number }>) {
          if (typeof g.Id === "number") myGroupIds.push(g.Id);
        }
      }
    } catch {
      /* every button stays offered */
    }

    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items` +
          `?$select=GroupId,Role,UnitTermGuid&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
      if (res.ok) {
        const rows = ((await res.json()).value ?? []) as Array<{
          GroupId?: number; Role?: string; UnitTermGuid?: string;
        }>;
        for (const r of rows) {
          // normalizeRoleValue, NOT a raw toUpperCase — see the same match in Requests.tsx. The
          // Group Map's Role is often the LONG form ("APPROVER"), and a raw compare skips such a
          // row, which would file the request against a unit nobody is recorded as approving for.
          const role = normalizeRoleValue(r.Role ?? "");
          const guid = (r.UnitTermGuid ?? "").trim();
          if (guid.length === 0) continue;

          // APR OR APRHC since 2026-08-24 — see the same match in Requests.tsx. An HC unit's approver
          // holds APRHC and no APR row at all, so matching APR alone routes its requests to nobody.
          // Site-wide, because this ROUTES a request rather than authorising one.
          if (role === "APR" || role === "APRHC") {
            if (next.approverUnits.indexOf(guid) === -1) next.approverUnits.push(guid);
          }

          /* ⚠ THE REST IS ABOUT THIS VIEWER ONLY, hence the GroupId test. Read the rows of groups
             they are NOT in and a PIC would lose their request buttons because somebody ELSE holds
             delete on their unit — which is precisely the person they are supposed to be asking. */
          if (myGroupIds.indexOf(r.GroupId ?? -1) === -1) continue;
          const add = (list: string[]): void => {
            if (list.indexOf(guid) === -1) list.push(guid);
          };
          // DEL / DELHC — delete an APPROVED document. Held by `hou` and by `hod` at department tier.
          if (role === "DEL" || role === "DELHC") add(next.directDelete);
          // DELS / DELSHC — delete one still awaiting approval, in the approval library. `hou` only.
          if (role === "DELS" || role === "DELSHC") add(next.directDeleteStaging);
          if (role === "SHARE" || role === "SHAREHC") add(next.directShare);
        }
      }
    } catch {
      /* the request still routes on the document's own deepest tier */
    }

    /**
     * The same probe that used to ask only whether the list exists, widened to fetch THIS PERSON'S
     * OWN requests.
     *
     * Added 2026-08-20 as the answer to "did my request go through?". Without it an uploader who
     * asked yesterday sees a page identical to one who never asked, and the natural response is to
     * ask again — so the Head of Unit gets the same file twice and has to work out whether that was
     * a mistake or a reminder.
     *
     * Filtered `RequestedBy eq me`, which costs nothing extra and keeps the page's one guarantee
     * intact: everything here is yours.
     */
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')/items` +
          `?$select=Id,ItemUniqueId,ItemName,RequestType,Status,RequestedBy,RequestedAt,DecidedBy,DecisionNote` +
          `&$filter=RequestedBy eq '${encodeURIComponent((context.pageContext.user.email ?? "").toLowerCase())}'` +
          `&$orderby=Id desc&$top=500`,
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
      // ONLY a 404 means absent. Any other failure leaves it true, so the uploader is told the write
      // failed rather than being sent to an administrator over what may be a transient error.
      if (res.status === 404) next.listExists = false;
      else if (res.ok) {
        const rows = ((await res.json()).value ?? []) as Array<Record<string, string>>;
        const byFile: Record<string, MyRequest> = {};
        const all: MyRequest[] = [];
        for (const r of rows) {
          // `Id desc`, so the FIRST row seen for a file is the most recent — and a file may have
          // been asked about more than once (asked, rejected, asked again). Only the latest is worth
          // showing; the rest are history the Requests page carries.
          const uid = (r.ItemUniqueId ?? "").replace(/[{}]/g, "").trim().toLowerCase();
          const one: MyRequest = {
            id: Number(r.Id ?? 0),
            type: r.RequestType === "Share" ? "Share" : "Deletion",
            // Parsed, not cast: a blind `as RequestStatus` lets a garbage value into the union and
            // out onto the screen. Same one definition the Requests page reads.
            status: parseRequestStatus(r.Status),
            at: r.RequestedAt ?? "",
            decidedBy: r.DecidedBy ?? "",
            note: r.DecisionNote ?? "",
            itemName: r.ItemName ?? "",
            requestedBy: r.RequestedBy ?? "",
          };
          // EVERY row goes in the list; only the FIRST per file goes in the index, since `Id desc`
          // makes that the most recent and the file banner is about what is outstanding now.
          all.push(one);
          if (uid && !byFile[uid]) byFile[uid] = one;
        }
        setMyRequests(byFile);
        setMyRequestList(all);
      }
    } catch {
      /* leave it true — see above. An unreadable list shows no request badges, never a wrong one. */
    }

    setPolicy(next);
  };

  /**
   * Withdraw a request that has not been decided yet.
   *
   * A MERGE on the row rather than a delete: the audit trail is the point of this list, and a
   * withdrawn request is a thing that happened. `Cancelled` is safe to introduce because `Status` is
   * a TEXT column — a value absent from a Choice column's choices fails the whole write, silently.
   *
   * `canCancel` is re-checked here as well as in the render. The button is only shown for the
   * requester's own pending rows, but the rule that matters is the one next to the write.
   */
  const cancelRequest = async (req: MyRequest): Promise<void> => {
    const me = (context.pageContext.user.email ?? "").toLowerCase();
    // Built explicitly rather than cast from MyRequest: the two shapes only overlap by accident, and
    // a cast would silently keep compiling if either drifted.
    const asRow: RequestRow = {
      type: req.type,
      status: req.status,
      itemUniqueId: "",
      itemName: req.itemName,
      segment: "",
      unit: "",
      requestedBy: req.requestedBy,
      requestedAt: req.at,
      reason: "",
    };
    if (!canCancel(asRow, me)) return;
    setRequestNotice(undefined);
    setCancelling(req.id);
    try {
      /* ⚠ RE-READ THE ROW BEFORE WRITING. `req.status` is what this page read at MOUNT, and the
         approver decides in their own session — so a page left open across an approval still shows
         Cancel, `canCancel` still passes on the stale value, and the MERGE (`IF-MATCH: "*"`, no
         concurrency check) would write **Cancelled over an Approved decision on a file already in
         the recycle bin**. The record would then say the request was withdrawn before any decision,
         on the one list whose entire purpose is being a record. Found on site 2026-08-21.
         Same shape as the stale-chain guard (gotcha 10b) and the upload-pause re-check: the state
         that matters is the one at WRITE time, never at render time.
         FAILS CLOSED, against this codebase's usual direction. Everywhere else an unreadable check
         costs a form for a minute; here it would overwrite a decision that has already been carried
         out, and the user can simply try again. */
      const listBase = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')`;
      const fresh = await context.spHttpClient.get(
        `${listBase}/items(${req.id})?$select=Status,DecidedBy`,
        SPHttpClient.configurations.v1,
        { headers: NO_CACHE },
      );
      if (!fresh.ok) {
        setCancelling(undefined);
        setRequestNotice(`Could not check whether this request has already been decided (HTTP ${fresh.status}). Nothing was changed — refresh and try again.`);
        return;
      }
      const now = (await fresh.json()) as { Status?: string; DecidedBy?: string };
      if ((now.Status ?? "").trim() !== "Pending") {
        setCancelling(undefined);
        // Names the decider and the outcome: "it is too late" without saying WHAT happened leaves
        // the requester unsure whether their document still exists.
        setRequestNotice(
          `Too late to withdraw — your approver already ${(now.Status ?? "decided").toLowerCase()} this request` +
            (now.DecidedBy ? ` (${now.DecidedBy})` : "") +
            ". Nothing was changed.",
        );
        await loadPolicy();
        return;
      }
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')/items(${req.id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify({ Status: "Cancelled" }),
        },
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      writeAudit(context.spHttpClient, siteUrl, {
        event: req.type === "Deletion" ? EVENT.deletionRequested : EVENT.shareRequested,
        outcome: "Failed",
        source: "MySubmissions",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: me,
        itemName: req.itemName,
        summary: `${req.type} request withdrawn — ${req.itemName}`,
        // Outcome "Failed" because the request did NOT result in what it asked for. A row reading
        // Success beside a document that was never deleted stops someone looking.
        details: ["Withdrawn by the requester before any decision was made."],
      }).catch(() => undefined);
      await loadPolicy();
    } catch (e) {
      // Surfaced, never swallowed: an uploader who thinks they withdrew a request and did not will
      // be surprised when the file disappears.
      setLoadError(`Could not cancel that request: ${(e as Error).message}`);
    } finally {
      setCancelling(undefined);
    }
  };


  /**
   * Raise a request against the open document.
   *
   * WHERE IT GOES is derived from the document's own tier fields, not from the folder path — folder
   * names are abbreviations, and the durable key is the term GUID. The deepest tier an approver is
   * actually mapped to wins: a below-Unit tier such as SubUnit carries a Tid column exactly like a
   * permissioned one, so routing on "deepest" alone would file the request where nobody can see it.
   */
  const submitRequest = async (row: Submission, type: RequestType): Promise<void> => {
    const ft = fieldText ?? {};
    const where = documentUnit(ft);
    const routed = routeToApprover(where, policy.approverUnits);

    // A row with a moderation status is still in an approval library; a Documents row has none,
    // because everything there is approved by definition. Same discriminator the read uses, so the
    // two cannot disagree.
    const stage: RequestStage = row.status === "Approved" ? "approved" : "pending";

    const draft: RequestDraft = {
      type,
      stage,
      /* ⚠ DERIVED FROM THE LIBRARY, and it must reach `validateDraft` or the refusal never fires.
         Every role holds only Read on the archive, so an approved request would fail in the
         APPROVER'S own session — recorded `Failed`, days after the requester was told yes.

         Read from the cache HERE rather than from the render-scope `arcSegs`, which is declared
         further down the component: this function is defined above it. Same module cache, same
         answer — and a value captured at render time must never decide a request anyway (the
         StagingAccess `listBase` bug, 2026-08-14). */
      archived: (() => {
        const a = cachedArchiveLibraries();
        return isArchivedRow(row.library, a ? { normal: a.normal.urlSegment, hc: a.hc?.urlSegment } : undefined);
      })(),
      itemUniqueId: row.uniqueId ?? "",
      itemName: row.name,
      segment: where.segment,
      unit: routed ? routed.value : where.unit,
      reason,
      shareWith: type === "Share" ? shareWith : undefined,
      sharePermission: type === "Share" ? permission : undefined,
      expiresAt: type === "Share" ? expiresAt : undefined,
    };

    const found = validateDraft(draft, {
      allowExternal: policy.allowExternal,
      tenantDomains: policy.tenantDomains,
    });
    if (found.length > 0) {
      setProblems(found);
      return;
    }

    setSending(true);
    setProblems([]);
    try {
      const guid = routed ? routed.guid : where.unitTermGuid;
      const body: Record<string, string> = {
        Title: `${type} — ${row.name}`.slice(0, 255),
        RequestType: type,
        Status: "Pending",
        Stage: stage,
        ItemUniqueId: draft.itemUniqueId,
        ItemName: row.name,
        ItemUrl: row.fileRef,
        Segment: where.segment,
        Unit: draft.unit,
        UnitTermGuid: guid,
        RequestedBy: (context.pageContext.user.email ?? "").toLowerCase(),
        // ISO, never M/D/YYYY. This is a plain /items POST through the OData layer, which answers a
        // locale string with "Cannot convert a primitive value to the expected type 'Edm.DateTime'"
        // — gotcha #1's format belongs to validateUpdateListItem. $filter needs ISO too, so writes
        // and filters share one format and cannot be mismatched.
        RequestedAt: new Date().toISOString(),
        Reason: reason.trim(),
      };
      if (type === "Share") {
        body.ShareWith = parseRecipients(shareWith).join("; ");
        body.SharePermission = permission;
        // OMITTED when blank rather than sent as "". A DateTime column rejects an empty string, and
        // the whole write fails over an optional field.
        if (expiresAt.trim().length > 0) body.ExpiresAt = new Date(`${expiresAt.trim()}T12:00:00`).toISOString();
      }

      const itemsUrl =
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')/items`;
      const send = async (payload: Record<string, string>): Promise<SPHttpClientResponse> =>
        context.spHttpClient.post(
          itemsUrl, SPHttpClient.configurations.v1,
          { headers: WRITE_HEADERS, body: JSON.stringify(payload) },
        );

      let res: SPHttpClientResponse = await send(body);
      /**
       * RETRY WITHOUT `Stage` on a 400.
       *
       * The column is provisioned by the Requests page, so on a site whose list predates 2026-08-20
       * it does not exist — and one unknown field name fails the WHOLE write. Refusing here would
       * stop an uploader raising a request at all, which is worse than raising one the approver's
       * queue labels imprecisely: the deletion resolves by UniqueId and works either way. The
       * Requests page detects the same absence and offers to add the column.
       *
       * Never the reverse — the field is always ATTEMPTED first, so a correctly provisioned site
       * always records the stage.
       */
      if (res.status === 400 && body.Stage !== undefined) {
        const without: Record<string, string> = { ...body };
        delete without.Stage;
        res = await send(without);
      }
      if (!res.ok) {
        const text = await res.text().catch(() => "");
        throw new Error(
          res.status === 404
            ? "the requests list does not exist yet — ask an administrator to open the Requests page, which creates it"
            : `HTTP ${res.status} ${text.slice(0, 160)}`,
        );
      }

      /* ⚠ UPDATE `myRequests` LOCALLY, RIGHT NOW — client, 2026-09-02: the Delete/Share buttons and
         "asked" badges kept showing until a manual refresh, because `myRequests` is only ever
         populated by the mount-time read; nothing wrote back into it after a successful submit. Both
         the row list (line ~2411/2446) and the detail view (`openRequest`) read the SAME state, so
         one update here fixes both without a second definition of "is a request pending" to drift.
         `res.json()` on a POST to `/items` returns the created row, including its `Id` — read it
         rather than re-reading the whole list, which is the mount-time read's job, not a per-request
         one. A failure here costs nothing but the instant reflect: the next load (or focus-triggered
         live refresh) still catches up from the server, same as before this existed. */
      try {
        const created = await res.json();
        const createdId = Number(created?.Id ?? 0);
        const uid = (draft.itemUniqueId ?? "").replace(/[{}]/g, "").trim().toLowerCase();
        if (uid && createdId > 0) {
          const mine: MyRequest = {
            id: createdId,
            type,
            status: "Pending",
            at: body.RequestedAt,
            decidedBy: "",
            note: "",
            itemName: row.name,
            requestedBy: (context.pageContext.user.email ?? "").toLowerCase(),
          };
          setMyRequests((prev) => ({ ...prev, [uid]: mine }));
          setMyRequestList((prev) => [mine, ...prev]);
        }
      } catch {
        /* Same reasoning as everywhere else here: a failed read of our OWN write must never be why
           the request itself is treated as failed. The buttons just wait for the next load. */
      }

      writeAudit(context.spHttpClient, siteUrl, {
        event: type === "Deletion" ? EVENT.deletionRequested : EVENT.shareRequested,
        source: "MySubmissions",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: (context.pageContext.user.email ?? "").toLowerCase(),
        library: documentsTitle(),
        itemName: row.name,
        itemUniqueId: draft.itemUniqueId,
        itemPath: row.fileRef,
        segment: where.segment,
        unitPath: draft.unit,
        summary: `${type} requested — ${row.name}`,
        details: [
          `Reason: ${reason.trim()}`,
          type === "Share" ? `Recipients: ${parseRecipients(shareWith).join(", ")} (${permission})` : "",
          type === "Share" && expiresAt.trim() ? `Expires: ${expiresAt.trim()}` : "",
          // Recorded because it decides whether anyone ever sees the request, and it is invisible
          // afterwards: the row looks identical either way.
          routed ? `Routed to the ${routed.label} approver: ${routed.value}` : "No approver mapping matched this document.",
        ].filter((d) => d.length > 0),
      }).catch(() => undefined);

      setAsking(undefined);
      setReason("");
      setShareWith("");
      setExpiresAt("");
      setPermission("View");
      setSent(
        routed
          ? `Sent. The ${routed.label} approver for ${routed.value} will see it on the Requests page.`
          : "Sent — but no approver is recorded for this document's unit, so it may sit unanswered. " +
              "Tell an administrator: an APR mapping is missing on the Folder Access page.",
      );
    } catch (e) {
      setProblems([`Could not send the request: ${(e as Error).message}`]);
    } finally {
      setSending(false);
    }
  };

  const [refreshing, setRefreshing] = useState(false);

  /* Extracted so both the mount effect and the manual refresh button below can call it — a plain
     const rather than a `useCallback`, since it closes over nothing that ever changes across
     renders (context/siteUrl are props, stampMissingRef is a ref). */
  const load = async (): Promise<void> => {
    /* Reset per load, or a single transient fallback would suppress the deleted state for the rest
         of the session — including across the live refresh, which re-runs this on focus. */
      stampMissingRef.current = false;
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const segment = await resolveSegment(documentsTitle());
      setDocsSegment(segment);
      const userId = await currentUserId();
      // Sequential, not Promise.all: per-call try/catch is required anyway because
      // Promise.allSettled is unavailable on this tsconfig target (CLAUDE.md #3), and either
      // library failing must produce the "could not read" state rather than a half list.
      const staging = await readLibrary(libraryTitle(), libraryUrlSegment(), true, userId);
      // The resolved segment, not the title — otherwise "Shared Documents" stays in every
      // approved file's folder trail.
      const documents = await readLibrary(documentsTitle(), segment, false, userId);
      /* The Highly Confidential pair, when the site has one.
         An HC uploader's files are invisible on this page without it, and "you have not uploaded
         anything yet" to someone who filed a Highly Confidential document last week is the worst
         possible answer — it invites them to upload it again.

         Read with the SAME AuthorId filter, so this adds no visibility whatsoever: a person sees
         their own HC files, and only where they already hold the grant. An uncleared uploader's
         read returns nothing because SharePoint refuses it, not because this code decided so. */
      const hc = cachedHcLibraries();
      let hcRows: Submission[] = [];
      /* ⚠ WHETHER EACH CHAIN WAS FULLY READ, and it decides whether a missing file may be called
         DELETED. The swallowed catches below are correct — an unreadable HC library is the normal
         case for everyone without clearance — but a swallowed failure leaves the live set silently
         incomplete, and judging a record against an incomplete set is how this page would tell an
         uploader their documents were destroyed. Tracked per CHAIN rather than globally: a single
         flag would be false for every uncleared user on every HC site, and then nothing would ever
         be reported as deleted at all. See `mergeRecords`. */
      let hcChainComplete = true;
      if (hc) {
        // Each in its own try. An HC library the viewer cannot read at all is the NORMAL case for
        // everyone without clearance, and must never take the page down or blank the two libraries
        // that already loaded successfully.
        try {
          hcRows = hcRows.concat(await readLibrary(hc.approval.title, hc.approval.urlSegment, true, userId));
        } catch {
          /* not cleared, or unreachable — either way, nothing of theirs to show */
          hcChainComplete = false;
        }
        try {
          hcRows = hcRows.concat(await readLibrary(hc.documents.title, hc.documents.urlSegment, false, userId));
        } catch {
          /* as above */
          hcChainComplete = false;
        }
      }
      /* ⚠ THE ARCHIVE IS DELIBERATELY NOT READ HERE (reversed 2026-09-02; was read 2026-08-22 to
         2026-09-02). Client: "only C level and system administrator have access to archive... that
         would mean My Submission should not update to archive either." Access to Archive/ArchiveHC
         narrowed to the two C-Level roles in FolderManager.tsx's LIBRARY_ROLES — and My Submissions is
         a PIC/HOU page (Upload-Form.aspx's policy, mirrored here), neither of whom holds a C-Level
         role. Attempting the read would now just 403/404 for every viewer of this page; not attempting
         it at all removes the wasted request and, per the client's own framing, the intent that a
         normal uploader should see nothing of the archive from here.

         `arc` itself is STILL read (a synchronous cache lookup, no request) — `isHcRecord` below still
         needs the archive HC library's title to correctly classify a SUBMISSION RECORD stamped against
         it, independent of whether this page ever fetches rows FROM that library. */
      const arc = cachedArchiveLibraries();
      /* Staging and Documents above are read un-guarded (no per-call try/catch): either failing throws
         out of this function before reaching here, so by the time this line runs both were read in
         full. Kept as a named flag, not a literal `true`, because `mergeRecords` below still takes it
         as a predicate alongside `hcChainComplete` and a literal would obscure why the two differ. */
      const normalChainComplete = true;
      const liveRows = [...staging, ...documents, ...hcRows];

      /* ── THE SUBMISSION RECORD ────────────────────────────────────────────────
         Read LAST and never awaited into the two library reads above: a file an uploader can see
         must never be hidden because a record could not be read.

         `undefined` means NOT-READ (unprovisioned list, throttle, ACL not yet applied) and is not
         `[]`. In that case nothing is merged and this page behaves exactly as it did before the
         feature — which is a degradation, not a lie. */
      let merged: MergedRow[] = liveRows;
      let note: { deleted: number; cancelled: number; archived: number; unknown: number } | undefined;
      try {
        const records = await readSubmissionRecords(context.spHttpClient, siteUrl, userId);
        if (records) {
          /* Which chain is this record's file living in? Decided by the library it was WRITTEN to,
             because that is what the record stores — and a file only ever moves within its own
             chain (approval → documents → archive), never across. */
          const hcTitles = hc
            ? [hc.approval.title.toLowerCase(), hc.documents.title.toLowerCase()]
            : [];
          if (arc?.hc) hcTitles.push(arc.hc.title.toLowerCase());
          const isHcRecord = (r: SubmissionRecord): boolean =>
            hcTitles.indexOf((r.libraryTitle ?? "").toLowerCase()) > -1;
          /* ⚠ NO STAMP ON THE LIVE ROWS MEANS NOTHING CAN BE JUDGED, WHATEVER ELSE SUCCEEDED. If any
             library had to fall back to a read without `SubmissionFileId`, its rows carry no key, so
             every record would fail to match and be reported as DELETED — the worst false positive
             this page can produce. Deliberately global rather than per chain: the flag does not say
             WHICH library fell back, and understating costs a grey "not checked" while overstating
             tells someone their documents were destroyed. */
          const stampReadable = !stampMissingRef.current;
          const result = mergeRecords(records, liveRows, (r) =>
            stampReadable && (isHcRecord(r) ? hcChainComplete : normalChainComplete));
          merged = result.rows;
          if (result.deleted > 0 || result.cancelled > 0 || result.archived > 0 || result.unknown > 0) {
            note = {
              deleted: result.deleted,
              cancelled: result.cancelled,
              archived: result.archived,
              unknown: result.unknown,
            };
          }
        }
      } catch {
        /* Never fatal. The record is an addition; the list of live files is the page. */
      }
      setRecordNote(note);
      setRows(sortNewestFirst(merged));
      setLoadError(undefined);
      // Last, and never awaited into the same try: this decides whether the request buttons can be
      // offered, and nothing about it may cost an uploader the list of their own files.
      loadPolicy().catch(() => setPolicy({ ...BLANK_POLICY, known: false }));
    };

  useEffect(() => {
    load().catch((e) => {
      // NEVER fall through to an empty list. An uploader told they have nothing, when a library was
      // merely unreachable, uploads the file again — and now there are two.
      setRows(undefined);
      setLoadError((e as Error).message);
    });
  }, []);

  /* Manual refresh (client, 2026-09-02: "add a refresh button... instead of refreshing the entire
     page"), same pattern as the Audit Log's icon-only header refresh. Deliberately does NOT clear
     `rows` on failure the way the mount load does — that fallback exists for when there is nothing
     on screen yet; wiping an already-loaded list because one refresh attempt failed would throw away
     good data over a transient error. `load()` itself clears `loadError` on success, so a stale
     banner does not survive a refresh that then succeeds. */
  const onRefresh = (): void => {
    setRefreshing(true);
    // `.finally` is unavailable on this tsconfig target (CLAUDE.md #3, same class as
    // `Promise.allSettled`) — reset the flag on both branches instead.
    load()
      .then(() => setRefreshing(false))
      .catch((e) => {
        setLoadError((e as Error).message);
        setRefreshing(false);
      });
  };

  /* ⚠ THE STATUS COUNTS EXCLUDE DELETED AND UNCHECKED ROWS. A deleted document HAS no approval
     outcome any more, and `rowFromRecord` sets `status` only because `Submission` demands a value —
     counting it would report a destroyed file as Pending or Approved. `liveRowsOnly` is the one place
     that decision lives, so the tabs, the tallies and the filters cannot disagree about it. */
  const counts = countByStatus(liveRowsOnly(rows ?? []));
  /* The Requests tab counts REQUESTS, not documents, so it cannot come from countByStatus — that
     function is about a file's approval status and has no idea any of this exists. */
  /* ⚠ `counts` covers the status tabs only, so "Submissions" read (0) beside a table listing
     fourteen of them — a page contradicting itself, which reads as a bug. Counted from the same
     `groupSubmissions` the tab renders, so the tally and the list cannot disagree. */
  /* ⚠ EVERY submission, bulk included. It used to subtract the bulk runs because they had a tab of
     their own; with that tab gone (2026-08-30) a subtraction here would leave the count disagreeing
     with the list underneath it — which is how a page starts reading as broken. */
  const allGroups = groupSubmissions(rows ?? []);
  const tabCounts: Record<string, number> = {
    ...counts,
    Submissions: allGroups.length,
    // Counted from the same helper the tab renders, so the tally and the list cannot disagree.
    Archive: archivedRowsOnly(rows ?? []).length,
    Requests: myRequestList.length,
  };
  /* ⚠ DELETED ROWS BELONG ON "All" AND NOT ON THE THREE STATUS TABS. `filterByTab` matches
     `r.status === tab`, and a record row carries `status: "Pending"` only because the type demands a
     value — so without this a destroyed document would be listed as awaiting approval, which is a
     claim about a file nobody can open. "All" is the complete history and must show them; the
     Submissions view shows them inside their own batch, which is the client's actual requirement. */
  /* ⚠ `Archive` CANNOT GO THROUGH `filterByTab`: that matches `r.status === tab`, and an archived
     row carries `status: "Pending"` only because the type demands a value. Handled here beside the
     `All` exception rather than as a fourth string comparison inside a function about approval
     outcomes. */
  const shown = tab === "All"
    ? filterByTab(rows ?? [], tab)
    : tab === "Archive"
      ? archivedRowsOnly(rows ?? [])
      : filterByTab(liveRowsOnly(rows ?? []), tab);
  /* The tabs that FILTER the flat file list. `Submissions` and `Requests` render their own views and
     their own empty states, so the shared ones below must not fire for them. */
  const isStatusTab = tab !== "Requests" && tab !== "Submissions";
  // Every library segment the trail builder must strip. Without the HC pair, an HC file's folder
  // trail would start with "HCApprovalDocument" — the library name presented as a folder.
  const hcLibs = cachedHcLibraries();
  /* The two HC URL segments, in the shape `isHcRow` takes. `undefined` when the pair is unresolved,
     which tags nothing — see that function for why guessing from a name prefix would be worse. */
  const hcSegs = hcLibs
    ? { approval: hcLibs.approval.urlSegment, documents: hcLibs.documents.urlSegment }
    : undefined;
  /* The archive segments, in the shape `isArchivedRow` takes. Same rules as `hcSegs`: `undefined`
     tags nothing, and a missing tag is safer than a wrong one. */
  const arcLibs = cachedArchiveLibraries();
  const arcSegs = arcLibs
    ? { normal: arcLibs.normal.urlSegment, hc: arcLibs.hc?.urlSegment }
    : undefined;
  const libs = [
    libraryUrlSegment(), docsSegment, documentsTitle(),
    ...(hcLibs ? [hcLibs.approval.urlSegment, hcLibs.documents.urlSegment] : []),
    /* ⚠ THE ARCHIVE SEGMENTS BELONG HERE TOO, or an archived file's folder trail reads
       "Archive › GHO › GF › TAX" — the library name presented as though it were a folder, which is
       exactly the "Shared Documents" bug of 2026-08-14 at a new library. */
    ...(arcLibs ? [arcLibs.normal.urlSegment, ...(arcLibs.hc ? [arcLibs.hc.urlSegment] : [])] : []),
  ];

  /* ── The detail view ──────────────────────────────────────────────────────────
     An INNER view, not a link out. The client's point: clicking a file used to leave the page for
     the document library, which is the thing this page exists to save them from.

     The preview strategy comes from shared/filePreview.ts — already built and tested for the
     approval page, and already handling the case that matters: SharePoint serves an Office file as
     a DOWNLOAD, so a raw URL in an iframe renders nothing at all. PDF, Office, image, text and an
     honest refusal are each handled there rather than re-guessed here. */
  /**
   * ONE request dialog, rendered by BOTH the list and the detail view.
   *
   * Built as a value rather than duplicated into each return: it carries the external-sharing
   * warning, the expiry rule and the pending-file wording, and a second copy is how one route ends
   * up permitting what the other refuses.
   *
   * `subject` is the row being asked about — set by whichever button was pressed — never `open`,
   * which is undefined when the dialog is opened from the list.
   */
  /* Debounced lookup. 3 characters minimum — a one-letter query returns noise and costs a request per
     keystroke. A FAILED search shows nothing and blocks nothing: the address can always be typed, so
     this is a convenience and must never stand between someone and their request.
     Declared here, ABOVE the detail view's early return, because a hook must run on every render. */
  useEffect(() => {
    const q = peopleQ.trim();
    if (q.length < 3) {
      setPeopleHits([]);
      return undefined;
    }
    let live = true;
    setPeopleBusy(true);
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then((hits) => {
          if (!live) return;
          // Only entries carrying an address: SP.Web.ShareObject shares with an EMAIL, so a result
          // without one cannot be acted on and offering it would be a dead row.
          setPeopleHits(hits.filter((h) => (h.email ?? "").indexOf("@") > -1));
          setPeopleBusy(false);
        })
        .catch(() => {
          if (!live) return;
          setPeopleHits([]);
          setPeopleBusy(false);
        });
    }, 300);
    return () => {
      live = false;
      clearTimeout(t);
    };
  }, [peopleQ, context.spHttpClient, siteUrl]);

  /** Add an address to the box, unless it is already there. Compared through `parseRecipients` so the
      check matches whatever separators the person actually typed. */
  const addRecipient = (email: string): void => {
    const clean = (email || "").trim();
    if (clean.length === 0) return;
    const have = parseRecipients(shareWith).map((r) => r.toLowerCase());
    if (have.indexOf(clean.toLowerCase()) === -1) {
      const base = shareWith.replace(/\s+$/, "");
      setShareWith(base.length > 0 ? base + "\n" + clean : clean);
    }
    setPeopleQ("");
    setPeopleHits([]);
  };

  /** Drop one recipient. Matched case-insensitively, because that is how it was added. */
  const removeRecipient = (email: string): void => {
    const gone = (email || "").trim().toLowerCase();
    setShareWith(parseRecipients(shareWith).filter((r) => r.toLowerCase() !== gone).join("\n"));
  };

  /* Keep the request statuses current without the uploader pressing refresh (client, 2026-08-21).
     Reloads ONLY `loadPolicy`, which re-reads the small requests list — never the two document
     libraries, which are the expensive reads and barely change while someone watches this page.
     Blocked while a request dialog is open, a send is in flight, or a cancel is running.
     Declared here, above the detail view's early return, because a hook must run every render. */
  useLiveRefresh(loadPolicy, sending || asking !== undefined || cancelling !== undefined);

  const requestDialog = ((): React.ReactNode => {
    const subject = askRow;
    if (asking === undefined || subject === undefined) return null;
    /* The recipients, parsed ONCE per render of this dialog. `shareWith` remains the single source of
       truth — the chip list and the external check both read this, so what is displayed and what is
       validated can never diverge. */
    const shareRecipients = parseRecipients(shareWith);
    return (
      <div style={s.modalBg} onClick={() => { if (!sending) setAsking(undefined); }}>
            <div style={s.modal} onClick={(e) => e.stopPropagation()}>
              {/* Client's wording, 2026-08-30: the title states the ACTION being requested rather
                  than narrating it. */}
              <p style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px" }}>
                {asking === "Deletion" ? "Delete this file?" : "Share this file?"}
              </p>
              <p style={{ fontSize: 12.5, color: "#605e5c", margin: "0 0 4px", lineHeight: 1.5 }}>
                {subject.name}
                {/* Repeated here, not only on the row: by the dialog the uploader has committed to a
                    file, and this is the last moment they can notice it is the HC copy rather than
                    the same-named ordinary one. */}
                {isHcRow(subject.library, hcSegs) && <span style={s.hcTag}>HC</span>}
                {isArchivedRow(subject.library, arcSegs) && <span style={s.arcTag}>Archived</span>}
              </p>
              <p style={{ fontSize: 12.5, color: "#605e5c", margin: "0 0 8px", lineHeight: 1.5 }}>
                {/* ⚠ "93 days", NOT the 90 in the client's mock — confirmed with them 2026-08-30 and
                    again 2026-09-03: *"stay 93, they might not know that is why they say 90."*
                    SharePoint's recycle bin really is 93, and that number is the only thing making an
                    approved deletion reversible. A promise of 90 would be wrong in the safe direction
                    on the day it mattered, but it would still be wrong, in a dialog people act on.
                    "HOD/HOU" -> "approver" is the client's own wording, 2026-09-03 — same fact,
                    said the way whoever is deciding actually holds the role (a Head of Unit for most
                    units, a Head of Department where they fan down). */}
                {asking !== "Deletion"
                  ? "Your approver must approve this request. If approved, the people below get access to this file — and nothing else."
                  : subject.status === "Approved"
                  ? "Your approver must approve this request. If approved, the file moves to the recycle bin and can be restored within 93 days."
                  // Named for what it IS from the uploader's side. "Withdraw" was considered and
                  // rejected: it reads as something they can do themselves, and the whole point of
                  // this change is that they no longer can.
                  /* Client's own copy, 2026-09-04, on two lines.
                     ⚠ 93 DAYS, NOT THE 90 IN THEIR MOCK. Same conflict, same resolution as
                     2026-08-30, when they were asked directly and said *"stay 93, they might not
                     know that is why they say 90"*. SharePoint's site recycle bin really is 93.
                     The dropped clause was *"so only you and your approver can see it"* — still
                     TRUE (Draft Item Security), just no longer stated here. */
                  : (
                    <>
                      This file has not been approved yet.
                      <br />
                      To delete this file, your approver must approve this request. If approved, the
                      file moves to the recycle bin and can be restored within 93 days.
                    </>
                  )}
              </p>

              <label style={s.label}>Reason</label>
              <textarea
                style={{ ...s.field, minHeight: 64, resize: "vertical" }}
                value={reason}
                onChange={(e) => setReason(e.target.value)}
              />

              {asking === "Share" && (
                <>
                  <label style={s.label}>Share with</label>
                  <input
                    style={s.field}
                    placeholder="Search for a person, or type an email address below"
                    value={peopleQ}
                    onChange={(e) => setPeopleQ(e.target.value)}
                  />
                  {peopleBusy && <div style={s.hint}>Searching&hellip;</div>}
                  {!peopleBusy && peopleQ.trim().length >= 3 && peopleHits.length === 0 && (
                    /* NOT an error, and must not read as one. Nobody matching is the NORMAL case for
                       a share request — the person is usually outside the organisation, which is the
                       whole reason this workflow exists. So it points at the box below rather than
                       reporting a failure. */
                    <div style={s.hint}>
                      Nobody on this site matches. If they are outside the organisation, type their
                      full email address in the box below.
                    </div>
                  )}
                  {peopleHits.length > 0 && (
                    <div style={s.pickList}>
                      {peopleHits.map((h) => (
                        <button
                          key={h.loginName}
                          type="button"
                          style={s.pickRow}
                          onClick={() => addRecipient(h.email)}
                        >
                          <strong>{h.displayName}</strong>
                          <span style={{ color: "#605e5c" }}> — {h.email}</span>
                        </button>
                      ))}
                    </div>
                  )}
                  {/* WARN: THIS WAS A FREE-TEXT textarea AND IS NOW READ-ONLY (client, 2026-09-03:
                      "Ensure client cannot type directly into the Email Address text box, but rather
                      let them type one by one in the share with"). Recipients are added through the
                      picker above and removed with the X here; `shareWith` is still the single
                      source of truth, so `parseRecipients`, `validateDraft` and the submit path are
                      all unchanged.

                      WARN: IT REVERSES A DELIBERATE DECISION, and why that is now safe is worth
                      recording. The box existed because the point of a share request is usually
                      somebody who is NOT on the site yet, and `searchTenantPeople` will not always
                      surface such an address — so typing was the only route to it. That case is
                      CLOSED BY POLICY: adding outsiders was withdrawn on 2026-08-27 (the client does
                      not want outsiders added), and where external sharing is off the external gate
                      refuses those addresses anyway. So the box had become the one way to type an
                      address the form would then refuse.
                      WARN: IF THE CLIENT EVER RE-ALLOWS OUTSIDE SHARING, this is the control to
                      revisit first — the picker alone may not reach every address they are then
                      permitted to share with.

                      WARN: SCROLLS AT 140px rather than growing without limit (client asked for it):
                      a long list would otherwise push Submit off the dialog, which has no scroll of
                      its own. Nothing inside is absolutely positioned — the people picker sits ABOVE
                      this box, not in it — so the cap cannot clip a popover, which is the trap that
                      has caught three other screens in this project. */}
                  <label style={s.label}>Email addresses</label>
                  <div
                    style={{
                      ...s.field, minHeight: 56, maxHeight: 140, overflowY: "auto", height: "auto",
                      display: "flex", flexWrap: "wrap", alignContent: "flex-start", gap: 6,
                      padding: "8px 10px", cursor: "default",
                    }}
                  >
                    {shareRecipients.length === 0 ? (
                      <span style={{ fontSize: 12, color: "#8a8886" }}>
                        Nobody yet — find a person in the box above.
                      </span>
                    ) : (
                      shareRecipients.map((addr) => (
                        <span key={addr} style={s.recipChip}>
                          {addr}
                          <button
                            type="button"
                            style={s.recipX}
                            title={"Remove " + addr}
                            aria-label={"Remove " + addr}
                            onClick={() => removeRecipient(addr)}
                          >
                            &times;
                          </button>
                        </span>
                      ))
                    )}
                  </div>
                  {/* Named as it is typed, not only after submitting: whether someone is outside the
                      organisation is the fact that decides whether this is refused, and finding out
                      at the end means retyping the lot. */}
                  {/* WARN: IT USED TO ASSERT "outside the organisation" WITH NO WAY TO CHECK, and
                      that is exactly what the client hit: a colleague on their own tenant was
                      flagged as outside it. `isExternal` fails CLOSED on purpose — NO DOMAINS
                      SUPPLIED MEANS UNKNOWN, AND UNKNOWN COUNTS AS EXTERNAL — so one sentence was
                      being produced by two completely different states and named only one of them.
                      The list now decides which sentence appears, and NAMES the domains it compared
                      against, so an admin sees in one glance whether the answer is "this really is
                      outside" or "this site has never been told what inside means". The check itself
                      is unchanged: guessing INTERNAL on an unconfigured site is how a document
                      leaves the organisation on the strength of nothing. */}
                  {shareRecipients.some((r) => isExternal(r, policy.tenantDomains)) && (
                    <div style={{ ...s.warnBox, marginTop: 8 }}>
                      {policy.tenantDomains.length === 0 ? (
                        <>
                          This site has not been told which email domains belong to your
                          organisation, so <strong>every</strong> address is treated as outside it.
                          An administrator sets <strong>tenantDomains</strong> on the CRS Config list
                          — until then no share request can be sent.
                        </>
                      ) : policy.allowExternal ? (
                        <>
                          Some of these are outside your organisation ({policy.tenantDomains.join(", ")}).
                          Your approver will be told.
                        </>
                      ) : (
                        <>
                          Some of these are outside your organisation ({policy.tenantDomains.join(", ")}),
                          and sharing outside it is switched off on this site — the request cannot be
                          sent until you remove them.
                        </>
                      )}
                    </div>
                  )}

                  {/* ⚠ "VIEW AND EDIT" REMOVED FROM THE DROPDOWN (client, 2026-09-03: "Should not
                      have the VIEW and EDIT here"). `permission` stays a fixed "View" — an
                      uploader-raised share request grants read access to the recipient; letting them
                      also ask for Edit rights was never the intent of this screen. `SharePermission`
                      itself is untouched (Requests.tsx still reads it), so this is display-only: the
                      state simply never becomes anything but "View" from this dialog now. */}
                  <label style={s.label}>They may</label>
                  <div style={{ ...s.field, display: "flex", alignItems: "center", background: "#f3f2f1", color: "#605e5c" }}>
                    View only
                  </div>

                  <label style={s.label}>Access ends on (optional)</label>
                  <input
                    type="date"
                    style={s.field}
                    value={expiresAt}
                    onChange={(e) => setExpiresAt(e.target.value)}
                  />
                  <p style={{ fontSize: 11.5, color: "#605e5c", marginTop: 4, lineHeight: 1.45 }}>
                    Leave it blank and the access is permanent until someone removes it.
                  </p>
                </>
              )}

              {problems.length > 0 && (
                <div style={{ ...s.rejectBox, marginTop: 12, marginBottom: 0 }}>
                  {problems.map((p) => <div key={p}>{p}</div>)}
                </div>
              )}

              <div style={{ display: "flex", gap: 8, marginTop: 16, flexWrap: "wrap" }}>
                <button
                  style={sending ? s.askOff : s.primary}
                  disabled={sending || fieldText === undefined}
                  onClick={() => { submitRequest(subject, asking).catch(() => undefined); }}
                >
                  {/* "Reading the file…" is a real state, not a spinner for its own sake: routing to the
                      right approver needs this file's tier fields, and FieldValuesAsText is a
                      per-ITEM endpoint the list has not called. Sending without it would route
                      on a blank unit — into nobody's queue, reported as sent. */}
                  {sending ? "Sending…" : fieldText === undefined ? "Reading the file…" : "Submit"}
                </button>
                <button style={s.askBtn} disabled={sending} onClick={() => setAsking(undefined)}>
                  Cancel
                </button>
              </div>
            </div>
      </div>
    );
  })();

  /* Why requests cannot be raised AT ALL right now — the half that is about the SITE rather than any
     one file, so both the detail view and every row in the list can say the same thing. Unknown is
     its own answer and clears itself in a second; treating it as "not set up" would send an admin
     hunting for a list that is merely still being probed. */
  const requestsUnavailable = !policy.known
    ? "Checking whether requests can be raised…"
    : !policy.listExists
      ? "Requests are not set up on this site yet — an administrator opens the Requests page once to create the list."
      : undefined;

  if (open !== undefined) {
    const preview = previewTarget(open.name, open.fileRef, tenantRoot, siteUrl);
    // FieldValuesAsText keys arrive double-encoded, so each row asks for both spellings — the same
    // reason pickField exists. Blank values are dropped rather than shown as empty rows.
    /* The metadata rows come from shared/documentDetails.ts, which DERIVES the tier rows instead of
       naming them.

       This list used to name `Department` and `Unit` literally, so on Upstream Operations Malaysia —
       whose tiers are Region and Estate/Mill — both read blank, blank rows were dropped, and the two
       values that decide where the file lives were missing from the panel. Reported 2026-08-14 on a
       file at `UPOPSMY › JHR › BKB › 2024 › Working File`, which showed the segment and nothing under
       it. Every segment onboarded from here has different tier names, so no list could stay right.

       `leading`/`trailing` are the rows only this screen knows. They are passed through blank-or-not,
       which is why each says "unknown" rather than being omitted: an absent row reads as a file with
       no location, and this screen can tell the difference. */
    const trail = trailText(folderTrail(open.fileRef, libs));
    const details = buildDetailRows({
      fieldText: fieldText ?? {},
      leading: [{ label: "Location", value: trail || "the library root" }],
      trailing: [
        // Both come from the list query, not FieldValuesAsText — one round trip already spent.
        { label: "File size", value: formatBytes(open.size) || "unknown" },
        // Uploaded is in the header line; Last updated is what changed when the approver acted.
        { label: "Last updated", value: formatSubmittedOn(open.modified) },
      ],
    });

    /* The latest request this person raised against the OPEN file, if any. */
    const openRequest = myRequests[(open.uniqueId ?? "").toLowerCase()];
    /* `RequestedAt` is a full ISO timestamp, so `new Date` is unambiguous here — the UTC-shift trap
       applies to DATE-ONLY strings, which land on the previous day west of Greenwich. An unparseable
       value yields "", so a bad stored date costs the date and never the whole banner. */
    const askedOn = ((): string => {
      const d = new Date(openRequest?.at ?? "");
      return openRequest?.at && !isNaN(d.getTime()) ? formatSubmittedOn(d) : "";
    })();

    /* Why a request cannot be raised for THIS file, or undefined when it can.
       A reason, never a disappeared button: an uploader who cannot see the control assumes the
       system does not do this and goes to ask someone in person. */
    const requestBlock = requestsUnavailable ?? (!open.uniqueId
      ? "This file's id could not be read, so a request would not find it. Reload the page and try again."
      : undefined);

    return (
      <section style={s.wrap}>
        <div style={s.backBand}>
          <button style={s.backLink} onClick={() => { setOpen(undefined); setFieldText(undefined); }}>
            ‹ Back to my submissions
          </button>
        </div>

        <h2 style={s.h2}>{open.name}</h2>
        {/* THE PENDING-REQUEST STATE IS A CHIP BESIDE THE STATUS, NOT A PARAGRAPH (client's design,
            2026-08-30). It says the same thing where people already look for a file's state, and it
            sits ALONGSIDE the approval status rather than replacing it - a file can be Approved AND
            have a deletion pending, which is exactly the case the amber banner made look like a
            contradiction. The banner's other job, saying nothing has happened yet, moves to a quiet
            line under the divider. */}
        {/* ⚠ BADGES ON THEIR OWN ROW, "Uploaded ..." ON THE ONE BELOW (client's mockup, 2026-09-03) —
            REVISES the 2026-08-30 single-line layout above. The client's own words: "our current one
            is almost correct already, its just that the text is not below." Same facts, same chip,
            just no longer sharing a line with the Uploaded/path text. */}
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 6 }}>
          <span style={badgeFor(open.status)}>{open.status}</span>
          {openRequest !== undefined && openRequest.status === "Pending" && (
            <span style={s.pendingChip}>
              {openRequest.type === "Share" ? "📤" : "🗑"}{" "}
              {openRequest.type === "Share" ? "Share" : "Deletion"} pending
            </span>
          )}
        </div>
        <p style={s.sub}>
          Uploaded {formatSubmittedAt(open.created)} &nbsp;·&nbsp;{" "}
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

        {/* ── Asking for something to be done to this file ─────────────────────────
            EVERY file, not only approved ones — changed 2026-08-20 when the PIC lost `DELS` (client:
            "For Staging PIC should not be able to delete, they have to request from HOU"). Until
            then a PIC deleted pending and rejected files outright, so offering a request here would
            have taught people to ask for something they already controlled. That is no longer true,
            and without this the removal of DELS left them no route at all.

            ⚠ THIS PAGE IS WHAT MAKES A REQUEST "YOUR OWN FILE". Both reads filter `AuthorId eq me`,
            so an uploader can only ever raise a request against something they filed — structurally,
            with no author check to write and therefore none to get wrong. The buttons were briefly
            moved to the library command bar on 2026-08-20 and moved straight back for exactly this
            reason: in `Documents` a PIC sees their whole unit's approved files, so a button there
            would have let them ask for a colleague's document to be deleted. **Do not move them
            again without replacing this guarantee with a real one.**

            SHARE STAYS APPROVED-ONLY. A pending file has been approved by nobody, and LIBRARY_ROLES
            keeps SHARE off both approval libraries — so a Head of Unit could not carry such a share
            out even having agreed to it. `validateDraft` refuses it too; hiding the button alone
            would make a rule into a UI decision. */}
        {/* ALREADY ASKED. Shown before the buttons, and it REPLACES them while a request is pending:
            a second request for the same file gives the Head of Unit two rows for one decision, and
            they cannot tell a mistake from a reminder. A decided request does NOT hide them — a
            rejected deletion may legitimately be asked for again once the reason is fixed. */}
        {openRequest !== undefined && (
          <div style={openRequest.status === "Pending" ? s.askQuiet : s.askDecided}>
            {openRequest.status === "Pending" ? (
              <>
                {openRequest.type === "Share" ? "Share" : "Deletion"} request submitted
                {askedOn ? ` on ${askedOn}` : ""}. Nothing has happened to the file yet.
              </>
            ) : (
              /* ⚠ EVERY STATUS IS NAMED, because the fallback was "attempted and failed" and TWO
                 statuses reach it that did not fail. Found live 2026-08-30 on a REVOKED share: the
                 uploader was told their request "was attempted and failed by <the approver>" when it
                 had in fact been approved, used, and later taken back — which reads as the approver
                 having broken something. `Cancelled` had the same defect for longer: withdrawing
                 your own request came back as a failure somebody else caused.

                 The lesson is the one this codebase keeps paying for: adding a value to a shared
                 union makes every READER a change site. `Revoked` was added for the Requests page
                 and this screen was never revisited. */
              <>
                <strong>
                  {openRequest.status === "Cancelled"
                    ? `You withdrew this ${openRequest.type === "Share" ? "share" : "deletion"} request`
                    : openRequest.status === "Revoked"
                      ? "This share was approved, and the access has since been taken back"
                      : `Your ${openRequest.type === "Share" ? "share" : "deletion"} request was ` +
                        (openRequest.status === "Approved"
                          ? "approved"
                          : openRequest.status === "Rejected"
                            ? "rejected"
                            : "attempted and failed")}
                </strong>
                {/* Nobody "decided" a withdrawal but the requester, and on a revoke `decidedBy` is
                    the person who APPROVED it — naming them beside "taken back" would credit the
                    wrong act to the wrong person. `revokedBy` is not carried on this page's row. */}
                {openRequest.status !== "Cancelled" && openRequest.status !== "Revoked" && openRequest.decidedBy
                  ? ` by ${openRequest.decidedBy}`
                  : ""}
                .
                {/* ⚠ THE NOTE IS WITHHELD ON A REVOKE, and not only for tidiness: after a revoke it
                    holds the appended line, which names EVERY principal dropped from the file —
                    including people granted outside CRS whom this uploader never named and has no
                    reason to be shown. For Approved/Rejected/Failed it is the reason they came for. */}
                {openRequest.status !== "Revoked" && openRequest.status !== "Cancelled" && openRequest.note
                  ? ` “${openRequest.note}”`
                  : ""}
                {openRequest.status === "Rejected" ? " You can ask again if something has changed." : ""}
                {openRequest.status === "Revoked" ? " You can ask again if it is still needed." : ""}
              </>
            )}
          </div>
        )}

        {/* ⚠ A REQUEST IS NOT OFFERED TO SOMEBODY WHO CAN ALREADY DO IT (client, 2026-08-30). The
            tiers come from the document's OWN `<Base>Tid` columns, so a Head of Department matches
            on their department term and a Head of Unit on their unit term — the fan-out falls out
            of the comparison. `fieldText` is undefined until the per-item read lands, which yields
            an empty chain and therefore offers the buttons: unknown must never HIDE a route. */}
        {(() => {
          const chain = fieldText ? documentUnit(fieldText).tiers.map((t) => t.guid) : [];
          const approved = open.status === "Approved";
          const canDeleteSelf = canActDirectly(
            chain,
            approved ? policy.directDelete : policy.directDeleteStaging,
          );
          const canShareSelf = approved && canActDirectly(chain, policy.directShare);
          const showDelete = !canDeleteSelf;
          const showShare = approved && !canShareSelf;
          return (
        <div style={s.askBar}>
          {sent === undefined && openRequest?.status !== "Pending" && (
            <>
              {showDelete && (
              <button
                style={requestBlock === undefined ? s.askBtn : s.askOff}
                disabled={requestBlock !== undefined}
                title={requestBlock}
                onClick={() => { setProblems([]); setAskRow(open); setAsking("Deletion"); }}
              >
                Request deletion
              </button>
              )}
              {showShare && (
                <button
                  style={requestBlock === undefined ? s.askBtn : s.askOff}
                  disabled={requestBlock !== undefined}
                  title={requestBlock}
                  onClick={() => { setProblems([]); setAskRow(open); setAsking("Share"); }}
                >
                  Request share
                </button>
              )}
            </>
          )}
          {/* SAYS WHY THE BUTTONS ARE ABSENT. A row that simply has no controls reads as a page
              that failed to finish loading — and this person can act, so the useful sentence names
              where. */}
          <span style={{ fontSize: 12, color: "#605e5c" }}>
            {!showDelete && !showShare
              ? "You can delete or share this document yourself, in the library — no request is needed."
              : requestBlock ??
                (approved
                  ? "Your Approver decides these."
                  : "Your Approver decides this. A file awaiting approval cannot be shared, only deleted.")}
          </span>
        </div>
          );
        })()}

        {sent !== undefined && <div style={s.okBox}>{sent}</div>}

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
                title={`Preview of ${open.name}`}
                allowFullScreen
              />
            )}
            {/* ⚠ "Open in a new tab" REMOVED HERE (client, 2026-09-03, screenshot of this exact
                page: "Remove the Open in a new tab"). Same reasoning as ApprovalDocument.tsx's
                identical removal — this was a supplementary fallback link shown for BOTH the image
                and iframe branches (and, before this, redundantly TWICE on the "none" branch, which
                has always had its own copy just above). That one stays: there the preview genuinely
                cannot render anything, so its link is the sole route to the document, not decoration. */}
          </div>

          <div>
            <div style={s.sectionTitle}>Details</div>
            {fieldText === undefined && <p style={s.empty}>Loading details&hellip;</p>}
            {/* The metadata is the part that can be missing — Location and the file facts always
                have a value, so `details` is never empty and cannot carry this state itself. */}
            {fieldText !== undefined && Object.keys(fieldText).length === 0 && (
              // Empty and unreadable look the same from here, so say the honest thing: the file and
              // its status are still correct above, which is what the page is for.
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

        {requestDialog}
      </section>
    );
  }

  return (
    <section style={s.wrap}>
      <div style={s.headRow}>
        <h2 style={s.h2}>My Submissions</h2>
        <button
          type="button"
          style={s.headRefresh}
          disabled={refreshing}
          title="Re-read your submissions"
          aria-label="Refresh"
          onClick={onRefresh}
        >
          {refreshing ? "…" : "↻ Refresh"}
        </button>
      </div>
      <p style={s.sub}>
        Every file you have uploaded, and where it has got to. Only your own files are listed here.
      </p>

      {requestNotice !== undefined && (
        /* Amber, not red: nothing went wrong. The approver simply decided first, which is the
           system working. The row list below has already been refreshed to show the decision. */
        <div style={s.warnBox}>{requestNotice}</div>
      )}

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
            {t} ({tabCounts[t] ?? 0})
          </button>
        ))}
      </div>

      {/* ── Outgoing requests ────────────────────────────────────────────────────
          Client, 2026-08-20: "Add two tabs too showing Outgoing Request and inside it shows the
          delete request or Share request and they can cancel the request."

          Both types in ONE tab with a Type column, because they are the same errand from the
          requester's side — what did I ask for, and what happened to it. Two tabs would split a
          four-row list in half and hide whichever kind they were not looking at. */}
      {/* ── Submissions: the uploader's own view of what they sent ─────────────────
          Client, 2026-08-22. Three levels, all read-only and all served from rows already loaded —
          drilling in filters what is on screen and issues no request. The per-file metadata panel is
          the existing detail view, reached from level 3, which is the one place that reads
          FieldValuesAsText (a per-ITEM endpoint, so it can only ever be one file at a time). */}
      {tab === "Submissions" && (() => {
        /* EVERY submission, however it was filed. The bulk runs used to be filtered out to their own
           tab; they are rows in this list now, told apart by the Upload column. */
        const groups = groupSubmissions(rows ?? []);
        /* ⚠ `some`, NOT `every`, and it matches the rule the old tab partition used. A group made by
           the import tool is entirely bulk, so the two agree in every real case — but an INFERRED
           group (folder + date, for files uploaded before submissions were recorded) can legitimately
           mix the two, and calling such a row "Batch" would hide that a bulk import is in it. */
        const isBulkGroup = (g: SubmissionGroup): boolean => g.files.some(isBulkUploadRow);
        const keyOf = (g: SubmissionGroup): string =>
          g.reference || (g.files[0] ? submissionKey(g.files[0]) : "");
        const current = groups.filter((g) => keyOf(g) === openSubmission)[0];
        const batch = current
          ? current.batches.filter((b) => (b.reference || "—") === openBatch)[0]
          : undefined;

        const countLine = (files: MergedRow[]): string => {
          /* Live rows only for the APPROVAL states — a deleted file has no approval outcome, and
             counting its placeholder `Pending` would report a destroyed document as awaiting
             approval. */
          const c = statusCounts(liveRowsOnly(files));
          // Only the non-zero states. "0 rejected" on every row is noise, and the one number that
          // matters is buried among the ones that do not.
          const parts = (["Approved", "Pending", "Rejected"] as const)
            .filter((k) => c[k] > 0)
            .map((k) => `${c[k]} ${k.toLowerCase()}`);
          /* ⚠ AND THE GONE ONES, COUNTED SEPARATELY AND SHOWN LAST (client, 2026-08-27: *"what would
             be great is if you put the status as deleted"*).
             Excluding them from the approval tally was right; leaving them out of the line ENTIRELY
             was not — a submission whose only file has been deleted rendered a BLANK Status cell,
             because all three approval counts were zero and nothing was left to join. Seen live on
             `SUB-20260827-FJNQ`. A blank cell on the page whose whole purpose is saying what became
             of a file is the one thing it must never show.
             `unknown` is worded differently on purpose: it means the libraries could not all be read,
             so those files may be perfectly fine. */
          /* ⚠ EVERY RECORD STATE, FROM ONE LIST — and it is a list because this line has now been
             short twice. `cancelled` was left out on 2026-08-28 and `archived` on 2026-09-03, each
             time leaving a submission whose files were ALL in that state with a **blank Status
             cell** (client: *"the files in Status for Archive is not showing"*). On the page whose
             whole job is saying what became of a file, blank is the one answer it must never give.
             `recordStateParts` is driven by a `Record` over the union, so a fifth state is a compile
             error rather than another blank cell. */
          return parts.concat(recordStateParts(recordCounts(files))).join(" · ");
        };

        if (rows === undefined) return <p style={s.empty}>Loading&hellip;</p>;
        if (groups.length === 0) {
          return <p style={s.empty}>You have not uploaded anything yet.</p>;
        }

        // ── Level 3: one batch, as the upload form that created it ───────────────
        /* Client, 2026-08-22: "when I say upload Form I literally meant a copy of the upload form
           with the details in it but view based only."

           So this is the form's own shape: the destination chosen ONCE at the top — the permissioned
           tiers plus Year and Document Type, which are what the folder IS — and then every file in
           the batch with the details that were typed for it, all expanded. Not a list to click
           through: the uploader filled this in as one screen and is checking it as one screen.

           ⚠ THIS IS THE ONE PLACE THAT COSTS REQUESTS — one `FieldValuesAsText` per file, because it
           is a per-ITEM endpoint. Bounded (a batch is a handful of files) and deliberate (they opened
           it). Everything above is served from rows already loaded. */
        if (current && batch) {
          // Keyed by `mergedKey`, matching what `loadBatchText` writes — a record row's item id
          // belongs to another list, so `submissionKey` could collide across the two kinds of row.
          /* ⚠ A KEY THAT IS NOT IN THE MAP IS `undefined`, NEVER `{}`. It used to answer `{}`, which
             says "read, and there was nothing" — and that is a different fact from "not read", which
             is what a missing key actually means. It went wrong the moment a file was approved:
             `mergedKey` is `library#itemId` for a live row, Auto-route re-creates the file in the
             approved-side library with a NEW id, and the map loaded while it was pending then had no
             key for it. The card fell back to `{}` and rendered File size alone — a document that
             looked like it had never carried any metadata. */
          const ftFor = (f: MergedRow): Record<string, string> | undefined =>
            batchText ? batchText[mergedKey(f)] : undefined;
          /* The destination, taken from whichever file in the batch could be read. A batch IS one
             folder, so any of them answers — but the FIRST one may be the one whose read failed, and
             falling back to it would show an empty destination for a batch that has one. */
          const destSource = batchText
            ? batch.files.map(ftFor).filter((ft) => ft && Object.keys(ft).length > 0)[0]
            : undefined;
          /* ⚠ AND WHEN NO FILE IN THE BATCH IS STILL IN A LIBRARY, THE FOLDER CARD COMES FROM THE
             RECORD. Every file archived or deleted means no live item to read, so `destSource` is
             undefined and this card showed nothing but Location — the client's report, verbatim:
             *"the archived files wont show Document folder information like a normal file does."*
             The snapshot holds the tiers, so a gone batch renders the same card a live one does.
             Taken from whichever gone file HAS a snapshot: one written before the record feature
             existed is empty, and it must not stand in for a sibling that has one. */
          const destRows = destSource
            ? buildBatchRows(destSource)
            : batch.files
              .map((f) => snapshotFolderRows(f.record?.metadata))
              .filter((r) => r.length > 0)[0] ?? [];
          return (
            <div>
              {/* ⚠ A BULK RUN GOES STRAIGHT BACK TO THE LIST. Its batch level is a list of one,
                  and it was skipped on the way in - so clearing only `openBatch` would land the
                  reader on a level they never saw and cannot act on. */}
              <button
                style={s.backBtn}
                onClick={() => {
                  setOpenBatch(undefined);
                  setBatchText(undefined);
                  /* ⚠ CLEARING THE MAP WITHOUT CLEARING ITS SIGNATURE WOULD STRAND THE NEXT VISIT.
                     Reopening the same batch computes the same signature, the effect would decide it
                     had already loaded it, and nothing would ever fill the map it just emptied. The
                     two are one fact and are cleared together. */
                  batchSigRef.current = "";
                  /* ⚠ KEYED ON THE GROUP, not on the tab it was reached from (the tab is gone).
                     A bulk run has ONE destination, so its batch level lists exactly one row and
                     answers nothing — the client asked for it to be skipped on the way in
                     (2026-08-28), so it has to be skipped on the way out too, or Back lands on the
                     empty level that was deliberately jumped. */
                  if (isBulkGroup(current)) setOpenSubmission(undefined);
                }}
              >
                &larr; Back to {isBulkGroup(current) ? "the list" : current.reference || "this submission"}
              </button>

              <div style={s.card}>
                <p style={s.cardHead}>Document folder information</p>
                <div style={s.grid2}>
                  {/* Location leads and is ALWAYS shown, blank or not: it comes from the file path
                      rather than the metadata, so it is the one row that is right even when the
                      per-item read failed. */}
                  <div style={s.detailRow}>
                    <div style={s.detailLabel}>Location</div>
                    <div style={s.detailValue}>
                      {trailText(folderTrail(batch.files[0]?.fileRef ?? "", libs)) || "the library root"}
                    </div>
                  </div>
                  {destRows.map((r) => (
                    <div key={r.label} style={s.detailRow}>
                      <div style={s.detailLabel}>{r.label}</div>
                      <div style={s.detailValue}>{r.value}</div>
                    </div>
                  ))}
                </div>
                {batchText === undefined && <div style={s.trail}>Reading the details&hellip;</div>}
                {/* ⚠ THE BATCH REFERENCE WAS SHOWN HERE TOO, AND IS NOW GONE (client, 2026-09-03,
                    same reasoning as the table two levels up on 2026-08-30: "User will be confused
                    with this info"). Still stored, still what groups these files — only the display
                    is gone, on both screens now. */}
              </div>

              <p style={s.cardHead}>
                Document details &mdash; {batch.files.length} file{batch.files.length === 1 ? "" : "s"}
              </p>
              {/* ⚠ SCROLLED (client, 2026-08-28). A Bulk Upload run holds up to 50 files, each
                  rendering a card with a metadata grid — long enough to bury the footer explaining
                  what Pending and Rejected mean, and to put the Back link a long scroll away.

                  The COUNT stays OUTSIDE the box so it never scrolls off: the same rule the
                  abbreviation editor follows, where its warning, collision banner and Save all sit
                  outside its own 58vh box.

                  Checked before adding `overflow`: nothing in this file is absolutely positioned, so
                  there is no descendant for the scroll container to clip. That check is not optional —
                  it has caught three screens here (the people-picker in GroupManager, the info panels
                  in `.dms-staged-row`, the add-box dropdown). Anything added inside these cards later
                  must re-check it. */}
              <div style={{ maxHeight: "60vh", overflowY: "auto" }}>
              {batch.files.map((f) => {
                /* ⚠ A GONE FILE'S DETAILS COME FROM THE RECORD, NOT FROM THE ITEM. `FieldValuesAsText`
                   is a per-ITEM endpoint, so for a deleted document it answers nothing and the card
                   would read "no details were recorded" — which is false and defeats the point of
                   keeping the row. The snapshot was written at upload for exactly this moment.
                   No `File size`: that came from the live list query and there is nothing to size. */
                const gone = f.recordState !== undefined;
                const ft = gone ? undefined : ftFor(f);
                /* ⚠ "STILL READING" IS A SEPARATE FACT FROM "NOTHING TO READ", AND `ft` CANNOT
                   CARRY BOTH. A gone file is never fetched, so its `ft` is `undefined` — and the
                   line below keyed the "Reading the details…" message on exactly that, so every
                   archived and deleted file sat under a message saying its details were on their
                   way, for ever, directly above the details themselves. Seen on site 2026-09-03. */
                const readingFt = !gone && ft === undefined;
                /* ⚠ ONLY THE PER-DOCUMENT HALF OF THE SNAPSHOT. It used to render every key, which
                   put Business Segment, Department and Unit on the file card — the folder card's
                   job, one card above, and the reason that card looked empty. */
                const rows2 = gone
                  ? snapshotFileRows(f.record?.metadata)
                  : ft
                    ? buildFileRows({
                      fieldText: ft,
                      // From the list query, not FieldValuesAsText — no extra round trip.
                      trailing: [{ label: "File size", value: formatBytes(f.size) || "unknown" }],
                    })
                    : [];
                return (
                  <div key={mergedKey(f)} style={f.recordState ? { ...s.fileCard, ...s.goneRow } : s.fileCard}>
                    <div style={s.fileHead}>
                      <strong style={{ fontSize: 13 }}>{f.name}</strong>
                      {isHcRow(f.library, hcSegs) && <span style={s.hcTag}>HC</span>}
                {isArchivedRow(f.library, arcSegs) && <span style={s.arcTag}>Archived</span>}
                      {/* ⚠ THE GONE STATES COME FIRST, and there is no `Open file`. This is the view
                          the client actually asked for — the file still listed inside its own batch,
                          with what was submitted — so the row must stay and only the actions go. */}
                      {f.recordState === "deleted" ? (
                        <span style={s.goneBadge}>Deleted</span>
                      ) : f.recordState === "cancelled" ? (
                        <span style={s.cancelBadge}>Cancelled</span>
                      ) : f.recordState === "archived" ? (
                        <span style={s.archivedBadge}>Archived</span>
                      ) : f.recordState === "unknown" ? (
                        <span style={s.unsureBadge}>Not checked</span>
                      ) : (
                        <span style={badgeFor(f.status)}>{f.status}</span>
                      )}
                      {/* The way to the full panel — preview, rejection reason, and the deletion and
                          share buttons. This view is read-only by design; that one is where an
                          uploader ACTS on a file. */}
                      {!f.recordState && (
                        <button style={s.openBtn} onClick={() => openRow(f)}>Open file</button>
                      )}
                    </div>
                    {/* ⚠ WHO AND WHEN, or "Cancelled" is just a politer "Deleted". The client asked
                        for this state so a replaced document reads as superseded rather than lost —
                        which only helps if the uploader can see who filed the newer version and go
                        and look at it. Rendered from the record, so it survives the document. */}
                    {/* ⚠ SAYS WHERE IT WENT, or "Archived" is just a politer "Deleted" again. The
                        uploader cannot open it from here and cannot browse to it — archive access is
                        C-Level only — so without this the row states a fact and leaves them with
                        nowhere to go. Naming who to ask is the actionable part. */}
                    {f.recordState === "archived" && (
                      <p style={s.trail}>
                        Moved to the archive
                        {f.record?.archivedAt
                          ? ` on ${f.record.archivedAt.toISOString().slice(0, 10)}`
                          : ""}
                        {" "}under the seven-year retention rule. It is still kept, but only C-level
                        users and system administrators can open the archive — ask them if you need a
                        copy. The details below are what <em>this</em> submission carried.
                      </p>
                    )}
                    {f.recordState === "cancelled" && (
                      <p style={s.trail}>
                        Replaced by a newer upload
                        {f.record?.replacedBy ? ` from ${f.record.replacedBy}` : ""}
                        {f.record?.replacedAt
                          ? ` on ${f.record.replacedAt.toISOString().slice(0, 10)}`
                          : ""}
                        . The details below are what <em>this</em> submission carried.
                      </p>
                    )}
                    {readingFt && <p style={s.trail}>Reading the details&hellip;</p>}
                    {/* Empty ≠ unknown: `{}` here is a read that failed or an item with nothing
                        recorded, and the two are indistinguishable from the response. Say so rather
                        than showing a card that looks like a document with no metadata.
                        Gone files reach this too, when the record predates the snapshot — the
                        sentence is true of both, and a blank card is what it exists to prevent. */}
                    {!readingFt && rows2.length === 0 && (
                      <p style={s.trail}>No details were recorded for this file, or they could not be read.</p>
                    )}
                    {rows2.length > 0 && (
                      <div style={s.grid2}>
                        {rows2.map((r) => (
                          <div key={r.label} style={s.detailRow}>
                            <div style={s.detailLabel}>{r.label}</div>
                            <div style={s.detailValue}>{r.value}</div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
              </div>
            </div>
          );
        }

        // ── Level 2: the batches inside one submission ───────────────────────────
        if (current) {
          return (
            <div>
              <button style={s.backBtn} onClick={() => setOpenSubmission(undefined)}>
                &larr; Back to all submissions
              </button>
              <p style={s.cardHead}>
                {current.reference || "Grouped by folder and date — uploaded before submissions were recorded"}
                {" — "}{formatSubmittedAt(current.at)}
              </p>
              <table style={s.table}>
                <thead>
                  <tr>
                    <th style={s.th}>Batch &mdash; where these files went</th>
                    <th style={s.th}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {current.batches.map((b) => (
                    <tr key={b.reference || submissionKey(b.files[0])}>
                      <td style={s.td}>
                        {/* ⚠ THIS DOES NOT LOAD THE DETAILS, DELIBERATELY. The effect near the top
                            watches `openBatch` and loads them — because level 1 opens a bulk run's
                            batch directly and forgot to, which left every card reading "Reading the
                            details…" for ever. One route in, one loader. */}
                        <button
                          style={s.nameBtn}
                          onClick={() => setOpenBatch(b.reference || "—")}
                        >
                          {trailText(folderTrail(b.files[0]?.fileRef ?? "", libs)) || "—"}
                        </button>
                        {/* ⚠ THE BATCH REFERENCE (`BAT-20260827-WFDZ`) IS NO LONGER SHOWN (client,
                            2026-08-30: *"User will be confused with this info"*). It is still
                            STORED and still what groups these files — only the display is gone. The
                            SUBMISSION reference above is kept: that is the one an uploader quotes to
                            their approver, which is the whole reason references are readable rather
                            than GUIDs. A batch is an internal subdivision of it. */}
                        <div style={s.trail}>
                          {b.files.length} file{b.files.length === 1 ? "" : "s"}
                        </div>
                      </td>
                      <td style={s.td}>{countLine(b.files)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        }

        // ── Level 1: every submission ────────────────────────────────────────────
        return (
          <div>
            {refsMissing && (
              /* Says WHY rather than looking broken. Not an error: grouping is an addition, and
                 everything below still lists correctly without it. */
              <div style={s.warnBox}>
                These libraries have not been reconciled since submissions were introduced, so older
                files carry no reference and are grouped by folder and date instead &mdash; a good
                guess, not a record. An administrator running <strong>Folder Reconciliation</strong>{" "}
                adds the columns; every upload after that is grouped by what was actually sent.
              </div>
            )}
            <div style={s.scroller}>
            <table style={s.table}>
              <thead>
                <tr>
                  <th style={s.th}>Submission</th>
                  <th style={s.th}>Batches</th>
                  {/* ⚠ THE "Upload" COLUMN IS GONE (client, 2026-09-03: "you forgot to remove the
                      upload column, add Bulk tagging beside the folder name"). It read `Batch` on
                      almost every row — the ordinary case, saying nothing — so a whole column was
                      spent distinguishing the rare one. The `Bulk` chip moved INLINE beside the
                      folder path in the Submission cell, and `Batch` is now simply the unmarked
                      default. Supersedes the 2026-08-30 note that put this column beside Batches. */}
                  <th style={s.th}>Files</th>
                  <th style={s.th}>Status</th>
                  <th style={s.th}>Uploaded</th>
                </tr>
              </thead>
              <tbody>
                {groups.map((g) => {
                  /* ⚠ THE ROW IS LABELLED BY ITS DESTINATION, NOT BY A REFERENCE OR A FILE NAME
                     (client, 2026-08-22, sketching it: `GHO › GF › TAX › 2024 › Term Sheet`).
                     A reference is what you QUOTE to an approver; a path is what you RECOGNISE, and
                     recognising the one you want is the whole job of this list. The reference stays,
                     below and in monospace, for the quoting. */
                  const where = trailText(folderTrail(g.batches[0]?.files[0]?.fileRef ?? "", libs));
                  const more = g.batches.length - 1;
                  return (
                  <tr key={keyOf(g)}>
                    <td style={s.td}>
                      <button
                        style={s.nameBtn}
                        onClick={() => {
                          setOpenSubmission(keyOf(g));
                          /* ⚠ A BULK RUN HAS ONE DESTINATION, so its batch level lists exactly one
                             row - a click that answers nothing. Opening the batch at the same time
                             takes the reader straight to the files (client, 2026-08-28: *"it will
                             open the second page immediately ... because there isn't any
                             batches"*). */
                          setOpenBatch(isBulkGroup(g) ? g.batches[0]?.reference || "—" : undefined);
                        }}
                      >
                        {where || g.files[0]?.name || "—"}
                        {/* Named rather than silently showing only the first: a submission that went
                            to three places must not look like one that went to one. */}
                        {more > 0 && ` + ${more} more destination${more === 1 ? "" : "s"}`}
                      </button>
                      {/* ⚠ ONLY BULK IS TAGGED (client, 2026-09-03). A batch upload is the ordinary
                          way files arrive here, so marking it said nothing; the tag now marks the
                          exception, beside the path it belongs to rather than in a column of its
                          own. `isBulkGroup` is unchanged — this is where its answer is shown, not
                          what it answers. */}
                      {isBulkGroup(g) && <span style={s.bulkChip}>Bulk</span>}
                      {g.inferred ? (
                        /* Said plainly. These rows are grouped by folder and day because nothing on
                           them records the upload — a good guess, and the screen must not pass it
                           off as a recorded submission. */
                        <div style={s.trail}>Grouped by folder and date &mdash; uploaded before submissions were recorded</div>
                      ) : (
                        <div style={s.refLine}>{g.reference}</div>
                      )}
                    </td>
                    <td style={s.td}>{g.batches.length}</td>
                    <td style={s.td}>{g.files.length}</td>
                    <td style={s.td}>{countLine(g.files)}</td>
                    <td style={s.td}>{formatSubmittedAt(g.at)}</td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
            </div>
          </div>
        );
      })()}

      {tab === "Requests" && (
        <>
          {!policy.known && <p style={s.empty}>Checking&hellip;</p>}
          {policy.known && myRequestList.length === 0 && (
            <p style={s.empty}>
              You have not asked for anything yet. Open one of your files and use{" "}
              <strong>Request deletion</strong> or <strong>Request share</strong>.
            </p>
          )}
          {myRequestList.length > 0 && (
            <div style={s.scroller}>
            <table style={s.table}>
              <thead>
                <tr>
                  <th style={s.th}>File</th>
                  <th style={s.th}>Asked for</th>
                  <th style={s.th}>Status</th>
                  <th style={s.th}>Raised</th>
                  <th style={s.th} />
                </tr>
              </thead>
              <tbody>
                {myRequestList.map((rq) => {
                  const on = new Date(rq.at);
                  return (
                    <tr key={rq.id}>
                      <td style={s.td}>{rq.itemName || <em style={{ color: "#605e5c" }}>unnamed</em>}</td>
                      <td style={s.td}>{rq.type === "Share" ? "Share" : "Deletion"}</td>
                      <td style={s.td}>
                        <span style={badgeFor(rq.status === "Cancelled" ? "Rejected" : rq.status)}>
                          {rq.status}
                        </span>
                        {rq.decidedBy ? (
                          <div style={{ fontSize: 11, color: "#605e5c", marginTop: 2 }}>by {rq.decidedBy}</div>
                        ) : null}
                        {rq.note ? (
                          <div style={{ fontSize: 11, color: "#605e5c", marginTop: 2 }}>&ldquo;{rq.note}&rdquo;</div>
                        ) : null}
                      </td>
                      <td style={s.td}>{isNaN(on.getTime()) ? "—" : formatSubmittedOn(on)}</td>
                      <td style={s.td}>
                        {/* Only the requester's own PENDING rows, and this list only ever holds their
                            own. Cancelling a DECIDED request is refused by canCancel: it could not
                            un-recycle a file or take back a share, so offering it would promise
                            something the button cannot do. */}
                        {rq.status === "Pending" ? (
                          <button
                            style={cancelling === rq.id ? s.askOff : s.askBtn}
                            disabled={cancelling !== undefined}
                            onClick={() => { cancelRequest(rq).catch(() => undefined); }}
                          >
                            {cancelling === rq.id ? "Cancelling…" : "Cancel"}
                          </button>
                        ) : null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            </div>
          )}
        </>
      )}

      {requestDialog}

      {/* ⚠ THESE BELONG TO THE STATUS TABS ALONE, and gating them on `!== "Requests"` was not enough
          once `Submissions` was added (2026-08-22). `filterByTab` has no case for it, so `shown` is
          empty and the page printed "Nothing is submissions right now" UNDERNEATH a table listing
          seven of them — a screen contradicting itself, which reads as a bug. Named as a predicate
          rather than a second literal, so the next tab added cannot walk back into it: anything that
          is not a status filter belongs here. */}
      {isStatusTab && rows === undefined && loadError === undefined && <p style={s.empty}>Loading&hellip;</p>}

      {/* ⚠ THE DELETED / REPLACED / ARCHIVED SENTENCES ARE GONE (client, 2026-09-03: *"Remove these
          text in My Submissions"*, quoting the whole paragraph). Each row already carries the same
          fact as a badge — Deleted, Cancelled, Archived — so the summary restated at the top of the
          page what the list says line by line, and on a site with real history it ran to four
          sentences before the table.

          ⚠ THE `unknown` SENTENCE IS DELIBERATELY KEPT. It is the only one of the four that is NOT
          restated by a badge in a way anybody can act on: it means a LIBRARY READ FAILED, so rows
          may be missing from this page entirely and the ones marked "not checked" may be perfectly
          fine. Dropping it would leave a page that is quietly incomplete with nothing saying so —
          the opposite of the other three, which describe rows that ARE listed. */}
      {recordNote !== undefined && recordNote.unknown > 0 && (tab === "All" || tab === "Submissions") && (
        <p style={s.goneNote}>
          {recordNote.unknown} could not be checked, because one of the libraries did not respond{" "}
          — {recordNote.unknown === 1 ? "it" : "they"} may be perfectly fine. Reload to try again.
        </p>
      )}

      {/* Three empty states, kept apart on purpose — spec §4.2. */}
      {isStatusTab && rows !== undefined && rows.length === 0 && (
        <p style={s.empty}>
          You have not uploaded anything yet. Files you submit on the upload form appear here, with
          their approval status.
        </p>
      )}
      {isStatusTab && rows !== undefined && rows.length > 0 && shown.length === 0 && (
        <p style={s.empty}>
          {/* ⚠ `Archive` NEEDS ITS OWN WORDING: the generic form is built from the tab name, which
              would read "Nothing is archive right now". It is also the one tab whose empty state is
              the NORMAL case for years — nothing on either site is close to seven years old — so it
              says why rather than implying something is missing. */}
          {tab === "Archive"
            ? "No files of yours have been archived yet. Documents move here seven years after they were filed."
            : `Nothing ${tab === "All" ? "here" : `is ${tab.toLowerCase()}`} right now. Your other files are under the tabs above.`}
        </p>
      )}

      {shown.length > 0 && (
        <div style={s.scroller}>
        <table style={s.table}>
          <thead>
            <tr>
              <th style={s.th}>File</th>
              <th style={s.th}>Status</th>
              <th style={s.th}>Uploaded</th>
              <th style={s.th} />
            </tr>
          </thead>
          <tbody>
            {shown.map((r) => (
              /* ⚠ KEYED BY `mergedKey`, NOT `submissionKey`. The latter is `library#itemId`, and a
                 record row's item id belongs to a DIFFERENT list — so two records could collide with
                 each other and with a live document, and React would silently drop a row. */
              <tr key={mergedKey(r)} style={r.recordState ? s.goneRow : undefined}>
                <td style={s.td}>
                  {/* ⚠ A GONE FILE IS NOT A BUTTON. There is nothing to open: no preview, no Open
                      file, and the detail view would read the item live and find nothing. Branching
                      on `recordState` BEFORE anything else is the rule the whole record rests on. */}
                  {r.recordState ? (
                    <span style={s.goneName}>{r.name}</span>
                  ) : (
                    /* A button, not a link. It opens the detail view INSIDE this page — the client's
                       request, and the fix for a click that used to land the uploader in the
                       document library. The file itself is still one click further, from there. */
                    <button style={s.nameBtn} onClick={() => openRow(r)}>
                      {r.name}
                    </button>
                  )}
                  {/* ⚠ THE SAME NAME CAN EXIST IN BOTH LIBRARIES — one upload classified Highly
                      Confidential, one not — and without this they are indistinguishable here: same
                      name, same tier path, same status. An uploader raising a deletion was choosing
                      blind, and only the recycle bin's Original location revealed which went. */}
                  {isHcRow(r.library, hcSegs) && <span style={s.hcTag}>HC</span>}
                {isArchivedRow(r.library, arcSegs) && <span style={s.arcTag}>Archived</span>}
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
                  {/* ⚠ NEVER THE APPROVAL BADGE FOR A GONE FILE. `status` is `Pending` on a record
                      row only because `Submission` demands a value; showing it would tell an uploader
                      a destroyed document is awaiting approval. `unknown` is its own answer — the
                      libraries could not all be read, so this says so rather than claiming either. */}
                  {r.recordState === "deleted" ? (
                    <span style={s.goneBadge}>Deleted</span>
                  ) : r.recordState === "cancelled" ? (
                    <span style={s.cancelBadge}>Cancelled</span>
                  ) : r.recordState === "archived" ? (
                    <span style={s.archivedBadge}>Archived</span>
                  ) : r.recordState === "unknown" ? (
                    <span style={s.unsureBadge}>Not checked</span>
                  ) : (
                    <span style={badgeFor(r.status)}>{r.status}</span>
                  )}
                  {/* A REQUEST is not a document status, so it gets its own quiet tag rather than
                      replacing the badge — a file can be Approved AND have a deletion pending, and
                      collapsing the two would hide whichever mattered less to whoever wrote the
                      code. Only PENDING is shown: a decided request is history, and the detail view
                      carries the outcome and the approver's note. */}
                  {myRequests[(r.uniqueId ?? "").toLowerCase()]?.status === "Pending" && (
                    // "asked" -> "requested" (client, 2026-09-03) — same fact, the word the rest of
                    // this screen already uses ("Your approver must approve this request").
                    <span style={s.askedTag}>
                      {myRequests[(r.uniqueId ?? "").toLowerCase()].type === "Share"
                        ? "Share Requested"
                        : "Deletion Requested"}
                    </span>
                  )}
                </td>
                <td style={s.td}>{formatSubmittedAt(r.created)}</td>
                {/* ── Ask from the ROW ────────────────────────────────────────────────
                    Client, 2026-08-20: "put a button beside the uploaded for each file so its
                    easier to share or delete instead of going into each file manually."

                    Same rules as the detail view, from the same places, so the two routes cannot
                    diverge: deletion on every file, share on APPROVED files only, and nothing at all
                    while a request is already pending — a second request for one file gives the Head
                    of Unit two rows for one decision.

                    Pressing one loads that file's tier metadata, because routing to the right
                    approver reads the document's own `<Base>Tid` fields and this list does not carry
                    them (FieldValuesAsText is a per-ITEM endpoint). The dialog's Send button waits
                    for it rather than sending with nothing, which would route on a blank unit. */}
                <td style={{ ...s.td, whiteSpace: "nowrap", textAlign: "right" }}>
                  {/* ⚠ NO REQUEST BUTTONS ON A GONE FILE. A deletion request for a document that is
                      already gone is meaningless, and a share request would be APPROVED by a Head of
                      Unit and then fail in their own session — after the requester was told it was
                      being handled. The approval executes against the document, and there is none. */}
                  {/* ⚠ NOTHING IS RENDERED FOR A GONE ROW, and the empty branch is doing the work.
                      The words "no longer here" / "replaced" / "in the archive" are gone (client,
                      2026-09-03) — each repeated, in the next column, what the row's own badge
                      already says. What must NOT be lost is the branch itself: a gone row gets no
                      Delete and no Share button. For an archived file that is not merely tidiness —
                      `validateDraft` REFUSES a deletion or share request against one outright
                      (2026-08-22), because every role holds only Read on the archive and an approved
                      request would fail in the approver's own session days later. */}
                  {/* ⚠ EMPTY WHILE A REQUEST IS PENDING, NOT the word "requested" (client,
                      2026-09-03). The chip beside the file name already says `Deletion Requested` /
                      `Share Requested`, so this column was the same fact a third time. The BRANCH
                      stays and is what matters: while a request is open the Delete and Share buttons
                      must not be offered, or the uploader raises a second row for a decision the
                      approver has not made yet — which is exactly what the 2026-08-20 "you already
                      asked" work exists to prevent. */}
                  {r.recordState ? undefined : myRequests[(r.uniqueId ?? "").toLowerCase()]?.status === "Pending" ? undefined : (
                    <>
                      <button
                        style={requestsUnavailable === undefined && r.uniqueId ? s.rowBtn : s.rowBtnOff}
                        disabled={requestsUnavailable !== undefined || !r.uniqueId}
                        title={requestsUnavailable ?? "Ask your Approver to delete this file"}
                        onClick={() => {
                          setProblems([]);
                          setAskRow(r);
                          setFieldText(undefined);
                          setAsking("Deletion");
                          loadFieldText(r).catch(() => setFieldText({}));
                        }}
                      >
                        Delete
                      </button>
                      {r.status === "Approved" && (
                        <button
                          style={requestsUnavailable === undefined && r.uniqueId ? s.rowBtn : s.rowBtnOff}
                          disabled={requestsUnavailable !== undefined || !r.uniqueId}
                          title={requestsUnavailable ?? "Ask your Approver to share this file"}
                          onClick={() => {
                            setProblems([]);
                            setAskRow(r);
                            setFieldText(undefined);
                            setAsking("Share");
                            loadFieldText(r).catch(() => setFieldText({}));
                          }}
                        >
                          Share
                        </button>
                      )}
                    </>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      )}

      {/* A BULLET LIST since 2026-08-30 (client's own wording). It was one paragraph of four facts,
          which is exactly the shape people skip — and the fourth, that the request buttons live
          inside an opened file, is the one nobody finds on their own. */}
      <ul style={s.note}>
        <li><strong>Pending</strong> means your file is waiting for your approver.</li>
        <li><strong>Rejected</strong> files stay here so you can see the reason for rejection.</li>
        <li>
          <strong>Approved</strong> files are moved to the main <strong>{documentsTitle()}</strong>{" "}
          library, where the rest of your unit can find them too.
        </li>
        <li>
          Open an approved file to request that it be <strong>deleted</strong> or{" "}
          <strong>shared</strong>. Your <strong>approver</strong> will review and decide.
        </li>
      </ul>
      {/* Its own paragraph, and rendered ONLY when it applies — an always-present empty <p> under
          the list is a gap nobody can explain. */}
      {commentsMissing && (
        <p style={s.note}>
          Rejection reasons cannot be shown on this site — the comment column could not be read.
        </p>
      )}
    </section>
  );
}
