import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { swapLibrarySegment } from "../../../shared/hcRouting";
import { markRecordReplaced } from "../../../shared/spSubmissionRecords";
import { documentsLibraryTitle } from "../../../shared/naming";
import { libraryHasColumns, APPROVED_BY_COLUMN } from "../../../shared/optionalColumns";
import { IApprovalDocumentProps } from "./IApprovalDocumentProps";
import {
  buildQueue,
  nextUndecidedIndex,
  statusToDecision,
  decisionFromLink,
  swapState,
  Decision,
  QueueEntry as QueueEntryOf,
} from "../../../shared/approvalQueue";
import { previewTarget } from "../../../shared/filePreview";
import { cachedHcLibraries, hcAvailable, libraryTitle, libraryUrlSegment } from "../../../shared/naming";
import { primeNames, listTitleEncoded, LIST_SUFFIX } from "../../../shared/spNaming";
import { permissionedTierCount } from "../../../shared/approvalDestination";
// ⚠ THE CHECKS THEMSELVES LIVE IN shared/approvalGuards.ts, shared with the bulk approve command
// set. Two implementations of "is it safe to approve this" is how a bulk route ends up weaker
// than the page it copies. Do not re-inline them here.
import {
  checkUnitFolderReady,
  checkDestinationClash,
  checkApproveRight,
  GuardResult,
} from "../../../shared/approvalGuards";

/**
 * A permissions-and-approvals screen must never render a cached answer: the destination guard's
 * mode-row read decides whether Approve is allowed to fire at all, and the SAME shape of bug
 * (a `304 Not Modified` silently replaying a stale body) already broke the Requests page and My
 * Submissions (2026-08-30) before it broke this guard too. `Cache-Control`/`Pragma` alone can be
 * ignored by a proxy or service worker; `bust()` — a unique query string on every call — cannot.
 */
const GET_FRESH = {
  Accept: "application/json;odata=nometadata",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
};
const bust = (): string => `&_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;

// ── Types ────────────────────────────────────────────────────────────────────

// Library URL segments used to map an approval-library path to its Documents twin.
// "Shared Documents" is the Documents library's URL segment even though its display name is
// "Documents" — a fixed built-in, so it stays a constant.
//
// The approval library's segment is NOT a constant: it is resolved at mount by primeNames,
// because its title and URL no longer match (title "Approval Document", URL
// "/ApprovalDocument"). Always read it through the function — a module-level const would be
// evaluated at import time, before priming, and silently freeze the legacy "Staging".
const DOCUMENTS_URL_SEGMENT = "Shared Documents";

/** Which library the document being reviewed lives in. */
type ApprovalLib = "normal" | "hc";

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

// Decision, the queue shape, and the queue rules all come from src/shared/approvalQueue.ts —
// pure and unit-tested there, because "which document is next" and "what a swap clears" are the
// parts of this feature that fail silently rather than visibly.
type QueueEntry = QueueEntryOf<IFileItem>;

// ── Helpers ──────────────────────────────────────────────────────────────────

function initials(name: string): string {
  return name.split(" ").slice(0, 2).map(n => n[0] ?? "").join("").toUpperCase();
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  // 08/Jan/2035 — DD/MMM/YYYY, matching the Staging views' column formatting.
  // en-GB gives "08 Jan 2035"; some ICU builds append a dot to the month, so strip it.
  const date = d
    .toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })
    .replace(/\./g, "")
    .replace(/\s+/g, "/");
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
  /* The Upload Form's page shell — see the note in AuditLog.tsx. This page had a BOTTOM padding and
     no horizontal one, so the preview and the approval panel both ran into the window edge.
     ⚠ `maxWidth` 1180 rather than the Upload Form's 960: the layout here is a two-column grid whose
     right column reserves a fixed 280px for the approval panel, so a narrow cap squeezes the document
     PREVIEW — the one thing an approver is on this page to read. Its own 40px bottom padding is kept
     rather than replaced by the shell's 48. */
  root:        { fontFamily: "'Segoe UI', Tahoma, sans-serif", color: "#323130", background: "#fff", maxWidth: 1180, margin: "32px auto", padding: "0 24px 40px" } as React.CSSProperties,
  // A full-width band, tinted from the same green as the link so the two read as
  // one control. The tint is an alpha of the brand green rather than a second
  // hex value — one colour to change if the brand shifts.
  backBand:    { background: "rgba(0, 104, 74, 0.08)", borderRadius: 4, padding: "10px 16px", marginBottom: 20 } as React.CSSProperties,
  backLink:    { color: "rgba(0, 104, 74, 1)", textDecoration: "none", fontSize: 14, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6 } as React.CSSProperties,
  // 22px, not 38. Filenames here are composed — [Project] - [Vendor] - [Name] -
  // [DDMMYY] — so they are long by design and ran to three enormous lines.
  // overflowWrap breaks a single unspaced run rather than letting it overhang.
  docTitle:    { margin: "0 0 6px", fontSize: 22, lineHeight: 1.3, fontWeight: 600, color: "#201f1e", overflowWrap: "break-word" as const } as React.CSSProperties,
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
  // maxHeight + scroll so a pasted essay in Remark or Details cannot push the
  // rest of the Details list below the fold — the approver would never scroll
  // past it to find Confidential Level.
  metaValue:   { fontWeight: 500, fontSize: 13, color: "#201f1e", lineHeight: 1.35, overflowWrap: "break-word" as const, maxHeight: 132, overflowY: "auto" as const } as React.CSSProperties,
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
  // Secondary popup button — only used when a primary is present, so the two are distinguishable.
  // No marginTop any more: the buttons live in a flex row (popupBtnRow) whose `gap` spaces them on
  // BOTH axes. A vertical margin did nothing when they sat side by side, which is how they ended up
  // touching each other (found on site 2026-08-18) — the card is text-align:center and a <button> is
  // inline, so they flowed onto one line with no separation at all.
  popupBtnGhost:{ background: "#fff", color: "#201f1e", border: "1px solid #8a8886" } as React.CSSProperties,
  popupBtnRow: { display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap", marginTop: 24 } as React.CSSProperties,
  navRow:      { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" as const, marginBottom: 8 } as React.CSSProperties,
  navControls: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" as const } as React.CSSProperties,
  navBtn:      { padding: "5px 12px", fontSize: 13, fontFamily: "inherit", fontWeight: 600, color: "#0f6c3f", background: "#fff", border: "1px solid #0f6c3f", borderRadius: 4, cursor: "pointer" } as React.CSSProperties,
  navBtnOff:   { color: "#a19f9d", borderColor: "#c8c6c4", cursor: "not-allowed" } as React.CSSProperties,
  navCount:    { fontSize: 13, color: "#605e5c", minWidth: 74, textAlign: "center" as const } as React.CSSProperties,
  navBadge:    { fontSize: 12, fontWeight: 700, padding: "2px 8px", borderRadius: 10 } as React.CSSProperties,
  navBadgeOk:  { color: "#0f6c3f", background: "#e7f4ec", border: "1px solid #b7dcc4" } as React.CSSProperties,
  navBadgeNo:  { color: "#a4262c", background: "#fde7e9", border: "1px solid #f1b0b3" } as React.CSSProperties,
  // Shared frame for the non-iframe previews, so an image and a "no preview" message occupy the
  // same space an iframe would and the three-column layout does not shift between documents.
  previewBox:  { display: "flex", alignItems: "center", justifyContent: "center", width: "100%", height: "calc(100vh - 240px)", minHeight: 600, background: "#faf9f8", border: "1px solid #edebe9", borderRadius: 4, overflow: "hidden", padding: 12, boxSizing: "border-box" as const } as React.CSSProperties,
  // The same frame, for IMAGES only, fitting to WIDTH and scrolling.
  //
  // `alignItems: flex-start` so a tall image starts at its TOP rather than being centred with its
  // head out of view, and `overflow: auto` so everything below the fold is reachable. Together with
  // the <img> rule below this is the whole fix for a screenshot rendering as a sliver.
  imageBox:    { display: "flex", alignItems: "flex-start", justifyContent: "center", width: "100%", height: "calc(100vh - 240px)", minHeight: 600, background: "#faf9f8", border: "1px solid #edebe9", borderRadius: 4, overflow: "auto", padding: 12, boxSizing: "border-box" as const } as React.CSSProperties,
  previewLink: { fontSize: 13, color: "#0f6cbd" } as React.CSSProperties,
};

// ── Component ─────────────────────────────────────────────────────────────────

const ApprovalDocument: React.FC<IApprovalDocumentProps> = ({ context }) => {
  const [item, setItem]           = useState<IFileItem | null>(null);
  /**
   * Whether THIS user may approve THIS document — asked of the item itself, not inferred from a role.
   *
   * WHY (2026-08-25, client: "is there no way to like to ensure the uploader for that group dont see
   * that page?"). The page grant on ApprovalDocument.aspx is per PAGE, not per unit — holding `APR`
   * in ANY unit opens it, and a page cannot be scoped to a unit. So a Head of Unit for one unit who
   * is also an uploader in another could open this page on their OWN upload in that second unit, see
   * a fully live Approval panel, and only discover at submit that they hold nothing there — via a raw
   * `UnauthorizedAccessException`.
   *
   * Not a leak: the folder ACL and Draft Item Security decide what is READABLE, and what they saw was
   * their own document. It is a missing pre-check, and the same one the upload form already makes
   * before offering a destination.
   *
   * FAILS OPEN. `unknown` shows the panel, because the 403 at submit is still the real backstop and a
   * transient read must never take the approval queue out of service for a genuine approver.
   */
  const [approveRight, setApproveRight] = useState<"granted" | "denied" | "unknown">("unknown");
  const [loading, setLoading]     = useState(true);
  const [decision, setDecision]   = useState<Decision>("Approved");
  const [comments, setComments]   = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState<Decision | null>(null);
  const [fetchError, setFetchError] = useState("");
  const [submitError, setSubmitError] = useState("");
  const [fieldText, setFieldText]   = useState<IFieldText>({});
  /**
   * The existing-document check, run PROACTIVELY as soon as the document loads rather than only at
   * the moment Approve is clicked (client, 2026-09-02: *"there is no need for two popup as the one we
   * made is enough … precheck the file against the Document Library and then show the text to them
   * under the comment box"*). `undefined` means not yet checked, or no item; a `GuardResult` with
   * `ok: true` means no clash — nothing to show. Displayed under the comment box; `submitDecision`
   * re-checks fresh immediately before approving (its own long-standing rule, unrelated to display)
   * rather than trusting this snapshot, so a clash appearing in the gap is still caught.
   */
  const [clashCheck, setClashCheck] = useState<GuardResult | undefined>(undefined);

  /**
   * Permissioned tier count per segment folder — `{ gho: 2, upopsmy: 2 }`, keyed lower-cased.
   *
   * The approval guard needs it to locate the unit folder, and nothing else on this page does. It is
   * NOT derivable from the document's own fields: every tier column has a `<Base>Tid` twin, so a
   * below-Unit tier such as SubUnit is indistinguishable from a permissioned one there (see
   * `documentDetails.documentUnit`). Only the mode row's `Levels` knows.
   *
   * `undefined` means the read has not finished or FAILED — never "no segments". The guard refuses on
   * undefined and names it, which is the same direction the rest of that guard already fails in: a
   * refused approval costs a retry, a wrong pass publishes a unit's documents to everyone.
   */
  const [tierCounts, setTierCounts] = useState<Record<string, number> | undefined>(undefined);

  /**
   * The approver's queue: the pending documents they can see, oldest first.
   *
   * Captured ONCE at mount and never re-queried. Deciding removes a document from the
   * server-side pending set, so a refetch would make items vanish from under the approver's
   * position and turn Next into an unpredictable jump. Instead each entry carries its own
   * `decided` flag, so positions stay fixed for the life of the page.
   */
  const [queue, setQueue]       = useState<QueueEntry[]>([]);
  const [pos, setPos]           = useState(0);
  /**
   * True while a swap's FieldValuesAsText is in flight.
   *
   * Not reusing `loading`: that blanks the whole page, which would flash the entire layout on
   * every Next press. The queue entry already holds everything the header and preview need, so
   * only the metadata labels are genuinely pending.
   */
  const [swapping, setSwapping] = useState(false);

  const webUrl = context.pageContext.web.absoluteUrl;

  /* ---------- Which approval library ------------------------------------------
     Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md

     This page reviews ONE document, reached by ?itemId=. Item ids are per-LIST, so that parameter
     alone cannot say which library the document is in — and pointing the read at the wrong one
     either 404s or, worse, opens a DIFFERENT document that happens to share the id.

     Resolved two ways, explicit first:
       1. `?lib=hc` on the link. The HC library's Name-column formatting should carry it, exactly as
          the normal library's carries the link to this page at all.
       2. Failing that, the normal library is tried and the HC one only if the item is not there.
          Normal-first preserves today's behaviour precisely, so an id present in both resolves the
          way it always has.

     A REF, not just state: the helpers below are called inside the same async pass that resolves
     the library, and state set mid-pass would not be visible to them. */
  const initialLib: ApprovalLib =
    new URLSearchParams(window.location.search).get("lib") === "hc" ? "hc" : "normal";
  const [lib, setLib] = useState<ApprovalLib>(initialLib);

  /* ---------- ?decision=approve|reject, from Crystal's [Approve]/[Reject] email links ----------
     Spec: docs/superpowers/specs/2026-08-28-email-bundling-and-templates-design.md §6.1.

     The link PRE-TICKS a radio. It never decides anything: `submitDecision` is still reached only
     by pressing the button, so the destination-folder guard, the name-clash check and the
     `ApproveItems` probe all still run. A link that decided from the inbox would skip all three,
     and would need a bearer trigger URL sitting in up to ten mailboxes.

     ⚠ A REF, READ ONCE. Three things depend on that:
       - the effect below STRIPS the parameter from the address bar, so the value must already be
         captured by the time it runs;
       - the queue rewrites `?itemId=` with replaceState as the approver moves through it, and a
         parameter that survived would pre-tick a decision from an email about a DIFFERENT
         document — the one failure here that could actually mislead someone;
       - a refresh would otherwise re-apply it after the approver had deliberately changed it.
     `useRef`'s initialiser runs on every render and only the first result is kept, which is exactly
     the once-per-mount semantics wanted. */
  const linkDecision = useRef<Decision | undefined>(
    decisionFromLink(new URLSearchParams(window.location.search).get("decision") ?? undefined),
  );

  useEffect(() => {
    if (!linkDecision.current) return;
    try {
      const u = new URL(window.location.href);
      u.searchParams.delete("decision");
      window.history.replaceState(undefined, "", u.toString());
    } catch { /* cosmetic — the radio is already set, and the value is held in the ref regardless */ }
  }, []);

  /* Plain state, no ref. `loadItem` is the only place that needs the resolved value before a render
     happens, and it passes it explicitly to the two loads it makes — so nothing here has to be
     mutated mid-pass, and every later render and callback reads the settled state. */

  /** The approval library holding this document. Falls back to the normal one if HC never resolved. */
  const libTitleOf = (which: ApprovalLib): string =>
    which === "hc" ? cachedHcLibraries()?.approval.title ?? libraryTitle() : libraryTitle();
  const libSegOf = (which: ApprovalLib): string =>
    which === "hc"
      ? cachedHcLibraries()?.approval.urlSegment ?? libraryUrlSegment()
      : libraryUrlSegment();
  const libTitle = (): string => libTitleOf(lib);
  const libSeg = (): string => libSegOf(lib);
  const libTitleEnc = (): string => encodeURIComponent(libTitle());
  /** Where an approved document lands — the OTHER half of whichever pair this one belongs to. */
  /* Display only — it names the destination library in the two clash warnings. `documentsLibraryTitle()`
     rather than the literal since 2026-08-28: the client retitled it to `Restricted & Confidential
     Document`, and a warning naming a library the approver cannot find on their own site is worse
     than no warning. Still falls back to the literal when unresolved, which is the loud direction. */
  const approvedLibTitle = (): string =>
    lib === "hc"
      ? cachedHcLibraries()?.documents.title ?? documentsLibraryTitle()
      : documentsLibraryTitle();
  /**
   * The destination library's URL SEGMENT, for building a path.
   *
   * Deliberately UNLIKE `approvedLibTitle` above, which falls back to the normal library when the HC
   * pair is unresolved. That fallback is harmless for a display string and is the worst thing this
   * one could do: it would point an HC document's readiness check — and the message an approver acts
   * on — at the open library, where the whole unit can read. Blank instead, which `unitFolderPath`
   * refuses by name. Same reasoning as `hcRouting.ts` returning undefined rather than falling back.
   */
  const approvedLibSeg = (): string =>
    lib === "hc" ? cachedHcLibraries()?.documents.urlSegment ?? "" : DOCUMENTS_URL_SEGMENT;

  const getItemId = (): number | null => {
    const p = new URLSearchParams(window.location.search);
    const id = p.get("itemId");
    return id ? parseInt(id, 10) : null;
  };

  // Back link → the file's own folder in Staging (not the library root), so the
  // approver lands where the document lives instead of having to drill back in.
  const backUrl = (): string => {
    const base = `${webUrl}/${encodeURIComponent(libSeg())}/Forms/AllItems.aspx`;
    const fileRef = item?.File?.ServerRelativeUrl;
    if (!fileRef) return base;
    const folder = fileRef.slice(0, fileRef.lastIndexOf("/"));
    return `${base}?id=${encodeURIComponent(folder)}`;
  };

  /** Per-item metadata labels. Cannot be batched into the queue query — FieldValuesAsText is
   *  a per-item endpoint — so it is the one thing a swap has to wait for. */
  // `which` defaults to the settled state; loadItem passes it explicitly, because on the first pass
  // the library has only just been resolved and the state has not repainted yet.
  const loadFieldText = async (itemId: number, which: ApprovalLib = lib): Promise<void> => {
    try {
      const textRes = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libTitleOf(which))}')/items(${itemId})/FieldValuesAsText`,
        SPHttpClient.configurations.v1,
      );
      if (textRes.ok) setFieldText(await textRes.json() as IFieldText);
    } catch { /* labels are non-critical */ }
  };

  /**
   * Build the queue: every Pending item this account can read, oldest first.
   *
   * SCOPE COMES FROM PERMISSIONS, not from a filter. Inheritance is broken per unit folder in
   * Staging, so SharePoint security-trims the query and an approver gets exactly their unit's
   * pending files — with no filter logic here and nothing to keep in sync as units are added.
   * A Full Control account therefore sees every pending file on the site, which is what $top=200
   * is really guarding against: not correctness, just how much an admin pulls while testing.
   *
   * A failure leaves the queue empty, which renders the page exactly as it behaved before this
   * feature. The queue is a convenience on top of a page that already works; failing to build it
   * must never block the decision the approver came to make.
   */
  const loadQueue = async (current: IFileItem, which: ApprovalLib = lib): Promise<void> => {
    try {
      // THE QUEUE STAYS INSIDE ONE LIBRARY. Merging the two would walk an HC approver from a Highly
      // Confidential document straight into an ordinary one and back, and — worse — would show a
      // plain approver nothing while quietly counting HC documents they cannot open.
      const url =
        `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libTitleOf(which))}')/items` +
        `?$filter=OData__ModerationStatus%20eq%202` +
        `&$expand=File,Author` +
        `&$select=ID,FileLeafRef,OData__ModerationStatus,Created,Author/Title,File/Length,File/ServerRelativeUrl` +
        `&$orderby=Created%20asc&$top=200`;
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json() as { value: IFileItem[] };
      const { queue: entries, index } = buildQueue(data.value ?? [], current);
      setQueue(entries);
      setPos(index);
    } catch (err) {
      console.error("[ApprovalDoc] queue query failed — continuing as a single document:", err);
      setQueue([]);
    }
  };

  const loadItem = async (): Promise<void> => {
    const itemId = getItemId();
    if (!itemId) {
      setFetchError("No document ID provided. The link should include ?itemId=123.");
      setLoading(false);
      return;
    }
    try {
      const read = async (which: ApprovalLib): Promise<SPHttpClientResponse> =>
        context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libTitleOf(which))}')/items(${itemId})` +
            `?$expand=File,Author` +
            `&$select=ID,FileLeafRef,OData__ModerationStatus,Created,Author/Title,File/Length,File/ServerRelativeUrl`,
          SPHttpClient.configurations.v1,
        );
      let where: ApprovalLib = lib;
      let res = await read(where);
      // Not in the library the link implied. On a site with Highly Confidential that is the ordinary
      // case for an HC document whose link carried no `lib=hc`, so try the other one before giving
      // up: an approver who cannot open the document cannot approve it, and an HC document that is
      // never approved never reaches the people it was filed for.
      if (!res.ok && where === "normal" && hcAvailable()) {
        const hcRes = await read("hc");
        if (hcRes.ok) {
          where = "hc";
          res = hcRes;
          setLib("hc");
        }
      }
      if (!res.ok) {
        const body = await res.text().catch(() => "");
        throw new Error(`HTTP ${res.status}: ${body.slice(0, 300)}`);
      }
      const data: IFileItem = await res.json();
      setItem(data);
      /* ⚠ THE LINK ONLY SPEAKS FOR A DOCUMENT THAT IS STILL PENDING. An already-decided one must
         show what actually happened to it — pre-ticking "Reject" on a document somebody approved
         an hour ago states something false on the one screen that is the record of the decision.
         The submit button is disabled for a decided document anyway, so this is about what the
         approver is TOLD, not about what they can do. */
      const current = statusToDecision(data.OData__ModerationStatus);
      setDecision(current === "Pending" ? linkDecision.current ?? current : current);
      await loadFieldText(itemId, where);
      // After the document, never before: a queue failure must not stop the page loading, and
      // the current item has to be known so it can be placed in the queue.
      await loadQueue(data, where);
    } catch (err) {
      setFetchError(err instanceof Error ? err.message : "Could not load this document.");
    } finally {
      setLoading(false);
    }
  };

  /**
   * Move to another position in the queue.
   *
   * Every piece of per-document state is reset here. `comments` is the one that matters beyond
   * cosmetics: carrying a comment onto the next document would attach one approver's reasoning
   * to a different document's audit trail.
   */
  const goTo = async (index: number): Promise<void> => {
    const entry = queue[index];
    if (!entry || index === pos) return;
    const reset = swapState(entry);
    setPos(index);
    setItem(entry.item);
    setDecision(reset.decision);
    setComments(reset.comments);
    setSubmitError(reset.submitError);
    setSubmitted(reset.submitted);
    setFieldText(reset.fieldText);
    // replaceState, not pushState: Back should return the approver to the Staging library they
    // came from, not walk them backwards through their own review session one document at a time.
    try {
      const u = new URL(window.location.href);
      u.searchParams.set("itemId", String(entry.item.ID));
      window.history.replaceState(undefined, "", u.toString());
    } catch { /* a failed URL rewrite is cosmetic — the page already shows the new document */ }
    setSwapping(true);
    try {
      await loadFieldText(entry.item.ID);
    } finally {
      setSwapping(false);
    }
  };

  // primeNames FIRST, and awaited: every read below addresses the library by title or by URL
  // segment, and both are wrong until it resolves. It never throws — an unresolvable name
  // leaves the legacy "Staging" pair rather than blocking the page.
  useEffect(() => {
    primeNames(context.spHttpClient, webUrl)
      .catch(() => undefined)
      .then(() => loadItem())
      .catch(() => undefined);
  }, []);

  /**
   * Each segment's permissioned tier count, for the approval guard.
   *
   * A SEPARATE effect from the one above, deliberately: it is not on the path to rendering the
   * document, so chaining it would delay the preview behind a read the approver only needs at the
   * moment they press Approve. It awaits `primeNames` INSIDE itself rather than relying on the other
   * effect having got there first — `cachedListTitle` answers the legacy `DMS Config` until priming
   * settles, which 404s on a renamed site, and that race emptied reconciliation's segment picker on
   * 2026-08-21. Priming is idempotent and cached, so awaiting it twice costs nothing.
   *
   * `SortOrder` is NOT selected: one unknown field name fails the whole request (gotcha #11) and
   * nothing here needs the order.
   */
  useEffect(() => {
    (async (): Promise<void> => {
      try {
        await primeNames(context.spHttpClient, webUrl).catch(() => undefined);
        const config = await listTitleEncoded(context.spHttpClient, webUrl, LIST_SUFFIX.config);
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${config}')/items` +
            `?$select=StagingFolder,Levels&$filter=ConfigType eq 'mode'&$top=200${bust()}`,
          SPHttpClient.configurations.v1,
          { headers: GET_FRESH },
        );
        if (!res.ok) {
          console.error("Approval guard: could not read the mode rows:", res.status);
          return;   // stays undefined — the guard refuses and says why
        }
        const rows = ((await res.json()).value ?? []) as Array<{ StagingFolder?: string; Levels?: string }>;
        const map: Record<string, number> = {};
        for (const r of rows) {
          const key = (r.StagingFolder ?? "").trim().toLowerCase();
          const n = permissionedTierCount(r.Levels);
          // A segment whose Levels could not be parsed is LEFT OUT rather than stored as 0: absent
          // reads as "unknown" downstream, which refuses; a stored 0 would claim we know it is flat.
          if (key.length > 0 && n !== undefined) map[key] = n;
        }
        setTierCounts(map);
      } catch (e) {
        console.error("Approval guard: mode rows unavailable", e);
      }
    })().catch(() => undefined);
  }, []);

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
    // Keyed on the segment FOLDER name, which IS the mode row's StagingFolder — so no term lookup is
    // needed and renaming a segment's LABEL cannot break the match.
    const segmentFolder = (() => {
      const prefix = `${webSru}/${libSeg()}/`.toLowerCase();
      if (stagingFileUrl.toLowerCase().indexOf(prefix) !== 0) return "";
      return stagingFileUrl.slice(prefix.length).split("/")[0] ?? "";
    })();
    // TEMPORARY DIAGNOSTIC (2026-09-01) — remove once the live refusal is explained. Every value
    // this guard's verdict depends on, in one line, so a live failure can be read off the console
    // instead of guessed at from the message alone.
    console.error("Approval guard diagnostic:", {
      stagingFileUrl,
      webSru,
      libSeg: libSeg(),
      prefix: `${webSru}/${libSeg()}/`.toLowerCase(),
      segmentFolder,
      lookupKey: segmentFolder.toLowerCase(),
      tierCounts,
      resolvedTierCount: tierCounts?.[segmentFolder.toLowerCase()],
    });
    return checkUnitFolderReady({
      sp: context.spHttpClient,
      webUrl,
      webSru,
      fileSru: stagingFileUrl,
      sourceSegment: libSeg(),
      destSegment: approvedLibSeg(),
      destLibTitle: approvedLibTitle(),
      permissionedTiers: tierCounts?.[segmentFolder.toLowerCase()],
    });
  };

  /**
   * Would approving this document REPLACE an existing one already in the destination library?
   *
   * Auto-route's `Copy file` step is configured to replace on a name clash (confirmed in the flow's
   * own Code view, 2026-08-24: `nameConflictBehavior: 1`) — so if a same-named file was filed into
   * the approved-side library after this one was uploaded, approving here silently overwrites it.
   * No error, no warning, a green run.
   *
   * `Form.tsx` already refuses a clash at UPLOAD time (`approvedClash`), but that check runs once,
   * before this document even exists in the approval library — it cannot see a name that lands in
   * Documents AFTER this upload and BEFORE this approval (a bulk import, or a second uploader).
   * This is the second half of that same protection, checked at the only other moment that matters:
   * immediately before the copy actually happens.
   *
   * ⚠ FAILS CLOSED, the OPPOSITE direction from `Form.tsx`'s upload-time check. That check runs on
   * every upload site-wide, so an unanswerable read there would take the WHOLE FORM out of service —
   * Form.tsx accepts a small risk of a missed clash to avoid that. This check runs once, at one
   * approval, and its neighbour `documentsUnitFolderReady` above already fails closed for the
   * identical reason: a retry costs the approver seconds, a wrong "proceed" overwrites a document
   * and nobody finds out. An inconclusive read must refuse here, not guess.
   *
   * The destination is the FULL file path with only the library segment swapped — NOT the unit
   * folder `documentsUnitFolderReady` computes, which deliberately stops at Unit. Everything below
   * Unit (Year/Document Type/Archive…) must be preserved exactly, because that is what Auto-route's
   * own `Compose_1` expression preserves when it builds the copy destination.
   */
  const documentsFileClash = async (
    stagingFileUrl: string,
    fileName: string,
    /* `GuardResult`, not a narrower literal. The local type used to spell out `{ ok, reason }`, which
       silently dropped `clash` - the flag that tells a confirmed collision apart from an unanswerable
       check, and therefore which of the two warnings the approver sees. */
  ): Promise<GuardResult> =>
    checkDestinationClash({
      sp: context.spHttpClient,
      webUrl,
      fileSru: stagingFileUrl,
      fileName,
      sourceSegment: libSeg(),
      destSegment: approvedLibSeg(),
    });

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
            `This unit's folder is not ready in the ${approvedLibTitle()} library, so the document was NOT approved (${ready.reason}). ` +
            `Ask an administrator to run Folder Reconciliation, then approve again.`,
          );
          // Leave `decision` as the approver chose it — the selection is still valid, it is
          // the destination that is not ready.
          setSubmitting(false);
          return;
        }
        // Second, SEPARATE check — the folder can be ready and still hold a same-named file that
        // this approval would silently replace. See `documentsFileClash` for why this cannot be
        // folded into the check above: it asks about a different folder (below-Unit, ensure-created
        // on demand) and fails closed for a different reason.
        const clash = await documentsFileClash(item.File.ServerRelativeUrl, item.FileLeafRef);
        if (!clash.ok) {
          /* ⚠ THIS NO LONGER REFUSES (client, 2026-08-28: *"just drop the guard and let them
             override … crosscheck and notify them its going to be overwritten and let them do it,
             because the overwritten is needed"*), AND NO LONGER ASKS A SECOND TIME EITHER
             (client, 2026-09-02: *"there is no need for two popup as the one we made is enough"*).

             The warning is shown BEFORE this point — under the comment box, from the `clashCheck`
             effect above, computed the moment the document loaded. By the time an approver reaches
             Approve they have already seen it; asking again here with a native `window.confirm()`
             was the redundant second popup. This re-check still runs (the same reasoning as before:
             a name that landed AFTER the page loaded and BEFORE this click must still be caught), it
             just no longer interrupts — it proceeds and, on a confirmed clash, does the bookkeeping
             below.

             ⚠ TWO DIFFERENT MEANINGS, because `ok: false` carries two of them. `clash.clash` marks a
             CONFIRMED collision, where approving really does replace a document. Every other refusal
             means the check could not be answered. Only the confirmed case marks a record replaced,
             below — the other lets the approval proceed with no consequence-tracking to fake.

             ⚠ THE PREVIOUS CONTENT DOES NOT SURVIVE AS A VERSION, AND THIS COMMENT SAID IT DID
             UNTIL 2026-08-31. Auto-route's `Copy file` with `nameConflictBehavior: 1` DELETES the
             destination item and creates a new one, so the filed document comes back as a fresh
             `1.0` holding the new content — proven on site: a 227.3 KB document by one uploader
             came back as a 117.4 KB one by another, with a single 1.0 in version history. Both
             approved-side libraries keep 500 major versions and it makes no difference, because
             version history is never reached. What DOES survive is the recycle bin — the deleted
             item sits there for 93 days with the ORIGINAL uploader still in Created By. */

          /* ── The record this approval is about to displace (2026-08-28) ──────
             Client: *"The older submission under My Submission will change to Cancelled status if
             replaced."* The upload form marks its own overwrites; this is the OTHER replacement
             path, and without it a document replaced by an approval reads as **deleted** to whoever
             filed it — the same false "your work was destroyed" the record feature exists to avoid.

             ⚠ MARKED BEFORE THE REPLACEMENT ACTUALLY HAPPENS, and that is safe ONLY because of the
             precedence in `mergeRecords`: a record that still RESOLVES wins as `live` regardless of
             `replacedAt`. Auto-route does the copy minutes later, so until it runs the old file is
             still there, still carries its stamp, and still reads live. The moment it is replaced,
             the stamp is gone and the row turns Cancelled. It self-corrects, and it also means an
             approval that Auto-route never completes leaves nothing wrongly marked.

             ⚠ CONFIRMED CLASHES ONLY (`clash.clash === true`). The other branch of this warning is
             "the check could not be answered" — marking a record replaced on the strength of a
             check that found nothing would assert a replacement that may never happen.

             Non-blocking and non-throwing throughout: the approval is the thing that matters, and
             `markRecordReplaced` logs its own failures. */
          if (clash.clash === true) {
            try {
              const destSeg = approvedLibSeg();
              const destFull = destSeg
                ? swapLibrarySegment(item.File.ServerRelativeUrl, libSeg(), destSeg)
                : undefined;
              if (destFull) {
                const priorRes = await context.spHttpClient.get(
                  `${webUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields` +
                    `?$select=SubmissionFileId&@f='${destFull.split("/").map(encodeURIComponent).join("/").replace(/'/g, "''")}'`,
                  SPHttpClient.configurations.v1,
                  { headers: { Accept: "application/json;odata=nometadata" } },
                );
                if (priorRes.ok) {
                  const prior = await priorRes.json();
                  const stamp =
                    typeof prior?.SubmissionFileId === "string" ? prior.SubmissionFileId.trim() : "";
                  if (stamp.length > 0) {
                    await markRecordReplaced(
                      context.spHttpClient,
                      webUrl,
                      stamp,
                      (context.pageContext.user.email ?? "").toLowerCase(),
                    );
                  }
                }
              }
            } catch {
              /* A bookkeeping read must never be why an approval fails. Worst case the displaced
                 record reads `deleted`, which is exactly what it did before this existed. */
            }
          }
        }
      }
      const digest      = await getDigest();
      const safeComment = comments.replace(/'/g, "''");
      // Path passed as an OData parameter alias (@f) appended to the query
      // string, not embedded inline in the URL path — inline literals hit
      // IIS's maxUrlLength once the server-relative path gets long/deep
      // (nested subcategory nesting can add up fast). See sp-rest-alias-vs-inline-literal-400-error memory.
      const safeUrl     = item.File.ServerRelativeUrl.replace(/'/g, "''");
      const itemBase    = `${webUrl}/_api/web/lists/getbytitle('${libTitleEnc()}')/items(${item.ID})`;
      const headers     = { "X-RequestDigest": digest, Accept: "application/json;odata=nometadata" };

      if (action === "Approved") {
        const res = await context.spHttpClient.post(
          `${webUrl}/_api/web/getfilebyserverrelativeurl(@f)/approve(comment='${safeComment}')?@f='${safeUrl}'`,
          SPHttpClient.configurations.v1,
          { headers },
        );
        if (!res.ok) throw new Error(`Approve returned ${res.status}: ${await res.text().catch(() => "")}`);
        /* ⚠ `File.approve()` DOES NOT RESTAMP `Editor` — proven live 2026-09-01: after
           clarencechojinheng approved a document uploaded by chocheetuck4, `Editor` still read
           chocheetuck4. Approving is a moderation-status change, not an edit, so SharePoint has no
           built-in "approved by" field and `Editor` is not a stand-in for one. Without this, the
           audit flow names the uploader as the approver, and Auto-route's self-approval suppression
           (which compares Author to Editor) matches on EVERY approval, silencing the notification
           email for everyone.
           A SEPARATE write, after the approval has already succeeded: a failure here must never make
           an approver believe their decision did not go through, the same reasoning as the
           `markRecordReplaced` call above. Guarded by `libraryHasColumns` so a site where
           reconciliation has not yet created the column degrades to today's behaviour — wrong actor,
           no email — rather than a failed approval. */
        try {
          if (await libraryHasColumns(context.spHttpClient, webUrl, libTitle(), [APPROVED_BY_COLUMN])) {
            const approverEmail = (context.pageContext.user.email ?? "").toLowerCase();
            // ⚠ SHAREPOINT REJECTS A MERGE THAT SETS OData__ModerationStatus ALONGSIDE ANY OTHER
            // FIELD — proven live 2026-09-01 via a 500 reading "You cannot change moderation status
            // and set other item properties at that same time." The earlier "fix" that combined them
            // into one MERGE could therefore never succeed; it 500'd silently on every approval, so
            // ApprovedBy was never written while approval/routing kept working (approve() below had
            // already set the status independently).
            //
            // So the two writes must stay SEPARATE: this MERGE sets ApprovedBy alone (an ordinary
            // field edit, no moderation status touched, so it does not hit the restriction above) —
            // and because ANY edit to a moderated item reverts its status to Pending unless that same
            // write re-asserts it, we then call approve() again immediately after to restore Approved.
            // approve() is a dedicated method, not a field MERGE, so IT doesn't hit the restriction
            // either. Net effect: ApprovedBy lands, and the item ends up Approved regardless of order.
            await context.spHttpClient.fetch(itemBase, SPHttpClient.configurations.v1, {
              method: "POST",
              headers: { ...headers, "X-HTTP-Method": "MERGE", "IF-MATCH": "*", "Content-Type": "application/json;odata=nometadata" },
              body: JSON.stringify({ ApprovedBy: approverEmail }),
            });
            await context.spHttpClient.post(
              `${webUrl}/_api/web/getfilebyserverrelativeurl(@f)/approve(comment='${safeComment}')?@f='${safeUrl}'`,
              SPHttpClient.configurations.v1,
              { headers },
            );
          }
        } catch {
          /* Same rule as the record-replacement bookkeeping above: the approval already succeeded,
             so a failed stamp must never be reported as a failed approval. */
        }
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
          `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(approvedLibTitle())}')/items?$filter=FileLeafRef eq '${safeLeaf}'&$select=FileRef&$top=1`,
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
          `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(approvedLibTitle())}')/items?$filter=FileLeafRef eq '${safeLeaf}'&$select=FileRef&$top=1`,
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
      // Mark this position decided so it stays navigable but cannot be resubmitted, and so
      // "Approve another file" can find the next piece of real work.
      setQueue((q) => q.map((e, i) => (i === pos ? { ...e, decided: action } : e)));
      setSubmitted(action);
    } catch (err) {
      console.error("[ApprovalDoc] submitDecision failed:", err);
      setSubmitError(err instanceof Error ? err.message : "Could not submit your decision. Please try again.");
    } finally {
      setSubmitting(false);
    }
  };

  /* ⚠ THIS HOOK MUST STAY ABOVE THE RENDER GUARDS BELOW.
     Placed after them (as it was on 2026-08-25, for about an hour) the component returns early
     while `loading` is true, the hook never runs on that render, and the moment loading finishes
     React sees more hooks than last time and throws "Rendered more hooks than during the previous
     render". The component then renders NOTHING — a blank web part, no error state, which is the
     one outcome this file's guards exist to prevent. Same rule the share-recipient picker follows
     in MySubmissions. */
  /**
   * May this user approve THIS document? Answered by `checkApproveRight` in shared/approvalGuards,
   * the same call the bulk approve command set makes — so the panel here and the command's
   * visibility there can never disagree about who may approve what.
   */
  useEffect(() => {
    if (!item) { setApproveRight("unknown"); return undefined; }
    let cancelled = false;
    checkApproveRight({
      sp: context.spHttpClient,
      webUrl,
      listTitle: libTitle(),
      itemId: item.ID,
    })
      .then((v) => { if (!cancelled) setApproveRight(v); })
      .catch(() => { if (!cancelled) setApproveRight("unknown"); });
    return () => { cancelled = true; };
  }, [item === null ? 0 : item.ID, lib]);

  /**
   * Precheck the existing-document clash as soon as the document loads, so the comment box already
   * carries the warning before Approve is ever clicked — see `clashCheck`'s own comment above for why.
   * MUST stay above the render guards below, same reason as the `approveRight` effect immediately
   * above it.
   */
  useEffect(() => {
    if (!item) { setClashCheck(undefined); return undefined; }
    let cancelled = false;
    documentsFileClash(item.File.ServerRelativeUrl, item.FileLeafRef)
      .then((r) => { if (!cancelled) setClashCheck(r); })
      .catch(() => { if (!cancelled) setClashCheck(undefined); });
    return () => { cancelled = true; };
  }, [item === null ? 0 : item.ID, lib]);

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

  /* ⚠ NEVER `return null` HERE — that rendered a BLANK PAGE, reported on site 2026-08-20.
     Approving the last document in the queue is the ordinary way to reach this state: Auto-route
     copies the file to Documents and DELETES it from the approval library, so the `?itemId=` the
     approver came in on no longer resolves. Loading has finished, no request errored, and there is
     simply no item — and an approver pressing Back after every approval met an empty screen with
     nothing on it to explain itself or to click.
     A page that renders nothing is the worst of both states this codebase distinguishes everywhere
     else: it says neither "could not read" nor "nothing here". */
  if (!item) {
    return (
      <div style={{ padding: 32, fontSize: 14, color: "#323130", maxWidth: 640, lineHeight: 1.6 }}>
        <div style={{ fontWeight: 600, marginBottom: 8 }}>This document is no longer in the approval queue.</div>
        <div style={{ color: "#605e5c" }}>
          If you have just approved it, that is expected — approved documents are moved into the
          Documents library and no longer appear here. A rejected document stays in the approval
          library, so you would still see it.
          {" "}Otherwise the link may point at a document that has since been moved or deleted.
        </div>
      </div>
    );
  }

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
    const nextIdx = nextUndecidedIndex(queue, pos);
    return (
      <div style={s.popupOverlay} role="dialog" aria-modal="true">
        <div style={s.popupCard}>
          <div style={{ marginBottom: 20 }}>{icon}</div>
          <div style={s.popupTitle}>{title}</div>
          <div style={s.popupMsg}>
            {message}
            {/* Only when the queue exists and is finished. Saying "that was the last one" when the
                queue query failed would be a claim we cannot support. */}
            {queue.length > 0 && nextIdx === -1 && (
              <div style={{ marginTop: 10, fontWeight: 600, color: "#0f6c3f" }}>
                That was the last document waiting for your approval.
              </div>
            )}
          </div>
          {/* Primary advances to the next UNDECIDED item, whereas Prev/Next move by position.
              Deliberate: after deciding, the approver wants the next piece of work; while
              browsing, they want the next document in the list. */}
          <div style={s.popupBtnRow}>
            {nextIdx !== -1 && (
              <button style={s.popupBtn} onClick={() => { goTo(nextIdx).catch(() => undefined); }}>
                Approve another file
              </button>
            )}
            <button
              style={nextIdx !== -1 ? { ...s.popupBtn, ...s.popupBtnGhost } : s.popupBtn}
              onClick={() => { window.location.href = backUrl(); }}
            >
              Back to library
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ── Main ───────────────────────────────────────────────────────────────────

  // Whether the position on screen has already been decided in this session. Drives the badge
  // and disables Ok, so stepping back to confirm a decision cannot resubmit it.
  const currentDecided: Decision | null = queue[pos]?.decided ?? null;

  // Per-type preview. PDFs used to be the only thing that rendered, because a raw file URL in an
  // iframe is a PDF viewer and nothing else — SharePoint serves an Office document as a download,
  // so the pane silently stayed blank. See src/shared/filePreview.ts.
  const preview = previewTarget(
    item.FileLeafRef,
    item.File.ServerRelativeUrl,
    webUrl.split('/sites/')[0],
    webUrl,
  );

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
    const after = item.File.ServerRelativeUrl.split(`/${libSeg()}/`)[1];
    if (!after) return "—";
    const parts = after.split("/");
    const org = parts.slice(0, Math.max(0, parts.length - 3));
    return org.length ? org.join(" › ") : "—";
  })();

  // Every column an approver is deciding on, in the same order as the Staging
  // views (memory `dms-staging-column-order`): where it is, then what it is,
  // then who it is for, then how sensitive, then the free text.
  //
  // FieldValuesAsText returns EVERY field on the item, so nothing here needs a
  // wider $select — only the right response key. It double-encodes underscores,
  // hence the _x005f_ form first with the plain internal name as a fallback.
  //
  // Business Segment / Department / Unit are not redundant with Location:
  // Location is the folder path, which is ABBREVIATED (GHO › GCA › EG), while
  // these three carry the terms' full labels. An approver needs the full label
  // to be sure which unit they are publishing to.
  const metadata: [string, string][] = [
    ["Location",             orgLocation],
    ["Business Segment",     pick("Business_x005f_x0020_x005f_Segment", "Business_x0020_Segment")],
    ["Department",           pick("Department")],
    ["Unit",                 pick("Unit")],
    ["Document Type",        pick("Document_x005f_x0020_x005f_Type", "Document_x0020_Type")],
    ["Year",                 pick("Year", "Year_x005f_x002f_x005f_Period", "Year_x002f_Period")],
    // Already formatted to the site locale by SharePoint — displayed as returned,
    // never re-parsed. Parsing it here would reintroduce the M/D/YYYY trap.
    ["Document Date",        pick("DocumentDate")],
    // ProjectName has no encoded characters, so its response key is unencoded.
    ["Project Name",         pick("ProjectName")],
    ["Vendor/Customer Name", pick("Vendor_x005f_x002f_x005f_CustomerName", "Vendor_x002f_CustomerName")],
    ["Confidential Level",   pick("Confidentiality_x005f_x0020_x005f_Level", "Confidentiality_x0020_Level")],
    // Yes/No comes back as the words, not a boolean — so a false reads "No"
    // rather than blank, which is the whole point on a privilege flag.
    ["Legally Privileged",   pick("LegallyPrivileged")],
    // The built-in Description. Its leading underscore is encoded too.
    ["Details",              pick("_x005f_ExtendedDescription", "_ExtendedDescription")],
    ["Remark",               pick("Remark")],
  ];

  // Pending is a system state only — approvers pick Approved or Rejected.
  /* ⚠ THE LABEL IS THE ACTION; THE `val` IS THE STORED DECISION, and they are deliberately
     different since 2026-08-30. The client asked for "Approve"/"Reject" — an imperative, which is
     what a control the approver is about to press should say. `Decision` is still `Approved` /
     `Rejected`, because that is what is written to the item and read back everywhere else; renaming
     the value would be a data change dressed as a copy change. */
  const radioOptions: { val: Decision; label: string }[] = [
    { val: "Approved", label: "Approve" },
    { val: "Rejected", label: "Reject" },
  ];

  return (
    <div style={s.root}>

      <div style={s.backBand}>
        <a href={backUrl()} style={s.backLink}>
          {/* A bare "<" must be escaped as an expression — JSX reads a literal
              left angle bracket in children as the start of a tag. */}
          <span style={{ fontSize: 16, lineHeight: 1, fontWeight: 700 }}>{"<"}</span>
          Back to document list
        </a>
      </div>

      <h1 style={s.docTitle}>{item.FileLeafRef}</h1>
      <div style={s.docMeta}>
        <img src={`${webUrl}/_layouts/15/images/ic${(item.FileLeafRef.split('.').pop() ?? 'txt').toLowerCase()}.png`} width={16} height={16} alt="" style={{ flexShrink: 0 }} />
        <span>{(item.FileLeafRef.split('.').pop() ?? 'File').toUpperCase()} document</span>
        <span style={s.dot}>•</span>
        <span>{formatSize(item.File.Length)}</span>
        <span style={s.dot}>•</span>
        <span>Uploaded on {formatDate(item.Created)}</span>
      </div>

      {/* The third track is DROPPED, not just emptied, when the approval panel is hidden.
          `s.grid` reserves 280px for it, so leaving the template alone left a blank column and a
          preview that stopped short of the page (client, 2026-08-25: "can you extend the preview, so
          empty on the right"). Overridden here rather than in `s.grid` because the style object is
          shared and static; this is the one thing about the layout that depends on state. */}
      <div
        style={{
          ...s.grid,
          gridTemplateColumns:
            approveRight === "denied" ? "196px minmax(0, 1fr)" : "196px minmax(0, 1fr) 280px",
        }}
      >

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
          {/* Muted rather than blanked during a swap: the labels come from a per-item endpoint
              that cannot be batched into the queue query, so they are always a moment behind the
              document itself. Hiding the rows would collapse the column and shift the layout on
              every Next press. */}
          <div style={swapping ? { opacity: 0.45, transition: "opacity .15s" } : undefined}>
            {metadata.map(([label, value]) => (
              <div key={label} style={s.detailRow}>
                <div style={s.metaLabel}>{label}</div>
                <div style={s.metaValue}>{value}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Center — Preview */}
        <div>
          {/* Queue navigation. Rendered only when a queue exists — with a single document, or
              after a failed queue query, the page keeps its original bare "Preview" heading. */}
          {queue.length > 1 ? (
            <div style={s.navRow}>
              <div style={s.sectionTitle}>Preview</div>
              <div style={s.navControls}>
                <button
                  style={{ ...s.navBtn, ...(pos === 0 || swapping ? s.navBtnOff : {}) }}
                  disabled={pos === 0 || swapping}
                  onClick={() => { goTo(pos - 1).catch(() => undefined); }}
                >
                  ‹ Prev
                </button>
                {/* Position in the WHOLE queue, decided items included. A number that moved
                    backwards as you worked would be worse than no number at all. */}
                <span style={s.navCount}>{pos + 1} of {queue.length}</span>
                <button
                  style={{ ...s.navBtn, ...(pos >= queue.length - 1 || swapping ? s.navBtnOff : {}) }}
                  disabled={pos >= queue.length - 1 || swapping}
                  onClick={() => { goTo(pos + 1).catch(() => undefined); }}
                >
                  Next ›
                </button>
                {currentDecided && (
                  <span style={{ ...s.navBadge, ...(currentDecided === "Approved" ? s.navBadgeOk : s.navBadgeNo) }}>
                    {currentDecided === "Approved" ? "✓ Approved" : "✕ Rejected"}
                  </span>
                )}
              </div>
            </div>
          ) : (
            <div style={s.sectionTitle}>Preview</div>
          )}
          {preview.kind === "image" ? (
            // An <img>, not an iframe, so the image can be sized directly.
            //
            // Fit to WIDTH and scroll — the same thing #view=FitH already does for PDFs, and for the
            // same reason. The first version fitted BOTH dimensions (`maxHeight: 100%` +
            // `objectFit: contain`), which is correct for a landscape photo and useless for the
            // common case: a full-page screenshot is far taller than it is wide, so fitting its
            // HEIGHT into the pane shrank its width to a sliver and the approver could read nothing.
            //
            // `maxWidth: 100%` with an auto height also never UPSCALES — a small image keeps its
            // natural size rather than being blown up blurry to fill the pane.
            <div>
              <div style={s.imageBox}>
                <img
                  src={preview.url}
                  alt={item.FileLeafRef}
                  style={{ maxWidth: "100%", height: "auto", display: "block" }}
                />
              </div>
              {/* ⚠ "Open in a new tab" REMOVED HERE (client, 2026-09-03: "Just remove this"), on the
                  IMAGE branch only — the preview already renders the whole document, so this was a
                  supplementary fallback link, never the only way to see the file. The "none" branch's
                  link below is untouched: there the preview genuinely could not render anything and
                  the link is the sole route to the document, not decoration. */}
            </div>
          ) : preview.kind === "none" ? (
            // Say so, and offer the file. A blank pane reads as a broken page, and an approver
            // who cannot see the document must not be nudged into deciding anyway.
            <div style={{ ...s.previewBox, flexDirection: "column", gap: 12, color: "#605e5c", fontSize: 14 }}>
              <div>
                No preview is available for <strong>{item.FileLeafRef}</strong>.
              </div>
              <a href={preview.openUrl} target="_blank" rel="noopener noreferrer" style={s.previewLink}>
                Open the file in a new tab
              </a>
            </div>
          ) : (
            <div>
              <iframe
                src={preview.url}
                style={{ width: "100%", height: "calc(100vh - 240px)", minHeight: 600, border: "none", borderRadius: 4, display: "block" }}
                title={`Preview of ${item.FileLeafRef}`}
                allowFullScreen
              />
              {/* ⚠ "Open in a new tab" REMOVED HERE TOO (client, 2026-09-03: "Just remove this"),
                  same reasoning as the image branch above. COST, STATED PLAINLY: this used to be the
                  approver's way through when Office Online refuses to render a file (a macro-enabled
                  workbook, a document open for editing elsewhere) without leaving the queue — that
                  fallback route is now gone. If Office Online preview failures start being reported,
                  this is why. */}
            </div>
          )}
        </div>

        {/* Right — Approval panel.
            HIDDEN ENTIRELY when this user cannot approve this document (client, 2026-08-25: "I think
            its best to remove the card"). An explanatory card was built first and rejected on sight:
            for someone who is only ever an uploader here, a panel headed "Approval" is a control they
            can never use, and explaining that on every visit is noise rather than help. What they
            came for — the document, its metadata, its status — is all on the left.

            `denied` ONLY. `unknown` keeps the panel, because a failed probe must never take the
            approval controls away from a real approver; the 403 at submit is still the backstop. */}
        {approveRight !== "denied" && (
        <div style={s.panel}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <span style={s.panelTitle}>Approval</span>
          </div>
          {/* ⚠ A DECIDED DOCUMENT GETS A READ-ONLY SUMMARY, NOT THE SAME FORM DISABLED (client's
              mockup, 2026-09-03: "having clickable radio buttons and CTA buttons ... displayed out
              in the layout is really confusing"). Before this, the panel kept showing live-looking
              radio buttons, a comment box and Approve/Cancel buttons after a decision had already
              been made in this session — every control merely `disabled`, which still READS as an
              interactive form. Status / Comment / who-and-when is the whole story once a decision
              exists; nothing below it can be acted on again. */}
          {currentDecided ? (
            <>
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: "#605e5c", marginBottom: 4 }}>STATUS</div>
                <span style={{ ...s.navBadge, ...(currentDecided === "Approved" ? s.navBadgeOk : s.navBadgeNo) }}>
                  {currentDecided === "Approved" ? "Approved" : "Rejected"}
                </span>
              </div>
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: "#605e5c", marginBottom: 4 }}>COMMENT</div>
                <div style={{ fontSize: 13 }}>{comments.trim() || <em style={{ color: "#a19f9d" }}>No comment was given.</em>}</div>
              </div>
              <div style={{ borderTop: "1px solid #edebe9", paddingTop: 10, fontSize: 12, color: "#605e5c" }}>
                Reviewed by <strong>{context.pageContext.user.displayName}</strong> on{" "}
                {formatDate(new Date().toISOString())} &mdash; no further action is needed.
              </div>
            </>
          ) : (
            <>
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
                {/* ⚠ REPLACES A NATIVE `window.confirm()` THAT USED TO FIRE ON APPROVE (client,
                    2026-09-02: *"the native one is annoying to client … its best to just add a text
                    under the comment box, precheck the file against the Document Library"*). Same
                    two messages the confirm used to carry — confirmed clash vs. an unanswerable check —
                    shown here instead, so Approve is a single click with nothing left to ask again. */}
                {clashCheck && !clashCheck.ok && (
                  <div style={{ ...s.charCount, textAlign: "left" as const, color: "#8a4b00", marginTop: 6 }}>
                    {clashCheck.clash === true
                      ? `A document called "${item.FileLeafRef}" is already filed in ${approvedLibTitle()} ` +
                        `for this folder. Approving REPLACES it — the document it replaces moves to the ` +
                        `site recycle bin, restorable for 93 days.`
                      : `The existing-document check could not be completed: ${clashCheck.reason}. ` +
                        `Approving may replace a document already filed in ${approvedLibTitle()}, or may ` +
                        `not — that could not be established.`}
                  </div>
                )}
              </div>

              <div style={{ marginBottom: 20 }}>
                <div style={s.publishLabel}>Publish to</div>
                <div style={{ display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap" as const }}>
                  {(() => {
                    // Segment › … › Unit, from the live Staging path (segment-agnostic).
                    const after = item.File.ServerRelativeUrl.split(`/${libSeg()}/`)[1];
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
                disabled={submitting || swapping || decision === "Pending"}
                style={{
                  ...s.btnApprove,
                  opacity: submitting || swapping || decision === "Pending" ? 0.7 : 1,
                  cursor: decision === "Pending" ? "not-allowed" : "pointer",
                }}
              >
                {submitting ? "Saving…" : "Proceed"}
              </button>
              <button onClick={() => { window.location.href = backUrl(); }} disabled={submitting} style={{ ...s.btnSendBack, opacity: submitting ? 0.7 : 1 }}>
                Cancel
              </button>
            </>
          )}
        </div>
        )}

      </div>
    </div>
  );
};

export default ApprovalDocument;
