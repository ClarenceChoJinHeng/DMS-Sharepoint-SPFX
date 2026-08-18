import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IApprovalDocumentProps } from "./IApprovalDocumentProps";
import {
  buildQueue,
  nextUndecidedIndex,
  statusToDecision,
  swapState,
  Decision,
  QueueEntry as QueueEntryOf,
} from "../../../shared/approvalQueue";
import { previewTarget } from "../../../shared/filePreview";
import { cachedHcLibraries, hcAvailable, libraryTitle, libraryUrlSegment } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

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
  root:        { fontFamily: "'Segoe UI', Tahoma, sans-serif", color: "#323130", background: "#fff", padding: "0 0 40px" } as React.CSSProperties,
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
  const [loading, setLoading]     = useState(true);
  const [decision, setDecision]   = useState<Decision>("Approved");
  const [comments, setComments]   = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState<Decision | null>(null);
  const [fetchError, setFetchError] = useState("");
  const [submitError, setSubmitError] = useState("");
  const [fieldText, setFieldText]   = useState<IFieldText>({});

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
  const approvedLibTitle = (): string =>
    lib === "hc" ? cachedHcLibraries()?.documents.title ?? "Documents" : "Documents";

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
      setDecision(statusToDecision(data.OData__ModerationStatus));
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
    const prefix = `${webSru}/${libSeg()}/`;
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
    // 404 covers two cases and SharePoint does not separate them: the folder is
    // not there, or the caller cannot even resolve it — which for an approver
    // means they lack library-level Read on Documents (the DMS_SITE_MEMBERS
    // grant). Both refuse, but they have different fixes, so name both.
    if (res.status === 404) {
      return { ok: false, reason: "it does not exist yet, or your account has no access to the Documents library at all" };
    }
    if (!res.ok) {
      console.error("Approval destination check failed:", res.status, unitDocs);
      return { ok: false, reason: `it could not be verified (HTTP ${res.status})` };
    }
    const d = await res.json();

    // Three distinct outcomes that this once collapsed into one message.
    //
    // `{"odata.null": true}` — the folder RESOLVED (or we would be in the 404
    // branch above) but its list item was security-trimmed away. This is a PASS,
    // and the reasoning is worth spelling out because the obvious reading is the
    // opposite one.
    //
    // An approver holds `_APR`, which is Staging-only by the isolation rule
    // (LIBRARY_ROLES in FolderManager.tsx). Their Documents access comes from
    // DMS_SITE_MEMBERS, which holds Read at LIBRARY level. So:
    //
    //   • Locked properly — reconciliation broke inheritance with
    //     copyRoleAssignments=false and granted only site Owners + the unit's
    //     MEMBER group. The library grant does not reach it, the approver is not
    //     on it, the item is trimmed → null. SAFE.
    //   • Created by Auto-route and still INHERITING — the library-level Read
    //     flows straight down, so the item IS readable and returns
    //     HasUniqueRoleAssignments: false. DANGEROUS, caught below.
    //   • Missing — GetFolderByServerRelativeUrl 404s. DANGEROUS, caught above.
    //
    // The dangerous state is the VISIBLE one, because inheritance is precisely
    // what makes it visible. So a trimmed item on a folder that resolved is
    // positive evidence of unique permissions excluding the caller — proof of
    // locking, not absence of proof.
    //
    // Previous versions got this backwards twice: first reporting "not locked
    // down" (a false claim about a correctly locked folder), then "missing or
    // invisible" — which blocked every correctly provisioned approver and implied
    // the fix was to add them to the unit's MEMBER group. That would have widened
    // Documents access for every approver on the site to work around a bug here.
    //
    // Rests on one assumption: the caller can resolve the folder at all, which
    // needs library-level Read from DMS_SITE_MEMBERS. Without it they 404 and are
    // refused — safe, and the 404 message names that cause.
    if (d === null || d["odata.null"] === true) {
      return { ok: true };
    }
    // Property genuinely absent, as opposed to false: the permissions could not be
    // evaluated from this account. Still refuses — a guard that cannot see must
    // not wave things through — but it must not claim to know what it did not read.
    if (typeof d.HasUniqueRoleAssignments !== "boolean") {
      console.error("Approval destination check: HasUniqueRoleAssignments not readable", d);
      return { ok: false, reason: "its permissions could not be read from your account" };
    }
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
      const itemBase    = `${webUrl}/_api/web/lists/getbytitle('${libTitleEnc()}')/items(${item.ID})`;
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
  const radioOptions: { val: Decision; label: string }[] = [
    { val: "Approved", label: "Approved" },
    { val: "Rejected", label: "Rejected" },
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
              {/* Offered here too, not just on the iframe branch. Fit-to-width answers the common
                  case; a very large scan still needs the browser's own zoom, and that lives in a
                  tab of its own. */}
              <div style={{ marginTop: 6, textAlign: "right" }}>
                <a href={preview.fileUrl} target="_blank" rel="noopener noreferrer" style={s.previewLink}>
                  Open in a new tab
                </a>
              </div>
            </div>
          ) : preview.kind === "none" ? (
            // Say so, and offer the file. A blank pane reads as a broken page, and an approver
            // who cannot see the document must not be nudged into deciding anyway.
            <div style={{ ...s.previewBox, flexDirection: "column", gap: 12, color: "#605e5c", fontSize: 14 }}>
              <div>
                No preview is available for <strong>{item.FileLeafRef}</strong>.
              </div>
              <a href={preview.fileUrl} target="_blank" rel="noopener noreferrer" style={s.previewLink}>
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
              {/* Always offered, whatever the kind. Office Online occasionally refuses a file it
                  cannot render — a macro-enabled workbook, a document open for editing elsewhere —
                  and the approver needs a way through that does not involve leaving the queue. */}
              <div style={{ marginTop: 6, textAlign: "right" }}>
                <a href={preview.fileUrl} target="_blank" rel="noopener noreferrer" style={s.previewLink}>
                  Open in a new tab
                </a>
              </div>
            </div>
          )}
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

          {/* Already decided in this session: the record exists, and resubmitting would fire the
              moderation and copy calls a second time. An approver can step back to confirm what
              they did — not to redo it. */}
          {currentDecided && (
            <div style={{ fontSize: 12, color: "#605e5c", marginBottom: 8 }}>
              You {currentDecided === "Approved" ? "approved" : "rejected"} this document in this
              session. Use Next to continue.
            </div>
          )}
          {/* Pending is no longer selectable — require an explicit Approved/Rejected choice. */}
          <button
            onClick={() => { submitDecision(decision).catch(() => undefined); }}
            disabled={submitting || swapping || decision === "Pending" || currentDecided !== null}
            style={{
              ...s.btnApprove,
              opacity: submitting || swapping || decision === "Pending" || currentDecided !== null ? 0.7 : 1,
              cursor: decision === "Pending" || currentDecided !== null ? "not-allowed" : "pointer",
            }}
          >
            {submitting ? "Saving…" : currentDecided ? "Already decided" : "Ok"}
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
