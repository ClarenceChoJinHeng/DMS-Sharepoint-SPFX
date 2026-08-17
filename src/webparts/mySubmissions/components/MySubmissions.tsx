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
// The metadata panel's rows. It DERIVES the tier rows from the item's own fields rather than naming
// them, which is what makes Region/Estate·Mill appear on a segment nobody wrote code for.
import { buildDetailRows, documentUnit, formatBytes, routeToApprover } from "../../../shared/documentDetails";
import {
  cachedHcLibraries,
  cachedListTitle,
  LIST_SUFFIX,
  libraryTitle,
  libraryUrlSegment,
} from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { normalizeRoleValue } from "../../../shared/groupMapModel";
// The request rules — validation, recipient parsing and the inside/outside test — live under test in
// shared/requests.ts and are shared with the approver's queue. Two copies of "what counts as external"
// is how one screen ends up permitting what the other refuses.
import {
  RequestDraft,
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
  // ── Requests ──
  askBar:   { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 16 },
  askBtn:   { padding: "6px 14px", fontSize: 12.5, fontFamily: "inherit", border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  askOff:   { padding: "6px 14px", fontSize: 12.5, fontFamily: "inherit", border: "1px solid #e6e6e6", borderRadius: 4, background: "#f4f4f4", color: "#9a9a9a", cursor: "not-allowed" },
  okBox:    { fontSize: 13, padding: "12px 14px", borderRadius: 6, border: "1px solid #b7dcc4", background: "#f3faf5", color: "#1c4d33", lineHeight: 1.55, marginBottom: 16 },
  warnBox:  { fontSize: 12.5, padding: "12px 14px", borderRadius: 6, border: "1px solid #f2c9a0", background: "#fff8f0", color: "#8a4b00", lineHeight: 1.55, marginBottom: 12 },
  modalBg:  { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 },
  modal:    { background: "#fff", borderRadius: 8, padding: 20, width: "min(560px, 94vw)", maxHeight: "86vh", overflowY: "auto", fontSize: 13 },
  label:    { display: "block", marginTop: 12, marginBottom: 4, fontSize: 12, fontWeight: 600, color: "#3b3a39" },
  field:    { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: 13, fontFamily: "inherit", border: "1px solid #c7c7c7", borderRadius: 4 },
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
  /** False only when the list answered 404. A different failure leaves this true and fails at the write. */
  listExists: boolean;
}

const BLANK_POLICY: RequestPolicy = {
  known: false, allowExternal: false, tenantDomains: [], approverUnits: [], listExists: true,
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

  /* ── Requests ──────────────────────────────────────────────────────────────
     An uploader cannot delete or share in Documents, so they ask and a Head of Unit decides.
     Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md */
  const [policy, setPolicy] = useState<RequestPolicy>(BLANK_POLICY);
  const [asking, setAsking] = useState<RequestType | undefined>(undefined);
  const [reason, setReason] = useState("");
  const [shareWith, setShareWith] = useState("");
  const [permission, setPermission] = useState<SharePermission>("View");
  const [expiresAt, setExpiresAt] = useState("");
  const [problems, setProblems] = useState<string[]>([]);
  const [sending, setSending] = useState(false);
  const [sent, setSent] = useState<string | undefined>(undefined);

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
        { headers: { Accept: "application/json;odata=nometadata" } },
      );

    let res = approvalLibrary
      ? await get(`${withStatus},OData__ModerationComments`)
      : await get(withStatus);
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
        { headers: { Accept: "application/json;odata=nometadata" } },
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
    /* Which library this row came from, by its URL segment.
       A two-way guess (`docsSegment ? Documents : approval`) sent every Highly Confidential row to
       the NORMAL approval library, which answers 404 for an item id it does not have — so the panel
       showed "no details were recorded" for exactly the documents whose metadata matters most. */
    const hc = cachedHcLibraries();
    const listTitle =
      row.library === docsSegment ? DOCUMENTS
      : hc && row.library === hc.approval.urlSegment ? hc.approval.title
      : hc && row.library === hc.documents.urlSegment ? hc.documents.title
      : libraryTitle();
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
    setAsking(undefined);
    setSent(undefined);
    setProblems([]);
    loadFieldText(row).catch(() => setFieldText({}));
  };

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

    // The signed-in user's own domain is not a guess — they are signed in to this tenant. A config
    // row adds more for a multi-domain tenant with no redeploy.
    const me = (context.pageContext.user.email ?? "").toLowerCase();
    const at = me.lastIndexOf("@");
    if (at > -1) next.tenantDomains.push(me.slice(at + 1));

    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Title,SettingValue&$filter=Title eq 'allowExternalSharing' or Title eq 'tenantDomains'&$top=20`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
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

    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items` +
          `?$select=Role,UnitTermGuid&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (res.ok) {
        for (const r of ((await res.json()).value ?? []) as Array<{ Role?: string; UnitTermGuid?: string }>) {
          /* APR ALONE since 2026-08-17 — see the same match in Requests.tsx. APRHC is retired, and
             `normalizeRoleValue` aliases a stored "APRHC" to "APR", so a legacy row still routes.

             normalizeRoleValue, NOT a raw toUpperCase: the Group Map's Role is often the LONG form
             ("APPROVER"), and a raw compare skips such a row — which would file the request against a
             unit nobody is recorded as approving for, silently. */
          const role = normalizeRoleValue(r.Role ?? "");
          if (role !== "APR") continue;
          const guid = (r.UnitTermGuid ?? "").trim();
          if (guid && next.approverUnits.indexOf(guid) === -1) next.approverUnits.push(guid);
        }
      }
    } catch {
      /* the request still routes on the document's own deepest tier */
    }

    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')/items?$select=Id&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      // ONLY a 404 means absent. Any other failure leaves it true, so the uploader is told the write
      // failed rather than being sent to an administrator over what may be a transient error.
      if (res.status === 404) next.listExists = false;
    } catch {
      /* leave it true — see above */
    }

    setPolicy(next);
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

    const draft: RequestDraft = {
      type,
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

      const res: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')/items`,
        SPHttpClient.configurations.v1,
        { headers: WRITE_HEADERS, body: JSON.stringify(body) },
      );
      if (!res.ok) {
        const text = await res.text().catch(() => "");
        throw new Error(
          res.status === 404
            ? "the requests list does not exist yet — ask an administrator to open the Requests page, which creates it"
            : `HTTP ${res.status} ${text.slice(0, 160)}`,
        );
      }

      writeAudit(context.spHttpClient, siteUrl, {
        event: type === "Deletion" ? EVENT.deletionRequested : EVENT.shareRequested,
        source: "MySubmissions",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: (context.pageContext.user.email ?? "").toLowerCase(),
        library: DOCUMENTS,
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
      /* The Highly Confidential pair, when the site has one.
         An HC uploader's files are invisible on this page without it, and "you have not uploaded
         anything yet" to someone who filed a Highly Confidential document last week is the worst
         possible answer — it invites them to upload it again.

         Read with the SAME AuthorId filter, so this adds no visibility whatsoever: a person sees
         their own HC files, and only where they already hold the grant. An uncleared uploader's
         read returns nothing because SharePoint refuses it, not because this code decided so. */
      const hc = cachedHcLibraries();
      let hcRows: Submission[] = [];
      if (hc) {
        // Each in its own try. An HC library the viewer cannot read at all is the NORMAL case for
        // everyone without clearance, and must never take the page down or blank the two libraries
        // that already loaded successfully.
        try {
          hcRows = hcRows.concat(await readLibrary(hc.approval.title, hc.approval.urlSegment, true, userId));
        } catch { /* not cleared, or unreachable — either way, nothing of theirs to show */ }
        try {
          hcRows = hcRows.concat(await readLibrary(hc.documents.title, hc.documents.urlSegment, false, userId));
        } catch { /* as above */ }
      }
      setRows(sortNewestFirst([...staging, ...documents, ...hcRows]));
      setLoadError(undefined);
      // Last, and never awaited into the same try: this decides whether the request buttons can be
      // offered, and nothing about it may cost an uploader the list of their own files.
      loadPolicy().catch(() => setPolicy({ ...BLANK_POLICY, known: false }));
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
  // Every library segment the trail builder must strip. Without the HC pair, an HC file's folder
  // trail would start with "HCApprovalDocument" — the library name presented as a folder.
  const hcLibs = cachedHcLibraries();
  const libs = [
    libraryUrlSegment(), docsSegment, DOCUMENTS,
    ...(hcLibs ? [hcLibs.approval.urlSegment, hcLibs.documents.urlSegment] : []),
  ];

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

    /* Why a request cannot be raised right now, or undefined when it can.
       A reason, never a disappeared button: an uploader who cannot see the control assumes the
       system does not do this and goes to ask someone in person. */
    const requestBlock =
      !policy.known
        ? "Checking whether requests can be raised…"
        : !policy.listExists
          ? "Requests are not set up on this site yet — an administrator opens the Requests page once to create the list."
          : !open.uniqueId
            ? "This file's id could not be read, so a request would not find it. Reload the page and try again."
            : undefined;

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

        {/* ── Asking for something to be done to this file ─────────────────────────
            Only for APPROVED files. Those live in Documents, where an uploader has Read and nothing
            more — so deleting or sharing one is a request. A pending or rejected file is still in the
            approval library, where a PIC holds Delete and can simply do it. Offering the same buttons
            there would teach people to ask permission for something they already control. */}
        {open.status === "Approved" && (
          <div style={s.askBar}>
            {sent === undefined && (
              <>
                <button
                  style={requestBlock === undefined ? s.askBtn : s.askOff}
                  disabled={requestBlock !== undefined}
                  title={requestBlock}
                  onClick={() => { setProblems([]); setAsking("Deletion"); }}
                >
                  Request deletion
                </button>
                <button
                  style={requestBlock === undefined ? s.askBtn : s.askOff}
                  disabled={requestBlock !== undefined}
                  title={requestBlock}
                  onClick={() => { setProblems([]); setAsking("Share"); }}
                >
                  Request share
                </button>
              </>
            )}
            <span style={{ fontSize: 12, color: "#605e5c" }}>
              {requestBlock ?? "Your Head of Unit decides these."}
            </span>
          </div>
        )}

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

        {asking !== undefined && (
          <div style={s.modalBg} onClick={() => { if (!sending) setAsking(undefined); }}>
            <div style={s.modal} onClick={(e) => e.stopPropagation()}>
              <p style={{ fontSize: 16, fontWeight: 600, margin: "0 0 6px" }}>
                {asking === "Deletion" ? "Ask for this file to be deleted" : "Ask for this file to be shared"}
              </p>
              <p style={{ fontSize: 12.5, color: "#605e5c", margin: "0 0 4px", lineHeight: 1.5 }}>
                {open.name}
              </p>
              <p style={{ fontSize: 12.5, color: "#605e5c", margin: "0 0 8px", lineHeight: 1.5 }}>
                {asking === "Deletion"
                  ? "Nothing happens until your Head of Unit approves. If they do, the file goes to the recycle bin, where it can be restored for 93 days."
                  : "Nothing happens until your Head of Unit approves. If they do, the people below get access to this file — and nothing else."}
              </p>

              <label style={s.label}>Reason — your approver sees only this</label>
              <textarea
                style={{ ...s.field, minHeight: 64, resize: "vertical" }}
                value={reason}
                onChange={(e) => setReason(e.target.value)}
              />

              {asking === "Share" && (
                <>
                  <label style={s.label}>Share with — email addresses, one per line or comma separated</label>
                  <textarea
                    style={{ ...s.field, minHeight: 56, resize: "vertical" }}
                    value={shareWith}
                    onChange={(e) => setShareWith(e.target.value)}
                  />
                  {/* Named as it is typed, not only after submitting: whether someone is outside the
                      organisation is the fact that decides whether this is refused, and finding out
                      at the end means retyping the lot. */}
                  {parseRecipients(shareWith).some((r) => isExternal(r, policy.tenantDomains)) && (
                    <div style={{ ...s.warnBox, marginTop: 8 }}>
                      {policy.allowExternal
                        ? "Some of these are outside the organisation. Your approver will be told."
                        : "Some of these are outside the organisation, and that is switched off on this site — the request cannot be sent until you remove them."}
                    </div>
                  )}

                  <label style={s.label}>They may</label>
                  <select
                    style={s.field}
                    value={permission}
                    onChange={(e) => setPermission(e.target.value === "Edit" ? "Edit" : "View")}
                  >
                    <option value="View">View only</option>
                    <option value="Edit">View and edit</option>
                  </select>

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
                  disabled={sending}
                  onClick={() => { submitRequest(open, asking).catch(() => undefined); }}
                >
                  {sending ? "Sending…" : "Send the request"}
                </button>
                <button style={s.askBtn} disabled={sending} onClick={() => setAsking(undefined)}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        )}
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
        library, where the rest of your unit can find them too.{" "}
        {/* Discoverability: the buttons live in the detail view, because a deletion request should be
            made looking at the file rather than at a row in a list. That is only obvious once you
            know it, so the page says it. */}
        Open an approved file to ask for it to be <strong>deleted</strong> or{" "}
        <strong>shared</strong> — your Head of Unit decides.
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
