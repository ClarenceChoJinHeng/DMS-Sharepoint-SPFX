/**
 * CRS Requests — the approver's queue.
 *
 * Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
 *
 * A PIC cannot delete or share in `Documents`; they raise a request from My Submissions and a Head of
 * Unit decides it here.
 *
 * THE APPROVAL EXECUTES IN THE APPROVER'S OWN SESSION. When they press Approve, THEIR browser recycles
 * the file or grants the access, because they hold the rights the requester lacks. No service account,
 * no Power Automate, nothing acting on anyone's behalf — so the audit row names the person who
 * actually did it, and an approval can never exceed the approver's own rights. It fails loudly instead.
 *
 * The rules live in `shared/requests.ts`, under test. This file is the screen and the requests.
 */
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { IRequestsProps } from "./IRequestsProps";
import {
  cachedListTitle,
  LIST_SUFFIX,
  libraryTitle,
  noteCreatedList,
  titleForNewList,
} from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { useLiveRefresh } from "../../../shared/liveRefresh";
import { writeAudit } from "../../../shared/spAuditLog";
import { EVENT } from "../../../shared/auditLog";
import { normalizeRoleValue } from "../../../shared/groupMapModel";
import { isSystemAdmin } from "../../../shared/spGroups";
import {
  RequestRow,
  RequestStatus,
  parseRequestStatus,
  REQUEST_STATUSES,
  applyDecision,
  canDecide,
  counts,
  decisionSummary,
  isExternal,
  isVisibleTo,
  queueFor,
  ViewerScope,
  stageOf,
  RequestType,
  collectSharedFiles,
  mergeShareAcl,
  currentlyShared,
  withoutGoneFiles,
  canRevoke,
  endsTheShare,
  revocationNote,
  localDateStamp,
  longDate,
  SharedFile,
  ShareState,
  FileAcl,
  AclPrincipal,
} from "../../../shared/requests";

/**
 * ⚠⚠ EVERY READ ON THIS PAGE IS NO-CACHE. This is not belt-and-braces — it is the fix for the
 * single most confusing defect of 2026-08-30.
 *
 * The browser was serving the REQUEST LIST from cache. A plain refresh replayed a response captured
 * before those rows were decided, so two requests that the list held as `Revoked` came back on
 * screen as `Pending`, with live Approve and Reject buttons. Pressing either was refused by the
 * pre-write re-read guard — correctly, since it asks the SERVER — and the client reported it as
 * *"I click approve and it doesn't work"*. Clearing the cache made the same rows vanish into
 * Decided, which is what finally identified it. Client: *"this is so messy."*
 *
 * It sent three separate diagnoses down the wrong road before that: a stale BUNDLE (the tab's
 * JavaScript), a status-parsing bug, and a miscounting summary line. **A stale RESPONSE and a stale
 * BUNDLE present almost identically, and the difference is that clearing the cache fixes the first
 * while only a hard reload fixes the second.**
 *
 * A permissions-and-approvals screen must never render a cached answer: everything here is a
 * decision somebody is about to act on. The cost is a handful of uncached requests per visit.
 */
const GET = {
  Accept: "application/json;odata=nometadata",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
};

/**
 * Kept as a NAME rather than a second header set — `GET` carries the same no-cache headers now, and
 * two objects that must stay identical are how one of them quietly stops being.
 */
const GET_FRESH = GET;

/** A per-call cache-buster: the only part of the guard a proxy or service worker cannot ignore. */
const bust = (): string =>
  `&_=${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;

const s: Record<string, React.CSSProperties> = {
  wrap: {
    fontFamily: '"Segoe UI", system-ui, sans-serif',
    color: "#242424",
    fontSize: 13,
    lineHeight: 1.5,
    /* From the client's own devtools edit on this web part's `<section>` (2026-09-04). They said
       `20px 60px` first and corrected it to `20px 80px` in the same message.

       ⚠ THIS IS OUR SECTION, NOT SHAREPOINT'S CANVAS ZONE — the distinction matters, because the
       first screenshot pointed at `.r_Ab1x8_y298L`, the ControlZone, which ALREADY carries its own
       `padding: 20px 60px` and is not ours to set. The second screenshot resolved it: the inspected
       element is `<section>` with the padding in `element.style`, which is this object. So the two
       paddings NEST, and the content sits 80px inside whatever the zone already gives it. */
    /* WARN: THIS REPLACES THE `20px 80px` SET EARLIER THE SAME DAY. The client asked for that
       first, then for *"the same padding spacing that the upload form is using"* - and the two are
       not the same thing. The Upload Form does not get its inset from padding: its padding is only
       24px, and what actually holds it off the edge is `max-width` plus `margin: auto`. Copying the
       24px ALONE would have put this page nearer the wall than before, which is the opposite of
       what was asked. The whole shell is what produces the effect, so the whole shell is copied.
       1100 matches My Submissions - its closest sibling, and the other card list in the product. */
    maxWidth: 1100,
    margin: "32px auto", padding: "0 24px 48px",
  },
  h2: { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 6px" },
  sub: { fontSize: 13, color: "#5f5f5f", margin: "0 0 20px", lineHeight: 1.55 },
  card: {
    border: "1px solid #e1e1e1",
    borderRadius: 8,
    background: "#fff",
    padding: 16,
    marginBottom: 16,
  },
  head: { fontSize: 15, fontWeight: 600, margin: "0 0 10px" },
  /* ⚠ WHITE WITH A SHADOW, AND SPACED (client, 2026-09-04: *"client is complaining it is blending in
     in the background making it difficult for them to see"*). It was `#fafafa` on a white card — all
     but invisible. White-on-white works only WITH the shadow; if the shadow is ever removed the
     background has to come back. `borderLeft` is set per row from `STATUS_ACCENT`, so it is
     deliberately absent here. */
  row: {
    border: "1px solid #ececec",
    borderRadius: 8,
    padding: "14px 16px",
    /* 40, from the client's own devtools edit (2026-09-04: *"Add margin bottom 40px for each box"*).
       Cards carry a status stripe and a shadow, so they read as separate objects rather than rows of
       a list — 14 packed them tightly enough that two adjacent ones looked like one. */
    marginBottom: 40,
    background: "#fff",
    boxShadow: "0 1px 4px rgba(0,0,0,.10)",
  },
  rowTop: {
    display: "flex",
    gap: 10,
    alignItems: "baseline",
    flexWrap: "wrap",
  },
  name: { fontWeight: 600, flex: "1 1 240px", wordBreak: "break-word" },
  pill: { fontSize: 11, fontWeight: 600, borderRadius: 10, padding: "2px 8px" },
  /* ⚠ THIS KEY WAS REFERENCED AND NEVER DEFINED, and it is the whole reason the accordion headers
     looked wrong (client, 2026-09-04: *"can you ensure the button is like this"*, over a screenshot
     of a bordered grey `<button>` box). `s.accHead` yielded `undefined`, so the element rendered with
     the BROWSER'S DEFAULT button chrome — which produced every complaint in that message at once: the
     box (UA border + background), the header not filling the card (a bare button shrink-wraps its
     text, so only that small area was clickable), and no room for the chevron to sit anywhere.
     ⚠ EXACTLY THE TRAP THE COMMENT ABOVE WARNS ABOUT — a missing key is invisible to `tsc` and to
     lint, and the page still renders, so nothing fails. Add every new key here.

     `appearance: none` is what actually removes the chrome in Chromium; `background`/`border` alone
     leave the inner button box in some engines. `width: 100%` is what makes the ENTIRE header
     clickable, which is what was asked for — a wrapping `div` with an `onClick` would also have done
     it and would have cost keyboard focus and `aria-expanded`, so the button stays and simply fills. */
  accHead: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    width: "100%",
    appearance: "none",
    background: "none",
    border: "none",
    borderRadius: 6,
    /* The right-hand padding is what keeps the chevron OFF the edge — client: *"put the arrow
       dropdown on the right side at the edge but dont make it stick."* */
    padding: "10px 12px",
    margin: 0,
    fontFamily: "inherit",
    fontSize: 15,
    fontWeight: 600,
    color: "#242424",
    textAlign: "left",
    cursor: "pointer",
  },
  meta: { fontSize: 11.5, color: "#6b7a71", marginTop: 4 },
  /* ⚠ `s` IS A `Record<string, CSSProperties>`: a key that does not exist yields `undefined`
     and the element renders UNSTYLED with a green build. Add every new key here. */
  reasonLabel: {
    marginTop: 10,
    fontSize: 12,
    fontWeight: 600,
    color: "#3b3a39",
  },
  reason: {
    marginTop: 8,
    fontSize: 12.5,
    background: "#fff",
    border: "1px solid #eceaea",
    borderRadius: 6,
    padding: "8px 10px",
  },
  actions: { display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" },
  approve: {
    padding: "6px 16px",
    fontSize: 12.5,
    background: "#0f6c3f",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    cursor: "pointer",
  },
  reject: {
    padding: "6px 16px",
    fontSize: 12.5,
    color: "#a4262c",
    border: "1px solid #a4262c",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  off: {
    padding: "6px 16px",
    fontSize: 12.5,
    background: "#e6e6e6",
    color: "#9a9a9a",
    border: "none",
    borderRadius: 4,
    cursor: "not-allowed",
  },
  ghost: {
    padding: "6px 14px",
    fontSize: 12.5,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  /* ── The decision dialog, to the client's mockup (2026-09-04) ──────────────────────────────────
     ⚠ EVERY KEY BELOW MUST EXIST HERE. `s` is a `Record<string, CSSProperties>`, so a missing one
     yields `undefined` and the element renders with no styling and a green build — which is exactly
     what happened to the accordion headers earlier today. */
  /** The filename, on its own line under the title. Quiet, because it identifies rather than states. */
  modalFile: { fontSize: 12, color: "#8a8886", margin: "0 0 10px", wordBreak: "break-word" },
  /** A textarea rather than the single-line input the dialog used: a reason runs to a sentence. */
  noteArea: {
    width: "100%",
    boxSizing: "border-box",
    padding: "8px 10px",
    fontSize: 13,
    fontFamily: "inherit",
    minHeight: 72,
    resize: "vertical",
    border: "1px solid #c7c7c7",
    borderRadius: 4,
  },
  /* ⚠ BLACK, AND ONLY IN THIS DIALOG — the client's mockup shows a black Approve here while the
     Approve on every request CARD stays green (`s.approve`). Deliberately not unified: the mockup is
     of this dialog, and changing the cards too would be a redesign nobody asked for. Raise it if the
     two ever need to match. */
  approveDark: {
    padding: "6px 18px",
    fontSize: 12.5,
    background: "#1b1b1b",
    color: "#fff",
    border: "1px solid #1b1b1b",
    borderRadius: 4,
    cursor: "pointer",
  },
  warn: {
    border: "1px solid #f2c9a0",
    background: "#fff8f0",
    borderRadius: 8,
    padding: "12px 14px",
    fontSize: 12.5,
    color: "#8a4b00",
    marginBottom: 16,
  },
  err: {
    border: "1px solid #f1b0b3",
    background: "#fdf3f4",
    borderRadius: 8,
    padding: "12px 14px",
    fontSize: 12.5,
    color: "#a4262c",
    marginBottom: 16,
  },
  ok: {
    border: "1px solid #b7dcc4",
    background: "#f3faf5",
    borderRadius: 8,
    padding: "12px 14px",
    fontSize: 12.5,
    color: "#1c4d33",
    marginBottom: 16,
  },
  quiet: { fontSize: 12.5, color: "#767676" },
  /* ⚠ `s` is a Record<string, CSSProperties>, so a key that does not exist yields `undefined` and
     the element renders UNSTYLED with a green build. Every style referenced below must be here —
     the trap that has now bitten MySubmissions, GroupManager and SubtreeMigrator. */
  // Section header row (emoji + label + pending-count pill), replacing the 2026-08-28 tab bar —
  // client's mockup, 2026-09-03: Delete and Share are one page now, not two tabs to click between.
  sectionHead: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    fontSize: 15,
    fontWeight: 600,
    margin: "0 0 10px",
  },
  topBar: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    margin: "0 0 16px",
  },
  fileRow: { padding: "12px 0", borderTop: "1px solid #f0f0f0" },
  recip: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    flexWrap: "wrap",
    padding: "6px 0 6px 14px",
    borderLeft: "2px solid #f0f0f0",
    marginLeft: 2,
  },
  chip: {
    fontSize: 11,
    fontWeight: 600,
    borderRadius: 10,
    padding: "2px 8px",
    whiteSpace: "nowrap",
  },
  small: { fontSize: 12, color: "#5f5f5f" },
  revoke: {
    padding: "4px 10px",
    fontSize: 12,
    border: "1px solid #a4262c",
    color: "#a4262c",
    background: "#fff",
    borderRadius: 4,
    cursor: "pointer",
    fontFamily: "inherit",
  },
  revokeAll: {
    padding: "5px 12px",
    fontSize: 12,
    border: "1px solid #a4262c",
    background: "#a4262c",
    color: "#fff",
    borderRadius: 4,
    cursor: "pointer",
    fontFamily: "inherit",
  },
  input: {
    width: "100%",
    maxWidth: 460,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
  },
  modalBg: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,.4)",
    zIndex: 100,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 16,
  },
  modal: {
    background: "#fff",
    borderRadius: 8,
    padding: 20,
    width: "min(560px, 94vw)",
    maxHeight: "86vh",
    overflowY: "auto",
  },
};

/**
 * The LEFT BORDER of a request card, one colour per status.
 *
 * Client, 2026-09-04: *"if its pending highlight the same color we have been using for the pending
 * tag and put it on the left border of the box. If its rejected show the same red we have been
 * using, if its Approve use the same color we have been using."*
 *
 * Each is the STRONG half of that status's existing pill — `Pending`'s pale amber background would
 * be invisible as a 4px rule, so its text colour is used instead, and Approved/Rejected are already
 * solid so their own background is the accent. Derived from the same three values `PILL` uses, so a
 * card's border and its badge can never disagree about what a status looks like.
 *
 * A `Record` over the union, so adding a `RequestStatus` is a compile error here until it is given a
 * colour — the same guard that has kept `PILL` complete through three additions.
 */
const STATUS_ACCENT: Record<RequestStatus, string> = {
  /* ⚠ THE TAG'S OWN AMBER, NOT ITS TEXT COLOUR (client, 2026-09-04: *"For the Pending one can you
     change the border color to follow the Pending color tag?"*). It was `#8a4b00` — the brown the
     word "Pending" is PRINTED in — which made Pending the odd one out: `Approved` and `Rejected`
     already take their tag's BACKGROUND (`#0f6c3f`, `#a4262c`), so only this row disagreed with the
     pill beside it.
     ⚠ COST, AND IT IS VISIBLE: `#fff2df` is a pale amber, so the stripe is far subtler than the other
     statuses' solid green and red. That is inherent in matching a pale tag; say so if the stripe
     should stay dark, and the fix is this one value. */
  Pending: "#fff2df",
  Approved: "#0f6c3f",
  Rejected: "#a4262c",
  Failed: "#a4262c",
  Cancelled: "#767676",
  Revoked: "#767676",
};

const PILL: Record<RequestStatus, React.CSSProperties> = {
  Pending: { background: "#fff2df", color: "#8a4b00" },
  Approved: { background: "#0f6c3f", color: "#fff" },
  /* SOLID RED, mirroring Approved's solid green — they are the two DECISIONS and must read as a
     matched pair. Grey put a refusal in the same visual register as Cancelled, i.e. as something
     that merely stopped rather than something an approver decided against. Distinct from `Failed`,
     which stays PALE red: "the approver said no" and "the action broke" need telling apart. */
  Rejected: { background: "#a4262c", color: "#fff" },
  Failed: { background: "#fdf3f4", color: "#a4262c" },
  // Grey and quiet — a withdrawn request is not a failure and must not read as one.
  Cancelled: { background: "#f3f2f1", color: "#767676" },
  // Same reasoning: a revocation is the system working. Amber rather than red — it is a deliberate
  // act somebody took, and it has to be distinguishable at a glance from a share that is still live.
  Revoked: { background: "#fff2df", color: "#8a4b00" },
};

/* Fixed order for every status filter. Options are DERIVED from the rows actually present, so this
   list being complete costs nothing: `Pending` is simply never offered on the decided list, because
   those rows sit in the queue above it — while `Your requests` does hold pending rows and needs it. */
/* Display order for the outcome filter — the SET comes from `REQUEST_STATUSES`, so a status can
   never be offered here and unparsed by `fromListItem`, which is the shape that produced the
   permanent-Pending bug. Ordered by how often an approver looks for it, not alphabetically. */
const STATUS_PREFERRED: RequestStatus[] = [
  "Pending",
  "Approved",
  "Rejected",
  "Revoked",
  "Cancelled",
  "Failed",
];
/* ⚠ DERIVED, so a status added later cannot go MISSING from the filter — it lands at the end
   rather than nowhere. The preferred list decides ORDER only; `REQUEST_STATUSES` decides membership,
   and it is the compiler-checked one. */
const STATUS_ORDER: RequestStatus[] = STATUS_PREFERRED.concat(
  REQUEST_STATUSES.filter((v) => STATUS_PREFERRED.indexOf(v) === -1),
);

/* Shared for every filter bar on this page. Kept as ONE definition rather than repeated inline: three
   bars that drift apart look like three different controls doing three different things. */
const FIL_WRAP: React.CSSProperties = {
  display: "flex",
  gap: 8,
  flexWrap: "wrap",
  alignItems: "center",
  margin: "0 0 10px",
};
const FIL_SELECT: React.CSSProperties = {
  padding: "6px 8px",
  border: "1px solid #c8c6c4",
  borderRadius: 2,
  fontSize: 13,
  minWidth: 130,
};
const FIL_INPUT: React.CSSProperties = {
  padding: "6px 8px",
  border: "1px solid #c8c6c4",
  borderRadius: 2,
  fontSize: 13,
  flex: "1 1 220px",
  minWidth: 180,
};
const FIL_CLEAR: React.CSSProperties = {
  padding: "6px 10px",
  border: "1px solid #c8c6c4",
  borderRadius: 2,
  background: "#fff",
  fontSize: 13,
  cursor: "pointer",
};
/* ⚠ 60vh, and safe ONLY because nothing on this page is absolutely positioned. A scroll container
   clips popovers — it has already broken Group Management's people picker, the upload form's info
   panels and the member-add dropdown. Re-check before adding one here. */
const SCROLLER: React.CSSProperties = {
  maxHeight: "60vh",
  overflowY: "auto",
  overflowX: "hidden",
  paddingRight: 4,
};

/**
 * ⚠ `unknown` IS DELIBERATELY THE LOUD ONE, and that inverts the usual instinct.
 *
 * Green "still has access" and grey "access ended" are both settled answers. `unknown` means the
 * permission read did not come back — so the honest rendering is a warning, not a neutral shrug. An
 * approver who reads it as "ended" stops looking, and the document stays reachable.
 */
const STATE_CHIP: Record<ShareState, React.CSSProperties> = {
  live: { background: "#e8f3ee", color: "#0f6c3f" },
  revoked: { background: "#f3f2f1", color: "#767676" },
  unknown: { background: "#fff2df", color: "#8a4b00" },
};

const STATE_LABEL: Record<ShareState, string> = {
  live: "has access",
  revoked: "access ended",
  // Never "no access" — see STATE_CHIP.
  unknown: "could not be checked",
};

/* ⚠ THE 2026-08-28 TAB BAR (Deletion / Share / Shared files) IS GONE, REPLACED BY THREE ALWAYS-
   VISIBLE SECTIONS ON ONE PAGE (client's mockup, 2026-09-03: "Separating Delete File & Share tabs
   are confusing"). Delete Requests and Share Requests render simultaneously, sharing ONE filter bar;
   Shared files becomes a third section, "Documents with Shared Access", further down the same page.
   `RequestType` itself ("Deletion"/"Share") is unchanged — only the tab mechanism that switched
   between them is gone; `ofType("Deletion")` etc. still work exactly as before. */
/* ⚠ "Your requests" WAS REMOVED 2026-08-30 (client: *"As an approver or HOD why do they need to
   have Your Request"*), and the reasoning behind it is the same one that removes the request BUTTONS
   from a Head of Unit's own files: a HoU and a HoD can delete and share directly, so they never need
   to ask anybody. Only a PIC raises requests, and a PIC cannot reach this page at all.

   An approver who somehow holds a request row can still see it on **My Submissions**, which is the
   uploader's own view and where a request they raised belongs. */

/** The columns the list must have. Internal names are space-free, as with the audit log. */
/* `type` is FieldTypeKind: 2 = Text, 3 = Note, 4 = DateTime. No other property is sent —
   see the NumberOfLines note in ensureColumns. */
const COLUMNS: Array<{ name: string; type: number }> = [
  { name: "RequestType", type: 2 },
  // TEXT, never Choice — as with RequestType and Status. A value absent from a Choice column's
  // `Choices` fails the WHOLE write, so the day a third stage is added every row of that stage
  // would be lost silently.
  { name: "Stage", type: 2 },
  { name: "Status", type: 2 },
  { name: "ItemUniqueId", type: 2 },
  { name: "ItemName", type: 2 },
  { name: "ItemUrl", type: 2 },
  { name: "Segment", type: 2 },
  { name: "Unit", type: 2 },
  { name: "UnitTermGuid", type: 2 },
  { name: "RequestedBy", type: 2 },
  { name: "RequestedAt", type: 4 },
  { name: "Reason", type: 3 },
  { name: "ShareWith", type: 2 },
  { name: "SharePermission", type: 2 },
  { name: "ExpiresAt", type: 4 },
  { name: "DecidedBy", type: 2 },
  { name: "DecidedAt", type: 4 },
  { name: "DecisionNote", type: 3 },
  { name: "RevokedBy", type: 2 },
];

type Load<T> =
  | { state: "loading" }
  | { state: "error"; message: string }
  | { state: "ready"; value: T };

export default function Requests({
  context,
}: IRequestsProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const me = (context.pageContext.user.email ?? "").toLowerCase();

  const [rows, setRows] = useState<Load<RequestRow[]>>({ state: "loading" });
  const [approverUnits, setApproverUnits] = useState<string[]>([]);
  /* Units under a department this person HEADS — approved-stage requests only. Separate state, not
     folded into `approverUnits`, because the two carry different powers: see `ViewerScope`. */
  const [hodUnits, setHodUnits] = useState<string[]>([]);
  /* ⚠ FAILS CLOSED AND SAYS SO. If the department could not be expanded, a Head of Department sees
     nothing extra — an over-wide fallback would show one department's requests to another's head.
     But "no requests" and "we could not work out your department" are different sentences, and a
     HoD told the first stops looking. Empty ≠ unknown, as everywhere else here. */
  const [hodScopeFailed, setHodScopeFailed] = useState(false);
  /**
   * A member of `CRS Owners`, or a site collection admin — decides EVERY request.
   *
   * ⚠ THIS SCREEN HAD NO ADMIN CHECK AT ALL, and that is why it was reported. Scope is built purely
   * from Group Map rows, and an admin holds NONE by design — `CRS Owners` grants through Full Control
   * rather than through a persona — so both unit arrays came back empty and an owner saw only the
   * requests they had raised themselves, under a banner telling them to go and add themselves a
   * mapping. Client, 2026-08-27: *"I notice the system admin don't show the full CRS Request."*
   *
   * ⚠ FIFTH SCREEN IN THIS GAP. Four checked `IsSiteAdmin` alone until `isSystemAdmin` was extracted
   * earlier the same day so Owners membership counted too; this one checked NEITHER. When a rule about
   * "who is an administrator" changes, grep every screen — the ones that never asked are invisible to
   * a search for the old check.
   *
   * `isSystemAdmin` FAILS CLOSED (false on any read failure), which is right here: a transient
   * failure costs an admin the queue for a moment, where a wrong `true` would hand every request on
   * the site to whoever asked.
   */
  const [systemAdmin, setSystemAdmin] = useState(false);
  const [tenantDomains, setTenantDomains] = useState<string[]>([]);
  const [listMissing, setListMissing] = useState(false);
  const [stageMissing, setStageMissing] = useState(false);
  /* Gates the WRITE below, not a banner. Attempting a column the list lacks fails the whole MERGE
     (gotcha #11), which would lose the revoke's DecisionNote line as well — and that line is where
     the revoker's name currently lives. Degrade to prose; never lose the record. */
  const [revokedByMissing, setRevokedByMissing] = useState(false);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState<string | undefined>(undefined);
  /* ⚠ IS THE NOTICE A FAILURE? Found live 2026-08-30: EVERY message rendered in the GREEN success
     style — "Recorded as failed", "Could not confirm this request is still open", "Could not
     finish" included. An approver whose click was refused was shown a green box saying so, at the
     TOP of a page they were scrolled well down, and reported it as the button doing nothing. */
  const [noticeBad, setNoticeBad] = useState(false);
  const [deciding, setDeciding] = useState<
    { row: RequestRow; approve: boolean } | undefined
  >(undefined);
  const [note, setNote] = useState("");
  /* ⚠ A NOTE IS NOW REQUIRED FOR EVERY DECISION, APPROVE OR REJECT (client, 2026-09-03, in their own
     capitals: "IF APPROVER WANTS TO REJECT A REQUEST ... APPROVER MUST INCLUDE THE REASON ... ELSE
     SYSTEM DOESN'T ALLOW TO PROCEED"). This REVERSES the 2026-08-30 design, whose own comment argued
     an approval "needs none" — that argument is superseded, not wrong for its time.
     Gated the same way every other required-field check in this codebase is: NEVER on load (a blank
     dialog must not look broken the instant it opens), only after a decision was attempted with
     nothing typed. Reset to false every time the dialog (re)opens. */
  const [showNoteError, setShowNoteError] = useState(false);
  /* ⚠ ONE SHARED PAIR NOW, ACROSS BOTH Delete Requests AND Share Requests (client, 2026-09-03: "yes
     share a single filter bar") — REVERSES the note below about a separate pair per tab, because
     there is no longer a tab to be separate FROM. Status still never hides a Pending row in EITHER
     section; see `applyFilters`. */
  const [statusFilter, setStatusFilter] = useState<string>("All");
  const [whoFilter, setWhoFilter] = useState<string>("");
  /* `Shared files` keeps its OWN pair, never folded into the one above. It filters on an access STATE
     (`live`/`unknown`), not a request status, so a shared control would carry a value that section
     cannot mean — and a filter silently matching nothing is how a list reads as empty. */
  const [sharedState, setSharedState] = useState<string>("All");
  const [sharedText, setSharedText] = useState<string>("");
  /**
   * The live verdict, keyed on the file's UniqueId.
   *
   * ⚠ A KEY THAT IS ABSENT MEANS "NOT PROBED" AND MUST NOT READ AS "NOT SHARED". `mergeShareAcl`
   * turns a missing entry into `unknown` for exactly that reason, and every render goes through it
   * rather than reading this map directly.
   */
  const [acls, setAcls] = useState<{ [uniqueId: string]: FileAcl | undefined }>(
    {},
  );
  const [aclsLoading, setAclsLoading] = useState(false);
  const [aclsRead, setAclsRead] = useState(false);
  const [revoking, setRevoking] = useState<string | undefined>(undefined);

  const listUrl = (): string =>
    `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.requests))}')`;

  /* JSON light with NO `__metadata`, and BOTH header halves saying nometadata. Verbose on one half
     gives "The property '__metadata' does not exist on type 'SP.List'"; verbose on both gives
     "Parsing JSON Light feeds or entries in requests without entity set is not supported", because
     SPFx attaches its own OData version header. `odata-version: ""` is needed for the same reason:
     SPFx injects 4.0, under which SharePoint cannot infer the entity set for a JSON-light entry.
     All learned on the audit log — do NOT "restore" ListItemEntityTypeFullName here. */
  const writeHeaders = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
    "odata-version": "",
  };

  const post = async (
    url: string,
    body?: unknown,
  ): Promise<SPHttpClientResponse> =>
    context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: writeHeaders,
      body: body === undefined ? undefined : JSON.stringify(body),
    });

  /* ── Load ─────────────────────────────────────────────────────────────── */

  const load = async (): Promise<void> => {
    await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);

    // Which units does this person approve for? From their group memberships matched against the
    // Group Map, keyed on the TERM GUID — a renamed unit must not silently empty someone's queue.
    const units: string[] = [];
    /* `${termSetGuid}|${departmentTermGuid}` pairs, from this viewer's DEPTVIEW rows. Kept as one
       string so the two halves cannot be separated — expanding a department needs both, and a
       department guid looked up in the wrong set answers 404. */
    const depts: string[] = [];
    try {
      // The CURRENT USER's groups, not every group on the site — the same endpoint the upload form
      // uses. `fetchAllSiteGroups` would list all of them and make every unit look like this
      // person's to approve.
      const mineRes = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser/groups?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      const myIds: string[] = mineRes.ok
        ? ((await mineRes.json()).value ?? []).map((g: { Id?: number }) =>
            String(g.Id ?? ""),
          )
        : [];
      const mapRes = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items` +
          `?$select=GroupId,Role,UnitTermGuid,Segment&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (mapRes.ok) {
        const data = await mapRes.json();
        for (const r of (data.value ?? []) as Array<{
          GroupId?: string;
          Role?: string;
          UnitTermGuid?: string;
          Segment?: string;
        }>) {
          // APR is the plain approver role; APRHC is the HC Head of Unit's, held INSTEAD of APR
          // since 2026-08-24. Both say "this person decides for this unit" — see the match below.
          //
          // normalizeRoleValue, NOT a raw toUpperCase — which this once was, and was a latent bug of
          // the same class formModel.ts documents. The Group Map's Role is often the LONG form
          // ("APPROVER", as "UPLOADER" was found live on 2026-08-07), and a raw compare against "APR"
          // skips such a row. The consequence is silent and total: that unit's approver queue stays
          // empty for ever while requests pile up behind it, with nothing on screen to say so.
          const role = normalizeRoleValue(r.Role ?? "");
          if (myIds.indexOf(String(r.GroupId ?? "")) === -1) continue;
          const guid = (r.UnitTermGuid ?? "").trim();
          if (!guid) continue;
          /* APR OR APRHC since 2026-08-24 — APRHC is a real role again (the HC Head of Unit holds it
             INSTEAD of APR), and the old normalizeRoleValue alias that rewrote APRHC to APR is gone
             with it. Matching APR alone would leave every HC unit's queue permanently empty while
             requests piled up behind it — the exact failure CLAUDE.md has always warned about. */
          if (role === "APR" || role === "APRHC") {
            if (units.indexOf(guid) === -1) units.push(guid);
            continue;
          }
          /* DEPTVIEW is a HEAD OF DEPARTMENT (2026-08-21). The row's `UnitTermGuid` holds the
             DEPARTMENT term, and `Segment` holds the segment's TERM SET guid — see
             `groupMapModel.ts` ~779. Both halves are needed to expand it, so keep them paired. */
          if (role === "DEPTVIEW") {
            const set = (r.Segment ?? "").trim();
            if (!set) continue;
            const key = `${set}|${guid}`;
            if (depts.indexOf(key) === -1) depts.push(key);
          }
        }
      }
    } catch {
      // An unreadable Group Map means an empty queue, not a broken page — said on screen below.
    }
    setApproverUnits(units);

    /* Expand each headed department to its child UNITS.
     *
     * ONE request per department, and a person heads one in the normal case. Requests are filed
     * against a unit term, so a department guid matches nothing on its own — the expansion is what
     * makes a HoD's queue possible at all.
     *
     * NOT a stored column on the request row: that is a migration, every row written before today
     * would lack it, and the `Stage` column two days ago showed exactly what that costs. Ancestry is
     * derivable, and a stored copy would describe a shape a re-parented term has since left.
     *
     * FAILS CLOSED per department and reports it — a partial expansion would silently shrink one
     * head's queue while another's looked complete. */
    const under: string[] = [];
    let expansionFailed = false;
    for (const pair of depts) {
      const cut = pair.indexOf("|");
      const setGuid = pair.slice(0, cut);
      const deptGuid = pair.slice(cut + 1);
      try {
        const res = await context.spHttpClient.get(
          `${siteUrl}/_api/v2.1/termStore/sets/${encodeURIComponent(setGuid)}` +
            `/terms/${encodeURIComponent(deptGuid)}/children?$select=id`,
          SPHttpClient.configurations.v1,
          { headers: GET },
        );
        if (!res.ok) {
          // Status named in the console, as gotcha #9 requires: a 404 is a term that has been
          // deleted or re-parented, a 403 is the term store not being readable by this person, and
          // the two need opposite fixes.
          console.warn(
            `[CRS Requests] department children ${deptGuid} returned HTTP ${res.status}`,
          );
          expansionFailed = true;
          continue;
        }
        const body = await res.json();
        for (const t of (body.value ?? []) as Array<{ id?: string }>) {
          const id = (t.id ?? "").trim();
          if (id && under.indexOf(id) === -1) under.push(id);
        }
      } catch {
        expansionFailed = true;
      }
    }
    setHodUnits(under);
    // Only a WARNING when they head a department at all. Someone with no DEPTVIEW row has nothing to
    // expand, and telling them their department could not be read would be a message about a role
    // they do not hold.
    setHodScopeFailed(depts.length > 0 && expansionFailed);

    /* THE TENANT'S OWN DOMAINS, so `isExternal` can tell inside from outside.
     *
     * ⚠ THE CONFIG ROW WINS OUTRIGHT WHEN IT EXISTS — the viewer's own domain is only a FALLBACK.
     * It used to be seeded unconditionally, on the reasoning that "they are signed in to this
     * tenant, so it is not a guess". **That is false for a guest**, and guests are how this project
     * is tested and how the agency accounts will exist on the client's tenant. Found live
     * 2026-08-30: signed in as a gmail guest, `clarencemindpalace@gmail.com` was shown as a
     * colleague and `clarence@trinergydigital.com` was flagged *outside the organisation* — both
     * are outside `dcidigitalcom`, so the label was wrong in BOTH directions on one screen. Client:
     * *"Shouldn't that be clarencemindpalace@gmail.com kinda misleading."*
     *
     * With no config row, seeding the viewer's domain is still the best available guess and keeps
     * an unconfigured site behaving as it did. With one, it is an ANSWER and must not be widened. */
    const configured: string[] = [];
    try {
      const cfg = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Title,SettingValue&$filter=Title eq 'tenantDomains'&$top=5`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (cfg.ok) {
        const data = await cfg.json();
        for (const r of (data.value ?? []) as Array<{
          SettingValue?: string;
        }>) {
          for (const d of (r.SettingValue ?? "").split(/[,;\s]+/)) {
            const clean = d.trim().toLowerCase();
            if (clean && configured.indexOf(clean) === -1)
              configured.push(clean);
          }
        }
      }
    } catch {
      /* falls back to the viewer's own domain below */
    }
    const domains: string[] = configured.slice();
    if (domains.length === 0) {
      const at = me.lastIndexOf("@");
      if (at > -1) domains.push(me.slice(at + 1));
    }
    setTenantDomains(domains);

    /**
     * `Stage` is asked for, and the read is RETRIED WITHOUT it on failure.
     *
     * The column is provisioned by this page, so on any site whose Requests list predates
     * 2026-08-20 it does not exist — and ONE unknown name in a `$select` fails the WHOLE request
     * with HTTP 400 (gotcha #11). Asking unconditionally would empty the approver's entire queue on
     * a site that is otherwise fine. Its absence is exactly the case `stageOf` already answers:
     * every row written before the column existed was an approved document.
     */
    const base =
      `${listUrl()}/items?$select=Id,RequestType,Status,ItemUniqueId,ItemName,ItemUrl,Segment,Unit,` +
      `UnitTermGuid,RequestedBy,RequestedAt,Reason,ShareWith,SharePermission,ExpiresAt,DecidedBy,` +
      `DecidedAt,DecisionNote`;
    /* ⚠ `bust()` ON THE LIST READ, not only the headers. This is THE read that was served from
       cache on 2026-08-30, showing decided requests as pending with live Approve buttons. Headers
       ask politely; a URL nothing has seen before cannot be answered from a store. */
    const tail = `&$top=2000&$orderby=Id desc${bust()}`;
    try {
      /* THREE RUNGS, newest column first. `Stage` (2026-08-20) and `RevokedBy` (2026-08-30) are both
         added by this page's own `addMissingColumns`, so a site can lack either or both — and ONE
         unknown name fails the WHOLE `$select` (gotcha #11), so a single retry is not enough: a site
         holding `Stage` but not `RevokedBy` would fall all the way through and lose the stage with
         it. Same ladder as `readLibrary` in My Submissions. */
      let noStage = false;
      let noRevokedBy = false;
      let res = await context.spHttpClient.get(
        `${base},Stage,RevokedBy${tail}`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      // 400 only. A 404 is the LIST missing and is answered below; retrying that would report an
      // unprovisioned list as an unreadable one.
      if (res.status === 400) {
        noRevokedBy = true;
        res = await context.spHttpClient.get(
          `${base},Stage${tail}`,
          SPHttpClient.configurations.v1,
          { headers: GET },
        );
      }
      if (res.status === 400) {
        /* Both flags now true, and `noRevokedBy` may be OVERSTATED — a site missing only `Stage`
           lands here with `RevokedBy` present. Deliberate: the flag's only job is to stop the write
           below attempting a column that may not exist, and a needless skip costs one prose line in
           the audit row, where a wrong guess fails the whole MERGE and loses the revoke's record. */
        noStage = true;
        res = await context.spHttpClient.get(
          `${base}${tail}`,
          SPHttpClient.configurations.v1,
          { headers: GET },
        );
      }
      if (res.status === 404) {
        setListMissing(true);
        setRows({ state: "ready", value: [] });
        return;
      }
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      setListMissing(false);
      // Reported, not repaired silently. Without the column a pending-file request is still raised
      // and still deletable, but it READS as an approved document in the queue — and the difference
      // is which library the file leaves. An approver must not be told the wrong one.
      setStageMissing(noStage);
      setRevokedByMissing(noRevokedBy);
      setRows({
        state: "ready",
        value: ((data.value ?? []) as Record<string, string>[]).map(
          fromListItem,
        ),
      });
    } catch (e) {
      // "Could not read" is never rendered as "there are none" — an approver told their queue is
      // empty stops looking, and a request sits unanswered.
      setRows({ state: "error", message: (e as Error).message });
    }
  };

  useEffect(() => {
    load().catch((e) =>
      setRows({ state: "error", message: (e as Error).message }),
    );
  }, [siteUrl]);

  /* Its OWN effect, not folded into `load`. That function reads the list and sets the queue, and an
     admin check failing must never be able to leave the page reporting it could not read requests —
     the two facts are independent and their failures mean different things. */
  useEffect(() => {
    let live = true;
    isSystemAdmin(context.spHttpClient, siteUrl)
      .then((yes) => {
        if (live) setSystemAdmin(yes);
      })
      .catch(() => {
        if (live) setSystemAdmin(false);
      });
    return () => {
      live = false;
    };
  }, [siteUrl]);

  /* ── Provisioning ─────────────────────────────────────────────────────── */

  /**
   * The columns the list already has, lower-cased. Empty when unreadable — which makes the caller
   * ATTEMPT every column, and an attempt on an existing column merely fails harmlessly. Reading
   * first turns "already exists" errors from the normal case into a real signal.
   */
  const readFieldNames = async (title: string): Promise<string[]> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/fields?$select=Title&$top=500`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!res.ok) return [];
      return ((await res.json()).value ?? []).map((f: { Title?: string }) =>
        (f.Title ?? "").toLowerCase(),
      );
    } catch {
      return [];
    }
  };

  /**
   * Add every column the list is missing. Used by provisioning AND by the repair button.
   *
   * ONE implementation, reusing `COLUMNS`, so there is never a second list of what the list should
   * hold. An "already exists" failure is indistinguishable from any other 400 here and is treated as
   * success-by-omission: the column is simply absent from `added`, and a genuinely missing one shows
   * up on the next run because this is repeatable.
   */
  const ensureColumns = async (
    title: string,
  ): Promise<{ added: string[]; failed: string[] }> => {
    const added: string[] = [];
    const failed: string[] = [];
    const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/fields`;
    const have = await readFieldNames(title);
    for (const c of COLUMNS) {
      if (have.indexOf(c.name.toLowerCase()) > -1) continue;
      /**
       * ⚠ NO `NumberOfLines`. It belongs to `SP.FieldMultiLineText`, not `SP.Field`, and with
       * `odata=nometadata` SharePoint infers the entity type from the ENDPOINT — so the extra
       * property is rejected outright with HTTP 400. `Reason` is the first Note column in the list
       * and was the exact point every provisioning run died (2026-08-20). FieldTypeKind 3 already
       * yields a multi-line column; the line count is display-only and defaults sensibly.
       */
      const r = await post(fieldsUrl, { Title: c.name, FieldTypeKind: c.type });
      if (r.ok) added.push(c.name);
      else failed.push(`${c.name} (HTTP ${r.status})`);
    }
    return { added, failed };
  };

  /* ── Live shares: probing, and taking them back ────────────────────────────
     Spec: docs/superpowers/specs/2026-08-28-requests-page-redesign-and-share-revoke-design.md */

  /**
   * Read one file's real permissions.
   *
   * ⚠ RETURNS `undefined` ON EVERY FAILURE, AND THAT IS NOT THE SAME AS AN EMPTY ACL. A 403, a 404,
   * a throttle and a malformed body all mean "we do not know", which `mergeShareAcl` renders as
   * `unknown`. Answering `{ hasUniqueRoleAssignments: false, principals: [] }` here would report the
   * share as ENDED on the strength of a read that failed — the one wrong answer this screen must
   * never give.
   *
   * `GetFileById` is WEB-scoped, so it reaches the HC libraries and the archive with no library name
   * to resolve and no chance of asking the wrong one.
   *
   * ⚠ IT MUST ASK FOR `RoleDefinitionBindings`, AND NOT DOING SO WAS A REAL DEFECT (2026-08-28).
   * Being listed in `RoleAssignments` is not the same as having access: SharePoint's automatic
   * **Limited Access** entry appears wherever a principal was granted on a child and survives some
   * revocations, so a revoked recipient went on reading `has access` on the one screen that is
   * supposed to answer that question from live data. `holdsRealAccess` is what discards them, and it
   * can only do its job if the bindings are actually read.
   */
  const readAcl = async (uniqueId: string): Promise<FileAcl | undefined> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFileById(guid'${encodeURIComponent(uniqueId)}')/ListItemAllFields` +
          `?$select=HasUniqueRoleAssignments,RoleAssignments/Member/Id,` +
          `RoleAssignments/Member/PrincipalType,RoleAssignments/Member/Email,` +
          `RoleAssignments/Member/LoginName,RoleAssignments/Member/Title,` +
          `RoleAssignments/RoleDefinitionBindings/Id,` +
          `RoleAssignments/RoleDefinitionBindings/Name,` +
          `RoleAssignments/RoleDefinitionBindings/RoleTypeKind` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` +
          bust(),
        SPHttpClient.configurations.v1,
        { headers: GET_FRESH },
      );
      /* ⚠ A 304 MUST NEVER REACH THE MERGE AS AN ANSWER. With the guards above it should not be
         possible — but if one is stripped in transit, "not modified" means we were given nothing,
         and nothing is `unknown`, never an empty principal list. An empty list would read as
         "everyone has been revoked". */
      if (res.status === 304) return undefined;
      /* ⚠ 404 IS A DIFFERENT ANSWER FROM A FAILED READ, and only this one removes the row.
         `GetFileById` answers 404 when the document has been deleted — and also when it is outside
         this viewer's reach, since SharePoint security-trims to 404 rather than 403. Both mean "not
         there for you", which is what the tab reports. Every OTHER status stays `undefined` and
         renders as "could not be checked". */
      if (res.status === 404) {
        return {
          hasUniqueRoleAssignments: false,
          principals: [],
          fileGone: true,
        };
      }
      if (!res.ok) return undefined;
      const d = await res.json();
      if (!d || typeof d.HasUniqueRoleAssignments !== "boolean")
        return undefined;
      const assignments = (d.RoleAssignments ?? []) as Array<{
        Member?: Record<string, unknown>;
        RoleDefinitionBindings?: Array<Record<string, unknown>>;
      }>;
      const principals: AclPrincipal[] = [];
      for (const a of assignments) {
        const m = a && a.Member;
        if (!m) continue;
        const id = Number(m.Id);
        if (!isFinite(id)) continue;
        /* ⚠ `undefined` when the key is ABSENT, never `[]`. `holdsRealAccess` reads an absent list
           as "we did not ask" and keeps the principal; an empty one would mean "holds nothing" and
           would hide them. Empty is not unknown, in the place where the cost is a disclosure. */
        const raw = a.RoleDefinitionBindings;
        const bindings = Array.isArray(raw)
          ? raw.map((b) => ({
              id: typeof b.Id === "number" ? b.Id : undefined,
              name: typeof b.Name === "string" ? b.Name : undefined,
              roleTypeKind:
                typeof b.RoleTypeKind === "number" ? b.RoleTypeKind : undefined,
            }))
          : undefined;
        principals.push({
          id,
          principalType:
            typeof m.PrincipalType === "number" ? m.PrincipalType : undefined,
          email: typeof m.Email === "string" ? m.Email : undefined,
          loginName: typeof m.LoginName === "string" ? m.LoginName : undefined,
          title: typeof m.Title === "string" ? m.Title : undefined,
          bindings,
        });
      }
      return {
        hasUniqueRoleAssignments: d.HasUniqueRoleAssignments,
        principals,
      };
    } catch {
      return undefined;
    }
  };

  /**
   * Probe every shared file this viewer can see.
   *
   * ⚠ FAILURE IS PER FILE. `Promise.allSettled` is unavailable here (gotcha #3 — the SPFx tsconfig
   * does not target ES2020), so `readAcl` swallows its own errors and one unreadable file leaves the
   * rest correct. Costs one request per shared file, on opening this tab only — shares are rare by
   * design, which is the whole reason per-file ACLs are workable here where they were not for
   * per-uploader isolation.
   */
  const loadAcls = async (files: SharedFile[]): Promise<void> => {
    if (files.length === 0) {
      setAclsRead(true);
      return;
    }
    setAclsLoading(true);
    try {
      const pairs = await Promise.all(
        files.map(async (f) => ({
          id: f.itemUniqueId,
          acl: await readAcl(f.itemUniqueId),
        })),
      );
      const next: { [id: string]: FileAcl | undefined } = {};
      for (const p of pairs) next[p.id] = p.acl;
      setAcls(next);
      setAclsRead(true);
    } finally {
      setAclsLoading(false);
    }
  };

  /**
   * Take a share back.
   *
   * TWO MECHANISMS, and only the second reclaims the permission SCOPE — which is the half that
   * matters, because every approved share creates one and that is the same 50,000-per-list ceiling
   * that killed per-uploader ACLs on 2026-08-06. Per-recipient revocation alone leaves the count
   * monotonic, i.e. exactly the problem §5.3 of the 2026-08-15 design was written to solve.
   *
   *   all=false → removeroleassignment(principalid) — drops one person, keeps the scope
   *   all=true  → resetroleinheritance             — drops everyone AND reclaims the scope
   *
   * ⚠ `resetroleinheritance` returns the file to its UNIT FOLDER's ACL, not to the library root.
   * Nothing is lost: the unit's own groups reach it exactly as they did before the share.
   *
   * Runs in THIS person's session with THEIR permissions — the founding principle of this whole
   * feature. `CRS Share` contains Manage Permissions, which is why it exists as a separate level
   * (FolderManager.tsx ~341) and why no new role was needed for any of this.
   */
  const revoke = async (
    file: SharedFile,
    all: boolean,
    one?: string,
  ): Promise<void> => {
    setRevoking(file.itemUniqueId);
    setNotice(undefined);
    setNoticeBad(false);
    try {
      const merged = mergeShareAcl(file, acls[file.itemUniqueId]);
      const base = `${siteUrl}/_api/web/GetFileById(guid'${encodeURIComponent(file.itemUniqueId)}')/ListItemAllFields`;
      const going = all
        ? merged.recipients
            .filter((r) => r.state !== "revoked")
            .map((r) => r.email)
        : [one ?? ""];

      let failure: string | undefined;
      if (all) {
        const res = await post(`${base}/resetroleinheritance`);
        if (!res.ok)
          failure = `Access could not be reset (HTTP ${res.status}).`;
      } else {
        const target = merged.recipients.filter((r) => r.email === one)[0];
        if (!target || typeof target.principalId !== "number") {
          /* No principal id means the live read never resolved them. Refusing is right: removing an
             assignment we could not identify means guessing, and a wrong guess takes access away
             from somebody nobody asked about. */
          failure =
            "That recipient could not be matched to a live permission, so nothing was changed.";
        } else {
          const res = await post(
            `${base}/roleassignments/removeroleassignment(principalid=${target.principalId})`,
          );
          if (!res.ok)
            failure = `Access could not be removed (HTTP ${res.status}).`;
        }
      }

      if (failure) {
        setNotice(failure);
        return;
      }

      /* ── Record it ────────────────────────────────────────────────────────
         ⚠ NOT via `writeAudit`. That is REFUSED for every non-Owner — `CRS Audit Log` restricts
         writes to Owners and the service account by design, which is exactly what made the
         `CRS — Audit request activity` flow necessary on 2026-08-26. A Head of Unit pressing Revoke
         is such a non-Owner, so a `writeAudit` call here would silently do nothing.

         The ROW is updated instead, and the existing created-or-modified flow picks it up. See §4.4
         of the spec: until that flow gains a `Revoked` branch this logs as `RequestRejected` — the
         actor and the row are right, only the label misleads. */
      const ended = endsTheShare(merged, going);
      /* ⚠ LOCAL date, never `toISOString()` — that is UTC, and a revoke before 08:00 local
         recorded the PREVIOUS day while the audit row beside it showed today. See
         `localDateStamp`. */
      const line = revocationNote(going, me, localDateStamp(new Date()));
      let rowFailed = false;
      for (const id of file.rowIds) {
        // APPEND to whatever the approver wrote when they let it through — that note is the record
        // of WHY access was granted, and overwriting it to record the revoke destroys half the trail.
        const existing = (rows.state === "ready" ? rows.value : []).filter(
          (r) => r.id === id,
        )[0];
        const prior = (existing?.decisionNote ?? "").trim();
        const body: Record<string, string> = {
          DecisionNote: prior.length > 0 ? `${prior}\n${line}` : line,
        };
        /* WHO REVOKED — written on EVERY revoke, not only the one that closes the request, so the
           value always names the person who acted last. `DecidedBy` is deliberately left alone: it
           records who let the share through, and a Head of Department can revoke a share a Head of
           Unit approved. Collapsing the two would credit the approver with someone else's
           revocation in `CRS Audit Log`, which is the one place that must not be guessed at. */
        if (!revokedByMissing) body.RevokedBy = me;
        // Only the LAST recipient going closes the request; a partial revoke leaves it Approved,
        // because the request's own outcome has not changed — only part of its grant.
        if (ended) body.Status = "Revoked";
        const upd = await context.spHttpClient.post(
          `${listUrl()}/items(${id})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              ...writeHeaders,
              "X-HTTP-Method": "MERGE",
              "IF-MATCH": "*",
            },
            body: JSON.stringify(body),
          },
        );
        if (!upd.ok) rowFailed = true;
      }

      /* Re-probe THIS file rather than trusting what we just did. This screen's whole premise is
         that the ACL is the verdict; believing our own write here would make the one screen that
         checks reality the one screen that assumes it. */
      const fresh = await readAcl(file.itemUniqueId);
      setAcls((prev) => ({ ...prev, [file.itemUniqueId]: fresh }));
      await load().catch(() => undefined);

      setNoticeBad(rowFailed);
      setNotice(
        rowFailed
          ? "Access was removed, but the request row could not be updated — the change is real even though the list still shows the old status."
          : all
            ? "Access removed for everyone, and the file's permission scope has been released."
            : `Access removed for ${one}.`,
      );
    } catch (e) {
      setNoticeBad(true);
      setNotice(
        `Access could not be removed (${e instanceof Error ? e.message : String(e)}).`,
      );
    } finally {
      setRevoking(undefined);
    }
  };

  /* ⚠ PROBES ONCE, ON LOAD — no longer gated on opening a "Shared files" tab, since that tab is gone
     (2026-09-03 redesign: the section is always on screen). `useLiveRefresh` re-reads this page on
     focus and every 45 seconds; re-probing every shared file on each of those turns a background
     convenience into a per-file request storm. `aclsRead` is deliberately STICKY, and the Re-check
     button is how an approver asks for a fresh answer. A revoke re-probes its own file regardless. */
  useEffect(() => {
    if (aclsRead || aclsLoading) return;
    const rowsNow = rows.state === "ready" ? rows.value : [];
    loadAcls(
      collectSharedFiles(rowsNow, {
        aprUnits: approverUnits,
        hodUnits,
        systemAdmin,
      }),
    ).catch(() => undefined);
  }, [aclsRead, aclsLoading, rows.state]);

  /**
   * Ensure the list, then ensure its columns. REPEATABLE by construction.
   *
   * ⚠ IT WAS NOT, AND THAT COST A HALF-BUILT LIST ON THE FIRST LIVE RUN (2026-08-20). The old
   * version created the list, then threw on the FIRST column that failed — leaving the list in place
   * with ten of seventeen columns. Pressing the button again tried to CREATE it, hit HTTP 500 because
   * it already existed, and there was no route forward at all: the repair button only appears once
   * the list can be read, and the list could not be read because it lacked the columns.
   *
   * So: create only when absent, attempt EVERY column, and report the failures together. A partial
   * result stated plainly beats an abort that hides how far it got.
   */
  const provision = async (): Promise<void> => {
    setBusy(true);
    setNotice(undefined);
    setNoticeBad(false);
    try {
      // titleForNewList, NOT cachedListTitle. A list that does not exist yet can never be in the
      // name cache, so cachedListTitle answers the LEGACY `DMS Requests` — which on a CRS-renamed
      // site creates the one list nobody can find. See resolvedPrefix.
      const title = titleForNewList(LIST_SUFFIX.requests);
      const base = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')`;

      const probe = await context.spHttpClient.get(
        `${base}?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (probe.status === 404) {
        const made = await post(`${siteUrl}/_api/web/lists`, {
          Title: title,
          BaseTemplate: 100,
          Description:
            "Deletion and share requests raised by uploaders and decided by the Head of Unit.",
        });
        if (!made.ok)
          throw new Error(`creating the list failed — HTTP ${made.status}`);
      } else if (!probe.ok) {
        // Neither present nor absent. Creating on top of that is how a duplicate list appears.
        throw new Error(
          `could not check whether "${title}" exists — HTTP ${probe.status}`,
        );
      }
      // Recorded before the columns, so a run that half-finishes still leaves every later read
      // pointing at the list that now exists.
      noteCreatedList(LIST_SUFFIX.requests, title);

      const { added, failed } = await ensureColumns(title);
      await load();

      if (failed.length > 0) {
        // NAMED, and it does NOT claim success. A list missing ShareWith accepts a share request and
        // silently drops who it was for.
        setNoticeBad(true);
        setNotice(
          `"${title}" exists, but these columns could not be added: ${failed.join(", ")}. ` +
            "Pressing the button again is safe and will retry them.",
        );
        return;
      }
      setNotice(
        (added.length > 0
          ? `Ready — added ${added.join(", ")}. `
          : `"${title}" was already complete. `) +
          "Its permissions are NOT set automatically — an administrator should restrict who may " +
          "edit it, the same as the audit log.",
      );
    } catch (e) {
      setNoticeBad(true);
      setNotice(`Could not finish: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  /** The repair button on an existing list — same operation, different entry point. */
  const addMissingColumns = async (): Promise<void> => {
    setBusy(true);
    setNotice(undefined);
    setNoticeBad(false);
    try {
      const title = cachedListTitle(LIST_SUFFIX.requests);
      const { added, failed } = await ensureColumns(title);
      setNoticeBad(failed.length > 0);
      setNotice(
        failed.length > 0
          ? `Could not add: ${failed.join(", ")}. Pressing again is safe and will retry them.`
          : added.length > 0
            ? `Added: ${added.join(", ")}. Everything else was already there.`
            : "Every column was already there.",
      );
      await load();
    } catch (e) {
      setNoticeBad(true);
      setNotice(`Could not finish: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  /* ── Carrying out a decision ──────────────────────────────────────────── */

  /**
   * Recycle, never delete outright — restorable for 93 days, which is what makes approving a deletion
   * reasonable at all. Resolved by UniqueId, so a rename or move since the request was raised does not
   * matter.
   */
  const performDeletion = async (
    row: RequestRow,
  ): Promise<string | undefined> => {
    const res = await post(
      `${siteUrl}/_api/web/GetFileById(guid'${row.itemUniqueId}')/recycle()`,
    );
    if (res.ok) return undefined;
    if (res.status === 404)
      return "That document no longer exists — it may already have been deleted.";
    if (res.status === 403)
      return "You do not have permission to delete that document.";
    return `The document could not be deleted (HTTP ${res.status}).`;
  };

  /**
   * One endpoint for internal and external alike.
   *
   * `SP.Web.ShareObject` is what SharePoint's own Share dialog calls, so it honours the tenant and
   * site sharing settings rather than working around them. If external sharing is off there, this
   * fails and says so — the correct outcome, and far better than a grant that half-works.
   */
  const performShare = async (row: RequestRow): Promise<string | undefined> => {
    const people = (row.shareWith ?? []).map((e) => ({ Key: e }));
    if (people.length === 0)
      return "No recipients were recorded on this request.";
    if (!row.itemUrl) return "This request has no document address recorded.";

    // SharePoint's own role values: 1073741826 = View, 1073741827 = Edit.
    const roleValue =
      row.sharePermission === "Edit" ? "role:1073741827" : "role:1073741826";
    const res = await post(`${siteUrl}/_api/SP.Web.ShareObject`, {
      url: `${window.location.origin}${row.itemUrl}`,
      peoplePickerInput: JSON.stringify(people),
      roleValue,
      groupId: 0,
      propagateAcl: false,
      sendEmail: true,
      includeAnonymousLinkInEmail: false,
      emailSubject: `A document has been shared with you: ${row.itemName}`,
      /* The invite itself is correctly attributed to the APPROVER by SharePoint (this call runs in
         their own session) — verified 2026-08-21. That surprised a requester reading the email, since
         they raised the request and never see themselves named anywhere in it. This line names the
         requester too, so both people involved are visible without changing who actually shared it. */
      emailBody: [
        row.reason,
        row.requestedBy ? `Requested by ${row.requestedBy}` : "",
      ]
        .filter((s) => (s ?? "").trim().length > 0)
        .join("\n\n"),
      useSimplifiedRoles: true,
    });
    if (!res.ok) {
      if (res.status === 403)
        return "You do not have permission to share that document.";
      return `The document could not be shared (HTTP ${res.status}).`;
    }
    // HTTP 200 does NOT mean it worked — the per-recipient result is in the body, exactly as with
    // validateUpdateListItem (gotcha #4). A refusal by tenant policy arrives here, not as a status.
    try {
      const body = await res.json();
      const results = (body?.value ?? []) as Array<{
        Status?: boolean;
        Message?: string;
        User?: string;
      }>;
      const failed = results.filter((r) => r && r.Status === false);
      if (failed.length > 0) {
        return failed
          .map((f) => `${f.User ?? "recipient"}: ${f.Message ?? "refused"}`)
          .join("; ");
      }
    } catch {
      /* an unreadable body after a 200 counts as success — the grant is what matters */
    }
    return undefined;
  };

  const decide = async (
    row: RequestRow,
    approve: boolean,
    text: string,
  ): Promise<void> => {
    setBusy(true);
    setNotice(undefined);
    setNoticeBad(false);
    try {
      /* ⚠ RE-READ BEFORE ACTING, AND BEFORE THE ACTION — NOT MERELY BEFORE THE STATUS WRITE.
         This queue is read at MOUNT, and the requester can withdraw in their own session — so a page
         left open across a cancellation still lists the row, and pressing Approve used to run
         `performShare`/`performDeletion` FIRST. Found on site 2026-08-21: a share was carried out and
         the recipient emailed for a request the uploader had already cancelled, while the row still
         read `Cancelled`. Granting access or recycling a document is not undoable by a status field.
         The mirror of the guard on Cancel in MySubmissions, and it has to come before the action
         rather than before the MERGE, because the action is the part that cannot be taken back.
         FAILS CLOSED: an unreadable check decides nothing and says so. */
      const check = await context.spHttpClient.get(
        /* ⚠ CACHE-BUSTED. This single read is what stands between an approver and re-granting
           access somebody deliberately removed; a cached answer here would defeat the guard
           entirely and it would look like it had passed. */
        `${listUrl()}/items(${row.id})?$select=Status,RequestedBy${bust()}`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!check.ok) {
        setNoticeBad(true);
        setNotice(
          `Could not confirm this request is still open (HTTP ${check.status}). Nothing was done — refresh and try again.`,
        );
        setBusy(false);
        return;
      }
      const current = (await check.json()) as { Status?: string };
      if ((current.Status ?? "").trim() !== "Pending") {
        // Names what happened: "no longer open" alone leaves the approver unsure whether their click
        // did something. Nothing was performed, and the queue is reloaded so the row goes.
        setNoticeBad(true);
        setNotice(
          `Nothing was done — this request is no longer open (${(current.Status ?? "already decided").toLowerCase()}).` +
            ((current.Status ?? "").trim() === "Cancelled"
              ? " The uploader withdrew it."
              : " Someone else decided it first."),
        );
        setDeciding(undefined);
        setNote("");
        setBusy(false);
        await load();
        return;
      }

      let failure: string | undefined;
      if (approve) {
        failure =
          row.type === "Deletion"
            ? await performDeletion(row)
            : await performShare(row);
      }
      const decided = applyDecision(row, {
        approve,
        by: me,
        at: new Date().toISOString(),
        note: text,
        failure,
      });

      const upd = await context.spHttpClient.post(
        `${listUrl()}/items(${row.id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            ...writeHeaders,
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({
            Status: decided.status,
            DecidedBy: decided.decidedBy,
            DecidedAt: decided.decidedAt,
            DecisionNote: decided.decisionNote,
          }),
        },
      );
      // The row could not be updated but the ACTION already happened. Said loudly: silently leaving
      // it Pending invites a second approver to do the same thing again.
      const rowWritten = upd.ok;

      writeAudit(context.spHttpClient, siteUrl, {
        event: approve ? EVENT.requestApproved : EVENT.requestRejected,
        outcome: failure || !rowWritten ? "Failed" : "Success",
        source: "Requests",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: me,
        library: libraryTitle(),
        itemName: row.itemName,
        itemUniqueId: row.itemUniqueId,
        summary: `${row.type} request ${decided.status.toLowerCase()} — ${row.itemName}`,
        details: [
          `Requested by ${row.requestedBy} on ${row.requestedAt}`,
          `Reason: ${row.reason}`,
          row.type === "Share"
            ? `Recipients: ${(row.shareWith ?? []).join(", ")}`
            : "",
          failure ? `The action failed: ${failure}` : "",
          rowWritten
            ? ""
            : "The request row could not be updated, so it may still show as pending.",
        ].filter((d) => d.length > 0),
      }).catch(() => undefined);

      setNoticeBad(failure !== undefined || !rowWritten);
      setNotice(
        failure
          ? `Recorded as failed — ${failure}`
          : !rowWritten
            ? "The action was carried out, but the request could not be updated. Refresh before deciding it again."
            : approve
              ? `Approved. ${decisionSummary(row)}`
              : "Request has been rejected.",
      );
      setDeciding(undefined);
      setNote("");
      await load();
    } catch (e) {
      setNoticeBad(true);
      setNotice(`Could not finish: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  /* Keep the queue current without the approver pressing refresh (client, 2026-08-21).
     Blocked while a decision dialog is open or a write is in flight: replacing the row someone is
     deciding on is worse than being a few seconds stale. */
  useLiveRefresh(load, busy || deciding !== undefined);

  /* ⚠ BRING THE ANSWER TO THE APPROVER. This banner sits at the TOP of the page, above the tabs,
     while the row being decided can be hundreds of pixels below — so a refusal, a failure and a
     success all looked identical from where the click happened: nothing. Reported on site
     2026-08-30 as "I click Approve and the list is still there".

     ⚠ DECLARED WITH THE OTHER HOOKS, ABOVE THE `rows.state` EARLY RETURNS. A `useEffect` below them
     runs a different number of times on the render after loading finishes — "Rendered more hooks
     than during the previous render" — which blanks the whole web part. Same trap as the approval
     panel's probe (1.0.243.0). */
  const noticeRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (notice === undefined) return;
    try {
      noticeRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
    } catch {
      /* an old browser without smooth scrolling must never break the page */
    }
  }, [notice]);

  /* ── The three accordions (client, 2026-09-04) — STATE ONLY; the rest is below ──
   *
   * ⚠⚠ THIS `useState` LIVES UP HERE BECAUSE IT SHIPPED BELOW THE EARLY RETURN AND BLANKED THE WHOLE
   * PAGE (found on site 2026-09-04, reported as *"it shows loading at first then blank"*). The
   * loading render returns at `rows.state === "loading"` and never reaches a hook declared after it,
   * so the render where loading finishes runs ONE MORE hook than the one before — React throws
   * *"Rendered more hooks than during the previous render"* (minified #310), and an SPFx component
   * that throws in render produces NOTHING: no message, no error boundary, a blank web part.
   *
   * ⚠ IT WAS ADDED TWELVE LINES BELOW THE COMMENT THAT WARNS ABOUT EXACTLY THIS, which is the part
   * worth remembering: the trap is not knowing the rule, it is that a `useState` added while writing
   * *rendering* code lands wherever the rendering code is. FOURTH INSTANCE IN THIS PROJECT — the
   * approval panel's probe (1.0.243.0, an hour lost), My Submissions' share picker, `FolderAdmin`'s
   * `maxIdx` (1.0.338.0), now this.
   *
   * ⚠ THE SYMPTOM NAMES THE CAUSE, AND IT IS THE DIAGNOSTIC WORTH KEEPING: *"loading, then blank"*
   * means the component MOUNTED and then threw — so it is never a stale bundle, which cannot render
   * a loading state at all. Blank from the very start is the bundle; blank after something showed is
   * the hooks.
   */
  const [openSections, setOpenSections] = useState<Record<string, boolean>>({});

  /* ── Render ───────────────────────────────────────────────────────────── */

  if (rows.state === "loading") return <p style={s.wrap}>Loading&hellip;</p>;

  const all = rows.state === "ready" ? rows.value : [];
  /* ONE scope object, built once and passed to every rule, so the queue, the empty state and the
     Approve button can never disagree about what this viewer may act on. */
  const scope: ViewerScope = { aprUnits: approverUnits, hodUnits, systemAdmin };
  const decidesSomething =
    systemAdmin || approverUnits.length > 0 || hodUnits.length > 0;
  const queue = queueFor(all, scope);
  /* ⚠ COUNTED OVER WHAT THIS VIEWER CAN SEE, NEVER OVER THE WHOLE LIST (2026-08-21).
     The read returns every row — the list is not security-trimmed per unit — so counting `all` had
     a Head of Department reading "2 pending" beneath a queue saying "Waiting for you (1)". Two
     faults in one line: the page contradicted itself, which reads as a bug, and it disclosed the
     existence of the pending-stage request that `ViewerScope` deliberately hides from them.
     `isVisibleTo` is the SAME rule the cards use, so the tally and the queue cannot disagree. */
  const visible = all.filter((r) => isVisibleTo(r, me, scope));
  const tally = counts(visible);

  /* ── Derived views ────────────────────────────────────────────────────────
     Every one of these filters the SAME `visible` array the tally uses, so a tab and the tally can
     never disagree — the 2026-08-21 defect, where a Head of Department read "2 pending" beneath a
     queue of one, disclosing the existence of a row `ViewerScope` deliberately hides from them. */
  const ofType = (t: RequestType): RequestRow[] =>
    visible.filter((r) => r.type === t);
  const decidedOf = (t: RequestType): RequestRow[] =>
    ofType(t)
      .filter((r) => r.status !== "Pending")
      .slice()
      // Newest first: what an approver came to check is almost always what just happened.
      .sort((a, b) =>
        (b.decidedAt || b.requestedAt || "").localeCompare(
          a.decidedAt || a.requestedAt || "",
        ),
      );

  /* Rows index, live ACL decides — see the module comment in `shared/requests.ts`. `mergeShareAcl`
     turns an absent probe into `unknown`, which is why this is safe to compute before the probe has
     run: it renders as "could not be checked", never as "access ended".

     `currentlyShared` then drops whoever the ACL says is gone, so this tab answers the present-tense
     question its heading asks. It keeps `unknown`, so an unreadable file never silently shortens the
     list — and the revoke itself is still recorded on the request row and shown on the Share tab. */
  const sharedFiles: SharedFile[] = currentlyShared(
    /* ⚠ GONE FILES ARE DROPPED FIRST, BEFORE `mergeShareAcl` EVER SEES THEM. A deleted document
       cannot be reached by anybody, so listing it under "who can reach these documents" answers the
       heading wrongly — and it used to render as `could not be checked`, which reads as a transient
       fault somebody should retry rather than a document that is simply not there. It also carried a
       Revoke button that could never work. Only a 404 qualifies: see `fileGone`.

       ⚠ THE REMOVAL IS SILENT BY THE CLIENT'S DECISION (2026-08-30). It was announced in a note
       counting the dropped files, on the usual rule that a tab about access must not quietly get
       shorter — and the client's answer was that the note itself is what confuses: it raises a
       question ("which documents? deleted by whom?") that this tab cannot answer and nobody asked.
       Deleting a document is not this page's business; the request rows still stand on the Share
       tab as the record. `goneFiles` in `shared/requests.ts` still returns the count if it is ever
       wanted again. */
    withoutGoneFiles(collectSharedFiles(all, scope), acls).map((f) =>
      mergeShareAcl(f, acls[f.itemUniqueId]),
    ),
  );

  /** One request, with the decide buttons only where they can actually be pressed. */
  /* ── The three accordions (client, 2026-09-04) ──────────────────────────────
   *
   * *"now we will have three accordion dropdown instead. One accordion dropdown for Deletion and one
   * for Share and one for Documents with Shared Access."*
   *
   * ⚠ ALL THREE OPEN BY DEFAULT, and that is not a style choice. A collapsed section can HIDE
   * PENDING WORK — the same rule that keeps the status filter from ever hiding a Pending row, and
   * that made "Waiting for you" announce what its text filter had hidden. Opening closed would mean
   * an approver could load this page, see three tidy headers, and miss four requests waiting on
   * them. Collapsing is then their own deliberate act, and the pending count stays in the header
   * while it is shut, so the work is still announced.
   *
   * ⚠ `openSections` ITSELF IS DECLARED WITH THE OTHER HOOKS, ABOVE THE EARLY RETURN — see the note
   * there. Only these two plain helpers live at this point, because nothing below the return may
   * introduce a hook.
   */
  const isOpen = (key: string): boolean => openSections[key] !== false;
  const toggle = (key: string): void =>
    setOpenSections((prev) => ({ ...prev, [key]: prev[key] === false }));

  /** The header of an accordion: chevron, title, and whatever the section wants to show beside it. */
  const accordionHead = (
    key: string,
    emoji: string,
    label: string,
    right?: React.ReactNode,
  ): React.ReactElement => (
    <button
      type="button"
      style={{
        ...s.accHead,
        /* ⚠ THE GAP EXISTS ONLY WHILE THE SECTION IS OPEN (client: *"ensure Delete Rqeuest label do
           not stick with the list box, same goes for the others"*). A constant `marginBottom` would
           leave dead space under every COLLAPSED header, which is the state three stacked sections
           are usually in — so the spacing follows the content it is separating from, not the header. */
        marginBottom: isOpen(key) ? 12 : 0,
      }}
      aria-expanded={isOpen(key)}
      onClick={() => toggle(key)}
    >
      <span aria-hidden="true">{emoji}</span>
      <span style={{ flex: "1 1 auto", textAlign: "left" }}>{label}</span>
      {right}
      {/* ⚠ THE CHEVRON IS LAST, so it sits at the RIGHT END (client's request). It follows `right`
          rather than preceding it, so the pending pill stays beside its label and the arrow is the
          outermost thing on the row — the header's padding is what stops it touching the edge.
          Rotated rather than swapped for a second glyph: one element, and the turn is the affordance. */}
      <span
        aria-hidden="true"
        style={{
          display: "inline-block",
          fontSize: 11,
          color: "#605e5c",
          /* `marginLeft: auto` is deliberately NOT used — the label already carries `flex: 1 1 auto`,
             so it does the pushing. Two things competing to absorb the free space is how the pill
             ends up drifting away from its own label. */
          transform: isOpen(key) ? "rotate(90deg)" : "none",
          transition: "transform .12s ease",
        }}
      >
        &#9654;
      </span>
    </button>
  );

  const requestCard = (
    r: RequestRow,
    actionable: boolean,
  ): React.ReactElement => (
    <div
      key={r.id}
      style={{ ...s.row, borderLeft: "4px solid " + STATUS_ACCENT[r.status] }}
    >
      {/* ⚠ THE BADGE LEADS (client, 2026-09-04: *"For each status tag ... client wants us to attach
          it on the left side of the file name itself"*). It was on the far right of the row, so on a
          wide screen the status and the filename it describes were the width of the page apart and
          a column of badges could not be scanned against a column of names. */}
      <div style={s.rowTop}>
        <span style={{ ...s.pill, ...PILL[r.status] }}>{r.status}</span>
        <span style={s.name}>{r.itemName}</span>
      </div>
      <div style={s.meta}>
        {r.requestedBy} · {r.unit} · {longDate(r.requestedAt)}
        {/* NAMED ON THE ROW. The two stages are deleted from different libraries and mean different
            things — an unapproved draft nobody else has seen, versus a document the unit has been
            filing against. An approver should not have to open the dialog to tell them apart. */}
        {stageOf(r) === "pending" && (
          <strong style={{ color: "#8a4b00" }}>
            {" "}
            · still awaiting approval
          </strong>
        )}
      </div>
      {r.type === "Share" && (
        <div style={s.meta}>
          {(r.shareWith ?? []).join(", ")} · {r.sharePermission ?? "View"}
          {r.expiresAt
            ? " · until " + longDate(r.expiresAt)
            : " · no expiry"}
          {/* Named on the ROW, not only in the dialog: this is the fact that decides the answer, and
              an approver should see it before reaching for a button. */}
          {(r.shareWith ?? []).some((e) => isExternal(e, tenantDomains)) && (
            <strong style={{ color: "#8a4b00" }}>
              {" "}
              · outside the organisation
            </strong>
          )}
        </div>
      )}
      {/* LABELLED 2026-08-30 at the client's request. The box below is the REQUESTER's words, and
          unlabelled it reads like a field the approver is meant to fill in — which is what the
          decision dialog is for. */}
      <div style={s.reasonLabel}>Reason</div>
      <div style={s.reason}>
        {r.reason || <em style={s.quiet}>No reason given.</em>}
      </div>
      {/* WHO DID WHAT — the half this page never showed. A decided request used to vanish unless the
          viewer happened to be the requester, so a Head of Unit could not answer "who approved that,
          and when" without opening the list itself. */}
      {!actionable && r.decidedBy && (
        <div style={s.meta}>
          {r.status === "Revoked" ? "Approved by " : "Decided by "}
          <strong>{r.decidedBy}</strong>
          {r.decidedAt ? " on " + longDate(r.decidedAt) : ""}
        </div>
      )}
      {/* BOTH names, never one replacing the other: who granted the access and who took it away are
          different facts about the same request, and a revoked share still needs the approver on
          record. Absent on rows predating the column, where the revoker's name survives in the
          decision note below. */}
      {!actionable && r.status === "Revoked" && r.revokedBy && (
        <div style={s.meta}>
          Revoked by <strong>{r.revokedBy}</strong>
        </div>
      )}
      {!actionable && r.decisionNote && (
        <div style={s.reason}>{r.decisionNote}</div>
      )}
      {actionable && (
        <div style={s.actions}>
          <button
            style={busy ? s.off : s.approve}
            disabled={busy}
            onClick={() => {
              setNote("");
              setShowNoteError(false);
              setDeciding({ row: r, approve: true });
            }}
          >
            Approve
          </button>
          <button
            style={busy ? s.off : s.reject}
            disabled={busy}
            onClick={() => {
              setNote("");
              setShowNoteError(false);
              setDeciding({ row: r, approve: false });
            }}
          >
            Reject
          </button>
        </div>
      )}
    </div>
  );

  /* ⚠ THE STATUS FILTER NEVER HIDES A PENDING ROW, and that is a safety rule rather than a layout
     one. A Pending row is now IN the same merged list as decided ones (client, 2026-09-03: "remove
     the Waiting for you, but just change the status from pending to Approved") — so without the
     `r.status === "Pending" ||` clause, picking "Approved" from the shared dropdown would silently
     empty out every unit's outstanding work, which is the one outcome no display control may cause.
     The TEXT filter still applies to everything, pending included — matching the pre-merge behaviour
     of the old "Waiting for you" box, which was always text-filterable (with its own warning when it
     hid something). */
  const matchesText = (r: RequestRow): boolean => {
    const q = whoFilter.trim().toLowerCase();
    return (
      q.length === 0 ||
      (r.requestedBy ?? "").toLowerCase().indexOf(q) !== -1 ||
      (r.itemName ?? "").toLowerCase().indexOf(q) !== -1
    );
  };
  const applyFilters = (list: RequestRow[]): RequestRow[] =>
    list.filter(
      (r) =>
        (r.status === "Pending" ||
          statusFilter === "All" ||
          r.status === statusFilter) &&
        matchesText(r),
    );

  /* ⚠ A FILE IS KEPT IF *ANY* RECIPIENT MATCHES, AND ITS RECIPIENT LIST IS NEVER NARROWED. Filtering
     the people inside a file would hide somebody who DOES have access, on the one tab whose whole job
     is answering "who can reach this document" — a disclosure-shaped error, not a display one. */
  const shownShared = sharedFiles.filter((f) => {
    const q = sharedText.trim().toLowerCase();
    const stateOk = sharedState === "All" || f.state === sharedState;
    const textOk =
      q.length === 0 ||
      (f.itemName ?? "").toLowerCase().indexOf(q) !== -1 ||
      f.recipients.some((r) => (r.email ?? "").toLowerCase().indexOf(q) !== -1);
    return stateOk && textOk;
  });
  /* Offered only where a row holds it — so `Could not be checked` appears only when a probe actually
     failed, rather than sitting in the dropdown implying the page is unsure when it is not. */
  const sharedStatesPresent = ["live", "unknown"].filter((v) =>
    sharedFiles.some((f) => f.state === v),
  );

  /* ⚠ MERGED — pending and decided requests of ONE type in a SINGLE list, not two stacked boxes
     (client, 2026-09-03: "remove the Waiting for you, but just change the status from pending to
     Approved"). Pending rows sort first (they are the work); decided rows follow, newest first, from
     `decidedOf` unchanged. `ofType(t)` is already "everything of this type visible to this viewer" —
     splitting it here rather than re-deriving keeps this in step with the tally above. */
  const typeSection = (
    t: RequestType,
    emoji: string,
    label: string,
  ): React.ReactElement => {
    const allPending = ofType(t).filter((r) => r.status === "Pending");
    const decided = decidedOf(t);
    const merged = [...allPending, ...decided];
    const shown = applyFilters(merged);
    const shownPending = shown.filter((r) => r.status === "Pending").length;
    const hiddenPending = allPending.length - shownPending;
    return (
      <div style={s.card}>
        {/* The pending count stays in the HEADER, so it is still announced while the section is
            collapsed — see `openSections` for why that matters. Amber, matching `PILL.Pending`, so
            the header and the cards below it agree about what "pending" looks like. */}
        {accordionHead(
          t,
          emoji,
          label,
          allPending.length > 0 ? (
            <span style={{ ...s.pill, ...PILL.Pending }}>
              {allPending.length} pending
            </span>
          ) : undefined,
        )}
        {!isOpen(t) ? undefined : merged.length === 0 ? (
          <p style={s.quiet}>Nothing here yet.</p>
        ) : (
          <>
            {/* ⚠ SAYS SO OUT LOUD, same rule as the pre-merge "Waiting for you" box: a filter that
                hides pending work must never look like the work is done. Text can hide a pending row
                (as before); status never can — see `applyFilters`. */}
            {hiddenPending > 0 && (
              <p style={s.quiet}>
                {hiddenPending} pending request{hiddenPending === 1 ? "" : "s"}{" "}
                {hiddenPending === 1 ? "is" : "are"} hidden by the filter and
                still need{hiddenPending === 1 ? "s" : ""} a decision — press
                Clear.
              </p>
            )}
            {shown.length === 0 ? (
              <p style={s.quiet}>No request matches the filter.</p>
            ) : (
              /* SCROLLS at 60vh. Safe here, unlike Group Management's group list and the upload
                 form's file rows, because nothing in this card is absolutely positioned. */
              <div style={SCROLLER}>
                {shown.map((r) =>
                  requestCard(r, r.status === "Pending" && canDecide(r, scope)),
                )}
              </div>
            )}
          </>
        )}
      </div>
    );
  };

  /* ⚠ ONE BAR, ABOVE BOTH `typeSection` CALLS, NEVER ONE PER SECTION (client, 2026-09-03: "yes share
     a single filter bar"). The outcome dropdown's options are the UNION of decided statuses across
     BOTH types now, since one control has to serve both. */
  const requestsFilterBar = ((): React.ReactElement | null => {
    const anyRow = ofType("Deletion").length + ofType("Share").length > 0;
    if (!anyRow) return null;
    const present = STATUS_ORDER.filter(
      (v) =>
        decidedOf("Deletion").some((r) => r.status === v) ||
        decidedOf("Share").some((r) => r.status === v),
    );
    const filtering = statusFilter !== "All" || whoFilter.trim().length > 0;
    return (
      <div style={{ ...FIL_WRAP, margin: "0 0 12px" }}>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          style={FIL_SELECT}
        >
          <option value="All">All outcomes</option>
          {present.map((v) => (
            <option key={v} value={v}>
              {v}
            </option>
          ))}
        </select>
        <input
          value={whoFilter}
          onChange={(e) => setWhoFilter(e.target.value)}
          placeholder="Filter by document or requester email"
          style={FIL_INPUT}
        />
        {filtering && (
          <button
            onClick={() => {
              setStatusFilter("All");
              setWhoFilter("");
            }}
            style={FIL_CLEAR}
          >
            Clear
          </button>
        )}
        {/* ⚠ SAID OUT LOUD. The outcome list holds no pending rows, so the dropdown cannot narrow
            pending work — without this line, setting it to `Approved` and watching pending items stay
            put reads as the control being broken. */}
        <span style={s.small}>Outcome applies to decided requests only.</span>
      </div>
    );
  })();

  return (
    <section style={s.wrap}>
      {/* ⚠ RENAMED (client's mockup, 2026-09-03): "Requests" → "Document Deletion & Sharing
          Approval". The page this renders on is a separate, manual SharePoint edit — see
          `docs/superpowers/specs/2026-09-03-requests-page-redesign-scope.md` for the full scope,
          including the still-unbuilt site-navigation restructuring. This heading is the part that
          IS code, so it changes here regardless of when the page/nav side happens. */}
      <h2 style={s.h2}>Document Deletion &amp; Sharing Approval</h2>
      <p style={s.sub}>Approve requests to delete or share documents.</p>

      {/* ⚠ THE "Access Audit" LINK IS GONE (client, 2026-09-04: *"Remove access audit"*). It was
          added on 2026-09-03 with this bar. The Audit Log page itself is untouched and still reachable
          from CRS Settings — only this shortcut went. The `auditLink` state, its Site Pages effect and the
          two now-unused imports went WITH it rather than being parked — an unused local is a lint
          warning, and this project tracks a zero-new-warning baseline. Re-adding it is the
          `resolveLink` + `readSitePages` pattern that five other files already use. */}
      <div style={s.topBar}>
        <span style={{ fontSize: 13, fontWeight: 600 }}>
          Requests {queue.length}
        </span>
      </div>

      {notice && (
        <div ref={noticeRef} style={noticeBad ? s.warn : s.ok}>
          {notice}
        </div>
      )}

      {/* ONE banner for BOTH optional columns, and it must stay that way: this is the only route
          to `addMissingColumns`, so gating it on `stageMissing` alone would leave `RevokedBy`
          unreachable on every site that already has `Stage` — which is all of them. A fix nobody can
          press is not a fix. */}
      {(stageMissing || revokedByMissing) && (
        <div style={s.warn}>
          <strong>
            This list is missing{" "}
            {stageMissing && revokedByMissing ? "two columns" : "a column"}.
          </strong>{" "}
          Requests still work, and nothing already recorded is lost.
          <ul style={{ margin: "8px 0 0 18px", padding: 0 }}>
            {stageMissing && (
              <li>
                <strong>Stage</strong> — a request for a file that is{" "}
                <em>awaiting approval</em> is shown here as though it were an
                approved document, and those live in different libraries.
              </li>
            )}
            {revokedByMissing && (
              <li>
                <strong>RevokedBy</strong> — when a share is revoked, the audit
                log names whoever approved it rather than whoever revoked it. A
                Head of Department can revoke a share a Head of Unit approved,
                so the two are often different people.
              </li>
            )}
          </ul>
          <div style={{ marginTop: 8 }}>
            <button
              style={busy ? s.off : s.approve}
              disabled={busy}
              onClick={() => {
                addMissingColumns().catch(() => undefined);
              }}
            >
              {busy
                ? "Working…"
                : stageMissing && revokedByMissing
                  ? "Add the missing columns"
                  : "Add the missing column"}
            </button>
          </div>
          <div style={{ marginTop: 6, fontSize: 12 }}>
            Adding {stageMissing && revokedByMissing ? "them" : "it"} fixes
            every request raised from now on. Requests already recorded are not
            changed.
          </div>
        </div>
      )}

      {rows.state === "error" && (
        <div style={s.err}>
          Could not read the requests ({rows.message}). This is{" "}
          <strong>not</strong> the same as there being none — do not treat the
          queue as empty until it clears.
        </div>
      )}

      {listMissing && (
        <div style={s.warn}>
          <p style={{ margin: "0 0 8px", fontWeight: 600 }}>
            The requests list does not exist yet
          </p>
          Nobody can raise a request until it does.
          <div style={s.actions}>
            <button
              style={busy ? s.off : s.approve}
              disabled={busy}
              onClick={() => {
                provision().catch(() => undefined);
              }}
            >
              {busy
                ? "Creating…"
                : `Create "${titleForNewList(LIST_SUFFIX.requests)}"`}
            </button>
          </div>
        </div>
      )}

      {/* Not shown to an admin: they already see every row regardless of any department expansion,
          so warning them that a department could not be read would report a limit they do not have. */}
      {!listMissing && hodScopeFailed && !systemAdmin && (
        <div style={s.warn}>
          Your department could not be read, so requests from the units under it
          are <strong>not shown</strong>. This is not a statement that there are
          none. Try again in a moment; if it persists, an administrator should
          check the term store is readable and that your{" "}
          <strong>Head of Department</strong> mapping still points at a live
          department term.
        </div>
      )}

      {/* ⚠ NOT SHOWN TO AN ADMIN — `decidesSomething` now includes them. Before this it fired for
          every owner, and its advice was actively wrong: it told them to have "an administrator add
          an APR or Head of Department mapping for your group", but an admin has no unit group, and
          mapping them would hand them a persona they should not hold and pollute the Group Map that
          reconciliation reads. Fixing the scope without fixing this sentence would have left the
          page contradicting itself — a full queue under a banner saying nothing is waiting. */}
      {!listMissing && !decidesSomething && (
        <div style={s.warn}>
          You are not recorded as the approver for any unit, nor as the head of
          any department, so nothing is waiting on you. If that is wrong, an
          administrator adds an <strong>APR</strong> or{" "}
          <strong>Head of Department</strong> mapping for your group on the{" "}
          <strong>Folder Access</strong> page.
        </div>
      )}

      {/* SAYS WHY THEY CAN SEE EVERYTHING. Without it an admin's queue looks like a leak: they hold
          no mapping for any of these units, so every other screen in this system would show them
          their own rows only. It also states the consequence, because deciding a request here is not
          reversible — an approved deletion recycles the file and an approved share grants access. */}
      {systemAdmin && (
        <div style={s.warn}>
          You are here as a <strong>system administrator</strong>, so you see
          and can decide <strong>every request on this site</strong> — not only
          your own units. Normally the unit&rsquo;s{" "}
          <strong>Head of Unit</strong> decides these; use this when their group
          is empty or nobody else can. Your name is recorded as the decider.
        </div>
      )}

      {hodUnits.length > 0 && approverUnits.length === 0 && !systemAdmin && (
        <div style={s.warn}>
          {/* ⚠ SAYS WHAT THEY DECIDE, NEVER WHY THE REST IS ABSENT (client, 2026-08-21: *"I dont
              even want them to be aware of approval libraries"*). The earlier copy explained the
              limit by naming the approval libraries and the permission a HoD does not hold — true,
              and it described internal structure to someone who has no business with it, in a
              sentence that reads as a shortcoming of their own account. The rule is unchanged; only
              the explanation is gone. */}
          You are here as a <strong>Head of Department</strong>. You decide
          requests about <strong>approved documents</strong> in your department.
          Anything still awaiting approval is handled by the unit&rsquo;s Head
          of Unit.
        </div>
      )}

      {/* ⚠ THE TAB BAR IS GONE — Delete Requests and Share Requests render TOGETHER now, sharing
          the one filter bar above them (client's mockup + confirmed answers, 2026-09-03). */}
      {requestsFilterBar}
      {typeSection("Deletion", "\u{1F5D1}️", "Delete Requests")}
      {typeSection("Share", "\u{1F4E4}", "Share Requests")}

      {/* ⚠ "Shared files" IS A THIRD SECTION NOW, NOT A TAB — always rendered, further down the same
          page (client's mockup: "Putting Shared File tab besides other tabs, doesn't communicate
          objectively the objective"). Retitled to match: "Documents with Shared Access". Everything
          below is otherwise UNCHANGED — same filter, same revoke logic, same admin-only note. */}
      {
        <div style={s.card}>
          {/* WARN: RE-CHECK IS NOT IN THE HEADER, deliberately: the header is a BUTTON that toggles
              the accordion, and a button inside a button is invalid HTML - the inner one swallows
              the click or the outer fires with it. It sits below instead, and only while the section
              is open, which is the only time its result is visible anyway. */}
          {accordionHead(
            "shared",
            "\u{1F517}",
            "Documents with Shared Access (" + String(sharedFiles.length) + ")",
          )}
          {isOpen("shared") && (
            <div style={{ ...s.rowTop, margin: "10px 0" }}>
              <span style={{ flex: "1 1 auto" }} />
              <button
                style={aclsLoading ? s.off : s.ghost}
                disabled={aclsLoading}
                onClick={() => {
                  setAclsRead(false);
                }}
              >
                {aclsLoading ? "Checking…" : "Re-check access"}
              </button>
            </div>
          )}
          {isOpen("shared") && (
            <>
              {/* ⚠ SAYS WHAT IT CANNOT SEE. This tab indexes files that HAVE a request row, so a document
              shared through SharePoint's OWN Share button — which a web part cannot intercept,
              settled 2026-07-23 — never appears here at all. A tab implying completeness about who
              can reach the company's documents would be worse than one that admits its edge. */}
              <p style={{ ...s.quiet, marginTop: 0 }}>
                Shares granted through CRS, checked against each file&rsquo;s
                real permissions. A document shared using SharePoint&rsquo;s own{" "}
                <strong>Share</strong> button is not listed here — only extra
                people found on the files below.
              </p>

              {aclsLoading && (
                <p style={s.quiet}>Checking each file&rsquo;s permissions…</p>
              )}

              {/* ⚠ THE TEXT BOX MATCHES THE DOCUMENT NAME *OR* ANY RECIPIENT'S EMAIL, because the question
              people actually arrive with is "does <person> still have access to anything?" — which a
              filename-only filter cannot answer. Labelled to say so; a box matching two things
              silently is worse than one that matches neither. */}
              {sharedFiles.length > 0 && (
                <div style={FIL_WRAP}>
                  <select
                    value={sharedState}
                    onChange={(e) => setSharedState(e.target.value)}
                    style={FIL_SELECT}
                  >
                    <option value="All">All access</option>
                    {sharedStatesPresent.map((v) => (
                      <option key={v} value={v}>
                        {v === "live" ? "Has access" : "Could not be checked"}
                      </option>
                    ))}
                  </select>
                  <input
                    value={sharedText}
                    onChange={(e) => setSharedText(e.target.value)}
                    placeholder="Filter by document or recipient email"
                    style={FIL_INPUT}
                  />
                  {(sharedState !== "All" || sharedText.length > 0) && (
                    <button
                      onClick={() => {
                        setSharedState("All");
                        setSharedText("");
                      }}
                      style={FIL_CLEAR}
                    >
                      Clear
                    </button>
                  )}
                </div>
              )}

              {sharedFiles.length > 0 && shownShared.length === 0 && (
                <p style={s.quiet}>
                  No document matches these filters. {sharedFiles.length}{" "}
                  {sharedFiles.length === 1 ? "is" : "are"} hidden — press Clear
                  to see {sharedFiles.length === 1 ? "it" : "them"}.
                </p>
              )}

              {sharedFiles.length === 0 ? (
                <p style={s.quiet}>
                  {/* ⚠ THREE STATES, NOT TWO. "Nothing is shared" and "everything that was shared has
                  been revoked" are different facts, and since revoked recipients stopped being
                  listed the second one reaches this branch too — saying there were never any share
                  requests would then contradict the Share tab sitting beside it. */}
                  {rows.state === "error"
                    ? "The requests could not be read, so this is not a statement that nothing is shared."
                    : ofType("Share").length > 0
                      ? "Nobody currently has access through a CRS share — every share here has been revoked or has ended."
                      : "No approved share requests, so CRS has granted nobody access to a document."}
                </p>
              ) : (
                <div style={SCROLLER}>
                  {shownShared.map((f) => {
                    const mayRevoke = canRevoke(f, scope);
                    const livePeople = f.recipients.filter(
                      (r) => r.state === "live",
                    );
                    const workingOn = revoking === f.itemUniqueId;
                    return (
                      <div key={f.itemUniqueId} style={s.fileRow}>
                        <div style={s.rowTop}>
                          <span style={s.name}>{f.itemName}</span>
                          <span style={{ ...s.chip, ...STATE_CHIP[f.state] }}>
                            {STATE_LABEL[f.state]}
                          </span>
                        </div>
                        <div style={s.meta}>
                          {f.segment} · {f.unit}
                        </div>

                        {f.recipients.length === 0 && (
                          <p style={{ ...s.quiet, marginLeft: 14 }}>
                            No recipients were recorded.
                          </p>
                        )}

                        {f.recipients.map((r) => (
                          <div key={r.email} style={s.recip}>
                            <span
                              style={{
                                flex: "1 1 220px",
                                wordBreak: "break-word",
                              }}
                            >
                              {r.email}
                              {isExternal(r.email, tenantDomains) && (
                                <strong style={{ color: "#8a4b00" }}>
                                  {" "}
                                  · outside the organisation
                                </strong>
                              )}
                            </span>
                            <span style={{ ...s.chip, ...STATE_CHIP[r.state] }}>
                              {STATE_LABEL[r.state]}
                            </span>
                            <span style={s.small}>
                              {r.permission === "Edit"
                                ? "can edit"
                                : "can view"}
                              {r.unrecorded
                                ? " · granted outside CRS"
                                : r.decidedBy
                                  ? " · approved by " +
                                    r.decidedBy +
                                    (r.decidedAt
                                      ? " on " + longDate(r.decidedAt)
                                      : "")
                                  : ""}
                              {!r.unrecorded && r.requestedBy
                                ? " · asked by " + r.requestedBy
                                : ""}
                              {r.expiresAt
                                ? " · until " + longDate(r.expiresAt)
                                : ""}
                            </span>
                            {mayRevoke && r.state === "live" && (
                              <button
                                style={workingOn || busy ? s.off : s.revoke}
                                disabled={workingOn || busy}
                                onClick={() => {
                                  const ok = window.confirm(
                                    'Remove access to "' +
                                      f.itemName +
                                      '" for ' +
                                      r.email +
                                      "?" +
                                      "\n\nThey are not told, and any link they have simply stops working.",
                                  );
                                  if (!ok) return;
                                  revoke(f, false, r.email).catch(
                                    () => undefined,
                                  );
                                }}
                              >
                                Revoke
                              </button>
                            )}
                          </div>
                        ))}

                        {mayRevoke && livePeople.length > 0 && (
                          <div style={s.actions}>
                            <button
                              style={workingOn || busy ? s.off : s.revokeAll}
                              disabled={workingOn || busy}
                              onClick={() => {
                                const ok = window.confirm(
                                  'Remove ALL shared access to "' +
                                    f.itemName +
                                    '"?' +
                                    "\n\n" +
                                    livePeople.length +
                                    " recipient(s) lose access immediately, and " +
                                    "the file goes back to its unit's normal permissions — everyone in the " +
                                    "unit keeps the access they already had." +
                                    "\n\nNobody is notified.",
                                );
                                if (!ok) return;
                                revoke(f, true).catch(() => undefined);
                              }}
                            >
                              {workingOn ? "Removing…" : "Revoke all access"}
                            </button>
                            {/* ⚠ SAID OUT LOUD, because it is the difference between tidying a list and
                          fixing the problem: only this button releases the file's permission scope,
                          and unreleased scopes are what the 50,000-per-list ceiling counts. */}
                            <span style={s.small}>
                              Also releases this file&rsquo;s permission scope.
                            </span>
                          </div>
                        )}

                        {!mayRevoke && livePeople.length > 0 && (
                          <p style={{ ...s.quiet, marginLeft: 14 }}>
                            Only this unit&rsquo;s Head of Unit, its Head of
                            Department or an administrator can remove access.
                          </p>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}

              {/* ⚠ ADMINS ONLY, AND THAT IS NOT TIDINESS — A HEAD OF UNIT CANNOT USE THIS.
              `CRS Share` carries Manage Permissions at the UNIT FOLDER, never at library scope, so an
              HoU or HoD opening a library's Permissions page gets AccessDenied. Showing them this
              would be an instruction they cannot follow, on the screen built for them. */}
              {systemAdmin && sharedFiles.length > 0 && (
                <div
                  style={{
                    ...s.recip,
                    borderLeft: "2px solid #f0c36d",
                    marginTop: 14,
                  }}
                >
                  <p style={{ ...s.quiet, margin: 0 }}>
                    <strong>
                      After revoking, SharePoint can leave a traversal entry
                      behind.
                    </strong>{" "}
                    It grants nothing — the person sees an empty library — but
                    to clear it by hand:{" "}
                    <a
                      href={`${siteUrl}/_layouts/15/user.aspx`}
                      target="_blank"
                      rel="noreferrer"
                    >
                      open permissions
                    </a>
                    , pick the library, then <em>Show users</em> in the yellow
                    banner, tick their name and choose{" "}
                    <em>Remove User Permissions</em>.{" "}
                    <strong>
                      Remove only rows whose Type is &ldquo;User&rdquo;. Never
                      tick the header checkbox: every group there also shows
                      Limited Access, and removing those takes away every
                      group&rsquo;s folder access across the whole library.
                    </strong>{" "}
                    Using <em>Revoke all access</em> above usually avoids the
                    leftover entirely, because it releases the file&rsquo;s
                    permission scope.
                  </p>
                </div>
              )}
            </>
          )}
        </div>
      }

      <p style={s.quiet}>
        {tally.Pending} pending · {tally.Approved} approved · {tally.Rejected}{" "}
        rejected · {tally.Revoked} revoked · {tally.Failed} failed
      </p>

      {deciding && (
        <div
          style={s.modalBg}
          onClick={() => {
            if (!busy) setDeciding(undefined);
          }}
        >
          <div style={s.modal} onClick={(e) => e.stopPropagation()}>
            <p style={{ ...s.head, fontSize: 16, margin: "0 0 4px" }}>
              {deciding.approve ? "Approve" : "Reject"} this request?
            </p>
            {/* ⚠ THE FILENAME IS SHOWN ON BOTH, though the client's REJECT mockup omits it. Only the
                Approve mockup carried it, and dropping it from Reject would leave an approver
                refusing a request with nothing on screen saying WHICH document — the one fact they
                cannot reconstruct once the dialog is open. Adding information the mockup left out is
                the safe direction; say so if it is unwanted. */}
            <p style={s.modalFile}>{deciding.row.itemName}</p>
            {/* ⚠ APPROVE KEEPS `decisionSummary`, WHICH ALREADY IS THE MOCKUP'S SENTENCE — *"This
                file will be moved to the recycle bin, where it can be restored for 93 days."* It must
                never be replaced by a fixed string: for a SHARE it names the recipients and the
                permission instead, and a hardcoded recycle-bin line would describe the wrong action
                entirely.
                ⚠ 93 DAYS, NOT THE 90 IN THE MOCKUP — the same conflict settled twice already
                (2026-08-30 and 2026-09-03: *"stay 93, they might not know that is why they say 90"*).
                SharePoint's site recycle bin really is 93, and that number is the only thing making an
                approved deletion reversible.
                REJECT now takes the mockup's prompt. It replaces *"X will not be touched, and Y will
                see your note"* — both facts still true, but the prompt is what the approver needs at
                the moment they are being asked to type. */}
            <p style={{ fontSize: 13, margin: "0 0 4px" }}>
              {deciding.approve
                ? decisionSummary(deciding.row)
                : "Please provide your reasoning below."}
            </p>
            {deciding.approve &&
              deciding.row.type === "Share" &&
              (deciding.row.shareWith ?? []).some((e) =>
                isExternal(e, tenantDomains),
              ) && (
                <div style={s.warn}>
                  This sends the document{" "}
                  <strong>outside the organisation</strong>. It stays accessible
                  until the share is removed
                  {deciding.row.expiresAt
                    ? ` or ${longDate(deciding.row.expiresAt)} passes`
                    : ""}
                  .
                </div>
              )}
            {/* ⚠ REQUIRED ON REJECT, OPTIONAL ON APPROVE (client's mockup, 2026-09-04: *"Note
                (optional)"* on Approve, *"Reason"* on Reject).

                ⚠ THIS NARROWS MY OWN OVER-APPLICATION RATHER THAN REVERSING THE CLIENT. Their
                2026-09-03 instruction was about rejecting — *"IF APPROVER WANTS TO REJECT A REQUEST …
                APPROVER MUST INCLUDE THE REASON … ELSE SYSTEM DOESN'T ALLOW TO PROCEED"* — and I
                extended it to approvals as well. The mockup settles it: a refusal has to be explained
                because the requester must know what to fix; an approval speaks for itself, since the
                thing they asked for simply happened. */}
            <label
              style={{
                display: "block",
                marginTop: 12,
                marginBottom: 4,
                fontSize: 12,
                fontWeight: 600,
              }}
            >
              {deciding.approve ? "Note (optional)" : "Reason"}
            </label>
            <textarea
              style={
                !deciding.approve && showNoteError && note.trim() === ""
                  ? {
                      ...s.noteArea,
                      borderColor: "#a4262c",
                      background: "#fdf6f6",
                    }
                  : s.noteArea
              }
              value={note}
              onChange={(e) => {
                setNote(e.target.value);
                if (showNoteError) setShowNoteError(false);
              }}
            />
            {/* Shown only after a rejection was ATTEMPTED with nothing typed — never on open, which
                would mark a dialog nobody has used yet as being in error. */}
            {!deciding.approve && showNoteError && note.trim() === "" && (
              <p
                style={{ fontSize: 11.5, color: "#a4262c", margin: "4px 0 0" }}
              >
                Type a reason before rejecting — the requester needs to see why.
              </p>
            )}
            <div style={s.actions}>
              <button
                style={
                  busy || !canDecide(deciding.row, scope)
                    ? s.off
                    : deciding.approve
                      ? s.approveDark
                      : s.reject
                }
                disabled={busy || !canDecide(deciding.row, scope)}
                onClick={() => {
                  /* ⚠ REJECT ONLY. An approval with no note proceeds — see the label comment. */
                  if (!deciding.approve && note.trim() === "") {
                    setShowNoteError(true);
                    return;
                  }
                  decide(deciding.row, deciding.approve, note).catch(
                    () => undefined,
                  );
                }}
              >
                {busy ? "Working…" : deciding.approve ? "Approve" : "Reject"}
              </button>
              <button
                style={s.ghost}
                disabled={busy}
                onClick={() => setDeciding(undefined)}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

/** A list item as this screen needs it. Field names are the list's, not the model's. */
function fromListItem(r: Record<string, string>): RequestRow {
  return {
    id: Number(r.Id),
    type: (r.RequestType === "Share"
      ? "Share"
      : "Deletion") as RequestRow["type"],
    // Absent column or absent value both mean an approved document — see stageOf.
    stage: r.Stage === "pending" ? "pending" : "approved",
    /* ⚠ ONE definition of the statuses, in `shared/requests.ts`. This was a hand-written allow-list
       missing `Revoked` and `Cancelled`, so both read as Pending and sat in the approver's queue
       permanently — undecidable, because the re-read guard rightly refuses a row that is no longer
       open. Never inline this list again. */
    status: parseRequestStatus(r.Status),
    itemUniqueId: r.ItemUniqueId ?? "",
    itemName: r.ItemName ?? "",
    itemUrl: r.ItemUrl ?? "",
    segment: r.Segment ?? "",
    unit: r.Unit ?? "",
    unitTermGuid: r.UnitTermGuid ?? "",
    requestedBy: r.RequestedBy ?? "",
    requestedAt: r.RequestedAt ?? "",
    reason: r.Reason ?? "",
    shareWith: (r.ShareWith ?? "").split(/[,;\s]+/).filter((x) => x.length > 0),
    sharePermission: (r.SharePermission === "Edit"
      ? "Edit"
      : "View") as RequestRow["sharePermission"],
    expiresAt: r.ExpiresAt ?? "",
    decidedBy: r.DecidedBy ?? "",
    decidedAt: r.DecidedAt ?? "",
    decisionNote: r.DecisionNote ?? "",
    revokedBy: r.RevokedBy ?? "",
  };
}
