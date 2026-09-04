/**
 * Deletion and share requests — the rules, with no React and no network.
 *
 * Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
 *       docs/superpowers/specs/2026-08-28-requests-page-redesign-and-share-revoke-design.md
 *
 * A PIC cannot delete or share in `Documents`; they ask, and a Head of Unit decides. Both request
 * types are ONE object with a `RequestType`, because they are the same shape — ask, decide, act — and
 * built separately they become two screens that drift, with two ideas of what "pending" means.
 *
 * THE PRINCIPLE THIS MODULE PROTECTS: the approval executes in the approver's own browser session.
 * Nothing here elevates anything. So the questions it answers are "is this request well formed", "may
 * this person decide it", and "what does the row say afterwards" — never "how do I get permission".
 *
 * Dates are ISO strings throughout. `requestedAt` and `decidedAt` go to a plain `/items` POST, which
 * parses through the OData layer and REJECTS a locale string — gotcha #1's `M/D/YYYY` belongs to
 * `validateUpdateListItem` and would 400 here (the audit-log lesson).
 */

export type RequestType = "Deletion" | "Share";

/**
 * `Failed` is a first-class outcome, not an error to swallow.
 *
 * An approver can approve a deletion for a file that has since been moved or deleted, or a share the
 * tenant refuses. Recording that as `Approved` teaches the approver the screen is lying; recording it
 * as `Rejected` blames them for a decision they did not make.
 */
export type RequestStatus =
  | "Pending"
  | "Approved"
  | "Rejected"
  | "Failed"
  | "Cancelled"
  /**
   * An approved SHARE whose access has since been taken back (2026-08-28).
   *
   * Terminal, and it supersedes `Approved` rather than contradicting it: the request WAS approved,
   * and the grant it produced has been withdrawn. Written only when the LAST recipient goes — a
   * partial revoke leaves the row `Approved` and appends to `DecisionNote`, because the request's
   * own outcome has not changed, only part of its grant.
   *
   * ⚠ Safe to add because `Status` is a **Text** column, never Choice. That was a deliberate choice
   * in the 2026-08-15 design for exactly this moment: a value absent from a Choice column's
   * `Choices` fails the whole write, silently, so every row of the new kind would be lost.
   */
  | "Revoked";

export type SharePermission = "View" | "Edit";

/**
 * Which side of approval the document sits on — and therefore which library holds it.
 *
 * Added 2026-08-20, when the PIC lost `DELS` (client: *"For Staging PIC should not be able to delete,
 * they have to request from HOU"*). Until then a PIC deleted pending and rejected files outright and
 * this workflow only ever saw APPROVED documents, so the distinction did not exist.
 *
 * `"approved"` — in `Documents` / `HC Documents`. Deletion or share, both decided by the Head of Unit.
 * `"pending"`  — in an approval library, still awaiting or refused approval. **Deletion only.**
 *
 * OPTIONAL, defaulting to `"approved"`, because rows written before today carry no value and every
 * one of them was an approved document. Same shape as `unitTermGuid`: a new field that must not
 * strand the rows that predate it.
 */
export type RequestStage = "approved" | "pending";

/** `stage` with the pre-2026-08-20 default applied. Never read `row.stage` directly. */
export function stageOf(row: { stage?: RequestStage } | undefined): RequestStage {
  return (row ?? {}).stage === "pending" ? "pending" : "approved";
}

/**
 * May this stage be SHARED, as opposed to only deleted?
 *
 * Sharing a pending file would hand someone a document **nobody has approved yet** — the same
 * objection that keeps `SHARE` out of `LIBRARY_ROLES` for both approval libraries, so a Head of Unit
 * could not perform such a share even if the screen offered it. Refused here rather than hidden
 * alone: a hidden control is a UI decision, and this is a rule.
 */
export function canShareStage(stage: RequestStage | undefined): boolean {
  return stage !== "pending";
}

export interface RequestRow {
  id?: number;
  type: RequestType;
  status: RequestStatus;
  /**
   * Which side of approval the file is on. Absent means `"approved"` — read it through `stageOf`,
   * never directly, or a pre-2026-08-20 row reads as neither.
   */
  stage?: RequestStage;
  /** The file's UniqueId — survives rename and move, which a URL does not. */
  itemUniqueId: string;
  itemName: string;
  itemUrl?: string;
  segment: string;
  /** The unit's LABEL — for display only. Never the matching key: terms get renamed. */
  unit: string;
  /**
   * The unit term's GUID, and the key an approver's queue actually matches on.
   *
   * Deleting a term orphans three lists at once and a rename changes every label, so a queue keyed on
   * the label would quietly stop showing an approver their own unit's requests the day someone tidied
   * up the term store — with nothing on screen to say a request had gone missing. Falls back to the
   * label when absent, so rows written before this existed still route.
   */
  unitTermGuid?: string;
  requestedBy: string;
  requestedAt: string;
  reason: string;
  /** Share only. */
  shareWith?: string[];
  sharePermission?: SharePermission;
  expiresAt?: string;
  decidedBy?: string;
  decidedAt?: string;
  decisionNote?: string;
  /**
   * Who revoked the share — SEPARATE from `decidedBy`, which must go on recording who let it
   * through. A Head of Department can revoke a share a Head of Unit approved, so collapsing the two
   * would credit the approver with someone else's revocation in the audit log. Absent on every row
   * written before 2026-08-30 and on any site whose Requests list lacks the column.
   */
  revokedBy?: string;
}

/** What the raise-a-request form collects, before it becomes a row. */
export interface RequestDraft {
  type: RequestType;
  /** Absent means `"approved"`, matching `RequestRow`. */
  stage?: RequestStage;
  itemUniqueId: string;
  itemName: string;
  /**
   * Has the document passed seven years and moved to the archive? (2026-08-22)
   *
   * Optional and defaulting to false, so every existing caller keeps today's behaviour — and so a
   * site with no archive can never accidentally refuse a valid request.
   */
  archived?: boolean;
  segment: string;
  unit: string;
  reason: string;
  shareWith?: string;
  sharePermission?: SharePermission;
  expiresAt?: string;
}

/* ── Recipients ─────────────────────────────────────────────────────────────── */

/**
 * Split a typed recipient list.
 *
 * Accepts commas, semicolons, spaces and newlines, because people paste from Outlook and from a
 * spreadsheet. Lower-cased and de-duplicated: granting the same person twice is not worth refusing,
 * but it should not produce two grants either.
 */
export function parseRecipients(raw: string): string[] {
  const parts = (raw ?? "")
    .split(/[,;\s]+/)
    .map((s) => s.trim().toLowerCase())
    .filter((s) => s.length > 0);
  const seen: string[] = [];
  for (const p of parts) if (seen.indexOf(p) === -1) seen.push(p);
  return seen;
}

/**
 * Deliberately permissive: one `@`, something either side, a dot in the domain.
 *
 * Stricter validation rejects real addresses (apostrophes, plus-addressing, long TLDs), and the cost
 * of a false reject here is a person unable to ask for something they are entitled to. A malformed
 * address that gets through fails visibly at the grant, which is a better place to find out.
 */
export function looksLikeEmail(value: string): boolean {
  return /^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test((value ?? "").trim());
}

/**
 * Is this recipient outside the organisation?
 *
 * Compared against the tenant's own domains, supplied by the caller — hardcoding `sdguthrie.com`
 * would classify every recipient as external on the test tenant and block the lot.
 *
 * NO DOMAINS SUPPLIED MEANS UNKNOWN, AND UNKNOWN COUNTS AS EXTERNAL. The one place in this module
 * that fails closed, matching the spec: everywhere else a failed read costs a form, here it would
 * send a document out of the organisation on the strength of something nobody could confirm.
 */
export function isExternal(email: string, tenantDomains: string[]): boolean {
  const at = (email ?? "").lastIndexOf("@");
  if (at < 0) return true;
  const domain = email.slice(at + 1).toLowerCase();
  const known = (tenantDomains ?? [])
    .map((d) => (d ?? "").trim().toLowerCase())
    .filter((d) => d.length > 0);
  if (known.length === 0) return true;
  return known.indexOf(domain) === -1;
}

/* ── Validation ─────────────────────────────────────────────────────────────── */

export interface ValidationContext {
  /** From the `DMS Config` row. Absent or unreadable must arrive here as `false` — see the spec. */
  allowExternal: boolean;
  tenantDomains: string[];
  /** Local date for the expiry check, so no string is ever parsed as UTC. */
  today?: Date;
}

/**
 * Everything wrong with a draft, in the order a person would fix it.
 *
 * Messages rather than a boolean: "invalid" gives someone nothing to act on, and this form is filled
 * in by an uploader who cannot ask an administrator what they got wrong.
 */
export function validateDraft(draft: RequestDraft, ctx: ValidationContext): string[] {
  const out: string[] = [];
  const d = draft ?? ({} as RequestDraft);

  if (!(d.itemUniqueId ?? "").trim()) out.push("No document selected.");
  // REQUIRED. An approver deciding with no reason in front of them will approve, every time — which
  // makes the whole workflow theatre.
  if (!(d.reason ?? "").trim()) out.push("Give a reason — the approver sees only this.");

  /* ⚠ AN ARCHIVED DOCUMENT CANNOT BE DELETED OR SHARED — BOTH types, unlike the pending rule below
     which refuses Share alone.

     THE FAILURE THIS PREVENTS IS SILENT. The approval executes in the APPROVER'S OWN SESSION with
     their own permissions, and every role resolves to Read on the archive (`READ_ONLY_LIBS`). So an
     approved request would fail at the last step — recorded `Failed`, days after the requester was
     told it was being handled, with the document still there.

     REFUSED IN THE RULES, NOT HIDDEN IN THE UI. Hiding a button turns a permission rule into a
     presentation decision, and the next screen to render a request forgets it. Same reasoning as the
     pending-share refusal directly below, which was written for the same reason.

     FIRST, and alone: the message says archiving is what changed, and a second complaint about a
     missing recipient beside it would read as something the requester could fix. */
  if (d.archived === true) {
    out.push(
      "This document has been archived, so it can no longer be deleted or shared. " +
        "Archived documents are kept as a read-only record — everyone who could read it still can, " +
        "but nobody can change it. Ask an administrator if it genuinely has to be removed.",
    );
    return out;
  }

  // A pending file cannot be shared, and this must REFUSE rather than rely on the form not offering
  // it. `LIBRARY_ROLES` keeps SHARE off both approval libraries, so an approved share would fail at
  // the last step in the approver's own session — after they had told the requester yes.
  if (d.type === "Share" && !canShareStage(d.stage)) {
    out.push(
      "This file is still awaiting approval, so it cannot be shared — only deleted. " +
        "Share it once it has been approved and moved into the Documents library.",
    );
  }

  if (d.type === "Share" && canShareStage(d.stage)) {
    const list = parseRecipients(d.shareWith ?? "");
    if (list.length === 0) {
      out.push("Add at least one person to share with.");
    } else {
      const bad = list.filter((e) => !looksLikeEmail(e));
      if (bad.length > 0) out.push(`Not an email address: ${bad.join(", ")}.`);

      const domains = (ctx?.tenantDomains ?? []).filter((d) => (d ?? "").trim().length > 0);
      const external = list.filter((e) => looksLikeEmail(e) && isExternal(e, domains));
      if (external.length > 0 && !(ctx && ctx.allowExternal)) {
        /* WARN: TWO STATES REACHED THIS MESSAGE AND IT NAMED ONLY ONE OF THEM. `isExternal` fails
           CLOSED — no domains supplied means UNKNOWN, and unknown counts as external — so a site
           with no `tenantDomains` row refuses EVERY address while saying each one is "outside the
           organisation". The client hit exactly that on 2026-09-03: a colleague on their own tenant
           was refused, with a sentence asserting they were an outsider and pointing at a setting
           that was not the problem.
           The refusal is UNCHANGED (guessing internal on an unconfigured site is how a document
           leaves the organisation on the strength of nothing); only the reason it gives is now the
           true one, and it names the fix that actually applies. */
        out.push(
          domains.length === 0
            ? "This site has not been told which email domains belong to your organisation, so every "
              + "address counts as outside it and no share request can be sent. An administrator sets "
              + "tenantDomains on the CRS Config list."
            : `Sharing outside the organisation is switched off, so ${external.join(", ")} cannot be added. `
              + `Your organisation is ${domains.join(", ")}. An administrator changes this on the CRS Config list.`,
        );
      }
    }

    const expires = (d.expiresAt ?? "").trim();
    if (expires.length > 0) {
      const today = ctx?.today ?? new Date();
      // Compared as strings against a locally-built ISO date. Never `new Date(iso)` — parsed as UTC,
      // it lands on the previous day west of Greenwich, so today would read as already past.
      const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
      const t = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
      if (expires.slice(0, 10) < t) out.push("The expiry date has already passed.");
    }
  }

  return out;
}

/* ── Who may decide ─────────────────────────────────────────────────────────── */

/**
 * Whose requests may this viewer act on?
 *
 * TWO SETS, and they are not interchangeable (2026-08-21, spec
 * `2026-08-21-requests-page-hod-access-design.md`):
 *   - `aprUnits` — units this person APPROVES for, from their `APR` mappings. Any stage.
 *   - `hodUnits` — units under a department this person HEADS, expanded from their `DEPTVIEW`
 *     mappings. **APPROVED-stage only.**
 *
 * ⚠ THE STAGE LIMIT ON `hodUnits` IS NOT A POLICY CHOICE — IT IS WHAT A HoD CAN PHYSICALLY DO.
 * `hod` is `DEPTVIEW + DEL + SHARE`, and none of those three is in `LIBRARY_ROLES.Staging`, so a
 * Head of Department holds NOTHING in either approval library. Since an approval executes in the
 * approver's own browser session, approving a pending-file deletion would fail as them and record
 * `Failed` — the requester told their request was handled, and nothing done. So those rows are
 * hidden from a HoD rather than shown un-actionable (client's choice, 2026-08-21).
 *
 * A person who is a Head of Unit somewhere AND a Head of Department elsewhere gets the UNION: the
 * sets are additive, never an override.
 */
export interface ViewerScope {
  aprUnits: string[];
  hodUnits: string[];
  /**
   * A system administrator — a member of the site's OWNERS group, or a site collection admin.
   *
   * ⚠ NOT A UNIT SCOPE, AND DELIBERATELY NOT EXPRESSED AS ONE. Client, 2026-08-27: *"their version
   * of system admin means they have access to everything."* Reported the same day: an owner opened
   * this page and saw only the four requests they had raised themselves, under a banner telling them
   * to have *"an administrator add an APR or Head of Department mapping for your group"* — wrong
   * advice, because an admin holds NO Group Map rows by design, and mapping them would hand them a
   * persona they should not have and pollute the list reconciliation reads.
   *
   * ⚠ WHY THIS IS SAFE HERE WHEN IT IS NOT FOR A HEAD OF DEPARTMENT: approving EXECUTES in the
   * decider's own session with their own permissions. `hod` is barred from pending-stage requests
   * because it holds nothing in either approval library, so its Approve would fail AFTER telling the
   * requester yes. An owner holds Full Control, so the recycle or the share genuinely succeeds.
   *
   * The operational need is real: `GHO_GCA_EG_APPROVER` sat EMPTY on 2026-08-27. With an approver
   * group empty or its Head of Unit gone, requests pile up with nobody able to decide, and the only
   * workaround was adding yourself to a unit's approver group — the pollution above.
   */
  systemAdmin?: boolean;
}

/**
 * Accept either shape.
 *
 * A bare `string[]` has always meant "the units this person approves for", so it is read as
 * `aprUnits` — which keeps every existing caller and test correct rather than rewriting them around
 * a new type. Nothing is inferred: an array can never grant HoD scope, and it can never grant
 * `systemAdmin` either — the widest right in this file must not arrive from the oldest call shape.
 */
export function toScope(scope: string[] | ViewerScope | undefined): ViewerScope {
  if (Array.isArray(scope)) return { aprUnits: scope, hodUnits: [], systemAdmin: false };
  const s = scope ?? { aprUnits: [], hodUnits: [] };
  return {
    aprUnits: s.aprUnits ?? [],
    hodUnits: s.hodUnits ?? [],
    // `=== true` so a truthy non-boolean cannot widen this, and `undefined` reads as NOT an admin.
    systemAdmin: s.systemAdmin === true,
  };
}

/**
 * May this person decide this request?
 *
 * Two conditions, and the second is the one easily forgotten: the request must still be PENDING.
 * Without it, two approvers on the same queue both press Approve and the action runs twice — a file
 * deleted twice is harmless, but a share granted twice leaves a second permission scope nobody tracks.
 *
 * `scope` is what this viewer may act on — see `ViewerScope`. An empty scope decides nothing, which
 * is what a PIC gets.
 */
export function canDecide(row: RequestRow, scope: string[] | ViewerScope): boolean {
  if (!row || row.status !== "Pending") return false;
  return inScope(row, scope);
}

/**
 * The minimum a thing needs for `inScope` to judge it.
 *
 * ⚠ STRUCTURAL ON PURPOSE, so a `SharedFile` can be scope-checked by the SAME predicate that judges
 * a `RequestRow` — with no cast, and with no second copy of the rule. A `RequestRow` still satisfies
 * it, so every existing caller is unchanged. Writing a separate `canRevoke` predicate was the
 * obvious alternative and is exactly how the revoke button and the queue would come to disagree
 * about which units a Head of Department may act on.
 */
type UnitScoped = { unit?: string; unitTermGuid?: string; stage?: RequestStage };

/**
 * Does this row belong to one of these units?
 *
 * Matches on the term GUID when the row carries one, and only falls back to the label otherwise —
 * see `RequestRow.unitTermGuid`.
 */
function matchesUnit(row: UnitScoped | undefined, approverUnits: string[]): boolean {
  const key = ((row?.unitTermGuid ?? "").trim() || (row?.unit ?? "").trim()).toLowerCase();
  if (key.length === 0) return false;
  const units = (approverUnits ?? [])
    .map((u) => (u ?? "").trim().toLowerCase())
    .filter((u) => u.length > 0);
  return units.indexOf(key) !== -1;
}

/**
 * May this viewer act on this row at all, ignoring its status?
 *
 * Shared by `canDecide` and `isVisibleTo` so the two can never disagree about which rows a Head of
 * Department is allowed to know about. Hidden means hidden — a row they cannot decide is a row they
 * do not see, because a visible-but-inert row reads as a broken button.
 */
function inScope(row: UnitScoped | undefined, scope: string[] | ViewerScope | undefined): boolean {
  const s = toScope(scope);
  /* A system administrator acts on EVERY unit and BOTH stages — see `ViewerScope.systemAdmin`.
     Placed in `inScope` rather than in the two callers so visibility and decide-rights cannot
     diverge: the invariant that a row they cannot decide is a row they do not see holds for an
     admin exactly as it does for everybody else. `canDecide` still requires the row to be PENDING,
     so this grants no power to re-decide something already settled. */
  if (s.systemAdmin === true) return true;
  if (matchesUnit(row, s.aprUnits)) return true;
  // APPROVED stage only for a department head. `stageOf` defaults to "approved", so every row
  // written before the `Stage` column existed is treated as the approved document it was.
  return stageOf(row) === "approved" && matchesUnit(row, s.hodUnits);
}

/**
 * Can this person see the request at all?
 *
 * Their own, or one for a unit they approve. NOT a confidentiality boundary — every PIC already reads
 * every approved document in their unit, so a request row names no file they could not already see.
 * It keeps the queue about the reader rather than showing them everyone else's admin.
 */
export function isVisibleTo(
  row: RequestRow,
  viewerEmail: string,
  scope: string[] | ViewerScope,
): boolean {
  const me = (viewerEmail ?? "").trim().toLowerCase();
  if (me.length > 0 && (row?.requestedBy ?? "").trim().toLowerCase() === me) return true;
  return inScope(row, scope);
}

/**
 * May this person withdraw this request?
 *
 * Added 2026-08-20 (client: *"they can cancel the request"*). Two conditions, and both are the point:
 *   - **Only the REQUESTER.** An approver already has Reject, which is a decision and is recorded as
 *     one. Letting them cancel instead would let a decision be taken with no decision on record.
 *   - **Only while PENDING.** Cancelling an approved deletion cannot un-recycle the file, and
 *     cancelling an approved share cannot take the access back — the row would say the request was
 *     withdrawn while the thing it asked for had already happened. `revoke` is a separate, unbuilt
 *     feature (spec §5.3); this must not look like it.
 *
 * Compared lower-cased and trimmed, because `RequestedBy` is stored text and a stored address can
 * differ in case from the one the page reads off the signed-in user.
 */
export function canCancel(row: RequestRow, viewerEmail: string): boolean {
  if (!row || row.status !== "Pending") return false;
  const me = (viewerEmail ?? "").trim().toLowerCase();
  return me.length > 0 && (row.requestedBy ?? "").trim().toLowerCase() === me;
}

/* ── Decisions ──────────────────────────────────────────────────────────────── */

export interface Decision {
  approve: boolean;
  by: string;
  at: string;
  note?: string;
  /** Set when the action itself failed after an approval — see `RequestStatus`. */
  failure?: string;
}

/**
 * The row as it stands after a decision.
 *
 * A FAILED ACTION IS RECORDED AS FAILED, never as approved. That distinction is the whole value of the
 * row: "approved" against a file that could not be deleted, or a share the tenant refused, tells the
 * next person the job is done when it is not.
 */
export function applyDecision(row: RequestRow, decision: Decision): RequestRow {
  const status: RequestStatus = !decision.approve
    ? "Rejected"
    : decision.failure
      ? "Failed"
      : "Approved";
  return {
    ...row,
    status,
    decidedBy: decision.by,
    decidedAt: decision.at,
    // The failure reason wins the note field: it is what someone needs in order to act, and a
    // rejection note and a failure reason can never both apply.
    decisionNote: decision.failure ?? decision.note ?? "",
  };
}

/** Pending requests for the units this person approves, oldest first — a queue, not a list. */
export function queueFor(rows: RequestRow[], scope: string[] | ViewerScope): RequestRow[] {
  return (rows ?? [])
    .filter((r) => canDecide(r, scope))
    .sort((a, b) => (a.requestedAt ?? "").localeCompare(b.requestedAt ?? ""));
}

/**
 * Every status, as DATA — and the compiler is what keeps it complete.
 *
 * ⚠ THE BUG THIS EXISTS TO END, found live 2026-08-30: the Requests page parsed a row with a
 * hand-written allow-list, `["Pending", "Approved", "Rejected", "Failed"]`, and anything outside it
 * fell back to **Pending**. `Revoked` and `Cancelled` were added to the union for the redesign and
 * never added to that literal — so a revoked share and a withdrawn request both came back to the
 * approver's queue reading `Pending`, FOR EVER. They could not be cleared: the pre-write re-read
 * guard correctly refused to decide a row that was no longer open, so every click answered
 * *"Nothing was done — this request is no longer open (revoked)"* while the row stayed put.
 *
 * The fallback's own comment defended itself — *"anything unrecognised reads as Pending, so a row
 * written by a newer version is still decidable rather than invisible"* — and that reasoning was
 * exactly backwards for the two statuses that reached it: both are TERMINAL, and presenting a
 * terminal row as decidable is what produced the phantom.
 *
 * Written as a `Record<RequestStatus, true>` on purpose: adding a member to the union without
 * adding it here is a COMPILE ERROR, where a `RequestStatus[]` literal would silently be short.
 */
const STATUS_SET: Record<RequestStatus, true> = {
  Pending: true, Approved: true, Rejected: true, Failed: true, Cancelled: true, Revoked: true,
};

export const REQUEST_STATUSES: RequestStatus[] = Object.keys(STATUS_SET) as RequestStatus[];

/**
 * The stored `Status` text as a `RequestStatus`.
 *
 * An unrecognised value still reads as `Pending`, and that is now safe rather than lucky: the set
 * above cannot be short, so the only way to reach the fallback is a row written by something other
 * than this code — which a human should see in the queue rather than have hidden.
 */
export function parseRequestStatus(raw: string | undefined): RequestStatus {
  const value = (raw ?? "").trim();
  return STATUS_SET[value as RequestStatus] === true ? (value as RequestStatus) : "Pending";
}

/** Counts for the page heading. Separate from the queue, so an empty queue still reports history. */
export function counts(rows: RequestRow[]): Record<RequestStatus, number> {
  const out: Record<RequestStatus, number> = {
    Pending: 0, Approved: 0, Rejected: 0, Failed: 0, Cancelled: 0, Revoked: 0,
  };
  for (const r of rows ?? []) if (out[r.status] !== undefined) out[r.status] += 1;
  return out;
}

/**
 * One sentence describing what approving this request will actually do.
 *
 * Shown in the confirmation, because "Approve" alone does not say whether a document is about to be
 * recycled or handed to someone outside the company.
 */
export function decisionSummary(row: RequestRow): string {
  if (!row) return "";
  if (row.type === "Deletion") {
    /* The NAME is shown on its own line by the dialog since 2026-08-30, so this sentence no longer
       repeats it. 93 days, not the 90 in the client's mock — confirmed with them: the recycle bin
       really is 93, and that number is what makes an approved deletion reversible. */
    return "This file will be moved to the recycle bin, where it can be restored for 93 days.";
  }
  const who = (row.shareWith ?? []).join(", ");
  const level = row.sharePermission === "Edit" ? "edit" : "view";
  const until = row.expiresAt ? ` until ${row.expiresAt.slice(0, 10)}` : ", with no expiry date";
  return `${who} will be able to ${level} ${row.itemName}${until}.`;
}

/* ── Live shares, and taking them back ──────────────────────────────────────────
   Spec: docs/superpowers/specs/2026-08-28-requests-page-redesign-and-share-revoke-design.md

   Completes §5.3 of the 2026-08-15 design, which specified a Revoke action and left it unbuilt:
   *"Without a way back out, the count only ever grows."* Every approved share creates a unique
   permission scope on its file, and that is the same 50,000-per-list ceiling that killed per-uploader
   ACLs on 2026-08-06.

   ⚠ THE GOVERNING RULE OF THIS SECTION: **approved Share ROWS are the INDEX; the file's live ACL is
   the VERDICT.** They disagree in BOTH directions, and neither direction is rare:

     * a share done through SharePoint's OWN Share button leaves no row at all — a web part cannot
       intercept that native control (settled in 2026-07-23-share-guard-retirement.md, not
       re-openable), so rows-only UNDER-reports;
     * a share revoked directly in SharePoint leaves its row still reading `Approved`, so rows-only
       OVER-reports.

   On a screen whose whole job is "who can currently reach this document", both are unacceptable.
   Same shape as `mergeRecords` in My Submissions: the record indexes, the live read decides. */

/**
 * Is a recipient's access still in place?
 *
 * ⚠ THREE STATES, NEVER TWO. `unknown` exists because a throttled or refused ACL read must never
 * render as "revoked" — that tells an approver access has ended when it has not, and they stop
 * looking. Empty ≠ unknown, in the place where the cost is believing a document is no longer shared.
 */
export type ShareState = "live" | "revoked" | "unknown";

/** One principal found on a file's ACL, as far as this module needs it. */
export interface AclPrincipal {
  /** SharePoint's principal id — what `removeroleassignment` needs. */
  id: number;
  /** `1` = user, `8` = SharePoint group, `4` = security group. Only users can be recipients. */
  principalType?: number;
  email?: string;
  loginName?: string;
  title?: string;
  /**
   * What levels this principal actually holds on the file. **Absent means the read did not ask** —
   * see `holdsRealAccess`, which then errs towards showing the grant.
   */
  bindings?: AclBinding[];
}

/** One permission level inside a role assignment. */
export interface AclBinding {
  id?: number;
  name?: string;
  /** `1` = Guest, which is what SharePoint calls Limited Access. `7` = System. */
  roleTypeKind?: number;
}

/** SharePoint's automatic traversal entry. Grants NOTHING on its own. */
const LIMITED_ACCESS_DEF_ID = 1073741825;
const LIMITED_ACCESS_ROLE_TYPE = 1;

/**
 * Does this principal actually hold anything on the file?
 *
 * ⚠ **A PRINCIPAL LISTED IN `RoleAssignments` HAS NOT NECESSARILY GOT ACCESS**, and reading mere
 * PRESENCE as access is what made a revoked recipient go on reading `has access` (2026-08-28).
 * SharePoint adds a **Limited Access** entry automatically wherever a principal was granted on a
 * child, and leaves one behind after some revocations — it is a traversal entry so a URL resolves,
 * and it lets nobody open, list or search anything. CLAUDE.md already records the same trap on the
 * Site Pages list, where *"almost every entry is Limited Access and grants NOTHING"*.
 *
 * ⚠ IT FAILS TOWARDS SHOWING THE GRANT, deliberately, and the asymmetry is the point. `bindings`
 * absent means the read did not ask (an older caller, or a response shape we did not anticipate) —
 * so we answer TRUE, exactly as an absent `principalType` is read as a user. Only bindings that were
 * READ and are ALL Limited Access answer false. On a screen about who can reach a document, an extra
 * name is a question somebody asks; a missing one is a disclosure nobody notices.
 */
export function holdsRealAccess(p: AclPrincipal): boolean {
  const bindings = p?.bindings;
  if (!bindings || bindings.length === 0) return true;
  return bindings.some(
    (b) =>
      b &&
      b.id !== LIMITED_ACCESS_DEF_ID &&
      b.roleTypeKind !== LIMITED_ACCESS_ROLE_TYPE,
  );
}

/** A file's live permission state. `undefined` for the whole thing means "not read", never "empty". */
export interface FileAcl {
  /** False ⇒ the file inherits its unit folder, so nothing is shared on it. Definitive. */
  hasUniqueRoleAssignments: boolean;
  principals: AclPrincipal[];
  /**
   * The document itself is not there — a 404 from `GetFileById`, i.e. deleted, or moved out of this
   * viewer's reach (SharePoint security-trims to 404, not 403).
   *
   * ⚠ SET ON 404 ALONE, NEVER ON ANY OTHER FAILURE. A throttle, a 403 or a malformed response all
   * stay `undefined` and render as "could not be checked", because `fileGone` REMOVES the row from
   * the Shared-files tab — and silently shortening the list of who can reach documents, on the
   * strength of a read that merely failed, is the one error this tab must not make.
   */
  fileGone?: boolean;
}

export interface ShareRecipient {
  email: string;
  permission: SharePermission;
  state: ShareState;
  /** From the live ACL. Absent ⇒ cannot be revoked individually — `undefined` is not `0`. */
  principalId?: number;
  /** Who asked for this share, and who let it through. The "who did what" the client asked for. */
  requestedBy?: string;
  decidedBy?: string;
  decidedAt?: string;
  expiresAt?: string;
  /**
   * On the file's ACL, but named by no request row.
   *
   * Almost always a share made through SharePoint's own Share button. Surfaced rather than hidden:
   * the question this tab answers is "who can reach this document", and a grant we did not make is
   * still a grant. Rendered as a note, not an alarm — a site administrator can legitimately appear.
   */
  unrecorded?: boolean;
}

export interface SharedFile {
  itemUniqueId: string;
  itemName: string;
  itemUrl?: string;
  segment: string;
  unit: string;
  unitTermGuid?: string;
  /** Always `"approved"` — `canShareStage` refuses to share a pending file. Carried for `inScope`. */
  stage?: RequestStage;
  recipients: ShareRecipient[];
  /** Worst-case rollup: any `unknown` recipient makes the file `unknown`, never `live`. */
  state: ShareState;
  /** The request rows behind this file, oldest first. Needed to write the revoke back. */
  rowIds: number[];
}

const lower = (s: string | undefined): string => (s ?? "").trim().toLowerCase();

/**
 * Which files does this viewer have approved shares on?
 *
 * Indexes on `itemUniqueId` — the one identifier that survives a rename or a move, which a URL does
 * not. Several requests can name one file (two people asked at different times, or a share was
 * extended), so recipients are MERGED across rows and the newest row wins on conflicting detail.
 *
 * ⚠ `Revoked` rows are included, deliberately. A fully revoked share still belongs on this screen —
 * it is the record that access was granted and taken back, and dropping it would make a revoke look
 * like the request never happened. The live probe marks its recipients `revoked` anyway.
 *
 * Scope is the SAME `inScope` the queue uses, so a file this viewer cannot act on is one they do not
 * see.
 */
export function collectSharedFiles(
  rows: RequestRow[],
  scope: string[] | ViewerScope,
): SharedFile[] {
  const byFile: { [key: string]: SharedFile } = {};
  const order: string[] = [];

  const relevant = (rows ?? [])
    .filter((r) => r && r.type === "Share")
    .filter((r) => r.status === "Approved" || r.status === "Revoked")
    .filter((r) => lower(r.itemUniqueId).length > 0)
    .filter((r) => inScope(r, scope))
    // Oldest first, so a plain overwrite in the merge below leaves the NEWEST row's detail standing.
    .slice()
    .sort((a, b) =>
      (a.decidedAt || a.requestedAt || "").localeCompare(b.decidedAt || b.requestedAt || ""),
    );

  for (const r of relevant) {
    const key = lower(r.itemUniqueId);
    if (!byFile[key]) {
      byFile[key] = {
        itemUniqueId: r.itemUniqueId,
        itemName: r.itemName,
        itemUrl: r.itemUrl,
        segment: r.segment,
        unit: r.unit,
        unitTermGuid: r.unitTermGuid,
        stage: "approved",
        recipients: [],
        // Nothing has been probed yet, and "not probed" is exactly what `unknown` means.
        state: "unknown",
        rowIds: [],
      };
      order.push(key);
    }
    const f = byFile[key];
    if (typeof r.id === "number") f.rowIds.push(r.id);
    // A later row may carry the name the file was renamed to; prefer it over the first sighting.
    if ((r.itemName || "").length > 0) f.itemName = r.itemName;
    if ((r.itemUrl || "").length > 0) f.itemUrl = r.itemUrl;

    for (const raw of r.shareWith || []) {
      const email = lower(raw);
      if (email.length === 0) continue;
      const existing = f.recipients.filter((x) => x.email === email)[0];
      const permission: SharePermission = r.sharePermission === "Edit" ? "Edit" : "View";
      if (existing) {
        existing.permission = permission;
        existing.requestedBy = r.requestedBy;
        existing.decidedBy = r.decidedBy;
        existing.decidedAt = r.decidedAt;
        existing.expiresAt = r.expiresAt;
      } else {
        f.recipients.push({
          email,
          permission,
          state: "unknown",
          requestedBy: r.requestedBy,
          decidedBy: r.decidedBy,
          decidedAt: r.decidedAt,
          expiresAt: r.expiresAt,
        });
      }
    }
  }

  return order.map((k) => byFile[k]);
}

/**
 * Apply a file's live permissions to what the rows claimed.
 *
 * `acl === undefined` means the read failed or was never made — every recipient stays `unknown`, and
 * so does the file. It must NOT mean "nothing is shared": that is the one wrong answer this whole
 * three-state design exists to prevent.
 *
 * ⚠ GROUP PRINCIPALS ARE SKIPPED. Breaking inheritance COPIES the unit folder's groups onto the
 * file, so every shared file carries its unit's own groups on its ACL. Counting them would report a
 * unit's ordinary access as a share — on the one screen that answers "who can reach this document".
 */
export function mergeShareAcl(file: SharedFile, acl: FileAcl | undefined): SharedFile {
  if (!acl) {
    return {
      ...file,
      recipients: file.recipients.map((r) => ({ ...r, state: "unknown" as ShareState })),
      state: "unknown",
    };
  }

  /* Inheriting is DEFINITIVE, and it is the state "revoke all" leaves behind: no unique scope means
     no per-file grant, so every recipient is gone. The one branch that can answer `revoked` without
     inspecting a single principal. */
  if (acl.hasUniqueRoleAssignments === false) {
    return {
      ...file,
      recipients: file.recipients.map((r) => ({
        ...r,
        state: "revoked" as ShareState,
        principalId: undefined,
      })),
      state: "revoked",
    };
  }

  /* Only USER principals can be a recipient. `principalType` ABSENT is treated as a user, so a
     response shape we did not anticipate errs towards SHOWING a grant rather than hiding one — the
     safe direction on a screen about who can reach a document. */
  const users = (acl.principals || []).filter(
    (p) => p && (p.principalType === undefined || p.principalType === 1) && holdsRealAccess(p),
  );
  const byEmail: { [key: string]: AclPrincipal } = {};
  for (const p of users) {
    const keys = [lower(p.email), emailFromLogin(p.loginName)];
    for (const k of keys) if (k.length > 0 && !byEmail[k]) byEmail[k] = p;
  }

  const recipients: ShareRecipient[] = file.recipients.map((r) => {
    const hit = byEmail[r.email];
    return hit
      ? { ...r, state: "live" as ShareState, principalId: hit.id }
      : { ...r, state: "revoked" as ShareState, principalId: undefined };
  });

  // Anyone holding the file whom no row accounts for — see `ShareRecipient.unrecorded`.
  const known: { [key: string]: true } = {};
  for (const r of recipients) known[r.email] = true;
  for (const p of users) {
    const email = lower(p.email) || emailFromLogin(p.loginName);
    if (email.length === 0 || known[email]) continue;
    known[email] = true;
    recipients.push({
      email,
      permission: "View",
      state: "live",
      principalId: p.id,
      unrecorded: true,
    });
  }

  return { ...file, recipients, state: rollUp(recipients) };
}

/** Worst-case: one unreadable recipient makes the whole file unreadable, never "live". */
function rollUp(recipients: ShareRecipient[]): ShareState {
  for (const r of recipients) if (r.state === "unknown") return "unknown";
  for (const r of recipients) if (r.state === "live") return "live";
  return "revoked";
}

/**
 * Pull an address out of a SharePoint claim.
 *
 * A member reads `i:0#.f|membership|someone@example.com`; a GUEST reads
 * `i:0#.f|membership|someone_example.com#ext#@tenant.onmicrosoft.com`, where the real address has had
 * its `@` turned into `_`. Both are matched, because a guest is the COMMON case for a share and
 * matching members alone would report every external recipient as revoked.
 */
function emailFromLogin(login: string | undefined): string {
  const raw = lower(login);
  if (raw.length === 0) return "";
  const tail = raw.indexOf("|") === -1 ? raw : raw.slice(raw.lastIndexOf("|") + 1);
  const ext = tail.indexOf("#ext#");
  if (ext === -1) return tail.indexOf("@") === -1 ? "" : tail;
  const guest = tail.slice(0, ext);
  const under = guest.lastIndexOf("_");
  return under === -1 ? "" : guest.slice(0, under) + "@" + guest.slice(under + 1);
}

/**
 * May this person take a share back?
 *
 * The SAME `inScope` that decides the request, reused rather than reimplemented — so the revoke
 * button and the queue can never disagree about which units a Head of Department may act on.
 *
 * All three deciding personas already hold the right: `hou`/`hod` carry `SHARE` and `hou_hc`/`hod`
 * carry `SHAREHC`, and **`CRS Share` contains Manage Permissions** — which is the reason it exists as
 * a separate level at all (`FolderManager.tsx` ~341). So this needs no new role, no new permission
 * level and no reconciliation run.
 *
 * ⚠ NOT the original requester. A PIC holds no Manage Permissions, so offering them the button would
 * fail in their own session AFTER the screen had offered it — the "visible but inert" state this
 * codebase already rejected for pending-stage requests shown to a Head of Department.
 */
export function canRevoke(file: SharedFile | undefined, scope: string[] | ViewerScope): boolean {
  if (!file) return false;
  return inScope({ unit: file.unit, unitTermGuid: file.unitTermGuid, stage: "approved" }, scope);
}

/**
 * Does revoking these recipients end the share entirely?
 *
 * Decides whether the row's `Status` becomes `Revoked` (nothing left) or stays `Approved` with a note
 * (part of the grant withdrawn).
 *
 * ⚠ `unknown` RECIPIENTS COUNT AS REMAINING. A row must not be marked fully revoked on the strength
 * of a read that did not answer — the same reason `mergeShareAcl` has three states at all.
 */
/**
 * What the **Shared files** tab shows: who can reach these documents *now*.
 *
 * Client, 2026-08-28, having revoked somebody and watched them stay on the list: *"it should remove
 * crystal from the label since she have no access anymore."* Right — the tab answers a
 * present-tense question, and a person who cannot reach the file is not part of that answer.
 *
 * ⚠ THIS REVERSES THE "REVOKED RECIPIENTS ARE KEPT" LINE IN THE ORIGINAL DESIGN, and only that line.
 * The record is not lost: the revoke appends to the request's own `DecisionNote`, the row turns
 * `Revoked` when the last recipient goes, and the **Share** tab's Decided section still lists it.
 * History belongs there; this tab is the live answer.
 *
 * ⚠ `unknown` IS KEPT, AND THAT IS THE WHOLE CARE IN THIS FUNCTION. A file whose ACL could not be
 * read must stay visible and stay amber — hiding it would let one throttled request quietly shorten
 * the list of who can reach the company's documents, which is the failure this screen exists to
 * prevent. Only `revoked`, which is a definite answer, is dropped.
 */
/**
 * Files whose DOCUMENT no longer exists, so there is nothing left to reach.
 *
 * Split out rather than folded into `currentlyShared` because the caller has to SAY how many were
 * dropped: a row vanishing with no explanation is how a tab about access quietly stops being
 * trusted.
 */
export function goneFiles(files: SharedFile[], acls: Record<string, FileAcl | undefined>): number {
  let n = 0;
  for (const f of files ?? []) if (f && acls[f.itemUniqueId]?.fileGone === true) n += 1;
  return n;
}

/** Drops files whose document is gone. Anything else — including an unreadable ACL — is kept. */
export function withoutGoneFiles(
  files: SharedFile[],
  acls: Record<string, FileAcl | undefined>,
): SharedFile[] {
  return (files ?? []).filter((f) => f && acls[f.itemUniqueId]?.fileGone !== true);
}

export function currentlyShared(files: SharedFile[]): SharedFile[] {
  const out: SharedFile[] = [];
  for (const f of files ?? []) {
    if (!f) continue;
    const recipients = (f.recipients ?? []).filter((r) => r && r.state !== "revoked");
    // Nobody left who can reach it — the file itself has stopped being shared.
    if (recipients.length === 0) continue;
    out.push({ ...f, recipients, state: rollUp(recipients) });
  }
  return out;
}

export function endsTheShare(file: SharedFile, revoking: string[]): boolean {
  const going: { [key: string]: true } = {};
  for (const e of revoking || []) going[lower(e)] = true;
  for (const r of file.recipients) {
    if (r.state !== "revoked" && !going[r.email]) return false;
  }
  return true;
}

/**
 * Today's date as the person reading it would write it — `YYYY-MM-DD` from LOCAL components.
 *
 * ⚠ NOT `toISOString().slice(0, 10)`, which is UTC. Verified live 2026-08-30: a revoke at 02:43
 * local (UTC+8) wrote *"on 2026-08-29"* into the record while the audit row beside it was stamped
 * 30/Aug 02:43 — one line of one record disagreeing with the next about which day it happened.
 * Anything done before 08:00 local was a day behind. Same family as Auto-route's `Created` stamp.
 *
 * The revoker's own local day is the right answer here: this sentence is a person saying when they
 * did something, not a machine-readable timestamp. `EventTime` on the audit row remains ISO/UTC,
 * which is what a `$filter` needs — the two serve different readers and are deliberately different.
 */
export function localDateStamp(when: Date): string {
  const pad = (n: number): string => (n < 10 ? "0" + n : String(n));
  return when.getFullYear() + "-" + pad(when.getMonth() + 1) + "-" + pad(when.getDate());
}

/**
 * The line appended to `DecisionNote` when access is taken back. Past tense — it has happened.
 *
 * `whenDate` is a plain `YYYY-MM-DD` (see `localDateStamp`); a full ISO string still works, since
 * only the first ten characters are read — but passing one puts the UTC day back into the record.
 */
export function revocationNote(emails: string[], actor: string, whenDate: string): string {
  const who = (emails || []).filter((e) => (e || "").length > 0).join(", ");
  return (
    "Access for " + (who || "the recipients") +
    " revoked by " + (actor || "an approver") +
    " on " + (whenDate || "").slice(0, 10) + "."
  );
}

/**
 * A stored date as a person reads it: `"2026-09-02"` → `"02 September 2026"`.
 *
 * Client, 2026-09-04: *"CHange the date format to 02 September 2026."*
 *
 * ⚠ PARSED FROM THE STRING, NEVER THROUGH `new Date()`. These columns are written and filtered ISO
 * (the 2026-08-13 rule for this list), so the value carries a UTC instant — and at UTC+8 anything
 * stamped before 08:00 local belongs to the PREVIOUS UTC day. Handing it to `new Date()` and reading
 * `getDate()` would silently shift roughly a third of all requests back by one day, which is the same
 * fault already recorded against Auto-route's `Created` stamp and against `revocationNote`. Reading
 * the first ten characters keeps the day exactly as it was stored.
 *
 * ⚠ ANYTHING UNPARSEABLE COMES BACK UNCHANGED, never `"Invalid Date"` and never blank. A date this
 * function cannot read is still information — the raw value on screen is recognisable and reportable,
 * where `Invalid Date` tells the reader nothing and an empty string hides that a date exists at all.
 */
export function longDate(iso: string | undefined): string {
  const raw = (iso ?? "").trim();
  if (raw.length === 0) return "";
  const day = raw.slice(0, 10);
  const parts = day.split("-");
  if (parts.length !== 3) return raw;
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);
  if (!isFinite(y) || !isFinite(m) || !isFinite(d)) return raw;
  if (m < 1 || m > 12 || d < 1 || d > 31) return raw;
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];
  // Two digits, matching the client's own example ("02 September 2026").
  return (d < 10 ? "0" + d : String(d)) + " " + months[m - 1] + " " + y;
}
