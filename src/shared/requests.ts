/**
 * Deletion and share requests — the rules, with no React and no network.
 *
 * Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
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
export type RequestStatus = "Pending" | "Approved" | "Rejected" | "Failed";

export type SharePermission = "View" | "Edit";

export interface RequestRow {
  id?: number;
  type: RequestType;
  status: RequestStatus;
  /** The file's UniqueId — survives rename and move, which a URL does not. */
  itemUniqueId: string;
  itemName: string;
  itemUrl?: string;
  segment: string;
  unit: string;
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
}

/** What the raise-a-request form collects, before it becomes a row. */
export interface RequestDraft {
  type: RequestType;
  itemUniqueId: string;
  itemName: string;
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

  if (d.type === "Share") {
    const list = parseRecipients(d.shareWith ?? "");
    if (list.length === 0) {
      out.push("Add at least one person to share with.");
    } else {
      const bad = list.filter((e) => !looksLikeEmail(e));
      if (bad.length > 0) out.push(`Not an email address: ${bad.join(", ")}.`);

      const external = list.filter((e) => looksLikeEmail(e) && isExternal(e, ctx?.tenantDomains ?? []));
      if (external.length > 0 && !(ctx && ctx.allowExternal)) {
        // Names the setting, because this is a policy refusal rather than a mistake by the requester.
        out.push(
          `Sharing outside the organisation is switched off, so ${external.join(", ")} cannot be added. ` +
            "An administrator changes this on the DMS Config list.",
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
 * May this person decide this request?
 *
 * Two conditions, and the second is the one easily forgotten: the request must still be PENDING.
 * Without it, two approvers on the same queue both press Approve and the action runs twice — a file
 * deleted twice is harmless, but a share granted twice leaves a second permission scope nobody tracks.
 *
 * `approverUnits` is the set of units this viewer approves for. An empty set decides nothing, which is
 * what a PIC gets.
 */
export function canDecide(row: RequestRow, approverUnits: string[]): boolean {
  if (!row || row.status !== "Pending") return false;
  const units = (approverUnits ?? [])
    .map((u) => (u ?? "").trim().toLowerCase())
    .filter((u) => u.length > 0);
  return units.indexOf((row.unit ?? "").trim().toLowerCase()) !== -1;
}

/**
 * Can this person see the request at all?
 *
 * Their own, or one for a unit they approve. NOT a confidentiality boundary — every PIC already reads
 * every approved document in their unit, so a request row names no file they could not already see.
 * It keeps the queue about the reader rather than showing them everyone else's admin.
 */
export function isVisibleTo(row: RequestRow, viewerEmail: string, approverUnits: string[]): boolean {
  const me = (viewerEmail ?? "").trim().toLowerCase();
  if (me.length > 0 && (row?.requestedBy ?? "").trim().toLowerCase() === me) return true;
  const units = (approverUnits ?? []).map((u) => (u ?? "").trim().toLowerCase());
  return units.indexOf((row?.unit ?? "").trim().toLowerCase()) !== -1;
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
export function queueFor(rows: RequestRow[], approverUnits: string[]): RequestRow[] {
  return (rows ?? [])
    .filter((r) => canDecide(r, approverUnits))
    .sort((a, b) => (a.requestedAt ?? "").localeCompare(b.requestedAt ?? ""));
}

/** Counts for the page heading. Separate from the queue, so an empty queue still reports history. */
export function counts(rows: RequestRow[]): Record<RequestStatus, number> {
  const out: Record<RequestStatus, number> = { Pending: 0, Approved: 0, Rejected: 0, Failed: 0 };
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
    return `${row.itemName} will be moved to the recycle bin, where it can be restored for 93 days.`;
  }
  const who = (row.shareWith ?? []).join(", ");
  const level = row.sharePermission === "Edit" ? "edit" : "view";
  const until = row.expiresAt ? ` until ${row.expiresAt.slice(0, 10)}` : ", with no expiry date";
  return `${who} will be able to ${level} ${row.itemName}${until}.`;
}
