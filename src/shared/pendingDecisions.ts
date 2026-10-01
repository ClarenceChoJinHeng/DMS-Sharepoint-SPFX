/**
 * Pending approve/reject decisions — the write, deferred.
 *
 * Spec: docs/superpowers/specs/2026-09-18-tag-approve-proxy-design.md
 *
 * Client: PIC and Approver lose direct Edit/Approve rights on the document libraries, so an
 * approver's own click can no longer flip `OData__ModerationStatus` directly — that write moves to
 * `CRS — Apply pending decisions` (Power Automate, running as crs@sdguthrie.com). This module is
 * the shape of the row that drives it: everything the flow needs to replay the decision, nothing
 * more.
 *
 * A NEW, purpose-built list, not an extension of `CRS Requests`. That list's whole design —
 * `RequestType: "Deletion" | "Share"`, a requester and a SEPARATE decider, the `ViewerScope`
 * visibility model — is built around *someone else decides on your behalf*. An approve/reject
 * decision (including a self-approve, where the same person both uploads and decides) is a
 * different shape: one action, not a request-then-decision pair. Bolting a third `RequestType`
 * onto `CRS Requests`' already heavily-worked machinery risks breaking what currently works, for
 * no benefit over a small new list.
 *
 * Pure and SPFx-free, same split as `submissionRecords.ts` / `spSubmissionRecords.ts`.
 */

export const PENDING_DECISION_COLUMNS: Array<{ name: string; type: number }> = [
  /** The document being decided on. The ONLY identifier — resolved by `GetFileById(guid'...')`,
   * web-scoped, works across every library regardless of normal/HC. Written fresh, moments before
   * this row is written, in the same browser session as the check that produced it — unlike a
   * `CRS Requests` row, there is no multi-day gap here, so no stamp-based fallback is needed. */
  { name: "ItemUniqueId", type: 2 },
  /** TEXT, never Choice — a value absent from a Choice column's `Choices` fails the whole write. */
  { name: "Decision", type: 2 },
  { name: "DecidedBy", type: 2 },
  { name: "DecidedAt", type: 4 },
  /** Only present for a Rejected decision. */
  { name: "RejectionComment", type: 3 },
  /** True when this came from `autoApproveOwnUpload`, never a human clicking Approve on someone
   * else's file. Exists so the applying flow — and any future reporting — can tell the two apart
   * without guessing from `DecidedBy` matching the file's own author. */
  { name: "IsSelfApprove", type: 8 },
  /** blank/absent ⇒ not yet processed. "Applied" ⇒ the flow performed the write. "Failed" ⇒ see
   * StatusError. Same three-state shape as `TagStatus` on `CRS Submissions`. */
  { name: "Status", type: 2 },
  { name: "StatusError", type: 3 },
];

/** The column the flow addresses the document by. Never name, never path — both can change. */
export const DECISION_JOIN_COLUMN = "ItemUniqueId";

export type DecisionOutcome = "Approved" | "Rejected";

export interface PendingDecision {
  id: number;
  itemUniqueId: string;
  decision: DecisionOutcome;
  decidedBy: string;
  decidedAt: Date;
  rejectionComment?: string;
  isSelfApprove: boolean;
  status?: string;
  statusError?: string;
}

/** Builds the write body for a NEW decision row. Never sets `Status` — the row starts Pending by
 * omission; only the flow that applies the decision ever writes `Applied`/`Failed`. */
export function buildDecisionPayload(
  input: Omit<PendingDecision, "id" | "status" | "statusError">,
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    ItemUniqueId: input.itemUniqueId,
    Decision: input.decision,
    DecidedBy: input.decidedBy,
    DecidedAt: input.decidedAt.toISOString(),
    IsSelfApprove: input.isSelfApprove,
  };
  if (input.rejectionComment !== undefined) {
    payload.RejectionComment = input.rejectionComment;
  }
  return payload;
}

/** Reads a raw list-item row back into a `PendingDecision`. `isSelfApprove` defaults to `false`
 * rather than `undefined` when the column is missing (an older row, or a read that omitted it) —
 * an ordinary human approval must never be misread as a self-approve for want of a value. */
export function parseDecisionRow(row: Record<string, unknown>): PendingDecision {
  const str = (k: string): string | undefined => {
    const v = row[k];
    return typeof v === "string" && v.length > 0 ? v : undefined;
  };
  return {
    id: Number(row.Id ?? 0),
    itemUniqueId: str("ItemUniqueId") ?? "",
    decision: (str("Decision") as DecisionOutcome) ?? "Approved",
    decidedBy: str("DecidedBy") ?? "",
    decidedAt: new Date(String(row.DecidedAt ?? "")),
    rejectionComment: str("RejectionComment"),
    isSelfApprove: row.IsSelfApprove === true,
    status: str("Status"),
    statusError: str("StatusError"),
  };
}
