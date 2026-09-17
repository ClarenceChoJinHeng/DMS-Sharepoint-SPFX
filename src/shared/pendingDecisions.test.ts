import {
  PENDING_DECISION_COLUMNS,
  DECISION_JOIN_COLUMN,
  buildDecisionPayload,
  parseDecisionRow,
  PendingDecision,
} from "./pendingDecisions";

describe("PENDING_DECISION_COLUMNS", () => {
  it("names every column CRS Pending Decisions needs", () => {
    const names = PENDING_DECISION_COLUMNS.map((c) => c.name);
    expect(names).toEqual([
      "ItemUniqueId",
      "Decision",
      "DecidedBy",
      "DecidedAt",
      "RejectionComment",
      "IsSelfApprove",
      "Status",
      "StatusError",
    ]);
  });

  it("Decision and Status are TEXT, never Choice", () => {
    const decision = PENDING_DECISION_COLUMNS.find((c) => c.name === "Decision");
    const status = PENDING_DECISION_COLUMNS.find((c) => c.name === "Status");
    expect(decision?.type).toBe(2);
    expect(status?.type).toBe(2);
  });

  it("IsSelfApprove is type 8 (Boolean)", () => {
    const col = PENDING_DECISION_COLUMNS.find((c) => c.name === "IsSelfApprove");
    expect(col?.type).toBe(8);
  });
});

describe("buildDecisionPayload", () => {
  const base: Omit<PendingDecision, "id" | "status" | "statusError"> = {
    itemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
    decision: "Approved",
    decidedBy: "approver@example.com",
    decidedAt: new Date("2026-09-18T02:30:00Z"),
    isSelfApprove: false,
  };

  it("writes the decision and actor, omits RejectionComment for an approval", () => {
    const payload = buildDecisionPayload(base);
    expect(payload.Decision).toBe("Approved");
    expect(payload.DecidedBy).toBe("approver@example.com");
    expect(payload.ItemUniqueId).toBe("f7d2d18c-1234-4a4a-9a9a-abcdefabcdef");
    expect(payload.RejectionComment).toBeUndefined();
  });

  it("includes RejectionComment for a rejection", () => {
    const payload = buildDecisionPayload({
      ...base,
      decision: "Rejected",
      rejectionComment: "Wrong vendor named on the document",
    });
    expect(payload.Decision).toBe("Rejected");
    expect(payload.RejectionComment).toBe("Wrong vendor named on the document");
  });

  it("marks IsSelfApprove true for a self-approve", () => {
    const payload = buildDecisionPayload({ ...base, isSelfApprove: true });
    expect(payload.IsSelfApprove).toBe(true);
  });

  it("never sets Status on the initial write — the row starts Pending by omission, the flow is what marks it Applied/Failed", () => {
    const payload = buildDecisionPayload(base);
    expect(payload.Status).toBeUndefined();
  });
});

describe("parseDecisionRow", () => {
  it("reads an applied decision back", () => {
    const row = parseDecisionRow({
      Id: 42,
      ItemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
      Decision: "Approved",
      DecidedBy: "approver@example.com",
      DecidedAt: "2026-09-18T02:30:00Z",
      IsSelfApprove: false,
      Status: "Applied",
    });
    expect(row.id).toBe(42);
    expect(row.decision).toBe("Approved");
    expect(row.status).toBe("Applied");
    expect(row.isSelfApprove).toBe(false);
  });

  it("leaves status undefined, never a guessed value, on a row with none", () => {
    const row = parseDecisionRow({
      Id: 42,
      ItemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
      Decision: "Approved",
      DecidedBy: "approver@example.com",
      DecidedAt: "2026-09-18T02:30:00Z",
      IsSelfApprove: false,
    });
    expect(row.status).toBeUndefined();
  });

  it("defaults isSelfApprove to false, never undefined, when the column is missing from an older row", () => {
    const row = parseDecisionRow({
      Id: 42,
      ItemUniqueId: "f7d2d18c-1234-4a4a-9a9a-abcdefabcdef",
      Decision: "Approved",
      DecidedBy: "approver@example.com",
      DecidedAt: "2026-09-18T02:30:00Z",
    });
    expect(row.isSelfApprove).toBe(false);
  });
});

describe("DECISION_JOIN_COLUMN", () => {
  it("is ItemUniqueId — the flow addresses the document by this, never by name or path", () => {
    expect(DECISION_JOIN_COLUMN).toBe("ItemUniqueId");
  });
});
