/**
 * Tests for deletion and share request rules.
 *
 * Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
 *
 * Weighted toward the cases where being wrong is silent or dangerous: a document leaving the
 * organisation because external sharing could not be confirmed, a share granted twice because two
 * approvers pressed the button, and a failed action recorded as a success.
 */

import {
  Decision,
  RequestDraft,
  RequestRow,
  applyDecision,
  canDecide,
  counts,
  decisionSummary,
  isExternal,
  isVisibleTo,
  looksLikeEmail,
  parseRecipients,
  queueFor,
  validateDraft,
} from "./requests";

const DOMAINS = ["example.com", "sdguthrie.com"];

const draft = (over: Partial<RequestDraft> = {}): RequestDraft => ({
  type: "Deletion",
  itemUniqueId: "a1b2c3d4-0000-0000-0000-000000000000",
  itemName: "Q3 Report.pdf",
  segment: "Group Head Office",
  unit: "Corporate",
  reason: "Superseded by the final version.",
  ...over,
});

const row = (over: Partial<RequestRow> = {}): RequestRow => ({
  type: "Deletion",
  status: "Pending",
  itemUniqueId: "a1b2c3d4-0000-0000-0000-000000000000",
  itemName: "Q3 Report.pdf",
  segment: "Group Head Office",
  unit: "Corporate",
  requestedBy: "pic@example.com",
  requestedAt: "2026-08-15T09:30:00.000Z",
  reason: "Superseded.",
  ...over,
});

describe("parseRecipients", () => {
  it("splits on commas, semicolons and whitespace", () => {
    expect(parseRecipients("a@x.com, b@x.com; c@x.com d@x.com")).toEqual([
      "a@x.com", "b@x.com", "c@x.com", "d@x.com",
    ]);
  });

  it("lower-cases and de-duplicates, so nobody is granted twice", () => {
    expect(parseRecipients("A@X.com, a@x.com")).toEqual(["a@x.com"]);
  });

  it("is empty for blank input", () => {
    expect(parseRecipients("")).toEqual([]);
    expect(parseRecipients("   ")).toEqual([]);
  });
});

describe("looksLikeEmail", () => {
  it("accepts ordinary and awkward-but-real addresses", () => {
    expect(looksLikeEmail("clarence@trinergydigital.com")).toBe(true);
    expect(looksLikeEmail("first.last+tag@sub.example.co.uk")).toBe(true);
    expect(looksLikeEmail("o'brien@example.com")).toBe(true);
  });

  it("rejects what is plainly not an address", () => {
    expect(looksLikeEmail("clarence")).toBe(false);
    expect(looksLikeEmail("clarence@localhost")).toBe(false);
    expect(looksLikeEmail("a b@example.com")).toBe(false);
    expect(looksLikeEmail("")).toBe(false);
  });
});

describe("isExternal", () => {
  it("knows the tenant's own domains", () => {
    expect(isExternal("someone@example.com", DOMAINS)).toBe(false);
    expect(isExternal("SOMEONE@Example.com", DOMAINS)).toBe(false);
  });

  it("flags anything else", () => {
    expect(isExternal("someone@vendor.com", DOMAINS)).toBe(true);
  });

  it("TREATS UNKNOWN AS EXTERNAL — the one fail-closed rule here", () => {
    // No domains means we could not confirm what internal is. Guessing "internal" would send a
    // document out of the organisation on the strength of a failed read.
    expect(isExternal("someone@example.com", [])).toBe(true);
    expect(isExternal("someone@example.com", ["  "])).toBe(true);
  });

  it("treats a malformed address as external rather than internal", () => {
    expect(isExternal("no-at-sign", DOMAINS)).toBe(true);
  });
});

describe("validateDraft", () => {
  const ctx = { allowExternal: true, tenantDomains: DOMAINS, today: new Date(2026, 7, 15) };

  it("accepts a well-formed deletion request", () => {
    expect(validateDraft(draft(), ctx)).toEqual([]);
  });

  it("REQUIRES a reason — an approver with nothing to read approves everything", () => {
    expect(validateDraft(draft({ reason: "" }), ctx)[0]).toContain("reason");
    expect(validateDraft(draft({ reason: "   " }), ctx).length).toBe(1);
  });

  it("requires a document", () => {
    expect(validateDraft(draft({ itemUniqueId: "" }), ctx)[0]).toContain("No document");
  });

  it("requires at least one recipient on a share", () => {
    expect(validateDraft(draft({ type: "Share", shareWith: "" }), ctx)[0]).toContain("at least one");
  });

  it("names the addresses it could not read", () => {
    const errs = validateDraft(draft({ type: "Share", shareWith: "good@example.com, nonsense" }), ctx);
    expect(errs[0]).toContain("nonsense");
    expect(errs[0]).not.toContain("good@example.com");
  });

  it("blocks external recipients when the setting is off, and NAMES the setting", () => {
    const off = { allowExternal: false, tenantDomains: DOMAINS, today: new Date(2026, 7, 15) };
    const errs = validateDraft(draft({ type: "Share", shareWith: "john@vendor.com" }), off);
    expect(errs[0]).toContain("john@vendor.com");
    expect(errs[0]).toContain("DMS Config");
  });

  it("allows internal recipients even when external is off", () => {
    const off = { allowExternal: false, tenantDomains: DOMAINS, today: new Date(2026, 7, 15) };
    expect(validateDraft(draft({ type: "Share", shareWith: "colleague@example.com" }), off)).toEqual([]);
  });

  it("blocks EVERY recipient when the tenant domains are unknown and external is off", () => {
    // The fail-closed path end to end: an unreadable domain list must not quietly permit sharing.
    const off = { allowExternal: false, tenantDomains: [], today: new Date(2026, 7, 15) };
    expect(validateDraft(draft({ type: "Share", shareWith: "colleague@example.com" }), off).length).toBe(1);
  });

  it("rejects an expiry already in the past, without parsing a date string", () => {
    const errs = validateDraft(
      draft({ type: "Share", shareWith: "a@example.com", expiresAt: "2026-08-14" }),
      ctx,
    );
    expect(errs[0]).toContain("already passed");
  });

  it("accepts today as an expiry", () => {
    expect(
      validateDraft(draft({ type: "Share", shareWith: "a@example.com", expiresAt: "2026-08-15" }), ctx),
    ).toEqual([]);
  });

  it("accepts no expiry at all", () => {
    expect(validateDraft(draft({ type: "Share", shareWith: "a@example.com" }), ctx)).toEqual([]);
  });

  it("does not apply share rules to a deletion", () => {
    expect(validateDraft(draft({ type: "Deletion", shareWith: "" }), ctx)).toEqual([]);
  });
});

describe("canDecide", () => {
  it("lets an approver for that unit decide a pending request", () => {
    expect(canDecide(row(), ["Corporate"])).toBe(true);
    expect(canDecide(row(), ["corporate"])).toBe(true);
  });

  it("refuses an approver for a different unit", () => {
    expect(canDecide(row(), ["Treasury"])).toBe(false);
  });

  it("refuses someone who approves nothing — which is what a PIC has", () => {
    expect(canDecide(row(), [])).toBe(false);
  });

  it("REFUSES A REQUEST THAT IS NO LONGER PENDING", () => {
    // Two approvers on one queue both press Approve. Without this the action runs twice — a second
    // deletion is harmless, a second share leaves a permission scope nobody is tracking.
    expect(canDecide(row({ status: "Approved" }), ["Corporate"])).toBe(false);
    expect(canDecide(row({ status: "Rejected" }), ["Corporate"])).toBe(false);
    expect(canDecide(row({ status: "Failed" }), ["Corporate"])).toBe(false);
  });
});

describe("isVisibleTo", () => {
  it("shows people their own requests", () => {
    expect(isVisibleTo(row(), "pic@example.com", [])).toBe(true);
    expect(isVisibleTo(row(), "PIC@Example.com", [])).toBe(true);
  });

  it("shows approvers their unit's requests", () => {
    expect(isVisibleTo(row(), "hou@example.com", ["Corporate"])).toBe(true);
  });

  it("hides everyone else's", () => {
    expect(isVisibleTo(row(), "other@example.com", ["Treasury"])).toBe(false);
  });

  it("still shows a decided request to its requester", () => {
    expect(isVisibleTo(row({ status: "Rejected" }), "pic@example.com", [])).toBe(true);
  });
});

describe("applyDecision", () => {
  const at = "2026-08-15T11:00:00.000Z";

  it("records an approval", () => {
    const out = applyDecision(row(), { approve: true, by: "hou@example.com", at });
    expect(out.status).toBe("Approved");
    expect(out.decidedBy).toBe("hou@example.com");
    expect(out.decidedAt).toBe(at);
  });

  it("records a rejection with its note", () => {
    const d: Decision = { approve: false, by: "hou@example.com", at, note: "Still needed for audit." };
    const out = applyDecision(row(), d);
    expect(out.status).toBe("Rejected");
    expect(out.decisionNote).toBe("Still needed for audit.");
  });

  it("RECORDS A FAILED ACTION AS FAILED, NEVER APPROVED", () => {
    // "Approved" against a file that could not be deleted tells the next person the job is done.
    const out = applyDecision(row(), {
      approve: true, by: "hou@example.com", at, failure: "The file no longer exists.",
    });
    expect(out.status).toBe("Failed");
    expect(out.decisionNote).toBe("The file no longer exists.");
  });

  it("lets the failure reason win over any note", () => {
    const out = applyDecision(row(), {
      approve: true, by: "x@example.com", at, note: "Looks fine", failure: "HTTP 403",
    });
    expect(out.decisionNote).toBe("HTTP 403");
  });

  it("a rejection is never Failed, even if something else went wrong", () => {
    const out = applyDecision(row(), { approve: false, by: "x@example.com", at, failure: "ignored" });
    expect(out.status).toBe("Rejected");
  });

  it("leaves the request's own details untouched", () => {
    const out = applyDecision(row(), { approve: true, by: "x@example.com", at });
    expect(out.itemUniqueId).toBe(row().itemUniqueId);
    expect(out.requestedBy).toBe("pic@example.com");
    expect(out.reason).toBe("Superseded.");
  });
});

describe("queueFor", () => {
  it("returns only pending requests for the approver's units, oldest first", () => {
    const rows = [
      row({ itemName: "c.pdf", requestedAt: "2026-08-15T12:00:00.000Z" }),
      row({ itemName: "a.pdf", requestedAt: "2026-08-15T09:00:00.000Z" }),
      row({ itemName: "other-unit.pdf", unit: "Treasury" }),
      row({ itemName: "done.pdf", status: "Approved" }),
    ];
    expect(queueFor(rows, ["Corporate"]).map((r) => r.itemName)).toEqual(["a.pdf", "c.pdf"]);
  });

  it("is empty for someone who approves nothing", () => {
    expect(queueFor([row()], [])).toEqual([]);
  });
});

describe("counts", () => {
  it("counts every status, including the ones with none", () => {
    const rows = [row(), row(), row({ status: "Approved" }), row({ status: "Failed" })];
    expect(counts(rows)).toEqual({ Pending: 2, Approved: 1, Rejected: 0, Failed: 1 });
  });

  it("is all zeroes for nothing", () => {
    expect(counts([])).toEqual({ Pending: 0, Approved: 0, Rejected: 0, Failed: 0 });
  });
});

describe("decisionSummary", () => {
  it("says a deletion is recoverable — which is what makes approving one reasonable", () => {
    expect(decisionSummary(row())).toContain("recycle bin");
    expect(decisionSummary(row())).toContain("93 days");
  });

  it("names who gets what, and says so when a share never expires", () => {
    const s = decisionSummary(row({
      type: "Share", shareWith: ["john@vendor.com"], sharePermission: "View",
    }));
    expect(s).toContain("john@vendor.com");
    expect(s).toContain("view");
    expect(s).toContain("no expiry date");
  });

  it("names the expiry when there is one", () => {
    const s = decisionSummary(row({
      type: "Share", shareWith: ["a@example.com"], sharePermission: "Edit", expiresAt: "2026-12-31",
    }));
    expect(s).toContain("edit");
    expect(s).toContain("2026-12-31");
  });
});

describe("matching an approver to a unit", () => {
  const guid = "11111111-2222-3333-4444-555555555555";

  it("matches on the term GUID when the row carries one", () => {
    expect(canDecide(row({ unitTermGuid: guid }), [guid])).toBe(true);
  });

  it("IGNORES THE LABEL once a GUID is present, so a renamed term still routes", () => {
    // The failure this prevents: someone tidies a term label, and an approver silently stops seeing
    // their own unit's requests — with nothing on screen to say any went missing.
    const renamed = row({ unit: "Corporate Renamed", unitTermGuid: guid });
    expect(canDecide(renamed, [guid])).toBe(true);
    expect(canDecide(renamed, ["Corporate"])).toBe(false);
  });

  it("falls back to the label for rows written before GUIDs were stored", () => {
    expect(canDecide(row(), ["Corporate"])).toBe(true);
  });

  it("matches nothing when the row identifies no unit at all", () => {
    expect(canDecide(row({ unit: "", unitTermGuid: "" }), ["Corporate"])).toBe(false);
  });
});
