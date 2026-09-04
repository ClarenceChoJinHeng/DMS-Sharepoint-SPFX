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
  RequestStatus,
  withoutGoneFiles,
  goneFiles,
  canCancel,
  canShareStage,
  stageOf,
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
  ViewerScope,
  toScope,
  collectSharedFiles,
  mergeShareAcl,
  canRevoke,
  endsTheShare,
  holdsRealAccess,
  currentlyShared,
  AclPrincipal,
  revocationNote,
  localDateStamp,
  parseRequestStatus,
  REQUEST_STATUSES,
  FileAcl,
  SharedFile,
  ShareState,
  SharePermission,
  longDate,
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
    expect(errs[0]).toContain("CRS Config");
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

  it("says the domains are UNCONFIGURED rather than calling the person an outsider", () => {
    /* The client's 2026-09-03 report: a colleague on their own tenant was refused with a sentence
       asserting they were outside the organisation. Both states refuse; only one of them is about
       the recipient, and pointing at the wrong one sends the admin to the wrong setting. */
    const off = { allowExternal: false, tenantDomains: [], today: new Date(2026, 7, 15) };
    const errs = validateDraft(draft({ type: "Share", shareWith: "colleague@example.com" }), off);
    expect(errs[0]).toContain("has not been told which email domains");
    expect(errs[0]).toContain("tenantDomains");
    // It must NOT claim this address is outside, because nothing here establishes that.
    expect(errs[0]).not.toContain("colleague@example.com");
  });

  it("names the configured domains when it DOES know them, so the answer is checkable", () => {
    const off = { allowExternal: false, tenantDomains: DOMAINS, today: new Date(2026, 7, 15) };
    const errs = validateDraft(draft({ type: "Share", shareWith: "john@vendor.com" }), off);
    expect(errs[0]).toContain("john@vendor.com");
    expect(errs[0]).toContain(DOMAINS[0]);
  });

  it("treats a blank entry in the domain list as no domain at all", () => {
    // A `tenantDomains` row of "  " parses to one empty string; counting it as configured would
    // make every address external while the message claimed the domains were known.
    const off = { allowExternal: false, tenantDomains: ["  "], today: new Date(2026, 7, 15) };
    const errs = validateDraft(draft({ type: "Share", shareWith: "colleague@example.com" }), off);
    expect(errs[0]).toContain("has not been told which email domains");
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

/* ── Head of Department scope (2026-08-21) ──────────────────────────────────
   Spec: docs/superpowers/specs/2026-08-21-requests-page-hod-access-design.md
   The stage limit is not a policy preference: a HoD holds nothing in either approval library, so a
   pending-file deletion would FAIL in their own session. These pin that it can never be offered. */
describe("ViewerScope — a Head of Department", () => {
  const hod: ViewerScope = { aprUnits: [], hodUnits: ["Corporate"] };

  it("decides an APPROVED-stage request for a unit in their department", () => {
    expect(canDecide(row({ stage: "approved" }), hod)).toBe(true);
    // A row written before the Stage column existed defaults to approved — it was an approved
    // document, because that is all requests could be raised against until 2026-08-20.
    expect(canDecide(row(), hod)).toBe(true);
  });

  it("NEVER decides a PENDING-stage request, because it would fail in their session", () => {
    expect(canDecide(row({ stage: "pending" }), hod)).toBe(false);
  });

  it("cannot even SEE a pending-stage request — hidden means hidden", () => {
    expect(isVisibleTo(row({ stage: "pending" }), "hod@example.com", hod)).toBe(false);
    expect(isVisibleTo(row({ stage: "approved" }), "hod@example.com", hod)).toBe(true);
  });

  it("still sees a pending-stage request they raised THEMSELVES", () => {
    // The requester branch comes first and is about authorship, not scope. A HoD who filed a
    // request must be able to watch it, whatever they can act on.
    const mine = row({ stage: "pending", requestedBy: "hod@example.com" });
    expect(isVisibleTo(mine, "hod@example.com", hod)).toBe(true);
  });

  it("decides nothing outside their department", () => {
    expect(canDecide(row({ unit: "Treasury", stage: "approved" }), hod)).toBe(false);
  });

  it("keeps its queue free of pending-stage rows", () => {
    const rows = [row({ stage: "approved" }), row({ stage: "pending" })];
    expect(queueFor(rows, hod)).toHaveLength(1);
  });
});

describe("ViewerScope — a Head of Unit, and someone who is both", () => {
  it("gives a Head of Unit every stage, unchanged", () => {
    const hou: ViewerScope = { aprUnits: ["Corporate"], hodUnits: [] };
    expect(canDecide(row({ stage: "pending" }), hou)).toBe(true);
    expect(canDecide(row({ stage: "approved" }), hou)).toBe(true);
  });

  it("UNIONS the two sets rather than letting one override the other", () => {
    // Head of Unit of Corporate, Head of Department over Treasury. Pending is decidable in
    // Corporate and not in Treasury — the narrower set must not cap the wider one.
    const both: ViewerScope = { aprUnits: ["Corporate"], hodUnits: ["Treasury"] };
    expect(canDecide(row({ unit: "Corporate", stage: "pending" }), both)).toBe(true);
    expect(canDecide(row({ unit: "Treasury", stage: "pending" }), both)).toBe(false);
    expect(canDecide(row({ unit: "Treasury", stage: "approved" }), both)).toBe(true);
  });
});

describe("toScope", () => {
  it("reads a bare array as APR units, which is what it has always meant", () => {
    expect(toScope(["a"])).toEqual({ aprUnits: ["a"], hodUnits: [], systemAdmin: false });
  });

  it("NEVER infers HoD scope from an array — an old caller cannot widen access by accident", () => {
    expect(toScope(["Corporate"]).hodUnits).toEqual([]);
    expect(canDecide(row({ stage: "pending" }), ["Corporate"])).toBe(true);
  });

  it("NEVER infers systemAdmin from an array — the widest right cannot come from the oldest shape", () => {
    expect(toScope(["Corporate"]).systemAdmin).toBe(false);
    expect(toScope([]).systemAdmin).toBe(false);
  });

  it("treats undefined and missing halves as empty, deciding nothing", () => {
    expect(toScope(undefined)).toEqual({ aprUnits: [], hodUnits: [], systemAdmin: false });
    expect(canDecide(row(), toScope(undefined))).toBe(false);
  });

  it("reads an ABSENT systemAdmin as false, never as unknown-so-allow", () => {
    expect(toScope({ aprUnits: [], hodUnits: [] }).systemAdmin).toBe(false);
  });

  it("normalises a truthy non-boolean to false rather than letting it widen scope", () => {
    // Guards against a value arriving from a loosely typed caller or a JSON round trip.
    const loose = { aprUnits: [], hodUnits: [], systemAdmin: "yes" } as unknown as ViewerScope;
    expect(toScope(loose).systemAdmin).toBe(false);
  });
});

/**
 * A system administrator decides every request, in every unit, at both stages.
 *
 * Client, 2026-08-27: *"their version of system admin means they have access to everything."*
 * Reported the same day — an owner saw only their OWN four requests, because scope is built purely
 * from Group Map rows and an admin holds none by design.
 */
describe("systemAdmin scope", () => {
  const admin: ViewerScope = { aprUnits: [], hodUnits: [], systemAdmin: true };

  it("decides a pending request for a unit it holds no mapping for", () => {
    expect(canDecide(row({ unit: "Treasury" }), admin)).toBe(true);
  });

  it("decides a PENDING-STAGE request, which a Head of Department may not", () => {
    // Safe here and not there: approving runs in the decider's own session, and an owner holds Full
    // Control, so the recycle or the share genuinely succeeds rather than failing after saying yes.
    expect(canDecide(row({ stage: "pending", unit: "Treasury" }), admin)).toBe(true);
  });

  it("STILL REFUSES A REQUEST THAT IS NO LONGER PENDING", () => {
    // The bypass widens WHICH rows, never the settled-request guard: two deciders must not both act.
    expect(canDecide(row({ status: "Approved" }), admin)).toBe(false);
    expect(canDecide(row({ status: "Rejected" }), admin)).toBe(false);
    expect(canDecide(row({ status: "Failed" }), admin)).toBe(false);
  });

  it("sees every request, whoever raised it and whatever unit", () => {
    expect(isVisibleTo(row({ unit: "Treasury" }), "admin@example.com", admin)).toBe(true);
    expect(isVisibleTo(row({ stage: "pending" }), "admin@example.com", admin)).toBe(true);
    expect(isVisibleTo(row({ status: "Rejected" }), "admin@example.com", admin)).toBe(true);
  });

  it("puts those requests in the queue rather than only in the visible list", () => {
    expect(queueFor([row({ unit: "Treasury" })], admin).length).toBe(1);
  });

  it("changes NOTHING when the flag is false — a non-admin is unaffected", () => {
    const plain: ViewerScope = { aprUnits: [], hodUnits: [], systemAdmin: false };
    expect(canDecide(row({ unit: "Treasury" }), plain)).toBe(false);
    expect(isVisibleTo(row({ unit: "Treasury" }), "other@example.com", plain)).toBe(false);
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
    expect(counts(rows)).toEqual({ Pending: 2, Approved: 1, Rejected: 0, Failed: 1, Cancelled: 0, Revoked: 0 });
  });

  it("is all zeroes for nothing", () => {
    expect(counts([])).toEqual({ Pending: 0, Approved: 0, Rejected: 0, Failed: 0, Cancelled: 0, Revoked: 0 });
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

describe("stage — a pending file may be deleted by request, never shared (2026-08-20)", () => {
  const ctx = { allowExternal: true, tenantDomains: ["sdguthrie.com"] };
  const base = {
    itemUniqueId: "9f1c2b3a-0000-4000-8000-000000000001",
    itemName: "Draft.pdf",
    segment: "GHO",
    unit: "Tax",
    reason: "Filed to the wrong unit.",
  };

  it("defaults an absent stage to approved, so rows written before today still route", () => {
    // Same rule as unitTermGuid: every row that predates the field WAS an approved document.
    expect(stageOf(undefined)).toBe("approved");
    expect(stageOf({})).toBe("approved");
    expect(stageOf({ stage: "pending" })).toBe("pending");
    expect(canShareStage(undefined)).toBe(true);
    expect(canShareStage("pending")).toBe(false);
  });

  it("accepts a DELETION request for a pending file", () => {
    // The whole point of the change: a PIC lost DELS, so this is now their only route.
    expect(validateDraft({ ...base, type: "Deletion", stage: "pending" }, ctx)).toEqual([]);
  });

  it("REFUSES a share of a pending file, even when the recipients are perfect", () => {
    // Refused in the RULE, not merely hidden in the form. LIBRARY_ROLES keeps SHARE off both
    // approval libraries, so an approved share would fail in the approver's own session — after
    // they had already told the requester yes.
    const errs = validateDraft(
      { ...base, type: "Share", stage: "pending", shareWith: "someone@sdguthrie.com", sharePermission: "View" },
      ctx,
    );
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("awaiting approval");
  });

  it("does not ALSO demand recipients on a refused pending share", () => {
    // One clear refusal, not that plus "Add at least one person to share with" — the second reads as
    // something the requester could fix, and they cannot.
    const errs = validateDraft({ ...base, type: "Share", stage: "pending" }, ctx);
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("awaiting approval");
  });

  it("still validates recipients normally for an APPROVED document", () => {
    expect(validateDraft({ ...base, type: "Share", stage: "approved" }, ctx))
      .toEqual(["Add at least one person to share with."]);
    expect(validateDraft(
      { ...base, type: "Share", shareWith: "nope", sharePermission: "View" }, ctx,
    ).join(" ")).toContain("Not an email address");
  });

  it("still requires a reason on a pending deletion", () => {
    expect(validateDraft({ ...base, reason: "", type: "Deletion", stage: "pending" }, ctx))
      .toEqual(["Give a reason — the approver sees only this."]);
  });
});

describe("canCancel — the requester withdraws, nobody else (2026-08-20)", () => {
  const row = (over: Partial<RequestRow>): RequestRow => ({
    type: "Deletion", status: "Pending", itemUniqueId: "u1", itemName: "a.pdf",
    segment: "GHO", unit: "Tax", requestedBy: "pic@sdguthrie.com",
    requestedAt: "2026-08-20T09:00:00.000Z", reason: "wrong unit", ...over,
  });

  it("lets the requester cancel their own pending request", () => {
    expect(canCancel(row({}), "pic@sdguthrie.com")).toBe(true);
    // Stored text, so case and padding must not decide whether someone can withdraw their own ask.
    expect(canCancel(row({}), "  PIC@SDGuthrie.com ")).toBe(true);
  });

  it("does NOT let anyone else cancel it — an approver has Reject, which is recorded", () => {
    // Cancelling instead of rejecting would take a decision with no decision on record.
    expect(canCancel(row({}), "hou@sdguthrie.com")).toBe(false);
    expect(canCancel(row({}), "")).toBe(false);
  });

  it("refuses once the request has been DECIDED", () => {
    // Cancelling an approved deletion cannot un-recycle the file, and cancelling an approved share
    // cannot take the access back. Revoke is a separate, unbuilt feature and must not be implied.
    for (const status of ["Approved", "Rejected", "Failed", "Cancelled"] as RequestStatus[]) {
      expect(canCancel(row({ status }), "pic@sdguthrie.com")).toBe(false);
    }
  });

  it("keeps a cancelled request out of every approver's queue", () => {
    // Falls out of canDecide's "must still be Pending" rule — no separate filter to keep in step.
    const cancelled = row({ status: "Cancelled", unitTermGuid: "t1" });
    expect(canDecide(cancelled, ["t1"])).toBe(false);
    expect(queueFor([cancelled], ["t1"])).toEqual([]);
  });

  it("counts Cancelled, so the tally still adds up to what was raised", () => {
    expect(counts([row({ status: "Cancelled" }), row({})])).toEqual({
      Pending: 1, Approved: 0, Rejected: 0, Failed: 0, Cancelled: 1, Revoked: 0,
    });
  });
});

/* -- an archived document is refused (2026-08-22) ----------------------------
 * Spec: docs/superpowers/specs/2026-08-22-seven-year-archive-design.md */
describe("validateDraft: archived documents", () => {
  // Its own ctx: the one above is scoped to its describe, and sharing state between blocks is how a
  // test starts depending on the order they run in.
  const ctx = { allowExternal: true, tenantDomains: DOMAINS, today: new Date(2026, 7, 15) };

  it("refuses a DELETION on an archived document", () => {
    // The approval executes in the approver's OWN session, and every role is Read on the archive —
    // so an approved request would fail at the last step, days after the requester was told yes.
    const errs = validateDraft(draft({ type: "Deletion", archived: true }), ctx);
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("archived");
  });

  it("refuses a SHARE on an archived document too", () => {
    // BOTH types, unlike the pending rule, which refuses Share alone.
    const errs = validateDraft(
      draft({ type: "Share", archived: true, shareWith: "someone@example.com" }),
      ctx,
    );
    expect(errs.length).toBe(1);
    expect(errs[0]).toContain("archived");
  });

  it("says ONLY that, with no second complaint beside it", () => {
    // An "add at least one person" alongside would read as something the requester could fix.
    const errs = validateDraft(draft({ type: "Share", archived: true, shareWith: "" }), ctx);
    expect(errs.length).toBe(1);
    expect(errs[0]).not.toContain("at least one person");
  });

  it("explains that archiving is what changed, not that they did something wrong", () => {
    const msg = validateDraft(draft({ archived: true }), ctx)[0];
    expect(msg).toContain("read-only record");
  });

  it("changes NOTHING when the flag is absent or false", () => {
    // Every caller predating this, and every site with no archive at all.
    expect(validateDraft(draft(), ctx)).toEqual([]);
    expect(validateDraft(draft({ archived: false }), ctx)).toEqual([]);
    expect(validateDraft(draft({ archived: undefined }), ctx)).toEqual([]);
  });
});

/* ── Live shares and revocation (2026-08-28) ──────────────────────────────────
 * Spec: docs/superpowers/specs/2026-08-28-requests-page-redesign-and-share-revoke-design.md
 *
 * Weighted, like the rest of this file, toward the answers that are wrong SILENTLY. Here that is
 * exactly one thing: reporting access as REVOKED when the read did not actually say so. An approver
 * told a share has ended stops looking, and the document stays reachable. */

const share = (over: Partial<RequestRow> = {}): RequestRow => ({
  type: "Share",
  status: "Approved",
  itemUniqueId: "F1",
  itemName: "Q3 Report.pdf",
  segment: "GHO",
  unit: "Tax",
  unitTermGuid: "t1",
  requestedBy: "pic@example.com",
  requestedAt: "2026-08-20T09:00:00.000Z",
  decidedBy: "hou@example.com",
  decidedAt: "2026-08-20T10:00:00.000Z",
  reason: "auditor needs it",
  shareWith: ["auditor@outside.com"],
  sharePermission: "View",
  ...over,
});

const SCOPE: ViewerScope = { aprUnits: ["t1"], hodUnits: [] };

describe("collectSharedFiles", () => {
  it("groups approved share rows by file, keeping who asked and who approved", () => {
    const files = collectSharedFiles([share({ id: 1 })], SCOPE);
    expect(files.length).toBe(1);
    expect(files[0].itemUniqueId).toBe("F1");
    expect(files[0].rowIds).toEqual([1]);
    expect(files[0].recipients[0].email).toBe("auditor@outside.com");
    expect(files[0].recipients[0].requestedBy).toBe("pic@example.com");
    expect(files[0].recipients[0].decidedBy).toBe("hou@example.com");
  });

  it("starts every recipient UNKNOWN — nothing has been probed yet", () => {
    const files = collectSharedFiles([share()], SCOPE);
    expect(files[0].state).toBe("unknown");
    expect(files[0].recipients[0].state).toBe("unknown");
  });

  it("ignores deletions, and anything not yet approved", () => {
    const rows = [
      share({ type: "Deletion" }),
      share({ status: "Pending", itemUniqueId: "F2" }),
      share({ status: "Rejected", itemUniqueId: "F3" }),
    ];
    expect(collectSharedFiles(rows, SCOPE)).toEqual([]);
  });

  it("KEEPS revoked rows — the record that access was granted and taken back", () => {
    const files = collectSharedFiles([share({ status: "Revoked" })], SCOPE);
    expect(files.length).toBe(1);
  });

  it("merges two requests for one file into one entry", () => {
    const rows = [
      share({ id: 1, shareWith: ["a@x.com"] }),
      share({ id: 2, shareWith: ["b@x.com"] }),
    ];
    const files = collectSharedFiles(rows, SCOPE);
    expect(files.length).toBe(1);
    expect(files[0].rowIds).toEqual([1, 2]);
    expect(files[0].recipients.map((r) => r.email).sort()).toEqual(["a@x.com", "b@x.com"]);
  });

  it("lets the NEWEST row win when one recipient appears twice", () => {
    const rows = [
      share({ id: 1, decidedAt: "2026-08-20T10:00:00.000Z", sharePermission: "View" }),
      share({ id: 2, decidedAt: "2026-08-25T10:00:00.000Z", sharePermission: "Edit",
              decidedBy: "hou2@example.com" }),
    ];
    const files = collectSharedFiles(rows, SCOPE);
    expect(files[0].recipients.length).toBe(1);
    expect(files[0].recipients[0].permission).toBe("Edit");
    expect(files[0].recipients[0].decidedBy).toBe("hou2@example.com");
  });

  it("shows a viewer nothing outside their scope", () => {
    expect(collectSharedFiles([share()], { aprUnits: ["other"], hodUnits: [] })).toEqual([]);
  });

  it("shows a system administrator every unit", () => {
    const files = collectSharedFiles([share()], { aprUnits: [], hodUnits: [], systemAdmin: true });
    expect(files.length).toBe(1);
  });

  it("shows a Head of Department their department's units — shares are always approved-stage", () => {
    const files = collectSharedFiles([share()], { aprUnits: [], hodUnits: ["t1"] });
    expect(files.length).toBe(1);
  });
});

describe("mergeShareAcl", () => {
  const file = (): SharedFile => collectSharedFiles([share({ id: 1 })], SCOPE)[0];

  it("⚠ an UNREAD acl leaves everything unknown — never 'revoked'", () => {
    const merged = mergeShareAcl(file(), undefined);
    expect(merged.state).toBe("unknown");
    expect(merged.recipients[0].state).toBe("unknown");
  });

  it("an inheriting file is DEFINITIVELY revoked — no unique scope, no per-file grant", () => {
    const merged = mergeShareAcl(file(), { hasUniqueRoleAssignments: false, principals: [] });
    expect(merged.state).toBe("revoked");
    expect(merged.recipients[0].state).toBe("revoked");
    expect(merged.recipients[0].principalId).toBeUndefined();
  });

  it("marks a recipient live and carries the principal id the revoke needs", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [{ id: 42, principalType: 1, email: "auditor@outside.com" }],
    };
    const merged = mergeShareAcl(file(), acl);
    expect(merged.recipients[0].state).toBe("live");
    expect(merged.recipients[0].principalId).toBe(42);
    expect(merged.state).toBe("live");
  });

  it("marks a recipient revoked when the ACL no longer names them", () => {
    const acl: FileAcl = { hasUniqueRoleAssignments: true, principals: [] };
    expect(mergeShareAcl(file(), acl).recipients[0].state).toBe("revoked");
  });

  it("matches a GUEST claim, where the address has its @ turned into _", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [{
        id: 7, principalType: 1,
        loginName: "i:0#.f|membership|auditor_outside.com#ext#@contoso.onmicrosoft.com",
      }],
    };
    expect(mergeShareAcl(file(), acl).recipients[0].state).toBe("live");
  });

  it("matches an ordinary member claim too", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [{ id: 8, principalType: 1, loginName: "i:0#.f|membership|auditor@outside.com" }],
    };
    expect(mergeShareAcl(file(), acl).recipients[0].state).toBe("live");
  });

  it("⚠ SKIPS GROUPS — breaking inheritance copies the unit's own groups onto the file", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [
        { id: 99, principalType: 8, title: "GHO_GF_TAX_UPLOADER" },
        { id: 42, principalType: 1, email: "auditor@outside.com" },
      ],
    };
    const merged = mergeShareAcl(file(), acl);
    expect(merged.recipients.length).toBe(1);
    expect(merged.recipients[0].email).toBe("auditor@outside.com");
  });

  it("surfaces a user no request row accounts for — a native Share button grant", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [
        { id: 42, principalType: 1, email: "auditor@outside.com" },
        { id: 43, principalType: 1, email: "someone.else@example.com" },
      ],
    };
    const merged = mergeShareAcl(file(), acl);
    const extra = merged.recipients.filter((r) => r.unrecorded)[0];
    expect(extra.email).toBe("someone.else@example.com");
    expect(extra.state).toBe("live");
    expect(extra.principalId).toBe(43);
  });

  it("treats a principal of unknown TYPE as a user — errs towards showing a grant", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [{ id: 42, email: "auditor@outside.com" }],
    };
    expect(mergeShareAcl(file(), acl).recipients[0].state).toBe("live");
  });

  it("⚠ rolls up worst-case: one unknown recipient makes the whole file unknown, never live", () => {
    const two = collectSharedFiles([share({ shareWith: ["a@x.com", "b@x.com"] })], SCOPE)[0];
    // Only `a` is on the ACL; `b` would normally read `revoked`. Prove the LIVE one does not win
    // the rollup on its own — the file is `live` here because nothing is unknown after a full read.
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [{ id: 1, principalType: 1, email: "a@x.com" }],
    };
    const merged = mergeShareAcl(two, acl);
    expect(merged.state).toBe("live");
    // …and an unread ACL makes the same file unknown rather than half-revoked.
    expect(mergeShareAcl(two, undefined).state).toBe("unknown");
  });
});

describe("canRevoke", () => {
  const file = (): SharedFile => collectSharedFiles([share()], SCOPE)[0];

  it("lets the unit's approver revoke", () => {
    expect(canRevoke(file(), SCOPE)).toBe(true);
  });

  it("lets a Head of Department revoke inside their department", () => {
    expect(canRevoke(file(), { aprUnits: [], hodUnits: ["t1"] })).toBe(true);
  });

  it("lets a system administrator revoke anywhere", () => {
    expect(canRevoke(file(), { aprUnits: [], hodUnits: [], systemAdmin: true })).toBe(true);
  });

  it("⚠ refuses somebody with no scope — a PIC holds no Manage Permissions", () => {
    expect(canRevoke(file(), { aprUnits: [], hodUnits: [] })).toBe(false);
  });

  it("refuses an absent file rather than throwing", () => {
    expect(canRevoke(undefined, SCOPE)).toBe(false);
  });
});

describe("holdsRealAccess — Limited Access grants NOTHING", () => {
  const p = (
    bindings?: Array<{ id?: number; name?: string; roleTypeKind?: number }>,
  ): AclPrincipal => ({ id: 9, principalType: 1, email: "x@y.com", bindings });

  it("⚠ ABSENT bindings mean the read did not ask — keeps the principal", () => {
    expect(holdsRealAccess(p(undefined))).toBe(true);
  });

  it("⚠ an EMPTY bindings array is treated the same way, never as 'holds nothing'", () => {
    expect(holdsRealAccess(p([]))).toBe(true);
  });

  it("discards a principal holding ONLY Limited Access, by role definition id", () => {
    expect(holdsRealAccess(p([{ id: 1073741825, name: "Limited Access" }]))).toBe(false);
  });

  it("discards it by RoleTypeKind too, when the id is not what we expect", () => {
    expect(holdsRealAccess(p([{ id: 999, name: "Limited Access", roleTypeKind: 1 }]))).toBe(false);
  });

  it("keeps a real grant that ALSO carries Limited Access alongside it", () => {
    expect(holdsRealAccess(p([
      { id: 1073741825, name: "Limited Access", roleTypeKind: 1 },
      { id: 1073741826, name: "Read", roleTypeKind: 2 },
    ]))).toBe(true);
  });

  it("keeps an ordinary Read", () => {
    expect(holdsRealAccess(p([{ id: 1073741826, name: "Read", roleTypeKind: 2 }]))).toBe(true);
  });
});

describe("mergeShareAcl — a Limited Access entry is not access", () => {
  const file = (): SharedFile => collectSharedFiles([share({ id: 1 })], SCOPE)[0];

  it("⚠ REGRESSION: a revoked recipient left behind as Limited Access reads as revoked", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [{
        id: 42, principalType: 1, email: "auditor@outside.com",
        bindings: [{ id: 1073741825, name: "Limited Access", roleTypeKind: 1 }],
      }],
    };
    expect(mergeShareAcl(file(), acl).recipients[0].state).toBe("revoked");
  });

  it("does not invent an 'unrecorded' recipient out of a Limited Access entry", () => {
    const acl: FileAcl = {
      hasUniqueRoleAssignments: true,
      principals: [
        { id: 42, principalType: 1, email: "auditor@outside.com",
          bindings: [{ id: 1073741826, name: "Read", roleTypeKind: 2 }] },
        { id: 77, principalType: 1, email: "passer.by@x.com",
          bindings: [{ id: 1073741825, name: "Limited Access", roleTypeKind: 1 }] },
      ],
    };
    const merged = mergeShareAcl(file(), acl);
    expect(merged.recipients.length).toBe(1);
    expect(merged.recipients[0].email).toBe("auditor@outside.com");
  });
});

describe("currentlyShared", () => {
  const withStates = (states: ShareState[], id = "F1"): SharedFile => ({
    itemUniqueId: id, itemName: "a.pdf", segment: "GHO", unit: "Tax", unitTermGuid: "t1",
    stage: "approved", state: "live", rowIds: [1],
    recipients: states.map((st, i) => ({
      email: "p" + i + "@x.com", permission: "View" as SharePermission, state: st,
    })),
  });

  it("drops a revoked recipient — the client's ask", () => {
    const out = currentlyShared([withStates(["live", "revoked"])]);
    expect(out.length).toBe(1);
    expect(out[0].recipients.length).toBe(1);
    expect(out[0].recipients[0].state).toBe("live");
  });

  it("drops the file entirely once nobody is left", () => {
    expect(currentlyShared([withStates(["revoked", "revoked"])]).length).toBe(0);
  });

  it("⚠ KEEPS an unknown recipient — a failed read must never shorten this list", () => {
    const out = currentlyShared([withStates(["unknown"])]);
    expect(out.length).toBe(1);
    expect(out[0].state).toBe("unknown");
  });

  it("re-rolls the file's own state from what survives", () => {
    const out = currentlyShared([withStates(["live", "revoked"])]);
    expect(out[0].state).toBe("live");
  });

  it("leaves an untouched list alone and never mutates the input", () => {
    const input = [withStates(["live", "revoked"])];
    currentlyShared(input);
    expect(input[0].recipients.length).toBe(2);
  });

  it("survives an empty or absent list", () => {
    expect(currentlyShared([]).length).toBe(0);
  });
});

describe("endsTheShare", () => {
  const withStates = (states: ShareState[]): SharedFile => ({
    itemUniqueId: "F1", itemName: "a.pdf", segment: "GHO", unit: "Tax", unitTermGuid: "t1",
    stage: "approved", state: "live", rowIds: [1],
    recipients: states.map((st, i) => ({
      email: "p" + i + "@x.com", permission: "View" as SharePermission, state: st,
    })),
  });

  it("is true when the last live recipient is going", () => {
    expect(endsTheShare(withStates(["live"]), ["p0@x.com"])).toBe(true);
  });

  it("is false while another live recipient remains", () => {
    expect(endsTheShare(withStates(["live", "live"]), ["p0@x.com"])).toBe(false);
  });

  it("is true when everyone was already revoked", () => {
    expect(endsTheShare(withStates(["revoked", "revoked"]), [])).toBe(true);
  });

  it("⚠ counts an UNKNOWN recipient as remaining — a read that did not answer is not a revocation", () => {
    expect(endsTheShare(withStates(["live", "unknown"]), ["p0@x.com"])).toBe(false);
  });
});

describe("parseRequestStatus — a terminal status must never come back as Pending", () => {
  /* ⚠ THE LIVE BUG, 2026-08-30: the Requests page carried its own allow-list of four statuses, so
     `Revoked` and `Cancelled` fell through to Pending and sat in the approver's queue permanently —
     undecidable, because the re-read guard rightly refuses a row that is no longer open. */
  it("keeps Revoked as Revoked", () => {
    expect(parseRequestStatus("Revoked")).toBe("Revoked");
  });

  it("keeps Cancelled as Cancelled", () => {
    expect(parseRequestStatus("Cancelled")).toBe("Cancelled");
  });

  it("round-trips EVERY member of the union, so none can be dropped again", () => {
    for (const v of REQUEST_STATUSES) expect(parseRequestStatus(v)).toBe(v);
  });

  it("carries all six — a short set is how the last one went missing", () => {
    expect(REQUEST_STATUSES.length).toBe(6);
  });

  it("trims, because a stored value can carry whitespace", () => {
    expect(parseRequestStatus(" Revoked ")).toBe("Revoked");
  });

  it("blank and unrecognised both read as Pending — a human should see the row, not lose it", () => {
    expect(parseRequestStatus("")).toBe("Pending");
    expect(parseRequestStatus(undefined)).toBe("Pending");
    expect(parseRequestStatus("Whatever")).toBe("Pending");
  });

  it("is case-SENSITIVE, matching what this codebase writes", () => {
    expect(parseRequestStatus("revoked")).toBe("Pending");
  });
});

describe("localDateStamp", () => {
  /* ⚠ THE REGRESSION THIS PINS, found live 2026-08-30: the note read "on 2026-08-29" for a revoke
     the audit row stamped 30/Aug 02:43, because `toISOString()` is UTC and 02:43 local (UTC+8) is
     still the previous day there. Anything before 08:00 local was a day behind. */
  it("uses LOCAL components, so an early-morning revoke is not recorded as yesterday", () => {
    // 02:43 local on the 30th. In UTC+8 this instant is 18:43 UTC on the 29th.
    const early = new Date(2026, 7, 30, 2, 43, 0);
    expect(localDateStamp(early)).toBe("2026-08-30");
  });

  it("pads the month and the day", () => {
    expect(localDateStamp(new Date(2026, 0, 5, 12, 0, 0))).toBe("2026-01-05");
  });

  it("is the last day of the year, not the first of the next", () => {
    expect(localDateStamp(new Date(2026, 11, 31, 23, 59, 0))).toBe("2026-12-31");
  });
});

describe("revocationNote", () => {
  it("names who lost access, who took it, and the day — past tense", () => {
    expect(revocationNote(["a@x.com"], "hou@example.com", "2026-08-28T11:00:00.000Z")).toBe(
      "Access for a@x.com revoked by hou@example.com on 2026-08-28.",
    );
  });

  it("still reads as a sentence when the detail is missing", () => {
    expect(revocationNote([], "", "2026-08-28T11:00:00.000Z")).toBe(
      "Access for the recipients revoked by an approver on 2026-08-28.",
    );
  });
});

describe("fileGone — a deleted document leaves the Shared-files tab", () => {
  const file = (id: string): SharedFile => ({
    itemUniqueId: id,
    itemName: id + ".pdf",
    segment: "GHO",
    unit: "TAX",
    recipients: [{ email: "a@x.com", permission: "View", state: "live" }],
    state: "live",
    rowIds: [1],
  });

  it("drops a file whose ACL says the document is gone", () => {
    const acls = { a: { hasUniqueRoleAssignments: false, principals: [], fileGone: true } };
    expect(withoutGoneFiles([file("a"), file("b")], acls).map((f) => f.itemUniqueId)).toEqual(["b"]);
  });

  it("KEEPS a file whose ACL could not be read — undefined is not gone", () => {
    const acls: Record<string, undefined> = { a: undefined };
    expect(withoutGoneFiles([file("a")], acls).map((f) => f.itemUniqueId)).toEqual(["a"]);
  });

  it("KEEPS a file whose ACL read fine — fileGone absent means present", () => {
    const acls = { a: { hasUniqueRoleAssignments: true, principals: [] } };
    expect(withoutGoneFiles([file("a")], acls).map((f) => f.itemUniqueId)).toEqual(["a"]);
  });

  it("counts only the gone ones", () => {
    const acls = {
      a: { hasUniqueRoleAssignments: false, principals: [], fileGone: true },
      b: undefined,
      c: { hasUniqueRoleAssignments: true, principals: [] },
    };
    expect(goneFiles([file("a"), file("b"), file("c")], acls)).toBe(1);
  });

  it("counts nothing when no ACL has been read at all", () => {
    expect(goneFiles([file("a"), file("b")], {})).toBe(0);
  });
});

describe("longDate", () => {
  it("renders the client's own example", () => {
    expect(longDate("2026-09-02")).toBe("02 September 2026");
  });

  /* The stored value is a full ISO instant; only the DAY may be read from it. Handing this to
     `new Date()` at UTC+8 would report 1 September. */
  it("reads the stored day, never a timezone-shifted one", () => {
    expect(longDate("2026-09-02T00:30:00Z")).toBe("02 September 2026");
  });

  it("pads a single-digit day, matching the example", () => {
    expect(longDate("2026-12-07")).toBe("07 December 2026");
  });

  /* Blank must stay blank: the callers already decide whether to render anything at all. */
  it("is blank for nothing", () => {
    expect(longDate(undefined)).toBe("");
    expect(longDate("")).toBe("");
    expect(longDate("   ")).toBe("");
  });

  /* Unreadable is returned AS IS - never "Invalid Date", never empty. A value we cannot parse is
     still information, and the raw text is reportable where "Invalid Date" is not. */
  it("returns anything unparseable unchanged", () => {
    expect(longDate("not a date")).toBe("not a date");
    expect(longDate("2026-13-02")).toBe("2026-13-02");
    expect(longDate("2026-00-02")).toBe("2026-00-02");
    expect(longDate("2026-09-00")).toBe("2026-09-00");
  });
});
