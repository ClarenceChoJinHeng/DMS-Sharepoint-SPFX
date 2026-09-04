import {
  ALL_EVENT_TYPES,
  AuditEvent,
  buildAuditRow,
  DETAILS_MAX,
  EVENT,
  EVENT_LABEL,
  eventLabelForRow,
  formatEventTime,
  joinDetails,
  segmentFromUnitPath,
  summarize,
  TITLE_MAX,
} from "./auditLog";

/** A minimal event. Tests override only the field under examination. */
function ev(over: Partial<AuditEvent> = {}): AuditEvent {
  return {
    event: EVENT.approved,
    source: "Test",
    at: new Date(2026, 7, 13, 9, 41),
    ...over,
  };
}

describe("formatEventTime", () => {
  // ISO, not the site-locale format. A plain /items POST goes through the OData layer, which answers
  // a locale string with "Cannot convert a primitive value to the expected type 'Edm.DateTime'".
  // Gotcha #1's M/D/YYYY rule belongs to validateUpdateListItem; applying it here cost a deploy.
  it("writes ISO 8601 in UTC", () => {
    expect(formatEventTime(new Date(Date.UTC(2026, 7, 13, 9, 41)))).toBe("2026-08-13T09:41:00.000Z");
  });

  it("ends in Z, so the value carries its own zone", () => {
    expect(formatEventTime(new Date(2026, 0, 5, 14, 5)).slice(-1)).toBe("Z");
  });

  // Asserted as a round trip rather than a literal: the input here is LOCAL time, so a literal
  // expectation would pass in one timezone and fail in another.
  it("round-trips to the same instant", () => {
    const d = new Date(2026, 0, 5, 14, 5, 33);
    expect(new Date(formatEventTime(d)).getTime()).toBe(d.getTime());
  });

  it("is the same shape $filter uses, so writes and filters cannot disagree", () => {
    expect(formatEventTime(new Date(Date.UTC(2026, 0, 1, 0, 0)))).toMatch(
      /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/,
    );
  });
});

describe("summarize", () => {
  it("uses the event label and the item name", () => {
    expect(summarize(ev({ itemName: "Q1 Tax Return.pdf" }))).toBe("Approved — Q1 Tax Return.pdf");
  });

  it("falls back to the label alone when there is no item", () => {
    expect(summarize(ev({ event: EVENT.policyChanged }))).toBe("File type policy changed");
  });

  it("prefers a caller-supplied summary, because only the caller knows the counts", () => {
    const s = summarize(
      ev({ event: EVENT.reconciliationRun, summary: "Reconciliation — 12 created, 3 renamed" }),
    );
    expect(s).toBe("Reconciliation — 12 created, 3 renamed");
  });

  it("marks a refused action", () => {
    expect(summarize(ev({ event: EVENT.deleted, outcome: "Refused", itemName: "a.pdf" })))
      .toBe("Deleted — a.pdf (refused)");
  });

  it("does NOT restate refusal for UploadRefused, whose label already says it", () => {
    expect(summarize(ev({ event: EVENT.uploadRefused, outcome: "Refused", itemName: "virus.exe" })))
      .toBe("Upload refused — virus.exe");
  });

  it("keeps an unknown event type readable rather than dropping it", () => {
    // Forward compatibility: a row written by a newer build must still say something in an older
    // viewer, and a Text EventType means such a row can exist.
    expect(summarize(ev({ event: "SomethingNew" }))).toBe("SomethingNew");
  });

  it("never exceeds the Title column's limit", () => {
    const long = `${"x".repeat(400)}.pdf`;
    const s = summarize(ev({ itemName: long }));
    expect(s.length).toBeLessThanOrEqual(TITLE_MAX);
    expect(s.slice(-1)).toBe("…");
  });
});

describe("joinDetails", () => {
  it("joins lines with newlines", () => {
    expect(joinDetails(["created GHO", "granted CORU_UPL"])).toBe("created GHO\ngranted CORU_UPL");
  });

  it("drops blank and whitespace-only lines", () => {
    expect(joinDetails(["a", "", "   ", "b"])).toBe("a\nb");
  });

  it("returns an empty string for no lines at all", () => {
    expect(joinDetails(undefined)).toBe("");
    expect(joinDetails([])).toBe("");
  });

  it("ANNOUNCES a truncation, with the real number of dropped lines", () => {
    const lines = Array.from({ length: 50 }, (_, i) => `line ${i} ${"y".repeat(20)}`);
    const out = joinDetails(lines, 200);
    expect(out).toContain("more lines not recorded");
    // The count must be the true remainder, or the announcement is worse than useless.
    const kept = out.split("\n").length - 1;
    expect(out).toContain(`… ${50 - kept} more lines not recorded`);
  });

  it("says 'line' not 'lines' when exactly one was dropped", () => {
    const out = joinDetails(["aaaa", "bbbb"], 65);
    expect(out).toContain("… 1 more line not recorded");
  });

  it("stays within the budget INCLUDING the announcement", () => {
    const lines = Array.from({ length: 200 }, (_, i) => `line ${i}`);
    const out = joinDetails(lines, 300);
    expect(out.length).toBeLessThanOrEqual(300);
  });

  it("does not truncate when everything fits", () => {
    const out = joinDetails(["a", "b", "c"], DETAILS_MAX);
    expect(out).toBe("a\nb\nc");
    expect(out.indexOf("not recorded")).toBe(-1);
  });
});

describe("buildAuditRow", () => {
  it("emits every column, as a string, even when the caller has nothing for it", () => {
    const row = buildAuditRow(ev());
    const expected = [
      "Title", "EventTime", "EventType", "Outcome", "ActorName", "ActorEmail", "Source",
      "LibraryName", "ItemUniqueId", "ItemName", "ItemPath", "Segment", "UnitPath", "Details",
    ];
    expect(Object.keys(row).sort()).toEqual(expected.sort());
    Object.keys(row).forEach((k) => {
      expect(typeof (row as unknown as Record<string, unknown>)[k]).toBe("string");
    });
  });

  it("defaults the outcome to Success", () => {
    expect(buildAuditRow(ev()).Outcome).toBe("Success");
  });

  it("records absent values as blank rather than omitting the field", () => {
    // An omitted field makes one missing value look like a different kind of event later.
    const row = buildAuditRow(ev());
    expect(row.ItemName).toBe("");
    expect(row.ItemPath).toBe("");
    expect(row.UnitPath).toBe("");
  });

  it("does NOT cap ItemPath, which is why it is a Note column", () => {
    // A truncated path cannot be matched back to a folder, defeating the point of recording it.
    const deep = `/sites/Example/ApprovalDocument/${"Level/".repeat(60)}file.pdf`;
    expect(deep.length).toBeGreaterThan(TITLE_MAX);
    expect(buildAuditRow(ev({ itemPath: deep })).ItemPath).toBe(deep);
  });

  it("trims surrounding whitespace on the text fields", () => {
    const row = buildAuditRow(ev({ actorEmail: "  someone@example.com  ", unitPath: " GHO/GF/CORU " }));
    expect(row.ActorEmail).toBe("someone@example.com");
    expect(row.UnitPath).toBe("GHO/GF/CORU");
  });

  it("writes the event time as ISO, which is what a plain /items POST accepts", () => {
    const at = new Date(Date.UTC(2026, 7, 13, 9, 41));
    expect(buildAuditRow(ev({ at })).EventTime).toBe("2026-08-13T09:41:00.000Z");
  });
});

describe("segmentFromUnitPath", () => {
  it("takes the leading segment", () => {
    expect(segmentFromUnitPath("GHO/GF/CORU")).toBe("GHO");
  });

  it("tolerates a leading slash", () => {
    expect(segmentFromUnitPath("/GHO/GF")).toBe("GHO");
  });

  it("returns empty rather than guessing", () => {
    expect(segmentFromUnitPath("")).toBe("");
    expect(segmentFromUnitPath(undefined)).toBe("");
    expect(segmentFromUnitPath("///")).toBe("");
  });
});

describe("the event catalogue", () => {
  it("labels every type it offers, so the viewer never shows a raw key", () => {
    ALL_EVENT_TYPES.forEach((t) => {
      expect(EVENT_LABEL[t]).toBeDefined();
      expect(EVENT_LABEL[t].length).toBeGreaterThan(0);
    });
  });

  it("offers every type defined in EVENT — a type absent here is unfilterable in the viewer", () => {
    const defined = Object.keys(EVENT).map((k) => (EVENT as unknown as Record<string, string>)[k]);
    expect(ALL_EVENT_TYPES.slice().sort()).toEqual(defined.slice().sort());
  });
});

describe("eventLabelForRow", () => {
  // Client, 2026-08-24: an HC routing row read "Moved to Documents" while the Library column beside
  // it said HC Documents. A row that contradicts itself reads as a bug.
  it("names the destination library on a routing event", () => {
    expect(eventLabelForRow(EVENT.routed, "HC Documents")).toBe("Moved to HC Documents");
    expect(eventLabelForRow(EVENT.routed, "Documents")).toBe("Moved to Documents");
  });

  // Every row written before LibraryName existed, and any row whose flow could not resolve it.
  it("falls back to the static label when the library is unknown", () => {
    expect(eventLabelForRow(EVENT.routed, "")).toBe(EVENT_LABEL[EVENT.routed]);
    expect(eventLabelForRow(EVENT.routed, undefined)).toBe(EVENT_LABEL[EVENT.routed]);
    expect(eventLabelForRow(EVENT.routed, "   ")).toBe(EVENT_LABEL[EVENT.routed]);
  });

  // ONLY routing is library-dependent: everything else happens IN a library, not BETWEEN two, so a
  // destination in the label would be wrong rather than merely redundant.
  it("leaves every other event type alone, library or not", () => {
    for (const t of ALL_EVENT_TYPES) {
      if (t === EVENT.routed) continue;
      expect(eventLabelForRow(t, "HC Documents")).toBe(EVENT_LABEL[t] ?? t);
    }
  });

  it("never renders blank for an unknown type", () => {
    expect(eventLabelForRow("SomethingNew", "HC Documents")).toBe("SomethingNew");
  });
});

/* ── Event types written by POWER AUTOMATE, not by this codebase ───────────────
 *
 * ⚠ NO TEST CAN CATCH THE GAP THESE CLOSE, and that is why they are pinned by name. A flow can add a
 * `Create item` with any `EventType` string — the column is TEXT, so the row is written and stored
 * correctly — and nothing in this repo fails. The only symptom is that the viewer's Action dropdown,
 * built from `ALL_EVENT_TYPES`, cannot offer it: the events are in the log and unfilterable. That is
 * how `Replaced` (Auto-route / HC Auto Route) and `ShareRevoked` (CRS — Audit request activity) went
 * missing until the client noticed on 2026-09-03.
 *
 * The existing consistency test proves EVENT and ALL_EVENT_TYPES agree; these prove the two
 * flow-written types are in EVENT at all.
 */
describe("event types the flows write", () => {
  it("offers Replaced as a filter — Auto-route and HC Auto Route both write it", () => {
    expect(ALL_EVENT_TYPES.indexOf("Replaced")).toBeGreaterThan(-1);
  });

  it("offers ShareRevoked as a filter — the request-activity flow writes it on a revoke", () => {
    expect(ALL_EVENT_TYPES.indexOf("ShareRevoked")).toBeGreaterThan(-1);
  });

  it("labels Replaced as what happened to THIS document, not what it did", () => {
    // The bare word reads as though this row's file did the replacing.
    expect(EVENT_LABEL.Replaced).toContain("newer upload");
  });
});
