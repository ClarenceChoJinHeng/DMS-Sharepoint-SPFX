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
} from "./mySubmissions";

const LIBS = ["ApprovalDocument", "Documents"];

function sub(over: Partial<Submission> = {}): Submission {
  return {
    itemId: 1,
    library: "ApprovalDocument",
    name: "q1.pdf",
    fileRef: "/sites/Example/ApprovalDocument/NBPOLHO/CDS/UPSUPPORT/2024/Tax Return/q1.pdf",
    status: "Pending",
    comment: "",
    ...over,
  };
}

// The live crash of 2026-08-14: "(intermediate value).trim is not a function", which took the whole
// page down and named no field. Document Type is managed metadata, so it arrives as an OBJECT.
describe("textOf — the fix for a page that would not load", () => {
  it("reads a TAXONOMY object, which is what broke it", () => {
    expect(textOf({ Label: "Working File", TermGuid: "abc", WssId: 12 })).toBe("Working File");
  });

  it("does not throw on the object shape that used to crash", () => {
    // The old code was (value ?? "").trim() — .trim on an object.
    expect(() => textOf({ Label: "x" })).not.toThrow();
    expect(() => textOf({})).not.toThrow();
  });

  it("still reads a plain string, trimmed", () => {
    expect(textOf("  Tax Return  ")).toBe("Tax Return");
  });

  it("treats null and undefined as blank rather than as text", () => {
    expect(textOf(undefined)).toBe("");
    expect(textOf(null)).toBe("");
    // Never the strings "null"/"undefined" in a cell.
    expect(textOf(null)).not.toBe("null");
  });

  it("handles numbers, so a numeric column is shown rather than blanked", () => {
    expect(textOf(0)).toBe("0");
    expect(textOf(42)).toBe("42");
  });

  it("joins a MULTI-value taxonomy field", () => {
    expect(textOf([{ Label: "A" }, { Label: "B" }])).toBe("A; B");
  });

  it("returns blank for a shape it does not recognise — a blank cell, never a crash", () => {
    expect(textOf({ Nope: 1 })).toBe("");
    expect(textOf({ Label: 7 })).toBe("");
  });
});

describe("pickField — the double-encoded key", () => {
  it("finds the plain key", () => {
    expect(pickField({ Document_x0020_Type: "Tax Return" }, "Document_x0020_Type")).toBe("Tax Return");
  });

  it("finds the DOUBLE-ENCODED key SharePoint may return instead", () => {
    // The OData layer re-encodes the underscore in an already-encoded internal name. Reading only
    // the name we wrote yields undefined — a permanently blank column with no error to explain it.
    const row = { Document_x005f_x0020_x005f_Type: { Label: "Working File" } };
    expect(pickField(row, "Document_x0020_Type", "Document_x005f_x0020_x005f_Type")).toBe("Working File");
  });

  it("prefers the first key that actually has a value", () => {
    const row = { A: "", B: "second" };
    expect(pickField(row, "A", "B")).toBe("second");
  });

  it("returns blank when no key matches, and survives a missing row", () => {
    expect(pickField({}, "A", "B")).toBe("");
    expect(pickField(undefined as unknown as Record<string, unknown>, "A")).toBe("");
  });
});

describe("status mapping is NOT redefined here", () => {
  it("re-exports the approver's own mapping", () => {
    // Two mappings of one SharePoint field is how the approver's screen and the uploader's screen
    // end up disagreeing about whether a document was approved.
    expect(statusToDecision(0)).toBe("Approved");
    expect(statusToDecision(1)).toBe("Rejected");
    expect(statusToDecision(2)).toBe("Pending");
  });

  it("treats Draft as Pending, which is what an uploader can act on", () => {
    // 3 = Draft. Not a state the uploader can do anything about, and a fourth badge nobody can
    // explain is worse than a coarser one that is true.
    expect(statusToDecision(3)).toBe("Pending");
  });
});

describe("submissionKey", () => {
  it("distinguishes the same id in different libraries", () => {
    // An approved file in Documents can share an item id with a pending one in the approval
    // library. Keyed on the id alone, one of them is silently dropped.
    const a = sub({ itemId: 7, library: "ApprovalDocument" });
    const b = sub({ itemId: 7, library: "Documents" });
    expect(submissionKey(a)).not.toBe(submissionKey(b));
  });
});

describe("folderTrail", () => {
  it("returns the folders between the library and the file", () => {
    expect(folderTrail(sub().fileRef, LIBS)).toEqual([
      "NBPOLHO", "CDS", "UPSUPPORT", "2024", "Tax Return",
    ]);
  });

  it("drops the file itself", () => {
    expect(folderTrail("/sites/X/Documents/GHO/GF/CORU/a.pdf", LIBS)).toEqual(["GHO", "GF", "CORU"]);
  });

  it("hides WHICH library the file is in — that is what the status badge is for", () => {
    const staging = folderTrail("/sites/X/ApprovalDocument/GHO/GF/CORU/a.pdf", LIBS);
    const documents = folderTrail("/sites/X/Documents/GHO/GF/CORU/a.pdf", LIBS);
    expect(staging).toEqual(documents);
  });

  it("matches the URL SEGMENT, not the display title", () => {
    // "Approval Document" (title) vs "/ApprovalDocument" (URL) — a title match never fires, which
    // would leave the library name sitting in the trail. Gotcha #12.
    expect(folderTrail(sub().fileRef, ["Approval Document"])[0]).not.toBe("NBPOLHO");
    expect(folderTrail(sub().fileRef, ["ApprovalDocument"])[0]).toBe("NBPOLHO");
  });

  it("is case-insensitive, because SharePoint URLs are", () => {
    expect(folderTrail("/sites/X/approvaldocument/GHO/GF/a.pdf", LIBS)).toEqual(["GHO", "GF"]);
  });

  it("falls back to the path below /sites/<site>/ when no library is recognised", () => {
    // A half-recognised path is still a useful hint; an empty cell tells the uploader nothing.
    expect(folderTrail("/sites/X/SomeOtherLib/GHO/GF/a.pdf", LIBS)).toEqual(["SomeOtherLib", "GHO", "GF"]);
  });

  it("handles a file sitting directly in the library root", () => {
    expect(folderTrail("/sites/X/Documents/a.pdf", LIBS)).toEqual([]);
  });

  it("survives blank and malformed input rather than throwing", () => {
    expect(folderTrail("", LIBS)).toEqual([]);
    expect(folderTrail("///", LIBS)).toEqual([]);
  });

  it("with NO library segments given, falls back and keeps the library name", () => {
    // Not a bug: nothing was recognised, so the fallback returns everything below /sites/<site>/,
    // library included. Better a trail with one extra segment than an empty cell — and the caller
    // always passes the real segments, so this is the degraded path, not the normal one.
    expect(folderTrail(sub().fileRef, [])).toEqual([
      "ApprovalDocument", "NBPOLHO", "CDS", "UPSUPPORT", "2024", "Tax Return",
    ]);
  });
});

describe("trailText", () => {
  it("joins with the separator the other screens use", () => {
    expect(trailText(["GHO", "GF", "CORU"])).toBe("GHO › GF › CORU");
  });

  it("is empty for a file at the library root, not a stray separator", () => {
    expect(trailText([])).toBe("");
  });
});

describe("countByStatus", () => {
  it("counts each status and the total", () => {
    const rows = [
      sub({ itemId: 1, status: "Pending" }),
      sub({ itemId: 2, status: "Pending" }),
      sub({ itemId: 3, status: "Approved" }),
      sub({ itemId: 4, status: "Rejected" }),
    ];
    const c = countByStatus(rows);
    expect(c.All).toBe(4);
    expect(c.Pending).toBe(2);
    expect(c.Approved).toBe(1);
    expect(c.Rejected).toBe(1);
  });

  it("reports zeroes rather than missing keys, so a tab never renders 'undefined'", () => {
    expect(countByStatus([])).toEqual({ All: 0, Pending: 0, Approved: 0, Rejected: 0 });
  });

  it("agrees with what filterByTab shows — a tally you cannot reconcile is worse than none", () => {
    const rows = [
      sub({ itemId: 1, status: "Pending" }),
      sub({ itemId: 2, status: "Approved" }),
      sub({ itemId: 3, status: "Approved" }),
    ];
    const c = countByStatus(rows);
    for (const tab of ["All", "Pending", "Approved", "Rejected"]) {
      expect(filterByTab(rows, tab).length).toBe(c[tab]);
    }
  });
});

describe("filterByTab", () => {
  const rows = [sub({ itemId: 1, status: "Pending" }), sub({ itemId: 2, status: "Rejected" })];

  it("All returns everything", () => {
    expect(filterByTab(rows, "All").length).toBe(2);
  });

  it("a status tab returns only that status", () => {
    expect(filterByTab(rows, "Rejected").map((r) => r.itemId)).toEqual([2]);
  });

  it("does not mutate the input", () => {
    const copy = rows.slice();
    filterByTab(rows, "All");
    expect(rows).toEqual(copy);
  });
});

describe("sortNewestFirst", () => {
  it("puts the most recent upload at the top", () => {
    const rows = [
      sub({ itemId: 1, created: new Date(2026, 0, 1) }),
      sub({ itemId: 2, created: new Date(2026, 7, 14) }),
      sub({ itemId: 3, created: new Date(2026, 3, 5) }),
    ];
    expect(sortNewestFirst(rows).map((r) => r.itemId)).toEqual([2, 3, 1]);
  });

  it("sorts an unknown date LAST, never first", () => {
    // An unknown date is not evidence of recency; on top it would displace the row the uploader
    // opened the page to find.
    const rows = [sub({ itemId: 1 }), sub({ itemId: 2, created: new Date(2026, 0, 1) })];
    expect(sortNewestFirst(rows).map((r) => r.itemId)).toEqual([2, 1]);
  });

  it("does not mutate the input", () => {
    const rows = [
      sub({ itemId: 1, created: new Date(2026, 0, 1) }),
      sub({ itemId: 2, created: new Date(2026, 7, 14) }),
    ];
    sortNewestFirst(rows);
    expect(rows.map((r) => r.itemId)).toEqual([1, 2]);
  });
});

describe("formatSubmittedOn", () => {
  it("is DD/MMM/YYYY, the agreed client format", () => {
    expect(formatSubmittedOn(new Date(2035, 0, 8))).toBe("08/Jan/2035");
  });

  it("pads the day", () => {
    expect(formatSubmittedOn(new Date(2026, 11, 1))).toBe("01/Dec/2026");
  });

  it("shows a dash rather than 'Invalid Date' when there is nothing to show", () => {
    expect(formatSubmittedOn(undefined)).toBe("—");
    expect(formatSubmittedOn(new Date("nonsense"))).toBe("—");
  });
});
