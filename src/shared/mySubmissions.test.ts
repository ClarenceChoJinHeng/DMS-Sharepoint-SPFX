import {
  Submission,
  countByStatus,
  filterByTab,
  folderTrail,
  formatSubmittedOn,
  canActDirectly,
  formatSubmittedAt,
  pickField,
  sortNewestFirst,
  statusToDecision,
  submissionKey,
  textOf,
  trailText,
  isHcRow,
  isArchivedRow,
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
    expect(formatSubmittedOn(new Date(2035, 0, 8))).toBe("8 Jan 2035");
  });

  it("pads the day", () => {
    expect(formatSubmittedOn(new Date(2026, 11, 1))).toBe("1 Dec 2026");
  });

  it("shows a dash rather than 'Invalid Date' when there is nothing to show", () => {
    expect(formatSubmittedOn(undefined)).toBe("—");
    expect(formatSubmittedOn(new Date("nonsense"))).toBe("—");
  });
});

describe("formatSubmittedAt", () => {
  it("appends HH:mm in 24-hour form, matching the audit log viewer", () => {
    expect(formatSubmittedAt(new Date(2026, 7, 27, 21, 50))).toBe("27 Aug 2026 21:50");
  });

  it("pads both the hour and the minute", () => {
    expect(formatSubmittedAt(new Date(2026, 7, 27, 9, 5))).toBe("27 Aug 2026 09:05");
    expect(formatSubmittedAt(new Date(2026, 7, 27, 0, 0))).toBe("27 Aug 2026 00:00");
  });

  it("keeps the date half identical to formatSubmittedOn", () => {
    // One definition of the date, reused - so the two screens cannot disagree about it.
    const d = new Date(2035, 0, 8, 13, 7);
    expect(formatSubmittedAt(d).indexOf(formatSubmittedOn(d))).toBe(0);
  });

  it("distinguishes two uploads on the same day", () => {
    // The whole reason it exists: date alone made every row of a busy day look identical.
    const a = formatSubmittedAt(new Date(2026, 7, 27, 21, 50));
    const b = formatSubmittedAt(new Date(2026, 7, 27, 22, 4));
    expect(a).not.toBe(b);
  });

  it("shows a dash rather than 'Invalid Date' when there is nothing to show", () => {
    expect(formatSubmittedAt(undefined)).toBe("—");
    expect(formatSubmittedAt(new Date("nonsense"))).toBe("—");
  });
});

/* ── isHcRow (2026-08-21) ────────────────────────────────────────────────────
   The uploader could not tell an HC document from a same-named ordinary one on My Submissions, so
   a deletion request was a coin-toss from their side. Found on site. */
describe("isHcRow", () => {
  const hc = { approval: "HCApprovalDocument", documents: "HCDocuments" };

  it("recognises both HC libraries", () => {
    expect(isHcRow("HCApprovalDocument", hc)).toBe(true);
    expect(isHcRow("HCDocuments", hc)).toBe(true);
  });

  it("matches case-insensitively — SharePoint URLs are", () => {
    expect(isHcRow("hcapprovaldocument", hc)).toBe(true);
    expect(isHcRow("  HCDOCUMENTS  ", hc)).toBe(true);
  });

  it("does NOT tag the ordinary libraries", () => {
    expect(isHcRow("ApprovalDocument", hc)).toBe(false);
    expect(isHcRow("Shared Documents", hc)).toBe(false);
  });

  it("returns false when the HC pair is unresolved, rather than guessing from the name", () => {
    // A wrong `true` brands an ordinary document as confidential. And it cannot hide a tag that was
    // needed: an uncleared uploader cannot see the HC libraries, so they have no HC rows.
    expect(isHcRow("HCApprovalDocument", undefined)).toBe(false);
    expect(isHcRow("HCApprovalDocument", {})).toBe(false);
  });

  it("survives a blank segment", () => {
    expect(isHcRow("", hc)).toBe(false);
    expect(isHcRow(undefined as unknown as string, hc)).toBe(false);
  });
});

/* -- isArchivedRow (2026-08-22) ---------------------------------------------
 * Spec: docs/superpowers/specs/2026-08-22-seven-year-archive-design.md */
describe("isArchivedRow", () => {
  const arc = { normal: "Archive", hc: "HCArchive" };

  it("tags a row from either archive library", () => {
    expect(isArchivedRow("Archive", arc)).toBe(true);
    expect(isArchivedRow("HCArchive", arc)).toBe(true);
  });

  it("leaves the working libraries alone", () => {
    expect(isArchivedRow("Shared Documents", arc)).toBe(false);
    expect(isArchivedRow("ApprovalDocument", arc)).toBe(false);
  });

  it("compares case-insensitively and ignores stray whitespace", () => {
    // SharePoint URLs are case-insensitive, and the segment is split out of a path.
    expect(isArchivedRow("  archive  ", arc)).toBe(true);
  });

  it("returns FALSE when the archive is unresolved, rather than guessing from the name", () => {
    // A wrong `true` would tell an uploader their live document is frozen and refuse a request they
    // are entitled to make. A site with no archive has no archived rows to mislabel.
    expect(isArchivedRow("Archive", undefined)).toBe(false);
    expect(isArchivedRow("Archive", {})).toBe(false);
  });

  it("never tags a blank segment", () => {
    expect(isArchivedRow("", arc)).toBe(false);
    expect(isArchivedRow(undefined as unknown as string, arc)).toBe(false);
  });

  it("handles a non-HC site, where only the normal archive exists", () => {
    expect(isArchivedRow("Archive", { normal: "Archive" })).toBe(true);
    expect(isArchivedRow("HCArchive", { normal: "Archive" })).toBe(false);
  });
});

describe("canActDirectly — no request where the viewer can already act", () => {
  const DEPT = "11111111-1111-1111-1111-111111111111";
  const UNIT = "22222222-2222-2222-2222-222222222222";
  const OTHER = "33333333-3333-3333-3333-333333333333";

  it("matches a Head of UNIT on the unit term", () => {
    expect(canActDirectly([DEPT, UNIT], [UNIT])).toBe(true);
  });

  /* ⚠ THE CASE THAT DECIDED THE DESIGN. A Head of Department holds DEL/SHARE on the DEPARTMENT
     term, and it fans down to every unit beneath. Comparing the document's unit alone would have
     answered false for every HoD — and needed a term-store expansion to fix. */
  it("matches a Head of DEPARTMENT on the department term in the chain", () => {
    expect(canActDirectly([DEPT, UNIT], [DEPT])).toBe(true);
  });

  it("does NOT match a grant on some other unit", () => {
    expect(canActDirectly([DEPT, UNIT], [OTHER])).toBe(false);
  });

  it("FAILS OPEN: nothing held means the request stays offered", () => {
    expect(canActDirectly([DEPT, UNIT], [])).toBe(false);
  });

  it("FAILS OPEN: an unread tier chain means the request stays offered", () => {
    expect(canActDirectly([], [UNIT])).toBe(false);
  });

  it("tolerates braces and case, because the two sources disagree about both", () => {
    expect(canActDirectly([`{${UNIT.toUpperCase()}}`], [UNIT])).toBe(true);
    expect(canActDirectly([UNIT], [` {${UNIT.toUpperCase()}} `])).toBe(true);
  });

  it("ignores blanks on either side rather than matching them together", () => {
    expect(canActDirectly(["", "  "], ["", "  "])).toBe(false);
  });
});
