/**
 * Tests for the submission record.
 * Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
 *
 * The two rules worth breaking a build over:
 *  - an unresolved record is DELETED only when the live read was complete (§2, `unknown`);
 *  - the join is the STAMP, never `UniqueId`, because Auto-route replaces the latter.
 */

import { Submission } from "./mySubmissions";
import {
  RECORD_COLUMNS,
  RECORD_JOIN_COLUMN,
  SubmissionRecord,
  normaliseFileId,
  isJoinable,
  indexLiveByFileId,
  rowFromRecord,
  mergedKey,
  RecordState,
  mergeRecords,
  recordCounts,
  liveRowsOnly,
  encodeMetadata,
  decodeMetadata,
  buildRecordPayload,
  parseRecordRow,
  RECORD_READ_SELECT,
  RECORD_READ_SELECT_NO_ARCHIVE,
  RECORD_READ_SELECT_LEGACY,
  REPLACEMENT_COLUMNS,
  ARCHIVE_COLUMNS,
  SNAPSHOT_FILE_KEYS,
  RECORD_STATE_LABEL,
  recordStateParts,
  archivedRowsOnly,
  snapshotFolderRows,
  snapshotFileRows,
  isBulkUploadRow,
  RECORD_SOURCE,
} from "./submissionRecords";

function rec(over: Partial<SubmissionRecord> = {}): SubmissionRecord {
  return {
    itemId: 1,
    submissionRef: "SUB-20260827-K4P2",
    batchRef: "BAT-20260827-GRWZ",
    fileId: "SFI-20260827-H7QM",
    uniqueId: "accdbebf-0000-0000-0000-000000000000",
    fileName: "ACME - VendorCo - Invoice - 27-08-26.pdf",
    itemPath: "/sites/Example/ApprovalDocument/GHO/GCA/EG/2024/Agreement/invoice.pdf",
    libraryTitle: "Approval Document",
    uploadedBy: "uploader@example.com",
    uploadedAt: new Date("2026-08-27T09:14:00Z"),
    metadata: { Unit: "EG", Year: "2024" },
    source: "Form",
    ...over,
  };
}

function live(over: Partial<Submission> = {}): Submission {
  return {
    itemId: 42,
    library: "Documents",
    name: "ACME - VendorCo - Invoice - 27-08-26.pdf",
    fileRef: "/sites/Example/Shared Documents/GHO/GCA/EG/2024/Agreement/invoice.pdf",
    status: "Approved",
    comment: "",
    ...over,
  };
}

describe("the list schema", () => {
  it("agrees with optionalColumns about the join column name", () => {
    // The literal in RECORD_COLUMNS cannot import the constant — optionalColumns pulls in
    // @microsoft/sp-http, which the test environment cannot resolve. Pinned here instead. A
    // disagreement would mean the writers stamp one name and the record filters another, silently.
    expect(RECORD_COLUMNS.filter((c) => c.name === RECORD_JOIN_COLUMN)).toHaveLength(1);
    expect(RECORD_JOIN_COLUMN).toBe("SubmissionFileId");
  });

  it("stores no status column", () => {
    // §3: status is joined live on every load. A stored copy is the one that goes stale, and the
    // client asked explicitly that the status track the approval.
    const names = RECORD_COLUMNS.map((c) => c.name.toLowerCase());
    for (const banned of ["status", "moderationstatus", "decision", "approved"]) {
      expect(names).not.toContain(banned);
    }
  });

  it("uses only Text, Note and DateTime — never Choice", () => {
    // A value absent from a Choice column's `Choices` fails the WHOLE write.
    for (const c of RECORD_COLUMNS) expect([2, 3, 4]).toContain(c.type);
  });

  it("declares no duplicate column", () => {
    const names = RECORD_COLUMNS.map((c) => c.name);
    expect(names.length).toBe(new Set(names).size);
  });
});

describe("normaliseFileId", () => {
  it("trims and lower-cases", () => {
    expect(normaliseFileId("  SFI-20260827-H7QM ")).toBe("sfi-20260827-h7qm");
  });

  it("answers empty for nothing", () => {
    expect(normaliseFileId(undefined)).toBe("");
    expect(normaliseFileId("   ")).toBe("");
  });
});

describe("isJoinable", () => {
  it("accepts a record with a stamp", () => {
    expect(isJoinable(rec())).toBe(true);
  });

  it("refuses one without", () => {
    // Such a row can never resolve, so showing it would be a permanent false "Deleted".
    expect(isJoinable(rec({ fileId: "" }))).toBe(false);
    expect(isJoinable(rec({ fileId: "   " }))).toBe(false);
  });
});

describe("indexLiveByFileId", () => {
  it("indexes case-insensitively", () => {
    const idx = indexLiveByFileId([live({ submissionFileId: "SFI-20260827-H7QM" })]);
    expect(idx["sfi-20260827-h7qm"]).toBeDefined();
  });

  it("omits rows with no stamp", () => {
    expect(Object.keys(indexLiveByFileId([live()]))).toEqual([]);
  });

  it("keeps the first of two rows sharing a stamp", () => {
    // Replace can put one file's columns onto another's, and mid-route the approval copy and the
    // routed copy briefly coexist. One live document is enough to answer "not deleted".
    const idx = indexLiveByFileId([
      live({ itemId: 1, submissionFileId: "SFI-1" }),
      live({ itemId: 2, submissionFileId: "SFI-1" }),
    ]);
    expect(idx["sfi-1"].itemId).toBe(1);
  });
});

describe("mergeRecords", () => {
  it("keeps every live row, recorded or not", () => {
    // Files uploaded before this existed must behave exactly as they do now.
    const r = mergeRecords([], [live({ itemId: 7 }), live({ itemId: 8 })], true);
    expect(r.rows).toHaveLength(2);
    expect(r.deleted).toBe(0);
  });

  it("adds nothing for a record that resolves", () => {
    const r = mergeRecords(
      [rec({ fileId: "SFI-1" })],
      [live({ submissionFileId: "SFI-1" })],
      true,
    );
    expect(r.rows).toHaveLength(1);
    expect(r.live).toBe(1);
    expect(r.deleted).toBe(0);
  });

  it("matches a resolved record across libraries", () => {
    // THE WHOLE POINT. The record was written against the approval library; the file is now in
    // Documents with a DIFFERENT UniqueId and a different item id. Only the stamp joins them.
    const r = mergeRecords(
      [rec({ fileId: "SFI-1", libraryTitle: "Approval Document", uniqueId: "old-guid" })],
      [live({ library: "Documents", submissionFileId: "SFI-1", uniqueId: "new-guid" })],
      true,
    );
    expect(r.live).toBe(1);
    expect(r.deleted).toBe(0);
    expect(r.rows).toHaveLength(1);
    expect(r.rows[0].recordState).toBeUndefined();
  });

  it("never joins on UniqueId", () => {
    // A live row sharing the record's UniqueId but carrying no stamp must NOT resolve it — that is
    // the agreed design's key, and relying on it is what marked approved files deleted.
    const r = mergeRecords(
      [rec({ fileId: "SFI-1", uniqueId: "same-guid" })],
      [live({ uniqueId: "same-guid" })],
      true,
    );
    expect(r.live).toBe(0);
    expect(r.deleted).toBe(1);
  });

  it("marks an unresolved record deleted when the read was complete", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [], true);
    expect(r.deleted).toBe(1);
    expect(r.unknown).toBe(0);
    expect(r.rows).toHaveLength(1);
    expect(r.rows[0].recordState).toBe("deleted");
  });

  it("marks it UNKNOWN when a library read failed", () => {
    // The rule that stops a throttled read telling an uploader their documents were destroyed.
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [], false);
    expect(r.deleted).toBe(0);
    expect(r.unknown).toBe(1);
    expect(r.rows[0].recordState).toBe("unknown");
  });

  it("drops an unjoinable record rather than showing it as a loss", () => {
    const r = mergeRecords([rec({ fileId: "" })], [], true);
    expect(r.rows).toHaveLength(0);
    expect(r.deleted).toBe(0);
  });

  it("carries name, path, batch and metadata onto a deleted row", () => {
    // "name/path/batch/metadata intact" — the client's requirement for a deleted file.
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [], true);
    const row = r.rows[0];
    expect(row.name).toBe("ACME - VendorCo - Invoice - 27-08-26.pdf");
    expect(row.fileRef).toContain("/GHO/GCA/EG/2024/Agreement/");
    expect(row.submissionId).toBe("SUB-20260827-K4P2");
    expect(row.batchId).toBe("BAT-20260827-GRWZ");
    expect(row.record?.metadata.Unit).toBe("EG");
  });

  it("does not mutate the live rows it was given", () => {
    const src = live({ submissionFileId: "SFI-1" });
    mergeRecords([rec({ fileId: "SFI-1" })], [src], true);
    expect((src as { recordState?: string }).recordState).toBeUndefined();
  });

  it("tolerates empty input", () => {
    expect(mergeRecords([], [], true).rows).toEqual([]);
  });

  it("takes a PREDICATE, so one chain's failed read cannot brand another's records", () => {
    /* The real shape on a site with HC: the HC read was swallowed (uncleared, or throttled) while
       the normal libraries read fine. The normal record must still be judged deleted; the HC one
       must not. A single boolean would have to be false and neither would be judged. */
    const normal = rec({ fileId: "SFI-N", libraryTitle: "Approval Document" });
    const hcRec = rec({ fileId: "SFI-H", libraryTitle: "HC Approval Document" });
    const r = mergeRecords(
      [normal, hcRec],
      [],
      (rr) => !rr.libraryTitle.startsWith("HC "),
    );
    expect(r.deleted).toBe(1);
    expect(r.unknown).toBe(1);
    const byStamp: Record<string, string | undefined> = {};
    for (const row of r.rows) byStamp[row.submissionFileId ?? ""] = row.recordState;
    expect(byStamp["SFI-N"]).toBe("deleted");
    expect(byStamp["SFI-H"]).toBe("unknown");
  });

  it("reads a THROWING predicate as not-complete, never as deleted", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [], () => {
      throw new Error("could not decide");
    });
    expect(r.deleted).toBe(0);
    expect(r.unknown).toBe(1);
  });

  it("still accepts a plain boolean", () => {
    expect(mergeRecords([rec({ fileId: "SFI-1" })], [], true).deleted).toBe(1);
    expect(mergeRecords([rec({ fileId: "SFI-1" })], [], false).unknown).toBe(1);
  });
});

describe("rowFromRecord", () => {
  it("does not claim an approval outcome", () => {
    // `status` exists only because Submission demands it. The screen must branch on recordState.
    const row = rowFromRecord(rec(), "deleted");
    expect(row.recordState).toBe("deleted");
    expect(row.record).toBeDefined();
  });
});

describe("mergedKey", () => {
  it("keys a live row by library and item id", () => {
    expect(mergedKey(live({ itemId: 9, library: "Documents" }))).toBe("Documents#9");
  });

  it("keys a record row by its stamp", () => {
    // Two records could share an item id with each other and with a live document — the record's id
    // belongs to a different list entirely.
    expect(mergedKey(rowFromRecord(rec({ fileId: "SFI-1" }), "deleted"))).toBe("rec#sfi-1");
  });

  it("gives two deleted records distinct keys", () => {
    const a = mergedKey(rowFromRecord(rec({ itemId: 1, fileId: "SFI-1" }), "deleted"));
    const b = mergedKey(rowFromRecord(rec({ itemId: 1, fileId: "SFI-2" }), "deleted"));
    expect(a).not.toBe(b);
  });
});

describe("recordCounts and liveRowsOnly", () => {
  it("counts the five states apart", () => {
    const rows = [
      live(),
      rowFromRecord(rec({ fileId: "SFI-1" }), "deleted"),
      rowFromRecord(rec({ fileId: "SFI-2" }), "unknown"),
      rowFromRecord(rec({ fileId: "SFI-3" }), "cancelled"),
      rowFromRecord(rec({ fileId: "SFI-4" }), "archived"),
    ];
    expect(recordCounts(rows)).toEqual({ live: 1, deleted: 1, unknown: 1, cancelled: 1, archived: 1 });
  });

  it("excludes every gone state from the approval counts", () => {
    const rows = [
      live(),
      rowFromRecord(rec({ fileId: "SFI-1" }), "deleted"),
      rowFromRecord(rec({ fileId: "SFI-2" }), "unknown"),
      rowFromRecord(rec({ fileId: "SFI-3" }), "cancelled"),
      rowFromRecord(rec({ fileId: "SFI-4" }), "archived"),
    ];
    expect(liveRowsOnly(rows)).toHaveLength(1);
  });
});

/* ── Replaced submissions read as Cancelled (2026-08-28) ───────────────────────
 * Client: *"The older submission under My Submission will change to Cancelled status if replaced"*,
 * then *"I know its weird but client want it to be Cancelled"*.
 *
 * The dangerous direction here is the same one this whole module exists to prevent, arriving through
 * a new door: a replaced document reading as DELETED tells its uploader it was destroyed. The other
 * direction matters too — claiming `cancelled` for a file that is demonstrably still there. */
describe("mergeRecords: replaced files", () => {
  const REPLACED = new Date("2026-08-28T09:14:00.000Z");

  it("reads a replaced record as cancelled, not deleted", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1", replacedAt: REPLACED })], [], true);
    expect(r.rows[0].recordState).toBe("cancelled");
    expect(r.cancelled).toBe(1);
    expect(r.deleted).toBe(0);
  });

  it("⚠ says cancelled even when the live read FAILED — it is observed, not derived", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1", replacedAt: REPLACED })], [], false);
    expect(r.rows[0].recordState).toBe("cancelled");
    expect(r.unknown).toBe(0);
  });

  it("⚠ but LIVE still wins: a record that resolves describes a document that exists", () => {
    const doc = live({ submissionFileId: "SFI-1" });
    const r = mergeRecords([rec({ fileId: "SFI-1", replacedAt: REPLACED })], [doc], true);
    expect(r.live).toBe(1);
    expect(r.cancelled).toBe(0);
    // The live row is the document itself, carrying no record state at all.
    expect(r.rows[0].recordState).toBeUndefined();
  });

  it("leaves an ordinary unresolved record as deleted", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [], true);
    expect(r.rows[0].recordState).toBe("deleted");
    expect(r.cancelled).toBe(0);
  });

  it("carries who replaced it onto the row, for the screen to show", () => {
    const r = mergeRecords(
      [rec({ fileId: "SFI-1", replacedAt: REPLACED, replacedBy: "hou@example.com" })],
      [],
      true,
    );
    expect(r.rows[0].record?.replacedBy).toBe("hou@example.com");
    expect(r.rows[0].record?.replacedAt).toEqual(REPLACED);
  });
});

/* ── Archived submissions read as Archived (2026-09-03) ────────────────────────
 * Client: *"why all the archive files consider as deleted, can we ensure if its archive it shows
 * archive instead?"*
 *
 * The dangerous direction is the SAME one this module exists to prevent, through yet another door:
 * a file the seven-year mover archived is retained, not destroyed, and this page cannot see the
 * archive at all (C-Level only since 2026-09-02) — so without the stamp every archived document
 * reads as a loss to the person who uploaded it. */
describe("mergeRecords: archived files", () => {
  const ARCHIVED = new Date("2026-09-02T16:00:00.000Z");

  it("reads an archived record as archived, not deleted", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1", archivedAt: ARCHIVED })], [], true);
    expect(r.rows[0].recordState).toBe("archived");
    expect(r.archived).toBe(1);
    expect(r.deleted).toBe(0);
  });

  it("⚠ says archived even when the live read FAILED — it is observed, not derived", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1", archivedAt: ARCHIVED })], [], false);
    expect(r.rows[0].recordState).toBe("archived");
    expect(r.unknown).toBe(0);
  });

  it("⚠ but LIVE still wins — a stale stamp must not hide a document the viewer can see", () => {
    const doc = live({ submissionFileId: "SFI-1" });
    const r = mergeRecords([rec({ fileId: "SFI-1", archivedAt: ARCHIVED })], [doc], true);
    expect(r.live).toBe(1);
    expect(r.archived).toBe(0);
    expect(r.rows[0].recordState).toBeUndefined();
  });

  it("⚠ a REPLACED file stays cancelled even if it also carries an archive stamp", () => {
    // They cannot honestly co-occur, so this pins the tie-break rather than a real scenario: the
    // replacement is the later fact about what became of THIS record.
    const r = mergeRecords(
      [rec({ fileId: "SFI-1", archivedAt: ARCHIVED, replacedAt: new Date("2026-09-03T01:00:00.000Z") })],
      [],
      true,
    );
    expect(r.rows[0].recordState).toBe("cancelled");
    expect(r.archived).toBe(0);
  });

  it("leaves an ordinary unresolved record as deleted", () => {
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [], true);
    expect(r.rows[0].recordState).toBe("deleted");
    expect(r.archived).toBe(0);
  });
});

describe("parseRecordRow: the archive column", () => {
  const base = {
    Id: 8, SubmissionRef: "SUB-20260903-K4P2", BatchRef: "BAT-20260903-GRWZ",
    SubmissionFileId: "SFI-20260903-BGWU", FileName: "a.pdf", ItemPath: "/sites/x/Documents/a.pdf",
    LibraryTitle: "Restricted & Confidential Document", UploadedBy: "pic@example.com", Source: "Form",
  };

  it("reads it when present", () => {
    expect(parseRecordRow({ ...base, ArchivedAt: "2026-09-02T16:00:00.000Z" }).archivedAt)
      .toEqual(new Date("2026-09-02T16:00:00.000Z"));
  });

  it("⚠ an unparseable date is undefined, NEVER Invalid Date", () => {
    // `archivedAt !== undefined` decides `archived`, so a truthy Invalid Date would report a file as
    // archived on the strength of a value nobody could read.
    expect(parseRecordRow({ ...base, ArchivedAt: "not a date" }).archivedAt).toBeUndefined();
  });

  it("treats absent and blank as not-archived — the pre-2026-09-03 behaviour", () => {
    expect(parseRecordRow(base).archivedAt).toBeUndefined();
    expect(parseRecordRow({ ...base, ArchivedAt: "" }).archivedAt).toBeUndefined();
  });
});

describe("parseRecordRow: replacement columns", () => {
  const base = {
    Id: 7, SubmissionRef: "SUB-20260828-K4P2", BatchRef: "BAT-20260828-GRWZ",
    SubmissionFileId: "SFI-20260828-BGWU", FileName: "a.pdf", ItemPath: "/sites/x/Documents/a.pdf",
    LibraryTitle: "Documents", UploadedBy: "pic@example.com", Source: "Form",
  };

  it("reads both columns when present", () => {
    const r = parseRecordRow({ ...base, ReplacedAt: "2026-08-28T09:14:00.000Z", ReplacedBy: "hou@example.com" });
    expect(r.replacedAt).toEqual(new Date("2026-08-28T09:14:00.000Z"));
    expect(r.replacedBy).toBe("hou@example.com");
  });

  it("leaves them undefined on a list that does not have them", () => {
    const r = parseRecordRow(base);
    expect(r.replacedAt).toBeUndefined();
    expect(r.replacedBy).toBeUndefined();
  });

  it("⚠ treats an UNPARSEABLE date as not-replaced, never as Invalid Date", () => {
    // `replacedAt !== undefined` is what decides `cancelled`, and an Invalid Date is truthy — so
    // this would mark a record replaced on the strength of a value nobody could read.
    const r = parseRecordRow({ ...base, ReplacedAt: "not a date" });
    expect(r.replacedAt).toBeUndefined();
  });

  it("treats a blank as not-replaced", () => {
    expect(parseRecordRow({ ...base, ReplacedAt: "", ReplacedBy: "" }).replacedAt).toBeUndefined();
  });
});

describe("RECORD_READ_SELECT_LEGACY", () => {
  it("is the full read minus exactly the optional columns", () => {
    const full = RECORD_READ_SELECT.split(",");
    const legacy = RECORD_READ_SELECT_LEGACY.split(",");
    const optional = REPLACEMENT_COLUMNS.concat(ARCHIVE_COLUMNS);
    expect(legacy).toEqual(full.filter((c) => optional.indexOf(c) === -1));
    for (const c of optional) expect(legacy).not.toContain(c);
  });

  /* ⚠ THE MIDDLE RUNG IS WHY THIS LADDER HAS THREE STEPS, and these pin the failure it prevents: a
     site holding `ReplacedAt` but not `ArchivedAt` must lose ONLY the archive column. Dropping
     straight to the legacy select would take the replacement state with it, and every replaced file
     there would silently go back to reading "Deleted" — the state the client asked us to stop
     showing. `RevokedBy` taught this on `CRS Requests` (2026-08-30). */
  it("has a middle rung that keeps the replacement columns and drops only the archive one", () => {
    const mid = RECORD_READ_SELECT_NO_ARCHIVE.split(",");
    for (const c of REPLACEMENT_COLUMNS) expect(mid).toContain(c);
    for (const c of ARCHIVE_COLUMNS) expect(mid).not.toContain(c);
    expect(mid).toContain(RECORD_JOIN_COLUMN);
  });

  it("declares the archive column on the list, or the write would 400 for ever", () => {
    const declared = RECORD_COLUMNS.map((c) => c.name);
    for (const c of ARCHIVE_COLUMNS) expect(declared).toContain(c);
  });

  it("⚠ still asks for the JOIN KEY — a fallback that drops it would resolve nothing", () => {
    expect(RECORD_READ_SELECT_LEGACY.split(",")).toContain(RECORD_JOIN_COLUMN);
  });

  it("declares both replacement columns on the list, or the write would 400 for ever", () => {
    const declared = RECORD_COLUMNS.map((c) => c.name);
    for (const c of REPLACEMENT_COLUMNS) expect(declared).toContain(c);
  });
});

describe("the metadata snapshot", () => {
  it("round-trips", () => {
    expect(decodeMetadata(encodeMetadata({ Unit: "EG", Year: "2024" })))
      .toEqual({ Unit: "EG", Year: "2024" });
  });

  it("drops blanks rather than storing empty fields", () => {
    expect(decodeMetadata(encodeMetadata({ Unit: "EG", Remark: "  " })))
      .toEqual({ Unit: "EG" });
  });

  it("answers {} for anything malformed, and never throws", () => {
    // A bad snapshot must cost the metadata table and nothing else — the row's purpose is to survive.
    const bad: Array<string | undefined> = ["", "   ", "not json", "[1,2]", "null", "\"text\"", undefined];
    for (const b of bad) expect(decodeMetadata(b)).toEqual({});
  });

  it("coerces numbers and booleans, ignores nested objects", () => {
    expect(decodeMetadata('{"a":1,"b":true,"c":{"d":1}}')).toEqual({ a: "1", b: "true" });
  });
});

describe("buildRecordPayload", () => {
  it("writes UploadedAt as ISO, never a locale string", () => {
    // A plain /items POST goes through the OData layer and rejects M/D/YYYY with "Cannot convert a
    // primitive value to the expected type 'Edm.DateTime'". That format belongs to
    // validateUpdateListItem alone.
    const p = buildRecordPayload(rec({ uploadedAt: new Date("2026-08-27T09:14:00Z") }));
    expect(p.UploadedAt).toBe("2026-08-27T09:14:00.000Z");
    expect(p.UploadedAt).not.toMatch(/^\d{1,2}\/\d{1,2}\/\d{4}/);
  });

  it("OMITS UploadedAt when there is no date, rather than sending an empty string", () => {
    // A DateTime column rejects "" and takes the WHOLE row with it — how a blank expiry once lost a
    // request row.
    const p = buildRecordPayload(rec({ uploadedAt: undefined }));
    expect("UploadedAt" in p).toBe(false);
  });

  it("omits ItemUniqueId when blank, and keeps it when present", () => {
    expect("ItemUniqueId" in buildRecordPayload(rec({ uniqueId: "" }))).toBe(false);
    expect("ItemUniqueId" in buildRecordPayload(rec({ uniqueId: "   " }))).toBe(false);
    expect(buildRecordPayload(rec({ uniqueId: "abc" })).ItemUniqueId).toBe("abc");
  });

  it("carries the join key and the two references", () => {
    const p = buildRecordPayload(rec());
    expect(p.SubmissionFileId).toBe("SFI-20260827-H7QM");
    expect(p.SubmissionRef).toBe("SUB-20260827-K4P2");
    expect(p.BatchRef).toBe("BAT-20260827-GRWZ");
  });

  it("puts the file name in Title so the list view is readable", () => {
    expect(buildRecordPayload(rec()).Title).toBe("ACME - VendorCo - Invoice - 27-08-26.pdf");
  });

  it("lower-cases UploadedBy", () => {
    expect(buildRecordPayload(rec({ uploadedBy: "Uploader@Example.COM" })).UploadedBy)
      .toBe("uploader@example.com");
  });

  it("writes no status-shaped field", () => {
    // §3: status is joined live. A stored copy is the one that goes stale.
    const keys = Object.keys(buildRecordPayload(rec())).map((k) => k.toLowerCase());
    for (const banned of ["status", "moderationstatus", "decision", "approved"]) {
      expect(keys).not.toContain(banned);
    }
  });

  it("writes only columns the list actually has", () => {
    // A field name the list does not carry fails the whole row. Title is a built-in.
    const declared = RECORD_COLUMNS.map((c) => c.name).concat(["Title"]);
    for (const k of Object.keys(buildRecordPayload(rec()))) expect(declared).toContain(k);
  });
});

describe("parseRecordRow", () => {
  it("round-trips a payload back into a record", () => {
    const p = buildRecordPayload(rec());
    const back = parseRecordRow({ Id: 7, ...p });
    expect(back.itemId).toBe(7);
    expect(back.fileId).toBe("SFI-20260827-H7QM");
    expect(back.submissionRef).toBe("SUB-20260827-K4P2");
    expect(back.batchRef).toBe("BAT-20260827-GRWZ");
    expect(back.metadata).toEqual({ Unit: "EG", Year: "2024" });
    expect(back.uploadedAt?.toISOString()).toBe("2026-08-27T09:14:00.000Z");
  });

  it("survives nulls on every column", () => {
    // SharePoint returns null for an empty column. Nothing here may throw.
    const back = parseRecordRow({
      Id: 1, SubmissionRef: null, BatchRef: null, SubmissionFileId: null, ItemUniqueId: null,
      FileName: null, ItemPath: null, LibraryTitle: null, UploadedBy: null, UploadedAt: null,
      MetadataSnapshot: null, Source: null,
    } as unknown as Record<string, unknown>);
    expect(back.fileId).toBe("");
    expect(back.uploadedAt).toBeUndefined();
    expect(back.metadata).toEqual({});
    expect(isJoinable(back)).toBe(false);
  });

  it("answers undefined for an unparseable date, never Invalid Date", () => {
    // `Invalid Date` formats on screen as "Invalid Date" and sorts unpredictably.
    const back = parseRecordRow({ Id: 1, UploadedAt: "not a date" });
    expect(back.uploadedAt).toBeUndefined();
  });

  it("does not throw on an object-valued field", () => {
    // A taxonomy value arrives as { Label, TermGuid, WssId } and took the page down on 2026-08-14.
    const back = parseRecordRow({ Id: 1, FileName: { Label: "x" } as unknown as string });
    expect(back.fileName).toBe("");
  });

  it("tolerates an empty object", () => {
    const back = parseRecordRow({});
    expect(back.itemId).toBe(0);
    expect(isJoinable(back)).toBe(false);
  });

  it("reads a record whose file has since been routed, and still joins", () => {
    // The end-to-end point: the record was written against the approval library, the file now lives
    // in Documents under a different UniqueId, and the stamp is what connects them.
    const back = parseRecordRow({ Id: 3, ...buildRecordPayload(rec({ fileId: "SFI-9" })) });
    const merged = mergeRecords([back], [live({ library: "Documents", submissionFileId: "SFI-9" })], true);
    expect(merged.live).toBe(1);
    expect(merged.deleted).toBe(0);
  });
});

describe("RECORD_READ_SELECT", () => {
  it("asks for every declared column plus Id", () => {
    // One unknown name fails the WHOLE $select (gotcha #11), and a missing one is a silently blank
    // field — so the select and the schema are pinned to each other.
    const asked = RECORD_READ_SELECT.split(",");
    expect(asked).toContain("Id");
    for (const c of RECORD_COLUMNS) expect(asked).toContain(c.name);
  });

  it("asks for nothing the list does not have", () => {
    const declared = RECORD_COLUMNS.map((c) => c.name).concat(["Id"]);
    for (const a of RECORD_READ_SELECT.split(",")) expect(declared).toContain(a);
  });
});

/**
 * Which tab a row belongs on.
 *
 * ⚠ THE POINT OF THESE IS THE FAIL DIRECTION. An unknown or absent source must leave a row exactly
 * where it already was — anything uploaded before the record feature has no source at all, and there
 * are far more of those than there are bulk imports.
 */
describe("isBulkUploadRow", () => {
  const row = (source?: string): Parameters<typeof isBulkUploadRow>[0] =>
    ({
      // Only `record` is read; the rest satisfies the row shape.
      name: "x.pdf", library: "Approval Document", fileRef: "/a/x.pdf",
      status: "Pending", uniqueId: "u", itemId: 1,
      record: source === undefined ? undefined : ({ source } as never),
    }) as never;

  it("recognises a bulk import", () => {
    expect(isBulkUploadRow(row(RECORD_SOURCE.bulk))).toBe(true);
  });

  it("leaves an ordinary upload where it is", () => {
    expect(isBulkUploadRow(row(RECORD_SOURCE.form))).toBe(false);
  });

  /* The commonest row on any site that predates the record feature. An exclusion test would sweep
     every one of these into the Bulk Upload tab. */
  it("leaves a row with NO record where it is", () => {
    expect(isBulkUploadRow(row(undefined))).toBe(false);
  });

  it("leaves an unrecognised source where it is", () => {
    expect(isBulkUploadRow(row("Migration"))).toBe(false);
    expect(isBulkUploadRow(row(""))).toBe(false);
  });

  /* Survives a round trip through a SharePoint text column: a stray space or a case change would
     otherwise empty the tab with nothing to explain why. */
  it("tolerates whitespace and case", () => {
    expect(isBulkUploadRow(row("  BulkUpload "))).toBe(true);
    expect(isBulkUploadRow(row("bulkupload"))).toBe(true);
  });

  /* ⚠ `UploadForm` is what Form.tsx writes to the AUDIT LOG, not to a record. Pinned so nobody
     "tidies" the two source constants into one list and starts matching the wrong field. */
  it("does not confuse the audit log's source with a record's", () => {
    expect(isBulkUploadRow(row("UploadForm"))).toBe(false);
  });
});

/* ── Every record-backed state keys as a record (2026-08-28) ───────────────────
 * `mergedKey` tested `deleted || unknown` by name and MISSED `cancelled` the day it was added, so a
 * replaced row fell through to `library#itemId` — an id belonging to the RECORD list, which can
 * collide with a live document's key. React drops a duplicate-keyed row SILENTLY.
 *
 * Pinned by iterating the states rather than naming two of them, so the next one added is covered
 * without anybody remembering to come back here. */
describe("mergedKey: record-backed rows", () => {
  const states: RecordState[] = ["deleted", "cancelled", "unknown"];

  it("keys EVERY gone state off the stamp, never off library#itemId", () => {
    for (const st of states) {
      const row = rowFromRecord(rec({ fileId: "SFI-1", itemId: 42 }), st);
      expect(mergedKey(row)).toBe("rec#sfi-1");
    }
  });

  it("⚠ cannot collide with a live document that shares the record's item id", () => {
    const liveRow = live({ itemId: 42, library: "Documents" });
    for (const st of states) {
      const goneRow = rowFromRecord(rec({ fileId: "SFI-1", itemId: 42 }), st);
      expect(mergedKey(goneRow)).not.toBe(mergedKey(liveRow));
    }
  });

  it("still keys a LIVE row by library and item id", () => {
    expect(mergedKey(live({ itemId: 7, library: "Documents" }))).toBe("Documents#7");
  });
});

/* ── A live row takes its date from its record (2026-08-28) ────────────────────
 * Verified live: a replacement and the submission it displaced were TEN MINUTES apart in the
 * records and both displayed 16:21, with the displaced one sorting ABOVE its own replacement.
 * `Files/Add(overwrite=true)` keeps the list item, so the replacing file inherits the original's
 * `Created`; the record, stamped after the original upload finished, is a few seconds later.
 *
 * `Created` describes the FILE. `uploadedAt` describes the SUBMISSION. This page groups by
 * submission, so the record wins. */
describe("mergeRecords: the record's timestamp wins on a live row", () => {
  const UPLOADED = new Date("2026-08-28T08:31:53.000Z");
  const STALE = new Date("2026-08-28T08:21:48.000Z");

  it("replaces a live row's created with the record's uploadedAt", () => {
    const doc = live({ submissionFileId: "SFI-1", created: STALE });
    const r = mergeRecords([rec({ fileId: "SFI-1", uploadedAt: UPLOADED })], [doc], true);
    expect(r.rows[0].created).toEqual(UPLOADED);
  });

  it("⚠ a replacement now sorts ABOVE the submission it displaced", () => {
    const replacement = live({ submissionFileId: "SFI-NEW", created: STALE });
    const merged = mergeRecords(
      [
        rec({ fileId: "SFI-NEW", uploadedAt: UPLOADED }),
        rec({ fileId: "SFI-OLD", uploadedAt: new Date("2026-08-28T08:21:56.000Z"),
              replacedAt: UPLOADED }),
      ],
      [replacement],
      true,
    );
    const sorted = merged.rows.slice().sort((a, b) =>
      (b.created ? b.created.getTime() : 0) - (a.created ? a.created.getTime() : 0));
    expect(sorted[0].submissionFileId).toBe("SFI-NEW");
    expect(sorted[1].recordState).toBe("cancelled");
  });

  it("keeps the live date when the record has no uploadedAt", () => {
    const doc = live({ submissionFileId: "SFI-1", created: STALE });
    const r = mergeRecords([rec({ fileId: "SFI-1", uploadedAt: undefined })], [doc], true);
    expect(r.rows[0].created).toEqual(STALE);
  });

  it("leaves a row with no record entirely alone — pre-feature uploads are untouched", () => {
    const doc = live({ submissionFileId: "", created: STALE });
    const r = mergeRecords([], [doc], true);
    expect(r.rows[0].created).toEqual(STALE);
    expect(r.rows[0].record).toBeUndefined();
  });

  it("⚠ does NOT set recordState on a live row — several call sites test it for truthiness", () => {
    // `row.recordState ? goneRow : fileCard` would grey out every recorded row and remove its
    // Open file button. Absent means live, and that is the contract.
    const doc = live({ submissionFileId: "SFI-1", created: STALE });
    const r = mergeRecords([rec({ fileId: "SFI-1", uploadedAt: UPLOADED })], [doc], true);
    expect(r.rows[0].recordState).toBeUndefined();
    expect(liveRowsOnly(r.rows)).toHaveLength(1);
  });

  it("attaches the record, so a LIVE bulk upload reaches the Bulk Upload tab", () => {
    // It read (0) on a site full of them: only gone rows carried a record, and `isBulkUploadRow`
    // reads `row.record?.source`.
    const doc = live({ submissionFileId: "SFI-1" });
    const r = mergeRecords([rec({ fileId: "SFI-1", source: "BulkUpload" })], [doc], true);
    expect(isBulkUploadRow(r.rows[0])).toBe(true);
  });

  it("still counts a resolved record as live, not as a second row", () => {
    const doc = live({ submissionFileId: "SFI-1" });
    const r = mergeRecords([rec({ fileId: "SFI-1" })], [doc], true);
    expect(r.rows).toHaveLength(1);
    expect(r.live).toBe(1);
  });
});

/* ── The snapshot partition ────────────────────────────────────────────────────
 *
 * A gone file's snapshot is the only record of what it carried, and the batch view renders it in TWO
 * cards — the folder tiers at the top, the per-document fields on the file. Until 2026-09-03 the file
 * card rendered every key and the folder card rendered nothing, which is what the client reported.
 * These pin the partition as exact: nothing lost, nothing shown twice.
 */
describe("snapshot rows: the folder half and the file half", () => {
  const snap = {
    BusinessSegment: "Group Head Office",
    Department: "Group Corporate Affairs",
    Unit: "Group Communications",
    "Document name": "board pack",
    "Project name": "",
    "Vendor/Customer": "Acme",
    "Document date": "2026-09-03",
    Confidentiality: "Highly Confidential",
    Remark: "",
    "Legally privileged": "Yes",
  };

  it("puts the tiers on the folder card and the form's fields on the file card", () => {
    expect(snapshotFolderRows(snap).map((r) => r.label)).toEqual([
      "BusinessSegment", "Department", "Unit",
    ]);
    expect(snapshotFileRows(snap).map((r) => r.label)).toEqual([
      "Document name", "Vendor/Customer", "Document date", "Confidentiality", "Legally privileged",
    ]);
  });

  it("loses nothing and shows nothing twice", () => {
    // The whole reason the partition is by key rather than by a guess at what a tier looks like.
    const folder = snapshotFolderRows(snap).map((r) => r.label);
    const file = snapshotFileRows(snap).map((r) => r.label);
    const filled = Object.keys(snap).filter((k) => (snap as Record<string, string>)[k].trim().length > 0);
    expect(folder.concat(file).sort()).toEqual(filled.slice().sort());
    expect(folder.filter((l) => file.indexOf(l) > -1)).toEqual([]);
  });

  it("drops blanks, so an optional Project Name leaves no empty row", () => {
    expect(snapshotFileRows(snap).map((r) => r.label)).not.toContain("Project name");
    expect(snapshotFolderRows({ Unit: "  " })).toEqual([]);
  });

  it("keeps the tiers in the order they were written, which is chain order", () => {
    // Insertion order is what makes the card read like the upload form, top down.
    const rows = snapshotFolderRows({ Region: "Sarawak", EstateMill: "Lavang" });
    expect(rows.map((r) => r.label)).toEqual(["Region", "EstateMill"]);
  });

  it("files `Bulk import` as a document field, never as a folder tier", () => {
    // Bulk Upload writes it. Left out of the key list it would claim the destination folder had a
    // level called "Bulk import".
    expect(snapshotFolderRows({ Unit: "u", "Bulk import": "Yes" }).map((r) => r.label)).toEqual(["Unit"]);
    expect(snapshotFileRows({ "Bulk import": "Yes" }).map((r) => r.label)).toEqual(["Bulk import"]);
  });

  it("survives an absent snapshot — a record written before the feature carries none", () => {
    expect(snapshotFolderRows(undefined)).toEqual([]);
    expect(snapshotFileRows(undefined)).toEqual([]);
  });

  it("names every field both upload screens write", () => {
    // A key missing here is silently filed as a folder tier, so the list is pinned rather than
    // trusted: these are the labels in `Form.tsx` and `BulkUpload.tsx`.
    expect(SNAPSHOT_FILE_KEYS).toEqual([
      "Document name", "Project name", "Vendor/Customer", "Document date",
      "Confidentiality", "Remark", "Legally privileged", "Bulk import",
    ]);
  });
});

/* ── The status line's record half ─────────────────────────────────────────────
 *
 * This line has been short twice: `cancelled` (2026-08-28) and `archived` (2026-09-03) were each
 * counted and then never mentioned, leaving a submission whose files were all in that state with a
 * blank Status cell. These pin that every state has a word and that none of them can be dropped.
 */
describe("recordStateParts: no state can go unmentioned", () => {
  const counts = (over: Partial<Record<string, number>>): {
    live: number; deleted: number; cancelled: number; archived: number; unknown: number;
  } => ({ live: 0, deleted: 0, cancelled: 0, archived: 0, unknown: 0, ...over });

  it("gives every non-live state a word", () => {
    // The guard against a fourth omission: a state added to RecordState fails to compile in
    // RECORD_STATE_LABEL, and this asserts the word is not left blank once it is added.
    for (const k of Object.keys(RECORD_STATE_LABEL)) {
      expect(RECORD_STATE_LABEL[k as keyof typeof RECORD_STATE_LABEL].length).toBeGreaterThan(0);
    }
    expect(Object.keys(RECORD_STATE_LABEL).sort()).toEqual(
      ["archived", "cancelled", "deleted", "unknown"],
    );
  });

  it("names each state when it has files, and stays silent when it does not", () => {
    for (const k of Object.keys(RECORD_STATE_LABEL)) {
      const parts = recordStateParts(counts({ [k]: 2 }));
      expect(parts).toHaveLength(1);
      expect(parts[0]).toBe(`2 ${RECORD_STATE_LABEL[k as keyof typeof RECORD_STATE_LABEL]}`);
    }
    expect(recordStateParts(counts({ live: 3 }))).toEqual([]);
  });

  it("says `replaced`, not `cancelled` — the file was superseded, not withdrawn", () => {
    expect(recordStateParts(counts({ cancelled: 1 }))).toEqual(["1 replaced"]);
  });

  it("says `not checked` for unknown, never anything implying the file is gone", () => {
    // It means a library read failed; those files may be perfectly fine.
    expect(recordStateParts(counts({ unknown: 1 }))).toEqual(["1 not checked"]);
  });

  it("reads worst news first and doubt last", () => {
    expect(recordStateParts(counts({ deleted: 1, cancelled: 2, archived: 3, unknown: 4 }))).toEqual([
      "1 deleted", "2 replaced", "3 archived", "4 not checked",
    ]);
  });
});

/* The `Archive` tab's rule. Separate from `filterByTab` because `archived` is a record state, not an
   approval outcome — see the function's own comment. */
describe("archivedRowsOnly", () => {
  it("keeps only the archived rows", () => {
    const rows = [
      rowFromRecord(rec({ fileId: "SFI-1" }), "archived"),
      rowFromRecord(rec({ fileId: "SFI-2" }), "deleted"),
      rowFromRecord(rec({ fileId: "SFI-3" }), "cancelled"),
      rowFromRecord(rec({ fileId: "SFI-4" }), "archived"),
    ];
    expect(archivedRowsOnly(rows).map((r) => r.record?.fileId)).toEqual(["SFI-1", "SFI-4"]);
  });

  it("keeps no LIVE row, whose recordState is undefined", () => {
    // A live document is not archived, and listing one here would offer no Delete or Share button
    // for a file that still has both.
    const live = { ...rowFromRecord(rec({ fileId: "SFI-1" }), "archived"), recordState: undefined };
    expect(archivedRowsOnly([live])).toEqual([]);
  });

  it("survives a missing list rather than throwing", () => {
    expect(archivedRowsOnly(undefined as never)).toEqual([]);
  });
});
