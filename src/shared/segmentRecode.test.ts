import {
  canOfferRecode,
  folderRecodeConflict,
  folderRecodeIsNoOp,
  recodeRefusalReason,
  recodeSummary,
} from "./segmentRecode";
import { ExistingSegment } from "./newSegment";
import { SegmentCounts, unknownCounts } from "./segmentDeletion";

const counted = (folders: number, documents: number): SegmentCounts => ({
  state: "counted",
  folders,
  documents,
  groupMapRows: 0,
  abbreviationRows: 0,
});

describe("canOfferRecode", () => {
  it("allows an empty, countable segment — folders may exist, documents must not", () => {
    expect(canOfferRecode(counted(0, 0))).toBe(true);
    // An empty tree reconciliation has already built is the ordinary case, not a red flag.
    expect(canOfferRecode(counted(12, 0))).toBe(true);
  });

  it("refuses a segment that holds even one document", () => {
    expect(canOfferRecode(counted(3, 1))).toBe(false);
  });

  it("fails CLOSED on an unreadable count — same direction as canOfferFolderDelete", () => {
    expect(canOfferRecode(unknownCounts("the folders could not be read"))).toBe(false);
  });
});

describe("recodeRefusalReason", () => {
  it("is blank exactly when canOfferRecode is true", () => {
    expect(recodeRefusalReason(counted(5, 0))).toBe("");
  });

  it("names the document count and the real route when documents exist", () => {
    const msg = recodeRefusalReason(counted(4, 7));
    expect(msg).toContain("7 documents");
    expect(msg).toContain("Move or archive");
  });

  it("uses singular wording for exactly one document", () => {
    expect(recodeRefusalReason(counted(1, 1))).toContain("1 document,");
  });

  it("names the read failure, never guesses empty, on an unknown count", () => {
    const msg = recodeRefusalReason(unknownCounts("the group map list could not be read"));
    expect(msg).toContain("could not be read");
    expect(msg).toContain("Re-coding is refused");
  });
});

describe("folderRecodeConflict", () => {
  const existing: ExistingSegment[] = [
    { key: "mode_gho", label: "Group Head Office", stagingFolder: "GHO" },
    { key: "mode_mho", label: "Minamas Head Office", stagingFolder: "MHO" },
  ];

  it("blocks a new folder that collides with ANOTHER segment, case-insensitively", () => {
    const msg = folderRecodeConflict("mho", existing, "mode_gho");
    expect(msg).toContain("Minamas Head Office");
  });

  it("does not treat the segment's OWN current folder as a clash with itself", () => {
    expect(folderRecodeConflict("GHO", existing, "mode_gho")).toBe("");
  });

  it("says nothing about a genuinely free name", () => {
    expect(folderRecodeConflict("NEWCODE", existing, "mode_gho")).toBe("");
  });

  it("says nothing about a blank draft — that is a required-field concern elsewhere", () => {
    expect(folderRecodeConflict("   ", existing, "mode_gho")).toBe("");
  });
});

describe("folderRecodeIsNoOp", () => {
  it("is true for the identical value", () => {
    expect(folderRecodeIsNoOp("GHO", "GHO")).toBe(true);
  });

  it("is true for a value that only differs by case or whitespace", () => {
    expect(folderRecodeIsNoOp("  gho ", "GHO")).toBe(true);
  });

  it("is false once the sanitized value genuinely differs", () => {
    expect(folderRecodeIsNoOp("GHONEW", "GHO")).toBe(false);
  });
});

describe("recodeSummary", () => {
  it("names both folders and the folder count, and tells the admin to reconcile next", () => {
    const msg = recodeSummary("GHO", "GHONEW", counted(3, 0));
    expect(msg).toContain('"GHO"');
    expect(msg).toContain('"GHONEW"');
    expect(msg).toContain("3 empty folders");
    expect(msg).toContain("Folder Reconciliation");
  });

  it("uses singular wording for exactly one folder", () => {
    expect(recodeSummary("A", "B", counted(1, 0))).toContain("1 empty folder ");
  });

  it("reports zero folders rather than throwing when the count was unknown", () => {
    expect(recodeSummary("A", "B", unknownCounts("x"))).toContain("0 empty folders");
  });
});
