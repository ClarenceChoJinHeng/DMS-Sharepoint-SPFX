import {
  SegmentCounts,
  canOfferFolderDelete,
  confirmationMatches,
  deletionSummary,
  needsTypedConfirmation,
  survivorLines,
  unknownCounts,
} from "./segmentDeletion";

/** A successful count. Tests override only the field under examination. */
function counted(over: Partial<SegmentCounts> = {}): SegmentCounts {
  return { state: "counted", folders: 0, documents: 0, groupMapRows: 0, ...over };
}

describe("canOfferFolderDelete — the one rule that fails CLOSED", () => {
  it("offers the option when the count succeeded, even at zero", () => {
    expect(canOfferFolderDelete(counted())).toBe(true);
    expect(canOfferFolderDelete(counted({ folders: 12, documents: 340 }))).toBe(true);
  });

  it("WITHHOLDS it when the count failed", () => {
    // The inverse of every other unreadable-data rule in this codebase, deliberately: offering it
    // here means deleting a tree the screen could not confirm was empty.
    expect(canOfferFolderDelete(unknownCounts("HTTP 500"))).toBe(false);
  });

  it("does not confuse an empty segment with an uncountable one", () => {
    const empty = counted();
    const unreadable = unknownCounts("timeout");
    // Both report zero documents. Only one of them means there are none.
    expect(empty.documents).toBe(unreadable.documents);
    expect(canOfferFolderDelete(empty)).not.toBe(canOfferFolderDelete(unreadable));
  });

  it("carries a reason when it withholds, so the screen can say why", () => {
    expect(unknownCounts("the library returned HTTP 403").reason).toContain("403");
  });
});

describe("needsTypedConfirmation", () => {
  it("does NOT gate a brand-new empty segment", () => {
    // A gate that fires on the harmless case is one people learn to type through without reading.
    expect(needsTypedConfirmation(counted(), false)).toBe(false);
  });

  it("gates whenever documents exist", () => {
    expect(needsTypedConfirmation(counted({ documents: 1 }), false)).toBe(true);
  });

  it("gates whenever folder-access mappings exist", () => {
    // Deleting these changes who can reach what, even though no file moves.
    expect(needsTypedConfirmation(counted({ groupMapRows: 1 }), false)).toBe(true);
  });

  it("always gates a folder delete, even of an empty segment", () => {
    expect(needsTypedConfirmation(counted(), true)).toBe(true);
  });

  it("always gates when the count failed, because nothing can be ruled out", () => {
    expect(needsTypedConfirmation(unknownCounts("timeout"), false)).toBe(true);
  });
});

describe("confirmationMatches", () => {
  it("accepts the label", () => {
    expect(confirmationMatches("Industrial", "Industrial")).toBe(true);
  });

  it("ignores case and surrounding whitespace", () => {
    // A long label would otherwise teach people to copy and paste it, defeating the check.
    expect(confirmationMatches("  upstream operations malaysia  ", "Upstream Operations Malaysia")).toBe(true);
  });

  it("rejects a different segment's name", () => {
    expect(confirmationMatches("Industrial", "Group Head Office")).toBe(false);
  });

  it("rejects a near miss rather than being generous", () => {
    expect(confirmationMatches("Industria", "Industrial")).toBe(false);
    expect(confirmationMatches("Industrial ops", "Industrial")).toBe(false);
  });

  it("NEVER passes on blank, in either direction", () => {
    // A mode row with an empty ModeLabel must not be deletable by typing nothing.
    expect(confirmationMatches("", "Industrial")).toBe(false);
    expect(confirmationMatches("   ", "Industrial")).toBe(false);
    expect(confirmationMatches("", "")).toBe(false);
    expect(confirmationMatches("anything", "")).toBe(false);
  });
});

describe("survivorLines", () => {
  it("always states that the columns and abbreviations stay", () => {
    const lines = survivorLines(false).join(" ");
    expect(lines).toContain("columns");
    expect(lines).toContain("abbreviations");
  });

  it("says the files stay put when the folders are NOT being deleted", () => {
    expect(survivorLines(false).join(" ")).toContain("stays exactly where it is");
  });

  it("names the recycle bin when they ARE — that is what makes the option acceptable", () => {
    expect(survivorLines(true).join(" ")).toContain("recycle bin");
  });

  it("keeps the columns and abbreviations promise even when the folders go", () => {
    // The documents may have been MOVED elsewhere, which is exactly this workflow — so their
    // metadata columns must survive the folders.
    const lines = survivorLines(true).join(" ");
    expect(lines).toContain("columns");
    expect(lines).toContain("abbreviations");
  });
});

describe("deletionSummary", () => {
  it("never claims nothing is being deleted", () => {
    expect(deletionSummary(counted(), false)).toBe("Deletes the segment only.");
  });

  it("counts the mappings", () => {
    expect(deletionSummary(counted({ groupMapRows: 1 }), false)).toBe(
      "Deletes the segment and 1 folder-access mapping.",
    );
    expect(deletionSummary(counted({ groupMapRows: 3 }), false)).toContain("3 folder-access mappings");
  });

  it("counts folders and documents only when the folders are being deleted", () => {
    const c = counted({ folders: 4, documents: 9, groupMapRows: 2 });
    expect(deletionSummary(c, false)).toBe("Deletes the segment and 2 folder-access mappings.");
    const withFolders = deletionSummary(c, true);
    expect(withFolders).toContain("4 folders");
    expect(withFolders).toContain("9 documents");
  });

  it("omits a zero document count rather than saying '0 documents'", () => {
    expect(deletionSummary(counted({ folders: 2 }), true)).toBe("Deletes the segment and 2 folders.");
  });

  it("never quotes folder counts it could not read", () => {
    expect(deletionSummary(unknownCounts("timeout"), true)).toBe("Deletes the segment only.");
  });
});
