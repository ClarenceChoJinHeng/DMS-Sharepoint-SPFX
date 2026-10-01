import {
  SegmentCounts,
  canDeleteArchive,
  canOfferFolderDelete,
  confirmationMatches,
  deletionSummary,
  needsTypedConfirmation,
  survivorLines,
  unknownCounts,
} from "./segmentDeletion";

/** A successful count. Tests override only the field under examination. */
function counted(over: Partial<SegmentCounts> = {}): SegmentCounts {
  return { state: "counted", folders: 0, documents: 0, groupMapRows: 0, abbreviationRows: 0, ...over };
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

describe("canDeleteArchive — a SEPARATE fail-closed rule from canOfferFolderDelete", () => {
  it("permits it only on a confirmed, explicit zero", () => {
    expect(canDeleteArchive(counted({ archiveDocuments: 0 }))).toBe(true);
  });

  it("refuses when the archive has content", () => {
    expect(canDeleteArchive(counted({ archiveDocuments: 7 }))).toBe(false);
  });

  it("refuses when the archive could not be counted — never guess empty on a failed read", () => {
    expect(canDeleteArchive(counted({ archiveDocuments: undefined }))).toBe(false);
    expect(canDeleteArchive(counted())).toBe(false); // field simply absent
  });

  it("is independent of the operational-library count — one can be empty while the other is not", () => {
    // The whole reason this is a separate field: a segment with real archived records but an
    // already-empty operational tree must still offer the ordinary folder delete.
    const c = counted({ documents: 0, archiveDocuments: 42 });
    expect(canOfferFolderDelete(c)).toBe(true);
    expect(canDeleteArchive(c)).toBe(false);
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

  it("gates whenever abbreviation rows exist — they are deleted unconditionally now", () => {
    expect(needsTypedConfirmation(counted({ abbreviationRows: 1 }), false)).toBe(true);
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
  it("states that the tier columns stay AND that the abbreviations are removed", () => {
    // ⚠ REVERSED 2026-09-11: abbreviations used to be promised to survive. Retire now deletes them
    // unconditionally, so this line has to warn about that, not reassure against it.
    const lines = survivorLines(false).join(" ");
    expect(lines).toContain("columns");
    expect(lines).toContain("abbreviations");
    expect(lines).toContain("deleted");
    expect(lines).not.toContain("stay, so re-creating");
  });

  it("says the files stay put when the folders are NOT being deleted", () => {
    expect(survivorLines(false).join(" ")).toContain("stays exactly where it is");
  });

  it("names the recycle bin when they ARE — that is what makes the option acceptable", () => {
    expect(survivorLines(true).join(" ")).toContain("recycle bin");
  });

  // The archive line: said only when the folders are actually going AND this site has an archive.
  // Client's decision 2026-08-26 — 7-year retained records outlive the segment that produced them,
  // so the retire skips those two libraries when they hold anything, and has to SAY so, or an admin
  // who ticked "delete the folders" finds Archive/<SEG> still standing in the next reconciliation
  // log. `archiveDocuments: 5` here — an explicit POSITIVE count — is what selects this branch.
  //
  // ⚠ SIMPLIFIED 2026-09-13: the "if this segment is ever recreated..." consequence that used to be
  // tested here is GONE, deliberately — the archive code-reuse guard on segment CREATION (same day)
  // makes that scenario structurally impossible now, so warning about it is stale advice rather
  // than a safety net. `undefined` and a positive count now share the SAME plain sentence, since the
  // displayed text no longer needs to distinguish "could not confirm" from "confirmed has content" —
  // only `canDeleteArchive`'s underlying check still cares about that distinction.
  it("says the archive is not removed when it has content on an archive site", () => {
    const lines = survivorLines(true, true, 5).join(" ");
    expect(lines).toContain("not removed");
    expect(lines).toContain("only empty ones are removed automatically");
    // The old, now-obsolete consequence wording must not reappear.
    expect(lines).not.toContain("recreated");
    expect(lines).not.toContain("Move existing folders");
    expect(lines).not.toContain("recognised");
  });

  it("says the archive is removed too when it is confirmed empty — never 'not removed'", () => {
    const lines = survivorLines(true, true, 0).join(" ");
    expect(lines).toContain("empty too");
    expect(lines).toContain("removed as well");
    expect(lines).not.toContain("not removed");
  });

  it("gives the archive an unknown count the SAME wording as a confirmed positive one", () => {
    // Not a distinction the client asked to keep — both mean "left alone", and `canDeleteArchive`
    // is the only place that still needs to tell them apart.
    expect(survivorLines(true, true, undefined).join(" ")).toBe(survivorLines(true, true, 5).join(" "));
  });

  // Omitting the third argument entirely must behave exactly like passing `undefined` explicitly —
  // no caller should have to remember to pass it just to stay in the safe, fail-closed state.
  it("treats an omitted archiveDocuments the same as an explicit undefined", () => {
    expect(survivorLines(true, true).join(" ")).toBe(survivorLines(true, true, undefined).join(" "));
  });

  it("says nothing about an archive the site does not have", () => {
    expect(survivorLines(true, false).join(" ")).not.toContain("archive");
    // Omitted is the same as absent: no caller should have to pass `false` to stay quiet.
    expect(survivorLines(true).join(" ")).not.toContain("archive");
  });

  it("does not mention the archive when no folder is being deleted at all", () => {
    // Nothing is going, so naming what survives among the folders would be noise — and it would
    // imply the OTHER folders are not surviving, which is the opposite of the truth here.
    expect(survivorLines(false, true).join(" ")).not.toContain("archive");
  });

  it("keeps the columns promise even when the folders go, and still warns abbreviations are gone", () => {
    // The documents may have been MOVED elsewhere, which is exactly this workflow — so their
    // metadata columns must survive the folders. Abbreviations do NOT survive either way.
    const lines = survivorLines(true).join(" ");
    expect(lines).toContain("columns");
    expect(lines).toContain("abbreviations");
    expect(lines).toContain("deleted");
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

  it("counts abbreviation rows, unconditionally — they go whether or not the folders do", () => {
    expect(deletionSummary(counted({ abbreviationRows: 1 }), false)).toBe(
      "Deletes the segment and 1 abbreviation.",
    );
    expect(deletionSummary(counted({ abbreviationRows: 5 }), false)).toContain("5 abbreviations");
    expect(deletionSummary(counted({ abbreviationRows: 5 }), true)).toContain("5 abbreviations");
  });

  // Added 2026-09-12 alongside canDeleteArchive — the summary (and the audit row that reuses it)
  // must say the archive went too, whenever it genuinely did.
  it("mentions the archive only when it was CONFIRMED empty and folders are being deleted", () => {
    expect(deletionSummary(counted({ archiveDocuments: 0 }), true)).toContain("its empty archive");
    // Not when the folders aren't being deleted at all — nothing archive-related is happening.
    expect(deletionSummary(counted({ archiveDocuments: 0 }), false)).not.toContain("archive");
    // Not when the archive has content — it was correctly left alone, so claiming it here would be
    // an audit row describing a deletion that never happened.
    expect(deletionSummary(counted({ archiveDocuments: 4 }), true)).not.toContain("archive");
    // Not when it could not be counted — same reasoning, unknown must never be reported as done.
    expect(deletionSummary(counted({ archiveDocuments: undefined }), true)).not.toContain("archive");
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
