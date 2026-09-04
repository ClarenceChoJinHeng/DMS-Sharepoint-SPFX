/**
 * Tests for the batched-upload rules.
 *
 * Spec: docs/superpowers/specs/2026-08-15-batched-multi-file-upload-design.md
 *
 * The cases that matter are the ones where getting it wrong is SILENT: two files quietly overwriting
 * each other, a retry uploading a file twice, and an unreadable term store marking every batch as
 * needing attention. Those get the most coverage.
 */

import {
  Batch,
  StagedFile,
  applyUploadResults,
  batchesNeedingRepick,
  canSaveBatch,
  collisionsWithin,
  decideClash,
  duplicateAcrossBatches,
  inheritDefaults,
  nextAvailableName,
  nextId,
  resetIds,
  resolveUploadName,
  stagedTotals,
  summarise,
  uploadableBatches,
  WARN_FILE_COUNT,
} from "./uploadBatches";

/** A `File` stand-in — these rules only ever read `name` and `size`. */
const fakeFile = (name: string, size = 1024): File => ({ name, size }) as unknown as File;

const staged = (id: string, name: string, typedName = "", size = 1024): StagedFile => ({
  id,
  file: fakeFile(name, size),
  typedName,
  meta: {},
});

const batch = (id: string, files: StagedFile[], over: Partial<Batch> = {}): Batch => ({
  id,
  segmentKey: "mode_gho",
  chainSignature: "Department|Unit|Year|Document Type",
  pathLabels: ["GHO", "Group Finance", "Corporate", "2026", "Tax Return"],
  destination: {},
  files,
  ...over,
});

describe("nextAvailableName", () => {
  it("leaves a free name alone", () => {
    expect(nextAvailableName("TEST.pdf", [])).toBe("TEST.pdf");
    expect(nextAvailableName("TEST.pdf", ["OTHER.pdf"])).toBe("TEST.pdf");
  });

  it("appends the copy suffix before the extension", () => {
    expect(nextAvailableName("TEST.pdf", ["TEST.pdf"])).toBe("TEST - Copy.pdf");
  });

  it("walks past taken suffixes rather than appending blindly", () => {
    expect(nextAvailableName("TEST.pdf", ["TEST.pdf", "TEST - Copy.pdf"])).toBe(
      "TEST - Copy (2).pdf",
    );
    expect(
      nextAvailableName("TEST.pdf", ["TEST.pdf", "TEST - Copy.pdf", "TEST - Copy (2).pdf"]),
    ).toBe("TEST - Copy (3).pdf");
  });

  it("takes the first gap", () => {
    expect(nextAvailableName("TEST.pdf", ["TEST.pdf", "TEST - Copy (2).pdf"])).toBe(
      "TEST - Copy.pdf",
    );
  });

  it("compares case-insensitively, as SharePoint does", () => {
    expect(nextAvailableName("TEST.pdf", ["test.pdf"])).toBe("TEST - Copy.pdf");
    expect(nextAvailableName("TEST.pdf", ["TEST.pdf", "test - copy.PDF"])).toBe(
      "TEST - Copy (2).pdf",
    );
  });

  it("splits on the LAST dot", () => {
    expect(nextAvailableName("report.final.docx", ["report.final.docx"])).toBe(
      "report.final - Copy.docx",
    );
  });

  it("handles a name with no extension", () => {
    expect(nextAvailableName("README", ["README"])).toBe("README - Copy");
  });

  it("treats a leading dot as part of the stem, not an extension", () => {
    expect(nextAvailableName(".gitignore", [".gitignore"])).toBe(".gitignore - Copy");
  });

  it("returns blank for a blank request rather than inventing a name", () => {
    expect(nextAvailableName("", ["TEST.pdf"])).toBe("");
    expect(nextAvailableName("   ", [])).toBe("");
  });

  it("preserves the composed convention this form actually uploads under", () => {
    // Real shape: [Project] - [Vendor] - [Document Name] - [Date].pdf
    expect(
      nextAvailableName("TEST - TEST - TEST - 26-08-26.pdf", ["TEST - TEST - TEST - 26-08-26.pdf"]),
    ).toBe("TEST - TEST - TEST - 26-08-26 - Copy.pdf");
  });

  /* THE REGRESSION THIS SCHEME EXISTS FOR (client, 2026-08-27). 1.0.273.0 stripped a trailing ` (n)`
     and resumed counting, which merged the numbering space of `X` and `X (2)` and offered BOTH files
     in one run the same name. Appending cannot do that: the suffix goes on the whole original. */
  it("never offers two DIFFERENT originals the same name", () => {
    const taken = ["TEST.pdf", "TEST (2).pdf"];
    const a = nextAvailableName("TEST (2).pdf", taken);
    const b = nextAvailableName("TEST.pdf", taken);
    expect(a).toBe("TEST (2) - Copy.pdf");
    expect(b).toBe("TEST - Copy.pdf");
    expect(a).not.toBe(b);
  });

  it("does not strip a trailing (n) — it is part of the name", () => {
    expect(nextAvailableName("TEST (2).pdf", ["TEST (2).pdf"])).toBe("TEST (2) - Copy.pdf");
    expect(
      nextAvailableName("TEST - TEST - TEST - 26-08-26 (2).pdf", [
        "TEST - TEST - TEST - 26-08-26 (2).pdf",
      ]),
    ).toBe("TEST - TEST - TEST - 26-08-26 (2) - Copy.pdf");
  });

  it("suffixes an already-copied name again rather than renumbering it", () => {
    expect(nextAvailableName("TEST - Copy.pdf", ["TEST - Copy.pdf"])).toBe(
      "TEST - Copy - Copy.pdf",
    );
  });
});

describe("resolveUploadName", () => {
  it("keeps the original filename when nothing is typed", () => {
    expect(resolveUploadName("Scan_001.pdf", "")).toBe("Scan_001.pdf");
    expect(resolveUploadName("Scan_001.pdf", "   ")).toBe("Scan_001.pdf");
  });

  it("replaces the stem and always preserves the extension", () => {
    expect(resolveUploadName("Scan_001.pdf", "Rewards B")).toBe("Rewards B.pdf");
  });

  it("does not double the extension when the user typed it", () => {
    expect(resolveUploadName("Scan_001.pdf", "Rewards B.pdf")).toBe("Rewards B.pdf");
    // Case-insensitively, because SharePoint file names are.
    expect(resolveUploadName("Scan_001.PDF", "Rewards B.pdf")).toBe("Rewards B.pdf");
  });

  it("strips illegal characters rather than substituting a separator", () => {
    // Substituting would INVENT a name; `sanitizeFolderSegment` made the same choice.
    expect(resolveUploadName("a.pdf", "Q1/Q2 report")).toBe("Q1Q2 report.pdf");
  });

  it("falls back to the original when the typed name is entirely illegal", () => {
    expect(resolveUploadName("Scan_001.pdf", "***")).toBe("Scan_001.pdf");
  });

  it("handles a file with no extension", () => {
    expect(resolveUploadName("README", "Notes")).toBe("Notes");
  });

  it("treats a leading dot as part of the name, not an extension", () => {
    expect(resolveUploadName(".gitignore", "")).toBe(".gitignore");
  });
});

describe("collisionsWithin", () => {
  it("finds nothing when every resolved name differs", () => {
    const b = batch("b1", [staged("f1", "a.pdf"), staged("f2", "b.pdf")]);
    expect(collisionsWithin(b)).toEqual([]);
  });

  it("flags BOTH files when two resolve to one name — which is wrong is not knowable", () => {
    const b = batch("b1", [
      staged("f1", "scan1.pdf", "Rewards B"),
      staged("f2", "scan2.pdf", "Rewards B"),
    ]);
    expect(collisionsWithin(b).sort()).toEqual(["f1", "f2"]);
  });

  it("compares case-insensitively", () => {
    const b = batch("b1", [
      staged("f1", "scan1.pdf", "rewards b"),
      staged("f2", "scan2.pdf", "REWARDS B"),
    ]);
    expect(collisionsWithin(b).sort()).toEqual(["f1", "f2"]);
  });

  it("catches a typed name colliding with an untouched original", () => {
    const b = batch("b1", [staged("f1", "Rewards B.pdf"), staged("f2", "scan2.pdf", "Rewards B")]);
    expect(collisionsWithin(b).sort()).toEqual(["f1", "f2"]);
  });

  it("returns every member of a three-way clash", () => {
    const b = batch("b1", [
      staged("f1", "a.pdf", "Same"),
      staged("f2", "b.pdf", "Same"),
      staged("f3", "c.pdf", "Same"),
    ]);
    expect(collisionsWithin(b).sort()).toEqual(["f1", "f2", "f3"]);
  });

  it("prefers finalName, because THAT is what reaches SharePoint", () => {
    // Different typed names, identical composed names — the case comparing typed names would miss.
    const a = { ...staged("f1", "scan1.pdf", "Rewards"), finalName: "P1 - V1 - 010826.pdf" };
    const b = { ...staged("f2", "scan2.pdf", "Bonuses"), finalName: "P1 - V1 - 010826.pdf" };
    expect(collisionsWithin(batch("b1", [a, b])).sort()).toEqual(["f1", "f2"]);
  });

  it("does not flag files whose typed names match but whose composed names differ", () => {
    const a = { ...staged("f1", "scan1.pdf", "Rewards"), finalName: "P1 - V1 - 010826.pdf" };
    const b = { ...staged("f2", "scan2.pdf", "Rewards"), finalName: "P2 - V1 - 010826.pdf" };
    expect(collisionsWithin(batch("b1", [a, b]))).toEqual([]);
  });

  it("falls back to the typed name when the form has not computed one", () => {
    const b = batch("b1", [staged("f1", "a.pdf", "Same"), staged("f2", "b.pdf", "Same")]);
    expect(collisionsWithin(b).sort()).toEqual(["f1", "f2"]);
  });

  it("does NOT flag the same name in different batches — that is the destination's problem", () => {
    const b1 = batch("b1", [staged("f1", "a.pdf", "Same")]);
    const b2 = batch("b2", [staged("f2", "b.pdf", "Same")]);
    expect(collisionsWithin(b1)).toEqual([]);
    expect(collisionsWithin(b2)).toEqual([]);
  });
});

describe("canSaveBatch", () => {
  it("refuses an empty batch", () => {
    expect(canSaveBatch(batch("b1", []))).toBe(false);
  });

  it("refuses a batch with no destination", () => {
    expect(canSaveBatch(batch("b1", [staged("f1", "a.pdf")], { pathLabels: [] }))).toBe(false);
  });

  it("refuses a batch with an internal name clash", () => {
    const b = batch("b1", [staged("f1", "a.pdf", "X"), staged("f2", "b.pdf", "X")]);
    expect(canSaveBatch(b)).toBe(false);
  });

  it("allows a well-formed batch", () => {
    expect(canSaveBatch(batch("b1", [staged("f1", "a.pdf")]))).toBe(true);
  });
});

describe("inheritDefaults", () => {
  it("returns nothing for the first file in a batch", () => {
    expect(inheritDefaults(batch("b1", []))).toEqual({});
  });

  it("copies the last file's metadata", () => {
    const f1 = { ...staged("f1", "a.pdf"), meta: { vendor: "Vendor A", project: "P1" } };
    const f2 = { ...staged("f2", "b.pdf"), meta: { vendor: "Vendor B", project: "P2" } };
    expect(inheritDefaults(batch("b1", [f1, f2]))).toEqual({ vendor: "Vendor B", project: "P2" });
  });

  it("returns a COPY, so editing the new file cannot mutate its sibling", () => {
    const f1 = { ...staged("f1", "a.pdf"), meta: { vendor: "Vendor A" } };
    const b = batch("b1", [f1]);
    const got = inheritDefaults(b);
    got.vendor = "Changed";
    expect(b.files[0].meta.vendor).toBe("Vendor A");
  });
});

describe("applyUploadResults", () => {
  it("removes what uploaded and keeps what failed, with its reason", () => {
    const b = batch("b1", [staged("f1", "a.pdf"), staged("f2", "b.pdf")]);
    const out = applyUploadResults([b], [
      { fileId: "f1", ok: true },
      { fileId: "f2", ok: false, error: "HTTP 403" },
    ]);
    expect(out).toHaveLength(1);
    expect(out[0].files.map((f) => f.id)).toEqual(["f2"]);
    expect(out[0].files[0].error).toBe("HTTP 403");
  });

  it("drops a batch once every file in it uploaded", () => {
    const b = batch("b1", [staged("f1", "a.pdf")]);
    expect(applyUploadResults([b], [{ fileId: "f1", ok: true }])).toEqual([]);
  });

  it("A RETRY CANNOT UPLOAD A FILE TWICE — success is what removes it", () => {
    const b = batch("b1", [staged("f1", "a.pdf"), staged("f2", "b.pdf")]);
    const afterRun1 = applyUploadResults([b], [
      { fileId: "f1", ok: true },
      { fileId: "f2", ok: false, error: "HTTP 500" },
    ]);
    // The retry can only see f2, so f1 is not in the second run at all.
    const retryIds = afterRun1.reduce(
      (acc: string[], x) => [...acc, ...x.files.map((f) => f.id)],
      [],
    );
    expect(retryIds).toEqual(["f2"]);
    const afterRun2 = applyUploadResults(afterRun1, [{ fileId: "f2", ok: true }]);
    expect(afterRun2).toEqual([]);
  });

  it("leaves an unattempted file untouched — a stopped run is not a failure", () => {
    const b = batch("b1", [staged("f1", "a.pdf"), staged("f2", "b.pdf")]);
    const out = applyUploadResults([b], [{ fileId: "f1", ok: false, error: "HTTP 403" }]);
    expect(out[0].files.map((f) => f.id)).toEqual(["f1", "f2"]);
    expect(out[0].files[1].error).toBeUndefined();
  });

  it("clears a stale error from an earlier run", () => {
    const withError = { ...staged("f1", "a.pdf"), error: "HTTP 500" };
    const b = batch("b1", [withError, staged("f2", "b.pdf")]);
    const out = applyUploadResults([b], [{ fileId: "f2", ok: false, error: "HTTP 403" }]);
    expect(out[0].files[0].error).toBeUndefined();
  });

  it("supplies a reason when a failure arrived without one", () => {
    const b = batch("b1", [staged("f1", "a.pdf")]);
    const out = applyUploadResults([b], [{ fileId: "f1", ok: false }]);
    expect(out[0].files[0].error).toBe("Upload failed");
  });

  it("survives an empty result set unchanged", () => {
    const b = batch("b1", [staged("f1", "a.pdf")]);
    expect(applyUploadResults([b], [])).toHaveLength(1);
  });

  it("spans batches, keeping only the failures", () => {
    const b1 = batch("b1", [staged("f1", "a.pdf"), staged("f2", "b.pdf")]);
    const b2 = batch("b2", [staged("f3", "c.pdf")]);
    const out = applyUploadResults([b1, b2], [
      { fileId: "f1", ok: true },
      { fileId: "f2", ok: true },
      { fileId: "f3", ok: false, error: "name taken" },
    ]);
    expect(out.map((b) => b.id)).toEqual(["b2"]);
  });
});

describe("summarise", () => {
  it("counts, never a verdict", () => {
    expect(summarise([{ fileId: "a", ok: true }, { fileId: "b", ok: false }]))
      .toEqual({ ok: 1, failed: 1, total: 2 });
  });

  it("handles an empty run", () => {
    expect(summarise([])).toEqual({ ok: 0, failed: 0, total: 0 });
  });
});

describe("stagedTotals", () => {
  it("adds up files and bytes across batches", () => {
    const b1 = batch("b1", [staged("f1", "a.pdf", "", 1000), staged("f2", "b.pdf", "", 2000)]);
    const b2 = batch("b2", [staged("f3", "c.pdf", "", 3000)]);
    const t = stagedTotals([b1, b2]);
    expect(t.files).toBe(3);
    expect(t.bytes).toBe(6000);
    expect(t.overCount).toBe(false);
    expect(t.overBytes).toBe(false);
  });

  it("warns past the file count, and the threshold is exclusive", () => {
    const files = [];
    for (let i = 0; i < WARN_FILE_COUNT; i++) files.push(staged(`f${i}`, `${i}.pdf`));
    expect(stagedTotals([batch("b1", files)]).overCount).toBe(false);
    expect(stagedTotals([batch("b2", [...files, staged("extra", "x.pdf")])]).overCount).toBe(true);
  });

  it("warns past the byte total", () => {
    const big = batch("b1", [staged("f1", "a.pdf", "", 201 * 1024 * 1024)]);
    expect(stagedTotals([big]).overBytes).toBe(true);
  });

  it("is zero for nothing staged", () => {
    expect(stagedTotals([])).toEqual({ files: 0, bytes: 0, overCount: false, overBytes: false });
  });
});

describe("duplicateAcrossBatches", () => {
  it("reports a file staged in two batches", () => {
    const b1 = batch("b1", [staged("f1", "Contract.pdf", "", 500)]);
    const b2 = batch("b2", [staged("f2", "Contract.pdf", "", 500)]);
    expect(duplicateAcrossBatches([b1, b2])).toEqual(["Contract.pdf"]);
  });

  it("does not confuse two different files that share a name", () => {
    const b1 = batch("b1", [staged("f1", "Contract.pdf", "", 500)]);
    const b2 = batch("b2", [staged("f2", "Contract.pdf", "", 900)]);
    expect(duplicateAcrossBatches([b1, b2])).toEqual([]);
  });

  it("reports nothing for distinct files", () => {
    const b1 = batch("b1", [staged("f1", "a.pdf")]);
    const b2 = batch("b2", [staged("f2", "b.pdf")]);
    expect(duplicateAcrossBatches([b1, b2])).toEqual([]);
  });
});

describe("batchesNeedingRepick", () => {
  const gho = batch("b1", [staged("f1", "a.pdf")], { segmentKey: "mode_gho", chainSignature: "A|B" });
  const nbpol = batch("b2", [staged("f2", "b.pdf")], { segmentKey: "mode_nbpol", chainSignature: "C|D" });

  it("marks only the batches on the segment that moved", () => {
    const out = batchesNeedingRepick([gho, nbpol], { mode_gho: "A|B|NEW", mode_nbpol: "C|D" });
    expect(out[0].needsRepick).toBe(true);
    expect(out[1].needsRepick).toBeFalsy();
  });

  it("marks nothing when every chain is unchanged", () => {
    const out = batchesNeedingRepick([gho, nbpol], { mode_gho: "A|B", mode_nbpol: "C|D" });
    expect(out.every((b) => !b.needsRepick)).toBe(true);
  });

  it("AN UNREADABLE CHAIN MARKS NOTHING — unknown is not changed", () => {
    // A transient error must never take the form out of service. Fail-open, as everywhere else here.
    const out = batchesNeedingRepick([gho, nbpol], {});
    expect(out.every((b) => !b.needsRepick)).toBe(true);
  });

  it("clears a mark once the batch is re-picked onto the current chain", () => {
    const marked = { ...gho, needsRepick: true };
    const out = batchesNeedingRepick([marked], { mode_gho: "A|B" });
    expect(out[0].needsRepick).toBe(false);
  });

  it("clears a mark when the chain becomes unreadable, rather than leaving it stuck", () => {
    const marked = { ...gho, needsRepick: true };
    expect(batchesNeedingRepick([marked], {})[0].needsRepick).toBe(false);
  });

  it("marks two batches that share one moved segment", () => {
    const alsoGho = batch("b3", [staged("f3", "c.pdf")], { segmentKey: "mode_gho", chainSignature: "A|B" });
    const out = batchesNeedingRepick([gho, alsoGho], { mode_gho: "A|B|NEW" });
    expect(out.map((b) => b.needsRepick)).toEqual([true, true]);
  });
});

describe("uploadableBatches", () => {
  it("holds back only what is waiting on a re-pick", () => {
    const ok = batch("b1", [staged("f1", "a.pdf")]);
    const blocked = batch("b2", [staged("f2", "b.pdf")], { needsRepick: true });
    expect(uploadableBatches([ok, blocked]).map((b) => b.id)).toEqual(["b1"]);
  });
});

describe("nextId", () => {
  it("never repeats within a session", () => {
    resetIds();
    const ids = [nextId("f"), nextId("f"), nextId("b")];
    expect(ids.length).toBe(3);
    expect(ids[0]).toBe("f_1");
    expect(ids[1]).toBe("f_2");
    expect(ids[2]).toBe("b_3");
  });
});

/**
 * The replacement decision (client, 2026-08-28).
 *
 * ⚠ THE POINT OF THESE TESTS IS THE ASYMMETRY BETWEEN THE TWO UPLOAD SCREENS. The upload form may
 * replace a pending draft; Bulk Upload may never. A future reader seeing the two screens differ will
 * reasonably assume it is a bug — these are what stop them "fixing" Bulk Upload into being able to
 * destroy somebody's unreviewed work.
 */
describe("decideClash", () => {
  const base = {
    inStaging: false,
    inApproved: false,
    replaceStaging: false,
    replaceApproved: false,
  };

  it("lets a free name straight through, with no overwrite", () => {
    expect(decideClash(base)).toEqual({ consented: true, overwrite: false });
  });

  it("refuses a staging clash until the uploader consents", () => {
    const d = decideClash({ ...base, inStaging: true });
    expect(d.where).toBe("staging");
    expect(d.consented).toBe(false);
    expect(d.overwrite).toBe(false);
  });

  it("overwrites a pending draft once the uploader consents", () => {
    const d = decideClash({ ...base, inStaging: true, replaceStaging: true });
    expect(d).toEqual({ where: "staging", consented: true, overwrite: true });
  });

  it("refuses an approved-side clash until the uploader consents", () => {
    const d = decideClash({ ...base, inApproved: true });
    expect(d.where).toBe("approved");
    expect(d.consented).toBe(false);
  });

  /* The heart of "send for approval as a replacement": it is an ORDINARY upload. The name is free in
     the approval library, so nothing is destroyed now — Auto-route replaces the filed copy later, and
     only if an approver approves. An overwrite here would destroy a filed record at upload time. */
  it("NEVER overwrites for an approved-side clash, even when consented", () => {
    const d = decideClash({ ...base, inApproved: true, replaceApproved: true });
    expect(d).toEqual({ where: "approved", consented: true, overwrite: false });
  });

  it("reports both libraries as one case", () => {
    const d = decideClash({ ...base, inStaging: true, inApproved: true });
    expect(d.where).toBe("both");
    expect(d.consented).toBe(false);
  });

  /* ⚠ `both` IS UNLOCKED BY THE STAGING CONSENT AND NOTHING ELSE. Replacing the pending draft is the
     single action that also causes the filed copy to be replaced at approval, so the dialog offers
     ONE button. If the approved consent could unlock it, that one button would come apart. */
  it("unlocks `both` with the staging consent, never the approved one", () => {
    const viaApproved = decideClash({
      ...base, inStaging: true, inApproved: true, replaceApproved: true,
    });
    expect(viaApproved.consented).toBe(false);
    expect(viaApproved.overwrite).toBe(false);

    const viaStaging = decideClash({
      ...base, inStaging: true, inApproved: true, replaceStaging: true,
    });
    expect(viaStaging).toEqual({ where: "both", consented: true, overwrite: true });
  });

  /* ── Bulk Upload ─────────────────────────────────────────────────────────── */

  it("never lets Bulk Upload overwrite a pending draft, however it consents", () => {
    for (const extra of [
      { replaceStaging: true },
      { replaceApproved: true },
      { replaceStaging: true, replaceApproved: true },
    ]) {
      const d = decideClash({
        ...base, inStaging: true, allowStagingReplace: false, ...extra,
      });
      expect(d.consented).toBe(false);
      expect(d.overwrite).toBe(false);
    }
  });

  it("never lets Bulk Upload past a `both` clash", () => {
    const d = decideClash({
      ...base, inStaging: true, inApproved: true, replaceStaging: true, allowStagingReplace: false,
    });
    expect(d.where).toBe("both");
    expect(d.consented).toBe(false);
    expect(d.overwrite).toBe(false);
  });

  /* Bulk Upload keeps the approved-side offer — that path overwrites nothing, so the reason to
     withhold it does not apply. Its clash table differs from the form's in exactly one row. */
  it("still lets Bulk Upload send an approved-side replacement", () => {
    const d = decideClash({
      ...base, inApproved: true, replaceApproved: true, allowStagingReplace: false,
    });
    expect(d).toEqual({ where: "approved", consented: true, overwrite: false });
  });

  it("treats a missing allowStagingReplace as permitted, so the form needs no flag", () => {
    const d = decideClash({ ...base, inStaging: true, replaceStaging: true });
    expect(d.overwrite).toBe(true);
  });

  /* An overwrite is only ever produced by a staging clash. Pinned as an invariant rather than a case,
     because it is the property that makes every other path non-destructive. */
  it("produces overwrite=true only when a pending draft is actually in the way", () => {
    for (const inStaging of [true, false]) {
      for (const inApproved of [true, false]) {
        for (const replaceStaging of [true, false]) {
          for (const replaceApproved of [true, false]) {
            const d = decideClash({ inStaging, inApproved, replaceStaging, replaceApproved });
            if (d.overwrite) {
              expect(inStaging).toBe(true);
              expect(replaceStaging).toBe(true);
            }
          }
        }
      }
    }
  });
});
