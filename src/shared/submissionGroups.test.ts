import { Submission } from "./mySubmissions";
import { groupSubmissions, newReference, statusCounts } from "./submissionGroups";

const row = (over: Partial<Submission> = {}): Submission => ({
  itemId: 1,
  library: "ApprovalDocument",
  name: "a.pdf",
  fileRef: "/sites/x/ApprovalDocument/GHO/GF/TAX/a.pdf",
  status: "Pending",
  comment: "",
  ...over,
});

describe("groupSubmissions", () => {
  it("groups files sharing a SubmissionId, and their batches within it", () => {
    const out = groupSubmissions([
      row({ itemId: 1, submissionId: "SUB-1", batchId: "BAT-A" }),
      row({ itemId: 2, submissionId: "SUB-1", batchId: "BAT-A" }),
      row({ itemId: 3, submissionId: "SUB-1", batchId: "BAT-B" }),
    ]);
    expect(out.length).toBe(1);
    expect(out[0].files.length).toBe(3);
    expect(out[0].batches.length).toBe(2);
    expect(out[0].batches.map((b) => b.files.length).sort()).toEqual([1, 2]);
  });

  it("keeps separate submissions apart", () => {
    const out = groupSubmissions([
      row({ itemId: 1, submissionId: "SUB-1" }),
      row({ itemId: 2, submissionId: "SUB-2" }),
    ]);
    expect(out.length).toBe(2);
  });

  /* ⚠ REVERSED 2026-08-22 by the client, on seeing the first build: referenceless rows were one row
     each, which is the flat list this feature exists to replace. They are now grouped by folder and
     day and flagged `inferred`. The tests below pin both the grouping and the flag. */
  it("groups referenceless rows filed into one folder on one day, and flags them inferred", () => {
    const day = new Date(2026, 7, 20, 9, 0);
    const out = groupSubmissions([
      row({ itemId: 1, created: day }),
      row({ itemId: 2, created: new Date(2026, 7, 20, 16, 30) }),
      row({ itemId: 3, created: day }),
    ]);
    expect(out.length).toBe(1);
    expect(out[0].files.length).toBe(3);
    expect(out[0].reference).toBe("");
    expect(out[0].inferred).toBe(true);
    // One folder means one batch, not one per file.
    expect(out[0].batches.length).toBe(1);
  });

  it("keeps a different DAY apart — a unit files into the same folder every month", () => {
    const out = groupSubmissions([
      row({ itemId: 1, created: new Date(2026, 7, 20, 9, 0) }),
      row({ itemId: 2, created: new Date(2026, 7, 21, 9, 0) }),
    ]);
    expect(out.length).toBe(2);
  });

  it("keeps a different FOLDER apart on the same day", () => {
    const at = new Date(2026, 7, 20, 9, 0);
    const out = groupSubmissions([
      row({ itemId: 1, created: at, fileRef: "/sites/x/ApprovalDocument/GHO/GF/TAX/a.pdf" }),
      row({ itemId: 2, created: at, fileRef: "/sites/x/ApprovalDocument/GHO/GF/CORU/b.pdf" }),
    ]);
    expect(out.length).toBe(2);
  });

  it("⚠ a row with NO date groups with nothing — unknown must never widen a group", () => {
    const out = groupSubmissions([row({ itemId: 1 }), row({ itemId: 2 })]);
    expect(out.length).toBe(2);
  });

  it("a stamped submission is never merged into an inferred one, or vice versa", () => {
    const at = new Date(2026, 7, 20, 9, 0);
    const out = groupSubmissions([
      row({ itemId: 1, created: at }),
      row({ itemId: 2, created: at, submissionId: "SUB-1" }),
    ]);
    expect(out.length).toBe(2);
    expect(out.filter((g) => g.inferred).length).toBe(1);
    expect(out.filter((g) => g.reference === "SUB-1")[0].inferred).toBe(false);
  });

  it("mixes old and new rows without either affecting the other", () => {
    const out = groupSubmissions([
      row({ itemId: 1, submissionId: "SUB-1" }),
      row({ itemId: 2 }),
      row({ itemId: 3, submissionId: "SUB-1" }),
    ]);
    expect(out.length).toBe(2);
    expect(out.filter((g) => g.reference === "SUB-1")[0].files.length).toBe(2);
  });

  it("treats whitespace as no reference rather than as a group key", () => {
    // Both fall to the inference; with no date they group with nothing, so they stay apart.
    const out = groupSubmissions([
      row({ itemId: 1, submissionId: "   " }),
      row({ itemId: 2, submissionId: "" }),
    ]);
    expect(out.length).toBe(2);
    expect(out.every((g) => g.inferred)).toBe(true);
  });

  it("dates a submission by its EARLIEST file", () => {
    const early = new Date("2026-08-20T09:00:00Z");
    const late = new Date("2026-08-20T11:00:00Z");
    const out = groupSubmissions([
      row({ itemId: 1, submissionId: "S", created: late }),
      row({ itemId: 2, submissionId: "S", created: early }),
    ]);
    expect(out[0].at).toEqual(early);
  });

  it("survives an empty or undefined list", () => {
    expect(groupSubmissions([])).toEqual([]);
    expect(groupSubmissions(undefined as unknown as Submission[])).toEqual([]);
  });
});

describe("statusCounts", () => {
  it("counts each state separately — a submission is not one status", () => {
    // Flattening a part-approved, part-rejected submission to one badge is how an uploader concludes
    // a rejected file was fine.
    const c = statusCounts([
      row({ status: "Approved" }), row({ status: "Approved" }),
      row({ status: "Pending" }), row({ status: "Rejected" }),
    ]);
    expect(c).toEqual({ Approved: 2, Pending: 1, Rejected: 1 });
  });

  it("is all zeroes for nothing", () => {
    expect(statusCounts([])).toEqual({ Approved: 0, Pending: 0, Rejected: 0 });
  });
});

describe("newReference", () => {
  it("reads as a reference a person could quote", () => {
    expect(newReference("SUB", new Date(2026, 7, 22), () => 0)).toBe("SUB-20260822-AAAA");
  });

  it("pads the month and day", () => {
    expect(newReference("BAT", new Date(2026, 0, 5), () => 0)).toBe("BAT-20260105-AAAA");
  });

  it("omits characters that are misread aloud or in handwriting", () => {
    // I, L, O, 0 and 1 are absent on purpose — this gets read out and written down.
    const ref = newReference("SUB", new Date(2026, 7, 22), () => 0.999999);
    const suffix = ref.split("-")[2];
    for (const bad of ["I", "L", "O", "0", "1"]) expect(suffix.indexOf(bad)).toBe(-1);
  });

  it("never runs off the end of the alphabet when rand returns 1", () => {
    expect(newReference("SUB", new Date(2026, 7, 22), () => 1)).toMatch(/^SUB-20260822-[A-Z2-9]{4}$/);
  });

  it("falls back to now on an invalid date rather than producing NaN in the reference", () => {
    expect(newReference("SUB", new Date("nonsense"), () => 0)).toMatch(/^SUB-\d{8}-AAAA$/);
  });
});
