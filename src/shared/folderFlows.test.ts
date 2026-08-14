import {
  FLOWS,
  Flow,
  FlowFacts,
  FlowStep,
  firstIncompleteStep,
  flowById,
  isLocked,
  labelMatches,
  lockReason,
  nearMatches,
  normaliseLabel,
  remainingCount,
  stepState,
} from "./folderFlows";

function flow(id: string): Flow {
  const f = flowById(id);
  if (!f) throw new Error(`no flow ${id}`);
  return f;
}

function step(flowId: string, stepId: string): FlowStep {
  const s = flow(flowId).steps.filter((x) => x.id === stepId)[0];
  if (!s) throw new Error(`no step ${stepId} in ${flowId}`);
  return s;
}

describe("the five flows", () => {
  it("offers exactly the five agreed flows", () => {
    expect(FLOWS.map((f) => f.id)).toEqual([
      "newSegment", "addUnit", "structure", "rename", "retire",
    ]);
  });

  it("marks only Retire as destructive", () => {
    expect(FLOWS.filter((f) => f.tone === "destructive").map((f) => f.id)).toEqual(["retire"]);
  });

  it("asks for a segment on every flow except Add a new segment", () => {
    // Flow 1 creates the segment, so it cannot require one first.
    expect(flow("newSegment").needsSegment).toBe(false);
    for (const id of ["addUnit", "structure", "rename", "retire"]) {
      expect(flow(id).needsSegment).toBe(true);
    }
  });

  it("asks for a subject only where that makes the term-store step checkable", () => {
    expect(flow("addUnit").asksSubject).toBe("add");
    expect(flow("rename").asksSubject).toBe("rename");
    expect(flow("structure").asksSubject).toBeUndefined();
  });

  it("makes Add a department or unit the tail of Add a new segment", () => {
    // The everyday job is steps 2-5 of the big one. If these ever diverge, one of them is wrong.
    const big = flow("newSegment").steps.map((s) => s.id);
    const small = flow("addUnit").steps.map((s) => s.id);
    expect(small.slice(1)).toEqual(big.slice(2));
  });

  it("does NOT offer reconciliation in the structure flow", () => {
    // Reconciliation walks the term tree, not Levels. Offering it invites a run that does nothing.
    expect(flow("structure").steps.map((s) => s.id)).not.toContain("reconcile");
  });

  it("puts abbreviations before reconciliation everywhere both appear", () => {
    for (const f of FLOWS) {
      const ids = f.steps.map((s) => s.id);
      if (ids.indexOf("abbreviations") === -1 || ids.indexOf("reconcile") === -1) continue;
      expect(ids.indexOf("abbreviations")).toBeLessThan(ids.indexOf("reconcile"));
    }
  });

  it("puts creating groups before mapping them", () => {
    // Mapping a group that does not exist yet is the ordering bug the landing page already had.
    for (const f of FLOWS) {
      const ids = f.steps.map((s) => s.id);
      if (ids.indexOf("groups") === -1 || ids.indexOf("folderAccess") === -1) continue;
      expect(ids.indexOf("groups")).toBeLessThan(ids.indexOf("folderAccess"));
    }
  });

  it("gives every step a hint — a step with no explanation is the problem we started with", () => {
    for (const f of FLOWS) {
      for (const s of f.steps) {
        expect(s.hint.trim().length).toBeGreaterThan(0);
        expect(s.label.trim().length).toBeGreaterThan(0);
      }
    }
  });
});

describe("isLocked — THE client's rule: never stop them doing the work", () => {
  it("locks reconciliation when terms are KNOWN to lack a code", () => {
    // The silent failure this whole flow exists to prevent.
    expect(isLocked(step("addUnit", "reconcile"), { abbreviationsMissing: 2 })).toBe(true);
    expect(lockReason(step("addUnit", "reconcile"), { abbreviationsMissing: 2 })).toContain("no folder code");
  });

  it("does NOT lock reconciliation when the count is zero", () => {
    expect(isLocked(step("addUnit", "reconcile"), { abbreviationsMissing: 0 })).toBe(false);
  });

  it("does NOT lock reconciliation when the count is UNKNOWN", () => {
    // The picker's state: the tree has not been walked. Unknown is not zero and is not a lock.
    expect(isLocked(step("addUnit", "reconcile"), {})).toBe(false);
    expect(isLocked(step("addUnit", "reconcile"), { abbreviationsMissing: undefined })).toBe(false);
  });

  it("locks only on a KNOWN-false fact, never on a missing one", () => {
    const migrate = step("structure", "migrate");
    expect(isLocked(migrate, { pendingLevels: false })).toBe(true);
    expect(isLocked(migrate, { pendingLevels: true })).toBe(false);
    // A read that failed leaves it undefined — must pass.
    expect(isLocked(migrate, {})).toBe(false);
  });

  it("locks the later steps of Add a new segment until the segment exists", () => {
    const abbrev = step("newSegment", "abbreviations");
    expect(isLocked(abbrev, { segmentExists: false })).toBe(true);
    expect(isLocked(abbrev, { segmentExists: true })).toBe(false);
    expect(isLocked(abbrev, {})).toBe(false);
  });

  it("does NOT lock abbreviations in the everyday flow — the segment is already there", () => {
    expect(step("addUnit", "abbreviations").lock).toBeUndefined();
  });

  it("locks NOTHING on a completely empty facts object", () => {
    // The state when every read failed. The flow must be fully walkable.
    for (const f of FLOWS) {
      for (const s of f.steps) {
        expect(isLocked(s, {})).toBe(false);
      }
    }
  });

  it("survives a missing facts object rather than throwing", () => {
    expect(isLocked(step("addUnit", "reconcile"), undefined as unknown as FlowFacts)).toBe(false);
  });

  it("locks only the agreed steps, and nothing else", () => {
    const locked: string[] = [];
    for (const f of FLOWS) {
      for (const s of f.steps) if (s.lock) locked.push(`${f.id}.${s.id}`);
    }
    expect(locked).toEqual([
      "newSegment.abbreviations", "newSegment.reconcile",
      "addUnit.reconcile", "structure.migrate", "rename.reconcile", "retire.delete",
    ]);
  });

  it("gives every lock a reason — a locked step with no explanation is a dead end", () => {
    for (const f of FLOWS) {
      for (const s of f.steps) {
        if (s.lock) expect(s.lock.reason.trim().length).toBeGreaterThan(0);
      }
    }
  });
});

describe("stepState", () => {
  it("is UNKNOWN, not todo, when nothing has been read", () => {
    // Rendered as "not checked". Calling it outstanding would nag about work already done.
    for (const f of FLOWS) {
      for (const s of f.steps) {
        expect(stepState(s, {})).toBe("unknown");
      }
    }
  });

  it("reads the term-set step off the segment existing", () => {
    // SegmentCreator refuses a set of the wrong depth, so a segment row proves the set was there.
    expect(stepState(step("newSegment", "termSet"), { segmentExists: true })).toBe("done");
    expect(stepState(step("newSegment", "termSet"), { segmentExists: false })).toBe("unknown");
  });

  it("marks the term-store step done ONLY because the flow asked for a subject", () => {
    expect(stepState(step("addUnit", "addTerm"), { subjectFound: true })).toBe("done");
    expect(stepState(step("addUnit", "addTerm"), { subjectFound: false })).toBe("todo");
    expect(stepState(step("addUnit", "addTerm"), {})).toBe("unknown");
  });

  it("marks abbreviations done only at zero missing, and unknown when unwalked", () => {
    expect(stepState(step("addUnit", "abbreviations"), { abbreviationsMissing: 0 })).toBe("done");
    expect(stepState(step("addUnit", "abbreviations"), { abbreviationsMissing: 1 })).toBe("todo");
    expect(stepState(step("addUnit", "abbreviations"), {})).toBe("unknown");
  });

  it("leaves migration and the Power Automate step permanently unknown", () => {
    // A migration leaves no marker, and Power Automate is unreachable from here.
    expect(stepState(step("structure", "migrate"), { pendingLevels: true })).toBe("unknown");
    expect(stepState(step("structure", "pauseFlows"), { segmentExists: true })).toBe("unknown");
  });
});

describe("firstIncompleteStep — resuming", () => {
  const facts: FlowFacts = { segmentExists: true, subjectFound: true, abbreviationsMissing: 3 };

  it("opens on the first step that is not done", () => {
    // Term added, codes outstanding → land on abbreviations.
    const f = flow("addUnit");
    expect(f.steps[firstIncompleteStep(f, facts)].id).toBe("abbreviations");
  });

  it("opens at the start when nothing is known", () => {
    expect(firstIncompleteStep(flow("addUnit"), {})).toBe(0);
  });

  it("ends on the LAST step when everything is done, not back at the beginning", () => {
    const f = flow("addUnit");
    const all: FlowFacts = {
      segmentExists: true, subjectFound: true, abbreviationsMissing: 0,
      groupsExist: true, folderAccessRows: true, foldersExist: true,
    };
    expect(firstIncompleteStep(f, all)).toBe(f.steps.length - 1);
  });

  it("never returns an out-of-range index, even for a flow with no steps", () => {
    const empty: Flow = {
      id: "x", label: "x", blurb: "x", tone: "normal", needsSegment: false, steps: [],
    };
    expect(firstIncompleteStep(empty, {})).toBe(0);
  });
});

describe("remainingCount", () => {
  it("counts everything not done, so unknown steps still read as outstanding work", () => {
    expect(remainingCount(flow("addUnit"), {})).toBe(flow("addUnit").steps.length);
  });

  it("reaches zero only when every step is done", () => {
    const all: FlowFacts = {
      segmentExists: true, subjectFound: true, abbreviationsMissing: 0,
      groupsExist: true, folderAccessRows: true, foldersExist: true,
    };
    expect(remainingCount(flow("addUnit"), all)).toBe(0);
  });
});

describe("normaliseLabel — the fullwidth ampersand trap", () => {
  it("matches a FULLWIDTH ＆ against a typed plain &", () => {
    // GHO really contains "Group Legal, Risk ＆ Compliance" — SharePoint requires the fullwidth form.
    // Without this, an admin who types & is told they never added the term.
    expect(normaliseLabel("Group Legal, Risk ＆ Compliance"))
      .toBe(normaliseLabel("Group Legal, Risk & Compliance"));
  });

  it("folds case and trims", () => {
    expect(normaliseLabel("  TREASURY  ")).toBe("treasury");
  });

  it("collapses runs of whitespace", () => {
    expect(normaliseLabel("Group   Finance")).toBe("group finance");
  });

  it("strips zero-width characters, which arrive by copy-paste", () => {
    expect(normaliseLabel("Trea​sury")).toBe("treasury");
  });

  it("survives blank input", () => {
    expect(normaliseLabel("")).toBe("");
    expect(normaliseLabel(undefined as unknown as string)).toBe("");
  });
});

describe("labelMatches", () => {
  const TREE = ["Group Finance", "Group Legal, Risk ＆ Compliance", "Treasury"];

  it("finds an exact term", () => {
    expect(labelMatches(TREE, "Treasury")).toBe(true);
  });

  it("finds one typed with a plain ampersand", () => {
    expect(labelMatches(TREE, "Group Legal, Risk & Compliance")).toBe(true);
  });

  it("does not match a term that is not there", () => {
    expect(labelMatches(TREE, "Tax")).toBe(false);
  });

  it("never matches on blank input — that would mark the step done for typing nothing", () => {
    expect(labelMatches(TREE, "")).toBe(false);
    expect(labelMatches(TREE, "   ")).toBe(false);
  });

  it("survives a missing tree", () => {
    expect(labelMatches(undefined as unknown as string[], "Treasury")).toBe(false);
  });
});

describe("nearMatches — for 'check the spelling'", () => {
  const TREE = ["Group Finance", "Group Finance Reporting", "Treasury"];

  it("suggests a containing label for a partial entry", () => {
    expect(nearMatches(TREE, "Group Finance Rep")).toContain("Group Finance Reporting");
  });

  it("suggests a label the entry contains", () => {
    expect(nearMatches(TREE, "Group Finance Reporting Unit")).toContain("Group Finance Reporting");
  });

  it("excludes the exact match — that is not a near miss", () => {
    expect(nearMatches(TREE, "Treasury")).not.toContain("Treasury");
  });

  it("says nothing for input too short to be meaningful", () => {
    expect(nearMatches(TREE, "G")).toEqual([]);
    expect(nearMatches(TREE, "")).toEqual([]);
  });

  it("caps the list, so the hint stays a hint", () => {
    const many = ["Group A", "Group B", "Group C", "Group D", "Group E"];
    expect(nearMatches(many, "Group", 3).length).toBeLessThanOrEqual(3);
  });
});
