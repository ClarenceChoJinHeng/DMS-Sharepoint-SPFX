import {
  stepUsesSegment,
  FLOWS,
  Flow,
  FlowFacts,
  FlowStep,
  blocksNext,
  firstBlockedStepIndex,
  firstIncompleteStep,
  isStepReachable,
  flowById,
  isLocked,
  labelMatches,
  lockReason,
  nearMatches,
  normaliseLabel,
  remainingCount,
  scopeFactsToFlow,
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

/* ⚠ Reported by the client 2026-09-09: Next was live on the migrate step BEFORE the migration had
   been run at all, on a segment plainly reading CHANGE PENDING. This step is the only thing that
   applies a staged chain, and step 5 turns uploads back on — which is the loop that cost three days
   on GHO.
   ⚠ THE EXHAUSTIVE SWEEP CANNOT COVER THIS ONE. It walks every step with every fact FALSE, and
   `pendingLevels: false` means "applied", which correctly gates nothing. A gate that fires on TRUE
   needs its own test or nothing guards it. */
describe("the migrate step holds Next while a change is still staged", () => {
  const migrate = step("structure", "migrate");

  it("gates while a change is staged, and releases once it is applied", () => {
    expect(blocksNext(migrate, { pendingLevels: true })).toContain("has not been applied yet");
    expect(blocksNext(migrate, { pendingLevels: false })).toBe("");
  });

  it("never gates on an unknown answer", () => {
    // The PendingLevels column does not exist on a site where nothing has ever staged a change,
    // and gating on a read that failed would strand every admin in this flow.
    expect(blocksNext(migrate, {})).toBe("");
    expect(blocksNext(migrate, { pendingLevels: undefined })).toBe("");
  });

  it("gates no other step, since the fact is read once for the whole flow", () => {
    for (const s of flow("structure").steps) {
      if (s.id === "migrate") continue;
      expect(blocksNext(s, { pendingLevels: true })).not.toContain("has not been applied yet");
    }
  });
});

describe("the flows", () => {
  /**
   * `runRecon` was added 2026-08-20 when "All tools" left the picker: reconciliation was the only
   * screen with no flow of its own, and it is a job an admin genuinely needs to start on its own.
   */
  it("offers exactly the agreed flows, in the agreed order", () => {
    expect(FLOWS.map((f) => f.id)).toEqual([
      "newSegment", "addUnit", "structure", "rename", "runRecon", "retire",
    ]);
  });

  it("marks only Retire as destructive", () => {
    expect(FLOWS.filter((f) => f.tone === "destructive").map((f) => f.id)).toEqual(["retire"]);
  });

  it("asks for a segment only where a step actually spends one", () => {
    // Flow 1 creates the segment, so it cannot require one first.
    expect(flow("newSegment").needsSegment).toBe(false);
    /* ⚠ RETIRE JOINED IT ON 2026-09-09. Its one remaining screen LISTS every segment with a Delete
       beside each, so a picker in front asked for a value the step never spends and a switcher above
       offered to change an answer that changed nothing — the `stepUsesSegment` shape, applied to a
       whole flow. It only became true when "Move the documents out" left: that step DID consume a
       segment, because the migrator works on one. */
    expect(flow("retire").needsSegment).toBe(false);
    for (const id of ["addUnit", "structure", "rename"]) {
      expect(flow(id).needsSegment).toBe(true);
    }
  });

  it("asks for a subject only where that makes the term-store step checkable", () => {
    expect(flow("addUnit").asksSubject).toBe("add");
    expect(flow("rename").asksSubject).toBe("rename");
    expect(flow("structure").asksSubject).toBeUndefined();
  });

  it("makes Add a department or unit the tail of Add a new segment", () => {
    /* The everyday job is the tail of the big one. If these ever diverge, one of them is wrong.
       ⚠ THE OFFSET CHANGED WITH THE 2026-09-06 REDESIGN and the invariant got STRONGER: `addUnit`
       lost its instruction step, so it is now EXACTLY the tail rather than the tail plus one. */
    const big = flow("newSegment").steps.map((s) => s.id);
    const small = flow("addUnit").steps.map((s) => s.id);
    expect(small).toEqual(big.slice(2));
  });

  it("keeps Add a department or unit at two steps", () => {
    // Pinned because the client asked for exactly two (2026-09-06), and because a third would
    // silently re-introduce the instruction screen the redesign removed.
    expect(flow("addUnit").steps.map((s) => s.id)).toEqual(["abbreviations", "reconcile"]);
  });

  /* ⚠ RENAME GAINED A THIRD STEP ON 2026-09-09, AND IT CLOSED A LIVE HAZARD. Its blurb promises
     *"or by changing its short code"* — which reconciliation delivers for Department and Unit and
     CANNOT below Unit, because that walk stops at `permissionedDepth`. So changing a below-Unit
     abbreviation used to write the row, run a reconciliation that touched nothing, leave the folder
     under its old name, and send the next upload to the new one: TWO FOLDERS FOR ONE TERM, reported
     as success. The migration is the only thing that renames a below-Unit folder. */
  /* ⚠ THE RETIRE FLOW'S STEPS WERE NEVER PINNED, so removing one broke no test — found on
     2026-09-09 when "Move the documents out" came out and the suite stayed green. Every other flow
     has a list pinned; this one now does too. */
  it("leaves Retire with the delete step alone", () => {
    expect(flow("retire").steps.map((s) => s.id)).toEqual(["delete"]);
  });

  /* ⚠ "Move the documents out" MOUNTED THE MIGRATOR, which re-shapes folders WITHIN one segment —
     there is no cross-segment move tool anywhere in this system. So the step was named after a thing
     the screen under it could not do, and its hint WAS the mechanism. Removed on the client's
     instruction; the warning it carried moved into the flow's blurb and the delete step's hint, which
     is what made removing it safe rather than merely tidier. */
  it("keeps the do-it-by-hand warning somewhere an admin retiring a segment will read", () => {
    const f = flow("retire");
    const text = `${f.blurb} ${f.steps.map((s) => s.hint).join(" ")}`;
    expect(text).toContain("by hand");
    expect(f.steps.map((s) => s.id)).not.toContain("moveOut");
  });

  /* ⚠ GAINED `pauseUploads`/`resumeUploads` ON 2026-09-11, MIRRORING `structure` EXACTLY — client:
     "Hard gate it same as Change the folder structure because we dont want to confuse the client".
     Both ids are reused with no new wiring: `blocksNext`/`stepState` key on `step.id` alone. */
  it("gives Rename a migrate step, between the abbreviation and the reconciliation, bracketed by the pause", () => {
    expect(flow("rename").steps.map((s) => s.id)).toEqual([
      "pauseUploads", "abbreviations", "migrate", "reconcile", "resumeUploads",
    ]);
  });

  it("gates the Rename flow's pause step exactly like structure's", () => {
    expect(
      blocksNext(step("rename", "pauseUploads"), { uploadsPaused: false }),
    ).toContain("Uploads are still switched on");
    expect(blocksNext(step("rename", "pauseUploads"), { uploadsPaused: true })).toBe("");
  });

  it("gates the Rename flow's closing step exactly like structure's", () => {
    expect(
      blocksNext(step("rename", "resumeUploads"), { uploadsPaused: true }),
    ).toContain("Uploads are still switched off");
    expect(blocksNext(step("rename", "resumeUploads"), { uploadsPaused: false })).toBe("");
  });

  it("does not put a segment picker in front of either Rename pause step", () => {
    expect(stepUsesSegment(step("rename", "pauseUploads"))).toBe(false);
    expect(stepUsesSegment(step("rename", "resumeUploads"))).toBe(false);
  });

  it("keeps the abbreviations step id stable however the flow labels it", () => {
    /* ⚠ THE LABEL VARIES, THE ID MUST NOT. `NEXT_GATED_STEPS` keys on the id, and it is what stops
       an admin reaching reconciliation with terms that have no code — the one silent failure on this
       screen (recon skips the term, creates no folder, and reports success). */
    for (const f of ["addUnit", "rename", "newSegment"]) {
      expect(flow(f).steps.filter((s) => s.id === "abbreviations")).toHaveLength(1);
    }
    expect(step("addUnit", "abbreviations").label).toBe("Create Term Abbreviation");
    expect(step("rename", "abbreviations").label).toBe("Rename Term Abbreviation");
  });

  it("has no Power Automate step in the structure flow", () => {
    // Removed 2026-09-06 (five steps to four). Pinned so it is not quietly restored: what it said is
    // recorded at the removal site in folderFlows.ts and is a real operational caveat.
    expect(flow("structure").steps.map((s) => s.id)).not.toContain("pauseFlows");
    // Six: the four the client's rail shows, plus reconciliation (2026-09-06) and term
    // abbreviations (2026-09-09, so a coded level's terms cannot reach the migration uncoded).
    expect(flow("structure").steps).toHaveLength(6);
  });

  /* ⚠ THIS TEST ONCE ASSERTED THE OPPOSITE, AND THE REASON IT DID IS WORTH KEEPING.
     Reconciliation walks the TERM TREE, not `Levels`, so a change to the levels below Unit gives it
     nothing to do — offering it invited an hour-long run that changes nothing, which is why it was
     excluded by design. The client asked for it anyway on 2026-09-06, having been shown that: it
     goes in as its own step, immediately before uploads resume, on the reading that re-asserting
     permissions before letting people back in is worth a run regardless. */
  it("offers reconciliation in the structure flow, just before uploads resume", () => {
    const ids = flow("structure").steps.map((s) => s.id);
    expect(ids).toEqual([
      "pauseUploads",
      "levels",
      "abbreviations",
      "migrate",
      "reconcile",
      "resumeUploads",
    ]);
  });

  it("puts abbreviations before reconciliation everywhere both appear", () => {
    for (const f of FLOWS) {
      const ids = f.steps.map((s) => s.id);
      if (ids.indexOf("abbreviations") === -1 || ids.indexOf("reconcile") === -1) continue;
      expect(ids.indexOf("abbreviations")).toBeLessThan(ids.indexOf("reconcile"));
    }
  });

  // The `folderAccess` step was removed on 2026-08-23 — Group Management adds the people itself — so
  // "groups before mapping them" no longer describes two steps. Pinned in the only form that still
  // means something: no flow may offer a step whose screen is the retired component id.
  it("offers no folderAccess step, on any flow", () => {
    for (const f of FLOWS) {
      for (const st of f.steps) {
        expect(st.id).not.toBe("folderAccess");
        expect(JSON.stringify(st.screen)).not.toContain("folderAccess");
      }
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
    const abbrev = step("newSegment", "abbreviations");
    expect(isLocked(abbrev, { segmentExists: false })).toBe(true);
    expect(isLocked(abbrev, { segmentExists: true })).toBe(false);
    // A read that failed leaves it undefined — must pass.
    expect(isLocked(abbrev, {})).toBe(false);
  });

  /* ⚠ THE MIGRATE STEP HAS NO LOCK AT ALL SINCE 2026-09-09, and this pins that it stays that way.
     It used to lock on `pendingLevels === false`, on the premise that a migration is only needed
     when a chain change is staged. Two cases disprove it and both stage NOTHING: subunit terms added
     to a unit that had none, and a below-Unit abbreviation changed. The scan is the only thing that
     can answer, so the step renders and `blocksNext` requires the scan to have been RUN. */
  it("never locks the migrate step — the scan is what answers", () => {
    for (const id of ["structure", "rename"]) {
      expect(step(id, "migrate").lock).toBeUndefined();
      expect(isLocked(step(id, "migrate"), { pendingLevels: false })).toBe(false);
    }
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
    // `runRecon.reconcile` carries the same lock because it is the SAME step object — one definition,
    // not a lock-free copy. It can never fire there: the flow has no abbreviations screen, so the count
    // stays `undefined`, and unknown never gates.
    /* `structure.reconcile` joined the list on 2026-09-06 for the same reason `runRecon.reconcile`
       is on it: the SAME step object, lock included, rather than a lock-free copy. It cannot fire
       there either — the structure flow has no abbreviations screen, so the count stays `undefined`
       and unknown never gates. */
    /* ⚠ `retire.delete` LEFT THIS LIST ON 2026-09-09, WITH ITS SEGMENT PICKER. Its lock asked
       "is the segment you chose real", which is unanswerable once nothing is chosen — and
       meaningless when the screen below is a list of the segments that ARE real, each with its own
       Delete. It would also have been inert: no segment picked leaves `segmentExists` undefined,
       and unknown never locks. A lock that can never fire is worse than no lock, because the next
       reader trusts it. */
    expect(locked).toEqual([
      "newSegment.abbreviations", "newSegment.reconcile",
      "addUnit.reconcile", "structure.reconcile",
      "rename.reconcile", "runRecon.reconcile",
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

  /* ⚠ NO FLOW CONTAINS `addTerm` OR `renameTerm` SINCE 2026-09-06, so this is tested against a
     literal step rather than through a flow. The arms are kept in `stepState` because the steps
     could return; if they never do, the `subjectFound` fact and these arms go together. */
  it("marks a term-store step done ONLY because the flow asked for a subject", () => {
    const parked: FlowStep = {
      id: "addTerm",
      label: "Add the term",
      hint: "",
      screen: { kind: "outside" },
    };
    expect(stepState(parked, { subjectFound: true })).toBe("done");
    expect(stepState(parked, { subjectFound: false })).toBe("todo");
    expect(stepState(parked, {})).toBe("unknown");
  });

  it("marks abbreviations done only at zero missing, and unknown when unwalked", () => {
    expect(stepState(step("addUnit", "abbreviations"), { abbreviationsMissing: 0 })).toBe("done");
    expect(stepState(step("addUnit", "abbreviations"), { abbreviationsMissing: 1 })).toBe("todo");
    expect(stepState(step("addUnit", "abbreviations"), {})).toBe("unknown");
  });

  it("leaves migration permanently unknown", () => {
    // A migration leaves no marker anywhere, so nothing can report it as done.
    // (The Power Automate step this also covered was removed on 2026-09-06.)
    expect(stepState(step("structure", "migrate"), { pendingLevels: true })).toBe("unknown");
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
      groupsExist: true, foldersExist: true,
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
      groupsExist: true, foldersExist: true,
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

describe("blocksNext", () => {
  // The client, 2026-08-17: "won't it make more sense once client finish filling up the New Segment and
  // saved and only next step is available?" Before this, Next advanced from a step that had not been
  // done — and the padlock meant to mark it could not fire either, because `segmentExists` is only ever
  // computed for a CHOSEN segment and flow 1 has no picker.
  it("blocks Next on New segment until the segment exists", () => {
    const st = step("newSegment", "createSegment");
    expect(blocksNext(st, { segmentExists: false })).toContain("Create the segment first");
    expect(blocksNext(st, { segmentExists: true })).toBe("");
  });

  it("does NOT block while the fact is unknown — a failed read must never trap anyone", () => {
    // The rule that matters most here, and the same one `isLocked` follows. A config list that 500s or
    // throttles would otherwise strand an admin on step 2 with no way forward and no explanation.
    expect(blocksNext(step("newSegment", "createSegment"), {})).toBe("");
    expect(blocksNext(step("newSegment", "abbreviations"), {})).toBe("");
  });

  it("blocks Next while any term still has no folder code", () => {
    const st = step("newSegment", "abbreviations");
    expect(blocksNext(st, { abbreviationsMissing: 3 })).toContain("no folder code");
    expect(blocksNext(st, { abbreviationsMissing: 0 })).toBe("");
  });

  it("holds Next while the term store is STILL BEING READ", () => {
    // In flight is not unknown. The count is `undefined` for the whole read, and unknown never
    // gates — so Next was clickable on a step reading "Reading the term store…", which is a step
    // whose answer is seconds away and definitely not yet known. Reported on site 2026-08-18.
    //
    // Gating on it strands nobody, because it clears itself. That is exactly the split that made
    // `subjectGiven` right: fail-open exists for reads that can FAIL, not for reads still running.
    const st = step("newSegment", "abbreviations");
    expect(blocksNext(st, { abbreviationsLoading: true })).toContain("Still reading");
    expect(blocksNext(st, { abbreviationsLoading: true, abbreviationsMissing: 0 })).toContain("Still reading");
  });

  it("stops holding Next the moment the read FINISHES", () => {
    const st = step("newSegment", "abbreviations");
    expect(blocksNext(st, { abbreviationsLoading: false, abbreviationsMissing: 0 })).toBe("");
    // A read that finished and failed leaves the count unknown, and unknown must NOT gate — that is
    // the case a throttled list produces, and holding there traps someone with no way forward.
    expect(blocksNext(st, { abbreviationsLoading: false })).toBe("");
  });

  it("NEVER blocks on an advisory step, however definite the fact looks", () => {
    /* ⚠ `reconcile` STANDS IN FOR "AN ADVISORY STEP" NOW. These assertions named `groups`, which
       left both flows on 2026-09-06 — so the point they make needs a step that is still there.
       `groupsExist` is read by the NAMING CONVENTION, which suggestGroupName only suggests, so a
       hand-named group reads as absent; gating on it would stop an admin who had already done the
       work, which is the exact failure the fail-open design exists to avoid. */
    expect(blocksNext(step("newSegment", "reconcile"), { groupsExist: false })).toBe("");
    // Loading gates the abbreviation step ONLY. It is set for the whole flow, so any other step
    // reading it would be held every time an admin opened the abbreviations screen.
    expect(blocksNext(step("newSegment", "reconcile"), { abbreviationsLoading: true })).toBe("");
    expect(blocksNext(step("newSegment", "createSegment"), { abbreviationsLoading: true, segmentExists: true })).toBe("");
    // A subject that merely differs in spelling must cost help, never progress. Asserted on the
    // reconciliation step since the instruction step it used to name was removed on 2026-09-06.
    expect(blocksNext(step("addUnit", "reconcile"), { subjectFound: false })).toBe("");
    // Reconciliation is advisory too — it is the step an admin most often arrives having already
    // run, and holding it would trap exactly them.
    expect(blocksNext(step("newSegment", "reconcile"), { foldersExist: false })).toBe("");
  });

  it("gates exactly three step ids, across every flow — pinned so a new gate must be deliberate", () => {
    const gated: string[] = [];
    for (const f of FLOWS) {
      for (const st of f.steps) {
        /* Every fact false at once: anything gateable will gate.
           ⚠ A FACT MISSING FROM HERE MAKES THIS SWEEP BLIND TO ITS GATE. `uploadsPaused` was absent
           until 2026-09-07, so adding the pauseUploads gate would have left this test passing
           unchanged — the one test whose whole job is to make a new gate deliberate. Any field added
           to FlowFacts belongs in this object. */
        const all: FlowFacts = {
          segmentExists: false, groupsExist: false,
          foldersExist: false, abbreviationsMissing: 9, pendingLevels: false, subjectFound: false,
          uploadsPaused: false,
          /* ⚠ LISTED, AND `false` IS THE RIGHT POLARITY. For a LOADING flag the gating value is
             `true`, so this sweep is deliberately not the test that proves the loading gate — it
             asserts the step list is unchanged by it. `uploadsPaused: false` above already puts
             `pauseUploads` in the list, and the loading gate has its own describe block below,
             exactly as `abbreviationsLoading` does. What this line buys is that a future change to
             the gate's polarity shows up here rather than nowhere. */
          uploadsPausedLoading: false,
          /* Enforced at the client's request on 2026-09-08 even though the step is not technically
             required — see NEXT_GATED_STEPS.reconcile. */
          reconcileRan: false,
          /* ⚠ LISTED, AND DELIBERATELY `undefined`. The no-segment gate is the one rule here that is
             keyed on a STEP PROPERTY rather than a step id, so folding it into this "exactly N ids"
             pin would muddy both: the id list would fill with steps whose gate has nothing to do
             with the allow-list. It gets its own exhaustive sweep, immediately below. */
          segmentChosen: undefined,
        };
        if (blocksNext(st, all).length > 0 && gated.indexOf(st.id) === -1) gated.push(st.id);
      }
    }
    expect(gated.sort()).toEqual([
      "abbreviations", "createSegment", "pauseUploads", "reconcile",
    ]);
  });

  describe("the pause setting is still being read", () => {
    /* ⚠ THE DEFECT (client, 2026-09-09): *"When I load in there is a split second I can click the
       next button without disablling the upload, when the internet is bad I can definitely click it
       fast enough."* `uploadsPaused` is `undefined` for the whole read and `undefined` never gates,
       so Next was green beside a panel reading "Reading the current setting…".
       The same shape as `abbreviationsLoading`, in a second place: in flight is not unknown. */
    it("holds Next while the read is in flight", () => {
      expect(
        blocksNext(step("structure", "pauseUploads"), { uploadsPausedLoading: true }),
      ).toContain("Still reading");
      // And it outranks a stale value: an answer being replaced is not an answer.
      expect(
        blocksNext(step("structure", "pauseUploads"), {
          uploadsPausedLoading: true,
          uploadsPaused: true,
        }),
      ).toContain("Still reading");
    });

    /* ⚠ THE CLOSING STEP TOO. Finish is gated on `resumeUploads`, so an open gate there is an open
       Finish — and a forgotten pause leaves the site refusing uploads behind a banner that makes it
       look deliberate. Same setting, same window, same hold. */
    it("holds the closing step as well", () => {
      expect(
        blocksNext(step("structure", "resumeUploads"), { uploadsPausedLoading: true }),
      ).toContain("Still reading");
    });

    it("lets go the moment the read FINISHES, however it finished", () => {
      // Read, and uploads are off: the step is done and Next opens.
      expect(
        blocksNext(step("structure", "pauseUploads"), {
          uploadsPausedLoading: false,
          uploadsPaused: true,
        }),
      ).toBe("");
      /* ⚠ FINISHED AND FAILED MUST NOT GATE. That is what a throttled config list produces, and
         holding there would strand an admin with no way forward — the fail-open rule this whole
         module is built on. The flag is only ever about the request being in the air. */
      expect(blocksNext(step("structure", "pauseUploads"), { uploadsPausedLoading: false })).toBe("");
      expect(blocksNext(step("structure", "resumeUploads"), { uploadsPausedLoading: false })).toBe("");
    });

    /* It is set for the WHOLE flow, so every step sees it — the same pin `subjectGiven` needed. */
    it("does not leak onto any other step, in any flow", () => {
      for (const f of FLOWS) {
        for (const st of f.steps) {
          if (st.id === "pauseUploads" || st.id === "resumeUploads") continue;
          expect(blocksNext(st, { uploadsPausedLoading: true })).toBe("");
        }
      }
    });

    /* `undefined` is a flow that never reported — the standalone tabs, and any caller predating the
       fact. It must behave exactly as it did before this existed. */
    it("gates nothing when nobody reported", () => {
      expect(blocksNext(step("structure", "pauseUploads"), {})).toBe("");
      expect(blocksNext(step("structure", "resumeUploads"), {})).toBe("");
    });
  });

  describe("reconciliation not run", () => {
    it("holds the step until a run has finished", () => {
      const st = step("structure", "reconcile");
      expect(blocksNext(st, { reconcileRan: false })).toContain("has not been run yet");
      expect(blocksNext(st, { reconcileRan: true })).toBe("");
    });

    /* ⚠ `undefined` NEVER GATES, as everywhere else here — a flow whose reconciliation screen has
       never been mounted has not reported anything, and that is not the same as "not run". */
    it("gates nothing when the flow never reported", () => {
      expect(blocksNext(step("structure", "reconcile"), {})).toBe("");
    });

    /* It is the LAST step of four of the five flows that have it, which is why Finish had to be
       gated in the runner as well — Next alone would enforce it in `structure` only. */
    it("applies in every flow that has the step", () => {
      for (const id of ["newSegment", "addUnit", "rename", "runRecon", "structure"]) {
        expect(blocksNext(step(id, "reconcile"), { reconcileRan: false }).length).toBeGreaterThan(0);
      }
    });

    /* ⚠ AND IT MUST NOT LEAK ONTO THE STEP THAT TURNS UPLOADS BACK ON. In `structure`, reconcile
       sits immediately before `resumeUploads`; if this gate reached that step too, a segment whose
       reconciliation would not complete could never have its uploads switched back on. */
    it("does not hold any other step in the structure flow", () => {
      /* `levels` has no gate of its own, and `migrate`'s reads `pendingLevels`, which is unknown
         here — so anything either returns could only have come from the reconcile gate. */
      for (const id of ["levels", "migrate"]) {
        expect(blocksNext(step("structure", id), { reconcileRan: false })).toBe("");
      }
      /* ⚠ THE TWO PAUSE STEPS CANNOT BE CHECKED THE SAME WAY: they read `uploadsPaused` in opposite
         directions, so whichever value is passed, one of them gates for its OWN reason. What matters
         is that the reason is never the reconcile one. */
      for (const paused of [true, false]) {
        for (const id of ["pauseUploads", "resumeUploads"]) {
          expect(blocksNext(step("structure", id), { reconcileRan: false, uploadsPaused: paused }))
            .not.toContain("Folder Reconciliation");
        }
      }
    });
  });

  /* Reported by the client 2026-09-08, on the migrate step: Next was clickable with the picker still
     reading "Select a segment...". The picker replaces that step's whole content, so walking on
     reaches step 5 — which turns uploads back on over a chain that was never applied. */
  describe("no segment picked", () => {
    it("holds every step that spends a segment, and no other, across every flow", () => {
      for (const f of FLOWS) {
        for (const st of f.steps) {
          const blocked = blocksNext(st, { segmentChosen: false }).length > 0;
          // The rule IS `stepUsesSegment`, so the two must agree step for step. Asserting a list of
          // ids instead would let a step added later slip through with nobody noticing.
          expect(blocked).toBe(stepUsesSegment(st));
        }
      }
    });

    it("names the picker and never the rail, which no longer navigates forward", () => {
      const st = step("structure", "migrate");
      const msg = blocksNext(st, { segmentChosen: false });
      expect(msg).toContain("Choose a segment above");
      // The old escape hatch. True while the rail navigated freely, false since 2026-08-30 — and a
      // gate pointing at a control that does nothing reads as a broken page.
      expect(msg).not.toContain("list of steps");
    });

    it("lets go the moment one is picked", () => {
      expect(blocksNext(step("structure", "migrate"), { segmentChosen: true })).toBe("");
    });

    /* ⚠ THE FAIL-OPEN CASE, and the reason this is one optional boolean rather than two fields: an
       unreadable segment list leaves nothing to pick FROM, so gating on it would strand an admin
       with no way forward at all. */
    it("gates nothing when the segment list could not be read", () => {
      for (const f of FLOWS) {
        for (const st of f.steps) {
          expect(blocksNext(st, { segmentChosen: undefined })).toBe("");
        }
      }
    });

    /* It must not reach the site-wide pause or the Folder levels screen — neither spends a segment,
       and the picker deliberately does not stand in front of either. Both are `stepUsesSegment`
       exclusions, so this pins the two that were reported as bugs in their own right. */
    it("does not hold the site-wide upload steps or the Folder levels screen", () => {
      for (const id of ["pauseUploads", "resumeUploads"]) {
        expect(blocksNext(step("structure", id), { segmentChosen: false })).toBe("");
      }
      expect(blocksNext(step("structure", "levels"), { segmentChosen: false })).toBe("");
    });
  });

  /* GHO, 2026-09-04: the structure change was staged, the migration started at 16:36, and a file
     arrived at 16:37 — one minute in, into a folder the scan had already passed. Four people kept
     filing for three days, so the tidy never met an empty folder and the chain never applied.
     `uploadsPaused` read `no` throughout, and the flow let the admin walk past step 1 regardless. */
  it("holds Next while uploads are still on, and lets go the moment they are paused", () => {
    const st = step("structure", "pauseUploads");
    expect(blocksNext(st, { uploadsPaused: false })).toContain("still switched on");
    expect(blocksNext(st, { uploadsPaused: true })).toBe("");
  });

  it("does NOT hold Next when the pause state could not be read", () => {
    // Same rule as every other gate here: unknown is not a lock. An unreadable config list must never
    // trap an admin mid-flow, and this one cannot even be worked around from another screen.
    expect(blocksNext(step("structure", "pauseUploads"), {})).toBe("");
  });

  /* ⚠ THIS TEST USED TO PIN THE OPPOSITE — "never holds the closing Enable Upload step" — on the
     reasoning that its `todo` is the state an admin arrives in, so gating would trap them. Reversed
     2026-09-08 at the client's request, because both halves of that reason had become false: the Back
     band is an open exit, and the toggle satisfying the gate is on that very step. Meanwhile uploads
     sat off site-wide for a day and a half because this step was never reached. */
  it("holds the closing step until uploads are back on", () => {
    const st = step("structure", "resumeUploads");
    expect(blocksNext(st, { uploadsPaused: true })).toContain("still switched off");
    expect(blocksNext(st, { uploadsPaused: false })).toBe("");
  });

  /* ⚠ AN UNREADABLE SETTING MUST NOT HOLD IT. This is the closing step of the flow and the only
     control that turns uploads back on — holding it over a throttled config read would be the one
     failure worse than the one the gate prevents. */
  it("does not hold the closing step when the setting could not be read", () => {
    expect(blocksNext(step("structure", "resumeUploads"), {})).toBe("");
  });

  /* ⚠ ONE FACT, TWO OPPOSITE READINGS, so no single facts object can gate both pause steps — which is
     why the exhaustive sweep above cannot cover this one and it needs its own test. */
  it("gates the two pause steps in opposite directions", () => {
    expect(blocksNext(step("structure", "pauseUploads"), { uploadsPaused: true })).toBe("");
    expect(blocksNext(step("structure", "resumeUploads"), { uploadsPaused: true }).length)
      .toBeGreaterThan(0);
    expect(blocksNext(step("structure", "pauseUploads"), { uploadsPaused: false }).length)
      .toBeGreaterThan(0);
    expect(blocksNext(step("structure", "resumeUploads"), { uploadsPaused: false })).toBe("");
  });

  it("asks Add a new segment for the segment's NAME, which is what makes the gate answerable", () => {
    // A row COUNT was the obvious alternative and is wrong: a baseline taken when the flow opens resets
    // on a page refresh, so the flow would then refuse work already done. A name is derived from data.
    expect(flow("newSegment").asksSubject).toBe("newSegment");
  });

  it("matches the typed name the way labelMatches does, so a fullwidth ＆ still counts as created", () => {
    // The client's GHO really contains "Group Legal, Risk ＆ Compliance"; a segment label could too, and
    // an admin typing a plain & must not be told their segment does not exist.
    expect(labelMatches(["Group Legal, Risk ＆ Compliance"], "group legal, risk & compliance")).toBe(true);
  });
});

describe("blocksNext — the blank name", () => {
  const st = (): FlowStep => step("newSegment", "createSegment");

  it("blocks Next until the created segment is confirmed, when the list WAS readable", () => {
    // Reported twice by the client (2026-08-17): "Umm I still can click next, I haven't fill up New
    // Segment yet...". The first build treated "not answered" as UNKNOWN and therefore gated nothing.
    const msg = blocksNext(st(), { subjectGiven: false });
    expect(msg).toContain("Create the segment above");
    /* ⚠ IT MUST NAME A WAY OUT, because the gate can legitimately be wrong after a page refresh —
       the segment list is read at MOUNT and Create happens after it. Until 2026-08-30 the way out
       was the rail ("jump straight on from the list of steps"); the client then asked for the rail
       to be locked forward, which made that sentence FALSE. It now names Refresh list, which is the
       escape that still exists. **A gate pointing at a control that no longer works is worse than a
       gate with no advice**: they try it, nothing happens, and the page reads as broken. */
    expect(msg).toContain("Refresh list");
    expect(msg).not.toContain("list of steps");
  });

  it("still does NOT block when the segment list could not be read", () => {
    // The distinction the first build got wrong. Fail-open exists for READS THAT CAN FAIL, not for a text
    // box nobody filled in — a throttled list cannot blank a local field, so gating on the field strands
    // nobody, while gating on an unreadable list strands everybody.
    expect(blocksNext(st(), {})).toBe("");
  });

  it("says nothing once the segment exists, even with no name typed", () => {
    // Order matters: a page reopened after the work is done has no typed name, and must not then be told
    // to create a segment that already exists.
    expect(blocksNext(st(), { segmentExists: true, subjectGiven: false })).toBe("");
  });

  it("moves on to 'create it first' once a segment IS picked but does not exist", () => {
    // Reachable only from stale facts — picking a segment implies it exists. Kept because the two
    // messages must stay distinguishable: "confirm what you made" and "you have not made it" are
    // different instructions, and collapsing them into one is how a gate stops being actionable.
    const msg = blocksNext(st(), { subjectGiven: true, segmentExists: false });
    expect(msg).toContain("Create the segment first");
    // Tracks the CURRENT wording of the other message. Left pointing at a retired string this assertion
    // would pass for ever while proving nothing, which is worse than not having it.
    expect(msg).not.toContain("Create the segment above");
  });

  it("allows Next when the typed name matches an existing segment", () => {
    expect(blocksNext(st(), { subjectGiven: true, segmentExists: true })).toBe("");
  });

  it("does not leak the blank-name gate onto any other step", () => {
    // `subjectGiven` is set for the whole flow, so every step sees it. Only createSegment may use it —
    // Term Abbreviations must gate on missing codes alone.
    expect(blocksNext(step("newSegment", "abbreviations"), { subjectGiven: false })).toBe("");
    expect(blocksNext(step("newSegment", "reconcile"), { subjectGiven: false })).toBe("");
    expect(blocksNext(step("addUnit", "reconcile"), { subjectGiven: false })).toBe("");
  });
});

describe("stepUsesSegment", () => {
  const stepsOf = (flowId: string): FlowStep[] =>
    (FLOWS.filter((f) => f.id === flowId)[0] ?? { steps: [] }).steps;

  /** Instruction-only steps: the client reported the picker standing in front of one. */
  it("is false for an outside step", () => {
    for (const st of stepsOf("structure")) {
      if (st.screen.kind === "outside") expect(stepUsesSegment(st)).toBe(false);
    }
  });

  /**
   * ⚠ REGRESSION GUARD. The upload pause is SITE-WIDE, and it is a `component` step — so the old
   * `kind !== "outside"` literal put a segment picker in front of it the day it was added.
   */
  it("is false for both upload-pause steps, which are site-wide", () => {
    const pause = stepsOf("structure").filter(
      (st) => st.screen.kind === "component" && st.screen.id === "pauseUploads",
    );
    expect(pause.length).toBe(2);
    for (const st of pause) expect(stepUsesSegment(st)).toBe(false);
  });

  /** The Folder levels screen lists every segment itself, so a picker in front of it asks twice. */
  it("is false for the Folder levels step, which lists every segment on its own", () => {
    for (const st of stepsOf("structure")) {
      if (st.id === "levels") expect(stepUsesSegment(st)).toBe(false);
    }
  });

  it("is true for the migrate step, which acts on one chosen segment", () => {
    for (const st of stepsOf("structure")) {
      if (st.id === "migrate") expect(stepUsesSegment(st)).toBe(true);
    }
  });
});

describe("scopeFactsToFlow — segment facts must not tick a subject-scoped flow", () => {
  const pick = (id: string): Flow => FLOWS.filter((f) => f.id === id)[0];
  // Found on site 2026-08-20: "Add a department or unit" on GHO showed Group Management, Folder
  // Access and Folder Reconciliation all Done before anything had been added, because each fact asks
  // "does the SEGMENT have any at all" and GHO has 308 groups / 790 rows / 67 folder rows.
  const provisioned: FlowFacts = {
    segmentExists: true,
    groupsExist: true,
    foldersExist: true,
    abbreviationsMissing: 2,
  };


  it("drops the segment-scoped facts for addUnit", () => {
    const out = scopeFactsToFlow(pick("addUnit"), provisioned);
    expect(out.groupsExist).toBeUndefined();
    expect(out.foldersExist).toBeUndefined();
  });

  it("drops them for rename too", () => {
    const out = scopeFactsToFlow(pick("rename"), provisioned);
    expect(out.groupsExist).toBeUndefined();
    expect(out.foldersExist).toBeUndefined();
  });

  it("KEEPS them for newSegment, where the segment IS the subject", () => {
    const out = scopeFactsToFlow(pick("newSegment"), provisioned);
    expect(out.groupsExist).toBe(true);
    expect(out.foldersExist).toBe(true);
  });

  it("keeps them for runRecon and structure, which are about the segment", () => {
    expect(scopeFactsToFlow(pick("runRecon"), provisioned).foldersExist).toBe(true);
    expect(scopeFactsToFlow(pick("structure"), provisioned).foldersExist).toBe(true);
  });

  it("keeps abbreviationsMissing — it is reported live and IS subject-aware", () => {
    expect(scopeFactsToFlow(pick("addUnit"), provisioned).abbreviationsMissing).toBe(2);
  });

  it("keeps segmentExists — the segment genuinely does exist and step 1 says so", () => {
    expect(scopeFactsToFlow(pick("addUnit"), provisioned).segmentExists).toBe(true);
  });

  it("those steps then read `unknown`, never `done`", () => {
    const out = scopeFactsToFlow(pick("addUnit"), provisioned);
    const byId = (id: string): FlowStep => pick("addUnit").steps.filter((x) => x.id === id)[0];
    expect(stepState(byId("reconcile"), out)).toBe("unknown");
  });

  it("does not mutate the facts it was given", () => {
    const input = { ...provisioned };
    scopeFactsToFlow(pick("addUnit"), input);
    expect(input.groupsExist).toBe(true);
  });

  it("survives an undefined flow and undefined facts", () => {
    expect(scopeFactsToFlow(undefined, provisioned).groupsExist).toBe(true);
    expect(scopeFactsToFlow(pick("addUnit"), {} as FlowFacts)).toEqual({});
  });
});

/**
 * The rail's reachability rule, end to end over the real "Change the folder structure" flow.
 *
 * Extracted from FolderAdmin on 2026-09-07 so it could be tested at all. The bug that prompted it
 * shipped and was found by the client within the hour, and no test could have caught it while the
 * rule lived inline in the component's JSX.
 */
describe("the rail cannot walk past a blocked step", () => {
  const steps = flow("structure").steps;
  /* ⚠ THIS COMMENT WENT STALE THE DAY THE ABBREVIATIONS STEP WAS ADDED (2026-09-09) and listed
     five steps against six constants — the kind of drift that makes every index below it suspect.
     Derived from the flow now, so it cannot lie: a step added or reordered fails the length pin
     rather than silently shifting what `MIGRATE` means. */
  const PAUSE = 0, LEVELS = 1, ABBREV = 2, MIGRATE = 3, RECONCILE = 4, RESUME = 5;

  it("pins the step order the indices below depend on", () => {
    expect(steps.map((st) => st.id)).toEqual([
      "pauseUploads", "levels", "abbreviations", "migrate", "reconcile", "resumeUploads",
    ]);
  });

  it("blocks at step 1 while uploads are on, and at the closing step while they are paused", () => {
    expect(firstBlockedStepIndex(steps, { uploadsPaused: false })).toBe(PAUSE);
    /* ⚠ WAS `steps.length` UNTIL 2026-09-08, when the closing step gained its own gate. It now
       reports RESUME, and that CANNOT narrow the rail: `isStepReachable` allows `i <= firstBlocked`,
       and RESUME is the last index — so every step stays reachable and nobody is trapped by it. That
       is what makes gating the closing step safe as well as correct. */
    expect(firstBlockedStepIndex(steps, { uploadsPaused: true })).toBe(RESUME);
  });

  it("does not block on an UNKNOWN pause state — an unreadable config traps nobody", () => {
    expect(firstBlockedStepIndex(steps, {})).toBe(steps.length);
  });

  /* THE REPORTED BUG (client, 2026-09-07): *"I can bypass two the next steps on the side panel even
     if the sharepoint status is on"*. They had walked to step 3 while paused, then switched uploads
     back on. Next correctly refused; the rail still offered steps 2 and 3, because `maxIdx`
     remembers where you have been and nothing consulted the gate. */
  it("refuses the reported bypass: walked to step 3, uploads switched back on", () => {
    const firstBlocked = firstBlockedStepIndex(steps, { uploadsPaused: false });
    const at = { maxIdx: MIGRATE, idx: PAUSE, firstBlocked };
    expect(isStepReachable(PAUSE, at)).toBe(true);
    expect(isStepReachable(LEVELS, at)).toBe(false);
    expect(isStepReachable(MIGRATE, at)).toBe(false);
  });

  it("still allows every walked step once uploads are paused", () => {
    const firstBlocked = firstBlockedStepIndex(steps, { uploadsPaused: true });
    const at = { maxIdx: MIGRATE, idx: PAUSE, firstBlocked };
    expect(isStepReachable(LEVELS, at)).toBe(true);
    expect(isStepReachable(ABBREV, at)).toBe(true);
    expect(isStepReachable(MIGRATE, at)).toBe(true);
  });

  it("never offers a step the admin has not yet reached", () => {
    // Earned progress (1.0.332.0) survives the new cap: nothing beyond maxIdx, blocked or not.
    const at = { maxIdx: LEVELS, idx: LEVELS, firstBlocked: steps.length };
    expect(isStepReachable(MIGRATE, at)).toBe(false);
    expect(isStepReachable(RECONCILE, at)).toBe(false);
  });

  /* ⚠ THE CONDITION THAT KEEPS BACKWARD NAVIGATION FREE. Without `i <= idx`, a fact turning false
     while the admin stands on a later step strands them there — unable even to reach the step that
     needs fixing, which is the opposite of what the gate wants. */
  it("lets an admin standing past the block still reach every earlier step", () => {
    const firstBlocked = firstBlockedStepIndex(steps, { uploadsPaused: false });
    const at = { maxIdx: RESUME, idx: RECONCILE, firstBlocked };
    for (const i of [PAUSE, LEVELS, MIGRATE, RECONCILE]) {
      expect(isStepReachable(i, at)).toBe(true);
    }
  });

  /* ⚠ THE KNOWN, ACCEPTED COST, PINNED SO IT IS A DECISION RATHER THAN A SURPRISE.
     At the END of the flow the admin turns uploads back ON — the correct final state — which makes
     step 1 block again. Stepping BACK from there narrows the rail to where they now stand, so
     already-walked steps ahead grey out. They are NOT trapped: Next still advances, because levels,
     neither `levels` nor `migrate` is gated by `uploadsPaused`. Widening this restores the bypass
     above. (`migrate` has had a gate of its own since 2026-09-09, but it reads `pendingLevels`.) */
  it("greys already-walked steps ahead once uploads are back on and the admin steps back", () => {
    const firstBlocked = firstBlockedStepIndex(steps, { uploadsPaused: false });
    const at = { maxIdx: RESUME, idx: LEVELS, firstBlocked };
    expect(isStepReachable(LEVELS, at)).toBe(true);
    expect(isStepReachable(MIGRATE, at)).toBe(false);
    // ...and Next is what carries them forward again, so this is conservative, not a dead end.
    expect(blocksNext(steps[LEVELS], { uploadsPaused: false })).toBe("");
  });

  /* ⚠ Client, 2026-09-09: *"once the at the last step its on then block from going back"*. The Back
     button alone would not have done it — from step 5 with uploads on, `i <= at.idx` evaluates true
     for the migrate step, whose Rebuild is NOT gated on uploads. */
  it("a finished flow offers only the step it is on", () => {
    const at = { maxIdx: RESUME, idx: RESUME, firstBlocked: 0, complete: true };
    expect(isStepReachable(RESUME, at)).toBe(true);
    for (const i of [0, LEVELS, MIGRATE]) expect(isStepReachable(i, at)).toBe(false);
  });

  it("leaves reachability untouched when the flow is not complete", () => {
    // `complete` absent and `complete: false` must both behave exactly as before.
    const base = { maxIdx: RESUME, idx: RESUME, firstBlocked: 0 };
    expect(isStepReachable(MIGRATE, base)).toBe(true);
    expect(isStepReachable(MIGRATE, { ...base, complete: false })).toBe(true);
  });

  it("caps nothing on a flow with no gated step at all", () => {
    // runRecon is a single reconciliation step: no gate, so the rail behaves exactly as before.
    const rr = flow("runRecon").steps;
    expect(firstBlockedStepIndex(rr, { uploadsPaused: false })).toBe(rr.length);
  });
});

describe("the structure flow, step by step", () => {
  const steps = flow("structure").steps;

  it("runs pause -> levels -> abbreviations -> migrate -> reconcile -> resume", () => {
    expect(steps.map((x) => x.id)).toEqual([
      "pauseUploads", "levels", "abbreviations", "migrate", "reconcile", "resumeUploads",
    ]);
  });

  /* ⚠ THE ORDER IS THE DESIGN, not a preference. Step 2 codes the level and reaches the Term Store;
     step 4 renames the existing folders to match. Codes missing at step 4 leave the migration
     building destinations it has no name for. */
  it("puts the abbreviations step between editing the levels and moving the folders", () => {
    const ids = steps.map((x) => x.id);
    expect(ids.indexOf("levels")).toBeLessThan(ids.indexOf("abbreviations"));
    expect(ids.indexOf("abbreviations")).toBeLessThan(ids.indexOf("migrate"));
  });

  /* ⚠ THE SCAN IS THE GATE NOW, NOT A LOCK ON WHAT WAS STAGED (2026-09-09). Reported by the client:
     step 3 said "there is nothing to move to" — WITHOUT LOOKING — after they had added subunit terms,
     and Next sat open beside it. Removing the lock alone would have left the step skippable, because
     `pendingLevels` holds Next only while a chain change is staged and those cases stage nothing. */
  it("holds Next until the check has actually been run", () => {
    const migrate = step("structure", "migrate");
    expect(blocksNext(migrate, { migrateScanRan: false })).toContain("Run the check first");
    expect(blocksNext(migrate, { migrateScanRan: true })).toBe("");
    // Unknown never gates — the standalone Migrate tab reports nothing at all.
    expect(blocksNext(migrate, {})).toBe("");
  });

  it("still holds Next while a staged change has not been applied", () => {
    expect(blocksNext(step("structure", "migrate"), { pendingLevels: true, migrateScanRan: true }))
      .toContain("has not been applied yet");
  });

  /* ⚠ THIS FLOW GAINED AN ABBREVIATIONS STEP ON 2026-09-09, so reconcile's abbreviation lock is no
     longer inert here — the count IS set now, and a term with no code holds reconciliation as well
     as Next. That is wanted (reconciliation skips an uncoded term and creates no folder) and it
     cannot deadlock: the abbreviations gate stops you reaching reconcile in that state anyway, and
     Back is always open. Undefined still never locks, which is what this pins. */
  it("never locks reconciliation in a flow that cannot know the abbreviation count", () => {
    expect(isLocked(step("structure", "reconcile"), {})).toBe(false);
  });

  /* ⚠ TWO SINCE 2026-09-09, AND THE SECOND IS THE WHOLE POINT OF ADDING THAT STEP. A level named by
     term codes cannot reach the migration with terms that have none, or the run builds destinations
     it has no name for. The gate itself is the one that has held the abbreviations step since it was
     written — the step was added so that rule finally applies in THIS flow. */
  it("gates the pause step and the abbreviations step, and nothing else", () => {
    const all: FlowFacts = {
      segmentExists: false, groupsExist: false, foldersExist: false,
      abbreviationsMissing: 9, pendingLevels: false, subjectFound: false,
      uploadsPaused: false,
    };
    const gated = steps.filter((x) => blocksNext(x, all).length > 0).map((x) => x.id);
    expect(gated).toEqual(["pauseUploads", "abbreviations"]);
  });

  it("releases the abbreviations step once every term has a code", () => {
    expect(blocksNext(step("structure", "abbreviations"), { abbreviationsMissing: 0 })).toBe("");
    // Unknown never gates: an unread count must not strand an admin mid-flow.
    expect(blocksNext(step("structure", "abbreviations"), {})).toBe("");
  });

  /* Both pause steps read one fact in OPPOSITE directions, and a rail that ticked the closing step
     while the site was still paused would confirm the wrong thing. */
  it("ticks the two pause steps in opposite directions", () => {
    expect(stepState(step("structure", "pauseUploads"), { uploadsPaused: true })).toBe("done");
    expect(stepState(step("structure", "resumeUploads"), { uploadsPaused: true })).toBe("todo");
    expect(stepState(step("structure", "pauseUploads"), { uploadsPaused: false })).toBe("todo");
    expect(stepState(step("structure", "resumeUploads"), { uploadsPaused: false })).toBe("done");
    expect(stepState(step("structure", "pauseUploads"), {})).toBe("unknown");
    expect(stepState(step("structure", "resumeUploads"), {})).toBe("unknown");
  });

  /* The segment picker must not stand in front of a step that spends no segment — the rule the
     site-wide upload pause walked into twice. */
  it("asks for a segment only on the steps that use one", () => {
    expect(stepUsesSegment(step("structure", "pauseUploads"))).toBe(false);
    expect(stepUsesSegment(step("structure", "resumeUploads"))).toBe(false);
    expect(stepUsesSegment(step("structure", "migrate"))).toBe(true);
    expect(stepUsesSegment(step("structure", "reconcile"))).toBe(true);
  });
});
