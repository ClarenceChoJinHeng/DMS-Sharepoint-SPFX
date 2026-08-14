/**
 * Folder Management's guided flows — what the steps are, and what may stop you.
 *
 * Spec: docs/superpowers/specs/2026-08-14-folder-management-guided-flows-design.md
 *
 * Pure and SPFx-free. The tab bar this fronts was already ordered by when the work happens; that was
 * not enough, because five equal doors do not say that four of them are steps of one job. The order is
 * the thing that breaks — miss Term Abbreviations and reconciliation creates NOTHING, silently.
 *
 * THE GOVERNING RULE, from the client (2026-08-14): *"the flow should not stop them from doing the
 * work."* So locks are few, each rests on ONE definitive read, and every other step is
 * marked-but-passable. `isLocked` is written so that anything unknown — a fact not read yet, or a read
 * that failed — is NOT a lock. Fail-open by construction rather than by remembering to.
 */

/** Where a step's work happens. */
export type StepScreen =
  /** Outside this tool entirely: the term store, or Power Automate. */
  | { kind: "outside" }
  /** A tab of FolderManager, driven via its `initialTab` prop. */
  | { kind: "tab"; tab: string }
  /** A component FolderManager does not host, mounted from userAccess. */
  | { kind: "component"; id: "groups" | "folderAccess" };

/**
 * A fact a step can be locked behind.
 *
 * Only three, and each is answerable by ONE read. Anything needing judgement, or a walk we cannot
 * afford, is not on this list and therefore cannot lock a step.
 */
export type LockFact = "segmentExists" | "abbreviationsComplete" | "pendingLevels";

export interface FlowStep {
  id: string;
  label: string;
  /** One line: what to do here. */
  hint: string;
  screen: StepScreen;
  /** Locked until this fact is known-true. Absent = never locked. */
  lock?: { fact: LockFact; reason: string };
}

export interface Flow {
  id: string;
  label: string;
  blurb: string;
  /** `destructive` is styled apart and never presented as a "get started" card. */
  tone: "normal" | "destructive";
  /** Needs an existing segment picked before the steps make sense. */
  needsSegment: boolean;
  /**
   * Asks for a subject term up front, which is what makes the term-store step checkable at all:
   * "adding Treasury under Group Finance" can be looked for in the tree; "you did something in the term
   * store" cannot.
   */
  asksSubject?: "add" | "rename";
  steps: FlowStep[];
}

/** What has been read about the chosen segment. EVERY field may be undefined, meaning not known. */
export interface FlowFacts {
  /** A DMS Config mode row exists. */
  segmentExists?: boolean;
  /** Site groups matching the segment's code prefix exist. Advisory — the convention is a suggestion. */
  groupsExist?: boolean;
  /** Group Map rows exist for this segment. */
  folderAccessRows?: boolean;
  /** Folder Map rows exist for this segment. */
  foldersExist?: boolean;
  /**
   * How many terms have no folder code. `undefined` means NOT WALKED YET — the state on the picker
   * screen — and must never be read as zero.
   */
  abbreviationsMissing?: number;
  /** The mode row carries a staged chain. */
  pendingLevels?: boolean;
  /** The subject term was found in the tree (flows 2 and 4). */
  subjectFound?: boolean;
}

const RECONCILE: FlowStep = {
  id: "reconcile",
  label: "Folder Reconciliation",
  hint: "Build the folders and apply the permissions.",
  screen: { kind: "tab", tab: "reconciliation" },
  // THE one lock that earns its place: reconciliation silently skips a term with no code and creates no
  // folder for it, with no error anywhere. Only bites once the count is known — see isLocked.
  lock: {
    fact: "abbreviationsComplete",
    reason: "Some terms still have no folder code. Reconciliation skips those and creates no folder for them.",
  },
};

const ABBREVIATIONS: FlowStep = {
  id: "abbreviations",
  label: "CRS Term Abbreviations",
  hint: "Give every term a short folder code. A term with no code gets no folder.",
  screen: { kind: "tab", tab: "abbreviations" },
};

const GROUPS: FlowStep = {
  id: "groups",
  label: "Group Management",
  hint: "Create the groups for the unit. If they already exist, there is nothing to do here.",
  screen: { kind: "component", id: "groups" },
};

const FOLDER_ACCESS: FlowStep = {
  id: "folderAccess",
  label: "Folder Access",
  hint: "Map each group to its segment, tier and role.",
  screen: { kind: "component", id: "folderAccess" },
};

export const FLOWS: Flow[] = [
  {
    id: "newSegment",
    label: "Add a new segment",
    blurb: "A whole new business segment or project, with its own term set and tiers.",
    tone: "normal",
    needsSegment: false,
    steps: [
      {
        id: "termSet",
        label: "Create the term set",
        hint: "In the term store, create the term set and its terms — nested exactly as deep as the tiers you will name.",
        screen: { kind: "outside" },
      },
      {
        id: "createSegment",
        label: "New segment",
        hint: "Name the segment and its tiers, and point at the term set. Refuses if the set's depth does not match.",
        screen: { kind: "tab", tab: "newsegment" },
      },
      {
        ...ABBREVIATIONS,
        // Only from here on: before the segment row exists there is no tree to code.
        lock: {
          fact: "segmentExists",
          reason: "Create the segment first — until it exists there is no term tree to give codes to.",
        },
      },
      GROUPS,
      FOLDER_ACCESS,
      RECONCILE,
    ],
  },
  {
    id: "addUnit",
    label: "Add a department or unit",
    blurb: "One more department or unit inside a segment that already exists. The everyday job.",
    tone: "normal",
    needsSegment: true,
    asksSubject: "add",
    steps: [
      {
        id: "addTerm",
        label: "Add the term",
        hint: "In the term store, add the term under its parent. The next step will show you whether it took.",
        screen: { kind: "outside" },
      },
      ABBREVIATIONS,
      GROUPS,
      FOLDER_ACCESS,
      RECONCILE,
    ],
  },
  {
    id: "structure",
    label: "Change the folder structure",
    blurb: "Add, reorder or remove a level below Unit, and move the existing folders to match.",
    tone: "normal",
    needsSegment: true,
    steps: [
      {
        id: "pauseFlows",
        label: "Power Automate (optional)",
        hint: "You can leave the flows running. Pausing Auto-route and folder approval avoids a lot of flow runs that do nothing.",
        screen: { kind: "outside" },
      },
      {
        id: "levels",
        label: "Folder Structure Management",
        hint: "Edit the levels. Saved as a pending change — nothing moves and nothing goes live yet.",
        screen: { kind: "tab", tab: "structure" },
      },
      {
        id: "migrate",
        label: "Move existing folders",
        hint: "Move the documents into the new shape. This applies the new structure as its last step.",
        screen: { kind: "tab", tab: "migrate" },
        lock: {
          fact: "pendingLevels",
          reason: "There is no pending structure change to move to — edit the levels first, or this has nothing to do.",
        },
      },
    ],
  },
  {
    id: "rename",
    label: "Rename or re-code a folder",
    blurb: "Change a folder's name by renaming its term, or by changing its short code.",
    tone: "normal",
    needsSegment: true,
    asksSubject: "rename",
    steps: [
      {
        id: "renameTerm",
        label: "Rename the term",
        hint: "In the term store, RENAME it — never delete and re-add. A new GUID orphans the abbreviation, folder and group rows at once.",
        screen: { kind: "outside" },
      },
      ABBREVIATIONS,
      RECONCILE,
    ],
  },
  {
    id: "retire",
    label: "Retire a segment",
    blurb: "Stop offering a segment. Its documents are not touched unless you ask separately.",
    tone: "destructive",
    needsSegment: true,
    steps: [
      {
        id: "moveOut",
        label: "Move the documents out",
        hint: "Move anything worth keeping somewhere else first.",
        screen: { kind: "tab", tab: "migrate" },
      },
      {
        id: "delete",
        label: "Segments → Delete",
        hint: "Removes the segment and its access rows. Deleting the folders is a separate, typed confirmation.",
        screen: { kind: "tab", tab: "newsegment" },
        lock: {
          fact: "segmentExists",
          reason: "There is no such segment to retire.",
        },
      },
    ],
  },
];

export function flowById(id: string): Flow | undefined {
  return FLOWS.filter((f) => f.id === id)[0];
}

/**
 * Is this step locked?
 *
 * **Only ever true when the fact is KNOWN and unmet.** Undefined — not read yet, or a read that failed —
 * is never a lock. That is the client's rule made structural: a transient error, or a check we have not
 * paid for, must not stop someone doing the work.
 */
export function isLocked(step: FlowStep, facts: FlowFacts): boolean {
  if (!step || !step.lock) return false;
  const f = facts ?? {};
  switch (step.lock.fact) {
    case "segmentExists":
      return f.segmentExists === false;
    case "pendingLevels":
      return f.pendingLevels === false;
    case "abbreviationsComplete":
      // A count of undefined means the tree has not been walked — the picker's state. Never a lock, and
      // never read as zero either.
      return typeof f.abbreviationsMissing === "number" && f.abbreviationsMissing > 0;
    default:
      return false;
  }
}

/** Why it is locked, for the UI. Blank when it is not. */
export function lockReason(step: FlowStep, facts: FlowFacts): string {
  return isLocked(step, facts) ? (step.lock?.reason ?? "") : "";
}

/**
 * How far along a step is.
 *
 * `unknown` is a first-class answer, rendered as "not checked" — never as done, never as outstanding.
 * Steps whose completion is unknowable (the term store, Power Automate) stay `unknown` unless a subject
 * makes them checkable.
 */
export type StepState = "done" | "todo" | "unknown";

export function stepState(step: FlowStep, facts: FlowFacts): StepState {
  const f = facts ?? {};
  switch (step?.id) {
    case "createSegment":
    case "delete":
      return f.segmentExists === undefined ? "unknown" : f.segmentExists ? "done" : "todo";
    case "termSet":
      // Implied by the segment existing: SegmentCreator refuses without a resolvable set of the right
      // depth, so a segment row is proof the term set was there.
      return f.segmentExists === true ? "done" : "unknown";
    case "addTerm":
    case "renameTerm":
      // Checkable ONLY because the flow asked what the subject was.
      return f.subjectFound === undefined ? "unknown" : f.subjectFound ? "done" : "todo";
    case "abbreviations":
      if (typeof f.abbreviationsMissing !== "number") return "unknown";
      return f.abbreviationsMissing === 0 ? "done" : "todo";
    case "groups":
      // Advisory: the naming convention is a suggestion, so a hand-named group reads as absent.
      return f.groupsExist === undefined ? "unknown" : f.groupsExist ? "done" : "todo";
    case "folderAccess":
      return f.folderAccessRows === undefined ? "unknown" : f.folderAccessRows ? "done" : "todo";
    case "reconcile":
      return f.foldersExist === undefined ? "unknown" : f.foldersExist ? "done" : "todo";
    case "levels":
      return f.pendingLevels === true ? "done" : "unknown";
    case "migrate":
    case "moveOut":
    case "pauseFlows":
      // Nothing to read. A migration run leaves no marker, and Power Automate is unreachable.
      return "unknown";
    default:
      return "unknown";
  }
}

/**
 * Where to open the flow.
 *
 * The first step that is not `done`, so a flow resumes where it was left — which matters because three
 * of the five leave the tool for the term store and will not be finished in one sitting. An `unknown`
 * step counts as somewhere to be, because it is the honest place to put someone who may still have work
 * there. All done → the LAST step, so a finished flow ends on its final screen rather than bouncing back
 * to the beginning.
 */
export function firstIncompleteStep(flow: Flow, facts: FlowFacts): number {
  const steps = flow?.steps ?? [];
  for (let i = 0; i < steps.length; i++) {
    if (stepState(steps[i], facts) !== "done") return i;
  }
  return Math.max(0, steps.length - 1);
}

/** Steps still outstanding, for the picker's summary line. */
export function remainingCount(flow: Flow, facts: FlowFacts): number {
  return (flow?.steps ?? []).filter((s) => stepState(s, facts) !== "done").length;
}

/**
 * Fold a term label for comparison.
 *
 * The trap this exists for: GHO contains `Group Legal, Risk ＆ Compliance` with a **FULLWIDTH ＆**, which
 * SharePoint requires. An admin typing a normal `&` would fail to match a term that exists perfectly
 * well — and the flow would tell them they had not done step one. Also folds case, collapses runs of
 * whitespace, and strips zero-width characters, which arrive via copy-paste from Word.
 */
export function normaliseLabel(s: string): string {
  return (s ?? "")
    .replace(/[​-‍﻿]/g, "")
    .replace(/＆/g, "&")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/** Does any of these labels match what was typed? */
export function labelMatches(labels: readonly string[], wanted: string): boolean {
  const want = normaliseLabel(wanted);
  if (want.length === 0) return false;
  return (labels ?? []).filter((l) => normaliseLabel(l) === want).length > 0;
}

/**
 * Labels that nearly match, for the "check the spelling" hint.
 *
 * Deliberately crude — containment either way. The goal is to catch a typo or a half-typed entry, not to
 * be a spell checker; a near-miss list that looks clever invites trusting it.
 */
export function nearMatches(labels: readonly string[], wanted: string, limit = 3): string[] {
  const want = normaliseLabel(wanted);
  if (want.length < 2) return [];
  return (labels ?? [])
    .filter((l) => {
      const n = normaliseLabel(l);
      return n !== want && n.length > 0 && (n.indexOf(want) !== -1 || want.indexOf(n) !== -1);
    })
    .slice(0, limit);
}
