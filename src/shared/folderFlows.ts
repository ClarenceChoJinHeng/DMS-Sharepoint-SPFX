/**
 * Folder Management's guided flows — what the steps are, and what may stop you.
 *
 * Spec: docs/superpowers/specs/2026-08-14-folder-management-guided-flows-design.md
 *
 * Pure and SPFx-free. The tab bar this fronts was already ordered by when the work happens; that was
 * not enough, because five equal doors do not say that four of them are steps of one job. The order is
 * the thing that breaks — miss Term Abbreviations and reconciliation creates NOTHING, silently.
 *
 * THE GOVERNING RULE: *"the flow should not stop them from doing the work."* **Recorded as a client
 * quote until 2026-08-17, when Clarence corrected it — it was never theirs.** Kept as a design rule
 * because it earns its place on its own terms, but it is OURS, so it can be traded off rather than
 * treated as a constraint handed down. The actual requirement is narrower and stronger: *"I just want to
 * ensure this flow is working properly and it should make them understand how it work."*
 *
 * What that means in practice: locks are few and each rests on ONE definitive read, but a step that is
 * definitively not done SHOULD say so and SHOULD hold the Next button (see `blocksNext`). `isLocked` is
 * still written so anything unknown — a fact not read yet, or a read that failed — is NOT a lock;
 * fail-open by construction rather than by remembering to. That protects against a wrong check, which is
 * a different thing from declining to check at all.
 */

/** Where a step's work happens. */
export type StepScreen =
  /** Outside this tool entirely: the term store, or Power Automate. */
  | { kind: "outside" }
  /** A tab of FolderManager, driven via its `initialTab` prop. */
  | { kind: "tab"; tab: string }
  /** A component FolderManager does not host, mounted from userAccess. */
  | { kind: "component"; id: "groups" | "pauseUploads" };

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
  asksSubject?: "add" | "rename" | "newSegment";
  steps: FlowStep[];
}

/** What has been read about the chosen segment. EVERY field may be undefined, meaning not known. */
export interface FlowFacts {
  /** A DMS Config mode row exists. */
  segmentExists?: boolean;
  /** Site groups matching the segment's code prefix exist. Advisory — the convention is a suggestion. */
  groupsExist?: boolean;
  /** Folder Map rows exist for this segment. */
  foldersExist?: boolean;
  /**
   * How many terms have no folder code. `undefined` means NOT WALKED YET — the state on the picker
   * screen — and must never be read as zero.
   */
  abbreviationsMissing?: number;
  /**
   * The abbreviations screen is READING the term store right now.
   *
   * A separate fact from the count, and it has to be. The count is `undefined` for the whole read,
   * `undefined` never gates, and so Next was clickable on a step that said "Reading the term store…"
   * — an answer seconds away and certainly not yet known (reported on site 2026-08-18).
   *
   * **In flight is not unknown**, the same split that made `subjectGiven` right: fail-open exists for
   * reads that can FAIL, not for reads still running. Gating here strands nobody, because it clears
   * itself; gating on a failed read would strand everybody.
   */
  abbreviationsLoading?: boolean;
  /** The mode row carries a staged chain. */
  pendingLevels?: boolean;
  /**
   * Are uploads paused site-wide right now?
   *
   * Drives the ticks on BOTH pause steps, and they read it in opposite directions: pausing is done
   * when this is true, resuming when it is false. `undefined` (unreadable config) leaves both
   * unknown — it must never claim uploads are on, because an admin who believes they paused when
   * they did not will migrate on top of live traffic.
   */
  uploadsPaused?: boolean;
  /** The subject term was found in the tree (flows 2 and 4). */
  subjectFound?: boolean;
  /**
   * The admin has typed the subject — for flow 1, the new segment's name.
   *
   * Set ONLY when the segment list was readable, so this never fires on a failed read. That split is
   * the point: **fail-open exists for reads that can fail, not for a text box nobody has filled in.**
   * A throttled list cannot blank a local field, so gating on it cannot strand anyone — whereas
   * treating blank as "unknown" left Next enabled on a step plainly not done, which is what the client
   * reported twice (2026-08-17).
   */
  subjectGiven?: boolean;
  /**
   * Has a segment been picked, on a flow that works on one?
   *
   * THREE STATES IN ONE OPTIONAL BOOLEAN, the same trick `subjectGiven` above uses, and for the same
   * reason: `true` a segment is chosen; `false` none is chosen AND the segment list read fine, so
   * the admin simply has not answered; `undefined` the list could not be read, or the flow does not
   * work on a single segment at all.
   *
   * ⚠ THE `undefined` CASE IS WHY THIS IS NOT TWO FIELDS. An unreadable segment list leaves nothing
   * to pick FROM, so gating on it would strand an admin with no way forward — and "unknown never
   * gates" is already the universal rule here, so one optional boolean expresses the whole thing
   * with no special branch.
   */
  segmentChosen?: boolean;
  /**
   * Has a reconciliation run FINISHED during this visit to the flow?
   *
   * ⚠ FINISHED, NOT SUCCEEDED, and that distinction is the whole safety of the gate it feeds. In the
   * structure flow `reconcile` sits before `resumeUploads` — so gating on SUCCESS would leave
   * uploads switched off site-wide for as long as reconciliation kept failing, and the admin could
   * not reach the step that turns them back on. A run that ends with errors satisfies this; the
   * errors are on screen for them to read. The client's ask was "ensure they run it", which is
   * exactly what this asks.
   *
   * ⚠ SESSION-SCOPED, so a page refresh clears it and the run is asked for again. Accepted rather
   * than solved: reconciliation is idempotent, so a second run is safe, and the durable
   * alternatives (an audit row, a stored marker) can each fail silently — which would strand the
   * flow over a read rather than over the work.
   */
  reconcileRan?: boolean;
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

/**
 * The abbreviations step, with the label and hint the FLOW needs.
 *
 * ⚠ A FACTORY SINCE 2026-09-06, AND THE `id` MUST STAY `abbreviations` IN EVERY ONE. The client's
 * redesign names it "Create Term Abbreviation" when adding and "Rename Term Abbreviation" when
 * renaming - but `NEXT_GATED_STEPS` keys on the id, and so does the padlock that stops an admin
 * reaching reconciliation with terms that have no code. Give one of these a new id and that gate
 * silently stops applying to it.
 */
function abbreviationsStep(label: string, hint: string): FlowStep {
  return { id: "abbreviations", label, hint, screen: { kind: "tab", tab: "abbreviations" } };
}

const ABBREVIATIONS: FlowStep = {
  id: "abbreviations",
  label: "CRS Term Abbreviations",
  /* The mandatory sentence the client asked for (2026-09-06) sits here rather than in the screen,
     because this is the hint the flow renders above it — the screen has no intro of its own once the
     top-folder and below-Unit notes came off. */
  /* WARN: THE NEWLINE IS LOAD-BEARING (client, 2026-09-06: *"Basically break line"*). It renders
     only because `s.stepHint` sets `white-space: pre-line`; strip that and the two sentences run
     together again with nothing failing. No other hint contains one, so nothing else changes shape. */
  hint:
    "Give every term a short folder code. A term with no code gets no folder.\n" +
    "All fields marked * are mandatory.",
  screen: { kind: "tab", tab: "abbreviations" },
};

/*
 * ⚠ THE `groups` STEP IS GONE FROM EVERY FLOW (client, 2026-09-06: *"Remove Group Management from
 * ALL Folder Management … Let client assign members under USER ACCESS MANAGEMENT > GROUP
 * MANAGEMENT"*). It stood between the abbreviations and reconciliation in `newSegment` and
 * `addUnit`.
 *
 * DELETED rather than parked. An unused const is a lint warning here, and this project's own rule
 * from `inviteToGroup` applies: something kept alive doing exactly the thing that was withdrawn is
 * only ever a matter of time before it is wired back up. Re-adding it is eight lines — a `label`, a
 * `hint`, and `screen: { kind: "component", id: "groups" }` — and the component id is still in the
 * `screen` union above, so the mount point survives.
 *
 * ⚠ NOTHING ABOUT PROVISIONING CHANGED. Bulk provisioning still creates every group and writes its
 * mappings, and reconciliation still grants from those rows, so a segment built through this flow is
 * complete without the step. What moved is WHO adds the people, and where.
 */

/*
 * ⚠ THE `folderAccess` STEP IS GONE (2026-08-23). It existed to ADD PEOPLE, which the Group
 * Management step above it now does itself — the member editor is mounted there. Client: *"I think we
 * can remove folder access and also in the step, reason being I separated Folder Access to allow
 * client to add the user separately but since the Group Management will be doing most of the job of
 * bulk group and allow client to add user on their own, that makes Folder Access redundant."*
 *
 * Its rule survives in the step that absorbed it: **membership must never gate Next.** Membership is
 * INTENT and intent is not checkable — nothing can tell whether the RIGHT people are in a group. An
 * empty group is a valid end state, because reconciliation grants to the GROUP and not to its
 * members, so the grant applies the moment somebody is added with nothing to re-run.
 */

export const FLOWS: Flow[] = [
  {
    id: "newSegment",
    label: "Add a new segment",
    blurb: "A whole new business segment or project, with its own term set and tiers.",
    tone: "normal",
    needsSegment: false,
    /**
     * Asks for the new segment's NAME up front, which is the only thing that makes "have they created
     * it yet" answerable (2026-08-17, client: the Next button should not be available until the New
     * segment form is filled in and saved).
     *
     * Counting mode rows was the obvious alternative and is wrong: a baseline taken when the flow opens
     * resets on a page refresh, so the flow would then refuse work the admin had already done — worse
     * than no check, and a direct breach of "the flow should not stop them from doing the work".
     * Matching a NAME is derived from data, so it survives a refresh, a second tab and a shared session.
     *
     * It also stops being an extra question: the name is what they were going to type into the form
     * anyway, and naming it lets steps 3-6 open on that segment instead of an empty picker.
     */
    asksSubject: "newSegment",
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
      RECONCILE,
    ],
  },
  {
    id: "addUnit",
    label: "Add a department or unit",
    blurb: "Add another department or unit under an existing segment.",
    tone: "normal",
    needsSegment: true,
    /* ⚠ KEPT THOUGH THE SUBJECT BOX IS GONE. The "which department did you add?" field lived on the
       `addTerm` step this flow no longer has, so nothing asks for it any more - but `asksSubject` is
       ALSO what `scopeFactsToFlow` reads to drop the segment-level facts, and without it the rail
       would tick Reconciliation as Done before anything was provisioned (the 1.0.189.0 defect). */
    asksSubject: "add",
    /* ⚠ FOUR STEPS DOWN TO TWO (client's redesign, 2026-09-06). The `addTerm` instruction step is
       gone; the Term Store link it carried now renders on the abbreviations step, which is where an
       admin is standing when they need it. Nothing about the TERM STORE work changed - it still has
       to be done there first, and the abbreviation screen still shows nothing for a term that does
       not exist yet, which is how this step confirms it took. */
    steps: [
      abbreviationsStep(
        "Create Term Abbreviation",
        "Add the term in the term store, then give it a short folder code below.",
      ),
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
        // FIRST, and not optional. The Power Automate step used to lead here and was the only
        // preparation offered — but Clarence, 2026-08-19: *"they are not going to go to Pause power
        // automate"*. A file uploaded mid-migration lands in the OLD shape and, if it arrives after
        // its folder was scanned, is never moved; the staged-apply guard then keeps the migration
        // pending rather than letting it go live, which is safe but can loop.
        id: "pauseUploads",
        label: "Temporarily disable uploads",
        hint: "Temporarily restrict document uploads across the site while you reorganise the folder structure.",
        screen: { kind: "component", id: "pauseUploads" },
      },
      /* ⚠ THE "Power Automate (optional)" STEP IS GONE (client's redesign, 2026-09-06 — five steps
         down to four), AND WITH IT THE ONLY PLACE THIS TOOL MENTIONS THE FLOWS.
         What it said is still TRUE and is now told to nobody: leaving Auto-route and the
         folder-approval flows running is SAFE — Auto-route reads moderation status and takes the
         False branch for a moved pending file — but a migration produces hundreds of no-op runs,
         which burn the daily quota, and an exhausted quota means the next real approval is never
         routed, silently.
         Raised with the client before removing it; accepted. If flow quota is ever exhausted during
         a structure change, this is why. */
      {
        id: "levels",
        label: "Folder Structure Management",
        hint:
          "Edit the folder levels. Changes are saved as pending and will not affect existing " +
          "folders until applied.",
        screen: { kind: "tab", tab: "structure" },
      },
      {
        id: "migrate",
        label: "Move existing folders",
        /* ⚠ THE CLIENT'S COPY DROPS THE ONE FACT THAT MAKES THIS STEP THE POINT OF NO RETURN, so it
           is kept on the end: the migration APPLIES the pending chain as its last action, and only
           when a fresh scan finds no drift left. Until it runs, uploads are still filing into the
           old shape - which is why "saved as pending" on the previous step is true and why walking
           past this one leaves the segment half-changed. */
        hint:
          "Apply the new folder structure to existing documents. This is what makes the pending " +
          "change live — until it runs, uploads keep using the old structure.",
        screen: { kind: "tab", tab: "migrate" },
        lock: {
          fact: "pendingLevels",
          reason: "There is no pending structure change to move to — edit the levels first, or this has nothing to do.",
        },
      },
      /* ⚠ RECONCILIATION IS IN THIS FLOW SINCE 2026-09-06, REVERSING A DELIBERATE EXCLUSION.
         It was kept OUT because reconciliation walks the TERM TREE, not `Levels` — so a change to
         the levels below Unit creates nothing for it to do, and offering it invited an hour-long run
         that changes nothing. There was a test named "does NOT offer reconciliation in the structure
         flow" saying exactly that.
         The client asked for it anyway, having been shown that reasoning and the two ways to read
         their own slides: *"A"* — a step of its own, before uploads resume.
         It is defensible on the OTHER reading of the request: a structure change moves folders, and
         re-asserting permissions before letting people back in is worth a run even when the tree is
         unchanged. What it must never be described as is "applying the structure change" — the
         migration already did that, as its own last step. */
      RECONCILE,
      {
        // ⚠ THE WHOLE REASON THE PAUSE IS MANUAL. Turning uploads back on is the step that gets
        // forgotten, and a forgotten pause is a DMS that quietly accepts no documents — with a
        // banner explaining why, which makes it look deliberate. It gets its own step so the flow
        // does not end until someone has looked at it.
        id: "resumeUploads",
        /* The client's title and description (2026-09-06).
           ⚠ THE OLD LABEL WAS DOING SAFETY WORK AND THE NEW ONE MUST KEEP DOING IT. This is the step
           that gets forgotten, and a forgotten pause is a DMS that quietly accepts no documents
           behind a banner that makes it look deliberate. The hint keeps the consequence on the end
           for that reason - "Enable upload access" alone states an action and not what happens if
           nobody performs it. */
        label: "Enable Upload",
        hint: "Enable upload access. Until you do, nobody can file a document.",
        screen: { kind: "component", id: "pauseUploads" },
      },
    ],
  },
  {
    id: "rename",
    label: "Rename or re-code a folder",
    blurb: "Change a folder's name by renaming its term, or by changing its short code.",
    tone: "normal",
    needsSegment: true,
    // Kept for `scopeFactsToFlow`, exactly as on `addUnit` — see the note there.
    asksSubject: "rename",
    /* ⚠ THE "NEVER DELETE AND RE-ADD" WARNING LOST ITS STEP, AND IT IS THE MOST EXPENSIVE MISTAKE ON
       THIS SCREEN. Deleting a term and re-adding it gives a NEW GUID, which orphans the abbreviation
       row, the folder-map row and every group mapping at once - reconciliation then repairs only a
       1:1 label match and reports the rest. The sentence now rides on the abbreviations hint below,
       because the client's redesign has no instruction step to put it on. If that hint is ever
       shortened, this warning needs somewhere else to live. */
    steps: [
      abbreviationsStep(
        "Rename Term Abbreviation",
        "Rename the folder in the Term Store and on this page, where applicable. Always RENAME a " +
          "term — never delete and re-add it, which orphans its abbreviation, folder and group rows.",
      ),
      RECONCILE,
    ],
  },
  {
    // The one job that is not a CHANGE — re-asserting what should already be true.
    //
    // It exists because "All tools" was removed from the picker (client, 2026-08-20: they did not want
    // the five screens reachable without a flow), and reconciliation was the only job left with no door
    // of its own. Every other screen belongs to a flow; this one was reachable ONLY as the tail of
    // three of them, so an admin needing to re-run it would have had to start the flow for adding a
    // unit — which is exactly the kind of instruction that gets misremembered.
    //
    // ⚠ NOT for adding people to a group. Reconciliation grants to the GROUP, never to its members, so
    // a new member inherits the grant the moment they join and there is nothing to re-run. Saying
    // otherwise would train an admin to run an hour-long job for no reason.
    id: "runRecon",
    label: "Run folder reconciliation",
    // ⚠ "after a group was DELETED" was in this line until 2026-08-20 and was WRONG — proven on
    // site. Reconciliation grants to groups that exist and NEVER creates them, so a deleted group
    // is not restored by any number of runs: the mapping rows survive, point at nothing, and are
    // silently ignored. The repair is Bulk provisioning first, THEN reconciliation. Sending an
    // admin here for a deleted group is a dead end that reports success.
    blurb: "Re-apply the existing folder structure and permissions to bring them back in line with the configured setup. Use this after re-creating a group, when folder permissions have been changed manually in SharePoint, or when you are unsure if a segment is configured correctly. This does not change your existing setup. If a group has been deleted, re-create it first using Bulk Provisioning in Group Management. Reconciliation can only re-apply access to existing groups; it cannot restore deleted groups.",
    tone: "normal",
    // The reconciliation screen has its own segment tick-list, so a picker in front of it would ask
    // twice — the same mistake already fixed for the Folder levels step.
    needsSegment: false,
    steps: [RECONCILE],
  },
  {
    id: "retire",
    label: "Retire a segment",
    blurb: "Remove the segment from the available options so users can no longer use it for new work. Any documents already stored in the segment will remain available and will not be deleted.",
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
 * The steps whose `todo` is DEFINITIVE enough to disable the Next button.
 *
 * An explicit list rather than "any todo step", because most of the facts behind `stepState` are
 * advisory and gating on them would trap someone who did the work a different way:
 *
 *   - `groups` reads site groups by the NAMING CONVENTION, which `suggestGroupName` only suggests. A
 *     hand-named group reads as absent, so gating here would stop an admin who had already made it.
 *   - `addTerm` / `renameTerm` match a typed subject against the tree. The subject is optional and the
 *     match folds fullwidth ＆, case and zero-width characters — but a genuine spelling difference
 *     still misses, and that must cost help, never progress.
 *   - `folderAccess` is a definitive read, but an admin may deliberately map groups after building the
 *     folders. Ordering is advice there, not a prerequisite.
 *
 * That leaves the three where `todo` means the next step CANNOT work: no segment row to code against,
 * a term with no code that reconciliation would silently skip, and uploads still running against the
 * folders a migration is about to move.
 *
 * ⚠ `resumeUploads` WAS DELIBERATELY ABSENT UNTIL 2026-09-08, and the reason it was excluded no longer
 * holds. It read: *its `todo` is the state an admin ARRIVES in, so gating it would trap them on the
 * closing step with no way to finish.* Both halves of that have since become false:
 *   - **The Back band is an open exit.** It calls the same `leaveFlow` and is held only by a run in
 *     flight or an unsaved edit, so nobody is trapped in a flow by a gate on Finish.
 *   - **The toggle that satisfies the gate is ON that step**, one click away — unlike a gate whose fix
 *     lives on another screen.
 * And the failure it prevents is not hypothetical: uploads were off site-wide from 2026-09-07 to
 * 2026-09-08 because this step was never reached, so nobody could file anything for a day and a half.
 * Client, on seeing it: *"I think its best you enforce the upload to be turn on then only finish or
 * not everyone cannot upload."* The old reason is kept above rather than deleted, so the same argument
 * is recognisable if anyone makes it again.
 */
const NEXT_GATED_STEPS: Record<string, string> = {
  createSegment:
    "Create the segment first — fill in this form and press Create. Until the segment exists there is " +
    "nothing for the next steps to point at.",
  abbreviations:
    "Some terms still have no folder code. Reconciliation skips those silently and creates no folder " +
    "for them, so finish here first.",
  /* ⚠ ADDED 2026-09-07, AFTER THIS EXACT STEP WAS SKIPPED ON A LIVE SITE AND COST THREE DAYS.
     GHO's structure change was staged on 2026-09-04; the migration was started at 16:36 and a file
     arrived at 16:37 — one minute in, into a folder the scan had already walked past. Four people went
     on filing into the segment for three days, so `2024` was never empty when the tidy ran, the fresh
     scan always found new drift, and the chain could never apply. `uploadsPaused` read `no` throughout.

     The screen already SAID so, in red: "Uploads must be turned off before continuing." It was a
     warning, and the flow let the admin walk straight past it. This is the one preparation step whose
     omission does not merely risk a worse outcome — it makes every later step in the flow futile, on a
     run that reports success every time. Nothing else in this flow has that property, which is why
     this is a gate and the rest are advice. */
  pauseUploads:
    "Uploads are still switched on. Turn them off above before continuing — a document uploaded during " +
    "the migration lands in the old folder shape, and one that arrives after its folder has been " +
    "scanned is never moved, which leaves the change stuck pending however many times you run it.",
  /* ⚠ ENFORCED AT THE CLIENT'S REQUEST (2026-09-08), AND NOT BECAUSE THE STEP IS TECHNICALLY
     REQUIRED — they were explicit: *"I actually told them its needed even though it isn't but its
     good to run it anyways"*. Reconciliation reads no `Levels` and manages no below-Unit folder, so
     a structure change does not depend on it; what it does is assert the columns, the grants and the
     page ACLs, which is worth doing after any of these flows.

     Gated on a run having FINISHED, never on one having SUCCEEDED — see `reconcileRan`. */
  reconcile:
    "Folder Reconciliation has not been run yet. Press Run on the panel above — it re-checks this " +
    "segment's folders, columns and group access, and it is safe to run as many times as you like.",
  /* Gated through the ordinary `stepState` fallthrough, which already reads this fact in the mirror
     direction: `todo` while uploads are still paused, and `unknown` — which never gates — when the
     config row could not be read. So an unreadable setting cannot hold the closing step. */
  resumeUploads:
    "Uploads are still switched off, so nobody can file a document. Turn them back on above before " +
    "finishing — a forgotten pause leaves the whole site quietly refusing uploads behind a banner " +
    "that makes it look deliberate.",
};

/**
 * Why the Next button should be disabled on this step. Blank means enabled.
 *
 * Enabled whenever the fact is UNKNOWN, exactly as `isLocked` is — a config list that could not be read
 * must never trap someone mid-flow. So this can only ever fire on a positive "not done yet".
 */
export function blocksNext(step: FlowStep, facts: FlowFacts): string {
  if (!step) return "";
  const f = facts ?? {};

  /* ⚠ CHECKED BEFORE `NEXT_GATED_STEPS`, AND DELIBERATELY NOT KEYED ON A STEP ID. "No segment
     picked" is not a property of one step — it is true of every step that spends a segment, so an
     entry in that allow-list could only ever cover one of them.

     Reported by the client on 2026-09-08, on the migrate step: Next was clickable with the picker
     still reading "Select a segment...". The picker REPLACES that step's whole content, so there is
     no scan, no Rebuild and no Apply behind it — and walking on reaches step 5, which turns uploads
     back on over a chain that was never applied. Nothing is written by pressing it; the cost is that
     the flow lets an admin finish without doing the one thing it exists for, on a run that reports
     success. That is the same property that made `pauseUploads` a gate rather than advice.

     `undefined` never gates, so an unreadable segment list cannot strand anyone — see
     `segmentChosen`. And because `firstBlockedStepIndex` reads this function, the rail is capped by
     the same rule in the same breath: gating Next alone would leave the side panel as a second
     route straight past it, which is how three separate holes were shipped on 2026-09-07. */
  if (f.segmentChosen === false && stepUsesSegment(step)) {
    /* ⚠ NAMES THE PICKER, NEVER "jump straight on from the list of steps". That escape hatch was
       true while the rail navigated freely and became false when it was locked forward; a gate that
       points at a control which no longer works is worse than one with no advice at all. */
    return (
      "Choose a segment above to carry on. This step works on one segment at a time, so until one " +
      "is picked there is nothing here to scan or move."
    );
  }

  const reason = NEXT_GATED_STEPS[step.id];
  if (!reason) return "";

  /* ⚠ ANSWERED HERE RATHER THAN THROUGH `stepState`, deliberately. That function also feeds
     `firstIncompleteStep`, which decides where a flow OPENS — teaching it that reconcile is "todo"
     until a run happens would open four of the five flows on their last step. */
  if (step.id === "reconcile") return f.reconcileRan === false ? reason : "";
  if (step.id === "abbreviations" && f.abbreviationsLoading === true) {
    return "Still reading the term store — the codes are being checked. This clears on its own.";
  }
  if (step.id === "createSegment") {
    // Already created — nothing to say, whatever else is unknown.
    if (f.segmentExists === true) return "";
    // Name not typed yet, on a site whose segment list WAS readable. Blank is not "unknown" here: it is
    // the admin not having answered, and no failed read can produce it (see `subjectGiven`). Without
    // this branch Next stayed enabled on a step visibly not done — reported twice by the client.
    if (f.subjectGiven === false) {
      /* ⚠ THE ESCAPE HATCH THIS USED TO NAME IS GONE. It ended "— or jump straight on from the list
         of steps", which was true while the rail navigated freely and became FALSE on 2026-08-30
         when the client asked for it to be locked forward. A gate that tells somebody to use a
         control that no longer works is worse than a gate with no advice at all: they try it, it
         does nothing, and the page reads as broken.

         What replaces it is the escape that still exists — Refresh list. The segment list is read at
         MOUNT and Create happens after it, so a segment made moments ago is genuinely missing from
         its own confirmation until the list is re-read. That is the case this branch is for. */
      return (
        "Create the segment above to carry on. If you created it earlier, pick it from the list " +
        "below — press Refresh list if it is not there yet."
      );
    }
  }
  return stepState(step, facts) === "todo" ? reason : "";
}

/**
 * How far along a step is.
 *
 * `unknown` is a first-class answer, rendered as "not checked" — never as done, never as outstanding.
 * Steps whose completion is unknowable (the term store, Power Automate) stay `unknown` unless a subject
 * makes them checkable.
 */
/**
 * Facts answered at SEGMENT scope, which a subject-scoped flow must not tick from.
 *
 * ⚠ THE DEFECT THIS EXISTS FOR, found on site 2026-08-20. `groupsExist`, `foldersExist` and the
 * since-retired `folderAccessRows` each ask "does this SEGMENT have any at all" — any group with the
 * segment's code prefix, any Folder Map row, any Group Map row. That is right for "Add a new segment",
 * where the answer starts false and turns true as the flow is walked.
 *
 * It is the WRONG question for "Add a department or unit". GHO has 308 groups, 790 mapping rows and
 * 67 folder rows, so all three answer yes the moment the segment is picked — and the flow whose whole
 * purpose is adding something NEW reported three of its five steps Done before the admin had done
 * anything. The fact was true; the granularity was wrong.
 *
 * Same shape as the scoped-reconciliation data loss: a question that is correct site-wide and wrong
 * once something narrows the subject. **Whenever a flow narrows what it is about, re-ask whether the
 * facts still answer the new question.**
 *
 * They are DROPPED, not re-scoped, because the honest answer is available and the correct one is not.
 * Scoping `groupsExist` to a new unit needs its abbreviation, which does not exist until a later step
 * of this very flow; scoping `foldersExist` needs the new term's GUID, which is only known when the
 * optional subject box was filled in. `unknown` renders as "Not checked", which is already this
 * codebase's first-class answer for a fact it cannot establish — and understating costs an admin a
 * glance at a screen, where overstating tells them they are finished when nothing was provisioned.
 */
const SEGMENT_SCOPED_FACTS: ReadonlyArray<keyof FlowFacts> = [
  "groupsExist",
  "foldersExist",
];

/**
 * Strip facts that do not answer the question THIS flow is asking.
 *
 * Applied once, where the facts are handed to the rail, so no caller has to remember. A flow with an
 * `add`/`rename` subject is about one new term inside an existing segment; every other flow is about
 * the segment itself, and keeps them.
 *
 * `abbreviationsMissing` is deliberately NOT stripped: it is reported live by the abbreviation screen
 * over the whole tree, so a newly added term with no code makes it non-zero and the step correctly
 * reads `todo`. It is the one fact here that is already subject-aware.
 */
export function scopeFactsToFlow(flow: Flow | undefined, facts: FlowFacts): FlowFacts {
  const subjectScoped = flow?.asksSubject === "add" || flow?.asksSubject === "rename";
  if (!subjectScoped) return facts ?? {};
  const out: FlowFacts = { ...(facts ?? {}) };
  for (const k of SEGMENT_SCOPED_FACTS) delete out[k];
  return out;
}

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
    case "reconcile":
      return f.foldersExist === undefined ? "unknown" : f.foldersExist ? "done" : "todo";
    case "levels":
      return f.pendingLevels === true ? "done" : "unknown";
    case "pauseUploads":
      // Done when uploads ARE paused — that is what this step exists to achieve.
      return f.uploadsPaused === undefined ? "unknown" : f.uploadsPaused ? "done" : "todo";
    case "resumeUploads":
      // The MIRROR: done when uploads are back ON. The closing step is the one that gets forgotten,
      // and a rail that ticks it while the site is still paused would confirm the wrong thing.
      return f.uploadsPaused === undefined ? "unknown" : f.uploadsPaused ? "todo" : "done";
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

/**
 * Does this step actually USE the segment the flow is about?
 *
 * The segment picker stands in front of every step that has a subject — without it those screens
 * have none, and half of them would show the wrong one. But standing it in front of a step that
 * makes no use of a segment hides the one thing that step exists to say, and asks for a value it
 * cannot spend.
 *
 * ⚠ THIS IS A RULE, NOT A LIST OF SCREEN KINDS, because the literal `kind !== "outside"` it replaces
 * was already wrong once. The client reported it on 2026-08-19 for instruction-only steps (*"weird
 * the first step is asking for the business segment when it is suppose to be showing only a sign to
 * users"*), it was fixed by exempting `outside` — and the very next step added to a flow, the
 * site-wide upload pause, was a `component` and walked straight back into it. Anything that does not
 * consume a segment belongs here, whatever kind of screen it renders.
 */
export function stepUsesSegment(step: FlowStep): boolean {
  // Instructions for work done elsewhere: the term store, Power Automate.
  if (step.screen.kind === "outside") return false;
  // The upload pause is SITE-WIDE by design — there is no per-segment pause to choose.
  if (step.screen.kind === "component" && step.screen.id === "pauseUploads") return false;
  // The Folder levels screen LISTS every segment with its own Edit button, so a picker in front of
  // it asks for a choice the screen then asks for again — and the client saw both at once
  // (2026-08-19: *"I choose Group Head Office and it shows two"*). Picking here changed nothing
  // about what that screen displayed, which reads as the choice having been ignored.
  if (step.screen.kind === "tab" && step.screen.tab === "structure") return false;
  return true;
}

/**
 * The lowest step whose Next is currently blocked, or `steps.length` when none is.
 *
 * Nothing PAST this index may be reached from the rail. Extracted from `FolderAdmin` on 2026-09-07
 * because the rule lived inline in the component, where nothing could test it — and the first
 * version of it shipped with a gap nobody could have caught by reading the JSX.
 */
export function firstBlockedStepIndex(steps: FlowStep[], facts: FlowFacts): number {
  const list = steps ?? [];
  for (let i = 0; i < list.length; i++) {
    if (blocksNext(list[i], facts).length > 0) return i;
  }
  return list.length;
}

/**
 * May the rail navigate to step `i`?
 *
 * THREE conditions, and each is load-bearing:
 *
 *   - `i <= maxIdx` — earned progress (1.0.332.0). The rail never jumps ahead of where the admin has
 *     actually walked.
 *   - `i <= firstBlocked` — the gate. Without it the rail is a SECOND ROUTE PAST IT, which is what the
 *     client found within the hour of the uploads gate shipping: uploads switched back on, Next
 *     correctly refusing, and steps 2 and 3 still clickable because `maxIdx` remembered them.
 *   - `i <= idx` — everything at or behind where the admin stands stays reachable. Without it a fact
 *     turning false strands them on a later step, unable even to go BACK to the one that needs
 *     fixing. That is the opposite of what the gate wants.
 *
 * ⚠ THE THIRD CONDITION HAS A KNOWN, ACCEPTED COST, AND IT IS WORTH UNDERSTANDING BEFORE CHANGING IT.
 * At the END of the structure flow the admin turns uploads back ON — which is the correct final
 * state, and which makes step 1 block again. Stepping BACK from there then narrows what is reachable
 * to where they now stand, so the already-walked steps ahead grey out. They are not trapped: Next
 * still advances, because `levels`/`migrate`/`reconcile` are not themselves gated. The rail is simply
 * more conservative than Next in that one window.
 *
 * Widening it (say, `i <= maxIdx` alone once a step has been visited) would restore the original
 * bypass, since the bypass WAS a visited step. Narrowing it (dropping `i <= idx`) strands people.
 * This is the middle, and the tests below pin all three behaviours.
 */
export function isStepReachable(
  i: number,
  at: { maxIdx: number; idx: number; firstBlocked: number },
): boolean {
  return i <= at.maxIdx && (i <= at.firstBlocked || i <= at.idx);
}
