// Folder Management — guided flows, with the five tabs kept behind "All tools".
//
// Spec: docs/superpowers/specs/2026-08-14-folder-management-guided-flows-design.md
//
// The client's complaint: the tab bar "is confusing and client doesnt know how it works". The tabs were
// already in the order the work happens, which was not enough — five equal doors do not say that four of
// them are steps of one job, and the order is what breaks. Miss Term Abbreviations and reconciliation
// creates NOTHING, silently.
//
// THE GOVERNING RULE: *"the flow should not stop them from doing the work."* Attributed to the client
// until 2026-08-17, when Clarence corrected it — it was never theirs, so it is a design rule of ours and
// can be traded off. The actual ask is *"ensure this flow is working properly and it should make them
// understand how it work"*, which is narrower and stronger: a step definitively not done must SAY so and
// hold the Next button, while a step we merely have not checked must not block anyone. The rules live in
// shared/folderFlows.ts, where `isLocked` keeps any unknown fact — unread, or a read that failed — from
// locking a step; that guards against a WRONG check, not against checking.
//
// It DRIVES FolderManager rather than dismantling it: reconciliation lives inline in a 4,000-line file and
// is the most site-verified code in the project, so a navigation change must not go near it. Steps
// FolderManager hosts render as `<FolderManager initialTab hideTabs />`; Group Management and Folder
// Access, which it does not host, are mounted straight from userAccess — one component, two mount points,
// never a second copy.
import * as React from "react";
import { UploadPauseToggle } from "../../userAccess/components/UploadPauseToggle";
import { stepUsesSegment, scopeFactsToFlow } from "../../../shared/folderFlows";
import {
  UPLOAD_PAUSE_SETTING,
  uploadsArePaused,
} from "../../../shared/uploadPause";
import { useEffect, useState, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import FolderManager from "./FolderManager";
import { IFolderManagerProps } from "./IFolderManagerProps";
import GroupManager from "../../userAccess/components/GroupManager";
import BulkGroupProvisioner from "../../userAccess/components/BulkGroupProvisioner";
import {
  FLOWS,
  Flow,
  FlowFacts,
  FlowStep,
  blocksNext,
  firstBlockedStepIndex,
  firstIncompleteStep,
  isLocked,
  isStepReachable,
  lockReason,
} from "../../../shared/folderFlows";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { tabFromHash } from "../../../shared/adminPages";
import { BackBand, BackToSettings } from "../../../shared/backToSettings";
import { FlowIcon } from "./flowIcons";

/**
 * What the URL asks for, read once on mount.
 *
 * `#flow=addUnit` opens a flow directly. `#tab=reconciliation` opens **All tools** — that is what the
 * hash meant before this screen gained flows, and a bookmark or an older landing-page link must still
 * land on the screen it named rather than on a picker that ignores it. FolderManager reads the same hash
 * itself, so the tab is honoured without being passed along.
 *
 * Anything unrecognised falls through to the picker, never to a blank screen.
 */
function readHash(): { flowId: string; wantsTabs: boolean } {
  const hash = typeof window === "undefined" ? "" : window.location.hash;
  const m = /[#&]flow=([^&]+)/.exec(hash);
  let flowId = "";
  if (m) {
    try {
      flowId = decodeURIComponent(m[1]).trim();
    } catch {
      flowId = "";
    }
  }
  return { flowId, wantsTabs: tabFromHash(hash).length > 0 };
}

/** A segment as the picker needs it. */
type Segment = {
  key: string;
  label: string;
  code: string;
  termSetGuid: string;
  itemId: number;
  /**
   * Whether this segment carries a staged structure change.
   *
   * `undefined` means NOT KNOWN - the column does not exist yet, or its read failed - and it must
   * stay distinct from `false`. `false` locks the migrate step; `undefined` locks nothing, which is
   * this codebase's rule for a fact that could not be read.
   */
  pendingLevels?: boolean;
  /**
   * Whether this segment already has FOLDERS, i.e. reconciliation has run for it.
   *
   * ⚠ THREE-STATE. `undefined` means the Folder Map could not be read, and the picker then shows the
   * segment — hiding one on the strength of a failed read would take away the only way to resume an
   * unfinished segment.
   */
  built?: boolean;
};

const s: Record<string, React.CSSProperties> = {
  /* The Upload Form's page shell - see the note in AuditLog.tsx. This page had NO container at
     all: no centring, no padding, hard against both edges.
     WARN: DELIBERATELY NO `maxWidth`. It hosts the widest content in the product - the
     reconciliation run panel, the migration scan and the group tables all render inside it - and
     capping the column would put a horizontal scrollbar on the screens that most need width. The
     padding alone answers what was asked: it no longer sticks to the wall. */
  /* ⚠ THE PADDING LEFT THIS OBJECT AND LIVES IN `SHELL_CSS` BELOW, because an inline style cannot be
     overridden by a class and the client asked for the side padding to go on a phone and stay on a
     desktop. Everything else stays inline. */
  wrap: {
    fontFamily: '"Segoe UI", system-ui, sans-serif',
    color: "#242424",
    margin: "32px auto",
  },
  h2: { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 6px" },
  sub: { fontSize: 13, color: "#5f5f5f", margin: "0 0 20px", lineHeight: 1.5 },
  /* THE CARD GRID (client design, 2026-08-30). Four across on a wide screen, and it reflows rather
     than scrolling: `auto-fit` with a `min()` floor keeps a card readable on a phone instead of
     squeezing five into the width. */
  cards: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 250px), 1fr))",
    gap: 16,
    alignItems: "stretch",
  },
  /* `display: flex` + `column` is what lets the step-count row sit on the BOTTOM of every card
     whatever the blurb's length — cards of different text lengths otherwise end with the meta line
     at a different height each, which reads as misalignment rather than as different content. */
  card: {
    display: "flex",
    flexDirection: "column",
    textAlign: "left",
    border: "1px solid #e6e6e6",
    borderRadius: 12,
    background: "#fff",
    padding: "20px 18px 16px",
    cursor: "pointer",
    font: "inherit",
  },
  cardIcon: { marginBottom: 14 },
  cardTitle: { fontSize: 15, fontWeight: 600, margin: 0, color: "#242424" },
  cardBlurb: {
    fontSize: 12.5,
    color: "#616161",
    margin: "6px 0 0",
    lineHeight: 1.5,
    flex: "1 1 auto",
  },
  /* The footer: step count left, arrow right — `marginTop: auto` is what pins it down. */
  cardFoot: {
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    justifyContent: "space-between",
    marginTop: 16,
  },
  cardMetaMuted: { fontSize: 11.5, color: "#8a8886" },
  cardGo: {
    width: 30,
    height: 30,
    borderRadius: 8,
    border: "1px solid #cfe3d5",
    background: "#fff",
    color: "#0f6c3f",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  /* Retiring is a BANNER, not a card (client design): it is the one destructive action here, and a
     fifth tile in the same grid invites it to be clicked as casually as the others. */
  cardDanger: {
    display: "flex",
    flexWrap: "wrap",
    alignItems: "flex-start",
    gap: 10,
    width: "100%",
    textAlign: "left",
    border: "1px solid #f3c9cb",
    borderRadius: 10,
    background: "#fdf4f4",
    padding: "14px 16px",
    cursor: "pointer",
    font: "inherit",
  },
  dangerName: {
    fontSize: 15,
    fontWeight: 600,
    color: "#a4262c",
    marginRight: 8,
  },
  dangerBlurb: { fontSize: 12.5, color: "#605e5c", lineHeight: 1.5 },
  /* The card arrow, in the danger palette (client, 2026-09-04: *"Add button for retire a segment as
     well"*). The banner was ALREADY a `<button>` — the whole strip has always been clickable — but it
     was the only entry on this page with no arrow, so it read as a notice rather than a way in.
     ⚠ IT IS STILL ONE BUTTON, NOT A BUTTON INSIDE A BANNER. A nested `<button>` is invalid HTML and
     the inner one swallows the click, which is the trap the Requests accordion header already hit.
     This is a `<span>` styled to match `cardGo`, so the whole strip stays the hit target. */
  dangerGo: {
    width: 30,
    height: 30,
    borderRadius: 8,
    border: "1px solid #f1b0b3",
    background: "#fff",
    color: "#a4262c",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
    marginLeft: "auto",
    alignSelf: "center",
  },
  sectionHead: {
    fontSize: 12,
    fontWeight: 600,
    color: "#605e5c",
    textTransform: "uppercase",
    letterSpacing: ".04em",
    margin: "26px 0 10px",
  },
  // The "All tools ›" link at the foot of the picker. It borrows the back band's TEXT style
  // (`backLinkStyle` in shared/backToSettings) without the band, because it is a descent into a
  // deeper screen rather than an exit — banding it would make two opposite moves look identical.
  back: {
    border: "none",
    background: "transparent",
    padding: 0,
    font: "inherit",
    color: "rgba(0, 104, 74, 1)",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
  },
  // FLEX, not `repeat(auto-fit, …)`. That grid made as many 240px columns as would fit, so on a wide
  // screen it produced a third empty column and the panel — spanning two — left a band of dead space to
  // its right (client, 2026-08-17: "CAn you remove the padding?"). Flex gives the same wrapping on a
  // narrow screen with no phantom tracks on a wide one.
  runner: {
    display: "flex",
    flexWrap: "wrap",
    gap: 20,
    alignItems: "flex-start",
  },
  /* THE STEPPER (client design, 2026-08-30). White card rather than the old grey block, so it reads
     as a panel beside the work rather than a sidebar behind it. `position: relative` anchors the
     connector line drawn between the numbers. */
  rail: {
    position: "relative",
    border: "1px solid #e6e6e6",
    borderRadius: 12,
    background: "#fff",
    padding: "18px 16px",
    flex: "0 1 260px",
  },
  railItem: {
    position: "relative",
    display: "flex",
    flexWrap: "wrap",
    gap: 12,
    width: "100%",
    textAlign: "left",
    border: "none",
    background: "transparent",
    font: "inherit",
    padding: "10px 8px",
    borderRadius: 8,
    cursor: "pointer",
    alignItems: "center",
  },
  railActive: { background: "#eef7f1" },
  /* The line joining one number to the next. Drawn per ITEM rather than as one line behind the
     list, because the last step must not trail a stub below it — which is exactly what a single
     absolutely-positioned line would do. */
  railLink: {
    position: "absolute",
    left: 21,
    top: 34,
    width: 2,
    bottom: -6,
    background: "#e2e2e2",
  },
  /* The number pill used to take its colour from the step STATE (done / todo / unknown). With the
     state gone it needs a resting look of its own — `s` is a `Record<string, CSSProperties>`,
     so a key that does not exist renders unstyled with a green build. */
  railNumIdle: { background: "#eaeaea", color: "#8a8886" },
  /* `zIndex` lifts the number above the connector line so the line appears to run BETWEEN the
     circles rather than through them. */
  railNum: {
    position: "relative",
    zIndex: 1,
    flexShrink: 0,
    width: 26,
    height: 26,
    borderRadius: "50%",
    fontSize: 12,
    fontWeight: 700,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: "#0f6c3f",
    color: "#fff",
  },
  railLabel: {
    fontSize: 13,
    fontWeight: 600,
    lineHeight: 1.35,
    color: "#242424",
  },
  railState: { fontSize: 11, lineHeight: 1.4, marginTop: 2, display: "block" },
  // `flex: 1 1 420px` — takes the rest of the row, wraps under the rail below ~700px. The old
  // `gridColumn: span 2` belonged to the auto-fit grid above and is what left the dead band.
  panel: { minWidth: 0, flex: "1 1 420px" },
  stepHead: { fontSize: 18, fontWeight: 600, margin: "0 0 4px" },
  stepHint: {
    fontSize: 13,
    color: "#5f5f5f",
    margin: "0 0 16px",
    lineHeight: 1.55,
    /* A step hint may carry its own line breaks (client, 2026-09-06 - the abbreviations hint puts
       "All fields marked * are mandatory." on its own line). `pre-line` honours a newline while
       still collapsing ordinary wrapping whitespace, unlike `pre`, which would also preserve the
       indentation of every concatenated string literal in `folderFlows.ts`. */
    whiteSpace: "pre-line",
  },
  /* ⚠ GREEN, NOT BLUE, since 2026-08-30 (client: *"keep the old design but tweak the color to match
     the other designs"*). The CONTENT of an `outside` step is unchanged and deliberately so — their
     mock replaced step 1 with a Create Term Set form, and **this app cannot create a term set or its
     terms**; the client's own answer was *"we do not have the time to built that interface"*. So the
     step still says the work happens in the term store, in the palette the rest of the page uses. */
  termStoreLink: {
    display: "inline-block",
    fontWeight: 600,
    color: "#0f6c3f",
    textDecoration: "underline",
  },
  termStoreNote: {
    display: "block",
    fontSize: 11.5,
    color: "#5f7a6b",
    marginTop: 4,
    lineHeight: 1.5,
  },
  outside: {
    border: "1px solid #cfe3d5",
    background: "#f2f9f5",
    borderRadius: 10,
    padding: "14px 16px",
    fontSize: 13,
    color: "#1c4d33",
    lineHeight: 1.6,
  },
  lockBox: {
    border: "1px solid #f2c9a0",
    background: "#fff8f0",
    borderRadius: 8,
    padding: "12px 14px",
    fontSize: 12.5,
    color: "#8a4b00",
    lineHeight: 1.55,
    marginBottom: 14,
  },
  label: { display: "block", fontWeight: 600, fontSize: 12, margin: "0 0 4px" },
  /* The always-visible segment switcher, above a step's own content. Separated by a rule rather than
     boxed, so it reads as a setting for the step below rather than as part of it.
     ⚠ `s` IS A `Record<string, CSSProperties>`: a key that does not exist yields `undefined` and the
     element renders unstyled with a GREEN BUILD. Every style referenced in the JSX must be here. */
  segSwitch: {
    margin: "0 0 16px",
    paddingBottom: 14,
    borderBottom: "1px solid #ebebeb",
  },
  select: {
    width: "100%",
    maxWidth: 420,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
  },
  input: {
    width: "100%",
    maxWidth: 420,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
  },
  railBusy: { cursor: "not-allowed", opacity: 0.55 },
  navBar: {
    display: "flex",
    gap: 8,
    marginTop: 20,
    paddingTop: 16,
    borderTop: "1px solid #eceaea",
    flexWrap: "wrap",
  },
  primary: {
    padding: "7px 16px",
    fontSize: 13,
    background: "#0f6c3f",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    cursor: "pointer",
  },
  off: {
    padding: "7px 16px",
    fontSize: 13,
    background: "#e6e6e6",
    color: "#9a9a9a",
    border: "none",
    borderRadius: 4,
    cursor: "not-allowed",
  },
  ghost: {
    padding: "7px 16px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  danger: {
    marginBottom: 16,
    padding: "10px 12px",
    border: "1px solid #f1b0b3",
    background: "#fdf3f4",
    borderRadius: 8,
    fontSize: 12.5,
    color: "#a4262c",
    lineHeight: 1.5,
  },
  allTools: { marginTop: 30, paddingTop: 16, borderTop: "1px solid #eceaea" },
  modeBar: { display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" },
  modeOn: {
    padding: "7px 14px",
    fontSize: 13,
    fontWeight: 600,
    background: "#0f6c3f",
    color: "#fff",
    border: "1px solid #0f6c3f",
    borderRadius: 4,
    cursor: "pointer",
  },
  modeOff: {
    padding: "7px 14px",
    fontSize: 13,
    background: "#fff",
    color: "#1b1b1b",
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    cursor: "pointer",
  },
  hint: { fontSize: 11.5, color: "#5f6f80", marginTop: 6, lineHeight: 1.5 },
  doneBox: {
    border: "1px solid #c6e3d1",
    background: "#f1f8f4",
    borderRadius: 8,
    padding: "12px 14px",
    fontSize: 13,
    color: "#0f6c3f",
    lineHeight: 1.55,
    marginBottom: 14,
  },
};

/* ⚠ REMOVED 2026-08-30 with the per-step status, at the client's request. `stepState` and
   `remainingCount` in `shared/folderFlows.ts` are UNTOUCHED and still tested — only this screen
   stopped rendering them. That matters: `isLocked` and `blocksNext` read the same `FlowFacts`, so
   the four real locks (segment exists, PendingLevels set, abbreviations complete) still work and
   still explain themselves beside the Next button. What went is the GRADING of every step, not the
   checking of the ones that can actually be checked. Re-render it by restoring this map. */

/* Client, 2026-09-08: *"ensure that section padding for mobile is gone but for desktop it is there"*.
   24px each side of a 360px screen is 13 percent of it spent on nothing, and the cards inside carry
   their own.

   ⚠⚠ @media HERE, NOT @container — AND THAT IS A DELIBERATE DEPARTURE FROM THE UPLOAD FORMS.
   A container query would need `container-type` on an ancestor of this section, and container-type
   applies LAYOUT CONTAINMENT, which makes that element the containing block for every
   `position: fixed` descendant. Three render inside this section: StructureManager's discard dialog,
   SubtreeMigrator's rename dialog and FolderManager's own modal. Containment would shrink all three
   from covering the window to covering the section.

   The cost, stated rather than discovered: a media query measures the WINDOW, so this does NOT fire
   in SharePoint's own Mobile preview, which narrows the content column and leaves the viewport at
   desktop width. It DOES fire on a real phone and in browser device emulation. If this page ever
   needs to respond to a narrow SECTION as well, the dialogs have to move out of the container first.

   NO BACKTICKS ANYWHERE IN THIS BLOCK - it is a JS template literal and one ends it. */
const SHELL_CSS = `
  .crs-shell { padding: 0 24px 48px; }
  @media (max-width: 480px) {
    .crs-shell { padding-left: 0; padding-right: 0; }
  }
`;

/* One definition of the page shell, used by all three of this component's roots — the flow picker,
   the flow runner and All tools. Rendering the style tag inside it means the CSS travels with
   whichever root is mounted, and there is never more than one. */
function Shell({ children }: { children: React.ReactNode }): React.ReactElement {
  return (
    <section style={s.wrap} className="crs-shell">
      <style>{SHELL_CSS}</style>
      {children}
    </section>
  );
}

export default function FolderAdmin({
  context,
}: IFolderManagerProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;

  // Read once, synchronously, so an addressed screen never flashes the picker first.
  const [flow, setFlow] = useState<Flow | undefined>(
    () => FLOWS.filter((f) => f.id === readHash().flowId)[0],
  );
  const [stepIdx, setStepIdx] = useState(0);
  /* ⚠ THE FURTHEST STEP REACHED, and the rail needs it — not the CURRENT one.
     The rail became forward-locked on 2026-08-30. Gating it on the current index alone meant that
     stepping BACK to check something also threw away the way forward: an admin on step 4 who looked
     at step 1 would have had to press Next three times to get back, on screens they had already
     finished. That is the "confusing" the client was worried about, not the locking itself.
     Reset with the flow, so a new flow starts closed again. */
  const [maxIdx, setMaxIdx] = useState(0);
  /**
   * The abbreviation screen's own `save()`, while that screen is mounted.
   *
   * ⚠ A REF, NOT STATE. It is re-registered on EVERY render of that screen (its closure has to stay
   * current over `rows`), and holding it in state would set state during render — an update loop.
   * Nothing renders from it; only the Next handler reads it.
   */
  const abbrevSaveRef = useRef<(() => Promise<boolean>) | undefined>(undefined);
  /** True while that save is in flight, so Next cannot be pressed twice. */
  const [abbrevSaving, setAbbrevSaving] = useState(false);
  const [allTools, setAllTools] = useState(() => readHash().wantsTabs);

  const [segments, setSegments] = useState<Segment[] | undefined>(undefined);
  const [segKey, setSegKey] = useState("");
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  /**
   * Bumped to re-read the segment list.
   *
   * The list was read once on mount, so a segment created on the New segment step never appeared here:
   * the rail said "Not checked" and Next stayed disabled while the panel beside it said the segment was
   * configured (client, 2026-08-17: *"Weird, I created the Segment but it is showing this."*). A screen
   * that contradicts itself is worse than one that says nothing.
   */
  const [reload, setReload] = useState(0);
  /* ⚠ PURELY SO THE PRESS IS VISIBLE (client, 2026-09-04: *"Button is not clickable"* against the
     Refresh list button). It always WAS clickable — no `disabled`, a live `onClick` — but a press
     re-reads a list that usually comes back identical, so nothing on screen moves and the button
     reads as dead. Exactly the same complaint they raised in the same message about the audit log's
     "History of this file" link, which genuinely did nothing.
     The read is fast, so this is mostly a flicker; what matters is that SOMETHING happens. */
  const [refreshing, setRefreshing] = useState(false);
  /**
   * Whether the New segment FORM is open on step 2.
   *
   * Closed by default and closed again once a segment is confirmed, because a creation form still sitting
   * under the confirmation of what was just created invites creating it twice (client, 2026-08-17: *"when
   * i create GHO and the Form still showing below it is really weird"*). Duplicating a segment is not
   * harmless either — `validateNewSegment` refuses a duplicate LABEL and a duplicate top folder, but two
   * segments differing only in punctuation still slug to one key, which presents as the new one shadowing
   * the old.
   */
  const [showForm, setShowForm] = useState(false);
  /**
   * Terms with no folder code, as reported by the abbreviations screen once it is open.
   *
   * NOT read here, deliberately: the count needs a walk of the whole term tree (~115 requests for GHO),
   * which is why the facts effect leaves it undefined. The screen that already walked it hands the number
   * over instead. Until it does, `undefined` gates nothing — which is why Next was clickable on a screen
   * covered in "no folder will be created" warnings (client, 2026-08-17).
   */
  const [abbrevMissing, setAbbrevMissing] = useState<number | undefined>(
    undefined,
  );
  /**
   * True while the abbreviations screen is reading. A SEPARATE fact from the count, because the count
   * is `undefined` both while loading and when the read failed — and only the first should hold Next
   * (client, on site 2026-08-18: Next was available while the panel said "Reading the term store…").
   */
  const [abbrevLoading, setAbbrevLoading] = useState(false);
  /**
   * True while a bulk group run is in flight, reported up by `BulkGroupProvisioner`.
   *
   * The run lives in that component's state with no resume, so a step change unmounts it and stops it
   * part-way — SILENTLY, which is the likely cause of 302 of 308 groups on the rehearsal site. So the
   * rail, Back, Next, Finish and the way out of the flow are all held for the duration.
   *
   * This is the one thing in the runner that DOES padlock navigation, and the exception is deliberate.
   * The client's rule is that the flow must not stop them doing the work — here navigation destroys
   * work already in progress, which is the same reasoning that makes a tab switch with unsaved
   * abbreviations a refusal rather than a "discard?" prompt. It is also temporary and self-clearing,
   * and the step itself offers Stop, so nobody is held longer than they choose to be.
   */
  /* ⚠⚠ ONE STATE PER SOURCE, DERIVED INTO ONE PADLOCK — never one shared boolean.
     Three independent runs report into this: reconciliation, the migration scan/move, and a bulk
     group run. With a single `runBusy` the LAST reporter won, and one of them re-asserted its value
     on EVERY RENDER (`onReconRunning` was a fresh function identity each render, and
     `FolderManager`'s effect depends on that identity) — so the migration scan set it true and the
     next render immediately pushed `false` back over it.
     Client, 2026-09-09: *"next is available when I am running this, that is dangerous"*. The
     migrate padlock had therefore NEVER held in the guided flow, on a step whose scan and move both
     run in the page with NO RESUME: pressing Next threw the work away silently.
     Deriving the padlock means a source can only ever report about ITSELF, so no source can release
     another's lock however often it re-reports. */
  const [reconBusy, setReconBusy] = useState(false);
  const [migrateBusy, setMigrateBusy] = useState(false);
  const [groupBusy, setGroupBusy] = useState(false);
  const runBusy = reconBusy || migrateBusy || groupBusy;
  /**
   * A reconciliation run has FINISHED during this visit, so the reconcile step may be left.
   *
   * ⚠ DERIVED FROM THE RUNNING FLAG GOING TRUE THEN FALSE, which is why the ref is needed:
   * `onReconRunningChange` is also called with `false` before anything has run, and treating that
   * as a completion would satisfy the gate without a run.
   *
   * ⚠ FINISHED, NOT SUCCEEDED — see `reconcileRan` in folderFlows. In the structure flow this step
   * sits before the one that switches uploads back ON, so a success requirement would keep the site
   * refusing uploads for as long as reconciliation kept failing.
   */
  const [reconRan, setReconRan] = useState(false);
  const reconStarted = React.useRef(false);
  /* ⚠ `useCallback` IS LOAD-BEARING, NOT TIDINESS. `FolderManager`'s effect lists this callback in
     its deps, so a fresh identity each render re-fires it on every render — which is how a stale
     `false` kept landing on top of another source's `true`. Stable identity, so it fires only when
     `reconRunning` actually changes. Empty deps: it touches only a ref and two setState functions,
     all of which are stable. */
  const onReconRunning = React.useCallback((running: boolean): void => {
    setReconBusy(running);
    if (running) reconStarted.current = true;
    else if (reconStarted.current) setReconRan(true);
  }, []);
  /**
   * The Folder levels screen holds unsaved edits.
   *
   * Worse here than the tab-switch case it already guards: the NEXT step reads `PendingLevels`,
   * which an unsaved edit has not written — so walking on reports "nothing to move" for a change
   * the admin believes they made, and the flow looks broken rather than incomplete.
   */
  const [structureDirty, setStructureDirty] = useState(false);
  /* ⚠ REPORTED BUT NO LONGER GATING (2026-09-06). Nothing reads the flag now that Next saves — the
     state it protected against cannot arise — but the screen still reports it, and dropping the
     receiver would make re-instating a dirty gate a two-file change instead of a one-line one.
     Named `_` so lint sees it consumed without pretending it is used. */
  const [, setAbbreviationsDirty] = useState(false);
  /**
   * The migration screen has scanned work it has not run.
   *
   * The client walked straight past Rebuild on the first run (2026-08-19). Doing so leaves the
   * segment staged and the folders in the old shape, and the NEXT step turns uploads back on over
   * it — so the flow would end with a half-applied structure and people filing into it.
   */
  const [migratePending, setMigratePending] = useState(false);
  /**
   * The segment whose migration finished during THIS visit to the flow.
   *
   * ⚠⚠ IT EXISTS BECAUSE THE MIGRATE STEP'S OWN SUCCESS FIRES ITS OWN LOCK, and that lock REPLACES
   * the screen (`{!locked && renderStep(step)}`) — so the moment a migration applied the chain,
   * `pendingLevels` flipped to false, the lock rendered, the migrator unmounted and **the run's own
   * result log was destroyed**. Reported the first time anyone completed one on 1.0.502.0: *"I can't
   * give you the log, after I finish running it immediately becomes like this."*
   *
   * That log is the only record of what moved, what was tidied, what was tagged and which folders
   * were left as strays — and an admin gets one chance to read it.
   *
   * ⚠ INTRODUCED BY 1.0.497.0's REFRESH SIGNAL. Before `onMigrateApplied` bumped `reload`, the fact
   * stayed stale and the screen survived; making the gate release correctly is what started
   * destroying the evidence. **A lock whose fact the step's own success changes must not hide that
   * step's result.**
   *
   * Keyed to the SEGMENT, and cleared on entering or leaving the flow: it means "a migration just
   * finished here", which is true of one segment for one visit. The switcher can change segment
   * mid-flow, and the previous segment's result must not unlock the new one's lock.
   */
  const [appliedFor, setAppliedFor] = useState<string | undefined>(undefined);

  /**
   * Hold navigation while a run is in flight, and RE-READ THE FACTS when it ends.
   *
   * `groupsExist` is read when the segment is picked, so after a run that created 300 groups the
   * rail still said the step was *To do* (defect 5, 2026-08-18). Cosmetic in the sense that nothing
   * is wrong underneath, and not cosmetic in the sense that matters: the rail is the thing telling
   * an admin what is left, and one entry known to be lying is enough to stop them trusting the rest.
   *
   * The same reason the segment list gained an explicit refresh — a mount-time read reflects nothing
   * a step below it has since written.
   */
  /**
   * Bumped when a bulk run FINISHES, which re-reads the rail's facts and the group list's mapping
   * badges. A separate signal from `busy` going false — "the run ended" is the fact that makes other
   * screens stale, and reading it off the falling edge of a UI flag would tie the two together.
   */
  const onRunComplete = (): void => setReload((n) => n + 1);

  /** The subject of flows 2 and 4 — what makes their term-store step checkable at all. */
  const [subject, setSubject] = useState("");

  // RAW reads only. The derived `effectiveFacts` below layers on the two facts that no list read can
  // answer: the abbreviation count (a term-tree walk, paid for by the screen that needs it anyway) and
  // flow 1's segmentExists (no picker, because the segment does not exist yet).
  const [baseFacts, setBaseFacts] = useState<FlowFacts>({});

  const GET = { Accept: "application/json;odata=nometadata" };

  /**
   * Segments, from the DMS Config mode rows.
   *
   * One read, and it also answers `pendingLevels` — so the structure flow's lock costs nothing extra.
   */
  useEffect(() => {
    const load = async (): Promise<void> => {
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      // PendingLevels IS REQUESTED SEPARATELY, BELOW, AND MUST STAY THAT WAY.
      //
      // The column is created ON DEMAND by StructureManager, the first time a structure change is
      // staged, so on any site where that has never happened it does not exist - and one unknown name
      // in a $select fails the WHOLE request with HTTP 400 rather than returning null (gotcha #11).
      // Asked for here, it took the segment list down with it: every guided flow reported "the segment
      // list could not be read" and offered no segment at all, on a site otherwise perfectly set up.
      // StructureManager and SubtreeMigrator both split this read for exactly this reason; this file
      // did not, and it is the one screen every folder job starts from.
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Id,Title,ModeLabel,StagingFolder,TermSetGuid,ConfigType&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!res.ok) throw new Error(`config HTTP ${res.status}`);
      const data = await res.json();
      const rows = ((data.value ?? []) as Array<Record<string, unknown>>)
        .filter(
          (r) =>
            String(r.ConfigType ?? "")
              .trim()
              .toLowerCase() === "mode",
        )
        .map((r) => ({
          key: String(r.Title ?? "").trim(),
          label: String(r.ModeLabel ?? r.Title ?? "").trim(),
          code: String(r.StagingFolder ?? "").trim(),
          termSetGuid: String(r.TermSetGuid ?? "").trim(),
          itemId: Number(r.Id ?? 0),
        }))
        .filter((r) => r.key.length > 0)
        .sort((a, b) => a.label.localeCompare(b.label));

      // The staged chains, in their own request. A 400 here means the column has never been created,
      // which is neither an error nor anything the admin can act on - so the flag is left UNDEFINED
      // rather than false, and undefined locks nothing.
      const staged: Record<number, boolean> = {};
      let stagedKnown = false;
      try {
        const p2: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
            `?$select=Id,PendingLevels&$filter=ConfigType eq 'mode'&$top=200`,
          SPHttpClient.configurations.v1,
          { headers: GET },
        );
        if (p2.ok) {
          stagedKnown = true;
          for (const r of ((await p2.json()).value ?? []) as Array<
            Record<string, unknown>
          >) {
            staged[Number(r.Id ?? 0)] =
              String(r.PendingLevels ?? "").trim().length > 0;
          }
        }
      } catch {
        // Leave it unknown. A transient failure must not lock the migrate step.
      }
      /* Which segments already have folders, for the "Continuing an earlier segment?" picker.
         ⚠ ONE READ, AND THE SAME ONE THE PER-SEGMENT FACTS EFFECT ALREADY MAKES UNFILTERED
         (`$select=Id,Section&$top=5000`) — a Folder Map row's `Section` holds the segment's top
         folder, so every segment's answer is in one response. `$top=5000`, not 500: a truncated read
         reports a built segment as unbuilt, which is the direction that puts a live segment back in
         a creation flow.
         Left UNDEFINED on any failure, so the picker falls back to listing everything. */
      let builtCodes: string[] | undefined;
      try {
        const fm: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.folderMap))}')/items` +
            `?$select=Id,Section&$top=5000`,
          SPHttpClient.configurations.v1,
          { headers: GET },
        );
        if (fm.ok) {
          const fmRows = ((await fm.json()).value ?? []) as Array<{ Section?: string }>;
          const codes: string[] = [];
          for (const r of fmRows) {
            const c = (r.Section ?? "").trim().toLowerCase();
            if (c && codes.indexOf(c) < 0) codes.push(c);
          }
          builtCodes = codes;
        }
      } catch {
        // Leave it undefined: the picker lists everything rather than hiding on a failed read.
      }
      // Merged only when the second read succeeded. Absent column or failed read leaves every
      // segment `undefined`, which reads as "not checked" and locks nothing.
      setSegments(
        rows.map((r) => ({
          ...r,
          ...(stagedKnown ? { pendingLevels: staged[r.itemId] === true } : {}),
          ...(builtCodes === undefined
            ? {}
            : { built: builtCodes.indexOf(r.code.toLowerCase()) >= 0 }),
        })),
      );
      setLoadError(undefined);
      setRefreshing(false);
    };
    load().catch((e) => {
      setRefreshing(false);
      // Unreadable ≠ none. A failed read must not present as "this site has no segments", which would
      // send an admin off to create one that already exists.
      setSegments(undefined);
      setLoadError((e as Error).message);
    });
  }, [siteUrl, reload]);

  const segment = (segments ?? []).filter((x) => x.key === segKey)[0];
  /**
   * The segments the "Continuing an earlier segment?" picker offers: the ones NOT already built.
   *
   * `built !== true` rather than `built === false`, so an unreadable Folder Map lists everything —
   * and the picked segment is kept whatever its state, or the control would display a different
   * segment from the one the flow is carrying.
   */
  const resumable = (segments ?? []).filter((x) => x.built !== true || x.key === segKey);

  /**
   * The cheap facts for the chosen segment: groups, mappings, folders.
   *
   * Three extra reads, each already performed by some other screen. Abbreviation completeness is
   * deliberately absent — it needs a term-tree walk (~115 requests for GHO) — so it stays `undefined`,
   * and `undefined` never locks. Every read is `.catch`ed to undefined for the same reason: a throttled
   * list must leave a step passable, not stuck.
   */
  useEffect(() => {
    if (!flow) return;
    let cancelled = false;
    const load = async (): Promise<void> => {
      const next: FlowFacts = {};
      const target = segment;
      if (target) {
        next.segmentExists = true;
        next.pendingLevels = target.pendingLevels;

        // Groups, by the naming convention. ADVISORY — `suggestGroupName` only suggests, so a
        // hand-named group reads as absent. This can mark, never lock.
        const groups = await context.spHttpClient
          // $top=5000, not 500: a provisioned segment is ~324 groups on its own, and a truncated
          // read would report an existing segment's groups as absent.
          .get(
            `${siteUrl}/_api/web/sitegroups?$select=Title&$top=5000`,
            SPHttpClient.configurations.v1,
            { headers: GET },
          )
          .then(async (r) =>
            r.ok
              ? (((await r.json()).value ?? []) as Array<{ Title?: string }>)
              : undefined,
          )
          .catch(() => undefined);
        if (groups) {
          const prefix = target.code.toLowerCase();
          next.groupsExist =
            prefix.length > 0 &&
            groups.filter(
              (g) => (g.Title ?? "").toLowerCase().indexOf(prefix) === 0,
            ).length > 0;
        }

        const mapRows = await context.spHttpClient
          .get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items` +
              `?$select=Id,Segment&$top=5000`,
            SPHttpClient.configurations.v1,
            { headers: GET },
          )
          .then(async (r) =>
            r.ok
              ? (((await r.json()).value ?? []) as Array<{ Segment?: string }>)
              : undefined,
          )
          .catch(() => undefined);
        const folderRows = await context.spHttpClient
          .get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.folderMap))}')/items` +
              `?$select=Id,Section&$top=5000`,
            SPHttpClient.configurations.v1,
            { headers: GET },
          )
          .then(async (r) =>
            r.ok
              ? (((await r.json()).value ?? []) as Array<{ Section?: string }>)
              : undefined,
          )
          .catch(() => undefined);
        if (folderRows) {
          next.foldersExist =
            folderRows.filter(
              (r) =>
                (r.Section ?? "").trim().toLowerCase() ===
                target.code.toLowerCase(),
            ).length > 0;
        }
      }
      // The site-wide upload pause, for the two pause steps' ticks. Its own request because the
      // setting row is site-wide and has nothing to do with the segment; `undefined` on failure, so
      // an unreadable config leaves both steps unknown rather than claiming uploads are on.
      try {
        const pRes = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
            `?$select=SettingValue&$filter=ConfigType eq 'setting' and Title eq '${UPLOAD_PAUSE_SETTING}'&$top=1`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (pRes.ok) {
          const pRows = ((await pRes.json()).value ?? []) as Array<{
            SettingValue?: string;
          }>;
          // No row is a valid state meaning "not paused"; only a failed READ is unknown.
          next.uploadsPaused =
            pRows.length > 0 ? uploadsArePaused(pRows[0].SettingValue) : false;
        }
      } catch {
        /* leaves uploadsPaused undefined — unknown, never "on" */
      }
      if (!cancelled) setBaseFacts(next);
    };
    load().catch(() => {
      if (!cancelled) setBaseFacts({});
    });
    return () => {
      cancelled = true;
    };
    // `reload` is in here so a bulk run's results are picked up the moment it ends — see
    // onRunBusyChange. It also refreshes the segment list, which is harmless and occasionally right.
    //
    // ⚠ `idx` IS IN HERE BECAUSE A STEP WRITES WHAT THE NEXT STEP READS. Saving a structure change on
    // step 3 sets `PendingLevels`, which is exactly the fact step 4 is locked behind — but the facts
    // were read when the flow opened, so step 4 said "Not ready yet" beside a screen that had just
    // saved, and step 3 stayed "Not checked" until the page was refreshed (reported 2026-08-19).
    // Fourth instance of the same trap in this feature: a screen that reads at mount lies about any
    // run beside it. Moving between steps is the cheapest honest moment to re-read.
  }, [flow ? flow.id : "", segKey, segments, reload, stepIdx]);

  /**
   * The facts, plus the one fact only the "add a new segment" flow can answer.
   *
   * That flow has no segment picker — the segment does not exist yet — so the effect above never sets
   * `segmentExists`, and `undefined` is never a lock. The result was a padlock and a Next gate that
   * could not fire in the one flow they were written for (found on the client's site 2026-08-17).
   *
   * Answered here by matching the NAME the admin said they were creating against the loaded mode rows,
   * via `labelMatches` so a fullwidth ＆, stray whitespace or a pasted zero-width character does not
   * read as "not created yet".
   *
   * DERIVED AT RENDER, deliberately, rather than added to the effect's dependencies: the effect performs
   * three list reads, and keying it on `subject` would fire all three on every keystroke. `segments` is
   * already in state, so this costs nothing.
   *
   * Three-state on purpose:
   *   - name blank, or the segment list unreadable  -> undefined, so nothing is gated (fail open);
   *   - a row matches                                -> true;
   *   - the list read fine and nothing matches       -> false, the only case that gates.
   */
  const effectiveFacts: FlowFacts = React.useMemo(() => {
    // The abbreviations screen's own count wins whenever it has one, in EVERY flow — four of the five
    // include that step. Spread first so the block below still overrides `segmentExists`.
    // Segment-scoped facts are dropped for a flow about ONE NEW TERM inside an existing segment —
    // otherwise GHO's 308 groups and 790 mapping rows tick three of "Add a department or unit"'s five
    // steps Done before the admin has done anything. Rule lives in folderFlows so it is testable.
    const facts: FlowFacts = {
      ...scopeFactsToFlow(flow, baseFacts),
      // `abbreviationsLoading` is carried ALWAYS, count or no count — it is the fact that tells the
      // gate apart "not read yet" from "read and failed", and only the first holds Next.
      abbreviationsLoading: abbrevLoading,
      ...(abbrevMissing === undefined
        ? {}
        : { abbreviationsMissing: abbrevMissing }),
      /* Has a segment been picked? Set ONLY on a flow that works on one, and ONLY when the list read
         fine — so `undefined` covers both "this flow is not about one segment" and "there is nothing
         to pick from", neither of which may gate anything. See `segmentChosen` in folderFlows. */
      ...(flow?.needsSegment !== true || segments === undefined
        ? {}
        : { segmentChosen: segment !== undefined }),
      /* Set for every flow that HAS the step, and deliberately after `scopeFactsToFlow`: this one is
         about the session rather than the segment, so the add/rename narrowing must not drop it. */
      ...(flow && flow.steps.filter((st) => st.id === "reconcile").length > 0
        ? { reconcileRan: reconRan }
        : {}),
    };
    if (!flow || flow.asksSubject !== "newSegment") return facts;
    // Unreadable list ⇒ change nothing, so nothing is gated. This is the ONLY fail-open case here.
    if (segments === undefined) return facts;
    // The list read fine, so "nothing picked" is the admin not having answered — not a failure.
    if (!segment) return { ...facts, subjectGiven: false };
    return { ...facts, subjectGiven: true, segmentExists: true };
  }, [flow, baseFacts, segments, segment, abbrevMissing, abbrevLoading, reconRan]);

  /** Open a flow on the first thing left to do. */
  /* ⚠ `segKey` IS CLEARED ON BOTH TRANSITIONS (client, 2026-09-06: leaving the structure flow and
     coming back "keeps holding the previous dropdown I selected").
     It survived because it is held HERE, in the host, so that one answer carries into every later
     step instead of each screen asking again — and nothing was resetting it. A stale segment is
     worse than an unanswered one on this flow in particular: the migrate step is pre-selected from
     it, so an admin returning for a DIFFERENT segment would have been shown the previous one already
     chosen, with the picker replaced by a static line saying so.
     Cleared on OPEN as well as on leave: a flow entered from a bookmark or a `#flow=` link never
     goes through `leaveFlow` first. */
  const openFlow = (f: Flow): void => {
    setFlow(f);
    setAppliedFor(undefined);
    setReconRan(false);
    reconStarted.current = false;
    setAllTools(false);
    setSubject("");
    setSegKey("");
    setStepIdx(firstIncompleteStep(f, baseFacts));
  };

  const leaveFlow = (): void => {
    setFlow(undefined);
    setAppliedFor(undefined);
    setReconRan(false);
    reconStarted.current = false;
    setBaseFacts({});
    setSegKey("");
    setStepIdx(0);
    setMaxIdx(0);
  };

  // ── The picker ─────────────────────────────────────────────────────────────
  /**
   * The step the runner is on, and the furthest one reached.
   *
   * BOTH OF THESE MUST STAY ABOVE EVERY EARLY RETURN. The effect below used to sit with the flow
   * runner, which crashed the web part with React error #310, "Rendered more hooks than during the
   * previous render": the picker returns before that point, so the effect ran on the flow path and
   * not on the picker path. Opening a flow then rendered NOTHING, with no error UI at all. Found on
   * site 2026-08-31 - and the comment on the effect already claimed it was "declared with the other
   * hooks, above every early return". It was not. A comment describing the fix is not the fix.
   *
   * THIRD TIME IN THIS PROJECT: the approval panel's ApproveItems probe and My Submissions'
   * share-recipient picker were the same fault. A useEffect below a return, in a component that
   * returns early, is a blank web part waiting to happen - and a blank SPFx web part has no error
   * UI, so it reads as a deployment problem rather than a crash.
   *
   * idx is computed defensively because there may be no flow yet: steps.length - 1 on an absent flow
   * would be -1, and the clamp has to survive that without the runner's `active` cast.
   */
  const stepCount = flow ? flow.steps.length : 0;
  const idx = stepCount > 0 ? Math.max(0, Math.min(stepIdx, stepCount - 1)) : 0;
  useEffect(() => {
    if (idx > maxIdx) setMaxIdx(idx);
  }, [idx, maxIdx]);

  if (!flow && !allTools) {
    return (
      <Shell>
        {/* The picker is the top of this page, so nothing else here points back to the directory that
            sent the admin — reported by the client 2026-08-15 as "I can't go back to CRS Settings".
            The deeper views have their own band back to the picker; this is the one that leaves. */}
        <BackToSettings context={context} siteUrl={siteUrl} />
        <h2 style={s.h2}>Folder Management</h2>
        <p style={s.sub}>
          Select an action below. You&rsquo;ll be guided through the required
          screens step by step.
        </p>

        {loadError && (
          <div style={s.danger}>
            Could not read this site&rsquo;s segments ({loadError}). This does{" "}
            <strong>not</strong> mean there are none — the flows still work,
            they just cannot tell you what is already done.
          </div>
        )}

        <div style={s.cards}>
          {FLOWS.filter((f) => f.tone === "normal").map((f) => (
            <button
              key={f.id}
              // ⚠ "Run folder reconciliation" SPANS THE FULL ROW (client, 2026-09-03: "Extend the
              // box Design for Folder recon till the end. Full width."). It is the one card with no
              // segment-picking flow behind it — every other card leads into a guided sequence of
              // steps — so it reads as the odd one out anyway; stretching it says so visually rather
              // than leaving it sized like a step-by-step flow it is not.
              style={
                f.id === "runRecon"
                  ? { ...s.card, gridColumn: "1 / -1" }
                  : s.card
              }
              onClick={() => openFlow(f)}
            >
              {/* Keyed on the FLOW ID, so a flow added later gets no icon rather than the wrong
                  one — `FlowIcon` renders nothing for a name it does not know. */}
              <span style={s.cardIcon}>
                <FlowIcon name={f.id} />
              </span>
              <div style={s.cardTitle}>{f.label}</div>
              <div style={s.cardBlurb}>{f.blurb}</div>
              <div style={s.cardFoot}>
                {/* The step count and "pick a segment first" came off on the client's instruction
                    (2026-09-06). Both were still true — the count comes from the flow itself and the
                    picker does appear first — but the card is a door, and neither fact changes which
                    door you take. The arrow keeps the footer from collapsing. */}
                <span />
                <span style={s.cardGo} aria-hidden="true">
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                    <path
                      d="M3 8h10M9 4l4 4-4 4"
                      stroke="currentColor"
                      strokeWidth="1.6"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                </span>
              </div>
            </button>
          ))}
        </div>

        {/* ⚠ NOT IN THE CARD GRID, and that is the point (client design, 2026-08-30). Retiring is the
            one destructive action on this page; a fifth tile among the others invites it to be
            clicked as casually as they are. A full-width banner in the danger palette reads as a
            different kind of thing, which is what it is. The "Take note" heading went with the
            redesign — the banner says everything the heading did, in the place it applies. */}
        {FLOWS.filter((f) => f.tone === "destructive").map((f) => (
          <button
            key={f.id}
            style={{ ...s.cardDanger, marginTop: 18 }}
            onClick={() => openFlow(f)}
          >
            <span
              style={{ flexShrink: 0, color: "#a4262c", marginTop: 1 }}
              aria-hidden="true"
            >
              <svg width="35" height="35" viewBox="0 0 16 16" fill="none">
                <path
                  d="M8 2.2 15 14H1L8 2.2z"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  strokeLinejoin="round"
                />
                <path
                  d="M8 6.4v3.2M8 11.6v.6"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  strokeLinecap="round"
                />
              </svg>
            </span>
            <span>
              <span style={s.dangerName}>{f.label}</span>
              <span style={s.dangerBlurb}>{f.blurb}</span>
            </span>
            {/* Same glyph as every card's, so "this is how you go in" is one visual idea across the
                page. `marginLeft: auto` pushes it to the far edge of the flex row. */}
            <span style={s.dangerGo} aria-hidden="true">
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                <path
                  d="M3 8h10M9 4l4 4-4 4"
                  stroke="currentColor"
                  strokeWidth="1.6"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </span>
          </button>
        ))}

        {/* ⚠ THE "All tools" LINK IS GONE FROM THE PICKER (client, 2026-08-20: *"I don't want client to
            touch and mess it all up without following proper flow"*). The five screens have a real
            ordering — an abbreviation before reconciliation, a saved pending chain before a migration —
            and opening one directly is how that order gets skipped.

            THE VIEW ITSELF STAYS, reachable by `#tab=<slug>`: those links were shipped on the CRS
            Settings page and may be bookmarked, and a public name that stops working is worse than one
            nobody clicks. Removing the invitation is not the same as removing the capability.

            What made this safe was giving reconciliation its own card (`runRecon`). It was the only
            screen with no flow of its own, so without it an admin needing to re-run reconciliation
            would have had to start the flow for ADDING A UNIT and jump the rail. */}
      </Shell>
    );
  }

  // ── All tools: exactly what this page was before ────────────────────────────
  if (allTools) {
    return (
      <Shell>
        <BackBand
          label="Back to Folder Management"
          onClick={() => setAllTools(false)}
        />
        <FolderManager context={context} />
      </Shell>
    );
  }

  // ── The flow runner ────────────────────────────────────────────────────────
  const active = flow as Flow;
  const steps = active.steps;
  const step = steps[idx];
  const needsPick = active.needsSegment && !segment;
  /* ⚠ THE MIGRATE STEP STAYS UNLOCKED AFTER ITS OWN RUN — see `appliedFor`. Everything else about
     the lock is unchanged: on a first visit with no pending change it renders exactly as before. */
  const locked =
    isLocked(step, effectiveFacts) &&
    !(step.id === "migrate" && appliedFor !== undefined && appliedFor === segKey);

  /**
   * The lowest step whose Next is currently blocked. Nothing PAST it may be reached from the rail.
   *
   * ⚠ WITHOUT THIS THE RAIL IS A SECOND ROUTE PAST THE GATE, and the client found it within the hour
   * (2026-09-07): uploads switched back ON, Next correctly held on step 1 — and steps 2 and 3 still
   * clickable in the side panel, because `reachable` was `i <= maxIdx` alone. `maxIdx` records the
   * furthest step ever REACHED, so once you have walked forward the rail keeps offering those steps
   * whatever the facts have since become. A gate on one control and not the other is not a gate.
   *
   * `steps.length` when nothing blocks, so the cap is inert on a clean flow.
   */
  const firstBlockedIdx = firstBlockedStepIndex(steps, effectiveFacts);
  const blockedReason = firstBlockedIdx < steps.length
    ? blocksNext(steps[firstBlockedIdx], effectiveFacts)
    : "";

  const renderStep = (st: FlowStep): React.ReactElement => {
    // The segment picker stands in front of every step that HAS a subject: without it the screens below
    // have none, and half of them would show the wrong one.
    //
    // But NOT in front of a step that makes no use of a segment (client, 2026-08-19: *"weird the first
    // step is asking for the business segment when it is suppose to be showing only a sign to users"*).
    // Those steps render instructions for work done elsewhere, or act site-wide. Replacing that with a
    // dropdown hides the one thing the step exists to say, and asks for a value it cannot spend.
    //
    // ⚠ THE TEST IS `stepUsesSegment`, NOT `kind !== "outside"`. That literal was the first fix, and
    // the very next step added to a flow — the SITE-WIDE upload pause — is a `component` and walked
    // straight back into the same bug. The rule lives in folderFlows.ts with a regression test.
    if (needsPick && stepUsesSegment(st)) {
      return (
        <div>
          <label style={s.label}>Which segment?</label>
          <select
            style={s.select}
            value={segKey}
            onChange={(e) => setSegKey(e.target.value)}
          >
            <option value="">Select a segment&hellip;</option>
            {(segments ?? []).map((x) => (
              <option key={x.key} value={x.key}>
                {x.label}
              </option>
            ))}
          </select>
          {segments !== undefined && segments.length === 0 && (
            <p style={s.hint}>
              No segments exist yet — use <strong>Add a new segment</strong>{" "}
              instead.
            </p>
          )}
          {segments === undefined && (
            <p style={s.hint}>
              The segment list could not be read, so it is empty here.
              Everything still works from <strong>All tools</strong>.
            </p>
          )}
        </div>
      );
    }

    if (st.screen.kind === "outside") {
      return (
        <div style={s.outside}>
          {/* ⚠ THE LINK GOES TO THE **CLASSIC, SITE-LEVEL** TERM STORE MANAGER, and that is the whole
              point of it (client, 2026-08-30, after hitting the wall themselves).

              An admin looking for the term store naturally lands on the MODERN one —
              `/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter` — which is the TENANT admin centre.
              It answers **"Access denied — you don't have access to this admin center operation"** to
              a site collection administrator, because it is gated on being a SharePoint
              Administrator for the whole tenant. The naming is backwards from the intuition:
              "modern" means newer, not site-scoped.

              `termstoremanager.aspx` is the per-site page an SCA can actually use, and it is where
              this project's term sets live — in the SITE-COLLECTION group (below the dashed line in
              its tree), which a site collection admin administers automatically. Verified on
              ClarenceDMSTesting 2026-08-30: the group is `DMS`, holding Group Head Office, Minamas
              Head Office, NBPOL Head Office and the shared sets.

              Built from `siteUrl` rather than hardcoded, so it is right on whichever site this is.

              ⚠ GATED ON THE THREE TERM-STORE STEPS, NOT ON `kind === "outside"`. `pauseFlows` is an
              `outside` step too, and it is about PAUSING POWER AUTOMATE — a term store link there
              would send an admin to the wrong tool at the one moment they are being asked to do
              something else entirely. Same trap as `stepUsesSegment` on 2026-08-19: the SCREEN KIND
              is not the question; what the step is ABOUT is. */}
          {(st.id === "termSet" ||
            st.id === "addTerm" ||
            st.id === "renameTerm") && (
            <p style={{ margin: "0 0 12px" }}>
              <a
                href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
                target="_blank"
                rel="noopener noreferrer"
                style={s.termStoreLink}
              >
                Open the Term Store Management Tool
                <span aria-hidden="true" style={{ marginLeft: 6 }}>
                  &#8599;
                </span>
              </a>
              <span style={s.termStoreNote}>
                Opens in a new tab. Use this page, not the term store in the
                SharePoint admin centre &mdash; that one needs tenant
                administrator rights and will refuse a site collection
                administrator.
              </span>
            </p>
          )}
          {/* NO HINT HERE. The panel prints `step.hint` immediately above this box, so repeating it
              rendered the same sentence twice, one under the other — which reads as a rendering fault
              rather than emphasis. This box carries the part that is NOT the hint. */}
          {active.asksSubject &&
            (st.id === "addTerm" || st.id === "renameTerm") && (
              <div style={{ marginTop: 14 }}>
                {/* PAST TENSE, and the segment named rather than ASKED for. "What are you adding?" over a
                  text box reads as "type it here and I'll add it" — but this is an `outside` step and
                  the box creates nothing; an admin could type a name, press Next, and believe the term
                  had been created. Naming the segment when one is already chosen gives the context the
                  client asked for WITHOUT putting a segment question in front of an instruction-only
                  step, which is the exact bug reported on 2026-08-19 (see `stepUsesSegment`). */}
                <label style={s.label}>
                  {active.asksSubject === "add"
                    ? `Which department or unit did you add${segment ? ` to ${segment.label}` : ""}?`
                    : `Which term did you rename${segment ? ` in ${segment.label}` : ""}?`}
                </label>
                <input
                  style={s.input}
                  value={subject}
                  placeholder="e.g. Treasury"
                  onChange={(e) => setSubject(e.target.value)}
                />
                {/* Naming it is what lets the next step point at the right row. OPTIONAL: leaving it blank
                  costs a little help, never progress — the flow must not stop them working. */}
                {/* WHY, not just "optional" — the client asked "why do I have to type in the term?".
                  Nothing can check whether someone renamed a term in the term store, so this box is
                  the only thing that can turn an uncheckable step into a checked one, and it carries
                  the name into the steps below. Blank costs help, never progress. */}
                <div style={s.hint}>
                  Type the name exactly as it appears in the term store — the
                  name, not a GUID. <strong>This box creates nothing.</strong>{" "}
                  Nothing can see into the term store, so naming the term is the
                  only way this step can confirm itself — and it takes you
                  straight to that term on the next step. Optional: leave it
                  blank and everything still works.
                </div>
              </div>
            )}
          {/* A step that asks for nothing has an empty box, which reads as a failed load. Say what the
              box is for instead: these steps are the ones this tool cannot do for you. */}
          {!(
            active.asksSubject &&
            (st.id === "addTerm" || st.id === "renameTerm")
          ) && (
            <div>
              This step happens outside this tool. Come back when it is done.
            </div>
          )}
        </div>
      );
    }

    if (st.screen.kind === "component") {
      // Mounted from ONE component for both the opening pause and the closing resume; `mode` only
      // changes the wording. Two components would drift, and the resume half is the one that would
      // never get tested — while forgetting it leaves a DMS that silently accepts no documents.
      if (st.screen.id === "pauseUploads") {
        return (
          <UploadPauseToggle
            context={context}
            siteUrl={siteUrl}
            mode={st.id === "resumeUploads" ? "resume" : "pause"}
            /* ⚠ UPDATES THE ONE FACT, NEVER `reload`. Bumping the reload counter would re-run the
               whole facts effect — three list reads — to answer a question this callback has already
               answered. Same reasoning as `GroupMembersEditor`'s `onChanged` calling `reloadMembers`
               rather than `reload`: re-reading everything to learn one thing is what makes a page
               feel broken.

               Without this the gate added in 1.0.452.0 contradicted the panel above it: the facts
               effect reads the setting once, so pausing with the toggle left `uploadsPaused` at
               `false` and Next stayed blocked saying uploads were on, clearable only by reloading the
               page (reported on site 2026-09-07). The toggle only ever calls this after a read or
               write that actually succeeded, so a failure still leaves the fact as it was. */
            onChanged={(paused) => setBaseFacts((f) => ({ ...f, uploadsPaused: paused }))}
          />
        );
      }
      /* ⚠ `groups` IS THE ONLY BRANCH LEFT HERE. The `folderAccess` step was removed on 2026-08-23
         — `GroupManager` mounts the member editor itself now, so a second step to "add the people"
         was the same screen a click later. The union in `folderFlows.ts` no longer carries that id,
         so adding a new component step is a compile error here rather than a silent fallthrough. */
      return (
        <div>
          {/* ⚠ THE "Create one group" MODE SWITCH IS GONE (2026-09-02, client: "remove the Create
                one group, since we are automatically creating the group for them"), matching the
                identical removal on the standalone Group Management page
                (`GroupManagementPage.tsx`). Bulk provisioning is the only creation path here now.

                `GroupManager`'s create-one-group FORM is NOT deleted, only permanently hidden
                (`hideCreateForm={true}`, unconditionally) — the same "kept, not deleted" pattern as
                `GroupMapBuilder`'s form. It still exists for CRS_SITE_MEMBERS and a genuine one-off
                nobody planned; re-enabling it here is passing `false` again, not rebuilding it. */}
          <BulkGroupProvisioner
            context={context}
            siteUrl={siteUrl}
            onBusyChange={setGroupBusy}
            onRunComplete={onRunComplete}
          />
          {/* ⚠ `RolesReference` REMOVED AGAIN HERE (2026-09-02, client's confirmed call, same
                request as removing it from the standalone page). It had been brought back
                2026-08-30 specifically because it is the ONLY place documenting how to create the
                three custom permission levels (CRS Upload/Approve/Delete) — without it, a newly
                provisioned unit can silently end up with no uploader grant at all, and reconciliation
                only reports the gap after the fact ("no CRS Approve role definition on site"). That
                risk is real again with it gone; flagged to the client before removing, and removed
                anyway on their explicit instruction. If this bites again, the fix is re-adding
                `<RolesReference spHttpClient={context.spHttpClient} siteUrl={siteUrl} />` here, not
                rebuilding it — the component itself is untouched. */}
          {/* The LIST always renders — hiding what already exists is how a group gets created twice. */}
          <GroupManager
            context={context}
            siteUrl={siteUrl}
            hideCreateForm={true}
            refreshKey={reload}
          />
        </div>
      );
    }

    /* THE `key` IS LOAD-BEARING, and its absence was a live bug (2026-08-17): stepping from
       "New segment" to "CRS Term Abbreviations" changed the heading and the rail highlight while the
       New segment FORM stayed on screen.

       FolderManager resolves `initialTab` in a `useState` INITIALISER, which React runs once per
       mounted instance. Without a key React reconciles the same element type across a step change,
       keeps the instance, and the new prop is never read — so the flow silently drives nothing and the
       page describes a screen the admin is not looking at. Worse than a dead link, because it looks
       like it worked: they would fill in one form under another step's heading.

       Keyed on the TAB rather than the step id on purpose: steps of different flows share a tab
       (ABBREVIATIONS appears in four), so keying on the tab re-uses the instance where the screen is
       genuinely the same and re-mounts only when it changes. The re-mount cost per step was accepted
       in the design — reconciliation is inline in a 4,000-line file, and extracting it to make it
       mountable would risk the most site-verified code here for a navigation change. */
    /* AFTER the form, a PICKER of the segments that exist — not a name to type.
       Client, 2026-08-17: *"The UX flow kinda doesn't make sense, what is What will the new segment be
       called?"* — and they were right twice over. Asking for the name up front asked for the same value
       the form below was already asking for, which reads as a bug; and the list was read once on mount,
       so a segment created on this very step never appeared, leaving the rail saying "Not checked" beside
       a panel saying "is configured".

       Picking is better than typing on every count: nothing to spell (so no fullwidth-＆ trap), it is the
       same control every other flow uses, and it is derived from data so it survives a refresh, a second
       tab and someone else's session. It also sets `segKey`, which is what carries the segment into steps
       3-6 instead of leaving each one to ask again. */
    const confirms =
      active.asksSubject === "newSegment" && st.id === "createSegment";
    // On step 2 the form is BEHIND A BUTTON, and closes once the segment is confirmed. Everywhere else
    // the step IS the screen, so it renders directly.
    if (confirms && !showForm && segment) {
      return (
        <div>
          <div style={s.doneBox}>
            <strong>{segment.label}</strong> is created — top folder{" "}
            <strong>{segment.code}</strong>. The remaining steps below are set
            up for it.
          </div>
          <button style={s.ghost} onClick={() => setShowForm(true)}>
            + Create another segment
          </button>
        </div>
      );
    }
    return (
      <div>
        {/* ⚠ THE TERM STORE LINK MOVED HERE ON 2026-09-06, and it had to move somewhere.
            `addUnit` and `rename` lost the instruction steps that carried it when the client's
            redesign cut them to two steps — and the work it points at is still the FIRST thing an
            admin has to do: a term that does not exist in the term store has no row on this screen,
            so without the link the step gives no way to create one.

            ⚠ THE LINK IS THE **CLASSIC, SITE-LEVEL** MANAGER, and that is the whole point of it. The
            modern one (`/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter`) is the TENANT admin
            centre and answers "Access denied" to a site collection administrator. Built from
            `siteUrl`, never hardcoded.

            Shown on the abbreviations step ONLY. It is gated on the step id rather than the screen
            kind for the reason `stepUsesSegment` exists: what the step is ABOUT is the question, not
            what it happens to render. */}
        {/* ⚠ WRAPPED IN `s.outside` — THE GREEN BANNER. When this moved off the instruction step it
            left that container behind and rendered as a bare link on white, which the client
            reported as the design banner "not being there". The box is what makes it read as the
            one thing to do before anything else on the screen. */}
        {st.id === "abbreviations" && (
          <div style={{ ...s.outside, marginBottom: 14 }}>
            <a
              href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
              target="_blank"
              rel="noopener noreferrer"
              style={s.termStoreLink}
            >
              Open the term store management tool
              <span aria-hidden="true" style={{ marginLeft: 6 }}>
                &#8599;
              </span>
            </a>
            <span style={s.termStoreNote}>
              Use this tool instead of the SharePoint admin centre version,
              which requires tenant administrator access.
            </span>
          </div>
        )}
        {confirms && !showForm && (
          <div style={{ marginBottom: 16 }}>
            <button style={s.primary} onClick={() => setShowForm(true)}>
              + Create a new segment
            </button>
            <div style={s.hint}>
              Already created it earlier? Pick it below to carry on — no need to
              open the form again.
            </div>
          </div>
        )}
        {(!confirms || showForm) && (
          <FolderManager
            key={st.screen.tab}
            context={context}
            initialTab={st.screen.tab}
            hideTabs
            /* The creation announces itself, so nothing has to be inferred: close the form, select the
               new segment, and re-read the list so it appears in the picker and in every later step.
               Before this the only signal was the admin choosing from a dropdown they had to refresh
               first — which left the form open under its own success message. */
            onSegmentCreated={(key) => {
              setSegKey(key);
              setShowForm(false);
              setReload((n) => n + 1);
              /* ⚠ AND MOVE ON BY ITSELF (client, 2026-09-06: *"the system should automatically
                 proceed to Step 3 screen … without requiring System Admin to click Next Step
                 button manually. This will make the process clearer and easier"*).

                 Safe only because the creation ANNOUNCED itself. This runs from `onCreated`, which
                 SegmentCreator calls after the mode row is written — so the segment provably exists
                 by the time the step changes, which is exactly what the abbreviations step is locked
                 behind. Advancing on anything weaker (a timer, a form close) could land on a locked
                 step with nothing to explain it.

                 ⚠ `stepIdx + 1`, NOT a hardcoded 2. The step lives at index 2 of `newSegment`
                 today; a step inserted above it would silently send the admin to the wrong screen,
                 and this is the only flow that auto-advances so nothing else would catch it.
                 `idx` is clamped on read, so an overrun cannot leave the runner out of range. */
              setStepIdx((n) => n + 1);
            }}
            /* Delete belongs to the Retire flow, which moves the documents out first and asks for a typed
               confirmation. Offering it beside a creation form gives the destructive half with none of
               that — and note this is scoped to THIS flow, not to `hideTabs`, because Retire mounts the
               same tab embedded and needs the button. */
            hideSegmentDelete={active.id === "newSegment"}
            hideSegmentCreate={active.id === "retire"}
            /* The abbreviations screen hands over its own count — the flow cannot afford the term-tree
               walk that produces it, and without it `undefined` gated nothing, so Next was clickable on a
               screen full of "no folder will be created" warnings. */
            onAbbreviationsMissingChange={setAbbrevMissing}
            onAbbreviationsLoadingChange={setAbbrevLoading}
            // Reconciliation gets the SAME padlock as a bulk group run: it lives in the page, has no
            // resume, and takes an hour at the client's scale — leaving the step stops it mid-folder.
            onReconRunningChange={onReconRunning}
            // THE SAME padlock, deliberately. The migration's scan and its move both run in this
            // page with no resume, so leaving the step throws the work away — identical to
            // reconciliation and to a bulk group run. One reason to hold navigation, one
            // implementation of holding it, one message.
            onMigrateRunningChange={setMigrateBusy}
            // NOT the padlock: unsaved edits are not a run, and holding Back would trap someone who
            // opened the screen by mistake. It gates NEXT only, with its own reason.
            onStructureDirtyChange={setStructureDirty}
            /* ⚠ RE-READS THE SEGMENT LIST, which is the ONE definition of `pendingLevels` — the
               fact the migrate step locks on. Bumping `reload` re-runs the read that owns it; the
               cheap-facts effect merely COPIES it, so re-running that alone re-copies the stale
               value. Before this, saving on the levels step left the very next step reporting
               "There is no pending structure change to move to" about the change just made. */
            onStructureSaved={() => setReload((n) => n + 1)}
            onAbbreviationsDirtyChange={setAbbreviationsDirty}
            onAbbreviationsRegisterSave={(fn) => {
              abbrevSaveRef.current = fn;
            }}
            // Gates NEXT only, like unsaved level edits — not the rail, because Back must stay open.
            onMigratePendingChange={setMigratePending}
            /* ⚠ RE-READS THE SEGMENT LIST, the ONE definition of `pendingLevels` — the fact the new
               `migrate` gate reads. The cheap-facts effect only COPIES it, so bumping `reload` is
               what makes Next open the moment the migration switches the new shape on. Same reason
               and same shape as `onStructureSaved` above. */
            onMigrateApplied={() => {
              // Remember it BEFORE the re-read: `reload` is what flips `pendingLevels` to false and
              // therefore what would fire this step's own lock.
              setAppliedFor(segKey);
              setReload((n) => n + 1);
            }}
            // The flow already asked which segment, and prints it in the header — so the migration
            // screen must not ask a second time. Same principle as `stepUsesSegment`: if the flow
            // knows, it does not ask.
            migrateInitialSegmentKey={segment?.key}
            /* The same fact `blocksNext` gates step 1 on, so the migrate screen's warning and the
               flow's own gate can never disagree about whether uploads are on. Read from
               `effectiveFacts` rather than `baseFacts` for exactly that reason — it is what the gate
               reads. `undefined` still shows the warning; only a known-paused state hides it. */
            migrateUploadsPaused={effectiveFacts.uploadsPaused}
            abbreviationsInitialSegmentKey={segment?.key}
          />
        )}
        {confirms && (
          <div style={{ ...s.card, marginTop: 18, marginBottom: 0 }}>
            {/* "Which segment did you just create?" described the case that no longer needs this control
                — creation selects itself since 1.0.136.0 — and the client read the dropdown as a way to
                CREATE one: *"I actually thought at first that I can recreate Group Head Office when I
                already create."* Naming it for RESUMING is the whole point: this exists for coming back
                to an unfinished segment, and it is also what carries the segment into steps 3-6. */}
            <label style={s.label} htmlFor="fa-newseg">
              Continuing an earlier segment?
            </label>
            <div
              style={{
                display: "flex",
                gap: 8,
                alignItems: "center",
                flexWrap: "wrap",
              }}
            >
              {/* ⚠ ALREADY-BUILT SEGMENTS ARE LEFT OUT (client, 2026-09-09: *"ensure that the Segment
                  that is already built do not need to be appearing under the Continuing an earlier
                  segment? It won't make sense, its already built"*). This control exists to RESUME an
                  unfinished creation, and a segment with folders is finished — offering it invited
                  walking a creation flow over a live segment, which is also what produced the earlier
                  *"I actually thought at first that I can recreate Group Head Office"*.
                  ⚠ FAIL-OPEN: `built === undefined` means the Folder Map could not be read, and such a
                  segment is LISTED. Hiding on a failed read would remove the only route back into an
                  unfinished segment.
                  ⚠ AND THE CURRENTLY-PICKED SEGMENT IS ALWAYS LISTED, whatever its state. A `<select>`
                  whose `value` matches no option renders the FIRST option while state keeps the old
                  one — the trap that shipped once on the Position dropdown — so the control would say
                  one segment while the flow carried another. */}
              <select
                id="fa-newseg"
                style={s.select}
                value={segKey}
                // Confirming closes the form: it has served its purpose, and leaving a creation form under
                // the confirmation of what was just created is what made this read as broken.
                onChange={(e) => {
                  setSegKey(e.target.value);
                  if (e.target.value) setShowForm(false);
                }}
              >
                <option value="">Select a segment&hellip;</option>
                {resumable.map((x) => (
                  <option key={x.key} value={x.key}>
                    {x.label}
                  </option>
                ))}
              </select>
              {/* The list is read on mount, and Create happens after — so without this the segment just
                  made is missing from its own confirmation. One list read. */}
              {/* `type="button"` is defensive rather than a fix for the report: there is no <form>
                  here today, but a button with no type submits one the day somebody adds it. */}
              <button
                type="button"
                style={s.ghost}
                disabled={refreshing}
                onClick={() => {
                  setRefreshing(true);
                  setReload((n) => n + 1);
                }}
              >
                {refreshing ? "Refreshing…" : "Refresh list"}
              </button>
            </div>
            <div style={s.hint}>
              {segments === undefined
                /* ⚠ IT USED TO END "use the list of steps to move between them", AND THAT WAS
                   FALSE. The rail stopped navigating forward on 2026-08-30, so this pointed at a
                   control that does nothing — the same dead escape hatch already corrected in the
                   `createSegment` gate message, missed on this sibling string. What is true is that
                   nothing is gated when the list cannot be read (see `segmentChosen`), so the way
                   on is Next, and the way to try again is the Refresh list button beside this. */
                ? "The segment list could not be read, so a segment cannot be picked here. Press Refresh list to try again — Next is not held, so you can carry on either way."
                : resumable.length === 0
                /* ⚠ SAID PLAINLY, because an empty dropdown beside a Refresh button reads as a list
                   that failed to load — and the admin's next move would be pressing Refresh for ever.
                   This is the NORMAL state on a settled site: every segment has folders, so none of
                   them is an unfinished creation. */
                ? "Every segment on this site is already set up, so there is nothing here to continue. Create one above, or press Refresh list if you have just made one."
                : "Only needed if you created it earlier, or reopened this page. A segment you create above is selected for you. Picking one takes the remaining steps straight to it — it does not create anything. Segments that are already set up are not listed."}
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <Shell>
      {/* Held during a run for the same reason as the rail: this is the widest exit on the screen.
          ⚠ ALSO held on an unsaved Structure or Abbreviations edit (found live 2026-08-26): `openFlow`
          and `leaveFlow` never reset `structureDirty`/`abbreviationsDirty`, so this was previously the
          one exit that let an admin walk straight out of a dirty screen into a DIFFERENT flow entirely
          — leaving the stale flag blocking Next there too, with a message about a screen they are no
          longer looking at. Same "refuse, don't confirm" rule as the tab-switch guard inside
          FolderManager itself ("Finish or clear what you are editing first"). */}
      <BackBand
        label="Back to Folder Management"
        onClick={leaveFlow}
        disabled={runBusy || structureDirty}
      />
      {structureDirty && !runBusy && (
        <p style={s.hint}>
          Finish or clear what you are editing first — leaving now would lose
          it, and the guided flow you land on would open holding a warning that
          belongs to this screen, not to it.
        </p>
      )}
      <h2 style={s.h2}>{active.label}</h2>
      <p style={s.sub}>
        {segment ? (
          <>
            Segment: <strong>{segment.label}</strong>.{" "}
          </>
        ) : undefined}
        {subject.trim().length > 0 ? (
          <>
            Subject: <strong>{subject.trim()}</strong>.{" "}
          </>
        ) : undefined}
        {/* ⚠ "N of M steps still to check" REMOVED, and the per-step tick with it (client QA,
            2026-08-30: *"JUST FORCE SYSTEM ADMIN TO FOLLOW EVERY STEP. NO NEED TO HAVE STEP 1
            CHECKED, STEP 2 CHECKED AND SO ON"*). Their objection was that on step 1 of a brand new
            segment it read "5 of 5 steps still to check", which states the obvious and reads like a
            warning. Step {idx + 1} of {steps.length} says where you are without grading you. */}
        Step {idx + 1} of {steps.length}.
      </p>

      <div style={s.runner}>
        <nav style={s.rail}>
          {steps.map((st, i) => {
            const isActive = i === idx;
            /* ⚠ THE RAIL NO LONGER NAVIGATES FORWARD, and this REVERSES the rule that governed it
               until today: *every step is reachable; the rail says what is outstanding and padlocks
               nothing*. The client asked for the opposite outright — *"ensure that side panel
               doesn't allow user to navigate"* — and confirmed it knowing the trade: *"just do as
               the client's requested, if any issue for the flow we will fix that flow."*

               BACKWARD is still allowed. Re-reading a step you have already passed changes nothing
               and is how somebody checks what they typed; blocking it would turn a wizard into a
               one-way corridor with no way to look back. Forward is what the sequence is for.

               ⚠ THE CONSEQUENCE TO WATCH: an admin who did step 3 outside the tool last week must
               now click through 1 and 2 to reach it. That was the original argument for a clickable
               rail, and it is now the client's cost to carry rather than ours to prevent. */
            /* ⚠ THREE CONDITIONS, AND THE LAST ONE KEEPS BACKWARD NAVIGATION FREE.
               `i <= maxIdx` is the earned-progress rule (1.0.332.0). `i <= firstBlockedIdx` is the
               new cap, so the rail cannot walk past a step whose Next is held. And `i <= idx` means
               everything at or behind where you stand stays clickable — without it, a fact turning
               false while an admin is on a later step would strand them there, unable even to go
               BACK to the step that needs fixing, which is the opposite of what the gate wants. */
            const reachable = isStepReachable(i, {
              maxIdx,
              idx,
              firstBlocked: firstBlockedIdx,
            });
            return (
              <button
                key={st.id}
                style={{
                  ...s.railItem,
                  ...(isActive ? s.railActive : {}),
                  ...(runBusy || !reachable ? s.railBusy : {}),
                }}
                // A bulk group run in flight still blocks EVERY move, backward included: leaving the
                // step unmounts the run and stops it part-way, silently.
                disabled={runBusy || !reachable}
                title={
                  runBusy
                    ? "Wait for the run to finish."
                    : reachable
                      ? undefined
                      : /* The gate's OWN reason when a gate is what holds it, so the rail and the
                           Next button never explain the same refusal two different ways. Falls back
                           to the earned-progress wording when nothing is blocked and the step is
                           simply ahead of where the admin has walked. */
                        i > firstBlockedIdx && blockedReason
                        ? blockedReason
                        : "Finish the step you are on first."
                }
                onClick={() => setStepIdx(i)}
              >
                {/* The connector, on every item BUT the last — a line below the final number would
                    trail into nothing, which is what makes a single line behind the whole list the
                    wrong shape here. */}
                {i < steps.length - 1 ? <span style={s.railLink} /> : undefined}
                {/* GREEN once reached, grey ahead. That is a POSITION, not a verdict: the per-step
                    Done / To do / Not checked grading was removed at the client's request the same
                    day, and this must not quietly reintroduce it. A reached step is one you have
                    walked past, never one that has been checked. */}
                <span
                  style={{
                    ...s.railNum,
                    ...(i <= maxIdx ? {} : s.railNumIdle),
                  }}
                >
                  {i + 1}
                </span>
                <span>
                  <span style={s.railLabel}>{st.label}</span>
                </span>
              </button>
            );
          })}
        </nav>

        <div style={s.panel}>
          <h3 style={s.stepHead}>{step.label}</h3>
          <p style={s.stepHint}>{step.hint}</p>

          {/* ⚠ THE CHOSEN SEGMENT STAYS CHANGEABLE (client, 2026-09-07: *"I accidentally selected buah
              and now I can't reselect another"*). `needsPick` is `needsSegment && !segment`, so the
              picker below vanished the moment anything was chosen — and the only way back was
              `< Back to Folder Management`, which clears `segKey` on both leave and open. That escape
              exists, but it reads like abandoning the whole flow rather than correcting one field.

              Rendered HERE rather than inside `renderStep` so it sits ABOVE the step's own content and
              is not one of the mutually exclusive branches below.

              ⚠⚠ AND ABOVE THE LOCK BOX, NOT BELOW IT — AND NOT GATED ON `!locked`. That gate was the
              first build and it recreated the very trap this control exists to remove: picking a
              segment whose step is LOCKED (Group Led Project has no pending change, so the migrate
              step reads "Not ready yet") hid the switcher, leaving the admin stuck on that segment
              with no way back but leaving the flow. **The state where you most need to change segment
              is precisely the one where the step cannot proceed.** Reported on site within minutes.

              ⚠ IT MUST NOT APPEAR ON A STEP THAT SPENDS NO SEGMENT — `stepUsesSegment` again, the rule
              that the site-wide upload pause walked into twice. Offering a segment picker above the
              pause toggle asks for a value that step cannot use.

              ⚠ AND CHANGING IT RESETS `maxIdx` TO THE CURRENT STEP. Forward navigation is earned by
              walking the flow (1.0.332.0); that progress was earned against the OLD segment, so
              carrying it would let an admin jump to a step they have never done for the segment now
              selected. The facts effect re-reads on its own — it keys on `segKey` — so the ticks and
              gates correct themselves. */}
          {!needsPick && active.needsSegment && stepUsesSegment(step) && (
            <div style={s.segSwitch}>
              <label style={s.label}>Segment</label>
              {/* ⚠ HELD BY `runBusy`, THE SAME PADLOCK AS THE RAIL, BACK AND NEXT — and it was the one
                  control left out of it (client, 2026-09-07: *"I notice I can select when the Group
                  Head office is running the scan"*).
                  Switching mid-run does not stop the scan — `SubtreeMigrator` keeps its own `chosen`
                  and its pre-select effect refuses to override one already made — so the run carries
                  on against the OLD segment while the flow header names the NEW one. Nothing breaks;
                  the screen simply states two different segments at once, and the admin has no way to
                  tell which the result belongs to. */}
              <select
                style={{ ...s.select, ...(runBusy ? s.off : {}) }}
                value={segKey}
                disabled={runBusy}
                onChange={(e) => {
                  setSegKey(e.target.value);
                  setMaxIdx(idx);
                }}
              >
                {(segments ?? []).map((x) => (
                  <option key={x.key} value={x.key}>
                    {x.label}
                  </option>
                ))}
              </select>
              <p style={s.hint}>
                {runBusy
                  ? "You cannot change segment while this step is running — the run would carry on against the segment it started with."
                  : "Change this to work on a different segment. The steps you have already passed are re-checked for whichever one you pick."}
              </p>
            </div>
          )}

          {/* BELOW the switcher deliberately: the lock is a fact about THIS segment, so an admin
              reading "Not ready yet" needs the control that changes the segment already in view
              above it, not hidden behind it. */}
          {locked && (
            <div style={s.lockBox}>
              <strong>Not ready yet.</strong> {lockReason(step, effectiveFacts)}
            </div>
          )}

          {!locked && renderStep(step)}

          {/* Next is unavailable until the step is actually done (2026-08-17, client: "won't it make
              more sense once client finish filling up the New Segment and saved and only next step is
              available?").

              Only for the two steps whose `todo` is definitive — see NEXT_GATED_STEPS. It stays enabled
              on every advisory step and on every UNKNOWN fact, so a list that could not be read cannot
              trap someone mid-flow. The rail entries remain clickable either way: this makes the
              recommended path obvious without making the others unreachable, which is the client's own
              rule that the flow must not stop them working. */}
          {(() => {
            const blocked = blocksNext(step, effectiveFacts);
            const last = idx >= steps.length - 1;
            return (
              <>
                <div style={s.navBar}>
                  <button
                    style={idx === 0 || runBusy ? s.off : s.ghost}
                    disabled={idx === 0 || runBusy}
                    onClick={() => setStepIdx(Math.max(0, idx - 1))}
                  >
                    Back
                  </button>
                  <button
                    style={
                      last ||
                      blocked ||
                      runBusy ||
                      structureDirty ||
                      abbrevSaving ||
                      migratePending
                        ? s.off
                        : s.primary
                    }
                    disabled={
                      last ||
                      blocked.length > 0 ||
                      runBusy ||
                      structureDirty ||
                      abbrevSaving ||
                      migratePending
                    }
                    /* ⚠ NEXT SAVES THE ABBREVIATIONS BEFORE IT ADVANCES (client, 2026-09-06 —
                       that screen's Save button is gone). It calls the SCREEN'S OWN `save()`, which
                       validates, writes and re-baselines its rows; the runner never writes them
                       itself, or there would be two definitions of what saving means.

                       ⚠ AND IT ADVANCES ONLY IF THAT SUCCEEDED. The codes live in that component's
                       state and nowhere else, so a step change after a failed write loses them
                       silently — which is precisely what the removed "save first" warning existed
                       to prevent. A refusal (two siblings sharing a name) counts as a failure here:
                       reconciliation would abort on it three steps later. The screen reports the
                       reason itself, in the panel already being read. */
                    onClick={() => {
                      const go = (): void =>
                        setStepIdx(Math.min(steps.length - 1, idx + 1));
                      const save =
                        step.id === "abbreviations" ? abbrevSaveRef.current : undefined;
                      if (!save) {
                        go();
                        return;
                      }
                      setAbbrevSaving(true);
                      save()
                        .then((ok: boolean) => {
                          if (ok) go();
                        })
                        .catch(() => undefined)
                        .then(() => setAbbrevSaving(false))
                        .catch(() => undefined);
                    }}
                  >
                    {abbrevSaving ? "Saving…" : "Next"}
                  </button>
                  {last && (
                    /* ⚠ FINISH IS GATED TOO, and it has to be: `reconcile` is the LAST step in four
                       of the five flows that have it, so gating Next alone would enforce the run in
                       the structure flow and nowhere else.

                       ⚠ THE BACK BAND STAYS OPEN, which is what keeps this an insistence rather than
                       a cage. It calls the same `leaveFlow` and is held only by `runBusy` and an
                       unsaved edit — so nobody is ever trapped in a flow by a reconciliation they
                       cannot get to complete. Do not gate that band on this. */
                    <button
                      style={
                        runBusy || structureDirty || blocked.length > 0
                          ? s.off
                          : s.ghost
                      }
                      disabled={runBusy || structureDirty || blocked.length > 0}
                      onClick={leaveFlow}
                    >
                      Finish
                    </button>
                  )}
                </div>
                {/* The reason sits BESIDE the disabled button, never only in a tooltip: a greyed button
                    with no explanation reads as a broken page, and the admin's next move is to reload
                    rather than to finish the step. */}
                {/* ⚠ NO LONGER `!last`. Finish is gated now, so the last step can carry a reason —
                    and suppressing it there would grey out the flow's own way of saying "done" with
                    no explanation, which is the broken-page reading this hint exists to prevent. */}
                {blocked.length > 0 &&
                  !runBusy &&
                  !structureDirty &&
                  !migratePending && (
                    <div style={s.hint}>
                      {blocked}
                      {last && " Once it has run, press Finish. You can also leave with Back to Folder Management."}
                    </div>
                  )}
                {/* Its own reason, again: "a run is in progress" would be wrong — nothing is running,
                    the admin simply has not pressed the button yet. */}
                {migratePending && !runBusy && (
                  <div style={s.hint}>
                    There are folders still to rebuild. Run them first — moving
                    on now would leave the segment half-changed and the next
                    step turns uploads back on over it. Choose a value for every
                    library listed, then run the rebuild.
                    {/* ⚠ THE ADVICE TO SKIP A LIBRARY IS GONE, because skipping is no longer possible:
                        1.0.502.0 holds Rebuild until every folder has a value. It read "set that
                        library to Leave this unit alone" until 1.0.499.0 and "Leave this library
                        alone" until now — THIRD correction of the same sentence in one day, each time
                        because it named a control or an action that had moved. An option label is a
                        string, so nothing type-checks the sentence that names it: grep every bundle
                        for the old label after renaming one. */}
                  </div>
                )}
                {/* Its own message, not folded into the run padlock: the fix is a click away on this
                    very screen, and telling someone "a run is in progress" when they simply have not
                    pressed Save would send them looking for something that is not happening. */}
                {structureDirty && !runBusy && (
                  <div style={s.hint}>
                    Save the structure first — the next step reads the saved
                    pending change, so it would report nothing to move. Press{" "}
                    <strong>Save structure</strong> above, or Cancel to discard
                    the edit.
                  </div>
                )}
                {/* ⚠ THE UNSAVED-ABBREVIATIONS GATE IS GONE, AND REMOVING IT WAS REQUIRED RATHER
                    THAN OPTIONAL. It held Back, Next and Finish and told the admin to "press Save
                    above" — a button that no longer exists, which would have left them with no way
                    out of the flow at all.

                    What replaces it is stronger: Next SAVES. The state it guarded against —
                    advancing with codes held only in component state — cannot arise, because
                    reconciliation is now reached through a save that had to succeed.

                    ⚠ LEAVING BY BACK OR THE BACK BAND STILL DISCARDS THE EDIT. Client, asked
                    directly, 2026-09-06: *"if they leave halfway and it is erased, that is their
                    problem."* The `beforeunload` guard inside the screen still catches a closed
                    tab. */}
                {/* Beside the greyed buttons, never only in a tooltip — the same rule as the Next
                    gate. A held navigation bar with no stated reason reads as a broken page. */}
                {/* ONE message for both runs, because it is one padlock. Naming neither specifically
                    keeps it true for whichever is in flight — a message naming the bulk run while
                    reconciliation was going would read as the wrong screen. */}
                {runBusy && (
                  <div style={s.hint}>
                    A run is in progress on this step. Moving away would stop it
                    part-way, so navigation is held until it finishes. Anything
                    already written stays written, and re-running picks up what
                    is missing.
                  </div>
                )}
              </>
            );
          })()}
        </div>
      </div>
    </Shell>
  );
}
