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
import { UPLOAD_PAUSE_SETTING, uploadsArePaused } from "../../../shared/uploadPause";
import { useEffect, useState } from "react";
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
  firstIncompleteStep,
  isLocked,
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
    try { flowId = decodeURIComponent(m[1]).trim(); } catch { flowId = ""; }
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
};

const s: Record<string, React.CSSProperties> = {
  /* The Upload Form's page shell - see the note in AuditLog.tsx. This page had NO container at
     all: no centring, no padding, hard against both edges.
     WARN: DELIBERATELY NO `maxWidth`. It hosts the widest content in the product - the
     reconciliation run panel, the migration scan and the group tables all render inside it - and
     capping the column would put a horizontal scrollbar on the screens that most need width. The
     padding alone answers what was asked: it no longer sticks to the wall. */
  wrap:      { fontFamily: '"Segoe UI", system-ui, sans-serif', color: "#242424", margin: "32px auto", padding: "0 24px 48px" },
  h2:        { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 6px" },
  sub:       { fontSize: 13, color: "#5f5f5f", margin: "0 0 20px", lineHeight: 1.5 },
  /* THE CARD GRID (client design, 2026-08-30). Four across on a wide screen, and it reflows rather
     than scrolling: `auto-fit` with a `min()` floor keeps a card readable on a phone instead of
     squeezing five into the width. */
  cards:     { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 250px), 1fr))", gap: 16, alignItems: "stretch" },
  /* `display: flex` + `column` is what lets the step-count row sit on the BOTTOM of every card
     whatever the blurb's length — cards of different text lengths otherwise end with the meta line
     at a different height each, which reads as misalignment rather than as different content. */
  card:      { display: "flex", flexDirection: "column", textAlign: "left", border: "1px solid #e6e6e6", borderRadius: 12, background: "#fff", padding: "20px 18px 16px", cursor: "pointer", font: "inherit" },
  cardIcon:  { marginBottom: 14 },
  cardTitle: { fontSize: 15, fontWeight: 600, margin: 0, color: "#242424" },
  cardBlurb: { fontSize: 12.5, color: "#616161", margin: "6px 0 0", lineHeight: 1.5, flex: "1 1 auto" },
  /* The footer: step count left, arrow right — `marginTop: auto` is what pins it down. */
  cardFoot:  { display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 16 },
  cardMetaMuted:{ fontSize: 11.5, color: "#8a8886" },
  cardGo:    { width: 30, height: 30, borderRadius: 8, border: "1px solid #cfe3d5", background: "#fff", color: "#0f6c3f", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  /* Retiring is a BANNER, not a card (client design): it is the one destructive action here, and a
     fifth tile in the same grid invites it to be clicked as casually as the others. */
  cardDanger:{ display: "flex", alignItems: "flex-start", gap: 10, width: "100%", textAlign: "left", border: "1px solid #f3c9cb", borderRadius: 10, background: "#fdf4f4", padding: "14px 16px", cursor: "pointer", font: "inherit" },
  dangerName:{ fontSize: 13.5, fontWeight: 600, color: "#a4262c", marginRight: 8 },
  dangerBlurb:{ fontSize: 12.5, color: "#605e5c", lineHeight: 1.5 },
  sectionHead:{ fontSize: 12, fontWeight: 600, color: "#605e5c", textTransform: "uppercase", letterSpacing: ".04em", margin: "26px 0 10px" },
  // The "All tools ›" link at the foot of the picker. It borrows the back band's TEXT style
  // (`backLinkStyle` in shared/backToSettings) without the band, because it is a descent into a
  // deeper screen rather than an exit — banding it would make two opposite moves look identical.
  back:      { border: "none", background: "transparent", padding: 0, font: "inherit", color: "rgba(0, 104, 74, 1)", fontSize: 14, fontWeight: 600, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6 },
  // FLEX, not `repeat(auto-fit, …)`. That grid made as many 240px columns as would fit, so on a wide
  // screen it produced a third empty column and the panel — spanning two — left a band of dead space to
  // its right (client, 2026-08-17: "CAn you remove the padding?"). Flex gives the same wrapping on a
  // narrow screen with no phantom tracks on a wide one.
  runner:    { display: "flex", flexWrap: "wrap", gap: 20, alignItems: "flex-start" },
  /* THE STEPPER (client design, 2026-08-30). White card rather than the old grey block, so it reads
     as a panel beside the work rather than a sidebar behind it. `position: relative` anchors the
     connector line drawn between the numbers. */
  rail:      { position: "relative", border: "1px solid #e6e6e6", borderRadius: 12, background: "#fff", padding: "18px 16px", flex: "0 1 260px" },
  railItem:  { position: "relative", display: "flex", gap: 12, width: "100%", textAlign: "left", border: "none", background: "transparent", font: "inherit", padding: "10px 8px", borderRadius: 8, cursor: "pointer", alignItems: "center" },
  railActive:{ background: "#eef7f1" },
  /* The line joining one number to the next. Drawn per ITEM rather than as one line behind the
     list, because the last step must not trail a stub below it — which is exactly what a single
     absolutely-positioned line would do. */
  railLink:  { position: "absolute", left: 21, top: 34, width: 2, bottom: -6, background: "#e2e2e2" },
  /* The number pill used to take its colour from the step STATE (done / todo / unknown). With the
     state gone it needs a resting look of its own — `s` is a `Record<string, CSSProperties>`,
     so a key that does not exist renders unstyled with a green build. */
  railNumIdle: { background: "#eaeaea", color: "#8a8886" },
  /* `zIndex` lifts the number above the connector line so the line appears to run BETWEEN the
     circles rather than through them. */
  railNum:   { position: "relative", zIndex: 1, flexShrink: 0, width: 26, height: 26, borderRadius: "50%", fontSize: 12, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", background: "#0f6c3f", color: "#fff" },
  railLabel: { fontSize: 13, fontWeight: 600, lineHeight: 1.35, color: "#242424" },
  railState: { fontSize: 11, lineHeight: 1.4, marginTop: 2, display: "block" },
  // `flex: 1 1 420px` — takes the rest of the row, wraps under the rail below ~700px. The old
  // `gridColumn: span 2` belonged to the auto-fit grid above and is what left the dead band.
  panel:     { minWidth: 0, flex: "1 1 420px" },
  stepHead:  { fontSize: 18, fontWeight: 600, margin: "0 0 4px" },
  stepHint:  { fontSize: 13, color: "#5f5f5f", margin: "0 0 16px", lineHeight: 1.55 },
  /* ⚠ GREEN, NOT BLUE, since 2026-08-30 (client: *"keep the old design but tweak the color to match
     the other designs"*). The CONTENT of an `outside` step is unchanged and deliberately so — their
     mock replaced step 1 with a Create Term Set form, and **this app cannot create a term set or its
     terms**; the client's own answer was *"we do not have the time to built that interface"*. So the
     step still says the work happens in the term store, in the palette the rest of the page uses. */
  termStoreLink: { display: "inline-block", fontWeight: 600, color: "#0f6c3f", textDecoration: "underline" },
  termStoreNote: { display: "block", fontSize: 11.5, color: "#5f7a6b", marginTop: 4, lineHeight: 1.5 },
  outside:   { border: "1px solid #cfe3d5", background: "#f2f9f5", borderRadius: 10, padding: "14px 16px", fontSize: 13, color: "#1c4d33", lineHeight: 1.6 },
  lockBox:   { border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 8, padding: "12px 14px", fontSize: 12.5, color: "#8a4b00", lineHeight: 1.55, marginBottom: 14 },
  label:     { display: "block", fontWeight: 600, fontSize: 12, margin: "0 0 4px" },
  select:    { width: "100%", maxWidth: 420, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  input:     { width: "100%", maxWidth: 420, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  railBusy:  { cursor: "not-allowed", opacity: 0.55 },
  navBar:    { display: "flex", gap: 8, marginTop: 20, paddingTop: 16, borderTop: "1px solid #eceaea", flexWrap: "wrap" },
  primary:   { padding: "7px 16px", fontSize: 13, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  off:       { padding: "7px 16px", fontSize: 13, background: "#e6e6e6", color: "#9a9a9a", border: "none", borderRadius: 4, cursor: "not-allowed" },
  ghost:     { padding: "7px 16px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  danger:    { marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 8, fontSize: 12.5, color: "#a4262c", lineHeight: 1.5 },
  allTools:  { marginTop: 30, paddingTop: 16, borderTop: "1px solid #eceaea" },
  modeBar:   { display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" },
  modeOn:    { padding: "7px 14px", fontSize: 13, fontWeight: 600, background: "#0f6c3f", color: "#fff", border: "1px solid #0f6c3f", borderRadius: 4, cursor: "pointer" },
  modeOff:   { padding: "7px 14px", fontSize: 13, background: "#fff", color: "#1b1b1b", border: "1px solid #c7c7c7", borderRadius: 4, cursor: "pointer" },
  hint:      { fontSize: 11.5, color: "#5f6f80", marginTop: 6, lineHeight: 1.5 },
  doneBox:   { border: "1px solid #c6e3d1", background: "#f1f8f4", borderRadius: 8, padding: "12px 14px", fontSize: 13, color: "#0f6c3f", lineHeight: 1.55, marginBottom: 14 },
};

/* ⚠ REMOVED 2026-08-30 with the per-step status, at the client's request. `stepState` and
   `remainingCount` in `shared/folderFlows.ts` are UNTOUCHED and still tested — only this screen
   stopped rendering them. That matters: `isLocked` and `blocksNext` read the same `FlowFacts`, so
   the four real locks (segment exists, PendingLevels set, abbreviations complete) still work and
   still explain themselves beside the Next button. What went is the GRADING of every step, not the
   checking of the ones that can actually be checked. Re-render it by restoring this map. */

export default function FolderAdmin({ context }: IFolderManagerProps): React.ReactElement {
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
  const [abbrevMissing, setAbbrevMissing] = useState<number | undefined>(undefined);
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
  const [runBusy, setRunBusy] = useState(false);
  /**
   * The Folder levels screen holds unsaved edits.
   *
   * Worse here than the tab-switch case it already guards: the NEXT step reads `PendingLevels`,
   * which an unsaved edit has not written — so walking on reports "nothing to move" for a change
   * the admin believes they made, and the flow looks broken rather than incomplete.
   */
  const [structureDirty, setStructureDirty] = useState(false);
  const [abbreviationsDirty, setAbbreviationsDirty] = useState(false);
  /**
   * The migration screen has scanned work it has not run.
   *
   * The client walked straight past Rebuild on the first run (2026-08-19). Doing so leaves the
   * segment staged and the folders in the old shape, and the NEXT step turns uploads back on over
   * it — so the flow would end with a half-applied structure and people filing into it.
   */
  const [migratePending, setMigratePending] = useState(false);

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
        .filter((r) => String(r.ConfigType ?? "").trim().toLowerCase() === "mode")
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
          for (const r of ((await p2.json()).value ?? []) as Array<Record<string, unknown>>) {
            staged[Number(r.Id ?? 0)] = String(r.PendingLevels ?? "").trim().length > 0;
          }
        }
      } catch {
        // Leave it unknown. A transient failure must not lock the migrate step.
      }
      // Merged only when the second read succeeded. Absent column or failed read leaves every
      // segment `undefined`, which reads as "not checked" and locks nothing.
      setSegments(
        stagedKnown
          ? rows.map((r) => ({ ...r, pendingLevels: staged[r.itemId] === true }))
          : rows,
      );
      setLoadError(undefined);
    };
    load().catch((e) => {
      // Unreadable ≠ none. A failed read must not present as "this site has no segments", which would
      // send an admin off to create one that already exists.
      setSegments(undefined);
      setLoadError((e as Error).message);
    });
  }, [siteUrl, reload]);

  const segment = (segments ?? []).filter((x) => x.key === segKey)[0];

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
          .get(`${siteUrl}/_api/web/sitegroups?$select=Title&$top=5000`, SPHttpClient.configurations.v1, { headers: GET })
          .then(async (r) => (r.ok ? ((await r.json()).value ?? []) as Array<{ Title?: string }> : undefined))
          .catch(() => undefined);
        if (groups) {
          const prefix = target.code.toLowerCase();
          next.groupsExist = prefix.length > 0 &&
            groups.filter((g) => (g.Title ?? "").toLowerCase().indexOf(prefix) === 0).length > 0;
        }

        const mapRows = await context.spHttpClient
          .get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items` +
              `?$select=Id,Segment&$top=5000`,
            SPHttpClient.configurations.v1, { headers: GET },
          )
          .then(async (r) => (r.ok ? ((await r.json()).value ?? []) as Array<{ Segment?: string }> : undefined))
          .catch(() => undefined);
        const folderRows = await context.spHttpClient
          .get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.folderMap))}')/items` +
              `?$select=Id,Section&$top=5000`,
            SPHttpClient.configurations.v1, { headers: GET },
          )
          .then(async (r) => (r.ok ? ((await r.json()).value ?? []) as Array<{ Section?: string }> : undefined))
          .catch(() => undefined);
        if (folderRows) {
          next.foldersExist = folderRows.filter(
            (r) => (r.Section ?? "").trim().toLowerCase() === target.code.toLowerCase(),
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
          const pRows = ((await pRes.json()).value ?? []) as Array<{ SettingValue?: string }>;
          // No row is a valid state meaning "not paused"; only a failed READ is unknown.
          next.uploadsPaused = pRows.length > 0 ? uploadsArePaused(pRows[0].SettingValue) : false;
        }
      } catch { /* leaves uploadsPaused undefined — unknown, never "on" */ }
      if (!cancelled) setBaseFacts(next);
    };
    load().catch(() => { if (!cancelled) setBaseFacts({}); });
    return () => { cancelled = true; };
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
      ...(abbrevMissing === undefined ? {} : { abbreviationsMissing: abbrevMissing }),
    };
    if (!flow || flow.asksSubject !== "newSegment") return facts;
    // Unreadable list ⇒ change nothing, so nothing is gated. This is the ONLY fail-open case here.
    if (segments === undefined) return facts;
    // The list read fine, so "nothing picked" is the admin not having answered — not a failure.
    if (!segment) return { ...facts, subjectGiven: false };
    return { ...facts, subjectGiven: true, segmentExists: true };
  }, [flow, baseFacts, segments, segment, abbrevMissing, abbrevLoading]);

  /** Open a flow on the first thing left to do. */
  const openFlow = (f: Flow): void => {
    setFlow(f);
    setAllTools(false);
    setSubject("");
    setStepIdx(firstIncompleteStep(f, baseFacts));
  };

  const leaveFlow = (): void => {
    setFlow(undefined);
    setBaseFacts({});
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
      <section style={s.wrap}>
        {/* The picker is the top of this page, so nothing else here points back to the directory that
            sent the admin — reported by the client 2026-08-15 as "I can't go back to CRS Settings".
            The deeper views have their own band back to the picker; this is the one that leaves. */}
        <BackToSettings context={context} siteUrl={siteUrl} />
        <h2 style={s.h2}>Folder Management</h2>
        <p style={s.sub}>
          Select an action below. You&rsquo;ll be guided through the required screens step by step.
        </p>

        {loadError && (
          <div style={s.danger}>
            Could not read this site&rsquo;s segments ({loadError}). This does <strong>not</strong> mean
            there are none — the flows still work, they just cannot tell you what is already done.
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
              style={f.id === "runRecon" ? { ...s.card, gridColumn: "1 / -1" } : s.card}
              onClick={() => openFlow(f)}>
              {/* Keyed on the FLOW ID, so a flow added later gets no icon rather than the wrong
                  one — `FlowIcon` renders nothing for a name it does not know. */}
              <span style={s.cardIcon}><FlowIcon name={f.id} /></span>
              <div style={s.cardTitle}>{f.label}</div>
              <div style={s.cardBlurb}>{f.blurb}</div>
              <div style={s.cardFoot}>
                <span style={s.cardMetaMuted}>
                  {f.steps.length} step{f.steps.length === 1 ? "" : "s"}
                  {f.needsSegment ? " · pick a segment first" : ""}
                </span>
                <span style={s.cardGo} aria-hidden="true">
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                    <path d="M3 8h10M9 4l4 4-4 4" stroke="currentColor" strokeWidth="1.6"
                      strokeLinecap="round" strokeLinejoin="round" />
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
          <button key={f.id} style={{ ...s.cardDanger, marginTop: 18 }} onClick={() => openFlow(f)}>
            <span style={{ flexShrink: 0, color: "#a4262c", marginTop: 1 }} aria-hidden="true">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M8 2.2 15 14H1L8 2.2z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
                <path d="M8 6.4v3.2M8 11.6v.6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
              </svg>
            </span>
            <span>
              <span style={s.dangerName}>{f.label}</span>
              <span style={s.dangerBlurb}>{f.blurb}</span>
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
      </section>
    );
  }

  // ── All tools: exactly what this page was before ────────────────────────────
  if (allTools) {
    return (
      <section style={s.wrap}>
        <BackBand label="Back to Folder Management" onClick={() => setAllTools(false)} />
        <FolderManager context={context} />
      </section>
    );
  }

  // ── The flow runner ────────────────────────────────────────────────────────
  const active = flow as Flow;
  const steps = active.steps;
  const step = steps[idx];
  const needsPick = active.needsSegment && !segment;
  const locked = isLocked(step, effectiveFacts);

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
          <select style={s.select} value={segKey} onChange={(e) => setSegKey(e.target.value)}>
            <option value="">Select a segment&hellip;</option>
            {(segments ?? []).map((x) => (
              <option key={x.key} value={x.key}>{x.label}</option>
            ))}
          </select>
          {segments !== undefined && segments.length === 0 && (
            <p style={s.hint}>
              No segments exist yet — use <strong>Add a new segment</strong> instead.
            </p>
          )}
          {segments === undefined && (
            <p style={s.hint}>
              The segment list could not be read, so it is empty here. Everything still works from{" "}
              <strong>All tools</strong>.
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
          {(st.id === "termSet" || st.id === "addTerm" || st.id === "renameTerm") && (
          <p style={{ margin: "0 0 12px" }}>
            <a
              href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
              target="_blank"
              rel="noopener noreferrer"
              style={s.termStoreLink}
            >
              Open the Term Store Management Tool
              <span aria-hidden="true" style={{ marginLeft: 6 }}>&#8599;</span>
            </a>
            <span style={s.termStoreNote}>
              Opens in a new tab. Use this page, not the term store in the SharePoint admin centre
              &mdash; that one needs tenant administrator rights and will refuse a site collection
              administrator.
            </span>
          </p>
          )}
          {/* NO HINT HERE. The panel prints `step.hint` immediately above this box, so repeating it
              rendered the same sentence twice, one under the other — which reads as a rendering fault
              rather than emphasis. This box carries the part that is NOT the hint. */}
          {active.asksSubject && (st.id === "addTerm" || st.id === "renameTerm") && (
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
                Type the name exactly as it appears in the term store — the name, not a GUID.{" "}
                <strong>This box creates nothing.</strong> Nothing can see into the term store, so
                naming the term is the only way this step can confirm itself — and it takes you
                straight to that term on the next step. Optional: leave it blank and everything still
                works.
              </div>
            </div>
          )}
          {/* A step that asks for nothing has an empty box, which reads as a failed load. Say what the
              box is for instead: these steps are the ones this tool cannot do for you. */}
          {!(active.asksSubject && (st.id === "addTerm" || st.id === "renameTerm")) && (
            <div>This step happens outside this tool. Come back when it is done.</div>
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
              onBusyChange={setRunBusy}
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
    const confirms = active.asksSubject === "newSegment" && st.id === "createSegment";
    // On step 2 the form is BEHIND A BUTTON, and closes once the segment is confirmed. Everywhere else
    // the step IS the screen, so it renders directly.
    if (confirms && !showForm && segment) {
      return (
        <div>
          <div style={s.doneBox}>
            <strong>{segment.label}</strong> is created — top folder <strong>{segment.code}</strong>.
            The remaining steps below are set up for it.
          </div>
          <button style={s.ghost} onClick={() => setShowForm(true)}>+ Create another segment</button>
        </div>
      );
    }
    return (
      <div>
        {confirms && !showForm && (
          <div style={{ marginBottom: 16 }}>
            <button style={s.primary} onClick={() => setShowForm(true)}>+ Create a new segment</button>
            <div style={s.hint}>
              Already created it earlier? Pick it below to carry on — no need to open the form again.
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
            onReconRunningChange={setRunBusy}
            // THE SAME padlock, deliberately. The migration's scan and its move both run in this
            // page with no resume, so leaving the step throws the work away — identical to
            // reconciliation and to a bulk group run. One reason to hold navigation, one
            // implementation of holding it, one message.
            onMigrateRunningChange={setRunBusy}
            // NOT the padlock: unsaved edits are not a run, and holding Back would trap someone who
            // opened the screen by mistake. It gates NEXT only, with its own reason.
            onStructureDirtyChange={setStructureDirty}
            onAbbreviationsDirtyChange={setAbbreviationsDirty}
            // Gates NEXT only, like unsaved level edits — not the rail, because Back must stay open.
            onMigratePendingChange={setMigratePending}
            // The flow already asked which segment, and prints it in the header — so the migration
            // screen must not ask a second time. Same principle as `stepUsesSegment`: if the flow
            // knows, it does not ask.
            migrateInitialSegmentKey={segment?.key}
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
            <label style={s.label} htmlFor="fa-newseg">Continuing an earlier segment?</label>
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
              <select
                id="fa-newseg"
                style={s.select}
                value={segKey}
                // Confirming closes the form: it has served its purpose, and leaving a creation form under
                // the confirmation of what was just created is what made this read as broken.
                onChange={(e) => { setSegKey(e.target.value); if (e.target.value) setShowForm(false); }}
              >
                <option value="">Select a segment&hellip;</option>
                {(segments ?? []).map((x) => (
                  <option key={x.key} value={x.key}>{x.label}</option>
                ))}
              </select>
              {/* The list is read on mount, and Create happens after — so without this the segment just
                  made is missing from its own confirmation. One list read. */}
              <button style={s.ghost} onClick={() => setReload((n) => n + 1)}>Refresh list</button>
            </div>
            <div style={s.hint}>
              {segments === undefined
                ? "The segment list could not be read, so a segment cannot be picked here — carry on, and use the list of steps to move between them."
                : "Only needed if you created it earlier, or reopened this page. A segment you create above is selected for you. Picking one takes the remaining steps straight to it — it does not create anything."}
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <section style={s.wrap}>
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
        disabled={runBusy || structureDirty || abbreviationsDirty}
      />
      {(structureDirty || abbreviationsDirty) && !runBusy && (
        <p style={s.hint}>
          Finish or clear what you are editing first — leaving now would lose it, and the guided flow
          you land on would open holding a warning that belongs to this screen, not to it.
        </p>
      )}
      <h2 style={s.h2}>{active.label}</h2>
      <p style={s.sub}>
        {segment ? <>Segment: <strong>{segment.label}</strong>. </> : undefined}
        {subject.trim().length > 0 ? <>Subject: <strong>{subject.trim()}</strong>. </> : undefined}
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
            const reachable = i <= maxIdx;
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
                  runBusy ? "Wait for the run to finish."
                    : reachable ? undefined
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
                <span style={{ ...s.railNum, ...(i <= maxIdx ? {} : s.railNumIdle) }}>{i + 1}</span>
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
                    &lsaquo; Back
                  </button>
                  <button
                    style={
                      last || blocked || runBusy || structureDirty || abbreviationsDirty || migratePending
                        ? s.off
                        : s.primary
                    }
                    disabled={
                      last ||
                      blocked.length > 0 ||
                      runBusy ||
                      structureDirty ||
                      abbreviationsDirty ||
                      migratePending
                    }
                    onClick={() => setStepIdx(Math.min(steps.length - 1, idx + 1))}
                  >
                    Next step &rsaquo;
                  </button>
                  {last && (
                    <button
                      style={runBusy || structureDirty || abbreviationsDirty ? s.off : s.ghost}
                      disabled={runBusy || structureDirty || abbreviationsDirty}
                      onClick={leaveFlow}
                    >
                      Finish
                    </button>
                  )}
                </div>
                {/* The reason sits BESIDE the disabled button, never only in a tooltip: a greyed button
                    with no explanation reads as a broken page, and the admin's next move is to reload
                    rather than to finish the step. */}
                {!last && blocked.length > 0 && !runBusy && !structureDirty && !abbreviationsDirty && !migratePending && <div style={s.hint}>{blocked}</div>}
                {/* Its own reason, again: "a run is in progress" would be wrong — nothing is running,
                    the admin simply has not pressed the button yet. */}
                {migratePending && !runBusy && (
                  <div style={s.hint}>
                    There are folders still to rebuild. Run them first — moving on now would leave the
                    segment half-changed and the next step turns uploads back on over it. If you meant
                    to skip a unit, set it to <strong>Leave this unit alone</strong> and check again.
                  </div>
                )}
                {/* Its own message, not folded into the run padlock: the fix is a click away on this
                    very screen, and telling someone "a run is in progress" when they simply have not
                    pressed Save would send them looking for something that is not happening. */}
                {structureDirty && !runBusy && (
                  <div style={s.hint}>
                    Save the structure first — the next step reads the saved pending change, so it
                    would report nothing to move. Press <strong>Save structure</strong> above, or
                    Cancel to discard the edit.
                  </div>
                )}
                {/* Same shape as structureDirty above, on the sibling screen it was never applied to
                    (found live 2026-08-26, client's own question: "wouldn't that mean someone can
                    click Next without saving?"). Reconciliation reads the SAVED rows, so an unsaved
                    edit here means the run reports success having done nothing for it — silently,
                    with the admin believing the rename happened. */}
                {abbreviationsDirty && !runBusy && (
                  <div style={s.hint}>
                    Save the abbreviation changes first — reconciliation reads the saved codes, so it
                    would rename nothing for this edit and report success anyway. Press{" "}
                    <strong>Save</strong> above, or discard the edit to continue without it.
                  </div>
                )}
                {/* Beside the greyed buttons, never only in a tooltip — the same rule as the Next
                    gate. A held navigation bar with no stated reason reads as a broken page. */}
                {/* ONE message for both runs, because it is one padlock. Naming neither specifically
                    keeps it true for whichever is in flight — a message naming the bulk run while
                    reconciliation was going would read as the wrong screen. */}
                {runBusy && (
                  <div style={s.hint}>
                    A run is in progress on this step. Moving away would stop it part-way, so
                    navigation is held until it finishes. Anything already written stays written, and
                    re-running picks up what is missing.
                  </div>
                )}
              </>
            );
          })()}
        </div>
      </div>
    </section>
  );
}
