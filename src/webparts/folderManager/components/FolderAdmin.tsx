// Folder Management — guided flows, with the five tabs kept behind "All tools".
//
// Spec: docs/superpowers/specs/2026-08-14-folder-management-guided-flows-design.md
//
// The client's complaint: the tab bar "is confusing and client doesnt know how it works". The tabs were
// already in the order the work happens, which was not enough — five equal doors do not say that four of
// them are steps of one job, and the order is what breaks. Miss Term Abbreviations and reconciliation
// creates NOTHING, silently.
//
// THE GOVERNING RULE, stated by the client: *"the flow should not stop them from doing the work."* So this
// screen marks what is outstanding and locks almost nothing. The rules live in shared/folderFlows.ts,
// where `isLocked` is written so any unknown fact — unread, or a read that failed — cannot lock a step.
//
// It DRIVES FolderManager rather than dismantling it: reconciliation lives inline in a 4,000-line file and
// is the most site-verified code in the project, so a navigation change must not go near it. Steps
// FolderManager hosts render as `<FolderManager initialTab hideTabs />`; Group Management and Folder
// Access, which it does not host, are mounted straight from userAccess — one component, two mount points,
// never a second copy.
import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import FolderManager from "./FolderManager";
import { IFolderManagerProps } from "./IFolderManagerProps";
import GroupManager from "../../userAccess/components/GroupManager";
import GroupMapBuilder from "../../userAccess/components/GroupMapBuilder";
import {
  FLOWS,
  Flow,
  FlowFacts,
  FlowStep,
  firstIncompleteStep,
  isLocked,
  lockReason,
  remainingCount,
  stepState,
} from "../../../shared/folderFlows";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { tabFromHash } from "../../../shared/adminPages";
import { BackBand, BackToSettings } from "../../../shared/backToSettings";

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
type Segment = { key: string; label: string; code: string; termSetGuid: string; pendingLevels: boolean };

const s: Record<string, React.CSSProperties> = {
  wrap:      { fontFamily: '"Segoe UI", system-ui, sans-serif', color: "#242424" },
  h2:        { fontSize: 22, fontWeight: 600, margin: "0 0 6px" },
  sub:       { fontSize: 13, color: "#5f5f5f", margin: "0 0 20px", lineHeight: 1.5 },
  cards:     { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 300px), 1fr))", gap: 14 },
  card:      { textAlign: "left", border: "1px solid #e1e1e1", borderRadius: 10, background: "#fff", padding: "16px 18px", cursor: "pointer", font: "inherit" },
  cardDanger:{ textAlign: "left", border: "1px solid #f1b0b3", borderRadius: 10, background: "#fdf9f9", padding: "16px 18px", cursor: "pointer", font: "inherit" },
  cardTitle: { fontSize: 15, fontWeight: 600, margin: 0 },
  cardBlurb: { fontSize: 12.5, color: "#616161", margin: "5px 0 0", lineHeight: 1.5 },
  cardMetaMuted:{ fontSize: 11.5, color: "#8a8886", marginTop: 10 },
  sectionHead:{ fontSize: 12, fontWeight: 600, color: "#605e5c", textTransform: "uppercase", letterSpacing: ".04em", margin: "26px 0 10px" },
  // The "All tools ›" link at the foot of the picker. It borrows the back band's TEXT style
  // (`backLinkStyle` in shared/backToSettings) without the band, because it is a descent into a
  // deeper screen rather than an exit — banding it would make two opposite moves look identical.
  back:      { border: "none", background: "transparent", padding: 0, font: "inherit", color: "rgba(0, 104, 74, 1)", fontSize: 14, fontWeight: 600, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6 },
  runner:    { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 240px), 1fr))", gap: 24, alignItems: "start" },
  rail:      { border: "1px solid #e8e6e6", borderRadius: 10, background: "#fafafa", padding: 12, maxWidth: 280 },
  railItem:  { display: "flex", gap: 10, width: "100%", textAlign: "left", border: "none", background: "transparent", font: "inherit", padding: "9px 8px", borderRadius: 8, cursor: "pointer", alignItems: "flex-start" },
  railActive:{ background: "#eef7f1" },
  railNum:   { flexShrink: 0, width: 20, height: 20, borderRadius: "50%", fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", marginTop: 1 },
  railLabel: { fontSize: 12.5, fontWeight: 600, lineHeight: 1.35 },
  railState: { fontSize: 11, lineHeight: 1.4, marginTop: 2, display: "block" },
  panel:     { minWidth: 0, gridColumn: "span 2" },
  stepHead:  { fontSize: 18, fontWeight: 600, margin: "0 0 4px" },
  stepHint:  { fontSize: 13, color: "#5f5f5f", margin: "0 0 16px", lineHeight: 1.55 },
  outside:   { border: "1px solid #cfd8e3", background: "#f4f7fb", borderRadius: 8, padding: "14px 16px", fontSize: 13, color: "#2b3f56", lineHeight: 1.6 },
  lockBox:   { border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 8, padding: "12px 14px", fontSize: 12.5, color: "#8a4b00", lineHeight: 1.55, marginBottom: 14 },
  label:     { display: "block", fontWeight: 600, fontSize: 12, margin: "0 0 4px" },
  select:    { width: "100%", maxWidth: 420, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  input:     { width: "100%", maxWidth: 420, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4 },
  navBar:    { display: "flex", gap: 8, marginTop: 20, paddingTop: 16, borderTop: "1px solid #eceaea", flexWrap: "wrap" },
  primary:   { padding: "7px 16px", fontSize: 13, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  off:       { padding: "7px 16px", fontSize: 13, background: "#e6e6e6", color: "#9a9a9a", border: "none", borderRadius: 4, cursor: "not-allowed" },
  ghost:     { padding: "7px 16px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  danger:    { marginBottom: 16, padding: "10px 12px", border: "1px solid #f1b0b3", background: "#fdf3f4", borderRadius: 8, fontSize: 12.5, color: "#a4262c", lineHeight: 1.5 },
  allTools:  { marginTop: 30, paddingTop: 16, borderTop: "1px solid #eceaea" },
  hint:      { fontSize: 11.5, color: "#5f6f80", marginTop: 6, lineHeight: 1.5 },
};

const STATE_STYLE: Record<string, { pill: React.CSSProperties; note: React.CSSProperties; text: string; mark: string }> = {
  done:    { pill: { background: "#0f6c3f", color: "#fff" }, note: { color: "#0f6c3f" }, text: "Done", mark: "✓" },
  todo:    { pill: { background: "#fff2df", color: "#8a4b00", border: "1px solid #f0d9b5" }, note: { color: "#8a4b00" }, text: "To do", mark: "" },
  unknown: { pill: { background: "#f0eff0", color: "#767676" }, note: { color: "#8a8886" }, text: "Not checked", mark: "" },
};

export default function FolderAdmin({ context }: IFolderManagerProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;

  // Read once, synchronously, so an addressed screen never flashes the picker first.
  const [flow, setFlow] = useState<Flow | undefined>(
    () => FLOWS.filter((f) => f.id === readHash().flowId)[0],
  );
  const [stepIdx, setStepIdx] = useState(0);
  const [allTools, setAllTools] = useState(() => readHash().wantsTabs);

  const [segments, setSegments] = useState<Segment[] | undefined>(undefined);
  const [segKey, setSegKey] = useState("");
  const [loadError, setLoadError] = useState<string | undefined>(undefined);

  /** The subject of flows 2 and 4 — what makes their term-store step checkable at all. */
  const [subject, setSubject] = useState("");

  const [facts, setFacts] = useState<FlowFacts>({});

  const GET = { Accept: "application/json;odata=nometadata" };

  /**
   * Segments, from the DMS Config mode rows.
   *
   * One read, and it also answers `pendingLevels` — so the structure flow's lock costs nothing extra.
   */
  useEffect(() => {
    const load = async (): Promise<void> => {
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Id,Title,ModeLabel,StagingFolder,TermSetGuid,PendingLevels,ConfigType&$top=200`,
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
          pendingLevels: String(r.PendingLevels ?? "").trim().length > 0,
        }))
        .filter((r) => r.key.length > 0)
        .sort((a, b) => a.label.localeCompare(b.label));
      setSegments(rows);
      setLoadError(undefined);
    };
    load().catch((e) => {
      // Unreadable ≠ none. A failed read must not present as "this site has no segments", which would
      // send an admin off to create one that already exists.
      setSegments(undefined);
      setLoadError((e as Error).message);
    });
  }, [siteUrl]);

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
          .get(`${siteUrl}/_api/web/sitegroups?$select=Title&$top=500`, SPHttpClient.configurations.v1, { headers: GET })
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
        if (mapRows) {
          next.folderAccessRows = mapRows.filter(
            (r) => (r.Segment ?? "").trim().toLowerCase() === target.termSetGuid.toLowerCase(),
          ).length > 0;
        }

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
      if (!cancelled) setFacts(next);
    };
    load().catch(() => { if (!cancelled) setFacts({}); });
    return () => { cancelled = true; };
  }, [flow ? flow.id : "", segKey, segments]);

  /** Open a flow on the first thing left to do. */
  const openFlow = (f: Flow): void => {
    setFlow(f);
    setAllTools(false);
    setSubject("");
    setStepIdx(firstIncompleteStep(f, facts));
  };

  const leaveFlow = (): void => {
    setFlow(undefined);
    setFacts({});
    setStepIdx(0);
  };

  // ── The picker ─────────────────────────────────────────────────────────────
  if (!flow && !allTools) {
    return (
      <section style={s.wrap}>
        {/* The picker is the top of this page, so nothing else here points back to the directory that
            sent the admin — reported by the client 2026-08-15 as "I can't go back to CRS Settings".
            The deeper views have their own band back to the picker; this is the one that leaves. */}
        <BackToSettings context={context} siteUrl={siteUrl} />
        <h2 style={s.h2}>Folder Management</h2>
        <p style={s.sub}>
          Pick what you are trying to do. Each one walks you through only the screens it needs, in order.
        </p>

        {loadError && (
          <div style={s.danger}>
            Could not read this site&rsquo;s segments ({loadError}). This does <strong>not</strong> mean
            there are none — the flows still work, they just cannot tell you what is already done.
          </div>
        )}

        <div style={s.cards}>
          {FLOWS.filter((f) => f.tone === "normal").map((f) => (
            <button key={f.id} style={s.card} onClick={() => openFlow(f)}>
              <div style={s.cardTitle}>{f.label}</div>
              <div style={s.cardBlurb}>{f.blurb}</div>
              <div style={s.cardMetaMuted}>
                {f.steps.length} step{f.steps.length === 1 ? "" : "s"}
                {f.needsSegment ? " · pick a segment first" : ""}
              </div>
            </button>
          ))}
        </div>

        <div style={s.sectionHead}>Careful</div>
        <div style={s.cards}>
          {FLOWS.filter((f) => f.tone === "destructive").map((f) => (
            <button key={f.id} style={s.cardDanger} onClick={() => openFlow(f)}>
              <div style={s.cardTitle}>{f.label}</div>
              <div style={s.cardBlurb}>{f.blurb}</div>
            </button>
          ))}
        </div>

        {/* The escape hatch. Deliberately quiet, and deliberately present: any job the flows do not cover
            would otherwise be unreachable, testing included. */}
        <div style={s.allTools}>
          <button style={s.back} onClick={() => setAllTools(true)}>
            All tools &rsaquo; open the five screens directly
          </button>
        </div>
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
  const idx = Math.max(0, Math.min(stepIdx, steps.length - 1));
  const step = steps[idx];
  const needsPick = active.needsSegment && !segment;
  const locked = isLocked(step, facts);

  const renderStep = (st: FlowStep): React.ReactElement => {
    // The segment picker stands in front of every step of a segment-scoped flow: without it the screens
    // below have no subject, and half of them would show the wrong one.
    if (needsPick) {
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
          <div>{st.hint}</div>
          {active.asksSubject && (st.id === "addTerm" || st.id === "renameTerm") && (
            <div style={{ marginTop: 14 }}>
              <label style={s.label}>
                {active.asksSubject === "add" ? "What are you adding?" : "Which term did you rename?"}
              </label>
              <input
                style={s.input}
                value={subject}
                placeholder="e.g. Treasury"
                onChange={(e) => setSubject(e.target.value)}
              />
              {/* Naming it is what lets the next step point at the right row. OPTIONAL: leaving it blank
                  costs a little help, never progress — the flow must not stop them working. */}
              <div style={s.hint}>Optional. Naming it lets the next step take you straight to it.</div>
            </div>
          )}
        </div>
      );
    }

    if (st.screen.kind === "component") {
      return st.screen.id === "groups"
        ? <GroupManager context={context} siteUrl={siteUrl} />
        : <GroupMapBuilder context={context} siteUrl={siteUrl} />;
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
    return (
      <FolderManager key={st.screen.tab} context={context} initialTab={st.screen.tab} hideTabs />
    );
  };

  return (
    <section style={s.wrap}>
      <BackBand label="Back to Folder Management" onClick={leaveFlow} />
      <h2 style={s.h2}>{active.label}</h2>
      <p style={s.sub}>
        {segment ? <>Segment: <strong>{segment.label}</strong>. </> : undefined}
        {subject.trim().length > 0 ? <>Subject: <strong>{subject.trim()}</strong>. </> : undefined}
        {remainingCount(active, facts)} of {steps.length} step{steps.length === 1 ? "" : "s"} still to check.
      </p>

      <div style={s.runner}>
        <nav style={s.rail}>
          {steps.map((st, i) => {
            const state = stepState(st, facts);
            const style = STATE_STYLE[state];
            const isActive = i === idx;
            return (
              <button
                key={st.id}
                style={{ ...s.railItem, ...(isActive ? s.railActive : {}) }}
                // EVERY step is reachable. The rail says what is outstanding; it does not padlock
                // navigation — the client's rule is that the flow must not stop them working.
                onClick={() => setStepIdx(i)}
              >
                <span style={{ ...s.railNum, ...style.pill }}>{style.mark || i + 1}</span>
                <span>
                  <span style={s.railLabel}>{st.label}</span>
                  <span style={{ ...s.railState, ...style.note }}>
                    {isLocked(st, facts) ? "Needs an earlier step" : style.text}
                  </span>
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
              <strong>Not ready yet.</strong> {lockReason(step, facts)}
            </div>
          )}

          {!locked && renderStep(step)}

          <div style={s.navBar}>
            <button
              style={idx === 0 ? s.off : s.ghost}
              disabled={idx === 0}
              onClick={() => setStepIdx(Math.max(0, idx - 1))}
            >
              &lsaquo; Back
            </button>
            <button
              style={idx >= steps.length - 1 ? s.off : s.primary}
              disabled={idx >= steps.length - 1}
              onClick={() => setStepIdx(Math.min(steps.length - 1, idx + 1))}
            >
              Next step &rsaquo;
            </button>
            {idx >= steps.length - 1 && (
              <button style={s.ghost} onClick={leaveFlow}>Finish</button>
            )}
          </div>
        </div>
      </div>
    </section>
  );
}
