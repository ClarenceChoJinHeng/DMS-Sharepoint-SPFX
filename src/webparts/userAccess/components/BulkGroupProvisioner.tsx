// Bulk group provisioning — every group a segment needs, in one run.
//
// Spec: docs/superpowers/specs/2026-08-18-group-creation-and-bulk-provisioning-design.md §5
//
// The client's site needs ~703 groups across 132 units. One at a time is days of clicking, and the mistake
// it invites is worse than the tedium: a wrong persona chosen once, then repeated by hand sixty times.
//
// The RULES live in shared/bulkGroups.ts, pure and tested, because they decide what gets created on a real
// tenant and the cost of being wrong is hundreds of groups deleted by hand. This file is the screen and the
// run: read, preview, create, log.
//
// PREVIEW BEFORE ANYTHING IS WRITTEN. Group deletion is manual and one at a time, so a plan this size is not
// an action to take on trust — the preview is the review step and the CSV is the cross-check.
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { parseLevels } from "../../../shared/formModel";
import { abbrevListTitle } from "../../../shared/folderAbbreviation";
import { AbbrevRowDraft } from "../../../shared/abbreviationDraft";
import {
  BulkPlan,
  BulkSegment,
  PlannedGroup,
  planBulkGroups,
  splitPlannedRows,
  toCreateCount,
} from "../../../shared/bulkGroups";
import {
  buildGroupMapRow,
  GroupMapRole,
  GroupMapWriteRow,
  normalizeRoleValue,
  normalizeScope,
  PERSONAS,
} from "../../../shared/groupMapModel";
import { createSiteGroup } from "../../../shared/spGroups";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { Toast, ToastKind } from "../../../shared/toast";

type Props = {
  context: WebPartContext;
  siteUrl: string;
  /**
   * Reported while a run is in flight, so a HOST can hold its own navigation.
   *
   * The run lives in component state and has no resume, so changing step, pressing Next or Back, or
   * leaving the page stops it part-way — and it stops SILENTLY, which is the likely cause of 302 of
   * the 308 groups on the rehearsal site. The component cannot disable a rail it does not own, so it
   * reports instead, the same report-upward shape as `onAbbreviationsMissingChange`.
   *
   * Optional: the standalone Group Management page has nothing to hold, and it must still be able to
   * mount this without knowing about flows.
   */
  onBusyChange?: (busy: boolean) => void;
  /**
   * Fired once a run has finished and its own lists have been re-read.
   *
   * A SEPARATE event from `onBusyChange`, not the falling edge of it, because "the run ended" and
   * "the button is clickable again" are different facts and only one of them means other screens are
   * now stale. The group LIST beside this one reads its mappings at mount, so after a run that wrote
   * 790 rows every group still showed `not mapped` — the same trap as the rail's tick, one screen
   * over, and the one that invites a needless second press.
   */
  onRunComplete?: () => void;
};

type Seg = BulkSegment & { key: string; label: string };

/** Everything one pass of the run reads, kept together so the loop can thread it pass to pass. */
type SegmentData = {
  rows: AbbrevRowDraft[];
  titles: Record<string, number>;
  existingRows: GroupMapWriteRow[];
};

/** What one pass of the run did, so the loop can tell "finished" from "more to do". */
type PassResult = {
  made: number;
  mapped: number;
  already: number;
  failed: number;
  stopped: boolean;
};

/**
 * How many times one press of Run will re-check its own work before giving up and saying so.
 *
 * A pass that writes NOTHING is the finish line: every planned row is on the list, so another pass
 * could only write nothing again. Anything else means the pass found work the pass before it did not
 * see, and stopping there is what left GHO and MHO part-provisioned across four separate presses on
 * 2026-08-24 — each one reporting a clean `0 failed` over a segment that was not done.
 *
 * BOUNDED, because a pass that keeps finding work forever is a defect, not a big segment, and looping
 * on it would hang the page instead of reporting it. Ten passes is far past what any real segment has
 * needed (GHO converged on the fourth), so hitting the cap is a finding and is reported as a warning.
 */
const MAX_PASSES = 10;

/** The personas a bulk run offers, in the order the client's own document lists them. */
// hou_hc joined 2026-08-24 ("not every HOU can upload into Highly Confidential, it will be the same
// pattern as PIC") — the HC Head-of-Unit group is provisioned for every unit exactly as pic_hc is,
// empty until somebody cleared is added. An empty group is a valid end state: the grant is in place
// and applies the moment a person joins, with nothing to re-run.
//
// ⚠ `clevel_global` JOINED 2026-09-02, AND ITS ABSENCE HAD BECOME A REAL BUG, NOT A DESIGN CHOICE.
// It was never offered here, and the ONLY other route to it — the "Create one group" hand-form on
// Group Management — was removed the same day for an unrelated reason (the client asked for it gone
// now that bulk provisioning creates groups for them). Together that left `C_LEVEL_GLOBAL` with NO
// path to existing at all, which is exactly what the client reported ("there isn't any global
// viewer"). `clevel_global`'s persona is `scope: "segment"`, so ticking it plans ONE group per run —
// but `suggestGroupName` returns the bare literal "C_LEVEL_GLOBAL" for it regardless of which
// segment's run creates it (see the comment there), and `planBulkGroups`'s existing-title check
// already makes a second segment's run MAP rather than re-create it. No special "once, site-wide"
// logic was needed — the general mechanism already does the right thing once this is ticked.
const OFFERED = [
  "pic",
  "hou",
  "hou_hc",
  "employee",
  "pic_hc",
  "employee_hc",
  "hod",
  "clevel_segment",
  "clevel_global",
];

/* ⚠ DISPLAY-ONLY, LOCAL TO THIS CHECKLIST (client's mockup, 2026-09-03: "Change the entire Create
   Group to this UI. Follow the exact copy, exact list"). This does NOT rename `PERSONAS` in
   `shared/groupMapModel.ts` — that array's `family`/`label` are read by Folder Access's tier
   picker, Quick Search's persona summary, and CSV export headers elsewhere, so renaming it there
   would relabel "SDG Employee" to "Viewer" everywhere at once, which is a wider change than one
   screen's copy. `ROW_ORDER` also reorders the checklist to match the client's list; `clevel_global`
   is APPENDED rather than dropped — it has no route to existing anywhere else (fixed 2026-09-02,
   "there isn't any global viewer"), and the client's list simply did not show a ninth row. */
const ROW_ORDER = [
  "pic",
  "pic_hc",
  "hou",
  "hou_hc",
  "employee",
  "employee_hc",
  "hod",
  "clevel_segment",
  "clevel_global",
];
const ROW_DISPLAY: Record<string, { family: string; label: string }> = {
  pic: { family: "PIC", label: "Upload only" },
  pic_hc: {
    family: "PIC (HC)",
    label: "Upload only, highly confidential documents",
  },
  hou: {
    family: "Head of unit",
    label: "Approve, upload, delete and share, own unit",
  },
  hou_hc: {
    family: "Head of unit (HC)",
    label:
      "Approve, upload, delete and share, own unit, highly confidential documents",
  },
  employee: { family: "Viewer", label: "View only" },
  employee_hc: {
    family: "Viewer (HC)",
    label: "View only, highly confidential documents",
  },
  hod: {
    family: "Head of department",
    label: "View, delete and share, department-wide",
  },
  clevel_segment: { family: "C-level", label: "View, one segment" },
  clevel_global: { family: "C-level (Global)", label: "View, every segment" },
};

const s: Record<string, React.CSSProperties> = {
  card: {
    border: "1px solid #e1e1e1",
    borderRadius: 6,
    padding: 16,
    marginBottom: 20,
    background: "#fafafa",
  },
  head: { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  label: {
    display: "block",
    fontSize: 12,
    fontWeight: 600,
    color: "#333",
    margin: "12px 0 4px",
  },
  select: {
    width: "100%",
    maxWidth: 460,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
  },
  /* ⚠ BLACK, from the client's mockup (2026-09-04) — it was the product green `#0f6c3f`. This is the
     same call as the decision dialog's `approveDark` in `Requests.tsx`: the mockup for THIS screen
     shows black, so this screen gets black. Green is still the primary elsewhere; if the two are ever
     meant to agree, that is a product-wide decision rather than a per-mockup one. */
  btn: {
    padding: "8px 18px",
    fontSize: 13,
    background: "#1b1b1b",
    color: "#fff",
    border: "1px solid #1b1b1b",
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
    padding: "6px 12px",
    fontSize: 12,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  hint: { fontSize: 11.5, color: "#666", marginTop: 6, lineHeight: 1.5 },
  /* ── The roles list, to the client's mockup (2026-09-04) ───────────────────────────────────────
     A single bordered box with a divider between rows, each row a padded line: checkbox, then the
     role in bold with its description beneath.

     ⚠ ONE COLUMN, NOT `auto-fit`. It was `repeat(auto-fit, minmax(min(100%, 280px), 1fr))`, so on a
     wide screen the eight roles reflowed into TWO columns and the reading order ran down the left and
     back up the right. `ROW_ORDER` is a deliberate sequence (PIC → PIC (HC) → Head of unit → …), so
     it must not be reflowed. */
  ticks: {
    border: "1px solid #e1e1e1",
    borderRadius: 8,
    background: "#fff",
    overflow: "hidden",
  },
  tick: {
    display: "flex",
    alignItems: "flex-start",
    gap: 10,
    fontSize: 13,
    padding: "12px 14px",
    cursor: "pointer",
    /* The divider sits on the TOP of each row and is suppressed on the first, so the box never ends
       on a stray rule — simpler than `:last-child`, which an inline style cannot reach. */
    borderTop: "1px solid #ececec",
  },
  /** A small, letter-spaced caption — the mockup's "ROLES TO SET UP". */
  sectionLabel: {
    display: "block",
    fontSize: 11,
    fontWeight: 600,
    letterSpacing: ".06em",
    textTransform: "uppercase",
    color: "#6b6b6b",
    margin: "18px 0 6px",
  },
  /** The grey line under each role name. */
  tickHint: { display: "block", fontSize: 12, color: "#6b6b6b", marginTop: 2 },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: 12.5,
    marginTop: 10,
  },
  th: {
    textAlign: "left",
    padding: "6px 8px",
    borderBottom: "1px solid #d9d9d9",
    fontWeight: 600,
    color: "#605e5c",
  },
  td: {
    padding: "5px 8px",
    borderBottom: "1px solid #f3f2f1",
    fontFamily: "Consolas, monospace",
  },
  scroll: {
    maxHeight: 340,
    overflowY: "auto",
    border: "1px solid #ececec",
    borderRadius: 4,
    background: "#fff",
    marginTop: 10,
  },
  warnBox: {
    padding: "10px 12px",
    border: "1px solid #f2c9a0",
    background: "#fff8f0",
    borderRadius: 4,
    fontSize: 12,
    color: "#8a4b00",
    lineHeight: 1.5,
    marginTop: 12,
  },
  okBox: {
    padding: "10px 12px",
    border: "1px solid #b7dcc4",
    background: "#f3faf5",
    borderRadius: 4,
    fontSize: 12,
    color: "#0f6c3f",
    lineHeight: 1.5,
    marginBottom: 12,
  },
  // Matches reconciliation's log panel (FolderManager `logBox`) rather than a terminal. Two run logs on
  // adjacent screens should not look like different products — and the dark box read as a developer
  // console, which is exactly the impression an admin tool should not give.
  logBox: {
    marginTop: 16,
    maxHeight: 260,
    overflowY: "auto",
    background: "#f5f5f5",
    borderRadius: 6,
    padding: "12px 16px",
    fontFamily: "Consolas, monospace",
    fontSize: 11.5,
    lineHeight: 1.6,
  },
  // Colour by outcome, since a 700-line log is scanned for the failures. Same palette as the recon tabs.
  logOk: { color: "#0f6c3f" },
  logBad: { color: "#d13438", fontWeight: 600 },
  logDim: { color: "#605e5c" },
  pill: {
    fontSize: 11,
    padding: "1px 7px",
    borderRadius: 10,
    background: "#eef4ff",
    border: "1px solid #cfe0ff",
    color: "#1b4b8a",
  },
};

const GET = { Accept: "application/json;odata=nometadata" };

/**
 * Retry a write that SharePoint throttled.
 *
 * ~700 group creations will meet 429 or 503, and without this a run dies part-way and the admin cannot tell
 * what was made. Non-throttle failures are NOT retried: a duplicate name or a permission error fails
 * identically the second time, so retrying only delays the report.
 */
async function withRetry<T>(fn: () => Promise<T>, attempts = 4): Promise<T> {
  let lastErr: Error | undefined;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (e) {
      const msg = (e as Error).message ?? "";
      if (msg.indexOf("429") === -1 && msg.indexOf("503") === -1) throw e;
      lastErr = e as Error;
      await new Promise((resolve) =>
        setTimeout(resolve, 2000 * Math.pow(2, i)),
      );
    }
  }
  throw lastErr ?? new Error("throttled");
}

export default function BulkGroupProvisioner({
  context,
  siteUrl,
  onBusyChange,
  onRunComplete,
}: Props): React.ReactElement {
  const [segments, setSegments] = useState<Seg[] | undefined>(undefined);
  const [chosen, setChosen] = useState("");
  /**
   * Guards `loadForSegment` against out-of-order resolution — found live 2026-08-26, switching from
   * GHO to MHO quickly enough that GHO's (slower) read resolved AFTER MHO's had already replaced it.
   * `seg`/`plan` were computed against the NEWLY chosen segment while `rows`/`existingRows` were still
   * GHO's, producing a hybrid: group names correctly prefixed `MHO_...` but built from GHO's abbrevi-
   * ation chain, matched as "already there" against GHO's real groups of a DIFFERENT name — and the
   * Create button is gated on `existingRows` being defined, not on `loading`, so it was clickable
   * against that wrong plan the whole time the race window was open, not merely a display glitch.
   * Every `loadForSegment` call takes a ticket; a response is applied only if its ticket is still the
   * most recent one issued, so a slow, superseded read can never overwrite a newer selection's data.
   */
  const loadSeq = useRef(0);
  const [rows, setRows] = useState<AbbrevRowDraft[] | undefined>(undefined);
  /** Existing site groups by lower-cased title, so one already there is MAPPED rather than re-created. */
  const [titles, setTitles] = useState<Record<string, number>>({});
  /**
   * The Group Map rows already on the list — what a re-run must NOT write again.
   *
   * `undefined` means the list could not be READ, never that it is empty, and the two must not
   * collapse into one: an unreadable list looks exactly like an empty one from here, and treating
   * it as empty is the 642-row bug arriving by a second route. This is the deliberate fail-CLOSED
   * case (as with `canOfferFolderDelete`) — everywhere else in this codebase an unreadable list
   * fails open, because the cost is a form that takes itself out of service for a minute. The cost
   * here is hundreds of duplicate rows that nothing on any screen will ever show you.
   */
  const [existingRows, setExistingRows] = useState<
    GroupMapWriteRow[] | undefined
  >(undefined);
  const [picked, setPicked] = useState<string[]>(OFFERED);
  // Collapsed by default (client's mockup, 2026-09-03: a "View all N groups" button reveals the
  // table; a segment can run to hundreds of rows and the preview COUNT above already answers most
  // questions without the full list on screen). Reset whenever a new plan is computed, so switching
  // segments never leaves a stale table open.
  const [tableOpen, setTableOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<string[]>([]);
  /**
   * Set by Stop, read by the run loop. A REF, not state: the loop is one long async function and a
   * state value it closed over at render time would stay false however many times Stop was pressed —
   * the same stale-closure trap as the counters below.
   */
  const stopRef = useRef(false);
  const [stopping, setStopping] = useState(false);
  const [toast, setToast] = useState<
    { kind: ToastKind; text: string } | undefined
  >(undefined);

  useEffect(() => {
    if (onBusyChange) onBusyChange(busy);
  }, [busy]);

  /**
   * Closing or reloading the tab mid-run stops it part-way, and until now said nothing at all.
   *
   * The browser prompt is the whole mitigation — there is no resume to offer — but the run is
   * repeatable since 1.0.151.0, so a stopped run is finished by pressing Run again rather than
   * being unrecoverable. Same guard as the staged batches in `BulkUpload.tsx`.
   */
  useEffect(() => {
    if (!busy) return undefined;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [busy]);

  const seg = (segments ?? []).filter((x) => x.key === chosen)[0];
  const abbrevList = (): string => encodeURIComponent(abbrevListTitle());
  const groupMapList = (): string =>
    encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap));

  /* ── Segments ─────────────────────────────────────────────────────────────── */

  useEffect(() => {
    const load = async (): Promise<void> => {
      await primeNames(context.spHttpClient, siteUrl).catch(() => undefined);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items` +
          `?$select=Title,ModeLabel,TermSetGuid,StagingFolder,Levels&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!res.ok) throw new Error(`config HTTP ${res.status}`);
      const data = await res.json();
      setSegments(
        ((data.value ?? []) as Array<Record<string, string>>)
          .filter((r) => (r.TermSetGuid ?? "").trim())
          .map((r) => ({
            key: (r.Title ?? "").trim(),
            label: (r.ModeLabel ?? r.Title ?? "").trim(),
            code: (r.StagingFolder ?? "").trim(),
            termSetGuid: (r.TermSetGuid ?? "").trim(),
            // PERMISSIONED tiers only. A below-Unit tier carries no group, and letting one through would
            // plan groups for Year and Document Type.
            levelNames: parseLevels(r.Levels ?? "")
              .filter((l) => l.permissioned !== false)
              .map((l) => l.label),
          })),
      );
    };
    // Unreadable ≠ none: `undefined` renders as "could not read", never as "this site has no segments".
    load().catch((e) => {
      setSegments(undefined);
      setToast({
        kind: "err",
        text: `Could not read the segments — ${(e as Error).message}`,
      });
    });
  }, [siteUrl]);

  /* ── Rows + existing groups for the chosen segment ────────────────────────── */

  /**
   * Every Group Map row on the list, so the run can tell a mapping that already exists from one it
   * still has to write. Throws rather than returning `[]` — see `existingRows`.
   *
   * PAGED, and that is not defensive decoration. `$top` caps a page, it does not raise the 5,000-item
   * list view threshold, so one segment on the client's site (~790 rows) plus a second segment puts
   * this list past a single page. A truncated read reports the rows it could not see as absent, which
   * is precisely the state that writes them all a second time — the bug, wearing the fix's clothes.
   *
   * Scope and Target are asked for and the request retried without them, exactly as GroupMapBuilder
   * does: a $select naming a column that does not exist fails the WHOLE request with HTTP 400
   * (CLAUDE.md #11), so a site that predates those columns would otherwise be unable to run at all.
   */
  const loadGroupMapRows = async (): Promise<GroupMapWriteRow[]> => {
    const base = `${siteUrl}/_api/web/lists/getbytitle('${groupMapList()}')/items`;
    const full = "Id,GroupId,GroupName,Segment,UnitTermGuid,Role,Scope,Target";
    const lean = "Id,GroupId,GroupName,Segment,UnitTermGuid,Role";
    let select = full;
    let url = `${base}?$select=${select}&$top=2000`;
    const out: GroupMapWriteRow[] = [];
    // Bounded: 50 pages is 100,000 rows, far past anything this list can hold, and a malformed
    // nextLink that pointed at itself would otherwise spin the page during a 700-group run.
    for (let page = 0; page < 50 && url; page++) {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: GET },
      );
      if (!res.ok) {
        if (page === 0 && select === full) {
          select = lean;
          url = `${base}?$select=${select}&$top=2000`;
          continue;
        }
        throw new Error(`HTTP ${res.status}`);
      }
      const data = await res.json();
      for (const r of (data.value ?? []) as Array<Record<string, string>>) {
        out.push({
          GroupId: r.GroupId ?? "",
          GroupName: r.GroupName ?? "",
          Segment: r.Segment ?? "",
          UnitTermGuid: r.UnitTermGuid ?? "",
          // A hand-authored row saying "UPLOADER" means UPL. Left raw it would not match the row
          // the run is about to write, and the run would write a second copy of it.
          Role: normalizeRoleValue(r.Role ?? "") as GroupMapRole,
          // Blank reads as Folder — every row written before the column existed is one.
          Scope: normalizeScope(r.Scope),
          Target: r.Target ?? "",
        });
      }
      const next = (data["odata.nextLink"] ??
        data["@odata.nextLink"] ??
        "") as string;
      url = next && next.indexOf("http") === 0 ? next : "";
    }
    return out;
  };

  /**
   * Everything one pass of the run needs, READ AND RETURNED rather than pushed into state.
   *
   * The run is a loop now (see `run`), and every pass after the first has to see what the pass before it
   * wrote. State cannot carry that: a value set during an async run is not visible to the closure already
   * running — the same stale-closure trap the local counters below exist to dodge. So the reader returns
   * its data and the loop threads it from one pass to the next; `loadForSegment` is the thin wrapper that
   * puts the same data on screen.
   */
  const readSegmentData = async (target: Seg): Promise<SegmentData> => {
    // Walk only as deep as there are permissioned levels — the same rule as the abbreviations page.
    const maxDepth = target.levelNames.length;
    const nodes: AbbrevRowDraft[] = [];
    const walk = async (parentId: string, depth: number): Promise<void> => {
      if (depth > maxDepth) return;
      const url = parentId
        ? `${siteUrl}/_api/v2.1/termStore/sets/${target.termSetGuid}/terms/${parentId}/children?$select=id,labels`
        : `${siteUrl}/_api/v2.1/termStore/sets/${target.termSetGuid}/children?$select=id,labels`;
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok)
        throw new Error(`the term store returned HTTP ${res.status}`);
      const kids = (
        ((await res.json()).value ?? []) as Array<{
          id?: string;
          labels?: Array<{ name?: string }>;
        }>
      ).map((t) => ({
        id: t.id ?? "",
        label: (t.labels ?? [])[0]?.name ?? "",
      }));
      for (const k of kids) {
        if (!k.id) continue;
        nodes.push({
          termGuid: k.id,
          label: k.label,
          level: target.levelNames[depth - 1] ?? `Level ${depth}`,
          parentGuid: parentId,
          abbreviation: "",
        });
        await walk(k.id, depth + 1);
      }
    };
    await walk("", 1);

    // Codes. A term with no row keeps its blank code, and the planner skips it naming the reason.
    const cur: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${abbrevList()}')/items?$select=TermGuid,Abbreviation&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    if (!cur.ok)
      throw new Error(`${abbrevListTitle()} returned HTTP ${cur.status}`);
    const codes: Record<string, string> = {};
    for (const r of ((await cur.json()).value ?? []) as Array<{
      TermGuid?: string;
      Abbreviation?: string;
    }>) {
      const k = (r.TermGuid ?? "").trim().toLowerCase();
      if (k) codes[k] = (r.Abbreviation ?? "").trim();
    }
    const readRows = nodes.map((n) => ({
      ...n,
      abbreviation: codes[n.termGuid.toLowerCase()] ?? "",
    }));

    // Existing groups, by title.
    //
    // `$top=5000`, NOT 1000, and the difference is a silent failure rather than a slow one. A capped
    // read is indistinguishable from the group being ABSENT (memory `sp-capped-read-reads-as-absent`),
    // and here "absent" means the run tries to CREATE a group that already exists, fails on the
    // duplicate name, and skips every one of its mapping rows with it. One provisioned segment is ~320
    // groups, so 1000 is under three segments — the client's live tenant crosses it. The shared readers
    // in `spGroups.ts` were raised to 5000 on 2026-08-21 for exactly this; this copy was missed.
    const gs: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/sitegroups?$select=Id,Title&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: GET },
    );
    const readTitles: Record<string, number> = {};
    if (gs.ok) {
      for (const g of ((await gs.json()).value ?? []) as Array<{
        Id: number;
        Title?: string;
      }>) {
        readTitles[(g.Title ?? "").trim().toLowerCase()] = g.Id;
      }
    }

    // Read LAST, so a failure here leaves the preview intact and gates only the Run button. The
    // admin can still review the plan and export the CSV while the reason is on screen.
    return {
      rows: readRows,
      titles: readTitles,
      existingRows: await loadGroupMapRows(),
    };
  };

  const loadForSegment = async (target: Seg): Promise<void> => {
    const ticket = ++loadSeq.current;
    setLoading(true);
    setRows(undefined);
    setExistingRows(undefined);
    try {
      const d = await readSegmentData(target);
      if (loadSeq.current !== ticket) return; // superseded by a later segment switch — discard
      setRows(d.rows);
      setTitles(d.titles);
      setExistingRows(d.existingRows);
    } catch (e) {
      if (loadSeq.current !== ticket) return;
      setRows(undefined);
      setExistingRows(undefined);
      setToast({
        kind: "err",
        text: `Could not read this segment — ${(e as Error).message}`,
      });
    } finally {
      if (loadSeq.current === ticket) setLoading(false);
    }
  };

  // `existingRows` is handed to the planner as well as to the dedupe: it is what lets a unit whose
  // folder code was RENAMED be recognised by its term instead of by a group name derived from the old
  // code. Without it a rename reads as a new unit and this run creates a second full set of groups.
  const plan: BulkPlan | undefined =
    seg && rows
      ? planBulkGroups(
          seg,
          rows,
          picked,
          Object.keys(titles),
          existingRows ?? [],
        )
      : undefined;

  /* ── The run ──────────────────────────────────────────────────────────────── */

  const postRow = async (row: GroupMapWriteRow): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${groupMapList()}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
        },
        body: JSON.stringify({ Title: row.GroupName || row.GroupId, ...row }),
      },
    );
    if (!res.ok) {
      throw new Error(
        `HTTP ${res.status} ${(await res.text().catch(() => "")).slice(0, 120)}`,
      );
    }
  };

  // `target` is PASSED, never read from the `seg` closure: a pass runs long after the render that
  // created it, and the segment it is provisioning must not be able to drift out from under it.
  const rowsFor = (
    g: PlannedGroup,
    groupId: string,
    target: Seg,
  ): GroupMapWriteRow[] => {
    const p = PERSONAS.filter((x) => x.key === g.personaKey)[0];
    if (!p) return [];
    return p.roles.map((r) =>
      buildGroupMapRow({
        groupId,
        groupName: g.name,
        role: r,
        segmentGuid: target.termSetGuid,
        tierGuid: g.tierGuid,
      }),
    );
  };

  /**
   * ONE pass over the plan: create what is missing, write the mapping rows that are not there yet.
   *
   * Takes its data as an argument rather than reading `plan`/`titles`/`existingRows` from the render
   * closure, because the loop in `run` calls it repeatedly with FRESHLY read data — a pass working from
   * the render-time values would re-do the first pass forever and never converge.
   */
  const runOnePass = async (
    target: Seg,
    data: SegmentData,
    say: (t: string) => void,
  ): Promise<PassResult> => {
    let made = 0;
    let mapped = 0;
    let already = 0;
    let failed = 0;
    let stopped = false;

    const passPlan = planBulkGroups(
      target,
      data.rows,
      picked,
      Object.keys(data.titles),
      data.existingRows,
    );
    // The rows the list holds, grown by every row this pass writes. A LOCAL copy for the same reason
    // the counters are local — state set during a run is not visible to the closure reading it — but
    // also because two groups can legitimately share a mapping and the second must see the first.
    const seen: GroupMapWriteRow[] = data.existingRows.slice();

    say(
      `${target.label}: ${passPlan.groups.length} planned, ${toCreateCount(passPlan)} to create.`,
    );
    for (const g of passPlan.groups) {
      // Checked BETWEEN groups, never inside one: a group whose rows are half written is the state
      // that needs a human to look at it, and this is a stop, not an abort. What has landed stays.
      if (stopRef.current) {
        stopped = true;
        say(
          `■ Stopped by you. Everything above is done and stays done — press Run again to finish.`,
        );
        break;
      }
      // A unit whose code was renamed is already provisioned under its OLD name, so the run maps to
      // THAT group. Creating `g.name` instead is exactly the duplicate set this closes.
      const useName = g.existingName ?? g.name;
      let id = data.titles[useName.toLowerCase()];
      if (id === undefined) {
        try {
          const created = await withRetry(() =>
            createSiteGroup(context.spHttpClient, siteUrl, g.name),
          );
          id = created.id;
          made++;
          say(`+ ${g.name}`);
        } catch (e) {
          // Its rows are skipped with it: a row naming a group that does not exist grants nothing and
          // leaves a mapping nobody can see on any screen. The next pass re-reads and tries again.
          failed++;
          say(`✗ ${g.name} — ${(e as Error).message}`);
          continue;
        }
      } else if (g.existingName) {
        // Named in the log, because the admin will not otherwise understand why a group they did not
        // ask for is being mapped — and because renaming it is the tidy-up this reports.
        say(
          `= ${g.existingName} (already there under its old name; code now says ${g.name})`,
        );
      } else {
        say(`= ${g.name} (already existed — mapping only)`);
      }
      // Rows the list already holds are NOT written again. Group creation was always idempotent, so
      // a second press read as safe while it re-wrote every mapping behind it — 642 rows on the
      // rehearsal site. Per row rather than per group, so a pass stopped part-way is finished by the
      // next one instead of being skipped as "that group is done".
      const { fresh, duplicate } = splitPlannedRows(
        seen,
        rowsFor({ ...g, name: useName }, String(id), target),
      );
      if (duplicate.length > 0) {
        already += duplicate.length;
        say(
          `  = ${duplicate.length} mapping(s) already there (${duplicate.map((r) => r.Role).join(", ")})`,
        );
      }
      for (const row of fresh) {
        try {
          await withRetry(() => postRow(row));
          mapped++;
          // Only after it lands. A failed write must stay writable by the next pass.
          seen.push(row);
        } catch (e) {
          failed++;
          say(`  ✗ ${g.name} → ${row.Role} — ${(e as Error).message}`);
        }
      }
    }
    return { made, mapped, already, failed, stopped };
  };

  /**
   * One press, run to completion — passes repeat until one of them finds nothing left to write.
   *
   * WHY THIS LOOPS AT ALL (2026-08-24, client: "Can't we keep running untill eveyrthing is there? … So
   * you expect that client to click multiple times manually?"). Rebuilding GHO took FOUR presses to
   * settle — 250 rows written, then 18, then 68, then 0 — and every one of those runs reported
   * `0 failed` over a segment that was demonstrably not finished. Whatever makes a pass see less than
   * the whole truth, the finish line is not "a pass completed", it is A PASS THAT WROTE NOTHING, and
   * only a further pass can establish that. Leaving it to the admin means a segment is finished when
   * somebody happens to press once more, which is not a condition anyone can check.
   *
   * SAFE TO REPEAT BY CONSTRUCTION, which is what makes this legitimate rather than brute force: every
   * group is matched before it is created and every row is checked against the list before it is written
   * (`splitPlannedRows` / `isDuplicateRow`, keyed on GroupId + Scope + Target + term + Role). An extra
   * pass over finished work writes nothing and costs one read.
   *
   * THE DATA IS RE-READ BETWEEN PASSES, never carried over. The whole reason a further pass finds work is
   * that the previous read did not describe the list correctly, so re-using it would reproduce the same
   * blind spot and the loop would converge on the same wrong answer, only faster.
   */
  const run = async (): Promise<void> => {
    if (!plan || !seg || !existingRows || !rows) return;
    stopRef.current = false;
    setStopping(false);
    setBusy(true);
    // A LOCAL log, mirrored into state. Counters read from state during a run would see their render-time
    // value and report zero — the same stale-closure trap as reconciliation's counts.
    const lines: string[] = [];
    const say = (t: string): void => {
      lines.push(t);
      setLog([...lines]);
    };

    // The first pass uses what is already on screen — it is the plan the admin just reviewed, and
    // re-reading it here would make Run mean something other than the preview above it.
    let data: SegmentData = { rows, titles, existingRows };
    let totalMade = 0;
    let totalMapped = 0;
    let totalFailed = 0;
    // The LAST pass's count, never a sum: "already there" describes the state of the list at the end,
    // and adding up every pass's view of it would report the same row several times over.
    let finalAlready = 0;
    let stopped = false;
    let settled = false;
    let passes = 0;

    for (let pass = 1; pass <= MAX_PASSES; pass++) {
      passes = pass;
      if (pass > 1)
        say(`— Pass ${pass}: checking whether anything is still missing …`);
      const r = await runOnePass(seg, data, say);
      totalMade += r.made;
      totalMapped += r.mapped;
      totalFailed += r.failed;
      finalAlready = r.already;
      stopped = r.stopped;
      if (stopped) break;
      // NOTHING WRITTEN IS THE FINISH LINE. A pass that created no group and wrote no row found every
      // planned mapping already on the list, so a further one could only find the same.
      if (r.made === 0 && r.mapped === 0) {
        settled = true;
        break;
      }
      if (pass === MAX_PASSES) break;
      try {
        data = await readSegmentData(seg);
      } catch (e) {
        // Reported, never swallowed, and it ENDS the run: another pass over data we could not refresh
        // would be the previous pass again, and its "already there" counts would be a guess.
        say(
          `⚠ Could not re-read ${seg.label} to check the last pass — ${(e as Error).message}`,
        );
        break;
      }
    }

    const passWord = passes === 1 ? "1 pass" : `${passes} passes`;
    say(
      `${stopped ? "Stopped" : "Done"} after ${passWord}. ${totalMade} created, ` +
        `${totalMapped} mappings written, ${finalAlready} already there (not re-written), ` +
        `${totalFailed} failed.`,
    );
    if (!stopped && !settled) {
      // Said out loud, because this is the one ending where the segment may still be unfinished and
      // the counts above look exactly like a clean run.
      say(
        `⚠ Still finding work after ${MAX_PASSES} passes — this segment may not be complete. ` +
          `Press Run again, and if it keeps writing rows, something needs looking at.`,
      );
    }
    setToast({
      // A stopped or unsettled run is a WARNING however clean its counts look — it is unfinished, and a
      // green toast over a half-provisioned segment is how 302 of 308 goes unnoticed a second time.
      kind: totalFailed > 0 || stopped || !settled ? "warn" : "ok",
      text:
        (stopped ? `Stopped part-way. ` : "") +
        `${totalMade} groups created and ${totalMapped} mappings written for ${seg.label}` +
        (passes > 1 ? ` over ${passWord}` : "") +
        "." +
        (stopped
          ? " Press Run again to finish the rest — nothing already written is touched."
          : "") +
        (!stopped && settled
          ? " A final pass found nothing left to write, so this segment is complete."
          : "") +
        (!stopped && !settled
          ? ` Still finding work after ${MAX_PASSES} passes — press Run again and check the log.`
          : "") +
        (finalAlready > 0
          ? ` ${finalAlready} mapping(s) were already there and were left alone.`
          : "") +
        (totalFailed > 0 ? ` ${totalFailed} failed — see the log.` : "") +
        " Nothing is granted until Folder Reconciliation runs.",
    });
    setBusy(false);
    setStopping(false);
    // stopRef is cleared at the START of the next run, not here: after the awaits above the linter
    // cannot prove the ref has not moved, and a reset that lands after a fresh press would arm a run
    // nobody asked to stop.
    // Re-read, so the screen shows what the run made instead of the plan it started from.
    await loadForSegment(seg).catch(() => undefined);
    // Told LAST, so anything listening re-reads a list this run has finished writing to.
    if (onRunComplete) onRunComplete();
  };

  const exportCsv = (): void => {
    if (!plan) return;
    const head = "Group name,Persona,Scope,Status";
    const body = plan.groups
      .map((g) =>
        [
          g.name,
          g.personaKey,
          g.scope,
          g.existingName
            ? `exists as ${g.existingName}`
            : g.exists
              ? "exists"
              : "to create",
        ].join(","),
      )
      .join("\n");
    const blob = new Blob([`${head}\n${body}\n`], {
      type: "text/csv;charset=utf-8",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `bulk-groups-${(seg?.code ?? "segment").toLowerCase()}.csv`;
    a.click();
  };

  /* ── Render ───────────────────────────────────────────────────────────────── */

  return (
    <div style={s.card}>
      {/* ⚠ COPY REPLACED VERBATIM (client's mockup, 2026-09-03: "Follow the exact copy, exact
          list"). The stale "set on Folder Access" reference is gone WITH this rewrite — Folder
          Access itself was retired 2026-08-23, and membership editing moved to Group Management;
          the client's own shorter sentence happens to drop the dangling reference along with it. */}
      <p style={s.head}>Create group</p>
      <div style={s.okBox}>
        Set up the groups a segment needs. Creating a group also writes its
        folder mapping — permissions apply once you run folder reconciliation.
      </div>

      <label style={s.label} htmlFor="bg-seg">
        Segment
      </label>
      <select
        id="bg-seg"
        style={s.select}
        value={chosen}
        disabled={busy}
        onChange={(e) => {
          setChosen(e.target.value);
          setLog([]);
          setTableOpen(false);
          const t = (segments ?? []).filter((x) => x.key === e.target.value)[0];
          if (t) loadForSegment(t).catch(() => undefined);
        }}
      >
        <option value="">— select a segment —</option>
        {(segments ?? []).map((x) => (
          <option key={x.key} value={x.key}>
            {x.label}
          </option>
        ))}
      </select>
      {segments === undefined && (
        <p style={s.hint}>
          The segment list could not be read, so none are offered here.
        </p>
      )}

      {seg && (
        <>
          <label style={s.sectionLabel}>Roles to set up</label>
          <div style={s.ticks}>
            {/* ⚠ `first` IS COUNTED OVER THE RENDERED ROWS, NOT THE INDEX IN `ROW_ORDER`. A persona
                whose key is missing from `PERSONAS` or `ROW_DISPLAY` renders nothing, so keying the
                divider on `i === 0` would leave a rule across the top of the box the moment the first
                entry is the one that dropped out. */}
            {(() => {
              let first = true;
              return ROW_ORDER.map((k) => {
                const p = PERSONAS.filter((x) => x.key === k)[0];
                const d = ROW_DISPLAY[k];
                if (!p || !d) return undefined;
                const style = first ? { ...s.tick, borderTop: "none" } : s.tick;
                first = false;
                return (
                  <label key={k} style={style}>
                    <input
                      type="checkbox"
                      checked={picked.indexOf(k) !== -1}
                      disabled={busy}
                      style={{ marginTop: 2 }}
                      onChange={(e) =>
                        setPicked(
                          e.target.checked
                            ? [...picked, k]
                            : picked.filter((x) => x !== k),
                        )
                      }
                    />
                    <span>
                      <strong>{d.family}</strong>
                      <span style={s.tickHint}>{d.label}</span>
                    </span>
                  </label>
                );
              });
            })()}
          </div>
          <p style={s.hint}>
            A term with no folder code is skipped automatically, since it
            wouldn&rsquo;t have a folder to grant.
          </p>
        </>
      )}

      {loading && <p style={s.hint}>Reading the term store and the codes…</p>}

      {plan && (
        <>
          {/* ⚠ THE "Preview — N groups..." SUMMARY LINE IS GONE (client's mockup, 2026-09-03) — the
              count now lives on the "View all N groups" toggle button, in the SAME row as Create and
              Export CSV, matching the mockup exactly. The table itself moved below that row, and only
              renders while `tableOpen`. */}
          {plan.skipped.length > 0 && (
            <div style={s.warnBox}>
              <strong>
                {plan.skipped.length} term{plan.skipped.length === 1 ? "" : "s"}{" "}
                left out.
              </strong>{" "}
              Give them a code on <strong>CRS Term Abbreviations</strong>, then
              run this again.
              <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                {plan.skipped.slice(0, 12).map((x) => (
                  <li key={`${x.level}-${x.label}`}>
                    {x.label} ({x.level}) — {x.reason}
                  </li>
                ))}
              </ul>
              {plan.skipped.length > 12 && (
                <p style={{ margin: "6px 0 0" }}>
                  …and {plan.skipped.length - 12} more.
                </p>
              )}
            </div>
          )}

          {existingRows === undefined && !loading && (
            <div style={s.warnBox}>
              <strong>
                {cachedListTitle(LIST_SUFFIX.groupMap)} could not be read, so
                this run is held.
              </strong>{" "}
              Without it there is no way to tell a mapping that already exists
              from one that is missing, and the run would write a second copy of
              every row it cannot see. Nothing on any screen shows a row twice,
              so that is a mistake you would not find. Re-pick the segment to
              try the read again.
            </div>
          )}

          <div
            style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}
          >
            {/* "Create missing groups" (client's mockup, 2026-09-03) — was "Create N group(s) and
                map M". The RUN is unchanged: it still creates whatever is missing and maps every
                group in the plan, whether new or already there. */}
            <button
              style={
                busy || plan.groups.length === 0 || !existingRows
                  ? s.off
                  : s.btn
              }
              disabled={busy || plan.groups.length === 0 || !existingRows}
              onClick={() => {
                run().catch(() => setBusy(false));
              }}
            >
              {busy ? "Working…" : "Create missing groups"}
            </button>
            <button style={s.ghost} disabled={busy} onClick={exportCsv}>
              Export CSV
            </button>
            {/* Collapsed by default, count on the button itself — matches the mockup.
                ⚠ PUSHED TO THE FAR RIGHT (`marginLeft: auto`), as the mockup shows. It is not a peer
                of Create and Export: those two ACT, while this only reveals the table below. Sitting
                in the same cluster made "View all 315 groups" read as a third thing that does
                something. */}
            <button
              type="button"
              style={{ ...s.ghost, marginLeft: "auto" }}
              disabled={busy}
              onClick={() => setTableOpen(!tableOpen)}
            >
              {tableOpen
                ? `Hide ${plan.groups.length} groups`
                : `View all ${plan.groups.length} groups`}
            </button>
            {/* Leaving the page is the only other way out of a run, and holding the flow's navigation
                without offering a way to stop would trap an admin for the length of a 300-group run.
                Safe to offer only because the run is repeatable: it stops between groups and what
                landed stays. */}
            {busy && (
              <button
                style={s.ghost}
                disabled={stopping}
                onClick={() => {
                  stopRef.current = true;
                  setStopping(true);
                }}
              >
                {stopping ? "Stopping after this group…" : "Stop"}
              </button>
            )}
          </div>
          {tableOpen && (
            <div style={s.scroll}>
              <table style={s.table}>
                <thead>
                  <tr>
                    <th style={s.th}>Group</th>
                    <th style={s.th}>Scope</th>
                    <th style={s.th}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {plan.groups.map((g) => (
                    <tr key={g.name}>
                      <td style={s.td}>{g.name}</td>
                      <td style={{ ...s.td, fontFamily: "inherit" }}>
                        {g.scope}
                      </td>
                      {/* THREE statuses. "Already there under its old name" is the one that carries
                          information the admin cannot get anywhere else: the unit is provisioned, so
                          nothing will be created — AND its groups no longer match the folder code,
                          which is the thing they may want to tidy. */}
                      <td
                        style={{
                          ...s.td,
                          fontFamily: "inherit",
                          color: g.existingName
                            ? "#8a4b00"
                            : g.exists
                              ? "#605e5c"
                              : "#0f6c3f",
                        }}
                      >
                        {g.existingName
                          ? `already there as ${g.existingName}`
                          : g.exists
                            ? "already exists"
                            : "new"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {existingRows !== undefined && (
            <p style={s.hint}>
              One press finishes the segment — the run keeps checking its own
              work and writing whatever is still missing, until a pass finds
              nothing left to write. {existingRows.length} mapping row(s) are
              already on {cachedListTitle(LIST_SUFFIX.groupMap)} and will be
              left alone; only what is missing is written, so running it again
              later is always safe.
            </p>
          )}
          {busy && (
            <p style={s.hint}>
              Leave this tab open and stay on this step — the run has no resume,
              and leaving stops it part-way. Nothing already written would be
              lost, and pressing Run again finishes the rest. A large segment
              may take several passes; the log says which pass it is on.
            </p>
          )}
        </>
      )}

      {log.length > 0 && (
        <div style={s.logBox}>
          {log.map((l, i) => (
            <div
              key={i}
              style={
                l.indexOf("✗") !== -1
                  ? s.logBad
                  : l.indexOf("+ ") === 0
                    ? s.logOk
                    : s.logDim
              }
            >
              {l}
            </div>
          ))}
        </div>
      )}

      {toast && (
        <Toast
          kind={toast.kind}
          text={toast.text}
          onDismiss={() => setToast(undefined)}
        />
      )}
    </div>
  );
}
