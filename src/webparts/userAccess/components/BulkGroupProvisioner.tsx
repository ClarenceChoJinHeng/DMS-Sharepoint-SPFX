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

/** The personas a bulk run offers, in the order the client's own document lists them. */
const OFFERED = ["pic", "hou", "employee", "pic_hc", "employee_hc", "hod", "clevel_segment"];

const s: Record<string, React.CSSProperties> = {
  card: { border: "1px solid #e1e1e1", borderRadius: 6, padding: 16, marginBottom: 20, background: "#fafafa" },
  head: { fontWeight: 600, fontSize: 13, margin: "0 0 8px" },
  label: { display: "block", fontSize: 12, fontWeight: 600, color: "#333", margin: "12px 0 4px" },
  select: { width: "100%", maxWidth: 460, boxSizing: "border-box", padding: "7px 10px", fontSize: 13, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff" },
  btn: { padding: "7px 16px", fontSize: 13, background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer" },
  off: { padding: "7px 16px", fontSize: 13, background: "#e6e6e6", color: "#9a9a9a", border: "none", borderRadius: 4, cursor: "not-allowed" },
  ghost: { padding: "6px 12px", fontSize: 12, border: "1px solid #c7c7c7", borderRadius: 4, background: "#fff", cursor: "pointer" },
  hint: { fontSize: 11.5, color: "#666", marginTop: 6, lineHeight: 1.5 },
  tick: { display: "flex", alignItems: "center", gap: 7, fontSize: 13, padding: "3px 0", cursor: "pointer" },
  ticks: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 280px), 1fr))", gap: 2 },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 12.5, marginTop: 10 },
  th: { textAlign: "left", padding: "6px 8px", borderBottom: "1px solid #d9d9d9", fontWeight: 600, color: "#605e5c" },
  td: { padding: "5px 8px", borderBottom: "1px solid #f3f2f1", fontFamily: "Consolas, monospace" },
  scroll: { maxHeight: 340, overflowY: "auto", border: "1px solid #ececec", borderRadius: 4, background: "#fff" },
  warnBox: { padding: "10px 12px", border: "1px solid #f2c9a0", background: "#fff8f0", borderRadius: 4, fontSize: 12, color: "#8a4b00", lineHeight: 1.5, marginTop: 12 },
  okBox: { padding: "10px 12px", border: "1px solid #b7dcc4", background: "#f3faf5", borderRadius: 4, fontSize: 12, color: "#0f6c3f", lineHeight: 1.5, marginBottom: 12 },
  // Matches reconciliation's log panel (FolderManager `logBox`) rather than a terminal. Two run logs on
  // adjacent screens should not look like different products — and the dark box read as a developer
  // console, which is exactly the impression an admin tool should not give.
  logBox: { marginTop: 16, maxHeight: 260, overflowY: "auto", background: "#f5f5f5", borderRadius: 6, padding: "12px 16px", fontFamily: "Consolas, monospace", fontSize: 11.5, lineHeight: 1.6 },
  // Colour by outcome, since a 700-line log is scanned for the failures. Same palette as the recon tabs.
  logOk: { color: "#0f6c3f" },
  logBad: { color: "#d13438", fontWeight: 600 },
  logDim: { color: "#605e5c" },
  pill: { fontSize: 11, padding: "1px 7px", borderRadius: 10, background: "#eef4ff", border: "1px solid #cfe0ff", color: "#1b4b8a" },
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
      await new Promise((resolve) => setTimeout(resolve, 2000 * Math.pow(2, i)));
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
  const [existingRows, setExistingRows] = useState<GroupMapWriteRow[] | undefined>(undefined);
  const [picked, setPicked] = useState<string[]>(OFFERED);
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
  const [toast, setToast] = useState<{ kind: ToastKind; text: string } | undefined>(undefined);

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
  const groupMapList = (): string => encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap));

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
      setToast({ kind: "err", text: `Could not read the segments — ${(e as Error).message}` });
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
        url, SPHttpClient.configurations.v1, { headers: GET },
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
      const next = (data["odata.nextLink"] ?? data["@odata.nextLink"] ?? "") as string;
      url = next && next.indexOf("http") === 0 ? next : "";
    }
    return out;
  };

  const loadForSegment = async (target: Seg): Promise<void> => {
    setLoading(true);
    setRows(undefined);
    setExistingRows(undefined);
    try {
      // Walk only as deep as there are permissioned levels — the same rule as the abbreviations page.
      const maxDepth = target.levelNames.length;
      const nodes: AbbrevRowDraft[] = [];
      const walk = async (parentId: string, depth: number): Promise<void> => {
        if (depth > maxDepth) return;
        const url = parentId
          ? `${siteUrl}/_api/v2.1/termStore/sets/${target.termSetGuid}/terms/${parentId}/children?$select=id,labels`
          : `${siteUrl}/_api/v2.1/termStore/sets/${target.termSetGuid}/children?$select=id,labels`;
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          url, SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } },
        );
        if (!res.ok) throw new Error(`the term store returned HTTP ${res.status}`);
        const kids = (
          ((await res.json()).value ?? []) as Array<{ id?: string; labels?: Array<{ name?: string }> }>
        ).map((t) => ({ id: t.id ?? "", label: (t.labels ?? [])[0]?.name ?? "" }));
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
        SPHttpClient.configurations.v1, { headers: GET },
      );
      if (!cur.ok) throw new Error(`${abbrevListTitle()} returned HTTP ${cur.status}`);
      const codes: Record<string, string> = {};
      for (const r of ((await cur.json()).value ?? []) as Array<{ TermGuid?: string; Abbreviation?: string }>) {
        const k = (r.TermGuid ?? "").trim().toLowerCase();
        if (k) codes[k] = (r.Abbreviation ?? "").trim();
      }
      setRows(nodes.map((n) => ({ ...n, abbreviation: codes[n.termGuid.toLowerCase()] ?? "" })));

      // Existing groups, by title.
      const gs: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/sitegroups?$select=Id,Title&$top=1000`,
        SPHttpClient.configurations.v1, { headers: GET },
      );
      if (gs.ok) {
        const map: Record<string, number> = {};
        for (const g of ((await gs.json()).value ?? []) as Array<{ Id: number; Title?: string }>) {
          map[(g.Title ?? "").trim().toLowerCase()] = g.Id;
        }
        setTitles(map);
      }

      // Read LAST, so a failure here leaves the preview intact and gates only the Run button. The
      // admin can still review the plan and export the CSV while the reason is on screen.
      setExistingRows(await loadGroupMapRows());
    } catch (e) {
      setRows(undefined);
      setExistingRows(undefined);
      setToast({ kind: "err", text: `Could not read this segment — ${(e as Error).message}` });
    } finally {
      setLoading(false);
    }
  };

  const plan: BulkPlan | undefined =
    seg && rows ? planBulkGroups(seg, rows, picked, Object.keys(titles)) : undefined;

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
      throw new Error(`HTTP ${res.status} ${(await res.text().catch(() => "")).slice(0, 120)}`);
    }
  };

  const rowsFor = (g: PlannedGroup, groupId: string): GroupMapWriteRow[] => {
    const p = PERSONAS.filter((x) => x.key === g.personaKey)[0];
    if (!p || !seg) return [];
    return p.roles.map((r) =>
      buildGroupMapRow({
        groupId,
        groupName: g.name,
        role: r,
        segmentGuid: seg.termSetGuid,
        tierGuid: g.tierGuid,
      }),
    );
  };

  const run = async (): Promise<void> => {
    if (!plan || !seg || !existingRows) return;
    stopRef.current = false;
    setStopping(false);
    setBusy(true);
    // A LOCAL log, mirrored into state. Counters read from state during a run would see their render-time
    // value and report zero — the same stale-closure trap as reconciliation's counts.
    const lines: string[] = [];
    const say = (t: string): void => { lines.push(t); setLog([...lines]); };
    let made = 0;
    let mapped = 0;
    let already = 0;
    let failed = 0;
    // The rows the list holds, grown by every row this run writes. A LOCAL copy for the same reason
    // the counters are local — state set during a run is not visible to the closure reading it — but
    // also because two groups can legitimately share a mapping and the second must see the first.
    const seen: GroupMapWriteRow[] = existingRows.slice();

    say(`${seg.label}: ${plan.groups.length} planned, ${toCreateCount(plan)} to create.`);
    let stopped = false;
    for (const g of plan.groups) {
      // Checked BETWEEN groups, never inside one: a group whose rows are half written is the state
      // that needs a human to look at it, and this is a stop, not an abort. What has landed stays.
      if (stopRef.current) {
        stopped = true;
        say(`■ Stopped by you. Everything above is done and stays done — press Run again to finish.`);
        break;
      }
      let id = titles[g.name.toLowerCase()];
      if (id === undefined) {
        try {
          const created = await withRetry(() => createSiteGroup(context.spHttpClient, siteUrl, g.name));
          id = created.id;
          made++;
          say(`+ ${g.name}`);
        } catch (e) {
          // Its rows are skipped with it: a row naming a group that does not exist grants nothing and
          // leaves a mapping nobody can see on any screen.
          failed++;
          say(`✗ ${g.name} — ${(e as Error).message}`);
          continue;
        }
      } else {
        say(`= ${g.name} (already existed — mapping only)`);
      }
      // Rows the list already holds are NOT written again. Group creation was always idempotent, so
      // a second press read as safe while it re-wrote every mapping behind it — 642 rows on the
      // rehearsal site. Per row rather than per group, so a run stopped part-way is finished by the
      // next press instead of being skipped as "that group is done".
      const { fresh, duplicate } = splitPlannedRows(seen, rowsFor(g, String(id)));
      if (duplicate.length > 0) {
        already += duplicate.length;
        say(`  = ${duplicate.length} mapping(s) already there (${duplicate.map((r) => r.Role).join(", ")})`);
      }
      for (const row of fresh) {
        try {
          await withRetry(() => postRow(row));
          mapped++;
          // Only after it lands. A failed write must stay writable by the next press.
          seen.push(row);
        } catch (e) {
          failed++;
          say(`  ✗ ${g.name} → ${row.Role} — ${(e as Error).message}`);
        }
      }
    }
    say(
      `${stopped ? "Stopped" : "Done"}. ${made} created, ${mapped} mappings written, ` +
      `${already} already there (not re-written), ${failed} failed.`,
    );
    setToast({
      // A stopped run is a WARNING however cleanly it stopped — it is unfinished, and a green toast
      // over a half-provisioned segment is how 302 of 308 goes unnoticed a second time.
      kind: failed > 0 || stopped ? "warn" : "ok",
      text:
        (stopped ? `Stopped part-way. ` : "") +
        `${made} groups created and ${mapped} mappings written for ${seg.label}.` +
        (stopped ? " Press Run again to finish the rest — nothing already written is touched." : "") +
        (already > 0 ? ` ${already} mapping(s) were already there and were left alone.` : "") +
        (failed > 0 ? ` ${failed} failed — see the log.` : "") +
        " Nothing is granted until Folder Reconciliation runs.",
    });
    setBusy(false);
    setStopping(false);
    // stopRef is cleared at the START of the next run, not here: after the awaits above the linter
    // cannot prove the ref has not moved, and a reset that lands after a fresh press would arm a run
    // nobody asked to stop.
    // Re-read, so a second press sees what the first made instead of trying to create it again.
    await loadForSegment(seg).catch(() => undefined);
    // Told LAST, so anything listening re-reads a list this run has finished writing to.
    if (onRunComplete) onRunComplete();
  };

  const exportCsv = (): void => {
    if (!plan) return;
    const head = "Group name,Persona,Scope,Status";
    const body = plan.groups
      .map((g) => [g.name, g.personaKey, g.scope, g.exists ? "exists" : "to create"].join(","))
      .join("\n");
    const blob = new Blob([`${head}\n${body}\n`], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `bulk-groups-${(seg?.code ?? "segment").toLowerCase()}.csv`;
    a.click();
  };

  /* ── Render ───────────────────────────────────────────────────────────────── */

  return (
    <div style={s.card}>
      <p style={s.head}>Create every group for a segment</p>
      <div style={s.okBox}>
        Creating groups grants <strong>nothing</strong>. This writes the groups and their{" "}
        <strong>Folder Access mappings</strong>; folder permissions are applied when you run{" "}
        <strong>Folder Reconciliation</strong>. Who is IN each group is set on <strong>Folder
        Access</strong>, per person, whenever you know — an empty group keeps its permissions.
      </div>

      <label style={s.label} htmlFor="bg-seg">Segment</label>
      <select
        id="bg-seg"
        style={s.select}
        value={chosen}
        disabled={busy}
        onChange={(e) => {
          setChosen(e.target.value);
          setLog([]);
          const t = (segments ?? []).filter((x) => x.key === e.target.value)[0];
          if (t) loadForSegment(t).catch(() => undefined);
        }}
      >
        <option value="">— select a segment —</option>
        {(segments ?? []).map((x) => <option key={x.key} value={x.key}>{x.label}</option>)}
      </select>
      {segments === undefined && (
        <p style={s.hint}>The segment list could not be read, so none are offered here.</p>
      )}

      {seg && (
        <>
          <label style={s.label}>Create for each unit, department and segment</label>
          <div style={s.ticks}>
            {OFFERED.map((k) => {
              const p = PERSONAS.filter((x) => x.key === k)[0];
              if (!p) return undefined;
              return (
                <label key={k} style={s.tick}>
                  <input
                    type="checkbox"
                    checked={picked.indexOf(k) !== -1}
                    disabled={busy}
                    onChange={(e) =>
                      setPicked(e.target.checked ? [...picked, k] : picked.filter((x) => x !== k))
                    }
                  />
                  <span>
                    {p.family} — {p.label} <span style={s.pill}>{p.scope}</span>
                  </span>
                </label>
              );
            })}
          </div>
          <p style={s.hint}>
            All ticked by default. A term with no folder code is left out and listed below — no code means no
            folder, so its groups would grant nothing.
          </p>
        </>
      )}

      {loading && <p style={s.hint}>Reading the term store and the codes…</p>}

      {plan && (
        <>
          <p style={{ ...s.head, marginTop: 16 }}>
            Preview — {plan.groups.length} group{plan.groups.length === 1 ? "" : "s"} ·{" "}
            {toCreateCount(plan)} to create · {plan.groups.length - toCreateCount(plan)} already there
          </p>
          <div style={s.scroll}>
            <table style={s.table}>
              <thead>
                <tr>
                  <th style={s.th}>Group name</th>
                  <th style={s.th}>Scope</th>
                  <th style={s.th}>Status</th>
                </tr>
              </thead>
              <tbody>
                {plan.groups.map((g) => (
                  <tr key={g.name}>
                    <td style={s.td}>{g.name}</td>
                    <td style={{ ...s.td, fontFamily: "inherit" }}>{g.scope}</td>
                    <td style={{ ...s.td, fontFamily: "inherit", color: g.exists ? "#605e5c" : "#0f6c3f" }}>
                      {g.exists ? "already exists — mapping only" : "new"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {plan.skipped.length > 0 && (
            <div style={s.warnBox}>
              <strong>{plan.skipped.length} term{plan.skipped.length === 1 ? "" : "s"} left out.</strong>{" "}
              Give them a code on <strong>CRS Term Abbreviations</strong>, then run this again.
              <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                {plan.skipped.slice(0, 12).map((x) => (
                  <li key={`${x.level}-${x.label}`}>
                    {x.label} ({x.level}) — {x.reason}
                  </li>
                ))}
              </ul>
              {plan.skipped.length > 12 && (
                <p style={{ margin: "6px 0 0" }}>…and {plan.skipped.length - 12} more.</p>
              )}
            </div>
          )}

          {existingRows === undefined && !loading && (
            <div style={s.warnBox}>
              <strong>{cachedListTitle(LIST_SUFFIX.groupMap)} could not be read, so this run is held.</strong>{" "}
              Without it there is no way to tell a mapping that already exists from one that is missing,
              and the run would write a second copy of every row it cannot see. Nothing on any screen
              shows a row twice, so that is a mistake you would not find. Re-pick the segment to try the
              read again.
            </div>
          )}

          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <button
              style={busy || plan.groups.length === 0 || !existingRows ? s.off : s.btn}
              disabled={busy || plan.groups.length === 0 || !existingRows}
              onClick={() => { run().catch(() => setBusy(false)); }}
            >
              {busy ? "Working…" : `Create ${toCreateCount(plan)} groups and map ${plan.groups.length}`}
            </button>
            <button style={s.ghost} disabled={busy} onClick={exportCsv}>Export CSV</button>
            {/* Leaving the page is the only other way out of a run, and holding the flow's navigation
                without offering a way to stop would trap an admin for the length of a 300-group run.
                Safe to offer only because the run is repeatable: it stops between groups and what
                landed stays. */}
            {busy && (
              <button
                style={s.ghost}
                disabled={stopping}
                onClick={() => { stopRef.current = true; setStopping(true); }}
              >
                {stopping ? "Stopping after this group…" : "Stop"}
              </button>
            )}
          </div>
          {existingRows !== undefined && (
            <p style={s.hint}>
              Safe to run twice: {existingRows.length} mapping row(s) are already on{" "}
              {cachedListTitle(LIST_SUFFIX.groupMap)} and will be left alone. Only what is missing is written.
            </p>
          )}
          {busy && (
            <p style={s.hint}>
              Leave this tab open and stay on this step — the run has no resume, and leaving stops it
              part-way. Nothing already written would be lost, and pressing Run again finishes the rest.
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
                l.indexOf("✗") !== -1 ? s.logBad : l.indexOf("+ ") === 0 ? s.logOk : s.logDim
              }
            >
              {l}
            </div>
          ))}
        </div>
      )}

      {toast && <Toast kind={toast.kind} text={toast.text} onDismiss={() => setToast(undefined)} />}
    </div>
  );
}
