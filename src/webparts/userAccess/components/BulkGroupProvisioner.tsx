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
import { useEffect, useState } from "react";
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
  toCreateCount,
} from "../../../shared/bulkGroups";
import { buildGroupMapRow, GroupMapWriteRow, PERSONAS } from "../../../shared/groupMapModel";
import { createSiteGroup } from "../../../shared/spGroups";
import { cachedListTitle, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { Toast, ToastKind } from "../../../shared/toast";

type Props = { context: WebPartContext; siteUrl: string };

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
  logBox: { marginTop: 12, maxHeight: 240, overflowY: "auto", background: "#1b1b1b", color: "#d7d7d7", fontFamily: "Consolas, monospace", fontSize: 11.5, padding: "10px 12px", borderRadius: 4, lineHeight: 1.55 },
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

export default function BulkGroupProvisioner({ context, siteUrl }: Props): React.ReactElement {
  const [segments, setSegments] = useState<Seg[] | undefined>(undefined);
  const [chosen, setChosen] = useState("");
  const [rows, setRows] = useState<AbbrevRowDraft[] | undefined>(undefined);
  /** Existing site groups by lower-cased title, so one already there is MAPPED rather than re-created. */
  const [titles, setTitles] = useState<Record<string, number>>({});
  const [picked, setPicked] = useState<string[]>(OFFERED);
  const [loading, setLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<string[]>([]);
  const [toast, setToast] = useState<{ kind: ToastKind; text: string } | undefined>(undefined);

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

  const loadForSegment = async (target: Seg): Promise<void> => {
    setLoading(true);
    setRows(undefined);
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
    } catch (e) {
      setRows(undefined);
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
    if (!plan || !seg) return;
    setBusy(true);
    // A LOCAL log, mirrored into state. Counters read from state during a run would see their render-time
    // value and report zero — the same stale-closure trap as reconciliation's counts.
    const lines: string[] = [];
    const say = (t: string): void => { lines.push(t); setLog([...lines]); };
    let made = 0;
    let mapped = 0;
    let failed = 0;

    say(`${seg.label}: ${plan.groups.length} planned, ${toCreateCount(plan)} to create.`);
    for (const g of plan.groups) {
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
      for (const row of rowsFor(g, String(id))) {
        try {
          await withRetry(() => postRow(row));
          mapped++;
        } catch (e) {
          failed++;
          say(`  ✗ ${g.name} → ${row.Role} — ${(e as Error).message}`);
        }
      }
    }
    say(`Done. ${made} created, ${mapped} mappings written, ${failed} failed.`);
    setToast({
      kind: failed > 0 ? "warn" : "ok",
      text:
        `${made} groups created and ${mapped} mappings written for ${seg.label}.` +
        (failed > 0 ? ` ${failed} failed — see the log.` : "") +
        " Nothing is granted until Folder Reconciliation runs.",
    });
    setBusy(false);
    // Re-read, so a second press sees what the first made instead of trying to create it again.
    await loadForSegment(seg).catch(() => undefined);
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
        <strong>Folder Reconciliation</strong>. Members are added per person, on the form above.
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

          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <button
              style={busy || plan.groups.length === 0 ? s.off : s.btn}
              disabled={busy || plan.groups.length === 0}
              onClick={() => { run().catch(() => setBusy(false)); }}
            >
              {busy ? "Working…" : `Create ${toCreateCount(plan)} groups and map ${plan.groups.length}`}
            </button>
            <button style={s.ghost} disabled={busy} onClick={exportCsv}>Export CSV</button>
          </div>
          {busy && (
            <p style={s.hint}>
              Leave this tab open — the run has no resume, and closing it stops it part-way.
            </p>
          )}
        </>
      )}

      {log.length > 0 && (
        <div style={s.logBox}>
          {log.map((l, i) => <div key={i}>{l}</div>)}
        </div>
      )}

      {toast && <Toast kind={toast.kind} text={toast.text} onDismiss={() => setToast(undefined)} />}
    </div>
  );
}
