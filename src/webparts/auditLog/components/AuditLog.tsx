import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";

import { ALL_EVENT_TYPES, EVENT_LABEL } from "../../../shared/auditLog";
import { csvCell, downloadCsv } from "../../../shared/groupExportCsv";
import { primeNames } from "../../../shared/spNaming";
import {
  AuditListState,
  auditListState,
  auditListTitle,
  AuditRecord,
  provisionAuditList,
  ProvisionReport,
  readAudit,
  readAuditPage,
} from "../../../shared/spAuditLog";
import { IAuditLogProps } from "./IAuditLogProps";

/**
 * CRS Audit Log — the viewer.
 *
 * Spec: docs/superpowers/specs/2026-08-13-audit-log-design.md
 *
 * Reads only. There is deliberately no UI to edit or delete a row: a log the audited party can amend
 * is not a log, and the absence of the control states that better than a disabled button would.
 *
 * Admin-only by decision D1, because every row carries a file name and a full path. The page's own
 * permissions are the real boundary — this check is what stops a non-admin who reaches the page from
 * seeing a confusing half-page, and it gates the provisioning control.
 */

const PAGE_SIZE = 100;

/** Preset windows. Every query keeps a date bound, so the filter stays on an indexed column. */
const RANGES: Array<{ key: string; label: string; days: number }> = [
  { key: "7", label: "Last 7 days", days: 7 },
  { key: "30", label: "Last 30 days", days: 30 },
  { key: "90", label: "Last 90 days", days: 90 },
  { key: "all", label: "All time", days: 0 },
];

const s: Record<string, React.CSSProperties> = {
  wrap: { maxWidth: 1180, margin: "0 auto", fontFamily: "'Segoe UI', sans-serif", color: "#1b1b1b" },
  h2: { fontSize: 20, fontWeight: 600, margin: "0 0 4px" },
  subtitle: { fontSize: 13, color: "#605e5c", margin: "0 0 18px", lineHeight: 1.5 },
  card: { border: "1px solid #e1dfdd", borderRadius: 8, padding: "16px 18px", marginBottom: 16 },
  cardTitle: {
    fontSize: 12, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase",
    color: "#0f6c3f", margin: "0 0 10px",
  },
  msg: { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.55 },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  info: { background: "#f3f2f1", border: "1px solid #e1dfdd", color: "#323130" },
  filters: { display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end", marginBottom: 14 },
  field: { display: "flex", flexDirection: "column", gap: 4 },
  label: { fontSize: 11, fontWeight: 600, color: "#605e5c" },
  input: {
    padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4,
    boxSizing: "border-box", minWidth: 180,
  },
  select: { padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4 },
  btn: {
    background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "8px 16px",
    fontSize: 13, cursor: "pointer",
  },
  ghost: {
    background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4,
    padding: "7px 14px", fontSize: 13, cursor: "pointer",
  },
  chipRow: { display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12 },
  chip: {
    fontSize: 11, borderRadius: 12, padding: "4px 10px", cursor: "pointer",
    border: "1px solid #c8c8c8", background: "#fff", color: "#323130",
  },
  chipOn: { border: "1px solid #0f6c3f", background: "#0f6c3f", color: "#fff" },
  head: {
    display: "grid", gridTemplateColumns: "150px 150px minmax(0,1fr) 160px", columnGap: 12,
    padding: "0 10px 8px", fontSize: 12, fontWeight: 600, color: "#605e5c",
    borderBottom: "1px solid #edebe9",
  },
  row: {
    display: "grid", gridTemplateColumns: "150px 150px minmax(0,1fr) 160px", columnGap: 12,
    padding: "10px", borderBottom: "1px solid #f3f2f1", alignItems: "start",
  },
  when: { fontSize: 12, color: "#323130", fontFamily: "Consolas, monospace" },
  type: { fontSize: 12, fontWeight: 600, color: "#0f6c3f" },
  title: { fontSize: 13, color: "#1b1b1b", wordBreak: "break-word" },
  path: { fontSize: 11, color: "#605e5c", wordBreak: "break-all", marginTop: 3 },
  who: { fontSize: 12, color: "#323130", wordBreak: "break-word" },
  details: {
    fontSize: 11, color: "#323130", background: "#faf9f8", border: "1px solid #edebe9",
    borderRadius: 4, padding: "8px 10px", marginTop: 6, whiteSpace: "pre-wrap",
    fontFamily: "Consolas, monospace",
  },
  link: {
    background: "none", border: "none", padding: 0, color: "#0f6c3f", fontSize: 11,
    cursor: "pointer", textDecoration: "underline",
  },
  disclose: {
    background: "none", border: "none", padding: 0, color: "#605e5c", fontSize: 11,
    cursor: "pointer", textAlign: "left",
  },
};

/**
 * `13/Aug/2026 09:41` — the agreed display format, with the time the feed needs.
 *
 * Parsing here is safe because `odata=nometadata` returns a DateTime field as ISO. This is NOT the
 * trap in gotcha #1: that one is re-parsing a value SharePoint has already formatted for display.
 */
function formatWhen(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${p(d.getDate())}/${months[d.getMonth()]}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function startOfDaysAgo(days: number): Date {
  const d = new Date();
  d.setDate(d.getDate() - days);
  d.setHours(0, 0, 0, 0);
  return d;
}

const CSV_HEADERS = [
  "Event time", "Event type", "Outcome", "Actor", "Actor email", "Source", "Library",
  "Item", "Path", "Segment", "Unit path", "Summary", "Details",
];

/** Reuses csvCell, which handles quoting AND Excel formula injection — folder names are client text. */
function toAuditCsv(rows: AuditRecord[]): string {
  const lines = [CSV_HEADERS.map(csvCell).join(",")];
  for (const r of rows) {
    lines.push([
      formatWhen(r.EventTime), r.EventType, r.Outcome, r.ActorName, r.ActorEmail, r.Source,
      r.LibraryName, r.ItemName, r.ItemPath, r.Segment, r.UnitPath, r.Title, r.Details,
    ].map((v) => csvCell(v ?? "")).join(","));
  }
  return lines.join("\r\n");
}

function csvName(now: Date): string {
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `CRS-audit-log-${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())}.csv`;
}

const AuditLog: React.FC<IAuditLogProps> = ({ context, siteUrl }) => {
  const sp: SPHttpClient = context.spHttpClient;

  const [booting, setBooting] = useState(true);
  const [admin, setAdmin] = useState(false);
  const [listState, setListState] = useState<AuditListState>("unknown");

  const [rows, setRows] = useState<AuditRecord[]>([]);
  const [next, setNext] = useState<string | undefined>(undefined);
  const [readFailed, setReadFailed] = useState(false);
  const [readStatus, setReadStatus] = useState<number | undefined>(undefined);
  const [loading, setLoading] = useState(false);

  const [range, setRange] = useState("7");
  const [types, setTypes] = useState<string[]>([]);
  const [actor, setActor] = useState("");
  const [text, setText] = useState("");
  /** Set when the user pivots to one file's history. */
  const [focus, setFocus] = useState<{ id: string; name: string } | undefined>(undefined);

  const [expanded, setExpanded] = useState<Record<number, boolean>>({});
  const [showLimits, setShowLimits] = useState(false);
  const [provisioning, setProvisioning] = useState(false);
  const [report, setReport] = useState<ProvisionReport | undefined>(undefined);

  /* -- Load ---------------------------------------------------------------- */

  const load = async (focusId?: string): Promise<void> => {
    setLoading(true);
    const days = RANGES.filter((r) => r.key === range)[0]?.days ?? 7;
    const page = await readAudit(sp, siteUrl, {
      // "All time" drops the date bound deliberately, and is the one query that is not
      // index-bounded. Offered because an investigation needs it; it is not the default.
      from: days > 0 ? startOfDaysAgo(days) : undefined,
      eventTypes: types,
      actorEmail: actor.trim(),
      text: text.trim(),
      itemUniqueId: focusId,
      top: PAGE_SIZE,
    });
    setRows(page.rows);
    setNext(page.next);
    setReadFailed(page.failed);
    setReadStatus(page.status);
    setExpanded({});
    setLoading(false);
  };

  useEffect(() => {
    const init = async (): Promise<void> => {
      // NAMES FIRST. An unprimed cache resolves to the legacy `DMS …` titles, which 404 on a renamed
      // site — the exact failure that took the abbreviation page down on 2026-08-12.
      await primeNames(sp, siteUrl).catch(() => undefined);

      let isAdmin = false;
      try {
        const res = await sp.get(
          `${siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (res.ok) {
          const me = await res.json();
          isAdmin = me.IsSiteAdmin === true;
        }
      } catch {
        // Treated as non-admin: this gates a CONTROL, so failing closed is right here. The page's
        // own permissions are what actually keep non-admins out.
      }
      setAdmin(isAdmin);

      const state = await auditListState(sp, siteUrl);
      setListState(state);
      setBooting(false);
      if (state === "ready") await load(undefined);
    };
    init().catch(() => setBooting(false));
  }, []);

  const more = async (): Promise<void> => {
    if (next === undefined) return;
    setLoading(true);
    const page = await readAuditPage(sp, next);
    // Append. A "load more" that replaced the rows would silently lose everything above it.
    setRows((prev) => prev.concat(page.rows));
    setNext(page.next);
    setLoading(false);
  };

  const provision = async (): Promise<void> => {
    setProvisioning(true);
    const r = await provisionAuditList(sp, siteUrl);
    setReport(r);
    const state = await auditListState(sp, siteUrl);
    setListState(state);
    setProvisioning(false);
    if (state === "ready") await load(undefined);
  };

  /** Pivot to one file's history. Applies immediately — a filter nobody applied is a filter nobody
   * trusts, and this one is a single click from a row they are already reading. */
  const focusOn = (id: string, name: string): void => {
    setFocus({ id, name });
    load(id).catch(() => undefined);
  };

  const clearFocus = (): void => {
    setFocus(undefined);
    load(undefined).catch(() => undefined);
  };

  const toggleType = (t: string): void => {
    setTypes((prev) => (prev.indexOf(t) === -1 ? prev.concat([t]) : prev.filter((x) => x !== t)));
  };

  /* -- Render -------------------------------------------------------------- */

  if (booting) {
    return (
      <div style={s.wrap}>
        <p style={s.subtitle}>Loading the audit log…</p>
      </div>
    );
  }

  const header = (
    <>
      <h2 style={s.h2}>Audit log</h2>
      <p style={s.subtitle}>
        What happened in the document management system, and who did it. Records are written
        automatically, and cannot be edited or deleted from this page.
      </p>
    </>
  );

  const limitsPanel = (
    <div style={s.card}>
      <button style={s.disclose} onClick={() => setShowLimits(!showLimits)}>
        {showLimits ? "▾" : "▸"} What this log cannot tell you
      </button>
      {showLimits && (
        <div style={{ fontSize: 12, color: "#323130", lineHeight: 1.6, marginTop: 10 }}>
          <p style={{ margin: "0 0 8px" }}>
            A log that looks complete but is not would be worse than none, so these gaps are stated
            rather than left to be discovered:
          </p>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            <li>
              <strong>Who viewed or downloaded a file.</strong> Nothing tells this system that a file
              was read. That record exists only in Microsoft Purview, which needs a tenant-wide
              administrator consent this site-scoped solution cannot request. Ask your Microsoft 365
              administrator for a Purview audit report.
            </li>
            <li><strong>Failed access attempts.</strong> A denied request never reaches this system.</li>
            <li>
              <strong>Anything from before this log was switched on.</strong> It starts from that day
              and cannot be filled in backwards.
            </li>
            <li>
              <strong>Permission changes made in SharePoint&apos;s own “Manage access” dialog</strong>{" "}
              instead of on the Folder Access page. Use the Folder Access page and they are recorded.
            </li>
            <li>
              <strong>Term store edits.</strong> Adding, renaming or deleting a term is not reported to
              this system. The consequence is recorded instead — the folder rename, or the orphaned
              term that reconciliation found.
            </li>
            <li>
              <strong>Rows changed directly in the list.</strong> The list&apos;s version history is
              the only trace, which is why it is switched on.
            </li>
          </ul>
        </div>
      )}
    </div>
  );

  // ---- The list is not there yet. Distinct from "unreadable" below.
  if (listState === "absent") {
    return (
      <div style={s.wrap}>
        {header}
        <div style={{ ...s.msg, ...s.info }}>
          <strong>The audit log is not set up yet.</strong>{" "}
          {admin
            ? "Nothing has been recorded so far. Setting it up creates the list that events are written to."
            : "Ask an administrator to open this page once — it sets itself up."}
        </div>
        {admin && (
          <div style={s.card}>
            <p style={s.cardTitle}>Set up</p>
            <p style={{ fontSize: 13, color: "#323130", lineHeight: 1.6, margin: "0 0 12px" }}>
              This creates the <strong>{auditListTitle()}</strong> list, its columns and its indexes,
              and turns on version history. It changes nothing else, and is safe to run again.
            </p>
            <p style={{ ...s.msg, ...s.warn, marginBottom: 12 }}>
              <strong>One step is left to you afterwards.</strong> In the new list&apos;s permission
              settings, stop inheriting permissions and give <em>Contribute</em> to the service account
              the Power Automate flows run as — keeping Full Control for owners, and removing everyone
              else. This is not automated on purpose: the code does not know which account that is, and
              a wrong guess would lock the flows out of the list they write to.
            </p>
            <button style={s.btn} disabled={provisioning} onClick={() => provision().catch(() => undefined)}>
              {provisioning ? "Setting up…" : "Set up the audit log"}
            </button>
          </div>
        )}
        {report && report.problems.length > 0 && (
          <div style={{ ...s.msg, ...s.err }}>
            {report.problems.map((p, i) => <div key={i}>{p}</div>)}
          </div>
        )}
        {limitsPanel}
      </div>
    );
  }

  // ---- We could not tell. Never presented as "nothing happened".
  if (listState === "unknown") {
    return (
      <div style={s.wrap}>
        {header}
        <div style={{ ...s.msg, ...s.err }}>
          <strong>The audit log could not be read.</strong> This is a permissions or connection
          problem, not an empty log — events may well have been recorded. Reload the page, and if it
          persists, check that you have access to the <strong>{auditListTitle()}</strong> list.
        </div>
        {limitsPanel}
      </div>
    );
  }

  if (!admin) {
    return (
      <div style={s.wrap}>
        {header}
        <div style={{ ...s.msg, ...s.info }}>
          This page is for administrators. Audit records name files and folder paths across every
          business segment, so they are not shown more widely.
        </div>
      </div>
    );
  }

  return (
    <div style={s.wrap}>
      {header}

      {report && (
        <div style={{ ...s.msg, ...(report.problems.length > 0 ? s.warn : s.ok) }}>
          {report.createdList && <div>Created the {auditListTitle()} list.</div>}
          {report.createdColumns.length > 0 && (
            <div>Added {report.createdColumns.length} columns.</div>
          )}
          {report.problems.map((p, i) => <div key={i}>{p}</div>)}
        </div>
      )}

      {focus && (
        <div style={{ ...s.msg, ...s.info }}>
          Showing the whole history of <strong>{focus.name || "this item"}</strong>, including any time
          it was moved or renamed.{" "}
          <button style={s.link} onClick={clearFocus}>Show everything again</button>
        </div>
      )}

      <div style={s.card}>
        <p style={s.cardTitle}>Filters</p>
        <div style={s.filters}>
          <div style={s.field}>
            <span style={s.label}>Period</span>
            <select style={s.select} value={range} onChange={(e) => setRange(e.target.value)}>
              {RANGES.map((r) => <option key={r.key} value={r.key}>{r.label}</option>)}
            </select>
          </div>
          <div style={s.field}>
            <span style={s.label}>Person (email)</span>
            <input
              style={s.input}
              value={actor}
              placeholder="someone@company.com"
              onChange={(e) => setActor(e.target.value)}
            />
          </div>
          <div style={s.field}>
            <span style={s.label}>File name or summary contains</span>
            <input
              style={s.input}
              value={text}
              placeholder="tax return"
              onChange={(e) => setText(e.target.value)}
            />
          </div>
          <button style={s.btn} disabled={loading} onClick={() => load(focus?.id).catch(() => undefined)}>
            {loading ? "Loading…" : "Apply"}
          </button>
          <button
            style={s.ghost}
            disabled={rows.length === 0}
            onClick={() => downloadCsv(toAuditCsv(rows), csvName(new Date()))}
          >
            Export what is shown
          </button>
        </div>

        <div style={s.chipRow}>
          {ALL_EVENT_TYPES.map((t) => {
            const on = types.indexOf(t) !== -1;
            return (
              <button
                key={t}
                style={on ? { ...s.chip, ...s.chipOn } : s.chip}
                onClick={() => toggleType(t)}
              >
                {EVENT_LABEL[t] ?? t}
              </button>
            );
          })}
        </div>
        <p style={{ fontSize: 11, color: "#605e5c", margin: 0 }}>
          No event types selected means all of them. Press Apply to use the filters.
        </p>
      </div>

      {readFailed && (
        <div style={{ ...s.msg, ...s.err }}>
          <strong>The log could not be read{readStatus ? ` (HTTP ${readStatus})` : ""}.</strong>{" "}
          This is not an empty result — do not read it as “nothing happened”.
        </div>
      )}

      {!readFailed && rows.length === 0 && (
        <div style={{ ...s.msg, ...s.info }}>
          <strong>No events match these filters.</strong> The log was read successfully and this
          filter simply found nothing — try a longer period, or clear the event types.
        </div>
      )}

      {rows.length > 0 && (
        <div style={s.card}>
          <div style={s.head}>
            <span>When</span><span>Event</span><span>What</span><span>Who</span>
          </div>
          {rows.map((r) => {
            const open = expanded[r.Id] === true;
            return (
              <div key={r.Id} style={s.row}>
                <span style={s.when}>{formatWhen(r.EventTime)}</span>
                <span style={s.type}>
                  {EVENT_LABEL[r.EventType] ?? r.EventType}
                  {r.Outcome && r.Outcome !== "Success" && (
                    <span style={{ color: "#a4262c" }}> · {r.Outcome}</span>
                  )}
                </span>
                <div>
                  <div style={s.title}>{r.Title}</div>
                  {r.ItemPath && <div style={s.path}>{r.ItemPath}</div>}
                  <div style={{ display: "flex", gap: 12, marginTop: 4, flexWrap: "wrap" }}>
                    {r.Details && (
                      <button
                        style={s.disclose}
                        onClick={() => setExpanded({ ...expanded, [r.Id]: !open })}
                      >
                        {open ? "▾ Hide details" : "▸ Details"}
                      </button>
                    )}
                    {r.ItemUniqueId && (
                      <button style={s.link} onClick={() => focusOn(r.ItemUniqueId, r.ItemName)}>
                        History of this file
                      </button>
                    )}
                  </div>
                  {open && r.Details && <div style={s.details}>{r.Details}</div>}
                </div>
                <div style={s.who}>
                  <div>{r.ActorName || r.ActorEmail || "—"}</div>
                  <div style={{ fontSize: 11, color: "#605e5c" }}>{r.Source}</div>
                </div>
              </div>
            );
          })}
          <div style={{ paddingTop: 12, display: "flex", gap: 10, alignItems: "center" }}>
            {next !== undefined && (
              <button style={s.ghost} disabled={loading} onClick={() => more().catch(() => undefined)}>
                {loading ? "Loading…" : "Load more"}
              </button>
            )}
            <span style={{ fontSize: 11, color: "#605e5c" }}>
              {rows.length} event{rows.length === 1 ? "" : "s"} shown
              {next !== undefined ? " — there are more" : ""}
            </span>
          </div>
        </div>
      )}

      {limitsPanel}
    </div>
  );
};

export default AuditLog;
