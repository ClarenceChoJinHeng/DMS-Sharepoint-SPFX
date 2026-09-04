import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";

import { ALL_EVENT_TYPES, EVENT_LABEL, eventLabelForRow } from "../../../shared/auditLog";
import { csvCell, downloadCsv } from "../../../shared/groupExportCsv";
import { primeNames } from "../../../shared/spNaming";
import { isSystemAdmin } from "../../../shared/spGroups";
import {
  AuditListState,
  auditListState,
  auditListTitle,
  AuditRecord,
  plannedAuditListTitle,
  provisionAuditList,
  ProvisionReport,
  readAudit,
  readAuditPage,
  AuditQuery,
} from "../../../shared/spAuditLog";
import { IAuditLogProps } from "./IAuditLogProps";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

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
const s: Record<string, React.CSSProperties> = {
  /* The Upload Form's page shell (client, 2026-09-04: *"the same padding spacing that the upload
     form is using ... they do not want the pages to stick at the wall"*). `margin: 32px auto` plus
     `padding: 0 24px 48px` is the pattern the Upload Form, My Submissions and the access pages
     already share; this page had the centring and NO padding, so its content met the window edge.
     The max-width is left alone deliberately - each page's is sized for its own content. */
  wrap: { maxWidth: 1180, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif", color: "#1b1b1b" },
  h2: { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle: { fontSize: 13, color: "#605e5c", margin: "0 0 18px", lineHeight: 1.5 },
  card: { border: "1px solid #e1dfdd", borderRadius: 8, padding: "16px 18px", marginBottom: 16 },
  cardTitle: {
    fontSize: 12, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase",
    color: "#0f6c3f", margin: "0 0 10px",
  },
  msg: { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.55 },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { ...NOTICE_ATTENTION },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  info: { background: "#f3f2f1", border: "1px solid #e1dfdd", color: "#323130" },
  /* THE FILTER PANEL (client design, 2026-08-30). Fields on the left in a two-column grid, the three
     buttons stacked on the right — which is what keeps Apply beside the fields it applies to instead
     of below a wrapping row where it reads as unrelated. */
  pager: { paddingTop: 14, display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", justifyContent: "space-between" },
  pagerNote: { fontSize: 11.5, color: "#605e5c" },
  pagerBtns: { display: "flex", gap: 4, alignItems: "center" },
  pageBtn: { minWidth: 30, height: 30, padding: "0 8px", border: "1px solid #e1dfdd", borderRadius: 6, background: "#fff", color: "#242424", fontSize: 12.5, cursor: "pointer" },
  pageBtnOn: { borderColor: "#0f6c3f", color: "#0f6c3f", fontWeight: 700 },
  filterGrid: { display: "flex", flexWrap: "wrap", gap: 20, alignItems: "flex-start" },
  filterFields: { flex: "1 1 520px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 },
  filterActions: { flex: "0 0 auto", display: "flex", flexDirection: "column", gap: 8, minWidth: 130 },
  actionsNote: { fontSize: 10.5, color: "#8a8886", textAlign: "center" },
  fieldPair: { display: "flex", flexDirection: "column", gap: 4, gridColumn: "1 / -1" },
  pairRow: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" },
  pairTag: { fontSize: 11.5, color: "#605e5c", flexShrink: 0 },
  /* The intro card. Icon left, heading and one line of explanation right — the same shape the CRS
     Settings cards use, so the two pages read as one system. */
  intro: { display: "flex", gap: 14, alignItems: "flex-start", marginBottom: 18 },
  introIcon: { flexShrink: 0, width: 54, height: 54, borderRadius: 10, background: "#D5EBD2", display: "flex", alignItems: "center", justifyContent: "center", color: "#00684A" },
  introHead: { fontSize: 14, fontWeight: 700, margin: "2px 0 4px", color: "#242424" },
  introBody: { fontSize: 12.5, color: "#605e5c", margin: 0, lineHeight: 1.55 },
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
  headWho: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  headRefresh: {
    border: "1px solid #c7c7c7", background: "#fff", borderRadius: 4, cursor: "pointer",
    fontSize: 13, lineHeight: 1, padding: "2px 7px", color: "#0f6c3f",
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

/* ⚠ `startOfDaysAgo` and `RANGES` were removed on 2026-08-30 with the "Last 7 days" period select.
   The dates are typed now, and both bounds go through one `bound()` helper inside `load` — which is
   also where the To bound is pushed to the END of its day, so picking a date includes it. */

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
  /* PAGING (client design, 2026-08-30). The old control was "Load more", which APPENDED — fine for
     reading forwards, useless for "go back to the second page".

     ⚠ THE MOCK'S `Showing 1 to 5 of 235 events` WITH A JUMP TO PAGE 47 CANNOT BE BUILT, and the two
     reasons are both about SharePoint rather than effort:
       • A TOTAL needs a count, and this list's `ItemCount` is a CACHED aggregate that lags in both
         directions — it read 2,419 for a list whose view was empty on 2026-08-24. A wrong total on
         screen is worse than none.
       • REST paging is a CONTINUATION TOKEN, so page 47 is only reachable by walking 1–46.
     So: numbered buttons for pages already REACHED, plus Next. Same shape, nothing invented.

     `tokens[i]` is the token that FETCHES page i. `tokens[0]` is undefined — page one needs none. */
  const [tokens, setTokens] = useState<Array<string | undefined>>([undefined]);
  const [pageIdx, setPageIdx] = useState(0);
  const [readFailed, setReadFailed] = useState(false);
  const [readStatus, setReadStatus] = useState<number | undefined>(undefined);
  const [loading, setLoading] = useState(false);

  /* ⚠ BOTH BOUNDS OPTIONAL AND INDEPENDENT. Blank From means "from the beginning" — the query that
     is NOT index-bounded, which used to be the deliberate "All time" choice and is now simply what
     an empty box means. Left empty on arrival so the first view is everything recent rather than a
     window the admin did not choose. */
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
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

  /**
   * The filters as a query.
   *
   * ⚠ ONE DEFINITION, because `goToPage` re-runs page one and MUST ask exactly what `load` asked.
   * Two copies of this would drift, and the symptom would be paging that quietly changes the result
   * set as you move through it — which reads as the log being wrong rather than the query being two
   * different queries.
   */
  const currentQuery = (focusId?: string): AuditQuery => {
    /* ⚠ AN UNPARSEABLE DATE IS NO BOUND, never today's. A half-typed `2026-08-` in a date input
       yields an Invalid Date, and passing that into the query would filter on NaN and return
       nothing at all — which reads as "there are no events" rather than "that date is incomplete". */
    const bound = (v: string): Date | undefined => {
      const d = v ? new Date(v) : undefined;
      return d && !isNaN(d.getTime()) ? d : undefined;
    };
    const to = bound(dateTo);
    return {
      from: bound(dateFrom),
      /* The To box means the whole of that DAY. Without this, picking the 30th excludes everything
         that happened on the 30th, because the value parses to midnight at its start. */
      to: to ? new Date(to.getTime() + 24 * 60 * 60 * 1000 - 1) : undefined,
      eventTypes: types,
      actorEmail: actor.trim(),
      text: text.trim(),
      itemUniqueId: focusId ?? focus?.id,
      top: PAGE_SIZE,
    };
  };

  const load = async (focusId?: string): Promise<void> => {
    setLoading(true);
    const page = await readAudit(sp, siteUrl, currentQuery(focusId));
    setRows(page.rows);
    setNext(page.next);
    /* ⚠ A NEW QUERY INVALIDATES EVERY TOKEN. A continuation token belongs to the query that
       produced it, so reusing one after the filters changed would page through the OLD result set
       while the screen showed the new filters — rows that match nothing the admin asked for, with
       nothing to explain them. */
    setTokens([undefined]);
    setPageIdx(0);
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

      /* Site Collection Admin OR a member of the site OWNERS group - see `isSystemAdmin`.
         WARN: THIS CHECKED `IsSiteAdmin` ALONE UNTIL 2026-08-27, which is why adding somebody to
         `CRS Owners` did not let them use this page: a group can never confer `IsSiteAdmin`, so the
         gate could only ever be satisfied one person at a time in Site Settings. Still fails closed -
         it gates a CONTROL, and the page's own permissions keep non-admins out. */
      setAdmin(await isSystemAdmin(sp, siteUrl));

      const state = await auditListState(sp, siteUrl);
      setListState(state);
      setBooting(false);
      if (state === "ready") await load(undefined);
    };
    init().catch(() => setBooting(false));
  }, []);

  /**
   * Show one page.
   *
   * REPLACES the rows rather than appending, which is the whole difference from the "Load more" this
   * grew out of: a page is a position, not a growing list.
   *
   * ⚠ ONLY A PAGE ALREADY REACHED, OR THE NEXT ONE. Anything further has no token, so there is
   * nothing to fetch it with — the buttons for those are not rendered, and this refuses as well, so
   * the rule holds even if a caller is added later.
   */
  const goToPage = async (idx: number): Promise<void> => {
    if (idx < 0 || idx > tokens.length) return;
    const token = idx === 0 ? undefined : tokens[idx];
    if (idx > 0 && token === undefined) return;
    setLoading(true);
    const page = token === undefined
      ? await readAudit(sp, siteUrl, currentQuery())
      : await readAuditPage(sp, token);
    setRows(page.rows);
    setPageIdx(idx);
    setNext(page.next);
    /* Remember how to fetch the page AFTER this one, so Next works from wherever we land — including
       after jumping backwards, where `next` would otherwise still describe the page we left. */
    setTokens((prev) => {
      const out = prev.slice(0, idx + 1);
      if (page.next !== undefined) out[idx + 1] = page.next;
      return out;
    });
    setExpanded({});
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

  /* `toggleType` went with the chip row — the Action select writes `types` directly. `types` stays an
     ARRAY so `readAudit` is unchanged and multi-select is a control swap away. */

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
      <h2 style={s.h2}>CRS Audit Log</h2>
      {/* ⚠ THE "cannot be edited or deleted" SENTENCE IS NOT DECORATION and is kept below. It is the
          claim that makes this log worth reading at all — writes are restricted to Owners and the
          service account by design, and an admin who does not know that has no reason to trust a row
          in front of them. */}
      <div style={s.intro}>
        <span style={s.introIcon} aria-hidden="true">
          <svg width="26" height="26" viewBox="0 0 24 24" fill="none">
            <circle cx="9" cy="8" r="3.2" stroke="currentColor" strokeWidth="1.6" />
            <path d="M3 19c0-3.2 2.7-5.3 6-5.3 1.4 0 2.7.4 3.7 1.1" stroke="currentColor"
              strokeWidth="1.6" strokeLinecap="round" />
            <circle cx="17" cy="15" r="3.2" stroke="currentColor" strokeWidth="1.6" />
            <path d="M13.5 21c.5-1.6 1.9-2.6 3.5-2.6s3 1 3.5 2.6" stroke="currentColor"
              strokeWidth="1.6" strokeLinecap="round" />
          </svg>
        </span>
        <div>
          <p style={s.introHead}>What can you do with the Audit Log?</p>
          <p style={s.introBody}>
            Track activities and changes across the system, including document actions, access
            changes, and folder updates. Records are written automatically, and cannot be edited or
            deleted from this page.
          </p>
        </div>
      </div>
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
              {/* plannedAuditListTitle, NOT auditListTitle: while the list is absent the latter has
                  fallen back to the legacy `DMS …` name, so this panel offered to create a list that
                  the button would not create. Both now come from one derivation. */}
              This creates the <strong>{plannedAuditListTitle()}</strong> list, its columns and its
              indexes, and turns on version history. It changes nothing else, and is safe to run again.
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
        <p style={s.cardTitle}>Filter</p>
        <div style={s.filterGrid}>
          <div style={s.filterFields}>
            {/* ⚠ EXPLICIT DATES REPLACED THE "Last 7 days / 30 / 90 / All time" PERIOD SELECT
                (client design, 2026-08-30). `AuditQuery` already carried a `to` bound, so this is a
                wiring change rather than a new capability.

                What was traded: "Last 7 days" was ONE click and is now two dates. What was gained:
                an investigation into a specific week no longer has to widen to 30 days and read past
                everything else. Both bounds are OPTIONAL and independently so — a From with no To
                means "since then", which is the old behaviour with a date the admin chose. */}
            <div style={s.fieldPair}>
              <span style={s.label}>Date</span>
              <div style={s.pairRow}>
                <span style={s.pairTag}>From</span>
                <input
                  type="date"
                  style={s.input}
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                />
                <span style={s.pairTag}>To</span>
                <input
                  type="date"
                  style={s.input}
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                />
              </div>
            </div>

            <div style={s.field}>
              <span style={s.label}>Action</span>
              {/* ⚠ ONE ACTION, WHERE THE CHIPS ALLOWED SEVERAL. The chip row let an admin pick
                  Uploaded AND Approved together; this cannot. Followed the mock because the client's
                  standing instruction is fewer inputs (*"don't give client too many features or
                  inputs, it will make them scared"*) — but `types` is still an ARRAY all the way into
                  `readAudit`, so restoring multi-select is a control swap, not a query change. */}
              <select
                style={s.select}
                value={types[0] ?? ""}
                onChange={(e) => setTypes(e.target.value ? [e.target.value] : [])}
              >
                <option value="">Select action</option>
                {ALL_EVENT_TYPES.map((t) => (
                  <option key={t} value={t}>{EVENT_LABEL[t] ?? t}</option>
                ))}
              </select>
            </div>

            <div style={s.field}>
              <span style={s.label}>Person (email address)</span>
              <input
                style={s.input}
                value={actor}
                placeholder="someone@abc.com"
                onChange={(e) => setActor(e.target.value)}
              />
            </div>

            <div style={s.field}>
              <span style={s.label}>Keyword</span>
              <input
                style={s.input}
                value={text}
                placeholder="Tax return"
                onChange={(e) => setText(e.target.value)}
              />
            </div>
          </div>

          <div style={s.filterActions}>
            <button style={s.btn} disabled={loading} onClick={() => load(focus?.id).catch(() => undefined)}>
              {loading ? "Loading…" : "Apply"}
            </button>
            {/* ⚠ RESET CLEARS AND RE-READS. Clearing the boxes without running the query leaves the
                admin looking at the OLD result under empty filters, which is the state that reads as
                "the page is broken". */}
            <button
              style={s.ghost}
              disabled={loading}
              onClick={() => {
                setDateFrom("");
                setDateTo("");
                setTypes([]);
                setActor("");
                setText("");
                load(focus?.id).catch(() => undefined);
              }}
            >
              ↻ Reset
            </button>
            <button
              style={s.ghost}
              disabled={rows.length === 0}
              onClick={() => downloadCsv(toAuditCsv(rows), csvName(new Date()))}
            >
              ↑ Export
            </button>
            <span style={s.actionsNote}>export what is shown</span>
          </div>
        </div>
        <p style={{ fontSize: 11, color: "#605e5c", margin: "10px 0 0" }}>
          Press Apply to use the filters.
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
          {/* A SECOND refresh, on the results header (client, 2026-08-24: *"can you add a refesh
              button like beside the column"*). The one in the filter card is 400–600px above this
              on a full page of rows, so the person watching for a flow to write a row has to scroll
              back up to ask for it again — and scrolling up to a filter panel is what makes people
              reload the whole page instead.

              Right-aligned INSIDE the last grid column, so the four headings stay aligned with the
              four columns of every row below. Icon-only with a `title`: a second "Refresh" wide
              enough to read would push "Who" out of alignment, and the glyph is unambiguous next to
              a table. Runs the identical `load` call, focus included. */}
          <div style={s.head}>
            <span>When</span><span>Event</span><span>What</span>
            <span style={s.headWho}>
              Who
              <button
                type="button"
                style={s.headRefresh}
                disabled={loading}
                title="Re-read the log with the same filters"
                aria-label="Refresh"
                onClick={() => load(focus?.id).catch(() => undefined)}
              >
                {loading ? "…" : "↻"}
              </button>
            </span>
          </div>
          {rows.map((r) => {
            const open = expanded[r.Id] === true;
            return (
              <div key={r.Id} style={s.row}>
                <span style={s.when}>{formatWhen(r.EventTime)}</span>
                <span style={s.type}>
                  {eventLabelForRow(r.EventType, r.LibraryName)}
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
                {/* ⚠ `Source` IS NO LONGER SHOWN UNDER THE NAME (client, 2026-08-30). It is still
                    STORED on every row and still EXPORTED in the CSV — only the column is gone.
                    Worth knowing what was given up: `Source` is what distinguishes a row a PERSON
                    wrote (`CrsConfiguration`, `Requests`) from one a FLOW wrote
                    (`Flow:DocumentsDeletions`), which is the difference between "somebody did this"
                    and "the system did this". If that question ever comes up in an investigation,
                    the CSV still answers it. */}
                <div style={s.who}>
                  <div>{r.ActorName || r.ActorEmail || "—"}</div>
                </div>
              </div>
            );
          })}
          {/* PAGINATION (client design, 2026-08-30). Numbered, but only over pages that have a
              token — see the note on `tokens`. `…` is shown when more exist beyond them, which is
              honest about there being more without claiming to know how many. */}
          <div style={s.pager}>
            <span style={s.pagerNote}>
              Showing {rows.length === 0 ? 0 : pageIdx * PAGE_SIZE + 1} to{" "}
              {pageIdx * PAGE_SIZE + rows.length} event{rows.length === 1 ? "" : "s"}
              {/* ⚠ NO TOTAL. The list's `ItemCount` is a cached aggregate that lags in BOTH
                  directions — it read 2,419 for a list whose view was empty (2026-08-24) — so a
                  number here would be wrong often enough to mislead. "and more" is what we can
                  actually stand behind. */}
              {next !== undefined ? " and more" : ""}
            </span>
            <div style={s.pagerBtns}>
              <button
                style={s.pageBtn}
                disabled={loading || pageIdx === 0}
                title="Previous page"
                onClick={() => goToPage(pageIdx - 1).catch(() => undefined)}
              >
                &lsaquo;
              </button>
              {tokens.map((_, i) => (
                <button
                  key={i}
                  style={i === pageIdx ? { ...s.pageBtn, ...s.pageBtnOn } : s.pageBtn}
                  disabled={loading}
                  onClick={() => goToPage(i).catch(() => undefined)}
                >
                  {i + 1}
                </button>
              ))}
              {next !== undefined && <span style={s.pagerNote}>&hellip;</span>}
              <button
                style={s.pageBtn}
                disabled={loading || next === undefined}
                title="Next page"
                onClick={() => goToPage(pageIdx + 1).catch(() => undefined)}
              >
                &rsaquo;
              </button>
            </div>
          </div>
        </div>
      )}

      {limitsPanel}
    </div>
  );
};

export default AuditLog;
