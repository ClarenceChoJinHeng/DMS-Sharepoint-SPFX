import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { AuditEvent, AuditRow, buildAuditRow } from "./auditLog";
import { cachedListTitle, LIST_SUFFIX } from "./naming";
import { ensureColumn } from "./spColumns";
import { writePrefix } from "./spNaming";

/**
 * Audit log — provisioning, the write, and the viewer's read.
 *
 * Spec: docs/superpowers/specs/2026-08-13-audit-log-design.md
 *
 * The SPFx half of the pair; auditLog.ts holds everything testable. Split the same way as
 * naming.ts / spNaming.ts.
 */

/* -- Schema ---------------------------------------------------------------- */

interface ColumnDef {
  name: string;
  display: string;
  kind: "Text" | "Note" | "DateTime";
  /** Indexed columns are the ones the viewer filters and sorts on. */
  indexed?: boolean;
}

/**
 * The 13 columns beyond the built-in `Title`.
 *
 * Internal names carry no spaces, and each is created UNDER its internal name and then retitled by
 * `ensureColumn` — an internal name is frozen at creation, permanently. `Vendor_x002f_CustomerName`
 * is what happens when this is left to chance.
 *
 * Only four are indexed. An index costs write throughput, and these are exactly the four the viewer
 * puts in a `$filter` or an `$orderby`; indexing the rest would slow every write for nothing.
 */
export const AUDIT_COLUMNS: ColumnDef[] = [
  { name: "EventTime", display: "Event time", kind: "DateTime", indexed: true },
  { name: "EventType", display: "Event type", kind: "Text", indexed: true },
  { name: "Outcome", display: "Outcome", kind: "Text" },
  { name: "ActorName", display: "Actor", kind: "Text" },
  { name: "ActorEmail", display: "Actor email", kind: "Text", indexed: true },
  { name: "Source", display: "Source", kind: "Text" },
  { name: "LibraryName", display: "Library", kind: "Text" },
  { name: "ItemUniqueId", display: "Item unique id", kind: "Text", indexed: true },
  { name: "ItemName", display: "Item", kind: "Text" },
  { name: "ItemPath", display: "Path", kind: "Note" },
  { name: "Segment", display: "Segment", kind: "Text" },
  { name: "UnitPath", display: "Unit path", kind: "Text" },
  { name: "Details", display: "Details", kind: "Note" },
];

/** Everything the viewer reads back, including the built-ins it needs. */
const READ_SELECT = [
  "Id", "Title", "EventTime", "EventType", "Outcome", "ActorName", "ActorEmail", "Source",
  "LibraryName", "ItemUniqueId", "ItemName", "ItemPath", "Segment", "UnitPath", "Details", "Created",
].join(",");

/* -- Helpers --------------------------------------------------------------- */

/** The live title, resolved through the shared cache. Callers must have primed names first. */
export function auditListTitle(): string {
  return cachedListTitle(LIST_SUFFIX.auditLog);
}

function listBase(siteUrl: string, title: string): string {
  return `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')`;
}

const JSON_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json",
};

/**
 * The list's item entity type, read from the list itself and cached for the session.
 *
 * Never derived from the title: RECREATING a list changes `ListItemEntityTypeFullName` while
 * renaming does not, so a derived value stays right until the day someone rebuilds the list, and
 * then every write breaks (gotcha #12). Same approach as FolderMap.tsx.
 */
let entityLookup: Promise<string | undefined> | undefined;

async function itemEntityType(sp: SPHttpClient, siteUrl: string): Promise<string | undefined> {
  // The in-flight PROMISE is cached, not the resolved value. Caching the value means a
  // check-then-await-then-assign, which is both a lint error and a real race: several writers can
  // fire at once on an admin screen, and each would issue its own lookup. This way the first caller
  // starts one request and every other awaits it.
  //
  // A failed lookup is cleared by the CALLER (writeAudit), not here: clearing it after the await
  // would be a post-await assignment to a module variable, the very thing the promise cache exists
  // to avoid. Either way a transient failure must not pin "unwritable" for the page's whole life.
  if (entityLookup === undefined) {
    entityLookup = (async () => {
      try {
        const res: SPHttpClientResponse = await sp.get(
          `${listBase(siteUrl, auditListTitle())}?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!res.ok) return undefined;
        const data = await res.json();
        const t = data?.ListItemEntityTypeFullName;
        return typeof t === "string" && t.length > 0 ? t : undefined;
      } catch {
        // Absent list, or a transient failure. The caller decides; this is not the place to guess.
        return undefined;
      }
    })();
  }
  return entityLookup;
}

/** Test seam, and the escape hatch after the list is recreated mid-session. */
export function clearAuditCache(): void {
  entityLookup = undefined;
}

/* -- State ----------------------------------------------------------------- */

/**
 * Whether the log is usable.
 *
 * `unknown` is distinct from `absent` on purpose, and the viewer renders them differently. A
 * transient read failure reported as "absent" would invite an admin to provision a list that
 * already exists, and — worse on a page whose whole job is completeness — would present a
 * permissions problem as "nothing has happened yet".
 */
export type AuditListState = "ready" | "absent" | "unknown";

export async function auditListState(sp: SPHttpClient, siteUrl: string): Promise<AuditListState> {
  try {
    const res: SPHttpClientResponse = await sp.get(
      `${listBase(siteUrl, auditListTitle())}?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (res.ok) return "ready";
    if (res.status === 404) return "absent";
    return "unknown";
  } catch {
    return "unknown";
  }
}

/* -- Provisioning ---------------------------------------------------------- */

export interface ProvisionReport {
  createdList: boolean;
  createdColumns: string[];
  indexed: string[];
  versioningEnabled: boolean;
  /** Non-fatal problems, named so an admin can act. An empty array means a clean run. */
  problems: string[];
}

/**
 * Create the list, its columns and its indexes, and turn version history on.
 *
 * Idempotent and safe to re-run: every step checks first, so a half-finished run is repaired by
 * running it again. The client cannot run PowerShell, so anything left to a script would simply not
 * get done — which is why the page provisions itself.
 *
 * Breaking inheritance is NOT done here. It needs to name the service account, which this code
 * cannot know; a guess would either grant the wrong identity or lock the flows out of the list they
 * are about to write to. The page states it as a manual step instead.
 */
export async function provisionAuditList(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<ProvisionReport> {
  const report: ProvisionReport = {
    createdList: false,
    createdColumns: [],
    indexed: [],
    versioningEnabled: false,
    problems: [],
  };

  const state = await auditListState(sp, siteUrl);
  if (state === "unknown") {
    // Creating a list that might already exist is how you end up with two logs, one of which
    // quietly receives nothing. Refuse rather than risk it.
    report.problems.push(
      "Could not tell whether the audit log list exists — the site did not answer. Nothing was changed; try again.",
    );
    return report;
  }

  // The title to CREATE comes from the prefix in force on this site, not from the resolved title:
  // while the list is absent, `auditListTitle()` has fallen back to the legacy `DMS …` name, and
  // creating that on a CRS site would leave behind a list nobody expects.
  let title = auditListTitle();
  if (state === "absent") {
    const prefix = await writePrefix(sp, siteUrl).catch(() => "");
    if (prefix.length > 0) title = `${prefix} ${LIST_SUFFIX.auditLog}`;

    const create: SPHttpClientResponse = await sp.post(
      `${siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: { ...JSON_HEADERS, Accept: "application/json;odata=verbose" },
        body: JSON.stringify({
          __metadata: { type: "SP.List" },
          Title: title,
          BaseTemplate: 100,
          Description:
            "Audit trail for the CRS document management system. Written automatically — do not edit rows by hand.",
        }),
      },
    );
    if (!create.ok) {
      const body = await create.text().catch(() => "");
      report.problems.push(
        `Could not create the "${title}" list (HTTP ${create.status}). ${body.slice(0, 200)}`,
      );
      return report;
    }
    report.createdList = true;
    clearAuditCache();
  }

  for (const col of AUDIT_COLUMNS) {
    try {
      const made = await ensureColumn(sp, siteUrl, title, col.name, col.display, col.kind);
      if (made) report.createdColumns.push(col.name);
    } catch (e) {
      // One failed column must not abandon the rest: a log missing `Segment` is still worth having,
      // and the next run repairs it. Naming it is what makes that possible.
      report.problems.push(`Column "${col.name}": ${(e as Error).message}`);
    }
  }

  const indexable = AUDIT_COLUMNS.filter((c) => c.indexed === true);
  for (const col of indexable) {
    const ok = await setIndexed(sp, siteUrl, title, col.name);
    if (ok) {
      report.indexed.push(col.name);
    } else {
      report.problems.push(
        `Could not index "${col.name}". The log still works, but filtering slows down as it grows.`,
      );
    }
  }

  const versioned = await enableVersioning(sp, siteUrl, title);
  report.versioningEnabled = versioned;
  if (!versioned) {
    report.problems.push(
      "Could not turn on version history. Without it an edited row leaves no trace — set it in list settings.",
    );
  }

  return report;
}

/** Index one column. A failure is degradation, not breakage, so it never throws. */
async function setIndexed(
  sp: SPHttpClient,
  siteUrl: string,
  title: string,
  internalName: string,
): Promise<boolean> {
  try {
    const res: SPHttpClientResponse = await sp.post(
      `${listBase(siteUrl, title)}/fields/getbyinternalnameortitle('${encodeURIComponent(internalName)}')`,
      SPHttpClient.configurations.v1,
      {
        headers: { ...JSON_HEADERS, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        body: JSON.stringify({ Indexed: true }),
      },
    );
    return res.ok;
  } catch {
    return false;
  }
}

async function enableVersioning(sp: SPHttpClient, siteUrl: string, title: string): Promise<boolean> {
  try {
    const res: SPHttpClientResponse = await sp.post(
      listBase(siteUrl, title),
      SPHttpClient.configurations.v1,
      {
        headers: { ...JSON_HEADERS, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
        body: JSON.stringify({ EnableVersioning: true }),
      },
    );
    return res.ok;
  } catch {
    return false;
  }
}

/* -- Writing --------------------------------------------------------------- */

/**
 * Write one row. NEVER throws, and never rejects.
 *
 * Returns true when the row was written. The contract matters: the caller has usually already
 * performed the action being logged, and a logger that throws at that point converts a logging gap
 * into a user-visible failure of something that actually succeeded.
 *
 * A `false` return is not nothing — an admin screen shows a non-blocking warning, because an audit
 * gap somebody knows about is worth far more than one nobody does. The console line carries the
 * status and body for diagnosis.
 */
export async function writeAudit(
  sp: SPHttpClient,
  siteUrl: string,
  event: AuditEvent,
): Promise<boolean> {
  let title = "";
  try {
    const row: AuditRow = buildAuditRow(event);
    title = row.Title;
    const type = await itemEntityType(sp, siteUrl);
    if (type === undefined) {
      // The list is absent, or the read failed. Drop the cached lookup so the NEXT write retries:
      // a transient failure must not disable logging for the rest of the page's life.
      clearAuditCache();
      console.warn(`[audit] not written — the audit log list could not be reached: ${title}`);
      return false;
    }

    const res: SPHttpClientResponse = await sp.post(
      `${listBase(siteUrl, auditListTitle())}/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { ...JSON_HEADERS, Accept: "application/json;odata=verbose" },
        body: JSON.stringify({ __metadata: { type }, ...row }),
      },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      console.warn(`[audit] not written (HTTP ${res.status}): ${title}. ${body.slice(0, 300)}`);
      return false;
    }
    return true;
  } catch (e) {
    console.warn(`[audit] not written: ${title} — ${(e as Error).message}`);
    return false;
  }
}

/* -- Reading --------------------------------------------------------------- */

export interface AuditQuery {
  /** Inclusive lower bound on EventTime. */
  from?: Date;
  /** Exclusive upper bound on EventTime. */
  to?: Date;
  eventTypes?: string[];
  actorEmail?: string;
  /** Substring match on the item name and the summary. */
  text?: string;
  /** One file's whole history, across path changes. */
  itemUniqueId?: string;
  top?: number;
}

export interface AuditRecord extends AuditRow {
  Id: number;
  Created: string;
}

export interface AuditPage {
  rows: AuditRecord[];
  /** Absolute URL for the next page, or undefined at the end. */
  next?: string;
  /** True when the list itself could not be read — distinct from an empty result. */
  failed: boolean;
  status?: number;
}

/**
 * ISO, for `$filter`.
 *
 * Deliberately NOT the write format. A row is written `M/D/YYYY h:mm tt` because the site locale
 * demands it (gotcha #1), while `$filter` wants an OData `datetime'…'` literal in ISO. Using the
 * write format here returns an empty result rather than an error, which reads as "nothing
 * happened" — the one thing this page must never say by accident.
 */
function odataDate(d: Date): string {
  return `datetime'${d.toISOString()}'`;
}

/** OData string literals escape a single quote by doubling it. */
function lit(v: string): string {
  return `'${v.replace(/'/g, "''")}'`;
}

/**
 * Build the `$filter`.
 *
 * Every clause targets an INDEXED column except the free-text one, which is why the viewer always
 * keeps a date window: an indexed clause keeps the query cheap however large the list gets.
 */
export function buildAuditFilter(q: AuditQuery): string {
  const clauses: string[] = [];
  if (q.from) clauses.push(`EventTime ge ${odataDate(q.from)}`);
  if (q.to) clauses.push(`EventTime lt ${odataDate(q.to)}`);
  if (q.itemUniqueId !== undefined && q.itemUniqueId.length > 0) {
    clauses.push(`ItemUniqueId eq ${lit(q.itemUniqueId)}`);
  }
  if (q.actorEmail !== undefined && q.actorEmail.length > 0) {
    clauses.push(`ActorEmail eq ${lit(q.actorEmail)}`);
  }
  const types = (q.eventTypes ?? []).filter((t) => t.length > 0);
  if (types.length > 0) {
    clauses.push(`(${types.map((t) => `EventType eq ${lit(t)}`).join(" or ")})`);
  }
  const text = (q.text ?? "").trim();
  if (text.length > 0) {
    clauses.push(`(substringof(${lit(text)},ItemName) or substringof(${lit(text)},Title))`);
  }
  return clauses.join(" and ");
}

/**
 * Read a page of the log, newest first.
 *
 * `failed` is separate from an empty `rows` array, and the viewer must keep them apart: "no events
 * match" and "the list could not be read" look identical on screen otherwise, and on an audit page
 * that difference is the whole point.
 */
export async function readAudit(
  sp: SPHttpClient,
  siteUrl: string,
  q: AuditQuery,
): Promise<AuditPage> {
  const filter = buildAuditFilter(q);
  const url =
    `${listBase(siteUrl, auditListTitle())}/items` +
    `?$select=${READ_SELECT}` +
    `&$orderby=EventTime desc,Id desc` +
    `&$top=${q.top ?? 100}` +
    (filter.length > 0 ? `&$filter=${encodeURIComponent(filter)}` : "");
  return fetchPage(sp, url);
}

/** Follow a `next` link from a previous page. */
export async function readAuditPage(sp: SPHttpClient, url: string): Promise<AuditPage> {
  return fetchPage(sp, url);
}

async function fetchPage(sp: SPHttpClient, url: string): Promise<AuditPage> {
  try {
    const res: SPHttpClientResponse = await sp.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata" },
    });
    if (!res.ok) return { rows: [], failed: true, status: res.status };
    const data = await res.json();
    const nextLink = data["odata.nextLink"];
    return {
      rows: (data.value ?? []) as AuditRecord[],
      next: typeof nextLink === "string" ? nextLink : undefined,
      failed: false,
    };
  } catch {
    return { rows: [], failed: true };
  }
}
