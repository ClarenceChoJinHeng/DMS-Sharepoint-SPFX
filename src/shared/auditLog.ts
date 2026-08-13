/**
 * Audit log — the row shape, the summaries and the truncation rules.
 *
 * Spec: docs/superpowers/specs/2026-08-13-audit-log-design.md
 *
 * SPFx-free by design, the same split as naming.ts / spNaming.ts: everything that decides what a
 * row SAYS lives here and is unit-tested, while spAuditLog.ts owns the request. The wording of a
 * summary and the announcement of a truncation are the parts a reader relies on months later, so
 * they are pinned by tests rather than by review.
 *
 * The caller supplies the event time. This module never reads the clock — a module that calls
 * `new Date()` cannot have its date formatting tested.
 */

/**
 * Event types, as written to the `EventType` column.
 *
 * `EventType` is a TEXT column, not Choice, deliberately: writing a value absent from a Choice
 * column's `Choices` FAILS the whole write, so the day someone adds a type here every row of that
 * type would be lost silently. Text cannot fail that way, and it indexes identically.
 */
export const EVENT = {
  uploaded: "Uploaded",
  approved: "Approved",
  rejected: "Rejected",
  routed: "Routed",
  deleted: "Deleted",
  uploadRefused: "UploadRefused",
  accessGranted: "AccessGranted",
  accessRevoked: "AccessRevoked",
  reconciliationRun: "ReconciliationRun",
  structureChanged: "StructureChanged",
  migrationRun: "MigrationRun",
  segmentCreated: "SegmentCreated",
  abbreviationChanged: "AbbreviationChanged",
  policyChanged: "PolicyChanged",
  groupMapChanged: "GroupMapChanged",
} as const;

export type AuditEventType = (typeof EVENT)[keyof typeof EVENT];

/** Human labels, for the summary line and the viewer's filter list. */
export const EVENT_LABEL: Record<string, string> = {
  [EVENT.uploaded]: "Uploaded",
  [EVENT.approved]: "Approved",
  [EVENT.rejected]: "Rejected",
  [EVENT.routed]: "Moved to Documents",
  [EVENT.deleted]: "Deleted",
  [EVENT.uploadRefused]: "Upload refused",
  [EVENT.accessGranted]: "Access granted",
  [EVENT.accessRevoked]: "Access revoked",
  [EVENT.reconciliationRun]: "Reconciliation run",
  [EVENT.structureChanged]: "Folder structure changed",
  [EVENT.migrationRun]: "Folder migration run",
  [EVENT.segmentCreated]: "Segment created",
  [EVENT.abbreviationChanged]: "Abbreviation changed",
  [EVENT.policyChanged]: "File type policy changed",
  [EVENT.groupMapChanged]: "Group Map changed",
};

/** Every type, in the order the viewer offers them. */
export const ALL_EVENT_TYPES: string[] = [
  EVENT.uploaded, EVENT.approved, EVENT.rejected, EVENT.routed, EVENT.deleted,
  EVENT.uploadRefused, EVENT.accessGranted, EVENT.accessRevoked,
  EVENT.reconciliationRun, EVENT.structureChanged, EVENT.migrationRun,
  EVENT.segmentCreated, EVENT.abbreviationChanged, EVENT.policyChanged,
  EVENT.groupMapChanged,
];

/**
 * Outcome of the AUDITED ACTION — never of the audit write itself.
 *
 * A row that was never written has no outcome to record; that is the honest limit of any
 * self-hosted log and is stated on the page, not smuggled in as a status value here.
 */
export type AuditOutcome = "Success" | "Refused" | "Failed";

/**
 * What a caller describes.
 *
 * Nearly everything is optional because real call sites genuinely lack it: a policy change has no
 * file, and a Power Automate deletion trigger hands over almost no metadata. An absent value is
 * recorded as blank, which the viewer reads as "not recorded" — never as a different event.
 */
export interface AuditEvent {
  event: string;
  /** Defaults to "Success". */
  outcome?: AuditOutcome;
  /** Which writer produced the row — `UploadForm`, `FolderAccess`, `Flow:ApprovalActivity`. */
  source: string;
  /** When it happened. Supplied by the caller so this module stays clock-free. */
  at: Date;
  actorName?: string;
  actorEmail?: string;
  library?: string;
  itemUniqueId?: string;
  itemName?: string;
  itemPath?: string;
  segment?: string;
  unitPath?: string;
  /** Readable lines. Joined, and truncated with an announcement if oversized. */
  details?: string[];
  /**
   * Overrides the derived summary.
   *
   * Run-style events (reconciliation, migration) need counts in the summary and only the caller
   * knows them; deriving "Reconciliation run" alone would make every run look identical in the feed.
   */
  summary?: string;
}

/** The row as POSTed. Field names are the list's internal names, exactly. */
export interface AuditRow {
  Title: string;
  EventTime: string;
  EventType: string;
  Outcome: string;
  ActorName: string;
  ActorEmail: string;
  Source: string;
  LibraryName: string;
  ItemUniqueId: string;
  ItemName: string;
  ItemPath: string;
  Segment: string;
  UnitPath: string;
  Details: string;
}

/** SharePoint's Text limit. `Title` is a Text column, so a long file name must not fail the write. */
export const TITLE_MAX = 255;

/**
 * Budget for `Details`.
 *
 * A `Note` column holds far more than this, but a reconciliation over twelve segments can emit
 * thousands of lines and there is no value in a row nobody can read. Well under any server limit,
 * so the write never fails for size — a lost row is worse than a shortened one.
 */
export const DETAILS_MAX = 30000;

function clean(v: string | undefined): string {
  return (v ?? "").toString().trim();
}

/**
 * `M/D/YYYY h:mm tt` — the format the site locale accepts on write (gotcha #1).
 *
 * Built by hand rather than with `toLocaleString`, whose output depends on the BROWSER's locale: an
 * admin with a UK locale would produce `13/08/2026`, which SharePoint either rejects or — far worse
 * — reads as day 8 of month 13 and silently shifts the event. Display formatting is a separate
 * concern and belongs in the viewer.
 */
export function formatEventTime(d: Date): string {
  const h24 = d.getHours();
  const suffix = h24 < 12 ? "AM" : "PM";
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  const mins = d.getMinutes();
  const mm = mins < 10 ? `0${mins}` : `${mins}`;
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${h12}:${mm} ${suffix}`;
}

/** Cut to a length, marking that it was cut. A silently clipped value reads as complete. */
function cap(text: string, max: number): string {
  if (text.length <= max) return text;
  return `${text.slice(0, Math.max(0, max - 1))}…`;
}

/**
 * Join detail lines within the budget, ALWAYS announcing what was dropped.
 *
 * The count of dropped lines is what makes a shortened block trustworthy: a reader can tell the
 * difference between "that is all that happened" and "there was more". A quietly clipped block is
 * indistinguishable from a complete one, which is the failure this whole feature exists to prevent.
 */
export function joinDetails(lines: readonly string[] | undefined, max: number = DETAILS_MAX): string {
  const all = (lines ?? []).map((l) => clean(l)).filter((l) => l.length > 0);
  const kept: string[] = [];
  let size = 0;
  // Room reserved for the marker, so the announcement itself can never be the thing that overflows.
  const budget = Math.max(0, max - 60);
  for (let i = 0; i < all.length; i++) {
    const line = all[i];
    if (size + line.length + 1 > budget) {
      const dropped = all.length - kept.length;
      kept.push(`… ${dropped} more line${dropped === 1 ? "" : "s"} not recorded`);
      return kept.join("\n");
    }
    kept.push(line);
    size += line.length + 1;
  }
  return kept.join("\n");
}

/**
 * The one-line summary for the `Title` column.
 *
 * A caller-supplied summary always wins. Otherwise: the event's label, plus the item name when
 * there is one. An event with neither is still readable — "File type policy changed" says enough on
 * its own, and `Details` carries the rest.
 */
export function summarize(e: AuditEvent): string {
  const supplied = clean(e.summary);
  if (supplied.length > 0) return cap(supplied, TITLE_MAX);

  const label = EVENT_LABEL[e.event] ?? clean(e.event);
  const name = clean(e.itemName);
  const base = name.length > 0 ? `${label} — ${name}` : label;
  // "Upload refused" already says it was refused; repeating it reads as a different, worse event.
  const restate = e.outcome === "Refused" && e.event !== EVENT.uploadRefused;
  return cap(restate ? `${base} (refused)` : base, TITLE_MAX);
}

/**
 * Build the row.
 *
 * Every field is emitted, always, as a string — never omitted when absent. A partial body makes one
 * missing value look like a different kind of event later, whereas an empty string reads correctly
 * in the viewer as "not recorded". Nothing here can throw: a logger that throws while describing an
 * action that already succeeded turns a logging gap into a user-visible failure.
 */
export function buildAuditRow(e: AuditEvent): AuditRow {
  return {
    Title: summarize(e),
    EventTime: formatEventTime(e.at),
    EventType: clean(e.event),
    Outcome: e.outcome ?? "Success",
    ActorName: cap(clean(e.actorName), TITLE_MAX),
    ActorEmail: cap(clean(e.actorEmail), TITLE_MAX),
    Source: cap(clean(e.source), TITLE_MAX),
    LibraryName: cap(clean(e.library), TITLE_MAX),
    ItemUniqueId: cap(clean(e.itemUniqueId), TITLE_MAX),
    ItemName: cap(clean(e.itemName), TITLE_MAX),
    // ItemPath is a Note column and so is NOT capped at 255 — that is the reason it is a Note.
    // A truncated path cannot be matched back to a folder, which defeats recording it at all.
    ItemPath: clean(e.itemPath),
    Segment: cap(clean(e.segment), TITLE_MAX),
    UnitPath: cap(clean(e.unitPath), TITLE_MAX),
    Details: joinDetails(e.details),
  };
}

/**
 * The leading segment of a unit path, for callers that hold the path but not the segment.
 *
 * Convenience only — a caller that knows the segment should pass it. Returns "" rather than
 * guessing when the path is blank, and tolerates a leading slash.
 */
export function segmentFromUnitPath(unitPath: string | undefined): string {
  const parts = clean(unitPath).split("/").filter((p) => p.length > 0);
  return parts.length > 0 ? parts[0] : "";
}
