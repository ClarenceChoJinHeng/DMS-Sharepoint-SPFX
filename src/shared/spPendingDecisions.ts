import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  PendingDecision,
  PENDING_DECISION_COLUMNS,
  buildDecisionPayload,
  parseDecisionRow,
} from "./pendingDecisions";
import {
  cachedListTitle,
  titleForNewList,
  noteCreatedList,
  LIST_SUFFIX,
} from "./naming";

/**
 * `CRS Pending Decisions` — the SPFx half of the pair; `pendingDecisions.ts` holds everything
 * testable. Same split as `naming.ts` / `spNaming.ts`, `auditLog.ts` / `spAuditLog.ts` and
 * `submissionRecords.ts` / `spSubmissionRecords.ts` — importing `@microsoft/sp-http` here would make
 * the pure rules untestable, exactly as it would for `PRIMED_SUFFIXES`.
 *
 * Spec: docs/superpowers/specs/2026-09-18-tag-approve-proxy-design.md
 */

/**
 * ⚠ ALL THREE HEADERS, AND THE THIRD IS THE LOAD-BEARING ONE.
 *
 * JSON light needs no `__metadata` envelope and no entity type, so `ListItemEntityTypeFullName` is
 * deliberately NOT read for this list — do not "restore" it per gotcha #12, which applies to the
 * verbose call sites in `FolderMap.tsx`.
 *
 * SPFx's `SPHttpClient` injects `odata-version: 4.0` on every request, and SharePoint cannot infer
 * the entity set for a JSON-light ENTRY payload under OData 4 — so a row POST returns 400 while
 * `/_api/web/lists` tolerates the identical headers and creates a list quite happily. That
 * asymmetry is what made this look like three unrelated bugs across 2026-08-13, and it cost the
 * file-type page again on 2026-08-27 (1.0.249.0). An empty value removes the header.
 */
const WRITE_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
  "odata-version": "",
};

// ⚠ NO-CACHE, same reason as `dmsFolderMap.ts`'s folder-permission probes and the
// `spSubmissionRecords.ts` fix: a stale response here would misjudge whether a decision this
// session just wrote has actually landed, or would re-serve a Pending row after the flow already
// applied it.
const GET_HEADERS = {
  Accept: "application/json;odata=nometadata",
  "Cache-Control": "no-cache",
  Pragma: "no-cache",
};

/** Read/write base, for a list assumed to already exist. Never used for creating one — see
 * `provisionPendingDecisionsList`, which resolves the title a different way. */
const listBase = (siteUrl: string): string =>
  `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.pendingDecisions))}')`;

/**
 * The columns the list already has, lower-cased. Empty when unreadable — which makes the caller
 * ATTEMPT every column even though some may already exist. That attempt does not "fail harmlessly"
 * on an existing column: SharePoint accepts a duplicate display name outright and silently renames
 * it with a numeric suffix instead of rejecting it — see the warning below for why that makes this
 * check load-bearing, not a nicety. Reading first turns "already exists" from the normal case (a
 * second provisioning run) into a real signal.
 *
 * ⚠ THIS IS NOT OPTIONAL. SharePoint does NOT reject a duplicate DISPLAY name outright — it silently
 * appends a numeric suffix (`Decision`, `Decision0`, `Decision1`, …) and returns success, so without
 * this check `provisionPendingDecisionsList` called a second time on an already-provisioned list
 * would create a fresh set of duplicate columns every time, reporting a clean run each time. This is
 * not hypothetical: `CRS Submissions` was hit by exactly this — 16 provisioning runs produced 14
 * duplicate columns each before it was caught, because the naive loop there had no existence check
 * at all. Mirrors `Requests.tsx`'s `readFieldNames` verbatim.
 */
const readFieldNames = async (
  sp: SPHttpClient,
  siteUrl: string,
  title: string,
): Promise<string[]> => {
  try {
    const res: SPHttpClientResponse = await sp.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/fields?$select=Title&$top=500`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!res.ok) return [];
    return ((await res.json()).value ?? []).map((f: { Title?: string }) =>
      (f.Title ?? "").toLowerCase(),
    );
  } catch {
    return [];
  }
};

/**
 * Ensure `CRS Pending Decisions` exists, with every column it needs.
 *
 * REPEATABLE by construction, the same lesson the requests list paid for on 2026-08-20: create the
 * list only when it is confirmed absent, then attempt EVERY MISSING column even if one fails, and
 * report the failures together rather than aborting on the first one. A run that half-finishes must
 * still leave a route forward — throwing on the first failed column strands the list with some
 * columns present and no way to add the rest except a second CREATE, which 500s because the list
 * already exists. "Missing" is checked via `readFieldNames` FIRST, before any column is attempted —
 * see that function's comment for why skipping this step silently duplicates columns rather than
 * failing loudly.
 */
export async function provisionPendingDecisionsList(
  sp: SPHttpClient,
  siteUrl: string,
): Promise<{ added: string[]; failed: string[] }> {
  // titleForNewList, NOT cachedListTitle. A list that does not exist yet can never be in the name
  // cache, so cachedListTitle would answer the LEGACY `DMS Pending Decisions` — which on a
  // CRS-renamed site creates the one list nobody can find. See `resolvedPrefix` in naming.ts.
  const title = titleForNewList(LIST_SUFFIX.pendingDecisions);
  const base = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')`;

  const probe: SPHttpClientResponse = await sp.get(
    `${base}?$select=Id`,
    SPHttpClient.configurations.v1,
    { headers: GET_HEADERS },
  );
  if (probe.status === 404) {
    const made: SPHttpClientResponse = await sp.post(
      `${siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: WRITE_HEADERS,
        body: JSON.stringify({
          Title: title,
          BaseTemplate: 100,
          Description:
            "Pending approve/reject decisions, applied by CRS — Apply pending decisions " +
            "(Power Automate) rather than by the approver's own click.",
        }),
      },
    );
    if (!made.ok) {
      throw new Error(`creating the list failed — HTTP ${made.status}`);
    }
  } else if (!probe.ok) {
    // Neither present nor absent. Creating on top of that is how a duplicate list appears.
    throw new Error(`could not check whether "${title}" exists — HTTP ${probe.status}`);
  }
  // Recorded before the columns, so a run that half-finishes on a column failure still leaves every
  // later read pointing at the list that now exists, rather than at the legacy cached name.
  noteCreatedList(LIST_SUFFIX.pendingDecisions, title);

  const added: string[] = [];
  const failed: string[] = [];
  const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/fields`;
  const have = await readFieldNames(sp, siteUrl, title);
  for (const c of PENDING_DECISION_COLUMNS) {
    // Skip a column already present — see `readFieldNames`'s own comment for why this check is
    // load-bearing rather than a nicety: SharePoint accepts a duplicate DISPLAY name and silently
    // renames it with a numeric suffix, so re-provisioning without this would create a fresh set of
    // duplicates on every run, reporting success each time.
    if (have.indexOf(c.name.toLowerCase()) > -1) continue;
    // `FieldTypeKind` IS the literal value already carried on `c.type` (2 Text / 3 Note /
    // 4 DateTime / 8 Boolean) — no NumberOfLines, no other property. `SP.FieldMultiLineText`-only
    // properties are rejected outright under `odata=nometadata`, because SharePoint infers the
    // entity type from the endpoint rather than from a supplied `__metadata` type — the exact
    // failure that stopped `CRS Requests`' provisioning cold on 2026-08-20.
    const r: SPHttpClientResponse = await sp.post(fieldsUrl, SPHttpClient.configurations.v1, {
      headers: WRITE_HEADERS,
      body: JSON.stringify({ Title: c.name, FieldTypeKind: c.type }),
    });
    if (r.ok) added.push(c.name);
    else failed.push(`${c.name} (HTTP ${r.status})`);
  }
  return { added, failed };
}

/**
 * Write one pending decision. **Never throws.**
 *
 * ⚠ NOT THE SAME DEGRADE AS `writeSubmissionRecord`, AND THAT DIFFERENCE IS DELIBERATE.
 * `writeSubmissionRecord` downgrades a 404 to `console.info`, because that row is a supplementary
 * record — the upload has already succeeded either way, and a missing row costs nothing but a future
 * label on My Submissions. This row IS the action: with direct Edit/Approve removed from the
 * approver's own permissions, this write is the ONLY record that an approve or reject ever happened
 * — nothing else performs it, and `CRS — Apply pending decisions` has nothing to replay if it never
 * lands. So every failure here is logged with `console.error`, uniformly, never downgraded — do NOT
 * "align" this with `writeSubmissionRecord`'s quieter logging; that would hide the one failure this
 * whole feature exists to surface. It still never THROWS: the caller decides how to tell the
 * approver their click did not take, this function only reports what happened.
 */
export async function writePendingDecision(
  sp: SPHttpClient,
  siteUrl: string,
  decision: Omit<PendingDecision, "id" | "status" | "statusError">,
): Promise<{ ok: true } | { ok: false; status: number; body: string }> {
  try {
    const res: SPHttpClientResponse = await sp.post(
      `${listBase(siteUrl)}/items`,
      SPHttpClient.configurations.v1,
      { headers: WRITE_HEADERS, body: JSON.stringify(buildDecisionPayload(decision)) },
    );
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      console.error(
        `[pendingDecisions] not written (HTTP ${res.status}): ${decision.itemUniqueId}. ${body.slice(0, 300)}`,
      );
      return { ok: false, status: res.status, body };
    }
    return { ok: true };
  } catch (e) {
    const message = (e as Error).message;
    console.error(`[pendingDecisions] not written: ${decision.itemUniqueId} — ${message}`);
    return { ok: false, status: 0, body: message };
  }
}

/**
 * Read the most recent decision row for a document, by its own `ItemUniqueId` — never by name or
 * path, both of which can change. `$orderby=Id desc&$top=1` because a document can in principle
 * gain more than one row over its life (a reject, then later an approve); the newest is the one
 * that describes its current state.
 *
 * Returns `undefined` on any failure or on an empty result — the caller must not treat that as "no
 * decision was ever made", only as "we could not confirm one right now".
 *
 * Silent on failure, deliberately, unlike `writePendingDecision`'s uniform `console.error`: this is
 * a READ meant to be POLLED — a screen re-checking whether the flow has applied a decision yet — so
 * a transient miss is the expected, repeatable case rather than a one-off event worth a log line.
 * Logging every poll tick would bury the one write failure that actually matters in noise.
 */
export async function readPendingDecision(
  sp: SPHttpClient,
  siteUrl: string,
  itemUniqueId: string,
): Promise<PendingDecision | undefined> {
  const id = (itemUniqueId ?? "").trim();
  if (id.length === 0) return undefined;
  try {
    const res: SPHttpClientResponse = await sp.get(
      `${listBase(siteUrl)}/items?$filter=ItemUniqueId eq '${encodeURIComponent(id.replace(/'/g, "''"))}'&$orderby=Id desc&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS },
    );
    if (!res.ok) return undefined;
    const rows = ((await res.json()).value ?? []) as Array<Record<string, unknown>>;
    if (rows.length === 0) return undefined;
    return parseDecisionRow(rows[0]);
  } catch {
    return undefined;
  }
}
