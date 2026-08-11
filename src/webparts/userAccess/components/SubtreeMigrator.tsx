import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels, PENDING_LEVELS_FIELD } from "../../../shared/formModel";
import { effectiveOnDemandTiers, splitChain, validateChain } from "../../../shared/folderChain";
import {
  backfillNeeds,
  Destination,
  EffectiveTier,
  effectiveTiers,
  FileRow,
  planTotals,
  planUnit,
  UnitPlan,
} from "../../../shared/subtreeMigration";
import { cachedListTitle, libraryTitle, libraryUrlSegment, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import {
  encodeServerRelativePath,
  ensureFolder,
  loadFolderMapRows,
  moveFolderTo,
} from "../../../shared/dmsFolderMap";

/**
 * Subtree Migration — move documents that are already filed into a new folder shape.
 *
 * Piece 3 of the configurable folder chain, spec
 * `docs/superpowers/specs/2026-08-11-subtree-migration-design.md`. Piece 2 lets an admin add a
 * level below Unit; that applies to new uploads only, so a segment already in use ends up with
 * two shapes side by side. This closes that.
 *
 * The rule that shapes the whole screen: it DERIVES what is misplaced from the term data and
 * shows the evidence, rather than asking "which level is new?". An admin can answer that
 * question wrongly, and a wrong answer files a unit's entire history under a value nobody
 * chose. That reasoning lives in `shared/subtreeMigration.ts` and is unit-tested; this file is
 * the REST and the rendering.
 */

const DOCUMENTS_LIST_TITLE = "Documents";
const DOCUMENTS_URL_SEGMENT = "Shared Documents";

const s: Record<string, React.CSSProperties> = {
  msg:      { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5 },
  err:      { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn:     { background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" },
  ok:       { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  card:     { border: "1px solid #e1e1e1", borderRadius: 8, padding: "14px 16px", marginBottom: 12, background: "#fff" },
  unitName: { fontSize: 14, fontWeight: 600, color: "#1b1b1b", fontFamily: "Consolas, monospace" },
  move:     { fontSize: 12, color: "#605e5c", fontFamily: "Consolas, monospace", wordBreak: "break-all", padding: "2px 0" },
  btn:      { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  ghost:    { background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  off:      { background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "not-allowed" },
  label:    { display: "block", fontSize: 12, fontWeight: 600, color: "#323130", margin: "10px 0 4px" },
  input:    { width: "100%", boxSizing: "border-box", padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4 },
  hint:     { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  log:      { fontSize: 12, fontFamily: "Consolas, monospace", maxHeight: 320, overflowY: "auto", border: "1px solid #e1e1e1", borderRadius: 6, padding: "8px 10px", background: "#fafafa" },
  logRow:   { padding: "2px 0", wordBreak: "break-all" },
  badge:    { fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 10, marginLeft: 8 },
  modalBg:  { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 },
  modal:    { background: "#fff", borderRadius: 8, padding: 24, maxWidth: 560, width: "100%", maxHeight: "80vh", overflowY: "auto" },
};

/** A DMS Config `mode` row, reduced to what this screen needs. */
interface SegmentRow {
  id: number;
  key: string;
  label: string;
  stagingFolder: string;
  /** The LIVE chain — what uploads are using right now. */
  chain: Level[];
  /**
   * A chain authored in the Folder levels tab but not yet applied. When present it is the
   * TARGET shape this screen migrates towards, and applying it is the last step of the run.
   */
  pending?: Level[];
  chainError?: string;
}

/** One library, resolved. Both hold the same tree and a unit can be adrift in one only. */
interface LibCtx {
  key: "Staging" | "Documents";
  title: string;
  urlSegment: string;
  /** Content approval on. Decides whether folders this tool creates must be stamped Approved. */
  moderated: boolean;
}

interface TermLite {
  id: string;
  label: string;
}

/** What the scan found for one unit in one library. */
interface UnitScan {
  lib: LibCtx;
  unitPath: string;
  /** `gho/gf/coru` — identity shared across libraries, so one choice serves both. */
  tail: string;
  children: string[];
  tiers: EffectiveTier[];
  /** Term options per EFFECTIVE tier, for the destination pickers. */
  options: TermLite[][];
  /** Set when this unit cannot be planned; it is reported and skipped. */
  unresolved?: string;
}

export interface SubtreeMigratorProps {
  context: WebPartContext;
  siteUrl: string;
}

export default function SubtreeMigrator({ context, siteUrl }: SubtreeMigratorProps): React.ReactElement {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [segments, setSegments] = useState<SegmentRow[]>([]);
  const [legacySets, setLegacySets] = useState<{ year: string; docType: string }>({ year: "", docType: "" });
  const [chosen, setChosen] = useState<string>("");
  const [scanning, setScanning] = useState(false);
  const [scanError, setScanError] = useState<string | undefined>(undefined);
  const [scans, setScans] = useState<UnitScan[] | undefined>(undefined);
  /** Destinations per unit tail, indexed by position in that unit's effective tier list. */
  const [dest, setDest] = useState<Record<string, Array<Destination | undefined>>>({});
  const [running, setRunning] = useState(false);
  const [log, setLog] = useState<Array<{ text: string; ok: boolean }>>([]);
  const [done, setDone] = useState<string | undefined>(undefined);
  const [confirm, setConfirm] = useState(false);
  const [confirmText, setConfirmText] = useState("");

  const webPath = new URL(siteUrl).pathname.replace(/\/$/, "");
  const say = (text: string, ok: boolean): void => setLog((prev) => prev.concat([{ text, ok }]));

  /* ---------- Load ------------------------------------------------------- */

  /**
   * Staged chains by config item id, read separately from the main query.
   *
   * `PendingLevels` does not exist on a site that has never staged a change, and one unknown
   * name in `$select` fails the WHOLE request with HTTP 400 rather than returning null
   * (gotcha #11) — so folding it in would make this screen unusable on exactly those sites.
   */
  const loadPendingChains = async (): Promise<Record<number, Level[]>> => {
    const out: Record<number, Level[]> = {};
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
          `/items?$select=Id,${PENDING_LEVELS_FIELD}&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return out; // no column means nothing is staged, by definition
      for (const r of ((await res.json()).value ?? []) as Array<Record<string, unknown>>) {
        const raw = r[PENDING_LEVELS_FIELD];
        if (typeof raw !== "string" || raw.trim() === "") continue;
        const parsed = parseLevels(raw);
        if (parsed.length > 0) out[Number(r.Id)] = parsed;
      }
    } catch {
      // Treated as "nothing staged". The consequence is a migration towards the LIVE shape,
      // which finds no drift and moves nothing — the safe direction to fail in.
    }
    return out;
  };

  const loadSegments = async (): Promise<SegmentRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
        `/items?$select=Id,Title,ModeLabel,StagingFolder,Levels&$filter=ConfigType eq 'mode'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`the configuration list returned HTTP ${res.status}`);
    const data = await res.json();
    const pending = await loadPendingChains();
    return ((data.value ?? []) as Array<{
      Id: number; Title?: string; ModeLabel?: string; StagingFolder?: string; Levels?: string;
    }>)
      .map((r) => {
        const chain = parseLevels(r.Levels ?? "");
        const row: SegmentRow = {
          id: r.Id,
          key: (r.Title ?? "").trim(),
          label: (r.ModeLabel ?? r.Title ?? "").trim(),
          stagingFolder: (r.StagingFolder ?? "").trim(),
          chain,
        };
        const staged = pending[r.Id];
        if (staged && staged.length > 0) row.pending = staged;
        // Validate the chain being migrated TOWARDS, not the live one. A malformed staged
        // chain must be refused here — it has no correct destination, and applying it at the
        // end of the run would then break every upload in the segment.
        const err = validateChain(row.pending ?? chain);
        if (err) row.chainError = err.message;
        return row;
      })
      .sort((a, b) => a.label.localeCompare(b.label));
  };

  /** The Year / Document Type term sets, for a segment still running on the built-in pair. */
  const loadLegacyTermSets = async (): Promise<{ year: string; docType: string }> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
        `/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return { year: "", docType: "" };
    const rows = ((await res.json()).value ?? []) as Array<{ Title?: string; SettingValue?: string }>;
    const get = (k: string): string =>
      (rows.filter((r) => (r.Title ?? "").trim() === k)[0]?.SettingValue ?? "").trim();
    return { year: get("termSet_yearPeriod"), docType: get("termSet_documentType") };
  };

  useEffect(() => {
    let cancelled = false;
    // Names first: an unprimed cache resolves to the legacy DMS titles, which 404 on a
    // renamed site and present as "the configuration could not be read".
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      .then(() => loadLegacyTermSets())
      .then((sets) => {
        if (!cancelled) setLegacySets(sets);
      })
      .catch(() => undefined)
      .then(() => loadSegments())
      .then((rows) => {
        if (cancelled) return;
        setSegments(rows);
        setLoading(false);
      })
      .catch((e: Error) => {
        if (cancelled) return;
        setLoadError(e.message);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  /* ---------- Term store ------------------------------------------------- */

  // A Map in a ref, not component state: the scan reads it inside one long async loop, and a
  // state update would not be visible to the next iteration — every unit would re-fetch the
  // same term set. A unit's own SubUnit call is unique, but the flat sets are shared by all.
  const termCache = React.useRef(new Map<string, TermLite[] | undefined>());

  /**
   * Read a term collection. `undefined` means the call FAILED — distinct from an empty array,
   * which means "this term has no children". Every optional-tier decision rests on that
   * difference: empty skips a tier, unknown must skip the whole unit.
   */
  const loadTerms = async (url: string): Promise<TermLite[] | undefined> => {
    const cache = termCache.current;
    if (cache.has(url)) return cache.get(url);
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) {
        cache.set(url, undefined);
        return undefined;
      }
      const data = await res.json();
      const out = ((data.value ?? []) as Array<{ id: string; labels?: Array<{ name?: string; isDefault?: boolean }> }>)
        .map((t) => ({
          id: t.id,
          label:
            (t.labels ?? []).filter((l) => l.isDefault !== false)[0]?.name ??
            (t.labels ?? [])[0]?.name ??
            "",
        }))
        .filter((t) => t.label !== "");
      cache.set(url, out);
      return out;
    } catch {
      cache.set(url, undefined);
      return undefined;
    }
  };

  const setChildren = (setGuid: string): Promise<TermLite[] | undefined> =>
    loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${setGuid}/children`);

  const termChildren = (setGuid: string, termId: string): Promise<TermLite[] | undefined> =>
    loadTerms(`${siteUrl}/_api/v2.1/termStore/sets/${setGuid}/terms/${termId}/children`);

  /* ---------- Folders ---------------------------------------------------- */

  /**
   * Child folders of a path. The path goes in as an OData alias, never an inline literal —
   * gotcha #9: a long encoded path in a quoted literal returns HTTP 400, which reads as a
   * missing folder, and these are the deepest paths in the system.
   */
  const childFolders = async (path: string): Promise<Array<{ name: string; url: string }>> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Folders?$select=Name,ServerRelativeUrl` +
        `&@f='${encodeServerRelativePath(path)}'&$orderby=Name`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return ((data.value ?? []) as Array<{ Name?: string; ServerRelativeUrl?: string }>)
      .filter((f) => (f.Name ?? "") !== "" && f.Name !== "Forms")
      .map((f) => ({ name: f.Name as string, url: f.ServerRelativeUrl as string }));
  };

  const isModerated = async (listTitle: string): Promise<boolean> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')?$select=EnableModeration`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      return res.ok && (await res.json()).EnableModeration === true;
    } catch {
      // Unreadable means "assume not moderated": writing OData__ModerationStatus to a library
      // without moderation fails the whole merge.
      return false;
    }
  };

  /** Stamp a folder Approved. Spec §2.2 — a Pending folder is invisible to the whole unit. */
  const approveFolder = async (path: string): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
        `?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        // Integer, never a quoted string — quoting it makes SharePoint reject the merge.
        body: JSON.stringify({ OData__ModerationStatus: 0 }),
      },
    );
    if (!res.ok) {
      const detail = (await res.text().catch(() => "")).slice(0, 160);
      throw new Error(`HTTP ${res.status} ${detail}`);
    }
  };

  /* ---------- Scan ------------------------------------------------------- */

  /** The chain being migrated TOWARDS: the staged one when there is one, else the live one. */
  const targetChain = (seg: SegmentRow): Level[] => seg.pending ?? seg.chain;

  const belowUnit = (seg: SegmentRow): Level[] =>
    effectiveOnDemandTiers(targetChain(seg), legacySets.year, legacySets.docType);

  /**
   * Make the staged chain live, and clear it.
   *
   * This is the LAST step of the run, never a separate button. Between "folders moved" and
   * "structure applied" every upload would land in the old shape again and re-create exactly
   * the drift just cleaned up — so the two must not be separable in the UI.
   */
  const activatePending = async (seg: SegmentRow): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items(${seg.id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify({
          Levels: JSON.stringify(seg.pending ?? seg.chain),
          [PENDING_LEVELS_FIELD]: "",
        }),
      },
    );
    if (!res.ok) {
      const detail = (await res.text().catch(() => "")).slice(0, 160);
      throw new Error(
        `The folders were moved, but the new structure could not be switched on ` +
          `(HTTP ${res.status}). ${detail} Uploads are still using the old shape — open ` +
          `"Move existing folders" again and it will finish the job.`,
      );
    }
  };

  /**
   * Options for every below-Unit tier of one unit, for CLASSIFICATION.
   *
   * A flat tier (has `termSet`) has one list for every unit. A cascading tier draws from the
   * term above, so tier 0 comes from the unit's own term. For a deeper cascading tier the
   * parent is not chosen yet, so the union of every possible parent's children is used: the
   * question here is only "could this folder name belong at this tier?", and a union answers
   * that without inventing a destination.
   */
  const tierOptionsFor = async (
    tiers: Level[],
    unitTermId: string | undefined,
    segmentSet: string,
  ): Promise<{ options: Array<TermLite[] | undefined>; note?: string }> => {
    const out: Array<TermLite[] | undefined> = [];
    for (let i = 0; i < tiers.length; i++) {
      const set = (tiers[i].termSet ?? "").trim();
      if (set) {
        out.push(await setChildren(set));
        continue;
      }
      if (i === 0) {
        if (!unitTermId) {
          return {
            options: out,
            note: "its folder is not in the Folder Map, so the values that belong under it cannot be read",
          };
        }
        out.push(await termChildren(segmentSet, unitTermId));
        continue;
      }
      const parents = out[i - 1];
      if (!parents) {
        out.push(undefined);
        continue;
      }
      const union: TermLite[] = [];
      let failed = false;
      for (const p of parents) {
        const kids = await termChildren(segmentSet, p.id);
        if (!kids) {
          failed = true;
          break;
        }
        for (const k of kids) if (union.filter((u) => u.id === k.id).length === 0) union.push(k);
      }
      out.push(failed ? undefined : union);
    }
    return { options: out };
  };

  /**
   * Find every unit and how its folders sit against the TARGET chain.
   *
   * Separate from `scan` so the run can call it again afterwards: whether the staged structure
   * gets switched on depends on whether any drift is LEFT, and the only trustworthy answer to
   * that is a fresh look at the folders rather than bookkeeping over what was attempted.
   */
  const collectScans = async (seg: SegmentRow): Promise<UnitScan[]> => {
    if (seg.chainError) throw new Error(seg.chainError);
      const { permissioned } = splitChain(targetChain(seg));
      const tiers = belowUnit(seg);
      if (tiers.length === 0) {
        throw new Error("this segment has no folder levels below Unit, so there is nothing to migrate into.");
      }
      if (!seg.stagingFolder) {
        throw new Error("this segment has no top folder name (StagingFolder) on its configuration row.");
      }

      // The segment term set is needed only when a tier cascades from the level above.
      let setGuid = "";
      if (tiers.filter((t) => !(t.termSet ?? "").trim()).length > 0) {
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
            `/items?$select=TermSetGuid&$filter=ConfigType eq 'mode' and Title eq '${encodeURIComponent(seg.key)}'&$top=1`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (res.ok) setGuid = (((await res.json()).value ?? [])[0]?.TermSetGuid ?? "").trim();
        if (!setGuid) {
          throw new Error(
            "one of the levels below Unit takes its values from the level above, which needs this " +
              "segment's term set — but its configuration row has no TermSetGuid.",
          );
        }
      }

      // Unit folder -> its term, so a cascading tier can read that unit's own values. Keyed on
      // the last (permissioned + 1) path segments, which is identical in both libraries.
      const depth = Math.max(1, permissioned.length);
      const tailOf = (path: string): string =>
        path.split("/").slice(-(depth + 1)).join("/").toLowerCase();
      const termByTail: Record<string, string> = {};
      try {
        for (const row of await loadFolderMapRows(context.spHttpClient, siteUrl)) {
          if (row.folderUrl && row.termGuid) termByTail[tailOf(row.folderUrl)] = row.termGuid;
        }
      } catch {
        // Only cascading tiers need this; those units are then reported as unresolved.
      }

      const libs: LibCtx[] = [
        { key: "Staging", title: libraryTitle(), urlSegment: libraryUrlSegment(), moderated: false },
        { key: "Documents", title: DOCUMENTS_LIST_TITLE, urlSegment: DOCUMENTS_URL_SEGMENT, moderated: false },
      ];
      for (const lib of libs) lib.moderated = await isModerated(lib.title);

      const found: UnitScan[] = [];
      for (const lib of libs) {
        const root = `${webPath}/${lib.urlSegment}/${seg.stagingFolder}`;
        // Walk down to the unit level: `permissioned` counts the levels below the segment folder.
        let level: Array<{ name: string; url: string }> = [{ name: seg.stagingFolder, url: root }];
        for (let d = 0; d < depth; d++) {
          const next: Array<{ name: string; url: string }> = [];
          for (const node of level) next.push(...(await childFolders(node.url)));
          level = next;
        }
        for (const unit of level) {
          const tail = tailOf(unit.url);
          const kids = (await childFolders(unit.url)).map((k) => k.name);
          const { options, note } = await tierOptionsFor(tiers, termByTail[tail], setGuid);
          // A tier whose options could not be READ makes every folder look misplaced, so the
          // unit is skipped rather than planned. Empty is fine — that tier does not apply here.
          const unreadable = options.filter((o) => o === undefined).length > 0;
          const eff = effectiveTiers(options.map((o) => (o ?? []).map((t) => t.label)));
          const row: UnitScan = {
            lib,
            unitPath: unit.url,
            tail,
            children: kids,
            tiers: eff,
            options: eff.map((t) => (options[t.chainIndex] ?? []) as TermLite[]),
          };
          if (note) row.unresolved = note;
          else if (unreadable) row.unresolved = "some of its folder values could not be read from the term store";
          else if (eff.length === 0) row.unresolved = "none of the levels below Unit have values for this unit";
          found.push(row);
        }
      }
      return found;
  };

  const scan = async (): Promise<void> => {
    const seg = segments.filter((x) => x.key === chosen)[0];
    if (!seg) return;
    setScanning(true);
    setScanError(undefined);
    setScans(undefined);
    setDest({});
    setLog([]);
    setDone(undefined);
    try {
      setScans(await collectScans(seg));
    } catch (e) {
      setScanError((e as Error).message);
    } finally {
      setScanning(false);
    }
  };

  /* ---------- Plan ------------------------------------------------------- */

  const planFor = (row: UnitScan): UnitPlan => {
    if (row.unresolved) {
      return { unitPath: row.unitPath, moves: [], strays: [], neededTiers: [], skipped: row.unresolved };
    }
    return planUnit(row.unitPath, row.children, row.tiers, dest[row.tail] ?? []);
  };

  const plans = (scans ?? []).map(planFor);
  const totals = planTotals(plans);

  /** Units needing attention, grouped so both libraries share one set of pickers. */
  const groupsOf = (): Array<{ tail: string; rows: UnitScan[]; needed: number }> => {
    const out: Array<{ tail: string; rows: UnitScan[]; needed: number }> = [];
    for (const row of scans ?? []) {
      const plan = planFor(row);
      const needed = plan.neededTiers.length;
      if (needed === 0 && plan.strays.length === 0) continue;
      const existing = out.filter((o) => o.tail === row.tail)[0];
      if (existing) {
        existing.rows.push(row);
        existing.needed = Math.max(existing.needed, needed);
      } else {
        out.push({ tail: row.tail, rows: [row], needed });
      }
    }
    return out;
  };

  const setDestination = (tail: string, index: number, value: Destination | undefined): void => {
    setDest((prev) => {
      const list = (prev[tail] ?? []).slice();
      list[index] = value;
      // A deeper cascading tier's options come from this choice, so anything below it is no
      // longer valid. Clearing is the honest response — keeping it would offer a value from a
      // different parent, which is how a folder ends up under a term that does not own it.
      for (let i = index + 1; i < list.length; i++) list[i] = undefined;
      return { ...prev, [tail]: list };
    });
  };

  /* ---------- Run -------------------------------------------------------- */

  /** One library's files, read once and reused for every unit. */
  interface LibFile extends FileRow {
    id: number;
  }

  /**
   * Every file in a library, with the tier columns needed to compare against the path.
   *
   * Read WHOLE-LIBRARY and filtered per unit in memory, rather than a CAML query per unit.
   * Two reasons, and the second is the one that matters: `$top` + `@odata.nextLink` pages
   * reliably, whereas `GetItems` returns no paging token in this shape — a per-unit CAML
   * query would silently stop at one page, or repeat page one forever while reporting
   * progress. It also carries `Id`, removing a resolve request per file.
   */
  const readLibraryFiles = async (lib: LibCtx, cols: string[]): Promise<LibFile[]> => {
    const out: LibFile[] = [];
    const select = ["Id", "FileRef", "FileSystemObjectType"].concat(cols).join(",");
    let url: string | undefined =
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lib.title)}')` +
      `/items?$select=${encodeURIComponent(select)}&$top=5000`;
    while (url) {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) {
        const detail = (await res.text().catch(() => "")).slice(0, 160);
        // A named column that does not exist fails the WHOLE request with 400, not a null
        // (gotcha #11) — so this reads as "cannot list the files" when it is a missing column.
        throw new Error(`could not read ${lib.title} (HTTP ${res.status}) ${detail}`);
      }
      const data = await res.json();
      for (const r of (data.value ?? []) as Array<Record<string, unknown>>) {
        if (r.FileSystemObjectType !== 0) continue; // folders describe nothing to stamp
        const values: Record<string, string> = {};
        for (const c of cols) values[c] = typeof r[c] === "string" ? (r[c] as string) : "";
        out.push({ id: Number(r.Id), path: String(r.FileRef ?? ""), values });
      }
      url = data["odata.nextLink"] ?? data["@odata.nextLink"] ?? undefined;
    }
    return out;
  };

  const stampFile = async (
    lib: LibCtx,
    id: number,
    values: Array<{ FieldName: string; FieldValue: string }>,
  ): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lib.title)}')/items(${id})/validateUpdateListItem`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
        // bNewDocumentUpdate stops this counting as a fresh upload, so no new version is cut
        // and no check-out is demanded.
        body: JSON.stringify({ formValues: values, bNewDocumentUpdate: true }),
      },
    );
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    // validateUpdateListItem returns 200 even when every field failed — gotcha #4.
    const results = ((await res.json()).value ?? []) as Array<{
      FieldName?: string; HasException?: boolean; ErrorMessage?: string;
    }>;
    const bad = results.filter((r) => r.HasException);
    if (bad.length > 0) {
      throw new Error(bad.map((b) => `${b.FieldName}: ${b.ErrorMessage ?? "rejected"}`).join("; "));
    }
  };

  /**
   * Switch the staged structure on, but ONLY once nothing is left in the old shape.
   *
   * Applying it while some units are still adrift would be the worst of both worlds: those units
   * would start receiving uploads in the new shape while their existing folders sat a tier above
   * — the exact side-by-side state staging exists to prevent, now with no pending change left
   * to tell anyone it is unfinished.
   *
   * So the check is a FRESH scan, not a tally of what this run attempted. A move that reported
   * success but landed somewhere unexpected can only be caught by looking.
   */
  const finishPending = async (seg: SegmentRow): Promise<string> => {
    if (!seg.pending) return " Turn the two flows back on.";
    let remaining = 0;
    try {
      for (const row of await collectScans(seg)) {
        // With empty destinations, `neededTiers` is "what would still need choosing" — which is
        // precisely the definition of a unit still sitting in the old shape.
        if (planUnit(row.unitPath, row.children, row.tiers, []).neededTiers.length > 0) remaining++;
      }
    } catch (e) {
      return (
        ` The new structure was NOT switched on, because the folders could not be re-checked ` +
        `afterwards (${(e as Error).message}). Uploads are still using the old shape — run this ` +
        `again once that is resolved.`
      );
    }
    if (remaining > 0) {
      return (
        ` The new structure is still waiting: ${remaining} unit(s) have folders in the old shape. ` +
        `Uploads carry on unchanged until every one of them is done, so nobody sees a ` +
        `half-changed library.`
      );
    }
    await activatePending(seg);
    setSegments((prev) =>
      prev.map((p) =>
        p.key === seg.key ? { ...p, chain: seg.pending as Level[], pending: undefined } : p,
      ),
    );
    return (
      ` Every folder is in the new shape, so the new structure is now LIVE — uploaders will see it ` +
      `on their next page load. Turn the two flows back on.`
    );
  };

  /**
   * Part B: make every file's tier columns agree with where the file actually sits.
   *
   * Separate from `run` because it must be reachable WITHOUT a move. The first live run proved
   * why: the folders moved, the tagging failed, and the second attempt found nothing to move —
   * so the only path to part B was closed and the metadata could never be repaired. A repair
   * step reachable only via the thing that broke it is not a repair step.
   *
   * Safe to run at any time and as often as wanted: it derives its work from the paths, writes
   * only what disagrees, and touches no folder.
   */
  const backfillMetadata = async (
    seg: SegmentRow,
    rows: UnitScan[],
  ): Promise<{ stamped: number; failed: number }> => {
    const tiers = belowUnit(seg);
    let stamped = 0;
    let failed = 0;

    /**
     * The column to compare and stamp for a tier — or `undefined` for one this tool must not
     * touch.
     *
     * **A tier with no `tidCol` is MANAGED METADATA**, and the migration leaves it alone. Two
     * independent reasons, either sufficient:
     *
     *   1. A taxonomy field needs `Label|GUID` (gotcha #5). The bare label fails with "The data
     *      returned from the tagging UI was not formatted correctly" — seen live 2026-08-11 on
     *      `Year` and `Document_x0020_Type` — and it took the valid `CreditCard` write down with
     *      it, because one bad field fails the whole call.
     *   2. Nothing needed writing. A migration INSERTS an ancestor tier; the `2024` and `Tax
     *      Return` folders keep their names and values, so those columns were already right.
     *      Reading one back as a plain string yields "" — the value is an object — so every file
     *      looked like it needed a stamp it did not need.
     *
     * The `tidCol` test is exact, not a proxy: the plain-text label+GUID pair is what the
     * multi-segment model writes, and the only tiers without one are the built-in Year /
     * Document Type pair that `effectiveOnDemandTiers` omits it for.
     */
    const colFor = (i: number): string | undefined => {
      const lvl = tiers[i];
      if (!lvl || !lvl.tidCol) return undefined;
      return lvl.labelCol ?? lvl.column;
    };
    const allCols: string[] = [];
    for (let i = 0; i < tiers.length; i++) {
      const col = colFor(i);
      if (col && allCols.indexOf(col) < 0) allCols.push(col);
    }
    if (allCols.length === 0) return { stamped, failed };

    // One read per library, not per unit.
    const filesByLib: Record<string, LibFile[]> = {};
    for (const lib of rows.map((r) => r.lib)) {
      if (filesByLib[lib.key]) continue;
      try {
        filesByLib[lib.key] = await readLibraryFiles(lib, allCols);
      } catch (e) {
        failed++;
        say(`${lib.title}: ${(e as Error).message}`, false);
        filesByLib[lib.key] = [];
      }
    }

    for (const row of rows) {
      if (row.unresolved) continue;
      const label = `${row.lib.key} · ${row.tail}`;
      try {
        const prefix = `${row.unitPath}/`;
        const files = (filesByLib[row.lib.key] ?? []).filter((f) => f.path.indexOf(prefix) === 0);
        const idByPath: Record<string, number> = {};
        for (const f of files) idByPath[f.path] = f.id;
        for (const need of backfillNeeds(row.unitPath, files, row.tiers, colFor)) {
          // Grouped per tier, not flattened, so one rejected tier can be retried alone. ONE bad
          // field name fails the WHOLE validateUpdateListItem call (gotcha #4) — how a single
          // unwritable column took every other column with it, the same failure that once made
          // bulk upload tag nothing at all.
          const groups: Array<{ label: string; values: Array<{ FieldName: string; FieldValue: string }> }> = [];
          for (const field of need.fields) {
            const lvl = tiers[field.chainIndex];
            if (!lvl) continue;
            const labelCol = lvl.labelCol ?? lvl.column;
            if (!labelCol) continue;
            const values = [{ FieldName: labelCol, FieldValue: field.label }];
            const at = row.tiers.map((t) => t.chainIndex).indexOf(field.chainIndex);
            const term = (row.options[at] ?? []).filter(
              (o) => o.label.trim().toLowerCase() === field.label.trim().toLowerCase(),
            )[0];
            // No matching term means no GUID: the label is still written and the id left alone,
            // rather than filled with a guess that would outlive the run.
            if (term) values.push({ FieldName: lvl.tidCol as string, FieldValue: term.id });
            groups.push({ label: lvl.label, values });
          }
          const all: Array<{ FieldName: string; FieldValue: string }> = [];
          for (const g of groups) for (const v of g.values) all.push(v);
          if (all.length === 0) continue;
          try {
            await stampFile(row.lib, idByPath[need.path], all);
            stamped++;
          } catch {
            // Retry tier by tier to salvage the ones that are fine and to name the one that is
            // not. "Could not tag this file", when four of five columns would have written, is
            // how a narrow problem becomes an opaque one.
            let wrote = 0;
            for (const g of groups) {
              try {
                await stampFile(row.lib, idByPath[need.path], g.values);
                wrote++;
              } catch (e2) {
                failed++;
                say(
                  `${label}: could NOT tag ${g.label} on ${need.path.split("/").pop()} — ${(e2 as Error).message}`,
                  false,
                );
              }
            }
            if (wrote > 0) stamped++;
          }
        }
      } catch (e) {
        failed++;
        say(`${label}: ${(e as Error).message}`, false);
      }
    }
    return { stamped, failed };
  };

  /** Part B on its own, from the no-folders-to-move state. */
  const runTagsOnly = async (): Promise<void> => {
    const seg = segments.filter((x) => x.key === chosen)[0];
    if (!seg || !scans) return;
    setRunning(true);
    setLog([]);
    setDone(undefined);
    try {
      const tags = await backfillMetadata(seg, scans);
      setDone(
        tags.stamped === 0 && tags.failed === 0
          ? "Every document's folder columns already match where it sits — nothing to change."
          : `Tagged ${tags.stamped} document(s).` +
            (tags.failed > 0 ? ` ${tags.failed} problem(s) listed above; running this again is safe.` : ""),
      );
    } finally {
      setRunning(false);
    }
  };

  const run = async (): Promise<void> => {
    const seg = segments.filter((x) => x.key === chosen)[0];
    if (!seg || !scans) return;
    setRunning(true);
    setConfirm(false);
    setConfirmText("");
    setLog([]);
    setDone(undefined);
    let moved = 0;
    let stamped = 0;
    let failed = 0;

    try {
      for (const row of scans) {
        const plan = planFor(row);
        const label = `${row.lib.key} · ${row.tail}`;
        if (plan.skipped) {
          say(`${label}: left alone — ${plan.skipped}`, false);
          continue;
        }
        for (const stray of plan.strays) {
          say(`${label}: "${stray}" matches no folder level — left alone`, false);
        }
        for (const move of plan.moves) {
          try {
            for (const ancestor of move.ancestors) {
              const cut = ancestor.path.lastIndexOf("/");
              const made = await ensureFolder(
                context.spHttpClient,
                siteUrl,
                ancestor.path.slice(0, cut),
                ancestor.path.slice(cut + 1),
              );
              if (!made) throw new Error(`could not create ${ancestor.path}`);
              // Left Pending, this folder — and every document moved into it — is invisible
              // to the whole unit. Spec §2.2.
              if (made.created && row.lib.moderated) await approveFolder(ancestor.path);
            }
            const res = await moveFolderTo(context.spHttpClient, siteUrl, move.from, move.to);
            if (!res.ok) {
              failed++;
              say(
                `${label}: could NOT move ${move.name} — ` +
                  (res.conflict ? "a folder of that name is already there" : `HTTP ${res.status}`),
                false,
              );
              continue;
            }
            moved++;
            say(`${label}: ${move.name} → ${move.to}`, true);
          } catch (e) {
            failed++;
            say(`${label}: could NOT move ${move.name} — ${(e as Error).message}`, false);
          }
        }
      }

      // Metadata second, read fresh: the paths have just changed, and re-reading rather than
      // remembering is what makes a half-finished run safe to repeat. Spec §5.1.
      const tags = await backfillMetadata(seg, scans);
      stamped += tags.stamped;
      failed += tags.failed;

      setDone(
        `Moved ${moved} folder(s) and tagged ${stamped} file(s).` +
          (failed > 0
            ? ` ${failed} problem(s) listed above — nothing was deleted, so running this again is safe and will retry them.`
            : "") +
          (await finishPending(seg)),
      );
      // The tree has changed underneath the plan, so a fresh scan is required before another run.
      setScans(undefined);
    } finally {
      setRunning(false);
    }
  };

  /* ---------- Render ----------------------------------------------------- */

  if (loading) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading…</p>;
  if (loadError) return <div style={{ ...s.msg, ...s.err }}>The segments could not be read: {loadError}</div>;

  const seg = segments.filter((x) => x.key === chosen)[0];
  const groups = groupsOf();
  const belowForSeg = seg ? belowUnit(seg) : [];

  return (
    <div>
      <div style={{ ...s.msg, ...s.warn }}>
        <strong>Turn off the two Power Automate flows before running this</strong> — Auto-route and
        the folder-approval flow. Moving a file re-triggers Auto-route, which then fails because the
        file has already moved. Nothing is damaged, but the run history fills with failures and a
        real one becomes hard to spot. Turn them back on afterwards.
      </div>

      <label style={s.label} htmlFor="mig-seg">Business segment</label>
      <select
        id="mig-seg"
        style={s.input}
        value={chosen}
        disabled={scanning || running}
        onChange={(e) => {
          setChosen(e.target.value);
          setScans(undefined);
          setDest({});
          setLog([]);
          setDone(undefined);
          setScanError(undefined);
        }}
      >
        <option value="">Choose a segment…</option>
        {segments.map((x) => (
          <option key={x.key} value={x.key}>{x.label}</option>
        ))}
      </select>
      {seg && (
        <p style={s.hint}>
          {seg.pending === undefined ? (
            <>
              Documents should be filed as{" "}
              {[seg.stagingFolder || seg.label, ...seg.chain.map((l) => `[${l.label}]`)].join(" / ")}
            </>
          ) : (
            // Name both shapes. "Should be filed as" alone would be ambiguous here — one of
            // these is what uploads are doing, the other is what they will do afterwards.
            <>
              Uploads currently use{" "}
              <strong>
                {[seg.stagingFolder || seg.label, ...seg.chain.map((l) => `[${l.label}]`)].join(" / ")}
              </strong>
              . Waiting to be applied:{" "}
              <strong>
                {[seg.stagingFolder || seg.label, ...(seg.pending ?? []).map((l) => `[${l.label}]`)].join(" / ")}
              </strong>
              . The new shape goes live when every folder has been moved.
            </>
          )}
        </p>
      )}

      <div style={{ marginTop: 14 }}>
        <button
          style={chosen && !scanning && !running ? s.btn : s.off}
          disabled={!chosen || scanning || running}
          onClick={scan}
        >
          {scanning ? "Checking…" : "Check for folders in the old shape"}
        </button>
      </div>

      {scanError && <div style={{ ...s.msg, ...s.err, marginTop: 16 }}>{scanError}</div>}

      {scans && groups.length === 0 && !scanError && (
        <div style={{ ...s.msg, ...s.ok, marginTop: 16 }}>
          Nothing to move — every folder in this segment already sits at the right level.
        </div>
      )}

      {/* Tagging must be reachable WITHOUT a move. Proved by the first live run: the folders
          moved, the tagging failed, and the retry found nothing to move — so the only route to
          the repair was closed, and the metadata could never be fixed. */}
      {scans && groups.length === 0 && !scanError && (
        <div style={{ ...s.card, marginTop: 4 }}>
          <p style={{ fontSize: 13, lineHeight: 1.6, margin: "0 0 12px" }}>
            The folder columns on each document should say which folder it is in. Checking is
            harmless and changes only what disagrees.
          </p>
          <button style={running ? s.off : s.ghost} disabled={running} onClick={runTagsOnly}>
            {running ? "Checking…" : "Check document tags"}
          </button>
        </div>
      )}

      {/* A staged change with no folders to move still has to be applyable, or it could never
          go live. Happens whenever the new level only affects units that hold no documents. */}
      {scans && groups.length === 0 && !scanError && seg?.pending !== undefined && (
        <div style={{ ...s.card, marginTop: 4 }}>
          <p style={{ fontSize: 13, lineHeight: 1.6, margin: "0 0 12px" }}>
            <strong>{seg.label}</strong> has a structure change waiting, and nothing needs moving.
            Applying it makes uploaders start using{" "}
            <span style={{ fontFamily: "Consolas, monospace" }}>
              {[seg.stagingFolder || seg.label, ...(seg.pending ?? []).map((l) => `[${l.label}]`)].join(" / ")}
            </span>
            .
          </p>
          <button
            style={running ? s.off : s.btn}
            disabled={running}
            onClick={() => {
              setRunning(true);
              setLog([]);
              activatePending(seg)
                .then(() => {
                  setSegments((prev) =>
                    prev.map((p) =>
                      p.key === seg.key ? { ...p, chain: seg.pending as Level[], pending: undefined } : p,
                    ),
                  );
                  setDone("The new structure is now live — uploaders will see it on their next page load.");
                  setScans(undefined);
                })
                .catch((e: Error) => setDone(e.message))
                .then(() => setRunning(false))
                .catch(() => undefined);
            }}
          >
            {running ? "Applying…" : "Apply the new structure"}
          </button>
        </div>
      )}

      {scans && groups.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3 style={{ fontSize: 15, margin: "0 0 4px" }}>
            {groups.length} unit(s) have folders in the old shape
          </h3>
          <p style={s.hint}>
            Choose where each unit&apos;s old folders should go. Everything inside them moves with
            them, and approved documents stay approved. A unit left as{" "}
            <em>Leave this unit alone</em> is not touched.
          </p>

          {groups.map((group) => {
            const first = group.rows[0];
            const chosenDest = dest[group.tail] ?? [];
            return (
              <div key={group.tail} style={s.card}>
                <div style={s.unitName}>{group.tail}</div>

                {Array.from({ length: group.needed }, (_x, i) => {
                  const tier = first.tiers[i];
                  const level = belowForSeg[tier ? tier.chainIndex : 0];
                  const opts = first.options[i] ?? [];
                  return (
                    <div key={i}>
                      <label style={s.label} htmlFor={`mig-d-${group.tail}-${i}`}>
                        {level ? level.label : `Level ${i + 1}`} for this unit
                      </label>
                      <select
                        id={`mig-d-${group.tail}-${i}`}
                        style={s.input}
                        disabled={running}
                        value={chosenDest[i] ? (chosenDest[i] as Destination).id : ""}
                        onChange={(e) => {
                          const picked = opts.filter((o) => o.id === e.target.value)[0];
                          setDestination(
                            group.tail,
                            i,
                            picked ? { label: picked.label, id: picked.id } : undefined,
                          );
                        }}
                      >
                        <option value="">Leave this unit alone</option>
                        {opts.map((o) => (
                          <option key={o.id} value={o.id}>{o.label}</option>
                        ))}
                      </select>
                    </div>
                  );
                })}

                {group.rows.map((row) => {
                  const plan = planFor(row);
                  return (
                    <div key={row.lib.key + row.unitPath} style={{ marginTop: 10 }}>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#605e5c" }}>
                        {row.lib.title}
                        {plan.skipped && (
                          <span style={{ ...s.badge, background: "#fff4e5", color: "#7a4f00" }}>
                            {plan.skipped}
                          </span>
                        )}
                      </div>
                      {plan.moves.map((m) => (
                        <div key={m.from} style={s.move}>{m.name} &rarr; {m.to}</div>
                      ))}
                      {plan.strays.map((x) => (
                        <div key={x} style={{ ...s.move, color: "#7a4f00" }}>
                          {x} — matches no folder level, will be left alone
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            );
          })}

          <div style={{ marginTop: 18, borderTop: "1px solid #edebe9", paddingTop: 16 }}>
            <button
              style={totals.moves > 0 && !running ? s.btn : s.off}
              disabled={totals.moves === 0 || running}
              onClick={() => setConfirm(true)}
            >
              {running ? "Moving…" : `Move ${totals.moves} folder(s)`}
            </button>
            {totals.moves === 0 && (
              <span style={{ ...s.hint, marginLeft: 10 }}>
                Choose a destination for at least one unit.
              </span>
            )}
          </div>
        </div>
      )}

      {log.length > 0 && (
        <div style={{ ...s.log, marginTop: 20 }}>
          {log.map((line, i) => (
            <div key={i} style={{ ...s.logRow, color: line.ok ? "#0f6c3f" : "#a4262c" }}>
              {line.text}
            </div>
          ))}
        </div>
      )}

      {done && <div style={{ ...s.msg, ...s.ok, marginTop: 16 }}>{done}</div>}

      {confirm && (
        <div style={s.modalBg} role="dialog" aria-modal="true">
          <div style={s.modal}>
            <h3 style={{ margin: "0 0 12px", fontSize: 17 }}>This moves documents that are already filed</h3>
            <p style={{ fontSize: 13, lineHeight: 1.6 }}>
              {totals.moves} folder(s) will move, with everything inside them. Approved documents
              stay approved and nothing is deleted — but <strong>there is no undo</strong>, and
              anyone holding a link to a moved folder will need a new one.
            </p>
            <p style={{ fontSize: 13, lineHeight: 1.6 }}>
              Type <strong>MOVE</strong> to continue.
            </p>
            <input
              style={s.input}
              value={confirmText}
              onChange={(e) => setConfirmText(e.target.value)}
              placeholder="MOVE"
            />
            <div style={{ marginTop: 16 }}>
              <button
                style={confirmText.trim().toUpperCase() === "MOVE" ? s.btn : s.off}
                disabled={confirmText.trim().toUpperCase() !== "MOVE"}
                onClick={run}
              >
                Move them
              </button>{" "}
              <button
                style={s.ghost}
                onClick={() => {
                  setConfirm(false);
                  setConfirmText("");
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
