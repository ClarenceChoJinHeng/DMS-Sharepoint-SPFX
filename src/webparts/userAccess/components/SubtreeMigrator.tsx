import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels, PENDING_LEVELS_FIELD } from "../../../shared/formModel";
import { effectiveOnDemandTiers, splitChain, validateChain } from "../../../shared/folderChain";
import {
  backfillNeeds,
  Collision,
  Destination,
  EffectiveTier,
  effectiveTiers,
  findCollisions,
  LeafFolder,
  LeafPlan,
  planLeaf,
  suggestRename,
} from "../../../shared/subtreeMigration";
import { cachedListTitle, libraryTitle, libraryUrlSegment, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import {
  deleteFolderIfEmpty,
  encodeServerRelativePath,
  ensureFolder,
  loadFolderMapRows,
  moveFileTo,
} from "../../../shared/dmsFolderMap";

/**
 * Subtree Migration — bring documents already filed into the shape the structure now describes.
 *
 * Piece 3 of the configurable folder chain, spec
 * `docs/superpowers/specs/2026-08-11-subtree-migration-design.md`.
 *
 * ONE model for all three edits (spec §3.0): work out which TIER each path segment belongs to,
 * then rebuild the path in tier order. A tier with no segment is a gap the admin fills (an ADD);
 * segments out of order come back sorted (a REORDER); a segment whose tier is gone is absent from
 * the result (a REMOVE, and therefore a collapse). The first build looked only at folders directly
 * under the Unit, which detects an insertion and nothing else — it reported "nothing to move" for a
 * reorder and activated the new structure against folders still in the old shape.
 *
 * Every decision lives in `shared/subtreeMigration.ts` and is unit-tested. This file is the REST
 * calls and the rendering.
 */

const DOCUMENTS_LIST_TITLE = "Documents";
const DOCUMENTS_URL_SEGMENT = "Shared Documents";

/** Depth guard for the folder walk. The deepest legitimate chain is nowhere near this. */
const MAX_DEPTH = 8;

/** Conflict rows rendered before the list is cut short. See the note beside it. */
const CONFLICT_LIMIT = 50;

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
  small:    { boxSizing: "border-box", padding: "5px 8px", fontSize: 12, border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "Consolas, monospace" },
  hint:     { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  log:      { fontSize: 12, fontFamily: "Consolas, monospace", maxHeight: 320, overflowY: "auto", border: "1px solid #e1e1e1", borderRadius: 6, padding: "8px 10px", background: "#fafafa" },
  logRow:   { padding: "2px 0", wordBreak: "break-all" },
  badge:    { fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 10, marginLeft: 8 },
  modalBg:  { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 },
  modal:    { background: "#fff", borderRadius: 8, padding: 24, maxWidth: 560, width: "100%", maxHeight: "80vh", overflowY: "auto" },
  conflict: { border: "1px solid #f0d9b5", background: "#fffdf8", borderRadius: 6, padding: "10px 12px", marginBottom: 10 },
  row:      { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginTop: 6 },
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
   * A chain authored in the Folder levels tab but not yet applied. When present it is the TARGET
   * shape this screen migrates towards, and applying it is the last step of the run.
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

/** Every file in one library, indexed the two ways this screen needs. */
interface FileIndex {
  /** File names by containing folder path — for collision detection. */
  byFolder: Record<string, string[]>;
}

/** What the scan found for one unit in one library. */
interface UnitScan {
  lib: LibCtx;
  unitPath: string;
  /** `nbpolho/cds/upsupport` — identity shared across libraries, so one choice serves both. */
  tail: string;
  /** Target chain tiers that apply to this unit. */
  tiers: EffectiveTier[];
  /** Options per chain index, for the destination pickers and the metadata stamp. */
  optionsByTier: Record<number, TermLite[]>;
  /**
   * Option lists of tiers that USED to exist and no longer do. Lets a segment belonging to a
   * removed tier be told apart from one belonging to nothing — a deliberate collapse versus a
   * folder nobody recognises.
   */
  removedOptions: string[][];
  leaves: LeafFolder[];
  /** Files in a folder that also has subfolders — reported, never moved. */
  looseFiles: string[];
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
  const [fileIndex, setFileIndex] = useState<Record<string, FileIndex>>({});
  /** Destinations per unit tail, keyed by chain index. */
  const [dest, setDest] = useState<Record<string, Record<number, Destination | undefined>>>({});
  /** Resolved collisions: new file name, keyed by SOURCE file path. */
  const [renames, setRenames] = useState<Record<string, string>>({});
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
   * `PendingLevels` does not exist on a site that has never staged a change, and one unknown name
   * in `$select` fails the WHOLE request with HTTP 400 rather than returning null (gotcha #11).
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
      if (!res.ok) return out;
      for (const r of ((await res.json()).value ?? []) as Array<Record<string, unknown>>) {
        const raw = r[PENDING_LEVELS_FIELD];
        if (typeof raw !== "string" || raw.trim() === "") continue;
        const parsed = parseLevels(raw);
        if (parsed.length > 0) out[Number(r.Id)] = parsed;
      }
    } catch {
      // Treated as "nothing staged": the migration then targets the LIVE shape, finds no drift and
      // moves nothing — the safe direction to fail in.
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
        // Validate the chain being migrated TOWARDS. A malformed staged chain has no correct
        // destination, and applying it at the end of the run would break every upload.
        const err = validateChain(row.pending ?? chain);
        if (err) row.chainError = err.message;
        return row;
      })
      .sort((a, b) => a.label.localeCompare(b.label));
  };

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
    // Names first: an unprimed cache resolves to the legacy DMS titles, which 404 on a renamed site
    // and present as "the configuration could not be read".
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

  // A Map in a ref, not state: the scan reads it inside one long async loop, where a state update
  // would not be visible to the next iteration — every unit would re-fetch the same flat set.
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

  /* ---------- Folders and files ------------------------------------------ */

  /**
   * Child folders of a path. The path goes in as an OData alias, never an inline literal —
   * gotcha #9: a long encoded path in a quoted literal returns HTTP 400, which reads as a missing
   * folder, and these are the deepest paths in the system.
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

  /**
   * Every LEAF folder below a unit — one with no subfolders.
   *
   * Leaves are where documents live, because the upload form always creates the whole chain before
   * writing the file. Planning per leaf is what lets one algorithm serve add, reorder and remove:
   * the leaf's own path segments say which tier values it represents.
   */
  const walkLeaves = async (unitPath: string): Promise<LeafFolder[]> => {
    const out: LeafFolder[] = [];
    const visit = async (path: string, segments: string[], depth: number): Promise<void> => {
      if (depth > MAX_DEPTH) return;
      const kids = await childFolders(path);
      if (kids.length === 0) {
        if (segments.length > 0) out.push({ path, segments, files: [] });
        return;
      }
      for (const k of kids) await visit(k.url, segments.concat([k.name]), depth + 1);
    };
    await visit(unitPath, [], 0);
    return out;
  };

  /**
   * Every file in a library, indexed by folder and returned as rows.
   *
   * Read WHOLE-LIBRARY and filtered in memory rather than a query per folder. `$top` +
   * `@odata.nextLink` pages reliably, whereas `GetItems` returns no paging token in this shape — a
   * per-folder query would silently stop at one page, or repeat page one while reporting progress.
   * The rows carry `Id`, which the metadata stamp needs.
   */
  const readLibraryFiles = async (
    lib: LibCtx,
    cols: string[],
  ): Promise<{ index: FileIndex; rows: Array<{ id: number; path: string; values: Record<string, string> }> }> => {
    const index: FileIndex = { byFolder: {} };
    const rows: Array<{ id: number; path: string; values: Record<string, string> }> = [];
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
        // A column named in $select that does not exist fails the WHOLE request with 400, not a
        // null (gotcha #11) — so this reads as "cannot list the files" when it is a missing column.
        throw new Error(`could not read ${lib.title} (HTTP ${res.status}) ${detail}`);
      }
      const data = await res.json();
      for (const r of (data.value ?? []) as Array<Record<string, unknown>>) {
        if (r.FileSystemObjectType !== 0) continue;
        const path = String(r.FileRef ?? "");
        if (!path) continue;
        const folder = path.slice(0, path.lastIndexOf("/"));
        const name = path.slice(path.lastIndexOf("/") + 1);
        if (!index.byFolder[folder]) index.byFolder[folder] = [];
        index.byFolder[folder].push(name);
        const values: Record<string, string> = {};
        for (const c of cols) values[c] = typeof r[c] === "string" ? (r[c] as string) : "";
        rows.push({ id: Number(r.Id), path, values });
      }
      url = data["odata.nextLink"] ?? data["@odata.nextLink"] ?? undefined;
    }
    return { index, rows };
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

  /* ---------- Chains ----------------------------------------------------- */

  /** The chain being migrated TOWARDS: the staged one when there is one, else the live one. */
  const targetChain = (seg: SegmentRow): Level[] => seg.pending ?? seg.chain;

  const belowUnitOf = (chain: Level[]): Level[] =>
    effectiveOnDemandTiers(chain, legacySets.year, legacySets.docType);

  const belowUnit = (seg: SegmentRow): Level[] => belowUnitOf(targetChain(seg));

  /** Below-Unit tiers the LIVE chain has and the target does not — removed by this edit. */
  const removedTiers = (seg: SegmentRow): Level[] => {
    const keys: Record<string, true> = {};
    for (const t of belowUnit(seg)) keys[(t.labelCol ?? t.column ?? "").toLowerCase()] = true;
    return belowUnitOf(seg.chain).filter((l) => !keys[(l.labelCol ?? l.column ?? "").toLowerCase()]);
  };

  /**
   * Options for a list of tiers, for one unit.
   *
   * A flat tier (has `termSet`) has one list for every unit. A cascading tier draws from the term
   * above, so tier 0 comes from the unit's own term. For a deeper cascading tier the parent is not
   * chosen yet, so the union of every possible parent's children is used: the question is only
   * "could this folder name belong at this tier?", and a union answers it without inventing a
   * destination.
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

  /* ---------- Scan ------------------------------------------------------- */

  /**
   * Find every unit, its leaf folders, and the term data needed to place them.
   *
   * Separate from `scan` so the run can call it again afterwards: whether the staged structure gets
   * switched on depends on whether any drift is LEFT, and the only trustworthy answer is a fresh
   * look at the folders rather than bookkeeping over what was attempted.
   */
  const collectScans = async (
    seg: SegmentRow,
  ): Promise<{ rows: UnitScan[]; files: Record<string, FileIndex> }> => {
    if (seg.chainError) throw new Error(seg.chainError);
    const { permissioned } = splitChain(targetChain(seg));
    const tiers = belowUnit(seg);
    const gone = removedTiers(seg);
    if (tiers.length === 0) {
      throw new Error("this segment has no folder levels below Unit, so there is nothing to migrate into.");
    }
    if (!seg.stagingFolder) {
      throw new Error("this segment has no top folder name (StagingFolder) on its configuration row.");
    }

    // The segment term set is needed whenever a tier — target OR removed — cascades from the level
    // above. A removed cascading tier still has to be recognisable, or its folders read as strays.
    let setGuid = "";
    if (tiers.concat(gone).filter((t) => !(t.termSet ?? "").trim()).length > 0) {
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

    // Unit folder -> its term, so a cascading tier can read that unit's own values. Keyed on the
    // last (permissioned + 1) path segments, identical in both libraries.
    const depth = Math.max(1, permissioned.length);
    const tailOf = (path: string): string =>
      path.split("/").slice(-(depth + 1)).join("/").toLowerCase();
    const termByTail: Record<string, string> = {};
    try {
      for (const row of await loadFolderMapRows(context.spHttpClient, siteUrl)) {
        if (row.folderUrl && row.termGuid) termByTail[tailOf(row.folderUrl)] = row.termGuid;
      }
    } catch {
      // Only cascading tiers need it; those units are reported as unresolved.
    }

    const libs: LibCtx[] = [
      { key: "Staging", title: libraryTitle(), urlSegment: libraryUrlSegment(), moderated: false },
      { key: "Documents", title: DOCUMENTS_LIST_TITLE, urlSegment: DOCUMENTS_URL_SEGMENT, moderated: false },
    ];
    for (const lib of libs) lib.moderated = await isModerated(lib.title);

    // Plain-text tier columns only, so one library read serves collision detection here and the
    // metadata stamp later. A managed-metadata tier is excluded — see `colFor` in backfillMetadata.
    const cols: string[] = [];
    for (const t of tiers) {
      const col = t.tidCol ? t.labelCol ?? t.column : undefined;
      if (col && cols.indexOf(col) < 0) cols.push(col);
    }

    const files: Record<string, FileIndex> = {};
    const rows: UnitScan[] = [];
    for (const lib of libs) {
      files[lib.key] = (await readLibraryFiles(lib, cols)).index;
      const root = `${webPath}/${lib.urlSegment}/${seg.stagingFolder}`;
      let level: Array<{ name: string; url: string }> = [{ name: seg.stagingFolder, url: root }];
      for (let d = 0; d < depth; d++) {
        const next: Array<{ name: string; url: string }> = [];
        for (const node of level) next.push(...(await childFolders(node.url)));
        level = next;
      }
      for (const unit of level) {
        const tail = tailOf(unit.url);
        const target = await tierOptionsFor(tiers, termByTail[tail], setGuid);
        const removed = await tierOptionsFor(gone, termByTail[tail], setGuid);
        const unreadable = target.options.filter((o) => o === undefined).length > 0;
        const eff = effectiveTiers(target.options.map((o) => (o ?? []).map((t) => t.label)));
        const optionsByTier: Record<number, TermLite[]> = {};
        for (const t of eff) optionsByTier[t.chainIndex] = (target.options[t.chainIndex] ?? []) as TermLite[];

        const leaves = await walkLeaves(unit.url);
        for (const leaf of leaves) leaf.files = (files[lib.key].byFolder[leaf.path] ?? []).slice();
        // Files in a folder that also has subfolders: their path does not say which tier values
        // they belong to, so there is no destination to derive. Reported, never moved.
        const leafPaths: Record<string, true> = {};
        for (const leaf of leaves) leafPaths[leaf.path] = true;
        const looseFiles: string[] = [];
        for (const folder of Object.keys(files[lib.key].byFolder)) {
          if (folder.indexOf(`${unit.url}/`) !== 0 || leafPaths[folder]) continue;
          for (const f of files[lib.key].byFolder[folder]) looseFiles.push(`${folder}/${f}`);
        }

        const row: UnitScan = {
          lib,
          unitPath: unit.url,
          tail,
          tiers: eff,
          optionsByTier,
          removedOptions: removed.options.map((o) => (o ?? []).map((t) => t.label)),
          leaves,
          looseFiles,
        };
        if (target.note) row.unresolved = target.note;
        else if (unreadable) row.unresolved = "some of its folder values could not be read from the term store";
        else if (eff.length === 0) row.unresolved = "none of the levels below Unit have values for this unit";
        rows.push(row);
      }
    }
    return { rows, files };
  };

  const scan = async (): Promise<void> => {
    const seg = segments.filter((x) => x.key === chosen)[0];
    if (!seg) return;
    setScanning(true);
    setScanError(undefined);
    setScans(undefined);
    setDest({});
    setRenames({});
    setLog([]);
    setDone(undefined);
    try {
      const { rows, files } = await collectScans(seg);
      setScans(rows);
      setFileIndex(files);
    } catch (e) {
      setScanError((e as Error).message);
    } finally {
      setScanning(false);
    }
  };

  /* ---------- Plan ------------------------------------------------------- */

  const plansFor = (row: UnitScan): LeafPlan[] => {
    if (row.unresolved) return [];
    const chosenDest = dest[row.tail] ?? {};
    return row.leaves.map((leaf) =>
      planLeaf(row.unitPath, leaf, row.tiers, chosenDest, row.removedOptions),
    );
  };

  /** Plans that would actually relocate something. */
  const movesOf = (row: UnitScan): LeafPlan[] =>
    plansFor(row).filter((p) => p.to !== undefined && p.to !== p.leaf.path);

  /**
   * Apply the admin's renames, then look for collisions again.
   *
   * Re-running the same tested detector over the RESOLVED names — rather than checking each form row
   * against its siblings — means "is it settled yet?" and "what collides?" cannot drift apart. They
   * are one computation.
   */
  const collisionsFor = (lib: string): Collision[] => {
    const plans: LeafPlan[] = [];
    for (const row of scans ?? []) {
      if (row.lib.key !== lib) continue;
      for (const p of movesOf(row)) {
        const renamed = p.leaf.files.map((f) => renames[`${p.leaf.path}/${f}`] ?? f);
        plans.push({ ...p, leaf: { ...p.leaf, files: renamed } });
      }
    }
    return findCollisions(plans, (fileIndex[lib] ?? { byFolder: {} }).byFolder);
  };

  /** The value that made a file distinct — what the suggested name should carry. */
  const distinguisherFor = (filePath: string): string => {
    const folder = filePath.slice(0, filePath.lastIndexOf("/"));
    for (const row of scans ?? []) {
      for (const p of plansFor(row)) {
        if (p.leaf.path !== folder) continue;
        if (p.dropped.length > 0) return p.dropped.join(" ");
        return p.leaf.segments[0] ?? "";
      }
    }
    return "";
  };

  const setDestination = (tail: string, chainIndex: number, value: Destination | undefined): void => {
    setDest((prev) => {
      const next: Record<number, Destination | undefined> = { ...(prev[tail] ?? {}) };
      if (value) next[chainIndex] = value;
      else delete next[chainIndex];
      return { ...prev, [tail]: next };
    });
  };

  /* ---------- Run -------------------------------------------------------- */

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
        // bNewDocumentUpdate stops this counting as a fresh upload, so no new version is cut and no
        // check-out is demanded.
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
   * Part B: make every file's tier columns agree with where the file actually sits.
   *
   * Reachable WITHOUT a move, by its own button. The first live run proved why: the folders moved,
   * the tagging failed, and the retry found nothing to move — so the only path to the repair was
   * closed. A repair step reachable only through the thing that broke it is not a repair step.
   */
  const backfillMetadata = async (
    seg: SegmentRow,
    rows: UnitScan[],
  ): Promise<{ stamped: number; failed: number }> => {
    const tiers = belowUnit(seg);
    let stamped = 0;
    let failed = 0;

    /**
     * The column to compare and stamp for a tier — or `undefined` for one this tool must not touch.
     *
     * **A tier with no `tidCol` is MANAGED METADATA**, and the migration leaves it alone. Two
     * independent reasons, either sufficient:
     *   1. A taxonomy field needs `Label|GUID` (gotcha #5). The bare label fails with "The data
     *      returned from the tagging UI was not formatted correctly" — seen live 2026-08-11 on
     *      `Year` and `Document_x0020_Type` — and it took the valid `CreditCard` write down with it,
     *      because one bad field fails the whole call.
     *   2. Nothing needed writing. A migration rearranges ANCESTORS; the `2024` and `Tax Return`
     *      folders keep their names and values. Reading one back as a plain string yields "" — the
     *      value is an object — so every file looked like it needed a stamp it did not need.
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

    const byLib: Record<string, Array<{ id: number; path: string; values: Record<string, string> }>> = {};
    for (const lib of rows.map((r) => r.lib)) {
      if (byLib[lib.key]) continue;
      try {
        byLib[lib.key] = (await readLibraryFiles(lib, allCols)).rows;
      } catch (e) {
        failed++;
        say(`${lib.title}: ${(e as Error).message}`, false);
        byLib[lib.key] = [];
      }
    }

    for (const row of rows) {
      if (row.unresolved) continue;
      const label = `${row.lib.key} · ${row.tail}`;
      try {
        const prefix = `${row.unitPath}/`;
        const files = (byLib[row.lib.key] ?? []).filter((f) => f.path.indexOf(prefix) === 0);
        const idByPath: Record<string, number> = {};
        for (const f of files) idByPath[f.path] = f.id;
        for (const need of backfillNeeds(row.unitPath, files, row.tiers, colFor)) {
          // Grouped per tier, not flattened, so one rejected tier can be retried alone. ONE bad
          // field name fails the WHOLE validateUpdateListItem call (gotcha #4).
          const groups: Array<{ label: string; values: Array<{ FieldName: string; FieldValue: string }> }> = [];
          for (const field of need.fields) {
            const lvl = tiers[field.chainIndex];
            if (!lvl) continue;
            const labelCol = lvl.labelCol ?? lvl.column;
            if (!labelCol) continue;
            const values = [{ FieldName: labelCol, FieldValue: field.label }];
            const term = (row.optionsByTier[field.chainIndex] ?? []).filter(
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
            // Retry tier by tier to salvage the ones that are fine and name the one that is not.
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
   * Switch the staged structure on, but ONLY once nothing movable is left in the old shape.
   *
   * The check is a FRESH scan, not a tally of what this run attempted: a move that reported success
   * but landed somewhere unexpected can only be caught by looking.
   *
   * A STRAY does not block activation — this tool cannot resolve one at all, and letting it hold a
   * structure change hostage forever would leave the client with no way forward. It is named in the
   * result instead, so the trade is visible rather than silent.
   */
  const finishPending = async (seg: SegmentRow): Promise<string> => {
    if (!seg.pending) return " Turn the two flows back on.";
    let outstanding = 0;
    let strays = 0;
    try {
      const fresh = await collectScans(seg);
      for (const row of fresh.rows) {
        if (row.unresolved) {
          outstanding++;
          continue;
        }
        const chosenDest = dest[row.tail] ?? {};
        for (const leaf of row.leaves) {
          const p = planLeaf(row.unitPath, leaf, row.tiers, chosenDest, row.removedOptions);
          if (p.strays.length > 0) strays++;
          else if (p.to === undefined || p.to !== p.leaf.path) outstanding++;
        }
      }
    } catch (e) {
      return (
        ` The new structure was NOT switched on, because the folders could not be re-checked ` +
        `afterwards (${(e as Error).message}). Uploads are still using the old shape — run this ` +
        `again once that is resolved.`
      );
    }
    if (outstanding > 0) {
      return (
        ` The new structure is still waiting: ${outstanding} folder(s) are not in the new shape yet. ` +
        `Uploads carry on unchanged until every one is done, so nobody sees a half-changed library.`
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
      `on their next page load.` +
      (strays > 0
        ? ` ${strays} folder(s) were left alone because their names match no folder level; they are ` +
          `listed above and still need a decision.`
        : "") +
      ` Turn the two flows back on.`
    );
  };

  const run = async (): Promise<void> => {
    const seg = segments.filter((x) => x.key === chosen)[0];
    if (!seg || !scans) return;
    setRunning(true);
    setConfirm(false);
    setConfirmText("");
    setLog([]);
    setDone(undefined);
    let movedFiles = 0;
    let failed = 0;
    let removedFolders = 0;
    const emptied: Array<{ lib: LibCtx; path: string; unitPath: string }> = [];

    try {
      for (const row of scans) {
        const label = `${row.lib.key} · ${row.tail}`;
        if (row.unresolved) {
          say(`${label}: left alone — ${row.unresolved}`, false);
          continue;
        }
        for (const loose of row.looseFiles) {
          say(`${label}: ${loose.split("/").pop()} sits outside the folder levels — left alone`, false);
        }
        for (const plan of plansFor(row)) {
          if (plan.strays.length > 0) {
            say(`${label}: "${plan.strays.join(", ")}" matches no folder level — left alone`, false);
            continue;
          }
          if (plan.to === undefined) continue; // needs a destination; reported in the UI
          if (plan.to === plan.leaf.path) continue;

          try {
            // Ancestors first, then the destination leaf folder itself.
            for (const ancestor of plan.ancestors) {
              const cut = ancestor.path.lastIndexOf("/");
              const made = await ensureFolder(
                context.spHttpClient,
                siteUrl,
                ancestor.path.slice(0, cut),
                ancestor.path.slice(cut + 1),
              );
              if (!made) throw new Error(`could not create ${ancestor.path}`);
              // Left Pending, this folder — and every document moved into it — is invisible to the
              // whole unit. Spec §2.2.
              if (made.created && row.lib.moderated) await approveFolder(ancestor.path);
            }
            const cut = plan.to.lastIndexOf("/");
            const leafMade = await ensureFolder(
              context.spHttpClient,
              siteUrl,
              plan.to.slice(0, cut),
              plan.to.slice(cut + 1),
            );
            if (!leafMade) throw new Error(`could not create ${plan.to}`);
            if (leafMade.created && row.lib.moderated) await approveFolder(plan.to);

            for (const file of plan.leaf.files) {
              const from = `${plan.leaf.path}/${file}`;
              const toName = renames[from] ?? file;
              const res = await moveFileTo(context.spHttpClient, siteUrl, from, `${plan.to}/${toName}`);
              if (!res.ok) {
                failed++;
                say(
                  `${label}: could NOT move ${file} — ` +
                    (res.conflict
                      ? `a file called ${toName} is already there`
                      : `HTTP ${res.status} ${res.detail ?? ""}`),
                  false,
                );
                continue;
              }
              movedFiles++;
              say(`${label}: ${file} → ${plan.to.slice(row.unitPath.length + 1)}/${toName}`, true);
            }
            // Queue the emptied source for cleanup, after every move is done.
            emptied.push({ lib: row.lib, path: plan.leaf.path, unitPath: row.unitPath });
          } catch (e) {
            failed++;
            say(`${label}: could NOT move ${plan.leaf.segments.join("/")} — ${(e as Error).message}`, false);
          }
        }
      }

      // Cleanup last, deepest first, and only folders this run emptied. `deleteFolderIfEmpty`
      // re-checks at the moment of deletion — a file uploaded mid-run lands in a folder the plan
      // believes it emptied, and deleting that would destroy a document nobody was migrating.
      const candidates: Array<{ lib: LibCtx; path: string }> = [];
      for (const e of emptied) {
        let path = e.path;
        while (path.length > e.unitPath.length && path.indexOf(e.unitPath) === 0) {
          if (candidates.filter((c) => c.path === path && c.lib.key === e.lib.key).length === 0) {
            candidates.push({ lib: e.lib, path });
          }
          path = path.slice(0, path.lastIndexOf("/"));
        }
      }
      candidates.sort((a, b) => b.path.split("/").length - a.path.split("/").length);
      for (const c of candidates) {
        const res = await deleteFolderIfEmpty(context.spHttpClient, siteUrl, c.path);
        if (res.deleted) {
          removedFolders++;
          continue;
        }
        // "Still holds something" is the normal case for an ancestor with other branches, so it is
        // only worth reporting when the folder failed for some other reason.
        if (res.reason && res.status !== 404 && !/still holds/.test(res.reason)) {
          say(`${c.lib.key}: could not tidy up ${c.path} — ${res.reason}`, false);
        }
      }

      const tags = await backfillMetadata(seg, scans);
      failed += tags.failed;

      setDone(
        `Moved ${movedFiles} document(s), tidied ${removedFolders} empty folder(s), tagged ` +
          `${tags.stamped} document(s).` +
          (failed > 0
            ? ` ${failed} problem(s) listed above — nothing holding a document was deleted, so ` +
              `running this again is safe and will retry them.`
            : "") +
          (await finishPending(seg)),
      );
      setScans(undefined);
    } finally {
      setRunning(false);
    }
  };

  /** Part B on its own, from the nothing-to-move state. */
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

  /* ---------- Render ----------------------------------------------------- */

  if (loading) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading…</p>;
  if (loadError) return <div style={{ ...s.msg, ...s.err }}>The segments could not be read: {loadError}</div>;

  const seg = segments.filter((x) => x.key === chosen)[0];
  const belowForSeg = seg ? belowUnit(seg) : [];
  const rowsWithWork = (scans ?? []).filter((r) => {
    if (r.unresolved !== undefined) return true;
    if (movesOf(r).length > 0) return true;
    return plansFor(r).filter((p) => p.strays.length > 0 || p.to === undefined).length > 0;
  });
  const conflicts: Array<{ lib: string; collision: Collision }> = [];
  for (const lib of ["Staging", "Documents"]) {
    for (const c of collisionsFor(lib)) conflicts.push({ lib, collision: c });
  }
  let totalMoves = 0;
  let needingChoice = 0;
  for (const r of scans ?? []) {
    totalMoves += movesOf(r).length;
    needingChoice += plansFor(r).filter((p) => p.missingTiers.length > 0).length;
  }
  const canRun = totalMoves > 0 && conflicts.length === 0;

  /** Units grouped by tail, so one set of pickers serves both libraries. */
  const groups: Array<{ tail: string; rows: UnitScan[] }> = [];
  for (const row of rowsWithWork) {
    const existing = groups.filter((g) => g.tail === row.tail)[0];
    if (existing) existing.rows.push(row);
    else groups.push({ tail: row.tail, rows: [row] });
  }

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
          setRenames({});
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
            <>
              Uploads currently use{" "}
              <strong>
                {[seg.stagingFolder || seg.label, ...seg.chain.map((l) => `[${l.label}]`)].join(" / ")}
              </strong>
              . Waiting to be applied:{" "}
              <strong>
                {[seg.stagingFolder || seg.label, ...(seg.pending ?? []).map((l) => `[${l.label}]`)].join(" / ")}
              </strong>
              . The new shape goes live when every folder is in it.
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
          Nothing to move — every folder in this segment already sits where the structure says.
        </div>
      )}

      {/* Tagging must be reachable WITHOUT a move: the first live run moved the folders, failed the
          tagging, then found nothing to move — closing the only route to the repair. */}
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
            {/* "0 folder(s) to rebuild" is true but useless: it reads as "nothing to do" when the
                real state is "waiting for you". Lead with what is being asked for. */}
            {totalMoves > 0
              ? `${totalMoves} folder(s) to rebuild across ${groups.length} unit(s)`
              : `${groups.length} unit(s) need a value chosen before anything can move`}
          </h3>
          <p style={s.hint}>
            Documents move into the shape the structure describes. Approved documents stay approved,
            and the only folders deleted are ones left completely empty.
            {needingChoice > 0 && ` ${needingChoice} folder(s) are waiting on a value for a new level.`}
          </p>

          {groups.map((group) => {
            const first = group.rows[0];
            const chosenDest = dest[group.tail] ?? {};
            // Counted, not just collected: a picker labelled only "Credit_Card for this unit" looks
            // like it is about to overwrite the folders that already have one. Saying how many
            // folders are actually missing it makes clear it only fills the gaps.
            const missingCount: Record<number, number> = {};
            for (const row of group.rows) {
              for (const p of plansFor(row)) {
                for (const t of p.missingTiers) missingCount[t] = (missingCount[t] ?? 0) + 1;
              }
            }
            const needed = Object.keys(missingCount)
              .map((k) => Number(k))
              .sort((a, b) => a - b);
            return (
              <div key={group.tail} style={s.card}>
                <div style={s.unitName}>{group.tail}</div>

                {needed.map((chainIndex) => {
                  const level = belowForSeg[chainIndex];
                  const opts = first.optionsByTier[chainIndex] ?? [];
                  return (
                    <div key={chainIndex}>
                      <label style={s.label} htmlFor={`mig-d-${group.tail}-${chainIndex}`}>
                        {level ? level.label : `Level ${chainIndex + 1}`} — for the{" "}
                        {missingCount[chainIndex]} folder(s) below that have no value for it
                      </label>
                      <select
                        id={`mig-d-${group.tail}-${chainIndex}`}
                        style={s.input}
                        disabled={running}
                        value={chosenDest[chainIndex] ? (chosenDest[chainIndex] as Destination).id : ""}
                        onChange={(e) => {
                          const picked = opts.filter((o) => o.id === e.target.value)[0];
                          setDestination(
                            group.tail,
                            chainIndex,
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

                {group.rows.map((row) => (
                  <div key={row.lib.key + row.unitPath} style={{ marginTop: 10 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: "#605e5c" }}>
                      {row.lib.title}
                      {row.unresolved && (
                        <span style={{ ...s.badge, background: "#fff4e5", color: "#7a4f00" }}>
                          {row.unresolved}
                        </span>
                      )}
                    </div>
                    {plansFor(row).map((p) => {
                      if (p.strays.length > 0) {
                        return (
                          <div key={p.leaf.path} style={{ ...s.move, color: "#7a4f00" }}>
                            {p.leaf.segments.join(" / ")} — &quot;{p.strays.join(", ")}&quot; matches no
                            folder level, will be left alone
                          </div>
                        );
                      }
                      if (p.to === undefined) {
                        return (
                          <div key={p.leaf.path} style={{ ...s.move, color: "#7a4f00" }}>
                            {p.leaf.segments.join(" / ")} — needs a value chosen above
                          </div>
                        );
                      }
                      if (p.to === p.leaf.path) return null;
                      return (
                        <div key={p.leaf.path} style={s.move}>
                          {p.leaf.segments.join(" / ")} &rarr; {p.to.slice(row.unitPath.length + 1)}
                          {p.leaf.files.length > 0 && ` (${p.leaf.files.length} document(s))`}
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            );
          })}

          {conflicts.length > 0 && (
            <div style={{ marginTop: 20 }}>
              <h3 style={{ fontSize: 15, margin: "0 0 4px", color: "#a4262c" }}>
                {conflicts.length} filename clash(es) to settle first
              </h3>
              <p style={s.hint}>
                Removing a level puts documents that were in separate folders into the same one.
                These share a name, so all but one of each group must be renamed — the extension is
                kept automatically. Nothing moves until every clash is settled.
              </p>
              <div style={{ marginTop: 10 }}>
                <button
                  style={running ? s.off : s.ghost}
                  disabled={running}
                  onClick={() => {
                    // One decision instead of N. The suggestion carries the value that made each
                    // file distinct, which is the information the removal is about to destroy.
                    setRenames((prev) => {
                      const next = { ...prev };
                      for (const { collision } of conflicts) {
                        const movers = collision.claimants.filter((c) => !c.existing);
                        // The first mover keeps its name only when nothing already at the
                        // destination owns it; otherwise every mover needs a new one.
                        const keepFirst = movers.length === collision.claimants.length;
                        movers.forEach((c, i) => {
                          if (keepFirst && i === 0) return;
                          const name = c.path.slice(c.path.lastIndexOf("/") + 1);
                          next[c.path] = suggestRename(name, distinguisherFor(c.path));
                        });
                      }
                      return next;
                    });
                  }}
                >
                  Use suggested names for all
                </button>
              </div>
              {conflicts.slice(0, CONFLICT_LIMIT).map(({ lib, collision }) => (
                <div key={`${lib}${collision.folder}${collision.name}`} style={{ ...s.conflict, marginTop: 10 }}>
                  <div style={{ fontSize: 12, fontFamily: "Consolas, monospace", color: "#605e5c" }}>
                    {lib} · {collision.folder}/<strong>{collision.name}</strong>
                  </div>
                  {collision.claimants.map((c) => {
                    const name = c.path.slice(c.path.lastIndexOf("/") + 1);
                    return (
                      <div key={c.path} style={s.row}>
                        <span style={{ ...s.move, flex: "1 1 260px" }}>
                          {c.existing ? "already there: " : "from "}
                          {c.path}
                        </span>
                        {c.existing ? (
                          <span style={{ ...s.hint, flex: "0 0 240px" }}>keeps its name</span>
                        ) : (
                          <input
                            style={{ ...s.small, flex: "0 0 240px" }}
                            disabled={running}
                            value={renames[c.path] ?? name}
                            onChange={(e) => setRenames((prev) => ({ ...prev, [c.path]: e.target.value }))}
                          />
                        )}
                      </div>
                    );
                  })}
                </div>
              ))}
              {conflicts.length > CONFLICT_LIMIT && (
                <p style={{ ...s.hint, color: "#a4262c" }}>
                  Showing the first {CONFLICT_LIMIT}. {conflicts.length} clashes is a sign the level
                  being removed carries real meaning — consider keeping it, or removing its values
                  one at a time.
                </p>
              )}
            </div>
          )}

          <div style={{ marginTop: 18, borderTop: "1px solid #edebe9", paddingTop: 16 }}>
            <button
              style={canRun && !running ? s.btn : s.off}
              disabled={!canRun || running}
              onClick={() => setConfirm(true)}
            >
              {running ? "Working…" : `Rebuild ${totalMoves} folder(s)`}
            </button>
            {conflicts.length > 0 && (
              <span style={{ ...s.hint, marginLeft: 10, color: "#a4262c" }}>
                Settle the filename clashes first.
              </span>
            )}
            {conflicts.length === 0 && totalMoves === 0 && (
              <span style={{ ...s.hint, marginLeft: 10 }}>Choose a value for at least one unit.</span>
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
              Documents in {totalMoves} folder(s) will be moved into the new shape. Approved documents
              stay approved, and the only folders deleted are ones left completely empty — but{" "}
              <strong>there is no undo</strong>, and anyone holding a link to a moved folder will need
              a new one.
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
                onClick={() => {
                  run().catch(() => undefined);
                }}
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
