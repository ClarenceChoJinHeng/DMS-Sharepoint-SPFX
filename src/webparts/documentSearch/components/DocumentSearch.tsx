// CRS Search — find a document by its metadata.
//
// Spec: docs/superpowers/specs/2026-08-16-document-search-design.md
//
// THIS WEB PART ENFORCES NO PERMISSIONS, and must never be changed to.
//
// Both engines are security-trimmed by SharePoint before a row reaches this code: the Search API
// applies ACLs at query time, and a REST list read returns only items the caller can open. So the
// hierarchy the client described — PIC sees their unit, Head of Department sees the department,
// C-Level sees the segment or everything — falls out of the folder ACLs reconciliation already
// grants. A role check written here could only ever be a second, drifting copy of a model whose
// source of truth is groupMapModel.ts and the live ACLs, and unlike this design it could be wrong
// in the dangerous direction.
//
// WHAT IT IS NOT: a privacy screen. A PIC searching finds their whole unit's approved documents,
// because that is what they can already open by browsing. My Submissions is the author-filtered
// page and stays that way.
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDocumentSearchProps } from "./IDocumentSearchProps";
import {
  LibraryResult,
  SearchCriteria,
  SearchHit,
  SearchLibrary,
  buildKql,
  buildListFilter,
  buildRecentFilter,
  emptyCriteria,
  failedLibraries,
  hasCriteria,
  isHcLibrary,
  kqlPathScope,
  mergeHits,
  recencyCutoff,
  searchState,
  sortByModified,
} from "../../../shared/documentSearch";
// The metadata panel DERIVES its tier rows from the item's own fields, which is what makes
// Region / Estate·Mill appear on a segment nobody wrote code for. A third hand-written copy is
// avoided deliberately: both existing copies carried the hardcoded-tier bug.
import { buildDetailRows, formatBytes } from "../../../shared/documentDetails";
// The preview strategy is already built and tested — SharePoint serves an Office file as a
// DOWNLOAD, so a raw URL in an iframe renders nothing. Reused rather than re-guessed.
import { previewTarget } from "../../../shared/filePreview";
import {
  DOCUMENTS_LIBRARY,
  LIST_SUFFIX,
  cachedHcLibraries,
  cachedListTitle,
  libApiTitle,
  libraryTitle,
} from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

/* ─────────────────────────────── Term set ids ─────────────────────────────── */

/**
 * The three fixed dropdowns' term sets.
 *
 * Read from the DMS Config `termSet_*` rows at mount; these are only the fallback for when that
 * read fails, and they are THIS site's ids. Term set ids are per site since the in-site term store
 * pivot (2026-07-27), so a hardcoded id is wrong on any other site — which is why the config read
 * wins and this is a last resort rather than the source.
 */
const FALLBACK_TERM_SETS = {
  documentType: "866c5754-258e-401f-8685-03d20ae59b1d",
  year: "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf",
  confidentiality: "0d6d1da8-27e5-477f-8684-e8cf169f8fb9",
};

/* ─────────────────────────────── Types ─────────────────────────────── */

/** One tier of a segment's chain, as the filter panel needs it. */
interface Level {
  label: string;
  /** The item column's internal name — what both engines filter on. */
  column: string;
  /** True unless the Levels entry says `"permissioned": false`. ABSENT MEANS TRUE. */
  permissioned: boolean;
  /** A below-Unit tier's own term set, when it declares one. */
  termSet: string;
}

/** A segment offered in the filter panel. */
interface Segment {
  /** The `Title` of the mode row — `mode_gho`. */
  key: string;
  /** What the client sees, and what is filed in `Business Segment`. */
  label: string;
  termSetGuid: string;
  levels: Level[];
}

/** One option in a dropdown. */
interface TermOption {
  id: string;
  label: string;
}

/** A row as either read returns it. An index signature, never a typed shape — see MySubmissions. */
type RawRow = Record<string, unknown>;

/**
 * Coerce whatever SharePoint returned into text.
 *
 * Managed metadata arrives as an OBJECT, and `(v ?? "").trim()` on one threw
 * "(intermediate value).trim is not a function" and killed the whole My Submissions page with an
 * error naming no field. A row interface typed `string` was the real culprit there: it made
 * TypeScript vouch for something SharePoint does not guarantee.
 */
function textOf(v: unknown): string {
  if (v === undefined || v === null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  const o = v as Record<string, unknown>;
  if (typeof o.Label === "string") return o.Label.trim();
  if (typeof o.Title === "string") return o.Title.trim();
  return "";
}

/** `DD/MMM/YYYY` — the agreed client format. Display only; sort and filter use the stored value. */
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function formatDate(iso: string): string {
  const raw = (iso ?? "").trim();
  if (raw.length === 0) return "";
  const d = new Date(raw);
  if (isNaN(d.getTime())) return raw;
  const day = `0${d.getDate()}`.slice(-2);
  return `${day}/${MONTHS[d.getMonth()]}/${d.getFullYear()}`;
}

/** The folder trail of a server-relative path, library segment and file name removed. */
function trailOf(path: string): string {
  return (path ?? "")
    .split("/")
    .filter((p) => p.length > 0)
    .slice(0, -1)
    .join(" › ");
}

/* ─────────────────────────────── Styles ─────────────────────────────── */

const s: Record<string, React.CSSProperties> = {
  root: { fontFamily: '"Segoe UI", system-ui, sans-serif', color: "#242424" },
  h2: { fontSize: 20, fontWeight: 600, margin: "0 0 4px" },
  sub: { fontSize: 13, color: "#616161", margin: "0 0 16px" },
  bar: { display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" },
  box: {
    flex: "1 1 260px", minWidth: 0, height: 36, padding: "0 12px", fontSize: 14,
    border: "1px solid #d1d1d1", borderRadius: 4, boxSizing: "border-box",
  },
  btn: {
    height: 36, padding: "0 18px", fontSize: 14, fontWeight: 600, border: "none",
    borderRadius: 4, background: "#107c10", color: "#fff", cursor: "pointer",
  },
  btnGhost: {
    height: 36, padding: "0 14px", fontSize: 14, border: "1px solid #d1d1d1",
    borderRadius: 4, background: "#fff", color: "#242424", cursor: "pointer",
  },
  filters: {
    display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
    gap: 10, padding: 14, background: "#faf9f8", border: "1px solid #edebe9",
    borderRadius: 6, marginBottom: 14,
  },
  field: { display: "flex", flexDirection: "column", gap: 4, minWidth: 0 },
  label: { fontSize: 12, fontWeight: 600, color: "#424242" },
  input: {
    height: 32, padding: "0 8px", fontSize: 13, border: "1px solid #d1d1d1",
    borderRadius: 4, background: "#fff", boxSizing: "border-box", maxWidth: "100%",
  },
  row: {
    display: "flex", gap: 12, alignItems: "flex-start", padding: "12px 10px",
    borderBottom: "1px solid #f0f0f0", cursor: "pointer", textAlign: "left",
    background: "none", border: "none", width: "100%", font: "inherit",
  },
  name: { fontSize: 14, fontWeight: 600, color: "#0f548c", wordBreak: "break-word" },
  meta: { fontSize: 12, color: "#616161", marginTop: 2, wordBreak: "break-word" },
  chip: {
    display: "inline-block", fontSize: 11, fontWeight: 600, padding: "1px 7px",
    borderRadius: 10, marginLeft: 8, verticalAlign: "middle",
  },
  chipHc: { background: "#fde7e9", color: "#a4262c" },
  chipPending: { background: "#fff4ce", color: "#8a6100" },
  note: { fontSize: 13, padding: "10px 12px", borderRadius: 4, marginBottom: 12 },
  noteWarn: { background: "#fff4ce", border: "1px solid #f2d16b", color: "#5d4200" },
  empty: { padding: "28px 12px", textAlign: "center", color: "#616161", fontSize: 14 },
  detailWrap: { display: "flex", gap: 20, flexWrap: "wrap", alignItems: "flex-start" },
  preview: { flex: "1 1 420px", minWidth: 0, border: "1px solid #edebe9", borderRadius: 6 },
  panel: { flex: "0 1 340px", minWidth: 260 },
  dRow: { display: "flex", gap: 10, padding: "6px 0", borderBottom: "1px solid #f5f5f5", fontSize: 13 },
  dLabel: { flex: "0 0 130px", color: "#616161" },
  dValue: { flex: "1 1 auto", minWidth: 0, wordBreak: "break-word" },
};

/* ─────────────────────────────── Component ─────────────────────────────── */

export default function DocumentSearch({ context, pageSize }: IDocumentSearchProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  // Origin with no /sites/… — a server-relative path already carries the site, so previewTarget
  // needs the bare host to build an absolute file URL.
  const tenantRoot = siteUrl.replace(/^(https?:\/\/[^/]+).*$/, "$1");
  const sp = context.spHttpClient;

  const [criteria, setCriteria] = useState<SearchCriteria>(emptyCriteria());
  const [segments, setSegments] = useState<Segment[]>([]);
  const [fixedOptions, setFixedOptions] = useState<Record<string, TermOption[]>>({});
  /** Options per tier column, filled as the cascade is walked. */
  const [tierOptions, setTierOptions] = useState<Record<string, TermOption[]>>({});
  /** The chosen term id per tier column — the cascade needs the id, the query needs the label. */
  const [tierIds, setTierIds] = useState<Record<string, string>>({});

  const [results, setResults] = useState<LibraryResult[]>([]);
  const [searched, setSearched] = useState(false);
  const [busy, setBusy] = useState(false);
  const [shown, setShown] = useState(pageSize);
  const [open, setOpen] = useState<SearchHit | undefined>(undefined);
  const [fieldText, setFieldText] = useState<Record<string, string> | undefined>(undefined);
  /** The Documents library's real URL segment — `Shared Documents`, not `Documents` (gotcha #12). */
  const [docsSegment, setDocsSegment] = useState(DOCUMENTS_LIBRARY);
  const [configWarning, setConfigWarning] = useState("");

  /* A run counter, so a slow earlier search cannot overwrite a faster later one. A ref rather than
     state: it must be readable inside an async closure at its CURRENT value, and state read there
     is the render-time value — the same stale-closure trap that recorded zero counts during
     reconciliation. */
  const runId = useRef(0);

  const jsonGet = async (url: string): Promise<{ ok: boolean; status: number; body: RawRow }> => {
    try {
      const res: SPHttpClientResponse = await sp.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: "application/json;odata=nometadata" },
      });
      if (!res.ok) return { ok: false, status: res.status, body: {} };
      return { ok: true, status: res.status, body: (await res.json()) as RawRow };
    } catch {
      // 0 = the request never completed. Kept distinct from a real status on screen, because a
      // network failure and a 403 share nothing except the word "failed".
      return { ok: false, status: 0, body: {} };
    }
  };

  const readTerms = async (url: string): Promise<TermOption[]> => {
    const r = await jsonGet(url);
    const out: TermOption[] = [];
    for (const t of (r.body.value as RawRow[]) ?? []) {
      const labels = (t.labels as RawRow[]) ?? [];
      const name = labels.length > 0 ? textOf(labels[0].name) : "";
      const tid = textOf(t.id);
      if (name.length > 0 && tid.length > 0) out.push({ id: tid, label: name });
    }
    return out.sort((a, b) => a.label.localeCompare(b.label));
  };

  /* ── Mount: names, segments, the three fixed term sets ─────────────────── */

  useEffect(() => {
    let cancelled = false;

    const boot = async (): Promise<void> => {
      await primeNames(sp, siteUrl).catch(() => undefined);
      if (cancelled) return;

      /* The Documents URL segment, resolved not assumed. Its title is `Documents` and its URL is
         `/Shared Documents/` — hardcoding one string for both put "Shared Documents" at the head of
         every folder trail in My Submissions. */
      const docs = await jsonGet(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(DOCUMENTS_LIBRARY)}')` +
          `/RootFolder?$select=ServerRelativeUrl`,
      );
      if (!cancelled && docs.ok) {
        const parts = textOf(docs.body.ServerRelativeUrl).split("/").filter((p) => p.length > 0);
        const seg = parts.length > 0 ? parts[parts.length - 1] : "";
        if (seg.length > 0) setDocsSegment(seg);
      }

      const cfg = cachedListTitle(LIST_SUFFIX.config);

      /* Segments and their tier chains, from the DMS Config mode rows — the same source the upload
         form cascades over. NOT a written-out Department/Unit pair: the admin NAMES the tiers at
         onboarding, so a hardcoded list is only ever right for the segments it was written for, and
         a filter that silently matches nothing reads as "there are no such documents". */
      const modes = await jsonGet(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cfg)}')/items` +
          `?$select=Title,ModeLabel,TermSetGuid,Levels,SortOrder&$filter=ConfigType eq 'mode'&$top=100`,
      );
      if (cancelled) return;
      if (!modes.ok) {
        // Named, not silent. Without segments the tier filters cannot be offered at all, and an
        // admin needs to know the cause is a config read rather than an empty repository.
        setConfigWarning(
          "Segment and tier filters are unavailable — the configuration list could not be read " +
            `(${modes.status === 0 ? "the request did not complete" : `HTTP ${modes.status}`}). ` +
            "Free-text and the other filters still work.",
        );
      } else {
        const list: Segment[] = [];
        for (const row of (modes.body.value as RawRow[]) ?? []) {
          const label = textOf(row.ModeLabel) || textOf(row.Title);
          if (label.length === 0) continue;
          const levels: Level[] = [];
          try {
            const parsed = (JSON.parse(textOf(row.Levels) || "[]") as RawRow[]) ?? [];
            for (const lv of parsed) {
              const column = textOf(lv.column);
              if (column.length === 0) continue;
              levels.push({
                label: textOf(lv.label) || column,
                column,
                // ABSENT MEANS TRUE, and only a literal false demotes a tier — the string "false"
                // does not. The same rule the upload form's chain follows.
                permissioned: lv.permissioned !== false,
                termSet: textOf(lv.termSet),
              });
            }
          } catch {
            // A malformed chain costs this one segment its tier filters, never the whole panel.
            levels.length = 0;
          }
          list.push({ key: textOf(row.Title), label, termSetGuid: textOf(row.TermSetGuid), levels });
        }
        if (!cancelled) setSegments(list);
      }

      /* The three fixed dropdowns. Ids come from DMS Config where present — they are per site since
         the in-site term store pivot, so the constants above are a last resort. */
      const settings = await jsonGet(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cfg)}')/items` +
          `?$select=Title,SettingValue&$filter=startswith(Title,'termSet_')&$top=50`,
      );
      const configured: Record<string, string> = {};
      for (const row of (settings.body.value as RawRow[]) ?? []) {
        configured[textOf(row.Title)] = textOf(row.SettingValue);
      }
      const setFor = (key: string, fallback: string): string => {
        const v = configured[`termSet_${key}`];
        return v !== undefined && v.length > 0 ? v : fallback;
      };
      const base = `${siteUrl}/_api/v2.1/termStore/sets`;
      const loaded = await Promise.all([
        readTerms(`${base}/${setFor("documentType", FALLBACK_TERM_SETS.documentType)}/children`),
        readTerms(`${base}/${setFor("yearPeriod", FALLBACK_TERM_SETS.year)}/children`),
        readTerms(`${base}/${setFor("confidentiality", FALLBACK_TERM_SETS.confidentiality)}/children`),
      ]);
      if (!cancelled) {
        setFixedOptions({ documentType: loaded[0], year: loaded[1], confidentiality: loaded[2] });
      }
    };

    boot().catch(() => undefined);
    return () => {
      cancelled = true;
    };
    // Mount only: neither the site nor the client changes under a live page.
  }, []);

  /* ── The tier cascade ──────────────────────────────────────────────────── */

  const segment = segments.filter((x) => x.label === criteria.segment)[0];

  /** Load one tier's options: children of the chosen parent term, or of the set at the top. */
  const loadTierOptions = async (seg: Segment, index: number, parentId: string): Promise<void> => {
    const level = seg.levels[index];
    if (!level) return;
    const base = `${siteUrl}/_api/v2.1/termStore/sets`;
    // A below-Unit tier has its own term set and no place in the segment's own tree.
    const url = level.permissioned
      ? parentId.length > 0
        ? `${base}/${seg.termSetGuid}/terms/${parentId}/children`
        : `${base}/${seg.termSetGuid}/children`
      : level.termSet.length > 0
        ? `${base}/${level.termSet}/children`
        : "";
    const options = url.length === 0 ? [] : await readTerms(url);
    setTierOptions((prev) => ({ ...prev, [level.column]: options }));
  };

  const chooseSegment = (label: string): void => {
    setCriteria((c) => ({ ...c, segment: label, tiers: [] }));
    setTierOptions({});
    setTierIds({});
    const seg = segments.filter((x) => x.label === label)[0];
    if (seg && seg.levels.length > 0) loadTierOptions(seg, 0, "").catch(() => undefined);
  };

  const chooseTier = (index: number, termId: string, label: string): void => {
    if (!segment) return;
    const level = segment.levels[index];
    if (!level) return;

    /* Everything BELOW the tier just changed is now meaningless — cleared rather than left in
       place, because a stale Unit under a newly chosen Department queries a combination that cannot
       exist and returns nothing at all, which reads as "there are no documents here". */
    const deeper = segment.levels.slice(index + 1).map((l) => l.column);
    setTierIds((prev) => {
      const next: Record<string, string> = { ...prev, [level.column]: termId };
      for (const col of deeper) delete next[col];
      return next;
    });
    setTierOptions((prev) => {
      const next = { ...prev };
      for (const col of deeper) delete next[col];
      return next;
    });
    setCriteria((c) => {
      const kept = (c.tiers ?? []).filter(
        (t) => t.column !== level.column && deeper.indexOf(t.column) === -1,
      );
      return {
        ...c,
        tiers: label.length > 0 ? [...kept, { column: level.column, value: label }] : kept,
      };
    });

    if (termId.length > 0 && index + 1 < segment.levels.length) {
      loadTierOptions(segment, index + 1, termId).catch(() => undefined);
    }
  };

  /* ── Reads ─────────────────────────────────────────────────────────────── */

  const libraryLabel = (lib: SearchLibrary): string => {
    const hc = cachedHcLibraries();
    if (lib === "Documents") return DOCUMENTS_LIBRARY;
    if (lib === "Staging") return libraryTitle();
    if (lib === "DocumentsHC") return hc ? hc.documents.title : "HC Documents";
    if (lib === "StagingHC") return hc ? hc.approval.title : "HC Approval Document";
    return lib;
  };

  /** Which approved-side library a crawled result came from, by its path. */
  const libraryOfPath = (path: string): SearchLibrary => {
    const hc = cachedHcLibraries();
    const p = (path ?? "").toLowerCase();
    if (hc && p.indexOf(`/${hc.documents.urlSegment.toLowerCase()}/`) !== -1) return "DocumentsHC";
    return "Documents";
  };

  /**
   * The Search half — ONE request covering both approved-side libraries.
   *
   * One rather than two because Search trims by ACL, so an uncleared user simply gets no HC rows.
   * There is no 403 to interpret here, which is why only the live reads carry a `refused` outcome.
   */
  const runSearch = async (c: SearchCriteria): Promise<LibraryResult[]> => {
    const hc = cachedHcLibraries();
    const segs = [docsSegment];
    if (hc) segs.push(hc.documents.urlSegment);
    const query = buildKql(c, kqlPathScope(siteUrl, segs));
    // Blank means "do not run" — never "match everything". An unscoped KQL query searches the whole
    // tenant, which would show documents from sites this system has nothing to do with.
    if (query.length === 0) return [];

    const props = "Path,Filename,Title,LastModifiedTime,Author,Size,UniqueId,ListItemID";
    const r = await jsonGet(
      `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(query)}'` +
        `&rowlimit=200&trimduplicates=false&selectproperties='${encodeURIComponent(props)}'`,
    );
    if (!r.ok) return [{ library: "Documents", outcome: "failed", status: r.status, hits: [] }];

    const primary = (r.body.PrimaryQueryResult as RawRow) ?? {};
    const relevant = (primary.RelevantResults as RawRow) ?? {};
    const table = (relevant.Table as RawRow) ?? {};
    const hits: SearchHit[] = [];
    for (const row of (table.Rows as RawRow[]) ?? []) {
      const map: Record<string, string> = {};
      for (const cell of (row.Cells as RawRow[]) ?? []) map[textOf(cell.Key)] = textOf(cell.Value);
      const path = map.Path ?? "";
      if (path.length === 0) continue;
      const name = map.Filename || map.Title || "";
      hits.push({
        // Search returns it braced and the list read does not — normalised so the two halves dedupe
        // against each other rather than showing the same document twice.
        uniqueId: (map.UniqueId ?? "").replace(/[{}]/g, ""),
        itemId: Number(map.ListItemID ?? "0") || 0,
        library: libraryOfPath(path),
        name: name.length > 0 ? name : path.split("/").filter((p) => p.length > 0).pop() ?? "",
        // Reduced to server-relative, so the trail and the preview agree with the live reads.
        path: path.replace(/^https?:\/\/[^/]+/, ""),
        modified: map.LastModifiedTime ?? "",
        author: map.Author ?? "",
        size: map.Size ?? "",
      });
    }
    return [{ library: "Documents", outcome: "ok", hits }];
  };

  /** One live list read — used for the approval libraries and for the recency top-up. */
  const runListRead = async (lib: SearchLibrary, filter: string): Promise<LibraryResult> => {
    const r = await jsonGet(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')/items` +
        `?$select=Id,FileLeafRef,FileRef,Modified,Author/Title,File/Length,File/UniqueId` +
        `&$expand=Author,File&$filter=${encodeURIComponent(filter)}&$orderby=Modified desc&$top=200`,
    );
    if (!r.ok) {
      /* A 403 is an uncleared user meeting an HC library, and that is the CORRECT answer to their
         query — not an error. Counting it as one would put a permanent warning in front of everyone
         who is not HC-cleared, and train them to ignore the banner on the day it means something. */
      const refused = isHcLibrary(lib) && (r.status === 403 || r.status === 401);
      return { library: lib, outcome: refused ? "refused" : "failed", status: r.status, hits: [] };
    }
    const hits: SearchHit[] = [];
    for (const row of (r.body.value as RawRow[]) ?? []) {
      const file = (row.File as RawRow) ?? {};
      const author = (row.Author as RawRow) ?? {};
      hits.push({
        uniqueId: textOf(file.UniqueId).replace(/[{}]/g, ""),
        itemId: Number(textOf(row.Id)) || 0,
        library: lib,
        name: textOf(row.FileLeafRef),
        path: textOf(row.FileRef),
        modified: textOf(row.Modified),
        author: textOf(author.Title),
        size: textOf(file.Length),
      });
    }
    return { library: lib, outcome: "ok", hits };
  };

  const search = async (): Promise<void> => {
    if (!hasCriteria(criteria)) return;
    runId.current += 1;
    const mine = runId.current;
    setBusy(true);
    setOpen(undefined);
    setShown(pageSize);

    const hc = cachedHcLibraries();
    const listFilter = buildListFilter(criteria);
    // Generous against a crawl measured in minutes: too tight leaves exactly the invisible gap this
    // exists to close, and the cost of being generous is a few extra rows to dedupe.
    const recentFilter = buildRecentFilter(criteria, recencyCutoff(new Date(), 24));

    const jobs: Promise<LibraryResult[]>[] = [
      runSearch(criteria),
      runListRead("Staging", listFilter).then((x) => [x]),
      /* The gap between the engines: a file approved two minutes ago has left the approval library
         and is not yet crawled, so it is in NEITHER. `Modified` is indexable and the window keeps
         the set tiny, so this stays threshold-safe at any library size — which is precisely why the
         main read could not be done this way. */
      runListRead("Documents", recentFilter).then((x) => [x]),
    ];
    if (hc) {
      jobs.push(runListRead("StagingHC", listFilter).then((x) => [x]));
      jobs.push(runListRead("DocumentsHC", recentFilter).then((x) => [x]));
    }

    // Promise.allSettled is unavailable (the tsconfig predates it), and each job already resolves to
    // an outcome rather than rejecting.
    const settled = await Promise.all(jobs);
    if (runId.current !== mine) return; // a later search has already answered
    const flat: LibraryResult[] = [];
    for (const group of settled) for (const one of group) flat.push(one);
    setResults(flat);
    setSearched(true);
    setBusy(false);
  };

  const runSearchSafely = (): void => {
    search().catch(() => setBusy(false));
  };

  const reset = (): void => {
    setCriteria(emptyCriteria());
    setTierOptions({});
    setTierIds({});
    setResults([]);
    setSearched(false);
    setOpen(undefined);
  };

  const openRow = (h: SearchHit): void => {
    setOpen(h);
    setFieldText(undefined);
    /* FieldValuesAsText is PER ITEM, which is why the list shows no metadata at all — one request
       per row would be hundreds. It also returns LABELS, which is what makes a taxonomy value
       readable instead of the bare lookup id that reached the screen on 2026-08-14. */
    jsonGet(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(h.library))}')` +
        `/items(${h.itemId})/FieldValuesAsText`,
    )
      .then((r) => setFieldText(r.ok ? (r.body as Record<string, string>) : {}))
      .catch(() => setFieldText({}));
  };

  /* ── Render ────────────────────────────────────────────────────────────── */

  const allHits = sortByModified(mergeHits(...results.map((r) => r.hits)));
  const state = searchState(searched, results);
  const failed = failedLibraries(results);

  if (open) {
    const preview = previewTarget(open.name, open.path, tenantRoot, siteUrl);
    const rows = buildDetailRows({
      fieldText: fieldText ?? {},
      leading: [
        { label: "Library", value: libraryLabel(open.library) },
        { label: "Location", value: trailOf(open.path) },
      ],
      trailing: [
        { label: "Uploaded by", value: open.author },
        { label: "Last updated", value: formatDate(open.modified) },
        { label: "File size", value: formatBytes(open.size) },
      ],
    });

    return (
      <div style={s.root}>
        <button
          type="button"
          style={{ ...s.btnGhost, marginBottom: 14 }}
          onClick={() => setOpen(undefined)}
        >
          ‹ Back to results
        </button>
        <h2 style={s.h2}>
          {open.name}
          {isHcLibrary(open.library) ? (
            <span style={{ ...s.chip, ...s.chipHc }}>Highly Confidential</span>
          ) : undefined}
        </h2>
        <div style={s.detailWrap}>
          <div style={s.preview}>
            {preview.kind === "none" ? (
              <div style={s.empty}>
                No preview for this file type.{" "}
                <a href={preview.fileUrl} target="_blank" rel="noreferrer">
                  Open it in SharePoint
                </a>
              </div>
            ) : preview.kind === "image" ? (
              // Fit to WIDTH and scroll — the same fix the approval page needed.
              <div style={{ maxHeight: 600, overflow: "auto" }}>
                <img src={preview.url} alt={open.name} style={{ width: "100%", display: "block" }} />
              </div>
            ) : (
              <iframe
                title={open.name}
                src={preview.url}
                style={{ width: "100%", height: 600, border: "none" }}
              />
            )}
          </div>
          <div style={s.panel}>
            {fieldText === undefined ? (
              <div style={s.meta}>Reading details…</div>
            ) : rows.length === 0 ? (
              <div style={s.meta}>No details were recorded for this document.</div>
            ) : (
              rows.map((r) => (
                <div key={`${r.label}:${r.value}`} style={s.dRow}>
                  <div style={s.dLabel}>{r.label}</div>
                  <div style={s.dValue}>{r.value}</div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={s.root}>
      <h2 style={s.h2}>Search documents</h2>
      <p style={s.sub}>
        Finds documents by their details. You will only ever see documents you are already allowed
        to open — searching does not give access to anything new.
      </p>

      {configWarning.length > 0 ? <div style={{ ...s.note, ...s.noteWarn }}>{configWarning}</div> : undefined}

      <div style={s.bar}>
        <input
          style={s.box}
          placeholder="Search by file name, project, vendor, remark or content…"
          value={criteria.text}
          onChange={(e) => setCriteria((c) => ({ ...c, text: e.target.value }))}
          onKeyDown={(e) => {
            if (e.key === "Enter") runSearchSafely();
          }}
        />
        <button type="button" style={s.btn} disabled={busy || !hasCriteria(criteria)} onClick={runSearchSafely}>
          {busy ? "Searching…" : "Search"}
        </button>
        <button type="button" style={s.btnGhost} onClick={reset}>
          Clear
        </button>
      </div>

      <div style={s.filters}>
        <div style={s.field}>
          <label style={s.label} htmlFor="crs-doctype">Document type</label>
          <select
            id="crs-doctype"
            style={s.input}
            value={criteria.documentType}
            onChange={(e) => setCriteria((c) => ({ ...c, documentType: e.target.value }))}
          >
            <option value="">Any</option>
            {(fixedOptions.documentType ?? []).map((o) => (
              <option key={o.id} value={o.label}>{o.label}</option>
            ))}
          </select>
        </div>

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-year">Year</label>
          <select
            id="crs-year"
            style={s.input}
            value={criteria.year}
            onChange={(e) => setCriteria((c) => ({ ...c, year: e.target.value }))}
          >
            <option value="">Any</option>
            {(fixedOptions.year ?? []).map((o) => (
              <option key={o.id} value={o.label}>{o.label}</option>
            ))}
          </select>
        </div>

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-conf">Confidentiality</label>
          <select
            id="crs-conf"
            style={s.input}
            value={criteria.confidentiality}
            onChange={(e) => setCriteria((c) => ({ ...c, confidentiality: e.target.value }))}
          >
            <option value="">Any</option>
            {(fixedOptions.confidentiality ?? []).map((o) => (
              <option key={o.id} value={o.label}>{o.label}</option>
            ))}
          </select>
        </div>

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-segment">Business segment</label>
          <select
            id="crs-segment"
            style={s.input}
            value={criteria.segment}
            onChange={(e) => chooseSegment(e.target.value)}
          >
            <option value="">Any</option>
            {segments.map((seg) => (
              <option key={seg.key} value={seg.label}>{seg.label}</option>
            ))}
          </select>
        </div>

        {/* Tier filters, DERIVED from the selected segment's own chain. Nothing here names
            Department or Unit — Upstream Ops files under Region and Estate/Mill, and every segment
            onboarded from here names its own. A tier appears once the tier above it is chosen,
            because its options are that term's children. */}
        {segment
          ? segment.levels.map((level, i) => {
              const options = tierOptions[level.column];
              const parentChosen =
                i === 0 || (tierIds[segment.levels[i - 1].column] ?? "").length > 0;
              if (!parentChosen) return undefined;
              const current = (criteria.tiers ?? []).filter((t) => t.column === level.column)[0];
              return (
                <div style={s.field} key={level.column}>
                  <label style={s.label} htmlFor={`crs-tier-${level.column}`}>{level.label}</label>
                  <select
                    id={`crs-tier-${level.column}`}
                    style={s.input}
                    value={current ? current.value : ""}
                    onChange={(e) => {
                      const picked = (options ?? []).filter((o) => o.label === e.target.value)[0];
                      chooseTier(i, picked ? picked.id : "", e.target.value);
                    }}
                  >
                    <option value="">{options === undefined ? "Loading…" : "Any"}</option>
                    {(options ?? []).map((o) => (
                      <option key={o.id} value={o.label}>{o.label}</option>
                    ))}
                  </select>
                </div>
              );
            })
          : undefined}

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-dfrom">Document date from</label>
          <input
            id="crs-dfrom"
            type="date"
            style={s.input}
            value={criteria.documentDateFrom}
            onChange={(e) => setCriteria((c) => ({ ...c, documentDateFrom: e.target.value }))}
          />
        </div>
        <div style={s.field}>
          <label style={s.label} htmlFor="crs-dto">Document date to</label>
          <input
            id="crs-dto"
            type="date"
            style={s.input}
            value={criteria.documentDateTo}
            onChange={(e) => setCriteria((c) => ({ ...c, documentDateTo: e.target.value }))}
          />
        </div>
        <div style={s.field}>
          <label style={s.label} htmlFor="crs-ufrom">Uploaded from</label>
          <input
            id="crs-ufrom"
            type="date"
            style={s.input}
            value={criteria.uploadedFrom}
            onChange={(e) => setCriteria((c) => ({ ...c, uploadedFrom: e.target.value }))}
          />
        </div>
        <div style={s.field}>
          <label style={s.label} htmlFor="crs-uto">Uploaded to</label>
          <input
            id="crs-uto"
            type="date"
            style={s.input}
            value={criteria.uploadedTo}
            onChange={(e) => setCriteria((c) => ({ ...c, uploadedTo: e.target.value }))}
          />
        </div>
      </div>

      {/* A PARTIAL failure shows its results AND says what is missing. Rendering it as a clean list
          silently understates what is there; rendering it as "no results" is how someone concludes
          a document was never filed. */}
      {failed.length > 0 ? (
        <div style={{ ...s.note, ...s.noteWarn }}>
          {failed.length === 1
            ? "One library could not be searched"
            : `${failed.length} libraries could not be searched`}
          {": "}
          {failed
            .map(
              (f) =>
                `${libraryLabel(f.library)} (${
                  f.status === 0 ? "the request did not complete" : `HTTP ${f.status}`
                })`,
            )
            .join(", ")}
          . Anything filed there is missing from these results.
        </div>
      ) : undefined}

      {state === "idle" ? (
        <div style={s.empty}>Type something above, or pick a filter, then choose Search.</div>
      ) : state === "error" ? (
        <div style={s.empty}>
          The search could not be completed, so this is <strong>not</strong> a statement that nothing
          matched. See the message above, then try again.
        </div>
      ) : state === "empty" ? (
        <div style={s.empty}>
          No documents matched. Note that a document approved in the last few minutes may not be
          findable yet.
        </div>
      ) : (
        <div>
          <div style={{ ...s.meta, marginBottom: 6 }}>
            {allHits.length === 1 ? "1 document" : `${allHits.length} documents`}
          </div>
          {allHits.slice(0, shown).map((h) => (
            <button
              type="button"
              key={`${h.library}#${h.itemId}#${h.uniqueId}`}
              style={s.row}
              onClick={() => openRow(h)}
            >
              <div style={{ minWidth: 0, flex: "1 1 auto" }}>
                <div style={s.name}>
                  {h.name}
                  {isHcLibrary(h.library) ? <span style={{ ...s.chip, ...s.chipHc }}>HC</span> : undefined}
                  {h.library === "Staging" || h.library === "StagingHC" ? (
                    <span style={{ ...s.chip, ...s.chipPending }}>Awaiting approval</span>
                  ) : undefined}
                </div>
                <div style={s.meta}>{trailOf(h.path)}</div>
                <div style={s.meta}>
                  {[h.author, formatDate(h.modified), formatBytes(h.size)]
                    .filter((x) => x.length > 0)
                    .join(" · ")}
                </div>
              </div>
            </button>
          ))}
          {allHits.length > shown ? (
            <button
              type="button"
              style={{ ...s.btnGhost, marginTop: 12 }}
              onClick={() => setShown((n) => n + pageSize)}
            >
              Show more ({allHits.length - shown} remaining)
            </button>
          ) : undefined}
        </div>
      )}
    </div>
  );
}
