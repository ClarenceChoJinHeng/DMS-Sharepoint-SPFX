import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IBulkUploadProps } from "./IBulkUploadProps";
import {
  lookupFolderMapping,
  resolveFolderServerUrl,
  probeFolderByPath,
  ensureFolder,
} from "../../../shared/dmsFolderMap";
import {
  parseLevels,
  collectMembership,
  isChainAuthorized,
  sanitizeFolderSegment,
  buildLevelFormValues,
  Level,
  GroupMapRow,
  Membership,
  ColumnPair,
} from "../../../shared/formModel";

/* ----------------------------------------------------------------------------
 * BULK UPLOAD — a duplicate of the `form` web part, with two behaviour changes:
 *   1. Writes straight into the Documents library, bypassing Staging (and so
 *      bypassing content approval and the Auto-route flow).
 *   2. Files are queued in up to MAX_BATCHES independent "batches" — each batch
 *      has its own destination folder + metadata and up to MAX_FILES files.
 *      "Upload all" processes the batches sequentially.
 * TEMPORARY TOOL. See:
 *   docs/superpowers/specs/2026-07-24-bulk-upload-two-batch-design.md
 *   docs/superpowers/specs/2026-07-23-bulk-upload-direct-to-documents-design.md
 *
 * Metadata tagging, group-based mode detection, the term cascade and per-Unit
 * folder routing are cloned from Form.tsx unchanged.
 * -------------------------------------------------------------------------- */

const MAX_FILES = 50; // per batch
const MAX_BATCHES = 2;

// The Documents library's real URL segment differs from its display title:
//   URL segment   = "Shared Documents"  (used to build server-relative paths)
//   display title = "Documents"         (used by lists/getbytitle for metadata)
// This matches the Auto-route flow's own concat('/Shared Documents/', …) target.
const DOCUMENTS_URL_SEGMENT = "Shared Documents";
const DOCUMENTS_LIST_TITLE = "Documents";

const FIELDS = {
  // Document Type column — internal name Document_x0020_Type (migrated from the old
  // frozen "Department_x0020_Type"; see 2026-07-24-document-type-internal-name-migration-design).
  documentType: "Document_x0020_Type",
  yearPeriod: "Year",            // site column internal name (client kept plain "Year")
  documentDate: "DocumentDate",
  confidentiality: "Confidentiality_x0020_Level",
  // "Vendor/CustomerName" — the "/" encodes to _x002f_ in the internal name.
  // Verified against /fields 2026-07-28. The old "Vendor" column was deleted.
  vendor: "Vendor_x002f_CustomerName",
  details: "_ExtendedDescription",
};

// FALLBACK map: logical Levels `column` key -> the two real internal names
// (label + term GUID). The preferred, portable source is each level's own
// labelCol/tidCol in DMS Config Levels JSON — those win over this table
// (see buildLevelFormValues), so a client tenant needs no code change here.
const LEVEL_COLUMNS: Record<string, ColumnPair> = {
  BusinessSegment: {
    label: "Business_x0020_Segment",
    tid: "BusinessSegmentTid",
  },
  Department: { label: "Department", tid: "DepartmentTid" },
  Unit: { label: "Unit", tid: "UnitTid" },
};

// validateUpdateListItem validates dates against the SITE's regional settings.
// This tenant is US locale (M/D/YYYY) — ISO YYYY-MM-DD is rejected.
const toSpDate = (iso: string): string => {
  const [y, m, d] = iso.split("-");
  return `${Number(m)}/${Number(d)}/${y}`;
};

// Managed-metadata single value: "Label|GUID". Empty when unset.
const taxVal = (label: string, id: string): string =>
  label && id ? `${label}|${id}` : "";

type TermOption = { id: string; label: string };
type ToastType = "error" | "success";

type UploadMode = {
  key: string;
  label: string;
  side: "BusinessSegment" | "Project";
  termSetGuid: string;
  stagingFolder: string;
  levels: Level[];
  sortOrder: number;
};

// A fully authorised upload path for a restricted (non-privileged) user.
type ValidPath = { modeKey: string; chain: TermOption[] };

// One picked file in the draft config panel (no per-file rename any more).
type PickedFile = { key: string; file: File };

type Outcome = "uploaded" | "skipped" | "failed" | "tagFailed";
type FileResult = { name: string; outcome: Outcome; detail?: string };

// A resolved level term (label + id) plus the pair of column internal names it
// writes to — everything needed to rebuild formValues without the live cascade.
type BatchSelection = {
  column: string;
  label: string;
  id: string;
  labelCol?: string;
  tidCol?: string;
};

// A fully-configured, self-contained batch. Snapshots the draft at Save time so
// it no longer depends on any live cascade/option state.
type Batch = {
  id: string;
  files: File[];
  modeKey: string;
  modeLabel: string;
  termSetGuid: string;
  levelSelections: BatchSelection[];
  docTypeId: string;
  docTypeLabel: string;
  yearId: string;
  yearLabel: string;
  confId: string;
  confLabel: string;
  vendorId: string;
  vendorLabel: string;
  documentDate: string; // ISO yyyy-mm-dd as entered; converted at upload
  destinationLabel: string;
};

// Per-batch upload outcome for the results panel.
type BatchOutcome = {
  batchId: string;
  label: string;
  destinationLabel: string;
  batchError?: string; // set when the whole batch could not be routed
  results: FileResult[];
};

// Live per-file state for the upload progress bars. Byte-accurate: the upload
// itself goes over XMLHttpRequest so `upload.onprogress` can drive a real
// percentage (spHttpClient exposes no progress events, which is why this used to
// be a looping indeterminate bar).
type FileState =
  | "pending"
  | "uploading"
  | "done"
  | "skipped"
  | "failed"
  | "tagFailed"
  | "deleted";
type LiveBatch = {
  label: string;
  destinationLabel: string;
  // `sru` is the uploaded file's ServerRelativeUrl, captured once a file lands in
  // the Documents library (done / tagFailed) so the per-row X can delete it after
  // the fact. `pct` is real bytes-sent progress, 0-100, only meaningful while
  // state === "uploading".
  files: { name: string; size: number; state: FileState; sru?: string; pct: number }[];
};

type OptionMap = {
  documentType: TermOption[];
  yearPeriod: TermOption[];
  confidentiality: TermOption[];
  vendor: TermOption[];
};

const EMPTY_OPTIONS: OptionMap = {
  documentType: [],
  yearPeriod: [],
  confidentiality: [],
  vendor: [],
};

// Fallback if DMS Config is missing/unreachable. DMS Config is the source of
// truth at runtime — these built-in modes only serve an offline fallback.
const DEFAULT_MODES: UploadMode[] = [
  {
    key: "gho",
    label: "Group Head Office",
    side: "BusinessSegment",
    termSetGuid: "df4b9afa-d3b9-4c04-9097-50dcaf5d8036",
    stagingFolder: "Group Head Office",
    levels: [
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ],
    sortOrder: 1,
  },
  {
    key: "upstream_my_ho",
    label: "Upstream Malaysia Head Office",
    side: "BusinessSegment",
    termSetGuid: "16a52947-57a3-4217-9a49-b48cb8b0dd31",
    stagingFolder: "Upstream Malaysia Head Office",
    levels: [
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ],
    sortOrder: 2,
  },
  {
    key: "minamas_ho",
    label: "Minamas Head Office",
    side: "BusinessSegment",
    termSetGuid: "c6b26d32-1c3e-441f-b4ea-78f053e12990",
    stagingFolder: "Minamas Head Office",
    levels: [
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ],
    sortOrder: 3,
  },
  {
    key: "nbpol_ho",
    label: "NBPOL Head Office",
    side: "BusinessSegment",
    termSetGuid: "303f2c38-086a-46ba-8ee0-85445f6bfa3a",
    stagingFolder: "NBPOL Head Office",
    levels: [
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ],
    sortOrder: 4,
  },
  {
    key: "upstream",
    label: "Group Upstream Operations",
    side: "BusinessSegment",
    termSetGuid: "REPLACE-UPSTREAM-TERMSET-GUID",
    stagingFolder: "Group Upstream Operations",
    levels: [
      { label: "Region", column: "Region" },
      { label: "Estate/Mill", column: "EstateMill" },
    ],
    sortOrder: 5,
  },
  {
    key: "sdgi",
    label: "Group SDGI Operations",
    side: "BusinessSegment",
    termSetGuid: "REPLACE-SDGI-TERMSET-GUID",
    stagingFolder: "Group SDGI Operations",
    levels: [
      { label: "Refinery", column: "Refinery" },
      { label: "Department", column: "Department" },
    ],
    sortOrder: 6,
  },
  {
    key: "it",
    label: "Group Innovation & Technology",
    side: "BusinessSegment",
    termSetGuid: "REPLACE-IT-TERMSET-GUID",
    stagingFolder: "Group Innovation & Technology",
    levels: [{ label: "I&T Operating Unit", column: "ITOperatingUnit" }],
    sortOrder: 7,
  },
  {
    key: "projects",
    label: "Group-led Projects",
    side: "Project",
    termSetGuid: "REPLACE-PROJECTS-TERMSET-GUID",
    stagingFolder: "Group-led Projects",
    levels: [
      // Renamed from "Project Name"/"ProjectName" 2026-07-28: the bare
      // ProjectName column is now the free-text field on every upload.
      { label: "Group Project Name", column: "GroupProjectName" },
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ],
    sortOrder: 8,
  },
];

type DmsSettings = {
  termSets: {
    documentType: string;
    yearPeriod: string;
    confidentiality: string;
    vendor: string;
  };
  // Metadata column INTERNAL names — portable via DMS Config col_* setting rows
  // (defaults match the current site so no config rows are required). Mirrors Form.tsx.
  columns: {
    documentType: string;
    yearPeriod: string;
    documentDate: string;
    confidentiality: string;
    vendor: string;
    businessSegmentLabel: string;
    businessSegmentTid: string;
  };
  stagingLibrary: string;
  allowedExtensions: string[];
};

const DEFAULT_SETTINGS: DmsSettings = {
  termSets: {
    documentType: "866c5754-258e-401f-8685-03d20ae59b1d",
    yearPeriod: "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf",
    confidentiality: "0d6d1da8-27e5-477f-8684-e8cf169f8fb9",
    vendor: "eaafd0e5-03fd-4d33-b1b1-e4252bec430a",
  },
  columns: {
    documentType: FIELDS.documentType,
    yearPeriod: FIELDS.yearPeriod,
    documentDate: FIELDS.documentDate,
    confidentiality: FIELDS.confidentiality,
    vendor: FIELDS.vendor,
    businessSegmentLabel: LEVEL_COLUMNS.BusinessSegment.label,
    businessSegmentTid: LEVEL_COLUMNS.BusinessSegment.tid,
  },
  stagingLibrary: "Staging",
  allowedExtensions: [".pdf", ".xls", ".xlsx"],
};

type XhrResult = { ok: boolean; status: number; body: string };

/**
 * POST a file to SharePoint with real byte-level progress.
 *
 * spHttpClient is fetch-based and exposes no upload progress events, so a
 * genuine percentage was impossible through it — hence the old looping
 * indeterminate bar. XMLHttpRequest still gives us `upload.onprogress`, so the
 * raw REST call is made here instead: same URL and same raw-File body as
 * before, plus the X-RequestDigest header spHttpClient used to attach for us.
 */
const postFileWithProgress = (
  url: string,
  file: File,
  digest: string,
  onPct: (pct: number) => void,
): Promise<XhrResult> =>
  new Promise<XhrResult>((resolve) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", url, true);
    xhr.setRequestHeader("Accept", "application/json;odata=nometadata");
    xhr.setRequestHeader("X-RequestDigest", digest);
    xhr.upload.onprogress = (e: ProgressEvent): void => {
      if (e.lengthComputable && e.total > 0) {
        // Capped at 99 — the last percent stands for "bytes sent, server still
        // committing the file", which only `onload` can confirm.
        onPct(Math.min(99, Math.round((e.loaded / e.total) * 100)));
      }
    };
    xhr.onload = (): void =>
      resolve({
        ok: xhr.status >= 200 && xhr.status < 300,
        status: xhr.status,
        body: xhr.responseText || "",
      });
    xhr.onerror = (): void =>
      resolve({ ok: false, status: 0, body: "Network error during upload." });
    xhr.onabort = (): void =>
      resolve({ ok: false, status: 0, body: "Upload was aborted." });
    xhr.send(file);
  });

// Human-friendly file size for the progress rows.
const fmtSize = (bytes: number): string =>
  bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    : `${Math.max(1, Math.round(bytes / 1024))} KB`;

/**
 * The animated fill for the file currently uploading.
 *
 * `target` is the real bytes-sent percentage, but the browser only fires
 * `upload.onprogress` a handful of times per file — a 2 MB PDF typically reports
 * once or twice, so the raw value parks on something like 50% and then jumps to
 * done. This drives the width from a rAF loop instead:
 *
 *   - when real progress arrives, ease toward it quickly (measured data wins);
 *   - when nothing new has been reported, creep slowly toward the midpoint of
 *     whatever is left, decelerating as it goes.
 *
 * So the bar always looks like it is loading, never overtakes what we actually
 * know, and never reaches 100% until the upload genuinely resolves.
 *
 * The loop writes to the DOM through refs rather than through state — one file
 * uploads at a time, but 60 setState calls a second to re-render a batch list of
 * 50 rows is a waste.
 */
const UploadingBar = ({
  target,
  label,
}: {
  target: number;
  label: string;
}): React.ReactElement => {
  const fillRef = useRef<HTMLSpanElement>(null);
  const pctRef = useRef<HTMLSpanElement>(null);
  const targetRef = useRef<number>(target);

  useEffect(() => {
    targetRef.current = target;
  }, [target]);

  useEffect(() => {
    let frame = 0;
    let shown = 0;
    const tick = (): void => {
      const t = targetRef.current;
      if (t > shown) {
        shown += (t - shown) * 0.15; // catch up to real progress
      } else {
        const ceiling = t + (99 - t) * 0.5; // halfway into the unknown, no further
        if (shown < ceiling) shown += (ceiling - shown) * 0.006;
      }
      if (shown > 99) shown = 99;
      if (fillRef.current) fillRef.current.style.width = `${shown}%`;
      if (pctRef.current) pctRef.current.textContent = `${Math.round(shown)}%`;
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(frame);
  }, []);

  return (
    <>
      <span
        className="dms-fp-bar"
        role="progressbar"
        aria-valuenow={target}
        aria-valuemin={0}
        aria-valuemax={100}
        aria-label={label}
      >
        <span className="dms-fp-fill" ref={fillRef} style={{ width: "0%" }} />
      </span>
      <span className="dms-fp-pct" ref={pctRef}>
        0%
      </span>
    </>
  );
};

export default function BulkUpload({
  context,
}: IBulkUploadProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const webSru = context.pageContext.web.serverRelativeUrl.replace(/\/+$/, "");
  const fileRef = useRef<HTMLInputElement>(null);
  const batchSeqRef = useRef<number>(1);
  // Cached form digest for the XHR upload path (spHttpClient handles its own).
  // Digests expire — SharePoint tells us when, and we refresh a minute early.
  const digestRef = useRef<{ value: string; expiresAt: number } | null>(null);

  const [options, setOptions] = useState<OptionMap>(EMPTY_OPTIONS);
  const [deptLoading, setDeptLoading] = useState<boolean>(true);
  // Privileged = site admin OR a GLOBAL-role uploader: bypasses tier detection
  // and gets the full manual cascade (may upload anywhere).
  const [privileged, setPrivileged] = useState<boolean>(false);

  // Generic N-level cascade state for the DRAFT config panel: one option list +
  // one selected term id per level.
  const [levelChoices, setLevelChoices] = useState<TermOption[][]>([]);
  const [levelValues, setLevelValues] = useState<string[]>([]);
  const [validPaths, setValidPaths] = useState<ValidPath[]>([]);

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);

  // ── Queued batches ──
  const [batches, setBatches] = useState<Batch[]>([]);
  // The inline config panel: open while adding/editing a batch. editingId is the
  // id of the batch being edited (null when adding a new one).
  const [panelOpen, setPanelOpen] = useState<boolean>(false);
  const [editingId, setEditingId] = useState<string | null>(null);

  // ── Draft state (bound to the config panel) ──
  const [uploadMode, setUploadMode] = useState<string>("");
  const [picked, setPicked] = useState<PickedFile[]>([]);
  const [documentType, setDocumentType] = useState<string>("");
  const [yearPeriod, setYearPeriod] = useState<string>("");
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [vendor, setVendor] = useState<string>("");

  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [batchResults, setBatchResults] = useState<BatchOutcome[] | null>(null);
  const [live, setLive] = useState<LiveBatch[] | null>(null);
  const [toast, setToast] = useState<{
    message: string;
    type: ToastType;
  } | null>(null);
  const toastTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const showToast = (message: string, type: ToastType): void => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToast({ message, type });
    toastTimerRef.current = setTimeout(() => setToast(null), 5000);
  };

  // Replace-file guard. Uploads run one file at a time, so a single shared prompt
  // (resolved via a promise) is enough — the loop awaits the user's answer.
  const [replaceAsk, setReplaceAsk] = useState<{ name: string } | null>(null);
  const replaceResolveRef = useRef<((ok: boolean) => void) | null>(null);
  const askReplace = (name: string): Promise<boolean> =>
    new Promise<boolean>((resolve) => {
      replaceResolveRef.current = resolve;
      setReplaceAsk({ name });
    });
  const answerReplace = (ok: boolean): void => {
    setReplaceAsk(null);
    const resolve = replaceResolveRef.current;
    replaceResolveRef.current = null;
    if (resolve) resolve(ok);
  };

  const activeMode = (): UploadMode | undefined =>
    modes.find((m) => m.key === uploadMode);

  /* ---------- Term Store helpers ------------------------------------------ */

  const loadTermSet = async (termSetId: string): Promise<TermOption[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error(`Term set ${termSetId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map(
      (t: { id: string; labels: Array<{ name: string }> }) => ({
        id: t.id,
        label: t.labels[0].name,
      }),
    );
  };

  const loadTermChildren = async (
    termSetId: string,
    termId: string,
  ): Promise<TermOption[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${termId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok)
      throw new Error(`Term ${termId} children returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map(
      (t: { id: string; labels: Array<{ name: string }> }) => ({
        id: t.id,
        label: t.labels[0].name,
      }),
    );
  };

  // Resolve the full ancestor chain [top ... leaf] for a unit term, as {label,id}.
  const loadTermPath = async (
    termSetId: string,
    leafTermId: string,
  ): Promise<TermOption[]> => {
    const chain: TermOption[] = [];
    let currentId: string | null = leafTermId;
    while (currentId) {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${currentId}?$select=id,labels&$expand=parent($select=id)`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) break;
      const t = await res.json();
      chain.unshift({ id: t.id, label: t.labels?.[0]?.name ?? "" });
      currentId = t.parent?.id ?? null;
    }
    return chain;
  };

  /* ---------- Config + group-map readers ---------------------------------- */

  const loadModes = async (): Promise<UploadMode[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,ModeLabel,Category,TermSetGuid,StagingFolder,Levels,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error("DMS Config list not found");
    const data = await res.json();
    return (data.value ?? []).map(
      (item: {
        Title: string;
        ModeLabel: string;
        Category: string;
        TermSetGuid: string;
        StagingFolder: string;
        Levels: string;
        SortOrder: number;
      }) => ({
        key: item.Title,
        label: item.ModeLabel,
        side: (item.Category === "Project" ? "Project" : "BusinessSegment") as
          | "BusinessSegment"
          | "Project",
        termSetGuid: item.TermSetGuid,
        stagingFolder: item.StagingFolder,
        levels: parseLevels(item.Levels),
        sortOrder: item.SortOrder,
      }),
    );
  };

  const loadGroupMap = async (): Promise<GroupMapRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Group%20Map')/items?$select=GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error("DMS Group Map list not found");
    const data = await res.json();
    return (data.value ?? []).map(
      (r: {
        GroupId: string;
        GroupName: string;
        Segment: string;
        UnitTermGuid: string;
        Role: string;
      }) => ({
        groupId: r.GroupId ?? "",
        groupName: r.GroupName ?? "",
        segment: r.Segment ?? "",
        termGuid: r.UnitTermGuid ?? "",
        role: r.Role ?? "",
      }),
    );
  };

  const loadSettings = async (): Promise<DmsSettings> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error("DMS Config list not found");
    const data = await res.json();
    const map: Record<string, string> = {};
    (data.value ?? []).forEach(
      (item: { Title: string; SettingValue: string }) => {
        map[item.Title] = item.SettingValue;
      },
    );
    const get = (key: string): string | undefined => map[key];
    return {
      termSets: {
        documentType:
          get("termSet_documentType") ?? DEFAULT_SETTINGS.termSets.documentType,
        yearPeriod:
          get("termSet_yearPeriod") ?? DEFAULT_SETTINGS.termSets.yearPeriod,
        confidentiality:
          get("termSet_confidentiality") ??
          DEFAULT_SETTINGS.termSets.confidentiality,
        vendor: get("termSet_vendor") ?? DEFAULT_SETTINGS.termSets.vendor,
      },
      columns: {
        documentType: get("col_documentType") ?? DEFAULT_SETTINGS.columns.documentType,
        yearPeriod: get("col_yearPeriod") ?? DEFAULT_SETTINGS.columns.yearPeriod,
        documentDate: get("col_documentDate") ?? DEFAULT_SETTINGS.columns.documentDate,
        confidentiality:
          get("col_confidentiality") ?? DEFAULT_SETTINGS.columns.confidentiality,
        vendor: get("col_vendor") ?? DEFAULT_SETTINGS.columns.vendor,
        businessSegmentLabel:
          get("col_businessSegment") ?? DEFAULT_SETTINGS.columns.businessSegmentLabel,
        businessSegmentTid:
          get("col_businessSegmentTid") ?? DEFAULT_SETTINGS.columns.businessSegmentTid,
      },
      stagingLibrary: get("stagingLibrary") ?? DEFAULT_SETTINGS.stagingLibrary,
      allowedExtensions: get("allowedExtensions")
        ? (get("allowedExtensions") as string).split(",").map((e) => e.trim())
        : DEFAULT_SETTINGS.allowedExtensions,
    };
  };

  /* ---------- Identity: group ids + admin --------------------------------- */

  // Native SP site-group Ids as strings (matches DMS Group Map GroupId, e.g. "27").
  const loadUserGroupIds = async (): Promise<string[]> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser/groups?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return [];
      const d = await res.json();
      return ((d.value ?? []) as Array<{ Id?: number }>)
        .map((g) => (g.Id != null ? String(g.Id) : ""))
        .filter(Boolean);
    } catch {
      return [];
    }
  };

  const loadIsAdmin = async (): Promise<boolean> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return false;
      const d = await res.json();
      return d.IsSiteAdmin === true;
    } catch {
      return false;
    }
  };

  /* ---------- Cascade builders (draft panel) ------------------------------ */

  const initCascade = async (mode: UploadMode): Promise<void> => {
    const tops = await loadTermSet(mode.termSetGuid).catch(
      () => [] as TermOption[],
    );
    setLevelChoices([tops]);
    setLevelValues([]);
  };

  // Rebuild the cascade option lists for a set of already-chosen level values
  // (used when editing a saved batch — privileged users).
  const restoreCascade = async (
    mode: UploadMode,
    values: string[],
  ): Promise<void> => {
    const tops = await loadTermSet(mode.termSetGuid).catch(
      () => [] as TermOption[],
    );
    const choices: TermOption[][] = [tops];
    for (let i = 0; i < values.length - 1; i++) {
      if (!values[i]) break;
      const kids = await loadTermChildren(mode.termSetGuid, values[i]).catch(
        () => [] as TermOption[],
      );
      choices[i + 1] = kids;
    }
    setLevelChoices(choices);
    setLevelValues(values);
  };

  const onLevelChange = async (
    mode: UploadMode,
    idx: number,
    termId: string,
  ): Promise<void> => {
    const values = levelValues.slice(0, idx);
    values[idx] = termId;
    setLevelValues(values);
    const choices = levelChoices.slice(0, idx + 1);
    if (idx + 1 < mode.levels.length && termId) {
      const kids = await loadTermChildren(mode.termSetGuid, termId).catch(
        () => [] as TermOption[],
      );
      choices[idx + 1] = kids;
    }
    setLevelChoices(choices);
  };

  /* ---------- Restricted (tiered-detection) cascade ----------------------- */

  const resolveValidPaths = async (
    loadedModes: UploadMode[],
    membership: Membership,
  ): Promise<ValidPath[]> => {
    const out: ValidPath[] = [];
    for (const leaf of membership.uploaderLeaves) {
      const mode = loadedModes.find((m) => m.termSetGuid === leaf.segment);
      if (!mode) continue;
      const chain = await loadTermPath(mode.termSetGuid, leaf.termGuid).catch(
        () => [] as TermOption[],
      );
      const requireSegment = mode.side === "BusinessSegment";
      if (
        isChainAuthorized(
          chain.map((c) => c.id),
          mode.termSetGuid,
          membership.memberTerms,
          requireSegment,
        )
      ) {
        out.push({ modeKey: mode.key, chain });
      }
    }
    return out;
  };

  const restrictedCascade = (
    paths: ValidPath[],
    mode: UploadMode,
    current: string[],
  ): { choices: TermOption[][]; values: string[] } => {
    const chains = paths
      .filter((p) => p.modeKey === mode.key)
      .map((p) => p.chain);
    const choices: TermOption[][] = [];
    const values: string[] = [];
    for (let i = 0; i < mode.levels.length; i++) {
      const consistent = chains.filter((ch) =>
        values.slice(0, i).every((v, j) => ch[j]?.id === v),
      );
      const seen = new Set<string>();
      const opts: TermOption[] = [];
      for (const ch of consistent) {
        const t = ch[i];
        if (t && !seen.has(t.id)) {
          seen.add(t.id);
          opts.push(t);
        }
      }
      choices[i] = opts;
      if (opts.length === 1) values[i] = opts[0].id; // auto-lock
      else {
        const prior = current[i] ?? "";
        values[i] = opts.some((o) => o.id === prior) ? prior : "";
      }
    }
    return { choices, values };
  };

  const applyRestrictedMode = (
    paths: ValidPath[],
    mode: UploadMode,
    current: string[],
  ): void => {
    const { choices, values } = restrictedCascade(paths, mode, current);
    setLevelChoices(choices);
    setLevelValues(values);
  };

  const isLevelLocked = (i: number): boolean =>
    !privileged && (levelChoices[i]?.length ?? 0) <= 1;

  /* ---------- Init -------------------------------------------------------- */

  useEffect(() => {
    const init = async (): Promise<void> => {
      const [loadedSettings, rawModes, groupMap, userGroupIds, admin] =
        await Promise.all([
          loadSettings().catch(() => DEFAULT_SETTINGS),
          loadModes().catch(() => DEFAULT_MODES),
          loadGroupMap().catch(() => [] as GroupMapRow[]),
          loadUserGroupIds(),
          loadIsAdmin(),
        ]);
      setSettings(loadedSettings);
      const usable = rawModes.filter(
        (m: UploadMode) => m.levels.length > 0 && !!m.termSetGuid,
      );
      const loadedModes = usable.length > 0 ? usable : DEFAULT_MODES;
      setModes(loadedModes);

      const membership = collectMembership(groupMap, userGroupIds);
      // Upload-anywhere is a site-admin privilege only. GLOBAL is a read-only role
      // (site-entry-access-layer spec) and no longer grants upload.
      const isPrivileged = admin;
      setPrivileged(isPrivileged);

      const [docTypes, years, confs, vendors] = await Promise.all([
        loadTermSet(loadedSettings.termSets.documentType).catch(
          () => [] as TermOption[],
        ),
        loadTermSet(loadedSettings.termSets.yearPeriod).catch(
          () => [] as TermOption[],
        ),
        loadTermSet(loadedSettings.termSets.confidentiality).catch(
          () => [] as TermOption[],
        ),
        loadTermSet(loadedSettings.termSets.vendor).catch(
          () => [] as TermOption[],
        ),
      ]);
      setOptions({
        documentType: docTypes,
        yearPeriod: years,
        confidentiality: confs,
        vendor: vendors,
      });

      if (!isPrivileged) {
        const paths = await resolveValidPaths(loadedModes, membership);
        setValidPaths(paths);
      }
      setDeptLoading(false);
    };

    init().catch((err) => {
      console.error("BulkUpload init failed:", err);
      showToast("Could not load form data. Please refresh the page.", "error");
      setDeptLoading(false);
    });
  }, []);

  /* ---------- Draft-panel helpers ----------------------------------------- */

  // Default mode to open the config panel on, honouring the user's access.
  const pickDefaultMode = (): UploadMode | undefined => {
    if (privileged) {
      const bs = modes.filter((m) => m.side === "BusinessSegment");
      return bs[0] ?? modes[0];
    }
    const offerable = new Set(validPaths.map((p) => p.modeKey));
    return (
      modes.find((m) => m.side === "BusinessSegment" && offerable.has(m.key)) ??
      modes.find((m) => offerable.has(m.key))
    );
  };

  const resetDraft = (): void => {
    setPicked([]);
    setDocumentType("");
    setLevelValues([]);
    setLevelChoices([]);
    setYearPeriod("");
    setDocumentDate("");
    setConfidentiality("");
    setVendor("");
    if (fileRef.current) fileRef.current.value = "";
  };

  const initDraftCascade = (mode: UploadMode): void => {
    if (privileged) {
      initCascade(mode).catch(() => {
        setLevelChoices([]);
        setLevelValues([]);
      });
    } else {
      applyRestrictedMode(validPaths, mode, []);
    }
  };

  const openAddPanel = (): void => {
    if (batches.length >= MAX_BATCHES) {
      showToast(`You can queue at most ${MAX_BATCHES} batches.`, "error");
      return;
    }
    resetDraft();
    setEditingId(null);
    setPanelOpen(true);
    setLive(null);
    setBatchResults(null);
    const dm = pickDefaultMode();
    if (dm) {
      setUploadMode(dm.key);
      initDraftCascade(dm);
    }
  };

  const editBatch = (batch: Batch): void => {
    // The tile is hidden while editingId === batch.id (see render); Save replaces
    // it in place, Cancel brings it back.
    setEditingId(batch.id);
    setPanelOpen(true);
    setUploadMode(batch.modeKey);
    const values = batch.levelSelections.map((s) => s.id);
    const mode = modes.find((m) => m.key === batch.modeKey);
    if (mode) {
      if (privileged) {
        restoreCascade(mode, values).catch(() => initDraftCascade(mode));
      } else {
        applyRestrictedMode(validPaths, mode, values);
      }
    }
    setPicked(
      batch.files.map((file, i) => ({
        key: `${batch.id}-${i}-${file.name}`,
        file,
      })),
    );
    setDocumentType(batch.docTypeId);
    setYearPeriod(batch.yearId);
    setDocumentDate(batch.documentDate);
    setConfidentiality(batch.confId);
    setVendor(batch.vendorId);
    setBatchResults(null);
    setLive(null);
  };

  const removeBatch = (id: string): void => {
    setBatches((prev) => prev.filter((b) => b.id !== id));
    setLive(null);
    setBatchResults(null);
  };

  const cancelPanel = (): void => {
    setPanelOpen(false);
    setEditingId(null);
    resetDraft();
  };

  const switchMode = (modeKey: string): void => {
    setUploadMode(modeKey);
    const mode = modes.find((m) => m.key === modeKey);
    if (!mode) return;
    initDraftCascade(mode);
  };

  const handleLevelChange = (
    mode: UploadMode,
    idx: number,
    termId: string,
  ): void => {
    if (privileged) {
      onLevelChange(mode, idx, termId).catch(() => undefined);
      return;
    }
    const values = levelValues.slice(0, idx);
    values[idx] = termId;
    applyRestrictedMode(validPaths, mode, values);
  };

  const validateDraft = (): string[] => {
    const missing: string[] = [];
    if (picked.length === 0) missing.push("File");
    const m = activeMode();
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) missing.push(lvl.label);
    });
    if (!documentType) missing.push("Document Type");
    if (!yearPeriod) missing.push("Year / Period");
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidentiality Level");
    return missing;
  };

  const saveBatch = (): void => {
    const missing = validateDraft();
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }

    // Files upload under their original names — duplicate names within the batch
    // would silently collide (first wins), so reject them here.
    const names = picked.map((p) => p.file.name);
    const collision = names.find(
      (n, i) =>
        names.findIndex((o) => o.toLowerCase() === n.toLowerCase()) !== i,
    );
    if (collision) {
      showToast(
        `Two or more files in this batch are named "${collision}". Remove the duplicate first.`,
        "error",
      );
      return;
    }

    const mode = activeMode();
    if (!mode || mode.levels.length === 0) {
      showToast("No upload mode configured.", "error");
      return;
    }

    const levelSelections: BatchSelection[] = mode.levels.map((lvl, i) => {
      const opt = (levelChoices[i] ?? []).find((o) => o.id === levelValues[i]);
      return {
        column: lvl.column,
        label: opt?.label ?? "",
        id: opt?.id ?? "",
        labelCol: lvl.labelCol,
        tidCol: lvl.tidCol,
      };
    });

    const dt = options.documentType.find((o) => o.id === documentType);
    const yr = options.yearPeriod.find((o) => o.id === yearPeriod);
    const cf = options.confidentiality.find((o) => o.id === confidentiality);
    const vd = options.vendor.find((o) => o.id === vendor);

    const destinationLabel = [
      mode.label,
      ...levelSelections.map((s) => s.label),
    ]
      .filter(Boolean)
      .join(" › ");

    const batch: Batch = {
      id: editingId ?? `b${batchSeqRef.current++}`,
      files: picked.map((p) => p.file),
      modeKey: mode.key,
      modeLabel: mode.label,
      termSetGuid: mode.termSetGuid,
      levelSelections,
      docTypeId: documentType,
      docTypeLabel: dt?.label ?? "",
      yearId: yearPeriod,
      yearLabel: yr?.label ?? "",
      confId: confidentiality,
      confLabel: cf?.label ?? "",
      vendorId: vendor,
      vendorLabel: vd?.label ?? "",
      documentDate,
      destinationLabel,
    };

    setBatches((prev) =>
      editingId
        ? prev.map((b) => (b.id === editingId ? batch : b))
        : [...prev, batch],
    );
    setPanelOpen(false);
    setEditingId(null);
    resetDraft();
  };

  const clearAll = (): void => {
    setBatches([]);
    setBatchResults(null);
    setLive(null);
    cancelPanel();
  };

  /* ---------- File selection (draft panel) -------------------------------- */

  const addFiles = (list: FileList | null): void => {
    if (!list || list.length === 0) return;
    const incoming = Array.from(list);
    const allowed: File[] = [];
    let rejectedCount = 0;
    incoming.forEach((f) => {
      const ok = settings.allowedExtensions.some((ext) =>
        f.name.toLowerCase().endsWith(ext),
      );
      if (ok) allowed.push(f);
      else rejectedCount++;
    });

    if (rejectedCount > 0) {
      showToast(
        `Skipped ${rejectedCount} file(s) with a disallowed type. Allowed: ${settings.allowedExtensions.join(", ")}`,
        "error",
      );
    }

    setPicked((prev) => {
      const room = MAX_FILES - prev.length;
      if (room <= 0) {
        showToast(`A batch can hold at most ${MAX_FILES} files.`, "error");
        return prev;
      }
      if (allowed.length > room) {
        showToast(
          `Only the first ${room} file(s) were added — the per-batch limit is ${MAX_FILES}.`,
          "error",
        );
      }
      const seq = batchSeqRef.current;
      const added = allowed.slice(0, room).map((file, i) => ({
        key: `${seq}-${prev.length + i}-${file.name}`,
        file,
      }));
      return [...prev, ...added];
    });

    // Clear the input so the same file can be re-picked after a removal.
    if (fileRef.current) fileRef.current.value = "";
  };

  const removeFile = (key: string): void => {
    setPicked((prev) => prev.filter((p) => p.key !== key));
  };

  // Delete an already-uploaded file from the Documents library via the per-row X
  // in the live progress panel (bulk upload writes straight to Documents). The
  // row keeps its `sru` (ServerRelativeUrl) captured at upload time, so the
  // delete targets the exact file regardless of any folder rename. Marks the row
  // "deleted" on success.
  const [deletingKey, setDeletingKey] = useState<string | null>(null);
  const deleteUploaded = async (
    batchIdx: number,
    fileIdx: number,
    sru: string,
    name: string,
  ): Promise<void> => {
    const key = `${batchIdx}:${fileIdx}`;
    if (!window.confirm(`Delete "${name}" from Documents? This cannot be undone.`)) {
      return;
    }
    setDeletingKey(key);
    try {
      const safeUrl = sru.replace(/'/g, "''");
      const delRes: SPHttpClientResponse = await context.spHttpClient.fetch(
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)?@f='${encodeURIComponent(safeUrl)}'`,
        SPHttpClient.configurations.v1,
        { method: "POST", headers: { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" } },
      );
      if (!delRes.ok && delRes.status !== 404) {
        throw new Error(`HTTP ${delRes.status}`);
      }
      setLive((prev) =>
        prev
          ? prev.map((lb, bi) =>
              bi === batchIdx
                ? {
                    ...lb,
                    files: lb.files.map((f, fi) =>
                      fi === fileIdx
                        ? { ...f, state: "deleted" as FileState, sru: undefined }
                        : f,
                    ),
                  }
                : lb,
            )
          : prev,
      );
      showToast(`"${name}" deleted from Documents.`, "success");
    } catch (err) {
      console.error("Delete from Staging failed:", name, err);
      showToast(`Could not delete "${name}". Please try again.`, "error");
    } finally {
      setDeletingKey(null);
    }
  };

  /* ---------- Documents-library path swap --------------------------------- */

  // The DMS Folder Map stores the STAGING unit folder's UniqueId. The Documents
  // library holds a mirrored tree, so the Documents unit folder is found by
  // swapping the library segment of the resolved Staging path:
  //   /sites/<web>/Staging/<rest>  ->  /sites/<web>/Shared Documents/<rest>
  // Anchored on the web-relative prefix rather than a global replace, so a
  // folder that happens to be named "Staging" deeper in the tree is not mangled.
  const toDocumentsPath = (stagingSru: string): string | null => {
    const prefix = `${webSru}/${settings.stagingLibrary}/`;
    if (stagingSru.toLowerCase().indexOf(prefix.toLowerCase()) !== 0) {
      return null;
    }
    const rest = stagingSru.slice(prefix.length);
    return `${webSru}/${DOCUMENTS_URL_SEGMENT}/${rest}`;
  };

  /* ---------- Form digest (XHR upload path) -------------------------------- */

  // A long batch can outlive one digest, so cache it with its real expiry and
  // re-request when close. `force` refetches unconditionally — used to retry once
  // after a 403, which is what an expired/rejected digest looks like.
  const fetchDigest = async (): Promise<string> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) {
      throw new Error(`Could not get a form digest (HTTP ${res.status}).`);
    }
    const json = await res.json();
    // nometadata puts it at the root; verbose nests it under GetContextWebInformation.
    const info =
      json.GetContextWebInformation ?? json.d?.GetContextWebInformation ?? json;
    const value: string = info.FormDigestValue;
    if (!value) throw new Error("Form digest response had no FormDigestValue.");
    const ttlSeconds: number = info.FormDigestTimeoutSeconds ?? 1800;
    digestRef.current = { value, expiresAt: Date.now() + ttlSeconds * 1000 };
    return value;
  };

  const getDigest = async (force?: boolean): Promise<string> => {
    const cached = digestRef.current;
    if (!force && cached && cached.expiresAt > Date.now() + 60_000) {
      return cached.value;
    }
    return fetchDigest();
  };

  /* ---------- Upload one batch -------------------------------------------- */

  // Resolves the batch's destination folder, ensures Year/DocType subfolders, and
  // uploads + tags each file. Never throws for routing problems — returns a
  // batchError instead, so the caller can continue to the next batch.
  const uploadBatch = async (
    batch: Batch,
    index: number,
    total: number,
    onState: (fileIndex: number, state: FileState, sru?: string) => void,
    onPct: (fileIndex: number, pct: number) => void,
  ): Promise<{ batchError?: string; results: FileResult[] }> => {
    const tag = `Batch ${index + 1} of ${total}`;
    const leaf = batch.levelSelections[batch.levelSelections.length - 1];
    if (!leaf || !leaf.id) {
      return { batchError: "No destination folder selected.", results: [] };
    }

    setStatus(`${tag}: locating destination folder…`);
    const mapping = await lookupFolderMapping(
      context.spHttpClient,
      siteUrl,
      leaf.id,
    ).catch((e: unknown) => {
      console.error("Folder map lookup error:", e);
      return null;
    });
    if (!mapping || !mapping.folderUniqueId) {
      return {
        batchError: `This folder hasn't been mapped yet. Ask an administrator to run the reconciliation tool. (term ${leaf.label})`,
        results: [],
      };
    }

    // Resolve the Staging unit folder's CURRENT path (rename-proof), then swap
    // the library segment to reach the mirrored Documents unit folder.
    const stagingSru = await resolveFolderServerUrl(
      context.spHttpClient,
      siteUrl,
      mapping.folderUniqueId,
    );
    if (!stagingSru) {
      return {
        batchError:
          "The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation.",
        results: [],
      };
    }

    const docsUnitPath = toDocumentsPath(stagingSru);
    if (!docsUnitPath) {
      console.error(
        "Could not swap library segment. stagingSru:",
        stagingSru,
        "expected prefix:",
        `${webSru}/${settings.stagingLibrary}/`,
      );
      return {
        batchError:
          "Could not work out the Documents path for this unit folder. Check the stagingLibrary setting in DMS Config.",
        results: [],
      };
    }

    // The unit folder must ALREADY exist in Documents. It is deliberately not
    // auto-created HERE: a folder created on the upload path would inherit the
    // Documents root ACL and silently widen access. The correct remedy is the
    // Reconciliation tool, which creates the mirrored tree in BOTH libraries and
    // locks each folder to its DMS Group Map groups (FolderManager.tsx — it loops
    // over ["Staging", "Documents"]). Hand-creating the folder is NOT equivalent:
    // it would inherit the root ACL, which is exactly what this guard prevents.
    const docsProbe = await probeFolderByPath(
      context.spHttpClient,
      siteUrl,
      docsUnitPath,
    );
    if (!docsProbe.folder) {
      // Only claim the folder is missing when SharePoint actually said 404.
      // Anything else (throttle, permission, malformed request) is "couldn't
      // tell" — and sending the user off to create a folder that already exists
      // is how the first report of this bug went wrong.
      return {
        batchError: docsProbe.confirmedMissing
          ? `The matching folder does not exist in Documents yet (${docsUnitPath}). Ask an administrator to run the reconciliation tool — it creates this folder with the correct permissions. Do not create it by hand: a hand-made folder inherits the library's root permissions and would widen access.`
          : `Could not verify the destination folder in Documents (HTTP ${docsProbe.status}). This usually means SharePoint was busy — most often because folder reconciliation is running at the same time. Wait for reconciliation to finish, then try again. The folder itself is probably fine.`,
        results: [],
      };
    }
    const docsUnitFolder = docsProbe.folder;

    const yearLabel = sanitizeFolderSegment(batch.yearLabel);
    const docTypeLabel = sanitizeFolderSegment(batch.docTypeLabel);
    if (!yearLabel || !docTypeLabel) {
      return {
        batchError: "Year and Document Type are required.",
        results: [],
      };
    }

    setStatus(`${tag}: preparing destination folders…`);
    // Year / Document Type subfolders are safe to create — they inherit the unit
    // folder's ACL, exactly as they do in Staging.
    const yearFolder = await ensureFolder(
      context.spHttpClient,
      siteUrl,
      docsUnitFolder.serverRelativeUrl,
      yearLabel,
    );
    if (!yearFolder) {
      return {
        batchError: `Could not create the "${yearLabel}" folder.`,
        results: [],
      };
    }
    const destFolder = await ensureFolder(
      context.spHttpClient,
      siteUrl,
      yearFolder.serverRelativeUrl,
      docTypeLabel,
    );
    if (!destFolder) {
      return {
        batchError: `Could not create the "${docTypeLabel}" folder.`,
        results: [],
      };
    }
    const folderId = destFolder.uniqueId;

    // Batch metadata — identical for every file in this batch.
    const selections: BatchSelection[] = batch.levelSelections.map((s) => ({
      ...s,
    }));
    selections.unshift({
      column: "BusinessSegment",
      label: batch.modeLabel,
      id: batch.termSetGuid,
      labelCol: undefined,
      tidCol: undefined,
    });

    // BusinessSegment column names come from settings.columns (portable); Department/Unit
    // fall back to LEVEL_COLUMNS but are normally overridden by DMS Config Levels JSON.
    const levelCols: Record<string, ColumnPair> = {
      ...LEVEL_COLUMNS,
      BusinessSegment: {
        label: settings.columns.businessSegmentLabel,
        tid: settings.columns.businessSegmentTid,
      },
    };
    const formValues: Array<{ FieldName: string; FieldValue: string }> = [
      {
        FieldName: settings.columns.documentType,
        FieldValue: taxVal(batch.docTypeLabel, batch.docTypeId),
      },
      {
        FieldName: settings.columns.yearPeriod,
        FieldValue: taxVal(batch.yearLabel, batch.yearId),
      },
      {
        FieldName: settings.columns.confidentiality,
        FieldValue: taxVal(batch.confLabel, batch.confId),
      },
      {
        FieldName: settings.columns.documentDate,
        FieldValue: toSpDate(batch.documentDate),
      },
      ...buildLevelFormValues(levelCols, selections),
    ];
    // Vendor is intentionally NOT written. The field is hidden in the UI (client
    // request, 2026-07-28) so batch.vendorId is always empty; the plumbing stays
    // in place so un-hiding the select is the only change needed to restore it.
    if (batch.vendorId) {
      formValues.push({
        FieldName: settings.columns.vendor,
        FieldValue: taxVal(batch.vendorLabel, batch.vendorId),
      });
    }

    /* ----- Sequential per-file upload; one failure never stops the batch -- */
    const results: FileResult[] = [];

    for (let i = 0; i < batch.files.length; i++) {
      const file = batch.files[i];
      const finalName = file.name;
      onPct(i, 0);
      onState(i, "uploading");
      setStatus(
        `${tag}: uploading ${i + 1} of ${batch.files.length} — ${finalName}`,
      );

      // Duplicate probe. On a clash we ask the user whether to overwrite; No
      // keeps the existing skip behaviour, Yes re-uploads with overwrite=true.
      let overwrite = false;
      try {
        const existsRes: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (existsRes.ok) {
          const confirmed = await askReplace(finalName);
          if (!confirmed) {
            results.push({
              name: finalName,
              outcome: "skipped",
              detail: "A file with this name already exists here.",
            });
            onState(i, "skipped");
            continue;
          }
          overwrite = true;
          onPct(i, 0);
          onState(i, "uploading");
        }
      } catch {
        // Network error on the existence check — proceed; the upload will
        // surface the real error.
      }

      try {
        const addUrl = `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=${overwrite})?$select=ServerRelativeUrl`;

        let uploadRes = await postFileWithProgress(
          addUrl,
          file,
          await getDigest(),
          (pct) => onPct(i, pct),
        );
        // A 403 here is almost always a stale digest — refresh once and retry.
        if (!uploadRes.ok && uploadRes.status === 403) {
          onPct(i, 0);
          uploadRes = await postFileWithProgress(
            addUrl,
            file,
            await getDigest(true),
            (pct) => onPct(i, pct),
          );
        }

        if (!uploadRes.ok) {
          let detail = `HTTP ${uploadRes.status}`;
          const bodyText = uploadRes.body;
          try {
            const errJson = JSON.parse(bodyText);
            const spMsg =
              errJson?.error?.message?.value ?? errJson?.error?.message;
            detail += spMsg ? ` — ${spMsg}` : ` — ${bodyText.slice(0, 200)}`;
          } catch {
            if (bodyText) detail += ` — ${bodyText.slice(0, 200)}`;
          }
          console.error("Upload failed:", finalName, folderId, detail);
          results.push({ name: finalName, outcome: "failed", detail });
          onState(i, "failed");
          continue;
        }
        const uploadJson = JSON.parse(uploadRes.body);
        const uploadedSru = uploadJson.ServerRelativeUrl;
        onPct(i, 100);

        const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeURIComponent(uploadedSru)}'`,
          SPHttpClient.configurations.v1,
        );
        if (!itemRes.ok) {
          results.push({
            name: finalName,
            outcome: "tagFailed",
            detail: "Uploaded, but the item could not be retrieved to tag.",
          });
          onState(i, "tagFailed", uploadedSru);
          continue;
        }
        const item = await itemRes.json();

        const metaRes: SPHttpClientResponse = await context.spHttpClient.post(
          `${siteUrl}/_api/web/lists/getbytitle('${DOCUMENTS_LIST_TITLE}')/items(${item.Id})/validateUpdateListItem`,
          SPHttpClient.configurations.v1,
          {
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ formValues }),
          },
        );
        if (!metaRes.ok) {
          results.push({
            name: finalName,
            outcome: "tagFailed",
            detail: `Uploaded, but tagging failed (HTTP ${metaRes.status}).`,
          });
          onState(i, "tagFailed", uploadedSru);
          continue;
        }
        // validateUpdateListItem returns HTTP 200 even on field errors —
        // HasException on each result is the real check.
        const metaJson = await metaRes.json();
        const fieldError = (metaJson.value ?? []).find(
          (v: { HasException?: boolean }) => v.HasException,
        );
        if (fieldError) {
          console.error("Field update error:", finalName, fieldError);
          results.push({
            name: finalName,
            outcome: "tagFailed",
            detail: `Uploaded, but a field failed: ${fieldError.FieldName} — ${fieldError.ErrorMessage}`,
          });
          onState(i, "tagFailed", uploadedSru);
          continue;
        }

        results.push({ name: finalName, outcome: "uploaded" });
        onState(i, "done", uploadedSru);
      } catch (err) {
        console.error("Upload threw:", finalName, err);
        results.push({
          name: finalName,
          outcome: "failed",
          detail: err instanceof Error ? err.message : "Unexpected error.",
        });
        onState(i, "failed");
      }
    }

    return { results };
  };

  /* ---------- Upload all -------------------------------------------------- */

  const handleUpload = async (): Promise<void> => {
    if (panelOpen) {
      showToast("Save or cancel the batch you're editing first.", "error");
      return;
    }
    if (batches.length === 0) {
      showToast("Add at least one batch first.", "error");
      return;
    }

    setBusy(true);
    setBatchResults(null);
    // Seed the live progress list — every file starts "pending".
    setLive(
      batches.map((b, i) => ({
        label: `Batch ${i + 1}`,
        destinationLabel: b.destinationLabel,
        files: b.files.map((f) => ({
          name: f.name,
          size: f.size,
          state: "pending" as FileState,
          pct: 0,
        })),
      })),
    );
    const outcomes: BatchOutcome[] = [];

    try {
      for (let b = 0; b < batches.length; b++) {
        const batch = batches[b];
        const onState = (
          fileIndex: number,
          state: FileState,
          sru?: string,
        ): void => {
          setLive((prev) =>
            prev
              ? prev.map((lb, idx) =>
                  idx === b
                    ? {
                        ...lb,
                        files: lb.files.map((f, fi) =>
                          fi === fileIndex
                            ? { ...f, state, ...(sru ? { sru } : {}) }
                            : f,
                        ),
                      }
                    : lb,
                )
              : prev,
          );
        };
        const onPct = (fileIndex: number, pct: number): void => {
          setLive((prev) =>
            prev
              ? prev.map((lb, idx) =>
                  idx === b
                    ? {
                        ...lb,
                        files: lb.files.map((f, fi) =>
                          fi === fileIndex ? { ...f, pct } : f,
                        ),
                      }
                    : lb,
                )
              : prev,
          );
        };
        const { batchError, results } = await uploadBatch(
          batch,
          b,
          batches.length,
          onState,
          onPct,
        );
        if (batchError) {
          // Routing failed before any file uploaded — show every row as failed.
          setLive((prev) =>
            prev
              ? prev.map((lb, idx) =>
                  idx === b
                    ? {
                        ...lb,
                        files: lb.files.map((f) => ({
                          ...f,
                          state: "failed" as FileState,
                        })),
                      }
                    : lb,
                )
              : prev,
          );
        }
        outcomes.push({
          batchId: batch.id,
          label: `Batch ${b + 1}`,
          destinationLabel: batch.destinationLabel,
          batchError,
          results,
        });
      }

      setBatchResults(outcomes);
      setStatus("");

      const totalOk = outcomes.reduce(
        (n, o) => n + o.results.filter((r) => r.outcome === "uploaded").length,
        0,
      );
      const anyProblem = outcomes.some(
        (o) => !!o.batchError || o.results.some((r) => r.outcome !== "uploaded"),
      );
      if (!anyProblem && totalOk > 0) {
        showToast(`All ${totalOk} file(s) uploaded to Documents.`, "success");
        // The queue has been fully processed — empty it so "Upload all" can't be
        // clicked a second time and re-send the same files (which would raise a
        // replace prompt per file). The results summary and the progress list are
        // deliberately left on screen, so nothing is lost by dropping the queue.
        // On a PARTIAL failure the batches are kept, so the user can edit and retry.
        setBatches([]);
      }
    } catch (err) {
      console.error("Bulk upload failed:", err);
      showToast("Upload failed. Please try again.", "error");
      setStatus("");
    } finally {
      setBusy(false);
    }
  };

  /* ---------- Render helpers ---------------------------------------------- */

  const renderSelect = (
    label: string,
    required: boolean,
    value: string,
    onChange: (v: string) => void,
    opts: TermOption[],
    disabled = false,
    placeholder = "--",
    wide = false,
  ): React.ReactElement => {
    const selectedLabel = opts.find((o) => o.id === value)?.label ?? "";
    return (
      <label
        className="dms-field"
        style={wide ? { gridColumn: "1 / -1" } : undefined}
      >
        <span>
          {label} {required && <em className="req">*</em>}
        </span>
        <select
          value={value}
          disabled={disabled}
          title={selectedLabel}
          onChange={(e) => onChange(e.target.value)}
        >
          <option value="">{placeholder}</option>
          {opts.map((o) => (
            <option key={o.id} value={o.id} title={o.label}>
              {o.label}
            </option>
          ))}
        </select>
      </label>
    );
  };

  const outcomeLabel: Record<Outcome, string> = {
    uploaded: "Uploaded",
    skipped: "Skipped",
    failed: "Failed",
    tagFailed: "Uploaded, not tagged",
  };

  const fpLabel: Record<FileState, string> = {
    pending: "Waiting",
    uploading: "Uploading",
    done: "Uploaded",
    skipped: "Skipped",
    failed: "Failed",
    tagFailed: "No tags",
    deleted: "Deleted",
  };

  // Overall progress across all live batches. A file counts as "settled" once it
  // reaches any terminal state; the one file in flight contributes its own byte
  // fraction so the bar creeps forward instead of jumping a whole file at a time.
  const liveTotals = (live ?? []).reduce(
    (acc, lb) => {
      lb.files.forEach((f) => {
        acc.total += 1;
        if (f.state !== "pending" && f.state !== "uploading") acc.done += 1;
        else if (f.state === "uploading") acc.partial += f.pct / 100;
      });
      return acc;
    },
    { total: 0, done: 0, partial: 0 },
  );
  const livePct =
    liveTotals.total > 0
      ? Math.min(
          100,
          Math.round(
            ((liveTotals.done + liveTotals.partial) / liveTotals.total) * 100,
          ),
        )
      : 0;

  const totalQueued = batches.reduce((n, b) => n + b.files.length, 0);
  const canAdd = !panelOpen && batches.length < MAX_BATCHES;
  const visibleBatches = batches.filter((b) => b.id !== editingId);

  /* ---------- Render ------------------------------------------------------ */

  return (
    <section className="dms-form">
      <style>{`
        .dms-form { max-width: 960px; margin: 32px auto; padding: 0 24px 48px; font-family: 'Segoe UI', sans-serif; }
        .dms-form h2 { margin: 0 0 4px; font-size: 24px; font-weight: 700; color: #1b1b1b; }
        .dms-subtitle { margin: 0 0 12px; font-size: 14px; color: #666; }
        .dms-warn { display: flex; gap: 10px; align-items: flex-start; background: #fff4e5; border: 1px solid #f0c070; border-radius: 6px; padding: 12px 14px; font-size: 13px; color: #7a4f00; margin: 0 0 24px; }
        .dms-section { background: #fff; border: 1px solid #e1e1e1; border-radius: 8px; padding: 24px; margin-bottom: 20px; }
        .dms-section-title { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #0f6c3f; margin: 0 0 16px; }
        .dms-drop { display: flex; align-items: center; justify-content: space-between; gap: 12px; border: 1px dashed #c8c8c8; border-radius: 6px; padding: 16px; }
        .dms-drop .size { color: #666; font-size: 12px; }
        .dms-link { background: none; border: none; color: #0f6c3f; cursor: pointer; font-weight: 600; padding: 0; font-size: 13px; }
        .dms-link:disabled { color: #9bbfaa; cursor: default; }
        /* Scrolls instead of paginating — the whole selection is always reachable. */
        .dms-filelist { margin-top: 16px; display: flex; flex-direction: column; gap: 10px; max-height: 340px; overflow-y: auto; overscroll-behavior: contain; padding-right: 6px; }
        .dms-filerow { display: grid; grid-template-columns: minmax(0,1fr) auto; gap: 12px; align-items: center; border: 1px solid #ececec; border-radius: 6px; padding: 10px 12px; background: #fafafa; }
        .dms-filerow .orig { font-size: 13px; font-weight: 600; word-break: break-all; }
        .dms-filerow .size { font-size: 12px; color: #666; }
        .dms-remove { background: none; border: none; cursor: pointer; color: #d13438; font-size: 15px; line-height: 1; padding: 4px; }
        .dms-count { font-size: 12px; color: #666; }
        .dms-filelist-foot { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-top: 10px; flex-wrap: wrap; }
        .dms-fileblock { margin-top: 20px; border-top: 1px solid #ececec; padding-top: 16px; }
        .dms-fileblock-title { font-size: 13px; font-weight: 700; color: #1b1b1b; margin: 0 0 12px; }
        /* Live per-file progress */
        .dms-progress { margin-top: 16px; }
        .dms-progress-batch { border: 1px solid #e1e1e1; border-radius: 8px; padding: 14px 16px; margin-bottom: 12px; background: #fff; }
        .dms-progress-head { font-size: 13px; font-weight: 700; color: #0f6c3f; margin-bottom: 12px; }
        .dms-progress-list { max-height: 360px; overflow-y: auto; overscroll-behavior: contain; display: flex; flex-direction: column; gap: 6px; padding-right: 6px; }
        /* One row per file: status tag · name · size · live bar · % · remove */
        .dms-fp-row { display: flex; align-items: center; gap: 12px; font-size: 13px; padding: 10px 12px; border-radius: 6px; background: #fafafa; }
        .dms-fp-row.done { background: #eaf7f0; }
        .dms-fp-row.skipped, .dms-fp-row.tagFailed { background: #fff4e5; }
        .dms-fp-row.failed { background: #fdf3f3; }
        .dms-fp-row.deleted { background: #f2f2f2; }
        .dms-fp-tag { flex: 0 0 84px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; color: #999; }
        .dms-fp-tag.done, .dms-fp-tag.uploading { color: #0f6c3f; }
        .dms-fp-tag.skipped, .dms-fp-tag.tagFailed { color: #7a4f00; }
        .dms-fp-tag.failed { color: #d13438; }
        .dms-fp-tag.deleted { color: #999; }
        .dms-fp-name { flex: 0 1 auto; min-width: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 600; color: #1b1b1b; }
        .dms-fp-size { flex: 0 0 auto; font-size: 12px; color: #888; }
        /* Real byte-driven fill. Width is animated frame-by-frame from the rAF
           loop in UploadingBar, so no CSS transition here — the two would fight. */
        .dms-fp-bar { flex: 1 1 auto; min-width: 40px; height: 8px; border-radius: 5px; background: #dde5e0; overflow: hidden; }
        /* display:block is REQUIRED — the fill is a <span>, and width/height are
           ignored on an inline box, so the rAF loop's style.width would silently
           do nothing. (.dms-fp-bar escapes this only because it is a flex item of
           .dms-fp-row and gets blockified.) */
        .dms-fp-fill { display: block; height: 100%; border-radius: 5px; background: #0f6c3f; }
        .dms-fp-pct { flex: 0 0 34px; text-align: right; font-size: 12px; color: #666; }
        .dms-fp-gap { flex: 1 1 auto; }
        .dms-fp-x { flex: 0 0 auto; background: none; border: none; cursor: pointer; color: #d13438; font-size: 14px; line-height: 1; padding: 2px 4px; }
        .dms-fp-x:disabled { color: #c9a3a4; cursor: default; }
        .dms-fp-xspacer { flex: 0 0 22px; }
        /* Overall progress bar (above the form) */
        .dms-overall { margin: 0 0 16px; }
        .dms-overall-head { display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; color: #333; margin-bottom: 6px; }
        .dms-overall-track { height: 10px; border-radius: 6px; background: #ececec; overflow: hidden; }
        .dms-overall-fill { height: 100%; background: #0f6c3f; border-radius: 6px; transition: width .3s ease; }
        /* Replace-file confirmation popup */
        .dms-popup-overlay { position: fixed; inset: 0; background: rgba(0,0,0,.25); backdrop-filter: blur(6px); -webkit-backdrop-filter: blur(6px); z-index: 9998; display: flex; align-items: center; justify-content: center; padding: 16px; }
        .dms-popup { background: #fff; border-radius: 16px; padding: 40px 40px 32px; text-align: center; max-width: 420px; width: 100%; box-shadow: 0 8px 40px rgba(0,0,0,.15); }
        .dms-popup-svg { width: 110px; height: 110px; display: block; margin: 0 auto 20px; }
        .dms-popup-title { font-size: 22px; font-weight: 700; color: #0f6c3f; margin: 0 0 12px; }
        .dms-popup-msg { font-size: 14px; color: #555; margin: 0 0 28px; line-height: 1.6; }
        .dms-popup-actions { display:flex; align-items:center; justify-content:center; gap: 12px; }
        .dms-popup-btn { min-width: 96px; border-radius: 6px; padding: 11px 22px; font-size: 14px; font-weight: 600; font-family: inherit; cursor: pointer; }
        .dms-popup-btn.confirm { background: #0f6c3f; color: #fff; border: none; }
        .dms-popup-btn.confirm:hover { background: #0a5230; }
        .dms-popup-btn.cancel { background: #fff; color: #0f6c3f; border: 1px solid #0f6c3f; }
        .dms-popup-btn.cancel:hover { background: #f0f6f2; }
        .dms-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 16px; font-size: 13px; }
        .dms-field > span { font-weight: 600; }
        .dms-field .req { color: #d13438; font-style: normal; }
        .dms-field select, .dms-field input[type="text"], .dms-field input[type="date"] { padding: 8px 10px; border: 1px solid #c8c8c8; border-radius: 4px; font: inherit; width: 100%; box-sizing: border-box; height: 38px; }
        .dms-field small { color: #666; font-size: 12px; font-weight: 400; }
        .dms-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0 24px; }
        .dms-radio-group { display: flex; gap: 24px; margin-bottom: 20px; align-items: center; }
        .dms-radio-group p { margin: 0; font-size: 13px; }
        .dms-radio-group label { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 600; cursor: pointer; color: #1b1b1b; }
        .dms-radio-group input[type="radio"] { accent-color: #0f6c3f; width: 16px; height: 16px; cursor: pointer; }
        .dms-dept-badge { display: inline-flex; align-items: center; gap: 8px; background: #e8f5ee; border: 1px solid #b3d9c4; border-radius: 20px; padding: 5px 14px; font-size: 13px; margin-bottom: 20px; }
        .dms-dept-badge .dept-label { font-weight: 400; color: #555; }
        .dms-dept-badge .dept-name { font-weight: 700; color: #0f6c3f; }
        .dms-dept-loading { font-size: 13px; color: #666; margin-bottom: 20px; }
        .dms-dept-error { font-size: 13px; color: #d13438; background: #fdf3f3; border: 1px solid #f1c0c0; border-radius: 4px; padding: 10px 14px; margin-bottom: 20px; }
        .dms-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 4px; }
        .dms-btn { padding: 9px 24px; border-radius: 4px; cursor: pointer; font: inherit; font-size: 14px; border: 1px solid transparent; }
        .dms-btn.primary { background: #0f6c3f; color: #fff; }
        .dms-btn.primary:disabled { background: #9bbfaa; cursor: default; }
        .dms-btn.secondary { background: #fff; border-color: #0f6c3f; color: #0f6c3f; }
        .dms-btn.secondary:disabled { border-color: #c8c8c8; color: #9b9b9b; cursor: default; }
        .dms-status { margin-top: 16px; font-size: 13px; }
        /* Batch tiles */
        .dms-batchlist { display: flex; flex-direction: column; gap: 12px; margin-bottom: 16px; }
        .dms-batch-tile { display: flex; gap: 14px; align-items: center; border: 1px solid #b3d9c4; border-radius: 8px; background: #fff; padding: 14px 16px; }
        .dms-batch-tile .ico { font-size: 30px; line-height: 1; }
        .dms-batch-tile .body { flex: 1; min-width: 0; }
        .dms-batch-tile .bt-title { font-weight: 700; color: #0f6c3f; font-size: 14px; }
        .dms-batch-tile .bt-path { font-size: 12px; color: #333; margin-top: 2px; word-break: break-word; }
        .dms-batch-tile .bt-meta { font-size: 11px; color: #666; margin-top: 4px; }
        .dms-batch-tile .bt-actions { display: flex; flex-direction: column; gap: 6px; }
        .dms-batch-tile .bt-actions button { background: none; border: none; cursor: pointer; font: inherit; font-size: 12px; padding: 0; }
        .dms-batch-tile .bt-edit { color: #0f6c3f; }
        .dms-batch-tile .bt-remove { color: #d13438; }
        .dms-batch-tile .bt-actions button:disabled { color: #9b9b9b; cursor: default; }
        .dms-add-batch { display: inline-flex; align-items: center; gap: 6px; background: #f4f8f5; border: 1px dashed #b3d9c4; color: #0f6c3f; border-radius: 6px; padding: 10px 16px; font: inherit; font-size: 13px; font-weight: 600; cursor: pointer; }
        .dms-add-batch:disabled { border-color: #ddd; color: #9b9b9b; cursor: default; }
        .dms-panel { border: 1px solid #d6e6dc; border-radius: 8px; padding: 20px; margin-top: 8px; background: #fbfdfc; }
        .dms-panel-title { font-size: 13px; font-weight: 700; color: #0f6c3f; margin: 0 0 16px; }
        /* Results */
        .dms-results { margin-top: 20px; }
        .dms-batch-result { border: 1px solid #e1e1e1; border-radius: 8px; padding: 16px; margin-bottom: 14px; }
        .dms-batch-result h4 { margin: 0 0 4px; font-size: 14px; color: #1b1b1b; }
        .dms-batch-result .bt-dest { font-size: 12px; color: #666; margin: 0 0 10px; }
        .dms-results-summary { font-size: 13px; font-weight: 600; margin: 0 0 10px; }
        .dms-batch-error { background: #fdf3f3; color: #d13438; border: 1px solid #f1c0c0; border-radius: 4px; padding: 10px 14px; font-size: 13px; }
        .dms-result { display: flex; gap: 10px; align-items: baseline; font-size: 13px; padding: 7px 10px; border-radius: 4px; margin-bottom: 6px; }
        .dms-result .tag { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }
        .dms-result .fname { word-break: break-all; }
        .dms-result .why { color: #666; font-size: 12px; }
        .dms-result.uploaded { background: #e8f5ee; } .dms-result.uploaded .tag { color: #0f6c3f; }
        .dms-result.skipped { background: #fff4e5; } .dms-result.skipped .tag { color: #7a4f00; }
        .dms-result.failed { background: #fdf3f3; } .dms-result.failed .tag { color: #d13438; }
        .dms-result.tagFailed { background: #fff4e5; } .dms-result.tagFailed .tag { color: #7a4f00; }
        .dms-toast { position: fixed; top: 24px; right: 24px; z-index: 9999; min-width: 300px; max-width: 460px; padding: 14px 40px 14px 16px; border-radius: 6px; font-size: 13px; font-family: 'Segoe UI', sans-serif; box-shadow: 0 4px 16px rgba(0,0,0,.18); animation: dms-slidein .2s ease; }
        .dms-toast.error { background: #d13438; color: #fff; }
        .dms-toast.success { background: #0f6c3f; color: #fff; }
        .dms-toast-close { position: absolute; top: 10px; right: 12px; background: none; border: none; cursor: pointer; font-size: 16px; color: inherit; opacity: .7; line-height: 1; }
        .dms-toast-close:hover { opacity: 1; }
        @keyframes dms-slidein { from { transform: translateX(60px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        @media (max-width: 640px) {
          .dms-grid { grid-template-columns: 1fr; }
          .dms-toast { left: 12px; right: 12px; min-width: unset; top: 12px; }
        }
      `}</style>

      <h2>Bulk upload</h2>
      <p className="dms-subtitle">
        Queue up to {MAX_BATCHES} batches of {MAX_FILES} files each — every batch
        goes to its own folder with its own metadata. All fields marked{" "}
        <strong>*</strong> are required.
      </p>
      <div className="dms-warn">
        <span aria-hidden="true">⚠</span>
        <span>
          <strong>Temporary tool.</strong> Files go straight into the{" "}
          <strong>Documents</strong> library — they skip Staging and the approval
          step entirely, and are visible to everyone with access to the
          destination folder as soon as they upload.
        </span>
      </div>

      {/* ── Live upload progress (above the form) ───────────────────────── */}
      {live && (
        <div className="dms-section">
          <p className="dms-section-title">Upload progress</p>
          {liveTotals.total > 0 && (
            <div className="dms-overall">
              <div className="dms-overall-head">
                <span>
                  {liveTotals.done} of {liveTotals.total} file
                  {liveTotals.total === 1 ? "" : "s"} processed
                </span>
                <span>{livePct}%</span>
              </div>
              <div className="dms-overall-track">
                <div
                  className="dms-overall-fill"
                  style={{ width: `${livePct}%` }}
                />
              </div>
            </div>
          )}
          <div className="dms-progress">
            {live.map((lb, bi) => (
              <div className="dms-progress-batch" key={bi}>
                <div className="dms-progress-head">
                  {lb.label} — {lb.destinationLabel}
                </div>
                <div className="dms-progress-list">
                  {lb.files.map((f, fi) => {
                    const canDelete =
                      (f.state === "done" || f.state === "tagFailed") && !!f.sru;
                    const rowKey = `${bi}:${fi}`;
                    return (
                      <div className={`dms-fp-row ${f.state}`} key={fi}>
                        <span className={`dms-fp-tag ${f.state}`}>
                          {fpLabel[f.state]}
                        </span>
                        <span className="dms-fp-name" title={f.name}>
                          {f.name}
                        </span>
                        <span className="dms-fp-size">{fmtSize(f.size)}</span>
                        {f.state === "uploading" ? (
                          <UploadingBar
                            target={f.pct}
                            label={`Uploading ${f.name}`}
                          />
                        ) : (
                          <span className="dms-fp-gap" aria-hidden="true" />
                        )}
                        {canDelete ? (
                          <button
                            type="button"
                            className="dms-fp-x"
                            aria-label={`Delete ${f.name} from Documents`}
                            title="Delete this file from Documents"
                            disabled={deletingKey === rowKey}
                            onClick={() => {
                              deleteUploaded(bi, fi, f.sru as string, f.name).catch(
                                () => undefined,
                              );
                            }}
                          >
                            ✕
                          </button>
                        ) : (
                          <span className="dms-fp-xspacer" aria-hidden="true" />
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ── Batches ─────────────────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Batches</p>

        {deptLoading ? (
          <p className="dms-dept-loading">Loading your access&hellip;</p>
        ) : !privileged && validPaths.length === 0 ? (
          <div className="dms-dept-error">
            Your account isn&apos;t fully provisioned to upload — you need
            membership at every level plus the unit uploader role. Contact your
            administrator.
          </div>
        ) : null}

        {visibleBatches.length > 0 && (
          <div className="dms-batchlist">
            {visibleBatches.map((b) => {
              const n = batches.findIndex((x) => x.id === b.id) + 1;
              return (
                <div className="dms-batch-tile" key={b.id}>
                  <div className="ico" aria-hidden="true">
                    📁
                  </div>
                  <div className="body">
                    <div className="bt-title">Batch {n}</div>
                    <div className="bt-path">{b.destinationLabel}</div>
                    <div className="bt-meta">
                      {[b.yearLabel, b.docTypeLabel, `${b.files.length} files`]
                        .filter(Boolean)
                        .join(" · ")}
                    </div>
                  </div>
                  <div className="bt-actions">
                    <button
                      type="button"
                      className="bt-edit"
                      disabled={busy || panelOpen}
                      onClick={() => editBatch(b)}
                    >
                      Edit
                    </button>
                    <button
                      type="button"
                      className="bt-remove"
                      disabled={busy || panelOpen}
                      onClick={() => removeBatch(b.id)}
                    >
                      Remove
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {!panelOpen && !deptLoading && (
          <button
            type="button"
            className="dms-add-batch"
            disabled={!canAdd || busy}
            onClick={openAddPanel}
            title={
              batches.length >= MAX_BATCHES
                ? `Maximum ${MAX_BATCHES} batches`
                : undefined
            }
          >
            + Add batch
          </button>
        )}

        {/* ── Inline config panel ──────────────────────────────────────── */}
        {panelOpen && (
          <div className="dms-panel">
            <p className="dms-panel-title">
              {editingId
                ? "Edit batch"
                : `Configure batch ${batches.length + 1}`}
            </p>

            {/* Files */}
            {/* The whole card is clickable to open the file picker — same as the
                single-file form. The label is a plain span, not a button, so the
                container is the only interactive element (no nested buttons). */}
            <div
              className="dms-drop"
              role="button"
              tabIndex={busy ? -1 : 0}
              aria-disabled={busy}
              style={{ cursor: busy ? "default" : "pointer" }}
              onClick={() => { if (!busy) fileRef.current?.click(); }}
              onKeyDown={(e) => {
                if (busy) return;
                if (e.key === "Enter" || e.key === " ") {
                  e.preventDefault();
                  fileRef.current?.click();
                }
              }}
            >
              <span className="size">
                {picked.length === 0
                  ? "Click here to select files"
                  : `${picked.length} file(s) selected`}
              </span>
              <span
                className="dms-link"
                style={busy ? { color: "#9bbfaa" } : undefined}
              >
                {picked.length === 0 ? "Select files" : "Add more files"}
              </span>
              <input
                ref={fileRef}
                type="file"
                multiple
                accept={settings.allowedExtensions.join(",")}
                style={{ display: "none" }}
                onChange={(e) => addFiles(e.target.files)}
              />
            </div>

            {/* (Selected-file list is rendered below the metadata form.) */}

            {/* Destination + metadata */}
            <div style={{ marginTop: 20 }}>
              <div className="dms-radio-group">
                <p>Upload into:</p>
                {(["BusinessSegment", "Project"] as const).map((side) => {
                  const sideModes = modes.filter((m) => m.side === side);
                  const offerable = privileged
                    ? sideModes
                    : sideModes.filter((m) =>
                        validPaths.some((p) => p.modeKey === m.key),
                      );
                  if (offerable.length === 0) return null;
                  const active = activeMode()?.side === side;
                  return (
                    <label key={side}>
                      <input
                        type="radio"
                        name="bulkSideToggle"
                        checked={active}
                        onChange={() => switchMode(offerable[0].key)}
                      />
                      {side === "BusinessSegment"
                        ? "Business Segment"
                        : "Project"}
                    </label>
                  );
                })}
              </div>

              {(() => {
                const side = activeMode()?.side;
                const sideModes = modes.filter((m) => m.side === side);
                const offerable = privileged
                  ? sideModes
                  : sideModes.filter((m) =>
                      validPaths.some((p) => p.modeKey === m.key),
                    );
                if (offerable.length <= 1) return null;
                return (
                  <label className="dms-field">
                    <span>Segment</span>
                    <select
                      value={uploadMode}
                      onChange={(e) => switchMode(e.target.value)}
                    >
                      {offerable.map((m) => (
                        <option key={m.key} value={m.key}>
                          {m.label}
                        </option>
                      ))}
                    </select>
                  </label>
                );
              })()}

              {!privileged && activeMode() && (
                <div className="dms-dept-badge">
                  <span className="dept-label">Uploading to:</span>
                  <span className="dept-name">
                    {[
                      activeMode()?.label,
                      ...(activeMode()?.levels ?? []).map(
                        (_lvl, i) =>
                          (levelChoices[i] ?? []).find(
                            (o) => o.id === levelValues[i],
                          )?.label,
                      ),
                    ]
                      .filter(Boolean)
                      .join(" › ")}
                  </span>
                </div>
              )}

              <div className="dms-grid">
                {(activeMode()?.levels ?? []).map((lvl, i) =>
                  renderSelect(
                    lvl.label,
                    true,
                    levelValues[i] ?? "",
                    (v) => {
                      const md = activeMode();
                      if (md) handleLevelChange(md, i, v);
                    },
                    levelChoices[i] ?? [],
                    deptLoading ||
                      isLevelLocked(i) ||
                      (i > 0 && !levelValues[i - 1]),
                    "--",
                    true,
                  ),
                )}

                {/* Graceful empty-state: a segment whose term set has no child terms
                    yet (non-GHO Head Offices before their trees are added) would show a
                    dead "--" dropdown. Name the missing level instead. */}
                {(() => {
                  const md = activeMode();
                  if (!md || md.levels.length === 0 || deptLoading || levelChoices.length === 0) return null;
                  for (let i = 0; i < md.levels.length; i++) {
                    const parentChosen = i === 0 || !!levelValues[i - 1];
                    if (parentChosen && (levelChoices[i]?.length ?? 0) === 0) {
                      return (
                        <div
                          key="dms-empty-level"
                          style={{ gridColumn: "1 / -1", padding: "8px 12px", background: "#fff8e1", border: "1px solid #f0c000", borderRadius: 4, fontSize: 13, color: "#7a5b00" }}
                        >
                          No {md.levels[i].label.toLowerCase()} options are configured for this segment yet — ask your administrator to add them in the term store before uploading here.
                        </div>
                      );
                    }
                  }
                  return null;
                })()}

                {renderSelect(
                  "Year / Period",
                  true,
                  yearPeriod,
                  setYearPeriod,
                  options.yearPeriod,
                )}

                {renderSelect(
                  "Document Type",
                  true,
                  documentType,
                  setDocumentType,
                  options.documentType,
                )}

                <label className="dms-field">
                  <span>
                    Document Date <em className="req">*</em>
                  </span>
                  <input
                    type="date"
                    value={documentDate}
                    max={(() => {
                      const d = new Date();
                      const mm = d.getMonth() + 1;
                      const day = d.getDate();
                      return `${d.getFullYear()}-${mm < 10 ? "0" + mm : mm}-${day < 10 ? "0" + day : day}`;
                    })()}
                    onChange={(e) => setDocumentDate(e.target.value)}
                  />
                </label>

                {renderSelect(
                  "Confidentiality Level",
                  true,
                  confidentiality,
                  setConfidentiality,
                  options.confidentiality,
                )}

                {/* Vendor is hidden for now (client request, 2026-07-28). The
                    state, term-set load and formValues push all stay wired up —
                    restoring the field is just un-commenting this select.
                    {renderSelect(
                      "Vendor (if applicable)",
                      false,
                      vendor,
                      setVendor,
                      options.vendor,
                    )} */}
              </div>

              {/* Selected files — listed below the form */}
              {picked.length > 0 && (
                <div className="dms-fileblock">
                  <p className="dms-fileblock-title">Selected files</p>
                  <div className="dms-filelist">
                    {picked.map((p) => (
                      <div className="dms-filerow" key={p.key}>
                        <div>
                          <div className="orig">{p.file.name}</div>
                          <div className="size">
                            {(p.file.size / 1024).toFixed(1)} KB
                          </div>
                        </div>
                        <button
                          type="button"
                          className="dms-remove"
                          aria-label={`Remove ${p.file.name}`}
                          disabled={busy}
                          onClick={() => removeFile(p.key)}
                        >
                          ✕
                        </button>
                      </div>
                    ))}
                  </div>
                  <div className="dms-filelist-foot">
                    <span className="dms-count">
                      {picked.length} file{picked.length === 1 ? "" : "s"}{" "}
                      selected ({MAX_FILES} max).
                    </span>
                  </div>
                </div>
              )}

              <div className="dms-actions">
                <button
                  type="button"
                  className="dms-btn secondary"
                  onClick={cancelPanel}
                  disabled={busy}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  className="dms-btn primary"
                  onClick={saveBatch}
                  disabled={busy}
                >
                  Save batch
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ── Actions ─────────────────────────────────────────────────────── */}
      <div className="dms-actions">
        <button
          type="button"
          className="dms-btn secondary"
          onClick={clearAll}
          disabled={busy || (batches.length === 0 && !panelOpen)}
        >
          Clear
        </button>
        <button
          type="button"
          className="dms-btn primary"
          onClick={handleUpload}
          disabled={busy || deptLoading || panelOpen || batches.length === 0}
          title={
            panelOpen
              ? "Save or cancel the batch you're editing first"
              : undefined
          }
        >
          {busy
            ? "Uploading…"
            : `Upload all${totalQueued > 0 ? ` (${totalQueued} files)` : ""}`}
        </button>
      </div>

      {status && <p className="dms-status">{status}</p>}

      {/* ── Per-batch results ───────────────────────────────────────────── */}
      {batchResults && (
        <div className="dms-results">
          {batchResults.map((bo, bi) => {
            const s = {
              uploaded: bo.results.filter((r) => r.outcome === "uploaded")
                .length,
              skipped: bo.results.filter((r) => r.outcome === "skipped").length,
              failed: bo.results.filter((r) => r.outcome === "failed").length,
              tagFailed: bo.results.filter((r) => r.outcome === "tagFailed")
                .length,
            };
            const allOk =
              !bo.batchError &&
              s.uploaded === bo.results.length &&
              bo.results.length > 0;
            const shownRows = allOk
              ? []
              : bo.results.filter((r) => r.outcome !== "uploaded");
            return (
              <div className="dms-batch-result" key={`${bo.batchId}-${bi}`}>
                <h4>{bo.label}</h4>
                <p className="bt-dest">{bo.destinationLabel}</p>
                {bo.batchError ? (
                  <div className="dms-batch-error">{bo.batchError}</div>
                ) : (
                  <>
                    <p className="dms-results-summary">
                      {s.uploaded} uploaded
                      {s.skipped > 0 && `, ${s.skipped} skipped`}
                      {s.failed > 0 && `, ${s.failed} failed`}
                      {s.tagFailed > 0 &&
                        `, ${s.tagFailed} uploaded without tags`}
                      .
                      {s.tagFailed > 0 &&
                        " Files listed as “uploaded, not tagged” are already in Documents — fix their metadata in the library rather than re-uploading."}
                    </p>
                    {shownRows.map((r, i) => (
                      <div
                        className={`dms-result ${r.outcome}`}
                        key={`${r.name}-${i}`}
                      >
                        <span className="tag">{outcomeLabel[r.outcome]}</span>
                        <span className="fname">{r.name}</span>
                        {r.detail && <span className="why">{r.detail}</span>}
                      </div>
                    ))}
                  </>
                )}
              </div>
            );
          })}
        </div>
      )}

      {replaceAsk && (
        <div className="dms-popup-overlay" role="dialog" aria-modal="true">
          <div className="dms-popup">
            <svg
              className="dms-popup-svg"
              viewBox="0 0 184 184"
              fill="none"
              xmlns="http://www.w3.org/2000/svg"
            >
              <circle opacity="0.3" cx="92.0001" cy="92" r="75.4872" fill="#FF952A" />
              <circle cx="92" cy="92" r="92" fill="#FF952A" fillOpacity="0.2" />
              <circle cx="92.0003" cy="91.9998" r="61.3333" fill="white" stroke="#FF952A" strokeWidth="3" />
              <path d="M93 66L93 100" stroke="#FF952A" strokeWidth="10" strokeLinecap="round" />
              <path d="M93 116.804L93 118" stroke="#FF952A" strokeWidth="10" strokeLinecap="round" />
            </svg>
            <p className="dms-popup-title">Replace Existing File</p>
            <p className="dms-popup-msg">
              We noticed there&rsquo;s a same name file
              {replaceAsk.name ? `: "${replaceAsk.name}"` : ""}.
              <br />
              Are you sure you want to override this file?
            </p>
            <div className="dms-popup-actions">
              <button
                className="dms-popup-btn confirm"
                onClick={() => answerReplace(true)}
              >
                Yes
              </button>
              <button
                className="dms-popup-btn cancel"
                onClick={() => answerReplace(false)}
              >
                No
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div className={`dms-toast ${toast.type}`} role="alert">
          {toast.message}
          <button
            className="dms-toast-close"
            onClick={() => setToast(null)}
            aria-label="Dismiss"
          >
            ✕
          </button>
        </div>
      )}
    </section>
  );
}
