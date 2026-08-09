import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IBulkUploadProps } from "./IBulkUploadProps";
import {
  lookupFolderMapping,
  resolveFolderServerUrl,
  probeFolderByPath,
  ensureFolder,
  encodeServerRelativePath,
} from "../../../shared/dmsFolderMap";
import {
  parseLevels,
  collectMembership,
  isLeafChainValid,
  sanitizeFolderSegment,
  buildLevelFormValues,
  Level,
  GroupMapRow,
  Membership,
  ColumnPair,
} from "../../../shared/formModel";
import {
  AllowedFileTypes,
  FALLBACK_FILE_TYPES,
  NO_TYPES_MESSAGE,
  readAllowedFileTypesField,
  resolveAllowedFileTypes,
} from "../../../shared/allowedFileTypes";
import { cachedListTitle, LIST_SUFFIX, libraryTitle, libraryUrlSegment } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";

/* ----------------------------------------------------------------------------
 * BULK UPLOAD — a duplicate of the `form` web part, with two behaviour changes:
 *   1. Writes straight into the Documents library, bypassing Staging (and so
 *      bypassing content approval and the Auto-route flow).
 *   2. Up to MAX_FILES files in ONE selection, sharing ONE destination folder
 *      and ONE metadata set.
 *
 * Batching (two independent destinations per run) was removed in Phase 1 — see
 *   docs/superpowers/specs/2026-08-03-bulk-upload-single-selection-design.md
 * which supersedes 2026-07-24-bulk-upload-two-batch-design.md. The screen exists
 * to migrate the client's HISTORICAL documents, which arrive already named and
 * belong to one folder at a time, so one shared metadata set is sufficient.
 *
 * Files keep their own names here — the Form's
 * [Project] - [Vendor] - [Name] - [Date] composition deliberately does NOT apply
 * (one metadata set would hand all 50 files the same name).
 *
 * TEMPORARY TOOL. See also:
 *   docs/superpowers/specs/2026-07-23-bulk-upload-direct-to-documents-design.md
 *
 * Metadata tagging, group-based mode detection, the term cascade and per-Unit
 * folder routing are cloned from Form.tsx unchanged.
 * -------------------------------------------------------------------------- */

const MAX_FILES = 50;

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
  // A dedicated "Remark" column, NOT the built-in _ExtendedDescription — matches Form.tsx.
  remark: "Remark",
  legallyPrivileged: "LegallyPrivileged",
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

// One picked file. `key` survives removals so React rows stay stable.
type PickedFile = { key: string; file: File };

type Outcome = "uploaded" | "skipped" | "failed" | "tagFailed";
type FileResult = { name: string; outcome: Outcome; detail?: string };

// A resolved level term (label + id) plus the pair of column internal names it
// writes to — everything needed to build formValues once for the whole run.
type LevelSelection = {
  column: string;
  label: string;
  id: string;
  labelCol?: string;
  tidCol?: string;
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

// `sru` is the uploaded file's ServerRelativeUrl, captured once a file lands in
// the Documents library (done / tagFailed) so the per-row X can delete it after
// the fact. `pct` is real bytes-sent progress, 0-100, only meaningful while
// state === "uploading".
type LiveFile = {
  name: string;
  size: number;
  state: FileState;
  sru?: string;
  pct: number;
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
    termSetGuid: "08dd94cb-f76c-431c-9b37-e9c98f739ffc",
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
    termSetGuid: "9ad00b00-a43c-4a8b-a39a-d0efa89ba706",
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
    termSetGuid: "77c3993b-0c3c-4a18-89d9-d69209886322",
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
    remark: string;
    legallyPrivileged: string;
    businessSegmentLabel: string;
    businessSegmentTid: string;
  };
  stagingLibrary: string;
  // Term GUID of the ONE confidentiality level that offers the Legally Privileged
  // tick. Empty = never offered. Mirrors Form.tsx.
  legallyPrivilegedFor: string;
  allowedFileTypes: AllowedFileTypes;
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
    remark: FIELDS.remark,
    legallyPrivileged: FIELDS.legallyPrivileged,
    businessSegmentLabel: LEVEL_COLUMNS.BusinessSegment.label,
    businessSegmentTid: LEVEL_COLUMNS.BusinessSegment.tid,
  },
  stagingLibrary: "Staging",
  // Empty by default: the tick appears only once a site sets legallyPrivilegedFor.
  legallyPrivilegedFor: "",
  // "unknown", not "configured": reaching this constant means DMS Config could not
  // be read, so the UI must not present these as configured values. Mirrors Form.tsx.
  allowedFileTypes: { kind: "unknown", types: FALLBACK_FILE_TYPES },
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

// Human-friendly file size for the file and progress rows.
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
 * uploads at a time, but 60 setState calls a second to re-render a list of
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
  const fileSeqRef = useRef<number>(1);
  // Cached form digest for the XHR upload path (spHttpClient handles its own).
  // Digests expire — SharePoint tells us when, and we refresh a minute early.
  const digestRef = useRef<{ value: string; expiresAt: number } | null>(null);

  const [options, setOptions] = useState<OptionMap>(EMPTY_OPTIONS);
  const [deptLoading, setDeptLoading] = useState<boolean>(true);
  // Privileged = site admin: bypasses tier detection and gets the full manual
  // cascade (may upload anywhere).
  const [privileged, setPrivileged] = useState<boolean>(false);

  // Generic N-level cascade state: one option list + one selected term id per level.
  const [levelChoices, setLevelChoices] = useState<TermOption[][]>([]);
  const [levelValues, setLevelValues] = useState<string[]>([]);
  const [validPaths, setValidPaths] = useState<ValidPath[]>([]);

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);

  // ── The one selection: files + the single metadata set they all share ──
  const [uploadMode, setUploadMode] = useState<string>("");
  const [picked, setPicked] = useState<PickedFile[]>([]);
  const [dragOver, setDragOver] = useState<boolean>(false);
  const [documentType, setDocumentType] = useState<string>("");
  const [yearPeriod, setYearPeriod] = useState<string>("");
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [legallyPrivileged, setLegallyPrivileged] = useState<boolean>(false);
  const [remark, setRemark] = useState<string>("");
  // Vendor stays wired but unset — the select is hidden (client request 2026-07-28).
  const [vendor] = useState<string>("");

  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [runError, setRunError] = useState<string | null>(null);
  const [results, setResults] = useState<FileResult[] | null>(null);
  const [live, setLive] = useState<LiveFile[] | null>(null);
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
  // Success dialog, matching the Form's. Separate from `toast` because this one is
  // modal and has to survive until the user chooses where to go next.
  const [doneOpen, setDoneOpen] = useState<boolean>(false);
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

  /**
   * Live title of one of our lists — "CRS Config" on a renamed site, "DMS Config" otherwise.
   * Primed per call so there is no ordering dependency; the cache short-circuits after the first.
   */
  const listName = async (suffix: string): Promise<string> => {
    await primeNames(context.spHttpClient, siteUrl);
    return encodeURIComponent(cachedListTitle(suffix));
  };

  const loadModes = async (): Promise<UploadMode[]> => {
    const config = await listName(LIST_SUFFIX.config);
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${config}')/items?$select=Title,ModeLabel,Category,TermSetGuid,StagingFolder,Levels,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error(`${decodeURIComponent(config)} list not found`);
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
    const groupMap = await listName(LIST_SUFFIX.groupMap);
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${groupMap}')/items?$select=GroupId,GroupName,Segment,UnitTermGuid,Role&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`${decodeURIComponent(groupMap)} list not found`);
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

  // See Form.tsx: a site whose DMS Config predates the AllowedFileTypes column
  // answers HTTP 400 to the ENTIRE request, which would drop every other setting
  // too. Retry without the new field. Spec 2026-07-30 §5.
  const SETTINGS_FIELDS = "Title,SettingValue,AllowedFileTypes";
  const SETTINGS_FIELDS_LEGACY = "Title,SettingValue";

  const fetchSettingRows = async (
    select: string,
  ): Promise<SPHttpClientResponse> =>
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${await listName(LIST_SUFFIX.config)}')/items?$select=${select}&$filter=ConfigType eq 'setting'`,
      SPHttpClient.configurations.v1,
      // Pinned to nometadata so multi-choice fields arrive as a plain array.
      { headers: { Accept: "application/json;odata=nometadata" } },
    );

  const loadSettings = async (): Promise<DmsSettings> => {
    let res = await fetchSettingRows(SETTINGS_FIELDS);
    if (!res.ok) {
      const body = await res.text();
      console.warn(
        `DMS Config read including AllowedFileTypes failed (HTTP ${res.status}). ` +
          `Retrying without that field. Response: ${body}`,
      );
      res = await fetchSettingRows(SETTINGS_FIELDS_LEGACY);
    }
    if (!res.ok) {
      const body = await res.text();
      throw new Error(
        `DMS Config read failed: HTTP ${res.status}. Response: ${body}`,
      );
    }
    const data = await res.json();
    const map: Record<string, string> = {};
    // Keyed on the property's PRESENCE, not its value: SharePoint sends null for
    // an emptied multi-choice field. Key missing = column absent (-> unknown);
    // key present = read it, null included (-> none). Verified 2026-07-30.
    let rawFileTypes: string[] | undefined;
    (data.value ?? []).forEach(
      (item: {
        Title: string;
        SettingValue: string;
        AllowedFileTypes?: unknown;
      }) => {
        map[item.Title] = item.SettingValue;
        if (item.Title === "allowedExtensions") {
          rawFileTypes = readAllowedFileTypesField(item);
        }
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
        remark: get("col_remark") ?? DEFAULT_SETTINGS.columns.remark,
        legallyPrivileged:
          get("col_legallyPrivileged") ?? DEFAULT_SETTINGS.columns.legallyPrivileged,
        businessSegmentLabel:
          get("col_businessSegment") ?? DEFAULT_SETTINGS.columns.businessSegmentLabel,
        businessSegmentTid:
          get("col_businessSegmentTid") ?? DEFAULT_SETTINGS.columns.businessSegmentTid,
      },
      // The library's LIVE title wins over the config row — see the same note in Form.tsx.
      // NOTE this holds the TITLE. The path-swap sites below need the URL SEGMENT, which is a
      // different string since the rename, and they read libraryUrlSegment() instead.
      stagingLibrary: ((): string => {
        const configured = (get("stagingLibrary") ?? "").trim();
        return configured.length > 0 && configured !== "Staging" ? configured : libraryTitle();
      })(),
      legallyPrivilegedFor:
        get("legallyPrivilegedFor") ?? DEFAULT_SETTINGS.legallyPrivilegedFor,
      // SettingValue is deliberately NOT consulted for file types any more —
      // AllowedFileTypes is the single source of truth. Spec 2026-07-30 §3.
      allowedFileTypes: resolveAllowedFileTypes(rawFileTypes),
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
        // Both branches spelled out. `!= null` meant "neither null nor undefined" — correct,
        // but invisible, and "fixing" it to `!== null` would let undefined through and turn a
        // missing id into the string "undefined".
        .map((g) => (g.Id !== null && g.Id !== undefined ? String(g.Id) : ""))
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

  /* ---------- Cascade builders -------------------------------------------- */

  const initCascade = async (mode: UploadMode): Promise<void> => {
    const tops = await loadTermSet(mode.termSetGuid).catch(
      () => [] as TermOption[],
    );
    setLevelChoices([tops]);
    setLevelValues([]);
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
      if (isLeafChainValid(chain.map((c) => c.id), leaf.termGuid)) {
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

  // `paths` and `isPrivileged` are passed in rather than read from state because
  // init() calls this in the same tick it resolves them — state would still be
  // the initial empty/false at that point.
  const initModeCascade = (
    mode: UploadMode,
    paths: ValidPath[],
    isPrivileged: boolean,
  ): void => {
    if (isPrivileged) {
      initCascade(mode).catch(() => {
        setLevelChoices([]);
        setLevelValues([]);
      });
    } else {
      applyRestrictedMode(paths, mode, []);
    }
  };

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

      const [docTypes, years, confs] = await Promise.all([
        loadTermSet(loadedSettings.termSets.documentType).catch(
          () => [] as TermOption[],
        ),
        loadTermSet(loadedSettings.termSets.yearPeriod).catch(
          () => [] as TermOption[],
        ),
        loadTermSet(loadedSettings.termSets.confidentiality).catch(
          () => [] as TermOption[],
        ),
      ]);
      setOptions({
        documentType: docTypes,
        yearPeriod: years,
        confidentiality: confs,
        // The vendor term set was deleted from the site 2026-07-29 — Vendor/Customer
        // Name is free text in the upload form now. Fetching it here was a guaranteed
        // 404 on every load. The hidden-field plumbing is left intact (client
        // request 2026-07-28), so this stays an empty list rather than being removed.
        vendor: [],
      });

      let paths: ValidPath[] = [];
      if (!isPrivileged) {
        paths = await resolveValidPaths(loadedModes, membership);
        setValidPaths(paths);
      }

      // Open on a destination the user can actually reach, so both cards are
      // populated from the first paint. The batch panel used to do this on
      // "+ Add batch"; with one selection there is no such moment.
      const offerable = new Set(paths.map((p) => p.modeKey));
      const defaultMode = isPrivileged
        ? loadedModes.filter((m) => m.side === "BusinessSegment")[0] ??
          loadedModes[0]
        : loadedModes.find(
            (m) => m.side === "BusinessSegment" && offerable.has(m.key),
          ) ?? loadedModes.find((m) => offerable.has(m.key));
      if (defaultMode) {
        setUploadMode(defaultMode.key);
        initModeCascade(defaultMode, paths, isPrivileged);
      }
      setDeptLoading(false);
    };

    init().catch((err) => {
      console.error("BulkUpload init failed:", err);
      showToast("Could not load form data. Please refresh the page.", "error");
      setDeptLoading(false);
    });
  }, []);

  /* ---------- Selection helpers ------------------------------------------- */

  const switchMode = (modeKey: string): void => {
    setUploadMode(modeKey);
    const mode = modes.find((m) => m.key === modeKey);
    if (!mode) return;
    initModeCascade(mode, validPaths, privileged);
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

  const resetForm = (): void => {
    setPicked([]);
    setDocumentType("");
    setYearPeriod("");
    setDocumentDate("");
    setConfidentiality("");
    setLegallyPrivileged(false);
    setRemark("");
    setResults(null);
    setRunError(null);
    setLive(null);
    setStatus("");
    if (fileRef.current) fileRef.current.value = "";
    const mode = activeMode();
    if (mode) initModeCascade(mode, validPaths, privileged);
  };

  /**
   * Add files to the selection — APPENDS, it never replaces. Every file in this
   * web part arrives here, from the picker or from a drop, which is the point:
   * `accept` filters the DIALOG only and a drop bypasses it entirely, so the
   * extension gate has to live at this one choke point. Spec 2026-08-03 §4.
   */
  const addFiles = (list: FileList | null): void => {
    if (!list || list.length === 0) return;
    const allowedTypes = settings.allowedFileTypes;
    if (allowedTypes.kind === "none") {
      showToast(NO_TYPES_MESSAGE, "error");
      return;
    }
    const incoming = Array.from(list);
    const allowed: File[] = [];
    const rejected: string[] = [];
    incoming.forEach((f) => {
      const ok = allowedTypes.types.some((ext) =>
        f.name.toLowerCase().endsWith(ext),
      );
      if (ok) allowed.push(f);
      else rejected.push(f.name);
    });

    // Named, not counted: "3 files were skipped" leaves the uploader hunting for
    // which three in a list of fifty. The rest of the pick still proceeds.
    if (rejected.length > 0) {
      showToast(
        `Not added — file type not allowed: ${rejected.join(", ")}. Allowed: ${allowedTypes.types.join(", ")}`,
        "error",
      );
    }

    setPicked((prev) => {
      // The 50 applies to the COMBINED selection after this add, not per pick.
      const room = MAX_FILES - prev.length;
      if (room <= 0) {
        showToast(
          `You can upload at most ${MAX_FILES} files at once. Remove some first.`,
          "error",
        );
        return prev;
      }
      if (allowed.length > room) {
        showToast(
          `Only the first ${room} of these ${allowed.length} files were added — the limit is ${MAX_FILES} files in one upload.`,
          "error",
        );
      }
      const added = allowed.slice(0, room).map((file) => ({
        key: `f${fileSeqRef.current++}`,
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
  // in the live progress list (bulk upload writes straight to Documents). The
  // row keeps its `sru` (ServerRelativeUrl) captured at upload time, so the
  // delete targets the exact file regardless of any folder rename. Marks the row
  // "deleted" on success.
  const [deletingKey, setDeletingKey] = useState<string | null>(null);
  const deleteUploaded = async (
    fileIdx: number,
    sru: string,
    name: string,
  ): Promise<void> => {
    const key = String(fileIdx);
    if (!window.confirm(`Delete "${name}" from Documents? This cannot be undone.`)) {
      return;
    }
    setDeletingKey(key);
    try {
      const safeUrl = sru.replace(/'/g, "''");
      const delRes: SPHttpClientResponse = await context.spHttpClient.fetch(
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)?@f='${encodeServerRelativePath(safeUrl)}'`,
        SPHttpClient.configurations.v1,
        { method: "POST", headers: { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" } },
      );
      if (!delRes.ok && delRes.status !== 404) {
        throw new Error(`HTTP ${delRes.status}`);
      }
      setLive((prev) =>
        prev
          ? prev.map((f, fi) =>
              fi === fileIdx
                ? { ...f, state: "deleted" as FileState, sru: undefined }
                : f,
            )
          : prev,
      );
      showToast(`"${name}" deleted from Documents.`, "success");
    } catch (err) {
      console.error("Delete from Documents failed:", name, err);
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
    // URL SEGMENT, not the title: this slices a server-relative path. Since the rename the
    // two differ ("Approval Document" vs "/ApprovalDocument"), and using the title here
    // matches nothing — the guard below returns null and the caller reports "could not work
    // out the Documents path" for every folder.
    const prefix = `${webSru}/${libraryUrlSegment()}/`;
    if (stagingSru.toLowerCase().indexOf(prefix.toLowerCase()) !== 0) {
      return null;
    }
    const rest = stagingSru.slice(prefix.length);
    return `${webSru}/${DOCUMENTS_URL_SEGMENT}/${rest}`;
  };

  /* ---------- Form digest (XHR upload path) -------------------------------- */

  // A long run can outlive one digest, so cache it with its real expiry and
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

  /* ---------- Validation --------------------------------------------------- */

  const validate = (): string[] => {
    const missing: string[] = [];
    if (picked.length === 0) missing.push("Files");
    const m = activeMode();
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) missing.push(lvl.label);
    });
    if (!yearPeriod) missing.push("Year");
    if (!documentType) missing.push("Document Type");
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidential Level");
    return missing;
  };

  /* ---------- Upload the selection ---------------------------------------- */

  // Resolves the destination folder, ensures Year/DocType subfolders, then
  // uploads + tags each file in turn. Never throws for routing problems —
  // returns a runError instead, so the caller can render it in place.
  const runUpload = async (
    files: File[],
    selections: LevelSelection[],
    labels: {
      modeLabel: string;
      termSetGuid: string;
      docTypeLabel: string;
      yearLabel: string;
      confLabel: string;
      vendorLabel: string;
    },
    onState: (fileIndex: number, state: FileState, sru?: string) => void,
    onPct: (fileIndex: number, pct: number) => void,
  ): Promise<{ runError?: string; results: FileResult[] }> => {
    const leaf = selections[selections.length - 1];
    if (!leaf || !leaf.id) {
      return { runError: "No destination folder selected.", results: [] };
    }

    setStatus("Locating destination folder…");
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
        // Names BOTH causes. "Re-run reconciliation" alone was wrong half the time:
        // since folder names come from DMS Term Abbreviation, a unit with no
        // abbreviation row is skipped by every run, so re-running changes nothing.
        runError: `"${leaf.label}" has no folder yet. Your DMS administrator needs to give it an abbreviation in the DMS Term Abbreviation list, then run folder reconciliation.`,
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
        runError:
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
        `${webSru}/${libraryUrlSegment()}/`,
        "(resolved library title:",
        settings.stagingLibrary,
        ")",
      );
      return {
        runError:
          "Could not work out the Documents path for this unit folder. The upload library's " +
          "folder path did not match what was expected — ask an administrator to check the " +
          "library has not been renamed since the last deployment.",
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
        runError: docsProbe.confirmedMissing
          ? `The matching folder does not exist in Documents yet (${docsUnitPath}). Ask an administrator to run the reconciliation tool — it creates this folder with the correct permissions. Do not create it by hand: a hand-made folder inherits the library's root permissions and would widen access.`
          : `Could not verify the destination folder in Documents (HTTP ${docsProbe.status}). This usually means SharePoint was busy — most often because folder reconciliation is running at the same time. Wait for reconciliation to finish, then try again. The folder itself is probably fine.`,
        results: [],
      };
    }
    const docsUnitFolder = docsProbe.folder;

    const yearLabel = sanitizeFolderSegment(labels.yearLabel);
    const docTypeLabel = sanitizeFolderSegment(labels.docTypeLabel);
    if (!yearLabel || !docTypeLabel) {
      return { runError: "Year and Document Type are required.", results: [] };
    }

    setStatus("Preparing destination folders…");
    // Year / Document Type subfolders are safe to create — they inherit the unit
    // folder's ACL, exactly as they do in Staging.
    const yearFolder = await ensureFolder(
      context.spHttpClient,
      siteUrl,
      docsUnitFolder.serverRelativeUrl,
      yearLabel,
    );
    if (!yearFolder) {
      return { runError: `Could not create the "${yearLabel}" folder.`, results: [] };
    }
    const destFolder = await ensureFolder(
      context.spHttpClient,
      siteUrl,
      yearFolder.serverRelativeUrl,
      docTypeLabel,
    );
    if (!destFolder) {
      return {
        runError: `Could not create the "${docTypeLabel}" folder.`,
        results: [],
      };
    }
    const folderId = destFolder.uniqueId;

    // One metadata set for every file in the selection — that is the point of
    // this screen (spec 2026-08-03 §4).
    const allSelections: LevelSelection[] = selections.map((s) => ({ ...s }));
    allSelections.unshift({
      column: "BusinessSegment",
      label: labels.modeLabel,
      id: labels.termSetGuid,
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
    // Re-derived here rather than trusted from state: hiding the tick does not
    // clear it, so a user who ticks it and then changes the level would otherwise
    // stamp true on a level that never offers it. Mirrors Form.tsx.
    const privilegedApplies =
      settings.legallyPrivilegedFor !== "" &&
      confidentiality === settings.legallyPrivilegedFor;
    const formValues: Array<{ FieldName: string; FieldValue: string }> = [
      {
        FieldName: settings.columns.documentType,
        FieldValue: taxVal(labels.docTypeLabel, documentType),
      },
      {
        FieldName: settings.columns.yearPeriod,
        FieldValue: taxVal(labels.yearLabel, yearPeriod),
      },
      {
        FieldName: settings.columns.confidentiality,
        FieldValue: taxVal(labels.confLabel, confidentiality),
      },
      {
        FieldName: settings.columns.legallyPrivileged,
        FieldValue: privilegedApplies && legallyPrivileged ? "true" : "false",
      },
      {
        FieldName: settings.columns.documentDate,
        FieldValue: toSpDate(documentDate),
      },
      { FieldName: settings.columns.remark, FieldValue: remark.trim() },
      ...buildLevelFormValues(levelCols, allSelections),
    ];
    // Vendor is intentionally NOT written. The field is hidden in the UI (client
    // request, 2026-07-28) so `vendor` is always empty; the plumbing stays in
    // place so un-hiding the select is the only change needed to restore it.
    if (vendor) {
      formValues.push({
        FieldName: settings.columns.vendor,
        FieldValue: taxVal(labels.vendorLabel, vendor),
      });
    }

    /* ----- Sequential per-file upload; one failure never stops the run ---- */
    const out: FileResult[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      // Files keep their own names — no composition on this screen (spec §4).
      const finalName = file.name;
      onPct(i, 0);
      onState(i, "uploading");
      setStatus(`Uploading ${i + 1} of ${files.length} — ${finalName}`);

      // Duplicate probe. On a clash we ask the user whether to overwrite; No
      // skips just this file, Yes re-uploads with overwrite=true. Either way the
      // remaining files carry on — spec §4.
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
            out.push({
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
          out.push({ name: finalName, outcome: "failed", detail });
          onState(i, "failed");
          continue;
        }
        const uploadJson = JSON.parse(uploadRes.body);
        const uploadedSru = uploadJson.ServerRelativeUrl;
        onPct(i, 100);

        const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeServerRelativePath(uploadedSru)}'`,
          SPHttpClient.configurations.v1,
        );
        if (!itemRes.ok) {
          out.push({
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
          out.push({
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
          out.push({
            name: finalName,
            outcome: "tagFailed",
            detail: `Uploaded, but a field failed: ${fieldError.FieldName} — ${fieldError.ErrorMessage}`,
          });
          onState(i, "tagFailed", uploadedSru);
          continue;
        }

        out.push({ name: finalName, outcome: "uploaded" });
        onState(i, "done", uploadedSru);
      } catch (err) {
        console.error("Upload threw:", finalName, err);
        out.push({
          name: finalName,
          outcome: "failed",
          detail: err instanceof Error ? err.message : "Unexpected error.",
        });
        onState(i, "failed");
      }
    }

    return { results: out };
  };

  const handleUpload = async (): Promise<void> => {
    const missing = validate();
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }

    // Files upload under their original names, so two identically named files in
    // one selection would silently collide (first wins). Reject up front rather
    // than half-way through the run.
    const names = picked.map((p) => p.file.name);
    const collision = names.find(
      (n, i) => names.findIndex((o) => o.toLowerCase() === n.toLowerCase()) !== i,
    );
    if (collision) {
      showToast(
        `Two or more of the selected files are named "${collision}". Remove the duplicate first.`,
        "error",
      );
      return;
    }

    const mode = activeMode();
    if (!mode || mode.levels.length === 0) {
      showToast("No upload mode configured.", "error");
      return;
    }

    const selections: LevelSelection[] = mode.levels.map((lvl, i) => {
      const opt = (levelChoices[i] ?? []).find((o) => o.id === levelValues[i]);
      return {
        column: lvl.column,
        label: opt?.label ?? "",
        id: opt?.id ?? "",
        labelCol: lvl.labelCol,
        tidCol: lvl.tidCol,
      };
    });

    const files = picked.map((p) => p.file);
    const labels = {
      modeLabel: mode.label,
      termSetGuid: mode.termSetGuid,
      docTypeLabel:
        options.documentType.find((o) => o.id === documentType)?.label ?? "",
      yearLabel: options.yearPeriod.find((o) => o.id === yearPeriod)?.label ?? "",
      confLabel:
        options.confidentiality.find((o) => o.id === confidentiality)?.label ?? "",
      vendorLabel: options.vendor.find((o) => o.id === vendor)?.label ?? "",
    };

    setBusy(true);
    setResults(null);
    setRunError(null);
    // Seed the live progress list — every file starts "pending".
    setLive(
      files.map((f) => ({
        name: f.name,
        size: f.size,
        state: "pending" as FileState,
        pct: 0,
      })),
    );

    const onState = (
      fileIndex: number,
      state: FileState,
      sru?: string,
    ): void => {
      setLive((prev) =>
        prev
          ? prev.map((f, fi) =>
              fi === fileIndex ? { ...f, state, ...(sru ? { sru } : {}) } : f,
            )
          : prev,
      );
    };
    const onPct = (fileIndex: number, pct: number): void => {
      setLive((prev) =>
        prev ? prev.map((f, fi) => (fi === fileIndex ? { ...f, pct } : f)) : prev,
      );
    };

    try {
      const outcome = await runUpload(files, selections, labels, onState, onPct);
      setStatus("");
      if (outcome.runError) {
        // Routing failed before any file uploaded — show every row as failed.
        setLive((prev) =>
          prev ? prev.map((f) => ({ ...f, state: "failed" as FileState })) : prev,
        );
        setRunError(outcome.runError);
        setResults(outcome.results);
        return;
      }
      setResults(outcome.results);

      const okCount = outcome.results.filter(
        (r) => r.outcome === "uploaded",
      ).length;
      const anyProblem = outcome.results.some((r) => r.outcome !== "uploaded");
      if (!anyProblem && okCount > 0) {
        // Same dialog as the single-file Form. A toast was easy to miss at the end
        // of a long run, and after 50 files the uploader needs an explicit "that
        // worked" plus a way to go again without reloading the page.
        setDoneOpen(true);
        // The selection has been fully processed — empty it so Upload can't be
        // clicked again and re-send the same files (which would raise a replace
        // prompt per file). The progress list is deliberately left on screen.
        setPicked([]);
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

  // READY / LOADING rather than Uploaded / Uploading — the two states the
  // uploading list is meant to distinguish at a glance (spec §3).
  const fpLabel: Record<FileState, string> = {
    pending: "Waiting",
    uploading: "Loading",
    done: "Ready",
    skipped: "Skipped",
    failed: "Failed",
    tagFailed: "No tags",
    deleted: "Deleted",
  };

  // Overall progress. A file counts as "settled" once it reaches any terminal
  // state; the one file in flight contributes its own byte fraction so the bar
  // creeps forward instead of jumping a whole file at a time.
  const liveTotals = (live ?? []).reduce(
    (acc, f) => {
      acc.total += 1;
      if (f.state !== "pending" && f.state !== "uploading") acc.done += 1;
      else if (f.state === "uploading") acc.partial += f.pct / 100;
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

  const resultCounts = {
    uploaded: (results ?? []).filter((r) => r.outcome === "uploaded").length,
    skipped: (results ?? []).filter((r) => r.outcome === "skipped").length,
    failed: (results ?? []).filter((r) => r.outcome === "failed").length,
    tagFailed: (results ?? []).filter((r) => r.outcome === "tagFailed").length,
  };
  const problemRows = (results ?? []).filter((r) => r.outcome !== "uploaded");

  /* ---------- Render ------------------------------------------------------ */

  return (
    <section className="dms-form">
      <style>{`
        .dms-form { max-width: 960px; margin: 32px auto; padding: 0 24px 48px; font-family: 'Segoe UI', sans-serif; }
        .dms-subtitle { margin: 0 0 16px; font-size: 14px; color: #666; }
        .dms-warn { display: flex; gap: 10px; align-items: flex-start; background: #fff4e5; border: 1px solid #f0c070; border-radius: 6px; padding: 12px 14px; font-size: 13px; color: #7a4f00; margin: 0 0 24px; }
        .dms-section { background: #fff; border: 1px solid #e1e1e1; border-radius: 8px; padding: 24px; margin-bottom: 20px; }
        .dms-section-title { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #0f6c3f; margin: 0 0 16px; }
        /* Drop zone — same component and styling as the single-file Form's; the
           only functional difference is the multiple attribute on the input. */
        .dms-dropzone { display: flex; flex-direction: column; align-items: center; justify-content: center;
          gap: 8px; text-align: center; border: 1px dashed #9bbfaa; border-radius: 10px;
          background: #f2f8f4; padding: 24px 16px; cursor: pointer; font-size: 13px;
          transition: background .12s, border-color .12s; }
        .dms-dropzone:hover, .dms-dropzone:focus-visible { border-color: #0f6c3f; background: #eaf4ee; }
        /* .over fires on dragover — without a visible change there is no confirmation
           the browser will accept the drop, and users let go over the wrong element. */
        .dms-dropzone.over { border-color: #0f6c3f; border-style: solid; background: #e2efe7; }
        .dms-dropzone-icon { width: 32px; height: 32px; color: #0f6c3f; }
        .dms-dropzone .hint { color: #666; font-size: 12px; }
        .dms-link { background: none; border: none; color: #0f6c3f; cursor: pointer; font-weight: 600; padding: 0; font-size: 13px; }
        .dms-link:disabled { color: #9bbfaa; cursor: default; }
        /* Selected summary bar: count on the left, "Add more" on the right. Add more
           APPENDS — the one behaviour someone coming from the single-file form would
           guess wrong, which is why it is a distinct control. */
        .dms-selbar { display: flex; align-items: center; justify-content: space-between; gap: 12px;
          background: #eaf4ee; border: 1px solid #b3d9c4; border-radius: 10px; padding: 12px 16px;
          font-size: 13px; font-weight: 600; color: #0f6c3f; }
        /* Scrolls instead of paginating — the whole selection is always reachable. */
        .dms-filelist { margin-top: 12px; display: flex; flex-direction: column; gap: 8px; max-height: 320px; overflow-y: auto; overscroll-behavior: contain; padding-right: 6px; }
        .dms-filerow { display: flex; align-items: center; gap: 12px; border: 1px solid #ececec; border-radius: 8px; padding: 10px 12px; background: #fafafa; font-size: 13px; }
        .dms-filerow .fname { flex: 1 1 auto; min-width: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 600; color: #1b1b1b; }
        .dms-filerow .size { flex: 0 0 auto; font-size: 12px; color: #666; }
        .dms-remove { flex: 0 0 auto; background: none; border: none; cursor: pointer; color: #d13438; font-size: 15px; line-height: 1; padding: 4px; }
        .dms-remove:disabled { color: #c9a3a4; cursor: default; }
        /* Live per-file progress */
        .dms-progress-list { margin-top: 12px; max-height: 360px; overflow-y: auto; overscroll-behavior: contain; display: flex; flex-direction: column; gap: 6px; padding-right: 6px; }
        /* One row per file: status tag · name · size · live bar · % · remove */
        .dms-fp-row { display: flex; align-items: center; gap: 12px; font-size: 13px; padding: 10px 12px; border-radius: 6px; background: #fafafa; }
        /* No green wash on a finished row — the READY tag carries the state in
           text, and 50 green rows drown out the ones that need attention. */
        .dms-fp-row.skipped, .dms-fp-row.tagFailed { background: #fff4e5; }
        .dms-fp-row.failed { background: #fdf3f3; }
        .dms-fp-row.deleted { background: #f2f2f2; }
        .dms-fp-tag { flex: 0 0 74px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; color: #999; }
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
        /* Overall progress header */
        .dms-overall-head { display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; color: #333; margin-bottom: 6px; }
        .dms-overall-track { height: 10px; border-radius: 6px; background: #ececec; overflow: hidden; }
        .dms-overall-fill { height: 100%; background: #0f6c3f; border-radius: 6px; transition: width .3s ease; }
        .dms-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 16px; font-size: 13px; }
        .dms-field > span { font-weight: 600; }
        .dms-field .req { color: #d13438; font-style: normal; }
        .dms-field select, .dms-field input[type="text"], .dms-field input[type="date"] { padding: 8px 10px; border: 1px solid #c8c8c8; border-radius: 10px; font: inherit; width: 100%; box-sizing: border-box; height: 38px; background: #fff; }
        .dms-field small { color: #666; font-size: 12px; font-weight: 400; }
        /* Document Date | Confidential Level | Legally Privileged on ONE line, per
           the client mockup. A flex row rather than grid cells because the tick
           exists for only one confidentiality level: when it is absent, date and
           level fall back to an even split, so the row does not reflow around a
           control that is not there. */
        .dms-detail-row { display: flex; gap: 24px; align-items: flex-end; flex-wrap: wrap; }
        .dms-detail-row > .dms-field { flex: 1 1 200px; min-width: 0; }
        .dms-lp { flex: 0 0 auto; display: flex; align-items: center; gap: 8px; height: 38px; margin-bottom: 16px; font-size: 13px; font-weight: 600; white-space: nowrap; cursor: pointer; }
        .dms-lp input { accent-color: #0f6c3f; width: 16px; height: 16px; margin: 0; cursor: pointer; }
        /* Info tooltips sit on the LABEL, beside the field name. They used to be
           absolutely positioned against the select's right edge, which had to be
           re-tuned every time the grid changed; on the label there is nothing to
           drift against. */
        .dms-labelrow { display: flex; align-items: center; gap: 6px; }
        .dms-info { position: relative; flex: 0 0 auto; width: 18px; height: 18px; border-radius: 50%; border: 1.5px solid #0f6c3f; background: transparent; color: #0f6c3f; font-size: 12px; font-weight: 700; font-style: normal; display: inline-flex; align-items: center; justify-content: center; cursor: help; box-sizing: border-box; }
        .dms-info-panel { display: none; position: absolute; top: calc(100% + 8px); left: 0; z-index: 30; width: 280px; max-width: calc(100vw - 48px); padding: 16px; background: #fff; border: 1px solid #e1e1e1; border-radius: 10px; box-shadow: 0 4px 16px rgba(0,0,0,.12); cursor: default; text-align: left; font-weight: 400; }
        /* The rightmost icon on the row would push its panel past the card edge. */
        .dms-info.align-right .dms-info-panel { left: auto; right: 0; }
        .dms-info:hover .dms-info-panel, .dms-info:focus .dms-info-panel, .dms-info:focus-within .dms-info-panel { display: block; }
        .dms-info-panel dl { margin: 0; }
        .dms-info-panel dt { margin-top: 12px; color: #0f6c3f; font-size: 13px; font-weight: 700; }
        .dms-info-panel dt:first-of-type { margin-top: 0; }
        .dms-info-panel dd { margin: 4px 0 0; color: #444; font-size: 12px; font-weight: 400; line-height: 1.45; }
        .dms-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 24px; }
        /* Unit | Year | Document Type on one row. Together they name exactly one
           destination folder, so they read better as a set than stacked. */
        .dms-grid-3 { grid-template-columns: 1.8fr 0.9fr 1.3fr; }
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
        /* Results */
        .dms-results { margin-top: 20px; border: 1px solid #e1e1e1; border-radius: 8px; padding: 16px; background: #fff; }
        .dms-results-summary { font-size: 13px; font-weight: 600; margin: 0 0 10px; }
        .dms-run-error { background: #fdf3f3; color: #d13438; border: 1px solid #f1c0c0; border-radius: 4px; padding: 10px 14px; font-size: 13px; }
        .dms-result { display: flex; gap: 10px; align-items: baseline; font-size: 13px; padding: 7px 10px; border-radius: 4px; margin-bottom: 6px; }
        .dms-result .tag { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }
        .dms-result .fname { word-break: break-all; }
        .dms-result .why { color: #666; font-size: 12px; }
        .dms-result.skipped { background: #fff4e5; } .dms-result.skipped .tag { color: #7a4f00; }
        .dms-result.failed { background: #fdf3f3; } .dms-result.failed .tag { color: #d13438; }
        .dms-result.tagFailed { background: #fff4e5; } .dms-result.tagFailed .tag { color: #7a4f00; }
        /* Popups */
        .dms-popup-overlay { position: fixed; inset: 0; background: rgba(0,0,0,.25); backdrop-filter: blur(6px); -webkit-backdrop-filter: blur(6px); z-index: 9998; display: flex; align-items: center; justify-content: center; padding: 16px; }
        .dms-popup { background: #fff; border-radius: 16px; padding: 40px 40px 32px; text-align: center; max-width: 420px; width: 100%; box-shadow: 0 8px 40px rgba(0,0,0,.15); }
        .dms-popup-svg { width: 110px; height: 110px; display: block; margin: 0 auto 20px; }
        .dms-popup-title { font-size: 22px; font-weight: 700; color: #0f6c3f; margin: 0 0 12px; }
        .dms-popup-msg { font-size: 14px; color: #555; margin: 0 0 28px; line-height: 1.6; }
        /* Stacked buttons for the success dialog — matches the Form. */
        .dms-popup-stack { display:flex; flex-direction: column; align-items:center; gap: 10px; }
        .dms-popup-stack .dms-popup-btn { width: 100%; max-width: 200px; }
        .dms-popup-actions { display:flex; align-items:center; justify-content:center; gap: 12px; }
        .dms-popup-btn { min-width: 96px; border-radius: 6px; padding: 11px 22px; font-size: 14px; font-weight: 600; font-family: inherit; cursor: pointer; }
        .dms-popup-btn.confirm { background: #0f6c3f; color: #fff; border: none; }
        .dms-popup-btn.confirm:hover { background: #0a5230; }
        .dms-popup-btn.cancel { background: #fff; color: #0f6c3f; border: 1px solid #0f6c3f; }
        .dms-popup-btn.cancel:hover { background: #f0f6f2; }
        .dms-toast { position: fixed; top: 24px; right: 24px; z-index: 9999; min-width: 300px; max-width: 460px; padding: 14px 40px 14px 16px; border-radius: 6px; font-size: 13px; font-family: 'Segoe UI', sans-serif; box-shadow: 0 4px 16px rgba(0,0,0,.18); animation: dms-slidein .2s ease; }
        .dms-toast.error { background: #d13438; color: #fff; }
        .dms-toast.success { background: #0f6c3f; color: #fff; }
        .dms-toast-close { position: absolute; top: 10px; right: 12px; background: none; border: none; cursor: pointer; font-size: 16px; color: inherit; opacity: .7; line-height: 1; }
        .dms-toast-close:hover { opacity: 1; }
        @keyframes dms-slidein { from { transform: translateX(60px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        @media (max-width: 640px) {
          .dms-grid, .dms-grid-3 { grid-template-columns: 1fr; }
          .dms-toast { left: 12px; right: 12px; min-width: unset; top: 12px; }
        }
      `}</style>

      <p className="dms-subtitle">
        Up to {MAX_FILES} documents in one go, all filed to the same folder with
        the same details. All fields marked <strong>*</strong> are required.
      </p>
      <div className="dms-warn">
        <span aria-hidden="true">⚠</span>
        <span>
          <strong>Temporary tool.</strong> Files go straight into the{" "}
          <strong>Documents</strong> library — they skip Staging and the approval
          step entirely, and are visible to everyone with access to the
          destination folder as soon as they upload. Files keep the names they
          already have, so name them before uploading.
        </span>
      </div>

      {/* ── Documents Folder Information ─────────────────────────────────── */}
      {/* Deliberately BEFORE Documents Details in SOURCE order, not reordered
          with CSS: tab order follows the DOM. Mirrors Form.tsx. */}
      <div className="dms-section">
        <p className="dms-section-title">Documents Folder Information</p>

        {deptLoading ? (
          <p className="dms-dept-loading">Loading your access&hellip;</p>
        ) : !privileged && validPaths.length === 0 ? (
          <div className="dms-dept-error">
            Your account isn&apos;t fully provisioned to upload — you need
            membership at every level plus the unit uploader role. Contact your
            administrator.
          </div>
        ) : null}

        {/* Business Segment | Project toggle — a side shows only if the user can
            actually upload there (privileged users see every configured side). */}
        <div className="dms-radio-group">
          <p>Upload to</p>
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
                  disabled={busy}
                  onChange={() => switchMode(offerable[0].key)}
                />
                {side === "BusinessSegment" ? "Business Segment" : "Project"}
              </label>
            );
          })}
        </div>

        {/* Segment picker. Shown whenever a segment is offered, even if there is
            only one — a single-option select still answers "where am I". */}
        {(() => {
          const side = activeMode()?.side;
          const sideModes = modes.filter((m) => m.side === side);
          const offerable = privileged
            ? sideModes
            : sideModes.filter((m) =>
                validPaths.some((p) => p.modeKey === m.key),
              );
          if (offerable.length === 0) return null;
          return (
            <label className="dms-field">
              <span>
                Segment <em className="req">*</em>
              </span>
              <select
                value={uploadMode}
                disabled={busy}
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

        {/* Restricted users: read-only breadcrumb of the resolved location. */}
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

        <div className="dms-grid dms-grid-3">
          {/* File-path fields, in folder order: Segment (above) -> level(s) ->
              Year -> Document Type. Intermediate levels span the full row; the
              deepest one shares a row of three with Year and Document Type,
              which is the set that identifies a single destination folder. */}
          {(activeMode()?.levels ?? []).map((lvl, i, arr) =>
            renderSelect(
              lvl.label,
              true,
              levelValues[i] ?? "",
              (v) => {
                const md = activeMode();
                if (md) handleLevelChange(md, i, v);
              },
              levelChoices[i] ?? [],
              busy ||
                deptLoading ||
                isLevelLocked(i) ||
                (i > 0 && !levelValues[i - 1]),
              `Select ${lvl.label}`,
              i < arr.length - 1,
            ),
          )}

          {renderSelect(
            "Year",
            true,
            yearPeriod,
            setYearPeriod,
            options.yearPeriod,
            busy,
          )}

          {renderSelect(
            "Document Type",
            true,
            documentType,
            setDocumentType,
            options.documentType,
            busy,
          )}

          {/* Graceful empty-state: a segment whose term set has no child terms yet
              (the non-GHO Head Offices before their Department/Unit trees are
              added) would otherwise show a dead "Select …" dropdown. */}
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

          {/* Spans the row, matching the Form. One shared remark for the whole
              selection, like every other field on this screen. */}
          <label className="dms-field" style={{ gridColumn: "1 / -1" }}>
            <span>Remark</span>
            <input
              type="text"
              value={remark}
              maxLength={250}
              disabled={busy}
              onChange={(e) => setRemark(e.target.value)}
            />
            <small>Max. 250 characters</small>
          </label>
        </div>
      </div>

      {/* ── Documents Details ───────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Documents Details</p>

        {/* The file area has three states: empty drop zone, selected list, and
            live upload progress. Spec 2026-08-03 §3. */}
        {settings.allowedFileTypes.kind === "none" ? (
          /* An empty AllowedFileTypes selection is a hard block, not a silent
             fallback — spec 2026-07-30 §3. The message names the column and the
             list because the client is the one who fixes it, in one click. */
          <div className="dms-dropzone" style={{ opacity: 0.6 }}>
            <span>{NO_TYPES_MESSAGE}</span>
          </div>
        ) : live ? (
          <>
            <div className="dms-overall-head">
              <span>
                {liveTotals.done} of {liveTotals.total} document
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
            {/* Completed rows stay visible and keep their order, so the header
                count can be checked against the list. */}
            <div className="dms-progress-list">
              {live.map((f, fi) => {
                const canDelete =
                  (f.state === "done" || f.state === "tagFailed") && !!f.sru;
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
                        disabled={busy || deletingKey === String(fi)}
                        onClick={() => {
                          deleteUploaded(fi, f.sru as string, f.name).catch(
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
          </>
        ) : picked.length === 0 ? (
          /* Click anywhere to open the picker, or drop files on it. */
          <div
            className={`dms-dropzone${dragOver ? " over" : ""}`}
            role="button"
            tabIndex={0}
            onClick={() => fileRef.current?.click()}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                fileRef.current?.click();
              }
            }}
            // preventDefault on dragOver is what makes the element a valid drop
            // target; without it the browser navigates to the file instead.
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              addFiles(e.dataTransfer?.files ?? null);
            }}
          >
            {/* Inlined rather than imported: an <img> would need an asset loader
                and a second network request for a 20-line glyph. */}
            <svg className="dms-dropzone-icon" viewBox="0 0 24 24" aria-hidden="true">
              <path d="M12 3v10m0 0 4-4m-4 4-4-4" fill="none" stroke="currentColor"
                strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              <path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" fill="none"
                stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
            </svg>
            <span>
              <span className="dms-link">Choose multiple documents</span> or drop
              them here
            </span>
            <span className="hint">Max. {MAX_FILES} files.</span>
          </div>
        ) : (
          <>
            <div className="dms-selbar">
              <span>
                {picked.length} document{picked.length === 1 ? "" : "s"} selected
              </span>
              <button
                type="button"
                className="dms-link"
                disabled={busy}
                onClick={() => fileRef.current?.click()}
              >
                Add more
              </button>
            </div>
            <div className="dms-filelist">
              {picked.map((p) => (
                <div className="dms-filerow" key={p.key}>
                  <span className="fname" title={p.file.name}>
                    {p.file.name}
                  </span>
                  <span className="size">{fmtSize(p.file.size)}</span>
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
          </>
        )}

        {/* One input serves the drop zone and Add more. `multiple` is the only
            functional difference from the Form's single-file picker; `accept`
            filters the DIALOG only, which is why addFiles re-checks. */}
        <input
          ref={fileRef}
          type="file"
          multiple
          accept={
            // undefined rather than "": an empty accept attribute means "no
            // filter" and would offer every file in the dialog. addFiles blocks
            // them anyway, but not offering them is clearer.
            settings.allowedFileTypes.kind === "none"
              ? undefined
              : settings.allowedFileTypes.types.join(",")
          }
          style={{ display: "none" }}
          onChange={(e) => addFiles(e.target.files)}
        />

        <div className="dms-detail-row" style={{ marginTop: 16 }}>
          <label className="dms-field">
            <span>
              Document Date <em className="req">*</em>
            </span>
            <input
              type="date"
              value={documentDate}
              disabled={busy}
              max={(() => {
                const d = new Date();
                const mm = d.getMonth() + 1;
                const day = d.getDate();
                return `${d.getFullYear()}-${mm < 10 ? "0" + mm : mm}-${day < 10 ? "0" + day : day}`;
              })()}
              onChange={(e) => setDocumentDate(e.target.value)}
            />
          </label>

          {/* Not renderSelect: the info icon belongs on the LABEL, and the label
              is a <span> holding a real <label htmlFor> so clicking the icon does
              not fall through and focus the select. */}
          <div className="dms-field">
            <span className="dms-labelrow">
              <label htmlFor="dms-bulk-conf">
                Confidential Level <em className="req">*</em>
              </label>
              <em
                className="dms-info"
                tabIndex={0}
                role="button"
                aria-label="What the confidentiality levels mean"
              >
                i
                <span className="dms-info-panel" role="tooltip">
                  {/* Legally Privileged is NOT defined here any more — it has its
                      own control and its own icon beside it, and defining it in two
                      places invites the two texts to drift apart.
                      Highly Confidential is deliberately absent too: its term is
                      removed from the term store for Phase 1, so the dropdown
                      cannot offer it. */}
                  <dl>
                    <dt>Confidential</dt>
                    <dd>
                      This applies to sensitive business information that is
                      intended strictly for use within the Group, on a need-to-know
                      basis.
                    </dd>
                    <dt>Restricted</dt>
                    <dd>
                      This applies to business information that may be disclosed to
                      external parties only if a non-disclosure agreement has been
                      signed.
                    </dd>
                  </dl>
                </span>
              </em>
            </span>
            <select
              id="dms-bulk-conf"
              value={confidentiality}
              disabled={busy}
              title={
                options.confidentiality.find((o) => o.id === confidentiality)
                  ?.label ?? ""
              }
              onChange={(e) => setConfidentiality(e.target.value)}
            >
              <option value="">--</option>
              {options.confidentiality.map((o) => (
                <option key={o.id} value={o.id} title={o.label}>
                  {o.label}
                </option>
              ))}
            </select>
          </div>

          {/* Offered only for the level named by `legallyPrivilegedFor` in DMS
              Config. Unset means never offered. The value is re-derived at upload
              time rather than trusted from here, because hiding the control does
              not clear the state behind it. */}
          {settings.legallyPrivilegedFor !== "" &&
            confidentiality === settings.legallyPrivilegedFor && (
              <>
                <label className="dms-lp">
                  <input
                    type="checkbox"
                    checked={legallyPrivileged}
                    disabled={busy}
                    onChange={(e) => setLegallyPrivileged(e.target.checked)}
                  />
                  <span>Legally Privileged</span>
                </label>
                <em
                  className="dms-info align-right"
                  tabIndex={0}
                  role="button"
                  aria-label="What Legally Privileged means"
                  style={{ marginBottom: 26 }}
                >
                  i
                  <span className="dms-info-panel" role="tooltip">
                    This applies to confidential communications (email, advice,
                    documents, conversations) between client and lawyer that are
                    protected by law from being disclosed in a court of law or
                    during legal proceedings.
                  </span>
                </em>
              </>
            )}
        </div>
      </div>

      {/* ── Actions ─────────────────────────────────────────────────────── */}
      <div className="dms-actions">
        <button
          type="button"
          className="dms-btn secondary"
          onClick={resetForm}
          disabled={busy}
        >
          Cancel
        </button>
        <button
          type="button"
          className="dms-btn primary"
          onClick={() => {
            handleUpload().catch(() => undefined);
          }}
          disabled={busy || deptLoading || picked.length === 0}
        >
          {busy
            ? "Uploading…"
            : `Upload${picked.length > 0 ? ` (${picked.length})` : ""}`}
        </button>
      </div>

      {status && <p className="dms-status">{status}</p>}

      {/* ── Results ─────────────────────────────────────────────────────── */}
      {results && (
        <div className="dms-results">
          {runError ? (
            <div className="dms-run-error">{runError}</div>
          ) : (
            <>
              <p className="dms-results-summary">
                {resultCounts.uploaded} uploaded
                {resultCounts.skipped > 0 && `, ${resultCounts.skipped} skipped`}
                {resultCounts.failed > 0 && `, ${resultCounts.failed} failed`}
                {resultCounts.tagFailed > 0 &&
                  `, ${resultCounts.tagFailed} uploaded without tags`}
                .
                {resultCounts.tagFailed > 0 &&
                  " Files listed as “uploaded, not tagged” are already in Documents — fix their metadata in the library rather than re-uploading."}
              </p>
              {problemRows.map((r, i) => (
                <div className={`dms-result ${r.outcome}`} key={`${r.name}-${i}`}>
                  <span className="tag">{outcomeLabel[r.outcome]}</span>
                  <span className="fname">{r.name}</span>
                  {r.detail && <span className="why">{r.detail}</span>}
                </div>
              ))}
            </>
          )}
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

      {doneOpen && (
        <div className="dms-popup-overlay" role="dialog" aria-modal="true">
          <div className="dms-popup">
            <svg
              className="dms-popup-svg"
              viewBox="0 0 184 184"
              fill="none"
              xmlns="http://www.w3.org/2000/svg"
            >
              <circle opacity="0.3" cx="92.0001" cy="92" r="75.4872" fill="#14C7A5" />
              <circle cx="92" cy="92" r="92" fill="#14C7A5" fillOpacity="0.2" />
              <circle cx="92.0003" cy="91.9998" r="61.3333" fill="white" stroke="#14C7A5" strokeWidth="3" />
              <path d="M67 90.7143L88.4286 110L117 74" stroke="#14C7A5" strokeWidth="10" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            <p className="dms-popup-title">Upload Successful</p>
            <p className="dms-popup-msg">
              Your document is awaiting approval.
              <br />
              Please visit Home page to track progress.
            </p>
            <div className="dms-popup-stack">
              <button
                className="dms-popup-btn confirm"
                onClick={() => {
                  window.location.href = siteUrl;
                }}
              >
                Back to Document
              </button>
              <button
                className="dms-popup-btn cancel"
                onClick={() => {
                  setDoneOpen(false);
                  resetForm();
                }}
              >
                Upload More
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
