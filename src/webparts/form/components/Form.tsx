import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IFormProps } from "./IFormProps";
import {
  lookupFolderMapping,
  resolveFolderServerUrl,
  ensureFolder,
  encodeServerRelativePath,
} from "../../../shared/dmsFolderMap";
import { formatFileSize } from "../../../shared/fileSize";
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
  CONFIG_UNREADABLE_MESSAGE,
  FALLBACK_FILE_TYPES,
  NO_TYPES_MESSAGE,
  readAllowedFileTypesField,
  resolveAllowedFileTypes,
} from "../../../shared/allowedFileTypes";

/* ----------------------------------------------------------------------------
 * CONFIG — hardcoded values are fallbacks only; live values load from DMS Config SP list
 * -------------------------------------------------------------------------- */

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
  // Free-text project name, written on every upload regardless of segment.
  // Distinct from the Group-led Projects "Group Project Name" folder level.
  projectName: "ProjectName",
  details: "_ExtendedDescription",
  // Free-text note from the uploader, shown to the approver. Distinct from
  // `details`/_ExtendedDescription, which is the built-in document Description.
  remark: "Remark",
  // Yes/No. Only meaningful when Confidentiality is the level named by the
  // `legallyPrivilegedFor` setting; written as "false" otherwise so a replaced
  // file cannot inherit a stale true from the document it overwrote.
  legallyPrivileged: "LegallyPrivileged",
};

// FALLBACK map: logical Levels `column` key -> the two real Staging internal
// names (label + term GUID). The preferred, portable source is each level's
// own labelCol/tidCol in DMS Config Levels JSON — those win over this table
// (see buildLevelFormValues), so a client tenant needs no code change here.
// Verified against the live /fields API (GHO pilot): SharePoint did NOT append
// "_Tid" — the GUID columns are BusinessSegmentTid / DepartmentTid / UnitTid
// (spaces stripped). BusinessSegment is injected by the form (not a config
// level), so it must stay here. Add the other segments' verified names too if
// you want them to work from the offline DEFAULT_MODES fallback.
const LEVEL_COLUMNS: Record<string, ColumnPair> = {
  BusinessSegment: { label: "Business_x0020_Segment", tid: "BusinessSegmentTid" },
  Department: { label: "Department", tid: "DepartmentTid" },
  Unit: { label: "Unit", tid: "UnitTid" },
};

const ILLEGAL_NAME_CHARS = /[\\/:*?"<>|#%&{}~]/g;

const getExtension = (name: string): string => {
  const dot = name.lastIndexOf(".");
  return dot >= 0 ? name.slice(dot) : "";
};

const buildUploadName = (originalName: string, typed: string): string => {
  const cleaned = typed.trim();
  if (!cleaned) return originalName;
  const ext = getExtension(originalName);
  let base = cleaned;
  const typedExt = getExtension(base);
  if (typedExt) base = base.slice(0, base.length - typedExt.length);
  base = base.replace(ILLEGAL_NAME_CHARS, "").replace(/\s+/g, " ").trim();
  return base ? `${base}${ext}` : originalName;
};

// validateUpdateListItem validates dates against the SITE's regional settings.
// This tenant is US locale (M/D/YYYY) — ISO YYYY-MM-DD is rejected.
const toSpDate = (iso: string): string => {
  const [y, m, d] = iso.split("-");
  return `${Number(m)}/${Number(d)}/${y}`;
};

// Document Date (ISO YYYY-MM-DD) -> DD-MM-YY for the auto-composed document name.
const dateDDMMYY = (iso: string): string => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  if (!y || !m || !d) return "";
  return `${d}-${m}-${y.slice(2)}`;
};

// Auto-compose the document name from Project Name + Vendor + Document Date:
// "<Project>-<Vendor>-<DD-MM-YY>", following the order the fields appear on the
// form. Any empty part is omitted, so a blank project still yields
// "<Vendor>-<DD-MM-YY>". Used until the user manually edits the name.
const composeDocName = (project: string, vendor: string, iso: string): string =>
  [project.trim(), vendor.trim(), dateDDMMYY(iso)].filter(Boolean).join("-");

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

// A fully authorised upload path for a restricted (non-privileged) user:
// the mode plus the resolved term chain [top … leaf], every tier of which the
// user is a member of, with the leaf held under the UPL role.
type ValidPath = { modeKey: string; chain: TermOption[] };

// Vendor is NOT here — it is free text (see the Vendor/Customer Name input), so it
// has no term set and no options list.
type OptionMap = {
  documentType: TermOption[];
  yearPeriod: TermOption[];
  confidentiality: TermOption[];
};

const EMPTY_OPTIONS: OptionMap = {
  documentType: [],
  yearPeriod: [],
  confidentiality: [],
};

// Fallback if DMS Config is missing/unreachable. DMS Config is the source of
// truth at runtime — these built-in modes only serve an offline fallback.
//   - The four Head Office segments (2026 pilot) have real term-set GUIDs.
//   - The 2027 segments carry REPLACE-* placeholder GUIDs: swap them for the real
//     term-set GUIDs (or, preferably, drive everything from DMS Config so no
//     code edit is needed). Each level's real Staging column internal names are
//     supplied per-level via DMS Config Levels JSON (labelCol/tidCol); the
//     logical `column` key here resolves via LEVEL_COLUMNS as a fallback only.
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
      // GroupProjectName/GroupProjectNameTid are not created yet — not needed
      // for the 2026 Head Office pilot.
      { label: "Group Project Name", column: "GroupProjectName" },
      { label: "Department", column: "Department" },
      { label: "Unit", column: "Unit" },
    ],
    sortOrder: 8,
  },
];

type DmsSettings = {
  // No `vendor` term set — Vendor/Customer Name is free text.
  termSets: {
    documentType: string;
    yearPeriod: string;
    confidentiality: string;
  };
  // Metadata column INTERNAL names. Portable: a new site sets these in DMS Config
  // (col_* setting rows) instead of editing code. Defaults below match the current
  // site so it works with no config rows. Department/Unit level columns are config-
  // driven separately via DMS Config Levels JSON (labelCol/tidCol).
  columns: {
    documentType: string;
    yearPeriod: string;
    documentDate: string;
    confidentiality: string;
    vendor: string;
    projectName: string;
    remark: string;
    legallyPrivileged: string;
    businessSegmentLabel: string;
    businessSegmentTid: string;
  };
  stagingLibrary: string;
  // Term GUID of the ONE confidentiality level that offers the Legally Privileged
  // tick. Config-driven rather than hardcoded because the levels are term-store
  // data: renaming or re-creating a level changes its GUID, and a hardcoded value
  // would silently stop offering the tick with nothing in the UI to explain why.
  // Empty means "never offer it", which is the correct behaviour for a site that
  // has not configured the row — better than guessing a level.
  legallyPrivilegedFor: string;
  allowedFileTypes: AllowedFileTypes;
};

const DEFAULT_SETTINGS: DmsSettings = {
  termSets: {
    documentType: "866c5754-258e-401f-8685-03d20ae59b1d",
    yearPeriod: "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf",
    confidentiality: "0d6d1da8-27e5-477f-8684-e8cf169f8fb9",
  },
  columns: {
    documentType: FIELDS.documentType,
    yearPeriod: FIELDS.yearPeriod,
    documentDate: FIELDS.documentDate,
    confidentiality: FIELDS.confidentiality,
    vendor: FIELDS.vendor,
    projectName: FIELDS.projectName,
    remark: FIELDS.remark,
    legallyPrivileged: FIELDS.legallyPrivileged,
    businessSegmentLabel: LEVEL_COLUMNS.BusinessSegment.label,
    businessSegmentTid: LEVEL_COLUMNS.BusinessSegment.tid,
  },
  stagingLibrary: "Staging",
  // Empty by default: the tick appears only once a site sets legallyPrivilegedFor
  // to a confidentiality term GUID. Defaulting to a guessed level would offer a
  // legal flag under the wrong heading.
  legallyPrivilegedFor: "",
  // "unknown", not "configured": reaching this constant means DMS Config could not
  // be read, and the UI must say so rather than present these as configured values.
  // The old list here was [".pdf", ".xls", ".xlsx"] — missing .doc/.docx, which is
  // part of why the silent fallback was so confusing to diagnose on 2026-07-30.
  allowedFileTypes: { kind: "unknown", types: FALLBACK_FILE_TYPES },
};

export default function Form({ context }: IFormProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const fileRef = useRef<HTMLInputElement>(null);

  const [options, setOptions] = useState<OptionMap>(EMPTY_OPTIONS);
  const [deptLoading, setDeptLoading] = useState<boolean>(true);
  // Privileged = site admin OR a GLOBAL-role uploader: bypasses tier detection
  // and gets the full manual cascade (may upload anywhere).
  const [privileged, setPrivileged] = useState<boolean>(false);

  // Generic N-level cascade state: one option list + one selected term id per level.
  const [levelChoices, setLevelChoices] = useState<TermOption[][]>([]);
  const [levelValues, setLevelValues] = useState<string[]>([]);
  // Restricted users: their fully-authorised upload paths (segment→…→leaf).
  const [validPaths, setValidPaths] = useState<ValidPath[]>([]);

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [uploadMode, setUploadMode] = useState<string>("");
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);
  const [file, setFile] = useState<File | undefined>(undefined);
  const [docName, setDocName] = useState<string>("");
  // True once the user manually edits the document name — stops the Vendor+Date
  // auto-composition from overwriting their custom text (reset when they clear it).
  const [docNameEdited, setDocNameEdited] = useState<boolean>(false);
  const [documentType, setDocumentType] = useState<string>("");
  const [yearPeriod, setYearPeriod] = useState<string>("");
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [vendor, setVendor] = useState<string>("");
  const [projectName, setProjectName] = useState<string>("");
  const [remark, setRemark] = useState<string>("");
  const [legallyPrivileged, setLegallyPrivileged] = useState<boolean>(false);
  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [toast, setToast] = useState<{
    message: string;
    type: ToastType;
  } | null>(null);
  const toastTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Replace-file guard. When a name clash is found we pause the upload and ask
  // the user whether to overwrite. The modal resolves a promise so handleUpload
  // can stay a single linear flow. Uploads are sequential, so one shared prompt
  // is safe.
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

  const showToast = (message: string, type: ToastType): void => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToast({ message, type });
    toastTimerRef.current = setTimeout(() => setToast(null), 5000);
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
    if (!res.ok)
      throw new Error(`Term set ${termSetId} returned ${res.status}`);
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

  // Two field lists: a site whose DMS Config predates the AllowedFileTypes column
  // answers HTTP 400 to the ENTIRE request, which would drop every setting on it —
  // term-set GUIDs, stagingLibrary, column names — not just the file types. So we
  // retry without the new field. Spec 2026-07-30 §5.
  const SETTINGS_FIELDS = "Title,SettingValue,AllowedFileTypes";
  const SETTINGS_FIELDS_LEGACY = "Title,SettingValue";

  const fetchSettingRows = async (
    select: string,
  ): Promise<SPHttpClientResponse> =>
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=${select}&$filter=ConfigType eq 'setting'`,
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
    // an emptied multi-choice field, so reading the value alone cannot tell
    // "unticked" (-> none, hard block) from "no such column" (-> unknown,
    // fallback). Verified against the live list 2026-07-30.
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
      },
      columns: {
        documentType: get("col_documentType") ?? DEFAULT_SETTINGS.columns.documentType,
        yearPeriod: get("col_yearPeriod") ?? DEFAULT_SETTINGS.columns.yearPeriod,
        documentDate: get("col_documentDate") ?? DEFAULT_SETTINGS.columns.documentDate,
        confidentiality:
          get("col_confidentiality") ?? DEFAULT_SETTINGS.columns.confidentiality,
        vendor: get("col_vendor") ?? DEFAULT_SETTINGS.columns.vendor,
        projectName: get("col_projectName") ?? DEFAULT_SETTINGS.columns.projectName,
        remark: get("col_remark") ?? DEFAULT_SETTINGS.columns.remark,
        legallyPrivileged:
          get("col_legallyPrivileged") ?? DEFAULT_SETTINGS.columns.legallyPrivileged,
        businessSegmentLabel:
          get("col_businessSegment") ?? DEFAULT_SETTINGS.columns.businessSegmentLabel,
        businessSegmentTid:
          get("col_businessSegmentTid") ?? DEFAULT_SETTINGS.columns.businessSegmentTid,
      },
      stagingLibrary: get("stagingLibrary") ?? DEFAULT_SETTINGS.stagingLibrary,
      legallyPrivilegedFor:
        get("legallyPrivilegedFor") ?? DEFAULT_SETTINGS.legallyPrivilegedFor,
      // SettingValue is deliberately NOT consulted for file types any more —
      // AllowedFileTypes is the single source of truth. Spec 2026-07-30 §3.
      allowedFileTypes: resolveAllowedFileTypes(rawFileTypes),
    };
  };

  /* ---------- Identity: group ids + admin --------------------------------- */

  // Detect the user's native SharePoint site-group Ids (direct membership only).
  // Matched as strings against DMS Group Map GroupId (SP group integer Id, e.g. "27").
  // No Graph, no admin consent. Nested Entra groups inside an SP group are NOT
  // returned — the DMS model adds users to site groups directly by design.
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

  /* ---------- Cascade builders -------------------------------------------- */

  // Load level-0 options (top terms of the mode's set); reset deeper levels.
  const initCascade = async (mode: UploadMode): Promise<void> => {
    const tops = await loadTermSet(mode.termSetGuid).catch(
      () => [] as TermOption[],
    );
    setLevelChoices([tops]);
    setLevelValues([]);
  };

  // When level `idx` changes to `termId`, load level idx+1 options and truncate below.
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

  // Hard-check every uploader leaf against the term tree: keep only the leaves
  // whose whole ancestor chain (and, for Business Segment modes, the segment
  // itself) the user is a member of. Returns the authorised paths.
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

  // Derive per-level options + values for a restricted user from their valid
  // paths: at each tier, offer only the terms their paths allow given the
  // choices so far; a tier with a single option is auto-selected (locked).
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

  // A tier is locked when the user's paths leave exactly one option for it.
  const isLevelLocked = (i: number): boolean =>
    !privileged && (levelChoices[i]?.length ?? 0) <= 1;

  /* ---------- Init -------------------------------------------------------- */

  useEffect(() => {
    const init = async (): Promise<void> => {
      const [loadedSettings, rawModes, groupMap, userGroupIds, admin] =
        await Promise.all([
          // Log the real failure. A silent fallback here is indistinguishable from
          // success and cost four rounds of diagnosis on 2026-07-30. Gotcha #9.
          loadSettings().catch((err) => {
            console.error(
              "DMS Config settings read failed — using built-in defaults.",
              err,
            );
            return DEFAULT_SETTINGS;
          }),
          loadModes().catch((err) => {
            console.error(
              "DMS Config mode rows read failed — using built-in modes.",
              err,
            );
            return DEFAULT_MODES;
          }),
          loadGroupMap().catch((err) => {
            console.error(
              "DMS Group Map read failed — no authorised upload paths.",
              err,
            );
            return [] as GroupMapRow[];
          }),
          loadUserGroupIds(),
          loadIsAdmin(),
        ]);
      setSettings(loadedSettings);
      // Ignore config rows that predate the Side/Levels schema (empty Levels) —
      // fall back to the built-in modes so the form never renders a broken cascade.
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

      // "unknown" means the config could not be read, or this site has no
      // AllowedFileTypes column — an admin problem, not something an uploader can
      // fix, so the toast is admin-only while the console line is always written.
      // Distinct from "none", which the file card handles. Spec 2026-07-30 §6.
      if (loadedSettings.allowedFileTypes.kind === "unknown") {
        console.warn(
          "AllowedFileTypes not supplied by DMS Config — running on built-in types:",
          loadedSettings.allowedFileTypes.types.join(", "),
        );
        if (isPrivileged) showToast(CONFIG_UNREADABLE_MESSAGE, "error");
      }

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
      });

      if (isPrivileged) {
        // Full manual cascade over the whole term set; default to first BS mode.
        const bsModes = loadedModes.filter(
          (m: UploadMode) => m.side === "BusinessSegment",
        );
        const defaultMode = bsModes[0] ?? loadedModes[0];
        if (defaultMode) {
          setUploadMode(defaultMode.key);
          await initCascade(defaultMode);
        }
      } else {
        // Restricted: resolve the user's authorised paths and lock the cascade.
        const paths = await resolveValidPaths(loadedModes, membership);
        setValidPaths(paths);
        const offerable = new Set(paths.map((p) => p.modeKey));
        const defaultMode =
          loadedModes.find(
            (m: UploadMode) =>
              m.side === "BusinessSegment" && offerable.has(m.key),
          ) ?? loadedModes.find((m: UploadMode) => offerable.has(m.key));
        if (defaultMode) {
          setUploadMode(defaultMode.key);
          applyRestrictedMode(paths, defaultMode, []);
        }
      }
      setDeptLoading(false);
    };

    init().catch((err) => {
      console.error("Form init failed:", err);
      showToast("Could not load form data. Please refresh the page.", "error");
      setDeptLoading(false);
    });
  }, []);

  /* ---------- Helpers ----------------------------------------------------- */

  const switchMode = (modeKey: string): void => {
    setUploadMode(modeKey);
    setStatus("");
    const mode = modes.find((m) => m.key === modeKey);
    if (!mode) return;
    if (privileged) {
      initCascade(mode).catch(() => {
        setLevelChoices([]);
        setLevelValues([]);
      });
    } else {
      applyRestrictedMode(validPaths, mode, []);
    }
  };

  // Level dropdown change: free choice for privileged (walk the term store);
  // for restricted users, re-derive the locked cascade from their valid paths.
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

  const toTaxValue = (opts: TermOption[], id: string): string => {
    const match = opts.find((o) => o.id === id);
    return match ? `${match.label}|${match.id}` : "";
  };

  // Project Name + Vendor + Document Date feed the document name until the user
  // types their own. Each handler passes its own new value plus the current
  // state of the other two, since its setState has not applied yet.
  const onProjectNameChange = (v: string): void => {
    setProjectName(v);
    if (!docNameEdited) setDocName(composeDocName(v, vendor, documentDate));
  };
  const onVendorChange = (v: string): void => {
    setVendor(v);
    if (!docNameEdited) setDocName(composeDocName(projectName, v, documentDate));
  };
  const onDocumentDateChange = (iso: string): void => {
    setDocumentDate(iso);
    if (!docNameEdited) setDocName(composeDocName(projectName, vendor, iso));
  };
  const onDocNameChange = (v: string): void => {
    setDocName(v);
    // Blank name re-enables auto-composition; any real text is treated as a manual override.
    setDocNameEdited(v.trim() !== "");
  };

  const resetForm = (): void => {
    setFile(undefined);
    setDocName("");
    setDocNameEdited(false);
    setDocumentType("");
    setLevelValues([]);
    setYearPeriod("");
    setDocumentDate("");
    setConfidentiality("");
    setVendor("");
    setProjectName("");
    setRemark("");
    setLegallyPrivileged(false);
    if (fileRef.current) fileRef.current.value = "";
  };

  /* ---------- Upload ------------------------------------------------------ */

  const handleUpload = async (): Promise<void> => {
    const missing: string[] = [];
    if (!file) missing.push("File");
    if (!documentType) missing.push("Document Type");
    const m = activeMode();
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) missing.push(lvl.label);
    });
    // These strings are shown to the user, so they must match the on-screen
    // field labels — renamed to "Year" / "Confidential Level" in the relayout.
    if (!yearPeriod) missing.push("Year");
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidential Level");
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }
    if (!file) return;

    const finalName = buildUploadName(file.name, docName);
    const allowed = settings.allowedFileTypes;
    if (allowed.kind === "none") {
      showToast(NO_TYPES_MESSAGE, "error");
      return;
    }
    if (!allowed.types.some((ext) => finalName.toLowerCase().endsWith(ext))) {
      showToast(
        `File type not allowed. Allowed: ${allowed.types.join(", ")}`,
        "error",
      );
      return;
    }

    const mode = activeMode();
    if (!mode || mode.levels.length === 0) {
      showToast("No upload mode configured.", "error");
      return;
    }
    // The leaf level's selected term is the permissioned Unit folder.
    const leafIdx = mode.levels.length - 1;
    const leafTerm = (levelChoices[leafIdx] ?? []).find(
      (o) => o.id === levelValues[leafIdx],
    );
    if (!leafTerm) {
      showToast("Please choose all folder levels before uploading.", "error");
      return;
    }

    setBusy(true);
    setStatus("Locating destination folder…");

    // Resolve the Unit folder by its stable UniqueId (rename-proof), NOT by a
    // name-built path. The selected leaf term is the lookup key.
    const mapping = await lookupFolderMapping(
      context.spHttpClient,
      siteUrl,
      leafTerm.id,
    ).catch((e: unknown) => {
      console.error("Folder map lookup error:", e);
      return null;
    });
    if (!mapping || !mapping.folderUniqueId) {
      showToast(
        `This folder hasn't been mapped yet. Ask an administrator to run the reconciliation tool. (term ${leafTerm.label})`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }

    // Rename-proof: resolve the Unit folder's CURRENT path from its UniqueId, then
    // ensure-create the Year and Document Type subfolders under it (they inherit its ACL).
    const unitSru = await resolveFolderServerUrl(
      context.spHttpClient,
      siteUrl,
      mapping.folderUniqueId,
    );
    if (!unitSru) {
      showToast(
        "The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation.",
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }
    const yearLabel = sanitizeFolderSegment(
      options.yearPeriod.find((o) => o.id === yearPeriod)?.label ?? "",
    );
    const docTypeLabel = sanitizeFolderSegment(
      options.documentType.find((o) => o.id === documentType)?.label ?? "",
    );
    if (!yearLabel || !docTypeLabel) {
      showToast("Year and Document Type are required.", "error");
      setStatus("");
      setBusy(false);
      return;
    }

    setStatus("Preparing destination folders…");
    const yearFolder = await ensureFolder(
      context.spHttpClient,
      siteUrl,
      unitSru,
      yearLabel,
    );
    if (!yearFolder) {
      showToast(`Could not create the "${yearLabel}" folder.`, "error");
      setStatus("");
      setBusy(false);
      return;
    }
    const destFolder = await ensureFolder(
      context.spHttpClient,
      siteUrl,
      yearFolder.serverRelativeUrl,
      docTypeLabel,
    );
    if (!destFolder) {
      showToast(`Could not create the "${docTypeLabel}" folder.`, "error");
      setStatus("");
      setBusy(false);
      return;
    }

    const folderId = destFolder.uniqueId; // upload target — a fresh, unit-scoped folder
    let uploadedServerRelativeUrl = "";

    setStatus("Checking for duplicates…");

    // Flipped to true only when the user confirms overwriting an existing file.
    let overwrite = false;
    try {
      // Duplicate check — GetFolderById targets the folder by UniqueId, so a
      // rename of that folder does not affect this lookup.
      const existsRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (existsRes.ok) {
        // A file with this name already exists — ask the user whether to replace
        // it rather than silently blocking the upload.
        setStatus("");
        const confirmed = await askReplace(finalName);
        if (!confirmed) {
          setBusy(false);
          return;
        }
        overwrite = true;
      }
    } catch {
      // Network error on existence check — proceed; upload will surface the real error.
    }

    setStatus("Uploading…");

    try {
      // Upload directly into the folder resolved by UniqueId. No name-path
      // existence probe is needed — GetFolderById either resolves (folder
      // still exists under its current name/location) or 404s (deleted).
      const uploadRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=${overwrite})?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1,
        { body: file },
      );
      if (!uploadRes.ok) {
        let detail = `HTTP ${uploadRes.status}`;
        try {
          const bodyText = await uploadRes.text();
          try {
            const errJson = JSON.parse(bodyText);
            const spMsg =
              errJson?.error?.message?.value ?? errJson?.error?.message;
            detail += spMsg
              ? ` — ${spMsg}`
              : bodyText
                ? ` — ${bodyText.slice(0, 300)}`
                : "";
          } catch {
            if (bodyText) detail += ` — ${bodyText.slice(0, 300)}`;
          }
        } catch {
          /* body already consumed or unreadable */
        }
        console.error("Upload failed:", folderId, detail);
        // A 404 here means the mapped folder no longer exists (deleted after mapping).
        const hint =
          uploadRes.status === 404
            ? " The mapped folder may have been deleted — ask an administrator to re-run the reconciliation tool."
            : "";
        showToast(`Upload failed (${detail}).${hint}`, "error");
        return;
      }
      const uploadJson = await uploadRes.json();
      uploadedServerRelativeUrl = uploadJson.ServerRelativeUrl;

      const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeServerRelativePath(uploadedServerRelativeUrl)}'`,
        SPHttpClient.configurations.v1,
      );
      if (!itemRes.ok) {
        showToast("Uploaded, but could not retrieve the item to tag.", "error");
        return;
      }
      const item = await itemRes.json();

      // One label + one term-GUID pair per level, plus the Business Segment
      // column (the mode's segment label / term-set GUID). Each level carries
      // its real Staging internal names (labelCol/tidCol) from DMS Config when
      // present; buildLevelFormValues falls back to LEVEL_COLUMNS otherwise.
      const selections = (mode.levels ?? []).map((lvl, i) => {
        const opt = (levelChoices[i] ?? []).find(
          (o) => o.id === levelValues[i],
        );
        return {
          column: lvl.column,
          label: opt?.label ?? "",
          id: opt?.id ?? "",
          labelCol: lvl.labelCol,
          tidCol: lvl.tidCol,
        };
      });
      selections.unshift({
        column: "BusinessSegment",
        label: mode.label,
        id: mode.termSetGuid,
        labelCol: undefined,
        tidCol: undefined,
      });

      // Level columns: BusinessSegment (injected, not a config level) takes its column
      // names from settings.columns so a new site needs no code change; Department/Unit
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
          FieldValue: toTaxValue(options.documentType, documentType),
        },
        {
          FieldName: settings.columns.yearPeriod,
          FieldValue: toTaxValue(options.yearPeriod, yearPeriod),
        },
        {
          FieldName: settings.columns.confidentiality,
          FieldValue: toTaxValue(options.confidentiality, confidentiality),
        },
        ...buildLevelFormValues(levelCols, selections),
      ];

      // Document Date is required, so validation above already guarantees a
      // value. The guard is a safety net: toSpDate("") yields the malformed
      // "NaN/NaN/", which SharePoint rejects with a HasException that surfaces
      // as a confusing "Uploaded, but a field failed" long after the cause.
      if (documentDate) {
        formValues.push({
          FieldName: settings.columns.documentDate,
          FieldValue: toSpDate(documentDate),
        });
      }

      // Project Name and Vendor are optional free text, but they are written
      // UNCONDITIONALLY — a blank field sends "" and clears the column.
      //
      // This matters on the replace path: Files/Add(overwrite=true) swaps the
      // file's content but reuses the same list item, so anything not written
      // here survives from the previous upload. Skipping blanks would leave an
      // approver looking at a vendor or project belonging to the document that
      // was just replaced. Sending "" is safe for a Text column (unlike the
      // date above, where an empty value would reach toSpDate).
      formValues.push({
        FieldName: settings.columns.projectName,
        FieldValue: projectName.trim(),
      });
      formValues.push({
        FieldName: settings.columns.vendor,
        FieldValue: vendor.trim(),
      });
      formValues.push({
        FieldName: settings.columns.remark,
        FieldValue: remark.trim(),
      });
      // Written on EVERY upload, and forced false unless the chosen confidentiality
      // level is the one configured to offer it. The tick is hidden when the level
      // changes, but hiding a control does not clear the state behind it, and on the
      // replace path the list item is reused — so a document could inherit a legal
      // flag from the file it overwrote. Deriving the value here rather than trusting
      // the checkbox makes that impossible.
      const privilegedApplies =
        settings.legallyPrivilegedFor !== "" &&
        confidentiality === settings.legallyPrivilegedFor;
      formValues.push({
        FieldName: settings.columns.legallyPrivileged,
        FieldValue: privilegedApplies && legallyPrivileged ? "true" : "false",
      });

      const metaRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${settings.stagingLibrary}')/items(${item.Id})/validateUpdateListItem`,
        SPHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ formValues }),
        },
      );
      if (!metaRes.ok) {
        showToast("Uploaded, but tagging metadata failed.", "error");
        return;
      }
      const metaJson = await metaRes.json();
      const fieldError = (metaJson.value ?? []).find(
        (v: { HasException?: boolean }) => v.HasException,
      );
      if (fieldError) {
        console.error("Field update error:", fieldError);
        showToast(
          `Uploaded, but a field failed: ${fieldError.FieldName} — ${fieldError.ErrorMessage}`,
          "error",
        );
        return;
      }

      showToast(
        "Document uploaded successfully and is pending review.",
        "success",
      );
      setStatus("");
      resetForm();
    } catch (err) {
      console.error("Upload failed:", err);
      showToast("Upload failed. Please try again.", "error");
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

  /* ---------- Render ------------------------------------------------------ */

  return (
    <section className="dms-form">
      <style>{`
        .dms-form { max-width: 960px; margin: 32px auto; padding: 0 24px 48px; font-family: 'Segoe UI', sans-serif; }
        .dms-form h2 { margin: 0 0 4px; font-size: 24px; font-weight: 700; color: #1b1b1b; }
        .dms-subtitle { margin: 0 0 28px; font-size: 14px; color: #666; }
        .dms-section { background: #fff; border: 1px solid #e1e1e1; border-radius: 8px; padding: 24px; margin-bottom: 20px; }
        .dms-section-title { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #0f6c3f; margin: 0 0 16px; }
        .dms-filecard { display: flex; align-items: center; justify-content: flex-start; gap: 16px; border: 1px dashed #8a8a8a; border-radius: 10px; padding: 16px; }
        .dms-filecard .name { font-weight: 600; }
        .dms-filecard .size { color: #666; font-size: 12px; }
        /* Textarea inherits the input styling so Remark matches the fields around it —
           without this it renders in the browser's default monospace at a random width. */
        .dms-field textarea { font: inherit; width: 100%; box-sizing: border-box; padding: 8px 10px;
          border: 1px solid #c8c8c8; border-radius: 4px; resize: vertical; }
        .dms-field textarea:focus { outline: 2px solid #0f6c3f; outline-offset: -1px; }
        /* Checkbox row: label beside the box, hint underneath and aligned with it. */
        .dms-check { display: grid; grid-template-columns: auto 1fr; gap: 2px 8px; align-items: center; }
        .dms-check input { margin: 0; }
        .dms-check small { grid-column: 2; color: #666; font-size: 12px; }
        .dms-link { background: none; border: none; color: #0f6c3f; cursor: pointer; font-weight: 600; padding: 0; font-size: 13px; }
        .dms-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 16px; font-size: 13px; }
        .dms-field > span { font-weight: 600; }
        .dms-field .req { color: #d13438; font-style: normal; }
        .dms-field select, .dms-field input[type="text"], .dms-field input[type="date"] { padding: 8px 10px; border: 1px solid #c8c8c8; border-radius: 10px; font: inherit; width: 100%; box-sizing: border-box; height: 38px; background: #fff; }
        .dms-field small { color: #666; font-size: 12px; font-weight: 400; }
        /* Confidentiality info tooltip. The icon is taken out of flow and placed
           in the grid gutter, so the select keeps the FULL column width and lines
           up with Vendor above it. bottom:10px centres the 18px icon on the 38px
           select; left keeps it inside the 24px gutter (6 + 18 = 24). */
        .dms-conf { position: relative; margin-bottom: 16px; }
        .dms-conf .dms-field { margin-bottom: 0; }
        .dms-info { position: absolute; left: calc(100% + 6px); bottom: 10px; width: 18px; height: 18px; border-radius: 50%; background: #0f6c3f; color: #fff; font-size: 12px; font-weight: 700; font-style: normal; display: inline-flex; align-items: center; justify-content: center; cursor: help; }
        /* Opens to the right of the icon, into the empty third grid column.
           280px keeps it inside the card rather than spilling past its edge. */
        .dms-info-panel { display: none; position: absolute; top: -8px; left: calc(100% + 8px); z-index: 30; width: 280px; max-width: calc(100vw - 48px); padding: 16px; background: #fff; border: 1px solid #e1e1e1; border-radius: 10px; box-shadow: 0 4px 16px rgba(0,0,0,.12); cursor: default; text-align: left; }
        .dms-info:hover .dms-info-panel, .dms-info:focus .dms-info-panel, .dms-info:focus-within .dms-info-panel { display: block; }
        .dms-info-panel dl { margin: 0; }
        .dms-info-panel dt { margin-top: 12px; color: #0f6c3f; font-size: 13px; font-weight: 700; }
        .dms-info-panel dt:first-of-type { margin-top: 0; }
        .dms-info-panel dd { margin: 4px 0 0; color: #444; font-size: 12px; font-weight: 400; line-height: 1.45; }
        .dms-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0 24px; }
        /* Folder card runs 2-up so the deepest level and Year pair evenly. */
        .dms-grid-2 { grid-template-columns: 1fr 1fr; }
        .dms-radio-group { display: flex; gap: 24px; margin-bottom: 20px; }
        .dms-radio-group label { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 600; cursor: pointer; color: #1b1b1b; }
        .dms-radio-group input[type="radio"] { accent-color: #0f6c3f; width: 16px; height: 16px; cursor: pointer; }
        .dms-dept-badge { display: inline-flex; align-items: center; gap: 8px; background: #e8f5ee; border: 1px solid #b3d9c4; border-radius: 20px; padding: 5px 14px; font-size: 13px; margin-bottom: 20px; }
        .dms-dept-badge .dept-label { font-weight: 400; color: #555; }
        .dms-dept-badge .dept-name { font-weight: 700; color: #0f6c3f; }
        .dms-dept-loading { font-size: 13px; color: #666; margin-bottom: 20px; }
        .dms-dept-error { font-size: 13px; color: #d13438; background: #fdf3f3; border: 1px solid #f1c0c0; border-radius: 4px; padding: 10px 14px; margin-bottom: 20px; }
        .dms-admin-row { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; }
        .dms-admin-badge { display: inline-flex; align-items: center; background: #fff4e5; border: 1px solid #f0c070; border-radius: 20px; padding: 5px 14px; font-size: 13px; font-weight: 700; color: #7a4f00; white-space: nowrap; }
        .dms-admin-row select { padding: 6px 10px; border: 1px solid #c8c8c8; border-radius: 4px; font: inherit; font-size: 13px; }
        .dms-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 4px; }
        .dms-btn { padding: 9px 24px; border-radius: 4px; cursor: pointer; font: inherit; font-size: 14px; border: 1px solid transparent; }
        .dms-btn.primary { background: #0f6c3f; color: #fff; }
        .dms-btn.primary:disabled { background: #9bbfaa; cursor: default; }
        .dms-btn.secondary { background: #fff; border-color: #0f6c3f; color: #0f6c3f; }
        .dms-status { margin-top: 16px; font-size: 13px; }
        .dms-toast { position: fixed; top: 24px; right: 24px; z-index: 9999; min-width: 300px; max-width: 460px; padding: 14px 40px 14px 16px; border-radius: 6px; font-size: 13px; font-family: 'Segoe UI', sans-serif; box-shadow: 0 4px 16px rgba(0,0,0,.18); animation: dms-slidein .2s ease; }
        .dms-toast.error { background: #d13438; color: #fff; }
        .dms-toast-close { position: absolute; top: 10px; right: 12px; background: none; border: none; cursor: pointer; font-size: 16px; color: inherit; opacity: .7; line-height: 1; }
        .dms-toast-close:hover { opacity: 1; }
        @keyframes dms-slidein { from { transform: translateX(60px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        .dms-popup-overlay { position: fixed; inset: 0; background: rgba(0,0,0,.25); backdrop-filter: blur(6px); -webkit-backdrop-filter: blur(6px); z-index: 9998; display: flex; align-items: center; justify-content: center; padding: 16px; animation: dms-fadein .15s ease; }
        .dms-popup { background: #fff; border-radius: 16px; padding: 40px 40px 32px; text-align: center; max-width: 420px; width: 100%; box-shadow: 0 8px 40px rgba(0,0,0,.15); animation: dms-popin .2s ease; }
        .dms-popup-icon { margin-bottom: 20px; }
        .dms-popup-svg { width: 120px; height: 120px; display: block; margin: 0 auto; }
        .dms-popup-title { font-size: 22px; font-weight: 700; color: #0f6c3f; margin: 0 0 12px; }
        .dms-popup-msg { font-size: 14px; color: #555; margin: 0 0 28px; line-height: 1.6; }
        .dms-popup-ok-container { display:flex; align-items:center; justify-content:center; }
        .dms-popup-ok { background: #0f6c3f; max-width: 183px; color: #fff; border: none; border-radius: 6px; padding: 12px 0; width: 100%; font-size: 14px; font-weight: 600; font-family: inherit; cursor: pointer; }
        .dms-popup-ok:hover { background: #0a5230; }
        .dms-popup-actions { display:flex; align-items:center; justify-content:center; gap: 12px; }
        .dms-popup-btn { min-width: 96px; border-radius: 6px; padding: 11px 22px; font-size: 14px; font-weight: 600; font-family: inherit; cursor: pointer; }
        .dms-popup-btn.confirm { background: #0f6c3f; color: #fff; border: none; }
        .dms-popup-btn.confirm:hover { background: #0a5230; }
        .dms-popup-btn.cancel { background: #fff; color: #0f6c3f; border: 1px solid #0f6c3f; }
        .dms-popup-btn.cancel:hover { background: #f0f6f2; }
        @media (max-width: 480px) {
          .dms-popup { padding: 28px 20px 24px; border-radius: 12px; }
          .dms-popup-svg { width: 88px; height: 88px; }
          .dms-popup-title { font-size: 18px; }
          .dms-popup-msg { margin-bottom: 20px; font-size: 13px; }
          .dms-toast { left: 12px; right: 12px; min-width: unset; top: 12px; }
        }
        @keyframes dms-fadein { from { opacity: 0; } to { opacity: 1; } }
        @keyframes dms-popin { from { transform: scale(.92); opacity: 0; } to { transform: scale(1); opacity: 1; } }
      `}</style>

      <h2>Upload Document Form</h2>
      <p className="dms-subtitle">
        All fields marked <strong>*</strong> are required.
      </p>

      {/* ── Document Details ────────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Document Details</p>
        <div style={{ marginBottom: 16 }}>
          <label className="dms-field">
            <span>Document Name</span>
            {/* Typeable before a file is picked: the name is only read at upload
                time, and leaving it enabled avoids a greyed-out first field. */}
            <input
              type="text"
              value={docName}
              maxLength={50}
              placeholder={file ? `Leave blank to keep "${file.name}"` : ""}
              onChange={(e) => onDocNameChange(e.target.value)}
            />
            <small>Max. 50 characters</small>
          </label>
        </div>
        {/* An empty AllowedFileTypes selection is a hard block, not a silent
            fallback — spec 2026-07-30 §3. The message names the column and the
            list because the client is the one who fixes it, in one click. */}
        {settings.allowedFileTypes.kind === "none" ? (
          <div className="dms-filecard" style={{ opacity: 0.6 }}>
            <span>{NO_TYPES_MESSAGE}</span>
          </div>
        ) : (
          /* The whole card is clickable to open the file picker. */
          <div
            className="dms-filecard"
            role="button"
            tabIndex={0}
            style={{ cursor: "pointer" }}
            onClick={() => fileRef.current?.click()}
            onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") fileRef.current?.click(); }}
          >
            {file ? (
              <>
                <span className="dms-filecard-ready">READY</span>
                <span className="name">{file.name}</span>
                <span className="size">{formatFileSize(file.size)}</span>
                <span className="dms-link dms-filecard-action">Change Document</span>
              </>
            ) : (
              <span className="dms-link">Choose a document</span>
            )}
            <input
              ref={fileRef}
              type="file"
              accept={settings.allowedFileTypes.types.join(",")}
              style={{ display: "none" }}
              onChange={(e) => {
                const picked = e.target.files?.[0];
                // Re-checked here because the narrowing above does not reach
                // inside the callback.
                const types =
                  settings.allowedFileTypes.kind === "none"
                    ? []
                    : settings.allowedFileTypes.types;
                if (
                  picked &&
                  !types.some((ext) => picked.name.toLowerCase().endsWith(ext))
                ) {
                  showToast(
                    `File type not allowed. Allowed: ${types.join(", ")}`,
                    "error",
                  );
                  setFile(undefined);
                  if (fileRef.current) fileRef.current.value = "";
                  return;
                }
                setStatus("");
                setFile(picked);
              }}
            />
          </div>
        )}

        <div className="dms-grid" style={{ marginTop: 16 }}>
          {/* Free-text Project Name — distinct from the Group-led Projects
              "Group Project Name" folder level in the card below. */}
          <label className="dms-field">
            <span>Project Name</span>
            <input
              type="text"
              value={projectName}
              maxLength={50}
              placeholder="Type the project name"
              onChange={(e) => onProjectNameChange(e.target.value)}
            />
            <small>Max. 50 characters</small>
          </label>

          {/* Vendor is free text. It also feeds the auto-composed document name. */}
          <label className="dms-field">
            <span>Vendor/Customer Name</span>
            <input
              type="text"
              value={vendor}
              maxLength={50}
              placeholder="Type the vendor or customer name"
              onChange={(e) => onVendorChange(e.target.value)}
            />
            <small>Max. 50 characters</small>
          </label>

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
              onChange={(e) => onDocumentDateChange(e.target.value)}
            />
          </label>

          {/* Document Type used to sit here. It moved into the folder card below:
              it is part of the destination path (Unit → Year → Document Type), not a
              property of the document, and grouping it with Unit and Year is what the
              client's mockup shows. */}

          {/* Confidential Level carries an info tooltip defining each term.
              tabIndex makes it keyboard-reachable; :focus-within keeps the
              panel open while it holds focus. */}
          <div className="dms-conf">
            {renderSelect(
              "Confidential Level",
              true,
              confidentiality,
              setConfidentiality,
              options.confidentiality,
            )}
            <em
              className="dms-info"
              tabIndex={0}
              role="button"
              aria-label="What the confidentiality levels mean"
            >
              i
              <span className="dms-info-panel" role="tooltip">
                {/* Highly Confidential is deliberately absent. Its term is removed
                    from the term store for Phase 1, so the dropdown cannot offer it,
                    and describing a level nobody can pick reads as a bug in UAT. The
                    definition returns with the HC libraries in Phase 2 — see the
                    highly-confidential-securing design on feat/hc-libraries. */}
                <dl>
                  <dt>Legally Privileged</dt>
                  <dd>
                    This applies to confidential communications (email, advice,
                    documents, conversations) between client and lawyer that are
                    protected by law from being disclosed in a court of law or
                    during legal proceedings.
                  </dd>
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
          </div>

          {/* Offered only for the level named by `legallyPrivilegedFor` in DMS Config.
              Unset means never offered — see the setting's note. The value is re-derived
              at upload time rather than trusted from here, because hiding the control
              does not clear the state behind it. */}
          {settings.legallyPrivilegedFor !== "" &&
            confidentiality === settings.legallyPrivilegedFor && (
              <label className="dms-check" style={{ gridColumn: "1 / -1" }}>
                <input
                  type="checkbox"
                  checked={legallyPrivileged}
                  onChange={(e) => setLegallyPrivileged(e.target.checked)}
                />
                <span>Legally Privileged</span>
                <small>
                  Tick if this is a protected communication between client and lawyer.
                </small>
              </label>
            )}

          {/* Full row: a remark is prose, and half a row wraps it to four lines. */}
          <label className="dms-field" style={{ gridColumn: "1 / -1" }}>
            <span>Remark</span>
            <textarea
              value={remark}
              maxLength={250}
              rows={3}
              placeholder="Anything the approver should know about this document"
              onChange={(e) => setRemark(e.target.value)}
            />
            <small>{remark.length}/250 characters</small>
          </label>
        </div>
      </div>

      {/* ── Document Folder Information ──────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Document Folder Information</p>

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
                  name="sideToggle"
                  checked={active}
                  onChange={() => switchMode(offerable[0].key)}
                />
                {side === "BusinessSegment" ? "Business Segment" : "Project"}
              </label>
            );
          })}
        </div>

        {/* Segment picker — only when the active side offers more than one mode. */}
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

        <div className="dms-grid dms-grid-2">
          {/* --- File-path fields, in folder order: Segment (above) -> level(s)
                 -> Year. Every level spans the full row EXCEPT the deepest one,
                 which shares its row with Year. Document Type also forms part of
                 the path but is ensure-created on demand, so it lives above. --- */}
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
              deptLoading ||
                isLevelLocked(i) ||
                (i > 0 && !levelValues[i - 1]),
              `Select ${lvl.label}`,
              i < arr.length - 1,
            ),
          )}

          {renderSelect("Year", true, yearPeriod, setYearPeriod, options.yearPeriod)}

          {/* Document Type completes the path: the deepest level is the permissioned
              Unit folder, and Year / Document Type are ensure-created beneath it on
              first use. It sits with Unit and Year because all three decide WHERE the
              file lands, unlike the fields above, which describe the file itself. */}
          {renderSelect(
            "Document Type",
            true,
            documentType,
            setDocumentType,
            options.documentType,
          )}

          {/* Graceful empty-state: a segment whose term set has no child terms yet
              (e.g. the non-GHO Head Offices before their Department/Unit trees are
              added) would otherwise show a dropdown with nothing but its
              "Select …" placeholder. Name the missing level instead. */}
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
          onClick={handleUpload}
          disabled={busy || deptLoading}
        >
          {busy ? "Uploading…" : "Upload"}
        </button>
      </div>

      {status && <p className="dms-status">{status}</p>}

      {toast && toast.type === "error" && (
        <div className="dms-toast error" role="alert">
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

      {replaceAsk && (
        <div className="dms-popup-overlay" role="dialog" aria-modal="true">
          <div className="dms-popup">
            <div className="dms-popup-icon">
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
            </div>
            <p className="dms-popup-title">Replace Existing File</p>
            <p className="dms-popup-msg">
              We noticed there&rsquo;s a same name file.
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

      {toast && toast.type === "success" && (
        <div className="dms-popup-overlay" role="dialog" aria-modal="true">
          <div className="dms-popup">
            <div className="dms-popup-icon">
              <svg
                className="dms-popup-svg"
                viewBox="0 0 184 184"
                fill="none"
                xmlns="http://www.w3.org/2000/svg"
              >
                <circle
                  opacity="0.3"
                  cx="92.0001"
                  cy="92"
                  r="75.4872"
                  fill="#14C7A5"
                />
                <circle
                  cx="92"
                  cy="92"
                  r="92"
                  fill="#14C7A5"
                  fillOpacity="0.2"
                />
                <circle
                  cx="92.0003"
                  cy="91.9998"
                  r="61.3333"
                  fill="white"
                  stroke="#14C7A5"
                  strokeWidth="3"
                />
                <path
                  d="M67 90.7143L88.4286 110L117 74"
                  stroke="#14C7A5"
                  strokeWidth="10"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </div>
            <p className="dms-popup-title">Upload Successful</p>
            <p className="dms-popup-msg">
              Your file is awaiting approval.
              <br />
              Please revisit the Home page to track progress.
            </p>
            <div className="dms-popup-ok-container">
              <button
                className="dms-popup-ok"
                onClick={() => {
                  window.location.href = siteUrl;
                }}
              >
                Track File Status
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
