import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IBulkUploadProps } from "./IBulkUploadProps";
import {
  lookupFolderMapping,
  resolveFolderServerUrl,
  resolveFolderByPath,
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
 * BULK UPLOAD — a duplicate of the `form` web part with two behaviour changes:
 *   1. Accepts up to MAX_FILES files per submission instead of one.
 *   2. Writes straight into the Documents library, bypassing Staging (and so
 *      bypassing content approval and the Auto-route flow).
 * TEMPORARY TOOL. See docs/superpowers/specs/2026-07-23-bulk-upload-direct-to-documents-design.md
 *
 * Everything else — metadata tagging, group-based mode detection, the term
 * cascade, per-Unit folder routing — is cloned from Form.tsx unchanged.
 * -------------------------------------------------------------------------- */

const MAX_FILES = 50;

// The Documents library's real URL segment differs from its display title:
//   URL segment   = "Shared Documents"  (used to build server-relative paths)
//   display title = "Documents"         (used by lists/getbytitle for metadata)
// This matches the Auto-route flow's own concat('/Shared Documents/', …) target.
const DOCUMENTS_URL_SEGMENT = "Shared Documents";
const DOCUMENTS_LIST_TITLE = "Documents";

const FIELDS = {
  // Internal name frozen as "Department_x0020_Type" (created as "Department Type",
  // then display-renamed to "Document Type" — verified against the live fields API).
  documentType: "Department_x0020_Type",
  yearPeriod: "Year_x002f_Period",
  documentDate: "DocumentDate",
  confidentiality: "Confidentiality_x0020_Level",
  vendor: "Vendor",
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

// One user-selected file plus its optional per-file rename.
type PickedFile = { key: string; file: File; rename: string };

type Outcome = "uploaded" | "skipped" | "failed" | "tagFailed";
type FileResult = { name: string; outcome: Outcome; detail?: string };

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
    termSetGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
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
    termSetGuid: "5ab1c7c4-78d2-43b4-869f-3eab4b1c375c",
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
    termSetGuid: "6ba9a64c-a363-48fd-afd1-324897df781c",
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
    termSetGuid: "21d7e6fe-8f71-4a56-bd2e-e4a2176995a7",
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
      { label: "Project Name", column: "ProjectName" },
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
  stagingLibrary: string;
  allowedExtensions: string[];
};

const DEFAULT_SETTINGS: DmsSettings = {
  termSets: {
    documentType: "0540e66e-7cb3-47ac-b0ef-4e3069387394",
    yearPeriod: "f7c578a1-e0e5-42ff-9e0c-d748cba42ede",
    confidentiality: "032534ab-9285-4b42-98c6-5c7b0df1f066",
    vendor: "cb3c0ab7-a959-4200-9b7b-d1e13397d240",
  },
  stagingLibrary: "Staging",
  allowedExtensions: [".pdf", ".xls", ".xlsx"],
};

export default function BulkUpload({
  context,
}: IBulkUploadProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const webSru = context.pageContext.web.serverRelativeUrl.replace(/\/+$/, "");
  const fileRef = useRef<HTMLInputElement>(null);

  const [options, setOptions] = useState<OptionMap>(EMPTY_OPTIONS);
  const [deptLoading, setDeptLoading] = useState<boolean>(true);
  // Privileged = site admin OR a GLOBAL-role uploader: bypasses tier detection
  // and gets the full manual cascade (may upload anywhere).
  const [privileged, setPrivileged] = useState<boolean>(false);

  // Generic N-level cascade state: one option list + one selected term id per level.
  const [levelChoices, setLevelChoices] = useState<TermOption[][]>([]);
  const [levelValues, setLevelValues] = useState<string[]>([]);
  const [validPaths, setValidPaths] = useState<ValidPath[]>([]);

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [uploadMode, setUploadMode] = useState<string>("");
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);
  const [picked, setPicked] = useState<PickedFile[]>([]);
  const [documentType, setDocumentType] = useState<string>("");
  const [yearPeriod, setYearPeriod] = useState<string>("");
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [vendor, setVendor] = useState<string>("");
  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [results, setResults] = useState<FileResult[] | null>(null);
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
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,ModeLabel,Side,TermSetGuid,StagingFolder,Levels,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error("DMS Config list not found");
    const data = await res.json();
    return (data.value ?? []).map(
      (item: {
        Title: string;
        ModeLabel: string;
        Side: string;
        TermSetGuid: string;
        StagingFolder: string;
        Levels: string;
        SortOrder: number;
      }) => ({
        key: item.Title,
        label: item.ModeLabel,
        side: (item.Side === "Project" ? "Project" : "BusinessSegment") as
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
      const isPrivileged = admin || membership.isGlobalUploader;
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

      if (isPrivileged) {
        const bsModes = loadedModes.filter(
          (m: UploadMode) => m.side === "BusinessSegment",
        );
        const defaultMode = bsModes[0] ?? loadedModes[0];
        if (defaultMode) {
          setUploadMode(defaultMode.key);
          await initCascade(defaultMode);
        }
      } else {
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
      console.error("BulkUpload init failed:", err);
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

  const resetForm = (): void => {
    setPicked([]);
    setDocumentType("");
    setLevelValues([]);
    setYearPeriod("");
    setDocumentDate("");
    setConfidentiality("");
    setVendor("");
    setResults(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  /* ---------- File selection ---------------------------------------------- */

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
        showToast(
          `You can upload at most ${MAX_FILES} files at a time.`,
          "error",
        );
        return prev;
      }
      if (allowed.length > room) {
        showToast(
          `Only the first ${room} file(s) were added — the limit is ${MAX_FILES}.`,
          "error",
        );
      }
      const stamp = Date.now();
      const added = allowed.slice(0, room).map((file, i) => ({
        key: `${stamp}-${i}-${file.name}`,
        file,
        rename: "",
      }));
      return [...prev, ...added];
    });

    setStatus("");
    setResults(null);
    // Clear the input so the same file can be re-picked after a removal.
    if (fileRef.current) fileRef.current.value = "";
  };

  const setRename = (key: string, value: string): void => {
    setPicked((prev) =>
      prev.map((p) => (p.key === key ? { ...p, rename: value } : p)),
    );
  };

  const removeFile = (key: string): void => {
    setPicked((prev) => prev.filter((p) => p.key !== key));
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

  /* ---------- Upload ------------------------------------------------------ */

  const handleUpload = async (): Promise<void> => {
    const missing: string[] = [];
    if (picked.length === 0) missing.push("File");
    if (!documentType) missing.push("Document Type");
    const m = activeMode();
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) missing.push(lvl.label);
    });
    if (!yearPeriod) missing.push("Year / Period");
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidentiality Level");
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }

    // Final names must be unique within the batch, or files silently collide:
    // the first wins and the rest are reported as duplicates.
    const finalNames = picked.map((p) => buildUploadName(p.file.name, p.rename));
    const collision = finalNames.find(
      (n, i) =>
        finalNames.findIndex((o) => o.toLowerCase() === n.toLowerCase()) !== i,
    );
    if (collision) {
      showToast(
        `Two or more files would be saved as "${collision}". Rename them first.`,
        "error",
      );
      return;
    }

    const mode = activeMode();
    if (!mode || mode.levels.length === 0) {
      showToast("No upload mode configured.", "error");
      return;
    }
    const leafIdx = mode.levels.length - 1;
    const leafTerm = (levelChoices[leafIdx] ?? []).find(
      (o) => o.id === levelValues[leafIdx],
    );
    if (!leafTerm) {
      showToast("Please choose all folder levels before uploading.", "error");
      return;
    }

    setBusy(true);
    setResults(null);
    setStatus("Locating destination folder…");

    try {
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
        return;
      }

      // Resolve the Staging unit folder's CURRENT path (rename-proof), then swap
      // the library segment to reach the mirrored Documents unit folder.
      const stagingSru = await resolveFolderServerUrl(
        context.spHttpClient,
        siteUrl,
        mapping.folderUniqueId,
      );
      if (!stagingSru) {
        showToast(
          "The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation.",
          "error",
        );
        setStatus("");
        return;
      }

      const docsUnitPath = toDocumentsPath(stagingSru);
      if (!docsUnitPath) {
        console.error(
          "Could not swap library segment. stagingSru:",
          stagingSru,
          "expected prefix:",
          `${webSru}/${settings.stagingLibrary}/`,
        );
        showToast(
          "Could not work out the Documents path for this unit folder. Check the stagingLibrary setting in DMS Config.",
          "error",
        );
        setStatus("");
        return;
      }

      // The unit folder must ALREADY exist in Documents. It is deliberately not
      // auto-created: a folder created here would inherit the Documents root
      // ACL and silently widen access. A missing one is an administrator task.
      const docsUnitFolder = await resolveFolderByPath(
        context.spHttpClient,
        siteUrl,
        docsUnitPath,
      );
      if (!docsUnitFolder) {
        showToast(
          `The matching folder does not exist in Documents yet (${docsUnitPath}). Ask an administrator to create it — it is not created automatically, so its permissions stay correct.`,
          "error",
        );
        setStatus("");
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
        return;
      }

      setStatus("Preparing destination folders…");
      // Year / Document Type subfolders are safe to create — they inherit the
      // unit folder's ACL, exactly as they do in Staging.
      const yearFolder = await ensureFolder(
        context.spHttpClient,
        siteUrl,
        docsUnitFolder.serverRelativeUrl,
        yearLabel,
      );
      if (!yearFolder) {
        showToast(`Could not create the "${yearLabel}" folder.`, "error");
        setStatus("");
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
        return;
      }
      const folderId = destFolder.uniqueId;

      // Batch metadata — identical for every file in this submission.
      const selections = (mode.levels ?? []).map((lvl, i) => {
        const opt = (levelChoices[i] ?? []).find((o) => o.id === levelValues[i]);
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

      const formValues: Array<{ FieldName: string; FieldValue: string }> = [
        {
          FieldName: FIELDS.documentType,
          FieldValue: toTaxValue(options.documentType, documentType),
        },
        {
          FieldName: FIELDS.yearPeriod,
          FieldValue: toTaxValue(options.yearPeriod, yearPeriod),
        },
        {
          FieldName: FIELDS.confidentiality,
          FieldValue: toTaxValue(options.confidentiality, confidentiality),
        },
        { FieldName: FIELDS.documentDate, FieldValue: toSpDate(documentDate) },
        ...buildLevelFormValues(LEVEL_COLUMNS, selections),
      ];
      if (vendor) {
        formValues.push({
          FieldName: FIELDS.vendor,
          FieldValue: toTaxValue(options.vendor, vendor),
        });
      }

      /* ----- Sequential per-file upload; one failure never stops the batch -- */
      const batch = [...picked];
      const collected: FileResult[] = [];
      const succeededKeys = new Set<string>();

      for (let i = 0; i < batch.length; i++) {
        const entry = batch[i];
        const finalName = buildUploadName(entry.file.name, entry.rename);
        setStatus(`Uploading ${i + 1} of ${batch.length} — ${finalName}`);

        // Duplicate probe. overwrite=false is kept, so an existing file is
        // never clobbered; it is reported as skipped instead.
        try {
          const existsRes: SPHttpClientResponse =
            await context.spHttpClient.get(
              `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
              SPHttpClient.configurations.v1,
              { headers: { Accept: "application/json;odata=nometadata" } },
            );
          if (existsRes.ok) {
            collected.push({
              name: finalName,
              outcome: "skipped",
              detail: "A file with this name already exists here.",
            });
            continue;
          }
        } catch {
          // Network error on the existence check — proceed; the upload will
          // surface the real error.
        }

        try {
          const uploadRes: SPHttpClientResponse =
            await context.spHttpClient.post(
              `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=false)?$select=ServerRelativeUrl`,
              SPHttpClient.configurations.v1,
              { body: entry.file },
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
                    ? ` — ${bodyText.slice(0, 200)}`
                    : "";
              } catch {
                if (bodyText) detail += ` — ${bodyText.slice(0, 200)}`;
              }
            } catch {
              /* body already consumed or unreadable */
            }
            console.error("Upload failed:", finalName, folderId, detail);
            collected.push({ name: finalName, outcome: "failed", detail });
            continue;
          }
          const uploadJson = await uploadRes.json();
          const uploadedSru = uploadJson.ServerRelativeUrl;

          const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
            `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeURIComponent(uploadedSru)}'`,
            SPHttpClient.configurations.v1,
          );
          if (!itemRes.ok) {
            collected.push({
              name: finalName,
              outcome: "tagFailed",
              detail: "Uploaded, but the item could not be retrieved to tag.",
            });
            succeededKeys.add(entry.key);
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
            collected.push({
              name: finalName,
              outcome: "tagFailed",
              detail: `Uploaded, but tagging failed (HTTP ${metaRes.status}).`,
            });
            succeededKeys.add(entry.key);
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
            collected.push({
              name: finalName,
              outcome: "tagFailed",
              detail: `Uploaded, but a field failed: ${fieldError.FieldName} — ${fieldError.ErrorMessage}`,
            });
            succeededKeys.add(entry.key);
            continue;
          }

          collected.push({ name: finalName, outcome: "uploaded" });
          succeededKeys.add(entry.key);
        } catch (err) {
          console.error("Upload threw:", finalName, err);
          collected.push({
            name: finalName,
            outcome: "failed",
            detail: err instanceof Error ? err.message : "Unexpected error.",
          });
        }
      }

      setResults(collected);
      setStatus("");
      // Keep only the rows that did not fully succeed, so the user can fix and
      // retry without re-picking everything.
      setPicked((prev) => prev.filter((p) => !succeededKeys.has(p.key)));

      const okCount = collected.filter((r) => r.outcome === "uploaded").length;
      if (okCount === collected.length && okCount > 0) {
        showToast(`All ${okCount} file(s) uploaded to Documents.`, "success");
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

  const summary = results
    ? {
        uploaded: results.filter((r) => r.outcome === "uploaded").length,
        skipped: results.filter((r) => r.outcome === "skipped").length,
        failed: results.filter((r) => r.outcome === "failed").length,
        tagFailed: results.filter((r) => r.outcome === "tagFailed").length,
      }
    : null;

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
        .dms-filelist { margin-top: 16px; display: flex; flex-direction: column; gap: 10px; }
        .dms-filerow { display: grid; grid-template-columns: minmax(0,1fr) minmax(0,1fr) auto; gap: 12px; align-items: start; border: 1px solid #ececec; border-radius: 6px; padding: 10px 12px; background: #fafafa; }
        .dms-filerow .orig { font-size: 13px; font-weight: 600; word-break: break-all; }
        .dms-filerow .size { font-size: 12px; color: #666; }
        .dms-filerow input[type="text"] { padding: 6px 8px; border: 1px solid #c8c8c8; border-radius: 4px; font: inherit; font-size: 13px; width: 100%; box-sizing: border-box; }
        .dms-filerow .preview { font-size: 11px; color: #666; margin-top: 4px; word-break: break-all; }
        .dms-remove { background: none; border: none; cursor: pointer; color: #d13438; font-size: 15px; line-height: 1; padding: 4px; }
        .dms-count { font-size: 12px; color: #666; margin-top: 10px; }
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
        .dms-status { margin-top: 16px; font-size: 13px; }
        .dms-results { margin-top: 20px; }
        .dms-results-summary { font-size: 13px; font-weight: 600; margin: 0 0 10px; }
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
          .dms-filerow { grid-template-columns: 1fr auto; }
          .dms-toast { left: 12px; right: 12px; min-width: unset; top: 12px; }
        }
      `}</style>

      <h2>Bulk upload</h2>
      <p className="dms-subtitle">
        All fields marked <strong>*</strong> are required. Up to {MAX_FILES}{" "}
        files per upload.
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

      {/* ── Files ───────────────────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Files</p>
        <div className="dms-drop">
          <span className="size">
            {picked.length === 0
              ? "No files selected"
              : `${picked.length} file(s) selected`}
          </span>
          <button
            type="button"
            className="dms-link"
            onClick={() => fileRef.current?.click()}
            disabled={busy}
          >
            {picked.length === 0 ? "Select files" : "Add more files"}
          </button>
          <input
            ref={fileRef}
            type="file"
            multiple
            accept={settings.allowedExtensions.join(",")}
            style={{ display: "none" }}
            onChange={(e) => addFiles(e.target.files)}
          />
        </div>

        {picked.length > 0 && (
          <>
            <div className="dms-filelist">
              {picked.map((p) => (
                <div className="dms-filerow" key={p.key}>
                  <div>
                    <div className="orig">{p.file.name}</div>
                    <div className="size">
                      {(p.file.size / 1024).toFixed(1)} KB
                    </div>
                  </div>
                  <div>
                    <input
                      type="text"
                      value={p.rename}
                      disabled={busy}
                      placeholder="Leave blank to keep this name"
                      aria-label={`Rename ${p.file.name}`}
                      onChange={(e) => setRename(p.key, e.target.value)}
                    />
                    {p.rename.trim() !== "" && (
                      <div className="preview">
                        Saved as: {buildUploadName(p.file.name, p.rename)}
                      </div>
                    )}
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
            <p className="dms-count">
              {picked.length} of {MAX_FILES} files.
            </p>
          </>
        )}
      </div>

      {/* ── Document Information ─────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">
          Document Information — applies to every file above
        </p>

        {deptLoading ? (
          <p className="dms-dept-loading">Loading your access&hellip;</p>
        ) : !privileged && validPaths.length === 0 ? (
          <div className="dms-dept-error">
            Your account isn&apos;t fully provisioned to upload — you need
            membership at every level plus the unit uploader role. Contact your
            administrator.
          </div>
        ) : null}

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
                {side === "BusinessSegment" ? "Business Segment" : "Project"}
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
                    (levelChoices[i] ?? []).find((o) => o.id === levelValues[i])
                      ?.label,
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
              deptLoading || isLevelLocked(i) || (i > 0 && !levelValues[i - 1]),
              "--",
              true,
            ),
          )}

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

          {renderSelect(
            "Vendor (if applicable)",
            false,
            vendor,
            setVendor,
            options.vendor,
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
          Clear
        </button>
        <button
          type="button"
          className="dms-btn primary"
          onClick={handleUpload}
          disabled={busy || deptLoading || picked.length === 0}
        >
          {busy
            ? "Uploading…"
            : `Upload ${picked.length > 0 ? picked.length : ""} file(s)`}
        </button>
      </div>

      {status && <p className="dms-status">{status}</p>}

      {/* ── Per-file results ────────────────────────────────────────────── */}
      {results && summary && (
        <div className="dms-results">
          <p className="dms-results-summary">
            {summary.uploaded} uploaded
            {summary.skipped > 0 && `, ${summary.skipped} skipped`}
            {summary.failed > 0 && `, ${summary.failed} failed`}
            {summary.tagFailed > 0 &&
              `, ${summary.tagFailed} uploaded without tags`}
            .
            {(summary.failed > 0 || summary.skipped > 0) &&
              " Failed and skipped files are still listed above so you can rename and retry."}
            {summary.tagFailed > 0 &&
              " Files listed as “uploaded, not tagged” are already in Documents — re-uploading would duplicate them; fix their metadata in the library instead."}
          </p>
          {results.map((r, i) => (
            <div className={`dms-result ${r.outcome}`} key={`${r.name}-${i}`}>
              <span className="tag">{outcomeLabel[r.outcome]}</span>
              <span className="fname">{r.name}</span>
              {r.detail && <span className="why">{r.detail}</span>}
            </div>
          ))}
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
