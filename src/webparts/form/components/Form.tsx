import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IFormProps } from "./IFormProps";
import {
  lookupFolderMapping,
  resolveFolderServerUrl,
  ensureFolder,
} from "../../../shared/dmsFolderMap";
import {
  parseLevels,
  matchUserPaths,
  sanitizeFolderSegment,
  buildLevelFormValues,
  Level,
  GroupMapRow,
  UserPath,
  ColumnPair,
} from "../../../shared/formModel";

/* ----------------------------------------------------------------------------
 * CONFIG — hardcoded values are fallbacks only; live values load from DMS Config SP list
 * -------------------------------------------------------------------------- */

const FIELDS = {
  // Internal name frozen as "Department_x0020_Type" (created as "Department Type",
  // then display-renamed to "Document Type" — verified against live Staging fields API).
  documentType: "Department_x0020_Type",
  yearPeriod: "Year_x002f_Period",
  documentDate: "DocumentDate",
  confidentiality: "Confidentiality_x0020_Level",
  vendor: "Vendor",
  details: "_ExtendedDescription",
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
//   - Group Head Office has a real term-set GUID (verified in the sandbox).
//   - The other four carry REPLACE-* placeholder GUIDs: swap them for the real
//     term-set GUIDs (or, preferably, drive everything from DMS Config so no
//     code edit is needed). Each level's real Staging column internal names are
//     supplied per-level via DMS Config Levels JSON (labelCol/tidCol); the
//     logical `column` key here resolves via LEVEL_COLUMNS as a fallback only.
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
    key: "upstream",
    label: "Group Upstream Operations",
    side: "BusinessSegment",
    termSetGuid: "REPLACE-UPSTREAM-TERMSET-GUID",
    stagingFolder: "Group Upstream Operations",
    levels: [
      { label: "Region", column: "Region" },
      { label: "Estate/Mill", column: "EstateMill" },
    ],
    sortOrder: 2,
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
    sortOrder: 3,
  },
  {
    key: "it",
    label: "Group Innovation & Technology",
    side: "BusinessSegment",
    termSetGuid: "REPLACE-IT-TERMSET-GUID",
    stagingFolder: "Group Innovation & Technology",
    levels: [{ label: "I&T Operating Unit", column: "ITOperatingUnit" }],
    sortOrder: 4,
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
    sortOrder: 5,
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

export default function Form({ context }: IFormProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const fileRef = useRef<HTMLInputElement>(null);

  const [options, setOptions] = useState<OptionMap>(EMPTY_OPTIONS);
  const [deptLoading, setDeptLoading] = useState<boolean>(true);
  const [isAdmin, setIsAdmin] = useState<boolean>(false);

  // Generic N-level cascade state: one option list + one selected term id per level.
  const [levelChoices, setLevelChoices] = useState<TermOption[][]>([]);
  const [levelValues, setLevelValues] = useState<string[]>([]);
  const [userPaths, setUserPaths] = useState<UserPath[]>([]);

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [uploadMode, setUploadMode] = useState<string>("");
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);
  const [file, setFile] = useState<File | undefined>(undefined);
  const [docName, setDocName] = useState<string>("");
  const [documentType, setDocumentType] = useState<string>("");
  const [yearPeriod, setYearPeriod] = useState<string>("");
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [vendor, setVendor] = useState<string>("");
  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
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
        unitTermGuid: r.UnitTermGuid ?? "",
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

  // Detect the user's M365 (Entra) group Object IDs via MS Graph /me/memberOf.
  // These are matched against DMS Group Map GroupId (Object ID) — names are cosmetic.
  const loadUserGroupIds = async (): Promise<string[]> => {
    try {
      const graph = await context.msGraphClientFactory.getClient("3");
      const memberOf = await graph.api("/me/memberOf").select("id").get();
      return (memberOf.value ?? [])
        .map((g: { id?: string }) => g.id ?? "")
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

  // Pre-fill all levels from a detected UserPath (walk the term ancestry).
  const prefillFromPath = async (
    mode: UploadMode,
    path: UserPath,
  ): Promise<void> => {
    const chain = await loadTermPath(mode.termSetGuid, path.unitTermGuid); // [top..leaf]
    const choices: TermOption[][] = [];
    const values: string[] = [];
    choices[0] = await loadTermSet(mode.termSetGuid).catch(
      () => [] as TermOption[],
    );
    for (let i = 0; i < chain.length && i < mode.levels.length; i++) {
      values[i] = chain[i].id;
      if (i + 1 < mode.levels.length) {
        choices[i + 1] = await loadTermChildren(
          mode.termSetGuid,
          chain[i].id,
        ).catch(() => [] as TermOption[]);
      }
    }
    setLevelChoices(choices);
    setLevelValues(values);
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
      // Ignore config rows that predate the Side/Levels schema (empty Levels) —
      // fall back to the built-in modes so the form never renders a broken cascade.
      const usable = rawModes.filter(
        (m: UploadMode) => m.levels.length > 0 && !!m.termSetGuid,
      );
      const loadedModes = usable.length > 0 ? usable : DEFAULT_MODES;
      setModes(loadedModes);
      setIsAdmin(admin);

      const paths = matchUserPaths(groupMap, userGroupIds);
      setUserPaths(paths);

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

      // Default toggle = BusinessSegment; default mode = first matched segment
      // path, else the first BusinessSegment mode.
      const bsModes = loadedModes.filter(
        (m: UploadMode) => m.side === "BusinessSegment",
      );
      const firstPath = paths.find((p) =>
        bsModes.some((m: UploadMode) => m.termSetGuid === p.segment),
      );
      const defaultMode =
        (firstPath &&
          bsModes.find(
            (m: UploadMode) => m.termSetGuid === firstPath.segment,
          )) ??
        bsModes[0] ??
        loadedModes[0];
      if (defaultMode) {
        setUploadMode(defaultMode.key);
        if (firstPath) await prefillFromPath(defaultMode, firstPath);
        else await initCascade(defaultMode);
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
    const path = userPaths.find((p) => p.segment === mode.termSetGuid);
    if (path) prefillFromPath(mode, path).catch(() => initCascade(mode));
    else
      initCascade(mode).catch(() => {
        setLevelChoices([]);
        setLevelValues([]);
      });
  };

  const toTaxValue = (opts: TermOption[], id: string): string => {
    const match = opts.find((o) => o.id === id);
    return match ? `${match.label}|${match.id}` : "";
  };

  const resetForm = (): void => {
    setFile(undefined);
    setDocName("");
    setDocumentType("");
    setLevelValues([]);
    setYearPeriod("");
    setDocumentDate("");
    setConfidentiality("");
    setVendor("");
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
    if (!yearPeriod) missing.push("Year / Period");
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidentiality Level");
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }
    if (!file) return;

    const finalName = buildUploadName(file.name, docName);
    if (
      !settings.allowedExtensions.some((ext) =>
        finalName.toLowerCase().endsWith(ext),
      )
    ) {
      showToast(
        `File type not allowed. Allowed: ${settings.allowedExtensions.join(", ")}`,
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

    try {
      // Duplicate check — GetFolderById targets the folder by UniqueId, so a
      // rename of that folder does not affect this lookup.
      const existsRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (existsRes.ok) {
        showToast(
          `A file named "${finalName}" already exists in this location. Rename your document or choose a different file.`,
          "error",
        );
        setStatus("");
        setBusy(false);
        return;
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
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=false)?$select=ServerRelativeUrl`,
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
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeURIComponent(uploadedServerRelativeUrl)}'`,
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
        .dms-filecard { display: flex; align-items: center; justify-content: space-between; gap: 12px; border: 1px dashed #c8c8c8; border-radius: 6px; padding: 16px; }
        .dms-filecard .name { font-weight: 600; }
        .dms-filecard .size { color: #666; font-size: 12px; }
        .dms-link { background: none; border: none; color: #0f6c3f; cursor: pointer; font-weight: 600; padding: 0; font-size: 13px; }
        .dms-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 16px; font-size: 13px; }
        .dms-field > span { font-weight: 600; }
        .dms-field .req { color: #d13438; font-style: normal; }
        .dms-field select, .dms-field input[type="text"], .dms-field input[type="date"] { padding: 8px 10px; border: 1px solid #c8c8c8; border-radius: 4px; font: inherit; width: 100%; box-sizing: border-box; height: 38px; }
        .dms-field small { color: #666; font-size: 12px; font-weight: 400; }
        .dms-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0 24px; }
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

      <h2>Upload a document</h2>
      <p className="dms-subtitle">
        All fields marked <strong>*</strong> are required.
      </p>

      {/* ── File ────────────────────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">File</p>
        <div className="dms-filecard">
          {file ? (
            <div>
              <div className="name">{file.name}</div>
              <div className="size">{(file.size / 1024).toFixed(1)} KB</div>
            </div>
          ) : (
            <span className="size">No file selected</span>
          )}
          <button
            type="button"
            className="dms-link"
            onClick={() => fileRef.current?.click()}
          >
            {file ? "Change file" : "Select file"}
          </button>
          <input
            ref={fileRef}
            type="file"
            accept={settings.allowedExtensions.join(",")}
            style={{ display: "none" }}
            onChange={(e) => {
              const picked = e.target.files?.[0];
              if (
                picked &&
                !settings.allowedExtensions.some((ext) =>
                  picked.name.toLowerCase().endsWith(ext),
                )
              ) {
                showToast(
                  `File type not allowed. Allowed: ${settings.allowedExtensions.join(", ")}`,
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
        <div style={{ marginTop: 16 }}>
          <label className="dms-field">
            <span>Document name</span>
            <input
              type="text"
              value={docName}
              disabled={!file}
              placeholder={
                file
                  ? `Leave blank to keep "${file.name}"`
                  : "Select a file first"
              }
              onChange={(e) => setDocName(e.target.value)}
            />
            {file && (
              <small>Saved as: {buildUploadName(file.name, docName)}</small>
            )}
          </label>
        </div>
      </div>

      {/* ── Document Information ─────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Document Information</p>

        {deptLoading ? (
          <p className="dms-dept-loading">Loading your access&hellip;</p>
        ) : userPaths.length === 0 && !isAdmin ? (
          <div className="dms-dept-error">
            Your account isn&apos;t mapped to any unit. Contact your
            administrator before uploading.
          </div>
        ) : null}

        {/* Business Segment | Project toggle — sides driven by DMS Config */}
        <div className="dms-radio-group">
          <p>Upload into:</p>
          {(["BusinessSegment", "Project"] as const).map((side) => {
            const first = modes.find((m) => m.side === side);
            if (!first) return null;
            const active = activeMode()?.side === side;
            return (
              <label key={side}>
                <input
                  type="radio"
                  name="sideToggle"
                  checked={active}
                  onChange={() => switchMode(first.key)}
                />
                {side === "BusinessSegment" ? "Business Segment" : "Project"}
              </label>
            );
          })}
        </div>

        {/* Segment picker — shown only when the active side has more than one mode */}
        {(() => {
          const side = activeMode()?.side;
          const sideModes = modes.filter((m) => m.side === side);
          if (sideModes.length <= 1) return null;
          return (
            <label className="dms-field">
              <span>Segment</span>
              <select
                value={uploadMode}
                onChange={(e) => switchMode(e.target.value)}
              >
                {sideModes.map((m) => (
                  <option key={m.key} value={m.key}>
                    {m.label}
                  </option>
                ))}
              </select>
            </label>
          );
        })()}

        <div className="dms-grid">
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
            "Vendor (if applicable)",
            false,
            vendor,
            setVendor,
            options.vendor,
          )}

          {(activeMode()?.levels ?? []).map((lvl, i) =>
            renderSelect(
              lvl.label,
              true,
              levelValues[i] ?? "",
              (v) => {
                const md = activeMode();
                if (md) onLevelChange(md, i, v).catch(() => undefined);
              },
              levelChoices[i] ?? [],
              deptLoading || (i > 0 && !levelValues[i - 1]),
              "--",
              true,
            ),
          )}

          {renderSelect(
            "Confidentiality Level",
            true,
            confidentiality,
            setConfidentiality,
            options.confidentiality,
          )}

          {renderSelect(
            "Year / Period",
            true,
            yearPeriod,
            setYearPeriod,
            options.yearPeriod,
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
                Track Status on Home
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
