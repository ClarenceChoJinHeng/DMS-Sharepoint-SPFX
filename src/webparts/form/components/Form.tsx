import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IFormProps } from "./IFormProps";
import { lookupFolderMapping } from "../../../shared/dmsFolderMap";

/* ----------------------------------------------------------------------------
 * CONFIG — hardcoded values are fallbacks only; live values load from DMS Settings SP list
 * -------------------------------------------------------------------------- */

const FIELDS = {
  // Internal name frozen as "Department_x0020_Type" (created as "Department Type",
  // then display-renamed to "Document Type" — verified against live Staging fields API).
  documentType: "Department_x0020_Type",
  department: "Department",
  yearPeriod: "Year_x002f_Period",
  documentDate: "DocumentDate",
  deptSubTeam: "Department_x0020_Team",
  projectSubTeam: "Project_x0020_Name",
  confidentiality: "Confidentiality_x0020_Level",
  vendor: "Vendor",
  details: "_ExtendedDescription",
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


type TermOption     = { id: string; label: string };
type ToastType = "error" | "success";

// A single entry in the flattened subcategory dropdown. `path` is the full
// ancestor chain (excluding the department itself) down to and including this
// term — used to rebuild the real nested folder path on upload.
type SubTeamNode = { term: TermOption; path: TermOption[] };

type UploadMode = {
  key: string;
  label: string;
  termSetGuid: string;
  stagingFolder: string;
  lookupStyle: "direct" | "parentMatch";
  subTeamLabel: string;
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

// Fallback if DMS Config list is missing or unreachable — preserves existing behaviour.
const DEFAULT_MODES: UploadMode[] = [
  {
    key: "department",
    label: "Department",
    termSetGuid: "eaba82e5-3e5f-4719-9a76-091f034ad407",
    stagingFolder: "Departments",
    lookupStyle: "direct",
    subTeamLabel: "Department Team",
    sortOrder: 1,
  },
  {
    key: "project",
    label: "Project",
    termSetGuid: "94ce322b-4515-4fda-8f50-35709f1f521d",
    stagingFolder: "Projects",
    lookupStyle: "parentMatch",
    subTeamLabel: "Sub Project",
    sortOrder: 2,
  },
];

type DmsSettings = {
  termSets: {
    documentType: string;
    department: string;
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
    department: "eaba82e5-3e5f-4719-9a76-091f034ad407",
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
  const [detectedDept, setDetectedDept] = useState<TermOption | null>(null);
  const [deptLoading, setDeptLoading] = useState<boolean>(true);
  const [isAdmin, setIsAdmin] = useState<boolean>(false);
  const [allDepts, setAllDepts] = useState<TermOption[]>([]);
  const [adminSelectedDept, setAdminSelectedDept] = useState<string>("");
  // Nested subcategory chain (e.g. Account D > Sub-1 > Sub-2 > ... > Sub-6),
  // flattened into one list for a single dropdown. Each node keeps its full
  // ancestor path so the real nested folder path can be rebuilt on upload.
  const [subTeamChoices, setSubTeamChoices] = useState<SubTeamNode[]>([]);
  const [subTeamValue, setSubTeamValue] = useState<string>("");

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [modeParents, setModeParents] = useState<Record<string, TermOption[]>>(
    {},
  );
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

  const loadModes = async (): Promise<UploadMode[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('DMS%20Config')/items?$select=Title,ModeLabel,TermSetGuid,StagingFolder,LookupStyle,SubTeamLabel,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error("DMS Config list not found");
    const data = await res.json();
    return (data.value ?? []).map(
      (item: {
        Title: string;
        ModeLabel: string;
        TermSetGuid: string;
        StagingFolder: string;
        LookupStyle: string;
        SubTeamLabel: string;
        SortOrder: number;
      }) => ({
        key: item.Title,
        label: item.ModeLabel,
        termSetGuid: item.TermSetGuid,
        stagingFolder: item.StagingFolder,
        lookupStyle: (item.LookupStyle || "direct") as "direct" | "parentMatch",
        subTeamLabel: item.SubTeamLabel || "Sub-team",
        sortOrder: item.SortOrder,
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
        department:
          get("termSet_department") ?? DEFAULT_SETTINGS.termSets.department,
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

  // Recursively walks every descendant of `rootId` (the whole nested chain,
  // however deep it goes) and flattens it into one list. Each node keeps its
  // full ancestor path so the real nested folder path can be rebuilt later.
  const loadDescendants = async (
    mode: UploadMode,
    rootId: string,
  ): Promise<SubTeamNode[]> => {
    const result: SubTeamNode[] = [];
    const walk = async (
      parentId: string,
      ancestorPath: TermOption[],
    ): Promise<void> => {
      const children = await loadTermChildren(mode.termSetGuid, parentId).catch(
        () => [] as TermOption[],
      );
      for (const child of children) {
        const path = [...ancestorPath, child];
        result.push({ term: child, path });
        await walk(child.id, path);
      }
    };
    await walk(rootId, []);
    return result;
  };

  // "direct" modes (e.g. Department) scope the chain to the uploader's own
  // department root. "parentMatch" modes (e.g. Project) have their own term
  // set whose top-level terms mirror department labels — find the matching
  // root by label, then flatten its chain.
  const loadSubTeamTree = async (
    mode: UploadMode,
    dept: TermOption | null,
    parents: TermOption[],
  ): Promise<SubTeamNode[]> => {
    if (!dept) return [];
    if (mode.lookupStyle === "direct") {
      return loadDescendants(mode, dept.id);
    }
    const parent = parents.find((p) =>
      dept.label.toLowerCase().includes(p.label.toLowerCase()),
    );
    if (!parent) return [];
    return loadDescendants(mode, parent.id);
  };

  /* ---------- Group detection --------------------------------------------- */

  // Uses MS Graph /me/memberOf to detect M365 (Domain) group membership.
  // SP REST /currentuser/groups only returns SharePoint-native groups and misses
  // M365 groups added from the admin center (e.g. SDG-IT-Uploaders).
  // Falls back to SP REST groups if Graph is unavailable.
  const detectDepartment = async (
    depts: TermOption[],
  ): Promise<{ dept: TermOption | null; admin: boolean }> => {
    const [graphClient, userRes] = await Promise.all([
      context.msGraphClientFactory.getClient("3"),
      context.spHttpClient.get(
        `${siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      ),
    ]);

    let groupNames: string[] = [];
    try {
      const memberOf = await graphClient
        .api("/me/memberOf")
        .select("displayName")
        .get();
      groupNames = (memberOf.value ?? []).map((g: { displayName?: string }) =>
        (g.displayName ?? "").toLowerCase(),
      );
    } catch {
      // Graph unavailable — fall back to SP REST groups (SharePoint-native only)
      try {
        const spRes = await context.spHttpClient.get(
          `${siteUrl}/_api/web/currentuser/groups?$select=Title`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json" } },
        );
        if (spRes.ok) {
          const data = await spRes.json();
          groupNames = (data.value ?? []).map((g: { Title: string }) =>
            g.Title.toLowerCase(),
          );
        }
      } catch {
        /* nothing — groupNames stays empty */
      }
    }

    // TEMP: simulate SDG-IT-Uploader membership for visual testing — REVERT before shipping
    groupNames.push("sdg-it-uploader");

    const uploaderGroups = groupNames.filter((g) => g.includes("uploader"));
    // Strip all non-alphanumeric chars and lowercase — used on both sides so
    // hyphens, spaces, and casing differences never cause a mismatch.
    const bare = (s: string): string =>
      s.toLowerCase().replace(/[^a-z0-9]/g, "");
    const dept =
      depts.find((d) => {
        const termBare = bare(d.label);
        return uploaderGroups.some((g) => {
          // Extract just the dept portion from the group name by stripping the
          // "SDG" prefix and "Uploader" suffix — whatever remains is the dept key.
          const core = bare(g.replace(/^sdg/i, "").replace(/uploaders?$/i, ""));
          // Term bare must start with the extracted core so "accountd" matches
          // core "account", and exact matches ("account" === "account") also work.
          return core.length > 0 && termBare.startsWith(core);
        });
      }) ?? null;

    let isSiteAdmin = false;
    if (userRes.ok) {
      const userData = await userRes.json();
      isSiteAdmin = userData.IsSiteAdmin === true;
    }
    const admin = isSiteAdmin || groupNames.some((g) => g.includes("owner"));
    return { dept, admin };
  };

  /* ---------- Init -------------------------------------------------------- */

  useEffect(() => {
    const init = async (): Promise<void> => {
      // Load config lists in parallel first — term sets depend on settings GUIDs.
      const [loadedSettings, loadedModes] = await Promise.all([
        loadSettings().catch(() => DEFAULT_SETTINGS),
        loadModes().catch(() => DEFAULT_MODES),
      ]);
      setSettings(loadedSettings);
      setModes(loadedModes);
      setUploadMode(loadedModes[0]?.key ?? "");

      const [docTypes, depts, years, confs, vendors] = await Promise.all([
        loadTermSet(loadedSettings.termSets.documentType).catch(
          () => [] as TermOption[],
        ),
        loadTermSet(loadedSettings.termSets.department).catch(
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

      // For parentMatch modes, load top-level terms now so switchMode can use them instantly.
      const parentEntries = await Promise.all(
        loadedModes
          .filter((m) => m.lookupStyle === "parentMatch")
          .map(async (m) => {
            const parents = await loadTermSet(m.termSetGuid).catch(
              () => [] as TermOption[],
            );
            return [m.key, parents] as [string, TermOption[]];
          }),
      );
      const parentMap: Record<string, TermOption[]> = {};
      parentEntries.forEach(([key, parents]) => {
        parentMap[key] = parents;
      });
      setModeParents(parentMap);

      const { dept, admin } = await detectDepartment(depts);
      setIsAdmin(false); // TEMP: force non-admin view — REVERT before shipping
      setAllDepts(depts);
      setDetectedDept(dept);
      setDeptLoading(false);

      if (loadedModes.length > 0) {
        const defaultMode = loadedModes[0];
        const tree = await loadSubTeamTree(
          defaultMode,
          dept,
          parentMap[defaultMode.key] ?? [],
        );
        setSubTeamChoices(tree);
        setSubTeamValue("");
      }
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
    setSubTeamValue("");
    setSubTeamChoices([]);
    setStatus("");

    if (isAdmin) {
      setAdminSelectedDept("");
      setDetectedDept(null);
      return;
    }

    if (!detectedDept) return;

    const mode = modes.find((m) => m.key === modeKey);
    if (!mode) return;

    loadSubTeamTree(mode, detectedDept, modeParents[modeKey] ?? [])
      .then(setSubTeamChoices)
      .catch(() => setSubTeamChoices([]));
  };

  const toTaxValue = (opts: TermOption[], id: string): string => {
    const match = opts.find((o) => o.id === id);
    return match ? `${match.label}|${match.id}` : "";
  };

  const resetForm = (): void => {
    setFile(undefined);
    setDocName("");
    setDocumentType("");
    setSubTeamValue("");
    setYearPeriod("");
    setDocumentDate("");
    setConfidentiality("");
    setVendor("");
    if (fileRef.current) fileRef.current.value = "";
  };

  const handleAdminDeptChange = (deptId: string): void => {
    setAdminSelectedDept(deptId);
    setSubTeamValue("");
    setSubTeamChoices([]);

    const mode = modes.find((m) => m.key === uploadMode);
    if (!mode) return;

    if (mode.lookupStyle === "parentMatch") {
      // deptId here is a Project-term-set id (the dropdown lists modeParents,
      // not allDepts, for this lookup style — see the render below).
      const parents = modeParents[mode.key] ?? [];
      const picked = parents.find((d) => d.id === deptId) ?? null;
      const deptForWrite =
        allDepts.find((d) =>
          d.label.toLowerCase().includes(picked?.label.toLowerCase() ?? ""),
        ) ?? null;
      if (picked && !deptForWrite) {
        // picked.id belongs to the Project term set, not the Department term
        // set the "Department" column is bound to — writing it as-is would
        // tag the item with a GUID that's real but in the wrong term set.
        showToast(
          `"${picked.label}" has no matching entry in the Department term set. Ask your administrator to align the Department and Project term sets before uploading here.`,
          "error",
        );
      }
      setDetectedDept(deptForWrite);
      if (picked) {
        loadDescendants(mode, picked.id)
          .then(setSubTeamChoices)
          .catch(() => setSubTeamChoices([]));
      }
      return;
    }

    const picked = allDepts.find((d) => d.id === deptId) ?? null;
    setDetectedDept(picked);
    if (picked) {
      loadDescendants(mode, picked.id)
        .then(setSubTeamChoices)
        .catch(() => setSubTeamChoices([]));
    }
  };

  /* ---------- Upload ------------------------------------------------------ */

  const handleUpload = async (): Promise<void> => {
    const missing: string[] = [];
    if (!file) missing.push("File");
    if (!detectedDept) missing.push("Department");
    if (!documentType) missing.push("Document Type");
    if (!subTeamValue)
      missing.push(
        modes.find((m) => m.key === uploadMode)?.subTeamLabel ?? "Sub-team",
      );
    if (!yearPeriod) missing.push("Year / Period");
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidentiality Level");
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }
    if (!file || !detectedDept) return;

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

    // Upload to Staging/Department/ or Staging/Projects/ depending on toggle mode.
    // These folders must be created once in the Staging library before first upload.
    const selectedChoice = subTeamChoices.find((c) => c.term.id === subTeamValue);
    if (!selectedChoice) {
      showToast(
        "Could not resolve the selected sub-team/project folder.",
        "error",
      );
      return;
    }
    // The picked term itself is the metadata value; its `path` is the full
    // nested chain from the department down to it (e.g. Sub-1/Sub-2/Sub-3).
    const selectedSubTeam = selectedChoice.term;

    // Resolve the destination folder by its stable UniqueId (rename-proof),
    // NOT by a name-built path. The selected leaf term is the lookup key.
    setBusy(true);
    setStatus("Locating destination folder…");

    const mapping = await lookupFolderMapping(
      context.spHttpClient,
      siteUrl,
      selectedSubTeam.id,
    ).catch((e: unknown) => {
      console.error("Folder map lookup error:", e);
      return null;
    });
    if (!mapping || !mapping.folderUniqueId) {
      showToast(
        `This folder hasn't been mapped yet. Ask an administrator to register it in the reconciliation tool. (term ${selectedSubTeam.label})`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }
    const folderId = mapping.folderUniqueId;
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
        // Department is auto-detected from SP group — never user-entered.
        {
          FieldName: FIELDS.department,
          FieldValue: `${detectedDept.label}|${detectedDept.id}`,
        },
      ];

      if (selectedSubTeam) {
        formValues.push({
          FieldName: uploadMode === "department" ? FIELDS.deptSubTeam : FIELDS.projectSubTeam,
          FieldValue: `${selectedSubTeam.label}|${selectedSubTeam.id}`,
        });
      }

      if (vendor) {
        formValues.push({
          FieldName: FIELDS.vendor,
          FieldValue: toTaxValue(options.vendor, vendor),
        });
      }

      // TEMP DIAGNOSTIC — remove after debugging blank-metadata issue.
      // Logs the item we're tagging and the EXACT payload we send. If any
      // FieldValue below is "" the tag will silently blank that column.
      console.log("[DMS DEBUG] item.Id =", item.Id);
      console.table(formValues);

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
      // TEMP DIAGNOSTIC — SharePoint echoes each field's stored value or its
      // exception here. This is the ground truth for whether tagging worked.
      console.log("[DMS DEBUG] validateUpdateListItem response:", JSON.stringify(metaJson));
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
          <p className="dms-dept-loading">Detecting your department&hellip;</p>
        ) : isAdmin ? (
          // Admin — always show dept picker regardless of own group membership
          <div className="dms-admin-row">
            <span className="dms-admin-badge">Admin</span>
            <select
              value={adminSelectedDept}
              onChange={(e) => handleAdminDeptChange(e.target.value)}
            >
              <option value="">
                {modes.find((m) => m.key === uploadMode)?.lookupStyle ===
                "parentMatch"
                  ? "Select project…"
                  : "Select department…"}
              </option>
              {(modes.find((m) => m.key === uploadMode)?.lookupStyle ===
              "parentMatch"
                ? (modeParents[uploadMode] ?? [])
                : allDepts
              ).map((d) => (
                <option key={d.id} value={d.id}>
                  {d.label}
                </option>
              ))}
            </select>
          </div>
        ) : detectedDept ? (
          <div className="dms-dept-badge">
            <span className="dept-label">Your department:</span>
            <span className="dept-name">{detectedDept.label}</span>
          </div>
        ) : (
          <div className="dms-dept-error">
            Your account is not assigned to a department group. Contact your
            administrator before uploading.
          </div>
        )}

        {/* Upload mode toggle — options driven by DMS Config list */}
        <div className="dms-radio-group">
        <p>Folder Location:</p>
          {modes.map((m) => (
            <label key={m.key}>
              <input
                type="radio"
                name="uploadMode"
                value={m.key}
                checked={uploadMode === m.key}
                onChange={() => switchMode(m.key)}
              />
              {m.label}
            </label>
          ))}
        </div>

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
                const m = d.getMonth() + 1;
                const day = d.getDate();
                return `${d.getFullYear()}-${m < 10 ? "0" + m : m}-${day < 10 ? "0" + day : day}`;
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

          {renderSelect(
            modes.find((m) => m.key === uploadMode)?.subTeamLabel ??
              "Sub-team",
            true,
            subTeamValue,
            setSubTeamValue,
            subTeamChoices.map((c) => c.term),
            deptLoading || !detectedDept,
            "--",
            true,
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
          disabled={busy || deptLoading || !detectedDept}
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
