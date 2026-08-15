import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IFormProps } from "./IFormProps";
import {
  lookupFolderMapping,
  loadFolderMapRows,
  resolveMappedFolder,
  ensureFolder,
  encodeServerRelativePath,
  probeFolderUploadAccess,
  FolderMapRow,
} from "../../../shared/dmsFolderMap";
import {
  AccessVerdict,
  filterProvisionedPaths,
  filterReachablePaths,
  mappedTermGuidSet,
  normalizeTermGuid,
  segmentProvisionState,
} from "../../../shared/segmentReadiness";
import { formatFileSize } from "../../../shared/fileSize";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX, libraryTitle } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import {
  Batch,
  BatchDestination,
  FileMeta,
  StagedFile,
  UploadResult,
  applyUploadResults,
  batchesNeedingRepick,
  canSaveBatch,
  collisionsWithin,
  duplicateAcrossBatches,
  inheritDefaults,
  nextId,
  stagedTotals,
  summarise,
  uploadableBatches,
} from "../../../shared/uploadBatches";
import {
  parseLevels,
  collectMembership,
  isLeafChainValid,
  buildLevelFormValues,
  Level,
  GroupMapRow,
  Membership,
  ColumnPair,
} from "../../../shared/formModel";
import {
  buildOnDemandSegments,
  decideTier,
  effectiveOnDemandTiers,
  splitChain,
  TierSelection,
  validateChain,
} from "../../../shared/folderChain";
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

/**
 * The uploaded file's name, always assembled in one order:
 *
 *   [Project Name] - [Vendor/Customer Name] - [Document Name] - [Date]
 *
 * Document Name is ONE SEGMENT of this, not the whole name. It used to be the
 * entire name, auto-filled from the other fields until the user typed over it —
 * which meant two people filing the same kind of document could end up with
 * unrelated names, and the ordering guarantee the client wants was impossible.
 *
 * Empty parts are dropped rather than leaving " -  - " gaps, so a document with
 * no project still reads "Acme - Invoice - 03-08-26".
 */
const composeUploadBase = (
  project: string,
  vendor: string,
  docName: string,
  iso: string,
): string =>
  [project.trim(), vendor.trim(), docName.trim(), dateDDMMYY(iso)]
    .filter(Boolean)
    .join(" - ");

type TermOption = { id: string; label: string };
type ToastType = "error" | "success";

type UploadMode = {
  key: string;
  label: string;
  side: "BusinessSegment" | "Project";
  termSetGuid: string;
  stagingFolder: string;
  /**
   * PERMISSIONED tiers only — the cascade down the segment term tree, ending at the
   * Unit folder. Deliberately not the whole chain: every use of this array indexes
   * it against `levelValues` / `levelChoices`, which the cascade fills, and the
   * cascade only ever walks permissioned tiers. Putting below-Unit entries in here
   * would shift every index by one and mis-label the whole form.
   */
  levels: Level[];
  /**
   * The full authored chain, permissioned prefix and below-Unit suffix together.
   * Optional because the DEFAULT_MODES fallbacks carry no below-Unit tiers — they
   * only surface when DMS Config is unreadable, and a site in that state gets the
   * legacy Year -> Document Type pair. Absent means "same as `levels`".
   */
  chain?: Level[];
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
  // Restricted users: their fully-authorised upload paths (segment→…→leaf), filtered
  // down to the ones whose leaf folder actually EXISTS. A path the user is authorised
  // for but which has no folder yet is withheld, not offered — spec
  // 2026-08-12-provisioned-segment-visibility-design.md.
  const [validPaths, setValidPaths] = useState<ValidPath[]>([]);
  // True when the user IS authorised somewhere but every one of those paths was
  // withheld for want of a folder. It is the only way to tell "you have no groups"
  // (a membership problem) from "your folders were never created" (an admin task) —
  // and they send the administrator to two completely different places.
  const [awaitingFolders, setAwaitingFolders] = useState<boolean>(false);
  /**
   * Folder Map sections (each row's segment top folder), kept for the ADMIN-ONLY "not set
   * up yet" marker on the segment picker — spec §9. `null` means the list could not be
   * read, and that must stay silent: marking every segment unbuilt over a transient error
   * would send an admin to reconcile an intact tree.
   *
   * Only the sections are held, not the rows — it is the one thing the picker needs, and
   * keeping up to 5,000 mapping rows alive for the life of the form to answer it is waste.
   */
  const [mapSections, setMapSections] = useState<{ section?: string }[] | null>(
    null,
  );

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [uploadMode, setUploadMode] = useState<string>("");
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);
  const [file, setFile] = useState<File | undefined>(undefined);
  // One SEGMENT of the final file name, not the whole name — see composeUploadBase.
  const [docName, setDocName] = useState<string>("");
  // Below-Unit tier selections, keyed by the tier's `column`. Replaces the old
  // dedicated documentType/yearPeriod state: those two WERE the hardcoded shape,
  // and a chain that varies in length cannot have one useState per tier.
  const [tierValues, setTierValues] = useState<Record<string, string>>({});
  // Term options keyed by TERM SET GUID, so two tiers bound to the same set cost
  // one fetch and switching modes reuses whatever is already loaded.
  const [termCache, setTermCache] = useState<Record<string, TermOption[]>>({});
  // Children of a term, keyed by PARENT term GUID — feeds cascading below-Unit tiers
  // (SubUnit), whose options depend on the Unit chosen rather than on a term set.
  //
  // `ok` separates "this term genuinely has no children" from "the call failed". Not every
  // Unit has SubUnits, so an empty list is a real answer that SKIPS the tier — and if a
  // failure produced that same empty list, a transient error would file a document one tier
  // too shallow, into a folder that exists and looks right. Failures stay unresolved and
  // block the upload instead.
  const [childCache, setChildCache] = useState<Record<string, { terms: TermOption[]; ok: boolean }>>({});
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [vendor, setVendor] = useState<string>("");
  const [projectName, setProjectName] = useState<string>("");
  const [remark, setRemark] = useState<string>("");
  const [legallyPrivileged, setLegallyPrivileged] = useState<boolean>(false);
  // ── Batches ─────────────────────────────────────────────────────────────────
  // Spec: 2026-08-15-batched-multi-file-upload-design.md. A batch is ONE destination folder plus the
  // files that belong in it; an upload carries several. The metadata fields below are the EDITOR for
  // whichever staged file is open, and `draftFiles` is the batch currently being built.
  //
  // Nothing here reaches SharePoint until Upload, and staged work cannot be persisted because a `File`
  // is not serialisable — hence the beforeunload guard, which is the whole mitigation.
  const [batches, setBatches] = useState<Batch[]>([]);
  const [draftFiles, setDraftFiles] = useState<StagedFile[]>([]);
  const [activeFileId, setActiveFileId] = useState<string>("");
  const [lastRun, setLastRun] = useState<{ ok: number; failed: number } | undefined>(undefined);
  const [dragOver, setDragOver] = useState<boolean>(false);
  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [toast, setToast] = useState<{
    message: string;
    type: ToastType;
  } | null>(null);
  const toastTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // The "Replace Existing File" prompt was REMOVED on 2026-08-15 with batching. A batch runs
  // unattended, so there is nobody to answer it, and the file it offered to overwrite may already be
  // Approved and routed to Documents — replacing that silently destroys a record an approver has
  // acted on. A name clash now fails THAT ONE FILE with its reason and lets its siblings through;
  // the uploader renames it and presses Upload again.

  const showToast = (message: string, type: ToastType): void => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToast({ message, type });
    // Errors self-dismiss; the success state must NOT. It renders as a modal
    // with two deliberate exits ("Back to Document" / "Upload More"), so a timer
    // yanks the dialog out from under someone mid-decision — and if it fires
    // while they are still reading, a completed upload looks like it vanished.
    if (type !== "success") {
      toastTimerRef.current = setTimeout(() => setToast(null), 5000);
    }
  };

  const activeMode = (): UploadMode | undefined =>
    modes.find((m) => m.key === uploadMode);

  /* ---------- Below-Unit tiers -------------------------------------------- */

  /**
   * The folders created beneath the Unit folder, in path order, for the active mode.
   *
   * A mode with no configured chain gets the synthetic [Year, Document Type] pair,
   * so an unmigrated site behaves exactly as before and there is ONE walk rather
   * than a legacy branch and a configured branch.
   */
  const belowUnitTiers = (): Level[] =>
    effectiveOnDemandTiers(
      activeMode()?.chain ?? activeMode()?.levels ?? [],
      settings.termSets.yearPeriod,
      settings.termSets.documentType,
    );

  /**
   * The term whose CHILDREN feed a cascading below-Unit tier (one with no `termSet`).
   * For the first such tier that is the deepest permissioned selection — the Unit;
   * for a later one, the tier immediately above it.
   *
   * This is what makes the client's SubUnit work: each Unit has its own SubUnits, so
   * those terms live inside the segment tree rather than in a flat set of their own.
   */
  /** The deepest permissioned selection — the Unit, in every configured segment so far. */
  const permissionedLeafTerm = (): string => {
    const perm = activeMode()?.levels ?? [];
    return perm.length > 0 ? levelValues[perm.length - 1] ?? "" : "";
  };

  /**
   * Walk the below-Unit chain, dropping tiers that do not apply here.
   *
   * A cascading tier applies only where the term above it HAS children — not every Unit
   * has SubUnits (client, 2026-08-10). Each tier's parent is the previous *applicable*
   * tier's selection, so a skipped tier does not orphan the one beneath it.
   *
   * `unresolved` collects tiers whose options are not known yet, or whose lookup failed.
   * Those must block the upload: treating "unknown" as "does not apply" would quietly
   * file the document a level too shallow.
   */
  const tierPlan = (): { tiers: Level[]; parents: string[]; unresolved: string[] } => {
    const tiers: Level[] = [];
    const parents: string[] = [];
    const unresolved: string[] = [];
    let parent = permissionedLeafTerm();
    for (const t of belowUnitTiers()) {
      const flat = (t.termSet ?? "").trim();
      if (flat) {
        tiers.push(t);
        parents.push("");
        continue;
      }
      // Until the tier above is chosen, whether this tier even APPLIES is unknown — so it
      // is hidden rather than shown disabled. Showing it would assert "this unit has
      // subunits" before anything says so, and on a unit that turns out to have none it
      // would then disappear, which reads as the form losing a field.
      //
      // Safe to leave out of `unresolved`: with no parent selected the tier above is itself
      // blank and already required, so the upload is blocked by that instead.
      if (!parent) continue;
      const entry = childCache[parent.trim().toLowerCase()];
      const decision = decideTier(t, entry !== undefined && entry.ok ? entry.terms.length : undefined);
      if (decision === "skip") continue;
      tiers.push(t);
      parents.push(parent);
      if (decision === "unresolved") unresolved.push(t.label);
      parent = tierValues[t.column] ?? "";
    }
    return { tiers, parents, unresolved };
  };

  /**
   * Options for one applicable tier. `termSet` present → that set's top terms, flat.
   * Absent → the children of the term above, cascading. Presence of `termSet` is the
   * discriminator; there is no separate flag (see the `termSet` comment on `Level`).
   */
  const tierOptions = (t: Level, parent: string): TermOption[] => {
    const set = (t.termSet ?? "").trim().toLowerCase();
    if (set) return termCache[set] ?? [];
    const key = parent.trim().toLowerCase();
    return key ? childCache[key]?.terms ?? [] : [];
  };

  /**
   * Only selections that still exist in their tier's CURRENT options are returned.
   *
   * Deliberate, and the safety net for a stale cascade: change the Unit and a
   * previously chosen SubUnit belongs to a different unit entirely. Rather than filing
   * under it, the selection drops out here, `buildOnDemandSegments` reports that tier
   * as missing, and the upload is blocked naming it. The loud failure — never a file
   * written into another unit's subunit folder.
   */
  const tierSelections = (): Record<string, TierSelection | undefined> => {
    const out: Record<string, TierSelection | undefined> = {};
    const plan = tierPlan();
    plan.tiers.forEach((t, i) => {
      const id = tierValues[t.column] ?? "";
      const opt = tierOptions(t, plan.parents[i]).find((o) => o.id === id);
      if (opt) out[t.column] = { id: opt.id, label: opt.label };
    });
    return out;
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

  // Fetch the term sets a configured chain adds. Runs on every mode change but does
  // nothing on an unmigrated site, where both sets are already cached from mount.
  // Failures are swallowed per set: one unreachable tier must not blank the others,
  // and the empty-state below names whichever tier came back with no options.
  // Declared after loadTermSet deliberately — a const arrow is not hoisted.
  useEffect(() => {
    let cancelled = false;
    const missing = belowUnitTiers()
      .map((t) => (t.termSet ?? "").trim())
      .filter((g) => g && !termCache[g.toLowerCase()]);
    if (missing.length === 0) return;
    (async () => {
      const loaded: Record<string, TermOption[]> = {};
      for (const guid of missing) {
        // Sequential, and Promise.allSettled is unavailable on this tsconfig
        // (CLAUDE.md gotcha #3) — a per-set try/catch is the supported shape.
        try {
          loaded[guid.toLowerCase()] = await loadTermSet(guid);
        } catch {
          loaded[guid.toLowerCase()] = [];
        }
      }
      if (!cancelled) setTermCache((prev) => ({ ...prev, ...loaded }));
    })().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [uploadMode, modes, settings.termSets.yearPeriod, settings.termSets.documentType]);

  const loadTermChildrenRef = useRef<
    ((termSetId: string, termId: string) => Promise<TermOption[]>) | undefined
  >(undefined);

  // Load the children a cascading below-Unit tier needs. Keyed on the selections
  // above it, so choosing a different Unit fetches that unit's own SubUnits.
  // Does nothing at all when no tier cascades, which is every site today.
  useEffect(() => {
    const md = activeMode();
    const load = loadTermChildrenRef.current;
    if (!md || !load) return;
    const needed: string[] = [];
    const plan = tierPlan();
    plan.tiers.forEach((t, i) => {
      if ((t.termSet ?? "").trim()) return;
      const parent = plan.parents[i].trim();
      if (parent && !childCache[parent.toLowerCase()]) needed.push(parent);
    });
    if (needed.length === 0) return;
    let cancelled = false;
    (async () => {
      const loaded: Record<string, { terms: TermOption[]; ok: boolean }> = {};
      for (const parent of needed) {
        // Per-parent try/catch: Promise.allSettled is unavailable on this tsconfig
        // (gotcha #3), and one unreachable tier must not blank the others.
        //
        // ok:false on failure, NOT an empty list. An empty list is a real answer that
        // skips the tier, so a failure recorded as empty would silently shorten the path.
        try {
          loaded[parent.toLowerCase()] = { terms: await load(md.termSetGuid, parent), ok: true };
        } catch {
          loaded[parent.toLowerCase()] = { terms: [], ok: false };
        }
      }
      if (!cancelled) setChildCache((prev) => ({ ...prev, ...loaded }));
    })().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [uploadMode, modes, levelValues, tierValues, childCache]);

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

  // Hand the loader to the cascading-tier effect above. A const arrow is not hoisted,
  // so the effect cannot call it directly without a use-before-define error, and
  // moving the effect below every loader would separate it from the other tier state.
  loadTermChildrenRef.current = loadTermChildren;

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
   *
   * Primed here rather than in a mount effect so there is no ordering dependency: every loader
   * primes, and after the first the cache short-circuits. Verified live 2026-08-05 — the client's
   * site has CRS lists and CRS permission levels but a DMS content type, DMS_SITE_MEMBERS and a
   * DMS entity type, which is why the prefix is resolved per artefact and never applied globally.
   */
  const listName = async (suffix: string): Promise<string> => {
    await primeNames(context.spHttpClient, siteUrl);
    return encodeURIComponent(cachedListTitle(suffix));
  };

  /**
   * A comparable fingerprint of a folder chain — order, names and columns.
   *
   * Compared rather than deep-equalled so a re-saved row with identical content does not block an
   * upload, while anything that would move a folder or write a different column does.
   */
  const chainSignature = (levels: Level[]): string =>
    (levels ?? [])
      .map((l) =>
        [
          l.label,
          l.column,
          l.labelCol ?? "",
          l.tidCol ?? "",
          l.termSet ?? "",
          l.permissioned === false ? "0" : "1",
        ].join("|"),
      )
      .join(">");

  /**
   * Re-read this mode's chain from DMS Config, for the staleness guard in the upload handler.
   *
   * `undefined` means "could not tell" — a failed read, or a row that has gone. The caller must NOT
   * block on that: refusing an upload because a config read failed takes the form down over a
   * transient error, and the wrong-shape risk it guards against is far rarer.
   */
  const freshChainFor = async (key: string): Promise<Level[] | undefined> => {
    try {
      const config = await listName(LIST_SUFFIX.config);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${config}')/items?$select=Levels` +
          `&$filter=ConfigType eq 'mode' and Title eq '${encodeURIComponent(key)}'&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return undefined;
      const rows = ((await res.json()).value ?? []) as Array<{ Levels?: string }>;
      if (rows.length === 0) return undefined;
      const parsed = parseLevels(rows[0].Levels ?? "");
      return parsed.length > 0 ? parsed : undefined;
    } catch {
      return undefined;
    }
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
        // `levels` is the permissioned prefix only — see the UploadMode comment.
        levels: splitChain(parseLevels(item.Levels)).permissioned,
        chain: parseLevels(item.Levels),
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
      // The library's LIVE title wins over the config row. That row predates the rename and
      // still reads "Staging" on a migrated site, which now 404s — honouring it would break
      // uploads on exactly the sites that migrated correctly. An admin who deliberately set
      // some other name still wins; only the stale legacy literal is ignored.
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
      if (!isLeafChainValid(chain.map((c) => c.id), leaf.termGuid)) continue;

      // A row on a NON-LEAF term (a department) authorises every unit beneath it —
      // the upload-side mirror of reconciliation's departmental fan-out. Without
      // it the two halves of the permission model disagree: reconciliation grants
      // a Head of Department Contribute on every unit folder, and then the form
      // offers them a Unit dropdown with nothing in it, because the chain
      // terminates at the department. Halves that disagree are worse than either
      // half alone — the user is told they cannot do what their folder ACL says
      // they can.
      //
      // This EXTENDS leaf-only authorisation rather than retreating from it. The
      // rule removed on 2026-07-29 demanded a MEMBER row at every tier and refused
      // correctly provisioned uploaders; this only adds paths a deliberate
      // department-tier row already grants, and never infers a row that is absent.
      //
      // Depth is bounded by the mode's Levels chain, so it terminates even if the
      // term store ever returned a cycle. A leaf row costs no extra call: prefix
      // length already equals the chain length, so expand() pushes and returns.
      const expand = async (prefix: TermOption[]): Promise<void> => {
        if (prefix.length >= mode.levels.length) {
          out.push({ modeKey: mode.key, chain: prefix });
          return;
        }
        const kids = await loadTermChildren(
          mode.termSetGuid,
          prefix[prefix.length - 1].id,
        ).catch(() => [] as TermOption[]);
        if (kids.length === 0) {
          // A term with no children IS a leaf, whatever the Levels chain claims.
          // A half-populated term store must not cost a user their own unit.
          out.push({ modeKey: mode.key, chain: prefix });
          return;
        }
        for (const k of kids) await expand([...prefix, k]);
      };
      await expand(chain);
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
      const [
        loadedSettings,
        rawModes,
        groupMap,
        userGroupIds,
        admin,
        folderMapRows,
      ] = await Promise.all([
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
          // Which folders EXIST. `null` on failure means "unknown", and unknown must
          // offer everything — an unreadable list proves nothing, and silently
          // emptying every dropdown takes the form down for the whole site. Same rule
          // as the stale-chain guard and AllowedFileTypes: empty is not unknown.
          loadFolderMapRows(context.spHttpClient, siteUrl).catch((err) => {
            console.error(
              "DMS Folder Map read failed — cannot tell which folders exist, so every authorised path will be offered.",
              err,
            );
            return null as FolderMapRow[] | null;
          }),
        ]);
      setSettings(loadedSettings);
      // Ignore config rows that predate the Side/Levels schema (empty Levels) —
      // fall back to the built-in modes so the form never renders a broken cascade.
      const usable = rawModes.filter(
        (m: UploadMode) => m.levels.length > 0 && !!m.termSetGuid,
      );
      const loadedModes = usable.length > 0 ? usable : DEFAULT_MODES;
      setModes(loadedModes);

      // Kept for the admin-only "not fully set up yet" marker (spec §9). Set for EVERY user,
      // including the privileged branch below which resolves no paths — that branch is
      // exactly the one the marker exists for.
      setMapSections(
        folderMapRows ? folderMapRows.map((r) => ({ section: r.section })) : null,
      );

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
      // Seed the by-term-set cache with the two sets every site has. A configured
      // chain's extra tiers are fetched by the effect below, which skips anything
      // already here — so an unmigrated site makes no additional requests at all.
      setTermCache({
        [loadedSettings.termSets.documentType.trim().toLowerCase()]: docTypes,
        [loadedSettings.termSets.yearPeriod.trim().toLowerCase()]: years,
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
        const authorised = await resolveValidPaths(loadedModes, membership);

        // Withhold any path whose leaf folder does not exist yet. Authorisation says
        // where the user MAY file; the Folder Map says where they CAN. Offering the
        // difference is what let a brand-new segment appear before it had folders, and
        // what let a newly grouped user pick a unit before reconciliation had run —
        // both ending in a failure at upload time that the uploader cannot act on.
        //
        // The gate is at the LEAF only. The tiers above it are derived from the
        // surviving paths, so a segment with nothing provisioned empties itself and
        // disappears from the picker with no per-segment rule.
        const provisioned = filterProvisionedPaths(
          authorised,
          mappedTermGuidSet(folderMapRows),
        );

        // Existence is not enough. Reconciliation creates folders from the term tree and
        // grants group ACLs in a SEPARATE pass, so a brand-new group's unit folder
        // usually already exists — it is the ACL that is missing. An existence check
        // waves those through (verified live 2026-08-12) and the uploader meets the 403
        // anyway. So ask the folder what THIS user may do with it.
        //
        // One request per surviving path, in parallel. A PIC has one or two; the
        // department fan-out that produces many belongs to Documents-side viewers, who
        // are not uploaders. Probes that cannot reach the server come back "unknown" and
        // are KEPT — a throttle must never empty the form.
        const reachable = filterReachablePaths(
          provisioned.paths,
          await Promise.all(
            provisioned.paths.map(async (p): Promise<AccessVerdict> => {
              const leaf = p.chain[p.chain.length - 1];
              const row = (folderMapRows ?? []).find(
                (r) => normalizeTermGuid(r.termGuid) === normalizeTermGuid(leaf?.id),
              );
              if (!row?.folderUniqueId) return "missing";
              return probeFolderUploadAccess(
                context.spHttpClient,
                siteUrl,
                row.folderUniqueId,
              ).catch(() => "unknown" as AccessVerdict);
            }),
          ),
        );

        const paths = reachable.paths;
        setValidPaths(paths);
        // Only claim the folders are not ready when we actually KNOW it: an unreadable
        // Folder Map or an inconclusive probe withholds nothing, and saying otherwise
        // would send an admin to reconcile an intact, correctly permissioned tree.
        setAwaitingFolders(
          provisioned.known &&
            reachable.known &&
            paths.length === 0 &&
            authorised.length > 0,
        );

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

  // Plain setters. These used to copy a composed name into the Document Name box
  // and stop as soon as the user typed; the name is now assembled at upload from
  // all four parts, so there is nothing to keep in sync and no "has the user
  // edited this yet" state to get wrong.
  const onProjectNameChange = (v: string): void => setProjectName(v);
  const onVendorChange = (v: string): void => setVendor(v);
  const onDocumentDateChange = (iso: string): void => setDocumentDate(iso);
  const onDocNameChange = (v: string): void => setDocName(v);

  /**
   * Single gate for a chosen file, whether it arrived from the picker or a drop.
   *
   * The extension check has to live here rather than only on the <input accept>:
   * `accept` filters the file dialog but is advisory, and a DROP bypasses it
   * entirely — so without this a drag-and-drop would happily hand an .exe to the
   * upload. See CLAUDE.md gotcha #10 for how the two checks disagreed before.
   */
  /* ---------- Batch editor bridge ----------------------------------------- */

  /**
   * The metadata fields are ONE editor, reused for whichever staged file is open. `captureEditor` and
   * `applyEditor` move values between that editor and the file.
   *
   * Kept as an explicit capture/apply pair rather than binding each input straight to
   * `draftFiles[i].meta`: every existing `onChange` handler (`onDocNameChange`, `onVendorChange`, …)
   * does more than `setState`, and rewiring all of them for a UI change would put site-verified
   * validation at risk for no gain.
   */
  const captureEditor = (): FileMeta => ({
    docName,
    projectName,
    vendor,
    remark,
    documentDate,
    confidentiality,
    // A boolean through a string map. Re-derived at upload time anyway (see `privilegedApplies`),
    // because hiding the tick does not clear the state behind it.
    legallyPrivileged: legallyPrivileged ? "1" : "",
  });

  const applyEditor = (m: FileMeta): void => {
    setDocName(m.docName ?? "");
    setProjectName(m.projectName ?? "");
    setVendor(m.vendor ?? "");
    setRemark(m.remark ?? "");
    setDocumentDate(m.documentDate ?? "");
    setConfidentiality(m.confidentiality ?? "");
    setLegallyPrivileged((m.legallyPrivileged ?? "") !== "");
  };

  /** The name a staged file will actually be saved under — the composed convention, not the typed name. */
  const finalNameFor = (f: File, m: FileMeta): string =>
    buildUploadName(
      f.name,
      composeUploadBase(m.projectName ?? "", m.vendor ?? "", m.docName ?? "", m.documentDate ?? ""),
    );

  /** Fold the editor back into the file it belongs to. Returns the updated list. */
  const commitEditor = (files: StagedFile[]): StagedFile[] => {
    if (!activeFileId) return files;
    const meta = captureEditor();
    return files.map((f) =>
      f.id === activeFileId
        ? { ...f, typedName: meta.docName ?? "", meta, finalName: finalNameFor(f.file, meta) }
        : f,
    );
  };

  /** Open a different file's panel, saving the one being left. */
  const selectFile = (id: string): void => {
    const saved = commitEditor(draftFiles);
    setDraftFiles(saved);
    if (id === activeFileId) {
      // Collapsing the open panel. The editor keeps its values; nothing is lost.
      setActiveFileId("");
      return;
    }
    const target = saved.find((f) => f.id === id);
    if (target) applyEditor(target.meta);
    setActiveFileId(id);
  };

  const stagedNow = stagedTotals([
    ...batches,
    { id: "draft", segmentKey: "", chainSignature: "", pathLabels: [], destination: {}, files: draftFiles },
  ]);
  const draftCollisions = collisionsWithin({
    id: "draft", segmentKey: "", chainSignature: "", pathLabels: [], destination: {},
    files: commitEditor(draftFiles),
  });
  const crossBatchDupes = duplicateAcrossBatches(batches);

  /**
   * The only protection staged work has.
   *
   * A `File` cannot be serialised, so there is no draft to restore and no localStorage fallback — a
   * closed tab loses every staged batch. The browser owns this dialog's wording; all we control is
   * whether it appears.
   */
  useEffect(() => {
    if (stagedNow.files === 0) return undefined;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [stagedNow.files]);

  /**
   * Is this file an allowed type — and if not, refuse it and record why.
   *
   * Split out from staging so a multi-select can check every file in one pass. Returns false for a
   * refusal, which is what keeps a disallowed type from ever entering a batch.
   */
  const admitFile = (picked: File | undefined): boolean => {
    if (!picked) return false;
    const types =
      settings.allowedFileTypes.kind === "none"
        ? []
        : settings.allowedFileTypes.types;
    if (!types.some((ext) => picked.name.toLowerCase().endsWith(ext))) {
      showToast(`File type not allowed. Allowed: ${types.join(", ")}`, "error");
      // Recorded — and this is the ONE event class Purview cannot see either: the file never reaches
      // SharePoint, so no server-side audit of it exists anywhere. Someone repeatedly offering an
      // .exe is exactly what an audit log should surface.
      //
      // No banner if this write fails, deliberately: an uploader can neither fix nor act on a logging
      // problem, and the refusal itself has already been reported to them.
      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.uploadRefused,
        outcome: "Refused",
        source: "UploadForm",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        library: libraryTitle(),
        itemName: picked.name,
        summary: `Upload refused — ${picked.name}`,
        details: [
          `Extension offered: ${getExtension(picked.name) || "(none)"}`,
          `Allowed at the time: ${types.length > 0 ? types.join(", ") : "(none — every upload is blocked)"}`,
          "The file was never uploaded.",
        ],
      }).catch(() => undefined);
      setFile(undefined);
      if (fileRef.current) fileRef.current.value = "";
      return false;
    }
    setStatus("");
    return true;
  };

  /**
   * Stage every picked file into the batch being built.
   *
   * Accumulates in ONE pass rather than calling a per-file helper in a loop: each call would read the
   * same stale `draftFiles`, so a five-file selection would stage only the last one — and it would look
   * like the picker had silently dropped four files.
   *
   * Each new file inherits the last-typed values in this batch (copied from a sibling the user typed,
   * visible and editable). The typed NAME is never inherited — that is the collision `collisionsWithin`
   * exists to block.
   */
  const acceptFiles = (list?: FileList | ReadonlyArray<File> | null): void => {
    if (!list) return;
    const picked: File[] = [];
    for (let i = 0; i < list.length; i++) {
      const f = (list as FileList)[i];
      if (admitFile(f)) picked.push(f);
    }
    if (picked.length === 0) return;

    let acc = commitEditor(draftFiles);
    let lastAdded: StagedFile | undefined;
    for (const f of picked) {
      const meta: FileMeta = {
        ...inheritDefaults({
          id: "draft", segmentKey: "", chainSignature: "", pathLabels: [], destination: {}, files: acc,
        }),
        docName: "",
      };
      lastAdded = { id: nextId("f"), file: f, typedName: "", meta, finalName: finalNameFor(f, meta) };
      acc = [...acc, lastAdded];
    }
    setDraftFiles(acc);
    // Open the last one added, so a single pick behaves exactly as before: choose a file, start typing.
    if (lastAdded) {
      applyEditor(lastAdded.meta);
      setActiveFileId(lastAdded.id);
      setFile(lastAdded.file);
    }
    if (fileRef.current) fileRef.current.value = "";
  };

  const resetForm = (): void => {
    setFile(undefined);
    setDocName("");
    setTierValues({});
    // Rebuild the cascade rather than blanking it. A restricted uploader's
    // Department/Unit are DERIVED from their group memberships and locked, so
    // clearing them to [] left the two fields empty with no way to re-select —
    // the tier detection runs once at mount, so only a page refresh brought them
    // back. That made "Upload More" strictly worse than reloading the form.
    const m = activeMode();
    if (!m) {
      setLevelValues([]);
    } else if (privileged) {
      initCascade(m).catch(() => {
        setLevelChoices([]);
        setLevelValues([]);
      });
    } else {
      applyRestrictedMode(validPaths, m, []);
    }
    setDocumentDate("");
    setConfidentiality("");
    setVendor("");
    setProjectName("");
    setRemark("");
    setLegallyPrivileged(false);
    if (fileRef.current) fileRef.current.value = "";
  };

  /* ---------- Upload ------------------------------------------------------ */

  /**
   * Per-file required fields, checked for EVERY staged file rather than only the one on screen.
   *
   * Returns the missing field names, prefixed with the file when there is more than one — "please
   * complete: Vendor/Customer Name" is unhelpful when four files are staged and one of them is short.
   */
  const missingForFile = (sf: StagedFile, many: boolean): string[] => {
    const meta = sf.id === activeFileId ? captureEditor() : sf.meta;
    const out: string[] = [];
    if (!(meta.documentDate ?? "")) out.push("Document Date");
    if (!(meta.confidentiality ?? "")) out.push("Confidential Level");
    // The three name parts are required so every saved file carries the full
    // [Project] - [Vendor] - [Document Name] - [Date] shape. composeUploadBase drops blank parts, so
    // leaving these optional silently produces a shorter name than the convention promises — and after
    // the fact a shortened name is indistinguishable from a deliberate one.
    if (!(meta.docName ?? "").trim()) out.push("Document Name");
    if (!(meta.projectName ?? "").trim()) out.push("Project Name");
    if (!(meta.vendor ?? "").trim()) out.push("Vendor/Customer Name");
    return many ? out.map((n) => `${sf.file.name}: ${n}`) : out;
  };

  /**
   * Stage the current draft as a batch.
   *
   * The DESTINATION IS SNAPSHOT HERE, not referenced. By Upload the pickers describe whatever batch is
   * being edited then, so a live read would write batch 1 into batch 3's department — silently, into a
   * folder that exists and looks correct.
   *
   * What is snapshot is the LEAF TERM ID, never the resolved folder URL. Resolving the folder from its
   * term at upload time is what makes this rename-proof today; freezing a URL at save would quietly
   * undo that for any folder renamed in between.
   */
  const saveBatch = (): boolean => {
    const files = commitEditor(draftFiles);
    if (files.length === 0) {
      showToast("Add at least one document to this batch.", "error");
      return false;
    }

    const m = activeMode();
    const missing: string[] = [];
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) missing.push(lvl.label);
    });
    missing.push(...buildOnDemandSegments(tierPlan().tiers, tierSelections()).missing);
    const many = files.length > 1;
    for (const sf of files) missing.push(...missingForFile(sf, many));
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return false;
    }

    // Same block as before, still checked here: a rename can turn an allowed file into a name whose
    // extension no longer matches, and the composed name is what actually gets sent.
    const allowed = settings.allowedFileTypes;
    if (allowed.kind === "none") {
      showToast(NO_TYPES_MESSAGE, "error");
      return false;
    }
    for (const sf of files) {
      const fn = sf.finalName ?? sf.file.name;
      if (!allowed.types.some((ext) => fn.toLowerCase().endsWith(ext))) {
        showToast(`${fn}: file type not allowed. Allowed: ${allowed.types.join(", ")}`, "error");
        return false;
      }
    }

    if (!m || m.levels.length === 0) {
      showToast("No upload mode configured.", "error");
      return false;
    }
    const chainError = validateChain(m.chain ?? m.levels ?? []);
    if (chainError) {
      showToast(
        `The folder structure for this segment is not set up correctly: ${chainError.message} ` +
          `Ask an administrator to check the DMS Config mode row.`,
        "error",
      );
      return false;
    }

    const plan = tierPlan();
    if (plan.unresolved.length > 0) {
      showToast(
        `Still checking ${plan.unresolved.join(" and ")} for this unit. ` +
          `If this does not clear, reload the page before saving this batch.`,
        "error",
      );
      return false;
    }
    const { segments } = buildOnDemandSegments(plan.tiers, tierSelections());

    const leafIdx = m.levels.length - 1;
    const leafTerm = (levelChoices[leafIdx] ?? []).find((o) => o.id === levelValues[leafIdx]);
    if (!leafTerm) {
      showToast("Please choose all folder levels before saving this batch.", "error");
      return false;
    }

    // Level label/GUID pairs, exactly as the single-file path built them.
    const levelSelections = (m.levels ?? []).map((lvl, i) => {
      const opt = (levelChoices[i] ?? []).find((o) => o.id === levelValues[i]);
      return {
        column: lvl.column,
        label: opt?.label ?? "",
        id: opt?.id ?? "",
        labelCol: lvl.labelCol,
        tidCol: lvl.tidCol,
      };
    });
    levelSelections.unshift({
      column: "BusinessSegment",
      label: m.label,
      id: m.termSetGuid,
      labelCol: undefined,
      tidCol: undefined,
    });

    // Below-Unit tier metadata. Only APPLICABLE tiers write: a unit with no subunits leaves SubUnit
    // empty rather than storing a value from some other unit's list.
    const tierFormValues: Array<{ FieldName: string; FieldValue: string }> = [];
    plan.tiers.forEach((t, i) => {
      const id = tierValues[t.column] ?? "";
      const opts = tierOptions(t, plan.parents[i]);
      const opt = opts.find((o) => o.id === id);
      const col = t.labelCol ?? t.column;
      if (!col || !opt) return;
      if (t.tidCol) {
        tierFormValues.push({ FieldName: col, FieldValue: opt.label });
        tierFormValues.push({ FieldName: t.tidCol, FieldValue: opt.id });
      } else {
        tierFormValues.push({ FieldName: col, FieldValue: toTaxValue(opts, id) });
      }
    });

    const destination: BatchDestination = {
      leafTermId: leafTerm.id,
      leafLabel: leafTerm.label,
      segments,
      levelSelections,
      tierFormValues,
    };

    const b: Batch = {
      id: nextId("b"),
      segmentKey: m.key,
      chainSignature: chainSignature(m.chain ?? m.levels ?? []),
      pathLabels: [...levelSelections.slice(1).map((l) => l.label), ...segments],
      destination,
      files,
    };
    if (!canSaveBatch(b)) {
      // The only remaining reason is an internal name clash, which the rows already flag.
      showToast("Two documents in this batch would be saved under the same name.", "error");
      return false;
    }

    setBatches([...batches, b]);
    setDraftFiles([]);
    setActiveFileId("");
    setFile(undefined);
    resetForm();
    showToast(`Batch saved — ${b.files.length} document${b.files.length === 1 ? "" : "s"} ready.`, "success");
    return true;
  };

  /**
   * Upload ONE staged file into an already-resolved folder.
   *
   * This is the old single-file handler's tail with two changes and no others: its values come from
   * `sf.meta` instead of component state, and it RETURNS a result instead of calling `showToast` and
   * returning. How a file is uploaded and tagged is untouched — that code is the most site-exercised
   * in the project, and a UI change is no reason to rewrite it.
   */
  const uploadStagedFile = async (
    b: Batch,
    folderId: string,
    sf: StagedFile,
  ): Promise<UploadResult> => {
    const meta = sf.meta;
    const finalName = sf.finalName ?? sf.file.name;
    const dest = b.destination as {
      levelSelections: Array<{ column: string; label: string; id: string; labelCol?: string; tidCol?: string }>;
      tierFormValues: Array<{ FieldName: string; FieldValue: string }>;
    };

    try {
      // NEVER overwrite. A clash fails this one file and lets its siblings through: the existing
      // document may already be Approved and routed to Documents, so replacing it silently would
      // destroy a record an approver has already acted on. The single-file form asked "replace?";
      // with a batch running unattended there is nobody to ask.
      const existsRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (existsRes.ok) {
        return {
          fileId: sf.id,
          ok: false,
          error: `A document called "${finalName}" is already in this folder — rename it, or ask an approver about the existing one.`,
        };
      }
    } catch {
      // Network error on the existence check — proceed; the upload will surface the real error.
    }

    try {
      const uploadRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=false)?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1,
        { body: sf.file },
      );
      if (!uploadRes.ok) {
        let detail = `HTTP ${uploadRes.status}`;
        try {
          const bodyText = await uploadRes.text();
          try {
            const errJson = JSON.parse(bodyText);
            const spMsg = errJson?.error?.message?.value ?? errJson?.error?.message;
            detail += spMsg ? ` — ${spMsg}` : bodyText ? ` — ${bodyText.slice(0, 300)}` : "";
          } catch {
            if (bodyText) detail += ` — ${bodyText.slice(0, 300)}`;
          }
        } catch {
          /* body already consumed or unreadable */
        }
        console.error("Upload failed:", folderId, detail);
        const hint =
          uploadRes.status === 404
            ? " The mapped folder may have been deleted — ask an administrator to re-run reconciliation."
            : "";
        return { fileId: sf.id, ok: false, error: `${detail}.${hint}` };
      }
      const uploadJson = await uploadRes.json();
      const uploadedServerRelativeUrl = uploadJson.ServerRelativeUrl;

      const itemRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=Id&@f='${encodeServerRelativePath(uploadedServerRelativeUrl)}'`,
        SPHttpClient.configurations.v1,
      );
      if (!itemRes.ok) {
        return { fileId: sf.id, ok: false, error: "Uploaded, but could not retrieve the item to tag." };
      }
      const item = await itemRes.json();

      const levelCols: Record<string, ColumnPair> = {
        ...LEVEL_COLUMNS,
        BusinessSegment: {
          label: settings.columns.businessSegmentLabel,
          tid: settings.columns.businessSegmentTid,
        },
      };

      const formValues: Array<{ FieldName: string; FieldValue: string }> = [
        ...dest.tierFormValues,
        {
          FieldName: settings.columns.confidentiality,
          FieldValue: toTaxValue(options.confidentiality, meta.confidentiality ?? ""),
        },
        ...buildLevelFormValues(levelCols, dest.levelSelections),
      ];

      // Guarded because toSpDate("") yields the malformed "NaN/NaN/", which SharePoint rejects with a
      // HasException surfacing far from its cause. Validation already guarantees a value.
      if (meta.documentDate) {
        formValues.push({
          FieldName: settings.columns.documentDate,
          FieldValue: toSpDate(meta.documentDate),
        });
      }

      // Written UNCONDITIONALLY, blanks included — a Text column set to "" is cleared rather than left
      // holding a value from somewhere else.
      formValues.push({ FieldName: settings.columns.projectName, FieldValue: (meta.projectName ?? "").trim() });
      formValues.push({ FieldName: settings.columns.vendor, FieldValue: (meta.vendor ?? "").trim() });
      formValues.push({ FieldName: settings.columns.remark, FieldValue: (meta.remark ?? "").trim() });

      // Re-derived, never trusted from the checkbox: hiding the control does not clear the state behind
      // it, so a file could otherwise carry a legal marker its confidentiality does not offer.
      const privilegedApplies =
        settings.legallyPrivilegedFor !== "" &&
        (meta.confidentiality ?? "") === settings.legallyPrivilegedFor;
      formValues.push({
        FieldName: settings.columns.legallyPrivileged,
        FieldValue: privilegedApplies && (meta.legallyPrivileged ?? "") !== "" ? "true" : "false",
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
        return { fileId: sf.id, ok: false, error: "Uploaded, but tagging metadata failed." };
      }
      const metaJson = await metaRes.json();
      // HTTP 200 even on field errors (gotcha #4) — the exception is per result, not per response.
      const fieldError = (metaJson.value ?? []).find((v: { HasException?: boolean }) => v.HasException);
      if (fieldError) {
        console.error("Field update error:", fieldError);
        return {
          fileId: sf.id,
          ok: false,
          error: `Uploaded, but a field failed: ${fieldError.FieldName} — ${fieldError.ErrorMessage}`,
        };
      }
      return { fileId: sf.id, ok: true };
    } catch (err) {
      console.error("Upload failed:", err);
      return { fileId: sf.id, ok: false, error: (err as Error).message || "Upload failed." };
    }
  };

  /**
   * Upload every staged batch.
   *
   * Each step exists for a reason learned the hard way:
   *
   * 1. Re-read each segment's chain ONCE — two batches on one segment cost one request, not two.
   * 2. Mark the batches whose segment moved under them. They stay staged and ask for a re-pick; the
   *    single-file form said "reload the page", which here would destroy every batch the client built.
   *    A chain that could not be READ marks nothing: unknown is not changed.
   * 3. Per batch, resolve the unit folder from its LEAF TERM (rename-proof) and ensure the below-Unit
   *    folders once — not once per file.
   * 4. Per file, upload and tag, collecting results rather than stopping at the first failure.
   * 5. Fold the results back: success removes, failure stays with its reason. That is what makes Retry
   *    safe — an uploaded file is no longer in the list, so it cannot be sent twice.
   */
  const handleUpload = async (): Promise<void> => {
    // An unsaved draft is the commonest way to lose work here: files staged, destination chosen, then
    // Upload pressed without Save batch. Save it rather than silently ignoring it. `saveBatch`'s
    // setState has not landed yet, so this run stops and the client presses Upload again — one extra
    // click, versus uploading a batch list that does not include what is on screen.
    if (draftFiles.length > 0) {
      if (!saveBatch()) return;
      showToast("Saved the open batch — press Upload again to send everything.", "success");
      return;
    }
    if (batches.length === 0) {
      showToast("Nothing to upload yet. Add documents and save a batch first.", "error");
      return;
    }

    setBusy(true);
    setLastRun(undefined);
    setStatus("Checking the folder structure…");

    // One read per distinct segment. `undefined` for a segment whose chain could not be read.
    const freshBySegment: Record<string, string | undefined> = {};
    const segmentKeys: string[] = [];
    for (const b of batches) if (segmentKeys.indexOf(b.segmentKey) === -1) segmentKeys.push(b.segmentKey);
    for (const key of segmentKeys) {
      const fresh = await freshChainFor(key);
      freshBySegment[key] = fresh ? chainSignature(fresh) : undefined;
    }

    const marked = batchesNeedingRepick(batches, freshBySegment);
    const stale = marked.filter((b) => b.needsRepick);
    const runnable = uploadableBatches(marked);
    if (stale.length > 0) {
      setBatches(marked);
      showToast(
        `The folder structure changed for ${stale.length} batch${stale.length === 1 ? "" : "es"} while ` +
          `this page was open. Choose ${stale.length === 1 ? "its" : "their"} destination again — nothing has been lost.`,
        "error",
      );
      if (runnable.length === 0) {
        setStatus("");
        setBusy(false);
        return;
      }
    }

    const results: UploadResult[] = [];
    let batchNo = 0;
    for (const b of runnable) {
      batchNo++;
      const dest = b.destination as { leafTermId: string; leafLabel: string; segments: string[] };

      setStatus(`Batch ${batchNo} of ${runnable.length} — locating destination folder…`);
      const mapping = await lookupFolderMapping(context.spHttpClient, siteUrl, dest.leafTermId).catch(
        (e: unknown) => {
          console.error("Folder map lookup error:", e);
          return null;
        },
      );
      if (!mapping || !mapping.folderUniqueId) {
        // Names BOTH causes: folder names come from DMS Term Abbreviation, so a unit with no
        // abbreviation row is skipped by every run and re-running reconciliation changes nothing.
        for (const sf of b.files) {
          results.push({
            fileId: sf.id,
            ok: false,
            error: `"${dest.leafLabel}" has no folder yet. An administrator needs to give it an abbreviation in the DMS Term Abbreviation list, then run folder reconciliation.`,
          });
        }
        continue;
      }

      const unitFolder = await resolveMappedFolder(
        context.spHttpClient,
        siteUrl,
        mapping.folderUniqueId,
        mapping.folderUrl,
      );
      const unitSru = unitFolder.serverRelativeUrl;
      if (!unitSru) {
        // Two different problems with two different fixes, and only one of them is reconciliation.
        const why = unitFolder.confirmedMissing
          ? "The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation."
          : `Your folder could not be opened (HTTP ${unitFolder.status}). That is usually a permissions problem — ask an administrator to check your access.`;
        for (const sf of b.files) results.push({ fileId: sf.id, ok: false, error: why });
        continue;
      }

      // Ensure-created ONCE per batch, not per file. Everything below Unit inherits the Unit's ACL.
      setStatus(`Batch ${batchNo} of ${runnable.length} — preparing folders…`);
      let parentSru = unitSru;
      let destFolder: Awaited<ReturnType<typeof ensureFolder>> | undefined;
      let folderError = "";
      for (const name of dest.segments) {
        const made = await ensureFolder(context.spHttpClient, siteUrl, parentSru, name);
        if (!made) {
          folderError = `Could not create the "${name}" folder.`;
          break;
        }
        parentSru = made.serverRelativeUrl;
        destFolder = made;
      }
      if (folderError || !destFolder) {
        const why = folderError || "No destination folder could be resolved for this batch.";
        for (const sf of b.files) results.push({ fileId: sf.id, ok: false, error: why });
        continue;
      }

      let fileNo = 0;
      for (const sf of b.files) {
        fileNo++;
        setStatus(
          `Batch ${batchNo} of ${runnable.length} · file ${fileNo} of ${b.files.length} — ${sf.finalName ?? sf.file.name}`,
        );
        results.push(await uploadStagedFile(b, destFolder.uniqueId, sf));
      }
    }

    // `marked`, not `runnable` — a batch held back for a re-pick has no results and must survive this
    // fold untouched.
    const remaining = applyUploadResults(marked, results);
    setBatches(remaining);
    const counts = summarise(results);
    setLastRun({ ok: counts.ok, failed: counts.failed });
    setStatus("");
    setBusy(false);
    if (counts.failed === 0) {
      showToast(`${counts.ok} document${counts.ok === 1 ? "" : "s"} uploaded and pending review.`, "success");
    } else {
      // Counts, never a verdict. What is left on screen is exactly what still needs doing.
      showToast(
        `Uploaded ${counts.ok} of ${counts.total}. ${counts.failed} could not be uploaded — see the reasons below.`,
        "error",
      );
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

  /**
   * The per-file metadata editor.
   *
   * Rendered INSIDE whichever staged row is open, not once below the list (client, 2026-08-15: each
   * file needs its own name, project, vendor, date and confidentiality). Below the list it read as one
   * form describing all of them, which is exactly the misreading that would give five files one
   * vendor.
   *
   * One instance, moved — not one per row. The fields are backed by the component-level editor state
   * that `selectFile` loads and saves, so duplicating them per row would mean five sets of inputs
   * bound to the same values, all changing together.
   */
  const activeStagedFile = draftFiles.find((x) => x.id === activeFileId);

  const metaEditor = (
    <>
              <div className="dms-grid" style={{ marginTop: 16 }}>
                <label className="dms-field" style={{ gridColumn: "1 / -1" }}>
                  <span>
                    Document Name <em className="req">*</em>
                  </span>
                  {/* Typeable before a file is picked: the name is only read at upload
                      time, and leaving it enabled avoids a greyed-out first field. */}
                  <input
                    type="text"
                    value={docName}
                    maxLength={50}
                    onChange={(e) => onDocNameChange(e.target.value)}
                  />
                  {/* The saved name is assembled from four fields, so showing the result
                      is the only way the uploader can tell what it will be called. Falls
                      back to the original filename when every part is blank, which is
                      exactly what buildUploadName does. */}
                  {/* The ACTIVE staged file, not `file` — that holds the last one added, so editing an
                      earlier row would preview a name belonging to a different document. */}
                  {activeStagedFile ? (
                    <small>
                      Saves as:{" "}
                      {buildUploadName(
                        activeStagedFile.file.name,
                        composeUploadBase(projectName, vendor, docName, documentDate),
                      )}
                    </small>
                  ) : (
                    <small>Max. 50 characters</small>
                  )}
                </label>

                {/* Free-text Project Name — distinct from the Group-led Projects
                    "Group Project Name" folder level in the card below. */}
                <label className="dms-field">
                  <span>
                    Project Name <em className="req">*</em>
                  </span>
                  <input
                    type="text"
                    value={projectName}
                    maxLength={50}
                    onChange={(e) => onProjectNameChange(e.target.value)}
                  />
                  <small>Max. 50 characters</small>
                </label>

                {/* Vendor is free text. It also feeds the auto-composed document name. */}
                <label className="dms-field">
                  <span>
                    Vendor/Customer Name <em className="req">*</em>
                  </span>
                  <input
                    type="text"
                    value={vendor}
                    maxLength={50}
                    onChange={(e) => onVendorChange(e.target.value)}
                  />
                  <small>Max. 50 characters</small>
                </label>

                {/* Document Type used to sit here. It moved into the folder card below:
                    it is part of the destination path (Unit → Year → Document Type), not a
                    property of the document, and grouping it with Unit and Year is what the
                    client's mockup shows. */}

                {/* Remark lives in the folder card above, per the mockup. */}
              </div>

              {/* Document Date | Confidential Level | Legally Privileged — one line. */}
              <div className="dms-detail-row">
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

                {/* Not renderSelect: the info icon belongs on the LABEL, and the label is
                    a <span> holding a real <label htmlFor> so clicking the icon does not
                    fall through and focus the select. */}
                <div className="dms-field">
                  <span className="dms-labelrow">
                    <label htmlFor="dms-form-conf">
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
                        {/* Legally Privileged is NOT defined here any more — it has its own
                            control and its own icon beside it, and defining it in two places
                            invites the two texts to drift apart.
                            Highly Confidential is deliberately absent too. Its term is
                            removed from the term store for Phase 1, so the dropdown cannot
                            offer it, and describing a level nobody can pick reads as a bug
                            in UAT. The definition returns with the HC libraries in Phase 2 —
                            see the highly-confidential-securing design on feat/hc-libraries. */}
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
                    id="dms-form-conf"
                    value={confidentiality}
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

                {/* Offered only for the level named by `legallyPrivilegedFor` in DMS Config.
                    Unset means never offered — see the setting's note. The value is re-derived
                    at upload time rather than trusted from here, because hiding the control
                    does not clear the state behind it. */}
                {settings.legallyPrivilegedFor !== "" &&
                  confidentiality === settings.legallyPrivilegedFor && (
                    <>
                      <label className="dms-lp">
                        <input
                          type="checkbox"
                          checked={legallyPrivileged}
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
    </>
  );

  return (
    <section className="dms-form">
      <style>{`
        .dms-form { max-width: 960px; margin: 32px auto; padding: 0 24px 48px; font-family: 'Segoe UI', sans-serif; }
        .dms-form h2 { margin: 0 0 4px; font-size: 24px; font-weight: 700; color: #1b1b1b; }
        .dms-subtitle { margin: 0 0 28px; font-size: 14px; color: #666; }
        .dms-section { background: #fff; border: 1px solid #e1e1e1; border-radius: 8px; padding: 24px; margin-bottom: 20px; }
        .dms-section-title { font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #0f6c3f; margin: 0 0 16px; }
        /* Drop zone. Stacked and centred while empty so the icon and prompt read as
           one target; switches to a single centred row once a file is chosen, where
           READY / name / size / action belong on one line. */
        .dms-dropzone { display: flex; flex-direction: column; align-items: center; justify-content: center;
          gap: 8px; text-align: center; border: 1px dashed #9bbfaa; border-radius: 10px;
          background: #f2f8f4; padding: 24px 16px; cursor: pointer; font-size: 13px;
          transition: background .12s, border-color .12s; }
        .dms-dropzone:hover, .dms-dropzone:focus-visible { border-color: #0f6c3f; background: #eaf4ee; }
        /* .over fires on dragover — without a visible change there is no confirmation
           the browser will accept the drop, and users let go over the wrong element. */
        .dms-dropzone.over { border-color: #0f6c3f; border-style: solid; background: #e2efe7; }
        .dms-dropzone.has-file { flex-direction: row; gap: 16px; padding: 16px; }
        .dms-dropzone-icon { width: 32px; height: 32px; color: #0f6c3f; }
        .dms-dropzone .name { font-weight: 600; }
        .dms-dropzone .size { color: #666; font-size: 12px; }
        /* READY badge and the trailing action. margin-left:auto pushes "Change
           Document" to the far edge so the row reads name-first, action-last. */
        .dms-filecard-ready { color: #0f6c3f; font-size: 11px; font-weight: 700;
          letter-spacing: .06em; }
        .dms-filecard-action { margin-left: auto; }
        /* Textarea inherits the input styling so Remark matches the fields around it —
           without this it renders in the browser's default monospace at a random width. */
        .dms-field textarea { font: inherit; width: 100%; box-sizing: border-box; padding: 8px 10px;
          border: 1px solid #c8c8c8; border-radius: 4px; resize: vertical; }
        .dms-field textarea:focus { outline: 2px solid #0f6c3f; outline-offset: -1px; }
        .dms-link { background: none; border: none; color: #0f6c3f; cursor: pointer; font-weight: 600; padding: 0; font-size: 13px; }
        .dms-field { display: flex; flex-direction: column; gap: 4px; margin-bottom: 16px; font-size: 13px; }
        .dms-field > span { font-weight: 600; }
        .dms-field .req { color: #d13438; font-style: normal; }
        .dms-field select, .dms-field input[type="text"], .dms-field input[type="date"] { padding: 8px 10px; border: 1px solid #c8c8c8; border-radius: 10px; font: inherit; width: 100%; box-sizing: border-box; height: 38px; background: #fff; }
        .dms-field small { color: #666; font-size: 12px; font-weight: 400; }
        /* Document Date | Confidential Level | Legally Privileged on ONE line, per
           the client mockup. A flex row rather than grid cells because the tick
           exists for only one confidentiality level: when it is absent, date and
           level fall back to an even split that lines up with Project Name and
           Vendor above, so the row does not reflow around a control that is not there. */
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
        /* Two columns, matching the mockup. Three made Document Date and
           Confidential Level share a row with Vendor and pushed the labels tight. */
        .dms-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 24px; }
        /* Folder card runs 2-up so the deepest level and Year pair evenly. */
        .dms-grid-2 { grid-template-columns: 1fr 1fr; }
        /* Unit | Year | Document Type on one row. Together they name exactly one
           destination folder, so they read better as a set than stacked. Unit gets
           the most room because its labels are long unit names, while Year holds
           four characters and needs almost none. */
        .dms-grid-3 { grid-template-columns: 1.8fr 0.9fr 1.3fr; }
        .dms-radio-group { display: flex; gap: 24px; margin-bottom: 20px; }
        .dms-radio-group label { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 600; cursor: pointer; color: #1b1b1b; }
        .dms-radio-group input[type="radio"] { accent-color: #0f6c3f; width: 16px; height: 16px; cursor: pointer; }
        .dms-dept-badge { display: inline-flex; align-items: center; gap: 8px; background: #e8f5ee; border: 1px solid #b3d9c4; border-radius: 20px; padding: 5px 14px; font-size: 13px; margin-bottom: 20px; }
        .dms-dept-badge .dept-label { font-weight: 400; color: #555; }
        .dms-dept-badge .dept-name { font-weight: 700; color: #0f6c3f; }
        .dms-dept-loading { font-size: 13px; color: #666; margin-bottom: 20px; }
        .dms-dept-error { font-size: 13px; color: #d13438; background: #fdf3f3; border: 1px solid #f1c0c0; border-radius: 4px; padding: 10px 14px; margin-bottom: 20px; }
        /* AMBER, not the red above: an unfinished segment is an admin task in progress,
           not a failure, and it is only ever shown to the administrator doing it. */
        .dms-setup-warn { font-size: 13px; line-height: 1.55; color: #7a4f00; background: #fff4e5; border: 1px solid #f0d9b5; border-radius: 4px; padding: 10px 14px; margin-bottom: 20px; }
        .dms-admin-row { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; }
        .dms-admin-badge { display: inline-flex; align-items: center; background: #fff4e5; border: 1px solid #f0c070; border-radius: 20px; padding: 5px 14px; font-size: 13px; font-weight: 700; color: #7a4f00; white-space: nowrap; }
        .dms-admin-row select { padding: 6px 10px; border: 1px solid #c8c8c8; border-radius: 4px; font: inherit; font-size: 13px; }
        .dms-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 4px; flex-wrap: wrap; }

        /* Batches and staged files (2026-08-15). Plain and quiet: this list sits above the fields and
           must not compete with them for attention. */
        .dms-batches { margin-top: 16px; border: 1px solid #d7e3da; border-radius: 8px; padding: 12px 14px; background: #f7fbf8; }
        .dms-batches-head { margin: 0 0 10px; font-size: 12.5px; color: #2f4c3a; }
        .dms-batch { border: 1px solid #e3ece6; border-radius: 6px; background: #fff; padding: 10px 12px; margin-bottom: 8px; }
        .dms-batch.needs-repick { border-color: #f2c9a0; background: #fff8f0; }
        .dms-batch-top { display: flex; align-items: baseline; gap: 10px; flex-wrap: wrap; }
        .dms-batch-no { font-weight: 600; font-size: 12.5px; }
        .dms-batch-path { font-size: 12px; color: #4a5a50; flex: 1 1 200px; word-break: break-word; }
        .dms-batch-count { font-size: 11.5px; color: #6b7a71; }
        .dms-batch-files { margin: 8px 0 0; padding-left: 18px; font-size: 12px; color: #4a5a50; }
        .dms-batch-files li { margin-bottom: 3px; }
        .dms-batch-err { display: block; color: #a4262c; font-style: normal; font-size: 11.5px; }
        .dms-batch-warn { margin: 8px 0 0; font-size: 11.5px; color: #8a4b00; line-height: 1.5; }
        .dms-batch-note { margin: 8px 0 0; font-size: 11.5px; color: #5f6f80; line-height: 1.5; }

        .dms-staged { margin-top: 16px; }
        .dms-staged-head { margin: 0 0 8px; font-size: 12.5px; font-weight: 600; }
        .dms-staged-row { border: 1px solid #e1e1e1; border-radius: 6px; margin-bottom: 6px; background: #fff; }
        .dms-staged-row.open { border-color: #0f6c3f; }
        .dms-staged-row.clash { border-color: #d0a05a; background: #fff8f0; }
        .dms-staged-btn { display: flex; width: 100%; gap: 10px; align-items: center; background: none; border: none; font: inherit; text-align: left; padding: 10px 12px; cursor: pointer; }
        .dms-staged-btn .name { flex: 1 1 auto; font-size: 13px; word-break: break-word; }
        .dms-staged-btn .size { font-size: 11.5px; color: #6b7a71; }
        .dms-staged-btn .chev { font-size: 10px; color: #6b7a71; }
        .dms-staged-del { display: inline-block; margin: 0 12px 10px; font-size: 12px; color: #a4262c; background: none; border: none; cursor: pointer; padding: 0; }
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
        /* Stacked, equal-width buttons for the success dialog. Side by side, the
           primary and secondary read as a yes/no pair; stacked they read as two
           destinations, which is what they are. */
        .dms-popup-stack { display:flex; flex-direction: column; align-items:center; gap: 10px; }
        .dms-popup-stack .dms-popup-btn { width: 100%; max-width: 200px; }
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

      {/* No page heading. The web part sits on a page that already titles itself,
          so an <h2> here repeated the title directly above it. The section headings
          below carry the document structure. */}
      <p className="dms-subtitle">
        All fields marked <strong>*</strong> are required.
      </p>

      {/* ── Document Folder Information ──────────────────────────────────── */}
      {/* Deliberately BEFORE Document Details in SOURCE order, not reordered with
          CSS: tab order follows the DOM, so a visual-only swap would have keyboard
          users moving through the form in a different sequence from what they see. */}
      <div className="dms-section">
        <p className="dms-section-title">Document Folder Information</p>

        {deptLoading ? (
          <p className="dms-dept-loading">Loading your access&hellip;</p>
        ) : !privileged && validPaths.length === 0 ? (
          <div className="dms-dept-error">
            {awaitingFolders ? (
              /* The user's GROUPS are correct — the folder side is not. Two different
                 causes with one fix: either the folder does not exist, or it exists and
                 this user's group was never granted access to it (the far more common
                 case, since creating a group grants nothing until reconciliation runs).
                 Reconciliation leads because it is what fixes both; the abbreviation is
                 named second because a unit lacking one is skipped by every run, so
                 reconciliation alone would silently change nothing for it.
                 The membership message below must not be shown here — it would send the
                 administrator to check groups that are already right. */
              <>
                Your unit isn&apos;t ready to receive uploads yet. Your DMS
                administrator needs to run folder reconciliation — and if the unit
                has no folder at all, give it an abbreviation in the DMS Term
                Abbreviation list first.
              </>
            ) : (
              <>
                Your account isn&apos;t fully provisioned to upload — you need
                membership at every level plus the unit uploader role. Contact
                your administrator.
              </>
            )}
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

        {/* Segment picker. Shown whenever a segment is offered, even if there is
            only one: the mockup includes it, and hiding it made the form open on
            "Department" with no indication of which segment those departments
            belonged to. A single-option select still answers "where am I". */}
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
                onChange={(e) => switchMode(e.target.value)}
              >
                {offerable.map((m) => (
                  <option key={m.key} value={m.key}>
                    {m.label}
                    {/* Admins see every segment, including half-built ones, so that they
                        CAN test one — which means an unbuilt segment otherwise looks
                        identical to a live one here. Labelled, never disabled: disabling
                        it would remove the only reason for the exemption. Spec §9. */}
                    {privileged &&
                    segmentProvisionState(m.stagingFolder, mapSections) ===
                      "unprovisioned"
                      ? " — not fully set up yet"
                      : ""}
                  </option>
                ))}
              </select>
            </label>
          );
        })()}

        {/* What "not fully set up yet" means, and what to do about it. Admin-only, and shown only
            when we actually KNOW the segment has no folders — `unknown` says nothing. */}
        {privileged &&
          segmentProvisionState(activeMode()?.stagingFolder, mapSections) ===
            "unprovisioned" && (
            <div className="dms-setup-warn">
              <strong>This segment isn&apos;t ready to receive uploads.</strong> No folders
              exist for it yet, so filing here will not work — and{" "}
              <strong>uploaders cannot see it at all</strong> until it is finished. On the
              Folder Administration page: give every term an{" "}
              <strong>abbreviation</strong> (a term without one is skipped, and gets no
              folder), add the groups and their Folder Access rows, then run{" "}
              <strong>Folder Reconciliation</strong>. You are seeing it because you are an
              administrator.
            </div>
          )}

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
          {/* --- File-path fields, in folder order: Segment (above) -> level(s)
                 -> Year -> Document Type. Intermediate levels span the full row;
                 the deepest one shares a row of three with Year and Document Type,
                 which is the set that identifies a single destination folder. --- */}
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

          {/* The below-Unit chain, in path order. Each tier is a third-width select,
              so it flows onto the row the deepest permissioned level started —
              [Unit][Year][Document Type] today, [Unit][Function][Year] then
              [Document Type] once a tier is added. Nothing changes on a site that
              has not configured one. */}
          {/* Only the tiers that apply to the chosen path. A unit with no subunits shows
              no SubUnit dropdown at all, rather than an empty required one it can never
              satisfy — which is what "not every unit has subunits" means on screen. */}
          {tierPlan().tiers.map((t, i) => {
            const parent = tierPlan().parents[i];
            return renderSelect(
              t.label,
              true,
              tierValues[t.column] ?? "",
              (v) => setTierValues((prev) => ({ ...prev, [t.column]: v })),
              tierOptions(t, parent),
              // A cascading tier is dead until the tier above it is chosen — the same rule
              // the permissioned cascade uses, so it greys out rather than offering nothing.
              !(t.termSet ?? "").trim() && !parent,
            );
          })}

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

          {/* Spans the row. Single-line rather than a textarea, matching the
              mockup — 250 characters is a sentence, not a paragraph, and a tall
              box invites people to write one. */}
          <label className="dms-field" style={{ gridColumn: "1 / -1" }}>
            <span>Remark</span>
            <input
              type="text"
              value={remark}
              maxLength={250}
              onChange={(e) => setRemark(e.target.value)}
            />
            <small>Max. 250 characters</small>
          </label>
        </div>
      </div>

      {/* ── Document Details ────────────────────────────────────────────── */}
      <div className="dms-section">
        <p className="dms-section-title">Document Details</p>
        {/* An empty AllowedFileTypes selection is a hard block, not a silent
            fallback — spec 2026-07-30 §3. The message names the column and the
            list because the client is the one who fixes it, in one click. */}
        {settings.allowedFileTypes.kind === "none" ? (
          <div className="dms-dropzone" style={{ opacity: 0.6 }}>
            <span>{NO_TYPES_MESSAGE}</span>
          </div>
        ) : (
          /* Click anywhere to open the picker, or drop a file on it. The card
             leads the section because choosing the document is the first thing
             anyone does, and the fields below describe what was chosen. */
          <div
            className={`dms-dropzone${dragOver ? " over" : ""}${file ? " has-file" : ""}`}
            role="button"
            tabIndex={0}
            onClick={() => fileRef.current?.click()}
            onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") fileRef.current?.click(); }}
            // preventDefault on dragOver is what makes the element a valid drop
            // target; without it the browser navigates to the file instead.
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              acceptFiles(e.dataTransfer?.files);
            }}
          >
            {/* Counts the batch, not one file. It used to name the single chosen document and offer
                "Change Document" — with several staged, that named whichever was added last and
                invited the client to swap it, when what they want is to add another. */}
            {draftFiles.length > 0 ? (
              <>
                <span className="dms-filecard-ready">READY</span>
                <span className="name">
                  {draftFiles.length} document{draftFiles.length === 1 ? "" : "s"} in this batch
                </span>
                <span className="size">{formatFileSize(draftFiles.reduce((n, x) => n + x.file.size, 0))}</span>
                <span className="dms-link dms-filecard-action">Add more documents</span>
              </>
            ) : (
              <>
                {/* Inlined rather than imported: an <img> would need an asset
                    loader and a second network request for a 20-line glyph. */}
                <svg className="dms-dropzone-icon" viewBox="0 0 24 24" aria-hidden="true">
                  <path d="M12 3v10m0 0 4-4m-4 4-4-4" fill="none" stroke="currentColor"
                    strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2" fill="none"
                    stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
                <span>
                  <span className="dms-link">Choose a document</span> or drop it here
                </span>
              </>
            )}
            <input
              ref={fileRef}
              type="file"
              multiple
              accept={settings.allowedFileTypes.types.join(",")}
              style={{ display: "none" }}
              onChange={(e) => acceptFiles(e.target.files)}
            />
          </div>
        )}

        {/* ── Saved batches ───────────────────────────────────────────────────
            Each is one destination folder plus its files. NOTHING here has been uploaded — said
            plainly, because the one belief a user must never form is that saving a batch sent it. */}
        {batches.length > 0 && (
          <div className="dms-batches">
            <p className="dms-batches-head">
              {batches.length} batch{batches.length === 1 ? "" : "es"} ready ·{" "}
              {stagedNow.files} document{stagedNow.files === 1 ? "" : "s"} —{" "}
              <strong>nothing has been uploaded yet</strong>
            </p>
            {batches.map((b, i) => (
              <div key={b.id} className={`dms-batch${b.needsRepick ? " needs-repick" : ""}`}>
                <div className="dms-batch-top">
                  <span className="dms-batch-no">Batch {i + 1}</span>
                  <span className="dms-batch-path">{b.pathLabels.join(" / ")}</span>
                  <span className="dms-batch-count">
                    {b.files.length} file{b.files.length === 1 ? "" : "s"}
                  </span>
                  <button
                    type="button"
                    className="dms-link"
                    disabled={busy}
                    onClick={() => setBatches(batches.filter((x) => x.id !== b.id))}
                  >
                    Remove
                  </button>
                </div>
                {b.needsRepick && (
                  <p className="dms-batch-warn">
                    The folder structure for this segment changed while this page was open. Remove this
                    batch and add it again with the current destination — its files are still listed
                    below, and nothing has been lost.
                  </p>
                )}
                <ul className="dms-batch-files">
                  {b.files.map((sf) => (
                    <li key={sf.id}>
                      <span>{sf.finalName ?? sf.file.name}</span>
                      {sf.error && <em className="dms-batch-err">{sf.error}</em>}
                    </li>
                  ))}
                </ul>
              </div>
            ))}
            {lastRun && lastRun.failed > 0 && (
              <p className="dms-batch-warn">
                Uploaded {lastRun.ok}. The {lastRun.failed} still listed above could not be uploaded —
                fix the reason shown and press Upload again. Nothing is sent twice.
              </p>
            )}
            {crossBatchDupes.length > 0 && (
              <p className="dms-batch-note">
                The same document appears in more than one batch: {crossBatchDupes.join(", ")}. That is
                allowed — it will be filed in each place.
              </p>
            )}
            {(stagedNow.overCount || stagedNow.overBytes) && (
              <p className="dms-batch-warn">
                {stagedNow.files} documents are waiting in this browser and none of them have been sent
                yet. Uploading now is safer than staging more — if this tab closes, they are lost.
              </p>
            )}
          </div>
        )}

        {/* ── The batch being built ───────────────────────────────────────────
            One row per staged file; the fields below edit whichever row is open. */}
        {draftFiles.length > 0 && (
          <div className="dms-staged">
            <p className="dms-staged-head">
              {draftFiles.length} document{draftFiles.length === 1 ? "" : "s"} selected
            </p>
            {draftFiles.map((sf) => {
              const open = sf.id === activeFileId;
              const clash = draftCollisions.indexOf(sf.id) !== -1;
              return (
                <div key={sf.id} className={`dms-staged-row${open ? " open" : ""}${clash ? " clash" : ""}`}>
                  <button type="button" className="dms-staged-btn" onClick={() => selectFile(sf.id)}>
                    <span className="name">{sf.finalName ?? sf.file.name}</span>
                    <span className="size">{formatFileSize(sf.file.size)}</span>
                    <span className="chev">{open ? "▲" : "▼"}</span>
                  </button>
                  {clash && (
                    <p className="dms-batch-warn">
                      Another document in this batch would be saved under this same name. The saved name
                      is built from Project, Vendor, Document Name and Date, so change one of those.
                    </p>
                  )}
                  {/* The editor belongs to the OPEN row. Only one row is open, so this renders once. */}
                  {open && metaEditor}
                  {open && (
                    <button
                      type="button"
                      className="dms-link dms-staged-del"
                      disabled={busy}
                      onClick={() => {
                        const left = draftFiles.filter((x) => x.id !== sf.id);
                        setDraftFiles(left);
                        setActiveFileId(left.length > 0 ? left[left.length - 1].id : "");
                        if (left.length > 0) applyEditor(left[left.length - 1].meta);
                      }}
                    >
                      Delete File
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        )}

      </div>

      {/* ── Actions ───────────────────────────────────────────────────────────
          Cancel CONFIRMS when anything is staged: the work exists only in this browser, so discarding
          it is unrecoverable — and it sits beside the button that sends everything. */}
      <div className="dms-actions">
        <button
          type="button"
          className="dms-btn secondary"
          onClick={() => {
            if (stagedNow.files > 0) {
              const n = stagedNow.files;
                const sure = window.confirm(
                `Discard ${n} document${n === 1 ? "" : "s"}? ${n === 1 ? "It has" : "They have"} not been uploaded, and this cannot be undone.`,
              );
              if (!sure) return;
            }
            setBatches([]);
            setDraftFiles([]);
            setActiveFileId("");
            setLastRun(undefined);
            resetForm();
          }}
          disabled={busy}
        >
          Cancel
        </button>
        {/* Save batch, not "Add batch": it files what is on screen and clears the pickers for the next
            destination. Hidden with nothing staged, so the single-file path is unchanged — pick a file,
            fill it in, press Upload, which saves the batch and sends it. */}
        {draftFiles.length > 0 && (
          <button
            type="button"
            className="dms-btn secondary"
            onClick={() => { saveBatch(); }}
            disabled={busy || deptLoading}
          >
            Save batch &amp; start another
          </button>
        )}
        <button
          type="button"
          className="dms-btn primary"
          onClick={handleUpload}
          disabled={busy || deptLoading}
        >
          {busy
            ? "Uploading…"
            : stagedNow.files > 0
              ? `Upload ${stagedNow.files} document${stagedNow.files === 1 ? "" : "s"}`
              : "Upload"}
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
              Your document is awaiting approval.
              <br />
              Please visit Home page to track progress.
            </p>
            {/* Two ways out, because the two things people do after uploading are
                "that was my last one" and "I have a stack of these". Previously the
                only button navigated away, so a second upload meant loading the form
                again from scratch. */}
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
                  setToast(null);
                  resetForm();
                }}
              >
                Upload More
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
