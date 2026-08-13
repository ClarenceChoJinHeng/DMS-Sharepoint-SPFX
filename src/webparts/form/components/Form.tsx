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
  const [dragOver, setDragOver] = useState<boolean>(false);
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
  const acceptFile = (picked: File | undefined): void => {
    if (!picked) return;
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
      return;
    }
    setStatus("");
    setFile(picked);
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

  const handleUpload = async (): Promise<void> => {
    const missing: string[] = [];
    if (!file) missing.push("File");
    const m = activeMode();
    // Permissioned tiers first, then the below-Unit chain — the order the fields
    // appear on screen, so the "please choose" list reads top to bottom.
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) missing.push(lvl.label);
    });
    // Derived from the chain, not from hardcoded strings. Every tier is required:
    // an optional one left blank would file documents at inconsistent depths inside
    // a single unit, which is the whole reason the tier exists.
    missing.push(...buildOnDemandSegments(tierPlan().tiers, tierSelections()).missing);
    // These strings are shown to the user, so they must match the on-screen
    // field labels — renamed to "Confidential Level" in the relayout.
    if (!documentDate) missing.push("Document Date");
    if (!confidentiality) missing.push("Confidential Level");
    // The three name parts are required so every saved file carries the full
    // [Project] - [Vendor] - [Document Name] - [Date] shape. composeUploadBase
    // drops blank parts, so leaving these optional silently produced a shorter
    // name than the convention promises — and after the fact a shortened name is
    // indistinguishable from a deliberate one. Trimmed: a space is not a value.
    if (!docName.trim()) missing.push("Document Name");
    if (!projectName.trim()) missing.push("Project Name");
    if (!vendor.trim()) missing.push("Vendor/Customer Name");
    if (missing.length > 0) {
      showToast(`Please complete: ${missing.join(", ")}.`, "error");
      return;
    }
    if (!file) return;

    const finalName = buildUploadName(
      file.name,
      composeUploadBase(projectName, vendor, docName, documentDate),
    );
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
        // Names BOTH causes. "Re-run reconciliation" alone was wrong half the time:
        // since folder names come from DMS Term Abbreviation, a unit with no
        // abbreviation row is skipped by every run, so re-running changes nothing.
        `"${leafTerm.label}" has no folder yet. Your DMS administrator needs to give it an abbreviation in the DMS Term Abbreviation list, then run folder reconciliation.`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }

    // Rename-proof: resolve the Unit folder's CURRENT path from its UniqueId, then
    // ensure-create the Year and Document Type subfolders under it (they inherit its ACL).
    const unitFolder = await resolveMappedFolder(
      context.spHttpClient,
      siteUrl,
      mapping.folderUniqueId,
      mapping.folderUrl,
    );
    const unitSru = unitFolder.serverRelativeUrl;
    if (!unitSru) {
      showToast(
        // Two different problems with two different fixes, and only one of them is
        // reconciliation. Saying "no longer exists" on a 403 or a throttle sends an admin to
        // re-provision an intact tree while the uploader stays blocked.
        unitFolder.confirmedMissing
          ? "The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation."
          : `Your folder could not be opened (HTTP ${unitFolder.status}). That is usually a permissions ` +
            `problem rather than a missing folder — ask an administrator to check your access to this unit.`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }
    const plan = tierPlan();
    const tiers = plan.tiers;
    // Never guess at a tier whose options are not known. "Not loaded" and "this unit has
    // no subunits" would produce the same shortened path, and the shorter one lands in a
    // folder that exists and looks correct — the document is simply filed in the wrong
    // place. Wait, or fail, but do not assume.
    if (plan.unresolved.length > 0) {
      showToast(
        `Still checking ${plan.unresolved.join(" and ")} for this unit. ` +
          `If this does not clear, reload the page before uploading.`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }
    // A malformed chain must never route a file to a partial path. It would land
    // in a folder that exists and looks right, one tier shallower than everything
    // else in the unit — silent misfiling, discovered only when someone cannot
    // find the document. Block the upload and name the config row instead.
    // A page open since before a structure change keeps filing into the OLD shape, because the
    // config is read once at mount (gotcha #10). Every such upload re-creates the two-shapes state
    // a migration just cleaned up — and it is invisible: the upload succeeds and lands in a folder
    // that looks perfectly reasonable. Seen live 2026-08-11, a file filed without the Credit_Card
    // level hours after that level went live.
    //
    // So re-read the chain and refuse if it moved. `undefined` means the read failed, which proves
    // nothing and must not block: the cure would be worse than the disease.
    const loadedMode = activeMode();
    if (loadedMode) {
      const fresh = await freshChainFor(loadedMode.key);
      const inUse = loadedMode.chain ?? loadedMode.levels ?? [];
      if (fresh && chainSignature(fresh) !== chainSignature(inUse)) {
        showToast(
          "The folder structure changed while this page was open, so this upload would be filed in " +
            "the wrong place. Please reload the page and upload again — nothing has been saved.",
          "error",
        );
        setStatus("");
        setBusy(false);
        return;
      }
    }

    const chainError = validateChain(activeMode()?.chain ?? activeMode()?.levels ?? []);
    if (chainError) {
      showToast(
        `The folder structure for this segment is not set up correctly: ${chainError.message} ` +
          `Ask an administrator to check the DMS Config mode row.`,
        "error",
      );
      setStatus("");
      setBusy(false);
      return;
    }
    const { segments, missing: missingTiers } = buildOnDemandSegments(tiers, tierSelections());
    if (missingTiers.length > 0) {
      showToast(`${missingTiers.join(" and ")} ${missingTiers.length > 1 ? "are" : "is"} required.`, "error");
      setStatus("");
      setBusy(false);
      return;
    }

    setStatus("Preparing destination folders…");
    // Walk the chain in order. Each folder inherits the Unit's ACL — nothing below
    // Unit breaks inheritance, which the client confirmed on 2026-08-06.
    let parentSru = unitSru;
    let destFolder: Awaited<ReturnType<typeof ensureFolder>> | undefined;
    for (const name of segments) {
      const made = await ensureFolder(context.spHttpClient, siteUrl, parentSru, name);
      if (!made) {
        showToast(`Could not create the "${name}" folder.`, "error");
        setStatus("");
        setBusy(false);
        return;
      }
      parentSru = made.serverRelativeUrl;
      destFolder = made;
    }
    if (!destFolder) {
      showToast("No destination folder could be resolved for this upload.", "error");
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
      // One entry per below-Unit tier, in chain order. How a tier writes is DERIVED,
      // not flagged: a tier with a `tidCol` writes a plain label + GUID text pair
      // exactly as the permissioned levels do, while a tier without one writes a
      // single managed-metadata column as "Label|GUID". That keeps Year and Document
      // Type byte-identical to what this form has always sent, and lets a new tier
      // use either column shape without a migration.
      const tierFormValues: Array<{ FieldName: string; FieldValue: string }> = [];
      // Only APPLICABLE tiers write metadata. A unit with no subunits leaves SubUnit and
      // SubUnitTid empty rather than storing a value from some other unit's list.
      const metaPlan = tierPlan();
      metaPlan.tiers.forEach((t, i) => {
        const id = tierValues[t.column] ?? "";
        const opts = tierOptions(t, metaPlan.parents[i]);
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
      const formValues: Array<{ FieldName: string; FieldValue: string }> = [
        ...tierFormValues,
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
              acceptFile(e.dataTransfer?.files?.[0]);
            }}
          >
            {file ? (
              <>
                <span className="dms-filecard-ready">READY</span>
                <span className="name">{file.name}</span>
                <span className="size">{formatFileSize(file.size)}</span>
                <span className="dms-link dms-filecard-action">Change Document</span>
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
              accept={settings.allowedFileTypes.types.join(",")}
              style={{ display: "none" }}
              onChange={(e) => acceptFile(e.target.files?.[0])}
            />
          </div>
        )}

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
            {file ? (
              <small>
                Saves as:{" "}
                {buildUploadName(
                  file.name,
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
