import * as React from "react";
import { UPLOAD_PAUSE_SETTING, UPLOAD_PAUSE_MESSAGE, uploadsArePaused } from "../../../shared/uploadPause";
import { libraryHasColumns, REF_COLUMNS, SUBMISSION_FILE_COLUMN, APPROVED_BY_COLUMN, KEYWORD_COLUMN } from "../../../shared/optionalColumns";
// Records the upload so My Submissions can still show the file after it is deleted (2026-08-27).
// Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
import { writeSubmissionRecord, markRecordReplaced } from "../../../shared/spSubmissionRecords";
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
  probeFolderApproveAccess,
  resolveFolderByPath,
  FolderMapRow,
} from "../../../shared/dmsFolderMap";
import { shouldAttemptSelfApprove } from "../../../shared/selfApprove";
import { isSystemAdmin } from "../../../shared/spGroups";
import { friendlyUploadError } from "../../../shared/networkErrors";
import { stripBlockedChars, blockedCharsMessage } from "../../../shared/inputSanitize";
// Highly Confidential routing. Every rule lives in the shared module under test, because each has a
// wrong version that looks identical on screen: the file uploads, the toast is green, and the
// mistake surfaces when the wrong person opens the document weeks later.
import {
  RoutingContext,
  effectiveHcLevel,
  isHcLevel,
  refuseReason,
  selectableLevels,
  swapLibrarySegment,
} from "../../../shared/hcRouting";
import {
  AccessVerdict,
  filterProvisionedPaths,
  filterReachablePaths,
  mappedTermGuidSet,
  normalizeTermGuid,
  segmentProvisionState,
} from "../../../shared/segmentReadiness";
import { formatFileSize } from "../../../shared/fileSize";
import { offersLegalPrivilege } from "../../../shared/legalPrivilege";
import { EVENT } from "../../../shared/auditLog";
import {
  LIST_SUFFIX,
  cachedHcLibraries,
  cachedListTitle,
  hcAvailable,
  libApiTitle,
  libraryTitle,
  libraryUrlSegment,
  DOCUMENTS_URL_SEGMENT,
} from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";

import {
  Batch,
  BatchDestination,
  ClashWhere,
  FileMeta,
  StagedFile,
  UploadResult,
  applyUploadResults,
  batchesNeedingRepick,
  canSaveBatch,
  collisionsWithin,
  decideClash,
  duplicateAcrossBatches,
  nextAvailableName,
  nextId,
  stagedTotals,
  summarise,
  uploadableBatches,
} from "../../../shared/uploadBatches";
import { newReference } from "../../../shared/submissionGroups";
import { useHcClearance } from "../../../shared/hcClearance";
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

/**
 * One line in the name-clash dialog.
 *
 * ⚠ REPLACES THE `clashOffers` / `clashNotes` PAIR (2026-08-28). That split was by whether a free
 * NAME could be computed, which is a property of the folder listing rather than of the uploader's
 * situation — so "already filed" rows landed in one list or the other depending on whether a
 * suggestion happened to be available, and the dialog needed two renderings that could disagree.
 * With a pending draft now replaceable, the split that matters is WHERE the clash is, and a missing
 * suggestion is simply an absent Rename button on that row.
 */
/* ⚠ `clashRowLine` WAS DELETED HERE ON 2026-09-04, with the per-row situation line the client asked
   to remove. It mapped each `ClashWhere` to a sentence, and one of those sentences is worth not
   losing sight of:

     hidden -> "A pending upload with this name already exists. It may be someone else's, and
                replacing it would discard their file."

   ⚠ THAT WARNING IS NO LONGER ON SCREEN ANYWHERE. The pre-check listed the folder and saw nothing,
   so a `hidden` clash is almost certainly a COLLEAGUE'S draft that Draft Item Security hides from
   this uploader — and pressing Yes destroys their unreviewed work without ever asking them. The
   generic copy ("There is already an existing file with the same name") is true but does not say
   whose. Raised with the client; recorded here so it is a known trade rather than a silent one.

   The four states themselves are still documented on `ClashWhere`, and the Yes button restates what
   each one does when pressed. */

interface ClashRow {
  fileId: string;
  /** The composed name that was refused. */
  from: string;
  /** A free name to offer. ABSENT when none could be computed — the row still shows, without Rename. */
  to?: string;
  where: ClashWhere;
  /** "Batch 2" — numbered as the cards are, so the dialog and the page agree. */
  batchLabel: string;
  /** "GHO / GCA / EG / 2024 / Agreement" (client, 2026-08-28: show the folder path per row). */
  pathLabel: string;
}
import {
  AllowedFileTypes,
  CONFIG_UNREADABLE_MESSAGE,
  FALLBACK_FILE_TYPES,
  NO_TYPES_MESSAGE,
  readAllowedFileTypesField,
  resolveAllowedFileTypes,
} from "../../../shared/allowedFileTypes";
import { NOTICE_ATTENTION, NOTICE_ATTENTION_CSS } from "../../../shared/noticeStyles";

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
// own labelCol/tidCol in CRS Config Levels JSON — those win over this table
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

/**
 * The most documents ONE batch may hold (client, 2026-08-26).
 *
 * Per BATCH, not per upload: a second batch takes another thirty, which is the whole point of the
 * batch model. Nothing enforced any ceiling here before — the design mockup named one ("up to 100
 * files") and it was never built, so a pick of any size was staged.
 *
 * Deliberately NOT shared with `BulkUpload`'s own `MAX_FILES` (50). That screen has one flat
 * selection writing straight to the approved side, so its limit answers a different question, and the
 * client set the two numbers separately.
 */
const MAX_FILES_PER_BATCH = 20;

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
/**
 * When a batch was staged, for the collapsed card - `23/Aug/2026 11:30`.
 *
 * DD/MMM/YYYY is the agreed client display format (gotcha #1), NOT the mockup's `14/06/2026`: a
 * numeric day/month is read the other way round by half the audience, and this string exists to be
 * glanced at. Display only - nothing parses it back.
 */
const STAGE_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const stagedAtLabel = (ms: number | undefined): string => {
  if (!ms) return "";
  const d = new Date(ms);
  const p = (n: number): string => (n < 10 ? "0" + n : String(n));
  return `${p(d.getDate())}/${STAGE_MONTHS[d.getMonth()]}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
};

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
/** "a, b and c" — a comma-list a person reads, rather than a machine-joined one. */
const listPhrase = (items: string[]): string => {
  const list = items.filter((s) => s.trim().length > 0);
  if (list.length <= 1) return list.join("");
  return `${list.slice(0, -1).join(", ")} and ${list[list.length - 1]}`;
};

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
/**
 * `success` is NOT a generic "that worked" — it renders the full "Upload Successful / awaiting
 * approval" modal, and nothing but a completed upload may claim it (client, 2026-08-15: it fired on
 * Save batch, which sends nothing). `notice` is the quiet slide-in for everything else that went
 * right.
 */
type ToastType = "error" | "success" | "notice";

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
  // Metadata column INTERNAL names. Portable: a new site sets these in CRS Config
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
  /**
   * The confidentiality LABEL that routes to the Highly Confidential library pair.
   *
   * Blank here means "not configured", NOT "no HC": `effectiveHcLevel` decides, and the libraries
   * existing is what turns routing on. A site with no HC libraries keeps Highly Confidential as the
   * ordinary metadata label it has always been, which is why this cannot simply default to it.
   */
  hcConfidentialityLevel: string;
  /**
   * `"yes"` turns on auto-approving a Head of Unit's own upload — `shared/selfApprove.ts`. Any
   * other value, including absent, means OFF: this is a NEW behaviour change, not an existing one
   * being relaxed, so it fails CLOSED toward the manual-approve behaviour every site already has —
   * the opposite direction from `uploadsPaused`/`legallyPrivilegedFor`, because the cost of a wrong
   * "on" here is a document going live that nobody looked at.
   */
  autoApproveOwnUpload: string;
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
  // Blank, and always read together with hcAvailable() — see the field's note.
  hcConfidentialityLevel: "",
  // Off by default — see the field's own note on `DmsSettings`.
  autoApproveOwnUpload: "",
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
  /* Free text the uploader types so a document can be found later by a word that appears nowhere
     else on it (client, 2026-09-04). PER FILE, like Document Name and unlike Remark — two documents
     in one batch are exactly the case where different keywords are wanted, and Remark's batch scope
     is a deliberate exception the client asked for, not the pattern. */
  const [keyword, setKeyword] = useState<string>("");
  const [legallyPrivileged, setLegallyPrivileged] = useState<boolean>(false);
  // ── Batches ─────────────────────────────────────────────────────────────────
  // Spec: 2026-08-15-batched-multi-file-upload-design.md. A batch is ONE destination folder plus the
  // files that belong in it; an upload carries several. The metadata fields below are the EDITOR for
  // whichever staged file is open, and `draftFiles` is the batch currently being built.
  //
  // Nothing here reaches SharePoint until Upload, and staged work cannot be persisted because a `File`
  // is not serialisable — hence the beforeunload guard, which is the whole mitigation.
  const [batches, setBatches] = useState<Batch[]>([]);
  /**
   * Which saved batches are expanded. Collapsed by DEFAULT - a saved batch is settled work, and
   * the room belongs to the batch still being filled in. Keyed on batch id, never index: rows are
   * removed mid-run, so an index would re-point at a different batch.
   */
  const [openBatches, setOpenBatches] = useState<Record<string, boolean>>({});
  /**
   * Is a batch form on screen? True at mount, so the page lands on Batch 1 ready to fill in.
   *
   * After a save it goes FALSE and stays false until the uploader presses Add another batch
   * (client, 2026-08-23: "there is no point for Add Another batch if the batch automatically appear
   * for them"). An auto-opened next batch also made the page look like it held unsaved work when it
   * held none.
   */
  const [draftOpen, setDraftOpen] = useState<boolean>(true);
  /**
   * Where the open draft came FROM, when re-opened from a saved batch — so saving puts it back in
   * its own position rather than appending it and silently renumbering the rest. `undefined` means
   * this is a new batch.
   */
  const [editingAt, setEditingAt] = useState<number | undefined>(undefined);
  /**
   * Has a save been ATTEMPTED? Required fields only turn red once it has.
   *
   * Before this the only signals were a toast naming a count and an amber badge on the file row —
   * neither of which marks the box that is actually empty (client, 2026-08-23: "the error is so
   * vague I as a user did not realize"). Gating on an attempt matters: colouring every empty
   * required field on load makes a blank form look broken, and people stop reading red.
   *
   * Not cleared per keystroke — each field's error is DERIVED from its own current value, so it
   * clears itself the moment that field is filled in, and the others stay marked.
   */
  const [showErrors, setShowErrors] = useState<boolean>(false);
  const [draftFiles, setDraftFiles] = useState<StagedFile[]>([]);
  const [activeFileId, setActiveFileId] = useState<string>("");
  const [lastRun, setLastRun] = useState<{ ok: number; failed: number } | undefined>(undefined);
  // Which staged files failed the last save attempt. Held so the ROWS can say so — a list of every
  // missing field of every file belongs on the rows, not in one toast.
  const [incompleteIds, setIncompleteIds] = useState<string[]>([]);
  /**
   * Name clashes from the last run that have a free name to offer, awaiting a decision.
   *
   * Held AFTER the run rather than before it (client, 2026-08-26). A pre-flight dialog would have to
   * hoist every per-batch resolution — folder-map lookup, unit resolve, ensure-create, HC routing —
   * out of the upload loop and hold that plan across a dialog, and `ensureFolder` CREATES folders, so
   * they would be created before the question was answered anyway. Running first costs nothing that
   * was not going to happen: the clashing files are refused, stay staged, and nothing is overwritten.
   */
  const [clashRows, setClashRows] = useState<ClashRow[]>([]);
  const [dragOver, setDragOver] = useState<boolean>(false);
  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  // Site-wide upload pause, set while an administrator reorganises the folders. Starts FALSE, so a
  // slow or failed config read never hides the form — the write-time re-check is the real guard.
  const [paused, setPaused] = useState<boolean>(false);
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

  /* ---------- Highly Confidential clearance -------------------------------
     Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md

     CLEARANCE IS A WRITE PROBE, NEVER GROUP MEMBERSHIP. Reconciliation grants folder ACLs in a pass
     separate from group creation, so a `..._UPL_HC` group can exist for days before it grants
     anything — the gap that produced the upload form's original HTTP 403. Here the symptom would be
     worse than a refused upload: an offered level with nowhere to file it.

     Cached per LEAF TERM, because the answer is per unit — a person may be cleared for one unit and
     not another. `undefined` means "not asked yet, or inconclusive", which `canOfferHc` reads as no. */
  /* THE STATE MOVED INTO `useHcClearance` (2026-08-27) — see the hook below. It was declared here
     with the probe inline, which is how Bulk Upload ended up with neither. */
  /* Flips once `primeNames` has resolved the HC pair. `hcAvailable()` reads a module cache, which
     React cannot watch — so without this the clearance effect below could run before the pair was
     known, return early, and never run again. */
  const [hcReady, setHcReady] = useState(false);
  useEffect(() => {
    let live = true;
    primeNames(context.spHttpClient, siteUrl)
      .then(() => { if (live) setHcReady(hcAvailable()); })
      .catch(() => undefined);
    return () => { live = false; };
  }, [context.spHttpClient, siteUrl]);
  /* Attempts per leaf, not a boolean. A single flag stranded a unit for the life of the page the
     moment ANY step bailed — see the effect below for why that reads to the uploader as "you are not
     cleared". Capped so a throttled site cannot produce a retry loop on the interaction path. */
  /* THE PROBE ITSELF LIVES IN `shared/hcClearance.ts` (2026-08-27). It was inline here and absent
     from Bulk Upload, which is exactly how the two drifted: this screen gained an `hcReady`
     re-render trigger on 2026-08-22 and Bulk did not, so the Highly Confidential level appeared
     intermittently on one and for the wrong reason on the other. One implementation now. */
  const hcWrite = useHcClearance(
    context.spHttpClient, siteUrl, permissionedLeafTerm(), hcReady,
  );

  /**
   * The routing context for the CURRENT selection.
   *
   * Built fresh on each read rather than held in state: `hcAvailable()` is answered by primeNames
   * after mount, so a context captured once would freeze the answer from before the probe returned.
   *
   * ⚠ TAKES THE LEAF EXPLICITLY. Saving a batch RESETS the pickers, so at Upload time
   * `permissionedLeafTerm()` is `""` and `hcWrite[""]` is undefined — which refused every Highly
   * Confidential upload until 2026-08-22. The write path passes the BATCH's own leaf; the default is
   * for the render-time call that decides which levels to offer while editing.
   */
  const hcContext = (leafTerm?: string): RoutingContext => ({
    hcLevel: effectiveHcLevel(settings.hcConfidentialityLevel, hcAvailable()),
    hcAvailable: hcAvailable(),
    canWriteHc: hcWrite[leafTerm ?? permissionedLeafTerm()],
  });

  /** The label behind a confidentiality term id — the routing rules compare labels, not GUIDs. */
  const confidentialityLabel = (termId: string): string =>
    options.confidentiality.find((o) => o.id === termId)?.label ?? "";

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

  /**
   * Read the site-wide upload pause, LIVE.
   *
   * Called twice on purpose: once at mount, so the form can say why it is disabled, and again
   * immediately before writing. The second call is the one that matters — settings are read once in a
   * mount-time effect (gotcha #10), so a tab opened before the pause would otherwise upload straight
   * through it. This is the same shape as the stale-chain guard.
   *
   * Returns `false` on any failure: an unreadable config row must never block every uploader on the
   * site. See `uploadsArePaused`.
   */
  const readUploadPause = async (): Promise<boolean> => {
    try {
      const config = await listName(LIST_SUFFIX.config);
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${config}')/items?$select=SettingValue` +
          `&$filter=ConfigType eq 'setting' and Title eq '${UPLOAD_PAUSE_SETTING}'&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return false;
      const rows = ((await res.json()).value ?? []) as Array<{ SettingValue?: string }>;
      return rows.length > 0 && uploadsArePaused(rows[0].SettingValue);
    } catch {
      return false;
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
        `CRS Config read including AllowedFileTypes failed (HTTP ${res.status}). ` +
          `Retrying without that field. Response: ${body}`,
      );
      res = await fetchSettingRows(SETTINGS_FIELDS_LEGACY);
    }
    if (!res.ok) {
      const body = await res.text();
      throw new Error(
        `CRS Config read failed: HTTP ${res.status}. Response: ${body}`,
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
      hcConfidentialityLevel:
        get("hcConfidentialityLevel") ?? DEFAULT_SETTINGS.hcConfidentialityLevel,
      autoApproveOwnUpload:
        get("autoApproveOwnUpload") ?? DEFAULT_SETTINGS.autoApproveOwnUpload,
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

  /* Site Collection Admin OR a member of the site OWNERS group - see `isSystemAdmin`. This checked
     `IsSiteAdmin` alone until 2026-08-27, so somebody in `CRS Owners` was an administrator on Group
     Management and nowhere else. */
  const loadIsAdmin = async (): Promise<boolean> =>
    isSystemAdmin(context.spHttpClient, siteUrl);

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
              "CRS Config settings read failed — using built-in defaults.",
              err,
            );
            return DEFAULT_SETTINGS;
          }),
          loadModes().catch((err) => {
            console.error(
              "CRS Config mode rows read failed — using built-in modes.",
              err,
            );
            return DEFAULT_MODES;
          }),
          loadGroupMap().catch((err) => {
            console.error(
              "CRS Group Map read failed — no authorised upload paths.",
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
              "CRS Folder Map read failed — cannot tell which folders exist, so every authorised path will be offered.",
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
          "AllowedFileTypes not supplied by CRS Config — running on built-in types:",
          loadedSettings.allowedFileTypes.types.join(", "),
        );
        if (isPrivileged) showToast(CONFIG_UNREADABLE_MESSAGE, "error");
      }

      // Mount-time read, for the banner only. The write-time re-check is what actually refuses an
      // upload — this exists so a paused uploader is told BEFORE filling the form in, rather than
      // after choosing a file and pressing Upload.
      setPaused(await readUploadPause());

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
  /* ⚠ THESE THREE ARE THE FILE NAME. `composeUploadBase` builds the uploaded name out of Project
     Name, Vendor/Customer, Document Name and the date — so a character SharePoint refuses in a file
     name, typed here, produces an upload SharePoint rejects, and the rejection names the FILE rather
     than the field. Blocked at the keystroke instead (client QA, 2026-08-30).

     STRIPPED, and the removal is ANNOUNCED. Silently dropping a character somebody typed is how a
     field comes to feel broken; `blockedChar` puts the offending character under the box while they
     still remember typing it. The note is cleared by the next clean keystroke. */
  const [blockedChar, setBlockedChar] = useState<Record<string, string | undefined>>({});
  const guard = (field: string, v: string, set: (x: string) => void): void => {
    const clean = stripBlockedChars(v);
    if (clean !== v) {
      setBlockedChar((prev) => ({ ...prev, [field]: blockedCharsMessage(v) }));
    } else if (blockedChar[field] !== undefined) {
      setBlockedChar((prev) => ({ ...prev, [field]: undefined }));
    }
    set(clean);
  };
  const onProjectNameChange = (v: string): void => guard("projectName", v, setProjectName);
  const onVendorChange = (v: string): void => guard("vendor", v, setVendor);
  const onDocumentDateChange = (iso: string): void => setDocumentDate(iso);
  const onDocNameChange = (v: string): void => guard("docName", v, setDocName);

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
    keyword,
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
    /* ⚠ REMARK IS NOT RESTORED HERE, deliberately — it is the one field on the BATCH card
       (client's mockup) while every other field in this editor is per file.

       Client, 2026-08-23: *"the remark disappear the moment I select a file."* Adding a document
       makes it the active file, which called this with an EMPTY meta, which blanked a box sitting in
       a different card entirely — so text typed before choosing a file was silently lost, and the
       most likely order of work is exactly that: fill in the folder card, then add the documents.

       Leaving it alone makes the box behave the way its POSITION promises: type it once for the
       batch. `saveBatch` then stamps it onto every file, so each document still carries its own
       Remark column value in SharePoint — no schema change, and `documentDetails.ts` keeps showing
       it per file. */
    /* Restored, UNLIKE `remark` above: this one is per file, so the active file's own value is the
       right thing to show. Blank meta blanks it, which is correct for a newly added document. */
    setKeyword(m.keyword ?? "");
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
  /**
   * ⚠ ONLY FILES WHOSE NAME IS ACTUALLY SETTLED ARE JUDGED FOR A CLASH (client, 2026-08-26: *"it
   * automatically force other files to have a rename when I already rename the first file"*).
   *
   * A newly added file inherits Project and Vendor from its sibling but NEVER the typed Document Name
   * — deliberately, since two files sharing a name is the clash this check exists to catch. So the
   * moment four files are picked, three have a blank Document Name, all three compose to the same
   * shorter name, and all three accuse each other of clashing before anyone has typed a character. The
   * complaint reads as "the form is forcing renames on me", and it was right: the real state is "these
   * are not filled in yet", which the incomplete-field marker beside them already says accurately.
   *
   * Filtered on the REQUIRED name parts rather than on `missingForFile`: Confidential Level is required
   * to save but contributes nothing to the filename, so a blank one must not suppress a REAL clash.
   *
   * WARN: PROJECT AND VENDOR ARE DELIBERATELY NOT TESTED, since 2026-08-27 made them optional. Blank is
   * now a legitimate END state, so requiring them would mean a file that correctly omits both is NEVER
   * settled — its row would show the original filename for ever and a genuine clash between two such
   * files would go unseen until Save. Testing only what is required is the honest reading.
   *
   * Residual, accepted: two files sharing a Document Name and Date can flag each other for the moment
   * between filling those in and typing a Project that would distinguish them. It clears itself as they
   * type, and it is advisory — `collisionsWithin` at Save is the real gate.
   *
   * This cannot let a genuine clash reach Save — an incomplete file blocks the save on its own, so by
   * the time every file is complete, every file is in this set.
   */
  const nameSettled = (sf: StagedFile): boolean => {
    const meta = sf.meta;
    return (
      (meta.docName ?? "").trim().length > 0 && (meta.documentDate ?? "").length > 0
    );
  };
  const draftCollisions = collisionsWithin({
    id: "draft", segmentKey: "", chainSignature: "", pathLabels: [], destination: {},
    files: commitEditor(draftFiles).filter(nameSettled),
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

    /* PER-BATCH CAP. 30 on 2026-08-26, lowered to 20 on 2026-08-27 at the client's request. Every
       message below is derived from the constant, so the number lives in exactly one place.
       Until now nothing enforced a ceiling here — the design mockup mentioned one and it was never
       built.

       ⚠ COUNTED AGAINST WHAT IS ALREADY STAGED, not against this selection. Checking `picked.length`
       alone would let six picks of twenty through, which is the way anyone actually reaches thirty.

       TAKES WHAT FITS rather than refusing the whole selection: a pick of forty stages thirty and
       names the number left out. Refusing all forty would lose the twenty-nine that were fine and
       leave the uploader re-picking from a file dialog with no idea which ones mattered. There is no
       silent truncation — the toast states both numbers.

       The cap is on ONE BATCH, not the upload: another batch takes another thirty, which is what the
       batch model is for. Said in the toast so the way forward is obvious. */
    const room = Math.max(0, MAX_FILES_PER_BATCH - acc.length);
    if (room === 0) {
      showToast(
        `This batch already holds the maximum of ${MAX_FILES_PER_BATCH} documents. ` +
          `Save it and add another batch for the rest.`,
        "error",
      );
      if (fileRef.current) fileRef.current.value = "";
      return;
    }
    if (picked.length > room) {
      const left = picked.length - room;
      showToast(
        `A batch holds at most ${MAX_FILES_PER_BATCH} documents — ${room} added, ` +
          `${left} not added. Save this batch and add another for the rest.`,
        "error",
      );
      picked.length = room;
    }
    let lastAdded: StagedFile | undefined;
    for (const f of picked) {
      /**
       * ⚠ NO PREFILL. Every added file starts completely BLANK (client, 2026-08-26: *"the user input
       * is suddenly halfway automatically filled"*).
       *
       * This deliberately drops `inheritDefaults`, which copied the previous file's Project, Vendor,
       * Date and Confidentiality — everything except the Document Name, which had to stay unique. That
       * asymmetry was the whole problem: a new row arrived two-thirds filled with values the uploader
       * had never typed for THAT document, and nothing on screen said where they came from. Offered the
       * choice between labelling the prefill and removing it, the client chose removal.
       *
       * The cost is real and was accepted: a ten-file batch to one vendor now means typing the same
       * Project, Vendor and Date ten times. If that returns as a complaint, the answer is a LABELLED
       * prefill or an explicit "copy from previous" button — NOT silently restoring this, which is the
       * behaviour that was just rejected.
       *
       * `inheritDefaults` stays in `uploadBatches.ts` with its tests, unused, so relabelling or
       * restoring it later is a small change rather than a rewrite.
       */
      const meta: FileMeta = {};
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
  const missingForFile = (sf: StagedFile): string[] => {
    const meta = sf.id === activeFileId ? captureEditor() : sf.meta;
    const out: string[] = [];
    // LISTED IN THE ORDER THEY APPEAR ON SCREEN, so someone reading "please complete: X, Y" can
    // scan down the form and meet them in that sequence. It was Date/Confidentiality first, which
    // matched neither the layout nor the filename and sent people hunting upwards.
    /* PROJECT NAME AND VENDOR/CUSTOMER NAME ARE OPTIONAL (client, 2026-08-27). They were required so
       every saved file carried the full `[Project] - [Vendor] - [Document Name] - [Date]` shape, since
       `composeUploadBase` drops blank parts and a shortened name is afterwards indistinguishable from
       a deliberate one. The client accepted that trade: not every document has a project or a vendor,
       and demanding a placeholder puts junk in a metadata column to satisfy a naming convention.

       WARN: THE COST IS THAT FILENAMES ARE NOW LESS UNIQUE, so clashes get more likely — two documents
       of the same kind, same day, no project or vendor, compose to the SAME name. That is handled
       rather than prevented: `collisionsWithin` blocks the save, and the six clash branches offer a
       ` - Copy` name at upload. Do not re-tighten this without the client asking. */
    if (!(meta.docName ?? "").trim()) out.push("Document Name");
    if (!(meta.documentDate ?? "")) out.push("Document Date");
    if (!(meta.confidentiality ?? "")) out.push("Confidential Level");
    return out;
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
    /* The remark box is batch-scoped (see `applyEditor`), so it is stamped onto EVERY file here
       rather than read from each file's captured meta — which would otherwise hold whatever was in
       the box at the moment that file happened to be switched away from. */
    const files = commitEditor(draftFiles).map((f) => ({
      ...f,
      meta: { ...f.meta, remark },
    }));
    if (files.length === 0) {
      showToast("Add at least one document to this batch.", "error");
      return false;
    }

    // TWO KINDS OF MISSING, reported separately and never enumerated per file.
    //
    // Listing every field of every document produced a 25-item wall of text for five files (client,
    // 2026-08-15: "for a client they won't understand"). It also put the information in the wrong
    // place: which document is short is a property OF THAT ROW, so the rows carry it now — an amber
    // "Details needed" badge, and the first incomplete one opens itself.
    //
    // The destination fields stay listed, because there is exactly one set of them and no row to
    // attach them to.
    const m = activeMode();
    const destMissing: string[] = [];
    (m?.levels ?? []).forEach((lvl, i) => {
      if (!levelValues[i]) destMissing.push(lvl.label);
    });
    destMissing.push(...buildOnDemandSegments(tierPlan().tiers, tierSelections()).missing);

    const short = files.filter((sf) => missingForFile(sf).length > 0);
    if (destMissing.length > 0 || short.length > 0) {
      setIncompleteIds(short.map((sf) => sf.id));
      // Open the first one that needs attention, so "fill these in" has somewhere to start.
      if (short.length > 0 && short.every((sf) => sf.id !== activeFileId)) {
        setDraftFiles(files);
        applyEditor(short[0].meta);
        setActiveFileId(short[0].id);
      }
      const parts: string[] = [];
      if (destMissing.length > 0) parts.push(`choose ${listPhrase(destMissing)}`);
      if (short.length > 0) {
        parts.push(
          short.length === files.length && files.length > 1
            ? `fill in the details for all ${files.length} documents`
            : `fill in the details for ${short.length} document${short.length === 1 ? "" : "s"}`,
        );
      }
      // Mark the actual BOXES from here on. The toast names a count and the row badge names the
      // field, but neither puts anything on the empty control — so the one thing the uploader has
      // to touch looked identical to the fields they had already filled in.
      setShowErrors(true);
      showToast(
        `Before saving this batch, ${parts.join(" and ")}.` +
          (short.length > 0 ? " The fields that need attention are marked in red below." : ""),
        "error",
      );
      return false;
    }
    setIncompleteIds([]);
    setShowErrors(false);

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
          `Ask an administrator to check the CRS Config mode row.`,
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
      createdAt: Date.now(),
      // What the PICKERS need to show this batch again. Separate from `destination`, which is what
      // the upload writes: re-opening a batch must never be able to change where it files.
      editState: {
        uploadMode,
        levelValues: [...levelValues],
        tierValues: { ...tierValues },
        remark,
      },
    };
    if (!canSaveBatch(b)) {
      // The only remaining reason is an internal name clash, which the rows already flag.
      showToast("Two documents in this batch would be saved under the same name.", "error");
      return false;
    }

    // A re-opened batch goes back where it came from. Appending it instead would renumber every
    // batch below it, so "Batch 2" would silently become "Batch 4" as a side effect of editing it.
    if (editingAt === undefined) {
      setBatches([...batches, b]);
    } else {
      const next = batches.slice();
      next.splice(editingAt, 0, b);
      setBatches(next);
    }
    setEditingAt(undefined);
    // The form CLOSES on save. Nothing opens again until Add another batch is pressed — otherwise
    // an empty card sits there looking like unsaved work, and the Add button has no purpose.
    setDraftOpen(false);
    setDraftFiles([]);
    setActiveFileId("");
    setFile(undefined);
    resetForm();
    // `notice`, never `success`: nothing has been sent, and the success modal says "awaiting approval".
    showToast(
      `Batch saved — ${b.files.length} document${b.files.length === 1 ? "" : "s"} ready to upload.`,
      "notice",
    );
    return true;
  };

  /**
   * Rebuild the whole cascade at once, for re-opening a saved batch.
   *
   * Declared HERE rather than beside `initCascade` because it calls `applyRestrictedMode`, which is
   * defined further down — `no-use-before-define` is on, and hoisting the call would only hide the
   * ordering rather than fix it.
   *
   * NOT a loop over `onLevelChange`: that reads `levelValues` and `levelChoices` out of the render
   * closure, so a sequence of awaited calls would each see the state as it stood BEFORE the first
   * one — and the batch would come back with only its first level filled in. Building both arrays
   * locally and setting them once is the only version that is correct for a chain deeper than two.
   *
   * A branch that cannot be read leaves that level's options empty while the VALUE is still
   * restored, so the uploader sees what was chosen and can re-pick. Losing the options is
   * recoverable; silently dropping the value is how a batch comes back filed somewhere else.
   */
  const restoreCascade = async (mode: UploadMode, values: string[]): Promise<void> => {
    if (!privileged) {
      applyRestrictedMode(validPaths, mode, values);
      return;
    }
    const choices: TermOption[][] = [];
    choices[0] = await loadTermSet(mode.termSetGuid).catch(() => [] as TermOption[]);
    for (let i = 0; i + 1 < mode.levels.length; i++) {
      if (!values[i]) break;
      choices[i + 1] = await loadTermChildren(mode.termSetGuid, values[i]).catch(
        () => [] as TermOption[],
      );
    }
    setLevelChoices(choices);
    setLevelValues(values);
  };

  /**
   * Re-open a saved batch for editing (client, 2026-08-23: "I cannot edit the batch after I save").
   *
   * It becomes the draft again — pickers, remark and files restored — and is REMOVED from the saved
   * list while being edited, so it cannot be uploaded half-changed. Saving puts it back at the same
   * index; discarding it loses it, which is why the discard confirms.
   *
   * ⚠ ONE BATCH AT A TIME, and that is the state model rather than a restriction we could lift
   * cheaply. There is a single set of pickers, and the destination is snapshot INTO the batch at
   * save. Editing two at once would mean per-batch picker state — the exact shape that produced the
   * silent HC-clearance failure of 2026-08-22, where a write-time read of the live dropdowns refused
   * every HC upload for weeks and read as a permissions problem.
   */
  const beginEditBatch = (b: Batch, index: number): void => {
    if (draftOpen) {
      showToast(
        "Save or discard the batch you are working on before editing another one.",
        "error",
      );
      return;
    }
    const es = (b.editState ?? {}) as {
      uploadMode?: string;
      levelValues?: string[];
      tierValues?: Record<string, string>;
      remark?: string;
    };
    // A batch saved before this field existed cannot be restored into the pickers. Say so rather
    // than opening a form pre-filled with whatever happened to be on screen — that would re-file it
    // somewhere nobody chose.
    const m = modes.find((x) => x.key === (es.uploadMode ?? b.segmentKey));
    if (!m) {
      showToast(
        "This batch cannot be re-opened — its segment is no longer configured. Remove it and add it again.",
        "error",
      );
      return;
    }
    setUploadMode(m.key);
    setTierValues({ ...(es.tierValues ?? {}) });
    setRemark(es.remark ?? "");
    setDraftFiles(b.files);
    const first = b.files[0];
    if (first) {
      applyEditor(first.meta);
      setActiveFileId(first.id);
    } else {
      setActiveFileId("");
    }
    setBatches(batches.filter((x) => x.id !== b.id));
    setEditingAt(index);
    setDraftOpen(true);
    // Last, and async: the cascade read can take a moment, and everything above is synchronous state
    // the uploader should see immediately.
    restoreCascade(m, [...(es.levelValues ?? [])]).catch(() =>
      showToast(
        "Re-opened, but the folder options could not be re-read — check the destination before saving.",
        "error",
      ),
    );
  };

  /**
   * Upload ONE staged file into an already-resolved folder.
   *
   * This is the old single-file handler's tail with two changes and no others: its values come from
   * `sf.meta` instead of component state, and it RETURNS a result instead of calling `showToast` and
   * returning. How a file is uploaded and tagged is untouched — that code is the most site-exercised
   * in the project, and a UI change is no reason to rewrite it.
   */
  /**
   * Does this library carry the submission-reference columns?
   *
   * The rule and every trap behind it live in `shared/optionalColumns.ts`, shared with Bulk Upload —
   * two copies deciding whether a field is safe to write is how one of them starts stripping metadata
   * the other preserves. Absent ⇒ the upload behaves exactly as it did before, ungrouped.
   */
  const libraryHasRefColumns = (libraryTitle: string): Promise<boolean> =>
    libraryHasColumns(context.spHttpClient, siteUrl, libraryTitle, REF_COLUMNS);

  /**
   * Does this library carry the per-file stamp the submission RECORD joins on?
   *
   * ⚠ A SEPARATE ASK FROM `libraryHasRefColumns`, and deliberately so. `libraryHasColumns` demands
   * every name it is given, so folding this into `REF_COLUMNS` would make a library reconciled before
   * 2026-08-27 fail the check for the PAIR as well — silently switching off the grouping that works
   * there today. Two asks, two cache keys, two independent degradations.
   *
   * Absent ⇒ no stamp and NO RECORD ROW. A row whose stamp never reached the document can never be
   * joined back to it, so it would show on My Submissions as a permanent false "Deleted" — worse than
   * no row at all.
   */
  const libraryHasRecordColumn = (libraryTitle: string): Promise<boolean> =>
    libraryHasColumns(context.spHttpClient, siteUrl, libraryTitle, [SUBMISSION_FILE_COLUMN]);

  const uploadStagedFile = async (
    b: Batch,
    folderId: string,
    sf: StagedFile,
    /**
     * The library the tagging call must address, by TITLE.
     *
     * ⚠ THE FILE IS PLACED BY FOLDER ID AND TAGGED BY LIBRARY TITLE, and those two have to agree.
     * `GetFolderById` reaches into any library, so an HC document landed in `HC Approval Document`
     * correctly — while `validateUpdateListItem` was hardcoded to `settings.stagingLibrary`, the
     * NORMAL approval library. Item ids are per-LIST, so that call looked for the new item's id in the
     * wrong list: 404 on a good day, and on a bad one it finds a DIFFERENT document with that id and
     * writes the metadata onto it, reporting success. HC uploads could never tag, on any site
     * (2026-08-19).
     *
     * Defaulted rather than required so no other caller changes behaviour.
     */
    libraryTitleForTagging: string = settings.stagingLibrary,
    /**
     * The submission this file belongs to, and the batch within it. Blank when the target library has
     * no reference columns — see `libraryHasRefColumns` for why that must degrade silently.
     */
    refs: { submissionId: string; batchId: string; fileId: string } =
      { submissionId: "", batchId: "", fileId: "" },
    /**
     * The names already in the destination folder, fetched LAZILY and only on a clash.
     *
     * A provider rather than an array so the listing costs nothing on the overwhelmingly common path
     * where no name clashes — and is fetched at most once per batch when one does. Undefined (or a
     * failed read) simply means no suggestion is offered, which degrades to the previous refuse-only
     * behaviour rather than blocking the upload.
     */
    listTaken?: () => Promise<string[] | undefined>,
    /**
     * Names this RUN has already offered, treated as taken.
     *
     * WARN: `nextAvailableName` sees only the folder listing, and neither of two clashing files in one
     * run has been written yet - so without this both are offered the SAME free name and the dialog
     * prints it twice (client, 2026-08-27, on Bulk Upload). Nothing is overwritten either way, but the
     * second file is refused a second time.
     */
    reserved: readonly string[] = [],
    /**
     * Files whose APPROVED-SIDE clash the uploader has chosen to proceed past, by staged-file id.
     *
     * Client's revised brief, 2026-08-27: *"if the files is in Document Library approved then allow
     * the client to upload the file to staging to overwrite so that the approval can approve. Never
     * allow anyone to overwrite a pending file in staging."* So a document already filed is a
     * REPLACEABLE RECORD - by going through approval like any other upload - while a pending draft is
     * untouchable.
     *
     * WARN: THIS SETS NO `overwrite` FLAG. The name is free in the APPROVAL library (its only
     * occupant is the filed copy, one library over), so this is an ordinary upload; all it does is
     * stop the approved-side check refusing. The filed copy is replaced later by Auto-route, and only
     * if an approver approves this one.
     */
    replaceApprovedIds: ReadonlySet<string> = new Set<string>(),
    /**
     * Staged-file ids whose clash IN THE APPROVAL LIBRARY the uploader chose to replace
     * (client, 2026-08-28 - *"allow them to replace file on staging, so the decision to replace will
     * be on their [side]"*).
     *
     * WARN: THIS IS THE ONE PATH IN THE CODEBASE THAT SETS `overwrite=true`, reinstated after being
     * deliberately deleted in 1.0.271.0. It destroys a pending draft - possibly somebody else's,
     * which the uploader could not even see - so it is reachable ONLY after the dialog has said so
     * and the uploader has chosen. Version history on both approval libraries is what makes it
     * recoverable; that is a hard prerequisite, verified 2026-08-28 (500 major versions kept).
     */
    replaceStagingIds: ReadonlySet<string> = new Set<string>(),
  ): Promise<UploadResult> => {
    const meta = sf.meta;
    const finalName = sf.finalName ?? sf.file.name;
    const dest = b.destination as {
      levelSelections: Array<{ column: string; label: string; id: string; labelCol?: string; tidCol?: string }>;
      tierFormValues: Array<{ FieldName: string; FieldValue: string }>;
    };

    /* HC is derived from the library being TAGGED, never from component state: that is the value
       the caller resolved and the one the file was actually placed against, so the two cannot
       disagree about which pair this document belongs to. */
    const isHc = libraryTitleForTagging === libApiTitle("StagingHC");
    /* HOISTED ABOVE BOTH try BLOCKS BELOW, 2026-08-24, so it can be called a SECOND time — right
       before the self-approve MERGE further down, not only once here at upload. Self-approve fires
       seconds after this first check, in the same request, but seconds are still a window: a second
       uploader or a bulk import can file the same name into the approved side in between. Re-running
       the identical check there closes the same race `documentsFileClash` (ApprovalDocument.tsx)
       closes for a HUMAN approval — this is that fix's counterpart for the SELF-approve path, which
       never goes through that component at all. Still fails OPEN here, deliberately: unlike the
       approval-time guard, a wrong "no clash" just leaves an upload where it already was — Pending,
       waiting for a human — never an overwrite, because self-approve simply does not fire.

       Originally declared inside the first `try` block below and went out of scope before the
       second one could see it — caught by `tsc`, not silently. */
    /**
     * Is this name already filed on the APPROVED side, and what else is in that folder?
     *
     * Returns the folder's names as well as the verdict (client, 2026-08-26: *"why not auto rename the
     * file for the Documents library version as well?"*). ⚠ A suggestion has to be free in BOTH
     * libraries: the file lands in the approval library NOW and is routed to the approved side LATER,
     * so a name free only here would simply move the collision to Auto-route.
     *
     * Lists the folder rather than probing one name — same single request, and it yields the names the
     * suggestion needs. Still FAILS OPEN (`clash: false`) on any failure: a 403, a throttle or a
     * missing folder must never become a new way to be blocked, since this is an extra guard.
     */
    const approvedClashInfo = async (): Promise<{ clash: boolean; taken?: string[] }> => {
      try {
        const here: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')?$select=ServerRelativeUrl`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!here.ok) return { clash: false };
        const hc = cachedHcLibraries();
        const mirrored = swapLibrarySegment(
          ((await here.json()) as { ServerRelativeUrl?: string }).ServerRelativeUrl,
          isHc ? hc?.approval.urlSegment : libraryUrlSegment(),
          isHc ? hc?.documents.urlSegment : DOCUMENTS_URL_SEGMENT,
        );
        if (!mirrored) return { clash: false };
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Files?$select=Name&$top=5000&@f='${encodeServerRelativePath(mirrored)}'`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!res.ok) return { clash: false };
        const rows = ((await res.json()).value ?? []) as Array<{ Name?: string }>;
        const names = rows.map((r) => r.Name ?? "").filter((n) => n.length > 0);
        const lower = finalName.toLowerCase();
        return { clash: names.some((n) => n.toLowerCase() === lower), taken: names };
      } catch {
        return { clash: false };
      }
    };

    const replaceStaging = replaceStagingIds.has(sf.id);
    const replaceApproved = replaceApprovedIds.has(sf.id);
    /* Read by the `Files/Add` call in the NEXT try block, which is why it is declared out here.
       TRUE only on a path the uploader has explicitly consented to in the dialog. */
    let overwritePending = false;
    /**
     * The `SubmissionFileId` of the document this upload is about to overwrite.
     *
     * ⚠ IT MUST BE READ BEFORE THE WRITE, AND THERE IS NO SECOND CHANCE. `Files/Add(overwrite=true)`
     * replaces the file and this upload then stamps its OWN id over the columns, so the displaced
     * value ceases to exist the moment the Add succeeds. Read it afterwards and there is nothing
     * left to read.
     *
     * Without it the displaced record stops resolving and My Submissions calls it **deleted** —
     * telling the original uploader their document was destroyed, when somebody in fact filed a
     * newer version of it. Client, 2026-08-28: *"The older submission under My Submission will
     * change to Cancelled status if replaced."*
     */
    let displacedFileId: string | undefined;

    try {
      /* ⚠ THIS PROBE CANNOT SEE A COLLEAGUE'S PENDING FILE. Draft Item Security on both approval
         libraries is *"only users who can approve items, and the author"*, so this listing returns
         only the signed-in uploader's own drafts. A clash with anybody else's is invisible here and
         surfaces at the WRITE instead — see the `hidden` branch further down. */
      const existsRes: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')?$select=Exists`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      const inStaging = existsRes.ok;
      /* Honoured BEFORE the approved-side call below, which is fail-open and may throw. A consent the
         uploader already gave must not be silently revoked by an unrelated check failing. */
      overwritePending = inStaging && replaceStaging;

      /* ⚠ THE APPROVED SIDE IS NOW ALWAYS ASKED, and the early return that used to sit here is gone.
         Until 2026-08-27 a staging clash returned immediately, which is what made a name taken in
         BOTH libraries come back rename-only. That was the client's rule, and its stated reason was
         the clause after the comma: *"don't allow them to overwrite, which is basically the don't
         allow them to overwrite the staging file"*. A pending draft is replaceable now, so that
         premise is gone and both facts are actionable — the uploader's choice depends on knowing
         whether a filed copy exists as well.

         Cost: one extra request, and only in the staging-clash case. The happy path already asked.
         `approvedClashInfo` answers `{ clash: false }` on any failure, so an unanswerable check
         degrades to "no clash" rather than becoming a new way to be blocked. */
      const approved = await approvedClashInfo();
      const inApproved = approved.clash;

      /* THE RULE ITSELF LIVES IN `uploadBatches.ts`, pinned by tests. Bulk Upload needs the same
         decision with `allowStagingReplace: false`, and a hand-written second copy is how the
         stricter of the two drifts - the copy nobody exercises being the one that rots. The table
         and the reasoning for `both` are documented on `decideClash`. */
      const decision = decideClash({ inStaging, inApproved, replaceStaging, replaceApproved });
      const where = decision.where;
      const consented = decision.consented;
      /* Identical to the pre-emptive assignment above for this screen; that one exists only so a
         throw inside `approvedClashInfo` cannot silently revoke a consent already given. */
      overwritePending = decision.overwrite;

      /* ── The record about to be displaced ────────────────────────────────────
         Read ONLY when this upload is going to overwrite, so the ordinary path costs nothing.

         ⚠ ITS OWN try/catch, and every failure is silent. The column is absent on a library that
         has not been reconciled since 2026-08-22, which fails the whole `$select` (gotcha #11) —
         and none of that is a reason to hold up an upload the uploader has already consented to.
         No id simply means the displaced record keeps reading `deleted`, i.e. yesterday's
         behaviour, which is always the correct way to degrade. */
      if (overwritePending) {
        try {
          const priorRes: SPHttpClientResponse = await context.spHttpClient.get(
            `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files('${encodeURIComponent(finalName)}')` +
              `/ListItemAllFields?$select=SubmissionFileId`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (priorRes.ok) {
            const prior = await priorRes.json();
            const raw = typeof prior?.SubmissionFileId === "string" ? prior.SubmissionFileId.trim() : "";
            if (raw.length > 0) displacedFileId = raw;
          }
        } catch {
          /* see above — the upload is not the place to report a bookkeeping read */
        }
      }

      if (!consented) {
        /* ONE suggestion routine for all three cases, where there used to be two near-copies.

           ⚠ IT MUST CLEAR BOTH LIBRARIES. A name free only in the approval library is accepted here
           and then collides inside Auto-route's `Copy file`, which renames it with SharePoint's own
           convention and tells nobody. That silent rename later is the whole reason for offering a
           visible one now.

           ⚠ `finalName` AND `reserved` are both included explicitly. The folder listing is CACHED PER
           FOLDER for the whole run, so it predates anything this run has since written; `reserved`
           holds the names already offered to earlier files in the same run. Without the first, a file
           is offered back the name that just failed; without the second, two different files are
           offered the same name. Both have happened. */
        let suggestedName: string | undefined;
        try {
          let taken = [finalName].concat(reserved).concat(approved.taken ?? []);
          try {
            const seen = listTaken ? await listTaken() : undefined;
            if (seen) taken = taken.concat(seen);
          } catch {
            /* One library's names still produce a usable offer. */
          }
          const free = nextAvailableName(finalName, taken);
          // Never offer a "free" name identical to the one that just failed.
          if (free && free.toLowerCase() !== finalName.toLowerCase()) suggestedName = free;
        } catch {
          // No suggestion; the refusal stands on its own and is still actionable.
        }
        const filed = isHc ? "HC Documents" : "Documents";
        return {
          fileId: sf.id,
          ok: false,
          clashWhere: where,
          suggestedName,
          error:
            where === "both"
              ? `"${finalName}" is already waiting for approval in this folder, and a document of ` +
                `that name has already been filed in ${filed}.`
              : where === "staging"
                ? `"${finalName}" is already waiting for approval in this folder.`
                : `A document called "${finalName}" has already been approved and filed in ${filed} ` +
                  `for this folder. Check whether it is the same document before uploading another copy.`,
        };
      }

    } catch {
      // Network error on the existence check — proceed; the upload will surface the real error.
    }

    try {
      const uploadRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=${overwritePending})?$select=ServerRelativeUrl`,
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
        /* ── A NAME CLASH THE CHECKS ABOVE COULD NOT SEE ──────────────────────
           Verified live 2026-08-26, and it is the client's ORIGINAL scenario: person A has a PENDING
           `TEST … .pdf` in this folder; person B uploads the same name. Draft Item Security on the
           approval library is *"Only users who can approve items (and the author)"*, so A's pending
           file is invisible to B — and therefore invisible to BOTH checks above:

             `Files('name')?$select=Exists`  -> security-trimmed 404, reads as "no clash"
             the folder listing               -> does not contain it, so no suggestion
             `Files/Add(overwrite=false)`     -> SharePoint knows, and answers HTTP 500

           So the offer has to be made HERE too, or it can never fire for the very case it was built
           for. The raw text SharePoint returns — *"A later version of this item has already been
           modified…"* — describes versioning to someone who only knows they picked a filename.

           ⚠ MATCHED ON SEVERAL SIGNALS, NOT ONE STATUS. The moderation conflict is a 500, an ordinary
           duplicate can be a 400 or a 409, and the wording differs between them. Matched loosely and
           deliberately: a false POSITIVE only offers a rename the uploader can decline, while a false
           negative puts a versioning error in front of them and hides the fix. */
        const lower = detail.toLowerCase();
        const looksLikeNameClash =
          lower.indexOf("later version") !== -1 ||
          lower.indexOf("already been modified") !== -1 ||
          lower.indexOf("already exists") !== -1 ||
          uploadRes.status === 409;
        if (looksLikeNameClash) {
          /* The visible names PLUS the one we now know is taken. The listing may be empty or
             trimmed — that is exactly how we got here — so `finalName` is added explicitly, which
             guarantees a suggestion even when nothing could be listed at all.

             If that suggestion ALSO collides with another invisible pending file, this same branch
             fires again on the retry and offers the next number. It converges, and each step is
             visible to the uploader rather than silently looping. */
          let taken: string[] = [finalName].concat(reserved);
          try {
            const seen = listTaken ? await listTaken() : undefined;
            if (seen) taken = taken.concat(seen);
          } catch {
            /* Keep the single known-taken name. */
          }
          try {
            // Both libraries, same reason as the two pre-check branches above: a name free only in
            // the approval library moves the collision into Auto-route's `Copy file`.
            taken = taken.concat((await approvedClashInfo()).taken ?? []);
          } catch {
            /* The approval-library names alone still produce a usable offer. */
          }
          const free = nextAvailableName(finalName, taken);
          return {
            fileId: sf.id,
            ok: false,
            error:
              `A document called "${finalName}" already exists in this folder. It may be a pending ` +
              `upload from someone else that you cannot see until it is approved.`,
            suggestedName:
              free && free.toLowerCase() !== finalName.toLowerCase() ? free : undefined,
            /* HIDDEN, not `staging`, even though the file is physically in the same library. The
               pre-check listed the folder and saw nothing, so this is almost certainly a colleague's
               draft that draft security hides. The dialog needs the distinction: replacing here
               discards someone else's unreviewed work, and they are never asked. */
            clashWhere: "hidden",
          };
        }
        const hint =
          uploadRes.status === 404
            ? " The mapped folder may have been deleted — ask an administrator to re-run reconciliation."
            : "";
        return { fileId: sf.id, ok: false, error: `${detail}.${hint}` };
      }
      const uploadJson = await uploadRes.json();
      const uploadedServerRelativeUrl = uploadJson.ServerRelativeUrl;

      /* `UniqueId` is asked for alongside `Id` for the submission record, and the retry below is
         INSURANCE ON THE HOT PATH. `UniqueId` is a built-in on every list item so this should never
         fail — but one unknown name fails the WHOLE `$select` (gotcha #11), and the cost of being
         wrong here is that no document can be tagged at all. The record's `ItemUniqueId` is only ever
         diagnostic (the join is the stamp), so it is never worth an upload for. */
      const itemUrl = (select: string): string =>
        `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=${select}&@f='${encodeServerRelativePath(uploadedServerRelativeUrl)}'`;
      let itemRes: SPHttpClientResponse = await context.spHttpClient.get(
        itemUrl("Id,UniqueId"), SPHttpClient.configurations.v1,
      );
      if (!itemRes.ok) {
        itemRes = await context.spHttpClient.get(itemUrl("Id"), SPHttpClient.configurations.v1);
      }
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
      // THE LABEL, not the dropdown's raw value — which is the TERM ID. HC routing already resolves
      // the label before comparing (`confidentialityLabel`, used by isHcLevel), and this did not, so
      // `legallyPrivilegedFor` had to hold a term GUID to match anything at all. That is why the row
      // on the client's site read `87f8481b-…`, why nothing on screen could explain it, and why a
      // second stray `term_highlyConfidential` GUID row exists that no code reads (2026-08-19).
      //
      // One shape for both rules now: the config row holds LEVEL NAMES, which an administrator can
      // read off the term store and verify by eye.
      const privilegedApplies = offersLegalPrivilege(
        confidentialityLabel(meta.confidentiality ?? ""),
        settings.legallyPrivilegedFor,
      );
      formValues.push({
        FieldName: settings.columns.legallyPrivileged,
        FieldValue: privilegedApplies && (meta.legallyPrivileged ?? "") !== "" ? "true" : "false",
      });

      /* Written ONLY when the caller confirmed the columns exist. Blank means the library has not been
         reconciled since these were introduced, and pushing the field anyway would fail the whole
         call and cost this document every other value it was about to carry. */
      if (refs.submissionId.length > 0 && refs.batchId.length > 0) {
        formValues.push({ FieldName: "SubmissionId", FieldValue: refs.submissionId });
        formValues.push({ FieldName: "BatchId", FieldValue: refs.batchId });
      }
      /* The per-file stamp, checked SEPARATELY from the pair above — see `libraryHasRecordColumn`.
         This is the key the submission record joins on, and it is on the document rather than
         inferred because `UniqueId` does NOT survive Auto-route: the routed copy gets a new one and
         the source carrying the old one is deleted. A stamped column travels with the copy. */
      if (refs.fileId.length > 0) {
        formValues.push({ FieldName: SUBMISSION_FILE_COLUMN, FieldValue: refs.fileId });
      }
      /* ⚠ THE KEYWORD IS ASKED FOR SEPARATELY TOO, and for the reason spelled out on
         `SUBMISSION_FILE_COLUMN`: `libraryHasColumns` demands EVERY name it is given, so folding
         `Keyword` in with either of the checks above would make a library reconciled before today
         fail that check as well — silently switching off a feature that works there now. One ask,
         one cache key, one independent degradation.

         Blank is written as blank rather than skipped: a REPLACED file inherits the columns of the
         document it overwrote, so omitting the field would leave the previous upload's keywords
         attached to entirely new content. Same reasoning as `legallyPrivileged` above. */
      /* Resolved HERE from the library actually being tagged, rather than threaded in as yet another
         parameter: `uploadStagedFile` already takes six, and the answer depends only on the target.
         `libraryHasColumns` caches per library, so a batch of twenty files costs one read.
         A FAILED read answers false, so the field is skipped rather than risking the whole call. */
      const hasKeywordColumn = await libraryHasColumns(
        context.spHttpClient, siteUrl, libraryTitleForTagging, [KEYWORD_COLUMN],
      );
      if (hasKeywordColumn) {
        formValues.push({ FieldName: KEYWORD_COLUMN, FieldValue: (meta.keyword ?? "").trim() });
      }

      const metaRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryTitleForTagging)}')/items(${item.Id})/validateUpdateListItem`,
        SPHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ formValues }),
        },
      );
      if (!metaRes.ok) {
        // NAMES THE STATUS AND THE LIBRARY. "tagging metadata failed" threw both away, and the two
        // causes need opposite fixes: a 404 is the wrong library title (which is exactly how the HC
        // mismatch above hid for so long), a 403 is permissions on that list, a 400 is a malformed
        // payload. An hour went into distinguishing them by hand on 2026-08-19.
        const body = await metaRes.text().catch(() => "");
        console.error("Tagging failed:", metaRes.status, libraryTitleForTagging, body);
        return {
          fileId: sf.id,
          ok: false,
          error:
            `Uploaded, but tagging failed — HTTP ${metaRes.status} on "${libraryTitleForTagging}"` +
            (metaRes.status === 404 ? " (no library with that title)" : ""),
        };
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

      /* RECORD THE UPLOAD, 2026-08-27 (client: *"Client wants to be able to record that file even if
         someone delete or replace"*). Spec:
         docs/superpowers/specs/2026-08-27-submission-record-design.md

         ⚠ ONLY WHEN THE STAMP WENT ON. `refs.fileId` is blank when the library has no
         `SubmissionFileId` column, and a row written without it could never be joined back to its
         document — it would sit on My Submissions as a permanent false "Deleted", which is worse than
         recording nothing. Degrading to the previous behaviour is always correct.

         ⚠ NEVER FAILS THE UPLOAD. The document is uploaded and tagged by this point;
         `writeSubmissionRecord` cannot throw, and its result is deliberately ignored — it logs its
         own failure. A file that is safely filed must not be reported as failed because a record row
         did not write.

         The snapshot is what an uploader filled in, so a DELETED row can still show it. Tier labels
         come from `levelSelections` rather than the internal column names, because this is read by a
         person, possibly years later. */
      if (refs.fileId.length > 0) {
        const snapshot: Record<string, string> = {};
        for (const l of dest.levelSelections ?? []) snapshot[l.column] = l.label;
        for (const [k, v] of [
          ["Document name", meta.documentName ?? ""],
          ["Project name", meta.projectName ?? ""],
          ["Vendor/Customer", meta.vendor ?? ""],
          ["Document date", meta.documentDate ?? ""],
          ["Confidentiality", confidentialityLabel(meta.confidentiality ?? "")],
          ["Remark", meta.remark ?? ""],
          // The derived value, not the checkbox — the same re-derivation the write above uses, so the
          // record cannot claim a legal marker the document does not carry.
          ["Legally privileged", privilegedApplies && (meta.legallyPrivileged ?? "") !== "" ? "Yes" : ""],
        ] as Array<[string, string]>) {
          snapshot[k] = v;
        }
        await writeSubmissionRecord(context.spHttpClient, siteUrl, {
          submissionRef: refs.submissionId,
          batchRef: refs.batchId,
          fileId: refs.fileId,
          uniqueId: typeof item.UniqueId === "string" ? item.UniqueId : undefined,
          fileName: finalName,
          itemPath: uploadedServerRelativeUrl,
          libraryTitle: libraryTitleForTagging,
          uploadedBy: (context.pageContext.user.email ?? "").toLowerCase(),
          uploadedAt: new Date(),
          metadata: snapshot,
          source: "Form",
        });
      }

      /* ── The record this upload displaced ────────────────────────────────────
         Client, 2026-08-28: *"The older submission under My Submission will change to Cancelled
         status if replaced"* — and, when the word was queried, *"I know its weird but client want
         it to be Cancelled"*.

         ⚠ AFTER the new record is written, deliberately. The two rows are independent, but if only
         one can succeed the NEW one matters more: without it this upload has no record at all and
         will itself read as deleted one day. Marking the old row first would spend the run's luck
         on the lesser of the two.

         ⚠ AND IT IS NOT GATED ON THE NEW ROW SUCCEEDING. `writeSubmissionRecord` returns false on
         an unprovisioned list, and the displaced record can perfectly well exist on a site where
         this upload's own row could not be written — refusing to mark it then would leave a
         document reading "deleted" for no reason connected to it.

         `markRecordReplaced` cannot throw; its result is ignored for the same reason as the write
         above — it logs its own failures, and none of them is the uploader's problem. */
      if (displacedFileId) {
        await markRecordReplaced(
          context.spHttpClient,
          siteUrl,
          displacedFileId,
          (context.pageContext.user.email ?? "").toLowerCase(),
        );
      }

      /* AUTO-APPROVE THE UPLOADER'S OWN FILE, 2026-08-24 (client: "if normal HOU or HC HOU upload,
         they dont need to do approval but rather it will automatically approve").
         Never blocks or fails the upload — this is the LAST thing that happens, after tagging has
         already succeeded, and a failed attempt here simply leaves the item Pending exactly as
         before. Fires for BOTH `hou` and `hou_hc` with no persona branch: `folderId` and
         `libraryTitleForTagging` already point at whichever library (normal or HC) this file was
         actually placed in, so the probe answers correctly for either without knowing which.

         `probeFolderApproveAccess` is a REAL ACL READ, not a role lookup — a role table can say
         "this persona approves" while reconciliation has not yet granted it, or a group was renamed.
         Only "granted" self-approves; "denied"/"missing"/"unknown" all leave the item Pending, the
         same as a PIC's upload, which is why this can never fire for a PIC — they hold no Approve
         role on any folder, so the probe always answers "denied" for them.

         The client asked for the existing "your file has been approved" email to keep going to the
         Head of Unit too — so this deliberately does NOT touch Auto-route or suppress anything.
         Once `OData__ModerationStatus` flips here, Auto-route's own poll picks it up exactly as it
         would a manual approval and sends the same email, to the same person, for the same reason. */
      if (shouldAttemptSelfApprove(settings.autoApproveOwnUpload)) {
        try {
          // RE-CHECKED HERE, not trusted from the read at the top of this function. Seconds have
          // passed — the upload and the tagging call — and a same-named file can have landed in the
          // approved side in that window. A clash now means SKIP self-approve, never fail the
          // upload: the file is already correctly uploaded and tagged, so it stays Pending for a
          // human to sort out, exactly the outcome a denied probe already produces below.
          // `.clash` only — the names this now also returns are for building a rename suggestion, which
          // has no meaning here: self-approve either fires or does not, and nothing is being renamed.
          const clashNow = (await approvedClashInfo()).clash;
          if (clashNow) {
            console.warn("Self-approve skipped: a same-named document already exists on the approved side.");
          } else {
            const approveAccess = await probeFolderApproveAccess(context.spHttpClient, siteUrl, folderId);
            if (approveAccess === "granted") {
              const itemUrl =
                `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryTitleForTagging)}')/items(${item.Id})`;
              const mergeHeaders = {
                Accept: "application/json;odata=nometadata",
                "Content-Type": "application/json;odata=nometadata",
                "X-HTTP-Method": "MERGE",
                "IF-MATCH": "*",
              };
              // ⚠ A SEPARATE, EARLIER MERGE — SharePoint REJECTS a MERGE that combines
              // OData__ModerationStatus with any other field (proven live 2026-09-01: HTTP 500
              // "You cannot change moderation status and set other item properties at that same
              // time"). This is a THIRD approval write path (alongside ApprovalDocument.tsx and
              // BulkApprovePanel.tsx) and the approver here IS the uploader, so ApprovedBy is
              // stamped with their own email before the status flip. Best-effort: a failure here
              // must never leave the moderation MERGE below unattempted.
              if (await libraryHasColumns(context.spHttpClient, siteUrl, libraryTitleForTagging, [APPROVED_BY_COLUMN])) {
                try {
                  await context.spHttpClient.post(itemUrl, SPHttpClient.configurations.v1, {
                    headers: mergeHeaders,
                    body: JSON.stringify({ ApprovedBy: (context.pageContext.user.email ?? "").toLowerCase() }),
                  });
                } catch {
                  // Swallowed — a missing stamp leaves ApprovedBy blank, which the audit/notification
                  // flows already treat as "send", never as an error to surface to the uploader.
                }
              }
              const approveRes: SPHttpClientResponse = await context.spHttpClient.post(itemUrl, SPHttpClient.configurations.v1, {
                headers: mergeHeaders,
                body: JSON.stringify({ OData__ModerationStatus: 0 }),
              });
              if (!approveRes.ok) {
                // Left Pending — the upload itself already succeeded and is reported as such below.
                console.warn("Self-approve MERGE failed, item left Pending:", approveRes.status);
              }
            }
          }
        } catch (e) {
          console.warn("Self-approve attempt threw, item left Pending:", e);
        }
      }

      return { fileId: sf.id, ok: true };
    } catch (err) {
      // Raw messages here are worth naming: "Failed to fetch" is a browser extension or content
      // blocker stopping the request before it reaches SharePoint (found live, 2026-08-24, Brave
      // Shields blocking the file upload POST while every other request on the page succeeded) —
      // not a document, permissions or DMS problem, and the raw string tells an uploader none of
      // that. console.error keeps the original for whoever actually has to debug it.
      console.error("Upload failed:", err);
      return { fileId: sf.id, ok: false, error: friendlyUploadError(err, "Upload failed.") };
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
  /**
   * @param override Batches to send INSTEAD of component state.
   *
   * Needed because the clash dialog's Proceed has to re-run with renamed files, and `setBatches` has
   * not landed by the time it calls back — reading state here would send the old names and refuse all
   * over again. Every read of the batch list below goes through `source`, never `batches`.
   */
  const handleUpload = async (
    override?: Batch[],
    /**
     * Staged-file ids whose APPROVED-SIDE clash the uploader chose to proceed past.
     * See `replaceApprovedIds` on `uploadStagedFile`.
     */
    replaceApprovedIds?: ReadonlySet<string>,
    /**
     * Staged-file ids whose APPROVAL-LIBRARY clash the uploader chose to REPLACE.
     * See `replaceStagingIds` on `uploadStagedFile` - this is the one consent that overwrites.
     */
    replaceStagingIds?: ReadonlySet<string>,
  ): Promise<void> => {
    const source = override ?? batches;
    // An unsaved draft is the commonest way to lose work here: files staged, destination chosen, then
    // Upload pressed without Save batch. Save it rather than silently ignoring it. `saveBatch`'s
    // setState has not landed yet, so this run stops and the client presses Upload again — one extra
    // click, versus uploading a batch list that does not include what is on screen.
    if (draftFiles.length > 0) {
      if (!saveBatch()) return;
      showToast("Saved the open batch — press Upload again to send everything.", "notice");
      return;
    }
    if (source.length === 0) {
      showToast("Nothing to upload yet. Add documents and save a batch first.", "error");
      return;
    }

    setBusy(true);
    setLastRun(undefined);
    setStatus("Checking the folder structure…");

    // Re-read the pause immediately before writing, NOT only at mount: a tab left open before an
    // administrator paused uploads would otherwise file straight into a structure being migrated.
    // Nothing is staged or lost — the batches stay exactly as they are and the upload can be retried.
    if (await readUploadPause()) {
      setPaused(true);
      setBusy(false);
      setStatus("");
      showToast(UPLOAD_PAUSE_MESSAGE, "error");
      return;
    }

    // One read per distinct segment. `undefined` for a segment whose chain could not be read.
    const freshBySegment: Record<string, string | undefined> = {};
    const segmentKeys: string[] = [];
    for (const b of source) if (segmentKeys.indexOf(b.segmentKey) === -1) segmentKeys.push(b.segmentKey);
    for (const key of segmentKeys) {
      const fresh = await freshChainFor(key);
      freshBySegment[key] = fresh ? chainSignature(fresh) : undefined;
    }

    const marked = batchesNeedingRepick(source, freshBySegment);
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
    /* Names this run has already OFFERED, fed back in so two clashing files can never be handed the
       same one. See `reserved` on `uploadStagedFile`. */
    const reserved: string[] = [];
    /* ONE reference per press of Upload — this is what the client means by "a submission" (2026-08-22).
       Generated here rather than per batch so everything sent together shares it, including files that
       went to different departments in the same press. */
    const submissionRef = newReference("SUB", new Date());
    let batchNo = 0;
    for (const b of runnable) {
      batchNo++;
      // One per DESTINATION within the submission, which is what an uploader opens to see "where did
      // these go". Regenerated per batch, never derived from `b.id` — that id is session-only and
      // would repeat if the same staged batch were uploaded twice after a partial failure.
      const batchRef = newReference("BAT", new Date());
      /* Blank unless the target library actually carries the columns. Resolved per library because a
         batch can send some files to the normal approval library and some to the HC one, and the two
         are provisioned independently. Cached, so this is one read per library per page. */
      /* ⚠ THE TWO CHECKS ARE INDEPENDENT. A library reconciled before 2026-08-27 carries the pair and
         not the per-file stamp, and must keep grouping exactly as it does today — so a missing stamp
         column costs the RECORD alone, and a missing pair costs the grouping alone. Folding them into
         one ask would let the newer column switch off the older feature.

         The stamp is minted PER FILE, here rather than per batch: it identifies one document, and it
         is the only thing that still connects that document to its record once Auto-route has copied
         it under a new `UniqueId`. */
      const refsFor = async (
        libraryTitle: string,
      ): Promise<{ submissionId: string; batchId: string; fileId: string }> => {
        const [hasPair, hasStamp] = [
          await libraryHasRefColumns(libraryTitle),
          await libraryHasRecordColumn(libraryTitle),
        ];
        return {
          submissionId: hasPair ? submissionRef : "",
          batchId: hasPair ? batchRef : "",
          fileId: hasStamp ? newReference("SFI", new Date()) : "",
        };
      };
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
            error: `"${dest.leafLabel}" has no folder yet. An administrator needs to give it an abbreviation in the CRS Term Abbreviation list, then run folder reconciliation.`,
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

      /* A BATCH IS ONE FOLDER, BUT NOT NECESSARILY ONE LIBRARY.
         Tiers, Year and Document Type belong to the batch; confidentiality is per FILE. So one saved
         batch can legitimately hold a Confidential document and a Highly Confidential one, and those
         go to different libraries. Resolved lazily, and only when a file actually needs it — a batch
         with no HC file must never pay for a folder resolution, and must never fail because the HC
         tree does not exist. */
      // The BATCH's unit, not the pickers — see `hcContext`. The pickers have been reset by now.
      const hcCtx = hcContext(dest.leafTermId);
      /** Resolved at most once per batch, and only if a file in it is Highly Confidential. */
      let hcDest: { id?: string; error?: string } | undefined;

      /**
       * The names already in a destination folder, read at most ONCE per folder per batch and only
       * when a clash actually happens.
       *
       * Keyed by folder id because one batch can write into TWO folders — the normal approval library
       * and the HC one, since confidentiality is per FILE. Sharing one cache across both would offer a
       * name checked against the wrong folder's contents.
       *
       * `undefined` on any failure, which means "no suggestion" rather than "no names": an unreadable
       * folder must never turn a rename offer into a blocked upload.
       */
      const fetchTaken = async (folderId: string): Promise<string[] | undefined> => {
        try {
          const res: SPHttpClientResponse = await context.spHttpClient.get(
            `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files?$select=Name&$top=5000`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (!res.ok) return undefined;
          const rows = ((await res.json()).value ?? []) as Array<{ Name?: string }>;
          return rows.map((r) => r.Name ?? "").filter((n) => n.length > 0);
        } catch {
          return undefined;
        }
      };
      /* ⚠ THE PROMISE IS CACHED, NOT THE RESOLVED ARRAY. Storing the value means an `await` sits
         between the cache check and the cache write, so two callers for one folder can both miss and
         both fetch — `require-atomic-updates` fails the build on exactly that shape, and it is right
         to: the uploads happen to be sequential today, and nothing here guarantees they stay that
         way. Caching the promise has no await between check and assign, and concurrent callers share
         one request. */
      const takenCache: Record<string, Promise<string[] | undefined>> = {};
      const takenFor = (folderId: string) => (): Promise<string[] | undefined> => {
        if (!takenCache[folderId]) takenCache[folderId] = fetchTaken(folderId);
        return takenCache[folderId];
      };
      const resolveHcFolder = async (): Promise<{ id?: string; error?: string }> => {
        const hcUnit = swapLibrarySegment(
          unitSru, libraryUrlSegment(), cachedHcLibraries()?.approval.urlSegment,
        );
        if (!hcUnit) return { error: "The Highly Confidential library could not be resolved on this site." };
        // THE UNIT FOLDER IS RESOLVED, NEVER CREATED. One created by an uploader would inherit the
        // LIBRARY root's permissions instead of carrying the unit's — which is exactly how a Highly
        // Confidential document becomes readable by the people the unit ACL exists to exclude. If it
        // is absent, reconciliation is the fix and the upload refuses until then.
        const unitHere = await resolveFolderByPath(context.spHttpClient, siteUrl, hcUnit);
        if (!unitHere) {
          return {
            error:
              "The Highly Confidential folder for this unit does not exist yet, or you cannot reach it. " +
              "Ask an administrator to run folder reconciliation.",
          };
        }
        // Below the unit, ensure-creation is correct and matches the normal path: everything below
        // Unit inherits the unit's ACL by design.
        let parent = unitHere.serverRelativeUrl;
        let folderId = unitHere.uniqueId;
        for (const name of dest.segments) {
          const made = await ensureFolder(context.spHttpClient, siteUrl, parent, name);
          if (!made) {
            return {
              error:
                `Could not create the "${name}" folder in the Highly Confidential library.`,
            };
          }
          parent = made.serverRelativeUrl;
          folderId = made.uniqueId;
        }
        return { id: folderId };
      };

      let fileNo = 0;
      for (const sf of b.files) {
        fileNo++;
        setStatus(
          `Batch ${batchNo} of ${runnable.length} · file ${fileNo} of ${b.files.length} — ${sf.finalName ?? sf.file.name}`,
        );
        const level = confidentialityLabel(sf.meta.confidentiality ?? "");
        if (isHcLevel(level, hcCtx)) {
          // RE-CHECKED HERE, not trusted from the dropdown. Settings and clearance are read once at
          // mount, so a tab left open across a clearance change would otherwise submit against the
          // old answer — gotcha 10b, where a stale chain filed into the old folder shape and nothing
          // looked wrong. Hiding a control never clears the state behind it either.
          const refused = refuseReason(level, hcCtx);
          if (refused) {
            results.push({ fileId: sf.id, ok: false, error: refused });
            continue;
          }
          if (hcDest === undefined) hcDest = await resolveHcFolder();
          if (!hcDest.id) {
            results.push({
              fileId: sf.id,
              ok: false,
              error: hcDest.error ?? "No Highly Confidential destination could be resolved.",
            });
            continue;
          }
          // The HC approval library BY NAME, resolved at runtime. `hcAvailable()` is already true here
          // — refuseReason and resolveHcFolder both ran — so the pair is resolved and this cannot fall
          // back to the normal library, which is the one outcome that must never happen quietly.
          const hcTitle = libApiTitle("StagingHC");
          const hcRes = await uploadStagedFile(
            b, hcDest.id, sf, hcTitle, await refsFor(hcTitle), takenFor(hcDest.id), reserved,
            replaceApprovedIds, replaceStagingIds,
          );
          if (hcRes.suggestedName) reserved.push(hcRes.suggestedName);
          results.push(hcRes);
          continue;
        }
        const res = await uploadStagedFile(
          b, destFolder.uniqueId, sf, settings.stagingLibrary, await refsFor(settings.stagingLibrary),
          takenFor(destFolder.uniqueId), reserved, replaceApprovedIds, replaceStagingIds,
        );
        if (res.suggestedName) reserved.push(res.suggestedName);
        results.push(res);
      }
    }

    // `marked`, not `runnable` — a batch held back for a re-pick has no results and must survive this
    // fold untouched.
    const remaining = applyUploadResults(marked, results);
    setBatches(remaining);

    /**
     * EXPAND WHAT FAILED (client, 2026-08-26: *"can you automatically make the dropdown show so I can
     * tell which file is wrong? Right now I cannot tell unless I manually click the dropdown."*).
     *
     * Saved batches are collapsed by default because a saved batch is settled work — right up to the
     * moment one of its files is refused, when the card holding the only explanation is shut. The
     * footer says "fix the reason shown" and the reason was not shown: the uploader is told to act on
     * information the page is hiding.
     *
     * Only batches that actually carry a per-file error are opened, and nothing is ever CLOSED here —
     * a batch the uploader expanded themselves stays expanded, and a clean batch stays as it was.
     * Merging into the existing map rather than replacing it is what keeps both of those true.
     */
    const failedOpen: Record<string, boolean> = {};
    for (const b of remaining) {
      if ((b.files ?? []).some((f) => (f.error ?? "").trim().length > 0)) failedOpen[b.id] = true;
    }
    if (Object.keys(failedOpen).length > 0) {
      setOpenBatches((prev) => ({ ...prev, ...failedOpen }));
    }

    /* ── The clash dialog's rows ───────────────────────────────────────
       EVERY file that failed on a name, whether or not a free one could be offered - a dialog that
       lists some of the failures reads as having missed one (client, 2026-08-26: *"that is why I
       dont see the first file"*).

       Only files STILL STAGED: one that somehow succeeded on a later attempt must not be offered a
       rename it no longer needs, which is why this is matched against `remaining` rather than
       trusting the results array alone.

       Driven from the BATCHES, not the results, so the rows come out in card order and each one can
       carry its batch number and folder path (client, 2026-08-28). */
    const stagedIds: Record<string, true> = {};
    for (const b of remaining) for (const f of b.files ?? []) stagedIds[f.id] = true;
    const rows: ClashRow[] = [];
    remaining.forEach((b, i) => {
      for (const f of b.files ?? []) {
        const r = results.filter((x) => x.fileId === f.id)[0];
        if (!r || r.ok || !r.clashWhere || !stagedIds[f.id]) continue;
        rows.push({
          fileId: f.id,
          from: f.finalName ?? f.file.name,
          to: r.suggestedName,
          where: r.clashWhere,
          // Numbered exactly as the cards are, so "Batch 2" in the dialog is the card headed "Batch 2".
          batchLabel: `Batch ${i + 1}`,
          pathLabel: (b.pathLabels ?? []).join(" / "),
        });
      }
    });
    setClashRows(rows);
    /**
     * ⚠ THERE MUST ALWAYS BE ONE FORM ON SCREEN, and a fully successful upload is the case that
     * broke it (client, 2026-08-23: *"after upload it doesnt reset back to one batch"*).
     *
     * SUCCESS REMOVES, FAILURE STAYS — the rule that makes a second press a safe retry — so a run
     * with nothing left over empties `batches` completely, and the page was then just a heading and
     * an "+ Add another batch" button. The client had already asked for this rule twice, for DISCARD
     * and for DELETE; the third route to an empty page was the one nobody walked, because it only
     * happens after everything WORKED.
     *
     * Opening the draft rather than leaving it closed also puts the next upload one field away,
     * which is what an uploader who has just filed one batch is about to do.
     */
    if (remaining.length === 0) {
      setDraftOpen(true);
      setEditingAt(undefined);
    }
    const counts = summarise(results);
    setLastRun({ ok: counts.ok, failed: counts.failed });
    setStatus("");
    setBusy(false);
    if (counts.failed === 0) {
      showToast(`${counts.ok} document${counts.ok === 1 ? "" : "s"} uploaded and pending review.`, "success");
    } else {
      /* ⚠ THE COUNTS TOAST CAME OFF 2026-08-30 at the client's request — it duplicated the message
         under the batch list, which says the same thing and stays on screen instead of fading. The
         toast is still RAISED, because a failure that announces itself nowhere is worse than one
         announced twice; it just no longer recites the arithmetic. */
      showToast("Some documents could not be uploaded — see the reasons below.", "error");
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
    // Marked only AFTER a save attempt (`showErrors`), never while someone is still typing their
    // way down the form — a field that turns red before it has been reached reads as a fault.
    const invalid = required && showErrors && !value;
    return (
      <label
        className={`dms-field${invalid ? " invalid" : ""}`}
        style={wide ? { gridColumn: "1 / -1" } : undefined}
      >
        <span>
          {label} {required && <em className="req">*</em>}
        </span>
        <select
          value={value}
          disabled={disabled}
          title={selectedLabel}
          aria-invalid={invalid || undefined}
          onChange={(e) => onChange(e.target.value)}
        >
          <option value="">{placeholder}</option>
          {opts.map((o) => (
            <option key={o.id} value={o.id} title={o.label}>
              {o.label}
            </option>
          ))}
        </select>
        {invalid && <small className="dms-err">Choose a {label.toLowerCase()}</small>}
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
  const metaEditor = (
    <>
              <div className="dms-grid" style={{ marginTop: 16 }}>
                {/* Free-text Project Name — distinct from the Group-led Projects
                    "Group Project Name" folder level in the card below.

                    ⚠ THE ORDER OF THESE THREE IS THE ORDER OF THE COMPOSED FILENAME
                    (client, 2026-08-23): `[Project] - [Vendor] - [Document Name] - [Date]`.
                    Document Name used to lead, which read as "the name of the file" when it
                    is in fact ONE SEGMENT of it — so someone typing there first was naming
                    the middle of a string whose start they had not filled in yet. Reordering
                    is the whole fix; `composeUploadBase` is untouched. Keep them in step: if
                    the convention ever changes, this layout has to move with it. */}
                {/* OPTIONAL since 2026-08-27 (client's request) — no `*`, no invalid state, and no
                    error line. `composeUploadBase` drops a blank part, so omitting both simply yields
                    the shorter `[Document Name] - [Date]`. See `missingForFile` for the trade. */}
                <label className="dms-field" style={{ gridColumn: "1 / -1" }}>
                  <span>Project Name</span>
                  <input
                    type="text"
                    value={projectName}
                    maxLength={50}
                    onChange={(e) => onProjectNameChange(e.target.value)}
                  />
                  {blockedChar.projectName ? (
                    <small className="dms-err">{blockedChar.projectName}</small>
                  ) : (
                    <small>Optional · max. 50 characters</small>
                  )}
                </label>

                {/* Vendor is free text. It also feeds the auto-composed document name. Optional. */}
                <label className="dms-field">
                  <span>Vendor/Customer Name</span>
                  <input
                    type="text"
                    value={vendor}
                    maxLength={50}
                    onChange={(e) => onVendorChange(e.target.value)}
                  />
                  {blockedChar.vendor ? (
                    <small className="dms-err">{blockedChar.vendor}</small>
                  ) : (
                    <small>Optional · max. 50 characters</small>
                  )}
                </label>

                <label className={`dms-field${showErrors && !docName.trim() ? " invalid" : ""}`}>
                  <span>
                    Document Name <em className="req">*</em>
                  </span>
                  {/* Typeable before a file is picked: the name is only read at upload
                      time, and leaving it enabled avoids a greyed-out first field. */}
                  <input
                    type="text"
                    value={docName}
                    maxLength={50}
                    aria-invalid={(showErrors && !docName.trim()) || undefined}
                    onChange={(e) => onDocNameChange(e.target.value)}
                  />
                  {/* The "Saves as" preview MOVED to the staged file row above (client, 2026-08-23).
                      It is assembled from four fields, so anchoring it to one of them read as a
                      preview of that box; on the row it reads as what that file will be called,
                      which is what it is. */}
                  {showErrors && !docName.trim() ? (
                    <small className="dms-err">Document Name is required</small>
                  ) : blockedChar.docName ? (
                    <small className="dms-err">{blockedChar.docName}</small>
                  ) : (
                    <small>Max. 50 characters</small>
                  )}
                </label>
                {/* Document Type used to sit here. It moved into the folder card below:
                    it is part of the destination path (Unit → Year → Document Type), not a
                    property of the document, and grouping it with Unit and Year is what the
                    client's mockup shows. */}

                {/* Remark lives in the folder card above, per the mockup. */}
              </div>

              {/* Document Date | Confidential Level | Legally Privileged — one line. */}
              <div className="dms-detail-row">
                <label className={`dms-field${showErrors && !documentDate ? " invalid" : ""}`}>
                  <span>
                    Document Date <em className="req">*</em>
                  </span>
                  {/* Back to the native picker (client, 2026-08-15: "lets just go back to default").

                      Three attempts to hold Fluent's DatePicker in line with the select beside it
                      failed intermittently, and the reason is structural: Fluent injects its styles
                      into <head> at runtime, so any rule of ours races its own. The styles prop was
                      the correct fix for that and still did not settle it, which is a strong signal
                      to stop paying for the control.

                      WHAT THIS COSTS, stated so it is not rediscovered: the displayed format follows
                      the BROWSER LOCALE again, so an en-US browser shows mm/dd/yyyy and no markup can
                      change it. The client asked for dd-mm-yyyy; this does not deliver it. What it
                      does deliver is a calendar with no typing, correct alignment, and 53KB off the
                      bundle. If dd-mm-yyyy is required, the route is a hand-built calendar, not
                      another pass at styling somebody else's. */}
                  <input
                    type="date"
                    value={documentDate}
                    max={(() => {
                      const d = new Date();
                      const mm = d.getMonth() + 1;
                      const day = d.getDate();
                      return `${d.getFullYear()}-${mm < 10 ? "0" + mm : mm}-${day < 10 ? "0" + day : day}`;
                    })()}
                    aria-invalid={(showErrors && !documentDate) || undefined}
                    onChange={(e) => onDocumentDateChange(e.target.value)}
                  />
                  {showErrors && !documentDate && (
                    <small className="dms-err">Document Date is required</small>
                  )}
                </label>

                {/* Not renderSelect: the info icon belongs on the LABEL, and the label is
                    a <span> holding a real <label htmlFor> so clicking the icon does not
                    fall through and focus the select. */}
                <div className={`dms-field${showErrors && !confidentiality ? " invalid" : ""}`}>
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
                    {/* HIDDEN, NOT DISABLED. A greyed-out "Highly Confidential" tells an uncleared
                        uploader that the level exists and that some of their unit's documents are
                        filed under it — information they have no need for. The client chose hidden.
                        On a site with no HC libraries nothing is filtered: the level stays the
                        ordinary metadata label it has always been. */}
                    {(() => {
                      const ctx = hcContext();
                      const allowed = selectableLevels(
                        options.confidentiality.map((o) => o.label), ctx,
                      );
                      const keep = new Set(allowed.map((l) => l.trim().toLowerCase()));
                      return options.confidentiality
                        .filter((o) => keep.has((o.label ?? "").trim().toLowerCase()))
                        .map((o) => (
                          <option key={o.id} value={o.id} title={o.label}>
                            {o.label}
                          </option>
                        ));
                    })()}
                  </select>
                  {showErrors && !confidentiality && (
                    <small className="dms-err">Confidential Level is required</small>
                  )}
                </div>

                {/* Offered only for the level named by `legallyPrivilegedFor` in CRS Config.
                    Unset means never offered — see the setting's note. The value is re-derived
                    at upload time rather than trusted from here, because hiding the control
                    does not clear the state behind it.

                    A LIST since 2026-08-19 (client: the tick must show for Highly Confidential too),
                    so `Confidential;Highly Confidential` offers it on both. One value behaves exactly
                    as it always did. */}
                {offersLegalPrivilege(confidentialityLabel(confidentiality), settings.legallyPrivilegedFor) && (
                    <div className="dms-lp-wrap">
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
                      >
                        i
                        <span className="dms-info-panel" role="tooltip">
                          This applies to confidential communications (email, advice,
                          documents, conversations) between client and lawyer that are
                          protected by law from being disclosed in a court of law or
                          during legal proceedings.
                        </span>
                      </em>
                    </div>
                  )}
              </div>

              {/* ── Keyword (client, 2026-09-04) ────────────────────────────────────────────────
                  *"Add a new field call Keyword add it in upload form and bulk upload - free text
                  field, it will be used to search throough the home page."*

                  Full width, on its own row: keywords run to several words, and pairing it with
                  another control would make it look like half of one thought.

                  OPTIONAL by design. It is a findability aid, not a property of the document, and
                  `missingForFile` is deliberately untouched — making it required would block every
                  upload of a document whose uploader has nothing to add.

                  It goes through the SAME `guard` as every other free-text box, so a character
                  SharePoint refuses is refused here rather than at the write. */}
              <label className="dms-field">
                <span>Keyword</span>
                <input
                  type="text"
                  maxLength={255}
                  value={keyword}
                  placeholder="Words to help find this document later"
                  onChange={(e) => guard("keyword", e.target.value, setKeyword)}
                />
                {blockedChar.keyword ? (
                  <small className="dms-err">{blockedChar.keyword}</small>
                ) : (
                  <small>Optional &middot; searchable from the home page</small>
                )}
              </label>
    </>
  );

  return (
    <section className="dms-form">
      <style>{`
        .dms-form { max-width: 960px; margin: 32px auto; padding: 0 24px 48px; font-family: 'Segoe UI', sans-serif; }
        .dms-form h2 { margin: 0 0 4px; font-size: 24px; font-weight: 700; color: #1b1b1b; }
        .dms-subtitle { margin: 0 0 28px; font-size: 14px; color: #666; }
        /* ⚠ THE PAGE TITLE IS RENDERED BY THE WEB PART AGAIN (client, 2026-09-03: "client wants the
           Upload Document title to be inside the Upload Form now"). It was REMOVED for the opposite
           reason — the page carried its own title directly above and this repeated it — so whoever
           deploys this must DELETE that text web part from the page, or the heading appears twice.
           Nothing in code can detect the duplicate: a text web part is not readable from here. */
        .dms-page-title { margin: 0 0 24px; font-size: 28px; font-weight: 700; color: #1b1b1b; }
        /* The same copy as .dms-subtitle above, but it now sits ABOVE THE OPEN BATCH rather than at
           the top of the page, so it needs its own spacing: it follows the Add-batch bar instead of
           a heading, and lands immediately on top of a card.
           (No backticks anywhere in this block — this whole <style> body is a JS template literal,
           and one backtick in a CSS comment ends it. That has broken this file three times.) */
        .dms-batch-intro { margin: 4px 0 10px; font-size: 14px; color: #666; line-height: 1.5; }
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
        /* Client, 2026-09-04: *"the icon arrow for each dropdown is too close to the border, move it
           away more"* — on BOTH upload forms. Chromium draws a native select's arrow inside the
           padding box, so padding-right is what moves it away from the border; there is no
           property that positions the arrow directly without replacing it with a background image,
           which would then need its own dark-mode and disabled states. The text keeps its 10px on
           the left, so only the right side changes. */
        /* NO BACKTICKS ANYWHERE IN THIS BLOCK - one ends the template literal, and the error
           names neither the cause nor the line. That is what happened writing this very rule. */
        .dms-field select { padding-right: 18px; }
        .dms-field small { color: #666; font-size: 12px; font-weight: 400; }
        /* Document Date | Confidential Level | Legally Privileged on ONE line, per
           the client mockup. A flex row rather than grid cells because the tick
           exists for only one confidentiality level: when it is absent, date and
           level fall back to an even split that lines up with Project Name and
           Vendor above, so the row does not reflow around a control that is not there. */
        /* TOP-aligned, with a fixed label height, so every control on the row starts at the same y.
           Bottom-aligning made each column's position depend on its own total height — and the two
           columns here are not the same shape: Confidential Level's label carries an info icon, and
           Fluent's date field reserves a strip beneath itself for a validation message. Either one
           shifts its box relative to its neighbour, which is why the date appeared to jump. */
        .dms-detail-row { display: flex; gap: 24px; align-items: flex-start; flex-wrap: wrap; }
        .dms-detail-row > .dms-field { flex: 1 1 200px; min-width: 0; }
        .dms-detail-row > .dms-field > span,
        .dms-detail-row > .dms-field > .dms-labelrow { min-height: 20px; display: flex; align-items: center; }
        /* Legally Privileged sits in the same flex row as Document Date and Confidential Level, but
           it has no label above it — so at flex-start it lined up with their LABELS instead of their
           inputs. The 20px top margin is the label row it does not have; the 38px height then centres
           it against the inputs beside it. Both the tick and its info icon live in the wrapper, so
           they move together. */
        /* Required-field error. Red border plus a red line under the box - the toast alone named a
           count, and the row badge named the field, but neither marked the control the uploader has
           to touch. Only applied after a save attempt - see the showErrors state.
           ⚠ NO BACKTICKS in comments in this block: it is a template literal, and a backtick ends
           it. That has broken this file three times today. */
        .dms-field.invalid > input,
        .dms-field.invalid > select { border-color: #a4262c; background: #fdf6f6; }
        .dms-field.invalid > input:focus,
        .dms-field.invalid > select:focus { outline-color: #a4262c; }
        .dms-field.invalid > span,
        .dms-field.invalid > .dms-labelrow label { color: #a4262c; }
        .dms-err { color: #a4262c; font-weight: 600; }
        .dms-lp-wrap { flex: 0 0 auto; display: flex; align-items: center; gap: 4px; height: 38px; margin-top: 20px; margin-bottom: 16px; }
        .dms-lp { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 600; white-space: nowrap; cursor: pointer; }
        .dms-lp input { accent-color: #0f6c3f; width: 16px; height: 16px; margin: 0; cursor: pointer; }
        /* Info tooltips sit on the LABEL, beside the field name. They used to be
           absolutely positioned against the select's right edge, which had to be
           re-tuned every time the grid changed; on the label there is nothing to
           drift against. */
        /* gap 6px -> 3px, and 8px -> 4px on the Legally Privileged wrapper (client, 2026-08-30:
           the info icon sat too far from what it explains). An icon that far from its label
           reads as belonging to the NEXT control along, which on this row is a different
           field entirely. */
        .dms-labelrow { display: flex; align-items: center; gap: 3px; }
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
        .dms-setup-warn { font-size: 13px; line-height: 1.55; ${NOTICE_ATTENTION_CSS} border-radius: 4px; padding: 10px 14px; margin-bottom: 20px; }
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
        /* ── Batch cards ─────────────────────────────────────────────────────
           A card per destination. The OPEN one is the batch being filled in; the
           collapsed ones above it are already staged. */
        .dms-batchbar { display: flex; margin: 0 0 16px; }
        .dms-btn.ghost { background: #fff; color: #0f6c3f; border-color: #b7d5c3; font-weight: 600; }
        .dms-btn.ghost:hover:enabled { background: #f2f8f4; border-color: #0f6c3f; }
        .dms-btn.ghost:disabled { color: #9aa8a0; border-color: #e6ebe8; cursor: default; }
        .dms-batchcard { border: 1px solid #e1e1e1; border-radius: 8px; background: #fff; margin-bottom: 16px; padding: 16px 18px; }
        .dms-batchcard.needs-repick { border-color: #f2c9a0; background: #fff8f0; }
        .dms-batchcard-head { display: flex; align-items: center; gap: 10px; }
        .dms-batchcard-no { flex: 0 0 auto; width: 20px; height: 20px; border-radius: 50%; background: #0f6c3f;
          color: #fff; font-size: 11.5px; font-weight: 700; display: inline-flex; align-items: center; justify-content: center; }
        .dms-batchcard-title { flex: 1 1 auto; font-size: 14px; font-weight: 600; color: #1b1b1b; }
        .dms-batchcard-icon { flex: 0 0 auto; background: none; border: none; cursor: pointer; padding: 4px 6px;
          color: #5f6f80; line-height: 0; border-radius: 4px; }
        .dms-batchcard-icon:hover:enabled { background: #f2f5f3; color: #1b1b1b; }
        .dms-batchcard-icon.danger:hover:enabled { color: #a4262c; background: #fdf3f4; }
        .dms-batchcard-icon:disabled { opacity: .4; cursor: not-allowed; }
        .dms-batchcard-meta { display: flex; flex-wrap: wrap; gap: 4px 18px; margin-top: 8px; font-size: 12px; color: #6b7a71; }
        .dms-batchcard-path { color: #4a5a50; word-break: break-word; }
        /* Numbered steps inside the open card. Plain dark text, not the green section
           caps: the green belongs to the panel headings INSIDE each step, and two
           levels of green reads as two levels of the same thing. */
        .dms-step { margin: 18px 0 8px; font-size: 13.5px; font-weight: 600; color: #1b1b1b; }
        .dms-step-sub { margin: -4px 0 10px; font-size: 12px; color: #6b7a71; }
        /* Save sits inside the card it acts on, right-aligned, clear of the page footer which
           acts on ALL batches. */
        .dms-cardactions { display: flex; justify-content: flex-end; padding: 4px 0 8px; }
        .dms-btn.secondary:disabled { border-color: #cfd8d3; color: #9aa8a0; cursor: default; }

        .dms-staged { margin-top: 16px; }
        .dms-staged-head { margin: 0 0 8px; font-size: 12.5px; font-weight: 600; }
        /* ⚠ NO overflow-hidden HERE — and note this comment lives inside a template literal, so it
           must never contain a backtick. The per-file editor renders inside this row, and the
           Confidential Level and Legally Privileged info panels are absolutely positioned children
           of it — clipping the row cut the tooltip off at the row's edge, so the definitions the
           icon exists to show were unreadable. The rounded corners come from .dms-staged-head,
           which sets its own radius, so nothing is lost by removing it. */
        .dms-staged-row { border: none; border-radius: 6px; margin-bottom: 6px; background: #fff; }
        /* A CLASH keeps its tint. It is the one state on this row that must be visible without
           opening it — two files heading for the same saved name, which SharePoint would not warn
           about. With the borders gone the background is all that is left to carry it. */
        .dms-staged-row.clash { background: #fff8f0; }
        /* The tinted row is the padding box. Everything inside it — the header, the clash warning and
           the open editor — is inset by the same 12px, so nothing sits flush against the tint. */
        .dms-staged-head { display: flex; align-items: stretch; background: rgba(250, 250, 250, 1); border-radius: 6px; }
        .dms-staged-row.clash .dms-staged-head { background: transparent; }
        .dms-staged-row.short .dms-staged-head { background: #fff8f0; }
        .dms-staged-badge { flex: 0 1 auto; min-width: 0; font-size: 11px; color: #8a4b00; background: #ffeed9; border-radius: 10px; padding: 2px 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .dms-staged-btn { display: flex; flex: 1 1 auto; min-width: 0; gap: 10px; align-items: center; background: none; border: none; font: inherit; text-align: left; padding: 10px 12px; cursor: pointer; }
        .dms-staged-btn .name { flex: 1 1 auto; font-size: 13px; word-break: break-word; }
        .dms-staged-btn .size { font-size: 11.5px; color: #6b7a71; }
        .dms-staged-btn .chev { font-size: 10px; color: #6b7a71; }
        .dms-staged-x { flex: 0 0 auto; background: none; border: none; cursor: pointer; padding: 0 12px; font-size: 13px; line-height: 1; color: #8a8886; }
        .dms-staged-x:hover:enabled { color: #a4262c; }
        .dms-staged-x:disabled { cursor: not-allowed; opacity: .5; }
        /* The composed upload name, under the row it belongs to. Indented to the row's text so it
           reads as a property of that file rather than of the section. */
        .dms-staged-saveas { margin: 0; padding: 0 12px 8px; font-size: 11.5px; color: #6b7a71; word-break: break-word; }
        .dms-staged-body { padding: 4px 12px 12px; }
        .dms-staged-row > .dms-batch-warn { padding: 0 12px 8px; margin-top: 6px; }
        .dms-btn { padding: 9px 24px; border-radius: 4px; cursor: pointer; font: inherit; font-size: 14px; border: 1px solid transparent; }
        .dms-btn.primary { background: #0f6c3f; color: #fff; }
        .dms-btn.primary:disabled { background: #9bbfaa; cursor: default; }
        .dms-btn.secondary { background: #fff; border-color: #0f6c3f; color: #0f6c3f; }
        .dms-status { margin-top: 16px; font-size: 13px; }
        .dms-toast { position: fixed; top: 24px; right: 24px; z-index: 9999; min-width: 300px; max-width: 460px; padding: 14px 40px 14px 16px; border-radius: 6px; font-size: 13px; font-family: 'Segoe UI', sans-serif; box-shadow: 0 4px 16px rgba(0,0,0,.18); animation: dms-slidein .2s ease; }
        .dms-toast.error { background: #d13438; color: #fff; }
        .dms-toast.notice { background: #0f6c3f; color: #fff; }
        /* Deliberately NO .ms-* overrides here. They raced Fluent's runtime-injected styles and lost
           at random; the date field is styled through its own styles prop instead.
           (No backticks in this block — it is a template literal, and one would close it.) */
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
        .dms-popup-btn.danger { background: #fff; color: #a4262c; border: 1px solid #a4262c; }
        .dms-popup-btn.danger:hover { background: #fdf3f4; }
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

      {/* ⚠ THE HEADING IS BACK, REVERSING ITS OWN REMOVAL (client, 2026-09-03). It was taken out
          because the page already titled itself and this repeated it — so THE PAGE'S OWN TITLE WEB
          PART MUST BE DELETED when this deploys, or Upload Document appears twice. That is a
          deployment step, not a code one; nothing here can see a text web part to detect it. */}
      <h1 className="dms-page-title">Upload Document</h1>

      {/* Shown ABOVE the form, not beside the Upload button: an uploader should learn this before
          filling anything in. Amber rather than red — nothing has failed and nothing they did is
          wrong, which is exactly what the message says. The form is left usable on purpose, so a
          paused uploader can still see their destination and prepare; the write-time re-check is
          what refuses the upload itself. */}
      {paused ? (
        <div
          role="status"
          style={{
            margin: "0 0 16px", padding: "10px 14px", borderRadius: 4,
            ...NOTICE_ATTENTION,
            fontSize: 13, lineHeight: 1.5,
          }}
        >
          <strong>Uploads are paused.</strong> {UPLOAD_PAUSE_MESSAGE.replace("Uploads are paused while an administrator reorganises the document folders. ", "")}
        </div>
      ) : undefined}

      {/* Add another batch — ABOVE every card, which is where the client put it. It OPENS an empty
          form and nothing else; saving is a separate act on the card itself. Those were one button
          until 2026-08-23, and combining them is what made a new batch appear unbidden after every
          save.

          Disabled while a batch is open, because there is one set of pickers: a second open form
          would have to share them. The title says which of the two states you are in, since a
          disabled button with no reason reads as a broken page. */}
      <div className="dms-batchbar">
        <button
          type="button"
          className="dms-btn ghost"
          disabled={busy || deptLoading || draftOpen}
          title={
            draftOpen
              ? "Finish the batch below first — save it, or discard it if you do not want it"
              : "Start another batch"
          }
          onClick={() => {
            setEditingAt(undefined);
            setDraftOpen(true);
            resetForm();
          }}
        >
          + Add another batch
        </button>
      </div>

      {/* Saved batches. One card per destination folder. NOTHING here has been uploaded - said
          plainly, because the one belief a user must never form is that saving a batch sent it.

          They render BELOW the Add button and ABOVE the batch being filled in, so numbering runs
          down the page and the open card is always the last one. Collapsed by DEFAULT: settled
          work, and the room belongs to the live one. */}
      {batches.length > 0 && (
        <>
          {/* ⚠ THE COUNT AND THE "nothing has been uploaded yet" WARNING CAME OFF 2026-08-30 at the
              client's request, replaced by the mandatory-fields line — which itself came off this
              spot on 2026-09-02 as a duplicate of `.dms-subtitle` above, per the client's own
              annotation. The staged-work-loss risk is unchanged; the `beforeunload` guard and the
              Cancel confirmation are the only things stating it now. */}
          {batches.map((b, i) => {
            const open = openBatches[b.id] === true;
            return (
              <div
                key={b.id}
                className={`dms-batchcard${b.needsRepick ? " needs-repick" : ""}`}
              >
                <div className="dms-batchcard-head">
                  <span className="dms-batchcard-no">{i + 1}</span>
                  <span className="dms-batchcard-title">Batch {i + 1}</span>
                  {/* Edit. Re-opens this batch as the draft — pickers, remark and files restored —
                      and takes it out of the saved list until it is saved again, so a half-changed
                      batch can never be uploaded. Disabled while another batch is open, because the
                      pickers are shared; the title says so rather than leaving a dead control. */}
                  <button
                    type="button"
                    className="dms-batchcard-icon"
                    disabled={busy || draftOpen}
                    aria-label={`Edit batch ${i + 1}`}
                    title={
                      draftOpen
                        ? "Save or discard the batch you are working on first"
                        : `Edit batch ${i + 1}`
                    }
                    onClick={() => beginEditBatch(b, i)}
                  >
                    <svg width="12" height="12" viewBox="0 0 16 16" aria-hidden="true">
                      <path
                        d="M11.2 2.3l2.5 2.5M2.5 11.5l8.1-8.1 2.5 2.5-8.1 8.1-3.2.7z"
                        fill="none" stroke="currentColor" strokeWidth="1.3"
                        strokeLinecap="round" strokeLinejoin="round"
                      />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className="dms-batchcard-icon"
                    aria-expanded={open}
                    aria-label={`${open ? "Collapse" : "Expand"} batch ${i + 1}`}
                    onClick={() => setOpenBatches((prev) => ({ ...prev, [b.id]: !open }))}
                  >
                    {/* Inline SVG, not a glyph: a unicode chevron renders at a different weight in
                        every font on the estate, and in one of them it is an emoji. */}
                    <svg width="11" height="11" viewBox="0 0 12 12" aria-hidden="true">
                      <path
                        d={open ? "M2 8 L6 4 L10 8" : "M2 4 L6 8 L10 4"}
                        fill="none" stroke="currentColor" strokeWidth="1.8"
                        strokeLinecap="round" strokeLinejoin="round"
                      />
                    </svg>
                  </button>
                  {/* Removing a batch throws away staged work that cannot be recovered - the files
                      are browser File handles, not anything re-readable. Confirmed for that reason,
                      and only when it actually holds something. */}
                  <button
                    type="button"
                    className="dms-batchcard-icon danger"
                    disabled={busy}
                    aria-label={`Remove batch ${i + 1}`}
                    onClick={() => {
                      if (
                        b.files.length > 0 &&
                                !window.confirm(
                          `Remove batch ${i + 1}? Its ${b.files.length} document` +
                            `${b.files.length === 1 ? "" : "s"} will have to be chosen again - ` +
                            `nothing has been uploaded yet.`,
                        )
                      ) {
                        return;
                      }
                      const left = batches.filter((x) => x.id !== b.id);
                      setBatches(left);
                      // ⚠ THERE IS ALWAYS A FORM. Removing the last saved batch with no card open
                      // left the page with nothing on it at all — no batches, nowhere to type, and
                      // an Add button that reads as "add a second one".
                      if (left.length === 0 && !draftOpen) {
                        setEditingAt(undefined);
                        setDraftOpen(true);
                        resetForm();
                      }
                    }}
                  >
                    <svg width="12" height="12" viewBox="0 0 16 16" aria-hidden="true">
                      <path
                        d="M3 4h10M6.5 4V2.5h3V4M4.5 4l.6 9h5.8l.6-9"
                        fill="none" stroke="currentColor" strokeWidth="1.4"
                        strokeLinecap="round" strokeLinejoin="round"
                      />
                    </svg>
                  </button>
                </div>

                {/* The summary line the mockup shows: where it goes, how much, and when it was
                    staged. Readable while collapsed, which is the state it is normally in. */}
                <div className="dms-batchcard-meta">
                  <span className="dms-batchcard-path">{b.pathLabels.join(" / ")}</span>
                  <span>
                    {b.files.length} file{b.files.length === 1 ? "" : "s"}
                  </span>
                  {b.createdAt ? <span>Created on {stagedAtLabel(b.createdAt)}</span> : null}
                </div>

                {/* A repick warning is NEVER hidden behind the collapse - it is the one thing on a
                    saved batch that needs acting on, and a warning nobody scrolls to is not one. */}
                {b.needsRepick && (
                  <p className="dms-batch-warn">
                    The folder structure for this segment changed while this page was open. Remove this
                    batch and add it again with the current destination - its files are still listed
                    here, and nothing has been lost.
                  </p>
                )}

                {open && (
                  <ul className="dms-batch-files">
                    {b.files.map((sf) => (
                      <li key={sf.id}>
                        <span>{sf.finalName ?? sf.file.name}</span>
                        {sf.error && <em className="dms-batch-err">{sf.error}</em>}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            );
          })}
          {lastRun && lastRun.failed > 0 && (
            <p className="dms-batch-warn">
              {/* "Nothing is sent twice" removed 2026-08-30 at the client's request. The GUARANTEE
                  is unchanged and is what makes a second press safe: a successful file leaves the
                  list, so what remains is exactly what still needs doing. Only the sentence is
                  gone. */}
              Uploaded {lastRun.ok}. The {lastRun.failed} still listed above could not be uploaded -
              fix the reason shown and press Upload again.
            </p>
          )}
          {crossBatchDupes.length > 0 && (
            <p className="dms-batch-note">
              The same document appears in more than one batch: {crossBatchDupes.join(", ")}. That is
              allowed - it will be filed in each place.
            </p>
          )}
          {(stagedNow.overCount || stagedNow.overBytes) && (
            <p className="dms-batch-warn">
              {stagedNow.files} documents are waiting in this browser and none of them have been sent
              yet. Uploading now is safer than staging more - if this tab closes, they are lost.
            </p>
          )}
        </>
      )}

      {/* The batch being filled in. ONE card is open at a time, and that is the STATE MODEL, not a
          styling choice: there is a single set of pickers, and the destination is SNAPSHOT into the
          batch at save (see saveBatch). Making every card independently editable would mean
          per-batch picker state - the exact shape that produced the silent HC-clearance failure of
          2026-08-22, where a write-time read of the live dropdowns refused every HC upload.

          Rendered only when `draftOpen`. After a save there is deliberately NO card until Add
          another batch is pressed. */}

      {/* ⚠ THE INSTRUCTIONS FOLLOW THE OPEN CARD, and this is why one line achieves it: there is
          exactly ONE open card at a time (the state model above), and it renders HERE, after every
          saved card. So gating this on the same `draftOpen` puts it above whichever batch is being
          filled in and removes it when none is — the client's rule, 2026-09-03, with nothing to keep
          in sync per batch.

          ⚠ IF THE OPEN CARD EVER MOVES OR MORE THAN ONE CAN BE OPEN, THIS BREAKS SILENTLY — it would
          keep rendering at this position while the card it describes is somewhere else. It belongs
          immediately above the card and must move with it. */}
      {draftOpen && (
        <p className="dms-batch-intro">
          Add folder information, upload documents and provide details for each document.
          <br />
          All fields marked <strong>*</strong> are mandatory.
        </p>
      )}

      {draftOpen && (
      <div className="dms-batchcard open">
        <div className="dms-batchcard-head">
          <span className="dms-batchcard-no">
            {editingAt === undefined ? batches.length + 1 : editingAt + 1}
          </span>
          <span className="dms-batchcard-title">
            Batch {editingAt === undefined ? batches.length + 1 : editingAt + 1}
            {editingAt === undefined ? "" : " (editing)"}
          </span>
          {/* Discard.
              ⚠ THERE IS ALWAYS A FORM ON THE PAGE (client, 2026-08-23: "ensure that there is always
              one upload form"). So this CLEARS the card and only CLOSES it when a saved batch is
              left behind to work from. Closing the last one stranded the uploader on a page with
              nowhere to type and an Add button that read as "add a second".

              An empty card clears without a dialog — there is nothing to lose, and confirming
              nothing trains people to click through the dialog that matters. With files staged it
              confirms, because a File handle cannot be recovered once dropped. */}
          <button
            type="button"
            className="dms-batchcard-icon danger"
            disabled={busy}
            aria-label={batches.length > 0 ? "Discard this batch" : "Clear this batch"}
            title={
              batches.length > 0
                ? "Discard this batch"
                : "Clear this batch — the form stays, there is always one"
            }
            onClick={() => {
              const n = draftFiles.length;
              if (
                n > 0 &&
                !window.confirm(
                  `${batches.length > 0 ? "Discard" : "Clear"} this batch? Its ${n} document` +
                    `${n === 1 ? "" : "s"} will have to be chosen again — nothing has been ` +
                    `uploaded yet.`,
                )
              ) {
                return;
              }
              setDraftFiles([]);
              setActiveFileId("");
              setFile(undefined);
              setEditingAt(undefined);
              // Close only if there is something else on the page. Otherwise the card stays, empty.
              if (batches.length > 0) setDraftOpen(false);
              resetForm();
            }}
          >
            <svg width="12" height="12" viewBox="0 0 16 16" aria-hidden="true">
              <path
                d="M3 4h10M6.5 4V2.5h3V4M4.5 4l.6 9h5.8l.6-9"
                fill="none" stroke="currentColor" strokeWidth="1.4"
                strokeLinecap="round" strokeLinejoin="round"
              />
            </svg>
          </button>
        </div>

      <p className="dms-step">1. Document folder information</p>
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
                Your unit isn&apos;t ready to receive uploads yet. Your
                administrator needs to run folder reconciliation — and if the unit
                has no folder at all, give it an abbreviation in the CRS Term
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
                {/* ⚠ HARDCODED to "Group Led Project", client's explicit instruction
                    (2026-09-02) — the pilot has exactly one Project-category segment today.
                    This label answers the whole CATEGORY radio, not one segment, so it will
                    need revisiting the day a second Project-category segment is onboarded
                    (the generic "Project" wording this replaces was correct for that case and
                    wrong for this one; there is no label that is right for both). */}
                {side === "BusinessSegment" ? "Business Segment" : "Group Led Project"}
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
                {/* Same split as SegmentCreator's "Project name" vs "Segment name" — this
                    field is genuinely a category-scoped picker (unlike the radio above it),
                    so "Project" is correct here even once a second Project-category segment
                    exists. */}
                {side === "Project" ? "Project" : "Segment"} <em className="req">*</em>
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
                    style={{ gridColumn: "1 / -1", padding: "8px 12px", ...NOTICE_ATTENTION, borderRadius: 4, fontSize: 13 }}
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
              onChange={(e) => guard("remark", e.target.value, setRemark)}
            />
            {blockedChar.remark ? (
              <small className="dms-err">{blockedChar.remark}</small>
            ) : (
              <small>Max. 250 characters</small>
            )}
          </label>
        </div>
      </div>

      <p className="dms-step">2. Upload documents</p>
      <div className="dms-section">
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
            // ⚠ Keyed on the STAGED LIST, not `file`. `file` holds only the last document ADDED, so
            // a batch re-opened for editing has files but no `file` — and the zone fell back to its
            // empty stacked layout while saying "1 document in this batch". The count and the
            // layout now read the same thing.
            className={`dms-dropzone${dragOver ? " over" : ""}${draftFiles.length > 0 ? " has-file" : ""}`}
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

      </div>

      <p className="dms-step">3. Document information for each file</p>
      <p className="dms-step-sub">Provide document information for each uploaded file.</p>
      <div className="dms-section">

        {/* ── The batch being built ───────────────────────────────────────────
            One row per staged file; the fields below edit whichever row is open. */}
        {draftFiles.length > 0 && (
          <div className="dms-staged">
            <p className="dms-staged-head">
              {draftFiles.length} document{draftFiles.length === 1 ? "" : "s"} selected
            </p>
            {/* WARN: CAPPED ONLY WHILE EVERY ROW IS COLLAPSED, and that condition is load-bearing.
                Client, 2026-08-27: *"if the file have 30 files it will be too long to scroll"* - true
                once MAX_FILES_PER_BATCH arrived. The cap was lowered to 20 later the same day, which
                shortens the list but does not remove the need for this: 20 collapsed rows still runs
                past the fold on a laptop.

                An OPEN row lifts the cap entirely, because the per-file editor's Confidential Level
                and Legally Privileged info panels are ABSOLUTELY POSITIONED children of the row, and
                a scroll container clips them - the exact bug the comment on `.dms-staged-row` records
                (the definitions the icon exists to show became unreadable), and the exact reason
                Group Management's group list lifts its own 60vh cap while a group is expanded. Third
                time this codebase has met it.

                The header stays OUTSIDE the box so the count does not scroll away, matching the
                abbreviation editor's warning and Save. */}
            <div
              className="dms-staged-scroll"
              style={
                activeFileId
                  ? undefined
                  : { maxHeight: "55vh", overflowY: "auto", paddingRight: 4 }
              }
            >
            {draftFiles.map((sf) => {
              const open = sf.id === activeFileId;
              const clash = draftCollisions.indexOf(sf.id) !== -1;
              // Recomputed live, so the badge clears as the fields are filled rather than lingering
              // until the next save attempt — a marker that outlives its cause is worse than none.
              const short = incompleteIds.indexOf(sf.id) !== -1 ? missingForFile(sf) : [];
              return (
                <div key={sf.id} className={`dms-staged-row${open ? " open" : ""}${clash ? " clash" : ""}${short.length > 0 ? " short" : ""}`}>
                  {/* The remove control is a SIBLING of the toggle, not inside it: a button nested in a
                      button is invalid HTML, and browsers resolve it by dropping one of the two. */}
                  <div className="dms-staged-head">
                    <button type="button" className="dms-staged-btn" onClick={() => selectFile(sf.id)}>
                      {/* ⚠ THE ORIGINAL FILENAME UNTIL THE COMPOSED NAME IS REAL (client, 2026-08-26:
                          *"I select one file and prefill the file name. Then I select more files with
                          different name and suddenly it name the three more files test test test"*).

                          A composed name built from fields that are not filled in yet is a name that
                          will NEVER be uploaded, because the blank Document Name blocks the save.
                          Showing it made the list read as though the form had renamed four different
                          documents to `TEST - TEST - …`, and the uploader could no longer tell which
                          row was which file.

                          Originally this bit hardest because `inheritDefaults` prefilled Project and
                          Vendor, so an untouched row already had two of the four name parts. That
                          prefill was REMOVED later the same day, so a new row is now entirely blank —
                          but this guard still earns its place: a half-typed row is exactly as
                          misleading, and the fix must not depend on a prefill that no longer exists.

                          So the row shows WHICH FILE THIS IS until the name is settled, and what it
                          BECOMES only once that is true. The open row's "Saves as" line already shows
                          the composed preview throughout, which is where that belongs. */}
                      <span className="name">
                        {nameSettled(sf) ? (sf.finalName ?? sf.file.name) : sf.file.name}
                      </span>
                      {short.length > 0 && (
                        // Names the fields on the row itself. The count alone would send someone
                        // opening five documents to find the one thing each is missing.
                        <span className="dms-staged-badge" title={`Still needed: ${listPhrase(short)}`}>
                          Needs {listPhrase(short)}
                        </span>
                      )}
                      <span className="size">{formatFileSize(sf.file.size)}</span>
                      <span className="chev">{open ? "▲" : "▼"}</span>
                    </button>
                    {/* On EVERY row, not just the open one. Each row names its own file, so there is no
                        ambiguity about what this removes — and requiring a file to be opened before it
                        can be dropped is friction on the commonest correction: picked the wrong file. */}
                    <button
                      type="button"
                      className="dms-staged-x"
                      title={`Remove ${sf.file.name} from this batch`}
                      aria-label={`Remove ${sf.file.name} from this batch`}
                      disabled={busy}
                      onClick={() => {
                        const left = draftFiles.filter((x) => x.id !== sf.id);
                        setDraftFiles(left);
                        // Only move the editor if the row being removed was the one open in it.
                        if (open) {
                          setActiveFileId(left.length > 0 ? left[left.length - 1].id : "");
                          if (left.length > 0) applyEditor(left[left.length - 1].meta);
                        }
                      }}
                    >
                      ✕
                    </button>
                  </div>
                  {/* The composed name, directly under the file it applies to (client, 2026-08-23).
                      It used to sit under Document Name, which is only ONE of the four parts that
                      make it — so it read as a preview of that box rather than of the file.

                      Only on the OPEN row: it is built from the live editor values, which describe
                      whichever row is open. Rendering it on every row would show that one name
                      against all of them. */}
                  {open && (
                    <p className="dms-staged-saveas">
                      Saves as:{" "}
                      {buildUploadName(
                        sf.file.name,
                        composeUploadBase(projectName, vendor, docName, documentDate),
                      )}
                    </p>
                  )}
                  {clash && (
                    <p className="dms-batch-warn">
                      Another document in this batch would be saved under this same name. The saved name
                      is built from Project, Vendor, Document Name and Date, so change one of those.
                    </p>
                  )}
                  {/* The editor belongs to the OPEN row. Only one row is open, so this renders once. */}
                  {open && <div className="dms-staged-body">{metaEditor}</div>}
                </div>
              );
            })}
            </div>
          </div>
        )}

      </div>

      {/* Save belongs to the CARD, not the page footer: it acts on this batch, and the footer acts
          on all of them. Nothing can be uploaded until every batch is saved, so this is the step
          that moves the work forward and it should sit where the work is. */}
      <div className="dms-cardactions">
        <button
          type="button"
          className="dms-btn secondary"
          disabled={busy || deptLoading || draftFiles.length === 0}
          title={
            draftFiles.length === 0
              ? "Add at least one document to this batch first"
              : "Save this batch"
          }
          onClick={() => { saveBatch(); }}
        >
          Save batch
        </button>
      </div>

      </div>
      )}

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
            // Back to the landing state, not to an empty page: Cancel clears the work, and the
            // uploader is still on the upload form. Leaving no card at all would make them press
            // Add another batch to start the FIRST one, which is not what that button means.
            setEditingAt(undefined);
            setDraftOpen(true);
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
        {/* EVERY batch must be saved before anything is sent (client, 2026-08-23). An open card
            holding files is unsaved work, and uploading around it would send the saved batches and
            silently leave that one behind — the uploader would believe all of it went.

            ⚠ IT SAYS SO ON CLICK RATHER THAN GOING GREY (client: "when client click on the upload
            all batches it will ask them to fully fill the batches first"). A disabled primary
            button states that something is wrong and not what, and the `title` only appears if you
            happen to hover it — so the one person who needs the message is the one who never sees
            it.

            ⚠ ANY open card blocks once a batch has been saved — not just one holding files. The
            first version tested for files, so a Batch 2 with its folder half chosen uploaded Batch 1
            and left Batch 2 sitting there (client, 2026-08-23: "a client can upload the all batches
            even when the second batch is not save yet"). From the uploader's side an open card IS
            unsaved work, whatever is in it.

            It cannot deadlock against "there is always a form": that guarantee only holds when NO
            batch is saved, and this only bites once one is. Discarding the open card closes it,
            because a saved batch is left to work from. */}
        <button
          type="button"
          className="dms-btn primary"
          onClick={() => {
            if (draftOpen && batches.length > 0) {
              const which = editingAt === undefined ? batches.length + 1 : editingAt + 1;
              showToast(
                `Batch ${which} has not been saved yet. Press Save batch on it — or remove it with ` +
                  `the bin icon — before uploading.`,
                "error",
              );
              return;
            }
            if (batches.length === 0) {
              showToast(
                "Nothing to upload yet. Fill in the batch and press Save batch first.",
                "error",
              );
              return;
            }
            handleUpload().catch(() => undefined);
          }}
          disabled={busy || deptLoading}
        >
          {busy
            ? "Uploading…"
            : stagedNow.files > 0
              ? "Upload all batches"
              : "Upload"}
        </button>
      </div>

      {status && <p className="dms-status">{status}</p>}

      {toast && (toast.type === "error" || toast.type === "notice") && (
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


      {/* ── The rename offer ─────────────────────────────────────────────────────
          Client, 2026-08-26: *"it will show that popup warning saying the existing file exist your
          filename will be TEST-TEST-TEST-2.pdf, then ask them if they want to proceed with the new
          file name or cancel and rename that specific file themselves."*

          ⚠ THIS NEVER REPLACES ANYTHING. The old "Replace Existing File?" prompt was removed on
          2026-08-15 because replacing destroys a record an approver may already have acted on. This
          asks a different question — may we FILE IT UNDER A FREE NAME — so both answers are safe, and
          the existing document is untouched either way.

          Rendered before the success popup so a partly-successful run shows this rather than a green
          tick: files were refused, and that is the outcome that needs a decision. */}
      {clashRows.length > 0 && (
        <div className="dms-popup-overlay" role="dialog" aria-modal="true">
          <div className="dms-popup" style={{ maxWidth: 560 }}>
            {/* The project's warning icon, RED (`#FF4646`, exclamation) — the client's own reference
                graphic (2026-09-03, "Group 70858.svg"), applied verbatim, replacing the amber
                (`#FF952A`) version this popup used before. Bulk Upload's own clash prompt gets the
                identical swap, for the same reason the comment used to give for keeping them amber:
                two prompts about the same situation must look like the same system. */}
            <div className="dms-popup-icon">
              <svg
                className="dms-popup-svg"
                viewBox="0 0 184 184"
                fill="none"
                xmlns="http://www.w3.org/2000/svg"
              >
                <circle opacity="0.3" cx="91.9999" cy="92" r="75.4872" fill="#FF4646" />
                <circle cx="92" cy="92" r="92" fill="#FF4646" fillOpacity="0.2" />
                <circle
                  cx="92.0003"
                  cy="91.9998"
                  r="61.3333"
                  fill="white"
                  stroke="#FF4646"
                  strokeWidth="3"
                />
                <path d="M93 66L93 100" stroke="#FF4646" strokeWidth="10" strokeLinecap="round" />
                <path d="M93 116.804L93 118" stroke="#FF4646" strokeWidth="10" strokeLinecap="round" />
              </svg>
            </div>
            <p className="dms-popup-title" style={{ color: "#a4262c" }}>
              {clashRows.length === 1 ? "Replace Existing File" : "Replace Existing Files"}
            </p>
            {/* ⚠ THE HEADER SUBCOPY IS GONE (client, 2026-08-28). It read *"nothing already filed
                will be changed or replaced"*, which stopped being true the moment replacement was
                allowed. It was also carrying the explanation for the whole dialog, so each ROW now
                states its own situation instead — more useful anyway, and it is where the batch
                number and folder path the client asked for naturally live.

                ONE scroll box, where there used to be two that could each reach 34vh. The buttons sit
                OUTSIDE it, so a long list never pushes them off screen — the rule the abbreviation
                editor already follows. */}
            <div
              style={{
                background: "#faf9f8", border: "1px solid #edebe9", borderRadius: 8,
                padding: "12px 14px", margin: "14px 0", maxHeight: "38vh", overflowY: "auto",
                textAlign: "left",
              }}
            >
              {clashRows.map((r, i) => (
                <div
                  key={r.fileId}
                  style={{
                    fontSize: 13, lineHeight: 1.6, wordBreak: "break-word",
                    paddingTop: i === 0 ? 0 : 10,
                    marginTop: i === 0 ? 0 : 10,
                    borderTop: i === 0 ? "none" : "1px solid #edebe9",
                  }}
                >
                  <div style={{ fontSize: 11, color: "#6b7a71", marginBottom: 3 }}>
                    {r.pathLabel ? `${r.batchLabel} · ${r.pathLabel}` : r.batchLabel}
                  </div>
                  {/* ONE plain filename (client, 2026-09-04: *"Just display which file name. No need
                      to cross off"*). The struck-through original with the suggested ` - Copy` name
                      beneath it went with the rename BUTTON — with only Yes/No there is no second
                      name to show, and a strikethrough implied the file was already gone.
                      `clashRowLine(r.where)` went too (*"Already waiting for approval in this
                      folder."*): the body copy below now states the situation once for the whole
                      dialog instead of repeating it per row. */}
                  <span style={{ color: "#323130" }}>{r.from}</span>
                </div>
              ))}
            </div>
            {/* ⚠ THIS REPLACES A SECOND, NATIVE `window.confirm()` THAT USED TO FIRE ON TOP OF THIS
                DIALOG (client, 2026-09-02: *"there is no need for two popup as the one we made is
                enough"*) — a plain browser confirm right after a deliberate click on this dialog's
                own danger button was the same redundant-popup pattern flagged on the approval page.
                The facts it carried (version history vs the recycle bin, and that a pending clash
                may be someone else's invisible file) are stated here instead, ALWAYS visible before
                either danger button is clicked, so nothing is lost — only the second click is. */}
            {/* ⚠ ONE SENTENCE, AND ONE DECISION (client, 2026-09-04). This replaced two amber
                paragraphs and THREE buttons — rename, replace-the-pending-draft, and
                send-as-a-replacement — with a plain Yes/No confirm. Their words: *"what client
                essentially wants is to have Yes or no button, Yes is basically Send for approval as
                a replacement."*

                ⚠ THE WORDING WAS AMBIGUOUS AND WAS QUESTIONED BEFORE BUILDING, WHICH WAS THE RIGHT
                CALL. The copy says "overwrite this file" while the arrows in the mockup pointed Yes
                at the RENAME button — opposite behaviours, and the wrong one destroys a pending
                document silently. Confirmed as replace.

                ⚠ WHAT IS LOST, DELIBERATELY: the ` - Copy` suggestion is no longer OFFERED. The
                machinery that computes it still runs (Bulk Upload uses it, and `reserved` needs it to
                stop two files in one run being offered the same name), it simply has no button here
                any more. An uploader who does not want to replace presses No and renames the file
                themselves. */}
            <p style={{ fontSize: 13, color: "#605e5c", textAlign: "left", margin: "0 0 14px", lineHeight: 1.55 }}>
              There is already an existing file with the same name. Please confirm if you would like
              to proceed to overwrite {clashRows.length === 1 ? "this file" : "these files"}
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "center", flexWrap: "wrap" }}>
              {/* ⚠ YES REPLACES EACH ROW BY WHICHEVER MECHANISM THAT ROW NEEDS, and the two are not
                  the same thing — which is why both id sets are passed in one call:
                    • `approved` → the file goes to the approver like any other upload and NOTHING is
                      destroyed by pressing this. Auto-route replaces the filed copy later, and only
                      if the approver approves.
                    • everything else (`staging`, `hidden`, `both`) → the ONLY path that sets
                      `overwrite=true`. It destroys a pending draft immediately, and for `hidden` that
                      draft belongs to somebody the uploader cannot see. Version history on both
                      approval libraries is what makes it recoverable and is a hard prerequisite.
                  `both` belongs with the second: replacing the draft means that draft goes through
                  approval, and Auto-route then replaces the filed copy anyway. */}
              <button
                className="dms-popup-btn confirm"
                onClick={() => {
                  const approved = new Set<string>(
                    clashRows.filter((r) => r.where === "approved").map((r) => r.fileId),
                  );
                  const pending = new Set<string>(
                    clashRows.filter((r) => r.where !== "approved").map((r) => r.fileId),
                  );
                  /* Clear any earlier per-file error so a retried file is not shown as failed while
                     it is being re-sent, and send the batches EXPLICITLY — `setBatches` has not
                     landed when `handleUpload` runs, so reading state here would send the old list. */
                  /* ⚠ NOT `new Set([...approved, ...pending])` — spreading a Set needs ES2015 and
                     this tsconfig targets ES5 (the same constraint that makes `Promise.allSettled`
                     unavailable, gotcha #3). Every clash row is being sent, so build it from the
                     rows directly. */
                  const touched = new Set<string>(clashRows.map((r) => r.fileId));
                  const cleared = batches.map((b) => ({
                    ...b,
                    files: (b.files ?? []).map((f) =>
                      touched.has(f.id) ? { ...f, error: undefined } : f,
                    ),
                  }));
                  setClashRows([]);
                  setBatches(cleared);
                  handleUpload(
                    cleared,
                    approved.size > 0 ? approved : undefined,
                    pending.size > 0 ? pending : undefined,
                  ).catch(() => undefined);
                }}
              >
                Yes
              </button>
              <button
                className="dms-popup-btn cancel"
                onClick={() => {
                  /* Nothing to undo — every file is already staged with its refusal showing, which is
                     where an uploader renames it by hand. Just close. */
                  setClashRows([]);
                }}
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
            {/* Counts, because a batch can be five. "Your document is awaiting approval" after
                uploading five reads as though four went missing. */}
            <p className="dms-popup-msg">
              {lastRun && lastRun.ok > 1
                ? `Your ${lastRun.ok} documents are awaiting approval.`
                : "Your document is awaiting approval."}
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
