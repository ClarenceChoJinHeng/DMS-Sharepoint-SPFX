import * as React from "react";
import {
  UPLOAD_PAUSE_SETTING,
  UPLOAD_PAUSE_MESSAGE,
  uploadsArePaused,
} from "../../../shared/uploadPause";
// Shared with the upload form deliberately: one definition of what a free name looks like, so the two
// screens can never disagree about it. See `nextAvailableName` for why it probes rather than appends.
import { nextAvailableName } from "../../../shared/uploadBatches";
import { useHcClearance } from "../../../shared/hcClearance";
import { isSystemAdmin } from "../../../shared/spGroups";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IBulkUploadProps } from "./IBulkUploadProps";
import {
  lookupFolderMapping,
  loadFolderMapRows,
  resolveMappedFolder,
  probeFolderByPath,
  probeFolderUploadAccess,
  ensureFolder,
  encodeServerRelativePath,
  FolderMapRow,
} from "../../../shared/dmsFolderMap";
/* ⚠ THE WRITE PROBE IS ON SINCE 2026-08-22, AND ITS ABSENCE WAS NEVER A BUG.
   This web part was existence-gated only *because* it wrote into DOCUMENTS, where a PIC holds Read
   by design (2026-08-09) — so an `AddListItems` probe there would have correctly and silently
   emptied this form for every uploader. That reasoning expired the day the target became the
   approval library, where a PIC holds `CRS Upload`.

   Now the probe is not merely safe but necessary: it is what stops an uploader picking a unit whose
   ACL has not been granted yet and meeting a 403 at write time — the same 403 that produced the
   upload form's original "not ready" empty state. Specs
   2026-08-12-provisioned-segment-visibility-design.md §2.2 and
   2026-08-22-bulk-upload-for-uploaders-design.md §2.3. */
import {
  AccessVerdict,
  filterProvisionedPaths,
  filterReachablePaths,
  mappedTermGuidSet,
  normalizeTermGuid,
} from "../../../shared/segmentReadiness";
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
  FALLBACK_FILE_TYPES,
  NO_TYPES_MESSAGE,
  readAllowedFileTypesField,
  resolveAllowedFileTypes,
} from "../../../shared/allowedFileTypes";
import {
  RoutingContext,
  effectiveHcLevel,
  isHcLevel,
  selectableLevels,
} from "../../../shared/hcRouting";
import {
  cachedHcLibraries,
  hcAvailable,
  cachedListTitle,
  LIST_SUFFIX,
  libraryTitle,
  libraryUrlSegment,
  libApiTitle,
} from "../../../shared/naming";
import {
  libraryHasColumns,
  REF_COLUMNS,
  BULK_IMPORT_COLUMN,
  SUBMISSION_FILE_COLUMN,
  KEYWORD_COLUMN,
} from "../../../shared/optionalColumns";
// Records the upload so My Submissions can still show the file after it is deleted (2026-08-27).
// Spec: docs/superpowers/specs/2026-08-27-submission-record-design.md
import { writeSubmissionRecord } from "../../../shared/spSubmissionRecords";
import { offersLegalPrivilege } from "../../../shared/legalPrivilege";
import { newReference } from "../../../shared/submissionGroups";
import { primeNames } from "../../../shared/spNaming";
import {
  stripBlockedChars,
  blockedCharsMessage,
} from "../../../shared/inputSanitize";
import {
  friendlyUploadError,
  blockedBeforeNetwork,
  BLOCKED_UPLOAD_MESSAGE,
} from "../../../shared/networkErrors";
import {
  NOTICE_ATTENTION,
  NOTICE_ATTENTION_CSS,
} from "../../../shared/noticeStyles";

/* ----------------------------------------------------------------------------
 * BULK UPLOAD — a duplicate of the `form` web part, with two behaviour changes:
 *   1. Files are marked `BulkImport`, which a Power Automate flow auto-approves — so they reach
 *      Documents through the SAME verified Auto-route flow as everything else, without a human
 *      approving each one. For HISTORICAL documents that were approved elsewhere already.
 *   2. Up to MAX_FILES files in ONE selection, sharing ONE destination folder
 *      and ONE metadata set.
 *
 * ⚠ THIS WROTE STRAIGHT INTO `Documents` UNTIL 2026-08-22, and several comments below still make
 * sense only against that history. It became an UPLOADER tool (client: "client doesnt want admin to
 * do the job"), and an uploader holds Read on `Documents` by design — so it writes into the approval
 * library like every other upload, and a flow moves it on. Spec
 * `2026-08-22-bulk-upload-for-uploaders-design.md`.
 *
 * ⚠ NOTHING HERE ENFORCES "historical only". A brand-new document put through this screen reaches
 * Documents unreviewed. That is the client's own rule, told to uploaders — not a guarantee the
 * software makes.
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

const FIELDS = {
  // Document Type column — internal name Document_x0020_Type (migrated from the old
  // frozen "Department_x0020_Type"; see 2026-07-24-document-type-internal-name-migration-design).
  documentType: "Document_x0020_Type",
  yearPeriod: "Year", // site column internal name (client kept plain "Year")
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
// labelCol/tidCol in CRS Config Levels JSON — those win over this table
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
/* `notice` added 2026-08-28 for the duplicate-name warning. It is neither an error (the files
   WERE added, and uploading them is allowed) nor a success — reporting it as either would
   misstate what happened. Amber, matching the clash dialog and every other advisory in the
   project. */
type ToastType = "error" | "success" | "notice";

type UploadMode = {
  key: string;
  label: string;
  side: "BusinessSegment" | "Project";
  termSetGuid: string;
  stagingFolder: string;
  /**
   * PERMISSIONED tiers only — indexed against levelValues/levelChoices, which the
   * segment-tree cascade fills. Below-Unit entries here would shift every index.
   * Same contract as Form.tsx.
   */
  levels: Level[];
  /** The full authored chain. Absent on the code fallbacks — see Form.tsx. */
  chain?: Level[];
  sortOrder: number;
};

// A fully authorised upload path for a restricted (non-privileged) user.
type ValidPath = { modeKey: string; chain: TermOption[] };

// One picked file. `key` survives removals so React rows stay stable.
type PickedFile = { key: string; file: File };

type Outcome = "uploaded" | "skipped" | "failed" | "tagFailed";
type FileResult = {
  name: string;
  outcome: Outcome;
  detail?: string;
  /**
   * A free name to offer, set only when this file was skipped for a NAME CLASH.
   *
   * Its presence is what makes the skip offerable — see the clash branch in `runUpload`. The
   * approved-side refusal deliberately leaves it undefined.
   */
  suggestedName?: string;
  /**
   * This skip was a NAME CLASH, whether or not a free name could be found.
   *
   * WARN: SEPARATE FROM `suggestedName`, because a clash with no free name must still REACH the
   * dialog. Client, 2026-08-27: *"the popup doesn't show two files like a normal upload form, it
   * shows only one for bulk upload."* Filtering the dialog on `suggestedName` alone silently omitted
   * the other refusal, and a dialog that accounts for SOME of the failures is worse than one that
   * accounts for none, because it looks complete.
   */
  nameClash?: boolean;
  /**
   * The clash was on the APPROVED side, so the original has already been through approval.
   *
   * Kept apart from `nameClash` because it only decides a warning line: a renamed copy of an approved
   * document is a SECOND document, which is a different thing from a first one.
   */
  approvedClash?: boolean;
};

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
  | "tagFailed";

// `sru` is the uploaded file's ServerRelativeUrl, captured once a file lands in
// the approval library (done / tagFailed). INERT since the per-row delete was
// removed on 2026-08-27 — nothing reads it. `pct` is real bytes-sent progress,
// 0-100, only meaningful while state === "uploading".
type LiveFile = {
  name: string;
  size: number;
  state: FileState;
  sru?: string;
  pct: number;
  /**
   * The File this row is for, so a post-run removal can take it out of the SELECTION.
   *
   * WARN: IDENTITY, NEVER NAME OR INDEX. This screen keeps original filenames, so two files added
   * from different folders can share one name and matching by name would drop the wrong row; and
   * successful files leave `picked` after a run, which shifts every index. Same reasoning as
   * `renameOverrides` and the post-run `setPicked` filter.
   */
  file?: File;
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
  /** The confidentiality label routing to the HC pair. Blank = not configured; see effectiveHcLevel. */
  hcConfidentialityLevel: string;
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
  // Blank, and always read together with hcAvailable() — see effectiveHcLevel.
  hcConfidentialityLevel: "",
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
/** The exact `body` an aborted XHR resolves with, so status 0 can be told from a blocked one. */
const ABORTED_BODY = "Upload was aborted.";

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
    /* WARN: THE ABORT MARKER IS LOAD-BEARING. An abort is ALSO status 0, and without a way to tell
       the two apart the caller would tell somebody who cancelled their own upload to go and check
       their browser extensions. */
    xhr.onabort = (): void =>
      resolve({ ok: false, status: 0, body: ABORTED_BODY });
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
  // Authorised paths, minus any whose leaf folder does not exist yet — spec
  // 2026-08-12-provisioned-segment-visibility-design.md.
  const [validPaths, setValidPaths] = useState<ValidPath[]>([]);
  // Authorised somewhere, but every path withheld for want of a folder. Distinguishes
  // an admin task (run reconciliation) from a membership problem, which send the
  // administrator to two different places.
  const [awaitingFolders, setAwaitingFolders] = useState<boolean>(false);

  const [modes, setModes] = useState<UploadMode[]>([]);
  const [settings, setSettings] = useState<DmsSettings>(DEFAULT_SETTINGS);

  // ── The one selection: files + the single metadata set they all share ──
  const [uploadMode, setUploadMode] = useState<string>("");
  const [picked, setPicked] = useState<PickedFile[]>([]);
  /**
   * Set once the HC library pair has been resolved, so the render re-evaluates.
   *
   * WARN: `hcAvailable()` READS A MODULE CACHE, WHICH REACT CANNOT SEE FILL. The pair resolves
   * asynchronously at mount, so `hcCtx()` computed during the first render answers "no HC" and
   * nothing re-renders when priming lands - the Highly Confidential level then appears only if some
   * other state change happens to force a repaint. That is the Bulk Upload half of the *"sometimes
   * shows and sometimes doesn't"* report of 2026-08-27; `Form.tsx` has had this flag since
   * 1.0.229.0 and this screen never got one.
   *
   * READ in `hcCtx()` rather than left unused, so nothing deletes it as dead. It can only ever be
   * true when `hcAvailable()` was true at priming, so it cannot widen the answer - it just makes the
   * recomputation happen.
   */
  const [hcReady, setHcReady] = useState(false);
  const [dragOver, setDragOver] = useState<boolean>(false);
  // Below-Unit tier selections keyed by the tier's `column`, mirroring Form.tsx.
  // Replaces the dedicated documentType/yearPeriod state, which WAS the hardcoded
  // shape — a chain that varies in length cannot have one useState per tier.
  const [tierValues, setTierValues] = useState<Record<string, string>>({});
  const [termCache, setTermCache] = useState<Record<string, TermOption[]>>({});
  // Children keyed by PARENT term GUID — feeds cascading below-Unit tiers (SubUnit).
  // `ok` separates "this term has no children" (a real answer that SKIPS the tier) from
  // "the call failed" (unknown, must block). Mirrors Form.tsx.
  const [childCache, setChildCache] = useState<
    Record<string, { terms: TermOption[]; ok: boolean }>
  >({});
  const [documentDate, setDocumentDate] = useState<string>("");
  const [confidentiality, setConfidentiality] = useState<string>("");
  const [legallyPrivileged, setLegallyPrivileged] = useState<boolean>(false);
  const [remark, setRemark] = useState<string>("");
  /* Free text for finding a document later (client, 2026-09-04). ONE value for the whole selection,
     like every other field on this screen — unlike the upload form, where it is per file. That is
     this screen's own model (one metadata set, one destination), not an inconsistency. */
  const [keyword, setKeyword] = useState<string>("");
  const [keywordBlocked, setKeywordBlocked] = useState<string | undefined>(undefined);
  const [remarkBlocked, setRemarkBlocked] = useState<string | undefined>(
    undefined,
  );
  // Vendor stays wired but unset — the select is hidden (client request 2026-07-28).
  const [vendor] = useState<string>("");

  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  /* True only while the pre-flight checks run - see `handleUpload`. Separate from `busy`, which
     means "files are being written": the two say different things to the person waiting. */
  const [preflight, setPreflight] = useState<boolean>(false);
  // Site-wide upload pause. Starts FALSE so a slow or failed read never hides the form; the
  // write-time re-check in handleUpload is the guard that actually refuses.
  const [paused, setPaused] = useState<boolean>(false);
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

  /* The per-file "Replace Existing File" prompt was REMOVED on 2026-08-27, along with the
     `overwrite=true` it enabled — see the clash branch in `runUpload`. It blocked a 50-file import on
     a modal per clash, and the answer it invited destroyed a pending document. Nothing replaces it
     per file; clashes are collected and decided ONCE below. */

  // Success dialog, matching the Form's. Separate from `toast` because this one is
  // modal and has to survive until the user chooses where to go next.
  const [doneOpen, setDoneOpen] = useState<boolean>(false);

  /**
   * Files skipped for a name clash that have a free name to offer, awaiting one decision.
   *
   * Keyed on the `File` OBJECT, not an index or a name: successful files leave the selection, which
   * shifts every index, and this screen keeps original filenames so two identical ones would collapse
   * together. Same reasoning as `runUpload`'s `renameOverrides`.
   */
  const [bulkClashes, setBulkClashes] = useState<
    { file: File; from: string; to: string; approvedSide: boolean }[]
  >([]);
  /** Set while a replace run is in flight, so the retry does not re-offer the same files. */

  /**
   * Clashes with NO free name to offer. They still have to appear in the dialog.
   *
   * WARN: A DIALOG THAT LISTS SOME OF THE FAILURES LOOKS COMPLETE, which is why this is a separate
   * list rather than an omission — client, 2026-08-27, on seeing one of two refusals in the popup.
   * Same split, and the same reason, as the upload form's own `clashNotes`.
   */
  const [bulkClashNotes, setBulkClashNotes] = useState<
    { name: string; detail: string }[]
  >([]);

  const activeMode = (): UploadMode | undefined =>
    modes.find((m) => m.key === uploadMode);

  /* ---------- Below-Unit tiers (mirrors Form.tsx) -------------------------- */

  const belowUnitTiers = (): Level[] =>
    effectiveOnDemandTiers(
      activeMode()?.chain ?? activeMode()?.levels ?? [],
      settings.termSets.yearPeriod,
      settings.termSets.documentType,
    );

  const permissionedLeafTerm = (): string => {
    const perm = activeMode()?.levels ?? [];
    return perm.length > 0 ? (levelValues[perm.length - 1] ?? "") : "";
  };

  /**
   * Which below-Unit tiers apply to this path, and each one's parent term. A cascading
   * tier applies only where the term above HAS children — not every Unit has SubUnits.
   * `unresolved` means "not known yet or lookup failed", and must block. Mirrors Form.tsx.
   */
  const tierPlan = (): {
    tiers: Level[];
    parents: string[];
    unresolved: string[];
  } => {
    const tiers: Level[] = [];
    const parents: string[] = [];
    const unresolved: string[] = [];
    let parent = permissionedLeafTerm();
    for (const t of belowUnitTiers()) {
      if ((t.termSet ?? "").trim()) {
        tiers.push(t);
        parents.push("");
        continue;
      }
      // Hidden, not disabled, until the tier above is chosen — whether it applies is not
      // yet knowable. The tier above is blank and already required, so nothing is lost.
      if (!parent) continue;
      const entry = childCache[parent.trim().toLowerCase()];
      const decision = decideTier(
        t,
        entry !== undefined && entry.ok ? entry.terms.length : undefined,
      );
      if (decision === "skip") continue;
      tiers.push(t);
      parents.push(parent);
      if (decision === "unresolved") unresolved.push(t.label);
      parent = tierValues[t.column] ?? "";
    }
    return { tiers, parents, unresolved };
  };

  /** `termSet` present → flat options from that set. Absent → children of the tier above. */
  const tierOptions = (t: Level, parent: string): TermOption[] => {
    const set = (t.termSet ?? "").trim().toLowerCase();
    if (set) return termCache[set] ?? [];
    const key = parent.trim().toLowerCase();
    return key ? (childCache[key]?.terms ?? []) : [];
  };

  // Drops any selection no longer present in its tier's current options — a stale SubUnit
  // from a different Unit is reported missing rather than filed under.
  const tierSelections = (): Record<string, TierSelection | undefined> => {
    const out: Record<string, TierSelection | undefined> = {};
    const plan = tierPlan();
    plan.tiers.forEach((t, i) => {
      const opt = tierOptions(t, plan.parents[i]).find(
        (o) => o.id === (tierValues[t.column] ?? ""),
      );
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

  // Fetch the term sets a configured chain adds. No-op on an unmigrated site, where
  // both are cached from mount. Per-set try/catch because Promise.allSettled is
  // unavailable on this tsconfig (CLAUDE.md gotcha #3). Declared after loadTermSet
  // deliberately — a const arrow is not hoisted.
  useEffect(() => {
    let cancelled = false;
    const missing = belowUnitTiers()
      .map((t) => (t.termSet ?? "").trim())
      .filter((g) => g && !termCache[g.toLowerCase()]);
    if (missing.length === 0) return;
    (async () => {
      const loaded: Record<string, TermOption[]> = {};
      for (const guid of missing) {
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
  }, [
    uploadMode,
    modes,
    settings.termSets.yearPeriod,
    settings.termSets.documentType,
  ]);

  const loadTermChildrenRef = useRef<
    ((termSetId: string, termId: string) => Promise<TermOption[]>) | undefined
  >(undefined);

  // Children for cascading below-Unit tiers, keyed on the selections above them, so
  // choosing a different Unit fetches that unit's own SubUnits. No-op when nothing
  // cascades, which is every site today. Mirrors Form.tsx.
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
        // Per-parent try/catch — Promise.allSettled is unavailable here (gotcha #3).
        //
        // ok:false on failure, NOT an empty list. Empty is a real answer that skips the
        // tier, so recording a failure as empty would silently shorten the path.
        try {
          loaded[parent.toLowerCase()] = {
            terms: await load(md.termSetGuid, parent),
            ok: true,
          };
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

  // Hand the loader to the cascading-tier effect above — a const arrow is not hoisted.
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
   * Primed per call so there is no ordering dependency; the cache short-circuits after the first.
   */
  const listName = async (suffix: string): Promise<string> => {
    await primeNames(context.spHttpClient, siteUrl);
    return encodeURIComponent(cachedListTitle(suffix));
  };

  /**
   * A comparable fingerprint of a folder chain — order, names and columns. Same contract as
   * Form.tsx: a re-saved row with identical content must not block an upload, but anything that
   * would move a folder or write a different column must.
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
   * Re-read this mode's chain from DMS Config, for the staleness guard in `handleUpload`.
   *
   * `undefined` means "could not tell" — a failed read, or a row that has gone — and the caller
   * must not block on it.
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
      const rows = ((await res.json()).value ?? []) as Array<{
        Levels?: string;
      }>;
      if (rows.length === 0) return undefined;
      const parsed = parseLevels(rows[0].Levels ?? "");
      return parsed.length > 0 ? parsed : undefined;
    } catch {
      return undefined;
    }
  };

  /**
   * Read the site-wide upload pause, LIVE. Same contract as the upload form's copy: called at mount
   * for the banner, and again immediately before writing, because settings are read once in a
   * mount-time effect and a tab opened before the pause would otherwise write straight through it.
   *
   * Returns `false` on any failure — an unreadable config row must never block every uploader.
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
      const rows = ((await res.json()).value ?? []) as Array<{
        SettingValue?: string;
      }>;
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
    if (!res.ok)
      throw new Error(`${decodeURIComponent(config)} list not found`);
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
    if (!res.ok)
      throw new Error(`${decodeURIComponent(groupMap)} list not found`);
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
        documentType:
          get("col_documentType") ?? DEFAULT_SETTINGS.columns.documentType,
        yearPeriod:
          get("col_yearPeriod") ?? DEFAULT_SETTINGS.columns.yearPeriod,
        documentDate:
          get("col_documentDate") ?? DEFAULT_SETTINGS.columns.documentDate,
        confidentiality:
          get("col_confidentiality") ??
          DEFAULT_SETTINGS.columns.confidentiality,
        vendor: get("col_vendor") ?? DEFAULT_SETTINGS.columns.vendor,
        remark: get("col_remark") ?? DEFAULT_SETTINGS.columns.remark,
        legallyPrivileged:
          get("col_legallyPrivileged") ??
          DEFAULT_SETTINGS.columns.legallyPrivileged,
        businessSegmentLabel:
          get("col_businessSegment") ??
          DEFAULT_SETTINGS.columns.businessSegmentLabel,
        businessSegmentTid:
          get("col_businessSegmentTid") ??
          DEFAULT_SETTINGS.columns.businessSegmentTid,
      },
      // The library's LIVE title wins over the config row — see the same note in Form.tsx.
      // NOTE this holds the TITLE. The path-swap sites below need the URL SEGMENT, which is a
      // different string since the rename, and they read libraryUrlSegment() instead.
      stagingLibrary: ((): string => {
        const configured = (get("stagingLibrary") ?? "").trim();
        return configured.length > 0 && configured !== "Staging"
          ? configured
          : libraryTitle();
      })(),
      legallyPrivilegedFor:
        get("legallyPrivilegedFor") ?? DEFAULT_SETTINGS.legallyPrivilegedFor,
      hcConfidentialityLevel:
        get("hcConfidentialityLevel") ??
        DEFAULT_SETTINGS.hcConfidentialityLevel,
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
      return (
        ((d.value ?? []) as Array<{ Id?: number }>)
          // Both branches spelled out. `!= null` meant "neither null nor undefined" — correct,
          // but invisible, and "fixing" it to `!== null` would let undefined through and turn a
          // missing id into the string "undefined".
          .map((g) => (g.Id !== null && g.Id !== undefined ? String(g.Id) : ""))
          .filter(Boolean)
      );
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
      if (
        isLeafChainValid(
          chain.map((c) => c.id),
          leaf.termGuid,
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
      if (opts.length === 1)
        values[i] = opts[0].id; // auto-lock
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
      // Mount-time read, for the banner only — so a paused uploader is told before filling the form
      // in rather than after choosing files. Deliberately its own statement, NOT another element in
      // the Promise.all below: that array is positionally destructured, so adding to it silently
      // rebinds every variable after it. Never throws; false on any failure.
      readUploadPause()
        .then(setPaused)
        .catch(() => setPaused(false));
      const [
        loadedSettings,
        rawModes,
        groupMap,
        userGroupIds,
        admin,
        folderMapRows,
      ] = await Promise.all([
        loadSettings().catch(() => DEFAULT_SETTINGS),
        loadModes().catch(() => DEFAULT_MODES),
        loadGroupMap().catch(() => [] as GroupMapRow[]),
        loadUserGroupIds(),
        loadIsAdmin(),
        // Which folders EXIST. `null` = unknown, and unknown offers everything: an
        // unreadable list proves nothing, and emptying every dropdown over a transient
        // error takes the form down for the whole site. Empty is not unknown.
        loadFolderMapRows(context.spHttpClient, siteUrl).catch((err) => {
          console.error(
            "CRS Folder Map read failed — cannot tell which folders exist, so every authorised path will be offered.",
            err,
          );
          return null as FolderMapRow[] | null;
        }),
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
      // Seed the by-term-set cache with the two sets every site has, so an
      // unmigrated site makes no extra requests for its below-Unit tiers.
      setTermCache({
        [loadedSettings.termSets.documentType.trim().toLowerCase()]: docTypes,
        [loadedSettings.termSets.yearPeriod.trim().toLowerCase()]: years,
      });
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
        const authorised = await resolveValidPaths(loadedModes, membership);
        // Authorisation says where the user MAY file; the Folder Map says where they
        // CAN. Offering the difference is what let a newly grouped user pick a unit
        // before reconciliation had run, and a new segment appear before it had any
        // folders. Gating the LEAF empties the tiers above it, so an unprovisioned
        // segment disappears with no per-segment rule.
        const provisioned = filterProvisionedPaths(
          authorised,
          mappedTermGuidSet(folderMapRows),
        );
        /* Existence is a property of the FOLDER; being able to upload is a property of the folder AND
           the user, and no amount of Folder Map reading answers the second. Reconciliation creates
           folders from the term tree and grants group ACLs in a SEPARATE pass, so a unit can have a
           folder for days before anyone can write into it.

           ⚠ An unreadable map or an inconclusive probe must offer EVERYTHING — a short verdict array
           fails open too. A transient error that empties every dropdown takes this form down
           site-wide, and the upload-time checks are the backstop that makes the degraded path safe. */
        const reachable = filterReachablePaths(
          provisioned.paths,
          await Promise.all(
            provisioned.paths.map(async (p): Promise<AccessVerdict> => {
              const leaf = p.chain[p.chain.length - 1];
              const row = (folderMapRows ?? []).find(
                (r) =>
                  normalizeTermGuid(r.termGuid) === normalizeTermGuid(leaf?.id),
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
        paths = reachable.paths;
        setValidPaths(paths);
        setAwaitingFolders(
          provisioned.known &&
            reachable.known &&
            paths.length === 0 &&
            authorised.length > 0,
        );
      }

      // Open on a destination the user can actually reach, so both cards are
      // populated from the first paint. The batch panel used to do this on
      // "+ Add batch"; with one selection there is no such moment.
      const offerable = new Set(paths.map((p) => p.modeKey));
      const defaultMode = isPrivileged
        ? (loadedModes.filter((m) => m.side === "BusinessSegment")[0] ??
          loadedModes[0])
        : (loadedModes.find(
            (m) => m.side === "BusinessSegment" && offerable.has(m.key),
          ) ?? loadedModes.find((m) => offerable.has(m.key)));
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
    setTierValues({});
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
  /**
   * Forget the previous run once the selection changes.
   *
   * WARN: WITHOUT THIS THE LOG OUTLIVES WHAT IT DESCRIBES (client, 2026-08-27: *"when I remove the
   * file it should also remove the log below for bulk upload"*). Seen live: a skipped-file panel
   * reading *"0 uploaded, 1 skipped - TEST - TEST - TEST - 26-08-26 (2).pdf"* sitting under a picker
   * that held a completely different document. The reader has no way to tell a stale result from a
   * fresh one, and on this screen the result is the only place a refusal is explained - so a stale
   * one is worse than none.
   *
   * `runError` goes too, for the same reason. `live` is NOT touched here: `dropFromRun` owns it and
   * deliberately nulls it only when the last row goes, because an empty array is truthy and would
   * leave a dead panel with no way back to the picker.
   */
  const clearRunLog = (): void => {
    setResults(null);
    setRunError(null);
  };

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
      const next = [...prev, ...added];

      /* ⚠ DUPLICATE NAMES IN ONE SELECTION WERE ACCEPTED SILENTLY UNTIL 2026-08-28 (found by the
         client: *"I am allowed to select two same file name and insert into bulk upload"*). This
         screen keeps ORIGINAL filenames — it composes nothing — so two files picked from different
         folders can carry the same name, and `addFiles` checked only the file type and the 50 cap.

         The upload form BLOCKS Save on the equivalent collision. This does not block, deliberately:
         two genuinely different documents can share a filename (different folder, different year),
         and refusing the selection would stop a legitimate import. The run-time checks already
         handle it correctly — the second file collides with the first in the approval library and is
         offered a free name — so nothing is ever lost. What was missing is being TOLD, at the moment
         of picking rather than after a fifty-file run.

         ⚠ THE CASE THIS REALLY EXISTS FOR: if both same-named files also clash with the approved
         side and the admin presses *Replace the filed document*, both are retried. One replaces the
         filed copy and the other is renamed — a correct end state reached from an instruction that
         sounded like it concerned a single document.

         Named, never counted, for the same reason as the rejected-type toast above: "2 duplicates"
         leaves an admin hunting through fifty rows. */
      const seen: Record<string, number> = {};
      for (const p of next) {
        const k = p.file.name.toLowerCase();
        seen[k] = (seen[k] ?? 0) + 1;
      }
      const dupes = Object.keys(seen)
        .filter((k) => seen[k] > 1)
        .map(
          (k) =>
            (
              next.filter((p) => p.file.name.toLowerCase() === k)[0] as {
                file: File;
              }
            ).file.name,
        );
      if (dupes.length > 0) {
        showToast(
          `More than one file is called ${dupes.join(", ")}. They can all be uploaded — the later ` +
            `${dupes.length === 1 ? "one is" : "ones are"} offered a free name when you upload — but ` +
            `check this is not the same document picked twice.`,
          "notice",
        );
      }
      return next;
    });

    // Clear the input so the same file can be re-picked after a removal.
    if (fileRef.current) fileRef.current.value = "";
    // WARN: THE PREVIOUS RUN LOG MUST GO WITH THE SELECTION CHANGE - see clearRunLog.
    clearRunLog();
  };

  const removeFile = (key: string): void => {
    setPicked((prev) => prev.filter((p) => p.key !== key));
    clearRunLog();
  };

  /**
   * Drop one row from the post-run progress list AND from the selection (client, 2026-08-27:
   * *"i want it so it will be easier for them to deselect"*).
   *
   * WARN: THIS IS NOT A DELETE. It touches nothing on the server - it is the same de-selection the
   * picker offers before a run, made reachable afterwards, because once a run has happened the
   * progress list REPLACES the picker and there was no way to drop a single file without pressing
   * Cancel and re-picking the whole set. The removed row was never uploaded (see `canDrop` at the
   * call site), so there is nothing anywhere to undo.
   *
   * Matched on the File OBJECT, never the index: a run removes its successful files from `picked`,
   * so the two lists stop lining up the moment anything succeeds.
   */
  const dropFromRun = (idx: number): void => {
    const row = (live ?? [])[idx];
    if (!row) return;
    /* WARN: EMPTIED MEANS null, NOT []. The render is `live ? progress : picked.length === 0 ?
       dropzone : picker`, and an EMPTY ARRAY IS TRUTHY - so dropping the last row would leave a dead
       empty panel with no dropzone and no "Add more", i.e. no way back to picking files at all. */
    setLive((prev) => {
      if (!prev) return prev;
      const next = prev.filter((_, i) => i !== idx);
      return next.length > 0 ? next : null;
    });
    if (row.file) {
      setPicked((prev) => prev.filter((p) => p.file !== row.file));
      // A pending rename offer for this file would otherwise outlive it and re-upload it on Proceed.
      setBulkClashes((prev) => prev.filter((c) => c.file !== row.file));
    }
    // The summary below described a run that included this row, so it no longer describes anything.
    clearRunLog();
  };

  /* WARN: THE PER-ROW DELETE WAS REMOVED (2026-08-27, client: *"pls remove the delete functionality
     ... don't let it be there since PIC cannot delete the file until they request"*). A `deleteUploaded`
     handler used to sit here, wired to an X on every completed row. Four things were wrong with it, and
     the client only had to spot the first:

       1. BULK UPLOAD IS NOT ADMIN-ONLY ANY MORE (2026-08-22). It is open to every uploader, and a PIC
          holds no delete anywhere — they lost `DELS` on 2026-08-20 and must raise a REQUEST that a Head
          of Unit decides. A delete button on their own screen contradicts the whole workflow.
       2. IT PURGED RATHER THAN RECYCLED. It sent a plain `X-HTTP-Method: DELETE`, which does NOT go to
          the recycle bin. The approved deletion-request flow deliberately calls `recycle()` so the
          document is restorable for 93 days — *that* is what makes approving one reasonable. This was
          strictly more destructive than the sanctioned route, with no approval in front of it.
       3. IT REPORTED SUCCESS AFTER DELETING NOTHING. It treated a 404 as success, and the `sru` it held
          points into the APPROVAL library — which auto-approve plus Auto-route empties within about a
          minute of the upload. So clicking it after the file had routed marked the row "deleted" and
          toasted *"deleted from Documents"* while the document sat untouched in `Documents`.
       4. EVERY STRING SAID "Documents", stale since this screen was repointed at the approval library.

     `sru` is still captured on the row and is now inert. Left in place so the plumbing does not have to
     be rebuilt if a future ADMIN-ONLY screen ever needs it — and it is data, not an action, so it
     cannot be clicked. If a delete is ever wanted here again it must recycle, must be gated on the
     persona rather than on the row's state, and must re-resolve the file rather than trusting `sru`. */

  /* ---------- Documents-library path swap --------------------------------- */

  // The DMS Folder Map stores the STAGING unit folder's UniqueId. The Documents
  // library holds a mirrored tree, so the Documents unit folder is found by
  // swapping the library segment of the resolved Staging path:
  //   /sites/<web>/Staging/<rest>  ->  /sites/<web>/Shared Documents/<rest>
  // Anchored on the web-relative prefix rather than a global replace, so a
  // folder that happens to be named "Staging" deeper in the tree is not mangled.
  /* ---------- Highly Confidential ------------------------------------------
     Spec: docs/superpowers/specs/2026-08-15-highly-confidential-library-design.md

     This page writes STRAIGHT into the approved library, bypassing approval entirely — so for a
     Highly Confidential document the destination is `HC Documents`, not `Documents`.

     NO WRITE PROBE HERE, unlike the upload form, and that is deliberate rather than an omission.
     This page is admin-only (`pageAccessPolicy` → adminOnly) and is already existence-gated only,
     for the documented reason that an AddListItems probe against Documents would empty the form for
     every PIC — `UPL` holds Read there by design. Admins are exempt from the segment probes
     elsewhere for the same reason. A write they are not entitled to fails LOUDLY here: the file
     lands nowhere and the run reports it per file. */
  /* Resolve the names, then announce it. `hcReady` is read nowhere except as a re-render trigger -
     see its declaration for why a plain module cache is not enough. */
  useEffect(() => {
    let live = true;
    primeNames(context.spHttpClient, siteUrl)
      .then(() => {
        if (live) setHcReady(hcAvailable());
      })
      .catch(() => undefined);
    return () => {
      live = false;
    };
  }, []);

  /**
   * Which units this admin can file Highly Confidential documents into.
   *
   * ⚠ THIS SCREEN HAD NO CLEARANCE PROBE AT ALL UNTIL 2026-08-27, and the comment above explained why
   * as "this page is admin-only". **IT STOPPED BEING ADMIN-ONLY ON 2026-08-22**, when Bulk Upload was
   * opened to every uploader - so `canWriteHc: hcAvailable()` offered the Highly Confidential level to
   * anyone who could open the page. Nothing was ever filed wrongly (the HC unit folder is RESOLVED,
   * never created, so the write refuses), but the level was shown to people who cannot use it - the
   * exact opposite of the upload form's rule, and what the client asked to close.
   *
   * The reasoning that DOES still hold is the one about `Documents`: an `AddListItems` probe against
   * the approved side would empty this form for every PIC, because `UPL` holds Read there by design.
   * That is why the existence gate on the normal path is unchanged - this probe asks about the HC
   * APPROVAL library, which is where an HC upload actually lands.
   */
  const hcWrite = useHcClearance(
    context.spHttpClient,
    siteUrl,
    permissionedLeafTerm(),
    hcReady,
  );

  const hcCtx = (): RoutingContext => {
    /* `hcReady` is read here on purpose: it is what ties this computation to the async resolution of
       the HC pair. `hcAvailable()` remains the source of truth - `hcReady` is only ever set FROM it. */
    const available = hcReady || hcAvailable();
    return {
      hcLevel: effectiveHcLevel(settings.hcConfidentialityLevel, available),
      hcAvailable: available,
      /* PER UNIT, not per site. `undefined` until the probe answers, which `canOfferHc` reads as no -
         fails closed, as it must: an over-offered HC level publishes a secret. */
      canWriteHc: hcWrite[permissionedLeafTerm()],
    };
  };

  /** True when the chosen level routes to the HC pair. */
  const uploadingHc = (): boolean =>
    isHcLevel(
      options.confidentiality.find((o) => o.id === confidentiality)?.label ??
        "",
      hcCtx(),
    );

  /**
   * The library this upload is WRITTEN to — the title metadata is tagged against.
   *
   * ⚠ THE APPROVAL LIBRARY SINCE 2026-08-22, not `Documents`. Bulk Upload became an uploader tool
   * (client: "client doesnt want admin to do the job"), and an uploader holds Read on `Documents` by
   * design — so it now writes where they can, is auto-approved by a flow keyed on `BulkImport`, and
   * reaches `Documents` through the SAME verified Auto-route flow as every other document. Spec
   * `2026-08-22-bulk-upload-for-uploaders-design.md`.
   *
   * ⚠ THE FILE IS PLACED BY FOLDER ID AND TAGGED BY LIBRARY TITLE, AND THOSE TWO MUST AGREE.
   * `GetFolderById` reaches into any library, so a mismatch places the file correctly and then tags
   * an item id in the WRONG list — 404 on a good day, and on a bad one it finds a different document
   * holding that id and writes this metadata onto it, reporting success (1.0.170.0, HC uploads).
   */
  const targetListTitle = (): string =>
    uploadingHc() ? libApiTitle("StagingHC") : libApiTitle("Staging");

  /**
   * The APPROVED-side folder this file will end up in, for the duplicate check only.
   *
   * ⚠ THE CHECK MUST FOLLOW THE FILE. Until 2026-08-22 this web part wrote into `Documents`, so its
   * existing duplicate probe was against `Documents` by construction. Repointing the WRITE to the
   * approval library would silently have repointed the CHECK with it — leaving nothing looking at
   * where the document actually lands, and surfacing the clash inside Auto-route's `Copy file`, where
   * the outcomes are a file stuck approved-but-not-routed or an approved record overwritten.
   */
  const approvedSideSegment = (): string =>
    uploadingHc()
      ? (cachedHcLibraries()?.documents.urlSegment ?? DOCUMENTS_URL_SEGMENT)
      : DOCUMENTS_URL_SEGMENT;

  /**
   * The library segment the file is WRITTEN into — the HC approval library for an HC upload.
   *
   * ⚠ RETURNS `undefined` RATHER THAN FALLING BACK when the HC pair is unresolved. A fallback here
   * files a Highly Confidential document into the open approval library, where every uploader in the
   * unit can see it while it is pending. `hcRouting.ts` fails closed for the same reason; so does this.
   */
  const writeSideSegment = (): string | undefined =>
    uploadingHc()
      ? cachedHcLibraries()?.approval.urlSegment
      : libraryUrlSegment();

  /**
   * Mirror a path from one library into another, keeping the folder tree.
   *
   * Slices on the URL SEGMENT, never the title: since the rename the two differ ("Approval Document"
   * vs "/ApprovalDocument"), and a title here matches nothing — the caller then reports it could not
   * work out the path, for every folder.
   */
  const mirrorPath = (
    sru: string,
    fromSegment: string | undefined,
    toSegment: string | undefined,
  ): string | null => {
    if (!fromSegment || !toSegment) return null;
    const prefix = `${webSru}/${fromSegment}/`;
    if (sru.toLowerCase().indexOf(prefix.toLowerCase()) !== 0) return null;
    return `${webSru}/${toSegment}/${sru.slice(prefix.length)}`;
  };

  /** The approved-side twin of a path in the library this upload writes to. For the clash check. */
  const toDocumentsPath = (sru: string): string | null =>
    mirrorPath(sru, writeSideSegment(), approvedSideSegment());

  /**
   * Is a file of this name already filed on the approved side?
   *
   * `true` only on a CONFIRMED hit. Anything else — a 404, a 403, a throttle, a network failure —
   * answers false, which lets the upload proceed: this is an additional guard, and an unanswerable
   * check must never become a new way to be blocked. Same direction as the approval-library probe
   * beside it, which also carries on when its request fails.
   *
   * ⚠ The folder is reached by PATH, so it goes through `probeFolderByPath` and its OData parameter
   * alias form. An inline quoted literal returns HTTP 400 rather than 404 once the path is deep
   * enough, which reads as "no clash" and is exactly the silent failure this exists to prevent
   * (gotcha #9).
   */
  /**
   * Whether `fileName` is already filed on the approved side, and every name that IS.
   *
   * WARN: IT LISTS THE FOLDER RATHER THAN PROBING ONE NAME (2026-08-27) — the same single request the
   * old `Exists` probe cost, and it yields the names a suggestion needs. Mirrors `approvedClashInfo`
   * in `Form.tsx`; the two screens must not disagree about what is already filed.
   *
   * Still FAILS OPEN: an unanswerable read reports no clash, which is the behaviour before the check
   * existed. `taken` then comes back empty, and the caller still has the known-taken name to work from.
   */
  const approvedSideClashInfo = async (
    folderPath: string,
    fileName: string,
  ): Promise<{ clash: boolean; taken: string[] }> => {
    try {
      const probe = await probeFolderByPath(
        context.spHttpClient,
        siteUrl,
        folderPath,
      );
      // No folder means nothing has been routed here yet — there cannot be a clash.
      if (!probe.folder) return { clash: false, taken: [] };
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${probe.folder.uniqueId}')/Files?$select=Name&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return { clash: false, taken: [] };
      const rows = ((await res.json()).value ?? []) as Array<{ Name?: string }>;
      const taken = rows.map((r) => r.Name ?? "").filter((x) => x.length > 0);
      return {
        clash: taken.some((x) => x.toLowerCase() === fileName.toLowerCase()),
        taken,
      };
    } catch {
      return { clash: false, taken: [] };
    }
  };

  /** Names already in the approval-library destination folder. Empty on any failure. */
  const approvalFolderNames = async (folderId: string): Promise<string[]> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files?$select=Name&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return [];
      const rows = ((await res.json()).value ?? []) as Array<{ Name?: string }>;
      return rows.map((r) => r.Name ?? "").filter((x) => x.length > 0);
    } catch {
      return [];
    }
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
    // Derived from the chain, not hardcoded. Every tier is required — an optional
    // one left blank files documents at inconsistent depths inside a single unit.
    missing.push(
      ...buildOnDemandSegments(tierPlan().tiers, tierSelections()).missing,
    );
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
      /** Below-Unit folder names in path order. Replaces the fixed year/docType pair. */
      tierSegments: string[];
      /** Metadata for those same tiers, already in validateUpdateListItem shape. */
      tierFormValues: Array<{ FieldName: string; FieldValue: string }>;
      confLabel: string;
      vendorLabel: string;
    },
    onState: (fileIndex: number, state: FileState, sru?: string) => void,
    onPct: (fileIndex: number, pct: number) => void,
    /**
     * Names to upload specific files under, replacing their own — the answer to the clash dialog.
     *
     * ⚠ KEYED ON THE `File` OBJECT, never on an index or a name. Indices shift the moment successful
     * files leave the selection, and this screen keeps ORIGINAL filenames, so two identical names in
     * one selection would collapse to one match. The File reference is stable and unique.
     */
    renameOverrides?: Map<File, string>,
    /**
     * Files whose APPROVED-SIDE clash the uploader chose to proceed past.
     *
     * Client's revised brief, 2026-08-27: a document already FILED may be replaced by sending a new
     * copy through approval; a PENDING draft may never be overwritten by anyone.
     *
     * WARN: SETS NO `overwrite` FLAG. The name is free in the approval library, so this is an ordinary
     * upload - it only stops the approved-side check refusing. The approval-library check is never
     * skipped.
     */
    replaceApprovedFiles?: Set<File>,
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
        runError: `"${leaf.label}" has no folder yet. Your administrator needs to give it an abbreviation in the CRS Term Abbreviation list, then run folder reconciliation.`,
        results: [],
      };
    }

    // Resolve the Staging unit folder's CURRENT path (rename-proof), then swap
    // the library segment to reach the mirrored Documents unit folder.
    const unitFolder = await resolveMappedFolder(
      context.spHttpClient,
      siteUrl,
      mapping.folderUniqueId,
      mapping.folderUrl,
    );
    const stagingSru = unitFolder.serverRelativeUrl;
    if (!stagingSru) {
      return {
        // Only a confirmed 404 means the folder is gone. Anything else — a 403, a throttle —
        // is a lookup that could not be answered, and sending an admin to re-run
        // reconciliation over an intact tree fixes nothing while the uploader stays blocked.
        runError: unitFolder.confirmedMissing
          ? "The mapped unit folder no longer exists. Ask an administrator to re-run reconciliation."
          : `Your folder could not be opened (HTTP ${unitFolder.status}). That is usually a permissions ` +
            `problem rather than a missing folder — ask an administrator to check your access to this unit.`,
        results: [],
      };
    }

    if (labels.tierSegments.length === 0) {
      return {
        runError: "No destination folder could be resolved for this upload.",
        results: [],
      };
    }

    /* ── The HC pair, when this upload is Highly Confidential ──────────────────
       ⚠ THE FILE MUST NOT BE WRITTEN INTO THE NORMAL APPROVAL LIBRARY. `stagingSru` comes from the
       Folder Map, which only ever describes the NORMAL library — so writing there directly files a
       Highly Confidential document where every uploader in the unit can see it while it is pending
       (draft security has been off since 2026-08-19). Found live on 2026-08-22, immediately after
       this web part was repointed: the old code reached HC only through the `toDocumentsPath` segment
       swap, and removing that swap took HC routing with it.

       It also produced the visible symptom rather than the dangerous one: the file landed in the
       normal library while `validateUpdateListItem` addressed the HC one, and **item ids are
       per-LIST**, so tagging failed with HTTP 400. The 400 was the lucky part.

       ⚠ RESOLVED, NEVER CREATED. A folder created here by an uploader inherits the HC library ROOT's
       permissions instead of carrying the unit's — which is exactly how an HC document becomes
       readable by the people the unit ACL excludes. Absent ⇒ refuse and name reconciliation. */
    let unitSru = stagingSru;
    if (uploadingHc()) {
      const hcPath = mirrorPath(
        stagingSru,
        libraryUrlSegment(),
        cachedHcLibraries()?.approval.urlSegment,
      );
      if (!hcPath) {
        return {
          runError:
            "The Highly Confidential libraries are not set up on this site, so this document cannot be " +
            "filed. Ask an administrator to create the HC library pair and run reconciliation — nothing " +
            "was uploaded.",
          results: [],
        };
      }
      const hcProbe = await probeFolderByPath(
        context.spHttpClient,
        siteUrl,
        hcPath,
      );
      if (!hcProbe.folder) {
        return {
          runError: hcProbe.confirmedMissing
            ? `This unit has no folder in the Highly Confidential library yet (${hcPath}). Ask an administrator to run folder reconciliation — it creates the folder with the correct permissions. Do not create it by hand: a hand-made folder inherits the library root's permissions and would widen access.`
            : `Could not open this unit's Highly Confidential folder (HTTP ${hcProbe.status}). That is usually a permissions problem rather than a missing folder — you may not be cleared for Highly Confidential documents in this unit.`,
          results: [],
        };
      }
      unitSru = hcProbe.folder.serverRelativeUrl;
    }

    /* ── The approved-side unit folder ─────────────────────────────────────────
       Resolved from `unitSru`, so on an HC upload it mirrors `HC Approval Document` into
       `HC Documents` rather than the normal pair. Computed AFTER the HC block above for exactly that
       reason — mirroring `stagingSru` would have described the wrong pair entirely.

       ⚠ STILL CHECKED EVEN THOUGH THE WRITE NO LONGER GOES HERE. Auto-route copies the approved file
       into this exact folder, so a missing one turns every file in this run into a document that is
       approved and never arrives — discovered days later by someone looking for it. Refusing now,
       naming the folder, is the same information at the only moment anyone can act on it.

       It is deliberately NOT auto-created: a folder created on the upload path would inherit the
       library root's ACL and silently widen access. Reconciliation creates the mirrored tree in all
       CRS libraries and locks each folder to its Group Map groups. */
    const docsUnitPath = toDocumentsPath(unitSru);
    if (!docsUnitPath) {
      console.error(
        "Could not mirror to the approved side. unitSru:",
        unitSru,
        "from segment:",
        writeSideSegment(),
        "to segment:",
        approvedSideSegment(),
      );
      return {
        runError:
          "Could not work out where this unit's approved documents live. The library's folder path " +
          "did not match what was expected — ask an administrator to check the libraries have not " +
          "been renamed since the last deployment.",
        results: [],
      };
    }
    const docsProbe = await probeFolderByPath(
      context.spHttpClient,
      siteUrl,
      docsUnitPath,
    );
    if (!docsProbe.folder) {
      // Only claim the folder is missing when SharePoint actually said 404. Anything else — a
      // throttle, a permission, a malformed request — is "could not tell", and sending someone to
      // create a folder that already exists is how the first report of this went wrong.
      return {
        runError: docsProbe.confirmedMissing
          ? `The matching folder does not exist yet (${docsUnitPath}). Ask an administrator to run folder reconciliation — it creates this folder with the correct permissions. Do not create it by hand: a hand-made folder inherits the library's root permissions and would widen access.`
          : `Could not verify the destination folder (HTTP ${docsProbe.status}). This usually means SharePoint was busy — most often because folder reconciliation is running at the same time. Wait for it to finish and try again.`,
        results: [],
      };
    }

    setStatus("Preparing destination folders…");
    /* Walk the configured below-Unit chain in the APPROVAL library — where the file is now written.
       Every folder here inherits the unit folder's ACL, exactly as Year / Document Type always did;
       nothing below Unit breaks inheritance (client decision, 2026-08-06).

       The mirrored folders on the approved side are created by AUTO-ROUTE when it copies the file,
       so they are deliberately not ensured here. */
    let parentSru = unitSru;
    let destFolder: Awaited<ReturnType<typeof ensureFolder>> | undefined;
    for (const name of labels.tierSegments) {
      const made = await ensureFolder(
        context.spHttpClient,
        siteUrl,
        parentSru,
        name,
      );
      if (!made) {
        return {
          runError: `Could not create the "${name}" folder.`,
          results: [],
        };
      }
      parentSru = made.serverRelativeUrl;
      destFolder = made;
    }
    if (!destFolder) {
      return {
        runError: "No destination folder could be resolved for this upload.",
        results: [],
      };
    }
    const folderId = destFolder.uniqueId;
    /* The matching path on the approved side, for the per-file duplicate check below. Derived from
       the folder actually written to, so the two can never describe different places. */
    const approvedSideLeaf = toDocumentsPath(destFolder.serverRelativeUrl);

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
    // THE LABEL, not the dropdown's raw value — which is the TERM ID. The HC check three lines below
    // already resolves the label; this did not, so `legallyPrivilegedFor` had to hold a term GUID to
    // match anything (2026-08-19). One shape for both rules: the config row holds LEVEL NAMES.
    const privilegedApplies = offersLegalPrivilege(
      options.confidentiality.find((o) => o.id === confidentiality)?.label ??
        "",
      settings.legallyPrivilegedFor,
    );
    const formValues: Array<{ FieldName: string; FieldValue: string }> = [
      // One entry per below-Unit tier, in chain order, resolved by the caller.
      ...labels.tierFormValues,
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
    /* ── The bulk-import marker, and the submission reference ────────────────
       `BulkImport` is what the auto-approve flow keys its trigger condition on, so it is the one
       thing separating a historical import from every other file in the approval library.

       ⚠ WRITTEN ONLY WHERE THE COLUMNS ARE CONFIRMED PRESENT. One unknown field name fails the WHOLE
       `validateUpdateListItem` call (gotcha #4) — every column lost, not just the missing one. The
       rule and the traps behind it live in `shared/optionalColumns.ts`, shared with the upload form.

       Degrading is SAFE AND VISIBLE here, unlike the submission stamp: with no marker the flow never
       fires and the files wait in the approval queue, where the Head of Unit can see and approve
       them. A failure that leaves work in a queue is not a silent one.

       The submission reference is new to this screen too — it was skipped while this was an
       admin-only tool writing to the approved side, and both halves of that reasoning are gone. */
    const targetTitle = targetListTitle();
    const canMark = await libraryHasColumns(
      context.spHttpClient,
      siteUrl,
      targetTitle,
      [BULK_IMPORT_COLUMN],
    );
    if (canMark) {
      formValues.push({ FieldName: BULK_IMPORT_COLUMN, FieldValue: "true" });
    }
    /* ⚠ ITS OWN ASK, never folded in with `BULK_IMPORT_COLUMN` or `REF_COLUMNS`. `libraryHasColumns`
       demands EVERY name it is given, so a library reconciled before today would fail the combined
       check and silently lose a feature that works there now. One ask, one degradation. */
    const canKeyword = await libraryHasColumns(
      context.spHttpClient,
      siteUrl,
      targetTitle,
      [KEYWORD_COLUMN],
    );
    if (canKeyword) {
      formValues.push({ FieldName: KEYWORD_COLUMN, FieldValue: keyword.trim() });
    }
    const canRef = await libraryHasColumns(
      context.spHttpClient,
      siteUrl,
      targetTitle,
      REF_COLUMNS,
    );
    /* One reference for the whole selection — this screen sends one destination, so a submission and
       its single batch are the same thing here. Generated per RUN, so pressing Upload again after a
       partial failure is a new submission rather than a silent merge into the last one. */
    const submissionRef = newReference("SUB", new Date());
    const batchRef = newReference("BAT", new Date());
    if (canRef) {
      formValues.push({ FieldName: REF_COLUMNS[0], FieldValue: submissionRef });
      formValues.push({ FieldName: REF_COLUMNS[1], FieldValue: batchRef });
    }
    /* ⚠ THE PER-FILE STAMP IS NOT PUSHED HERE, AND THAT IS THE WHOLE POINT. `formValues` is built
       ONCE for the entire selection — every file in a bulk run carries identical metadata — so a
       `SubmissionFileId` added here would be the SAME on all fifty documents, and every record row
       would resolve to whichever one happened to be found first. It is appended per file inside the
       loop instead.

       Checked SEPARATELY from `REF_COLUMNS` (see `libraryHasColumns`): a library reconciled before
       2026-08-27 has the pair and not the stamp, and must keep grouping exactly as it does today. */
    const canStamp = await libraryHasColumns(
      context.spHttpClient,
      siteUrl,
      targetTitle,
      [SUBMISSION_FILE_COLUMN],
    );

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

    /* WARN: NAMES ALREADY OFFERED IN THIS RUN, FED BACK IN AS IF THEY WERE TAKEN.
       `nextAvailableName` sees only the folder listing, so two files clashing in one run would each be
       offered the SAME free name and the dialog would print it twice - which is what the client caught
       on 2026-08-27. The listing cannot help: neither name has been written yet. Nothing is ever
       overwritten either way (`overwrite=false`), but the second file would be refused a second time,
       turning one round trip into two. */
    const reserved: string[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      // Files keep their own names — no composition on this screen (spec §4).
      /* An override REPLACES this file's own name — and the clash check below still runs on it, so a
         suggestion that has itself been taken in the meantime is skipped and offered the next one.
         Same convergence as the upload form. */
      const finalName = renameOverrides?.get(file) ?? file.name;
      // Read once, above the approved-side check that consumes it.
      const replaceApproved = replaceApprovedFiles?.has(file) === true;
      onPct(i, 0);
      onState(i, "uploading");
      setStatus(`Uploading ${i + 1} of ${files.length} — ${finalName}`);

      // Duplicate probe in the approval library.
      /* ⚠ NO OVERWRITE, AND NO PER-FILE PROMPT (2026-08-27). This branch used to ask "Are you sure you
         want to override this file?" and, on Yes, upload with `overwrite=true` — the ONLY path in the
         whole system that could destroy a pending document. Two things were wrong with it:

           1. It DESTROYED. A pending file has been approved by nobody, which was the original
              argument for allowing it, but it is still somebody's upload and there is no undo.
           2. It ASKED PER FILE, blocking a 50-file import on a modal. An admin importing historical
              documents is exactly the person who clicks Yes through the fifth prompt without reading
              it, and every Yes was a document gone.

         Now it SKIPS and offers a free name, decided ONCE at the end of the run — the same model as
         the upload form, so the two screens behave alike. The listing gives the names a suggestion
         needs, in the same single request the old `Exists` probe cost. */
      try {
        const listRes: SPHttpClientResponse = await context.spHttpClient.get(
          `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files?$select=Name&$top=5000`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (listRes.ok) {
          const rows = ((await listRes.json()).value ?? []) as Array<{
            Name?: string;
          }>;
          const here = rows
            .map((r) => r.Name ?? "")
            .filter((n) => n.length > 0);
          if (here.some((n) => n.toLowerCase() === finalName.toLowerCase())) {
            /* WARN: THE SUGGESTION MUST CLEAR BOTH LIBRARIES, and this branch was only clearing one.
               The file lands in the approval library NOW and is routed to the approved side LATER by
               auto-approve plus Auto-route, so a name free only here just moves the collision to
               `Copy file` - which renames it with SharePoint's own convention and tells nobody. The
               approved-side branch above has said this since 1.0.274.0; its sibling had not. */
            const taken = here
              .concat(
                approvedSideLeaf
                  ? (await approvedSideClashInfo(approvedSideLeaf, finalName))
                      .taken
                  : [],
              )
              .concat(reserved);
            const free = nextAvailableName(finalName, taken);
            if (free && free.toLowerCase() !== finalName.toLowerCase())
              reserved.push(free);
            out.push({
              name: finalName,
              outcome: "skipped",
              detail: "A document with this name is already in this folder.",
              nameClash: true,
              suggestedName:
                free && free.toLowerCase() !== finalName.toLowerCase()
                  ? free
                  : undefined,
            });
            onState(i, "skipped");
            continue;
          }
        }
      } catch {
        // Unreadable folder — proceed. `overwrite=false` on the write is the real guarantee, and an
        // unanswerable check must not become a new way to be blocked.
      }

      /* WARN: THE STAGING CLASH IS EVALUATED FIRST, AND THE ORDER IS THE CLIENT'S RULE.
         *"if the file is existing in document library and also exist in staging library, don't allow
         them to overwrite"* (2026-08-27). A name taken in BOTH libraries must come back as the
         RENAME-ONLY refusal, never as the replaceable one - the staging branch above `continue`s,
         so `approvedClash` is never set and the dialog cannot offer *Send for approval as a
         replacement* for a file whose real blocker is a pending draft. */
      /* ── Duplicate on the APPROVED side ────────────────────────────────────
         Checked FIRST, and it is a refusal rather than a prompt.

         SUPERSEDED 2026-08-27 — IT NOW OFFERS A FREE NAME TOO, matching the upload form (client:
         *"the popup doesn't show two files like a normal upload form"*, and earlier, of the form
         itself, *"why not auto rename the file for the Documents library version as well?"*). The
         earlier explain-only stance argued that nobody should be one click from filing a second copy
         of an approved record. THE DECIDING ARGUMENT IT MISSED IS THAT AUTO-ROUTE RENAMES IT ANYWAY:
         this screen writes into the APPROVAL library and stamps `BulkImport`, so auto-approve fires
         and Auto-route copies with `nameConflictBehavior: 2` — producing a SharePoint-renamed
         `TEST1.pdf` at routing time and telling nobody. A visible rename beats an invisible one.

         Nothing is ever replaced either way: the write is `overwrite=false` and an uploader holds
         only Read on the approved side.

         WARN: THE SUGGESTION MUST CLEAR BOTH LIBRARIES. The file lands in the approval library NOW
         and is routed to the approved side LATER, so a name free only in the approval library just
         moves the collision to Auto-route.

         ⚠ FAILS OPEN. An unanswerable check lets the upload through — which is exactly the behaviour
         before this existed. It is an extra guard, and a throttle must not become a new way to be
         blocked. `probeFolderByPath` already uses the OData parameter alias form, which is what
         keeps a deep path from returning 400 instead of 404 (gotcha #9). */
      if (approvedSideLeaf && !replaceApproved) {
        const approved = await approvedSideClashInfo(
          approvedSideLeaf,
          finalName,
        );
        if (approved.clash) {
          const taken = [finalName]
            .concat(approved.taken)
            .concat(await approvalFolderNames(folderId))
            .concat(reserved);
          const free = nextAvailableName(finalName, taken);
          if (free && free.toLowerCase() !== finalName.toLowerCase())
            reserved.push(free);
          out.push({
            name: finalName,
            outcome: "skipped",
            detail: `A document with this name is already filed in ${uploadingHc() ? "HC Documents" : "Documents"}. Nothing filed has been replaced — the original has already been approved, so check this is not the same document.`,
            nameClash: true,
            approvedClash: true,
            suggestedName:
              free && free.toLowerCase() !== finalName.toLowerCase()
                ? free
                : undefined,
          });
          onState(i, "skipped");
          continue;
        }
      }

      try {
        // HARDCODED FALSE, matching the upload form. Nothing on this screen may replace a document.
        const addUrl = `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files/Add(url='${encodeURIComponent(finalName)}',overwrite=false)?$select=ServerRelativeUrl`;

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
          /* ⚠ A BLOCKED REQUEST FIRST, BEFORE THE STATUS IS FORMATTED AT ALL (2026-08-27, seen live
             as `HTTP 0 — Network error during upload.`). This screen writes the file through XHR for
             the progress bar, and a request stopped by an extension resolves with status 0 rather
             than throwing — so `friendlyUploadError`, which matches what `fetch` THROWS, never saw
             it. The uploader got a raw status for the one failure that has nothing to do with their
             document, their permissions or the DMS.
             An abort is excluded: it is also status 0 and it is the uploader's own doing. */
          if (
            blockedBeforeNetwork(uploadRes.status) &&
            uploadRes.body !== ABORTED_BODY
          ) {
            console.error(
              "Upload blocked before the network:",
              finalName,
              folderId,
              uploadRes.body,
            );
            out.push({
              name: finalName,
              outcome: "failed",
              detail: BLOCKED_UPLOAD_MESSAGE,
            });
            onState(i, "failed");
            continue;
          }
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

          /* ── A CLASH THE LISTING ABOVE COULD NOT SEE ──────────────────────────
             ⚠ THE COMMON CASE FOR THIS SCREEN, not an edge one, and it is worse here than in the
             upload form. Bulk Upload is used by ordinary uploaders (since 2026-08-22), and Draft Item
             Security on the approval library is *"Only users who can approve items (and the author)"* —
             so an uploader's folder listing contains only their OWN pending files. Every pending file
             belonging to anybody else is invisible to the clash check, and only `Files/Add` knows.

             Without this the uploader gets SharePoint's raw *"A later version of this item has already
             been modified…"*, which describes versioning to someone who only chose a filename, and no
             rename is offered for the very case the offer exists for.

             Matched loosely and deliberately, exactly as in `Form.tsx`: a false positive only offers a
             rename that can be declined, while a false negative hides the fix behind a 500. */
          const lower = detail.toLowerCase();
          const looksLikeNameClash =
            lower.indexOf("later version") !== -1 ||
            lower.indexOf("already been modified") !== -1 ||
            lower.indexOf("already exists") !== -1 ||
            uploadRes.status === 409;
          if (looksLikeNameClash) {
            /* The name we now KNOW is taken, plus whatever the listing could see. The listing may have
               been trimmed to nothing — that is how we got here — so `finalName` is included
               explicitly, which guarantees a suggestion even when nothing was visible at all. A
               suggestion that is itself taken comes back through this same branch on the retry and is
               offered the next number. */
            let taken: string[] = [finalName].concat(reserved);
            try {
              const againRes: SPHttpClientResponse =
                await context.spHttpClient.get(
                  `${siteUrl}/_api/web/GetFolderById(guid'${folderId}')/Files?$select=Name&$top=5000`,
                  SPHttpClient.configurations.v1,
                  { headers: { Accept: "application/json;odata=nometadata" } },
                );
              if (againRes.ok) {
                const rows = ((await againRes.json()).value ?? []) as Array<{
                  Name?: string;
                }>;
                taken = taken.concat(
                  rows.map((r) => r.Name ?? "").filter((n) => n.length > 0),
                );
              }
            } catch {
              /* The single known-taken name is enough to produce an offer. */
            }
            if (approvedSideLeaf) {
              // Both libraries, same reason as the pre-check branches above.
              taken = taken.concat(
                (await approvedSideClashInfo(approvedSideLeaf, finalName))
                  .taken,
              );
            }
            const free = nextAvailableName(finalName, taken);
            if (free && free.toLowerCase() !== finalName.toLowerCase())
              reserved.push(free);
            out.push({
              name: finalName,
              outcome: "skipped",
              detail:
                "A document with this name is already in this folder. It may be a pending upload from " +
                "someone else that you cannot see until it is approved.",
              nameClash: true,
              suggestedName:
                free && free.toLowerCase() !== finalName.toLowerCase()
                  ? free
                  : undefined,
            });
            onState(i, "skipped");
            continue;
          }

          out.push({ name: finalName, outcome: "failed", detail });
          onState(i, "failed");
          continue;
        }
        const uploadJson = JSON.parse(uploadRes.body);
        const uploadedSru = uploadJson.ServerRelativeUrl;
        onPct(i, 100);

        /* `UniqueId` for the submission record, with a retry on the bare `Id` as INSURANCE: one
           unknown name fails the WHOLE `$select` (gotcha #11), and the cost of being wrong here is
           that nothing can be tagged. The record's `ItemUniqueId` is only ever diagnostic — the join
           is the stamp — so it is never worth an upload for. */
        const itemUrl = (select: string): string =>
          `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@f)/ListItemAllFields?$select=${select}&@f='${encodeServerRelativePath(uploadedSru)}'`;
        let itemRes: SPHttpClientResponse = await context.spHttpClient.get(
          itemUrl("Id,UniqueId"),
          SPHttpClient.configurations.v1,
        );
        if (!itemRes.ok) {
          itemRes = await context.spHttpClient.get(
            itemUrl("Id"),
            SPHttpClient.configurations.v1,
          );
        }
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

        /* One stamp per FILE, appended to a COPY of the shared values — see the note where `canStamp`
           is resolved. Mutating `formValues` itself would accumulate one stamp per file across the
           run and fail the call on a duplicate field name. */
        const fileStamp = canStamp ? newReference("SFI", new Date()) : "";
        const fileValues =
          fileStamp.length > 0
            ? formValues.concat([
                { FieldName: SUBMISSION_FILE_COLUMN, FieldValue: fileStamp },
              ])
            : formValues;

        const metaRes: SPHttpClientResponse = await context.spHttpClient.post(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(targetListTitle())}')/items(${item.Id})/validateUpdateListItem`,
          SPHttpClient.configurations.v1,
          {
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ formValues: fileValues }),
          },
        );
        if (!metaRes.ok) {
          /* ⚠ THE BODY IS WHERE SHAREPOINT NAMES THE CAUSE, and dropping it cost an hour on
             2026-08-22: a 400 here said only "tagging failed", and the real fault was that the file
             had been written to one library while this call addressed another. The three causes need
             opposite fixes — 404 is the wrong title, 403 is permissions on that list, 400 is a
             malformed payload or an unknown field name — so the status, the LIBRARY and the body all
             have to survive to the screen. Same lesson as gotcha #9, learned here for the third time. */
          const why = await metaRes.text().catch(() => "");
          out.push({
            name: finalName,
            outcome: "tagFailed",
            detail: `Uploaded, but tagging failed (HTTP ${metaRes.status}) against "${targetListTitle()}". ${why.slice(0, 200)}`,
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

        /* RECORD THE UPLOAD (2026-08-27). Only when the stamp went on: a row that cannot be joined
           back to its document would show on My Submissions as a permanent false "Deleted". Never
           fails the upload — the file is filed and tagged by this point, and `writeSubmissionRecord`
           cannot throw. Its result is ignored deliberately; it logs its own failure.

           The snapshot is the run's shared metadata, so a deleted row still shows what was imported.
           `Source` is `BulkUpload`, which is how a historical import is told apart from an uploader's
           own submission when somebody reads this list years from now. */
        if (fileStamp.length > 0) {
          const snapshot: Record<string, string> = {};
          for (const l of allSelections ?? []) snapshot[l.column] = l.label;
          snapshot["Document date"] = documentDate;
          // The LABEL, not the term id — the record is read by a person, and a bare GUID says nothing.
          snapshot.Confidentiality = labels.confLabel;
          snapshot.Remark = remark.trim();
          snapshot["Bulk import"] = "Yes";
          await writeSubmissionRecord(context.spHttpClient, siteUrl, {
            submissionRef: canRef ? submissionRef : "",
            batchRef: canRef ? batchRef : "",
            fileId: fileStamp,
            uniqueId:
              typeof item.UniqueId === "string" ? item.UniqueId : undefined,
            fileName: finalName,
            itemPath: uploadedSru,
            libraryTitle: targetListTitle(),
            uploadedBy: (context.pageContext.user.email ?? "").toLowerCase(),
            uploadedAt: new Date(),
            metadata: snapshot,
            source: "BulkUpload",
          });
        }

        out.push({ name: finalName, outcome: "uploaded" });
        onState(i, "done", uploadedSru);
      } catch (err) {
        console.error("Upload threw:", finalName, err);
        out.push({
          name: finalName,
          outcome: "failed",
          // See Form.tsx's identical use of friendlyUploadError for why: a raw "Failed to fetch"
          // is a browser extension or content blocker, not this document, this admin's
          // permissions, or the DMS — and this list is admin-only, but admins hit blockers too.
          detail: friendlyUploadError(err, "Unexpected error."),
        });
        onState(i, "failed");
      }
    }

    return { results: out };
  };

  /**
   * @param retry Files to send INSTEAD of the current selection, under forced names.
   *
   * Used by the clash dialog's Proceed. Passed explicitly rather than read from `picked`, because the
   * successful files were filtered out of it moments earlier and that setState has not landed —
   * reading state here would re-send everything that just worked.
   */
  const startUpload = async (retry?: {
    files: File[];
    overrides: Map<File, string>;
    replaceApproved?: Set<File>;
  }): Promise<void> => {
    // Before anything else, and re-read rather than trusted from mount: an administrator may have
    // paused uploads since this page was opened. Nothing is lost — the selection stays as it is.
    if (await readUploadPause()) {
      setPaused(true);
      showToast(UPLOAD_PAUSE_MESSAGE, "error");
      return;
    }
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
      (n, i) =>
        names.findIndex((o) => o.toLowerCase() === n.toLowerCase()) !== i,
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
    // A page open since before a structure change keeps filing into the OLD shape, because the
    // config is read once at mount (gotcha #10). Every such upload re-creates the two-shapes state
    // a migration just cleaned up, and it is invisible — the upload succeeds and lands in a folder
    // that looks perfectly reasonable. Seen live 2026-08-11 on the single-file form.
    //
    // `undefined` means the read failed, which proves nothing and must not block: refusing to
    // upload because a config read failed is worse than the risk it guards against.
    const freshChain = await freshChainFor(mode.key);
    if (
      freshChain &&
      chainSignature(freshChain) !== chainSignature(mode.chain ?? mode.levels)
    ) {
      showToast(
        "The folder structure changed while this page was open, so these files would be filed in " +
          "the wrong place. Please reload the page and try again — nothing has been uploaded.",
        "error",
      );
      return;
    }

    // A malformed chain must never route files to a partial path — they would land
    // one tier shallower than everything else in the unit, in a folder that exists
    // and looks right. Block and name the config row instead.
    const chainError = validateChain(mode.chain ?? mode.levels);
    if (chainError) {
      showToast(
        `The folder structure for this segment is not set up correctly: ${chainError.message} ` +
          `Ask an administrator to check the CRS Config mode row.`,
        "error",
      );
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

    const files = retry ? retry.files : picked.map((p) => p.file);
    const labels = {
      modeLabel: mode.label,
      termSetGuid: mode.termSetGuid,
      // The below-Unit path and its metadata, resolved once here and carried to the
      // runner. Passing the ordered segments rather than two named labels is what
      // lets the runner stay agnostic about how deep the chain is.
      tierSegments: buildOnDemandSegments(tierPlan().tiers, tierSelections())
        .segments,
      // Only APPLICABLE tiers write metadata: a unit with no subunits leaves SubUnit and
      // SubUnitTid empty rather than storing a value from another unit's list.
      tierFormValues: tierPlan().tiers.reduce(
        (acc: Array<{ FieldName: string; FieldValue: string }>, t, i) => {
          const opts = tierOptions(t, tierPlan().parents[i]);
          const opt = opts.find((o) => o.id === (tierValues[t.column] ?? ""));
          const col = t.labelCol ?? t.column;
          if (!col || !opt) return acc;
          // Derived, not flagged — a tier with a tidCol writes a text label+GUID
          // pair like the permissioned levels; one without writes a single
          // managed-metadata column. Keeps Year/Document Type byte-identical.
          if (t.tidCol) {
            acc.push({ FieldName: col, FieldValue: opt.label });
            acc.push({ FieldName: t.tidCol, FieldValue: opt.id });
          } else {
            acc.push({ FieldName: col, FieldValue: taxVal(opt.label, opt.id) });
          }
          return acc;
        },
        [],
      ),
      confLabel:
        options.confidentiality.find((o) => o.id === confidentiality)?.label ??
        "",
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
        file: f,
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
        prev
          ? prev.map((f, fi) => (fi === fileIndex ? { ...f, pct } : f))
          : prev,
      );
    };

    try {
      const outcome = await runUpload(
        files,
        selections,
        labels,
        onState,
        onPct,
        retry?.overrides,
        retry?.replaceApproved,
      );
      setStatus("");
      if (outcome.runError) {
        // Routing failed before any file uploaded — show every row as failed.
        setLive((prev) =>
          prev
            ? prev.map((f) => ({ ...f, state: "failed" as FileState }))
            : prev,
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

      /* ⚠ SUCCESS REMOVES, FAILURE STAYS — and this used to be all-or-nothing (fixed 2026-08-22,
         found on site). The clear was gated on `!anyProblem`, so ONE skipped file left the whole
         selection on screen — including every file that had already uploaded. Pressing Upload again
         then re-sent them, which is a replace prompt per file at best and, if someone clicks through,
         a document overwritten in the approval library.

         The rule is the one `uploadBatches.ts` already states for the batched form: an uploaded file
         leaves the list so it cannot be sent twice, and what remains on screen is exactly what still
         needs doing. That is what makes pressing Upload again a safe RETRY rather than a re-send.

         Matched by INDEX, not by name: `files` is `picked.map(p => p.file)` and `runUpload` pushes one
         result per file in order, so position is exact. Names are not — this screen keeps the original
         filenames, and two identical names in one selection would collapse to one match. A result
         missing for an index (a run that threw part-way) counts as NOT uploaded and stays. */
      /* ⚠ MATCHED BY FILE IDENTITY, not by index (changed 2026-08-27). `outcome.results` is aligned
         with the array actually SENT, which on a retry is just the clashing files — so an index into
         it means nothing to `picked`. Identity is right for both cases and removes the fragility the
         old comment above was defending against. A file with no result counts as NOT uploaded and
         stays, exactly as before. */
      const uploadedFiles = new Set<File>();
      outcome.results.forEach((r, i) => {
        if (r.outcome === "uploaded" && files[i]) uploadedFiles.add(files[i]);
      });
      setPicked((prev) => prev.filter((p) => !uploadedFiles.has(p.file)));

      /* Clashes with a free name to offer, decided once. Mapped back to the File that produced each
         result so the retry can force the name onto the right object. */
      const clashes: {
        file: File;
        from: string;
        to: string;
        approvedSide: boolean;
      }[] = [];
      const notes: { name: string; detail: string }[] = [];
      outcome.results.forEach((r, i) => {
        if (r.outcome !== "skipped" || !r.nameClash) return;
        if (r.suggestedName && files[i]) {
          clashes.push({
            file: files[i],
            from: r.name,
            to: r.suggestedName,
            approvedSide: r.approvedClash === true,
          });
        } else {
          // No free name could be found — still say so, or the dialog omits it silently.
          notes.push({ name: r.name, detail: r.detail ?? "" });
        }
      });
      setBulkClashes(clashes);
      setBulkClashNotes(notes);

      if (!anyProblem && okCount > 0) {
        // Same dialog as the single-file Form. A toast was easy to miss at the end of a long run, and
        // after 50 files the uploader needs an explicit "that worked" plus a way to go again without
        // reloading the page. The progress list is deliberately left on screen.
        setDoneOpen(true);
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

  /**
   * The button's entry point. Nothing but a responsiveness wrapper around `startUpload`.
   *
   * WARN: WITHOUT THIS THE BUTTON LOOKS DEAD FOR A SECOND OR TWO (client, 2026-08-27: *"I have to
   * click the upload button twice for bulk"*, then *"its a bit slow sometimes"* - the same
   * observation, correctly diagnosed the second time). `startUpload` performs TWO awaited network
   * reads before it reaches `setBusy(true)` - the upload-pause row and the fresh chain - so for that
   * whole window the button stayed enabled, still reading "Upload (1)", with nothing on screen
   * changed. The second click was then swallowed by `disabled={busy}` once it finally flipped, which
   * is why one press appeared to be ignored.
   *
   * THE UPLOAD FORM NEVER HAD THIS: it sets `setBusy(true)` BEFORE re-reading the pause. Same gap,
   * one screen fixed and the sibling not - the third time today a fix landed on one upload path and
   * not the other. Grep the other screen.
   *
   * A WRAPPER RATHER THAN A FLAG THREADED THROUGH THE BODY, deliberately: `startUpload` has more
   * than half a dozen early returns, and a flag set at the top must be cleared at every one of them.
   * Missing a single return would leave Upload permanently disabled - a worse bug than the one being
   * fixed. `finally` cannot miss one.
   */
  const handleUpload = async (retry?: {
    files: File[];
    overrides: Map<File, string>;
    replaceApproved?: Set<File>;
  }): Promise<void> => {
    setPreflight(true);
    try {
      await startUpload(retry);
    } finally {
      setPreflight(false);
    }
  };

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
        /* Same rule as Form.tsx, so the two upload pages carry the SAME heading — the client asked
           for this one "same like a normal upload form". If one is restyled, restyle both. */
        .dms-page-title { margin: 0 0 24px; font-size: 28px; font-weight: 700; color: #1b1b1b; }
        .dms-subtitle { margin: 0 0 16px; font-size: 14px; color: #666; }
        .dms-warn { display: flex; gap: 10px; align-items: flex-start; ${NOTICE_ATTENTION_CSS} border-radius: 6px; padding: 12px 14px; font-size: 13px; margin: 0 0 24px; max-width: 626px; }
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
        .dms-fp-x { flex: 0 0 auto; background: none; border: none; cursor: pointer; color: #605e5c; font-size: 14px; line-height: 1; padding: 2px 4px; }
        .dms-fp-x:hover { color: #d13438; }
        .dms-fp-xspacer { flex: 0 0 22px; }
        /* Overall progress header */
        .dms-overall-head { display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; color: #333; margin-bottom: 6px; }
        .dms-overall-track { height: 10px; border-radius: 6px; background: #ececec; overflow: hidden; }
        .dms-overall-fill { height: 100%; background: #0f6c3f; border-radius: 6px; transition: width .3s ease; }
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
        .dms-popup-btn.danger { background: #fff; color: #a4262c; border: 1px solid #a4262c; }
        .dms-popup-btn.danger:hover { background: #fdf3f4; }
        .dms-popup-btn.cancel { background: #fff; color: #0f6c3f; border: 1px solid #0f6c3f; }
        .dms-popup-btn.cancel:hover { background: #f0f6f2; }
        .dms-toast { position: fixed; top: 24px; right: 24px; z-index: 9999; min-width: 300px; max-width: 460px; padding: 14px 40px 14px 16px; border-radius: 6px; font-size: 13px; font-family: 'Segoe UI', sans-serif; box-shadow: 0 4px 16px rgba(0,0,0,.18); animation: dms-slidein .2s ease; }
        .dms-toast.error { background: #d13438; color: #fff; }
        .dms-toast.success { background: #0f6c3f; color: #fff; }
        .dms-toast.notice { background: #8a4b00; color: #fff; }
        .dms-toast-close { position: absolute; top: 10px; right: 12px; background: none; border: none; cursor: pointer; font-size: 16px; color: inherit; opacity: .7; line-height: 1; }
        .dms-toast-close:hover { opacity: 1; }
        @keyframes dms-slidein { from { transform: translateX(60px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        @media (max-width: 640px) {
          .dms-grid, .dms-grid-3 { grid-template-columns: 1fr; }
          .dms-toast { left: 12px; right: 12px; min-width: unset; top: 12px; }
        }
      `}</style>

      {/* ⚠ THE PAGE MUST HAVE ITS OWN TITLE WEB PART DELETED, or "Bulk Upload" appears TWICE — the
          same required deployment step the upload form needed (2026-09-03). Nothing in code can
          detect the duplicate: a text web part is not readable from inside a web part, so this will
          not fail, warn, or look wrong to any check. It renders twice until somebody opens the page.
          Client: *"Same like a normal upload form, add the title back into the form for bulk
          upload."* */}
      <h1 className="dms-page-title">Bulk Upload</h1>

      {/* Above the form, so it is read before any work is done. Amber, not red: nothing has failed
          and nothing the uploader did is wrong. The write-time re-check is what refuses. */}
      {paused ? (
        <div
          role="status"
          style={{
            margin: "0 0 16px",
            padding: "10px 14px",
            borderRadius: 4,
            ...NOTICE_ATTENTION,
            fontSize: 13,
            lineHeight: 1.5,
          }}
        >
          <strong>Uploads are paused.</strong> An administrator is reorganising
          the document folders. Nothing is wrong with your files — please try
          again shortly.
        </div>
      ) : undefined}
      <div className="dms-warn">
        {/* ⚠ INLINED, NOT PROVISIONED — the client's own `Icon.svg`, verbatim apart from the two
            attributes JSX needs to own (`aria-hidden`, and the sizing left to the markup). It is NOT
            referenced as a file: shipping an image asset in this package has failed to provision on
            this tenant TWICE, both times silently (the `+ New Folder` customizer, then the
            bulk-approve command set), so a src-referenced icon would render as a broken box with
            nothing explaining it. Same reason the bulk-approve command set uses a base64 data URI
            rather than a packaged file.
            `flexShrink: 0` because `.dms-warn` is a flex row: without it the icon is squeezed as the
            sentence grows. `fill` is the client's `#DA3B3B`, kept verbatim rather than switched to
            `currentColor` — it is deliberately a shade off the banner's `#a4262c` text in their
            design, and guessing otherwise would quietly redesign it. */}
        <svg
          width="35"
          height="35"
          viewBox="0 0 21 19"
          fill="none"
          aria-hidden="true"
          style={{ flexShrink: 0, marginTop: 1 }}
        >
          <path
            d="M11.2354 6.85311C11.2328 6.38367 10.8501 6.00526 10.3806 6.00789C9.9112 6.01053 9.53279 6.39322 9.53543 6.86265L10.3854 6.85788L11.2354 6.85311ZM9.55786 10.8563C9.5605 11.3258 9.94319 11.7042 10.4126 11.7015C10.8821 11.6989 11.2605 11.3162 11.2578 10.8468L10.4078 10.8516L9.55786 10.8563ZM11.2354 13.8479C11.2354 13.3784 10.8549 12.9979 10.3854 12.9979C9.91597 12.9979 9.53541 13.3784 9.53541 13.8479H10.3854H11.2354ZM9.53541 13.8579C9.53541 14.3273 9.91597 14.7079 10.3854 14.7079C10.8549 14.7079 11.2354 14.3273 11.2354 13.8579H10.3854H9.53541ZM8.65456 1.84753L7.91895 1.42165V1.42165L8.65456 1.84753ZM1.12346 14.8558L0.387845 14.4299H0.387845L1.12346 14.8558ZM19.6474 14.8558L18.9118 15.2817V15.2817L19.6474 14.8558ZM12.1163 1.84754L11.3807 2.27342V2.27342L12.1163 1.84754ZM10.3854 6.85788L9.53543 6.86265L9.55786 10.8563L10.4078 10.8516L11.2578 10.8468L11.2354 6.85311L10.3854 6.85788ZM10.3854 13.8479H9.53541V13.8579H10.3854H11.2354V13.8479H10.3854ZM8.65456 1.84753L7.91895 1.42165L0.387845 14.4299L1.12346 14.8558L1.85907 15.2817L9.39017 2.27342L8.65456 1.84753ZM2.85431 17.8579V18.7079H17.9165V17.8579V17.0079H2.85431V17.8579ZM19.6474 14.8558L20.383 14.4299L12.8519 1.42166L12.1163 1.84754L11.3807 2.27342L18.9118 15.2817L19.6474 14.8558ZM17.9165 17.8579V18.7079C20.112 18.7079 21.483 16.3299 20.383 14.4299L19.6474 14.8558L18.9118 15.2817C19.3556 16.0484 18.8024 17.0079 17.9165 17.0079V17.8579ZM1.12346 14.8558L0.387845 14.4299C-0.712152 16.3299 0.658856 18.7079 2.85431 18.7079V17.8579V17.0079C1.96843 17.0079 1.41521 16.0484 1.85907 15.2817L1.12346 14.8558ZM8.65456 1.84753L9.39017 2.27342C9.83311 1.50834 10.9377 1.50834 11.3807 2.27342L12.1163 1.84754L12.8519 1.42166C11.7542 -0.474401 9.01667 -0.474411 7.91895 1.42165L8.65456 1.84753Z"
            fill="#DA3B3B"
          />
        </svg>
        {/* Client's own copy, verbatim, 2026-09-03. Shorter than what it replaced, which spelled out
            that the files are visible to everyone with access to the destination folder as soon as
            they upload — that consequence is now unstated on screen. Their call; it is the ONE thing
            an admin might not infer from "without going through approval". */}
        <span>
          <strong>Temporary Tool.</strong> Upload files directly to the{" "}
          <strong>Documents</strong> library without going through approval.
          Please make sure the file names are correct before uploading.
        </span>
      </div>

      {/* ⚠ THE 50-FILE CEILING AND "same folder, same details" CAME OFF 2026-08-30 at the client's
          request. Both are still TRUE and still enforced — `MAX_FILES` refuses the 51st file with a
          toast, and this screen has one destination and one metadata set by construction. Only the
          statement of them is gone, so an uploader now meets the cap at the moment they exceed it
          rather than before they start. */}
      <p className="dms-subtitle">
        All fields marked <strong>*</strong> are required.
      </p>

      {/* ── Documents Folder Information ─────────────────────────────────── */}
      {/* Deliberately BEFORE Documents Details in SOURCE order, not reordered
          with CSS: tab order follows the DOM. Mirrors Form.tsx. */}
      <div className="dms-section">
        <p className="dms-section-title">Documents Folder Information</p>

        {deptLoading ? (
          <p className="dms-dept-loading">Loading your access&hellip;</p>
        ) : !privileged && validPaths.length === 0 ? (
          <div className="dms-dept-error">
            {awaitingFolders ? (
              /* Access is correct; the folders were never created. The membership
                 wording below would send the administrator to check groups that are
                 already right. Both fixes are named in order — a unit with no
                 abbreviation row is skipped by every reconciliation run, so
                 "re-run reconciliation" alone is wrong half the time. */
              <>
                Your unit&apos;s folders haven&apos;t been created yet. Your CRS
                administrator needs to give every unit an abbreviation in the
                DMS Term Abbreviation list, then run folder reconciliation.
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
                  name="bulkSideToggle"
                  checked={active}
                  disabled={busy}
                  onChange={() => switchMode(offerable[0].key)}
                />
                {/* ⚠ HARDCODED to "Group Led Project", client's explicit instruction
                    (2026-09-02) — mirrors Form.tsx's identical change; see that file's comment
                    for why this will need revisiting once a second Project-category segment
                    exists. Keep the two upload screens' wording in sync deliberately — this
                    project's own history is full of the two drifting apart. */}
                {side === "BusinessSegment"
                  ? "Business Segment"
                  : "Group Led Project"}
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
                {side === "Project" ? "Project" : "Segment"}{" "}
                <em className="req">*</em>
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
                    (levelChoices[i] ?? []).find((o) => o.id === levelValues[i])
                      ?.label,
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

          {/* The below-Unit chain, in path order — third-width, so tiers flow onto
              the row the deepest permissioned level starts. Unchanged on a site
              that has configured none. */}
          {/* Only the tiers that apply here — a unit with no subunits shows no SubUnit
              dropdown at all rather than an empty required one. */}
          {tierPlan().tiers.map((t, i) => {
            const parent = tierPlan().parents[i];
            return renderSelect(
              t.label,
              true,
              tierValues[t.column] ?? "",
              (v) => setTierValues((prev) => ({ ...prev, [t.column]: v })),
              tierOptions(t, parent),
              // A cascading tier stays disabled until the tier above it is chosen.
              busy || (!(t.termSet ?? "").trim() && !parent),
            );
          })}

          {/* Graceful empty-state: a segment whose term set has no child terms yet
              (the non-GHO Head Offices before their Department/Unit trees are
              added) would otherwise show a dead "Select …" dropdown. */}
          {(() => {
            const md = activeMode();
            if (
              !md ||
              md.levels.length === 0 ||
              deptLoading ||
              levelChoices.length === 0
            )
              return null;
            for (let i = 0; i < md.levels.length; i++) {
              const parentChosen = i === 0 || !!levelValues[i - 1];
              if (parentChosen && (levelChoices[i]?.length ?? 0) === 0) {
                return (
                  <div
                    key="dms-empty-level"
                    style={{
                      gridColumn: "1 / -1",
                      padding: "8px 12px",
                      ...NOTICE_ATTENTION,
                      borderRadius: 4,
                      fontSize: 13,
                    }}
                  >
                    No {md.levels[i].label.toLowerCase()} options are configured
                    for this segment yet — ask your administrator to add them in
                    the term store before uploading here.
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
              onChange={(e) => {
                /* Same guard as the upload form (client QA, 2026-08-30). Stripped rather than
                   refused, and the removal is SAID - a character vanishing with no explanation is
                   how a field comes to feel broken. */
                const clean = stripBlockedChars(e.target.value);
                setRemarkBlocked(
                  clean !== e.target.value
                    ? blockedCharsMessage(e.target.value)
                    : undefined,
                );
                setRemark(clean);
              }}
            />
            {remarkBlocked ? (
              <small className="dms-err">{remarkBlocked}</small>
            ) : (
              <small>Max. 250 characters</small>
            )}
          </label>

          {/* Keyword (client, 2026-09-04). Spans the row, as Remark does, and optional for the same
              reason it is optional on the upload form: it is a findability aid, not a property of the
              document, so requiring it would block an import of files nobody has words for. */}
          <label className="dms-field" style={{ gridColumn: "1 / -1" }}>
            <span>Keyword</span>
            <input
              type="text"
              value={keyword}
              maxLength={255}
              disabled={busy}
              placeholder="Words to help find these documents later"
              onChange={(e) => {
                /* The same strip-and-say guard as Remark above — a character vanishing with no
                   explanation is how a field comes to feel broken. */
                const clean = stripBlockedChars(e.target.value);
                setKeywordBlocked(
                  clean !== e.target.value
                    ? blockedCharsMessage(e.target.value)
                    : undefined,
                );
                setKeyword(clean);
              }}
            />
            {keywordBlocked ? (
              <small className="dms-err">{keywordBlocked}</small>
            ) : (
              <small>Optional &middot; searchable from the home page</small>
            )}
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
                    {/* WARN: DE-SELECT, NOT DELETE. See `dropFromRun`, and `sru` on LiveFile for why
                        a server-side delete must not come back here.

                        Offered ONLY on a row that was NOT uploaded. `done` is excluded because the row
                        is the only record that the file went, and `tagFailed` IS offered because that
                        file stays in the selection and would otherwise be re-sent on the next press.
                        Held entirely while a run is in flight. */}
                    {!busy &&
                    (f.state === "skipped" ||
                      f.state === "failed" ||
                      f.state === "tagFailed") ? (
                      <button
                        type="button"
                        className="dms-fp-x"
                        aria-label={`Remove ${f.name} from the list`}
                        title="Remove from the list (nothing is deleted)"
                        onClick={() => dropFromRun(fi)}
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
            onDragOver={(e) => {
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              addFiles(e.dataTransfer?.files ?? null);
            }}
          >
            {/* Inlined rather than imported: an <img> would need an asset loader
                and a second network request for a 20-line glyph. */}
            <svg
              className="dms-dropzone-icon"
              viewBox="0 0 24 24"
              aria-hidden="true"
            >
              <path
                d="M12 3v10m0 0 4-4m-4 4-4-4"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
              <path
                d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
              />
            </svg>
            <span>
              <span className="dms-link">Choose multiple documents</span> or
              drop them here
            </span>
            <span className="hint">Max. {MAX_FILES} files.</span>
          </div>
        ) : (
          // The list itself is now ALSO a drop target, not only the empty dropzone above —
          // dragging more files onto an already-picked list previously did nothing (the
          // browser just navigated to the file), which read as "the system didn't allow it".
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              addFiles(e.dataTransfer?.files ?? null);
            }}
            style={
              dragOver
                ? {
                    outline: "2px dashed #0f6c3f",
                    outlineOffset: 4,
                    borderRadius: 6,
                  }
                : undefined
            }
          >
            <div className="dms-selbar">
              <span>
                {picked.length} document{picked.length === 1 ? "" : "s"}{" "}
                selected
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
              {picked.map((p) => {
                /* Marked on the ROW as well as in the toast: a toast is gone in seconds and a
                   fifty-row list is not, so an admin scrolling back has no way to find which two
                   files the message was about. Both rows are marked — which of them is the mistake
                   is not knowable here, the same reasoning as the abbreviation sibling check. */
                const duplicate =
                  picked.filter(
                    (q) =>
                      q.file.name.toLowerCase() === p.file.name.toLowerCase(),
                  ).length > 1;
                return (
                  <div className="dms-filerow" key={p.key}>
                    <span className="fname" title={p.file.name}>
                      {p.file.name}
                    </span>
                    {duplicate && (
                      <span
                        style={{
                          fontSize: 11,
                          color: "#8a4b00",
                          whiteSpace: "nowrap",
                        }}
                        title="Another selected file has this name. All of them can be uploaded — the later ones are offered a free name — but check this is not the same document picked twice."
                      >
                        same name
                      </span>
                    )}
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
                );
              })}
            </div>
          </div>
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
                      intended strictly for use within the Group, on a
                      need-to-know basis.
                    </dd>
                    <dt>Restricted</dt>
                    <dd>
                      This applies to business information that may be disclosed
                      to external parties only if a non-disclosure agreement has
                      been signed.
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
              {/* Hidden rather than greyed out, as on the upload form. On a site with no HC
                  libraries nothing is filtered and the level stays an ordinary label. */}
              {(() => {
                const ctx = hcCtx();
                const keep = new Set(
                  selectableLevels(
                    options.confidentiality.map((o) => o.label),
                    ctx,
                  ).map((l) => l.trim().toLowerCase()),
                );
                return options.confidentiality
                  .filter((o) => keep.has((o.label ?? "").trim().toLowerCase()))
                  .map((o) => (
                    <option key={o.id} value={o.id} title={o.label}>
                      {o.label}
                    </option>
                  ));
              })()}
            </select>
          </div>

          {/* Offered only for the levels LISTED by `legallyPrivilegedFor` in DMS
              Config (a list since 2026-08-19; one value behaves as before). Unset
              means never offered. The value is re-derived at upload time rather
              than trusted from here, because hiding the control does not clear the
              state behind it. */}
          {offersLegalPrivilege(
            options.confidentiality.find((o) => o.id === confidentiality)
              ?.label ?? "",
            settings.legallyPrivilegedFor,
          ) && (
            // Wrapped together so the icon sits tight beside the checkbox label instead of
            // inheriting `.dms-detail-row`'s 24px row gap — the two were previously separate
            // flex children of that row, which is what put daylight between them.
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                flex: "0 0 auto",
                marginBottom: 16,
              }}
            >
              <label className="dms-lp" style={{ margin: 0 }}>
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
          disabled={busy || preflight || deptLoading || picked.length === 0}
        >
          {/* `busy` first: once the run starts both are true, and "Uploading" is the truer word. */}
          {busy
            ? "Uploading…"
            : preflight
              ? "Checking…"
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
                {resultCounts.skipped > 0 &&
                  `, ${resultCounts.skipped} skipped`}
                {resultCounts.failed > 0 && `, ${resultCounts.failed} failed`}
                {resultCounts.tagFailed > 0 &&
                  `, ${resultCounts.tagFailed} uploaded without tags`}
                .
                {resultCounts.tagFailed > 0 &&
                  " Files listed as “uploaded, not tagged” are already in Documents — fix their metadata in the library rather than re-uploading."}
              </p>
              {problemRows.map((r, i) => (
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
      )}

      {/* ── Name clashes, decided ONCE after the run ─────────────────────────────
          Replaces the per-file "Replace Existing File" prompt and its `overwrite=true`. Same shape as
          the upload form's dialog on purpose: an admin who has met one should recognise the other.
          Icon is RED (`#FF4646`, client's reference graphic 2026-09-03), matching the same swap made
          on the upload form's identical dialog — was amber (`#FF952A`) on both. */}
      {(bulkClashes.length > 0 || bulkClashNotes.length > 0) && (
        <div className="dms-popup-overlay" role="dialog" aria-modal="true">
          <div className="dms-popup" style={{ maxWidth: 520 }}>
            <svg
              className="dms-popup-svg"
              viewBox="0 0 184 184"
              fill="none"
              xmlns="http://www.w3.org/2000/svg"
            >
              <circle
                opacity="0.3"
                cx="91.9999"
                cy="92"
                r="75.4872"
                fill="#FF4646"
              />
              <circle cx="92" cy="92" r="92" fill="#FF4646" fillOpacity="0.2" />
              <circle
                cx="92.0003"
                cy="91.9998"
                r="61.3333"
                fill="white"
                stroke="#FF4646"
                strokeWidth="3"
              />
              <path
                d="M93 66L93 100"
                stroke="#FF4646"
                strokeWidth="10"
                strokeLinecap="round"
              />
              <path
                d="M93 116.804L93 118"
                stroke="#FF4646"
                strokeWidth="10"
                strokeLinecap="round"
              />
            </svg>
            <p className="dms-popup-title" style={{ color: "#a4262c" }}>
              {bulkClashes.length + bulkClashNotes.length === 1
                ? "A document with that name already exists"
                : `${bulkClashes.length + bulkClashNotes.length} documents with those names already exist`}
            </p>
            <p className="dms-popup-msg" style={{ marginBottom: 16 }}>
              {bulkClashes.length + bulkClashNotes.length === 1
                ? "It was"
                : "They were"}{" "}
              not uploaded. Nothing already filed has been changed or replaced.
            </p>
            <div
              style={{
                background: "#faf9f8",
                border: "1px solid #edebe9",
                borderRadius: 8,
                padding: "12px 14px",
                marginBottom: 20,
                maxHeight: "40vh",
                overflowY: "auto",
                textAlign: "left",
              }}
            >
              {bulkClashes.map((c) => (
                <div
                  key={c.from + String(c.file.size)}
                  style={{
                    fontSize: 13,
                    lineHeight: 1.7,
                    wordBreak: "break-word",
                    marginBottom: 8,
                  }}
                >
                  <span
                    style={{ color: "#605e5c", textDecoration: "line-through" }}
                  >
                    {c.from}
                  </span>
                  <br />
                  <strong style={{ color: "#0f6c3f" }}>{c.to}</strong>
                  {c.approvedSide && (
                    <>
                      <br />
                      <span style={{ color: "#8a4b00", fontSize: 12 }}>
                        The original has already been approved and filed — check
                        this is not the same document.
                      </span>
                    </>
                  )}
                </div>
              ))}
              {bulkClashNotes.map((x) => (
                <div
                  key={`note-${x.name}`}
                  style={{
                    fontSize: 13,
                    lineHeight: 1.7,
                    wordBreak: "break-word",
                    marginBottom: 8,
                  }}
                >
                  <strong style={{ color: "#8a4b00" }}>{x.name}</strong>
                  <br />
                  <span style={{ color: "#605e5c", fontSize: 12 }}>
                    {x.detail}
                  </span>
                </div>
              ))}
            </div>
            {/* WARN: THE SAME LAYOUT AS THE UPLOAD FORM, INLINE AND WRAPPING (client,
                2026-08-27: the bulk popup did not match the form). `dms-popup-actions` has no
                flex-wrap, so THREE buttons - the normal case once a document is already filed on
                the approved side - were squeezed onto one row here while the form let the third
                drop to a second line. Two screens asking the same question must not look like
                two different systems. */}
            <div
              style={{
                display: "flex",
                gap: 10,
                justifyContent: "center",
                flexWrap: "wrap",
              }}
            >
              {/* No Proceed when nothing can be renamed — Close is then the only honest action. */}
              {bulkClashes.length > 0 && (
                <button
                  className="dms-popup-btn confirm"
                  onClick={() => {
                    /* Retry ONLY these files, under the new names. The list is passed explicitly rather
                     than read from `picked`: the successful files were removed from it a moment ago
                     and that setState has not landed, so reading it here would re-send them. */
                    const overrides = new Map<File, string>();
                    const retryFiles: File[] = [];
                    for (const c of bulkClashes) {
                      overrides.set(c.file, c.to);
                      retryFiles.push(c.file);
                    }
                    setBulkClashes([]);
                    setBulkClashNotes([]);
                    handleUpload({ files: retryFiles, overrides }).catch(
                      () => undefined,
                    );
                  }}
                >
                  {bulkClashes.length === 1
                    ? "Upload with the new name"
                    : "Upload with the new names"}
                </button>
              )}
              {/* FILE A REPLACEMENT (client's revised brief, 2026-08-27). Offered ONLY for a document
                  already filed on the APPROVED side. A clash in the approval library is a PENDING
                  draft, which the client ruled may never be overwritten by anyone, so those stay
                  rename-only.

                  WARN: NO `overwrite` FLAG. The file goes to the approver under the same name; the
                  filed copy is replaced by the ROUTING flow, and Auto-route currently copies with a
                  NEW NAME on conflict - so without the flow change this does not replace anything. */}
              {bulkClashes.some((c) => c.approvedSide) && (
                <button
                  className="dms-popup-btn danger"
                  onClick={() => {
                    const targets = bulkClashes.filter((c) => c.approvedSide);
                    const one = targets.length === 1;
                    /* ⚠ THIS USED TO SAY "sent for approval … replaced only once an approver approves",
                     copied from the upload form where it IS true. It is not true here, and the
                     difference is the whole risk: Bulk Upload stamps `BulkImport`, the auto-approve
                     flow fires within a minute, and Auto-route replaces the filed document. NOBODY
                     REVIEWS IT. The old wording told an admin there was a human gate between their
                     click and an approved record being overwritten, and there is none.

                     Corrected 2026-08-28, after the client described this path in their own words —
                     *"Bulk upload is allowed to overwrite Document Library"*, which is what actually
                     happens. The behaviour was already right; only the description was wrong. */
                    const ok = window.confirm(
                      `Replace the ${one ? "document" : `${targets.length} documents`} already filed in ` +
                        `${uploadingHc() ? "HC Documents" : "Documents"}?\n\n` +
                        `${one ? "It is" : "They are"} replaced AUTOMATICALLY — bulk imports skip ` +
                        `approval, so nobody reviews this first.\n\n` +
                        /* ⚠ NOT "stays in version history", which is what this said until 2026-08-31.
                         Auto-route's Copy file DELETES the destination item and creates a new one,
                         so version history never sees the old content — proven on site. The recycle
                         bin is the only recovery, and it keeps the ORIGINAL uploader's name. */
                        `The ${one ? "document it replaces is" : "documents they replace are"} moved ` +
                        `to the site recycle bin, restorable for 93 days.`,
                    );
                    if (!ok) return;
                    const replaceApproved = new Set<File>();
                    const retryFiles: File[] = [];
                    for (const c of targets) {
                      replaceApproved.add(c.file);
                      retryFiles.push(c.file);
                    }
                    setBulkClashes([]);
                    setBulkClashNotes([]);
                    // No renames: these keep their own names, which is the point.
                    handleUpload({
                      files: retryFiles,
                      overrides: new Map<File, string>(),
                      replaceApproved,
                    }).catch(() => undefined);
                  }}
                >
                  {/* "Replace", not "Send for approval" — see the confirm above. A bulk import is
                    auto-approved, so this button IS the replacement, not a request for one. */}
                  {bulkClashes.filter((c) => c.approvedSide).length === 1
                    ? "Replace the filed document"
                    : "Replace the filed documents"}
                </button>
              )}
              <button
                className="dms-popup-btn cancel"
                onClick={() => {
                  setBulkClashes([]);
                  setBulkClashNotes([]);
                }}
              >
                {/* The form's own wording, so the two dialogs read alike: the button says what the
                    uploader is choosing to do instead, not merely that it closes. */}
                {bulkClashes.length === 0
                  ? "Close"
                  : `Cancel, I’ll rename ${bulkClashes.length === 1 ? "it" : "them"}`}
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
              <circle
                opacity="0.3"
                cx="92.0001"
                cy="92"
                r="75.4872"
                fill="#14C7A5"
              />
              <circle cx="92" cy="92" r="92" fill="#14C7A5" fillOpacity="0.2" />
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
