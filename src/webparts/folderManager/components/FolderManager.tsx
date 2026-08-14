import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { searchSiteGroups, fetchAllSiteGroups, getGroupMembers, addGroupMember } from "../../../shared/spGroups";
import { ensureSiteEntryGroup } from "../../../shared/siteEntryGroup";
import { findSiteEntryGroup, siteEntryGroupTitle, isForbiddenPageTarget, normalizeRoleValue } from "../../../shared/groupMapModel";
import { EVENT } from "../../../shared/auditLog";
import { cachedListTitle, LIST_SUFFIX, libApiTitle } from "../../../shared/naming";
// One parser for the `#tab=` deep link, shared with the CRS Settings page that writes it — the two
// halves of one contract, so they cannot drift.
import { tabFromHash } from "../../../shared/adminPages";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { IFolderManagerProps } from "./IFolderManagerProps";
import {
  loadFolderMapRows,
  FolderMapRow,
  resolveFolderByPath,
  probeFolderByPath,
  probeFolderById,
  writeFolderMapping,
  updateFolderMapping,
  renameFolder,
  deleteFolderMapRow,
  ensureFolder,
  encodeServerRelativePath,
} from "../../../shared/dmsFolderMap";
import { sanitizeFolderSegment, parseLevels, parseReconModes, RawModeRow, Level } from "../../../shared/formModel";
import { gridPlan, isPermissioned, splitChain, validateChain } from "../../../shared/folderChain";
import {
  abbrevListTitle,
  AbbrevCollision,
  AbbrevRow,
  AbbrevTarget,
  OrphanAbbrevRow,
  UnclaimedTerm,
  buildAbbrevIndex,
  findCollisions,
  lookupAbbrev,
  planOrphanRepairs,
} from "../../../shared/folderAbbreviation";
import { FULL_NAME_COLUMN_TITLE, pickFullNameField, SpFieldLite } from "../../../shared/folderFullName";
import AbbreviationManager from "./AbbreviationManager";
// The three Folder Structure screens are MOUNTED here, not copied: they keep living in the
// userAccess web part so the `Folder Structure` page stays working for any site that already has it
// on a page (spec `2026-08-12-term-abbreviation-page-design.md` §6). Two copies of a screen that
// rewrites `Levels` and creates columns in both libraries is exactly the drift this import avoids.
import StructureManager from "../../userAccess/components/StructureManager";
import SubtreeMigrator from "../../userAccess/components/SubtreeMigrator";
import SegmentCreator from "../../userAccess/components/SegmentCreator";

// A "mode" is a top-level container folder under the library root. These used to
// be hardcoded (Departments / Projects); they are now discovered dynamically so
// the tool works with the multi-segment model (Group Head Office, Group Upstream
// Operations, …) or any future top-level folder naming.
type Mode      = string;
/**
 * The two libraries, as a LOGICAL key — not necessarily what either is called on the site.
 *
 * "Staging" is kept as the key rather than renamed to "Approval Document" because it is also
 * the value stored in the Group Map's `Target` column on every library-scope row, and it
 * indexes LIBRARY_ROLES and the progress feeds. Renaming the key would silently orphan that
 * stored data. Translate to the real title at the API boundary instead — libApiTitle().
 */
type LibTarget = "Staging" | "Documents";

// libApiTitle — the title SharePoint actually answers to, for a logical library key — moved to
// shared/naming.ts on 2026-08-14. It was private here while this was the only screen building a URL
// from a LibTarget; StagingAccess turned out to need it too and, lacking it, had been asking for a
// library called "Staging" that no longer exists.
// Folder-derived content type that carries the Full Name column, so the value shows in
// the details pane. Matched by NAME because its id differs per library. Optional: a site
// without it still provisions normally, it just shows nothing in the pane.
//
// Two names are accepted, newest first, for the same reason list titles are probed per
// suffix rather than from one global prefix: the client renames every DMS-named artefact
// to CRS by hand, one at a time, so a site legitimately spends time with CRS lists and a
// DMS-named content type. Probing both is what keeps a half-renamed site working — and
// unlike a list, a miss here is silent, because an absent content type is a supported
// state rather than an error.
const FOLDER_CONTENT_TYPE_CANDIDATES = ["CRS Folder", "DMS Folder"];
// Whichever candidate this library actually has, recorded by loadFolderContentTypeId.
// Only ever read for log lines, so the fallback to the preferred name is cosmetic.
let resolvedFolderCtName: string | undefined;
// The top-level tabs. The two library tabs drive the folder tree; Reconciliation is the
// term-store-driven provisioner (create + lock + map).
//
// The four ACCESS surfaces (folder, site, library, page) left this web part on 2026-08-07 for
// their own "User Access" page — spec `2026-08-07-access-webpart-split-design.md`. They were
// sub-tabs here, which buried the most frequent task (a person joins or moves) one level below
// the least frequent ones. What remains is structure and provisioning only.
/**
 * `Staging` and `Documents` are no longer offered in the tab bar (client, 2026-08-12: "I am honestly
 * not using it") — reconciliation and the Folder Access page cover what that manual folder tree did.
 * They stay in the union because the tree's render branch and its helpers are still here; deleting
 * that code is a separate cleanup, and keeping the branch type-reachable leaves the file compiling
 * and lint-clean in the meantime.
 */
type Tab       = LibTarget | "Reconciliation" | "Abbreviations" | "Levels" | "Migrate" | "NewSegment";

/**
 * Tabs the CRS Settings landing page may deep-link to, by slug.
 *
 * Only the three it actually links are listed. A slug is a PUBLIC name once shipped — the landing page
 * writes it into a URL an admin may bookmark — so this map is the contract, and renaming a Tab value
 * must not silently break it. Anything unrecognised falls through to the default tab, never to a blank
 * screen.
 */
const DEEP_LINK_TABS: Record<string, Tab> = {
  abbreviations: "Abbreviations",
  structure: "Levels",
  reconciliation: "Reconciliation",
  migrate: "Migrate",
  newsegment: "NewSegment",
};

/** The tab named by the URL hash, or Term Abbreviations. */
function tabFromDeepLink(): Tab {
  // `window` is always present in a web part, but a guard costs nothing and keeps this callable from
  // a test later.
  const hash = typeof window === "undefined" ? "" : window.location.hash;
  return DEEP_LINK_TABS[tabFromHash(hash)] ?? "Abbreviations";
}

// Reconciliation "modes" — mirror Form.tsx / the retired Reconciliation web part.
// Each maps a term set to the segment container folder its terms live under.
// Pilot slice: the four Head Office segments (2026); add the 2027 segments once
// their term sets are onboarded (term-set GUID + its container folder name).
type ReconMode = { key: string; termSetGuid: string; stagingFolder: string };
const RECON_MODES: ReconMode[] = [
  { key: "gho", termSetGuid: "08dd94cb-f76c-431c-9b37-e9c98f739ffc", stagingFolder: "Group Head Office" },
  // PLACEHOLDER — Upstream Malaysia is NOT onboarded (client scope 2026-07-29 is the other
  // three head offices). This GUID pre-dates the 2026-07-29 term-set rebuild and is stale.
  // Inert: with no `mode` row in DMS Config the segment is never offered. Replace the GUID
  // when the client creates the term set — nothing else needs to change.
  { key: "upstream_my_ho", termSetGuid: "16a52947-57a3-4217-9a49-b48cb8b0dd31", stagingFolder: "Upstream Malaysia Head Office" },
  { key: "minamas_ho", termSetGuid: "9ad00b00-a43c-4a8b-a39a-d0efa89ba706", stagingFolder: "Minamas Head Office" },
  { key: "nbpol_ho", termSetGuid: "77c3993b-0c3c-4a18-89d9-d69209886322", stagingFolder: "NBPOL Head Office" },
];
type TermLite = { id: string; label: string };

// Year/Period and Document Type term sets. Under each leaf (unit) folder the
// provisioner pre-creates the full Year × Document Type grid (path order matches
// the upload form: Unit / Year / Document Type). These inherit the unit's ACL.
// Offline fallback only — the grid term sets are read at runtime from the DMS Config
// `setting` rows (termSet_yearPeriod / termSet_documentType) via loadReconGridTermSets,
// so a different tenant needs no code edit. These GUIDs are the sandbox values.
const YEAR_TERMSET    = "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf";
const DOCTYPE_TERMSET = "866c5754-258e-401f-8685-03d20ae59b1d";

/**
 * How much of the Year × Document Type grid to pre-create under each unit folder.
 * Set live from DMS Config (`recon_gridMode` setting row) — no redeploy to change.
 *
 *   off         — create nothing. The upload form ensure-creates Year/Document Type on
 *                 first upload, and the Auto-route flow creates them in Documents as
 *                 approved files land. Viewers then see only folders that hold documents.
 *   currentYear — pre-create the current year's row only.
 *   full        — every year × every document type (the original behaviour).
 *
 * Why `off` is the default: the grid is ~98% of a full run. With 3 years and 20 document
 * types it is 63 folders per unit per library — 16,758 operations for 133 units, turning a
 * ~6 minute run into ~4h20m, to produce folders that are overwhelmingly empty and that
 * make browsing NOISIER for viewers. The structural folders that actually carry
 * permissions are only ~360 operations.
 */
type GridMode = "off" | "currentYear" | "full";
const DEFAULT_GRID_MODE: GridMode = "off";

// One folder the provisioner will ensure exists + lock. termGuid is null for the
// segment container folder (not a term); term folders (department, unit, …) carry
// their GUID so Staging can be mapped for rename-proof routing. isLeaf marks the
// deepest terms (upload targets) that get the Year × Document Type grid beneath.
// assignTerm is the term used to look up DMS Group Map rows for THIS folder's tier:
// the term-set GUID for the segment container, the term GUID for dept/unit folders.
// fullName is the RAW term label for this one folder, written to the Full Name column
// so the abbreviated folder still says what it is. It is deliberately not `label`:
// label is a breadcrumb ("GHO > Group Finance > Treasury") built for the log and the
// map row Title, and putting a breadcrumb in a per-folder column would repeat the
// whole path on every row.
// ancestorTerms is every term GUID ABOVE this folder on its own branch, outermost
// first. It exists for departmental fan-out: a Group Map row on a non-leaf term
// grants its role on every folder beneath it, so each folder needs to know which
// terms could be reaching down onto it. Empty for the top tier.
// termSetGuid identifies which MODE this target belongs to, so the below-Unit grid
// can follow that segment's own configured chain. Without it every segment would
// share one hardcoded Year → Document Type shape, which is the bug this replaces.
type ProvTarget = { termGuid: string | null; assignTerm: string; ancestorTerms: string[]; relPath: string; label: string; fullName: string; section: string; isLeaf: boolean; termSetGuid: string };

// DMS Group Map role → SharePoint permission level. GLOBAL is a privileged
// uploader bypass (not folder-scoped) and is never assigned to a folder.
//
// APR was "Design" until 2026-08-03. Design includes Add Items and Delete Items,
// so approvers could always upload and "approve only" was never enforceable —
// which is exactly the Head-of #2 persona. DMS Approve and DMS Delete are custom
// levels the client creates once per site (copy Contribute / copy Read; see the
// role-personas spec §3.1).
//
// Deploying before those levels exist fails SAFELY and visibly: the assignment
// loop logs `no "<name>" role definition on site` and skips that one grant.
// Nobody loses access — an existing approver keeps the Design grant already on
// the folder — the change simply does not take effect until the level is made.
//
// UPL dropped from "Contribute" to "DMS Upload" on 2026-08-04. Contribute includes
// Delete Items, so every uploader — PIC included — could delete any file in their
// unit folder, not just their own uploads (Delete Items is folder-scoped, never
// author-scoped). That made the client's rule "a PIC must get approval from the
// head of unit before deleting" unenforceable. DMS Upload is Contribute minus
// Delete Items; PICs raise a deletion REQUEST instead. See the
// deletion-request-approval spec.
//
// DELS restores Staging delete for the people who should have it — the head-of
// groups that also upload (Head of Department/Unit 1 and 4). It is a SEPARATE role
// rather than a second upload level so the capability is carried by an explicit
// group name (`_DELS`) instead of an admin remembering which of two upload
// suffixes means "with delete". A wrong pick between two upload roles is invisible
// until someone deletes something.
//
// DELS and DEL share the "DMS Delete" level — Read + Delete Items is what a deleter
// needs in either library — and are kept apart by LIBRARY_ROLES below: DELS is
// Staging-only, DEL is Documents-only. So the client still creates only three
// custom levels.
//
// HC is absent by design: Highly Confidential is Phase 2 (see groupMapModel).
const ROLE_TO_PERMISSION: Record<string, string> = {
  MEMBER: "Read",
  // Re-pointed at the site's actual prefix by applyPermissionPrefix() below. The DMS values are
  // the legacy default, used until the role definitions have been read.
  UPL: "DMS Upload",
  APR: "DMS Approve",
  DEL: "DMS Delete",
  DELS: "DMS Delete",
  // The C-level view role, 2026-08-04. Plain Read, like MEMBER — the difference is not
  // the level but how far the grant travels: every folder in every segment.
  // Documents-only (see LIBRARY_ROLES); a viewer on Staging would be reading other
  // people's pending drafts, which is the isolation rule the whole model rests on.
  //
  GLOBAL: "Read",
  // SEGVIEW — un-retired 2026-08-07 for the client's second C-Level shape, "view its own
  // business segment only". Same level as GLOBAL and the same fan-DOWN; the difference is
  // only how far it travels. A GLOBAL row is termless and reaches every segment; a SEGVIEW
  // row carries a SEGMENT term and reaches that segment's folders alone.
  //
  // Documents-only, like GLOBAL, and for the same reason — see LIBRARY_ROLES. This is the
  // one role where a mistake is both quiet and wide: a SEGVIEW row wrongly accepted on
  // Staging hands one person every unapproved draft in an entire business segment.
  SEGVIEW: "Read",
  // Library entry, 2026-08-04. Plain Read on the LIST so an uploader/approver can open the
  // library at all — Limited Access on the parent chain lets a direct folder URL through but
  // confers no View Items on the list itself, so AllItems.aspx returns Access Denied without
  // this. Deliberately Read and not the group's own level: DMS Upload at library scope would
  // reach everything in the library that inherits, not just their unit folder.
  //
  // Safe at folder scope by construction: ENTRY is absent from LIBRARY_ROLES below, and the
  // folder loop grants only what that table lists. A hand-written ENTRY row at Folder scope
  // is skipped rather than approximated.
  ENTRY: "Read",
};

/**
 * Re-point the three custom levels at whatever prefix this site uses.
 *
 * MUTATES the map in place rather than replacing it, so the half-dozen existing readers
 * (`ROLE_TO_PERMISSION[g.role]`, the accepts() guard, the op counter) pick up the change with no
 * re-wiring — and so a reader that runs before detection still gets a usable legacy value instead
 * of undefined, which those guards read as "this role grants nothing".
 *
 * Resolved from the site's OWN role definitions, never from the list prefix. Those two genuinely
 * diverge: verified live 2026-08-05, the client's site has CRS lists and CRS levels but a DMS
 * content type and a DMS_SITE_MEMBERS group. Deriving one from the other is right today and wrong
 * on the next site.
 */
function applyPermissionPrefix(levelNames: string[]): void {
  const prefix = levelNames.indexOf("CRS Upload") !== -1 ? "CRS"
    : levelNames.indexOf("DMS Upload") !== -1 ? "DMS"
    : undefined;
  // No match: leave the legacy values. Reconciliation then logs `no "DMS Upload" role definition
  // on site` and grants nothing for that role — visible, and better than guessing at a name.
  if (prefix === undefined) return;
  ROLE_TO_PERMISSION.UPL = `${prefix} Upload`;
  ROLE_TO_PERMISSION.APR = `${prefix} Approve`;
  ROLE_TO_PERMISSION.DEL = `${prefix} Delete`;
  ROLE_TO_PERMISSION.DELS = `${prefix} Delete`;
}

// Which roles each library accepts — the isolation rule that keeps viewers off
// pending documents.
//
// A table, not the inline ternary it replaces. That ternary read "Staging gets
// everything except MEMBER", which silently admitted every FUTURE role: DEL
// would have landed on Staging the day it was added, handing deleters other
// people's pending documents. Listing roles explicitly means a new role reaches
// no library until someone names it here.
//
// DELS and DEL both map to "DMS Delete" but must never cross: a DELS row on
// Documents would hand a Staging deleter other people's approved documents, and a
// DEL row on Staging would hand a Documents deleter other people's pending ones.
// The level cannot tell them apart, so this table is the only thing that does.
// SEGVIEW joins GLOBAL on Documents ONLY (2026-08-07). Both are C-Level view roles and
// neither may ever appear in the Staging list: a segment-wide viewer on Staging reads every
// unapproved draft in that segment, which is the exact isolation this table exists to hold.
// UPL and APR joined Documents on 2026-08-09 so one group serves a person in both libraries —
// a PIC no longer needs the unit's base group just to read the approved archive. See
// 2026-08-09-persona-driven-folder-access-design.md. DELS stays Staging-only: a Head of Unit
// deletes PENDING work, never an approved document.
const LIBRARY_ROLES: Record<LibTarget, string[]> = {
  Staging: ["UPL", "APR", "DELS"],
  Documents: ["MEMBER", "DEL", "GLOBAL", "SEGVIEW", "UPL", "APR"],
};

/**
 * Roles whose Documents grant is READ, whatever they mean on the approval library.
 *
 * This is the whole safety of letting one group serve both libraries. ROLE_TO_PERMISSION is
 * flat — one level per role — so listing UPL under Documents WITHOUT this would grant
 * "CRS Upload" there: Contribute minus Delete. Every PIC could then add and edit APPROVED
 * documents in their unit, with no approval step and nothing in any log to show for it. The
 * request was read-only; the flat table alone would have delivered write.
 *
 * A short exception list rather than a second full table, because the table is what an editor
 * reads to answer "what does this role do", and two of them would let the answer depend on
 * which one they happened to open.
 */
const DOCUMENTS_READ_ONLY_ROLES: string[] = ["UPL", "APR"];

/** The level a role grants IN A GIVEN LIBRARY. Always use this, never the raw table. */
function permissionForRole(lib: LibTarget, role: string): string | undefined {
  if (lib === "Documents" && DOCUMENTS_READ_ONLY_ROLES.indexOf(role) > -1) return "Read";
  return ROLE_TO_PERMISSION[role];
}

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

// Throttle-safety tuning for reconciliation. These are the DEFAULTS/fallbacks —
// they can be overridden live via DMS Config `setting` rows (recon_writeDelayMs /
// recon_batchSize / recon_cooldownMs) so the client can tune speed without a redeploy.
// withThrottleRetry() is the safety net: any write that still hits a 429/503 backs
// off and retries, so a shorter delay is safe — an occasional throttle self-heals.
// Lower delay = faster; too low risks 429s whose Retry-After penalty can be large,
// so ~150ms is a sensible aggressive floor. Grid is sequential (the old parallel
// burst was what tripped the throttle originally).
const RECON_WRITE_DELAY_MS = 200;    // pause between folder/permission writes (was 500)
const RECON_BATCH_SIZE      = 300;   // writes before an automatic cooldown (was 150)
const RECON_COOLDOWN_MS     = 2500;  // cooldown length, masked in the UI as "work" (was 4000)
const RECON_EST_HTTP_MS     = 250;   // rough per-write network+server time, on top of the delay
                                     // (used only for the up-front estimate before a live rate exists)

// Human-friendly duration: "45s" or "3m 07s".
const fmtDur = (ms: number): string => {
  const s = Math.max(0, Math.round(ms / 1000));
  if (s < 60) return `${s}s`;
  const sec = s % 60;
  return `${Math.floor(s / 60)}m ${sec < 10 ? "0" : ""}${sec}s`;
};
// Rotating status text shown during the cooldown so the pause reads as progress.
const COOLDOWN_MESSAGES = ["Discombobulating…", "Generating folders…"];

// Retry a SharePoint request on transient throttling, honoring Retry-After (seconds)
// with exponential backoff. Returns the final response (ok or not) after up to 5 retries.
// Beyond the obvious 429/503, SharePoint Online's anti-abuse system can surface a
// sustained WRITE burst as `403 E_ACCESSDENIED` / UnauthorizedAccessException even when
// the caller has full permissions — indistinguishable at the status line from a real
// denial except by the burst context. We treat that signature as retryable but only for
// the first few attempts, so a GENUINE 403 still fails fast (after ~3 backoffs) instead
// of masquerading as success forever.
async function withThrottleRetry(
  doPost: () => Promise<SPHttpClientResponse>,
  onWait?: (ms: number) => void,
): Promise<SPHttpClientResponse> {
  for (let attempt = 0; ; attempt++) {
    const res = await doPost();
    const throttled = res.status === 429 || res.status === 503;
    let burst403 = false;
    if (res.status === 403 && attempt < 3) {
      // Peek at a CLONE so the original body stays readable for the caller.
      const body = await res.clone().text().catch(() => "");
      burst403 = /E_ACCESSDENIED|UnauthorizedAccessException/i.test(body);
    }
    if ((!throttled && !burst403) || attempt >= 5) return res;
    const ra = Number(res.headers.get("Retry-After"));
    const waitMs = ra > 0 ? ra * 1000 : Math.min(30000, 1000 * 2 ** attempt);
    if (onWait) onWait(waitMs);
    await sleep(waitMs);
  }
}

// Live reconciliation progress feed item (rendered in the two-panel progress view).
type ProgItem = { text: string; status: "run" | "ok" | "skip" | "fail" | "warn" | "admin" };

type GroupMapRow = { groupId: string; groupName: string; role: string };

// SharePoint document libraries keep a system "Forms" folder (and other names
// starting with "_") at the root — never show those as manageable sections.
const isSystemFolder = (name: string): boolean =>
  name === "Forms" || name.startsWith("_");

const sanitize = (str: string): string => str.replace(/[\\/:*?"<>|#%]/g, "").trim();
const uid      = (): string => Math.random().toString(36).slice(2, 9);

// Reserved tree key holding NEW top-level folders staged for creation directly
// under the library root. "*" is illegal in SharePoint folder names, so this key
// can never collide with a real discovered section (top-level folder) name.
const NEW_TOP_LEVEL = "*pending-top-level*";

/* ── Types ─────────────────────────────────────────────────────────────────── */

type RoleDef        = { id: number; name: string };
type GroupPick      = { id: string; displayName: string };
type ExistingAssign = { uid: string; principalId: number; title: string; roleDefId: number; kept: boolean };
type PendingAssign  = { uid: string; group: GroupPick; roleDefId: number };
type LogEntry       = { msg: string; ok: boolean };
type LogTab         = "All" | "Documents" | "Staging" | "Warnings" | "Errors";

// Permission state now lives directly on the folder node so a single "Update"
// commit can apply renames, new-folder creation, and permission edits together.
type PermDraft = {
  existing: ExistingAssign[];
  pending: PendingAssign[];
  loaded: boolean;
  loading: boolean;
  isUnique: boolean | null;
  open: boolean;
};

// Arbitrary-depth folder tree. `isNew` folders don't exist in SharePoint yet —
// `path` stays null for them until the Update commit creates them. Existing
// folders always recompute their real path from parentPath + name at commit
// time (never trust a stored path directly) so an ancestor rename earlier in
// the same commit correctly cascades to every descendant.
type FolderNode = {
  id: string;
  name: string;
  newName: string;
  path: string | null;
  isNew: boolean;
  // Existing (non-new) folders only: staged for deletion but not yet recycled.
  // confirmingDelete gates a second click before isDeleted is actually set.
  isDeleted: boolean;
  confirmingDelete: boolean;
  childrenLoaded: boolean;
  children: FolderNode[];
  perm: PermDraft;
};

/* ── Styles ─────────────────────────────────────────────────────────────────── */

const s: Record<string, React.CSSProperties> = {
  wrap:          { maxWidth: 880, margin: "32px auto", padding: "0 24px 48px", fontFamily: "'Segoe UI', sans-serif" },
  h2:            { fontSize: 22, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle:      { fontSize: 13, color: "#666", margin: "0 0 24px" },
  toggleWrap:    { display: "flex", justifyContent: "center", marginBottom: 24 },
  // flexWrap + narrower padding since the tab count reached seven: without wrapping the bar
  // overflows the web part on a laptop and the last tabs become unreachable.
  seg:           { display: "flex", flexWrap: "wrap", justifyContent: "center", maxWidth: "100%", border: "1px solid #0f6c3f", borderRadius: 8, overflow: "hidden" },
  segBtn:        { padding: "8px 16px", fontSize: 13, fontFamily: "'Segoe UI', sans-serif", fontWeight: 600, cursor: "pointer", background: "#fff", color: "#0f6c3f", border: "none", borderRight: "1px solid #0f6c3f" },
  segActive:     { background: "#0f6c3f", color: "#fff" },
  secHeader:     { display: "flex", alignItems: "center", gap: 8, cursor: "pointer", userSelect: "none", margin: "0 0 10px", padding: "4px 0" },
  secTitle:      { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".07em", color: "#0f6c3f", margin: 0 },
  ico:           { fontSize: 11, color: "#0f6c3f", lineHeight: 1, flexShrink: 0 },
  badge:         { fontSize: 11, color: "#888", background: "#f3f3f3", borderRadius: 10, padding: "1px 7px", flexShrink: 0 },
  scrollPane:    { maxHeight: 560, overflowY: "auto", border: "1px solid #e0e0e0", borderRadius: 6, padding: "12px 16px", marginBottom: 8 },
  parentRow:     { display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #f0f0f0" },
  childRow:      { display: "flex", alignItems: "center", gap: 8, padding: "4px 0" },
  chevBtn:       { background: "none", border: "none", cursor: "pointer", padding: "2px 4px", fontSize: 11, color: "#666", lineHeight: 1, flexShrink: 0 },
  renameIn:      { padding: "5px 9px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 13, width: 200, boxSizing: "border-box" },
  wasLabel:      { fontSize: 11, color: "#aaa", fontStyle: "italic", whiteSpace: "nowrap" },
  permBtn:       { marginLeft: "auto", background: "none", border: "1px solid #c8c8c8", borderRadius: 4, padding: "3px 10px", fontSize: 11, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", color: "#444", flexShrink: 0, whiteSpace: "nowrap" },
  permBtnOpen:   { borderColor: "#0f6c3f", color: "#0f6c3f" },
  childrenPane:  { marginLeft: 40, borderLeft: "2px solid #e8f5ee", paddingLeft: 12, marginBottom: 4 },
  permPanel:     { margin: "2px 0 8px", background: "#f8faf8", border: "1px solid #d0e8d8", borderRadius: 6, padding: "10px 12px" },
  permTitle:     { fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em", color: "#0f6c3f", margin: "0 0 8px" },
  assignRow:     { display: "flex", alignItems: "center", gap: 6, marginBottom: 5 },
  chip:          { display: "inline-flex", alignItems: "center", gap: 5, color: "#0f6c3f", fontWeight: 600, fontSize: 11, border: "1px solid #cfe8da", borderRadius: 4, padding: "3px 7px", background: "#fff", maxWidth: 190, boxSizing: "border-box", flexShrink: 0 },
  chipName:      { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  chipX:         { background: "none", border: "none", color: "#888", cursor: "pointer", fontSize: 11, lineHeight: 1, padding: 0, flexShrink: 0 },
  roleSelect:    { flex: "0 0 130px", padding: "4px 6px", border: "1px solid #c8c8c8", borderRadius: 4, fontSize: 12, fontFamily: "'Segoe UI', sans-serif", background: "#fff" },
  undoLink:      { background: "none", border: "none", color: "#0f6c3f", cursor: "pointer", fontSize: 11, padding: 0, fontFamily: "'Segoe UI', sans-serif" },
  searchWrap:    { position: "relative", marginBottom: 6 },
  searchIn:      { padding: "5px 9px", border: "1px solid #c8c8c8", borderRadius: 4, fontFamily: "'Segoe UI', sans-serif", fontSize: 12, width: "100%", boxSizing: "border-box" },
  dropdown:      { position: "absolute", top: 30, left: 0, right: 0, background: "#fff", border: "1px solid #d0d0d0", borderRadius: 4, boxShadow: "0 6px 18px rgba(0,0,0,.14)", zIndex: 100, maxHeight: 180, overflowY: "auto" },
  dropItem:      { padding: "7px 10px", cursor: "pointer", borderBottom: "1px solid #f2f2f2", fontSize: 12 },
  addFolderBtn:  { background: "none", border: "1px dashed #0f6c3f", color: "#0f6c3f", borderRadius: 4, padding: "4px 12px", fontSize: 12, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", marginTop: 8 },
  actions:       { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 16 },
  btn:           { padding: "8px 22px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 13 },
  logBox:        { marginTop: 20, background: "#f5f5f5", borderRadius: 6, padding: "12px 16px" },
  logTitle:      { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".05em", color: "#555", margin: "0 0 8px" },
  toast:         { position: "fixed", top: 24, right: 24, color: "#fff", padding: "14px 44px 14px 16px", borderRadius: 6, fontSize: 13, zIndex: 9999, minWidth: 280, maxWidth: 420, boxShadow: "0 4px 16px rgba(0,0,0,.18)" },
  toastClose:    { position: "absolute", top: 10, right: 12, background: "none", border: "none", cursor: "pointer", color: "#fff", fontSize: 16, opacity: .7, lineHeight: "1" },
  confirmBar:    { display: "flex", alignItems: "center", gap: 10, margin: "4px 0 8px", padding: "8px 10px", background: "#fdf3f3", border: "1px solid #f1c0c0", borderRadius: 4, fontSize: 12, color: "#a4262c" },
  dangerBtn:     { padding: "5px 14px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 12, border: "none", background: "#a4262c", color: "#fff", flexShrink: 0 },
  ghostBtn:      { padding: "5px 14px", borderRadius: 4, cursor: "pointer", fontFamily: "'Segoe UI', sans-serif", fontSize: 12, border: "1px solid #d0d0d0", background: "#fff", color: "#333", flexShrink: 0 },
};

/* ── GroupSearch ─────────────────────────────────────────────────────────────── */

const GroupSearch: React.FC<{
  disabled: boolean;
  placeholder?: string;
  onSearch: (q: string) => Promise<GroupPick[]>;
  onPick: (g: GroupPick) => void;
}> = ({ disabled, placeholder = "Search for a group…", onSearch, onPick }) => {
  const [q,         setQ]         = useState("");
  const [results,   setResults]   = useState<GroupPick[]>([]);
  const [open,      setOpen]      = useState(false);
  const [searching, setSearching] = useState(false);
  const [focused,   setFocused]   = useState(false);

  useEffect(() => {
    if (!focused) return undefined;
    const h = setTimeout(() => {
      setSearching(true);
      onSearch(q.trim())
        .then(r => { setResults(r); setOpen(true); setSearching(false); })
        .catch(() => { setResults([]); setOpen(false); setSearching(false); });
    }, q.trim().length === 0 ? 0 : 300);
    return () => clearTimeout(h);
  }, [q, focused]);

  return (
    <div style={s.searchWrap}>
      <input
        style={s.searchIn}
        placeholder={placeholder}
        value={q}
        disabled={disabled}
        onFocus={() => { setFocused(true); setOpen(true); }}
        onBlur={() => { window.setTimeout(() => { setFocused(false); setOpen(false); }, 150); }}
        onChange={e => setQ(e.target.value)}
      />
      {searching && <span style={{ position: "absolute", right: 8, top: 7, fontSize: 11, color: "#aaa" }}>Searching…</span>}
      {open && (
        <div style={s.dropdown}>
          {results.length > 0 ? results.map(g => (
            <div key={g.id} style={s.dropItem}
              onMouseDown={() => { onPick(g); setOpen(false); setFocused(false); setResults([]); setQ(""); }}
            >
              <div style={{ fontWeight: 600 }}>{g.displayName}</div>
            </div>
          )) : (
            <div style={{ ...s.dropItem, color: "#888", cursor: "default" }}>
              {searching ? "Searching…" : "No groups found"}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

/* ── Main ────────────────────────────────────────────────────────────────────── */

export default function FolderManager({ context }: IFolderManagerProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;

  // Active tab. `Staging`/`Documents` are no longer OFFERED (see the Tab type) — the folder tree
  // they drove is retired, and `libTarget` now only ever holds its initial value, which keeps the
  // tree's helpers compiling until that code is deleted.
  // Opens on Term Abbreviations because that is where the work starts: a term with no code gets no
  // folder, so reconciliation has nothing to build until this tab is filled in.
  //
  // …unless the URL names one. The CRS Settings landing page has three separate rows — Term
  // Abbreviations, Folder Structure Management, Folder Reconciliations — which are all TABS of this
  // single page, so without the deep link they would all land here and two of the three would look
  // broken. An absent or unrecognised hash falls back to the default rather than showing nothing.
  const [tab,          setTab]          = useState<Tab>(() => tabFromDeepLink());
  const [libTarget]                     = useState<LibTarget>("Staging");
  /**
   * True while a mounted structure screen holds unsaved changes. Switching tabs UNMOUNTS it, which
   * discards the edit silently, so the switch is REFUSED rather than confirmed — the Save button is
   * a few pixels away, and a "discard?" prompt would put losing the work one click behind something
   * that looks like ordinary navigation. Copied in behaviour from FolderStructurePage, whose three
   * tabs now live here.
   */
  const [dirty,        setDirty]        = useState(false);
  const [tabBlocked,   setTabBlocked]   = useState(false);
  // Top-level container folders discovered under the library root, in display order.
  const [sections,     setSections]     = useState<Mode[]>([]);
  const [tree,         setTree]         = useState<Record<Mode, FolderNode[]>>({});
  const [libRoot,      setLibRoot]      = useState<string | null>(null);
  // Has primeNames() resolved? The tree CANNOT load before it has. Every library read goes
  // through libApiTitle(), which maps the logical key "Staging" to the live title — and until
  // priming lands that returns the legacy default, so getbytitle('Staging') 404s on this site
  // (the library is titled "Approval Document") and the tab renders "No top-level folders under
  // Staging yet" on a library holding 200 of them. Documents hid the bug: its title never
  // changed, so it loaded correctly whether primed or not.
  const [namesReady,   setNamesReady]   = useState(false);
  const [loading,      setLoading]      = useState(true);
  const [busy,         setBusy]         = useState(false);
  const [roleDefs,     setRoleDefs]     = useState<RoleDef[]>([]);
  const [ownerGroupId, setOwnerGroupId] = useState<number | null>(null);
  const [log,          setLog]          = useState<LogEntry[]>([]);
  // Which slice of the log is on screen. "All" is the default because the prune,
  // orphan-repair and site-entry passes name neither library, so they are reachable
  // from nowhere else.
  const [logTab,       setLogTab]       = useState<LogTab>("All");
  const [toast,        setToast]        = useState<{ message: string; error: boolean } | null>(null);
  const [expandedIds,  setExpandedIds]  = useState<Record<string, boolean>>({});
  // Section collapse state, keyed by section name; sections default to open.
  const [modeOpen,     setModeOpen]     = useState<Record<Mode, boolean>>({});
  // Confirm gate for the Reconciliation tab's provisioning run.
  const [reconConfirm, setReconConfirm] = useState(false);
  // Live reconciliation progress (two-panel view + rotating cooldown text).
  const [reconRunning, setReconRunning] = useState(false);
  const [reconPhase,   setReconPhase]   = useState("");
  // One feed per library per kind — four panels. A single merged pair made a
  // 700-step run read as one undifferentiated wall; the client's question is
  // always "how is Staging doing", never "how is the run doing".
  const emptyFeeds = (): Record<LibTarget, ProgItem[]> => ({ Staging: [], Documents: [] });
  const [folderFeeds,  setFolderFeeds]  = useState<Record<LibTarget, ProgItem[]>>(emptyFeeds);
  const [assignFeeds,  setAssignFeeds]  = useState<Record<LibTarget, ProgItem[]>>(emptyFeeds);
  const [reconCounts,  setReconCounts]  = useState<{ folders: number; assigns: number }>({ folders: 0, assigns: 0 });
  // ETA: planned total throttled ops, run start time, and a ticking "now" so the
  // elapsed/remaining estimate repaints every second even between ops.
  const [reconPlanned, setReconPlanned] = useState(0);
  /**
   * Steps ATTEMPTED, which is what the ETA divides by.
   *
   * This used to reuse `reconCounts`, which counts things that CHANGED — folders
   * created, groups assigned. On a re-run almost nothing changes, so the numerator
   * stayed near zero while the denominator stayed at the full planned total, and
   * the estimate ran away: a re-run that finished in three minutes advertised
   * "~197m left". Attempted-vs-planned is the only pair that measures the same
   * thing on both sides.
   */
  const [reconDone,    setReconDone]    = useState(0);
  const [reconStartMs, setReconStartMs] = useState<number | undefined>(undefined);
  const [reconNow,     setReconNow]     = useState(0);

  useEffect(() => {
    if (!reconRunning) return;
    const id = setInterval(() => setReconNow(Date.now()), 1000);
    return () => clearInterval(id);
  }, [reconRunning]);

  // Guard against losing a run to a stray refresh or tab close. Reconciliation executes
  // entirely in this page — there is no server-side job — so navigating away stops it
  // mid-operation. Work already committed survives and a re-run resumes safely, but a
  // folder interrupted between "created" and "inheritance broken" is briefly left
  // INHERITING its parent's permissions until the next run repairs it. Worth a prompt.
  useEffect(() => {
    if (!reconRunning) return;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      // Browsers show their own wording and ignore ours, but a non-empty returnValue is
      // still what triggers the prompt at all.
      e.returnValue = "Reconciliation is still running. Leaving now will stop it.";
      return e.returnValue;
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [reconRunning]);

  // Keep the live feeds short so hundreds of ops don't flood the DOM.
  const FEED_CAP = 40;
  const pushFolder = (lib: LibTarget, text: string, status: ProgItem["status"]): void =>
    setFolderFeeds((f) => ({ ...f, [lib]: [...f[lib].slice(-(FEED_CAP - 1)), { text, status }] }));
  const setLastFolder = (lib: LibTarget, text: string, status: ProgItem["status"]): void =>
    setFolderFeeds((f) => ({
      ...f,
      [lib]: f[lib].length ? [...f[lib].slice(0, -1), { text, status }] : [{ text, status }],
    }));
  const pushAssign = (lib: LibTarget, text: string, status: ProgItem["status"]): void =>
    setAssignFeeds((a) => ({ ...a, [lib]: [...a[lib].slice(-(FEED_CAP - 1)), { text, status }] }));
  /** One planned step attempted — succeeded, skipped or failed alike. */
  const step = (n = 1): void => setReconDone((d) => d + n);
  // Surfaces a 429 backoff wait in the rotating status line (safety-net path).
  const reconWaitNote = (ms: number): void =>
    setReconPhase(`Easing off — SharePoint is busy (${Math.round(ms / 1000)}s)…`);

  // Status glyph for a progress-feed row (spinner while running, else a colored mark).
  const progIcon = (status: ProgItem["status"]): React.ReactElement => {
    if (status === "run") {
      return <span style={{ display: "inline-block", width: 11, height: 11, border: "2px solid #cfe4d8", borderTopColor: "#0f6c3f", borderRadius: "50%", animation: "fmspin 0.8s linear infinite", flexShrink: 0 }} />;
    }
    const marks: Record<string, [string, string]> = {
      ok: ["✓", "#0f6c3f"], skip: ["○", "#999"], fail: ["✗", "#c0392b"], admin: ["⚠", "#b45309"], warn: ["⚠", "#b45309"],
    };
    const [ch, color] = marks[status] ?? ["•", "#666"];
    return <span style={{ color, fontSize: 12, width: 11, textAlign: "center", flexShrink: 0 }}>{ch}</span>;
  };

  const showToast = (message: string, error: boolean): void => {
    setToast({ message, error });
    setTimeout(() => setToast(null), 5000);
  };

  const roleName = (id: number): string => roleDefs.find(r => r.id === id)?.name ?? String(id);

  /* ── REST ────────────────────────────────────────────────────────────────────── */

  const getLibraryRoot = async (lib: string): Promise<string | null> => {
    // Retry-wrapped: a single throttled GET here used to null out the WHOLE library
    // pass in reconciliation (e.g. Documents skipped → no MEMBER grants land). Log the
    // real HTTP status on genuine failure rather than reporting a false "not found"
    // (CLAUDE.md gotcha #9 — a transient status is not missing data).
    const res: SPHttpClientResponse = await withThrottleRetry(() => context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    ), reconWaitNote);
    if (!res.ok) {
      console.warn(`getLibraryRoot('${lib}') failed: HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
      return null;
    }
    const data = await res.json();
    return data.ServerRelativeUrl ?? null;
  };

  const getFolders = async (folderPath: string): Promise<Array<{ Name: string; ServerRelativeUrl: string }>> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')/Folders?$select=Name,ServerRelativeUrl&$orderby=Name`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    return data.value ?? [];
  };

  const moveFolder = async (oldPath: string, newPath: string): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@s)/MoveTo(newUrl=@d)?@s='${encodeURIComponent(oldPath)}'&@d='${encodeURIComponent(newPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const json = await res.json();
        const sp = json?.error?.message?.value ?? json?.error?.message ?? json?.["odata.error"]?.message?.value;
        if (sp) msg += ` — ${sp}`;
      } catch {
        msg += ` — ${(await res.text().catch(() => "")).slice(0, 200)}`;
      }
      throw new Error(msg);
    }
  };

  const getRoleAssignments = async (folderPath: string): Promise<ExistingAssign[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments?$expand=Member,RoleDefinitionBindings&@f='${encodeServerRelativePath(folderPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return [];
    const data = await res.json();
    const result: ExistingAssign[] = [];
    for (const ra of (data.value ?? [])) {
      if (ownerGroupId !== null && ra.PrincipalId === ownerGroupId) continue;
      const rawBindings = ra.RoleDefinitionBindings;
      const bindings: Array<{ RoleTypeKind: number; Id: number }> = Array.isArray(rawBindings) ? rawBindings : (rawBindings?.value ?? rawBindings?.results ?? []);
      const valid = bindings.find(b => b.RoleTypeKind !== 1 && b.RoleTypeKind !== 7);
      if (!valid) continue;
      result.push({ uid: uid(), principalId: ra.PrincipalId, title: ra.Member?.Title ?? String(ra.PrincipalId), roleDefId: valid.Id, kept: true });
    }
    return result;
  };

  const getHasUniquePerms = async (folderPath: string): Promise<boolean | null> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields?$select=HasUniqueRoleAssignments&@f='${encodeServerRelativePath(folderPath)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return null;
    const data = await res.json();
    return typeof data.HasUniqueRoleAssignments === "boolean" ? data.HasUniqueRoleAssignments : null;
  };

  const folderExists = async (path: string): Promise<boolean> => {
    const res = await context.spHttpClient.get(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    return res.ok;
  };

  const createFolder = async (path: string): Promise<void> => {
    if (await folderExists(path)) return;
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/folders/AddUsingPath(DecodedUrl=@d,overwrite=false)?@d='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" } },
    ), reconWaitNote);
    if (!res.ok) throw new Error(`create folder HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
  };

  const breakInheritance = async (path: string): Promise<void> => {
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ), reconWaitNote);
    if (!res.ok) throw new Error(`breakroleinheritance HTTP ${res.status}`);
  };

  // Recycles the folder (and everything inside it) to the site Recycle Bin —
  // recoverable, not a permanent delete.
  const deleteFolder = async (path: string): Promise<void> => {
    const res = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/recycle()?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`delete folder HTTP ${res.status} — ${(await res.text().catch(() => "")).slice(0, 200)}`);
  };

  // SP site group: the group's integer Id IS the role-assignment principal id.
  // No ensureuser, no federateddirectoryclaimprovider claim. A GUID here means a
  // legacy Entra row that must be recreated via the Group Map tab.
  const spGroupPrincipalId = (groupId: string): number => {
    const n = Number((groupId ?? "").trim());
    if (!(n > 0) || n % 1 !== 0) {
      throw new Error(
        `"${groupId}" is not a SharePoint site-group id — recreate this Group Map row with the Group Map tab`,
      );
    }
    return n;
  };

  // The only call in this component that REMOVES access. Deliberately narrow: it takes
  // a principal, not a role definition, because SharePoint's removeroleassignment drops
  // every binding that principal holds on that folder. Callers must therefore have
  // already established that the ONLY binding there is the one they mean to remove —
  // see the ancestor-read revoke in the assignment loop, which checks the group holds
  // exactly "Read" and holds no Group Map row of its own at that tier.
  //
  // Limited Access is NOT removed by this and must not be: SharePoint maintains it on
  // parent folders so a user can reach a child they are granted on. Stripping it would
  // break access to the unit folder itself, which is the opposite of the intent.
  // UNREACHABLE ON PURPOSE, and kept. `revokeAncestorRead` is hard-coded false in both
  // ReconSettings defaults — the config row is deliberately ignored, because revoking
  // ancestor Read fights the browse corridor that lets a user click down to their own
  // folder (CLAUDE.md, RBAC section). Retained rather than deleted so the capability and
  // the reasoning above survive if the client ever asks for it; deleting it would mean
  // rediscovering the Limited Access caveat from scratch.
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const removeRoleAssignment = async (path: string, principalId: number): Promise<void> => {
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/removeroleassignment(principalid=${principalId})?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ));
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const json = await res.json();
        const sp = json?.error?.message?.value ?? json?.error?.message ?? json?.["odata.error"]?.message?.value;
        if (sp) msg += ` — ${sp}`;
      } catch {
        msg += ` — ${(await res.text().catch(() => "")).slice(0, 200)}`;
      }
      throw new Error(msg);
    }
  };

  // Library-scope grant. Separate from addRoleAssignment because that one addresses a
  // FOLDER by server-relative path; a list is addressed by its GUID, and the two endpoints
  // are not interchangeable. Takes the `/_api/web/lists(guid'…')` base already built by the
  // caller, so the GUID is resolved once per library rather than per grant.
  const addRoleAssignmentToList = async (listBase: string, principalId: number, roleDefId: number): Promise<void> => {
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${listBase}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ));
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const json = await res.json();
        const sp = json?.error?.message?.value ?? json?.error?.message ?? json?.["odata.error"]?.message?.value;
        if (sp) msg += ` — ${sp}`;
      } catch {
        msg += ` — ${(await res.text().catch(() => "")).slice(0, 200)}`;
      }
      throw new Error(msg);
    }
  };

  const addRoleAssignment = async (path: string, principalId: number, roleDefId: number): Promise<void> => {
    const res = await withThrottleRetry(() => context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})?@f='${encodeServerRelativePath(path)}'`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    ), reconWaitNote);
    if (!res.ok) throw new Error(`addroleassignment HTTP ${res.status}`);
  };

  const searchGroups = async (query: string): Promise<GroupPick[]> => {
    const groups = await searchSiteGroups(context.spHttpClient, siteUrl, query);
    return groups.map((g) => ({ id: String(g.id), displayName: g.title }));
  };

  /* ── Init ────────────────────────────────────────────────────────────────────── */

  useEffect(() => {
    const loadRoleDefs = async (): Promise<void> => {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/roledefinitions?$select=Id,Name,Hidden,RoleTypeKind`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return;
      const data = await res.json();
      const all = (data.value ?? []) as Array<{ Id: number; Name: string; Hidden: boolean; RoleTypeKind: number }>;
      // Detect the custom-level prefix from the UNFILTERED list, before the hidden/system levels
      // are dropped — the filter is about what an admin may pick, not about what exists.
      applyPermissionPrefix(all.map(r => r.Name));
      setRoleDefs(all.filter(r => !r.Hidden && r.RoleTypeKind !== 1 && r.RoleTypeKind !== 7).map(r => ({ id: r.Id, name: r.Name })));
    };
    const loadOwnerGroup = async (): Promise<void> => {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/AssociatedOwnerGroup?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return;
      const data = await res.json();
      // Spelled out rather than `!= null`: the loose form meant "neither null nor undefined",
      // which is right, but "fixing" it to `!== null` would pass undefined into a number.
      if (data.Id !== null && data.Id !== undefined) setOwnerGroupId(data.Id as number);
    };
    // Names first, then everything else. Every list read in this component resolves through the
    // cache, and an unprimed cache falls back to the legacy DMS titles — which 404 on a
    // CRS-renamed site and would present as an empty term store rather than a naming problem.
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      // Set even when priming FAILED: the cache then holds the legacy defaults, which is the
      // best guess available and is what this component used before priming existed. Leaving
      // the flag false would strand the tree on a permanent spinner over a transient GET.
      .then(() => setNamesReady(true))
      .then(() => Promise.all([loadRoleDefs(), loadOwnerGroup()]))
      .catch(() => undefined);
  }, []);

  /* ── Tree ────────────────────────────────────────────────────────────────────── */

  const makeNode = (name: string, path: string): FolderNode => ({
    id: uid(), name, newName: name, path, isNew: false, isDeleted: false, confirmingDelete: false, childrenLoaded: false, children: [],
    perm: { existing: [], pending: [], loaded: false, loading: false, isUnique: null, open: false },
  });

  const makeNewNode = (): FolderNode => ({
    id: uid(), name: "", newName: "", path: null, isNew: true, isDeleted: false, confirmingDelete: false, childrenLoaded: true, children: [],
    perm: { existing: [], pending: [], loaded: true, loading: false, isUnique: null, open: true },
  });

  const loadTree = async (): Promise<void> => {
    setLoading(true);
    setExpandedIds({});
    const root = await getLibraryRoot(libTarget);
    setLibRoot(root);
    if (!root) {
      setSections([]);
      setTree({});
      setLoading(false);
      showToast(`Could not find the "${libTarget}" library.`, true);
      return;
    }
    // Discover the top-level container folders under the library root (e.g. the
    // segment folders) instead of assuming a fixed Departments/Projects layout.
    const topFolders = (await getFolders(root))
      .filter(f => !isSystemFolder(f.Name));
    const sectionNames = topFolders.map(f => f.Name);
    const result: Record<Mode, FolderNode[]> = {};
    for (const f of topFolders) {
      const parents = await getFolders(f.ServerRelativeUrl);
      result[f.Name] = parents.map(p => makeNode(p.Name, p.ServerRelativeUrl));
    }
    setSections(sectionNames);
    setTree(result);
    setLoading(false);
  };

  // Waits for namesReady — see the flag's declaration. Re-runs when it flips, so the tree
  // loads as soon as the titles are known rather than needing a manual Refresh.
  useEffect(() => {
    if (!namesReady) return;
    // The manual folder tree is no longer reachable from the tab bar, so this crawl — dozens of
    // requests down every branch of a library — would run on every page load for a view nobody can
    // open, and its failure toast would accuse the admin of a permissions problem on a tab that is
    // not there. Reconciliation reads its own state and does not depend on this.
    if (tab !== "Staging" && tab !== "Documents") return;
    loadTree().catch(() => { setLoading(false); showToast("Could not load folders. Check your permissions.", true); });
  }, [libTarget, namesReady, tab]);

  /* ── Generic tree mutation helpers (recursive, keyed by node id) ───────────────── */

  const mapTree = (nodes: FolderNode[], id: string, updater: (n: FolderNode) => FolderNode): FolderNode[] =>
    nodes.map(n => n.id === id ? updater(n) : (n.children.length > 0 ? { ...n, children: mapTree(n.children, id, updater) } : n));

  const updateNode = (id: string, updater: (n: FolderNode) => FolderNode): void =>
    setTree(prev => {
      const next: Record<Mode, FolderNode[]> = {};
      for (const k of Object.keys(prev)) next[k] = mapTree(prev[k], id, updater);
      return next;
    });

  const filterTree = (nodes: FolderNode[], id: string): FolderNode[] =>
    nodes.filter(n => n.id !== id).map(n => n.children.length > 0 ? { ...n, children: filterTree(n.children, id) } : n);

  const discardNode = (id: string): void =>
    setTree(prev => {
      const next: Record<Mode, FolderNode[]> = {};
      for (const k of Object.keys(prev)) next[k] = filterTree(prev[k], id);
      return next;
    });

  /* ── Expand / lazy-load children ────────────────────────────────────────────── */

  const toggleExpand = async (node: FolderNode): Promise<void> => {
    const isOpen = !!expandedIds[node.id];
    if (isOpen) { setExpandedIds(prev => ({ ...prev, [node.id]: false })); return; }
    setExpandedIds(prev => ({ ...prev, [node.id]: true }));
    if (node.childrenLoaded || node.isNew || !node.path) return;
    const kids = await getFolders(node.path).catch(() => []);
    updateNode(node.id, n => ({ ...n, childrenLoaded: true, children: kids.map(c => makeNode(c.Name, c.ServerRelativeUrl)) }));
  };

  /* ── Add / discard folders (staged in-memory until Update) ─────────────────────── */

  const addNewChild = (mode: Mode, parentId: string | null): void => {
    const node = makeNewNode();
    if (parentId === null) {
      setTree(prev => ({ ...prev, [mode]: [...(prev[mode] ?? []), node] }));
      return;
    }
    setExpandedIds(prev => ({ ...prev, [parentId]: true }));
    updateNode(parentId, n => ({ ...n, children: [...n.children, node] }));
  };

  // Stage a brand-new TOP-LEVEL folder, created directly under the library root
  // on Update. Works even when the library currently has no folders at all.
  const addTopLevel = (): void =>
    setTree(prev => ({ ...prev, [NEW_TOP_LEVEL]: [...(prev[NEW_TOP_LEVEL] ?? []), makeNewNode()] }));

  /* ── Delete existing folders (staged in-memory, requires confirmation, applied on Update) ── */

  const requestDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, confirmingDelete: true }));

  const cancelDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, confirmingDelete: false }));

  const confirmDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, confirmingDelete: false, isDeleted: true }));

  const undoDelete = (id: string): void =>
    updateNode(id, n => ({ ...n, isDeleted: false }));

  /* ── Permission draft mutations (existing + new nodes share the same shape) ────── */

  const togglePermPanel = async (node: FolderNode): Promise<void> => {
    if (node.isNew || node.perm.loaded) {
      updateNode(node.id, n => ({ ...n, perm: { ...n.perm, open: !n.perm.open } }));
      return;
    }
    updateNode(node.id, n => ({ ...n, perm: { ...n.perm, open: true, loading: true } }));
    const path = node.path as string;
    const [assignments, isUnique] = await Promise.all([
      getRoleAssignments(path).catch(() => [] as ExistingAssign[]),
      getHasUniquePerms(path).catch(() => null as boolean | null),
    ]);
    updateNode(node.id, n => ({ ...n, perm: { ...n.perm, loading: false, loaded: true, existing: assignments, isUnique } }));
  };

  const permToggleKept = (id: string, aUid: string): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, existing: n.perm.existing.map(a => a.uid === aUid ? { ...a, kept: !a.kept } : a) } }));

  const permSetExistingRole = (id: string, aUid: string, roleDefId: number): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, existing: n.perm.existing.map(a => a.uid === aUid ? { ...a, roleDefId } : a) } }));

  const permAddGroup = (id: string, group: GroupPick): void =>
    updateNode(id, n => {
      if (n.perm.pending.some(a => a.group.id === group.id)) return n;
      const def = roleDefs.find(r => r.name === "Read") ?? roleDefs[0];
      return { ...n, perm: { ...n.perm, pending: [...n.perm.pending, { uid: uid(), group, roleDefId: def?.id ?? 0 }] } };
    });

  const permSetPendingRole = (id: string, aUid: string, roleDefId: number): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, pending: n.perm.pending.map(a => a.uid === aUid ? { ...a, roleDefId } : a) } }));

  const permRemovePending = (id: string, aUid: string): void =>
    updateNode(id, n => ({ ...n, perm: { ...n.perm, pending: n.perm.pending.filter(a => a.uid !== aUid) } }));

  /* ── Change detection ────────────────────────────────────────────────────────── */

  const nodeHasChanges = (node: FolderNode): boolean => {
    if (node.isNew) return true;
    if (node.isDeleted) return true;
    const renamed = node.newName.trim() !== "" && node.newName.trim() !== node.name;
    const permDirty = node.perm.loaded && (node.perm.existing.some(a => !a.kept) || node.perm.pending.length > 0);
    return renamed || permDirty || node.children.some(nodeHasChanges);
  };

  const hasChanges =
    sections.some(mode => (tree[mode] ?? []).some(nodeHasChanges)) ||
    (tree[NEW_TOP_LEVEL] ?? []).some(nodeHasChanges);

  /* ── Update (rename + create + permissions, all in one commit) ─────────────────── */

  // Validation walks the WHOLE tree regardless of what's currently expanded on
  // screen, so an error can point at a folder buried inside a collapsed section.
  // Each error carries the mode + ancestor chain needed to reveal it in the UI.
  type ValidationError = { message: string; mode: Mode; ancestorIds: string[] };

  const validateTree = (): ValidationError[] => {
    const errors: ValidationError[] = [];
    const walk = (nodes: FolderNode[], mode: Mode, ancestorIds: string[], parentLabel: string): void => {
      for (const n of nodes) {
        // A folder marked for deletion is recycled whole — its children go with
        // it, so any staged edits inside it are moot and don't need validating.
        if (!n.isNew && n.isDeleted) continue;
        if (n.isNew) {
          const nm = n.newName.trim();
          if (!nm) errors.push({ message: `A new folder under "${parentLabel}" is missing a name.`, mode, ancestorIds });
          else if (n.perm.pending.length === 0) errors.push({ message: `"${nm}" (under "${parentLabel}") needs at least one group assigned.`, mode, ancestorIds });
        } else if (n.perm.loaded) {
          const dirty = n.perm.existing.some(a => !a.kept) || n.perm.pending.length > 0;
          if (dirty) {
            const remaining = n.perm.existing.filter(a => a.kept).length + n.perm.pending.length;
            if (remaining === 0) errors.push({ message: `"${n.newName.trim() || n.name}" would end up with no groups assigned.`, mode, ancestorIds });
          }
        }
        if (n.children.length > 0) walk(n.children, mode, [...ancestorIds, n.id], n.newName.trim() || n.name || "(unnamed folder)");
      }
    };
    sections.forEach(mode => walk(tree[mode] ?? [], mode, [], mode));
    // New top-level folders live under the library root; label them by library.
    walk(tree[NEW_TOP_LEVEL] ?? [], NEW_TOP_LEVEL, [], libTarget);
    return errors;
  };

  const handleUpdate = async (): Promise<void> => {
    const problems = validateTree();
    if (problems.length > 0) {
      const first = problems[0];
      setModeOpen(prev => ({ ...prev, [first.mode]: true }));
      if (first.ancestorIds.length > 0) {
        setExpandedIds(prev => {
          const next = { ...prev };
          first.ancestorIds.forEach(id => { next[id] = true; });
          return next;
        });
      }
      showToast(first.message, true);
      return;
    }
    if (!libRoot) { showToast("Library root not found.", true); return; }

    setBusy(true);
    const entries: LogEntry[] = [];
    const fullCtrlId = roleDefs.find(r => r.name === "Full Control")?.id;
    const ensurePrincipal = async (group: GroupPick): Promise<number> =>
      spGroupPrincipalId(group.id);

    // Pre-order walk: parents are created/renamed before their children are
    // processed, so each child always receives its parent's up-to-date path.
    const processNode = async (node: FolderNode, parentPath: string): Promise<string | null> => {
      const trimmedNew = node.newName.trim();

      if (!node.isNew && node.isDeleted) {
        const currentPath = `${parentPath}/${node.name}`;
        try {
          await deleteFolder(currentPath);
          entries.push({ msg: `"${node.name}" — deleted ✓`, ok: true });
        } catch (e) {
          entries.push({ msg: `"${node.name}" — delete FAILED: ${(e as Error).message}`, ok: false });
        }
        // Recycling the folder takes its whole subtree with it — nothing left to descend into.
        return null;
      }

      if (node.isNew) {
        const fp = `${parentPath}/${sanitize(trimmedNew)}`;
        try {
          await createFolder(fp);
          await breakInheritance(fp);
          if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(fp, ownerGroupId, fullCtrlId);
          for (const a of node.perm.pending) {
            const pid = await ensurePrincipal(a.group);
            await addRoleAssignment(fp, pid, a.roleDefId);
          }
          const summary = node.perm.pending.map(a => `${a.group.displayName}=${roleName(a.roleDefId)}`).join(", ");
          entries.push({ msg: `"${trimmedNew}" — created · ${summary}`, ok: true });
          for (const child of node.children) await processNode(child, fp);
          return fp;
        } catch (e) {
          entries.push({ msg: `"${trimmedNew}" — create FAILED: ${(e as Error).message}`, ok: false });
          return null;
        }
      }

      // Existing node: always derive the CURRENT real path from parentPath + the
      // original name, since an ancestor rename earlier in this same commit
      // shifts every descendant's real URL — never trust the stored node.path here.
      const currentPath = `${parentPath}/${node.name}`;
      let resolvedPath = currentPath;

      if (trimmedNew && trimmedNew !== node.name) {
        const newPath = `${parentPath}/${sanitize(trimmedNew)}`;
        try {
          await moveFolder(currentPath, newPath);
          entries.push({ msg: `"${node.name}" → "${trimmedNew}" ✓`, ok: true });
          resolvedPath = newPath;
        } catch (e) {
          entries.push({ msg: `"${node.name}" → "${trimmedNew}" FAILED: ${(e as Error).message}`, ok: false });
        }
      }

      const perm = node.perm;
      const permDirty = perm.loaded && (perm.existing.some(a => !a.kept) || perm.pending.length > 0);
      if (permDirty) {
        const kept = perm.existing.filter(a => a.kept);
        try {
          await breakInheritance(resolvedPath);
          if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(resolvedPath, ownerGroupId, fullCtrlId);
          for (const a of kept) await addRoleAssignment(resolvedPath, a.principalId, a.roleDefId);
          for (const a of perm.pending) {
            const pid = await ensurePrincipal(a.group);
            await addRoleAssignment(resolvedPath, pid, a.roleDefId);
          }
          entries.push({ msg: `"${trimmedNew || node.name}" — permissions updated ✓`, ok: true });
        } catch (e) {
          entries.push({ msg: `"${trimmedNew || node.name}" — permissions FAILED: ${(e as Error).message}`, ok: false });
        }
      }

      for (const child of node.children) await processNode(child, resolvedPath);
      return resolvedPath;
    };

    for (const mode of sections) {
      const modeRootPath = `${libRoot}/${mode}`;
      for (const node of (tree[mode] ?? [])) {
        await processNode(node, modeRootPath);
      }
    }
    // New top-level folders are created directly under the library root.
    for (const node of (tree[NEW_TOP_LEVEL] ?? [])) {
      await processNode(node, libRoot);
    }

    setLog(entries);
    setBusy(false);
    const failed = entries.filter(e => !e.ok).length;
    const ok     = entries.filter(e => e.ok).length;
    showToast(
      failed > 0 ? `${ok} update${ok !== 1 ? "s" : ""} applied, ${failed} failed — see log.` :
      ok > 0     ? `${ok} update${ok !== 1 ? "s" : ""} applied.` :
                   "No changes to update.",
      failed > 0,
    );
    await loadTree();
  };

  /* ── Folder Reconciliation (term-store provisioner) ────────────────────────────── */

  // Term-store readers (v2.1 taxonomy API) — same calls the retired Reconciliation
  // web part used. Top-level terms = the first level under the set (departments for
  // GHO); children recurse to the leaf (unit) terms.
  const loadReconTops = async (termSetId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    if (!res.ok) throw new Error(`Term set ${termSetId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({ id: t.id, label: t.labels[0].name }));
  };

  const loadReconChildren = async (termSetId: string, parentId: string): Promise<TermLite[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}/terms/${parentId}/children`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    // THROWS on failure — do NOT soften this to `return []`. A silently empty child
    // list truncates the term tree, and prune reads "term not enumerated" as "term
    // deleted". Swallowing this error would let a throttle wipe a segment's map rows.
    if (!res.ok) throw new Error(`Term children ${parentId} returned ${res.status}`);
    const data = await res.json();
    return (data.value ?? []).map((t: { id: string; labels: Array<{ name: string }> }) => ({ id: t.id, label: t.labels[0].name }));
  };

  // Name of a term SET (not a term), used as the full name of the segment container
  // folder. Soft-fails to undefined: a missing label is cosmetic, and unlike the term
  // TREE this value feeds nothing that prune or routing depends on.
  const loadTermSetName = async (termSetId: string): Promise<string | undefined> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${termSetId}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (!res.ok) return undefined;
      const data = await res.json();
      const name = (data.localizedNames ?? [])[0]?.name;
      return typeof name === "string" && name.trim() !== "" ? name : undefined;
    } catch {
      return undefined;
    }
  };

  // Resolve the Full Name column's INTERNAL name on a library, live. Never guessed —
  // see the folderFullName module for why the display name does not determine it.
  // Returns undefined when the column has not been added to that library yet, which
  // is a soft state: the run proceeds and simply writes no full names.
  const loadFullNameField = async (
    lib: LibTarget,
  ): Promise<{ internalName?: string; note?: string }> => {
    // No $filter. A filter that fails returns the same undefined as a column that is
    // genuinely absent, and the first run against a real site could not tell the two
    // apart — CLAUDE.md gotcha #9. Reading all fields and matching in code costs one
    // unfiltered read per library per run and removes the ambiguity.
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')/fields` +
      `?$select=Title,InternalName,TypeAsString,ReadOnlyField&$top=500`;
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) {
        const body = (await res.text().catch(() => "")).slice(0, 200);
        return { note: `field read failed: HTTP ${res.status} ${body}` };
      }
      const data = await res.json();
      const fields = (data.value ?? []) as SpFieldLite[];
      const internalName = pickFullNameField(fields);
      if (internalName) return { internalName };
      // Name the near-misses. "Column absent" and "column present but the wrong type"
      // need different fixes, and the client cannot tell which one they are looking at.
      const near = fields.filter((f) =>
        (f.Title ?? "").trim().toLowerCase().indexOf("full") >= 0 ||
        (f.InternalName ?? "").toLowerCase().indexOf("full") >= 0,
      );
      if (near.length > 0) {
        const desc = near
          .map((f) => `"${f.Title}" (${f.InternalName}, ${f.TypeAsString}${f.ReadOnlyField ? ", read-only" : ""})`)
          .join("; ");
        return { note: `no usable "${FULL_NAME_COLUMN_TITLE}" column. Closest match: ${desc}` };
      }
      return { note: `no column titled "${FULL_NAME_COLUMN_TITLE}" (${fields.length} fields read)` };
    } catch (e) {
      return { note: `field read threw: ${(e as Error).message}` };
    }
  };

  // Resolve the DMS Folder content type's id on a library, by NAME.
  //
  // SharePoint's built-in Folder content type cannot take a custom column: its list
  // settings page offers only "Delete this content type", and adding a field link by API
  // fails too (verified 2026-08-03). So a folder that should show Full Name in the details
  // pane has to carry a folder-derived content type that does accept columns, and this run
  // stamps it — a folder created through the REST API gets the plain Folder type whatever
  // the library's default-content-type setting says, because that governs the New button
  // only.
  //
  // Resolved per library: the two libraries hold separate list-scoped copies of the same
  // site content type, with DIFFERENT ids. Undefined when it is absent, which is a soft
  // state — a site that has not provisioned it still gets all its folders.
  const loadFolderContentTypeId = async (lib: LibTarget): Promise<string | undefined> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')/ContentTypes?$select=Id,Name`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return undefined;
      const data = await res.json();
      const rows = (data.value ?? []) as Array<{ Id?: { StringValue?: string }; Name?: string }>;
      // Preference order, not "whichever matches first in the library's list": a library
      // mid-rename can carry BOTH, and stamping the one being retired would mean the next
      // run has to rewrite every folder again.
      for (const candidate of FOLDER_CONTENT_TYPE_CANDIDATES) {
        const hit = rows.find(
          (c) => (c.Name ?? "").trim().toLowerCase() === candidate.toLowerCase(),
        );
        if (hit?.Id?.StringValue) {
          resolvedFolderCtName = candidate;
          return hit.Id.StringValue;
        }
      }
      return undefined;
    } catch {
      return undefined;
    }
  };

  // Read what a folder's list item already holds, so an unchanged folder costs a cheap GET
  // instead of a throttled write. Undefined fields mean "unknown" (empty, or the read
  // failed) and the caller writes.
  const getFolderItemState = async (
    serverRelativeUrl: string,
    internalName?: string,
    withModeration?: boolean,
  ): Promise<{ fullName?: string; contentTypeId?: string; moderationStatus?: number }> => {
    try {
      // The column is optional: a library can carry the content type before anyone adds
      // Full Name to it, and the content type must still be stamped in that state.
      //
      // OData__ModerationStatus is requested ONLY when the library has content approval on.
      // Naming a field that does not exist fails the WHOLE request with HTTP 400 — not a null
      // (CLAUDE.md #11) — so on a library without moderation this would break the Full Name
      // and content-type writes too.
      const base = internalName ? `${internalName},ContentTypeId` : "ContentTypeId";
      const select = withModeration ? `${base},OData__ModerationStatus` : base;
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
          `?@f='${encodeServerRelativePath(serverRelativeUrl)}'&$select=${select}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return {};
      const data = await res.json();
      const v = internalName ? data[internalName] : undefined;
      const ct = data.ContentTypeId;
      const mod = data.OData__ModerationStatus;
      return {
        fullName: typeof v === "string" && v !== "" ? v : undefined,
        contentTypeId: typeof ct === "string" && ct !== "" ? ct : undefined,
        moderationStatus: typeof mod === "number" ? mod : undefined,
      };
    } catch {
      return {};
    }
  };

  // Write the given fields onto a folder's list item. Throws on failure so the caller can
  // report it against the folder rather than losing it silently.
  const setFolderItemFields = async (
    serverRelativeUrl: string,
    // Numbers as well as strings: OData__ModerationStatus is an integer, and quoting it
    // makes SharePoint reject the merge (memory sp-column-formatting-gotchas records the
    // same trap on the read side — the status is an integer, never a label).
    values: Record<string, string | number>,
  ): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/ListItemAllFields` +
        `?@f='${encodeServerRelativePath(serverRelativeUrl)}'`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          // MERGE leaves every other field alone; IF-MATCH:* skips the etag round-trip.
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify(values),
      },
    );
    // validateUpdateListItem returns 200 on field errors, but a plain MERGE does not —
    // it 4xxs. Still surface the body: "field does not exist" and "read-only" read
    // identically as a bare status code.
    if (!res.ok) {
      const detail = (await res.text().catch(() => "")).slice(0, 200);
      throw new Error(`HTTP ${res.status} ${detail}`);
    }
  };

  // Load DMS Group Map keyed by term (lowercased UnitTermGuid) → its group rows.
  // Same list/fields the upload form reads. A term can have several rows (one per
  // role), so a folder gets every mapped group at its role's permission level.
  const loadGroupMapForAssign = async (): Promise<Map<string, GroupMapRow[]>> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items?$select=GroupName,GroupId,UnitTermGuid,Role&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    const map = new Map<string, GroupMapRow[]>();
    if (!res.ok) return map;
    const data = await res.json();
    for (const r of (data.value ?? []) as Array<{ GroupName?: string; GroupId?: string; UnitTermGuid?: string; Role?: string }>) {
      const term = (r.UnitTermGuid ?? "").toLowerCase();
      // Accepts the long-form value an admin naturally types now that the GROUP NAMES use
      // long suffixes ("UPLOADER" for UPL). Unrecognised values pass through and still fail
      // accepts() below, so this widens what works without widening what is granted.
      const role = normalizeRoleValue(r.Role ?? "");
      if (!r.GroupId) continue;
      // A GLOBAL row carries NO term by design — it is not scoped to a segment, so
      // buildGroupMapRow forces Segment and UnitTermGuid empty. The termless guard
      // below therefore used to discard it, which is why GLOBAL granted nothing
      // anywhere. Keyed under "" instead, which is where the C-level fan-down looks
      // for it. Every OTHER termless row is still dropped: without a term there is no
      // folder to grant on, so it is a broken row, not a wide one.
      if (!term && role !== "GLOBAL") continue;
      const arr = map.get(term) ?? [];
      arr.push({ groupId: r.GroupId, groupName: r.GroupName ?? r.GroupId, role });
      map.set(term, arr);
    }
    return map;
  };

  /**
   * Group Map rows with their item Ids, for the orphan-repair pass.
   *
   * Separate from loadGroupMapForAssign because that one keys by term and drops the
   * Id — repair needs to WRITE specific rows, and re-pointing the wrong one is a
   * permissions change.
   */
  /**
   * Immediate subfolder names of a folder, by server-relative path.
   *
   * Uses the OData parameter alias, not an inline quoted literal: an inline path
   * returns HTTP 400 once it is long enough (CLAUDE.md gotcha #9), which reads as a
   * malformed request but looks like "no subfolders" to a caller that ignores it.
   * Returns undefined — not [] — when the listing FAILED, so "could not read" is
   * distinguishable from "genuinely empty". Reporting a folder as unclaimed on the
   * strength of a throttled read would be the same class of bug the prune guard exists for.
   */
  const listSubfolders = async (
    serverRelativeUrl: string,
  ): Promise<string[] | undefined> => {
    const res: SPHttpClientResponse = await withThrottleRetry(() =>
      context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Folders?$select=Name&@f='${encodeServerRelativePath(serverRelativeUrl)}'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      ),
    );
    if (!res.ok) return undefined;
    const data = await res.json();
    return ((data.value ?? []) as Array<{ Name?: string }>)
      .map((f) => f.Name ?? "")
      .filter((n) => n.length > 0);
  };

  const loadGroupMapRowsForRepair = async (): Promise<
    Array<{ itemId: number; termGuid: string; groupName: string }>
  > => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items?$select=Id,GroupName,UnitTermGuid&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`DMS Group Map read failed: HTTP ${res.status}. ${body}`);
    }
    const data = await res.json();
    return ((data.value ?? []) as Array<{ Id?: number; GroupName?: string; UnitTermGuid?: string }>)
      .map((r) => ({
        itemId: r.Id ?? 0,
        termGuid: (r.UnitTermGuid ?? "").trim(),
        groupName: r.GroupName ?? "",
      }))
      .filter((r) => r.itemId > 0 && r.termGuid.length > 0);
  };

  /** MERGE one field on one list item. Throws with the body — a bare status hides
   *  "field does not exist" behind the same 400 as "read-only". */
  const patchListItem = async (
    listTitle: string,
    itemId: number,
    fields: Record<string, string>,
  ): Promise<void> => {
    const res: SPHttpClientResponse = await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },
        body: JSON.stringify(fields),
      },
    );
    if (!res.ok) {
      const detail = (await res.text().catch(() => "")).slice(0, 200);
      throw new Error(`HTTP ${res.status} ${detail}`);
    }
  };

  /**
   * Term GUID → folder-name abbreviation.
   *
   * Lives here rather than in the shared module because a shared file that imports
   * `@microsoft/sp-http` cannot be unit tested — see the note on abbrevListTitle().
   *
   * Deliberately NOT wrapped in a catch returning an empty map: an unreadable list
   * must abort the run. An empty index makes every term look unmapped, which would
   * skip every folder and report 175 false "needs attention" rows.
   */
  const loadAbbrevRows = async (): Promise<OrphanAbbrevRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(abbrevListTitle())}')/items?$select=Id,TermGuid,Title,Level,Abbreviation&$top=5000`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`${abbrevListTitle()} read failed: HTTP ${res.status}. ${body}`);
    }
    const data = await res.json();
    // Id/Title/Level are read for the orphan-repair pass, which matches a dead row
    // to a re-created term by LABEL — the one thing that survives a delete-and-re-add.
    return ((data.value ?? []) as Array<{
      Id?: number;
      TermGuid?: string;
      Title?: string;
      Level?: string;
      Abbreviation?: string;
    }>).map((r) => ({
      itemId: r.Id ?? 0,
      termGuid: r.TermGuid ?? "",
      title: r.Title ?? "",
      level: r.Level ?? "",
      abbreviation: r.Abbreviation ?? "",
    }));
  };

  const abbrevIndexOf = (rows: readonly OrphanAbbrevRow[]): Map<string, string> =>
    buildAbbrevIndex(
      rows.map(
        (r): AbbrevRow => ({ termGuid: r.termGuid, abbreviation: r.abbreviation }),
      ),
    );

  /**
   * Term set GUID → the segment's level names, in depth order, from the same DMS
   * Config `Levels` JSON the upload form reads.
   *
   * Needed because the abbreviation list stores a level NAME ("Unit"), and an
   * orphan is only ever matched to a live term at the SAME level. Falls back to
   * the pilot's shared chain, which is what all four Head Office segments use.
   */
  const FALLBACK_LEVEL_NAMES = ["Department", "Unit"];

  /**
   * Term set GUID → the mode's below-Unit tiers, in path order.
   *
   * Empty for every site that has not configured one, which is all of them today —
   * `needsLegacyBelowUnit` then keeps the hardcoded Year → Document Type grid.
   */
  const loadReconOnDemandTiers = async (): Promise<Map<string, Level[]>> => {
    const out = new Map<string, Level[]>();
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items?$select=TermSetGuid,Levels&$filter=ConfigType eq 'mode'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return out;
      const data = await res.json();
      ((data.value ?? []) as Array<{ TermSetGuid?: string; Levels?: string }>).forEach((r) => {
        const guid = (r.TermSetGuid ?? "").trim().toLowerCase();
        if (!guid) return;
        const levels = parseLevels(r.Levels ?? "");
        // A malformed chain must not half-build a tree. Reconciliation reports it and
        // falls back to the legacy shape rather than creating folders in an order the
        // upload form will not agree with.
        if (validateChain(levels)) return;
        const { onDemand } = splitChain(levels);
        if (onDemand.length > 0) out.set(guid, onDemand);
      });
    } catch {
      // Same posture as the level names below: a config hiccup costs the configured
      // shape, not the run, and the legacy grid is correct for every live site.
    }
    return out;
  };

  const loadReconLevelNames = async (): Promise<Map<string, string[]>> => {
    const out = new Map<string, string[]>();
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items?$select=TermSetGuid,Levels&$filter=ConfigType eq 'mode'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return out;
      const data = await res.json();
      ((data.value ?? []) as Array<{ TermSetGuid?: string; Levels?: string }>).forEach((r) => {
        const guid = (r.TermSetGuid ?? "").trim().toLowerCase();
        if (!guid) return;
        // PERMISSIONED TIERS ONLY. These names are indexed by depth in the term tree,
        // which reconciliation walks — and that tree contains only permissioned tiers.
        // Leave the below-Unit entries in and every name shifts, so a missing
        // abbreviation gets reported against the wrong tier.
        const names = parseLevels(r.Levels ?? "").filter(isPermissioned).map((l) => l.label);
        if (names.length > 0) out.set(guid, names);
      });
    } catch {
      // A config hiccup costs the level NAMES, not the run — the fallback below
      // is correct for every segment onboarded so far.
    }
    return out;
  };

  // Segments to provision come from the SAME DMS Config `mode` rows the upload form
  // reads, so onboarding a segment is data-only (add a mode row → Run) — no redeploy.
  // Falls back to the built-in RECON_MODES (GHO) if the config is empty/unreachable.
  const loadReconModes = async (): Promise<Array<{ termSetGuid: string; stagingFolder: string }>> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items?$select=TermSetGuid,StagingFolder,Levels,SortOrder&$filter=ConfigType eq 'mode'&$orderby=SortOrder`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return RECON_MODES;
      const data = await res.json();
      const modes = parseReconModes((data.value ?? []) as RawModeRow[]);
      return modes.length > 0 ? modes : RECON_MODES;
    } catch {
      return RECON_MODES;
    }
  };

  // The Year × Document Type grid term sets come from the SAME DMS Config `setting`
  // rows the upload form reads (termSet_yearPeriod / termSet_documentType), so the
  // grid matches the form on any tenant with no code edit. Falls back to the built-in
  // YEAR_TERMSET / DOCTYPE_TERMSET constants per-key if the row or the list is missing.
  type ReconSettings = { year: string; docType: string; gridMode: GridMode; fanOut: boolean; revokeAncestorRead: boolean };
  const RECON_SETTINGS_FALLBACK: ReconSettings = {
    year: YEAR_TERMSET, docType: DOCTYPE_TERMSET, gridMode: DEFAULT_GRID_MODE, fanOut: false,
    revokeAncestorRead: false,
  };
  const loadReconGridTermSets = async (): Promise<ReconSettings> => {
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return RECON_SETTINGS_FALLBACK;
      const data = await res.json();
      const map: Record<string, string> = {};
      ((data.value ?? []) as Array<{ Title: string; SettingValue: string }>).forEach(
        (item) => { map[item.Title] = (item.SettingValue ?? "").trim(); },
      );
      const raw = (map.recon_gridMode || "").toLowerCase();
      const gridMode: GridMode =
        raw === "full" ? "full"
        : raw === "currentyear" ? "currentYear"
        : raw === "off" ? "off"
        : DEFAULT_GRID_MODE;
      return {
        year: map.termSet_yearPeriod || YEAR_TERMSET,
        docType: map.termSet_documentType || DOCTYPE_TERMSET,
        gridMode,
        // Departmental fan-out is OPT-IN, and defaults off even though the design
        // wants it on. The reason is historical data, not caution for its own sake:
        // until 2026-07-29 the old isChainAuthorized required a MEMBER row at EVERY
        // tier, so sites provisioned before then can still carry leftover
        // department-tier MEMBER rows. Fanning those would silently grant Read on
        // every unit folder in Documents — precisely the cross-unit leak the
        // isolation model exists to prevent, applied to rows nobody remembers
        // creating.
        //
        // So the run REPORTS what would fan while this is off (see the assignment
        // loop) and grants nothing. An admin reads that list, deletes the leftovers,
        // and only then sets the row to "on".
        // Defaults ON as of 2026-08-07. Head of Department is a department-scoped persona,
        // and unit folders have unique permissions, so with this off an HoD row grants Read
        // on the department folder and nothing else — a folder that appears, to them, to
        // contain no units. That is the persona not working rather than working narrowly.
        // An explicit "off" still turns it off.
        fanOut: (map.recon_departmentFanOut || "on").toLowerCase() === "on",
        // SUPERSEDED 2026-08-07 — kept for the history, because the reasoning below is why
        // the field still exists rather than being deleted outright. The client withdrew the
        // rule it served; ancestor Read is granted again on every run, and this is hard-off.
        //
        // Original note: Ancestor browse Read is no longer granted at all: the client's rule
        // is that a Head of Unit cannot see the department folder and a Head of Department
        // cannot see the segment folder. Only a C-level (segment-tier) row sees from
        // the segment down.
        //
        // But reconciliation has only ever ADDED assignments, so every site already
        // provisioned carries the ancestor Read grants made by earlier runs. Removing
        // the granting code fixes new folders and changes nothing on existing ones —
        // the requirement would read as met while every current user still saw their
        // parents, and the run log would report a clean pass.
        //
        // Hence a revoke pass, and hence it is OPT-IN. This is the first operation in
        // this codebase that deletes anything, so the default REPORTS every candidate
        // and removes none. An admin reads that list, satisfies themselves it contains
        // only what they expect, and then sets the row to "on".
        // HARD OFF as of 2026-08-07, config row ignored. The requirement it served was
        // withdrawn, and the ancestor Read it stripped is now granted again on every run —
        // so an admin flipping this row would have the two passes fight, and would break
        // in-library navigation for every user. Kept as a field rather than deleted so the
        // setting name stays recognisable if it reappears in a client's config list.
        revokeAncestorRead: false,
      };
    } catch {
      return RECON_SETTINGS_FALLBACK;
    }
  };

  // Reconciliation speed knobs, tunable live from DMS Config `setting` rows without a
  // redeploy: recon_writeDelayMs / recon_batchSize / recon_cooldownMs. Falls back to
  // the module defaults. Lets the client dial the delay down to find their tenant's
  // throttle floor (withThrottleRetry catches any 429 that slips through).
  const loadReconThrottle = async (): Promise<{ delayMs: number; batchSize: number; cooldownMs: number }> => {
    const fallback = { delayMs: RECON_WRITE_DELAY_MS, batchSize: RECON_BATCH_SIZE, cooldownMs: RECON_COOLDOWN_MS };
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return fallback;
      const data = await res.json();
      const map: Record<string, string> = {};
      ((data.value ?? []) as Array<{ Title: string; SettingValue: string }>).forEach(
        (item) => { map[item.Title] = (item.SettingValue ?? "").trim(); },
      );
      const num = (v: string | undefined, d: number): number => {
        const n = Number(v);
        return Number.isFinite(n) && n > 0 ? n : d;
      };
      return {
        delayMs: num(map.recon_writeDelayMs, fallback.delayMs),
        batchSize: num(map.recon_batchSize, fallback.batchSize),
        cooldownMs: num(map.recon_cooldownMs, fallback.cooldownMs),
      };
    } catch {
      return fallback;
    }
  };

  // Flatten the term store into the folders to provision, parent-before-child so a
  // parent always exists before we create its child. The segment container folder
  // (mode.stagingFolder) has no term; every term below it carries its GUID.
  /**
   * Enumerate every folder target from the term store.
   *
   * Also reports which segments could NOT be fully enumerated. This matters far more
   * than it looks: prune decides a row is orphaned by asking "is its term in the
   * enumerated set?", so a segment that silently came back empty would make every one
   * of its rows look deleted. Callers MUST treat a non-empty `incomplete` as a reason
   * to skip pruning.
   */
  const buildProvisionTargets = async (): Promise<{
    targets: ProvTarget[];
    incomplete: string[];
    missingAbbrev: UnclaimedTerm[];
    collisions: AbbrevCollision[];
    abbrevRows: OrphanAbbrevRow[];
  }> => {
    const out: ProvTarget[] = [];
    const incomplete: string[] = [];
    // Terms with no abbreviation are skipped, not guessed at, and reported here.
    // These are also the "unclaimed" half of an orphan repair: a re-created term
    // has a new GUID, so it arrives here looking like a brand-new term.
    const missingAbbrev: UnclaimedTerm[] = [];
    // Every named target, so siblings sharing an abbreviation can be caught before
    // a single folder is created.
    const abbrevTargets: AbbrevTarget[] = [];
    const abbrevRows = await loadAbbrevRows();
    const abbrevIndex = abbrevIndexOf(abbrevRows);
    const levelNamesBySet = await loadReconLevelNames();
    const modes = await loadReconModes();
    for (const mode of modes) {
      // Segment container: not a mapped term, but groups target it via the term-set GUID.
      // Its full name is the TERM SET's name, read live. mode.stagingFolder cannot serve
      // here — in DMS Config it already holds the abbreviation ("GHO"), which is exactly
      // the string Full Name exists to explain. Falls back to the folder name if the read
      // fails, so a term-store hiccup costs a label, not the run.
      const segmentFullName = (await loadTermSetName(mode.termSetGuid)) ?? mode.stagingFolder;
      // The term-set GUID IS now placed in every descendant's ancestorTerms (2026-08-09),
      // so a segment-tier row can reach the segment's folders — which is what the C-Level
      // "view one business segment" persona is.
      //
      // The original reason it was withheld still stands and is still enforced, just one
      // layer down instead of here: sites provisioned before 2026-07-29 can carry leftover
      // segment-tier MEMBER rows, and fanning one would grant Read across an entire
      // business segment. The fanned loop therefore accepts a segment-tier inheritance for
      // **SEGVIEW only** (see segmentTermSets there). Withholding the term entirely was the
      // blunter version of the same rule and made the intended role unusable.
      out.push({ termGuid: null, assignTerm: mode.termSetGuid, ancestorTerms: [], relPath: `/${mode.stagingFolder}`, label: mode.stagingFolder, fullName: segmentFullName, section: mode.stagingFolder, isLeaf: false, termSetGuid: mode.termSetGuid });
      // A failure anywhere in this segment's tree marks the WHOLE segment incomplete.
      // Targets gathered before the failure are kept (creating a subset of folders is
      // harmless and idempotent) — but prune must not run against a partial picture.
      try {
      const tops = await loadReconTops(mode.termSetGuid);
      for (const top of tops) {
        // Term labels are NOT safe as folder names. A "/" is the worst case — it is both
        // rejected by SharePoint (HTTP 400, SPException -2130575245) and read as a path
        // separator, so it silently implies an extra folder level. Real example: the
        // Minamas unit "Value Creation / Value Transformation".
        // Path segments are sanitised; the DISPLAY label and the map row Title keep the
        // raw term text, so the log and the index still read like the term store.
        // Folder names come from the abbreviation list, NOT the term label. The
        // labels are long and their fullwidth ampersands cost 9 encoded characters
        // each, which put the worst-case path within 73 characters of the ~330
        // limit where GetFolderByServerRelativeUrl starts returning 400.
        // See the 2026-07-30 folder-abbreviation-naming spec.
        //
        // sanitizeFolderSegment still applies: a "/" in a name is both rejected by
        // SharePoint and read as a path separator.
        //
        // A term with no abbreviation is SKIPPED and reported, never guessed at.
        // Falling back to the label would create a folder at a path the next run
        // does not expect, and uploads resolve by UniqueId so nobody would notice.
        //
        // `depth` is 1-based (top = the first level in the segment's Levels chain).
        // It resolves the level NAME the orphan-repair pass matches on, so a dead
        // "Legal" department can never be repaired from a live "Legal" unit.
        const levelNames =
          levelNamesBySet.get((mode.termSetGuid ?? "").trim().toLowerCase()) ??
          FALLBACK_LEVEL_NAMES;
        // How deep the PERMISSIONED tiers go — `levelNames` is already filtered to
        // them (see loadReconLevelNames). Caps the walk below.
        //
        // Until 2026-08-10 the walk recursed until a term had no children, which was
        // correct only because nothing was ever nested below Unit. The client's
        // SubUnit tier nests INSIDE the segment term set but INHERITS the unit's ACL,
        // so an uncapped walk would create an ACL'd folder per subunit carrying only
        // the owners group — invisible to the people who need it — and would stop
        // Units being leaves, moving the Year × Document Type grid onto subunits.
        // The term tree and the Levels chain now agree deliberately, not by accident.
        const permissionedDepth = Math.max(1, levelNames.length);
        const seg = (termGuid: string, label: string, depth: number): string | undefined => {
          const abbrev = lookupAbbrev(abbrevIndex, termGuid);
          if (abbrev === undefined) {
            missingAbbrev.push({
              termGuid,
              label,
              level: levelNames[depth - 1] ?? `Level ${depth}`,
            });
            return undefined;
          }
          return sanitizeFolderSegment(abbrev) || abbrev;
        };
        const topSeg = seg(top.id, top.label, 1);
        if (topSeg === undefined) continue; // reported; its children are unreachable
        abbrevTargets.push({ parentPath: `/${mode.stagingFolder}`, termGuid: top.id, abbreviation: topSeg, label: top.label });
        // ancestorTerms leads with the SEGMENT (the term-set GUID) since 2026-08-09, so a
        // segment-tier row reaches this department and everything under it. Restricted to
        // SEGVIEW in the fanned loop — see segmentTermSets there.
        const topTarget: ProvTarget = { termGuid: top.id, assignTerm: top.id, ancestorTerms: [mode.termSetGuid], relPath: `/${mode.stagingFolder}/${topSeg}`, label: `${mode.stagingFolder} > ${top.label}`, fullName: top.label, section: mode.stagingFolder, isLeaf: false, termSetGuid: mode.termSetGuid };
        out.push(topTarget);
        // Recurse; returns whether the term had children. A term with no children
        // is a leaf (the upload target) and gets the Year × Document Type grid.
        // `ancestors` carries raw labels for display, `pathAncestors` the sanitised
        // segments for the folder path — they can differ and must not be conflated.
        // `termAncestors` mirrors `ancestors` but carries term GUIDs rather than
        // labels, and INCLUDES top (which `ancestors` excludes) because a row on
        // the top tier fans down just like any other non-leaf row.
        const walk = async (
          parentId: string,
          ancestors: string[],
          pathAncestors: string[],
          termAncestors: string[],
        ): Promise<boolean> => {
          // STOP at the permissioned boundary. `ancestors` excludes top, so a direct
          // child of top is depth 2. Returning false (rather than skipping inside the
          // loop) is what makes the caller mark this term a LEAF — the deepest
          // permissioned folder, which is the upload target and where the grid hangs.
          // Anything deeper in the term tree is a below-Unit tier: created on demand
          // by the upload form, inheriting this folder's ACL, never provisioned here.
          if (ancestors.length + 2 > permissionedDepth) return false;
          const children = await loadReconChildren(mode.termSetGuid, parentId);
          for (const child of children) {
            // ancestors excludes `top`, so a direct child of top has depth 2.
            const childSeg = seg(child.id, child.label, ancestors.length + 2);
            if (childSeg === undefined) continue; // reported; skip this subtree
            const chain = [...ancestors, child.label];
            const pathChain = [...pathAncestors, childSeg];
            const termChain = [...termAncestors, child.id];
            const parentPath = `/${mode.stagingFolder}/${topSeg}${pathAncestors.length > 0 ? "/" + pathAncestors.join("/") : ""}`;
            abbrevTargets.push({ parentPath, termGuid: child.id, abbreviation: childSeg, label: child.label });
            const childTarget: ProvTarget = {
              termGuid: child.id,
              assignTerm: child.id,
              ancestorTerms: termAncestors,
              relPath: `/${mode.stagingFolder}/${topSeg}/${pathChain.join("/")}`,
              label: `${mode.stagingFolder} > ${top.label} > ${chain.join(" > ")}`,
              fullName: child.label,
              section: mode.stagingFolder,
              isLeaf: false,
              termSetGuid: mode.termSetGuid,
            };
            out.push(childTarget);
            childTarget.isLeaf = !(await walk(child.id, chain, pathChain, termChain));
          }
          return children.length > 0;
        };
        // termAncestors starts with the SEGMENT then the top term, so every descendant
        // inherits both. Outermost first, matching the ProvTarget contract.
        topTarget.isLeaf = !(await walk(top.id, [], [], [mode.termSetGuid, top.id]));
      }
      } catch (e) {
        incomplete.push(`${mode.stagingFolder} — ${(e as Error).message}`);
      }
    }
    return {
      targets: out,
      incomplete,
      missingAbbrev,
      collisions: findCollisions(abbrevTargets),
      abbrevRows,
    };
  };

  // Provision from the term store into BOTH libraries in one run: create each folder
  // if missing and break inheritance at every level (segment → department → unit).
  // Owner group is re-added as Full Control so admins keep access. Then AUTO-ASSIGN
  // the DMS Group Map groups to each folder by role. Staging gets UPL→Contribute and
  // APR→Design only (MEMBER is Documents-only for isolation); Documents gets MEMBER→Read
  // only; GLOBAL is skipped everywhere. Staging term folders are also mapped (term → UniqueId) for
  // rename-proof upload routing; Documents is not mapped. Idempotent: folders/maps
  // are not duplicated; role assignments merge (re-adding an existing one is a no-op),
  // and manual extra grants survive. Folders whose term has no group-map rows are
  // logged as a warning (locked admin-only until groups are added).
  const runReconciliation = async (): Promise<void> => {
    setReconConfirm(false);
    setBusy(true);
    setReconRunning(true);
    setFolderFeeds(emptyFeeds());
    setAssignFeeds(emptyFeeds());
    setReconCounts({ folders: 0, assigns: 0 });
    setReconDone(0);
    setReconPlanned(0);
    setReconStartMs(undefined);
    setReconPhase("Generating folders…");
    const entries: LogEntry[] = [];
    // Throttle governor: pause between writes, and auto-cooldown every N writes. The
    // cooldown is invisible — the spinner keeps running and the status text rotates so
    // it reads as continuous work (no user click needed to resume). Speed knobs come
    // from DMS Config (recon_writeDelayMs / recon_batchSize / recon_cooldownMs) so the
    // client can tune without a redeploy; withThrottleRetry catches any 429 that slips.
    const throttle = await loadReconThrottle();
    let opCount = 0;
    const tick = async (): Promise<void> => {
      opCount++;
      await sleep(throttle.delayMs);
      if (opCount % throttle.batchSize === 0) {
        for (const msg of COOLDOWN_MESSAGES) {
          setReconPhase(msg);
          await sleep(throttle.cooldownMs / COOLDOWN_MESSAGES.length);
        }
        setReconPhase("Generating folders…");
      }
    };
    const bumpFolders = (): void => setReconCounts((c) => ({ ...c, folders: c.folders + 1 }));
    const bumpAssigns = (): void => setReconCounts((c) => ({ ...c, assigns: c.assigns + 1 }));
    try {
      const fullCtrlId = roleDefs.find(r => r.name === "Full Control")?.id;
      const readId = roleDefs.find(r => r.name === "Read")?.id;
      // Read now serves two purposes: recognising a leftover ancestor browse grant (they
      // are all exactly Read) and granting site entry below. Without it the run still
      // provisions folders correctly — it just cannot do either of those.
      if (readId === undefined) entries.push({ msg: `⚠ "Read" role definition not found — site entry and ancestor-browse cleanup will be skipped`, ok: false });

      // ── Site entry, FIRST ──────────────────────────────────────────────────────
      //
      // Find or create DMS_SITE_MEMBERS and give it Read on the web. Until now this was a
      // sentence in a runbook ("Set up site entry"), and the only thing that noticed it had
      // been skipped was a warning at the very END of the run — by which point every folder
      // was already locked. A step that exists only in prose gets skipped.
      //
      // Order is not cosmetic. A folder grant alone confers Limited Access: the user can
      // open that folder by direct link, but the site root denies them, so they cannot
      // reach anything by navigating. Lock folders first and grant entry afterwards and
      // there is a window in which a correctly provisioned uploader can reach nothing at
      // all — and if the run dies in that window, that is the state the site is left in.
      // Granting entry first means the worst case is a user who can open the site and sees
      // nothing yet, which the next run resolves.
      //
      // See the access-scope-mapping spec §4 and the site-entry-access-layer spec.
      try {
        setReconPhase("Ensuring site entry…");
        // Find-or-create through the shared module (siteEntryGroup.ts). It THROWS rather than
        // creating when the group list cannot be read — creating on an unreadable list would make
        // a second entry group alongside the real one, both looking correct. The catch below
        // already treats that as non-fatal and names it in the log.
        const ensured = await ensureSiteEntryGroup(context.spHttpClient, siteUrl);
        if (ensured.created) entries.push({ msg: `${siteEntryGroupTitle()} created`, ok: true });
        const entryId = ensured.group.id;
        if (readId === undefined) {
          entries.push({ msg: `⚠ ${siteEntryGroupTitle()}: cannot grant site Read — no "Read" role definition`, ok: false });
        } else {
          // The root web always has unique permissions, so there is no inheritance to break
          // here — unlike a library or a page. Grant directly.
          //
          // Checked before granting rather than leaning on addroleassignment being
          // idempotent, purely so the log distinguishes "already had it" from "granted
          // now". Both outcomes are fine; only one of them is news.
          const webRas = await context.spHttpClient.get(
            `${siteUrl}/_api/web/roleassignments?$select=PrincipalId`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          let holds = false;
          if (webRas.ok) {
            const raJson = await webRas.json();
            holds = ((raJson.value ?? []) as Array<{ PrincipalId?: number }>)
              .some((ra) => ra.PrincipalId === entryId);
          }
          if (holds) {
            entries.push({ msg: `${siteEntryGroupTitle()}: already holds a role on the site ✓`, ok: true });
          } else {
            const grant = await withThrottleRetry(() => context.spHttpClient.post(
              `${siteUrl}/_api/web/roleassignments/addroleassignment(principalid=${entryId},roledefid=${readId})`,
              SPHttpClient.configurations.v1,
              { headers: { Accept: "application/json;odata=nometadata" } },
            ));
            entries.push(grant.ok
              ? { msg: `${siteEntryGroupTitle()} → Read on the site ✓`, ok: true }
              : { msg: `⚠ ${siteEntryGroupTitle()} → Read on the site FAILED (HTTP ${grant.status}) — users will reach folders by direct link only`, ok: false });
          }
        }
      } catch (e) {
        // Never fatal. Folder provisioning is still worth doing, and the failure is named
        // rather than swallowed so it is not mistaken for a folder problem.
        entries.push({ msg: `⚠ Site entry could not be ensured — ${(e as Error).message}`, ok: false });
      }

      // ── Library-scope grants ───────────────────────────────────────────────────
      //
      // A Group Map row with Scope = Library grants its role on a whole LIBRARY rather
      // than on a unit folder. Target holds the library title. See the
      // access-scope-mapping spec.
      //
      // Runs after site entry and before the folder passes: breaking a library's
      // inheritance resets what flows down to its folders, so doing it after the folders
      // were locked would mean the next run is the first one that is actually correct.
      try {
        setReconPhase("Applying library access…");
        // Read the scoped rows directly. loadGroupMapForAssign keys by term and drops
        // Scope/Target, which a library row does not have and does not need.
        //
        // Scope/Target are newer than this list, and a $select naming a column that does
        // not exist fails the WHOLE request with 400 — so a site without the columns must
        // not lose its folder provisioning over it. Absent columns simply mean no library
        // rows exist yet.
        const scopedRes = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items?$select=GroupId,GroupName,Role,Scope,Target&$top=5000`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!scopedRes.ok) {
          entries.push({ msg: `Library access: skipped — Scope/Target columns not present on DMS Group Map`, ok: true });
        } else {
          const scopedJson = await scopedRes.json();
          const libRows = ((scopedJson.value ?? []) as Array<{ GroupId?: string; GroupName?: string; Role?: string; Scope?: string; Target?: string }>)
            .filter((r) => (r.Scope ?? "").trim().toLowerCase() === "library")
            .filter((r) => (r.Target ?? "").trim() !== "" && (r.GroupId ?? "") !== "");
          if (libRows.length === 0) {
            entries.push({ msg: `Library access: no Library-scope mappings`, ok: true });
          } else {
            // Needed to re-grant after a break — see below. Resolved here rather than
            // threaded out of the site-entry pass so this block stands alone.
            const groupsNow = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
            const entryPid = findSiteEntryGroup(groupsNow)?.id;
            const brokenThisRun = new Set<string>();
            for (const row of libRows) {
              const lib = (row.Target ?? "").trim();
              const role = normalizeRoleValue(row.Role ?? "");
              // Library scope, but resolved the same way as folder scope on purpose. A
              // hand-written UPL row targeting Documents would otherwise grant CRS Upload at
              // the LIBRARY ROOT — write access to every folder that inherits. Downgrading it
              // to Read here costs nothing for the rows that belong at this scope (ENTRY is
              // Read either way) and closes that off.
              const levelName = permissionForRole(lib as LibTarget, role);
              const roleDefId = roleDefs.find((r) => r.name === levelName)?.id;
              const label = `${row.GroupName || row.GroupId} → ${lib}`;
              if (levelName === undefined) {
                entries.push({ msg: `  ⚠ ${label}: role "${role}" grants nothing (retired or unknown) — skipped`, ok: false });
                continue;
              }
              if (roleDefId === undefined) {
                entries.push({ msg: `  ⚠ ${label}: no "${levelName}" role definition on site — skipped`, ok: false });
                continue;
              }
              const listRes = await context.spHttpClient.get(
                `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')?$select=Id,HasUniqueRoleAssignments`,
                SPHttpClient.configurations.v1,
                { headers: { Accept: "application/json;odata=nometadata" } },
              );
              if (!listRes.ok) {
                entries.push({ msg: `  ⚠ ${label}: library "${lib}" not found — check the Target value`, ok: false });
                continue;
              }
              const listJson = await listRes.json();
              const listId: string = listJson.Id;
              const listBase = `${siteUrl}/_api/web/lists(guid'${listId}')`;
              // Did the site-entry group have access to this library BEFORE we touched it?
              //
              // This decides whether it is restored after the break, and it is asked rather
              // than assumed. Documents MUST keep it — the approval guard resolves the
              // destination folder as the approver and depends on that Read. Staging must
              // NOT gain it: site entry means "can open the site", and a plain member who is
              // neither uploader nor approver has no business reaching Staging at all.
              //
              // Deciding by library NAME would bake that into a constant and be wrong the
              // moment a library is renamed or a third one appears. Preserving whatever was
              // already in force cannot be wrong about either.
              let entryHadAccess = false;
              if (entryPid !== undefined) {
                const before = await context.spHttpClient.get(
                  `${listBase}/roleassignments?$select=PrincipalId`,
                  SPHttpClient.configurations.v1,
                  { headers: { Accept: "application/json;odata=nometadata" } },
                );
                if (before.ok) {
                  const bj = await before.json();
                  entryHadAccess = ((bj.value ?? []) as Array<{ PrincipalId?: number }>)
                    .some((ra) => ra.PrincipalId === entryPid);
                }
              }
              if (listJson.HasUniqueRoleAssignments !== true && !brokenThisRun.has(listId)) {
                // copyRoleAssignments=false, always. With true, every inherited grant is
                // copied forward, so the library stays visible to exactly the same people
                // and the run reports success — a failure that is invisible from the log.
                const broke = await withThrottleRetry(() => context.spHttpClient.post(
                  `${listBase}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
                  SPHttpClient.configurations.v1,
                  { headers: { Accept: "application/json;odata=nometadata" } },
                ));
                if (!broke.ok) {
                  entries.push({ msg: `  ✗ ${lib}: could not break inheritance (HTTP ${broke.status}) — nothing granted`, ok: false });
                  continue;
                }
                brokenThisRun.add(listId);
                entries.push({ msg: `  ↳ ${lib}: inheritance broken (no permissions copied)`, ok: true });
                // Two principals go back on, and BOTH are load-bearing.
                //
                // Owners: with nothing copied, the only remaining access is site collection
                // administrators. An owner who is not also a site collection admin would
                // lose the library.
                if (fullCtrlId !== undefined) {
                  try {
                    await addRoleAssignmentToList(listBase, ownerGroupId as number, fullCtrlId);
                    entries.push({ msg: `  ↳ ${lib}: site Owners → Full Control restored`, ok: true });
                  } catch (e) {
                    entries.push({ msg: `  ✗ ${lib}: could not restore site Owners — ${(e as Error).message}`, ok: false });
                  }
                }
                // Site entry: the approval guard resolves the destination folder in
                // Documents AS THE APPROVER, and that works only because the site-entry
                // group holds Read on the library by INHERITANCE. Breaking inheritance
                // discards it, and every approver on the site then 404s on every
                // destination folder — approval refused for everyone, with a log that says
                // the run succeeded. See the access-scope-mapping spec §5.
                if (!entryHadAccess) {
                  // It did not have access before, so it does not get any now. Logged rather
                  // than silent: on Staging this is the correct and intended outcome, and an
                  // unexplained absence here would look like the restore had failed.
                  entries.push({ msg: `  ↳ ${lib}: ${siteEntryGroupTitle()} had no access before — not granted (site entry is not library access)`, ok: true });
                } else if (entryPid !== undefined && readId !== undefined) {
                  try {
                    await addRoleAssignmentToList(listBase, entryPid, readId);
                    entries.push({ msg: `  ↳ ${lib}: ${siteEntryGroupTitle()} → Read restored (keeps approval working)`, ok: true });
                  } catch (e) {
                    entries.push({ msg: `  ✗ ${lib}: could not restore ${siteEntryGroupTitle()} — approvals may fail — ${(e as Error).message}`, ok: false });
                  }
                } else {
                  entries.push({ msg: `  ⚠ ${lib}: ${siteEntryGroupTitle()} or "Read" not resolved — approvals may fail until it holds Read here`, ok: false });
                }
              }
              try {
                await addRoleAssignmentToList(listBase, spGroupPrincipalId(row.GroupId ?? ""), roleDefId);
                entries.push({ msg: `  ↳ ${label} → ${levelName} (library)`, ok: true });
                pushAssign("Documents", `${lib} → ${row.GroupName} (${levelName}, library scope)`, "ok");
                bumpAssigns();
                await tick();
              } catch (e) {
                entries.push({ msg: `  ✗ ${label} → ${levelName} FAILED: ${(e as Error).message}`, ok: false });
              }
            }
          }
        }
      } catch (e) {
        entries.push({ msg: `⚠ Library access skipped — ${(e as Error).message}`, ok: false });
      }

      // ── PAGE ACCESS PASS ──────────────────────────────────────────────────────
      // Re-asserts every `Scope = Page` row, so page access self-heals the same way folder
      // and library access do.
      //
      // The client's rule, 2026-08-05: page and library access are DERIVED from the role, not
      // curated per page. An uploader has the upload form because they are an uploader — so a
      // grant removed by hand in SharePoint is drift to be repaired, not a decision to respect.
      // The supported way to revoke is to remove the mapping (which deletes the row and the
      // grant together) or to delete the group.
      //
      // That is a deliberate reversal of the "curated state should not self-heal" argument, and
      // it holds only because the row IS the record of the role. If page grants ever become
      // hand-curated exceptions, this pass has to become report-only.
      try {
        setReconPhase("Applying page access…");
        const pgRes = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items?$select=GroupId,GroupName,Role,Scope,Target&$top=5000`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
        if (!pgRes.ok) {
          entries.push({ msg: `Page access: skipped — Scope/Target columns not present on DMS Group Map`, ok: true });
        } else {
          const pgJson = await pgRes.json();
          const pageRows = ((pgJson.value ?? []) as Array<{ GroupId?: string; GroupName?: string; Role?: string; Scope?: string; Target?: string }>)
            .filter((r) => (r.Scope ?? "").trim().toLowerCase() === "page")
            .filter((r) => (r.Target ?? "").trim() !== "" && (r.GroupId ?? "") !== "");
          if (pageRows.length === 0) {
            entries.push({ msg: `Page access: no Page-scope mappings`, ok: true });
          } else if (readId === undefined) {
            entries.push({ msg: `⚠ Page access: no "Read" role definition on site — skipped`, ok: false });
          } else {
            const PAGES_LIST = "Site Pages";
            const pagesBase = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(PAGES_LIST)}')`;
            // The site's real welcome page, not just the "home.aspx" constant: a renamed welcome
            // page would slip past the constant, and locking it makes the whole site unreachable
            // for everyone who is not an administrator.
            let welcome = "";
            try {
              const wRes = await context.spHttpClient.get(
                `${siteUrl}/_api/web/RootFolder?$select=WelcomePage`,
                SPHttpClient.configurations.v1,
                { headers: { Accept: "application/json;odata=nometadata" } },
              );
              if (wRes.ok) {
                const wj = await wRes.json();
                welcome = ((wj.WelcomePage ?? "") as string).split("/").pop()?.toLowerCase() ?? "";
              }
            } catch { /* fall back to the constant alone */ }

            const itemsRes = await context.spHttpClient.get(
              `${pagesBase}/items?$select=Id,FileLeafRef,HasUniqueRoleAssignments&$top=500`,
              SPHttpClient.configurations.v1,
              { headers: { Accept: "application/json;odata=nometadata" } },
            );
            if (!itemsRes.ok) {
              entries.push({ msg: `⚠ Page access: could not read ${PAGES_LIST} (HTTP ${itemsRes.status}) — skipped`, ok: false });
            } else {
              const itemsJson = await itemsRes.json();
              const items = ((itemsJson.value ?? []) as Array<{ Id: number; FileLeafRef?: string; HasUniqueRoleAssignments?: boolean }>)
                .map((p) => ({ id: p.Id, file: (p.FileLeafRef ?? "").toLowerCase(), unique: p.HasUniqueRoleAssignments === true }));

              // Grouped by page so inheritance is broken ONCE per page rather than once per row.
              const byPage = new Map<string, typeof pageRows>();
              for (const r of pageRows) {
                const key = (r.Target ?? "").trim().toLowerCase();
                byPage.set(key, [...(byPage.get(key) ?? []), r]);
              }

              for (const [file, rowsForPage] of Array.from(byPage.entries())) {
                if (isForbiddenPageTarget(file) || file === welcome) {
                  // Refused by name, every run, rather than applied once and regretted. The row
                  // is left in place: deleting authored data on the client's behalf is not this
                  // pass's job, and the refusal is logged so it can be removed deliberately.
                  entries.push({ msg: `  ⚠ ${file}: the site home page cannot be restricted — ${rowsForPage.length} mapping(s) refused`, ok: false });
                  continue;
                }
                const item = items.find((p) => p.file === file);
                if (!item) {
                  entries.push({ msg: `  ⚠ ${file}: page not found in ${PAGES_LIST} — check the Target value`, ok: false });
                  continue;
                }
                const itemBase = `${pagesBase}/items(${item.id})`;
                if (!item.unique) {
                  // copyRoleAssignments=false, as everywhere else: with true, every inherited
                  // grant is carried forward, so the page stays visible to exactly the same
                  // people and the run reports success.
                  const broke = await withThrottleRetry(() => context.spHttpClient.post(
                    `${itemBase}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
                    SPHttpClient.configurations.v1,
                    { headers: { Accept: "application/json;odata=nometadata" } },
                  ));
                  if (!broke.ok) {
                    entries.push({ msg: `  ✗ ${file}: could not break inheritance (HTTP ${broke.status}) — nothing granted`, ok: false });
                    continue;
                  }
                  entries.push({ msg: `  ↳ ${file}: inheritance broken (no permissions copied)`, ok: true });
                  // typeof, not !== undefined: ownerGroupId is `number | null` when the
                  // associated owner group could not be resolved, and null would slip past an
                  // undefined check straight into the request as "null".
                  if (fullCtrlId !== undefined && typeof ownerGroupId === "number") {
                    try {
                      await addRoleAssignmentToList(itemBase, ownerGroupId, fullCtrlId);
                      entries.push({ msg: `  ↳ ${file}: site Owners → Full Control restored`, ok: true });
                    } catch (e) {
                      entries.push({ msg: `  ✗ ${file}: could not restore site Owners — ${(e as Error).message}`, ok: false });
                    }
                  }
                }
                for (const row of rowsForPage) {
                  const label = `${row.GroupName || row.GroupId} → ${file}`;
                  // ALWAYS Read, never the row's own level. A page is opened or it is not, and
                  // granting "DMS Upload" on a page item would be a meaningless binding that
                  // reads, in the permissions UI, like an upload right on the page.
                  if (normalizeRoleValue(row.Role ?? "") !== "ENTRY") {
                    entries.push({ msg: `  ⚠ ${label}: role "${row.Role}" on a Page row — granting Read (page access is Read by definition)`, ok: true });
                  }
                  try {
                    await addRoleAssignmentToList(itemBase, spGroupPrincipalId(row.GroupId ?? ""), readId);
                    entries.push({ msg: `  ↳ ${label} → Read (page)`, ok: true });
                    pushAssign("Documents", `${file} → ${row.GroupName} (Read, page scope)`, "ok");
                    bumpAssigns();
                    await tick();
                  } catch (e) {
                    entries.push({ msg: `  ✗ ${label} FAILED: ${(e as Error).message}`, ok: false });
                  }
                }
              }
            }
          }
        }
      } catch (e) {
        entries.push({ msg: `⚠ Page access skipped — ${(e as Error).message}`, ok: false });
      }

      // Ancestor rel-paths of a target, excluding the folder itself and the library root.
      // e.g. "/A/B/C" -> ["/A", "/A/B"]. Used to grant each unit group Read up its own path
      // so members can browse down to their folder; siblings without a grant stay
      // security-trimmed (invisible) in the view.
      const ancestorRelPaths = (relPath: string): string[] => {
        const parts = relPath.split("/").filter(Boolean);
        const out: string[] = [];
        let cur = "";
        for (let i = 0; i < parts.length - 1; i++) { cur += `/${parts[i]}`; out.push(cur); }
        return out;
      };
      // Full rows, not just term GUIDs: the presence of a row is NOT proof the row is
      // still correct, so each one gets verified below. See the folder-map-integrity spec.
      const mapRows = await loadFolderMapRows(context.spHttpClient, siteUrl);
      const mapByTerm = new Map<string, FolderMapRow>();
      for (const r of mapRows) {
        if (r.termGuid) mapByTerm.set(r.termGuid.toLowerCase(), r);
      }
      // Each folder's name BEFORE this run, snapshotted once.
      //
      // Read live from the row instead and the Staging pass — which rewrites
      // FolderUrl after a successful rename — would make the later Documents pass
      // believe the name was already correct. Documents would skip the rename and
      // the create step would make an empty folder at the new name, stranding the
      // real one under the old name. Snapshotting keeps every library renaming
      // from the same origin regardless of pass order.
      const oldNameByTerm = new Map<string, string>();
      for (const r of mapRows) {
        const name = (r.folderUrl ?? "").split("/").pop() ?? "";
        if (r.termGuid && name) oldNameByTerm.set(r.termGuid.toLowerCase(), name);
      }
      const groupMap = await loadGroupMapForAssign();
      const { targets, incomplete: incompleteSegments, missingAbbrev, collisions, abbrevRows } = await buildProvisionTargets();
      // The segment tier, identified by term-set GUID. Derived from the targets already in
      // hand — a segment container is the one target with no term of its own — rather than
      // re-reading the modes, so the two can never disagree about what "a segment" is.
      // Used by the fan-down to let SEGVIEW, and only SEGVIEW, inherit from this tier.
      const segmentTermSets = new Set<string>(
        targets.filter((t) => t.termGuid === null).map((t) => t.assignTerm.toLowerCase()),
      );
      // relPath → the term that folder stands for. Every ancestor folder is itself a
      // target (the segment folder and each department folder both get one), so this
      // covers the whole tree. The ancestor-read revoke needs it to answer the one
      // question that separates a leftover browse grant from a legitimate one: does
      // this group hold a Group Map row AT this tier? A department-tier viewer's Read
      // on its own department folder must survive; a unit group's Read on that same
      // folder must not.
      const termByRelPath = new Map<string, string>();
      for (const t of targets) termByRelPath.set(t.relPath, (t.assignTerm ?? "").toLowerCase());
      // Role assignments per ancestor path, fetched once. A twelve-unit department
      // would otherwise re-read the same department folder twelve times per library.
      const ancAssignCache = new Map<string, ExistingAssign[]>();
      // Ancestor browse grants MADE by this run, keyed `${ancestorFullPath}|${principalId}`.
      // Kept apart from ancAssignCache rather than appended to it: an ExistingAssign carries
      // uid/title/kept read back from SharePoint, and inventing those to represent a grant we
      // just made would put fabricated data into a structure other passes read as truth.
      // Twelve units under one department means eleven repeat grants without this.
      const ancGranted = new Set<string>();
      // A collision aborts BEFORE anything is created. Two siblings resolving to
      // one path means one folder, one ACL, and two units' documents inside it —
      // the isolation the whole permission model rests on. A partial run would
      // create that merged folder before anyone read the log.
      if (collisions.length > 0) {
        for (const c of collisions) {
          entries.push({ msg: `✖ COLLISION in ${c.parentPath}: "${c.abbreviation}" is used by ${c.labels.join(" | ")}`, ok: false });
        }
        entries.push({ msg: `Nothing was created. Give each of these a distinct abbreviation in ${abbrevListTitle()}, then run again.`, ok: false });
        setLog((prev) => [...prev, ...entries]);
        showToast(`${collisions.length} abbreviation collision(s) — nothing was created.`, false);
        setBusy(false);
        return;
      }
      // Not fatal: every other term still provisions. But it must be loud, because
      // a unit with no folder has no map row and its uploaders are blocked.
      for (const m of missingAbbrev) {
        entries.push({ msg: `⚠ SKIPPED (no abbreviation): ${m.label} — add a row to ${abbrevListTitle()} for term ${m.termGuid}`, ok: false });
      }
      // Surface enumeration failures as real errors. Without this they were invisible:
      // the segment just produced no targets, the run looked clean, and prune would
      // then delete every map row for it. See the prune guard below.
      for (const seg of incompleteSegments) {
        entries.push({ msg: `⚠ could not fully read the term store for ${seg} — folders may be missing and pruning is disabled for this run`, ok: false });
      }
      if (targets.length === 0) {
        showToast("No terms found in the term store to provision.", false);
        setBusy(false);
        return;
      }
      // Year × Document Type grid labels (from their term sets) — pre-created under
      // each leaf/unit folder. Both are flat term sets, so top-level terms suffice.
      // Sanitize labels to safe folder names with the SAME helper the upload form uses
      // (formModel.sanitizeFolderSegment), so reconciliation and Form.tsx always agree on
      // the folder name — e.g. a Document Type term containing illegal chars like "/".
      const gridSets = await loadReconGridTermSets();
      // State the fan-out mode up front. Off is the default and is the state in
      // which a Head of Department silently does not work, so it must be visible
      // at the top of the log rather than inferred from the absence of grants.
      entries.push({
        msg: gridSets.fanOut
          ? `Departmental fan-out: ON — a mapping on a department reaches every unit beneath it`
          : `Departmental fan-out: off (recon_departmentFanOut) — parent-tier mappings are reported, not granted`,
        ok: true,
      });
      // Unconditional since 2026-08-07: there is no longer an "off" state to report. The
      // banner used to describe the revoke pass, which is hard-off — so on a run that was
      // busily GRANTING ancestor Read it announced "nothing is removed", which is true and
      // entirely beside the point. Stated positively, it now matches the ↳ lines below it.
      entries.push({
        msg: `Ancestor browse Read: GRANTING — each group gets Read up its own path so it can browse down to its folder; siblings stay security-trimmed`,
        ok: true,
      });
      let yearLabels: string[] = [];
      let docTypeLabels: string[] = [];
      // Term set GUID → the folder names of each below-Unit tier, in path order.
      // One entry per SEGMENT, because each mode configures its own chain.
      const gridTiersBySet = new Map<string, string[][]>();
      const onDemandTiers = await loadReconOnDemandTiers();
      if (gridSets.gridMode === "off") {
        // Nothing to pre-create. Year/Document Type folders are made on demand by the
        // upload form, and in Documents by the Auto-route flow as approved files land.
        entries.push({ msg: `Year × Document Type grid: skipped (recon_gridMode = off) — folders are created on first use`, ok: true });
      } else {
        yearLabels    = (await loadReconTops(gridSets.year).catch(() => [] as TermLite[])).map(y => sanitizeFolderSegment(y.label)).filter(Boolean);
        docTypeLabels = (await loadReconTops(gridSets.docType).catch(() => [] as TermLite[])).map(d => sanitizeFolderSegment(d.label)).filter(Boolean);
        if (gridSets.gridMode === "currentYear" && yearLabels.length > 0) {
          const thisYear = String(new Date().getFullYear());
          const match = yearLabels.filter((y) => y === thisYear);
          // No term for the current year (e.g. the client has not added 2027 yet) — fall
          // back to the LAST year in the set rather than silently building all of them.
          yearLabels = match.length > 0 ? match : yearLabels.slice(-1);
        }
        entries.push({
          msg: `Year × Document Type grid: ${gridSets.gridMode} — ${yearLabels.length} year(s) × ${docTypeLabels.length} document type(s) per unit`,
          ok: true,
        });

        // Per-segment below-Unit chains. A mode with none configured keeps the two
        // legacy arrays above, so a site migrates one mode row at a time instead of
        // all at once — and a half-configured site still builds a complete path.
        const tierLabelCache = new Map<string, string[]>();
        const labelsForSet = async (setGuid: string): Promise<string[]> => {
          const key = (setGuid ?? "").trim().toLowerCase();
          if (!key) return [];
          const hit = tierLabelCache.get(key);
          if (hit) return hit;
          const tops: TermLite[] = await loadReconTops(setGuid).catch(() => [] as TermLite[]);
          const labels = tops.map((tl: TermLite) => sanitizeFolderSegment(tl.label)).filter(Boolean);
          tierLabelCache.set(key, labels);
          return labels;
        };
        // forEach into an array first: the SPFx tsconfig does not target ES2015, so
        // `for…of` over a Map is a compile error (same family as gotcha #3 in CLAUDE.md).
        const onDemandEntries: Array<{ setGuid: string; tiers: Level[] }> = [];
        onDemandTiers.forEach((tiers, setGuid) => onDemandEntries.push({ setGuid, tiers }));
        for (const { setGuid, tiers } of onDemandEntries) {
          const resolved: string[][] = [];
          for (const tier of tiers) {
            let labels = await labelsForSet(tier.termSet ?? "");
            // currentYear is a property of the YEAR set, not of a fixed position in
            // the chain — so it follows the term set wherever the admin puts the tier.
            if (
              gridSets.gridMode === "currentYear" &&
              (tier.termSet ?? "").trim().toLowerCase() === (gridSets.year ?? "").trim().toLowerCase() &&
              labels.length > 0
            ) {
              const thisYear = String(new Date().getFullYear());
              const match = labels.filter((y) => y === thisYear);
              labels = match.length > 0 ? match : labels.slice(-1);
            }
            resolved.push(labels);
          }
          const plan = gridPlan(resolved);
          gridTiersBySet.set(setGuid, plan.tiers);
          entries.push({
            msg:
              `Below-Unit structure for this segment: ${tiers.map((t) => t.label).join(" → ")} ` +
              `— ${plan.total} folder(s) per unit` +
              (plan.tiers.length < tiers.length
                ? `; "${tiers[plan.tiers.length].label}" has no terms, so nothing below it is pre-created`
                : ""),
            ok: true,
          });
        }
      }

      /**
       * The below-Unit tiers for a target's own segment, falling back to the legacy
       * Year → Document Type pair. Resolved per target, never once per run: two
       * segments may configure different shapes.
       */
      const gridTiersFor = (t: ProvTarget): string[][] =>
        gridTiersBySet.get((t.termSetGuid ?? "").trim().toLowerCase()) ??
        gridPlan([yearLabels, docTypeLabels]).tiers;
      // Estimate the workload up front: count every throttled op (each incurs the
      // inter-write delay). Structural folder (1) + Year×DocType grid per leaf +
      // applicable group grants per lib. Ancestor browse grants aren't throttled, so
      // they're excluded — the live rate absorbs their real time. Worst case (assumes
      // nothing exists yet); re-runs finish faster as existing folders skip.
      // Per SEGMENT, not per run — each mode may configure a different chain, and the
      // estimate divides elapsed time by ops completed, so counting a different set
      // here than the build loop attempts is what makes the "time left" figure lie.
      const gridPerLeafFor = (t: ProvTarget): number => gridPlan(gridTiersFor(t)).total;
      // Resolved once per library, before the estimate, because whether the column
      // exists changes the op count. Absent on a library = that library gets no full
      // names and is told so once, rather than once per folder.
      const fullNameFields = new Map<LibTarget, string>();
      const folderCtIds = new Map<LibTarget, string>();
      /** Libraries with content approval on — the only ones a moderation status may be written to. */
      const moderatedLibs = new Set<LibTarget>();
      for (const lib of ["Staging", "Documents"] as LibTarget[]) {
        const ct = await loadFolderContentTypeId(lib);
        if (ct) folderCtIds.set(lib, ct);
        // Names both candidates: "no CRS Folder content type" on a site that still has the
        // DMS-named one would read as a missing artefact rather than a rename half-done.
        else entries.push({ msg: `⚠ ${lib}: no ${FOLDER_CONTENT_TYPE_CANDIDATES.map(n => `"${n}"`).join(" or ")} content type — folders keep the built-in Folder type and the details pane will not show Full Name`, ok: true });
        const f = await loadFullNameField(lib);
        if (f.internalName) fullNameFields.set(lib, f.internalName);
        // ok:true deliberately. This is a warning, not an error: `errorsBeforePrune`
        // counts !ok entries and blocks the orphan prune, and that guard exists because
        // a partial TERM STORE read returns a short target list that makes healthy map
        // rows look deleted. A missing display column cannot shorten the target list, so
        // gating prune on it would silently disable self-healing over a cosmetic column.
        // The ⚠ still puts it in "Needs attention" where an admin will see it.
        else entries.push({ msg: `⚠ ${lib}: ${f.note ?? "no Full Name column"} — folders will show only their abbreviation`, ok: true });

        // Does this library moderate? Folders created in a content-approval library arrive
        // PENDING (verified live 2026-08-07: every folder in Approval Document was status 2),
        // and a pending FOLDER is hidden from anyone who cannot see drafts. Today that is
        // nobody, because Draft Item Security is "any user who can read items" — but the
        // moment it is tightened to approver-only for per-uploader isolation, every folder
        // vanishes for every non-approver and the library renders empty. Same failure as
        // memory dms-content-approval-blocks-uploader, reached from a different direction.
        //
        // So approve the folders as they are provisioned. Files are untouched: they are the
        // things actually under review, and approving them here would defeat the whole point.
        //
        // SCOPE: structural folders (segment / department / unit) ONLY — NOT the Year ×
        // Document Type grid under each unit. See the note at the grid loop for why, and why
        // Draft Item Security must therefore stay at "any user who can read items".
        try {
          const modRes: SPHttpClientResponse = await context.spHttpClient.get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')?$select=EnableModeration`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (modRes.ok && (await modRes.json()).EnableModeration === true) {
            moderatedLibs.add(lib);
            entries.push({ msg: `${lib}: content approval is on — provisioned folders will be approved so they stay visible`, ok: true });
          }
        } catch {
          // Unreadable means "assume not moderated": writing OData__ModerationStatus to a
          // library without moderation fails the whole merge, taking Full Name with it.
        }
      }
      let plannedOps = 0;
      for (const lib of ["Staging", "Documents"] as LibTarget[]) {
        for (const t of targets) {
          plannedOps += 1;
          if (fullNameFields.has(lib)) plannedOps += 1;
          if (t.isLeaf) plannedOps += gridPerLeafFor(t);
          // Count exactly what the assignment loop will attempt: this folder's own
          // rows PLUS any fanned down from a parent tier, de-duplicated on
          // group + role the same way. The estimate divides elapsed time by ops
          // COMPLETED, so counting a different set here is what produced the
          // "~197m left" on a three-minute run — the numerator and denominator
          // must measure the same thing.
          const countKeys = new Set<string>();
          // Ancestors only when fan-out is actually on — a reported-but-not-granted
          // row costs no write, and counting it would inflate the estimate.
          const terms = gridSets.fanOut ? [t.assignTerm, ...t.ancestorTerms] : [t.assignTerm];
          for (const term of terms) {
            for (const g of groupMap.get(term.toLowerCase()) ?? []) {
              if (permissionForRole(lib, g.role) === undefined) continue;
              if (LIBRARY_ROLES[lib].indexOf(g.role) === -1) continue;
              countKeys.add(`${g.groupId}|${g.role}`);
            }
          }
          plannedOps += countKeys.size;
        }
      }
      const reconStart = Date.now();
      setReconPlanned(plannedOps);
      setReconStartMs(reconStart);
      setReconNow(reconStart);
      // ── RENAME PASS ────────────────────────────────────────────────────────
      // Runs to completion across EVERY library before the create/lock loop, and
      // before anything writes to the Folder Map.
      //
      // Two earlier attempts put this inside the library loop and both failed the
      // same way. The Staging map row is the only record of a folder's previous
      // name, so renaming Staging and refreshing that row destroyed the very
      // information the Documents pass needed: Documents then saw "already
      // correct", skipped the rename, and the create step made an empty folder at
      // the new name while the real one kept the old one. Persisted, so re-running
      // could not repair it either.
      //
      // Renaming first, then updating the row once, removes the ordering entirely.
      // It also has to precede the create step regardless: `full` points at the new
      // name, so creating first would put an empty folder exactly where the rename
      // needs to land, and the rename would report a collision against it.
      //
      // Note this reverts a folder renamed by hand. Deliberate — the abbreviation
      // list is the single source of truth and every library must agree — and
      // harmless, because uploads resolve by UniqueId rather than by path.
      const renameLibs = ["Staging", "Documents"] as LibTarget[];
      const renameRoots = new Map<string, string>();
      for (const lib of renameLibs) {
        const r = await getLibraryRoot(lib);
        if (r) renameRoots.set(lib, r);
      }
      for (const t of targets) {
        if (!t.termGuid) continue; // segment containers have no term, so no abbreviation
        const wantName = t.relPath.split("/").pop() ?? "";
        const oldName = oldNameByTerm.get(t.termGuid.toLowerCase()) ?? "";
        // Case-insensitive: SharePoint treats sibling names as case-insensitive for
        // uniqueness, so CORU → Coru would collide with itself and report a
        // conflict that is not one.
        if (!oldName || !wantName || oldName.toLowerCase() === wantName.toLowerCase()) continue;
        const parentRel = t.relPath.slice(0, t.relPath.lastIndexOf("/"));
        let newStagingUrl = "";
        for (const lib of renameLibs) {
          const libRoot = renameRoots.get(lib);
          if (!libRoot) continue;
          const oldProbe = await probeFolderByPath(
            context.spHttpClient,
            siteUrl,
            `${libRoot}${parentRel}/${oldName}`,
          );
          if (!oldProbe.folder) continue; // nothing under the old name here
          const renamed = await renameFolder(
            context.spHttpClient,
            siteUrl,
            oldProbe.folder.serverRelativeUrl,
            wantName,
          );
          if (renamed.ok) {
            entries.push({ msg: `  ✎ ${lib}${parentRel}: renamed ${oldName} → ${wantName}`, ok: true });
            if (lib === "Staging") newStagingUrl = renamed.serverRelativeUrl ?? "";
            await tick();
          } else if (renamed.conflict) {
            // Two terms want one folder name. Forcing it would merge two units'
            // documents behind a single ACL — the isolation failure findCollisions
            // exists to prevent — so this needs a person.
            entries.push({
              msg: `  ⚠ ${t.label} — cannot rename "${oldName}" to "${wantName}" in ${lib}: a folder of that name is already there. Fix the abbreviation, then re-run.`,
              ok: false,
            });
          } else {
            entries.push({
              msg: `  ✗ ${t.label} — rename "${oldName}" → "${wantName}" in ${lib} FAILED (HTTP ${renamed.status}) ${renamed.detail ?? ""}`,
              ok: false,
            });
          }
        }
        // Refresh the stored path only now, with every library already handled. The
        // UniqueId is untouched by a rename, so the verification step later treats
        // the row as valid and would otherwise leave FolderUrl permanently stale —
        // it self-heals only when the folder is missing.
        const row = mapByTerm.get(t.termGuid.toLowerCase());
        if (newStagingUrl && row) {
          await updateFolderMapping(context.spHttpClient, siteUrl, row.itemId, {
            folderUniqueId: row.folderUniqueId,
            folderUrl: newStagingUrl,
            title: t.label,
          });
          row.folderUrl = newStagingUrl;
        }
      }

      for (const lib of ["Staging", "Documents"] as LibTarget[]) {
        const root = await getLibraryRoot(lib);
        if (!root) {
          entries.push({ msg: `${lib}: library root not found — skipped`, ok: false });
          continue;
        }
        for (const t of targets) {
          const full = `${root}${t.relPath}`;
          const folderLabel = `${lib}${t.relPath}`;
          try {
            pushFolder(lib, `${folderLabel} — creating…`, "run");
            const existed = await folderExists(full);
            if (!existed) {
              await createFolder(full);
              await breakInheritance(full);
              if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(full, ownerGroupId, fullCtrlId);
              entries.push({ msg: `${folderLabel} — created + locked ✓`, ok: true });
              setLastFolder(lib, `${folderLabel} — created + locked`, "ok");
              bumpFolders();
              await tick();
            } else {
              const isUnique = await getHasUniquePerms(full);
              if (isUnique === false) {
                await breakInheritance(full);
                if (ownerGroupId !== null && fullCtrlId !== undefined) await addRoleAssignment(full, ownerGroupId, fullCtrlId);
                entries.push({ msg: `${folderLabel} — existed, locked ✓`, ok: true });
                setLastFolder(lib, `${folderLabel} — existed, locked`, "ok");
                bumpFolders();
                await tick();
              } else {
                entries.push({ msg: `${folderLabel} — already locked, skipped`, ok: true });
                setLastFolder(lib, `${folderLabel} — already there`, "skip");
              }
            }
            // Counted whichever branch ran, including "already there". Counting only
            // the branches that WROTE something is what made the estimate run away
            // on a re-run.
            step();
            // Full Name: the abbreviation is the folder's name, so the raw term label
            // goes on the item to keep the tree readable in the details pane.
            //
            // Written on every run, not only at creation, for two reasons: it backfills
            // the trees that already exist (this column arrives after they were built),
            // and it is how a RENAMED term's label catches up — the rename pass above
            // moves the folder when the ABBREVIATION changes, but a term relabelled in
            // the term store with its abbreviation untouched changes nothing on disk and
            // would otherwise keep its stale full name forever.
            //
            // The read guard keeps that cheap: an unchanged value costs one GET, which
            // is not throttled, instead of a write that is. On a settled tree this is
            // the normal case, so the extra pass is close to free.
            const fullNameField = fullNameFields.get(lib);
            const wantCtId = folderCtIds.get(lib);
            if (fullNameField || wantCtId) {
              // Planned as one step per target per library, so it is counted here
              // whether the merge turns out to be needed or not.
              if (fullNameFields.has(lib)) step();
              try {
                const moderated = moderatedLibs.has(lib);
                const state = await getFolderItemState(full, fullNameField, moderated);
                const values: Record<string, string | number> = {};
                if (fullNameField && state.fullName !== t.fullName) values[fullNameField] = t.fullName;
                // Stamp the content type in the SAME merge — no extra request, no extra
                // throttle cost. Compared case-insensitively because SharePoint is not
                // consistent about the hex casing it returns.
                if (wantCtId && (state.contentTypeId ?? "").toLowerCase() !== wantCtId.toLowerCase()) {
                  values.ContentTypeId = wantCtId;
                }
                // Moderation status MUST travel in its own merge. SharePoint rejects any
                // request that sets it alongside another field:
                //   "You cannot change moderation status and set other item properties at
                //    that same time." (-2146232832, HTTP 500)
                //
                // It must also run AFTER the property merge, never before: in a moderated
                // library ANY property write re-pends the item. Approving first and then
                // writing Full Name leaves the folder Pending, which is not cosmetic — a
                // pending folder is invisible to an uploader under approver-only draft
                // security, and the upload form then reports "the mapped unit folder no
                // longer exists" because its UniqueId lookup 404s for that user.
                const wrote: string[] = [];
                // The property write and the approve are caught SEPARATELY, because they are
                // not equally serious. A failed Full Name is cosmetic — the folder still
                // controls access correctly, just with a blank label in the details pane. A
                // failed APPROVE leaves the folder Pending, which hides it from every
                // uploader but its creator and makes the upload form report "the mapped unit
                // folder no longer exists". One shared catch reported both as ok:true, so the
                // serious one hid inside a wall of the cosmetic one.
                try {
                  if (Object.keys(values).length > 0) {
                    await setFolderItemFields(full, values);
                    if (fullNameField && values[fullNameField] !== undefined) {
                      wrote.push(`${FULL_NAME_COLUMN_TITLE} = ${t.fullName}`);
                    }
                    if (values.ContentTypeId !== undefined) {
                      wrote.push(`content type → ${resolvedFolderCtName ?? FOLDER_CONTENT_TYPE_CANDIDATES[0]}`);
                    }
                  }
                } catch (e) {
                  // Not fatal, and deliberately ok:true — a label failure must not stop the
                  // folder's ACL work, which is the part that actually controls access, nor
                  // gate the orphan prune.
                  entries.push({ msg: `  ⚠ ${folderLabel} — could not set ${FULL_NAME_COLUMN_TITLE}: ${(e as Error).message}`, ok: true });
                }
                // Re-approve when the folder was already pending, OR when the write above
                // just re-pended it. Missing that second case is what silently un-approves
                // a settled tree on every run.
                //
                // Runs even if the property write threw: a folder left Pending by a failed
                // Full Name still needs approving, and skipping it would turn a cosmetic
                // failure into an access one.
                const rePended = Object.keys(values).length > 0;
                const wasPending = state.moderationStatus !== undefined && state.moderationStatus !== 0;
                if (moderated && (rePended || wasPending)) {
                  try {
                    await setFolderItemFields(full, { OData__ModerationStatus: 0 });
                    wrote.push("approved (a pending folder is invisible under approver-only draft security)");
                  } catch (e) {
                    // ok:FALSE. An access failure, not a label failure: the folder stays
                    // Pending and its unit cannot reach their own files inside it. It must
                    // read as an error and it must gate the clean-run guards.
                    entries.push({
                      msg: `  ✗ ${folderLabel} — COULD NOT APPROVE, folder stays pending and will be invisible to its uploaders: ${(e as Error).message}`,
                      ok: false,
                    });
                  }
                }
                if (wrote.length > 0) {
                  entries.push({ msg: `  ↳ ${wrote.join(", ")}`, ok: true });
                  await tick();
                }
              } catch (e) {
                // Only getFolderItemState can reach here now — both writes catch their own.
                // A failed READ means we could not tell what the folder already had, so
                // nothing was attempted. ok:true: it changed nothing and must not gate the
                // orphan prune, but it IS worth surfacing, because a folder whose state is
                // unreadable is also a folder we did not approve.
                entries.push({ msg: `  ⚠ ${folderLabel} — could not read folder state, ${FULL_NAME_COLUMN_TITLE} and approval skipped: ${(e as Error).message}`, ok: true });
              }
            }
            // Map Staging term folders only (the segment container has no term).
            if (lib === "Staging" && t.termGuid) {
              const existingRow = mapByTerm.get(t.termGuid.toLowerCase());
              if (!existingRow) {
                // Unmapped term → create the row.
                const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, full);
                if (resolved) {
                  await writeFolderMapping(context.spHttpClient, siteUrl, {
                    termGuid: t.termGuid,
                    folderUniqueId: resolved.uniqueId,
                    title: t.label,
                    folderUrl: resolved.serverRelativeUrl,
                    section: t.section,
                  });
                  entries.push({ msg: `  ↳ mapped ${t.label} → ${resolved.uniqueId}`, ok: true });
                }
              } else {
                // Mapped already — VERIFY the stored UniqueId instead of assuming it is
                // right. Ask whether THAT folder still exists; do NOT compare against
                // whatever sits at the term-label path. A renamed or moved folder keeps
                // its UniqueId, so it still resolves and the row is left alone — that is
                // the rename-proofing working. Comparing by path would see a difference
                // and repoint the row at a freshly created empty folder, abandoning the
                // real one and its documents.
                const probe = await probeFolderById(
                  context.spHttpClient,
                  siteUrl,
                  existingRow.folderUniqueId,
                );
                if (probe.folder) {
                  // Row still valid. Nothing to do — any rename already happened
                  // before the folder was created, further up this iteration.
                } else if (probe.confirmedMissing) {
                  // The mapped folder is gone (deleted, then recreated by this run or by
                  // hand). Repoint the row at the folder that is actually there now —
                  // this is the self-heal that the old skip-if-mapped logic prevented.
                  const resolved = await resolveFolderByPath(context.spHttpClient, siteUrl, full);
                  if (resolved) {
                    await updateFolderMapping(context.spHttpClient, siteUrl, existingRow.itemId, {
                      folderUniqueId: resolved.uniqueId,
                      folderUrl: resolved.serverRelativeUrl,
                      title: t.label,
                    });
                    existingRow.folderUniqueId = resolved.uniqueId;
                    existingRow.folderUrl = resolved.serverRelativeUrl;
                    entries.push({
                      msg: `  ↻ remapped ${t.label} — stale folder id replaced with ${resolved.uniqueId}`,
                      ok: true,
                    });
                  } else {
                    entries.push({
                      msg: `  ⚠ ${t.label} — mapped folder is gone and no folder found at ${full}; row left as-is`,
                      ok: false,
                    });
                  }
                } else {
                  // Throttle, permission error, malformed request — NOT evidence of
                  // deletion. Changing the row on this would be acting on a guess.
                  entries.push({
                    msg: `  ⚠ ${t.label} — could not verify mapped folder (HTTP ${probe.status}); row left unchanged`,
                    ok: false,
                  });
                }
              }
            }
            // Auto-assign DMS Group Map groups to this folder by role. Staging gets
            // UPL/APR only; Documents gets MEMBER (viewer) groups only. Idempotent
            // (add-role merges). A folder whose term has no rows is flagged.
            const groupRows = groupMap.get(t.assignTerm.toLowerCase()) ?? [];

            // Departmental fan-out. A row on a NON-LEAF term (a department) applies
            // to that folder AND to every folder beneath it, at the row's own level —
            // the mirror of the ancestor Read fan-UP below. Without it, Head of
            // Department is indistinguishable from Head of Unit: unit folders have
            // unique permissions, so a department-tier grant stops dead at the
            // department folder and never reaches a single unit.
            //
            // Leaf rows never fan, because a leaf has no descendants — unit isolation
            // holds by construction rather than by a special case that could be got
            // wrong. And since the model is leaf-only today (no department rows exist
            // on any site), this is inert until an admin deliberately creates one.
            //
            // NOTE for deeper Levels chains than [Department, Unit]: with a 2027
            // segment (Region -> Estate/Mill) a Region-tier row reaches everything
            // below it. That is the honest meaning of a Region row, but it is wider
            // than "departmental fan-out" suggests — review before onboarding those
            // segments. Spec §4.1.
            const inherited: Array<{ row: GroupMapRow; fromTerm: string }> = [];
            for (const anc of t.ancestorTerms) {
              for (const g of groupMap.get(anc.toLowerCase()) ?? []) {
                inherited.push({ row: g, fromTerm: anc });
              }
            }

            // Isolation rule: which roles this library accepts. MEMBER (base/viewer)
            // groups are Documents-only and must NEVER land on Staging, else a viewer
            // could read pending documents. A fanned row is filtered identically —
            // inheriting a grant must not widen which library it reaches.
            const accepts = (role: string): boolean =>
              permissionForRole(lib, role) !== undefined &&
              LIBRARY_ROLES[lib].indexOf(role) !== -1;

            const applicable = groupRows.filter(g => accepts(g.role));
            const fanned = inherited.filter(i => accepts(i.row.role));

            // One list, direct rows first, de-duplicated on group + role. A group
            // holding BOTH a unit row and a department row for the same role would
            // otherwise be granted twice: harmless (add-role merges) but it doubles
            // the log and makes the run look like it did more work than it did.
            // Direct wins, so the log attributes the grant to the nearer row.
            const seenGrant = new Set<string>();
            const toGrant: Array<{ row: GroupMapRow; viaTerm?: string }> = [];
            const g0 = (r: GroupMapRow): string => r.groupName || r.groupId;
            for (const g of applicable) {
              const k = `${g.groupId}|${g.role}`;
              if (seenGrant.has(k)) continue;
              seenGrant.add(k);
              toGrant.push({ row: g });
            }
            for (const i of fanned) {
              const k = `${i.row.groupId}|${i.row.role}`;
              if (seenGrant.has(k)) continue;
              seenGrant.add(k);
              // Opt-in. While recon_departmentFanOut is off we REPORT the grant and
              // make none — see loadReconGridTermSets for why a leftover
              // department-tier row makes silently granting unacceptable. Reported
              // as a warning, not an error: this is the configured behaviour, and an
              // error would gate the orphan prune on it.
              //
              // SEGVIEW is EXEMPT from the gate, for exactly the reason GLOBAL needs no gate:
              // the role name is itself the consent. The gate exists because a leftover
              // segment- or department-tier MEMBER row is indistinguishable BY TIER from a
              // deliberate wide grant. SEGVIEW cannot be that — it granted nothing on any
              // site until 2026-08-07, so every row that exists was written deliberately, and
              // a segment-tier row IS its intended shape rather than a legacy accident.
              if (!gridSets.fanOut && i.row.role !== "SEGVIEW") {
                entries.push({
                  msg: `  ⚠ ${g0(i.row)} would inherit ${permissionForRole(lib, i.row.role)} on ${folderLabel} from a parent-tier mapping — not granted (recon_departmentFanOut is off)`,
                  ok: true,
                });
                continue;
              }
              // A SEGMENT-tier row reaches descendants for SEGVIEW alone (2026-08-09). The
              // segment term-set GUID is now in every descendant's ancestorTerms so the
              // C-Level "one business segment" persona can work at all — but a leftover
              // segment-tier MEMBER row from before 2026-07-29 is indistinguishable BY TIER
              // from a deliberate one, and honouring it would hand a viewer Read across an
              // entire business segment. SEGVIEW cannot be such a leftover: it granted
              // nothing on any site until 2026-08-07, so every row that exists was written
              // on purpose. The role name is the consent, exactly as it is for GLOBAL.
              //
              // NOT gated on recon_departmentFanOut either: that switch is about DEPARTMENT
              // rows, and turning it off must not silently disable a C-Level.
              if (segmentTermSets.has((i.fromTerm ?? "").toLowerCase()) && i.row.role !== "SEGVIEW") {
                entries.push({
                  msg: `  ⚠ ${g0(i.row)} sits on the SEGMENT tier and would reach ${folderLabel} — not granted (only SEGVIEW fans down from a segment)`,
                  ok: true,
                });
                continue;
              }
              toGrant.push({ row: i.row, viaTerm: i.fromTerm });
            }

            // C-level view fan-down. A GLOBAL row carries no term at all and reaches
            // every folder in every segment, at Read, in Documents only.
            //
            // NOT gated by recon_departmentFanOut, and that is the whole reason it is a
            // role of its own. The gate exists because a leftover segment- or
            // department-tier MEMBER row from before 2026-07-29 is indistinguishable,
            // BY TIER, from a deliberate wide grant, so tier-based fanning has to be
            // opt-in. A GLOBAL row has no tier to be mistaken for and has never granted
            // anything until now: the explicit role name is the consent, so no switch is
            // needed and none should be added.
            //
            // MEMBER keeps its guard untouched — a segment-tier MEMBER row still fans
            // nowhere, at any setting.
            //
            // GLOBAL rows are stored with an empty UnitTermGuid, so they key on "".
            for (const g of groupMap.get("") ?? []) {
              if (g.role !== "GLOBAL" || !accepts(g.role)) continue;
              const k = `${g.groupId}|${g.role}`;
              if (seenGrant.has(k)) continue;
              seenGrant.add(k);
              toGrant.push({ row: g, viaTerm: "(all segments)" });
            }

            // Only a LEAF (unit) folder is expected to carry Group Map rows — the model is
            // leaf-only by design (see CLAUDE.md / the leaf-only authorization spec), so the
            // segment and department tiers having none is the correct state, not a problem.
            // Warning on them buried the real warnings under ~260 structurally unfixable
            // ones. Parent tiers stay admin-only and silent; members reach their unit
            // through the ancestor Read grants below.
            //
            // Tested against toGrant, not applicable: a unit reached ONLY by a
            // department row is properly provisioned, and calling it "admin-only"
            // would send an admin hunting for a row that should not exist.
            if (toGrant.length === 0 && t.isLeaf) {
              entries.push({ msg: `  ⚠ ${folderLabel} — no group-map groups for this unit (locked admin-only)`, ok: true });
            }
            const grantedPids: Array<{ groupName: string; pid: number }> = [];
            for (const { row: g, viaTerm } of toGrant) {
              // Fanned grants are logged distinctly. "Why does this group hold
              // DMS Approve on a unit folder with no row for it" is otherwise
              // unanswerable from the log, and an unexplained grant is
              // indistinguishable from a bug.
              const arrow = viaTerm ? "↳↓" : "↳";
              const via = viaTerm ? " (inherited from a parent-tier mapping)" : "";
              // Resolved ONCE, per library. Read straight from ROLE_TO_PERMISSION and an
              // uploader's Documents grant would say "CRS Upload" in the log while the
              // assignment said Read — or worse, actually be CRS Upload.
              const levelName = permissionForRole(lib, g.role);
              const roleDefId = roleDefs.find(r => r.name === levelName)?.id;
              if (roleDefId === undefined) {
                entries.push({ msg: `  ⚠ ${g.groupName} — no "${levelName}" role definition on site`, ok: false });
                pushAssign(lib, `${t.label}: "${levelName}" role missing on site`, "warn");
                step();
                continue;
              }
              step();
              try {
                const pid = spGroupPrincipalId(g.groupId);
                await addRoleAssignment(full, pid, roleDefId);
                grantedPids.push({ groupName: g.groupName, pid });
                entries.push({ msg: `  ${arrow} ${g.groupName} → ${levelName}${via}`, ok: true });
                pushAssign(lib, `${t.label} → ${g.groupName} (${levelName})${via}`, "ok");
                bumpAssigns();
                await tick();
              } catch (e) {
                entries.push({ msg: `  ✗ ${g.groupName} → ${levelName} FAILED: ${(e as Error).message}`, ok: false });
                // Most common cause: the SP group doesn't exist yet (or is a legacy Entra
                // row). Surface the admin-needs-to-create-it message in the right panel.
                pushAssign(lib, `Group "${g.groupName}" not found — ask an administrator to create it`, "admin");
              }
            }
            // Ancestor browse Read: granted, so a user can click down to their folder.
            //
            // This pass used to grant each group Read on every ancestor folder on its
            // path so members could click down to their unit.
            //
            // RESTORED 2026-08-07, having been removed on 2026-08-04. The 2026-08-04 rule
            // ("a Head of Unit must not see the department folder, a Head of Department
            // must not see the segment folder") was WITHDRAWN by the client's next
            // restatement: navigation now starts at the business segment for every family,
            // in both libraries. "View the unit folder ONLY" meant not seeing SIBLING
            // units — which siblings already are, since they carry no grant and SharePoint
            // security-trims them.
            //
            // Removing the grant was never visible on an existing site, because earlier
            // runs had already made the assignments and reconciliation only ever added.
            // It bites on a NEW library: every folder is new, none gets ancestor Read, and
            // the library root renders empty for every non-admin. That is exactly the state
            // the recreated Approval Document library is in, which is how this was caught.
            if (grantedPids.length > 0 && readId !== undefined) {
              for (const anc of ancestorRelPaths(t.relPath)) {
                const ancFull = `${root}${anc}`;
                // Declared with a definite type rather than inferred from the cache read:
                // the loop below reassigns it after each grant, which loses the narrowing
                // that the undefined-check would otherwise give.
                const cached = ancAssignCache.get(ancFull);
                let existing: ExistingAssign[];
                if (cached === undefined) {
                  existing = await getRoleAssignments(ancFull);
                  ancAssignCache.set(ancFull, existing);
                } else {
                  existing = cached;
                }
                for (const gp of grantedPids) {
                  // The site-entry group reaches the site, never a folder. It is not part
                  // of anyone's browse path and must not be granted one.
                  if (gp.groupName === siteEntryGroupTitle()) continue;
                  // Idempotent: a second unit under the same department must not re-grant
                  // Read its parent already holds. The cache makes that one probe per
                  // ancestor per run rather than one per unit.
                  const grantKey = `${ancFull}|${gp.pid}`;
                  if (ancGranted.has(grantKey)) continue;
                  if (existing.some(a => a.principalId === gp.pid && a.roleDefId === readId)) continue;
                  try {
                    await addRoleAssignment(ancFull, gp.pid, readId);
                    ancGranted.add(grantKey);
                    entries.push({ msg: `  ↳ ${gp.groupName} — Read (browse) on ${anc}`, ok: true });
                    pushAssign(lib, `${anc} → ${gp.groupName} Read (ancestor browse)`, "ok");
                    await tick();
                  } catch (e) {
                    entries.push({ msg: `  ✗ ${gp.groupName} — failed to grant Read on ${anc}: ${(e as Error).message}`, ok: false });
                  }
                }
              }
            }
            // Under leaf (unit) folders, pre-create the Year × Document Type grid.
            // These inherit the unit's ACL (no lock, no map). Idempotent via
            // ensureFolder. Logged as a per-unit count, not one line per folder.
            //
            // These grid folders are NOT approved, and deliberately so — two reasons it
            // cannot be done here:
            //   1. The fast path below settles the entire grid on ONE probe, so a per-folder
            //      approve in this loop would never fire on a site that already has its grid.
            //   2. ensureFolder is shared with Form.tsx / BulkUpload.tsx, where the caller is
            //      an UPLOADER. Writing OData__ModerationStatus = 0 needs ApproveItems, which
            //      a PIC does not hold — so uploaders keep creating Pending grid folders
            //      between runs no matter what this pass does.
            // Consequence: Draft Item Security must stay "any user who can read items". Under
            // approver-only, a PIC sees an empty unit folder and cannot browse to their own
            // pending files. Closing it needs an end-of-run library sweep (page by `ID gt`,
            // never a $filter on FSObjType — the grid puts these libraries over the 5,000-item
            // list-view threshold). Not built: the only requirement that wanted approver-only
            // was per-uploader isolation, which is out of scope.
            const tierNames = t.isLeaf ? gridTiersFor(t) : [];
            const plan = gridPlan(tierNames);
            if (t.isLeaf && plan.total > 0) {
              const gridTotal = plan.total;
              let grid = 0;
              // FAST PATH. The grid is built in order, so if the LAST year's LAST
              // document-type folder is there, the whole grid is there. One probe
              // replaces 63 round-trips per leaf per library. Without this a no-op
              // re-run cost the same as a first run (measured: 6,090 folders in
              // 111m for 48 units, of which only 6 were real group assignments).
              // Trade-off: a folder hand-deleted from the middle of a COMPLETE grid
              // is not restored here — the upload form ensure-creates it on demand.
              // probeFolderByPath retries 429s, so a throttle cannot fake "incomplete"
              // and trigger a needless full rebuild.
              const gridProbe = await probeFolderByPath(
                context.spHttpClient,
                siteUrl,
                `${full}/${plan.lastPath.join("/")}`,
              );
              if (gridProbe.folder) {
                entries.push({ msg: `  ↳ ${lib}${t.relPath} — Year × Document Type grid already complete (${gridTotal}), skipped`, ok: true });
                setLastFolder(lib, `${t.label} grid: already complete, skipped`, "skip");
                // The fast path settles every planned grid step in one probe; the
                // estimate has to see them land or it keeps counting them as pending.
                step(gridTotal);
              } else {
              pushFolder(lib, `${t.label} grid: 0 / ${gridTotal}`, "run");
              // One walk down the configured chain, replacing the old fixed
              // Year-then-Document-Type pair. Depth-first and SEQUENTIAL (never
              // parallel) — a parallel burst is what tripped the 429 throttle.
              const buildTier = async (parent: string, depth: number): Promise<void> => {
                if (depth >= plan.tiers.length) return;
                for (const name of plan.tiers[depth]) {
                  const made = await ensureFolder(context.spHttpClient, siteUrl, parent, name);
                  step();
                  if (made) { grid++; bumpFolders(); }
                  setLastFolder(lib, `${t.label} grid: ${grid} / ${gridTotal}`, "run");
                  // Only pace REAL writes. Charging the throttle delay to a folder that
                  // already existed is what made a no-op re-run as slow as a first run.
                  if (made?.created) await tick();
                  if (!made) continue;
                  await buildTier(made.serverRelativeUrl, depth + 1);
                }
              };
              await buildTier(full, 0);
              }
              entries.push({ msg: `  ↳ ${lib}${t.relPath} — Year × Document Type grid: ${grid} folder(s) ensured`, ok: true });
              setLastFolder(lib, `${t.label} grid: ${grid} / ${gridTotal} ✓`, "ok");
            }
          } catch (e) {
            entries.push({ msg: `${lib}${t.relPath} — FAILED: ${(e as Error).message}`, ok: false });
          }
        }
      }
      // Site-entry self-heal: ensure every member of a group we MANAGE is also in
      // DMS_SITE_MEMBERS, so users added the native way (bypassing the web part's
      // auto-add) can still open the site. See site-entry-access-layer spec §7a.
      //
      // "Managed" comes from the Group Map's GroupId column, NOT from a title prefix.
      // Until 2026-08-04 this filtered on `title.startsWith("DMS_")`, which the client's
      // new prefix-less convention (GHO_GF_CORU_UPLOADER) matches zero of — and the failure
      // is silent: nobody gets site entry and the run still reports ✓, producing "only SOME
      // users cannot open the site", the hardest version of this to diagnose.
      //
      // Reading the list is also strictly more accurate than the name test ever was: it
      // finds a group whatever it is called, covers every Scope in one query, and skips a
      // group somebody hand-named with our prefix but never actually mapped.
      try {
        setReconPhase("Syncing site-entry group…");
        const allGroups = await fetchAllSiteGroups(context.spHttpClient, siteUrl);
        const entryGroup = findSiteEntryGroup(allGroups);
        if (!entryGroup) {
          entries.push({ msg: `⚠ ${siteEntryGroupTitle()} not found — run "Set up site entry" first (site-entry sync skipped)`, ok: false });
        } else {
          const entryMembers = await getGroupMembers(context.spHttpClient, siteUrl, entryGroup.id);
          const already = new Set(entryMembers.map((m) => m.loginName.toLowerCase()));
          // GroupId ONLY. That column predates Scope/Target, so this cannot hit the
          // whole-request HTTP 400 that naming a nonexistent $select column causes
          // (CLAUDE.md #11) — no fallback query needed, and every scope is included.
          const managedIds = new Set<number>();
          const idRes = await context.spHttpClient.get(
            `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.groupMap))}')/items?$select=GroupId&$top=5000`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: "application/json;odata=nometadata" } },
          );
          if (idRes.ok) {
            const idJson = await idRes.json();
            for (const r of (idJson.value ?? []) as Array<{ GroupId?: string }>) {
              const n = Number((r.GroupId ?? "").toString().trim());
              if (n > 0 && n % 1 === 0) managedIds.add(n);
            }
          } else {
            entries.push({ msg: `⚠ site-entry sync: could not read Group Map ids (HTTP ${idRes.status}) — no members synced`, ok: false });
          }
          managedIds.delete(entryGroup.id);
          const managedGroups = allGroups.filter((g) => managedIds.has(g.id));
          if (idRes.ok && managedIds.size === 0) {
            entries.push({ msg: `${siteEntryGroupTitle()}: no mapped groups in Group Map — nothing to sync`, ok: true });
          }
          let healed = 0;
          for (const g of managedGroups) {
            const members = await getGroupMembers(context.spHttpClient, siteUrl, g.id).catch(() => []);
            for (const m of members) {
              if (already.has(m.loginName.toLowerCase())) continue;
              try {
                await addGroupMember(context.spHttpClient, siteUrl, entryGroup.id, m.loginName);
                already.add(m.loginName.toLowerCase());
                healed++;
              } catch { /* skip a member that can't be added; not fatal */ }
            }
          }
          entries.push({
            msg: healed > 0
              ? `${siteEntryGroupTitle()}: added ${healed} member(s) missing site entry ✓`
              : `${siteEntryGroupTitle()}: all mapped-group members already have site entry ✓`,
            ok: true,
          });
        }
      } catch (e) {
        entries.push({ msg: `Site-entry sync skipped — ${(e as Error).message}`, ok: false });
      }

      // ── Prune orphaned Folder Map rows ────────────────────────────────────────
      // A row whose TERM no longer exists is never visited by the loop above, so it
      // would linger forever. This is the only place that can see them: the run has
      // just enumerated every valid term.
      //
      // Deletes LIST ROWS ONLY, never folders. A row is derived data that a later run
      // can rebuild from the term store + the folder tree; the folder holds documents
      // that exist nowhere else. A folder left behind is reported so a human can decide.
      const errorsBeforePrune = entries.filter((e) => !e.ok).length;
      if (incompleteSegments.length > 0) {
        // Checked separately from the error count even though an entry was already
        // pushed above — this is THE condition prune must never run under, and it
        // should not depend on that entry still being there.
        entries.push({
          msg: `Prune skipped — the term store could not be fully read (${incompleteSegments.length} segment(s)); rows are never pruned against a partial term list`,
          ok: true,
        });
      } else if (errorsBeforePrune > 0) {
        // Hard guard. A half-failed term-store read returns a SHORT target list, and a
        // short list makes perfectly healthy rows look orphaned — pruning on that would
        // wipe the map wholesale. No clean run, no prune.
        entries.push({
          msg: `Prune skipped — run had ${errorsBeforePrune} error(s); orphaned map rows are only removed after a clean run`,
          ok: true,
        });
      } else {
        try {
          setReconPhase("Pruning orphaned folder map rows…");
          const enumerated = new Set(
            targets
              .filter((t) => t.termGuid)
              .map((t) => (t.termGuid as string).toLowerCase()),
          );
          let pruned = 0;
          for (const row of mapRows) {
            if (!row.termGuid) continue;
            if (enumerated.has(row.termGuid.toLowerCase())) continue; // term alive — handled above
            const probe = await probeFolderById(
              context.spHttpClient,
              siteUrl,
              row.folderUniqueId,
            );
            if (!probe.folder && !probe.confirmedMissing) {
              // Could not tell. Keep the row; never delete on a guess.
              entries.push({
                msg: `  ⚠ orphaned row "${row.title}" — could not verify its folder (HTTP ${probe.status}); row kept`,
                ok: false,
              });
              continue;
            }
            try {
              await deleteFolderMapRow(context.spHttpClient, siteUrl, row.itemId);
              pruned++;
              entries.push({
                msg: probe.folder
                  ? `  ✂ pruned "${row.title}" — term deleted from the term store. Its FOLDER still exists at ${probe.folder.serverRelativeUrl} and was NOT touched — delete it manually if that is intended.`
                  : `  ✂ pruned "${row.title}" — term and folder both gone`,
                ok: true,
              });
            } catch (e) {
              entries.push({
                msg: `  ✗ could not prune "${row.title}": ${(e as Error).message}`,
                ok: false,
              });
            }
          }
          entries.push({
            msg: pruned > 0
              ? `Folder Map: pruned ${pruned} orphaned row(s) ✓`
              : `Folder Map: no orphaned rows ✓`,
            ok: true,
          });
        } catch (e) {
          entries.push({ msg: `Prune skipped — ${(e as Error).message}`, ok: false });
        }

        // ── Repair term-GUID orphans ───────────────────────────────────────────
        // Deleting a term and re-adding it under the same name orphans a row in
        // DMS Term Abbreviation and DMS Group Map at once; the re-created term
        // carries a NEW guid, so nothing joins them. Only the LABEL survives, and
        // the abbreviation row keeps it in Title — that is the whole repair.
        //
        // The abbreviation row is NEVER auto-deleted, unlike a Folder Map row. A
        // map row is derivable (term + folder on disk rebuilds it); an abbreviation
        // exists nowhere else, and deleting one lets a later re-created term take a
        // DIFFERENT abbreviation, which Task 4's rename pass would then apply to a
        // live folder full of documents. Repair, or report. See spec 2026-08-02 §7.
        //
        // Inside the same else-branch as the prune, so it inherits both guards: a
        // partial term-store read must never look like a mass deletion.
        try {
          setReconPhase("Repairing term-GUID orphans…");
          const enumeratedTerms = new Set(
            targets
              .filter((t) => t.termGuid)
              .map((t) => (t.termGuid as string).toLowerCase()),
          );
          const orphanRows = abbrevRows.filter(
            (r) =>
              r.termGuid.trim().length > 0 &&
              !enumeratedTerms.has(r.termGuid.trim().toLowerCase()),
          );
          if (orphanRows.length === 0) {
            entries.push({ msg: `${abbrevListTitle()}: no orphaned rows ✓`, ok: true });
          } else {
            const plan = planOrphanRepairs(orphanRows, missingAbbrev);
            const groupRows =
              plan.repairs.length > 0 ? await loadGroupMapRowsForRepair() : [];

            for (const rep of plan.repairs) {
              const oldGuid = rep.orphan.termGuid.trim().toLowerCase();
              // Every row holding the dead GUID is the same unit by definition —
              // the join cannot DISCOVER the new GUID, but it is how the answer is
              // applied once the label has established it.
              const affected = groupRows.filter(
                (g) => g.termGuid.toLowerCase() === oldGuid,
              );
              try {
                await patchListItem(abbrevListTitle(), rep.orphan.itemId, {
                  TermGuid: rep.term.termGuid,
                });
              } catch (e) {
                entries.push({
                  msg: `  ✗ could not re-point "${rep.orphan.abbreviation}" (${rep.orphan.title}): ${(e as Error).message}`,
                  ok: false,
                });
                continue;
              }
              let groupsFixed = 0;
              for (const g of affected) {
                try {
                  await patchListItem(cachedListTitle(LIST_SUFFIX.groupMap), g.itemId, {
                    UnitTermGuid: rep.term.termGuid,
                  });
                  groupsFixed++;
                } catch (e) {
                  entries.push({
                    msg: `  ✗ re-pointed the abbreviation but NOT the group "${g.groupName}": ${(e as Error).message}. That folder will be locked admin-only until this row is fixed by hand.`,
                    ok: false,
                  });
                }
              }
              entries.push({
                msg: `  ✎ ${rep.orphan.abbreviation} — "${rep.orphan.title}" was re-created; re-pointed to ${rep.term.termGuid}${affected.length > 0 ? ` (+${groupsFixed} of ${affected.length} group-map row(s))` : ""}. Run again to create its folder.`,
                ok: true,
              });
            }

            for (const amb of plan.ambiguous) {
              entries.push({
                msg:
                  `  ? AMBIGUOUS: "${amb.orphan.title}" (${amb.orphan.level}, ${amb.orphan.abbreviation}) — ` +
                  `${amb.orphanCandidates.length} orphaned row(s) and ${amb.termCandidates.length} re-created term(s) share that name ` +
                  `[${amb.termCandidates.map((t) => t.termGuid).join(", ")}]. ` +
                  `Not repaired: the same name exists under more than one parent, and re-pointing the wrong one would grant another department's groups access to this folder. Set TermGuid by hand.`,
                ok: false,
              });
            }

            for (const u of plan.unmatched) {
              entries.push({
                msg:
                  `  ⚠ ORPHANED: "${u.title}" (${u.level}, ${u.abbreviation}) — its term ${u.termGuid} no longer exists and no re-created term matches the name. ` +
                  `The row was KEPT: it holds the only copy of the abbreviation. Delete it by hand once you are sure the term is gone for good.`,
                ok: false,
              });
            }

            entries.push({
              msg: `${abbrevListTitle()}: ${plan.repairs.length} repaired, ${plan.ambiguous.length} ambiguous, ${plan.unmatched.length} orphaned`,
              ok: plan.ambiguous.length === 0 && plan.unmatched.length === 0,
            });
          }
        } catch (e) {
          entries.push({
            msg: `Orphan repair skipped — ${(e as Error).message}`,
            ok: false,
          });
        }

        // ── Report folders no live term claims ─────────────────────────────────
        // Reconciliation only ever walks term store → libraries, never the reverse,
        // so a permanently deleted term leaves its folder behind with broken
        // inheritance INTACT — still granting Contribute and Design to that unit's
        // groups. Deleting a term revokes nobody's access; the folder stays
        // reachable by direct link or by browsing the library. Spec 2026-08-02 §8.
        //
        // Report only. NEVER deleted: there are documents behind it.
        try {
          setReconPhase("Checking for folders with no term…");
          const expected = new Set(
            targets.map((t) => t.relPath.toLowerCase()),
          );
          // Descend only into non-leaf targets. Below a leaf sit the Year ×
          // Document Type grid folders, which have no term by design and would
          // otherwise be reported as unclaimed — hundreds of false positives.
          const descendFrom = targets.filter((t) => !t.isLeaf);
          let unclaimed = 0;
          let unreadable = 0;
          for (const lib of ["Staging", "Documents"] as LibTarget[]) {
            const root = await getLibraryRoot(lib);
            if (!root) continue;
            for (const t of descendFrom) {
              const names = await listSubfolders(`${root}${t.relPath}`);
              if (names === undefined) {
                unreadable++;
                continue;
              }
              for (const name of names) {
                if (expected.has(`${t.relPath}/${name}`.toLowerCase())) continue;
                unclaimed++;
                entries.push({
                  msg: `  ⚠ NO TERM: ${lib}${t.relPath}/${name} — no live term maps to this folder, but its permissions are unchanged and its documents are still reachable. Review it; nothing was deleted.`,
                  ok: false,
                });
              }
            }
          }
          if (unreadable > 0) {
            entries.push({
              msg: `  ⚠ could not list ${unreadable} folder(s) while checking for unclaimed folders — that part of the tree was not checked`,
              ok: false,
            });
          }
          if (unclaimed === 0 && unreadable === 0) {
            entries.push({ msg: `Folders: every folder maps to a live term ✓`, ok: true });
          }
        } catch (e) {
          entries.push({
            msg: `Unclaimed-folder check skipped — ${(e as Error).message}`,
            ok: false,
          });
        }
      }

      setLog(entries);
      const failed = entries.filter(e => !e.ok).length;

      // ONE row per run, carrying the whole log. Two reasons that is worth more than it looks: the
      // on-screen log is lost the moment anyone navigates away, and one row per FOLDER would bury
      // every other event in the audit log the first time somebody reconciles.
      //
      // Counts come from LOCALS, never from `reconCounts`: that is React state set during the run,
      // so this closure still sees its render-time value — reading it here would faithfully record
      // zero.
      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.reconciliationRun,
        outcome: failed > 0 ? "Failed" : "Success",
        source: "FolderManager",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        library: "Approval Document + Documents",
        summary:
          `Reconciliation — ${targets.length} folder(s) reconciled` +
          (failed > 0 ? `, ${failed} error(s)` : ", no errors"),
        details: entries.map(e => `${e.ok ? "✓" : "✗"} ${e.msg}`),
      }).catch(() => undefined);

      showToast(
        failed > 0
          ? `Reconciled with ${failed} error(s) — see log.`
          : `Reconciled ${targets.length} folder(s) across Staging + Documents — locked + groups assigned from DMS Group Map.`,
        failed > 0,
      );
    } catch (e) {
      showToast(`Reconciliation failed: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
      setReconRunning(false);
      setReconPhase("");
    }
  };

  /* ── Render helpers ──────────────────────────────────────────────────────────── */

  const renderPermPanel = (node: FolderNode): React.ReactElement | null => {
    const panel = node.perm;
    if (!panel.open) return null;
    return (
      <div style={s.permPanel}>
        <p style={s.permTitle}>{node.isNew ? "Assign groups" : "Permissions"}</p>
        {panel.loading ? (
          <p style={{ fontSize: 12, color: "#666", margin: 0 }}>Loading…</p>
        ) : (
          <>
            {!node.isNew && (
              <>
                <p style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase" as const, letterSpacing: ".05em", color: "#555", margin: "0 0 6px" }}>
                  Current assignments
                </p>
                {panel.existing.length === 0 ? (
                  <p style={{ fontSize: 12, color: "#888", margin: "0 0 10px", fontStyle: "italic" }}>
                    {panel.isUnique === false
                      ? "This folder inherits permissions from its parent — no unique assignments are set at the folder level. Add groups below to break inheritance and assign explicit access."
                      : "No groups assigned to this folder yet. Add groups below."}
                  </p>
                ) : (
                  panel.existing.map(a => (
                    <div key={a.uid} style={{ ...s.assignRow, opacity: a.kept ? 1 : .45 }}>
                      <span style={{ ...s.chip, ...(a.kept ? {} : { textDecoration: "line-through" } as React.CSSProperties) }}>
                        <span style={s.chipName} title={a.title}>{a.title}</span>
                      </span>
                      <select style={s.roleSelect} value={a.roleDefId} disabled={busy || !a.kept}
                        onChange={e => permSetExistingRole(node.id, a.uid, Number(e.target.value))}>
                        {roleDefs.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
                      </select>
                      {a.kept
                        ? <button style={s.chipX} disabled={busy} title="Remove" onClick={() => permToggleKept(node.id, a.uid)}>✕</button>
                        : <button style={s.undoLink} disabled={busy} onClick={() => permToggleKept(node.id, a.uid)}>Undo</button>
                      }
                    </div>
                  ))
                )}
              </>
            )}

            <p style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase" as const, letterSpacing: ".05em", color: "#555", margin: "10px 0 6px" }}>
              {node.isNew ? "Groups (required)" : "Add groups"}
            </p>
            {panel.pending.map(a => (
              <div key={a.uid} style={s.assignRow}>
                <span style={s.chip}><span style={s.chipName} title={a.group.displayName}>{a.group.displayName}</span></span>
                <select style={s.roleSelect} value={a.roleDefId} disabled={busy}
                  onChange={e => permSetPendingRole(node.id, a.uid, Number(e.target.value))}>
                  {roleDefs.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
                </select>
                <button style={s.chipX} disabled={busy} onClick={() => permRemovePending(node.id, a.uid)}>✕</button>
              </div>
            ))}
            <GroupSearch disabled={busy} placeholder={node.isNew ? "Add group (required)…" : "Search for a group to add…"} onSearch={searchGroups} onPick={g => permAddGroup(node.id, g)} />

            <p style={{ fontSize: 11, color: "#888", marginTop: 8, marginBottom: 0 }}>
              {node.isNew
                ? "This folder and its permissions will be created when you click Update."
                : "Changes are applied when you click Update below."}
            </p>
          </>
        )}
      </div>
    );
  };

  const renderAddForm = (mode: Mode, parentNode: FolderNode | null): React.ReactElement => (
    <div style={{ paddingTop: 8 }}>
      <button style={s.addFolderBtn} disabled={busy}
        onClick={() => addNewChild(mode, parentNode ? parentNode.id : null)}>
        {`+ ${parentNode ? "Add subfolder" : "Add folder"}`}
      </button>
    </div>
  );

  const renderNode = (node: FolderNode, mode: Mode, depth: number): React.ReactElement => {
    const isOpen = !!expandedIds[node.id];
    const changed = !node.isNew && !node.isDeleted && node.newName.trim() !== "" && node.newName.trim() !== node.name;
    const permCount = node.perm.existing.filter(a => a.kept).length + node.perm.pending.length;
    const noPendingChanges = !node.perm.existing.some(a => !a.kept) && node.perm.pending.length === 0;
    const permLabel = (() => {
      const arrow = node.perm.open ? "▴" : "▾";
      if (node.isNew) return `Groups ${arrow}`;
      if (!node.perm.loaded) return `Permissions ${arrow}`;
      if (node.perm.isUnique === false && noPendingChanges) return `Permissions (inherited) ${arrow}`;
      return `Permissions (${permCount} group${permCount !== 1 ? "s" : ""}) ${arrow}`;
    })();

    return (
      <div key={node.id}>
        <div style={depth === 0 ? s.parentRow : s.childRow}>
          <button style={s.chevBtn}
            onClick={() => { toggleExpand(node).catch(() => undefined); }}
            title={isOpen ? "Collapse" : "Expand subfolders"}
          >
            {isOpen ? "▾" : "▸"}
          </button>
          <span style={{ fontSize: depth === 0 ? 16 : 14, flexShrink: 0 }}>
            {node.isNew ? "🆕" : depth === 0 ? "📁" : "📂"}
          </span>
          <input className="fm-in" type="text" value={node.newName} disabled={busy || node.isDeleted}
            placeholder={node.isNew ? "New folder name" : undefined}
            onChange={e => updateNode(node.id, n => ({ ...n, newName: e.target.value }))}
            style={{
              ...s.renameIn,
              fontWeight: depth === 0 && !node.isNew ? 600 : 400,
              borderColor: (changed || node.isNew) ? "#0f6c3f" : "#c8c8c8",
              textDecoration: node.isDeleted ? "line-through" : undefined,
              color: node.isDeleted ? "#a4262c" : undefined,
              opacity: node.isDeleted ? .6 : 1,
            }}
          />
          {node.childrenLoaded && node.children.length > 0 && (
            <span style={s.badge}>{node.children.length} subfolder{node.children.length !== 1 ? "s" : ""}</span>
          )}
          {changed && <span style={s.wasLabel}>was: {node.name}</span>}
          {node.isNew && <span style={{ ...s.wasLabel, color: "#0f6c3f" }}>new — not yet created</span>}
          {node.isDeleted && <span style={{ ...s.wasLabel, color: "#a4262c" }}>marked for deletion</span>}

          {!node.isDeleted && (
            <button
              style={{ ...s.permBtn, ...(node.perm.open ? s.permBtnOpen : {}) }}
              onClick={() => { togglePermPanel(node).catch(() => undefined); }}
            >
              {permLabel}
            </button>
          )}
          {node.isNew && (
            <button style={{ ...s.permBtn, color: "#a4262c", borderColor: "#a4262c" }} disabled={busy}
              onClick={() => discardNode(node.id)}>
              Discard
            </button>
          )}
          {!node.isNew && !node.isDeleted && !node.confirmingDelete && (
            <button style={{ ...s.permBtn, color: "#a4262c", borderColor: "#a4262c" }} disabled={busy}
              onClick={() => requestDelete(node.id)}>
              Delete
            </button>
          )}
          {!node.isNew && node.isDeleted && (
            <button style={s.undoLink} disabled={busy} onClick={() => undoDelete(node.id)}>
              Undo
            </button>
          )}
        </div>

        {node.confirmingDelete && (
          <div style={s.confirmBar}>
            <span>Delete &quot;{node.name}&quot; and everything inside it? This moves it to the site Recycle Bin.</span>
            <button style={s.dangerBtn} disabled={busy} onClick={() => confirmDelete(node.id)}>Yes, delete</button>
            <button style={s.ghostBtn} disabled={busy} onClick={() => cancelDelete(node.id)}>Cancel</button>
          </div>
        )}

        {!node.isDeleted && renderPermPanel(node)}

        {!node.isDeleted && (
          <div style={s.childrenPane}>
            {isOpen && node.children.map(child => renderNode(child, mode, depth + 1))}
            {renderAddForm(mode, node)}
          </div>
        )}
      </div>
    );
  };

  /* ── Render ──────────────────────────────────────────────────────────────────── */

  /**
   * The retired manual folder tree. Always false now that the tab bar cannot reach `Staging` or
   * `Documents` — kept as a named test so the tree's own controls stay attached to the tree rather
   * than to "not Reconciliation", which since this restructure also means Abbreviations, Levels,
   * Migrate and New segment. That inverted read would have rendered a Refresh/Update pair over
   * every one of the mounted screens, each with its own Save.
   */
  const treeTab = tab === "Staging" || tab === "Documents";

  return (
    <section style={s.wrap}>
      <style>{`.fm-in:focus { outline: none; box-shadow: 0 0 0 2px rgba(15,108,63,.18); }`}</style>

      <h2 style={s.h2}>Folder Administration</h2>
      <p style={s.subtitle}>
        Name the folders a segment&rsquo;s terms produce, shape the levels beneath Unit, move what is
        already filed, then build the tree. Who can see a folder is set on the{" "}
        <strong>Folder Access</strong> page.
      </p>

      {/*
        One home for folder administration — spec `2026-08-12-term-abbreviation-page-design.md` §6.
        The order is the order the work happens in: name the terms, shape the levels, move what is
        already filed, reconcile. Segments sits last because it is the rarest — it both creates a
        segment and, since 2026-08-14, deletes one.

        `Staging` and `Documents` are gone (client, 2026-08-12: "I am honestly not using it"). They
        were a manual folder tree — reconciliation and the Folder Access page now cover it from data.
      */}
      <div style={s.toggleWrap}>
        <div style={s.seg}>
          {([
            ["Abbreviations",  "Term Abbreviations"],
            ["Levels",         "Folder levels"],
            ["Migrate",        "Move existing folders"],
            ["Reconciliation", "Folder Reconciliation"],
            ["NewSegment",     "Segments"],
          ] as Array<[Tab, string]>).map(([t, label], i, arr) => (
            <button key={t}
              onClick={() => {
                if (t === tab) return;
                // An unsaved edit refuses the switch rather than losing it — see `dirty`.
                if (dirty) { setTabBlocked(true); return; }
                setTabBlocked(false);
                setTab(t);
                setReconConfirm(false);
              }}
              style={{ ...s.segBtn, ...(i === arr.length - 1 ? { borderRight: "none" } : {}), ...(tab === t ? s.segActive : {}) }}
            >
              {label}
            </button>
          ))}
        </div>
      </div>

      {tabBlocked && (
        <div style={{ fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5, background: "#fff4e5", border: "1px solid #f0d9b5", color: "#7a4f00" }}>
          Finish or clear what you are editing first — leaving this tab would lose it.
        </div>
      )}

      {tab === "Abbreviations" ? (
        <AbbreviationManager
          context={context}
          siteUrl={siteUrl}
          onDirtyChange={(d) => {
            setDirty(d);
            // Clear the refusal as soon as its reason is gone, so a saved edit does not leave a
            // warning telling them to do what they just did.
            if (!d) setTabBlocked(false);
          }}
        />
      ) : tab === "Levels" ? (
        <StructureManager
          context={context}
          siteUrl={siteUrl}
          onDirtyChange={(d) => { setDirty(d); if (!d) setTabBlocked(false); }}
        />
      ) : tab === "Migrate" ? (
        <SubtreeMigrator context={context} siteUrl={siteUrl} />
      ) : tab === "NewSegment" ? (
        // Same dirty guard: a half-typed segment costs more to retype than a level edit, and
        // losing it to a tab click would be the same silent discard.
        <SegmentCreator
          context={context}
          siteUrl={siteUrl}
          onDirtyChange={(d) => { setDirty(d); if (!d) setTabBlocked(false); }}
        />
      ) : tab === "Reconciliation" ? (
        <div>
          <p style={{ fontSize: 13, color: "#444", lineHeight: 1.5, margin: "0 0 16px" }}>
            Build the folder tree from the <strong>term store</strong> in both{" "}
            <strong>Staging</strong> and <strong>Documents</strong>. Every level
            (segment → department → unit) is created if missing and{" "}
            <strong>locked</strong> (inheritance broken) so nobody can see a folder
            until its group is assigned. Groups are then{" "}
            <strong>auto-assigned from DMS Group Map</strong> by role (MEMBER → Read,
            UPL → Contribute, APR → Design; Documents gets viewer/MEMBER groups only).
            Under each unit the full <strong>Year × Document Type</strong> grid is
            pre-created (inheriting the unit folder permissions). Staging folders are
            also mapped for rename-proof upload routing. Safe to re-run: existing
            folders and mappings are not duplicated, and role assignments merge. A
            tier with no group-map rows is flagged in the log. Note: a full run can
            take a minute or two.
          </p>
          {reconConfirm ? (
            <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "4px 0 8px", padding: "8px 10px", background: "#f0f7f2", border: "1px solid #cfe4d8", borderRadius: 4, fontSize: 12, color: "#0f6c3f" }}>
              <span>
                Create + lock the term-store folder tree in <strong>Staging</strong> and <strong>Documents</strong>, then auto-assign groups from DMS Group Map by role. Safe to re-run.
                <br />
                <strong>Keep this tab open until it finishes.</strong> The run happens in your browser — refreshing, closing the tab or navigating away stops it partway. Nothing is lost and you can simply run it again, but do re-run before letting users in: a folder interrupted at the wrong moment stays unlocked until the next run.
              </span>
              <button style={{ ...s.btn, padding: "5px 14px", fontSize: 12, background: "#0f6c3f", color: "#fff", border: "none", flexShrink: 0 }} disabled={busy} onClick={() => { runReconciliation().catch(() => undefined); }}>
                {busy ? "Running…" : "Yes, run reconciliation"}
              </button>
              <button style={s.ghostBtn} disabled={busy} onClick={() => setReconConfirm(false)}>Cancel</button>
            </div>
          ) : (
            <button
              onClick={() => setReconConfirm(true)}
              disabled={busy}
              style={{ ...s.btn, background: "#0f6c3f", color: "#fff", border: "none" }}
            >
              Run reconciliation
            </button>
          )}

          {(reconRunning ||
            folderFeeds.Staging.length > 0 || folderFeeds.Documents.length > 0 ||
            assignFeeds.Staging.length > 0 || assignFeeds.Documents.length > 0) && (
            <div style={{ marginTop: 18 }}>
              <style>{"@keyframes fmspin{to{transform:rotate(360deg)}}"}</style>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                {reconRunning && (
                  <span style={{ display: "inline-block", width: 16, height: 16, border: "2px solid #cfe4d8", borderTopColor: "#0f6c3f", borderRadius: "50%", animation: "fmspin 0.8s linear infinite" }} />
                )}
                <strong style={{ fontSize: 13, color: "#0f6c3f" }}>
                  {reconRunning ? (reconPhase || "Working…") : "Reconciliation complete"}
                  {reconRunning && (
                    <span style={{ marginLeft: 8, fontWeight: 400, color: "#8a5a00" }}>
                      — keep this tab open; leaving stops the run
                    </span>
                  )}
                </strong>
                <span style={{ fontSize: 12, color: "#666" }}>
                  {reconCounts.folders} folders created · {reconCounts.assigns} groups assigned
                </span>
                {(() => {
                  // reconDone, NOT reconCounts: the estimate must divide by steps
                  // ATTEMPTED. reconCounts only rises when something changed, so on a
                  // re-run it sits near zero against a full planned total and the
                  // estimate runs away — the "~197m left" on a three-minute re-run.
                  const elapsedMs = reconStartMs ? Math.max(0, reconNow - reconStartMs) : 0;
                  if (!reconRunning) {
                    // Final line after a run completes.
                    return reconStartMs ? (
                      <span style={{ fontSize: 12, color: "#666" }}>· took {fmtDur(elapsedMs)}</span>
                    ) : null;
                  }
                  const remainingOps = Math.max(0, reconPlanned - reconDone);
                  // Use the live rate once there are enough samples to mean anything.
                  // 25, not 5: the first steps are all cheap "already there" skips, and
                  // extrapolating a whole run from them under-reads it as badly as the
                  // old counter over-read it.
                  const remainMs = reconDone >= 25
                    ? remainingOps * (elapsedMs / reconDone)
                    : remainingOps * (RECON_WRITE_DELAY_MS + RECON_EST_HTTP_MS)
                      + Math.floor(reconPlanned / RECON_BATCH_SIZE) * RECON_COOLDOWN_MS;
                  const pct = reconPlanned > 0
                    ? Math.min(100, Math.round((reconDone / reconPlanned) * 100))
                    : 0;
                  return (
                    <span style={{ fontSize: 12, color: "#0f6c3f", fontWeight: 600 }}>
                      · {pct}% · ~{fmtDur(remainMs)} left
                      <span style={{ color: "#999", fontWeight: 400 }}> ({fmtDur(elapsedMs)} elapsed{reconPlanned > 0 ? `, ${reconDone}/${reconPlanned} steps` : ""})</span>
                    </span>
                  );
                })()}
              </div>
              {reconPlanned > 0 && reconRunning && (
                <div style={{ height: 6, borderRadius: 4, background: "#ececec", overflow: "hidden", marginBottom: 12 }}>
                  <div style={{ height: "100%", background: "#0f6c3f", borderRadius: 4, transition: "width .3s ease", width: `${Math.min(100, (reconDone / reconPlanned) * 100)}%` }} />
                </div>
              )}
              {/* One row per library, two panels each. Grouping by library rather than
                  by kind is what the client actually reads: "is Staging done" is a
                  question, "are all folders done across both libraries" is not. */}
              {(["Staging", "Documents"] as LibTarget[]).map((lib) => (
                <div key={lib} style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#0f6c3f", textTransform: "uppercase", letterSpacing: ".05em", marginBottom: 6 }}>{lib}</div>
                  {/* flexWrap + flex-basis makes the two panels sit side-by-side on wide
                      screens and stack on narrow (mobile) — no media query needed. */}
                  <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                    {([
                      { title: `${lib} folders`, feed: folderFeeds[lib] },
                      { title: `${lib} group assignments`, feed: assignFeeds[lib] },
                    ] as const).map((panel) => (
                      <div key={panel.title} style={{ flex: "1 1 280px", minWidth: 0, border: "1px solid #e5e5e5", borderRadius: 4, overflow: "hidden" }}>
                        <div style={{ padding: "6px 10px", background: "#f7f7f7", fontSize: 12, fontWeight: 600, color: "#444", borderBottom: "1px solid #eee" }}>{panel.title}</div>
                        {/* overflowX:auto lets the client slide left/right to read full paths;
                            rows keep nowrap (no ellipsis clip) so the whole message is reachable. */}
                        <div style={{ maxHeight: 220, overflowY: "auto", overflowX: "auto", padding: "4px 0" }}>
                          {panel.feed.length === 0 ? (
                            <div style={{ padding: "6px 10px", fontSize: 12, color: "#aaa" }}>—</div>
                          ) : (
                            panel.feed.map((it, i) => (
                              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "3px 10px", fontSize: 12, color: it.status === "admin" ? "#b45309" : "#333" }}>
                                {progIcon(it.status)}
                                <span style={{ whiteSpace: "nowrap" }}>{it.text}</span>
                              </div>
                            ))
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      ) : loading ? (
        <p style={{ fontSize: 13, color: "#666" }}>Loading folders…</p>
      ) : (
        /* UNREACHABLE from the tab bar: the manual folder tree below served the retired
           `Staging`/`Documents` tabs. Kept compiling so its removal is a separate, reviewable
           change rather than a 200-line deletion buried in a tab restructure. */
        <>
          {sections.length === 0 && (tree[NEW_TOP_LEVEL] ?? []).length === 0 && (
            <p style={{ fontSize: 13, color: "#999" }}>
              No top-level folders under {libTarget} yet — create the first one below.
            </p>
          )}

          {sections.map(mode => {
            const open = modeOpen[mode] !== false; // sections default to open
            const nodes = tree[mode] ?? [];
            return (
              <div key={mode} style={{ marginBottom: 28 }}>
                <div style={s.secHeader} onClick={() => setModeOpen(prev => ({ ...prev, [mode]: open ? false : true }))}>
                  <span style={s.ico}>{open ? "▾" : "▸"}</span>
                  <p style={s.secTitle}>{mode}</p>
                  {nodes.length > 0 && <span style={s.badge}>{nodes.length}</span>}
                </div>

                {open && (
                  <>
                    {nodes.length === 0 ? (
                      <p style={{ fontSize: 13, color: "#999" }}>No folders found under {libTarget}/{mode}.</p>
                    ) : (
                      <div style={s.scrollPane}>
                        {nodes.map(node => renderNode(node, mode, 0))}
                      </div>
                    )}

                    {libRoot && renderAddForm(mode, null)}
                  </>
                )}
              </div>
            );
          })}

          {/* New top-level folders staged for creation directly under the library root */}
          {(tree[NEW_TOP_LEVEL] ?? []).length > 0 && (
            <div style={{ marginBottom: 20 }}>
              <p style={s.secTitle}>New top-level folder{(tree[NEW_TOP_LEVEL] ?? []).length !== 1 ? "s" : ""}</p>
              <div style={s.scrollPane}>
                {(tree[NEW_TOP_LEVEL] ?? []).map(node => renderNode(node, NEW_TOP_LEVEL, 0))}
              </div>
            </div>
          )}

          {libRoot && (
            <div style={{ paddingTop: 4 }}>
              <button style={s.addFolderBtn} disabled={busy} onClick={addTopLevel}>
                + Add top-level folder
              </button>
            </div>
          )}
        </>
      )}

      {treeTab && (
        <div style={s.actions}>
          <button onClick={() => loadTree().catch(() => undefined)} disabled={busy || loading}
            style={{ ...s.btn, marginRight: "auto", background: "#fff", color: "#0f6c3f", border: "1px solid #0f6c3f" }}>
            Refresh
          </button>
          <button onClick={() => { handleUpdate().catch(() => undefined); }}
            disabled={busy || loading || !hasChanges}
            style={{ ...s.btn, background: !busy && !loading && hasChanges ? "#0f6c3f" : "#9bbfaa", color: "#fff", border: "none", cursor: !busy && !loading && hasChanges ? "pointer" : "default" }}>
            {busy ? "Updating…" : "Update"}
          </button>
        </div>
      )}

      {/* The log belongs to the run that produced it. Ungated it would sit under the mounted
          structure screens too, where a stale reconciliation report reads as that screen's output. */}
      {(tab === "Reconciliation" || treeTab) && log.length > 0 && (() => {
        /* The log is split by WHERE and by SEVERITY, because those answer different
           questions: "what happened in Documents" and "what do I have to fix".
           Warnings and Errors are filtered VIEWS, so an entry appears both in its
           library tab and in its severity tab — that is the point of triage.

           Severity comes from the marker the message carries, not from `ok`. The
           two disagree on purpose elsewhere in this file: several genuine warnings
           are pushed with ok:true so they cannot gate the orphan prune, and a
           missing abbreviation is pushed with ok:false. The glyph is the author's
           actual intent; the flag is a control signal. */
        const isError = (e: LogEntry): boolean => /✗|FAILED|✖/.test(e.msg) || (!e.ok && !/⚠|\?/.test(e.msg));
        const isWarning = (e: LogEntry): boolean => !isError(e) && /⚠|(^|\s)\?\s/.test(e.msg);
        const errors = log.filter(isError);
        const warnings = log.filter(isWarning);
        // An entry naming both libraries belongs to both; one naming neither (the
        // prune, orphan repair and site-entry passes) is reachable only from All,
        // which is why All exists and is the default.
        const forLib = (lib: string): LogEntry[] => log.filter((e) => e.msg.indexOf(lib) !== -1);
        const tabs = [
          { key: "All", rows: log },
          { key: "Documents", rows: forLib("Documents") },
          { key: "Staging", rows: forLib("Staging") },
          { key: "Warnings", rows: warnings },
          { key: "Errors", rows: errors },
        ] as const;
        const active = tabs.find((t) => t.key === logTab) ?? tabs[0];
        const colourOf = (e: LogEntry): string =>
          isError(e) ? "#d13438" : isWarning(e) ? "#b45309" : "#0f6c3f";
        const glyphOf = (e: LogEntry): string =>
          isError(e) ? "✗" : isWarning(e) ? "⚠" : "✓";
        // Most messages were written with their own leading glyph, from before the
        // renderer added one — hence "⚠ ⚠ Documents/…". Strip it at RENDER only:
        // isError/isWarning classify by reading that glyph out of the text, so
        // removing it from the stored message would silently demote every warning
        // to a success. Leading spaces are preserved because indentation is how
        // per-folder detail lines are nested under their folder.
        const textOf = (e: LogEntry): string =>
          e.msg.replace(/^(\s*)[✓⚠✗✖]\s*/, "$1");
        return (
          <div style={s.logBox}>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
              {tabs.map((t) => {
                const on = t.key === active.key;
                const alert = (t.key === "Errors" && t.rows.length > 0)
                  ? "#d13438"
                  : (t.key === "Warnings" && t.rows.length > 0) ? "#b45309" : undefined;
                return (
                  <button
                    key={t.key}
                    onClick={() => setLogTab(t.key)}
                    style={{
                      padding: "5px 12px", borderRadius: 14, fontSize: 12, fontWeight: 600,
                      cursor: "pointer", fontFamily: "inherit",
                      border: `1px solid ${on ? (alert ?? "#0f6c3f") : "#d8d8d8"}`,
                      background: on ? (alert ?? "#0f6c3f") : "#fff",
                      color: on ? "#fff" : (alert ?? "#555"),
                    }}
                  >
                    {t.key} ({t.rows.length})
                  </button>
                );
              })}
            </div>
            {active.rows.length === 0 ? (
              <div style={{ fontSize: 12, color: active.key === "Errors" || active.key === "Warnings" ? "#0f6c3f" : "#999" }}>
                {active.key === "Errors"
                  ? "No errors. 🎉"
                  : active.key === "Warnings"
                    ? "No warnings — every folder got a group."
                    : "Nothing logged for this view."}
              </div>
            ) : (
              // Scrolls rather than growing the page: a full run logs thousands of
              // lines and the tab bar has to stay reachable.
              <div style={{ maxHeight: 420, overflowY: "auto" }}>
                {active.rows.map((entry, i) => (
                  <div key={i} style={{ fontSize: 12, color: colourOf(entry), marginBottom: 4, wordBreak: "break-word" }}>
                    {glyphOf(entry)} {textOf(entry)}
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      {toast && (
        <div style={{ ...s.toast, background: toast.error ? "#d13438" : "#0f6c3f" }}>
          {toast.message}
          <button onClick={() => setToast(null)} style={s.toastClose}>✕</button>
        </div>
      )}
    </section>
  );
}
