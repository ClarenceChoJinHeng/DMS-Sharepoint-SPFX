import * as React from "react";
import { useState, useEffect } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Level, parseLevels, PENDING_LEVELS_FIELD, sanitizeFolderSegment } from "../../../shared/formModel";
import {
  allowedTierPositions, builtInTierFor, canMoveBelowUnitTier, clampTierPosition,
  effectiveOnDemandTiers, isFixedBelowUnitTier, perUnitTierExists, splitChain, validateChain,
} from "../../../shared/folderChain";
import { EVENT } from "../../../shared/auditLog";
import { allLibraryTitles, cachedHcLibraries, cachedListTitle, documentsLibraryTitle, libraryTitle, libraryUrlSegment, LIST_SUFFIX } from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
import { writeAudit } from "../../../shared/spAuditLog";
import { ensureColumn } from "../../../shared/spColumns";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";
import {
  existingColumnReason,
  reservedColumnReason,
} from "../../../shared/tierColumnGuard";

/**
 * Folder Structure — add, reorder and remove the folder levels BENEATH Unit.
 *
 * Piece 2 of `2026-08-06-configurable-folder-structure-chain-design.md`, specced in
 * `2026-08-10-structure-manager-ui-design.md`. Piece 1 made the structure data-driven;
 * it lives as `Levels` JSON on a DMS Config `mode` row, and the client cannot edit JSON —
 * which is the whole reason the feature exists. This is the editor that closes that.
 *
 * THE JSON IS NEVER SHOWN. An admin edits a list of levels and sees the resulting path.
 *
 * Scope, deliberately: below-Unit levels only. The permissioned tiers (Segment, Department,
 * Unit) carry folder ACLs and are set at provisioning time; renaming one renames live ACL'd
 * folders and reordering one asks reconciliation to walk a tree that does not exist. Neither
 * is a config edit, so they render locked.
 */

/* ⚠ RESOLVED, NOT LITERAL (2026-08-28). This screen creates the tier columns in BOTH libraries, so
   a 404 here means the approved side silently never gets them — and a MOVE carries over only the
   columns that EXIST at the destination. That is the `Remark`/`LegallyPrivileged` gap of
   2026-08-10 exactly, arriving through a renamed library. */
const documentsListTitle = (): string => documentsLibraryTitle();
/** Documents' URL segment differs from its title, like the approval library's does. */
const DOCUMENTS_URL_SEGMENT = "Shared Documents";

/** A DMS Config `mode` row, reduced to what this screen edits. */
interface SegmentRow {
  id: number;
  key: string;           // Title, e.g. "mode_gho"
  label: string;         // ModeLabel, e.g. "Group Head Office"
  stagingFolder: string; // top folder name, e.g. "GHO"
  chain: Level[];
  /**
   * A structure change authored but not yet applied. Present only on a segment that already
   * holds documents; applying it is the last step of the folder migration. See
   * PENDING_LEVELS_FIELD.
   */
  pending?: Level[];
  /** Parsed-chain problem, if any — shown instead of letting them edit a broken row. */
  chainError?: string;
  /** Whether documents are already filed under this segment. Drives the warning. */
  hasDocuments?: boolean;
  /**
   * The segment's own term set. Read separately (see `loadSegmentTermSets`) and therefore optional:
   * absent means the sub unit check cannot run, which warns rather than blocking.
   */
  termSetGuid?: string;
}

/** The in-progress "add a level" form. */
interface DraftTier {
  label: string;
  /**
   * TRUE = the values sit under each Unit in the term store, so every unit offers its own — the
   * SubUnit shape (client, 2026-08-10; asked for as an explicit button 2026-08-17).
   *
   * It exists because the stored discriminator is the ABSENCE of `termSet`, and a blank text box is
   * a poor way to express a mode. The admin who needs this most is the one who does not know that
   * leaving a field empty is a decision — they would paste some other segment's set ID, and every
   * unit would then be offered every other unit's subunits, which reads as working.
   *
   * When true the GUID is CLEARED rather than merely ignored, so a half-typed ID cannot survive into
   * the saved chain by a later toggle back.
   */
  fromUnit: boolean;

  termSetGuid: string; // only meaningful when fromUnit is false
  position: number;    // index within the below-Unit list
}

/**
 * What we know about the term set ID the admin typed.
 *
 * A wrong ID here is the worst failure this screen can produce, because nothing reports
 * it: the tier saves, reconciliation ignores below-Unit tiers entirely, and the only
 * symptom is an upload form with one permanently empty dropdown that blocks the upload.
 * The admin who typed it is not the person who hits it. So resolve the ID against the
 * term store while they are still looking at the field.
 *
 * `unknown` is deliberately NOT treated as bad. A 500 or a dropped connection says
 * nothing about the ID, and refusing to save on it would block a correct one.
 */
/**
 * What we know about whether any unit in this segment has sub unit terms under it.
 *
 * `none` is the only state that blocks Add, and it is the only one that is a definite negative:
 * the walk finished, every unit term was read, and not one of them has a child. Everything else —
 * a failed read, a capped walk, a segment with no term set recorded — is `unknown` and warns.
 */
type UnitCheck =
  | { state: "idle" }
  | { state: "checking" }
  | { state: "none"; units: number }
  | { state: "some"; withTerms: number; units: number }
  | { state: "unknown" };

type SetCheck =
  | { state: "blank" }
  | { state: "malformed" }
  | { state: "checking" }
  /**
   * `nested` — how many of the top-level terms have terms of their OWN.
   *
   * ⚠ OPTIONAL, AND `undefined` MEANS "NOT ESTABLISHED" rather than none. A shared level only ever
   * offers the TOP level of its set (`sets/{guid}/children`), so anything authored beneath those
   * terms can never appear — the client hit this with a set whose second term had a child: the screen
   * said *"Found — 2 values"*, enabled Add, and would have ignored the nested term for ever
   * (2026-09-08: *"it doesn't really show an error or refuse if the Shared folder term Term structure
   * is not the same as Year"*).
   *
   * ⚠ IT REFUSES ADD SINCE 2026-09-09, at the client's instruction (*"isn't it better to force that
   * the Add doesn't work until they completely ensure the Term level follows the Year"*) — reversing
   * the warn-and-allow it shipped with a day earlier. Only a DEFINITE positive blocks: `undefined`
   * allows, because a tenant that does not report `childrenCount` must not lose the ability to add a
   * perfectly flat set.
   */
  | { state: "found"; name: string; count: number; nested?: number }
  | { state: "notfound" }
  | { state: "unknown"; status: number };

/** Accepts the braced/whitespaced forms people paste out of SharePoint UI and URLs. */
function normalizeGuid(raw: string): string {
  return raw.trim().replace(/^[{(]|[)}]$/g, "").trim().toLowerCase();
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

/**
 * Derive the column internal name from the level name — never typed by the admin.
 *
 * A SharePoint column's internal name is fixed at creation, permanently, from the title
 * it is created with. Letting an untrained admin type it offers a decision that cannot
 * be undone and whose consequence is invisible until a metadata write silently fails.
 * So strip everything SharePoint would encode, create under that, then set the display
 * title to what they typed.
 */
function columnNameFor(label: string): string {
  return sanitizeFolderSegment(label).replace(/[^A-Za-z0-9]/g, "");
}

/** The verdict, in the admin's language. Never shows the GUID back — they can see it. */
/**
 * The sub unit verdict, as something to render.
 *
 * A NODE rather than a string, because the `none` case has to carry a LINK. Everything else is
 * plain text, so the caller does not have to care which it got.
 *
 * ⚠ THE TERM STORE LINK GOES TO THE CLASSIC, SITE-LEVEL PAGE (`termstoremanager.aspx`), and that is
 * load-bearing rather than incidental — this is its third mount point in the project, all three the
 * same URL for the same reason. An admin looking for the term store naturally reaches the MODERN one
 * (`/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter`), which is the TENANT admin centre and
 * answers "Access denied — you don't have access to this admin center operation" to a site
 * collection administrator, because it is gated on being a SharePoint Administrator for the whole
 * tenant. The client hit that wall themselves on 2026-08-30. Built from `siteUrl`, never hardcoded.
 */
function unitCheckNote(
  c: UnitCheck, unitLabel: string, segLabel: string, siteUrl: string,
): React.ReactNode {
  switch (c.state) {
    case "checking":
      return `Checking whether any ${unitLabel.toLowerCase()} has sub unit terms under it…`;
    case "none":
      /* ⚠ THE CLIENT'S OWN WORDING (2026-09-08), and shorter than what it replaced for a reason:
         they read the first version and asked what it meant. "authored under it" is term-store
         jargon, and the person who meets this next is further from the term store than they are. It
         NAMES THE SEGMENT, because this screen lists five and the message is about one of them. */
      return (
        <>
          {segLabel} does not have any sub units under its {unitLabel} terms yet. Add one in the{" "}
          <a
            href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
            target="_blank"
            rel="noopener noreferrer"
          >
            Term Store
          </a>{" "}
          first, then you can add a Sub unit level here.
        </>
      );
    case "some":
      // The COUNT is the useful part, and the admin has no other way to get it: it says at a glance
      // whether the terms were authored where they meant to author them. `some` is never a warning -
      // a partial spread is the designed shape, not a problem to report.
      return (
        `${c.withTerms} of ${c.units} ${unitLabel.toLowerCase()}s have sub unit terms under them. ` +
        `The level is only offered where they exist; the rest file one level shallower, which needs ` +
        `no configuration.`
      );
    case "unknown":
      return (
        `Could not check whether sub unit terms exist — the term store did not answer, or this ` +
        `segment has no term set recorded. Adding the level is still allowed; if no ` +
        `${unitLabel.toLowerCase()} has terms under it, the level simply never appears.`
      );
    default:
      return "";
  }
}

function setCheckMessage(c: SetCheck): string {
  switch (c.state) {
    case "checking":
      return "Checking…";
    case "malformed":
      return "That is not a term set ID. Copy the ID from the term store — it looks like " +
        "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf.";
    case "notfound":
      return "No term set with that ID exists on this site. If you copied it from another " +
        "site, or copied a term instead of the term set, it will not work here.";
    case "unknown":
      return `Could not check that ID right now${c.status ? ` (HTTP ${c.status})` : ""}. ` +
        "You can still add the level, but confirm the ID is right — a wrong one leaves the " +
        "dropdown empty and blocks uploads.";
    case "found": {
      if (c.count === 0) {
        return `Found "${c.name}", but it has no terms yet. Add the terms before anyone uploads, ` +
          "or the dropdown will be empty and block them.";
      }
      const found = `Found "${c.name}" — ${c.count} ${c.count === 1 ? "value" : "values"}.`;
      /* A shared level offers the TOP level of its set and nothing below, so a nested term
         contributes one option and hides its own children — the client's set had a term with a child
         and the screen called it two clean values.
         ⚠ THIS MESSAGE IS PAIRED WITH THE REFUSAL IN `canAddTier`, and the two read the same
         condition. It ends by naming the fix rather than offering to proceed, because proceeding is
         no longer possible: an offer beside a disabled button is how a screen reads as broken.
         `undefined` says nothing at all: the tenant may not report `childrenCount`, and "we could
         not establish it" must not read as "there is none". */
      if (c.nested === undefined || c.nested === 0) return found;
      return (
        `${found} ⚠ ${c.nested} of them ${c.nested === 1 ? "has" : "have"} terms nested underneath. ` +
        "A shared level only ever offers the top level — like Year and Document Type — so those " +
        "nested terms will never appear as folders. Ensure that the Term Set is only one level."
      );
    }
    default:
      return "";
  }
}

/**
 * Whether a re-check is worth offering.
 *
 * ⚠ ONLY WHERE THE VERDICT CAME FROM THE SERVER. `blank` and `malformed` are decided locally from
 * the field itself, so re-asking cannot change either — a button there would be a control that
 * visibly does nothing, which is how the ones that DO work stop being trusted.
 */
function setCheckIsRecheckable(c: SetCheck): boolean {
  return (
    c.state === "checking" || c.state === "found" ||
    c.state === "notfound" || c.state === "unknown"
  );
}

function setCheckStyle(c: SetCheck): React.CSSProperties {
  // Red is what BLOCKS Add, amber is what merely warns — so a nested set moved from amber to red
  // when it started blocking. Keeping it amber beside a disabled button would put a "you may
  // continue" colour on the one verdict that stops the admin.
  if (
    c.state === "notfound" || c.state === "malformed" ||
    (c.state === "found" && (c.nested ?? 0) > 0)
  ) {
    return { color: "#a4262c" };
  }
  if (c.state === "unknown" || (c.state === "found" && c.count === 0)) {
    return { color: "#7a4f00" };
  }
  if (c.state === "found") return { color: "#0f6c3f" };
  return {};
}

/**
 * Block Add only where the ID is KNOWN to be wrong.
 *
 * A `found` set with no terms is allowed through: authoring the tier before the terms is a
 * legitimate order of work, and the warning says what is still owed. `unknown` is allowed
 * for the reason on `SetCheck`. `checking` blocks for the few hundred ms it lasts, so a
 * fast click cannot outrun the verdict.
 */
function canAddTier(
  adding: DraftTier, check: SetCheck, unitCheck: UnitCheck, perUnitTaken: boolean,
): boolean {
  const label = adding.label.trim();
  if (!label || !columnNameFor(label)) return false;
  // One sub unit per unit — a second per-unit tier is refused rather than offered a slot.
  if (adding.fromUnit && perUnitTaken) return false;
  // A per-unit tier has no ID to validate, so the TERM-SET verdict must not gate it. Without this
  // the two could disagree: `setCheck` is computed from whatever is in the GUID field, and a
  // "notfound" left over from a paste before the toggle would keep Add disabled with a message
  // about a term set the admin can no longer see.
  //
  // What DOES gate it is the unit walk, and only its one definite negative. `checking` blocks for
  // as long as the walk lasts, so a fast click cannot outrun the verdict; `unknown` never blocks,
  // because an unreachable term store says nothing about whether the terms are there.
  if (adding.fromUnit) return unitCheck.state !== "none" && unitCheck.state !== "checking";
  // A SHARED-LIST tier must actually name a set. Blank is not neutral here — it is the stored
  // discriminator for per-unit, so saving it would silently produce the other kind of tier from a
  // screen that says "one shared list".
  if (!normalizeGuid(adding.termSetGuid)) return false;
  // ⚠ A NESTED SET IS REFUSED (2026-09-09, client's instruction). A shared level only ever offers the
  // top level, so depth is silently discarded — and the client's rule is that such a set must be
  // flattened before it is used, not used in the knowledge that part of it is ignored.
  //
  // `?? 0` is what keeps this safe: `nested` is `undefined` when the depth could not be established
  // (the read failed, or the tenant does not report `childrenCount`), and unknown must never block —
  // otherwise a tenant that never reports the field could add no shared level at all.
  if (check.state === "found" && (check.nested ?? 0) > 0) return false;
  return check.state !== "malformed" && check.state !== "notfound" && check.state !== "checking";
}

const s: Record<string, React.CSSProperties> = {
  msg:       { fontSize: 13, padding: "10px 12px", borderRadius: 6, marginBottom: 16, lineHeight: 1.5 },
  err:       { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { ...NOTICE_ATTENTION },
  ok:        { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  card:      { border: "1px solid #e6e6e6", borderRadius: 12, padding: "18px 20px", marginBottom: 14, background: "#fff" },
  radioRow:  { display: "flex", gap: 8, alignItems: "flex-start", fontSize: 13, lineHeight: 1.5, cursor: "pointer" },
  radioHint: { fontSize: 12, color: "#605e5c" },
  segRow:    { display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 16, flexWrap: "wrap" },
  segName:   { fontSize: 16, fontWeight: 600, color: "#242424" },
  /* ⚠ NOT MONOSPACE ANY MORE (client design, 2026-08-30). The chain reads as a BREADCRUMB —
     `GHO › [Department] › [Unit]` — which is how it is drawn and how people read a path; the fixed
     pitch made it look like a stored value rather than a shape. `wordBreak` stays: a deep chain on a
     narrow screen must wrap rather than widen the card. */
  path:      { fontSize: 12.5, color: "#605e5c", marginTop: 6, wordBreak: "break-word", lineHeight: 1.6 },
  badge:     { fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em", padding: "4px 10px", borderRadius: 999, whiteSpace: "nowrap" },
  badgeUsed: { background: "#fff4e5", color: "#7a4f00" },
  badgeFree: { background: "#f1f8f4", color: "#0f6c3f" },
  /* ⚠ ITS OWN TONE, not `badgeUsed`. "In use" and "Change pending" appear TOGETHER on a staged
     segment, and two identical amber pills side by side read as one repeated badge rather than two
     different facts. */
  badgePending: { background: "#fdf1e7", color: "#8a4b00", border: "1px solid #f0d5a8" },
  /* ⚠ THE PANEL IS ABSOLUTELY POSITIONED, and that is only safe because nothing on this form is a
     scroll container — the file's one `overflow` is `s.modal`, a dialog this form never renders
     inside. A scroll cap has clipped an absolutely positioned child on three other screens in this
     project (Group Management's people picker, the upload form's two info panels, the member-add
     dropdown), so re-check this if a scroller is ever added here. */
  labelRow:  { display: "flex", alignItems: "center", gap: 5, marginTop: 14 },
  infoWrap:  { position: "relative", display: "inline-flex", alignItems: "center" },
  /* Sits on the label row, so it reads as help for THAT field rather than as a page-level action —
     the term set ID is the one value on this form that has to be fetched from somewhere else. */
  labelLink: { fontSize: 12, fontWeight: 400, marginLeft: "auto" },
  resultList: { margin: "6px 0 0", paddingLeft: 18 },
  resultItem: { marginTop: 3 },
  /* The `↻ Refresh` treatment My Submissions established, at label-row scale (12px, to sit beside
     `labelLink` rather than tower over it). Sits AFTER the Term Store link so the row reads in the
     order the work happens: open the store, fix the set, re-check. */
  recheckBtn: { border: "1px solid #c7c7c7", background: "#fff", borderRadius: 4, cursor: "pointer", fontSize: 12, lineHeight: 1, padding: "3px 8px", color: "#0f6c3f", fontWeight: 400 },
  recheckOff: { border: "1px solid #e1dfdd", background: "#f3f2f1", borderRadius: 4, cursor: "not-allowed", fontSize: 12, lineHeight: 1, padding: "3px 8px", color: "#a19f9d", fontWeight: 400 },
  infoBtn:   { display: "inline-flex", background: "none", border: "none", padding: 0, cursor: "pointer", lineHeight: 1, color: "#605e5c" },
  infoPanel: { position: "absolute", top: "calc(100% + 8px)", left: 0, zIndex: 30, width: 300, maxWidth: "80vw", padding: 14, background: "#fff", border: "1px solid #e1e1e1", borderRadius: 10, boxShadow: "0 4px 16px rgba(0,0,0,.12)", fontSize: 12.5, lineHeight: 1.55, color: "#323130", fontWeight: 400, textAlign: "left" },
  btn:       { background: "#0f6c3f", color: "#fff", border: "none", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  ghost:     { background: "#fff", color: "#1b1b1b", border: "1px solid #c8c8c8", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "pointer" },
  danger:    { background: "#fff", color: "#a4262c", border: "1px solid #e6b3b5", borderRadius: 4, padding: "4px 9px", fontSize: 12, cursor: "pointer" },
  /* The SAME red, at DIALOG scale. `danger` above is sized for the inline Remove button on a tier
     row - 4px/9px at 12px - and the discard dialog was using it beside `ghost` and `btn`, which are
     7px/14px at 13px. Three buttons in one row at two different sizes (client, 2026-09-08: "The
     discard button is not the same size"). Derived from `danger` rather than restated, so the red
     has one definition and only the metrics differ. */
  dangerLg:  { ...{ background: "#fff", color: "#a4262c", border: "1px solid #e6b3b5", borderRadius: 4, cursor: "pointer" }, padding: "7px 14px", fontSize: 13 },
  iconBtn:   { background: "#fff", border: "1px solid #c8c8c8", borderRadius: 4, padding: "3px 8px", fontSize: 12, cursor: "pointer", marginRight: 4 },
  off:       { background: "#f3f2f1", color: "#a19f9d", border: "1px solid #e1dfdd", borderRadius: 4, padding: "7px 14px", fontSize: 13, cursor: "not-allowed" },
  tierRow:   { display: "flex", alignItems: "center", gap: 12, padding: "12px 14px", borderRadius: 8, background: "#fafafa", marginBottom: 8, fontSize: 13, flexWrap: "wrap" },
  /* A LOCKED row is visibly a different KIND of row, not a disabled version of an editable one:
     the two lists sit one above the other, and the client's design distinguishes them by weight
     rather than by the absence of buttons alone. */
  tierLock:  { background: "#f3f2f1", color: "#605e5c" },
  tierName:  { fontWeight: 600, flex: "0 0 170px", color: "#242424" },
  tierMeta:  { fontSize: 11.5, color: "#8a8886", flex: "1 1 160px" },
  codeTog:   { display: "inline-flex", alignItems: "center", gap: 5, fontSize: 11.5, color: "#605e5c", marginRight: 10, cursor: "pointer", whiteSpace: "nowrap" },
  label:     { display: "block", fontSize: 12, fontWeight: 600, color: "#323130", margin: "14px 0 4px" },
  input:     { width: "100%", boxSizing: "border-box", padding: "7px 9px", fontSize: 13, border: "1px solid #c8c8c8", borderRadius: 4 },
  hint:      { fontSize: 11, color: "#8a8886", marginTop: 3, lineHeight: 1.5 },
  modalBg:   { position: "fixed", inset: 0, background: "rgba(0,0,0,.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 },
  modal:     { background: "#fff", borderRadius: 8, padding: 24, maxWidth: 520, width: "100%", maxHeight: "80vh", overflowY: "auto" },
};

export interface StructureManagerProps {
  context: WebPartContext;
  siteUrl: string;
  /**
   * Fired whenever the editor gains or loses unsaved changes, so the page can stop a tab
   * switch from discarding them. Optional — the component guards the browser-level exits by
   * itself, and a caller that does not care still gets those.
   */
  onDirtyChange?: (dirty: boolean) => void;
  /**
   * Fired after a structure change is SAVED, so a page hosting this can re-read what it thought it
   * knew.
   *
   * ⚠ THIS EXISTS BECAUSE A SAVE HERE INVALIDATES A FACT THE GUIDED FLOW READ AT MOUNT, and nothing
   * told it. `FolderAdmin` locks its migrate step on `pendingLevels`, which it copies out of a
   * segment list read once on `[siteUrl, reload]` — so saving a change on the levels step left the
   * NEXT step insisting *"There is no pending structure change to move to"* about the change just
   * made. Found live 2026-09-08.
   *
   * ⚠ AND RE-READING THE DERIVED FACTS DID NOT HELP, which is the subtle part: that flow already
   * re-reads its facts on every step change (`stepIdx` is in the effect's deps, added 2026-08-19 for
   * exactly this class of problem). But the effect COPIES the flag out of the segment list rather
   * than re-reading it, so re-running the consumer re-copied the same stale value. **Re-reading a
   * derived fact cannot fix staleness that lives in the upstream read.**
   */
  onSaved?: () => void;
}

export default function StructureManager({
  context,
  siteUrl,
  onDirtyChange,
  onSaved,
}: StructureManagerProps): React.ReactElement {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [segments, setSegments] = useState<SegmentRow[]>([]);
  const [editing, setEditing] = useState<string | undefined>(undefined);
  const [draft, setDraft] = useState<Level[]>([]);
  /**
   * The chain the editor was opened with, for comparison against `draft`.
   *
   * Held rather than re-derived because `startEdit` seeds from `effectiveOnDemandTiers`, so
   * "what was on screen when they started" is not simply the mode row's chain — comparing
   * against the row would report a segment on the built-in Year/Document Type pair as dirty
   * the instant it was opened, and an always-on warning is one nobody reads.
   */
  const [baseline, setBaseline] = useState<string>("");
  const [confirmDiscard, setConfirmDiscard] = useState(false);
  const [adding, setAdding] = useState<DraftTier | undefined>(undefined);
  /**
   * Why the last Add was refused, shown beside the button.
   *
   * Set ONLY by a failed Add — never on load or while typing. Marking a name as clashing while it is
   * half-typed would flag every level on its way to a valid name, which is how a form comes to feel
   * like it is arguing with you.
   */
  const [addError, setAddError] = useState<string | undefined>(undefined);
  const [busy, setBusy] = useState(false);
  /* `text` is the headline; `lines` are the bullets under it.
     ⚠ THE MESSAGE USED TO BE ONE PARAGRAPH and it ran to ~600 characters, naming all twelve
     created columns (client, 2026-09-09: *"Wow that message is long.. what are you trying to
     convey?"*). Nobody needs the list — it is two columns times six libraries — and the one fact
     that matters, that the change is NOT live yet, was buried in the middle of it. */
  const [result, setResult] = useState<
    { text: string; ok: boolean; lines?: string[] } | undefined
  >(undefined);
  // The Year / Document Type term sets, so a segment still running on the built-in pair can
  // be shown the levels it is EFFECTIVELY using rather than an empty list.
  const [legacySets, setLegacySets] = useState<{ year: string; docType: string }>({ year: "", docType: "" });
  const [setCheck, setSetCheck] = useState<SetCheck>({ state: "blank" });
  /* Bumped by the Re-check button, and read by the term-set effect purely as a trigger.
     ⚠ IT IS A COUNTER RATHER THAN A BOOLEAN, so a second press while a check is already settled
     still changes the value and still re-runs the effect. A flag would need clearing afterwards,
     and a missed clear is a button that works once. */
  const [recheck, setRecheck] = useState(0);
  const [unitCheck, setUnitCheck] = useState<UnitCheck>({ state: "idle" });
  /* The "why this name is permanent" panel. Opened on hover AND on click: hover does not
     exist on a touch screen, and this is the one explanation on the form that has to be
     reachable without a mouse. Declared with the other hooks, above every early return. */
  const [nameInfo, setNameInfo] = useState(false);

  const editingSegment = (): SegmentRow | undefined => segments.filter((x) => x.key === editing)[0];

  /* ---------- Load ------------------------------------------------------- */

  /**
   * Staged chains by config item id. Read in its own request so a site without the column
   * still loads: a `$select` naming a nonexistent field returns HTTP 400 for the whole query,
   * so folding this into the main read would break every site that has never staged a change.
   */
  const loadPendingChains = async (): Promise<Record<number, Level[]>> => {
    const out: Record<number, Level[]> = {};
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
          `/items?$select=Id,${PENDING_LEVELS_FIELD}&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return out; // column absent on this site: nothing is staged, by definition
      for (const r of ((await res.json()).value ?? []) as Array<Record<string, unknown>>) {
        const raw = r[PENDING_LEVELS_FIELD];
        if (typeof raw !== "string" || raw.trim() === "") continue;
        const parsed = parseLevels(raw);
        if (parsed.length > 0) out[Number(r.Id)] = parsed;
      }
    } catch {
      // Same reading: unreadable means nothing is staged. The migrator makes the same call and
      // reports its own failure, so a stale badge here cannot cause a wrong migration.
    }
    return out;
  };

  /**
   * The segment term-set GUID per config item id.
   *
   * ⚠ ITS OWN REQUEST, for the reason `loadPendingChains` above states: a `$select` naming a column
   * this site does not have returns 400 for the WHOLE query, and the main segment read failing is
   * the one that takes this screen down with "Could not read the configuration". `TermSetGuid` is
   * written by `SegmentCreator`, so a hand-authored `mode` row can legitimately lack it.
   *
   * A missing GUID is `undefined`, never blank-and-usable: it means the sub unit check cannot run,
   * which reads as `unknown` and therefore warns rather than blocking.
   */
  const loadSegmentTermSets = async (): Promise<Record<number, string>> => {
    const out: Record<number, string> = {};
    try {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
          `/items?$select=Id,TermSetGuid&$filter=ConfigType eq 'mode'&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );
      if (!res.ok) return out;
      for (const r of ((await res.json()).value ?? []) as Array<Record<string, unknown>>) {
        const raw = r.TermSetGuid;
        if (typeof raw === "string" && raw.trim() !== "") out[Number(r.Id)] = raw.trim();
      }
    } catch {
      // Unreadable: every segment reads as "no GUID", so the check reports unknown and warns.
    }
    return out;
  };

  const loadSegments = async (): Promise<SegmentRow[]> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
        // PendingLevels may not exist yet on a site that has never staged a change, and one
        // unknown name in $select fails the WHOLE request with 400 rather than returning null
        // (gotcha #11) — which would read as "the configuration list is unreadable". So it is
        // requested separately, below, where a 400 costs nothing.
        `/items?$select=Id,Title,ModeLabel,StagingFolder,Levels` +
        `&$filter=ConfigType eq 'mode'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`the configuration list returned HTTP ${res.status}`);
    const data = await res.json();
    const pending = await loadPendingChains();
    const termSets = await loadSegmentTermSets();
    const rows = ((data.value ?? []) as Array<{
      Id: number; Title?: string; ModeLabel?: string; StagingFolder?: string; Levels?: string;
    }>).map((r) => {
      const chain = parseLevels(r.Levels ?? "");
      const err = validateChain(chain);
      const row: SegmentRow = {
        id: r.Id,
        key: (r.Title ?? "").trim(),
        label: (r.ModeLabel ?? r.Title ?? "").trim(),
        stagingFolder: (r.StagingFolder ?? "").trim(),
        chain,
      };
      const staged = pending[r.Id];
      if (staged && staged.length > 0) row.pending = staged;
      if (termSets[r.Id]) row.termSetGuid = termSets[r.Id];
      // A row whose chain does not validate is shown but NOT editable: saving would
      // rewrite a structure nobody has seen, and the fix may lie outside this screen.
      if (err) row.chainError = err.message;
      return row;
    });
    return rows.sort((a, b) => a.label.localeCompare(b.label));
  };

  /**
   * The Year / Document Type term-set GUIDs from the `setting` rows the upload form reads.
   * Needed only to seed the built-in pair when a segment has no below-Unit levels yet.
   */
  const loadLegacyTermSets = async (): Promise<{ year: string; docType: string }> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')` +
        `/items?$select=Title,SettingValue&$filter=ConfigType eq 'setting'&$top=200`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) return { year: "", docType: "" };
    const data = await res.json();
    const rows = (data.value ?? []) as Array<{ Title?: string; SettingValue?: string }>;
    const get = (key: string): string =>
      (rows.filter((r) => (r.Title ?? "").trim() === key)[0]?.SettingValue ?? "").trim();
    return { year: get("termSet_yearPeriod"), docType: get("termSet_documentType") };
  };

  /**
   * Resolve a typed term set ID against the term store: does it exist, what is it called,
   * and how many top-level terms does it hold.
   *
   * The name is the part that actually catches mistakes. A GUID either resolving or not
   * only rules out typos; showing "Year / Period" back when they meant SubUnit catches the
   * commonest error by far — pasting the ID of the wrong set, or of a TERM instead of a
   * set. A term's GUID is well-formed and 404s here, which is exactly the outcome we want.
   *
   * The count is read separately because a set can exist and be empty, and an empty set
   * produces the same permanently-blank dropdown as a wrong ID.
   */
  const checkTermSet = async (guid: string): Promise<SetCheck> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      `${siteUrl}/_api/v2.1/termStore/sets/${guid}?$select=id,localizedNames`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json" } },
    );
    // 404 is the answer, not an error: no set with that ID is readable from this site.
    // Term sets are per-site here, so one copied from another site's config lands here too.
    if (res.status === 404) return { state: "notfound" };
    if (!res.ok) return { state: "unknown", status: res.status };

    const set = (await res.json()) as { localizedNames?: Array<{ name?: string }> };
    const name = (set.localizedNames ?? [])[0]?.name ?? "";

    let count = 0;
    try {
      const kids: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${guid}/children?$select=id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (kids.ok) count = (((await kids.json()).value ?? []) as unknown[]).length;
    } catch {
      count = 0; // reported as "0 terms" below, which is a warning and not a block
    }

    /* Is the set FLAT? A shared level shows `sets/{guid}/children` and nothing below it, so a term
       with children of its own contributes exactly one option and hides the rest.

       ⚠ ITS OWN REQUEST, AND ITS FAILURE CHANGES NOTHING. Folding `childrenCount` into the read
       above would risk the whole `$select` being rejected on a tenant that does not expose it —
       which would report a perfectly good set as having 0 terms, i.e. turn a missing warning into a
       false one. Here a failure, or a tenant that omits the field, simply leaves `nested`
       undefined and says nothing. */
    let nested: number | undefined;
    try {
      const deep: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/v2.1/termStore/sets/${guid}/children?$select=id,childrenCount`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json" } },
      );
      if (deep.ok) {
        const rows = (((await deep.json()).value ?? []) as Array<{ childrenCount?: number }>);
        // Absent on every row means the tenant does not report it — unknown, not zero.
        if (rows.filter((r) => typeof r.childrenCount === "number").length > 0) {
          nested = rows.filter((r) => (r.childrenCount ?? 0) > 0).length;
        }
      }
    } catch {
      // Leave it undefined: nothing is said rather than something wrong being said.
    }
    return { state: "found", name, count, nested };
  };

  /**
   * Do the units in this segment actually HAVE sub unit terms authored under them?
   *
   * ⚠ THE TIER'S NAME CANNOT BE VALIDATED AND THIS DOES NOT TRY TO. A per-unit tier's options are
   * the CHILDREN of the term above it, whatever those children happen to be called — the label only
   * names the tier and derives its column. So "does `Sub unit` exist in the term store" is not a
   * question with an answer. What IS answerable, and is the thing worth knowing, is whether any unit
   * term has children at all: if none does, the tier can never appear for anybody, which is the same
   * silent-skip shape that filed an HC document onto a unit folder on 2026-09-08.
   *
   * ⚠ PARTIAL IS THE DESIGNED SHAPE AND MUST NEVER BE REFUSED. Some units have sub units and some do
   * not, and they are unit-specific — that is exactly why they are authored under each unit rather
   * than in a flat set of their own, and why optionality needs no configuration. Only a definite
   * ZERO is refusable.
   *
   * ⚠ AND AN UNREACHABLE TERM STORE WARNS, NEVER BLOCKS. That is already this screen's rule for a
   * shared-list ID: malformed and 404 block Add, a resolvable-but-empty set and an unreachable store
   * warn only. An outage must not stop an admin authoring a legitimate tier.
   *
   * Bounded at `UNIT_WALK_CAP` requests, `UNIT_WALK_CONCURRENCY` at a time — GHO is 8 departments
   * and 62 units, so ~71 reads, in the same range as the abbreviation screen's own tree walk.
   * Hitting the cap is `unknown`, never "none": an incomplete walk that answered zero would refuse a
   * tier on a segment too large to check.
   */
  const UNIT_WALK_CAP = 400;
  const UNIT_WALK_CONCURRENCY = 8;

  const childIds = async (url: string): Promise<string[] | undefined> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      url, SPHttpClient.configurations.v1, { headers: { Accept: "application/json" } },
    );
    if (!res.ok) return undefined;
    const rows = (((await res.json()).value ?? []) as Array<{ id?: string }>);
    return rows.map((r) => r.id ?? "").filter((id) => id !== "");
  };

  const checkUnitTerms = async (guid: string, permissioned: number): Promise<UnitCheck> => {
    if (!guid || permissioned < 1) return { state: "unknown" };
    let spent = 0;
    const setBase = `${siteUrl}/_api/v2.1/termStore/sets/${guid}`;

    // Descend to the terms at the LAST permissioned tier — the Unit level. `permissioned` is the
    // count of those tiers, so the set's own children are level 1 and Unit is level `permissioned`.
    let level = await childIds(`${setBase}/children?$select=id`);
    spent++;
    if (level === undefined) return { state: "unknown" };

    for (let depth = 1; depth < permissioned; depth++) {
      const next: string[] = [];
      let failed = false;
      let i = 0;
      const parents = level;
      const workers = Array.from(
        { length: Math.min(UNIT_WALK_CONCURRENCY, parents.length) },
        async () => {
          while (i < parents.length) {
            const id = parents[i++];
            if (spent >= UNIT_WALK_CAP) { failed = true; return; }
            spent++;
            const kids = await childIds(`${setBase}/terms/${id}/children?$select=id`);
            // ⚠ ONE UNREADABLE BRANCH ABANDONS THE WHOLE ANSWER, deliberately. Counting it as "no
            // children" would report a SHALLOWER tree than exists — the same reasoning as the
            // segment depth check, where too-shallow is the silent direction.
            if (kids === undefined) { failed = true; return; }
            for (const k of kids) next.push(k);
          }
        },
      );
      await Promise.all(workers);
      if (failed) return { state: "unknown" };
      level = next;
    }

    if (level.length === 0) return { state: "unknown" }; // no unit terms at all: not this check's business
    const units = level;
    let withKids = 0;
    let failed = false;
    let i = 0;
    const workers = Array.from(
      { length: Math.min(UNIT_WALK_CONCURRENCY, units.length) },
      async () => {
        while (i < units.length) {
          const id = units[i++];
          if (spent >= UNIT_WALK_CAP) { failed = true; return; }
          spent++;
          const kids = await childIds(`${setBase}/terms/${id}/children?$select=id`);
          if (kids === undefined) { failed = true; return; }
          if (kids.length > 0) withKids++;
        }
      },
    );
    await Promise.all(workers);
    if (failed) return { state: "unknown" };
    return withKids === 0
      ? { state: "none", units: units.length }
      : { state: "some", withTerms: withKids, units: units.length };
  };

  /**
   * Whether ANY document is already filed under this segment's top folder.
   *
   * Existence, not a count — it is all the warning needs, and one CAML query with
   * RowLimit 1 answers it whatever the library size. A count would mean paging through
   * a library the Year × Document Type grid can push past the 5,000-item threshold.
   *
   * Files only (`FSObjType = 0`). Folders exist from the moment reconciliation runs, so
   * counting them would report every provisioned segment as "in use" before anyone had
   * uploaded anything.
   *
   * Checks the approval library too: a segment with nothing approved can still hold
   * pending files whose paths a structure change would leave behind.
   */
  const segmentHasDocuments = async (folder: string): Promise<boolean> => {
    if (!folder) return false;
    const webPath = new URL(siteUrl).pathname.replace(/\/$/, "");
    const libs: Array<[string, string]> = [
      [libraryTitle(), libraryUrlSegment()],
      [documentsListTitle(), DOCUMENTS_URL_SEGMENT],
    ];
    // The HC pair counts as "in use" too. A segment whose only documents are Highly Confidential
    // would otherwise read as empty and activate a structure change immediately — stranding exactly
    // the documents nobody can go and look at to notice.
    const hc = cachedHcLibraries();
    if (hc) {
      libs.push([hc.approval.title, hc.approval.urlSegment]);
      libs.push([hc.documents.title, hc.documents.urlSegment]);
    }
    for (const [title, segment] of libs) {
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/GetItems`,
        SPHttpClient.configurations.v1,
        {
          headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/json" },
          body: JSON.stringify({
            query: {
              ViewXml:
                `<View Scope="RecursiveAll"><Query><Where>` +
                `<Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq>` +
                `</Where></Query><RowLimit>1</RowLimit></View>`,
              FolderServerRelativeUrl: `${webPath}/${segment}/${folder}`,
            },
          }),
        },
      );
      if (!res.ok) continue;
      const data = await res.json();
      if (((data.value ?? []) as unknown[]).length > 0) return true;
    }
    return false;
  };

  useEffect(() => {
    let cancelled = false;
    // Names FIRST — every read below goes through cachedListTitle, and an unprimed cache
    // resolves to the legacy DMS titles, which 404 on a renamed site and would present
    // as "the config could not be read" rather than "the list is called something else".
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      .then(() => loadLegacyTermSets())
      .then((sets) => {
        if (!cancelled) setLegacySets(sets);
      })
      .catch(() => undefined)
      .then(() => loadSegments())
      .then(async (rows) => {
        if (cancelled) return;
        setSegments(rows);
        setLoading(false);
        // Per-segment document checks are slower than the list read, so fill them in
        // afterwards rather than holding the whole page on them.
        for (const row of rows) {
          let used = true;
          try {
            used = await segmentHasDocuments(row.stagingFolder);
          } catch {
            // Cannot prove it is empty, so assume in use. A needless warning costs one
            // click; a missing one costs a split folder tree.
            used = true;
          }
          if (cancelled) return;
          setSegments((prev) => prev.map((p) => (p.key === row.key ? { ...p, hasDocuments: used } : p)));
        }
      })
      .catch((e: Error) => {
        if (cancelled) return;
        setLoadError(e.message);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  // Check the term set ID as it is typed. Debounced, and `stale` discards a slow reply
  // that lands after the field has moved on — otherwise a verdict about an earlier,
  // half-typed GUID would sit under the field looking like a verdict about this one.
  //
  // ⚠ `recheck` IS A DEP SO THE BUTTON RE-RUNS THIS EFFECT, rather than calling `checkTermSet`
  // itself. A second caller would be a second definition of what a verdict means — the debounce,
  // the `stale` guard and the `unknown` fallback all live here — and the drifting copy would be the
  // rarely-pressed one. The 400ms debounce applies to a manual press too, which is harmless and
  // keeps `checking` on screen long enough for the click to register.
  //
  // It exists because the fix happens SOMEWHERE ELSE: the Term Store link opens a new tab, so an
  // admin edits the set there and comes back to a field that has not changed — and nothing else
  // would re-ask. Sixth instance of "a screen that reads once lies about anything changed beside
  // it", and the first where the change is deliberately made in another tab.
  const typedSet = adding === undefined ? "" : adding.termSetGuid;
  useEffect(() => {
    const guid = normalizeGuid(typedSet);
    if (!guid) {
      setSetCheck({ state: "blank" });
      return undefined;
    }
    if (!GUID_RE.test(guid)) {
      setSetCheck({ state: "malformed" });
      return undefined;
    }
    setSetCheck({ state: "checking" });
    let stale = false;
    const timer = setTimeout(() => {
      checkTermSet(guid)
        .then((r) => {
          if (!stale) setSetCheck(r);
        })
        .catch(() => {
          if (!stale) setSetCheck({ state: "unknown", status: 0 });
        });
    }, 400);
    return () => {
      stale = true;
      clearTimeout(timer);
    };
  }, [typedSet, recheck]);

  /**
   * The folder path a chain produces, with level names in brackets where a real value
   * goes. The only representation of the change an admin can reason about — which is why
   * it sits under every editor and the JSON sits nowhere.
   */
  const pathPreview = (seg: SegmentRow, chain: Level[]): string =>
    [seg.stagingFolder || seg.label, ...chain.map((l) => `[${l.label}]`)].join(" / ");

  /* ---------- Editing --------------------------------------------------- */

  /**
   * Open the editor on the levels the segment is EFFECTIVELY using, not just the ones
   * written down.
   *
   * A segment with nothing below Unit runs on the built-in Year → Document Type pair, and
   * showing an empty list would be a lie with teeth: adding one level would silently
   * replace that pair, dropping Year and Document Type from every future path. Seeding
   * the draft with them makes adding a level an insertion into a list the admin can see,
   * and saving writes all of it explicitly.
   *
   * The seeded pair comes from `effectiveOnDemandTiers`, so it carries the same column
   * names the form has always written and — importantly — NO `tidCol`, because Year and
   * Document Type are managed-metadata columns that neither have nor want a companion.
   */
  const startEdit = (seg: SegmentRow): void => {
    // Edit the STAGED chain when there is one, not the live one. Otherwise a second edit
    // before the migration silently discards the first, and the migration would then apply a
    // shape nobody had reviewed.
    const base = seg.pending ?? seg.chain;
    const { permissioned } = splitChain(base);
    const below = effectiveOnDemandTiers(base, legacySets.year, legacySets.docType);
    const seeded = [...permissioned, ...below].map((l) => ({ ...l }));
    setEditing(seg.key);
    setDraft(seeded);
    setBaseline(JSON.stringify(seeded));
    setAdding(undefined);
    setResult(undefined);
  };

  /**
   * Whether the editor holds changes that have not been saved.
   *
   * Compared as JSON against what was seeded, so reordering a level and putting it back reads
   * as clean — a warning that fires on a no-op change is one people learn to dismiss.
   *
   * Exists because a level added and then navigated away from is lost SILENTLY: the list looks
   * finished while it is being edited, so there is nothing on screen afterwards to say the work
   * did not land. Reported 2026-08-11 after it happened.
   */
  const dirty = editing !== undefined && JSON.stringify(draft) !== baseline;

  /**
   * A half-filled Add form that Save would throw away.
   *
   * ⚠ THE OTHER HALF OF THE 2026-09-09 REPORT. Save writes `draft`, and the Add form is a SEPARATE
   * piece of state that only joins `draft` when Add is pressed — so an admin who fills the form in,
   * misses that step and presses Save loses a complete level with nothing on screen saying so. The
   * screen looks saved, because it was: just not with the level they authored.
   *
   * ⚠ AN UNTOUCHED FORM BLOCKS NOTHING. Opening the form and thinking better of it leaves nothing
   * to lose, and refusing there would make Cancel a compulsory step on the way to every save — the
   * same reason `showErrors` never lights a blank form red.
   */
  const addPending =
    adding !== undefined &&
    (adding.label.trim() !== "" || normalizeGuid(adding.termSetGuid) !== "");

  const cancelEdit = (): void => {
    setEditing(undefined);
    setDraft([]);
    setBaseline("");
    setAdding(undefined);
    setConfirmDiscard(false);
  };

  /** Cancel asks first when there is something to lose, and only then. */
  const onCancelClicked = (): void => {
    if (dirty) setConfirmDiscard(true);
    else cancelEdit();
  };

  // The browser's own leave-the-page prompt. It is the only thing that can interrupt a closed
  // tab or a typed URL, which is exactly how the change was lost. Deliberately not attached
  // when clean: an unconditional prompt is noise, and noise gets clicked through.
  useEffect(() => {
    if (!dirty) return undefined;
    const warn = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      // Browsers ignore custom text and show their own wording; returnValue is still required
      // for the prompt to appear at all.
      e.returnValue = "You have unsaved folder structure changes.";
      return e.returnValue;
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);

  // Let the page block a tab switch while an edit is in flight.
  useEffect(() => {
    if (onDirtyChange) onDirtyChange(dirty);
  }, [dirty]);

  const permissionedCount = (chain: Level[]): number => splitChain(chain).permissioned.length;

  /**
   * What the DEEPEST permissioned tier is actually called, for the sub unit messages.
   *
   * ⚠ NEVER THE LITERAL "unit". The admin names the permissioned tiers at onboarding, so this is
   * `Estate/Mill` on Upstream Ops, `Buah Number` on Buah and `Department` on SDGI. Hardcoding
   * "unit" is the same mistake as the hardcoded Department/Unit rows in the approver's detail panel,
   * which read blank on every segment that names its tiers differently.
   */
  const unitLabel = (chain: Level[]): string => {
    const perm = splitChain(chain).permissioned;
    return perm.length > 0 ? perm[perm.length - 1].label : "unit";
  };


  /**
   * Run the sub unit walk when, and only when, a per-unit tier is actually being authored.
   *
   * ⚠ DRIVEN FROM AN EFFECT, NOT FROM THE RADIO'S onChange. There are two routes into the per-unit
   * state - opening the form (which defaults to it) and toggling back to it - and a handler on one
   * of them starts from zero on the other. That is the "a new route into an existing path silently
   * starts from zero" trap this project has now paid for three times, most recently on the two
   * clash-dialog failure paths. One effect covers every route, including a third added later.
   *
   * ⚠ ABOVE EVERY EARLY RETURN. This component returns early for `loadError` and again for the
   * editing branch, and a hook declared below either renders a different number of hooks on the
   * pass after loading finishes - "Rendered more hooks than during the previous render" - which
   * blanks the whole web part with no error UI. Four screens in this project have paid for that.
   *
   * Keyed on the segment and the toggle, never on the label or the position: retyping the tier's
   * name has no bearing on whether the terms exist, and re-walking 71 reads per keystroke would
   * make the field unusable.
   */
  useEffect(() => {
    const seg = editingSegment();
    if (!adding || !adding.fromUnit || !seg) {
      setUnitCheck({ state: "idle" });
      return;
    }
    if (!seg.termSetGuid) {
      // No term set recorded for this segment, so the walk has no root. Unknown, which warns.
      setUnitCheck({ state: "unknown" });
      return;
    }
    let stale = false;
    setUnitCheck({ state: "checking" });
    checkUnitTerms(seg.termSetGuid, permissionedCount(draft))
      .then((r) => { if (!stale) setUnitCheck(r); })
      .catch(() => { if (!stale) setUnitCheck({ state: "unknown" }); });
    return () => { stale = true; };
    /* `adding?.fromUnit` is `true | false | undefined`, and `undefined` uniquely means the form is
       closed - so it covers open/closed on its own and no second dep is needed for it.
       `draft` is deliberately ABSENT: the permissioned prefix renders locked on this screen, so the
       depth this walk descends to cannot change while the form is open, and depending on the draft
       would re-walk ~71 reads on every tier edit. */
  }, [editing, adding?.fromUnit]);


  /** Move a below-Unit level. Bounded to the below-Unit region — the prefix cannot move. */
  const moveTier = (index: number, delta: number): void => {
    /* ⚠ THE SAME CHECK THE BUTTON MAKES, and not merely belt-and-braces: a disabled attribute is a
       rendering, and this handler is the thing that actually mutates the chain. Guarding only the
       button would leave the one path that can produce the broken chain unguarded. */
    if (!canMoveBelowUnitTier(draft, index, delta)) return;
    const first = permissionedCount(draft);
    const to = index + delta;
    if (to < first || to >= draft.length) return;
    const next = draft.slice();
    const item = next[index];
    next.splice(index, 1);
    next.splice(to, 0, item);
    setDraft(next);
  };

  const removeTier = (index: number): void => setDraft(draft.filter((_l, i) => i !== index));

  /**
   * The type of an existing column, or `undefined` when there is none.
   *
   * ⚠ THROWS WHEN THE LIST COULD NOT BE READ, and the caller must not turn that into `undefined`.
   * "No such column" is safe; "we could not look" is not, and they are otherwise the same value —
   * the trap `existingColumnReason` documents.
   *
   * ⚠ `$filter`ed server-side, NEVER a `$top` over the whole field collection. A document library
   * carries hundreds of fields and a newly added one sorts LAST, so a capped read would report the
   * very column being asked about as absent. That has bitten five reads in this project.
   *
   * Asked of the approved-side library, which carries the fullest schema — every tier column plus
   * Remark, Keyword, the submission stamps and the rest. A collision on one of the OTHER libraries
   * and not this one is possible in principle and is not checked: it would cost a request per
   * library, and the schemas are created together by this same screen.
   */
  const readFieldType = async (internalName: string): Promise<string | undefined> => {
    const lib = documentsLibraryTitle();
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(lib)}')/fields` +
      `?$select=InternalName,TypeAsString&$filter=InternalName eq '${encodeURIComponent(internalName)}'`;
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const rows = ((await res.json()).value ?? []) as Array<{ TypeAsString?: string }>;
    return rows.length === 0 ? undefined : rows[0].TypeAsString ?? "";
  };

  const addTier = async (): Promise<void> => {
    if (!adding) return;
    const label = adding.label.trim();
    const col = columnNameFor(label);
    if (!label || !col) return;
    setAddError(undefined);
    // Re-adding `Year` or `Document Type` restores the BUILT-IN shape, never a derived one.
    // Those columns already exist as managed metadata, and a derived `tidCol` makes every
    // writer treat them as plain text — which fails the tagging call on every document.
    // See builtInTierFor.
    const builtIn = builtInTierFor(label, legacySets.year, legacySets.docType);

    /* ── Is this column safe to bind to? ──────────────────────────────────────
       `ensureTextColumn` SKIPS a column that already exists whatever its TYPE, so a level whose
       name derives an existing column binds to it silently — no error, and a whole segment stops
       tagging. Client, 2026-09-06: *"put an error? force them to create a new column with new name
       instead of binding to the same one if it already exist?"*

       ⚠ EXEMPT FOR A BUILT-IN TIER, and that exemption is required rather than convenient. `Year`
       and `Document Type` ARE taxonomy columns and MUST be reused — checking them would refuse the
       one case `builtInTierFor` exists to make safe.

       ⚠ FAILS CLOSED on an unreadable read, against this codebase's usual direction. Everywhere
       else a failed read costs a form for a minute; here it would restore exactly the silent bind
       this check exists to prevent. Adding a level is a deliberate admin action, so a retry costs
       seconds — and the message says to retry rather than leaving a dead button. */
    if (!builtIn) {
      const reserved = reservedColumnReason(col);
      if (reserved) {
        setAddError(reserved);
        return;
      }
      let existing: string | undefined;
      try {
        existing = await readFieldType(col);
      } catch (e) {
        setAddError(
          `Could not check whether a column called "${col}" already exists (${(e as Error).message}). ` +
            `Press Add again — the level was not added.`,
        );
        return;
      }
      const clash = existingColumnReason(col, existing);
      if (clash) {
        setAddError(clash);
        return;
      }
    }
    const tier: Level = builtIn ?? {
      label,
      column: col,
      labelCol: col,
      tidCol: `${col}Tid`,
      permissioned: false,
    };
    // Blank term set = values come from the level above (how each Unit gets its own
    // SubUnits). Presence of termSet is the discriminator; there is no extra flag.
    // Store the normalized GUID, never the raw paste: braces and stray whitespace survive
    // into the URL and 404 the term-store call at upload time, long after this screen.
    const set = normalizeGuid(adding.termSetGuid);
    if (set) tier.termSet = set;
    const next = draft.slice();
    next.splice(permissionedCount(draft) + adding.position, 0, tier);
    setDraft(next);
    setAdding(undefined);
  };

  /* ---------- Save ------------------------------------------------------ */

  /** Create a text column if absent. Idempotent; returns whether it created one. */
  /**
   * Delegates to the shared `ensureColumn`, which slice B (SegmentCreator) also uses.
   *
   * The internal-name trick and the `Options: 8` choice are both load-bearing and both invisible
   * when wrong, so they live in one place rather than in two copies that drift.
   */
  const ensureTextColumn = async (
    listTitle: string,
    internalName: string,
    displayName: string,
  ): Promise<boolean> =>
    ensureColumn(context.spHttpClient, siteUrl, listTitle, internalName, displayName);

  /**
   * Create the `PendingLevels` column on DMS Config if it is absent.
   *
   * A Note field, not Text: a chain of five levels with internal names and term-set GUIDs
   * passes 255 characters easily, and a silently TRUNCATED chain is the worst possible
   * outcome here — it would parse as valid JSON only by luck, and if it did parse it would
   * describe a shorter structure than the admin authored.
   *
   * Created here rather than documented as a provisioning step because a missing column would
   * fail the save on a site that is otherwise correctly set up, and the person hitting it
   * would have no way to know why.
   */
  const ensurePendingColumn = async (): Promise<void> => {
    await ensureColumn(
      context.spHttpClient,
      siteUrl,
      cachedListTitle(LIST_SUFFIX.config),
      PENDING_LEVELS_FIELD,
      PENDING_LEVELS_FIELD,
      "Note",
    );
  };

  const saveStructure = async (): Promise<void> => {
    const seg = editingSegment();
    if (!seg) return;
    setBusy(true);
    setResult(undefined);
    try {
      // 1. VALIDATE FIRST, before a single column is created. A rejected chain must
      //    leave nothing behind.
      const err = validateChain(draft);
      if (err) throw new Error(err.message);

      // 2. Columns, in BOTH libraries. Order matters: a column with no chain is a
      //    harmless orphan, whereas a chain naming a missing column breaks every upload
      //    in the segment — and validateUpdateListItem returns HTTP 200 with
      //    HasException, so that failure is not even loud.
      /* ⚠ ARRAYS, NOT `Set`. Spreading a Set is TS2802 at this tsconfig target — the same ES-level
         limitation that makes `Promise.allSettled` unavailable in this project. An `indexOf` guard
         dedupes without needing the iteration protocol. */
      const createdCols: string[] = [];
      const createdLibs: string[] = [];
      for (const tier of splitChain(draft).onDemand) {
        const cols: Array<[string, string]> = [];
        // ensureTextColumn skips a column that already exists, whatever its type — so the
        // seeded Year / Document Type pair touches nothing, and only genuinely new levels
        // get columns created. A tier with no tidCol (managed metadata) never gets one
        // invented for it here.
        if (tier.labelCol) cols.push([tier.labelCol, tier.label]);
        if (tier.tidCol) cols.push([tier.tidCol, `${tier.label} ID`]);
        for (const [internal, display] of cols) {
          // Four libraries once the site has HC — see allLibraryTitles.
          for (const lib of allLibraryTitles()) {
            if (await ensureTextColumn(lib, internal, display)) {
            if (createdCols.indexOf(internal) < 0) createdCols.push(internal);
            if (createdLibs.indexOf(lib) < 0) createdLibs.push(lib);
          }
          }
        }
      }

      // 3. Write the chain — to `Levels` on an EMPTY segment, to `PendingLevels` on one that
      //    already holds documents.
      //
      //    A segment in use must not adopt a new shape the moment it is authored: its existing
      //    folders are still a tier above, so uploads would start landing in the new shape
      //    beside old folders in the old one. The people who meet that state first are
      //    uploaders nobody told, and it reads as a broken system rather than an unfinished
      //    admin task. So the change is STAGED, and the Move-existing-folders tab applies it
      //    as the last step of the migration.
      //
      //    An empty segment skips all of that: there is nothing to move, and making someone
      //    run a migration over zero folders teaches them to click through it.
      /* ⚠ STAGE ONLY A CHANGE THAT ACTUALLY CHANGES THE FOLDER SHAPE (client, 2026-09-09: *"I
         can click Save structure before I add and this causes the entire MHO to be change pending
         even there is nothing change in the structure"*). It used to write `PendingLevels` on any
         save of an in-use segment, so a save with the Add form still open — nothing added —
         stamped CHANGE PENDING on a segment nothing had moved in.
         That badge is not cosmetic: it holds the migrate step, invites a scan that finds nothing to
         move, and the only way back out is pressing Apply on a change that was never made.
         ⚠ COMPARED AGAINST THE LIVE CHAIN **NORMALISED THE WAY `startEdit` NORMALISES IT**, never
         against `seg.chain` raw. A segment with nothing below Unit runs on the IMPLICIT Year →
         Document Type pair and the editor seeds them explicitly, so a raw compare would call every
         such segment changed and reproduce the bug for exactly the segments that never touched their
         own structure. */
      const liveSeeded = [
        ...splitChain(seg.chain).permissioned,
        ...effectiveOnDemandTiers(seg.chain, legacySets.year, legacySets.docType),
      ];
      const sameAsLive = JSON.stringify(draft) === JSON.stringify(liveSeeded);
      /* Writing the identical shape to `Levels` is safe on an in-use segment precisely BECAUSE it is
         identical: staging exists to stop uploads landing in a new shape beside folders in the old
         one, and there is no new shape here. It also makes an implicit pair explicit, which is what
         the editor showed, and it CLEARS a pending chain — so reverting a staged change back to the
         live shape and saving now takes the badge off, which nothing else could do. */
      const staged = seg.hasDocuments === true && !sameAsLive;
      if (staged) await ensurePendingColumn();
      const body: Record<string, string> = staged
        ? { [PENDING_LEVELS_FIELD]: JSON.stringify(draft) }
        // Clearing pending on an empty segment matters: one could have been staged and then
        // emptied, and a stale pending chain would offer to apply a change already applied.
        : { Levels: JSON.stringify(draft), [PENDING_LEVELS_FIELD]: "" };
      const write: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cachedListTitle(LIST_SUFFIX.config))}')/items(${seg.id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify(body),
        },
      );
      if (!write.ok) {
        throw new Error(
          `The columns were created but the structure could not be saved (HTTP ${write.status}). ` +
            `Those columns are harmless and will be reused when you try again.`,
        );
      }

      setSegments((prev) =>
        prev.map((p) =>
          p.key === seg.key
            ? staged
              ? { ...p, pending: draft.map((l) => ({ ...l })) }
              : { ...p, chain: draft.map((l) => ({ ...l })), pending: undefined }
            : p,
        ),
      );
      /* SHORT, AND THE FIRST BULLET IS THE ONE THAT MATTERS. An admin who reads nothing else must
         still learn whether this is live. Everything explanatory moved below it, and the twelve
         created columns collapsed to what they are and where they went — the LIST was the bulk of
         the old message and answered a question nobody asks. */
      const lines: string[] = [];
      if (staged) {
        lines.push("Not live yet. Uploads carry on in the old shape.");
        lines.push('Next: step 3, "Move existing folders". The new shape goes live at the end of it.');
      } else if (sameAsLive) {
        lines.push("Nothing changed in the folder shape, so there is nothing to migrate.");
      } else {
        lines.push("Live now. Uploads use the new shape once people reload the form.");
      }
      if (createdCols.length > 0) {
        const libs = createdLibs.length;
        lines.push(
          `Added ${createdCols.join(" and ")} to ${libs} ${libs === 1 ? "library" : "libraries"}.`,
        );
        // Kept, shortened: an invisible new column looks exactly like the save having failed, so
        // the deliberate part has to be said somewhere.
        lines.push('Hidden in library views on purpose — add them with "Show or hide columns".');
      }
      setResult({ ok: true, text: `Saved for ${seg.label}.`, lines });

      // STAGED vs LIVE is the load-bearing part of this record. A row saying only "structure
      // changed" would imply uploads had moved to the new shape when, for a segment in use, they
      // deliberately have not — and that misreading is the very thing the staging exists to prevent.
      writeAudit(context.spHttpClient, siteUrl, {
        event: EVENT.structureChanged,
        source: "StructureManager",
        at: new Date(),
        actorName: context.pageContext.user.displayName,
        actorEmail: context.pageContext.user.email,
        segment: seg.label,
        summary: staged
          ? `Folder structure STAGED for ${seg.label} — not live yet`
          : sameAsLive
          ? `Folder structure re-saved for ${seg.label} — no change to the folder shape`
          : `Folder structure changed for ${seg.label} — live now`,
        details: [
          staged
            ? "Written to PendingLevels. Levels is untouched, so uploads carry on in the old shape."
            : "Written to Levels. Uploads use the new shape as soon as people reload the form.",
          `Old chain: ${(seg.pending ?? seg.chain).map((l) => l.label).join(" → ")}`,
          `New chain: ${draft.map((l) => l.label).join(" → ")}`,
          createdCols.length > 0
            ? `Columns created: ${createdCols.join(", ")} in ${createdLibs.join(", ")}`
            : "No columns created — every one it needs already existed.",
          staged
            ? "Goes live at the end of the folder migration, not before."
            : sameAsLive
            // ⚠ NOT "the segment held no documents" — an in-use segment reaches this branch too
            // when the shape did not move, and claiming it was empty would misrepresent the record.
            ? "The chain matches what was already live, so nothing moved and no pending change was staged."
            : "The segment held no documents, so there was nothing to migrate.",
        ],
      }).catch(() => undefined);

      cancelEdit();
      /* ⚠ AFTER `cancelEdit`, and only on the success path. The audit write above is deliberately
         `.catch`ed to undefined, so reaching here means the chain really was written — and a caller
         re-reading on a FAILED save would refresh itself into the same state while implying
         something changed. */
      if (onSaved) onSaved();
    } catch (e) {
      setResult({ ok: false, text: (e as Error).message });
    } finally {
      setBusy(false);
    }
  };

  /**
   * No confirmation gate here any more.
   *
   * Saving used to go live immediately, so a segment in use got a typed-CHANGE modal warning
   * about two folder shapes side by side. Staging removed the thing it was warning about:
   * nothing moves, uploads carry on unchanged, and the chain can be edited again before it is
   * applied. The gate that matters is now the typed MOVE on the migration, which is the step
   * that actually relocates documents — and a scary modal on an inert action is exactly what
   * teaches someone to click through the one that counts.
   */
  const onSaveClicked = (): void => {
    if (!editingSegment()) return;
    // REFUSED, not confirmed — the same rule as a tab switch with unsaved abbreviations. Both ways
    // out (Add, Cancel) are a few pixels above this button, and a "discard?" prompt would put losing
    // an authored level one click behind an ordinary-looking Save.
    if (addPending) return;
    saveStructure().catch(() => undefined);
  };

  /**
   * The save/failure box. ONE definition — it renders at two places (inside the editor and under
   * the segment list), and two copies would drift the moment either grew a bullet.
   */
  const resultBox = (r: { text: string; ok: boolean; lines?: string[] }): React.ReactElement => (
    <div style={{ ...s.msg, ...(r.ok ? s.ok : s.err) }}>
      <div>{r.text}</div>
      {r.lines !== undefined && r.lines.length > 0 && (
        <ul style={s.resultList}>
          {r.lines.map((l) => (
            <li key={l} style={s.resultItem}>{l}</li>
          ))}
        </ul>
      )}
    </div>
  );

  /* ---------- Render ---------------------------------------------------- */

  if (loading) return <p style={{ fontSize: 13, color: "#605e5c" }}>Loading folder structure&hellip;</p>;

  if (loadError !== undefined) {
    return (
      <div style={{ ...s.msg, ...s.err }}>
        Could not read the configuration: {loadError}. Nothing is safe to edit without it — this
        screen will not guess at a structure.
      </div>
    );
  }

  const seg = editingSegment();

  if (seg) {
    const first = permissionedCount(draft);
    const onDemand = draft.slice(first);
    const perUnitTaken = perUnitTierExists(onDemand);
    return (
      <>
        {result !== undefined && resultBox(result)}

        <div style={s.card}>
          <div style={s.segName}>{seg.label}</div>
          <div style={s.path}>{pathPreview(seg, draft)}</div>
        </div>

        {/* ⚠ THE EXPLANATION IS GONE, THE RULE IS NOT (client, 2026-09-06: *"Remove ALL COMPONENTS /
            TEXTS with red slash"*). These levels still carry the folder permissions and still cannot
            be edited here — the padlock beside each one is now the only thing that says so. */}
        <p style={s.label}>Fixed levels</p>
        {draft.slice(0, first).map((l, i) => (
          <div style={{ ...s.tierRow, ...s.tierLock }} key={`p${i}`}>
            {/* Drawn, not the padlock EMOJI it replaced: an emoji renders in the platform's own
                colour and style, which beside flat green line-art reads as a stray character. */}
            <span style={{ flex: "0 0 16px", color: "#8a8886", display: "flex" }} aria-hidden="true">
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                <rect x="3.5" y="7" width="9" height="6.5" rx="1.5" stroke="currentColor" strokeWidth="1.4" />
                <path d="M5.8 7V5.2a2.2 2.2 0 0 1 4.4 0V7" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
              </svg>
            </span>
            <span style={s.tierName}>{l.label}</span>
            <span style={s.tierMeta}>
              column: {l.labelCol ?? l.column}
              {l.tidCol !== undefined ? ` + ${l.tidCol}` : ""}
            </span>
          </div>
        ))}

        {/* Same removal as "Fixed levels" above. The inheritance is unchanged: everything below Unit
            takes the unit folder's ACL, which is why these need no abbreviation, group or mapping. */}
        <p style={s.label}>Folders below Unit</p>
        {onDemand.length === 0 && (
          <p style={{ fontSize: 12, color: "#8a8886", margin: "0 0 8px", lineHeight: 1.5 }}>
            None — uploads would stop at the Unit folder. Add at least one level.
          </p>
        )}
        {onDemand.map((l, k) => {
          const i = first + k;
          /* ⚠ YEAR AND DOCUMENT TYPE ARE FIXED (client, 2026-09-06: *"enforce Year and DocType to not
             move or be deleted"*, confirmed 2026-09-07). They are managed-metadata columns bound to
             SITE-WIDE term sets and every segment shares them, so removing one takes a tier out of
             the path for a column that still exists and still holds values — and re-adding it by hand
             is what produced the 1.0.195.0 corruption (an ordinary `tidCol` derived onto a taxonomy
             column, after which every document in the segment failed to tag).

             ⚠ ONLY THEIR OWN CONTROLS ARE DISABLED — every OTHER tier keeps ↑ / ↓ / Remove, which is
             what makes the client's actual requirement work: a new level can be placed **before,
             between, or after** the pair by adding it and moving it into position. Moving a custom
             tier past Year shifts Year's index, and that is fine — the two can never change order
             relative to EACH OTHER, because neither can be moved directly. */
          const fixed = isFixedBelowUnitTier(l);
          /* ⚠ ALSO REFUSED WHEN THE MOVE WOULD BREAK THE CHAIN, not only at the ends. A shared-list
             tier moved UP above a per-unit one is the `per-unit-not-contiguous` shape that made every
             unit in Buah unreadable — the save refuses it, but only after the move has visibly
             happened. The mirror is a per-unit tier moved DOWN past a shared-list one. */
          const canUp = !fixed && k > 0 && canMoveBelowUnitTier(draft, i, -1);
          const canDown =
            !fixed && k < onDemand.length - 1 && canMoveBelowUnitTier(draft, i, 1);
          return (
            <div style={s.tierRow} key={`d${i}`}>
              <span style={s.tierName}>{l.label}</span>
              <span style={s.tierMeta}>
                {l.termSet !== undefined ? "own list of values" : "values come from the level above"}
                {" · column: "}
                {l.labelCol ?? l.column}
                {/* ⚠ DERIVED, NOT A SETTING (client, 2026-09-09). Every below-Unit level but the
                    fixed pair is named by its terms' abbreviations, so this states a fact rather than
                    reflecting a choice — there is no longer a control to reflect. */}
                {!fixed && " · folders named by abbreviation"}
                {fixed && " · fixed — cannot be moved or removed"}
              </span>
              <span style={{ flex: "0 0 auto" }}>
                {/* ⚠ ABSENT, NOT GREYED, WHEN A LEVEL CAN MOVE NEITHER WAY (client, 2026-09-09) —
                    the rule Remove below already follows. Three levels reach it for two different
                    reasons: Year and Document Type are `fixed`, a SUB UNIT is refused both ways by
                    `canMoveBelowUnitTier` because it may only sit directly under Unit. Derived from
                    the predicates the buttons already obeyed, so it names none of them and covers a
                    future immovable level too.

                    ⚠ ONE DIRECTION STILL SHOWS BOTH: a greyed ↑ beside a usable ↓ says "you are at
                    the end" and becomes usable when another level is added. Only "neither" hides. */}
                {(canUp || canDown) && (
                  <>
                    <button
                      style={s.iconBtn}
                      disabled={!canUp}
                      onClick={() => moveTier(i, -1)}
                      title="Move up"
                    >
                      &#8593;
                    </button>
                    <button
                      style={s.iconBtn}
                      disabled={!canDown}
                      onClick={() => moveTier(i, 1)}
                      title="Move down"
                    >
                      &#8595;
                    </button>
                  </>
                )}
                {/* Absent rather than greyed: a disabled Remove on a level that can NEVER be removed
                    is a control that will never do anything, and the meta line above already says
                    why it is not there. Greying it invites repeated clicking. */}
                {!fixed && (
                  <button style={s.danger} onClick={() => removeTier(i)}>Remove</button>
                )}
              </span>
            </div>
          );
        })}

        {adding === undefined ? (
          <button
            style={s.ghost}
            // Defaults to per-unit, because SubUnit is the tier the client is actually adding and
            // the shared-list kind (Year, Document Type) already exists on every segment.
            onClick={() =>
              setAdding({
                label: "", fromUnit: true, termSetGuid: "",
                position: clampTierPosition(onDemand, true, onDemand.length),
              })
            }
          >
            + Add folder level
          </button>
        ) : (
          <div style={{ ...s.card, marginTop: 12 }}>
            {/* ⚠ THE CHOICE COMES FIRST NOW (client's design, 2026-09-06), and that ordering is
                better than it looks: it decides whether a Term set ID field appears at all, so
                asking for the name first meant the form changed shape UNDER the admin after they had
                already filled something in.

                ⚠ THE LABELS CHANGED, THE MEANING DID NOT. "Sub unit" is still the tier whose values
                are terms authored UNDER each unit, and "Shared folder term" is still the one that
                needs a term set. The stored discriminator is unchanged: a below-Unit tier with no
                `termSet` is per-unit. Rename these two if you like; do not let them come to mean
                anything else. */}
            <span style={s.label}>Folder setup</span>
            {/* An EXPLICIT choice, not a blank field (client, 2026-08-17: "can you add a button call
                have subunit?"). The stored discriminator is the absence of `termSet`, so this used to
                be expressed by leaving a text box empty — undiscoverable, and the wrong guess is
                silent: paste any other set's ID and every unit is offered every other unit's
                subunits, with nothing on screen to say so. */}
            <div style={{ display: "grid", gap: 8, marginBottom: 14 }}>
              <label style={s.radioRow} htmlFor="sm-src-unit">
                <input
                  id="sm-src-unit"
                  type="radio"
                  name="sm-src"
                  checked={adding.fromUnit}
                  // Clearing the GUID is deliberate, not tidiness: a half-typed ID that survived the
                  // toggle would be saved the moment anyone switched back, producing a shared-list
                  // tier from a screen the admin last saw set to per-unit.
                  onChange={() => setAdding({
                    ...adding, fromUnit: true, termSetGuid: "",
                    position: clampTierPosition(onDemand, true, adding.position),
                  })}
                />
                <span>
                  <strong>Sub unit</strong>
                  <br />
                  <span style={s.radioHint}>
                    Create sub unit folder under each unit.
                  </span>
                </span>
              </label>
              <label style={s.radioRow} htmlFor="sm-src-set">
                <input
                  id="sm-src-set"
                  type="radio"
                  name="sm-src"
                  checked={!adding.fromUnit}
                  onChange={() => setAdding({
                    ...adding, fromUnit: false,
                    position: clampTierPosition(onDemand, false, adding.position),
                  })}
                />
                <span>
                  {/* ⚠ RENAMED ON THE CLIENT'S INSTRUCTION (2026-09-09). Note that BOTH radios add a
                      folder layer, so the name does not by itself say which is which — the hint under
                      it is doing that work. The distinction they answer is where the VALUES come from:
                      each unit's own child terms, or one list shared by every unit. Earlier labels
                      were "One shared list" (2026-08-17) and then "Shared folder term". */}
                  <strong>New Folder Layer</strong>
                  <br />
                  <span style={s.radioHint}>
                    One list of options, shared by every unit folder.
                  </span>
                </span>
              </label>
            </div>

            {/* ⚠ NO LONGER A CHOICE (client, 2026-09-09: *"honestly just enforce the term abbreviation
                to be created"*, then *"can we enforce them to use term abbreviations when adding new
                subunit or share folders"*). A new below-Unit level is ALWAYS named by abbreviation,
                so there is nothing to tick and nothing to forget.

                ⚠ THE CHECKBOX THAT WAS HERE ALSO READ AS A DUPLICATE OF THE ONE ON THE TIER ROW
                ABOVE — client: *"I can tick both btw"*. They govern different levels (that row's
                level, and the one being added), which nothing on screen said.

                ⚠ NOT APPLIED TO YEAR OR DOCUMENT TYPE — see `addTier`. There is no useful
                abbreviation for `2024`, and coding either would move every document in the system.

                ⚠ THE TIER-ROW CHECKBOX STAYS, and only until the four levels that predate this are
                switched on and migrated (`Minasmas Archive 2`, GHO's and Buah's `Shared Folder`,
                TO's `SDG`). It is the migration path, not an option — delete it once they are done
                and the rule becomes absolute with nothing left to get wrong. */}
            <label style={{ ...s.label, marginTop: 14 }}>Folder naming</label>
            <p style={{ ...s.hint, marginTop: 0 }}>
              Folders on this level are named by each term&rsquo;s <strong>abbreviation</strong>
              {" "}rather than its full label, which keeps paths short. Every term on it needs an
              abbreviation before anyone can file into it.
            </p>

            {/* ⚠ HIDDEN WHEN A SUB UNIT LEVEL ALREADY EXISTS (client, 2026-09-09). Naming a level
                that cannot be added is work thrown away — typed, then a greyed Add, then the reason.
                ⚠ GATED ON THE COMBINATION, NEVER `perUnitTaken` ALONE: switching to Shared folder
                term is the way out, so the field must come straight back when they do. */}
            {!(adding.fromUnit && perUnitTaken) && (
              <>
            <span style={s.labelRow}>
              <label style={{ ...s.label, marginTop: 0 }} htmlFor="sm-label">Folder level name</label>
              {/* ⚠ THE NEUTRAL INFO AFFORDANCE, NOT A RED WARNING — matching the upload form's own
                  icon, which is what the client meant by "like how upload form works". A red mark
                  sitting beside a label on a form nobody has filled in yet reads as an error state,
                  and this project's own rule is that colouring an untouched form red is how people
                  stop reading red. The CAUTION is in the panel; the icon only says there is
                  something here to read.

                  ⚠ Form.tsx draws the same affordance from a base64 PNG behind CSS `:hover`. This
                  file has no <style> block, so hover is React state instead — and adding one just
                  for this would walk into the backtick-in-a-template-literal trap that has broken
                  the build five times. Unifying the two icons is worth its own change. */}
              <span
                style={s.infoWrap}
                onMouseEnter={() => setNameInfo(true)}
                onMouseLeave={() => setNameInfo(false)}
              >
                <button
                  type="button"
                  style={s.infoBtn}
                  aria-expanded={nameInfo}
                  aria-label="Why the folder level name is permanent"
                  onClick={() => setNameInfo(!nameInfo)}
                  onFocus={() => setNameInfo(true)}
                  onBlur={() => setNameInfo(false)}
                >
                  <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
                    <circle cx="8" cy="8" r="7" fill="none" stroke="currentColor" strokeWidth="1.3" />
                    <circle cx="8" cy="4.6" r="0.95" fill="currentColor" />
                    <path d="M8 7v5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
                  </svg>
                </button>
                {nameInfo && (
                  <span style={s.infoPanel} role="tooltip">
                    <strong>This name is permanent in practice.</strong> It becomes the level&rsquo;s
                    label on the upload form and in the folder chain, and it creates the columns{" "}
                    {/* ⚠ THE DERIVED COLUMN NAME LIVES HERE, NOT UNDER THE FIELD. A standing preview
                        was removed on the client's request (2026-09-06) and re-adding one would
                        reverse that; behind the icon it answers "what is it tied to" — which is what
                        they asked for — and it is also the one thing that makes a typo visible,
                        since the name is what the column is called for ever. */}
                    {columnNameFor(adding.label) ? (
                      <>
                        <code>{columnNameFor(adding.label)}</code> and{" "}
                        <code>{columnNameFor(adding.label)}Tid</code>
                      </>
                    ) : (
                      <>it derives</>
                    )}{" "}
                    in all four libraries. A column&rsquo;s internal name is fixed when it is created
                    and can never be changed, and removing the level later deletes neither the column
                    nor the values already written to it. Type the name you want to keep.
                  </span>
                )}
              </span>
            </span>
            <input
              id="sm-label"
              style={s.input}
              value={adding.label}
              onChange={(e) => setAdding({ ...adding, label: e.target.value })}
              placeholder="Folder Name"
            />
              </>
            )}
            {/* ⚠ THE COLUMN NAME IS NO LONGER SHOWN (client, 2026-09-06), AND IT WAS NOT DECORATION.
                `ensureTextColumn` SKIPS a column that already exists whatever its TYPE, so a level
                named into an existing column binds to it silently — that is how a re-added `Year`
                once wrote a bare label into a taxonomy field and every upload in the segment stopped
                tagging. `builtInTierFor` guards the two known cases (Year, Document Type) by name;
                nothing guards a collision with a column the client added themselves, and the derived
                name is no longer on screen for anyone to notice it.
                `columnNameFor` is still what CREATES the column — only the preview is gone. */}
            {!adding.fromUnit && (
              <>
                <span style={s.labelRow}>
                  <label style={{ ...s.label, marginTop: 0 }} htmlFor="sm-set">Term set ID</label>
                  {/* ⚠ THE CLASSIC, SITE-LEVEL PAGE — fourth mount point, same URL, same reason. The
                      MODERN term store (`/_layouts/15/SiteAdmin.aspx#/termStoreAdminCenter`) is the
                      TENANT admin centre and answers "Access denied" to a site collection
                      administrator; the client hit that wall themselves on 2026-08-30. Built from
                      `siteUrl`, never hardcoded.

                      New tab, which matters here: this screen holds an unsaved-changes guard, so a
                      same-tab jump would either prompt or lose the half-filled Add form — and the
                      admin is going there precisely to fetch a value to paste back into it. */}
                  <a
                    style={s.labelLink}
                    href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    Open the Term Store
                  </a>
                  {/* Offered only where the verdict came from the server — see
                      `setCheckIsRecheckable`. Disabled WHILE checking rather than hidden: hiding it
                      moves the Term Store link as the row reflows, so the control an admin is
                      reaching for jumps out from under the cursor. */}
                  {setCheckIsRecheckable(setCheck) && (
                    <button
                      type="button"
                      style={setCheck.state === "checking" ? s.recheckOff : s.recheckBtn}
                      disabled={setCheck.state === "checking"}
                      title="Check this term set again — use it after editing the set in the Term Store"
                      onClick={() => setRecheck((n) => n + 1)}
                    >
                      {setCheck.state === "checking" ? "…" : "↻ Re-check"}
                    </button>
                  )}
                </span>
                <input
                  id="sm-set"
                  style={s.input}
                  value={adding.termSetGuid}
                  onChange={(e) => setAdding({ ...adding, termSetGuid: e.target.value })}
                  placeholder="023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf"
                />
                {setCheck.state !== "blank" && (
                  <p style={{ ...s.hint, marginTop: 6, fontSize: 12, ...setCheckStyle(setCheck) }}>
                    {setCheckMessage(setCheck)}
                  </p>
                )}
              </>
            )}

            {/* ⚠ SHOWN ONLY FOR A PER-UNIT TIER. For a shared list the walk never runs and the
                state is `idle`, so there is nothing to say — and a stale count from a previous
                toggle must not sit under a field it does not describe. `some` renders in the
                ordinary hint tone, not the warning one: a partial spread is the designed shape,
                and colouring it amber would train people to ignore the state that matters. */}
            {/* ⚠ SHOWN INSTEAD OF THE TERM COUNT, not beside it. That count walks to the UNIT level,
                which is the right level for the FIRST per-unit tier and the wrong one for a second —
                so leaving it on screen here would put a true-but-irrelevant number under a control
                that cannot be used. Name the refusal and point at the alternative. */}
            {adding.fromUnit && perUnitTaken && (
              <p style={{ ...s.msg, ...s.warn }}>
                This segment already has a sub unit level (
                {onDemand.filter((l) => !(l.termSet ?? "").trim())[0]?.label}), continue to add more
                in the{" "}
                {/* ⚠ THE CLASSIC, SITE-LEVEL PAGE — fifth mount point, same URL, same reason. The
                    MODERN term store is the TENANT admin centre and answers "Access denied" to a
                    site collection administrator; the client hit that wall on 2026-08-30. New tab,
                    because this screen holds an unsaved-changes guard. */}
                <a
                  href={`${siteUrl}/_layouts/15/termstoremanager.aspx`}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Term Store
                </a>
                .
              </p>
            )}
            {adding.fromUnit && !perUnitTaken && unitCheck.state !== "idle" && (
              <p style={unitCheck.state === "none" ? { ...s.msg, ...s.warn } : s.hint}>
                {unitCheckNote(unitCheck, unitLabel(draft), seg.label, siteUrl)}
              </p>
            )}

            {/* ⚠ THE LABEL FOR ONE SLOT, DERIVED ONCE AND USED BY BOTH BRANCHES BELOW. Two copies
                of this expression would be two answers to "what is this position called", and the
                one that drifts is the branch nobody looks at. */}
            {(() => {
              const slots = allowedTierPositions(onDemand, adding.fromUnit);
              const name = (p: number): string =>
                onDemand.length === 0
                  ? "First level below Unit"
                  : p === 0
                    ? `Before ${onDemand[0].label}`
                    : p === onDemand.length
                      ? `After ${onDemand[onDemand.length - 1].label}`
                      : `Between ${onDemand[p - 1].label} and ${onDemand[p].label}`;

              /* ⚠ ONE SLOT ⇒ NO CONTROL (client, 2026-09-08: *"what is the point of the Dropdown
                 being able to select, just disable it since it will always be under Unit"*).

                 Stated as text rather than DISABLED, which is the treatment this project already
                 settled on for the migrate screen's segment picker — client, 2026-09-06: *"You
                 haven't remove the select... just put a message indicator is enough."* A greyed
                 control still reads as something to interact with, and its reason only appears on
                 hover.

                 ⚠ BUT NOT HARDCODED TO "under Unit", because it is NOT always one slot: once a
                 per-unit tier exists, a SECOND one may legitimately go above or below it, and a
                 shared-list tier gains every slot from the end of the per-unit run downwards. The
                 control disappears when there is nothing to decide and comes back when there is. */
              if (slots.length === 1) {
                return (
                  <>
                    <label style={s.label}>Position</label>
                    <p style={{ ...s.hint, marginTop: 0 }}>
                      {/* ⚠ SLOT 0 IS NAMED FOR ITS ANCHOR, NOT ITS NEIGHBOUR (client, 2026-09-08:
                          *"it says before Year but another Shared Folder Term is before year as
                          well"*). "Before Year" is true only because Year happens to be first
                          today; the actual rule is "first level below Unit", so the moment anything
                          is inserted the sentence stops describing the rule while staying
                          technically true. The anchor cannot move — the contiguity rule closes slot
                          0 to every shared-list tier — so this stays correct for the life of the
                          chain.

                          The comparative labels stay in the DROPDOWN branch, where "Before Year /
                          Between Year and Document Type / After Document Type" is one ordered list
                          and consistency across its options is what makes it readable. */}
                      <strong>{slots[0] === 0 ? `Directly under ${unitLabel(draft)}` : name(slots[0])}</strong>
                      {" — "}
                      {adding.fromUnit
                        ? "a sub unit level takes its values from the terms under each unit, so this " +
                          "is the only place it works."
                        : "a shared level cannot sit above a sub unit level, which cascades from the " +
                          "Unit term."}
                    </p>
                  </>
                );
              }
              return (
                <>
                  <label style={s.label} htmlFor="sm-pos">Position</label>
                  <select
                    id="sm-pos"
                    style={s.input}
                    value={String(adding.position)}
                    onChange={(e) => setAdding({ ...adding, position: Number(e.target.value) })}
                  >
                    {slots.map((p) => (
                      <option key={p} value={String(p)}>{name(p)}</option>
                    ))}
                  </select>
                </>
              );
            })()}
            {/* ⚠ THE COST OF POSITION IS NO LONGER STATED (client, 2026-09-06) AND IT IS REAL.
                Near the top means one folder per unit; at the bottom means one for every Year and
                Document Type combination — the same level placed last can multiply the folder count
                by the size of two term sets. Nothing enforces a limit, so a badly placed level is
                discovered as a slow reconciliation rather than as an error. */}

            <div style={{ marginTop: 14 }}>
              <button
                style={canAddTier(adding, setCheck, unitCheck, perUnitTaken) ? s.btn : s.off}
                disabled={!canAddTier(adding, setCheck, unitCheck, perUnitTaken)}
                // `addTier` is async now — it asks the site whether the derived column already
                // exists. The rejection is swallowed here because every failure path inside it
                // already sets `addError`; an unhandled rejection would show nothing at all.
                onClick={() => { addTier().catch(() => undefined); }}
              >
                Add
              </button>{" "}
              <button
                style={s.ghost}
                // Clearing the refusal matters: it names a column, and leaving it on screen after
                // the form is dismissed would describe a level that no longer exists.
                onClick={() => { setAdding(undefined); setAddError(undefined); }}
              >
                Cancel
              </button>
            </div>
            {/* BESIDE THE BUTTON, not as a toast. It names the column and what already owns it, and
                the fix is to edit the field a few pixels above — a message that fades takes the name
                with it. */}
            {addError !== undefined && (
              <div style={{ ...s.msg, ...s.err, marginTop: 12, marginBottom: 0 }}>
                {addError}
              </div>
            )}
          </div>
        )}

        <div style={{ marginTop: 24, borderTop: "1px solid #edebe9", paddingTop: 16 }}>
          {/* ⚠ NO REASON IS RENDERED BESIDE THIS WHEN `addPending` GREYS IT (client, 2026-09-09:
              *"its plain obvious they need to click add to save, so remove it"*), which is a
              deliberate departure from the house rule that an unexplained disabled button reads as a
              broken page. The judgement is theirs: the Add form is directly above with its own Add
              and Cancel, so the way out is on screen even though nothing points at it. Restore a
              message here if anyone ever reports Save as broken. */}
          <button
            style={busy || addPending ? s.off : s.btn}
            disabled={busy || addPending}
            onClick={onSaveClicked}
          >
            {busy ? "Saving…" : "Save structure"}
          </button>{" "}
          <button style={s.ghost} disabled={busy} onClick={onCancelClicked}>Cancel</button>
          {dirty && (
            // Sits beside the button that fixes it. A banner at the top of a long editor is
            // scrolled off exactly when someone is about to leave.
            <span style={{ ...s.hint, marginLeft: 12, color: "#7a4f00", fontWeight: 600 }}>
              Not saved yet — these changes are lost if you leave this page.
            </span>
          )}
        </div>

        {confirmDiscard && (
          <div style={s.modalBg} role="dialog" aria-modal="true">
            <div style={s.modal}>
              <h3 style={{ margin: "0 0 12px", fontSize: 17 }}>Discard your changes?</h3>
              <p style={{ fontSize: 13, lineHeight: 1.6 }}>
                The folder levels you edited have not been saved. Leaving now loses them.
              </p>
              <div style={{ marginTop: 16, textAlign: "right" }}>
                <button style={s.ghost} onClick={() => setConfirmDiscard(false)}>
                  Keep editing
                </button>{" "}
                <button style={s.btn} onClick={() => saveStructure().catch(() => undefined)}>
                  Save them
                </button>{" "}
                <button style={s.dangerLg} onClick={cancelEdit}>
                  Discard
                </button>
              </div>
            </div>
          </div>
        )}
      </>
    );
  }

  return (
    <>
      {result !== undefined && resultBox(result)}
      {segments.length === 0 && (
        <div style={{ ...s.msg, ...s.warn }}>
          No business segments are configured, so there is no structure to edit. Segments exist as{" "}
          <strong>mode</strong> rows in the configuration list.
        </div>
      )}
      {segments.map((row) => (
        <div style={s.card} key={row.key}>
          <div style={s.segRow}>
            <div style={{ flex: "1 1 300px" }}>
              <div style={s.segName}>{row.label}</div>
              {/* ONE LINE, NOT TWO (client, 2026-09-06: *"no need to duplicate"*). It used to print
                  the live chain and the staged one stacked so the two could be compared - and the
                  common case is a change that alters ONE level, which rendered as two nearly
                  identical lines and read as a display fault rather than a comparison.

                  ⚠ WITH A CHANGE PENDING THIS SHOWS THE PENDING SHAPE, WHICH IS NOT WHAT UPLOADS
                  ARE USING. That is exactly why it is red, and why the `Change pending` badge sits
                  beside the name: between them they say "this is the shape you asked for, and it is
                  not live yet". Remove either and the line becomes a statement about the CURRENT
                  structure that stays false until the migration applies it. */}
              <div
                style={
                  row.pending === undefined
                    ? s.path
                    : { ...s.path, color: "#a4262c" }
                }
              >
                {pathPreview(row, row.pending ?? row.chain)}
              </div>
            </div>
            <div style={{ flex: "0 0 auto", display: "flex", alignItems: "center", gap: 10 }}>
              <span
                style={{
                  ...s.badge,
                  ...(row.hasDocuments === undefined ? {} : row.hasDocuments ? s.badgeUsed : s.badgeFree),
                }}
              >
                {row.hasDocuments === undefined ? "checking…" : row.hasDocuments ? "In use" : "Empty"}
              </span>
              {row.pending !== undefined && (
                <span style={{ ...s.badge, ...s.badgePending }}>Change pending</span>
              )}
              <button
                style={row.chainError === undefined ? s.ghost : s.off}
                disabled={row.chainError !== undefined}
                onClick={() => startEdit(row)}
              >
                Edit structure
              </button>
            </div>
          </div>
          {row.chainError !== undefined && (
            <div style={{ ...s.msg, ...s.err, marginTop: 10, marginBottom: 0 }}>
              This segment&rsquo;s structure has a problem, so it cannot be edited here:{" "}
              {row.chainError}
            </div>
          )}
        </div>
      ))}
      <p style={{ fontSize: 12, color: "#8a8886", marginTop: 16, lineHeight: 1.6 }}>
        <strong>In use</strong> means documents are already filed under that segment, so adding a
        level leaves two folder shapes side by side. <strong>Empty</strong> means the structure can
        be changed freely.
      </p>
    </>
  );
}
