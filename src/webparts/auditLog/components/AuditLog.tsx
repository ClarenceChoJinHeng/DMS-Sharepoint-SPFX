import * as React from "react";
import { useEffect, useState } from "react";
// The Action list is 26 entries and grows with every new flow event — too long for a native menu.
import { FilterSelect } from "../../../shared/filterSelect";
// The Event column prints a SHORT library name - see `shortLibrary`.
import { documentsLibraryTitle } from "../../../shared/naming";
import { SPHttpClient } from "@microsoft/sp-http";

import {
  ALL_EVENT_TYPES,
  EVENT,
  EVENT_LABEL,
  eventLabelForRow,
} from "../../../shared/auditLog";
import { csvCell, downloadCsv } from "../../../shared/groupExportCsv";
import { primeNames } from "../../../shared/spNaming";
import { isSystemAdmin } from "../../../shared/spGroups";
import {
  AuditListState,
  auditListState,
  auditListTitle,
  AuditCount,
  AuditRecord,
  countAudit,
  plannedAuditListTitle,
  provisionAuditList,
  ProvisionReport,
  readAudit,
  readAuditPage,
  AuditQuery,
} from "../../../shared/spAuditLog";
import { IAuditLogProps } from "./IAuditLogProps";
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/**
 * CRS Audit Log — the viewer.
 *
 * Spec: docs/superpowers/specs/2026-08-13-audit-log-design.md
 *
 * Reads only. There is deliberately no UI to edit or delete a row: a log the audited party can amend
 * is not a log, and the absence of the control states that better than a disabled button would.
 *
 * Admin-only by decision D1, because every row carries a file name and a full path. The page's own
 * permissions are the real boundary — this check is what stops a non-admin who reaches the page from
 * seeing a confusing half-page, and it gates the provisioning control.
 */

/* 10 (client, 2026-09-04: *"Audit log - show 10 list per page, then the next"*). It was 100, then 30
   earlier the same day, then this. An audit row is three lines tall once its path is shown, so ten is
   about a screen. The pager below handles the rest; this only changes how much arrives at once. */
const PAGE_SIZE = 10;

/**
 * The library name as the Event column should PRINT it (client, 2026-09-04: *"For this Restricted &
 * Confidential Document, change to Document only."*).
 *
 * "Moved to Restricted & Confidential Document" wrapped onto two lines in a 150px column and pushed
 * every row taller. The stored `LibraryName` is UNTOUCHED — the row, the CSV and every filter still
 * carry the real title; this shortens the LABEL only.
 *
 * ⚠ MATCHED AGAINST THE LIVE TITLE, NEVER A HARDCODED STRING. This client renames libraries
 * routinely — three times in two days at one point — and a literal "Restricted & Confidential
 * Document" here would silently stop matching on the next rename, putting the long name back on
 * screen with nothing to explain it.
 *
 * ⚠ THE HC LIBRARY IS DELIBERATELY LEFT ALONE. "Moved to Highly Confidential Document" is short
 * enough AND load-bearing: it is the one word in that column telling a reader the document went to
 * the restricted vertical, and collapsing it to "Document" would make an HC routing
 * indistinguishable from an ordinary one.
 */
function shortLibrary(libraryName?: string): string | undefined {
  const raw = (libraryName ?? "").trim();
  if (raw.length === 0) return libraryName;
  return raw.toLowerCase() === documentsLibraryTitle().trim().toLowerCase()
    ? "Document"
    : raw;
}

/** Preset windows. Every query keeps a date bound, so the filter stays on an indexed column. */
const s: Record<string, React.CSSProperties> = {
  /* The Upload Form's page shell (client, 2026-09-04: *"the same padding spacing that the upload
     form is using ... they do not want the pages to stick at the wall"*). `margin: 32px auto` plus
     `padding: 0 24px 48px` is the pattern the Upload Form, My Submissions and the access pages
     already share; this page had the centring and NO padding, so its content met the window edge.
     The max-width is left alone deliberately - each page's is sized for its own content. */
  wrap: {
    maxWidth: 1180,
    margin: "32px auto",
    padding: "0 24px 48px",
    fontFamily: "'Segoe UI', sans-serif",
    color: "#1b1b1b",
  },
  h2: { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  subtitle: {
    fontSize: 13,
    color: "#605e5c",
    margin: "0 0 18px",
    lineHeight: 1.5,
  },
  card: {
    border: "1px solid rgba(217, 217, 217, 1)",
    borderRadius: 8,
    padding: "18px 22px",
    marginBottom: 16,
    background: "rgba(251, 251, 252, 1)",
  },
  cardTitle: {
    fontSize: 16,
    fontWeight: 700,
    textTransform: "uppercase",
    color: "rgba(0, 104, 74, 1)",
    margin: "0 0 10px",
  },
  msg: {
    fontSize: 13,
    padding: "10px 12px",
    borderRadius: 6,
    marginBottom: 16,
    lineHeight: 1.55,
  },
  err: { background: "#fdf3f3", border: "1px solid #f1c9c9", color: "#a4262c" },
  warn: { ...NOTICE_ATTENTION },
  ok: { background: "#f1f8f4", border: "1px solid #c6e3d1", color: "#0f6c3f" },
  info: {
    background: "#f3f2f1",
    border: "1px solid #e1dfdd",
    color: "#323130",
  },
  /* THE FILTER PANEL (client design, 2026-08-30). Fields on the left in a two-column grid, the three
     buttons stacked on the right — which is what keeps Apply beside the fields it applies to instead
     of below a wrapping row where it reads as unrelated. */
  pager: {
    paddingTop: 14,
    display: "flex",
    flexWrap: "wrap",
    gap: 12,
    alignItems: "center",
    justifyContent: "space-between",
  },
  pagerNote: { fontSize: 11.5, color: "#605e5c" },
  pagerBtns: { display: "flex", gap: 4, alignItems: "center" },
  pageBtn: {
    minWidth: 30,
    height: 30,
    padding: "0 8px",
    border: "1px solid #e1dfdd",
    borderRadius: 6,
    background: "#fff",
    color: "#242424",
    fontSize: 12.5,
    cursor: "pointer",
  },
  pageBtnOn: { borderColor: "#0f6c3f", color: "#0f6c3f", fontWeight: 700 },
  filterGrid: {
    display: "flex",
    flexWrap: "wrap",
    gap: 20,
    alignItems: "flex-start",
  },
  filterFields: {
    flex: "1 1 520px",
    display: "grid",
    /* ⚠ EXACTLY TWO COLUMNS (client's design, 2026-09-07): Date | Action, then Person | Keyword.
       `auto-fit` was giving three or four columns on a wide page and one on a narrow one, so the
       pairing the design depends on only held at certain widths.
       The DOM order IS the reading order — Date, Action, Person, Keyword — so the rows come out as
       drawn with no explicit placement. Reorder those four and the design breaks silently.
       `minmax(0, 1fr)` rather than a pixel floor: a grid track defaults to `min-content`, which a
       long placeholder would push wider than its share, forcing the panel to scroll sideways. */
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: 14,
  },
  filterActions: {
    flex: "0 0 auto",
    display: "flex",
    flexDirection: "column",
    gap: 8,
    minWidth: 130,
  },
  /* Flush with the RIGHT edge of the button stack (client, 2026-09-07: *"Ensure the export what is
     shown is align with the buttons"*). It was centred, which left it sitting proud of the buttons
     above and reading as a stray line rather than a caption on Export.
     `alignSelf: stretch` is what makes "right" mean the button edge: without it the span is only as
     wide as its own text and has nothing to align against. */
  actionsNote: {
    fontSize: 10.5,
    color: "#8a8886",
    textAlign: "right",
    alignSelf: "stretch",
  },
  /* ⚠ NO `gridColumn: "1 / -1"` — the Date pair occupies ONE column, so Action lands beside it
     (client, 2026-09-07: *"Can you move the Action to be on the same row as From date"*). Spanning
     both columns pushed Action onto a row of its own and made the panel three rows tall.
     Halving its width is safe because `pairRow` wraps: From/To fall onto two lines inside their own
     column on a narrow page rather than forcing the panel to scroll sideways. */
  fieldPair: {
    display: "flex",
    flexDirection: "column",
    gap: 4,
  },
  pairRow: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" },
  pairTag: {
    fontSize: 14,
    color: "rgba(50, 49, 48, 1)",
    flexShrink: 0,
    fontWeight: 600,
  },
  /* The intro card. Icon left, heading and one line of explanation right — the same shape the CRS
     Settings cards use, so the two pages read as one system. */
  intro: {
    display: "flex",
    flexWrap: "wrap",
    gap: 14,
    alignItems: "flex-start",
    marginBottom: 18,
  },
  introIcon: {
    flexShrink: 0,
    width: 54,
    height: 54,
    /* No `background` and no `borderRadius`: the client's mark draws its own circle, and a rounded
       square behind it showed as a pale box with a circle floating inside. `color` is gone too - the
       artwork is coloured by its own `fill`, not by inheritance. */
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  introHead: {
    fontSize: 14,
    fontWeight: 700,
    margin: "2px 0 4px",
    color: "#242424",
  },
  introBody: { fontSize: 12.5, color: "#605e5c", margin: 0, lineHeight: 1.55 },
  filters: {
    display: "flex",
    flexWrap: "wrap",
    gap: 10,
    alignItems: "flex-end",
    marginBottom: 14,
  },
  field: { display: "flex", flexDirection: "column", gap: 4 },
  label: { fontSize: 14, fontWeight: 600, color: "rgba(50, 49, 48, 1)" },
  // 8px on the filter controls (client, 2026-09-06), matching the upload form's fields. The BUTTONS
  // below keep 4 — the client's design shows square-ish buttons beside rounded inputs.
  input: {
    padding: "7px 9px",
    fontSize: 13,
    border: "1px solid #c8c8c8",
    borderRadius: 8,
    boxSizing: "border-box",
    minWidth: 180,
  },
  select: {
    padding: "7px 9px",
    fontSize: 13,
    border: "1px solid #c8c8c8",
    borderRadius: 8,
  },
  btn: {
    background: "#0f6c3f",
    color: "#fff",
    border: "none",
    borderRadius: 4,
    padding: "8px 16px",
    fontSize: 13,
    fontWeight: 700,
    cursor: "pointer",
    maxWidth: 106,
  },
  ghost: {
    background: "#fff",
    color: "rgba(3, 101, 52, 1)",
    border: "1px solid rgba(0, 104, 74, 1)",
    borderRadius: 4,
    padding: "7px 14px",
    fontWeight: 700,
    fontSize: 13,
    cursor: "pointer",
    maxWidth: 106,
    /* Reset and Export now hold an SVG beside their label (2026-09-06), so the button has to lay its
       children out rather than relying on them being text. `center` on both axes, or the 14px icon
       sits on the text baseline. */
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
  },
  chipRow: { display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12 },
  chip: {
    fontSize: 11,
    borderRadius: 12,
    padding: "4px 10px",
    cursor: "pointer",
    border: "1px solid #c8c8c8",
    background: "#fff",
    color: "#323130",
  },
  chipOn: { border: "1px solid #0f6c3f", background: "#0f6c3f", color: "#fff" },
  head: {
    display: "grid",
    /* ⚠ THE GAP BEFORE "Who" IS ITS OWN COLUMN, not a bigger `columnGap` (client, 2026-09-04:
       *"add more gaps for the Who part ... space out more away from What"*). `columnGap` applies
       between EVERY pair, so widening it would also push "When" away from "Event", which are meant to
       read together. An empty 40px track separates only the pair that needed it.
       The head and the row MUST carry the same template — they are separate grids, and a column added
       to one and not the other misaligns every heading from its column. */
    gridTemplateColumns: "minmax(0,150px) minmax(0,150px) minmax(0,1fr) minmax(0,40px) minmax(0,160px)",
    columnGap: 12,
    padding: "0 10px 8px",
    fontSize: 12,
    fontWeight: 600,
    color: "#605e5c",
    borderBottom: "1px solid #edebe9",
  },
  row: {
    display: "grid",
    gridTemplateColumns: "minmax(0,150px) minmax(0,150px) minmax(0,1fr) minmax(0,40px) minmax(0,160px)",
    columnGap: 12,
    padding: "10px",
    borderBottom: "1px solid #f3f2f1",
    alignItems: "start",
  },
  /* The only clickable heading. Underlined on purpose: a heading that does something must not look
     exactly like the three beside it that do not. */
  sortHead: {
    cursor: "pointer",
    userSelect: "none",
    textDecoration: "underline",
    textDecorationStyle: "dotted",
    color: "#0f6c3f",
  },
  headWho: {
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 8,
  },
  /* Bigger, and LABELLED (client, 2026-09-04: *"Make the refresh button bigger"*). It was a 13px
     glyph in 2px of padding — a hit target smaller than the text beside it, on the control an admin
     presses repeatedly while waiting for a flow to write a row. The word is added because the button
     now RESETS as well as re-reads, and a bare glyph cannot say that. */
  headRefresh: {
    border: "1px solid #c7c7c7",
    background: "#fff",
    borderRadius: 4,
    cursor: "pointer",
    fontSize: 13,
    lineHeight: 1.2,
    padding: "6px 12px",
    color: "#0f6c3f",
    fontWeight: 600,
    whiteSpace: "nowrap",
    // Same reason as `ghost` — this button carries the refresh mark beside its label now.
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
  },
  when: { fontSize: 12, color: "#323130", fontFamily: "Consolas, monospace" },
  type: { fontSize: 12, fontWeight: 600, color: "#0f6c3f" },
  title: { fontSize: 13, color: "#1b1b1b", wordBreak: "break-word" },
  path: {
    fontSize: 11,
    color: "#605e5c",
    wordBreak: "break-all",
    marginTop: 3,
  },
  who: { fontSize: 12, color: "#323130", wordBreak: "break-word" },
  details: {
    fontSize: 11,
    color: "#323130",
    background: "#faf9f8",
    border: "1px solid #edebe9",
    borderRadius: 4,
    padding: "8px 10px",
    marginTop: 6,
    whiteSpace: "pre-wrap",
    fontFamily: "Consolas, monospace",
  },
  link: {
    background: "none",
    border: "none",
    padding: 0,
    color: "#0f6c3f",
    fontSize: 11,
    cursor: "pointer",
    textDecoration: "underline",
  },
  disclose: {
    background: "none",
    border: "none",
    padding: 0,
    color: "#605e5c",
    fontSize: 11,
    cursor: "pointer",
    textAlign: "left",
  },
};

/* ── The client's icons (2026-09-06) ──────────────────────────────────────────
 *
 * WARN: INLINED AS JSX, NEVER REFERENCED AS FILES. Shipping an image asset in this package has
 * failed to provision on this tenant TWICE, both times silently - the New Folder customizer, then
 * the bulk-approve command set - so a `src`-referenced SVG would render as a broken box with nothing
 * explaining it. The root `Group 634044.svg`, `refresh-Icon.svg` and `download-Icon.svg` are the
 * design sources and are NOT built or copied anywhere.
 *
 * WARN: THE COLOURS ARE THE CLIENT'S, NOT `currentColor`. Switching them to inherit would quietly
 * redesign their assets - and the two small ones are drawn with `stroke`, not `fill`, so a `fill`
 * override would do nothing at all and look like the swap had failed.
 */

/** The people mark for the intro card. Carries its OWN pale circle, so its wrapper draws none. */
function AuditLogoIcon(): React.ReactElement {
  return (
    <svg width="54" height="54" viewBox="0 0 61 61" fill="none" aria-hidden="true">
      <circle cx="30.5" cy="30.5" r="30.5" fill="#F1F7EF" />
      <path
        d="M27.2861 30.5C29.8869 30.5001 32.3773 31.5702 34.2109 33.4688C36.0441 35.3669 37.0713 37.938 37.0713 40.6152V42.5381C37.0713 43.3286 36.4479 44 35.6426 44H18.9287C18.1234 44 17.5 43.3286 17.5 42.5381V40.6152C17.5 37.938 18.5272 35.3669 20.3604 33.4688C22.1941 31.5701 24.6852 30.5 27.2861 30.5ZM36.3486 31.8125C36.6234 31.0648 37.447 30.6666 38.1992 30.9629C40.0539 31.6935 41.6475 32.9892 42.7715 34.6748C43.8953 36.3602 44.4975 38.3577 44.5 40.4023V42.5381C44.5 43.3286 43.8765 44 43.0713 44H40.2861C39.4809 44 38.8574 43.3286 38.8574 42.5381C38.8576 41.7477 39.4811 41.0771 40.2861 41.0771H41.6426V40.4053L41.6377 40.1318C41.5859 38.7678 41.1605 37.4462 40.4111 36.3223C39.6117 35.1234 38.4812 34.2062 37.1719 33.6904C36.4291 33.3975 36.0768 32.5524 36.3486 31.8125ZM27.2861 33.4229C25.4543 33.4229 23.6934 34.1765 22.3926 35.5234C21.0912 36.871 20.3575 38.7022 20.3574 40.6152V41.0771H34.2139V40.6152C34.2138 38.7022 33.48 36.871 32.1787 35.5234C30.878 34.1766 29.1178 33.423 27.2861 33.4229ZM27.2861 18C30.3988 18.0002 32.8924 20.6081 32.8926 23.7881C32.8926 26.9682 30.399 29.5769 27.2861 29.5771C24.1731 29.5771 21.6787 26.9684 21.6787 23.7881C21.6789 20.608 24.1732 18 27.2861 18ZM34.7139 18C36.2067 18 37.6355 18.6138 38.6855 19.7012C39.7352 20.7879 40.3212 22.2585 40.3213 23.7881C40.3213 25.3176 39.735 26.7882 38.6855 27.875C37.6355 28.9624 36.2067 29.5771 34.7139 29.5771C33.9088 29.5769 33.2861 28.9057 33.2861 28.1152C33.2862 27.3248 33.9088 26.6545 34.7139 26.6543C35.4373 26.6543 36.1353 26.3566 36.6523 25.8213C37.1701 25.2851 37.4639 24.5535 37.4639 23.7881C37.4638 23.0228 37.17 22.2919 36.6523 21.7559C36.1353 21.2205 35.4374 20.9229 34.7139 20.9229C33.9089 20.9226 33.2863 20.2522 33.2861 19.4619C33.2861 18.6715 33.9088 18.0002 34.7139 18ZM27.2861 20.9229C25.7838 20.9229 24.5363 22.1892 24.5361 23.7881C24.5361 25.3872 25.7836 26.6543 27.2861 26.6543C28.7884 26.6541 30.0361 25.387 30.0361 23.7881C30.0359 22.1893 28.7883 20.9231 27.2861 20.9229Z"
        fill="#00684A"
      />
    </svg>
  );
}

/** The circular arrow on Reset and on the header's Refresh. `stroke`, not `fill`. */
function RefreshIcon(): React.ReactElement {
  return (
    <svg
      width="14"
      height="15"
      viewBox="0 0 14 15"
      fill="none"
      aria-hidden="true"
      style={{ flexShrink: 0 }}
    >
      <path
        d="M12.1307 12.1307C9.52722 14.7342 5.30612 14.7342 2.70262 12.1307C0.0991263 9.52722 0.099126 5.30612 2.70262 2.70262C5.30612 0.0991262 9.52722 0.0991262 12.1307 2.70262M12.75 0.75V3.35C12.75 3.38682 12.7202 3.41667 12.6833 3.41667H10.0833"
        stroke="#00684A"
        strokeWidth="1.5"
        strokeLinecap="round"
      />
    </svg>
  );
}

/** The download arrow on Export. A single filled path, unlike the other two. */
function DownloadIcon(): React.ReactElement {
  return (
    <svg
      width="14"
      height="14"
      viewBox="0 0 14 14"
      fill="none"
      aria-hidden="true"
      style={{ flexShrink: 0 }}
    >
      <path
        d="M5.99998 9.24209C5.99998 9.71153 6.38053 10.0921 6.84998 10.0921C7.31942 10.0921 7.69998 9.71153 7.69998 9.24209H6.84998H5.99998ZM7.69998 1.90875C7.69998 1.43931 7.31942 1.05875 6.84998 1.05875C6.38053 1.05875 5.99998 1.43931 5.99998 1.90875H6.84998L7.69998 1.90875ZM8.9156 3.84313C9.24755 4.17507 9.78574 4.17507 10.1177 3.84313C10.4496 3.51118 10.4496 2.97299 10.1177 2.64105L9.51664 3.24209L8.9156 3.84313ZM7.32138 1.04682L6.72034 1.64787L7.32138 1.04682ZM6.37857 1.04682L6.97961 1.64787L6.97961 1.64787L6.37857 1.04682ZM3.58227 2.64105C3.25032 2.97299 3.25032 3.51118 3.58227 3.84313C3.91421 4.17507 4.4524 4.17507 4.78435 3.84313L4.18331 3.24209L3.58227 2.64105ZM1.69998 9.24209C1.69998 8.77264 1.31942 8.39209 0.849976 8.39209C0.380534 8.39209 -2.44379e-05 8.77264 -2.44379e-05 9.24209H0.849976H1.69998ZM13.7 9.24209C13.7 8.77264 13.3194 8.39209 12.85 8.39209C12.3805 8.39209 12 8.77264 12 9.24209H12.85H13.7ZM11.758 12.3574L11.3721 11.6001L11.3721 11.6001L11.758 12.3574ZM12.632 11.4834L13.3893 11.8693V11.8693L12.632 11.4834ZM1.06796 11.4834L0.310607 11.8693L1.06796 11.4834ZM1.94199 12.3574L1.5561 13.1148H1.5561L1.94199 12.3574ZM6.84998 9.24209H7.69998L7.69998 1.90875L6.84998 1.90875H5.99998L5.99998 9.24209H6.84998ZM9.51664 3.24209L10.1177 2.64105L7.92242 0.445784L7.32138 1.04682L6.72034 1.64787L8.9156 3.84313L9.51664 3.24209ZM6.37857 1.04682L5.77753 0.445784L3.58227 2.64105L4.18331 3.24209L4.78435 3.84313L6.97961 1.64787L6.37857 1.04682ZM7.32138 1.04682L7.92242 0.445784C7.33013 -0.146512 6.36982 -0.146511 5.77753 0.445784L6.37857 1.04682L6.97961 1.64787C6.90802 1.71946 6.79194 1.71946 6.72034 1.64787L7.32138 1.04682ZM0.849976 9.24209H-2.44379e-05V9.37542H0.849976H1.69998V9.24209H0.849976ZM4.04998 12.5754V13.4254H9.64998V12.5754V11.7254H4.04998V12.5754ZM12.85 9.37542H13.7V9.24209H12.85H12V9.37542H12.85ZM9.64998 12.5754V13.4254C10.196 13.4254 10.6566 13.4261 11.0328 13.3953C11.4189 13.3638 11.7901 13.295 12.1438 13.1148L11.758 12.3574L11.3721 11.6001C11.298 11.6378 11.1753 11.678 10.8944 11.701C10.6035 11.7248 10.2241 11.7254 9.64998 11.7254V12.5754ZM12.85 9.37542H12C12 9.9495 11.9993 10.3289 11.9756 10.6198C11.9526 10.9007 11.9124 11.0234 11.8746 11.0975L12.632 11.4834L13.3893 11.8693C13.5696 11.5156 13.6384 11.1443 13.6699 10.7582C13.7006 10.3821 13.7 9.92145 13.7 9.37542H12.85ZM11.758 12.3574L12.1438 13.1148C12.6801 12.8415 13.1161 12.4056 13.3893 11.8693L12.632 11.4834L11.8746 11.0975C11.7644 11.3139 11.5885 11.4898 11.3721 11.6001L11.758 12.3574ZM0.849976 9.37542H-2.44379e-05C-2.44379e-05 9.92145 -0.000685513 10.3821 0.0300469 10.7582C0.06159 11.1443 0.130383 11.5156 0.310607 11.8693L1.06796 11.4834L1.82532 11.0975C1.78756 11.0234 1.74735 10.9007 1.7244 10.6198C1.70064 10.3289 1.69998 9.9495 1.69998 9.37542H0.849976ZM4.04998 12.5754V11.7254C3.4759 11.7254 3.09646 11.7248 2.8056 11.701C2.52466 11.678 2.402 11.6378 2.32789 11.6001L1.94199 12.3574L1.5561 13.1148C1.90981 13.295 2.28109 13.3638 2.66716 13.3953C3.04331 13.4261 3.50395 13.4254 4.04998 13.4254V12.5754ZM1.06796 11.4834L0.310607 11.8693C0.583846 12.4056 1.01984 12.8415 1.5561 13.1148L1.94199 12.3574L2.32789 11.6001C2.1115 11.4898 1.93557 11.3139 1.82532 11.0975L1.06796 11.4834Z"
        fill="#00684A"
      />
    </svg>
  );
}

/**
 * `13/Aug/2026 09:41` — the agreed display format, with the time the feed needs.
 *
 * Parsing here is safe because `odata=nometadata` returns a DateTime field as ISO. This is NOT the
 * trap in gotcha #1: that one is re-parsing a value SharePoint has already formatted for display.
 */
function formatWhen(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${p(d.getDate())}/${months[d.getMonth()]}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/* ⚠ `startOfDaysAgo` and `RANGES` were removed on 2026-08-30 with the "Last 7 days" period select.
   The dates are typed now, and both bounds go through one `bound()` helper inside `load` — which is
   also where the To bound is pushed to the END of its day, so picking a date includes it. */

const CSV_HEADERS = [
  "Event time",
  "Event type",
  "Outcome",
  "Actor",
  "Actor email",
  "Source",
  "Library",
  "Item",
  "Path",
  "Segment",
  "Unit path",
  "Summary",
  "Details",
];

/** Reuses csvCell, which handles quoting AND Excel formula injection — folder names are client text. */
function toAuditCsv(rows: AuditRecord[]): string {
  const lines = [CSV_HEADERS.map(csvCell).join(",")];
  for (const r of rows) {
    lines.push(
      [
        formatWhen(r.EventTime),
        r.EventType,
        r.Outcome,
        r.ActorName,
        r.ActorEmail,
        r.Source,
        r.LibraryName,
        r.ItemName,
        r.ItemPath,
        r.Segment,
        r.UnitPath,
        r.Title,
        r.Details,
      ]
        .map((v) => csvCell(v ?? ""))
        .join(","),
    );
  }
  return lines.join("\r\n");
}

function csvName(now: Date): string {
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `CRS-audit-log-${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())}.csv`;
}

const AuditLog: React.FC<IAuditLogProps> = ({ context, siteUrl }) => {
  const sp: SPHttpClient = context.spHttpClient;

  const [booting, setBooting] = useState(true);
  /* Newest first by default — an audit log is read for what just happened. The toggle is on the
     "When" heading, where the thing it sorts is. */
  const [oldestFirst, setOldestFirst] = useState(false);
  const [admin, setAdmin] = useState(false);
  const [listState, setListState] = useState<AuditListState>("unknown");

  const [rows, setRows] = useState<AuditRecord[]>([]);
  const [next, setNext] = useState<string | undefined>(undefined);
  /* PAGING (client design, 2026-08-30). The old control was "Load more", which APPENDED — fine for
     reading forwards, useless for "go back to the second page".

     ⚠ THE MOCK'S `Showing 1 to 5 of 235 events` WITH A JUMP TO PAGE 47 CANNOT BE BUILT, and the two
     reasons are both about SharePoint rather than effort:
       • A TOTAL needs a count, and this list's `ItemCount` is a CACHED aggregate that lags in both
         directions — it read 2,419 for a list whose view was empty on 2026-08-24. A wrong total on
         screen is worse than none.
       • REST paging is a CONTINUATION TOKEN, so page 47 is only reachable by walking 1–46.
     So: numbered buttons for pages already REACHED, plus Next. Same shape, nothing invented.

     `tokens[i]` is the token that FETCHES page i. `tokens[0]` is undefined — page one needs none. */
  const [tokens, setTokens] = useState<Array<string | undefined>>([undefined]);
  const [pageIdx, setPageIdx] = useState(0);
  /* How many rows the CURRENT query matches, counted rather than taken from `ItemCount` — see
     `countAudit`. `undefined` means the count has not landed or could not be made, and the line then
     falls back to the honest "and more" this page showed before. */
  const [total, setTotal] = useState<AuditCount | undefined>(undefined);
  const [readFailed, setReadFailed] = useState(false);
  const [readStatus, setReadStatus] = useState<number | undefined>(undefined);
  const [loading, setLoading] = useState(false);

  /* ⚠ BOTH BOUNDS OPTIONAL AND INDEPENDENT. Blank From means "from the beginning" — the query that
     is NOT index-bounded, which used to be the deliberate "All time" choice and is now simply what
     an empty box means. Left empty on arrival so the first view is everything recent rather than a
     window the admin did not choose. */
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [types, setTypes] = useState<string[]>([]);
  const [actor, setActor] = useState("");
  const [text, setText] = useState("");
  /** Set when the user pivots to one file's history. */
  const [focus, setFocus] = useState<{ id: string; name: string } | undefined>(
    undefined,
  );

  const [expanded, setExpanded] = useState<Record<number, boolean>>({});
  const [showLimits, setShowLimits] = useState(false);
  const [provisioning, setProvisioning] = useState(false);
  const [report, setReport] = useState<ProvisionReport | undefined>(undefined);

  /* -- Load ---------------------------------------------------------------- */

  /**
   * The filters as a query.
   *
   * ⚠ ONE DEFINITION, because `goToPage` re-runs page one and MUST ask exactly what `load` asked.
   * Two copies of this would drift, and the symptom would be paging that quietly changes the result
   * set as you move through it — which reads as the log being wrong rather than the query being two
   * different queries.
   */
  /**
   * ⚠ `overrides` EXISTS BECAUSE `setState` HAS NOT LANDED INSIDE THE HANDLER THAT CALLS THIS.
   *
   * Every field below is read from React state, so a handler that clears the boxes and then calls
   * `load` in the same tick builds its query from the values that were there BEFORE the clear — the
   * screen shows empty filters and the rows come back filtered. That is precisely the state the Reset
   * button's own comment says it exists to avoid, and Reset had the defect too.
   *
   * Passing the cleared values explicitly is what actually drops them from the query; the setters
   * only keep the controls agreeing with it on the next render.
   */
  const currentQuery = (
    focusId?: string,
    overrides?: Partial<AuditQuery>,
  ): AuditQuery => {
    /* ⚠ AN UNPARSEABLE DATE IS NO BOUND, never today's. A half-typed `2026-08-` in a date input
       yields an Invalid Date, and passing that into the query would filter on NaN and return
       nothing at all — which reads as "there are no events" rather than "that date is incomplete". */
    const bound = (v: string): Date | undefined => {
      const d = v ? new Date(v) : undefined;
      return d && !isNaN(d.getTime()) ? d : undefined;
    };
    const to = bound(dateTo);
    return {
      from: bound(dateFrom),
      /* The To box means the whole of that DAY. Without this, picking the 30th excludes everything
         that happened on the 30th, because the value parses to midnight at its start. */
      to: to ? new Date(to.getTime() + 24 * 60 * 60 * 1000 - 1) : undefined,
      eventTypes: types,
      actorEmail: actor.trim(),
      text: text.trim(),
      itemUniqueId: focusId ?? focus?.id,
      top: PAGE_SIZE,
      oldestFirst,
      ...(overrides ?? {}),
    };
  };

  const load = async (
    focusId?: string,
    overrides?: Partial<AuditQuery>,
  ): Promise<void> => {
    setLoading(true);
    const q = currentQuery(focusId, overrides);
    /* ⚠ THE QUERY IS BUILT ONCE AND SHARED. Calling `currentQuery` a second time for the count would
       be a second reading of the same state — and with `overrides` in play (Reset, Refresh) the two
       could differ, giving a total that describes a query nobody ran. */
    setTotal(undefined);
    /* Counted in parallel, and deliberately NOT awaited: the rows must not wait on it. The line
       reads "and more" for the moment it is in flight, which is what it said permanently until
       today. `countAudit` never throws; the catch is belt and braces. */
    countAudit(sp, siteUrl, q)
      .then((c) => setTotal(c))
      .catch(() => setTotal(undefined));
    const page = await readAudit(sp, siteUrl, q);
    setRows(page.rows);
    setNext(page.next);
    /* ⚠ A NEW QUERY INVALIDATES EVERY TOKEN. A continuation token belongs to the query that
       produced it, so reusing one after the filters changed would page through the OLD result set
       while the screen showed the new filters — rows that match nothing the admin asked for, with
       nothing to explain them. */
    setTokens([undefined]);
    setPageIdx(0);
    setReadFailed(page.failed);
    setReadStatus(page.status);
    setExpanded({});
    setLoading(false);
  };

  useEffect(() => {
    const init = async (): Promise<void> => {
      // NAMES FIRST. An unprimed cache resolves to the legacy `DMS …` titles, which 404 on a renamed
      // site — the exact failure that took the abbreviation page down on 2026-08-12.
      await primeNames(sp, siteUrl).catch(() => undefined);

      /* Site Collection Admin OR a member of the site OWNERS group - see `isSystemAdmin`.
         WARN: THIS CHECKED `IsSiteAdmin` ALONE UNTIL 2026-08-27, which is why adding somebody to
         `CRS Owners` did not let them use this page: a group can never confer `IsSiteAdmin`, so the
         gate could only ever be satisfied one person at a time in Site Settings. Still fails closed -
         it gates a CONTROL, and the page's own permissions keep non-admins out. */
      setAdmin(await isSystemAdmin(sp, siteUrl));

      const state = await auditListState(sp, siteUrl);
      setListState(state);
      setBooting(false);
      if (state === "ready") await load(undefined);
    };
    init().catch(() => setBooting(false));
  }, []);

  /**
   * Show one page.
   *
   * REPLACES the rows rather than appending, which is the whole difference from the "Load more" this
   * grew out of: a page is a position, not a growing list.
   *
   * ⚠ ONLY A PAGE ALREADY REACHED, OR THE NEXT ONE. Anything further has no token, so there is
   * nothing to fetch it with — the buttons for those are not rendered, and this refuses as well, so
   * the rule holds even if a caller is added later.
   */
  const goToPage = async (idx: number): Promise<void> => {
    if (idx < 0 || idx > tokens.length) return;
    const token = idx === 0 ? undefined : tokens[idx];
    if (idx > 0 && token === undefined) return;
    setLoading(true);
    const page =
      token === undefined
        ? await readAudit(sp, siteUrl, currentQuery())
        : await readAuditPage(sp, token);
    setRows(page.rows);
    setPageIdx(idx);
    setNext(page.next);
    /* Remember how to fetch the page AFTER this one, so Next works from wherever we land — including
       after jumping backwards, where `next` would otherwise still describe the page we left. */
    setTokens((prev) => {
      const out = prev.slice(0, idx + 1);
      if (page.next !== undefined) out[idx + 1] = page.next;
      return out;
    });
    setExpanded({});
    setLoading(false);
  };

  const provision = async (): Promise<void> => {
    setProvisioning(true);
    const r = await provisionAuditList(sp, siteUrl);
    setReport(r);
    const state = await auditListState(sp, siteUrl);
    setListState(state);
    setProvisioning(false);
    if (state === "ready") await load(undefined);
  };

  /** Pivot to one file's history. Applies immediately — a filter nobody applied is a filter nobody
   * trusts, and this one is a single click from a row they are already reading. */
  /* What "no filters" means, in ONE place. Refresh and Reset both clear, and two literals would be
     two chances for one of them to keep a field the other drops. `itemUniqueId` is NOT here: Reset
     keeps the focus (it is a filter-card control, sitting with the fields it clears) while Refresh
     drops it, so that difference is passed at the call site rather than buried in here. */
  const CLEARED: Partial<AuditQuery> = {
    from: undefined,
    to: undefined,
    eventTypes: [],
    actorEmail: "",
    text: "",
  };

  const focusOn = (id: string, name: string): void => {
    setFocus({ id, name });
    load(id).catch(() => undefined);
  };

  const clearFocus = (): void => {
    setFocus(undefined);
    load(undefined).catch(() => undefined);
  };

  /* `toggleType` went with the chip row — the Action select writes `types` directly. `types` stays an
     ARRAY so `readAudit` is unchanged and multi-select is a control swap away. */

  /* -- Render -------------------------------------------------------------- */

  if (booting) {
    return (
      <div style={s.wrap}>
        <p style={s.subtitle}>Loading the audit log…</p>
      </div>
    );
  }

  const header = (
    <>
      <h2 style={s.h2}>CRS Audit Log</h2>
      {/* ⚠ THE "cannot be edited or deleted" SENTENCE IS NOT DECORATION and is kept below. It is the
          claim that makes this log worth reading at all — writes are restricted to Owners and the
          service account by design, and an admin who does not know that has no reason to trust a row
          in front of them. */}
      <div style={s.intro}>
        {/* The client's mark replaces the hand-drawn one (2026-09-06). It carries its own pale
            circle, so `s.introIcon` no longer draws a background or a corner radius - leaving them
            would put a rounded square behind a circle. */}
        <span style={s.introIcon} aria-hidden="true">
          <AuditLogoIcon />
        </span>
        <div>
          {/* ⚠ THE HEADING IS REMOVED, THE BODY IS NOT (client, 2026-09-07, marking just that line
              REMOVE). The paragraph below still says what the log is and — the load-bearing half —
              that records are written automatically and cannot be edited or deleted from this page,
              which is the tamper-resistance claim the whole list exists to make.
              `s.introHead` is left in the style record on purpose: it costs nothing, and deleting a
              style to match one removal is how the next heading added here loses its formatting. */}
          <p style={s.introBody}>
            Track activities and changes across the system, including document
            actions, access changes, and folder updates. Records are written
            automatically, and cannot be edited or deleted from this page.
          </p>
        </div>
      </div>
    </>
  );

  const limitsPanel = (
    <div style={s.card}>
      <button style={s.disclose} onClick={() => setShowLimits(!showLimits)}>
        {showLimits ? "▾" : "▸"} What this log cannot tell you
      </button>
      {showLimits && (
        <div
          style={{
            fontSize: 12,
            color: "#323130",
            lineHeight: 1.6,
            marginTop: 10,
          }}
        >
          <p style={{ margin: "0 0 8px" }}>
            A log that looks complete but is not would be worse than none, so
            these gaps are stated rather than left to be discovered:
          </p>
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            <li>
              <strong>Who viewed or downloaded a file.</strong> Nothing tells
              this system that a file was read. That record exists only in
              Microsoft Purview, which needs a tenant-wide administrator consent
              this site-scoped solution cannot request. Ask your Microsoft 365
              administrator for a Purview audit report.
            </li>
            <li>
              <strong>Failed access attempts.</strong> A denied request never
              reaches this system.
            </li>
            <li>
              <strong>Anything from before this log was switched on.</strong> It
              starts from that day and cannot be filled in backwards.
            </li>
            <li>
              <strong>
                Permission changes made in SharePoint&apos;s own “Manage access”
                dialog
              </strong>{" "}
              instead of on the Folder Access page. Use the Folder Access page
              and they are recorded.
            </li>
            <li>
              <strong>Term store edits.</strong> Adding, renaming or deleting a
              term is not reported to this system. The consequence is recorded
              instead — the folder rename, or the orphaned term that
              reconciliation found.
            </li>
            <li>
              <strong>Rows changed directly in the list.</strong> The
              list&apos;s version history is the only trace, which is why it is
              switched on.
            </li>
          </ul>
        </div>
      )}
    </div>
  );

  // ---- The list is not there yet. Distinct from "unreadable" below.
  if (listState === "absent") {
    return (
      <div style={s.wrap}>
        {header}
        <div style={{ ...s.msg, ...s.info }}>
          <strong>The audit log is not set up yet.</strong>{" "}
          {admin
            ? "Nothing has been recorded so far. Setting it up creates the list that events are written to."
            : "Ask an administrator to open this page once — it sets itself up."}
        </div>
        {admin && (
          <div style={s.card}>
            <p style={s.cardTitle}>Set up</p>
            <p
              style={{
                fontSize: 13,
                color: "#323130",
                lineHeight: 1.6,
                margin: "0 0 12px",
              }}
            >
              {/* plannedAuditListTitle, NOT auditListTitle: while the list is absent the latter has
                  fallen back to the legacy `DMS …` name, so this panel offered to create a list that
                  the button would not create. Both now come from one derivation. */}
              This creates the <strong>{plannedAuditListTitle()}</strong> list,
              its columns and its indexes, and turns on version history. It
              changes nothing else, and is safe to run again.
            </p>
            <p style={{ ...s.msg, ...s.warn, marginBottom: 12 }}>
              <strong>One step is left to you afterwards.</strong> In the new
              list&apos;s permission settings, stop inheriting permissions and
              give <em>Contribute</em> to the service account the Power Automate
              flows run as — keeping Full Control for owners, and removing
              everyone else. This is not automated on purpose: the code does not
              know which account that is, and a wrong guess would lock the flows
              out of the list they write to.
            </p>
            <button
              style={s.btn}
              disabled={provisioning}
              onClick={() => provision().catch(() => undefined)}
            >
              {provisioning ? "Setting up…" : "Set up the audit log"}
            </button>
          </div>
        )}
        {report && report.problems.length > 0 && (
          <div style={{ ...s.msg, ...s.err }}>
            {report.problems.map((p, i) => (
              <div key={i}>{p}</div>
            ))}
          </div>
        )}
        {limitsPanel}
      </div>
    );
  }

  // ---- We could not tell. Never presented as "nothing happened".
  if (listState === "unknown") {
    return (
      <div style={s.wrap}>
        {header}
        <div style={{ ...s.msg, ...s.err }}>
          <strong>The audit log could not be read.</strong> This is a
          permissions or connection problem, not an empty log — events may well
          have been recorded. Reload the page, and if it persists, check that
          you have access to the <strong>{auditListTitle()}</strong> list.
        </div>
        {limitsPanel}
      </div>
    );
  }

  if (!admin) {
    return (
      <div style={s.wrap}>
        {header}
        <div style={{ ...s.msg, ...s.info }}>
          This page is for administrators. Audit records name files and folder
          paths across every business segment, so they are not shown more
          widely.
        </div>
      </div>
    );
  }

  return (
    <div style={s.wrap}>
      {header}

      {report && (
        <div
          style={{ ...s.msg, ...(report.problems.length > 0 ? s.warn : s.ok) }}
        >
          {report.createdList && (
            <div>Created the {auditListTitle()} list.</div>
          )}
          {report.createdColumns.length > 0 && (
            <div>Added {report.createdColumns.length} columns.</div>
          )}
          {report.problems.map((p, i) => (
            <div key={i}>{p}</div>
          ))}
        </div>
      )}

      {focus && (
        <div style={{ ...s.msg, ...s.info }}>
          {/* Named for the control that opens it (client, 2026-09-04) — a banner that calls this
              "the whole history" while the button says "File Activity Log" leaves the reader working
              out whether they are two different things. */}
          <strong>File Activity Log</strong> for{" "}
          <strong>{focus.name || "this item"}</strong> — every event, including
          any time it was moved or renamed.{" "}
          <button style={s.link} onClick={clearFocus}>
            Show everything again
          </button>
        </div>
      )}

      <div style={s.card}>
        <p style={s.cardTitle}>Filter</p>
        <div style={s.filterGrid}>
          <div style={s.filterFields}>
            {/* ⚠ EXPLICIT DATES REPLACED THE "Last 7 days / 30 / 90 / All time" PERIOD SELECT
                (client design, 2026-08-30). `AuditQuery` already carried a `to` bound, so this is a
                wiring change rather than a new capability.

                What was traded: "Last 7 days" was ONE click and is now two dates. What was gained:
                an investigation into a specific week no longer has to widen to 30 days and read past
                everything else. Both bounds are OPTIONAL and independently so — a From with no To
                means "since then", which is the old behaviour with a date the admin chose. */}
            <div style={s.fieldPair}>
              <span style={s.label}>Date</span>
              <div style={s.pairRow}>
                <span style={s.pairTag}>From</span>
                <input
                  type="date"
                  style={s.input}
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                />
                <span style={s.pairTag}>To</span>
                <input
                  type="date"
                  style={s.input}
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                />
              </div>
            </div>

            <div style={s.field}>
              <span style={s.label}>Action</span>
              {/* ⚠ ONE ACTION, WHERE THE CHIPS ALLOWED SEVERAL. The chip row let an admin pick
                  Uploaded AND Approved together; this cannot. Followed the mock because the client's
                  standing instruction is fewer inputs (*"don't give client too many features or
                  inputs, it will make them scared"*) — but `types` is still an ARRAY all the way into
                  `readAudit`, so restoring multi-select is a control swap, not a query change. */}
              {/* ⚠ A CUSTOM CONTROL, NOT A NATIVE `<select>` (client, 2026-09-04: *"can you use a
                  customize dropdown? Is so long"*). `ALL_EVENT_TYPES` is 26 entries and grows every
                  time a flow learns a new event, so the native menu had become a wall.

                  ⚠ THE VALUE CONTRACT IS UNCHANGED: `types` is still a `string[]` into `readAudit`,
                  still holding at most one entry. So this is a CONTROL swap, not a query change — and
                  restoring multi-select later remains the same one-line possibility the old comment
                  described.

                  The label falls back to the raw stored value for an event type with no `EVENT_LABEL`
                  entry, which is the same fallback the rows use. That matters: a flow can introduce a
                  type this codebase has never heard of, and a filter that silently omitted it would
                  make those rows unfindable — the exact gap `Replaced` and `ShareRevoked` sat in
                  until today. */}
              <FilterSelect
                value={types[0] ?? ""}
                onChange={(v) => setTypes(v ? [v] : [])}
                anyLabel="Select action"
                placeholder="Select action"
                searchPlaceholder="Type to filter actions…"
                /* WARN: THE TWO ACCESS EVENTS ARE HIDDEN FROM THE PICKER, NOT FROM THE LOG (client,
                   2026-09-06: *"REMOVE THESE OPTIONS"*). Reconciliation still WRITES
                   `AccessGranted`/`AccessRevoked` rows on every run, so they remain in the list and
                   in the CSV - they simply cannot be filtered TO any more. Raised with the client
                   and accepted. Filtered here rather than removed from `ALL_EVENT_TYPES`, which is
                   also what the viewer reads back to label a row: dropping them there would leave
                   those rows showing a raw `AccessGranted` instead of "Access granted". */
                /* ⚠ `shareRevoked` JOINS THE HIDDEN SET (client, 2026-09-06). The revoke control was
                   removed from the Requests page, so nothing in the app raises this any more — but
                   ROWS ALREADY EXIST from before it was removed, and the `CRS — Audit request
                   activity` flow still writes one if `RevokedBy` is ever set. They stay in the log
                   and in the CSV; they simply cannot be filtered TO. */
                options={ALL_EVENT_TYPES.filter(
                  (t) =>
                    t !== EVENT.accessGranted &&
                    t !== EVENT.accessRevoked &&
                    t !== EVENT.shareRevoked,
                ).map((t) => ({
                  value: t,
                  label: EVENT_LABEL[t] ?? t,
                }))}
              />
            </div>

            <div style={s.field}>
              <span style={s.label}>Person (email address)</span>
              <input
                style={s.input}
                value={actor}
                placeholder="someone@abc.com"
                onChange={(e) => setActor(e.target.value)}
              />
            </div>

            <div style={s.field}>
              <span style={s.label}>Keyword</span>
              <input
                style={s.input}
                value={text}
                placeholder="Tax return"
                onChange={(e) => setText(e.target.value)}
              />
            </div>
          </div>

          <div style={s.filterActions}>
            <button
              style={s.btn}
              disabled={loading}
              onClick={() => load(focus?.id).catch(() => undefined)}
            >
              {loading ? "Loading…" : "Apply"}
            </button>
            {/* ⚠ RESET CLEARS AND RE-READS. Clearing the boxes without running the query leaves the
                admin looking at the OLD result under empty filters, which is the state that reads as
                "the page is broken". */}
            <button
              style={s.ghost}
              disabled={loading}
              onClick={() => {
                setDateFrom("");
                setDateTo("");
                setTypes([]);
                setActor("");
                setText("");
                /* ⚠ THE CLEARED VALUES ARE PASSED, not left to state — see `currentQuery`. Without
                   this, Reset emptied the boxes and re-read with the OLD filters, which is exactly
                   the "old result under empty filters" the comment above says it avoids. */
                load(focus?.id, CLEARED).catch(() => undefined);
              }}
            >
              <RefreshIcon />
              Reset
            </button>
            <button
              style={s.ghost}
              disabled={rows.length === 0}
              onClick={() => downloadCsv(toAuditCsv(rows), csvName(new Date()))}
            >
              <DownloadIcon />
              Export
            </button>
            <span style={s.actionsNote}>export what is shown</span>
          </div>
        </div>
        <p style={{ fontSize: 11, color: "#605e5c", margin: "10px 0 0" }}>
          Press Apply to use the filters.
        </p>
      </div>

      {readFailed && (
        <div style={{ ...s.msg, ...s.err }}>
          <strong>
            The log could not be read{readStatus ? ` (HTTP ${readStatus})` : ""}
            .
          </strong>{" "}
          This is not an empty result — do not read it as “nothing happened”.
        </div>
      )}

      {!readFailed && rows.length === 0 && (
        <div style={{ ...s.msg, ...s.info }}>
          <strong>No events match these filters.</strong> The log was read
          successfully and this filter simply found nothing — try a longer
          period, or clear the event types.
        </div>
      )}

      {rows.length > 0 && (
        <div style={s.card}>
          {/* A SECOND refresh, on the results header (client, 2026-08-24: *"can you add a refesh
              button like beside the column"*). The one in the filter card is 400–600px above this
              on a full page of rows, so the person watching for a flow to write a row has to scroll
              back up to ask for it again — and scrolling up to a filter panel is what makes people
              reload the whole page instead.

              Right-aligned INSIDE the last grid column, so the four headings stay aligned with the
              four columns of every row below. Icon-only with a `title`: a second "Refresh" wide
              enough to read would push "Who" out of alignment, and the glyph is unambiguous next to
              a table. Runs the identical `load` call, focus included. */}
          <div style={s.head}>
            {/* ⚠ CHANGING THE SORT RE-READS FROM THE SERVER, and passes the new direction
                EXPLICITLY — `setState` has not landed when `load` runs in this handler, so relying on
                state would sort by the PREVIOUS direction while the arrow showed the new one. Same
                trap as the cleared filters a few lines below.
                It also returns to page 1: a continuation token belongs to the query that produced it,
                and reversing the order makes every token meaningless. `load` already resets them. */}
            <span
              role="button"
              tabIndex={0}
              style={s.sortHead}
              title={
                oldestFirst
                  ? "Showing oldest first — click for newest"
                  : "Showing newest first — click for oldest"
              }
              onClick={() => {
                const next = !oldestFirst;
                setOldestFirst(next);
                load(focus?.id, { oldestFirst: next }).catch(() => undefined);
              }}
              onKeyDown={(e) => {
                if (e.key !== "Enter" && e.key !== " ") return;
                e.preventDefault();
                const next = !oldestFirst;
                setOldestFirst(next);
                load(focus?.id, { oldestFirst: next }).catch(() => undefined);
              }}
            >
              When{" "}
              {/* Bigger than the surrounding label (client, 2026-09-06). It is the only control in
                  this header row, and at the text's own size it read as punctuation rather than as
                  something clickable. `lineHeight: 0` keeps the taller glyph from growing the row. */}
              <span
                aria-hidden="true"
                style={{ fontSize: 18, lineHeight: 0, verticalAlign: "-2px" }}
              >
                {oldestFirst ? "↑" : "↓"}
              </span>
            </span>
            <span>Event</span>
            <span>What</span>
            {/* The empty 40px track that holds "Who" away from "What" — see `s.head`. */}
            <span aria-hidden="true" />
            <span style={s.headWho}>
              Who
              {/* ⚠ REFRESH NOW CLEARS THE FILTERS AND THE FOCUS (client, 2026-09-04: *"ensure if it
                  refreshes it reset and show everything again"*). It used to re-read with whatever
                  was set, focus included.

                  ⚠ THIS MAKES IT BEHAVE LIKE THE "Reset" BUTTON IN THE FILTER CARD, which does the
                  same clearing but does NOT drop the focus. They are no longer distinct enough to
                  need both; Reset is left in place because it sits with the fields it clears, where
                  an admin looks for it. Say the word and one of them goes.

                  Clearing STATE and passing `undefined` both matter: `load` builds its query from
                  `currentQuery`, which reads the state — but `setState` has not landed inside this
                  handler, so the explicit `undefined` is what actually drops the focus for THIS
                  read. The setters are what keep the boxes on screen agreeing with it. */}
              <button
                type="button"
                style={s.headRefresh}
                disabled={loading}
                title="Clear the filters and re-read the whole log"
                aria-label="Refresh and show everything"
                onClick={() => {
                  setDateFrom("");
                  setDateTo("");
                  setTypes([]);
                  setActor("");
                  setText("");
                  setFocus(undefined);
                  load(undefined, {
                    ...CLEARED,
                    itemUniqueId: undefined,
                  }).catch(() => undefined);
                }}
              >
                {loading ? (
                  "…"
                ) : (
                  <>
                    <RefreshIcon />
                    Refresh
                  </>
                )}
              </button>
            </span>
          </div>
          {rows.map((r) => {
            const open = expanded[r.Id] === true;
            return (
              <div key={r.Id} style={s.row}>
                <span style={s.when}>{formatWhen(r.EventTime)}</span>
                <span style={s.type}>
                  {eventLabelForRow(r.EventType, shortLibrary(r.LibraryName))}
                  {r.Outcome && r.Outcome !== "Success" && (
                    <span style={{ color: "#a4262c" }}> · {r.Outcome}</span>
                  )}
                </span>
                <div>
                  <div style={s.title}>{r.Title}</div>
                  {r.ItemPath && <div style={s.path}>{r.ItemPath}</div>}
                  <div
                    style={{
                      display: "flex",
                      gap: 12,
                      marginTop: 4,
                      flexWrap: "wrap",
                    }}
                  >
                    {r.Details && (
                      <button
                        style={s.disclose}
                        onClick={() =>
                          setExpanded({ ...expanded, [r.Id]: !open })
                        }
                      >
                        {open ? "▾ Hide details" : "▸ Details"}
                      </button>
                    )}
                    {/* ⚠ HIDDEN ONCE ALREADY FOCUSED, and that is a real bug fix rather than
                        tidying (client, 2026-09-04: *"the list is showing HIstory of file which when
                        I click it doesnt do anything"*). Inside one file's log EVERY row is that same
                        file, so `focusOn` was being handed the id it is already focused on: same
                        query, same rows, nothing observable. A control that does nothing reads as
                        broken, and an admin clicks it repeatedly before concluding the page is at
                        fault.

                        Gated on `focus === undefined` rather than on `r.ItemUniqueId !== focus.id` —
                        every row in a focused view carries that id by definition, so the narrower
                        test is the same test with more ways to go wrong. */}
                    {r.ItemUniqueId && focus === undefined && (
                      <button
                        style={s.link}
                        onClick={() => focusOn(r.ItemUniqueId, r.ItemName)}
                      >
                        File Activity Log
                      </button>
                    )}
                  </div>
                  {open && r.Details && (
                    <div style={s.details}>{r.Details}</div>
                  )}
                </div>
                {/* ⚠ `Source` IS NO LONGER SHOWN UNDER THE NAME (client, 2026-08-30). It is still
                    STORED on every row and still EXPORTED in the CSV — only the column is gone.
                    Worth knowing what was given up: `Source` is what distinguishes a row a PERSON
                    wrote (`CrsConfiguration`, `Requests`) from one a FLOW wrote
                    (`Flow:DocumentsDeletions`), which is the difference between "somebody did this"
                    and "the system did this". If that question ever comes up in an investigation,
                    the CSV still answers it. */}
                {/* The empty 40px track — see `s.head`. It sits in BOTH grids or the headings stop
                    lining up with their columns. */}
                <span aria-hidden="true" />
                <div style={s.who}>
                  {/* WARN: DISPLAY ONLY - THE ACTOR IS STILL STORED, AND THE CSV STILL CARRIES IT.
                      Archiving and routing are done BY A SCHEDULED FLOW, so whatever name the row
                      holds is whoever the flow ran as, not a person who decided anything - and
                      showing it invites exactly that reading (client, 2026-09-06). The row keeps
                      `ActorName`/`ActorEmail` for an investigation; only this cell is blanked. */}
                  <div>
                    {r.EventType === EVENT.archived ||
                    r.EventType === EVENT.routed
                      ? "-"
                      : r.ActorName || r.ActorEmail || "—"}
                  </div>
                </div>
              </div>
            );
          })}
          {/* PAGINATION (client design, 2026-08-30). Numbered, but only over pages that have a
              token — see the note on `tokens`. `…` is shown when more exist beyond them, which is
              honest about there being more without claiming to know how many. */}
          <div style={s.pager}>
            <span style={s.pagerNote}>
              Showing {rows.length === 0 ? 0 : pageIdx * PAGE_SIZE + 1} to{" "}
              {pageIdx * PAGE_SIZE + rows.length}
              {/* THE TOTAL (client, 2026-09-04, on their mock's `Showing 1 to 5 of 235 events`:
                  *"Add it, just follow what the mockup would want."*).

                  ⚠ IT IS COUNTED, NOT `ItemCount`. That aggregate is cached and lags in BOTH
                  directions — it read 2,419 for a list whose view was empty (2026-08-24) — and it
                  cannot answer a filtered question at all. `countAudit` walks Ids instead.

                  ⚠ THREE STATES, and they must stay apart. A counted total; a floor (`25,000+`)
                  when the walk hit its cap, because a floor shown as a total is the one thing this
                  page must never do; and `and more` when the count could not be made, which is
                  exactly what this line said before today. */}
              {total === undefined
                ? ""
                : ` of ${total.total.toLocaleString()}${total.exact ? "" : "+"}`}{" "}
              event
              {(total !== undefined ? total.total : rows.length) === 1
                ? ""
                : "s"}
              {total === undefined && next !== undefined ? " and more" : ""}
            </span>
            <div style={s.pagerBtns}>
              <button
                style={s.pageBtn}
                disabled={loading || pageIdx === 0}
                title="Previous page"
                onClick={() => goToPage(pageIdx - 1).catch(() => undefined)}
              >
                &lsaquo;
              </button>
              {tokens.map((_, i) => (
                <button
                  key={i}
                  style={
                    i === pageIdx ? { ...s.pageBtn, ...s.pageBtnOn } : s.pageBtn
                  }
                  disabled={loading}
                  onClick={() => goToPage(i).catch(() => undefined)}
                >
                  {i + 1}
                </button>
              ))}
              {next !== undefined && <span style={s.pagerNote}>&hellip;</span>}
              <button
                style={s.pageBtn}
                disabled={loading || next === undefined}
                title="Next page"
                onClick={() => goToPage(pageIdx + 1).catch(() => undefined)}
              >
                &rsaquo;
              </button>
            </div>
          </div>
        </div>
      )}

      {limitsPanel}
    </div>
  );
};

export default AuditLog;
