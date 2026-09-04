// CRS Search — find a document by its metadata.
//
// Spec: docs/superpowers/specs/2026-08-16-document-search-design.md
//
// THIS WEB PART ENFORCES NO PERMISSIONS, and must never be changed to.
//
// Both engines are security-trimmed by SharePoint before a row reaches this code: the Search API
// applies ACLs at query time, and a REST list read returns only items the caller can open. So the
// hierarchy the client described — PIC sees their unit, Head of Department sees the department,
// C-Level sees the segment or everything — falls out of the folder ACLs reconciliation already
// grants. A role check written here could only ever be a second, drifting copy of a model whose
// source of truth is groupMapModel.ts and the live ACLs, and unlike this design it could be wrong
// in the dangerous direction.
//
// WHAT IT IS NOT: a privacy screen. A PIC searching finds their whole unit's approved documents,
// because that is what they can already open by browsing. My Submissions is the author-filtered
// page and stays that way.
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDocumentSearchProps } from "./IDocumentSearchProps";
import {
  LibraryResult,
  SearchCriteria,
  SearchHit,
  SearchLibrary,
  buildKql,
  buildListFilter,
  buildRecentFilter,
  emptyCriteria,
  failedLibraries,
  hasCriteria,
  hasMetadataFilter,
  isHcLibrary,
  METADATA_FILTER_FIELDS,
  metadataFilterMatches,
  kqlPathScope,
  mergeHits,
  recencyCutoff,
  searchState,
  sortByModified,
} from "../../../shared/documentSearch";
// The metadata panel DERIVES its tier rows from the item's own fields, which is what makes
// Region / Estate·Mill appear on a segment nobody wrote code for. A third hand-written copy is
// avoided deliberately: both existing copies carried the hardcoded-tier bug.
import { buildDetailRows, formatBytes } from "../../../shared/documentDetails";
// The preview strategy is already built and tested — SharePoint serves an Office file as a
// DOWNLOAD, so a raw URL in an iframe renders nothing. Reused rather than re-guessed.
import { previewTarget } from "../../../shared/filePreview";
import {
  DOCUMENTS_URL_SEGMENT,
  documentsLibraryTitle,
  LIST_SUFFIX,
  cachedHcLibraries,
  cachedArchiveLibraries,
  cachedListTitle,
  libApiTitle,
  libraryTitle,
} from "../../../shared/naming";
import { primeNames } from "../../../shared/spNaming";
// The `Keyword` column is created by reconciliation, so a library provisioned earlier may not have
// it — and a $filter naming an absent column 400s the whole read. Probed, never assumed.
import { libraryHasColumns, KEYWORD_COLUMN } from "../../../shared/optionalColumns";
// Admin-only, for the banner diagnostic below: owners-group membership counts, not just IsSiteAdmin.
import { isSystemAdmin } from "../../../shared/spGroups";
// One palette for every attention banner — the client asked for the Retire-a-segment colour.
import { NOTICE_ATTENTION } from "../../../shared/noticeStyles";

/* ─────────────────────────────── Term set ids ─────────────────────────────── */

/**
 * The three fixed dropdowns' term sets.
 *
 * Read from the DMS Config `termSet_*` rows at mount; these are only the fallback for when that
 * read fails, and they are THIS site's ids. Term set ids are per site since the in-site term store
 * pivot (2026-07-27), so a hardcoded id is wrong on any other site — which is why the config read
 * wins and this is a last resort rather than the source.
 */
const FALLBACK_TERM_SETS = {
  documentType: "866c5754-258e-401f-8685-03d20ae59b1d",
  year: "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf",
  confidentiality: "0d6d1da8-27e5-477f-8684-e8cf169f8fb9",
};

/* ─────────────────────────────── Types ─────────────────────────────── */

/** One tier of a segment's chain, as the filter panel needs it. */
interface Level {
  label: string;
  /** The item column's internal name — what both engines filter on. */
  column: string;
  /** True unless the Levels entry says `"permissioned": false`. ABSENT MEANS TRUE. */
  permissioned: boolean;
  /** A below-Unit tier's own term set, when it declares one. */
  termSet: string;
}

/** A segment offered in the filter panel. */
interface Segment {
  /** The `Title` of the mode row — `mode_gho`. */
  key: string;
  /** What the client sees, and what is filed in `Business Segment`. */
  label: string;
  termSetGuid: string;
  levels: Level[];
}

/** One option in a dropdown. */
interface TermOption {
  id: string;
  label: string;
}

/** A row as either read returns it. An index signature, never a typed shape — see MySubmissions. */
type RawRow = Record<string, unknown>;

/**
 * Coerce whatever SharePoint returned into text.
 *
 * Managed metadata arrives as an OBJECT, and `(v ?? "").trim()` on one threw
 * "(intermediate value).trim is not a function" and killed the whole My Submissions page with an
 * error naming no field. A row interface typed `string` was the real culprit there: it made
 * TypeScript vouch for something SharePoint does not guarantee.
 */
function textOf(v: unknown): string {
  if (v === undefined || v === null) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  const o = v as Record<string, unknown>;
  if (typeof o.Label === "string") return o.Label.trim();
  if (typeof o.Title === "string") return o.Title.trim();
  return "";
}

/** `DD/MMM/YYYY` — the agreed client format. Display only; sort and filter use the stored value. */
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function formatDate(iso: string): string {
  const raw = (iso ?? "").trim();
  if (raw.length === 0) return "";
  const d = new Date(raw);
  if (isNaN(d.getTime())) return raw;
  const day = `0${d.getDate()}`.slice(-2);
  return `${day}/${MONTHS[d.getMonth()]}/${d.getFullYear()}`;
}

/** The folder trail of a server-relative path, library segment and file name removed. */
function trailOf(path: string): string {
  return (path ?? "")
    .split("/")
    .filter((p) => p.length > 0)
    .slice(0, -1)
    .join(" › ");
}

/* ─────────────────────────────── Styles ─────────────────────────────── */

const s: Record<string, React.CSSProperties> = {
  root: { fontFamily: '"Segoe UI", system-ui, sans-serif', color: "#242424" },

  /* ---- The hero (client design, 2026-08-30) --------------------------------
     ⚠ Every key used anywhere in this file MUST exist in this object: it is a
     `Record<string, CSSProperties>`, so a key that does not yields `undefined` and the element
     renders UNSTYLED with a completely green build. Four screens in this project have been caught
     by that. */
  hero: {
    position: "relative", padding: "48px 20px 28px", textAlign: "center",
    /* ⚠⚠ NO FULL-BLEED. This was `width: 100vw` with `marginLeft/Right: calc(50% - 50vw)` so the
       banner could escape its column and reach the page edges. It did - and it also pushed the
       heading and search bar visibly LEFT of the banner's centre, on every screen size and in both
       view and edit mode.
       ⚠ The first explanation offered for that was the left navigation's width halved, which the
       arithmetic happened to fit on one screenshot. It was WRONG: the offset persists in edit mode,
       where there is no left nav. Fitting a formula to one measurement is not a diagnosis.
       Removing the bleed removes the whole class of problem instead of compensating for it: the hero
       fills its own column, and `heroInner`'s `margin: 0 auto` then centres against something real.
       ⚠ THE BANNER'S WIDTH IS NOW THE SECTION'S. It reaches the page edges only in a FULL-WIDTH
       section (Edit page -> section settings -> Full width). In a normal one-column section it picks
       up the page's ~30px padding, which is correct rather than broken - but it is a page setting, so
       if the banner ever stops touching the edges, check the section before the code. */
    width: "100%",
    backgroundSize: "cover", backgroundPosition: "center", backgroundRepeat: "no-repeat",
    marginBottom: 20, overflow: "hidden",
    /* Client's own values, arrived at in DevTools on 2026-09-03.
       ⚠ `minHeight` GIVES THE PHOTO ROOM TO BE SEEN. The hero's height was whatever its content
       needed, so with the filter panel COLLAPSED the banner shrank to about the height of the title
       and the search bar — and `backgroundSize: cover` then cropped the image to a thin strip, which
       is not what a key visual is for. A floor rather than a fixed height, so expanding Advanced
       Filters still grows it instead of overflowing.
       ⚠ AND THE FLEX CENTRING IS WHAT MAKES THE FLOOR USABLE: with a minimum height the content
       would otherwise sit at the top against the padding, leaving the gap below it. `alignItems`
       centres it vertically in whatever height the banner actually has. It is `heroInner`'s
       `width: 100%` that keeps the content full width once this becomes a flex container — a flex
       item shrinks to its content by default, so removing one without the other narrows the search
       bar to the width of the words in it. */
    minHeight: 275, display: "flex", alignItems: "center",
  },
  /* ⚠ THE DARKENING OVERLAY IS GONE (client, 2026-09-01): `rgba(6, 42, 26, .55)` over the banner.
     It existed so white text stayed legible whatever photo was chosen — a `text-shadow` fails against
     a bright sky where an overlay does not. The client's own banner is dark enough without it, and
     the veil was visibly washing the photo out.

     ⚠ SO LEGIBILITY NOW DEPENDS ON THE IMAGE. A pale banner would leave white text unreadable, with
     nothing in the code to prevent it. If that ever happens, this is what to put back — do not reach
     for a text-shadow instead. */
  /* ⚠ `maxWidth` WAS 720, AND THAT IS WHY THE FILTER GRID COULD NEVER SHOW 4 COLUMNS (client,
     2026-09-02: "the layout should be grid 4 not grid 3" / "the box doesnt extend to the end of the
     page"). `advPanel` needs 4 x minmax(170px,1fr) + 3 x 14px gaps = 722px minimum — TWO PIXELS over
     the old cap — so `auto-fit` had no choice but to wrap to 3-then-1 every time, on every screen
     wide enough to hold 4 easily. 720 was sized for a single search bar, before Advanced Filters
     existed; it was never widened when the grid was added. Raised generously past the threshold
     rather than to the exact minimum, so normal column padding/gap rounding cannot reopen the same
     wrap on a slightly different width. This also directly answers the "doesn't extend to the end of
     the page" report — `s.hero` itself already spans the full section width; only this inner wrapper
     was capped, and NOT by anything SharePoint controls. */
  /* ⚠ `width: "100%"` IS LOAD-BEARING, NOT DECORATION — see `hero`'s `display: flex` above. As a
     flex item this div sizes to its content without it, so the title and the search bar collapse to
     their text width and the banner reads as broken. `maxWidth` still caps it; the two work
     together. */
  heroInner: { position: "relative", zIndex: 1, maxWidth: 1100, width: "100%", margin: "0 auto" },
  heroTitle: {
    color: "#fff", fontSize: 26, fontWeight: 700, letterSpacing: 3, textTransform: "uppercase",
    margin: "0 0 22px", lineHeight: 1.3,
  },
  /* The pill. `overflow: hidden` is what clips the button's square corners to the rounded end. */
  heroBar: {
    display: "flex", alignItems: "center", background: "#fff", borderRadius: 6,
    overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,.18)",
  },
  heroIcon: { flexShrink: 0, padding: "0 10px 0 14px", color: "#8a8886", display: "flex", alignItems: "center" },
  heroInput: {
    flex: "1 1 auto", minWidth: 0, height: 44, border: "none", outline: "none",
    fontSize: 14, color: "#242424", background: "transparent",
  },
  /* 36 high inside a 44 bar, with its own radius and right margin, so it reads as a button SITTING
     IN the field rather than as the field's end cap — the client's spec, 2026-09-01. */
  heroBtn: {
    flexShrink: 0, height: 36, padding: "0 26px", border: "none", borderRadius: 3,
    marginRight: 10, background: "#0f6c3f",
    color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer",
  },
  /* On the dark banner, so white — an amber or red would vanish against the green. */
  heroWarn: {
    color: "#ffd9a0", fontSize: 13, fontWeight: 600, margin: "10px 0 0", textAlign: "left",
  },
  /* A button, not a link: it toggles state on this page rather than navigating. Styled flat so it
     reads as the disclosure the design shows. */
  /* ⚠ LEFT, not centred. `hero` sets `textAlign: center` for the title and the search bar, which
     the toggle inherited — the client's design has it under the LEFT edge of the search box, where
     it reads as belonging to that control rather than floating beneath it. A wrapper rather than
     `alignSelf`, because the parent is not a flex container. */
  advToggleWrap: { textAlign: "left" },
  advToggle: {
    display: "inline-flex", alignItems: "center", gap: 6, marginTop: 12, padding: 0,
    background: "none", border: "none", color: "#fff", fontSize: 13, fontWeight: 600,
    cursor: "pointer", fontFamily: "inherit",
  },
  advPanel: {
    display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))",
    gap: 14, marginTop: 16, textAlign: "left",
  },
  advLabel: { display: "block", color: "#fff", fontSize: 12, fontWeight: 600, marginBottom: 6 },
  advSelect: {
    width: "100%", height: 38, padding: "0 10px", fontSize: 13, borderRadius: 6,
    border: "none", background: "#fff", color: "#242424", boxSizing: "border-box",
  },
  advActions: { display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 14, marginTop: 16 },
  advReset: {
    background: "none", border: "none", color: "#fff", fontSize: 13, cursor: "pointer",
    padding: 0, fontFamily: "inherit",
  },
  advApply: {
    height: 36, padding: "0 20px", border: "none", borderRadius: 6, background: "#0f6c3f",
    color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer",
  },
  h2: { fontSize: 28, fontWeight: 700, color: "#1b1b1b", margin: "0 0 4px" },
  sub: { fontSize: 13, color: "#616161", margin: "0 0 16px" },
  bar: { display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" },
  box: {
    flex: "1 1 260px", minWidth: 0, height: 36, padding: "0 12px", fontSize: 14,
    border: "1px solid #d1d1d1", borderRadius: 4, boxSizing: "border-box",
  },
  btn: {
    height: 36, padding: "0 18px", fontSize: 14, fontWeight: 600, border: "none",
    borderRadius: 4, background: "#107c10", color: "#fff", cursor: "pointer",
  },
  btnGhost: {
    height: 36, padding: "0 14px", fontSize: 14, border: "1px solid #d1d1d1",
    borderRadius: 4, background: "#fff", color: "#242424", cursor: "pointer",
  },
  filters: {
    display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
    gap: 10, padding: 14, background: "#faf9f8", border: "1px solid #edebe9",
    borderRadius: 6, marginBottom: 14,
  },
  field: { display: "flex", flexDirection: "column", gap: 4, minWidth: 0 },
  /* ⚠ THESE TWO NOW SIT ON THE DARK HERO, so the label is WHITE and the control has no border of
     its own — a grey 1px border against a photograph reads as a rendering fault. Used only inside
     the Advanced Filters panel (every occurrence is between the panel's open and close tags); if
     either is ever reused on a light background it needs its own key rather than an override. */
  label: { fontSize: 12, fontWeight: 600, color: "#fff" },
  input: {
    height: 38, padding: "0 10px", fontSize: 13, border: "none",
    borderRadius: 6, background: "#fff", boxSizing: "border-box", maxWidth: "100%", width: "100%",
  },
  row: {
    display: "flex", gap: 12, alignItems: "flex-start", padding: "12px 10px",
    borderBottom: "1px solid #f0f0f0", cursor: "pointer", textAlign: "left",
    background: "none", border: "none", width: "100%", font: "inherit",
  },
  name: { fontSize: 14, fontWeight: 600, color: "#0f548c", wordBreak: "break-word" },
  meta: { fontSize: 12, color: "#616161", marginTop: 2, wordBreak: "break-word" },
  chip: {
    display: "inline-block", fontSize: 11, fontWeight: 600, padding: "1px 7px",
    borderRadius: 10, marginLeft: 8, verticalAlign: "middle",
  },
  chipHc: { background: "#fde7e9", color: "#a4262c" },
  chipPending: { background: "#fff4ce", color: "#8a6100" },
  note: { fontSize: 13, padding: "10px 12px", borderRadius: 4, marginBottom: 12 },
  /* One palette for every banner — see shared/noticeStyles.ts. */
  noteWarn: { ...NOTICE_ATTENTION },
  empty: { padding: "28px 12px", textAlign: "center", color: "#616161", fontSize: 14 },
  detailWrap: { display: "flex", gap: 20, flexWrap: "wrap", alignItems: "flex-start" },
  preview: { flex: "1 1 420px", minWidth: 0, border: "1px solid #edebe9", borderRadius: 6 },
  panel: { flex: "0 1 340px", minWidth: 260 },
  dRow: { display: "flex", gap: 10, padding: "6px 0", borderBottom: "1px solid #f5f5f5", fontSize: 13 },
  dLabel: { flex: "0 0 130px", color: "#616161" },
  dValue: { flex: "1 1 auto", minWidth: 0, wordBreak: "break-word" },
};

/* ─────────────────────────────── Component ─────────────────────────────── */

export default function DocumentSearch(
  { context, pageSize, heroImageUrl, heroTitle }: IDocumentSearchProps,
): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  // Origin with no /sites/… — a server-relative path already carries the site, so previewTarget
  // needs the bare host to build an absolute file URL.
  const tenantRoot = siteUrl.replace(/^(https?:\/\/[^/]+).*$/, "$1");
  const sp = context.spHttpClient;

  const [criteria, setCriteria] = useState<SearchCriteria>(emptyCriteria());
  const [segments, setSegments] = useState<Segment[]>([]);
  const [fixedOptions, setFixedOptions] = useState<Record<string, TermOption[]>>({});

  const [results, setResults] = useState<LibraryResult[]>([]);
  const [searched, setSearched] = useState(false);
  const [busy, setBusy] = useState(false);
  /* Pressing Search with nothing filled in used to be prevented by DISABLING the button. The client
     asked for the opposite (2026-09-01): keep the button solid green and say what is wrong when it is
     pressed. A disabled control explains nothing — somebody who has not typed yet sees a dimmed
     button and no reason for it. Cleared the moment any criterion changes, so the message cannot
     outlive the state that caused it. */
  const [emptySearch, setEmptySearch] = useState(false);
  /* The filter panel is CLOSED on arrival (client design, 2026-08-30). Most searches are a few
     typed words, and a wall of dropdowns in front of that reads as a form to fill in rather than
     a box to type in. */
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [shown, setShown] = useState(pageSize);
  const [open, setOpen] = useState<SearchHit | undefined>(undefined);
  const [fieldText, setFieldText] = useState<Record<string, string> | undefined>(undefined);
  /** The Documents library's real URL segment — `Shared Documents`, not `Documents` (gotcha #12).
      ⚠ SEEDED WITH THE SEGMENT, NOT THE TITLE. It held `DOCUMENTS_LIBRARY` until 2026-08-28, which
      was the wrong string for a URL even when the two matched — and once the client retitled the
      library to `Restricted & Confidential Document` it was not even close. Overwritten by the boot
      fetch below either way; this is what the page uses until that lands. */
  const [docsSegment, setDocsSegment] = useState(DOCUMENTS_URL_SEGMENT);
  const [configWarning, setConfigWarning] = useState("");

  /* A run counter, so a slow earlier search cannot overwrite a faster later one. A ref rather than
     state: it must be readable inside an async closure at its CURRENT value, and state read there
     is the render-time value — the same stale-closure trap that recorded zero counts during
     reconciliation. */
  const runId = useRef(0);

  const jsonGet = async (
    url: string,
    headers?: Record<string, string>,
  ): Promise<{ ok: boolean; status: number; body: RawRow }> => {
    try {
      const res: SPHttpClientResponse = await sp.get(url, SPHttpClient.configurations.v1, {
        headers: headers ?? { Accept: "application/json;odata=nometadata" },
      });
      if (!res.ok) return { ok: false, status: res.status, body: {} };
      return { ok: true, status: res.status, body: (await res.json()) as RawRow };
    } catch {
      // 0 = the request never completed. Kept distinct from a real status on screen, because a
      // network failure and a 403 share nothing except the word "failed".
      return { ok: false, status: 0, body: {} };
    }
  };

  /**
   * ⚠ `_api/search/query` NEEDS `odata-version: ""`, AND THIS IS THE SIXTH TIME THIS PROJECT HAS
   * HIT THE SAME MECHANISM (spAuditLog.ts's `WRITE_HEADERS`, FileTypeSettings.tsx, FolderManager.tsx,
   * Requests.tsx, MySubmissions.tsx, UploadPauseToggle.tsx — all writes). This is the first GET to
   * need it: SPFx's `SPHttpClient` injects `odata-version: 4.0` on EVERY request regardless of what
   * Accept says, and the legacy Search REST surface — older than the JSON-light conventions the rest
   * of this codebase's `/_api/web/lists` calls tolerate — throws a bare, bodyless HTTP 500 under that
   * combination. Found live 2026-09-02: every free-text search against `Documents` failed with
   * "One library could not be searched: <title> (HTTP 500)", on a query as plain as "ocbc".
   * `/_api/web/lists` calls elsewhere in THIS SAME FILE are untouched — the injected header is
   * harmless there, so only this one call gets the override, not the shared default.
   */
  const SEARCH_HEADERS = {
    Accept: "application/json;odata=nometadata",
    "odata-version": "",
  };

  const readTerms = async (url: string): Promise<TermOption[]> => {
    const r = await jsonGet(url);
    const out: TermOption[] = [];
    for (const t of (r.body.value as RawRow[]) ?? []) {
      const labels = (t.labels as RawRow[]) ?? [];
      const name = labels.length > 0 ? textOf(labels[0].name) : "";
      const tid = textOf(t.id);
      if (name.length > 0 && tid.length > 0) out.push({ id: tid, label: name });
    }
    return out.sort((a, b) => a.label.localeCompare(b.label));
  };

  /* ── The banner, when it does not appear ────────────────────────────────────
   *
   * ⚠ A BANNER THAT FAILS TO LOAD IS SILENT BY DESIGN, AND THAT SILENCE HAS NOW COST TWO ROUNDS OF
   * GUESSING. The gradient sits behind the image precisely so a viewer never meets a broken page —
   * right for a viewer, and it means nobody can tell a WRONG URL from a PERMISSION problem, which
   * need opposite fixes:
   *   404 → the file is not at that address (never uploaded, renamed, or the property-pane override
   *         points somewhere else). Fix the file or the override.
   *   403 → the file is there and this account cannot read it.
   * ⚠ SharePoint security-trims to 404 as readily as it answers 403, so a 404 does NOT rule out
   * permissions — it only rules IN "this account cannot resolve it", which is why the message names
   * both and does not pick one.
   *
   * ADMINS ONLY. An uploader can do nothing with this and it would read as the page being broken —
   * the same rule as the unready-segment label, which is shown to the admin building the segment and
   * to nobody else. The probe is skipped entirely for everyone else, so it costs them nothing.
   *
   * ⚠⚠ AND THAT GATE MAKES IT BLIND TO THE CAUSE THAT ACTUALLY BITES. Confirmed twice on site
   * (2026-09-02 and 2026-09-03): **Site Assets does not grant Read to `CRS_SITE_MEMBERS`**, so an
   * uploader or approver gets SharePoint's own AccessDenied page for this URL while an ADMINISTRATOR
   * loads it perfectly. An admin therefore never sees a 403 — so this note catches a wrong or
   * missing URL, which affects everybody equally, and **cannot catch the permission case at all**.
   * Do not read its silence as "the banner is fine for everyone".
   *
   * The real check is "can the site-members group read Site Assets", which is a question about a
   * GROUP rather than about this session, and it belongs in **reconciliation** beside the existing
   * Site Pages assertion (`⚠ NOBODY CAN OPEN THE SITE HOME PAGE`) — same shape, same fail-open rule,
   * and it would then cover SDG on its next run instead of waiting for somebody to notice a green
   * rectangle. NOT BUILT.
   */
  const [bannerNote, setBannerNote] = useState<string | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    if (heroImageUrl.length === 0) return undefined;
    const check = async (): Promise<void> => {
      // The admin check FIRST, so a non-admin never issues the image request at all.
      const admin = await isSystemAdmin(sp, siteUrl).catch(() => false);
      if (cancelled || !admin) return;
      try {
        const res: SPHttpClientResponse = await sp.get(
          heroImageUrl,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "*/*" } },
        );
        if (cancelled) return;
        // ⚠ The STATUS only — the body is an image and is never read. `ok` covers 200 and 304.
        if (!res.ok) setBannerNote(String(res.status));
      } catch {
        /* A network-level failure says nothing about the file — a browser extension can block a
           request before it leaves (the Brave Shields case of 2026-08-24). Report nothing rather
           than blaming the banner. */
      }
    };
    check().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [heroImageUrl, sp, siteUrl]);

  /* ── Mount: names, segments, the three fixed term sets ─────────────────── */

  useEffect(() => {
    let cancelled = false;

    const boot = async (): Promise<void> => {
      await primeNames(sp, siteUrl).catch(() => undefined);
      if (cancelled) return;

      /* The Documents URL segment, resolved not assumed. Its title is `Documents` and its URL is
         `/Shared Documents/` — hardcoding one string for both put "Shared Documents" at the head of
         every folder trail in My Submissions. */
      const docs = await jsonGet(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(documentsLibraryTitle())}')` +
          `/RootFolder?$select=ServerRelativeUrl`,
      );
      if (!cancelled && docs.ok) {
        const parts = textOf(docs.body.ServerRelativeUrl).split("/").filter((p) => p.length > 0);
        const seg = parts.length > 0 ? parts[parts.length - 1] : "";
        if (seg.length > 0) setDocsSegment(seg);
      }

      const cfg = cachedListTitle(LIST_SUFFIX.config);

      /* Segments and their tier chains, from the DMS Config mode rows — the same source the upload
         form cascades over. NOT a written-out Department/Unit pair: the admin NAMES the tiers at
         onboarding, so a hardcoded list is only ever right for the segments it was written for, and
         a filter that silently matches nothing reads as "there are no such documents". */
      const modes = await jsonGet(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cfg)}')/items` +
          `?$select=Title,ModeLabel,TermSetGuid,Levels,SortOrder&$filter=ConfigType eq 'mode'&$top=100`,
      );
      if (cancelled) return;
      if (!modes.ok) {
        // Named, not silent. Without segments the tier filters cannot be offered at all, and an
        // admin needs to know the cause is a config read rather than an empty repository.
        setConfigWarning(
          "Segment and tier filters are unavailable — the configuration list could not be read " +
            `(${modes.status === 0 ? "the request did not complete" : `HTTP ${modes.status}`}). ` +
            "Free-text and the other filters still work.",
        );
      } else {
        const list: Segment[] = [];
        for (const row of (modes.body.value as RawRow[]) ?? []) {
          const label = textOf(row.ModeLabel) || textOf(row.Title);
          if (label.length === 0) continue;
          const levels: Level[] = [];
          try {
            const parsed = (JSON.parse(textOf(row.Levels) || "[]") as RawRow[]) ?? [];
            for (const lv of parsed) {
              const column = textOf(lv.column);
              if (column.length === 0) continue;
              levels.push({
                label: textOf(lv.label) || column,
                column,
                // ABSENT MEANS TRUE, and only a literal false demotes a tier — the string "false"
                // does not. The same rule the upload form's chain follows.
                permissioned: lv.permissioned !== false,
                termSet: textOf(lv.termSet),
              });
            }
          } catch {
            // A malformed chain costs this one segment its tier filters, never the whole panel.
            levels.length = 0;
          }
          list.push({ key: textOf(row.Title), label, termSetGuid: textOf(row.TermSetGuid), levels });
        }
        if (!cancelled) setSegments(list);
      }

      /* The three fixed dropdowns. Ids come from DMS Config where present — they are per site since
         the in-site term store pivot, so the constants above are a last resort. */
      const settings = await jsonGet(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(cfg)}')/items` +
          `?$select=Title,SettingValue&$filter=startswith(Title,'termSet_')&$top=50`,
      );
      const configured: Record<string, string> = {};
      for (const row of (settings.body.value as RawRow[]) ?? []) {
        configured[textOf(row.Title)] = textOf(row.SettingValue);
      }
      const setFor = (key: string, fallback: string): string => {
        const v = configured[`termSet_${key}`];
        return v !== undefined && v.length > 0 ? v : fallback;
      };
      const base = `${siteUrl}/_api/v2.1/termStore/sets`;
      const loaded = await Promise.all([
        readTerms(`${base}/${setFor("documentType", FALLBACK_TERM_SETS.documentType)}/children`),
        readTerms(`${base}/${setFor("yearPeriod", FALLBACK_TERM_SETS.year)}/children`),
        readTerms(`${base}/${setFor("confidentiality", FALLBACK_TERM_SETS.confidentiality)}/children`),
      ]);
      if (!cancelled) {
        setFixedOptions({ documentType: loaded[0], year: loaded[1], confidentiality: loaded[2] });
      }
    };

    boot().catch(() => undefined);
    return () => {
      cancelled = true;
    };
    // Mount only: neither the site nor the client changes under a live page.
  }, []);

  /* ── The business segment filter ──────────────────────────────────────────────────── */



  const chooseSegment = (label: string): void => {
    // `tiers: []` still cleared: a segment change must not leave a tier value from the previous one
    // in the criteria, even though nothing sets one any more.
    setCriteria((c) => ({ ...c, segment: label, tiers: [] }));
  };

  /* ⚠ `chooseTier` AND THE TIER CASCADE WERE REMOVED with the tier filters (client, 2026-08-30).
     `loadTierOptions` goes with it — nothing else called it. `SearchCriteria.tiers` REMAINS, and
     `buildKql` still honours it, so restoring the cascade is UI work rather than a rewrite.

     What the removed code did, if it is ever wanted back: choosing a tier cleared every tier BELOW
     it, because a stale Unit under a newly chosen Department queries a combination that cannot
     exist and returns nothing — which reads as *"there are no documents here"* rather than as a
     stale filter. Anyone re-adding this must re-add that clearing too. */

  /* ── Reads ─────────────────────────────────────────────────────────────── */

  const libraryLabel = (lib: SearchLibrary): string => {
    const hc = cachedHcLibraries();
    if (lib === "Documents") return documentsLibraryTitle();
    if (lib === "Staging") return libraryTitle();
    if (lib === "DocumentsHC") return hc ? hc.documents.title : "HC Documents";
    if (lib === "StagingHC") return hc ? hc.approval.title : "HC Approval Document";
    return lib;
  };

  /** Which approved-side library a crawled result came from, by its path. */
  const libraryOfPath = (path: string): SearchLibrary => {
    const hc = cachedHcLibraries();
    const p = (path ?? "").toLowerCase();
    if (hc && p.indexOf(`/${hc.documents.urlSegment.toLowerCase()}/`) !== -1) return "DocumentsHC";
    return "Documents";
  };

  /**
   * The Search half — ONE request covering both approved-side libraries.
   *
   * One rather than two because Search trims by ACL, so an uncleared user simply gets no HC rows.
   * There is no 403 to interpret here, which is why only the live reads carry a `refused` outcome.
   */
  const runSearch = async (c: SearchCriteria): Promise<LibraryResult[]> => {
    const hc = cachedHcLibraries();
    const segs = [docsSegment];
    if (hc) segs.push(hc.documents.urlSegment);
    /* ⚠ THE ARCHIVE IS SEARCHABLE (2026-08-22), and leaving it out would be the worse default. A
       document that passed seven years would become unfindable — and search is the ONLY route to it
       once it is out of the folders people browse. Nobody would report that as a bug; they would
       conclude the document had been deleted.

       Safe by the same argument as the HC library beside it: Search applies ACLs at query time, so a
       reader gets archive rows only where they already hold Read. This web part enforces no
       permissions and must never be changed to — a bug here can only ever return FEWER rows. */
    const arc = cachedArchiveLibraries();
    if (arc) {
      segs.push(arc.normal.urlSegment);
      if (arc.hc) segs.push(arc.hc.urlSegment);
    }
    const query = buildKql(c, kqlPathScope(siteUrl, segs));
    // Blank means "do not run" — never "match everything". An unscoped KQL query searches the whole
    // tenant, which would show documents from sites this system has nothing to do with.
    if (query.length === 0) return [];

    const props = "Path,Filename,Title,LastModifiedTime,Author,Size,UniqueId,ListItemID";
    const r = await jsonGet(
      `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(query)}'` +
        `&rowlimit=200&trimduplicates=false&selectproperties='${encodeURIComponent(props)}'`,
      SEARCH_HEADERS,
    );
    if (!r.ok) return [{ library: "Documents", outcome: "failed", status: r.status, hits: [] }];

    const primary = (r.body.PrimaryQueryResult as RawRow) ?? {};
    const relevant = (primary.RelevantResults as RawRow) ?? {};
    const table = (relevant.Table as RawRow) ?? {};
    const hits: SearchHit[] = [];
    for (const row of (table.Rows as RawRow[]) ?? []) {
      const map: Record<string, string> = {};
      for (const cell of (row.Cells as RawRow[]) ?? []) map[textOf(cell.Key)] = textOf(cell.Value);
      const path = map.Path ?? "";
      if (path.length === 0) continue;
      const name = map.Filename || map.Title || "";
      hits.push({
        // Search returns it braced and the list read does not — normalised so the two halves dedupe
        // against each other rather than showing the same document twice.
        uniqueId: (map.UniqueId ?? "").replace(/[{}]/g, ""),
        itemId: Number(map.ListItemID ?? "0") || 0,
        library: libraryOfPath(path),
        name: name.length > 0 ? name : path.split("/").filter((p) => p.length > 0).pop() ?? "",
        // Reduced to server-relative, so the trail and the preview agree with the live reads.
        path: path.replace(/^https?:\/\/[^/]+/, ""),
        modified: map.LastModifiedTime ?? "",
        author: map.Author ?? "",
        size: map.Size ?? "",
      });
    }
    return [{ library: "Documents", outcome: "ok", hits }];
  };

  /** One live list read — used for the approval libraries and for the recency top-up. */
  const runListRead = async (lib: SearchLibrary, filter: string): Promise<LibraryResult> => {
    const base = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(lib))}')/items`;
    const core = `Id,FileLeafRef,FileRef,Modified,Author/Title,File/Length,File/UniqueId`;
    /* ⚠ THE THREE TAXONOMY COLUMNS ARE SELECTED, NEVER FILTERED. `$filter` cannot touch them — `eq`
       answers 400 and the `TaxCatchAllLabel` workaround answers 500, both failing the WHOLE request
       (see `METADATA_FILTER_FIELDS` in shared/documentSearch.ts for the two live failures that
       establish this). Selecting them is a different operation and is well supported, so the labels
       travel with the row and the narrowing happens below. */
    const meta = [
      METADATA_FILTER_FIELDS.documentType,
      METADATA_FILTER_FIELDS.year,
      METADATA_FILTER_FIELDS.confidentiality,
    ].join(",");
    const url = (select: string): string =>
      `${base}?$select=${select}&$expand=Author,File` +
      `&$filter=${encodeURIComponent(filter)}&$orderby=Modified desc&$top=200`;

    let r = await jsonGet(url(`${core},${meta}`));
    let narrowed = true;
    /* ⚠ RETRY WITHOUT THEM ON 400 ONLY — one unknown field name fails the whole `$select` (gotcha
       #11), and a library provisioned before those columns existed must still be searchable. 400
       only: a 403 is a permission answer and a 404 is the library, and retrying either asks the same
       unanswerable question twice. The same three-rung habit as the record read on My Submissions. */
    if (!r.ok && r.status === 400) {
      r = await jsonGet(url(core));
      narrowed = false;
    }
    if (!r.ok) {
      /* A 403 is an uncleared user meeting an HC library, and that is the CORRECT answer to their
         query — not an error. Counting it as one would put a permanent warning in front of everyone
         who is not HC-cleared, and train them to ignore the banner on the day it means something. */
      const refused = isHcLibrary(lib) && (r.status === 403 || r.status === 401);
      /* ⚠ THE SERVER'S OWN MESSAGE, TO THE CONSOLE. "HTTP 500" alone is what cost two blind fixes to
         this filter: SharePoint names the offending field and the reason in the BODY, and nothing was
         reading it. Console rather than the banner — it is developer detail, and the banner already
         names the library and the status. */
      if (!refused) {
        console.error(`CRS Search: ${libApiTitle(lib)} read failed (HTTP ${r.status})`, r.body);
      }
      return { library: lib, outcome: refused ? "refused" : "failed", status: r.status, hits: [] };
    }
    const hits: SearchHit[] = [];
    for (const row of (r.body.value as RawRow[]) ?? []) {
      const file = (row.File as RawRow) ?? {};
      const author = (row.Author as RawRow) ?? {};
      /* The narrowing REST could not do. `textOf` already handles a taxonomy field's
         `{Label, TermGuid, WssId}` shape as well as a bare string. */
      if (narrowed && !metadataFilterMatches(criteria, {
        documentType: textOf(row[METADATA_FILTER_FIELDS.documentType]),
        year: textOf(row[METADATA_FILTER_FIELDS.year]),
        confidentiality: textOf(row[METADATA_FILTER_FIELDS.confidentiality]),
      })) {
        continue;
      }
      hits.push({
        uniqueId: textOf(file.UniqueId).replace(/[{}]/g, ""),
        itemId: Number(textOf(row.Id)) || 0,
        library: lib,
        name: textOf(row.FileLeafRef),
        path: textOf(row.FileRef),
        modified: textOf(row.Modified),
        author: textOf(author.Title),
        size: textOf(file.Length),
      });
    }
    /* `unnarrowed` only when there was actually something to narrow BY — a search with none of those
       three filters set is fully answered by the `$filter`, and flagging it would warn about a
       limitation that did not apply. */
    return {
      library: lib,
      outcome: "ok",
      hits,
      unnarrowed: !narrowed && hasMetadataFilter(criteria) ? true : undefined,
    };
  };

  const search = async (): Promise<void> => {
    if (!hasCriteria(criteria)) return;
    runId.current += 1;
    const mine = runId.current;
    setBusy(true);
    setOpen(undefined);
    setShown(pageSize);

    const hc = cachedHcLibraries();

    /* ⚠ RESOLVED PER LIBRARY, NOT ONCE (2026-09-04). One filter string is sent to up to four
       libraries, and they are provisioned independently — a site can carry `Keyword` on the normal
       pair and not on the HC pair, or vice versa. Since a `$filter` naming a column the library does
       not have returns 400 and fails the WHOLE read, a single shared flag would either lose the
       keyword search everywhere or break search on whichever library is behind.

       `libraryHasColumns` caches per library, so this is at most one extra read per library per
       page, and it answers FALSE on any failure — the safe direction, because the cost of a wrong
       `true` is a dead search and the cost of a wrong `false` is one field not being matched. */
    const hasKw = async (t: "Staging" | "Documents" | "StagingHC" | "DocumentsHC"): Promise<boolean> => {
      try {
        return await libraryHasColumns(
          context.spHttpClient, siteUrl, libApiTitle(t), [KEYWORD_COLUMN],
        );
      } catch {
        return false;
      }
    };
    const kwStaging = await hasKw("Staging");
    const kwDocuments = await hasKw("Documents");
    const kwStagingHc = hc ? await hasKw("StagingHC") : false;
    const kwDocumentsHc = hc ? await hasKw("DocumentsHC") : false;

    const listFilter = buildListFilter(criteria, kwStaging);
    const listFilterHc = buildListFilter(criteria, kwStagingHc);
    // Generous against a crawl measured in minutes: too tight leaves exactly the invisible gap this
    // exists to close, and the cost of being generous is a few extra rows to dedupe.
    const cutoff = recencyCutoff(new Date(), 24);
    const recentFilter = buildRecentFilter(criteria, cutoff, kwDocuments);
    const recentFilterHc = buildRecentFilter(criteria, cutoff, kwDocumentsHc);

    const jobs: Promise<LibraryResult[]>[] = [
      runSearch(criteria),
      runListRead("Staging", listFilter).then((x) => [x]),
      /* The gap between the engines: a file approved two minutes ago has left the approval library
         and is not yet crawled, so it is in NEITHER. `Modified` is indexable and the window keeps
         the set tiny, so this stays threshold-safe at any library size — which is precisely why the
         main read could not be done this way. */
      runListRead("Documents", recentFilter).then((x) => [x]),
    ];
    if (hc) {
      jobs.push(runListRead("StagingHC", listFilterHc).then((x) => [x]));
      jobs.push(runListRead("DocumentsHC", recentFilterHc).then((x) => [x]));
    }

    // Promise.allSettled is unavailable (the tsconfig predates it), and each job already resolves to
    // an outcome rather than rejecting.
    const settled = await Promise.all(jobs);
    if (runId.current !== mine) return; // a later search has already answered
    const flat: LibraryResult[] = [];
    for (const group of settled) for (const one of group) flat.push(one);
    setResults(flat);
    setSearched(true);
    setBusy(false);
  };

  const runSearchSafely = (): void => {
    // The guard lives here rather than in `search`, which returns silently on empty criteria and is
    // also reached from Enter. One place decides, and it is the one place that can show a message.
    if (!hasCriteria(criteria)) {
      setEmptySearch(true);
      return;
    }
    setEmptySearch(false);
    search().catch(() => setBusy(false));
  };

  const reset = (): void => {
    setCriteria(emptyCriteria());
    setResults([]);
    setSearched(false);
    setOpen(undefined);
  };

  const openRow = (h: SearchHit): void => {
    setOpen(h);
    setFieldText(undefined);
    /* FieldValuesAsText is PER ITEM, which is why the list shows no metadata at all — one request
       per row would be hundreds. It also returns LABELS, which is what makes a taxonomy value
       readable instead of the bare lookup id that reached the screen on 2026-08-14. */
    jsonGet(
      `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libApiTitle(h.library))}')` +
        `/items(${h.itemId})/FieldValuesAsText`,
    )
      .then((r) => setFieldText(r.ok ? (r.body as Record<string, string>) : {}))
      .catch(() => setFieldText({}));
  };

  /* ── Render ────────────────────────────────────────────────────────────── */

  const allHits = sortByModified(mergeHits(...results.map((r) => r.hits)));
  const state = searchState(searched, results);
  const failed = failedLibraries(results);
  /* Read straight off the results rather than recomputed: the read itself is the only thing that
     knows whether its `$select` was accepted. */
  const unnarrowedLibraries = results.filter((r) => r.unnarrowed === true).map((r) => r.library);

  if (open) {
    const preview = previewTarget(open.name, open.path, tenantRoot, siteUrl);
    const rows = buildDetailRows({
      fieldText: fieldText ?? {},
      leading: [
        { label: "Library", value: libraryLabel(open.library) },
        { label: "Location", value: trailOf(open.path) },
      ],
      trailing: [
        { label: "Uploaded by", value: open.author },
        { label: "Last updated", value: formatDate(open.modified) },
        { label: "File size", value: formatBytes(open.size) },
      ],
    });

    return (
      <div style={s.root}>
        <button
          type="button"
          style={{ ...s.btnGhost, marginBottom: 14 }}
          onClick={() => setOpen(undefined)}
        >
          ‹ Back to results
        </button>
        <h2 style={s.h2}>
          {open.name}
          {isHcLibrary(open.library) ? (
            <span style={{ ...s.chip, ...s.chipHc }}>Highly Confidential</span>
          ) : undefined}
        </h2>
        <div style={s.detailWrap}>
          <div style={s.preview}>
            {preview.kind === "none" ? (
              <div style={s.empty}>
                No preview for this file type.{" "}
                <a href={preview.fileUrl} target="_blank" rel="noreferrer">
                  Open it in SharePoint
                </a>
              </div>
            ) : preview.kind === "image" ? (
              // Fit to WIDTH and scroll — the same fix the approval page needed.
              <div style={{ maxHeight: 600, overflow: "auto" }}>
                <img src={preview.url} alt={open.name} style={{ width: "100%", display: "block" }} />
              </div>
            ) : (
              <iframe
                title={open.name}
                src={preview.url}
                style={{ width: "100%", height: 600, border: "none" }}
              />
            )}
          </div>
          <div style={s.panel}>
            {fieldText === undefined ? (
              <div style={s.meta}>Reading details…</div>
            ) : rows.length === 0 ? (
              <div style={s.meta}>No details were recorded for this document.</div>
            ) : (
              rows.map((r) => (
                <div key={`${r.label}:${r.value}`} style={s.dRow}>
                  <div style={s.dLabel}>{r.label}</div>
                  <div style={s.dValue}>{r.value}</div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={s.root}>
      {/* THE HERO (client design, 2026-08-30). The old heading, one-line explanation and flat search
          row are replaced by a full-bleed banner. What the explanation used to say — that searching
          shows only what you can already open — has NOT been dropped: it moves under the results,
          where somebody wondering why a document is missing will actually read it. */}
      <div
        style={{
          ...s.hero,
          /* ⚠⚠ TWO LAYERS IN `backgroundImage`, NEVER THE `background` SHORTHAND. The first attempt
             set `background: linear-gradient(...)` alongside the image — and the shorthand RESETS
             every `background-*` longhand, so it silently wiped the `backgroundSize: cover`,
             `backgroundPosition: center` and `backgroundRepeat` declared in `s.hero`. The photo then
             rendered at its natural size with the gradient showing down both sides, which is what
             "the background is not fully covered" was. Layers keep `cover` intact.

             The image is the FIRST layer (nearest the viewer) and the gradient sits behind it, so a
             banner that 404s degrades to the plain green rather than to nothing. */
          backgroundImage:
            heroImageUrl.length > 0
              ? `url("${heroImageUrl}"), linear-gradient(135deg, #0b4f2e 0%, #14724a 60%, #0b4f2e 100%)`
              : "linear-gradient(135deg, #0b4f2e 0%, #14724a 60%, #0b4f2e 100%)",
        }}
      >
        <div style={s.heroInner}>
          <h2 style={s.heroTitle}>{heroTitle}</h2>

          <div style={s.heroBar}>
            <span style={s.heroIcon} aria-hidden="true">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <circle cx="7" cy="7" r="4.5" stroke="currentColor" strokeWidth="1.6" />
                <path d="M10.5 10.5L14 14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
              </svg>
            </span>
            <input
              style={s.heroInput}
              placeholder="Search file..."
              aria-label="Search documents"
              value={criteria.text}
              onChange={(e) => setCriteria((c) => ({ ...c, text: e.target.value }))}
              onKeyDown={(e) => {
                if (e.key === "Enter") runSearchSafely();
              }}
            />
            <button
              type="button"
              /* ⚠ DISABLED ONLY WHILE A SEARCH IS RUNNING. Empty criteria no longer dim it — the
                 press is what surfaces the message. */
              style={{ ...s.heroBtn, ...(busy ? { opacity: 0.6, cursor: "default" } : {}) }}
              disabled={busy}
              onClick={runSearchSafely}
            >
              {busy ? "Searching…" : "Search"}
            </button>
          </div>

          {/* ⚠ DERIVED, not cleared by hand. Keying this on the flag alone meant clearing it in every
              onChange — the text box, six filter dropdowns and Reset — and the one that got missed
              would leave "type something to search for" sitting above a filled-in filter. ANY
              criterion becoming non-empty hides it, with nothing to remember. */}
          {emptySearch && !hasCriteria(criteria) ? (
            <p style={s.heroWarn}>Type something to search for, or pick a filter, then choose Search.</p>
          ) : undefined}

          <div style={s.advToggleWrap}>
          <button
            type="button"
            style={s.advToggle}
            aria-expanded={showAdvanced}
            onClick={() => setShowAdvanced((v) => !v)}
          >
            <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
              <path d="M2 3h12l-4.6 5.4V13L6.6 11.4V8.4L2 3z" stroke="currentColor" strokeWidth="1.4"
                strokeLinejoin="round" />
            </svg>
            Advanced Filters
            <svg width="12" height="12" viewBox="0 0 16 16" fill="none" aria-hidden="true"
              style={{ transform: showAdvanced ? "rotate(180deg)" : undefined }}>
              <path d="M3.5 6L8 10.5 12.5 6" stroke="currentColor" strokeWidth="1.6"
                strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </button>
          </div>

          {showAdvanced ? (
            <>
      <div style={s.advPanel}>
        <div style={s.field}>
          <label style={s.label} htmlFor="crs-doctype">Document type</label>
          <select
            id="crs-doctype"
            style={s.input}
            value={criteria.documentType}
            onChange={(e) => setCriteria((c) => ({ ...c, documentType: e.target.value }))}
          >
            <option value="">Any</option>
            {(fixedOptions.documentType ?? []).map((o) => (
              <option key={o.id} value={o.label}>{o.label}</option>
            ))}
          </select>
        </div>

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-year">Year</label>
          <select
            id="crs-year"
            style={s.input}
            value={criteria.year}
            onChange={(e) => setCriteria((c) => ({ ...c, year: e.target.value }))}
          >
            <option value="">Any</option>
            {(fixedOptions.year ?? []).map((o) => (
              <option key={o.id} value={o.label}>{o.label}</option>
            ))}
          </select>
        </div>

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-conf">Confidentiality</label>
          <select
            id="crs-conf"
            style={s.input}
            value={criteria.confidentiality}
            onChange={(e) => setCriteria((c) => ({ ...c, confidentiality: e.target.value }))}
          >
            <option value="">Any</option>
            {(fixedOptions.confidentiality ?? []).map((o) => (
              <option key={o.id} value={o.label}>{o.label}</option>
            ))}
          </select>
        </div>

        <div style={s.field}>
          <label style={s.label} htmlFor="crs-segment">Business segment</label>
          <select
            id="crs-segment"
            style={s.input}
            value={criteria.segment}
            onChange={(e) => chooseSegment(e.target.value)}
          >
            <option value="">Any</option>
            {segments.map((seg) => (
              <option key={seg.key} value={seg.label}>{seg.label}</option>
            ))}
          </select>
        </div>

        {/* ⚠ THE TIER AND DATE FILTERS WERE REMOVED HERE (client, 2026-08-30: *"just follow the
            mock, keep it simple, don't give client too many features or inputs, it will make them
            scared"*). What went: the segment's own cascading tiers (Department/Unit, Region/Estate
            — derived, never hardcoded), Document date from/to, and Uploaded from/to.

            ⚠ `SearchCriteria` STILL CARRIES ALL OF THEM, and `buildKql`/`buildRestFilter` still
            honour them, with 72 tests untouched. Only the CONTROLS are gone — the fields simply stay
            empty, which every builder already treats as "no constraint". Restoring any of them is
            re-adding a `<select>`, not re-implementing a filter. Deleting the model instead would
            have thrown away tested behaviour to satisfy a layout. */}
      </div>

              <div style={s.advActions}>
                {/* Clears the filters AND the typed text — the same `reset` the old Clear button
                    called, so there is one definition of "start again". */}
                <button type="button" style={s.advReset} onClick={reset}>Reset Filter</button>
                <button
                  type="button"
                  style={{ ...s.advApply, ...(busy ? { opacity: 0.6, cursor: "default" } : {}) }}
                  disabled={busy}
                  onClick={runSearchSafely}
                >
                  {busy ? "Searching…" : "Apply Filter"}
                </button>
              </div>
            </>
          ) : undefined}
        </div>
      </div>

      {configWarning.length > 0 ? <div style={{ ...s.note, ...s.noteWarn }}>{configWarning}</div> : undefined}

      {/* A PARTIAL failure shows its results AND says what is missing. Rendering it as a clean list
          silently understates what is there; rendering it as "no results" is how someone concludes
          a document was never filed. */}
      {failed.length > 0 ? (
        <div style={{ ...s.note, ...s.noteWarn }}>
          {failed.length === 1
            ? "One library could not be searched"
            : `${failed.length} libraries could not be searched`}
          {": "}
          {failed
            .map(
              (f) =>
                `${libraryLabel(f.library)} (${
                  f.status === 0 ? "the request did not complete" : `HTTP ${f.status}`
                })`,
            )
            .join(", ")}
          . Anything filed there is missing from these results.
        </div>
      ) : undefined}

      {/* ⚠ SAYS SO WHEN A FILTER COULD NOT BE APPLIED, rather than presenting a wider result as an
          exact one. The three taxonomy filters are narrowed on the rows, so a library whose `$select`
          of those columns was rejected returns everything matching the OTHER filters — correct rows,
          just not narrowed. Silently showing them would overstate the match; dropping the library
          would hide documents that may well qualify. Naming it is the only honest option, and it
          appears only when one of those three filters was actually set. */}
      {unnarrowedLibraries.length > 0 ? (
        <div style={{ ...s.note, ...s.noteWarn }}>
          Document type, Year and Confidentiality could not be applied to{" "}
          {unnarrowedLibraries.map((l) => libraryLabel(l)).join(", ")}, so results from{" "}
          {unnarrowedLibraries.length === 1 ? "it" : "them"} may include documents that do not match
          those three filters. Every other filter was applied everywhere.
        </div>
      ) : undefined}

      {/* ⚠ THE ACCESS SENTENCE MOVED HERE FROM THE OLD SUBTITLE (2026-08-30) rather than being
          dropped with it. The hero has no room for prose, but the fact still matters and belongs
          where somebody wondering "why is my colleague's document not showing" will read it —
          beside the results, not above the box. Always rendered, in every state: it explains an
          EMPTY result as readily as a short one. */}
      <p style={{ ...s.sub, marginTop: 14, marginBottom: 10 }}>
        You will only ever see documents you are already allowed to open — searching does not give
        access to anything new.
      </p>

      {/* ⚠ ADMIN-ONLY, AND ONLY WHEN THE BANNER ACTUALLY FAILED. `bannerNote` is set for nobody else
          (the probe is skipped for them), so an uploader never meets a message about a file they
          cannot fix. It names the URL because the property-pane override can point anywhere, and
          both causes because SharePoint's 404 and 403 are interchangeable here. */}
      {bannerNote !== undefined && (
        <p style={{ ...s.sub, marginTop: 0, marginBottom: 10, color: "#8a4b00" }}>
          <strong>Banner image not loading (HTTP {bannerNote}).</strong>{" "}
          The page fell back to the plain green. The address it tried is{" "}
          <code>{heroImageUrl}</code> — either nothing is there, or this account cannot read it.
          Open that address in a new tab: if it downloads or displays, the file is fine and{" "}
          <strong>Site Assets</strong> needs Read for the site members group; if it fails, upload the
          image to Site Assets under that exact name, or set a different address in this web part&rsquo;s
          property pane. Only administrators see this line.
        </p>
      )}

      {state === "idle" ? (
        <div style={s.empty}>Type something above, or pick a filter, then choose Search.</div>
      ) : state === "error" ? (
        <div style={s.empty}>
          The search could not be completed, so this is <strong>not</strong> a statement that nothing
          matched. See the message above, then try again.
        </div>
      ) : state === "empty" ? (
        /* ⚠ THE ATTENTION BOX, NOT PLAIN TEXT (client, 2026-09-04: *"when something is not found can
            you use the red box design like the one is using in Retire a segment?"*). Same
            `NOTICE_ATTENTION` palette as every other banner in the product, so there is one
            definition of what an attention box looks like.

            ⚠ THE `error` STATE ABOVE DELIBERATELY STAYS PLAIN. These two are the distinction this
            three-state render exists to draw — *"nothing matched"* is an ANSWER, while *"the search
            could not be completed"* is the absence of one, and its own text says outright that it is
            not a statement about the documents. Giving both the same box puts them back into one
            visual register, which is what the split was for. If the error state should be red too,
            that needs its own decision. */
        <div style={{ ...s.noteWarn, padding: "12px 14px", borderRadius: 8 }}>
          No documents matched. Note that a document approved in the last few minutes may not be
          findable yet.
        </div>
      ) : (
        <div>
          <div style={{ ...s.meta, marginBottom: 6 }}>
            {allHits.length === 1 ? "1 document" : `${allHits.length} documents`}
          </div>
          {allHits.slice(0, shown).map((h) => (
            <button
              type="button"
              key={`${h.library}#${h.itemId}#${h.uniqueId}`}
              style={s.row}
              onClick={() => openRow(h)}
            >
              <div style={{ minWidth: 0, flex: "1 1 auto" }}>
                <div style={s.name}>
                  {h.name}
                  {isHcLibrary(h.library) ? <span style={{ ...s.chip, ...s.chipHc }}>HC</span> : undefined}
                  {h.library === "Staging" || h.library === "StagingHC" ? (
                    <span style={{ ...s.chip, ...s.chipPending }}>Awaiting approval</span>
                  ) : undefined}
                </div>
                <div style={s.meta}>{trailOf(h.path)}</div>
                <div style={s.meta}>
                  {[h.author, formatDate(h.modified), formatBytes(h.size)]
                    .filter((x) => x.length > 0)
                    .join(" · ")}
                </div>
              </div>
            </button>
          ))}
          {allHits.length > shown ? (
            <button
              type="button"
              style={{ ...s.btnGhost, marginTop: 12 }}
              onClick={() => setShown((n) => n + pageSize)}
            >
              Show more ({allHits.length - shown} remaining)
            </button>
          ) : undefined}
        </div>
      )}
    </div>
  );
}
