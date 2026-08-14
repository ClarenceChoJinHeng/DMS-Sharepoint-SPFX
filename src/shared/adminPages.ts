/**
 * The CRS Settings landing page — what it lists, and how each row finds its page.
 *
 * Spec: docs/superpowers/specs/2026-08-14-crs-settings-landing-page-design.md
 *
 * Pure and SPFx-free, because the page CONTENT is the deliverable here and the risky part is link
 * resolution. The client's complaint is that they cannot tell what to operate; a landing page whose
 * links quietly point at the wrong page, or at nothing, makes that worse rather than better.
 *
 * Links are matched against the site's real Site Pages rather than hardcoded, for one specific
 * reason: the client renames every DMS-named list, group and content type at import (memory
 * `dms-to-crs-rename-pending`), and page names will not survive either. A hardcoded
 * `Folder-Administration.aspx` fails as a DEAD LINK — no error, no clue — on the one page whose whole
 * job is telling an admin where to go.
 *
 * This module deliberately does NOT decide who may open these pages. That is `pageAccessPolicy.ts`,
 * which answers a different question about the same pages; a second copy of that rule is how the two
 * would come to disagree.
 */

/** One page in the site's Site Pages library, as this module needs it. */
export interface SitePage {
  /** e.g. "Folder-Administration.aspx" */
  fileName: string;
  /** The display title, which the client may have changed independently of the file name. */
  title: string;
  /** Server-relative path — the only thing safe to navigate to. */
  serverRelativeUrl: string;
}

/** A row (or a whole card) that navigates somewhere. */
export interface AdminLink {
  /** Stable id. Used for the property-pane override, so it must never change once shipped. */
  key: string;
  /** What the client reads. */
  label: string;
  /** Matched against BOTH the file name and the title. */
  match: RegExp;
  /**
   * Tab to open on the destination, appended as `#tab=<slug>`.
   *
   * Needed because Folder Administration is ONE page with five tabs, so three of the rows below
   * share a destination. Without this they would all land on the same tab and two of the three would
   * look broken — precisely the confusion this page exists to remove.
   */
  tab?: string;
  /**
   * The page this row lives on, when that is not its own label.
   *
   * Only needed for the rows that SHARE a destination: "Folder Reconciliations" lives on Folder
   * Administration, so telling an admin to create `Folder-Reconciliations.aspx` would have them make
   * a page the pattern can never match — advice that leads nowhere is worse than none. Caught by the
   * test that every suggestion must resolve.
   */
  pageName?: string;
}

export interface AdminCard {
  key: string;
  title: string;
  blurb: string;
  /** Which icon to draw. The component owns the artwork; this only names it. */
  icon: "folders" | "access" | "config" | "audit" | "bulk";
  /** Left or right column of the two-column grid. */
  column: 1 | 2;
  /** Rows under the card. Empty when the card is itself a single link. */
  links: AdminLink[];
  /** Set when the whole card is one link — the arrow sits on the card header. */
  self?: AdminLink;
}

/**
 * The Folder Administration page, referenced by three different rows.
 *
 * `folder.?admin|folder.?manage|folder.?structure` and NOT a bare `folder`: `Folder Access` is a
 * different page and a loose pattern would claim it, sending an admin who clicked "Folder
 * Reconciliations" to a permissions screen instead. That collision is why these patterns are pinned
 * by test.
 */
const FOLDER_ADMIN = /folder.?admin|folder.?manage|folder.?structure/i;

export const CARDS: AdminCard[] = [
  {
    key: "folders",
    title: "Folder Management",
    blurb: "Create, rename, and assign permissions to folders at any level — all at once.",
    icon: "folders",
    column: 1,
    links: [
      // In the order the work happens, which is also the order the tabs sit in. Abbreviations first
      // because that is the step whose omission makes reconciliation create nothing, silently.
      // All three name Folder Administration as their page: they are TABS of it, not pages of their
      // own, so the "create a page called…" advice has to say the page that exists.
      { key: "abbreviations", label: "CRS Term Abbreviations", match: FOLDER_ADMIN, tab: "abbreviations", pageName: "Folder Administration" },
      { key: "structure", label: "Folder Structure Management", match: FOLDER_ADMIN, tab: "structure", pageName: "Folder Administration" },
      { key: "reconciliation", label: "Folder Reconciliations", match: FOLDER_ADMIN, tab: "reconciliation", pageName: "Folder Administration" },
    ],
  },
  {
    key: "access",
    title: "User Access Management",
    blurb: "Manage who can access sites, approval libraries, pages and folders.",
    icon: "access",
    column: 1,
    links: [
      // Group Management leads, and is not on the client's mockup. It has to be here: Folder Access
      // maps an EXISTING group, so an admin starting there has nothing to pick.
      { key: "groups", label: "Group Management", match: /group.?manage/i },
      { key: "siteAccess", label: "Site Access", match: /site.?access/i },
      { key: "libraryAccess", label: "Approval Library Access", match: /approval.?library|library.?access/i },
      { key: "pageAccess", label: "Page Access", match: /page.?access/i },
      { key: "folderAccess", label: "Folder Access", match: /folder.?access/i },
    ],
  },
  {
    key: "config",
    title: "CRS Configuration",
    blurb: "Review and update the central repository settings and defaults.",
    icon: "config",
    column: 2,
    links: [],
    self: { key: "config", label: "CRS Configuration", match: /configuration|crs.?config/i },
  },
  {
    key: "audit",
    title: "CRS Audit Log",
    blurb: "See who changed what, and when — access, structure and folder runs.",
    icon: "audit",
    column: 2,
    links: [],
    self: { key: "audit", label: "CRS Audit Log", match: /audit/i },
  },
  {
    key: "bulk",
    title: "Bulk Upload",
    blurb: "Upload many documents at once into an existing folder.",
    icon: "bulk",
    column: 2,
    links: [],
    self: { key: "bulk", label: "Bulk Upload", match: /bulk/i },
  },
];

/** Every link on the page, flattened — for resolution and for the property-pane overrides. */
export function allLinks(cards: readonly AdminCard[] = CARDS): AdminLink[] {
  const out: AdminLink[] = [];
  for (const c of cards ?? []) {
    if (c.self) out.push(c.self);
    for (const l of c.links) out.push(l);
  }
  return out;
}

/**
 * Where a link points.
 *
 * Three states, not two. `missing` is not an error to swallow: the row renders disabled with the name
 * to give the page, because a dead arrow on a navigation page is the worst failure available here —
 * it teaches the client the tool is broken. `ambiguous` still navigates (a best guess beats nothing)
 * but says so, so a duplicated or half-renamed page is visible rather than silently picked.
 */
export type LinkTarget =
  | { state: "resolved"; url: string }
  | { state: "ambiguous"; url: string; others: string[] }
  | { state: "missing" };

/** Append the tab hash, if the link names one. */
function withTab(url: string, tab?: string): string {
  return tab ? `${url}#tab=${encodeURIComponent(tab)}` : url;
}

/**
 * Resolve one link against the site's pages.
 *
 * An `override` (from the property pane) wins outright and is never pattern-checked — it is the
 * escape hatch for a site whose page names this module cannot guess, and second-guessing it would
 * remove the only fix available without a redeploy.
 *
 * Otherwise the pattern is matched against the file name AND the title, because the client may rename
 * either independently. On several matches the SHORTEST file name wins: a duplicate is almost always
 * the longer name (`Folder-Administration-Copy.aspx`, `…-old.aspx`), so the canonical page is the
 * terse one. Explainable, and reported rather than hidden.
 */
export function resolveLink(
  link: AdminLink,
  pages: readonly SitePage[],
  override?: string,
): LinkTarget {
  const manual = (override ?? "").trim();
  if (manual.length > 0) return { state: "resolved", url: withTab(manual, link.tab) };

  const hits = (pages ?? []).filter(
    (p) => link.match.test(p.fileName ?? "") || link.match.test(p.title ?? ""),
  );
  if (hits.length === 0) return { state: "missing" };

  const sorted = hits.slice().sort((a, b) => {
    const byLength = (a.fileName ?? "").length - (b.fileName ?? "").length;
    return byLength !== 0 ? byLength : (a.fileName ?? "").localeCompare(b.fileName ?? "");
  });
  const chosen = sorted[0];
  const url = withTab(chosen.serverRelativeUrl, link.tab);
  if (sorted.length === 1) return { state: "resolved", url };
  return { state: "ambiguous", url, others: sorted.slice(1).map((p) => p.fileName) };
}

/** Resolve every link, keyed for the component. */
export function resolveAll(
  pages: readonly SitePage[],
  overrides: Record<string, string> = {},
  cards: readonly AdminCard[] = CARDS,
): Record<string, LinkTarget> {
  const out: Record<string, LinkTarget> = {};
  for (const l of allLinks(cards)) out[l.key] = resolveLink(l, pages, (overrides ?? {})[l.key]);
  return out;
}

/**
 * A page name to suggest when a link resolves to nothing.
 *
 * `pageName` when the row is a tab of a shared page, otherwise the label. The distinction is not
 * cosmetic: three rows here are tabs of Folder Administration, and suggesting
 * `Folder-Reconciliations.aspx` would send an admin to create a page the pattern can never match —
 * so the row would stay dead after they did exactly as they were told. Pinned by a test asserting
 * every suggestion resolves.
 */
export function suggestedPageName(link: AdminLink): string {
  const base = (link.pageName ?? link.label ?? "").trim();
  return `${base.replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "")}.aspx`;
}

/** Cards for one column, in declaration order. */
export function columnCards(column: 1 | 2, cards: readonly AdminCard[] = CARDS): AdminCard[] {
  return (cards ?? []).filter((c) => c.column === column);
}

/**
 * The tab slug from a URL hash, for the destination page to read.
 *
 * Returns "" for anything unrecognised, which every caller must treat as "open on the default tab" —
 * a stray hash must never leave a tabbed page showing nothing.
 */
export function tabFromHash(hash: string): string {
  const m = /[#&]tab=([^&]+)/.exec(hash ?? "");
  if (!m) return "";
  try {
    return decodeURIComponent(m[1]).trim().toLowerCase();
  } catch {
    return "";
  }
}
