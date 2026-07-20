// Pure, SPFx-free helpers for the segment Documents view web part (multi-site Phase 2).
// Config comes from the "DMS Site Map" list: which segment lives on which site.
// No imports from @microsoft/* — keep this unit-testable in plain Jest.

/** One raw row from the DMS Site Map list. */
export interface SiteMapRow {
  title: string; // segment display label (list Title)
  siteUrl: string; // absolute segment site URL
  documentsLibrary: string; // library name (blank -> "Documents")
}

/** A normalised, browsable segment target. */
export interface SegmentTarget {
  segment: string; // display label
  siteUrl: string; // trimmed, no trailing slash
  documentsLibrary: string; // never blank
}

const noTrailing = (s: string): string => (s ?? "").trim().replace(/\/+$/, "");

/**
 * Reduce raw DMS Site Map rows to valid, normalised targets.
 * Skips rows with no site URL; defaults a blank library to "Documents";
 * falls back to the site URL as the label when Title is blank.
 */
export function parseSiteMapRows(rows: SiteMapRow[]): SegmentTarget[] {
  const out: SegmentTarget[] = [];
  for (const r of rows ?? []) {
    const siteUrl = noTrailing(r?.siteUrl);
    if (!siteUrl) continue;
    out.push({
      segment: (r.title ?? "").trim() || siteUrl,
      siteUrl,
      documentsLibrary: (r.documentsLibrary ?? "").trim() || "Documents",
    });
  }
  return out;
}

/** Absolute origin of a site URL (for building file open links). "" if invalid. */
export function siteOrigin(siteUrl: string): string {
  try {
    return new URL(siteUrl).origin;
  } catch {
    return "";
  }
}

/** Server-relative path to a library's root folder. "" if the URL is invalid. */
export function libraryRootPath(siteUrl: string, library: string): string {
  try {
    const u = new URL(siteUrl);
    const lib = (library ?? "").trim() || "Documents";
    return `${noTrailing(u.pathname)}/${lib}`;
  } catch {
    return "";
  }
}
