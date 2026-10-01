/**
 * Human-readable file sizes for the upload cards.
 *
 * Shared because Form and BulkUpload both show it, and they must agree: two
 * different renderings of the same file is the kind of inconsistency that gets
 * reported as a bug in whichever one the user saw second.
 *
 * No @microsoft/* imports — that is what keeps this covered by plain Jest. See
 * the note on folderAbbreviation.ts's ABBREV_LIST for why that matters.
 */

/** Binary units, matching what Windows and SharePoint both report. */
const KB = 1024;
const MB = 1048576;

/**
 * Format a byte count the way the client mockup shows it: `2MB`, no space.
 *
 * Sub-megabyte files fall back to KB rather than rendering as `0MB`, which is
 * what a plain megabyte conversion produces for the small PDFs that make up most
 * uploads. A trailing `.0` is stripped so a 2 MB file reads `2MB` rather than
 * `2.0MB`, matching the mockup exactly.
 */
export function formatFileSize(bytes: number): string {
  if (!isFinite(bytes) || bytes < 0) return "";
  if (bytes < KB) return `${Math.round(bytes)}B`;
  if (bytes < MB) return `${Math.round(bytes / KB)}KB`;
  return `${(bytes / MB).toFixed(1).replace(/\.0$/, "")}MB`;
}
