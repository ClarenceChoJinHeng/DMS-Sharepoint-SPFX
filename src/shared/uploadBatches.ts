/**
 * Batched multi-file upload — the rules, with no React and no network.
 *
 * Spec: docs/superpowers/specs/2026-08-15-batched-multi-file-upload-design.md
 *
 * A BATCH is one destination folder plus the files that belong in it. The client asked for "each file
 * to go to different places", then refined it to batches — so the destination is chosen per batch and
 * one upload carries several. Save stages; nothing reaches SharePoint until Upload.
 *
 * WHY THIS IS A SEPARATE FILE AT ALL. Batching was built into `BulkUpload.tsx` on 2026-07-24 and
 * REMOVED on 2026-08-03, for a reason that applies here unchanged: `Batch`, `BatchSelection`,
 * `BatchOutcome`, `LiveBatch` and their panel cost 169 references in a 2,586-line file already over the
 * 2,000-line lint limit. `Form.tsx` is 2,653 lines. Repeating that shape inline would repeat that
 * outcome, so every rule decidable without the network lives here, under test.
 *
 * It is also written so Bulk Upload's deferred "Phase 2" can adopt it: the difference there is that a
 * batch carries ONE metadata set rather than one per file — a difference in what a batch CONTAINS, not
 * in what a batch IS. Hence `StagedFile.meta` is opaque; this module never looks inside it.
 */

/** Opaque per-file metadata. This module never reads a field; `Form.tsx` owns the shape. */
export type FileMeta = Record<string, string>;

/** One file staged for upload. */
export interface StagedFile {
  /** Stable within a session. Never an array index — rows are removed mid-run (see applyUploadResults). */
  id: string;
  /** The picked browser file. Not serialisable, which is why staged work cannot survive a closed tab. */
  file: File;
  /** What the user typed in "Document Name". Blank keeps the original filename. */
  typedName: string;
  meta: FileMeta;
  /** Set only after a failed upload attempt — the reason shown on the row. */
  error?: string;
}

/** One destination folder plus its files. */
export interface Batch {
  id: string;
  /**
   * The mode/segment key this destination belongs to. Groups the chain re-read, so two batches on one
   * segment cost one request rather than two.
   */
  segmentKey: string;
  /** `chainSignature(levels)` captured when the batch was saved — see `batchesNeedingRepick`. */
  chainSignature: string;
  /** Folder names in chain order, for display: ["GHO", "Group Finance", "Corporate", "2026", "Tax Return"]. */
  pathLabels: string[];
  /** Everything `Form.tsx` needs to rebuild the write. Opaque here, as with `FileMeta`. */
  destination: Record<string, string>;
  files: StagedFile[];
  /** Set by `batchesNeedingRepick` when this batch's segment changed shape under it. */
  needsRepick?: boolean;
}

/** The result of trying to upload one staged file. */
export interface UploadResult {
  fileId: string;
  ok: boolean;
  /** Required when `ok` is false — a row that failed with no reason is worse than no row at all. */
  error?: string;
}

/* ── Names ──────────────────────────────────────────────────────────────────── */

/**
 * The name a file will be uploaded under.
 *
 * Mirrors `buildUploadName` in Form.tsx: the typed name replaces the stem, the original extension is
 * always preserved, and blank keeps the original filename. Duplicated rather than imported because
 * Form.tsx is a web part module and not Jest-loadable — if that helper ever moves to `shared/`, delete
 * this one rather than keeping both.
 */
export function resolveUploadName(original: string, typed: string): string {
  const name = original ?? "";
  const dot = name.lastIndexOf(".");
  const ext = dot > 0 ? name.slice(dot) : "";
  const stem = (typed ?? "").trim();
  if (stem.length === 0) return name;
  // Illegal SharePoint characters are STRIPPED, not substituted — the same choice as
  // `sanitizeFolderSegment`, for the same reason: a substituted separator invents a name.
  const clean = stem.replace(/[*:<>?/\\|"#%{}~&]/g, "").trim();
  if (clean.length === 0) return name;
  return clean.toLowerCase().endsWith(ext.toLowerCase()) ? clean : `${clean}${ext}`;
}

/**
 * File ids whose resolved upload name clashes with another file in the SAME batch.
 *
 * Blocking at save time is possible only because the rule is pure. Two files resolving to one name
 * would put one on top of the other, and SharePoint would not complain.
 *
 * Case-insensitive, because SharePoint file names are. EVERY member of a clashing set is returned —
 * which one is "wrong" is not knowable, exactly as with the abbreviation sibling check.
 */
export function collisionsWithin(batch: Batch): string[] {
  const byName = new Map<string, string[]>();
  for (const f of batch?.files ?? []) {
    const key = resolveUploadName(f.file?.name ?? "", f.typedName).toLowerCase();
    const seen = byName.get(key);
    if (seen) seen.push(f.id);
    else byName.set(key, [f.id]);
  }
  const out: string[] = [];
  byName.forEach((ids) => {
    if (ids.length > 1) out.push(...ids);
  });
  return out;
}

/** A batch may be saved only when it has files, a destination, and no internal name clash. */
export function canSaveBatch(batch: Batch): boolean {
  if (!batch || (batch.files ?? []).length === 0) return false;
  if ((batch.pathLabels ?? []).length === 0) return false;
  return collisionsWithin(batch).length === 0;
}

/* ── Defaults ───────────────────────────────────────────────────────────────── */

/**
 * Values a newly added file should start with: those of the last file already in this batch.
 *
 * Not invented data — copied from a sibling the user typed, shown, and editable. Without it a ten-file
 * batch means typing one vendor ten times, and effort is the whole complaint.
 *
 * The typed NAME is deliberately never inherited: two files sharing a name is precisely the collision
 * `collisionsWithin` exists to block.
 */
export function inheritDefaults(batch: Batch): FileMeta {
  const files = batch?.files ?? [];
  if (files.length === 0) return {};
  return { ...files[files.length - 1].meta };
}

/* ── Partial failure ────────────────────────────────────────────────────────── */

/**
 * Fold a run's results back into the staged batches.
 *
 * SUCCESS REMOVES; FAILURE STAYS. That one rule is what makes Retry safe: an uploaded file is no longer
 * in the list, so it cannot be sent twice, and what remains on screen is exactly what still needs
 * doing. Rollback was rejected — it would mean deleting files that already uploaded, which can fail on
 * its own, and could delete a document an approver has already opened.
 *
 * A file with NO result is left untouched: the run may have stopped early, and marking it failed would
 * report a failure that never happened. Its stale error from an earlier run IS cleared, so a row never
 * shows a reason that no longer applies. An emptied batch is dropped.
 */
export function applyUploadResults(batches: Batch[], results: UploadResult[]): Batch[] {
  const byId = new Map<string, UploadResult>();
  for (const r of results ?? []) byId.set(r.fileId, r);

  const out: Batch[] = [];
  for (const b of batches ?? []) {
    const kept: StagedFile[] = [];
    for (const f of b.files ?? []) {
      const r = byId.get(f.id);
      if (r && r.ok) continue; // uploaded — it leaves the staging area
      if (r && !r.ok) {
        kept.push({ ...f, error: r.error ?? "Upload failed" });
        continue;
      }
      kept.push(f.error === undefined ? f : { ...f, error: undefined });
    }
    if (kept.length > 0) out.push({ ...b, files: kept });
  }
  return out;
}

/** Counts for the result line — *"Uploaded 4 of 6. 2 could not be uploaded."* Counts, never a verdict. */
export function summarise(results: UploadResult[]): { ok: number; failed: number; total: number } {
  const list = results ?? [];
  const ok = list.filter((r) => r.ok).length;
  return { ok, failed: list.length - ok, total: list.length };
}

/* ── Totals and limits ──────────────────────────────────────────────────────── */

/** Advisory thresholds. A WARNING, never a block — see `stagedTotals`. */
export const WARN_FILE_COUNT = 25;
export const WARN_TOTAL_BYTES = 200 * 1024 * 1024;

/**
 * How much is staged, and whether that is worth a warning.
 *
 * Advisory only. The browser holds every byte of every staged file, so a client queueing a hundred
 * scans should hear about it before the tab dies — but a hard block would refuse work the client
 * explicitly asked to be able to do, and whether the tab survives is their call, not ours.
 */
export function stagedTotals(batches: Batch[]): {
  files: number;
  bytes: number;
  overCount: boolean;
  overBytes: boolean;
} {
  let files = 0;
  let bytes = 0;
  for (const b of batches ?? []) {
    for (const f of b.files ?? []) {
      files++;
      bytes += f.file?.size ?? 0;
    }
  }
  return {
    files,
    bytes,
    overCount: files > WARN_FILE_COUNT,
    overBytes: bytes > WARN_TOTAL_BYTES,
  };
}

/**
 * Names staged in more than one batch.
 *
 * Warned, never blocked: filing one document in two places is unusual but legitimate, and refusing it
 * would invent a rule the client never asked for. Matched on name + size, because two `File` objects
 * from separate picks are never identity-equal even for the same file on disk.
 */
export function duplicateAcrossBatches(batches: Batch[]): string[] {
  const seen = new Map<string, number>();
  const label = new Map<string, string>();
  for (const b of batches ?? []) {
    for (const f of b.files ?? []) {
      const name = f.file?.name ?? "";
      const key = `${name.toLowerCase()}|${f.file?.size ?? 0}`;
      seen.set(key, (seen.get(key) ?? 0) + 1);
      label.set(key, name);
    }
  }
  const out: string[] = [];
  seen.forEach((n, key) => {
    if (n > 1) out.push(label.get(key) ?? "");
  });
  return out;
}

/* ── The stale-chain guard ──────────────────────────────────────────────────── */

/**
 * Mark the batches whose segment changed shape while they sat staged.
 *
 * `Form.tsx` already refuses to write when the mode row's `Levels` moved under it (gotcha 10b) — a
 * stale page otherwise files into the OLD folder shape, succeeds, and looks entirely normal. Its
 * remedy is "reload the page", which with staged batches would destroy every batch the client built.
 *
 * So the guard becomes gentler exactly as it becomes more likely: only batches on the moved segment are
 * marked, they STAY staged, and the client re-picks that one destination. Batches on other segments are
 * untouched and upload normally.
 *
 * `freshBySegment` OMITS a segment whose chain could not be read. That is not "unchanged" — it is
 * unknown, and unknown must never mark a batch: a transient error would otherwise take the form out of
 * service, which is worse than the drift it guards against. Fail-open, as everywhere else here.
 */
export function batchesNeedingRepick(
  batches: Batch[],
  freshBySegment: Record<string, string | undefined>,
): Batch[] {
  return (batches ?? []).map((b) => {
    const fresh = (freshBySegment ?? {})[b.segmentKey];
    if (fresh === undefined) return b.needsRepick ? { ...b, needsRepick: false } : b;
    const moved = fresh !== b.chainSignature;
    if (moved === Boolean(b.needsRepick)) return b;
    return { ...b, needsRepick: moved };
  });
}

/** Batches safe to upload right now — everything not waiting on a re-pick. */
export function uploadableBatches(batches: Batch[]): Batch[] {
  return (batches ?? []).filter((b) => !b.needsRepick);
}

/* ── Ids ────────────────────────────────────────────────────────────────────── */

/**
 * Session-unique ids.
 *
 * A counter, not `Math.random()` or a timestamp: ids need only be unique within one page, and a
 * deterministic sequence keeps the tests readable.
 */
let seq = 0;
export function nextId(prefix: string): string {
  seq += 1;
  return `${prefix}_${seq}`;
}

/** Test-only: make id sequences reproducible across cases. */
export function resetIds(): void {
  seq = 0;
}
