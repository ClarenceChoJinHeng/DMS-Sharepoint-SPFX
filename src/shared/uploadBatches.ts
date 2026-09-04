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

/**
 * Opaque destination snapshot. `Form.tsx` owns the shape; nothing here inspects it.
 *
 * Deliberately not narrowed: the moment this module knows what a destination contains, it knows about
 * term GUIDs and SharePoint columns, and it stops being the thing that can be tested without a tenant.
 */
export type BatchDestination = Record<string, unknown>;

/** One file staged for upload. */
export interface StagedFile {
  /** Stable within a session. Never an array index — rows are removed mid-run (see applyUploadResults). */
  id: string;
  /** The picked browser file. Not serialisable, which is why staged work cannot survive a closed tab. */
  file: File;
  /** What the user typed in "Document Name". Blank keeps the original filename. */
  typedName: string;
  /**
   * The name this file will actually be uploaded under, computed by `Form.tsx`.
   *
   * NOT the typed name. The form composes `[Project] - [Vendor] - [Document Name] - [Date]` via
   * `composeUploadBase`, then extension-proofs it with `buildUploadName` — so the typed name is one
   * SEGMENT of the result, and two files with different typed names can still collide (same project,
   * same vendor, same date). Comparing typed names would miss exactly that case.
   *
   * Kept here rather than re-derived, because re-deriving would mean a second copy of the naming
   * convention living somewhere it could drift from the one that does the upload.
   */
  finalName?: string;
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
  /**
   * A SNAPSHOT of everything `Form.tsx` needs to write this batch, taken when the batch was saved.
   * Opaque here, as with `FileMeta`.
   *
   * A snapshot rather than a reference to the pickers, because by Upload the pickers describe whatever
   * batch is being edited NOW — batch 1 would be written using batch 3's department. It holds the
   * unit's server-relative URL, the below-Unit folder names, the level selections for
   * `buildLevelFormValues`, and the already-built tier field/value pairs.
   */
  destination: BatchDestination;
  files: StagedFile[];
  /** Set by `batchesNeedingRepick` when this batch's segment changed shape under it. */
  needsRepick?: boolean;
  /**
   * Epoch ms when the batch was staged, for the collapsed card's "Created on …" line.
   *
   * OPTIONAL so every existing construction site and test stays valid, and because nothing here
   * reads it — it is display only. This module never calls the clock (same rule as `auditLog.ts`):
   * the caller supplies the time, so the value can be asserted in a test.
   */
  createdAt?: number;
  /**
   * The FORM state this batch was saved from, so it can be re-opened and edited.
   *
   * Opaque here, exactly like `destination` - this module must not learn what a term GUID is. It is
   * a SEPARATE snapshot from `destination` on purpose: `destination` is what the UPLOAD writes and
   * must never change once taken, while this is what the PICKERS need to show the batch again. One
   * field serving both would mean the act of re-opening a batch could alter where it files.
   */
  editState?: BatchDestination;
}

/** The result of trying to upload one staged file. */
export interface UploadResult {
  fileId: string;
  ok: boolean;
  /** Required when `ok` is false — a row that failed with no reason is worse than no row at all. */
  error?: string;
  /**
   * A free name to offer, set ONLY when the failure was a name clash in the destination folder.
   *
   * Its presence is what tells the form this refusal is OFFERABLE — a clash on the APPROVED side
   * deliberately leaves it undefined, because that document has been through approval and filing a
   * renamed second copy is not the right answer there.
   */
  suggestedName?: string;
  /**
   * WHERE the name was already taken. Set only when the failure was a clash.
   *
   * ⚠ REPLACES THE BOOLEAN `approvedClash` (2026-08-28), and the widening is the point. A boolean
   * could express "approved side, or not", which was enough only while a pending draft was
   * untouchable. Now the uploader may replace one, so there are FOUR answers and the dialog offers a
   * different action for each.
   *
   * It exists so the dialog accounts for EVERY file that failed rather than silently omitting some —
   * a popup listing one of two failures reads as a bug (client, 2026-08-26: *"that is why I dont see
   * the first file"*).
   */
  clashWhere?: ClashWhere;
}

/**
 * The four ways a name can already be taken, from the uploader's point of view.
 *
 * - `staging`  — a document is already waiting for approval in this folder. Replaceable since
 *                2026-08-28: the uploader decides.
 * - `approved` — already approved and filed. Uploading proceeds normally and nothing is touched now;
 *                the FILED copy is replaced later, if an approver approves this one.
 * - `both`     — taken in both. ⚠ ONE action, not two: replacing the pending draft means that draft
 *                goes through approval, and Auto-route then replaces the filed copy anyway. Choosing
 *                "replace staging" IS choosing "replace Documents, later", so offering the two
 *                separately would imply a combination that cannot exist.
 * - `hidden`   — the pre-checks saw nothing and the WRITE was refused. Draft Item Security is
 *                *"only approvers and the author"*, so this is almost always a colleague's pending
 *                upload the uploader cannot see. ⚠ The dialog must SAY so: replacing here discards
 *                someone else's unreviewed work, and they are never asked.
 */
export type ClashWhere = "staging" | "approved" | "both" | "hidden";

/** What the two existence checks found, and what the uploader has already consented to. */
export interface ClashInputs {
  /** The name is taken in the APPROVAL library — a document already waiting for approval. */
  inStaging: boolean;
  /** The name is taken on the APPROVED side (`Documents` / `HC Documents`). */
  inApproved: boolean;
  /** The uploader chose "Replace the waiting document" for this file. */
  replaceStaging: boolean;
  /** The uploader chose "Send for approval as a replacement" for this file. */
  replaceApproved: boolean;
  /**
   * ⚠ FALSE FOR BULK UPLOAD, and this is the whole reason the flag exists.
   *
   * Bulk Upload may never overwrite a pending draft (client, 2026-08-28) — it files historical
   * documents, so a clash there is always somebody's live work. Passing `false` makes `staging` and
   * `both` permanently un-consentable, so that screen can only ever rename.
   *
   * Defaults to TRUE because the upload form is the common caller; a missing flag must not silently
   * disable a feature the client asked for.
   */
  allowStagingReplace?: boolean;
}

/** What to do about it. */
export interface ClashDecision {
  /** Absent when the name is free in both libraries. */
  where?: ClashWhere;
  /** May the upload proceed? False means refuse and offer the dialog. */
  consented: boolean;
  /** ⚠ The `overwrite` flag for `Files/Add`. True ONLY when deliberately replacing a pending draft. */
  overwrite: boolean;
}

/**
 * Given what the checks found, decide what happens.
 *
 * ⚠ EXTRACTED FROM `Form.tsx` (2026-08-28) SO IT CAN BE TESTED. The rule used to sit inside the
 * upload loop, tangled with the network calls that produce its inputs, which meant it could only be
 * verified by uploading a file and looking. It decides whether a document is destroyed, and Bulk
 * Upload needs the SAME rule with one deliberate difference — the exact shape where a hand-written
 * second copy drifts, and where the copy that drifts is the stricter one nobody exercises.
 *
 * ── The table (client, 2026-08-28) ───────────────────────────────────────────
 *   staging | approved | consent needed        | overwrite
 *   --------|----------|-----------------------|----------
 *      -    |    -     | none                  | no
 *     yes   |    -     | replaceStaging        | YES
 *      -    |   yes    | replaceApproved       | no
 *     yes   |   yes    | replaceStaging        | YES
 *
 * ⚠ `both` IS UNLOCKED BY THE STAGING CONSENT, NEVER THE APPROVED ONE. The pending draft is what
 * physically blocks the write, and replacing it is the single action that ALSO causes Auto-route to
 * replace the filed copy at approval. They are not independent choices, so the dialog offers one
 * button — and this function is what guarantees the two cannot come apart.
 *
 * ⚠ AN APPROVED-SIDE CLASH SETS NO `overwrite`. The name is free in the approval library, so that
 * upload is ordinary; the filed copy is replaced later by Auto-route, and only if an approver
 * approves it. Nothing is destroyed at upload time.
 *
 * The `hidden` case is NOT decided here: it is only knowable from a rejected write, long after this
 * runs, and it can only arise when consent was absent (a consented replace would have succeeded).
 */
export function decideClash(i: ClashInputs): ClashDecision {
  const allowStagingReplace = i.allowStagingReplace !== false;
  const where: ClashWhere | undefined =
    i.inStaging && i.inApproved
      ? "both"
      : i.inStaging
        ? "staging"
        : i.inApproved
          ? "approved"
          : undefined;

  if (where === undefined) return { consented: true, overwrite: false };

  const consented =
    where === "approved" ? i.replaceApproved : allowStagingReplace && i.replaceStaging;

  // `inStaging` is what makes an overwrite necessary at all, so an approved-side consent can never
  // produce one - which is what keeps "send for approval as a replacement" non-destructive.
  return { where, consented, overwrite: consented && i.inStaging };
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

/** How many suffixes to try before giving up. A folder with 999 same-named files is a different problem. */
const MAX_NAME_ATTEMPTS = 999;

/** Windows' own convention, and the reason this is a WORD rather than a number is in `nextAvailableName`. */
const COPY_SUFFIX = " - Copy";

/**
 * The first free name derived from `desired`, given the names already `taken`.
 *
 * Client, 2026-08-26: person A has a pending `TEST.pdf`; person B uploads `TEST.pdf` and should be
 * OFFERED a free name rather than refused - *"ask them if they want to proceed with the new file name
 * or cancel and rename that specific file themselves."*
 *
 * WARN: IT APPENDS AND NEVER STRIPS, AND THAT INVARIANT IS THE WHOLE CORRECTNESS ARGUMENT.
 * The suffix goes on the WHOLE original name, so two DIFFERENT originals can never be offered the SAME
 * name. `X.pdf` becomes `X - Copy.pdf` while `X (2).pdf` becomes `X (2) - Copy.pdf`.
 *
 * That property was broken for one day. 1.0.273.0 tried to be clever and INCREMENT a name already
 * ending ` (n)` - `X (2).pdf` -> `X (3).pdf` - which reads better per file and silently MERGES the
 * numbering space of `X` and `X (2)`. Client, 2026-08-27, reading a dialog that offered ` (3)` twice:
 * *"look at the second clash checking Documents Library, it increments to (3), so if both files go
 * through Approval Document to be auto approve for bulk upload, won't both files clash at the same
 * time?"* Exactly so. Stripping is what let two originals collapse onto one counter; appending cannot.
 *
 * A WORD, not a bare number, for the reason the client gave on 2026-08-26 when they rejected `TEST-2`:
 * a digit reads as one more segment of `[Project] - [Vendor] - [Name] - [Date]`, and ` - Copy` cannot
 * be mistaken for one. It is also what Windows itself produces, so nobody has to be taught it.
 *
 * WARN: IT PROBES RATHER THAN ASSUMING. ` - Copy` may itself be taken (this is the third person to
 * upload that name), so a blind append hands back a name that clashes just as hard and turns one
 * refusal into two. It walks ` - Copy`, ` - Copy (2)`, ` - Copy (3)`, ... until something is free.
 *
 * WARN: IT CANNOT SEE OTHER FILES IN THE SAME RUN. Two files whose names are identical would each be
 * offered the same free name, because each call knows only what is in `taken`. Callers running a batch
 * must feed every name they have already OFFERED back in - see `reserved` in `Form.tsx`/`BulkUpload.tsx`.
 *
 * WARN: SUFFIXED BEFORE THE EXTENSION, never after: `TEST - Copy.pdf`, not `TEST.pdf - Copy`. The
 * extension decides what opens the file, and the whole convention here preserves it.
 *
 * Case-insensitive, because SharePoint file names are - `test.pdf` and `TEST.pdf` cannot coexist in one
 * folder, so treating them as distinct would suggest a name that is not actually free.
 */
export function nextAvailableName(desired: string, taken: readonly string[]): string {
  const name = (desired ?? "").trim();
  if (name.length === 0) return name;
  const lower: Record<string, true> = {};
  (taken ?? []).forEach((t) => {
    lower[(t ?? "").trim().toLowerCase()] = true;
  });
  if (!lower[name.toLowerCase()]) return name;

  // Split on the LAST dot, and only when it is not the first character: `.gitignore` is all stem.
  const dot = name.lastIndexOf(".");
  const stem = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : "";

  const first = `${stem}${COPY_SUFFIX}${ext}`;
  if (!lower[first.toLowerCase()]) return first;

  for (let n = 2; n <= MAX_NAME_ATTEMPTS; n++) {
    const candidate = `${stem}${COPY_SUFFIX} (${n})${ext}`;
    if (!lower[candidate.toLowerCase()]) return candidate;
  }
  /* Nothing free in 999 tries. Returning `desired` would be the WORST answer - it is known-taken, so
     the caller would offer a rename that changes nothing. The timestamp cannot collide in practice and
     the upload's own `overwrite=false` is still behind it. */
  return `${stem}${COPY_SUFFIX} (${Date.now()})${ext}`;
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
    // `finalName` when the form has computed it — the composed
    // `[Project] - [Vendor] - [Name] - [Date]`, which is what actually reaches SharePoint. The
    // fallback keeps this module usable on its own; the form always supplies the real name.
    const key = (f.finalName && f.finalName.trim().length > 0
      ? f.finalName
      : resolveUploadName(f.file?.name ?? "", f.typedName)
    ).toLowerCase();
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
