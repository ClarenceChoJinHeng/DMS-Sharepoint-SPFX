/**
 * Turning a flat list of uploaded files into submissions, batches and files.
 *
 * Client, 2026-08-22: My Submissions should show what the uploader actually DID — a submission — and
 * let them open it to see the batches, and open a batch to see each file's details as they filled them
 * in. Spec: docs/superpowers/specs/2026-08-22-submission-grouping-design.md
 *
 * ⚠ THE IDENTITY IS STAMPED, NEVER INFERRED. `Batch.id` in `uploadBatches.ts` is session-only, so
 * until 2026-08-22 nothing on a document said which upload it came from. Two text columns now carry it
 * (`SubmissionId` per press of Upload, `BatchId` per destination folder), and everything here reads
 * those.
 *
 * ⚠ FILES UPLOADED BEFORE THAT CARRY NO STAMP, AND ARE GROUPED BY FOLDER AND DAY — an inference the
 * spec first rejected and the client overruled on seeing the result (see `legacyKey`). Such a group
 * is flagged `inferred` so the screen can label it as a guess rather than as a recorded submission.
 */

import { Submission, SubmissionStatus, sortNewestFirst } from "./mySubmissions";

/* Generic over the row type, defaulted so every existing use is unchanged. The page groups
   `MergedRow`s — a Submission plus `recordState` — and a non-generic group would erase the field it
   then has to branch on, which is exactly how a deleted file gets rendered as a live one. */
export interface BatchGroup<T extends Submission = Submission> {
  /** `BatchId` as stored. "" for a file that carries none. */
  reference: string;
  files: T[];
}

export interface SubmissionGroup<T extends Submission = Submission> {
  /** `SubmissionId` as stored. "" for a file uploaded before this existed — see `groupSubmissions`. */
  reference: string;
  /**
   * True when this group was INFERRED from folder and date rather than read from a stamped
   * reference — i.e. every file in it predates submissions being recorded.
   *
   * The screen must say so. An inferred group is a good guess and nothing more, and presenting it
   * as a submission the uploader actually made is the failure mode the spec warned about.
   */
  inferred: boolean;
  /** Earliest upload time in the submission. Undefined only when no file carried one. */
  at?: Date;
  files: T[];
  batches: BatchGroup<T>[];
}

function key(v: string | undefined): string {
  return (v ?? "").trim();
}

/** The server-relative folder a file sits in — its path with the file name removed. */
export function folderOf(fileRef: string | undefined): string {
  const p = (fileRef ?? "").trim();
  const cut = p.lastIndexOf("/");
  return cut > 0 ? p.slice(0, cut) : p;
}

/** `20260822`, or "" when the row carried no time. Local, matching what the screen displays. */
function dayOf(at: Date | undefined): string {
  if (!(at instanceof Date) || isNaN(at.getTime())) return "";
  return `${at.getFullYear()}${`0${at.getMonth() + 1}`.slice(-2)}${`0${at.getDate()}`.slice(-2)}`;
}

/**
 * The grouping key for a file with NO stamped reference.
 *
 * ⚠ THIS IS AN INFERENCE, AND IT REPLACED A DELIBERATE REFUSAL TO INFER (client, 2026-08-22, on
 * seeing the first build: *"client doesn't want to go through each file one by one figuring out and
 * noting down which one"*). The original rule — one row per referenceless file — is what the data
 * honestly supports, and on a site with a year of history it produces exactly the flat list this
 * whole feature exists to replace.
 *
 * FOLDER **AND DAY**, never folder alone: a unit files into the same folder every month, so the
 * folder by itself would present a year of unrelated uploads as one submission. Same day into the
 * same destination is the strongest evidence available that these went together, and it is the shape
 * a real batch has.
 *
 * The cost, accepted and stated on screen: two people filing into one folder on one day are shown as
 * one group. Nothing is hidden and no file moves — only how the rows are stacked — and `inferred`
 * makes the screen label it rather than pass it off as a recorded submission.
 *
 * A row with NO date falls back to its own item path, so it groups with nothing: a missing timestamp
 * is unknown, and unknown must never widen a group. Empty ≠ unknown, again.
 */
function legacyKey(row: Submission): string {
  const day = dayOf(row.created);
  if (day.length === 0) return `?${row.library}#${row.itemId}`;
  return `${folderOf(row.fileRef).toLowerCase()}|${day}`;
}

/**
 * Group rows into submissions, each with its batches.
 *
 * ⚠ A ROW WITH NO `SubmissionId` IS GROUPED BY FOLDER AND DAY — see `legacyKey`, which carries why
 * this reverses the original refusal to infer. Such a group is flagged `inferred`, and the screen
 * must say so: it is a presentation of evidence, not a record of what someone did.
 *
 * Newest first, reusing `sortNewestFirst` so this list and the flat one cannot disagree about order.
 */
export function groupSubmissions<T extends Submission>(rows: readonly T[]): SubmissionGroup<T>[] {
  const bySubmission: SubmissionGroup<T>[] = [];
  const index: Record<string, SubmissionGroup<T>> = {};

  for (const row of sortNewestFirst(rows ?? [])) {
    const sid = key(row.submissionId);
    const inferred = sid.length === 0;
    // A stamped id is the group key; without one, folder + day is the inference. The two key spaces
    // are kept apart by the `~` prefix, so a stored id can never collide with a derived one.
    const gk = inferred ? `~${legacyKey(row)}` : sid;
    let group: SubmissionGroup<T> | undefined = index[gk];
    if (!group) {
      group = { reference: sid, inferred, at: row.created, files: [], batches: [] };
      bySubmission.push(group);
      index[gk] = group;
    }
    group.files.push(row);
    if (row.created && (!group.at || row.created < group.at)) group.at = row.created;

    const bid = key(row.batchId);
    // An inferred group's batch is its folder — which is the same for every file in it by
    // construction, so it lands as one batch rather than one per file.
    const bk = bid.length > 0 ? bid : `~${folderOf(row.fileRef).toLowerCase()}`;
    let batch = group.batches.filter((b) => (b.reference || `~${folderOf(b.files[0]?.fileRef).toLowerCase()}`) === bk)[0];
    if (!batch) {
      batch = { reference: bid, files: [] };
      group.batches.push(batch);
    }
    batch.files.push(row);
  }
  return bySubmission;
}

/**
 * How many files in each state.
 *
 * ⚠ NEVER COLLAPSED TO A SINGLE STATUS. A submission can be part approved and part pending, and
 * showing one badge for it is how an uploader concludes a rejected file was fine. The caller renders
 * whichever counts are non-zero.
 */
export function statusCounts<T extends Submission>(files: readonly T[]): Record<SubmissionStatus, number> {
  const out: Record<SubmissionStatus, number> = { Pending: 0, Approved: 0, Rejected: 0 };
  for (const f of files ?? []) if (out[f.status] !== undefined) out[f.status] += 1;
  return out;
}

const ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";

/**
 * A human-readable reference: `SUB-20260822-K4P2`.
 *
 * Readable rather than a GUID because it appears on screen as the thing an uploader quotes to their
 * approver — the client's own claim system shows a Reference No, and that is the shape they expect.
 * Uniqueness comes from the random suffix; the date is there to be read, not to disambiguate.
 *
 * `I`, `L`, `O`, `0` and `1` are absent from the alphabet: this gets read aloud and written down.
 *
 * `rand` is injected so the format can be pinned by test without a stubbed global.
 */
export function newReference(prefix: string, now: Date, rand: () => number = Math.random): string {
  const d = now instanceof Date && !isNaN(now.getTime()) ? now : new Date();
  const yyyy = d.getFullYear();
  const mm = `0${d.getMonth() + 1}`.slice(-2);
  const dd = `0${d.getDate()}`.slice(-2);
  let suffix = "";
  for (let i = 0; i < 4; i += 1) {
    const pick = Math.floor(rand() * ALPHABET.length);
    suffix += ALPHABET.charAt(Math.min(Math.max(pick, 0), ALPHABET.length - 1));
  }
  return `${prefix}-${yyyy}${mm}${dd}-${suffix}`;
}
