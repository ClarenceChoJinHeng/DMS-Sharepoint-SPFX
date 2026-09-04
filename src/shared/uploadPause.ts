// Site-wide upload pause, for the duration of a folder-structure change.
//
// Spec: docs/superpowers/specs/2026-08-19-upload-pause-design.md
//
// Clarence, 2026-08-19: *"they are not going to go to Pause power automate, so that is why I was
// thinking of building that toggle to stop user from uploading, both in normal upload form and bulk
// upload. I will tell the client to only do file structure changes after work hours."*
//
// ⚠ WHAT THIS PREVENTS. A file uploaded DURING a migration lands in the OLD shape — correctly, since
// `Levels` is still the old chain until the run's last step. The problem is timing: if it arrives
// after the migrator has scanned its folder, it is not moved. The staged-apply model already stops
// that going live silently (the pending chain is applied only when a FRESH scan finds no drift, so a
// slipped file leaves the migration pending), but on a live site that can loop indefinitely.
//
// ⚠ WHAT IT CANNOT PREVENT, and the client is told so rather than reassured: a browser tab opened
// BEFORE the pause holds its own state. The write-time re-check below closes most of that window —
// but not a session that never talks to the server again until it uploads. The operational answer is
// the client's: make structure changes outside working hours.
//
// SITE-WIDE, not per segment (Clarence's decision). A per-segment pause is narrower, but it is one
// more thing to reason about at the moment an admin is already mid-migration, and the work happens
// after hours when nobody is uploading anywhere.

/** The DMS Config `setting` row that holds the pause. */
export const UPLOAD_PAUSE_SETTING = "uploadsPaused";

/**
 * The message every blocked uploader sees.
 *
 * Says WHAT is happening, that nothing is wrong with their file, and roughly when to come back. An
 * unexplained disabled upload form reads as a broken system, and the uploader's next move is to mail
 * an administrator who is in the middle of a migration.
 */
export const UPLOAD_PAUSE_MESSAGE =
  "Uploads are paused while an administrator reorganises the document folders. " +
  "Nothing is wrong with your file — please try again shortly, or check with your CRS administrator.";

/**
 * Is the pause ON?
 *
 * ⚠ FAILS OPEN, deliberately, and against the instinct that a safety switch should fail safe. The
 * cost of a wrong `true` is **every uploader on the site blocked** by a transient config read; the
 * cost of a wrong `false` is one file filed in the old shape, which the migration's fresh-scan guard
 * already catches and reports. Same reasoning as the stale-chain guard (gotcha 10b): a read failure
 * proves nothing, and taking the form down over one is worse than the risk it guards against.
 *
 * So only an EXPLICIT affirmative pauses. `undefined` (row missing, column absent, read failed),
 * blank, and anything unrecognised all mean "not paused".
 */
export function uploadsArePaused(settingValue: string | null | undefined): boolean {
  const v = (settingValue ?? "").trim().toLowerCase();
  if (v.length === 0) return false;
  return v === "yes" || v === "true" || v === "on" || v === "1" || v === "paused";
}

/**
 * The value written when an administrator turns the pause on or off.
 *
 * Written as words rather than a checkbox value because a human reads this row in a list view, and
 * `yes`/`no` needs no legend. Turning it OFF writes `no` rather than clearing the cell: a blank cell
 * is indistinguishable from a row nobody has ever set, and an admin checking whether they remembered
 * to switch uploads back on deserves a positive answer.
 */
export function pauseSettingValue(paused: boolean): string {
  return paused ? "yes" : "no";
}
