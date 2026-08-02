// Folder names are abbreviations (GHO/GF/TREASURY), short enough to keep the
// server-relative path under the ~330-character ceiling where GetFolderByServerRelativeUrl
// starts returning 400, but unreadable on their own. The full term label is written to a
// "Full Name" column on each folder's list item so the details pane still shows what the
// folder actually is.
//
// The column's INTERNAL name is discovered at runtime rather than assumed. CLAUDE.md is
// explicit that internal names must never be guessed from display names: SharePoint's
// encoding depends on how the column was created — `Full_x0020_Name` if it was created as
// "Full Name", but plain `FullName` if it was created without the space and renamed
// afterwards, a very common path when an admin fixes a typo. Both are legitimate and only
// the live /fields response can say which one this site has.

/** Display name the client is asked to create on Staging and Documents. */
export const FULL_NAME_COLUMN_TITLE = "Full Name";

/** The subset of an SP field entry this module needs. */
export interface SpFieldLite {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  ReadOnlyField: boolean;
}

/**
 * Resolve the internal name of the Full Name column from a library's /fields response.
 *
 * Returns undefined when the column is absent or unusable, which is a soft state: the
 * reconciliation run carries on and simply does not write a full name. A missing display
 * column must never abort provisioning — the folders and their ACLs are the thing that
 * matters, and this is a label.
 */
/**
 * Reduce a display or internal name to a comparable key: strip SharePoint's `_x0020_`
 * space encoding, drop every non-alphanumeric character, lowercase.
 *
 * So "Full Name", "FullName", "full name" and "Full_x0020_Name" all collapse to
 * "fullname". Two people creating the same column on two libraries do not necessarily
 * type the same spacing — that is exactly what happened on 2026-08-02, where Staging
 * got "Full Name" and Documents got "FullName", and an exact-match comparison filled
 * one library while silently skipping ~190 folders in the other. The variance is not
 * worth a support round-trip, and nothing else plausibly collapses to this key.
 */
function nameKey(s: string): string {
  return (s ?? "").replace(/_x0020_/gi, "").replace(/[^a-z0-9]/gi, "").toLowerCase();
}

export function pickFullNameField(fields: SpFieldLite[]): string | undefined {
  const wanted = nameKey(FULL_NAME_COLUMN_TITLE);
  const usable = fields.filter(
    (f) =>
      (nameKey(f.Title) === wanted || nameKey(f.InternalName) === wanted) &&
      // A calculated or otherwise read-only field with this title would accept the MERGE
      // request and then silently ignore the value, so it is not a candidate at all.
      !f.ReadOnlyField &&
      // Only free-text types. A Choice or Lookup named "Full Name" cannot hold an
      // arbitrary term label, and writing to it would fail per-item rather than up front.
      (f.TypeAsString === "Text" || f.TypeAsString === "Note"),
  );
  if (usable.length === 0) return undefined;
  if (usable.length === 1) return usable[0].InternalName;

  // More than one candidate. Adding a content type brings its own column, so a library
  // can end up with "FullName" (created by hand) beside "Full Name" (from the content
  // type) — both Text, both writable, both matching the relaxed key. Live 2026-08-03:
  // `FullName` held the data while `FullName0`/"Full Name" was the one the details pane
  // rendered, so every folder looked empty.
  //
  // Resolve on the EXACT documented title. A list cannot hold two columns with the same
  // display name, so this matches at most one, and it picks the column the client was
  // told to create — the one their content type and details pane are bound to.
  const exact = usable.find(
    (f) => (f.Title ?? "").trim().toLowerCase() === FULL_NAME_COLUMN_TITLE.toLowerCase(),
  );
  if (exact) return exact.InternalName;

  // Several near-names and no exact title: there is no principled way to choose, and
  // guessing writes hundreds of folders to a column that may be the wrong one. Refuse —
  // the caller lists the candidates so a person can delete the stray column.
  return undefined;
}
