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
export function pickFullNameField(fields: SpFieldLite[]): string | undefined {
  const usable = fields.filter(
    (f) =>
      (f.Title ?? "").trim().toLowerCase() === FULL_NAME_COLUMN_TITLE.toLowerCase() &&
      // A calculated or otherwise read-only field with this title would accept the MERGE
      // request and then silently ignore the value, so it is not a candidate at all.
      !f.ReadOnlyField &&
      // Only free-text types. A Choice or Lookup named "Full Name" cannot hold an
      // arbitrary term label, and writing to it would fail per-item rather than up front.
      (f.TypeAsString === "Text" || f.TypeAsString === "Note"),
  );
  // Prefer single-line text when a site somehow has both: it is what the provisioning
  // step asks for, and it renders on one line in the details pane.
  const single = usable.find((f) => f.TypeAsString === "Text");
  return (single ?? usable[0])?.InternalName;
}
