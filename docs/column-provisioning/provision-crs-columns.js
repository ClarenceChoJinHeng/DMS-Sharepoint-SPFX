/* CRS column provisioning — BROWSER CONSOLE snippet (JavaScript, NOT PowerShell).
 *
 * Supersedes provision-dms-columns.js, deleted because it was DANGEROUS rather than merely stale:
 * it created `Year_x002f_Period` (the column is now plain `Year`) and a `Vendor` taxonomy column
 * (retired 2026-07-28 for free-text `Vendor_x002f_CustomerName`). An internal name is frozen at
 * creation, permanently, so running it would have baked in two unfixable columns.
 *
 * HOW TO RUN
 *   1. Open any page on the target site, signed in.  2. F12 -> Console.  3. Paste, Enter.
 *   4. Read the summary. RE-RUNNABLE: existing columns are skipped, missing lists are named.
 *
 * WHY A SCRIPT AT ALL
 *   An internal name is derived from the title a column is CREATED with, once, permanently, and the
 *   UI gives no way to set it. Create "Document Date" in the UI and you get `Document_x0020_Date`;
 *   the code reads `DocumentDate`. The column then looks right, reads right in the view, and every
 *   write silently drops it. `CreateFieldAsXml` sets Name/StaticName independently of DisplayName,
 *   which is exactly why shared/spColumns.ts uses it.
 *
 *   And the cost of one mistake is not one column: ONE unknown field name fails a whole
 *   `validateUpdateListItem` call, so a single misnamed column in one library loses EVERY column on
 *   that write. Auto-route separately drops any column absent at the destination, with no error and
 *   a green run. Column parity across all four libraries is not tidiness.
 *
 * THE MANIFEST IS TAKEN FROM THE LIVE TEST SITE (2026-08-17), not inferred from the code. That read
 * corrected four things a code-derived list had wrong: FolderUrl is Note (a deep path exceeds 255
 * chars), Folder Map has a `Library` column, Term Abbreviation has a `Level` column, and
 * `Documents` there carries `FullName` while its three siblings carry `Full_x0020_Name` — which is
 * why pickFullNameField matches on display title OR internal name. We create it consistently here.
 *
 * DELIBERATE DIVERGENCE FROM THE TEST SITE: Role, Scope, ConfigType, Category, Library and Level are
 * Choice columns there and TEXT here. A Choice column fails the WHOLE write when a value is absent
 * from its list, silently — and Role has gained DELS, SHARE, UPLHC and APRHC since that column was
 * made. The test site already holds both "UPLOADER" and "UPL", so it only works because someone kept
 * it fed. normalizeRoleValue accepts every form, so Text cannot fail this way and loses only a
 * dropdown nobody hand-edits. AllowedFileTypes stays MultiChoice — its Choices ARE the setting, read
 * by readAllowedFileTypesField.
 *
 * NOT COPIED FROM THE TEST SITE: Function, Estate, Company, Stage, CreditCard, Testing, Testing11 and
 * their Tid twins — experiment debris and per-segment tier columns. Tier columns for a new segment
 * are created by the Add-segment tool, in all four libraries, which is the only thing that should
 * ever create them.
 *
 * TAXONOMY BINDING CANNOT BE DONE VIA REST (verified 2026-07-27): the MERGE fails with "A type named
 * 'SP.TaxonomyField' could not be resolved by the model." The three taxonomy FIELDS are created here
 * with the correct internal names — the permanent, error-prone half — and must then be BOUND by hand:
 * Library settings -> column -> Term Set Settings -> the site-collection CRS term group.
 */
(async () => {
  const SITE = '/sites/CRS';   // <-- server-relative, no trailing slash
  const DRY_RUN = false;       // true = report what it WOULD do, write nothing

  const TEXT = 'Text', NOTE = 'Note', DATE = 'DateTime', NUM = 'Number', BOOL = 'Bool',
        MC = 'MultiChoice', TAX = 'Taxonomy';

  const FILE_TYPES = ['.pdf', '.doc', '.docx', '.xls', '.xlsx'];

  /* Every CRS library gets an IDENTICAL set — Documents included. It was missing Remark,
     LegallyPrivileged, ProjectName and Vendor/CustomerName on the test site for MONTHS, which
     silently stripped metadata from every routed document and tagged nothing on bulk upload. */
  const LIBRARY_COLUMNS = [
    ['Business_x0020_Segment', 'Business Segment', TEXT],
    ['BusinessSegmentTid', 'BusinessSegmentTid', TEXT],
    ['Department', 'Department', TEXT],
    ['DepartmentTid', 'DepartmentTid', TEXT],
    ['Unit', 'Unit', TEXT],
    ['UnitTid', 'UnitTid', TEXT],
    // SubUnit is a FIXED below-Unit tier that inherits: no group, no abbreviation, no Folder Map
    // row — just these two columns and one Levels entry. Its terms are authored UNDER the unit.
    ['SubUnit', 'SubUnit', TEXT],
    ['SubUnitTid', 'SubUnitTid', TEXT],
    // DateOnly, matching the test site. The internal name has no _x0020_ precisely because it is
    // created as "DocumentDate" and renamed afterwards — the whole point of this script.
    ['DocumentDate', 'Document Date', DATE],
    ['Vendor_x002f_CustomerName', 'Vendor/CustomerName', TEXT],
    ['ProjectName', 'ProjectName', TEXT],
    ['Remark', 'Remark', TEXT],
    ['LegallyPrivileged', 'LegallyPrivileged', BOOL],
    // On FOLDERS, not files: the term's real label behind the abbreviated folder name. Needs the
    // CRS Folder content type to reach the details pane.
    ['Full_x0020_Name', 'Full Name', TEXT],
    // Created UNBOUND. Bind each to its term set in the UI — see the header.
    ['Document_x0020_Type', 'Document Type', TAX],
    ['Year', 'Year', TAX],
    ['Confidentiality_x0020_Level', 'Confidentiality Level', TAX],
  ];

  const MANIFEST = {
    'CRS Config': [
      ['ConfigType', 'ConfigType', TEXT],
      ['SettingValue', 'SettingValue', TEXT],
      ['ModeLabel', 'ModeLabel', TEXT],
      ['Category', 'Category', TEXT],
      ['TermSetGuid', 'TermSetGuid', TEXT],
      ['StagingFolder', 'StagingFolder', TEXT],
      ['SortOrder', 'SortOrder', NUM],
      ['Levels', 'Levels', NOTE],
      ['PendingLevels', 'PendingLevels', NOTE],
      // The one Choice column in the system, and required. Nothing ticked is a deliberate hard block
      // on uploading; the column absent falls back to code defaults with an admin-only warning.
      // FillInChoice stays FALSE — enabling it restores the free-text typo risk it exists to remove.
      ['AllowedFileTypes', 'AllowedFileTypes', MC],
    ],
    // GroupId is TEXT: GroupMapWriteRow types it `string` and posts it as one.
    'CRS Group Map': [
      ['GroupId', 'GroupId', TEXT],
      ['GroupName', 'GroupName', TEXT],
      ['Segment', 'Segment', TEXT],
      ['UnitTermGuid', 'UnitTermGuid', TEXT],
      ['Role', 'Role', TEXT],
      ['Scope', 'Scope', TEXT],
      ['Target', 'Target', TEXT],
    ],
    'CRS Folder Map': [
      ['TermGuid', 'TermGuid', TEXT],
      ['FolderUniqueId', 'FolderUniqueId', TEXT],
      // NOTE, not Text. A deep folder path exceeds 255 characters, and Text would truncate the one
      // value that routes a file — silently.
      ['FolderUrl', 'FolderUrl', NOTE],
      ['Section', 'Section', TEXT],
      ['Library', 'Library', TEXT],
    ],
    'CRS Term Abbreviation': [
      // Title (built-in) holds the term's LABEL. Not decoration: it is what orphan repair matches on
      // when a term GUID dies, because the label is the only thing that survives.
      ['TermGuid', 'TermGuid', TEXT],
      ['Abbreviation', 'Abbreviation', TEXT],
      ['Level', 'Level', TEXT],
    ],
    'Approval Document': LIBRARY_COLUMNS,
    'Documents': LIBRARY_COLUMNS,
    'HC Approval Document': LIBRARY_COLUMNS,
    'HC Documents': LIBRARY_COLUMNS,
  };

  // CRS Audit Log and CRS Requests are absent on purpose — each is provisioned by a button on its own
  // page, with the exact types and indexes the code expects. Hand-making either is worse than
  // skipping it: a Choice column where the code writes free text fails every row, silently.

  const H = { Accept: 'application/json;odata=nometadata' };
  const digest = (await (await fetch(`${SITE}/_api/contextinfo`, { method: 'POST', headers: H })).json()).FormDigestValue;

  const schemaXml = (name, type) => {
    const head = `DisplayName="${name}" Name="${name}" StaticName="${name}"`;
    switch (type) {
      case NOTE: return `<Field Type="Note" ${head} NumLines="6" RichText="FALSE" />`;
      case DATE: return `<Field Type="DateTime" ${head} Format="DateOnly" />`;
      case NUM:  return `<Field Type="Number" ${head} />`;
      case BOOL: return `<Field Type="Boolean" ${head}><Default>0</Default></Field>`;
      case TAX:  return `<Field Type="TaxonomyFieldType" ${head}></Field>`;
      case MC:   return `<Field Type="MultiChoice" ${head} FillInChoice="FALSE"><CHOICES>` +
                        FILE_TYPES.map((c) => `<CHOICE>${c}</CHOICE>`).join('') + '</CHOICES></Field>';
      default:   return `<Field Type="Text" ${head} MaxLength="255" />`;
    }
  };

  const out = { created: [], kept: [], failed: [], missingLists: [], toBind: [] };

  for (const [listTitle, columns] of Object.entries(MANIFEST)) {
    const base = `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;

    const check = await fetch(`${base}/fields?$select=InternalName&$top=500`, { headers: H });
    if (!check.ok) {
      // Absent is normal on a first pass — the libraries come later. NAMED, never silent: a skipped
      // list is otherwise indistinguishable from a completed one.
      out.missingLists.push(`${listTitle} (HTTP ${check.status})`);
      continue;
    }
    const have = new Set(((await check.json()).value || []).map((f) => (f.InternalName || '').toLowerCase()));

    for (const [internal, display, type] of columns) {
      // Matched on internal name alone, whatever the TYPE. A column already present is never
      // converted: a type change on a populated column loses its data.
      if (have.has(internal.toLowerCase())) { out.kept.push(`${listTitle} · ${internal}`); continue; }
      if (DRY_RUN) { out.created.push(`${listTitle} · ${internal} (dry run)`); continue; }

      const res = await fetch(`${base}/fields/CreateFieldAsXml`, {
        method: 'POST',
        headers: { ...H, 'Content-Type': 'application/json', 'X-RequestDigest': digest },
        // Options 8 = AddToAllContentTypes. NOT 12, which also adds each column to the default VIEW
        // and would reshape every view the client arranged, once per column.
        body: JSON.stringify({ parameters: { SchemaXml: schemaXml(internal, type), Options: 8 } }),
      });
      if (!res.ok) {
        out.failed.push(`${listTitle} · ${internal} — HTTP ${res.status} ${(await res.text()).slice(0, 140)}`);
        continue;
      }
      out.created.push(`${listTitle} · ${internal}`);
      if (type === TAX) out.toBind.push(`${listTitle} · ${display}`);

      // Rename to the human title AFTER creation. The internal name is already frozen, so this only
      // changes what is displayed — and a failure here is cosmetic and must never be reported as a
      // failure of the column itself.
      if (display !== internal) {
        await fetch(`${base}/fields/getbyinternalnameortitle('${encodeURIComponent(internal)}')`, {
          method: 'POST',
          headers: { ...H, 'Content-Type': 'application/json', 'X-RequestDigest': digest,
                     'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
          body: JSON.stringify({ Title: display }),
        }).catch(() => undefined);
      }
    }
  }

  console.log(`CREATED ${out.created.length}`);           if (out.created.length) console.table(out.created);
  console.log(`ALREADY THERE ${out.kept.length}`);
  if (out.missingLists.length) { console.warn('LISTS NOT FOUND — run again once they exist:'); console.table(out.missingLists); }
  if (out.failed.length)       { console.error('FAILED:'); console.table(out.failed); }
  if (out.toBind.length) {
    console.warn('BIND THESE TO THEIR TERM SET BY HAND (REST cannot do it) — library settings -> column -> Term Set Settings:');
    console.table(out.toBind);
  }
  console.log('Then VERIFY: diff /fields?$select=Title,InternalName across all four libraries. '
    + 'A matching display name over a DIFFERENT internal name fails exactly like an absent column, and looks right.');
})();
