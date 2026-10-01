# SharePoint REST

- Never hardcode list or library titles. The client renames them. Use `shared/naming.ts` (`libraryTitle`, `documentsLibraryTitle`, `libApiTitle`…).
- Dates: `M/D/YYYY` for `validateUpdateListItem`, ISO for `/items` and `$filter`.
- `validateUpdateListItem` returns 200 even on failure. Check `HasException`. One unknown field fails the whole write.
- One unknown column in `$select` fails the whole read. Use a retry ladder, one rung per optional column.
- Paths go through the alias form: `GetFolderByServerRelativeUrl(@f)?@f='…'`.
- A `$top` read can't prove something is absent. Filter server-side or page it.
- Reads that decide an action are cache-busted (`Cache-Control: no-cache` + a unique URL param).
- Never MERGE `OData__ModerationStatus` together with any other field (500). Write the field first, then the status.
- Taxonomy value = `"Label|GUID"`. Compare terms by `TermGuid`: REST `/items` often returns the term's WssId (a number) in `Label`; the stored name is fine.
- Yes/No in `$filter`: compare to `1`/`0`, not `true`/`false`.
