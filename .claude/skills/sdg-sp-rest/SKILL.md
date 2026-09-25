---
name: sdg-sp-rest
description: Use when writing or changing code that reads or writes SharePoint lists, libraries, fields, groups or permissions through REST in this project.
---
# SharePoint REST, the house way

Hard rules are in `rules/sharepoint-rest.md`. This is how to apply them.

1. **Names**: get list/library titles from `src/shared/naming.ts` (`libraryTitle()`, `documentsLibraryTitle()`, `libApiTitle(key)`, `cachedListTitle(LIST_SUFFIX.x)`). Call `await primeNames(...)` (`spNaming.ts`) inside your own reader first.
2. **Reads**:
   - Reuse existing readers before writing new ones (grep `src/shared/sp*.ts`).
   - Optional columns: a `$select` ladder. Try with the column, and on **400 only** retry without it. One rung per column.
   - Decides an action? Add `Cache-Control: no-cache` and a unique URL param.
   - Failure returns `undefined`, never `[]`.
3. **Writes**:
   - Metadata → `validateUpdateListItem`, dates `M/D/YYYY`, check each `HasException`.
   - Item fields → MERGE with `IF-MATCH: *`. Never in the same MERGE as `OData__ModerationStatus`.
   - Columns that may not exist → check `libraryHasColumns` (`optionalColumns.ts`) first.
4. **Paths**: `GetFolderByServerRelativeUrl(@f)?@f='<encoded>'`.
5. **Throttling**: wrap in `withThrottleRetry` (`throttleRetry.ts`).
6. **Permissions checks**: ask the item (`EffectiveBasePermissions`, see the probes in `dmsFolderMap.ts`). Never rebuild the answer from the Group Map alone.
7. **Deletes and tagging** never happen directly. Write the `CRS Requests` / `CRS Submissions` row and let the flow do it.
8. Put pure logic in `src/shared/<x>.ts` with a test. Put the fetch calls in `sp<X>.ts`.
