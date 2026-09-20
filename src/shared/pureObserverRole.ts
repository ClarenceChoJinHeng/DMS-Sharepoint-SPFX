import { normalizeRoleValue } from "./groupMapModel";

/**
 * Does this viewer hold ANY role, anywhere in the Group Map, that could ever legitimately justify
 * offering a delete or share action — direct or via a request?
 *
 * Client, 2026-09-21, on `Document-Viewer.aspx` offering "Request deletion"/"Request share" to a
 * Viewer account: "remove the Request Share and Request deletion for Viewer and C Level only."
 *
 * ⚠ DERIVED, NEVER A HARDCODED "Viewer or C-Level" PERSONA LIST. Under the current persona model
 * (`groupMapModel.ts`'s `PERSONAS`), Viewer (`employee`/`employee_hc` = MEMBER/MEMBERHC) and C-Level
 * (`clevel_global`/`clevel_segment` = GLOBAL/SEGVIEW) are the ONLY personas holding none of
 * UPL/UPLHC/APR/APRHC/DEL/DELHC/SHARE/SHAREHC/DELS/DELSHC/DEPTVIEW — so checking for the ABSENCE of
 * every acting role is functionally identical to "is Viewer or C-Level" today, and stays correct
 * automatically if the persona model changes again (it already has, several times, in this project).
 * A hardcoded name list would be a second definition of the same fact, free to drift from
 * `PERSONAS` the way several other things in this codebase already have.
 *
 * ⚠ `DEPTVIEW` IS IN THE ACTING SET, THOUGH IT LOOKS LIKE A PURE-VIEW ROLE BY NAME. A Head of
 * Department's persona is `[DEPTVIEW, SHARE, SHAREHC]` — DEPTVIEW never travels alone under the
 * current model, so a `DEPTVIEW` row is proof the account also carries `SHARE`/`SHAREHC` somewhere
 * in the same Group Map, and excluding it would wrongly treat a genuine HOD as a pure observer the
 * moment `SHARE`'s own row happens to sort after `DEPTVIEW`'s in the response. Including it costs
 * nothing when it is wrong (a stray pre-2026-08-17 DEPTVIEW-only row simply keeps the buttons
 * visible for someone who cannot use them, which fails toward showing rather than hiding).
 */
const ACTING_ROLES = new Set([
  "UPL",
  "UPLHC",
  "APR",
  "APRHC",
  "DEL",
  "DELHC",
  "SHARE",
  "SHAREHC",
  "DELS",
  "DELSHC",
  "DEPTVIEW",
]);

export function isPureObserverRole(
  myGroupIds: readonly string[],
  mapRows: ReadonlyArray<{ GroupId?: string; Role?: string }>,
): boolean {
  const mine = new Set(myGroupIds.map((id) => String(id)));
  for (const r of mapRows) {
    if (!mine.has(String(r.GroupId ?? ""))) continue;
    if (ACTING_ROLES.has(normalizeRoleValue(r.Role ?? ""))) return false;
  }
  return true;
}
