import { normalizeRoleValue } from "./groupMapModel";

/**
 * Which of three destinations should the CURRENT viewer land on when a column-formatting link
 * (Documents/HC Documents/Archive/HC Archive) sends them to the shared Viewer/C-Level/Head of
 * Department page?
 *
 * Spec/context: 2026-09-21, client's own words - "Ensure only PIC can open My Submission and only
 * Approver can open File Permission... Viewer, C Level, HOD [get] a separate page."
 *
 * SHAREPOINT COLUMN FORMATTING CANNOT ROUTE BY VIEWER ROLE - it only ever sees the ROW's own data
 * (filename, UniqueId), never who is looking at it. So every click lands on ONE fixed page, and
 * THAT page is what decides, on load, whether to silently redirect the viewer elsewhere before
 * anything renders. This function is the decision; the redirect itself is the caller's job.
 *
 * `myGroupIds` is the viewer's OWN SharePoint group ids (`currentuser/groups`) - never every
 * group on the site, or every unit would look like this person's.
 *
 * ORDER MATTERS: APR/APRHC is checked BEFORE UPL/UPLHC. A Head of Unit's own persona has carried
 * `UPL` as one of its roles since 2026-09-17 (so they can upload too), so an Approver's own group
 * memberships legitimately include an uploader-role row - checking uploader first would send an
 * Approver to My Submissions instead of File Permission.
 */
export function classifyViewerForFileRoute(
  myGroupIds: readonly string[],
  mapRows: ReadonlyArray<{ GroupId?: string; Role?: string }>,
): "approver" | "uploader" | "other" {
  const mine = new Set(myGroupIds.map((id) => String(id)));
  let sawUploader = false;
  for (const r of mapRows) {
    if (!mine.has(String(r.GroupId ?? ""))) continue;
    const role = normalizeRoleValue(r.Role ?? "");
    if (role === "APR" || role === "APRHC") return "approver";
    if (role === "UPL" || role === "UPLHC") sawUploader = true;
  }
  return sawUploader ? "uploader" : "other";
}
