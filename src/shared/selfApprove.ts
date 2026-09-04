/**
 * The one pure decision behind auto-approving a Head of Unit's own upload: is the feature turned on
 * for this site at all.
 *
 * Everything else — whether THIS uploader holds Approve on THIS folder — is a live ACL read
 * (`probeFolderApproveAccess` in `dmsFolderMap.ts`) and cannot be pure. This function exists so that
 * one fact IS testable without a mock HTTP client.
 *
 * ── Client's request, 2026-08-24, verbatim ─────────────────────────────────────────────────────
 * "can we make it that if normal HOU or HC HOU upload, they dont need to do approval but rather it
 * will automatically approve" — for both `hou` and `hou_hc`, no persona distinction needed here: the
 * caller's `probeFolderApproveAccess` answers correctly for whichever library and role the uploader
 * actually holds, so this file needs no HC branch at all.
 *
 * ── Default is OFF, deliberately ────────────────────────────────────────────────────────────────
 * A `CRS Config` row (`Title = "autoApproveOwnUpload"`, `ConfigType = "setting"`) turns it on;
 * absent, blank or unrecognised all mean OFF. This is the codebase's usual direction for a NEW
 * behaviour change reversed: everywhere else an unreadable config fails OPEN because the cost is a
 * form taking itself out of service for a minute (`uploadsPaused`, `legallyPrivilegedFor`). Here the
 * cost of a wrong "on" is a document going live with nobody having looked at it, which is the more
 * expensive failure — so this fails CLOSED, toward the existing manual-approve behaviour, exactly
 * as `allowExternalSharing` fails closed for the same reason (a wrong "yes" sends a document outside
 * the organisation).
 */
export function shouldAttemptSelfApprove(configValue: string | undefined): boolean {
  return (configValue ?? "").trim().toLowerCase() === "yes";
}
