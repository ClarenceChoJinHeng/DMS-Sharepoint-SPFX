import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMySubmissionsProps {
  context: WebPartContext;
  /**
   * A SECOND MOUNT of this same web part (2026-09-21), for Viewer/C-Level/Head of Department —
   * roles that browse Documents/HC Documents/Archive/HC Archive but hold no upload rights, so they
   * have nothing of their own to list here. When true:
   *  - the personal submissions list/tabs is never loaded or rendered — there is nothing to load;
   *  - on mount, the viewer's own role is checked, and a PIC or Approver landing here (via the SAME
   *    column-formatting link, since SharePoint cannot route by viewer role) is silently redirected
   *    to their OWN real page (My Submissions / File Permission) before anything renders;
   *  - everyone else (Viewer, C-Level, HOD) simply stays and sees the ?file=<UniqueId> detail view,
   *    with Delete/Share offered exactly where probeFileRights says they may act — HOD's own SHARE
   *    grant on their department's approved documents falls out of this for free, with no
   *    special-casing: the same check that grants Approver a Share button grants HOD one too.
   *  - Omitted/false means the ORIGINAL page, unchanged: the viewer's own submissions list, plus the
   *    existing "any document, not just mine" fallback added the same day.
   */
  viewerOnlyMode?: boolean;
}
