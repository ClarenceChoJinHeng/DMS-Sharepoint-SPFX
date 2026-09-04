// Approve or reject several documents at once, from the library's own command bar.
//
// Client, 2026-08-25: a side panel in the NATIVE library view to approve multiple files at once. The
// approval web part reviews one document at a time, which is right for reading a document and wrong
// for clearing a backlog of twenty.
//
// ⚠ THIS RUNS THE SAME CHECKS AS THE APPROVAL PAGE, from `shared/approvalGuards.ts`. A second
// approval route that skipped them would not be a convenience, it would be a faster way to do the
// damage they prevent — an unlocked destination folder, a silently replaced document, an approval
// the approver had no right to make. There is one implementation; this file must never grow its own.
//
// ⚠ IT DOES NOT REPLACE SharePoint's OWN Approve/Reject command, which sets the moderation field
// directly and is unaware any of this exists (CLAUDE.md: "THE NATIVE SHAREPOINT APPROVE COMMAND
// BYPASSES BOTH OVERWRITE GUARDS"). This narrows the reason anyone would reach for that; only the
// same check inside Auto-route closes it.
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters,
  type IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";
import { checkApproveRight } from "../../shared/approvalGuards";
import { cachedHcLibraries, libraryTitle } from "../../shared/naming";
import { primeNames } from "../../shared/spNaming";
import { openBulkApprovePanel, type SelectedDoc } from "./components/BulkApprovePanel";

const LOG_SOURCE = "CrsBulkApprove";
const COMMAND = "CRS_BULK_APPROVE";

export default class BulkApproveCommandSet extends BaseListViewCommandSet<Record<string, never>> {
  /** Resolved once — the two libraries this command is allowed to appear on. */
  private approvalTitles: string[] = [];

  /**
   * The last selection we probed, and what it answered.
   *
   * `onListViewUpdated` fires on every selection change and cannot await, so the probe runs in the
   * background and calls `raiseOnChange()` when it settles. Without this cache it would re-probe on
   * every mouse click, including the ones that change nothing.
   */
  private probedKey = "";
  private probeVerdict: "granted" | "denied" | "unknown" = "unknown";

  public async onInit(): Promise<void> {
    // Await priming HERE, in this component, rather than trusting something else to have done it.
    // `cachedListTitle` answers the LEGACY name until priming settles (the 1.0.207.0 race), and on a
    // CRS-renamed site that would leave `approvalTitles` holding a library that does not exist — so
    // the command would never appear, silently.
    await primeNames(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl).catch(
      () => undefined,
    );
    const hc = cachedHcLibraries();
    this.approvalTitles = [libraryTitle(), hc ? hc.approval.title : undefined]
      .filter((t): t is string => typeof t === "string" && t.length > 0)
      .map((t) => t.toLowerCase());
    Log.info(LOG_SOURCE, `Bulk approve armed for: ${this.approvalTitles.join(", ") || "(none)"}`);
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const cmd = this.tryGetCommand(COMMAND);
    if (!cmd) return;

    const listTitle = (this.context.pageContext.list?.title ?? "").toLowerCase();
    const rows = event.selectedRows ?? [];

    // ── Gate 1: the library. Instant, and the only check that is about WHERE you are. ───────────
    // Never on Documents, HC Documents, the archives, or anything else on the site.
    if (this.approvalTitles.indexOf(listTitle) === -1 || rows.length === 0) {
      cmd.visible = false;
      return;
    }

    // ── Gate 2: the SELECTION, not the person. ─────────────────────────────────────────────────
    //
    // There is no such thing as "an approver" site-wide: `ApproveItems` is granted per UNIT FOLDER,
    // so a Head of Unit for one unit holds it there and nowhere else. A list-level permission check
    // would therefore hide this from almost everyone, since approvers hold nothing at library scope.
    // So the question is "can you approve THESE files", and the answer changes as you move between
    // unit folders in the same library — which is correct.
    const ids = rows.map((r) => String(r.getValueByName("ID"))).sort();
    const key = `${listTitle}|${ids.join(",")}`;
    if (key !== this.probedKey) {
      this.probedKey = key;
      this.probeVerdict = "unknown";
      // Probes the FIRST selected item only. A mixed selection (some approvable, some not) still
      // shows the command, and the panel reports per file — refusing all four because one belongs to
      // another unit would be maddening on a large selection.
      const firstId = Number(rows[0].getValueByName("ID"));
      // A row whose id will not parse is not an answer — leave the verdict `unknown`, which now
      // HIDES, rather than firing `items(NaN)` and reading its 404 as "cannot approve".
      if (!isFinite(firstId)) return;
      checkApproveRight({
        sp: this.context.spHttpClient,
        webUrl: this.context.pageContext.web.absoluteUrl,
        listTitle: this.context.pageContext.list?.title ?? "",
        itemId: firstId,
      })
        .then((v) => {
          if (this.probedKey !== key) return; // selection moved on; this answer is stale
          this.probeVerdict = v;
          this.raiseOnChange();
        })
        .catch(() => undefined);
    }

    /* ⚠ ONLY a confirmed `granted` shows this. `unknown` HIDES.
       Built the other way round first — `denied` hides, `unknown` shows — on the reasoning that a
       transient read must not take the feature from a real approver, and that the approval page's
       panel follows that rule. Wrong here, and it showed on site immediately: the probe is ASYNC, so
       every selection is `unknown` for a moment, and in that window a PIC who cannot approve
       anything sees the button. The client's requirement is that non-approvers never see it, and a
       rule that is right most of the time does not satisfy that.
       The asymmetry with the approval page is deliberate. There, hiding the panel on doubt would
       strand an approver on the one screen built for approving. Here the cost of hiding is that an
       approver briefly does not see a shortcut, with the page still available — which is the cheaper
       failure by a wide margin. */
    cmd.visible = this.probeVerdict === "granted";
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== COMMAND) return;
    const listTitle = this.context.pageContext.list?.title ?? "";
    const docs: SelectedDoc[] = (event.selectedRows ?? []).map((r) => ({
      id: Number(r.getValueByName("ID")),
      name: String(r.getValueByName("FileLeafRef") ?? ""),
      fileSru: String(r.getValueByName("FileRef") ?? ""),
    }));
    openBulkApprovePanel({
      sp: this.context.spHttpClient,
      webUrl: this.context.pageContext.web.absoluteUrl,
      webSru: this.context.pageContext.web.serverRelativeUrl,
      listTitle,
      docs,
      approverEmail: (this.context.pageContext.user.email ?? "").toLowerCase(),
      // The view is re-read rather than patched: a decided document leaves the pending view, and
      // reproducing that in the DOM would be a second, drifting idea of what the list contains.
      onDone: () => window.location.reload(),
    });
  }
}
