import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters,
  type IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";

const UPLOAD_PAGE =
  "/sites/SPFX-Sandbox-Testing-Ground/SitePages/Upload-Document.aspx";

export default class UploadCommandCommandSet extends BaseListViewCommandSet<Record<string, never>> {
  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters,
  ): void {
    const cmd = this.tryGetCommand("UPLOAD_DOCUMENT");
    if (cmd) {
      cmd.visible = this.context.pageContext.list?.title === "Staging";
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId === "UPLOAD_DOCUMENT") {
      const base = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, "");
      window.location.href = `${base}${UPLOAD_PAGE}`;
    }
  }
}
