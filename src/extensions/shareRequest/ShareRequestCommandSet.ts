import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters,
  type IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";

const REQUEST_SHARE_PAGE = "/SitePages/Request-Share.aspx";

export default class ShareRequestCommandSet extends BaseListViewCommandSet<Record<string, never>> {
  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const cmd = this.tryGetCommand("REQUEST_SHARE");
    if (!cmd) return;
    const lib = this.context.pageContext.list?.title;
    // Visible on Staging/Documents, exactly one row selected.
    cmd.visible = (lib === "Staging" || lib === "Documents") && event.selectedRows.length === 1;
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== "REQUEST_SHARE" || event.selectedRows.length !== 1) return;
    const row = event.selectedRows[0];
    const itemUrl = row.getValueByName("FileRef") as string;          // server-relative path
    const uniqueId = String(row.getValueByName("UniqueId") ?? "").replace(/[{}]/g, "");
    const fsObjType = row.getValueByName("FSObjType");                 // "1" = folder
    const itemType = String(fsObjType) === "1" ? "Folder" : "File";
    const library = this.context.pageContext.list?.title === "Staging" ? "Staging" : "Documents";
    const base = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, "");
    const q = `itemUrl=${encodeURIComponent(itemUrl)}&itemId=${encodeURIComponent(uniqueId)}`
      + `&itemType=${itemType}&library=${library}`;
    window.location.href = `${base}${REQUEST_SHARE_PAGE}?${q}`;
  }
}
