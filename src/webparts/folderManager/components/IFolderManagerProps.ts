import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFolderManagerProps {
  context: WebPartContext;
  /**
   * Which tab to open on, as a deep-link slug (`abbreviations`, `structure`, `migrate`,
   * `reconciliation`, `newsegment`).
   *
   * Added 2026-08-14 so the guided flows in `FolderAdmin` can drive this component one step at a time
   * without reconciliation being extracted from it. Absent = the URL hash decides, then Term
   * Abbreviations. See `DEEP_LINK_TABS` in FolderManager for the slug contract.
   *
   * Ignored by `FolderMap`, which shares this interface.
   */
  initialTab?: string;
  /**
   * Hide the tab bar and the page heading.
   *
   * Set when a flow is showing one step: the flow's own rail is the navigation, and a second row of tabs
   * beside it would offer a way out of the flow that looks like part of it.
   */
  hideTabs?: boolean;
}
