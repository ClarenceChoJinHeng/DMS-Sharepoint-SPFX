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
  /**
   * Fired when the New segment tab actually writes a segment, with its derived key and label.
   *
   * Passed straight through to `SegmentCreator`. It exists so the guided flow can close the creation
   * form and select the new segment on the creation ITSELF, rather than inferring it from a re-read plus
   * a question to the admin — which is why the form used to sit open underneath its own success message.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onSegmentCreated?: (key: string, label: string) => void;
  /**
   * Hide **Delete** on the New segment tab's list of existing segments.
   *
   * Set by the "Add a new segment" flow only. NOT keyed off `hideTabs`, because the **Retire** flow mounts
   * this same tab with `hideTabs` set and Delete is the whole point of that step — so the two cases must
   * be told apart explicitly. The list itself is never hidden: seeing what exists is what stops someone
   * creating a segment twice.
   */
  hideSegmentDelete?: boolean;
  /**
   * Passed through to `AbbreviationManager`: how many terms have no folder code, or `undefined` when that
   * is not knowable.
   *
   * The guided flow gates its Next button on this and cannot work it out itself — the count needs a walk
   * of the whole term tree (~115 requests for GHO), which the abbreviations screen has already paid for.
   * Ignored by `FolderMap`, which shares this interface.
   */
  onAbbreviationsMissingChange?: (missing: number | undefined) => void;
  /**
   * True while the abbreviations screen is reading — the segment list or the term tree.
   *
   * The count alone cannot answer this: it is `undefined` both while loading and when the read
   * failed, and only the first should hold the flow's Next button. Ignored by `FolderMap`, which
   * shares this interface.
   */
  onAbbreviationsLoadingChange?: (loading: boolean) => void;
}
