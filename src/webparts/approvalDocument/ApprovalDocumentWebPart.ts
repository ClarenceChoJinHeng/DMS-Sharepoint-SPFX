import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import ApprovalDocument from "./components/ApprovalDocument";
import { IApprovalDocumentProps } from "./components/IApprovalDocumentProps";

export default class ApprovalDocumentWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    this.widenCanvas();
    const element = React.createElement<IApprovalDocumentProps>(ApprovalDocument, {
      context: this.context,
    });
    ReactDom.render(element, this.domElement);
  }

  /**
   * Let the page use the whole browser width.
   *
   * A modern SharePoint page caps its canvas at 1204px, so on a wide monitor the
   * approval screen leaves a large empty margin to the right of the panel no matter
   * what the web part's own CSS does — the ceiling lives outside our DOM. The section
   * elements are the only thing carrying that cap, so lift it on the ancestors we
   * sit inside.
   *
   * Deliberately narrow: it only ever RAISES a max-width, never sets a width, so the
   * page reflows normally and a narrow window is unaffected. Everything is
   * null-guarded and matched loosely — these are Microsoft's class names and they are
   * free to change them, in which case the page falls back to 1204px rather than
   * breaking.
   *
   * The alternative is asking the client to convert the section to a full-width
   * section in the page editor: a per-page manual step, on a screen they have to be
   * walked to, that gets skipped.
   */
  private widenCanvas(): void {
    const CANVAS_CLASSES = ["CanvasZone", "CanvasSection", "CanvasZoneContainer"];
    let node: HTMLElement | null = this.domElement;
    while (node && node !== document.body) {
      const cls = node.className;
      if (typeof cls === "string") {
        for (let i = 0; i < CANVAS_CLASSES.length; i++) {
          if (cls.indexOf(CANVAS_CLASSES[i]) !== -1) {
            node.style.maxWidth = "none";
            break;
          }
        }
      }
      node = node.parentElement;
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
