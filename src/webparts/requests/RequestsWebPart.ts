import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";


import Requests from "./components/Requests";
import { IRequestsProps } from "./components/IRequestsProps";

/**
 * CRS Requests — deletion and share requests, raised by uploaders and decided by the Head of Unit.
 *
 * Spec: docs/superpowers/specs/2026-08-15-deletion-and-share-requests-design.md
 *
 * Its own page rather than a tab on Approval Document. That page is the queue of documents awaiting
 * approval; this is a queue of decisions ABOUT documents already approved. Mixing two kinds of
 * decision on one screen is how the wrong button gets pressed.
 *
 * No web part properties. The list title, whether that list exists, which units the viewer approves
 * for and whether external sharing is permitted are all resolved at runtime — a property stored in
 * the page goes stale the moment the client renames a list.
 */
export default class RequestsWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IRequestsProps> = React.createElement(Requests, {
      context: this.context,
    });

    /* NO "Back to CRS Settings" band, deliberately — removed 2026-08-20 when the client asked
       "why is Request webpart inside CRS Settings? HOU don't have access to CRS Settings".

       They are right, and it was a real dead end: `pageAccessPolicy` classifies any page whose name
       matches /setting/ as `adminOnly`, so a Head of Unit following that link gets AccessDenied — on
       the one page they are meant to work from. This belongs with the other uploader/approver pages
       (Upload Form, Approval Document, My Submissions), none of which carry the band, and NOT with
       the admin screens the CRS Settings directory actually lists. */
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}
