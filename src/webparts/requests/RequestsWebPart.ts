import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { withBackToSettings } from "../../shared/backToSettings";

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
    // Wrapped at the web part boundary — see withBackToSettings.
    ReactDom.render(withBackToSettings(this.context, element), this.domElement);
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
