import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { withBackToSettings } from "../../shared/backToSettings";

import FolderAccessPage from "./components/FolderAccessPage";
import { IAccessProps } from "./components/IAccessProps";

export default class FolderAccessWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IAccessProps> = React.createElement(
      FolderAccessPage,
      { context: this.context },
    );
    // Wrapped at the WEB PART boundary, not inside the component: several of these components are
    // also mounted as steps of a Folder Management guided flow, where a band offering the way OUT
    // would look like part of the flow. render() is the one place an embedded mount cannot reach.
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
