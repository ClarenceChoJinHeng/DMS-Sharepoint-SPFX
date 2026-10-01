import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import FolderStructurePage from "./components/FolderStructurePage";
import { IAccessProps } from "./components/IAccessProps";

export default class FolderStructureWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IAccessProps> = React.createElement(
      FolderStructurePage,
      { context: this.context },
    );
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
