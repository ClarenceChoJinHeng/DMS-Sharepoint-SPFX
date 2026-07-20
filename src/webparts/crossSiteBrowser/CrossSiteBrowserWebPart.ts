import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import CrossSiteBrowser from "./components/CrossSiteBrowser";
import { ICrossSiteBrowserProps } from "./components/ICrossSiteBrowserProps";

// TEMPORARY proof-of-concept web part. Remove once the multi-site view
// approach is finalised (delete this folder + the config.json bundle entry +
// the componentId from package-solution.json).
export interface ICrossSiteBrowserWebPartProps {
  targetSiteUrl: string;
  libraryName: string;
}

export default class CrossSiteBrowserWebPart extends BaseClientSideWebPart<ICrossSiteBrowserWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ICrossSiteBrowserProps> = React.createElement(
      CrossSiteBrowser,
      {
        context: this.context,
        targetSiteUrl: (this.properties.targetSiteUrl ?? "").trim(),
        libraryName: (this.properties.libraryName ?? "").trim() || "Documents",
      },
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
    return {
      pages: [
        {
          header: {
            description:
              "Segments come from the 'DMS Site Map' list. The fields below are an optional fallback used only when that list is missing/empty.",
          },
          groups: [
            {
              groupName: "Fallback target (optional)",
              groupFields: [
                PropertyPaneTextField("targetSiteUrl", {
                  label: "Fallback site URL (absolute)",
                  placeholder: "https://tenant.sharepoint.com/sites/Upstream",
                }),
                PropertyPaneTextField("libraryName", {
                  label: "Fallback library name",
                  placeholder: "Documents",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
