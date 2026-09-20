import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import MySubmissions from "./components/MySubmissions";
import { IMySubmissionsProps } from "./components/IMySubmissionsProps";

export interface IMySubmissionsWebPartProps {
  viewerOnlyMode?: boolean;
}

export default class MySubmissionsWebPart extends BaseClientSideWebPart<IMySubmissionsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IMySubmissionsProps> = React.createElement(MySubmissions, {
      context: this.context,
      viewerOnlyMode: this.properties.viewerOnlyMode === true,
    });
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
          header: { description: "Viewer / C-Level / Head of Department page (2026-09-21)" },
          groups: [
            {
              groupName: "Mode",
              groupFields: [
                PropertyPaneCheckbox("viewerOnlyMode", {
                  text:
                    "This instance is the Viewer/C-Level/HOD page — hides the personal " +
                    "submissions list, and silently redirects a PIC or Approver who lands here to " +
                    "their own page instead. Leave unticked on the real My Submissions page.",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
