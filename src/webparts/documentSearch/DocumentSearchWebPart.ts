import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import DocumentSearch from "./components/DocumentSearch";
import { IDocumentSearchProps } from "./components/IDocumentSearchProps";

export interface IDocumentSearchWebPartProps {
  pageSize: number;
}

export default class DocumentSearchWebPart extends BaseClientSideWebPart<IDocumentSearchWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDocumentSearchProps> = React.createElement(DocumentSearch, {
      context: this.context,
      // The slider cannot produce 0, but a page saved before this property existed can.
      pageSize: this.properties.pageSize > 0 ? this.properties.pageSize : 10,
    });
    /* NO "Back to CRS Settings" band, deliberately — the same exclusion the upload form, the
       approval page and My Submissions carry, and for the same reason. This web part lives on the
       HOME page and is for everyone, so a back link would assert the visitor arrived from an admin
       page they have most likely never opened and may not be permitted to open. The band belongs to
       pages the CRS Settings directory actually links to; this is not one. */
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
          header: { description: "How many results to show before 'show more'." },
          groups: [
            {
              groupName: "Results",
              groupFields: [
                PropertyPaneSlider("pageSize", {
                  label: "Results per page",
                  min: 5,
                  max: 50,
                  step: 5,
                  showValue: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
