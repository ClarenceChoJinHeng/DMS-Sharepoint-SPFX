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
import { withBackToSettings } from "../../shared/backToSettings";

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
    /* Wrapped at the WEB PART boundary, never inside the component — the same seam every other CRS
       page uses. It is load-bearing: a component that renders its own back band would carry it into
       any future embedded mount, where a link OUT of the host page reads as part of the host. */
    ReactDom.render(withBackToSettings(this.context, element), this.domElement);
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
