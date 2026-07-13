import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import Reconciliation from "./components/Reconciliation";
import { IReconciliationProps } from "./components/IReconciliationProps";

export default class ReconciliationWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IReconciliationProps> = React.createElement(
      Reconciliation,
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
