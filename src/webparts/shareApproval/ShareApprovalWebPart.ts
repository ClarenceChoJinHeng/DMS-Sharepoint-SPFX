import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import ShareApproval from "./components/ShareApproval";
import { IShareApprovalProps } from "./components/IShareApprovalProps";

export default class ShareApprovalWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element = React.createElement<IShareApprovalProps>(ShareApproval, {
      context: this.context,
    });
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
