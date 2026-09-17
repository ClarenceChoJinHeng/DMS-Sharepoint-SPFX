import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { withBackToSettings } from "../../shared/backToSettings";

import FileTypeSettings, { FileTypeSettingsProps } from "./components/FileTypeSettings";

/**
 * CRS Configuration — settings that live on the `CRS Config` list.
 *
 * Spec: docs/superpowers/specs/2026-08-13-file-type-settings-page-design.md
 *
 * ⚠ A second panel (Bulk Upload access, 2026-09-13) was built and mounted here the same day the
 * client reviewed it and said "no need" — removed the same day, per
 * `docs/superpowers/specs/2026-09-13-bulk-upload-access-toggle-design.md`, which is kept as the
 * record rather than deleted. This web part is back to the one panel that shipped: file types.
 */
export default class CrsConfigurationWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const fileTypes: React.ReactElement<FileTypeSettingsProps> = React.createElement(
      FileTypeSettings,
      { context: this.context, siteUrl },
    );
    // Wrapped at the WEB PART boundary, not inside the component: several of these components are
    // also mounted as steps of a Folder Management guided flow, where a band offering the way OUT
    // would look like part of the flow. render() is the one place an embedded mount cannot reach.
    ReactDom.render(
      withBackToSettings(this.context, fileTypes),
      this.domElement,
    );
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
