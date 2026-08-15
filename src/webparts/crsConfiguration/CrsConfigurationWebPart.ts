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
 * One panel today: file types. The name is a container, so a second panel becomes a sibling component
 * rather than a rewrite — but no tab framework until there is a second panel.
 */
export default class CrsConfigurationWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<FileTypeSettingsProps> = React.createElement(
      FileTypeSettings,
      {
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      },
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
