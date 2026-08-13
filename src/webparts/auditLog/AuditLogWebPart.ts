import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import AuditLog from "./components/AuditLog";
import { IAuditLogProps } from "./components/IAuditLogProps";

/**
 * CRS Audit Log — what happened in the DMS, and who did it.
 *
 * Spec: docs/superpowers/specs/2026-08-13-audit-log-design.md
 *
 * Its own web part rather than a tab on CRS Configuration: a log is not a setting, and it needs its
 * own page so page-level permissions can restrict it to administrators.
 *
 * No web part properties. Everything it needs — the list title, whether that list exists, whether
 * the viewer is an admin — is resolved at runtime, because a property stored in the page would go
 * stale the moment the client renames a list.
 */
export default class AuditLogWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IAuditLogProps> = React.createElement(AuditLog, {
      context: this.context,
      siteUrl: this.context.pageContext.web.absoluteUrl,
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
    return { pages: [] };
  }
}
