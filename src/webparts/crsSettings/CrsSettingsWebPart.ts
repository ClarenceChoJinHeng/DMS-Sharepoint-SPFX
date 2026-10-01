import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { allLinks, suggestedPageName } from "../../shared/adminPages";
import CrsSettings, { CrsSettingsProps } from "./components/CrsSettings";

/**
 * CRS Settings — the landing page for every admin screen.
 *
 * Spec: docs/superpowers/specs/2026-08-14-crs-settings-landing-page-design.md
 *
 * A directory, at the client's decision. Links resolve themselves against the site's own Site Pages,
 * so nothing here hardcodes a page name — the client renames everything at import, and a hardcoded
 * name fails as a dead link on the one page whose job is telling an admin where to go.
 */
export interface ICrsSettingsWebPartProps {
  heading: string;
  subheading: string;
  /**
   * Per-link overrides, stored flat as `link_<key>` rather than as a nested object.
   *
   * SPFx persists web part properties as JSON and the property pane binds one field to one path, so a
   * nested `overrides` object would need a custom accessor per key for no benefit. Flat keys also
   * survive a link being added or removed: an orphaned property is inert, where a stale nested shape
   * would have to be migrated.
   */
  [key: string]: string;
}

export default class CrsSettingsWebPart extends BaseClientSideWebPart<ICrsSettingsWebPartProps> {
  /** The `link_<key>` properties, gathered into what the component expects. */
  private overrides(): Record<string, string> {
    const out: Record<string, string> = {};
    for (const l of allLinks()) {
      const v = this.properties[`link_${l.key}`];
      if (typeof v === "string" && v.trim().length > 0) out[l.key] = v.trim();
    }
    return out;
  }

  public render(): void {
    const element: React.ReactElement<CrsSettingsProps> = React.createElement(CrsSettings, {
      context: this.context,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      overrides: this.overrides(),
      heading: this.properties.heading,
      subheading: this.properties.subheading,
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  /**
   * Heading, then one optional address per link.
   *
   * The overrides exist so a site whose page names cannot be guessed is fixable without a redeploy.
   * Each placeholder names the page auto-detection is looking for, so the pane doubles as the answer
   * to "why is this row greyed out".
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "The heading, and where each link points." },
          groups: [
            {
              groupName: "Heading",
              groupFields: [
                PropertyPaneTextField("heading", { label: "Title", placeholder: "CRS Settings" }),
                PropertyPaneTextField("subheading", {
                  label: "Subtitle",
                  placeholder: "Manage repository settings, folders, access, and mappings.",
                }),
              ],
            },
            {
              groupName: "Link addresses (optional)",
              groupFields: allLinks().map((l) =>
                PropertyPaneTextField(`link_${l.key}`, {
                  label: l.label,
                  placeholder: `auto — looks for ${suggestedPageName(l)}`,
                }),
              ),
            },
          ],
        },
      ],
    };
  }
}
