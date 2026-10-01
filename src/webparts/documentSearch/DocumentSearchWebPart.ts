import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import DocumentSearch from "./components/DocumentSearch";
import { IDocumentSearchProps } from "./components/IDocumentSearchProps";

export interface IDocumentSearchWebPartProps {
  pageSize: number;
  /**
   * The hero photograph behind the search bar.
   *
   * ⚠ A PROPERTY RATHER THAN A PACKAGED ASSET, and that is not a preference. Provisioning files
   * with this solution DOES NOT WORK on this tenant — `elements.xml` has failed to provision twice
   * (the `+ New Folder` customizer, then the bulk-approve command set), both times silently. An
   * image shipped in the package would render as a broken box with nothing explaining it. Point
   * this at something already uploaded to **Site Assets**.
   *
   * Blank is a supported state: the hero falls back to a green gradient, which reads as deliberate
   * rather than as a failed image.
   */
  heroImageUrl: string;
  /** The hero heading. Free text so the client can retitle the site without a redeploy. */
  heroTitle: string;
}

export default class DocumentSearchWebPart extends BaseClientSideWebPart<IDocumentSearchWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDocumentSearchProps> = React.createElement(DocumentSearch, {
      context: this.context,
      // The slider cannot produce 0, but a page saved before this property existed can.
      pageSize: this.properties.pageSize > 0 ? this.properties.pageSize : 10,
      /* Defaults to the client's banner in Site Assets, SITE-RELATIVE so the same package works on
         ClarenceDMSTesting and on SDG without a per-site literal. The property still wins when set.
         The file is NOT packaged: provisioning image assets has failed silently on this tenant twice
         (the `+ New Folder` customizer, then the bulk-approve command set), and a 333 KB JPEG inlined
         as base64 would add ~445 KB to every page load of this bundle. Upload it once to Site Assets
         instead. If it is missing the banner falls back to the gradient, exactly as before. */
      heroImageUrl:
        (this.properties.heroImageUrl ?? "").trim() ||
        `${this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, "")}/SiteAssets/img_kv-banner.jpg`,
      // A page saved before this property existed has neither; the default is the client's wording.
      heroTitle: (this.properties.heroTitle ?? "").trim() || "Welcome to Guthrie Repository",
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
          header: { description: "The banner, and how many results to show before 'show more'." },
          groups: [
            {
              groupName: "Banner",
              groupFields: [
                PropertyPaneTextField("heroTitle", {
                  label: "Heading",
                  placeholder: "Welcome to Guthrie Repository",
                }),
                PropertyPaneTextField("heroImageUrl", {
                  label: "Background image URL",
                  // Says where to put the file, because the usual instinct — attach an image to the
                  // web part — is not available here.
                  description:
                    "Upload the image to Site Assets and paste its address. Leave blank for a plain green banner.",
                }),
              ],
            },
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
