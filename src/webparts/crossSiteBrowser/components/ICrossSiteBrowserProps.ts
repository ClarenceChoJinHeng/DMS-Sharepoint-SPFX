import { WebPartContext } from "@microsoft/sp-webpart-base";

// TEMPORARY proof-of-concept web part — cross-site folder/file browser.
// Remove once the multi-site view approach is finalised.
export interface ICrossSiteBrowserProps {
  context: WebPartContext;
  targetSiteUrl: string; // absolute URL of the segment site to read from
  libraryName: string; // e.g. "Staging" or "Documents"
}
