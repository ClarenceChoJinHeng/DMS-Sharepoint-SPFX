import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAuditLogProps {
  context: WebPartContext;
  siteUrl: string;
}
