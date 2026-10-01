import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBulkUploadProps {
  context: WebPartContext;
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}
