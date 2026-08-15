import { WebPartContext } from "@microsoft/sp-webpart-base";

/**
 * Context only.
 *
 * Everything else this screen needs — the list title, the viewer's approver units, whether external
 * sharing is allowed — is resolved at runtime. A property stored in the page goes stale the moment
 * the client renames a list or moves someone between units.
 */
export interface IRequestsProps {
  context: WebPartContext;
}
