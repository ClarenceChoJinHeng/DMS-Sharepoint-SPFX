import { WebPartContext } from "@microsoft/sp-webpart-base";

/** Shared by all four access web parts — they differ only in which surface they render. */
export interface IAccessProps {
  context: WebPartContext;
}
