import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentSearchProps {
  context: WebPartContext;
  /**
   * How many results to show before the "show more" control.
   *
   * A property rather than a constant because this web part sits on the HOME page, where a long
   * list pushes everything else off the screen, and may also stand alone on its own page, where it
   * should not. One placement should not have to inherit the other's shape.
   */
  pageSize: number;
  /** Empty is supported — the hero falls back to a gradient. See the web part for why this is a
   *  property and not a packaged asset. */
  heroImageUrl: string;
  heroTitle: string;
}
