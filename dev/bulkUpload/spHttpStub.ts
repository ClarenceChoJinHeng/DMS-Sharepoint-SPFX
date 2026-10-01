/* Dev-harness stub for @microsoft/sp-http.
 * esbuild aliases the real package to this file so the standalone browser
 * harness has no SharePoint runtime dependency. The component only ever uses
 * `SPHttpClient.configurations.v1` as an opaque argument (the mock client
 * ignores it) and `SPHttpClientResponse` as a type. */

export class SPHttpClient {
  public static configurations = { v1: {} as unknown };
}

export class SPHttpClientResponse {}
