import * as React from "react";
import AccessShell from "./AccessShell";
import PageAccess from "./PageAccess";
import { IAccessProps } from "./IAccessProps";

export default function PageAccessPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      className="crs-pa-shell"
      title="Page Access"
    >
      {/* Client, 2026-09-27: no shell padding on this page. important beats the inline wrap style. */}
      <style>{".crs-pa-shell { padding: 0 !important; }"}</style>
      <PageAccess context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}
