import * as React from "react";
import AccessShell from "./AccessShell";
import PageAccess from "./PageAccess";
import { IAccessProps } from "./IAccessProps";

export default function PageAccessPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Page Access"
      subtitle="Who may open each page on this site. Each page is granted separately."
    >
      <PageAccess context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}
