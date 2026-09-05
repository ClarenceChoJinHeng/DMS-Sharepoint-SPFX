import * as React from "react";
import AccessShell from "./AccessShell";
import PageAccess from "./PageAccess";
import { IAccessProps } from "./IAccessProps";

export default function PageAccessPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Page Access"
    >
      <PageAccess context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}
