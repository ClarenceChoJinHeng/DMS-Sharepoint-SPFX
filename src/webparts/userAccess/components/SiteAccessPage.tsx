import * as React from "react";
import AccessShell from "./AccessShell";
import SiteAccess from "./SiteAccess";
import { IAccessProps } from "./IAccessProps";

export default function SiteAccessPage({ context }: IAccessProps): React.ReactElement {
  return (
    <AccessShell
      title="Site Access"
      subtitle="Who may open this site at all. Without site entry, a folder grant alone only reaches that folder by direct link."
    >
      <SiteAccess context={context} siteUrl={context.pageContext.web.absoluteUrl} />
    </AccessShell>
  );
}
