import * as React from "react";
import { useEffect, useState } from "react";
import AccessShell from "./AccessShell";
import StagingAccess from "./StagingAccess";
import { IAccessProps } from "./IAccessProps";
import { primeNames } from "../../../shared/spNaming";
import { libraryTitle } from "../../../shared/naming";

export default function ApprovalLibraryAccessPage({ context }: IAccessProps): React.ReactElement {
  const siteUrl = context.pageContext.web.absoluteUrl;
  // The heading is the LIVE library title ("Approval Document"), never the logical key.
  // Hardcoding "Staging" here would be the one place a client still sees the retired name.
  const [libLabel, setLibLabel] = useState<string>(libraryTitle());

  useEffect(() => {
    primeNames(context.spHttpClient, siteUrl)
      .catch(() => undefined)
      .then(() => setLibLabel(libraryTitle()))
      .catch(() => undefined);
  }, []);

  return (
    <AccessShell
      title={`${libLabel} Access`}
      subtitle={`Who may open the ${libLabel} library at all. Uploaders and approvers need this on top of their folder grant.`}
    >
      {/* "Staging" is the LOGICAL LibTarget key, mapped to the live title at the API
          boundary by libApiTitle — not the library's name. */}
      <StagingAccess context={context} siteUrl={siteUrl} library="Staging" />
    </AccessShell>
  );
}
