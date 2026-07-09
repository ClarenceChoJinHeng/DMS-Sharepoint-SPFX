import * as React from "react";
import FolderManager from "./FolderManager";
import { IFolderManagerProps } from "./IFolderManagerProps";

export default function FolderAdmin({ context }: IFolderManagerProps): React.ReactElement {
  return <FolderManager context={context} />;
}
