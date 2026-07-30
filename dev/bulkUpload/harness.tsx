/* Standalone browser harness for the BulkUpload web part.
 * Renders the REAL component with a mock SharePoint context (no site needed).
 * Build: see dev/bulkUpload/README.md */
import * as React from "react";
import * as ReactDOM from "react-dom";
import BulkUpload from "../../src/webparts/bulkUpload/components/BulkUpload";
import { makeMockContext } from "./mockContext";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const ctx = makeMockContext() as any;

const root = document.getElementById("root");
ReactDOM.render(React.createElement(BulkUpload, { context: ctx }), root);
