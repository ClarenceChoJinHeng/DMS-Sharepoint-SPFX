/**
 * REFERENCE BACKUP — Form.tsx as of when you wiped it to rebuild from scratch.
 * This file is OUTSIDE src/ so TypeScript won't compile it.
 *
 * Key gotchas you hit (don't fall into them again):
 *
 * 1. LIBRARY NAMING
 *    - URL path = "Shared Documents" (used in getfolderbyserverrelativeurl)
 *    - Display title = "Documents" (used in lists/getbytitle)
 *    - These DIFFER for the default Documents library.
 *
 * 2. FOLDER vs DOCUMENT CONTENT TYPES
 *    - Columns added to a library may NOT apply to folders by default.
 *    - To tag folders: enable content type management → add column to
 *      Folder content type.
 *
 * 3. MANAGED METADATA FIELD VALUE FORMAT (for validateUpdateListItem)
 *    - Single: "Label|GUID"
 *    - Multi:  "Label1|GUID1;Label2|GUID2"   (semicolon-separated)
 *    - NOT the "-1;#Label|GUID;#-1;#Label2|GUID2" format — that's legacy.
 *
 * 4. validateUpdateListItem RETURNS 200 EVEN ON FIELD ERRORS
 *    - Always parse the response body and check each result for HasException.
 *    - HTTP 200 + HasException = silent failure.
 *
 * 5. MODERN FOLDER CREATION
 *    - Use POST /_api/web/folders/addUsingPath(decodedurl='/full/path/to/newfolder')
 *    - The full server-relative path goes in decodedurl. No body needed.
 *    - Legacy POST /_api/web/folders with body often hits permission walls.
 *
 * 6. FILE UPLOAD — RAW BINARY, NOT FormData
 *    - /Files/Add expects the raw file bytes as request body.
 *    - { body: file } where file is a File/Blob.
 *    - FormData would corrupt the upload (envelope saved as file content).
 *
 * 7. CONTROLLED vs UNCONTROLLED FILE INPUT
 *    - <input type="file"> is uncontrolled — state changes don't clear it.
 *    - To reset after submit: change the React key (e.g. new Date.now() id)
 *      so the element unmounts/remounts.
 *
 * 8. ModernTaxonomyPicker CONTEXT CAST
 *    - WebPartContext vs BaseComponentContext have a private _serviceScope
 *      field mismatch (SPFx/PnP version drift).
 *    - Workaround: cast context to `any` in a typed local before passing.
 *
 * 9. CONSIDER USING PnPjs (@pnp/sp) NEXT TIME
 *    - Replaces SPHttpClient verbose REST calls with sp.web.folders.addUsingPath(...) etc.
 *    - Cleaner code, typed, autocomplete. Has its own version-pinning pain
 *      but worth it for new projects.
 *
 * 10. SHAREPOINT CONTENT APPROVAL
 *     - Library setting: Versioning settings → Require content approval = Yes
 *     - Items default to "Pending", approval changes status only.
 *     - Moving to destination folder on approval is a SEPARATE step
 *       (manual Copy To, Power Automate, or Content Organizer routing rules).
 */

import * as React from "react";
import { useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ModernTaxonomyPicker } from "@pnp/spfx-controls-react/lib/ModernTaxonomyPicker";
import { IFormProps } from "./IFormProps";

const DEPARTMENT_TERM_SET_ID = "8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f";

type PickedTerm = {
  id: string;
  labels: Array<{ name: string; isDefault: boolean; languageTag: string }>;
};

type FormEntry = {
  id: number;
  file: File[];
  department: PickedTerm[];
};

export default function Form({ context }: IFormProps): React.ReactElement {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pickerContext = context as any;

  const [status, setStatus] = useState<string>("");

  const [entries, setEntries] = useState<FormEntry[]>([
    {
      id: Date.now(),
      file: [],
      department: [],
    },
  ]);

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();

    const invalid = entries.some(
      (entry) => entry.file.length === 0 || entry.department.length === 0,
    );
    if (invalid) {
      setStatus(
        "Every File Set needs at least one file and at least one Department.",
      );
      return;
    }

    const siteUrl = context.pageContext.web.absoluteUrl;
    const library = "Shared Documents";
    const libraryTitle = "Documents";
    const pendingFolder = "Pending";
    const libraryServerRelativeUrl = `${context.pageContext.web.serverRelativeUrl}/${library}`;

    const timestamp = new Date()
      .toISOString()
      .replace(/[:.]/g, "-")
      .replace("T", "_")
      .slice(0, 19);

    setStatus("Uploading…");

    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      const folderName = `${timestamp}_set${i + 1}`;
      const folderServerRelativeUrl = `${libraryServerRelativeUrl}/${pendingFolder}/${folderName}`;

      const folderRes: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/folders/addUsingPath(decodedurl='${encodeURIComponent(folderServerRelativeUrl)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/json" },
        },
      );
      if (!folderRes.ok) {
        setStatus(`Could not create folder for File Set ${i + 1}.`);
        return;
      }

      for (const file of entry.file) {
        const uploadRes: SPHttpClientResponse = await context.spHttpClient.post(
          `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${encodeURIComponent(folderServerRelativeUrl)}')/Files/Add(url='${encodeURIComponent(file.name)}',overwrite=true)`,
          SPHttpClient.configurations.v1,
          { body: file },
        );
        if (!uploadRes.ok) {
          setStatus(`Upload failed for ${file.name}.`);
          return;
        }
      }

      const folderItemRes = await context.spHttpClient.get(
        `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${encodeURIComponent(folderServerRelativeUrl)}')/ListItemAllFields?$select=Id`,
        SPHttpClient.configurations.v1,
      );
      if (!folderItemRes.ok) {
        setStatus(`Could not find folder list item for File Set ${i + 1}.`);
        return;
      }
      const folderItem = await folderItemRes.json();

      const departmentFieldValue = entry.department
        .map((term) => `${term.labels[0].name}|${term.id}`)
        .join(";");

      const metaRes = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${libraryTitle}')/items(${folderItem.Id})/validateUpdateListItem`,
        SPHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            formValues: [
              {
                FieldName: "Department",
                FieldValue: departmentFieldValue,
              },
            ],
          }),
        },
      );
      if (!metaRes.ok) {
        setStatus(`Could not tag folder for File Set ${i + 1}.`);
        return;
      }
      const metaJson = await metaRes.json();
      const fieldError = (metaJson.value ?? []).find(
        (v: { HasException?: boolean }) => v.HasException,
      );
      if (fieldError) {
        console.error("Field update error:", fieldError);
        setStatus(
          `Could not tag folder for File Set ${i + 1}: ${fieldError.ErrorMessage}`,
        );
        return;
      }
    }

    setStatus("Done!");
    setEntries([{ id: Date.now(), file: [], department: [] }]);
  };

  const addContainerOne = (): void => {
    setEntries((prev) => [
      ...prev,
      { id: Date.now(), file: [], department: [] },
    ]);
  };

  const removeContainerOne = (id: number): void => {
    setEntries((prev) =>
      prev.length > 1 ? prev.filter((e) => e.id !== id) : prev,
    );
  };

  const updateEntry = (id: number, patch: Partial<FormEntry>): void => {
    setEntries((prev) =>
      prev.map((e) => (e.id === id ? { ...e, ...patch } : e)),
    );
  };

  return (
    <section>
      <style>
        {`
        .container-one {
            display:flex;
            justify-content: space-between;
            padding: 10px;
            }

            .h1 {
            margin-top: 0px;
            }

            .container-two {
              display:flex;
              gap: 10px;
              max-height:30px;
            }

            .form-container {
            width: 100%;
            display:flex;
            flex-direction: column;
            gap: 10px;
            }

            .flex-wrapper {
            display:flex;
            justify-content: space-between;
            }

          `}
      </style>
      <div className="container-two">
        <button onClick={addContainerOne}>Add +</button>
      </div>

      <div className="container-one">
        <form className="form-container" onSubmit={handleSubmit}>
          {entries.map((entry, index) => (
            <div className="flex-wrapper" key={entry.id}>
              <div>
                <div>
                  <h1 className="h1">File Set {index + 1} :</h1>
                </div>
                <div className="fileUpload">
                  <label htmlFor="file">File Upload:</label>
                  <input
                    type="file"
                    name="file"
                    multiple
                    onChange={(e) => {
                      updateEntry(entry.id, {
                        file: Array.from(e.target.files ?? []),
                      });
                    }}
                  />
                </div>
                <div>
                  <ModernTaxonomyPicker
                    allowMultipleSelections={true}
                    termSetId={DEPARTMENT_TERM_SET_ID}
                    panelTitle="Select Department"
                    label="Send to:"
                    context={pickerContext}
                    onChange={(terms) =>
                      updateEntry(entry.id, { department: terms ?? [] })
                    }
                  />
                </div>
              </div>

              {index !== 0 && (
                <div className="container-two">
                  <button onClick={() => removeContainerOne(entry.id)}>
                    Remove -
                  </button>
                </div>
              )}
            </div>
          ))}
          <button type="submit">Upload</button>
          {status && <p>{status}</p>}
        </form>
      </div>
    </section>
  );
}
