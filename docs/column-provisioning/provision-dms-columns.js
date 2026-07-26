/*
 * DMS column provisioning — BROWSER CONSOLE snippet (JavaScript, NOT PowerShell).
 *
 * HOW TO RUN:
 *   1. Open any page on  https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting
 *      (you must be signed in — the snippet uses your session).
 *   2. Press F12 → Console tab.
 *   3. Paste this whole file → Enter.
 *   4. Watch the log. Re-runnable: existing columns are skipped (idempotent).
 *
 * WHAT IT DOES:
 *   - Staging   : 6 text + DocumentDate + 4 taxonomy = 11 columns
 *   - Documents : same 11 + Approval Status (Choice, default "Approved") = 12
 *   - Exact internal names via AddFieldInternalNameHint (Options: 8).
 *
 * ⚠ TAXONOMY BINDING DOES NOT WORK VIA REST (verified 2026-07-27):
 *   The bindTax() MERGE fails with "A type named 'SP.TaxonomyField' could not be
 *   resolved by the model." The 4 taxonomy FIELDS get created with correct internal
 *   names, but must be BOUND to their term set via the UI:
 *   Library settings → column → Term Set Settings → Site level term groups → DMS → <set>.
 *   Sets: Document Type→866c5754… · Year→023a866a… · Confidentiality Level→0d6d1da8… · Vendor→eaafd0e5…
 *
 * VERIFY AFTER: open a new item in each library and confirm each taxonomy column
 * offers its terms (Document Type shows the 20, etc.). If a column errors in the
 * log, copy the error line back to Claude.
 */
(async () => {
  const web   = "https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting";
  const SSP_ID = "df05a518-5c5e-411c-9444-cf6e1bd10a88"; // term store "Unique identifier"

  const digest = async () =>
    (await (await fetch(web + "/_api/contextinfo", {
      method: "POST", headers: { Accept: "application/json;odata=verbose" }
    })).json()).d.GetContextWebInformation.FormDigestValue;

  const post = async (url, body, extra = {}) => fetch(web + url, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": await digest(),
      ...extra
    },
    body: JSON.stringify(body)
  });

  const exists = async (list, internal) =>
    (await fetch(web + `/_api/web/lists/getbytitle('${list}')/fields/getbyinternalnameortitle('${internal}')?$select=Id`,
      { headers: { Accept: "application/json;odata=verbose" } })).ok;

  const addXml = async (list, internal, xml) => {
    if (await exists(list, internal)) { console.log(`= ${list}.${internal} exists`); return; }
    const r = await post(`/_api/web/lists/getbytitle('${list}')/fields/createfieldasxml`, {
      parameters: { "__metadata": { type: "SP.XmlSchemaFieldCreationInformation" }, SchemaXml: xml, Options: 8 }
    });
    console.log(r.ok ? `+ ${list}.${internal}` : `✗ ${list}.${internal}: ${await r.text()}`);
  };

  const bindTax = async (list, internal, termSetId) => {
    const id = (await (await fetch(web + `/_api/web/lists/getbytitle('${list}')/fields/getbyinternalnameortitle('${internal}')?$select=Id`,
      { headers: { Accept: "application/json;odata=verbose" } })).json()).d.Id;
    const r = await post(`/_api/web/lists/getbytitle('${list}')/fields(guid'${id}')`, {
      "__metadata": { type: "SP.TaxonomyField" },
      SspId: SSP_ID, TermSetId: termSetId,
      AnchorId: "00000000-0000-0000-0000-000000000000", TargetTemplate: null
    }, { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" });
    console.log(r.ok ? `  ↳ bound ${internal} → ${termSetId}` : `  ✗ bind ${internal}: ${await r.text()}`);
  };

  const text   = (n, d) => `<Field Type="Text" Name="${n}" StaticName="${n}" DisplayName="${d}"/>`;
  const date   = (n, d) => `<Field Type="DateTime" Name="${n}" StaticName="${n}" DisplayName="${d}" Format="DateOnly"/>`;
  const tax    = (n, d) => `<Field Type="TaxonomyFieldType" Name="${n}" StaticName="${n}" DisplayName="${d}"></Field>`;
  const choice = (n, d) => `<Field Type="Choice" Name="${n}" StaticName="${n}" DisplayName="${d}"><CHOICES><CHOICE>Approved</CHOICE></CHOICES><Default>Approved</Default></Field>`;

  const TEXT = [
    ["Business_x0020_Segment", "Business Segment"],
    ["BusinessSegmentTid",     "Business Segment Tid"],
    ["Department",             "Department"],
    ["DepartmentTid",          "Department Tid"],
    ["Unit",                   "Unit"],
    ["UnitTid",                "Unit Tid"],
  ];
  const TAX = [
    ["Document_x0020_Type",         "Document Type",         "866c5754-258e-401f-8685-03d20ae59b1d"],
    ["Year_x002f_Period",           "Year",                  "023a866a-5c0b-4f1b-ad42-2ddf7a9e7abf"],
    ["Confidentiality_x0020_Level", "Confidentiality Level", "0d6d1da8-27e5-477f-8684-e8cf169f8fb9"],
    ["Vendor",                      "Vendor",                "eaafd0e5-03fd-4d33-b1b1-e4252bec430a"],
  ];

  const provision = async (list, withApproval) => {
    console.log(`--- ${list} ---`);
    for (const [n, d] of TEXT) await addXml(list, n, text(n, d));
    await addXml(list, "DocumentDate", date("DocumentDate", "Document Date"));
    for (const [n, d, ts] of TAX) { await addXml(list, n, tax(n, d)); await bindTax(list, n, ts); }
    if (withApproval) await addXml(list, "Approval_x0020_Status", choice("Approval_x0020_Status", "Approval Status"));
  };

  await provision("Staging", false);
  await provision("Documents", true);
  console.log("DONE — open a new item in each library and confirm the taxonomy columns show their terms.");
})();
