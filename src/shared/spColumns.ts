import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/**
 * Create a column only if it is absent, under an internal name we control.
 *
 * Extracted from StructureManager so slice B (adding a segment) does not carry a second copy.
 * Both screens create level columns in BOTH libraries, and the two things that make this
 * function subtle — the internal name and the view behaviour — are exactly what a divergent
 * copy would get wrong months later, silently.
 *
 * `CreateFieldAsXml`, not the `/fields` collection: it is the only form that sets the internal
 * name (Name/StaticName) independently of the display name, which is what keeps the internal
 * name free of `_x0020_` encoding. An internal name is frozen at creation, permanently, so
 * there is no second chance at this.
 *
 * `Options: 8` = AddToAllContentTypes. Deliberately NOT 12, which also adds the column to the
 * default VIEW and would silently reshape every library view the client arranged — once per
 * level anyone ever adds.
 *
 * Returns true if it created the column, false if it was already there. Existence is matched on
 * internal name alone, whatever the field's TYPE: a column already present is never converted,
 * because a type change on a populated column loses data. A caller needing a specific type must
 * check for itself.
 */
export async function ensureColumn(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listTitle: string,
  internalName: string,
  displayName: string,
  kind: "Text" | "Note" | "DateTime" = "Text",
): Promise<boolean> {
  const listBase = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;

  const check: SPHttpClientResponse = await spHttpClient.get(
    `${listBase}/fields?$select=InternalName&$top=500`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata" } },
  );
  if (!check.ok) {
    throw new Error(`${listTitle}: could not read its columns (HTTP ${check.status})`);
  }
  const data = await check.json();
  const exists =
    ((data.value ?? []) as Array<{ InternalName?: string }>).filter(
      (f) => (f.InternalName ?? "").toLowerCase() === internalName.toLowerCase(),
    ).length > 0;
  if (exists) return false;

  // `Format="DateTime"` keeps the TIME component. The default for a DateTime field is date-only,
  // which would collapse every event in a day to the same instant and destroy the ordering an
  // audit feed exists to show — while still looking like a working date column.
  const xml =
    kind === "Note"
      ? `<Field Type="Note" DisplayName="${internalName}" Name="${internalName}" ` +
        `StaticName="${internalName}" NumLines="6" RichText="FALSE" />`
      : kind === "DateTime"
        ? `<Field Type="DateTime" DisplayName="${internalName}" Name="${internalName}" ` +
          `StaticName="${internalName}" Format="DateTime" />`
        : `<Field Type="Text" DisplayName="${internalName}" Name="${internalName}" ` +
          `StaticName="${internalName}" MaxLength="255" />`;

  const create: SPHttpClientResponse = await spHttpClient.post(
    `${listBase}/fields/CreateFieldAsXml`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ parameters: { SchemaXml: xml, Options: 8 } }),
    },
  );
  if (!create.ok) {
    const body = await create.text().catch(() => "");
    throw new Error(
      `${listTitle}: could not create the "${internalName}" column (HTTP ${create.status}). ` +
        body.slice(0, 180),
    );
  }

  // Rename to the human title afterwards. The internal name is already frozen by the create
  // above, so this only changes what is displayed — and a failure here is cosmetic, so it must
  // never fail the operation that just created a correct column.
  if (displayName !== internalName) {
    await spHttpClient
      .post(
        `${listBase}/fields/getbyinternalnameortitle('${encodeURIComponent(internalName)}')`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({ Title: displayName }),
        },
      )
      .catch(() => undefined);
  }
  return true;
}
