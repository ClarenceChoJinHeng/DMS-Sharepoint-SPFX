/* Dev-harness mock WebPartContext for BulkUpload.
 *
 * Returns a fake `context` whose `spHttpClient` answers every REST call the
 * component (and dmsFolderMap) makes with canned, synthetic data — so the whole
 * batch UX runs in a plain browser with no SharePoint site. The current user is
 * reported as a site admin, so the PRIVILEGED (full manual cascade) path is
 * exercised. All uploads "succeed".
 *
 * NOT bundled into the shipped web part — lives under dev/ only. */

type MockResponse = {
  ok: boolean;
  status: number;
  json: () => Promise<unknown>;
  text: () => Promise<string>;
};

const res = (body: unknown, status = 200): MockResponse => ({
  ok: status >= 200 && status < 300,
  status,
  json: () => Promise.resolve(body),
  text: () =>
    Promise.resolve(typeof body === "string" ? body : JSON.stringify(body)),
});

const delay = (ms: number): Promise<void> =>
  new Promise((r) => setTimeout(r, ms));

/* ---- Synthetic term trees (Department -> Unit) per segment term set ------- */
type Unit = { id: string; label: string };
type Dept = { id: string; label: string; units: Unit[] };

const SEGMENT_TREES: Record<string, Dept[]> = {
  // Group Head Office
  "efa87c6a-9536-4f7c-910f-011bf7413b80": [
    {
      id: "gho-legal",
      label: "Group Legal",
      units: [
        { id: "gho-legal-compliance", label: "Group Compliance" },
        { id: "gho-legal-ops", label: "Group Legal Ops" },
      ],
    },
    {
      id: "gho-risk",
      label: "Risk & Compliance",
      units: [{ id: "gho-risk-grp", label: "Group Risk" }],
    },
  ],
  // Minamas Head Office
  "6ba9a64c-a363-48fd-afd1-324897df781c": [
    {
      id: "min-fin",
      label: "Finance",
      units: [
        { id: "min-fin-ap", label: "Accounts Payable" },
        { id: "min-fin-treasury", label: "Treasury" },
      ],
    },
    {
      id: "min-ops",
      label: "Operations",
      units: [{ id: "min-ops-mill", label: "Milling" }],
    },
  ],
};

/* ---- Flat metadata term sets --------------------------------------------- */
const FLAT_SETS: Record<string, Array<{ id: string; label: string }>> = {
  // documentType
  "0540e66e-7cb3-47ac-b0ef-4e3069387394": [
    { id: "dt-contracts", label: "Contracts" },
    { id: "dt-invoices", label: "Invoices" },
    { id: "dt-reports", label: "Reports" },
  ],
  // yearPeriod
  "f7c578a1-e0e5-42ff-9e0c-d748cba42ede": [
    { id: "yr-2024", label: "2024" },
    { id: "yr-2025", label: "2025" },
    { id: "yr-2026", label: "2026" },
  ],
  // confidentiality
  "032534ab-9285-4b42-98c6-5c7b0df1f066": [
    { id: "cf-public", label: "Public" },
    { id: "cf-internal", label: "Internal" },
    { id: "cf-confidential", label: "Confidential" },
  ],
  // vendor
  "cb3c0ab7-a959-4200-9b7b-d1e13397d240": [
    { id: "vd-acme", label: "Acme Sdn Bhd" },
    { id: "vd-globex", label: "Globex" },
  ],
};

/* ---- DMS Config "mode" rows ---------------------------------------------- */
const LEVELS_JSON = JSON.stringify([
  { label: "Department", column: "Department" },
  { label: "Unit", column: "Unit" },
]);

const MODE_ROWS = [
  {
    Title: "gho",
    ModeLabel: "Group Head Office",
    Side: "BusinessSegment",
    TermSetGuid: "efa87c6a-9536-4f7c-910f-011bf7413b80",
    StagingFolder: "Group Head Office",
    Levels: LEVELS_JSON,
    SortOrder: 1,
  },
  {
    Title: "minamas_ho",
    ModeLabel: "Minamas Head Office",
    Side: "BusinessSegment",
    TermSetGuid: "6ba9a64c-a363-48fd-afd1-324897df781c",
    StagingFolder: "Minamas Head Office",
    Levels: LEVELS_JSON,
    SortOrder: 3,
  },
];

const asTerms = (
  items: Array<{ id: string; label: string }>,
): { value: unknown[] } => ({
  value: items.map((t) => ({ id: t.id, labels: [{ name: t.label }] })),
});

let itemIdSeq = 1000;

/* ---- Mock spHttpClient --------------------------------------------------- */
const mockGet = async (url: string): Promise<MockResponse> => {
  // Term store: children of a term (units under a department)
  let m = url.match(/termStore\/sets\/([^/]+)\/terms\/([^/?]+)\/children/);
  if (m) {
    const termId = m[2];
    for (const depts of Object.values(SEGMENT_TREES)) {
      const dept = depts.find((d) => d.id === termId);
      if (dept) return res(asTerms(dept.units));
    }
    return res({ value: [] });
  }
  // Term store: single term (loadTermPath — restricted only; unused for admin)
  m = url.match(/termStore\/sets\/([^/]+)\/terms\/([^/?]+)/);
  if (m) {
    const termId = m[2];
    return res({ id: termId, labels: [{ name: termId }], parent: null });
  }
  // Term store: children of a set (top level: departments, or a flat set)
  m = url.match(/termStore\/sets\/([^/]+)\/children/);
  if (m) {
    const setId = m[1];
    if (SEGMENT_TREES[setId]) {
      return res(
        asTerms(
          SEGMENT_TREES[setId].map((d) => ({ id: d.id, label: d.label })),
        ),
      );
    }
    if (FLAT_SETS[setId]) return res(asTerms(FLAT_SETS[setId]));
    return res({ value: [] });
  }

  // DMS Config: modes vs settings
  if (url.indexOf("DMS%20Config") >= 0 || url.indexOf("DMS Config") >= 0) {
    if (url.indexOf("eq 'mode'") >= 0) return res({ value: MODE_ROWS });
    return res({ value: [] }); // settings -> component uses its DEFAULT_SETTINGS
  }

  // DMS Group Map (admin, so unused for routing)
  if (url.indexOf("DMS%20Group%20Map") >= 0) return res({ value: [] });

  // Identity
  if (url.indexOf("currentuser/groups") >= 0) return res({ value: [] });
  if (url.indexOf("IsSiteAdmin") >= 0) return res({ IsSiteAdmin: true });

  // DMS Folder Map lookup by term guid
  if (url.indexOf("DMS%20Folder%20Map") >= 0) {
    const tg = url.match(/TermGuid eq '([^']+)'/);
    const guid = tg ? tg[1] : "unknown";
    return res({
      value: [
        {
          Title: guid,
          TermGuid: guid,
          FolderUniqueId: `unit-${guid}`,
          FolderUrl: `/sites/dev/Staging/${guid}`,
          Section: "mock",
        },
      ],
    });
  }

  // File existence probe -> report "not found" so the upload proceeds
  if (url.indexOf("/Files(") >= 0) return res({ error: "not found" }, 404);

  // Folder by UniqueId -> current Staging server-relative path
  m = url.match(/GetFolderById\(guid'([^']+)'\)/);
  if (m && url.indexOf("ServerRelativeUrl") >= 0) {
    return res({ ServerRelativeUrl: `/sites/dev/Staging/${m[1]}` });
  }

  // Folder by path -> always "exists" (so ensureFolder short-circuits)
  m = url.match(/GetFolderByServerRelativeUrl\(@f\)\?@f='([^']*)'/);
  if (m) {
    const path = decodeURIComponent(m[1]);
    return res({ UniqueId: `fld-${path.length}`, ServerRelativeUrl: path });
  }

  // Uploaded file's list item
  if (url.indexOf("/ListItemAllFields") >= 0) {
    return res({ Id: itemIdSeq++ });
  }

  return res({ value: [] });
};

const mockPost = async (
  url: string,
  _config: unknown,
  opts?: { body?: unknown },
): Promise<MockResponse> => {
  // File upload
  if (url.indexOf("/Files/Add(") >= 0) {
    await delay(120); // make the per-file progress visible
    const nm = url.match(/url='([^']*)'/);
    const name = nm ? decodeURIComponent(nm[1]) : "file";
    return res({ ServerRelativeUrl: `/sites/dev/Shared Documents/${name}` });
  }
  // Create folder (rarely hit — folders "exist")
  if (url.indexOf("AddUsingPath") >= 0) {
    const m = url.match(/@u='([^']*)'/);
    const path = m ? decodeURIComponent(m[1]) : "/sites/dev/x";
    return res({ UniqueId: `fld-${path.length}`, ServerRelativeUrl: path });
  }
  // Metadata tagging
  if (url.indexOf("validateUpdateListItem") >= 0) {
    let fields: Array<{ FieldName: string }> = [];
    try {
      const body = JSON.parse(String((opts && opts.body) ?? "{}"));
      fields = body.formValues ?? [];
    } catch {
      /* ignore */
    }
    return res({
      value: fields.map((f) => ({
        FieldName: f.FieldName,
        HasException: false,
        ErrorMessage: null,
      })),
    });
  }
  return res({});
};

const mockSpHttpClient = {
  get: (url: string) => mockGet(url),
  post: (url: string, config: unknown, opts?: { body?: unknown }) =>
    mockPost(url, config, opts),
};

export function makeMockContext(): unknown {
  return {
    spHttpClient: mockSpHttpClient,
    pageContext: {
      web: {
        absoluteUrl: "https://localhost/sites/dev",
        serverRelativeUrl: "/sites/dev",
      },
      user: { displayName: "Dev Admin" },
    },
    sdks: {},
  };
}
