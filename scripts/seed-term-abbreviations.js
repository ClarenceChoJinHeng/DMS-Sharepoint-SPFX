/**
 * Seeds the `DMS Term Abbreviation` list from
 * docs/term-store-import/10-per-level-abbreviations.csv.
 *
 * ── Why a console script rather than a web part ──────────────────────────────
 * Phase 1 of the abbreviation work needs the list populated once, with 175 rows.
 * The admin UI that maintains it afterwards is Phase 2. This unblocks testing
 * without waiting for that UI, the same way the term store and DMS Config have
 * been inspected throughout this project. JavaScript rather than PowerShell
 * because PS scripts do not run reliably in this environment.
 *
 * ── How to run ───────────────────────────────────────────────────────────────
 * 1. Create the list first: Custom List named `DMS Term Abbreviation` with Text
 *    columns TermGuid and Abbreviation, and a Choice column Level
 *    (Segment | Department | Unit). Title holds the full label.
 * 2. Open any page of the site in a browser, signed in.
 * 3. Paste the CSV contents into CSV_TEXT below.
 * 4. Paste this whole file into the DevTools console and press Enter.
 *
 * Re-runnable: an existing row for the same TermGuid is updated, not duplicated.
 *
 * Segments are skipped deliberately — a segment container has no term, so its
 * folder name comes from `StagingFolder` on the DMS Config mode row instead.
 * Set those by hand: mode_gho -> GHO, mode_minamas_ho -> MHO, mode_nbpol_ho -> NBPOLHO.
 */
const SITE = "/sites/ClarenceDMSTesting";
const LIST = "DMS Term Abbreviation";

// Per-site term set GUIDs. These are ClarenceDMSTesting's — a different site has
// its own, because the term store is site-collection local (see CLAUDE.md).
const SETS = {
  "Group Head Office": "08dd94cb-f76c-431c-9b37-e9c98f739ffc",
  "Minamas Head Office": "9ad00b00-a43c-4a8b-a39a-d0efa89ba706",
  "NBPOL Head Office": "77c3993b-0c3c-4a18-89d9-d69209886322",
};

const CSV_TEXT = `"Level","BusinessSegment","Department","Unit","FullName","Abbreviation","Source","Notes"
"Segment","Group Head Office","","","Group Head Office","GHO","from client tokens",""
"Department","Group Head Office","Group Corporate Affairs","","Group Corporate Affairs","GCA","corrected: client typed CGA on 12 of 13 rows; GCA on the other",""
"Unit","Group Head Office","Group Corporate Affairs","Elaeis Garden (Hospitality)","Elaeis Garden (Hospitality)","EG","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Global Marketing & Branding - Product Marketing B2C","Global Marketing & Branding - Product Marketing B2C","GMB_PMB2C","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Global Marketing & Branding - Strategic & Customer Marketing","Global Marketing & Branding - Strategic & Customer Marketing","GMB_SCM","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Global Marketing & Branding - Strategic Communications","Global Marketing & Branding - Strategic Communications","GMB_STRATCOMMS","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Communications - Brand Communications","Group Communications - Brand Communications","GC_BRANDCOMMS","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Communications - Events & Protocol","Group Communications - Events & Protocol","GC_EP","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Communications - Strategic Communications","Group Communications - Strategic Communications","GC_STRATCOMMS","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Operation Services - Agri Sumber Prestari","Group Operation Services - Agri Sumber Prestari","GOS_ASP","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Operation Services - Chairman Office","Group Operation Services - Chairman Office","GOS_CO","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Operation Services - Government Relations","Group Operation Services - Government Relations","GOS_GOVTREL","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Group Operation Services - Workforce Management Unit","Group Operation Services - Workforce Management Unit","GOS_WMU","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Investor Relations","Investor Relations","IR","from client tokens",""
"Unit","Group Head Office","Group Corporate Affairs","Yayasan Guthrie","Yayasan Guthrie","YG","from client tokens",""
"Department","Group Head Office","Group Corporate Secretarial","","Group Corporate Secretarial","COSEC","from client tokens",""
"Unit","Group Head Office","Group Corporate Secretarial","Guthrie","Guthrie","GUTHRIE","from client tokens",""
"Unit","Group Head Office","Group Corporate Secretarial","SDGI","SDGI","SDGI","from client tokens",""
"Department","Group Head Office","Group Finance","","Group Finance","GF","from client tokens",""
"Unit","Group Head Office","Group Finance","Compliance & Operational Risk (CORU)","Compliance & Operational Risk (CORU)","CORU","from client tokens",""
"Unit","Group Head Office","Group Finance","Finance Land & SDGRE","Finance Land & SDGRE","LAND_SDGRE","from client tokens",""
"Unit","Group Head Office","Group Finance","Finance Operations","Finance Operations","FINOPS","from client tokens",""
"Unit","Group Head Office","Group Finance","Finance Upstream Malaysia","Finance Upstream Malaysia","FINUPMY","from client tokens",""
"Unit","Group Head Office","Group Finance","Global Strategic Procurement","Global Strategic Procurement","GSTRATPROC","from client tokens",""
"Unit","Group Head Office","Group Finance","Group IT & Projects","Group IT & Projects","GIT","from client tokens",""
"Unit","Group Head Office","Group Finance","Strategy & Price Risk Management","Strategy & Price Risk Management","STRAT_PRM","from client tokens",""
"Unit","Group Head Office","Group Finance","Tax","Tax","TAX","from client tokens",""
"Unit","Group Head Office","Group Finance","Treasury","Treasury","TREASURY","from client tokens",""
"Department","Group Head Office","Group Human Resources","","Group Human Resources","GHR","from client tokens",""
"Unit","Group Head Office","Group Human Resources","HR Head Office & Org Design - HR Head Office (HR HO)","HR Head Office & Org Design - HR Head Office (HR HO)","HRHO","from client tokens",""
"Unit","Group Head Office","Group Human Resources","HR Head Office & Org Design - Org Design","HR Head Office & Org Design - Org Design","HROD","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Industrial Relations & Compliance","Industrial Relations & Compliance","IRC","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Org Culture & Workplace Services - Admin & Sports","Org Culture & Workplace Services - Admin & Sports","OCWS_ANS","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Org Culture & Workplace Services - Integrated Facilities Management (IFM)","Org Culture & Workplace Services - Integrated Facilities Management (IFM)","OCWS_IFM","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Org Culture & Workplace Services - Org Culture & Transformation (OCT)","Org Culture & Workplace Services - Org Culture & Transformation (OCT)","OCWS_OCT","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Performance & Analytics - Analytics","Performance & Analytics - Analytics","ANALYTICS","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Performance & Analytics - Performance","Performance & Analytics - Performance","PERFORMANCE","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Rewards & HRIS - GHR Payroll","Rewards & HRIS - GHR Payroll","PAYROLL","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Rewards & HRIS - Nadi HRIS","Rewards & HRIS - Nadi HRIS","NHRIS","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Rewards & HRIS - Rewards","Rewards & HRIS - Rewards","REWARDS","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Talent & Learning - GH Campus","Talent & Learning - GH Campus","GHC","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Talent & Learning - Learning & Development (L&D)","Talent & Learning - Learning & Development (L&D)","LND","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Talent & Learning - Talent Acquisition (TA)","Talent & Learning - Talent Acquisition (TA)","TA","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Talent & Learning - Talent Management (TM)","Talent & Learning - Talent Management (TM)","TM","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Talent & Learning - Talent Strategy (TS)","Talent & Learning - Talent Strategy (TS)","TALENTSTRAT","from client tokens",""
"Unit","Group Head Office","Group Human Resources","Talent & Learning - The Crest","Talent & Learning - The Crest","THECREST","from client tokens",""
"Department","Group Head Office","Group Integrity, Governance & Assurance","","Group Integrity, Governance & Assurance","GIGA","from client tokens",""
"Unit","Group Head Office","Group Integrity, Governance & Assurance","GCA IT & Analytics","GCA IT & Analytics","GCAIT_ANALYTICS","from client tokens",""
"Unit","Group Head Office","Group Integrity, Governance & Assurance","GCA SDGI","GCA SDGI","GCA_SDGI","from client tokens",""
"Unit","Group Head Office","Group Integrity, Governance & Assurance","GCA UPMY, R&D & HO","GCA UPMY, R&D & HO","GCA_UPMY_RDHO","from client tokens",""
"Unit","Group Head Office","Group Integrity, Governance & Assurance","Group Fraud & Corruption Risk Mgmt (GFCRM)","Group Fraud & Corruption Risk Mgmt (GFCRM)","GRCRM","from client tokens",""
"Unit","Group Head Office","Group Integrity, Governance & Assurance","Group Integrity & Governance Unit (GIG)","Group Integrity & Governance Unit (GIG)","GIG","from client tokens",""
"Unit","Group Head Office","Group Integrity, Governance & Assurance","Practice Management (PM)","Practice Management (PM)","PM","from client tokens",""
"Department","Group Head Office","Group Legal, Risk & Compliance","","Group Legal, Risk & Compliance","GLRC","from client tokens",""
"Unit","Group Head Office","Group Legal, Risk & Compliance","Group Compliance (GCO)","Group Compliance (GCO)","GCO","from client tokens",""
"Unit","Group Head Office","Group Legal, Risk & Compliance","Group Legal","Group Legal","LEGAL","from client tokens",""
"Unit","Group Head Office","Group Legal, Risk & Compliance","Group Risk","Group Risk","RISK","from client tokens",""
"Department","Group Head Office","Group Sustainability","","Group Sustainability","GS","from client tokens",""
"Unit","Group Head Office","Group Sustainability","Climate Change & Reporting","Climate Change & Reporting","CLIMATE_CR","from client tokens",""
"Unit","Group Head Office","Group Sustainability","Conservation & Biodiversity","Conservation & Biodiversity","CNB","from client tokens",""
"Unit","Group Head Office","Group Sustainability","Global HSE","Global HSE","GLOBALHSE","from client tokens",""
"Unit","Group Head Office","Group Sustainability","Sustainability Compliance","Sustainability Compliance","SC","from client tokens",""
"Unit","Group Head Office","Group Sustainability","Sustainability Head Office","Sustainability Head Office","SUS_HO","from client tokens",""
"Unit","Group Head Office","Group Sustainability","Traceability","Traceability","TRACEABILITY","from client tokens",""
"Department","Group Head Office","President's Office","","President's Office","PO","from client tokens",""
"Unit","Group Head Office","President's Office","Strategic Governance & Communications","Strategic Governance & Communications","STRATGOV_COMMS","from client tokens",""
"Unit","Group Head Office","President's Office","Strategic Planning","Strategic Planning","SP","from client tokens",""
"Unit","Group Head Office","President's Office","Strategy Partner Downstream","Strategy Partner Downstream","SPDNSTRM","from client tokens",""
"Unit","Group Head Office","President's Office","Strategy Partner Malaysia & Indonesia","Strategy Partner Malaysia & Indonesia","SP_MY_ID","from client tokens",""
"Unit","Group Head Office","President's Office","Strategy Partner Midstream & Investments","Strategy Partner Midstream & Investments","SPMID_INVT","from client tokens",""
"Unit","Group Head Office","President's Office","Strategy Partner PNG & SI","Strategy Partner PNG & SI","SP_PNGSI","from client tokens",""
"Segment","Minamas Head Office","","","Minamas Head Office","MHO","from client tokens",""
"Department","Minamas Head Office","CEO Office Administration","","CEO Office Administration","CEOOA","from client tokens",""
"Unit","Minamas Head Office","CEO Office Administration","Expatriate Formalities","Expatriate Formalities","EXPATF","from client tokens",""
"Unit","Minamas Head Office","CEO Office Administration","General Administration & Building Management","General Administration & Building Management","GABM","from client tokens",""
"Department","Minamas Head Office","Corporate Communication","","Corporate Communication","CORPCOMM","corrected: client used CEOOA, which belongs to CEO Office Administration",""
"Unit","Minamas Head Office","Corporate Communication","Corporate Social Responsibility & Operations Comms","Corporate Social Responsibility & Operations Comms","CSR_OPSCOMMS","from client tokens",""
"Unit","Minamas Head Office","Corporate Communication","Internal & External Comms","Internal & External Comms","INEXCOMMS","corrected","row carried no unit token"
"Department","Minamas Head Office","Corporate Services","","Corporate Services","CS","from client tokens",""
"Unit","Minamas Head Office","Corporate Services","Government Relations","Government Relations","GOVTREL","from client tokens",""
"Unit","Minamas Head Office","Corporate Services","Land Regulation","Land Regulation","LANDREL","from client tokens",""
"Unit","Minamas Head Office","Corporate Services","Licensing & Operation Security","Licensing & Operation Security","LOS","from client tokens",""
"Unit","Minamas Head Office","Corporate Services","Minority Shareholder Relation","Minority Shareholder Relation","MSR","from client tokens",""
"Unit","Minamas Head Office","Corporate Services","Partnership Development","Partnership Development","PARTDEV","from client tokens",""
"Department","Minamas Head Office","General Accounting","","General Accounting","GA","from client tokens",""
"Unit","Minamas Head Office","General Accounting","Asset Management","Asset Management","AM","from client tokens",""
"Unit","Minamas Head Office","General Accounting","General Ledger","General Ledger","GL","from client tokens",""
"Unit","Minamas Head Office","General Accounting","Holding & Consolidation","Holding & Consolidation","HCONSOL","from client tokens",""
"Unit","Minamas Head Office","General Accounting","Tax Compliance","Tax Compliance","TAX","from client tokens",""
"Department","Minamas Head Office","Group Compliance","","Group Compliance","GCO","from client tokens",""
"Unit","Minamas Head Office","Group Compliance","Group Compliance","Group Compliance","GCO","from client tokens","department and unit share a name"
"Department","Minamas Head Office","Group Corporate Assurance","","Group Corporate Assurance","GCA","corrected: one row omitted the department token",""
"Unit","Minamas Head Office","Group Corporate Assurance","Data Intelligence & Investigations","Data Intelligence & Investigations","DIINVESTN","from client tokens",""
"Unit","Minamas Head Office","Group Corporate Assurance","Estate & Other Operations","Estate & Other Operations","EOOPS","from client tokens",""
"Unit","Minamas Head Office","Group Corporate Assurance","Practice Management","Practice Management","PM","from client tokens",""
"Unit","Minamas Head Office","Group Corporate Assurance","Project & Mill Operations","Project & Mill Operations","PMOPS","corrected","row carried no unit token"
"Department","Minamas Head Office","Group Integrity & Governance","","Group Integrity & Governance","GIG","from client tokens",""
"Unit","Minamas Head Office","Group Integrity & Governance","Group Integrity & Governance","Group Integrity & Governance","GIG","from client tokens","department and unit share a name"
"Department","Minamas Head Office","Human Resource Management","","Human Resource Management","HRM","from client tokens",""
"Unit","Minamas Head Office","Human Resource Management","HR Business Partner and Industrial Relations","HR Business Partner and Industrial Relations","HRBP_IR","from client tokens",""
"Unit","Minamas Head Office","Human Resource Management","HR Rewards, Services, and Performance","HR Rewards, Services, and Performance","HRRSP","from client tokens",""
"Unit","Minamas Head Office","Human Resource Management","Organisation Development and Talent Management","Organisation Development and Talent Management","ODTM","from client tokens",""
"Department","Minamas Head Office","Information Technology Services","","Information Technology Services","ITS","from client tokens",""
"Unit","Minamas Head Office","Information Technology Services","Business Application & Development","Business Application & Development","BADEV","from client tokens",""
"Unit","Minamas Head Office","Information Technology Services","Infrastructure & IT Security","Infrastructure & IT Security","INFRA_ITS","from client tokens",""
"Unit","Minamas Head Office","Information Technology Services","SAP Strategic Project & Data","SAP Strategic Project & Data","SAPSTRATPD","from client tokens",""
"Department","Minamas Head Office","Legal & Corporate Secretary","","Legal & Corporate Secretary","LCS","from client tokens",""
"Unit","Minamas Head Office","Legal & Corporate Secretary","Corporate Secretary","Corporate Secretary","COSEC","from client tokens",""
"Unit","Minamas Head Office","Legal & Corporate Secretary","Legal","Legal","LEGAL","from client tokens",""
"Department","Minamas Head Office","Performance Monitoring Unit & Mechanisation","","Performance Monitoring Unit & Mechanisation","PMUM","from client tokens",""
"Unit","Minamas Head Office","Performance Monitoring Unit & Mechanisation","Mechanisation","Mechanisation","MEC","from client tokens",""
"Unit","Minamas Head Office","Performance Monitoring Unit & Mechanisation","Performance Monitoring Unit","Performance Monitoring Unit","PMU","from client tokens",""
"Department","Minamas Head Office","Plantation Operation","","Plantation Operation","POP","from client tokens",""
"Unit","Minamas Head Office","Plantation Operation","Estate Report","Estate Report","ER","from client tokens",""
"Unit","Minamas Head Office","Plantation Operation","Management Report","Management Report","MR","from client tokens",""
"Unit","Minamas Head Office","Plantation Operation","Mill Report","Mill Report","MILLR","from client tokens",""
"Unit","Minamas Head Office","Plantation Operation","Task Force Operation","Task Force Operation","TFO","from client tokens",""
"Department","Minamas Head Office","Production & Engineering","","Production & Engineering","PNE","corrected: client used both PNE and PROD",""
"Unit","Minamas Head Office","Production & Engineering","Engineering","Engineering","ENGINEERING","from client tokens",""
"Unit","Minamas Head Office","Production & Engineering","Production & Operation","Production & Operation","OPS","from client tokens",""
"Department","Minamas Head Office","Risk Management","","Risk Management","GRM","from client tokens",""
"Unit","Minamas Head Office","Risk Management","Risk Management","Risk Management","GRM","from client tokens","department and unit share a name"
"Department","Minamas Head Office","Supply Chain","","Supply Chain","SC","from client tokens",""
"Unit","Minamas Head Office","Supply Chain","Logistics","Logistics","LOGS","from client tokens",""
"Unit","Minamas Head Office","Supply Chain","Procurement","Procurement","PROC","from client tokens",""
"Department","Minamas Head Office","Sustainability","","Sustainability","SUS","corrected: client used SC on 3 rows, which belongs to Supply Chain",""
"Unit","Minamas Head Office","Sustainability","Environment, Social & Governance","Environment, Social & Governance","ESG","from client tokens",""
"Unit","Minamas Head Office","Sustainability","Health & Medical Services","Health & Medical Services","HMS","from client tokens",""
"Unit","Minamas Head Office","Sustainability","Health, Safety, and Environment","Health, Safety, and Environment","HSE","from client tokens",""
"Unit","Minamas Head Office","Sustainability","Sustainability Compliance & Quality Management","Sustainability Compliance & Quality Management","QM","from client tokens",""
"Department","Minamas Head Office","Taxation","","Taxation","TAX","from client tokens",""
"Unit","Minamas Head Office","Taxation","Taxation","Taxation","TAX","from client tokens","department and unit share a name"
"Department","Minamas Head Office","Treasury","","Treasury","TREAS","from client tokens",""
"Unit","Minamas Head Office","Treasury","Account Payable","Account Payable","AC","from client tokens",""
"Unit","Minamas Head Office","Treasury","Funding Operations","Funding Operations","FO","from client tokens",""
"Unit","Minamas Head Office","Treasury","Insurance Management","Insurance Management","INSURM","from client tokens",""
"Unit","Minamas Head Office","Treasury","Payment","Payment","PAYMENT","from client tokens",""
"Department","Minamas Head Office","Upstream Productivity","","Upstream Productivity","UPPROD","from client tokens",""
"Unit","Minamas Head Office","Upstream Productivity","Upstream Productivity","Upstream Productivity","UPPROD","from client tokens","department and unit share a name"
"Department","Minamas Head Office","Value Creation / Value Transformation","","Value Creation / Value Transformation","VCVT","from client tokens",""
"Unit","Minamas Head Office","Value Creation / Value Transformation","Value Creation / Value Transformation","Value Creation / Value Transformation","VCVT","from client tokens","department and unit share a name"
"Segment","NBPOL Head Office","","","NBPOL Head Office","NBPOLHO","from client tokens",""
"Department","NBPOL Head Office","Chief Operating Officer's Office","","Chief Operating Officer's Office","COO","from client tokens",""
"Unit","NBPOL Head Office","Chief Operating Officer's Office","Chief Operating Officer's Office","Chief Operating Officer's Office","COO","from client tokens","department and unit share a name"
"Department","NBPOL Head Office","Commercial","","Commercial","COMMERCIAL","from client tokens",""
"Unit","NBPOL Head Office","Commercial","Commercial","Commercial","COMMERCIAL","from client tokens","department and unit share a name"
"Department","NBPOL Head Office","Corporate Development & Services","","Corporate Development & Services","CDS","corrected: row carried only the unit token",""
"Unit","NBPOL Head Office","Corporate Development & Services","Upstream Support","Upstream Support","UPSUPPORT","corrected","row carried no unit token"
"Department","NBPOL Head Office","Engineering & Mill Services","","Engineering & Mill Services","EMS","from client tokens",""
"Unit","NBPOL Head Office","Engineering & Mill Services","Engineering Services","Engineering Services","ES","from client tokens",""
"Department","NBPOL Head Office","Finance","","Finance","FIN","from client tokens",""
"Unit","NBPOL Head Office","Finance","Finance Reporting","Finance Reporting","REPORTING","from client tokens",""
"Unit","NBPOL Head Office","Finance","Tax","Tax","TAX","from client tokens",""
"Department","NBPOL Head Office","Group Corporate Assurance","","Group Corporate Assurance","GCA","from client tokens",""
"Unit","NBPOL Head Office","Group Corporate Assurance","Group Corporate Assurance","Group Corporate Assurance","GCA","from client tokens","department and unit share a name"
"Department","NBPOL Head Office","Human Resources","","Human Resources","HR","from client tokens",""
"Unit","NBPOL Head Office","Human Resources","HR Operations","HR Operations","OPS","from client tokens",""
"Unit","NBPOL Head Office","Human Resources","HR Strategy & Talent","HR Strategy & Talent","STRATTALENT","from client tokens",""
"Unit","NBPOL Head Office","Human Resources","HRMIS","HRMIS","HRMIS","from client tokens",""
"Unit","NBPOL Head Office","Human Resources","Management Academy","Management Academy","MGMTACAD","from client tokens",""
"Unit","NBPOL Head Office","Human Resources","Travel","Travel","TRAVEL","from client tokens",""
"Department","NBPOL Head Office","Information Technology","","Information Technology","IT","from client tokens",""
"Unit","NBPOL Head Office","Information Technology","NBPOL IT Application Development","NBPOL IT Application Development","APPDEV","from client tokens",""
"Unit","NBPOL Head Office","Information Technology","NBPOL IT SAP Support","NBPOL IT SAP Support","SAP","from client tokens",""
"Department","NBPOL Head Office","Land Department","","Land Department","LAND","from client tokens",""
"Unit","NBPOL Head Office","Land Department","Land Management","Land Management","LM","from client tokens",""
"Unit","NBPOL Head Office","Land Department","Surveying","Surveying","SURVEYING","from client tokens",""
"Department","NBPOL Head Office","Legal","","Legal","LEGAL","from client tokens",""
"Unit","NBPOL Head Office","Legal","Legal","Legal","LEGAL","from client tokens","department and unit share a name"
"Department","NBPOL Head Office","NBPOL Procurement","","NBPOL Procurement","PROC","from client tokens",""
"Unit","NBPOL Head Office","NBPOL Procurement","Corporate Office","Corporate Office","CO","from client tokens",""
"Department","NBPOL Head Office","Performance Management Unit","","Performance Management Unit","PMU","corrected: row carried only the unit token",""
"Unit","NBPOL Head Office","Performance Management Unit","Standard Crop Recovery Assessment","Standard Crop Recovery Assessment","SCRA","corrected","row carried no unit token"
"Department","NBPOL Head Office","Sustainability","","Sustainability","SUS","from client tokens",""
"Unit","NBPOL Head Office","Sustainability","Certification","Certification","CERT","from client tokens",""
"Unit","NBPOL Head Office","Sustainability","Environment, Safety & Health","Environment, Safety & Health","ESH","from client tokens",""
"Department","NBPOL Head Office","Transformation","","Transformation","TRANSFORM","from client tokens",""
"Unit","NBPOL Head Office","Transformation","Transformation","Transformation","TRANSFORM","from client tokens","department and unit share a name"
`;

/**
 * Folds the differences that are presentation, not identity, so a CSV label
 * matches its live term.
 *
 * The fullwidth ＆ is the one that actually bit: the term store requires it
 * (see CLAUDE.md), the client's spreadsheet uses plain &, and the first run
 * matched only 102 of 175 rows because of it. A department mismatch takes its
 * units down with it, since a unit path is Segment|Department|Unit.
 *
 * The quote, dash and space rules are pre-emptive: Word and Excel rewrite those
 * silently and the client's data came through both.
 */
function normalizeLabel(s) {
  return (s || "")
    .replace(/＆/g, "&") // fullwidth ampersand
    .replace(/[‘’ʼ]/g, "'") // curly / modifier apostrophes
    .replace(/[“”]/g, '"')
    .replace(/[‐‑‒–—―]/g, "-") // en/em dashes and friends
    .replace(/ /g, " ") // non-breaking space
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

const getJson = async (url) => {
  const r = await fetch(url, {
    headers: { Accept: "application/json;odata=nometadata" },
  });
  if (!r.ok) throw new Error(`HTTP ${r.status} on ${url}: ${await r.text()}`);
  return r.json();
};

/**
 * Minimal CSV reader. A naive split on "," is wrong here: the department
 * "Group Legal, Risk & Compliance" contains a comma and is therefore quoted.
 */
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') inQuotes = false;
      else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (c !== "\r") field += c;
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  const head = rows.shift();
  return rows
    .filter((r) => r.length === head.length)
    .map((r) => {
      const o = {};
      head.forEach((h, i) => (o[h.trim()] = r[i]));
      return o;
    });
}

(async () => {
  // 1. Walk every term set, building "Segment|Department[|Unit]" -> term GUID.
  //    The CSV is keyed by label path because the client supplied names, not
  //    GUIDs. Both sides of that key go through normalizeLabel, so a fullwidth
  //    ＆ in the store still matches a plain & in the spreadsheet.
  const guidByPath = new Map();
  for (const segName of Object.keys(SETS)) {
    const setId = SETS[segName];
    const walk = async (url, prefix) => {
      const d = await getJson(url);
      for (const t of d.value || []) {
        const label = t.labels[0].name;
        const path = `${prefix}|${normalizeLabel(label)}`;
        // The LIVE label, not the CSV's FullName: this becomes Title, which
        // Task 5 writes to the folder's Full Name column. The client must see
        // the same text there as in the upload dropdown, fullwidth ＆ included.
        guidByPath.set(path, { id: t.id, label: label });
        await walk(
          `${SITE}/_api/v2.1/termStore/sets/${setId}/terms/${t.id}/children`,
          path,
        );
      }
    };
    await walk(
      `${SITE}/_api/v2.1/termStore/sets/${setId}/children`,
      normalizeLabel(segName),
    );
  }
  console.log(`resolved ${guidByPath.size} terms from the term store`);

  // 2. Read existing rows so re-running updates rather than duplicates.
  const existing = new Map();
  const cur = await getJson(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(
      LIST,
    )}')/items?$select=Id,TermGuid&$top=5000`,
  );
  (cur.value || []).forEach((r) =>
    existing.set((r.TermGuid || "").toLowerCase(), r.Id),
  );
  console.log(`list already holds ${existing.size} rows`);

  // 3. Write.
  const ctx = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" },
  });
  const digest = (await ctx.json()).FormDigestValue;

  let written = 0;
  const missing = [];
  for (const r of parseCsv(CSV_TEXT)) {
    if (r.Level === "Segment") continue; // named via StagingFolder, not this list
    const parts =
      r.Level === "Department"
        ? [r.BusinessSegment, r.Department]
        : [r.BusinessSegment, r.Department, r.Unit];
    const path = parts.map(normalizeLabel).join("|");
    const term = guidByPath.get(path);
    if (!term) {
      missing.push(path);
      continue;
    }
    const id = existing.get(term.id.toLowerCase());
    const base = `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(LIST)}')/items`;
    const res = await fetch(id ? `${base}(${id})` : base, {
      method: "POST",
      headers: Object.assign(
        {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest,
        },
        id ? { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" } : {},
      ),
      body: JSON.stringify({
        Title: term.label,
        TermGuid: term.id,
        Abbreviation: r.Abbreviation,
        Level: r.Level,
      }),
    });
    if (!res.ok) {
      console.error(`FAILED ${path}: HTTP ${res.status} ${await res.text()}`);
      continue;
    }
    written++;
  }

  console.log(`wrote ${written} rows`);
  if (missing.length) {
    // A label path that does not resolve means the CSV text and the live term
    // label differ. Investigate the term store rather than editing the CSV to
    // match — the CSV was generated from the client's own data.
    console.warn(`could not resolve ${missing.length} label paths:`, missing);
  }
})();
