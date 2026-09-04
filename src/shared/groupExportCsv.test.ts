import {
  csvCell,
  toCsv,
  exportFileName,
  GroupExportRow,
  NO_MEMBERS,
} from "./groupExportCsv";

const row = (over: Partial<GroupExportRow> = {}): GroupExportRow => ({
  group: "DMS_GHO_GF_CORU_APR",
  segment: "Group Head Office",
  tier1: "Group Finance",
  tier2: "Compliance ＆ Operational Risk (CORU)",
  role: "APR",
  memberName: "Jane Doe",
  memberEmail: "jane.doe@example.com",
  ...over,
});

describe("csvCell", () => {
  it("leaves a plain value alone", () => {
    expect(csvCell("Group Head Office")).toBe("Group Head Office");
  });

  // Real unit name: "Group Legal, Risk ＆ Compliance" would otherwise split into
  // two columns and shift every later column left.
  it("quotes a value containing a comma", () => {
    expect(csvCell("Group Legal, Risk ＆ Compliance")).toBe(
      '"Group Legal, Risk ＆ Compliance"',
    );
  });

  it("doubles embedded quotes and wraps", () => {
    expect(csvCell('He said "hi"')).toBe('"He said ""hi"""');
  });

  it("quotes a value containing a newline", () => {
    expect(csvCell("line1\nline2")).toBe('"line1\nline2"');
  });

  it("neutralises a leading = so Excel does not evaluate it", () => {
    expect(csvCell("=1+1")).toBe("'=1+1");
  });

  it("neutralises a leading +", () => {
    expect(csvCell("+cmd")).toBe("'+cmd");
  });

  it("neutralises a leading -", () => {
    expect(csvCell("-cmd")).toBe("'-cmd");
  });

  it("neutralises a leading @", () => {
    expect(csvCell("@cmd")).toBe("'@cmd");
  });

  it("treats undefined as empty", () => {
    expect(csvCell(undefined as unknown as string)).toBe("");
  });
});

describe("toCsv", () => {
  it("emits the header row even with no data", () => {
    expect(toCsv([])).toBe("Group,Segment,Tier 1,Tier 2,Role,Member Name,Member Email");
  });

  it("emits one line per member row, CRLF separated", () => {
    const csv = toCsv([
      row(),
      row({ memberName: "John Roe", memberEmail: "john.roe@example.com" }),
    ]);
    const lines = csv.split("\r\n");
    expect(lines).toHaveLength(3);
    expect(lines[1]).toContain("jane.doe@example.com");
    expect(lines[2]).toContain("john.roe@example.com");
  });

  it("keeps a group with no members as a single placeholder row", () => {
    const csv = toCsv([row({ memberName: NO_MEMBERS, memberEmail: "" })]);
    expect(csv.split("\r\n")[1]).toBe(
      `DMS_GHO_GF_CORU_APR,Group Head Office,Group Finance,Compliance ＆ Operational Risk (CORU),APR,${NO_MEMBERS},`,
    );
  });

  it("keeps column alignment when a tier contains a comma", () => {
    const csv = toCsv([row({ tier1: "Group Legal, Risk ＆ Compliance", tier2: "Legal" })]);
    expect(csv.split("\r\n")[1]).toContain('"Group Legal, Risk ＆ Compliance"');
  });
});

describe("exportFileName", () => {
  it("pads month and day", () => {
    expect(exportFileName(new Date(2026, 6, 5))).toBe("CRS-group-members-2026-07-05.csv");
  });
});
