import {
  FIXED_FIELDS,
  buildDetailRows,
  discoverTierFields,
  encodedKey,
  fieldKeys,
  formatBytes,
  labelFromInternalName,
  readField,
  tidBase,
  tierRows,
} from "./documentDetails";

/**
 * A file on Upstream Operations Malaysia — the segment that exposed the bug. Its tiers are Region
 * and Estate/Mill; it has no Department and no Unit, so a hardcoded panel showed neither and dropped
 * both as blank.
 */
const UPOPSMY: Record<string, string> = {
  Business_x005f_x0020_x005f_Segment: "Upstream Operations Malaysia",
  BusinessSegmentTid: "11111111-1111-1111-1111-111111111111",
  Region: "Johor",
  RegionTid: "22222222-2222-2222-2222-222222222222",
  EstateMill: "Bukit Example",
  EstateMillTid: "33333333-3333-3333-3333-333333333333",
  Year: "2024",
  Document_x005f_x0020_x005f_Type: "Working File",
  DocumentDate: "8/13/2026 12:00 AM",
  Confidentiality_x005f_x0020_x005f_Level: "Confidential",
  LegallyPrivileged: "Yes",
  Remark: "test",
};

/** A Group Head Office file — Department and Unit, no Region. The other half of the same problem. */
const GHO: Record<string, string> = {
  Business_x005f_x0020_x005f_Segment: "Group Head Office",
  BusinessSegmentTid: "aaaaaaaa-1111-1111-1111-111111111111",
  Department: "Group Finance",
  DepartmentTid: "bbbbbbbb-2222-2222-2222-222222222222",
  Unit: "Corporate Reporting",
  UnitTid: "cccccccc-3333-3333-3333-333333333333",
  Year: "2026",
};

describe("encodedKey / fieldKeys — the double-encoded name, derived not typed", () => {
  it("double-encodes every underscore", () => {
    expect(encodedKey("Business_x0020_Segment")).toBe("Business_x005f_x0020_x005f_Segment");
  });

  it("handles a LEADING underscore, which is the built-in Description", () => {
    expect(encodedKey("_ExtendedDescription")).toBe("_x005f_ExtendedDescription");
  });

  it("leaves a name with no underscore alone, and offers only one key for it", () => {
    expect(encodedKey("DocumentDate")).toBe("DocumentDate");
    expect(fieldKeys("DocumentDate")).toEqual(["DocumentDate"]);
  });

  it("puts the ENCODED spelling first — that is what a live response uses", () => {
    expect(fieldKeys("Vendor_x002f_CustomerName")[0]).toBe("Vendor_x005f_x002f_x005f_CustomerName");
  });

  it("returns nothing for a blank name rather than a key that matches everything", () => {
    expect(fieldKeys("")).toEqual([]);
    expect(fieldKeys("   ")).toEqual([]);
  });
});

describe("readField", () => {
  it("finds a value under the double-encoded key", () => {
    expect(readField(UPOPSMY, "Business_x0020_Segment")).toBe("Upstream Operations Malaysia");
  });

  it("finds a value under the plain key", () => {
    expect(readField(UPOPSMY, "Region")).toBe("Johor");
  });

  it("is blank for an absent field, and survives a missing response", () => {
    expect(readField(UPOPSMY, "Department")).toBe("");
    expect(readField(undefined as unknown as Record<string, string>, "Region")).toBe("");
  });

  it("treats a whitespace-only value as blank, so it is dropped rather than shown empty", () => {
    expect(readField({ Remark: "   " }, "Remark")).toBe("");
  });
});

describe("tidBase", () => {
  it("strips encoded characters to give the name the Tid column is derived from", () => {
    expect(tidBase("Business_x0020_Segment")).toBe("BusinessSegment");
  });

  it("leaves an already-sanitized name alone", () => {
    // sanitizeFolderSegment REMOVES illegal characters, so "Estate/Mill" became "EstateMill".
    expect(tidBase("EstateMill")).toBe("EstateMill");
    expect(tidBase("Department")).toBe("Department");
  });
});

describe("labelFromInternalName", () => {
  it("recovers the real display name when the space was encoded", () => {
    expect(labelFromInternalName("Business_x0020_Segment")).toBe("Business Segment");
  });

  it("splits camel case for a name sanitized at creation", () => {
    // "Estate/Mill" lost its slash permanently. "Estate Mill" is readable and does not invent
    // punctuation nobody typed.
    expect(labelFromInternalName("EstateMill")).toBe("Estate Mill");
    expect(labelFromInternalName("SubUnit")).toBe("Sub Unit");
  });

  it("leaves a single-word name untouched", () => {
    expect(labelFromInternalName("Region")).toBe("Region");
    expect(labelFromInternalName("Unit")).toBe("Unit");
  });

  it("does not re-split a name the decoding already spaced", () => {
    expect(labelFromInternalName("Project_x0020_Name")).toBe("Project Name");
  });

  it("keeps an acronym together rather than exploding it", () => {
    expect(labelFromInternalName("ITOperatingUnit")).toBe("IT Operating Unit");
  });

  it("is blank for a blank name", () => {
    expect(labelFromInternalName("")).toBe("");
  });
});

describe("discoverTierFields — the fix for the vanished tiers", () => {
  it("finds Region and Estate/Mill on a segment nobody wrote code for", () => {
    // The whole point: no list anywhere names these.
    expect(discoverTierFields(UPOPSMY)).toEqual(["Business_x0020_Segment", "Region", "EstateMill"]);
  });

  it("finds Department and Unit on Group Head Office", () => {
    expect(discoverTierFields(GHO)).toEqual(["Business_x0020_Segment", "Department", "Unit"]);
  });

  it("PINS Business Segment first even when it arrives last", () => {
    const reordered: Record<string, string> = {
      Region: "Johor",
      RegionTid: "x",
      Business_x005f_x0020_x005f_Segment: "Upstream Operations Malaysia",
      BusinessSegmentTid: "y",
    };
    expect(discoverTierFields(reordered)[0]).toBe("Business_x0020_Segment");
  });

  it("EXCLUDES the Tid columns themselves — they hold GUIDs", () => {
    const found = discoverTierFields(UPOPSMY);
    expect(found.filter((f) => /Tid$/.test(f))).toEqual([]);
    expect(JSON.stringify(tierRows(UPOPSMY))).not.toContain("2222-2222");
  });

  it("ignores a field with no Tid twin — that is what makes it not a tier", () => {
    // Year and Document Type are managed metadata and deliberately have no Tid column.
    const found = discoverTierFields(UPOPSMY);
    expect(found).not.toContain("Year");
    expect(found).not.toContain("Document_x0020_Type");
    expect(found).not.toContain("Remark");
  });

  it("does not list one field twice when both spellings are present", () => {
    const both: Record<string, string> = {
      Business_x005f_x0020_x005f_Segment: "A",
      Business_x0020_Segment: "A",
      BusinessSegmentTid: "x",
    };
    expect(discoverTierFields(both)).toEqual(["Business_x0020_Segment"]);
  });

  it("returns nothing rather than throwing on an empty or missing response", () => {
    expect(discoverTierFields({})).toEqual([]);
    expect(discoverTierFields(undefined as unknown as Record<string, string>)).toEqual([]);
  });
});

describe("tierRows", () => {
  it("labels and orders the tiers as the hierarchy reads", () => {
    expect(tierRows(UPOPSMY)).toEqual([
      { label: "Business Segment", value: "Upstream Operations Malaysia" },
      { label: "Region", value: "Johor" },
      { label: "Estate Mill", value: "Bukit Example" },
    ]);
  });

  it("drops a tier column that exists but has no value for this document", () => {
    // An optional SubUnit does not apply to every unit; an em dash row would be noise.
    const withEmptySub = { ...UPOPSMY, SubUnit: "", SubUnitTid: "" };
    expect(tierRows(withEmptySub).map((r) => r.label)).not.toContain("Sub Unit");
  });
});

describe("buildDetailRows", () => {
  it("puts the tiers ABOVE the fixed fields — where it lives, then what it is", () => {
    const labels = buildDetailRows({ fieldText: UPOPSMY }).map((r) => r.label);
    expect(labels.indexOf("Region")).toBeLessThan(labels.indexOf("Document Type"));
    expect(labels.indexOf("Estate Mill")).toBeLessThan(labels.indexOf("Document Type"));
  });

  it("includes every fixed field that has a value, in the library's own order", () => {
    const labels = buildDetailRows({ fieldText: UPOPSMY }).map((r) => r.label);
    expect(labels).toEqual([
      "Business Segment", "Region", "Estate Mill",
      "Document Type", "Year", "Document Date", "Confidentiality",
      "Legally Privileged", "Remark",
    ]);
  });

  it("drops blank fixed fields rather than padding with em dashes", () => {
    // Most columns are empty for any given document on a library serving twelve segments.
    const labels = buildDetailRows({ fieldText: UPOPSMY }).map((r) => r.label);
    expect(labels).not.toContain("Project Name");
    expect(labels).not.toContain("Vendor / Customer");
  });

  it("keeps a caller row EXACTLY as given, blank or not — 'unknown' is meant", () => {
    const rows = buildDetailRows({
      fieldText: {},
      leading: [{ label: "Location", value: "" }],
      trailing: [{ label: "File size", value: "unknown" }],
    });
    expect(rows).toEqual([
      { label: "Location", value: "" },
      { label: "File size", value: "unknown" },
    ]);
  });

  it("puts leading rows first and trailing rows last", () => {
    const rows = buildDetailRows({
      fieldText: GHO,
      leading: [{ label: "Location", value: "GHO › GF › CORU" }],
      trailing: [{ label: "File size", value: "1.4 MB" }],
    });
    expect(rows[0].label).toBe("Location");
    expect(rows[rows.length - 1].label).toBe("File size");
  });

  it("returns nothing but the caller's rows when the metadata could not be read", () => {
    // The screen distinguishes "could not read" from "nothing recorded"; this returns [] for both
    // and lets it.
    expect(buildDetailRows({ fieldText: {} })).toEqual([]);
  });

  it("names LegallyPrivileged so a 'No' is shown, not blanked", () => {
    // Yes/No comes back as the words, which is the whole point on a privilege flag.
    const rows = buildDetailRows({ fieldText: { LegallyPrivileged: "No" } });
    expect(rows).toEqual([{ label: "Legally Privileged", value: "No" }]);
  });

  it("never re-parses the document date — it arrives formatted by SharePoint", () => {
    // Re-parsing is how the M/D/YYYY trap of gotcha #1 gets reintroduced.
    const rows = buildDetailRows({ fieldText: { DocumentDate: "8/13/2026 12:00 AM" } });
    expect(rows[0].value).toBe("8/13/2026 12:00 AM");
  });
});

describe("FIXED_FIELDS", () => {
  it("names no tier column — those are discovered, never listed", () => {
    // The bug: Department and Unit were in this list, so Region and Estate/Mill could never appear.
    const names = FIXED_FIELDS.map((f) => f.field);
    expect(names).not.toContain("Department");
    expect(names).not.toContain("Unit");
    expect(names).not.toContain("Business_x0020_Segment");
  });
});

describe("formatBytes", () => {
  it("formats a byte string, which is what File/Length returns", () => {
    expect(formatBytes("1483776")).toBe("1.4 MB");
    expect(formatBytes(2048)).toBe("2 KB");
  });

  it("shows bytes below 1 KB", () => {
    expect(formatBytes("512")).toBe("512 B");
  });

  it("drops the decimal above 10, where it is noise", () => {
    expect(formatBytes(867532)).toBe("847 KB");
  });

  it("is blank rather than 'NaN' for anything unreadable", () => {
    expect(formatBytes(undefined)).toBe("");
    expect(formatBytes("")).toBe("");
    expect(formatBytes("not a number")).toBe("");
    expect(formatBytes(-5)).toBe("");
  });

  it("handles zero as a real size, not as missing", () => {
    expect(formatBytes(0)).toBe("0 B");
  });
});
