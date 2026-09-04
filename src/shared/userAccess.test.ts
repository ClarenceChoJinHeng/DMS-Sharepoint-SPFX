import {
  AccessRow,
  UserGroupRef,
  personDisplay,
  personaForRoles,
  summarizeUserAccess,
  describeGroupAccess,
} from "./userAccess";
import { PERSONAS, siteEntryGroupTitle } from "./groupMapModel";

const resolve = {
  segmentLabel: (guid: string): string =>
    guid === "seg-gho" ? "Group Head Office" : guid === "seg-mho" ? "Minamas Head Office" : "",
  tierChain: (guid: string): string[] =>
    guid === "term-tax" ? ["Group Finance", "Tax"] : guid === "term-gf" ? ["Group Finance"] : [],
};

const row = (over: Partial<AccessRow> & { groupId: number }): AccessRow => ({
  segment: "seg-gho",
  unitTermGuid: "term-tax",
  role: "UPL",
  scope: "Folder",
  target: "Staging",
  ...over,
});

const g = (id: number, title: string): UserGroupRef => ({ id, title });

describe("personaForRoles", () => {
  it("matches a persona on its exact role set, whatever the order or case", () => {
    expect(personaForRoles(["UPL"])?.key).toBe("pic");
    expect(personaForRoles(["upl"])?.key).toBe("pic");
    // hod's set gained DELHC/SHAREHC on 2026-08-24, so it is a five-role match now.
    expect(personaForRoles(["SHARE", "DEL", "DEPTVIEW", "DELHC", "SHAREHC"])?.key).toBe("hod");
    // hou_hc's full set since the 2026-08-24 clarification ("only HC HOU can see and approve") —
    // APRHC in place of APR, plus DELHC and SHAREHC, none of which the plain hou holds any more.
    expect(
      personaForRoles(["APRHC", "DELS", "DEL", "SHARE", "UPLHC", "DELSHC", "DELHC", "SHAREHC"])?.key,
    ).toBe("hou_hc");
    // The plain hou set: no HC role reaches it at all — a role held by two personas cannot grant to
    // one and withhold from the other, so any single HC role here would match nothing.
    expect(personaForRoles(["APR", "DELS", "DEL", "SHARE", "UPL"])?.key).toBe("hou");
  });

  it("ignores duplicates, which a row set legitimately contains", () => {
    expect(personaForRoles(["UPL", "UPL"])?.key).toBe("pic");
  });

  // The whole reason this is an exact match. `employee` is ["MEMBER"] and `employee_hc` is
  // ["MEMBER","MEMBERHC"], so a subset rule would let a CLEARED viewer group answer as the plain
  // one — presenting HC clearance as absent, which is the dangerous direction.
  it("does NOT match a subset or a superset", () => {
    expect(personaForRoles(["MEMBER"])?.key).toBe("employee");
    expect(personaForRoles(["MEMBER", "MEMBERHC"])?.key).toBe("employee_hc");
    expect(personaForRoles(["MEMBER", "MEMBERHC", "UPL"])).toBeUndefined();
    expect(personaForRoles(["APR"])).toBeUndefined();
  });

  it("answers undefined for nothing and for an unknown role", () => {
    expect(personaForRoles([])).toBeUndefined();
    expect(personaForRoles([""])).toBeUndefined();
    expect(personaForRoles(["NOT_A_ROLE"])).toBeUndefined();
  });

  it("can name every persona in the model, so none is unreachable", () => {
    for (const p of PERSONAS) {
      expect(personaForRoles(p.roles as unknown as string[])?.key).toBe(p.key);
    }
  });
});

describe("summarizeUserAccess", () => {
  it("reports the persona, the roles and the place for a mapped group", () => {
    const s = summarizeUserAccess([g(42, "GHO_GF_TAX_UPLOADER")], [row({ groupId: 42 })], resolve);
    expect(s.none).toBe(false);
    expect(s.groups).toHaveLength(1);
    expect(s.groups[0].persona?.key).toBe("pic");
    expect(s.groups[0].roles).toEqual(["UPL"]);
    expect(s.groups[0].places).toEqual([
      {
        segmentLabel: "Group Head Office",
        tierChain: ["Group Finance", "Tax"],
        scope: "Folder",
        target: "Staging",
      },
    ]);
  });

  // An approver group carries six role rows at ONE term. Listing that term six times says nothing
  // the first line did not.
  it("collapses many rows at one place into a single place", () => {
    // Plain hou's role set as of 2026-08-24 — DELSHC left this persona for hou_hc alone, so a
    // fixture mixing it with UPL/APR (as this once did) matches no persona at all.
    const roles = ["APR", "DELS", "DEL", "SHARE", "UPL"];
    const s = summarizeUserAccess(
      [g(7, "GHO_GF_TAX_APPROVER")],
      roles.map((r) => row({ groupId: 7, role: r })),
      resolve,
    );
    expect(s.groups[0].places).toHaveLength(1);
    expect(s.groups[0].roles).toEqual(roles);
    expect(s.groups[0].persona?.key).toBe("hou");
  });

  it("keeps genuinely different places apart", () => {
    const s = summarizeUserAccess(
      [g(7, "X")],
      [
        row({ groupId: 7, unitTermGuid: "term-tax" }),
        row({ groupId: 7, unitTermGuid: "term-gf" }),
        row({ groupId: 7, segment: "seg-mho", unitTermGuid: "term-tax" }),
      ],
      resolve,
    );
    expect(s.groups[0].places).toHaveLength(3);
  });

  // The state that exists because creating a group stopped implying a mapping (2026-08-14), and the
  // most useful thing this page can tell someone asking why a person cannot get in.
  it("flags a group with no rows as granting nothing", () => {
    const s = summarizeUserAccess([g(9, "GHO_GF_TAX_EMPLOYEE")], [], resolve);
    expect(s.groups[0].unmapped).toBe(true);
    expect(s.groups[0].places).toEqual([]);
    expect(s.unmappedCount).toBe(1);
    expect(describeGroupAccess(s.groups[0])).toContain("Grants nothing yet");
  });

  it("flags the site-entry group, which is not a document grant", () => {
    const s = summarizeUserAccess([g(1, siteEntryGroupTitle())], [], resolve);
    expect(s.groups[0].siteEntry).toBe(true);
    expect(describeGroupAccess(s.groups[0])).toContain("Opens the site");
  });

  // Joining on the NAME would break on a rename (the id survives, the stored GroupName goes stale)
  // and on two groups renamed alike.
  it("joins rows on the group id, never the title", () => {
    const s = summarizeUserAccess([g(42, "RENAMED_SINCE")], [row({ groupId: 42 })], resolve);
    expect(s.groups[0].roles).toEqual(["UPL"]);
    const other = summarizeUserAccess([g(43, "RENAMED_SINCE")], [row({ groupId: 42 })], resolve);
    expect(other.groups[0].unmapped).toBe(true);
  });

  // Reporting an unresolved chain as a bare leaf presents a UNIT as a DEPARTMENT — the difference
  // between granting one unit and granting all of them.
  it("leaves an unresolvable chain EMPTY rather than guessing a leaf", () => {
    const s = summarizeUserAccess(
      [g(42, "X")],
      [row({ groupId: 42, unitTermGuid: "term-unknown", segment: "seg-unknown" })],
      resolve,
    );
    expect(s.groups[0].places[0].tierChain).toEqual([]);
    expect(s.groups[0].places[0].segmentLabel).toBe("");
  });

  it("defaults a blank scope to Folder and tolerates missing fields", () => {
    const s = summarizeUserAccess([g(42, "X")], [{ groupId: 42, role: "UPL" }], resolve);
    expect(s.groups[0].places[0].scope).toBe("Folder");
    expect(s.groups[0].places[0].target).toBe("");
  });

  it("says so when the person is in no groups", () => {
    const s = summarizeUserAccess([], [row({ groupId: 42 })], resolve);
    expect(s.none).toBe(true);
    expect(s.groups).toEqual([]);
  });

  it("ignores rows belonging to groups the person is not in", () => {
    const s = summarizeUserAccess([g(42, "X")], [row({ groupId: 99 })], resolve);
    expect(s.groups[0].unmapped).toBe(true);
  });
});

describe("describeGroupAccess", () => {
  const base = {
    groupId: 1,
    groupTitle: "X",
    places: [],
    unmapped: false,
    siteEntry: false,
  };

  it("prefers the persona label", () => {
    expect(
      describeGroupAccess({
        ...base,
        roles: ["UPL"],
        roleLabels: ["Upload"],
        persona: PERSONAS.filter((p) => p.key === "pic")[0],
      }),
    ).toBe("Upload only");
  });

  // An unrecognised role set must still say something true — a hand-made group with one odd role is
  // exactly the case an admin is investigating.
  it("falls back to the role labels, never to blank", () => {
    expect(describeGroupAccess({ ...base, roles: ["APR"], roleLabels: ["Approve"] })).toBe("Approve");
  });

  it("says something even with rows but no roles", () => {
    const d = describeGroupAccess({ ...base, roles: [], roleLabels: [] });
    expect(d.length).toBeGreaterThan(0);
    expect(d).toContain("No roles recorded");
  });
});

describe("personDisplay", () => {
  // The reported symptom, twice over: a guest whose display name is its own local part, and a
  // free-text address that resolves to itself.
  it("drops a display name that is only the address again", () => {
    expect(personDisplay("clarencechojinheng", "clarencechojinheng@gmail.com"))
      .toBe("clarencechojinheng@gmail.com");
    expect(personDisplay("test@gmail.com", "test@gmail.com")).toBe("test@gmail.com");
    expect(personDisplay("CLARENCECHOJINHENG", "clarencechojinheng@gmail.com"))
      .toBe("clarencechojinheng@gmail.com");
  });

  // A real name is the USEFUL half and must survive.
  it("keeps a real name alongside the address", () => {
    expect(personDisplay("Wyanet Falasifa", "wyanet@trinergydigital.com"))
      .toBe("Wyanet Falasifa · wyanet@trinergydigital.com");
  });

  it("never renders blank, whichever half is missing", () => {
    expect(personDisplay("Crystal Kong", "")).toBe("Crystal Kong");
    expect(personDisplay("", "crystal@trinergydigital.com")).toBe("crystal@trinergydigital.com");
    expect(personDisplay(undefined, undefined)).toBe("");
    expect(personDisplay("  ", " a@b.com ")).toBe("a@b.com");
  });

  // A prefix match is not enough: "clarence" must not swallow a colleague's address.
  it("matches the local part exactly, not as a prefix", () => {
    expect(personDisplay("clarence", "clarencechojinheng@gmail.com"))
      .toBe("clarence · clarencechojinheng@gmail.com");
  });
});
