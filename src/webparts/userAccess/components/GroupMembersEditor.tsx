// Who is in one SharePoint group — read, add, remove.
//
// Spec: docs/superpowers/specs/2026-08-18-folder-access-membership-design.md §3
//
// MOVED here from GroupManager, not copied. Membership now lives on Folder Access, because putting a
// person into their unit's group is the one genuinely recurring job in this system, and it sat on a
// page an admin otherwise visits twice in a segment's life. Two lists of the same people drift — and
// re-merging the two pages is exactly what the client called confusing on 2026-08-14 — so there is
// one implementation and one mount point.
//
// ⚠ THIS IS A GROUP CHANGE, and can only be. A group's grants are held at FOLDER scope, so adding
// someone here gives them that unit's folder in every library the group is mapped to; removing them
// takes all of it away. It is never scoped to the screen it is rendered on, however the surrounding
// page reads, which is why the caveat is repeated in the hint and in every toast.
//
// NOT accessMemberUi.tsx, which is removal-only and built around the library/page grant rows on
// StagingAccess and PageAccess — including a removalVerdict that has no meaning here. Those screens
// gain no Add, deliberately: adding a person there would look scoped to that library or page, and
// never is.
import * as React from "react";
import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  getGroupMembers,
  removeGroupMember,
  ensureSiteUser,
  setSiteAdmin,
  countSiteAdmins,
  searchTenantPeople,
  SpGroupMember,
  PersonPick,
} from "../../../shared/spGroups";
import { addMemberWithSiteEntry } from "../../../shared/siteEntryGroup";
import {
  siteEntryGroupTitle,
  isSiteEntryGroupTitle,
} from "../../../shared/groupMapModel";
import { personDisplay } from "../../../shared/userAccess";
import { EVENT } from "../../../shared/auditLog";
import { writeAudit } from "../../../shared/spAuditLog";

type Props = {
  context: WebPartContext;
  siteUrl: string;
  /** The group being edited. Identified by ID — a stored name can be stale or duplicated. */
  group: { id: number; title: string };
  /** Surfaced by the host, so one page has one toast. */
  showToast: (message: string, error: boolean) => void;
  /**
   * Fired after a successful add or remove, so a host showing a member COUNT beside the group can
   * re-read it. Optional: Folder Access shows no count and passes nothing.
   *
   * A screen that reads a list at mount lies about any write beside it — the fourth instance of that
   * in this feature (the flow rail's tick, the mapping badges, the segment picker's Refresh).
   */
  onChanged?: () => void;
  /**
   * Membership of this group also carries SITE COLLECTION ADMINISTRATOR.
   *
   * Client, 2026-08-27: *"its ok to allow System Admin to have SCA, their version of system admin
   * means they have access to everything."* Offered as two separate decisions and deliberately
   * coupled, because in their model the role IS "access to everything".
   *
   * WARN: SET ON THE OWNERS MOUNT ONLY. This component is mounted for EVERY group on Group
   * Management and again on Folder Access; passing it anywhere else would promote a unit's uploader
   * to site collection administrator.
   *
   * WARN: THE PROMOTION IS REPORTED SEPARATELY AND NEVER FAILS THE ADD. Only an existing SCA may
   * promote, and a guest may be ineligible on some tenants - so a refusal is a real outcome, and the
   * person IS in the group either way. Silently swallowing it leaves an administrator who cannot
   * administer, with nothing on screen to say why.
   */
  alsoSiteAdmin?: boolean;
};

const s: Record<string, React.CSSProperties> = {
  wrap: {
    padding: "8px 10px 14px 24px",
    background: "#fdf3f4",
    borderTop: "1px solid #f5f4f4",
    marginTop: 10,
  },
  row: { display: "flex", alignItems: "center", gap: 10, padding: "4px 0" },
  name: { flex: 1, fontSize: 13 },
  sub: { color: "#666", fontSize: 12 },
  ghost: {
    padding: "5px 12px",
    fontSize: 12,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  danger: {
    padding: "5px 12px",
    fontSize: 12,
    color: "#a4262c",
    border: "1px solid #a4262c",
    borderRadius: 4,
    background: "#fff",
    cursor: "pointer",
  },
  input: {
    width: "100%",
    maxWidth: 460,
    boxSizing: "border-box",
    padding: "7px 10px",
    fontSize: 13,
    border: "1px solid #c7c7c7",
    borderRadius: 4,
  },
  ddwrap: { position: "relative", maxWidth: 460, marginTop: 10 },
  dd: {
    position: "absolute",
    top: "100%",
    left: 0,
    right: 0,
    zIndex: 20,
    background: "#fff",
    border: "1px solid #c7c7c7",
    borderRadius: 4,
    maxHeight: 240,
    overflowY: "auto",
    boxShadow: "0 4px 12px rgba(0,0,0,.12)",
  },
  ddItem: {
    padding: "7px 10px",
    cursor: "pointer",
    borderBottom: "1px solid #f0f0f0",
    fontSize: 13,
  },
  inputRow: { display: "flex", gap: 8, alignItems: "center", maxWidth: 460 },
  inputBad: { borderColor: "#a4262c", background: "#fdf3f4" },
  hint: { fontSize: 11, color: "#666", marginTop: 6, lineHeight: 1.45 },
  err: { fontSize: 11.5, color: "#a4262c", marginTop: 6, lineHeight: 1.45 },
};

export default function GroupMembersEditor({
  context,
  siteUrl,
  group,
  showToast,
  onChanged,
  alsoSiteAdmin,
}: Props): React.ReactElement {
  /**
   * THREE states, never two. `undefined` is loading and `failed` is a read that did not work —
   * neither is "nobody is in this group". An admin told a group is empty stops looking for the person
   * they came for, and on this page they would then add a second copy of someone already in it.
   */
  const [members, setMembers] = useState<SpGroupMember[] | undefined>(
    undefined,
  );
  const [failed, setFailed] = useState(false);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<PersonPick[]>([]);
  const [confirmRemove, setConfirmRemove] = useState<number | undefined>(
    undefined,
  );
  const [busy, setBusy] = useState(false);
  /**
   * Why the last add did not work, IN PLAIN WORDS, shown under the box.
   *
   * Client, 2026-08-23: *"When an email can't be found then show an error not a toast with error
   * client cannot understand"*. The toast carried SharePoint's own text —
   * `HTTP 404 {"error":{"code":"-2130575276, Microsoft.SharePoint.SPException","message":"The user
   * does not exist or is not unique."}}` — next to the box the address was typed into, then
   * vanished after seven seconds. An error about ONE field belongs beside that field, and it must
   * name the fix rather than the exception.
   */
  const [addError, setAddError] = useState<string | undefined>(undefined);

  const reload = async (): Promise<void> => {
    try {
      setMembers(
        await getGroupMembers(context.spHttpClient, siteUrl, group.id),
      );
      setFailed(false);
    } catch {
      setMembers(undefined);
      setFailed(true);
    }
  };

  // Keyed on the group id, so expanding a different group re-reads rather than showing the last
  // one's people under this one's name.
  useEffect(() => {
    setMembers(undefined);
    setFailed(false);
    setQuery("");
    setResults([]);
    setConfirmRemove(undefined);
    setAddError(undefined);
    reload().catch(() => undefined);
  }, [group.id]);

  useEffect(() => {
    const q = query.trim();
    if (q.length < 2) {
      setResults([]);
      return undefined;
    }
    const t = setTimeout(() => {
      searchTenantPeople(context.spHttpClient, siteUrl, q)
        .then(setResults)
        .catch(() => setResults([]));
    }, 350);
    return () => clearTimeout(t);
  }, [query]);

  /** Fire-and-forget: the act has happened, and no admin action should fail because the log was down. */
  const log = (
    summary: string,
    details: string[],
    outcome: "Success" | "Failed" = "Success",
  ): void => {
    writeAudit(context.spHttpClient, siteUrl, {
      event: EVENT.membersChanged,
      outcome,
      source: "FolderAccess",
      at: new Date(),
      actorName: context.pageContext.user.displayName,
      actorEmail: context.pageContext.user.email,
      summary,
      details,
    }).catch(() => undefined);
  };

  /**
   * SharePoint's add-member failures, translated.
   *
   * `The user does not exist or is not unique` is by far the common one and means one of two very
   * different things: the address belongs to nobody (a typo, or somebody outside the organisation who
   * has never been invited), or it matches MORE than one account — which happens here, because a
   * guest can exist twice for one address (observed 2026-08-18). Both are said, because the admin
   * cannot tell which from the outside and the fix differs.
   *
   * The raw text is kept on the end for anything unrecognised: an unknown error must stay visible.
   */
  const explainAddFailure = (person: string, raw: string): string => {
    /* ⚠ THIS BRANCH ALSO PROMISED AN INVITATION and was corrected with the one below (2026-08-27):
       it read *"they have to be invited to the site once before they can be put in a group"*, which
       describes a step nobody here is allowed to take. The duplicate-account half is kept — it is
       real on this tenant, where one address can exist twice as a guest. */
    if (/does not exist or is not unique/i.test(raw)) {
      return (
        `${person} could not be added — SharePoint does not recognise that address on this site.` +
        ` Check the spelling. Only people who already have an account here can be put in a group;` +
        ` somebody from outside the organisation cannot be added from this page. The same address` +
        ` can also match two accounts, in which case pick the person from the list instead of` +
        ` typing the address.`
      );
    }
    /* ⚠ `ensureuser` USES A DIFFERENT WORDING, and the first build fell straight through to the raw
       JSON because of it (seen on site 2026-08-27):
         "The specified user reenelow@hotmail.com could not be found."
       versus the people-picker's "The user does not exist or is not unique". Same situation, two
       messages - so matching one pattern left the other showing an SPException code to an admin who
       only typed an address. THE LESSON: every route into this function needs its own phrase checked,
       not assumed. */
    /* ⚠ THIS IS THE MESSAGE THE CLIENT ASKED FOR: *"show an error message why it is not allowed."*
       It has been wrong in two OPPOSITE directions in one day, and both were misleading:
         1. It asserted this site's guest sharing was OFF — a guess dressed as a diagnosis, disproved
            by an external address that was already sitting in a group.
         2. It then told the admin to press Add because that would invite them — true for about an
            hour, until the client ruled out adding outsiders at all.
       So it now states the RULE, and says outright that this is a decision rather than a fault. That
       last clause is the load-bearing one: without it an admin reports the refusal as a bug, which is
       exactly what happened twice before anybody had written the rule down.

       ⚠ IT NAMES BOTH CAUSES, BECAUSE THEY CANNOT BE TOLD APART FROM THE RESPONSE. SharePoint answers
       identically for a typo and for a real person with no account here, and a domain check could not
       separate them either — `MySubmissions`'s `tenantDomains` already counts a guest's own domain as
       internal. Naming one cause would send half the admins looking the wrong way. */
    if (/could not be found|cannot be found/i.test(raw)) {
      return (
        `${person} could not be added — there is no account for that address on this site.` +
        ` Only people who already have an account here can be put in a group, and somebody from` +
        ` outside the organisation cannot be added or invited from this page. That is deliberate,` +
        ` not a fault. If ${person} should have access: check the spelling first, then ask IT to set` +
        ` up their account — once it exists they will appear in the list of suggestions.`
      );
    }
    if (/403|access denied|unauthorized/i.test(raw)) {
      return (
        `${person} could not be added — this account does not have permission to change group` +
        ` membership. Managing groups needs Full Control on the site.`
      );
    }
    return `${person} could not be added. ${raw}`;
  };

  /** Anything shaped like an address. Deliberately loose — SharePoint decides what is real. */
  const looksLikeEmail = (v: string): boolean =>
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v.trim());

  /**
   * ONE definition of "the add failed": the explanation goes INLINE, and the dropdown CLOSES.
   *
   * ⚠ THIS EXISTS BECAUSE THE DROPDOWN-CLOSE WAS MISSED ON THE SECOND ROUTE INTO THIS PATH. It was
   * added to `onAdd`'s catch on 2026-08-23 (client: *"the dropdown doesnt clear making it difficult
   * to see"*) and the Add button, built 2026-08-27, went straight past it — so a failed TYPED add
   * rendered the explanation underneath the absolutely-positioned candidate list that caused it.
   * Reported by the client the same day: *"The dropdown keeps showing even after I add already"*,
   * with the first line of the message hidden behind it.
   *
   * ⚠ SECOND TIME IN ONE DAY that a new route into this failure path missed something the old one
   * already did — the other was `explainAddFailure`'s wording, which matched the people-picker's
   * phrase and not `ensureuser`'s. **Both catches call this now, so a third route cannot miss it.**
   * A copied `setResults([])` would have been a second copy of the rule, which is how this happened.
   *
   * The TYPED TEXT is deliberately kept: a typo is fixed by editing it, not by retyping it.
   */
  const failAdd = (person: string, raw: string): void => {
    setAddError(explainAddFailure(person, raw));
    setResults([]);
  };

  const onAdd = async (p: PersonPick): Promise<void> => {
    setBusy(true);
    setAddError(undefined);
    try {
      const res = await addMemberWithSiteEntry(
        context.spHttpClient,
        siteUrl,
        { id: group.id, title: group.title },
        p.loginName,
      );
      await reload();
      setQuery("");
      setResults([]);
      /* SITE COLLECTION ADMIN, on the Owners mount only. Reported on its own line and never allowed
         to fail the add: only an existing SCA may promote, and a guest may be ineligible, so this
         genuinely can refuse while the person IS in the group. */
      let scaNote = "";
      if (alsoSiteAdmin === true) {
        try {
          await setSiteAdmin(context.spHttpClient, siteUrl, p.loginName, true);
          /* ⚠ SUCCESS SAYS NOTHING NOW (client, 2026-09-04: the toast was too long). Safe to drop
             ONLY because the card this mount sits on states it permanently and in bolder terms:
             *"Warning: Adding a user also makes them a Site Collection Administrator..."*. The
             FAILURE below is NOT dropped — see there. */
          scaNote = "";
        } catch (e) {
          scaNote =
            ` They were NOT made a site collection administrator (${(e as Error).message}).` +
            " Only an existing site collection administrator can grant that, and some tenants refuse" +
            " guest accounts. Add them by hand in Site settings if it is needed.";
        }
      }
      // res.note is never dropped: it means they ARE in the group but may not be able to open the
      // site, which neither they nor the admin would discover until they tried.
      // ⚠ THE TOAST MUST SAY WHAT THEY NOW REACH (client, 2026-08-23: *"The toast should tell them
      // they also have access to the folder"*). It said only "added to <group>", which reads as a
      // list change — and an admin who does not realise this hands over FOLDER access has just
      // granted documents without knowing it. It is the same sentence the hint below the box makes,
      // said at the moment the change actually happens.
      /* ⚠ TWO SENTENCES COLLAPSED TO TWO WORDS (client, 2026-09-04: *"just make it if a user added
         or remove it shows — [user] added, [user] remove"*).

         WHAT WENT, AND WHY IT IS SURVIVABLE: the toast used to state that the person now has this
         group's folder access in every library it is mapped to, that they can open the site, and
         that they may need to sign out and back in. All three are still TRUE and all three are still
         written to the AUDIT ROW below, which is the actual record — a toast is a receipt, not the
         documentation. The sign-out caveat is the one genuinely worth missing (it cost an hour of
         diagnosis on 2026-08-18), so if "I added them and they still cannot get in" starts arriving,
         that sentence is the answer and this is where it was removed from.

         ⚠ A PROBLEM STILL SPEAKS IN FULL, and that is the half that must not be shortened. `res.note`
         means they ARE in the group but may not be able to open the site at all, and a failed SCA
         promotion leaves an administrator who cannot administer — neither would be discovered by
         anyone until they tried. A short toast is for the ordinary case; there is no short way to say
         "this half worked". */
      const addProblem = (res.note ? ` ${res.note}` : "") + scaNote;
      showToast(
        `${p.displayName} added` + addProblem,
        !!res.note || scaNote.indexOf("NOT made") !== -1,
      );
      if (onChanged) onChanged();
      log(
        `Member added — ${p.displayName} into ${group.title}`,
        [
          `Group: ${group.title}`,
          `Added: ${p.displayName}${p.email ? ` <${p.email}>` : ""}`,
          isSiteEntryGroupTitle(group.title)
            ? "This IS the site-entry group."
            : res.note || `Also added to ${siteEntryGroupTitle()}.`,
        ],
        res.note ? "Failed" : "Success",
      );
    } catch (e) {
      // INLINE, not a toast: it is about the box directly above it, and it must still be on screen
      // when the admin looks back at what they typed. `failAdd` also closes the dropdown — see there
      // for why that is shared rather than written out here.
      failAdd(personDisplay(p.displayName, p.email), (e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  /* DECLARED AFTER `onAdd` because it calls it, and `no-use-before-define` is on. */
  /**
   * Add whatever was TYPED, without waiting for the picker to suggest it.
   *
   * Client, 2026-08-27: *"I cannot add new email at all, is there no button for this."* Adding was
   * only possible by clicking a dropdown candidate, and `searchTenantPeople` suggests nothing for
   * some addresses this tenant's picker will not surface — so an existing guest could not be added
   * at all, and there was no button and no Enter handler to fall back on.
   *
   * ⚠⚠ IT RESOLVES ONLY. IT MUST NOT INVITE, AND THAT IS A POLICY DECISION RATHER THAN A LIMIT.
   * An invite path was built the same afternoon (`inviteToGroup`, `SP.Web.ShareObject` with
   * `roleValue: "group:<id>"`) and REVERTED within the hour on the client's instruction: *"client
   * doesn't want to allow us to add outsiders, so we got to revert again."* So `ensureuser` answering
   * `could not be found` for somebody from outside is now the CORRECT and FINAL outcome — a refusal
   * to be EXPLAINED, not a gap to close. Do not re-add the invite without the client asking.
   *
   * ⚠ WHAT THE REVERT HAD TO CHANGE BESIDES DELETING THE CALL — each of these otherwise goes on
   * describing a capability the page no longer has:
   *   - BOTH branches of `explainAddFailure` told the admin an outsider "has to be invited to the
   *     site once", one of them naming the Add button as the way to do it.
   *   - The hint under the box said an outside address is "invited and put in this group in one step".
   * A message that promises a button will do something it cannot is worse than no message: the admin
   * presses it, gets the same refusal, and reports a bug — which is how this arrived twice already.
   *
   * The button still earns its place: an address that EXISTS but which the picker declines to
   * surface — an existing guest, or a work account it fails to match — is added by typing it.
   */
  const onAddTyped = async (): Promise<void> => {
    const typed = query.trim();
    if (!looksLikeEmail(typed)) return;
    setBusy(true);
    setAddError(undefined);
    try {
      const person = await ensureSiteUser(context.spHttpClient, siteUrl, typed);
      setBusy(false);
      await onAdd(person);
    } catch (e) {
      setBusy(false);
      // Closes the dropdown as well as showing the reason: it is absolutely positioned over this
      // message, and the client hit exactly that on 2026-08-27.
      failAdd(typed, (e as Error).message);
    }
  };

  const onRemove = async (m: SpGroupMember): Promise<void> => {
    setBusy(true);
    try {
      await removeGroupMember(context.spHttpClient, siteUrl, group.id, m.id);
      /* SITE COLLECTION ADMIN GOES WITH IT, on the Owners mount only. Leaving it behind would mean
         somebody removed from the administrators group keeps full control of the site collection -
         worse than never coupling them, because the screen would say the access was revoked.

         ⚠ TWO GUARDS, AND BOTH PREVENT A LOCK-OUT NOTHING IN THIS APP COULD UNDO:
           1. NEVER THE SIGNED-IN USER. Demoting yourself mid-action removes the very right this call
              needs, so anything after it fails and the state is half-changed.
           2. NEVER THE LAST ONE. A site collection with no administrator cannot be administered, and
              no screen here could put one back — only a tenant administrator could.
         `countSiteAdmins` answers `undefined` when it could not count, and an uncountable answer
         REFUSES the demotion. Guessing "there must be others" is exactly how the site is lost. */
      let scaNote = "";
      if (alsoSiteAdmin === true) {
        const mine = (context.pageContext.user.loginName ?? "").toLowerCase();
        if (m.loginName.toLowerCase() === mine) {
          scaNote =
            " Their site collection administrator rights were left in place: this would have" +
            " removed your own, and you would lose the rights this page needs.";
        } else {
          const admins = await countSiteAdmins(context.spHttpClient, siteUrl);
          if (admins === undefined) {
            scaNote =
              " Their site collection administrator rights were left in place: the number of" +
              " administrators could not be read, and removing the last one leaves nobody able to" +
              " administer this site.";
          } else if (admins <= 1) {
            scaNote =
              " Their site collection administrator rights were left in place: they are the" +
              " only one, and a site with no administrator cannot be recovered from here.";
          } else {
            try {
              await setSiteAdmin(
                context.spHttpClient,
                siteUrl,
                m.loginName,
                false,
              );
              /* Silent on success, as with the add: the ordinary outcome is what the button said it
                 would do. The FAILURE branch below stays in full — it leaves someone holding full
                 control of the site after a screen told the admin their access was revoked. */
              scaNote = "";
            } catch (e) {
              scaNote =
                ` ⚠ Their site collection administrator rights could NOT be removed` +
                ` (${(e as Error).message}) — they still have full control of this site. Remove them` +
                ` by hand in Site settings.`;
            }
          }
        }
      }
      await reload();
      setConfirmRemove(undefined);
      /* Shortened with the add above. The dropped clause — *"that group's folder access is revoked
         immediately"* — is a correction of a real misconception (deleting a MAPPING leaves the grant
         until reconciliation runs; removing a PERSON does not), and it survives in the audit row.
         ⚠ The failure note is kept: "could NOT be removed" means they still have full control. */
      showToast(
        `${m.title} removed` + scaNote,
        scaNote.indexOf("could NOT") !== -1 ||
          scaNote.indexOf("left in place") !== -1,
      );
      if (onChanged) onChanged();
      log(`Member removed — ${m.title} from ${group.title}`, [
        `Group: ${group.title}`,
        `Removed: ${m.title}${m.email ? ` <${m.email}>` : ""}`,
        // Removing from one group is not removing from the site. Saying so stops a reader
        // concluding the person was de-provisioned.
        `This removes ONE group. Their membership of ${siteEntryGroupTitle()} and any other group is unchanged.`,
      ]);
    } catch (e) {
      showToast(`Could not remove ${m.title}: ${(e as Error).message}`, true);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={s.wrap}>
      {members === undefined && !failed && (
        <p style={s.hint}>Loading members&hellip;</p>
      )}
      {failed && (
        <p style={s.err}>
          Could not read this group&apos;s members — so this is{" "}
          <strong>not</strong> a statement that nobody is in it. Adding someone
          still works.
        </p>
      )}
      {members !== undefined && members.length === 0 && (
        <p style={s.hint}>
          Nobody is in this group yet. Its folder permissions exist and apply
          the moment someone is added.
        </p>
      )}
      {(members ?? []).map((m) => (
        <div key={m.id} style={s.row}>
          <span style={s.name}>{personDisplay(m.title, m.email)}</span>
          {confirmRemove === m.id ? (
            <>
              <button
                type="button"
                style={s.danger}
                disabled={busy}
                onClick={() => {
                  onRemove(m).catch(() => undefined);
                }}
              >
                Confirm remove
              </button>
              <button
                type="button"
                style={s.ghost}
                onClick={() => setConfirmRemove(undefined)}
              >
                Cancel
              </button>
            </>
          ) : (
            <button
              type="button"
              style={s.ghost}
              disabled={busy}
              onClick={() => setConfirmRemove(m.id)}
            >
              Remove
            </button>
          )}
        </div>
      ))}

      <div style={s.ddwrap}>
        <div style={s.inputRow}>
          <input
            // Red while the last add failed — the client could not see which field the message was
            // about. Cleared by typing, because the next keystroke starts a new search.
            style={
              addError === undefined ? s.input : { ...s.input, ...s.inputBad }
            }
            value={query}
            placeholder="Add a person&hellip;"
            disabled={busy}
            onChange={(e) => {
              setQuery(e.target.value);
              setAddError(undefined);
            }}
            onKeyDown={(e) => {
              if (e.key === "Escape") {
                setQuery("");
                setResults([]);
                setAddError(undefined);
              }
              // Enter adds the typed address — the keystroke every admin tries first.
              if (e.key === "Enter" && looksLikeEmail(query)) {
                e.preventDefault();
                onAddTyped().catch(() => undefined);
              }
            }}
          />
          {/* There was NO way to empty this box (client: *"once I select the email I cannot clear
              it"*) — selecting a candidate leaves the text behind, and on a failed add it stays with
              the dropdown over it. Rendered only when there is something to clear. */}
          {/* ADD, for an address that EXISTS but which the picker will not suggest — an existing
              guest, or a work account it fails to match. Shown only when the text is address-shaped,
              so it never invites a click that cannot work. ⚠ IT DOES NOT INVITE ANYBODY: an invite
              path was built and reverted on 2026-08-27 (see `onAddTyped`). Picking a candidate from
              the dropdown remains the normal path. */}
          {looksLikeEmail(query) && (
            <button
              type="button"
              style={s.ghost}
              disabled={busy}
              onClick={() => {
                onAddTyped().catch(() => undefined);
              }}
            >
              Add
            </button>
          )}
          {query.length > 0 && (
            <button
              type="button"
              style={s.ghost}
              disabled={busy}
              onClick={() => {
                setQuery("");
                setResults([]);
                setAddError(undefined);
              }}
            >
              Clear
            </button>
          )}
        </div>
        {results.length > 0 && (
          <div style={s.dd}>
            {results.map((p) => (
              <div
                key={p.loginName}
                style={s.ddItem}
                onClick={() => {
                  onAdd(p).catch(() => undefined);
                }}
              >
                {personDisplay(p.displayName, p.email)}
              </div>
            ))}
          </div>
        )}
      </div>
      {addError !== undefined && <p style={s.err}>{addError}</p>}

      {/* ⚠ THREE PARAGRAPHS BECAME ONE (client, 2026-09-04), and two real facts went with them.
          Their replacement copy, verbatim. Recorded here because none of the three was decoration and
          the last one in particular has already cost somebody an hour:

          1. THE SITE-ENTRY AUTO-JOIN. `addMemberWithSiteEntry` still adds every person to the site
             entry group as well — the behaviour is unchanged, it is simply no longer stated. Nobody
             reading this page now learns that adding someone to one group also grants them site
             entry, or that removal is ASYMMETRIC (`onRemove` does not take it away again).
          2. THAT MEMBERSHIP IS A *GROUP* CHANGE, reaching every folder the group is mapped to in every
             library — not only whatever the screen above happens to list. The add toast still says it.
          3. ⚠ THE SIGN-OUT/SIGN-IN GOTCHA, OBSERVED ON SITE 2026-08-18 AND THE ONE WORTH ARGUING
             ABOUT. A guest who already had the site open was bounced off a page their new group
             grants, while EVERY permission read back correct — group membership, the page role
             assignment, the Read binding, even their own `currentuser/groups`. A full sign-out and
             sign-in fixed it instantly. It is the signed-in SESSION that is stale, not the
             permissions. With this line gone, the next person who hits it has nothing on screen to
             suggest the tool worked, and the diagnosis starts again from zero.

          If "I added them and they still cannot get in" becomes a support call, this paragraph was
          the answer and putting it back is one line. */}
      <p style={s.hint}>
        <strong>Note:</strong> Only people who already have an account on this
        site can be added. External users must have an account set up first.
      </p>
    </div>
  );
}
