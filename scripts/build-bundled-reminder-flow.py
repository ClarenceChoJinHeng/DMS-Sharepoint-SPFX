"""Builds the 'CRS — Remind approvers (bundled)' import package (design: vault Specs/2026-09-30-bundled-reminder-email-design.md).
Replaces 'CRS - Reminding Approver to Approve' + 'CRS - HC approval reminder' (one email per file per approver).
Same email + audiences as the bundled approver email; subject 'REMINDER: Approval Required: …'.
Keeps the old reminder's rules: every 30 min, due 3 days after upload then every 3 days (`NextReminderAt`),
no bulk imports, no uploader, no gdc, no site admins.
Usage: python scripts/build-bundled-reminder-flow.py test|crs   (writes the import zip to C:/tmp/flows-for-test/)
"""
import importlib.util, os, sys

_spec = importlib.util.spec_from_file_location(
    "approver", os.path.join(os.path.dirname(__file__), "build-bundled-approver-flow.py"))
ap = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ap)

DISPLAY_NAME = "CRS — Remind approvers (bundled)"
OUT = {
    "test": r"C:\tmp\flows-for-test\TEST_CRS—Remindapproversbundled.zip",
    "crs": r"C:\tmp\flows-for-test\CRS—Remindapproversbundled.zip",
}
ALL_PAGES = {"paginationPolicy": {"minimumItemCount": 5000}}


def blank(expr):
    return "empty(coalesce(" + expr + ", ''))"


# One upload + one unit. Old files without SubmissionId fall back to their own file id, then their path.
UPLOAD_KEY = ("concat(if(" + blank("item()?['SubmissionId']") + ", if(" + blank("item()?['SubmissionFileId']") + ", "
              "item()?['{FullPath}'], item()?['SubmissionFileId']), item()?['SubmissionId']), '#', "
              "coalesce(item()?['UnitTid'], ''))")
# Due = NextReminderAt, or 3 days after upload when it was never reminded.
DUE_AT = ("if(" + blank("item()?['NextReminderAt']") + ", addDays(item()?['Created'], 3), item()?['NextReminderAt'])")


def subject(files):
    return ("@if(equals(length(" + files + "), 1), concat('REMINDER: Approval Required: ', first(" + files + ")?['{FilenameWithExtension}']), "
            "concat('REMINDER: Approval Required: ', string(length(" + files + ")), ' documents'))")


def send_email(t, to, files, show_normal, show_hc):
    params = {"emailMessage/To": to, "emailMessage/Subject": subject(files),
              "emailMessage/Body": ap.email_body(t["site"], files, show_normal, show_hc),
              "emailMessage/Importance": "Normal"}
    if t["bcc"]:
        params["emailMessage/Bcc"] = t["bcc"]
    return {"type": "OpenApiConnection", "runAfter": {},
            "inputs": {"parameters": params, "host": ap.MAIL_HOST, "authentication": "@parameters('$authentication')"}}


def definition(t):
    site = t["site"]
    after = ap.after
    has_normal_files = "greater(length(body('NormalFiles')), 0)"
    has_hc_files = "greater(length(body('HcFiles')), 0)"
    hc_audience = "greater(length(outputs('HcEmails')), 0)"
    normal_audience = "and(greater(length(outputs('NormalEmails')), 0), " + has_normal_files + ")"

    each_group = {
        "type": "Foreach", "foreach": "@body('GetApproverGroups')?['value']", "runAfter": after("Reset_Members"),
        "runtimeConfiguration": ap.ONE_AT_A_TIME,
        "actions": {
            "GetMembers": ap.sp_http(site, "GET",
                                     "_api/web/sitegroups(@{items('Each_group')?['GroupId']})/users?$select=Email,IsSiteAdmin", {}),
            "Each_member": {"type": "Foreach", "foreach": "@coalesce(body('GetMembers')?['value'], createArray())",
                            "runAfter": after("GetMembers"), "runtimeConfiguration": ap.ONE_AT_A_TIME,
                            "actions": {"Append_member": {"type": "AppendToArrayVariable", "runAfter": {},
                                                          "inputs": {"name": "Members",
                                                                     "value": "@addProperty(item(), 'Role', items('Each_group')?['Role'])"}}}},
        },
    }
    has_approver_actions = {
        "Reset_Members": {"type": "SetVariable", "runAfter": {}, "inputs": {"name": "Members", "value": []}},
        "Each_group": each_group,
        "PickEmails": {"type": "Query", "runAfter": after("Each_group"),
                       "inputs": {"from": "@variables('Members')",
                                  "where": ("@and(not(empty(coalesce(item()?['Email'], ''))), "
                                            "not(equals(toLower(item()?['Email']), toLower(coalesce(first(body('KeyFiles'))?['Author']?['Email'], '---')))), "
                                            "not(equals(toLower(item()?['Email']), 'gdc@sdguthrie.com')), "
                                            "not(equals(item()?['IsSiteAdmin'], true)))")}},
    }
    has_approver_actions.update(ap.distinct_emails("Hc", "PickEmails", "@equals(item()?['Role'], 'APRHC')"))
    has_approver_actions.update(ap.distinct_emails(
        "Normal", "HcEmails",
        "@and(equals(item()?['Role'], 'APR'), not(contains(outputs('HcEmails'), toLower(item()?['Email']))))"))
    has_approver_actions.update({
        "NormalFiles": {"type": "Query", "runAfter": after("NormalEmails"),
                        "inputs": {"from": "@body('KeyFiles')", "where": "@startsWith(item()?['{Path}'], 'ApprovalDocument/')"}},
        "HcFiles": {"type": "Query", "runAfter": after("NormalFiles"),
                    "inputs": {"from": "@body('KeyFiles')", "where": "@not(startsWith(item()?['{Path}'], 'ApprovalDocument/'))"}},
        "SendToHc": {"type": "If", "runAfter": after("HcFiles"), "expression": ap.is_true("@" + hc_audience),
                     "actions": {"Send_hc_email": send_email(t, "@join(outputs('HcEmails'), ';')", "body('KeyFiles')",
                                                             has_normal_files, has_hc_files)},
                     "else": {"actions": {}}},
        "SendToNormal": {"type": "If", "runAfter": after("SendToHc"), "expression": ap.is_true("@" + normal_audience),
                         "actions": {"Send_normal_email": send_email(t, "@join(outputs('NormalEmails'), ';')",
                                                                     "body('NormalFiles')", "true", "false")},
                         "else": {"actions": {}}},
    })
    has_approver = {
        "type": "If", "runAfter": after("GetApproverGroups"),
        "expression": ap.is_true("@greater(length(coalesce(body('GetApproverGroups')?['value'], createArray())), 0)"),
        "actions": has_approver_actions,
        "else": {"actions": {}},
    }
    # Every pending file of the upload gets the same next date, so the set stays together next time.
    # Runs even when no email went out (no unit, no approvers), as the old reminder did.
    list_of_file = ("if(startsWith(items('Each_file')?['{Path}'], 'ApprovalDocument/'), "
                    "'" + t["approval_list"] + "', '" + t["hc_approval_list"] + "')")
    set_next_reminder = {
        "type": "Foreach", "foreach": "@body('KeyFiles')", "runAfter": {"HasApprover": ["Succeeded", "Failed", "Skipped"]},
        "runtimeConfiguration": ap.ONE_AT_A_TIME,
        "actions": {
            "SetNextReminder": ap.sp_http(site, "POST",
                                          "_api/web/lists(guid'@{" + list_of_file + "}')/items(@{items('Each_file')?['ID']})",
                                          {}, "{ \"NextReminderAt\": \"@{addDays(utcNow(), 3)}\" }"),
        },
    }
    each_key = {
        "type": "Foreach", "foreach": "@union(body('DueKeys'), body('DueKeys'))", "runAfter": after("DueKeys"),
        "runtimeConfiguration": ap.ONE_AT_A_TIME,
        "actions": {
            # All pending files of this upload + unit, due or not.
            "KeyFiles": {"type": "Query", "runAfter": {},
                         "inputs": {"from": "@body('Live')", "where": "@equals(" + UPLOAD_KEY + ", items('Each_key'))"}},
            "GetApproverGroups": ap.sp_http(site, "GET",
                                            "_api/web/lists/getbytitle('CRS Group Map')/items?$select=GroupId,Role&$filter=(Role eq 'APR' or Role eq 'APRHC') and UnitTermGuid eq '@{first(body('KeyFiles'))?['UnitTid']}'&$top=10",
                                            after("KeyFiles")),
            "HasApprover": has_approver,
            "Each_file": set_next_reminder,
        },
    }
    pending = {"$filter": "OData__ModerationStatus eq 2", "$top": 5000}
    actions = {
        "Init_Members": {"type": "InitializeVariable", "runAfter": {},
                         "inputs": {"variables": [{"name": "Members", "type": "array", "value": []}]}},
        "GetPending": dict(ap.sp("GetItems", dict(pending, dataset=site, table=t["approval_list"]), after("Init_Members")),
                           runtimeConfiguration=ALL_PAGES),
        "GetPendingHC": dict(ap.sp("GetItems", dict(pending, dataset=site, table=t["hc_approval_list"]), after("GetPending")),
                             runtimeConfiguration=ALL_PAGES),
        "AllPending": {"type": "Compose", "runAfter": after("GetPendingHC"),
                       "inputs": "@union(coalesce(body('GetPending')?['value'], createArray()), coalesce(body('GetPendingHC')?['value'], createArray()))"},
        "Live": {"type": "Query", "runAfter": after("AllPending"),
                 "inputs": {"from": "@outputs('AllPending')",
                            "where": ("@and(equals(item()?['{IsFolder}'], false), not(equals(item()?['BulkImport'], true)), "
                                      "equals(item()?['{ModerationStatus}'], 'Pending'))")}},
        "Due": {"type": "Query", "runAfter": after("Live"),
                "inputs": {"from": "@body('Live')", "where": "@lessOrEquals(ticks(" + DUE_AT + "), ticks(utcNow()))"}},
        "DueKeys": {"type": "Select", "runAfter": after("Due"),
                    "inputs": {"from": "@body('Due')", "select": "@" + UPLOAD_KEY}},
        "Each_key": each_key,
    }
    recurrence = {"frequency": "Minute", "interval": 30, "timeZone": "Singapore Standard Time"}
    return {
        "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {"$connections": {"defaultValue": {}, "type": "Object"},
                       "$authentication": {"defaultValue": {}, "type": "SecureObject"}},
        # One run at a time: an overlapping run (Test + schedule) would send the same reminders twice.
        "triggers": {"Recurrence": {"recurrence": recurrence, "evaluatedRecurrence": recurrence, "type": "Recurrence",
                                    "runtimeConfiguration": {"concurrency": {"runs": 1}}}},
        "actions": actions,
    }


if __name__ == "__main__":
    which = sys.argv[1]
    ap.build(ap.TARGETS[which], DISPLAY_NAME, definition, OUT[which])
