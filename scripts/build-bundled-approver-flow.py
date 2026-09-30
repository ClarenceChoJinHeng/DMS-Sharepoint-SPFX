"""Builds the 'CRS — Notify approvers (bundled)' import package (runbook: vault Specs/2026-09-29-bundled-approver-email-runbook.md, section 4).
Email = the client's template (2026-09-30): count, Submitted By, links to the Pending Files views. No per-set list.
Option B (2026-09-30): one pass per upload + unit covers BOTH approval libraries.
Audiences (2026-09-30): APRHC members get every file + both links; plain APR members get only the normal files + the
normal link (they cannot open HC). Someone in both groups gets the APRHC email only.
Usage: python scripts/build-bundled-approver-flow.py test|crs   (writes the import zip to C:/tmp/flows-for-test/)
"""
import json, sys, uuid, zipfile

TARGETS = {
    "test": {
        "site": "https://dcidigitalcom.sharepoint.com/sites/ClarenceDMSTesting",
        "approval_list": "a9342528-66be-458c-ba70-a6e3248a5133",
        "hc_approval_list": "2311cd83-90aa-4077-a647-65d251a5f426",
        "submissions_list": "f413972a-49ae-476a-8a8f-df4562349918",
        "bcc": "",
        "base": r"C:\tmp\flows-for-test\TEST_NotifyApprovers_20260925035912.zip",
        "out": r"C:\tmp\flows-for-test\TEST_CRS—Notifyapproversbundled.zip",
    },
    "crs": {
        "site": "https://sdguthrie.sharepoint.com/sites/CRS",
        "approval_list": "eeb1bb19-ec53-4c41-8dc9-ecf255979c9b",
        "hc_approval_list": "d935aa6d-dc48-4832-ba7e-eca485ecdff3",
        "submissions_list": "dd1bc5bd-daa7-491a-9dc7-b81212c391ab",
        "bcc": "gdc@sdguthrie.com",
        "base": r"C:\laragon\www\Work\Projects\sd-gatrie\PowerAutomateFlowsSDG\NotifyApprovers_20260925035912.zip",
        "out": r"C:\tmp\flows-for-test\CRS—Notifyapproversbundled.zip",
    },
}
DISPLAY_NAME = "CRS — Notify approvers (bundled)"
SP_HOST = {"apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
           "connectionName": "shared_sharepointonline"}
MAIL_HOST = {"apiId": "/providers/Microsoft.PowerApps/apis/shared_office365",
             "connectionName": "shared_office365", "operationId": "SendEmailV2"}
ONE_AT_A_TIME = {"concurrency": {"repetitions": 1}}


def sp(op, params, run_after=None):
    return {"runAfter": run_after or {}, "type": "OpenApiConnection",
            "inputs": {"parameters": params, "host": dict(SP_HOST, operationId=op),
                       "authentication": "@parameters('$authentication')"}}


def sp_http(site, method, uri, run_after, body=None):
    headers = {"Accept": "application/json;odata=nometadata"}
    params = {"dataset": site, "parameters/method": method, "parameters/uri": uri, "parameters/headers": headers}
    if body is not None:
        headers.update({"Content-Type": "application/json;odata=nometadata", "IF-MATCH": "*", "X-HTTP-Method": "MERGE"})
        params["parameters/body"] = body
    return sp("HttpRequest", params, run_after)


def after(name):
    return {name: ["Succeeded"]}


def is_true(expr):
    return {"and": [{"equals": [expr, True]}]}


SUBMISSION_KEY = ("concat(if(empty(coalesce(item()?['SubmissionId'], '')), item()?['SubmissionFileId'], "
                  "item()?['SubmissionId']), '#', item()?['UnitTid'])")
THREE_DAYS_AGO = "@{formatDateTime(addDays(utcNow(), -3), 'yyyy-MM-ddTHH:mm:ssZ')}"


def email_body(site, files, show_normal, show_hc):
    """The client's template. `files` = the files this audience can open; each link shows only when it applies."""
    normal_link = ("<a href=\"" + site + "/ApprovalDocument/Forms/Pending%20Files.aspx\">"
                   "View Documents – Restricted &amp; Confidential</a><br>")
    hc_link = ("<a href=\"" + site + "/HCApprovalDocument/Forms/Pending%20Files.aspx\">"
               "View Documents – Highly Confidential</a><br>")
    return ("<div style=\"font-family: Arial, sans-serif; font-size: 13px;\">\n"
            "<p>Dear Approver,</p>\n"
            "<p>@{length(" + files + ")} document(s) have been submitted for your review and approval.</p>\n"
            "<p><b>Submitted By:</b> @{first(body('KeyFiles'))?['Author']?['DisplayName']}</p>\n"
            "<p>Please review the submitted files using the links below:<br>\n"
            "@{if(" + show_normal + ", '" + normal_link + "', '')}\n"
            "@{if(" + show_hc + ", '" + hc_link + "', '')}</p>\n"
            "<p>Thank you,<br><b>Guthrie Document Centre</b></p>\n"
            "<p><br>--- This is an auto-generated e-mail. Please do not reply to this email. ---</p>\n"
            "</div>")


def subject(files):
    return ("@if(equals(length(" + files + "), 1), concat('Approval Requested: ', first(" + files + ")?['{FilenameWithExtension}']), "
            "concat('Approval Requested: ', string(length(" + files + ")), ' documents'))")


def send_email(t, to, files, show_normal, show_hc):
    params = {"emailMessage/To": to, "emailMessage/Subject": subject(files),
              "emailMessage/Body": email_body(t["site"], files, show_normal, show_hc),
              "emailMessage/Importance": "Normal"}
    if t["bcc"]:
        params["emailMessage/Bcc"] = t["bcc"]
    return {"type": "OpenApiConnection", "runAfter": {},
            "inputs": {"parameters": params, "host": MAIL_HOST, "authentication": "@parameters('$authentication')"}}


def distinct_emails(prefix, run_after, where):
    """Filter the picked members, lower-case their emails, drop duplicates."""
    return {
        prefix + "Members": {"type": "Query", "runAfter": after(run_after),
                             "inputs": {"from": "@body('PickEmails')", "where": where}},
        prefix + "List": {"type": "Select", "runAfter": after(prefix + "Members"),
                          "inputs": {"from": "@body('" + prefix + "Members')", "select": "@toLower(item()?['Email'])"}},
        # union() of a list with itself drops the duplicates.
        prefix + "Emails": {"type": "Compose", "runAfter": after(prefix + "List"),
                            "inputs": "@take(union(body('" + prefix + "List'), body('" + prefix + "List')), 20)"},
    }


def definition(t):
    site = t["site"]
    has_normal_files = "greater(length(body('NormalFiles')), 0)"
    has_hc_files = "greater(length(body('HcFiles')), 0)"
    hc_audience = "greater(length(outputs('HcEmails')), 0)"
    normal_audience = "and(greater(length(outputs('NormalEmails')), 0), " + has_normal_files + ")"

    mark_each_file = {
        "type": "Foreach", "foreach": "@body('KeyFiles')", "runAfter": {},
        "runtimeConfiguration": ONE_AT_A_TIME,
        "actions": {
            "GetRecord": sp("GetItems", {"dataset": site, "table": t["submissions_list"],
                                         "$filter": "SubmissionFileId eq '@{items('Each_file')?['SubmissionFileId']}'",
                                         "$top": 1}),
            "MarkNotified": sp_http(site, "POST",
                                    "_api/web/lists/getbytitle('CRS Submissions')/items(@{first(body('GetRecord')?['value'])?['ID']})",
                                    after("GetRecord"), "{ \"ApproverNotifiedAt\": \"@{utcNow()}\" }"),
        },
    }
    # Members of the unit's APR and APRHC groups, each tagged with its group's role.
    each_group = {
        "type": "Foreach", "foreach": "@body('GetApproverGroups')?['value']", "runAfter": after("Reset_Members"),
        "runtimeConfiguration": ONE_AT_A_TIME,
        "actions": {
            "GetMembers": sp_http(site, "GET",
                                  "_api/web/sitegroups(@{items('Each_group')?['GroupId']})/users?$select=Email", {}),
            "Each_member": {"type": "Foreach", "foreach": "@coalesce(body('GetMembers')?['value'], createArray())",
                            "runAfter": after("GetMembers"), "runtimeConfiguration": ONE_AT_A_TIME,
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
                                            "not(equals(toLower(item()?['Email']), 'gdc@sdguthrie.com')))")}},
    }
    # APRHC sees both libraries; plain APR (not also APRHC) gets only the normal files and link.
    has_approver_actions.update(distinct_emails("Hc", "PickEmails", "@equals(item()?['Role'], 'APRHC')"))
    has_approver_actions.update(distinct_emails(
        "Normal", "HcEmails",
        "@and(equals(item()?['Role'], 'APR'), not(contains(outputs('HcEmails'), toLower(item()?['Email']))))"))
    has_approver_actions.update({
        "NormalFiles": {"type": "Query", "runAfter": after("NormalEmails"),
                        "inputs": {"from": "@body('KeyFiles')", "where": "@startsWith(item()?['{Path}'], 'ApprovalDocument/')"}},
        "HcFiles": {"type": "Query", "runAfter": after("NormalFiles"),
                    "inputs": {"from": "@body('KeyFiles')", "where": "@not(startsWith(item()?['{Path}'], 'ApprovalDocument/'))"}},
        "SendToHc": {"type": "If", "runAfter": after("HcFiles"), "expression": is_true("@" + hc_audience),
                     "actions": {"Send_hc_email": send_email(t, "@join(outputs('HcEmails'), ';')", "body('KeyFiles')",
                                                             has_normal_files, has_hc_files)},
                     "else": {"actions": {}}},
        "SendToNormal": {"type": "If", "runAfter": after("SendToHc"), "expression": is_true("@" + normal_audience),
                         "actions": {"Send_normal_email": send_email(t, "@join(outputs('NormalEmails'), ';')",
                                                                     "body('NormalFiles')", "true", "false")},
                         "else": {"actions": {}}},
        # Mark only after the emails went out: a missed mark means a repeat email, never a silent one.
        "AnySent": {"type": "If", "runAfter": after("SendToNormal"),
                    "expression": is_true("@or(" + hc_audience + ", " + normal_audience + ")"),
                    "actions": {"Each_file": mark_each_file}, "else": {"actions": {}}},
    })
    has_approver = {
        "type": "If", "runAfter": after("GetApproverGroups"),
        "expression": is_true("@greater(length(coalesce(body('GetApproverGroups')?['value'], createArray())), 0)"),
        "actions": has_approver_actions,
        "else": {"actions": {}},
    }
    ready = {
        "type": "If", "runAfter": after("StillArriving"),
        "expression": is_true("@equals(length(body('StillArriving')), 0)"),
        "actions": {
            "GetApproverGroups": sp_http(site, "GET",
                                         "_api/web/lists/getbytitle('CRS Group Map')/items?$select=GroupId,Role&$filter=(Role eq 'APR' or Role eq 'APRHC') and UnitTermGuid eq '@{first(body('KeyFiles'))?['UnitTid']}'&$top=10",
                                         {}),
            "HasApprover": has_approver,
        },
        "else": {"actions": {}},
    }
    each_key = {
        "type": "Foreach", "foreach": "@union(body('KeyList'), body('KeyList'))", "runAfter": after("KeyList"),
        "runtimeConfiguration": ONE_AT_A_TIME,
        "actions": {
            "KeyFiles": {"type": "Query", "runAfter": {},
                         "inputs": {"from": "@body('Candidates')", "where": "@equals(" + SUBMISSION_KEY + ", items('Each_key'))"}},
            "StillArriving": {"type": "Query", "runAfter": after("KeyFiles"),
                              "inputs": {"from": "@outputs('AllPending')",
                                         "where": ("@and(not(empty(coalesce(item()?['SubmissionId'], ''))), "
                                                   "equals(item()?['SubmissionId'], first(body('KeyFiles'))?['SubmissionId']), "
                                                   "or(and(empty(coalesce(item()?['UnitTid'], '')), greater(ticks(item()?['Created']), ticks(addMinutes(utcNow(), -30)))), "
                                                   "greater(ticks(item()?['Modified']), ticks(addMinutes(utcNow(), -3)))))")}},
            "Ready": ready,
        },
    }
    pending_filter = "OData__ModerationStatus eq 2 and Created ge '" + THREE_DAYS_AGO + "'"
    actions = {
        "Init_Members": {"type": "InitializeVariable", "runAfter": {},
                         "inputs": {"variables": [{"name": "Members", "type": "array", "value": []}]}},
        "GetPending": sp("GetItems", {"dataset": site, "table": t["approval_list"], "$filter": pending_filter,
                                      "$top": 500}, after("Init_Members")),
        "GetPendingHC": sp("GetItems", {"dataset": site, "table": t["hc_approval_list"], "$filter": pending_filter,
                                        "$top": 500}, after("GetPending")),
        # Normal + HC files together.
        "AllPending": {"type": "Compose", "runAfter": after("GetPendingHC"),
                       "inputs": "@union(coalesce(body('GetPending')?['value'], createArray()), coalesce(body('GetPendingHC')?['value'], createArray()))"},
        "Tagged": {"type": "Query", "runAfter": after("AllPending"),
                   "inputs": {"from": "@outputs('AllPending')",
                              "where": ("@and(equals(item()?['{IsFolder}'], false), not(equals(item()?['BulkImport'], true)), "
                                        "equals(item()?['{ModerationStatus}'], 'Pending'), not(empty(coalesce(item()?['UnitTid'], ''))), "
                                        "empty(coalesce(item()?['NextReminderAt'], '')), not(empty(coalesce(item()?['SubmissionFileId'], ''))))")}},
        "GetUnnotified": sp("GetItems", {"dataset": site, "table": t["submissions_list"],
                                         "$filter": "ApproverNotifiedAt eq null and UploadedAt ge '" + THREE_DAYS_AGO + "'",
                                         "$top": 500}, after("Tagged")),
        "SelectIds": {"type": "Select", "runAfter": after("GetUnnotified"),
                      "inputs": {"from": "@body('GetUnnotified')?['value']", "select": "@item()?['SubmissionFileId']"}},
        "UnnotifiedIds": {"type": "Compose", "runAfter": after("SelectIds"),
                          "inputs": "@concat('|', join(body('SelectIds'), '|'), '|')"},
        "Candidates": {"type": "Query", "runAfter": after("UnnotifiedIds"),
                       "inputs": {"from": "@body('Tagged')",
                                  "where": "@contains(outputs('UnnotifiedIds'), concat('|', item()?['SubmissionFileId'], '|'))"}},
        "KeyList": {"type": "Select", "runAfter": after("Candidates"),
                    "inputs": {"from": "@body('Candidates')", "select": "@" + SUBMISSION_KEY}},
        "Each_key": each_key,
    }
    recurrence = {"frequency": "Minute", "interval": 5, "timeZone": "Singapore Standard Time"}
    return {
        "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {"$connections": {"defaultValue": {}, "type": "Object"},
                       "$authentication": {"defaultValue": {}, "type": "SecureObject"}},
        "triggers": {"Recurrence": {"recurrence": recurrence, "evaluatedRecurrence": recurrence, "type": "Recurrence"}},
        "actions": actions,
    }


def build(t, display_name, definition_fn, out):
    """Writes an import zip: the base export's package with a new flow id, name and definition."""
    new_id = str(uuid.uuid4())
    with zipfile.ZipFile(t["base"]) as zin:
        names = zin.namelist()
        old_id = [n for n in names if n.endswith("definition.json")][0].split("/")[2]
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                text = zin.read(n).decode("utf-8-sig")
                if n.endswith("definition.json"):
                    j = json.loads(text)
                    old_def = j["properties"]["definition"]
                    new_def = definition_fn(t)
                    new_def["parameters"] = old_def.get("parameters", new_def["parameters"])
                    j["properties"]["definition"] = new_def
                    j["properties"]["displayName"] = display_name
                    j["name"] = new_id
                    text = json.dumps(j, ensure_ascii=False)
                elif n == "manifest.json":
                    m = json.loads(text)
                    m["details"]["displayName"] = display_name
                    res = m["resources"].pop(old_id)
                    res["details"]["displayName"] = display_name
                    res["suggestedCreationType"] = "New"
                    m["resources"][new_id] = res
                    text = json.dumps(m, ensure_ascii=False)
                text = text.replace(old_id, new_id)
                zout.writestr(n.replace(old_id, new_id), text.encode("utf-8"))
    print("wrote", out, "flow id", new_id)


if __name__ == "__main__":
    target = TARGETS[sys.argv[1]]
    build(target, DISPLAY_NAME, definition, target["out"])
