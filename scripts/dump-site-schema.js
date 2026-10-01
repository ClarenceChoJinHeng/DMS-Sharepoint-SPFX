// Dumps the site's full schema (lists, libraries, columns, views, content types, permissions, term
// sets, config rows, pages → web parts, extensions) into one JSON file for the design spec.
// READ ONLY: every request is a GET. Run: open the CRS site, F12 → Console, paste, Enter.
// A file crs-schema-<date>.json downloads. A failed read is recorded in `errors`, never as "empty".
(async () => {
  // ---- CONFIG ----
  const SITE_OVERRIDE = "";      // e.g. "https://sdguthrie.sharepoint.com/sites/CRS"
  const MAX_TERM_DEPTH = 8;
  const MAX_TERMS = 5000;
  // CRS/DMS lists whose ROWS are dumped (setup data, no personal data). Other lists: counts only.
  const ROW_DUMP_SUFFIXES = ["Config", "Term Abbreviation", "Group Map", "Folder Map"];

  // ---- HELPERS ----
  function resolveWeb() {
    if (SITE_OVERRIDE) return SITE_OVERRIDE;
    try {
      if (window._spPageContextInfo) return window._spPageContextInfo.webAbsoluteUrl;
    } catch (e) {
      // not on a classic page; fall through
    }
    const match = location.pathname.match(/^\/(sites|teams)\/[^/]+/);
    return location.origin + (match ? match[0] : "");
  }
  const web = resolveWeb();
  const bust = () => "_=" + Date.now() + Math.random().toString(36).slice(2);
  const errors = [];
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  async function getJson(path, accept) {
    const url = web + path + (path.includes("?") ? "&" : "?") + bust();
    try {
      const res = await fetch(url, {
        headers: { Accept: accept || "application/json;odata=nometadata", "Cache-Control": "no-cache" },
      });
      const type = res.headers.get("content-type") || "";
      if (!res.ok || !type.includes("json")) {
        errors.push({ path, status: res.status });
        return { ok: false, status: res.status };
      }
      return { ok: true, data: await res.json() };
    } catch (e) {
      errors.push({ path, status: "network", message: String(e) });
      return { ok: false, status: "network" };
    }
  }

  // Reads every page. Returns undefined if any page fails (unknown, never "empty").
  async function getAll(path, accept) {
    const rows = [];
    let next = path;
    while (next) {
      const r = await getJson(next, accept);
      if (!r.ok) return undefined;
      rows.push(...(r.data.value || []));
      const link = r.data["odata.nextLink"] || r.data["@odata.nextLink"] || "";
      next = link ? link.replace(web, "") : "";
    }
    return rows;
  }

  // /sitegroups has no nextLink: $top alone truncates. Page by Id instead.
  async function getAllGroups() {
    let groups = [];
    let lastId = 0;
    for (;;) {
      const page = await getAll("/_api/web/sitegroups?$select=Id,Title,OwnerTitle,Description"
        + "&$orderby=Id&$filter=Id gt " + lastId + "&$top=500");
      if (page === undefined) return undefined;
      groups = groups.concat(page);
      if (page.length < 500) return groups;
      lastId = page[page.length - 1].Id;
    }
  }

  function pick(obj, keys) {
    const out = {};
    for (const k of keys) if (obj[k] !== undefined && obj[k] !== null && obj[k] !== "") out[k] = obj[k];
    return out;
  }

  const FIELD_KEYS = ["Title", "InternalName", "StaticName", "Id", "TypeAsString", "TypeDisplayName",
    "Required", "Hidden", "ReadOnlyField", "FromBaseType", "Sealed", "Indexed", "EnforceUniqueValues",
    "DefaultValue", "Description", "Group", "MaxLength", "Choices", "FillInChoice", "AllowMultipleValues",
    "LookupList", "LookupField", "DisplayFormat", "Formula", "OutputType", "SelectionMode",
    "SelectionGroup", "RichText", "NumberOfLines", "AppendOnly", "TermSetId", "SspId", "AnchorId",
    "IsKeyword", "CreateValuesInEditForm", "Open", "TextField", "ValidationFormula", "ValidationMessage"];

  const LIST_KEYS = ["Id", "Title", "Description", "BaseTemplate", "BaseType", "Hidden", "ItemCount",
    "EnableModeration", "DraftVersionVisibility", "EnableVersioning", "EnableMinorVersions",
    "MajorVersionLimit", "ForceCheckout", "EnableAttachments", "ContentTypesEnabled",
    "HasUniqueRoleAssignments", "NoCrawl", "Created", "LastItemModifiedDate"];

  const isCrsList = (title) => /^(CRS|DMS) /.test(title);
  const wantsRows = (title) => isCrsList(title) && ROW_DUMP_SUFFIXES.some((s) => title.endsWith(" " + s));

  async function readRoleAssignments(base) {
    const rows = await getAll(base + "/roleassignments?$expand=Member,RoleDefinitionBindings"
      + "&$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name");
    if (rows === undefined) return undefined;
    return rows.map((r) => ({
      principalId: r.PrincipalId,
      member: r.Member ? r.Member.Title : "",
      principalType: r.Member ? r.Member.PrincipalType : "",
      roles: (r.RoleDefinitionBindings || []).map((d) => d.Name),
    }));
  }

  // One rung per optional part: views with their columns first, then without.
  async function readViews(base) {
    const select = "$select=Title,ServerRelativeUrl,DefaultView,Hidden,PersonalView,RowLimit,ViewQuery";
    let rows = await getAll(base + "/views?" + select + ",ViewFields/Items&$expand=ViewFields");
    if (rows === undefined) rows = await getAll(base + "/views?" + select);
    if (rows === undefined) return undefined;
    return rows.filter((v) => !v.Hidden).map((v) => ({
      title: v.Title, url: v.ServerRelativeUrl, isDefault: v.DefaultView, rowLimit: v.RowLimit,
      query: v.ViewQuery, columns: v.ViewFields ? v.ViewFields.Items : undefined,
    }));
  }

  async function readList(list) {
    const base = "/_api/web/lists(guid'" + list.Id + "')";
    const fieldsRaw = await getAll(base + "/fields");
    const fields = fieldsRaw === undefined ? undefined : fieldsRaw
      .filter((f) => !f.Hidden || !f.FromBaseType)
      .map((f) => pick(f, FIELD_KEYS));
    const out = Object.assign(pick(list, LIST_KEYS), {
      url: list.RootFolder ? list.RootFolder.ServerRelativeUrl : "",
      fields,
      contentTypes: await getAll(base + "/contenttypes?$select=Name,StringId,Description"),
      views: await readViews(base),
      roleAssignments: list.HasUniqueRoleAssignments ? await readRoleAssignments(base) : "inherits",
      customActions: await getAll(base + "/usercustomactions?$select=Title,Location,ClientSideComponentId"),
    });
    if (wantsRows(list.Title) && fields !== undefined) {
      const custom = fields.filter((f) => !f.FromBaseType).map((f) => f.InternalName);
      const items = await getAll(base + "/items?$top=2000");
      out.rows = items === undefined ? undefined : items.map((it) => pick(it, ["Id", "Title"].concat(custom)));
    }
    return out;
  }

  function collectTermSetIds(lists) {
    const ids = new Set();
    for (const l of lists) {
      for (const f of l.fields || []) {
        if (f.TermSetId && GUID.test(f.TermSetId) && !/^0{8}-/.test(f.TermSetId)) ids.add(f.TermSetId.toLowerCase());
      }
      for (const r of l.rows || []) {
        for (const k of Object.keys(r)) {
          if (/TermSet/i.test(k) && GUID.test(String(r[k]))) ids.add(String(r[k]).toLowerCase());
        }
      }
    }
    return Array.from(ids);
  }

  // Term store (v2.1 needs plain application/json). Walks each set's tree.
  async function readTermSet(setId) {
    const accept = "application/json";
    const head = await getJson("/_api/v2.1/termStore/sets/" + setId, accept);
    const set = { id: setId, ok: head.ok, status: head.status };
    if (!head.ok) return set;
    set.names = (head.data.localizedNames || []).map((n) => n.name);
    set.description = head.data.description || "";
    set.terms = [];
    set.complete = true;
    const walk = async (parentId, depth) => {
      if (depth > MAX_TERM_DEPTH || set.terms.length >= MAX_TERMS) { set.complete = false; return; }
      const path = "/_api/v2.1/termStore/sets/" + setId
        + (parentId ? "/terms/" + parentId + "/children" : "/children") + "?$select=id,labels";
      const kids = await getAll(path, accept);
      if (kids === undefined) { set.complete = false; return; }
      for (const k of kids) {
        set.terms.push({ id: k.id, parentId, depth, labels: (k.labels || []).map((l) => l.name) });
        if (set.terms.length % 25 === 0) console.log("  terms read:", set.terms.length);
        await walk(k.id, depth + 1);
      }
    };
    await walk("", 1);
    return set;
  }

  function decodeHtml(text) {
    const box = document.createElement("textarea");
    box.innerHTML = text;
    return box.value;
  }

  function webPartsOnPage(canvas) {
    const parts = [];
    const re = /data-sp-webpartdata="([^"]*)"/g;
    let m;
    while ((m = re.exec(canvas || ""))) {
      try {
        const data = JSON.parse(decodeHtml(m[1]));
        parts.push({ id: data.id, title: data.title });
      } catch (e) {
        parts.push({ id: "unreadable" });
      }
    }
    return parts;
  }

  function download(name, text) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([text], { type: "application/json" }));
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
  }

  // ---- MAIN ----
  console.log("Site:", web, "| READ ONLY");
  const result = { site: web, dumpedAt: new Date().toISOString(), errors };

  const webInfo = await getJson("/_api/web?$select=Id,Title,Url,ServerRelativeUrl,WebTemplate,Configuration,"
    + "Language,UIVersion,Created,HasUniqueRoleAssignments");
  result.web = webInfo.ok ? webInfo.data : undefined;
  const siteInfo = await getJson("/_api/site?$select=Id,Url");
  result.siteCollection = siteInfo.ok ? siteInfo.data : undefined;

  result.permissionLevels = await getAll("/_api/web/roledefinitions?$select=Id,Name,Description,Hidden,"
    + "RoleTypeKind,Order");
  result.siteRoleAssignments = await readRoleAssignments("/_api/web");
  const groups = await getAllGroups();
  result.siteGroups = groups === undefined ? undefined : { count: groups.length, groups };
  result.webCustomActions = await getAll("/_api/web/usercustomactions?$select=Title,Location,"
    + "ClientSideComponentId,RegistrationId,RegistrationType");
  result.siteCustomActions = await getAll("/_api/site/usercustomactions?$select=Title,Location,"
    + "ClientSideComponentId,RegistrationId,RegistrationType");

  const lists = await getAll("/_api/web/lists?$expand=RootFolder&$select=" + LIST_KEYS.join(",")
    + ",RootFolder/ServerRelativeUrl");
  if (lists === undefined) {
    console.warn("Could not read the site's lists. Nothing concluded.");
    console.table(errors);
    window.__result = result;
    return;
  }
  const chosen = lists.filter((l) => !l.Hidden || isCrsList(l.Title));
  result.lists = [];
  for (let i = 0; i < chosen.length; i++) {
    console.log("List " + (i + 1) + "/" + chosen.length + ": " + chosen[i].Title);
    result.lists.push(await readList(chosen[i]));
  }

  result.termSets = [];
  for (const id of collectTermSetIds(result.lists)) {
    console.log("Term set", id);
    result.termSets.push(await readTermSet(id));
  }

  const pagesList = lists.find((l) => l.BaseTemplate === 119);
  if (pagesList) {
    const pages = await getAll("/_api/web/lists(guid'" + pagesList.Id + "')/items"
      + "?$select=Id,Title,FileLeafRef,CanvasContent1&$top=500");
    result.pages = pages === undefined ? undefined : pages.map((p) => ({
      file: p.FileLeafRef, title: p.Title, webParts: webPartsOnPage(p.CanvasContent1),
    }));
  }

  window.__result = result;
  console.table(result.lists.map((l) => ({
    title: l.Title, url: l.url, items: l.ItemCount,
    columns: l.fields ? l.fields.length : "READ FAILED",
    approval: l.EnableModeration, uniquePerms: l.HasUniqueRoleAssignments,
  })));
  console.log("Term sets: " + result.termSets.map((s) => (s.names || [s.id]).join("/") + " → "
    + (s.ok ? s.terms.length + " terms" + (s.complete ? "" : " (INCOMPLETE)") : "HTTP " + s.status)).join(" | "));
  console.log("Pages:", result.pages ? result.pages.length : "not read",
    "| Groups:", result.siteGroups ? result.siteGroups.count : "not read",
    "| Failed reads:", errors.length);
  if (errors.length) console.table(errors);

  const name = "crs-schema-" + new Date().toISOString().slice(0, 10) + ".json";
  download(name, JSON.stringify(result, null, 1));
  console.log("Saved " + name + " (also on window.__result). Send me the file and this summary.");
})();
