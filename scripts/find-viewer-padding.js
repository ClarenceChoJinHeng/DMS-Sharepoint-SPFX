// Read-only. Paste into DevTools Console on Document-Viewer.aspx (narrow the window first).
// Walks up from the web part and lists every ancestor that adds left/right space.
(function findViewerPadding() {
  var start = document.querySelector(".crs-ms-wrap");
  if (!start) { console.warn("No .crs-ms-wrap on this page. Open a document first."); return; }
  var rows = [];
  var node = start;
  while (node && node !== document.documentElement) {
    var cs = getComputedStyle(node);
    var pl = parseFloat(cs.paddingLeft) || 0, pr = parseFloat(cs.paddingRight) || 0;
    var ml = parseFloat(cs.marginLeft) || 0, mr = parseFloat(cs.marginRight) || 0;
    if (pl || pr || ml || mr) {
      rows.push({
        tag: node.tagName.toLowerCase(),
        id: node.id || "",
        className: String(node.className || "").slice(0, 80),
        dataAutomationId: node.getAttribute("data-automation-id") || "",
        padding: pl + " / " + pr,
        margin: ml + " / " + mr,
        width: Math.round(node.getBoundingClientRect().width),
      });
    }
    node = node.parentElement;
  }
  console.log("Viewport width:", window.innerWidth);
  console.table(rows);
  window.__result = rows;
})();
