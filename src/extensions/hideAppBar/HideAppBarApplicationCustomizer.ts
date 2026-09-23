import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

export default class HideAppBarApplicationCustomizer extends BaseApplicationCustomizer<Record<string, never>> {
  private _styleEl: HTMLStyleElement | null = null;
  private _observer: MutationObserver | null = null;

  public onInit(): Promise<void> {
    this._styleEl = document.createElement("style");
    this._styleEl.textContent = `#m365-app-bar { display: none !important; }`;
    document.head.appendChild(this._styleEl);

    this._observer = new MutationObserver(() => this._onDomChange());
    this._observer.observe(document.body, { childList: true, subtree: true });

    return Promise.resolve();
  }

  /**
   * Every library this system files into, under every created AND retitled name.
   *
   * Mirrors LIBRARY_CANDIDATES / HC_APPROVAL_CANDIDATES / DOCUMENTS_CANDIDATES /
   * HC_DOCUMENTS_CANDIDATES in shared/naming.ts. Compared here as literals rather than imported
   * because an application customizer loads on EVERY page of the site, and pulling in naming.ts
   * would drag its resolution cache into every page load to answer a question a string comparison
   * answers.
   *
   * ⚠⚠ THIS LIST WENT STALE AND SILENTLY DISABLED THE WHOLE CUSTOMIZER ON "Approval for Document"
   * (found live 2026-09-23) — it had only "approval document" (no "for"), which is a DIFFERENT
   * string from the actual live-renamed title. `_isCrsLibrary()` answered false, so NOTHING here
   * ran on that page: not the `+ New Folder` injection, not the upload-menu hiding, not the
   * Approve/Reject hiding. The client renames these libraries as a matter of routine (naming.ts's
   * own comment: twice in one afternoon on 2026-08-27, again the next day) — this hand-copied
   * mirror WILL drift again the next time that happens, and it will fail exactly this silently.
   * Re-sync it against naming.ts's four candidate arrays whenever CLAUDE.md records a new rename.
   */
  private static readonly CRS_LIBRARIES = [
    "approval for document",
    "approval document",
    "approvaldocument",
    "staging",
    "approval for highly confidential document",
    "hc approval document",
    "hc approval documents",
    "hcapprovaldocument",
    "highly confidential approval document",
    "documents",
    "restricted & confidential document",
    "restricted and confidential document",
    "shared documents",
    "hc documents",
    "hc document",
    "hcdocuments",
    "highly confidential document",
    "highly confidential documents",
  ];

  /**
   * True on a library this system owns.
   *
   * Matched on the list TITLE, not on the URL. The previous test was
   * `pathname.includes("/staging")`, which silently stopped matching anything the day the library
   * was recreated as `Approval Document` at `/ApprovalDocument` — so the upload items came back on
   * the one library they had been hidden from, and nothing reported it. A title is also the only
   * thing that identifies `Documents`, whose URL segment is `Shared Documents` (gotcha #12).
   *
   * Absent list context means this is not a library page — a site page, the home page — so the
   * answer is false and no menu is touched.
   */
  private _isCrsLibrary(): boolean {
    const title = (this.context.pageContext.list?.title ?? "").trim().toLowerCase();
    if (title.length === 0) return false;
    return HideAppBarApplicationCustomizer.CRS_LIBRARIES.indexOf(title) !== -1;
  }

  private _onDomChange(): void {
    if (!this._isCrsLibrary()) return;
    this._injectNewFolderButton();
    // Every upload is meant to arrive through the upload form, in every library this system owns —
    // that is what gives a document its metadata, its folder routing and its approval trail. A file
    // dragged straight into a library has none of them, and looks completely normal in the view.
    this._hideUploadMenuItems();
    // The native command bypasses this project's own approval guards entirely — it flips
    // OData__ModerationStatus directly and never runs ApprovalDocument.tsx's destination-folder or
    // name-clash checks (see CLAUDE.md, "THE NATIVE APPROVE COMMAND BYPASSES BOTH GUARDS"). A web
    // part cannot intercept a native SharePoint control, only hide the invitation to use it — the
    // real safety is still the copy-with-a-new-name behaviour on Auto-route's own Copy file step.
    this._hideApproveRejectCommand();
  }

  private _injectNewFolderButton(): void {
    if (document.getElementById("sdg-new-folder-btn")) return;

    const nativeBtn = this._findCreateOrUploadButton();
    if (!nativeBtn || !nativeBtn.parentElement) return;

    // Move off-screen but keep in DOM so we can .click() it to open the dropdown
    nativeBtn.style.cssText = "position:fixed;left:-9999px;top:-9999px;";

    const btn = document.createElement("button");
    btn.id = "sdg-new-folder-btn";
    btn.type = "button";
    btn.textContent = "+ New Folder";
    btn.style.cssText = [
      "display:inline-flex",
      "align-items:center",
      "height:32px",
      "padding:0 16px",
      "margin:4px 8px 4px 0",
      "background:#107c10",
      "color:#fff",
      "border:none",
      "border-radius:4px",
      "font-size:14px",
      "font-family:inherit",
      "font-weight:600",
      "cursor:pointer",
      "white-space:nowrap",
    ].join(";");

    btn.addEventListener("click", () => this._triggerNativeFolderDialog());
    nativeBtn.parentElement.insertBefore(btn, nativeBtn);
  }

  private _triggerNativeFolderDialog(): void {
    // Open the hidden native dropdown
    const nativeBtn = this._findCreateOrUploadButton();
    if (!nativeBtn) return;
    nativeBtn.click();

    // Poll for the "Folder" item in the dropdown and click it
    const poll = (remaining: number): void => {
      if (remaining <= 0) return;

      // Try known automationid first, then fall back to text match
      const byId = document.querySelector<HTMLElement>('[data-automationid="newFolder"]');
      if (byId) { byId.click(); return; }

      const byText = Array.from(
        document.querySelectorAll<HTMLElement>(".ms-ContextualMenu-link, [role='menuitem']")
      ).find((el) => el.textContent?.trim().toLowerCase() === "folder");
      if (byText) { byText.click(); return; }

      setTimeout(() => poll(remaining - 1), 100);
    };
    setTimeout(() => poll(20), 150);
  }

  private _findCreateOrUploadButton(): HTMLElement | null {
    const el = document.querySelector<HTMLElement>('[data-automationid="newCommand"]');
    if (el) return el;
    return (
      Array.from(document.querySelectorAll<HTMLElement>("button")).find((b) =>
        b.textContent?.toLowerCase().includes("create or upload")
      ) ?? null
    );
  }

  private _hideUploadMenuItems(): void {
    const toHide = [
      '[data-automationid="uploadFile"]',
      '[data-automationid="uploadFolder"]',
    ];
    toHide.forEach((sel) => {
      document.querySelectorAll<HTMLElement>(sel).forEach((el) => {
        const li = el.closest("li");
        if (li) (li as HTMLElement).style.display = "none";
        else el.style.display = "none";
      });
    });

    const textsToHide = ["link", "edit new menu", "add template"];
    document.querySelectorAll<HTMLElement>(".ms-ContextualMenu-link").forEach((el) => {
      const text = el.textContent?.trim().toLowerCase() || "";
      if (textsToHide.some((t) => text === t)) {
        const li = el.closest("li");
        if (li) (li as HTMLElement).style.display = "none";
      }
    });
  }

  /**
   * Hides the native Approve/Reject command everywhere SharePoint renders it.
   *
   * PRIMARY match: `[data-automationid="approveReject"]` — confirmed live via DevTools 2026-09-23,
   * on the "Integrate" flyout menu entry (this command bar renders it nested there, not as its own
   * top-level button). This is the reliable route.
   *
   * FALLBACK match: the visible label text, but only on `.ms-ContextualMenu-itemText` — the LEAF
   * span holding just the clean label — never the containing button/link/menuitem. The first
   * version of this matched the whole menu item's `textContent`, which also picks up the icon's own
   * glyph: Fluent UI icons render their character as an actual Unicode code point inside the `<i>`
   * element, so the button's full text was never an exact match for the plain string
   * "approve/reject" — it silently matched nothing. This fallback exists only for a render context
   * that might lack the automationid, such as the view switcher's own "Approve/reject Items" view
   * entry, which was never confirmed live.
   */
  private _hideApproveRejectCommand(): void {
    document
      .querySelectorAll<HTMLElement>('[data-automationid="approveReject"]')
      .forEach((el) => this._hideMenuItem(el));

    /* ⚠⚠ "Approve or reject" IS NOT THIS COMMAND — DO NOT ADD IT BACK. It was added here briefly
       on 2026-09-23 from a live screenshot, and the client caught it within minutes: that exact
       wording is THIS PROJECT'S OWN Bulk Approve command (the `bulkApprove` ListViewCommandSet, see
       CLAUDE.md "BULK APPROVE FROM THE LIBRARY COMMAND BAR"), which runs through the app's own
       approval guards and must stay reachable. Hiding it here would silently disable a feature this
       project built on purpose while looking, from this file alone, like it was hiding the native
       one. Whether the NATIVE command also has its own separate entry in that same row-level "..."
       overflow menu — and under what exact wording — has NOT been confirmed; only the "Integrate"
       flyout's "Approve/Reject" has been verified live so far. Confirm before adding a third string
       here, rather than guessing one back in. */
    const textsToHide = ["approve/reject", "approve/reject items"];
    document
      .querySelectorAll<HTMLElement>(".ms-ContextualMenu-itemText")
      .forEach((el) => {
        const text = (el.textContent ?? "").trim().toLowerCase();
        if (textsToHide.some((t) => text === t)) this._hideMenuItem(el);
      });
  }

  /** Hides the enclosing `<li>` when present (a clean row removal), else the element itself. */
  private _hideMenuItem(el: HTMLElement): void {
    const li = el.closest("li");
    if (li) (li as HTMLElement).style.display = "none";
    else el.style.display = "none";
  }

  protected onDispose(): void {
    if (this._observer) { this._observer.disconnect(); this._observer = null; }
    if (this._styleEl) { this._styleEl.remove(); this._styleEl = null; }
    const btn = document.getElementById("sdg-new-folder-btn");
    if (btn) btn.remove();
  }
}
