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

  private _onDomChange(): void {
    this._injectNewFolderButton();
    // Upload item hiding is Staging-only
    if (window.location.pathname.toLowerCase().includes("/staging")) {
      this._hideUploadMenuItems();
    }
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

  protected onDispose(): void {
    if (this._observer) { this._observer.disconnect(); this._observer = null; }
    if (this._styleEl) { this._styleEl.remove(); this._styleEl = null; }
    const btn = document.getElementById("sdg-new-folder-btn");
    if (btn) btn.remove();
  }
}
