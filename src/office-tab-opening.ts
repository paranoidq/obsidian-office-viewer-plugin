import { parseLinktext, TFile, type App } from "obsidian";
import { reportNonFatalError } from "./report-error";

export function registerOfficeTabOpening(
  app: App,
  canPreview: (file: TFile) => boolean,
): () => void {
  const { workspace, metadataCache, vault } = app;
  const openLinkText = workspace.openLinkText;
  let disposed = false;
  const openOfficeLink: typeof openLinkText = function (
    this: typeof workspace, linktext, sourcePath, newLeaf, openViewState,
  ) {
    if (!disposed && !newLeaf) {
      const { path } = parseLinktext(linktext);
      const file = metadataCache.getFirstLinkpathDest(path, sourcePath);
      if (file !== null && canPreview(file)) newLeaf = "tab";
    }
    return openLinkText.call(this, linktext, sourcePath, newLeaf, openViewState);
  };
  workspace.openLinkText = openOfficeLink;

  const listeners = new Map<Window, (event: MouseEvent) => void>();
  const attachExplorerListener = (window: Window) => {
    if (listeners.has(window)) return;
    const onClick = (event: MouseEvent) => {
      if (event.defaultPrevented || event.button !== 0
        || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
      const ElementConstructor = window.document.defaultView?.Element;
      const target = event.target;
      if (ElementConstructor === undefined || !(target instanceof ElementConstructor)
        || target.closest('[contenteditable="true"], input')) return;
      const path = target.closest(".nav-file-title[data-path]")?.getAttribute("data-path");
      if (!path) return;
      const file = vault.getAbstractFileByPath(path);
      if (!(file instanceof TFile) || !canPreview(file)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      const leaf = workspace.getLeaf("tab");
      void leaf.openFile(file, { active: true }).catch((error: unknown) => {
        reportNonFatalError("Failed to open the Office preview tab", error);
      });
    };
    listeners.set(window, onClick);
    window.addEventListener("click", onClick, true);
  };
  const detachExplorerListener = (window: Window) => {
    const onClick = listeners.get(window);
    if (onClick === undefined) return;
    window.removeEventListener("click", onClick, true);
    listeners.delete(window);
  };
  attachExplorerListener(window);
  workspace.iterateAllLeaves((leaf) => {
    const window = leaf.view.containerEl.ownerDocument.defaultView;
    if (window !== null) attachExplorerListener(window);
  });
  const onWindowOpen = workspace.on("window-open", (_, window) => attachExplorerListener(window));
  const onWindowClose = workspace.on("window-close", (_, window) => detachExplorerListener(window));
  return () => {
    disposed = true;
    if (workspace.openLinkText === openOfficeLink) workspace.openLinkText = openLinkText;
    workspace.offref(onWindowOpen);
    workspace.offref(onWindowClose);
    for (const window of listeners.keys()) detachExplorerListener(window);
  };
}
