import {
  FileView,
  Scope,
  setIcon,
  TFile,
  type WorkspaceLeaf,
} from "obsidian";
import {
  createOfficeViewerErrorSurface,
  createOfficeViewerToolbar,
  decorateOfficeViewerIconButton,
} from "../office-viewer-chrome";
import { registerOfficeEmbedTextCopy } from "../office-embed-text-copy";
import { registerDocxAttachmentOpening } from "./docx-attachment-controls";
import type { DocxAttachmentFile } from "./docx-attachments";
import type {
  DocxMessageTranslator,
} from "./docx-messages";
import { DocxOpenError } from "./docx-open-error";
import { DocxSearchPanel } from "./docx-search-panel";
import {
  inspectDocxPackage,
  type DocxSemanticModel,
  type DocxSemanticParagraph,
} from "./docx-semantic-model";
import type {
  DocxRenderMode,
  DocxRendererAdapter,
  DocxRendererSession,
} from "./renderer/docx-renderer-adapter";
import { DocxRenderSource } from "./docx-render-source";
import { createDocxModeButton } from "./docx-mode-button";
import { createDocxOutline, scrollToDocxHeading } from "./docx-outline";
import { createDocxZoomControls } from "./docx-zoom-controls";

export const DOCX_VIEW_TYPE = "docx-viewer";

export interface DocxViewDependencies {
  readonly renderer: DocxRendererAdapter;
  readonly messages: DocxMessageTranslator;
  readonly readBinary?: (file: TFile, signal: AbortSignal) => Promise<ArrayBuffer>;
  readonly openExternalUrl?: (url: string) => Promise<void>;
  readonly openAttachment?: (file: DocxAttachmentFile) => Promise<void>;
  readonly openInDefaultApplication?: (file: TFile) => Promise<void>;
}

export interface DocxViewPerformanceDiagnostics {
  readonly candidate: string | null;
  readonly sourceBytes: number | null;
  readonly paragraphCount: number | null;
  readonly firstReadableMs: number | null;
  readonly searchReadyMs: number | null;
}

function button(label: string): HTMLButtonElement {
  const element = createEl("button");
  element.type = "button";
  element.textContent = label;
  return element;
}

function safeExternalProtocol(url: string): boolean {
  try {
    const protocol = new URL(url).protocol;
    return (
      protocol === "https:" ||
      protocol === "http:" ||
      protocol === "mailto:"
    );
  } catch {
    return false;
  }
}

export class DocxFileView extends FileView {
  private readonly root = createDiv();
  private readonly toolbar = createOfficeViewerToolbar({
    format: "DOCX",
    extraClassName: "office-viewer-docx-toolbar",
  });
  private readonly searchButton: HTMLButtonElement;
  private readonly externalButton: HTMLButtonElement;
  private readonly modeToggle: ReturnType<typeof createDocxModeButton>;
  private readonly zoomControls: ReturnType<typeof createDocxZoomControls>;
  private readonly outline: ReturnType<typeof createDocxOutline>;
  private readonly notices = createDiv();
  private readonly noticesIcon = createSpan();
  private readonly noticesText = createSpan();
  private readonly actionStatus = createDiv();
  private readonly main = createDiv();
  private readonly searchRail = createDiv();
  private readonly readingBody = createDiv();
  private readonly searchPanel: DocxSearchPanel;
  private currentFile: TFile | null = null;
  private model: DocxSemanticModel | null = null;
  private renderSource: DocxRenderSource | null = null;
  private rendererSession: DocxRendererSession | null = null;
  private activeParagraph: DocxSemanticParagraph | null = null;
  private activeParagraphElement: HTMLElement | null = null;
  private readonly unregisterTextCopy: () => void;
  private readonly unregisterAttachmentOpening: () => void;
  private abortController: AbortController | null = null;
  private disposed = false;
  private diagnostics: DocxViewPerformanceDiagnostics = {
    candidate: null,
    sourceBytes: null,
    paragraphCount: null,
    firstReadableMs: null,
    searchReadyMs: null,
  };

  constructor(
    leaf: WorkspaceLeaf,
    private readonly dependencies: DocxViewDependencies,
    private readonly onDisposed: () => void = () => {},
  ) {
    super(leaf);
    this.root.className = "office-viewer-docx-shell";
    this.root.dataset.state = "idle";
    this.searchButton = button("");
    decorateOfficeViewerIconButton(this.searchButton, "lucide-search");
    this.searchButton.title = dependencies.messages.text("searchOpen");
    this.searchButton.setAttribute("data-action", "open-docx-search");
    this.searchButton.setAttribute(
      "aria-label",
      dependencies.messages.text("searchOpen"),
    );
    this.searchButton.setAttribute("aria-pressed", "false");
    this.externalButton = button("");
    decorateOfficeViewerIconButton(this.externalButton, "lucide-external-link");
    this.externalButton.title = dependencies.messages.text("openDefault");
    this.externalButton.setAttribute("aria-label", this.externalButton.title);
    this.externalButton.setAttribute("data-action", "open-externally");
    this.modeToggle = createDocxModeButton(
      this.root.ownerDocument, dependencies.messages, (mode) => {
        void this.renderDocument(mode);
      },
    );
    this.zoomControls = createDocxZoomControls(this.root.ownerDocument, dependencies.messages);
    this.notices.className = "office-viewer-docx-notices";
    this.notices.setAttribute("role", "status");
    this.notices.hidden = true;
    this.noticesIcon.className = "office-viewer-docx-notices__icon";
    this.noticesIcon.setAttribute("aria-hidden", "true");
    setIcon(this.noticesIcon, "lucide-info");
    this.noticesText.className = "office-viewer-docx-notices__text";
    this.notices.append(this.noticesIcon, this.noticesText);
    this.actionStatus.className = "office-viewer-docx-action-status";
    this.actionStatus.setAttribute("aria-live", "polite");
    this.main.className = "office-viewer-docx-main";
    this.searchRail.className = "office-viewer-docx-search-rail";
    this.readingBody.className = "office-viewer-docx-reading-body";
    this.readingBody.setAttribute("role", "document");
    this.unregisterTextCopy = registerOfficeEmbedTextCopy(this.readingBody);
    this.unregisterAttachmentOpening = registerDocxAttachmentOpening(
      this.readingBody, () => this.renderSource, dependencies.openAttachment,
      dependencies.messages, this.actionStatus,
    );

    this.searchPanel = new DocxSearchPanel(this.root, this.searchRail, {
      messages: dependencies.messages,
      getModel: () => this.model,
      currentParagraphOrdinal: () => this.activeParagraph?.ordinal ?? null,
      onNavigate: (paragraphOrdinal) => {
        this.activateOrdinal(paragraphOrdinal);
      },
      onDismiss: () => this.closeSearch(),
    });
    this.outline = createDocxOutline(this.root.ownerDocument, dependencies.messages, (ordinal) => {
      const element = this.rendererSession?.revealParagraph(ordinal);
      if (element == null) return false;
      this.setActiveParagraph(null);
      scrollToDocxHeading(this.readingBody, element);
      return true;
    }, () => {
      this.searchPanel.close();
      this.updateSearchButton();
    });

    this.toolbar.primary.append(this.outline.button, this.searchButton);
    this.toolbar.secondary.append(this.modeToggle.button, this.zoomControls.root, this.externalButton);
    this.main.append(this.outline.panel, this.searchRail, this.readingBody);
    this.root.append(
      this.toolbar.root,
      this.notices,
      this.actionStatus,
      this.main,
    );
    this.contentEl.replaceChildren(this.root);

    this.searchButton.addEventListener("click", () => this.toggleSearch());
    this.externalButton.addEventListener("click", () => {
      void this.openCurrentFileExternally();
    });
    this.readingBody.addEventListener("click", (event) => {
      void this.handleReadingBodyClick(event);
    });
    this.readingBody.addEventListener("auxclick", (event) => {
      void this.handleReadingBodyClick(event);
    });
    this.scope = new Scope(this.app.scope);
    this.scope.register(["Mod"], "f", () => {
      this.openSearch();
      return false;
    });
  }

  override getViewType(): string {
    return DOCX_VIEW_TYPE;
  }

  override getDisplayText(): string {
    return this.file?.basename ??
      this.dependencies.messages.text("fallbackTitle");
  }

  getPerformanceDiagnostics(): DocxViewPerformanceDiagnostics {
    return { ...this.diagnostics };
  }

  override async onLoadFile(file: TFile): Promise<void> {
    const startedAt = performance.now();
    this.abortController?.abort();
    const controller = new AbortController();
    this.abortController = controller;
    this.zoomControls.setZoom();
    this.outline.setModel();
    this.rendererSession?.dispose();
    this.rendererSession = null;
    this.renderSource = null;
    this.modeToggle.setMode("full");
    this.modeToggle.button.disabled = true;
    this.currentFile = file;
    this.root.dataset.state = "loading";
    delete this.root.dataset.errorCategory;
    delete this.root.dataset.renderer;
    delete this.root.dataset.firstReadableMs;
    delete this.root.dataset.searchReadyMs;
    this.model = null;
    this.setActiveParagraph(null);
    this.searchPanel.close();
    this.updateSearchButton();
    this.restoreReadingShell();
    this.readingBody.replaceChildren();
    this.setNotice(this.dependencies.messages.text("loading"));
    this.actionStatus.textContent = "";
    this.searchButton.disabled = false;
    this.externalButton.disabled =
      this.dependencies.openInDefaultApplication === undefined;
    this.diagnostics = {
      candidate: null,
      sourceBytes: file.stat.size,
      paragraphCount: null,
      firstReadableMs: null,
      searchReadyMs: null,
    };

    try {
      const buffer = await (this.dependencies.readBinary?.(file, controller.signal)
        ?? this.app.vault.readBinary(file));
      controller.signal.throwIfAborted();
      const model = await inspectDocxPackage(buffer, controller.signal);
      controller.signal.throwIfAborted();
      this.model = model;
      this.renderSource = new DocxRenderSource(buffer, model);
      this.modeToggle.button.disabled = false;
      this.diagnostics = {
        ...this.diagnostics,
        sourceBytes: buffer.byteLength,
        paragraphCount: model.paragraphs.length,
        searchReadyMs: performance.now() - startedAt,
      };
      await this.renderDocument("full", startedAt);
    } catch (error) {
      if (!controller.signal.aborted) this.renderOpenError(error);
    }
  }

  private async renderDocument(mode: DocxRenderMode, startedAt = performance.now()): Promise<void> {
    const source = this.renderSource;
    if (source === null || this.disposed) return;
    this.abortController?.abort();
    const controller = new AbortController();
    this.abortController = controller;
    this.modeToggle.setMode(mode);
    this.restoreReadingShell();
    this.root.dataset.state = "loading";
    delete this.root.dataset.errorCategory;
    this.setNotice(this.dependencies.messages.text("loading"));
    this.actionStatus.textContent = "";
    const previous = this.rendererSession;
    const scrollRange = this.readingBody.scrollHeight - this.readingBody.clientHeight;
    const scrollProgress = scrollRange > 0 ? this.readingBody.scrollTop / scrollRange : 0;

    try {
      const session = await source.open(
        this.dependencies.renderer, this.readingBody, mode, controller.signal,
      );
      if (controller.signal.aborted) {
        session.dispose();
        return;
      }
      previous?.dispose();
      this.rendererSession = session;
      this.zoomControls.setZoom(session.zoom);
      const model = source.model;
      this.outline.setModel(model);
      if (session.managesUnavailableContent !== true) {
        this.installUnavailablePlaceholders(model, session);
      }
      this.searchButton.disabled = false;
      if (this.activeParagraph !== null) {
        this.setActiveParagraph(this.activeParagraph);
      } else if (scrollProgress > 0) {
        this.readingBody.scrollTop = scrollProgress *
          (this.readingBody.scrollHeight - this.readingBody.clientHeight);
      }
      const firstReadableMs = performance.now() - startedAt;
      this.diagnostics = { ...this.diagnostics, candidate: session.candidate, firstReadableMs };
      this.root.dataset.state = "ready";
      this.root.dataset.renderer = session.candidate;
      this.root.dataset.firstReadableMs = firstReadableMs.toFixed(1);
      const searchReadyMs = this.diagnostics.searchReadyMs;
      if (searchReadyMs !== null) this.root.dataset.searchReadyMs = searchReadyMs.toFixed(1);
      this.renderDocumentNotices(model, session);
    } catch (error) {
      if (controller.signal.aborted) return;
      if (previous !== null) {
        this.modeToggle.setMode(previous.candidate === "bounded-semantic" ? "simplified" : "full");
        this.root.dataset.state = "ready";
        this.renderDocumentNotices(source.model, previous);
        this.actionStatus.textContent = this.dependencies.messages.text("modeSwitchFailed");
      } else {
        this.renderOpenError(error, () => { void this.renderDocument(mode); });
      }
    }
  }

  override async onClose(): Promise<void> {
    this.dispose();
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.abortController?.abort();
    this.abortController = null;
    this.searchPanel.dispose();
    this.unregisterTextCopy();
    this.unregisterAttachmentOpening();
    this.zoomControls.setZoom();
    this.outline.setModel();
    this.rendererSession?.dispose();
    this.rendererSession = null;
    this.model = null;
    this.renderSource = null;
    this.currentFile = null;
    this.activeParagraph = null;
    this.contentEl.replaceChildren();
    this.onDisposed();
  }

  private openSearch(): void {
    this.outline.close();
    if (this.searchPanel.isOpen) {
      this.searchPanel.open();
      return;
    }
    this.searchPanel.open();
    this.updateSearchButton();
  }

  private closeSearch(): void {
    const wasOpen = this.searchPanel.isOpen;
    this.searchPanel.close();
    this.updateSearchButton();
    if (wasOpen) this.searchButton.focus();
  }

  private toggleSearch(): void {
    if (this.searchPanel.isOpen) this.closeSearch();
    else this.openSearch();
  }

  private updateSearchButton(): void {
    const open = this.searchPanel.isOpen;
    this.searchButton.setAttribute("aria-pressed", String(open));
    const key = open ? "searchClose" : "searchOpen";
    this.searchButton.title = this.dependencies.messages.text(key);
    this.searchButton.setAttribute(
      "aria-label",
      this.dependencies.messages.text(key),
    );
  }

  private renderDocumentNotices(
    model: DocxSemanticModel,
    session: DocxRendererSession,
  ): void {
    const notices: string[] = [];
    if (session.candidate === "bounded-semantic") {
      notices.push(this.dependencies.messages.text("showingSimplified"));
    }
    if (session.warnings.includes("preview-paragraph-mapping-degraded")) {
      notices.push(this.dependencies.messages.text("mappingDegraded"));
    }
    if (model.hasUnavailableBodyContent) {
      notices.push(this.dependencies.messages.text("unavailableContent"));
    }
    this.setNotice(notices.join(" "));
  }

  private setNotice(message: string): void {
    this.noticesText.textContent = message;
    this.notices.hidden = message.length === 0;
  }

  private installUnavailablePlaceholders(
    model: DocxSemanticModel,
    session: DocxRendererSession,
  ): void {
    const placeholderText =
      this.dependencies.messages.text("unavailablePlaceholder");
    for (const paragraph of model.paragraphs) {
      if (paragraph.unavailableContent.length === 0) continue;
      const element = session.paragraphElements.get(paragraph.ordinal);
      if (element === undefined) continue;
      const placeholder = createSpan();
      placeholder.className = "office-viewer-docx-unavailable-content";
      placeholder.setAttribute("role", "note");
      placeholder.dataset.docxUnavailableKinds =
        paragraph.unavailableContent.join(",");
      placeholder.textContent = placeholderText;
      element.append(placeholder);
    }
    for (const block of model.unavailableBodyBlocks) {
      const placeholder = createDiv();
      placeholder.className = "office-viewer-docx-unavailable-content";
      placeholder.setAttribute("role", "note");
      placeholder.dataset.docxUnavailableKinds = block.kinds.join(",");
      placeholder.textContent = placeholderText;
      const preceding =
        session.paragraphElements.get(block.afterParagraphOrdinal);
      if (preceding === undefined) {
        this.readingBody.prepend(placeholder);
      } else {
        preceding.after(placeholder);
      }
    }
  }

  private setActiveParagraph(paragraph: DocxSemanticParagraph | null): void {
    this.activeParagraphElement?.classList.remove("is-active-docx-paragraph");
    this.activeParagraphElement?.removeAttribute("aria-current");
    this.activeParagraphElement = null;
    this.activeParagraph = paragraph;
    if (paragraph === null) {
      this.searchPanel.syncCurrentResult();
      return;
    }
    const element = this.rendererSession?.revealParagraph(paragraph.ordinal);
    if (element == null) return;
    this.activeParagraphElement = element;
    element.classList.add("is-active-docx-paragraph");
    element.setAttribute("aria-current", "true");
    element.scrollIntoView({ block: "center", behavior: "auto" });
    this.searchPanel.syncCurrentResult();
  }

  private activateOrdinal(ordinal: number): boolean {
    const paragraph = this.model?.paragraphs[ordinal - 1];
    if (paragraph === undefined || paragraph.ordinal !== ordinal) return false;
    this.setActiveParagraph(paragraph);
    return true;
  }

  private async handleReadingBodyClick(event: MouseEvent): Promise<void> {
    const ElementConstructor = this.readingBody.ownerDocument.defaultView?.Element;
    if (ElementConstructor === undefined) return;
    const target = event.composedPath().find((node): node is Element =>
      node instanceof ElementConstructor,
    );
    if (target === undefined) return;
    const external = target.closest<HTMLAnchorElement>(
      "a[data-docx-external-link]",
    );
    if (external !== null) {
      event.preventDefault();
      event.stopPropagation();
      if (event.type !== "click") return;
      const url = external.getAttribute("href") ?? "";
      if (
        !safeExternalProtocol(url) ||
        this.dependencies.openExternalUrl === undefined
      ) {
        this.actionStatus.textContent =
          this.dependencies.messages.text("blockedLink");
        return;
      }
      try {
        await this.dependencies.openExternalUrl(url);
      } catch {
        this.actionStatus.textContent =
          this.dependencies.messages.text("blockedLink");
      }
      return;
    }
    const bookmark = target.closest<HTMLElement>("[data-docx-bookmark]");
    if (bookmark !== null) {
      event.preventDefault();
      if (event.type !== "click") return;
      const tocTarget = bookmark.dataset.docxTocTarget;
      if (tocTarget !== undefined) {
        const heading = this.rendererSession?.revealParagraph(Number(tocTarget));
        if (heading == null) this.actionStatus.textContent = this.dependencies.messages.text("bookmarkUnavailable");
        else {
          this.setActiveParagraph(null);
          scrollToDocxHeading(this.readingBody, heading);
        }
        return;
      }
      const name = bookmark.dataset.docxBookmark ?? "";
      const targets =
        this.model?.bookmarkTargets.filter(
          (candidate) => candidate.bookmark === name,
        ) ?? [];
      if (
        targets.length !== 1 ||
        !this.activateOrdinal(targets[0]?.paragraphOrdinal ?? Number.NaN)
      ) {
        this.actionStatus.textContent =
          this.dependencies.messages.text("bookmarkUnavailable");
      }
    }
  }

  private async openCurrentFileExternally(): Promise<void> {
    const file = this.currentFile;
    const open = this.dependencies.openInDefaultApplication;
    if (file === null || open === undefined) return;
    try {
      await open(file);
    } catch {
      this.actionStatus.textContent =
        this.dependencies.messages.text("openDefaultFailure");
    }
  }

  private renderOpenError(error: unknown, retryRender?: () => void): void {
    this.zoomControls.setZoom();
    this.outline.setModel();
    this.readingBody.replaceChildren();
    this.rendererSession = null;
    this.searchPanel.close();
    this.updateSearchButton();
    this.setNotice("");
    this.actionStatus.textContent = "";
    const key =
      error instanceof DocxOpenError
        ? error.category === "resource-exhausted"
          ? "resourceExhausted"
          : error.category
        : "unknown";
    this.root.dataset.state = "error";
    this.root.dataset.errorCategory = key;
    this.searchButton.disabled = true;
    const file = this.currentFile;
    const { panel } = createOfficeViewerErrorSurface({
      title: this.dependencies.messages.text(key),
      safetyNote: this.dependencies.messages.text("sourceUnmodified"),
      retry: {
        label: this.dependencies.messages.text("retry"),
        action: "retry",
        onClick: () => {
          if (retryRender !== undefined) retryRender();
          else if (file !== null) void this.onLoadFile(file);
        },
      },
      openExternal:
        file !== null &&
        this.dependencies.openInDefaultApplication !== undefined
          ? {
              label: this.dependencies.messages.text("openDefault"),
              action: "open-externally",
              onClick: () => {
                void this.openCurrentFileExternally();
              },
            }
          : undefined,
      classNames: {
        root: "office-viewer-docx-error",
      },
    });
    this.main.replaceChildren(panel);
  }

  private restoreReadingShell(): void {
    if (this.main.contains(this.readingBody)) return;
    this.main.replaceChildren(this.outline.panel, this.searchRail, this.readingBody);
  }
}
