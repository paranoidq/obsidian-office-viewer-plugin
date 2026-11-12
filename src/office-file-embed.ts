import { MarkdownRenderChild, type MarkdownPostProcessorContext } from "obsidian";
import type { DocxMessageTranslator } from "./docx/docx-messages";
import { registerDocxAttachmentOpening } from "./docx/docx-attachment-controls";
import type { DocxAttachmentFile } from "./docx/docx-attachments";
import { createDocxModeButton } from "./docx/docx-mode-button";
import { DocxOpenError } from "./docx/docx-open-error";
import { createDocxOutline, scrollToDocxHeading } from "./docx/docx-outline";
import { DocxRenderSource } from "./docx/docx-render-source";
import { inspectDocxPackage } from "./docx/docx-semantic-model";
import { createDocxZoomControls } from "./docx/docx-zoom-controls";
import type {
  DocxRenderMode,
  DocxRendererAdapter,
  DocxRendererSession,
} from "./docx/renderer/docx-renderer-adapter";
import type { MessageTranslator } from "./i18n";
import { LegacyOfficeUnavailable } from "./legacy-office-preview";
import { registerOfficeEmbedTextCopy } from "./office-embed-text-copy";
import {
  parseOfficeFileEmbedTarget,
  type OfficeFileEmbedTarget,
} from "./office-file-embed-target";
import { decorateOfficeViewerIconButton } from "./office-viewer-chrome";
import type { PptxRendererAdapter } from "./renderer/pptx-renderer-adapter";
import { mapSlideEmbedOpenError } from "./slide-embed-core";
import type { SlideEmbedScheduler } from "./slide-embed-scheduler";
import { observeSlideEmbedVisibility } from "./slide-embed-visibility";

export interface OfficeFileEmbedOptions<TFile> {
  readonly resolveFile: (sourcePath: string, notePath: string) => TFile | null;
  readonly readBinary: (file: TFile, signal: AbortSignal) => Promise<ArrayBuffer>;
  readonly openInTab: (file: TFile) => Promise<void>;
  readonly openExternally?: (file: TFile) => Promise<void>;
  readonly openExternalUrl: (url: string) => Promise<void>;
  readonly openAttachment?: (file: DocxAttachmentFile) => Promise<void>;
  readonly pptxRenderer: PptxRendererAdapter;
  readonly docxRenderer: DocxRendererAdapter;
  readonly scheduler: SlideEmbedScheduler;
  readonly messages: MessageTranslator;
  readonly docxMessages: DocxMessageTranslator;
}

function iconButton(parent: HTMLElement, action: string, label: string, icon: string) {
  const button = parent.createEl("button", {
    attr: { type: "button", "data-action": action, "aria-label": label, title: label },
  });
  decorateOfficeViewerIconButton(button, icon);
  return button;
}

/** Mount a whole-file preview for either Markdown Reading View or Live Preview. */
export function mountOfficeFileEmbed<TFile>(
  host: HTMLElement,
  target: OfficeFileEmbedTarget,
  notePath: string,
  options: OfficeFileEmbedOptions<TFile>,
): () => void {
  const { sourcePath, format } = target;
  const file = options.resolveFile(sourcePath, notePath);
  const { messages, docxMessages } = options;
  const filename = sourcePath.split("/").pop() ?? sourcePath;
  host.classList.add("office-file-embed");
  host.dataset.format = format;
  host.dataset.state = file === null ? "missing-source" : "waiting";
  host.setAttribute("role", "group");
  host.setAttribute("aria-label", filename);
  const toolbar = host.createDiv({ cls: "office-file-embed__toolbar" });
  toolbar.createSpan({ cls: "office-file-embed__title", text: filename, attr: { title: filename } });
  const navigation = toolbar.createDiv({ cls: "office-file-embed__navigation" });
  navigation.hidden = true;
  const status = host.createDiv({
    cls: "office-file-embed__status",
    attr: { role: "status" },
  });
  const actionStatus = host.createDiv({
    cls: "office-file-embed__status",
    attr: { role: "status" },
  });
  if (file === null) {
    status.textContent = messages.text("fileEmbed.sourceMissing");
    return () => host.replaceChildren();
  }
  const modeToggle = format === "docx"
    ? createDocxModeButton(host.ownerDocument, docxMessages, (mode) => startPreview(mode))
    : null;
  if (modeToggle !== null) toolbar.append(modeToggle.button);
  const zoomControls = format === "docx"
    ? createDocxZoomControls(host.ownerDocument, docxMessages)
    : null;
  if (zoomControls !== null) toolbar.append(zoomControls.root);
  let docxSession: DocxRendererSession | undefined;
  const outline = format === "docx"
    ? createDocxOutline(host.ownerDocument, docxMessages, (ordinal) => {
      const element = docxSession?.revealParagraph(ordinal);
      if (element == null) return false;
      scrollToDocxHeading(body, element);
      return true;
    })
    : null;
  if (outline !== null) toolbar.append(outline.button);
  iconButton(toolbar, "open-in-tab", messages.text("fileEmbed.openInTab"), "panel-top")
    .addEventListener("click", (event) => {
      event.stopPropagation();
      actionStatus.textContent = "";
      void options.openInTab(file).catch(() => {
        actionStatus.textContent = messages.text("fileEmbed.openInTabFailure");
      });
    });
  const open = options.openExternally;
  if (open !== undefined) {
    iconButton(toolbar, "open-externally", messages.text("external.open"), "external-link")
      .addEventListener("click", (event) => {
        event.stopPropagation();
        actionStatus.textContent = "";
        void open(file).catch(() => {
          actionStatus.textContent = messages.text("external.failure");
        });
      });
  }
  const layout = outline === null ? host : host.createDiv({ cls: "office-file-embed__docx-layout" });
  if (outline !== null) layout.append(outline.panel);
  const body = layout.createDiv({ cls: "office-file-embed__body" });
  const unregisterTextCopy = registerOfficeEmbedTextCopy(body);
  let disposeSession: (() => void) | undefined;
  let cancel: (() => void) | undefined;
  let disposed = false;
  let docxSource: DocxRenderSource | undefined;
  const unregisterAttachmentOpening = format === "docx"
    ? registerDocxAttachmentOpening(body, () => docxSource, options.openAttachment, docxMessages, actionStatus)
    : undefined;
  let renderedDocxMode: DocxRenderMode | undefined;
  const showStatus = (message: string) => { status.textContent = message; };
  const loading = format === "pptx" ? messages.text("viewer.loading") : docxMessages.text("loading");
  const showDocxStatus = (mode: DocxRenderMode) => {
    showStatus(mode === "simplified"
      ? docxMessages.text("showingSimplified")
      : docxSource?.model.hasUnavailableBodyContent
        ? messages.text("fileEmbed.docxPartial")
        : "");
  };
  const startPreview = (mode: DocxRenderMode = "full") => {
    if (disposed) return;
    cancel?.();
    modeToggle?.setMode(mode);
    showStatus(loading);
    actionStatus.textContent = "";
    cancel = options.scheduler.schedule(async (signal) => {
      host.dataset.state = "loading";
      try {
        if (format === "pptx") {
          const buffer = await options.readBinary(file, signal);
          signal.throwIfAborted();
          const session = await options.pptxRenderer.open(buffer, body, signal);
          if (signal.aborted) {
            session.dispose();
            return;
          }
          disposeSession = () => session.dispose();
          body.style.aspectRatio = `${session.slideWidth} / ${session.slideHeight}`;
          const previous = iconButton(
            navigation, "previous-slide", messages.text("navigation.previous"), "chevron-left",
          );
          const counter = navigation.createSpan({ attr: { role: "status" } });
          const next = iconButton(
            navigation, "next-slide", messages.text("navigation.next"), "chevron-right",
          );
          let index = 0;
          let rendering = false;
          const render = async (targetIndex: number) => {
            if (rendering || signal.aborted || targetIndex < 0 || targetIndex >= session.slideCount) return;
            rendering = true;
            previous.disabled = next.disabled = true;
            index = targetIndex;
            try {
              await session.renderSlide(index);
              if (signal.aborted) return;
              host.dataset.state = "ready";
              showStatus("");
              counter.textContent = messages.text("page.counter", {
                current: index + 1, total: session.slideCount,
              });
            } catch (error) {
              if (signal.aborted) return;
              if (host.dataset.state !== "ready") {
                host.dataset.state = "error";
                showStatus(mapSlideEmbedOpenError(messages, error));
              } else {
                showStatus(messages.text("navigation.renderFailure", { slide: index + 1 }));
              }
            } finally {
              rendering = false;
              previous.disabled = index === 0;
              next.disabled = index === session.slideCount - 1;
            }
          };
          navigation.hidden = session.slideCount < 2;
          previous.addEventListener("click", () => { void render(index - 1); });
          next.addEventListener("click", () => { void render(index + 1); });
          await render(0);
        } else {
          if (docxSource === undefined) {
            const buffer = await options.readBinary(file, signal);
            const model = await inspectDocxPackage(buffer, signal);
            signal.throwIfAborted();
            docxSource = new DocxRenderSource(buffer, model);
            if (modeToggle !== null) modeToggle.button.disabled = false;
          }
          const scrollRange = body.scrollHeight - body.clientHeight;
          const scrollProgress = scrollRange > 0 ? body.scrollTop / scrollRange : 0;
          const session = await docxSource.open(
            options.docxRenderer, body, mode, signal,
          );
          if (signal.aborted) {
            session.dispose();
            return;
          }
          disposeSession?.();
          disposeSession = () => session.dispose();
          docxSession = session;
          outline?.setModel(docxSource.model);
          zoomControls?.setZoom(session.zoom);
          renderedDocxMode = mode;
          host.dataset.renderer = session.candidate;
          host.dataset.state = "ready";
          showDocxStatus(mode);
          if (scrollProgress > 0) {
            body.scrollTop = scrollProgress * (body.scrollHeight - body.clientHeight);
          }
        }
      } catch (error) {
        if (signal.aborted) return;
        if (renderedDocxMode !== undefined) {
          modeToggle?.setMode(renderedDocxMode);
          host.dataset.state = "ready";
          showDocxStatus(renderedDocxMode);
          actionStatus.textContent = docxMessages.text("modeSwitchFailed");
          return;
        }
        disposeSession?.();
        disposeSession = undefined;
        host.dataset.state = "error";
        showStatus(error instanceof LegacyOfficeUnavailable
          ? messages.text("fileEmbed.legacyUnavailable")
          : format === "pptx"
          ? mapSlideEmbedOpenError(messages, error)
          : error instanceof DocxOpenError && error.category === "render-failed"
          ? docxMessages.text("renderFailed")
          : messages.text("fileEmbed.renderFailure"));
      }
    });
  };
  showStatus(loading);
  const stopObserving = observeSlideEmbedVisibility(host, (visible) => {
    if (visible && cancel === undefined) startPreview();
  });
  // Intercept DOCX links before the note or Electron can navigate. composedPath
  // reaches links inside the page renderer's shadow root as well.
  if (format === "docx") {
    const handleLink = (event: MouseEvent) => {
      const Anchor = host.ownerDocument.defaultView?.HTMLAnchorElement;
      if (Anchor === undefined) return;
      const link = event.composedPath().find((node): node is HTMLAnchorElement =>
        node instanceof Anchor,
      );
      if (link === undefined || !link.hasAttribute("href")) return;
      event.preventDefault();
      event.stopPropagation();
      if (event.type !== "click") return;
      const tocTarget = link.dataset.docxTocTarget;
      if (tocTarget !== undefined) {
        const heading = docxSession?.revealParagraph(Number(tocTarget));
        if (heading == null) actionStatus.textContent = docxMessages.text("bookmarkUnavailable");
        else scrollToDocxHeading(body, heading);
        return;
      }
      void options.openExternalUrl(link.getAttribute("href") ?? "").catch(() => {
        actionStatus.textContent = docxMessages.text("blockedLink");
      });
    };
    body.addEventListener("click", handleLink);
    body.addEventListener("auxclick", handleLink);
  }
  return () => {
    disposed = true;
    unregisterTextCopy();
    unregisterAttachmentOpening?.();
    stopObserving();
    cancel?.();
    zoomControls?.setZoom();
    outline?.setModel();
    docxSession = undefined;
    disposeSession?.();
    host.replaceChildren();
  };
}

class OfficeFileEmbedChild extends MarkdownRenderChild {
  private disposePreview: (() => void) | undefined;
  private nativeEmbed: HTMLElement;
  private observer: MutationObserver | undefined;

  constructor(
    host: HTMLElement,
    nativeEmbed: HTMLElement,
    private readonly src: string,
    private readonly mount: () => () => void,
    private readonly unregister: (child: MarkdownRenderChild) => void,
  ) {
    super(host);
    this.nativeEmbed = nativeEmbed;
  }

  override onload(): void {
    const hideNativeEmbed = () => {
      const previous = this.containerEl.previousElementSibling;
      if (previous !== this.nativeEmbed && previous?.matches(".internal-embed")
        && previous.getAttribute("src") === this.src) {
        this.nativeEmbed.classList.remove("office-file-embed__native");
        this.nativeEmbed = previous as HTMLElement;
      }
      this.nativeEmbed.classList.add("office-file-embed__native");
    };
    hideNativeEmbed();
    const parent = this.containerEl.parentElement;
    const Observer = parent?.ownerDocument.defaultView?.MutationObserver;
    if (parent !== null && Observer !== undefined) {
      this.observer = new Observer(hideNativeEmbed);
      this.observer.observe(parent, { childList: true });
    }
    this.disposePreview = this.mount();
  }

  override onunload(): void {
    this.observer?.disconnect();
    this.disposePreview?.();
    this.nativeEmbed.classList.remove("office-file-embed__native");
    this.containerEl.remove();
    this.unregister(this);
  }
}

export function processOfficeFileEmbeds<TFile>(
  root: HTMLElement,
  context: MarkdownPostProcessorContext,
  options: OfficeFileEmbedOptions<TFile>,
  lifecycle: { register(child: MarkdownRenderChild): void; unregister(child: MarkdownRenderChild): void },
): void {
  const selector = ".internal-embed[src]";
  const embeds = [...root.querySelectorAll<HTMLElement>(selector)];
  if (root.matches(selector)) embeds.unshift(root);
  for (const nativeEmbed of embeds) {
    if (nativeEmbed.parentElement === null
      || nativeEmbed.classList.contains("office-file-embed__native")) continue;
    const src = nativeEmbed.getAttribute("src") ?? "";
    const target = parseOfficeFileEmbedTarget(src);
    if (target === null) continue;
    if (nativeEmbed.nextElementSibling?.getAttribute("data-office-file-embed-src") === src) continue;
    const host = root.ownerDocument.createElement("div");
    host.dataset.officeFileEmbedSrc = src;
    nativeEmbed.classList.add("office-file-embed__native");
    nativeEmbed.after(host);
    const child = new OfficeFileEmbedChild(
      host, nativeEmbed, src,
      () => mountOfficeFileEmbed(host, target, context.sourcePath, options),
      lifecycle.unregister,
    );
    lifecycle.register(child);
    context.addChild(child);
  }
}
