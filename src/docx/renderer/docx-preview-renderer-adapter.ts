import { renderAsync } from "docx-preview";
import { createDocxAttachmentButton, DOCX_ATTACHMENT_STYLES } from "../docx-attachment-controls";
import { createDocxMessageTranslator } from "../docx-messages";
import { DOCX_EQUATION_STYLES, renderDocxEquation } from "../docx-equations";
import { DOCX_TOC_STYLES, layoutDocxTocEntry } from "./docx-preview-toc";
import { DocxOpenError } from "../docx-open-error";
import type { DocxSemanticModel } from "../docx-semantic-model";
import {
  compactDocxPreviewMargins,
  fitWideDocxPageTables,
  prepareRenderedDocxPageSpacing,
} from "./docx-preview-page-layout";
import { createDocxZoom } from "./docx-wheel-zoom";
import {
  alignRenderedParagraphs,
  comparableParagraphText,
  mapRenderedParagraphs,
  markDocxHeading,
  prepareRenderedDocxReadingLayout,
  sanitizeRenderedDocx,
  type DocxRendererAdapter,
  type DocxRendererSession,
} from "./docx-renderer-adapter";

function renderedBodyText(element: HTMLElement): string {
  if (element.querySelector("math[data-docx-equation]") === null) return element.textContent ?? "";
  const walker = element.ownerDocument.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let text = "";
  for (let child = walker.nextNode(); child !== null; child = walker.nextNode()) {
    if (!child.parentElement?.closest("math[data-docx-equation]")) text += child.textContent ?? "";
  }
  return text;
}

function revealMappedOrUniqueText(
  staging: HTMLElement,
  paragraphElements: ReadonlyMap<number, HTMLElement>,
  model: DocxSemanticModel,
  ordinal: number,
): HTMLElement | null {
  const mapped = paragraphElements.get(ordinal);
  if (mapped !== undefined) return mapped;
  const semantic = model.paragraphs[ordinal - 1];
  if (semantic === undefined || semantic.ordinal !== ordinal) return null;
  const needle = comparableParagraphText(semantic.text);
  if (needle.length === 0) return null;
  const matches = Array.from(
    staging.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, h6, li"),
  ).filter(
    (element) =>
      element.querySelector("p, h1, h2, h3, h4, h5, h6, li") === null &&
      comparableParagraphText(renderedBodyText(element)) === needle,
  );
  if (matches.length !== 1) return null;
  const element = matches[0];
  if (element === undefined) return null;
  element.dataset.docxParagraphOrdinal = String(ordinal);
  markDocxHeading(element, semantic.headingLevel);
  return element;
}

export class DocxPreviewRendererAdapter implements DocxRendererAdapter {
  constructor(
    private readonly pageLayout = false,
    private readonly messages = createDocxMessageTranslator("en"),
  ) {}

  async open(
    buffer: ArrayBuffer,
    container: HTMLElement,
    model: DocxSemanticModel,
    signal: AbortSignal,
  ): Promise<DocxRendererSession> {
    signal.throwIfAborted();
    const staging = container.ownerDocument.createElement("div");
    staging.className = "office-viewer-docx office-viewer-docx--preview";
    // Keep note themes and styles from other DOCX previews out of the pages.
    const content = this.pageLayout
      ? staging.attachShadow({ mode: "open" }).appendChild(
          container.ownerDocument.createElement("div"),
        )
      : staging;
    try {
      await renderAsync(buffer.slice(0), content, content, {
        inWrapper: true,
        hideWrapperOnPrint: false,
        ignoreWidth: !this.pageLayout,
        ignoreHeight: !this.pageLayout,
        ignoreFonts: false,
        breakPages: this.pageLayout,
        debug: false,
        experimental: false,
        className: "docx",
        trimXmlDeclaration: true,
        renderHeaders: false,
        renderFooters: false,
        renderFootnotes: false,
        renderEndnotes: false,
        ignoreLastRenderedPageBreak: !this.pageLayout,
        useBase64URL: true,
        renderChanges: false,
        renderComments: false,
        renderAltChunks: false,
      });
    } catch (error) {
      signal.throwIfAborted();
      throw new DocxOpenError(
        "render-failed",
        "docx-preview could not render the DOCX",
        { cause: error },
      );
    }
    signal.throwIfAborted();
    sanitizeRenderedDocx(content);
    if (!this.pageLayout) prepareRenderedDocxReadingLayout(content);
    const warnings: string[] = [];
    let paragraphElements: ReadonlyMap<number, HTMLElement>;
    try {
      paragraphElements = mapRenderedParagraphs(content, model.paragraphs);
    } catch {
      // Keep the readable preview: bind only order-preserving exact-character
      // matches instead of failing the whole document open.
      paragraphElements = alignRenderedParagraphs(content, model.paragraphs);
      if (paragraphElements.size < model.paragraphs.length) {
        warnings.push("preview-paragraph-mapping-degraded");
      }
      if (paragraphElements.size === 0 && model.paragraphs.length > 0) {
        throw new DocxOpenError(
          "render-failed",
          "docx-preview output does not map to the project-owned paragraph model",
        );
      }
    }
    signal.throwIfAborted();
    for (const paragraph of model.paragraphs) {
      if (paragraph.headingLevel === null) continue;
      const element = paragraphElements.get(paragraph.ordinal);
      if (element !== undefined) markDocxHeading(element, paragraph.headingLevel);
    }
    if (model.attachments !== undefined || model.equations !== undefined) {
      const markers = new Map(Array.from(content.querySelectorAll<HTMLElement>("[id]"),
        (element) => [element.id, element] as const));
      for (const equation of model.equations ?? []) {
        markers.get(equation.marker)?.replaceWith(renderDocxEquation(content.ownerDocument, equation));
      }
      for (const attachment of model.attachments ?? []) {
        const marker = markers.get(attachment.marker);
        if (marker === undefined) continue;
        const following = marker.nextElementSibling as HTMLElement | null;
        const preview = attachment.hasPreview && following?.querySelector("img[src]")
          ? following : undefined;
        marker.replaceWith(createDocxAttachmentButton(
          content.ownerDocument, attachment, this.messages, preview,
        ));
      }
      const style = content.ownerDocument.createElement("style");
      style.textContent = (model.attachments === undefined ? "" : DOCX_ATTACHMENT_STYLES)
        + (model.equations === undefined ? "" : DOCX_EQUATION_STYLES);
      content.append(style);
    }
    container.replaceChildren(staging);
    let pageLayout: ReturnType<typeof fitDocxPages>;
    if (this.pageLayout) {
      staging.classList.add("office-viewer-docx--paged");
      for (const entry of model.tocEntries ?? []) {
        const paragraph = paragraphElements.get(entry.paragraphOrdinal);
        if (paragraph !== undefined) layoutDocxTocEntry(paragraph, entry);
      }
      const style = container.ownerDocument.createElement("style");
      style.textContent = `
        :host { font: 11pt "Calibri", "Carlito", sans-serif; }
        ${model.tocEntries === undefined ? "" : DOCX_TOC_STYLES}
        .docx p:where([data-docx-heading-level]) {
          font-size: var(--office-viewer-docx-heading-size);
          font-weight: 600;
        }
        .docx :where([data-docx-heading-level]) span:where(:not([class])) {
          font-size: inherit;
          font-weight: inherit;
        }
        .office-viewer-docx-table-scroll {
          max-width: 100%;
          overflow-x: auto;
        }
        .is-active-docx-paragraph {
          background: var(--text-highlight-bg, #fff1ae);
          outline: 2px solid var(--interactive-accent, #7761bd);
          outline-offset: 3px;
        }
        .office-viewer-docx-unavailable-content {
          display: block;
          border: 1px dashed var(--color-yellow, #c18e26);
          padding: 4px 8px;
        }
      `;
      staging.shadowRoot?.append(style);
      prepareRenderedDocxPageSpacing(content);
      pageLayout = fitDocxPages(container, content);
    }
    return {
      candidate: "docx-preview",
      paragraphElements,
      warnings,
      zoom: pageLayout?.zoom,
      revealParagraph: (ordinal) =>
        revealMappedOrUniqueText(
          content,
          paragraphElements,
          model,
          ordinal,
        ),
      dispose: () => {
        pageLayout?.dispose();
        content.replaceChildren();
        if (staging.parentElement === container) container.replaceChildren();
      },
    };
  }
}

function fitDocxPages(container: HTMLElement, content: HTMLElement) {
  const wrapper = content.querySelector<HTMLElement>(".docx-wrapper");
  if (wrapper === null) return;
  wrapper.style.padding = "12px";
  wrapper.style.background = "transparent";
  wrapper.style.gap = "12px";
  // A centered page wider than the viewport needs a reachable left edge.
  wrapper.style.alignItems = "safe center";
  const pages = Array.from(wrapper.children).filter(
    (child): child is HTMLElement => child.matches("section.docx"),
  );
  const tablesByPage = new Map<HTMLElement, HTMLTableElement[]>();
  for (const page of pages) {
    if (page.style.width === "") page.style.width = "612pt";
    if (page.style.minHeight === "") page.style.minHeight = "792pt";
    compactDocxPreviewMargins(page);
    page.style.margin = "0";
    page.style.boxShadow = "0 1px 4px rgb(0 0 0 / 25%)";
    const tables = Array.from(page.querySelectorAll("table")).filter(
      (table) => table.parentElement?.closest("table") == null,
    );
    if (tables.length > 0) tablesByPage.set(page, tables);
  }
  const pageElements = new Set<EventTarget>(pages);
  let zoom = 1;
  const resize = () => {
    const available = wrapper.clientWidth - 24;
    if (available <= 0) return;
    const widths = pages.map((page) => page.offsetWidth);
    for (const [index, page] of pages.entries()) {
      const width = widths[index] ?? 0;
      if (width > 0) {
        page.style.setProperty("zoom", String(Math.min(1, available / width) * zoom));
      }
    }
  };
  let disposed = false;
  const fitLayout = () => {
    if (disposed) return;
    for (const [page, tables] of tablesByPage) fitWideDocxPageTables(page, tables);
    resize();
  };
  fitLayout();
  const Observer = container.ownerDocument.defaultView?.ResizeObserver;
  const observer = Observer === undefined ? null : new Observer(fitLayout);
  observer?.observe(wrapper);
  content.addEventListener("load", fitLayout, true);
  void container.ownerDocument.fonts?.ready.then(fitLayout);
  const zoomControl = createDocxZoom(
    container,
    (x, y, event) => event?.composedPath().find(
      (target): target is HTMLElement => pageElements.has(target),
    ) ?? (content.getRootNode() as ShadowRoot).elementFromPoint?.(x, y)
      ?.closest<HTMLElement>("section.docx") ?? wrapper,
    (scale) => {
      zoom = scale;
      resize();
    },
  );
  return {
    zoom: zoomControl,
    dispose: () => {
      disposed = true;
      content.removeEventListener("load", fitLayout, true);
      zoomControl.dispose();
      observer?.disconnect();
    },
  };
}
