import { renderAsync } from "docx-preview";
import { DocxOpenError } from "../docx-open-error";
import type { DocxSemanticModel } from "../docx-semantic-model";
import { prepareRenderedDocxPageSpacing } from "./docx-preview-page-layout";
import {
  alignRenderedParagraphs,
  comparableParagraphText,
  mapRenderedParagraphs,
  prepareRenderedDocxReadingLayout,
  sanitizeRenderedDocx,
  type DocxRendererAdapter,
  type DocxRendererSession,
} from "./docx-renderer-adapter";

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
      comparableParagraphText(element.textContent ?? "") === needle,
  );
  if (matches.length !== 1) return null;
  const element = matches[0];
  if (element === undefined) return null;
  element.dataset.docxParagraphOrdinal = String(ordinal);
  return element;
}

export class DocxPreviewRendererAdapter implements DocxRendererAdapter {
  constructor(private readonly pageLayout = false) {}

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
        "incompatible",
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
          "incompatible",
          "docx-preview output does not map to the project-owned paragraph model",
        );
      }
    }
    signal.throwIfAborted();
    container.replaceChildren(staging);
    let stopFitting: (() => void) | undefined;
    if (this.pageLayout) {
      staging.classList.add("office-viewer-docx--paged");
      const style = container.ownerDocument.createElement("style");
      style.textContent = `
        :host { font: 11pt "Calibri", "Carlito", sans-serif; }
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
      staging.shadowRoot?.prepend(style);
      prepareRenderedDocxPageSpacing(content);
      stopFitting = fitDocxPages(container, content);
    }
    return {
      candidate: "docx-preview",
      paragraphElements,
      warnings,
      revealParagraph: (ordinal) =>
        revealMappedOrUniqueText(
          content,
          paragraphElements,
          model,
          ordinal,
        ),
      dispose: () => {
        stopFitting?.();
        content.replaceChildren();
        if (staging.parentElement === container) container.replaceChildren();
      },
    };
  }
}

function fitDocxPages(container: HTMLElement, content: HTMLElement): () => void {
  const wrapper = content.querySelector<HTMLElement>(".docx-wrapper");
  if (wrapper === null) return () => undefined;
  wrapper.style.padding = "12px";
  wrapper.style.background = "transparent";
  wrapper.style.gap = "12px";
  const pages = [...wrapper.querySelectorAll<HTMLElement>(":scope > section.docx")];
  for (const page of pages) {
    if (page.style.width === "") page.style.width = "612pt";
    if (page.style.minHeight === "") page.style.minHeight = "792pt";
    for (const side of ["top", "right", "bottom", "left"]) {
      const property = `padding-${side}`;
      if (page.style.getPropertyValue(property) === "") {
        page.style.setProperty(property, "72pt");
      }
    }
    page.style.margin = "0";
    page.style.boxShadow = "0 1px 4px rgb(0 0 0 / 25%)";
  }
  const resize = () => {
    const available = wrapper.clientWidth - 24;
    if (available <= 0) return;
    const widths = pages.map((page) => page.offsetWidth);
    for (const [index, page] of pages.entries()) {
      const width = widths[index] ?? 0;
      if (width > 0) {
        page.style.setProperty("zoom", String(Math.min(1, available / width)));
      }
    }
  };
  resize();
  const Observer = container.ownerDocument.defaultView?.ResizeObserver;
  const observer = Observer === undefined ? null : new Observer(resize);
  observer?.observe(wrapper);
  return () => observer?.disconnect();
}
