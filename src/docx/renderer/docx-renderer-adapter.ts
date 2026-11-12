import type {
  DocxSemanticModel,
  DocxSemanticParagraph,
} from "../docx-semantic-model";

export type DocxRendererCandidate =
  | "docx-preview"
  | "bounded-semantic";

export type DocxRenderMode = "full" | "simplified";

export interface DocxRendererSession {
  readonly candidate: DocxRendererCandidate;
  readonly paragraphElements: ReadonlyMap<number, HTMLElement>;
  readonly warnings: readonly string[];
  readonly managesUnavailableContent?: boolean;
  revealParagraph(ordinal: number): HTMLElement | null;
  dispose(): void;
}

export interface DocxRendererAdapter {
  open(
    buffer: ArrayBuffer,
    container: HTMLElement,
    model: DocxSemanticModel,
    signal: AbortSignal,
    mode?: DocxRenderMode,
  ): Promise<DocxRendererSession>;
}

export function normalizedRenderedParagraphText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

/**
 * Mapping comparison ignores whitespace differences that docx-preview
 * introduces around breaks and hyperlink boundaries, while still requiring
 * identical visible characters in document order.
 */
export function comparableParagraphText(value: string): string {
  return value.replace(/\s+/g, "");
}

const PARAGRAPH_SELECTOR = "p, h1, h2, h3, h4, h5, h6, li";

function leafRenderedParagraphs(container: HTMLElement): HTMLElement[] {
  return Array.from(
    container.querySelectorAll<HTMLElement>(PARAGRAPH_SELECTOR),
  ).filter(
    (element) => element.querySelector(PARAGRAPH_SELECTOR) === null,
  );
}

function visibleRenderedParagraphs(container: HTMLElement): HTMLElement[] {
  return leafRenderedParagraphs(container).filter(
    (element) =>
      comparableParagraphText(element.textContent ?? "").length > 0,
  );
}

export function mapRenderedParagraphs(
  container: HTMLElement,
  paragraphs: readonly DocxSemanticParagraph[],
): ReadonlyMap<number, HTMLElement> {
  const rendered = visibleRenderedParagraphs(container);
  if (rendered.length !== paragraphs.length) {
    throw new Error(
      `Renderer paragraph count ${rendered.length} does not match semantic paragraph count ${paragraphs.length}`,
    );
  }
  const mapping = new Map<number, HTMLElement>();
  for (let index = 0; index < paragraphs.length; index += 1) {
    const semantic = paragraphs[index];
    const element = rendered[index];
    if (semantic === undefined || element === undefined) {
      throw new Error("Renderer paragraph mapping is incomplete");
    }
    const semanticText = comparableParagraphText(semantic.text);
    const renderedText = comparableParagraphText(element.textContent ?? "");
    if (semanticText !== renderedText) {
      throw new Error(
        `Renderer paragraph ${semantic.ordinal} text does not match the semantic model`,
      );
    }
    element.dataset.docxParagraphOrdinal = String(semantic.ordinal);
    mapping.set(semantic.ordinal, element);
  }
  return mapping;
}

/**
 * Order-preserving exact-character mapping used when positional mapping fails.
 * Only binds paragraphs whose comparable text matches; never invents a binding.
 */
export function alignRenderedParagraphs(
  container: HTMLElement,
  paragraphs: readonly DocxSemanticParagraph[],
): ReadonlyMap<number, HTMLElement> {
  const rendered = visibleRenderedParagraphs(container);
  const mapping = new Map<number, HTMLElement>();
  let renderedIndex = 0;
  for (const semantic of paragraphs) {
    const semanticText = comparableParagraphText(semantic.text);
    if (semanticText.length === 0) continue;
    let found = -1;
    for (let index = renderedIndex; index < rendered.length; index += 1) {
      const element = rendered[index];
      if (element === undefined) continue;
      if (comparableParagraphText(element.textContent ?? "") === semanticText) {
        found = index;
        break;
      }
    }
    if (found < 0) continue;
    const element = rendered[found];
    if (element === undefined) continue;
    element.dataset.docxParagraphOrdinal = String(semantic.ordinal);
    mapping.set(semantic.ordinal, element);
    renderedIndex = found + 1;
  }
  return mapping;
}

export function sanitizeRenderedDocx(container: HTMLElement): void {
  for (const element of container.querySelectorAll(
    "script, iframe, frame, object, embed, form, input, textarea, select, button, meta, base",
  )) {
    element.remove();
  }
  for (const element of container.querySelectorAll<HTMLElement>("*")) {
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      if (
        name.startsWith("on") ||
        name === "srcdoc" ||
        name === "formaction"
      ) {
        element.removeAttribute(attribute.name);
      }
    }
    if (element.instanceOf(HTMLAnchorElement)) {
      const rawHref = element.getAttribute("href");
      if (rawHref === null) continue;
      if (rawHref.startsWith("#")) {
        element.dataset.docxBookmark = rawHref.slice(1);
        element.removeAttribute("href");
        continue;
      }
      let protocol: string | null = null;
      try {
        protocol = new URL(rawHref).protocol;
      } catch {
        protocol = null;
      }
      if (
        protocol !== "https:" &&
        protocol !== "http:" &&
        protocol !== "mailto:"
      ) {
        element.removeAttribute("href");
      } else {
        element.dataset.docxExternalLink = "true";
        element.setAttribute("rel", "noopener noreferrer");
      }
    }
    if (element.instanceOf(HTMLImageElement)) {
      const source = element.getAttribute("src") ?? "";
      if (!source.startsWith("data:") && !source.startsWith("blob:")) {
        element.removeAttribute("src");
        element.dataset.docxUnavailableImage = "true";
      }
    }
  }
  for (const style of container.querySelectorAll("style")) {
    const css = style.textContent ?? "";
    if (
      /@import/i.test(css) ||
      /url\(\s*(['"]?)\s*(?:https?:|file:|javascript:)/i.test(css)
    ) {
      style.remove();
    }
  }
}

/**
 * Adds project-owned hooks around docx-preview media and tables so reading-mode
 * CSS can shrink fixed-size content without depending on the renderer's
 * inline-style shape. Authored widths remain upper bounds on wide panes.
 */
export function prepareRenderedDocxReadingLayout(
  container: HTMLElement,
): void {
  const wrapper = container.querySelector<HTMLElement>(".docx-wrapper");
  if (wrapper !== null) {
    wrapper.style.removeProperty("align-items");
    wrapper.style.removeProperty("background");
    wrapper.style.removeProperty("padding");
    wrapper.style.removeProperty("width");
  }
  const page = wrapper?.querySelector<HTMLElement>(":scope > section.docx");
  if (page !== null && page !== undefined) {
    page.style.removeProperty("box-shadow");
    page.style.removeProperty("margin");
    page.style.removeProperty("padding");
    page.style.removeProperty("padding-inline");
    page.style.removeProperty("width");
  }
  for (const image of container.querySelectorAll<HTMLImageElement>("img")) {
    image.classList.add("office-viewer-docx-media");
    image.style.removeProperty("height");
    image.style.removeProperty("max-width");
    image.style.removeProperty("width");
    const wrapper = image.parentElement;
    if (
      wrapper === null ||
      wrapper.tagName !== "DIV" ||
      wrapper.style.display !== "inline-block" ||
      wrapper.children.length !== 1 ||
      wrapper.firstElementChild !== image
    ) {
      continue;
    }
    wrapper.classList.add("office-viewer-docx-media-wrapper");
    const authoredWidth = wrapper.style.width.trim();
    if (authoredWidth.length > 0) {
      wrapper.style.setProperty(
        "--office-viewer-docx-media-width",
        authoredWidth,
      );
    }
    wrapper.style.removeProperty("height");
    wrapper.style.removeProperty("max-width");
    wrapper.style.removeProperty("width");
  }
  for (const table of container.querySelectorAll<HTMLTableElement>("table")) {
    table.classList.add("office-viewer-docx-table");
    const authoredWidth = table.style.width.trim();
    if (authoredWidth.length > 0 && authoredWidth !== "auto") {
      table.style.setProperty(
        "--office-viewer-docx-table-width",
        authoredWidth,
      );
    }
    table.style.removeProperty("max-width");
    table.style.removeProperty("table-layout");
    table.style.removeProperty("width");
    const columns = Array.from(
      table.querySelectorAll<HTMLTableColElement>("col"),
    ).filter((column) => column.closest("table") === table);
    const parsedWidths = columns.map((column) => {
      const match = column.style.width
        .trim()
        .match(/^([0-9]+(?:\.[0-9]+)?)([a-z]+)$/i);
      if (match === null) return null;
      const value = Number.parseFloat(match[1] ?? "");
      const unit = match[2]?.toLowerCase() ?? "";
      return Number.isFinite(value) && value > 0 ? { value, unit } : null;
    });
    const unit = parsedWidths[0]?.unit;
    if (
      unit === undefined ||
      parsedWidths.some((width) => width === null || width.unit !== unit)
    ) {
      continue;
    }
    const total = parsedWidths.reduce(
      (sum, width) => sum + (width?.value ?? 0),
      0,
    );
    if (total <= 0) continue;
    columns.forEach((column, index) => {
      const width = parsedWidths[index];
      if (width === undefined || width === null) return;
      column.style.setProperty(
        "--office-viewer-docx-column-width",
        `${(width.value / total) * 100}%`,
      );
      column.style.removeProperty("min-width");
      column.style.removeProperty("width");
    });
  }
}
