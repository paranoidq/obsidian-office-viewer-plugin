import type {
  DocxSemanticModel,
  DocxSemanticParagraph,
} from "../docx-semantic-model";
import type {
  DocxRenderMode,
  DocxRendererAdapter,
  DocxRendererSession,
} from "./docx-renderer-adapter";
import { markDocxHeading } from "./docx-renderer-adapter";
import { createDocxZoom } from "./docx-wheel-zoom";

export interface BoundedDocxRendererOptions {
  readonly windowSize: number;
  readonly unavailablePlaceholder: string;
}

const DEFAULT_OPTIONS: BoundedDocxRendererOptions = {
  windowSize: 240,
  unavailablePlaceholder: "This document content cannot be displayed.",
};

function positiveInteger(value: number, name: string): number {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new TypeError(`${name} must be a positive integer`);
  }
  return value;
}

function hyperlinkElement(
  hyperlink: DocxSemanticParagraph["hyperlinks"][number],
): HTMLElement {
  if (hyperlink.kind === "blocked") {
    const blocked = createSpan();
    blocked.className = "office-viewer-docx-blocked-link";
    blocked.dataset.docxBlockedLink = "true";
    blocked.textContent = hyperlink.label;
    return blocked;
  }
  const anchor = createEl("a");
  anchor.textContent = hyperlink.label;
  if (hyperlink.kind === "external") {
    anchor.href = hyperlink.target;
    anchor.dataset.docxExternalLink = "true";
    anchor.rel = "noopener noreferrer";
  } else {
    anchor.href = "#";
    anchor.dataset.docxBookmark = hyperlink.bookmark;
  }
  return anchor;
}

function appendParagraphText(
  element: HTMLElement,
  paragraph: DocxSemanticParagraph,
): void {
  let cursor = 0;
  for (const hyperlink of paragraph.hyperlinks) {
    const index = hyperlink.label.length === 0
      ? -1
      : paragraph.text.indexOf(hyperlink.label, cursor);
    if (index < 0) continue;
    element.append(document.createTextNode(paragraph.text.slice(cursor, index)));
    element.append(hyperlinkElement(hyperlink));
    cursor = index + hyperlink.label.length;
  }
  element.append(document.createTextNode(paragraph.text.slice(cursor)));
}

function unavailablePlaceholder(
  kinds: readonly string[],
  text: string,
): HTMLElement {
  const placeholder = createSpan();
  placeholder.className = "office-viewer-docx-unavailable-content";
  placeholder.setAttribute("role", "note");
  placeholder.dataset.docxUnavailableKinds = kinds.join(",");
  placeholder.textContent = text;
  return placeholder;
}

function paragraphElement(
  paragraph: DocxSemanticParagraph,
  placeholderText: string,
): HTMLElement {
  const heading = paragraph.headingLevel;
  const tag = heading === null || heading > 6
    ? paragraph.listItem
      ? "li"
      : "p"
    : (`h${heading}` as keyof HTMLElementTagNameMap);
  const element = createEl(tag);
  appendParagraphText(element, paragraph);
  element.dataset.docxParagraphOrdinal = String(paragraph.ordinal);
  if (paragraph.listItem) {
    element.classList.add("office-viewer-docx-bounded-list-item");
  }
  if (paragraph.tableDepth > 0) {
    element.classList.add("office-viewer-docx-bounded-table-paragraph");
    element.setAttribute("role", "cell");
  }
  markDocxHeading(element, heading);
  for (const kind of paragraph.unavailableContent) {
    element.append(unavailablePlaceholder([kind], placeholderText));
  }
  for (let index = 0; index < paragraph.inlineImageCount; index += 1) {
    element.append(unavailablePlaceholder(["inline-image"], placeholderText));
  }
  return element;
}

export class BoundedDocxRendererAdapter implements DocxRendererAdapter {
  private readonly options: BoundedDocxRendererOptions;

  constructor(
    private readonly delegate: DocxRendererAdapter,
    options: Partial<BoundedDocxRendererOptions> = {},
  ) {
    this.options = {
      windowSize: positiveInteger(
        options.windowSize ?? DEFAULT_OPTIONS.windowSize,
        "windowSize",
      ),
      unavailablePlaceholder:
        options.unavailablePlaceholder ??
        DEFAULT_OPTIONS.unavailablePlaceholder,
    };
  }

  async open(
    buffer: ArrayBuffer,
    container: HTMLElement,
    model: DocxSemanticModel,
    signal: AbortSignal,
    mode: DocxRenderMode = "full",
  ): Promise<DocxRendererSession> {
    return mode === "simplified"
      ? this.openBoundedSemantic(container, model, signal)
      : this.delegate.open(buffer, container, model, signal);
  }

  private async openBoundedSemantic(
    container: HTMLElement,
    model: DocxSemanticModel,
    signal: AbortSignal,
  ): Promise<DocxRendererSession> {
    signal.throwIfAborted();

    const root = createDiv();
    root.className =
      "office-viewer-docx office-viewer-docx--bounded-semantic";
    const topSpacer = createDiv();
    topSpacer.className = "office-viewer-docx-bounded-spacer";
    topSpacer.setAttribute("aria-hidden", "true");
    const window = createDiv();
    window.className = "office-viewer-docx-bounded-window";
    const bottomSpacer = createDiv();
    bottomSpacer.className = "office-viewer-docx-bounded-spacer";
    bottomSpacer.setAttribute("aria-hidden", "true");
    root.append(topSpacer, window, bottomSpacer);

    const mounted = new Map<number, HTMLElement>();
    const bodyBlocks = new Map<number, readonly string[]>();
    for (const block of model.unavailableBodyBlocks) {
      bodyBlocks.set(block.afterParagraphOrdinal, [
        ...(bodyBlocks.get(block.afterParagraphOrdinal) ?? []),
        ...block.kinds,
      ]);
    }
    const count = model.paragraphs.length;
    const size = Math.min(this.options.windowSize, count);
    const estimatedParagraphHeight = 36;
    let zoom = 1;
    let start = -1;

    const renderWindow = (targetOrdinal: number): HTMLElement | null => {
      const targetIndex = Math.max(
        0,
        Math.min(count - 1, targetOrdinal - 1),
      );
      const nextStart = Math.max(
        0,
        Math.min(count - size, targetIndex - Math.floor(size / 2)),
      );
      if (nextStart !== start) {
        start = nextStart;
        mounted.clear();
        const fragment = createFragment();
        for (let index = start; index < start + size; index += 1) {
          const paragraph = model.paragraphs[index];
          if (paragraph === undefined) continue;
          if (paragraph.ordinal === 1) {
            const leadingKinds = bodyBlocks.get(0);
            if (leadingKinds !== undefined) {
              fragment.append(
                unavailablePlaceholder(
                  leadingKinds,
                  this.options.unavailablePlaceholder,
                ),
              );
            }
          }
          const element = paragraphElement(
            paragraph,
            this.options.unavailablePlaceholder,
          );
          mounted.set(paragraph.ordinal, element);
          fragment.append(element);
          const trailingKinds = bodyBlocks.get(paragraph.ordinal);
          if (trailingKinds !== undefined) {
            fragment.append(
              unavailablePlaceholder(
                trailingKinds,
                this.options.unavailablePlaceholder,
              ),
            );
          }
        }
        window.replaceChildren(fragment);
        topSpacer.style.height =
          `${start * estimatedParagraphHeight}px`;
        bottomSpacer.style.height =
          `${Math.max(0, count - start - size) * estimatedParagraphHeight}px`;
      }
      return mounted.get(targetOrdinal) ?? null;
    };

    const onScroll = (): void => {
      const approximateOrdinal =
        Math.floor(container.scrollTop / (estimatedParagraphHeight * zoom)) + 1;
      if (
        approximateOrdinal < start + Math.floor(size / 4) ||
        approximateOrdinal > start + Math.floor((size * 3) / 4)
      ) {
        renderWindow(approximateOrdinal);
      }
    };

    renderWindow(1);
    signal.throwIfAborted();
    container.replaceChildren(root);
    const zoomControl = createDocxZoom(container, () => root, (scale) => {
      zoom = scale;
      root.style.setProperty("zoom", String(zoom));
    });
    container.addEventListener("scroll", onScroll, { passive: true });

    return {
      candidate: "bounded-semantic",
      paragraphElements: mounted,
      warnings: [],
      managesUnavailableContent: true,
      zoom: zoomControl,
      revealParagraph: (ordinal) => {
        if (!Number.isSafeInteger(ordinal) || ordinal < 1 || ordinal > count) {
          return null;
        }
        const element = renderWindow(ordinal);
        container.scrollTop = (ordinal - 1) * estimatedParagraphHeight * zoom;
        return element;
      },
      dispose: () => {
        zoomControl.dispose();
        container.removeEventListener("scroll", onScroll);
        mounted.clear();
        root.replaceChildren();
        if (root.parentElement === container) container.replaceChildren();
      },
    };
  }
}
