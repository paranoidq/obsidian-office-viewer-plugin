import type { DocxTocEntry } from "../docx-semantic-model";

export const DOCX_TOC_STYLES = `
  .docx p[data-docx-toc-level] {
    align-items: baseline;
    display: grid;
    grid-template-columns: minmax(0, auto) minmax(0.75em, 1fr) max-content;
    column-gap: 0.28em;
    padding-inline-start: var(--office-viewer-docx-toc-inset);
    text-indent: 0;
  }
  .docx p[data-docx-toc-level="1"] { font-weight: 600; }
  .office-viewer-docx-toc-title { overflow-wrap: anywhere; min-width: 0; }
  .office-viewer-docx-toc-leader {
    border-block-end: 0.075em dotted currentColor;
    align-self: baseline;
    line-height: inherit;
  }
  .office-viewer-docx-toc-leader[data-leader="dash"] { border-block-end-style: dashed; }
  .office-viewer-docx-toc-leader[data-leader="solid"] { border-block-end-style: solid; }
  .office-viewer-docx-toc-leader[data-leader="none"] { border-block-end-style: none; }
  .office-viewer-docx-toc-leader::before { content: "\\00a0"; }
  .docx .office-viewer-docx-toc-page { white-space: nowrap; font-variant-numeric: tabular-nums; }
  .docx a[data-docx-toc-target] { cursor: pointer; }
  .docx a[data-docx-toc-target]:is(:hover, :focus-visible) { text-decoration: underline; }
`;

export function layoutDocxTocEntry(paragraph: HTMLElement, entry: DocxTocEntry): void {
  const document = paragraph.ownerDocument;
  const range = document.createRange();
  // docx-preview renders w:tab as a single em space inside its own span.
  const tabs = Array.from(paragraph.querySelectorAll("span")).reverse();
  const separator = tabs.find((tab) => {
    if (tab.childElementCount !== 0 || tab.textContent !== "\u2003") return false;
    range.setStartAfter(tab);
    range.setEnd(paragraph, paragraph.childNodes.length);
    return range.toString().trim() === entry.page;
  });
  if (separator === undefined) return;

  const style = document.defaultView?.getComputedStyle(paragraph);
  const fontSize = parseFloat(style?.fontSize ?? "") || 16;
  const margin = parseFloat(style?.marginInlineStart ?? "") || 0;
  const firstLine = parseFloat(style?.textIndent ?? "") || 0;
  const inset = Math.max(0, (entry.level - 1) * fontSize * 1.15 - margin, firstLine);
  const page = document.createElement("span");
  page.className = "office-viewer-docx-toc-page";
  page.append(range.extractContents());
  separator.remove();
  const title = document.createElement("span");
  title.className = "office-viewer-docx-toc-title";
  title.append(...paragraph.childNodes);
  if (entry.number !== null) {
    const number = document.createElement("span");
    number.className = "office-viewer-docx-toc-number";
    number.textContent = `${entry.number} `;
    (title.querySelector("a[data-docx-bookmark]") ?? title).prepend(number);
  }
  if (entry.targetOrdinal !== null) {
    for (const part of [title, page]) {
      for (const link of part.querySelectorAll<HTMLAnchorElement>("a[data-docx-bookmark]")) {
        if (link.dataset.docxBookmark !== entry.bookmark) continue;
        link.dataset.docxTocTarget = String(entry.targetOrdinal);
        link.href = "#";
      }
    }
  }
  const leader = document.createElement("span");
  leader.className = "office-viewer-docx-toc-leader";
  leader.dataset.leader = entry.leader;
  leader.setAttribute("aria-hidden", "true");
  paragraph.dataset.docxTocLevel = String(entry.level);
  paragraph.style.setProperty("--office-viewer-docx-toc-inset", `${inset}px`);
  paragraph.style.textIndent = "0";
  paragraph.append(title, leader, page);
}
