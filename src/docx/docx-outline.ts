import { decorateOfficeViewerIconButton } from "../office-viewer-chrome";
import type { DocxMessageTranslator } from "./docx-messages";
import type { DocxSemanticModel } from "./docx-semantic-model";

function numberedOutlineTitle(text: string, headingNumber?: string): string {
  const title = text.replace(/\s+/g, " ").trim();
  const number = headingNumber?.trim();
  if (!number) return title;
  const punctuation = /[、：．]$/u.test(number);
  const alreadyNumbered = title.startsWith(number) &&
    (title.length === number.length || punctuation || /^\s/.test(title.slice(number.length)));
  if (alreadyNumbered) return title;
  return punctuation ? `${number}${title}` : `${number} ${title}`;
}

export function createDocxOutline(
  document: Document,
  messages: DocxMessageTranslator,
  onNavigate: (ordinal: number) => boolean,
  onOpen?: () => void,
) {
  const button = document.createElement("button");
  button.type = "button";
  button.dataset.action = "toggle-docx-outline";
  button.setAttribute("aria-label", messages.text("outline"));
  decorateOfficeViewerIconButton(button, "lucide-list-tree");

  const panel = document.createElement("nav");
  panel.className = "office-viewer-docx-outline";
  panel.setAttribute("aria-label", messages.text("outline"));
  const label = document.createElement("div");
  label.className = "office-viewer-docx-outline__title";
  label.textContent = messages.text("outline");
  const list = document.createElement("ul");
  list.className = "office-viewer-docx-outline__list";
  const notice = document.createElement("div");
  notice.className = "office-viewer-docx-outline__notice";
  notice.setAttribute("role", "status");
  panel.append(label, list, notice);

  let model: DocxSemanticModel | undefined;
  let current: HTMLButtonElement | undefined;
  let populated = false;

  const setOpen = (open: boolean) => {
    panel.hidden = !open;
    button.setAttribute("aria-pressed", String(open));
    button.title = messages.text(open ? "outlineClose" : "outlineOpen");
  };

  const populate = () => {
    if (populated || model === undefined) return;
    populated = true;
    const entries = document.createDocumentFragment();
    const ancestors: Array<{
      level: number;
      item: HTMLLIElement;
      children?: HTMLUListElement;
    }> = [];
    for (const paragraph of model.paragraphs) {
      const level = paragraph.headingLevel;
      if (level === null || paragraph.tableDepth > 0) continue;
      let parent = ancestors.at(-1);
      while (parent !== undefined && parent.level >= level) {
        ancestors.pop();
        parent = ancestors.at(-1);
      }
      const item = document.createElement("li");
      const entry = document.createElement("button");
      entry.type = "button";
      entry.className = "office-viewer-docx-outline__entry";
      entry.dataset.action = "docx-outline-entry";
      entry.dataset.docxHeadingLevel = String(level);
      entry.dataset.paragraphOrdinal = String(paragraph.ordinal);
      const title = numberedOutlineTitle(paragraph.text, paragraph.headingNumber);
      entry.textContent = title;
      entry.setAttribute("aria-label", messages.text("outlineEntry", { level, title }));
      entry.addEventListener("click", (event) => {
        event.stopPropagation();
        if (!onNavigate(paragraph.ordinal)) {
          notice.textContent = messages.text("outlineUnavailable");
          return;
        }
        notice.textContent = "";
        current?.removeAttribute("aria-current");
        current = entry;
        entry.setAttribute("aria-current", "location");
      });
      item.append(entry);
      if (parent === undefined) {
        entries.append(item);
      } else {
        if (parent.children === undefined) {
          parent.children = document.createElement("ul");
          parent.children.className = "office-viewer-docx-outline__branch";
          parent.item.append(parent.children);
        }
        parent.children.append(item);
      }
      ancestors.push({ level, item });
    }
    notice.textContent = entries.childElementCount === 0 ? messages.text("outlineEmpty") : "";
    list.append(entries);
  };

  button.addEventListener("click", (event) => {
    event.stopPropagation();
    if (panel.hidden) {
      onOpen?.();
      populate();
      setOpen(true);
    } else {
      setOpen(false);
    }
  });
  panel.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    event.stopPropagation();
    setOpen(false);
    button.focus();
  });
  setOpen(false);
  button.disabled = true;

  return {
    button,
    panel,
    close: () => setOpen(false),
    setModel(next?: DocxSemanticModel) {
      if (model === next) return;
      model = next;
      current = undefined;
      populated = false;
      list.replaceChildren();
      notice.textContent = "";
      setOpen(false);
      button.disabled = next === undefined;
    },
  };
}

export function scrollToDocxHeading(viewport: HTMLElement, element: HTMLElement): void {
  viewport.scrollTop += element.getBoundingClientRect().top
    - viewport.getBoundingClientRect().top - viewport.clientTop - 8;
}
