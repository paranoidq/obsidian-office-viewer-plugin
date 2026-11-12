import { decorateOfficeViewerIconButton } from "../office-viewer-chrome";
import type { DocxMessageTranslator } from "./docx-messages";
import type { DocxSemanticModel } from "./docx-semantic-model";

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
    for (const paragraph of model.paragraphs) {
      const level = paragraph.headingLevel;
      if (level === null || paragraph.tableDepth > 0) continue;
      const item = document.createElement("li");
      const entry = document.createElement("button");
      entry.type = "button";
      entry.className = "office-viewer-docx-outline__entry";
      entry.dataset.action = "docx-outline-entry";
      entry.dataset.docxHeadingLevel = String(level);
      entry.dataset.paragraphOrdinal = String(paragraph.ordinal);
      entry.style.setProperty("--office-viewer-docx-outline-indent", `${(level - 1) * 8}px`);
      const title = paragraph.text.replace(/\s+/g, " ").trim();
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
      entries.append(item);
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
