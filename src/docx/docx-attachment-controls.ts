import { setIcon } from "obsidian";
import type { DocxAttachment, DocxAttachmentFile } from "./docx-attachments";
import { UnsupportedDocxAttachment } from "./docx-attachments";
import type { DocxMessageTranslator } from "./docx-messages";
import type { DocxRenderSource } from "./docx-render-source";

export function createDocxAttachmentButton(
  document: Document,
  attachment: DocxAttachment,
  messages: DocxMessageTranslator,
  preview?: HTMLElement,
): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "office-viewer-docx-attachment";
  button.dataset.docxAttachment = String(attachment.objectIndex);
  const extension = attachment.extension;
  const name = attachment.name ?? messages.text("attachmentName", { number: attachment.objectIndex + 1 });
  const label = attachment.name === null && extension !== null ? `${name} (${extension.toUpperCase()})` : name;
  const icon = document.createElement("span");
  icon.setAttribute("aria-hidden", "true");
  setIcon(icon, extension === "xlsx" || extension === "xls" ? "lucide-sheet"
    : extension === "pptx" || extension === "ppt" ? "lucide-presentation" : "lucide-file-text");
  const title = document.createElement("span");
  title.textContent = label;
  if (preview === undefined) button.append(icon, title);
  else {
    button.classList.add("office-viewer-docx-attachment--preview");
    if (attachment.centerPreview) button.classList.add("office-viewer-docx-attachment--centered");
    button.append(preview);
  }
  button.title = messages.text("attachmentOpen", { name: label });
  button.setAttribute("aria-label", button.title);
  return button;
}

export const DOCX_ATTACHMENT_STYLES = `
  button.office-viewer-docx-attachment {
    align-items: center;
    background: var(--background-secondary, #f2f4f7);
    border: 1px solid var(--background-modifier-border, #ced3dc);
    border-radius: var(--radius-s, 4px);
    color: var(--text-normal, #21252b);
    cursor: pointer;
    display: inline-flex;
    font: var(--font-ui-small, 13px) var(--font-interface, sans-serif);
    gap: 6px;
    margin: 2px 4px;
    max-width: 95%;
    padding: 4px 8px;
    text-align: start;
    vertical-align: middle;
  }
  .office-viewer-docx-attachment svg { width: 16px; height: 16px; }
  .office-viewer-docx-attachment > span:last-child { overflow-wrap: anywhere; }
  button.office-viewer-docx-attachment--preview {
    background: transparent;
    border-color: transparent;
    margin: 0;
    padding: 0;
  }
  button.office-viewer-docx-attachment--preview:is(:hover, :focus-visible) {
    border-color: var(--interactive-accent, #595ebf);
  }
  .office-viewer-docx-attachment--preview img { pointer-events: none; }
  button.office-viewer-docx-attachment--centered {
    display: flex;
    inline-size: fit-content;
    margin-inline: auto;
  }
  button.office-viewer-docx-attachment[aria-busy="true"] { cursor: progress; }
`;

export function registerDocxAttachmentOpening(
  viewport: HTMLElement,
  getSource: () => DocxRenderSource | null | undefined,
  open: ((file: DocxAttachmentFile) => Promise<void>) | undefined,
  messages: DocxMessageTranslator,
  status: HTMLElement,
): () => void {
  const pending = new Set<HTMLButtonElement>();
  const controller = new AbortController();
  const handle = (event: MouseEvent) => {
    const Button = viewport.ownerDocument.defaultView?.HTMLButtonElement;
    if (Button === undefined) return;
    const button = event.composedPath().find((target): target is HTMLButtonElement =>
      target instanceof Button && target.hasAttribute("data-docx-attachment"),
    );
    if (button === undefined) return;
    event.stopPropagation();
    if (event.button !== 0 || event.detail > 1 || pending.has(button)) return;
    const source = getSource();
    if (source == null || open === undefined) {
      status.textContent = messages.text("attachmentOpenFailed");
      return;
    }
    pending.add(button);
    button.setAttribute("aria-busy", "true");
    status.textContent = "";
    void (async () => {
      try {
        const file = await source.readAttachment(Number(button.dataset.docxAttachment));
        if (!controller.signal.aborted && source === getSource()) await open(file);
      } catch (error) {
        if (!controller.signal.aborted && source === getSource()) {
          status.textContent = messages.text(error instanceof UnsupportedDocxAttachment
            ? "attachmentUnsupported" : "attachmentOpenFailed");
        }
      } finally {
        pending.delete(button);
        button.removeAttribute("aria-busy");
      }
    })();
  };
  viewport.addEventListener("click", handle);
  return () => {
    controller.abort();
    viewport.removeEventListener("click", handle);
  };
}
