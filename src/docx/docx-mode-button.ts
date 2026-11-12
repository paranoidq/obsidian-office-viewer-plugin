import { decorateOfficeViewerIconButton } from "../office-viewer-chrome";
import type { DocxMessageTranslator } from "./docx-messages";
import type { DocxRenderMode } from "./renderer/docx-renderer-adapter";

export function createDocxModeButton(
  document: Document,
  messages: DocxMessageTranslator,
  onToggle: (mode: DocxRenderMode) => void,
) {
  const button = document.createElement("button");
  button.type = "button";
  button.dataset.action = "toggle-docx-mode";
  button.setAttribute("aria-label", messages.text("simplifiedView"));
  button.disabled = true;
  decorateOfficeViewerIconButton(button, "lucide-align-left");
  let mode: DocxRenderMode = "full";
  const setMode = (next: DocxRenderMode) => {
    mode = next;
    button.setAttribute("aria-pressed", String(mode === "simplified"));
    button.title = messages.text(mode === "full" ? "switchToSimplified" : "switchToFull");
  };
  setMode(mode);
  button.addEventListener("click", (event) => {
    event.stopPropagation();
    onToggle(mode === "full" ? "simplified" : "full");
  });
  return { button, setMode };
}
