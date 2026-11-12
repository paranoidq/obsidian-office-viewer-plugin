type ClipboardWriter = (text: string) => void;

function writeDesktopClipboard(text: string): void {
  // Desktop Obsidian exposes the system clipboard from its renderer process.
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { clipboard } = require("electron") as {
    clipboard: { writeText(text: string): void };
  };
  clipboard.writeText(text);
}

function previewFor<T>(node: Node | null, previews: ReadonlyMap<Node, T>): T | undefined {
  while (node !== null) {
    const preview = previews.get(node);
    if (preview !== undefined) return preview;
    // DOCX pages are rendered inside a shadow root to isolate their styles.
    node = node.parentNode ?? (node.nodeType === 11 && "host" in node
      ? (node as ShadowRoot).host
      : null);
  }
  return undefined;
}

const documents = new WeakMap<Document, {
  previews: Map<Node, { write: ClipboardWriter }>;
  detach: () => void;
}>();

export function registerOfficeEmbedTextCopy(
  preview: HTMLElement,
  write: ClipboardWriter = writeDesktopClipboard,
): () => void {
  const document = preview.ownerDocument;
  const window = document.defaultView;
  if (window === null) return () => {};
  let entry = documents.get(document);
  if (entry === undefined) {
    const previews = new Map<Node, { write: ClipboardWriter }>();
    const selectionForCopy = (event: Event) => {
      if (event.composedPath().some((target) =>
        target instanceof window.HTMLInputElement || target instanceof window.HTMLTextAreaElement
          || target instanceof window.HTMLSelectElement,
      )) return null;
      const selection = document.getSelection();
      if (selection === null || selection.isCollapsed) return null;
      const preview = previewFor(selection.anchorNode, previews);
      if (preview === undefined || previewFor(selection.focusNode, previews) !== preview) return null;
      const text = selection.toString();
      return text.length === 0 ? null : { text, write: preview.write };
    };
    const onCopy = (event: ClipboardEvent) => {
      if (event.clipboardData == null) return;
      const selected = selectionForCopy(event);
      if (selected === null) return;
      event.clipboardData.setData("text/plain", selected.text);
      event.preventDefault();
      event.stopImmediatePropagation();
    };
    const onKeydown = (event: KeyboardEvent) => {
      if (event.key.toLowerCase() !== "c" || (!event.metaKey && !event.ctrlKey)
        || event.altKey || event.shiftKey || event.isComposing) return;
      const selected = selectionForCopy(event);
      if (selected === null) return;
      try {
        selected.write(selected.text);
      } catch {
        // Leave the browser's copy action available if the desktop API fails.
        return;
      }
      event.preventDefault();
      event.stopImmediatePropagation();
    };
    // Live Preview can retain editor focus while the selection is in a widget.
    // Handle the shortcut before CodeMirror can use its Markdown selection.
    window.addEventListener("keydown", onKeydown, true);
    window.addEventListener("copy", onCopy, true);
    entry = {
      previews,
      detach: () => {
        window.removeEventListener("keydown", onKeydown, true);
        window.removeEventListener("copy", onCopy, true);
      },
    };
    documents.set(document, entry);
  }

  const registration = entry;
  registration.previews.set(preview, { write });
  return () => {
    if (!registration.previews.delete(preview)) return;
    if (registration.previews.size === 0) {
      registration.detach();
      documents.delete(document);
    }
  };
}
