import { expect, it, vi } from "vitest";
import { registerOfficeEmbedTextCopy } from "../src/office-embed-text-copy";

it("copies a DOCX shadow selection even while the Markdown editor owns keyboard focus", () => {
  const editor = document.createElement("div");
  const noteText = document.createTextNode("note text");
  const preview = document.createElement("div");
  const shadowHost = preview.appendChild(document.createElement("div"));
  const text = shadowHost.attachShadow({ mode: "open" })
    .appendChild(document.createTextNode("预览文字"));
  editor.append(noteText, preview);
  document.body.append(editor);

  const writer = vi.fn();
  const unregister = registerOfficeEmbedTextCopy(preview, writer);
  const selection = vi.spyOn(document, "getSelection");
  const setSelection = (anchorNode: Node, focusNode: Node, value: string) => {
    selection.mockReturnValue({
      anchorNode, focusNode, isCollapsed: false, toString: () => value,
    } as Selection);
  };
  const pressCopy = () => {
    const event = new KeyboardEvent("keydown", {
      key: "c", metaKey: true, bubbles: true, cancelable: true,
    });
    editor.dispatchEvent(event);
    return event.defaultPrevented;
  };

  try {
    setSelection(text, text, "预览");
    expect(pressCopy()).toBe(true);
    expect(writer).toHaveBeenCalledExactlyOnceWith("预览");

    const setData = vi.fn();
    const copy = new Event("copy", { bubbles: true, cancelable: true });
    Object.defineProperty(copy, "clipboardData", { value: { setData } });
    editor.dispatchEvent(copy);
    expect(copy.defaultPrevented).toBe(true);
    expect(setData).toHaveBeenCalledWith("text/plain", "预览");

    setSelection(noteText, noteText, "note");
    expect(pressCopy()).toBe(false);
    setSelection(noteText, text, "note preview");
    expect(pressCopy()).toBe(false);

    unregister();
    setSelection(text, text, "预览");
    expect(pressCopy()).toBe(false);
    expect(writer).toHaveBeenCalledTimes(1);
  } finally {
    unregister();
    selection.mockRestore();
    editor.remove();
  }
});
