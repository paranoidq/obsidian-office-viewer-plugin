import { readFile } from "node:fs/promises";
import path from "node:path";
import { TFile } from "obsidian";
import { describe, expect, it, vi } from "vitest";

import { DocxFileView } from "../../src/docx/docx-file-view";
import { createDocxMessageTranslator } from "../../src/docx/docx-messages";
import { BoundedDocxRendererAdapter } from "../../src/docx/renderer/bounded-docx-renderer-adapter";
import type {
  DocxRendererAdapter,
  DocxRendererSession,
} from "../../src/docx/renderer/docx-renderer-adapter";

async function fixtureBuffer(name: string): Promise<ArrayBuffer> {
  const bytes = await readFile(
    path.resolve("tests/fixtures/docx-exploration", name),
  );
  return Uint8Array.from(bytes).buffer;
}

class SemanticTestRenderer implements DocxRendererAdapter {
  async open(
    _buffer: ArrayBuffer,
    container: HTMLElement,
    model: Parameters<DocxRendererAdapter["open"]>[2],
    signal: AbortSignal,
  ): Promise<DocxRendererSession> {
    signal.throwIfAborted();
    const mapping = new Map<number, HTMLElement>();
    const body = document.createElement("div");
    for (const paragraph of model.paragraphs) {
      const element = document.createElement("p");
      element.textContent = paragraph.text;
      element.dataset.docxParagraphOrdinal = String(paragraph.ordinal);
      mapping.set(paragraph.ordinal, element);
      body.append(element);
    }
    container.replaceChildren(body);
    return {
      candidate: "docx-preview",
      paragraphElements: mapping,
      warnings: [],
      revealParagraph: (ordinal) => mapping.get(ordinal) ?? null,
      dispose: () => body.remove(),
    };
  }
}

describe("DOCX reading and search view", () => {
  it("reads and searches a document without paragraph reference controls", async () => {
    const source = await fixtureBuffer("read-search-only.docx");
    const sourceBefore = new Uint8Array(source).slice();
    const app = {
      scope: undefined,
      vault: { readBinary: vi.fn(async () => source) },
    };
    const view = new DocxFileView(
      { app } as never,
      {
        renderer: new SemanticTestRenderer(),
        messages: createDocxMessageTranslator("en"),
      },
    );
    const file = Object.assign(new TFile(), {
      path: "docx-exploration/read-search-only.docx",
      basename: "read-search-only",
      extension: "docx",
      stat: { size: source.byteLength, mtime: 1 },
    });
    Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
      configurable: true,
      value: vi.fn(),
    });

    await view.onLoadFile(file);

    const root = view.contentEl.querySelector<HTMLElement>(
      ".office-viewer-docx-shell",
    )!;
    expect(root.dataset.state).toBe("ready");
    expect(root.dataset.renderer).toBe("docx-preview");
    expect(root.textContent).toContain("This paragraph has a native identity.");
    expect(root.textContent).toContain(
      "This generated paragraph has no native identity.",
    );
    expect(root.textContent).not.toContain("stable source identities");
    expect(
      root.querySelector('[data-action="open-docx-search"]'),
    ).not.toBeNull();
    expect(
      root.querySelector('[data-action="open-externally"]')
        ?.getAttribute("aria-label"),
    ).toBe("Open in default application");
    const brand = root.querySelector("[data-office-viewer-brand]")!;
    expect(brand).not.toBeNull();
    expect(brand.querySelector(".office-viewer-brand__product")?.textContent)
      .toBe("Office Viewer");
    expect(brand.querySelector(".office-viewer-brand__format")?.textContent)
      .toBe("DOCX");
    expect(brand.querySelector(".office-viewer-brand__creator")).toBeNull();
    expect(
      root.querySelector(".office-viewer-toolbar__primary")
        ?.previousElementSibling,
    ).toBe(brand);

    root
      .querySelector<HTMLButtonElement>('[data-action="open-docx-search"]')!
      .click();
    expect(root.dataset.searchOpen).toBe("true");
    const search = root.querySelector<HTMLInputElement>(
      '[data-action="docx-search-input"]',
    )!;
    search.value = "generated paragraph";
    search.dispatchEvent(new Event("input"));
    const result = root.querySelector<HTMLButtonElement>(
      ".office-viewer-docx-search-result",
    )!;
    expect(result.getAttribute("aria-label")).toContain(
      "Paragraph 2, matches: 1",
    );
    expect(result.textContent).toContain("generated paragraph");
    result.click();
    expect(
      root
        .querySelector('[data-docx-paragraph-ordinal="2"]')
        ?.classList.contains("is-active-docx-paragraph"),
    ).toBe(true);
    expect(result.getAttribute("aria-current")).toBe("location");
    expect(new Uint8Array(source)).toEqual(sourceBefore);
  });

  it("exposes a stable local error state without changing the source", async () => {
    const source = Uint8Array.from([1, 2, 3, 4]).buffer;
    const sourceBefore = new Uint8Array(source).slice();
    const app = {
      scope: undefined,
      vault: { readBinary: vi.fn(async () => source) },
    };
    const view = new DocxFileView(
      { app } as never,
      {
        renderer: new SemanticTestRenderer(),
        messages: createDocxMessageTranslator("en"),
        openInDefaultApplication: vi.fn(async () => undefined),
      },
    );
    const file = Object.assign(new TFile(), {
      path: "broken.docx",
      basename: "broken",
      extension: "docx",
      stat: { size: source.byteLength, mtime: 1 },
    });

    await view.onLoadFile(file);

    const root = view.contentEl.querySelector<HTMLElement>(
      ".office-viewer-docx-shell",
    )!;
    expect(root.dataset.state).toBe("error");
    expect(root.dataset.errorCategory).toBe("malformed");
    expect(root.querySelector(".office-viewer-error")).not.toBeNull();
    expect(root.querySelector('[data-action="retry"]')?.textContent).toBe(
      "Retry",
    );
    expect(root.querySelector('[data-action="open-externally"]')?.getAttribute("aria-label"))
      .toBe("Open in default application");
    expect(root.querySelectorAll("[data-office-viewer-brand]")).toHaveLength(1);
    expect(
      root.querySelector(".office-viewer-brand__format")?.textContent,
    ).toBe("DOCX");
    expect(root.textContent).toContain("original DOCX file was not modified");
    expect(new Uint8Array(source)).toEqual(sourceBefore);
  });

  it("searches and reveals a distant result while keeping DOM bounded", async () => {
    const source = await fixtureBuffer("body-led-reference.docx");
    const app = {
      scope: undefined,
      vault: { readBinary: vi.fn(async () => source) },
    };
    const renderer = new BoundedDocxRendererAdapter(
      new SemanticTestRenderer(),
      { windowSize: 3 },
    );
    const view = new DocxFileView(
      { app } as never,
      {
        renderer,
        messages: createDocxMessageTranslator("en"),
      },
    );
    const file = Object.assign(new TFile(), {
      path: "docx-exploration/body-led-reference.docx",
      basename: "body-led-reference",
      extension: "docx",
      stat: { size: source.byteLength, mtime: 1 },
    });
    Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
      configurable: true,
      value: vi.fn(),
    });

    await view.onLoadFile(file);
    const root = view.contentEl.querySelector<HTMLElement>(
      ".office-viewer-docx-shell",
    )!;
    expect(root.dataset.renderer).toBe("docx-preview");
    const toggle = root.querySelector<HTMLButtonElement>('[data-action="toggle-docx-mode"]')!;
    expect(toggle.getAttribute("aria-pressed")).toBe("false");
    toggle.click();
    await vi.waitFor(() => expect(root.dataset.renderer).toBe("bounded-semantic"));
    expect(root.textContent).toContain("Showing simplified view.");
    expect(
      root.querySelectorAll(".office-viewer-docx-reading-body *").length,
    ).toBeLessThanOrEqual(20);

    root
      .querySelector<HTMLButtonElement>('[data-action="open-docx-search"]')!
      .click();
    const search = root.querySelector<HTMLInputElement>(
      '[data-action="docx-search-input"]',
    )!;
    search.value = "中文段落";
    search.dispatchEvent(new Event("input"));
    root
      .querySelector<HTMLButtonElement>(".office-viewer-docx-search-result")!
      .click();

    const active = root.querySelector<HTMLElement>(
      '[data-docx-paragraph-ordinal="11"]',
    );
    expect(active?.classList.contains("is-active-docx-paragraph")).toBe(true);
    expect(
      root.querySelectorAll(".office-viewer-docx-reading-body *").length,
    ).toBeLessThanOrEqual(20);
    toggle.click();
    await vi.waitFor(() => expect(root.dataset.renderer).toBe("docx-preview"));
    expect(root.querySelector('[data-docx-paragraph-ordinal="11"]')?.getAttribute("aria-current"))
      .toBe("true");
    expect(search.value).toBe("中文段落");
    expect(app.vault.readBinary).toHaveBeenCalledOnce();
    view.dispose();
  });
});
