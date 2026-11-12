import { readFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import type { MarkdownRenderChild } from "obsidian";
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { createDocxMessageTranslator } from "../src/docx/docx-messages";
import { createDocxRendererAdapter } from "../src/docx/renderer/create-docx-renderer-adapter";
import { ENGLISH_MESSAGE_TRANSLATOR } from "../src/i18n";
import {
  processOfficeFileEmbeds,
  type OfficeFileEmbedOptions,
} from "../src/office-file-embed";
import { SlideEmbedScheduler } from "../src/slide-embed-scheduler";

beforeEach(() => vi.stubGlobal("IntersectionObserver", undefined));
afterEach(() => vi.unstubAllGlobals());

function fixture(link: string, overrides: Partial<OfficeFileEmbedOptions<string>>) {
  const root = document.createElement("div");
  const native = root.createSpan({ cls: "internal-embed", attr: { src: link } });
  const children = new Set<MarkdownRenderChild>();
  const scheduler = new SlideEmbedScheduler();
  const options: OfficeFileEmbedOptions<string> = {
    resolveFile: vi.fn((sourcePath) => sourcePath),
    readBinary: vi.fn(async () => new ArrayBuffer(0)),
    openInTab: vi.fn(async () => undefined),
    openExternally: vi.fn(async () => undefined),
    openExternalUrl: vi.fn(async () => undefined),
    pptxRenderer: { open: async () => { throw new Error("Unexpected PPTX"); } },
    docxRenderer: createDocxRendererAdapter({ pageLayout: true }),
    scheduler,
    messages: ENGLISH_MESSAGE_TRANSLATOR,
    docxMessages: createDocxMessageTranslator("en"),
    ...overrides,
  };
  const process = () => processOfficeFileEmbeds(root, {
    sourcePath: "notes/note.md",
    addChild: (child: MarkdownRenderChild) => child.load(),
  } as never, options, {
    register: (child) => children.add(child),
    unregister: (child) => children.delete(child),
  });
  process();
  return { root, native, options, process, dispose: () => {
    for (const child of children) child.unload();
    scheduler.dispose();
  } };
}

it("previews a whole deck in place and opens the resolved local source only on click", async () => {
  const renderSlide = vi.fn(async (index: number) => {
    if (index === 1) throw new Error("Could not render slide 2");
  });
  const dispose = vi.fn();
  const f = fixture("slides/Quarterly Report.PPTX|Quarterly Report", {
    pptxRenderer: {
      open: async () => ({
        slideCount: 3, slideWidth: 16, slideHeight: 9,
        capabilities: { thumbnails: false, prefetch: false },
        renderSlide, dispose,
      }),
    },
  });
  try {
    f.process();
    expect(f.root.querySelectorAll(".office-file-embed")).toHaveLength(1);
    const embed = f.root.querySelector<HTMLElement>(".office-file-embed")!;
    await vi.waitFor(() => expect(embed.dataset.state).toBe("ready"));
    expect(f.options.resolveFile).toHaveBeenCalledWith("slides/Quarterly Report.PPTX", "notes/note.md");
    expect(f.options.openInTab).not.toHaveBeenCalled();
    const tabButton = embed.querySelector<HTMLButtonElement>('[data-action="open-in-tab"]')!;
    expect(tabButton.getAttribute("aria-label")).toBe("Open in new tab");
    tabButton.click();
    expect(f.options.openInTab).toHaveBeenCalledWith("slides/Quarterly Report.PPTX");
    expect(f.options.openExternally).not.toHaveBeenCalled();
    embed.querySelector<HTMLButtonElement>('[data-action="open-externally"]')!.click();
    expect(f.options.openExternally).toHaveBeenCalledWith("slides/Quarterly Report.PPTX");

    const next = embed.querySelector<HTMLButtonElement>('[data-action="next-slide"]')!;
    next.click();
    await vi.waitFor(() => expect(embed.textContent).toContain("Slide 2 could not be rendered"));
    next.click();
    await vi.waitFor(() => expect(embed.textContent).toContain("3 / 3"));
    expect(renderSlide.mock.calls).toEqual([[0], [1], [2]]);
  } finally {
    f.dispose();
  }
  expect(dispose).toHaveBeenCalledOnce();
  expect(f.root.children).toHaveLength(1);
  expect(f.native.classList.contains("office-file-embed__native")).toBe(false);
});

it("keeps DOCX page geometry, fixed table columns, and readable single spacing", async () => {
  const input = await readFile(path.resolve("tests/vault/docx-exploration/read-search-only.docx"));
  const zip = await JSZip.loadAsync(input);
  const part = "word/document.xml";
  const xml = await zip.file(part)!.async("string");
  zip.file(part, xml.replace("<w:sectPr/>", `
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    <w:p><w:r><w:t>Second page content</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblpPr w:leftFromText="180" w:rightFromText="180"
          w:horzAnchor="page" w:vertAnchor="text" w:tblpX="1200" w:tblpY="140"/>
        <w:tblW w:type="auto" w:w="0"/>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="10"/><w:gridCol w:w="9600"/><w:gridCol w:w="10"/></w:tblGrid>
      <w:tr>
        <w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>
        <w:tc><w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
          <w:r><w:t>单倍行距正文</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr>
          <w:p><w:pPr><w:spacing w:line="300" w:lineRule="exact"/></w:pPr>
            <w:r><w:t>Exact spacing</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:left="1080" w:right="1080" w:top="1440" w:bottom="1440"/>
    </w:sectPr>`));
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  const readBinary = vi.fn(async () => buffer);
  const f = fixture("attachments/报告.docx", { readBinary });
  try {
    const embed = f.root.querySelector<HTMLElement>(".office-file-embed")!;
    await vi.waitFor(() => expect(embed.dataset.state).toBe("ready"));
    const pages = embed.querySelector(".office-viewer-docx--preview")?.shadowRoot
      ?.querySelectorAll<HTMLElement>("section.docx");
    expect(pages).toHaveLength(2);
    expect(pages?.[0]?.style.width).toBe("595.3pt");
    expect(pages?.[0]?.style.paddingLeft).toBe("54pt");
    expect(pages?.[1]?.textContent).toContain("Second page content");
    const table = pages?.[1]?.querySelector("table");
    expect(table?.style.tableLayout).toBe("fixed");
    expect(table?.style.width).toBe("481pt");
    expect(table?.style.marginLeft).toBe("6pt");
    expect(table?.style.marginTop).toBe("7pt");
    expect(table?.rows[0]?.cells[1]?.querySelector("p")?.style.lineHeight).toBe("normal");
    expect(table?.rows[1]?.cells[0]?.querySelector("p")?.style.lineHeight).toBe("15pt");
    const toggle = embed.querySelector<HTMLButtonElement>('[data-action="toggle-docx-mode"]')!;
    expect(toggle.getAttribute("aria-pressed")).toBe("false");
    toggle.click();
    await vi.waitFor(() => expect(embed.dataset.renderer).toBe("bounded-semantic"));
    expect(embed.textContent).toContain("Second page content");
    expect(toggle.getAttribute("aria-pressed")).toBe("true");
    toggle.click();
    await vi.waitFor(() => expect(embed.dataset.renderer).toBe("docx-preview"));
    expect(toggle.getAttribute("aria-pressed")).toBe("false");
    expect(readBinary).toHaveBeenCalledOnce();
  } finally {
    f.dispose();
  }
});

it("leaves fragment references to their existing embed handlers", () => {
  const f = fixture("deck.pptx#slide-id=256&slide=1", {});
  try {
    expect(f.root.querySelector(".office-file-embed")).toBeNull();
    expect(f.options.resolveFile).not.toHaveBeenCalled();
  } finally {
    f.dispose();
  }
});
