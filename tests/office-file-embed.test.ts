import { readFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import CFB from "cfb";
import type { MarkdownRenderChild } from "obsidian";
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { createDocxMessageTranslator } from "../src/docx/docx-messages";
import type { DocxAttachmentFile } from "../src/docx/docx-attachments";
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

it("compacts DOCX margins while preserving page width, breaks, tables, and spacing", async () => {
  const input = await readFile(path.resolve("tests/vault/docx-exploration/read-search-only.docx"));
  const zip = await JSZip.loadAsync(input);
  const part = "word/document.xml";
  const xml = await zip.file(part)!.async("string");
  zip.file(part, xml.replace("<w:sectPr/>", `
    <w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>First section</w:t></w:r></w:p>
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    <w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>Subsection</w:t></w:r></w:p>
    <w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Second page content</w:t></w:r></w:p>
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
    <w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Next section</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:left="1080" w:right="1080" w:top="240" w:bottom="1440"/>
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
    expect(pages?.[0]?.style.paddingLeft).toBe("24pt");
    expect(pages?.[0]?.style.paddingRight).toBe("24pt");
    expect(pages?.[0]?.style.paddingTop).toBe("12pt");
    expect(pages?.[0]?.style.paddingBottom).toBe("24pt");
    expect(pages?.[0]?.style.minHeight).toBe("793.9pt");
    expect(pages?.[1]?.textContent).toContain("Second page content");
    const heading = pages?.[1]?.querySelector<HTMLElement>('[role="heading"][aria-level="3"]')!;
    expect(heading.textContent).toBe("Second page content");
    embed.querySelector<HTMLButtonElement>('[data-action="toggle-docx-outline"]')!.click();
    const list = embed.querySelector(".office-viewer-docx-outline__list")!;
    const roots = Array.from(list.children);
    expect(roots.map((root) => root.firstElementChild?.textContent))
      .toEqual(["First section", "Next section"]);
    const subsection = roots[0]?.lastElementChild?.firstElementChild;
    expect(subsection?.firstElementChild?.textContent).toBe("Subsection");
    const outlineEntry = subsection?.lastElementChild?.firstElementChild
      ?.firstElementChild as HTMLButtonElement;
    expect(outlineEntry.textContent).toBe("Second page content");
    expect(outlineEntry.dataset.docxHeadingLevel).toBe("3");
    outlineEntry.click();
    expect(outlineEntry.getAttribute("aria-current")).toBe("location");
    const table = pages?.[1]?.querySelector("table");
    expect(table?.style.tableLayout).toBe("fixed");
    expect(table?.style.width).toBe("481pt");
    expect(table?.style.marginLeft).toBe("6pt");
    expect(table?.style.marginTop).toBe("7pt");
    expect(table?.rows[0]?.cells[1]?.querySelector("p")?.style.lineHeight).toBe("normal");
    expect(table?.rows[1]?.cells[0]?.querySelector("p")?.style.lineHeight).toBe("15pt");
    const ratio = embed.querySelector<HTMLButtonElement>('[data-action="docx-zoom-reset"]')!;
    expect(ratio.textContent).toBe("100%");
    embed.querySelector<HTMLButtonElement>('[data-action="docx-zoom-in"]')!.click();
    expect(ratio.textContent).toBe("110%");
    pages?.[0]?.dispatchEvent(new WheelEvent("wheel", {
      ctrlKey: true, deltaY: -100, bubbles: true, composed: true, cancelable: true,
    }));
    expect(ratio.textContent).toBe("134%");
    const toggle = embed.querySelector<HTMLButtonElement>('[data-action="toggle-docx-mode"]')!;
    expect(toggle.getAttribute("aria-pressed")).toBe("false");
    toggle.click();
    await vi.waitFor(() => expect(embed.dataset.renderer).toBe("bounded-semantic"));
    expect(embed.textContent).toContain("Second page content");
    expect(embed.querySelector('[role="heading"][aria-level="3"]')?.textContent)
      .toBe("Second page content");
    expect(toggle.getAttribute("aria-pressed")).toBe("true");
    expect(ratio.textContent).toBe("134%");
    ratio.click();
    expect(ratio.textContent).toBe("100%");
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

it("opens DOCX attachments only on activation and keeps unsupported objects inert", async () => {
  const nestedDocx = await readFile("tests/vault/docx-exploration/read-search-only.docx");
  const source = await JSZip.loadAsync(nestedDocx);
  const part = "word/document.xml";
  const xml = await source.file(part)!.async("string");
  const object = (id: string) => `<w:p><w:r><w:object><o:OLEObject
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    Type="Embed" r:id="${id}"/></w:object></w:r></w:p>`;
  source.file(part, xml.replace("<w:sectPr/>", `${object("pdf")}${object("word")}${object("unknown")}<w:sectPr/>`));
  const relPath = "word/_rels/document.xml.rels";
  const rels = await source.file(relPath)!.async("string");
  const rel = (id: string, type: string, target: string) =>
    `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="embeddings/${target}"/>`;
  source.file(relPath, rels.replace("</Relationships>", [
    rel("pdf", "oleObject", "oleObject1.bin"),
    rel("word", "package", "report.docx"),
    rel("unknown", "oleObject", "oleObject2.bin"),
    "</Relationships>",
  ].join("")));
  const pdf = new TextEncoder().encode("%PDF-1.4\n%%EOF\n");
  const packager = CFB.utils.cfb_new();
  const metadata = new TextEncoder().encode("attachment.pdf\0C:\\attachment.pdf\0");
  const record = new Uint8Array(6 + metadata.length + 12 + pdf.length);
  const data = new DataView(record.buffer);
  data.setUint32(0, record.length - 4, true);
  data.setUint16(4, 2, true);
  record.set(metadata, 6);
  data.setUint32(6 + metadata.length + 8, pdf.length, true);
  record.set(pdf, record.length - pdf.length);
  CFB.utils.cfb_add(packager, "\x01Ole10Native", record);
  source.file("word/embeddings/oleObject1.bin", CFB.write(packager, { type: "buffer" }));
  source.file("word/embeddings/report.docx", nestedDocx);
  source.file("word/embeddings/oleObject2.bin", "Unsupported attachment");
  const buffer = await source.generateAsync({ type: "arraybuffer" });
  const openAttachment = vi.fn(async (_file: DocxAttachmentFile) => undefined);
  const f = fixture("with-attachments.docx", { readBinary: async () => buffer, openAttachment });
  const click = (element: Element, detail = 1) =>
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, composed: true, detail }));
  try {
    const embed = f.root.querySelector<HTMLElement>(".office-file-embed")!;
    await vi.waitFor(() => expect(embed.dataset.state).toBe("ready"));
    const page = embed.querySelector(".office-viewer-docx--paged")?.shadowRoot;
    const buttons = page!.querySelectorAll<HTMLButtonElement>("[data-docx-attachment]");
    expect(buttons).toHaveLength(3);
    expect(openAttachment).not.toHaveBeenCalled();
    click(buttons[0]!);
    await vi.waitFor(() => expect(openAttachment).toHaveBeenCalledTimes(1));
    click(buttons[0]!, 2);
    expect(openAttachment.mock.lastCall?.[0]).toMatchObject({ name: "attachment.pdf", extension: "pdf" });
    expect(Array.from(openAttachment.mock.lastCall![0].bytes)).toEqual(Array.from(pdf));
    click(buttons[2]!);
    await vi.waitFor(() => expect(embed.textContent).toContain("does not contain a supported document"));
    expect(openAttachment).toHaveBeenCalledTimes(1);

    embed.querySelector<HTMLButtonElement>('[data-action="toggle-docx-mode"]')!.click();
    await vi.waitFor(() => expect(embed.dataset.renderer).toBe("bounded-semantic"));
    const simplified = embed.querySelectorAll<HTMLButtonElement>("[data-docx-attachment]");
    expect(simplified).toHaveLength(3);
    simplified[1]!.click();
    await vi.waitFor(() => expect(openAttachment).toHaveBeenCalledTimes(2));
    expect(openAttachment.mock.lastCall?.[0]).toMatchObject({ name: "report.docx", extension: "docx" });
    expect(Array.from(openAttachment.mock.lastCall![0].bytes)).toEqual(Array.from(nestedDocx));
  } finally {
    f.dispose();
  }
});

it("lays out the body table of contents and jumps to the matching heading inside the embed", async () => {
  const source = await JSZip.loadAsync(await readFile("tests/vault/docx-exploration/read-search-only.docx"));
  const part = "word/document.xml";
  const xml = await source.file(part)!.async("string");
  const toc = (level: number, title: string, page: string, anchor: string, leader: string) =>
    `<w:p><w:pPr><w:pStyle w:val="TOC${level}"/><w:tabs><w:tab w:val="right" w:leader="${leader}" w:pos="9000"/></w:tabs></w:pPr>
      <w:hyperlink w:anchor="${anchor}"><w:r><w:t>${title}</w:t></w:r><w:r><w:tab/><w:t>${page}</w:t></w:r></w:hyperlink></w:p>`;
  const heading = (level: number, title: string, anchor: string) =>
    `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/><w:numPr><w:ilvl w:val="${level - 1}"/><w:numId w:val="25"/></w:numPr></w:pPr>
      <w:bookmarkStart w:id="${level}" w:name="${anchor}"/><w:r><w:t>${title}</w:t></w:r><w:bookmarkEnd w:id="${level}"/></w:p>`;
  source.file(part, xml.replace("<w:sectPr/>", [
    toc(1, "Overview", "4", "_Toc100", "dot"),
    toc(2, "3.1 Scope", "9", "_Toc101", "none"),
    heading(1, "Overview", "_Toc100"),
    heading(2, "Scope", "_Toc101"),
    "<w:sectPr/>",
  ].join("")));
  source.file("word/numbering.xml", `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="25">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
      <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
    </w:abstractNum><w:num w:numId="25"><w:abstractNumId w:val="25"/>
      <w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride></w:num>
  </w:numbering>`);
  const buffer = await source.generateAsync({ type: "arraybuffer" });
  const f = fixture("with-toc.docx", { readBinary: async () => buffer });
  try {
    const embed = f.root.querySelector<HTMLElement>(".office-file-embed")!;
    await vi.waitFor(() => expect(embed.dataset.state).toBe("ready"));
    const body = embed.querySelector<HTMLElement>(".office-file-embed__body")!;
    const preview = body.querySelector(".office-viewer-docx--paged")!.shadowRoot!;
    const rows = Array.from(preview.querySelectorAll<HTMLElement>("[data-docx-toc-level]"));
    expect(rows.map((row) => [
      row.dataset.docxTocLevel,
      row.querySelector(".office-viewer-docx-toc-title")?.textContent,
      row.querySelector<HTMLElement>(".office-viewer-docx-toc-leader")?.dataset.leader,
      row.querySelector(".office-viewer-docx-toc-page")?.textContent,
    ])).toEqual([
      ["1", "3 Overview", "dot", "4"],
      ["2", "3.1 Scope", "none", "9"],
    ]);
    const scope = preview.querySelector<HTMLElement>('[data-docx-heading-level="2"]')!;
    vi.spyOn(scope, "getBoundingClientRect").mockReturnValue(new DOMRect(0, 240, 100, 20));
    rows[1]!.querySelector<HTMLAnchorElement>("a[data-docx-toc-target]")!.click();
    expect(body.scrollTop).toBe(232);
    expect(f.options.openExternalUrl).not.toHaveBeenCalled();
  } finally {
    f.dispose();
  }
});

it("keeps complete numbered headings in the sidebar even without a body table of contents", async () => {
  const source = await JSZip.loadAsync(await readFile("tests/vault/docx-exploration/read-search-only.docx"));
  const part = "word/document.xml";
  const xml = await source.file(part)!.async("string");
  const heading = (level: number, title: string) =>
    `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/><w:numPr><w:numId w:val="25"/></w:numPr></w:pPr>
      <w:r><w:t>${title}</w:t></w:r></w:p>`;
  source.file(part, xml.replace("<w:sectPr/>", [
    heading(1, "主要需求"),
    heading(2, "身份验证"),
    heading(2, "1.2 已编号的标题"),
    "<w:sectPr/>",
  ].join("")));
  source.file("word/numbering.xml", `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="25">
      <w:lvl w:ilvl="0"><w:pStyle w:val="Heading1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/></w:lvl>
      <w:lvl w:ilvl="1"><w:pStyle w:val="Heading2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:isLgl/></w:lvl>
    </w:abstractNum><w:num w:numId="25"><w:abstractNumId w:val="25"/></w:num>
  </w:numbering>`);
  const buffer = await source.generateAsync({ type: "arraybuffer" });
  const f = fixture("numbered-headings.docx", { readBinary: async () => buffer });
  try {
    const embed = f.root.querySelector<HTMLElement>(".office-file-embed")!;
    await vi.waitFor(() => expect(embed.dataset.state).toBe("ready"));
    embed.querySelector<HTMLButtonElement>('[data-action="toggle-docx-outline"]')!.click();
    const headings = Array.from(embed.querySelectorAll('[data-action="docx-outline-entry"]'));
    expect(headings.map((entry) => entry.textContent)).toEqual([
      "一、主要需求", "1.1 身份验证", "1.2 已编号的标题",
    ]);
    expect(headings[1]?.getAttribute("aria-label")).toContain("1.1 身份验证");
  } finally {
    f.dispose();
  }
});
