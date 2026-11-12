import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { inspectDocxPackage } from "../../src/docx/docx-semantic-model";
import { DocxRenderSource } from "../../src/docx/docx-render-source";
import { createDocxRendererAdapter } from "../../src/docx/renderer/create-docx-renderer-adapter";
import {
  mapRenderedParagraphs,
  prepareRenderedDocxReadingLayout,
  sanitizeRenderedDocx,
  type DocxRendererAdapter,
} from "../../src/docx/renderer/docx-renderer-adapter";
import { DocxPreviewRendererAdapter } from "../../src/docx/renderer/docx-preview-renderer-adapter";

const W =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

async function rendererFixture(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="word/document.xml"/>
      </Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
      <w:document xmlns:w="${W}" xmlns:w14="${W14}">
        <w:body>
          <w:p w14:paraId="00000001">
            <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
            <w:r><w:t>Quarterly outlook</w:t></w:r>
          </w:p>
          <w:p w14:paraId="00000002">
            <w:r><w:t>Demand remains resilient.</w:t></w:r>
          </w:p>
          <w:tbl><w:tr><w:tc>
            <w:p w14:paraId="00000003">
              <w:r><w:t>Table evidence</w:t></w:r>
            </w:p>
          </w:tc></w:tr></w:tbl>
          <w:sectPr/>
        </w:body>
      </w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
}

it("preserves Word blank lines and inherited character indents in both preview modes", async () => {
  const zip = await JSZip.loadAsync(await rendererFixture());
  zip.file("word/document.xml", `<w:document xmlns:w="${W}"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
    <w:p><w:pPr><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr><w:r><w:cr/></w:r></w:p>
    <w:p/>
    <w:p><w:pPr><w:ind w:left="480" w:hangingChars="100"/></w:pPr><w:r><w:t>Last</w:t></w:r></w:p>
  </w:body></w:document>`);
  zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file("word/styles.xml", `<w:styles xmlns:w="${W}">
    <w:style w:type="paragraph" w:styleId="Base">
      <w:pPr><w:ind w:firstLineChars="200"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Base"/></w:style>
  </w:styles>`);
  const signal = new AbortController().signal;
  const bytes = await zip.generateAsync({ type: "arraybuffer" });
  const model = await inspectDocxPackage(bytes, signal);
  expect(model.paragraphs.map((p) => p.text)).toEqual(["First", "Last"]);
  const source = new DocxRenderSource(bytes, model);
  const renderer = createDocxRendererAdapter({ pageLayout: true });
  const container = document.createElement("div");
  const full = await source.open(renderer, container, "full", signal);
  const paragraphs = container.firstElementChild?.shadowRoot?.querySelectorAll("article p");
  expect(paragraphs).toHaveLength(4);
  expect(paragraphs?.[1]?.querySelector("br")).not.toBeNull();
  expect(paragraphs?.[1]?.querySelector<HTMLElement>("span[style]")?.style.fontSize).toBe("18pt");
  expect(paragraphs?.[2]?.textContent).toBe("\u00a0");
  expect(full.revealParagraph(1)?.style.textIndent).toBe("24pt");
  expect(full.revealParagraph(2)?.style.textIndent).toBe("-11pt");
  expect(full.revealParagraph(2)?.style.marginInlineStart).toBe("24pt");
  full.dispose();

  const simplified = await source.open(renderer, container, "simplified", signal);
  const first = simplified.revealParagraph(1);
  const last = simplified.revealParagraph(2);
  expect(first?.style.textIndent).toBe("24pt");
  expect(last?.style.textIndent).toBe("-11pt");
  const blank = first?.nextElementSibling as HTMLElement;
  expect(blank.getAttribute("aria-hidden")).toBe("true");
  expect(parseFloat(blank.style.height)).toBeCloseTo((2 * 18 + 11) * 1.2 * 4 / 3);
  expect(blank.nextElementSibling).toBe(last);
  simplified.dispose();
});

describe("DOCX renderer mapping and sanitization", () => {
  it("maps only exact semantic paragraph text in document order", () => {
    const container = document.createElement("div");
    container.innerHTML = `
      <h1>Quarterly outlook</h1>
      <p>Demand remains resilient.</p>
    `;
    const mapping = mapRenderedParagraphs(container, [
      {
        ordinal: 1,
        text: "Quarterly outlook",
        searchText: "quarterly outlook",
        styleId: "Heading1",
        headingLevel: 1,
        listItem: false,
        tableDepth: 0,
        bookmarks: [],
        hyperlinks: [],
        inlineImageCount: 0,
        unavailableContent: [],
      },
      {
        ordinal: 2,
        text: "Demand remains resilient.",
        searchText: "demand remains resilient.",
        styleId: null,
        headingLevel: null,
        listItem: false,
        tableDepth: 0,
        bookmarks: [],
        hyperlinks: [],
        inlineImageCount: 0,
        unavailableContent: [],
      },
    ]);

    expect(mapping.get(2)?.dataset.docxParagraphOrdinal).toBe("2");
    expect(() =>
      mapRenderedParagraphs(container, [
        {
          ordinal: 1,
          text: "Different text",
          searchText: "different text",
          styleId: null,
          headingLevel: null,
          listItem: false,
          tableDepth: 0,
          bookmarks: [],
          hyperlinks: [],
          inlineImageCount: 0,
          unavailableContent: [],
        },
        {
          ordinal: 2,
          text: "Demand remains resilient.",
          searchText: "demand remains resilient.",
          styleId: null,
          headingLevel: null,
          listItem: false,
          tableDepth: 0,
          bookmarks: [],
          hyperlinks: [],
          inlineImageCount: 0,
          unavailableContent: [],
        },
      ])
    ).toThrow(/does not match/);
  });

  it("treats hyperlink break whitespace as equivalent for mapping", () => {
    const container = document.createElement("div");
    container.innerHTML = `
      <p>【腾讯文档】记录https://docs.qq.com/sheet/A</p>
      <p>Next paragraph</p>
    `;
    const mapping = mapRenderedParagraphs(container, [
      {
        ordinal: 1,
        text: "【腾讯文档】记录\nhttps://docs.qq.com/sheet/A",
        searchText: "【腾讯文档】记录\nhttps://docs.qq.com/sheet/a",
        styleId: null,
        headingLevel: null,
        listItem: false,
        tableDepth: 0,
        bookmarks: [],
        hyperlinks: [],
        inlineImageCount: 0,
        unavailableContent: [],
      },
      {
        ordinal: 2,
        text: "Next paragraph",
        searchText: "next paragraph",
        styleId: null,
        headingLevel: null,
        listItem: false,
        tableDepth: 0,
        bookmarks: [],
        hyperlinks: [],
        inlineImageCount: 0,
        unavailableContent: [],
      },
    ]);
    expect(mapping.get(1)?.dataset.docxParagraphOrdinal).toBe("1");
    expect(mapping.size).toBe(2);
  });

  it("blocks active DOM and unsafe navigation while retaining allowlisted links", () => {
    const container = document.createElement("div");
    container.innerHTML = `
      <script>window.bad = true</script>
      <a id="safe" href="https://example.com">safe</a>
      <a id="local" href="file:///tmp/private">local</a>
      <a id="bookmark" href="#Target">jump</a>
      <img id="remote" src="https://example.com/private.png">
      <img id="inline" src="data:image/png;base64,AA==">
    `;

    sanitizeRenderedDocx(container);

    expect(container.querySelector("script")).toBeNull();
    expect(container.querySelector("#safe")?.getAttribute("href")).toBe(
      "https://example.com",
    );
    expect(container.querySelector("#local")?.getAttribute("href")).toBeNull();
    expect(
      (container.querySelector("#bookmark") as HTMLElement | null)?.dataset
        .docxBookmark,
    ).toBe("Target");
    expect(container.querySelector("#remote")?.getAttribute("src")).toBeNull();
    expect(container.querySelector("#inline")?.getAttribute("src")).toContain(
      "data:image/png",
    );
  });

  it("sanitizes links and images created by another window", () => {
    const frame = document.createElement("iframe");
    document.body.append(frame);
    const frameNode = (frame.contentWindow as unknown as { Node: typeof Node })
      .Node;
    Object.defineProperty(frameNode.prototype, "instanceOf", {
      configurable: true,
      value: Node.prototype.instanceOf,
    });
    const frameDocument = frame.contentDocument;
    expect(frameDocument).not.toBeNull();
    const container = frameDocument!.createElement("div");
    container.innerHTML = `
      <a id="safe" href="https://example.com">safe</a>
      <a id="local" href="file:///tmp/private">local</a>
      <img id="remote" src="https://example.com/private.png">
    `;

    sanitizeRenderedDocx(container);

    expect(container.querySelector("#safe")?.getAttribute("href")).toBe(
      "https://example.com",
    );
    expect(container.querySelector("#local")?.getAttribute("href")).toBeNull();
    expect(container.querySelector("#remote")?.getAttribute("src")).toBeNull();
    frame.remove();
  });

  it("marks fixed-size docx-preview content for responsive reading", () => {
    const container = document.createElement("div");
    container.innerHTML = `
      <section class="docx">
        <div id="large" style="display: inline-block; width: 428.65pt; height: 172.2pt">
          <img id="large-image" src="data:image/png;base64,AA==" style="width: 428.65pt; height: 172.2pt">
        </div>
        <div id="small" style="display: inline-block; width: 72pt; height: 66pt">
          <img id="small-image" src="data:image/png;base64,AA==" style="width: 72pt; height: 66pt">
        </div>
        <div id="complex" style="display: inline-block; width: 200pt">
          <span>caption</span>
          <img id="complex-image" src="data:image/png;base64,AA==">
        </div>
        <table id="wide-table" style="width: 457.95pt">
          <colgroup><col id="narrow-column" style="width: 100pt"><col id="wide-column" style="width: 300pt"></colgroup>
          <tr><td>Label</td><td>Evidence</td></tr>
        </table>
      </section>
    `;

    prepareRenderedDocxReadingLayout(container);

    const large = container.querySelector<HTMLElement>("#large");
    const small = container.querySelector<HTMLElement>("#small");
    expect(large?.classList.contains("office-viewer-docx-media-wrapper")).toBe(
      true,
    );
    expect(
      large?.style.getPropertyValue("--office-viewer-docx-media-width"),
    ).toBe("428.65pt");
    expect(large?.style.width).toBe("");
    expect(large?.style.height).toBe("");
    expect(small?.classList.contains("office-viewer-docx-media-wrapper")).toBe(
      true,
    );
    expect(
      small?.style.getPropertyValue("--office-viewer-docx-media-width"),
    ).toBe("72pt");
    expect(
      container
        .querySelector("#complex")
        ?.classList.contains("office-viewer-docx-media-wrapper"),
    ).toBe(false);
    expect(
      Array.from(container.querySelectorAll("img")).every((image) =>
        image.classList.contains("office-viewer-docx-media"),
      ),
    ).toBe(true);
    expect(container.querySelector<HTMLElement>("#large-image")?.style.width)
      .toBe("");
    expect(container.querySelector<HTMLElement>("#large-image")?.style.height)
      .toBe("");
    const table = container.querySelector<HTMLElement>("#wide-table");
    expect(table?.classList.contains("office-viewer-docx-table")).toBe(true);
    expect(table?.style.width).toBe("");
    expect(
      table?.style.getPropertyValue("--office-viewer-docx-table-width"),
    ).toBe("457.95pt");
    expect(
      container
        .querySelector<HTMLElement>("#narrow-column")
        ?.style.getPropertyValue("--office-viewer-docx-column-width"),
    ).toBe("25%");
    expect(
      container
        .querySelector<HTMLElement>("#wide-column")
        ?.style.getPropertyValue("--office-viewer-docx-column-width"),
    ).toBe("75%");
    expect(container.querySelector<HTMLElement>("#narrow-column")?.style.width)
      .toBe("");
  });
});

describe("docx-preview renderer adapter", () => {
  const createAdapter = () => new DocxPreviewRendererAdapter();
  it("maps the same fixture to the project-owned semantic paragraphs", async () => {
    const bytes = await rendererFixture();
    const signal = new AbortController().signal;
    const model = await inspectDocxPackage(bytes, signal);
    const container = document.createElement("div");
    const adapter: DocxRendererAdapter = createAdapter();

    const session = await adapter.open(bytes, container, model, signal);

    expect(session.paragraphElements.size).toBe(3);
    expect(session.paragraphElements.get(1)?.textContent).toContain(
      "Quarterly outlook",
    );
    expect(session.paragraphElements.get(3)?.textContent).toContain(
      "Table evidence",
    );
    expect(container.querySelectorAll("[data-docx-paragraph-ordinal]"))
      .toHaveLength(3);
    session.dispose();
    expect(container.childElementCount).toBe(0);
  });

  it("keeps a readable preview when only one hyperlink paragraph disagrees on whitespace", async () => {
    const zip = new JSZip();
    const R =
      "http://schemas.openxmlformats.org/package/2006/relationships";
    const PR =
      "http://schemas.openxmlformats.org/package/2006/content-types";
    const OR =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="${PR}">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>`,
    );
    zip.file(
      "_rels/.rels",
      `<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="${R}">
          <Relationship Id="rId1" Type="${OR}/officeDocument" Target="word/document.xml"/>
        </Relationships>`,
    );
    zip.file(
      "word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="${R}">
          <Relationship Id="rId1" Type="${OR}/hyperlink"
            Target="https://docs.qq.com/sheet/A" TargetMode="External"/>
        </Relationships>`,
    );
    zip.file(
      "word/document.xml",
      `<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="${W}" xmlns:r="${OR}" xmlns:w14="${W14}">
          <w:body>
            <w:p w14:paraId="00000001">
              <w:r><w:t>Before</w:t></w:r>
            </w:p>
            <w:p w14:paraId="00000002">
              <w:r><w:t>【腾讯文档】记录</w:t></w:r>
              <w:r><w:br/></w:r>
              <w:hyperlink r:id="rId1">
                <w:r><w:t>https://docs.qq.com/sheet/A</w:t></w:r>
              </w:hyperlink>
            </w:p>
            <w:p w14:paraId="00000003">
              <w:r><w:t>After</w:t></w:r>
            </w:p>
            <w:sectPr/>
          </w:body>
        </w:document>`,
    );
    const bytes = await zip.generateAsync({ type: "arraybuffer" });
    const signal = new AbortController().signal;
    const model = await inspectDocxPackage(bytes, signal);
    const container = document.createElement("div");
    const session = await createAdapter().open(bytes, container, model, signal);
    expect(session.candidate).toBe("docx-preview");
    expect(session.paragraphElements.size).toBe(model.paragraphs.length);
    expect(container.textContent).toContain("【腾讯文档】记录");
    expect(container.textContent).toContain("https://docs.qq.com/sheet/A");
    session.dispose();
  });

  it("does not mutate the destination when already aborted", async () => {
    const bytes = await rendererFixture();
    const controller = new AbortController();
    controller.abort();
    const model = await inspectDocxPackage(
      bytes,
      new AbortController().signal,
    );
    const container = document.createElement("div");
    container.textContent = "unchanged";

    await expect(
      createAdapter().open(bytes, container, model, controller.signal),
    ).rejects.toMatchObject({ name: "AbortError" });
    expect(container.textContent).toBe("unchanged");
  });
});
