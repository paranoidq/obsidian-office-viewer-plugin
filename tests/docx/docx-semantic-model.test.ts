import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import { DocxOpenError } from "../../src/docx/docx-open-error";
import {
  createSafeDocxRendererBuffer,
  DOCX_ZIP_LIMITS,
  inspectDocxPackage,
  searchDocxBody,
} from "../../src/docx/docx-semantic-model";

const W =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const R =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const M =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

interface FixtureOptions {
  readonly body: string;
  readonly relationships?: string;
  readonly contentTypeOverrides?: string;
  readonly extraParts?: Readonly<Record<string, string | Uint8Array>>;
}

async function docxFixture(options: FixtureOptions): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        ${options.contentTypeOverrides ?? ""}
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
      <w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}"
        xmlns:m="${M}" xmlns:a="${A}">
        <w:body>${options.body}<w:sectPr/></w:body>
      </w:document>`,
  );
  if (options.relationships !== undefined) {
    zip.file(
      "word/_rels/document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          ${options.relationships}
        </Relationships>`,
    );
  }
  for (const [path, value] of Object.entries(options.extraParts ?? {})) {
    zip.file(path, value);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

function paragraph(
  id: string | null,
  text: string,
  properties = "",
): string {
  const identity = id === null ? "" : ` w14:paraId="${id}"`;
  return `<w:p${identity}>${properties}<w:r><w:t>${text}</w:t></w:r></w:p>`;
}

async function inspect(buffer: ArrayBuffer) {
  return inspectDocxPackage(buffer, new AbortController().signal);
}

describe("DOCX project-owned semantic model", () => {
  it("resolves direct, inherited, and localized heading levels from the styles relationship", async () => {
    const model = await inspect(await docxFixture({
      relationships: `<Relationship Id="styles" Type="${R}/styles" Target="format/custom.xml"/>`,
      extraParts: {
        "word/format/custom.xml": `<w:styles xmlns:w="${W}">
          <w:style w:type="paragraph" w:styleId="Custom"><w:basedOn w:val="Parent"/></w:style>
          <w:style w:type="paragraph" w:styleId="Parent"><w:basedOn w:val="Base"/></w:style>
          <w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="Localized"><w:name w:val="标题 5"/></w:style>
          <w:style w:type="paragraph" w:styleId="Cycle"><w:basedOn w:val="Cycle"/></w:style>
        </w:styles>`,
      },
      body: [
        paragraph(null, "Inherited", '<w:pPr><w:pStyle w:val="Custom"/></w:pPr>'),
        paragraph(null, "Override", '<w:pPr><w:pStyle w:val="Custom"/><w:outlineLvl w:val="8"/></w:pPr>'),
        paragraph(null, "Body", '<w:pPr><w:pStyle w:val="Custom"/><w:outlineLvl w:val="9"/></w:pPr>'),
        paragraph(null, "Localized", '<w:pPr><w:pStyle w:val="Localized"/></w:pPr>'),
        paragraph(null, "Built in", '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'),
        paragraph(null, "Cycle", '<w:pPr><w:pStyle w:val="Cycle"/></w:pPr>'),
        paragraph(null, "Ordinary text"),
      ].join(""),
    }));
    expect(model.paragraphs.map(({ headingLevel }) => headingLevel)).toEqual([3, 9, null, 5, 1, null, null]);
  });

  it("extracts searchable final-view body paragraphs in canonical order", async () => {
    const buffer = await docxFixture({
      body: `
        <w:p w14:paraId="00000001">
          <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
          <w:r><w:t>Report heading</w:t></w:r>
        </w:p>
        <w:p w14:paraId="00000002">
          <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>
          <w:r><w:t>Visible </w:t></w:r>
          <w:ins><w:r><w:t>addition</w:t></w:r></w:ins>
          <w:del><w:r><w:delText> removed</w:delText></w:r></w:del>
          <w:r><w:rPr><w:vanish/></w:rPr><w:t> hidden</w:t></w:r>
          <w:txbxContent>
            <w:p w14:paraId="00000099"><w:r><w:t>textbox</w:t></w:r></w:p>
          </w:txbxContent>
        </w:p>
        <w:tbl><w:tr><w:tc>
          ${paragraph("00000003", "Cell paragraph")}
        </w:tc></w:tr></w:tbl>
        ${paragraph("00000004", "   ")}
      `,
    });

    const model = await inspect(buffer);

    expect(model.paragraphs).toEqual([
      expect.objectContaining({
        ordinal: 1,
        text: "Report heading",
        styleId: "Heading1",
        listItem: false,
        tableDepth: 0,
      }),
      expect.objectContaining({
        ordinal: 2,
        text: "Visible addition",
        listItem: true,
        tableDepth: 0,
      }),
      expect.objectContaining({
        ordinal: 3,
        text: "Cell paragraph",
        tableDepth: 1,
      }),
    ]);
  });

  it("reads paragraphs regardless of missing or duplicate native identities", async () => {
    const missing = await inspect(
      await docxFixture({
        body:
          paragraph("00000001", "First") +
          paragraph(null, "Second"),
      }),
    );
    const duplicate = await inspect(
      await docxFixture({
        body:
          paragraph("00000001", "First") +
          paragraph("00000001", "Second"),
      }),
    );

    expect(missing).not.toHaveProperty("referenceCapable");
    expect(duplicate).not.toHaveProperty("referenceCapable");
    expect(missing.paragraphs[1]).not.toHaveProperty("paragraphId");
    expect(duplicate.paragraphs[1]).not.toHaveProperty("paragraphId");
    expect(missing.paragraphs).toHaveLength(2);
    expect(duplicate.paragraphs).toHaveLength(2);
  });

  it("returns one search result per matching paragraph", async () => {
    const model = await inspect(
      await docxFixture({
        body:
          paragraph("00000001", "Risk risk response") +
          paragraph("00000002", "No match") +
          paragraph("00000003", "Another RISK"),
      }),
    );

    expect(searchDocxBody(model, "risk")).toEqual([
      {
        paragraphOrdinal: 1,
        matchCount: 2,
        text: "Risk risk response",
      },
      {
        paragraphOrdinal: 3,
        matchCount: 1,
        text: "Another RISK",
      },
    ]);
    expect(searchDocxBody(model, "  ")).toEqual([]);
  });

  it("classifies allowlisted, blocked, and bookmark hyperlinks without fetching", async () => {
    const relationships = `
      <Relationship Id="rHttp"
        Type="${R}/hyperlink" Target="https://example.com/report" TargetMode="External"/>
      <Relationship Id="rFile"
        Type="${R}/hyperlink" Target="file:///tmp/private" TargetMode="External"/>
    `;
    const body = `
      <w:p w14:paraId="00000001">
        <w:bookmarkStart w:id="1" w:name="Target"/>
        <w:r><w:t>Target paragraph. </w:t></w:r>
        <w:hyperlink r:id="rHttp"><w:r><w:t>Website</w:t></w:r></w:hyperlink>
        <w:hyperlink r:id="rFile"><w:r><w:t>Local file</w:t></w:r></w:hyperlink>
        <w:hyperlink w:anchor="Target"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>
      </w:p>
    `;

    const model = await inspect(await docxFixture({ body, relationships }));

    expect(model.bookmarkTargets).toEqual([
      { bookmark: "Target", paragraphOrdinal: 1 },
    ]);
    expect(model.paragraphs[0]?.hyperlinks).toEqual([
      {
        kind: "external",
        label: "Website",
        target: "https://example.com/report",
      },
      {
        kind: "blocked",
        label: "Local file",
        target: "file:///tmp/private",
      },
      { kind: "bookmark", label: "Jump", bookmark: "Target" },
    ]);
  });

  it("does not expose duplicate bookmark targets", async () => {
    const withBookmark = (id: string, name: string, text: string) => `
      <w:p w14:paraId="${id}">
        <w:bookmarkStart w:id="${id}" w:name="${name}"/>
        <w:r><w:t>${text}</w:t></w:r>
      </w:p>`;
    const model = await inspect(
      await docxFixture({
        body:
          withBookmark("00000001", "Duplicate", "First") +
          withBookmark("00000002", "Duplicate", "Second"),
      }),
    );

    expect(model.bookmarkTargets).toEqual([]);
  });

  it("detects unavailable main-body content and external images", async () => {
    const model = await inspect(
      await docxFixture({
        relationships: `
          <Relationship Id="rImage"
            Type="${R}/image" Target="https://example.com/image.png" TargetMode="External"/>
        `,
        body: `
          <w:p w14:paraId="00000001">
            <w:r><w:t>Mixed content</w:t></w:r>
            <m:oMath><m:phant><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant></m:oMath>
            <w:object/>
            <w:r><w:drawing><a:blip r:embed="rImage"/></w:drawing></w:r>
          </w:p>
        `,
      }),
    );

    expect(model.hasUnavailableBodyContent).toBe(true);
    expect(model.unavailableBodyBlocks).toEqual([]);
    expect(model.paragraphs[0]?.unavailableContent).toEqual([
      "equation",
      "embedded-object",
      "external-image",
    ]);
  });

  it("materializes native charts into PNG pictures for the renderer", async () => {
    const chartXml = `<?xml version="1.0" encoding="UTF-8"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart>
          <c:title><c:tx><c:rich><a:p><a:r><a:t>Demo</a:t></a:r></a:p></c:rich></c:tx></c:title>
          <c:plotArea>
            <c:lineChart>
              <c:ser>
                <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:tx>
                <c:cat><c:strRef><c:strCache>
                  <c:pt idx="0"><c:v>202301</c:v></c:pt>
                  <c:pt idx="1"><c:v>202302</c:v></c:pt>
                </c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:numCache>
                  <c:pt idx="0"><c:v>0.99</c:v></c:pt>
                  <c:pt idx="1"><c:v>0.98</c:v></c:pt>
                </c:numCache></c:numRef></c:val>
              </c:ser>
            </c:lineChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;
    const source = await docxFixture({
      relationships: `
        <Relationship Id="rChart" Type="${R}/chart" Target="charts/chart1.xml"/>
      `,
      extraParts: { "word/charts/chart1.xml": chartXml },
      body: `
        <w:p w14:paraId="00000001">
          <w:r><w:t>Before chart</w:t></w:r>
        </w:p>
        <w:p w14:paraId="00000002">
          <w:r><w:drawing>
            <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
              <wp:extent cx="5274310" cy="2139315"/>
              <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rChart"/>
            </wp:inline>
          </w:drawing></w:r>
        </w:p>
        <w:p w14:paraId="00000003">
          <w:r><w:t>After chart</w:t></w:r>
        </w:p>
      `,
    });

    const safe = await createSafeDocxRendererBuffer(
      source,
      new AbortController().signal,
    );
    const safeZip = await JSZip.loadAsync(safe);
    const safeDocument = await safeZip.file("word/document.xml")!.async("text");
    const safeRels = await safeZip
      .file("word/_rels/document.xml.rels")!
      .async("text");

    expect(safeDocument).not.toContain("<c:chart");
    expect(safeDocument).toContain("a:blip");
    expect(safeDocument).toContain("Before chart");
    expect(safeDocument).toContain("After chart");
    expect(safeRels).toContain("docx-chart-");
    expect(
      Object.keys(safeZip.files).some((name) =>
        /^word\/media\/docx-chart-\d+\.png$/.test(name),
      ),
    ).toBe(true);
  });

  it("detects and removes externally linked images that use r:link", async () => {
    const relationships = `
      <Relationship Id="rImage" Type="${R}/image"
        Target="https://example.com/linked.png" TargetMode="External"/>
    `;
    const source = await docxFixture({
      relationships,
      body: `
        <w:p w14:paraId="00000001">
          <w:r><w:t>Linked image</w:t></w:r>
          <w:r><w:drawing><a:blip r:link="rImage"/></w:drawing></w:r>
        </w:p>
      `,
    });

    const model = await inspect(source);
    const safe = await createSafeDocxRendererBuffer(
      source,
      new AbortController().signal,
    );
    const safeZip = await JSZip.loadAsync(safe);
    const safeDocument = await safeZip.file("word/document.xml")!.async("text");
    const safeRelationships = await safeZip
      .file("word/_rels/document.xml.rels")!
      .async("text");

    expect(model.paragraphs[0]?.unavailableContent).toContain("external-image");
    expect(safeDocument).not.toContain("r:link=\"rImage\"");
    expect(safeRelationships).not.toContain("Id=\"rImage\"");
    expect(new Uint8Array(source)).not.toHaveLength(0);
  });

  it("creates a sanitized renderer derivative without mutating source bytes", async () => {
    const source = await docxFixture({
      relationships: `
        <Relationship Id="rImage"
          Type="${R}/image" Target="https://example.com/image.png" TargetMode="External"/>
      `,
      body: `
        <w:p w14:paraId="00000001">
          <w:r><w:t>Visible</w:t></w:r>
          <w:del><w:r><w:delText>Deleted</w:delText></w:r></w:del>
          <w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t></w:r>
          <w:r><w:drawing><a:blip r:embed="rImage"/></w:drawing></w:r>
        </w:p>
      `,
    });
    const before = new Uint8Array(source).slice();

    const derivative = await createSafeDocxRendererBuffer(
      source,
      new AbortController().signal,
    );

    expect(new Uint8Array(source)).toEqual(before);
    expect(new Uint8Array(derivative)).not.toEqual(before);
    const renderedModel = await inspect(derivative);
    expect(renderedModel.paragraphs[0]?.text).toBe("Visible");
    expect(renderedModel.paragraphs[0]?.unavailableContent).toEqual([]);
  });

  it("retains object-only body content as a non-ordinal unavailable block", async () => {
    const model = await inspect(
      await docxFixture({
        body:
          paragraph("00000001", "Before") +
          `<w:p w14:paraId="00000002"><w:object/></w:p>` +
          paragraph("00000003", "After") +
          `<w:altChunk r:id="rChunk"/>`,
      }),
    );

    expect(model.paragraphs.map((item) => item.text)).toEqual([
      "Before",
      "After",
    ]);
    expect(model.unavailableBodyBlocks).toEqual([
      { afterParagraphOrdinal: 1, kinds: ["embedded-object"] },
      { afterParagraphOrdinal: 2, kinds: ["alt-chunk"] },
    ]);
  });

  it("rejects active content and malformed XML with stable categories", async () => {
    const active = await docxFixture({
      body: paragraph("00000001", "Body"),
      contentTypeOverrides: `
        <Override PartName="/word/vbaProject.bin"
          ContentType="application/vnd.ms-office.vbaProject"/>
      `,
      extraParts: {
        "word/vbaProject.bin": new Uint8Array([1, 2, 3]),
      },
    });
    await expect(inspect(active)).rejects.toMatchObject({
      category: "incompatible",
    } satisfies Partial<DocxOpenError>);

    const malformedZip = new JSZip();
    malformedZip.file("[Content_Types].xml", "<Types>");
    malformedZip.file("_rels/.rels", "<Relationships/>");
    malformedZip.file("word/document.xml", "<broken");
    await expect(
      inspect(await malformedZip.generateAsync({ type: "arraybuffer" })),
    ).rejects.toMatchObject({
      category: "malformed",
    } satisfies Partial<DocxOpenError>);
  });

  it("rejects encrypted OOXML and unsupported external relationships", async () => {
    const protectedBytes = await readFile(
      "tests/fixtures/failure/protected-encrypted.pptx",
    );
    await expect(
      inspect(Uint8Array.from(protectedBytes).buffer),
    ).rejects.toMatchObject({
      category: "protected",
    } satisfies Partial<DocxOpenError>);

    const external = await docxFixture({
      body: paragraph("00000001", "Body"),
      relationships: `
        <Relationship Id="rTemplate" Type="${R}/attachedTemplate"
          Target="https://example.com/template.dotx" TargetMode="External"/>
      `,
    });
    await expect(inspect(external)).rejects.toMatchObject({
      category: "incompatible",
    } satisfies Partial<DocxOpenError>);
  });

  it("enforces ZIP entry and XML-part limits before parsing", async () => {
    const entryBomb = await JSZip.loadAsync(
      await docxFixture({ body: paragraph("00000001", "Body") }),
    );
    for (let index = 0; index < DOCX_ZIP_LIMITS.maxEntries; index += 1) {
      entryBomb.file(`custom/entry-${index}.bin`, "x");
    }
    await expect(
      inspect(await entryBomb.generateAsync({ type: "arraybuffer" })),
    ).rejects.toMatchObject({
      category: "resource-exhausted",
    } satisfies Partial<DocxOpenError>);

    const xmlBomb = await docxFixture({
      body: paragraph(
        "00000001",
        "x".repeat(DOCX_ZIP_LIMITS.maxXmlPartBytes + 1),
      ),
    });
    await expect(inspect(xmlBomb)).rejects.toMatchObject({
      category: "resource-exhausted",
    } satisfies Partial<DocxOpenError>);
  });

  it("honors cancellation before package work begins", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      inspectDocxPackage(new ArrayBuffer(0), controller.signal),
    ).rejects.toMatchObject({ name: "AbortError" });
  });
});
