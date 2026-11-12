import { readFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";

import { inspectDocxPackage } from "../../src/docx/docx-semantic-model";
import type { DocxSemanticModel } from "../../src/docx/docx-semantic-model";
import { BoundedDocxRendererAdapter } from "../../src/docx/renderer/bounded-docx-renderer-adapter";
import type { DocxRendererAdapter } from "../../src/docx/renderer/docx-renderer-adapter";

async function stressFixture(): Promise<ArrayBuffer> {
  const bytes = await readFile(
    path.resolve(
      "tests/fixtures/docx-exploration/performance-stress-5000-paragraphs.docx",
    ),
  );
  return Uint8Array.from(bytes).buffer;
}

describe("bounded DOCX renderer", () => {
  it("never materializes the full stress document and can reveal a distant paragraph", async () => {
    const source = await stressFixture();
    const signal = new AbortController().signal;
    const model = await inspectDocxPackage(source, signal);
    const delegateOpen = vi.fn();
    const delegate = { open: delegateOpen } as unknown as DocxRendererAdapter;
    const renderer = new BoundedDocxRendererAdapter(delegate, {
      windowSize: 240,
    });
    const container = document.createElement("div");

    const session = await renderer.open(source, container, model, signal, "simplified");

    expect(delegateOpen).not.toHaveBeenCalled();
    expect(
      container.querySelectorAll("[data-docx-paragraph-ordinal]").length,
    ).toBeLessThanOrEqual(240);
    expect(container.querySelectorAll("*").length).toBeLessThanOrEqual(1_200);

    const distant = session.revealParagraph(4_999);
    expect(distant?.textContent).toContain("Stress paragraph 4999");
    expect(distant?.dataset.docxParagraphOrdinal).toBe("4999");
    expect(
      container.querySelectorAll("[data-docx-paragraph-ordinal]").length,
    ).toBeLessThanOrEqual(240);
    expect(container.querySelectorAll("*").length).toBeLessThanOrEqual(1_200);
  });

  it("does not silently omit links, images, or unavailable body content", async () => {
    const paragraphs = Array.from({ length: 1_001 }, (_, index) => ({
      ordinal: index + 1,
      text: index === 1_000 ? "Read evidence and equation" : `Paragraph ${index + 1}`,
      searchText: index === 1_000 ? "read evidence and equation" : `paragraph ${index + 1}`,
      styleId: null,
      headingLevel: null,
      listItem: false,
      tableDepth: 0,
      bookmarks: [],
      hyperlinks: index === 1_000
        ? [{ kind: "external" as const, label: "evidence", target: "https://example.com/evidence" }]
        : [],
      unavailableContent: index === 1_000 ? ["equation" as const] : [],
      inlineImageCount: index === 1_000 ? 1 : 0,
    }));
    const model: DocxSemanticModel = {
      paragraphs,
      bookmarkTargets: [],
      unavailableBodyBlocks: [
        { afterParagraphOrdinal: 1_001, kinds: ["alt-chunk"] },
      ],
      hasUnavailableBodyContent: true,
    };
    const renderer = new BoundedDocxRendererAdapter(
      { open: vi.fn() } as unknown as DocxRendererAdapter,
      { windowSize: 240 },
    );
    const container = document.createElement("div");

    const session = await renderer.open(
      new ArrayBuffer(0),
      container,
      model,
      new AbortController().signal,
      "simplified",
    );
    const distant = session.revealParagraph(1_001)!;

    const external = distant.querySelector<HTMLAnchorElement>(
      'a[data-docx-external-link="true"]',
    );
    expect(external?.textContent).toBe("evidence");
    expect(external?.href).toBe("https://example.com/evidence");
    expect(
      container.querySelectorAll(".office-viewer-docx-unavailable-content"),
    ).toHaveLength(3);
  });
});
