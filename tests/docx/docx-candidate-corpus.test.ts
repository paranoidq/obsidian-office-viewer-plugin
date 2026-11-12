import { readFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import {
  createSafeDocxRendererBuffer,
  inspectDocxPackage,
} from "../../src/docx/docx-semantic-model";
import type { DocxRendererAdapter } from "../../src/docx/renderer/docx-renderer-adapter";
import { DocxPreviewRendererAdapter } from "../../src/docx/renderer/docx-preview-renderer-adapter";

function fixturePath(name: string): string {
  return path.resolve("tests/fixtures/docx-exploration", name);
}

async function fixtureBytes(name: string): Promise<ArrayBuffer> {
  const bytes = await readFile(fixturePath(name));
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

async function inspect(name: string) {
  return inspectDocxPackage(
    await fixtureBytes(name),
    new AbortController().signal,
  );
}

describe("committed DOCX semantic corpus", () => {
  it("reads documents with or without native paragraph identities", async () => {
    const reference = await inspect("body-led-reference.docx");
    const readOnly = await inspect("read-search-only.docx");

    expect(reference.paragraphs).toHaveLength(11);
    expect(readOnly.paragraphs).toHaveLength(2);
    expect(reference).not.toHaveProperty("referenceCapable");
    expect(readOnly).not.toHaveProperty("referenceCapable");
  });

  it("extracts only final revision body text", async () => {
    const model = await inspect("final-revisions.docx");

    expect(model.paragraphs.map((paragraph) => paragraph.text)).toEqual([
      "Approved new language",
    ]);
  });

  it("detects remaining unavailable content alongside supported equations", async () => {
    const model = await inspect("unavailable-content.docx");

    expect(model.hasUnavailableBodyContent).toBe(true);
    expect(model.equations).toHaveLength(1);
    expect(
      new Set([
        ...model.paragraphs.flatMap(
          (paragraph) => paragraph.unavailableContent,
        ),
        ...model.unavailableBodyBlocks.flatMap((block) => block.kinds),
      ]),
    ).toEqual(
      new Set([
        "embedded-object",
        "external-image",
        "alt-chunk",
      ]),
    );
  });

  it("rejects the committed active-content fixture", async () => {
    await expect(inspect("active-content.docx")).rejects.toMatchObject({
      category: "incompatible",
    });
  });
});

describe("docx-preview committed DOCX corpus", () => {
  const createAdapter = () => new DocxPreviewRendererAdapter();
  it.each([
    "body-led-reference.docx",
    "read-search-only.docx",
    "final-revisions.docx",
    "unavailable-content.docx",
  ])("maps every semantic paragraph for %s", async (name) => {
    const bytes = await fixtureBytes(name);
    const signal = new AbortController().signal;
    const model = await inspectDocxPackage(bytes, signal);
    const rendererBytes = await createSafeDocxRendererBuffer(bytes, signal, undefined, model);
    const container = document.createElement("div");
    const adapter: DocxRendererAdapter = createAdapter();

    const session = await adapter.open(
      rendererBytes,
      container,
      model,
      signal,
    );

    expect(session.paragraphElements.size).toBe(model.paragraphs.length);
    expect(container.querySelectorAll("math[data-docx-equation]")).toHaveLength(model.equations?.length ?? 0);
    session.dispose();
  });
});
