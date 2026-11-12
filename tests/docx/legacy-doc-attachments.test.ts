import { readFile } from "node:fs/promises";
import CFB from "cfb";
import JSZip from "jszip";
import { expect, it } from "vitest";
import { extractDocxAttachment } from "../../src/docx/docx-attachments";
import { inspectDocxPackage } from "../../src/docx/docx-semantic-model";
import { restoreLegacyDocAttachments } from "../../src/docx/legacy-doc-attachments";

function legacyDocWithObject(previewName: string, payload: Uint8Array): Uint8Array {
  const document = Buffer.alloc(1536);
  document.writeUInt16LE(0xa5ec);
  document.writeUInt16LE(34, 36);
  const textOffset = 512;
  [0x13, 0x14, 1, 0x15].forEach((value, index) =>
    document.writeUInt16LE(value, textOffset + index * 2),
  );
  const fkp = document.subarray(1024);
  fkp[511] = 4;
  for (let i = 0; i <= 4; i++) fkp.writeUInt32LE(textOffset + i * 2, i * 4);
  const separator = Buffer.from("0a0801560801550801036a7b000000", "hex");
  const picture = Buffer.from("550801036a00000000", "hex");
  fkp[21] = 40;
  fkp[22] = 50;
  fkp[80] = separator.length;
  fkp[100] = picture.length;
  separator.copy(fkp, 81);
  picture.copy(fkp, 101);
  const bte = Buffer.alloc(12);
  bte.writeUInt32LE(textOffset);
  bte.writeUInt32LE(textOffset + 8, 4);
  bte.writeUInt32LE(2, 8);
  const fields = Buffer.alloc(22);
  for (let i = 0; i <= 3; i++) fields.writeUInt32LE(i, 4 * i);
  Buffer.from([0x13, 0x3a, 0x14, 0, 0x15, 0x80]).copy(fields, 16);
  // The last field CP is a sentinel, so the field end is the third CP.
  fields.writeUInt32LE(3, 8);
  fields.writeUInt32LE(4, 12);
  const clx = Buffer.alloc(21);
  clx[0] = 2;
  clx.writeUInt32LE(16, 1);
  clx.writeUInt32LE(4, 9);
  clx.writeUInt32LE(textOffset, 15);
  for (const [index, offset, bytes] of [
    [12, 0, bte], [16, bte.length, fields], [33, bte.length + fields.length, clx],
  ] as const) {
    document.writeUInt32LE(offset, 38 + index * 8);
    document.writeUInt32LE(bytes.length, 42 + index * 8);
  }
  const name = Buffer.from(`${previewName}\0`, "utf16le");
  const picf = Buffer.alloc(68 + 8 + 8 + 6 + name.length);
  picf.writeUInt32LE(picf.length);
  picf.writeUInt16LE(68, 4);
  picf.writeUInt16LE(0x000f, 68);
  picf.writeUInt16LE(0xf004, 70);
  picf.writeUInt32LE(8 + 6 + name.length, 72);
  picf.writeUInt16LE(0x0013, 76);
  picf.writeUInt16LE(0xf00b, 78);
  picf.writeUInt32LE(6 + name.length, 80);
  picf.writeUInt16LE(0x8380, 84);
  picf.writeUInt32LE(name.length, 86);
  name.copy(picf, 90);
  const file = CFB.utils.cfb_new();
  CFB.utils.cfb_add(file, "WordDocument", document);
  CFB.utils.cfb_add(file, "0Table", Buffer.concat([bte, fields, clx]));
  CFB.utils.cfb_add(file, "Data", picf);
  CFB.utils.cfb_add(file, "ObjectPool/_123/package", payload);
  return new Uint8Array(CFB.write(file, { type: "buffer" }));
}

it("restores an Excel attachment only at its named legacy preview shape", async () => {
  const workbook = new JSZip();
  workbook.file("xl/workbook.xml", "<workbook/>");
  workbook.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  </Types>`);
  const payload = await workbook.generateAsync({ type: "uint8array" });
  const source = legacyDocWithObject("Excel preview", payload);
  const zip = await JSZip.loadAsync(await readFile("tests/vault/docx-exploration/read-search-only.docx"));
  const path = "word/document.xml";
  const xml = await zip.file(path)!.async("string");
  const drawing = (name: string) => `<w:p><w:r><w:drawing><wp:inline
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:docPr id="1" name="${name}"/></wp:inline></w:drawing></w:r></w:p>`;
  zip.file(path, xml.replace("<w:sectPr/>", drawing("Other picture") + drawing("Excel preview") + "<w:sectPr/>"));
  const signal = new AbortController().signal;
  const restored = await restoreLegacyDocAttachments(source, await zip.generateAsync({ type: "arraybuffer" }), signal);
  const model = await inspectDocxPackage(restored, signal);
  expect(model.attachments).toHaveLength(1);
  const attachment = await extractDocxAttachment(restored, model.attachments![0]!);
  expect(attachment.extension).toBe("xlsx");
  expect(Array.from(attachment.bytes)).toEqual(Array.from(payload));
  const updatedXml = await (await JSZip.loadAsync(restored)).file(path)!.async("string");
  expect(updatedXml).toContain("Other picture");
  expect(updatedXml).not.toContain("Excel preview");
});
