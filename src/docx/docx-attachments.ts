import CFB from "cfb";
import JSZip, { type JSZipObject } from "jszip";

const OOXML_TYPES = {
  docx: ["word/document.xml", "wordprocessingml.document"],
  xlsx: ["xl/workbook.xml", "spreadsheetml.sheet"],
  pptx: ["ppt/presentation.xml", "presentationml.presentation"],
} as const;
const LEGACY_STREAMS = {
  doc: ["WordDocument"],
  xls: ["Workbook", "Book"],
  ppt: ["PowerPoint Document"],
} as const;
const MAX_EMBEDDED_BYTES = 32 * 1024 * 1024;
const COMPOUND_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export type DocxAttachmentExtension = "pdf" | keyof typeof OOXML_TYPES | keyof typeof LEGACY_STREAMS;
export interface DocxAttachment {
  readonly objectIndex: number;
  readonly partPath: string;
  readonly marker: string;
  readonly name: string | null;
  readonly extension: DocxAttachmentExtension | null;
  readonly afterParagraphOrdinal: number;
  readonly hasPreview?: boolean;
  readonly centerPreview?: boolean;
}

export interface DocxAttachmentFile {
  readonly name: string;
  readonly extension: DocxAttachmentExtension;
  readonly bytes: Uint8Array;
}

export class UnsupportedDocxAttachment extends Error {}

const EXTENSIONS = new Set<string>(["pdf", ...Object.keys(OOXML_TYPES), ...Object.keys(LEGACY_STREAMS)]);

export function docxAttachmentExtension(name: string): DocxAttachmentExtension | null {
  const extension = /\.([a-z0-9]+)$/i.exec(name)?.[1]?.toLowerCase();
  return extension !== undefined && EXTENSIONS.has(extension)
    ? extension as DocxAttachmentExtension
    : null;
}

export function docxAttachmentPart(zip: JSZip, path: string): JSZipObject | null {
  const entry = zip.file(path);
  const size = (entry as (JSZipObject & { _data?: { uncompressedSize?: number } }) | null)
    ?._data?.uncompressedSize;
  return size !== undefined && Number.isSafeInteger(size) && size > 0 && size <= MAX_EMBEDDED_BYTES
    ? entry
    : null;
}

function startsWith(bytes: Uint8Array, signature: readonly number[]): boolean {
  return signature.every((byte, index) => bytes[index] === byte);
}

function rootStream(file: CFB.CFB$Container, name: string): Uint8Array | undefined {
  const entry = CFB.find(file, `/${name}`);
  return entry?.type === 2 ? new Uint8Array(entry.content) : undefined;
}

// The native size prefix and Packager record fields follow MS-OLEDS and Apache POI.
function unpackOleNative(stream: Uint8Array): { bytes: Uint8Array; name?: string } | undefined {
  if (stream.length < 6) return;
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const end = 4 + view.getUint32(0, true);
  if (end > stream.length) return;
  if (view.getUint16(4, true) !== 2) return { bytes: stream.subarray(4, end) };
  const record = stream.subarray(0, end);
  const nameEnd = record.indexOf(0, 6);
  const sourceEnd = nameEnd < 0 ? -1 : record.indexOf(0, nameEnd + 1);
  if (nameEnd <= 6 || nameEnd - 6 > 4096 || sourceEnd < 0 || sourceEnd + 9 > end) return;
  const commandLength = view.getUint32(sourceEnd + 5, true);
  const sizeOffset = sourceEnd + 9 + commandLength;
  if (sizeOffset + 4 > end) return;
  const length = view.getUint32(sizeOffset, true);
  if (length === 0 || sizeOffset + 4 + length > end) return;
  return {
    name: new TextDecoder().decode(record.subarray(6, nameEnd)),
    bytes: record.subarray(sizeOffset + 4, sizeOffset + 4 + length),
  };
}

function unpackOle(bytes: Uint8Array): { bytes: Uint8Array; name?: string } {
  const file = CFB.read(bytes, { type: "array" });
  const packageStream = rootStream(file, "Package");
  if (packageStream !== undefined) return { bytes: packageStream };
  const native = rootStream(file, "\x01Ole10Native");
  if (native !== undefined) return unpackOleNative(native) ?? { bytes };
  const contents = rootStream(file, "CONTENTS");
  return contents === undefined ? { bytes } : { bytes: contents };
}

async function detectDocument(bytes: Uint8Array): Promise<DocxAttachmentExtension | null> {
  if (startsWith(bytes, [0x25, 0x50, 0x44, 0x46, 0x2d])) return "pdf";
  if (startsWith(bytes, [0x50, 0x4b, 0x03, 0x04])) {
    const zip = await JSZip.loadAsync(bytes);
    const entries = Object.keys(zip.files);
    if (entries.length > 4000 || entries.some(
      (path) => /(?:^|\/)(?:vbaproject\.bin|activex|macrosheets)(?:\/|$)/i.test(path),
    )) return null;
    const typesPart = docxAttachmentPart(zip, "[Content_Types].xml");
    if (typesPart === null) return null;
    const types = await typesPart.async("text");
    if (types.length > 1024 * 1024 || /macroenabled|vbaproject|activex|macrosheet/i.test(types)) return null;
    const xml = new DOMParser().parseFromString(types, "application/xml");
    const overrides = Array.from(xml.getElementsByTagNameNS(
      "http://schemas.openxmlformats.org/package/2006/content-types", "Override",
    ));
    for (const extension of Object.keys(OOXML_TYPES) as (keyof typeof OOXML_TYPES)[]) {
      const [part, type] = OOXML_TYPES[extension];
      if (zip.file(part) !== null && overrides.some((element) =>
        element.getAttribute("PartName") === `/${part}` &&
        element.getAttribute("ContentType") === `application/vnd.openxmlformats-officedocument.${type}.main+xml`,
      )) return extension;
    }
  }
  if (startsWith(bytes, COMPOUND_SIGNATURE)) {
    const file = CFB.read(bytes, { type: "array" });
    if (file.FullPaths.some((path) => /(?:^|\/)(?:VBA|Macros|_VBA_PROJECT_CUR|ObjectPool|EncryptedPackage)(?:\/|$)/i.test(path))) {
      return null;
    }
    for (const extension of Object.keys(LEGACY_STREAMS) as (keyof typeof LEGACY_STREAMS)[]) {
      if (LEGACY_STREAMS[extension].some((name) => rootStream(file, name) !== undefined)) return extension;
    }
  }
  return null;
}

export async function extractDocxAttachment(
  buffer: ArrayBuffer,
  attachment: DocxAttachment,
): Promise<DocxAttachmentFile> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const part = docxAttachmentPart(zip, attachment.partPath);
    if (part === null) throw new UnsupportedDocxAttachment();
    const extracted = await readEmbeddedDocument(await part.async("uint8array"), attachment.name);
    if (attachment.name !== null && attachment.extension !== extracted.extension) {
      throw new UnsupportedDocxAttachment();
    }
    return extracted;
  } catch (error) {
    throw new UnsupportedDocxAttachment("The object does not contain a supported document", { cause: error });
  }
}

export async function readEmbeddedDocument(
  original: Uint8Array,
  name: string | null = null,
): Promise<DocxAttachmentFile> {
  if (original.byteLength === 0 || original.byteLength > MAX_EMBEDDED_BYTES) {
    throw new UnsupportedDocxAttachment();
  }
  const extracted = startsWith(original, COMPOUND_SIGNATURE) ? unpackOle(original) : { bytes: original };
  const extension = await detectDocument(extracted.bytes);
  if (extension === null) throw new UnsupportedDocxAttachment();
  return {
    name: extracted.name ?? name ?? `attachment.${extension}`,
    extension,
    bytes: extracted.bytes,
  };
}
