import CFB from "cfb";
import JSZip from "jszip";
import { docxAttachmentExtension, readEmbeddedDocument, type DocxAttachmentExtension } from "./docx-attachments";
import { renderDocxAttachmentEmf } from "./docx-attachment-emf";
import { readLegacyDocObjectFields } from "./legacy-doc-object-fields";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DRAWING_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE = "urn:schemas-microsoft-com:office:office";
const TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
const PROGRAMS: Record<DocxAttachmentExtension, string> = {
  pdf: "AcroExch.Document", doc: "Word.Document.8", docx: "Word.Document.12",
  xls: "Excel.Sheet.8", xlsx: "Excel.Sheet.12",
  ppt: "PowerPoint.Show.8", pptx: "PowerPoint.Show.12",
};

function objectBytes(file: CFB.CFB$Container, storage: string): Uint8Array | undefined {
  const root = `${file.FullPaths[0]}ObjectPool/${storage}/`;
  const streams = file.FullPaths.flatMap((path, index) => {
    const entry = file.FileIndex[index];
    return entry !== undefined && path.startsWith(root) && path !== root
      ? [{ name: path.slice(root.length), entry }] : [];
  });
  if (streams.some(({ name, entry }) => entry.type !== 2 || name.includes("/") ||
    /^(?:VBA|Macros|_VBA_PROJECT_CUR|EncryptedPackage)$/i.test(name))) return;
  const packaged = streams.find(({ name }) => /^package$/i.test(name));
  if (packaged !== undefined) return new Uint8Array(packaged.entry.content);
  if (!streams.some(({ name }) =>
    /^(?:WordDocument|Workbook|Book|PowerPoint Document|CONTENTS|\x01Ole10Native)$/i.test(name))) return;
  const standalone = CFB.utils.cfb_new();
  for (const { name, entry } of streams) CFB.utils.cfb_add(standalone, name, entry.content);
  return new Uint8Array(CFB.write(standalone, { type: "buffer" }));
}

function legacyFields(file: CFB.CFB$Container) {
  const read = (name: string) => {
    const stream = CFB.find(file, `/${name}`);
    return stream?.type === 2 ? Buffer.from(stream.content) : Buffer.alloc(0);
  };
  const doc = read("WordDocument");
  if (doc.length < 34 || doc.readUInt16LE(0) !== 0xa5ec) return [];
  const flags = doc.readUInt16LE(10);
  if ((flags & 0x8100) !== 0) return [];
  return readLegacyDocObjectFields(doc, read(flags & 0x0200 ? "1Table" : "0Table"), read("Data"));
}

/** Restore file data beside the exact preview shape emitted by LibreOffice. */
export async function restoreLegacyDocAttachments(
  source: Uint8Array,
  converted: ArrayBuffer,
  signal: AbortSignal,
): Promise<ArrayBuffer> {
  const file = CFB.read(source, { type: "array" });
  if (!file.FullPaths.some((path) => /\/ObjectPool\/[^/]+\/$/i.test(path))) return converted;
  const fields = legacyFields(file);
  if (fields.length === 0) return converted;
  const zip = await JSZip.loadAsync(converted);
  const documentPart = zip.file("word/document.xml");
  const relationshipPart = zip.file("word/_rels/document.xml.rels");
  const typePart = zip.file("[Content_Types].xml");
  if (documentPart === null || relationshipPart === null || typePart === null) return converted;
  const parser = new DOMParser();
  const parse = async (part: typeof documentPart) =>
    parser.parseFromString(await part.async("text"), "application/xml");
  const [document, relationships, types] = await Promise.all([
    parse(documentPart), parse(relationshipPart), parse(typePart),
  ]);
  const body = document.getElementsByTagNameNS(WORD, "body")[0];
  if (body === undefined) return converted;
  const names = new Map<string, Element[]>();
  for (const drawing of body.getElementsByTagNameNS(WORD, "drawing")) {
    const name = drawing.getElementsByTagNameNS(DRAWING, "docPr")[0]?.getAttribute("name");
    if (!name) continue;
    const matches = names.get(name) ?? [];
    matches.push(drawing);
    names.set(name, matches);
  }
  const fieldCounts = new Map<string, number>();
  for (const { previewName } of fields) {
    if (previewName !== undefined) fieldCounts.set(previewName, (fieldCounts.get(previewName) ?? 0) + 1);
  }
  const rels = new Map(Array.from(relationships.getElementsByTagNameNS(PACKAGE_REL, "Relationship"),
    (relationship) => [relationship.getAttribute("Id"), relationship] as const));
  let added = 0;
  let generatedPng = false;
  for (const { storage, previewName } of fields) {
    signal.throwIfAborted();
    if (previewName === undefined || fieldCounts.get(previewName) !== 1) continue;
    const matches = names.get(previewName);
    const drawing = matches?.[0];
    if (matches?.length !== 1 || drawing === undefined || drawing.parentElement?.localName !== "r") continue;
    let attachment;
    try {
      const bytes = objectBytes(file, storage);
      if (bytes === undefined) continue;
      attachment = await readEmbeddedDocument(bytes);
    } catch {
      continue;
    }
    signal.throwIfAborted();
    let id: string;
    let target: string;
    do {
      added++;
      id = `officeViewerLegacyAttachment${added}`;
      target = `embeddings/office-viewer-legacy-${added}.bin`;
    } while (rels.has(id) || rels.has(`${id}Preview`) || zip.file(`word/${target}`) !== null ||
      zip.file(`word/media/office-viewer-legacy-preview-${added}.png`) !== null);
    const relation = relationships.createElementNS(PACKAGE_REL, "Relationship");
    relation.setAttribute("Id", id);
    relation.setAttribute("Type", `${REL}/oleObject`);
    relation.setAttribute("Target", target);
    relationships.documentElement.append(relation);
    zip.file(`word/${target}`, attachment.bytes);

    const object = document.createElementNS(WORD, "w:object");
    const ole = document.createElementNS(OFFICE, "o:OLEObject");
    ole.setAttribute("Type", "Embed");
    ole.setAttribute("ProgID", PROGRAMS[attachment.extension]);
    ole.setAttributeNS(REL, "r:id", id);
    object.append(ole);
    const blip = drawing.getElementsByTagNameNS(DRAWING_MAIN, "blip")[0];
    const imageId = blip?.getAttributeNS(REL, "embed");
    const image = rels.get(imageId ?? null);
    const imageTarget = image?.getAttribute("Type") === `${REL}/image`
      && image.getAttribute("TargetMode")?.toLowerCase() !== "external"
      ? image.getAttribute("Target") ?? "" : "";
    let raster = /^media\/[^/]+\.(?:png|jpe?g|gif|bmp)$/i.test(imageTarget);
    if (blip !== undefined && /^media\/[^/]+\.emf$/i.test(imageTarget)) {
      const part = zip.file(`word/${imageTarget}`);
      const preview = part === null ? null : await part.async("uint8array")
        .then((bytes) => renderDocxAttachmentEmf(bytes, globalThis.document))
        .catch(() => null);
      signal.throwIfAborted();
      if (preview !== null) {
        const previewId = `${id}Preview`;
        const previewTarget = `media/office-viewer-legacy-preview-${added}.png`;
        const previewRelationship = relationships.createElementNS(PACKAGE_REL, "Relationship");
        previewRelationship.setAttribute("Id", previewId);
        previewRelationship.setAttribute("Type", `${REL}/image`);
        previewRelationship.setAttribute("Target", previewTarget);
        relationships.documentElement.append(previewRelationship);
        zip.file(`word/${previewTarget}`, preview.png);
        blip.setAttributeNS(REL, "r:embed", previewId);
        if (docxAttachmentExtension(preview.caption) === attachment.extension) {
          object.setAttribute("officeViewerAttachmentName", preview.caption);
        }
        const paragraph = drawing.parentElement?.parentElement;
        if (paragraph?.localName === "p" &&
          paragraph.getElementsByTagNameNS(WORD, "drawing").length === 1) {
          object.setAttribute("officeViewerAttachmentCentered", "true");
        }
        raster = generatedPng = true;
      }
    }
    drawing.replaceWith(object);
    if (raster) object.append(drawing);
  }
  if (added === 0) return converted;
  const defaults = Array.from(types.getElementsByTagNameNS(TYPES, "Default"));
  const addDefault = (extension: string, contentType: string) => {
    if (defaults.some((entry) => entry.getAttribute("Extension")?.toLowerCase() === extension)) return;
    const type = types.createElementNS(TYPES, "Default");
    type.setAttribute("Extension", extension);
    type.setAttribute("ContentType", contentType);
    types.documentElement.append(type);
  };
  addDefault("bin", "application/vnd.openxmlformats-officedocument.oleObject");
  if (generatedPng) addDefault("png", "image/png");
  const serializer = new XMLSerializer();
  zip.file(documentPart.name, serializer.serializeToString(document));
  zip.file(relationshipPart.name, serializer.serializeToString(relationships));
  zip.file(typePart.name, serializer.serializeToString(types));
  const result = await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  signal.throwIfAborted();
  return result;
}
