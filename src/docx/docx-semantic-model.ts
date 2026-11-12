import CFB from "cfb";
import JSZip, { type JSZipObject } from "jszip";
import { convertWindowsMetafileToPng } from "./convert-windows-metafile";
import { DocxOpenError } from "./docx-open-error";
import { materializeDocxChartsAsImages } from "./materialize-docx-charts";

const WORDPROCESSING_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MATH_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/math";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const IMAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const ALT_CHUNK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
const OLE_SIGNATURE = [
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
] as const;
const ZIP_SIGNATURE = [0x50, 0x4b, 0x03, 0x04] as const;
const REQUIRED_PARTS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "word/document.xml",
] as const;

export const DOCX_ZIP_LIMITS = Object.freeze({
  maxEntries: 4_000,
  maxEntryUncompressedBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxMediaBytes: 192 * 1024 * 1024,
  maxXmlPartBytes: 8 * 1024 * 1024,
  maxXmlBytes: 32 * 1024 * 1024,
});

type SizedZipObject = JSZipObject & {
  _data?: { uncompressedSize?: number };
};

export interface DocxExternalHyperlink {
  readonly kind: "external";
  readonly label: string;
  readonly target: string;
}

export interface DocxBookmarkHyperlink {
  readonly kind: "bookmark";
  readonly label: string;
  readonly bookmark: string;
}

export interface DocxBlockedHyperlink {
  readonly kind: "blocked";
  readonly label: string;
  readonly target: string;
}

export type DocxHyperlink =
  | DocxExternalHyperlink
  | DocxBookmarkHyperlink
  | DocxBlockedHyperlink;

export type DocxUnavailableContentKind =
  | "equation"
  | "embedded-object"
  | "alt-chunk"
  | "external-image";

function isChartGraphicElement(element: Element): boolean {
  return element.localName === "chart" || element.localName === "chartEx";
}

export interface DocxSemanticParagraph {
  readonly ordinal: number;
  readonly text: string;
  readonly searchText: string;
  readonly styleId: string | null;
  readonly listItem: boolean;
  readonly tableDepth: number;
  readonly bookmarks: readonly string[];
  readonly hyperlinks: readonly DocxHyperlink[];
  readonly inlineImageCount: number;
  readonly unavailableContent: readonly DocxUnavailableContentKind[];
}

export interface DocxBookmarkTarget {
  readonly bookmark: string;
  readonly paragraphOrdinal: number;
}

export interface DocxUnavailableBodyBlock {
  readonly afterParagraphOrdinal: number;
  readonly kinds: readonly DocxUnavailableContentKind[];
}

export interface DocxSemanticModel {
  readonly paragraphs: readonly DocxSemanticParagraph[];
  readonly bookmarkTargets: readonly DocxBookmarkTarget[];
  readonly unavailableBodyBlocks: readonly DocxUnavailableBodyBlock[];
  readonly hasUnavailableBodyContent: boolean;
}

interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

function malformed(message: string, cause?: unknown): DocxOpenError {
  return new DocxOpenError("malformed", message, { cause });
}

function incompatible(message: string, cause?: unknown): DocxOpenError {
  return new DocxOpenError("incompatible", message, { cause });
}

function exhausted(message: string): DocxOpenError {
  return new DocxOpenError("resource-exhausted", message);
}

function hasSignature(
  bytes: Uint8Array,
  signature: readonly number[],
): boolean {
  return signature.every((byte, index) => bytes[index] === byte);
}

function uncompressedSize(part: JSZipObject): number {
  const size = (part as SizedZipObject)._data?.uncompressedSize;
  if (!Number.isSafeInteger(size) || size === undefined || size < 0) {
    throw malformed(`Invalid ZIP size metadata for ${part.name}`);
  }
  return size;
}

function enforceZipLimits(zip: JSZip): void {
  const entries = Object.values(zip.files).filter((part) => !part.dir);
  if (entries.length > DOCX_ZIP_LIMITS.maxEntries) {
    throw exhausted("The DOCX package contains too many ZIP entries");
  }
  let totalBytes = 0;
  let mediaBytes = 0;
  let xmlBytes = 0;
  for (const part of entries) {
    const size = uncompressedSize(part);
    if (size > DOCX_ZIP_LIMITS.maxEntryUncompressedBytes) {
      throw exhausted(`OOXML part ${part.name} exceeds the safe size limit`);
    }
    totalBytes += size;
    if (part.name.startsWith("word/media/")) mediaBytes += size;
    if (part.name.endsWith(".xml") || part.name.endsWith(".rels")) {
      if (size > DOCX_ZIP_LIMITS.maxXmlPartBytes) {
        throw exhausted(`OOXML part ${part.name} exceeds the XML size limit`);
      }
      xmlBytes += size;
    }
  }
  if (totalBytes > DOCX_ZIP_LIMITS.maxTotalUncompressedBytes) {
    throw exhausted("The DOCX package exceeds the expanded-size limit");
  }
  if (mediaBytes > DOCX_ZIP_LIMITS.maxMediaBytes) {
    throw exhausted("The DOCX package exceeds the media-size limit");
  }
  if (xmlBytes > DOCX_ZIP_LIMITS.maxXmlBytes) {
    throw exhausted("The DOCX package contains too much XML to inspect safely");
  }
}

function compoundStreamBytes(
  container: CFB.CFB$Container,
  name: string,
): Uint8Array | null {
  const entry = CFB.find(container, name);
  return entry?.type === 2 ? new Uint8Array(entry.content) : null;
}

function isEncryptedOoxmlCompoundFile(bytes: Uint8Array): boolean {
  let container: CFB.CFB$Container;
  try {
    container = CFB.parse(bytes, { WTF: true });
  } catch (error) {
    throw malformed("Unable to parse compound-file input", error);
  }
  const encryptionInfo = compoundStreamBytes(container, "EncryptionInfo");
  const encryptedPackage = compoundStreamBytes(container, "EncryptedPackage");
  return (
    encryptionInfo !== null &&
    encryptionInfo.byteLength >= 8 &&
    encryptedPackage !== null &&
    encryptedPackage.byteLength >= 16
  );
}

async function readXml(
  part: JSZipObject,
  signal: AbortSignal,
): Promise<Document> {
  signal.throwIfAborted();
  let xml: string;
  try {
    xml = await part.async("text");
  } catch (error) {
    throw malformed(`Unable to read OOXML part ${part.name}`, error);
  }
  signal.throwIfAborted();
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.getElementsByTagName("parsererror").length > 0) {
    throw malformed(`Invalid XML in OOXML part ${part.name}`);
  }
  return document;
}

function relationshipMap(document: Document): ReadonlyMap<string, Relationship> {
  const relationships = new Map<string, Relationship>();
  for (const element of document.getElementsByTagNameNS("*", "Relationship")) {
    const id = element.getAttribute("Id");
    const type = element.getAttribute("Type");
    const target = element.getAttribute("Target");
    if (id === null || type === null || target === null) {
      throw malformed("DOCX relationship is missing Id, Type, or Target");
    }
    if (relationships.has(id)) {
      throw malformed(`DOCX relationship contains duplicate Id ${id}`);
    }
    relationships.set(id, {
      id,
      type,
      target,
      external:
        element.getAttribute("TargetMode")?.toLowerCase() === "external",
    });
  }
  return relationships;
}

function isDescendantOf(
  element: Element,
  namespace: string,
  localName: string,
): boolean {
  let parent = element.parentElement;
  while (parent !== null) {
    if (
      parent.namespaceURI === namespace &&
      parent.localName === localName
    ) {
      return true;
    }
    parent = parent.parentElement;
  }
  return false;
}

function nearestAncestor(
  element: Element,
  namespace: string,
  localName: string,
): Element | null {
  let parent = element.parentElement;
  while (parent !== null) {
    if (
      parent.namespaceURI === namespace &&
      parent.localName === localName
    ) {
      return parent;
    }
    parent = parent.parentElement;
  }
  return null;
}

function closestElement(
  element: Element,
  namespace: string,
  localName: string,
): Element | null {
  let candidate: Element | null = element;
  while (candidate !== null) {
    if (
      candidate.namespaceURI === namespace &&
      candidate.localName === localName
    ) {
      return candidate;
    }
    candidate = candidate.parentElement;
  }
  return null;
}

function runIsHidden(element: Element): boolean {
  const run = nearestAncestor(element, WORDPROCESSING_NAMESPACE, "r");
  if (run === null) return false;
  for (const property of run.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "rPr",
  )) {
    if (property.parentElement !== run) continue;
    if (
      property.getElementsByTagNameNS(
        WORDPROCESSING_NAMESPACE,
        "vanish",
      ).length > 0 ||
      property.getElementsByTagNameNS(
        WORDPROCESSING_NAMESPACE,
        "webHidden",
      ).length > 0
    ) {
      return true;
    }
  }
  return false;
}

function paragraphText(paragraph: Element): string {
  let text = "";
  const append = (node: Node): void => {
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const element = node as Element;
    if (
      element.namespaceURI === WORDPROCESSING_NAMESPACE &&
      (element.localName === "del" ||
        element.localName === "txbxContent" ||
        element.localName === "instrText" ||
        element.localName === "delText")
    ) {
      return;
    }
    if (
      element.namespaceURI === WORDPROCESSING_NAMESPACE &&
      element.localName === "t"
    ) {
      if (!runIsHidden(element)) text += element.textContent ?? "";
      return;
    }
    if (
      element.namespaceURI === WORDPROCESSING_NAMESPACE &&
      element.localName === "tab"
    ) {
      if (!runIsHidden(element)) text += "\t";
      return;
    }
    if (
      element.namespaceURI === WORDPROCESSING_NAMESPACE &&
      (element.localName === "br" || element.localName === "cr")
    ) {
      if (!runIsHidden(element)) text += "\n";
      return;
    }
    for (const child of element.childNodes) append(child);
  };
  for (const child of paragraph.childNodes) append(child);
  return text;
}

function textOfElement(element: Element): string {
  let text = "";
  for (const descendant of element.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "t",
  )) {
    if (
      !isDescendantOf(descendant, WORDPROCESSING_NAMESPACE, "del") &&
      !isDescendantOf(descendant, WORDPROCESSING_NAMESPACE, "txbxContent") &&
      !runIsHidden(descendant)
    ) {
      text += descendant.textContent ?? "";
    }
  }
  return text;
}

function allowedExternalTarget(target: string): boolean {
  let url: URL;
  try {
    url = new URL(target);
  } catch {
    return false;
  }
  return (
    url.protocol === "https:" ||
    url.protocol === "http:" ||
    url.protocol === "mailto:"
  );
}

function paragraphHyperlinks(
  paragraph: Element,
  relationships: ReadonlyMap<string, Relationship>,
): readonly DocxHyperlink[] {
  const hyperlinks: DocxHyperlink[] = [];
  for (const element of paragraph.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "hyperlink",
  )) {
    if (isDescendantOf(element, WORDPROCESSING_NAMESPACE, "txbxContent")) {
      continue;
    }
    const label = textOfElement(element);
    const anchor = element.getAttributeNS(
      WORDPROCESSING_NAMESPACE,
      "anchor",
    );
    if (anchor !== null) {
      hyperlinks.push({ kind: "bookmark", label, bookmark: anchor });
      continue;
    }
    const relationshipId = element.getAttributeNS(
      OFFICE_RELATIONSHIP_NAMESPACE,
      "id",
    );
    if (relationshipId === null) continue;
    const relationship = relationships.get(relationshipId);
    if (
      relationship === undefined ||
      relationship.type !== HYPERLINK_RELATIONSHIP ||
      !relationship.external
    ) {
      hyperlinks.push({ kind: "blocked", label, target: "" });
      continue;
    }
    hyperlinks.push(
      allowedExternalTarget(relationship.target)
        ? { kind: "external", label, target: relationship.target }
        : { kind: "blocked", label, target: relationship.target },
    );
  }
  return hyperlinks;
}

function paragraphBookmarks(paragraph: Element): readonly string[] {
  const bookmarks = new Set<string>();
  for (const element of paragraph.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "bookmarkStart",
  )) {
    if (isDescendantOf(element, WORDPROCESSING_NAMESPACE, "txbxContent")) {
      continue;
    }
    const name = element.getAttributeNS(WORDPROCESSING_NAMESPACE, "name");
    if (name !== null && name.length > 0 && !name.startsWith("_")) {
      bookmarks.add(name);
    }
  }
  return [...bookmarks];
}

function unavailableContent(
  paragraph: Element,
  relationships: ReadonlyMap<string, Relationship>,
): readonly DocxUnavailableContentKind[] {
  const kinds = new Set<DocxUnavailableContentKind>();
  if (
    paragraph.getElementsByTagNameNS(MATH_NAMESPACE, "oMath").length > 0 ||
    paragraph.getElementsByTagNameNS(MATH_NAMESPACE, "oMathPara").length > 0
  ) {
    kinds.add("equation");
  }
  if (
    paragraph.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "object",
    ).length > 0 ||
    paragraph.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "pict",
    ).length > 0
  ) {
    kinds.add("embedded-object");
  }
  if (
    paragraph.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "altChunk",
    ).length > 0
  ) {
    kinds.add("alt-chunk");
  }
  for (const element of paragraph.getElementsByTagName("*")) {
    for (const attribute of Array.from(element.attributes)) {
      if (
        attribute.namespaceURI !== OFFICE_RELATIONSHIP_NAMESPACE ||
        (attribute.localName !== "embed" && attribute.localName !== "link")
      ) {
        continue;
      }
      const relationship = relationships.get(attribute.value);
      if (
        relationship?.type === IMAGE_RELATIONSHIP &&
        relationship.external
      ) {
        kinds.add("external-image");
      }
    }
  }
  return [...kinds];
}

function inlineImageCount(
  paragraph: Element,
  relationships: ReadonlyMap<string, Relationship>,
): number {
  let count = 0;
  for (const element of paragraph.getElementsByTagName("*")) {
    for (const attribute of Array.from(element.attributes)) {
      if (
        attribute.namespaceURI !== OFFICE_RELATIONSHIP_NAMESPACE ||
        attribute.localName !== "embed"
      ) {
        continue;
      }
      const relationship = relationships.get(attribute.value);
      if (
        relationship?.type === IMAGE_RELATIONSHIP &&
        !relationship.external
      ) {
        count += 1;
      }
    }
  }
  return count;
}

function tableDepth(paragraph: Element): number {
  let depth = 0;
  let parent = paragraph.parentElement;
  while (parent !== null) {
    if (
      parent.namespaceURI === WORDPROCESSING_NAMESPACE &&
      parent.localName === "tbl"
    ) {
      depth += 1;
    }
    parent = parent.parentElement;
  }
  return depth;
}

function paragraphStyleId(paragraph: Element): string | null {
  for (const properties of paragraph.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "pPr",
  )) {
    if (properties.parentElement !== paragraph) continue;
    for (const style of properties.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "pStyle",
    )) {
      return style.getAttributeNS(WORDPROCESSING_NAMESPACE, "val");
    }
  }
  return null;
}

function paragraphIsListItem(paragraph: Element): boolean {
  for (const properties of paragraph.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "pPr",
  )) {
    if (properties.parentElement !== paragraph) continue;
    if (
      properties.getElementsByTagNameNS(
        WORDPROCESSING_NAMESPACE,
        "numPr",
      ).length > 0
    ) {
      return true;
    }
  }
  return false;
}

function rejectActiveContent(zip: JSZip, contentTypes: Document): void {
  const names = Object.values(zip.files)
    .filter((part) => !part.dir)
    .map((part) => part.name.toLowerCase());
  const types = Array.from(
    contentTypes.getElementsByTagNameNS("*", "Override"),
    (element) => element.getAttribute("ContentType")?.toLowerCase() ?? "",
  );
  if (
    names.some(
      (name) =>
        name.endsWith("/vbaproject.bin") ||
        name.startsWith("word/activex/"),
    ) ||
    types.some(
      (type) =>
        type.includes("macroenabled") ||
        type.includes("vbaproject") ||
        type.includes("activex"),
    )
  ) {
    throw incompatible("The DOCX package contains active content");
  }
}

function verifyExternalRelationships(
  relationships: ReadonlyMap<string, Relationship>,
  document: Document,
): void {
  const hyperlinkIds = new Set<string>();
  for (const hyperlink of document.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "hyperlink",
  )) {
    const id = hyperlink.getAttributeNS(OFFICE_RELATIONSHIP_NAMESPACE, "id");
    if (id !== null) hyperlinkIds.add(id);
  }
  for (const relationship of relationships.values()) {
    if (!relationship.external) continue;
    if (
      relationship.type === HYPERLINK_RELATIONSHIP &&
      hyperlinkIds.has(relationship.id)
    ) {
      continue;
    }
    if (
      relationship.type === IMAGE_RELATIONSHIP ||
      relationship.type === ALT_CHUNK_RELATIONSHIP
    ) {
      continue;
    }
    throw incompatible(
      `Unsupported external DOCX relationship ${relationship.id}`,
    );
  }
}

export async function inspectDocxPackage(
  buffer: ArrayBuffer,
  signal: AbortSignal,
): Promise<DocxSemanticModel> {
  signal.throwIfAborted();
  const bytes = new Uint8Array(buffer);
  if (hasSignature(bytes, OLE_SIGNATURE)) {
    if (isEncryptedOoxmlCompoundFile(bytes)) {
      throw new DocxOpenError(
        "protected",
        "Encrypted or protected OOXML compound-file container",
      );
    }
    throw malformed("Compound-file input is not a DOCX package");
  }
  if (!hasSignature(bytes, ZIP_SIGNATURE)) {
    throw malformed("DOCX input is not a ZIP-based OOXML package");
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (error) {
    throw malformed("Unable to parse the DOCX ZIP container", error);
  }
  signal.throwIfAborted();
  enforceZipLimits(zip);
  for (const path of REQUIRED_PARTS) {
    if (zip.file(path) === null) {
      throw malformed(`Missing required OOXML part ${path}`);
    }
  }

  const documentPart = zip.file("word/document.xml");
  const contentTypesPart = zip.file("[Content_Types].xml");
  if (documentPart === null || contentTypesPart === null) {
    throw malformed("Missing required DOCX package parts");
  }
  const [document, contentTypes] = await Promise.all([
    readXml(documentPart, signal),
    readXml(contentTypesPart, signal),
  ]);
  const relationshipsPart = zip.file("word/_rels/document.xml.rels");
  const relationships =
    relationshipsPart === null
      ? new Map<string, Relationship>()
      : relationshipMap(await readXml(relationshipsPart, signal));
  rejectActiveContent(zip, contentTypes);
  verifyExternalRelationships(relationships, document);

  const body = document.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "body",
  )[0];
  if (body === undefined) throw malformed("DOCX main document has no body");

  const paragraphs: DocxSemanticParagraph[] = [];
  const unavailableBodyBlocks: DocxUnavailableBodyBlock[] = [];
  for (const paragraph of body.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "p",
  )) {
    if (
      isDescendantOf(paragraph, WORDPROCESSING_NAMESPACE, "txbxContent") ||
      isDescendantOf(paragraph, WORDPROCESSING_NAMESPACE, "del")
    ) {
      continue;
    }
    const text = paragraphText(paragraph);
    const unavailable = unavailableContent(paragraph, relationships);
    if (text.trim().length === 0) {
      if (unavailable.length > 0) {
        unavailableBodyBlocks.push({
          afterParagraphOrdinal: paragraphs.length,
          kinds: unavailable,
        });
      }
      continue;
    }
    paragraphs.push({
      ordinal: paragraphs.length + 1,
      text,
      searchText: text.toLocaleLowerCase(),
      styleId: paragraphStyleId(paragraph),
      listItem: paragraphIsListItem(paragraph),
      tableDepth: tableDepth(paragraph),
      bookmarks: paragraphBookmarks(paragraph),
      hyperlinks: paragraphHyperlinks(paragraph, relationships),
      inlineImageCount: inlineImageCount(paragraph, relationships),
      unavailableContent: unavailable,
    });
  }

  for (const altChunk of body.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "altChunk",
  )) {
    if (
      nearestAncestor(altChunk, WORDPROCESSING_NAMESPACE, "p") === null
    ) {
      unavailableBodyBlocks.push({
        afterParagraphOrdinal: paragraphs.length,
        kinds: ["alt-chunk"],
      });
    }
  }

  const bookmarkOwners = new Map<string, number | null>();
  for (const paragraph of paragraphs) {
    for (const bookmark of paragraph.bookmarks) {
      bookmarkOwners.set(
        bookmark,
        bookmarkOwners.has(bookmark) ? null : paragraph.ordinal,
      );
    }
  }
  const bookmarkTargets: DocxBookmarkTarget[] = [];
  for (const [bookmark, paragraphOrdinal] of bookmarkOwners) {
    if (paragraphOrdinal !== null) {
      bookmarkTargets.push({ bookmark, paragraphOrdinal });
    }
  }

  return {
    paragraphs,
    bookmarkTargets,
    unavailableBodyBlocks,
    hasUnavailableBodyContent:
      unavailableBodyBlocks.length > 0 ||
      paragraphs.some(
        (paragraph) => paragraph.unavailableContent.length > 0,
      ),
  };
}

async function convertEmbeddedWindowsMetafiles(
  zip: JSZip,
  signal: AbortSignal,
): Promise<void> {
  const renames: Array<{ readonly from: string; readonly to: string }> = [];
  for (const [path, part] of Object.entries(zip.files)) {
    if (part.dir || !/^word\/media\/.+\.(?:wmf|emf)$/i.test(path)) continue;
    signal.throwIfAborted();
    const bytes = new Uint8Array(await part.async("arraybuffer"));
    const png = convertWindowsMetafileToPng(bytes);
    if (png === null) continue;
    const pngPath = path.replace(/\.(?:wmf|emf)$/i, ".png");
    zip.remove(path);
    zip.file(pngPath, png);
    renames.push({
      from: path.slice("word/".length),
      to: pngPath.slice("word/".length),
    });
  }
  if (renames.length === 0) return;

  for (const [path, part] of Object.entries(zip.files)) {
    if (part.dir || !path.endsWith(".rels")) continue;
    let text = await part.async("string");
    let changed = false;
    for (const { from, to } of renames) {
      if (!text.includes(from)) continue;
      text = text.split(from).join(to);
      changed = true;
    }
    if (changed) zip.file(path, text);
  }

  const contentTypesPart = zip.file("[Content_Types].xml");
  if (contentTypesPart === null) return;
  let contentTypes = await contentTypesPart.async("string");
  if (!/Extension="png"/i.test(contentTypes)) {
    contentTypes = contentTypes.replace(
      /<Types([^>]*)>/,
      '<Types$1><Default Extension="png" ContentType="image/png"/>',
    );
    zip.file("[Content_Types].xml", contentTypes);
  }
}

export async function createSafeDocxRendererBuffer(
  buffer: ArrayBuffer,
  signal: AbortSignal,
  prepareLayout?: (document: Document) => void,
): Promise<ArrayBuffer> {
  signal.throwIfAborted();
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (error) {
    throw malformed("Unable to prepare the DOCX renderer input", error);
  }
  await convertEmbeddedWindowsMetafiles(zip, signal);
  const documentPart = zip.file("word/document.xml");
  if (documentPart === null) throw malformed("Missing DOCX main document");
  const document = await readXml(documentPart, signal);
  const relationshipsPart = zip.file("word/_rels/document.xml.rels");
  const externalImageIds = new Set<string>();
  const removableRelationshipIds = new Set<string>();
  let relationshipDocument: Document | null = null;
  if (relationshipsPart !== null) {
    relationshipDocument = await readXml(relationshipsPart, signal);
    for (const relationship of relationshipDocument.getElementsByTagNameNS(
      "*",
      "Relationship",
    )) {
      const id = relationship.getAttribute("Id");
      const type = relationship.getAttribute("Type");
      const external =
        relationship.getAttribute("TargetMode")?.toLowerCase() === "external";
      if (id === null || !external) continue;
      if (type === IMAGE_RELATIONSHIP) {
        externalImageIds.add(id);
        removableRelationshipIds.add(id);
      } else if (type === ALT_CHUNK_RELATIONSHIP) {
        removableRelationshipIds.add(id);
      }
    }
    for (const relationship of Array.from(
      relationshipDocument.getElementsByTagNameNS("*", "Relationship"),
    )) {
      const id = relationship.getAttribute("Id");
      if (id !== null && removableRelationshipIds.has(id)) {
        relationship.remove();
      }
    }
  }

  for (const deletion of Array.from(
    document.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "del"),
  )) {
    deletion.remove();
  }
  for (const textBox of Array.from(
    document.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "txbxContent",
    ),
  )) {
    textBox.remove();
  }
  for (const chunk of Array.from(
    document.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "altChunk",
    ),
  )) {
    chunk.remove();
  }
  for (const object of [
    ...Array.from(
      document.getElementsByTagNameNS(
        WORDPROCESSING_NAMESPACE,
        "object",
      ),
    ),
    ...Array.from(
      document.getElementsByTagNameNS(
        WORDPROCESSING_NAMESPACE,
        "pict",
      ),
    ),
    ...Array.from(document.getElementsByTagNameNS(MATH_NAMESPACE, "oMath")),
    ...Array.from(
      document.getElementsByTagNameNS(MATH_NAMESPACE, "oMathPara"),
    ),
  ]) {
    object.remove();
  }
  for (const run of Array.from(
    document.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "r"),
  )) {
    let hidden = false;
    for (const properties of run.getElementsByTagNameNS(
      WORDPROCESSING_NAMESPACE,
      "rPr",
    )) {
      if (properties.parentElement !== run) continue;
      hidden =
        properties.getElementsByTagNameNS(
          WORDPROCESSING_NAMESPACE,
          "vanish",
        ).length > 0 ||
        properties.getElementsByTagNameNS(
          WORDPROCESSING_NAMESPACE,
          "webHidden",
        ).length > 0;
      if (hidden) break;
    }
    if (hidden) run.remove();
  }
  if (externalImageIds.size > 0) {
    for (const element of Array.from(document.getElementsByTagName("*"))) {
      const relationshipId =
        element.getAttributeNS(OFFICE_RELATIONSHIP_NAMESPACE, "embed") ??
        element.getAttributeNS(OFFICE_RELATIONSHIP_NAMESPACE, "link");
      if (
        relationshipId !== null &&
        externalImageIds.has(relationshipId)
      ) {
        (
          closestElement(element, WORDPROCESSING_NAMESPACE, "drawing") ??
          element
        ).remove();
      }
    }
  }
  // docx-preview cannot paint DrawingML charts. Rasterize cached chart series
  // into PNG pictures first; strip any leftover chart drawings.
  await materializeDocxChartsAsImages(
    zip,
    document,
    relationshipDocument,
    signal,
  );
  for (const element of Array.from(document.getElementsByTagName("*"))) {
    if (!isChartGraphicElement(element)) continue;
    (
      closestElement(element, WORDPROCESSING_NAMESPACE, "drawing") ?? element
    ).remove();
  }
  signal.throwIfAborted();
  prepareLayout?.(document);
  if (relationshipsPart !== null && relationshipDocument !== null) {
    zip.file(
      relationshipsPart.name,
      new XMLSerializer().serializeToString(relationshipDocument),
    );
  }
  zip.file(
    documentPart.name,
    new XMLSerializer().serializeToString(document),
  );
  const rendererBuffer = await zip.generateAsync({ type: "arraybuffer" });
  signal.throwIfAborted();
  return rendererBuffer;
}

export interface DocxSearchResult {
  readonly paragraphOrdinal: number;
  readonly matchCount: number;
  readonly text: string;
}

export function searchDocxBody(
  model: DocxSemanticModel,
  query: string,
): readonly DocxSearchResult[] {
  const normalizedQuery = query.trim().toLocaleLowerCase();
  if (normalizedQuery.length === 0) return [];
  const results: DocxSearchResult[] = [];
  for (const paragraph of model.paragraphs) {
    const text = paragraph.searchText;
    let count = 0;
    let offset = 0;
    while (true) {
      const match = text.indexOf(normalizedQuery, offset);
      if (match < 0) break;
      count += 1;
      offset = match + Math.max(normalizedQuery.length, 1);
    }
    if (count > 0) {
      results.push({
        paragraphOrdinal: paragraph.ordinal,
        matchCount: count,
        text: paragraph.text,
      });
    }
  }
  return results;
}
