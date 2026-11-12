import CFB from "cfb";
import JSZip, { type JSZipObject } from "jszip";
import { convertWindowsMetafileToPng } from "./convert-windows-metafile";
import { DocxOpenError } from "./docx-open-error";
import {
  canRenderDocxBlankLine,
  createDocxParagraphLayoutResolver,
  type DocxParagraphIndents,
} from "./docx-paragraph-layout";
import { createDocxNumbering } from "./docx-numbering";
import { docxAttachmentExtension, docxAttachmentPart, type DocxAttachment } from "./docx-attachments";
import {
  createDocxHeadingResolver,
  orderDocxStylesForPreview,
  wordChild,
  wordValue,
} from "./docx-heading-level";
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
const ATTACHMENT_RELATIONSHIPS = new Set([
  `${OFFICE_RELATIONSHIP_NAMESPACE}/package`,
  `${OFFICE_RELATIONSHIP_NAMESPACE}/oleObject`,
]);
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
  readonly headingLevel: number | null;
  readonly headingNumber?: string;
  readonly indents?: DocxParagraphIndents;
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
  readonly attachments?: readonly DocxAttachment[];
  readonly tocEntries?: readonly DocxTocEntry[];
  readonly blankLineHeights?: readonly { afterParagraphOrdinal: number; twips: number }[];
}

export interface DocxTocEntry {
  readonly paragraphOrdinal: number;
  readonly level: number;
  readonly page: string;
  readonly number: string | null;
  readonly targetOrdinal: number | null;
  readonly bookmark: string | null;
  readonly leader: "dot" | "dash" | "solid" | "none";
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

function docxRelatedPart(
  zip: JSZip,
  relationships: ReadonlyMap<string, Relationship>,
  kind: "styles" | "numbering",
): JSZipObject | null {
  const relationship = Array.from(relationships.values()).find(
    ({ type }) => type === `${OFFICE_RELATIONSHIP_NAMESPACE}/${kind}`,
  );
  if (relationship === undefined) return zip.file(`word/${kind}.xml`);
  if (relationship.external) return null;
  try {
    const base = "https://docx.invalid/word/document.xml";
    const part = new URL(relationship.target, base);
    if (part.origin !== new URL(base).origin || part.search || part.hash) return null;
    return zip.file(decodeURIComponent(part.pathname.slice(1)));
  } catch {
    return null;
  }
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
    if (name !== null && name.length > 0 &&
      (!name.startsWith("_") || /^(?:_Toc|__RefHeading___Toc)/i.test(name))) {
      bookmarks.add(name);
    }
  }
  return [...bookmarks];
}

function docxTocStyles(styles: Document | null) {
  const byId = new Map(Array.from(styles?.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "style") ?? [],
    (style) => [style.getAttributeNS(WORDPROCESSING_NAMESPACE, "styleId"), style] as const));
  const cache = new Map<string, number | null>();
  const levelFor = (styleId: string | null): number | null => {
    const visited = new Set<string>();
    let level: number | null = null;
    while (styleId !== null && !visited.has(styleId)) {
      const cached = cache.get(styleId);
      if (cached !== undefined) { level = cached; break; }
      visited.add(styleId);
      const style = byId.get(styleId) ?? null;
      const pattern = /^(?:toc|contents|目录|目錄)[\s_-]*([1-9])$/i;
      const match = pattern.exec(styleId) ?? pattern.exec(wordValue(style, "name") ?? "");
      if (match !== null) { level = Number(match[1]); break; }
      styleId = wordValue(style, "basedOn");
    }
    for (const id of visited) cache.set(id, level);
    return level;
  };
  const leaderFrom = (properties: Element | null): DocxTocEntry["leader"] | null => {
    const tabs = Array.from(wordChild(properties, "tabs")?.children ?? []);
    const right = tabs.reverse().find((tab) => tab.localName === "tab" &&
      ["right", "end"].includes(tab.getAttributeNS(WORDPROCESSING_NAMESPACE, "val") ?? ""));
    if (right === undefined) return null;
    const value = right.getAttributeNS(WORDPROCESSING_NAMESPACE, "leader");
    if (value === "dot" || value === "middleDot") return "dot";
    if (value === "hyphen") return "dash";
    if (value === "underscore" || value === "heavy") return "solid";
    return "none";
  };
  return {
    levelFor,
    leaderFor(properties: Element | null, styleId: string | null): DocxTocEntry["leader"] {
      let leader = leaderFrom(properties);
      const visited = new Set<string>();
      while (leader === null && styleId !== null && !visited.has(styleId)) {
        visited.add(styleId);
        const style = byId.get(styleId) ?? null;
        leader = leaderFrom(wordChild(style, "pPr"));
        styleId = wordValue(style, "basedOn");
      }
      return leader ?? "dot";
    },
  };
}

function unavailableContent(
  paragraph: Element,
  relationships: ReadonlyMap<string, Relationship>,
  attachments: ReadonlySet<Element>,
): readonly DocxUnavailableContentKind[] {
  const kinds = new Set<DocxUnavailableContentKind>();
  if (
    paragraph.getElementsByTagNameNS(MATH_NAMESPACE, "oMath").length > 0 ||
    paragraph.getElementsByTagNameNS(MATH_NAMESPACE, "oMathPara").length > 0
  ) {
    kinds.add("equation");
  }
  if (
    Array.from(paragraph.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "object"))
      .some((object) => !attachments.has(object)) ||
    Array.from(paragraph.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "pict"))
      .some((picture) => {
        const object = nearestAncestor(picture, WORDPROCESSING_NAMESPACE, "object");
        return object === null || !attachments.has(object);
      })
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

function attachmentPart(
  object: Element,
  relationships: ReadonlyMap<string, Relationship>,
  zip: JSZip,
) {
  const ole = object.getElementsByTagNameNS("urn:schemas-microsoft-com:office:office", "OLEObject")[0];
  const id = ole?.getAttributeNS(OFFICE_RELATIONSHIP_NAMESPACE, "id");
  if (!id || ole?.getAttribute("Type")?.toLowerCase() === "link" || runIsHidden(object)) return;
  const relationship = relationships.get(id);
  if (relationship === undefined || relationship.external ||
    !ATTACHMENT_RELATIONSHIPS.has(relationship.type)) return;
  try {
    const url = new URL(relationship.target, "https://docx.invalid/word/document.xml");
    if (url.origin !== "https://docx.invalid" || url.search || url.hash) return;
    const partPath = decodeURIComponent(url.pathname.slice(1));
    if (!/^word\/embeddings\/(?!\.{1,2}(?:\/|$))[^/\\\x00]+$/i.test(partPath) ||
      docxAttachmentPart(zip, partPath) === null) return;
    const name = partPath.split("/").at(-1)!;
    const extension = docxAttachmentExtension(name);
    if (extension !== null) return { partPath, name, extension };
    if (!/\.bin$/i.test(name)) return;
    const program = ole?.getAttribute("ProgID") ?? "";
    const hint = /^AcroExch\.Document/i.test(program) ? "pdf"
      : /^Word\.Document\.12$/i.test(program) ? "docx"
      : /^Word\.Document\.8$/i.test(program) ? "doc"
      : /^Excel\.Sheet\.12$/i.test(program) ? "xlsx"
      : /^Excel\.Sheet\.8$/i.test(program) ? "xls"
      : /^PowerPoint\.Show\.12$/i.test(program) ? "pptx"
      : /^PowerPoint\.Show\.8$/i.test(program) ? "ppt"
      : null;
    const previewName = object.getAttribute("officeViewerAttachmentName");
    return {
      partPath,
      name: previewName && hint !== null && docxAttachmentExtension(previewName) === hint
        ? previewName : null,
      extension: hint,
    } as const;
  } catch {
    return;
  }
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
      relationship.type === ALT_CHUNK_RELATIONSHIP ||
      ATTACHMENT_RELATIONSHIPS.has(relationship.type)
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
  const stylesPart = docxRelatedPart(zip, relationships, "styles");
  const styles = stylesPart === null ? null : await readXml(stylesPart, signal);
  const headings = createDocxHeadingResolver(styles);
  const paragraphLayout = createDocxParagraphLayoutResolver(styles);

  const body = document.getElementsByTagNameNS(
    WORDPROCESSING_NAMESPACE,
    "body",
  )[0];
  if (body === undefined) throw malformed("DOCX main document has no body");
  const tocStyles = docxTocStyles(styles);
  const hasToc = Array.from(body.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "pStyle"))
    .some((style) => tocStyles.levelFor(style.getAttributeNS(WORDPROCESSING_NAMESPACE, "val")) !== null);
  const numberingPart = docxRelatedPart(zip, relationships, "numbering");
  const listNumber = numberingPart === null ? null
    : createDocxNumbering(await readXml(numberingPart, signal), styles);

  const paragraphs: DocxSemanticParagraph[] = [];
  const tocRows: Array<Pick<DocxTocEntry, "level" | "page" | "bookmark" | "leader"> & {
    ordinal: number;
    title: string;
  }> = [];
  const unavailableBodyBlocks: DocxUnavailableBodyBlock[] = [];
  const blankLineHeights = new Map<number, number>();
  const attachments: DocxAttachment[] = [];
  const objectIndexes = new Map(Array.from(
    document.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "object"),
    (object, index) => [object, index] as const,
  ));
  let markerPrefix: string | undefined;
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
    const hasText = text.trim().length > 0;
    const properties = wordChild(paragraph, "pPr");
    const recognized = new Set<Element>();
    for (const object of paragraph.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "object")) {
      if (nearestAncestor(object, WORDPROCESSING_NAMESPACE, "p") !== paragraph ||
        isDescendantOf(object, WORDPROCESSING_NAMESPACE, "del") ||
        isDescendantOf(object, WORDPROCESSING_NAMESPACE, "txbxContent")) continue;
      const part = attachmentPart(object, relationships, zip);
      const objectIndex = objectIndexes.get(object);
      if (part === undefined || objectIndex === undefined) continue;
      markerPrefix ??= `_office-viewer-attachment-${crypto.randomUUID()}-`;
      attachments.push({
        ...part,
        objectIndex,
        marker: markerPrefix + objectIndex,
        afterParagraphOrdinal: paragraphs.length + Number(hasText),
        hasPreview: wordChild(object, "drawing") !== null,
        centerPreview: object.getAttribute("officeViewerAttachmentCentered") === "true",
      });
      recognized.add(object);
    }
    const unavailable = unavailableContent(paragraph, relationships, recognized);
    if (!hasText) {
      if (unavailable.length > 0) {
        unavailableBodyBlocks.push({
          afterParagraphOrdinal: paragraphs.length,
          kinds: unavailable,
        });
      } else if (recognized.size === 0 && canRenderDocxBlankLine(paragraph)) {
        const { lineHeight, spacing } = paragraphLayout(properties);
        const lines = 1 + (text.match(/\n/g)?.length ?? 0);
        const previous = blankLineHeights.get(paragraphs.length) ?? 0;
        blankLineHeights.set(paragraphs.length, previous + Math.max(0, lineHeight) * lines + spacing);
      }
      continue;
    }
    const { indents } = paragraphLayout(properties);
    const styleId = wordValue(properties, "pStyle");
    const level = hasToc ? tocStyles.levelFor(styleId) : null;
    const headingLevel = level === null ? headings.level(properties, styleId) : null;
    const ordinal = paragraphs.length + 1;
    const hyperlinks = paragraphHyperlinks(paragraph, relationships);
    const number = listNumber?.(properties, styleId);
    if (level !== null) {
      const separator = text.lastIndexOf("\t");
      const title = text.slice(0, separator).trim();
      const page = text.slice(separator + 1).trim();
      if (separator > 0 && title && /^(?:(?:[a-z]+-)?\d+(?:[-.]\d+)*|[ivxlcdm]+)$/i.test(page)) {
        const bookmark = hyperlinks.find((link) => link.kind === "bookmark")?.bookmark ?? null;
        tocRows.push({ ordinal, level, title, page, bookmark, leader: tocStyles.leaderFor(properties, styleId) });
      }
    }
    paragraphs.push({
      ordinal,
      text,
      searchText: text.toLocaleLowerCase(),
      styleId,
      headingLevel,
      ...(headingLevel !== null && number ? { headingNumber: number } : {}),
      ...(Object.keys(indents).length > 0 ? { indents } : {}),
      listItem: paragraphIsListItem(paragraph),
      tableDepth: tableDepth(paragraph),
      bookmarks: paragraphBookmarks(paragraph),
      hyperlinks,
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
  const tocEntries = tocRows.map(({ ordinal, level, title, page, bookmark, leader }): DocxTocEntry => {
    const targetOrdinal = bookmark === null ? null : bookmarkOwners.get(bookmark) ?? null;
    const target = targetOrdinal === null ? undefined : paragraphs[targetOrdinal - 1];
    const number = target !== undefined && target.headingLevel !== null &&
      target.text.replace(/\s/g, "") === title.replace(/\s/g, "")
      ? target.headingNumber ?? null : null;
    return { paragraphOrdinal: ordinal, level, page, number, targetOrdinal, bookmark, leader };
  });

  return {
    paragraphs,
    bookmarkTargets,
    unavailableBodyBlocks,
    ...(attachments.length > 0 ? { attachments } : {}),
    ...(tocEntries.length > 0 ? { tocEntries } : {}),
    ...(blankLineHeights.size > 0 ? { blankLineHeights: Array.from(
      blankLineHeights, ([afterParagraphOrdinal, twips]) => ({ afterParagraphOrdinal, twips }),
    ) } : {}),
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
  prepareLayout?: (document: Document, styles: Document | null) => void,
  attachments: readonly DocxAttachment[] = [],
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
  if (attachments.length > 0) {
    const objects = Array.from(document.getElementsByTagNameNS(WORDPROCESSING_NAMESPACE, "object"));
    for (const { objectIndex, marker, hasPreview } of attachments) {
      const object = objects[objectIndex];
      const paragraph = object === undefined ? null : nearestAncestor(object, WORDPROCESSING_NAMESPACE, "p");
      if (object === undefined || paragraph === null) continue;
      let insertion = object;
      while (insertion.parentElement !== paragraph && insertion.parentElement !== null) {
        insertion = insertion.parentElement;
      }
      const start = document.createElementNS(WORDPROCESSING_NAMESPACE, "w:bookmarkStart");
      start.setAttributeNS(WORDPROCESSING_NAMESPACE, "w:id", String(1_000_000 + objectIndex));
      start.setAttributeNS(WORDPROCESSING_NAMESPACE, "w:name", marker);
      insertion.before(start);
      const drawing = hasPreview ? wordChild(object, "drawing") : null;
      if (drawing !== null) {
        const preview = document.createElementNS(WORDPROCESSING_NAMESPACE, "w:r");
        preview.append(drawing);
        insertion.before(preview);
      }
    }
  }
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
      if (id === null) continue;
      if (ATTACHMENT_RELATIONSHIPS.has(type ?? "")) {
        removableRelationshipIds.add(id);
        continue;
      }
      if (!external) continue;
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
  if (prepareLayout !== undefined) {
    const stylesPart = docxRelatedPart(
      zip, relationshipDocument === null ? new Map() : relationshipMap(relationshipDocument), "styles",
    );
    const styles = stylesPart === null ? null : await readXml(stylesPart, signal);
    prepareLayout(document, styles);
    if (stylesPart !== null && styles !== null && orderDocxStylesForPreview(styles)) {
      zip.file(stylesPart.name, new XMLSerializer().serializeToString(styles));
    }
  }
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
