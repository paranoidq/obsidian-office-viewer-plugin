import { wordChild, wordValue } from "./docx-heading-level";
import { DOCX_EQUATION_MARKER } from "./docx-equations";

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const INDENTS = ["start", "end", "firstLine"] as const;
type IndentSide = typeof INDENTS[number];
type Indent = { twips?: number; chars?: number };
type ParagraphProperties = Partial<Record<IndentSide, Indent>> & {
  fontTwips?: number;
  line?: number;
  lineRule?: string;
  before?: number;
  after?: number;
};

export type DocxParagraphIndents = Partial<Record<IndentSide, number>>;

function integer(value: string | null | undefined): number | undefined {
  if (value == null || !/^-?\d+$/.test(value)) return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) && Math.abs(number) <= 1_000_000 ? number : undefined;
}

function attr(element: Element | null, ...names: string[]): number | undefined {
  for (const name of names) {
    const value = integer(element?.getAttributeNS(WORD_NS, name));
    if (value !== undefined) return value;
  }
}

function fontTwips(run: Element | null): number | undefined {
  const size = integer(wordValue(run, "sz") ?? wordValue(run, "szCs"));
  return size !== undefined && size > 0 && size <= 3276 ? size * 10 : undefined;
}

function mergeProperties(
  base: ParagraphProperties,
  pPr: Element | null,
  run: Element | null = null,
): ParagraphProperties {
  const ind = wordChild(pPr, "ind");
  const spacing = wordChild(pPr, "spacing");
  const hanging = attr(ind, "hanging");
  const hangingChars = attr(ind, "hangingChars");
  return {
    start: {
      twips: attr(ind, "start", "left") ?? base.start?.twips,
      chars: attr(ind, "startChars", "leftChars") ?? base.start?.chars,
    },
    end: {
      twips: attr(ind, "end", "right") ?? base.end?.twips,
      chars: attr(ind, "endChars", "rightChars") ?? base.end?.chars,
    },
    firstLine: {
      twips: hanging === undefined
        ? attr(ind, "firstLine") ?? base.firstLine?.twips : -hanging,
      chars: hangingChars === undefined || hangingChars === 0
        ? attr(ind, "firstLineChars") ?? (hangingChars === 0 ? 0 : base.firstLine?.chars)
        : -hangingChars,
    },
    fontTwips: fontTwips(wordChild(pPr, "rPr")) ?? fontTwips(run) ?? base.fontTwips,
    line: attr(spacing, "line") ?? base.line,
    lineRule: spacing?.getAttributeNS(WORD_NS, "lineRule") ?? base.lineRule,
    before: attr(spacing, "before") ?? base.before,
    after: attr(spacing, "after") ?? base.after,
  };
}

export function createDocxParagraphLayoutResolver(styles: Document | null) {
  const defaults = wordChild(styles?.documentElement ?? null, "docDefaults");
  const base = mergeProperties({}, wordChild(wordChild(defaults, "pPrDefault"), "pPr"),
    wordChild(wordChild(defaults, "rPrDefault"), "rPr"));
  const definitions = new Map<string, Element>();
  const cache = new Map<string, ParagraphProperties>();
  let defaultStyle: string | null = null;
  for (const style of styles?.documentElement.children ?? []) {
    if (style.namespaceURI !== WORD_NS || style.localName !== "style" ||
      style.getAttributeNS(WORD_NS, "type") !== "paragraph") continue;
    const id = style.getAttributeNS(WORD_NS, "styleId");
    if (id === null) continue;
    definitions.set(id, style);
    if (["1", "true", "on"].includes(style.getAttributeNS(WORD_NS, "default") ?? "")) {
      defaultStyle = id;
    }
  }

  function inherited(id: string | null): ParagraphProperties {
    const chain = new Map<string, Element>();
    while (id !== null && !chain.has(id) && !cache.has(id)) {
      const style = definitions.get(id);
      if (style === undefined) break;
      chain.set(id, style);
      id = wordValue(style, "basedOn");
    }
    let result = (id === null ? undefined : cache.get(id)) ?? base;
    for (const [styleId, style] of [...chain].reverse()) {
      result = mergeProperties(result, wordChild(style, "pPr"), wordChild(style, "rPr"));
      cache.set(styleId, result);
    }
    return result;
  }

  return (pPr: Element | null) => {
    const props = mergeProperties(inherited(wordValue(pPr, "pStyle") ?? defaultStyle), pPr);
    const size = props.fontTwips ?? 220;
    const indents: DocxParagraphIndents = {};
    const characterIndents: DocxParagraphIndents = {};
    for (const side of INDENTS) {
      const indent = props[side];
      if (indent?.chars) {
        const twips = Math.round(indent.chars * size / 100);
        characterIndents[side] = twips;
        indents[side] = twips;
      } else if (indent?.twips !== undefined) {
        indents[side] = indent.twips;
      }
    }
    const singleLine = size * 1.2;
    const line = props.line ?? 240;
    const lineHeight = props.lineRule === "exact" ? line
      : props.lineRule === "atLeast" ? Math.max(singleLine, line)
      : singleLine * line / 240;
    return { indents, characterIndents, size, lineHeight,
      spacing: Math.max(0, props.before ?? 0) + Math.max(0, props.after ?? 0) };
  };
}

export function canRenderDocxBlankLine(paragraph: Element): boolean {
  for (const name of ["drawing", "object", "pict", "sym", "sectPr", "lastRenderedPageBreak"]) {
    if (paragraph.getElementsByTagNameNS(WORD_NS, name).length > 0) return false;
  }
  for (const element of paragraph.getElementsByTagNameNS(WORD_NS, "br")) {
    const type = element.getAttributeNS(WORD_NS, "type");
    if (type === "page" || type === "column") return false;
  }
  for (const bookmark of paragraph.getElementsByTagNameNS(WORD_NS, "bookmarkStart")) {
    const name = bookmark.getAttributeNS(WORD_NS, "name");
    if (name?.startsWith("_office-viewer-attachment-") || name?.startsWith(DOCX_EQUATION_MARKER)) return false;
  }
  return true;
}

export function prepareDocxPreviewParagraphs(document: Document, styles: Document | null): void {
  const resolve = createDocxParagraphLayoutResolver(styles);
  const body = wordChild(document.documentElement, "body");
  if (body === null) return;
  const create = (name: string) => document.createElementNS(WORD_NS, `w:${name}`);
  const set = (element: Element, name: string, value: number) =>
    element.setAttributeNS(WORD_NS, `w:${name}`, String(value));
  // docx-preview handles w:br, but drops the equivalent w:cr.
  for (const lineBreak of Array.from(body.getElementsByTagNameNS(WORD_NS, "cr"))) {
    lineBreak.replaceWith(create("br"));
  }
  for (const paragraph of body.getElementsByTagNameNS(WORD_NS, "p")) {
    let properties = wordChild(paragraph, "pPr");
    const { characterIndents, size } = resolve(properties);
    // docx-preview only understands the absolute form of indentation.
    if (Object.keys(characterIndents).length > 0) {
      if (properties === null) {
        properties = create("pPr");
        paragraph.prepend(properties);
      }
      let indent = wordChild(properties, "ind");
      if (indent === null) {
        indent = create("ind");
        properties.append(indent);
      }
      const { start, end, firstLine } = characterIndents;
      if (start !== undefined) {
        indent.removeAttributeNS(WORD_NS, "left");
        set(indent, "start", start);
      }
      if (end !== undefined) {
        indent.removeAttributeNS(WORD_NS, "right");
        set(indent, "end", end);
      }
      if (firstLine !== undefined) {
        indent.removeAttributeNS(WORD_NS, "firstLine");
        indent.removeAttributeNS(WORD_NS, "hanging");
        set(indent, firstLine < 0 ? "hanging" : "firstLine", Math.abs(firstLine));
      }
    }
    if (Array.from(paragraph.getElementsByTagNameNS(WORD_NS, "t"))
      .some((text) => text.textContent?.trim()) || !canRenderDocxBlankLine(paragraph)) continue;
    // An empty HTML paragraph has no line box. Give it the Word paragraph mark's font.
    const run = create("r");
    const mark = wordChild(properties, "rPr");
    const runProperties = create("rPr");
    const fonts = wordChild(mark, "rFonts");
    if (fonts !== null) runProperties.append(fonts.cloneNode(true));
    const fontSize = create("sz");
    set(fontSize, "val", size / 10);
    runProperties.append(fontSize);
    const text = create("t");
    text.textContent = "\u00a0";
    run.append(runProperties, text);
    paragraph.append(run);
  }
}
