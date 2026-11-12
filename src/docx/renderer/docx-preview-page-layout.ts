const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function child(parent: Element, name: string): Element | undefined {
  return Array.from(parent.children).find((element) =>
    element.namespaceURI === WORD_NS && element.localName === name,
  );
}

function twips(element: Element | undefined, name: string): number | undefined {
  const value = element?.getAttributeNS(WORD_NS, name);
  if (value == null || !/^-?\d+$/.test(value)) return undefined;
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : undefined;
}

function setWordAttribute(element: Element, name: string, value: string | number): void {
  element.setAttributeNS(WORD_NS, `w:${name}`, String(value));
}

/** Repair the table properties docx-preview uses, on the sanitized copy only. */
export function prepareDocxPreviewPageTables(document: Document): void {
  const body = child(document.documentElement, "body");
  if (body === undefined) return;
  const sections = body.getElementsByTagNameNS(WORD_NS, "sectPr");
  const section = sections.length === 1 ? sections.item(0) : null;
  const page = section === null ? undefined : child(section, "pgSz");
  const margins = section === null ? undefined : child(section, "pgMar");
  const pageWidth = twips(page, "w");
  const left = twips(margins, "left");
  const right = twips(margins, "right");
  const textWidth = pageWidth !== undefined && left !== undefined && right !== undefined
    ? pageWidth - left - right
    : undefined;

  for (const table of body.getElementsByTagNameNS(WORD_NS, "tbl")) {
    const properties = child(table, "tblPr");
    if (properties === undefined) continue;
    const layout = child(properties, "tblLayout");
    if (layout?.getAttributeNS(WORD_NS, "type") !== "fixed") continue;
    // docx-preview 0.3.6 looks for w:val. CSS fixed layout also needs a width.
    setWordAttribute(layout, "val", "fixed");
    const grid = child(table, "tblGrid");
    const columns = Array.from(grid?.children ?? [])
      .filter((element) => element.namespaceURI === WORD_NS && element.localName === "gridCol")
      .map((column) => twips(column, "w"));
    if (columns.length === 0 || columns.some((width) => width === undefined || width < 0)) continue;
    const gridWidth = columns.reduce<number>((sum, width) => sum + (width ?? 0), 0);
    if (!Number.isSafeInteger(gridWidth) || gridWidth <= 0) continue;
    let width = child(properties, "tblW");
    const widthType = width?.getAttributeNS(WORD_NS, "type");
    if (width === undefined || widthType === "auto" || widthType === "nil") {
      if (width === undefined) {
        width = document.createElementNS(WORD_NS, "w:tblW");
        properties.insertBefore(width, layout);
      }
      setWordAttribute(width, "type", "dxa");
      setWordAttribute(width, "w", gridWidth);
    }

    // The renderer treats a floating table's wrap distance as its position.
    // A nearly full-width table cannot wrap prose beside it. Use its authored
    // offset when the page and text origins are unambiguous.
    const position = child(properties, "tblpPr");
    const tableWidth = width.getAttributeNS(WORD_NS, "type") === "dxa"
      ? twips(width, "w")
      : undefined;
    if (position === undefined || table.parentElement !== body || left === undefined
      || textWidth === undefined || textWidth <= 0
      || tableWidth === undefined || tableWidth < textWidth * 0.95) continue;
    const x = twips(position, "tblpX");
    const horizontalAnchor = position.getAttributeNS(WORD_NS, "horzAnchor");
    if (x !== undefined && (horizontalAnchor === "page" || horizontalAnchor === "text")) {
      setWordAttribute(position, "leftFromText", horizontalAnchor === "page" ? x - left : x);
      setWordAttribute(position, "rightFromText", 0);
    }
    const y = twips(position, "tblpY");
    if (y !== undefined && y >= 0 && position.getAttributeNS(WORD_NS, "vertAnchor") === "text") {
      setWordAttribute(position, "topFromText", y);
    }
  }
}

function normalizeLineSpacing(style: CSSStyleDeclaration): void {
  const spacing = style.lineHeight;
  if (!/^\d+(?:\.\d+)?$/.test(spacing)) return;
  // OOXML automatic spacing uses a font's normal line box as one line;
  // CSS unitless spacing instead uses its font size. For other multiples,
  // approximate the normal line box while preserving the authored ratio.
  const multiple = Number(spacing);
  style.lineHeight = multiple === 1 ? "normal" : String(Math.round(multiple * 120) / 100);
}

export function prepareRenderedDocxPageSpacing(content: HTMLElement): void {
  for (const paragraph of content.querySelectorAll<HTMLElement>("p[style]")) {
    normalizeLineSpacing(paragraph.style);
  }
  for (const sheet of content.querySelectorAll<HTMLStyleElement>("style")) {
    for (const rule of Array.from(sheet.sheet?.cssRules ?? [])) {
      if ("style" in rule) normalizeLineSpacing((rule as CSSStyleRule).style);
    }
  }
}
