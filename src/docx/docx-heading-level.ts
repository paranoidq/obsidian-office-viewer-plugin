const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export function wordChild(parent: Element | null, name: string): Element | null {
  return Array.from(parent?.children ?? []).find(
    (element) => element.namespaceURI === WORD_NS && element.localName === name,
  ) ?? null;
}

export function wordValue(parent: Element | null, name: string): string | null {
  return wordChild(parent, name)?.getAttributeNS(WORD_NS, "val") ?? null;
}

function outlineLevel(properties: Element | null): number | null | undefined {
  const value = wordValue(properties, "outlineLvl");
  if (value === null || !/^[0-9]$/.test(value)) return undefined;
  return value === "9" ? null : Number(value) + 1;
}

function namedHeadingLevel(name: string | null): number | undefined {
  const match = /^(?:heading|标题|標題)[\s_]*([1-9])$/i.exec(name?.trim() ?? "");
  return match === null ? undefined : Number(match[1]);
}

export function createDocxHeadingResolver(styles: Document | null) {
  const definitions = new Map<string, Element>();
  let defaultStyle: string | null = null;
  for (const style of Array.from(styles?.documentElement.children ?? [])) {
    if (style.namespaceURI !== WORD_NS || style.localName !== "style"
      || style.getAttributeNS(WORD_NS, "type") !== "paragraph") continue;
    const id = style.getAttributeNS(WORD_NS, "styleId");
    if (id === null) continue;
    definitions.set(id, style);
    if (["1", "true", "on"].includes(style.getAttributeNS(WORD_NS, "default") ?? "")) {
      defaultStyle = id;
    }
  }
  const cached = new Map<string, number | null>();
  const styleLevel = (id: string | null): number | null => {
    const visited = new Set<string>();
    let level: number | null = null;
    while (id !== null && !visited.has(id)) {
      const existing = cached.get(id);
      if (existing !== undefined) {
        level = existing;
        break;
      }
      visited.add(id);
      const style = definitions.get(id) ?? null;
      const authored = outlineLevel(wordChild(style, "pPr"));
      const named = namedHeadingLevel(wordValue(style, "name")) ?? namedHeadingLevel(id);
      if (authored !== undefined || named !== undefined) {
        level = authored === undefined ? named ?? null : authored;
        break;
      }
      id = wordValue(style, "basedOn");
    }
    for (const visitedId of visited) cached.set(visitedId, level);
    return level;
  };
  return {
    level(properties: Element | null, styleId: string | null): number | null {
      const direct = outlineLevel(properties);
      return direct === undefined ? styleLevel(styleId ?? defaultStyle) : direct;
    },
  };
}

/** docx-preview copies base style properties in document order. */
export function orderDocxStylesForPreview(styles: Document): boolean {
  const elements = Array.from(styles.documentElement.children).filter(
    (style) => style.namespaceURI === WORD_NS && style.localName === "style",
  );
  const byId = new Map(elements.map((style) => [style.getAttributeNS(WORD_NS, "styleId"), style]));
  const seen = new Set<Element>();
  const ordered: Element[] = [];
  for (const style of elements) {
    const chain: Element[] = [];
    let current: Element | undefined = style;
    while (current !== undefined && !seen.has(current)) {
      chain.push(current);
      seen.add(current);
      const base = wordValue(current, "basedOn");
      current = base === null ? undefined : byId.get(base);
    }
    for (const ancestor of chain.reverse()) ordered.push(ancestor);
  }
  if (ordered.every((style, index) => style === elements[index])) return false;
  for (const style of ordered) styles.documentElement.appendChild(style);
  return true;
}
