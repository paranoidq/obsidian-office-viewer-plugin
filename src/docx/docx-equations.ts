export const OFFICE_MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MATHML_NAMESPACE = "http://www.w3.org/1998/Math/MathML";
export const DOCX_EQUATION_MARKER = "_office-viewer-equation-";

type MathTag =
  | "math" | "mrow" | "mi" | "mn" | "mo" | "mtext"
  | "mfrac" | "msqrt" | "mroot" | "msub" | "msup" | "msubsup"
  | "munder" | "mover" | "munderover" | "mmultiscripts" | "mprescripts"
  | "mtable" | "mtr" | "mtd";

type MathNode = {
  readonly tag: MathTag;
  readonly children: readonly MathNode[];
  readonly text?: string;
  readonly attributes?: Readonly<Record<string, string>>;
};

export interface DocxEquation {
  readonly index: number;
  readonly marker: string;
  readonly afterParagraphOrdinal: number;
  readonly textOffset: number | null;
  readonly nodes: readonly MathNode[];
}

export function isDocxEquation(element: Element): boolean {
  return element.namespaceURI === OFFICE_MATH_NAMESPACE &&
    (element.localName === "oMath" || element.localName === "oMathPara");
}

export function docxEquationElements(document: Document): Element[] {
  return Array.from(document.getElementsByTagNameNS(OFFICE_MATH_NAMESPACE, "*")).filter(
    (element) => isDocxEquation(element) &&
      !(element.parentElement !== null && isDocxEquation(element.parentElement)),
  );
}

function mathChild(element: Element | undefined, name: string): Element | undefined {
  return Array.from(element?.children ?? []).find(
    (child) => child.namespaceURI === OFFICE_MATH_NAMESPACE && child.localName === name,
  );
}

function property(properties: Element | undefined, name: string): string | null | undefined {
  return mathChild(properties, name)?.getAttributeNS(OFFICE_MATH_NAMESPACE, "val");
}

function enabled(value: string | null | undefined): boolean {
  return value !== undefined && value !== "0" && value !== "off" && value !== "false";
}

function node(tag: MathTag, children: readonly MathNode[], attributes?: MathNode["attributes"]): MathNode {
  return { tag, children, attributes };
}

function token(tag: "mi" | "mn" | "mo" | "mtext", text: string, attributes?: MathNode["attributes"]): MathNode {
  return { tag, text, attributes, children: [] };
}

function textTokens(text: string, properties: Element | undefined): MathNode[] {
  if (enabled(property(properties, "nor"))) return [token("mtext", text)];
  const style = property(properties, "sty");
  return Array.from(text.matchAll(/\d+(?:\.\d+)?|[\p{L}\p{M}]+|\s+|[^]/gu), ([part]) => {
    if (/^\s+$/u.test(part)) return token("mtext", part.replace(/\s/g, "\u00a0"));
    if (/^\d/u.test(part)) return token("mn", part);
    if (/^[\p{L}\p{M}]/u.test(part)) {
      return token("mi", part, style === "p" ? { mathvariant: "normal" } : undefined);
    }
    return token("mo", part);
  });
}

const UNSUPPORTED_EQUATION = Symbol();

/** Translate a bounded set of OMML constructs without copying source markup or URLs. */
export function parseDocxEquation(element: Element): DocxEquation["nodes"] | null {
  let visited = 0;
  let characters = 0;
  const fail = (): never => { throw UNSUPPORTED_EQUATION; };
  const parse = (source: Element, depth: number): MathNode => {
    if (depth > 40 || ++visited > 4096 || source.namespaceURI !== OFFICE_MATH_NAMESPACE) fail();
    const kind = source.localName;
    const properties = mathChild(source, `${kind}Pr`);
    const parts = Array.from(source.children).filter((child) =>
      child !== properties && !(child.namespaceURI === OFFICE_MATH_NAMESPACE &&
        (child.localName === "argPr" || child.localName === "ctrlPr")),
    );
    const children = () => parts.map((part) => parse(part, depth + 1));
    const argument = (name: string) => {
      const part = mathChild(source, name);
      return part === undefined ? fail() : parse(part, depth + 1);
    };
    const base = () => argument("e");

    switch (kind) {
      case "oMath":
      case "e":
      case "num":
      case "den":
      case "deg":
      case "sub":
      case "sup":
      case "fName":
      case "lim":
      case "box":
        return node("mrow", children());
      case "r": {
        // Run properties carry formatting only. Text is the only input to MathML tokens.
        let text = "";
        for (const part of parts) {
          if (part.namespaceURI === OFFICE_MATH_NAMESPACE && part.localName === "t") {
            const value = part.textContent ?? "";
            characters += value.length;
            if (characters > 4096) fail();
            text += value;
          } else if (part.localName !== "rPr") {
            fail();
          }
        }
        return node("mrow", textTokens(text, properties));
      }
      case "f": {
        const type = property(properties, "type");
        const numerator = argument("num");
        const denominator = argument("den");
        if (type === "lin" || type === "skw") {
          return node("mrow", [numerator, token("mo", "/"), denominator]);
        }
        return node("mfrac", [numerator, denominator],
          type === "noBar" ? { linethickness: "0" } : undefined);
      }
      case "sSub":
        return node("msub", [base(), argument("sub")]);
      case "sSup":
        return node("msup", [base(), argument("sup")]);
      case "sSubSup":
        return node("msubsup", [base(), argument("sub"), argument("sup")]);
      case "sPre":
        return node("mmultiscripts", [base(), node("mprescripts", []), argument("sub"), argument("sup")]);
      case "rad":
        return enabled(property(properties, "degHide")) || !mathChild(source, "deg")?.textContent?.trim()
          ? node("msqrt", [base()]) : node("mroot", [base(), argument("deg")]);
      case "d": {
        const opening = property(properties, "begChr") ?? "(";
        const closing = property(properties, "endChr") ?? ")";
        const separator = property(properties, "sepChr") ?? "|";
        const items: MathNode[] = [];
        const fence = (char: string) => token("mo", char, { stretchy: "true" });
        if (opening) items.push(fence(opening));
        children().forEach((item, index) => {
          if (index > 0 && separator) items.push(fence(separator));
          items.push(item);
        });
        if (closing) items.push(fence(closing));
        return node("mrow", items);
      }
      case "nary": {
        const op = token("mo", property(properties, "chr") ?? "∫", { largeop: "true" });
        const lower = !enabled(property(properties, "subHide"));
        const upper = !enabled(property(properties, "supHide"));
        const side = property(properties, "limLoc") === "subSup";
        const operator = lower && upper ? node(side ? "msubsup" : "munderover", [op, argument("sub"), argument("sup")])
          : lower ? node(side ? "msub" : "munder", [op, argument("sub")])
          : upper ? node(side ? "msup" : "mover", [op, argument("sup")]) : op;
        return node("mrow", [operator, base()]);
      }
      case "limLow":
        return node("munder", [base(), argument("lim")]);
      case "limUpp":
        return node("mover", [base(), argument("lim")]);
      case "func":
        return node("mrow", [argument("fName"), token("mo", "\u2061"), base()]);
      case "acc":
        return node("mover", [base(), token("mo", property(properties, "chr") ?? "̂")], { accent: "true" });
      case "bar":
        return property(properties, "pos") === "bot"
          ? node("munder", [base(), token("mo", "_")], { accentunder: "true" })
          : node("mover", [base(), token("mo", "¯")], { accent: "true" });
      case "m":
        return node("mtable", children());
      case "mr":
        return node("mtr", children().map((entry) => node("mtd", [entry])));
      case "eqArr":
        return node("mtable", children().map((entry) => node("mtr", [node("mtd", [entry])])));
      default:
        return fail();
    }
  };

  const display = element.localName === "oMathPara";
  const formulas = display ? Array.from(element.children).filter(isDocxEquation) : [element];
  if (formulas.length === 0) return null;
  const align = property(mathChild(element, "oMathParaPr"), "jc");
  try {
    return formulas.map((formula) => node("math", [parse(formula, 0)], display
      ? { display: "block", "data-alignment": align === "left" || align === "right" ? align : "center" }
      : undefined));
  } catch (error) {
    if (error !== UNSUPPORTED_EQUATION) throw error;
    return null;
  }
}

export function renderDocxEquation(document: Document, equation: DocxEquation): DocumentFragment {
  const create = (part: MathNode): Element => {
    const element = document.createElementNS(MATHML_NAMESPACE, part.tag);
    for (const [name, value] of Object.entries(part.attributes ?? {})) {
      element.setAttribute(name, value);
    }
    if (part.text !== undefined) element.textContent = part.text;
    for (const child of part.children) element.append(create(child));
    return element;
  };
  const fragment = document.createDocumentFragment();
  for (const part of equation.nodes) {
    const math = create(part);
    math.setAttribute("data-docx-equation", "");
    fragment.append(math);
  }
  return fragment;
}

export const DOCX_EQUATION_STYLES = `
  math[data-docx-equation][display="block"][data-alignment="left"] { text-align: left; }
  math[data-docx-equation][display="block"][data-alignment="right"] { text-align: right; }
`;
