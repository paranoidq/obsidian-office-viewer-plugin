import { wordChild, wordValue } from "./docx-heading-level";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

interface NumberingLevel {
  readonly start: number;
  readonly format: string;
  readonly text: string;
  readonly restart: number;
  readonly legal: boolean;
  readonly paragraphStyle: string | null;
}

function readLevel(level: Element, index: number): NumberingLevel {
  return {
    start: Number(wordValue(level, "start") ?? 1),
    format: wordValue(level, "numFmt") ?? "decimal",
    text: wordValue(level, "lvlText") ?? "",
    restart: Number(wordValue(level, "lvlRestart") ?? index),
    legal: wordChild(level, "isLgl") !== null && wordValue(level, "isLgl") !== "0",
    paragraphStyle: wordValue(level, "pStyle"),
  };
}

function chineseCounting(value: number): string | null {
  if (value > 9999) return null;
  const digits = "零一二三四五六七八九";
  if (value === 0) return digits[0]!;
  let text = "";
  let zeroPending = false;
  for (const [place, unit] of [[1000, "千"], [100, "百"], [10, "十"], [1, ""]] as const) {
    const digit = Math.floor(value / place) % 10;
    if (digit === 0) {
      zeroPending = text !== "";
      continue;
    }
    if (zeroPending) text += digits[0];
    if (digit !== 1 || place !== 10 || text !== "") text += digits[digit];
    text += unit;
    zeroPending = false;
  }
  return text;
}

function numberText(value: number, format: string): string | null {
  if (!Number.isSafeInteger(value) || value < 0) return null;
  if (format === "decimal") return String(value);
  if (format === "decimalZero") return String(value).padStart(2, "0");
  if (format === "chineseCounting" || format === "chineseCountingThousand") {
    return chineseCounting(value);
  }
  if (format === "ideographDigital") {
    return String(value).replace(/\d/g, (digit) => "〇一二三四五六七八九"[Number(digit)]!);
  }
  if (value < 1) return null;
  if (format === "upperLetter" || format === "lowerLetter") {
    if (value > 676) return null;
    const letter = String.fromCharCode((format === "upperLetter" ? 65 : 97) + (value - 1) % 26);
    return letter.repeat(Math.ceil(value / 26));
  }
  if ((format === "upperRoman" || format === "lowerRoman") && value <= 3999) {
    let text = "";
    for (const [count, digits] of [
      [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
      [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
    ] as const) {
      text += digits.repeat(Math.floor(value / count));
      value %= count;
    }
    return format === "lowerRoman" ? text.toLowerCase() : text;
  }
  return null;
}

export function createDocxNumbering(numbering: Document | null, styles: Document | null) {
  const styleById = new Map(Array.from(styles?.getElementsByTagNameNS(WORD, "style") ?? [],
    (style) => [style.getAttributeNS(WORD, "styleId"), style] as const));
  const abstract = new Map<string, Map<number, NumberingLevel>>();
  for (const list of Array.from(numbering?.getElementsByTagNameNS(WORD, "abstractNum") ?? [])) {
    const id = list.getAttributeNS(WORD, "abstractNumId");
    if (id === null) continue;
    const levels = new Map<number, NumberingLevel>();
    for (const level of list.getElementsByTagNameNS(WORD, "lvl")) {
      const index = Number(level.getAttributeNS(WORD, "ilvl"));
      if (Number.isInteger(index) && index >= 0 && index < 9) levels.set(index, readLevel(level, index));
    }
    abstract.set(id, levels);
  }
  const lists = new Map<string, { levels: Map<number, NumberingLevel>; counts: Map<number, number> }>();
  for (const list of Array.from(numbering?.getElementsByTagNameNS(WORD, "num") ?? [])) {
    const id = list.getAttributeNS(WORD, "numId");
    const base = abstract.get(wordValue(list, "abstractNumId") ?? "");
    if (id === null || base === undefined) continue;
    const levels = new Map(base);
    for (const override of list.getElementsByTagNameNS(WORD, "lvlOverride")) {
      const index = Number(override.getAttributeNS(WORD, "ilvl"));
      if (!Number.isInteger(index) || index < 0 || index > 8) continue;
      const own = wordChild(override, "lvl");
      const level = own === null ? levels.get(index) : readLevel(own, index);
      if (level === undefined) continue;
      levels.set(index, { ...level, start: Number(wordValue(override, "startOverride") ?? level.start) });
    }
    lists.set(id, { levels, counts: new Map() });
  }
  return (properties: Element | null, styleId: string | null): string | null => {
    if (lists.size === 0) return null;
    const direct = wordChild(properties, "numPr");
    let id = wordValue(direct, "numId");
    let level = wordValue(direct, "ilvl");
    const visited = new Set<string>();
    while ((id === null || level === null) && styleId !== null && !visited.has(styleId)) {
      visited.add(styleId);
      const style = styleById.get(styleId) ?? null;
      const styleNumbering = wordChild(wordChild(style, "pPr"), "numPr");
      id ??= wordValue(styleNumbering, "numId");
      level ??= wordValue(styleNumbering, "ilvl");
      styleId = wordValue(style, "basedOn");
    }
    if (id === null || id === "0") return null;
    const list = lists.get(id);
    const styleLevel = level === null && list !== undefined
      ? Array.from(list.levels).find(([, definition]) =>
        definition.paragraphStyle !== null && visited.has(definition.paragraphStyle),
      )?.[0] : undefined;
    const index = Number(level ?? styleLevel ?? 0);
    const current = list?.levels.get(index);
    if (list === undefined || current === undefined) return null;
    list.counts.set(index, (list.counts.get(index) ?? current.start - 1) + 1);
    for (const [descendant, definition] of list.levels) {
      if (descendant > index && definition.restart > index) list.counts.delete(descendant);
    }
    let valid = true;
    const text = current.text.replace(/%([1-9])/g, (_, position: string) => {
      const ancestor = Number(position) - 1;
      const definition = list.levels.get(ancestor);
      const count = list.counts.get(ancestor);
      const formatted = definition === undefined || count === undefined
        ? null : numberText(count, current.legal ? "decimal" : definition.format);
      if (formatted === null) valid = false;
      return formatted ?? "";
    });
    return valid && /%[1-9]/.test(current.text) ? text : null;
  };
}
