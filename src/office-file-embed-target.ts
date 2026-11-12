export interface OfficeFileEmbedTarget {
  readonly sourcePath: string;
  readonly format: "pptx" | "docx";
}

export function parseOfficeFileEmbedTarget(link: string): OfficeFileEmbedTarget | null {
  const sourcePath = (link.split("|", 1)[0] ?? "").trim();
  // Fragments belong to the separate slide/paragraph reference contracts.
  if (/[#\[\]\r\n]/.test(sourcePath)) return null;
  const extension = /\.(pptx?|docx?)$/i.exec(sourcePath)?.[1]?.toLowerCase();
  if (extension === undefined) return null;
  return { sourcePath, format: extension.startsWith("ppt") ? "pptx" : "docx" };
}

export function matchStandaloneOfficeFileEmbedLine(line: string) {
  const trimmed = line.trim();
  const match = /^!\[\[([^\[\]\r\n]+)\]\]$/.exec(trimmed);
  const target = match?.[1] === undefined ? null : parseOfficeFileEmbedTarget(match[1]);
  if (target === null) return null;
  const fromOffset = line.indexOf(trimmed);
  return { target, fromOffset, toOffset: fromOffset + trimmed.length };
}
