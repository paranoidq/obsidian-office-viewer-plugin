import {
  resolveMessageLocale,
  type MessageLocale,
  type MessageParameters,
} from "../i18n";

const ENGLISH_DOCX_MESSAGES = {
  fallbackTitle: "DOCX reader",
  loading: "Loading document…",
  attachmentName: "Attachment {number}",
  attachmentOpen: "Open {name} in the default application",
  attachmentUnsupported: "This object does not contain a supported document.",
  attachmentOpenFailed: "Unable to open the attachment in the default application.",
  outline: "Document outline",
  outlineOpen: "Show document outline",
  outlineClose: "Hide document outline",
  outlineEmpty: "This document has no headings.",
  outlineEntry: "Level {level} heading. {title}",
  outlineUnavailable: "This heading could not be located in the preview.",
  zoom: "Document zoom",
  zoomIn: "Zoom in",
  zoomOut: "Zoom out",
  zoomReset: "Reset to the default zoom",
  zoomResetAt: "{percent}. Reset to the default zoom",
  searchOpen: "Search document body",
  searchClose: "Close document search",
  searchLabel: "Search document body",
  searchPlaceholder: "Search document body",
  resultsLabel: "Matching paragraphs",
  resultCount: "Matching paragraphs: {count}",
  resultRange: "Showing {start}–{end} of {count}",
  previousResults: "Previous search results",
  nextResults: "Next search results",
  noResults: "No matching document body text.",
  resultLabel: "Paragraph {paragraph}, matches: {matches}",
  paragraph: "Paragraph {paragraph}",
  matchCount: "Matches: {count}",
  openDefault: "Open in default application",
  openDefaultFailure: "Unable to open the default application.",
  retry: "Retry",
  unavailableContent:
    "Some main-body content cannot be represented here. In-flow placeholders mark every detected omission.",
  unavailablePlaceholder: "This document content cannot be displayed.",
  simplifiedView: "Simplified view",
  switchToSimplified: "Switch to simplified view",
  switchToFull: "Switch to full layout",
  showingSimplified: "Showing simplified view.",
  modeSwitchFailed: "Could not switch the document view.",
  mappingDegraded:
    "Some paragraphs could not be linked for search navigation; the document remains readable.",
  blockedLink: "This document link is blocked for safety.",
  bookmarkUnavailable:
    "The internal document link target is unavailable.",
  malformed: "This DOCX is damaged or incomplete.",
  protected: "This DOCX is encrypted or password-protected.",
  incompatible: "This DOCX contains content that cannot be displayed safely.",
  renderFailed: "The full layout could not be displayed. Try switching to simplified view.",
  resourceExhausted:
    "This DOCX is too large or complex to open within the viewer's safety limits.",
  cancelled: "Loading this DOCX was cancelled.",
  unknown: "An unexpected error prevented this DOCX from opening.",
  sourceUnmodified: "The original DOCX file was not modified.",
} as const;

export type DocxMessageKey = keyof typeof ENGLISH_DOCX_MESSAGES;
type DocxMessageCatalog = Readonly<Record<DocxMessageKey, string>>;

const SIMPLIFIED_CHINESE_DOCX_MESSAGES = {
  fallbackTitle: "DOCX 阅读器",
  loading: "正在加载文档…",
  attachmentName: "附件 {number}",
  attachmentOpen: "点击用默认应用打开 {name}",
  attachmentUnsupported: "此对象没有可打开的受支持文档。",
  attachmentOpenFailed: "无法用默认应用打开附件。",
  outline: "文档目录",
  outlineOpen: "显示文档目录",
  outlineClose: "隐藏文档目录",
  outlineEmpty: "此文档没有可用于目录的标题。",
  outlineEntry: "{level} 级标题。{title}",
  outlineUnavailable: "无法在预览中定位此标题。",
  zoom: "文档缩放",
  zoomIn: "放大",
  zoomOut: "缩小",
  zoomReset: "恢复默认缩放",
  zoomResetAt: "当前比例 {percent}。恢复默认缩放",
  searchOpen: "搜索文档正文",
  searchClose: "关闭文档搜索",
  searchLabel: "搜索文档正文",
  searchPlaceholder: "搜索文档正文",
  resultsLabel: "匹配段落",
  resultCount: "找到 {count} 段",
  resultRange: "显示第 {start}–{end} 项，共 {count} 项",
  previousResults: "上一组搜索结果",
  nextResults: "下一组搜索结果",
  noResults: "未在文档正文中找到匹配内容。",
  resultLabel: "第 {paragraph} 段，{matches} 处匹配",
  paragraph: "第 {paragraph} 段",
  matchCount: "{count} 处匹配",
  openDefault: "在默认应用中打开",
  openDefaultFailure: "无法打开默认应用。",
  retry: "重试",
  unavailableContent:
    "部分正文内容无法在此呈现；所有检测到的缺失位置均已显示占位提示。",
  unavailablePlaceholder: "无法显示此文档内容。",
  simplifiedView: "简化视图",
  switchToSimplified: "切换到简化视图",
  switchToFull: "切换到完整排版",
  showingSimplified: "正在使用简化视图。",
  modeSwitchFailed: "无法切换文档视图。",
  mappingDegraded: "部分段落无法关联到搜索定位，但文档仍可阅读。",
  blockedLink: "出于安全原因，已阻止此文档链接。",
  bookmarkUnavailable: "此文档内部链接的目标不可用。",
  malformed: "此 DOCX 已损坏或不完整。",
  protected: "此 DOCX 已加密或受密码保护。",
  incompatible: "此 DOCX 包含无法安全显示的内容。",
  renderFailed: "无法显示完整排版。可尝试切换到简化视图。",
  resourceExhausted: "此 DOCX 过大或过于复杂，超出了查看器的安全限制。",
  cancelled: "已取消加载此 DOCX。",
  unknown: "发生意外错误，无法打开此 DOCX。",
  sourceUnmodified: "原始 DOCX 文件未被修改。",
} as const satisfies DocxMessageCatalog;

const TRADITIONAL_CHINESE_DOCX_MESSAGES = {
  fallbackTitle: "DOCX 閱讀器",
  loading: "正在載入文件…",
  attachmentName: "附件 {number}",
  attachmentOpen: "點擊以預設應用程式開啟 {name}",
  attachmentUnsupported: "此物件沒有可開啟的受支援文件。",
  attachmentOpenFailed: "無法以預設應用程式開啟附件。",
  outline: "文件目錄",
  outlineOpen: "顯示文件目錄",
  outlineClose: "隱藏文件目錄",
  outlineEmpty: "此文件沒有可用於目錄的標題。",
  outlineEntry: "{level} 級標題。{title}",
  outlineUnavailable: "無法在預覽中定位此標題。",
  zoom: "文件縮放",
  zoomIn: "放大",
  zoomOut: "縮小",
  zoomReset: "恢復預設縮放",
  zoomResetAt: "目前比例 {percent}。恢復預設縮放",
  searchOpen: "搜尋文件正文",
  searchClose: "關閉文件搜尋",
  searchLabel: "搜尋文件正文",
  searchPlaceholder: "搜尋文件正文",
  resultsLabel: "相符段落",
  resultCount: "找到 {count} 段",
  resultRange: "顯示第 {start}–{end} 項，共 {count} 項",
  previousResults: "上一組搜尋結果",
  nextResults: "下一組搜尋結果",
  noResults: "未在文件正文中找到相符內容。",
  resultLabel: "第 {paragraph} 段，{matches} 處相符",
  paragraph: "第 {paragraph} 段",
  matchCount: "{count} 處相符",
  openDefault: "在預設應用程式中開啟",
  openDefaultFailure: "無法開啟預設應用程式。",
  retry: "重試",
  unavailableContent:
    "部分正文內容無法在此呈現；所有偵測到的缺失位置均已顯示占位提示。",
  unavailablePlaceholder: "無法顯示此文件內容。",
  simplifiedView: "簡化檢視",
  switchToSimplified: "切換到簡化檢視",
  switchToFull: "切換到完整排版",
  showingSimplified: "正在使用簡化檢視。",
  modeSwitchFailed: "無法切換文件檢視。",
  mappingDegraded: "部分段落無法關聯到搜尋定位，但文件仍可閱讀。",
  blockedLink: "基於安全原因，已阻止此文件連結。",
  bookmarkUnavailable: "此文件內部連結的目標不可用。",
  malformed: "此 DOCX 已損毀或不完整。",
  protected: "此 DOCX 已加密或受密碼保護。",
  incompatible: "此 DOCX 包含無法安全顯示的內容。",
  renderFailed: "無法顯示完整排版。可嘗試切換到簡化檢視。",
  resourceExhausted: "此 DOCX 過大或過於複雜，超出檢視器的安全限制。",
  cancelled: "已取消載入此 DOCX。",
  unknown: "發生未預期的錯誤，無法開啟此 DOCX。",
  sourceUnmodified: "原始 DOCX 檔案未經修改。",
} as const satisfies DocxMessageCatalog;

const CATALOGS: Readonly<Record<MessageLocale, DocxMessageCatalog>> = {
  en: ENGLISH_DOCX_MESSAGES,
  "zh-Hans": SIMPLIFIED_CHINESE_DOCX_MESSAGES,
  "zh-Hant": TRADITIONAL_CHINESE_DOCX_MESSAGES,
};

function interpolate(
  template: string,
  parameters: MessageParameters = {},
): string {
  return template.replace(/\{([a-z][a-zA-Z0-9]*)\}/g, (
    placeholder: string,
    name: string,
  ) => {
    const value: string | number | undefined = parameters[name];
    return value === undefined ? placeholder : String(value);
  });
}

export interface DocxMessageTranslator {
  readonly locale: MessageLocale;
  text(key: DocxMessageKey, parameters?: MessageParameters): string;
}

export function createDocxMessageTranslator(
  language: string,
): DocxMessageTranslator {
  const locale = resolveMessageLocale(language);
  const catalog = CATALOGS[locale];
  return {
    locale,
    text: (key, parameters) => interpolate(catalog[key], parameters),
  };
}
