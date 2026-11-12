export type MessageLocale = "en" | "zh-Hans" | "zh-Hant";

const ENGLISH_MESSAGES = {
  "viewer.empty": "Open a PPTX file from your Vault to start reading.",
  "viewer.loading": "Loading presentation…",
  "viewer.fallbackTitle": "PPTX viewer",
  "navigation.previous": "Previous",
  "navigation.next": "Next",
  "navigation.slideNumber": "Slide number",
  "navigation.go": "Go",
  "navigation.slide": "Slide",
  "navigation.invalidPage": "Enter a slide number from 1 to {total}.",
  "navigation.renderFailure":
    "Slide {slide} could not be rendered. The previous slide is still shown. Try another slide or open it in the default application.",
  "page.counter": "{current} / {total}",
  "thumbnails.toggle": "Thumbnails",
  "thumbnails.toggleLabel": "Toggle slide thumbnails",
  "thumbnails.railLabel": "Slide thumbnails",
  "thumbnails.slideLabel": "Slide {slide}",
  "thumbnails.previewUnavailable": "Slide {slide} preview unavailable",
  "thumbnails.resizeLabel": "Resize slide thumbnails",
  "thumbnails.resizeTitle":
    "Drag to resize thumbnails; double-click to reset",
  "thumbnails.resizeValue": "{pixels} pixels",
  "notes.toggle": "Speaker notes",
  "notes.toggleLabel": "Toggle speaker notes",
  "notes.panelLabel": "Speaker notes",
  "notes.empty": "This slide has no speaker notes.",
  "notes.unavailable":
    "Speaker notes are unavailable for this presentation.",
  "notes.copy": "Copy speaker notes",
  "notes.copied": "Speaker notes copied with slide reference.",
  "notes.copyFailure": "Unable to copy the speaker notes.",
  "search.open": "Search slide text",
  "search.close": "Close slide search",
  "search.openPresentation": "Search presentation content",
  "search.closePresentation": "Close presentation search",
  "search.inputLabel": "Search slide text",
  "search.placeholder": "Search slide text",
  "search.inputLabelPresentation": "Search presentation content",
  "search.placeholderPresentation": "Search slides and speaker notes",
  "search.resultsLabel": "Matching slides",
  "search.resultCount": "Matching slides: {count}",
  "search.resultRange": "Showing {start}–{end} of {count}",
  "search.previousResults": "Previous search results",
  "search.nextResults": "Next search results",
  "search.noResults":
    "No matching slide text. Images and speaker notes are not searched.",
  "search.noResultsPresentation":
    "No matching slide text or speaker notes. Images, charts, and SmartArt are not searched.",
  "search.noResultsSlides":
    "No matching slide text. Images, charts, SmartArt, and speaker notes are outside this scope.",
  "search.noResultsNotes":
    "No matching speaker notes. Slide text, images, charts, and SmartArt are outside this scope.",
  "search.resultLabel": "Slide {slide}, matches: {matches}",
  "search.slide": "Slide {slide}",
  "search.matchCount": "Matches: {count}",
  "search.scopeLabel": "Search scope",
  "search.scopeAll": "All",
  "search.scopeSlides": "Slides",
  "search.scopeNotes": "Notes",
  "search.provenanceSlideText": "Slide text",
  "search.provenanceNotes": "Speaker notes",
  "search.slideMatchLabel":
    "Slide {slide}, slide text matches: {matches}",
  "search.notesMatchLabel":
    "Slide {slide}, speaker-note matches: {matches}",
  "fullscreen.button": "Full screen",
  "fullscreen.enterLabel": "Enter full screen",
  "fullscreen.exit": "Exit full screen",
  "fullscreen.failure": "Unable to change full-screen mode.",
  "external.open": "Open in default application",
  "external.failure": "Unable to open the default application.",
  "reference.copy": "Copy slide reference",
  "reference.copyEmbed": "Copy slide embed",
  "reference.copied": "Slide reference copied.",
  "reference.embedCopied": "Slide embed copied.",
  "reference.copyFailure": "Unable to copy the slide reference.",
  "selection.copyWithSource": "Copy selected text with source",
  "selection.copiedWithSource": "Selected text copied with slide reference.",
  "selection.copyFailure": "Unable to copy the selected text.",
  "reference.alias": "{name} — Slide {slide}",
  "reference.moved":
    "This reference was created for slide {created}; the same slide is now slide {current}.",
  "reference.missing":
    "The referenced slide is no longer available in this presentation.",
  "reference.openPresentation": "Open presentation",
  "embed.loading": "Loading slide…",
  "embed.currentSlide": "{name} — Slide {slide}",
  "embed.sourceMissing": "The source presentation is no longer available.",
  "embed.renderFailure": "This slide could not be rendered safely.",
  "fileEmbed.sourceMissing": "The source file could not be found.",
  "fileEmbed.openInTab": "Open in new tab",
  "fileEmbed.openInTabFailure": "Unable to open the preview in a new tab.",
  "fileEmbed.renderFailure": "This document could not be previewed.",
  "fileEmbed.legacyUnavailable": "This legacy Office file needs LibreOffice for in-note preview. You can open the source in the default application.",
  "fileEmbed.docxPartial": "Some formatting or content is simplified in this preview.",
  "compatibility.unsupportedContent":
    "Some presentation content may not render correctly. Compare with the default application when accuracy matters.",
  "compatibility.fontSubstitution":
    "One or more presentation fonts are unavailable and may be substituted, which can change the layout.",
  "diagnostics.copy": "Copy diagnostic summary",
  "diagnostics.copied": "Diagnostic summary copied.",
  "diagnostics.copyFailure": "Unable to copy the diagnostic summary.",
  "error.unsupportedLegacy":
    "Legacy PPT files are not supported. Open this file in the default application.",
  "error.malformed": "This PPTX is damaged or incomplete.",
  "error.protected": "This PPTX is encrypted or password-protected.",
  "error.incompatible":
    "This PPTX uses content this viewer cannot safely display.",
  "error.resourceExhausted":
    "This PPTX is too large or complex to open within the viewer's safety limits.",
  "error.cancelled": "Loading this PPTX was cancelled.",
  "error.unknown": "An unexpected error prevented this PPTX from opening.",
  "error.sourceUnmodified": "The original PPTX file was not modified.",
  "error.sourceUnmodifiedLegacy": "The original file was not modified.",
  "error.retry": "Retry",
  "companion.open": "Open companion note",
  "companion.openLabel": "Open presentation companion note",
  "companion.created": "Companion note created.",
  "companion.adopted": "Companion note opened.",
  "companion.opened": "Companion note opened.",
  "companion.migrated": "Companion note moved to the matching location.",
  "companion.conflict":
    "The companion note could not move because {path} is occupied. The previous note remains open.",
  "companion.targetOccupied":
    "Cannot create a companion note because {path} is occupied.",
  "companion.writeFailure": "Unable to create the companion note.",
  "companion.openFailure": "Unable to open the companion note.",
  "settings.rememberPosition": "Remember reading position",
  "settings.rememberPositionDescription":
    "Store only the last slide number and a local file-change fingerprint. Explicitly claimed companion-note path pairs are kept separately and are not cleared by this setting.",
  "settings.localProcessing": "Local processing and privacy",
  "settings.localProcessingDescription":
    "Presentation bytes stay on this device. The source PPTX remains read-only. Companion notes are Markdown files created only by an explicit action and store no slide text in plugin data. Office Viewer does not upload files or include telemetry.",
  "settings.compatibility": "Compatibility and safety",
  "settings.compatibilityDescription":
    "Rendering is a read-only preview. Blocking errors always stay visible. Detectable non-blocking compatibility warnings and the diagnostic copy control appear only when Diagnostic summary is enabled.",
  "settings.diagnostics": "Diagnostic summary",
  "settings.diagnosticsDescription":
    "Off by default. When enabled, detectable compatibility warnings and the copy control appear the next time you open, retry, or reload a file. The copied summary includes versions, file size, slide count, timings, and stable categories. It excludes filenames, paths, slide text, images, and author metadata.",
  "settings.supportDevelopment":
    "If Office Viewer helps your reading workflow, you can support its continued development here.",
  "settings.supportDevelopmentGitHub": "GitHub Sponsors",
  "settings.supportDevelopmentCoffee": "Buy Me a Coffee",
} as const;

export type MessageKey = keyof typeof ENGLISH_MESSAGES;
export type MessageCatalog = Readonly<Record<MessageKey, string>>;
export type MessageParameters = Readonly<Record<string, string | number>>;

const SIMPLIFIED_CHINESE_MESSAGES = {
  "viewer.empty": "从仓库中打开 PPTX 文件即可开始阅读。",
  "viewer.loading": "正在加载演示文稿…",
  "viewer.fallbackTitle": "PPTX 查看器",
  "navigation.previous": "上一页",
  "navigation.next": "下一页",
  "navigation.slideNumber": "幻灯片编号",
  "navigation.go": "跳转",
  "navigation.slide": "幻灯片",
  "navigation.invalidPage": "请输入 1 到 {total} 之间的幻灯片编号。",
  "navigation.renderFailure":
    "无法渲染第 {slide} 张幻灯片。仍显示上一张幻灯片。请尝试其他幻灯片，或使用默认应用打开。",
  "page.counter": "{current} / {total}",
  "thumbnails.toggle": "缩略图",
  "thumbnails.toggleLabel": "切换幻灯片缩略图",
  "thumbnails.railLabel": "幻灯片缩略图",
  "thumbnails.slideLabel": "第 {slide} 张幻灯片",
  "thumbnails.previewUnavailable": "第 {slide} 张幻灯片的预览不可用",
  "thumbnails.resizeLabel": "调整幻灯片缩略图大小",
  "thumbnails.resizeTitle": "拖动以调整缩略图大小；双击以重置",
  "thumbnails.resizeValue": "{pixels} 像素",
  "notes.toggle": "讲者备注",
  "notes.toggleLabel": "切换讲者备注",
  "notes.panelLabel": "讲者备注",
  "notes.empty": "此幻灯片没有讲者备注。",
  "notes.unavailable": "无法提供此演示文稿的讲者备注。",
  "notes.copy": "复制讲者备注",
  "notes.copied": "已复制讲者备注及幻灯片引用。",
  "notes.copyFailure": "无法复制讲者备注。",
  "search.open": "搜索幻灯片文字",
  "search.close": "关闭幻灯片文字搜索",
  "search.openPresentation": "搜索演示文稿内容",
  "search.closePresentation": "关闭演示文稿内容搜索",
  "search.inputLabel": "搜索幻灯片文字",
  "search.placeholder": "搜索幻灯片文字",
  "search.inputLabelPresentation": "搜索演示文稿内容",
  "search.placeholderPresentation": "搜索幻灯片与讲者备注",
  "search.resultsLabel": "匹配的幻灯片",
  "search.resultCount": "找到 {count} 张幻灯片",
  "search.resultRange": "显示第 {start}–{end} 项，共 {count} 项",
  "search.previousResults": "上一组搜索结果",
  "search.nextResults": "下一组搜索结果",
  "search.noResults": "未在幻灯片文字中找到结果；图片和讲者备注尚未搜索。",
  "search.noResultsPresentation":
    "未在幻灯片文字或讲者备注中找到结果；图片、图表和 SmartArt 不在搜索范围内。",
  "search.noResultsSlides":
    "未在幻灯片文字中找到结果；图片、图表、SmartArt 和讲者备注不在此范围内。",
  "search.noResultsNotes":
    "未在讲者备注中找到结果；幻灯片文字、图片、图表和 SmartArt 不在此范围内。",
  "search.resultLabel": "第 {slide} 张幻灯片，{matches} 处匹配",
  "search.slide": "第 {slide} 张幻灯片",
  "search.matchCount": "{count} 处匹配",
  "search.scopeLabel": "搜索范围",
  "search.scopeAll": "全部",
  "search.scopeSlides": "幻灯片",
  "search.scopeNotes": "讲者备注",
  "search.provenanceSlideText": "幻灯片文字",
  "search.provenanceNotes": "讲者备注",
  "search.slideMatchLabel": "第 {slide} 张幻灯片，幻灯片文字 {matches} 处匹配",
  "search.notesMatchLabel": "第 {slide} 张幻灯片，讲者备注 {matches} 处匹配",
  "fullscreen.button": "全屏",
  "fullscreen.enterLabel": "进入全屏",
  "fullscreen.exit": "退出全屏",
  "fullscreen.failure": "无法切换全屏模式。",
  "external.open": "在默认应用中打开",
  "external.failure": "无法打开默认应用。",
  "reference.copy": "复制幻灯片引用",
  "reference.copyEmbed": "复制幻灯片嵌入",
  "reference.copied": "已复制幻灯片引用。",
  "reference.embedCopied": "已复制幻灯片嵌入。",
  "reference.copyFailure": "无法复制幻灯片引用。",
  "selection.copyWithSource": "复制所选文字并附来源",
  "selection.copiedWithSource": "已复制所选文字及幻灯片引用。",
  "selection.copyFailure": "无法复制所选文字。",
  "reference.alias": "{name} — 第 {slide} 张幻灯片",
  "reference.moved":
    "此引用创建时指向第 {created} 张；同一张幻灯片现在是第 {current} 张。",
  "reference.missing": "此演示文稿中已找不到被引用的幻灯片。",
  "reference.openPresentation": "打开演示文稿",
  "embed.loading": "正在加载幻灯片…",
  "embed.currentSlide": "{name} — 第 {slide} 张幻灯片",
  "embed.sourceMissing": "源演示文稿已不存在。",
  "embed.renderFailure": "无法安全渲染此幻灯片。",
  "fileEmbed.sourceMissing": "找不到源文件。",
  "fileEmbed.openInTab": "在新标签页打开",
  "fileEmbed.openInTabFailure": "无法在新标签页中打开预览。",
  "fileEmbed.renderFailure": "无法预览此文档。",
  "fileEmbed.legacyUnavailable": "此旧版 Office 文件需要 LibreOffice 才能在笔记中预览。你可以在系统默认应用中打开原文件。",
  "fileEmbed.docxPartial": "此预览中的部分格式或内容已简化。",
  "compatibility.unsupportedContent":
    "部分演示文稿内容可能无法正确显示。需要确认准确性时，请与默认应用中的效果进行比较。",
  "compatibility.fontSubstitution":
    "一个或多个演示文稿字体不可用，可能会被替换并导致版式变化。",
  "diagnostics.copy": "复制诊断摘要",
  "diagnostics.copied": "已复制诊断摘要。",
  "diagnostics.copyFailure": "无法复制诊断摘要。",
  "error.unsupportedLegacy": "不支持旧版 PPT 文件。请在默认应用中打开此文件。",
  "error.malformed": "此 PPTX 已损坏或不完整。",
  "error.protected": "此 PPTX 已加密或受密码保护。",
  "error.incompatible": "此 PPTX 包含此查看器无法安全显示的内容。",
  "error.resourceExhausted":
    "此 PPTX 过大或过于复杂，超出了查看器的安全限制。",
  "error.cancelled": "已取消加载此 PPTX。",
  "error.unknown": "发生意外错误，无法打开此 PPTX。",
  "error.sourceUnmodified": "原始 PPTX 文件未被修改。",
  "error.sourceUnmodifiedLegacy": "原始文件未被修改。",
  "error.retry": "重试",
  "companion.open": "打开伴生笔记",
  "companion.openLabel": "打开演示文稿伴生笔记",
  "companion.created": "已创建伴生笔记。",
  "companion.adopted": "已打开伴生笔记。",
  "companion.opened": "已打开伴生笔记。",
  "companion.migrated": "伴生笔记已移动到对应位置。",
  "companion.conflict":
    "伴生笔记无法移动，因为 {path} 已被占用。仍打开原先的笔记。",
  "companion.targetOccupied": "无法创建伴生笔记，因为 {path} 已被占用。",
  "companion.writeFailure": "无法创建伴生笔记。",
  "companion.openFailure": "无法打开伴生笔记。",
  "settings.rememberPosition": "记住阅读位置",
  "settings.rememberPositionDescription":
    "仅存储上次阅读的幻灯片编号和用于检测本地文件更改的信息。明确认领的伴生笔记路径对单独保存，不会因此设置而被清除。",
  "settings.localProcessing": "本地处理与隐私",
  "settings.localProcessingDescription":
    "演示文稿数据始终保留在此设备上。源 PPTX 保持只读。伴生笔记是仅在显式操作时创建的 Markdown 文件，插件数据中不保存幻灯片正文。Office Viewer 不会上传文件，也不包含遥测。",
  "settings.compatibility": "兼容性与安全",
  "settings.compatibilityDescription":
    "渲染结果是只读预览。阻断性错误始终可见。可检测的非阻断兼容性提示和诊断复制入口仅在开启诊断摘要后显示。",
  "settings.diagnostics": "诊断摘要",
  "settings.diagnosticsDescription":
    "默认关闭。开启后，下一次打开、重试或重新加载文件时会显示可检测的兼容性提示和复制入口。复制的摘要包含版本、文件大小、幻灯片数量、耗时和稳定分类，不包含文件名、路径、幻灯片文本、图像或作者元数据。",
  "settings.supportDevelopment":
    "如果 Office Viewer 对你的阅读流程有帮助，你可以在这里支持它继续开发。",
  "settings.supportDevelopmentGitHub": "GitHub Sponsors",
  "settings.supportDevelopmentCoffee": "Buy Me a Coffee",
} as const satisfies MessageCatalog;

const TRADITIONAL_CHINESE_MESSAGES = {
  "viewer.empty": "從儲存庫開啟 PPTX 檔案即可開始閱讀。",
  "viewer.loading": "正在載入簡報…",
  "viewer.fallbackTitle": "PPTX 檢視器",
  "navigation.previous": "上一頁",
  "navigation.next": "下一頁",
  "navigation.slideNumber": "投影片編號",
  "navigation.go": "跳至",
  "navigation.slide": "投影片",
  "navigation.invalidPage": "請輸入 1 到 {total} 之間的投影片編號。",
  "navigation.renderFailure":
    "無法呈現第 {slide} 張投影片。仍顯示上一張投影片。請嘗試其他投影片，或使用預設應用程式開啟。",
  "page.counter": "{current} / {total}",
  "thumbnails.toggle": "縮圖",
  "thumbnails.toggleLabel": "切換投影片縮圖",
  "thumbnails.railLabel": "投影片縮圖",
  "thumbnails.slideLabel": "第 {slide} 張投影片",
  "thumbnails.previewUnavailable": "第 {slide} 張投影片的預覽無法使用",
  "thumbnails.resizeLabel": "調整投影片縮圖大小",
  "thumbnails.resizeTitle": "拖曳以調整縮圖大小；按兩下以重設",
  "thumbnails.resizeValue": "{pixels} 像素",
  "notes.toggle": "講者備註",
  "notes.toggleLabel": "切換講者備註",
  "notes.panelLabel": "講者備註",
  "notes.empty": "此投影片沒有講者備註。",
  "notes.unavailable": "無法提供此簡報的講者備註。",
  "notes.copy": "複製講者備註",
  "notes.copied": "已複製講者備註及投影片引用。",
  "notes.copyFailure": "無法複製講者備註。",
  "search.open": "搜尋投影片文字",
  "search.close": "關閉投影片文字搜尋",
  "search.openPresentation": "搜尋簡報內容",
  "search.closePresentation": "關閉簡報內容搜尋",
  "search.inputLabel": "搜尋投影片文字",
  "search.placeholder": "搜尋投影片文字",
  "search.inputLabelPresentation": "搜尋簡報內容",
  "search.placeholderPresentation": "搜尋投影片與講者備註",
  "search.resultsLabel": "相符的投影片",
  "search.resultCount": "找到 {count} 張投影片",
  "search.resultRange": "顯示第 {start}–{end} 項，共 {count} 項",
  "search.previousResults": "上一組搜尋結果",
  "search.nextResults": "下一組搜尋結果",
  "search.noResults": "在投影片文字中找不到結果；尚未搜尋圖片和講者備註。",
  "search.noResultsPresentation":
    "在投影片文字或講者備註中找不到結果；圖片、圖表和 SmartArt 不在搜尋範圍內。",
  "search.noResultsSlides":
    "在投影片文字中找不到結果；圖片、圖表、SmartArt 和講者備註不在此範圍內。",
  "search.noResultsNotes":
    "在講者備註中找不到結果；投影片文字、圖片、圖表和 SmartArt 不在此範圍內。",
  "search.resultLabel": "第 {slide} 張投影片，{matches} 處相符",
  "search.slide": "第 {slide} 張投影片",
  "search.matchCount": "{count} 處相符",
  "search.scopeLabel": "搜尋範圍",
  "search.scopeAll": "全部",
  "search.scopeSlides": "投影片",
  "search.scopeNotes": "講者備註",
  "search.provenanceSlideText": "投影片文字",
  "search.provenanceNotes": "講者備註",
  "search.slideMatchLabel": "第 {slide} 張投影片，投影片文字 {matches} 處相符",
  "search.notesMatchLabel": "第 {slide} 張投影片，講者備註 {matches} 處相符",
  "fullscreen.button": "全螢幕",
  "fullscreen.enterLabel": "進入全螢幕",
  "fullscreen.exit": "結束全螢幕",
  "fullscreen.failure": "無法切換全螢幕模式。",
  "external.open": "在預設應用程式中開啟",
  "external.failure": "無法開啟預設應用程式。",
  "reference.copy": "複製投影片引用",
  "reference.copyEmbed": "複製投影片嵌入",
  "reference.copied": "已複製投影片引用。",
  "reference.embedCopied": "已複製投影片嵌入。",
  "reference.copyFailure": "無法複製投影片引用。",
  "selection.copyWithSource": "複製所選文字並附來源",
  "selection.copiedWithSource": "已複製所選文字及投影片引用。",
  "selection.copyFailure": "無法複製所選文字。",
  "reference.alias": "{name} — 第 {slide} 張投影片",
  "reference.moved":
    "此引用建立時指向第 {created} 張；同一張投影片目前是第 {current} 張。",
  "reference.missing": "此簡報中已找不到被引用的投影片。",
  "reference.openPresentation": "開啟簡報",
  "embed.loading": "正在載入投影片…",
  "embed.currentSlide": "{name} — 第 {slide} 張投影片",
  "embed.sourceMissing": "來源簡報已不存在。",
  "embed.renderFailure": "無法安全呈現此投影片。",
  "fileEmbed.sourceMissing": "找不到來源檔案。",
  "fileEmbed.openInTab": "在新分頁開啟",
  "fileEmbed.openInTabFailure": "無法在新分頁中開啟預覽。",
  "fileEmbed.renderFailure": "無法預覽此文件。",
  "fileEmbed.legacyUnavailable": "此舊版 Office 檔案需要 LibreOffice 才能在筆記中預覽。你可以在系統預設應用程式中開啟原始檔案。",
  "fileEmbed.docxPartial": "此預覽中的部分格式或內容已簡化。",
  "compatibility.unsupportedContent":
    "部分簡報內容可能無法正確顯示。需要確認準確性時，請與預設應用程式中的效果進行比較。",
  "compatibility.fontSubstitution":
    "一個或多個簡報字型無法使用，可能會被替代並導致版面配置變化。",
  "diagnostics.copy": "複製診斷摘要",
  "diagnostics.copied": "已複製診斷摘要。",
  "diagnostics.copyFailure": "無法複製診斷摘要。",
  "error.unsupportedLegacy": "不支援舊版 PPT 檔案。請在預設應用程式中開啟此檔案。",
  "error.malformed": "此 PPTX 已損毀或不完整。",
  "error.protected": "此 PPTX 已加密或受密碼保護。",
  "error.incompatible": "此 PPTX 包含此檢視器無法安全顯示的內容。",
  "error.resourceExhausted":
    "此 PPTX 過大或過於複雜，超出檢視器的安全限制。",
  "error.cancelled": "已取消載入此 PPTX。",
  "error.unknown": "發生未預期的錯誤，無法開啟此 PPTX。",
  "error.sourceUnmodified": "原始 PPTX 檔案未經修改。",
  "error.sourceUnmodifiedLegacy": "原始檔案未經修改。",
  "error.retry": "重試",
  "companion.open": "開啟伴生筆記",
  "companion.openLabel": "開啟簡報伴生筆記",
  "companion.created": "已建立伴生筆記。",
  "companion.adopted": "已開啟伴生筆記。",
  "companion.opened": "已開啟伴生筆記。",
  "companion.migrated": "伴生筆記已移動到對應位置。",
  "companion.conflict":
    "伴生筆記無法移動，因為 {path} 已被占用。仍開啟原先的筆記。",
  "companion.targetOccupied": "無法建立伴生筆記，因為 {path} 已被占用。",
  "companion.writeFailure": "無法建立伴生筆記。",
  "companion.openFailure": "無法開啟伴生筆記。",
  "settings.rememberPosition": "記住閱讀位置",
  "settings.rememberPositionDescription":
    "僅儲存上次閱讀的投影片編號，以及用於偵測本機檔案變更的資訊。明確認領的伴生筆記路徑對單獨保存，不會因此設定而被清除。",
  "settings.localProcessing": "本機處理與隱私",
  "settings.localProcessingDescription":
    "簡報資料始終保留在此裝置上。來源 PPTX 保持唯讀。伴生筆記是僅在明確操作時建立的 Markdown 檔案，外掛資料中不保存投影片正文。Office Viewer 不會上傳檔案，也不包含遙測。",
  "settings.compatibility": "相容性與安全",
  "settings.compatibilityDescription":
    "呈現結果是唯讀預覽。阻斷性錯誤始終可見。可偵測的非阻斷相容性提示和診斷複製入口僅在開啟診斷摘要後顯示。",
  "settings.diagnostics": "診斷摘要",
  "settings.diagnosticsDescription":
    "預設關閉。開啟後，下一次開啟、重試或重新載入檔案時會顯示可偵測的相容性提示和複製入口。複製的摘要包含版本、檔案大小、投影片數量、耗時和穩定分類，不包含檔名、路徑、投影片文字、影像或作者中繼資料。",
  "settings.supportDevelopment":
    "如果 Office Viewer 對你的閱讀流程有幫助，你可以在這裡支持它持續開發。",
  "settings.supportDevelopmentGitHub": "GitHub Sponsors",
  "settings.supportDevelopmentCoffee": "Buy Me a Coffee",
} as const satisfies MessageCatalog;

export const MESSAGE_CATALOGS: Readonly<
  Record<MessageLocale, MessageCatalog>
> = {
  en: ENGLISH_MESSAGES,
  "zh-Hans": SIMPLIFIED_CHINESE_MESSAGES,
  "zh-Hant": TRADITIONAL_CHINESE_MESSAGES,
};

export interface MessageTranslator {
  readonly locale: MessageLocale;
  text(key: MessageKey, parameters?: MessageParameters): string;
}

const SIMPLIFIED_CHINESE_REGIONS = new Set(["cn", "sg"]);
const TRADITIONAL_CHINESE_REGIONS = new Set(["tw", "hk", "mo"]);

export function resolveMessageLocale(language: string): MessageLocale {
  const normalized = language.trim().replaceAll("_", "-").toLowerCase();
  if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(normalized)) return "en";

  const [languageCode, ...subtags] = normalized.split("-");
  if (languageCode === "en") return "en";
  if (languageCode !== "zh") return "en";

  if (
    subtags.includes("hant") ||
    subtags.some((subtag) => TRADITIONAL_CHINESE_REGIONS.has(subtag))
  ) return "zh-Hant";

  if (
    subtags.length === 0 ||
    subtags.includes("hans") ||
    subtags.some((subtag) => SIMPLIFIED_CHINESE_REGIONS.has(subtag))
  ) return "zh-Hans";

  return "en";
}

function interpolate(
  template: string,
  parameters: MessageParameters = {},
): string {
  return template.replace(/\{([a-z][a-zA-Z0-9]*)\}/g, (placeholder, name) => {
    const key = name as keyof MessageParameters;
    const value: string | number | undefined = parameters[key];
    return value === undefined ? placeholder : String(value);
  });
}

export function createMessageTranslator(
  language: string,
  catalogOverride?: Partial<MessageCatalog>,
): MessageTranslator {
  const locale = resolveMessageLocale(language);
  const selectedCatalog = catalogOverride ?? MESSAGE_CATALOGS[locale];
  return {
    locale,
    text: (key, parameters) => {
      const selectedMessage = selectedCatalog[key];
      const template = selectedMessage?.trim()
        ? selectedMessage
        : ENGLISH_MESSAGES[key];
      return interpolate(template, parameters);
    },
  };
}

export const ENGLISH_MESSAGE_TRANSLATOR = createMessageTranslator("en");
