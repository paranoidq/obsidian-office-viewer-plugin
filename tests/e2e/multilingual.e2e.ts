import { browser, expect } from "@wdio/globals";
import { obsidianPage } from "wdio-obsidian-service";
import {
  assertNoNetworkRequests,
  installNetworkGuard,
} from "../compatibility/browser-environment";
import {
  closeSettings,
  DIAGNOSTIC_SUMMARY_LABELS,
  setDiagnosticSummaryEnabled,
  type DiagnosticSummaryHostLanguage,
} from "./office-viewer-settings";

const EXPECTED_UI = {
  en: {
    previous: "Previous",
    next: "Next",
    slideNumber: "Slide number",
    go: "Go",
    thumbnails: "Thumbnails",
    thumbnailRail: "Slide thumbnails",
    firstThumbnail: "Slide 1",
    fullscreen: "Full screen",
    external: "Open in default application",
    diagnostics: "Copy diagnostic summary",
    copyReference: "Copy slide reference",
    copyEmbed: "Copy slide embed",
    companionOpen: "Open presentation companion note",
    searchOpen: "Search slide text",
    searchResults: "Matching slides",
    searchCount: "Matching slides: 1",
    searchNoResults:
      "No matching slide text. Images and speaker notes are not searched.",
    notesToggle: "Speaker notes",
    notesToggleLabel: "Toggle speaker notes",
    notesPanelLabel: "Speaker notes",
    notesCopy: "Copy speaker notes",
    notesCopied: "Speaker notes copied with slide reference.",
    presentationSearchOpen: "Search presentation content",
    presentationSearchClose: "Close presentation search",
    presentationSearchPlaceholder: "Search slides and speaker notes",
    searchScope: "Search scope",
    scopeAll: "All",
    scopeSlides: "Slides",
    scopeNotes: "Notes",
    provenanceSlide: "Slide text",
    provenanceNotes: "Speaker notes",
    slideMatchLabel: "Slide 1, slide text matches: 1",
    notesMatchLabel: "Slide 1, speaker-note matches: 4",
    presentationNoResults:
      "No matching slide text or speaker notes. Images, charts, and SmartArt are not searched.",
    slidesNoResults:
      "No matching slide text. Images, charts, SmartArt, and speaker notes are outside this scope.",
    notesNoResults:
      "No matching speaker notes. Slide text, images, charts, and SmartArt are outside this scope.",
    embedCurrent: "representative-12-slides — Slide 6",
    openInTab: "Open in new tab",
    invalidPage: "Enter a slide number from 1 to 12.",
  },
  zh: {
    previous: "上一页",
    next: "下一页",
    slideNumber: "幻灯片编号",
    go: "跳转",
    thumbnails: "缩略图",
    thumbnailRail: "幻灯片缩略图",
    firstThumbnail: "第 1 张幻灯片",
    fullscreen: "全屏",
    external: "在默认应用中打开",
    diagnostics: "复制诊断摘要",
    copyReference: "复制幻灯片引用",
    copyEmbed: "复制幻灯片嵌入",
    companionOpen: "打开演示文稿伴生笔记",
    searchOpen: "搜索幻灯片文字",
    searchResults: "匹配的幻灯片",
    searchCount: "找到 1 张幻灯片",
    searchNoResults: "未在幻灯片文字中找到结果；图片和讲者备注尚未搜索。",
    notesToggle: "讲者备注",
    notesToggleLabel: "切换讲者备注",
    notesPanelLabel: "讲者备注",
    notesCopy: "复制讲者备注",
    notesCopied: "已复制讲者备注及幻灯片引用。",
    presentationSearchOpen: "搜索演示文稿内容",
    presentationSearchClose: "关闭演示文稿内容搜索",
    presentationSearchPlaceholder: "搜索幻灯片与讲者备注",
    searchScope: "搜索范围",
    scopeAll: "全部",
    scopeSlides: "幻灯片",
    scopeNotes: "讲者备注",
    provenanceSlide: "幻灯片文字",
    provenanceNotes: "讲者备注",
    slideMatchLabel: "第 1 张幻灯片，幻灯片文字 1 处匹配",
    notesMatchLabel: "第 1 张幻灯片，讲者备注 4 处匹配",
    presentationNoResults:
      "未在幻灯片文字或讲者备注中找到结果；图片、图表和 SmartArt 不在搜索范围内。",
    slidesNoResults:
      "未在幻灯片文字中找到结果；图片、图表、SmartArt 和讲者备注不在此范围内。",
    notesNoResults:
      "未在讲者备注中找到结果；幻灯片文字、图片、图表和 SmartArt 不在此范围内。",
    embedCurrent: "representative-12-slides — 第 6 张幻灯片",
    openInTab: "在新标签页打开",
    invalidPage: "请输入 1 到 12 之间的幻灯片编号。",
  },
  "zh-TW": {
    previous: "上一頁",
    next: "下一頁",
    slideNumber: "投影片編號",
    go: "跳至",
    thumbnails: "縮圖",
    thumbnailRail: "投影片縮圖",
    firstThumbnail: "第 1 張投影片",
    fullscreen: "全螢幕",
    external: "在預設應用程式中開啟",
    diagnostics: "複製診斷摘要",
    copyReference: "複製投影片引用",
    copyEmbed: "複製投影片嵌入",
    companionOpen: "開啟簡報伴生筆記",
    searchOpen: "搜尋投影片文字",
    searchResults: "相符的投影片",
    searchCount: "找到 1 張投影片",
    searchNoResults: "在投影片文字中找不到結果；尚未搜尋圖片和講者備註。",
    notesToggle: "講者備註",
    notesToggleLabel: "切換講者備註",
    notesPanelLabel: "講者備註",
    notesCopy: "複製講者備註",
    notesCopied: "已複製講者備註及投影片引用。",
    presentationSearchOpen: "搜尋簡報內容",
    presentationSearchClose: "關閉簡報內容搜尋",
    presentationSearchPlaceholder: "搜尋投影片與講者備註",
    searchScope: "搜尋範圍",
    scopeAll: "全部",
    scopeSlides: "投影片",
    scopeNotes: "講者備註",
    provenanceSlide: "投影片文字",
    provenanceNotes: "講者備註",
    slideMatchLabel: "第 1 張投影片，投影片文字 1 處相符",
    notesMatchLabel: "第 1 張投影片，講者備註 4 處相符",
    presentationNoResults:
      "在投影片文字或講者備註中找不到結果；圖片、圖表和 SmartArt 不在搜尋範圍內。",
    slidesNoResults:
      "在投影片文字中找不到結果；圖片、圖表、SmartArt 和講者備註不在此範圍內。",
    notesNoResults:
      "在講者備註中找不到結果；投影片文字、圖片、圖表和 SmartArt 不在此範圍內。",
    embedCurrent: "representative-12-slides — 第 6 張投影片",
    openInTab: "在新分頁開啟",
    invalidPage: "請輸入 1 到 12 之間的投影片編號。",
  },
  fr: {
    previous: "Previous",
    next: "Next",
    slideNumber: "Slide number",
    go: "Go",
    thumbnails: "Thumbnails",
    thumbnailRail: "Slide thumbnails",
    firstThumbnail: "Slide 1",
    fullscreen: "Full screen",
    external: "Open in default application",
    diagnostics: "Copy diagnostic summary",
    copyReference: "Copy slide reference",
    copyEmbed: "Copy slide embed",
    companionOpen: "Open presentation companion note",
    searchOpen: "Search slide text",
    searchResults: "Matching slides",
    searchCount: "Matching slides: 1",
    searchNoResults:
      "No matching slide text. Images and speaker notes are not searched.",
    notesToggle: "Speaker notes",
    notesToggleLabel: "Toggle speaker notes",
    notesPanelLabel: "Speaker notes",
    notesCopy: "Copy speaker notes",
    notesCopied: "Speaker notes copied with slide reference.",
    presentationSearchOpen: "Search presentation content",
    presentationSearchClose: "Close presentation search",
    presentationSearchPlaceholder: "Search slides and speaker notes",
    searchScope: "Search scope",
    scopeAll: "All",
    scopeSlides: "Slides",
    scopeNotes: "Notes",
    provenanceSlide: "Slide text",
    provenanceNotes: "Speaker notes",
    slideMatchLabel: "Slide 1, slide text matches: 1",
    notesMatchLabel: "Slide 1, speaker-note matches: 4",
    presentationNoResults:
      "No matching slide text or speaker notes. Images, charts, and SmartArt are not searched.",
    slidesNoResults:
      "No matching slide text. Images, charts, SmartArt, and speaker notes are outside this scope.",
    notesNoResults:
      "No matching speaker notes. Slide text, images, charts, and SmartArt are outside this scope.",
    embedCurrent: "representative-12-slides — Slide 6",
    openInTab: "Open in new tab",
    invalidPage: "Enter a slide number from 1 to 12.",
  },
} as const;

type HostLanguage = keyof typeof EXPECTED_UI;

const EXPECTED_HOST_LANGUAGE = {
  "en-US": "en",
  "zh-CN": "zh",
  "zh-TW": "zh-TW",
  fr: "fr",
} as const satisfies Readonly<Record<string, HostLanguage>>;

async function vaultSha256(path: string): Promise<string> {
  return browser.executeObsidian(
    async ({ app, obsidian, require }, vaultPath) => {
      const file = app.vault.getAbstractFileByPath(vaultPath);
      if (!(file instanceof obsidian.TFile)) throw new Error(`Missing ${vaultPath}`);
      const buffer = await app.vault.readBinary(file);
      const { createHash } = require("node:crypto") as typeof import("node:crypto");
      return createHash("sha256").update(new Uint8Array(buffer)).digest("hex");
    },
    path,
  );
}

async function installClipboardProbe(): Promise<void> {
  await browser.execute(() => {
    Object.defineProperty(navigator.clipboard, "writeText", {
      configurable: true,
      value: async () => {},
    });
  });
}

describe("multilingual installed smoke", () => {
  it("follows the Obsidian host language and preserves local-only read-only behavior", async () => {
    await installNetworkGuard();
    const requestedCapabilities = browser.requestedCapabilities as {
      "goog:chromeOptions"?: { args?: string[] };
    };
    const languageArgument = requestedCapabilities["goog:chromeOptions"]?.args
      ?.find((argument) => argument.startsWith("--lang="));
    const requestedLanguage = languageArgument?.slice("--lang=".length);
    if (!requestedLanguage || !(requestedLanguage in EXPECTED_HOST_LANGUAGE)) {
      throw new Error(`Missing expected host-language capability: ${languageArgument}`);
    }
    const expectedHostLanguage = EXPECTED_HOST_LANGUAGE[
      requestedLanguage as keyof typeof EXPECTED_HOST_LANGUAGE
    ];
    const hostLanguage = await browser.executeObsidian(({ obsidian }) =>
      obsidian.getLanguage()
    );
    expect(hostLanguage).toBe(expectedHostLanguage);
    const expected = EXPECTED_UI[expectedHostLanguage];
    const diagnosticHostLanguage =
      expectedHostLanguage in DIAGNOSTIC_SUMMARY_LABELS
        ? (expectedHostLanguage as DiagnosticSummaryHostLanguage)
        : "en";
    const diagnosticLabel = DIAGNOSTIC_SUMMARY_LABELS[diagnosticHostLanguage];
    const path = "performance/representative-12-slides.pptx";
    const before = await vaultSha256(path);
    const speakerNotesPath = "speaker-notes.pptx";
    const speakerNotesBefore = await vaultSha256(speakerNotesPath);

    await setDiagnosticSummaryEnabled(true, diagnosticLabel);
    await closeSettings();

    await obsidianPage.openFile(path);
    const root = await browser.$(
      '.workspace-leaf.mod-active .pptx-viewer[data-state="ready"]',
    );
    await expect(root).toExist();
    await browser.waitUntil(
      async () => (await root.$(".pptx-viewer__page-counter").getText())
        .endsWith(" / 12"),
      { timeout: 30_000, timeoutMsg: "Representative deck did not become active" },
    );
    await expect(root.$(".office-viewer-brand__product")).toHaveText(
      "Office Viewer",
    );
    await expect(root.$(".office-viewer-brand__format")).toHaveText("PPTX");
    await expect(root.$(".office-viewer-brand__creator")).not.toExist();
    await expect(root.$('[data-action="previous-slide"]')).toHaveAttribute(
      "aria-label",
      expected.previous,
    );
    await expect(root.$('[data-action="next-slide"]')).toHaveAttribute(
      "aria-label",
      expected.next,
    );
    await expect(root.$('[data-action="page-number"]')).toHaveAttribute(
      "aria-label",
      expected.slideNumber,
    );
    await expect(root.$('[data-action="jump-to-slide"]')).toHaveText(expected.go);
    await expect(root.$('[data-action="toggle-thumbnails"]')).toHaveText(
      expected.thumbnails,
    );
    await expect(root.$('.pptx-viewer__thumbnail-rail')).toHaveAttribute(
      "aria-label",
      expected.thumbnailRail,
    );
    await expect(root.$('[data-action="thumbnail-slide"]')).toHaveAttribute(
      "aria-label",
      expected.firstThumbnail,
    );
    await expect(root.$('[data-action="toggle-fullscreen"]')).toHaveText(
      expected.fullscreen,
    );
    await expect(root.$('[data-action="open-externally"]')).toHaveAttribute(
      "aria-label",
      expected.external,
    );
    await expect(root.$('[data-action="copy-diagnostics"]')).toHaveAttribute(
      "aria-label",
      expected.diagnostics,
    );
    await expect(root.$('[data-action="copy-slide-reference"]')).toHaveAttribute(
      "aria-label",
      expected.copyReference,
    );
    await expect(root.$('[data-action="copy-slide-embed"]')).toHaveAttribute(
      "aria-label",
      expected.copyEmbed,
    );
    await expect(root.$('[data-action="open-companion-note"]')).toHaveAttribute(
      "aria-label",
      expected.companionOpen,
    );
    await root.$('[data-action="next-slide"]').click();
    await expect(root.$('.pptx-viewer__page-counter')).toHaveText("2 / 12");
    const pageInput = root.$('[data-action="page-number"]');
    await pageInput.setValue("13");
    await root.$('[data-action="jump-to-slide"]').click();
    await expect(root).toHaveText(expect.stringContaining(expected.invalidPage));

    const searchButton = root.$('[data-action="open-slide-search"]');
    await expect(searchButton).toHaveAttribute("aria-label", expected.searchOpen);
    await searchButton.click();
    const searchInput = root.$('[data-action="slide-search-input"]');
    await expect(searchInput).toHaveAttribute("aria-label", expected.searchOpen);
    await expect(root.$('[role="list"]')).toHaveAttribute(
      "aria-label",
      expected.searchResults,
    );
    await searchInput.setValue("unique representative marker 4");
    await expect(root.$(".pptx-viewer__slide-search-summary"))
      .toHaveText(expected.searchCount);
    await searchInput.setValue("missing search marker");
    await expect(root.$(".pptx-viewer__slide-search-summary"))
      .toHaveText(expected.searchNoResults);
    await searchInput.click();
    await browser.keys(["Escape"]);

    await obsidianPage.openFile(speakerNotesPath);
    const notesRoot = await browser.$(
      '.workspace-leaf.mod-active .pptx-viewer[data-state="ready"]',
    );
    await notesRoot.waitForExist({ timeout: 30_000 });
    await browser.waitUntil(
      async () => (await notesRoot.$(".pptx-viewer__page-counter").getText())
        .endsWith(" / 3"),
      { timeout: 30_000, timeoutMsg: "Speaker-notes deck did not become active" },
    );
    const notesToggle = notesRoot.$('[data-action="toggle-notes"]');
    await expect(notesToggle).toHaveText(expected.notesToggle);
    await expect(notesToggle).toHaveAttribute(
      "aria-label",
      expected.notesToggleLabel,
    );
    await expect(notesRoot.$(".pptx-viewer__notes-panel")).toHaveAttribute(
      "aria-label",
      expected.notesPanelLabel,
    );
    const notesCopy = notesRoot.$('[data-action="copy-speaker-notes"]');
    await expect(notesCopy).toHaveAttribute("aria-label", expected.notesCopy);
    await installClipboardProbe();
    await notesCopy.click();
    await expect(notesRoot.$(".pptx-viewer__action-status"))
      .toHaveText(expected.notesCopied);

    const presentationSearch = notesRoot.$(
      '[data-action="open-slide-search"]',
    );
    await expect(presentationSearch).toHaveAttribute(
      "aria-label",
      expected.presentationSearchOpen,
    );
    await presentationSearch.click();
    await expect(presentationSearch).toHaveAttribute(
      "aria-label",
      expected.presentationSearchClose,
    );
    const presentationInput = notesRoot.$(
      '[data-action="slide-search-input"]',
    );
    await expect(presentationInput).toHaveAttribute(
      "aria-label",
      expected.presentationSearchOpen,
    );
    await expect(presentationInput).toHaveAttribute(
      "placeholder",
      expected.presentationSearchPlaceholder,
    );
    await expect(notesRoot.$(".pptx-viewer__slide-search-scopes"))
      .toHaveAttribute("aria-label", expected.searchScope);
    await expect(notesRoot.$(
      '[data-action="search-scope"][data-search-scope="all"]',
    )).toHaveText(expected.scopeAll);
    await expect(notesRoot.$(
      '[data-action="search-scope"][data-search-scope="slides"]',
    )).toHaveText(expected.scopeSlides);
    await expect(notesRoot.$(
      '[data-action="search-scope"][data-search-scope="notes"]',
    )).toHaveText(expected.scopeNotes);

    await presentationInput.setValue("author");
    const slideMatch = notesRoot.$(
      '[data-action="slide-search-slide-match"]',
    );
    const notesMatch = notesRoot.$(
      '[data-action="slide-search-notes-match"]',
    );
    await expect(slideMatch).toHaveAttribute(
      "aria-label",
      expected.slideMatchLabel,
    );
    await expect(notesMatch).toHaveAttribute(
      "aria-label",
      expected.notesMatchLabel,
    );
    await expect(slideMatch.$(".pptx-viewer__slide-search-provenance"))
      .toHaveText(expected.provenanceSlide);
    await expect(notesMatch.$(".pptx-viewer__slide-search-provenance"))
      .toHaveText(expected.provenanceNotes);

    await presentationInput.setValue("missing presentation marker");
    const presentationSummary = notesRoot.$(
      ".pptx-viewer__slide-search-summary",
    );
    await expect(presentationSummary).toHaveText(expected.presentationNoResults);
    await notesRoot.$(
      '[data-action="search-scope"][data-search-scope="slides"]',
    ).click();
    await expect(presentationSummary).toHaveText(expected.slidesNoResults);
    await notesRoot.$(
      '[data-action="search-scope"][data-search-scope="notes"]',
    ).click();
    await expect(presentationSummary).toHaveText(expected.notesNoResults);

    await obsidianPage.openFile("embed-note.md");
    await browser.executeObsidian(({ app }) => {
      (app as unknown as {
        commands: { executeCommandById(id: string): boolean };
      }).commands.executeCommandById("markdown:toggle-preview");
    });
    const embed = await browser.$(
      '.workspace-leaf.mod-active .pptx-slide-embed[data-state="ready"]',
    );
    await embed.waitForExist({ timeout: 10_000 });
    await expect(embed).toHaveAttribute("aria-label", expected.embedCurrent);
    await expect(embed.$('[data-action="open-in-tab"]')).toHaveAttribute("aria-label", expected.openInTab);

    expect(await vaultSha256(path)).toBe(before);
    expect(await vaultSha256(speakerNotesPath)).toBe(speakerNotesBefore);
    await assertNoNetworkRequests();
  });
});
