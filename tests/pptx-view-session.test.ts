import { describe, expect, it, vi } from "vitest";
import { createMessageTranslator } from "../src/i18n";
import { PptxViewSession } from "../src/pptx-view-session";
import { PptxOpenError } from "../src/pptx-open-error";
import type {
  PptxRendererAdapter,
  PptxRendererSession,
  PptxSourceAuthoredSlideText,
  PptxSpeakerNoteContent,
} from "../src/renderer/pptx-renderer-adapter";

function makeRenderer(
  slideCount = 1,
  slideIdentities?: readonly number[],
  sourceAuthoredSlideText?: readonly PptxSourceAuthoredSlideText[],
  speakerNoteContent?: readonly PptxSpeakerNoteContent[],
) {
  const rendererSession: PptxRendererSession = {
    slideCount,
    slideIdentities,
    sourceAuthoredSlideText,
    speakerNoteContent,
    slideWidth: 960,
    slideHeight: 540,
    capabilities: { thumbnails: false, prefetch: false },
    renderSlide: vi.fn(async () => {}),
    dispose: vi.fn(),
  };
  const adapter: PptxRendererAdapter = {
    open: vi.fn(async (_buffer, container) => {
      container.textContent = "Obsidian PPTX smoke test";
      return rendererSession;
    }),
  };
  return { adapter, rendererSession };
}

function makeM2Renderer(slideCount = 10) {
  const rendererSession: PptxRendererSession = {
    slideCount,
    slideWidth: 960,
    slideHeight: 540,
    capabilities: { thumbnails: true, prefetch: true },
    renderSlide: vi.fn(async () => {}),
    renderThumbnail: vi.fn((index, container) => {
      container.textContent = `Preview ${index + 1}`;
      return { ready: Promise.resolve(), dispose: vi.fn() };
    }),
    prefetchSlide: vi.fn(async () => {}),
    dispose: vi.fn(),
  };
  const adapter: PptxRendererAdapter = {
    open: vi.fn(async (_buffer, container) => {
      container.textContent = "Obsidian PPTX smoke test";
      return rendererSession;
    }),
  };
  return { adapter, rendererSession };
}

function makeFullscreen() {
  let activeRoot: HTMLElement | null = null;
  const listeners = new Set<() => void>();
  return {
    api: {
      isActive: vi.fn((root: HTMLElement) => root === activeRoot),
      enter: vi.fn(async (root: HTMLElement) => {
        activeRoot = root;
        listeners.forEach((listener) => listener());
      }),
      exit: vi.fn(async () => {
        activeRoot = null;
        listeners.forEach((listener) => listener());
      }),
      subscribe: vi.fn((listener: () => void) => {
        listeners.add(listener);
        return () => {
          listeners.delete(listener);
        };
      }),
    },
    listeners,
  };
}

describe("PptxViewSession", () => {
  it("finds source-authored text and opens the matching stable slide", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeRenderer(
      2,
      [256, 257],
      [
        { slideId: 256, text: ["Opening summary"] },
        { slideId: 257, text: ["Quarterly revenue", "Revenue target"] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");

    root.querySelector<HTMLButtonElement>('[data-action="open-slide-search"]')
      ?.click();
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    );
    expect(input).not.toBeNull();
    input!.value = "revenue";
    input!.dispatchEvent(new Event("input", { bubbles: true }));

    const result = root.querySelector<HTMLButtonElement>(
      '[data-action="slide-search-result"]',
    );
    expect(result?.textContent).toContain("Slide 2");
    expect(result?.textContent).toContain("Matches: 2");
    result?.click();

    await vi.waitFor(() => {
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(1);
    });
  });

  it("opens search from the keyboard and cycles results from the current slide", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const { adapter, rendererSession } = makeRenderer(
      3,
      [256, 257, 258],
      [
        { slideId: 256, text: ["Needle first"] },
        { slideId: 257, text: ["Current slide"] },
        { slideId: 258, text: ["Needle last"] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        positions: {
          initialSlideFor: () => 1,
          record: vi.fn(),
        },
      },
    );
    await session.open("deck.pptx");

    root.dispatchEvent(new KeyboardEvent("keydown", {
      key: "f",
      ctrlKey: true,
      bubbles: true,
    }));
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    expect(document.activeElement).toBe(input);
    input.value = "needle";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    input.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Enter",
      bubbles: true,
    }));
    await vi.waitFor(() => {
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(2);
    });
    input.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Enter",
      bubbles: true,
    }));
    await vi.waitFor(() => {
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(0);
    });
    input.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Enter",
      shiftKey: true,
      bubbles: true,
    }));
    await vi.waitFor(() => {
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(2);
    });

    input.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Escape",
      bubbles: true,
    }));
    expect(
      root.querySelector<HTMLElement>(".pptx-viewer__thumbnail-rail")?.dataset
        .searchOpen,
    ).toBeUndefined();
    session.dispose();
    root.remove();
  });

  it("temporarily reveals a collapsed thumbnail rail while searching", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const { adapter } = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Needle"] }],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");
    const thumbnails = root.querySelector<HTMLButtonElement>(
      '[data-action="toggle-thumbnails"]',
    )!;
    thumbnails.click();
    expect(root.dataset.thumbnailsCollapsed).toBe("true");

    root.dispatchEvent(new KeyboardEvent("keydown", {
      key: "f",
      metaKey: true,
      bubbles: true,
    }));
    expect(root.dataset.thumbnailsCollapsed).toBe("false");
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    input.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Escape",
      bubbles: true,
    }));

    expect(root.dataset.thumbnailsCollapsed).toBe("true");
    expect(thumbnails.getAttribute("aria-expanded")).toBe("false");
    session.dispose();
    root.remove();
  });

  it("restores the thumbnail scroll position after search navigation", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      2,
      [256, 257],
      [
        { slideId: 256, text: ["First needle"] },
        { slideId: 257, text: ["Second needle"] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");
    const rail = root.querySelector<HTMLElement>(
      ".pptx-viewer__thumbnail-rail",
    )!;
    rail.scrollTop = 123;
    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    rail.scrollTop = 456;

    root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!.dispatchEvent(new KeyboardEvent("keydown", {
      key: "Escape",
      bubbles: true,
    }));

    expect(rail.scrollTop).toBe(123);
    session.dispose();
  });

  it("restores thumbnail mode when the search query is cleared", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Needle"] }],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");
    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const rail = root.querySelector<HTMLElement>(
      ".pptx-viewer__thumbnail-rail",
    )!;
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    input.value = "needle";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    expect(rail.dataset.searchHasQuery).toBe("true");

    input.value = "";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    expect(rail.dataset.searchOpen).toBeUndefined();
    expect(rail.dataset.searchHasQuery).toBeUndefined();
    expect(root.querySelector('[data-action="thumbnail-slide"]')).not.toBeNull();
    session.dispose();
  });

  it("bounds query input and keeps search setup failure non-fatal", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Needle"] }],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");
    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    expect(input.maxLength).toBe(200);
    input.value = "x".repeat(250);
    input.dispatchEvent(new Event("input", { bubbles: true }));
    expect(input.value).toHaveLength(200);
    session.dispose();

    const failingRoot = document.createElement("div");
    const failing = makeRenderer(
      1,
      [512],
      [{ slideId: 512, text: ["Needle"] }],
    );
    const normalize = vi.spyOn(String.prototype, "normalize").mockImplementation(() => {
      throw new Error("search normalization unavailable");
    });
    const failingSession = new PptxViewSession(
      failingRoot,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      failing.adapter,
    );

    await expect(failingSession.open("deck.pptx")).resolves.toBeUndefined();
    normalize.mockRestore();
    expect(failingRoot.dataset.state).toBe("ready");
    expect(failingRoot.querySelector('[data-action="open-slide-search"]'))
      .toBeNull();
    failingSession.dispose();
  });

  it("keeps slide search state independent between workspace panes", async () => {
    const firstRoot = document.createElement("div");
    const secondRoot = document.createElement("div");
    document.body.append(firstRoot, secondRoot);
    const first = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Alpha only"] }],
    );
    const second = makeRenderer(
      1,
      [512],
      [{ slideId: 512, text: ["Beta only"] }],
    );
    const adapter: PptxRendererAdapter = {
      open: vi.fn()
        .mockImplementationOnce(first.adapter.open)
        .mockImplementationOnce(second.adapter.open),
    };
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const firstSession = new PptxViewSession(firstRoot, reader, adapter);
    const secondSession = new PptxViewSession(secondRoot, reader, adapter);
    await Promise.all([
      firstSession.open("first.pptx"),
      secondSession.open("second.pptx"),
    ]);

    firstRoot.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const firstInput = firstRoot.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    firstInput.value = "alpha";
    firstInput.dispatchEvent(new Event("input", { bubbles: true }));

    const secondInput = secondRoot.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    expect(secondInput.value).toBe("");
    expect(secondRoot.querySelector('[data-action="slide-search-result"]'))
      .toBeNull();
    expect(
      secondRoot.querySelector<HTMLElement>(".pptx-viewer__thumbnail-rail")
        ?.dataset.searchOpen,
    ).toBeUndefined();

    firstSession.dispose();
    secondSession.dispose();
    firstRoot.remove();
    secondRoot.remove();
  });

  it("clears search content and listeners when reopening or disposing", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const first = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Private query needle"] }],
    );
    const second = makeRenderer(
      1,
      [512],
      [{ slideId: 512, text: ["Replacement slide"] }],
    );
    const adapter: PptxRendererAdapter = {
      open: vi.fn()
        .mockImplementationOnce(first.adapter.open)
        .mockImplementationOnce(second.adapter.open),
    };
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("first.pptx");
    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const oldInput = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    oldInput.value = "private query";
    oldInput.dispatchEvent(new Event("input", { bubbles: true }));
    expect(JSON.stringify(session.getPerformanceDiagnostics()))
      .not.toContain("private query");

    await session.open("second.pptx");

    expect(oldInput.isConnected).toBe(false);
    const newInput = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    expect(newInput.value).toBe("");
    oldInput.dispatchEvent(new Event("input", { bubbles: true }));
    expect(root.textContent).not.toContain("Private query needle");

    session.dispose();
    expect(root.childElementCount).toBe(0);
    root.dispatchEvent(new KeyboardEvent("keydown", {
      key: "f",
      ctrlKey: true,
      bubbles: true,
    }));
    expect(root.childElementCount).toBe(0);
    root.remove();
  });

  it("keeps slide-only search labels when speaker-note entries are all empty", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      2,
      [256, 257],
      [
        { slideId: 256, text: ["Current needle"] },
        { slideId: 257, text: ["Other needle"] },
      ],
      [
        { slideId: 256, paragraphs: [] },
        { slideId: 257, paragraphs: [] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");

    expect(
      root.querySelector('[data-action="open-slide-search"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("Search slide text");
    expect(root.querySelector('[data-search-scope="all"]')).toBeNull();
    session.dispose();
  });

  it("announces trustworthy no-result coverage and the current result", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      2,
      [256, 257],
      [
        { slideId: 256, text: ["Current needle"] },
        { slideId: 257, text: ["Other needle"] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");
    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    expect(input.getAttribute("aria-label")).toBe("Search slide text");
    expect(root.querySelector('[role="search"]')?.getAttribute("aria-label"))
      .toBe("Search slide text");
    expect(root.querySelector('[role="list"]')?.getAttribute("aria-label"))
      .toBe("Matching slides");

    input.value = "needle";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    expect(root.querySelector('[role="status"]')?.textContent).not.toBeNull();
    expect(
      root.querySelector('[data-action="slide-search-result"]')
        ?.getAttribute("aria-current"),
    ).toBe("location");

    input.value = "missing";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    expect(root.textContent).toContain(
      "Images and speaker notes are not searched.",
    );
    session.dispose();
  });

  it.each([
    [
      "en",
      "Search slide text",
      "Matching slides",
      "Matching slides: 51",
      "Showing 1–50 of 51",
      "Previous search results",
      "Next search results",
      "No matching slide text. Images and speaker notes are not searched.",
    ],
    [
      "zh-CN",
      "搜索幻灯片文字",
      "匹配的幻灯片",
      "找到 51 张幻灯片",
      "显示第 1–50 项，共 51 项",
      "上一组搜索结果",
      "下一组搜索结果",
      "未在幻灯片文字中找到结果；图片和讲者备注尚未搜索。",
    ],
    [
      "zh-TW",
      "搜尋投影片文字",
      "相符的投影片",
      "找到 51 張投影片",
      "顯示第 1–50 項，共 51 項",
      "上一組搜尋結果",
      "下一組搜尋結果",
      "在投影片文字中找不到結果；尚未搜尋圖片和講者備註。",
    ],
  ] as const)(
    "localizes the complete search surface for %s",
    async (
      language,
      searchLabel,
      resultsLabel,
      resultCount,
      resultRange,
      previousResults,
      nextResults,
      noResults,
    ) => {
      const root = document.createElement("div");
      const slides = Array.from({ length: 51 }, (_, index) => ({
        slideId: 256 + index,
        text: [`Needle ${index + 1}`],
      }));
      const { adapter } = makeRenderer(
        slides.length,
        slides.map(({ slideId }) => slideId),
        slides,
      );
      const session = new PptxViewSession(
        root,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        adapter,
        { messages: createMessageTranslator(language) },
      );
      await session.open("deck.pptx");
      const searchButton = root.querySelector<HTMLButtonElement>(
        '[data-action="open-slide-search"]',
      )!;
      expect(searchButton.getAttribute("aria-label")).toBe(searchLabel);
      searchButton.click();
      const input = root.querySelector<HTMLInputElement>(
        '[data-action="slide-search-input"]',
      )!;
      expect(input.getAttribute("aria-label")).toBe(searchLabel);
      expect(root.querySelector('[role="list"]')?.getAttribute("aria-label"))
        .toBe(resultsLabel);
      expect(
        root.querySelector('[data-action="previous-search-results"]')
          ?.getAttribute("aria-label"),
      ).toBe(previousResults);
      expect(
        root.querySelector('[data-action="next-search-results"]')
          ?.getAttribute("aria-label"),
      ).toBe(nextResults);

      input.value = "needle";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      expect(root.querySelector(".pptx-viewer__slide-search-summary")?.textContent)
        .toBe(resultCount);
      expect(root.querySelector(".pptx-viewer__slide-search-range")?.textContent)
        .toBe(resultRange);

      input.value = "missing";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      expect(root.querySelector(".pptx-viewer__slide-search-summary")?.textContent)
        .toBe(noResults);
      session.dispose();
    },
  );

  it("bounds mounted search results and measures large-deck query latency", async () => {
    const root = document.createElement("div");
    const sourceAuthoredSlideText = Array.from({ length: 200 }, (_, index) => ({
      slideId: 256 + index,
      text: [`Needle on stress slide ${index + 1}`],
    }));
    const { adapter } = makeRenderer(
      sourceAuthoredSlideText.length,
      sourceAuthoredSlideText.map(({ slideId }) => slideId),
      sourceAuthoredSlideText,
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("large.pptx");
    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    input.value = "needle";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    const rail = root.querySelector<HTMLElement>(
      ".pptx-viewer__thumbnail-rail",
    )!;
    expect(root.querySelectorAll('[data-action="slide-search-result"]'))
      .toHaveLength(50);
    expect(rail.dataset.mountedSearchResultCount).toBe("50");
    expect(Number(rail.dataset.lastSearchMs)).toBeGreaterThanOrEqual(0);
    expect(root.textContent).toContain("Showing 1–50 of 200");

    root.querySelector<HTMLButtonElement>(
      '[data-action="next-search-results"]',
    )!.click();
    expect(root.textContent).toContain("Showing 51–100 of 200");
    expect(
      root.querySelector('[data-action="slide-search-result"]')?.textContent,
    ).toContain("Slide 51");

    session.dispose();
    expect(rail.dataset.mountedSearchResultCount).toBeUndefined();
    expect(rail.dataset.lastSearchMs).toBeUndefined();
  });

  it("renders the core reading loop in Simplified Chinese", async () => {
    const root = document.createElement("div");
    let finishRead!: (bytes: ArrayBuffer) => void;
    const reader = {
      readBinary: vi.fn(
        () => new Promise<ArrayBuffer>((resolve) => {
          finishRead = resolve;
        }),
      ),
    };
    const { adapter } = makeRenderer(3);
    const session = new PptxViewSession(root, reader, adapter, {
      messages: createMessageTranslator("zh-CN"),
      openExternally: vi.fn(async () => {}),
    });

    expect(root.textContent).toContain("从仓库中打开 PPTX 文件即可开始阅读。");

    const opening = session.open("deck.pptx");
    expect(root.textContent).toContain("正在加载演示文稿…");
    expect(
      root.querySelector('[data-action="previous-slide"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("上一页");
    expect(
      root.querySelector('[data-action="next-slide"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("下一页");
    expect(
      root.querySelector('[data-action="page-number"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("幻灯片编号");
    expect(root.querySelector('[data-action="jump-to-slide"]')?.textContent)
      .toBe("跳转");
    expect(root.querySelector('[data-action="toggle-thumbnails"]')?.textContent)
      .toBe("缩略图");
    expect(
      root.querySelector('[data-action="toggle-thumbnails"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("切换幻灯片缩略图");
    expect(root.querySelector('[data-action="toggle-notes"]')?.textContent)
      .toBe("讲者备注");
    expect(
      root.querySelector('[data-action="toggle-notes"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("切换讲者备注");
    expect(root.querySelector('[data-action="toggle-fullscreen"]')?.textContent)
      .toBe("全屏");
    expect(
      root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("进入全屏");
    expect(root.querySelector('[data-action="open-externally"]')?.getAttribute("aria-label"))
      .toBe("在默认应用中打开");

    finishRead(new ArrayBuffer(1));
    await opening;

    expect(root.textContent).toContain("1 / 3");
  });

  it("renders the core reading loop in Traditional Chinese", async () => {
    const root = document.createElement("div");
    let finishRead!: (bytes: ArrayBuffer) => void;
    const { adapter } = makeRenderer(3);
    const session = new PptxViewSession(
      root,
      {
        readBinary: vi.fn(
          () => new Promise<ArrayBuffer>((resolve) => {
            finishRead = resolve;
          }),
        ),
      },
      adapter,
      {
        messages: createMessageTranslator("zh-TW"),
        openExternally: vi.fn(async () => {}),
      },
    );

    expect(root.textContent).toContain("從儲存庫開啟 PPTX 檔案即可開始閱讀。");

    const opening = session.open("deck.pptx");
    expect(root.textContent).toContain("正在載入簡報…");

    expect(
      root.querySelector('[data-action="previous-slide"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("上一頁");
    expect(
      root.querySelector('[data-action="next-slide"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("下一頁");
    expect(root.querySelector('[data-action="jump-to-slide"]')?.textContent)
      .toBe("跳至");
    expect(root.querySelector('[data-action="toggle-thumbnails"]')?.textContent)
      .toBe("縮圖");
    expect(
      root.querySelector('[data-action="toggle-thumbnails"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("切換投影片縮圖");
    expect(root.querySelector('[data-action="toggle-notes"]')?.textContent)
      .toBe("講者備註");
    expect(
      root.querySelector('[data-action="toggle-notes"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("切換講者備註");
    expect(root.querySelector('[data-action="toggle-fullscreen"]')?.textContent)
      .toBe("全螢幕");
    expect(
      root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("進入全螢幕");
    expect(
      root.querySelector('[data-action="page-number"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("投影片編號");
    expect(root.querySelector('[data-action="open-externally"]')?.getAttribute("aria-label"))
      .toBe("在預設應用程式中開啟");

    finishRead(new ArrayBuffer(1));
    await opening;

    expect(
      root.querySelector('.pptx-viewer__thumbnail-rail')?.getAttribute(
        "aria-label",
      ),
    ).toBe("投影片縮圖");
    expect(
      root.querySelector('[data-action="resize-thumbnails"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("調整投影片縮圖大小");
  });

  it.each([
    [
      "en",
      "Enter a slide number from 1 to 3.",
      "Unable to change full-screen mode.",
      "Unable to open the default application.",
      "Slide 2 could not be rendered. The previous slide is still shown. Try another slide or open it in the default application.",
    ],
    [
      "zh-CN",
      "请输入 1 到 3 之间的幻灯片编号。",
      "无法切换全屏模式。",
      "无法打开默认应用。",
      "无法渲染第 2 张幻灯片。仍显示上一张幻灯片。请尝试其他幻灯片，或使用默认应用打开。",
    ],
    [
      "zh-TW",
      "請輸入 1 到 3 之間的投影片編號。",
      "無法切換全螢幕模式。",
      "無法開啟預設應用程式。",
      "無法呈現第 2 張投影片。仍顯示上一張投影片。請嘗試其他投影片，或使用預設應用程式開啟。",
    ],
  ] as const)(
    "renders validation and recoverable failures for %s",
    async (
      language,
      invalidPage,
      fullscreenFailure,
      externalFailure,
      renderFailure,
    ) => {
      const root = document.createElement("div");
      const rendererSession: PptxRendererSession = {
        slideCount: 3,
        slideWidth: 960,
        slideHeight: 540,
        capabilities: { thumbnails: false, prefetch: false },
        renderSlide: vi.fn(async (index) => {
          if (index === 1) throw new Error("candidate detail");
        }),
        dispose: vi.fn(),
      };
      const adapter: PptxRendererAdapter = {
        open: vi.fn(async () => rendererSession),
      };
      const fullscreen = {
        isActive: vi.fn(() => false),
        enter: vi.fn(async () => {
          throw new Error("platform detail");
        }),
        exit: vi.fn(async () => {}),
        subscribe: vi.fn(() => () => {}),
      };
      const session = new PptxViewSession(
        root,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        adapter,
        {
          messages: createMessageTranslator(language),
          fullscreen,
          openExternally: vi.fn(async () => {
            throw new Error("filesystem detail");
          }),
        },
      );
      await session.open("deck.pptx");

      const input = root.querySelector<HTMLInputElement>(
        '[data-action="page-number"]',
      )!;
      input.value = "0";
      root.querySelector<HTMLButtonElement>('[data-action="jump-to-slide"]')!
        .click();
      expect(root.textContent).toContain(invalidPage);

      root.querySelector<HTMLButtonElement>(
        '[data-action="toggle-fullscreen"]',
      )!.click();
      await vi.waitFor(() => expect(root.textContent).toContain(fullscreenFailure));

      root.querySelector<HTMLButtonElement>(
        '[data-action="open-externally"]',
      )!.click();
      await vi.waitFor(() => expect(root.textContent).toContain(externalFailure));

      root.querySelector<HTMLButtonElement>('[data-action="next-slide"]')!
        .click();
      await vi.waitFor(() => expect(root.textContent).toContain(renderFailure));
      expect(root.textContent).not.toContain("candidate detail");
      expect(root.textContent).not.toContain("platform detail");
      expect(root.textContent).not.toContain("filesystem detail");
    },
  );

  it.each([
    ["en", "Exit full screen", "Full screen", "Enter full screen"],
    ["zh-CN", "退出全屏", "全屏", "进入全屏"],
    ["zh-TW", "結束全螢幕", "全螢幕", "進入全螢幕"],
  ] as const)(
    "updates full-screen visible and accessible text for %s",
    async (language, exit, enterButton, enterLabel) => {
      const root = document.createElement("div");
      const { adapter } = makeRenderer(2);
      const fullscreen = makeFullscreen();
      const session = new PptxViewSession(
        root,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        adapter,
        {
          messages: createMessageTranslator(language),
          fullscreen: fullscreen.api,
        },
      );
      await session.open("deck.pptx");
      const button = root.querySelector<HTMLButtonElement>(
        '[data-action="toggle-fullscreen"]',
      )!;

      button.click();
      await vi.waitFor(() => expect(root.dataset.fullscreen).toBe("true"));
      expect(button.textContent).toBe(exit);
      expect(button.getAttribute("aria-label")).toBe(exit);

      button.click();
      await vi.waitFor(() => expect(root.dataset.fullscreen).toBe("false"));
      expect(button.textContent).toBe(enterButton);
      expect(button.getAttribute("aria-label")).toBe(enterLabel);
    },
  );

  it.each([
    ["en", "unsupported-legacy", "Legacy PPT files are not supported. Open this file in the default application.", "The original file was not modified.", "Retry"],
    ["en", "malformed", "This PPTX is damaged or incomplete.", "The original PPTX file was not modified.", "Retry"],
    ["en", "protected", "This PPTX is encrypted or password-protected.", "The original PPTX file was not modified.", "Retry"],
    ["en", "incompatible", "This PPTX uses content this viewer cannot safely display.", "The original PPTX file was not modified.", "Retry"],
    ["en", "resource-exhausted", "This PPTX is too large or complex to open within the viewer's safety limits.", "The original PPTX file was not modified.", "Retry"],
    ["en", "cancelled", "Loading this PPTX was cancelled.", "The original PPTX file was not modified.", "Retry"],
    ["en", "unknown", "An unexpected error prevented this PPTX from opening.", "The original PPTX file was not modified.", "Retry"],
    ["zh-CN", "malformed", "此 PPTX 已损坏或不完整。", "原始 PPTX 文件未被修改。", "重试"],
    ["zh-CN", "unsupported-legacy", "不支持旧版 PPT 文件。请在默认应用中打开此文件。", "原始文件未被修改。", "重试"],
    ["zh-CN", "protected", "此 PPTX 已加密或受密码保护。", "原始 PPTX 文件未被修改。", "重试"],
    ["zh-CN", "incompatible", "此 PPTX 包含此查看器无法安全显示的内容。", "原始 PPTX 文件未被修改。", "重试"],
    ["zh-CN", "resource-exhausted", "此 PPTX 过大或过于复杂，超出了查看器的安全限制。", "原始 PPTX 文件未被修改。", "重试"],
    ["zh-CN", "cancelled", "已取消加载此 PPTX。", "原始 PPTX 文件未被修改。", "重试"],
    ["zh-CN", "unknown", "发生意外错误，无法打开此 PPTX。", "原始 PPTX 文件未被修改。", "重试"],
    ["zh-TW", "malformed", "此 PPTX 已損毀或不完整。", "原始 PPTX 檔案未經修改。", "重試"],
    ["zh-TW", "unsupported-legacy", "不支援舊版 PPT 檔案。請在預設應用程式中開啟此檔案。", "原始檔案未經修改。", "重試"],
    ["zh-TW", "protected", "此 PPTX 已加密或受密碼保護。", "原始 PPTX 檔案未經修改。", "重試"],
    ["zh-TW", "incompatible", "此 PPTX 包含此檢視器無法安全顯示的內容。", "原始 PPTX 檔案未經修改。", "重試"],
    ["zh-TW", "resource-exhausted", "此 PPTX 過大或過於複雜，超出檢視器的安全限制。", "原始 PPTX 檔案未經修改。", "重試"],
    ["zh-TW", "cancelled", "已取消載入此 PPTX。", "原始 PPTX 檔案未經修改。", "重試"],
    ["zh-TW", "unknown", "發生未預期的錯誤，無法開啟此 PPTX。", "原始 PPTX 檔案未經修改。", "重試"],
  ] as const)(
    "renders the %s %s blocking error surface",
    async (language, category, title, safety, retry) => {
      const root = document.createElement("div");
      const adapter: PptxRendererAdapter = {
        open: vi.fn(async () => {
          throw new PptxOpenError(category, "private candidate detail");
        }),
      };
      const session = new PptxViewSession(
        root,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        adapter,
        { messages: createMessageTranslator(language) },
      );

      await session.open("deck.pptx");

      expect(root.textContent).toContain(title);
      expect(root.textContent).toContain(safety);
      expect(root.querySelector('[data-action="retry"]')?.textContent).toBe(
        retry,
      );
      expect(root.textContent).not.toContain("private candidate detail");
    },
  );

  it.each([
    ["en", "Some presentation content may not render correctly", "One or more presentation fonts are unavailable"],
    ["zh-CN", "部分演示文稿内容可能无法正确显示", "一个或多个演示文稿字体不可用"],
    ["zh-TW", "部分簡報內容可能無法正確顯示", "一個或多個簡報字型無法使用"],
  ] as const)(
    "keeps known compatibility limitations visible in %s",
    async (language, unsupported, fontSubstitution) => {
      const root = document.createElement("div");
      const { adapter, rendererSession } = makeRenderer(2);
      Object.defineProperty(rendererSession, "compatibilityWarnings", {
        value: ["unsupported-content", "font-substitution"],
      });
      const session = new PptxViewSession(
        root,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        adapter,
        {
          messages: createMessageTranslator(language),
          openExternally: vi.fn(async () => {}),
        },
      );

      await session.open("deck.pptx");

      expect(root.dataset.state).toBe("ready");
      expect(root.querySelector('[data-warning-category="unsupported-content"]')
        ?.textContent).toContain(unsupported);
      expect(root.querySelector('[data-warning-category="font-substitution"]')
        ?.textContent).toContain(fontSubstitution);
      expect(root.querySelector('[data-action="open-externally"]')).not.toBeNull();
    },
  );

  it("hides compatibility warnings and diagnostic copy when diagnostic summary is disabled", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeRenderer(2);
    Object.defineProperty(rendererSession, "compatibilityWarnings", {
      value: ["unsupported-content", "font-substitution"],
    });
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        diagnostics: {
          environment: {
            pluginVersion: "0.0.1",
            obsidianVersion: "1.13.1",
            rendererVersion: "1.2.4",
            operatingSystem: "darwin-arm64",
          },
          rememberReadingPosition: () => false,
          enabled: () => false,
          copy: vi.fn(async () => {}),
        },
      },
    );

    await session.open("deck.pptx");

    expect(root.dataset.state).toBe("ready");
    expect(root.querySelector("[data-warning-category]")).toBeNull();
    expect(root.querySelector('[data-action="copy-diagnostics"]')).toBeNull();
    expect(root.dataset.warningCategories).toBeUndefined();
  });

  it.each([
    ["en", "Copy diagnostic summary", "Diagnostic summary copied."],
    ["zh-CN", "复制诊断摘要", "已复制诊断摘要。"],
    ["zh-TW", "複製診斷摘要", "已複製診斷摘要。"],
  ] as const)("copies a content-free diagnostic summary in %s", async (
    language,
    copyLabel,
    copiedStatus,
  ) => {
    const root = document.createElement("div");
    const copy = vi.fn(async (_summary: string) => {});
    const { adapter } = makeRenderer(2);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(42)) },
      adapter,
      {
        messages: createMessageTranslator(language),
        diagnostics: {
          environment: {
            pluginVersion: "0.0.1",
            obsidianVersion: "1.13.1",
            rendererVersion: "1.2.4",
            operatingSystem: "darwin-arm64",
          },
          rememberReadingPosition: () => true,
          enabled: () => true,
          copy,
        },
      },
    );

    await session.open("private/customer-roadmap.pptx");
    const copyButton = root.querySelector<HTMLButtonElement>(
      '[data-action="copy-diagnostics"]',
    )!;
    expect(copyButton.getAttribute("aria-label")).toBe(copyLabel);
    copyButton.click();
    await vi.waitFor(() => expect(copy).toHaveBeenCalledOnce());

    const copied = copy.mock.calls[0]![0];
    expect(JSON.parse(copied)).toMatchObject({
      sourceBytes: 42,
      slideCount: 2,
      lifecyclePhase: "ready",
      errorCategory: null,
    });
    expect(copied).not.toContain("customer-roadmap");
    expect(root.textContent).toContain(copiedStatus);
  });

  it.each([
    ["en", "Unable to copy the diagnostic summary."],
    ["zh-CN", "无法复制诊断摘要。"],
    ["zh-TW", "無法複製診斷摘要。"],
  ] as const)("reports a diagnostic copy failure in %s", async (
    language,
    failureStatus,
  ) => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer();
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        messages: createMessageTranslator(language),
        diagnostics: {
          environment: {
            pluginVersion: "0.0.1",
            obsidianVersion: "1.13.1",
            rendererVersion: "1.2.4",
            operatingSystem: "darwin-arm64",
          },
          rememberReadingPosition: () => false,
          enabled: () => true,
          copy: vi.fn(async () => { throw new Error("clipboard unavailable"); }),
        },
      },
    );

    await session.open("deck.pptx");
    root.querySelector<HTMLButtonElement>('[data-action="copy-diagnostics"]')!
      .click();

    await vi.waitFor(() =>
      expect(root.textContent).toContain(failureStatus)
    );
  });

  it("classifies a current AbortError as a recoverable cancelled load", async () => {
    const root = document.createElement("div");
    const adapter: PptxRendererAdapter = {
      open: vi.fn(async () => {
        throw new DOMException("cancelled", "AbortError");
      }),
    };
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );

    await session.open("deck.pptx");

    expect(root.dataset.errorCategory).toBe("cancelled");
    expect(root.textContent).toContain("Loading this PPTX was cancelled.");
  });

  it("keeps fit-to-window automatic and exposes no manual zoom controls", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(3);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );

    await session.open("deck.pptx");

    expect(root.querySelector('[data-action="zoom-out"]')).toBeNull();
    expect(root.querySelector('[data-action="zoom-in"]')).toBeNull();
    expect(root.querySelector('[data-action="fit-slide"]')).toBeNull();
    expect(root.dataset.zoomMode).toBeUndefined();
    expect(root.dataset.zoomPercent).toBeUndefined();
    expect(root.querySelector('[data-action="open-externally"]')).toBeNull();

    session.dispose();
  });

  it("exposes an accessible thumbnail divider with keyboard resize and reset", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeM2Renderer(12);
    const recordWidth = vi.fn();
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        thumbnailRail: {
          initialWidth: () => 300,
          recordWidth,
        },
      },
    );

    await session.open("deck.pptx");
    const divider = root.querySelector<HTMLElement>(
      '[role="separator"][aria-orientation="vertical"]',
    );
    const rail = root.querySelector<HTMLElement>(
      ".pptx-viewer__thumbnail-rail",
    );

    expect(divider?.getAttribute("aria-label")).toBe(
      "Resize slide thumbnails",
    );
    expect(divider?.tabIndex).toBe(0);
    expect(divider?.getAttribute("aria-valuenow")).toBe("300");
    expect(rail?.style.getPropertyValue("--pptx-thumbnail-rail-width")).toBe(
      "300px",
    );

    divider?.dispatchEvent(
      new KeyboardEvent("keydown", {
        bubbles: true,
        cancelable: true,
        key: "ArrowRight",
      }),
    );
    expect(divider?.getAttribute("aria-valuenow")).toBe("316");
    expect(recordWidth).toHaveBeenLastCalledWith(316);
    await Promise.resolve();
    expect(rendererSession.renderSlide).toHaveBeenCalledTimes(1);
    expect(root.textContent).toContain("1 / 12");

    divider?.dispatchEvent(
      new KeyboardEvent("keydown", {
        bubbles: true,
        cancelable: true,
        key: "ArrowLeft",
        shiftKey: true,
      }),
    );
    expect(divider?.getAttribute("aria-valuenow")).toBe("268");
    expect(recordWidth).toHaveBeenLastCalledWith(268);

    divider?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(divider?.getAttribute("aria-valuenow")).toBe("168");
    expect(recordWidth).toHaveBeenLastCalledWith(168);

    session.dispose();
  });

  it("updates thumbnail layout while dragging and rerenders only after release", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeM2Renderer(12);
    const recordWidth = vi.fn();
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        thumbnailRail: {
          initialWidth: () => 300,
          recordWidth,
        },
      },
    );
    await session.open("deck.pptx");
    const divider = root.querySelector<HTMLElement>(
      '[data-action="resize-thumbnails"]',
    )!;
    const rail = root.querySelector<HTMLElement>(
      ".pptx-viewer__thumbnail-rail",
    )!;
    const pointer = (type: string, clientX: number) => {
      const event = new MouseEvent(type, {
        bubbles: true,
        button: 0,
        cancelable: true,
        clientX,
      });
      Object.defineProperty(event, "pointerId", { value: 7 });
      return event;
    };
    await vi.waitFor(() =>
      expect(rendererSession.renderThumbnail).toHaveBeenCalled(),
    );
    const callsBeforeDrag = vi.mocked(rendererSession.renderThumbnail!).mock
      .calls.length;

    divider.dispatchEvent(pointer("pointerdown", 300));
    window.dispatchEvent(pointer("pointermove", 400));

    expect(rail.style.getPropertyValue("--pptx-thumbnail-rail-width")).toBe(
      "400px",
    );
    expect(recordWidth).not.toHaveBeenCalled();
    expect(vi.mocked(rendererSession.renderThumbnail!).mock.calls).toHaveLength(
      callsBeforeDrag,
    );

    window.dispatchEvent(pointer("pointerup", 400));

    expect(recordWidth).toHaveBeenLastCalledWith(400);
    await vi.waitFor(() =>
      expect(
        vi.mocked(rendererSession.renderThumbnail!).mock.calls.some(
          (call) => call[3] === 376,
        ),
      ).toBe(true),
    );

    session.dispose();
  });

  it("synchronizes the Vault-wide thumbnail width across open sessions", async () => {
    let width = 168;
    const listeners = new Set<(nextWidth: number) => void>();
    const thumbnailRail = {
      initialWidth: () => width,
      recordWidth: (nextWidth: number) => {
        width = nextWidth;
        listeners.forEach((listener) => listener(nextWidth));
      },
      subscribeWidth: (listener: (nextWidth: number) => void) => {
        listeners.add(listener);
        return () => listeners.delete(listener);
      },
    };
    const firstRoot = document.createElement("div");
    const secondRoot = document.createElement("div");
    const first = makeM2Renderer(3);
    const second = makeM2Renderer(3);
    const firstSession = new PptxViewSession(
      firstRoot,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      first.adapter,
      { thumbnailRail },
    );
    const secondSession = new PptxViewSession(
      secondRoot,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      second.adapter,
      { thumbnailRail },
    );
    await Promise.all([
      firstSession.open("first.pptx"),
      secondSession.open("second.pptx"),
    ]);

    firstRoot.querySelector<HTMLElement>('[data-action="resize-thumbnails"]')
      ?.dispatchEvent(new KeyboardEvent("keydown", {
        bubbles: true,
        cancelable: true,
        key: "ArrowRight",
      }));

    expect(width).toBe(184);
    expect(
      secondRoot.querySelector<HTMLElement>('[data-action="resize-thumbnails"]')
        ?.getAttribute("aria-valuenow"),
    ).toBe("184");
    firstSession.dispose();
    secondSession.dispose();
    expect(listeners.size).toBe(0);
  });

  it("renders a restored page exactly once and records only later successful navigation", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter, rendererSession } = makeM2Renderer(10);
    const record = vi.fn();
    const session = new PptxViewSession(root, reader, adapter, {
      positions: { initialSlideFor: vi.fn(() => 7), record },
    });

    await session.open("deck.pptx");

    expect(reader.readBinary).toHaveBeenCalledOnce();
    expect(rendererSession.renderSlide).toHaveBeenCalledTimes(1);
    expect(rendererSession.renderSlide).toHaveBeenCalledWith(7);
    expect(record).not.toHaveBeenCalled();
    root.querySelector<HTMLButtonElement>('[data-action="next-slide"]')!.click();
    await vi.waitFor(() => expect(root.textContent).toContain("9 / 10"));
    expect(record).toHaveBeenCalledOnce();
    expect(record).toHaveBeenCalledWith("deck.pptx", 8);
    session.dispose();
  });

  it("keeps reading when a position callback fails", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(3);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        positions: {
          initialSlideFor: () => 0,
          record: () => { throw new Error("private persistence detail"); },
        },
      },
    );
    await session.open("deck.pptx");

    root.querySelector<HTMLButtonElement>('[data-action="next-slide"]')!.click();

    await vi.waitFor(() => expect(root.textContent).toContain("2 / 3"));
    expect(root.dataset.state).toBe("ready");
    expect(root.textContent).not.toContain("private persistence detail");
    session.dispose();
  });

  it("keeps page state independent between sessions", async () => {
    const firstRoot = document.createElement("div");
    const secondRoot = document.createElement("div");
    document.body.append(firstRoot, secondRoot);
    const first = makeM2Renderer(4);
    const second = makeM2Renderer(4);
    const adapter: PptxRendererAdapter = {
      open: vi.fn().mockImplementationOnce(first.adapter.open).mockImplementationOnce(second.adapter.open),
    };
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const firstSession = new PptxViewSession(firstRoot, reader, adapter);
    const secondSession = new PptxViewSession(secondRoot, reader, adapter);
    await Promise.all([firstSession.open("a.pptx"), secondSession.open("b.pptx")]);

    firstRoot.querySelector<HTMLButtonElement>('[data-action="next-slide"]')!.click();
    await vi.waitFor(() => expect(firstRoot.textContent).toContain("2 / 4"));

    expect(secondRoot.textContent).toContain("1 / 4");
    expect(second.rendererSession.renderSlide).toHaveBeenCalledTimes(1);
    firstSession.dispose();
    secondSession.dispose();
  });

  it("supports accessible keyboard, thumbnail, collapse, and full-screen actions", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const { adapter, rendererSession } = makeM2Renderer(6);
    const fullscreen = makeFullscreen();
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    await session.open("deck.pptx");

    expect(document.activeElement).toBe(root);
    expect(root.tabIndex).toBe(0);
    expect(root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute("aria-label")).toBe("Enter full screen");
    expect(root.querySelector('[aria-label="Slide thumbnails"]')).not.toBeNull();
    expect(root.querySelector('[role="status"][aria-live="polite"]')).not.toBeNull();

    for (const [key, page] of [["ArrowRight", "2 / 6"], ["PageDown", "3 / 6"], ["ArrowLeft", "2 / 6"], ["PageUp", "1 / 6"]] as const) {
      const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key });
      root.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
      await vi.waitFor(() => expect(root.textContent).toContain(page));
    }
    const boundary = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "ArrowLeft" });
    root.dispatchEvent(boundary);
    expect(boundary.defaultPrevented).toBe(false);

    const input = root.querySelector<HTMLInputElement>('[data-action="page-number"]')!;
    const editableEvent = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "ArrowRight" });
    input.dispatchEvent(editableEvent);
    expect(editableEvent.defaultPrevented).toBe(false);
    expect(rendererSession.renderSlide).toHaveBeenCalledTimes(5);

    root.querySelector<HTMLButtonElement>('[data-slide-index="2"]')!.click();
    await vi.waitFor(() => expect(root.textContent).toContain("3 / 6"));
    expect(root.querySelector('[aria-current="page"]')?.getAttribute("aria-label")).toBe("Slide 3");

    root.querySelector<HTMLButtonElement>('[data-action="toggle-thumbnails"]')!.click();
    expect(root.dataset.thumbnailsCollapsed).toBe("true");
    expect(root.querySelector('[data-action="toggle-thumbnails"]')?.getAttribute("aria-expanded")).toBe("false");

    root.querySelector<HTMLButtonElement>('[data-action="toggle-fullscreen"]')!.click();
    await vi.waitFor(() => expect(root.dataset.fullscreen).toBe("true"));
    expect(fullscreen.api.enter).toHaveBeenCalledWith(root);
    expect(root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute("aria-label")).toBe("Exit full screen");
    fullscreen.api.isActive.mockReturnValueOnce(false);
    root.querySelector<HTMLButtonElement>('[data-action="toggle-fullscreen"]')!.click();
    await vi.waitFor(() => expect(root.dataset.fullscreen).toBe("false"));
    expect(fullscreen.api.exit).toHaveBeenCalledOnce();

    const actions = Array.from(root.querySelectorAll<HTMLElement>("button, input"), (item) => item.dataset.action).filter(Boolean);
    expect(actions.indexOf("previous-slide")).toBeLessThan(actions.indexOf("page-number"));
    expect(actions.indexOf("page-number")).toBeLessThan(actions.indexOf("toggle-fullscreen"));
    session.dispose();
  });

  it("keeps full-screen failures local and disposes session-owned work", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const { adapter, rendererSession } = makeM2Renderer(20);
    const fullscreen = makeFullscreen();
    fullscreen.api.enter.mockRejectedValueOnce(new Error("platform detail"));
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    await session.open("deck.pptx");

    root.querySelector<HTMLButtonElement>('[data-action="toggle-fullscreen"]')!.click();
    await vi.waitFor(() => expect(root.textContent).toContain("Unable to change full-screen mode."));
    expect(root.dataset.fullscreen).toBe("false");
    expect(root.textContent).not.toContain("platform detail");
    expect(session.getPerformanceDiagnostics()).toMatchObject({
      mountedThumbnails: expect.any(Number),
    });
    expect(Number(root.dataset.mountedThumbnailCount)).toBeGreaterThan(0);

    session.dispose();
    expect(fullscreen.listeners.size).toBe(0);
    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(session.getPerformanceDiagnostics()).toMatchObject({
      backgroundPending: 0,
      backgroundRunning: 0,
      mountedThumbnails: 0,
    });
    const afterDispose = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "ArrowRight" });
    root.dispatchEvent(afterDispose);
    expect(afterDispose.defaultPrevented).toBe(false);
  });
  it("starts in an explicit empty state before a file is selected", () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter } = makeRenderer();

    new PptxViewSession(root, reader, adapter);

    expect(root.dataset.state).toBe("empty");
    expect(root.dataset.lifecyclePhase).toBe("idle");
    expect(root.textContent).toContain(
      "Open a PPTX file from your Vault to start reading.",
    );
  });

  it("exposes a candidate-independent in-flight phase after Vault reading completes", async () => {
    const root = document.createElement("div");
    let finishRead: ((bytes: ArrayBuffer) => void) | undefined;
    let finishAdapterOpen: ((session: PptxRendererSession) => void) | undefined;
    const rendererSession = makeRenderer().rendererSession;
    const reader = {
      readBinary: vi.fn(
        () =>
          new Promise<ArrayBuffer>((resolve) => {
            finishRead = resolve;
          }),
      ),
    };
    const adapter: PptxRendererAdapter = {
      open: vi.fn(
        () =>
          new Promise<PptxRendererSession>((resolve) => {
            finishAdapterOpen = resolve;
          }),
      ),
    };
    const session = new PptxViewSession(root, reader, adapter);

    const opening = session.open("stress.pptx");
    expect(root.dataset.lifecyclePhase).toBe("reading");

    finishRead?.(new ArrayBuffer(1));
    await vi.waitFor(() => expect(adapter.open).toHaveBeenCalledOnce());
    expect(root.dataset.lifecyclePhase).toBe("adapter-opening");
    expect(session.getPerformanceDiagnostics().lifecyclePhase).toBe(
      "adapter-opening",
    );

    finishAdapterOpen?.(rendererSession);
    await opening;
    expect(root.dataset.lifecyclePhase).toBe("ready");

    session.dispose();
    expect(root.dataset.lifecyclePhase).toBeUndefined();
    expect(session.getPerformanceDiagnostics().lifecyclePhase).toBe("disposed");
  });

  it("exposes adapter work diagnostics through open and disposal", async () => {
    const root = document.createElement("div");
    let finishRead: ((value: ArrayBuffer) => void) | undefined;
    const reader = {
      readBinary: vi.fn(
        () =>
          new Promise<ArrayBuffer>((resolve) => {
            finishRead = resolve;
          }),
      ),
    };
    const { adapter } = makeRenderer();
    const session = new PptxViewSession(root, reader, adapter);

    const opening = session.open("deck.pptx");
    expect(session.getPerformanceDiagnostics()).toMatchObject({
      openPending: true,
      rendererActive: false,
      disposed: false,
    });
    expect(
      root.querySelector<HTMLInputElement>('[data-action="page-number"]')
        ?.disabled,
    ).toBe(true);
    expect(
      root.querySelector<HTMLButtonElement>('[data-action="jump-to-slide"]')
        ?.disabled,
    ).toBe(true);

    finishRead?.(new ArrayBuffer(1));
    await opening;
    expect(session.getPerformanceDiagnostics()).toMatchObject({
      openPending: false,
      rendererActive: true,
      disposed: false,
    });
    expect(
      root.querySelector<HTMLInputElement>('[data-action="page-number"]')
        ?.disabled,
    ).toBe(false);
    expect(
      root.querySelector<HTMLButtonElement>('[data-action="jump-to-slide"]')
        ?.disabled,
    ).toBe(false);

    session.dispose();
    expect(session.getPerformanceDiagnostics()).toMatchObject({
      openPending: false,
      rendererActive: false,
      disposed: true,
    });
  });

  it("reads once, renders slide 1, and exposes ready state", async () => {
    const root = document.createElement("div");
    const bytes = new ArrayBuffer(8);
    const reader = { readBinary: vi.fn(async () => bytes) };
    const { adapter, rendererSession } = makeRenderer(3);
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("deck.pptx");

    expect(reader.readBinary).toHaveBeenCalledOnce();
    expect(reader.readBinary).toHaveBeenCalledWith("deck.pptx", expect.any(AbortSignal));
    expect(adapter.open).toHaveBeenCalledOnce();
    expect(adapter.open).toHaveBeenCalledWith(
      bytes,
      expect.any(HTMLElement),
      expect.any(AbortSignal),
    );
    expect(rendererSession.renderSlide).toHaveBeenCalledWith(0);
    expect(root.dataset.state).toBe("ready");
    expect(root.textContent).toContain("1 / 3");
    expect(root.textContent).toContain("Obsidian PPTX smoke test");
  });

  it("opens a stable slide identity, discloses reordering, and copies canonical targets", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(8)) };
    const { adapter, rendererSession } = makeRenderer(3, [256, 300, 512]);
    const copy = vi.fn(async () => {});
    const positions = {
      initialSlideFor: vi.fn(() => 0),
      record: vi.fn(),
    };
    const session = new PptxViewSession(root, reader, adapter, {
      positions,
      slideReferences: { copy },
    });

    await session.open("deck.pptx", {
      slideId: 512,
      createdSlideNumber: 2,
    });

    expect(rendererSession.renderSlide).toHaveBeenCalledWith(2);
    expect(root.dataset.referenceSlideId).toBe("512");
    expect(root.dataset.referenceCreatedSlide).toBe("2");
    expect(root.dataset.referenceCurrentSlide).toBe("3");
    expect(root.textContent).toContain(
      "This reference was created for slide 2; the same slide is now slide 3.",
    );
    expect(positions.initialSlideFor).toHaveBeenCalledOnce();

    expect(root.querySelector(
      '[data-action="copy-slide-reference"]',
    )?.getAttribute("aria-label")).toBe("Copy slide reference");
    expect(root.querySelector(
      '[data-action="copy-slide-embed"]',
    )?.getAttribute("aria-label")).toBe("Copy slide embed");

    root.querySelector<HTMLButtonElement>(
      '[data-action="copy-slide-reference"]',
    )!.click();
    await vi.waitFor(() => expect(copy).toHaveBeenCalledWith(
      "deck.pptx",
      { slideId: 512, createdSlideNumber: 3 },
      false,
    ));
    root.querySelector<HTMLButtonElement>(
      '[data-action="copy-slide-embed"]',
    )!.click();
    await vi.waitFor(() => expect(copy).toHaveBeenLastCalledWith(
      "deck.pptx",
      { slideId: 512, createdSlideNumber: 3 },
      true,
    ));

    root.querySelector<HTMLButtonElement>(
      '[data-action="previous-slide"]',
    )!.click();
    await vi.waitFor(() => expect(root.textContent).toContain("2 / 3"));
    expect(root.querySelector(".pptx-viewer__reference-notice")?.textContent)
      .toBe("");
    root.querySelector<HTMLButtonElement>('[data-action="next-slide"]')!.click();
    await vi.waitFor(() => expect(root.textContent).toContain("3 / 3"));
    expect(root.querySelector(".pptx-viewer__reference-notice")?.textContent)
      .toContain("created for slide 2");
  });

  it("copies a slide-only text selection with its stable source", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const { adapter } = makeRenderer(1, [256]);
    const copySelection = vi.fn(async () => {});
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(8)) },
      adapter,
      {
        slideReferences: {
          copy: vi.fn(async () => {}),
          copySelection,
        },
      },
    );

    await session.open("deck.pptx");

    const toolbarAction = root.querySelector<HTMLButtonElement>(
      '[data-action="copy-selected-slide-text"]',
    )!;
    const floatingAction = root.querySelector<HTMLButtonElement>(
      '[data-action="copy-selected-slide-text-floating"]',
    )!;
    expect(toolbarAction.disabled).toBe(true);
    expect(floatingAction.hidden).toBe(true);

    const slide = root.querySelector<HTMLElement>(".pptx-viewer__slide")!;
    const text = slide.firstChild!;
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, 8);
    Object.defineProperty(range, "getBoundingClientRect", {
      value: () => ({
        bottom: 40,
        height: 20,
        left: 10,
        right: 90,
        top: 20,
        width: 80,
        x: 10,
        y: 20,
        toJSON: () => ({}),
      }),
    });
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));

    expect(toolbarAction.disabled).toBe(false);
    expect(floatingAction.hidden).toBe(false);
    expect(floatingAction.textContent).toBe("Copy selected text with source");
    const nativeCopy = new Event("copy", { bubbles: true, cancelable: true });
    slide.dispatchEvent(nativeCopy);
    expect(nativeCopy.defaultPrevented).toBe(false);

    toolbarAction.click();
    await vi.waitFor(() => expect(copySelection).toHaveBeenCalledWith(
      "deck.pptx",
      { slideId: 256, createdSlideNumber: 1 },
      "Obsidian",
    ));
    expect(root.querySelector(".pptx-viewer__action-status")?.textContent)
      .toBe("Selected text copied with slide reference.");

    selection.removeAllRanges();
    document.dispatchEvent(new Event("selectionchange"));
    expect(toolbarAction.disabled).toBe(true);
    expect(floatingAction.hidden).toBe(true);

    session.dispose();
    root.remove();
  });

  it("ignores selections outside the rendered slide", async () => {
    const root = document.createElement("div");
    const outside = document.createElement("p");
    outside.textContent = "Outside text";
    document.body.append(root, outside);
    const { adapter } = makeRenderer(1, [256]);
    const copySelection = vi.fn(async () => {});
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(8)) },
      adapter,
      {
        slideReferences: {
          copy: vi.fn(async () => {}),
          copySelection,
        },
      },
    );
    await session.open("deck.pptx");

    const range = document.createRange();
    range.selectNodeContents(outside);
    const selection = document.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));

    const toolbarAction = root.querySelector<HTMLButtonElement>(
      '[data-action="copy-selected-slide-text"]',
    )!;
    expect(toolbarAction.disabled).toBe(true);
    toolbarAction.click();
    expect(copySelection).not.toHaveBeenCalled();

    selection.removeAllRanges();
    session.dispose();
    root.remove();
    outside.remove();
  });

  it("fails a missing stable identity honestly without ordinal fallback", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(8)) };
    const { adapter, rendererSession } = makeRenderer(3, [256, 300, 512]);
    const session = new PptxViewSession(root, reader, adapter, {
      openExternally: vi.fn(async () => {}),
      slideReferences: { copy: vi.fn(async () => {}) },
    });

    await session.open("deck.pptx", {
      slideId: 999,
      createdSlideNumber: 2,
    });

    expect(rendererSession.renderSlide).not.toHaveBeenCalled();
    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.dataset.state).toBe("stale-reference");
    expect(root.textContent).toContain(
      "The referenced slide is no longer available in this presentation.",
    );
    expect(root.querySelector('[data-action="open-presentation"]')).not.toBeNull();
    expect(root.querySelector('[data-action="open-externally"]')).not.toBeNull();
  });

  it("exposes full-open timings and navigates with product controls", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter, rendererSession } = makeRenderer(3);
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("deck.pptx");

    expect(root.dataset.metadataMs).toMatch(/^\d/);
    expect(root.dataset.firstReadableMs).toMatch(/^\d/);
    const previous = root.querySelector<HTMLButtonElement>(
      '[data-action="previous-slide"]',
    );
    const next = root.querySelector<HTMLButtonElement>(
      '[data-action="next-slide"]',
    );
    expect(previous?.disabled).toBe(true);
    expect(next?.disabled).toBe(false);

    next?.click();

    await vi.waitFor(() =>
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(1),
    );
    await vi.waitFor(() => expect(root.textContent).toContain("2 / 3"));
    expect(root.dataset.lastSlideSwitchMs).toMatch(/^\d/);
    expect(previous?.disabled).toBe(false);
    expect(next?.disabled).toBe(false);
  });

  it("jumps to a valid one-based slide number", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter, rendererSession } = makeRenderer(3);
    const session = new PptxViewSession(root, reader, adapter);
    await session.open("deck.pptx");

    const input = root.querySelector<HTMLInputElement>(
      '[data-action="page-number"]',
    );
    const jump = root.querySelector<HTMLButtonElement>(
      '[data-action="jump-to-slide"]',
    );
    expect(input).not.toBeNull();
    expect(jump).not.toBeNull();

    input!.value = "3";
    jump!.click();

    await vi.waitFor(() =>
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(2),
    );
    await vi.waitFor(() => expect(root.textContent).toContain("3 / 3"));
    expect(root.dataset.state).toBe("ready");
    expect(input?.value).toBe("3");
  });

  it.each(["", "0", "4", "1.5"])(
    "rejects invalid page input %j without changing the readable slide",
    async (value) => {
      const root = document.createElement("div");
      const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
      const { adapter, rendererSession } = makeRenderer(3);
      const session = new PptxViewSession(root, reader, adapter);
      await session.open("deck.pptx");
      const input = root.querySelector<HTMLInputElement>(
        '[data-action="page-number"]',
      )!;

      input.value = value;
      root
        .querySelector<HTMLButtonElement>('[data-action="jump-to-slide"]')!
        .click();

      expect(rendererSession.renderSlide).toHaveBeenCalledTimes(1);
      expect(root.textContent).toContain("1 / 3");
      expect(root.textContent).toContain(
        "Enter a slide number from 1 to 3.",
      );
      expect(root.dataset.state).toBe("ready");
    },
  );

  it("captures slide-switch timing before updating product UI", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter } = makeRenderer(3);
    const session = new PptxViewSession(root, reader, adapter);
    const textSeenAtTimingBoundaries: string[] = [];
    const now = vi.spyOn(performance, "now").mockImplementation(() => {
      textSeenAtTimingBoundaries.push(root.textContent ?? "");
      return textSeenAtTimingBoundaries.length;
    });

    try {
      await session.open("deck.pptx");
      textSeenAtTimingBoundaries.length = 0;

      root
        .querySelector<HTMLButtonElement>('[data-action="next-slide"]')
        ?.click();

      await vi.waitFor(() =>
        expect(root.dataset.lastSlideSwitchMs).toMatch(/^\d/),
      );
      expect(textSeenAtTimingBoundaries.slice(0, 2)).toEqual([
        expect.stringContaining("1 / 3"),
        expect.stringContaining("1 / 3"),
      ]);
    } finally {
      now.mockRestore();
    }
  });

  it("disposes the renderer and clears the view", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter, rendererSession } = makeRenderer();
    const session = new PptxViewSession(root, reader, adapter);
    await session.open("deck.pptx");

    session.dispose();
    session.dispose();

    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.childElementCount).toBe(0);
    expect(root.dataset.metadataMs).toBeUndefined();
    expect(root.dataset.firstReadableMs).toBeUndefined();
    expect(root.dataset.lastSlideSwitchMs).toBeUndefined();
  });

  it("continues ordered teardown and clears the root when component disposers throw", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeM2Renderer(20);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");
    const internals = session as unknown as {
      thumbnailRail: { dispose(): void };
      viewerController: { dispose(): void };
      backgroundQueue: { dispose(): void };
    };
    const railDispose = vi.spyOn(internals.thumbnailRail, "dispose").mockImplementation(() => {
      throw new Error("rail cleanup failed");
    });
    const controllerDispose = vi.spyOn(internals.viewerController, "dispose");
    const queueDispose = vi.spyOn(internals.backgroundQueue, "dispose");
    vi.mocked(rendererSession.dispose).mockImplementation(() => {
      throw new Error("renderer cleanup failed");
    });

    expect(() => session.dispose()).not.toThrow();

    expect(railDispose).toHaveBeenCalledOnce();
    expect(controllerDispose).toHaveBeenCalledOnce();
    expect(queueDispose).toHaveBeenCalledOnce();
    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.childElementCount).toBe(0);
    expect(root.dataset.state).toBeUndefined();
    expect(session.getPerformanceDiagnostics()).toMatchObject({
      backgroundPending: 0,
      backgroundRunning: 0,
      mountedThumbnails: 0,
      rendererActive: false,
    });
  });

  it("detaches every owned resource before a cleanup callback re-enters dispose", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeM2Renderer(20);
    let session!: PptxViewSession<string>;
    let diagnosticsDuringCleanup: ReturnType<PptxViewSession<string>["getPerformanceDiagnostics"]> | undefined;
    let reentered = false;
    const unsubscribe = vi.fn(() => {
      diagnosticsDuringCleanup = session.getPerformanceDiagnostics();
      if (!reentered) {
        reentered = true;
        session.dispose();
      }
    });
    const fullscreen = makeFullscreen();
    fullscreen.api.subscribe.mockReturnValue(unsubscribe);
    session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    await session.open("deck.pptx");
    const internals = session as unknown as {
      thumbnailRail: { dispose(): void };
      viewerController: { dispose(): void };
      backgroundQueue: { dispose(): void };
    };
    const railDispose = vi.spyOn(internals.thumbnailRail, "dispose");
    const controllerDispose = vi.spyOn(internals.viewerController, "dispose");
    const queueDispose = vi.spyOn(internals.backgroundQueue, "dispose");

    expect(() => session.dispose()).not.toThrow();

    expect(diagnosticsDuringCleanup).toMatchObject({
      backgroundPending: 0,
      backgroundRunning: 0,
      mountedThumbnails: 0,
      rendererActive: false,
    });
    expect(unsubscribe).toHaveBeenCalledOnce();
    expect(railDispose).toHaveBeenCalledOnce();
    expect(controllerDispose).toHaveBeenCalledOnce();
    expect(queueDispose).toHaveBeenCalledOnce();
    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.childElementCount).toBe(0);
    expect(root.dataset.state).toBeUndefined();
  });

  it("reports a synchronous full-screen action failure locally", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(2);
    const fullscreen = makeFullscreen();
    fullscreen.api.enter.mockImplementation(() => {
      throw new Error("synchronous platform failure");
    });
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    await session.open("deck.pptx");

    root.querySelector<HTMLButtonElement>('[data-action="toggle-fullscreen"]')!.click();

    await vi.waitFor(() =>
      expect(root.textContent).toContain("Unable to change full-screen mode."),
    );
    expect(root.dataset.state).toBe("ready");
    expect(root.dataset.fullscreen).toBe("false");
    expect(root.textContent).not.toContain("synchronous platform failure");
    session.dispose();
  });

  it("keeps the viewer ready when the initial full-screen state probe throws", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(2);
    const fullscreen = makeFullscreen();
    fullscreen.api.isActive.mockImplementation(() => {
      throw new Error("initial full-screen probe failed");
    });
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );

    await expect(session.open("deck.pptx")).resolves.toBeUndefined();

    expect(root.dataset.state).toBe("ready");
    expect(root.dataset.fullscreen).toBe("false");
    expect(root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute("aria-label"))
      .toBe("Enter full screen");
    expect(root.textContent).not.toContain("initial full-screen probe failed");
    session.dispose();
  });

  it("isolates a throwing full-screen probe from the subscribed event", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(2);
    const fullscreen = makeFullscreen();
    fullscreen.api.isActive
      .mockReturnValueOnce(false)
      .mockImplementation(() => {
        throw new Error("event full-screen probe failed");
      });
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    await session.open("deck.pptx");

    expect(() => {
      fullscreen.listeners.forEach((listener) => listener());
    }).not.toThrow();

    expect(root.dataset.state).toBe("ready");
    expect(root.dataset.fullscreen).toBe("false");
    expect(root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute("aria-label"))
      .toBe("Enter full screen");
    session.dispose();
  });

  it("keeps full-screen failure status stable when the recovery probe throws", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(2);
    const fullscreen = makeFullscreen();
    fullscreen.api.isActive
      .mockReturnValueOnce(false)
      .mockReturnValueOnce(false)
      .mockImplementation(() => {
        throw new Error("recovery full-screen probe failed");
      });
    fullscreen.api.enter.mockRejectedValueOnce(new Error("enter failed"));
    const unhandled = vi.fn();
    process.on("unhandledRejection", unhandled);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    try {
      await session.open("deck.pptx");

      root.querySelector<HTMLButtonElement>('[data-action="toggle-fullscreen"]')!.click();

      await vi.waitFor(() =>
        expect(root.textContent).toContain("Unable to change full-screen mode."),
      );
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(unhandled).not.toHaveBeenCalled();
      expect(root.dataset.fullscreen).toBe("false");
      expect(root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute("aria-label"))
        .toBe("Enter full screen");
      expect(root.textContent).not.toContain("recovery full-screen probe failed");
    } finally {
      process.off("unhandledRejection", unhandled);
      session.dispose();
    }
  });

  it("retains the last known state without rejection when a successful toggle recovery probe throws", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(2);
    const fullscreen = makeFullscreen();
    fullscreen.api.isActive
      .mockReturnValueOnce(false)
      .mockReturnValueOnce(false)
      .mockImplementation(() => {
        throw new Error("successful recovery probe failed");
      });
    const unhandled = vi.fn();
    process.on("unhandledRejection", unhandled);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      { fullscreen: fullscreen.api },
    );
    try {
      await session.open("deck.pptx");

      root.querySelector<HTMLButtonElement>('[data-action="toggle-fullscreen"]')!.click();

      await vi.waitFor(() => expect(fullscreen.api.enter).toHaveBeenCalledOnce());
      await new Promise((resolve) => setTimeout(resolve, 0));
      expect(unhandled).not.toHaveBeenCalled();
      expect(root.dataset.fullscreen).toBe("false");
      expect(root.querySelector('[data-action="toggle-fullscreen"]')?.getAttribute("aria-label"))
        .toBe("Enter full screen");
      expect(root.textContent).not.toContain("Unable to change full-screen mode.");
      expect(root.textContent).not.toContain("successful recovery probe failed");
    } finally {
      process.off("unhandledRejection", unhandled);
      session.dispose();
    }
  });

  it("survives a throwing previous renderer disposer while reopening", async () => {
    const root = document.createElement("div");
    const first = makeRenderer(2);
    const second = makeRenderer(2);
    vi.mocked(first.rendererSession.dispose).mockImplementation(() => {
      throw new Error("old renderer cleanup failed");
    });
    const adapter: PptxRendererAdapter = {
      open: vi.fn().mockImplementationOnce(first.adapter.open).mockImplementationOnce(second.adapter.open),
    };
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("first.pptx");

    await expect(session.open("second.pptx")).resolves.toBeUndefined();

    expect(first.rendererSession.dispose).toHaveBeenCalledOnce();
    expect(second.rendererSession.dispose).not.toHaveBeenCalled();
    expect(root.dataset.state).toBe("ready");
    expect(root.textContent).toContain("1 / 2");
    session.dispose();
  });

  it("finishes the open-error surface when renderer cleanup throws", async () => {
    const root = document.createElement("div");
    const { adapter, rendererSession } = makeRenderer(1);
    vi.mocked(rendererSession.renderSlide).mockRejectedValueOnce(new Error("render failed"));
    vi.mocked(rendererSession.dispose).mockImplementation(() => {
      throw new Error("renderer cleanup failed");
    });
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );

    await expect(session.open("broken.pptx")).resolves.toBeUndefined();

    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.dataset.state).toBe("error");
    expect(root.dataset.errorCategory).toBe("incompatible");
    expect(session.getPerformanceDiagnostics().rendererActive).toBe(false);
  });

  it("keeps the mounted-thumbnail data attribute live and stale-safe", async () => {
    const root = document.createElement("div");
    const { adapter } = makeM2Renderer(200);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("large.pptx");
    const rail = root.querySelector<HTMLElement>('[aria-label="Slide thumbnails"]')!;
    const initial = Number(root.dataset.mountedThumbnailCount);

    Object.defineProperty(rail, "clientHeight", { configurable: true, value: 100 });
    rail.dispatchEvent(new Event("scroll"));

    expect(Number(root.dataset.mountedThumbnailCount)).toBe(
      rail.querySelectorAll('[data-action="thumbnail-slide"]').length,
    );
    expect(Number(root.dataset.mountedThumbnailCount)).not.toBe(initial);

    session.dispose();
    expect(root.dataset.mountedThumbnailCount).toBeUndefined();
    rail.dispatchEvent(new Event("scroll"));
    expect(root.dataset.mountedThumbnailCount).toBeUndefined();
  });

  it("exposes a thumbnail-ready count only after readiness and clears it on dispose", async () => {
    const root = document.createElement("div");
    const pending = new Promise<void>((resolve) => {
      (root as HTMLElement & { resolveThumbnail?: () => void }).resolveThumbnail = resolve;
    });
    const { adapter, rendererSession } = makeM2Renderer(2);
    rendererSession.renderThumbnail = vi.fn((index, container) => {
      container.textContent = `Allocated preview ${index + 1}`;
      return { ready: pending, dispose: vi.fn() };
    });
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );

    await session.open("ready-signal.pptx");
    await vi.waitFor(() => expect(rendererSession.renderThumbnail).toHaveBeenCalled());
    expect(root.dataset.readyThumbnailCount).toBe("0");

    (root as HTMLElement & { resolveThumbnail?: () => void }).resolveThumbnail?.();
    await vi.waitFor(() => expect(Number(root.dataset.readyThumbnailCount)).toBeGreaterThan(0));
    expect(session.getPerformanceDiagnostics().readyThumbnails).toBeGreaterThan(0);

    session.dispose();
    expect(root.dataset.readyThumbnailCount).toBeUndefined();
    expect(session.getPerformanceDiagnostics().readyThumbnails).toBe(0);
  });

  it("disposes the previous renderer before reopening", async () => {
    const root = document.createElement("div");
    const bytes = new ArrayBuffer(1);
    let finishSecondRead: (() => void) | undefined;
    const reader = {
      readBinary: vi
        .fn()
        .mockResolvedValueOnce(bytes)
        .mockImplementationOnce(
          () =>
            new Promise<ArrayBuffer>((resolve) => {
              finishSecondRead = () => resolve(bytes);
            }),
        ),
    };
    const first = makeRenderer();
    const second = makeRenderer();
    const adapter: PptxRendererAdapter = {
      open: vi
        .fn()
        .mockImplementationOnce(first.adapter.open)
        .mockImplementationOnce(second.adapter.open),
    };
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("first.pptx");
    root.dataset.lastSlideSwitchMs = "12.345";
    const reopening = session.open("second.pptx");

    expect(first.rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.dataset.metadataMs).toBeUndefined();
    expect(root.dataset.firstReadableMs).toBeUndefined();
    expect(root.dataset.lastSlideSwitchMs).toBeUndefined();

    finishSecondRead?.();
    await reopening;

    expect(second.rendererSession.dispose).not.toHaveBeenCalled();
    expect(root.dataset.metadataMs).toMatch(/^\d/);
    expect(root.dataset.firstReadableMs).toMatch(/^\d/);
    expect(root.dataset.lastSlideSwitchMs).toBeUndefined();
  });

  it("does not apply a completed navigation after reopening", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    let finishNavigation: (() => void) | undefined;
    const first = makeRenderer(3);
    first.rendererSession.renderSlide = vi
      .fn()
      .mockResolvedValueOnce(undefined)
      .mockImplementationOnce(
        () =>
          new Promise<void>((resolve) => {
            finishNavigation = resolve;
          }),
      );
    const second = makeRenderer(2);
    const adapter: PptxRendererAdapter = {
      open: vi
        .fn()
        .mockImplementationOnce(first.adapter.open)
        .mockImplementationOnce(second.adapter.open),
    };
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("first.pptx");
    root
      .querySelector<HTMLButtonElement>('[data-action="next-slide"]')
      ?.click();
    await vi.waitFor(() =>
      expect(first.rendererSession.renderSlide).toHaveBeenLastCalledWith(1),
    );

    await session.open("second.pptx");
    finishNavigation?.();
    await Promise.resolve();

    expect(root.textContent).toContain("1 / 2");
    expect(root.dataset.lastSlideSwitchMs).toBeUndefined();
  });

  it("reports navigation failures and restores the controls", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter, rendererSession } = makeRenderer(3);
    rendererSession.renderSlide = vi
      .fn()
      .mockResolvedValueOnce(undefined)
      .mockRejectedValueOnce(new Error("render failed"))
      .mockResolvedValueOnce(undefined);
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("deck.pptx");
    const previous = root.querySelector<HTMLButtonElement>(
      '[data-action="previous-slide"]',
    );
    const next = root.querySelector<HTMLButtonElement>(
      '[data-action="next-slide"]',
    );

    next?.click();

    await vi.waitFor(() => expect(root.dataset.state).toBe("degraded"));
    expect(root.textContent).toContain(
      "Slide 2 could not be rendered. The previous slide is still shown. Try another slide or open it in the default application.",
    );
    expect(root.textContent).toContain("Obsidian PPTX smoke test");
    expect(root.textContent).toContain("1 / 3");
    expect(
      root.querySelector<HTMLInputElement>('[data-action="page-number"]')
        ?.value,
    ).toBe("1");
    expect(root.dataset.lastSlideSwitchMs).toBeUndefined();
    expect(previous?.disabled).toBe(true);
    expect(next?.disabled).toBe(false);

    next?.click();

    await vi.waitFor(() => expect(root.textContent).toContain("2 / 3"));
    expect(root.dataset.state).toBe("ready");
    expect(root.textContent).not.toContain("Unable to render this slide.");
    expect(root.dataset.lastSlideSwitchMs).toMatch(/^\d/);
    expect(previous?.disabled).toBe(false);
    expect(next?.disabled).toBe(false);
  });

  it("ignores a rejected navigation after disposal", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    let failNavigation: ((error: Error) => void) | undefined;
    const { adapter, rendererSession } = makeRenderer(3);
    rendererSession.renderSlide = vi
      .fn()
      .mockResolvedValueOnce(undefined)
      .mockImplementationOnce(
        () =>
          new Promise<void>((_resolve, reject) => {
            failNavigation = reject;
          }),
      );
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("deck.pptx");
    root
      .querySelector<HTMLButtonElement>('[data-action="next-slide"]')
      ?.click();
    await vi.waitFor(() =>
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(1),
    );

    session.dispose();
    failNavigation?.(new Error("disposed"));
    await Promise.resolve();

    expect(root.childElementCount).toBe(0);
    expect(root.dataset.state).toBeUndefined();
    expect(root.dataset.lastSlideSwitchMs).toBeUndefined();
  });

  const errorCases = [
    ["unsupported-legacy", "Legacy PPT files are not supported."],
    ["malformed", "This PPTX is damaged or incomplete."],
    ["protected", "This PPTX is encrypted or password-protected."],
    [
      "incompatible",
      "This PPTX uses content this viewer cannot safely display.",
    ],
    [
      "resource-exhausted",
      "This PPTX is too large or complex to open within the viewer's safety limits.",
    ],
    ["cancelled", "Loading this PPTX was cancelled."],
    ["unknown", "An unexpected error prevented this PPTX from opening."],
  ] as const;

  for (const [category, message] of errorCases) {
    it(`shows a stable ${category} failure without rejecting the view load`, async () => {
      const root = document.createElement("div");
      const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
      const adapter: PptxRendererAdapter = {
        open: vi.fn(async () => {
          throw new PptxOpenError(category, "candidate-specific details");
        }),
      };
      const session = new PptxViewSession(root, reader, adapter);

      await expect(session.open("deck.pptx")).resolves.toBeUndefined();

      expect(root.dataset.state).toBe("error");
      expect(root.dataset.errorCategory).toBe(category);
      expect(root.textContent).toContain(message);
      expect(root.textContent).toContain(
        category === "unsupported-legacy"
          ? "The original file was not modified."
          : "The original PPTX file was not modified.",
      );
      expect(
        root.querySelector<HTMLButtonElement>('[data-action="retry"]'),
      ).not.toBeNull();
      expect(root.textContent).not.toContain("candidate-specific details");
    });
  }

  it("classifies an unexpected Vault read failure as unknown", async () => {
    const root = document.createElement("div");
    const reader = {
      readBinary: vi.fn(async () => {
        throw new Error("private filesystem details");
      }),
    };
    const { adapter } = makeRenderer();
    const session = new PptxViewSession(root, reader, adapter);

    await expect(session.open("deck.pptx")).resolves.toBeUndefined();

    expect(root.dataset.errorCategory).toBe("unknown");
    expect(root.textContent).not.toContain("private filesystem details");
  });

  it("retries the same file and reaches ready after a recoverable failure", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const recovered = makeRenderer(2);
    const adapter: PptxRendererAdapter = {
      open: vi
        .fn()
        .mockRejectedValueOnce(new PptxOpenError("malformed", "first failure"))
        .mockImplementationOnce(recovered.adapter.open),
    };
    const session = new PptxViewSession(root, reader, adapter);
    await session.open("deck.pptx");
    const brand = root.querySelector("[data-office-viewer-brand]");
    expect(brand).not.toBeNull();
    expect(root.querySelector(".office-viewer-brand__format")?.textContent)
      .toBe("PPTX");
    expect(root.querySelectorAll("[data-office-viewer-brand]")).toHaveLength(1);

    root.querySelector<HTMLButtonElement>('[data-action="retry"]')?.click();

    await vi.waitFor(() => expect(root.dataset.state).toBe("ready"));
    expect(reader.readBinary).toHaveBeenCalledTimes(2);
    expect(root.textContent).toContain("1 / 2");
    expect(root.dataset.errorCategory).toBeUndefined();
    expect(root.querySelector("[data-office-viewer-brand]")).toBe(brand);
    expect(root.querySelectorAll("[data-office-viewer-brand]")).toHaveLength(1);
  });

  it("offers the injected default-application fallback", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const adapter: PptxRendererAdapter = {
      open: vi.fn(async () => {
        throw new PptxOpenError("protected", "protected");
      }),
    };
    const openExternally = vi.fn(async () => {});
    const session = new PptxViewSession(root, reader, adapter, {
      openExternally,
    });
    await session.open("deck.pptx");

    root
      .querySelector<HTMLButtonElement>('[data-action="open-externally"]')
      ?.click();

    await vi.waitFor(() =>
      expect(openExternally).toHaveBeenCalledWith("deck.pptx"),
    );
  });

  it("offers the default-application fallback while a deck is readable", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const { adapter } = makeRenderer(2);
    const openExternally = vi.fn(async () => {});
    const session = new PptxViewSession(root, reader, adapter, {
      openExternally,
    });
    await session.open("deck.pptx");

    root
      .querySelector<HTMLButtonElement>('[data-action="open-externally"]')
      ?.click();

    await vi.waitFor(() =>
      expect(openExternally).toHaveBeenCalledWith("deck.pptx"),
    );
    expect(root.dataset.state).toBe("ready");
  });

  it("disposes renderer resources when first-slide rendering fails", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const rendererSession: PptxRendererSession = {
      slideCount: 1,
      slideWidth: 960,
      slideHeight: 540,
      capabilities: { thumbnails: false, prefetch: false },
      renderSlide: vi.fn(async () => {
        throw new Error("renderer exploded");
      }),
      dispose: vi.fn(),
    };
    const adapter: PptxRendererAdapter = {
      open: vi.fn(async () => rendererSession),
    };
    const session = new PptxViewSession(root, reader, adapter);

    await session.open("deck.pptx");

    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.dataset.errorCategory).toBe("incompatible");
  });

  it("disposes a late renderer result after the view closes", async () => {
    const root = document.createElement("div");
    const reader = { readBinary: vi.fn(async () => new ArrayBuffer(1)) };
    const rendererSession: PptxRendererSession = {
      slideCount: 1,
      slideWidth: 960,
      slideHeight: 540,
      capabilities: { thumbnails: false, prefetch: false },
      renderSlide: vi.fn(async () => {}),
      dispose: vi.fn(() => {
        throw new Error("late candidate cleanup failed");
      }),
    };
    let resolveOpen!: (value: PptxRendererSession) => void;
    const adapter: PptxRendererAdapter = {
      open: vi.fn(
        async () =>
          new Promise<PptxRendererSession>((resolve) => {
            resolveOpen = resolve;
          }),
      ),
    };
    const session = new PptxViewSession(root, reader, adapter);
    const opening = session.open("deck.pptx");
    await vi.waitFor(() => expect(adapter.open).toHaveBeenCalledOnce());

    session.dispose();
    resolveOpen(rendererSession);
    await expect(opening).resolves.toBeUndefined();

    expect(rendererSession.dispose).toHaveBeenCalledOnce();
    expect(root.childElementCount).toBe(0);
  });

  it("keeps the speaker-notes panel collapsed until toggled and preserves that choice while navigating", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      2,
      [256, 257],
      undefined,
      [
        {
          slideId: 256,
          paragraphs: ["AUTHOR_NOTE_P1", "AUTHOR_NOTE_P2"],
        },
        { slideId: 257, paragraphs: [] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");

    const toggle = root.querySelector<HTMLButtonElement>(
      '[data-action="toggle-notes"]',
    );
    const panel = root.querySelector<HTMLElement>(".pptx-viewer__notes-panel");
    expect(toggle).not.toBeNull();
    expect(panel).not.toBeNull();
    expect(root.dataset.notesCollapsed).toBe("true");
    expect(toggle?.getAttribute("aria-expanded")).toBe("false");
    expect(panel?.getAttribute("aria-hidden")).toBe("true");
    expect(panel?.getAttribute("aria-label")).toBe("Speaker notes");
    expect(
      root.querySelector('[data-notes-state="ready"]')?.textContent,
    ).toContain("AUTHOR_NOTE_P1");
    expect(
      root.querySelector('[data-notes-state="ready"]')?.textContent,
    ).toContain("AUTHOR_NOTE_P2");

    toggle?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(root.dataset.notesCollapsed).toBe("false");
    expect(toggle?.getAttribute("aria-expanded")).toBe("true");
    expect(panel?.getAttribute("aria-hidden")).toBe("false");

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
    );
    await vi.waitFor(() =>
      expect(root.querySelector('[data-notes-state="empty"]')?.textContent)
        .toBe("This slide has no speaker notes.")
    );
    expect(root.dataset.notesCollapsed).toBe("false");

    await session.open("other.pptx");
    expect(root.dataset.notesCollapsed).toBe("true");
    expect(
      root.querySelector('[data-action="toggle-notes"]')?.getAttribute(
        "aria-expanded",
      ),
    ).toBe("false");

    session.dispose();
    expect(root.querySelector(".pptx-viewer__notes-panel")).toBeNull();
  });

  it("announces unavailable speaker notes when optional extraction failed", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(1, [256], undefined, undefined);
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");

    expect(
      root.querySelector('[data-notes-state="unavailable"]')?.textContent,
    ).toBe("Speaker notes are unavailable for this presentation.");
    session.dispose();
  });

  it("copies current-slide speaker notes with the canonical slide reference", async () => {
    const root = document.createElement("div");
    const copyNotes = vi.fn(async () => {});
    const { adapter } = makeRenderer(
      2,
      [256, 257],
      [
        { slideId: 256, text: ["Slide one"] },
        { slideId: 257, text: ["Slide two"] },
      ],
      [
        { slideId: 256, paragraphs: ["Note A", "Note B"] },
        { slideId: 257, paragraphs: [] },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        slideReferences: {
          copy: vi.fn(async () => {}),
          copyNotes,
        },
      },
    );
    await session.open("deck.pptx");

    const copyButton = root.querySelector<HTMLButtonElement>(
      '[data-action="copy-speaker-notes"]',
    );
    expect(copyButton?.disabled).toBe(false);
    copyButton?.click();
    await vi.waitFor(() => expect(copyNotes).toHaveBeenCalledOnce());
    expect(copyNotes).toHaveBeenCalledWith(
      "deck.pptx",
      { slideId: 256, createdSlideNumber: 1 },
      ["Note A", "Note B"],
    );
    expect(root.querySelector(".pptx-viewer__action-status")?.textContent)
      .toBe("Speaker notes copied with slide reference.");

    root.dispatchEvent(
      new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
    );
    await vi.waitFor(() =>
      expect(
        root.querySelector<HTMLButtonElement>(
          '[data-action="copy-speaker-notes"]',
        )?.disabled,
      ).toBe(true)
    );
    session.dispose();
  });

  it("reports a clipboard failure without changing the current speaker notes", async () => {
    const root = document.createElement("div");
    const copyNotes = vi.fn(async () => {
      throw new Error("clipboard unavailable");
    });
    const { adapter, rendererSession } = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Slide one"] }],
      [{ slideId: 256, paragraphs: ["Note A", "Note B"] }],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
      {
        slideReferences: {
          copy: vi.fn(async () => {}),
          copyNotes,
        },
      },
    );
    await session.open("deck.pptx");

    root.querySelector<HTMLButtonElement>(
      '[data-action="copy-speaker-notes"]',
    )!.click();
    await vi.waitFor(() =>
      expect(root.querySelector(".pptx-viewer__action-status")?.textContent)
        .toBe("Unable to copy the speaker notes.")
    );
    expect(copyNotes).toHaveBeenCalledWith(
      "deck.pptx",
      { slideId: 256, createdSlideNumber: 1 },
      ["Note A", "Note B"],
    );
    expect(rendererSession.renderSlide).toHaveBeenCalledTimes(1);
    expect(root.querySelector('[data-notes-state="ready"]')?.textContent)
      .toContain("Note A");
    expect(
      root.querySelector<HTMLButtonElement>(
        '[data-action="copy-speaker-notes"]',
      )?.disabled,
    ).toBe(false);
    session.dispose();
  });

  it.each([
    {
      language: "en",
      toggle: "Speaker notes",
      toggleLabel: "Toggle speaker notes",
      panelLabel: "Speaker notes",
      empty: "This slide has no speaker notes.",
      unavailable: "Speaker notes are unavailable for this presentation.",
      copy: "Copy speaker notes",
      copied: "Speaker notes copied with slide reference.",
      copyFailure: "Unable to copy the speaker notes.",
      searchOpen: "Search presentation content",
      searchClose: "Close presentation search",
      searchInput: "Search presentation content",
      searchPlaceholder: "Search slides and speaker notes",
      scopeLabel: "Search scope",
      scopeAll: "All",
      scopeSlides: "Slides",
      scopeNotes: "Notes",
      provenanceSlide: "Slide text",
      provenanceNotes: "Speaker notes",
      slideMatchLabel: "Slide 1, slide text matches: 1",
      notesMatchLabel: "Slide 1, speaker-note matches: 1",
      noResultsAll:
        "No matching slide text or speaker notes. Images, charts, and SmartArt are not searched.",
      noResultsSlides:
        "No matching slide text. Images, charts, SmartArt, and speaker notes are outside this scope.",
      noResultsNotes:
        "No matching speaker notes. Slide text, images, charts, and SmartArt are outside this scope.",
    },
    {
      language: "zh-CN",
      toggle: "讲者备注",
      toggleLabel: "切换讲者备注",
      panelLabel: "讲者备注",
      empty: "此幻灯片没有讲者备注。",
      unavailable: "无法提供此演示文稿的讲者备注。",
      copy: "复制讲者备注",
      copied: "已复制讲者备注及幻灯片引用。",
      copyFailure: "无法复制讲者备注。",
      searchOpen: "搜索演示文稿内容",
      searchClose: "关闭演示文稿内容搜索",
      searchInput: "搜索演示文稿内容",
      searchPlaceholder: "搜索幻灯片与讲者备注",
      scopeLabel: "搜索范围",
      scopeAll: "全部",
      scopeSlides: "幻灯片",
      scopeNotes: "讲者备注",
      provenanceSlide: "幻灯片文字",
      provenanceNotes: "讲者备注",
      slideMatchLabel: "第 1 张幻灯片，幻灯片文字 1 处匹配",
      notesMatchLabel: "第 1 张幻灯片，讲者备注 1 处匹配",
      noResultsAll:
        "未在幻灯片文字或讲者备注中找到结果；图片、图表和 SmartArt 不在搜索范围内。",
      noResultsSlides:
        "未在幻灯片文字中找到结果；图片、图表、SmartArt 和讲者备注不在此范围内。",
      noResultsNotes:
        "未在讲者备注中找到结果；幻灯片文字、图片、图表和 SmartArt 不在此范围内。",
    },
    {
      language: "zh-TW",
      toggle: "講者備註",
      toggleLabel: "切換講者備註",
      panelLabel: "講者備註",
      empty: "此投影片沒有講者備註。",
      unavailable: "無法提供此簡報的講者備註。",
      copy: "複製講者備註",
      copied: "已複製講者備註及投影片引用。",
      copyFailure: "無法複製講者備註。",
      searchOpen: "搜尋簡報內容",
      searchClose: "關閉簡報內容搜尋",
      searchInput: "搜尋簡報內容",
      searchPlaceholder: "搜尋投影片與講者備註",
      scopeLabel: "搜尋範圍",
      scopeAll: "全部",
      scopeSlides: "投影片",
      scopeNotes: "講者備註",
      provenanceSlide: "投影片文字",
      provenanceNotes: "講者備註",
      slideMatchLabel: "第 1 張投影片，投影片文字 1 處相符",
      notesMatchLabel: "第 1 張投影片，講者備註 1 處相符",
      noResultsAll:
        "在投影片文字或講者備註中找不到結果；圖片、圖表和 SmartArt 不在搜尋範圍內。",
      noResultsSlides:
        "在投影片文字中找不到結果；圖片、圖表、SmartArt 和講者備註不在此範圍內。",
      noResultsNotes:
        "在講者備註中找不到結果；投影片文字、圖片、圖表和 SmartArt 不在此範圍內。",
    },
  ] as const)(
    "renders the complete speaker-note and Presentation search surface in $language",
    async (expected) => {
      const root = document.createElement("div");
      let copyAttempt = 0;
      const copyNotes = vi.fn(async () => {
        copyAttempt += 1;
        if (copyAttempt === 2) throw new Error("clipboard unavailable");
      });
      const { adapter } = makeRenderer(
        2,
        [256, 257],
        [
          { slideId: 256, text: ["Shared author marker"] },
          { slideId: 257, text: ["Other slide"] },
        ],
        [
          { slideId: 256, paragraphs: ["Shared author marker"] },
          { slideId: 257, paragraphs: [] },
        ],
      );
      const options = {
        messages: createMessageTranslator(expected.language),
        slideReferences: {
          copy: vi.fn(async () => {}),
          copyNotes,
        },
      };
      const session = new PptxViewSession(
        root,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        adapter,
        options,
      );
      await session.open("deck.pptx");

      const toggle = root.querySelector<HTMLButtonElement>(
        '[data-action="toggle-notes"]',
      )!;
      expect(toggle.textContent).toBe(expected.toggle);
      expect(toggle.getAttribute("aria-label")).toBe(expected.toggleLabel);
      expect(root.querySelector(".pptx-viewer__notes-panel")?.getAttribute(
        "aria-label",
      )).toBe(expected.panelLabel);
      const copy = root.querySelector<HTMLButtonElement>(
        '[data-action="copy-speaker-notes"]',
      )!;
      expect(copy.getAttribute("aria-label")).toBe(expected.copy);
      copy.click();
      await vi.waitFor(() =>
        expect(root.querySelector(".pptx-viewer__action-status")?.textContent)
          .toBe(expected.copied)
      );
      copy.click();
      await vi.waitFor(() =>
        expect(root.querySelector(".pptx-viewer__action-status")?.textContent)
          .toBe(expected.copyFailure)
      );

      const searchButton = root.querySelector<HTMLButtonElement>(
        '[data-action="open-slide-search"]',
      )!;
      expect(searchButton.getAttribute("aria-label")).toBe(expected.searchOpen);
      searchButton.click();
      expect(searchButton.getAttribute("aria-label")).toBe(expected.searchClose);
      expect(root.querySelector('[role="search"]')?.getAttribute("aria-label"))
        .toBe(expected.searchOpen);
      const input = root.querySelector<HTMLInputElement>(
        '[data-action="slide-search-input"]',
      )!;
      expect(input.getAttribute("aria-label")).toBe(expected.searchInput);
      expect(input.placeholder).toBe(expected.searchPlaceholder);
      const scopeGroup = root.querySelector('[role="group"]')!;
      expect(scopeGroup.getAttribute("aria-label")).toBe(expected.scopeLabel);
      expect(root.querySelector(
        '[data-action="search-scope"][data-search-scope="all"]',
      )?.textContent)
        .toBe(expected.scopeAll);
      expect(root.querySelector(
        '[data-action="search-scope"][data-search-scope="slides"]',
      )?.textContent)
        .toBe(expected.scopeSlides);
      expect(root.querySelector(
        '[data-action="search-scope"][data-search-scope="notes"]',
      )?.textContent)
        .toBe(expected.scopeNotes);

      input.value = "shared author marker";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      const slideMatch = root.querySelector<HTMLElement>(
        '[data-action="slide-search-slide-match"]',
      )!;
      const notesMatch = root.querySelector<HTMLElement>(
        '[data-action="slide-search-notes-match"]',
      )!;
      expect(slideMatch.getAttribute("aria-label"))
        .toBe(expected.slideMatchLabel);
      expect(notesMatch.getAttribute("aria-label"))
        .toBe(expected.notesMatchLabel);
      expect(slideMatch.querySelector(
        ".pptx-viewer__slide-search-provenance",
      )?.textContent).toBe(expected.provenanceSlide);
      expect(notesMatch.querySelector(
        ".pptx-viewer__slide-search-provenance",
      )?.textContent).toBe(expected.provenanceNotes);

      input.value = "missing";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      const summary = root.querySelector(
        ".pptx-viewer__slide-search-summary",
      );
      expect(summary?.textContent).toBe(expected.noResultsAll);
      root.querySelector<HTMLButtonElement>(
        '[data-action="search-scope"][data-search-scope="slides"]',
      )!.click();
      expect(summary?.textContent).toBe(expected.noResultsSlides);
      root.querySelector<HTMLButtonElement>(
        '[data-action="search-scope"][data-search-scope="notes"]',
      )!.click();
      expect(summary?.textContent).toBe(expected.noResultsNotes);

      root.dispatchEvent(
        new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }),
      );
      await vi.waitFor(() =>
        expect(root.querySelector('[data-notes-state="empty"]')?.textContent)
          .toBe(expected.empty)
      );
      session.dispose();

      const unavailableRoot = document.createElement("div");
      const { adapter: unavailableAdapter } = makeRenderer(
        1,
        [256],
        [{ slideId: 256, text: ["Slide"] }],
        undefined,
      );
      const unavailableSession = new PptxViewSession(
        unavailableRoot,
        { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
        unavailableAdapter,
        { messages: createMessageTranslator(expected.language) },
      );
      await unavailableSession.open("unavailable.pptx");
      expect(
        unavailableRoot.querySelector(
          '[data-notes-state="unavailable"]',
        )?.textContent,
      ).toBe(expected.unavailable);
      unavailableSession.dispose();
    },
  );

  it("searches notes-only markers, expands the panel, and highlights the match", async () => {
    const root = document.createElement("div");
    document.body.append(root);
    const { adapter, rendererSession } = makeRenderer(
      2,
      [256, 257],
      [
        { slideId: 256, text: ["Canvas only"] },
        { slideId: 257, text: ["Other canvas"] },
      ],
      [
        { slideId: 256, paragraphs: [] },
        {
          slideId: 257,
          paragraphs: ["Hidden NOTE_ONLY marker", "Trailing note"],
        },
      ],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");

    expect(
      root.querySelector('[data-action="open-slide-search"]')?.getAttribute(
        "aria-label",
      ),
    ).toBe("Search presentation content");

    root.dispatchEvent(new KeyboardEvent("keydown", {
      key: "f",
      ctrlKey: true,
      bubbles: true,
    }));
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    input.value = "NOTE_ONLY";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    expect(root.querySelector('[data-search-scope="all"]')).not.toBeNull();
    const notesMatch = root.querySelector<HTMLButtonElement>(
      '[data-action="slide-search-notes-match"]',
    );
    expect(notesMatch?.textContent).toContain("Speaker notes");
    expect(notesMatch?.textContent).toContain("NOTE_ONLY");
    notesMatch?.click();

    await vi.waitFor(() => {
      expect(rendererSession.renderSlide).toHaveBeenLastCalledWith(1);
    });
    expect(root.dataset.notesCollapsed).toBe("false");
    expect(
      root.querySelector(".pptx-viewer__notes-highlight")?.textContent,
    ).toBe("NOTE_ONLY");

    root.querySelector<HTMLButtonElement>(
      '[data-search-scope="slides"]',
    )?.click();
    expect(
      root.querySelector('[data-action="slide-search-notes-match"]'),
    ).toBeNull();
    expect(root.querySelector(".pptx-viewer__slide-search-summary")?.textContent)
      .toContain("outside this scope");

    session.dispose();
    root.remove();
  });

  it("highlights the raw note range for a normalized search match", async () => {
    const root = document.createElement("div");
    const { adapter } = makeRenderer(
      1,
      [256],
      [{ slideId: 256, text: ["Canvas only"] }],
      [{ slideId: 256, paragraphs: ["Before ＡＢＣ\t  target After"] }],
    );
    const session = new PptxViewSession(
      root,
      { readBinary: vi.fn(async () => new ArrayBuffer(1)) },
      adapter,
    );
    await session.open("deck.pptx");

    root.querySelector<HTMLButtonElement>(
      '[data-action="open-slide-search"]',
    )!.click();
    const input = root.querySelector<HTMLInputElement>(
      '[data-action="slide-search-input"]',
    )!;
    input.value = "abc target";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    root.querySelector<HTMLButtonElement>(
      '[data-action="slide-search-notes-match"]',
    )!.click();

    expect(root.querySelector(".pptx-viewer__notes-highlight")?.textContent)
      .toBe("ＡＢＣ\t  target");
    expect(root.dataset.notesCollapsed).toBe("false");
    session.dispose();
  });
});
