import { afterEach, describe, expect, it, vi } from "vitest";
import { EditorSelection, EditorState, StateField } from "@codemirror/state";
import { EditorView } from "@codemirror/view";
import { ENGLISH_MESSAGE_TRANSLATOR } from "../src/i18n";
import { createLivePreviewSlideEmbedExtension } from "../src/live-preview-slide-embed";
import type { PptxRendererSession } from "../src/renderer/pptx-renderer-adapter";
import { SlideEmbedScheduler } from "../src/slide-embed-scheduler";

const originalIntersectionObserver = globalThis.IntersectionObserver;

afterEach(() => {
  globalThis.IntersectionObserver = originalIntersectionObserver;
  vi.restoreAllMocks();
  document.body.replaceChildren();
});

function makeSession(
  identities: readonly number[] = [256, 261, 300],
): PptxRendererSession {
  return {
    slideCount: identities.length,
    slideIdentities: identities,
    slideWidth: 16,
    slideHeight: 9,
    capabilities: { thumbnails: false, prefetch: false },
    compatibilityWarnings: [],
    renderSlide: vi.fn(async (index: number) => {
      document.querySelector(".pptx-slide-embed__canvas")?.append(
        Object.assign(document.createElement("div"), {
          textContent: `rendered ${index}`,
        }),
      );
    }),
    dispose: vi.fn(),
  };
}

function createHarness(doc: string, options?: {
  readonly livePreview?: boolean;
  readonly cursor?: number;
  readonly openSource?: (linkTarget: string) => void;
  readonly ownerDocument?: Document;
}) {
  // Match Reading View focused tests: without IntersectionObserver the widget
  // mounts work immediately for deterministic unit coverage.
  globalThis.IntersectionObserver = undefined as never;
  const livePreviewField = StateField.define<boolean>({
    create: () => options?.livePreview ?? true,
    update: (value) => value,
  });
  const sourcePathField = StateField.define<string>({
    create: () => "note.md",
    update: (value) => value,
  });
  const file = { basename: "deck", path: "deck.pptx" };
  const session = makeSession();
  const open = vi.fn(async () => session);
  const openSource = options?.openSource ?? vi.fn();
  const ownerDocument = options?.ownerDocument ?? document;
  const parent = ownerDocument.createElement("div");
  ownerDocument.body.append(parent);
  // Default to the end of the document so a leading standalone embed is not
  // immediately suppressed by a cursor that touches its range.
  const cursor = options?.cursor ?? doc.length;
  const view = new EditorView({
    parent,
    state: EditorState.create({
      doc,
      selection: EditorSelection.cursor(cursor),
      extensions: [
        livePreviewField,
        sourcePathField,
        createLivePreviewSlideEmbedExtension({
          livePreviewField,
          getSourcePath: (state) => state.field(sourcePathField),
          resolveFile: (sourcePath) =>
            sourcePath === "deck.pptx" ? file : null,
          readBinary: async () => new ArrayBuffer(8),
          renderer: { open },
          scheduler: new SlideEmbedScheduler(2),
          messages: ENGLISH_MESSAGE_TRANSLATOR,
          showDiagnostics: () => false,
          openSource,
        }),
      ],
    }),
  });
  return { view, session, open, openSource, parent };
}

describe("Live Preview slide embed extension", () => {
  it("renders a standalone canonical embed only while Live Preview is active", async () => {
    const embed = "![[deck.pptx#slide-id=261&slide=1|deck — Slide 1]]";
    const { view, session, open } = createHarness(`intro\n\n${embed}\n\noutro`);

    await vi.waitFor(() => {
      expect(view.dom.querySelector(".pptx-slide-embed")?.getAttribute("data-state"))
        .toBe("ready");
    });
    expect(open).toHaveBeenCalledOnce();
    expect(session.renderSlide).toHaveBeenCalledWith(1);
    expect(view.state.doc.toString()).toContain(embed);

    view.destroy();
    expect(session.dispose).toHaveBeenCalled();
  });

  it("keeps syntax visible when a selection touches the embed range", async () => {
    const embed = "![[deck.pptx#slide-id=261&slide=1|deck — Slide 1]]";
    const prefix = "intro\n\n";
    const from = prefix.length;
    const { view, open } = createHarness(`${prefix}${embed}\n\noutro`, {
      cursor: from + 3,
    });

    await Promise.resolve();
    expect(view.dom.querySelector(".pptx-slide-embed")).toBeNull();
    expect(open).not.toHaveBeenCalled();
    expect(view.state.doc.toString()).toContain(embed);
    view.destroy();
  });

  it("reveals exact syntax when the slide canvas is clicked", async () => {
    const embed = "![[deck.pptx#slide-id=261&slide=1|deck — Slide 1]]";
    const prefix = "intro\n\n";
    const from = prefix.length;
    const { view } = createHarness(`${prefix}${embed}\n\noutro`);

    await vi.waitFor(() => {
      expect(view.dom.querySelector(".pptx-slide-embed")?.getAttribute("data-state"))
        .toBe("ready");
    });
    view.dom.querySelector<HTMLElement>(".pptx-slide-embed__canvas")!.click();

    await vi.waitFor(() => {
      expect(view.dom.querySelector(".pptx-slide-embed")).toBeNull();
    });
    expect(view.state.selection.main.from).toBeGreaterThanOrEqual(from);
    expect(view.state.selection.main.from).toBeLessThanOrEqual(from + embed.length);
    expect(view.state.doc.toString()).toContain(embed);
    view.destroy();
  });

  it("creates and handles widgets in the editor's pop-out window realm", async () => {
    const frame = document.createElement("iframe");
    document.body.append(frame);
    const popoutDocument = frame.contentDocument!;
    const popoutElementPrototype = (frame.contentWindow as unknown as {
      HTMLElement: typeof HTMLElement;
    }).HTMLElement.prototype;
    for (const helper of ["createEl", "createDiv", "createSpan"] as const) {
      Object.defineProperty(popoutElementPrototype, helper, {
        configurable: true,
        value: HTMLElement.prototype[helper],
      });
    }
    const embed = "![[deck.pptx#slide-id=261&slide=1|deck — Slide 1]]";
    const { view } = createHarness(`intro\n\n${embed}\n\noutro`, {
      ownerDocument: popoutDocument,
    });

    await vi.waitFor(() => {
      expect(view.dom.querySelector(".pptx-slide-embed")?.getAttribute("data-state"))
        .toBe("ready");
    });
    const host = view.dom.querySelector<HTMLElement>(".pptx-slide-embed")!;
    expect(host.ownerDocument).toBe(popoutDocument);
    host.querySelector<HTMLElement>(".pptx-slide-embed__canvas")!.click();
    await vi.waitFor(() => {
      expect(view.dom.querySelector(".pptx-slide-embed")).toBeNull();
    });
    view.destroy();
    frame.remove();
  });

  it("opens the exact PPTX target from the explicit source action", async () => {
    const embed = "![[deck.pptx#slide-id=261&slide=1|deck — Slide 1]]";
    const openSource = vi.fn();
    const { view } = createHarness(`\n${embed}\n`, {
      openSource,
      cursor: 0,
    });

    await vi.waitFor(() => {
      expect(view.dom.querySelector(".pptx-slide-embed")?.getAttribute("data-state"))
        .toBe("ready");
    });
    const openInTab = view.dom.querySelector<HTMLAnchorElement>('[data-action="open-in-tab"]')!;
    expect(openInTab.getAttribute("aria-label")).toBe("Open in new tab");
    openInTab.click();
    expect(openSource).toHaveBeenCalledWith(
      "deck.pptx#slide-id=261&slide=1",
      "note.md",
    );
    view.destroy();
  });

  it("leaves non-standalone and source-mode documents untouched", async () => {
    const mixed = "see ![[deck.pptx#slide-id=261&slide=1]]";
    const { view: mixedView, open: mixedOpen } = createHarness(mixed);
    await Promise.resolve();
    expect(mixedView.dom.querySelector(".pptx-slide-embed")).toBeNull();
    expect(mixedOpen).not.toHaveBeenCalled();
    mixedView.destroy();

    const embed = "![[deck.pptx#slide-id=261&slide=1]]";
    const { view, open } = createHarness(`\n${embed}\n`, {
      livePreview: false,
      cursor: 0,
    });
    await Promise.resolve();
    expect(view.dom.querySelector(".pptx-slide-embed")).toBeNull();
    expect(open).not.toHaveBeenCalled();
    view.destroy();
  });

  it("ignores stale async completion after the editor is destroyed", async () => {
    let finishOpen!: (session: PptxRendererSession) => void;
    const openGate = new Promise<PptxRendererSession>((resolve) => {
      finishOpen = resolve;
    });
    const livePreviewField = StateField.define<boolean>({
      create: () => true,
      update: (value) => value,
    });
    const sourcePathField = StateField.define<string>({
      create: () => "note.md",
      update: (value) => value,
    });
    const session = makeSession();
    const open = vi.fn(async () => openGate);
    const parent = document.createElement("div");
    document.body.append(parent);
    globalThis.IntersectionObserver = undefined as never;
    const view = new EditorView({
      parent,
      state: EditorState.create({
        doc: "\n![[deck.pptx#slide-id=261&slide=1]]\n",
        selection: EditorSelection.cursor(0),
        extensions: [
          // cursor stays on the blank first line so the embed can mount
          livePreviewField,
          sourcePathField,
          createLivePreviewSlideEmbedExtension({
            livePreviewField,
            getSourcePath: (state) => state.field(sourcePathField),
            resolveFile: () => ({ basename: "deck", path: "deck.pptx" }),
            readBinary: async () => new ArrayBuffer(8),
            renderer: { open },
            scheduler: new SlideEmbedScheduler(2),
            messages: ENGLISH_MESSAGE_TRANSLATOR,
            showDiagnostics: () => false,
          }),
        ],
      }),
    });

    await vi.waitFor(() => expect(open).toHaveBeenCalledOnce());
    view.destroy();
    finishOpen(session);
    await Promise.resolve();
    await Promise.resolve();
    expect(session.renderSlide).not.toHaveBeenCalled();
    expect(session.dispose).toHaveBeenCalledOnce();
  });
});
