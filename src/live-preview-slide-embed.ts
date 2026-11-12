import {
  EditorSelection,
  Prec,
  RangeSetBuilder,
  StateEffect,
  StateField,
  type EditorState,
  type Extension,
} from "@codemirror/state";
import { syntaxTree } from "@codemirror/language";
import {
  Decoration,
  EditorView,
  ViewPlugin,
  WidgetType,
  type DecorationSet,
} from "@codemirror/view";
import type { MessageTranslator } from "./i18n";
import { mountOfficeFileEmbed, type OfficeFileEmbedOptions } from "./office-file-embed";
import {
  matchStandaloneOfficeFileEmbedLine,
  type OfficeFileEmbedTarget,
} from "./office-file-embed-target";
import type { PptxRendererAdapter } from "./renderer/pptx-renderer-adapter";
import {
  SlideEmbedController,
  type SlideEmbedSourceFile,
} from "./slide-embed-core";
import {
  matchStandaloneSlideEmbedLine,
  type SlideReferenceTarget,
} from "./slide-reference";
import { SlideEmbedScheduler } from "./slide-embed-scheduler";
import {
  observeSlideEmbedVisibility,
  type SlideEmbedVisibilityDisposer,
} from "./slide-embed-visibility";

export interface LivePreviewSlideEmbedOptions<
  TFile extends SlideEmbedSourceFile = SlideEmbedSourceFile,
> {
  readonly livePreviewField: StateField<boolean>;
  readonly getSourcePath: (state: EditorState) => string;
  readonly resolveFile: (sourcePath: string, notePath: string) => TFile | null;
  readonly readBinary: (file: TFile) => Promise<ArrayBuffer>;
  readonly renderer: PptxRendererAdapter;
  readonly scheduler: SlideEmbedScheduler;
  readonly messages: MessageTranslator;
  readonly showDiagnostics: () => boolean;
  readonly openExternally?: (file: TFile) => Promise<void>;
  readonly openSource?: (linkTarget: string, notePath: string) => void | Promise<void>;
  readonly officeFiles?: OfficeFileEmbedOptions<TFile>;
}

class LivePreviewOfficeFileWidget<TFile> extends WidgetType {
  private cleanup: (() => void) | undefined;

  constructor(
    private readonly target: OfficeFileEmbedTarget,
    private readonly notePath: string,
    private readonly options: OfficeFileEmbedOptions<TFile>,
  ) {
    super();
  }

  override eq(other: LivePreviewOfficeFileWidget<TFile>): boolean {
    return this.target.sourcePath === other.target.sourcePath
      && this.notePath === other.notePath
      && this.options === other.options;
  }

  override toDOM(view: EditorView): HTMLElement {
    const host = view.dom.ownerDocument.createElement("div");
    this.cleanup = mountOfficeFileEmbed(host, this.target, this.notePath, this.options);
    return host;
  }

  override destroy(): void {
    this.cleanup?.();
    this.cleanup = undefined;
  }

  override ignoreEvent(): boolean {
    return true;
  }
}

class LivePreviewSlideEmbedWidget<
  TFile extends SlideEmbedSourceFile = SlideEmbedSourceFile,
> extends WidgetType {
  private controller: SlideEmbedController<TFile> | null = null;
  private stopObserving: SlideEmbedVisibilityDisposer | null = null;

  constructor(
    private readonly from: number,
    private readonly sourcePath: string,
    private readonly target: SlideReferenceTarget,
    private readonly notePath: string,
    private readonly options: LivePreviewSlideEmbedOptions<TFile>,
  ) {
    super();
  }

  override eq(other: LivePreviewSlideEmbedWidget<TFile>): boolean {
    return this.from === other.from
      && this.sourcePath === other.sourcePath
      && this.target.slideId === other.target.slideId
      && this.target.createdSlideNumber === other.target.createdSlideNumber
      && this.notePath === other.notePath;
  }

  override toDOM(view: EditorView): HTMLElement {
    const host = view.dom.createDiv();
    host.remove();
    const file = this.options.resolveFile(this.sourcePath, this.notePath);
    const { options, notePath, sourcePath, target, from } = this;
    const { openSource } = options;
    this.controller = new SlideEmbedController(host, {
      readBinary: options.readBinary,
      renderer: options.renderer,
      scheduler: options.scheduler,
      messages: options.messages,
      showDiagnostics: options.showDiagnostics,
      openExternally: options.openExternally,
      openSource: openSource === undefined
        ? undefined
        : (linkTarget) => openSource(linkTarget, notePath),
    });
    this.controller.mount({ file, sourcePath, target });
    if (file !== null) {
      this.stopObserving = observeSlideEmbedVisibility(host, (visible) => {
        this.controller?.setVisible(visible);
      });
    }
    host.addEventListener("click", (event) => {
      const el = event.target;
      const ElementConstructor = host.ownerDocument.defaultView?.Element;
      if (ElementConstructor === undefined || !(el instanceof ElementConstructor)) {
        return;
      }
      if (el.closest("a.internal-link, [data-action='open-externally']")) return;
      const selection = host.ownerDocument.getSelection();
      if (selection !== null && !selection.isCollapsed
        && host.contains(selection.anchorNode) && host.contains(selection.focusNode)) return;
      event.preventDefault();
      event.stopPropagation();
      view.dispatch({ selection: EditorSelection.cursor(from + 1) });
    });
    return host;
  }

  override destroy(): void {
    this.stopObserving?.();
    this.stopObserving = null;
    this.controller?.dispose();
    this.controller = null;
  }

  override ignoreEvent(): boolean {
    return true;
  }
}

function selectionTouches(state: EditorState, from: number, to: number): boolean {
  return state.selection.ranges.some((range) => range.from <= to && range.to >= from);
}

function overlapsCodeSyntax(state: EditorState, from: number, to: number): boolean {
  let overlaps = false;
  syntaxTree(state).iterate({
    from,
    to,
    enter: ({ name }) => {
      if (name.toLowerCase().includes("code")) {
        overlaps = true;
        return false;
      }
      return true;
    },
  });
  return overlaps;
}

function isLivePreviewEditorSurface(view: EditorView, livePreview: boolean): boolean {
  if (!livePreview) return false;
  const sourceView = view.dom.closest(".markdown-source-view");
  if (sourceView === null) return true;
  const display = sourceView.ownerDocument.defaultView
    ?.getComputedStyle(sourceView).display;
  return sourceView.classList.contains("is-live-preview")
    && display !== "none";
}

function buildDecorations<TFile extends SlideEmbedSourceFile>(
  state: EditorState,
  visibleRanges: readonly { readonly from: number; readonly to: number }[],
  options: LivePreviewSlideEmbedOptions<TFile>,
  surfaceActive: boolean,
): DecorationSet {
  if (!surfaceActive || !state.field(options.livePreviewField, false)) {
    return Decoration.none;
  }
  const notePath = options.getSourcePath(state);
  const builder = new RangeSetBuilder<Decoration>();
  let lastVisitedLine = 0;
  for (const range of visibleRanges) {
    const firstLine = state.doc.lineAt(range.from).number;
    const lastPosition = Math.max(range.from, range.to - 1);
    const lastLine = state.doc.lineAt(lastPosition).number;
    for (
      let lineNumber = Math.max(firstLine, lastVisitedLine + 1);
      lineNumber <= lastLine;
      lineNumber += 1
    ) {
      lastVisitedLine = lineNumber;
      const line = state.doc.line(lineNumber);
      const slideMatch = matchStandaloneSlideEmbedLine(line.text);
      const officeMatch = slideMatch === null && options.officeFiles !== undefined
        ? matchStandaloneOfficeFileEmbedLine(line.text)
        : null;
      const match = slideMatch ?? officeMatch;
      if (match === null) continue;
      const from = line.from + match.fromOffset;
      const to = line.from + match.toOffset;
      if (
        overlapsCodeSyntax(state, from, to)
        || selectionTouches(state, from, to)
      ) continue;
      const blockTo = lineNumber < state.doc.lines ? line.to + 1 : to;
      const widget = slideMatch !== null
        ? new LivePreviewSlideEmbedWidget(
            from, slideMatch.sourcePath, slideMatch.target, notePath, options,
          )
        : new LivePreviewOfficeFileWidget(officeMatch!.target, notePath, options.officeFiles!);
      builder.add(
        line.from,
        blockTo,
        Decoration.replace({
          widget,
          block: true,
        }),
      );
    }
  }
  return builder.finish();
}

export function createLivePreviewSlideEmbedExtension<
  TFile extends SlideEmbedSourceFile = SlideEmbedSourceFile,
>(
  options: LivePreviewSlideEmbedOptions<TFile>,
): Extension {
  const setSurfaceActive = StateEffect.define<boolean>();
  const setVisibleRanges = StateEffect.define<
    readonly { readonly from: number; readonly to: number }[]
  >();
  const surfaceActiveField = StateField.define<boolean>({
    create: () => true,
    update: (value, tr) => {
      for (const effect of tr.effects) {
        if (effect.is(setSurfaceActive)) return effect.value;
      }
      return value;
    },
  });

  const visibleRangesField = StateField.define<
    readonly { readonly from: number; readonly to: number }[]
  >({
    create: () => [],
    update: (ranges, tr) => {
      let next = tr.docChanged
        ? ranges.map(({ from, to }) => ({
            from: tr.changes.mapPos(from, -1),
            to: tr.changes.mapPos(to, 1),
          }))
        : ranges;
      for (const effect of tr.effects) {
        if (effect.is(setVisibleRanges)) next = effect.value;
      }
      return next;
    },
  });

  const decorationsField = StateField.define<DecorationSet>({
    create: () => Decoration.none,
    update: (decorations, tr) => {
      // Obsidian adds/removes plugin editor extensions via Compartment
      // reconfigure. CodeMirror still runs update() after reconfigure, with a
      // startState that may not contain newly added (or already-removed)
      // sibling fields. Required field() throws and aborts plugin load.
      const nextRanges = tr.state.field(visibleRangesField, false);
      const nextLivePreview = tr.state.field(options.livePreviewField, false);
      const nextSurfaceActive = tr.state.field(surfaceActiveField, false);
      if (
        nextRanges === undefined
        || nextLivePreview === undefined
        || nextSurfaceActive === undefined
      ) {
        return decorations.map(tr.changes);
      }
      const rangesChanged =
        tr.startState.field(visibleRangesField, false) !== nextRanges;
      if (
        tr.docChanged
        || tr.selection
        || rangesChanged
        || tr.startState.field(options.livePreviewField, false) !== nextLivePreview
        || tr.startState.field(surfaceActiveField, false) !== nextSurfaceActive
      ) {
        return buildDecorations(
          tr.state,
          nextRanges,
          options,
          nextSurfaceActive,
        );
      }
      return decorations.map(tr.changes);
    },
    provide: (value) => Prec.highest(EditorView.decorations.from(value)),
  });

  const viewportMonitor = ViewPlugin.fromClass(class {
    private destroyed = false;
    private pending = false;
    private readonly resizeObserver: ResizeObserver | null;
    private readonly onScroll = (): void => this.scheduleSync();

    constructor(private readonly view: EditorView) {
      view.scrollDOM.addEventListener("scroll", this.onScroll, { passive: true });
      const Observer = view.dom.ownerDocument.defaultView?.ResizeObserver;
      this.resizeObserver = Observer === undefined
        ? null
        : new Observer(() => this.scheduleSync());
      this.resizeObserver?.observe(view.scrollDOM);
      this.scheduleSync();
    }

    update(update: { readonly docChanged: boolean }): void {
      if (update.docChanged) this.scheduleSync();
    }

    destroy(): void {
      this.destroyed = true;
      this.view.scrollDOM.removeEventListener("scroll", this.onScroll);
      this.resizeObserver?.disconnect();
    }

    private scheduleSync(): void {
      if (this.pending) return;
      this.pending = true;
      queueMicrotask(() => {
        this.pending = false;
        if (this.destroyed) return;
        // Use the viewport's document range rather than visibleRanges. Obsidian's
        // native file-embed decoration hides the source range from visibleRanges;
        // the higher-precedence slide widget still needs to inspect that bounded
        // range so it can replace the native widget.
        const ranges = [{
          from: this.view.viewport.from,
          to: this.view.viewport.to,
        }];
        const previous = this.view.state.field(visibleRangesField, false);
        if (previous === undefined) return;
        if (
          ranges.length === previous.length
          && ranges.every((range, index) =>
            range.from === previous[index]?.from && range.to === previous[index]?.to)
        ) {
          return;
        }
        this.view.dispatch({ effects: setVisibleRanges.of(ranges) });
      });
    }
  });

  const surfaceMonitor = ViewPlugin.fromClass(class {
    private readonly observer: MutationObserver | null;
    private destroyed = false;
    private pending = false;
    constructor(private readonly view: EditorView) {
      const Observer = view.dom.ownerDocument.defaultView?.MutationObserver;
      this.observer = Observer === undefined
        ? null
        : new Observer(() => this.scheduleSync());
      const leaf = view.dom.closest(".workspace-leaf-content");
      if (leaf !== null) {
        this.observer?.observe(leaf, {
          attributes: true,
          attributeFilter: ["data-mode", "class"],
        });
      }
      const source = view.dom.closest(".markdown-source-view");
      if (source !== null) {
        this.observer?.observe(source, {
          attributes: true,
          attributeFilter: ["class", "style"],
        });
      }
      this.scheduleSync();
    }
    update(): void {
      this.scheduleSync();
    }
    destroy(): void {
      this.destroyed = true;
      this.observer?.disconnect();
    }
    private scheduleSync(): void {
      if (this.pending) return;
      this.pending = true;
      queueMicrotask(() => {
        this.pending = false;
        if (!this.destroyed) this.sync();
      });
    }
    private sync(): void {
      const livePreview = this.view.state.field(options.livePreviewField, false);
      const surfaceActive = this.view.state.field(surfaceActiveField, false);
      if (livePreview === undefined || surfaceActive === undefined) return;
      const active = isLivePreviewEditorSurface(this.view, livePreview);
      if (active !== surfaceActive) {
        this.view.dispatch({ effects: setSurfaceActive.of(active) });
      }
    }
  });

  return [
    surfaceActiveField,
    visibleRangesField,
    decorationsField,
    viewportMonitor,
    surfaceMonitor,
  ];
}
