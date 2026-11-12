import { registerOfficeEmbedTextCopy } from "./office-embed-text-copy";
import type {
  PptxCompatibilityWarningCategory,
  PptxRendererAdapter,
  PptxRendererSession,
} from "./renderer/pptx-renderer-adapter";
import {
  ENGLISH_MESSAGE_TRANSLATOR,
  type MessageKey,
  type MessageTranslator,
} from "./i18n";
import {
  createDiagnosticSummary,
  type DiagnosticEnvironment,
} from "./diagnostic-summary";
import {
  createOfficeViewerErrorSurface,
  createOfficeViewerToolbar,
  decorateOfficeViewerIconButton,
} from "./office-viewer-chrome";
import {
  PptxOpenError,
  type PptxOpenErrorCategory,
} from "./pptx-open-error";
import { PptxViewerController } from "./pptx-viewer-controller";
import { RenderTaskQueue } from "./render-task-queue";
import { ThumbnailRail } from "./thumbnail-rail";
import {
  DEFAULT_THUMBNAIL_RAIL_WIDTH,
  resolveThumbnailRailWidth,
  thumbnailPreviewWidth,
} from "./thumbnail-rail-sizing";
import { ThumbnailRailResizer } from "./thumbnail-rail-resizer";
import type { SlideReferenceTarget } from "./slide-reference";
import {
  normalizedDisplayMatchRange,
  type SlideSearchSnippet,
} from "./slide-content-search";
import { SlideSearchRail } from "./slide-search-rail";
import {
  hasUsableSpeakerNoteContent,
  indexSpeakerNoteContent,
} from "./speaker-note-content";

export interface VaultBinaryReader<FileRef> {
  readBinary(file: FileRef, signal: AbortSignal): Promise<ArrayBuffer>;
}

export interface PptxViewSessionDiagnostics {
  readonly generation: number;
  readonly openPending: boolean;
  readonly rendererActive: boolean;
  readonly disposed: boolean;
  readonly lifecyclePhase: PptxLifecyclePhase;
  readonly backgroundPending: number;
  readonly backgroundRunning: number;
  readonly mountedThumbnails: number;
  readonly readyThumbnails: number;
}

export type PptxLifecyclePhase =
  | "idle"
  | "reading"
  | "adapter-opening"
  | "first-slide-rendering"
  | "ready"
  | "degraded"
  | "error"
  | "disposed";

export interface PptxViewOptions<FileRef> {
  messages?: MessageTranslator;
  openExternally?: (file: FileRef) => Promise<void>;
  positions?: {
    initialSlideFor(file: FileRef, slideCount: number): number;
    record(file: FileRef, slideIndex: number): void;
  };
  thumbnailRail?: {
    initialWidth(): number;
    recordWidth(width: number): void;
    subscribeWidth?(listener: (width: number) => void): () => void;
  };
  fullscreen?: PptxFullscreenActions;
  diagnostics?: {
    readonly environment: DiagnosticEnvironment;
    rememberReadingPosition(): boolean;
    enabled(): boolean;
    copy(summary: string): Promise<void>;
  };
  slideReferences?: {
    copy(
      file: FileRef,
      target: SlideReferenceTarget,
      embed: boolean,
    ): Promise<void>;
    copyNotes?(
      file: FileRef,
      target: SlideReferenceTarget,
      paragraphs: readonly string[],
    ): Promise<void>;
    copySelection?(
      file: FileRef,
      target: SlideReferenceTarget,
      selectedText: string,
    ): Promise<void>;
  };
  supportsPptxActions?(file: FileRef): boolean;
  /**
   * File-level Presentation companion note action. Independent of Vault APIs
   * inside the session; the host supplies create/open behavior.
   */
  companionNote?: {
    open(file: FileRef): Promise<string>;
  };
}

export interface PptxFullscreenActions {
  isActive(root: HTMLElement): boolean;
  enter(root: HTMLElement): Promise<void>;
  exit(): Promise<void>;
  subscribe(listener: () => void): () => void;
}

function createDefaultFullscreenActions(): PptxFullscreenActions {
  return {
    isActive: (root) => document.fullscreenElement === root,
    enter: async (root) => root.requestFullscreen(),
    exit: async () => document.exitFullscreen(),
    subscribe: (listener) => {
      document.addEventListener("fullscreenchange", listener);
      return () => document.removeEventListener("fullscreenchange", listener);
    },
  };
}

function isEditableTarget(target: EventTarget | null): boolean {
  return target instanceof HTMLElement &&
    (target.isContentEditable ||
      ["INPUT", "BUTTON", "SELECT", "TEXTAREA"].includes(target.tagName));
}

interface SlideTextSelection {
  readonly text: string;
  readonly range: Range;
}

function readSlideTextSelection(
  slideContainer: HTMLElement,
): SlideTextSelection | null {
  const selection = slideContainer.ownerDocument.getSelection();
  if (selection === null || selection.isCollapsed || selection.rangeCount !== 1) {
    return null;
  }
  const range = selection.getRangeAt(0);
  if (
    !slideContainer.contains(range.startContainer) ||
    !slideContainer.contains(range.endContainer)
  ) {
    return null;
  }
  const text = selection.toString();
  return text.trim().length === 0 ? null : { text, range };
}

const errorMessageKeys: Record<PptxOpenErrorCategory, MessageKey> = {
  "unsupported-legacy": "error.unsupportedLegacy",
  malformed: "error.malformed",
  protected: "error.protected",
  incompatible: "error.incompatible",
  "resource-exhausted": "error.resourceExhausted",
  cancelled: "error.cancelled",
  unknown: "error.unknown",
};

const warningMessageKeys: Record<PptxCompatibilityWarningCategory, MessageKey> = {
  "unsupported-content": "compatibility.unsupportedContent",
  "font-substitution": "compatibility.fontSubstitution",
};

export class PptxViewSession<FileRef> {
  private abortController: AbortController | null = null;
  private rendererSession: PptxRendererSession | null = null;
  private viewerController: PptxViewerController | null = null;
  private backgroundQueue: RenderTaskQueue | null = null;
  private thumbnailRail: ThumbnailRail | null = null;
  private readonly runCleanups = new Set<() => void>();
  private generation = 0;
  private openPending = false;
  private disposed = false;
  private lifecyclePhase: PptxLifecyclePhase = "idle";
  private readonly messages: MessageTranslator;
  private diagnosticSourceBytes: number | null = null;
  private diagnosticSlideCount: number | null = null;
  private diagnosticWarnings: readonly PptxCompatibilityWarningCategory[] = [];
  private diagnosticError: PptxOpenErrorCategory | null = null;
  private diagnosticThumbnails = false;
  private diagnosticPrefetch = false;
  private referenceNavigator: ((target: SlideReferenceTarget) => void) | null = null;
  private openSlideContentSearchAction: (() => void) | null = null;
  private toolbar: ReturnType<typeof createOfficeViewerToolbar> | null = null;
  private contentHost: HTMLElement | null = null;

  constructor(
    private readonly root: HTMLElement,
    private readonly reader: VaultBinaryReader<FileRef>,
    private readonly renderer: PptxRendererAdapter,
    private readonly options: PptxViewOptions<FileRef> = {},
  ) {
    this.messages = options.messages ?? ENGLISH_MESSAGE_TRANSLATOR;
    root.classList.add("pptx-viewer");
    root.replaceChildren();
    root.createDiv({
      cls: "pptx-viewer__empty",
      text: this.messages.text("viewer.empty"),
    });
    root.dataset.state = "empty";
    this.setLifecyclePhase("idle");
  }

  private ensurePersistentShell(): {
    toolbar: ReturnType<typeof createOfficeViewerToolbar>;
    content: HTMLElement;
  } {
    if (
      this.toolbar !== null &&
      this.contentHost !== null &&
      this.root.contains(this.toolbar.root) &&
      this.root.contains(this.contentHost)
    ) {
      return { toolbar: this.toolbar, content: this.contentHost };
    }
    this.root.replaceChildren();
    this.toolbar = createOfficeViewerToolbar({
      format: "PPTX",
      extraClassName: "pptx-viewer__toolbar",
    });
    this.contentHost = createDiv();
    this.contentHost.className = "pptx-viewer__content";
    this.root.append(this.toolbar.root, this.contentHost);
    return { toolbar: this.toolbar, content: this.contentHost };
  }

  private clearShellActionSlots(
    toolbar: ReturnType<typeof createOfficeViewerToolbar>,
  ): void {
    toolbar.primary.replaceChildren();
    toolbar.secondary.replaceChildren();
  }

  private disableShellSearchActions(
    toolbar: ReturnType<typeof createOfficeViewerToolbar>,
  ): void {
    for (const button of toolbar.primary.querySelectorAll("button")) {
      button.disabled = true;
    }
  }

  openSlideContentSearch(): boolean {
    const openSearch = this.openSlideContentSearchAction;
    if (openSearch === null) return false;
    openSearch();
    return true;
  }

  async open(
    file: FileRef,
    referenceTarget?: SlideReferenceTarget,
  ): Promise<void> {
    this.teardownOpenResources();
    this.clearTimings();
    this.resetDiagnosticState();
    const generation = ++this.generation;
    this.openPending = true;
    this.disposed = false;
    const controller = new AbortController();
    this.abortController = controller;

    this.root.tabIndex = 0;
    this.root.dataset.thumbnailsCollapsed = "false";
    this.root.dataset.notesCollapsed = "true";
    this.root.dataset.fullscreen = "false";
    this.root.dataset.mountedThumbnailCount = "0";
    this.root.dataset.readyThumbnailCount = "0";
    delete this.root.dataset.referenceSlideId;
    delete this.root.dataset.referenceCreatedSlide;
    delete this.root.dataset.referenceCurrentSlide;

    const { toolbar, content } = this.ensurePersistentShell();
    this.clearShellActionSlots(toolbar);
    content.replaceChildren();

    const header = content.createDiv({
      cls: "pptx-viewer__header pptx-viewer__status",
    });
    const status = header.createDiv({
      cls: "pptx-viewer__status-text",
      text: this.messages.text("viewer.loading"),
      attr: { role: "status", "aria-live": "polite" },
    });
    const headerActions = toolbar.secondary;

    const compatibility = content.createDiv({
      cls: "pptx-viewer__compatibility",
      attr: { role: "note" },
    });
    const referenceNotice = content.createDiv({
      cls: "pptx-viewer__reference-notice",
      attr: { role: "note" },
    });

    const controls = content.createDiv({
      cls: "pptx-viewer__controls",
    });

    const navGroup = controls.createDiv({
      cls: "pptx-viewer__control-group",
    });

    const previousButton = navGroup.createEl("button", {
      type: "button",
      attr: { "data-action": "previous-slide" },
    });
    previousButton.title = this.messages.text("navigation.previous");
    previousButton.setAttribute(
      "aria-label",
      this.messages.text("navigation.previous"),
    );
    decorateOfficeViewerIconButton(previousButton, "lucide-chevron-left");
    previousButton.disabled = true;

    const pageCounter = navGroup.createDiv({
      cls: "pptx-viewer__page-counter",
      attr: { role: "status", "aria-live": "polite" },
    });

    const nextButton = navGroup.createEl("button", {
      type: "button",
      attr: { "data-action": "next-slide" },
    });
    nextButton.title = this.messages.text("navigation.next");
    nextButton.setAttribute(
      "aria-label",
      this.messages.text("navigation.next"),
    );
    decorateOfficeViewerIconButton(nextButton, "lucide-chevron-right");
    nextButton.disabled = true;

    const jumpForm = controls.createEl("form", {
      cls: "pptx-viewer__page-jump pptx-viewer__control-group",
    });
    jumpForm.createSpan({
      text: this.messages.text("navigation.slide"),
    });
    const pageInput = jumpForm.createEl("input", {
      type: "number",
      attr: {
        min: "1",
        step: "1",
        "data-action": "page-number",
        "aria-label": this.messages.text("navigation.slideNumber"),
      },
    });
    pageInput.value = "1";
    pageInput.disabled = true;

    const jumpButton = jumpForm.createEl("button", {
      type: "button",
      text: this.messages.text("navigation.go"),
      attr: { "data-action": "jump-to-slide" },
    });
    jumpButton.disabled = true;

    const viewGroup = controls.createDiv({
      cls: "pptx-viewer__control-group",
    });

    const toggleFullscreen = viewGroup.createEl("button", {
      type: "button",
      text: this.messages.text("fullscreen.button"),
      attr: {
        "data-action": "toggle-fullscreen",
        "aria-label": this.messages.text("fullscreen.enterLabel"),
      },
    });

    const toggleThumbnails = viewGroup.createEl("button", {
      type: "button",
      text: this.messages.text("thumbnails.toggle"),
      attr: {
        "data-action": "toggle-thumbnails",
        "aria-label": this.messages.text("thumbnails.toggleLabel"),
        "aria-expanded": "true",
      },
    });

    const notesPanelId = `pptx-viewer-speaker-notes-${this.generation}`;
    const toggleNotes = viewGroup.createEl("button", {
      type: "button",
      text: this.messages.text("notes.toggle"),
      attr: {
        "data-action": "toggle-notes",
        "aria-label": this.messages.text("notes.toggleLabel"),
        "aria-expanded": "false",
        "aria-controls": notesPanelId,
      },
    });

    let copyTarget: SlideReferenceTarget | null = null;
    const slideReferences = this.options.supportsPptxActions?.(file) === false
      ? undefined
      : this.options.slideReferences;
    const copyReferenceButton = slideReferences === undefined
      ? null
      : headerActions.createEl("button", {
          type: "button",
          title: this.messages.text("reference.copy"),
          attr: {
            "data-action": "copy-slide-reference",
            "aria-label": this.messages.text("reference.copy"),
          },
        });
    if (copyReferenceButton) {
      decorateOfficeViewerIconButton(copyReferenceButton, "lucide-link");
    }
    const copyEmbedButton = slideReferences === undefined
      ? null
      : headerActions.createEl("button", {
          type: "button",
          title: this.messages.text("reference.copyEmbed"),
          attr: {
            "data-action": "copy-slide-embed",
            "aria-label": this.messages.text("reference.copyEmbed"),
          },
        });
    if (copyEmbedButton) {
      decorateOfficeViewerIconButton(copyEmbedButton, "lucide-panel-top");
    }
    const copyNotesButton =
      slideReferences?.copyNotes === undefined
        ? null
        : headerActions.createEl("button", {
            type: "button",
            title: this.messages.text("notes.copy"),
            attr: {
              "data-action": "copy-speaker-notes",
              "aria-label": this.messages.text("notes.copy"),
            },
          });
    if (copyNotesButton) {
      decorateOfficeViewerIconButton(copyNotesButton, "lucide-notebook-text");
    }
    const copySelectionButton =
      slideReferences?.copySelection === undefined
        ? null
        : headerActions.createEl("button", {
            type: "button",
            title: this.messages.text("selection.copyWithSource"),
            attr: {
              "data-action": "copy-selected-slide-text",
              "aria-label": this.messages.text("selection.copyWithSource"),
            },
          });
    if (copySelectionButton) {
      decorateOfficeViewerIconButton(copySelectionButton, "lucide-copy");
    }
    if (copyReferenceButton) copyReferenceButton.disabled = true;
    if (copyEmbedButton) copyEmbedButton.disabled = true;
    if (copyNotesButton) copyNotesButton.disabled = true;
    if (copySelectionButton) copySelectionButton.disabled = true;

    const actionStatus = content.createDiv({
      cls: "pptx-viewer__action-status",
      attr: { role: "status", "aria-live": "polite" },
    });

    const companionButton = this.createCompanionNoteButton(
      file,
      generation,
      actionStatus,
      () => this.options.fullscreen ?? createDefaultFullscreenActions(),
    );
    if (companionButton) headerActions.prepend(companionButton);

    const openExternally = this.createExternalOpenButton(
      file,
      generation,
      actionStatus,
    );
    if (openExternally) headerActions.append(openExternally);
    const diagnosticButton = this.createDiagnosticButton(actionStatus);
    if (diagnosticButton) {
      decorateOfficeViewerIconButton(diagnosticButton, "lucide-clipboard-copy");
      diagnosticButton.setAttribute(
        "aria-label",
        this.messages.text("diagnostics.copy"),
      );
      diagnosticButton.title = this.messages.text("diagnostics.copy");
      headerActions.append(diagnosticButton);
    }

    const readingBody = content.createDiv({
      cls: "pptx-viewer__reading-body",
    });
    const thumbnailRoot = readingBody.createDiv();
    const slideStage = readingBody.createDiv({
      cls: "pptx-viewer__slide-stage",
    });
    const slideContainer = slideStage.createDiv({
      cls: "pptx-viewer__slide",
    });
    const floatingCopySelectionButton =
      slideReferences?.copySelection === undefined
        ? null
        : slideStage.createEl("button", {
            type: "button",
            cls: "pptx-viewer__selection-copy",
            text: this.messages.text("selection.copyWithSource"),
            attr: {
              "data-action": "copy-selected-slide-text-floating",
              "aria-label": this.messages.text("selection.copyWithSource"),
            },
          });
    if (floatingCopySelectionButton) floatingCopySelectionButton.hidden = true;
    const notesPanel = slideStage.createDiv({
      cls: "pptx-viewer__notes-panel",
      attr: {
        id: notesPanelId,
        role: "region",
        "aria-label": this.messages.text("notes.panelLabel"),
        "aria-hidden": "true",
      },
    });
    const notesContent = notesPanel.createDiv({
      cls: "pptx-viewer__notes-content",
      attr: {
        role: "status",
        "aria-live": "polite",
      },
    });
    const setNotesCollapsed = (collapsed: boolean) => {
      this.root.dataset.notesCollapsed = String(collapsed);
      toggleNotes.setAttribute("aria-expanded", String(!collapsed));
      notesPanel.setAttribute("aria-hidden", String(collapsed));
    };
    this.root.dataset.state = "loading";
    this.setLifecyclePhase("reading");
    delete this.root.dataset.errorCategory;
    delete this.root.dataset.warningCategories;
    let phase: "read" | "renderer" = "read";

    try {
      const openedAt = performance.now();
      const buffer = await this.reader.readBinary(file, controller.signal);
      this.diagnosticSourceBytes = buffer.byteLength;
      controller.signal.throwIfAborted();
      this.setLifecyclePhase("adapter-opening");
      phase = "renderer";
      const rendererSession = await this.renderer.open(
        buffer,
        slideContainer,
        controller.signal,
      );
      if (generation !== this.generation || controller.signal.aborted) {
        rendererSession.dispose();
        return;
      }
      this.rendererSession = rendererSession;
      this.runCleanups.add(registerOfficeEmbedTextCopy(slideContainer));
      this.diagnosticSlideCount = rendererSession.slideCount;
      this.diagnosticThumbnails = rendererSession.capabilities.thumbnails;
      this.diagnosticPrefetch = rendererSession.capabilities.prefetch;
      this.root.dataset.metadataMs = (performance.now() - openedAt).toFixed(3);
      this.setLifecyclePhase("first-slide-rendering");
      const isCurrentRun = () =>
        generation === this.generation &&
        !controller.signal.aborted &&
        this.rendererSession === rendererSession;
      let selectedSlideText: string | null = null;
      const clearSlideTextSelectionAction = () => {
        selectedSlideText = null;
        if (copySelectionButton) copySelectionButton.disabled = true;
        if (floatingCopySelectionButton) floatingCopySelectionButton.hidden = true;
      };
      const updateSlideTextSelectionAction = () => {
        if (!isCurrentRun()) {
          clearSlideTextSelectionAction();
          return;
        }
        const selected = readSlideTextSelection(slideContainer);
        selectedSlideText = selected?.text ?? null;
        const available = selected !== null && copyTarget !== null;
        if (copySelectionButton) copySelectionButton.disabled = !available;
        if (floatingCopySelectionButton === null) return;
        floatingCopySelectionButton.hidden = !available;
        if (!available || selected === null) return;
        if (typeof selected.range.getBoundingClientRect !== "function") {
          floatingCopySelectionButton.hidden = true;
          return;
        }
        try {
          const selectionRect = selected.range.getBoundingClientRect();
          const stageRect = slideStage.getBoundingClientRect();
          const maximumLeft = Math.max(8, stageRect.width - 8);
          const left = Math.min(
            maximumLeft,
            Math.max(
              8,
              selectionRect.left - stageRect.left + selectionRect.width / 2,
            ),
          );
          const top = Math.max(8, selectionRect.bottom - stageRect.top + 8);
          floatingCopySelectionButton.style.left = `${left}px`;
          floatingCopySelectionButton.style.top = `${top}px`;
        } catch {
          floatingCopySelectionButton.hidden = true;
        }
      };
      let notesHighlight: {
        readonly slideId: number;
        readonly snippet: SlideSearchSnippet;
      } | null = null;
      const notesBySlideId = indexSpeakerNoteContent(
        rendererSession.slideIdentities ?? [],
        rendererSession.speakerNoteContent,
      );
      let currentNoteParagraphs: readonly string[] = [];
      const updateCopyNotesAvailability = () => {
        if (!copyNotesButton) return;
        copyNotesButton.disabled =
          copyTarget === null || currentNoteParagraphs.length === 0;
      };
      const renderHighlightedParagraph = (
        container: HTMLElement,
        paragraph: string,
        highlight: SlideSearchSnippet | null,
      ) => {
        const node = container.createEl("p", {
          cls: "pptx-viewer__notes-paragraph",
        });
        if (highlight === null) {
          node.textContent = paragraph;
          return;
        }
        const matchRange = normalizedDisplayMatchRange(
          paragraph,
          highlight.match,
        );
        if (matchRange === null) {
          node.textContent = paragraph;
          return;
        }
        if (matchRange.start > 0) {
          node.append(
            document.createTextNode(paragraph.slice(0, matchRange.start)),
          );
        }
        node.createEl("mark", {
          cls: "pptx-viewer__notes-highlight",
          text: paragraph.slice(matchRange.start, matchRange.end),
        });
        if (matchRange.end < paragraph.length) {
          node.append(document.createTextNode(paragraph.slice(matchRange.end)));
        }
      };
      const renderSpeakerNotes = (slideIndex: number) => {
        notesContent.replaceChildren();
        if (notesBySlideId === undefined) {
          currentNoteParagraphs = [];
          notesContent.createDiv({
            cls: "pptx-viewer__notes-message",
            text: this.messages.text("notes.unavailable"),
            attr: { "data-notes-state": "unavailable" },
          });
          updateCopyNotesAvailability();
          return;
        }
        const slideId = rendererSession.slideIdentities?.[slideIndex];
        const paragraphs = slideId === undefined
          ? []
          : notesBySlideId.get(slideId) ?? [];
        currentNoteParagraphs = paragraphs;
        if (paragraphs.length === 0) {
          notesContent.createDiv({
            cls: "pptx-viewer__notes-message",
            text: this.messages.text("notes.empty"),
            attr: { "data-notes-state": "empty" },
          });
          updateCopyNotesAvailability();
          return;
        }
        const list = notesContent.createDiv({
          cls: "pptx-viewer__notes-paragraphs",
          attr: { "data-notes-state": "ready" },
        });
        const highlightForSlide =
          notesHighlight !== null &&
          slideId !== undefined &&
          notesHighlight.slideId === slideId
            ? notesHighlight.snippet
            : null;
        for (const paragraph of paragraphs) {
          renderHighlightedParagraph(list, paragraph, highlightForSlide);
        }
        updateCopyNotesAvailability();
      };
      const queue = new RenderTaskQueue();
      this.backgroundQueue = queue;
      let initialCommitPending = true;
      let initialRenderFailed = false;
      let navigationStartedAt: number | undefined;
      let rail: ThumbnailRail | null = null;
      let searchRail: SlideSearchRail | null = null;
      let openSearch = () => {};
      let closeSearch = () => {};
      let thumbnailScrollTopBeforeSearch = 0;
      let viewController!: PptxViewerController;
      const restoreControlState = () => {
        const currentSlideIndex = viewController.state.currentSlideIndex;
        previousButton.disabled = currentSlideIndex === 0;
        nextButton.disabled =
          currentSlideIndex >= rendererSession.slideCount - 1;
        pageInput.disabled = false;
        jumpButton.disabled = false;
        if (copyReferenceButton) copyReferenceButton.disabled = copyTarget === null;
        if (copyEmbedButton) copyEmbedButton.disabled = copyTarget === null;
        updateSlideTextSelectionAction();
        updateCopyNotesAvailability();
      };
      const updateMountedCount = () => {
        this.root.dataset.mountedThumbnailCount = String(rail?.mountedCount ?? 0);
      };
      const navigate = (targetIndex: number): void => {
        if (!isCurrentRun()) return;
        const current = viewController.state.currentSlideIndex;
        if (
          !Number.isInteger(targetIndex) ||
          targetIndex < 0 ||
          targetIndex >= rendererSession.slideCount ||
          targetIndex === current
        ) return;
        navigationStartedAt = performance.now();
        void viewController.navigate(targetIndex);
      };
      let initialSlideIndex = 0;
      let referenceNoticeSlideIndex: number | null = null;
      let referenceNoticeText = "";
      try {
        initialSlideIndex = this.options.positions?.initialSlideFor(
          file,
          rendererSession.slideCount,
        ) ?? 0;
      } catch {
        // Persistence is optional; private storage failures must not block reading.
      }
      if (referenceTarget !== undefined) {
        const referenceIndex = rendererSession.slideIdentities?.indexOf(
          referenceTarget.slideId,
        ) ?? -1;
        if (referenceIndex < 0) {
          this.teardownOpenResources();
          this.showMissingReference(file, generation);
          return;
        }
        initialSlideIndex = referenceIndex;
        referenceNoticeSlideIndex = referenceIndex;
        this.root.dataset.referenceSlideId = String(referenceTarget.slideId);
        this.root.dataset.referenceCreatedSlide = String(
          referenceTarget.createdSlideNumber,
        );
        this.root.dataset.referenceCurrentSlide = String(referenceIndex + 1);
        if (referenceTarget.createdSlideNumber !== referenceIndex + 1) {
          referenceNoticeText = this.messages.text("reference.moved", {
            created: referenceTarget.createdSlideNumber,
            current: referenceIndex + 1,
          });
          referenceNotice.textContent = referenceNoticeText;
        }
      }
      viewController = new PptxViewerController(rendererSession, queue, {
        setNavigationPending: (pending) => {
          if (!isCurrentRun()) return;
          if (pending) clearSlideTextSelectionAction();
          previousButton.disabled = pending;
          nextButton.disabled = pending;
          pageInput.disabled = pending;
          jumpButton.disabled = pending;
          if (copyReferenceButton) copyReferenceButton.disabled = true;
          if (copyEmbedButton) copyEmbedButton.disabled = true;
          if (copyNotesButton) copyNotesButton.disabled = true;
          if (copySelectionButton) copySelectionButton.disabled = true;
          if (!pending) restoreControlState();
        },
        commitSlide: (index) => {
          if (!isCurrentRun()) return;
          clearSlideTextSelectionAction();
          const isInitialCommit = initialCommitPending;
          if (isInitialCommit) {
            initialCommitPending = false;
            this.root.dataset.firstReadableMs = (
              performance.now() - openedAt
            ).toFixed(3);
          } else {
            const startedAt = navigationStartedAt ?? performance.now();
            this.root.dataset.lastSlideSwitchMs = (
              performance.now() - startedAt
            ).toFixed(3);
          }
          navigationStartedAt = undefined;
          pageInput.value = String(index + 1);
          pageCounter.textContent = this.messages.text("page.counter", {
            current: index + 1,
            total: rendererSession.slideCount,
          });
          referenceNotice.textContent = index === referenceNoticeSlideIndex
            ? referenceNoticeText
            : "";
          rail?.setCurrentSlide(index);
          if (searchRail?.isOpen) {
            thumbnailRoot.scrollTop = thumbnailScrollTopBeforeSearch;
          }
          searchRail?.setCurrentSlide(index);
          const slideId = rendererSession.slideIdentities?.[index];
          copyTarget = slideId === undefined
            ? null
            : { slideId, createdSlideNumber: index + 1 };
          renderSpeakerNotes(index);
          updateSlideTextSelectionAction();
          updateMountedCount();
          this.root.dataset.state = "ready";
          this.setLifecyclePhase("ready");
          status.textContent = "";
          if (!isInitialCommit) {
            try {
              this.options.positions?.record(file, index);
            } catch {
              // The readable page remains authoritative if persistence fails.
            }
          }
        },
        reportNavigationFailure: (index) => {
          if (!isCurrentRun()) return;
          if (initialCommitPending) initialRenderFailed = true;
          copyTarget = null;
          currentNoteParagraphs = [];
          notesHighlight = null;
          if (copyReferenceButton) copyReferenceButton.disabled = true;
          if (copyEmbedButton) copyEmbedButton.disabled = true;
          if (copyNotesButton) copyNotesButton.disabled = true;
          clearSlideTextSelectionAction();
          navigationStartedAt = undefined;
          pageInput.value = String(viewController.state.currentSlideIndex + 1);
          this.root.dataset.state = "degraded";
          this.setLifecyclePhase("degraded");
          status.textContent = this.messages.text("navigation.renderFailure", {
            slide: index + 1,
          });
        },
      }, { initialSlideIndex });
      this.viewerController = viewController;
      this.referenceNavigator = (target) => {
        if (!isCurrentRun()) return;
        const referenceIndex = rendererSession.slideIdentities?.indexOf(
          target.slideId,
        ) ?? -1;
        if (referenceIndex < 0) {
          this.teardownOpenResources();
          this.showMissingReference(file, generation);
          return;
        }
        referenceNoticeSlideIndex = referenceIndex;
        referenceNoticeText = target.createdSlideNumber === referenceIndex + 1
          ? ""
          : this.messages.text("reference.moved", {
              created: target.createdSlideNumber,
              current: referenceIndex + 1,
            });
        this.root.dataset.referenceSlideId = String(target.slideId);
        this.root.dataset.referenceCreatedSlide = String(
          target.createdSlideNumber,
        );
        this.root.dataset.referenceCurrentSlide = String(referenceIndex + 1);
        referenceNotice.textContent =
          viewController.state.currentSlideIndex === referenceIndex
            ? referenceNoticeText
            : "";
        navigate(referenceIndex);
      };
      await viewController.start();
      controller.signal.throwIfAborted();
      if (!isCurrentRun()) return;
      if (initialRenderFailed) {
        throw new PptxOpenError("incompatible", "Initial slide render failed");
      }
      this.showCompatibilityWarnings(
        compatibility,
        rendererSession.compatibilityWarnings ?? [],
      );
      if (rendererSession.detectCompatibilityWarnings) {
        const detectWarnings = () => {
          if (!isCurrentRun()) return;
          try {
            this.showCompatibilityWarnings(
              compatibility,
              rendererSession.detectCompatibilityWarnings!(),
            );
          } catch {
            // Optional compatibility inspection must not disrupt reading.
          }
        };
        let removeDeferredFocus = () => {};
        const warningTimer = window.setTimeout(() => {
          if (this.root.closest(".workspace-leaf.mod-active")) {
            detectWarnings();
            return;
          }
          const onFocus = () => {
            removeDeferredFocus();
            detectWarnings();
          };
          this.root.addEventListener("focusin", onFocus, { once: true });
          removeDeferredFocus = () =>
            this.root.removeEventListener("focusin", onFocus);
        }, 5_000);
        this.runCleanups.add(() => {
          window.clearTimeout(warningTimer);
          removeDeferredFocus();
        });
      }

      let preferredThumbnailRailWidth = DEFAULT_THUMBNAIL_RAIL_WIDTH;
      try {
        preferredThumbnailRailWidth = this.options.thumbnailRail?.initialWidth() ??
          DEFAULT_THUMBNAIL_RAIL_WIDTH;
      } catch {
        // Preference persistence must not interrupt reading.
      }
      const initialThumbnailRailWidth = resolveThumbnailRailWidth(
        readingBody.clientWidth,
        preferredThumbnailRailWidth,
      );
      rail = new ThumbnailRail(thumbnailRoot, rendererSession, queue, {
        messages: this.messages,
        onMountedCountChange: (count) => {
          if (isCurrentRun() && this.thumbnailRail === rail) {
            this.root.dataset.mountedThumbnailCount = String(count);
          }
        },
        onReadyCountChange: (count) => {
          if (isCurrentRun() && this.thumbnailRail === rail) {
            this.root.dataset.readyThumbnailCount = String(count);
          }
        },
        onNavigate: navigate,
        thumbnailWidth: thumbnailPreviewWidth(initialThumbnailRailWidth),
      });
      this.thumbnailRail = rail;
      rail.start(viewController.state.currentSlideIndex);
      const searchableSlides = rendererSession.sourceAuthoredSlideText;
      if (searchableSlides?.length === rendererSession.slideCount) {
        let thumbnailsCollapsedBeforeSearch = false;
        const presentationSearchAvailable =
          rendererSession.speakerNoteContent?.length ===
            rendererSession.slideCount &&
          hasUsableSpeakerNoteContent(rendererSession.speakerNoteContent);
        const setThumbnailsCollapsed = (collapsed: boolean) => {
          this.root.dataset.thumbnailsCollapsed = String(collapsed);
          toggleThumbnails.setAttribute("aria-expanded", String(!collapsed));
          if (!collapsed) rail?.refresh();
          updateMountedCount();
        };
        let createdSearchRail: SlideSearchRail | null = null;
        try {
          createdSearchRail = new SlideSearchRail(thumbnailRoot, searchableSlides, {
            messages: this.messages,
            speakerNoteContent: rendererSession.speakerNoteContent,
            currentSlideIndex: () => viewController.state.currentSlideIndex,
            onNavigate: (slideId, intent) => {
              const targetIndex = rendererSession.slideIdentities?.indexOf(slideId) ?? -1;
              if (targetIndex < 0) return;
              if (intent?.surface === "speaker-notes" && intent.highlight) {
                notesHighlight = { slideId, snippet: intent.highlight };
                setNotesCollapsed(false);
              } else {
                notesHighlight = null;
              }
              if (targetIndex === viewController.state.currentSlideIndex) {
                renderSpeakerNotes(targetIndex);
                return;
              }
              navigate(targetIndex);
            },
            onDismiss: () => closeSearch(),
          });
        } catch {
          // Optional search setup must not invalidate an otherwise readable deck.
        }
        if (createdSearchRail !== null) {
          const activeSearchRail = createdSearchRail;
          searchRail = activeSearchRail;
          const searchOpenKey = presentationSearchAvailable
            ? "search.openPresentation"
            : "search.open";
          const searchCloseKey = presentationSearchAvailable
            ? "search.closePresentation"
            : "search.close";
          const searchButton = toolbar.primary.createEl("button", {
            type: "button",
            title: this.messages.text(searchOpenKey),
            attr: {
              "data-action": "open-slide-search",
              "aria-label": this.messages.text(searchOpenKey),
              "aria-pressed": "false",
            },
          });
          decorateOfficeViewerIconButton(searchButton, "lucide-search");
          const updateSearchButton = () => {
            searchButton.setAttribute(
              "aria-pressed",
              String(activeSearchRail.isOpen),
            );
            searchButton.title = this.messages.text(
              activeSearchRail.isOpen ? searchCloseKey : searchOpenKey,
            );
            searchButton.setAttribute(
              "aria-label",
              this.messages.text(
                activeSearchRail.isOpen ? searchCloseKey : searchOpenKey,
              ),
            );
          };
          openSearch = () => {
            if (searchRail?.isOpen) {
              searchRail.open();
              return;
            }
            thumbnailsCollapsedBeforeSearch =
              this.root.dataset.thumbnailsCollapsed === "true";
            thumbnailScrollTopBeforeSearch = thumbnailRoot.scrollTop;
            if (thumbnailsCollapsedBeforeSearch) setThumbnailsCollapsed(false);
            searchRail?.open();
            updateSearchButton();
          };
          closeSearch = () => {
            if (!searchRail?.isOpen) return;
            searchRail.close();
            thumbnailRoot.scrollTop = thumbnailScrollTopBeforeSearch;
            if (thumbnailsCollapsedBeforeSearch) setThumbnailsCollapsed(true);
            updateSearchButton();
            searchButton.focus();
          };
          this.openSlideContentSearchAction = openSearch;
          const toggleSearch = () => {
            if (searchRail?.isOpen) closeSearch();
            else openSearch();
          };
          searchButton.addEventListener("click", toggleSearch);
          this.runCleanups.add(() => {
            searchButton.removeEventListener("click", toggleSearch);
            activeSearchRail.dispose();
          });
        }
      }
      const railResizer = new ThumbnailRailResizer(
        readingBody,
        thumbnailRoot,
        rail,
        {
          messages: this.messages,
          preferredWidth: preferredThumbnailRailWidth,
          onCommit: (width) => this.options.thumbnailRail?.recordWidth(width),
        },
      );
      readingBody.insertBefore(railResizer.element, slideStage);
      this.runCleanups.add(() => railResizer.dispose());
      try {
        const unsubscribe = this.options.thumbnailRail?.subscribeWidth?.(
          (width) => railResizer.setPreferredWidth(width),
        );
        if (unsubscribe !== undefined) this.runCleanups.add(unsubscribe);
      } catch {
        // Cross-view preference updates are optional and non-blocking.
      }
      updateMountedCount();

      previousButton.addEventListener("click", () => {
        navigate(viewController.state.currentSlideIndex - 1);
      });
      nextButton.addEventListener("click", () => {
        navigate(viewController.state.currentSlideIndex + 1);
      });
      const jumpToRequestedPage = () => {
        const requestedPage = Number(pageInput.value);
        if (
          !Number.isInteger(requestedPage) ||
          requestedPage < 1 ||
          requestedPage > rendererSession.slideCount
        ) {
          status.textContent = this.messages.text("navigation.invalidPage", {
            total: rendererSession.slideCount,
          });
          pageInput.focus();
          return;
        }
        navigate(requestedPage - 1);
      };
      jumpButton.addEventListener("click", jumpToRequestedPage);
      jumpForm.addEventListener("submit", (event) => {
        event.preventDefault();
        jumpToRequestedPage();
      });

      const currentSlideIndex = viewController.state.currentSlideIndex;
      pageCounter.textContent = this.messages.text("page.counter", {
        current: currentSlideIndex + 1,
        total: rendererSession.slideCount,
      });
      pageInput.value = String(currentSlideIndex + 1);
      pageInput.max = String(rendererSession.slideCount);
      restoreControlState();
      status.textContent = "";
      this.root.dataset.state = "ready";
      this.setLifecyclePhase("ready");
      toggleThumbnails.addEventListener("click", () => {
        if (searchRail?.isOpen) closeSearch();
        const collapsed = this.root.dataset.thumbnailsCollapsed !== "true";
        this.root.dataset.thumbnailsCollapsed = String(collapsed);
        toggleThumbnails.setAttribute("aria-expanded", String(!collapsed));
        if (!collapsed) rail?.refresh();
        updateMountedCount();
      });
      toggleNotes.addEventListener("click", () => {
        if (!isCurrentRun()) return;
        setNotesCollapsed(this.root.dataset.notesCollapsed !== "true");
      });
      const copySlideMarkup = (embed: boolean) => {
        const target = copyTarget;
        if (target === null || !isCurrentRun()) return;
        actionStatus.textContent = "";
        void this.options.slideReferences?.copy(file, target, embed).then(
          () => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text(
                embed ? "reference.embedCopied" : "reference.copied",
              );
            }
          },
          () => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text(
                "reference.copyFailure",
              );
            }
          },
        );
      };
      const copySpeakerNotes = () => {
        const target = copyTarget;
        const paragraphs = currentNoteParagraphs;
        if (
          target === null ||
          paragraphs.length === 0 ||
          !isCurrentRun() ||
          this.options.slideReferences?.copyNotes === undefined
        ) {
          return;
        }
        actionStatus.textContent = "";
        void this.options.slideReferences.copyNotes(file, target, paragraphs).then(
          () => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text("notes.copied");
            }
          },
          () => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text("notes.copyFailure");
            }
          },
        );
      };
      const copySelectedSlideText = () => {
        const target = copyTarget;
        const selectedText = selectedSlideText;
        if (
          target === null ||
          selectedText === null ||
          !isCurrentRun() ||
          this.options.slideReferences?.copySelection === undefined
        ) {
          return;
        }
        actionStatus.textContent = "";
        void this.options.slideReferences.copySelection(
          file,
          target,
          selectedText,
        ).then(
          () => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text(
                "selection.copiedWithSource",
              );
            }
          },
          () => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text(
                "selection.copyFailure",
              );
            }
          },
        );
      };
      copyReferenceButton?.addEventListener("click", () => copySlideMarkup(false));
      copyEmbedButton?.addEventListener("click", () => copySlideMarkup(true));
      copyNotesButton?.addEventListener("click", copySpeakerNotes);
      copySelectionButton?.addEventListener("click", copySelectedSlideText);
      floatingCopySelectionButton?.addEventListener(
        "pointerdown",
        (event) => event.preventDefault(),
      );
      floatingCopySelectionButton?.addEventListener(
        "click",
        copySelectedSlideText,
      );
      const selectionDocument = slideContainer.ownerDocument;
      selectionDocument.addEventListener(
        "selectionchange",
        updateSlideTextSelectionAction,
      );
      slideContainer.addEventListener("pointerup", updateSlideTextSelectionAction);
      slideContainer.addEventListener("keyup", updateSlideTextSelectionAction);
      slideContainer.addEventListener("scroll", updateSlideTextSelectionAction);
      this.runCleanups.add(() => {
        selectionDocument.removeEventListener(
          "selectionchange",
          updateSlideTextSelectionAction,
        );
        slideContainer.removeEventListener(
          "pointerup",
          updateSlideTextSelectionAction,
        );
        slideContainer.removeEventListener("keyup", updateSlideTextSelectionAction);
        slideContainer.removeEventListener(
          "scroll",
          updateSlideTextSelectionAction,
        );
      });
      const fullscreen = this.options.fullscreen ?? createDefaultFullscreenActions();
      let lastKnownFullscreenState = false;
      const applyFullscreenState = (active: boolean) => {
        this.root.dataset.fullscreen = String(active);
        toggleFullscreen.textContent = this.messages.text(
          active ? "fullscreen.exit" : "fullscreen.button",
        );
        toggleFullscreen.setAttribute(
          "aria-label",
          this.messages.text(
            active ? "fullscreen.exit" : "fullscreen.enterLabel",
          ),
        );
      };
      const probeFullscreenState = (): {
        active: boolean;
        determinate: boolean;
      } => {
        if (!isCurrentRun()) {
          return { active: lastKnownFullscreenState, determinate: false };
        }
        try {
          lastKnownFullscreenState = fullscreen.isActive(this.root);
          applyFullscreenState(lastKnownFullscreenState);
          return { active: lastKnownFullscreenState, determinate: true };
        } catch {
          return { active: lastKnownFullscreenState, determinate: false };
        }
      };
      const unsubscribeFullscreen = fullscreen.subscribe(() => {
        probeFullscreenState();
      });
      this.runCleanups.add(unsubscribeFullscreen);
      probeFullscreenState();
      toggleFullscreen.addEventListener("click", () => {
        actionStatus.textContent = "";
        void Promise.resolve()
          .then(() => {
            // Preserve the state represented by the button if Chromium briefly
            // clears fullscreenElement while dispatching the second click.
            const wasActive = lastKnownFullscreenState;
            const { active } = probeFullscreenState();
            return active || wasActive
              ? fullscreen.exit()
              : fullscreen.enter(this.root);
          })
          .then(() => {
            probeFullscreenState();
          })
          .catch(() => {
            if (isCurrentRun()) {
              actionStatus.textContent = this.messages.text(
                "fullscreen.failure",
              );
              probeFullscreenState();
            }
          });
      });

      const onKeyDown = (event: KeyboardEvent) => {
        if (!isCurrentRun()) return;
        if (
          searchRail !== null &&
          (event.metaKey || event.ctrlKey) &&
          !event.altKey &&
          event.key.toLowerCase() === "f"
        ) {
          event.preventDefault();
          openSearch();
          return;
        }
        if (event.key === "Escape" && searchRail?.isOpen) {
          event.preventDefault();
          closeSearch();
          this.root.focus();
          return;
        }
        if (event.key === "Escape" && selectedSlideText !== null) {
          event.preventDefault();
          const selection = slideContainer.ownerDocument.getSelection();
          if (
            selection !== null &&
            selection.rangeCount === 1 &&
            slideContainer.contains(selection.getRangeAt(0).startContainer)
          ) {
            selection.removeAllRanges();
          }
          clearSlideTextSelectionAction();
          return;
        }
        if (isEditableTarget(event.target)) return;
        const delta = event.key === "ArrowLeft" || event.key === "PageUp"
          ? -1
          : event.key === "ArrowRight" || event.key === "PageDown"
            ? 1
            : 0;
        if (delta === 0) return;
        const target = viewController.state.currentSlideIndex + delta;
        if (target < 0 || target >= rendererSession.slideCount) return;
        event.preventDefault();
        navigate(target);
      };
      this.root.addEventListener("keydown", onKeyDown);
      this.runCleanups.add(() => this.root.removeEventListener("keydown", onKeyDown));
      this.root.focus();
      if (this.abortController === controller) this.abortController = null;
    } catch (error) {
      if (controller.signal.aborted || generation !== this.generation) return;
      this.teardownOpenResources();
      this.setLifecyclePhase("error");
      const openError =
        error instanceof PptxOpenError
          ? error
          : typeof error === "object" &&
              error !== null &&
              "name" in error &&
              error.name === "AbortError"
            ? new PptxOpenError("cancelled", "PPTX open was cancelled", {
                cause: error,
              })
          : new PptxOpenError(
              phase === "renderer" ? "incompatible" : "unknown",
              "Unexpected PPTX open failure",
              { cause: error },
            );
      this.showError(file, openError, generation);
    } finally {
      if (generation === this.generation) this.openPending = false;
    }
  }

  openUnsupportedLegacy(file: FileRef, sourceBytes?: number): void {
    this.teardownOpenResources();
    this.clearTimings();
    this.resetDiagnosticState();
    if (Number.isSafeInteger(sourceBytes) && sourceBytes! >= 0) {
      this.diagnosticSourceBytes = sourceBytes!;
    }
    const generation = ++this.generation;
    this.openPending = false;
    this.disposed = false;
    this.setLifecyclePhase("error");
    this.showError(
      file,
      new PptxOpenError(
        "unsupported-legacy",
        "Legacy PPT is intentionally unsupported",
      ),
      generation,
    );
  }

  getPerformanceDiagnostics(): PptxViewSessionDiagnostics {
    const queue = this.backgroundQueue?.diagnostics;
    return {
      generation: this.generation,
      openPending: this.openPending,
      rendererActive: this.rendererSession !== null,
      disposed: this.disposed,
      lifecyclePhase: this.lifecyclePhase,
      backgroundPending: queue?.pending ?? 0,
      backgroundRunning: queue?.running ?? 0,
      mountedThumbnails: this.thumbnailRail?.mountedCount ?? 0,
      readyThumbnails: this.thumbnailRail?.readyCount ?? 0,
    };
  }

  navigateToReference(target: SlideReferenceTarget): boolean {
    if (this.referenceNavigator === null) return false;
    this.referenceNavigator(target);
    return true;
  }

  dispose(): void {
    this.generation += 1;
    this.openPending = false;
    this.disposed = true;
    this.lifecyclePhase = "disposed";
    this.teardownOpenResources();
    this.toolbar = null;
    this.contentHost = null;
    this.root.replaceChildren();
    delete this.root.dataset.state;
    delete this.root.dataset.lifecyclePhase;
    this.clearTimings();
    delete this.root.dataset.errorCategory;
    delete this.root.dataset.warningCategories;
    delete this.root.dataset.thumbnailsCollapsed;
    delete this.root.dataset.notesCollapsed;
    delete this.root.dataset.fullscreen;
    delete this.root.dataset.mountedThumbnailCount;
    delete this.root.dataset.readyThumbnailCount;
    delete this.root.dataset.referenceSlideId;
    delete this.root.dataset.referenceCreatedSlide;
    delete this.root.dataset.referenceCurrentSlide;
  }

  private showError(
    file: FileRef,
    error: PptxOpenError,
    generation: number,
  ): void {
    const { toolbar, content } = this.ensurePersistentShell();
    this.disableShellSearchActions(toolbar);
    content.replaceChildren();
    const { panel, actionStatus } = createOfficeViewerErrorSurface({
      title: this.messages.text(errorMessageKeys[error.category]),
      safetyNote: this.messages.text(
        error.category === "unsupported-legacy"
          ? "error.sourceUnmodifiedLegacy"
          : "error.sourceUnmodified",
      ),
      retry: {
        label: this.messages.text("error.retry"),
        action: "retry",
        onClick: () => {
          if (error.category === "unsupported-legacy") {
            this.openUnsupportedLegacy(
              file,
              this.diagnosticSourceBytes ?? undefined,
            );
            return;
          }
          void this.open(file);
        },
      },
      classNames: {
        root: "pptx-viewer__error",
        status: "pptx-viewer__status",
        safetyNote: "pptx-viewer__safety-note",
        actions: "pptx-viewer__actions",
        actionStatus: "pptx-viewer__action-status",
      },
    });

    const actions = panel.querySelector(".office-viewer-error__actions");
    if (actions !== null) {
      if (error.category !== "unsupported-legacy") {
        const companionButton = this.createCompanionNoteButton(
          file,
          generation,
          actionStatus,
          () => this.options.fullscreen ?? createDefaultFullscreenActions(),
        );
        if (companionButton) actions.append(companionButton);
      }
      const openExternally = this.createExternalOpenButton(
        file,
        generation,
        actionStatus,
      );
      if (openExternally) actions.append(openExternally);
      const diagnosticButton = this.createDiagnosticButton(actionStatus);
      if (diagnosticButton) actions.append(diagnosticButton);
    }

    content.append(panel);
    this.root.dataset.state = "error";
    this.root.dataset.errorCategory = error.category;
    this.diagnosticError = error.category;
    delete this.root.dataset.warningCategories;
  }

  private createCompanionNoteButton(
    file: FileRef,
    generation: number,
    actionStatus: HTMLElement,
    resolveFullscreen: () => PptxFullscreenActions,
  ): HTMLButtonElement | null {
    if (this.options.companionNote === undefined
      || this.options.supportsPptxActions?.(file) === false) return null;
    const button = createEl("button", {
      type: "button",
      title: this.messages.text("companion.open"),
      attr: {
        "data-action": "open-companion-note",
        "aria-label": this.messages.text("companion.openLabel"),
      },
    });
    decorateOfficeViewerIconButton(button, "lucide-pen-line");
    button.addEventListener("click", () => {
      if (generation !== this.generation || this.disposed) return;
      actionStatus.textContent = "";
      const fullscreen = resolveFullscreen();
      const run = async () => {
        if (fullscreen.isActive(this.root)) {
          try {
            await fullscreen.exit();
          } catch {
            actionStatus.textContent = this.messages.text("fullscreen.failure");
            return;
          }
        }
        if (generation !== this.generation || this.disposed) return;
        try {
          const status = await this.options.companionNote!.open(file);
          if (generation !== this.generation || this.disposed) return;
          actionStatus.textContent = status;
        } catch {
          if (generation !== this.generation || this.disposed) return;
          actionStatus.textContent = this.messages.text("companion.openFailure");
        }
      };
      void run();
    });
    return button;
  }

  private showMissingReference(file: FileRef, generation: number): void {
    const { toolbar, content } = this.ensurePersistentShell();
    this.disableShellSearchActions(toolbar);
    content.replaceChildren();
    const panel = content.createDiv({ cls: "pptx-viewer__error" });
    panel.createDiv({
      cls: "pptx-viewer__status",
      text: this.messages.text("reference.missing"),
      attr: { role: "status" },
    });
    const actions = panel.createDiv({ cls: "pptx-viewer__actions" });
    const openPresentation = actions.createEl("button", {
      type: "button",
      text: this.messages.text("reference.openPresentation"),
      attr: { "data-action": "open-presentation" },
    });
    openPresentation.addEventListener("click", () => {
      void this.open(file);
    });
    const actionStatus = panel.createDiv({
      cls: "pptx-viewer__action-status",
      attr: { role: "status", "aria-live": "polite" },
    });
    const openExternally = this.createExternalOpenButton(
      file,
      generation,
      actionStatus,
    );
    if (openExternally) actions.append(openExternally);
    this.root.dataset.state = "stale-reference";
    delete this.root.dataset.errorCategory;
    delete this.root.dataset.warningCategories;
    this.setLifecyclePhase("degraded");
  }

  private createExternalOpenButton(
    file: FileRef,
    generation: number,
    actionStatus: HTMLElement,
  ): HTMLButtonElement | null {
    if (!this.options.openExternally) return null;
    const button = createEl("button", {
      type: "button",
      attr: {
        "data-action": "open-externally",
        "aria-label": this.messages.text("external.open"),
        title: this.messages.text("external.open"),
      },
    });
    decorateOfficeViewerIconButton(button, "lucide-external-link");
    button.addEventListener("click", () => {
      actionStatus.textContent = "";
      void this.options.openExternally?.(file).catch(() => {
        if (generation === this.generation) {
          actionStatus.textContent = this.messages.text("external.failure");
        }
      });
    });
    return button;
  }

  private showCompatibilityWarnings(
    container: HTMLElement,
    categories: readonly PptxCompatibilityWarningCategory[],
  ): void {
    const unique = [...new Set(categories)].sort();
    this.diagnosticWarnings = unique;
    if (
      this.options.diagnostics !== undefined &&
      this.options.diagnostics.enabled() !== true
    ) {
      container.replaceChildren();
      delete this.root.dataset.warningCategories;
      return;
    }
    container.replaceChildren();
    for (const category of unique) {
      container.createEl("p", {
        text: this.messages.text(warningMessageKeys[category]),
        attr: { "data-warning-category": category },
      });
    }
    if (unique.length > 0) {
      this.root.dataset.warningCategories = unique.join(",");
    } else {
      delete this.root.dataset.warningCategories;
    }
  }

  private createDiagnosticButton(
    actionStatus: HTMLElement,
  ): HTMLButtonElement | null {
    const diagnostics = this.options.diagnostics;
    if (!diagnostics || diagnostics.enabled() !== true) return null;
    const button = createEl("button", {
      type: "button",
      text: this.messages.text("diagnostics.copy"),
      attr: { "data-action": "copy-diagnostics" },
    });
    button.addEventListener("click", () => {
      actionStatus.textContent = "";
      const summary = createDiagnosticSummary({
        environment: diagnostics.environment,
        sourceBytes: this.diagnosticSourceBytes,
        slideCount: this.diagnosticSlideCount,
        lifecyclePhase: this.lifecyclePhase,
        warningCategories: this.diagnosticWarnings,
        errorCategory: this.diagnosticError,
        timingsMs: {
          metadata: this.readTiming("metadataMs"),
          firstReadable: this.readTiming("firstReadableMs"),
          lastSlideSwitch: this.readTiming("lastSlideSwitchMs"),
        },
        features: {
          thumbnails: this.diagnosticThumbnails,
          prefetch: this.diagnosticPrefetch,
          rememberReadingPosition: diagnostics.rememberReadingPosition(),
          externalOpen: this.options.openExternally !== undefined,
        },
      });
      void diagnostics.copy(summary).then(
        () => { actionStatus.textContent = this.messages.text("diagnostics.copied"); },
        () => {
          actionStatus.textContent = this.messages.text(
            "diagnostics.copyFailure",
          );
        },
      );
    });
    return button;
  }

  private readTiming(key: "metadataMs" | "firstReadableMs" | "lastSlideSwitchMs"):
    number | null {
    const value = Number(this.root.dataset[key]);
    return Number.isFinite(value) ? value : null;
  }

  private resetDiagnosticState(): void {
    this.diagnosticSourceBytes = null;
    this.diagnosticSlideCount = null;
    this.diagnosticWarnings = [];
    this.diagnosticError = null;
    this.diagnosticThumbnails = false;
    this.diagnosticPrefetch = false;
  }

  private teardownOpenResources(): void {
    const abortController = this.abortController;
    const cleanups = [...this.runCleanups];
    const thumbnailRail = this.thumbnailRail;
    const viewerController = this.viewerController;
    const backgroundQueue = this.backgroundQueue;
    const rendererSession = this.rendererSession;

    this.abortController = null;
    this.runCleanups.clear();
    this.thumbnailRail = null;
    this.viewerController = null;
    this.backgroundQueue = null;
    this.rendererSession = null;
    this.referenceNavigator = null;
    this.openSlideContentSearchAction = null;
    if (this.root.dataset.readyThumbnailCount !== undefined) {
      this.root.dataset.readyThumbnailCount = "0";
    }

    try {
      abortController?.abort();
    } catch {
      // Continue releasing owned resources if cancellation hooks misbehave.
    }
    for (const cleanup of cleanups) {
      try {
        cleanup();
      } catch {
        // One detached platform listener must not block remaining cleanup.
      }
    }
    for (const dispose of [
      () => thumbnailRail?.dispose(),
      () => viewerController?.dispose(),
      () => backgroundQueue?.dispose(),
      () => rendererSession?.dispose(),
    ]) {
      try {
        dispose();
      } catch {
        // A candidate or component cleanup failure must not strand later owners.
      }
    }
  }

  private clearTimings(): void {
    delete this.root.dataset.metadataMs;
    delete this.root.dataset.firstReadableMs;
    delete this.root.dataset.lastSlideSwitchMs;
  }

  private setLifecyclePhase(phase: PptxLifecyclePhase): void {
    this.lifecyclePhase = phase;
    this.root.dataset.lifecyclePhase = phase;
  }
}
