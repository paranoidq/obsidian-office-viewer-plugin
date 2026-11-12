import {
  FileView,
  MarkdownView,
  Scope,
  TFile,
  type WorkspaceLeaf,
} from "obsidian";
import {
  ENGLISH_MESSAGE_TRANSLATOR,
  type MessageTranslator,
} from "./i18n";
import { PptxViewSession } from "./pptx-view-session";
import type { PptxViewSessionDiagnostics } from "./pptx-view-session";
import type { DiagnosticEnvironment } from "./diagnostic-summary";
import { createPptxRendererAdapter } from "./renderer/create-pptx-renderer-adapter";
import {
  formatSelectedSlideTextCopyMarkup,
  formatSlideReferenceMarkup,
  formatSpeakerNotesCopyMarkup,
  parseSlideReferenceFragment,
  type SlideReferenceTarget,
} from "./slide-reference";
import { createExternalOpenAction } from "./external-open";
import type { CompanionNoteEnsureResult } from "./presentation-companion-note-service";
import { canonicalCompanionNotePath } from "./presentation-companion-note";

export const PPTX_VIEW_TYPE = "pptx-viewer";

export interface PptxFileViewState {
  initialSlideFor(file: TFile, slideCount: number): number;
  record(file: TFile, slideIndex: number): void;
  initialThumbnailRailWidth(): number;
  recordThumbnailRailWidth(width: number): void;
  subscribeThumbnailRailWidth(listener: (width: number) => void): () => void;
  rememberReadingPosition(): boolean;
  diagnosticSummary(): boolean;
  ensureCompanionNote?(
    sourcePath: string,
  ): Promise<CompanionNoteEnsureResult>;
}

export class PptxFileView extends FileView {
  private readonly session: PptxViewSession<TFile>;
  private disposed = false;
  private currentFile: TFile | null = null;
  private pendingReferenceTarget: SlideReferenceTarget | undefined;
  private companionLeaf: WorkspaceLeaf | null = null;
  private companionNotePath: string | null = null;
  private companionActionGeneration = 0;

  constructor(
    leaf: WorkspaceLeaf,
    private readonly onDisposed: () => void = () => {},
    private readonly state?: PptxFileViewState,
    private readonly messages: MessageTranslator = ENGLISH_MESSAGE_TRANSLATOR,
    diagnosticEnvironment?: DiagnosticEnvironment,
    private readonly previewReader?: (file: TFile, signal: AbortSignal) => Promise<ArrayBuffer>,
  ) {
    super(leaf);
    this.contentEl.replaceChildren();
    const root = this.contentEl.createDiv();
    this.session = new PptxViewSession(
      root,
      { readBinary: previewReader ?? ((file) => this.app.vault.readBinary(file)) },
      createPptxRendererAdapter(),
      {
        messages: this.messages,
        supportsPptxActions: (file) => file.extension.toLowerCase() !== "ppt",
        openExternally: createExternalOpenAction(this.app),
        positions: state,
        thumbnailRail: state === undefined
          ? undefined
          : {
              initialWidth: () => state.initialThumbnailRailWidth(),
              recordWidth: (width) => state.recordThumbnailRailWidth(width),
              subscribeWidth: (listener) =>
                state.subscribeThumbnailRailWidth(listener),
            },
        diagnostics: diagnosticEnvironment === undefined
          ? undefined
          : {
              environment: diagnosticEnvironment,
              rememberReadingPosition: () =>
                state?.rememberReadingPosition() ?? false,
              enabled: () => state?.diagnosticSummary() ?? false,
              copy: async (summary) => navigator.clipboard.writeText(summary),
            },
        slideReferences: {
          copy: async (file, target, embed) => {
            const alias = this.messages.text("reference.alias", {
              name: file.basename,
              slide: target.createdSlideNumber,
            });
            await navigator.clipboard.writeText(formatSlideReferenceMarkup({
              sourcePath: file.path,
              alias,
              slideId: target.slideId,
              createdSlideNumber: target.createdSlideNumber,
              embed,
            }));
          },
          copyNotes: async (file, target, paragraphs) => {
            const alias = this.messages.text("reference.alias", {
              name: file.basename,
              slide: target.createdSlideNumber,
            });
            await navigator.clipboard.writeText(formatSpeakerNotesCopyMarkup(
              paragraphs,
              {
                sourcePath: file.path,
                alias,
                slideId: target.slideId,
                createdSlideNumber: target.createdSlideNumber,
                embed: false,
              },
            ));
          },
          copySelection: async (file, target, selectedText) => {
            const alias = this.messages.text("reference.alias", {
              name: file.basename,
              slide: target.createdSlideNumber,
            });
            await navigator.clipboard.writeText(
              formatSelectedSlideTextCopyMarkup(selectedText, {
                sourcePath: file.path,
                alias,
                slideId: target.slideId,
                createdSlideNumber: target.createdSlideNumber,
                embed: false,
              }),
            );
          },
        },
        companionNote: state?.ensureCompanionNote === undefined
          ? undefined
          : {
              open: (file) => this.openCompanionNote(file),
            },
      },
    );
    this.scope = new Scope(this.app.scope);
    this.scope.register(["Mod"], "f", () =>
      this.session.openSlideContentSearch() ? false : undefined
    );
  }

  override getViewType(): string {
    return PPTX_VIEW_TYPE;
  }

  override getDisplayText(): string {
    return this.file?.basename ?? this.messages.text("viewer.fallbackTitle");
  }

  getPerformanceDiagnostics(): PptxViewSessionDiagnostics {
    return this.session.getPerformanceDiagnostics();
  }

  override async onLoadFile(file: TFile): Promise<void> {
    if (this.currentFile !== null && this.currentFile.path !== file.path) {
      this.releaseCompanionLeaf();
    }
    this.currentFile = file;
    this.companionActionGeneration += 1;
    if (file.extension.toLowerCase() === "ppt") {
      this.pendingReferenceTarget = undefined;
      this.releaseCompanionLeaf();
      if (this.previewReader === undefined) {
        this.session.openUnsupportedLegacy(file, file.stat.size);
        return;
      }
    }
    const referenceTarget = this.pendingReferenceTarget;
    this.pendingReferenceTarget = undefined;
    await this.session.open(file, referenceTarget);
    if (this.currentFile !== file || this.pendingReferenceTarget === undefined) {
      return;
    }
    const lateReferenceTarget = this.pendingReferenceTarget;
    if (this.session.navigateToReference(lateReferenceTarget)) {
      this.pendingReferenceTarget = undefined;
    }
  }

  override setEphemeralState(state: unknown): void {
    super.setEphemeralState(state);
    const subpath = typeof state === "object" && state !== null &&
        "subpath" in state && typeof state.subpath === "string"
      ? state.subpath
      : undefined;
    const target = subpath === undefined
      ? undefined
      : parseSlideReferenceFragment(subpath) ?? undefined;
    this.pendingReferenceTarget = target;
    if (
      target !== undefined &&
      this.currentFile !== null &&
      this.currentFile.extension.toLowerCase() === "pptx"
    ) {
      if (this.session.navigateToReference(target)) {
        this.pendingReferenceTarget = undefined;
      }
    }
  }

  override async onClose(): Promise<void> {
    this.dispose();
  }

  dispose = (): void => {
    if (this.disposed) return;
    this.disposed = true;
    this.companionActionGeneration += 1;
    this.currentFile = null;
    this.pendingReferenceTarget = undefined;
    this.releaseCompanionLeaf();
    this.session.dispose();
    this.contentEl.replaceChildren();
    this.onDisposed();
  };

  private async openCompanionNote(file: TFile): Promise<string> {
    const state = this.state;
    if (state?.ensureCompanionNote === undefined) {
      return this.messages.text("companion.openFailure");
    }
    const actionGeneration = this.companionActionGeneration;
    const sourcePath = file.path;
    const result = await state.ensureCompanionNote(sourcePath);
    if (
      this.disposed ||
      actionGeneration !== this.companionActionGeneration ||
      this.currentFile?.path !== sourcePath
    ) {
      return this.messages.text("companion.openFailure");
    }

    if (result.status === "failure") {
      return result.reason === "vault-write"
        ? this.messages.text("companion.writeFailure")
        : this.messages.text("companion.openFailure");
    }

    if (result.status === "target-occupied") {
      return this.messages.text("companion.targetOccupied", {
        path: result.notePath,
      });
    }

    try {
      await this.focusCompanionNote(result.notePath, actionGeneration, sourcePath);
    } catch {
      return this.messages.text("companion.openFailure");
    }

    if (
      this.disposed ||
      actionGeneration !== this.companionActionGeneration ||
      this.currentFile?.path !== sourcePath
    ) {
      return this.messages.text("companion.openFailure");
    }

    if (result.conflict) {
      return this.messages.text("companion.conflict", {
        path: canonicalCompanionNotePath(sourcePath) ?? result.notePath,
      });
    }

    switch (result.status) {
      case "created":
        return this.messages.text("companion.created");
      case "adopted":
        return this.messages.text("companion.adopted");
      case "migrated":
        return this.messages.text("companion.migrated");
      case "opened":
        return this.messages.text("companion.opened");
    }
  }

  private async focusCompanionNote(
    notePath: string,
    actionGeneration: number,
    sourcePath: string,
  ): Promise<void> {
    const note = this.app.vault.getAbstractFileByPath(notePath);
    if (!(note instanceof TFile) || note.extension !== "md") {
      throw new Error(`Companion note missing: ${notePath}`);
    }

    if (
      this.disposed ||
      actionGeneration !== this.companionActionGeneration ||
      this.currentFile?.path !== sourcePath
    ) {
      return;
    }

    if (this.leafShowsCompanionNote(this.companionLeaf, notePath)) {
      this.app.workspace.setActiveLeaf(this.companionLeaf!, { focus: true });
      return;
    }

    this.releaseCompanionLeaf();
    if (
      this.disposed ||
      actionGeneration !== this.companionActionGeneration ||
      this.currentFile?.path !== sourcePath
    ) {
      return;
    }

    const companionLeaf = this.app.workspace.createLeafBySplit(
      this.leaf,
      "vertical",
      false,
    );
    await companionLeaf.openFile(note);
    if (
      this.disposed ||
      actionGeneration !== this.companionActionGeneration ||
      this.currentFile?.path !== sourcePath
    ) {
      return;
    }
    this.app.workspace.setActiveLeaf(companionLeaf, { focus: true });
    this.companionLeaf = companionLeaf;
    this.companionNotePath = notePath;
  }

  private leafShowsCompanionNote(
    leaf: WorkspaceLeaf | null,
    notePath: string,
  ): boolean {
    if (leaf === null || !this.isLeafInWorkspace(leaf)) {
      return false;
    }
    const view = leaf.view;
    return view instanceof MarkdownView && view.file?.path === notePath;
  }

  private isLeafInWorkspace(leaf: WorkspaceLeaf): boolean {
    let found = false;
    this.app.workspace.iterateAllLeaves((candidate) => {
      if (candidate === leaf) found = true;
    });
    return found;
  }

  private releaseCompanionLeaf(): void {
    this.companionLeaf = null;
    this.companionNotePath = null;
  }
}
