import {
  apiVersion,
  editorInfoField,
  editorLivePreviewField,
  getLanguage,
  Plugin,
  TFile,
  type FileSystemAdapter,
  type MarkdownRenderChild,
} from "obsidian";
import manifest from "../manifest.json" with { type: "json" };
import releaseContract from "../release-contract.json" with { type: "json" };
import { createMessageTranslator } from "./i18n";
import { createLivePreviewSlideEmbedExtension } from "./live-preview-slide-embed";
import { OfficeViewerSettingTab } from "./office-viewer-setting-tab";
import { PptxFileView, PPTX_VIEW_TYPE } from "./pptx-file-view";
import {
  OfficeViewerDataStore,
  type FileFingerprint,
} from "./office-viewer-data-store";
import { reportNonFatalError } from "./report-error";
import { processPptxSlideEmbeds } from "./pptx-slide-embed";
import { createPptxRendererAdapter } from "./renderer/create-pptx-renderer-adapter";
import { SlideEmbedScheduler } from "./slide-embed-scheduler";
import { createExternalOpenAction } from "./external-open";
import { resolveSlideEmbedFile } from "./resolve-slide-embed-file";
import { createCompanionNoteVault } from "./companion-note-vault";
import { PresentationCompanionNoteService } from "./presentation-companion-note-service";
import { DocxFileView, DOCX_VIEW_TYPE } from "./docx/docx-file-view";
import { createDocxMessageTranslator } from "./docx/docx-messages";
import { createDocxRendererAdapter } from "./docx/renderer/create-docx-renderer-adapter";
import { createDocxAttachmentOpener } from "./docx/open-docx-attachment";
import { processOfficeFileEmbeds, type OfficeFileEmbedOptions } from "./office-file-embed";
import { createLegacyOfficePreview, LegacyOfficeUnavailable } from "./legacy-office-preview";
import { registerOfficeTabOpening } from "./office-tab-opening";

async function openExternalUrl(url: string): Promise<void> {
  const protocol = new URL(url).protocol;
  if (
    protocol !== "https:" &&
    protocol !== "http:" &&
    protocol !== "mailto:"
  ) {
    throw new Error("Blocked external URL protocol");
  }
  // eslint-disable-next-line @typescript-eslint/no-require-imports -- Electron shell is available in desktop Obsidian
  const { shell } = require("electron") as {
    shell: { openExternal(url: string): Promise<void> };
  };
  await shell.openExternal(url);
}

function fingerprint(file: TFile): FileFingerprint {
  return {
    path: file.path,
    size: file.stat.size,
    mtime: file.stat.mtime,
  };
}

export default class OfficeViewerPlugin extends Plugin {
  private readonly views = new Set<PptxFileView>();
  private readonly docxViews = new Set<DocxFileView>();
  private store: OfficeViewerDataStore | undefined;
  private companionNotes: PresentationCompanionNoteService | undefined;
  private embedScheduler: SlideEmbedScheduler | undefined;
  private attachmentOpener: ReturnType<typeof createDocxAttachmentOpener> | undefined;
  private readonly embedChildren = new Set<MarkdownRenderChild>();
  private unloading = false;

  override async onload(): Promise<void> {
    this.unloading = false;
    const language = getLanguage();
    const messages = createMessageTranslator(language);
    const diagnosticEnvironment = {
      pluginVersion: manifest.version,
      obsidianVersion: apiVersion,
      rendererVersion: process.env.PPTX_RENDERER_VERSION ?? "unknown",
      operatingSystem: `${process.platform}-${process.arch}`,
    };
    const store = new OfficeViewerDataStore({
      loadData: () => this.loadData(),
      saveData: (data) => this.saveData(data),
    });
    this.store = store;

    try {
      await store.initialize();
    } catch (error) {
      if (this.store === store) this.store = undefined;
      await store.dispose().catch(() => undefined);
      throw error;
    }

    if (this.unloading || this.store !== store) {
      await store.dispose().catch(() => undefined);
      return;
    }

    const companionNotes = new PresentationCompanionNoteService(
      store,
      createCompanionNoteVault(this.app),
    );
    this.companionNotes = companionNotes;
    companionNotes.reconcile();

    const embedScheduler = new SlideEmbedScheduler(2);
    this.embedScheduler = embedScheduler;

    this.addSettingTab(
      new OfficeViewerSettingTab(this.app, this, store, messages),
    );
    this.registerEvent(
      this.app.vault.on("rename", (file, oldPath) => {
        if (!(file instanceof TFile)) return;
        store.rename(oldPath, fingerprint(file));
        const extension = file.extension.toLowerCase();
        if (extension === "pptx") {
          void companionNotes.handleSourceRename(oldPath, file.path).catch(
            (error: unknown) => {
              reportNonFatalError(
                "Failed to migrate presentation companion note",
                error,
              );
            },
          );
        } else if (extension === "md") {
          void companionNotes.handleNoteRename(oldPath, file.path).catch(
            (error: unknown) => {
              reportNonFatalError(
                "Failed to handle companion note rename",
                error,
              );
            },
          );
        }
      }),
    );
    this.registerEvent(
      this.app.vault.on("delete", (file) => {
        if (!(file instanceof TFile)) return;
        store.delete(file.path);
        void companionNotes.handleDelete(file.path);
      }),
    );
    const adapter = this.app.vault.adapter as Partial<FileSystemAdapter> | undefined;
    const legacyPreview = process.env.OFFICE_VIEWER_LEGACY_SOFFICE
      ? createLegacyOfficePreview(process.env.OFFICE_VIEWER_LEGACY_SOFFICE, adapter?.getBasePath?.() ?? "")
      : undefined;
    const readPreviewBinary = (file: TFile, signal: AbortSignal) => {
      const extension = file.extension.toLowerCase();
      const readSource = () => this.app.vault.readBinary(file);
      if (extension !== "ppt" && extension !== "doc") return readSource();
      if (legacyPreview === undefined) throw new LegacyOfficeUnavailable();
      return legacyPreview(extension, file.stat.size, readSource, signal);
    };
    this.registerView(PPTX_VIEW_TYPE, (leaf) => {
      let view: PptxFileView;
      view = new PptxFileView(
        leaf,
        () => this.views.delete(view),
        {
          initialSlideFor: (file, slideCount) =>
            store.resolve(fingerprint(file), slideCount),
          record: (file, slideIndex) =>
            store.record(fingerprint(file), slideIndex),
          initialThumbnailRailWidth: () => store.settings.thumbnailRailWidth,
          recordThumbnailRailWidth: (width) =>
            store.setThumbnailRailWidth(width),
          subscribeThumbnailRailWidth: (listener) =>
            store.subscribeThumbnailRailWidth(listener),
          rememberReadingPosition: () =>
            store.settings.rememberReadingPosition,
          diagnosticSummary: () => store.settings.diagnosticSummary,
          ensureCompanionNote: (sourcePath) =>
            companionNotes.ensureCompanionNote(sourcePath),
        },
        messages,
        diagnosticEnvironment,
        legacyPreview === undefined ? undefined : readPreviewBinary,
      );
      this.views.add(view);
      return view;
    });
    const docxMessages = createDocxMessageTranslator(language);
    const docxRenderer = createDocxRendererAdapter({
      pageLayout: true,
      messages: docxMessages,
      unavailablePlaceholder: docxMessages.text("unavailablePlaceholder"),
    });
    const attachmentOpener = this.attachmentOpener = createDocxAttachmentOpener();
    const openExternally = createExternalOpenAction(this.app);
    this.registerView(DOCX_VIEW_TYPE, (leaf) => {
      let view: DocxFileView;
      view = new DocxFileView(
        leaf,
        {
          renderer: docxRenderer,
          messages: docxMessages,
          readBinary: readPreviewBinary,
          openExternalUrl,
          openAttachment: attachmentOpener.open,
          openInDefaultApplication:
            openExternally === undefined
              ? undefined
              : (file: TFile) => openExternally(file),
        },
        () => this.docxViews.delete(view),
      );
      this.docxViews.add(view);
      return view;
    });
    const embedRenderer = createPptxRendererAdapter();
    const presentationExtensions = releaseContract.extensionRoutes[PPTX_VIEW_TYPE];
    const documentExtensions = legacyPreview === undefined
      ? releaseContract.extensionRoutes[DOCX_VIEW_TYPE]
      : [...releaseContract.extensionRoutes[DOCX_VIEW_TYPE], "doc"];
    const officeExtensions = new Set([
      ...presentationExtensions, ...documentExtensions,
    ]);
    const canPreview = (file: TFile) => officeExtensions.has(file.extension.toLowerCase());
    const resolveFile = (sourcePath: string, notePath: string) =>
      resolveSlideEmbedFile(this.app, sourcePath, notePath);
    const openSource = (linkTarget: string, notePath: string) =>
      this.app.workspace.openLinkText(linkTarget, notePath, "tab");
    const officeFiles: OfficeFileEmbedOptions<TFile> = {
      resolveFile: (sourcePath, notePath) => {
        const file = resolveFile(sourcePath, notePath);
        return file !== null
          && sourcePath.toLowerCase().endsWith(`.${file.extension.toLowerCase()}`)
          ? file
          : null;
      },
      readBinary: readPreviewBinary,
      openInTab: async (file) => {
        if (!canPreview(file)) throw new Error("No preview view is registered for this format");
        await this.app.workspace.getLeaf("tab").openFile(file, { active: true });
      },
      openExternally,
      openExternalUrl,
      openAttachment: attachmentOpener.open,
      pptxRenderer: embedRenderer,
      docxRenderer,
      scheduler: embedScheduler,
      messages,
      docxMessages,
    };
    const lifecycle = {
      register: (child: MarkdownRenderChild) => this.embedChildren.add(child),
      unregister: (child: MarkdownRenderChild) => this.embedChildren.delete(child),
    };
    this.registerMarkdownPostProcessor((element, context) => {
      processPptxSlideEmbeds(element, context, {
        app: this.app,
        renderer: embedRenderer,
        scheduler: embedScheduler,
        messages,
        showDiagnostics: () => store.settings.diagnosticSummary,
        openExternally,
        openSource,
        lifecycle,
      });
      processOfficeFileEmbeds(element, context, officeFiles, lifecycle);
    }, 100);
    this.registerEditorExtension(
      createLivePreviewSlideEmbedExtension({
        livePreviewField: editorLivePreviewField,
        getSourcePath: (state) =>
          state.field(editorInfoField, false)?.file?.path ?? "",
        resolveFile,
        readBinary: (file) => this.app.vault.readBinary(file),
        renderer: embedRenderer,
        scheduler: embedScheduler,
        messages,
        showDiagnostics: () => store.settings.diagnosticSummary,
        openExternally,
        officeFiles,
        openSource,
      }),
    );
    this.registerExtensions(presentationExtensions, PPTX_VIEW_TYPE);
    this.registerExtensions(documentExtensions, DOCX_VIEW_TYPE);
    if (typeof this.app.workspace?.openLinkText === "function") {
      this.register(registerOfficeTabOpening(this.app, canPreview));
    }
  }

  override onunload(): void {
    this.unloading = true;
    void this.attachmentOpener?.dispose();
    this.attachmentOpener = undefined;
    for (const view of [...this.views]) view.dispose();
    this.views.clear();
    for (const view of [...this.docxViews]) view.dispose();
    this.docxViews.clear();
    for (const child of [...this.embedChildren]) child.unload();
    this.embedChildren.clear();
    this.embedScheduler?.dispose();
    this.embedScheduler = undefined;
    this.companionNotes = undefined;
    const store = this.store;
    this.store = undefined;
    if (store !== undefined) {
      void store.dispose().catch((error: unknown) => {
        reportNonFatalError(
          "Failed to save PPTX reading positions during unload",
          error,
        );
      });
    }
  }
}
