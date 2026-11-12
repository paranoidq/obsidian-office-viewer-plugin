import type { MessageKey, MessageTranslator } from "./i18n";
import { registerOfficeEmbedTextCopy } from "./office-embed-text-copy";
import { decorateOfficeViewerIconButton } from "./office-viewer-chrome";
import {
  PptxOpenError,
  type PptxOpenErrorCategory,
} from "./pptx-open-error";
import type {
  PptxRendererAdapter,
  PptxRendererSession,
} from "./renderer/pptx-renderer-adapter";
import {
  formatSlideReferenceLinkTarget,
  type SlideReferenceTarget,
} from "./slide-reference";
import { SlideEmbedScheduler } from "./slide-embed-scheduler";

export type SlideEmbedDisplayState =
  | "waiting"
  | "queued"
  | "loading"
  | "ready"
  | "stale-reference"
  | "error"
  | "missing-source";

/** Minimal source-file surface needed by the host-agnostic embed core. */
export interface SlideEmbedSourceFile {
  readonly basename: string;
}

export interface SlideEmbedCorePorts<
  TFile extends SlideEmbedSourceFile = SlideEmbedSourceFile,
> {
  readonly readBinary: (file: TFile) => Promise<ArrayBuffer>;
  readonly renderer: PptxRendererAdapter;
  readonly scheduler: SlideEmbedScheduler;
  readonly messages: MessageTranslator;
  readonly showDiagnostics: () => boolean;
  readonly openSource?: (linkTarget: string) => void | Promise<void>;
  readonly openExternally?: (file: TFile) => Promise<void>;
  readonly now?: () => number;
}

export interface SlideEmbedCoreInput<
  TFile extends SlideEmbedSourceFile = SlideEmbedSourceFile,
> {
  readonly file: TFile | null;
  readonly sourcePath: string;
  readonly target: SlideReferenceTarget;
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

interface EmbedElementOptions {
  readonly className?: string;
  readonly text?: string;
  readonly attributes?: Readonly<Record<string, string>>;
}

function appendEmbedElement<K extends keyof HTMLElementTagNameMap>(
  parent: HTMLElement,
  tagName: K,
  options: EmbedElementOptions = {},
): HTMLElementTagNameMap[K] {
  const element = parent.createEl(tagName);
  if (options.className !== undefined) element.className = options.className;
  if (options.text !== undefined) element.textContent = options.text;
  for (const [name, value] of Object.entries(options.attributes ?? {})) {
    element.setAttribute(name, value);
  }
  return element;
}

export function mapSlideEmbedOpenError(
  messages: MessageTranslator,
  error: unknown,
): string {
  if (error instanceof PptxOpenError) {
    return messages.text(errorMessageKeys[error.category]);
  }
  return messages.text("embed.renderFailure");
}

/**
 * Host-agnostic single-slide embed rendering core.
 *
 * Owns shared chrome, scheduling, identity resolution, diagnostics, recovery
 * actions, cancellation, and disposal. Host adapters own MarkdownRenderChild /
 * CodeMirror widget lifecycle and visibility observation.
 */
export class SlideEmbedController<
  TFile extends SlideEmbedSourceFile = SlideEmbedSourceFile,
> {
  private canvas: HTMLElement | null = null;
  private unregisterTextCopy: (() => void) | undefined;
  private status: HTMLElement | null = null;
  private compatibility: HTMLElement | null = null;
  private input: SlideEmbedCoreInput<TFile> | null = null;
  private cancelScheduled: (() => void) | null = null;
  private rendererSession: PptxRendererSession | null = null;
  private generation = 0;
  private renderRequestedAt: number | null = null;
  private disposed = false;

  constructor(
    private readonly host: HTMLElement,
    private readonly ports: SlideEmbedCorePorts<TFile>,
  ) {}

  mount(input: SlideEmbedCoreInput<TFile>): void {
    if (this.disposed) return;
    this.unregisterTextCopy?.();
    this.unregisterTextCopy = undefined;
    this.input = input;
    this.host.replaceChildren();
    this.host.classList.add("pptx-slide-embed");
    this.host.dataset.sourcePath = input.sourcePath;
    this.host.dataset.slideId = String(input.target.slideId);
    this.host.dataset.createdSlide = String(input.target.createdSlideNumber);
    this.host.setAttribute("role", "group");

    if (input.file === null) {
      this.canvas = null;
      this.compatibility = null;
      this.host.dataset.state = "missing-source";
      this.host.setAttribute(
        "aria-label",
        this.ports.messages.text("embed.sourceMissing"),
      );
      this.status = appendEmbedElement(this.host, "div", {
        className: "pptx-slide-embed__status",
        text: this.ports.messages.text("embed.sourceMissing"),
        attributes: { role: "status", "aria-live": "polite" },
      });
      this.mountFooter(input.sourcePath, input.target, null);
      return;
    }

    this.host.dataset.state = "waiting";
    this.host.setAttribute(
      "aria-label",
      this.ports.messages.text("embed.loading"),
    );
    this.status = appendEmbedElement(this.host, "div", {
      className: "pptx-slide-embed__status",
      text: this.ports.messages.text("embed.loading"),
      attributes: { role: "status", "aria-live": "polite" },
    });
    this.canvas = appendEmbedElement(this.host, "div", {
      className: "pptx-slide-embed__canvas",
    });
    this.unregisterTextCopy = registerOfficeEmbedTextCopy(this.canvas);
    this.compatibility = appendEmbedElement(this.host, "div", {
      className: "pptx-slide-embed__compatibility",
      attributes: { role: "note" },
    });
    this.mountFooter(input.sourcePath, input.target, input.file);
  }

  setVisible(visible: boolean): void {
    if (this.disposed || this.input?.file === null) return;
    if (visible) this.startRendering();
    else this.releaseRenderingWindow();
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.unregisterTextCopy?.();
    this.unregisterTextCopy = undefined;
    this.releaseRenderingWindow();
  }

  private mountFooter(
    sourcePath: string,
    target: SlideReferenceTarget,
    file: TFile | null,
  ): void {
    const footer = appendEmbedElement(this.host, "div", {
      className: "pptx-slide-embed__footer",
    });
    const linkTarget = formatSlideReferenceLinkTarget(sourcePath, target);
    const sourceLink = appendEmbedElement(footer, "a", {
      className: "internal-link pptx-slide-embed__action",
      attributes: { href: linkTarget },
    });
    sourceLink.dataset.href = linkTarget;
    if (file === null) {
      sourceLink.textContent = this.ports.messages.text("reference.openPresentation");
    } else {
      const label = this.ports.messages.text("fileEmbed.openInTab");
      sourceLink.dataset.action = "open-in-tab";
      sourceLink.setAttribute("aria-label", label);
      sourceLink.title = label;
      decorateOfficeViewerIconButton(sourceLink, "panel-top");
    }
    const actionStatus = appendEmbedElement(footer, "span", {
      attributes: { role: "status", "aria-live": "polite" },
    });
    const { openSource, openExternally } = this.ports;
    if (openSource !== undefined) {
      sourceLink.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        actionStatus.textContent = "";
        try {
          await openSource(linkTarget);
        } catch {
          actionStatus.textContent = this.ports.messages.text("fileEmbed.openInTabFailure");
        }
      });
    }
    if (file !== null && openExternally !== undefined) {
      const externalButton = appendEmbedElement(footer, "button", {
        className: "pptx-slide-embed__action",
        text: this.ports.messages.text("external.open"),
        attributes: {
          type: "button",
          "data-action": "open-externally",
        },
      });
      externalButton.addEventListener("click", () => {
        actionStatus.textContent = "";
        void openExternally(file).catch(() => {
          actionStatus.textContent = this.ports.messages.text(
            "external.failure",
          );
        });
      });
    }
    footer.append(actionStatus);
  }

  private now(): number {
    return this.ports.now?.()
      ?? this.host.ownerDocument.defaultView?.performance.now()
      ?? performance.now();
  }

  private startRendering(): void {
    if (
      this.disposed
      || this.input?.file === null
      || this.input === null
      || this.cancelScheduled !== null
      || this.rendererSession !== null
      || this.canvas === null
      || this.status === null
    ) {
      return;
    }
    const generation = ++this.generation;
    const file = this.input.file;
    const target = this.input.target;
    this.renderRequestedAt = this.now();
    this.host.dataset.state = "queued";
    this.status.textContent = this.ports.messages.text("embed.loading");
    this.cancelScheduled = this.ports.scheduler.schedule(async (signal) => {
      if (generation !== this.generation || this.disposed) return;
      this.host.dataset.state = "loading";
      const startedAt = this.renderRequestedAt ?? this.now();
      try {
        const buffer = await this.ports.readBinary(file);
        signal.throwIfAborted();
        const session = await this.ports.renderer.open(
          buffer,
          this.canvas!,
          signal,
        );
        if (signal.aborted || generation !== this.generation || this.disposed) {
          session.dispose();
          return;
        }
        this.rendererSession = session;
        const slideIndex = session.slideIdentities?.indexOf(target.slideId) ?? -1;
        if (slideIndex < 0) {
          session.dispose();
          this.rendererSession = null;
          this.canvas!.replaceChildren();
          this.host.dataset.state = "stale-reference";
          this.status!.textContent = this.ports.messages.text(
            "reference.missing",
          );
          this.host.setAttribute(
            "aria-label",
            this.ports.messages.text("reference.missing"),
          );
          return;
        }
        await session.renderSlide(slideIndex);
        signal.throwIfAborted();
        if (generation !== this.generation || this.disposed) return;
        const currentSlide = slideIndex + 1;
        this.host.dataset.currentSlide = String(currentSlide);
        this.host.dataset.firstReadableMs = (
          this.now() - startedAt
        ).toFixed(3);
        this.canvas!.style.aspectRatio =
          `${session.slideWidth} / ${session.slideHeight}`;
        const label = this.ports.messages.text("embed.currentSlide", {
          name: file.basename,
          slide: currentSlide,
        });
        this.status!.textContent = label;
        this.host.setAttribute("aria-label", label);
        if (currentSlide !== target.createdSlideNumber) {
          this.status!.append(
            this.host.ownerDocument.createTextNode(
              ` — ${this.ports.messages.text("reference.moved", {
                created: target.createdSlideNumber,
                current: currentSlide,
              })}`,
            ),
          );
        }
        this.showCompatibilityWarnings(session);
        this.host.dataset.state = "ready";
      } catch (error) {
        if (signal.aborted || generation !== this.generation || this.disposed) {
          return;
        }
        this.disposeRendererSession();
        this.host.dataset.state = "error";
        const message = mapSlideEmbedOpenError(this.ports.messages, error);
        this.status!.textContent = message;
        this.host.setAttribute("aria-label", message);
      } finally {
        if (generation === this.generation) {
          this.cancelScheduled = null;
          this.renderRequestedAt = null;
        }
      }
    });
  }

  private showCompatibilityWarnings(session: PptxRendererSession): void {
    if (this.compatibility === null) return;
    this.compatibility.replaceChildren();
    try {
      if (!this.ports.showDiagnostics()) return;
    } catch {
      return;
    }
    let categories = session.compatibilityWarnings ?? [];
    try {
      categories = session.detectCompatibilityWarnings?.() ?? categories;
    } catch {
      // Optional compatibility inspection must not disrupt a readable slide.
    }
    for (const category of categories) {
      appendEmbedElement(this.compatibility, "div", {
        text: this.ports.messages.text(
          category === "font-substitution"
            ? "compatibility.fontSubstitution"
            : "compatibility.unsupportedContent",
        ),
      });
    }
  }

  private releaseRenderingWindow(): void {
    ++this.generation;
    this.cancelScheduled?.();
    this.cancelScheduled = null;
    this.renderRequestedAt = null;
    this.disposeRendererSession();
    this.canvas?.replaceChildren();
    delete this.host.dataset.currentSlide;
    delete this.host.dataset.firstReadableMs;
    this.compatibility?.replaceChildren();
    if (this.input?.file === null) return;
    this.host.dataset.state = "waiting";
    if (this.status !== null) {
      this.status.textContent = this.ports.messages.text("embed.loading");
    }
    this.host.setAttribute(
      "aria-label",
      this.ports.messages.text("embed.loading"),
    );
  }

  private disposeRendererSession(): void {
    this.rendererSession?.dispose();
    this.rendererSession = null;
  }
}
