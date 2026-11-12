import { createSafeDocxRendererBuffer, type DocxSemanticModel } from "./docx-semantic-model";
import { prepareDocxPreviewPageTables } from "./renderer/docx-preview-page-layout";
import type { DocxRenderMode, DocxRendererAdapter } from "./renderer/docx-renderer-adapter";

/** Keep the inspected document available to both render modes for this open view. */
export class DocxRenderSource {
  private prepared = false;

  constructor(
    private buffer: ArrayBuffer,
    readonly model: DocxSemanticModel,
  ) {}

  async open(
    renderer: DocxRendererAdapter,
    container: HTMLElement,
    mode: DocxRenderMode,
    signal: AbortSignal,
  ) {
    if (mode === "full" && !this.prepared) {
      const buffer = await createSafeDocxRendererBuffer(
        this.buffer, signal, prepareDocxPreviewPageTables,
      );
      signal.throwIfAborted();
      this.buffer = buffer;
      this.prepared = true;
    }
    return renderer.open(this.buffer, container, this.model, signal, mode);
  }
}
