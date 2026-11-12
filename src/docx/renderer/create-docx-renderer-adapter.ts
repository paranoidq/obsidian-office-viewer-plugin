import { BoundedDocxRendererAdapter } from "./bounded-docx-renderer-adapter";
import { DocxPreviewRendererAdapter } from "./docx-preview-renderer-adapter";

export interface CreateDocxRendererOptions {
  readonly unavailablePlaceholder?: string;
  readonly pageLayout?: boolean;
}

export function createDocxRendererAdapter(
  options: CreateDocxRendererOptions = {},
): BoundedDocxRendererAdapter {
  return new BoundedDocxRendererAdapter(
    new DocxPreviewRendererAdapter(options.pageLayout),
    {
      windowSize: 240,
      unavailablePlaceholder: options.unavailablePlaceholder,
    },
  );
}
