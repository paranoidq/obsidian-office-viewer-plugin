import { BoundedDocxRendererAdapter } from "./bounded-docx-renderer-adapter";
import { DocxPreviewRendererAdapter } from "./docx-preview-renderer-adapter";
import type { DocxMessageTranslator } from "../docx-messages";

export interface CreateDocxRendererOptions {
  readonly unavailablePlaceholder?: string;
  readonly pageLayout?: boolean;
  readonly messages?: DocxMessageTranslator;
}

export function createDocxRendererAdapter(
  options: CreateDocxRendererOptions = {},
): BoundedDocxRendererAdapter {
  return new BoundedDocxRendererAdapter(
    new DocxPreviewRendererAdapter(options.pageLayout, options.messages),
    {
      windowSize: 240,
      unavailablePlaceholder: options.unavailablePlaceholder,
      messages: options.messages,
    },
  );
}
