import { decorateOfficeViewerIconButton } from "../office-viewer-chrome";
import type { DocxMessageTranslator } from "./docx-messages";
import { MAX_DOCX_ZOOM, MIN_DOCX_ZOOM, type DocxZoom } from "./renderer/docx-wheel-zoom";

export function createDocxZoomControls(document: Document, messages: DocxMessageTranslator) {
  const root = document.createElement("div");
  root.className = "office-viewer-docx-zoom";
  root.setAttribute("role", "group");
  root.setAttribute("aria-label", messages.text("zoom"));
  const percent = new Intl.NumberFormat(messages.locale, {
    style: "percent",
    maximumFractionDigits: 0,
  });
  const zoomOut = document.createElement("button");
  const reset = document.createElement("button");
  const zoomIn = document.createElement("button");
  for (const button of [zoomOut, reset, zoomIn]) button.type = "button";
  zoomOut.dataset.action = "docx-zoom-out";
  reset.dataset.action = "docx-zoom-reset";
  zoomIn.dataset.action = "docx-zoom-in";
  reset.className = "office-viewer-docx-zoom__ratio";
  decorateOfficeViewerIconButton(zoomOut, "lucide-minus");
  decorateOfficeViewerIconButton(zoomIn, "lucide-plus");
  zoomOut.title = messages.text("zoomOut");
  zoomIn.title = messages.text("zoomIn");
  reset.title = messages.text("zoomReset");
  zoomOut.setAttribute("aria-label", zoomOut.title);
  zoomIn.setAttribute("aria-label", zoomIn.title);
  root.append(zoomOut, reset, zoomIn);

  let zoom: DocxZoom | undefined;
  let unsubscribe: (() => void) | undefined;
  const refresh = () => {
    const scale = zoom?.scale ?? 1;
    reset.textContent = percent.format(scale);
    reset.setAttribute("aria-label", messages.text("zoomResetAt", { percent: reset.textContent }));
    zoomOut.disabled = zoom === undefined || scale <= MIN_DOCX_ZOOM;
    zoomIn.disabled = zoom === undefined || scale >= MAX_DOCX_ZOOM;
    reset.disabled = zoom === undefined;
  };
  const step = (points: number) => {
    if (zoom !== undefined) zoom.setScale((Math.round(zoom.scale * 100) + points) / 100);
  };
  zoomOut.addEventListener("click", (event) => {
    event.stopPropagation();
    step(-10);
  });
  zoomIn.addEventListener("click", (event) => {
    event.stopPropagation();
    step(10);
  });
  reset.addEventListener("click", (event) => {
    event.stopPropagation();
    zoom?.setScale(1);
  });
  refresh();

  return {
    root,
    setZoom(next?: DocxZoom) {
      const scale = zoom?.scale ?? 1;
      unsubscribe?.();
      zoom = next;
      zoom?.setScale(scale);
      unsubscribe = zoom?.onChange(refresh);
      refresh();
    },
  };
}
