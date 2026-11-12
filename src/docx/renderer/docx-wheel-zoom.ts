export const MIN_DOCX_ZOOM = 0.5;
export const MAX_DOCX_ZOOM = 3;
const SCROLL_PIXELS_PER_ZOOM = 500;

export interface DocxZoom {
  readonly scale: number;
  setScale(scale: number): void;
  onChange(listener: () => void): () => void;
}

export function createDocxZoom(
  viewport: HTMLElement,
  findAnchor: (x: number, y: number, event?: WheelEvent) => HTMLElement,
  applyZoom: (scale: number) => void,
): DocxZoom & { dispose(): void } {
  let scale = 1;
  const listeners = new Set<() => void>();

  // Hold the pointer position for gestures and the viewport center for buttons.
  const setScale = (requestedScale: number, event?: WheelEvent) => {
    const nextScale = Math.min(MAX_DOCX_ZOOM, Math.max(MIN_DOCX_ZOOM, requestedScale));
    if (nextScale === scale) return;

    const bounds = viewport.getBoundingClientRect();
    const clientX = event?.clientX ?? bounds.left + viewport.clientWidth / 2;
    const clientY = event?.clientY ?? bounds.top + viewport.clientHeight / 2;
    const anchor = findAnchor(clientX, clientY, event);
    const before = anchor.getBoundingClientRect();
    scale = nextScale;
    applyZoom(scale);
    const after = anchor.getBoundingClientRect();
    if (before.width > 0) {
      const x = (clientX - before.left) / before.width;
      viewport.scrollLeft += after.left + x * after.width - clientX;
    }
    if (before.height > 0) {
      const y = (clientY - before.top) / before.height;
      viewport.scrollTop += after.top + y * after.height - clientY;
    }
    for (const listener of listeners) listener();
  };

  const onWheel = (event: WheelEvent) => {
    if ((!event.ctrlKey && !event.metaKey) || event.defaultPrevented || !event.cancelable) {
      return;
    }
    const pixelsPerUnit = [1, 16, viewport.clientHeight][event.deltaMode] ?? 1;
    const delta = event.deltaY * pixelsPerUnit;
    if (delta === 0) return;

    event.preventDefault();
    event.stopPropagation();
    setScale(scale * Math.exp(-delta / SCROLL_PIXELS_PER_ZOOM), event);
  };
  viewport.addEventListener("wheel", onWheel, { passive: false });
  return {
    get scale() { return scale; },
    setScale,
    onChange: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose: () => {
      viewport.removeEventListener("wheel", onWheel);
      listeners.clear();
    },
  };
}
