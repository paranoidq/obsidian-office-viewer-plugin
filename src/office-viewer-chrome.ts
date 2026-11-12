import { setIcon } from "obsidian";
import officeViewerMarkDataUrl from "../assets/brand/office-viewer-mark.svg";

export type OfficeViewerFormat = "DOCX" | "PPTX";

export const OFFICE_VIEWER_PRODUCT_NAME = "Office Viewer";
export const OFFICE_VIEWER_MARK_DATA_URL = officeViewerMarkDataUrl;

export interface OfficeViewerErrorAction {
  readonly label: string;
  readonly action: string;
  readonly onClick: () => void;
  readonly title?: string;
  readonly ariaLabel?: string;
}

export interface OfficeViewerErrorSurfaceOptions {
  readonly title: string;
  readonly safetyNote: string;
  readonly retry: OfficeViewerErrorAction;
  readonly openExternal?: OfficeViewerErrorAction;
  readonly extraActions?: readonly HTMLElement[];
  readonly classNames?: {
    readonly root?: string;
    readonly status?: string;
    readonly safetyNote?: string;
    readonly actions?: string;
    readonly actionStatus?: string;
  };
}

export interface OfficeViewerErrorSurface {
  readonly panel: HTMLElement;
  readonly actionStatus: HTMLElement;
}

export interface OfficeViewerToolbarOptions {
  readonly format: OfficeViewerFormat;
  readonly extraClassName?: string;
}

export interface OfficeViewerToolbar {
  readonly root: HTMLElement;
  readonly primary: HTMLElement;
  readonly secondary: HTMLElement;
}

/**
 * Shared Office Viewer recovery card: title, safety note, primary retry,
 * optional external open, and host-supplied extra actions.
 */
export function createOfficeViewerErrorSurface(
  options: OfficeViewerErrorSurfaceOptions,
): OfficeViewerErrorSurface {
  const panel = createDiv();
  panel.className = classNames(
    "office-viewer-error",
    options.classNames?.root,
  );

  const status = createDiv();
  status.className = classNames(
    "office-viewer-error__status",
    options.classNames?.status,
  );
  status.textContent = options.title;
  status.setAttribute("role", "status");

  const safety = createEl("p");
  safety.className = classNames(
    "office-viewer-error__safety-note",
    options.classNames?.safetyNote,
  );
  safety.textContent = options.safetyNote;

  const actions = createDiv();
  actions.className = classNames(
    "office-viewer-error__actions",
    options.classNames?.actions,
  );

  const retry = createActionButton(options.retry);
  actions.append(retry);

  for (const extra of options.extraActions ?? []) {
    actions.append(extra);
  }

  if (options.openExternal !== undefined) {
    actions.append(createActionButton(options.openExternal));
  }

  const actionStatus = createDiv();
  actionStatus.className = classNames(
    "office-viewer-error__action-status",
    options.classNames?.actionStatus,
  );
  actionStatus.setAttribute("role", "status");
  actionStatus.setAttribute("aria-live", "polite");

  panel.append(status, safety, actions, actionStatus);
  return { panel, actionStatus };
}

/**
 * Shared chrome row: brand identity, primary actions left, secondary actions
 * right. Callers own only the action slots.
 */
export function createOfficeViewerToolbar(
  options: OfficeViewerToolbarOptions,
): OfficeViewerToolbar {
  const root = createDiv();
  root.className = classNames(
    "office-viewer-toolbar",
    options.extraClassName,
  );
  const brand = createOfficeViewerBrand(options.format);
  const primary = createDiv();
  primary.className = "office-viewer-toolbar__primary";
  const secondary = createDiv();
  secondary.className = "office-viewer-toolbar__secondary";
  root.append(brand, primary, secondary);
  return { root, primary, secondary };
}

/** Apply an Obsidian icon and the shared toolbar hit target. */
export function decorateOfficeViewerIconButton(
  control: HTMLButtonElement | HTMLAnchorElement,
  iconId: string,
): void {
  control.classList.add("office-viewer-icon-button");
  control.textContent = "";
  setIcon(control, iconId);
}

function createOfficeViewerBrand(format: OfficeViewerFormat): HTMLElement {
  const brand = createDiv();
  brand.className = "office-viewer-brand";
  brand.dataset.officeViewerBrand = "true";
  brand.dataset.officeFormat = format;

  const mark = createEl("img");
  mark.className = "office-viewer-brand__mark";
  mark.src = OFFICE_VIEWER_MARK_DATA_URL;
  mark.alt = "";
  mark.setAttribute("aria-hidden", "true");
  mark.draggable = false;

  const product = createSpan();
  product.className = "office-viewer-brand__product";
  product.textContent = OFFICE_VIEWER_PRODUCT_NAME;

  const formatBadge = createSpan();
  formatBadge.className = "office-viewer-brand__format";
  formatBadge.textContent = format;

  brand.append(mark, product, formatBadge);
  return brand;
}

function createActionButton(
  action: OfficeViewerErrorAction,
): HTMLButtonElement {
  const button = createEl("button");
  button.type = "button";
  button.textContent = action.label;
  button.setAttribute("data-action", action.action);
  if (action.title !== undefined) button.title = action.title;
  if (action.ariaLabel !== undefined) {
    button.setAttribute("aria-label", action.ariaLabel);
  }
  button.addEventListener("click", action.onClick);
  return button;
}

function classNames(
  ...parts: Array<string | undefined>
): string {
  return parts.filter((part): part is string =>
    part !== undefined && part.length > 0
  ).join(" ");
}
