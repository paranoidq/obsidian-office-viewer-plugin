import { StateField } from "@codemirror/state";
import { vi } from "vitest";

export const getLanguage = vi.fn(() => "en");
export const apiVersion = "1.13.1";

export function parseLinktext(linktext: string): { path: string; subpath: string } {
  const fragment = linktext.indexOf("#");
  return fragment < 0
    ? { path: linktext, subpath: "" }
    : { path: linktext.slice(0, fragment), subpath: linktext.slice(fragment) };
}

/** Stand-in for Obsidian's lucide icon injector. */
export function setIcon(parent: HTMLElement, iconId: string): void {
  parent.replaceChildren();
  parent.dataset.icon = iconId;
}

export class Plugin {
  app: unknown;
  registerView = vi.fn();
  registerExtensions = vi.fn();
  registerEvent = vi.fn();
  registerMarkdownPostProcessor = vi.fn();
  registerEditorExtension = vi.fn();
  addSettingTab = vi.fn();
  loadData = vi.fn(async (): Promise<unknown> => undefined);
  saveData = vi.fn(async (_data: unknown): Promise<void> => undefined);

  constructor(app: unknown) {
    this.app = app;
  }
}

export class PluginSettingTab {
  containerEl = document.createElement("div") as HTMLElement & {
    empty(): void;
  };

  constructor(
    public app: unknown,
    public plugin: Plugin,
  ) {
    this.containerEl.empty = () => this.containerEl.replaceChildren();
  }
}

export class ToggleComponent {
  private value = false;
  private change: ((value: boolean) => unknown) | undefined;
  readonly toggleEl = document.createElement("label");

  constructor() {
    const input = document.createElement("input");
    input.type = "checkbox";
    this.toggleEl.append(input);
  }

  setValue(value: boolean): this {
    this.value = value;
    return this;
  }

  onChange(change: (value: boolean) => unknown): this {
    this.change = change;
    return this;
  }

  async trigger(value: boolean): Promise<void> {
    this.value = value;
    await this.change?.(value);
  }

  triggerWithoutAwait(value: boolean): void {
    this.value = value;
    this.change?.(value);
  }

  getValue(): boolean {
    return this.value;
  }
}

export class Setting {
  readonly toggle = new ToggleComponent();
  readonly settingEl = document.createElement("div");
  readonly nameEl = document.createElement("div");
  readonly descEl = document.createElement("div");

  constructor(readonly containerEl: HTMLElement) {
    this.settingEl.className = "setting-item";
    this.settingEl.append(this.nameEl, this.descEl);
    containerEl.append(this.settingEl);
  }

  setName(name: string): this {
    this.nameEl.textContent = name;
    return this;
  }

  setDesc(description: string): this {
    this.descEl.textContent = description;
    return this;
  }

  addToggle(configure: (toggle: ToggleComponent) => unknown): this {
    configure(this.toggle);
    this.settingEl.append(this.toggle.toggleEl);
    Object.defineProperty(this.settingEl, "testToggle", {
      value: this.toggle,
    });
    return this;
  }
}

export class TFile {
  constructor(
    public path = "",
    public stat: { size: number; mtime: number } = { size: 0, mtime: 0 },
    public basename = path.replace(/^.*\//, "").replace(/\.pptx$/i, ""),
    public extension = "pptx",
  ) {}
}

export class FileView {
  app: unknown;
  contentEl = document.createElement("div");
  file: { basename?: string } | null = null;
  scope: Scope | null = null;

  constructor(leaf: { app: unknown }) {
    this.app = leaf.app;
  }

  setEphemeralState(_state: unknown): void {}
}

export class Scope {
  readonly handlers: Array<{
    modifiers: string[] | null;
    key: string | null;
    func: (event: KeyboardEvent, context: unknown) => unknown;
  }> = [];

  constructor(readonly parent?: unknown) {}

  register(
    modifiers: string[] | null,
    key: string | null,
    func: (event: KeyboardEvent, context: unknown) => unknown,
  ): { scope: Scope; modifiers: string | null; key: string | null } {
    this.handlers.push({ modifiers, key, func });
    return {
      scope: this,
      modifiers: modifiers?.join("+") ?? null,
      key,
    };
  }
}

export class MarkdownRenderChild {
  constructor(public containerEl: HTMLElement) {}
  onload(): void {}
  onunload(): void {}
  load(): void {
    this.onload();
  }
  unload(): void {
    this.onunload();
  }
}

/** Controllable stand-in for Obsidian's public Live Preview state field. */
export const editorLivePreviewField = StateField.define<boolean>({
  create: () => true,
  update: (value) => value,
});

/** Controllable stand-in for Obsidian's public editor info field. */
export const editorInfoField = StateField.define<{ file: { path: string } | null }>({
  create: () => ({ file: { path: "note.md" } }),
  update: (value) => value,
});
