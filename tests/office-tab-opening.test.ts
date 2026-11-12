import { TFile, type App } from "obsidian";
import { expect, it, vi } from "vitest";
import { registerOfficeTabOpening } from "../src/office-tab-opening";

it("opens Office clicks in separate tabs and passes slide references to Obsidian", async () => {
  const file = Object.assign(new TFile(), { path: "Deck.pptx", extension: "pptx" });
  const note = Object.assign(new TFile(), { path: "Note.md", extension: "md" });
  const files = new Map([[file.path, file], [note.path, note]]);
  const openFile = vi.fn(async () => {});
  const nativeOpenLink = vi.fn(async () => {});
  const workspace = {
    openLinkText: nativeOpenLink,
    getLeaf: vi.fn(() => ({ openFile })),
    iterateAllLeaves: vi.fn(),
    on: vi.fn(),
    offref: vi.fn(),
  };
  const unregister = registerOfficeTabOpening({
    workspace,
    metadataCache: { getFirstLinkpathDest: (path: string) => files.get(path) ?? null },
    vault: { getAbstractFileByPath: (path: string) => files.get(path) ?? null },
  } as unknown as App, (candidate) => candidate.extension === "pptx");
  const row = document.createElement("div");
  row.className = "nav-file-title";
  row.dataset.path = file.path;
  document.body.append(row);
  try {
    row.click();
    expect(workspace.getLeaf).toHaveBeenCalledExactlyOnceWith("tab");
    expect(openFile).toHaveBeenCalledExactlyOnceWith(file, { active: true });
    row.dataset.path = note.path;
    row.click();
    expect(openFile).toHaveBeenCalledTimes(1);

    const openLinkText = (workspace as unknown as App["workspace"]).openLinkText;
    const reference = "Deck.pptx#slide-id=256&slide=1";
    await openLinkText.call(workspace, reference, "Note.md", false);
    await openLinkText.call(workspace, "Note.md", "Index.md", false);
    await openLinkText.call(workspace, reference, "Note.md", "split");
    expect(nativeOpenLink.mock.calls).toEqual([
      [reference, "Note.md", "tab", undefined],
      ["Note.md", "Index.md", false, undefined],
      [reference, "Note.md", "split", undefined],
    ]);
  } finally {
    unregister();
    row.remove();
  }
  expect(workspace.openLinkText).toBe(nativeOpenLink);
});
