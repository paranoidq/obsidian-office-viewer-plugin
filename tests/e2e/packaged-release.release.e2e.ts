import path from "node:path";
import { browser, expect } from "@wdio/globals";
import { obsidianPage } from "wdio-obsidian-service";
import {
  assertNoNetworkRequests,
  installNetworkGuard,
} from "../compatibility/browser-environment";
import {
  closeSettings,
  setDiagnosticSummaryEnabled,
} from "./office-viewer-settings";

async function sourceHash(vaultPath: string): Promise<string> {
  return browser.executeObsidian(async ({ app, obsidian, require }, path) => {
    const file = app.vault.getAbstractFileByPath(path);
    if (!(file instanceof obsidian.TFile)) throw new Error(`Missing ${path}`);
    const bytes = await app.vault.readBinary(file);
    return require("node:crypto").createHash("sha256")
      .update(new Uint8Array(bytes)).digest("hex");
  }, vaultPath);
}

describe("packaged release lifecycle", () => {
  it("installs in a clean Vault, rehearses an upgrade, and uninstalls cleanly", async () => {
    await installNetworkGuard();
    const pptxBefore = await sourceHash("minimal.pptx");
    const docxBefore = await sourceHash("read-search-only.docx");

    await obsidianPage.openFile("minimal.pptx");
    const installed = await browser.$('.pptx-viewer[data-state="ready"]');
    await expect(installed).toHaveText(expect.stringContaining("Obsidian PPTX smoke test"));
    await expect(installed.$('[data-action="copy-diagnostics"]')).not.toExist();
    expect(await sourceHash("minimal.pptx")).toBe(pptxBefore);

    await obsidianPage.openFile("read-search-only.docx");
    const docx = await browser.$(
      '.office-viewer-docx-shell[data-state="ready"]',
    );
    await expect(docx.$("section.docx")).toHaveText(
      expect.stringContaining("generated paragraph has no native identity"),
    );
    await docx.$('[data-action="open-docx-search"]').click();
    await docx
      .$('[data-action="docx-search-input"]')
      .setValue("generated paragraph");
    await expect(docx.$(".office-viewer-docx-search-result")).toExist();
    expect(await sourceHash("read-search-only.docx")).toBe(docxBefore);

    await setDiagnosticSummaryEnabled(true, "Diagnostic summary");
    await closeSettings();
    await obsidianPage.openFile("minimal.pptx");
    await expect(
      browser.$('.pptx-viewer[data-state="ready"] [data-action="copy-diagnostics"]'),
    ).toExist();
    expect(await sourceHash("minimal.pptx")).toBe(pptxBefore);

    await browser.executeObsidian(
      async ({ app, require }, stagedPlugin) => {
        const fs = require("node:fs/promises") as typeof import("node:fs/promises");
        const nodePath = require("node:path") as typeof import("node:path");
        const plugins = (app as unknown as {
          plugins: {
            disablePlugin(id: string): Promise<void>;
            enablePlugin(id: string): Promise<void>;
          };
        }).plugins;
        app.workspace.detachLeavesOfType("pptx-viewer");
        app.workspace.detachLeavesOfType("docx-viewer");
        await plugins.disablePlugin("office-viewer");
        const vaultRoot = (app.vault.adapter as unknown as {
          getBasePath(): string;
        }).getBasePath();
        const target = nodePath.join(vaultRoot, app.vault.configDir, "plugins", "office-viewer");
        for (const name of await fs.readdir(stagedPlugin)) {
          await fs.copyFile(nodePath.join(stagedPlugin, name), nodePath.join(target, name));
        }
        await plugins.enablePlugin("office-viewer");
      },
      path.resolve("artifacts/release/plugin"),
    );

    await obsidianPage.openFile("minimal.pptx");
    await expect(browser.$('.pptx-viewer[data-state="ready"]')).toHaveText(
      expect.stringContaining("Obsidian PPTX smoke test"),
    );
    expect(await sourceHash("minimal.pptx")).toBe(pptxBefore);

    await obsidianPage.openFile("read-search-only.docx");
    await expect(
      browser.$('.office-viewer-docx-shell[data-state="ready"]'),
    ).toHaveText(expect.stringContaining("generated paragraph has no native identity"));
    expect(await sourceHash("read-search-only.docx")).toBe(docxBefore);

    await obsidianPage.disablePlugin("office-viewer");
    const removed = await browser.executeObsidian(async ({ app, require }) => {
      const fs = require("node:fs/promises") as typeof import("node:fs/promises");
      const nodePath = require("node:path") as typeof import("node:path");
      const vaultRoot = (app.vault.adapter as unknown as {
        getBasePath(): string;
      }).getBasePath();
      const target = nodePath.join(vaultRoot, app.vault.configDir, "plugins", "office-viewer");
      await fs.rm(target, { recursive: true, force: true });
      return fs.access(target).then(() => false, () => true);
    });
    expect(removed).toBe(true);
    expect(await browser.$$(".pptx-viewer")).toHaveLength(0);
    expect(await browser.$$(".office-viewer-docx-shell")).toHaveLength(0);
    expect(await sourceHash("minimal.pptx")).toBe(pptxBefore);
    expect(await sourceHash("read-search-only.docx")).toBe(docxBefore);
    await assertNoNetworkRequests();
  });
});
