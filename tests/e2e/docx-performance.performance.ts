import { browser, expect } from "@wdio/globals";
import { obsidianPage } from "wdio-obsidian-service";
import {
  assertNoNetworkRequests,
  installNetworkGuard,
} from "../compatibility/browser-environment";

const REPRESENTATIVE =
  "docx-exploration/performance-representative-1000-paragraphs.docx";
const STRESS =
  "docx-exploration/performance-stress-5000-paragraphs.docx";

function p95(samples: readonly number[]): number {
  const sorted = [...samples].sort((left, right) => left - right);
  return sorted[Math.ceil(sorted.length * 0.95) - 1] ?? Number.POSITIVE_INFINITY;
}

async function closeActiveDocx(): Promise<number> {
  return browser.executeObsidian(async ({ app }) => {
    const started = performance.now();
    app.workspace.activeLeaf?.detach();
    while (document.querySelector(".office-viewer-docx-shell") !== null) {
      if (performance.now() - started > 5_000) break;
      await new Promise((resolve) => requestAnimationFrame(() => resolve(undefined)));
    }
    return performance.now() - started;
  });
}

async function rendererHeapUsed(): Promise<number> {
  return browser.executeObsidian(({ require }) => {
    const processModule = require("node:process") as typeof import("node:process");
    return processModule.memoryUsage().heapUsed;
  });
}

describe("installed DOCX performance gates", () => {
  it("meets representative p95 and bounded stress budgets", async () => {
    await installNetworkGuard();
    const firstReadable: number[] = [];
    const searchReady: number[] = [];
    const query: number[] = [];
    const cleanup: number[] = [];

    for (let sample = 0; sample < 5; sample += 1) {
      await obsidianPage.openFile(REPRESENTATIVE);
      const root = await browser.$(
        '.workspace-leaf.mod-active .office-viewer-docx-shell[data-state="ready"]',
      );
      await expect(root).toHaveAttribute("data-renderer", "docx-preview");
      firstReadable.push(Number(await root.getAttribute("data-first-readable-ms")));
      searchReady.push(Number(await root.getAttribute("data-search-ready-ms")));
      await root.$('[data-action="open-docx-search"]').click();
      query.push(await browser.execute(() => {
        const input = document.querySelector<HTMLInputElement>(
          ".workspace-leaf.mod-active [data-action=\"docx-search-input\"]",
        );
        if (input === null) throw new Error("DOCX search input is unavailable");
        const started = performance.now();
        input.value = "Representative paragraph";
        input.dispatchEvent(new Event("input", { bubbles: true }));
        if (
          document.querySelector(".workspace-leaf.mod-active .office-viewer-docx-search-result") === null
        ) {
          throw new Error("DOCX query did not synchronously produce a result");
        }
        return performance.now() - started;
      }));
      cleanup.push(await closeActiveDocx());
    }

    expect(p95(firstReadable)).toBeLessThanOrEqual(3_000);
    expect(p95(searchReady)).toBeLessThanOrEqual(3_000);
    expect(p95(query)).toBeLessThanOrEqual(100);
    expect(p95(cleanup)).toBeLessThanOrEqual(2_000);

    const memoryBefore = await rendererHeapUsed();
    await obsidianPage.openFile(STRESS);
    const stressRoot = await browser.$(
      '.workspace-leaf.mod-active .office-viewer-docx-shell[data-state="ready"]',
    );
    await stressRoot.$('[data-action="toggle-docx-mode"]').click();
    await expect(stressRoot).toHaveAttribute("data-renderer", "bounded-semantic");
    const stressElements = await browser.execute(() =>
      document.querySelectorAll(
        ".workspace-leaf.mod-active .office-viewer-docx-reading-body *",
      ).length);
    const memoryAfter = await rendererHeapUsed();
    const memoryDelta = Math.max(0, memoryAfter - memoryBefore);
    expect(stressElements).toBeLessThanOrEqual(1_200);
    expect(memoryDelta).toBeLessThanOrEqual(256 * 1024 * 1024);
    cleanup.push(await closeActiveDocx());

    process.stdout.write(`${JSON.stringify({
      environment: {
        obsidian: "1.12.7",
        installer: "1.12.7",
        renderer: "docx-preview@0.3.6",
      },
      samples: { firstReadable, searchReady, query, cleanup },
      p95: {
        firstReadableMs: p95(firstReadable),
        searchReadyMs: p95(searchReady),
        queryMs: p95(query),
        cleanupMs: p95(cleanup),
      },
      stress: { bodyDomElements: stressElements, memoryDeltaBytes: memoryDelta },
    })}\n`);
    await assertNoNetworkRequests();
  });
});
