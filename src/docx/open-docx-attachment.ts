import { createHash, randomUUID } from "node:crypto";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { openLocalPath } from "../external-open";
import type { DocxAttachmentFile } from "./docx-attachments";

export function createDocxAttachmentOpener() {
  const copies = new Map<string, Promise<string>>();
  let disposed = false;

  const remove = (directory: string) =>
    rm(directory, { recursive: true, force: true }).catch(() => undefined);

  async function createCopy(file: DocxAttachmentFile): Promise<string> {
    const directory = await mkdtemp(join(tmpdir(), "office-viewer-attachment-"));
    try {
      const base = (file.name.split(/[\\/]/).at(-1) ?? "")
        .replace(/\.[^.]*$/, "")
        .replace(/[\p{C}<>:"/\\|?*]/gu, "_")
        .replace(/^[.\s]+|[.\s]+$/g, "");
      const stem = Array.from(base).slice(0, 40).join("") || "attachment";
      const safeStem = /^(?:con|prn|aux|nul|(?:com|lpt)[1-9¹²³])(?:\.|$)/i.test(stem)
        ? `attachment-${stem}` : stem;
      // Spreadsheet apps can reject identical basenames even in separate directories.
      const suffix = randomUUID().replaceAll("-", "").slice(0, 12);
      const path = join(directory, `${safeStem}--${suffix}.${file.extension}`);
      await writeFile(path, file.bytes, { mode: 0o600, flag: "wx" });
      return path;
    } catch (error) {
      await remove(directory);
      throw error;
    }
  }

  return {
    async open(file: DocxAttachmentFile): Promise<void> {
      if (disposed) throw new Error("The Office Viewer was unloaded");
      const digest = createHash("sha256").update(file.bytes).digest("hex");
      const key = JSON.stringify([file.name, file.extension, digest]);
      let copy = copies.get(key);
      if (copy === undefined) {
        copy = createCopy(file).catch((error: unknown) => {
          copies.delete(key);
          throw error;
        });
        copies.set(key, copy);
      }
      const path = await copy;
      if (disposed) throw new Error("The Office Viewer was unloaded");
      await openLocalPath(path);
    },
    async dispose(): Promise<void> {
      disposed = true;
      const pending = [...copies.values()];
      copies.clear();
      await Promise.allSettled(pending.map(async (copy) => remove(dirname(await copy))));
    },
  };
}
