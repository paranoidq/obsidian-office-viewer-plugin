import { readFile } from "node:fs/promises";
import { basename } from "node:path";
import { afterEach, expect, it, vi } from "vitest";
import type { DocxAttachmentFile } from "../../src/docx/docx-attachments";
import { createDocxAttachmentOpener } from "../../src/docx/open-docx-attachment";
import { openLocalPath } from "../../src/external-open";

vi.mock("../../src/external-open", () => ({ openLocalPath: vi.fn() }));

const openLocal = vi.mocked(openLocalPath);
const file: DocxAttachmentFile = {
  name: "业务申请表.xlsx",
  extension: "xlsx",
  bytes: new Uint8Array([1, 2, 3]),
};

afterEach(() => openLocal.mockReset());

it("reopens the same extracted path across clicks and concurrent requests", async () => {
  const opener = createDocxAttachmentOpener();
  let path: string;
  try {
    await Promise.all([opener.open(file), opener.open({ ...file, bytes: file.bytes.slice() })]);
    await opener.open(file);
    path = openLocal.mock.calls[0]![0];
    expect(openLocal.mock.calls).toEqual([[path], [path], [path]]);
    expect(await readFile(path)).toEqual(Buffer.from(file.bytes));
  } finally {
    await opener.dispose();
  }
  await expect(readFile(path)).rejects.toMatchObject({ code: "ENOENT" });
});

it("uses distinct filenames for different workbooks and after a plugin reload", async () => {
  const opener = createDocxAttachmentOpener();
  const reloaded = createDocxAttachmentOpener();
  try {
    await opener.open(file);
    await opener.open({ ...file, bytes: new Uint8Array([3, 2, 1]) });
    await opener.dispose();
    await reloaded.open(file);
    const names = openLocal.mock.calls.map(([path]) => basename(path));
    expect(new Set(names).size).toBe(3);
    for (const name of names) expect(name).toMatch(/^业务申请表--[0-9a-f]{12}\.xlsx$/);
  } finally {
    await Promise.all([opener.dispose(), reloaded.dispose()]);
  }
});
