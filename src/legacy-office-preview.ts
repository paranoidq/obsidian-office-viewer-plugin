import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, readFile, realpath, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { isAbsolute, join } from "node:path";
import { pathToFileURL } from "node:url";
import { DocxOpenError } from "./docx/docx-open-error";
import { PptxOpenError } from "./pptx-open-error";

export type LegacyOfficeExtension = "ppt" | "doc";

export class LegacyOfficeUnavailable extends Error {}

const MAX_SOURCE_BYTES = 64 * 1024 * 1024;
const MAX_OUTPUT_BYTES = 128 * 1024 * 1024;
const COMPOUND_FILE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const CONVERSION_TIMEOUT_MS = 30_000;

const PROFILE = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
  <item oor:path="/org.openoffice.Office.Common/Security/Scripting">
    <prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop>
    <prop oor:name="DisableActiveContent" oor:op="fuse"><value>true</value></prop>
    <prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop>
  </item>
</oor:items>`;

const SANDBOX = `(version 1)
(allow default)
(deny network*)
(deny file-write* (require-not (subpath (param "WORK_DIR"))))
(deny file-read-data (subpath (param "VAULT_DIR")))`;

export function createLegacyOfficePreview(executable: string, vaultPath: string) {
  if (process.platform !== "darwin" || !isAbsolute(executable)
    || !isAbsolute(vaultPath) || !existsSync(executable)) {
    return undefined;
  }

  return async (
    extension: LegacyOfficeExtension,
    sourceSize: number,
    readSource: () => Promise<ArrayBuffer>,
    signal: AbortSignal,
  ): Promise<ArrayBuffer> => {
    const OpenError = extension === "ppt" ? PptxOpenError : DocxOpenError;
    signal.throwIfAborted();
    if (sourceSize > MAX_SOURCE_BYTES) {
      throw new OpenError("resource-exhausted", "Legacy Office source exceeds the conversion limit");
    }
    const source = new Uint8Array(await readSource());
    signal.throwIfAborted();
    if (source.byteLength > MAX_SOURCE_BYTES) {
      throw new OpenError("resource-exhausted", "Legacy Office source exceeds the conversion limit");
    }
    if (!COMPOUND_FILE_SIGNATURE.every((byte, index) => source[index] === byte)) {
      throw new OpenError("malformed", "Expected a binary Office compound document");
    }

    const root = await realpath(tmpdir());
    const vault = await realpath(vaultPath);
    const work = await mkdtemp(join(root, "office-viewer-legacy-"));
    try {
      const profile = join(work, "profile");
      const userProfile = join(profile, "user");
      const input = join(work, `source.${extension}`);
      const outputExtension = extension === "ppt" ? "pptx" : "docx";
      const inputFilter = extension === "ppt" ? "MS PowerPoint 97" : "MS Word 97";
      const outputFilter = extension === "ppt"
        ? "Impress MS PowerPoint 2007 XML"
        : "Office Open XML Text";
      await mkdir(userProfile, { recursive: true });
      await writeFile(join(userProfile, "registrymodifications.xcu"), PROFILE, { signal });
      await writeFile(input, source, { signal });
      signal.throwIfAborted();
      await runConverter(executable, vault, work, [
        `-env:UserInstallation=${pathToFileURL(profile).href}`,
        "--headless", "--norestore", `--infilter=${inputFilter}`,
        "--convert-to", `${outputExtension}:${outputFilter}`,
        "--outdir", work, input,
      ], signal);
      signal.throwIfAborted();
      const output = join(work, `source.${outputExtension}`);
      const outputInfo = await stat(output);
      if (outputInfo.size > MAX_OUTPUT_BYTES) {
        throw new OpenError("resource-exhausted", "Converted Office file exceeds the preview limit");
      }
      const bytes = await readFile(output, { signal });
      return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    } catch (error) {
      signal.throwIfAborted();
      if (error instanceof PptxOpenError || error instanceof DocxOpenError) throw error;
      throw new OpenError("incompatible", "Legacy Office conversion did not produce a preview", { cause: error });
    } finally {
      await rm(work, { recursive: true, force: true });
    }
  };
}

function runConverter(
  executable: string,
  vault: string,
  work: string,
  args: readonly string[],
  signal: AbortSignal,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn("/usr/bin/sandbox-exec", [
      "-D", `WORK_DIR=${work}`, "-D", `VAULT_DIR=${vault}`, "-p", SANDBOX,
      executable, ...args,
    ], {
      cwd: work,
      env: { ...process.env, TMPDIR: work, XDG_CACHE_HOME: join(work, "cache") },
      stdio: "ignore",
      signal,
      killSignal: "SIGKILL",
    });
    const timeout = setTimeout(() => child.kill("SIGKILL"), CONVERSION_TIMEOUT_MS);
    let spawnError: Error | undefined;
    child.once("error", (error) => { spawnError = error; });
    child.once("close", (code) => {
      clearTimeout(timeout);
      if (code === 0 && spawnError === undefined) resolve();
      else reject(spawnError ?? new Error("LibreOffice conversion process failed"));
    });
  });
}
