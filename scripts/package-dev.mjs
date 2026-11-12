import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";

const root = new URL("../", import.meta.url);
const manifest = JSON.parse(await readFile(new URL("manifest.json", root), "utf8"));
manifest.id += "-dev";
manifest.name += "-dev";
const output = new URL(`artifacts/dev/${manifest.id}/`, root);
await mkdir(output, { recursive: true });
await Promise.all([
  copyFile(new URL("styles.css", root), new URL("styles.css", output)),
  writeFile(new URL("manifest.json", output), JSON.stringify(manifest, null, 2) + "\n"),
]);
console.log(`Development plugin: ${output.pathname}`);
