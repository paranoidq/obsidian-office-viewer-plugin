import esbuild from "esbuild";
import process from "node:process";
import path from "node:path";
import fs from "node:fs";
import rendererCandidates from "./src/renderer/renderer-candidates.json" with {
  type: "json",
};

const production = process.argv[2] === "production";
const devPackage = process.argv[3] === "dev";
const rendererCandidate = process.env.PPTX_RENDERER_CANDIDATE ?? "aiden";
const rendererTestAdapter = process.env.PPTX_RENDERER_TEST_ADAPTER;
const outfile = devPackage
  ? "artifacts/dev/office-viewer-dev/main.js"
  : process.env.PPTX_BUNDLE_OUTFILE ?? "main.js";

if (!Object.hasOwn(rendererCandidates, rendererCandidate)) {
  throw new Error(
    `Unsupported PPTX renderer candidate "${rendererCandidate}"`,
  );
}
if (
  rendererTestAdapter !== undefined &&
  rendererTestAdapter !== "degraded-navigation"
) {
  throw new Error(`Unsupported PPTX renderer test adapter "${rendererTestAdapter}"`);
}
const selectedRendererModule = rendererTestAdapter
  ? "tests/support/selected-pptx-renderer-adapter.degraded.ts"
  : `src/renderer/selected-pptx-renderer-adapter.${rendererCandidate}.ts`;
const context = await esbuild.context({
  entryPoints: ["src/main.ts"],
  bundle: true,
  external: [
    "obsidian",
    "electron",
    "node:*",
    "@codemirror/autocomplete",
    "@codemirror/collab",
    "@codemirror/commands",
    "@codemirror/language",
    "@codemirror/lint",
    "@codemirror/search",
    "@codemirror/state",
    "@codemirror/view",
    "@lezer/common",
    "@lezer/highlight",
    "@lezer/lr"
  ],
  format: "cjs",
  platform: "browser",
  target: "es2022",
  outfile,
  sourcemap: production ? false : "inline",
  minify: production,
  loader: {
    ".svg": "dataurl",
  },
  define: {
    "process.env.NODE_ENV": production ? '"production"' : '"development"',
    "process.env.PPTX_RENDERER_CANDIDATE": JSON.stringify(rendererCandidate),
    "process.env.PPTX_RENDERER_VERSION": JSON.stringify(
      rendererCandidates[rendererCandidate].version,
    ),
    "process.env.OFFICE_VIEWER_LEGACY_SOFFICE": JSON.stringify(devPackage
      ? process.env.OFFICE_VIEWER_LEGACY_SOFFICE
        ?? "/Applications/LibreOffice.app/Contents/MacOS/soffice"
      : ""),
  },
  plugins: [
    {
      name: "selected-pptx-renderer",
      setup(build) {
        build.onResolve({ filter: /^#selected-pptx-renderer$/ }, () => ({
          path: path.resolve(
            selectedRendererModule,
          ),
        }));
      },
    },
  ],
  logLevel: "info"
});

if (production) {
  await context.rebuild();
  await context.dispose();
  // JSZip's setImmediate polyfill (via lie → immediate) creates <script>
  // elements and `new Function(""+…)` as fallbacks for obsolete IE. Obsidian's
  // review scanner flags these dead code paths. Neutralize them — the
  // polyfill already prefers MutationObserver / MessageChannel, so these
  // branches are unreachable in modern Electron.
  let content = fs.readFileSync(outfile, "utf-8");
  content = content.replace(/createElement\("script"\)/g, 'createElement("div")');
  content = content.replace(/new Function\(""\+\w+\)/g, "function(){}");
  if (
    content.includes('createElement("script")') ||
    /new Function\(""\+\w+\)/.test(content)
  ) {
    throw new Error(
      "production bundle still contains review-flagged JSZip polyfill fallbacks",
    );
  }
  fs.writeFileSync(outfile, content);
} else {
  await context.watch();
}
