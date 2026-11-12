# Contributing

## Development

Use Node.js 22 and run:

```bash
npm ci
npm run verify
npm run test:e2e
npm run test:e2e:docx
npm run test:compatibility
npm run test:compatibility:docx
npm run test:performance:baseline
npm run test:performance:docx
npm run release:check
```

Keep renderer-specific objects behind their PPTX or DOCX adapter, read Office
file bytes only through the Vault API, and do not add network, editing,
telemetry, or source-write behavior. Any English user-facing message change
must update the Simplified and Traditional Chinese catalogs and retain the
human-review evidence described in `docs/globalization/multilingual-ui.md`.

## Feedback

Bugs and feature requests go through
[GitHub Issues](https://github.com/jerry4pan/obsidian-office-viewer-plugin/issues)
only. Do not open a pull request to report a bug or propose a feature; Issues
are the request surface. Security vulnerabilities use private reporting in
`SECURITY.md`.

### Bug reports

Before opening an issue, retry in the latest version and use **Copy diagnostic
summary** in the affected viewer. Include:

- exact Office Viewer and Obsidian versions;
- operating system;
- steps to reproduce;
- expected and actual behavior;
- the content-free diagnostic summary;
- whether **Open in default application** displays the file correctly.

Do not upload confidential presentations or Word documents, screenshots
containing sensitive content, filenames, paths, extracted text, or raw renderer
errors. If a minimal file is needed, create a repository-authored reproduction
with synthetic data.

### Feature requests

Describe the problem you are trying to solve, your reading workflow, and any
alternatives you considered. Check **Current boundaries** in the README first;
requests that ask for editing, cloud services, mobile support, or other
out-of-scope work are unlikely to be accepted.

### Pull requests

Open a pull request only after an Issue exists for the change, or when fixing
something already tracked. Follow the Development constraints above.
Compatibility changes must update the corpus expectation, visual evidence, and
known-limit documentation together. Runner-specific visual tolerances must be
bounded and documented separately from the approved reference baselines.
Performance misses are evidence and must not be hidden by loosening latency,
lifecycle, or resource budgets. Committed performance reports remain immutable
evidence snapshots; the sole current bundle-size allowance is the documented
5% regression budget in `docs/performance/README.md`.

## Local development package

Run `npm run build:dev` to create `artifacts/dev/office-viewer-dev/`. Copy that
directory to the test vault's `.obsidian/plugins/` folder. Disable the regular
`office-viewer` plugin in that vault before enabling `office-viewer-dev`;
both register the same Office file extensions.

On macOS, the development package uses LibreOffice for previews of Office
97–2003 `.ppt` and `.doc` files. The build uses
`/Applications/LibreOffice.app/Contents/MacOS/soffice` by default. Set
`OFFICE_VIEWER_LEGACY_SOFFICE` to the absolute path of another installed
`soffice` executable when building. Conversion uses a temporary profile
and the macOS sandbox. The output goes through the existing OOXML readers.
