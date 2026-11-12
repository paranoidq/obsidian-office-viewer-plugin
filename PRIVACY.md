# Privacy

Office Viewer processes presentation and Word document files locally inside
desktop Obsidian.
The normal viewing path does not upload files, call a cloud renderer, follow
external media relationships, or send telemetry or analytics.

The plugin reads Office source bytes through Obsidian's Vault API and never
writes back to the source. The macOS development build can convert binary
`.ppt` and `.doc` using an installed LibreOffice. It uses a temporary copy and
profile in a process with network access blocked, then deletes the temporary
files. Optional PPTX reading-position history
stores only a Vault-relative path, file size, modification time, zero-based
slide index, and update timestamp in the plugin data store. It can be disabled
and cleared from settings.

Opening a DOCX attachment in its default application writes a temporary copy
outside the Vault. The plugin removes these copies on unload.

DOCX preview may rasterize Windows metafiles and chart cached series into local
PNG images for display. That conversion stays in-process, uses only package
bytes already present in the open document, and does not fetch external chart
workbooks or media.

Explicitly claimed **Presentation companion notes** are ordinary Vault Markdown
files. The plugin creates or adopts them only when the reader invokes **Open
companion note**. Plugin data stores only the normalized Vault-relative
`sourcePath` and `notePath` for each claimed pair. Disabling reading-position
history does not clear companion relationships. Companion note bodies are never
copied into plugin data, diagnostics, or telemetry.

**Diagnostic summary** is off by default. When you enable it, detectable
compatibility warnings and the copy control appear on the next open, retry, or
reload of a file. Enabling the setting does not upload data or submit a report
automatically.

The user-triggered diagnostic summary contains plugin, Obsidian, renderer, and
operating-system versions; source byte size; slide count when known; timings;
stable warning/error categories; and anonymous feature flags. It excludes
filenames, paths, presentation text, images, author metadata, URLs, raw errors,
and rendered content. Copying the summary writes it only to the local clipboard.

Compatibility-check results exist only for the current view lifecycle and are
not persisted as a document archive.

Presentation content search and speaker-note viewing are local to one open
presentation and one view lifecycle. The query, source-authored slide text and
speaker-note paragraphs read from package XML, normalized comparable text,
snippets, highlights, scope choice, and results are not written to plugin data,
diagnostics, logs, Markdown, or the source presentation. Search performance
evidence contains only elapsed time and mounted-result counts.

DOCX body search is likewise limited to one open document and one view
lifecycle. The query, extracted body paragraphs, snippets, match counts, and
active result are not persisted in plugin data, Markdown, diagnostics, or the
source document. DOCX support does not create paragraph references,
companion notes, or a Vault-wide index.

Safe `http`, `https`, and `mailto` document links are opened only after an
explicit click. They are never prefetched. Other external relationships are
blocked or represented as unavailable content without a network request.

**Open in default application** is an explicit action. After it is selected,
the operating system and chosen application control any subsequent processing
or network behavior.
