# Security

## Supported release

Security fixes target the latest published Office Viewer release on GitHub and
in Obsidian Community Plugins.

## Security model

- PPTX and DOCX packages are inspected before renderer allocation.
- ZIP entry, expanded-size, media, and XML limits bound resource use.
- OOXML checks reject or remove macros, VBA, ActiveX, embedded executable
  content, and external media before rendering.
- In the macOS development build, legacy conversion uses a temporary
  LibreOffice profile that disables macros and active embedded content.
  The converter runs with network access blocked and writes restricted to
  its temporary directory. The converted OOXML still passes the existing
  package checks.
- External hyperlinks are not followed automatically; DOCX links are limited
  to `http`, `https`, and `mailto` after an explicit click.
- DOCX attachment buttons extract supported document data to a private
  temporary directory after an explicit open action. The viewer checks the
  extracted format and passes a generated local file path to the default app.
  Linked OLE objects and embedded executable commands are excluded.
- Source files are read-only and are never modified by the viewer.
- Rendering is isolated per Obsidian view and cancellation/disposal is scoped
  to the open generation.
- Renderer exceptions are mapped to stable categories before user display or
  diagnostic export.
- Non-blocking compatibility warnings can be disabled by turning off
  **Diagnostic summary**. Security rejections for macros, ActiveX, external
  content, resource limits, and other blocking failures are not controlled by
  that setting.

Office Viewer is a previewer, not a malware scanner or a complete PowerPoint or
Word security boundary. Do not treat a successful preview as proof that a file
is safe to open in another application.

## Reporting a vulnerability

Do not attach confidential presentations or Word documents, or paste document
content into a public issue. Use [GitHub's private vulnerability reporting](
https://github.com/jerry4pan/obsidian-office-viewer-plugin/security/advisories/new)
for this repository. Include the Office Viewer version, Obsidian version,
operating system, stable error category, reproducible steps using a
non-confidential fixture, and the content-free diagnostic summary if relevant.
