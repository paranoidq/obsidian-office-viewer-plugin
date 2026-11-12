# Obsidian Office Viewer

Obsidian Office Viewer provides trustworthy, local, read-only Office document
reading experiences inside Obsidian.

## Knowledge workflow

**Knowledge material**:
A local Office document that a reader uses as input to knowledge work and
needs to locate, reference, or reuse inside Obsidian without changing the
source document. Format coverage is valuable only when it advances this
workflow, not as a standalone measure of product breadth.
_Avoid_: Supported file, Office file

**Body-led DOCX knowledge material**:
DOCX **Knowledge material** whose reader value resides primarily in ordered
main-body prose, headings, lists, tables, and inline images rather than exact
print layout or auxiliary content stories. Reports, research material,
proposals, memoranda, and explanatory documents are the first target; forms,
mail-merge templates, brochures, and other layout-led documents are not.
_Avoid_: Word file, General DOCX, Print artifact

**Presentation companion note**:
The single presentation-level Markdown note a reader explicitly associates
with one PPTX **Knowledge material** for reader-authored understanding and
source-preserving references. It is not a separate note for each slide.
_Avoid_: Slide note, Per-slide note, PPTX copy

**Companion note path conflict**:
The state in which an associated **Presentation companion note** cannot follow
its PPTX to the required same-name location because that location is occupied.
The previously associated note remains authoritative until the conflict is
resolved.
_Avoid_: Duplicate companion note, Automatic reassociation

**Knowledge reference loop**:
The complete reader journey from finding relevant content, opening its precise
location, copying or embedding it with its source, and returning to that
location later. Reading and search may deliver useful format coverage without
claiming that this complete loop exists.
_Avoid_: Preview flow, Open-file flow

**Document reading workflow**:
The local, read-only journey of opening one DOCX, reading its main-body content,
searching within that document, and navigating to a matching paragraph during
the current session. It does not create Markdown references, persistent search
state, or a Vault-wide index.
_Avoid_: Knowledge reference loop, DOCX reference workflow, Word preview

**Knowledge reference loop technical exploration**:
A bounded engineering phase that tests the technical feasibility of supporting
the **Knowledge reference loop** for **Knowledge material**. It may finish while
post-release validation remains open, but it does not establish real-reader
workflow value or commit the project to ship a complete v0.2 product.
_Avoid_: v0.2 implementation, Workflow validation, M4 completion

**Real-reader workflow validation**:
Evidence that target readers repeatedly complete the **Knowledge reference
loop** with their own work Vaults and **Knowledge material** during real work.
Maintainer tests, generated fixtures, download counts, and one-time successful
opens do not establish it.
_Avoid_: Technical validation, Installation proof, Download validation

**Slide content search**:
A local search within one open PPTX **Knowledge material** that returns matching
**Source-authored slide text** as precise slide locations, allowing the reader to
continue through the **Knowledge reference loop**. **Speaker note content** is not
part of this search surface. The search does not imply a persistent or Vault-wide
content index; its query, indexed text, and results do not outlive the active
reading session.
_Avoid_: Vault search, File search

**Source-authored slide text**:
Reader-visible text directly authored in a slide's own titles, body content,
text boxes, shapes, or table cells. It excludes speaker notes, master or layout
text, chart or SmartArt data, and text contained in images.
_Avoid_: Extracted text, All slide text

**Source-authored document body text**:
Reader-visible text directly authored in the main body of DOCX **Knowledge
material**, including headings, ordinary paragraphs, list items, and paragraphs
inside table cells. Inserted revision content is included as final body text;
headers, footers, footnotes, endnotes, comments, text boxes, hidden text,
deleted revision content, and revision metadata are excluded.
_Avoid_: Extracted text, All document text, Word text

**Speaker note content**:
Reader-authored explanatory text associated with one slide but not visible on
the slide canvas. It is part of the presentation's **Knowledge material** when
a reader needs to locate, reference, or reuse it; notes-master text, headers,
footers, dates, and slide numbers are not speaker note content.
_Avoid_: Notes metadata, All notes text

**Presentation content search**:
A session-local search within one open PPTX across both **Source-authored slide
text** and **Speaker note content**. Results remain anchored to one slide and
identify which content surface matched so notes are never presented as visible
slide text. A single search is limited by its **Presentation search scope**.
_Avoid_: All slide text, Vault search

**Presentation search scope**:
The content surfaces included in one **Presentation content search**:
source-authored slide text only, speaker note content only, or both.
_Avoid_: Search filter, Search mode

**Document body search**:
A session-local search within one open DOCX **Knowledge material** that returns
matching **Source-authored document body text** as precise paragraph locations.
Its query, temporary index, and results do not outlive the active reading
session and never become a persistent or Vault-wide content index.
_Avoid_: Vault search, File search, Word search

**Document paragraph search result**:
One paragraph-level result from **Document body search**, anchored to that
paragraph and summarizing one or more matches within it. Multiple matches do
not create separate results for character positions.
_Avoid_: Text occurrence, Character match, Search hit

**Active document paragraph**:
The single searchable non-empty body paragraph a reader has selected in the
**DOCX reading view** by direct activation or search navigation. It is a
session-local reading location, not a character selection, deep link, viewport
guess, or persisted reading position.
_Avoid_: Selected text, Current page, Reading position

**Slide search result**:
One slide-level result from **Slide content search** or **Presentation content
search**, anchored to the slide's stable identity and summarizing one or more
matches within that slide. Multiple matches do not create separate results for
character positions or content surfaces.
_Avoid_: Text occurrence, Search hit

**Slide reference**:
A source-preserving Obsidian reference to one precise slide in a PPTX
**Knowledge material**. It follows the slide's stable identity when the source
presentation is reordered and distinguishes the creation-time ordinal from the
current ordinal. Following it returns the reader to that presentation and
slide, while a deleted target is reported as unavailable rather than silently
redirected to the same ordinal position.
_Avoid_: Page link, Deep link, PPTX link

**DOCX reading view**:
A local, read-only, continuously flowing representation of DOCX
**Knowledge material** that preserves its reader-meaningful content structure.
It does not promise Word pagination, print-layout fidelity, or pixel-level
formatting equivalence.
_Avoid_: Word preview, Page view, Print preview

**Unavailable document body content**:
Reader-visible content detected in the main body of DOCX **Knowledge material**
but not representable in the **DOCX reading view**. It retains an in-flow
placeholder and a visible document-level degradation notice instead of being
silently omitted; formatting differences and deliberately excluded non-body
content are not unavailable document body content.
_Avoid_: Omitted content, Hidden degradation, Formatting difference

**Document hyperlink**:
A source-authored hyperlink in **Source-authored document body text** whose
visible label remains part of its paragraph and whose `https`, `http`, or
`mailto` target opens only after explicit reader activation. Targets are never
prefetched or opened during rendering, search, or reference resolution; all
other external protocols and local paths are blocked.
_Avoid_: External resource, Embedded asset, Automatic navigation

**Document bookmark navigation**:
An explicit reader-activated jump from a source-authored internal DOCX link to
the unique main-body paragraph containing its bookmark target. It activates the
paragraph only for the current reading session and does not create a persistent
reference; excluded, missing, or ambiguous targets are unavailable.
_Avoid_: Persistent document reference, Persisted bookmark, External link

**Slide embed**:
An inline, read-only view of the current local source slide identified by a
**Slide reference** inside an Obsidian note. It reflects source changes when the
note is viewed again and preserves the source relationship instead of becoming
an unrelated screenshot or copied asset.
_Avoid_: Screenshot, Image attachment, Full-deck embed

**Live Preview slide embed**:
A **Slide embed** shown inline while its Markdown note remains editable in
Obsidian Live Preview. Touching it with a cursor or selection reveals the
canonical embed syntax; the rendered slide itself remains read-only and never
edits the source presentation.
_Avoid_: Editable slide, PPTX edit mode, Live Preview PPTX editor

**Whole-file Office embed**:
A local, read-only preview of a PPTX or DOCX inside an Obsidian note. A PPTX
embed can navigate the deck while a DOCX embed previews the body with the
source's recorded page layout. The embed can open the original file in the
system's default application and does not target a stable slide or paragraph.
_Avoid_: Slide embed, Paragraph reference

## Example dialogue

> **Developer:** What does the first DOCX release promise?
>
> **Domain expert:** A Document reading workflow: local main-body reading,
> current-document search, and session-local result navigation. It deliberately
> does not claim a stable DOCX reference or complete the Knowledge reference
> loop.

## Globalization

**Supported locale**:
A plugin interface locale whose user-facing text has been reviewed by a human
proficient in that language, whose catalog is complete against the English
source catalog, and whose critical reading flows pass the locale acceptance
matrix. The first supported locales are English, Simplified Chinese, and
Traditional Chinese; English is the fallback when the Obsidian locale has no
supported match.
_Avoid_: Available language, translated language

**Message locale**:
The supported locale selected from Obsidian's language when the plugin loads
and used for plugin-owned interface text. English variants resolve to `en`,
Simplified Chinese variants to `zh-Hans`, and Traditional Chinese variants to
one region-neutral `zh-Hant`; unsupported or invalid values resolve to `en`.
Changing the Obsidian language takes effect when Obsidian or the plugin next
loads.
_Avoid_: Language setting, display language

**User-facing message**:
Plugin-owned text that a reader can see in the interface or that assistive
technology can announce, including controls, settings, status and validation
messages, error surfaces, tooltips, and accessible labels. Developer logs,
internal exceptions, diagnostic categories, and development reports are not
user-facing messages.
_Avoid_: Log message, diagnostic detail
