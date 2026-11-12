---
status: accepted
---

# Limit the First DOCX Release to Reading and Search

The first public DOCX scope is the **Document reading workflow**: local,
read-only main-body reading, current-document search, and session-local result
navigation. Stable paragraph references, sourced paragraph copy, DOCX embeds,
persistent reading position, companion notes, and Vault-wide indexing are
excluded. Microsoft Word validation showed that native paragraph IDs cannot
satisfy the zero-false-binding reference contract, while reading and search do
not depend on that identity. The rejected reference design and its evidence
remain recorded in ADR-0004 and the 2026-07-31 exploration report.

The subsequent local development build adds a whole-file DOCX embed in notes.
It resolves the source file by its Vault link and does not depend on paragraph
IDs. This expands the preview surface without changing the first-release scope
or reopening the rejected paragraph reference contract.
