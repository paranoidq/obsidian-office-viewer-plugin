# ADR-0001: Select Aiden PPTX Renderer for M0

- Status: Accepted
- Date: 2026-07-14
- Decision owners: project maintainers
- Scope: local, read-only PPTX preview in Obsidian M0/M1

## Context

The plugin needs a browser-native PPTX renderer that works from an Obsidian
Vault `ArrayBuffer`, keeps document content local, renders one slide at a time,
and can be stopped and cleaned up when a leaf closes or a new file replaces an
in-flight open. M0 requires at least 80% readable main content on the fixed
compatibility corpus and a first-readable p95 no greater than 3 seconds in the
installed Obsidian protocol.

Two exact package versions were run through the same product UI, shared
`PptxRendererAdapter` contract and preflight decorator, malformed-file
fixtures, compatibility corpus, visual baselines, performance protocol, and
bundle build. Each package has its own candidate-specific adapter:

- `@aiden0z/pptx-renderer@1.2.4`
- `pptx-preview@1.0.7`

The complete evidence and source links are summarized in
[`renderer-comparison.md`](../research/renderer-comparison.md).

## Decision

Use `@aiden0z/pptx-renderer@1.2.4` as the selected renderer for M0 and the M1
foundation. Keep it behind `PptxRendererAdapter` and
`PreflightPptxRendererAdapter`; product code must not depend directly on its
viewer API.

Retain `pptx-preview@1.0.7` only as an exact-pinned development dependency so
the rejected candidate's evidence remains reproducible. It must not ship in
the default production bundle.

Security maintenance note (2026-08-02): npm overrides keep the rejected
candidate's transitive ECharts and UUID packages, and the test runner's
transitive packages, above their published patched versions. The selected
Aiden production bundle remains unchanged. The committed `pptx-preview@1.0.7`
results remain historical evidence for the original dependency tree; current
candidate runs use the secure transitive tree. An A/B run with and without the
ECharts override produced the same 14 / 20 readable result under the newer
shared reading shell, so the candidate's already-stale visual baseline and
readability drift are not caused by the security override.

The M0 prerequisite PRs have landed, their linked issues are closed, and the
integrated branch passes the full verification suite. This decision authorizes
M1 engineering. It does not authorize a public release or a claim of
pixel-perfect PowerPoint fidelity.

## Decision evidence

| Gate | Required | Aiden result | Decision |
| --- | ---: | ---: | --- |
| Readable main content | >= 80% | 18 / 20 (90%) | pass |
| First-readable p95 | <= 3,000 ms | 43.9 ms, 10 / 10 samples | pass |
| Slide-switch p95 | <= 100 ms | 2.3 ms, 40 / 40 samples | pass |
| Safe abnormal inputs | stable local failure | all shared fixtures pass | pass |
| In-flight cancellation | 5 complete runs | 5 / 5; p95 12.9 ms | pass |
| Resource completion | <= 2,000 ms | 15 / 15; p95 1,853.7 ms | pass |
| License floor | open distribution and commercial evolution | Apache-2.0 | pass with distribution obligations |

The 3-second and 100-millisecond budgets are unchanged. No missing sample is
treated as zero and no candidate receives a different corpus or threshold.

## Alternatives considered

### `pptx-preview@1.0.7`

Rejected. It reaches the compatibility threshold at exactly 16 / 20 readable
units, but all 13 cold/warmup/measured opens of the unchanged representative
12-slide file fail with no usable slide. It therefore has no valid
first-readable or slide-switch measurement. It also has lower visual fidelity,
no parser-level abort/resource-limit API, a module-global event registry,
ambiguous closed-source redistribution language, and three moderate production
audit nodes in its transitive runtime chain.

### LibreOffice or Microsoft Office conversion/automation

Not selected for M0. These paths add an external application or service,
platform-specific process management, conversion latency, temporary-file
lifecycle, and a different trust boundary. They remain fallback research
options if browser-native rendering cannot meet later real-world fidelity
requirements.

The later local macOS development build uses a sandboxed LibreOffice process
to prepare a temporary OOXML derivative of Office 97–2003 files. Aiden
continues to render PPTX. Stable slide references require IDs from the
original PPTX package.

### Commercial Office rendering SDK

Not selected while the browser-native candidate meets M0. Reconsider only if
public-release corpus results show that unsupported business-critical content
cannot be handled safely with honest degradation messaging.

## Consequences

### Positive

- The selected candidate passes the fixed readability and latency gates with
  complete raw evidence.
- It accepts local binary input, supports one-slide rendering, exposes native
  abort and ZIP limits, and has deterministic destruction hooks.
- Apache-2.0 permits open distribution and future commercial Pro evolution.
  Distribution must include the license, retain applicable copyright and
  attribution notices, mark modified files, and include an upstream `NOTICE`
  only when one is supplied. Version 1.2.4 contains `LICENSE` but no `NOTICE`.
- The adapter, build alias, and candidate-neutral acceptance suite preserve a
  practical replacement seam.

### Negative and accepted risks

- Embedded SVG content is currently rendered as a broken placeholder in the
  selected implementation; the corpus records two degraded fixtures.
- EMF/WMF and embedded-PDF fallbacks are incomplete while optional PDF.js
  support is disabled.
- The dependency and its current API are new, and maintenance is concentrated.
- The 1.14 MB candidate bundle is material for an Obsidian plugin.
- Fonts unavailable on the user's machine can change metrics and layout even
  when text remains readable.
- Passing a generated corpus does not establish compatibility with every
  PowerPoint feature or every real business deck.

## M1 dependency boundary

M1 may rely on these project-owned capabilities:

- `ArrayBuffer` input from the Vault;
- package preflight before candidate allocation;
- slide count and zero-based single-slide rendering;
- abort-aware open and idempotent disposal;
- stable `malformed`, `protected`, `incompatible`, and `unknown` failure
  categories with candidate details redacted;
- candidate-neutral lifecycle diagnostics and installed acceptance commands.

M1 must continue to isolate these candidate details:

- renderer DOM structure and CSS;
- package-specific exceptions and parse messages;
- chart/image cleanup behavior;
- ZIP limit implementation and optional format fallbacks;
- renderer metadata, dependency version, and build selection.

## Replacement conditions

Reopen this ADR if any of the following occurs:

- the selected version fails the shared safety, >=80% readability, <=3-second
  first-readable, or <=100-millisecond switch gate;
- real-deck testing reveals a business-critical unsupported feature rate that
  cannot be addressed with bounded adapter work;
- license, security, maintenance, or API changes make distribution unsafe;
- another candidate passes the unchanged corpus and protocol with materially
  better fidelity or lifecycle guarantees;
- product scope expands to editing, legacy `.ppt`, animations, or other
  requirements outside the current adapter contract.

Any replacement must rerun the existing evidence suite. It cannot substitute a
friendlier corpus or silently change a budget.

## Operational evidence policy

Added 2026-07-17: the selected renderer's committed installed-performance run
remains an immutable evidence snapshot. The current production `main.js` may be
up to 5% larger than the bundle recorded by that run; exceeding the threshold
requires a new accepted installed run or an explicit revision of this policy.
The threshold does not change the latency, cancellation, cleanup, or resource
completion gates in this decision. See the
[performance evidence policy](../performance/README.md).
