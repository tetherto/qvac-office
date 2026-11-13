---
id: ADR-2026-09-28-office-files-are-edited-by-the-harness-engine
status: active
introduced_by: PR-150
supersedes: []
---

# Office files are edited by an engine inside the harness

## Decision

`.pptx`, `.docx` and `.xlsx` files are read, edited and created by the TypeScript engine in
`lib/office/engine/`, not by Python. It is harness source under the house style: a format-neutral
package layer (`opc/`: zip container, content types, relationships, part graph), shared DrawingML
text and charts (`drawingml/`), and one module per format (`pptx/`, `docx/`, `xlsx/`). Its only new
dependency is `@xmldom/xmldom`; compression goes through `#zlib`, which is `bare-zlib` on Bare and
`zlib` on Node, and the zip container is our own (`opc/zip.ts`).

## Context

The python skills run on Pyodide, which does not ship on mobile, so office editing was desktop-only.
The engine started as a standalone package (`bare-docs`) and moved in so the harness owns it
without a separate package to publish, and so the office tools and the engine change in one PR.
fflate was dropped on the way: under ESM, Bare resolves it to a Node build that imports `module`,
and its browser build is ESM in a `.js` file that Bare loads as CommonJS.

## Alternatives Considered

- Publishing the engine as its own package: the apps install the harness from npm and rebuild its
  worker from `node_modules`, so a `file:` dependency cannot reach them; publishing was the cost we
  wanted to avoid.
- Keeping the engine as CommonJS with `.d.cts` types: breaks "no new raw .js/.cjs product sources".

## Consequences

- Every edit keeps untouched parts byte for byte, refuses or reports silent losses, and fails with a
  `DocsError` whose code is registered in `lib/office/engine/codes.ts`; `test/office-engine/` holds
  a fuzz test and a corpus of real files (`test/fixtures/office-corpus/`) that pin this.
- `Package.strict` makes `save()` fail when a part changed without `touch()`; the tests run strict.
- Excel row and column changes move references the way Excel does; pivot tables, comments, data
  tables, straddled array formulas and table header, totals or column deletes are still refused.
- `scripts/office-validate/` checks output with the Open XML SDK validator; it is not wired into CI.
- Revisit if an app needs the engine without the harness: extract it again as a package then.

## When Agents Should Read This

Read this when changing the paths, workflows, runtime boundaries, or platform
behavior listed for this ADR in `docs/decisions/INDEX.md`.
