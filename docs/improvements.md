# Improvements

What would make `@qvac/office` better for its main caller, an LLM using it through tool calls.
The ordering is by value to that caller. Items marked _seen_ came up while building the editor.

Status, 2026-10-28:

- Sections 1 to 6 are implemented.
- Section 7 needs the layout map from bare-x-to-pdf; the library side is in place (slide size, stable ids, run-level read).
- Section 8 is a change to the harness.
- Two gaps remain:
  - Errors raised by code older than this work carry no `details` yet.
  - The Open XML validator runs in CI; it was not run locally.

## 1. A tool surface owned by the library

- **Ship the operation catalog and JSON Schemas.** The harness keeps its own op table
  (`lib/tools/office/ops.ts`) and the editor needed a second one. One `tools()` export that
  returns each method's name, summary and JSON Schema, plus `apply(doc, { op, args })`, gives
  every host the same contract and stops the tables drifting. _seen_
- **Atomic batches.** `apply(doc, ops[])` that runs every op or none. Today a method that
  throws part way can leave the DOM half edited. The editor works around it by reopening the
  last saved bytes. _seen_
- **Units in the call.** Accept `"2in"`, `"5cm"` and `"12pt"` wherever a length is taken, so
  callers stop converting to EMU. Every host does that conversion today.
- **Structured error details.** Errors already carry a `code` and name what exists. Adding
  `details` (`{ field, expected, got, candidates }`) would let a model fix its next call
  without parsing prose.

## 2. Addressing that survives a plan

- **Stable references.** Block indexes and slide numbers move after an insert or delete, so a
  plan of several calls breaks part way through. Return stable ids (`w14:paraId` for docx
  paragraphs, shape ids for pptx), accept them wherever an index is taken, and return the new
  ids from every insert.
- **Find.** `find({ text | style | heading | placeholder })` returns references, so a model can
  say "the paragraph that starts with Revenue" without reading the whole document first.

## 3. Edits at the scale a model makes them

- **Text ranges.** `replaceText(target, { find, replace, all })` and formatting a substring
  (bold one word, link a phrase), keeping the formatting around it. Today `setText` replaces a
  whole paragraph or shape, and `setStyle` and `formatParagraph` style all of it.
- **Markdown in, markdown out.** Models write markdown naturally. `appendMarkdown(md)` /
  `insertMarkdown(at, md)` for docx (headings, lists, tables, bold, links) and a markdown
  `read({ as: 'markdown' })` would replace most hand-built `ParagraphSpec` arrays.
- **Run-level read.** `read({ detail: 'runs' })` returns each run's formatting, so a model can
  keep emphasis when it rewrites text. A visual editor needs this too (see 6).

## 4. Reads that fit a context window

- **Outline mode.** Headings with block ranges for docx, titles and layouts for pptx, used
  ranges and header rows for xlsx, at a few hundred tokens for any size of file.
- **Budgets and cursors.** `maxChars` and a `next` cursor on every read, in place of
  `from`/`count` arithmetic.
- **Diff after an edit.** `diff(before, after)` summarising what changed, so a model can check
  and report its own work.

## 5. Output that looks right

- **Cached formula values.** `setCells` writes a formula with no cached value, so anything that
  does not recalculate (bare-x-to-pdf, Quick Look, mobile previews) shows an empty cell, and
  `read` returns `null` for it. Evaluating the common functions (`SUM`, `AVERAGE`, arithmetic,
  `IF`, references) when writing would fix both. _seen_
- **Layout-aware placement.** A new pptx chart with no box overlapped the body placeholder.
  Report `overlaps` in the result, take `into: 'body'` to fill a placeholder, and give the
  slide size in `read()` (no method exposes it today). _seen_
- **A render hook.** An optional `render(doc, { pages })` that returns PNGs through
  bare-x-to-pdf and bare-pdfium, so a multimodal model can look at what it made.
- **Validation in CI.** `scripts/validate/` checks output against the Open XML SDK but is not
  wired into CI.

## 6. Missing capabilities

- **Word:** add and resolve comments, add footnotes, hyperlinks, explicit bulleted and
  numbered lists, find and replace, page setup (margins, orientation), a table of contents,
  and accepting or rejecting tracked changes (today a tracked paragraph is refused).
- **PowerPoint:** shape fill and outline, duplicating a shape, z-order, grouping, hyperlinks,
  editing layouts and the theme, and slide size.
- **Excel:** `insertColumns` (only `deleteColumns` exists), merge and unmerge, row height,
  hiding rows and columns, freeze panes, autofilter and sort, conditional formatting, data
  validation, named ranges and tables (ListObjects).
- **Templates:** open `.dotx`, `.potx` and `.xltx` and save them as documents.

## 7. Toward a visual editor

Today's editor shows a structured view next to a PDF preview. To edit directly on the page:

- **pptx** is closest: shapes already carry boxes in EMU. With the slide size, the editor can lay
  hit areas over the rendered slide for select, drag (`setBox`) and in-place text
  (`setText`), with no change to the converter.
- **docx** needs a map from the rendered page back to the model. bare-x-to-pdf would emit, for
  each paragraph, table cell and run, its page and bounding box tagged with the engine's id
  (see 2). Clicking the preview then selects a block, and an overlay text box commits through
  `setText` or `replaceText`. Conversion takes about 0.1 s warm, so re-rendering after each
  commit keeps up.
- **xlsx** already edits in place in the grid. It needs cell styles in `read()` to look like
  the sheet.

The other routes are an HTML renderer of our own (a second layout engine whose output drifts
from the PDF) and embedding Collabora (full editing, but its edits bypass this library, which
the editor exists to test).

## 8. Finishing the extraction

Switch the harness to depend on `@qvac/office`, delete `lib/office/engine/` and its tests there,
and update the 2026-09-28 decision record to point here.
