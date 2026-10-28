# @qvac/office

Read, edit and create `.pptx`, `.docx` and `.xlsx` files in TypeScript, on Bare and Node, without
Python or native code. The engine behind the QVAC office tools: every edit is a method call that
returns a plain report, and every failure is a `DocsError` with a stable `code`.

```sh
npm install @qvac/office
```

## Usage

```js
import { pptx, docx, xlsx, units, DocsError } from '@qvac/office'

const deck = pptx.open(bytes)
deck.setText({ slide: 0, shape: 'title' }, 'Why Sunsets Are Red')
const out = deck.save()
```

Each format also has its own entry: `@qvac/office/pptx`, `@qvac/office/docx` and
`@qvac/office/xlsx`. Lengths are EMU; `units` converts from inches, points and centimetres.

## Layout

```
opc/        zip container, content types, relationships, the part graph   (every format)
xml.ts      namespace-aware DOM helpers over @xmldom/xmldom                 (every format)
drawingml/  text bodies, run styling and charts                            (pptx, docx, xlsx)
pptx/       slides, layouts, placeholders, shapes, notes, sections, new decks
docx/       blocks, runs, styles, sections, tables, headers, new documents
xlsx/       sheets, cells, strings, styles, formulas, reference moving, drawings, new workbooks
```

`opc/zip.ts` is our own zip reader and writer; compression goes through `#zlib`
(`bare-zlib` on Bare, `zlib` on Node).

## PowerPoint

Slides are numbered from 0. Wherever a slide is taken it can also be `{ id }`, the stable
`p:sldId/@id` that `read()`, `outline()`, `find()` and `addSlide()` report, so a plan of several
calls survives inserts, moves and deletes. A shape is `{ slide, shape }`, where `shape` is the id
`read()` reports or `'title'`, `'subtitle'` or `'body'`.

- `read({ from, count, detail } | { slides, detail })`, `layouts()` — with `detail: 'runs'`
  each paragraph also lists its runs and the formatting set on each (`bold`, `italic`,
  `underline`, `size`, `color`, `font`, `link`, `field`); inherited formatting is not resolved
- `readPage({ from, maxChars, detail })` — `{ slides, next }`: stops before the slides' JSON
  would pass `maxChars` (always at least one slide); pass `next` back as `from`. `read()` keeps
  returning a plain array and refuses `maxChars`
- `outline()` — per slide its index, id, layout, title, shape count, whether it has notes and
  its chart, table and picture counts
- `find({ text, placeholder, kind, name, regex, matchCase })` — `[{ slide, slideId, shape, name,
kind, text }]`; `text` and `name` match a substring (or a regular expression with `regex`),
  ignoring case unless `matchCase`; `placeholder` is `'title'`, `'subtitle'`, `'body'` or a
  placeholder type
- `replaceText({ find, replace, all, matchCase, slide, shape, notes })` — across the deck, a
  slide or a shape (table cells included, notes with `notes: true`); matches across runs and keeps
  the formatting of the run where the match starts; matches that touch a field or a line break
  are left alone and counted in `refused`. Returns `{ replaced, shapes: [{ slide, shape }] }`
- `formatText(target, { find, all, matchCase }, style)` — styles only the matched text,
  splitting runs; `link(target, { find, all, matchCase }, url)` — links the matched text to an
  `http`, `https` or `mailto` URL
- `setText(target, text, { fit: 'shrink' })` — keeps run and paragraph formatting, reports
  `removed` links and fields and an `overflow` estimate
- `setCellText({ slide, shape, row, column }, text)`
- `setStyle(target, style, { paragraph })`, `format(...)` (same)
- `addSlide({ like, layout, at, title, body, image })` — returns its `index` and `id`,
  `deleteSlide(slide)`, `moveSlide(slide, { to })`
- `addPicture(slide, image, box)`, `replacePicture(target, image, { stretch })`
- `addTextBox(slide, opts)`, `addTable(slide, opts)`, `setBackground(slide, color)`
- `addChart(slide, spec)`, `setChartData(target, data)`, `setChartTitle(target, title)`
- `deleteShape(target)`, `setBox(target, box)`, `setNotes(slide, text)`
- `pptx.create({ size, fonts, colors, title })`

## Word

The body is a list of blocks — paragraphs and tables — numbered from 0.

- `read({ from, count })` — blocks, headers, footers, footnotes and comments
- `setText(index, text)` — refuses tracked changes, reports `removed` links, images,
  fields, content controls, equations, objects and page or column breaks
- `insertParagraphs(at, paragraphs)`, `appendParagraphs(paragraphs)`
- `deleteBlocks(index, { count })`, `sections()`, `replaceSection(index, opts)`,
  `deleteSection(index)`
- `formatParagraph(index, format)`, `format(...)` (same), `formatStyle(name, format)`
- `addTable(at, opts)`, `setCell({ table, row, column }, text)`, `addRow(table, values)`,
  `deleteRow({ table, row })`
- `addImage(at, image, opts)`, `addChart(at, spec)`, `setChartData(index, data)`,
  `setChartTitle(index, title)`
- `setHeader(text, { type })`, `setFooter(text, { type })`
- `docx.create({ title, page, size, fonts })`

## Excel

A sheet is its name or 0-based index; cells use A1 references and rows are numbered from 1.

- `sheets()`, `read(sheet, { range, limit })`, `get(sheet, ref)`
- `setCells(sheet, cells)` — text starting with `=` must parse as a formula
- `format(sheet, range, style)`, `appendRows(sheet, rows)`
- `insertRows(sheet, { at, rows | count })`, `deleteRows(sheet, { at, count })`,
  `deleteColumns(sheet, { column, count })` — move every reference the way Excel does; pivot
  tables, comments, data tables, straddled array formulas and table header, totals or column
  deletes are refused
- `insertColumns(sheet, { column, count | values })` — `values` is one array per new column,
  from row 1 down; moves references and refuses the same things `deleteColumns` does
- `setColumnWidth(sheet, { column, width })`, `setRowHeight(sheet, { row, height })` (points, or
  a length like `"0.5in"`), `hide(sheet, { rows, columns, hidden })` with a number, letter or
  `[from, to]`
- `merge(sheet, range)` keeps the top-left value and reports the cells it cleared;
  `unmerge(sheet, range)` drops every merge the range touches
- `freeze(sheet, { rows, columns })`; `{ rows: 0, columns: 0 }` unfreezes
- `autofilter(sheet, range | null)`, `sort(sheet, { range, by: [{ column, descending }], header })`
  — sort refuses formulas that refer inside the range, merged cells and ranges straddling a table
- `addConditionalFormat(sheet, range, rule)` with `cellIs`, `colorScale`, `dataBar` or `text`
  rules, `deleteConditionalFormats(sheet, range?)`
- `addValidation(sheet, range, { type, values | min | max, operator, allowBlank, prompt, error })`
  for `list`, `whole`, `decimal`, `date` and `textLength`, `deleteValidations(sheet, range?)`
- `names()`, `defineName(name, ref, { sheet })`, `deleteName(name, { sheet })` — names follow
  Excel's rules and share one namespace with tables
- `tables(sheet?)`, `addTable(sheet, range, { name, header, style, totals })`,
  `deleteTable(name)`, which turns structured references to the table into plain ranges
- `find(query, { sheet, regex, matchCase, formulas, limit })` and `outline()`, a short summary of
  every sheet: used range, header row guess, size, charts and tables
- `addSheet(name, { at })`, `renameSheet(sheet, name)`, `deleteSheet(sheet)`
- `addImage(sheet, image, opts)`, `addChart(sheet, spec)`, `removeMacros()`
- `xlsx.create({ sheets, title })`

## Tools

`@qvac/office/tools` is the surface an LLM calls. Each operation is a method of the format
class, described by a JSON Schema whose fields are the method's arguments, with option
objects flattened beside the target.

```js
import { tools } from '@qvac/office'

const definition = tools.tool('docx', { name: 'word' })
const { doc, results } = tools.apply(document, [
  { op: 'appendParagraphs', paragraphs: [{ text: 'Plan', style: 'Heading 1' }] },
  { op: 'addImage', at: 1, image: 'data:image/png;base64,...', width: '3in' }
])
```

- `tool(format, { name, only })` is one tool definition whose input is `{ ops: [...] }`.
  `operations(format)` lists each operation with its own schema.
- `apply(doc, ops, { resolveImage })` runs the ops in order on a copy and returns it with
  each op's result. If one fails, none apply, and the `DocsError` names the op as
  `details.index` and `details.op`.
- Images arrive as base64, a `data:` URL or bytes. `resolveImage` maps anything else, such as
  an attachment id.
- `detect(bytes)` and `open(bytes)` work out the format from the file.
- A length is EMU or a string such as `"2in"`, `"5cm"`, `"40mm"`, `"72pt"` or `"96px"`.
- Errors serialise to `{ code, message, details }`. `details` carries fields such as `field`,
  `expected`, `got` and `candidates` when the next call can use them.

## Diff and render

- `diff.diff(before, after, { limit, excerpt })` takes two versions of a document, as bytes
  or open documents. It returns what was added, removed and changed, by block, slide and
  shape, or cell.
- `@qvac/office/render` draws pages to PNG on Bare with
  `render(doc, { pages, scale })`. `toPdf(doc)` returns the PDF.
  - They need the optional peers `bare-x-to-pdf`, `bare-pdfium` and `bare-png`.
  - Without them, or on Node, they fail with `RENDER_UNAVAILABLE`.

## Guarantees

- Parts an edit does not touch are saved byte for byte.
- Every failure is a `DocsError` whose `code` is in `codes.ts` and whose message names what
  does exist.
- Opening bounds entry count, size, compression ratio and XML depth, and refuses a DOCTYPE.
- `Package.strict` fails a save when a part changed without `touch()`; the tests run strict.
- `save({ modified, modifiedBy })` records the edit in the document properties.

## Testing

`test/` covers every format, a fuzz test (`FUZZ_ROUNDS` raises its rounds) and a
corpus of real files from the Apache POI test data in `test/fixtures/corpus/`, with the
recorded outcome of each in `expected.json`. After changing behaviour, check the corpus with
`node scripts/validate/corpus.ts` and record a reviewed change with `--write`.

`scripts/validate/` also validates output against the Open XML schemas:

```sh
node scripts/validate/outputs.ts validate-out
dotnet run --project scripts/validate -- validate-out
```

`scripts/validate/render.mjs` renders LibreOffice's PDF conversion of those files with
[bare-pdfium](https://github.com/tetherto/bare-pdfium) to catch blank output; install it first
(`npm i --no-save bare-pdfium`) and run it with `bare`.

## Development

```sh
npm install
npm run lint
npm test
npm run build
```

`npm run build` compiles `lib/` to `dist/` with declarations; the package ships only `dist/`.
After adding a test file, regenerate the runner with `npm run test:generate`.

## License

Apache-2.0
