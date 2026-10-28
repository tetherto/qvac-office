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

Slides are numbered from 0. A shape is `{ slide, shape }`, where `shape` is the id `read()`
reports or `'title'`, `'subtitle'` or `'body'`.

- `read({ from, count } | { slides })`, `layouts()`
- `setText(target, text, { fit: 'shrink' })` — keeps run and paragraph formatting, reports
  `removed` links and fields and an `overflow` estimate
- `setCellText({ slide, shape, row, column }, text)`
- `setStyle(target, style, { paragraph })`, `format(...)` (same)
- `addSlide({ like, layout, at, title, body, image })`, `deleteSlide(slide)`, `moveSlide(slide, { to })`
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
- `setColumnWidth(sheet, { column, width })`
- `addSheet(name, { at })`, `renameSheet(sheet, name)`, `deleteSheet(sheet)`
- `addImage(sheet, image, opts)`, `addChart(sheet, spec)`, `removeMacros()`
- `xlsx.create({ sheets, title })`

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
