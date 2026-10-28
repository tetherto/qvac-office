import fs from 'fs'
import path from 'path'
import process from 'process'
import { pptx, docx, xlsx, units } from '../../lib/index.ts'
import { edits, expected, readCorpus, type Format } from '../../test/helpers/corpus.ts'

type Edit = (bytes: Uint8Array) => Uint8Array

const FIXTURES = new URL('../../test/fixtures/engine/', import.meta.url)
const out = process.argv[2] ?? 'validate-out'
const png = fixture('red.png')
const chart = {
  categories: ['North', 'South', 'East'],
  series: [
    { name: 'Units', values: [3, 5, 2] },
    { name: 'Returns', values: [1, 0, 1] }
  ]
}

function fixture(name: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(new URL(name, FIXTURES)))
}

function write(name: string, bytes: Uint8Array): void {
  fs.writeFileSync(path.join(out, name), bytes)
}

function pair(name: string, input: Uint8Array, edit: Edit): void {
  write(name.replace(/(\.\w+)$/, '.orig$1'), input)
  write(name, edit(input))
}

function createdDeck(): Uint8Array {
  const deck = pptx.create({ title: 'Validation', colors: { accent1: '1A73E8' } })
  deck.addSlide({ layout: 'Title Slide', title: 'Cover', body: 'Subtitle' })
  deck.addSlide({ title: 'Bullets', body: ['One', { text: 'Two', level: 1 }] })
  deck.addSlide({ layout: 'Section Header', title: 'Section', body: 'Intro' })
  const two = deck.addSlide({ layout: 'Two Content', title: 'Compare' })
  const left = two.shapes[1]?.id
  if (left !== undefined && left !== null) {
    deck.setText({ slide: two.index, shape: left }, 'Left')
  }

  for (const type of ['column', 'bar', 'line', 'pie', 'doughnut', 'area'] as const) {
    const slide = deck.addSlide({ layout: 'Title Only', title: type })
    deck.addChart(slide.index, {
      type,
      title: type,
      stacked: type === 'column' ? 'percent' : false,
      ...chart
    })
  }

  const table = deck.addSlide({ layout: 'Title Only', title: 'Table' })
  deck.addTable(table.index, {
    rows: [
      ['A', 'B'],
      ['1', 2]
    ],
    height: units.inches(1)
  })
  const blank = deck.addSlide({ layout: 'Blank' })
  deck.setBackground(blank.index, '101820')
  deck.addTextBox(blank.index, {
    text: 'Thanks',
    fill: 'F1F3F4',
    style: { size: 40, bold: true, align: 'center' }
  })
  deck.addPicture(blank.index, png, { width: units.inches(1) })
  deck.setNotes(1, 'Speaker notes')
  const lines = Array.from({ length: 20 }, (_, i) => `Line ${i}`)
  deck.setText({ slide: 1, shape: 'body' }, lines, { fit: 'shrink' })

  return deck.save({ modified: new Date('2026-01-01T00:00:00Z') })
}

function editedDeck(bytes: Uint8Array): Uint8Array {
  const d = pptx.open(bytes)
  d.setText({ slide: 0, shape: 'title' }, 'Edited')
  d.format({ slide: 0, shape: 'title' }, { color: 'C0392B', font: 'Georgia' })
  d.setCellText({ slide: 2, shape: 4, row: 1, column: 1 }, 'x')
  d.addSlide({ title: 'Added', body: 'Body', at: 1 })
  d.moveSlide(3, { to: 0 })
  d.deleteShape({ slide: 0, shape: 6 })
  d.setBox({ slide: 1, shape: 'title' }, { y: units.inches(0.3) })
  d.replacePicture({ slide: 0, shape: 3 }, png)
  d.setNotes(2, 'Notes')
  d.deleteSlide(2)
  return textEdits(d).save()
}

function textEdits(d: pptx.Presentation): pptx.Presentation {
  const [hit] = d.find({ text: 'scatter' })
  d.replaceText({ find: 'scatter', replace: 'spread', all: true, notes: true })
  d.formatText({ slide: { id: hit.slideId }, shape: hit.shape }, { find: 'spread' }, { bold: true })
  d.link(
    { slide: { id: hit.slideId }, shape: hit.shape },
    { find: 'spread' },
    'https://example.com'
  )
  d.moveSlide({ id: hit.slideId }, { to: 0 })

  return d
}

function editedCharts(bytes: Uint8Array): Uint8Array {
  const d = pptx.open(bytes)
  const series = [1, 2, 3].map((n) => ({ name: `S${n}`, values: [n, n + 1, n + 2, n + 3] }))
  d.setChartData({ slide: 0, shape: 3 }, { categories: ['Q1', 'Q2', 'Q3', 'Q4'], series })
  d.setChartTitle({ slide: 0, shape: 3 }, 'Retitled')
  d.setChartTitle({ slide: 1, shape: 2 }, 'Added title')
  return d.save()
}

function createdDocument(): Uint8Array {
  const doc = docx.create({ title: 'Validation', page: 'a4' })
  doc.appendParagraphs([
    { text: 'Title', style: 'Title' },
    { text: 'Heading', style: 'Heading 1' },
    { text: 'Bullet', style: 'List Bullet' },
    { text: 'Nested', style: 'List Bullet', level: 1 },
    { text: 'Number', style: 'List Number' },
    { runs: ['Plain ', { text: 'bold', bold: true, color: 'C0392B' }, '\ttab'] },
    { pageBreak: true }
  ])
  doc.addTable(doc.blockCount, {
    rows: [
      ['A', 'B'],
      ['1', 2]
    ]
  })
  doc.addImage(doc.blockCount, png, { align: 'center' })
  doc.addChart(doc.blockCount, { type: 'bar', title: 'Chart', ...chart })
  doc.setHeader('Header')
  doc.setFooter('Footer')
  doc.setHeader('First', { type: 'first' })
  doc.formatStyle('Normal', { size: 12, font: 'Georgia' })

  return doc.save()
}

function editedReport(bytes: Uint8Array): Uint8Array {
  const d = docx.open(bytes)
  d.setText(3, { runs: ['Revenue ', { text: 'grew', italic: true }] })
  d.insertParagraphs(6, [{ text: 'East', style: 'List Bullet', level: 1 }])
  d.replaceSection(7, { title: 'Areas', body: ['All reported.'] })
  d.setCell({ table: 11, row: 1, column: 2 }, '38%')
  d.addRow(11, ['Headcount', 1, 2])
  d.formatParagraph(1, { align: 'center', size: 14 })
  d.addImage(0, png)
  d.setFooter('Footer')
  return d.save()
}

function textEditedReport(bytes: Uint8Array): Uint8Array {
  const d = docx.open(bytes)
  const { ids } = d.insertParagraphs(2, ['Summary in brief.', 'Second note.'])
  d.deleteBlocks({ id: ids[1] })
  d.replaceText({ find: 'grew twelve', replace: 'rose eleven' })
  d.replaceText({ find: 'year', replace: 'yr', all: true, in: { id: ids[0] } })
  d.formatText({ id: ids[0] }, { find: 'brief' }, { italic: true, color: 'C0392B' })
  d.link(4, { find: 'eleven percent' }, 'https://example.com/revenue')
  d.addTable({ id: ids[0] }, { rows: [['A', 'B']] })
  d.addRow(2, ['1', '2'])
  return d.save()
}

function createdWorkbook(): Uint8Array {
  const book = xlsx.create({ sheets: ['Data', 'Summary'] })
  book.appendRows('Data', [
    ['Month', 'Spend'],
    ['Jan', 120],
    ['Feb', 80.5]
  ])
  book.format('Data', 'A1:B1', { bold: true, fill: 'FFF2CC' })
  book.setCells('Data', {
    B4: '=SUM(B2:B3)',
    C2: { date: '2026-01-31' },
    D2: { value: 0.25, format: '0.0%' }
  })
  book.setCells('Summary', { A1: 'Total', B1: '=Data!B4' })
  book.setColumnWidth('Data', { column: 'A', width: 14 })
  book.addImage('Summary', png, { cell: 'D2' })
  book.addChart('Data', {
    type: 'column',
    title: 'Spend',
    categories: 'A2:A3',
    series: [{ values: 'B2:B3', nameCell: 'B1' }]
  })
  book.insertRows('Data', { at: 3, rows: [['Jan.5', 10]] })
  book.deleteRows('Data', { at: 3 })

  return book.save()
}

function sheetFeatures(): Uint8Array {
  const book = xlsx.create({ sheets: ['Data', 'Lists'] })
  book.appendRows('Data', [
    ['Region', 'Units', 'Price'],
    ['North', 3, 2.5],
    ['South', 5, 1.5],
    ['East', 2, 4]
  ])
  book.insertColumns('Data', { column: 'C', values: [['Status', 'open', 'done', 'open']] })
  book.setCells('Data', { F1: 'Note', F5: '=SUM(B2:B4)' })
  book.merge('Data', 'F2:G3')
  book.setRowHeight('Data', { row: 1, height: 24 })
  book.hide('Data', { rows: 6, columns: 'H' })
  book.freeze('Data', { rows: 1, columns: 1 })
  book.sort('Data', { range: 'A1:D4', header: true, by: [{ column: 'B', descending: true }] })
  book.addTable('Data', 'A1:D4', { name: 'Sales', totals: true })
  book.addConditionalFormat('Data', 'B2:B4', {
    kind: 'cellIs',
    operator: 'greaterThan',
    value: 2,
    style: { fill: 'FFC7CE', color: '9C0006' }
  })
  book.addConditionalFormat('Data', 'D2:D4', { kind: 'colorScale', colors: ['F8696B', '63BE7B'] })
  book.addConditionalFormat('Data', 'B2:B4', { kind: 'dataBar', color: '638EC6' })
  book.addConditionalFormat('Data', 'C2:C4', {
    kind: 'text',
    contains: 'open',
    style: { bold: true }
  })
  book.addValidation('Data', 'C2:C4', { type: 'list', values: ['open', 'done'], prompt: 'Pick' })
  book.addValidation('Data', 'B2:B4', { type: 'whole', min: 0, max: 100, error: 'Whole units' })
  book.addValidation('Data', 'J2:J9', { type: 'date', min: '2026-01-01' })
  book.setCells('Lists', { A1: 'x', A2: 'y', B1: 'z' })
  book.autofilter('Lists', 'A1:B2')
  book.defineName('TaxRate', '=0.2')
  book.defineName('Picks', 'A1:A2', { sheet: 'Lists' })

  return book.save()
}

function editedSales(bytes: Uint8Array): Uint8Array {
  const w = xlsx.open(bytes)
  w.setCells('Sales', { B3: 125, F4: { value: 0.3, format: '0.0%' }, G4: { date: '2026-06-30' } })
  w.appendRows('Sales', [['Q4', 80, 10.5, '=B8*C8']])
  w.renameSheet('Notes', 'Commentary')
  w.addSheet('Extra')
  w.addChart('Sales', {
    type: 'line',
    categories: 'A3:A5',
    series: [{ values: 'B3:B5', name: 'Units' }]
  })
  w.insertRows('Sales', { at: 4, rows: [['Q1.5', 130, 9.6, '=B4*C4']] })
  w.deleteRows('Sales', { at: 3 })
  w.deleteColumns('Sales', { column: 'E' })

  return w.save()
}

function laidOutDeck(): Uint8Array {
  const deck = pptx.create({ size: '4:3' })
  deck.setSize('16:9', { scale: true })
  deck.setTheme({ colors: { accent1: '1A73E8' }, fonts: { heading: 'Georgia' } })
  deck.addSlide({ title: 'Chart' })
  const placed = deck.addChart(0, { type: 'column', into: 'body', ...chart })
  const cover = deck.addSlide({ layout: 'Two Content', title: 'Shapes', body: 'Left' })
  const box = deck.addTextBox(cover.index, { text: 'Boxed', x: '1in', y: '5in', width: '3in' })
  deck.setFill({ slide: cover.index, shape: box.id }, 'F1F3F4')
  deck.setLine(
    { slide: cover.index, shape: box.id },
    { color: '1A73E8', width: '2pt', dash: 'dash' }
  )
  const picture = deck.addPicture(cover.index, png, { into: cover.shapes[2]?.id ?? 'body' })
  deck.duplicateShape({ slide: cover.index, shape: picture.id })
  deck.duplicateShape({ slide: 0, shape: placed.id }, { slide: cover.index, x: '6in', y: '5in' })
  deck.reorder({ slide: cover.index, shape: box.id }, 'front')
  const grouped = deck.group(cover.index, [box.id, picture.id], { name: 'Pair' })
  deck.addSlide({ layout: 'Title Slide', title: 'Moved', body: 'Subtitle' })
  deck.setLayout(2, 'Title Only')
  deck.ungroup({ slide: cover.index, shape: grouped.id })

  return deck.save()
}

function editedLayout(bytes: Uint8Array): Uint8Array {
  const d = pptx.open(bytes)
  d.setLayout(1, 'Two Content')
  d.group(2, [3, 4])
  d.duplicateShape({ slide: 2, shape: 5 }, { slide: 0 })
  d.setTheme({ colors: { accent2: 'C0392B' } })
  d.setSize('16:9', { scale: true })

  return d.save()
}

fs.mkdirSync(out, { recursive: true })
write('created.pptx', createdDeck())
write('created.docx', createdDocument())
write('created.xlsx', createdWorkbook())
write('sheet-features.xlsx', sheetFeatures())
pair('edited-basic.pptx', fixture('basic.pptx'), editedDeck)
pair('edited-charts.pptx', fixture('charts.pptx'), editedCharts)
pair('edited-report.docx', fixture('report.docx'), editedReport)
pair('edited-report-text.docx', fixture('report.docx'), textEditedReport)
pair('edited-sales.xlsx', fixture('sales.xlsx'), editedSales)
write('laid-out.pptx', laidOutDeck())
pair('edited-layout.pptx', fixture('basic.pptx'), editedLayout)
for (const [name, outcome] of Object.entries(expected())) {
  if (outcome !== 'ok') continue
  const [format, file] = name.split('/')
  if (format === 'pptx' || format === 'docx' || format === 'xlsx') {
    const edit: Edit = edits[format satisfies Format]
    pair(`corpus-${file}`, readCorpus(name), edit)
  }
}
