import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import * as units from '../lib/units.ts'
import * as utf8 from '../lib/utf8.ts'
import { fixture, corpusFile, decode, unzip, zipFiles } from './helpers/engine.ts'

const sales = (): xlsx.Workbook => xlsx.open(fixture('sales.xlsx'))
const budget = (): xlsx.Workbook => xlsx.open(fixture('shared.xlsx'))
const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const partXml = (book: xlsx.Workbook, name: string): string => decode(unzip(book.save())[name])

function unchecked(value: unknown): value is never {
  return value !== undefined
}

function styleIndex(sheet: string, ref: string): string | null {
  const found = new RegExp(`<c r="${ref}" s="(\\d+)"`).exec(sheet)
  return found ? found[1] : null
}

test('lists sheets and reads values, formulas, dates, booleans and merges', (t) => {
  const book = sales()
  t.alike(
    book.sheets().map((s) => [s.name, s.range]),
    [
      ['Sales', 'A1:H6'],
      ['Notes', 'A1:B2']
    ]
  )
  const read = book.read('Sales')
  t.alike(read.rows[2], ['Q1', 120, 9.5, null, null, 0.25, '2026-03-31', true])
  t.is(read.formulas.D6, '=SUM(D3:D5)')
  t.alike(read.merged, ['A1:D1'])
  t.alike(book.get('Sales', 'F3'), { value: 0.25, formula: null, format: '0.0%' })
  t.alike(book.read(1).formulas, { B1: '=Sales!D6*2' })
})

test('reads shared strings and cached formula results', (t) => {
  const read = budget().read(0)
  t.alike(read.rows, [
    ['Item', 'Cost'],
    ['Rent', 1200],
    ['Food', 450.5],
    ['Total', 1650.5]
  ])
})

test('read takes a range and a cell limit', (t) => {
  const book = sales()
  const part = book.read('Sales', { range: 'B3:C4' })
  t.is(part.range, 'B3:C4')
  t.alike(part.rows, [
    [120, 9.5],
    [150, 9.75]
  ])
  const capped = book.read('Sales', { limit: 16 })
  t.ok(capped.truncated)
  t.is(capped.rows.length, 2)
})

test('rejects files that are not workbooks and names what exists', (t) => {
  t.exception(() => xlsx.open(fixture('report.docx')), /not an Excel workbook/)
  const book = sales()
  t.exception(
    () => book.read('Summary'),
    /no sheet "Summary"; the workbook has 0 "Sales", 1 "Notes"/
  )
  t.exception(() => book.get('Sales', 'Z'), /is not a cell reference/)
})

test('setCells writes numbers, text, booleans, formulas and dates', (t) => {
  const book = budget()
  book.setCells('Budget', {
    B2: 1250,
    A5: 'Travel',
    B5: 300,
    C1: 'Paid',
    C2: true,
    B4: '=SUM(B2:B3,B5)',
    D1: { date: '2026-04-01' },
    E1: { value: '=not a formula' }
  })
  const again = reopen(book)
  const read = again.read(0)
  t.alike(read.rows[0], ['Item', 'Cost', 'Paid', '2026-04-01', '=not a formula'])
  t.alike(read.rows[1].slice(0, 3), ['Rent', 1250, true])
  t.alike(read.rows[4].slice(0, 2), ['Travel', 300])
  t.is(read.formulas.B4, '=SUM(B2:B3,B5)')
  t.is(read.rows[3][1], null)

  const workbook = partXml(again, 'xl/workbook.xml')
  t.ok(workbook.includes('fullCalcOnLoad="1"'))
  const strings = partXml(again, 'xl/sharedStrings.xml')
  t.ok(strings.includes('<si><t>Travel</t></si>'))
  t.ok(strings.includes('uniqueCount="8"'))
})

test('setCells applies styles and reuses identical ones', (t) => {
  const book = budget()
  book.setCells(0, {
    C2: { value: 0.25, format: '0.0%', bold: true },
    C3: { value: 0.5, format: '0.0%', bold: true },
    A1: { fill: 'FFF2CC', align: 'center' }
  })
  const again = reopen(book)
  t.is(again.get(0, 'C2').format, '0.0%')
  t.is(again.get(0, 'A1').value, 'Item')
  const styles = partXml(again, 'xl/styles.xml')
  t.ok(styles.includes('formatCode="0.0%"'))
  t.ok(styles.includes('<fgColor rgb="FFFFF2CC"/>'))
  t.ok(styles.includes('horizontal="center"'))
  const sheet = partXml(again, 'xl/worksheets/sheet1.xml')
  const s2 = styleIndex(sheet, 'C2')
  t.ok(s2)
  t.is(s2, styleIndex(sheet, 'C3'))
  const odd: unknown = { A1: { colour: 'red' } }
  if (unchecked(odd)) {
    t.exception(() => book.setCells(0, odd), /unknown cell style "colour"/)
  }
})

test('setCells refuses merged cells and bad input', (t) => {
  const book = sales()
  t.exception(
    () => book.setCells('Sales', { B1: 'x' }),
    /B1 is inside the merged range A1:D1; write its top-left cell A1/
  )
  t.exception(() => book.setCells('Sales', { A1: { date: '2026-02-30' } }), /not a real date/)
  t.exception(
    () => book.setCells('Sales', { A1: { value: 1, formula: '=1' } }),
    /only one of value, formula or date/
  )
  t.exception(() => book.setCells('Sales', {}), /cells must be an object/)
  t.is(book.changed, false)
})

test('writing a cell drops the stale calculation chain', (t) => {
  const book = budget()
  book.package.add(
    '/xl/calcChain.xml',
    utf8.encode('<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'
  )
  book.package
    .rels(book.main)
    .add(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
      '/xl/calcChain.xml'
    )
  book.setCells(0, { B2: 1 })
  const files = unzip(book.save())
  t.absent(files['xl/calcChain.xml'])
  t.absent(decode(files['xl/_rels/workbook.xml.rels']).includes('calcChain'))
})

test('appendRows adds after the last used row', (t) => {
  const book = budget()
  t.alike(
    book.appendRows(0, [
      ['Gym', 40],
      ['Books', 25.5]
    ]),
    { from: 5, to: 6 }
  )
  const read = reopen(book).read(0)
  t.alike(read.rows.slice(4), [
    ['Gym', 40],
    ['Books', 25.5]
  ])
  t.is(read.range, 'A1:B6')
})

test('insertRows and deleteRows move rows and keep formats', (t) => {
  const book = xlsx.create()
  book.setCells(0, {
    A1: 'Year',
    B1: 'Units',
    A2: 2016,
    B2: { value: 278, format: '0.00' },
    A3: 2017,
    B3: 300
  })
  t.alike(book.insertRows(0, { at: 2, rows: [[2015, 250]] }), { from: 2, to: 2 })
  let read = reopen(book).read(0)
  t.alike(read.rows, [
    ['Year', 'Units'],
    [2015, 250],
    [2016, 278],
    [2017, 300]
  ])
  t.is(reopen(book).get(0, 'B2').format, '0.00')
  t.alike(book.deleteRows(0, { at: 2, count: 2 }), { deleted: 2 })
  read = reopen(book).read(0)
  t.alike(read.rows, [
    ['Year', 'Units'],
    [2017, 300]
  ])
})

test('insertRows takes exactly one of rows or count', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 1 })
  t.exception(
    () => book.insertRows(0, { at: 1, rows: [[1]], count: 1 }),
    /give the rows to insert, or how many empty rows/
  )
  t.exception(() => book.insertRows(0, { at: 1 }), /give the rows to insert, or how many/)
  t.exception(() => book.insertRows(0, { at: 0, count: 1 }), /at is a row number from 1/)
  t.exception(() => book.deleteRows(0, { at: 1, count: 0 }), /count must be at least 1/)
  const bare: unknown = 2
  if (unchecked(bare)) {
    t.exception(() => book.deleteRows(0, bare), /deleteRows takes an options object/)
  }

  t.alike(book.insertRows(0, { at: 1, count: 2 }), { from: 1, to: 2 })
  t.is(reopen(book).get(0, 'A3').value, 1)
})

test('row shifts on a sheet with formulas move every reference to it', (t) => {
  const book = sales()
  book.insertRows('Sales', { at: 3, rows: [['Q0', 100, 9]] })
  const again = reopen(book)
  const read = again.read('Sales')
  t.alike(read.rows[2].slice(0, 3), ['Q0', 100, 9])
  t.is(read.formulas.D4, '=B4*C4')
  t.is(read.formulas.B7, '=SUM(B4:B6)')
  t.alike(read.merged, ['A1:D1'])
  t.is(again.read('Notes').formulas.B1, '=Sales!D7*2')
  t.ok(partXml(again, 'xl/workbook.xml').includes('>Sales!$B$7</definedName>'))
})

test('deleteColumns shifts cells and column widths', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 'a', B1: 'b', C1: 'c', D1: 'd' })
  book.setColumnWidth(0, { column: 'A', width: 10 })
  book.setColumnWidth(0, { column: 'C', width: 30 })
  t.alike(book.setColumnWidth(0, { column: 4, width: 40 }), { changed: true })
  t.alike(book.deleteColumns(0, { column: 'B' }), { deleted: 1 })
  const again = reopen(book)
  t.alike(again.read(0).rows, [['a', 'c', 'd']])
  const sheet = partXml(again, 'xl/worksheets/sheet1.xml')
  t.ok(
    sheet.includes(
      '<col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="3" width="40" customWidth="1"/>'
    )
  )
  t.exception(() => book.setColumnWidth(0, { column: 'A', width: 0 }), /from 0 to 255/)
})

test('sheets can be added, renamed and deleted', (t) => {
  const book = sales()
  t.alike(book.addSheet('Summary', { at: 0 }), { index: 0, name: 'Summary' })
  book.setCells('Summary', { A1: '=Sales!D6' })
  t.alike(book.renameSheet('Sales', 'Q Sales'), { changed: true, formulas: 3 })
  const again = reopen(book)
  t.alike(
    again.sheets().map((s) => s.name),
    ['Summary', 'Q Sales', 'Notes']
  )
  t.is(again.read('Summary').formulas.A1, "='Q Sales'!D6")
  t.is(again.read('Notes').formulas.B1, "='Q Sales'!D6*2")
  t.ok(partXml(again, 'xl/workbook.xml').includes(">'Q Sales'!$B$6</definedName>"))

  t.exception(
    () => again.deleteSheet('Q Sales'),
    /"Q Sales" is used by formulas on "Summary", formulas on "Notes", defined names TotalUnits/
  )
  again.deleteSheet('Summary')
  t.alike(
    reopen(again)
      .sheets()
      .map((s) => s.name),
    ['Q Sales', 'Notes']
  )
  t.absent(Object.keys(unzip(again.save())).includes('xl/worksheets/sheet3.xml'))
})

test('sheet names are validated', (t) => {
  const book = sales()
  t.exception(() => book.addSheet('notes'), /already has a sheet named "Notes"/)
  t.exception(() => book.addSheet('a/b'), /cannot contain/)
  t.exception(() => book.addSheet('x'.repeat(32)), /1 to 31 characters/)
  t.exception(() => book.addSheet('History'), /reserved/)
})

test('the last visible sheet cannot be deleted', (t) => {
  const book = budget()
  t.exception(() => book.deleteSheet(0), /only visible sheet/)
})

test('addImage anchors a picture to a cell in a drawing', (t) => {
  const book = budget()
  const placed = book.addImage(0, fixture('red.png'), { cell: 'D2', width: units.inches(2) })
  t.alike(placed, { cell: 'D2', width: 1828800, height: 914400 })
  book.addImage(0, fixture('red.png'), { cell: 'D8' })
  const files = unzip(book.save())
  t.ok(files['xl/drawings/drawing1.xml'])
  t.ok(files['xl/media/image1.png'])
  t.ok(files['xl/media/image2.png'])
  const drawing = decode(files['xl/drawings/drawing1.xml'])
  t.is(drawing.split('<xdr:oneCellAnchor>').length - 1, 2)
  t.ok(drawing.includes('<xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>'))
  t.ok(decode(files['xl/worksheets/sheet1.xml']).includes('<drawing r:id="rId1"'))
})

test('addImage with no cell sits beside the data, not over it', (t) => {
  const book = sales()
  t.is(book.addImage('Sales', fixture('red.png')).cell, 'J1')
  t.is(xlsx.create().addImage(0, fixture('red.png')).cell, 'A1')
})

test('a picture with no cell stacks below a chart or picture already beside the data', (t) => {
  const book = xlsx.create({ sheets: ['Data'] })
  book.setCells('Data', { A1: 'Month', B1: 'Sales', A2: 'Jan', B2: 10, A3: 'Feb', B3: 12 })
  const chart = book.addChart('Data', {
    type: 'bar',
    categories: 'A2:A3',
    series: [{ name: 'Sales', values: 'B2:B3' }]
  })

  t.is(chart.cell, 'D1')
  t.is(book.addImage('Data', fixture('red.png')).cell, 'D19', 'below the chart, not over it')
  t.is(book.addImage('Data', fixture('red.png'), { cell: 'D1' }).cell, 'D1', 'a named cell is kept')

  const empty = xlsx.create()
  empty.addImage(0, fixture('red.png'))
  t.is(empty.addImage(0, fixture('red.png')).cell, 'A3')
})

test('macros are reported and can be removed', (t) => {
  const book = budget()
  book.package.contentTypes.setOverride(
    book.main,
    'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
  )
  book.package.add('/xl/vbaProject.bin', new Uint8Array(8), 'application/vnd.ms-office.vbaProject')
  book.package
    .rels(book.main)
    .add('http://schemas.microsoft.com/office/2006/relationships/vbaProject', '/xl/vbaProject.bin')
  const macro = reopen(book)
  t.ok(macro.hasMacros)
  t.alike(macro.removeMacros(), { changed: true })
  const clean = reopen(macro)
  t.absent(clean.hasMacros)
  t.absent(Object.keys(unzip(clean.save())).includes('xl/vbaProject.bin'))
})

test('create builds a workbook with named sheets', (t) => {
  const book = xlsx.create({ sheets: ['Summary', 'Data'], title: 'Budget' })
  book.appendRows('Data', [
    ['Month', 'Spend'],
    ['Jan', 120],
    ['Feb', 80]
  ])
  book.setCells('Data', {
    A1: { value: 'Month', bold: true },
    B1: { value: 'Spend', bold: true },
    B4: '=SUM(B2:B3)'
  })
  book.setCells('Summary', { A1: 'Total', B1: '=Data!B4', C1: { date: '2026-01-31T09:30' } })
  book.setColumnWidth('Data', { column: 'A', width: 14 })
  const again = reopen(book)
  t.alike(
    again.sheets().map((s) => s.name),
    ['Summary', 'Data']
  )
  t.alike(again.read('Data').rows.slice(0, 3), [
    ['Month', 'Spend'],
    ['Jan', 120],
    ['Feb', 80]
  ])
  t.is(again.read('Summary').rows[0][2], '2026-01-31T09:30:00')
  t.is(again.get('Summary', 'C1').format, 'yyyy-mm-dd hh:mm')
  const files = unzip(again.save())
  t.ok(files['xl/styles.xml'])
  t.ok(files['xl/sharedStrings.xml'])
  t.ok(decode(files['docProps/core.xml']).includes('<dc:title>Budget</dc:title>'))
  t.exception(() => xlsx.create({ sheets: ['A', 'a'] }), /repeated/)
})

test('an untouched workbook saves byte for byte and an edited one survives cycles', (t) => {
  const input = fixture('sales.xlsx')
  const same = unzip(xlsx.open(input).save())
  const original = unzip(input)
  for (const name of Object.keys(original)) {
    t.alike(same[name], original[name], name)
  }

  let book = sales()
  book.setCells('Notes', { A3: 'third' })
  book = reopen(book)
  book.addSheet('Extra')
  book.appendRows('Extra', [[1, 2]])
  book = reopen(book)
  t.alike(book.read('Notes').rows[2], ['third', null])
  t.alike(book.read('Extra').rows, [[1, 2]])
})

test('shared strings read and write the _xHHHH_ escapes Excel uses', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 'placeholder' })
  const files = unzip(book.save())
  const strings = decode(files['xl/sharedStrings.xml']).replace(
    /<si>.*<\/si>/,
    '<si><t>Line1_x000D_\nLine2</t></si><si><r><t>a_x0009_</t></r><r><t>_x005F_x0041_</t></r></si>'
  )
  const excel = xlsx.open(zipFiles({ ...files, 'xl/sharedStrings.xml': utf8.encode(strings) }))
  excel.setCells(0, { A2: { formula: '=1' } })
  t.is(excel.read(0).rows[0][0], 'Line1\r\nLine2')
  excel.setCells(0, { B1: 'x_x000D_y', C1: 'tab\u000bbell\u0007end', D1: 'cr\rlf' })
  const again = reopen(excel)
  t.alike(again.read(0).rows[0], [
    'Line1\r\nLine2',
    'x_x000D_y',
    'tab\u000bbell\u0007end',
    'cr\rlf'
  ])
  const written = partXml(again, 'xl/sharedStrings.xml')
  t.ok(written.includes('x_x005F_x000D_y'))
  t.ok(written.includes('tab_x000B_bell_x0007_end'))
  t.ok(written.includes('cr_x000D_lf'))
  t.absent(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(written))
})

test('inline strings decode _xHHHH_ escapes and rich text runs', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 1 })
  const files = unzip(book.save())
  const sheet = decode(files['xl/worksheets/sheet1.xml']).replace(
    /<c r="A1"[^>]*>.*?<\/c>/,
    '<c r="A1" t="inlineStr"><is><r><t>one_x000A_</t></r><r><t>two</t></r><rPh sb="0" eb="1"><t>ignored</t></rPh></is></c>'
  )
  const opened = xlsx.open(zipFiles({ ...files, 'xl/worksheets/sheet1.xml': utf8.encode(sheet) }))
  t.is(opened.get(0, 'A1').value, 'one\ntwo')
})

test('formulas with control characters still save as XML a strict parser accepts', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: '="a\u000bb"' })
  const sheet = partXml(book, 'xl/worksheets/sheet1.xml')
  t.absent(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(sheet))
})

test('a sheet whose rows and cells carry no r attribute takes writes in place', (t) => {
  const book = xlsx.open(corpusFile('xlsx/59746_NoRowNums.xlsx'))
  const before = book.read(0)
  const sheet = book.sheets()[0].name
  book.setCells(sheet, { B2: 'edited', A40: 'appended', L1: 'extra' })
  t.is(book.get(sheet, 'B2').value, 'edited')
  const again = reopen(book)
  t.is(again.get(sheet, 'B2').value, 'edited')
  t.is(again.get(sheet, 'A40').value, 'appended')
  t.is(again.get(sheet, 'L1').value, 'extra')
  t.is(again.get(sheet, 'C2').value, before.rows[1][2])
  t.alike(again.read(sheet).rows[2].slice(0, before.rows[2].length), before.rows[2])
  const xmlText = partXml(again, 'xl/worksheets/sheet.xml')
  const rows = [...xmlText.matchAll(/<x:row[^>]*\br="(\d+)"/g)].map((m) => Number(m[1]))
  t.alike(
    rows,
    rows.slice().sort((a, b) => a - b),
    'rows stay in order'
  )
  t.is(rows.length, new Set(rows).size)
})
