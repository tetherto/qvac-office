import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import * as xml from '../lib/xml.ts'
import { corpusFile, decode, unzip } from './helpers/engine.ts'

const X = xml.ns.x
const SHEET = '/xl/worksheets/sheet1.xml'
const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const partXml = (book: xlsx.Workbook, name: string): string => decode(unzip(book.save())[name])

function ledger(): xlsx.Workbook {
  const book = xlsx.create({ sheets: ['Data', 'Summary'] })
  book.appendRows('Data', [
    ['Item', 'Q1', 'Q2'],
    ['Rent', 100, 110],
    ['Food', 50, 60]
  ])
  book.setCells('Data', { D1: 'Total', D2: '=SUM(B2:C2)', D3: '=SUM(B3:C3)', E1: '=C1' })
  book.setCells('Summary', { A1: '=Data!D2', B1: '=SUM(Data!B2:C3)', C1: '=Data!A1' })
  book.addChart('Summary', {
    type: 'column',
    categories: 'Data!A2:A3',
    series: [{ values: 'Data!C2:C3', name: 'Q2' }]
  })
  book.defineName('Totals', 'Data!$D$2:$D$3')

  return reopen(book)
}

test('insertColumns moves cells and every reference to them', (t) => {
  const book = ledger()
  t.alike(book.insertColumns('Data', { column: 'C', count: 2 }), { from: 'C', to: 'D' })
  const again = reopen(book)
  const read = again.read('Data')
  t.alike(read.rows[0], ['Item', 'Q1', null, null, 'Q2', 'Total', 'Q2'])
  t.alike(read.formulas, { F2: '=SUM(B2:E2)', F3: '=SUM(B3:E3)', G1: '=E1' })
  t.alike(again.read('Summary').formulas, {
    A1: '=Data!F2',
    B1: '=SUM(Data!B2:E3)',
    C1: '=Data!A1'
  })
  t.alike(again.names(), [{ name: 'Totals', ref: 'Data!$F$2:$F$3' }])
  t.ok(partXml(again, 'xl/charts/chart1.xml').includes('<c:f>Data!$E$2:$E$3</c:f>'))
  t.ok(partXml(again, 'xl/charts/chart1.xml').includes('<c:f>Data!$A$2:$A$3</c:f>'))
})

test('insertColumns writes values one column at a time and takes the style on its right', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 'Name', B1: { value: 'Score', bold: true }, B2: 9 })
  const span = book.insertColumns(0, { column: 2, values: [['Age', 31], ['City']] })
  t.alike(span, { from: 'B', to: 'C' })
  const again = reopen(book)
  t.alike(again.read(0).rows, [
    ['Name', 'Age', 'City', 'Score'],
    [null, 31, null, 9]
  ])
  const sheet = again.package.xml(SHEET)
  const style = (ref: string): string | null =>
    xml
      .descendants(sheet, X, 'c')
      .find((c) => c.getAttribute('r') === ref)
      ?.getAttribute('s') ?? null
  t.is(style('B1'), style('D1'))
  t.ok(style('D1'))
})

test('insertColumns moves merges, validations, conditional formats and column widths', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 'a', B1: 'b', C1: 'c' })
  book.setColumnWidth(0, { column: 'B', width: 20 })
  book.merge(0, 'B2:C2')
  book.addValidation(0, 'C3:C5', { type: 'whole', min: 0 })
  book.addConditionalFormat(0, 'B1:C9', { kind: 'dataBar', color: '638EC6' })
  book.autofilter(0, 'A1:C9')
  book.insertColumns(0, { column: 'B' })
  const sheet = partXml(book, 'xl/worksheets/sheet1.xml')
  t.ok(sheet.includes('<col min="3" max="3" width="20" customWidth="1"/>'))
  t.ok(sheet.includes('<mergeCell ref="C2:D2"/>'))
  t.ok(sheet.includes('sqref="D3:D5"'))
  t.ok(sheet.includes('<conditionalFormatting sqref="C1:D9">'))
  t.ok(sheet.includes('<autoFilter ref="A1:D9"/>'))
})

test('insertColumns refuses the same things deleteColumns does', (t) => {
  t.exception(
    () => xlsx.open(corpusFile('xlsx/ExcelTables.xlsx')).insertColumns(0, { column: 'A' }),
    /cannot insert columns .* it has tables/
  )
  t.exception(
    () => xlsx.open(corpusFile('xlsx/comments.xlsx')).insertColumns(0, { column: 'A' }),
    /it has comments/
  )
  const book = xlsx.create()
  book.setCells(0, { XFC1: 1 })
  t.exception(() => book.insertColumns(0, { column: 'A', count: 2 }), /past column XFD/)
  t.exception(() => book.insertColumns(0, { column: 'A', count: 0 }), /count must be at least 1/)
  t.exception(() => book.insertColumns(0, { column: 'A', count: 2, values: [[1]] }), /give one/)
})
