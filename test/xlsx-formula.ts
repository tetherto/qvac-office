import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import * as xml from '../lib/xml.ts'
import * as formula from '../lib/xlsx/formula.ts'
import { corpusFile, decode, unzip } from './helpers/engine.ts'
import type { Change } from '../lib/xlsx/formula.ts'

const X = xml.ns.x
const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const partXml = (book: xlsx.Workbook, name: string): string => decode(unzip(book.save())[name])
const corpus = (name: string): xlsx.Workbook => xlsx.open(corpusFile(`xlsx/${name}`))
const rows = (sheet: string, at: number, count = 1): Change => ({
  sheet,
  axis: 'row',
  at,
  count,
  kind: 'insert'
})

function inject(book: xlsx.Workbook, part: string, markup: string): void {
  const doc = book.package.xml(part)
  const root = xml.root(doc)
  const margins = xml.descendants(root, X, 'pageMargins')[0] ?? null
  for (const node of xml.children(xml.root(xml.parse(`<w xmlns="${X}">${markup}</w>`)))) {
    root.insertBefore(doc.importNode(node, true), margins)
  }

  book.package.touch(part)
}

test('structured references are left alone when rows move', (t) => {
  const change = rows('Data', 0, 2)
  t.is(formula.adjust('SUM(Sales[Q1])', 'Data', change), 'SUM(Sales[Q1])')
  t.is(formula.adjust('Sales[[#This Row],[Q1]]*B2', 'Data', change), 'Sales[[#This Row],[Q1]]*B4')
  t.is(formula.adjust('SUM(T1[Q1])+[@Q2]', 'Data', change), 'SUM(T1[Q1])+[@Q2]')
  t.is(formula.adjust("Sales['[Q1']]", 'Data', change), "Sales['[Q1']]")
  t.alike(formula.tablesIn('SUM(Sales[Q1])+T1[[#Totals],[Q2]]'), ['Sales', 'T1'])
})

test('a structured reference survives insertRows on its own sheet', (t) => {
  const book = xlsx.create({ sheets: ['Data'] })
  book.setCells('Data', { A1: 'Q1', A2: 1, B5: '=SUM(Sales[Q1])' })
  book.insertRows('Data', { at: 1, count: 1 })
  t.alike(reopen(book).read('Data').formulas, { B6: '=SUM(Sales[Q1])' })
})

test('references into another workbook are never moved or renamed', (t) => {
  t.is(formula.adjust('[1]Sheet1!B2+B2', 'Sheet1', rows('Sheet1', 0)), '[1]Sheet1!B2+B3')
  t.is(formula.adjust("'[1]Sheet 1'!B2", 'Sheet 1', rows('Sheet 1', 0)), "'[1]Sheet 1'!B2")
  t.is(formula.adjust('[1]!Total+B2', 'Sheet1', rows('Sheet1', 0)), '[1]!Total+B3')
  t.is(
    formula.renameSheet('[1]Sheet1!B2+Sheet1!B2', { from: 'Sheet1', to: 'Data' }),
    '[1]Sheet1!B2+Data!B2'
  )
  t.absent(formula.mentions('[1]Sheet1!B2', 'Sheet1'))
  t.alike(formula.tablesIn('[1]Sheet1!Sales[Q1]'), [])
})

test('inserting rows keeps references on the sheet', (t) => {
  const change = rows('S', 3, 2)
  t.is(formula.adjust('SUM(B2:B1048576)', 'S', change), 'SUM(B2:B1048576)')
  t.is(formula.adjust('SUM(B5:B1048575)', 'S', change), 'SUM(B7:B1048576)')
  t.is(formula.adjust('B1048575+1', 'S', change), '#REF!+1')
  t.is(formula.adjust('SUM(B1048575:B1048576)', 'S', change), 'SUM(#REF!)')
  t.is(formula.adjust('SUM(1048575:1048576)', 'S', change), 'SUM(#REF!)')
  t.is(formula.adjustRange('B1:B1048576', rows('S', 0)), 'B2:B1048576')
  t.is(formula.adjustRange('B1048576', rows('S', 0)), null)
  const cols: Change = { sheet: 'S', axis: 'col', at: 0, count: 1, kind: 'insert' }
  t.is(formula.adjust('SUM(B1:XFD1)+XFD2', 'S', cols), 'SUM(C1:XFD1)+#REF!')
})

test('insertRows leaves whole-column ranges and rules inside the sheet', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 1, C1: '=SUM(B2:B1048576)' })
  inject(
    book,
    '/xl/worksheets/sheet1.xml',
    '<conditionalFormatting sqref="B1:B1048576 D1048576"><cfRule type="expression" priority="1"><formula>B1&gt;2</formula></cfRule></conditionalFormatting>'
  )
  book.insertRows(0, { at: 1, count: 1 })
  const sheet = partXml(book, 'xl/worksheets/sheet1.xml')
  t.ok(sheet.includes('<conditionalFormatting sqref="B2:B1048576">'))
  t.alike(reopen(book).read(0).formulas, { C2: '=SUM(B3:B1048576)' })
})

test('3D references see both of their sheets', (t) => {
  t.alike(formula.sheetsIn('SUM(Jan:Mar!B2)'), ['Jan', 'Mar'])
  t.alike(formula.sheetsIn("SUM('Jan 1:Mar 2'!B2)"), ['Jan 1', 'Mar 2'])
  t.is(
    formula.renameSheet('SUM(Jan:Mar!B2)', { from: 'Mar', to: 'March 2026' }),
    "SUM('Jan:March 2026'!B2)"
  )
  t.is(
    formula.renameSheet('SUM(Jan:Mar!B2)', { from: 'Jan', to: 'January' }),
    'SUM(January:Mar!B2)'
  )
  t.is(
    formula.renameSheet("SUM('Jan 1:Mar''s'!B2)", { from: 'Jan 1', to: 'Start' }),
    "SUM('Start:Mar''s'!B2)"
  )
  t.ok(formula.mentions('SUM(Jan:Mar!B2)', 'Jan'))
  t.is(formula.adjust('SUM(Jan:Mar!B2)', 'Summary', rows('Jan', 0)), 'SUM(Jan:Mar!B2)')

  const book = xlsx.create({ sheets: ['Jan', 'Feb', 'Mar', 'Summary'] })
  book.setCells('Summary', { A1: '=SUM(Jan:Mar!B2)' })
  book.renameSheet('Mar', 'March 2026')
  book.renameSheet('Jan', 'January')
  t.alike(reopen(book).read('Summary').formulas, { A1: "=SUM('January:March 2026'!B2)" })
  t.exception(
    () => book.deleteSheet('January'),
    /SHEET_IS_REFERENCED|is used by formulas on "Summary"/
  )
})

test('rows do not move under a pivot table placed on the sheet', (t) => {
  const book = corpus('ExcelPivotTableSample.xlsx')
  t.exception(
    () => book.insertRows('Tabelle2', { at: 1, count: 1 }),
    /cannot insert rows on "Tabelle2": it has pivot tables/
  )
})

test('column deletes re-index the autoFilter columns', (t) => {
  const book = xlsx.create()
  book.appendRows(0, [
    ['a', 'b', 'c', 'd'],
    [1, 2, 3, 4],
    [5, 6, 7, 8]
  ])
  inject(
    book,
    '/xl/worksheets/sheet1.xml',
    '<autoFilter ref="A1:D3"><filterColumn colId="1"><filters><filter val="2"/></filters></filterColumn><filterColumn colId="3"><filters><filter val="4"/></filters></filterColumn></autoFilter>'
  )
  book.deleteColumns(0, { column: 'B' })
  const sheet = partXml(book, 'xl/worksheets/sheet1.xml')
  t.ok(sheet.includes('<autoFilter ref="A1:C3"><filterColumn colId="2">'))
  t.absent(sheet.includes('colId="1"'))
  t.absent(sheet.includes('colId="3"'))
})

const LIST_RULES =
  '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1=Lists!$A$1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>Lists!$A$1:$A$3</formula1></dataValidation></dataValidations>'

test('renameSheet rewrites conditional formatting and validation formulas', (t) => {
  const book = xlsx.create({ sheets: ['Form', 'Lists'] })
  book.appendRows('Lists', [['red'], ['green'], ['blue']])
  inject(book, '/xl/worksheets/sheet1.xml', LIST_RULES)
  t.alike(book.renameSheet('Lists', 'Choices'), { changed: true, formulas: 2 })
  const sheet = partXml(book, 'xl/worksheets/sheet1.xml')
  t.ok(sheet.includes('<formula1>Choices!$A$1:$A$3</formula1>'))
  t.ok(sheet.includes('<formula>A1=Choices!$A$1</formula>'))
})

test('deleteSheet refuses a sheet that validation or formatting rules use', (t) => {
  const book = xlsx.create({ sheets: ['Form', 'Lists'] })
  inject(book, '/xl/worksheets/sheet1.xml', LIST_RULES)
  t.exception(() => book.deleteSheet('Lists'), /"Lists" is used by formulas on "Form"/)
})

test('deleteSheet refuses a sheet whose table other sheets use', (t) => {
  const book = corpus('ExcelTables.xlsx')
  book.addSheet('Report')
  book.setCells('Report', { A1: '=SUM(TableName[Age])' })
  t.exception(
    () => book.deleteSheet('ExcelTable'),
    /"ExcelTable" is used by formulas on "Report" that use table TableName/
  )
})

test('a new workbook names QVAC as the application that made it', (t) => {
  t.ok(partXml(xlsx.create(), 'docProps/app.xml').includes('<Application>QVAC</Application>'))
})
