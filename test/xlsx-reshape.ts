import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import * as xml from '../lib/xml.ts'
import { fixture, corpusFile, decode, unzip } from './helpers/engine.ts'
import type { Document, Element } from '../lib/xml.ts'

const X = xml.ns.x
const SHEET = '/xl/worksheets/sheet1.xml'
const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const partXml = (book: xlsx.Workbook, name: string): string => decode(unzip(book.save())[name])
const corpus = (name: string): xlsx.Workbook => xlsx.open(corpusFile(`xlsx/${name}`))

interface FormulaCell {
  ref: string
  type: string
  text?: string
  range?: string
}

function formulaCell(doc: Document, spec: FormulaCell): Element {
  const c = xml.named(doc, 'c', X)
  c.setAttribute('r', spec.ref)
  const f = xml.named(doc, 'f', X)
  f.setAttribute('t', spec.type)
  if (spec.type === 'shared') {
    f.setAttribute('si', '0')
  }

  if (spec.range) {
    f.setAttribute('ref', spec.range)
  }

  if (spec.text) {
    f.appendChild(doc.createTextNode(spec.text))
  }

  c.appendChild(f)

  return c
}

function cachedTotal(book: xlsx.Workbook): xlsx.Workbook {
  const sheet = book.package.xml(SHEET)
  const total = xml.descendants(sheet, X, 'c').find((c) => c.getAttribute('r') === 'B5')
  if (!total) {
    throw new Error('no B5 cell')
  }

  const v = xml.named(sheet, 'v', X)
  v.appendChild(sheet.createTextNode('1690.5'))
  total.appendChild(v)
  book.package.touch(SHEET)

  return reopen(book)
}

function budget(): xlsx.Workbook {
  const book = xlsx.create({ sheets: ['Budget', 'Summary'] })
  book.appendRows('Budget', [
    ['Item', 'Cost'],
    ['Rent', 1200],
    ['Food', 450.5],
    ['Gym', 40],
    ['Total', '=SUM(B2:B4)']
  ])
  book.setCells('Summary', { A1: 'Monthly total', B1: '=Budget!B5', C1: '=Budget!B4' })
  book.addChart('Summary', {
    type: 'pie',
    categories: 'Budget!A2:A4',
    series: [{ values: 'Budget!B2:B4', name: 'Cost' }]
  })
  const doc = book.package.xml(book.main)
  const names = xml.root(
    xml.parse(
      `<definedNames xmlns="${X}"><definedName name="MonthlyTotal">Budget!$B$5</definedName></definedNames>`
    )
  )
  const calc = xml.descendants(doc, X, 'calcPr')[0]
  xml.root(doc).insertBefore(doc.importNode(names, true), calc)
  book.package.touch(book.main)

  return cachedTotal(reopen(book))
}

test('dropping a row inside a SUM shrinks it and moves every reference', (t) => {
  const book = budget()
  t.is(book.read('Budget').rows[4][1], 1690.5)
  t.alike(book.deleteRows('Budget', { at: 4 }), { deleted: 1 })
  const again = reopen(book)
  const read = again.read('Budget')
  t.alike(read.rows, [
    ['Item', 'Cost'],
    ['Rent', 1200],
    ['Food', 450.5],
    ['Total', 1650.5]
  ])
  t.alike(read.formulas, { B4: '=SUM(B2:B3)' })
  t.alike(again.read('Summary').formulas, { B1: '=Budget!B4', C1: '=Budget!#REF!' })
  t.ok(partXml(again, 'xl/workbook.xml').includes('>Budget!$B$4</definedName>'))
  t.ok(partXml(again, 'xl/workbook.xml').includes('fullCalcOnLoad="1"'))
  const chart = partXml(again, 'xl/charts/chart1.xml')
  t.ok(chart.includes('<c:f>Budget!$B$2:$B$3</c:f>'))
  t.ok(chart.includes('<c:f>Budget!$A$2:$A$3</c:f>'))
})

test('inserting inside a range grows it, and above it moves it', (t) => {
  const book = budget()
  book.insertRows('Budget', { at: 3, rows: [['Phone', 30]] })
  book.insertRows('Budget', { at: 2, count: 1 })
  const read = reopen(book).read('Budget')
  t.is(read.formulas.B7, '=SUM(B3:B6)')
  t.alike(read.rows[3], ['Phone', 30])
  t.alike(read.rows[1], [null, null])
})

test('deleting every row of a range turns references to it into #REF!', (t) => {
  const book = budget()
  book.deleteRows('Budget', { at: 2, count: 3 })
  t.alike(reopen(book).read('Budget').formulas, { B2: '=SUM(#REF!)' })
})

test('sheet structures follow the rows', (t) => {
  const book = xlsx.create()
  book.appendRows(0, [['Title'], ['a', 1], ['b', 2], ['c', 3], ['d', 4]])
  const doc = book.package.xml(SHEET)
  const extra = xml.root(
    xml.parse(
      `<w xmlns="${X}"><autoFilter ref="A2:B5"/><mergeCells count="3"><mergeCell ref="A1:B1"/><mergeCell ref="A3:B3"/><mergeCell ref="A4:B5"/></mergeCells><conditionalFormatting sqref="B2:B5 D3"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>B2&gt;2</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" sqref="B3"><formula1>0</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="A3" location="Sheet1!A1"/><hyperlink ref="A5" location="Sheet1!A1"/></hyperlinks><rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks></w>`
    )
  )
  const root = xml.root(doc)
  const margins = xml.descendants(root, X, 'pageMargins')[0]
  for (const node of xml.children(extra)) {
    const copy = doc.importNode(node, true)
    if (node.localName === 'rowBreaks') {
      root.appendChild(copy)
    } else {
      root.insertBefore(copy, margins)
    }
  }

  book.package.touch(SHEET)
  book.addImage(0, fixture('red.png'), { cell: 'D5' })
  book.deleteRows(0, { at: 3 })
  const sheet = partXml(book, 'xl/worksheets/sheet1.xml')
  t.ok(sheet.includes('<autoFilter ref="A2:B4"/>'))
  t.ok(
    sheet.includes(
      '<mergeCells count="2"><mergeCell ref="A1:B1"/><mergeCell ref="A3:B4"/></mergeCells>'
    )
  )
  t.ok(sheet.includes('<conditionalFormatting sqref="B2:B4">'))
  t.ok(sheet.includes('<formula>B2&gt;2</formula>'))
  t.absent(sheet.includes('<dataValidations'))
  t.ok(sheet.includes('<hyperlinks><hyperlink ref="A4" location="Sheet1!A1"/></hyperlinks>'))
  t.ok(sheet.includes('<brk id="3" max="16383" man="1"/>'))
  const drawing = partXml(book, 'xl/drawings/drawing1.xml')
  t.ok(drawing.includes('<xdr:row>3</xdr:row>'))
})

test('shared formulas are expanded before rows move', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 1, A2: 2, A3: 3, A4: 4 })
  const doc = book.package.xml(SHEET)
  const rows = xml.descendants(doc, X, 'row')
  for (let i = 0; i < 4; i++) {
    const first = i === 0
    rows[i].appendChild(
      formulaCell(doc, {
        ref: `B${i + 1}`,
        type: 'shared',
        range: first ? 'B1:B4' : undefined,
        text: first ? 'A1*2' : undefined
      })
    )
  }

  book.package.touch(SHEET)
  book.deleteRows(0, { at: 1 })
  const again = reopen(book)
  t.alike(again.read(0).formulas, { B1: '=A1*2', B2: '=A2*2', B3: '=A3*2' })
  t.absent(partXml(again, 'xl/worksheets/sheet1.xml').includes('t="shared"'))
})

test('deleteColumns moves column references', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 1, B1: 2, C1: 3, D1: '=SUM(A1:C1)', E1: '=C1' })
  book.deleteColumns(0, { column: 'B' })
  t.alike(reopen(book).read(0).formulas, { C1: '=SUM(A1:B1)', D1: '=B1' })
})

test('column deletes through tables, comments and straddling array formulas still refuse', (t) => {
  t.exception(
    () => corpus('ExcelTables.xlsx').deleteColumns(0, { column: 'A' }),
    /it has tables, which cannot be moved yet/
  )
  t.exception(() => corpus('comments.xlsx').insertRows(0, { at: 2, count: 1 }), /it has comments/)
  const book = xlsx.create()
  book.setCells(0, { A1: 1, A2: 2, A3: 3 })
  const doc = book.package.xml(SHEET)
  const array = formulaCell(doc, { ref: 'B1', type: 'array', range: 'B1:B3', text: 'A1:A3*2' })
  xml.descendants(doc, X, 'row')[0].appendChild(array)
  book.package.touch(SHEET)
  t.exception(() => book.deleteRows(0, { at: 2 }), /an array formula over B1:B3/)
  t.execution(() => book.insertRows(0, { at: 5, count: 1 }))
})

test('rows move through tables, which grow and shrink with them', (t) => {
  const book = corpus('simple-monthly-budget.xlsx')
  const table = (): string => decode(unzip(book.save())['xl/tables/table1.xml'])
  t.ok(table().includes('ref="B4:C7"><autoFilter ref="B4:C7"/>'))
  book.insertRows(0, { at: 6, count: 1 })
  t.ok(table().includes('ref="B4:C8"><autoFilter ref="B4:C8"/>'))
  book.deleteRows(0, { at: 5, count: 2 })
  t.ok(table().includes('ref="B4:C6"><autoFilter ref="B4:C6"/>'))
  t.exception(() => book.deleteRows(0, { at: 4 }), /the header or totals row of table tblIncome/)
  t.exception(() => book.deleteColumns(0, { column: 'C' }), /it has tables/)
  t.execution(() => xlsx.open(book.save()).read(0))
})
